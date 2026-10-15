import React, { useState, useCallback, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

/* ---------- configuración ---------- */

const MODELO = "claude-sonnet-4-6";
const MAX_TOKENS_RESPUESTA = 16000; // antes 1000: con registros largos la respuesta se cortaba
const PAGINAS_POR_PARTE = 8; // los PDF largos se envían en partes de 8 páginas
const MAX_BYTES_POR_PARTE = 18 * 1024 * 1024; // la API acepta hasta 32 MB por solicitud (en base64, ~24 MB binarios)
const MAX_DIM_IMAGEN = 2000; // px del lado mayor; más grande no mejora la lectura y solo pesa más
const MAX_BYTES_IMAGEN = 3.5 * 1024 * 1024; // la API acepta hasta 5 MB por imagen (en base64)
const PARTES_EN_PARALELO = 3;
const REINTENTOS = 4;
const TIMEOUT_MS = 180000;
const UMBRAL = 0.6; // más del 60%, ya sea por documento o por nombre
const UMBRAL_DEBIL = 0.75;
const ESTADO_PRIORIDAD = { ACTIVO: 0, BAJA: 1, VISITA: 2 };

const CAMPOS_DISPONIBLES = [
  { key: "nombre", label: "Apellidos y Nombres" },
  { key: "dni", label: "DNI o CE" },
  { key: "cargo", label: "Cargo o Puesto" },
  { key: "fechaIngreso", label: "Fecha de Ingreso" },
  { key: "fechaNacimiento", label: "Fecha de Nacimiento" },
  { key: "celular", label: "Celular" },
  { key: "frente", label: "Frente" },
];

const tick = () => new Promise((r) => setTimeout(r, 0));

function formatoMB(bytes) {
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

/* ---------- utilidades de texto ---------- */

function stripAccents(s) {
  return s.normalize("NFD").replace(/[̀-ͯ]/g, "");
}
export function normName(s) {
  if (!s) return "";
  return stripAccents(String(s))
    .toUpperCase()
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
export function onlyDigitsAlnum(s) {
  if (!s) return "";
  return String(s).toUpperCase().replace(/[^0-9A-Z]/g, "");
}

// Levenshtein con dos filas reutilizadas (sin crear una matriz por comparación) y corte temprano:
// si la distancia ya supera "max", deja de calcular.
let LEV_A = new Int32Array(64);
let LEV_B = new Int32Array(64);
export function levenshtein(a, b, max = Infinity) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  if (Math.abs(m - n) > max) return max + 1;
  if (n + 1 > LEV_A.length) {
    LEV_A = new Int32Array(n + 1);
    LEV_B = new Int32Array(n + 1);
  }
  let prev = LEV_A, cur = LEV_B;
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    let rowMin = i;
    const ca = a.charCodeAt(i - 1);
    for (let j = 1; j <= n; j++) {
      let v = prev[j - 1] + (ca === b.charCodeAt(j - 1) ? 0 : 1);
      const del = prev[j] + 1;
      const ins = cur[j - 1] + 1;
      if (del < v) v = del;
      if (ins < v) v = ins;
      cur[j] = v;
      if (v < rowMin) rowMin = v;
    }
    if (rowMin > max) return max + 1;
    const t = prev; prev = cur; cur = t;
  }
  return prev[n];
}

// similitud 0..1; si no puede llegar a "min", devuelve 0 sin terminar el cálculo
function similarity(a, b, min = 0) {
  const L = Math.max(a.length, b.length);
  if (!L) return 0;
  const maxDist = min > 0 ? Math.floor((1 - min) * L + 1e-9) : Infinity;
  const d = levenshtein(a, b, maxDist);
  if (d > maxDist) return 0;
  return 1 - d / L;
}

export function prepName(s) {
  const n = normName(s);
  const tokens = [...new Set(n.split(" ").filter(Boolean))].slice(0, 30);
  return { n, tokens, sorted: [...tokens].sort().join(" "), cache: null };
}

// dos palabras cuentan como iguales si difieren en 1 letra (2 si son largas): "QUISPE" ~ "QUIZPE"
function tokensParecidos(a, b) {
  if (a === b) return true;
  if (a.length < 4 || b.length < 4) return false;
  const maxD = Math.max(a.length, b.length) <= 6 ? 1 : 2;
  if (Math.abs(a.length - b.length) > maxD) return false;
  return levenshtein(a, b, maxD) <= maxD;
}

// Los apellidos y nombres se repiten mucho en el maestro, así que la comparación de cada palabra
// del maestro contra las palabras leídas se calcula una sola vez y se guarda (q.cache).
function mascaraToken(q, et) {
  let m = q.cache.get(et);
  if (m === undefined) {
    m = 0;
    for (let i = 0; i < q.tokens.length; i++) if (tokensParecidos(q.tokens[i], et)) m |= 1 << i;
    q.cache.set(et, m);
  }
  return m;
}

function jaccardTolerante(q, et) {
  if (!q.tokens.length || !et.length) return 0;
  let usados = 0, shared = 0;
  for (const t of et) {
    const libres = mascaraToken(q, t) & ~usados;
    if (libres) { usados |= libres & -libres; shared++; }
  }
  return shared / (q.tokens.length + et.length - shared);
}

// misma fórmula que antes (60% palabras compartidas + 40% parecido letra a letra), pero:
// - tolera errores de una letra en cada palabra,
// - compara también con las palabras ordenadas ("JUAN PEREZ" vs "PEREZ JUAN"),
// - si el candidato no puede superar "floor", se descarta sin calcular Levenshtein.
export function nameScore(q, e, floor = 0) {
  if (!q.n || !e.n) return 0;
  if (!q.cache) q.cache = new Map();
  const jac = jaccardTolerante(q, e.tokens);
  const L = Math.max(q.n.length, e.n.length);
  const cotaLetras = 1 - Math.abs(q.n.length - e.n.length) / L;
  if (0.6 * jac + 0.4 * cotaLetras < floor) return 0;
  const minLetras = Math.max(0, (floor - 0.6 * jac) / 0.4);
  const c1 = similarity(q.n, e.n, minLetras);
  const c2 = q.sorted === q.n && e.sorted === e.n ? c1 : similarity(q.sorted, e.sorted, minLetras);
  return 0.6 * jac + 0.4 * Math.max(c1, c2);
}

export function docScore(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.length === b.length) {
    let same = 0;
    for (let i = 0; i < a.length; i++) if (a[i] === b[i]) same++;
    return same / a.length;
  }
  return similarity(a, b);
}

/* ---------- normalización de Cargo/Puesto: expande abreviaturas comunes antes de comparar ---------- */

const CARGO_ABREVIATURAS = {
  OP: "OPERARIO",
  OPER: "OPERARIO",
  SUP: "SUPERVISOR",
  SUPERV: "SUPERVISOR",
  ING: "INGENIERO",
  ASIST: "ASISTENTE",
  AYUD: "AYUDANTE",
  AYTE: "AYUDANTE",
  TEC: "TECNICO",
  ALM: "ALMACENERO",
  ALMAC: "ALMACENERO",
  SEG: "SEGURIDAD",
  ADM: "ADMINISTRATIVO",
  ADMIN: "ADMINISTRATIVO",
  RESID: "RESIDENTE",
  CAP: "CAPATAZ",
  MAQ: "MAQUINISTA",
  ELEC: "ELECTRICISTA",
  SOLD: "SOLDADOR",
  GASF: "GASFITERO",
  ALB: "ALBANIL",
  ALBAN: "ALBANIL",
  PEON: "PEON",
  PNO: "PEON",
  VIG: "VIGILANTE",
  ENC: "ENCARGADO",
  CHOF: "CHOFER",
  MEC: "MECANICO",
  SSOMA: "SSOMA",
  PREV: "PREVENCIONISTA",
  JEF: "JEFE",
  GTE: "GERENTE",
  COORD: "COORDINADOR",
};

// sinónimos: mismo cargo, distinta palabra (no es abreviatura). Se normalizan a un término único.
const CARGO_SINONIMOS = {
  CHOFER: "CONDUCTOR",
  MOTORISTA: "CONDUCTOR",
  GUARDIAN: "VIGILANTE",
  SERENO: "VIGILANTE",
  OBRERO: "OPERARIO",
  MAESTRO: "CAPATAZ", // "maestro de obra" suele equivaler a capataz
};

// normName ya quita tildes y la Ñ queda como N, por eso las claves van sin tilde ("TEC", "ALBANIL")
export function expandCargo(s) {
  if (!s) return "";
  return normName(s)
    .split(" ")
    .filter(Boolean)
    .map((tok) => {
      const expandido = CARGO_ABREVIATURAS[tok] || tok;
      return CARGO_SINONIMOS[expandido] || expandido;
    })
    .join(" ");
}

/* ---------- lectura del maestro "fuerza laboral" ---------- */

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row = (rows[i] || []).map((c) => stripAccents(String(c || "").toUpperCase()));
    const hasDoc = row.some((c) => c.includes("DNI") || c.includes("DOC"));
    const hasName = row.some((c) => c.includes("APELLIDOS") || c.includes("NOMBRES"));
    if (hasDoc && hasName) return i;
  }
  return -1;
}

function formatCellValue(raw) {
  if (raw instanceof Date && !isNaN(raw)) {
    const dd = String(raw.getDate()).padStart(2, "0");
    const mm = String(raw.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${raw.getFullYear()}`;
  }
  if (raw === null || raw === undefined) return "";
  return String(raw).trim();
}

function parseSheet(ws, estado) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true, blankrows: false });
  const headerIdx = findHeaderRow(rows);
  if (headerIdx === -1) return [];
  const headerN = rows[headerIdx].map((c) => stripAccents(String(c || "").toUpperCase()));
  const col = (...kws) => headerN.findIndex((c) => kws.some((k) => c.includes(k)));
  const docCol = col("DNI", "DOC");
  const nameCol = col("APELLIDOS", "NOMBRES");
  const cargoCol = col("CARGO", "PUESTO", "POSICION");
  const ingresoCol = col("INGRESO");
  const nacimientoCol = col("NACIMIENTO", "NAC.");
  const celularCol = col("CELULAR", "TELEFONO", "CONTACTO");
  const frenteCol = col("FRENTE");
  if (docCol === -1 || nameCol === -1) return [];
  const get = (r, c) => (c === -1 ? "" : formatCellValue(r[c]));
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const rawDoc = r[docCol];
    let doc = onlyDigitsAlnum(rawDoc);
    // si Excel guardó el DNI como número, pudo perder los ceros iniciales (DNI peruano = 8 dígitos)
    if (typeof rawDoc === "number" && /^[0-9]{5,7}$/.test(doc)) doc = doc.padStart(8, "0");
    const nombre = String(r[nameCol] || "").trim();
    if (!doc && !nombre) continue;
    out.push({
      dni: doc,
      nombre,
      cargo: get(r, cargoCol),
      fechaIngreso: get(r, ingresoCol),
      fechaNacimiento: get(r, nacimientoCol),
      celular: get(r, celularCol),
      frente: get(r, frenteCol),
      estado,
    });
  }
  return out;
}

// Lee solo las hojas DATA / BAJA / VISITA. Primero se leen únicamente los nombres de hoja (rápido)
// y después solo esas hojas, sin estilos ni fórmulas: en archivos pesados es varias veces más rápido.
export function leerFuerzaLaboral(data) {
  const nombres = XLSX.read(data, { type: "array", bookSheets: true }).SheetNames;
  const up = (n) => stripAccents(n.toUpperCase());
  // se leen TODAS las hojas cuyo nombre contenga DATA, BAJA o VISITA (p. ej. "DATA" y "DATA SUBCONTRATAS")
  const plan = nombres
    .map((n) => [n, up(n).includes("BAJA") ? "BAJA" : up(n).includes("VISITA") ? "VISITA" : up(n).includes("DATA") ? "ACTIVO" : null])
    .filter(([, estado]) => estado);
  if (plan.length === 0) return { master: [], sheetsFound: [], conteo: [], hojasDisponibles: nombres };

  const wb = XLSX.read(data, {
    type: "array",
    sheets: plan.map(([n]) => n),
    cellDates: true,
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellStyles: false,
    cellText: false,
  });
  let master = [];
  const conteo = [];
  for (const [hoja, estado] of plan) {
    const rows = wb.Sheets[hoja] ? parseSheet(wb.Sheets[hoja], estado) : [];
    master = master.concat(rows);
    conteo.push(`${hoja} (${rows.length})`);
  }
  return { master, sheetsFound: plan.map(([n]) => n), conteo, hojasDisponibles: nombres };
}

/* ---------- guardado de la fuerza laboral (en partes, para que entren maestros grandes) ---------- */

const CLAVE_MAESTRO = "fuerza-laboral:master";
const TAM_PARTE_STORAGE = 900000; // caracteres por clave, por debajo del límite de window.storage
const CAMPOS_MAESTRO = ["dni", "nombre", "cargo", "fechaIngreso", "fechaNacimiento", "celular", "frente", "estado"];

async function guardarMaestro({ master, sheetsFound, conteo, name, savedAt }) {
  // formato compacto: arreglos en vez de objetos (repite menos texto, ocupa ~40% menos)
  const json = JSON.stringify(master.map((m) => CAMPOS_MAESTRO.map((k) => m[k] || "")));
  const partes = Math.max(1, Math.ceil(json.length / TAM_PARTE_STORAGE));
  for (let i = 0; i < partes; i++) {
    await window.storage.set(`${CLAVE_MAESTRO}:p${i}`, json.slice(i * TAM_PARTE_STORAGE, (i + 1) * TAM_PARTE_STORAGE), false);
  }
  // el índice se escribe al final: si algo falló antes, se conserva el guardado anterior completo
  await window.storage.set(CLAVE_MAESTRO, JSON.stringify({ v: 2, partes, name, savedAt, sheetsFound, conteo }), false);
}

async function cargarMaestro() {
  const meta = await window.storage.get(CLAVE_MAESTRO, false);
  if (!meta || !meta.value) return null;
  const p = JSON.parse(meta.value);
  if (p.v !== 2) return p.master ? p : null; // formato anterior (un solo bloque)
  const trozos = await Promise.all(
    Array.from({ length: p.partes }, (_, i) => window.storage.get(`${CLAVE_MAESTRO}:p${i}`, false))
  );
  if (trozos.some((t) => !t || typeof t.value !== "string")) return null;
  const filas = JSON.parse(trozos.map((t) => t.value).join(""));
  const master = filas.map((f) => Object.fromEntries(CAMPOS_MAESTRO.map((k, j) => [k, f[j]])));
  return { ...p, master };
}

/* ---------- índice del maestro: se arma una sola vez, no en cada fila del registro ---------- */

export function indexarMaestro(master) {
  const entradas = master.map((m, i) => ({
    ...m,
    _i: i,
    _name: prepName(m.nombre),
    _cargo: expandCargo(m.cargo),
  }));
  const byDni = new Map();
  const byCargo = new Map();
  const byToken = new Map();
  const add = (map, k, e) => {
    const l = map.get(k);
    if (l) l.push(e); else map.set(k, [e]);
  };
  for (const e of entradas) {
    if (e.dni) {
      add(byDni, e.dni, e);
      // también sin prefijo ("CE00123456" → "00123456"), para encontrarlo se lea como se lea
      const limpio = onlyDigitsAlnum(cleanDni(e.dni));
      if (limpio && limpio !== e.dni) add(byDni, limpio, e);
    }
    if (e._cargo) add(byCargo, e._cargo, e);
    for (const t of e._name.tokens) add(byToken, t, e);
  }
  return { entradas, byDni, byCargo, byToken };
}

export function cleanDni(raw) {
  if (!raw) return "";
  let s = String(raw).trim();
  // quita prefijos tipo "DNI", "DNI:", "N°", "CE" antes de los dígitos
  s = s.replace(/^\s*(DNI|D\.?N\.?I\.?|N[°º]?|CE)\s*[:\-]?\s*/i, "");
  // si queda una sola letra "D" pegada justo antes de los dígitos, quítala
  s = s.replace(/^D(?=\d{6,9}$)/i, "");
  return s.trim();
}

const prioridad = (m) => ESTADO_PRIORIDAD[m.estado] ?? 9;

// Busca el mejor candidato. Primero revisa a quienes comparten alguna palabra del nombre
// (así encuentra pronto un buen puntaje) y luego al resto, descartando sin calcular a quien
// no pueda superar el mejor puntaje ya encontrado.
function buscarMejor(q, dniLeido, candidatos, semillas, filtro) {
  let best = null, bestScore = -1, bestSec = -1;
  const visto = new Set();
  const evalua = (m) => {
    if (visto.has(m._i)) return;
    visto.add(m._i);
    const doc = dniLeido && m.dni ? docScore(dniLeido, m.dni) : 0;
    const name = nameScore(q, m._name, bestScore - 1e-9);
    const score = Math.max(doc, name);
    const sec = Math.min(doc, name);
    const mejor =
      score > bestScore ||
      (score === bestScore && best && (sec > bestSec || (sec === bestSec && prioridad(m) < prioridad(best))));
    if (mejor) { best = m; bestScore = score; bestSec = sec; }
  };
  for (const m of semillas) if (!filtro || filtro(m)) evalua(m);
  for (const m of candidatos) evalua(m);
  return best ? { best, score: bestScore } : null;
}

const notaEstado = (m) =>
  m.estado === "BAJA" ? "Figura como Dado de Baja, no como personal activo." : m.estado === "VISITA" ? "Figura en hoja VISITA." : "";

function filaDesdeMaestro(m, estado, obs, dudoso = false) {
  return {
    dni: m.dni, // copiado literal del maestro
    nombre: m.nombre, // copiado literal del maestro
    cargo: m.cargo || "",
    fechaIngreso: m.fechaIngreso || "",
    fechaNacimiento: m.fechaNacimiento || "",
    celular: m.celular || "",
    frente: m.frente || "",
    estado,
    obs: obs.filter(Boolean).join(" "),
    dudoso,
  };
}

export function matchAgainstMaster(extracted, indice) {
  const dniLeido = onlyDigitsAlnum(cleanDni(extracted.dni));
  const q = prepName(extracted.nombre);
  const cargoLeidoExp = expandCargo(extracted.cargo);
  const leido = `en el registro se leyó "${extracted.dni}" / "${extracted.nombre}"`;

  // 1) DNI exacto: búsqueda directa en el índice (también con los ceros iniciales que Excel o la lectura pudieron perder)
  if (dniLeido.length >= 6) {
    const hits =
      indice.byDni.get(dniLeido) ||
      indice.byDni.get(onlyDigitsAlnum(extracted.dni)) || // con prefijo, p. ej. un CE escrito "CE00123456"
      (/^\d{5,7}$/.test(dniLeido) ? indice.byDni.get(dniLeido.padStart(8, "0")) : null);
    if (hits) {
      const m = hits.reduce((a, b) => (prioridad(b) < prioridad(a) ? b : a));
      const nombreSim = nameScore(q, m._name);
      const nombreDistinto = q.n && nombreSim < 0.5;
      return filaDesdeMaestro(m, nombreDistinto ? "REVISAR" : "OK", [
        nombreDistinto
          ? `El DNI coincide, pero el nombre leído ("${extracted.nombre}") es distinto al del maestro (${Math.round(nombreSim * 100)}%); verificar.`
          : "",
        notaEstado(m),
      ], nombreDistinto);
    }
  }

  // 2) búsqueda por similitud. Si el registro trae cargo, primero se busca solo entre quienes tienen ese cargo
  //    (sobre todas las hojas a la vez); si ahí no hay una coincidencia clara, se busca en toda la fuerza laboral.
  const semillas = q.tokens.flatMap((t) => indice.byToken.get(t) || []);
  let res = null;
  let filtradoPorCargo = false;
  let cargoNoCoincide = false;
  const porCargo = cargoLeidoExp ? indice.byCargo.get(cargoLeidoExp) : null;
  if (porCargo) {
    res = buscarMejor(q, dniLeido, porCargo, semillas, (m) => m._cargo === cargoLeidoExp);
    filtradoPorCargo = true;
  }
  if (!res || res.score < UMBRAL_DEBIL) {
    const global = buscarMejor(q, dniLeido, indice.entradas, semillas, null);
    if (global && (!res || global.score > res.score)) {
      cargoNoCoincide = filtradoPorCargo && global.score > UMBRAL;
      res = global;
      filtradoPorCargo = false;
    }
  }

  // Una coincidencia solo por parecido de DNI (2 o más dígitos distintos) con un nombre que no se parece
  // suele ser otra persona: con miles de trabajadores, que 5 de 8 dígitos coincidan pasa por azar.
  let soloDniDebil = false;
  if (res && res.score > UMBRAL && dniLeido && q.n) {
    const d = docScore(dniLeido, res.best.dni);
    const n = nameScore(q, res.best._name);
    soloDniDebil = d >= n && d < UMBRAL_DEBIL && n < 0.4;
  }

  if (res && res.score > UMBRAL && !soloDniDebil) {
    const m = res.best;
    // puntajes exactos (sin descarte) para el informe
    const d = dniLeido && m.dni ? docScore(dniLeido, m.dni) : 0;
    const n = nameScore(q, m._name);
    // OK solo si lo leído coincide tal cual; si se corrigió el DNI o el nombre, queda como CORREGIDO
    const isExact = res.score === 1 && (!dniLeido || d === 1) && (!q.n || n === 1);
    return filaDesdeMaestro(m, isExact ? "OK" : "CORREGIDO", [
      !isExact && d === 1
        ? `DNI exacto; el nombre se tomó del maestro (${leido}).`
        : !isExact
        ? `Coincidencia por ${d >= n ? "documento" : "nombre"} (${Math.round(res.score * 100)}% — DNI:${Math.round(d * 100)}%, Nombre:${Math.round(n * 100)}%); datos tomados del maestro (${leido}).`
        : "",
      filtradoPorCargo ? `Búsqueda acotada por cargo/puesto ("${extracted.cargo}").` : "",
      cargoNoCoincide ? `El cargo leído ("${extracted.cargo}") no coincide con el del maestro ("${m.cargo}").` : "",
      res.score < UMBRAL_DEBIL ? "Similitud moderada, verificar manualmente." : "",
      notaEstado(m),
    ], res.score < UMBRAL_DEBIL);
  }

  // sin similitud suficiente: se deja lo leído y, si hay algo cercano, se sugiere para revisión manual
  const sugerencia = soloDniDebil
    ? ` El DNI se parece al de ${res.best.nombre} (${res.best.dni}, ${Math.round(res.score * 100)}%), pero el nombre no coincide; verificar manualmente.`
    : res && res.score >= 0.4 ? ` Lo más parecido: ${res.best.nombre} (${res.best.dni || "sin doc"}) con ${Math.round(res.score * 100)}%.` : "";
  return {
    dni: extracted.dni,
    nombre: extracted.nombre,
    cargo: extracted.cargo || "",
    fechaIngreso: "",
    fechaNacimiento: "",
    celular: "",
    frente: "",
    estado: "NO ENCONTRADO",
    obs: "No se encontró en la fuerza laboral (ni por documento ni por nombre con más de 60% de similitud). Se dejó el dato tal como se leyó." + sugerencia,
    dudoso: true,
  };
}

// los candidatos más parecidos de la fuerza laboral (por nombre y por DNI), para la segunda lectura
export function candidatosPara(extracted, indice, kNombre = 4, kDoc = 3) {
  const dni = onlyDigitsAlnum(cleanDni(extracted.dni));
  const q = prepName(extracted.nombre);
  const topN = [], topD = [];
  const meter = (arr, k, item) => {
    if (arr.length >= k && item.s <= arr[arr.length - 1].s) return;
    arr.push(item);
    arr.sort((a, b) => b.s - a.s);
    if (arr.length > k) arr.pop();
  };
  for (const m of indice.entradas) {
    if (q.n) {
      const piso = topN.length < kNombre ? 0.3 : topN[topN.length - 1].s;
      const sN = nameScore(q, m._name, piso);
      if (sN >= 0.3) meter(topN, kNombre, { m, s: sN });
    }
    if (dni.length >= 6 && m.dni) {
      const sD = docScore(dni, m.dni);
      if (sD >= 0.6) meter(topD, kDoc, { m, s: sD });
    }
  }
  const vistos = new Set();
  return [...topN, ...topD].map((x) => x.m).filter((m) => !vistos.has(m._i) && vistos.add(m._i));
}

// convierte el resultado del cruce en la fila que se muestra y se copia a Excel
function aFila(r, e) {
  return {
    dni: cleanDni(r.dni),
    nombre: r.nombre || "",
    cargo: r.cargo || "",
    fechaIngreso: r.fechaIngreso || "",
    fechaNacimiento: r.fechaNacimiento || "",
    celular: r.celular || "",
    frente: r.frente || "",
    status: r.estado,
    obs: r.obs || "",
    dudoso: !!r.dudoso,
    raw: { nombre: e.nombre, dni: e.dni, cargo: e.cargo || "" },
  };
}

// misma persona marcada dos veces en el registro (o dos filas que cayeron en la misma persona del maestro)
export function marcarDuplicados(rows) {
  const porDni = new Map();
  rows.forEach((r, i) => {
    if (r.status === "NO ENCONTRADO" || !r.dni) return;
    const l = porDni.get(r.dni);
    if (l) l.push(i); else porDni.set(r.dni, [i]);
  });
  for (const filas of porDni.values()) {
    if (filas.length < 2) continue;
    const txt = `Aparece ${filas.length} veces en el registro (filas ${filas.map((i) => i + 1).join(", ")}).`;
    for (const i of filas) rows[i].obs = rows[i].obs ? `${rows[i].obs} ${txt}` : txt;
  }
  return rows;
}

/* ---------- preparación de archivos pesados: imágenes comprimidas y PDF en partes ---------- */

class DemasiadoGrande extends Error {}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = () => reject(new Error("No se pudo leer el archivo en el navegador."));
    r.readAsDataURL(blob);
  });
}

const TIPOS_IMAGEN_API = ["image/jpeg", "image/png", "image/gif", "image/webp"];

async function prepararImagen(file) {
  let bmp = null;
  try {
    bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
  } catch {
    bmp = null;
  }
  if (!bmp) {
    if (TIPOS_IMAGEN_API.includes(file.type) && file.size <= MAX_BYTES_IMAGEN) {
      return { base64: await blobToBase64(file), mediaType: file.type };
    }
    throw new Error(`No se pudo abrir la imagen (formato ${file.type || "desconocido"}). Conviértela a JPG o PNG.`);
  }
  const escala0 = Math.min(1, MAX_DIM_IMAGEN / Math.max(bmp.width, bmp.height));
  if (escala0 === 1 && TIPOS_IMAGEN_API.includes(file.type) && file.size <= MAX_BYTES_IMAGEN) {
    bmp.close?.();
    return { base64: await blobToBase64(file), mediaType: file.type };
  }
  // se reduce y se convierte a JPEG; si aún pesa mucho, se baja calidad y tamaño
  let escala = escala0, calidad = 0.9, blob = null;
  for (let intento = 0; intento < 5; intento++) {
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(bmp.width * escala));
    canvas.height = Math.max(1, Math.round(bmp.height * escala));
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
    blob = await new Promise((r) => canvas.toBlob(r, "image/jpeg", calidad));
    if (blob && blob.size <= MAX_BYTES_IMAGEN) break;
    calidad = Math.max(0.6, calidad - 0.1);
    escala *= 0.85;
  }
  bmp.close?.();
  if (!blob || blob.size > MAX_BYTES_IMAGEN) throw new Error("La imagen sigue siendo demasiado pesada después de comprimirla.");
  return { base64: await blobToBase64(blob), mediaType: "image/jpeg" };
}

// pdf-lib se carga desde cdnjs solo cuando hace falta dividir un PDF
let pdfLibPromise = null;
function cargarPdfLib() {
  if (window.PDFLib) return Promise.resolve(window.PDFLib);
  if (!pdfLibPromise) {
    pdfLibPromise = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js";
      s.onload = () => (window.PDFLib ? resolve(window.PDFLib) : reject(new Error("pdf-lib no disponible")));
      s.onerror = () => reject(new Error("No se pudo cargar pdf-lib"));
      document.head.appendChild(s);
    }).catch((e) => {
      pdfLibPromise = null;
      throw e;
    });
  }
  return pdfLibPromise;
}

function trabajoArchivoCompleto(file, etiqueta, mediaType) {
  return {
    etiqueta,
    contexto: "",
    obtenerDatos: async () => {
      if (file.size > MAX_BYTES_POR_PARTE) throw new DemasiadoGrande(`pesa ${formatoMB(file.size)}`);
      return { base64: await blobToBase64(file), mediaType };
    },
    dividir: null,
  };
}

function trabajoPdfRango(PDFLib, src, file, desde, hasta, total) {
  const completo = desde === 0 && hasta === total;
  const etiqueta = completo ? file.name : `${file.name} (págs. ${desde + 1}–${hasta} de ${total})`;
  return {
    etiqueta,
    contexto: completo ? "" : `Este archivo contiene las páginas ${desde + 1} a ${hasta} de un documento de ${total} páginas.`,
    obtenerDatos: async () => {
      if (completo && file.size <= MAX_BYTES_POR_PARTE) {
        return { base64: await blobToBase64(file), mediaType: "application/pdf" };
      }
      const doc = await PDFLib.PDFDocument.create();
      const idx = Array.from({ length: hasta - desde }, (_, i) => desde + i);
      const paginas = await doc.copyPages(src, idx);
      paginas.forEach((p) => doc.addPage(p));
      const bytes = await doc.save();
      if (bytes.length > MAX_BYTES_POR_PARTE) throw new DemasiadoGrande(`la parte pesa ${formatoMB(bytes.length)}`);
      return { base64: await blobToBase64(new Blob([bytes], { type: "application/pdf" })), mediaType: "application/pdf" };
    },
    dividir:
      hasta - desde > 1
        ? async () => {
            const mitad = desde + Math.ceil((hasta - desde) / 2);
            return [
              trabajoPdfRango(PDFLib, src, file, desde, mitad, total),
              trabajoPdfRango(PDFLib, src, file, mitad, hasta, total),
            ];
          }
        : null,
  };
}

async function trabajosPdf(file, addLog) {
  let PDFLib = null, src = null;
  try {
    PDFLib = await cargarPdfLib();
    src = await PDFLib.PDFDocument.load(await file.arrayBuffer(), { ignoreEncryption: true });
  } catch {
    addLog(`⚠ No se pudo dividir "${file.name}" en partes; se enviará completo.`);
    return [trabajoArchivoCompleto(file, file.name, "application/pdf")];
  }
  const total = src.getPageCount();
  if (total <= PAGINAS_POR_PARTE) return [trabajoPdfRango(PDFLib, src, file, 0, total, total)];
  const out = [];
  for (let d = 0; d < total; d += PAGINAS_POR_PARTE) {
    out.push(trabajoPdfRango(PDFLib, src, file, d, Math.min(total, d + PAGINAS_POR_PARTE), total));
  }
  addLog(`"${file.name}": ${total} páginas, se lee en ${out.length} partes.`);
  return out;
}

async function prepararTrabajos(files, addLog) {
  const trabajos = [];
  for (const file of files) {
    const esPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name);
    if (esPdf) {
      trabajos.push(...(await trabajosPdf(file, addLog)));
    } else {
      trabajos.push({ etiqueta: file.name, contexto: "", obtenerDatos: () => prepararImagen(file), dividir: null });
    }
  }
  return trabajos;
}

/* ---------- llamada a la API: SOLO lee el registro (el cruce se hace después, en código, no en el modelo) ---------- */

const PROMPT_OCR = `El archivo adjunto (imagen o PDF, puede tener varias páginas) contiene un registro de asistencia/inducción/capacitación manuscrito, y posiblemente TAMBIÉN otras páginas ajenas (exámenes, evaluaciones, anexos con preguntas y respuestas marcadas).

Identifica y usa ÚNICAMENTE la(s) página(s) que correspondan a la tabla de asistentes de un registro de asistencia/inducción/capacitación. Ignora POR COMPLETO cualquier página de examen, evaluación u otro anexo: no la leas, no extraigas nada de ella, no la menciones ni la cuentes como fila.

De la tabla de asistentes identificada, extrae cada fila (Apellidos, Nombres, DNI/CE y Cargo/Puesto si la columna existe), tal como está escrita a mano, aunque sospeches error de escritura o de lectura. Si una celda está ilegible, deja el valor como cadena vacía. No inventes filas ni datos. No intentes corregir ni adivinar el documento, el nombre o el cargo: transcribe literalmente lo que está escrito, la corrección contra el maestro se hace en otro paso.

Devuelve el resultado como TEXTO PLANO, una fila por línea, con 3 campos separados por un TABULADOR (no uses JSON, no uses markdown, no uses comillas, no numeres las filas, no agregues encabezado ni texto adicional antes o después). Si la tabla no tiene columna de Cargo/Puesto, deja ese tercer campo vacío (igual deja el tabulador):
DOCUMENTO_LEIDO<TAB>NOMBRE_LEIDO<TAB>CARGO_LEIDO

Si en el archivo no hay ninguna tabla de asistentes, responde únicamente: VACIO`;

class ApiError extends Error {
  constructor(message, status, reintentable) {
    super(message);
    this.status = status;
    this.reintentable = reintentable;
  }
}

function esperar(ms, signal) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener("abort", () => { clearTimeout(t); reject(new DOMException("Cancelado", "AbortError")); }, { once: true });
  });
}

// reintenta con espera creciente ante saturación (429/529), errores 5xx, cortes de red y tiempos de espera
async function llamarApi(body, signal) {
  let ultimo = null;
  for (let intento = 0; intento < REINTENTOS; intento++) {
    if (signal?.aborted) throw new DOMException("Cancelado", "AbortError");
    const ctrl = new AbortController();
    const onAbort = () => ctrl.abort();
    signal?.addEventListener("abort", onAbort);
    const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal: ctrl.signal,
      });
      let data = null;
      try { data = await res.json(); } catch { data = null; }
      if (res.ok && data && !data.error) return data;
      const msg = data?.error?.message || `Error HTTP ${res.status}`;
      ultimo = new ApiError(msg, res.status, res.status === 429 || res.status >= 500 || res.status === 0);
    } catch (e) {
      if (signal?.aborted) throw new DOMException("Cancelado", "AbortError");
      ultimo = new ApiError(
        e.name === "AbortError" ? "Tiempo de espera agotado" : `Falló la conexión con Claude (${e.message || "error de red"})`,
        0,
        true
      );
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
    }
    if (!ultimo.reintentable || intento === REINTENTOS - 1) break;
    // ante cortes de conexión se reintenta una sola vez: si persiste, quien llama prueba con partes más chicas
    if (ultimo.status === 0 && intento >= 1) break;
    await esperar(Math.min(30000, 2000 * 2 ** intento) + Math.random() * 500, signal);
  }
  throw ultimo;
}

const ES_ENCABEZADO = /^(DNI|DOC|DOCUMENTO|N[°º]?|CE)\b/i;

export function parsearRespuestaOcr(text, truncated) {
  let lines = text
    .split("\n")
    .map((l) => l.replace(/\r$/, ""))
    .filter((l) => l.trim().length > 0 && l.includes("\t"));
  if (truncated && lines.length > 0) lines = lines.slice(0, -1); // la última línea puede venir incompleta
  const rows = [];
  for (const line of lines) {
    let campos = line.split("\t").map((c) => c.trim());
    // si el modelo numeró las filas igual, se quita el número
    if (campos.length >= 4 && /^\d{1,3}$/.test(campos[0])) campos = campos.slice(1);
    const [dni = "", nombre = "", cargo = ""] = campos;
    if (ES_ENCABEZADO.test(dni) && /APELLIDO|NOMBRE/i.test(nombre)) continue; // fila de encabezado
    if (!dni && !nombre) continue;
    rows.push({ dni, nombre, cargo });
  }
  return rows;
}

async function ocrParte({ base64, mediaType }, contexto, signal) {
  const fileBlock =
    mediaType === "application/pdf"
      ? { type: "document", source: { type: "base64", media_type: mediaType, data: base64 } }
      : { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } };
  const data = await llamarApi(
    {
      model: MODELO,
      max_tokens: MAX_TOKENS_RESPUESTA,
      temperature: 0,
      messages: [{ role: "user", content: [fileBlock, { type: "text", text: contexto ? `${PROMPT_OCR}\n\nNota: ${contexto}` : PROMPT_OCR }] }],
    },
    signal
  );
  if (data.stop_reason === "refusal") throw new ApiError("El modelo no procesó este archivo.", 200, false);
  const text = (data.content || []).map((b) => (b.type === "text" ? b.text || "" : "")).join("\n");
  const truncated = data.stop_reason === "max_tokens";
  return { rows: parsearRespuestaOcr(text, truncated), truncated };
}

const esMuyGrande = (e) =>
  e instanceof DemasiadoGrande || e.status === 413 || /too large|exceed|maximum|demasiado/i.test(e.message || "");

// lee una parte; si es muy pesada o la respuesta se corta, la divide en dos y lee cada mitad
async function procesarTrabajo(job, signal, addLog) {
  try {
    const datos = await job.obtenerDatos();
    const r = await ocrParte(datos, job.contexto, signal);
    if (r.truncated && job.dividir) {
      addLog(`${job.etiqueta}: la respuesta se cortó, se vuelve a leer en dos mitades…`);
      return combinar(await Promise.all((await job.dividir()).map((s) => procesarTrabajo(s, signal, addLog))));
    }
    return {
      rows: r.rows.map((x) => ({ ...x, _job: job })),
      avisos: r.truncated ? [`${job.etiqueta}: la respuesta se cortó; puede faltar la última fila.`] : [],
      errores: [],
    };
  } catch (e) {
    if (e.name === "AbortError") throw e;
    // un corte de conexión que se repite en cada reintento suele deberse a un envío demasiado pesado
    // (el entorno de los artefactos acepta menos que la API): se prueba con partes más chicas
    const fallaRed = e instanceof ApiError && e.status === 0;
    if ((esMuyGrande(e) || fallaRed) && job.dividir) {
      addLog(`${job.etiqueta}: ${fallaRed ? "falló la conexión" : "muy pesado para un solo envío"}; se reintenta en dos partes más pequeñas…`);
      return combinar(await Promise.all((await job.dividir()).map((s) => procesarTrabajo(s, signal, addLog))));
    }
    const detalle = e instanceof ApiError && e.status ? ` (HTTP ${e.status})` : "";
    const sugerencia =
      esMuyGrande(e) || fallaRed
        ? " Si se repite, el archivo puede ser muy pesado: escanéalo a menor resolución (150–200 ppp) o en blanco y negro."
        : "";
    return { rows: [], avisos: [], errores: [`${job.etiqueta}: ${e.message}${detalle}.${sugerencia}`] };
  }
}

/* ---------- segunda lectura de filas dudosas ---------- */

function promptVerificacion(items) {
  const lista = items
    .map((it) => {
      const cands = it.candidatos
        .map((m, j) => `   ${j + 1}) DNI ${m.dni || "(sin doc)"} · ${m.nombre}${m.cargo ? ` · ${m.cargo}` : ""}`)
        .join("\n");
      return `[F${it.id}] se leyó: DNI "${it.e.dni}" / nombre "${it.e.nombre}"${it.e.cargo ? ` / cargo "${it.e.cargo}"` : ""}\n${cands || "   (sin candidatos parecidos)"}`;
    })
    .join("\n\n");
  return `Te adjunto otra vez el registro de asistencia manuscrito. Algunas filas que se leyeron de él no se pudieron cruzar con seguridad contra la fuerza laboral.

Para cada fila de la lista: ubícala en la tabla de asistentes del documento, vuelve a leer con cuidado el DNI (dígito por dígito) y el nombre tal como están escritos, y compáralos con los candidatos de la fuerza laboral que aparecen debajo de la fila.

Elige un candidato SOLO si lo escrito en el documento respalda que es la misma persona (el DNI coincide o casi coincide, y/o los apellidos y nombres corresponden aunque la letra sea difícil). No elijas un candidato solo porque es el más parecido de la lista. Si ninguno corresponde, responde NINGUNO.

Responde en TEXTO PLANO, una línea por fila, con 4 campos separados por TABULADOR, sin encabezado ni texto adicional:
ID<TAB>OPCION<TAB>DNI_RELEIDO<TAB>NOMBRE_RELEIDO
donde ID es el código de la fila (por ejemplo F3), OPCION es el número del candidato elegido o NINGUNO, y DNI_RELEIDO / NOMBRE_RELEIDO son tu nueva lectura literal del documento.

Filas a revisar:

${lista}`;
}

export function parsearVerificacion(text) {
  const out = new Map();
  for (const line of text.split("\n")) {
    const c = line.split("\t").map((x) => x.trim());
    const id = /^\[?F(\d+)\]?$/i.exec(c[0] || "");
    if (!id || c.length < 2) continue;
    const opcion = /^\d+$/.test(c[1]) ? Number(c[1]) : null;
    out.set(Number(id[1]), { opcion, dni: c[2] || "", nombre: c[3] || "" });
  }
  return out;
}

async function verificarFilas(job, items, signal) {
  const datos = await job.obtenerDatos();
  const fileBlock =
    datos.mediaType === "application/pdf"
      ? { type: "document", source: { type: "base64", media_type: datos.mediaType, data: datos.base64 } }
      : { type: "image", source: { type: "base64", media_type: datos.mediaType, data: datos.base64 } };
  const data = await llamarApi(
    {
      model: MODELO,
      max_tokens: MAX_TOKENS_RESPUESTA,
      temperature: 0,
      messages: [{ role: "user", content: [fileBlock, { type: "text", text: promptVerificacion(items) }] }],
    },
    signal
  );
  const text = (data.content || []).map((b) => (b.type === "text" ? b.text || "" : "")).join("\n");
  return parsearVerificacion(text);
}

// aplica la respuesta de la segunda lectura a una fila dudosa; nunca inventa: solo acepta personas del maestro
export function aplicarVerificacion(fila, it, v, indice) {
  const nota = (t) => (fila.obs ? `${fila.obs} ${t}` : t);
  if (!v) return { ...fila, obs: nota("La segunda lectura no devolvió respuesta para esta fila.") };
  if (v.opcion && it.candidatos[v.opcion - 1]) {
    const m = it.candidatos[v.opcion - 1];
    return aFila(
      filaDesdeMaestro(m, "CORREGIDO", [
        `Confirmado en segunda lectura del documento (se releyó "${v.dni}" / "${v.nombre}"; en la primera lectura: "${it.e.dni}" / "${it.e.nombre}").`,
        notaEstado(m),
      ]),
      it.e
    );
  }
  // ninguno de los candidatos: se vuelve a cruzar con la nueva lectura, por si el DNI o el nombre se leyeron mal la primera vez
  const cambio = onlyDigitsAlnum(v.dni) !== onlyDigitsAlnum(it.e.dni) || normName(v.nombre) !== normName(it.e.nombre);
  if (cambio && (v.dni || v.nombre)) {
    const r2 = matchAgainstMaster({ dni: v.dni, nombre: v.nombre, cargo: it.e.cargo }, indice);
    if (r2.estado !== "NO ENCONTRADO" && !r2.dudoso) {
      return aFila({ ...r2, estado: "CORREGIDO", obs: `Encontrado con la segunda lectura del documento ("${v.dni}" / "${v.nombre}"). ${r2.obs}`.trim() }, it.e);
    }
  }
  if (fila.status === "NO ENCONTRADO") {
    return {
      ...fila,
      dni: cleanDni(v.dni || fila.dni),
      nombre: v.nombre || fila.nombre,
      obs: `Revisado dos veces contra el documento: ninguna persona de la fuerza laboral corresponde. Probablemente no está registrado en el maestro (alta reciente o en otra hoja); verificar.${
        cambio ? ` Segunda lectura: "${v.dni}" / "${v.nombre}".` : ""
      }`,
    };
  }
  return { ...fila, obs: nota("La segunda lectura no confirmó ningún candidato; verificar manualmente.") };
}

function combinar(resultados) {
  return {
    rows: resultados.flatMap((r) => r.rows),
    avisos: resultados.flatMap((r) => r.avisos),
    errores: resultados.flatMap((r) => r.errores),
  };
}

// ejecuta fn sobre items con un máximo de "limite" a la vez, conservando el orden
async function enParalelo(items, limite, fn) {
  const out = new Array(items.length);
  let sig = 0;
  const worker = async () => {
    while (sig < items.length) {
      const i = sig++;
      out[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limite, items.length) }, worker));
  return out;
}

/* ---------- UI ---------- */

const COLORS = {
  bg: "#1B1D21",
  panel: "#222529",
  panelAlt: "#26292E",
  border: "#34383F",
  text: "#ECE9E2",
  dim: "#9BA0A8",
  amber: "#F0A631",
  amberDim: "#5A4520",
  green: "#5EA97A",
  red: "#E2574C",
};

function StatusPill({ status }) {
  const map = {
    OK: { c: COLORS.green, l: "OK" },
    CORREGIDO: { c: COLORS.amber, l: "CORREGIDO" },
    "NO ENCONTRADO": { c: COLORS.red, l: "NO ENCONTRADO" },
    REVISAR: { c: COLORS.red, l: "REVISAR" },
  };
  const s = map[status] || map["NO ENCONTRADO"];
  return (
    <span
      style={{
        color: s.c,
        border: `1px solid ${s.c}66`,
        background: `${s.c}14`,
        fontFamily: "'IBM Plex Mono', monospace",
        fontSize: "11px",
        letterSpacing: "0.06em",
        padding: "2px 8px",
        borderRadius: "3px",
        whiteSpace: "nowrap",
      }}
    >
      {s.l}
    </span>
  );
}

function UploadSlot({ label, hint, accept, files, onFiles, icon, multiple = false, disabled = false }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const entregar = (lista) => {
    const arr = Array.from(lista || []);
    if (arr.length) onFiles(multiple ? arr : arr.slice(0, 1));
  };
  const total = files ? files.reduce((s, f) => s + f.size, 0) : 0;
  const nombre = !files || files.length === 0 ? null : files.length === 1 ? files[0].name : `${files.length} archivos`;
  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        if (!disabled) entregar(e.dataTransfer.files);
      }}
      onClick={() => !disabled && inputRef.current.click()}
      style={{
        flex: 1,
        minHeight: "148px",
        border: `1px dashed ${dragOver ? COLORS.amber : COLORS.border}`,
        background: dragOver ? `${COLORS.amber}0d` : COLORS.panel,
        borderRadius: "6px",
        padding: "20px",
        cursor: disabled ? "wait" : "pointer",
        opacity: disabled ? 0.6 : 1,
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        transition: "border-color 120ms, background 120ms",
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        multiple={multiple}
        style={{ display: "none" }}
        onChange={(e) => { entregar(e.target.files); e.target.value = ""; }}
      />
      <div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, letterSpacing: "0.08em", marginBottom: "6px" }}>
          {icon} {label.toUpperCase()}
        </div>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: "15px", color: COLORS.text, wordBreak: "break-word" }}>
          {nombre ? `${nombre} · ${formatoMB(total)}` : hint}
        </div>
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: nombre ? COLORS.green : COLORS.dim }}>
        {nombre ? "✓ cargado — clic para reemplazar" : multiple ? "clic o arrastra uno o varios archivos aquí" : "clic o arrastra el archivo aquí"}
      </div>
    </div>
  );
}

export default function App() {
  const tsvRef = useRef(null);
  const abortRef = useRef(null);
  const [regFiles, setRegFiles] = useState([]);
  const [xlsxFile, setXlsxFile] = useState(null);
  const [master, setMaster] = useState(null);
  const [sheetsFound, setSheetsFound] = useState([]);
  const [conteoHojas, setConteoHojas] = useState([]);
  const [savedInfo, setSavedInfo] = useState(null); // {name, savedAt}
  const [loadingStored, setLoadingStored] = useState(true);
  const [leyendoExcel, setLeyendoExcel] = useState(false);
  const [avisoGuardado, setAvisoGuardado] = useState(null);
  const [log, setLog] = useState([]);
  const [rows, setRows] = useState(null);
  const [detail, setDetail] = useState(false);
  const [segundaLectura, setSegundaLectura] = useState(true);
  const [camposSeleccionados, setCamposSeleccionados] = useState(["nombre", "dni"]);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  const [copied, setCopied] = useState(false);

  const toggleCampo = (key) => {
    setCamposSeleccionados((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]));
  };

  // el índice se arma una vez por maestro, no en cada procesamiento
  const indice = useMemo(() => (master ? indexarMaestro(master) : null), [master]);

  // al abrir la app, intenta recuperar la última fuerza laboral guardada
  useEffect(() => {
    (async () => {
      try {
        const parsed = await cargarMaestro();
        if (parsed && parsed.master) {
          setMaster(parsed.master);
          setSheetsFound(parsed.sheetsFound || []);
          setConteoHojas(parsed.conteo || []);
          setSavedInfo({ name: parsed.name, savedAt: parsed.savedAt });
        }
      } catch {
        // no había nada guardado todavía, o falló la lectura: se pide subir manualmente
      } finally {
        setLoadingStored(false);
      }
    })();
  }, []);

  const handleReg = useCallback((files) => {
    setRegFiles(files);
    setRows(null);
    setErr(null);
  }, []);

  const handleXlsx = useCallback(async (files) => {
    const file = files[0];
    setXlsxFile(file);
    setRows(null);
    setErr(null);
    setAvisoGuardado(null);
    setLeyendoExcel(true);
    await tick(); // deja que se pinte "Leyendo Excel…" antes del trabajo pesado
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const { master: m, sheetsFound: sf, conteo, hojasDisponibles } = leerFuerzaLaboral(data);
      if (m.length === 0) {
        setXlsxFile(null);
        setErr(
          sf.length === 0
            ? `No se encontraron las hojas DATA / DADO DE BAJA / VISITA. Hojas del archivo: ${hojasDisponibles.join(", ") || "ninguna"}.`
            : "Se encontraron las hojas, pero no filas válidas. Verifica que tengan columnas de DNI y Apellidos y Nombres en las primeras 30 filas."
        );
        return;
      }
      setMaster(m);
      setSheetsFound(sf);
      setConteoHojas(conteo);
      const savedAt = new Date().toLocaleString("es-PE");
      setSavedInfo({ name: file.name, savedAt });
      try {
        await guardarMaestro({ master: m, sheetsFound: sf, conteo, name: file.name, savedAt });
      } catch (storageErr) {
        console.error("No se pudo guardar la fuerza laboral:", storageErr);
        setAvisoGuardado("No se pudo guardar la fuerza laboral para la próxima vez (el archivo es muy grande para el almacenamiento). Se usará solo en esta sesión.");
      }
    } catch (e) {
      setXlsxFile(null);
      setErr("No se pudo leer el archivo (.xlsx/.xlsb): " + e.message);
    } finally {
      setLeyendoExcel(false);
    }
  }, []);

  const addLog = useCallback((msg) => setLog((l) => [...l, msg]), []);

  const cancelar = () => abortRef.current?.abort();

  const procesar = async () => {
    if (regFiles.length === 0 || !indice) return;
    const ctrl = new AbortController();
    abortRef.current = ctrl;
    setBusy(true); setErr(null); setRows(null); setLog([]);
    try {
      addLog("Preparando archivos…");
      const trabajos = await prepararTrabajos(regFiles, addLog);
      addLog(`Leyendo tabla del registro manuscrito (${trabajos.length} parte(s), hasta ${PARTES_EN_PARALELO} a la vez)…`);

      let hechas = 0;
      const resultados = await enParalelo(trabajos, PARTES_EN_PARALELO, async (t) => {
        const r = await procesarTrabajo(t, ctrl.signal, addLog);
        hechas++;
        addLog(`Parte ${hechas}/${trabajos.length} lista — ${t.etiqueta}: ${r.rows.length} fila(s)${r.errores.length ? " (con errores)" : ""}.`);
        return r;
      });
      const { rows: extracted, avisos, errores } = combinar(resultados);
      avisos.forEach((a) => addLog("⚠ " + a));

      if (extracted.length === 0) {
        setErr(
          errores.length
            ? `No se pudo leer el registro:\n${errores.join("\n")}`
            : "No se pudo leer ninguna fila del registro. Verifica que el archivo tenga una tabla de asistentes legible."
        );
        return;
      }
      addLog(`${extracted.length} fila(s) detectada(s) en el registro.`);
      addLog(`Cruzando contra fuerza laboral (${sheetsFound.join(", ") || "sin hojas reconocidas"}, ${master.length} personas)…`);

      let matched = [];
      for (let i = 0; i < extracted.length; i++) {
        matched.push(aFila(matchAgainstMaster(extracted[i], indice), extracted[i]));
        if (i % 25 === 24) await tick(); // mantiene la página respondiendo con registros largos
      }

      // segunda lectura: las filas dudosas se vuelven a revisar contra el documento original,
      // mostrándole al modelo los candidatos reales de la fuerza laboral para que elija (o diga que ninguno)
      const dudosas = matched.map((f, i) => i).filter((i) => matched[i].dudoso && extracted[i]._job);
      if (segundaLectura && dudosas.length) {
        addLog(`Segunda lectura de ${dudosas.length} fila(s) dudosa(s) contra el documento…`);
        const porParte = new Map();
        for (const i of dudosas) {
          const it = { id: i + 1, idx: i, e: extracted[i], candidatos: candidatosPara(extracted[i], indice) };
          const l = porParte.get(extracted[i]._job);
          if (l) l.push(it); else porParte.set(extracted[i]._job, [it]);
          await tick();
        }
        const tareas = [];
        for (const [job, items] of porParte) for (let k = 0; k < items.length; k += 25) tareas.push([job, items.slice(k, k + 25)]);
        let resueltas = 0;
        await enParalelo(tareas, PARTES_EN_PARALELO, async ([job, items]) => {
          let resp = null;
          try {
            resp = await verificarFilas(job, items, ctrl.signal);
          } catch (e) {
            if (e.name === "AbortError") throw e;
            addLog(`⚠ ${job.etiqueta}: no se pudo hacer la segunda lectura (${e.message}).`);
            return;
          }
          for (const it of items) {
            const nueva = aplicarVerificacion(matched[it.idx], it, resp.get(it.id), indice);
            if (nueva.status !== matched[it.idx].status) resueltas++;
            matched[it.idx] = nueva;
          }
        });
        addLog(`Segunda lectura completada: ${resueltas} fila(s) resuelta(s).`);
      }
      marcarDuplicados(matched);
      addLog("Cruce completado.");
      setRows(matched);

      if (errores.length) {
        setErr(`Se procesaron ${extracted.length} fila(s), pero algunas partes no se pudieron leer y pueden faltar filas:\n${errores.join("\n")}`);
      }
    } catch (e) {
      if (e.name === "AbortError") {
        addLog("Procesamiento cancelado.");
      } else {
        setErr(e.message);
        addLog("Error: " + e.message);
      }
    } finally {
      abortRef.current = null;
      setBusy(false);
    }
  };

  const camposActivos = useMemo(
    () => camposSeleccionados.map((key) => CAMPOS_DISPONIBLES.find((c) => c.key === key)).filter(Boolean),
    [camposSeleccionados]
  );
  const tsv = useMemo(
    () =>
      rows
        ? rows
            .map((r) => {
              const base = camposActivos.map((c) => r[c.key] || "");
              const extra = detail ? [r.status, r.obs] : [];
              return [...base, ...extra].join("\t");
            })
            .join("\n")
        : "",
    [rows, camposActivos, detail]
  );
  const flagged = useMemo(() => (rows ? rows.filter((r) => r.status !== "OK" || r.obs) : []), [rows]);
  const resumen = useMemo(() => {
    if (!rows) return null;
    const c = { OK: 0, CORREGIDO: 0, REVISAR: 0, "NO ENCONTRADO": 0 };
    rows.forEach((r) => { c[r.status] = (c[r.status] || 0) + 1; });
    return c;
  }, [rows]);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(tsv);
      setCopied(true);
      setTimeout(() => setCopied(false), 1600);
      return;
    } catch {
      // algunos navegadores bloquean la Clipboard API; respaldo abajo
    }
    try {
      const el = tsvRef.current;
      el.focus();
      el.select();
      document.execCommand("copy");
      setCopied(true);
      setTimeout(() => setCopied(false), 1600);
    } catch {}
  };

  const puedeProcesar = regFiles.length > 0 && !!indice && !busy && !leyendoExcel;

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, color: COLORS.text, fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        button:focus-visible, div[tabindex]:focus-visible { outline: 2px solid #F0A631; outline-offset: 2px; }
      `}</style>

      <div style={{ maxWidth: "920px", margin: "0 auto", padding: "40px 24px 80px" }}>
        {/* header */}
        <div style={{ marginBottom: "36px", borderBottom: `1px solid ${COLORS.border}`, paddingBottom: "20px" }}>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.amber, letterSpacing: "0.12em", marginBottom: "8px" }}>
            PROTOTIPO · RECONOCIMIENTO DE REGISTROS
          </div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: "28px", fontWeight: 600, margin: 0, letterSpacing: "-0.01em" }}>
            Registro manuscrito → Fuerza laboral
          </h1>
          <p style={{ color: COLORS.dim, fontSize: "14px", marginTop: "8px", maxWidth: "560px" }}>
            Sube el registro escaneado y el Excel de fuerza laboral. El sistema lee la tabla, la cruza contra "DATA" / "DADO DE BAJA" / "VISITA" y entrega un bloque listo para pegar en Excel.
          </p>
        </div>

        {/* uploads */}
        <div style={{ display: "flex", flexWrap: "wrap", gap: "16px", marginBottom: "12px" }}>
          <UploadSlot
            label="Registro del día"
            hint="Imágenes, escaneos o PDF (uno o varios; si traen exámenes u otros anexos, se ignoran)"
            accept="image/*,application/pdf"
            files={regFiles}
            onFiles={handleReg}
            icon="①"
            multiple
            disabled={busy}
          />
          <UploadSlot
            label="Fuerza laboral"
            hint={savedInfo ? "Ya guardada — sube un archivo solo si quieres actualizarla" : "Excel maestro (.xlsx o .xlsb)"}
            accept=".xlsx,.xls,.xlsb"
            files={xlsxFile ? [xlsxFile] : []}
            onFiles={handleXlsx}
            icon="②"
            disabled={busy || leyendoExcel}
          />
        </div>

        {loadingStored ? (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.dim, marginBottom: "20px" }}>
            Buscando fuerza laboral guardada…
          </div>
        ) : leyendoExcel ? (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.amber, marginBottom: "20px" }}>
            Leyendo Excel… (los archivos grandes pueden tardar unos segundos)
          </div>
        ) : savedInfo && !xlsxFile ? (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.green, marginBottom: "20px" }}>
            ✓ Usando fuerza laboral guardada: {savedInfo.name} (cargada el {savedInfo.savedAt}) — {conteoHojas.length > 0 ? conteoHojas.join(", ") : `${master ? master.length : 0} registro(s), hojas: ${sheetsFound.join(", ") || "—"}`}
          </div>
        ) : master ? (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.dim, marginBottom: "20px" }}>
            Hojas reconocidas: {conteoHojas.join(", ")} — {master.length} registro(s) en total.{avisoGuardado ? "" : " Se guardó para la próxima vez."}
          </div>
        ) : null}

        {avisoGuardado && (
          <div style={{ fontSize: "12px", color: COLORS.amber, marginTop: "-12px", marginBottom: "20px" }}>⚠ {avisoGuardado}</div>
        )}

        {/* selector de columnas: se marca ANTES de procesar */}
        <div
          style={{
            border: `1px solid ${COLORS.border}`,
            borderRadius: "6px",
            padding: "12px 14px",
            marginBottom: "16px",
            background: COLORS.panel,
          }}
        >
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, letterSpacing: "0.06em", marginBottom: "8px" }}>
            DATOS A INCLUIR (tomados de la fuerza laboral) — se generan en el orden en que los marques
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "14px" }}>
            {CAMPOS_DISPONIBLES.map((c) => {
              const posicion = camposSeleccionados.indexOf(c.key);
              return (
                <label
                  key={c.key}
                  style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12.5px", color: COLORS.text, cursor: "pointer" }}
                >
                  <input type="checkbox" checked={posicion !== -1} onChange={() => toggleCampo(c.key)} />
                  {c.label}
                  {posicion !== -1 && (
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "10px", color: COLORS.amber }}>#{posicion + 1}</span>
                  )}
                </label>
              );
            })}
          </div>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: "8px", fontSize: "12.5px", color: COLORS.text, cursor: "pointer", marginBottom: "12px" }}>
          <input type="checkbox" checked={segundaLectura} onChange={(e) => setSegundaLectura(e.target.checked)} disabled={busy} />
          Segunda lectura de filas dudosas
          <span style={{ color: COLORS.dim }}>— vuelve a revisar el documento con los candidatos de la fuerza laboral (más preciso, tarda un poco más)</span>
        </label>

        {/* action */}
        <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
          <button
            onClick={procesar}
            disabled={!puedeProcesar}
            style={{
              flex: 1,
              padding: "14px",
              borderRadius: "6px",
              border: "none",
              background: puedeProcesar ? COLORS.amber : COLORS.panelAlt,
              color: puedeProcesar ? "#1B1D21" : COLORS.dim,
              fontFamily: "'Space Grotesk', sans-serif",
              fontWeight: 600,
              fontSize: "14px",
              cursor: puedeProcesar ? "pointer" : "not-allowed",
            }}
          >
            {busy ? "Procesando…" : "Procesar registro"}
          </button>
          {busy && (
            <button
              onClick={cancelar}
              style={{
                padding: "14px 18px",
                borderRadius: "6px",
                border: `1px solid ${COLORS.border}`,
                background: "none",
                color: COLORS.text,
                fontFamily: "'Space Grotesk', sans-serif",
                fontSize: "14px",
                cursor: "pointer",
              }}
            >
              Cancelar
            </button>
          )}
        </div>

        {/* log */}
        {log.length > 0 && (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "12px", color: COLORS.dim, marginBottom: "24px", lineHeight: 1.8 }}>
            {log.map((l, i) => <div key={i}>› {l}</div>)}
          </div>
        )}

        {err && (
          <div style={{ border: `1px solid ${COLORS.red}66`, background: `${COLORS.red}14`, color: COLORS.red, padding: "12px 16px", borderRadius: "6px", fontSize: "13px", marginBottom: "24px", whiteSpace: "pre-wrap" }}>
            {err}
          </div>
        )}

        {/* results */}
        {rows && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px", marginBottom: "10px" }}>
              <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: "16px", fontWeight: 600, margin: 0 }}>
                Resultado — {rows.length} fila(s)
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, fontWeight: 400, marginLeft: "10px" }}>
                  {resumen.OK} OK · {resumen.CORREGIDO} corregido(s) · {resumen.REVISAR ? `${resumen.REVISAR} a revisar · ` : ""}{resumen["NO ENCONTRADO"]} no encontrado(s)
                </span>
              </h2>
              <label style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, display: "flex", alignItems: "center", gap: "6px", cursor: "pointer" }}>
                <input type="checkbox" checked={detail} onChange={(e) => setDetail(e.target.checked)} />
                agregar estado/observación
              </label>
            </div>

            {/* ledger */}
            <div style={{ border: `1px solid ${COLORS.border}`, borderRadius: "6px", overflow: "auto", maxHeight: "520px", marginBottom: "16px" }}>
              {rows.map((r, i) => (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "12px",
                    padding: "9px 14px",
                    borderTop: i === 0 ? "none" : `1px solid ${COLORS.border}`,
                    background: i % 2 === 0 ? COLORS.panel : COLORS.panelAlt,
                  }}
                >
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, width: "28px" }}>{i + 1}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "13px", width: "110px", flexShrink: 0 }}>{r.dni || "—"}</span>
                  <span style={{ fontSize: "13px", flex: 1 }}>{r.nombre || "—"}</span>
                  <StatusPill status={r.status} />
                </div>
              ))}
            </div>

            {/* copy block */}
            <div style={{ marginBottom: "10px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "6px" }}>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.dim, letterSpacing: "0.06em" }}>
                  BLOQUE PARA PEGAR EN EXCEL {camposActivos.length > 0 ? `(${camposActivos.map((c) => c.label).join(" / ")}${detail ? " / Estado / Observación" : ""})` : ""}
                </span>
                <button
                  onClick={copy}
                  style={{
                    background: "none",
                    border: `1px solid ${COLORS.border}`,
                    color: copied ? COLORS.green : COLORS.text,
                    fontFamily: "'IBM Plex Mono', monospace",
                    fontSize: "11px",
                    padding: "4px 10px",
                    borderRadius: "4px",
                    cursor: "pointer",
                  }}
                >
                  {copied ? "✓ copiado" : "copiar todo"}
                </button>
              </div>
              {camposActivos.length === 0 && (
                <div style={{ fontSize: "11px", color: COLORS.red, marginBottom: "6px" }}>
                  Marca al menos una columna arriba para generar el bloque.
                </div>
              )}
              <textarea
                ref={tsvRef}
                readOnly
                value={tsv}
                onFocus={(e) => e.target.select()}
                style={{
                  width: "100%",
                  minHeight: "160px",
                  background: "#FFFFFF",
                  border: `1px solid ${COLORS.border}`,
                  borderRadius: "6px",
                  padding: "14px",
                  fontFamily: "'IBM Plex Mono', monospace",
                  fontSize: "10pt",
                  color: "#000000",
                  whiteSpace: "pre",
                  overflowX: "auto",
                  resize: "vertical",
                  margin: 0,
                  boxSizing: "border-box",
                }}
              />
              <div style={{ fontSize: "11px", color: COLORS.dim, marginTop: "6px" }}>
                Usa el botón "copiar todo" — si seleccionas y copias el texto manualmente, el navegador puede perder el tabulador al pegar en Excel.
              </div>
            </div>

            {/* observations */}
            {flagged.length > 0 && (
              <div style={{ marginTop: "24px" }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: "11px", color: COLORS.amber, letterSpacing: "0.06em", marginBottom: "8px" }}>
                  PARA REVISAR MANUALMENTE ({flagged.length})
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                  {flagged.map((r, i) => (
                    <div key={i} style={{ fontSize: "12.5px", color: COLORS.dim }}>
                      <span style={{ color: COLORS.text }}>{r.nombre || r.raw.nombre || "(sin nombre)"} — {r.dni || r.raw.dni || "(sin doc)"}:</span> {r.obs}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
