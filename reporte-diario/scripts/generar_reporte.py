#!/usr/bin/env python3
"""Genera el Excel del reporte diario y el LOTE_REPORTE_*.json para Power Automate.

Uso:
    python3 generar_reporte.py actividades.json destinatarios.json [carpeta_salida]

actividades.json:
{
  "fecha": "2026-11-22",
  "actividades": [
    {"actividad": "Carga de VCC", "detalle": "Cargadas hasta el 20/11/2026",
     "avance": 0.9, "pendientes": "10 VCC pendientes"}
  ]
}
"avance" va de 0 a 1 (0.9 = 90 %). Usa null si no hay porcentaje.
Opcional: "tiempo_perdido": [{"actividad", "solicitado_por", "detalle"}] agrega el
apartado INDICADOR DE TIEMPO PERDIDO debajo de la tabla. "tiempo_perdido_titulo" y
"tiempo_perdido_col3" cambian el título del apartado y el encabezado de la 3.ª columna.

destinatarios.json: ver destinatarios.ejemplo.json.
"""
import base64, html, json, sys, uuid
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule, CellIsRule

F = "Arial"
AZUL, AZUL2, CELESTE = "1F4E78", "2E75B6", "DDEBF7"
COLOR_ESTADO = {"Pendiente": "F8CBAD", "En proceso": "FFE699", "Completado": "C6EFCE"}
HDR = ["N°", "Actividad", "Detalle / Situación", "Avance (%)", "Estado", "Pendientes / Próximos pasos"]


def estado(av):
    if av is None:
        return "En proceso"
    if av >= 1:
        return "Completado"
    return "Pendiente" if av == 0 else "En proceso"


def fecha_txt(f):
    return f.strftime("%d/%m/%Y")


def crear_excel(fecha, acts, ruta, tiempo_perdido=None, tp_titulo="INDICADOR DE TIEMPO PERDIDO", tp_col3="Solicitado por"):
    wb = Workbook(); ws = wb.active; ws.title = f"Reporte {fecha.strftime('%d-%m-%Y')}"
    thin = Side(style="thin", color="808080"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells("A1:F1"); ws["A1"] = "REPORTE DIARIO DE ACTIVIDADES"
    ws["A1"].font = Font(name=F, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=AZUL)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center"); ws.row_dimensions[1].height = 26
    ws["A2"] = "Fecha:"; ws["A2"].font = Font(name=F, bold=True)
    ws["B2"] = fecha; ws["B2"].number_format = "dd/mm/yyyy"; ws["B2"].font = Font(name=F)
    ws["B2"].alignment = Alignment(horizontal="left")
    for i, h in enumerate(HDR, 1):
        c = ws.cell(4, i, h); c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL2); c.border = B
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r0 = 5
    for n, a in enumerate(acts):
        r = r0 + n
        vals = [n + 1, a["actividad"], a.get("detalle", ""), a.get("avance"),
                f'=IF(D{r}="","En proceso",IF(D{r}>=1,"Completado",IF(D{r}=0,"Pendiente","En proceso")))',
                a.get("pendientes", "")]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v); c.font = Font(name=F, size=10); c.border = B
            c.alignment = Alignment(vertical="center", wrap_text=True,
                                    horizontal="center" if i in (1, 4, 5) else "left")
        ws.cell(r, 4).number_format = "0%"
    last = r0 + len(acts) - 1; t = last + 1
    ws.merge_cells(f"A{t}:C{t}"); ws[f"A{t}"] = "AVANCE PROMEDIO"
    ws[f"D{t}"] = f'=IFERROR(AVERAGE(D{r0}:D{last}),"")'; ws[f"D{t}"].number_format = "0%"
    ws.merge_cells(f"E{t}:F{t}")
    ws[f"E{t}"] = (f'=COUNTIF(E{r0}:E{last},"Pendiente")&" pendientes / "&'
                   f'COUNTIF(E{r0}:E{last},"En proceso")&" en proceso / "&'
                   f'COUNTIF(E{r0}:E{last},"Completado")&" completadas"')
    for i in range(1, 7):
        c = ws.cell(t, i); c.font = Font(name=F, bold=True, size=10)
        c.fill = PatternFill("solid", fgColor=CELESTE); c.border = B
        c.alignment = Alignment(horizontal="center" if i >= 4 else "right", vertical="center")
    rng = f"D{r0}:D{last}"
    ws.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num",
                                                   end_value=1, color="63BE7B"))
    for k, col in COLOR_ESTADO.items():
        ws.conditional_formatting.add(f"E{r0}:E{last}", CellIsRule(
            operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=col)))
    for col, w in zip("ABCDEF", [5, 42, 34, 12, 14, 48]):
        ws.column_dimensions[col].width = w
    if tiempo_perdido:
        r = t + 2
        ws.merge_cells(f"A{r}:F{r}"); c = ws[f"A{r}"]; c.value = tp_titulo
        c.font = Font(name=F, size=12, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center"); ws.row_dimensions[r].height = 22
        r += 1
        for col, h in zip("ABCD", ["N°", "Actividad", tp_col3, "Detalle / Motivo"]):
            c = ws[f"{col}{r}"]; c.value = h
        ws.merge_cells(f"D{r}:F{r}")
        for i in range(1, 7):
            c = ws.cell(r, i); c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=AZUL2); c.border = B
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for n, x in enumerate(tiempo_perdido, 1):
            r += 1
            vals = [n, x["actividad"], x.get("solicitado_por", ""), x.get("detalle", "")]
            for i, v in enumerate(vals, 1):
                ws.cell(r, i, v)
            ws.merge_cells(f"D{r}:F{r}")
            for i in range(1, 7):
                c = ws.cell(r, i); c.font = Font(name=F, size=10); c.border = B
                c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center" if i == 1 else "left")
            ws.row_dimensions[r].height = 45
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(ruta)


def crear_html(fecha, acts, cfg, tiempo_perdido=None, tp_titulo="INDICADOR DE TIEMPO PERDIDO", tp_col3="Solicitado por"):
    """Replica el modelo del correo enviado el 22/11/2026 (tabla pegada desde Excel)."""
    e = lambda x: html.escape(str(x if x is not None else ""))
    W = ["54.0pt", "288.9pt", "234.25pt", "82.7pt", "96.7pt", "319.9pt"]
    pad = "padding:0cm 3.5pt 0cm 3.5pt"
    bd = "border:solid gray 1.0pt"
    P = lambda txt, al="", b=False: (f'<p class="MsoNormal"{f" style={chr(34)}text-align:{al}{chr(34)}" if al else ""}>'
                                     f'{"<b>" if b else ""}{txt}{"</b>" if b else ""}</p>')

    def celda(i, txt, al="", b=False, bg="", h="20.6pt", first_row=False, span=1, w=None):
        borde = bd if first_row or i == 0 else f"{bd};border-left:none"
        if not first_row:
            borde += ";border-top:none"
        fondo = f";background:{bg}" if bg else ""
        cs = f' colspan="{span}"' if span > 1 else ""
        return (f'<td{cs} style="width:{w or W[i]};{borde}{fondo};{pad};height:{h}">{P(txt, al, b)}</td>')

    menciones = "/ ".join(
        f'<a id="OWAAM{uuid.uuid4().hex.upper()}" href="mailto:{e(x["correo"])}">@{e(x["alias"])}</a>'
        for x in cfg["para"] if x.get("alias"))
    saludo = f'<p>{e(cfg.get("saludo_inicio", "Estimados"))} {menciones}</p>'
    intro = f"<p>Adjunto mi reporte del día {fecha_txt(fecha)}<br /><br /></p>"

    filas = []
    filas.append(f'<tr style="height:37.15pt"><td colspan="6" style="width:1076.6pt;background:#1F4E78;{pad};height:37.15pt">'
                 f'<p class="MsoNormal" style="text-align:center"><b><span style="color:white;font-size:14pt">REPORTE DIARIO DE ACTIVIDADES</span></b></p></td></tr>')
    vacio = lambda i: f'<td style="width:{W[i]};{pad};height:20.6pt"></td>'
    filas.append(f'<tr style="height:20.6pt"><td style="width:{W[0]};{pad};height:20.6pt">{P("Fecha:", b=True)}</td>'
                 f'<td style="width:{W[1]};{pad};height:20.6pt">{P(fecha_txt(fecha))}</td>' + "".join(vacio(i) for i in range(2, 6)) + "</tr>")
    filas.append('<tr style="height:20.6pt">' + "".join(vacio(i) for i in range(6)) + "</tr>")
    hdr = "".join(celda(i, h, "center", True, "#2E75B6", "39.5pt", first_row=True) for i, h in enumerate(HDR))
    filas.append(f'<tr style="height:39.5pt">{hdr}</tr>')
    con_av = []
    for n, a in enumerate(acts, 1):
        av = a.get("avance")
        if av is not None:
            con_av.append(av)
        h = "37.75pt" if len(a["actividad"]) > 45 or len(a.get("pendientes") or "") > 60 else "20.6pt"
        vals = [(str(n), "center"), (e(a["actividad"]), ""), (e(a.get("detalle")), ""),
                ("—" if av is None else f"{round(av * 100)}%", "center"), (estado(av), "center"),
                (e(a.get("pendientes")), "")]
        filas.append(f'<tr style="height:{h}">' + "".join(celda(i, t, al, h=h) for i, (t, al) in enumerate(vals)) + "</tr>")
    prom = f"{round(sum(con_av) / len(con_av) * 100)}%" if con_av else "—"
    cnt = {k: sum(estado(a.get("avance")) == k for a in acts) for k in ("Pendiente", "En proceso", "Completado")}
    etq = {"Pendiente": ("pendiente", "pendientes"), "En proceso": ("en proceso", "en proceso"),
           "Completado": ("completada", "completadas")}
    partes = [f"{cnt[k]} {v[cnt[k] != 1]}" for k, v in etq.items() if cnt[k]]
    filas.append('<tr style="height:20.6pt">'
                 + celda(0, "AVANCE PROMEDIO", "right", True, "#DDEBF7", span=3, w="577.2pt")
                 + celda(3, prom, "center", True, "#DDEBF7")
                 + celda(4, " / ".join(partes), "center", True, "#DDEBF7", span=2, w="416.65pt")
                 + "</tr>")
    tabla = ('<table class="MsoNormalTable" style="width:1076.6pt;border-collapse:collapse"><tbody>'
             + "".join(filas) + "</tbody></table>")
    if tiempo_perdido:
        tw = ["54.0pt", "288.9pt", "234.25pt", "499.3pt"]
        tf = [f'<tr style="height:24pt"><td colspan="4" style="width:1076.6pt;background:#1F4E78;{pad};height:24pt">'
              f'<p class="MsoNormal" style="text-align:center"><b><span style="color:white;font-size:12pt">{e(tp_titulo)}</span></b></p></td></tr>']
        tf.append('<tr style="height:30pt">' + "".join(
            celda(min(i, 5), h, "center", True, "#2E75B6", "30pt", first_row=True, w=tw[i], span=1)
            for i, h in enumerate(["N°", "Actividad", e(tp_col3), "Detalle / Motivo"])) + "</tr>")
        for n, x in enumerate(tiempo_perdido, 1):
            v = [(str(n), "center"), (e(x["actividad"]), ""), (e(x.get("solicitado_por")), ""), (e(x.get("detalle")), "")]
            tf.append('<tr style="height:37.75pt">' + "".join(
                celda(i, t, al, h="37.75pt", w=tw[i]) for i, (t, al) in enumerate(v)) + "</tr>")
        tabla += ('<p class="MsoNormal"><br /></p><table class="MsoNormalTable" style="width:1076.6pt;border-collapse:collapse"><tbody>'
                  + "".join(tf) + "</tbody></table>")
    cierre = f'<p class="MsoNormal"><br /><br /></p><p class="MsoNormal">{e(cfg.get("despedida", "Saludos;"))}</p>'
    return f'<div class="WordSection1">{saludo}{intro}{tabla}{cierre}</div>'


def main():
    acts_p, dest_p = Path(sys.argv[1]), Path(sys.argv[2])
    out = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(".")
    out.mkdir(parents=True, exist_ok=True)
    data = json.loads(acts_p.read_text(encoding="utf-8"))
    cfg = json.loads(dest_p.read_text(encoding="utf-8"))
    fecha = date.fromisoformat(data["fecha"]); acts = data["actividades"]
    tp = (data.get("tiempo_perdido_titulo", "INDICADOR DE TIEMPO PERDIDO"), data.get("tiempo_perdido_col3", "Solicitado por"))
    xlsx = out / f"Reporte_Diario_{fecha.isoformat()}.xlsx"
    crear_excel(fecha, acts, xlsx, data.get("tiempo_perdido"), *tp)
    asunto = cfg.get("asunto", "Reporte diario de actividades – {fecha}").format(fecha=fecha_txt(fecha))
    rec = lambda lst: [{"emailAddress": {"address": x["correo"], "name": x.get("nombre", "")}} for x in lst]
    lote = {
        "tipo": "REPORTE_DIARIO",
        "fecha": fecha.isoformat(),
        "archivo_excel": xlsx.name,
        "mensaje": {
            "subject": asunto,
            "importance": "normal",
            "body": {"contentType": "HTML", "content": crear_html(fecha, acts, cfg, data.get("tiempo_perdido"), *tp)},
            "toRecipients": rec(cfg["para"]),
            "ccRecipients": rec(cfg.get("cc", [])),
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": xlsx.name,
                "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "contentBytes": base64.b64encode(xlsx.read_bytes()).decode(),
            }],
        },
    }
    lp = out / f"LOTE_REPORTE_{fecha.strftime('%Y%m%d')}.json"
    lp.write_text(json.dumps(lote, ensure_ascii=False), encoding="utf-8")
    print(f"EXCEL {xlsx}\nLOTE {lp}")


if __name__ == "__main__":
    main()
