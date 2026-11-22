---
name: "reporte-diario"
description: "Arma el reporte diario de actividades de Joseph (tabla en Excel) y deja el LOTE_REPORTE_*.json en REPORTE DIARIO\\ENTRADA para que Power Automate cree el borrador en Outlook con la tabla y el Excel adjunto."
---

# Reporte diario → borrador en Outlook

Se usa cuando Joseph dicta sus actividades del día y pide su reporte diario o el correo del reporte.

## Ubicación (OneDrive personal de Joseph, conector Microsoft 365)
- PC: `C:\Users\joseph.flores\OneDrive - JJC\Escritorio\REPOS\REPORTE DIARIO`
- driveId: `b!kFzZASJFPkuhhJmcVcW095ffmQGkWZFGnMo8_C1bHNcpm_KSRCerTK8qokziB5YV`
- `REPORTE DIARIO`: `01XZZ665YKF6G6BM4O3BH2IMOZFHYKQFFR`
- `ENTRADA`: `01XZZ6657J6UJG4CSJWBBY4XHB7CI4FA46`
- `CONFIG`: `01XZZ6654ISDLFDTDF4VA2HEU7T5KYXV2N` (contiene `destinatarios.json` y `GUIA_FLUJO.md`)
- `REPORTES` (Excel archivados por el flujo): `01XZZ6652EJCW25VFMJBEYVEEGBE34HOLK`

## Pasos
1. Convierte lo que dicta Joseph en `actividades.json`:
   - `fecha`: hoy, en formato AAAA-MM-DD.
   - Por cada actividad: `actividad`, `detalle`, `avance` (de 0 a 1, o `null` si no da %), `pendientes`.
   - Si dice "pendiente" sin dar %, usa `avance: 0`. Si da una fecha en lugar de %, usa `null` y pon la fecha en `detalle`.
2. Lee `CONFIG/destinatarios.json` con `read_resource`. Si Joseph cambia los destinatarios, actualiza ese archivo (`sharepoint_upload_file` con `conflictBehavior: replace`).
3. Ejecuta `python3 generar_reporte.py actividades.json destinatarios.json salida/`. El script está en `reporte-diario/scripts/` de este repositorio.
4. **Muestra la previsualización:** la tabla en el chat, con Para, CC y asunto. Espera el visto bueno de Joseph.
5. Con su visto bueno, sube `LOTE_REPORTE_AAAAMMDD.json` a `ENTRADA`:
   - Si la PC está conectada, usa `device_commit_files` (sube el archivo tal cual, sin copiar su contenido).
   - Si no, usa `sharepoint_upload_file` con el texto del JSON. Power Automate crea el borrador, guarda el Excel en `REPORTES` y mueve el JSON a `PROCESADOS`.
6. Entrega también el Excel en el chat.

## Formato del correo (modelo: el reporte enviado el 22/11/2026)
- Saludo `Estimados @Edinson/ @Roberto` con menciones de Outlook (campo `alias` de `para`).
- `Adjunto mi reporte del día DD/MM/AAAA`.
- Tabla igual a la del Excel: título azul, fila de fecha, encabezado azul con bordes grises y fila AVANCE PROMEDIO en celeste.
- `Saludos;` y el Excel adjunto. La firma con imagen no va en el lote: Joseph la agrega al revisar el borrador.

## Reglas
- Nunca envíes el correo: solo se crea el borrador.
- No subas nada fuera de `REPORTE DIARIO\ENTRADA` y `REPORTE DIARIO\CONFIG`.
- Si el lote aparece en `ERRORES`, revisa el JSON y vuelve a subirlo corregido.
