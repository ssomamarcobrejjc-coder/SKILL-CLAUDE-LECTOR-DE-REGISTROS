---
name: "reporte-diario"
description: "Arma el reporte diario de actividades de Joseph (tabla en Excel) y deja el LOTE_REPORTE_*.json en REPORTE DIARIO\\ENTRADA para que Power Automate cree el borrador en Outlook con la tabla y el Excel adjunto."
---

# Reporte diario → borrador en Outlook

Se usa cuando Joseph dicta sus actividades del día y pide su reporte diario o el correo del reporte.

## Ubicación (OneDrive de One Marcobre, conector Microsoft 365)
- driveId: `b!zBaSpFjYjEm1glOVzE7cDegcK7fx8QREnWLmfaeq4zrrjNqSKNCcR6pZo3E1xYP_`
- `REPORTE DIARIO`: `01A2SIUE67PLNFPKTMUJHZS5ZAJBHYNARS`
- `ENTRADA`: `01A2SIUE6POLGXH5M5XVGLN2733NJPOSBX`
- `CONFIG`: `01A2SIUE7DQVZ4Q4DKOZFK5JC6G5O6U4PQ` (contiene `destinatarios.json` y `generar_reporte.py`)
- Ruta: `SSOMA\01. Capacitaciones\1. AUTOMATIZACIONES\REPORTE DIARIO`

## Pasos
1. Convierte lo que dicta Joseph en `actividades.json`:
   - `fecha`: hoy, en formato AAAA-MM-DD.
   - Por cada actividad: `actividad`, `detalle`, `avance` (de 0 a 1, o `null` si no da %), `pendientes`.
   - Si dice "pendiente" sin dar %, usa `avance: 0`. Si da una fecha en lugar de %, usa `null` y pon la fecha en `detalle`.
2. Lee `CONFIG/destinatarios.json` con `read_resource`. Si Joseph cambia los destinatarios, actualiza ese archivo (`sharepoint_upload_file` con `conflictBehavior: replace`).
3. Ejecuta `python3 generar_reporte.py actividades.json destinatarios.json salida/`. El script está en `scripts/` de este repositorio y en `CONFIG`.
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
- No subas nada fuera de `REPORTE DIARIO\ENTRADA` y `CONFIG`.
- Si el lote aparece en `ERRORES`, revisa el JSON y vuelve a subirlo corregido.
