# Flujo de Power Automate: "Reporte diario → borrador Outlook"

Crea un **borrador** en Outlook con el reporte diario (tabla con líneas y colores y el Excel adjunto) cada vez que aparece un `LOTE_REPORTE_*.json` en la carpeta ENTRADA. Archiva el Excel y mueve el JSON. Es independiente de los flujos de `ENTRADA`, `ENTRADA_DOCS` y `ENTRADA_VCC`.

## Carpetas (OneDrive de One Marcobre)

Sitio: `https://perujjc-my.sharepoint.com/personal/onemarcobre_jjc_com_pe`, biblioteca **Documentos**.

```
SSOMA/01. Capacitaciones/1. AUTOMATIZACIONES/REPORTE DIARIO/
├── ENTRADA      ← Claude deja aquí LOTE_REPORTE_AAAAMMDD.json
├── PROCESADOS   ← el flujo mueve aquí el JSON procesado
├── ERRORES      ← el flujo mueve aquí el JSON si algo falla
├── REPORTES     ← el flujo guarda aquí el Excel del día
└── CONFIG       ← destinatarios.json, script y esta guía
```

## Qué trae el lote

```json
{
  "tipo": "REPORTE_DIARIO",
  "fecha": "2026-11-22",
  "archivo_excel": "Reporte_Diario_2026-11-22.xlsx",
  "mensaje": { ...mensaje de Microsoft Graph listo para enviar... }
}
```

`mensaje` ya trae el asunto, Para, CC, el cuerpo HTML con formato y el Excel adjunto. El flujo solo lo pasa a Outlook sin tocarlo.

## Pasos para crear el flujo

> En las carpetas, usa el ícono de carpeta para elegirlas en lugar de escribir la ruta. Así Power Automate pone el formato correcto.

En Power Automate: **Crear → Flujo de nube automatizado**. Nombre: `Reporte diario - borrador Outlook`.

### 1. Desencadenador: SharePoint → "Cuando se crea un archivo (solo propiedades)"
- **Dirección del sitio:** `https://perujjc-my.sharepoint.com/personal/onemarcobre_jjc_com_pe` (si no aparece en la lista, elige "Escribir valor personalizado").
- **Nombre de la biblioteca:** `Documentos`.
- **Carpeta:** `/Documents/SSOMA/01. Capacitaciones/1. AUTOMATIZACIONES/REPORTE DIARIO/ENTRADA`.
- **Configuración (⋯ → Configuración) → Condiciones del desencadenador → + Agregar**, para que solo se active con los lotes:
  ```
  @startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], 'LOTE_REPORTE_')
  ```

### 2. Ámbito "Intentar"
Agrega un control **Ámbito** llamado `Intentar` y dentro pon los pasos 2.1 a 2.5.

**2.1 SharePoint → "Obtener contenido de archivo"**
- Dirección del sitio y biblioteca: los mismos del desencadenador.
- **Identificador de archivo:** contenido dinámico **Identificador** del desencadenador.

**2.2 Componer** (renómbralo `Lote`). En **Entradas**, pega esta expresión:
```
json(base64ToString(body('Obtener_contenido_de_archivo')?['$content']))
```
> Si falla con "The template language function 'base64ToString'…", cambia la expresión por `json(body('Obtener_contenido_de_archivo'))`.

**2.3 Office 365 Outlook → "Enviar una solicitud HTTP"** (renómbralo `Crear borrador`)
- **URI:** `https://graph.microsoft.com/v1.0/me/messages`
- **Método:** `POST`
- **Cuerpo:** la expresión `outputs('Lote')?['mensaje']`
- **Tipo de contenido:** `application/json`

Esto **solo crea el borrador** en tu carpeta Borradores, no lo envía.

**2.4 SharePoint → "Crear archivo"** (guarda el Excel)
- **Ruta de acceso de la carpeta:** `/Documents/SSOMA/01. Capacitaciones/1. AUTOMATIZACIONES/REPORTE DIARIO/REPORTES`
- **Nombre de archivo:** la expresión `outputs('Lote')?['archivo_excel']`
- **Contenido del archivo:** la expresión `base64ToBinary(outputs('Lote')?['mensaje']?['attachments']?[0]?['contentBytes'])`

> Si ese día ya hay un Excel con el mismo nombre, el paso falla. Para que lo reemplace, usa "Actualizar archivo" o borra el anterior.

**2.5 SharePoint → "Mover archivo"**
- **Archivo para mover:** **Identificador** del desencadenador.
- **Carpeta de destino:** `/Documents/SSOMA/01. Capacitaciones/1. AUTOMATIZACIONES/REPORTE DIARIO/PROCESADOS`
- **Si ya hay otro archivo:** `Reemplazar`.

### 3. Ámbito "Si falla"
Agrega otro **Ámbito** debajo de `Intentar`, llamado `Si falla`. Abre **⋯ → Configurar ejecución posterior** y marca solo **ha fallado** y **ha agotado el tiempo de espera**.

Dentro:
1. **SharePoint → Mover archivo:** el **Identificador** del desencadenador a `.../REPORTE DIARIO/ERRORES` (Si ya hay otro archivo: Reemplazar).
2. **Office 365 Outlook → Enviar un correo electrónico (V2):**
   - **Para:** tu correo.
   - **Asunto:** `Reporte diario: no se pudo crear el borrador`.
   - **Cuerpo:** `El lote @{triggerOutputs()?['body/{FilenameWithExtension}']} pasó a ERRORES. Revisa el historial del flujo.`

### 4. Guardar y probar
1. Guarda el flujo.
2. Avísale a Claude para que suba a `ENTRADA` el lote de prueba del 22/11/2026.
3. En 1 a 5 minutos deberías tener el borrador en Outlook, el Excel en `REPORTES` y el JSON en `PROCESADOS`.

## Errores comunes

| Síntoma | Causa y solución |
|---|---|
| El flujo no se activa | La carpeta del desencadenador está mal escrita o la condición no coincide con el nombre del archivo. |
| `Crear borrador` da 400 | El JSON está incompleto: abre el lote en `ERRORES` y avísale a Claude. |
| `Crear borrador` da 403 | La conexión de Office 365 Outlook no es la de tu cuenta (`joseph.flores@…`). Corrígela en Conexiones. |
| El borrador no tiene tu firma | Es normal: los borradores creados por Graph no llevan firma. Insértala desde Outlook (Insertar → Firma) antes de enviar. |
| El correo llega sin colores | No debería pasar con Graph. Si pasa, revisa que el paso 2.3 use `outputs('Lote')?['mensaje']` sin modificarlo. |
