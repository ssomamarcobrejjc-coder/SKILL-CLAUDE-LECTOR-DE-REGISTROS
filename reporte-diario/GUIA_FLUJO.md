# Flujo de Power Automate: "Reporte diario → borrador Outlook"

Crea un **borrador** en Outlook con el reporte diario cada vez que aparece un `LOTE_REPORTE_*.json` en la carpeta ENTRADA. El borrador lleva la tabla con líneas y colores y el Excel adjunto. Además, el flujo archiva el Excel y mueve el JSON. Es independiente de los demás flujos de automatización.

## Carpetas (tu OneDrive - JJC)

En la PC: `C:\Users\joseph.flores\OneDrive - JJC\Escritorio\REPOS\REPORTE DIARIO`

```
Escritorio/REPOS/REPORTE DIARIO/
├── ENTRADA      ← Claude deja aquí LOTE_REPORTE_AAAAMMDD.json
├── PROCESADOS   ← el flujo mueve aquí el JSON procesado
├── ERRORES      ← el flujo mueve aquí el JSON si algo falla
├── REPORTES     ← el flujo guarda aquí el Excel del día
└── CONFIG       ← destinatarios.json y esta guía
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

`mensaje` ya trae el asunto, Para, CC, el cuerpo HTML con el formato de tu reporte y el Excel adjunto. El flujo lo pasa a Outlook sin tocarlo.

## Pasos para crear el flujo

En Power Automate: **Crear → Flujo de nube automatizado**. Nombre: `Reporte diario - borrador Outlook`.

> En cada campo de carpeta, usa el ícono de carpeta para elegirla en lugar de escribir la ruta.

### 1. Desencadenador: OneDrive para la Empresa → "Cuando se crea un archivo"
- **Carpeta:** `/Escritorio/REPOS/REPORTE DIARIO/ENTRADA`
- **Incluir subcarpetas:** No.
- **⋯ → Configuración → Condiciones del desencadenador → + Agregar**, para que solo se active con los lotes:
  ```
  @startsWith(triggerOutputs()?['headers/x-ms-file-name'], 'LOTE_REPORTE_')
  ```

### 2. Ámbito "Intentar"
Agrega un control **Ámbito** llamado `Intentar` y pon dentro los pasos 2.1 a 2.4.

**2.1 Componer** (renómbralo `Lote`). En **Entradas**, pega esta expresión:
```
json(base64ToString(triggerBody()?['$content']))
```
> Si da error, usa `json(triggerBody())`.

**2.2 Office 365 Outlook → "Enviar una solicitud HTTP"** (renómbralo `Crear borrador`)
- **URI:** `https://graph.microsoft.com/v1.0/me/messages`
- **Método:** `POST`
- **Cuerpo:** la expresión `outputs('Lote')?['mensaje']`
- **Tipo de contenido:** `application/json`

Este paso **solo crea el borrador** en tu carpeta Borradores; no lo envía.

**2.3 OneDrive para la Empresa → "Crear archivo"** (guarda el Excel)
- **Ruta de acceso de la carpeta:** `/Escritorio/REPOS/REPORTE DIARIO/REPORTES`
- **Nombre de archivo:** la expresión `outputs('Lote')?['archivo_excel']`
- **Contenido del archivo:** la expresión `base64ToBinary(outputs('Lote')?['mensaje']?['attachments']?[0]?['contentBytes'])`

**2.4 OneDrive para la Empresa → "Mover o cambiar el nombre de un archivo"**
- **Archivo:** la expresión `triggerOutputs()?['headers/x-ms-file-id']`
- **Ruta de acceso de destino:** `/Escritorio/REPOS/REPORTE DIARIO/PROCESADOS/@{triggerOutputs()?['headers/x-ms-file-name']}`
- **Sobrescribir:** Sí.

### 3. Ámbito "Si falla"
Agrega otro **Ámbito** debajo de `Intentar`, llamado `Si falla`. Entra a **⋯ → Configurar ejecución posterior** y marca solo **ha fallado** y **ha agotado el tiempo de espera**.

Dentro:
1. **OneDrive para la Empresa → Mover o cambiar el nombre de un archivo**: el mismo archivo, con destino `/Escritorio/REPOS/REPORTE DIARIO/ERRORES/@{triggerOutputs()?['headers/x-ms-file-name']}` y Sobrescribir: Sí.
2. **Office 365 Outlook → Enviar un correo electrónico (V2)**:
   - **Para:** tu correo.
   - **Asunto:** `Reporte diario: no se pudo crear el borrador`
   - **Cuerpo:** `El lote @{triggerOutputs()?['headers/x-ms-file-name']} pasó a ERRORES. Revisa el historial del flujo.`

### 4. Guardar y probar
1. Guarda el flujo.
2. Pídele a Claude que suba a `ENTRADA` el lote de prueba del 22/11/2026.
3. En unos minutos deberías tener el borrador en Outlook, el Excel en `REPORTES` y el JSON en `PROCESADOS`.

## Errores comunes

| Síntoma | Causa y solución |
|---|---|
| El flujo no se activa | La carpeta del desencadenador está mal elegida, o el nombre del archivo no empieza con `LOTE_REPORTE_`. |
| `Crear borrador` da error 400 | El JSON está incompleto. Abre el lote que quedó en `ERRORES` y avísale a Claude. |
| `Crear borrador` da error 403 | La conexión de Office 365 Outlook no es la de tu cuenta. Corrígela en **Conexiones**. |
| `Crear archivo` falla | Ya existe un Excel con ese nombre en `REPORTES`. Bórralo o cambia la acción por **Actualizar archivo**. |
| El borrador no tiene tu firma | Es normal: los borradores creados por Graph no llevan firma. Insértala desde Outlook (**Insertar → Firma**) antes de enviar. |
