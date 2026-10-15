# Lector de registros manuscritos → Fuerza laboral

App React (artefacto de claude.ai) que lee registros de asistencia, inducción o capacitación escritos a mano y los cruza contra el Excel de fuerza laboral. Entrega un bloque listo para pegar en Excel.

## Cómo funciona

1. **Lectura:** Claude transcribe la tabla de asistentes del registro (imágenes o PDF). Ignora las páginas de exámenes y anexos.
2. **Cruce:** el código, no el modelo, busca a cada persona en las hojas `DATA`, `DADO DE BAJA` y `VISITA` por DNI y por nombre. Tolera errores de lectura, nombres en otro orden y abreviaturas de cargo. Los datos de salida se copian del maestro; nunca se inventan.
3. **Segunda lectura:** las filas dudosas se vuelven a revisar contra el documento original. Claude elige entre los candidatos reales de la fuerza laboral o responde que ninguno corresponde.

## Uso

1. En un chat de claude.ai, pide crear un artefacto React con el contenido de `App.jsx`.
2. Sube el Excel de fuerza laboral (`.xlsx` o `.xlsb`). Queda guardado para la próxima vez.
3. Sube el registro (una o varias imágenes o PDF), marca las columnas que quieres y pulsa **Procesar registro**.
4. Usa **copiar todo** y pega en Excel.

## Archivos

| Archivo | Qué es |
|---|---|
| `App.jsx` | Código completo de la app |
| `cambios-para-el-artefacto.txt` | Instrucciones para actualizar un artefacto ya abierto sin pegar todo el código |

## Configuración

Los valores ajustables están al inicio de `App.jsx`: modelo, páginas por parte, partes en paralelo, tamaños máximos y umbrales de similitud.
