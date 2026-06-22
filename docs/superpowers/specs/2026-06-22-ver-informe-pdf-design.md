# Ver Informe en PDF (Tigo + WOM)

## Alcance

`server.js`, `informe_clima_app.html` (Tigo), `informe_wom_app.html` (WOM).
Boton nuevo aparece solo justo despues de generar un informe (no en el
historial/registro).

## Conversion docx -> pdf

Se usa LibreOffice headless (`soffice --headless --convert-to pdf`) sobre
el `.docx` ya escrito en disco por `/generar` o `/generar-wom`. No se
duplica la plantilla del informe en HTML — se convierte el archivo real,
garantiza fidelidad exacta con el Word.

- Ruta del ejecutable: `process.env.LIBREOFFICE_PATH || 'soffice'`
  (en Windows, tipicamente `C:\Program Files\LibreOffice\program\soffice.exe`
  si no esta en PATH).
- Salida: carpeta `informes/pdf_tmp/` (creada si no existe, agregada a
  `.gitignore`).
- Sin cache: cada request a `/ver-pdf*` reconvierte y sobrescribe el PDF.
  Archivo pequeno, conversion ~1-2s, evita logica de invalidacion.
- Si `soffice` falla o no esta instalado: responder 500 con mensaje
  `"LibreOffice no instalado o no encontrado. Instala LibreOffice o
  configura LIBREOFFICE_PATH."` — sin fallback silencioso.

## Endpoints nuevos

- `GET /ver-pdf/:id` (Tigo): busca el registro en `informes_clima`
  (mismo lookup que `/descargar/:id`), obtiene el `.docx` (storage o
  `DOCS_DIR` local), convierte a pdf, responde
  `Content-Type: application/pdf` + `Content-Disposition: inline;
  filename="..."` (sin forzar descarga, para que el navegador lo
  muestre).
- `GET /ver-pdf-wom/:id` (WOM): igual, usando `informes_wom` /
  `DOCS_DIR_WOM`.

## Vinculo generar -> ver

`/generar` y `/generar-wom` ya devuelven el buffer del docx como body de
la respuesta (`res.send(buffer)`). Se agrega header
`X-Informe-Id: entry.id` a la respuesta, y se agrega `X-Informe-Id` a
`Access-Control-Expose-Headers` (junto a `Content-Disposition` que ya
esta expuesto) para que el cliente lo pueda leer.

## Frontend (ambos formularios)

- Tras un `generateDocx()` / equivalente WOM exitoso, leer
  `response.headers.get('X-Informe-Id')` y guardarlo en una variable
  (`lastGeneratedId`).
- Mostrar boton `👁 Ver Informe` (oculto por defecto) junto al boton
  `⬇ Descargar Informe Word`, habilitado solo cuando hay
  `lastGeneratedId` valido del ultimo generado.
- `onclick`: `window.open(`${SERVER}/ver-pdf/${lastGeneratedId}`,
  '_blank')` (Tigo) / `/ver-pdf-wom/` (WOM).
- Si la conversion falla (500), mostrar el mensaje de error del server en
  el `statusMsg` existente (mismo patron que errores de `/generar`).

## Fuera de alcance

- Boton "Ver Informe" en el historial/registro de informes ya guardados.
- Cache o limpieza periodica de `pdf_tmp/`.
- Instalacion de LibreOffice (la hace el usuario manualmente).
