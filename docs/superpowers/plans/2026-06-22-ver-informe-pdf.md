# Ver Informe en PDF (Tigo + WOM) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar boton "Ver Informe" que abre el informe recien generado como PDF en una pestana nueva del navegador, en Tigo y WOM.

**Architecture:** El servidor convierte el `.docx` ya generado a PDF en el momento, usando LibreOffice headless (`soffice --convert-to pdf`) via `child_process.execFile`. Dos endpoints nuevos (`/ver-pdf/:id`, `/ver-pdf-wom/:id`) reusan el lookup de registro que ya usan `/descargar/:id` y `/descargar-wom/:id`. El frontend recibe el id del informe recien creado por un header `X-Informe-Id` en la respuesta de `/generar` y `/generar-wom`, y lo usa para abrir `window.open` al endpoint de PDF.

**Tech Stack:** Node.js/Express (`server.js`), `child_process.execFile` (nativo), LibreOffice headless (binario externo, no es dependencia npm), HTML/JS vanilla en `informe_clima_app.html` y `informe_wom_app.html`.

Este proyecto no tiene suite de tests automatizados (sin Jest/Mocha, sin carpeta `tests/`). La verificacion de cada tarea es manual: reiniciar el servidor y probar con `curl` o en el navegador, igual que se hizo para el cambio de nomenclatura de archivos en `server.js`.

## Global Constraints

- No instalar LibreOffice automaticamente — lo instala el usuario manualmente (puede no estar disponible aun en esta maquina).
- Si `soffice` falla: responder 500 con mensaje exacto `"LibreOffice no instalado o no encontrado. Instala LibreOffice o configura LIBREOFFICE_PATH."` — sin fallback silencioso.
- Sin cache de PDFs: cada request a `/ver-pdf*` reconvierte.
- Boton "Ver Informe" solo aparece tras generar un informe nuevo (no en historial/registro).
- Salida de PDFs en `informes/pdf_tmp/` (ya cubierto por `informes/` en `.gitignore`, no requiere cambio de `.gitignore`).
- Ruta del binario configurable via `process.env.LIBREOFFICE_PATH`, default `'soffice'`.

---

### Task 1: Helper de conversion docx -> pdf en server.js

**Files:**
- Modify: `server.js:1-25` (requires y helpers de nombre de archivo)
- Modify: `server.js:86-94` (consts de directorios)

**Interfaces:**
- Produces: `PDFS_DIR` (const, string, path absoluto a `informes/pdf_tmp`), `convertDocxToPdf(docxPath: string): Promise<string>` — resuelve con el path absoluto al `.pdf` generado, o rechaza con `Error` cuyo `.message` es el texto exacto de la constraint global.

- [ ] **Step 1: Agregar `require('child_process')` y declarar `PDFS_DIR`**

En `server.js`, la linea 5 dice:
```js
const fs = require('fs');
```
Cambiar a:
```js
const fs = require('fs');
const { execFile } = require('child_process');
```

En `server.js`, las lineas 86-94 dicen:
```js
const DOCS_DIR      = path.join(__dirname, 'informes');
const PAPELERA_DIR  = path.join(__dirname, 'papelera');
const DB_FILE       = path.join(__dirname, 'registro.json');
const PAPELERA_FILE = path.join(__dirname, 'papelera.json');

if (!fs.existsSync(DOCS_DIR))     fs.mkdirSync(DOCS_DIR);
if (!fs.existsSync(PAPELERA_DIR)) fs.mkdirSync(PAPELERA_DIR);
if (!fs.existsSync(DB_FILE))      fs.writeFileSync(DB_FILE, '[]');
if (!fs.existsSync(PAPELERA_FILE))fs.writeFileSync(PAPELERA_FILE, '[]');
```
Cambiar a (agrega `PDFS_DIR` y su `mkdir`):
```js
const DOCS_DIR      = path.join(__dirname, 'informes');
const PAPELERA_DIR  = path.join(__dirname, 'papelera');
const PDFS_DIR      = path.join(DOCS_DIR, 'pdf_tmp');
const DB_FILE       = path.join(__dirname, 'registro.json');
const PAPELERA_FILE = path.join(__dirname, 'papelera.json');

if (!fs.existsSync(DOCS_DIR))     fs.mkdirSync(DOCS_DIR);
if (!fs.existsSync(PAPELERA_DIR)) fs.mkdirSync(PAPELERA_DIR);
if (!fs.existsSync(PDFS_DIR))     fs.mkdirSync(PDFS_DIR, { recursive: true });
if (!fs.existsSync(DB_FILE))      fs.writeFileSync(DB_FILE, '[]');
if (!fs.existsSync(PAPELERA_FILE))fs.writeFileSync(PAPELERA_FILE, '[]');
```

- [ ] **Step 2: Agregar la funcion `convertDocxToPdf`**

Inmediatamente despues del bloque anterior (despues de la linea con `fs.writeFileSync(PAPELERA_FILE, '[]');`), agregar:
```js

function convertDocxToPdf(docxPath) {
  return new Promise((resolve, reject) => {
    const soffice = process.env.LIBREOFFICE_PATH || 'soffice';
    execFile(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', PDFS_DIR, docxPath], (err) => {
      if (err) {
        return reject(new Error('LibreOffice no instalado o no encontrado. Instala LibreOffice o configura LIBREOFFICE_PATH.'));
      }
      const pdfPath = path.join(PDFS_DIR, path.basename(docxPath, '.docx') + '.pdf');
      if (!fs.existsSync(pdfPath)) {
        return reject(new Error('LibreOffice no instalado o no encontrado. Instala LibreOffice o configura LIBREOFFICE_PATH.'));
      }
      resolve(pdfPath);
    });
  });
}
```

- [ ] **Step 3: Verificar que el servidor arranca sin errores de sintaxis**

Run: `node -c server.js`
Expected: sin output (exit code 0, significa sintaxis valida).

- [ ] **Step 4: Commit**

```bash
git add server.js
git commit -m "Agregar helper convertDocxToPdf con LibreOffice headless"
```

---

### Task 2: Endpoint /ver-pdf/:id + header X-Informe-Id (Tigo)

**Files:**
- Modify: `server.js:883-907` (respuesta de `/generar` y endpoint `/descargar/:id`)

**Interfaces:**
- Consumes: `convertDocxToPdf(docxPath)` de Task 1, `dbClimaFind(id)`, `storageDownload(storagePath)`, `DOCS_DIR`, `PDFS_DIR` (ya existentes/Task 1).
- Produces: endpoint `GET /ver-pdf/:id` que responde `application/pdf` inline; header `X-Informe-Id` en la respuesta de `POST /generar`.

- [ ] **Step 1: Agregar header `X-Informe-Id` a la respuesta de `/generar`**

En `server.js`, dentro de `app.post('/generar', ...)`, las lineas actuales dicen:
```js
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro', async (req,res) => {
```
Cambiar a:
```js
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('X-Informe-Id', entry.id);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition, X-Informe-Id');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro', async (req,res) => {
```

- [ ] **Step 2: Agregar el endpoint `/ver-pdf/:id`**

En `server.js`, justo despues del endpoint `/descargar/:id` (que termina con):
```js
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition',`attachment; filename="${entry.filename}"`);
  res.send(buffer);
});
```
agregar inmediatamente despues:
```js

app.get('/ver-pdf/:id', async (req, res) => {
  const entry = await dbClimaFind(req.params.id);
  if (!entry) return res.status(404).json({ error: 'No encontrado' });
  const docxPath = path.join(DOCS_DIR, entry.filename);
  if (!fs.existsSync(docxPath)) {
    const buffer = await storageDownload(`clima/${entry.filename}`);
    if (!buffer) return res.status(404).json({ error: 'Archivo no existe' });
    fs.writeFileSync(docxPath, buffer);
  }
  try {
    const pdfPath = await convertDocxToPdf(docxPath);
    const pdfBuffer = fs.readFileSync(pdfPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${entry.filename.replace(/\.docx$/, '.pdf')}"`);
    res.send(pdfBuffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
```

- [ ] **Step 3: Verificar sintaxis**

Run: `node -c server.js`
Expected: sin output.

- [ ] **Step 4: Reiniciar servidor y probar el flujo completo**

Matar el proceso que escucha en :3000 y arrancar de nuevo:
```bash
PID=$(netstat -ano | grep ":3000" | grep LISTENING | awk '{print $5}'); taskkill //PID $PID //F
npm start &
sleep 1
```
Generar un informe de prueba y capturar el id del header:
```bash
curl -s -D - -o /tmp/test.docx -X POST http://localhost:3000/generar \
  -H "Content-Type: application/json" \
  -d '{"nombreSitio":"Sitio Test","tecnico":"Juan Perez","equipo":"1","circuito":"1","sala":"Sala A","fecha":"22-06-2026","codInforme":"TEST-VERPDF"}' \
  | grep -i x-informe-id
```
Expected: una linea `X-Informe-Id: <numero>` (el id es el timestamp `Date.now()` del registro).

Con ese id, probar `/ver-pdf/:id`:
```bash
curl -s -D - -o /tmp/test.pdf http://localhost:3000/ver-pdf/<ID_OBTENIDO_ARRIBA>
```
Expected (si LibreOffice ya esta instalado): `HTTP/1.1 200`, `Content-Type: application/pdf`, y `/tmp/test.pdf` es un PDF valido (`file /tmp/test.pdf` dice `PDF document`).
Expected (si LibreOffice NO esta instalado todavia): `HTTP/1.1 500` con body `{"error":"LibreOffice no instalado o no encontrado. Instala LibreOffice o configura LIBREOFFICE_PATH."}` — esto confirma que el manejo de error funciona aunque el binario no este presente.

- [ ] **Step 5: Commit**

```bash
git add server.js
git commit -m "Agregar endpoint /ver-pdf/:id y header X-Informe-Id en /generar (Tigo)"
```

---

### Task 3: Endpoint /ver-pdf-wom/:id + header X-Informe-Id (WOM)

**Files:**
- Modify: `server.js` (respuesta de `/generar-wom` y endpoint `/descargar-wom/:id`)

**Interfaces:**
- Consumes: `convertDocxToPdf(docxPath)` de Task 1, `dbWomFind(id)`, `storageDownload(storagePath)`, `DOCS_DIR_WOM`, `PDFS_DIR`.
- Produces: endpoint `GET /ver-pdf-wom/:id`; header `X-Informe-Id` en la respuesta de `POST /generar-wom`.

- [ ] **Step 1: Agregar header `X-Informe-Id` a la respuesta de `/generar-wom`**

En `server.js`, dentro de `app.post('/generar-wom', ...)`, las lineas actuales dicen:
```js
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro-wom', async (_req, res) => res.json(await dbWomList()));
```
Cambiar a:
```js
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('X-Informe-Id', entry.id);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition, X-Informe-Id');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro-wom', async (_req, res) => res.json(await dbWomList()));
```

- [ ] **Step 2: Agregar el endpoint `/ver-pdf-wom/:id`**

En `server.js`, ubicar el endpoint `/descargar-wom/:id` (usa `dbWomFind`, `DOCS_DIR_WOM`, prefijo de storage `wom/`). Justo despues de que ese endpoint termine con su `res.send(buffer);\n});`, agregar:
```js

app.get('/ver-pdf-wom/:id', async (req, res) => {
  const entry = await dbWomFind(req.params.id);
  if (!entry) return res.status(404).json({ error: 'No encontrado' });
  const docxPath = path.join(DOCS_DIR_WOM, entry.filename);
  if (!fs.existsSync(docxPath)) {
    const buffer = await storageDownload(`wom/${entry.filename}`);
    if (!buffer) return res.status(404).json({ error: 'Archivo no existe' });
    fs.writeFileSync(docxPath, buffer);
  }
  try {
    const pdfPath = await convertDocxToPdf(docxPath);
    const pdfBuffer = fs.readFileSync(pdfPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${entry.filename.replace(/\.docx$/, '.pdf')}"`);
    res.send(pdfBuffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
```

- [ ] **Step 3: Verificar sintaxis**

Run: `node -c server.js`
Expected: sin output.

- [ ] **Step 4: Reiniciar servidor y probar el flujo completo**

```bash
PID=$(netstat -ano | grep ":3000" | grep LISTENING | awk '{print $5}'); taskkill //PID $PID //F
npm start &
sleep 1
curl -s -D - -o /tmp/test-wom.docx -X POST http://localhost:3000/generar-wom \
  -H "Content-Type: application/json" \
  -d '{"ticket":"TICKET-VERPDF","instalacion":"Sitio Test WOM","tecnicos":["Juan Perez"]}' \
  | grep -i x-informe-id
```
Expected: linea `X-Informe-Id: <numero>`.
```bash
curl -s -D - -o /tmp/test-wom.pdf http://localhost:3000/ver-pdf-wom/<ID_OBTENIDO_ARRIBA>
```
Expected: mismo comportamiento que Task 2 Step 4 (200+PDF si LibreOffice esta, o 500 con el mensaje exacto si no esta).

- [ ] **Step 5: Commit**

```bash
git add server.js
git commit -m "Agregar endpoint /ver-pdf-wom/:id y header X-Informe-Id en /generar-wom"
```

---

### Task 4: Boton "Ver Informe" en Tigo (informe_clima_app.html)

**Files:**
- Modify: `informe_clima_app.html:628-631` (boton de generar, tab 6)
- Modify: `informe_clima_app.html:766-767` (variables globales)
- Modify: `informe_clima_app.html:1638-1684` (funcion `generateDocx`)
- Modify: `informe_clima_app.html:1306-1335` (funcion `nuevoInforme`)

**Interfaces:**
- Consumes: endpoint `GET /ver-pdf/:id` y header `X-Informe-Id` de Task 2.
- Produces: funcion global `verInformePdf()`, variable global `lastInformeId`.

- [ ] **Step 1: Declarar la variable global `lastInformeId`**

En `informe_clima_app.html`, la linea 767 dice:
```js
let serverOnline=false, allRecords=[], pendingDeleteId=null, pendingEmailId=null;
```
Cambiar a:
```js
let serverOnline=false, allRecords=[], pendingDeleteId=null, pendingEmailId=null, lastInformeId=null;
```

- [ ] **Step 2: Agregar el boton oculto en el HTML**

En `informe_clima_app.html`, las lineas 628-631 dicen:
```html
  <div style="padding:4px 0 20px">
    <button class="btn-gen" id="btnDocx" onclick="generateDocx()">⬇ Descargar Informe Word (.docx)</button>
    <div class="status" id="statusMsg"></div>
  </div>
```
Cambiar a:
```html
  <div style="padding:4px 0 20px">
    <button class="btn-gen" id="btnDocx" onclick="generateDocx()">⬇ Descargar Informe Word (.docx)</button>
    <button class="btn btn-secondary" id="btnVerPdf" onclick="verInformePdf()" style="display:none;width:100%;margin-top:8px">👁 Ver Informe (PDF)</button>
    <div class="status" id="statusMsg"></div>
  </div>
```

- [ ] **Step 3: Capturar el id y mostrar el boton al generar OK**

En `informe_clima_app.html`, dentro de `generateDocx()`, las lineas actuales dicen:
```js
    const blob=await resp.blob();
    const disp=resp.headers.get('Content-Disposition')||'';
    const fnMatch=disp.match(/filename="(.+)"/);
    const fname=fnMatch?fnMatch[1]:`${cod}.docx`;
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=fname;a.click();URL.revokeObjectURL(url);
    status.className='status ok';status.textContent=`✅ "${fname}" generado y guardado en el historial`;
    setTimeout(()=>loadHist(true),800);
```
Cambiar a:
```js
    const blob=await resp.blob();
    const disp=resp.headers.get('Content-Disposition')||'';
    const fnMatch=disp.match(/filename="(.+)"/);
    const fname=fnMatch?fnMatch[1]:`${cod}.docx`;
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=fname;a.click();URL.revokeObjectURL(url);
    lastInformeId=resp.headers.get('X-Informe-Id');
    document.getElementById('btnVerPdf').style.display=lastInformeId?'block':'none';
    status.className='status ok';status.textContent=`✅ "${fname}" generado y guardado en el historial`;
    setTimeout(()=>loadHist(true),800);
```

- [ ] **Step 4: Agregar la funcion `verInformePdf`**

En `informe_clima_app.html`, justo despues del cierre de `generateDocx` (la linea que dice `}finally{btn.disabled=false;}\n}`), agregar:
```js

function verInformePdf(){
  if(!lastInformeId)return;
  window.open(`${SERVER}/ver-pdf/${lastInformeId}`,'_blank');
}
```

- [ ] **Step 5: Ocultar el boton y limpiar el id al crear un informe nuevo**

En `informe_clima_app.html`, dentro de `nuevoInforme()`, la linea actual dice:
```js
  document.getElementById('statusMsg').style.display='none';
```
Cambiar a:
```js
  document.getElementById('statusMsg').style.display='none';
  lastInformeId=null;
  document.getElementById('btnVerPdf').style.display='none';
```

- [ ] **Step 6: Probar en el navegador**

Abrir `http://localhost:3000/tigo`, completar un informe minimo, ir al tab "Generar" y presionar "Descargar Informe Word". Confirmar que:
- El `.docx` se descarga como antes.
- El boton "👁 Ver Informe (PDF)" aparece debajo, antes oculto.
- Al hacer click, se abre una pestana nueva a `/ver-pdf/<id>` (PDF si LibreOffice esta instalado, o un JSON de error 500 si no).
- Presionar "Crear Nuevo Informe" oculta el boton de nuevo.

- [ ] **Step 7: Commit**

```bash
git add informe_clima_app.html
git commit -m "Agregar boton Ver Informe (PDF) tras generar en Tigo"
```

---

### Task 5: Boton "Ver Informe" en WOM (informe_wom_app.html)

**Files:**
- Modify: `informe_wom_app.html:496-501` (acciones del panel 3)
- Modify: `informe_wom_app.html:754-755` (variables globales)
- Modify: `informe_wom_app.html:920-975` (funcion `generarInforme`)

**Interfaces:**
- Consumes: endpoint `GET /ver-pdf-wom/:id` y header `X-Informe-Id` de Task 3.
- Produces: funcion global `verInformePdfWom()`, variable global `lastInformeIdWom`.

- [ ] **Step 1: Declarar la variable global `lastInformeIdWom`**

En `informe_wom_app.html`, la linea 754 dice:
```js
let photos = [], captions = [], photoTransforms = [];
```
Cambiar a:
```js
let photos = [], captions = [], photoTransforms = [];
let lastInformeIdWom = null;
```

- [ ] **Step 2: Agregar el boton oculto en el HTML**

En `informe_wom_app.html`, las lineas 496-501 dicen:
```html
  <div class="actions">
    <button class="btn btn-secondary" onclick="goTab(2)">← Anterior</button>
    <button class="btn btn-primary" id="btn-generar" onclick="generarInforme()">
      Generar Informe WOM
    </button>
  </div>
</div>
```
Cambiar a:
```html
  <div class="actions">
    <button class="btn btn-secondary" onclick="goTab(2)">← Anterior</button>
    <button class="btn btn-primary" id="btn-generar" onclick="generarInforme()">
      Generar Informe WOM
    </button>
  </div>
  <button class="btn btn-secondary" id="btn-ver-pdf-wom" onclick="verInformePdfWom()" style="display:none;width:100%;margin-top:8px">👁 Ver Informe (PDF)</button>
</div>
```

- [ ] **Step 3: Capturar el id y mostrar el boton al generar OK**

En `informe_wom_app.html`, dentro de `generarInforme()`, las lineas actuales dicen:
```js
    const blob = await resp.blob();
    const cd   = resp.headers.get('Content-Disposition') || '';
    const name = cd.match(/filename="?([^"]+)"?/)?.[1] || `${payload.ticket}_WOM.docx`;
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = name; a.click();
    URL.revokeObjectURL(url);

    toast(`✓ Informe "${name}" generado`, 'ok');
    clearState();
```
Cambiar a:
```js
    const blob = await resp.blob();
    const cd   = resp.headers.get('Content-Disposition') || '';
    const name = cd.match(/filename="?([^"]+)"?/)?.[1] || `${payload.ticket}_WOM.docx`;
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = name; a.click();
    URL.revokeObjectURL(url);

    lastInformeIdWom = resp.headers.get('X-Informe-Id');
    const verBtn = g('btn-ver-pdf-wom');
    if (verBtn) verBtn.style.display = lastInformeIdWom ? 'block' : 'none';

    toast(`✓ Informe "${name}" generado`, 'ok');
    clearState();
```

- [ ] **Step 4: Agregar la funcion `verInformePdfWom`**

En `informe_wom_app.html`, justo despues del cierre de `generarInforme` (la linea que dice `}\n}` al final de esa funcion, antes del comentario `// ── Historial`), agregar:
```js

function verInformePdfWom() {
  if (!lastInformeIdWom) return;
  window.open(`/ver-pdf-wom/${lastInformeIdWom}`, '_blank');
}
```

- [ ] **Step 5: Probar en el navegador**

Abrir `http://localhost:3000/wom`, completar un informe minimo, presionar "Generar Informe WOM". Confirmar que:
- El `.docx` se descarga como antes.
- El boton "👁 Ver Informe (PDF)" aparece debajo, antes oculto.
- Al hacer click, se abre una pestana nueva a `/ver-pdf-wom/<id>` (PDF si LibreOffice esta instalado, o un JSON de error 500 si no).

- [ ] **Step 6: Commit**

```bash
git add informe_wom_app.html
git commit -m "Agregar boton Ver Informe (PDF) tras generar en WOM"
```
