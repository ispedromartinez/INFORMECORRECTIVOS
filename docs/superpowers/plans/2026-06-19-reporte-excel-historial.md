# Reporte Excel desde Historial (Tigo + WOM) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar un boton "Reporte" en Historial (Tigo y WOM) que abre un modal con checklist de columnas + checklist de informes, y genera un `.xlsx` descargable con exactamente lo marcado.

**Architecture:** Dos endpoints nuevos en `server.js` (`POST /reporte`, `POST /reporte-wom`) que reusan `dbClimaList`/`dbWomList` ya existentes, filtran por `ids` recibidos y arman la hoja con la libreria `xlsx` (SheetJS). Frontend reusa el patron fetch+blob+download que ya usan `generateDocx()` (Tigo) y `generarInforme()` (WOM, llamado `generarWom` segun el archivo) para `/generar`. Tigo reusa su CSS de modal existente; WOM no tiene infraestructura de modal y se le agrega copiando el patron de Tigo con acento `--mag`.

**Tech Stack:** Express (rutas existentes), libreria `xlsx` (SheetJS, ya en el proyecto como devDependency), HTML/CSS/JS inline tal como el resto de la app — sin build step.

## Global Constraints

- Columnas exportables = solo campos que YA persisten en el historial hoy. Tigo: `fecha`, `codInforme`, `nombreSitio`, `codigoSitio`, `tecnico`, `supervisor`, `numOT`, `photoCount`. WOM: `fechaInicio`, `ticket`, `codInterno`, `instalacion`, `tipoActividad`, `tecnicos`, `photoCount`. No se agregan campos nuevos a lo que se guarda (LPU, tickets individuales, resumen, observaciones quedan fuera, ver spec).
- Boton "Generar Excel" debe quedar deshabilitado si no hay al menos 1 columna Y 1 informe marcados.
- Tigo: la lista de informes del modal respeta el filtro activo de `#histSearch` (mismo criterio que `filterHist`). WOM: no tiene busqueda en Historial, el modal lista todos los informes de `/registro-wom`.
- Sin test runner en el proyecto — verificacion via curl + revision manual en navegador.
- Spec completo: `docs/superpowers/specs/2026-06-19-reporte-excel-historial-design.md`.

---

### Task 1: Backend — endpoints `/reporte` y `/reporte-wom`

**Files:**
- Modify: `server.js:6` (requires)
- Modify: `server.js:824-834` (insertar endpoint despues de `app.delete('/papelera', ...)`)
- Modify: `server.js:1564` (insertar endpoint despues de `app.delete('/papelera-wom', ...)`, antes del bloque `PORT`)
- Modify: `package.json` (mover `xlsx` de `devDependencies` a `dependencies`)

**Interfaces:**
- Produces: `POST /reporte` body `{columns:string[], ids:string[]}` -> `.xlsx` binario (Content-Type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`) o 400 JSON `{error}` si falta columns/ids o ninguna columna es valida.
- Produces: `POST /reporte-wom`, mismo contrato, usando el set de columnas WOM.
- Consumes: `dbClimaList(q)` (linea 133, ya existente, llamar con `null` para traer todo) y `dbWomList()` (linea 872, ya existente, sin parametros).

- [ ] **Step 1: Agregar require de `xlsx`**

En `server.js`, reemplazar:

```js
const path = require('path');
const nodemailer = require('nodemailer');
```

por:

```js
const path = require('path');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
```

- [ ] **Step 2: Mover `xlsx` a `dependencies` en `package.json`**

Reemplazar todo el archivo `package.json` por:

```json
{
  "name": "informe-clima",
  "version": "3.0.0",
  "description": "Generador de Informes Mantenimiento Clima",
  "main": "server.js",
  "scripts": {
    "start": "node server.js"
  },
  "dependencies": {
    "@supabase/supabase-js": "^2.108.0",
    "cors": "^2.8.5",
    "docx": "^8.5.0",
    "express": "^4.18.2",
    "express-rate-limit": "^8.5.2",
    "helmet": "^8.2.0",
    "nodemailer": "^6.9.7",
    "xlsx": "^0.18.5"
  }
}
```

- [ ] **Step 3: Endpoint `/reporte` (Tigo)**

En `server.js`, ubicar el bloque que termina en:

```js
app.delete('/papelera', async (req,res) => {
  const papelera = await dbPapeleraList(null);
  if (papelera.length) {
    await storageRemove(papelera.map(e => `clima/papelera/${e.filename}`));
    papelera.forEach(e => {
      try { const f = path.join(PAPELERA_DIR, e.filename); if (fs.existsSync(f)) fs.unlinkSync(f); } catch(e2) {}
    });
  }
  await dbPapeleraClear();
  res.json({ok:true});
});
```

Justo despues de ese `});` (y antes del comentario `// MÓDULO WOM`), agregar:

```js
const REPORTE_COLUMNAS_CLIMA = {
  fecha: 'Fecha', codInforme: 'Código Informe', nombreSitio: 'Sitio',
  codigoSitio: 'Código Sitio', tecnico: 'Técnico', supervisor: 'Supervisor',
  numOT: 'OT', photoCount: 'Cant. Fotos'
};

app.post('/reporte', async (req, res) => {
  const { columns, ids } = req.body;
  if (!Array.isArray(columns) || !columns.length || !Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'Selecciona al menos una columna y un informe' });
  }
  const validColumns = columns.filter(c => REPORTE_COLUMNAS_CLIMA[c]);
  if (!validColumns.length) return res.status(400).json({ error: 'Columnas inválidas' });
  const all = await dbClimaList(null);
  const selected = all.filter(r => ids.includes(r.id));
  const headerRow = validColumns.map(c => REPORTE_COLUMNAS_CLIMA[c]);
  const dataRows = selected.map(r => validColumns.map(c => r[c] ?? ''));
  const ws = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Historial');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="reporte-tigo.xlsx"');
  res.send(buffer);
});
```

- [ ] **Step 4: Endpoint `/reporte-wom`**

En `server.js`, ubicar el bloque que termina en:

```js
app.delete('/papelera-wom', async (_req, res) => {
  const papelera = await dbPapeleraWomList();
  if (papelera.length) {
    await storageRemove(papelera.map(e => `wom/papelera/${e.filename}`));
    papelera.forEach(e => {
      try { const fp = path.join(PAPELERA_DIR_WOM, e.filename); if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch(e2) {}
    });
  }
  await dbPapeleraWomClear();
  res.json({ok:true});
});
```

Justo despues de ese `});` (y antes de `const PORT = process.env.PORT || 3000;`), agregar:

```js
const REPORTE_COLUMNAS_WOM = {
  fechaInicio: 'Fecha Inicio', ticket: 'Ticket', codInterno: 'Código Interno',
  instalacion: 'Instalación', tipoActividad: 'Tipo Actividad',
  tecnicos: 'Técnicos', photoCount: 'Cant. Fotos'
};

app.post('/reporte-wom', async (req, res) => {
  const { columns, ids } = req.body;
  if (!Array.isArray(columns) || !columns.length || !Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'Selecciona al menos una columna y un informe' });
  }
  const validColumns = columns.filter(c => REPORTE_COLUMNAS_WOM[c]);
  if (!validColumns.length) return res.status(400).json({ error: 'Columnas inválidas' });
  const all = await dbWomList();
  const selected = all.filter(r => ids.includes(r.id));
  const headerRow = validColumns.map(c => REPORTE_COLUMNAS_WOM[c]);
  const dataRows = selected.map(r => validColumns.map(c => r[c] ?? ''));
  const ws = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Historial');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="reporte-wom.xlsx"');
  res.send(buffer);
});
```

- [ ] **Step 5: Instalar dependencia y reiniciar servidor**

```bash
cd "C:/Users/Pedro Luis Martinez/Documents/InformesClima"
npm install
```

Expected: instala `xlsx` como dependency normal (ya estaba en `node_modules` por ser devDependency, esto solo formaliza el `package.json`/`package-lock.json`).

Si el servidor esta corriendo en background (puerto 3000), encontrar el proceso y reiniciarlo:

```bash
netstat -ano | grep ":3000" | grep LISTENING
taskkill //PID <pid encontrado> //F
nohup node server.js > /tmp/reporte_test.log 2>&1 &
```

Esperar 2 segundos y verificar:

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/tigo
```

Expected: `200`

- [ ] **Step 6: Verificar endpoints manualmente**

Generar un informe de prueba Tigo (ajustar payload minimo):

```bash
curl -s -X POST http://localhost:3000/generar -H "Content-Type: application/json" -d '{"codInforme":"TEST-REP-001","fecha":"19/06/2026","nombreSitio":"Sitio QA Reporte","tecnico":"Tec QA","numOT":"OT-1","photos":[]}' -o /tmp/test_rep.docx
ID=$(curl -s "http://localhost:3000/registro?q=TEST-REP-001" | grep -o '"id":"[0-9]*"' | head -1 | grep -o '[0-9]*')
echo "ID=$ID"
curl -s -o /tmp/test_reporte.xlsx -w "%{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d "{\"columns\":[\"nombreSitio\",\"numOT\"],\"ids\":[\"$ID\"]}"
file /tmp/test_reporte.xlsx
```

Expected: `200`, `file` reporta `Microsoft Excel 2007+` (o similar zip/xlsx).

Probar validacion de body vacio:

```bash
curl -s -w "\n%{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d '{}'
```

Expected: `400` con `{"error":"Selecciona al menos una columna y un informe"}`.

Limpiar datos de prueba:

```bash
curl -s -X DELETE http://localhost:3000/registro/$ID > /dev/null
curl -s -X DELETE http://localhost:3000/papelera > /dev/null
rm -f /tmp/test_rep.docx /tmp/test_reporte.xlsx
```

- [ ] **Step 7: Commit**

```bash
git add server.js package.json package-lock.json
git commit -m "$(cat <<'EOF'
Agregar endpoints /reporte y /reporte-wom para exportar Excel

Reusan dbClimaList/dbWomList, filtran por ids recibidos y arman la
hoja con xlsx (SheetJS) usando solo las columnas pedidas. xlsx pasa
de devDependency a dependency porque ahora se usa en runtime.
EOF
)"
```

---

### Task 2: Frontend Tigo — boton, modal y JS de Reporte

**Files:**
- Modify: `informe_clima_app.html:577-580` (boton junto a refresh en Historial)
- Modify: `informe_clima_app.html` (nuevo modal, justo despues del modal de email, antes del comentario `<!-- DELETE CONFIRM -->`)
- Modify: `informe_clima_app.html` (JS: nuevas funciones, cerca de `filterHist`/`renderHist`)

**Interfaces:**
- Consumes: `POST /reporte` de Task 1 (body `{columns, ids}`, devuelve `.xlsx` binario o 400 JSON `{error}`); variables globales ya existentes `allRecords` (array de informes) y funcion `norm()` (linea 892).
- Produces: nada (hoja terminal para Tigo; Task 3 es independiente en otro archivo).

- [ ] **Step 1: Agregar boton "Reporte" junto al refresh**

En `informe_clima_app.html`, reemplazar:

```html
  <div class="hist-search-wrap">
    <input type="text" id="histSearch" placeholder="🔍 Buscar por sitio, código, técnico, OT..." oninput="filterHist(this.value)">
    <button class="btn-refresh" onclick="loadHist()">↻</button>
  </div>
```

por:

```html
  <div class="hist-search-wrap">
    <input type="text" id="histSearch" placeholder="🔍 Buscar por sitio, código, técnico, OT..." oninput="filterHist(this.value)">
    <button class="btn-refresh" onclick="loadHist()">↻</button>
    <button class="btn-refresh" onclick="openReporteModal()">📊 Reporte</button>
  </div>
```

- [ ] **Step 2: Agregar el modal de Reporte**

En `informe_clima_app.html`, ubicar el cierre del modal de email:

```html
    <button class="btn-gen" style="font-size:14px;padding:13px" onclick="sendEmail()">📤 Enviar</button>
    <div class="status" id="emailStatus"></div>
  </div>
</div>

<!-- DELETE CONFIRM -->
```

Reemplazar por (agrega el modal nuevo entre el email modal y el delete confirm):

```html
    <button class="btn-gen" style="font-size:14px;padding:13px" onclick="sendEmail()">📤 Enviar</button>
    <div class="status" id="emailStatus"></div>
  </div>
</div>

<!-- REPORTE MODAL -->
<div class="modal-overlay" id="reporteModal">
  <div class="modal-box" style="max-width:520px;max-height:80vh;overflow-y:auto;">
    <div class="modal-title">📊 Generar reporte Excel<button class="modal-close" onclick="closeReporteModal()">✕</button></div>
    <div class="modal-info">Elige las columnas y los informes a incluir.</div>
    <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;margin:10px 0 6px">Columnas</div>
    <div id="reporteColumnas" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px;"></div>
    <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;margin:10px 0 6px">Informes a incluir</div>
    <div id="reporteInformes" style="max-height:220px;overflow-y:auto;border:1px solid var(--border);border-radius:8px;padding:8px;margin-bottom:14px;"></div>
    <button class="btn-gen" id="btnGenerarReporte" style="font-size:14px;padding:13px" onclick="generarReporte()">📥 Generar Excel</button>
  </div>
</div>

<!-- DELETE CONFIRM -->
```

- [ ] **Step 3: Agregar JS de Reporte**

En `informe_clima_app.html`, ubicar la funcion `filterHist`:

```js
function filterHist(q){
  if(!q){renderHist(allRecords);return;}
  const nq=norm(q);
  renderHist(allRecords.filter(r=>norm(r.nombreSitio||'').includes(nq)||norm(r.codInforme||'').includes(nq)||norm(r.tecnico||'').includes(nq)||norm(r.numOT||'').includes(nq)));
}
```

Reemplazar por (agrega las funciones de Reporte justo despues):

```js
function filterHist(q){
  if(!q){renderHist(allRecords);return;}
  const nq=norm(q);
  renderHist(allRecords.filter(r=>norm(r.nombreSitio||'').includes(nq)||norm(r.codInforme||'').includes(nq)||norm(r.tecnico||'').includes(nq)||norm(r.numOT||'').includes(nq)));
}

// ══ REPORTE EXCEL ══
const REPORTE_COLUMNAS=[
  {key:'fecha',label:'Fecha'},{key:'codInforme',label:'Código Informe'},
  {key:'nombreSitio',label:'Sitio'},{key:'codigoSitio',label:'Código Sitio'},
  {key:'tecnico',label:'Técnico'},{key:'supervisor',label:'Supervisor'},
  {key:'numOT',label:'OT'},{key:'photoCount',label:'Cant. Fotos'}
];
function getVisibleHistRecords(){
  const q=document.getElementById('histSearch')?.value||'';
  if(!q)return allRecords;
  const nq=norm(q);
  return allRecords.filter(r=>norm(r.nombreSitio||'').includes(nq)||norm(r.codInforme||'').includes(nq)||norm(r.tecnico||'').includes(nq)||norm(r.numOT||'').includes(nq));
}
function updateReporteBtnState(){
  const anyCol=document.querySelector('.reporte-col:checked');
  const anyInf=document.querySelector('.reporte-inf:checked');
  document.getElementById('btnGenerarReporte').disabled=!(anyCol&&anyInf);
}
function openReporteModal(){
  const colsDiv=document.getElementById('reporteColumnas');
  colsDiv.innerHTML=REPORTE_COLUMNAS.map(c=>`<label style="display:flex;align-items:center;gap:5px;font-size:12px;"><input type="checkbox" class="reporte-col" value="${c.key}" checked>${c.label}</label>`).join('');
  const recs=getVisibleHistRecords();
  const infDiv=document.getElementById('reporteInformes');
  infDiv.innerHTML=recs.length?recs.map(r=>`<label style="display:flex;align-items:center;gap:6px;font-size:12px;padding:5px 0;"><input type="checkbox" class="reporte-inf" value="${r.id}" checked>${r.nombreSitio||'—'} · ${r.codInforme||''} · ${r.fecha||''}</label>`).join(''):'<p style="font-size:12px;color:var(--muted)">No hay informes para incluir.</p>';
  colsDiv.onchange=updateReporteBtnState;
  infDiv.onchange=updateReporteBtnState;
  updateReporteBtnState();
  document.getElementById('reporteModal').classList.add('open');
}
function closeReporteModal(){document.getElementById('reporteModal').classList.remove('open');}
async function generarReporte(){
  const columns=[...document.querySelectorAll('.reporte-col:checked')].map(c=>c.value);
  const ids=[...document.querySelectorAll('.reporte-inf:checked')].map(c=>c.value);
  if(!columns.length||!ids.length)return;
  const btn=document.getElementById('btnGenerarReporte');
  btn.disabled=true;
  try{
    const resp=await fetch(SERVER+'/reporte',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({columns,ids})});
    if(!resp.ok){const e=await resp.json();throw new Error(e.error||'Error del servidor');}
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=`reporte-tigo-${Date.now()}.xlsx`;a.click();URL.revokeObjectURL(url);
    closeReporteModal();
  }catch(err){alert('Error: '+err.message);}
  finally{btn.disabled=false;}
}
```

- [ ] **Step 4: Verificar que el servidor sigue sirviendo el archivo**

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/tigo
curl -s http://localhost:3000/tigo | grep -c "reporteModal"
```

Expected: `200` y `1` (el modal aparece una vez en el HTML).

- [ ] **Step 5: Revision visual manual**

Abrir `http://localhost:3000/tigo`, ir a Historial, click "📊 Reporte": confirmar que el modal abre con columnas marcadas y la lista de informes visibles actuales. Desmarcar una columna y un informe, click "Generar Excel", confirmar que descarga un `.xlsx` y que el boton se deshabilita si desmarcas todas las columnas o todos los informes.

- [ ] **Step 6: Commit**

```bash
git add informe_clima_app.html
git commit -m "$(cat <<'EOF'
Agregar boton Reporte (export Excel) en Historial Tigo

Modal con checklist de columnas + informes visibles, genera .xlsx
via POST /reporte reusando el patron fetch+blob de generateDocx.
EOF
)"
```

---

### Task 3: Frontend WOM — modal CSS nuevo, boton y JS de Reporte

**Files:**
- Modify: `informe_wom_app.html` (CSS: agregar clases de modal, no existen hoy)
- Modify: `informe_wom_app.html:410-416` (boton en `.hist-header`)
- Modify: `informe_wom_app.html` (nuevo modal markup)
- Modify: `informe_wom_app.html` (JS: nuevas funciones, cerca de `loadHistorial`)

**Interfaces:**
- Consumes: `POST /reporte-wom` de Task 1 (body `{columns, ids}` -> `.xlsx` o 400 `{error}`); `GET /registro-wom` ya existente; funcion `toast(msg,type)` ya existente (linea 538).
- Produces: nada (hoja terminal para WOM).

- [ ] **Step 1: Agregar CSS de modal (no existe en este archivo)**

En `informe_wom_app.html`, ubicar:

```css
/* ── TOAST ── */
```

Reemplazar por (agrega las clases de modal justo antes del toast):

```css
/* ── MODAL ── */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:500;display:none;align-items:flex-end;justify-content:center;backdrop-filter:blur(2px);}
.modal-overlay.open{display:flex;}
.modal-box{background:#fff;border-radius:20px 20px 0 0;width:100%;max-width:520px;padding:22px;max-height:80vh;overflow-y:auto;}
.modal-title{font-size:15px;font-weight:700;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center;}
.modal-close{background:#f1f5f9;border:none;font-size:16px;cursor:pointer;color:var(--muted);width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;}
.modal-info{font-size:12px;color:var(--muted);margin-bottom:14px;padding:9px 13px;background:#f8fafc;border-radius:8px;}

/* ── TOAST ── */
```

- [ ] **Step 2: Agregar boton "Reporte" en el header de Historial**

En `informe_wom_app.html`, reemplazar:

```html
    <div class="hist-header">
      <span class="hist-title">Informes Generados</span>
      <span class="hist-count" id="hist-count">Cargando...</span>
    </div>
```

por:

```html
    <div class="hist-header">
      <span class="hist-title">Informes Generados</span>
      <span style="display:flex;align-items:center;gap:8px;">
        <span class="hist-count" id="hist-count">Cargando...</span>
        <button class="btn-dl" onclick="openReporteModal()">📊 Reporte</button>
      </span>
    </div>
```

- [ ] **Step 3: Agregar el modal de Reporte**

En `informe_wom_app.html`, ubicar el cierre del panel de Historial:

```html
    <div id="hist-body"><div class="hist-empty">Cargando historial...</div></div>
  </div>
  <div class="actions" style="margin-top:12px;">
    <button class="btn btn-secondary" onclick="goTab(0)">← Inicio</button>
  </div>
</div>
```

Reemplazar por (agrega el modal justo despues de cerrar el panel):

```html
    <div id="hist-body"><div class="hist-empty">Cargando historial...</div></div>
  </div>
  <div class="actions" style="margin-top:12px;">
    <button class="btn btn-secondary" onclick="goTab(0)">← Inicio</button>
  </div>
</div>

<!-- REPORTE MODAL -->
<div class="modal-overlay" id="reporteModal">
  <div class="modal-box">
    <div class="modal-title">📊 Generar reporte Excel<button class="modal-close" onclick="closeReporteModal()">✕</button></div>
    <div class="modal-info">Elige las columnas y los informes a incluir.</div>
    <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;margin:10px 0 6px">Columnas</div>
    <div id="reporteColumnas" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px;"></div>
    <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;margin:10px 0 6px">Informes a incluir</div>
    <div id="reporteInformes" style="max-height:220px;overflow-y:auto;border:1px solid var(--border);border-radius:8px;padding:8px;margin-bottom:14px;"></div>
    <button class="btn-primary btn" id="btnGenerarReporte" style="width:100%;justify-content:center;" onclick="generarReporte()">📥 Generar Excel</button>
  </div>
</div>
```

- [ ] **Step 4: Agregar JS de Reporte**

En `informe_wom_app.html`, ubicar la funcion `loadHistorial` completa:

```js
async function loadHistorial() {
  const body = g('hist-body');
  if (!body) return;
  body.innerHTML = '<div class="hist-empty">Cargando...</div>';
  try {
    const data = await fetch('/registro-wom').then(r=>r.json());
    g('hist-count').textContent = `${data.length} informe(s)`;
    if (!data.length) { body.innerHTML = '<div class="hist-empty">No hay informes generados aún.</div>'; return; }
    body.innerHTML = data.map(e => `
      <div class="hist-row">
        <div class="hist-info">
          <div class="hist-ticket">${e.ticket||'—'}</div>
          <div class="hist-sub">
            ${e.tipoActividad||''} · ${e.instalacion||''} · ${e.fechaInicio||''}
            ${e.codInterno ? `<span class="hist-inc"> · INC-${e.codInterno}</span>` : ''}
          </div>
          <div class="hist-sub">${(e.tecnicos||'').split(',').join(' / ')} · ${e.photoCount||0} foto(s)</div>
        </div>
        <div class="hist-actions">
          <button class="btn-dl" onclick="descargar('${e.id}')">↓ Descargar</button>
          <button class="btn-del" onclick="eliminar('${e.id}',this)">🗑</button>
        </div>
      </div>`).join('');
  } catch { body.innerHTML = '<div class="hist-empty">Error al cargar historial.</div>'; }
}
```

Reemplazar por (agrega las funciones de Reporte justo despues, sin modificar `loadHistorial`):

```js
async function loadHistorial() {
  const body = g('hist-body');
  if (!body) return;
  body.innerHTML = '<div class="hist-empty">Cargando...</div>';
  try {
    const data = await fetch('/registro-wom').then(r=>r.json());
    g('hist-count').textContent = `${data.length} informe(s)`;
    if (!data.length) { body.innerHTML = '<div class="hist-empty">No hay informes generados aún.</div>'; return; }
    body.innerHTML = data.map(e => `
      <div class="hist-row">
        <div class="hist-info">
          <div class="hist-ticket">${e.ticket||'—'}</div>
          <div class="hist-sub">
            ${e.tipoActividad||''} · ${e.instalacion||''} · ${e.fechaInicio||''}
            ${e.codInterno ? `<span class="hist-inc"> · INC-${e.codInterno}</span>` : ''}
          </div>
          <div class="hist-sub">${(e.tecnicos||'').split(',').join(' / ')} · ${e.photoCount||0} foto(s)</div>
        </div>
        <div class="hist-actions">
          <button class="btn-dl" onclick="descargar('${e.id}')">↓ Descargar</button>
          <button class="btn-del" onclick="eliminar('${e.id}',this)">🗑</button>
        </div>
      </div>`).join('');
  } catch { body.innerHTML = '<div class="hist-empty">Error al cargar historial.</div>'; }
}

// ── Reporte Excel ────────────────────────────────────────────────
const REPORTE_COLUMNAS_WOM=[
  {key:'fechaInicio',label:'Fecha Inicio'},{key:'ticket',label:'Ticket'},
  {key:'codInterno',label:'Código Interno'},{key:'instalacion',label:'Instalación'},
  {key:'tipoActividad',label:'Tipo Actividad'},{key:'tecnicos',label:'Técnicos'},
  {key:'photoCount',label:'Cant. Fotos'}
];
let reporteWomData=[];
function updateReporteBtnState(){
  const anyCol=document.querySelector('.reporte-col:checked');
  const anyInf=document.querySelector('.reporte-inf:checked');
  document.getElementById('btnGenerarReporte').disabled=!(anyCol&&anyInf);
}
async function openReporteModal(){
  const colsDiv=document.getElementById('reporteColumnas');
  colsDiv.innerHTML=REPORTE_COLUMNAS_WOM.map(c=>`<label style="display:flex;align-items:center;gap:5px;font-size:12px;"><input type="checkbox" class="reporte-col" value="${c.key}" checked>${c.label}</label>`).join('');
  const infDiv=document.getElementById('reporteInformes');
  infDiv.innerHTML='<p style="font-size:12px;color:var(--muted)">Cargando...</p>';
  document.getElementById('reporteModal').classList.add('open');
  try{ reporteWomData=await fetch('/registro-wom').then(r=>r.json()); }catch{ reporteWomData=[]; }
  infDiv.innerHTML=reporteWomData.length?reporteWomData.map(e=>`<label style="display:flex;align-items:center;gap:6px;font-size:12px;padding:5px 0;"><input type="checkbox" class="reporte-inf" value="${e.id}" checked>${e.ticket||'—'} · ${e.instalacion||''} · ${e.fechaInicio||''}</label>`).join(''):'<p style="font-size:12px;color:var(--muted)">No hay informes para incluir.</p>';
  colsDiv.onchange=updateReporteBtnState;
  infDiv.onchange=updateReporteBtnState;
  updateReporteBtnState();
}
function closeReporteModal(){document.getElementById('reporteModal').classList.remove('open');}
async function generarReporte(){
  const columns=[...document.querySelectorAll('.reporte-col:checked')].map(c=>c.value);
  const ids=[...document.querySelectorAll('.reporte-inf:checked')].map(c=>c.value);
  if(!columns.length||!ids.length)return;
  const btn=document.getElementById('btnGenerarReporte');
  btn.disabled=true;
  try{
    const resp=await fetch('/reporte-wom',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({columns,ids})});
    if(!resp.ok){const e=await resp.json();throw new Error(e.error||'Error del servidor');}
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=`reporte-wom-${Date.now()}.xlsx`;a.click();URL.revokeObjectURL(url);
    closeReporteModal();
    toast('✓ Reporte generado','ok');
  }catch(err){ toast('Error: '+err.message,'err'); }
  finally{ btn.disabled=false; }
}
```

- [ ] **Step 5: Verificar que el servidor sigue sirviendo el archivo**

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/wom
curl -s http://localhost:3000/wom | grep -c "reporteModal"
```

Expected: `200` y `1`.

- [ ] **Step 6: Revision visual manual**

Abrir `http://localhost:3000/wom`, ir a Historial, click "📊 Reporte": confirmar que el modal abre (fondo blur, caja blanca abajo) con columnas marcadas y la lista completa de informes. Desmarcar todo y confirmar que el boton "Generar Excel" se deshabilita. Marcar algo y confirmar que descarga `.xlsx`.

- [ ] **Step 7: Commit**

```bash
git add informe_wom_app.html
git commit -m "$(cat <<'EOF'
Agregar boton Reporte (export Excel) en Historial WOM

Se agrega infraestructura de modal (no existia en este archivo) y
el mismo flujo de checklist columnas+informes que Tigo, generando
.xlsx via POST /reporte-wom.
EOF
)"
```
