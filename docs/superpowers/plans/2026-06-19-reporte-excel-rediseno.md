# Reporte Excel: rediseño (periodo en vez de informes) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reemplazar el checklist de informes del modal "Reporte" (Tigo + WOM) por un selector de periodo (Todos/Mes/Semana/Fecha) que filtra por `fechaCreacion`, pasar el checklist de columnas a vertical, y restylear el modal reusando el patron de tarjeta (`.card`/`.card-title`) que ya usa el resto del sistema.

**Architecture:** Backend: se reemplaza el contrato `{columns, ids}` por `{columns, periodo}` en `/reporte` y `/reporte-wom`; una funcion compartida `computeRangoPeriodo(periodo)` calcula el rango de fechas (o `null` para "todos") y ambos endpoints filtran `dbClimaList(null)`/`dbWomList()` por `fechaCreacion` cayendo en ese rango antes de armar la hoja. Frontend: cada app reemplaza su modal+JS de Reporte; Tigo y WOM quedan con codigo estructuralmente igual (mismos nombres de funcion), solo cambia el accent color y el detalle de fetch (`SERVER+'/reporte'`+`alert` en Tigo vs `'/reporte-wom'`+`toast` en WOM, ya establecido en el feature anterior).

**Tech Stack:** Express + `xlsx` (sin cambios de libreria), HTML/CSS/JS inline existente, reutiliza clases `.card`/`.card-title`/`.modal-overlay`/`.field`/`.row2` (Tigo) / `.row-2` (WOM) ya definidas.

## Global Constraints

- Las columnas exportables y sus headers en español NO cambian (mismo `REPORTE_COLUMNAS_CLIMA`/`REPORTE_COLUMNAS_WOM` ya existente).
- El filtro de periodo usa `fechaCreacion` (timestamp ISO), nunca el campo `fecha`/`fechaInicio` de texto libre.
- No se persiste ningun campo nuevo — los rangos se calculan al vuelo en el endpoint.
- El acento de color de cada app no cambia: `var(--blue)` en Tigo, `var(--mag)` en WOM. Ninguna sección introduce un segundo color.
- Sin test runner — verificacion via curl + revision manual.
- Spec completo: `docs/superpowers/specs/2026-06-19-reporte-excel-rediseno-design.md`.

---

### Task 1: Backend — contrato `{columns, periodo}` en `/reporte` y `/reporte-wom`

**Files:**
- Modify: `server.js:837-861` (bloque `REPORTE_COLUMNAS_CLIMA` + endpoint `/reporte`)
- Modify: `server.js:1593-1617` (bloque `REPORTE_COLUMNAS_WOM` + endpoint `/reporte-wom`)

**Interfaces:**
- Produces: `function computeRangoPeriodo(periodo)` -> `{inicio:Date|null, fin:Date|null}` o `{error:string}`. `inicio`/`fin` ambos `null` significa "todos, sin filtro". Usada por Task 2 y Task 3 indirectamente (son ellas las que construyen el `periodo` que este endpoint consume, no llaman la funcion directamente — la funcion vive solo en `server.js`).
- Consumes: `dbClimaList(q)` (linea 133) y `dbWomList()` (linea 872), sin cambios en sus firmas.

- [ ] **Step 1: Reemplazar bloque `/reporte` (Tigo)**

En `server.js`, reemplazar exactamente:

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

por:

```js
function computeRangoPeriodo(periodo) {
  if (!periodo || typeof periodo !== 'object') return { error: 'Periodo inválido' };
  const { tipo } = periodo;
  if (tipo === 'todos') return { inicio: null, fin: null };
  if (tipo === 'mes') {
    if (!periodo.mes || !/^\d{4}-\d{2}$/.test(periodo.mes)) return { error: 'Mes inválido' };
    const [y, m] = periodo.mes.split('-').map(Number);
    return { inicio: new Date(y, m - 1, 1, 0, 0, 0, 0), fin: new Date(y, m, 0, 23, 59, 59, 999) };
  }
  if (tipo === 'semana') {
    if (!periodo.fecha) return { error: 'Fecha inválida' };
    const d = new Date(periodo.fecha + 'T00:00:00');
    if (isNaN(d)) return { error: 'Fecha inválida' };
    const dow = (d.getDay() + 6) % 7;
    const inicio = new Date(d); inicio.setDate(d.getDate() - dow); inicio.setHours(0, 0, 0, 0);
    const fin = new Date(inicio); fin.setDate(inicio.getDate() + 6); fin.setHours(23, 59, 59, 999);
    return { inicio, fin };
  }
  if (tipo === 'fecha') {
    if (!periodo.desde || !periodo.hasta) return { error: 'Rango de fecha inválido' };
    const inicio = new Date(periodo.desde + 'T00:00:00');
    const fin = new Date(periodo.hasta + 'T23:59:59.999');
    if (isNaN(inicio) || isNaN(fin)) return { error: 'Rango de fecha inválido' };
    return { inicio, fin };
  }
  return { error: 'Tipo de periodo inválido' };
}

const REPORTE_COLUMNAS_CLIMA = {
  fecha: 'Fecha', codInforme: 'Código Informe', nombreSitio: 'Sitio',
  codigoSitio: 'Código Sitio', tecnico: 'Técnico', supervisor: 'Supervisor',
  numOT: 'OT', photoCount: 'Cant. Fotos'
};

app.post('/reporte', async (req, res) => {
  const { columns, periodo } = req.body;
  if (!Array.isArray(columns) || !columns.length) {
    return res.status(400).json({ error: 'Selecciona al menos una columna' });
  }
  const validColumns = columns.filter(c => REPORTE_COLUMNAS_CLIMA[c]);
  if (!validColumns.length) return res.status(400).json({ error: 'Columnas inválidas' });
  const rango = computeRangoPeriodo(periodo);
  if (rango.error) return res.status(400).json({ error: rango.error });
  const all = await dbClimaList(null);
  const selected = rango.inicio
    ? all.filter(r => { const f = new Date(r.fechaCreacion); return f >= rango.inicio && f <= rango.fin; })
    : all;
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

- [ ] **Step 2: Reemplazar bloque `/reporte-wom`**

En `server.js`, reemplazar exactamente:

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

por:

```js
const REPORTE_COLUMNAS_WOM = {
  fechaInicio: 'Fecha Inicio', ticket: 'Ticket', codInterno: 'Código Interno',
  instalacion: 'Instalación', tipoActividad: 'Tipo Actividad',
  tecnicos: 'Técnicos', photoCount: 'Cant. Fotos'
};

app.post('/reporte-wom', async (req, res) => {
  const { columns, periodo } = req.body;
  if (!Array.isArray(columns) || !columns.length) {
    return res.status(400).json({ error: 'Selecciona al menos una columna' });
  }
  const validColumns = columns.filter(c => REPORTE_COLUMNAS_WOM[c]);
  if (!validColumns.length) return res.status(400).json({ error: 'Columnas inválidas' });
  const rango = computeRangoPeriodo(periodo);
  if (rango.error) return res.status(400).json({ error: rango.error });
  const all = await dbWomList();
  const selected = rango.inicio
    ? all.filter(r => { const f = new Date(r.fechaCreacion); return f >= rango.inicio && f <= rango.fin; })
    : all;
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

Nota: `computeRangoPeriodo` se define una sola vez (en el Step 1, antes de `/reporte`) y `/reporte-wom` la usa tambien — las declaraciones `function` en JavaScript se "hoistean" al alcance del modulo, asi que es valida aunque `/reporte-wom` este mas abajo en el archivo. No declarar la funcion una segunda vez en el Step 2.

- [ ] **Step 3: Reiniciar servidor**

```bash
netstat -ano | grep ":3000" | grep LISTENING
```

Tomar el PID de la salida y:

```bash
taskkill //PID <pid> //F
cd "C:/Users/Pedro Luis Martinez/Documents/InformesClima"
nohup node server.js > /tmp/reporte_v2_test.log 2>&1 &
```

Esperar 2 segundos y verificar:

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/tigo
```

Expected: `200`. Revisar `/tmp/reporte_v2_test.log` no tenga errores de sintaxis.

- [ ] **Step 4: Verificar el nuevo contrato con curl**

Generar un informe de prueba y exportar con `periodo:{tipo:'todos'}`:

```bash
curl -s -X POST http://localhost:3000/generar -H "Content-Type: application/json" -d '{"codInforme":"TEST-PER-001","fecha":"19/06/2026","nombreSitio":"Sitio Periodo QA","tecnico":"Tec QA","numOT":"OT-1","photos":[]}' -o /tmp/test_per.docx
curl -s -o /tmp/test_per.xlsx -w "todos HTTP %{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d '{"columns":["nombreSitio"],"periodo":{"tipo":"todos"}}'
file /tmp/test_per.xlsx
```

Expected: `200`, `Microsoft Excel 2007+`.

Probar `tipo:"semana"` con la fecha de hoy (debe incluir el informe recien creado, porque su `fechaCreacion` es ahora):

```bash
HOY=$(date +%Y-%m-%d)
curl -s -o /tmp/test_semana.xlsx -w "semana HTTP %{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d "{\"columns\":[\"nombreSitio\"],\"periodo\":{\"tipo\":\"semana\",\"fecha\":\"$HOY\"}}"
mkdir -p /tmp/xc_sem && cd /tmp/xc_sem && unzip -o -q /tmp/test_semana.xlsx && grep -l "Sitio Periodo QA" xl/worksheets/sheet1.xml
cd "C:/Users/Pedro Luis Martinez/Documents/InformesClima"
```

Expected: `200`, y el grep encuentra "Sitio Periodo QA" en `xl/worksheets/sheet1.xml` (confirma que la semana de hoy incluyo el informe).

Probar un rango de fecha que NO cubre hoy (debe excluirlo):

```bash
curl -s -o /tmp/test_fecha_vieja.xlsx -w "fecha-vieja HTTP %{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d '{"columns":["nombreSitio"],"periodo":{"tipo":"fecha","desde":"2020-01-01","hasta":"2020-01-31"}}'
mkdir -p /tmp/xc_old && cd /tmp/xc_old && unzip -o -q /tmp/test_fecha_vieja.xlsx && grep -c "Sitio Periodo QA" xl/worksheets/sheet1.xml
cd "C:/Users/Pedro Luis Martinez/Documents/InformesClima"
```

Expected: `200`, y el `grep -c` da `0` (el informe de hoy no aparece en un rango de enero 2020).

Probar validacion (periodo invalido):

```bash
curl -s -w "\nHTTP %{http_code}\n" -X POST http://localhost:3000/reporte -H "Content-Type: application/json" -d '{"columns":["nombreSitio"],"periodo":{"tipo":"mes"}}'
```

Expected: `400` con `{"error":"Mes inválido"}`.

Limpiar:

```bash
ID=$(curl -s "http://localhost:3000/registro?q=TEST-PER-001" | grep -o '"id":"[0-9]*"' | head -1 | grep -o '[0-9]*')
curl -s -X DELETE http://localhost:3000/registro/$ID > /dev/null
curl -s -X DELETE http://localhost:3000/papelera > /dev/null
rm -f /tmp/test_per.docx /tmp/test_per.xlsx /tmp/test_semana.xlsx /tmp/test_fecha_vieja.xlsx
rm -rf /tmp/xc_sem /tmp/xc_old
```

- [ ] **Step 5: Commit**

```bash
git add server.js
git commit -m "$(cat <<'EOF'
Cambiar /reporte y /reporte-wom a filtrar por periodo en vez de ids

Nuevo contrato {columns, periodo} con tipo todos/mes/semana/fecha.
computeRangoPeriodo() calcula el rango y filtra por fechaCreacion
(timestamp real de generacion), no por el campo fecha de texto
libre. Sin persistir nada nuevo.
EOF
)"
```

---

### Task 2: Frontend Tigo — modal con periodo + columnas verticales

**Files:**
- Modify: `informe_clima_app.html:302` (CSS, agregar `.rep-tab`/`.rep-tab.active` despues de `.modal-info`)
- Modify: `informe_clima_app.html:651-662` (markup del modal de Reporte)
- Modify: `informe_clima_app.html:1614-1659` (JS de Reporte completo)

**Interfaces:**
- Consumes: `POST /reporte` de Task 1 (body `{columns, periodo}`, `periodo` = `{tipo:'todos'}` | `{tipo:'mes',mes:'YYYY-MM'}` | `{tipo:'semana',fecha:'YYYY-MM-DD'}` | `{tipo:'fecha',desde:'YYYY-MM-DD',hasta:'YYYY-MM-DD'}`).
- Produces: nada nuevo para otras tasks (Task 3 es WOM, archivo separado, mismo patron pero independiente).

- [ ] **Step 1: Agregar CSS de tabs de periodo**

En `informe_clima_app.html`, reemplazar:

```css
.modal-info{font-size:12px;color:var(--muted);margin-bottom:14px;padding:9px 13px;background:#f8fafc;border-radius:8px;}
```

por:

```css
.modal-info{font-size:12px;color:var(--muted);margin-bottom:14px;padding:9px 13px;background:#f8fafc;border-radius:8px;}
.rep-tab{flex:1;text-align:center;padding:7px 4px;border-radius:8px;font-size:11px;font-weight:600;cursor:pointer;color:var(--muted);}
.rep-tab.active{background:#fff;color:var(--blue);box-shadow:0 1px 3px rgba(0,0,0,.12);}
```

- [ ] **Step 2: Reemplazar el modal de Reporte**

En `informe_clima_app.html`, reemplazar exactamente:

```html
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
```

por:

```html
<!-- REPORTE MODAL -->
<div class="modal-overlay" id="reporteModal">
  <div class="modal-box" style="max-width:480px;max-height:85vh;overflow-y:auto;">
    <div class="modal-title">📊 Generar reporte Excel<button class="modal-close" onclick="closeReporteModal()">✕</button></div>
    <div class="card" style="margin-bottom:14px;">
      <div class="card-title">Periodo</div>
      <div style="display:flex;gap:4px;background:#EDEEF2;border-radius:10px;padding:4px;margin-bottom:10px;">
        <div class="rep-tab active" data-tipo="todos" onclick="selectPeriodoTipo('todos')">Todos</div>
        <div class="rep-tab" data-tipo="mes" onclick="selectPeriodoTipo('mes')">Mes</div>
        <div class="rep-tab" data-tipo="semana" onclick="selectPeriodoTipo('semana')">Semana</div>
        <div class="rep-tab" data-tipo="fecha" onclick="selectPeriodoTipo('fecha')">Fecha</div>
      </div>
      <div id="periodoInputMes" class="field" style="display:none;margin-bottom:0;"><label>Mes</label><input type="month" id="repMes" oninput="updateReporteBtnState()" style="width:100%;border:1.5px solid var(--border);border-radius:10px;padding:10px 13px;font-size:14px;font-family:inherit;"></div>
      <div id="periodoInputSemana" class="field" style="display:none;margin-bottom:0;"><label>Cualquier día de la semana</label><input type="date" id="repSemana" oninput="updateReporteBtnState()"></div>
      <div id="periodoInputFecha" style="display:none;">
        <div class="row2">
          <div class="field" style="margin-bottom:0;"><label>Desde</label><input type="date" id="repDesde" oninput="updateReporteBtnState()"></div>
          <div class="field" style="margin-bottom:0;"><label>Hasta</label><input type="date" id="repHasta" oninput="updateReporteBtnState()"></div>
        </div>
      </div>
    </div>
    <div class="card" style="margin-bottom:14px;">
      <div class="card-title">Columnas</div>
      <div id="reporteColumnas" style="display:flex;flex-direction:column;gap:9px;"></div>
    </div>
    <button class="btn-gen" id="btnGenerarReporte" style="font-size:14px;padding:13px" onclick="generarReporte()">📥 Generar Excel</button>
  </div>
</div>
```

- [ ] **Step 3: Reemplazar el JS de Reporte**

En `informe_clima_app.html`, reemplazar exactamente (todo el bloque entre el comentario `// ══ REPORTE EXCEL ══` y el `}` que cierra `generarReporte`, justo antes de `function renderHist`):

```js
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

por:

```js
// ══ REPORTE EXCEL ══
const REPORTE_COLUMNAS=[
  {key:'fecha',label:'Fecha'},{key:'codInforme',label:'Código Informe'},
  {key:'nombreSitio',label:'Sitio'},{key:'codigoSitio',label:'Código Sitio'},
  {key:'tecnico',label:'Técnico'},{key:'supervisor',label:'Supervisor'},
  {key:'numOT',label:'OT'},{key:'photoCount',label:'Cant. Fotos'}
];
let periodoTipoActivo='todos';
function selectPeriodoTipo(tipo){
  periodoTipoActivo=tipo;
  document.querySelectorAll('.rep-tab').forEach(t=>t.classList.toggle('active',t.dataset.tipo===tipo));
  document.getElementById('periodoInputMes').style.display=tipo==='mes'?'block':'none';
  document.getElementById('periodoInputSemana').style.display=tipo==='semana'?'block':'none';
  document.getElementById('periodoInputFecha').style.display=tipo==='fecha'?'block':'none';
  updateReporteBtnState();
}
function buildPeriodoPayload(){
  if(periodoTipoActivo==='mes')return{tipo:'mes',mes:document.getElementById('repMes').value};
  if(periodoTipoActivo==='semana')return{tipo:'semana',fecha:document.getElementById('repSemana').value};
  if(periodoTipoActivo==='fecha')return{tipo:'fecha',desde:document.getElementById('repDesde').value,hasta:document.getElementById('repHasta').value};
  return{tipo:'todos'};
}
function periodoValido(){
  if(periodoTipoActivo==='mes')return!!document.getElementById('repMes').value;
  if(periodoTipoActivo==='semana')return!!document.getElementById('repSemana').value;
  if(periodoTipoActivo==='fecha')return!!document.getElementById('repDesde').value&&!!document.getElementById('repHasta').value;
  return true;
}
function updateReporteBtnState(){
  const anyCol=document.querySelector('.reporte-col:checked');
  document.getElementById('btnGenerarReporte').disabled=!(anyCol&&periodoValido());
}
function openReporteModal(){
  periodoTipoActivo='todos';
  document.querySelectorAll('.rep-tab').forEach(t=>t.classList.toggle('active',t.dataset.tipo==='todos'));
  document.getElementById('periodoInputMes').style.display='none';
  document.getElementById('periodoInputSemana').style.display='none';
  document.getElementById('periodoInputFecha').style.display='none';
  document.getElementById('repMes').value='';
  document.getElementById('repSemana').value='';
  document.getElementById('repDesde').value='';
  document.getElementById('repHasta').value='';
  const colsDiv=document.getElementById('reporteColumnas');
  colsDiv.innerHTML=REPORTE_COLUMNAS.map(c=>`<label style="display:flex;align-items:center;gap:7px;font-size:13px;"><input type="checkbox" class="reporte-col" value="${c.key}" checked>${c.label}</label>`).join('');
  colsDiv.onchange=updateReporteBtnState;
  updateReporteBtnState();
  document.getElementById('reporteModal').classList.add('open');
}
function closeReporteModal(){document.getElementById('reporteModal').classList.remove('open');}
async function generarReporte(){
  const columns=[...document.querySelectorAll('.reporte-col:checked')].map(c=>c.value);
  if(!columns.length||!periodoValido())return;
  const periodo=buildPeriodoPayload();
  const btn=document.getElementById('btnGenerarReporte');
  btn.disabled=true;
  try{
    const resp=await fetch(SERVER+'/reporte',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({columns,periodo})});
    if(!resp.ok){const e=await resp.json();throw new Error(e.error||'Error del servidor');}
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=`reporte-tigo-${Date.now()}.xlsx`;a.click();URL.revokeObjectURL(url);
    closeReporteModal();
  }catch(err){alert('Error: '+err.message);}
  finally{btn.disabled=false;}
}
```

Nota: `getVisibleHistRecords` desaparece (ya no hace falta, no hay checklist de informes). `allRecords` y `norm()` siguen existiendo y usandose en otras partes del archivo (busqueda de Historial) — no tocar esas.

- [ ] **Step 4: Verificar que el servidor sirve el archivo actualizado**

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/tigo
curl -s http://localhost:3000/tigo | grep -c "rep-tab"
curl -s http://localhost:3000/tigo | grep -c "reporteInformes"
```

Expected: primer curl `200`; segundo curl >= `5` (4 tabs + la regla CSS `.rep-tab`); tercer curl `0` (el checklist de informes ya no existe en el HTML).

- [ ] **Step 5: Revision visual manual**

Abrir `http://localhost:3000/tigo`, Historial, click "📊 Reporte": confirmar que el modal muestra dos tarjetas ("Periodo" con tabs Todos/Mes/Semana/Fecha, "Columnas" con checklist vertical), que cambiar de tab muestra/oculta el input correspondiente, que "Generar Excel" esta deshabilitado si la tab activa requiere un input vacio, y que la descarga funciona para "Todos".

- [ ] **Step 6: Commit**

```bash
git add informe_clima_app.html
git commit -m "$(cat <<'EOF'
Rediseñar modal Reporte en Tigo: periodo en vez de informes

Selector de periodo (Todos/Mes/Semana/Fecha) reemplaza el checklist
de informes; columnas en checklist vertical; secciones con estilo
de tarjeta (.card/.card-title) igual al resto del sistema.
EOF
)"
```

---

### Task 3: Frontend WOM — modal con periodo + columnas verticales

**Files:**
- Modify: `informe_wom_app.html:151` (CSS, agregar `.rep-tab`/`.rep-tab.active` despues de `.modal-info`)
- Modify: `informe_wom_app.html:433-444` (markup del modal de Reporte)
- Modify: `informe_wom_app.html:861-903` (JS de Reporte completo)

**Interfaces:**
- Consumes: `POST /reporte-wom` de Task 1 (mismo contrato `{columns, periodo}` que Tigo).
- Produces: nada (ultimo task del plan).

- [ ] **Step 1: Agregar CSS de tabs de periodo**

En `informe_wom_app.html`, reemplazar:

```css
.modal-info{font-size:12px;color:var(--muted);margin-bottom:14px;padding:9px 13px;background:#f8fafc;border-radius:8px;}
```

por:

```css
.modal-info{font-size:12px;color:var(--muted);margin-bottom:14px;padding:9px 13px;background:#f8fafc;border-radius:8px;}
.rep-tab{flex:1;text-align:center;padding:7px 4px;border-radius:8px;font-size:11px;font-weight:600;cursor:pointer;color:var(--muted);}
.rep-tab.active{background:#fff;color:var(--mag);box-shadow:0 1px 3px rgba(0,0,0,.12);}
```

- [ ] **Step 2: Reemplazar el modal de Reporte**

En `informe_wom_app.html`, reemplazar exactamente:

```html
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

por:

```html
<!-- REPORTE MODAL -->
<div class="modal-overlay" id="reporteModal">
  <div class="modal-box" style="max-width:480px;max-height:85vh;overflow-y:auto;">
    <div class="modal-title">📊 Generar reporte Excel<button class="modal-close" onclick="closeReporteModal()">✕</button></div>
    <div class="card" style="margin-bottom:14px;">
      <div class="card-title">Periodo</div>
      <div style="display:flex;gap:4px;background:#EDEEF2;border-radius:10px;padding:4px;margin-bottom:10px;">
        <div class="rep-tab active" data-tipo="todos" onclick="selectPeriodoTipo('todos')">Todos</div>
        <div class="rep-tab" data-tipo="mes" onclick="selectPeriodoTipo('mes')">Mes</div>
        <div class="rep-tab" data-tipo="semana" onclick="selectPeriodoTipo('semana')">Semana</div>
        <div class="rep-tab" data-tipo="fecha" onclick="selectPeriodoTipo('fecha')">Fecha</div>
      </div>
      <div id="periodoInputMes" class="field" style="display:none;margin-bottom:0;"><label>Mes</label><input type="month" id="repMes" oninput="updateReporteBtnState()"></div>
      <div id="periodoInputSemana" class="field" style="display:none;margin-bottom:0;"><label>Cualquier día de la semana</label><input type="date" id="repSemana" oninput="updateReporteBtnState()"></div>
      <div id="periodoInputFecha" style="display:none;">
        <div class="row-2">
          <div class="field" style="margin-bottom:0;"><label>Desde</label><input type="date" id="repDesde" oninput="updateReporteBtnState()"></div>
          <div class="field" style="margin-bottom:0;"><label>Hasta</label><input type="date" id="repHasta" oninput="updateReporteBtnState()"></div>
        </div>
      </div>
    </div>
    <div class="card" style="margin-bottom:14px;">
      <div class="card-title">Columnas</div>
      <div id="reporteColumnas" style="display:flex;flex-direction:column;gap:9px;"></div>
    </div>
    <button class="btn-primary btn" id="btnGenerarReporte" style="width:100%;justify-content:center;" onclick="generarReporte()">📥 Generar Excel</button>
  </div>
</div>
```

Nota: `.field input,.field select,.field textarea` en WOM ya estiliza cualquier `input` dentro de `.field` sin importar el `type`, asi que `repMes` (`type="month"`) y `repSemana` (`type="date"`) quedan bien estilizados sin estilo inline extra (a diferencia de Tigo, que si necesita el inline en `repMes`).

- [ ] **Step 3: Reemplazar el JS de Reporte**

En `informe_wom_app.html`, reemplazar exactamente (todo el bloque entre el comentario `// ── Reporte Excel ──` y el `}` que cierra `generarReporte`, al final del archivo):

```js
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

por:

```js
// ── Reporte Excel ────────────────────────────────────────────────
const REPORTE_COLUMNAS_WOM=[
  {key:'fechaInicio',label:'Fecha Inicio'},{key:'ticket',label:'Ticket'},
  {key:'codInterno',label:'Código Interno'},{key:'instalacion',label:'Instalación'},
  {key:'tipoActividad',label:'Tipo Actividad'},{key:'tecnicos',label:'Técnicos'},
  {key:'photoCount',label:'Cant. Fotos'}
];
let periodoTipoActivo='todos';
function selectPeriodoTipo(tipo){
  periodoTipoActivo=tipo;
  document.querySelectorAll('.rep-tab').forEach(t=>t.classList.toggle('active',t.dataset.tipo===tipo));
  document.getElementById('periodoInputMes').style.display=tipo==='mes'?'block':'none';
  document.getElementById('periodoInputSemana').style.display=tipo==='semana'?'block':'none';
  document.getElementById('periodoInputFecha').style.display=tipo==='fecha'?'block':'none';
  updateReporteBtnState();
}
function buildPeriodoPayload(){
  if(periodoTipoActivo==='mes')return{tipo:'mes',mes:document.getElementById('repMes').value};
  if(periodoTipoActivo==='semana')return{tipo:'semana',fecha:document.getElementById('repSemana').value};
  if(periodoTipoActivo==='fecha')return{tipo:'fecha',desde:document.getElementById('repDesde').value,hasta:document.getElementById('repHasta').value};
  return{tipo:'todos'};
}
function periodoValido(){
  if(periodoTipoActivo==='mes')return!!document.getElementById('repMes').value;
  if(periodoTipoActivo==='semana')return!!document.getElementById('repSemana').value;
  if(periodoTipoActivo==='fecha')return!!document.getElementById('repDesde').value&&!!document.getElementById('repHasta').value;
  return true;
}
function updateReporteBtnState(){
  const anyCol=document.querySelector('.reporte-col:checked');
  document.getElementById('btnGenerarReporte').disabled=!(anyCol&&periodoValido());
}
function openReporteModal(){
  periodoTipoActivo='todos';
  document.querySelectorAll('.rep-tab').forEach(t=>t.classList.toggle('active',t.dataset.tipo==='todos'));
  document.getElementById('periodoInputMes').style.display='none';
  document.getElementById('periodoInputSemana').style.display='none';
  document.getElementById('periodoInputFecha').style.display='none';
  document.getElementById('repMes').value='';
  document.getElementById('repSemana').value='';
  document.getElementById('repDesde').value='';
  document.getElementById('repHasta').value='';
  const colsDiv=document.getElementById('reporteColumnas');
  colsDiv.innerHTML=REPORTE_COLUMNAS_WOM.map(c=>`<label style="display:flex;align-items:center;gap:7px;font-size:13px;"><input type="checkbox" class="reporte-col" value="${c.key}" checked>${c.label}</label>`).join('');
  colsDiv.onchange=updateReporteBtnState;
  updateReporteBtnState();
  document.getElementById('reporteModal').classList.add('open');
}
function closeReporteModal(){document.getElementById('reporteModal').classList.remove('open');}
async function generarReporte(){
  const columns=[...document.querySelectorAll('.reporte-col:checked')].map(c=>c.value);
  if(!columns.length||!periodoValido())return;
  const periodo=buildPeriodoPayload();
  const btn=document.getElementById('btnGenerarReporte');
  btn.disabled=true;
  try{
    const resp=await fetch('/reporte-wom',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({columns,periodo})});
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

Nota: `openReporteModal` deja de ser `async` (ya no hace fetch al abrir) y `reporteWomData` desaparece (no se necesita mas).

- [ ] **Step 4: Verificar que el servidor sirve el archivo actualizado**

```bash
curl -s -o /dev/null -w "%{http_code}\n" http://localhost:3000/wom
curl -s http://localhost:3000/wom | grep -c "rep-tab"
curl -s http://localhost:3000/wom | grep -c "reporteInformes"
```

Expected: primer curl `200`; segundo >= `5`; tercero `0`.

- [ ] **Step 5: Revision visual manual**

Abrir `http://localhost:3000/wom`, Historial, click "📊 Reporte": mismas confirmaciones que Tigo (tarjetas Periodo/Columnas, tabs cambian el input visible, columnas en vertical, boton se deshabilita correctamente, descarga funciona para "Todos").

- [ ] **Step 6: Commit**

```bash
git add informe_wom_app.html
git commit -m "$(cat <<'EOF'
Rediseñar modal Reporte en WOM: periodo en vez de informes

Mismo cambio que Tigo: selector de periodo reemplaza el checklist
de informes, columnas en vertical, tarjetas con acento magenta
propio de WOM.
EOF
)"
```
