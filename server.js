const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { createClient } = require('@supabase/supabase-js');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
        VerticalAlign, Header } = require('docx');

const supabase = (process.env.SUPABASE_URL && process.env.SUPABASE_KEY)
  ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY)
  : null;
const SUPABASE_BUCKET = process.env.SUPABASE_BUCKET || 'documentos-word';

if (supabase) {
  supabase.from('informes_clima').select('id').limit(1)
    .then(({ error }) => {
      if (error) console.error('⚠️  Supabase conectado pero error de acceso:', error.message);
      else console.log('✅ Supabase conectado correctamente');
    });
} else {
  console.warn('⚠️  Supabase NO configurado — falta SUPABASE_URL o SUPABASE_KEY. Usando archivos locales.');
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '80mb' }));
app.use(express.static(__dirname));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'selector.html')));
app.get('/tigo', (req, res) => res.sendFile(path.join(__dirname, 'informe_clima_app.html')));
app.get('/wom', (req, res) => res.sendFile(path.join(__dirname, 'informe_wom_app.html')));

const DOCS_DIR      = path.join(__dirname, 'informes');
const PAPELERA_DIR  = path.join(__dirname, 'papelera');
const DB_FILE       = path.join(__dirname, 'registro.json');
const PAPELERA_FILE = path.join(__dirname, 'papelera.json');

if (!fs.existsSync(DOCS_DIR))     fs.mkdirSync(DOCS_DIR);
if (!fs.existsSync(PAPELERA_DIR)) fs.mkdirSync(PAPELERA_DIR);
if (!fs.existsSync(DB_FILE))      fs.writeFileSync(DB_FILE, '[]');
if (!fs.existsSync(PAPELERA_FILE))fs.writeFileSync(PAPELERA_FILE, '[]');

// ── Local fallback helpers ─────────────────────────────────────
function loadDBLocal()       { try { return JSON.parse(fs.readFileSync(DB_FILE,'utf8')); }       catch { return []; } }
function saveDBLocal(d)      { fs.writeFileSync(DB_FILE, JSON.stringify(d, null, 2)); }
function loadPapeleraLocal() { try { return JSON.parse(fs.readFileSync(PAPELERA_FILE,'utf8')); } catch { return []; } }
function savePapeleraLocal(d){ fs.writeFileSync(PAPELERA_FILE, JSON.stringify(d, null, 2)); }

// ── Row mappers – Clima ────────────────────────────────────────
const fromClima = r => ({
  id: r.id, fecha: r.fecha, fechaCreacion: r.fecha_creacion,
  codInforme: r.cod_informe, nombreSitio: r.nombre_sitio,
  codigoSitio: r.codigo_sitio, tecnico: r.tecnico,
  supervisor: r.supervisor, numOT: r.num_ot,
  photoCount: r.photo_count, filename: r.filename
});
const toClima = e => ({
  id: e.id, fecha: e.fecha, fecha_creacion: e.fechaCreacion,
  cod_informe: e.codInforme, nombre_sitio: e.nombreSitio,
  codigo_sitio: e.codigoSitio, tecnico: e.tecnico,
  supervisor: e.supervisor, num_ot: e.numOT,
  photo_count: e.photoCount, filename: e.filename
});

// ── Supabase Storage helpers ───────────────────────────────────
async function storageUpload(buffer, storagePath) {
  if (!supabase) return;
  const { error } = await supabase.storage.from(SUPABASE_BUCKET)
    .upload(storagePath, buffer, {
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      upsert: true
    });
  if (error) console.error('storageUpload error:', error.message);
}

async function storageDownload(storagePath) {
  if (!supabase) return null;
  const { data, error } = await supabase.storage.from(SUPABASE_BUCKET).download(storagePath);
  if (error || !data) return null;
  return Buffer.from(await data.arrayBuffer());
}

async function storageMove(fromPath, toPath) {
  if (!supabase) return;
  const { error } = await supabase.storage.from(SUPABASE_BUCKET).move(fromPath, toPath);
  if (error) console.error('storageMove error:', error.message);
}

async function storageRemove(paths) {
  if (!supabase) return;
  const { error } = await supabase.storage.from(SUPABASE_BUCKET).remove(paths);
  if (error) console.error('storageRemove error:', error.message);
}

// ── Async DB – Informes Clima ──────────────────────────────────
async function dbClimaList(q) {
  if (supabase) {
    const { data, error } = await supabase.from('informes_clima')
      .select('*').order('fecha_creacion', { ascending: false });
    if (!error) {
      const rows = (data||[]).map(fromClima);
      if (!q) return rows;
      const ql = q.toLowerCase();
      return rows.filter(r => ['nombreSitio','codInforme','tecnico','numOT']
        .some(k => (r[k]||'').toLowerCase().includes(ql)));
    }
    console.error('dbClimaList:', error.message);
  }
  const db = loadDBLocal();
  if (!q) return db;
  const ql = q.toLowerCase();
  return db.filter(r => ['nombreSitio','codInforme','tecnico','numOT']
    .some(k => (r[k]||'').toLowerCase().includes(ql)));
}

async function dbClimaInsert(entry) {
  if (supabase) {
    const { error } = await supabase.from('informes_clima').insert(toClima(entry));
    if (error) console.error('dbClimaInsert:', error.message);
  } else {
    const db = loadDBLocal(); db.unshift(entry); saveDBLocal(db);
  }
}

async function dbClimaFind(id) {
  if (supabase) {
    const { data, error } = await supabase.from('informes_clima')
      .select('*').eq('id', id).single();
    if (!error && data) return fromClima(data);
  }
  return loadDBLocal().find(r => r.id === id) || null;
}

async function dbClimaDelete(id) {
  if (supabase) {
    const { error } = await supabase.from('informes_clima').delete().eq('id', id);
    if (error) console.error('dbClimaDelete:', error.message);
  } else {
    saveDBLocal(loadDBLocal().filter(r => r.id !== id));
  }
}

// ── Async DB – Papelera Clima ──────────────────────────────────
async function dbPapeleraList(q) {
  if (supabase) {
    const { data, error } = await supabase.from('papelera_clima')
      .select('*').order('fecha_eliminado', { ascending: false });
    if (!error) {
      const rows = (data||[]).map(r => ({ ...fromClima(r), fechaEliminado: r.fecha_eliminado }));
      if (!q) return rows;
      const ql = q.toLowerCase();
      return rows.filter(r => ['nombreSitio','codInforme','tecnico']
        .some(k => (r[k]||'').toLowerCase().includes(ql)));
    }
    console.error('dbPapeleraList:', error.message);
  }
  const p = loadPapeleraLocal();
  if (!q) return p;
  const ql = q.toLowerCase();
  return p.filter(r => ['nombreSitio','codInforme','tecnico']
    .some(k => (r[k]||'').toLowerCase().includes(ql)));
}

async function dbPapeleraInsert(entry) {
  if (supabase) {
    const row = { ...toClima(entry), fecha_eliminado: entry.fechaEliminado };
    const { error } = await supabase.from('papelera_clima').insert(row);
    if (error) console.error('dbPapeleraInsert:', error.message);
  } else {
    const p = loadPapeleraLocal(); p.unshift(entry); savePapeleraLocal(p);
  }
}

async function dbPapeleraFind(id) {
  if (supabase) {
    const { data, error } = await supabase.from('papelera_clima')
      .select('*').eq('id', id).single();
    if (!error && data) return { ...fromClima(data), fechaEliminado: data.fecha_eliminado };
  }
  return loadPapeleraLocal().find(r => r.id === id) || null;
}

async function dbPapeleraDelete(id) {
  if (supabase) {
    const { error } = await supabase.from('papelera_clima').delete().eq('id', id);
    if (error) console.error('dbPapeleraDelete:', error.message);
  } else {
    savePapeleraLocal(loadPapeleraLocal().filter(r => r.id !== id));
  }
}

async function dbPapeleraClear() {
  if (supabase) {
    const { error } = await supabase.from('papelera_clima').delete().neq('id', '');
    if (error) console.error('dbPapeleraClear:', error.message);
  } else {
    savePapeleraLocal([]);
  }
}

// Logo base64 embedded
const LOGO_B64 = fs.readFileSync(path.join(__dirname, 'logo.jpeg'), null) || null;

// ── Design tokens (matching original exactly) ─────────────
const BL  = 'DEEAF6';   // azul claro — celdas etiqueta
const BL2 = 'D9E2F3';   // azul medio — sub-headers
const GH  = 'D9D9D9';   // gris — headers sección
const BC  = '7B7B7B';   // color borde
const TW  = 9869;       // ancho tabla principal (igual al header original)

// Borders: thin=4 (like original sz:4), used everywhere
const thinB = (color=BC) => ({ style: BorderStyle.SINGLE, size: 12, color });
const allThin = { top: thinB(), bottom: thinB(), left: thinB(), right: thinB() };

const mkPara = (text, opts={}) => new Paragraph({
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
  spacing: { before: 0, after: 0 },
  children: [new TextRun({
    text: text || '', bold: opts.bold||false, italics: opts.italics||false,
    size: opts.sz || 18, font: 'Calibri', color: opts.color||'000000'
  })]
});

// Label cell (blue bg)
const LC = (text, w, span=1, fill=BL) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  borders: allThin,
  shading: { fill, type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 40, bottom: 40, left: 70, right: 40 },
  children: [mkPara(text, { bold: true, sz: 16 })]
});

// Value cell (white)
const VC = (text, w, span=1) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  borders: allThin,
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 40, bottom: 40, left: 70, right: 40 },
  children: [mkPara(text || '', { sz: 16, center: true })]
});

// Column header cell (blue2 bg, centered, bold)
const HC = (text, w, span=1, rowSpan=1) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  ...(rowSpan > 1 ? { rowSpan } : {}),
  borders: allThin,
  shading: { fill: BL2, type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 30, bottom: 30, left: 30, right: 30 },
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: text||'', bold: true, size: 14, font: 'Calibri' })]
  })]
});

// Section header row (full-width gray, ALL CAPS bold)
const secRow = (text) => new TableRow({
  height: { value: 300 },
  children: [new TableCell({
    width: { size: TW, type: WidthType.DXA }, columnSpan: 14,
    borders: allThin,
    shading: { fill: GH, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 100, right: 40 },
    children: [mkPara(text, { bold: true, sz: 18 })]
  })]
});

// ── Build DOCX ─────────────────────────────────────────────
async function buildDocx(d) {
  const v = s => (s||'').toString().trim() || 'N/A';
  // Separa el resumen por punto seguido de espacio, cada oración en su propia viñeta
  const splitSentences = text => {
    if (!text || text.trim() === 'N/A') return [text || 'N/A'];
    const parts = text.split(/(?<=\.)\s+/).map(s => s.trim()).filter(Boolean);
    return parts.length > 1 ? parts : [text.trim()];
  };
  const mkBullet = text => new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: text || '', size: 16, font: 'Calibri', color: '000000' })]
  });

  // ── HEADER with logo ──────────────────────────────────
  const HDR_BLUE  = '1A3A6C';   // azul corporativo ICETEL
  const HDR_BLUE2 = 'D6E4F0';   // azul claro etiquetas COD/FECHA
  const HDR_BRD   = '1A3A6C';   // borde header
  const thinHdr   = (color=HDR_BRD) => ({ style: BorderStyle.SINGLE, size: 12, color });
  const allHdr    = { top: thinHdr(), bottom: thinHdr(), left: thinHdr(), right: thinHdr() };

  let headerChildren = [];
  try {
    // Intenta logo.png primero, luego logo.jpeg como respaldo
    let logoPath = path.join(__dirname, 'logo.png');
    let logoType = 'png';
    if (!fs.existsSync(logoPath)) { logoPath = path.join(__dirname, 'logo.jpeg'); logoType = 'jpeg'; }
    const logoData = fs.readFileSync(logoPath);
    const headerTable = new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [2563, 4625, 745, 1936],
      borders: {
        top: thinHdr(), bottom: thinHdr(),
        left: thinHdr(), right: thinHdr(),
        insideH: thinHdr(), insideV: thinHdr()
      },
      rows: [
        // Row 1: Logo | CENTRALES CLIMA | COD. | value
        new TableRow({
          height: { value: 450 },
          children: [
            // Logo cell — fondo blanco
            new TableCell({
              width: { size: 2563, type: WidthType.DXA },
              rowSpan: 2,
              borders: allHdr,
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 30, bottom: 30, left: 60, right: 60 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new ImageRun({ data: logoData, transformation: { width: 120, height: 65 }, type: logoType })]
              })]
            }),
            // Title cell — azul corporativo, texto blanco
            new TableCell({
              width: { size: 4625, type: WidthType.DXA },
              borders: allHdr,
              shading: { fill: HDR_BLUE, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 70, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'CENTRALES CLIMA', bold: true, size: 24, font: 'Calibri', color: 'FFFFFF' })]
              })]
            }),
            // COD label
            new TableCell({
              width: { size: 745, type: WidthType.DXA },
              borders: allHdr,
              shading: { fill: HDR_BLUE2, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 40, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'COD.', bold: true, size: 16, font: 'Calibri', color: HDR_BLUE })]
              })]
            }),
            // COD value
            new TableCell({
              width: { size: 1936, type: WidthType.DXA },
              borders: allHdr,
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 40, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: v(d.codInforme), size: 16, font: 'Calibri', bold: true })]
              })]
            }),
          ]
        }),
        // Row 2: (logo merged) | INFORME CORRECTIVO CLIMA | FECHA | value
        new TableRow({
          height: { value: 360 },
          children: [
            new TableCell({
              width: { size: 4625, type: WidthType.DXA },
              borders: allHdr,
              shading: { fill: HDR_BLUE, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 70, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'INFORME CORRECTIVO CLIMA', bold: true, size: 22, font: 'Calibri', color: 'FFFFFF' })]
              })]
            }),
            new TableCell({
              width: { size: 745, type: WidthType.DXA },
              borders: allHdr,
              shading: { fill: HDR_BLUE2, type: ShadingType.CLEAR },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 40, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: 'FECHA', bold: true, size: 16, font: 'Calibri', color: HDR_BLUE })]
              })]
            }),
            new TableCell({
              width: { size: 1936, type: WidthType.DXA },
              borders: allHdr,
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 20, bottom: 20, left: 40, right: 40 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: v(d.fecha), size: 16, font: 'Calibri' })]
              })]
            }),
          ]
        })
      ]
    });
    headerChildren = [headerTable, new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })];
  } catch(e) {
    console.log('Logo not found, skipping:', e.message);
    headerChildren = [new Paragraph({ children: [] })];
  }

  // ── MAIN TABLE ────────────────────────────────────────
  // Col widths matching original exactly (14 cols summing to TW=9869)
  // Proportionally scaled from 11231 → 9869
  const scale = 9869 / 11231;
  const cw = [2827,517,528,782,269,1056,1050,262,787,527,523,790,261,1052];
  const scaled = cw.map(w => Math.round(w * scale));
  // Fix rounding to sum exactly to TW
  const diff = TW - scaled.reduce((a,b)=>a+b,0);
  scaled[0] += diff;

  // ── INFO GENERAL ──────────────────────────────────────
  const row_ig = secRow('INFORMACION GENERAL');
  const row_sitio = new TableRow({ height:{value:280}, children:[
    LC('Nombre de Sitio', scaled[0]+scaled[1], 2),
    VC(v(d.nombreSitio), scaled[2]+scaled[3]+scaled[4]+scaled[5], 4),
    LC('Código de Sitio', scaled[6]+scaled[7]+scaled[8]+scaled[9], 4),
    VC(v(d.codigoSitio), scaled[10]+scaled[11]+scaled[12]+scaled[13], 4)
  ]});
  const row_dir = new TableRow({ height:{value:280}, children:[
    LC('Dirección', scaled[0]+scaled[1], 2),
    VC(v(d.direccion), TW-(scaled[0]+scaled[1]), 12)
  ]});

  const w_tk_label = scaled[0]+scaled[1];
  const w_tk_inc   = scaled[2]+scaled[3];
  const w_tk_te    = scaled[4]+scaled[5];
  const w_tk_ti    = scaled[6]+scaled[7];
  const w_tk_red   = scaled[8]+scaled[9];
  const w_tk_ot    = scaled[10]+scaled[11];
  const w_tk_otv   = scaled[12]+scaled[13];

  const row_tk = new TableRow({ height:{value:280}, children:[
    new TableCell({ width:{size:w_tk_label,type:WidthType.DXA}, columnSpan:2, rowSpan:2,
      borders:allThin, shading:{fill:BL,type:ShadingType.CLEAR},
      verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:70,right:40},
      children:[mkPara('Números de Tickets',{bold:true,sz:16})] }),
    LC('Inc.', w_tk_inc, 2, BL2), LC('TE', w_tk_te, 2, BL2),
    LC('TI', w_tk_ti, 2, BL2), LC('RED', w_tk_red, 2, BL2),
    LC('Numero de OT', w_tk_ot, 2, BL2), VC('', w_tk_otv, 2)
  ]});
  const row_tk2 = new TableRow({ height:{value:260}, children:[
    VC(v(d.ticketInc), w_tk_inc, 2), VC(v(d.ticketTE), w_tk_te, 2),
    VC(v(d.ticketTI), w_tk_ti, 2), VC(v(d.ticketRED), w_tk_red, 2),
    new TableCell({ width:{size:w_tk_ot,type:WidthType.DXA}, columnSpan:2,
      borders:allThin, shading:{fill:BL2,type:ShadingType.CLEAR},
      verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:40,right:40},
      children:[mkPara('')] }),
    VC(v(d.numOT), w_tk_otv, 2)
  ]});

  const row_sala = new TableRow({ height:{value:280}, children:[
    LC('Sala', w_tk_label, 2),
    VC(v(d.sala), w_tk_inc+w_tk_te+w_tk_ti+w_tk_red, 8),
    LC('Fecha Ejecución', w_tk_ot, 2, BL2),
    VC(v(d.fecha), w_tk_otv, 2)
  ]});
  const row_tec = new TableRow({ height:{value:280}, children:[
    LC('Técnico Ejecutante', w_tk_label, 2),
    VC(v(d.tecnico), w_tk_inc+w_tk_te, 4),
    LC('Supervisor', w_tk_ti+w_tk_red, 4),
    VC(v(d.supervisor), w_tk_ot+w_tk_otv, 4)
  ]});

  // ── RESUMEN ───────────────────────────────────────────
  const row_rs = secRow('RESUMEN DE LA ACTIVIDAD');
  const row_rs2 = new TableRow({ height:{value:2200}, children:[
    new TableCell({ width:{size:TW,type:WidthType.DXA}, columnSpan:14,
      borders:allThin, margins:{top:60,bottom:60,left:100,right:100},
      children: splitSentences(v(d.resumen)).map(linea => mkBullet(linea)) })
  ]});

  // ── EQUIPAMIENTO ─────────────────────────────────────
  const row_eq  = secRow('DATOS GENERALES DEL EQUIPAMIENTO');
  const w_eq1 = Math.round(TW*0.23), w_eq2 = Math.round(TW*0.17), w_eq3 = Math.round(TW*0.19),
        w_eq4 = Math.round(TW*0.22), w_eq5 = TW - w_eq1 - w_eq2 - w_eq3 - w_eq4;
  const row_eq_h = new TableRow({ height:{value:280}, children:[
    HC('Sala',w_eq1,1), HC('N° Equipo',w_eq2,4), HC('Tipo',w_eq3,2),
    HC('Marca',w_eq4,4), HC('Modelo / Serie',w_eq5,3)
  ]});
  const row_eq_d = new TableRow({ height:{value:280}, children:[
    VC(v(d.eqSala),w_eq1,1), VC(v(d.eqNumero),w_eq2,4),
    VC(v(d.eqTipo),w_eq3,2), VC(v(d.eqMarca),w_eq4,4), VC(v(d.eqModelo),w_eq5,3)
  ]});

  // ── MEDICIONES ────────────────────────────────────────
  const row_med = secRow('MEDICIONES GENERALES');
  const mw = Math.floor(TW/14);
  const row_med_h1 = new TableRow({ height:{value:260}, children:[
    HC('N° De Equipo', mw*2, 1, 2),
    HC('Consumo Compresor COMP 1', mw*3, 4),
    HC('Consumo Evaporador', mw*2, 2),
    HC('Consumo Condensador', mw*3, 4),
    HC('Temperatura', TW - mw*2 - mw*3 - mw*2 - mw*3, 3)
  ]});
  const mw2 = Math.floor((mw*3)/2);
  const row_med_h2 = new TableRow({ height:{value:260}, children:[
    HC('V.Prom\n(Volt)', mw2, 2), HC('Corriente\n(Amp)', mw*3-mw2, 2),
    HC('V.Prom\n(Volt)', mw, 1), HC('Corriente\n(Amp)', mw, 1),
    HC('V.Prom\n(Volt)', mw2, 2), HC('Corriente\n(Amp)', mw*3-mw2, 2),
    HC('Inyección\n(°C)', mw, 2),
    HC('Retorno\n(°C)', TW - mw*2 - mw*3 - mw*2 - mw*3 - mw, 1)
  ]});
  const row_med_d = new TableRow({ height:{value:300}, children:[
    VC(v(d.eqNumero), mw*2, 1),
    VC(v(d.m_cv), mw2, 2), VC(v(d.m_ca), mw*3-mw2, 2),
    VC(v(d.m_ev), mw, 1), VC(v(d.m_ea), mw, 1),
    VC(v(d.m_condv), mw2, 2), VC(v(d.m_conda), mw*3-mw2, 2),
    VC(v(d.m_tinj), mw, 2),
    VC(v(d.m_tret), TW - mw*2 - mw*3 - mw*2 - mw*3 - mw, 1)
  ]});

  // ── OBSERVACIONES ─────────────────────────────────────
  const row_obs  = secRow('OBSERVACIONES Y RECOMENDACIONES');
  const row_obs2 = new TableRow({ height:{value:1400}, children:[
    new TableCell({ width:{size:TW,type:WidthType.DXA}, columnSpan:14,
      borders:allThin, margins:{top:60,bottom:60,left:100,right:100},
      children:[mkPara(v(d.observaciones),{sz:16})] })
  ]});

  // ── FOTOS — only rows with at least 1 photo ───────────
  const row_foto_hdr = secRow('REGISTRO FOTOGRAFICO');
  const photoRows = [];

  for (let r = 0; r < 6; r++) {
    const i1 = r*2, i2 = r*2+1;
    const p1 = d.photos && d.photos[i1];
    const p2 = d.photos && d.photos[i2];
    if (!p1 && !p2) continue; // skip empty rows entirely

    const cells = [];
    for (const [idx, photoData] of [[i1,p1],[i2,p2]]) {
      let children = [];
      if (photoData) {
        const base64 = photoData.replace(/^data:image\/\w+;base64,/, '');
        const imgBuf = Buffer.from(base64, 'base64');
        const ext = photoData.startsWith('data:image/png') ? 'png' : 'jpeg';
        // Image sized to fill cell nicely
        const descText = (d.photoDescs && d.photoDescs[idx]) ? d.photoDescs[idx] : `Fig. ${idx+1}`;
        children = [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [new ImageRun({ data: imgBuf, transformation: { width: 210, height: 158 }, type: ext })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 0 },
            children: [new TextRun({ text: descText, italics: true, bold: true, size: 14, font: 'Calibri' })]
          })
        ];
      } else {
        // Empty slot — just label, no box/border fill
        children = [
          new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun({text:'',size:14})] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 20, after: 0 },
            children: [new TextRun({ text: `Fig. ${idx+1}`, italics: true, bold: true, size: 14, font: 'Calibri', color: 'CCCCCC' })]
          })
        ];
      }
      cells.push(new TableCell({
        width: { size: Math.floor(TW/2), type: WidthType.DXA },
        columnSpan: 7,
        borders: allThin,
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 40, bottom: 40, left: 60, right: 60 },
        children
      }));
    }
    photoRows.push(new TableRow({ height: { value: 2400 }, children: cells }));
  }

  const mainTable = new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: scaled,
    rows: [
      row_ig, row_sitio, row_dir, row_tk, row_tk2, row_sala, row_tec,
      row_rs, row_rs2,
      row_eq, row_eq_h, row_eq_d,
      row_med, row_med_h1, row_med_h2, row_med_d,
      row_obs, row_obs2,
    ]
  });

  const sectionChildren = [mainTable];

  if (photoRows.length > 0) {
    const photoTable = new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: scaled,
      rows: [row_foto_hdr, ...photoRows]
    });
    sectionChildren.push(
      new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [] }),
      photoTable
    );
  }

  const doc = new Document({
    sections: [{
      headers: { default: new Header({ children: headerChildren }) },
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1701, bottom: 1417, left: 1701, header: 284 }
        }
      },
      children: sectionChildren
    }]
  });

  return Packer.toBuffer(doc);
}

// ── Routes ────────────────────────────────────────────────
app.get('/ping', (req,res) => res.json({ok:true}));

app.get('/ping-supabase', async (req, res) => {
  if (!supabase) return res.json({ ok: false, error: 'SUPABASE_URL o SUPABASE_KEY no configuradas' });
  const { error } = await supabase.from('informes_clima').select('id').limit(1);
  if (error) return res.json({ ok: false, error: error.message });
  res.json({ ok: true, bucket: SUPABASE_BUCKET });
});

app.get('/test-insert', async (req, res) => {
  if (!supabase) return res.json({ ok: false, error: 'Supabase no configurado' });
  const testId = 'test-' + Date.now();
  const { error: insertError } = await supabase.from('informes_clima').insert({
    id: testId, fecha: '2026-01-01', fecha_creacion: new Date().toISOString(),
    cod_informe: 'TEST-001', nombre_sitio: 'Sitio Test', codigo_sitio: 'TST',
    tecnico: 'Test', supervisor: 'Test', num_ot: '000',
    photo_count: 0, filename: 'test.docx'
  });
  if (insertError) return res.json({ ok: false, paso: 'insert', error: insertError.message });
  const { data, error: selectError } = await supabase.from('informes_clima').select('*').eq('id', testId).single();
  if (selectError) return res.json({ ok: false, paso: 'select', error: selectError.message });
  await supabase.from('informes_clima').delete().eq('id', testId);
  res.json({ ok: true, mensaje: 'Insert y select funcionan correctamente', registro: data });
});

app.get('/version', (req,res) => {
  try {
    const mtime = fs.statSync(path.join(__dirname, 'informe_clima_app.html')).mtimeMs;
    res.json({ v: mtime });
  } catch { res.json({ v: 0 }); }
});

app.post('/generar', async (req,res) => {
  try {
    const d = req.body;
    const buffer = await buildDocx(d);
    const sitePart = (d.nombreSitio||'Clima').replace(/[^a-zA-Z0-9]/g,'_').slice(0,25);
    const fname = `${d.codInforme||'Informe'}_${sitePart}.docx`;
    fs.writeFileSync(path.join(DOCS_DIR, fname), buffer);
    await storageUpload(buffer, `clima/${fname}`);

    const { photos } = d;
    const entry = {
      id: Date.now().toString(),
      fecha: d.fecha, fechaCreacion: new Date().toISOString(),
      codInforme: d.codInforme, nombreSitio: d.nombreSitio,
      codigoSitio: d.codigoSitio, tecnico: d.tecnico,
      supervisor: d.supervisor, numOT: d.numOT,
      photoCount: (photos||[]).filter(Boolean).length,
      filename: fname
    };
    await dbClimaInsert(entry);

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro', async (req,res) => {
  const q = (req.query.q||'').toLowerCase().trim();
  res.json(await dbClimaList(q || null));
});

app.get('/descargar/:id', async (req,res) => {
  const entry = await dbClimaFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  let buffer = await storageDownload(`clima/${entry.filename}`);
  if (!buffer) {
    const fpath = path.join(DOCS_DIR, entry.filename);
    if (!fs.existsSync(fpath)) return res.status(404).json({error:'Archivo no existe'});
    buffer = fs.readFileSync(fpath);
  }
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition',`attachment; filename="${entry.filename}"`);
  res.send(buffer);
});

app.post('/enviar/:id', async (req,res) => {
  const entry = await dbClimaFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  let buffer = await storageDownload(`clima/${entry.filename}`);
  if (!buffer) {
    const fpath = path.join(DOCS_DIR, entry.filename);
    if (!fs.existsSync(fpath)) return res.status(404).json({error:'Archivo no existe'});
    buffer = fs.readFileSync(fpath);
  }
  const {to,smtpHost,smtpPort,smtpUser,smtpPass} = req.body;
  if (!to) return res.status(400).json({error:'Email requerido'});
  try {
    const t = nodemailer.createTransport({ host:smtpHost||'smtp.gmail.com', port:smtpPort||587, secure:false, auth:{user:smtpUser,pass:smtpPass} });
    await t.sendMail({ from:smtpUser, to, subject:`Informe - ${entry.nombreSitio} - ${entry.codInforme}`,
      text:`Adjunto informe.\nSitio: ${entry.nombreSitio}\nFecha: ${entry.fecha}\nTécnico: ${entry.tecnico}`,
      attachments:[{filename:entry.filename, content:buffer}] });
    res.json({ok:true});
  } catch(err){ res.status(500).json({error:err.message}); }
});

// Mover a papelera (soft delete)
app.delete('/registro/:id', async (req,res) => {
  const entry = await dbClimaFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageMove(`clima/${entry.filename}`, `clima/papelera/${entry.filename}`);
  try {
    const srcPath = path.join(DOCS_DIR, entry.filename);
    if (fs.existsSync(srcPath)) fs.renameSync(srcPath, path.join(PAPELERA_DIR, entry.filename));
  } catch(e) {}
  await dbClimaDelete(entry.id);
  await dbPapeleraInsert({ ...entry, fechaEliminado: new Date().toISOString() });
  res.json({ok:true});
});

// List papelera
app.get('/papelera', async (req,res) => {
  const q = (req.query.q||'').toLowerCase().trim();
  res.json(await dbPapeleraList(q || null));
});

// Restore from papelera
app.post('/papelera/restaurar/:id', async (req,res) => {
  const entry = await dbPapeleraFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageMove(`clima/papelera/${entry.filename}`, `clima/${entry.filename}`);
  try {
    const srcPath = path.join(PAPELERA_DIR, entry.filename);
    if (fs.existsSync(srcPath)) fs.renameSync(srcPath, path.join(DOCS_DIR, entry.filename));
  } catch(e) {}
  const { fechaEliminado, ...clean } = entry;
  await dbPapeleraDelete(entry.id);
  await dbClimaInsert(clean);
  res.json({ok:true});
});

// Delete permanently from papelera
app.delete('/papelera/:id', async (req,res) => {
  const entry = await dbPapeleraFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageRemove([`clima/papelera/${entry.filename}`]);
  try {
    const fpath = path.join(PAPELERA_DIR, entry.filename);
    if (fs.existsSync(fpath)) fs.unlinkSync(fpath);
  } catch(e) {}
  await dbPapeleraDelete(entry.id);
  res.json({ok:true});
});

// Empty entire papelera
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

// ═══════════════════════════════════════════════════════════════
// MÓDULO WOM
// ═══════════════════════════════════════════════════════════════
const DOCS_DIR_WOM      = path.join(__dirname, 'informes_wom');
const PAPELERA_DIR_WOM  = path.join(__dirname, 'papelera_wom');
const DB_FILE_WOM       = path.join(__dirname, 'registro_wom.json');
const PAPELERA_FILE_WOM = path.join(__dirname, 'papelera_wom.json');

if (!fs.existsSync(DOCS_DIR_WOM))      fs.mkdirSync(DOCS_DIR_WOM);
if (!fs.existsSync(PAPELERA_DIR_WOM))  fs.mkdirSync(PAPELERA_DIR_WOM);
if (!fs.existsSync(DB_FILE_WOM))       fs.writeFileSync(DB_FILE_WOM, '[]');
if (!fs.existsSync(PAPELERA_FILE_WOM)) fs.writeFileSync(PAPELERA_FILE_WOM, '[]');

// ── Local fallback helpers WOM ─────────────────────────────────
function loadDBWomLocal()        { try { return JSON.parse(fs.readFileSync(DB_FILE_WOM,'utf8')); }        catch { return []; } }
function saveDBWomLocal(d)       { fs.writeFileSync(DB_FILE_WOM, JSON.stringify(d, null, 2)); }
function loadPapeleraWomLocal()  { try { return JSON.parse(fs.readFileSync(PAPELERA_FILE_WOM,'utf8')); }  catch { return []; } }
function savePapeleraWomLocal(d) { fs.writeFileSync(PAPELERA_FILE_WOM, JSON.stringify(d, null, 2)); }

// ── Row mappers – WOM ──────────────────────────────────────────
const fromWom = r => ({
  id: r.id, fechaCreacion: r.fecha_creacion,
  ticket: r.ticket, codInterno: r.cod_interno,
  fechaInicio: r.fecha_inicio, instalacion: r.instalacion,
  tipoActividad: r.tipo_actividad, tecnicos: r.tecnicos,
  photoCount: r.photo_count, filename: r.filename
});
const toWom = e => ({
  id: e.id, fecha_creacion: e.fechaCreacion,
  ticket: e.ticket, cod_interno: e.codInterno,
  fecha_inicio: e.fechaInicio, instalacion: e.instalacion,
  tipo_actividad: e.tipoActividad, tecnicos: e.tecnicos,
  photo_count: e.photoCount, filename: e.filename
});

// ── Async DB – Informes WOM ────────────────────────────────────
async function dbWomList() {
  if (supabase) {
    const { data, error } = await supabase.from('informes_wom')
      .select('*').order('fecha_creacion', { ascending: false });
    if (!error) return (data||[]).map(fromWom);
    console.error('dbWomList:', error.message);
  }
  return loadDBWomLocal();
}

async function dbWomInsert(entry) {
  if (supabase) {
    const { error } = await supabase.from('informes_wom').insert(toWom(entry));
    if (error) console.error('dbWomInsert:', error.message);
  } else {
    const db = loadDBWomLocal(); db.unshift(entry); saveDBWomLocal(db);
  }
}

async function dbWomFind(id) {
  if (supabase) {
    const { data, error } = await supabase.from('informes_wom')
      .select('*').eq('id', id).single();
    if (!error && data) return fromWom(data);
  }
  return loadDBWomLocal().find(r => r.id === id) || null;
}

async function dbWomDelete(id) {
  if (supabase) {
    const { error } = await supabase.from('informes_wom').delete().eq('id', id);
    if (error) console.error('dbWomDelete:', error.message);
  } else {
    saveDBWomLocal(loadDBWomLocal().filter(r => r.id !== id));
  }
}

// ── Async DB – Papelera WOM ────────────────────────────────────
async function dbPapeleraWomList() {
  if (supabase) {
    const { data, error } = await supabase.from('papelera_wom')
      .select('*').order('deleted_at', { ascending: false });
    if (!error) return (data||[]).map(r => ({ ...fromWom(r), deletedAt: r.deleted_at }));
    console.error('dbPapeleraWomList:', error.message);
  }
  return loadPapeleraWomLocal();
}

async function dbPapeleraWomInsert(entry) {
  if (supabase) {
    const { error } = await supabase.from('papelera_wom')
      .insert({ ...toWom(entry), deleted_at: entry.deletedAt });
    if (error) console.error('dbPapeleraWomInsert:', error.message);
  } else {
    const p = loadPapeleraWomLocal(); p.unshift(entry); savePapeleraWomLocal(p);
  }
}

async function dbPapeleraWomFind(id) {
  if (supabase) {
    const { data, error } = await supabase.from('papelera_wom')
      .select('*').eq('id', id).single();
    if (!error && data) return { ...fromWom(data), deletedAt: data.deleted_at };
  }
  return loadPapeleraWomLocal().find(r => r.id === id) || null;
}

async function dbPapeleraWomDelete(id) {
  if (supabase) {
    const { error } = await supabase.from('papelera_wom').delete().eq('id', id);
    if (error) console.error('dbPapeleraWomDelete:', error.message);
  } else {
    savePapeleraWomLocal(loadPapeleraWomLocal().filter(r => r.id !== id));
  }
}

async function dbPapeleraWomClear() {
  if (supabase) {
    const { error } = await supabase.from('papelera_wom').delete().neq('id', '');
    if (error) console.error('dbPapeleraWomClear:', error.message);
  } else {
    savePapeleraWomLocal([]);
  }
}

const RSO_SITES = {
  'RSO CONCEPCION':   'ANIBAL PINTO 105, CONCEPCION',
  'RSO ANTOFAGASTA':  'FELIX GARCIA 581, ANTOFAGASTA',
  'RSO PUERTO MONTT': 'AV LO CELIS S/N PUERTO MONTT',
  'RSO PUNTA ARENAS': 'AV PDTE EDUARDO FREI MONTALVA 5, PUNTA ARENAS'
};

const ACTIVIDADES_WOM = [
  'Trabajo Correctivo',
  'Mantenimiento preventivo',
  'Atención de emergencia',
  'Inspección',
  'Retiro de equipos'
];

async function buildDocxWom(d) {
  const v  = s => (s||'').toString().trim();

  // ── Medidas exactas del documento de referencia ──────────
  // Header: tabla exterior invisible, col izq (logos) + col der (ORDEN DE TRABAJO)
  const HDR_L   = 5616;   // col izquierda header (logos)
  const HDR_R   = 5020;   // col derecha header (ORDEN DE TRABAJO = 2×2510)
  const OT_COL  = 2510;   // columnas dentro de la mini-tabla OT

  // Cuerpo: 9 tablas independientes
  const LBL_W   = 1911;   // etiqueta azul (body)
  const VAL_W   = 8453;   // valor (body)
  const FULL_W  = LBL_W + VAL_W;   // 10364
  const TEC_W   = 5067;             // técnicos (media página)
  const DAT_LBL = 1199;
  const DAT_VAL = 9165;
  const RES_L   = 5171;   // resumen col izq
  const RES_R   = 5245;   // resumen col der (exacto del ref: 5171+5245=10416)
  const RES_W   = RES_L + RES_R;   // 10416

  // ── Colores ──────────────────────────────────────────────
  const BLU = '1F497D';
  const WHT = 'FFFFFF';
  const GRN = '008000';
  const BLK = '000000';

  // ── Bordes ───────────────────────────────────────────────
  const thin    = () => ({ style: BorderStyle.SINGLE, size: 4, color: 'auto' });
  const noneB   = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  // Sin borde izquierdo: evita la línea vertical en el margen izquierdo
  const brd     = { top:thin(), bottom:thin(), left:noneB, right:thin() };
  const tblBrd  = { top:thin(), bottom:thin(), left:noneB, right:thin(), insideH:thin(), insideV:thin() };
  const noBrd   = { top:noneB, bottom:noneB, left:noneB, right:noneB };
  const noTblBrd= { top:noneB, bottom:noneB, left:noneB, right:noneB, insideH:noneB, insideV:noneB };

  // ── Helpers ──────────────────────────────────────────────
  const para = (children, align='left', before=0) => new Paragraph({
    alignment: align==='center' ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before, after: 0 },
    children: Array.isArray(children) ? children : [children]
  });
  const run = (text, opts={}) => new TextRun({
    text: text||'', size: opts.sz||18, font: 'Calibri',
    bold: !!opts.bold, italics: !!opts.it, color: opts.c||BLK
  });

  // Celda OT (texto centrado, sin relleno)
  const otCell = (text, w, opts={}) => new TableCell({
    width:{size:w,type:WidthType.DXA}, borders:brd,
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:14,right:14},
    children:[para(run(text,opts),'center',110)]
  });
  // Celda OT azul (ORDEN DE TRABAJO / Fecha OT)
  const otBlu = (text, w, span, sz) => new TableCell({
    width:{size:w,type:WidthType.DXA}, ...(span>1?{columnSpan:span}:{}),
    borders:brd, shading:{fill:BLU,type:ShadingType.CLEAR},
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:60,right:60},
    children:[para(run(text,{bold:true,sz,c:WHT}),'center',sz===30?284:110)]
  });
  // Celda código interno verde centrada
  const otCod = (codVal, w) => new TableCell({
    width:{size:w,type:WidthType.DXA}, borders:brd,
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:14,right:14},
    children:[para(codVal
      ? run(`INC-${codVal}`,{bold:true,c:GRN,sz:18})
      : run('',{sz:18}),
    'center',110)]
  });

  // Mini-tabla ORDEN DE TRABAJO — usa tblBrd (ya sin borde izquierdo)
  const otTblBrd = tblBrd;
  const otTable = new Table({
    width:{size:HDR_R,type:WidthType.DXA}, columnWidths:[OT_COL,OT_COL], borders:otTblBrd,
    indent:{size:0,type:WidthType.DXA},
    rows:[
      new TableRow({height:{value:558},children:[otBlu('ORDEN DE TRABAJO',OT_COL*2,2,30)]}),
      new TableRow({height:{value:404},children:[
        otCell('Código Interno',OT_COL), otCod(v(d.codInterno),OT_COL)
      ]}),
      new TableRow({height:{value:404},children:[
        otCell('Ticket',OT_COL), otCell(v(d.ticket),OT_COL,{bold:true})
      ]}),
      new TableRow({height:{value:404},children:[otBlu('Fecha OT',OT_COL*2,2,18)]}),
      new TableRow({height:{value:393},children:[
        otCell('Inicio:',OT_COL), otCell(`${v(d.fechaInicio)}  ${v(d.horaInicio)}`,OT_COL)
      ]}),
      new TableRow({height:{value:371},children:[
        otCell('Término:',OT_COL), otCell(`${v(d.fechaTermino)}  ${v(d.horaTermino)}`,OT_COL)
      ]})
    ]
  });

  // ── Logos (ICETEL + WOM) ──────────────────────────────────
  let icetelLogo = null, womLogo = null;
  try { const p=path.join(__dirname,'icetel-logo.jpeg'); if(fs.existsSync(p)) icetelLogo=fs.readFileSync(p); } catch(e){}
  try { const p=path.join(__dirname,'wom-logo.png');     if(fs.existsSync(p)) womLogo   =fs.readFileSync(p); } catch(e){}

  const logoParas = [];
  if (icetelLogo) logoParas.push(
    para(new ImageRun({data:icetelLogo, transformation:{width:314,height:175}}), 'left', 0)
  );
  if (womLogo) logoParas.push(
    new Paragraph({
      spacing:{before:40,after:0},
      indent:{left:420},
      children:[new ImageRun({data:womLogo, transformation:{width:220,height:102}})]
    })
  );
  if (!logoParas.length) logoParas.push(para(run('ICETEL / WOM',{bold:true,sz:20}),'left',0));

  // Celda logos (col izq header, sin bordes)
  const logoCell = new TableCell({
    width:{size:HDR_L,type:WidthType.DXA}, borders:noBrd,
    verticalAlign:VerticalAlign.TOP, margins:{top:40,bottom:40,left:0,right:40},
    children:logoParas
  });

  // Celda derecha header: contiene la mini-tabla OT
  const otCell_ = new TableCell({
    width:{size:HDR_R,type:WidthType.DXA}, borders:noBrd,
    verticalAlign:VerticalAlign.TOP, margins:{top:0,bottom:0,left:0,right:0},
    children:[otTable]
  });

  // Tabla exterior header (invisible)
  const hdrTable = new Table({
    width:{size:HDR_L+HDR_R,type:WidthType.DXA},
    columnWidths:[HDR_L,HDR_R],
    borders:noTblBrd,
    rows:[new TableRow({children:[logoCell, otCell_]})]
  });

  // ── Helpers para tablas de cuerpo ────────────────────────
  const bluCell = (text, w, span=1) => new TableCell({
    width:{size:w,type:WidthType.DXA}, ...(span>1?{columnSpan:span}:{}),
    borders:brd, shading:{fill:BLU,type:ShadingType.CLEAR},
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:100,right:80},
    children:[para(run(text,{bold:true,sz:18,c:WHT}))]
  });
  const valCell = (text, w, opts={}) => new TableCell({
    width:{size:w,type:WidthType.DXA}, borders:brd,
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:100,right:60},
    children:[para(run(text,{sz:18,bold:opts.bold||false,c:opts.c||BLK}))]
  });
  const whtCell = (text, w) => new TableCell({
    width:{size:w,type:WidthType.DXA}, borders:brd,
    shading:{fill:WHT,type:ShadingType.CLEAR},
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:100,right:60},
    children:[para(run(text,{bold:true}))]
  });
  const multiCell = (paragraphs, w) => new TableCell({
    width:{size:w,type:WidthType.DXA}, borders:brd,
    margins:{top:60,bottom:60,left:100,right:80}, children:paragraphs
  });

  // ── TABLE 2: Cliente/Sistemas/Tipo actividad/Sitio/Dirección ─
  const bodyTable = new Table({
    width:{size:FULL_W,type:WidthType.DXA}, columnWidths:[LBL_W,VAL_W], borders:tblBrd,
    rows:[
      new TableRow({height:{value:404},children:[bluCell('Cliente',LBL_W),        valCell('WOM',VAL_W,{bold:true})]}),
      new TableRow({height:{value:404},children:[bluCell('Sistemas',LBL_W),        valCell(v(d.infraestructura),VAL_W)]}),
      new TableRow({height:{value:404},children:[bluCell('Tipo de actividad',LBL_W),valCell(v(d.tipoActividad),VAL_W)]}),
      new TableRow({height:{value:404},children:[bluCell('Instalación',LBL_W),     valCell(v(d.instalacion),VAL_W)]}),
      new TableRow({height:{value:404},children:[bluCell('Dirección',LBL_W),       valCell(v(d.direccion),VAL_W)]})
    ]
  });

  // ── TABLE 3: Trabajos Realizados ──────────────────────────
  const trabajosParas = (v(d.trabajos)||'').split('\n').filter(l=>l.trim())
    .map(line => new Paragraph({spacing:{before:0,after:60},children:[run(line.trim())]}));
  if (!trabajosParas.length) trabajosParas.push(new Paragraph({spacing:{before:0,after:0},children:[run('')]}));
  const trabajosTable = new Table({
    width:{size:FULL_W,type:WidthType.DXA}, columnWidths:[FULL_W], borders:tblBrd,
    rows:[
      new TableRow({height:{value:404},children:[bluCell('Trabajos Realizados',FULL_W)]}),
      new TableRow({height:{value:Math.max(608,trabajosParas.length*300)},children:[multiCell(trabajosParas,FULL_W)]})
    ]
  });

  // ── TABLE 4: Observaciones ────────────────────────────────
  const obsTable = new Table({
    width:{size:FULL_W,type:WidthType.DXA}, columnWidths:[FULL_W], borders:tblBrd,
    rows:[
      new TableRow({height:{value:404},children:[bluCell('Observaciones',FULL_W)]}),
      new TableRow({height:{value:404},children:[multiCell(
        [para(run(v(d.observaciones)||'Sin observaciones adicionales'))], FULL_W
      )]})
    ]
  });

  // ── TABLE 5: Separador vacío ──────────────────────────────
  const emptyTable = new Table({
    width:{size:FULL_W,type:WidthType.DXA}, columnWidths:[FULL_W], borders:tblBrd,
    rows:[new TableRow({height:{value:150},children:[
      new TableCell({width:{size:FULL_W,type:WidthType.DXA},borders:brd,children:[para(run(''))]})
    ]})]
  });

  // ── TABLE 6: Técnicos (5067 DXA) ─────────────────────────
  const tecNames = (d.tecnicos||[]).filter(Boolean);
  const tecTable = new Table({
    width:{size:TEC_W,type:WidthType.DXA}, columnWidths:[TEC_W], borders:tblBrd,
    rows:[
      new TableRow({height:{value:404},children:[bluCell('Técnico(s) Responsable(s)',TEC_W)]}),
      new TableRow({height:{value:404},children:[new TableCell({
        width:{size:TEC_W,type:WidthType.DXA}, borders:brd,
        verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:100,right:60},
        children:[para([run('Nombre y Apellido:  ',{bold:true}), run(tecNames.join('    /    '))])]
      })]})
    ]
  });

  // ── TABLE 7: Datos Generales (1199|9165) ─────────────────
  const datTable = new Table({
    width:{size:FULL_W,type:WidthType.DXA}, columnWidths:[DAT_LBL,DAT_VAL], borders:tblBrd,
    rows:[
      new TableRow({height:{value:404},children:[
        new TableCell({width:{size:FULL_W,type:WidthType.DXA},columnSpan:2,borders:brd,
          shading:{fill:BLU,type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
          margins:{top:40,bottom:40,left:100,right:80},
          children:[para(run('Datos generales:',{bold:true,c:WHT}))]})
      ]}),
      new TableRow({height:{value:404},children:[whtCell('Sala:',DAT_LBL),   valCell(v(d.sala),DAT_VAL)]}),
      new TableRow({height:{value:404},children:[whtCell('Equipo:',DAT_LBL), valCell(v(d.equipo),DAT_VAL)]}),
      new TableRow({height:{value:404},children:[whtCell('Marca:',DAT_LBL),  valCell(v(d.marca),DAT_VAL)]}),
      new TableRow({height:{value:404},children:[whtCell('Modelo:',DAT_LBL), valCell(v(d.modelo),DAT_VAL)]})
    ]
  });

  // ── TABLE 8: Resumen + Fotos (5171|5245 = 10416) ─────────
  const photos   = d.photos   || [];
  const captions = d.captions || [];

  const mkPhotoCell = (b64, w) => {
    if (!b64) return new TableCell({width:{size:w,type:WidthType.DXA},borders:brd,
      verticalAlign:VerticalAlign.CENTER,margins:{top:40,bottom:40,left:40,right:40},
      children:[para(run(''))]});
    try {
      const buf = Buffer.from(b64.replace(/^data:image\/\w+;base64,/,''),'base64');
      return new TableCell({width:{size:w,type:WidthType.DXA},borders:brd,
        verticalAlign:VerticalAlign.CENTER,margins:{top:40,bottom:40,left:40,right:40},
        children:[para(new ImageRun({data:buf,transformation:{width:235,height:175}}),'center')]});
    } catch(e) {
      return new TableCell({width:{size:w,type:WidthType.DXA},borders:brd,
        children:[para(run('[error foto]',{sz:14}))]});
    }
  };
  const mkCapCell = (idx, w) => new TableCell({
    width:{size:w,type:WidthType.DXA},borders:brd,
    verticalAlign:VerticalAlign.CENTER,margins:{top:30,bottom:30,left:100,right:60},
    children:[para(run(captions[idx]||'',{sz:16,it:true}))]
  });

  const resRows = [
    new TableRow({height:{value:394},children:[
      new TableCell({width:{size:RES_W,type:WidthType.DXA},columnSpan:2,borders:brd,
        shading:{fill:BLU,type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,
        margins:{top:40,bottom:40,left:100,right:80},
        children:[para(run('RESUMEN DE ACTIVIDAD:',{bold:true,c:WHT}))]})
    ]}),
    new TableRow({height:{value:414},children:[
      new TableCell({width:{size:RES_L,type:WidthType.DXA},borders:brd,
        margins:{top:40,bottom:40,left:100,right:60},children:[para(run(v(d.resumen1)))]}),
      new TableCell({width:{size:RES_R,type:WidthType.DXA},borders:brd,
        margins:{top:40,bottom:40,left:100,right:60},children:[para(run(v(d.resumen2)))]})
    ]})
  ];

  for (let i=0; i<Math.min(photos.length,8); i+=2) {
    const ph = i===0 ? 3231 : 3826;
    resRows.push(new TableRow({height:{value:ph},children:[
      mkPhotoCell(photos[i]||null,RES_L), mkPhotoCell(photos[i+1]||null,RES_R)
    ]}));
    resRows.push(new TableRow({height:{value:458},children:[
      mkCapCell(i,RES_L), mkCapCell(i+1,RES_R)
    ]}));
  }

  const resTable = new Table({
    width:{size:RES_W,type:WidthType.DXA}, columnWidths:[RES_L,RES_R], borders:tblBrd, rows:resRows
  });

  // ── Documento final ───────────────────────────────────────
  const gap = new Paragraph({spacing:{before:0,after:60},children:[]});
  const doc = new Document({
    sections:[{
      properties:{
        page:{
          margin:{top:720,right:708,bottom:280,left:566},
          size:{width:11910,height:16840}
        }
      },
      children:[
        hdrTable, gap,
        bodyTable, gap,
        trabajosTable, gap,
        obsTable, gap,
        emptyTable, gap,
        tecTable, gap,
        datTable, gap,
        resTable
      ]
    }]
  });
  return await Packer.toBuffer(doc);
}

// ── (old buildDocxWom removed — replaced above) ──────────────
async function buildDocxWom_UNUSED(d) {
  const v  = s => (s||'').toString().trim();
  const WTW   = 9869;
  const wCol  = Math.floor(WTW / 10);
  const W_BC  = '7B7B7B';
  const W_LBL = 'D6E4F0';
  const W_GRY = 'D9D9D9';
  const W_BLU = '1A3A6C';
  const W_GRN = '2E7D32';
  const W_MAG = 'E2007A';

  const thinW  = (c=W_BC) => ({ style: BorderStyle.SINGLE, size: 12, color: c });
  const allW   = { top: thinW(), bottom: thinW(), left: thinW(), right: thinW() };
  const thinBl = (c=W_BLU) => ({ style: BorderStyle.SINGLE, size: 12, color: c });
  const allBl  = { top: thinBl(), bottom: thinBl(), left: thinBl(), right: thinBl() };

  const para = (runs, center=false) => new Paragraph({
    alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: Array.isArray(runs) ? runs : [runs]
  });
  const run  = (text, opts={}) => new TextRun({ text: text||'', size: opts.sz||16, font: 'Calibri',
    bold: opts.bold||false, italics: opts.italics||false, color: opts.color||'000000' });

  const WL = (text, span=1) => new TableCell({
    width: { size: wCol*span, type: WidthType.DXA }, ...(span>1?{columnSpan:span}:{}),
    borders: allW, shading: { fill: W_LBL, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top:40, bottom:40, left:70, right:40 },
    children: [para(run(text, { bold:true, sz:16 }))]
  });
  const WV = (text, span=1, opts={}) => new TableCell({
    width: { size: wCol*span, type: WidthType.DXA }, ...(span>1?{columnSpan:span}:{}),
    borders: allW, verticalAlign: VerticalAlign.CENTER,
    margins: { top:40, bottom:40, left:70, right:40 },
    children: [para(run(text, { sz:16, bold:opts.bold||false, color:opts.color||'000000' }), opts.center||false)]
  });
  const wSecRow = (text) => new TableRow({ height:{value:300}, children:[new TableCell({
    width:{size:WTW,type:WidthType.DXA}, columnSpan:10,
    borders:allW, shading:{fill:W_GRY,type:ShadingType.CLEAR},
    verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:100,right:40},
    children:[para(run(text,{bold:true,sz:18}))]
  })]});

  // ── Header con logo ──────────────────────────────────────
  let headerTable = null;
  try {
    let logoPath = path.join(__dirname, 'logo.png');
    if (!fs.existsSync(logoPath)) logoPath = path.join(__dirname, 'logo.jpeg');
    const logoData = fs.readFileSync(logoPath);
    headerTable = new Table({
      width: { size: WTW, type: WidthType.DXA },
      columnWidths: [2400, 5069, 2400],
      borders: { top:thinBl(),bottom:thinBl(),left:thinBl(),right:thinBl(),insideH:thinBl(),insideV:thinBl() },
      rows: [new TableRow({ height:{value:820}, children:[
        new TableCell({
          width:{size:2400,type:WidthType.DXA}, borders:allBl,
          shading:{fill:W_BLU,type:ShadingType.CLEAR},
          verticalAlign:VerticalAlign.CENTER, margins:{top:60,bottom:60,left:100,right:100},
          children:[para(new ImageRun({ data:logoData, transformation:{width:100,height:44} }), true)]
        }),
        new TableCell({
          width:{size:5069,type:WidthType.DXA}, borders:allBl,
          shading:{fill:W_BLU,type:ShadingType.CLEAR},
          verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:60,right:60},
          children:[para(run('ORDEN DE TRABAJO',{bold:true,sz:28,color:'FFFFFF'}), true)]
        }),
        new TableCell({
          width:{size:2400,type:WidthType.DXA}, borders:allBl,
          shading:{fill:W_MAG,type:ShadingType.CLEAR},
          verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:60,right:60},
          children:[para(run('WOM',{bold:true,sz:36,color:'FFFFFF',italics:true}), true)]
        })
      ]})]
    });
  } catch(e) { /* logo opcional */ }

  // ── Código Interno (verde si hay valor) ─────────────────
  const codVal = v(d.codInterno);
  const codCell = new TableCell({
    width:{size:wCol*3,type:WidthType.DXA}, columnSpan:3,
    borders:allW, verticalAlign:VerticalAlign.CENTER,
    margins:{top:40,bottom:40,left:70,right:40},
    children:[para(codVal ? run(`INC-${codVal}`,{bold:true,sz:18,color:W_GRN}) : run('',{sz:16}))]
  });

  // ── Trabajos: split por saltos de línea ─────────────────
  const trabajosParas = (v(d.trabajos)||'').split('\n').filter(l=>l.trim()).map(line =>
    new Paragraph({ spacing:{before:0,after:60}, children:[run(line.trim(),{sz:16})] })
  );
  if (!trabajosParas.length) trabajosParas.push(new Paragraph({spacing:{before:0,after:0},children:[run('',{sz:16})]}));

  // ── Técnicos ────────────────────────────────────────────
  const tecnicosStr = (d.tecnicos||[]).filter(Boolean).join('    /    ');

  // ── Fotos ────────────────────────────────────────────────
  const photos   = d.photos   || [];
  const captions = d.captions || [];
  const photoRows = [];
  for (let r = 0; r < 4; r++) {
    const i1 = r*2, i2 = r*2+1;
    const p1 = photos[i1], p2 = photos[i2];
    if (!p1 && !p2) continue;
    const mkPhotoCell = (b64) => {
      const halfW = Math.floor(WTW/2);
      if (!b64) return new TableCell({
        width:{size:halfW,type:WidthType.DXA},columnSpan:5,
        borders:allW,verticalAlign:VerticalAlign.CENTER,
        margins:{top:40,bottom:40,left:40,right:40},
        children:[para(run('',{sz:16}))]
      });
      try {
        const buf = Buffer.from(b64.replace(/^data:image\/\w+;base64,/,''),'base64');
        return new TableCell({
          width:{size:halfW,type:WidthType.DXA},columnSpan:5,
          borders:allW,verticalAlign:VerticalAlign.CENTER,
          margins:{top:40,bottom:40,left:40,right:40},
          children:[para(new ImageRun({data:buf,transformation:{width:250,height:185}}),true)]
        });
      } catch { return new TableCell({
        width:{size:Math.floor(WTW/2),type:WidthType.DXA},columnSpan:5,
        borders:allW,children:[para(run('[error foto]',{sz:14}))]
      }); }
    };
    const mkCapCell = (idx) => new TableCell({
      width:{size:Math.floor(WTW/2),type:WidthType.DXA},columnSpan:5,
      borders:allW,verticalAlign:VerticalAlign.CENTER,
      margins:{top:30,bottom:30,left:70,right:40},
      children:[para(run(captions[idx]||'',{sz:14,italics:true}),true)]
    });
    photoRows.push(new TableRow({height:{value:2200},children:[mkPhotoCell(p1),mkPhotoCell(p2)]}));
    photoRows.push(new TableRow({height:{value:270},children:[mkCapCell(i1),mkCapCell(i2)]}));
  }

  // ── Tabla principal ──────────────────────────────────────
  const mainTable = new Table({
    width:{size:WTW,type:WidthType.DXA},
    columnWidths: Array(10).fill(wCol),
    borders:{top:thinW(),bottom:thinW(),left:thinW(),right:thinW(),insideH:thinW(),insideV:thinW()},
    rows:[
      new TableRow({height:{value:320},children:[WL('Código Interno',2),codCell,WL('Ticket',2),WV(v(d.ticket),3)]}),
      new TableRow({height:{value:300},children:[WL('Fecha OT',2),WV(`Inicio:   ${v(d.fechaInicio)}   ${v(d.horaInicio)}`,4),WV(`Término:   ${v(d.fechaTermino)}   ${v(d.horaTermino)}`,4)]}),
      new TableRow({height:{value:300},children:[WL('Cliente',2),WV('WOM',8,{bold:true})]}),
      new TableRow({height:{value:300},children:[WL('Infraestructura',2),WV(v(d.infraestructura),8)]}),
      new TableRow({height:{value:300},children:[WL('Tipo de actividad',2),WV(v(d.tipoActividad),8)]}),
      new TableRow({height:{value:300},children:[WL('Instalación',2),WV(v(d.instalacion),3),WL('Dirección',2),WV(v(d.direccion),3)]}),
      wSecRow('Trabajos Realizados'),
      new TableRow({height:{value:1600},children:[new TableCell({
        width:{size:WTW,type:WidthType.DXA},columnSpan:10,
        borders:allW,margins:{top:60,bottom:60,left:100,right:100},
        children:trabajosParas
      })]}),
      wSecRow('Observaciones'),
      new TableRow({height:{value:1200},children:[new TableCell({
        width:{size:WTW,type:WidthType.DXA},columnSpan:10,
        borders:allW,margins:{top:60,bottom:60,left:100,right:100},
        children:[para(run(v(d.observaciones)||'Sin observaciones adicionales',{sz:16}))]
      })]}),
      wSecRow('Técnico(s) Responsable(s)'),
      new TableRow({height:{value:350},children:[WL('Nombre y Apellido:',2),WV(tecnicosStr,8)]}),
      new TableRow({height:{value:300},children:[WL('Sala:',1),WV(v(d.sala),2),WL('Equipo:',1),WV(v(d.equipo),2),WL('Marca:',1),WV(v(d.marca),2),WL('Modelo:',1),WV(v(d.modelo),1)]}),
      wSecRow('RESUMEN DE ACTIVIDAD'),
      new TableRow({height:{value:1400},children:[
        new TableCell({width:{size:Math.floor(WTW/2),type:WidthType.DXA},columnSpan:5,borders:allW,margins:{top:60,bottom:60,left:100,right:100},children:[para(run(v(d.resumen1),{sz:16}))]}),
        new TableCell({width:{size:Math.floor(WTW/2),type:WidthType.DXA},columnSpan:5,borders:allW,margins:{top:60,bottom:60,left:100,right:100},children:[para(run(v(d.resumen2),{sz:16}))]})
      ]}),
      ...(photoRows.length ? [wSecRow('REGISTRO FOTOGRÁFICO'),...photoRows] : [])
    ]
  });

  const sections = [];
  if (headerTable) sections.push(headerTable);
  sections.push(new Paragraph({spacing:{before:0,after:80},children:[]}));
  sections.push(mainTable);

  const doc = new Document({
    sections:[{
      properties:{ page:{ margin:{top:720,right:720,bottom:720,left:720}, size:{width:12240,height:15840} } },
      children: sections
    }]
  });
  return await Packer.toBuffer(doc);
}

// ── WOM Routes ─────────────────────────────────────────────────
app.get('/sitios-rso', (_req, res) => res.json(RSO_SITES));
app.get('/actividades-wom', (_req, res) => res.json(ACTIVIDADES_WOM));

app.post('/generar-wom', async (req, res) => {
  try {
    const d = req.body;
    const buffer = await buildDocxWom(d);
    const ticket = (d.ticket||'WOM').replace(/[^a-zA-Z0-9\-_]/g,'_').slice(0,50);
    const fname  = `${ticket}_WOM.docx`;
    fs.writeFileSync(path.join(DOCS_DIR_WOM, fname), buffer);
    await storageUpload(buffer, `wom/${fname}`);

    const { photos, captions } = d;
    const entry = {
      id: Date.now().toString(),
      fechaCreacion: new Date().toISOString(),
      ticket: d.ticket, codInterno: d.codInterno,
      fechaInicio: d.fechaInicio, instalacion: d.instalacion,
      tipoActividad: d.tipoActividad,
      tecnicos: (d.tecnicos||[]).filter(Boolean).join(', '),
      photoCount: (photos||[]).filter(Boolean).length,
      filename: fname
    };
    await dbWomInsert(entry);

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro-wom', async (_req, res) => res.json(await dbWomList()));

app.get('/descargar-wom/:id', async (req, res) => {
  const entry = await dbWomFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  let buffer = await storageDownload(`wom/${entry.filename}`);
  if (!buffer) {
    const fpath = path.join(DOCS_DIR_WOM, entry.filename);
    if (!fs.existsSync(fpath)) return res.status(404).json({error:'Archivo no encontrado'});
    buffer = fs.readFileSync(fpath);
  }
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition',`attachment; filename="${entry.filename}"`);
  res.send(buffer);
});

app.delete('/registro-wom/:id', async (req, res) => {
  const entry = await dbWomFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageMove(`wom/${entry.filename}`, `wom/papelera/${entry.filename}`);
  try {
    const fp = path.join(DOCS_DIR_WOM, entry.filename);
    if (fs.existsSync(fp)) fs.renameSync(fp, path.join(PAPELERA_DIR_WOM, entry.filename));
  } catch(e) {}
  await dbWomDelete(entry.id);
  await dbPapeleraWomInsert({ ...entry, deletedAt: new Date().toISOString() });
  res.json({ok:true});
});

app.get('/papelera-wom', async (_req, res) => res.json(await dbPapeleraWomList()));

app.post('/papelera-wom/restaurar/:id', async (req, res) => {
  const entry = await dbPapeleraWomFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageMove(`wom/papelera/${entry.filename}`, `wom/${entry.filename}`);
  try {
    const fp = path.join(PAPELERA_DIR_WOM, entry.filename);
    if (fs.existsSync(fp)) fs.renameSync(fp, path.join(DOCS_DIR_WOM, entry.filename));
  } catch(e) {}
  const { deletedAt, ...clean } = entry;
  await dbPapeleraWomDelete(entry.id);
  await dbWomInsert(clean);
  res.json({ok:true});
});

app.delete('/papelera-wom/:id', async (req, res) => {
  const entry = await dbPapeleraWomFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  await storageRemove([`wom/papelera/${entry.filename}`]);
  try {
    const fp = path.join(PAPELERA_DIR_WOM, entry.filename);
    if (fs.existsSync(fp)) fs.unlinkSync(fp);
  } catch(e) {}
  await dbPapeleraWomDelete(entry.id);
  res.json({ok:true});
});

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

const PORT = process.env.PORT || 3000;
const os = require('os');
function getLocalIP() {
  for (const ifaces of Object.values(os.networkInterfaces())) {
    for (const iface of ifaces) {
      if (iface.family === 'IPv4' && !iface.internal) return iface.address;
    }
  }
  return 'localhost';
}
app.listen(PORT, '0.0.0.0', () => {
  const ip = getLocalIP();
  console.log(`✅ Servidor corriendo:`);
  console.log(`   PC:     http://localhost:${PORT}`);
  console.log(`   Celular (misma red WiFi): http://${ip}:${PORT}`);
  console.log(`   Informes guardados en: ./informes/`);
});
