const express = require('express');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSXStyle = require('xlsx-js-style');
const nodemailer = require('nodemailer');
const { createClient } = require('@supabase/supabase-js');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
        VerticalAlign, Header } = require('docx');

const supabase = (process.env.SUPABASE_URL && process.env.SUPABASE_KEY)
  ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY)
  : null;
const SUPABASE_BUCKET = process.env.SUPABASE_BUCKET || 'documentos-word';

function sanitizeFnamePart(s, max = 20) {
  return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, max);
}
function initialsFnamePart(name) {
  return (name || '').trim().split(/\s+/).map(w => w[0] || '').join('').toUpperCase();
}

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
app.use(helmet({ contentSecurityPolicy: false })); // CSP off: HTML usa scripts inline

const PORT_FOR_CORS = process.env.PORT || 3000;
const allowedOrigins = [
  `http://localhost:${PORT_FOR_CORS}`,
  `http://127.0.0.1:${PORT_FOR_CORS}`,
  'null', // file:// (HTML abierto directo con doble clic, ver LEEME.txt)
  ...(process.env.ALLOWED_ORIGIN ? [process.env.ALLOWED_ORIGIN] : []),
  // En Render el propio dominio se expone aquí: permite el same-origin del deploy
  ...(process.env.RENDER_EXTERNAL_URL ? [process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '')] : []),
];
const lanOriginPattern = new RegExp(`^http://192\\.168\\.\\d{1,3}\\.\\d{1,3}:${PORT_FOR_CORS}$`);
// Cualquier subdominio https de Render (el sitio llamándose a sí mismo)
const renderOriginPattern = /^https:\/\/[a-z0-9-]+\.onrender\.com$/i;

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)
        || lanOriginPattern.test(origin) || renderOriginPattern.test(origin)) {
      return callback(null, true);
    }
    // Rechazo limpio (sin cabeceras CORS) en vez de lanzar un Error,
    // que Express convertiría en un 500 "Internal Server Error".
    callback(null, false);
  },
}));
const generalLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Demasiadas peticiones, intenta de nuevo en un momento.' },
});
const heavyLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Demasiadas peticiones a este endpoint, intenta de nuevo en un momento.' },
});

app.use(generalLimiter);
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

// ── Row mappers – Clima ────────────────────────────────────────
const fromClima = r => ({
  id: r.id, fecha: r.fecha, fechaCreacion: r.fecha_creacion,
  codInforme: r.cod_informe, nombreSitio: r.nombre_sitio,
  codigoSitio: r.codigo_sitio, tecnico: r.tecnico,
  supervisor: r.supervisor, numOT: r.num_ot,
  lpu: r.lpu, inc: r.inc, equipo: r.equipo, circuito: r.circuito,
  tipoEquipo: r.tipo_equipo, marca: r.marca,
  photoCount: r.photo_count, filename: r.filename
});
const toClima = e => ({
  id: e.id, fecha: e.fecha, fecha_creacion: e.fechaCreacion,
  cod_informe: e.codInforme, nombre_sitio: e.nombreSitio,
  codigo_sitio: e.codigoSitio, tecnico: e.tecnico,
  supervisor: e.supervisor, num_ot: e.numOT,
  lpu: e.lpu, inc: e.inc, equipo: e.equipo, circuito: e.circuito,
  tipo_equipo: e.tipoEquipo, marca: e.marca,
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

// ── Repositorio genérico (Supabase + fallback local JSON) ──────
// Genera toda la capa de acceso a datos de un módulo a partir de su config.
function makeRepo({ table, papelera, dbFile, papeleraFile, from, to,
                    searchFields = [], papSearchFields = [], delCol, delKey }) {
  const loadDB  = () => { try { return JSON.parse(fs.readFileSync(dbFile, 'utf8')); }       catch { return []; } };
  const saveDB  = d  => fs.writeFileSync(dbFile, JSON.stringify(d, null, 2));
  const loadPap = () => { try { return JSON.parse(fs.readFileSync(papeleraFile, 'utf8')); } catch { return []; } };
  const savePap = d  => fs.writeFileSync(papeleraFile, JSON.stringify(d, null, 2));
  const filt = (rows, q, fields) => {
    if (!q) return rows;
    const ql = q.toLowerCase();
    return rows.filter(r => fields.some(k => (r[k] || '').toLowerCase().includes(ql)));
  };
  return {
    async list(q) {
      if (supabase) {
        const { data, error } = await supabase.from(table).select('*').order('fecha_creacion', { ascending: false });
        if (!error) return filt((data || []).map(from), q, searchFields);
        console.error(`${table} list:`, error.message);
      }
      return filt(loadDB(), q, searchFields);
    },
    async insert(entry) {
      if (supabase) {
        const { error } = await supabase.from(table).insert(to(entry));
        if (error) console.error(`${table} insert:`, error.message);
      } else { const db = loadDB(); db.unshift(entry); saveDB(db); }
    },
    async find(id) {
      if (supabase) {
        const { data, error } = await supabase.from(table).select('*').eq('id', id).single();
        if (!error && data) return from(data);
      }
      return loadDB().find(r => r.id === id) || null;
    },
    async remove(id) {
      if (supabase) {
        try {
          const { data, error } = await supabase.from(table).delete().eq('id', id).select();
          if (error) { console.error(`${table} delete:`, error); return { error: `${table}: ${error.message}${error.hint ? ' ('+error.hint+')' : ''}` }; }
          if (!data || !data.length) return { error: `${table}: Supabase no borró ninguna fila (falta política RLS de DELETE o usa la clave service_role en Render).` };
        } catch (e) { console.error(`${table} delete (excepción):`, e); return { error: `${table}: ${e.message || e}` }; }
      } else { saveDB(loadDB().filter(r => r.id !== id)); }
      return { ok: true };
    },
    async papList(q) {
      if (supabase) {
        const { data, error } = await supabase.from(papelera).select('*').order(delCol, { ascending: false });
        if (!error) return filt((data || []).map(r => ({ ...from(r), [delKey]: r[delCol] })), q, papSearchFields);
        console.error(`${papelera} list:`, error.message);
      }
      return filt(loadPap(), q, papSearchFields);
    },
    async papInsert(entry) {
      if (supabase) {
        try {
          const { error } = await supabase.from(papelera).insert({ ...to(entry), [delCol]: entry[delKey] });
          if (error) { console.error(`${papelera} insert:`, error); return { error: `${papelera}: ${error.message}${error.hint ? ' ('+error.hint+')' : ''}` }; }
        } catch (e) { console.error(`${papelera} insert (excepción):`, e); return { error: `${papelera}: ${e.message || e}` }; }
      } else { const p = loadPap(); p.unshift(entry); savePap(p); }
      return { ok: true };
    },
    async papFind(id) {
      if (supabase) {
        const { data, error } = await supabase.from(papelera).select('*').eq('id', id).single();
        if (!error && data) return { ...from(data), [delKey]: data[delCol] };
      }
      return loadPap().find(r => r.id === id) || null;
    },
    async papDelete(id) {
      if (supabase) {
        try {
          const { data, error } = await supabase.from(papelera).delete().eq('id', id).select();
          if (error) { console.error(`${papelera} delete:`, error); return { error: `${papelera}: ${error.message}${error.hint ? ' ('+error.hint+')' : ''}` }; }
          if (!data || !data.length) return { error: `${papelera}: Supabase no borró ninguna fila (falta política RLS de DELETE o usa la clave service_role en Render).` };
        } catch (e) { console.error(`${papelera} delete (excepción):`, e); return { error: `${papelera}: ${e.message || e}` }; }
      } else { savePap(loadPap().filter(r => r.id !== id)); }
      return { ok: true };
    },
    async papClear() {
      if (supabase) {
        const { error } = await supabase.from(papelera).delete().neq('id', '');
        if (error) { console.error(`${papelera} clear:`, error.message); return { error: error.message }; }
      } else { savePap([]); }
      return { ok: true };
    }
  };
}

// ── Informes + Papelera Clima ──────────────────────────────────
const { list: dbClimaList, insert: dbClimaInsert, find: dbClimaFind, remove: dbClimaDelete,
        papList: dbPapeleraList, papInsert: dbPapeleraInsert, papFind: dbPapeleraFind,
        papDelete: dbPapeleraDelete, papClear: dbPapeleraClear } = makeRepo({
  table: 'informes_clima', papelera: 'papelera_clima',
  dbFile: DB_FILE, papeleraFile: PAPELERA_FILE,
  from: fromClima, to: toClima,
  searchFields: ['nombreSitio', 'codInforme', 'tecnico', 'numOT'],
  papSearchFields: ['nombreSitio', 'codInforme', 'tecnico'],
  delCol: 'fecha_eliminado', delKey: 'fechaEliminado'
});

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

  const row_lpu = new TableRow({ height:{value:280}, children:[
    LC('LPU', w_tk_label, 2),
    VC(v(d.lpu), TW-w_tk_label, 12)
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
    HC('Consumo Compresor', mw*3, 4),
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
      row_ig, row_sitio, row_dir, row_tk, row_tk2, row_lpu, row_sala, row_tec,
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

// ── Auto-envío de informes por correo ──────────────────────────
// Configurable por variables de entorno (Render). Si faltan MAIL_USER/MAIL_PASS
// el auto-envío queda desactivado y el sistema sigue funcionando normalmente.
// Outlook personal:  MAIL_HOST=smtp-mail.outlook.com  MAIL_PORT=587
// Microsoft 365:     MAIL_HOST=smtp.office365.com      MAIL_PORT=587
// Gmail (def.):      MAIL_HOST=smtp.gmail.com          MAIL_PORT=587
// Brevo (Sendinblue):MAIL_HOST=smtp-relay.brevo.com    MAIL_PORT=587
//   MAIL_USER = login SMTP de Brevo (p.ej. 8xxxxx@smtp-brevo.com)
//   MAIL_PASS = clave SMTP de Brevo (NO la contraseña de la cuenta)
//   MAIL_FROM = remitente verificado en Brevo (Brevo rechaza enviar desde MAIL_USER)
const MAIL_HOST = process.env.MAIL_HOST || 'smtp.gmail.com';
const MAIL_PORT = parseInt(process.env.MAIL_PORT || '587', 10);
// .trim() defensivo: en Render es fácil dejar un espacio al copiar/pegar y Brevo
// rechaza el envío con "email is not valid". Soporta varios destinos por coma.
const MAIL_TO   = (process.env.MAIL_TO || process.env.MAIL_USER || '').trim();
// Remitente del correo. En Gmail/Outlook coincide con el login; en Brevo debe ser
// una dirección verificada distinta del login SMTP.
const MAIL_FROM = (process.env.MAIL_FROM || process.env.MAIL_USER || '').trim();

// API HTTP de Brevo. Render gratis BLOQUEA los puertos SMTP salientes (25/465/587),
// así que en producción el envío debe ir por la API (HTTPS, puerto 443).
// Si BREVO_API_KEY está definida, se usa la API; si no, se usa SMTP (nodemailer).
const BREVO_API_KEY = process.env.BREVO_API_KEY || '';

let mailTransport = null;
function getMailTransport() {
  if (!process.env.MAIL_USER || !process.env.MAIL_PASS) return null;
  if (!mailTransport) {
    const secure = MAIL_PORT === 465;
    mailTransport = nodemailer.createTransport({
      host: MAIL_HOST, port: MAIL_PORT, secure,
      // En 587 (Outlook/Gmail) forzar STARTTLS: Outlook rechaza la conexión sin esto.
      ...(secure ? {} : { requireTLS: true, tls: { ciphers: 'TLSv1.2' } }),
      // Fallar rápido si el host SMTP no responde (evita colgar el servidor → 502).
      connectionTimeout: 12000, greetingTimeout: 12000, socketTimeout: 15000,
      auth: { user: process.env.MAIL_USER, pass: process.env.MAIL_PASS }
    });
  }
  return mailTransport;
}

// ¿Hay alguna vía de correo configurada?
function mailConfigured() {
  return !!BREVO_API_KEY || (!!process.env.MAIL_USER && !!process.env.MAIL_PASS);
}

// Envío vía API HTTP de Brevo (no usa SMTP, funciona en Render gratis).
async function sendViaBrevoApi({ from, to, subject, text, filename, buffer }) {
  // Acepta uno o varios destinatarios separados por coma (limpia espacios y vacíos).
  const recipients = String(to).split(',').map(e => e.trim()).filter(Boolean).map(email => ({ email }));
  const body = {
    sender: { email: from },
    to: recipients,
    subject,
    textContent: text
  };
  if (buffer && filename) {
    body.attachment = [{ name: filename, content: Buffer.from(buffer).toString('base64') }];
  }
  const resp = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: { 'accept': 'application/json', 'content-type': 'application/json', 'api-key': BREVO_API_KEY },
    body: JSON.stringify(body)
  });
  if (!resp.ok) {
    const detail = await resp.text().catch(() => '');
    throw new Error(`Brevo API ${resp.status}: ${detail}`);
  }
}

// Envío unificado: usa la API de Brevo si hay BREVO_API_KEY, si no, SMTP.
async function sendEmail({ to, subject, text, filename, buffer }) {
  if (!MAIL_FROM) throw new Error('Falta MAIL_FROM (remitente verificado).');
  if (BREVO_API_KEY) {
    return sendViaBrevoApi({ from: MAIL_FROM, to, subject, text, filename, buffer });
  }
  const transport = getMailTransport();
  if (!transport) throw new Error('Sin configuración de correo: define BREVO_API_KEY o MAIL_USER/MAIL_PASS.');
  await transport.sendMail({
    from: MAIL_FROM, to, subject, text,
    attachments: (buffer && filename) ? [{ filename, content: buffer }] : []
  });
}

// Envía el informe recién creado al correo configurado. No bloquea ni lanza:
// si falla, solo registra el error para no afectar la descarga del usuario.
async function autoEnviarInforme({ buffer, filename, subject, text }) {
  if (!mailConfigured()) { console.warn('✉️  Auto-envío desactivado: falta BREVO_API_KEY o MAIL_USER/MAIL_PASS'); return; }
  if (!MAIL_TO)          { console.warn('✉️  Auto-envío desactivado: falta MAIL_TO'); return; }
  try {
    await sendEmail({ to: MAIL_TO, subject, text, filename, buffer });
    console.log(`✉️  Informe enviado automáticamente a ${MAIL_TO}: ${filename}`);
  } catch (err) {
    console.error('✉️  Error en auto-envío de informe:', err.message);
  }
}

// ── Routes ────────────────────────────────────────────────
app.get('/ping', (req,res) => res.json({ok:true}));

// Prueba de configuración de correo: envía un email de test a MAIL_TO.
// Uso: abrir /probar-correo en el navegador. Devuelve el resultado real.
app.get('/probar-correo', async (req, res) => {
  const via = BREVO_API_KEY ? 'API Brevo (HTTPS)' : `SMTP ${MAIL_HOST}:${MAIL_PORT}`;
  if (!mailConfigured()) {
    return res.json({ ok: false, error: 'Correo desactivado: define BREVO_API_KEY (recomendado en Render) o MAIL_USER/MAIL_PASS.' });
  }
  if (!MAIL_FROM) {
    return res.json({ ok: false, error: 'Falta MAIL_FROM (remitente verificado en Brevo).' });
  }
  if (!MAIL_TO) {
    return res.json({ ok: false, error: 'Falta MAIL_TO para saber a quién enviar.' });
  }
  try {
    await sendEmail({
      to: MAIL_TO,
      subject: '✅ Prueba de correo - Informes ICETEL',
      text: `Este es un correo de prueba.\n\nSi lo recibiste, el auto-envío de informes está configurado correctamente.\n\nVía: ${via}\nRemitente: ${MAIL_FROM}\nDestino: ${MAIL_TO}\nFecha: ${new Date().toLocaleString()}`
    });
    res.json({ ok: true, mensaje: `Correo de prueba enviado correctamente a ${MAIL_TO}. Revisa tu bandeja (y spam).`, via });
  } catch (err) {
    res.json({ ok: false, error: err.message, via, sugerencia: BREVO_API_KEY
      ? 'Verifica que BREVO_API_KEY sea válida y que MAIL_FROM sea un remitente verificado en Brevo.'
      : 'Render gratis bloquea SMTP (587): usa BREVO_API_KEY en su lugar. En local, verifica MAIL_USER/MAIL_PASS.' });
  }
});

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

app.post('/generar', heavyLimiter, async (req,res) => {
  try {
    const d = req.body;
    const buffer = await buildDocx(d);
    const fechaPart = (d.fecha || '').split(' - ')[0].trim();
    const fnameParts = [
      sanitizeFnamePart(d.nombreSitio) || 'Sitio',
      initialsFnamePart(d.tecnico) || 'TEC',
      d.equipo ? `E${sanitizeFnamePart(d.equipo, 8)}` : '',
      d.circuito ? `C${sanitizeFnamePart(d.circuito, 8)}` : '',
      sanitizeFnamePart(d.sala, 20),
      sanitizeFnamePart(fechaPart, 10)
    ].filter(Boolean);
    const fname = `${fnameParts.join('_')}.docx`;
    fs.writeFileSync(path.join(DOCS_DIR, fname), buffer);
    await storageUpload(buffer, `clima/${fname}`);

    const { photos } = d;
    const entry = {
      id: Date.now().toString(),
      fecha: d.fecha, fechaCreacion: new Date().toISOString(),
      codInforme: d.codInforme, nombreSitio: d.nombreSitio,
      codigoSitio: d.codigoSitio, tecnico: d.tecnico,
      supervisor: d.supervisor, numOT: d.numOT,
      lpu: d.lpu, inc: d.ticketInc,
      equipo: d.equipo, circuito: d.circuito,
      tipoEquipo: d.eqTipo, marca: d.eqMarca,
      photoCount: (photos||[]).filter(Boolean).length,
      filename: fname
    };
    await dbClimaInsert(entry);

    // Auto-envío por correo (fire-and-forget: no retrasa la descarga)
    autoEnviarInforme({
      buffer, filename: fname,
      subject: `Nuevo informe Clima - ${(d.nombreSitio||'').trim()} - ${(d.codInforme||'').trim()}`.trim(),
      text: `Se generó un nuevo informe correctivo de clima.\n\n` +
            `Sitio: ${d.nombreSitio||'N/A'}\nCódigo informe: ${d.codInforme||'N/A'}\n` +
            `Código sitio: ${d.codigoSitio||'N/A'}\nFecha: ${d.fecha||'N/A'}\n` +
            `Técnico: ${d.tecnico||'N/A'}\nSupervisor: ${d.supervisor||'N/A'}\nN° OT: ${d.numOT||'N/A'}`
    });

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

app.post('/enviar/:id', heavyLimiter, async (req,res) => {
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
 try {
  const entry = await dbClimaFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  // 1) Guardar en la papelera PRIMERO (si falla, no se pierde el informe)
  const ins = await dbPapeleraInsert({ ...entry, fechaEliminado: new Date().toISOString() });
  if (ins && ins.error) return res.status(500).json({ error: 'No se pudo mover a la papelera: ' + ins.error });
  // 2) Borrar del registro principal (si falla, deshacer el paso 1)
  const del = await dbClimaDelete(entry.id);
  if (del && del.error) { await dbPapeleraDelete(entry.id); return res.status(500).json({ error: del.error }); }
  // 3) Mover el archivo (best-effort; no debe romper el borrado)
  try {
    await storageMove(`clima/${entry.filename}`, `clima/papelera/${entry.filename}`);
    const srcPath = path.join(DOCS_DIR, entry.filename);
    if (fs.existsSync(srcPath)) fs.renameSync(srcPath, path.join(PAPELERA_DIR, entry.filename));
  } catch(e) { console.error('mover archivo a papelera:', e.message); }
  res.json({ok:true});
 } catch(e) { console.error('DELETE /registro:', e); res.status(500).json({ error: 'No se pudo borrar: '+(e.message||e) }); }
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
  const del = await dbPapeleraDelete(entry.id);
  if (del && del.error) return res.status(500).json({ error: del.error });
  await storageRemove([`clima/papelera/${entry.filename}`]);
  try {
    const fpath = path.join(PAPELERA_DIR, entry.filename);
    if (fs.existsSync(fpath)) fs.unlinkSync(fpath);
  } catch(e) {}
  res.json({ok:true});
});

// Empty entire papelera
app.delete('/papelera', async (req,res) => {
  const papelera = await dbPapeleraList(null);
  const cleared = await dbPapeleraClear();
  if (cleared && cleared.error) return res.status(500).json({ error: cleared.error });
  if (papelera.length) {
    await storageRemove(papelera.map(e => `clima/papelera/${e.filename}`));
    papelera.forEach(e => {
      try { const f = path.join(PAPELERA_DIR, e.filename); if (fs.existsSync(f)) fs.unlinkSync(f); } catch(e2) {}
    });
  }
  res.json({ok:true});
});

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
  fecha: 'Fecha', codInforme: 'Código Informe', lpu: 'LPU',
  nombreSitio: 'Sitio', equipo: 'Equipo', circuito: 'Circuito',
  tecnico: 'Técnico', tipoEquipo: 'Tipo de Equipo', marca: 'Marca',
  inc: 'Incidencia (Inc)'
};

// ── Excel con estilo de tabla (cabecera color, bordes, zebra, anchos) ──
const XL_AZUL = 'FF0073EA';   // cabecera
const XL_ZEBRA = 'FFEAF3FF';  // filas pares
const XL_BORDE = 'FFB9C4DE';  // líneas de la tabla
function buildReporteExcel(colMap, validColumns, rows, sheetName = 'Historial') {
  const headerRow = validColumns.map(c => colMap[c]);
  const dataRows = rows.map(r => validColumns.map(c => String(r[c] ?? '')));
  const ws = XLSXStyle.utils.aoa_to_sheet([headerRow, ...dataRows]);

  const thin = { style: 'thin', color: { rgb: XL_BORDE } };
  const border = { top: thin, bottom: thin, left: thin, right: thin };
  const ncols = validColumns.length;

  for (let R = 0; R <= dataRows.length; R++) {
    for (let C = 0; C < ncols; C++) {
      const ref = XLSXStyle.utils.encode_cell({ r: R, c: C });
      if (!ws[ref]) ws[ref] = { t: 's', v: '' };
      if (R === 0) {
        ws[ref].s = {
          font: { bold: true, sz: 11, color: { rgb: 'FFFFFFFF' }, name: 'Calibri' },
          fill: { patternType: 'solid', fgColor: { rgb: XL_AZUL } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border
        };
      } else {
        ws[ref].s = {
          font: { sz: 10, color: { rgb: 'FF323338' }, name: 'Calibri' },
          fill: { patternType: 'solid', fgColor: { rgb: R % 2 === 0 ? XL_ZEBRA : 'FFFFFFFF' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: false },
          border
        };
      }
    }
  }

  // Anchos de columna = ajuste al contenido (margen de +2, entre 12 y 45)
  ws['!cols'] = validColumns.map((c, i) => {
    const lens = [colMap[c].length, ...dataRows.map(row => (row[i] || '').length)];
    return { wch: Math.min(Math.max(Math.max(...lens) + 2, 12), 45) };
  });
  ws['!rows'] = [{ hpt: 22 }];  // alto de cabecera
  const lastCol = XLSXStyle.utils.encode_col(ncols - 1);
  ws['!autofilter'] = { ref: `A1:${lastCol}${dataRows.length + 1}` };

  const wb = XLSXStyle.utils.book_new();
  XLSXStyle.utils.book_append_sheet(wb, ws, sheetName);
  return XLSXStyle.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// Handler genérico de reporte (Clima y WOM comparten lógica)
function reporteHandler(colMap, listFn, filename) {
  return async (req, res) => {
    const { columns, periodo } = req.body;
    if (!Array.isArray(columns) || !columns.length) {
      return res.status(400).json({ error: 'Selecciona al menos una columna' });
    }
    const validColumns = columns.filter(c => colMap[c]);
    if (!validColumns.length) return res.status(400).json({ error: 'Columnas inválidas' });
    const rango = computeRangoPeriodo(periodo);
    if (rango.error) return res.status(400).json({ error: rango.error });
    const all = await listFn();
    const selected = rango.inicio
      ? all.filter(r => { const f = new Date(r.fechaCreacion); return f >= rango.inicio && f <= rango.fin; })
      : all;
    const buffer = buildReporteExcel(colMap, validColumns, selected);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  };
}

app.post('/reporte', reporteHandler(REPORTE_COLUMNAS_CLIMA, () => dbClimaList(null), 'reporte-tigo.xlsx'));

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

// ── Informes + Papelera WOM ────────────────────────────────────
const { list: dbWomList, insert: dbWomInsert, find: dbWomFind, remove: dbWomDelete,
        papList: dbPapeleraWomList, papInsert: dbPapeleraWomInsert, papFind: dbPapeleraWomFind,
        papDelete: dbPapeleraWomDelete, papClear: dbPapeleraWomClear } = makeRepo({
  table: 'informes_wom', papelera: 'papelera_wom',
  dbFile: DB_FILE_WOM, papeleraFile: PAPELERA_FILE_WOM,
  from: fromWom, to: toWom,
  delCol: 'deleted_at', delKey: 'deletedAt'
});

const RSO_SITES = {
  'RSO CONCEPCION':   'ANIBAL PINTO 105, CONCEPCION',
  'RSO ANTOFAGASTA':  'FELIX GARCIA 581, ANTOFAGASTA',
  'RSO PUERTO MONTT': 'AV LO CELIS S/N PUERTO MONTT',
  'RSO PUNTA ARENAS': 'AV PDTE EDUARDO FREI MONTALVA 5, PUNTA ARENAS',
  'MSO ROSAS':        'ROSAS 2451, SANTIAGO CENTRO',
  'MSO QUILICURA':    'AV CAÑAVERAL 34, QUILICURA'
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

// ── WOM Routes ─────────────────────────────────────────────────
app.get('/sitios-rso', (_req, res) => res.json(RSO_SITES));
app.get('/actividades-wom', (_req, res) => res.json(ACTIVIDADES_WOM));

app.post('/generar-wom', heavyLimiter, async (req, res) => {
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

    // Auto-envío por correo (fire-and-forget: no retrasa la descarga)
    autoEnviarInforme({
      buffer, filename: fname,
      subject: `Nuevo informe WOM - ${(d.ticket||'').trim()}`.trim(),
      text: `Se generó un nuevo informe WOM.\n\n` +
            `Ticket: ${d.ticket||'N/A'}\nCódigo interno: ${d.codInterno||'N/A'}\n` +
            `Instalación: ${d.instalacion||'N/A'}\nFecha inicio: ${d.fechaInicio||'N/A'}\n` +
            `Tipo actividad: ${d.tipoActividad||'N/A'}\n` +
            `Técnicos: ${(d.tecnicos||[]).filter(Boolean).join(', ')||'N/A'}`
    });

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
 try {
  const entry = await dbWomFind(req.params.id);
  if (!entry) return res.status(404).json({error:'No encontrado'});
  // 1) Guardar en la papelera PRIMERO (si falla, no se pierde el informe)
  const ins = await dbPapeleraWomInsert({ ...entry, deletedAt: new Date().toISOString() });
  if (ins && ins.error) return res.status(500).json({ error: 'No se pudo mover a la papelera: ' + ins.error });
  // 2) Borrar del registro principal (si falla, deshacer el paso 1)
  const del = await dbWomDelete(entry.id);
  if (del && del.error) { await dbPapeleraWomDelete(entry.id); return res.status(500).json({ error: del.error }); }
  // 3) Mover el archivo (best-effort)
  try {
    await storageMove(`wom/${entry.filename}`, `wom/papelera/${entry.filename}`);
    const fp = path.join(DOCS_DIR_WOM, entry.filename);
    if (fs.existsSync(fp)) fs.renameSync(fp, path.join(PAPELERA_DIR_WOM, entry.filename));
  } catch(e) { console.error('mover archivo a papelera (wom):', e.message); }
  res.json({ok:true});
 } catch(e) { console.error('DELETE /registro-wom:', e); res.status(500).json({ error: 'No se pudo borrar: '+(e.message||e) }); }
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
  const del = await dbPapeleraWomDelete(entry.id);
  if (del && del.error) return res.status(500).json({ error: del.error });
  await storageRemove([`wom/papelera/${entry.filename}`]);
  try {
    const fp = path.join(PAPELERA_DIR_WOM, entry.filename);
    if (fs.existsSync(fp)) fs.unlinkSync(fp);
  } catch(e) {}
  res.json({ok:true});
});

app.delete('/papelera-wom', async (_req, res) => {
  const papelera = await dbPapeleraWomList();
  const cleared = await dbPapeleraWomClear();
  if (cleared && cleared.error) return res.status(500).json({ error: cleared.error });
  if (papelera.length) {
    await storageRemove(papelera.map(e => `wom/papelera/${e.filename}`));
    papelera.forEach(e => {
      try { const fp = path.join(PAPELERA_DIR_WOM, e.filename); if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch(e2) {}
    });
  }
  res.json({ok:true});
});

const REPORTE_COLUMNAS_WOM = {
  fechaInicio: 'Fecha Inicio', ticket: 'Ticket', codInterno: 'Código Interno',
  instalacion: 'Instalación', tipoActividad: 'Tipo Actividad',
  tecnicos: 'Técnicos', photoCount: 'Cant. Fotos'
};

app.post('/reporte-wom', reporteHandler(REPORTE_COLUMNAS_WOM, () => dbWomList(), 'reporte-wom.xlsx'));

// ── Auto-purga de papelera: borra definitivamente lo que lleve +30 días ──
const RETENCION_DIAS = Number(process.env.RETENCION_DIAS) || 30;
async function purgarPapelera(nombre, { papList, papDelete, prefix, dir, delKey }) {
  try {
    const items = await papList(null);
    const limite = Date.now() - RETENCION_DIAS * 24 * 60 * 60 * 1000;
    let n = 0;
    for (const e of items) {
      const t = new Date(e[delKey]).getTime();
      if (isNaN(t) || t >= limite) continue;            // sin fecha válida o aún reciente → conservar
      const del = await papDelete(e.id);
      if (del && del.error) { console.error(`purga ${nombre} (${e.filename}):`, del.error); continue; }
      await storageRemove([`${prefix}/papelera/${e.filename}`]);
      try { const fp = path.join(dir, e.filename); if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch {}
      n++;
    }
    if (n) console.log(`🗑️  Papelera ${nombre}: ${n} informe(s) borrados definitivamente (>${RETENCION_DIAS} días).`);
  } catch (e) { console.error(`purgarPapelera ${nombre}:`, e.message || e); }
}
async function purgarTodasLasPapeleras() {
  await purgarPapelera('clima', { papList: dbPapeleraList,    papDelete: dbPapeleraDelete,    prefix: 'clima', dir: PAPELERA_DIR,     delKey: 'fechaEliminado' });
  await purgarPapelera('wom',   { papList: dbPapeleraWomList, papDelete: dbPapeleraWomDelete, prefix: 'wom',   dir: PAPELERA_DIR_WOM, delKey: 'deletedAt' });
}

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
  // Purga de papelera al arrancar y luego cada 24 h
  purgarTodasLasPapeleras();
  setInterval(purgarTodasLasPapeleras, 24 * 60 * 60 * 1000);
});
