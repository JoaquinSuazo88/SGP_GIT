/**
 * _gen_html.js — Generador + Watcher del visor HTML de documentación SGP
 *
 * Uso:
 *   node _gen_html.js            → genera visor_documentos.html una vez
 *   node _gen_html.js --watch    → genera y vigila cambios en los .md
 */

const fs   = require('fs');
const path = require('path');

const BASE    = __dirname;
const OUTPUT  = path.join(BASE, 'visor_documentos.html');
const WATCH   = process.argv.includes('--watch');
const IGNORED = new Set(['node_modules', '.git']);   // dirs a ignorar

/* ─── Metadatos por palabras clave en nombre de directorio ── */
const META_RULES = [
  { keys: ['produccion','producción'],  icon:'🏭', color:'#1e3a5f', bg:'rgba(30,58,95,.15)' },
  { keys: ['minuta','receta'],          icon:'📋', color:'#c05c00', bg:'rgba(232,93,4,.14)' },
  { keys: ['informe','reporte'],        icon:'📊', color:'#185c37', bg:'rgba(24,92,55,.14)' },
  { keys: ['kpi','indicador'],          icon:'📈', color:'#6b21a8', bg:'rgba(107,33,168,.14)' },
  { keys: ['pedido','compra'],          icon:'🛒', color:'#9f1239', bg:'rgba(159,18,57,.14)' },
  { keys: ['inventario','stock','bod'], icon:'📦', color:'#0e5f8a', bg:'rgba(14,95,138,.14)' },
  { keys: ['venta','cafeteria'],        icon:'💰', color:'#7c6300', bg:'rgba(180,130,0,.14)' },
  { keys: ['usuario','manual'],         icon:'👤', color:'#374151', bg:'rgba(55,65,81,.13)' },
  { keys: ['integracion','sap','flms'], icon:'🔗', color:'#1a5276', bg:'rgba(26,82,118,.14)' },
  { keys: ['cierre'],                   icon:'🔒', color:'#6b2020', bg:'rgba(150,30,30,.13)' },
  { keys: ['planif'],                   icon:'📅', color:'#155724', bg:'rgba(21,87,36,.13)' },
];
const DEFAULT_META = { icon:'📄', color:'#374151', bg:'rgba(55,65,81,.12)' };

function getMetaForDir(dirName) {
  const lower = dirName.toLowerCase().replace(/[-_]/g, '');
  for (const rule of META_RULES) {
    if (rule.keys.some(k => lower.includes(k.replace(/[-_]/g, '')))) return rule;
  }
  return DEFAULT_META;
}

/* ─── Extraer título del primer H1 del .md ─────────────────── */
// Palabras genéricas que no sirven como título de módulo
const GENERIC_TITLES = ['levantamiento','documento','documentacion','descripcion','introduccion','resumen','indice','contenido','drf'];
function extractTitle(mdText, dirName) {
  const m = mdText.match(/^#\s+(.+)/m);
  if (m) {
    const t = m[1].replace(/\*\*/g,'').trim();
    const lower = t.toLowerCase().replace(/[^a-z]/g,'');
    if (!GENERIC_TITLES.some(g => lower.startsWith(g) && t.length < 40)) return t;
  }
  // Fallback: limpiar el nombre del directorio (quitar prefijos DRF_, guiones)
  return dirName.replace(/^DRF[_-]?/i,'').replace(/[-_]/g,' ').replace(/\s+/g,' ').trim();
}

/* ─── Extraer subtítulo del primer H2 ──────────────────────── */
function extractSub(mdText, fallback) {
  const m = mdText.match(/^##\s+(.+)/m);
  return m ? m[1].replace(/\*\*/g, '').trim() : fallback;
}

/* ─── Descubrir módulos automáticamente ─────────────────────── */
function discoverMods() {
  const entries = fs.readdirSync(BASE, { withFileTypes: true });
  const mods = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (IGNORED.has(entry.name) || entry.name.startsWith('.') || entry.name.startsWith('_')) continue;

    const dirPath = path.join(BASE, entry.name);
    const mdFiles = fs.readdirSync(dirPath).filter(f => f.toLowerCase().endsWith('.md'));
    if (!mdFiles.length) continue;

    const mdFile = mdFiles[0];   // primer .md encontrado
    const mdPath = path.join(dirPath, mdFile);
    const mdText = fs.readFileSync(mdPath, 'utf8');

    const id    = entry.name.toLowerCase().replace(/[^a-z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
    const meta  = getMetaForDir(entry.name);
    const title = extractTitle(mdText, entry.name);
    const sub   = extractSub(mdText, entry.name);

    mods.push({
      id,
      title,
      sub:   sub === title ? entry.name : sub,
      file:  `${entry.name}/${mdFile}`,
      dir:   entry.name,
      icon:  meta.icon,
      color: meta.color,
      bg:    meta.bg,
    });
  }

  // Ordenar por nombre de directorio para consistencia
  mods.sort((a, b) => a.dir.localeCompare(b.dir));
  return mods;
}

/* ─── Leer y preparar contenido de cada módulo ─────────────── */
function loadDocs(mods) {
  const docs = {};
  for (const m of mods) {
    const fp  = path.join(BASE, m.file);
    let   txt = fs.readFileSync(fp, 'utf8');
    // Corregir rutas de imágenes relativas
    txt = txt.replace(/!\[([^\]]*)\]\(imagenes\//g, `![$1](./${m.dir}/imagenes/`);
    docs[m.id] = txt;
  }
  return docs;
}

/* ─── Generar el HTML ───────────────────────────────────────── */
function generate() {
  const ts   = new Date().toLocaleString('es-CL');
  console.log(`[${ts}] Escaneando módulos…`);

  let mods, docs;
  try {
    mods = discoverMods();
    docs = loadDocs(mods);
  } catch (err) {
    console.error('  ✗ Error al leer archivos:', err.message);
    return;
  }

  mods.forEach(m => {
    const kb = (Buffer.byteLength(docs[m.id], 'utf8') / 1024).toFixed(0);
    console.log(`  ✓ ${m.title} (${kb} KB) ← ${m.file}`);
  });

  const modsJson = JSON.stringify(mods);
  const docsJson = JSON.stringify(docs);

  /* ── HTML template ──────────────────────────────────────── */
  const html = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SGP Upgrade – Documentación Funcional</title>
  <!-- Generado automáticamente el ${ts} -->
  <script src="https://cdn.jsdelivr.net/npm/marked@4/marked.min.js"><\/script>
  <style>
    :root {
      --primary:#1e3a5f; --primary-l:#2d5a8e; --accent:#e85d04;
      --sbw:268px; --tocw:248px; --hh:58px;
      --bg:#f0f4f8; --card:#fff; --text:#1a202c; --muted:#64748b; --border:#e2e8f0;
    }
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--text);height:100vh;overflow:hidden;display:flex;flex-direction:column}

    /* HEADER */
    .hdr{position:fixed;top:0;left:0;right:0;height:var(--hh);background:var(--primary);display:flex;align-items:center;gap:18px;padding:0 20px;z-index:200;box-shadow:0 2px 12px rgba(0,0,0,.25)}
    .hdr-brand{display:flex;flex-direction:column;gap:1px;cursor:pointer}
    .hdr-brand strong{font-size:17px;font-weight:800;color:#fff}
    .hdr-brand strong span{color:var(--accent)}
    .hdr-brand small{font-size:11px;color:rgba(255,255,255,.48)}
    .hdr-sp{flex:1}
    .srch-w{position:relative}
    .srch-ic{position:absolute;left:11px;top:50%;transform:translateY(-50%);font-size:13px;color:rgba(255,255,255,.45);pointer-events:none}
    .srch-inp{padding:7px 14px 7px 33px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:20px;color:#fff;font-size:13px;width:210px;outline:none;transition:all .2s}
    .srch-inp::placeholder{color:rgba(255,255,255,.42)}
    .srch-inp:focus{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.4);width:260px}
    .srch-cnt{font-size:11px;color:rgba(255,255,255,.55);min-width:72px}
    .hbtn{padding:6px 13px;border-radius:8px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:#fff;font-size:12px;cursor:pointer;font-weight:500;transition:background .2s;display:flex;align-items:center;gap:5px}
    .hbtn:hover{background:rgba(255,255,255,.2)}

    /* LAYOUT */
    .layout{display:flex;margin-top:var(--hh);height:calc(100vh - var(--hh));overflow:hidden}

    /* SIDEBAR */
    .sb{width:var(--sbw);flex-shrink:0;background:var(--primary);height:100%;overflow-y:auto;padding:10px 0 20px}
    .sb::-webkit-scrollbar{width:4px}
    .sb::-webkit-scrollbar-thumb{background:rgba(255,255,255,.15);border-radius:4px}
    .sb-lbl{padding:12px 18px 5px;font-size:10px;letter-spacing:1.3px;text-transform:uppercase;font-weight:700;color:rgba(255,255,255,.33)}
    .mod-btn{display:flex;align-items:center;gap:11px;width:100%;padding:11px 14px 11px 16px;background:none;border:none;border-left:3px solid transparent;cursor:pointer;text-align:left;transition:background .17s,border-color .17s}
    .mod-btn:hover{background:rgba(255,255,255,.07)}
    .mod-btn.act{background:rgba(255,255,255,.11);border-left-color:var(--accent)}
    .mod-ico{width:36px;height:36px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
    .mod-nm{font-size:13px;font-weight:600;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .mod-sub{font-size:11px;color:rgba(255,255,255,.46);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px}

    /* MAIN */
    .main{flex:1;display:flex;overflow:hidden}
    .cw{flex:1;overflow-y:auto;padding:28px 34px 60px;scrollbar-width:thin;scrollbar-color:#cbd5e0 var(--bg)}
    .cw::-webkit-scrollbar{width:8px}
    .cw::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:4px}

    /* PROGRESS */
    .prog{position:sticky;top:0;height:3px;background:var(--border);margin:-28px -34px 22px;z-index:10}
    .prog-fill{height:100%;background:var(--accent);width:0%;transition:width .1s}

    /* WELCOME */
    .welcome{max-width:660px;margin:48px auto;text-align:center}
    .welcome .emo{font-size:58px;margin-bottom:14px}
    .welcome h1{font-size:25px;font-weight:800;color:var(--primary);margin-bottom:10px}
    .welcome p{font-size:14px;color:var(--muted);line-height:1.7;margin-bottom:34px}
    .wgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:14px}
    .wcard{padding:22px 14px;background:var(--card);border:1px solid var(--border);border-radius:10px;cursor:pointer;transition:all .2s;text-align:center}
    .wcard:hover{box-shadow:0 6px 24px rgba(0,0,0,.1);transform:translateY(-3px);border-color:var(--primary)}
    .wcard-ico{font-size:28px;margin-bottom:9px}
    .wcard-nm{font-size:13px;font-weight:700;color:var(--primary)}
    .wcard-sub{font-size:11px;color:var(--muted);margin-top:3px}

    /* DOC HEADER */
    .dh{display:flex;align-items:center;gap:13px;margin-bottom:24px;padding-bottom:18px;border-bottom:2px solid var(--border)}
    .dh-ico{width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0}
    .dh-info h2{font-size:20px;font-weight:800}
    .dh-info p{font-size:12px;color:var(--muted);margin-top:2px}
    .dh-acts{margin-left:auto;display:flex;gap:8px}
    .dbtn{padding:7px 13px;border-radius:8px;border:1px solid var(--border);background:var(--card);font-size:12px;cursor:pointer;color:var(--text);font-weight:500;transition:all .17s;display:flex;align-items:center;gap:5px}
    .dbtn:hover{background:var(--bg);border-color:var(--primary);color:var(--primary)}

    /* SECTIONS */
    .doc-sec{display:none}
    .doc-sec.show{display:block}

    /* MARKDOWN */
    .mdc{max-width:880px;line-height:1.76;font-size:14px}
    .mdc h1{font-size:22px;color:var(--primary);font-weight:800;margin:32px 0 13px;padding-bottom:8px;border-bottom:2px solid var(--border);scroll-margin-top:18px}
    .mdc h1:first-child{margin-top:0}
    .mdc h2{font-size:17px;color:var(--primary);font-weight:700;margin:26px 0 10px;padding-bottom:6px;border-bottom:1px solid var(--border);scroll-margin-top:18px}
    .mdc h3{font-size:15px;color:var(--primary-l);font-weight:700;margin:18px 0 8px;scroll-margin-top:18px}
    .mdc h4{font-size:13px;color:var(--text);font-weight:700;margin:14px 0 6px;scroll-margin-top:18px}
    .mdc p{margin:0 0 11px}
    .mdc ul,.mdc ol{margin:6px 0 12px 22px}
    .mdc li{margin-bottom:4px}
    .mdc table{width:100%;border-collapse:collapse;font-size:12.5px;margin:16px 0;background:var(--card);border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.07)}
    .mdc th{background:var(--primary);color:#fff;padding:9px 12px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.35px}
    .mdc td{padding:8px 12px;border-bottom:1px solid var(--border);vertical-align:top}
    .mdc tr:last-child td{border-bottom:none}
    .mdc tr:nth-child(even) td{background:#f8fafc}
    .mdc code{background:#f7f8fa;border:1px solid var(--border);padding:1px 6px;border-radius:4px;font-size:12px;font-family:Consolas,Monaco,monospace;color:#c53030}
    .mdc pre{background:#f7f8fa;border:1px solid var(--border);padding:14px 16px;border-radius:8px;overflow-x:auto;margin:12px 0}
    .mdc pre code{background:none;border:none;padding:0;color:var(--text);font-size:12.5px}
    .mdc blockquote{border-left:4px solid var(--accent);margin:12px 0;padding:10px 15px;background:#fff8f5;border-radius:0 8px 8px 0;color:#7c3009;font-size:13px}
    .mdc img{max-width:100%;border-radius:8px;box-shadow:0 2px 14px rgba(0,0,0,.12);margin:10px 0;cursor:zoom-in;display:block}
    .mdc hr{border:none;border-top:2px solid var(--border);margin:22px 0}
    .mdc strong{font-weight:700}
    .mdc em{font-style:italic}
    .mdc u{text-decoration:underline}
    .mdc del{text-decoration:line-through;color:var(--muted)}
    .mdc a{color:var(--accent);text-decoration:none}
    .mdc a:hover{text-decoration:underline}
    mark{background:#fef08a;padding:0 2px;border-radius:2px}

    /* RESULT MARK */
    .result-mark{background:#fef08a;outline:2px solid var(--accent);border-radius:3px;padding:0 2px;scroll-margin-top:120px}
    .result-mark.pulse{animation:markpulse 1.4s ease-out 2}
    @keyframes markpulse{0%{background:#fef08a;outline-color:var(--accent)}50%{background:#fed7aa;outline-color:#ff4500;outline-width:3px}100%{background:#fef08a;outline-color:var(--accent)}}

    /* SEARCH RESULTS */
    .sv-hd{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:22px;padding-bottom:16px;border-bottom:2px solid var(--border)}
    .sv-hd-left{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
    .sv-title{font-size:16px;font-weight:700;color:var(--text)}
    .sv-title strong{color:var(--primary)}
    .sv-badge{background:var(--primary);color:#fff;font-size:11px;font-weight:700;padding:2px 9px;border-radius:10px}
    .sr-mod{margin-bottom:20px;border-radius:10px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.07)}
    .sr-mod-hd{display:flex;align-items:center;gap:9px;padding:10px 14px;color:#fff;font-size:13px;font-weight:700}
    .sr-mod-cnt{font-size:11px;background:rgba(255,255,255,.22);padding:2px 8px;border-radius:10px;margin-left:auto}
    .sr-item{padding:10px 16px;background:var(--card);border-top:1px solid var(--border);cursor:pointer;transition:background .14s}
    .sr-item:hover{background:#eef5ff}
    .sr-path{font-size:11px;color:var(--muted);margin-bottom:4px;display:flex;align-items:center;gap:4px;flex-wrap:wrap}
    .sr-path-seg{color:var(--primary);font-weight:600}
    .sr-path-sep{color:var(--border)}
    .sr-snip{font-size:13px;color:var(--text);line-height:1.55}
    .sr-snip mark{background:#fef08a;padding:0 2px;border-radius:2px;font-style:normal}
    .sr-empty{text-align:center;padding:70px 20px 40px;color:var(--muted)}
    .sr-empty-ico{font-size:50px;margin-bottom:14px}
    .sr-empty p{font-size:14px;line-height:1.6}

    /* TOC */
    .toc{width:var(--tocw);flex-shrink:0;height:100%;overflow-y:auto;background:var(--card);border-left:1px solid var(--border);padding:16px 0}
    .toc::-webkit-scrollbar{width:4px}
    .toc::-webkit-scrollbar-thumb{background:#cbd5e0;border-radius:4px}
    .toc-hd{padding:0 14px 10px;font-size:10px;text-transform:uppercase;letter-spacing:1.2px;font-weight:700;color:var(--muted);border-bottom:1px solid var(--border);margin-bottom:5px}
    .toc-a{display:block;padding:4px 13px;font-size:12px;color:var(--text);text-decoration:none;line-height:1.4;border-left:2px solid transparent;transition:all .14s}
    .toc-a:hover{background:var(--bg);color:var(--primary)}
    .toc-a.act{color:var(--accent);border-left-color:var(--accent);background:#fff8f5;font-weight:600}
    .toc-a.l1{font-weight:600;font-size:12.5px}
    .toc-a.l2{padding-left:23px}
    .toc-a.l3{padding-left:35px;font-size:11px;color:var(--muted)}

    /* IMAGE MODAL */
    .imo{display:none;position:fixed;inset:0;background:rgba(0,0,0,.88);z-index:9999;align-items:center;justify-content:center;padding:22px;cursor:zoom-out}
    .imo.on{display:flex}
    .imo img{max-width:92vw;max-height:90vh;border-radius:8px;object-fit:contain;box-shadow:0 8px 48px rgba(0,0,0,.6);cursor:default}
    .imo-close{position:fixed;top:17px;right:20px;width:38px;height:38px;background:rgba(255,255,255,.15);border:none;border-radius:50%;cursor:pointer;font-size:19px;color:#fff;display:flex;align-items:center;justify-content:center}
    .imo-close:hover{background:rgba(255,255,255,.25)}

    /* BACK TOP */
    .bktop{position:fixed;bottom:26px;right:26px;width:42px;height:42px;border-radius:50%;background:var(--primary);color:#fff;border:none;cursor:pointer;font-size:18px;box-shadow:0 4px 16px rgba(0,0,0,.2);opacity:0;pointer-events:none;transition:opacity .22s;display:flex;align-items:center;justify-content:center}
    .bktop.show{opacity:1;pointer-events:auto}
    .bktop:hover{background:var(--primary-l)}

    /* PRINT */
    @media print{
      .hdr,.sb,.toc,.prog,.bktop,.imo{display:none !important}
      .layout{margin-top:0;height:auto}
      .main,.cw{overflow:visible}
      .cw{padding:0}
      .doc-sec{display:block !important}
      .dh-acts{display:none}
    }
    @media (max-width:1200px){.toc{display:none}}
    @media (max-width:780px){
      :root{--sbw:56px}
      .mod-nm,.mod-sub,.sb-lbl{display:none}
      .cw{padding:16px 14px 60px}
    }
  </style>
</head>
<body>

<!-- HEADER -->
<header class="hdr">
  <div class="hdr-brand" onclick="goHome()" title="Ir al inicio">
    <strong>SGP <span>Upgrade</span></strong>
    <small>Documentación Funcional – Sodexo Chile</small>
  </div>
  <div class="hdr-sp"></div>
  <div class="srch-w">
    <span class="srch-ic">⌕</span>
    <input type="text" class="srch-inp" id="si" placeholder="Buscar en todos los módulos…" autocomplete="off">
  </div>
  <span class="srch-cnt" id="sc"></span>
  <button class="hbtn" onclick="window.print()">🖨 Imprimir</button>
</header>

<!-- LAYOUT -->
<div class="layout">
  <aside class="sb" id="sb"></aside>
  <main class="main">
    <div class="cw" id="cw">
      <div class="prog"><div class="prog-fill" id="pf"></div></div>

      <!-- BIENVENIDA -->
      <div id="wl">
        <div class="welcome">
          <div class="emo">📚</div>
          <h1>Documentación Funcional SGP</h1>
          <p>Selecciona un módulo del panel izquierdo para visualizar su documentación.<br>
             Todos los documentos pertenecen al proyecto <strong>SGP Upgrade – Sodexo Chile</strong>.</p>
          <div class="wgrid" id="wg"></div>
          <p style="margin-top:28px;font-size:11px;color:var(--muted)">Generado el ${ts} · ${mods.length} módulos</p>
        </div>
      </div>

      <!-- RESULTADOS DE BÚSQUEDA -->
      <div id="sv" style="display:none">
        <div class="sv-hd">
          <div class="sv-hd-left">
            <span class="sv-title">Resultados para "<strong id="sq-text"></strong>"</span>
            <span class="sv-badge" id="sq-cnt"></span>
          </div>
          <button class="dbtn" onclick="clearSrch()">✕ Cerrar búsqueda</button>
        </div>
        <div id="sr-list"></div>
      </div>

      <!-- SECCIONES DE DOCUMENTOS -->
      <div id="dv" style="display:none">
        <div class="dh" id="dh"></div>
        <div id="ds"></div>
      </div>
    </div>

    <!-- TOC -->
    <aside class="toc" id="toc" style="display:none">
      <div class="toc-hd">Contenido</div>
      <nav id="tn"></nav>
    </aside>
  </main>
</div>

<!-- MODAL IMAGEN -->
<div class="imo" id="imo" onclick="closeMod(event)">
  <button class="imo-close" onclick="closeMod()">✕</button>
  <img id="mi" src="" alt="">
</div>

<!-- BOTÓN ARRIBA -->
<button class="bktop" id="bt" onclick="goTop()" title="Volver arriba">↑</button>

<script>
const MODS = ${modsJson};
const DOCS = ${docsJson};

marked.setOptions({ gfm:true, breaks:false });

let obsIntersect = null;
let currentId    = null;

/* ── Render secciones al cargar ───────────────────────── */
function initSections() {
  const ds = document.getElementById('ds');
  ds.innerHTML = '';
  for (const m of MODS) {
    const sec = document.createElement('div');
    sec.className = 'doc-sec mdc';
    sec.id = 'sec-' + m.id;
    sec.innerHTML = marked.parse(DOCS[m.id]);
    ds.appendChild(sec);
    addIds(sec);
    sec.querySelectorAll('img').forEach(img => { img.onclick = () => openMod(img.src); });
  }
}

function addIds(container) {
  const seen = {};
  container.querySelectorAll('h1,h2,h3,h4').forEach(h => {
    if (h.id) return;
    const base = h.textContent.trim().toLowerCase()
      .replace(/[áàäâã]/g,'a').replace(/[éèëê]/g,'e')
      .replace(/[íìïî]/g,'i').replace(/[óòöôõ]/g,'o')
      .replace(/[úùüû]/g,'u').replace(/ñ/g,'n')
      .replace(/[^a-z0-9\\s]/g,'').replace(/\\s+/g,'-')
      .replace(/-+/g,'-').replace(/^-|-$/g,'');
    seen[base] = (seen[base]||0)+1;
    h.id = seen[base]>1 ? base+'-'+seen[base] : base;
  });
}

/* ── Sidebar + Welcome grid ───────────────────────────── */
function buildUI() {
  const sb = document.getElementById('sb');
  const wg = document.getElementById('wg');
  sb.innerHTML = '<div class="sb-lbl">Módulos</div>';
  wg.innerHTML = '';
  MODS.forEach(m => {
    const btn = document.createElement('button');
    btn.className = 'mod-btn'; btn.id = 'sb-'+m.id;
    btn.innerHTML = \`<div class="mod-ico" style="background:\${m.bg};color:\${m.color}">\${m.icon}</div>
      <div><div class="mod-nm">\${m.title}</div><div class="mod-sub">\${m.sub}</div></div>\`;
    btn.onclick = () => showMod(m);
    sb.appendChild(btn);

    const c = document.createElement('div');
    c.className = 'wcard';
    c.innerHTML = \`<div class="wcard-ico">\${m.icon}</div>
      <div class="wcard-nm">\${m.title}</div>
      <div class="wcard-sub">\${m.sub}</div>\`;
    c.onclick = () => showMod(m);
    wg.appendChild(c);
  });
}

/* ── Mostrar módulo ───────────────────────────────────── */
function showMod(m) {
  clearResultMarks();
  currentId = m.id;
  document.querySelectorAll('.mod-btn').forEach(b => b.classList.remove('act'));
  document.getElementById('sb-'+m.id).classList.add('act');
  document.getElementById('wl').style.display = 'none';
  document.getElementById('sv').style.display = 'none';
  document.getElementById('dv').style.display = 'block';
  document.getElementById('toc').style.display = 'block';
  document.getElementById('dh').innerHTML = \`
    <div class="dh-ico" style="background:\${m.bg};color:\${m.color}">\${m.icon}</div>
    <div class="dh-info"><h2 style="color:\${m.color}">\${m.title}</h2><p>\${m.sub}</p></div>
    <div class="dh-acts">
      <button class="dbtn" onclick="window.print()">🖨 Imprimir</button>
      <button class="dbtn" onclick="goHome()">🏠 Inicio</button>
    </div>\`;
  document.querySelectorAll('.doc-sec').forEach(s => s.classList.remove('show'));
  document.getElementById('sec-'+m.id).classList.add('show');
  document.getElementById('si').value = '';
  document.getElementById('sc').textContent = '';
  document.getElementById('cw').scrollTop = 0;
  buildToc(m.id);
}

/* ── TOC ──────────────────────────────────────────────── */
function buildToc(id) {
  const nav = document.getElementById('tn');
  nav.innerHTML = '';
  if (obsIntersect) obsIntersect.disconnect();
  const sec = document.getElementById('sec-'+id);
  const cw  = document.getElementById('cw');
  const hs  = [...sec.querySelectorAll('h1,h2,h3')];
  hs.forEach(h => {
    if (!h.id) return;
    const lv = parseInt(h.tagName[1]);
    const a  = document.createElement('a');
    a.href='#'+h.id; a.className='toc-a l'+lv; a.dataset.t=h.id;
    a.textContent=h.textContent.trim();
    a.onclick = e => { e.preventDefault(); h.scrollIntoView({behavior:'smooth',block:'start'}); setTocAct(h.id); };
    nav.appendChild(a);
  });
  obsIntersect = new IntersectionObserver(entries => {
    entries.forEach(en => { if(en.isIntersecting) setTocAct(en.target.id); });
  }, { root:cw, rootMargin:'-15% 0% -75% 0%', threshold:0 });
  hs.forEach(h => { if(h.id) obsIntersect.observe(h); });
}
function setTocAct(id) {
  document.querySelectorAll('.toc-a').forEach(a => a.classList.toggle('act', a.dataset.t===id));
}

/* ── Búsqueda global ──────────────────────────────────── */
let srchTid;
document.getElementById('si').addEventListener('input', function() {
  clearTimeout(srchTid);
  const q = this.value.trim();
  if (!q) { clearSrch(); return; }
  if (q.length < 2) return;
  srchTid = setTimeout(() => globalSearch(q), 350);
});

function escRx(s) { return s.replace(/[-.*+?^$()|[\]\\\\]/g,'\\\\$&'); }

function buildSnippet(text, q) {
  const clean = text.replace(/\\s+/g,' ').trim();
  const idx   = clean.toLowerCase().indexOf(q.toLowerCase());
  const start = Math.max(0, idx-50);
  const end   = Math.min(clean.length, idx+q.length+90);
  const snip  = (start>0?'…':'')+clean.substring(start,end)+(end<clean.length?'…':'');
  return snip.replace(new RegExp('('+escRx(q)+')','gi'),'<mark>$1</mark>');
}

function getNearestHeading(el, container) {
  let node = el;
  while (node && node !== container) {
    let prev = node.previousElementSibling;
    while (prev) { if (/^H[1-3]$/.test(prev.tagName)) return prev; prev=prev.previousElementSibling; }
    node = node.parentElement;
  }
  return null;
}

function globalSearch(q) {
  document.getElementById('wl').style.display  = 'none';
  document.getElementById('dv').style.display  = 'none';
  document.getElementById('toc').style.display = 'none';
  document.getElementById('sv').style.display  = 'block';
  document.getElementById('sq-text').textContent = q;
  document.querySelectorAll('.mod-btn').forEach(b => b.classList.remove('act'));

  const srList = document.getElementById('sr-list');
  srList.innerHTML = '';
  let totalHits = 0;

  for (const m of MODS) {
    const sec = document.getElementById('sec-'+m.id);
    if (!sec) continue;
    const results = [];
    const elements = sec.querySelectorAll('p,li,td,h1,h2,h3,h4,blockquote');
    for (const el of elements) {
      if (!el.textContent.toLowerCase().includes(q.toLowerCase())) continue;
      const heading  = getNearestHeading(el, sec);
      const headText = heading ? heading.textContent.trim() : m.title;
      const headId   = heading ? heading.id : null;
      if (!results.find(r => r.headingId === headId)) {
        results.push({ headingText:headText, headingId:headId, snippet:buildSnippet(el.textContent, q) });
      }
      if (results.length >= 40) break;
    }
    if (!results.length) continue;
    totalHits += results.length;

    const grp = document.createElement('div');
    grp.className = 'sr-mod';
    grp.innerHTML = \`<div class="sr-mod-hd" style="background:\${m.color}">
      <span>\${m.icon}</span> \${m.title}
      <span class="sr-mod-cnt">\${results.length} resultado\${results.length>1?'s':''}</span></div>\`;
    results.forEach(r => {
      const item = document.createElement('div');
      item.className = 'sr-item';
      item.innerHTML = \`<div class="sr-path">
        <span class="sr-path-seg">\${m.title}</span>
        \${r.headingText!==m.title?'<span class="sr-path-sep">›</span><span class="sr-path-seg">'+r.headingText+'</span>':''}
        </div><div class="sr-snip">\${r.snippet}</div>\`;
      item.onclick = () => goToResult(m, r.headingId, q);
      grp.appendChild(item);
    });
    srList.appendChild(grp);
  }

  document.getElementById('sq-cnt').textContent = totalHits
    ? totalHits+' resultado'+(totalHits!==1?'s':'') : '';
  if (!totalHits) srList.innerHTML = \`<div class="sr-empty">
    <div class="sr-empty-ico">🔍</div>
    <p>No se encontraron resultados para "<strong>\${q}</strong>"</p>
    <p style="margin-top:6px;font-size:13px">Intenta con términos más cortos o generales.</p></div>\`;
  document.getElementById('cw').scrollTop = 0;
}

function goToResult(m, headingId, q) {
  showMod(m);
  setTimeout(() => {
    if (headingId) {
      const h = document.getElementById(headingId);
      if (h) h.scrollIntoView({ behavior:'smooth', block:'start' });
    }
    if (q) highlightTermInSection(m.id, headingId, q);
  }, 80);
}

function highlightTermInSection(modId, headingId, q) {
  clearResultMarks();
  const sec = document.getElementById('sec-'+modId);
  if (!sec) return;
  const walker = document.createTreeWalker(sec, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const p = node.parentElement;
      if (!p) return NodeFilter.FILTER_REJECT;
      if (['SCRIPT','STYLE','CODE','PRE'].includes(p.tagName)) return NodeFilter.FILTER_REJECT;
      return node.nodeValue.toLowerCase().includes(q.toLowerCase())
        ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
    }
  });
  const textNodes = [];
  while (walker.nextNode()) textNodes.push(walker.currentNode);

  const rx = new RegExp('('+escRx(q)+')','gi');
  const allMarks = [];
  textNodes.forEach(node => {
    if (!node.parentNode) return;
    const parts = node.nodeValue.split(rx);
    if (parts.length < 2) return;
    const frag = document.createDocumentFragment();
    parts.forEach(part => {
      if (part.toLowerCase() === q.toLowerCase()) {
        const mark = document.createElement('mark');
        mark.className = 'result-mark'; mark.textContent = part;
        allMarks.push(mark); frag.appendChild(mark);
      } else { frag.appendChild(document.createTextNode(part)); }
    });
    node.parentNode.replaceChild(frag, node);
  });

  if (!allMarks.length) return;
  let targetMark = allMarks[0];
  if (headingId) {
    const hEl = document.getElementById(headingId);
    if (hEl) {
      const hTop = hEl.getBoundingClientRect().top;
      for (const mk of allMarks) { if (mk.getBoundingClientRect().top >= hTop) { targetMark=mk; break; } }
    }
  }
  setTimeout(() => { targetMark.scrollIntoView({behavior:'smooth',block:'center'}); targetMark.classList.add('pulse'); }, 250);
}

function clearResultMarks() {
  document.querySelectorAll('mark.result-mark').forEach(mk => {
    const p = mk.parentNode;
    if (!p) return;
    p.replaceChild(document.createTextNode(mk.textContent), mk);
    p.normalize();
  });
}

function clearSrch() {
  clearResultMarks();
  document.getElementById('si').value = '';
  document.getElementById('sc').textContent = '';
  document.getElementById('sv').style.display = 'none';
  if (currentId) {
    document.getElementById('dv').style.display  = 'block';
    document.getElementById('toc').style.display = 'block';
  } else { document.getElementById('wl').style.display = 'block'; }
}

/* ── Modal imagen ─────────────────────────────────────── */
function openMod(src) { document.getElementById('mi').src=src; document.getElementById('imo').classList.add('on'); }
function closeMod(e) { if (!e||e.target!==document.getElementById('mi')) document.getElementById('imo').classList.remove('on'); }
document.addEventListener('keydown', e => {
  if (e.key==='Escape') document.getElementById('imo').classList.remove('on');
  if ((e.ctrlKey||e.metaKey)&&e.key==='f') { e.preventDefault(); const i=document.getElementById('si'); i.focus(); i.select(); }
});

/* ── Scroll / Progreso ────────────────────────────────── */
const cwEl = document.getElementById('cw');
const pfEl = document.getElementById('pf');
const btEl = document.getElementById('bt');
cwEl.addEventListener('scroll', () => {
  const p = cwEl.scrollTop/(cwEl.scrollHeight-cwEl.clientHeight)*100;
  pfEl.style.width = Math.min(p,100)+'%';
  btEl.classList.toggle('show', cwEl.scrollTop>400);
});
function goTop() { cwEl.scrollTo({top:0,behavior:'smooth'}); }
function goHome() {
  clearResultMarks();
  currentId = null;
  document.querySelectorAll('.mod-btn').forEach(b => b.classList.remove('act'));
  document.getElementById('dv').style.display  = 'none';
  document.getElementById('toc').style.display = 'none';
  document.getElementById('sv').style.display  = 'none';
  document.getElementById('si').value = '';
  document.getElementById('sc').textContent = '';
  document.getElementById('wl').style.display  = 'block';
  cwEl.scrollTo({top:0,behavior:'smooth'});
}

/* ── INIT ─────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => { initSections(); buildUI(); });
<\/script>
</body>
</html>`;

  fs.writeFileSync(OUTPUT, html, 'utf8');
  const sz = (fs.statSync(OUTPUT).size / 1024 / 1024).toFixed(2);
  console.log(`  → visor_documentos.html actualizado (${sz} MB)\n`);
}

/* ─── Modo watcher ──────────────────────────────────────────── */
if (WATCH) {
  generate();
  console.log('👁  Modo vigilancia activo. Esperando cambios en archivos .md…');
  console.log('   (Ctrl+C para detener)\n');

  // Debounce: esperar 800 ms tras el último cambio antes de regenerar
  let debounceTimer = null;
  function onFileChange(event, filename) {
    if (!filename || !filename.toLowerCase().endsWith('.md')) return;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      console.log(`\n📝 Cambio detectado: ${filename}`);
      generate();
    }, 800);
  }

  // Vigilar cada subdirectorio existente
  const watched = new Set();
  function watchDir(dir) {
    if (watched.has(dir)) return;
    watched.add(dir);
    try {
      fs.watch(dir, { recursive: false }, (event, filename) => onFileChange(event, filename));
    } catch(e) { /* ignorar directorios que no se pueden vigilar */ }
  }

  // Vigilar el directorio raíz (para detectar nuevas carpetas)
  fs.watch(BASE, { recursive: false }, (event, filename) => {
    // Si se crea un nuevo directorio, añadirlo a la vigilancia
    if (!filename) return;
    const newPath = path.join(BASE, filename);
    try {
      if (fs.statSync(newPath).isDirectory()) {
        if (!watched.has(newPath)) {
          watchDir(newPath);
          console.log(`\n📁 Nuevo directorio detectado: ${filename}`);
        }
      }
    } catch(e) {}
    onFileChange(event, filename);
  });

  // Vigilar subdirectorios ya existentes
  fs.readdirSync(BASE, { withFileTypes: true })
    .filter(d => d.isDirectory() && !d.name.startsWith('.') && !d.name.startsWith('_'))
    .forEach(d => watchDir(path.join(BASE, d.name)));

} else {
  // Generación única
  generate();
}
