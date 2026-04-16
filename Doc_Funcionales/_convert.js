/**
 * _convert.js — Convierte DRF_Inventario.docx a Markdown
 * Paso a paso según prompt_lectura_word_a_md.md
 */
'use strict';
const fs    = require('fs');
const path  = require('path');
const JSZip = require('jszip');
const { XMLParser } = require('fast-xml-parser');
const sharp = require('sharp');

/* ── Parámetros ────────────────────────────────────────── */
const DOCX    = path.join(__dirname, 'Words', 'DRF_Inventario.docx');
const OUT_DIR = path.join(__dirname, 'MD', 'DRF_Inventario');
const IMG_DIR = path.join(OUT_DIR, 'imagenes');
const OUT_MD  = path.join(OUT_DIR, 'DRF_Inventario.md');

/* ── Parser XML con orden preservado (para document.xml) ── */
const P_OPTS = {
  ignoreAttributes:   false,
  attributeNamePrefix:'@_',
  textNodeName:       '#text',
  preserveOrder:      true,
  trimValues:         false,
  parseAttributeValue:false,
};
const parser = new XMLParser(P_OPTS);

/* ── Parser plano (para rels, comments, numbering) ───────── */
const parserFlat = new XMLParser({
  ignoreAttributes:   false,
  attributeNamePrefix:'@_',
  textNodeName:       '#text',
});

/* ── Estilos → nivel de heading ─────────────────────────  */
const H_LEVEL = {
  'Ttulo1':1,'Ttulo2':2,'Ttulo3':3,'Ttulo4':4,'Ttulo5':5,
  'Titulo1':1,'Titulo2':2,'Titulo3':3,'Titulo4':4,
  'Heading1':1,'Heading2':2,'Heading3':3,'Heading4':4,
  '1':1,'2':2,'3':3,'4':4,  // algunos documentos usan el índice
};

/* ── Estilos de lista ───────────────────────────────────── */
const LIST_STYLES = new Set([
  'Prrafodelista','Prrafodelist','ListBullet','ListBullet2',
  'ListNumber','ListNumber2','ListParagraph',
]);

/* ── Contadores de numeración automática ───────────────── */
const numCounters = {};
function incrNum(numId, ilvl) {
  const k = `${numId}:${ilvl}`;
  numCounters[k] = (numCounters[k] || 0) + 1;
  // Reiniciar niveles hijos
  for (let l = ilvl + 1; l < 10; l++) delete numCounters[`${numId}:${l}`];
  const parts = [];
  for (let l = 0; l <= ilvl; l++) parts.push(numCounters[`${numId}:${l}`] || 1);
  return parts.join('.') + '.';
}

/* ══════════════════════════════════════════════════════════
   NAVEGADORES del formato preserveOrder
   ══════════════════════════════════════════════════════════ */
function tagName(node) {
  return Object.keys(node).find(k => k !== ':@' && k !== '#text');
}
function children(node) {
  const t = tagName(node);
  return t ? (node[t] || []) : [];
}
function attrs(node) { return node[':@'] || {}; }
function attr(node, name) { return (node[':@'] || {})[name] || ''; }

function findFirst(arr, tag) {
  if (!Array.isArray(arr)) return null;
  const found = arr.find(c => c[tag] !== undefined);
  return found ? found[tag] : null;
}
function findAll(arr, tag) {
  if (!Array.isArray(arr)) return [];
  return arr.filter(c => c[tag] !== undefined).map(c => c[tag]);
}

/* ══════════════════════════════════════════════════════════
   EXTRACCIÓN DE TEXTO E IMÁGENES DE UN PÁRRAFO
   ══════════════════════════════════════════════════════════ */

// Extrae el rId de una imagen dentro de w:drawing
function extractImageRid(drawingChildren) {
  // Buscar recursivamente a:blip[@r:embed]
  function walk(arr) {
    if (!Array.isArray(arr)) return null;
    for (const node of arr) {
      if (typeof node !== 'object') continue;
      const tag = tagName(node);
      if (!tag) continue;
      if (tag === 'a:blip') {
        const rid = attr(node, '@_r:embed');
        if (rid) return rid;
      }
      const rid = walk(children(node));
      if (rid) return rid;
    }
    return null;
  }
  return walk(drawingChildren);
}

// Procesa runs de un párrafo; retorna { text, imageRids[] }
function processParaContent(paraChildren, hyperlinkMap) {
  let text = '';
  const imageRids = [];

  for (const node of (paraChildren || [])) {
    const tag = tagName(node);
    if (!tag) continue;

    if (tag === 'w:r') {
      const rc = children(node);
      // Verificar si tiene imagen
      const drawing = findFirst(rc, 'w:drawing');
      if (drawing) {
        const rid = extractImageRid(drawing);
        if (rid) imageRids.push(rid);
        continue;
      }
      // Texto normal
      const tNodes = findAll(rc, 'w:t');
      let runText = '';
      for (const tChildren of tNodes) {
        for (const tc of tChildren) {
          if (tc['#text'] !== undefined) runText += String(tc['#text']);
        }
      }
      if (!runText) continue;
      // Formato del run
      const rPr = findFirst(rc, 'w:rPr') || [];
      const isBold   = rPr.some(c => c['w:b']      !== undefined && attr(c,'@_w:val') !== 'false');
      const isItal   = rPr.some(c => c['w:i']      !== undefined && attr(c,'@_w:val') !== 'false');
      const isUline  = rPr.some(c => c['w:u']      !== undefined && attr(c,'@_w:val') !== 'none');
      const isStrike = rPr.some(c => c['w:strike'] !== undefined && attr(c,'@_w:val') !== 'false');
      if (isStrike) runText = `~~${runText}~~`;
      if (isUline)  runText = `<u>${runText}</u>`;
      if (isBold && isItal) runText = `***${runText}***`;
      else if (isBold)  runText = `**${runText}**`;
      else if (isItal)  runText = `*${runText}*`;
      text += runText;

    } else if (tag === 'w:hyperlink') {
      // Extraer texto del hyperlink
      const hc = children(node);
      let hText = '';
      const hRid = attr(node, '@_r:id');
      for (const hNode of hc) {
        if (tagName(hNode) === 'w:r') {
          const rc2 = children(hNode);
          const tNodes = findAll(rc2, 'w:t');
          for (const tChildren of tNodes) {
            for (const tc of tChildren) {
              if (tc['#text'] !== undefined) hText += String(tc['#text']);
            }
          }
        }
      }
      const url = hyperlinkMap[hRid] || '';
      if (hText && url) text += `[${hText}](${url})`;
      else if (hText) text += hText;

    } else if (tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd') {
      // Se manejan a nivel superior
    }
  }

  return { text: text.trim(), imageRids };
}

// Texto simple de un párrafo (sin formato, para tablas/celdas)
function plainText(paraChildren) {
  let t = '';
  for (const node of (paraChildren || [])) {
    const tag = tagName(node);
    if (tag === 'w:r') {
      const tNodes = findAll(children(node), 'w:t');
      for (const tc of tNodes)
        for (const x of tc)
          if (x['#text'] !== undefined) t += String(x['#text']);
    } else if (tag === 'w:hyperlink') {
      t += plainText(children(node));
    }
  }
  return t.trim();
}

// Obtener IDs de comentarios referenciados en el párrafo
function getCommentIds(paraChildren) {
  const ids = [];
  for (const node of (paraChildren || [])) {
    const tag = tagName(node);
    if (tag === 'w:commentRangeEnd') {
      const id = attr(node, '@_w:id');
      if (id) ids.push(id);
    }
  }
  return ids;
}

/* ══════════════════════════════════════════════════════════
   MAIN
   ══════════════════════════════════════════════════════════ */
async function main() {
  fs.mkdirSync(IMG_DIR, { recursive: true });

  /* ── Cargar DOCX ─────────────────────────────────────── */
  const data = fs.readFileSync(DOCX);
  const zip  = await JSZip.loadAsync(data);
  console.log('✓ DOCX cargado');

  /* ── PASO 2: Extraer imágenes ────────────────────────── */
  // Construir mapa rId → filename original  (parser PLANO, no preserveOrder)
  const relsXml  = await zip.file('word/_rels/document.xml.rels').async('string');
  const relsFlat = parserFlat.parse(relsXml);
  const relArr   = [].concat(relsFlat?.Relationships?.Relationship || []);
  const ridToFile = {};
  for (const r of relArr) {
    const id     = r['@_Id']     || '';
    const type   = r['@_Type']   || '';
    const target = r['@_Target'] || '';
    if (type.endsWith('/image')) ridToFile[id] = path.basename(target);
  }

  // Extraer y convertir imágenes en orden de nombre (para tabla de correspondencia)
  const mediaFiles = Object.keys(zip.files)
    .filter(f => f.startsWith('word/media/') && !zip.files[f].dir)
    .sort((a, b) => {
      // Ordenar numéricamente: image1 < image2 < image10
      const na = parseInt(path.basename(a).replace(/\D/g,''), 10) || 0;
      const nb = parseInt(path.basename(b).replace(/\D/g,''), 10) || 0;
      return na - nb;
    });

  // fileToImgNum: image1.png → imagen_01.jpg
  const fileToImgNum = {};
  let imgCounter = 1;
  for (const mf of mediaFiles) {
    const base = path.basename(mf);
    const outName = `imagen_${String(imgCounter).padStart(2,'0')}.jpg`;
    fileToImgNum[base] = outName;
    const outPath = path.join(IMG_DIR, outName);
    if (!fs.existsSync(outPath)) {
      const buf = await zip.file(mf).async('nodebuffer');
      try {
        await sharp(buf)
          .flatten({ background: { r: 255, g: 255, b: 255 } })  // PNG transparente → fondo blanco
          .jpeg({ quality: 90 })
          .toFile(outPath);
      } catch (e) {
        // Copiar sin convertir si sharp falla (EMF/WMF/etc)
        fs.writeFileSync(outPath.replace('.jpg', path.extname(base)), buf);
        fileToImgNum[base] = `imagen_${String(imgCounter).padStart(2,'0')}${path.extname(base)}`;
        console.warn(`  ⚠ No convertida: ${base} → copiada como original`);
      }
    }
    imgCounter++;
  }
  console.log(`✓ ${mediaFiles.length} imágenes extraídas → imagenes/`);

  // Mapa rId → nombre de archivo de salida
  const ridToImgFile = {};
  for (const [rid, fname] of Object.entries(ridToFile)) {
    if (fileToImgNum[fname]) ridToImgFile[rid] = fileToImgNum[fname];
  }

  /* ── Mapa de hyperlinks ──────────────────────────────── */
  const hyperlinkMap = {};
  for (const r of relArr) {
    if ((r['@_Type'] || '').endsWith('/hyperlink'))
      hyperlinkMap[r['@_Id']] = r['@_Target'] || '';
  }

  /* ── PASO 3d: Comentarios (parser PLANO) ─────────────── */
  const commentMap = {};
  if (zip.file('word/comments.xml')) {
    const cXml  = await zip.file('word/comments.xml').async('string');
    const cFlat = parserFlat.parse(cXml);
    const cArr  = [].concat(cFlat?.['w:comments']?.['w:comment'] || []);
    for (const c of cArr) {
      const id     = c['@_w:id']     || '';
      const author = c['@_w:author'] || '';
      const date   = (c['@_w:date'] || '').substring(0, 10);
      // Extraer texto del comentario (plano)
      const paras  = [].concat(c?.['w:p'] || []);
      let txt = '';
      for (const p of paras) {
        const runs = [].concat(p?.['w:r'] || []);
        for (const r of runs) {
          const t = r?.['w:t'];
          if (typeof t === 'string') txt += t;
          else if (t?.['#text']) txt += t['#text'];
        }
      }
      commentMap[id] = { author, date, text: txt.trim() };
    }
  }
  console.log(`✓ ${Object.keys(commentMap).length} comentarios leídos`);

  /* ── PASO 3-4: Procesar documento ───────────────────── */
  const docXml    = await zip.file('word/document.xml').async('string');
  const docParsed = parser.parse(docXml);
  const docElem   = docParsed.find(c => c['w:document']);
  const docBody   = findFirst(docElem['w:document'], 'w:body') || [];

  const mdLines  = [];
  const headings = []; // { level, text, anchor }
  let coverSkipped = false;
  let inToc       = false;
  let imgInDoc    = 0; // contador de imágenes encontradas en orden del documento

  // Mapa para que las imágenes salgan en el orden en que aparecen en el doc
  // rId → imagen_NN.jpg (se asigna al primer encuentro)
  const ridToDocOrder = {};
  let docImgCounter = 1;

  function resolveImgRid(rid) {
    if (!ridToDocOrder[rid]) {
      // Asignar número en orden de aparición en el documento
      const originalFile = ridToFile[rid];
      if (originalFile && fileToImgNum[originalFile]) {
        // Usar el número basado en nombre de archivo original
        ridToDocOrder[rid] = fileToImgNum[originalFile];
      } else {
        ridToDocOrder[rid] = `imagen_${String(docImgCounter).padStart(2,'0')}.jpg`;
        docImgCounter++;
      }
    }
    return ridToDocOrder[rid];
  }

  /* ── Helpers ──────────────────────────────────────────── */
  function makeAnchor(text) {
    return text.toLowerCase()
      .replace(/[^\w\s\-áéíóúüñàèìòùâêîôûäëïöü]/g, '')
      .replace(/\s+/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
  }

  function processTable(tblChildren) {
    const rows = findAll(tblChildren, 'w:tr');
    const lines = [];
    let headerDone = false;
    for (const rowChildren of rows) {
      const cells = findAll(rowChildren, 'w:tc');
      const cellTexts = cells.map(cellChildren => {
        const paraNodes = findAll(cellChildren, 'w:p');
        return paraNodes.map(pc => plainText(pc)).join(' ').replace(/\|/g, '\\|').trim();
      });
      lines.push('| ' + cellTexts.join(' | ') + ' |');
      if (!headerDone) {
        lines.push('| ' + cellTexts.map(() => '---').join(' | ') + ' |');
        headerDone = true;
      }
    }
    return lines;
  }

  /* ── Detección de section break en párrafo ───────────── */
  function hasSectionBreak(pChildren) {
    const pPr = findFirst(pChildren, 'w:pPr') || [];
    return pPr.some(c => tagName(c) === 'w:sectPr');
  }

  /* ── Detección de TOC ────────────────────────────────── */
  function isTocSdt(sdtChildren) {
    const sdtPr = findFirst(sdtChildren, 'w:sdtPr') || [];
    for (const child of sdtPr) {
      const t = tagName(child);
      if (t === 'w:tag') {
        const val = (attr(child,'@_w:val') || '').toLowerCase();
        if (val.includes('toc')) return true;
      }
    }
    // Verificar instrText con TOC
    function hasTocInstr(arr) {
      if (!Array.isArray(arr)) return false;
      for (const n of arr) {
        if (tagName(n) === 'w:instrText') {
          const itChildren = children(n);
          for (const c of itChildren) {
            if (typeof c['#text'] === 'string' && c['#text'].includes('TOC')) return true;
          }
        }
        if (hasTocInstr(children(n))) return true;
      }
      return false;
    }
    return hasTocInstr(sdtChildren);
  }

  function isTocPara(styleId) {
    const l = (styleId || '').toLowerCase().replace(/\s/g,'');
    return l.startsWith('toc') || l === 'tableofcontents';
  }

  /* ── Procesar cada elemento del body ────────────────── */
  for (const bodyNode of docBody) {
    const tag = tagName(bodyNode);
    if (!tag || tag === 'w:sectPr') continue;

    /* Saltear TOC (w:sdt) */
    if (tag === 'w:sdt') {
      const sdtC = children(bodyNode);
      if (!isTocSdt(sdtC)) {
        // No es TOC: procesar contenido interno
        const sdtContent = findFirst(sdtC, 'w:sdtContent') || [];
        for (const inner of sdtContent) {
          const innerTag = tagName(inner);
          if (innerTag === 'w:p' || innerTag === 'w:tbl') {
            // Procesar normalmente (ver abajo)
          }
        }
      }
      continue;
    }

    /* Tablas */
    if (tag === 'w:tbl') {
      if (!coverSkipped) continue;
      const tableLines = processTable(children(bodyNode));
      if (tableLines.length) {
        mdLines.push('');
        mdLines.push(...tableLines);
        mdLines.push('');
      }
      continue;
    }

    /* Párrafos */
    if (tag === 'w:p') {
      const pC = children(bodyNode);

      // Obtener estilo
      const pPr  = findFirst(pC, 'w:pPr') || [];
      const styleNode = pPr.find(c => c['w:pStyle'] !== undefined);
      const styleId   = styleNode ? attr(styleNode, '@_w:val') : 'Normal';

      // Detectar portada
      if (!coverSkipped) {
        const hLevel = H_LEVEL[styleId];
        if (hLevel || styleId.startsWith('Ttulo') || styleId.startsWith('Titulo')) {
          coverSkipped = true;
        } else {
          if (hasSectionBreak(pC)) coverSkipped = true;
          continue;
        }
      }

      // Saltar estilos TOC
      if (isTocPara(styleId)) continue;

      // Obtener numPr
      const numPrNode = pPr.find(c => c['w:numPr'] !== undefined);
      const numPrC    = numPrNode ? children(numPrNode) : null;
      let numId = null, ilvl = 0;
      if (numPrC) {
        const ilvlNode  = numPrC.find(c => c['w:ilvl'] !== undefined);
        const numIdNode = numPrC.find(c => c['w:numId'] !== undefined);
        ilvl  = parseInt(attr(ilvlNode,  '@_w:val') || '0', 10);
        numId = attr(numIdNode, '@_w:val') || null;
      }

      // Contenido del párrafo
      const { text, imageRids } = processParaContent(pC, hyperlinkMap);
      const commentIds = getCommentIds(pC);

      // Imágenes en este párrafo
      const imgLines = [];
      for (const rid of imageRids) {
        const imgFile = resolveImgRid(rid);
        imgLines.push(`![Imagen ${docImgCounter}](imagenes/${imgFile})`);
      }

      // Nivel de heading
      const hLevel = H_LEVEL[styleId];
      if (hLevel !== undefined) {
        // Heading
        let headText = text;

        // Numeración automática: añadir si numId > 0 y el texto no comienza ya con número
        if (numId && numId !== '0' && !/^\d/.test(headText)) {
          const numStr = incrNum(numId, ilvl);
          headText = `${numStr} ${headText}`;
        } else if (numId && numId === '0') {
          // Sin numeración automática (como Historial de Versiones)
        }

        const anchor = makeAnchor(headText);
        headings.push({ level: hLevel, text: headText, anchor });
        const hashes = '#'.repeat(hLevel);
        mdLines.push('');
        mdLines.push(`${hashes} ${headText}`);
        mdLines.push('');

      } else if (LIST_STYLES.has(styleId) || styleId.toLowerCase().includes('list')) {
        // Lista
        const indent = '  '.repeat(ilvl);
        if (text) {
          mdLines.push(`${indent}- ${text}`);
          for (const il of imgLines) mdLines.push(il);
        }

      } else {
        // Párrafo normal
        if (imgLines.length) {
          mdLines.push('');
          for (const il of imgLines) mdLines.push(il);
          mdLines.push('');
        }
        if (text) {
          mdLines.push('');
          mdLines.push(text);
          mdLines.push('');
        }
      }

      // Comentarios al final del párrafo
      for (const cid of commentIds) {
        const c = commentMap[cid];
        if (c) {
          const fecha = c.date || '';
          mdLines.push('');
          mdLines.push(`> 💬 **Comentario — ${c.author} (${fecha}):** ${c.text}`);
          mdLines.push('');
        }
      }
    }
  }

  console.log(`✓ Documento procesado: ${mdLines.length} líneas MD generadas`);
  console.log(`✓ Encabezados encontrados: ${headings.length}`);

  /* ── PASO 5: Generar índice ──────────────────────────── */
  const idxLines = ['', '---', '', '## Índice', ''];
  for (const h of headings) {
    const indent = '  '.repeat(h.level - 1);
    idxLines.push(`${indent}- [${h.text}](#${h.anchor})`);
  }
  idxLines.push('', '---', '');

  /* ── PASO 6: Escribir archivo final ──────────────────── */
  // Título del MD según regla del prompt:
  // DRF_Inventario.md → "DRF - Inventario"
  const mdName    = path.basename(OUT_MD, '.md');
  const mdTitle   = mdName.replace(/_/g,' - ').replace(/-/g,' ');
  // "DRF  Inventario" → "DRF - Inventario"
  const titleLine = `# ${mdName.replace(/_/g,' - ')}`;

  // Limpiar líneas en blanco múltiples consecutivas
  const cleanLines = [];
  let prevBlank = false;
  for (const line of [titleLine, ...idxLines, ...mdLines]) {
    const isBlank = line.trim() === '';
    if (isBlank && prevBlank) continue;
    cleanLines.push(line);
    prevBlank = isBlank;
  }

  fs.writeFileSync(OUT_MD, cleanLines.join('\n'), 'utf8');

  const stats  = fs.statSync(OUT_MD);
  const sizeKB = (stats.size / 1024).toFixed(0);
  console.log(`\n✅ Archivo generado: ${OUT_MD}`);
  console.log(`   Tamaño: ${sizeKB} KB`);
  console.log(`   Encabezados: ${headings.length}`);
  console.log(`   Imágenes referenciadas: ${Object.keys(ridToDocOrder).length}`);
  console.log(`   Comentarios insertados: ${Object.keys(commentMap).length}`);

  // Verificaciones
  console.log('\n── Verificaciones ──────────────────────────');
  const mdContent = fs.readFileSync(OUT_MD, 'utf8');
  const imgRefs = (mdContent.match(/!\[.*?\]\(imagenes\//g) || []).length;
  console.log(`  Imágenes en carpeta:       ${mediaFiles.length}`);
  console.log(`  Imágenes referenciadas MD: ${imgRefs}`);
  console.log(`  Índice presente:           ${mdContent.includes('## Índice') ? 'SÍ' : 'NO'}`);
  console.log(`  Comentarios en MD:         ${(mdContent.match(/💬 \*\*Comentario/g) || []).length}`);
  const firstHeadings = headings.slice(0,5).map(h => h.text);
  console.log('  Primeros 5 encabezados:');
  firstHeadings.forEach(h => console.log(`    · ${h}`));
}

main().catch(err => { console.error('ERROR:', err); process.exit(1); });
