const fs      = require('fs');
const path    = require('path');
const JSZip   = require('jszip');
const { XMLParser } = require('fast-xml-parser');

const DOCX = path.join(__dirname, 'Words', 'DRF_Inventario.docx');
const parserOpts = { ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text' };
const parser = new XMLParser(parserOpts);

function getText(el) {
  if (!el) return '';
  if (typeof el === 'string') return el;
  if (typeof el === 'number') return String(el);
  const t = el['w:t'];
  if (t !== undefined) {
    if (Array.isArray(t)) return t.map(x => typeof x === 'object' ? (x['#text'] || '') : String(x)).join('');
    if (typeof t === 'object') return t['#text'] || '';
    return String(t);
  }
  const r = el['w:r'];
  if (r !== undefined) {
    const arr = Array.isArray(r) ? r : [r];
    return arr.map(getText).join('');
  }
  return '';
}

async function main() {
  const data = fs.readFileSync(DOCX);
  const zip  = await JSZip.loadAsync(data);

  console.log('\n=== ARCHIVOS EN EL DOCX ===');
  Object.keys(zip.files).sort().forEach(f => console.log('  ' + f));

  const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('word/media/') && !zip.files[f].dir);
  console.log(`\n=== IMÁGENES (${mediaFiles.length}) ===`);
  mediaFiles.sort().forEach(f => console.log('  ' + f));

  const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  const relsDoc = parser.parse(relsXml);
  const relArr  = [].concat(relsDoc?.Relationships?.Relationship || []);
  console.log('\n=== RELACIONES ===');
  relArr.forEach(r => {
    const type = (r['@_Type'] || '').split('/').pop();
    console.log(`  ${r['@_Id']} | ${type} | ${r['@_Target']}`);
  });

  const docXml = await zip.file('word/document.xml').async('string');
  const doc    = parser.parse(docXml);
  const body   = doc?.['w:document']?.['w:body'];

  // Contar tablas
  const tables = [].concat(body?.['w:tbl'] || []);
  console.log(`\n=== TABLAS: ${tables.length} ===`);

  // Párrafos: style + texto
  const paras = [].concat(body?.['w:p'] || []);
  console.log(`\n=== PÁRRAFOS (${paras.length} total, mostrando con texto) ===`);
  let shown = 0;
  paras.forEach((p, i) => {
    const style = p?.['w:pPr']?.['w:pStyle']?.['@_w:val'] || 'Normal';
    const numPr = p?.['w:pPr']?.['w:numPr'];
    const text  = getText(p).trim();
    if (text && shown < 60) {
      const num = numPr ? ` [numPr ilvl=${numPr['w:ilvl']?.['@_w:val']} numId=${numPr['w:numId']?.['@_w:val']}]` : '';
      console.log(`  [${i}] [${style}${num}] ${text.substring(0, 100)}`);
      shown++;
    }
  });

  // Comentarios
  if (zip.file('word/comments.xml')) {
    const cXml  = await zip.file('word/comments.xml').async('string');
    const cDoc  = parser.parse(cXml);
    const cArr  = [].concat(cDoc?.['w:comments']?.['w:comment'] || []);
    console.log(`\n=== COMENTARIOS: ${cArr.length} ===`);
    cArr.slice(0, 5).forEach(c => {
      const paras = [].concat(c?.['w:p'] || []);
      const txt   = paras.map(getText).join(' ').trim();
      console.log(`  [id=${c['@_w:id']}] ${c['@_w:author']} (${(c['@_w:date']||'').substring(0,10)}): ${txt.substring(0,80)}`);
    });
    if (cArr.length > 5) console.log(`  ... y ${cArr.length - 5} más`);
  }

  if (zip.file('word/numbering.xml')) {
    console.log('\n=== numbering.xml: PRESENTE ===');
    const nXml = await zip.file('word/numbering.xml').async('string');
    const nDoc = parser.parse(nXml);
    const nums = [].concat(nDoc?.['w:numbering']?.['w:num'] || []);
    console.log(`  ${nums.length} definiciones de numeración`);
  }

  if (zip.file('word/styles.xml')) {
    const sXml   = await zip.file('word/styles.xml').async('string');
    const sDoc   = parser.parse(sXml);
    const styles = [].concat(sDoc?.['w:styles']?.['w:style'] || []);
    const headings = styles.filter(s => (s['@_w:styleId']||'').startsWith('Heading'));
    console.log('\n=== ESTILOS DE HEADING ===');
    headings.forEach(s => console.log(`  ${s['@_w:styleId']}`));
  }
}

main().catch(console.error);
