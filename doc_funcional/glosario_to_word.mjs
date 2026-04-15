import { readFileSync, writeFileSync } from 'fs';
import {
  Document, Packer, Table, TableRow, TableCell,
  Paragraph, TextRun, WidthType, BorderStyle,
  AlignmentType, HeadingLevel, ShadingType
} from './node_modules/docx/dist/index.mjs';

const MD_PATH  = './md_pantallas/MD_Glosario.md';
const OUT_PATH = './md_pantallas/MD_Glosario.docx';

// --- Parsear MD ---
const md    = readFileSync(MD_PATH, 'utf-8');
const lines = md.split(/\r?\n/);

const title = lines[0].replace(/^#+\s*/, '').trim();
const desc  = (lines.find(l => l.startsWith('>')) ?? '').replace(/^>\s*/, '').trim();

const tableLines = lines.filter(l => l.startsWith('|') && !l.match(/^\|[-:| ]+\|$/));

function parseCell(raw) {
  return raw
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/\*(.*?)\*/g, '$1')
    .replace(/`(.*?)`/g, '$1')
    .trim();
}
function parseCells(line) {
  const parts = line.split('|').map(s => s.trim()).filter((_, i, a) => i > 0 && i < a.length - 1);
  return parts.map(parseCell);
}

const headerCells = parseCells(tableLines[0]);
const rows        = tableLines.slice(1).map(l => parseCells(l));

// --- Estilos ---
const PRIMARY = '1E3A5F';
const ROW_ALT = 'EBF3FB';
const ROW_NRM = 'FFFFFF';
const BORDER  = { style: BorderStyle.SINGLE, size: 6, color: 'B0C4D8' };

function cell(text, { bold = false, bg = ROW_NRM, fg = '333333', pct = 50, sz = 19 } = {}) {
  return new TableCell({
    width:   { size: pct, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: bg, color: 'auto' },
    borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
    margins: { top: 80, bottom: 80, left: 130, right: 130 },
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children:  [new TextRun({ text, bold, color: fg, size: sz, font: 'Calibri' })]
    })]
  });
}

const headerRow = new TableRow({
  tableHeader: true,
  children: [
    cell(headerCells[0], { bold: true, bg: PRIMARY, fg: 'FFFFFF', pct: 22, sz: 20 }),
    cell(headerCells[1], { bold: true, bg: PRIMARY, fg: 'FFFFFF', pct: 78, sz: 20 })
  ]
});

const dataRows = rows.map((cells, i) => {
  const bg = i % 2 === 0 ? ROW_NRM : ROW_ALT;
  return new TableRow({ children: [
    cell(cells[0] ?? '', { bold: true,  bg, fg: PRIMARY, pct: 22 }),
    cell(cells[1] ?? '', { bold: false, bg, fg: '333333', pct: 78 })
  ]});
});

// --- Documento ---
const doc = new Document({
  sections: [{
    properties: { page: { margin: { top: 900, bottom: 900, left: 1100, right: 1100 } } },
    children: [
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: title, bold: true, color: PRIMARY, size: 36, font: 'Calibri' })]
      }),
      new Paragraph({
        spacing: { after: 280 },
        children: [new TextRun({ text: desc, italics: true, color: '666666', size: 19, font: 'Calibri' })]
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows:  [headerRow, ...dataRows]
      }),
      new Paragraph({
        spacing: { before: 200 },
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({
          text: `Generado: ${new Date().toLocaleDateString('es-CL')} — SGP Local`,
          italics: true, color: 'AAAAAA', size: 16, font: 'Calibri'
        })]
      })
    ]
  }]
});

const buf = await Packer.toBuffer(doc);
writeFileSync(OUT_PATH, buf);
console.log(`Listo: ${OUT_PATH}`);
console.log(`Términos: ${rows.length}`);
