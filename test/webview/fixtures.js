// Test workbooks for the webview suites, built in memory with ExcelJS (no binary fixtures in the repo).
// Each builder returns a Promise<Buffer>; harness.sheetHost loads it once per file name with the real renderer.
// The layouts follow the workbooks the 0.1.x bug hunts used (H4 / H5 / fix-D / fix-U); the comments say which test
// needs which cell.
'use strict';
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const edge = (style, argb = 'FF000000') => ({ style, color: { argb } });
const box = (e) => ({ top: e, left: e, bottom: e, right: e });
const LONG = 'This is a long piece of text that overflows';
const write = (wb) => wb.xlsx.writeBuffer().then((b) => Buffer.from(b));

/** Navigation workbook (H5): data blocks for Ctrl+Arrow, merges, links, number texts, hidden row / column, frozen panes, hidden sheets. */
async function nav() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  // Column A: 1..10, gap, 15..20, lone values at rows 600 / 1000 / 1500 (beyond the first loaded row blocks).
  for (let r = 1; r <= 10; r++) ws.getCell(r, 1).value = r;
  for (let r = 15; r <= 20; r++) ws.getCell(r, 1).value = r;
  ws.getCell(600, 1).value = 600;
  ws.getCell(1000, 1).value = 1000;
  ws.getCell(1500, 1).value = 1500;
  ws.getCell('B2').value = 'M1';
  ws.mergeCells('B2:C3');
  ws.getCell('D5').value = 'wide';
  ws.mergeCells('D5:F5');
  ws.getCell('G1').value = { text: 'ext', hyperlink: 'https://example.com/' };
  ws.getCell('G2').value = { text: 'internal', hyperlink: "#'Other Sheet'!B5" };
  ws.getCell('G3').value = { formula: 'HYPERLINK("other.xlsx#Sheet1!A1","file link")', result: 'file link' };
  ws.getCell('G4').value = { text: 'bad', hyperlink: '#Nope!A1' };
  ws.getCell('G5').value = { formula: 'HYPERLINK("[other.xlsx]Sheet1!A1","bracket link")', result: 'bracket link' };
  ws.getCell('G6').value = { text: 'far', hyperlink: '#Data!A5000' };
  ws.getCell('H1').value = 'a\tb';
  ws.getCell('H2').value = 'line1\nline2';
  ws.getCell('H3').value = 'say "hi"';
  ws.getCell('H4').value = 1234.5;
  ws.getCell('H4').numFmt = '#,##0.00';
  ws.getCell('H5').value = 0.5;
  ws.getCell('H5').numFmt = '0%';
  ws.getCell('H6').value = new Date(Date.UTC(2024, 0, 15));
  ws.getCell('H6').numFmt = 'yyyy-mm-dd';
  ws.getCell('H7').value = 1e21;
  ws.getCell('H8').value = 0.000000000001;
  ws.getCell('H9').value = '123';
  ws.getCell('H10').value = { error: '#DIV/0!' };
  ws.getCell('H11').value = 0.1;
  ws.getCell('H12').value = 0.2;
  ws.getCell('H13').value = -5;
  for (let r = 28; r <= 32; r++) ws.getCell(r, 9).value = r;
  ws.getRow(30).hidden = true;
  ws.getCell('J1').value = 'hidden col';
  ws.getColumn(10).hidden = true;
  ws.getCell('K1').value = 'after hidden';
  const other = wb.addWorksheet('Other Sheet');
  other.getCell('B5').value = 'target';
  const fz = wb.addWorksheet('Frozen', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4', activeCell: 'C4' }] });
  for (let r = 1; r <= 300; r++) for (let c = 1; c <= 40; c++) fz.getCell(r, c).value = r * 100 + c;
  wb.addWorksheet('Secret', { state: 'hidden' }).getCell('A1').value = 'secret';
  wb.addWorksheet('VerySecret', { state: 'veryHidden' }).getCell('A1').value = 'very';
  return write(wb);
}

/** Merges (H5-F6 / H5-F14b): every cell has its address as text; B3:D4, C6:C9 and E2:F2 merged. */
async function merges() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('M');
  for (let r = 1; r <= 12; r++) for (let c = 1; c <= 8; c++) ws.getCell(r, c).value = `${String.fromCharCode(64 + c)}${r}`;
  ws.mergeCells('B3:D4');
  ws.mergeCells('C6:C9');
  ws.mergeCells('E2:F2');
  return write(wb);
}

/** Copy, cell text box and header workbook (U1-U3, R2-F1/F5/F6/F7/F8/F11/F12/F15, U4b). */
async function cells() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Copy');
  for (let c = 1; c <= 10; c++) ws.getColumn(c).width = 14;
  ws.getCell('D1').value = 1;
  ws.getCell('D2').value = 2;
  ws.getCell('D3').value = 'text';
  ws.getCell('D4').value = 4;
  ws.getCell('B2').value = 'Merged title B2:C2';
  ws.mergeCells('B2:C2');
  ws.getCell('C5').value = 'a\tb';
  ws.getCell('B6').value = 'Short';
  ws.getCell('C6').value = 'Line one\nLine two\nLine three';
  ws.getCell('C6').alignment = { wrapText: true, vertical: 'top' };
  ws.getRow(6).height = 48;
  ws.getCell('E8').value = { text: 'Example site', hyperlink: 'https://example.com/' };
  ws.getCell('E10').value = 'noted';
  ws.getCell('E10').note = 'A cell note';
  // R2-F5: text cells that Excel would turn into numbers / dates / formulas on paste; A14 a number.
  ws.getCell('A10').value = '007';
  ws.getCell('A11').value = '1-2';
  ws.getCell('A12').value = '=1+1';
  ws.getCell('A13').value = '3/4';
  ws.getCell('A14').value = 12.5;
  // R2-F1 / R2-F8: hidden row 31, hidden column G.
  ws.getCell('B30').value = 'r30';
  ws.getCell('B31').value = 'r31';
  ws.getCell('B32').value = 'r32';
  ws.getRow(31).hidden = true;
  ws.getCell('F30').value = 'f30';
  ws.getCell('G30').value = 'g30';
  ws.getCell('H30').value = 'h30';
  ws.getColumn(7).hidden = true;
  // R2-F12: header clicks next to merges.
  ws.getCell('B35').value = 'merge B35:D35';
  ws.mergeCells('B35:D35');
  ws.getCell('B37').value = 'merge B37:C38';
  ws.mergeCells('B37:C38');
  const rtl = wb.addWorksheet('RTL', { views: [{ rightToLeft: true }] });
  rtl.getColumn(2).width = 20;
  rtl.getCell('B2').value = 'مرحبا بالعالم';
  wb.addWorksheet('Hid', { state: 'hidden' }).getCell('A1').value = 'hidden';
  wb.addWorksheet('VeryHid', { state: 'veryHidden' }).getCell('A1').value = 'very hidden';
  return write(wb);
}

/** U4a: the active tab is a hidden sheet (Excel opens on it). */
async function hiddenActive() {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('One').getCell('A1').value = 'one';
  wb.addWorksheet('Two', { state: 'hidden' }).getCell('A1').value = 'two';
  wb.addWorksheet('Three').getCell('A1').value = 'three';
  wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];
  return write(wb);
}

/** U4c: a link into a hidden sheet. */
async function hiddenLink() {
  const wb = new ExcelJS.Workbook();
  const one = wb.addWorksheet('One');
  one.getCell('A1').value = 'links';
  one.getCell('A2').value = { text: 'to hidden Two!B2', hyperlink: '#Two!B2' };
  wb.addWorksheet('Two', { state: 'hidden' }).getCell('B2').value = 'target';
  return write(wb);
}

/** Many tabs (H5-F12 / R2-F13 tab-scroll arrows, H9-C4 coloured tabs). */
async function tabs() {
  const wb = new ExcelJS.Workbook();
  const colors = ['FFFF0000', 'FF00B050', 'FF0070C0', 'FFFFFF00', 'FF000000', 'FFFFFFFF', 'FF7030A0', 'FF002060'];
  for (let i = 0; i < 40; i++) {
    const ws = wb.addWorksheet(`Sheet ${i + 1}`, i < 8 ? { properties: { tabColor: { argb: colors[i] } } } : undefined);
    ws.getCell('A1').value = `sheet ${i + 1}`;
    ws.getCell('B2').value = i;
  }
  return write(wb);
}

/** Rendering cases (H4): overflow / clip / alignment / merges / borders / decorations. Columns 12 wide. */
async function render() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Render');
  for (let c = 1; c <= 12; c++) ws.getColumn(c).width = 12;
  // 1: H4-F2 formula returning "" in B1 blocks the overflow of A1.
  ws.getCell('A1').value = LONG;
  ws.getCell('B1').value = { formula: 'IF(1=0,"x","")', result: '' };
  // 3: H4-F1 right-aligned text in a merge, too long -> Excel shows the END.
  ws.mergeCells('B3:C3');
  ws.getCell('B3').value = 'START merged right-aligned text that is too long END';
  ws.getCell('B3').alignment = { horizontal: 'right' };
  // 5: H4-F1 centred text in a merge, too long -> Excel shows the middle.
  ws.mergeCells('B5:C5');
  ws.getCell('B5').value = 'START merged centred text that is too long END';
  ws.getCell('B5').alignment = { horizontal: 'center' };
  // 7: H4-F4 Center Across Selection over A7:E7.
  ws.getCell('A7').value = 'Title';
  for (let c = 1; c <= 5; c++) ws.getCell(7, c).alignment = { horizontal: 'centerContinuous' };
  // 9: H4-F5 fill alignment.
  ws.getCell('A9').value = '-';
  ws.getCell('A9').alignment = { horizontal: 'fill' };
  ws.getCell('B9').value = 'ab';
  ws.getCell('B9').alignment = { horizontal: 'fill' };
  ws.getCell('D9').value = 'fill-aligned long text must not overflow';
  ws.getCell('D9').alignment = { horizontal: 'fill' };
  // 11: H4-F6 distributed + indent 1.
  ws.getCell('D11').value = 'distrib ind';
  ws.getCell('D11').alignment = { horizontal: 'distributed', indent: 1 };
  // 13: H4-F13 double underline + strike; H4-F14 superscript run of 24 pt.
  ws.getCell('A13').value = 'dbl+strike';
  ws.getCell('A13').font = { underline: 'double', strike: true };
  ws.getCell('C13').value = { richText: [{ text: 'x' }, { text: '2', font: { vertAlign: 'superscript', size: 24 } }, { text: ' big sup', font: { size: 24 } }] };
  ws.getRow(13).height = 36;
  // 15: H4-F10 shrink-to-fit rich text with a 14 pt run.
  ws.getCell('B15').value = { richText: [{ text: 'Rich ', font: { size: 11 } }, { text: 'shrink run 14pt text', font: { size: 14, bold: true } }] };
  ws.getCell('B15').alignment = { shrinkToFit: true };
  // 17: H4-F15 wrapped numbers -> ### (column L is narrow).
  ws.getColumn(12).width = 8;
  ws.getCell('L17').value = 12345678901;
  ws.getCell('L17').numFmt = '0';
  ws.getCell('L17').alignment = { wrapText: true };
  ws.getCell('L18').value = 3.14159265358979;
  ws.getCell('L18').alignment = { wrapText: true };
  ws.getRow(17).height = 45;
  ws.getRow(18).height = 45;
  // 20: H4-F11 rotated text.
  ws.getRow(20).height = 70;
  [45, -45, 30].forEach((r, i) => {
    const cell = ws.getCell(20, i + 1);
    cell.value = `rot ${r}`;
    cell.alignment = { textRotation: r };
  });
  // H4-F8 merges whose edge cells carry the borders (Excel's layout), on their own sheet.
  const m = wb.addWorksheet('Merges');
  for (let c = 1; c <= 12; c++) m.getColumn(c).width = 10;
  const blue = edge('thin', 'FF0000FF');
  const merge = (r0, c0, r1, c1, master, covered = () => ({})) => {
    m.mergeCells(r0, c0, r1, c1);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) m.getCell(r, c).style = r === r0 && c === c0 ? master : covered(r, c);
  };
  merge(2, 2, 4, 4, { border: box(edge('medium', 'FFC00000')) });
  m.getCell('B2').value = 'master-only box';
  merge(2, 6, 2, 8, { border: { top: blue, left: blue, bottom: blue }, fill: solid('FFFFF2CC') }, (r, c) => ({ border: { top: blue, bottom: blue, ...(c === 8 ? { right: blue } : {}) } }));
  m.getCell('F2').value = 'edge cells';
  // H4-F9 conflicting borders on a shared edge at the frozen column edge (B | C) and away from it (G | H).
  const f = wb.addWorksheet('Frozen', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4' }] });
  for (let c = 1; c <= 10; c++) f.getColumn(c).width = 14;
  const red = { style: 'thick', color: { argb: 'FFFF0000' } };
  const navy = { style: 'thick', color: { argb: 'FF0000FF' } };
  f.getCell('B10').border = { right: red };
  f.getCell('C10').border = { left: navy };
  f.getCell('B10').value = 'B10 thick red R';
  f.getCell('C10').value = 'C10 thick blue L';
  f.getCell('G10').border = { right: red };
  f.getCell('H10').border = { left: navy };
  f.getCell('G10').value = 'G10 thick red R';
  f.getCell('H10').value = 'H10 thick blue L';
  f.getCell('F3').border = { bottom: red };
  f.getCell('F4').border = { top: navy };
  f.getCell('F3').value = 'F3 bottom';
  f.getCell('F4').value = 'F4 top';
  return write(wb);
}

/** K8: a right-to-left sheet (first, active) and its left-to-right twin. */
async function rtl() {
  const wb = new ExcelJS.Workbook();
  for (const dir of ['RTL', 'LTR']) {
    const ws = wb.addWorksheet(dir, { views: [{ rightToLeft: dir === 'RTL' }] });
    for (let c = 1; c <= 6; c++) ws.getColumn(c).width = 14;
    ws.getRow(1).height = 70;
    ws.getCell('A1').value = 'rot 30';
    ws.getCell('A1').alignment = { textRotation: 30 };
    ws.getCell('B2').value = 'Left';
    ws.getCell('B2').alignment = { horizontal: 'left' };
    ws.getCell('C3').value = 1234;
  }
  return write(wb);
}

// ===== WORD =====

/**
 * A minimal .docx around `body` (WordprocessingML of <w:body>), plus optional hyperlink relationships
 * ({ id, target }) — enough for mammoth (the real renderDocx) to read.
 */
async function docx(body, links = []) {
  const zip = new JSZip();
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      links.map((l) => `<Relationship Id="${l.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${l.target}" TargetMode="External"/>`).join('') +
      '</Relationships>',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** A paragraph of plain runs: para('text'), para('text', '<w:bidi/>'). */
const para = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

// ===== PDF =====

/**
 * A minimal PDF: pages = [{ w, h, lines: ['text', ...] }] in Helvetica 18 pt (pdf.js can find the text). No pages
 * gives a valid file with an empty page tree.
 * @param {{ w?: number, h?: number, lines?: string[] }[]} pages
 */
function pdf(pages) {
  const objs = ['<< /Type /Catalog /Pages 2 0 R >>', '', '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
  const kids = [];
  const esc = (t) => t.replace(/[\\()]/g, (c) => `\\${c}`);
  for (const p of pages) {
    const w = p.w || 612;
    const h = p.h || 792;
    const content = `BT /F1 18 Tf ${(p.lines || []).map((t, i) => `1 0 0 1 72 ${h - 72 - i * 28} Tm (${esc(t)}) Tj`).join(' ')} ET`;
    objs.push(`<< /Length ${Buffer.byteLength(content, 'latin1')} >>\nstream\n${content}\nendstream`);
    const c = objs.length;
    objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /Font << /F1 3 0 R >> >> /Contents ${c} 0 R >>`);
    kids.push(objs.length);
  }
  objs[1] = `<< /Type /Pages /Kids [${kids.map((k) => `${k} 0 R`).join(' ')}] /Count ${kids.length} >>`;
  let out = '%PDF-1.7\n%\xe2\xe3\xcf\xd3\n';
  const offsets = objs.map((o, i) => {
    const at = Buffer.byteLength(out, 'latin1');
    out += `${i + 1} 0 obj\n${o}\nendobj\n`;
    return at;
  });
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
  out += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Promise.resolve(Buffer.from(out, 'latin1'));
}

module.exports = { nav, merges, cells, hiddenActive, hiddenLink, tabs, render, rtl, docx, para, pdf };
