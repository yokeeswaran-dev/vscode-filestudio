// Spreadsheet model (src/renderers/sheet.ts): styled.xlsx -> WorkbookModel, and CSV / TSV / PSV / SSV parsing.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const sample = (name) => fs.readFileSync(path.join(__dirname, '..', 'samples', name));

let modelPromise;
/** styled.xlsx loaded once for the whole file. */
function styledModel() {
  modelPromise ??= S.loadWorkbook(new Uint8Array(sample('styled.xlsx'))).then((wb) => S.workbookToModel(wb));
  return modelPromise;
}

/** Cells of rows [start, end) of a sheet as { 'A1': cell }, with each cell's resolved style in `style`. */
function cellsOf(model, sheet, start, end) {
  const styles = new Map(model.getMeta().styles.map((style, i) => [i, style]));
  const { rows, styles: added } = model.getRows(sheet, start, end);
  for (const [i, style] of added) styles.set(i, style);
  const cells = {};
  for (const row of rows) {
    for (const cell of row.cells) {
      cells[`${String.fromCharCode(65 + cell.c)}${row.r + 1}`] = { ...cell, style: styles.get(cell.s) ?? {} };
    }
  }
  return cells;
}

// ----- xlsx -----

test('styled.xlsx: sheet names and visibility states', async () => {
  const { sheets } = (await styledModel()).getMeta();
  assert.deepEqual(
    sheets.map((s) => s.name),
    ['Formatting', 'Numbers', 'CondFormat', 'Hidden', 'Images', 'RTL', 'VeryHidden'],
  );
  assert.deepEqual(
    sheets.map((s) => s.state),
    ['visible', 'visible', 'visible', 'hidden', 'visible', 'visible', 'veryHidden'],
  );
});

test('styled.xlsx: frozen panes, merges and sheet view options', async () => {
  const { sheets } = (await styledModel()).getMeta();
  const [formatting, numbers, , , images, rtl] = sheets;
  assert.deepEqual(formatting.frozen, { rows: 2, cols: 1 });
  assert.deepEqual(numbers.frozen, { rows: 1, cols: 0 });
  assert.equal(formatting.merges.length, 6);
  assert.ok(formatting.merges.some((m) => m.r0 === 1 && m.c0 === 1 && m.r1 === 1 && m.c1 === 9), 'B2:J2 is merged');
  assert.deepEqual([images.images.length, images.showGridLines, images.zoom], [3, false, 125]);
  assert.equal(rtl.rightToLeft, true);
});

test('styled.xlsx: theme colours with tint resolve like Excel', async () => {
  const cells = cellsOf(await styledModel(), 0, 9, 15);
  assert.equal(cells.D10.style.font.color, '#8EA9DB'); // theme 4, tint 0.4
  assert.equal(cells.E10.style.font.color, '#C65911'); // theme 5, tint -0.25
  assert.equal(cells.B15.style.fill, '#D9E1F2'); // theme 4, tint 0.8
  assert.equal(cells.E15.style.fill, '#203764'); // theme 4, tint -0.5
});

test('styled.xlsx: number formats give Excel display texts', async () => {
  const cells = cellsOf(await styledModel(), 1, 0, 41);
  // Column A describes the format, column C holds the formatted value.
  const shown = {};
  for (let r = 2; r <= 41; r++) if (cells[`A${r}`] && cells[`C${r}`]) shown[cells[`A${r}`].w] = cells[`C${r}`].w;
  const expected = {
    'Thousands, 2 decimals (negative)': '-9,876.50',
    Currency: '$1,234.50',
    Percent: '26%',
    Scientific: '1.23E+04',
    Fraction: '1 1/4',
    'Date d-mmm-yy': '15-Mar-24',
    'Long date': 'Friday, March 15, 2024',
    'Time 12-hour': '2:30 PM',
    'Elapsed hours': '36:00:00',
    'Leading zeros': '00501',
    'Conditional sections (millions)': '2.5M',
    'Negative in parentheses': '(4,321)',
  };
  for (const [format, text] of Object.entries(expected)) assert.equal(shown[format], text, format);
  assert.equal(cells.C22.t, 'd', 'a date-formatted number is a date cell');
});

test('styled.xlsx: getRows (formulas, column window) and getStats', async () => {
  const model = await styledModel();
  const cells = cellsOf(model, 1, 43, 50);
  assert.deepEqual([cells.D44.f, cells.D44.v, cells.D44.w], ['B44*C44', 5, '$5.00']);
  assert.equal(cells.E45.f, 'D45/$D$50', 'shared formula is translated to its own cell');
  const window = model.getRows(1, 43, 46, 1, 1).rows;
  assert.deepEqual(window.map((row) => row.cells.map((c) => c.c)), [[1], [1], [1]]);
  // B44:B49 = quantities 10, 24, 5, 12, 3, 8
  const stats = model.getStats(1, [{ r0: 43, c0: 1, r1: 48, c1: 1 }]);
  assert.deepEqual(stats, { count: 6, numCount: 6, sum: 62, avg: 62 / 6, min: 3, max: 24 });
  const text = model.getStats(1, [{ r0: 43, c0: 0, r1: 48, c1: 0 }]);
  assert.deepEqual([text.count, text.numCount, text.sum], [6, 0, 0], 'text cells are counted, not summed');
});

test('loadWorkbook rejects a file that is not a workbook with a readable message', async () => {
  await assert.rejects(S.loadWorkbook(new TextEncoder().encode('hello')), /not a valid \.xlsx/);
});

// ----- xlsx files written by other tools -----

/**
 * An .xlsx written by ExcelJS (which, like openpyxl and pandas, stores number formats as given): `fill` fills its one
 * sheet, `edit` may then change the zip parts (to add what Excel writes and ExcelJS does not).
 */
async function makeXlsx(fill, edit) {
  const workbook = new ExcelJS.Workbook();
  fill(workbook.addWorksheet('Sheet1'));
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  if (edit) await edit(zip);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

/** Adds a <Relationship> to a .rels part of the zip. */
async function addRelationship(zip, rels, id, type, target) {
  const xml = await zip.file(rels).async('string');
  zip.file(rels, xml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`));
}

/** loadWorkbook + workbookToModel, and the cells of the first sheet as { A1: cell }. */
async function loadCells(bytes) {
  const model = S.workbookToModel(await S.loadWorkbook(bytes));
  return { model, cells: cellsOf(model, 0, 0, 100) };
}

test('number formats stored as typed (not escaped as Excel saves them) display like Excel 16', async () => {
  const cases = [
    // [format, value, Excel 16's text (Range.Text)]
    ['dd.mm.yyyy', 45000, '15.03.2023'],
    ['dd.mm.yyyy hh:mm', 45000 + 12.5 / 24, '15.03.2023 12:30'],
    ['hh.mm', 0.78125, '18.45'],
    ['yyyy-mm-ddThh:mm:ss', 45000.5, '2023-03-15T12:00:00'],
    ['yyyy年m月d日', 45000.5, '2023年3月15日'],
    ['dddd, dd.mm.yyyy', 45000.5, 'Wednesday, 15.03.2023'],
    ['0.00\\ E+00', 45000.5, '4.50 E+04'],
    ['0.00\\ E+00', -5, '-5.00 E+00'],
    ['##0.0\\ E+0', 45000.5, '45.0 E+3'],
    ['0.0x', -2.25, '-2.3x'],
    ['0.0 °C', 21.5, '21.5 °C'],
    ['0.0 °C', -3.25, '-3.3 °C'],
    ['General', 8238230.9475, '8238230.948'],
    ['General', -1234567.8905, '-1234567.891'],
    ['m/d/yyyy h:mm', 2958465.5, '12/31/9999 12:00'],
    ['h:mm', 2958465.75, '18:00'],
    ['dddd', 2958465.5, 'Friday'],
    ['[h]:mm', 2958465.5, '71003172:00'],
    ['[h]:mm:ss', 2958465.0001, '71003160:00:09'],
    ['[mm]:ss.000', 2958465.99999, '4260191039:59.136'],
    ['[ss]', 2958465.123456, '255611386667'],
  ];
  const { cells } = await loadCells(
    await makeXlsx((ws) =>
      cases.forEach(([numFmt, value], i) => {
        const cell = ws.getCell(i + 1, 1);
        cell.value = value;
        cell.numFmt = numFmt;
      }),
    ),
  );
  cases.forEach(([numFmt, value, text], i) => assert.equal(cells[`A${i + 1}`].w, text, `${numFmt} ${value}`));
  assert.equal(cells.A1.t, 'd', 'a date format with literal dots is a date');
  assert.equal(cells.A7.t, 'n', "'0.00\\ E+00' is scientific, not a date");
});

test('threaded comments show the thread instead of Excel\'s fallback note', async () => {
  const person = (id, name) => `<person displayName="${name}" id="{${id}}" providerId="None"/>`;
  const comment = (id, parent, person, text) =>
    `<threadedComment ref="A2" dT="2026-10-06T15:05:34.08" personId="{${person}}" id="{${id}}"${parent ? ` parentId="{${parent}}"` : ''}><text>${text}</text></threadedComment>`;
  const ns = 'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"';
  const bytes = await makeXlsx(
    (ws) => {
      ws.getCell('A2').value = 'x';
      // the legacy note Excel writes next to each thread, for older versions
      ws.getCell('A2').note = '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, ...\n\nComment:\n    first post';
      ws.getCell('C1').note = 'legacy note';
    },
    async (zip) => {
      zip.file('xl/persons/person.xml', `<personList ${ns}>${person('P1', 'Ada Lovelace')}${person('P2', 'Alan Turing')}</personList>`);
      zip.file(
        'xl/threadedComments/threadedComment1.xml',
        `<ThreadedComments ${ns}>${comment('T1', '', 'P1', 'first post &amp; more')}${comment('T2', 'T1', 'P2', 'a reply')}</ThreadedComments>`,
      );
      await addRelationship(zip, 'xl/_rels/workbook.xml.rels', 'rIdP', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml');
      await addRelationship(
        zip,
        'xl/worksheets/_rels/sheet1.xml.rels',
        'rIdT',
        'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
        '../threadedComments/threadedComment1.xml',
      );
    },
  );
  const { cells } = await loadCells(bytes);
  assert.equal(cells.A2.note, 'Ada Lovelace: first post & more\nAlan Turing: a reply');
  assert.equal(cells.C1.note, 'legacy note');
});

test('dialog sheets and Excel 4.0 macro sheets are reported as not displayed', async () => {
  const bytes = await makeXlsx(
    (ws) => {
      ws.getCell('A1').value = 1;
    },
    async (zip) => {
      const workbook = await zip.file('xl/workbook.xml').async('string');
      zip.file('xl/workbook.xml', workbook.replace('</sheets>', '<sheet name="Dialog1" sheetId="2" r:id="rIdD"/><sheet name="Macro1" sheetId="3" r:id="rIdM"/></sheets>'));
      const sheet = (tag) => `<${tag} xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></${tag}>`;
      zip.file('xl/dialogsheets/sheet1.xml', sheet('dialogsheet'));
      zip.file('xl/macrosheets/sheet1.xml', sheet('macrosheet'));
      await addRelationship(zip, 'xl/_rels/workbook.xml.rels', 'rIdD', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet', 'dialogsheets/sheet1.xml');
      await addRelationship(zip, 'xl/_rels/workbook.xml.rels', 'rIdM', 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet', 'macrosheets/sheet1.xml');
    },
  );
  const { model } = await loadCells(bytes);
  assert.deepEqual(model.getMeta().sheets.map((s) => s.name), ['Sheet1'], 'only worksheets get a tab');
  assert.deepEqual(await S.detectLossyFeatures(bytes), ['Dialog sheets', 'Macro sheets (Excel 4.0)']);
});

test('_xHHHH_ escapes in inline strings and formula results are decoded like Excel', async () => {
  const rows = [
    // ="x"&CHAR(13)&CHAR(10)&"y" as Excel saves it
    '<c r="A1" t="str"><f>"x"&amp;CHAR(13)&amp;CHAR(10)&amp;"y"</f><v>x_x000D_\ny</v></c>',
    '<c r="A2" t="inlineStr"><is><t>a_x000d_b</t></is></c>',
    '<c r="A3" t="inlineStr"><is><t>lit _x005F_x000D_ end</t></is></c>',
    '<c r="A4" t="str"><v>p_x0009_q</v></c>',
    '<c r="A5" t="inlineStr"><is><t>_xD83D__xDE00_ _x00zz_</t></is></c>',
  ];
  const bytes = await makeXlsx(
    (ws) => {
      ws.getCell('A1').value = 1;
    },
    async (zip) => {
      const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const sheetData = `<sheetData>${rows.map((c, i) => `<row r="${i + 1}">${c}</row>`).join('')}</sheetData>`;
      zip.file('xl/worksheets/sheet1.xml', xml.replace(/<sheetData>[\s\S]*<\/sheetData>/, sheetData));
    },
  );
  const { cells } = await loadCells(bytes);
  // Excel 16 (Range.Value2): one left-to-right pass, hex digits in either case
  assert.deepEqual(
    ['A1', 'A2', 'A3', 'A4', 'A5'].map((a) => cells[a].v),
    ['x\r\ny', 'a\rb', 'lit _x000D_ end', 'p\tq', '\u{1F600} _x00zz_'],
  );
  assert.equal(cells.A1.w, 'x\r\ny');
});

test('selection statistics: error cells are counted apart, values come in the first number\'s format (Excel 16)', async () => {
  const bytes = await makeXlsx((ws) => {
    ws.getCell('A1').value = { formula: '1/0', result: { error: '#DIV/0!' } };
    ws.getCell('A2').value = { error: '#N/A' };
    ws.getCell('A3').value = 0.125;
    ws.getCell('A3').numFmt = '0.0%';
    ws.getCell('A4').value = 1234.5;
    ws.getCell('A4').numFmt = '#,##0.00';
    ws.getCell('A5').value = 'text';
    ws.getCell('B1').value = 46303;
    ws.getCell('B1').numFmt = 'yyyy-mm-dd';
    ws.getCell('B2').value = 46303;
    ws.getCell('B2').numFmt = 'm/d/yyyy';
    ws.getCell('C1').value = 5;
    ws.getCell('C1').numFmt = '"$"#,##0.00';
    ws.getCell('C2').value = 21;
    ws.getCell('D1').value = 3;
    ws.getCell('D2').value = 4;
    ws.getCell('D2').numFmt = '0.0%';
  });
  const { model } = await loadCells(bytes);
  const stats = (r0, c0, r1, c1) => model.getStats(0, [{ r0, c0, r1, c1 }]);
  // A1:A5: Excel shows only "Count 5" when the selection holds an error value; errors are in count, not in numCount.
  const withErrors = stats(0, 0, 4, 0);
  assert.deepEqual([withErrors.count, withErrors.numCount, withErrors.errors], [5, 2, 2]);
  // A3:A4: 0.0% then #,##0.00 -> both in the first format ("Average 61731.3% | Sum 123462.5%").
  assert.deepEqual(stats(2, 0, 3, 0).text, { sum: '123462.5%', avg: '61731.3%', min: '12.5%', max: '123450.0%' });
  assert.equal(stats(2, 0, 3, 0).errors, undefined, 'no error cells: no errors field');
  // 'A4,A3': the first number in the selection's order decides (Excel 16 'B10,B9': "Average 617.31 | Sum 1,234.63").
  const reversed = model.getStats(0, [{ r0: 3, c0: 0, r1: 3, c1: 0 }, { r0: 2, c0: 0, r1: 2, c1: 0 }]);
  assert.deepEqual([reversed.text.sum, reversed.text.avg], ['1,234.63', '617.31']);
  // B1:B2: dates -> "Average 2026-10-08 | Sum 2153-07-17".
  assert.deepEqual(stats(0, 1, 1, 1).text, { sum: '2153-07-17', avg: '2026-10-08', min: '2026-10-08', max: '2026-10-08' });
  // C1:C2: currency -> "Average $13.00 | Sum $26.00".
  assert.deepEqual([stats(0, 2, 1, 2).text.sum, stats(0, 2, 1, 2).text.avg], ['$26.00', '$13.00']);
  // D1:D2: the first number is General -> no text (the viewer formats plain numbers).
  assert.equal(stats(0, 3, 1, 3).text, undefined);
  assert.equal(stats(0, 3, 1, 3).sum, 7);
});

// ----- csv / tsv / psv / ssv -----

/** Parses like src/viewerProvider.ts (parseDelimited) for the file's extension. */
function parseDelimited(text, ext) {
  switch (ext) {
    case '.tsv':
      return S.parseCsv(text, { delimiter: '\t' });
    case '.psv':
      return S.parseCsv(text, { delimiter: '|' });
    case '.ssv':
      return S.detectSsvDelimiter(text) === ';'
        ? S.parseCsv(text, { delimiter: ';' })
        : S.parseCsv(text, { delimiter: ' ', collapseSpaces: true });
    default:
      return S.parseCsv(text);
  }
}

const DELIMITED = [
  // file, delimiter, BOM, rows, a field that must survive parsing
  ['sample.csv', ',', false, 51, [1, 1, 'Smith, John']],
  ['sample-semicolon.csv', ';', true, 16, [1, 3, '8,99']],
  ['sample.tsv', '\t', false, 26, [25, 1, 'Tab\tseparated name']],
  ['sample.psv', '|', false, 31, [1, 2, '/api/search?q=a|b']],
  ['sample.ssv', ';', true, 16, [1, 1, 'Café moulu 250 g']],
  ['sample-space.ssv', ' ', false, 13, [1, 1, 'Ada Lovelace']],
];

for (const [file, delimiter, hasBom, rowCount, [r, c, field]] of DELIMITED) {
  test(`${file}: parses and serializes back byte-identical`, () => {
    const bytes = sample(file);
    const model = parseDelimited(bytes.toString('utf8'), path.extname(file));
    assert.deepEqual([model.delimiter, model.hasBom, model.rows.length], [delimiter, hasBom, rowCount]);
    assert.equal(model.rows[r][c], field);
    assert.ok(Buffer.from(S.serializeCsv(model), 'utf8').equals(bytes), 'serializeCsv(parseCsv(file)) === file');
  });
}

test('sample-space.ssv is read in space mode (runs of spaces, quoted fields)', () => {
  const model = parseDelimited(sample('sample-space.ssv').toString('utf8'), '.ssv');
  assert.equal(model.collapseSpaces, true);
  assert.deepEqual(model.rows[1], ['1', 'Ada Lovelace', '1815-12-10', 'London', 'Mathematics', '98.5', 'First published algorithm']);
});

test('an edited CSV keeps the exact text of the records that did not change', () => {
  const text = sample('sample.csv').toString('utf8');
  const model = S.parseCsv(text);
  assert.equal(model.newline, '\r\n', 'sample.csv has CRLF line endings (.gitattributes keeps them)');
  model.rows[1][1] = 'Doe, Jane';
  const before = text.split(model.newline);
  const after = S.serializeCsv(model).split(model.newline);
  assert.equal(after[1], before[1].replace('"Smith, John"', '"Doe, Jane"'));
  assert.deepEqual(after.filter((_, i) => i !== 1), before.filter((_, i) => i !== 1));
});

test("Excel's 'sep=' first line sets the delimiter and is not a row", () => {
  const text = 'sep=;\r\na;b;c\r\n1;2;3\r\n';
  const model = S.parseCsv(text);
  assert.deepEqual([model.delimiter, model.sepLine, model.rows], [';', 'sep=;\r\n', [['a', 'b', 'c'], ['1', '2', '3']]]);
  assert.equal(S.serializeCsv(model), text, 'the line is written back');
  // after a BOM, any case, also with a forced delimiter (.tsv)
  const bom = S.parseCsv('\uFEFFSEP=,\na,b\n');
  assert.deepEqual([bom.hasBom, bom.delimiter, bom.rows], [true, ',', [['a', 'b']]]);
  assert.deepEqual(S.parseCsv('sep=\t\na\tb\n', { delimiter: '\t' }).rows, [['a', 'b']]);
  // only the first line is a directive
  assert.deepEqual(S.parseCsv('a;b\nsep=;\n1;2\n').rows, [['a', 'b'], ['sep=', ''], ['1', '2']]);
});

test('a CSV sheet name is cut to 31 characters and has ( ) for [ ], like Excel', () => {
  const name = (file) => new S.CsvGridModel(S.parseCsv('a,b\r\n1,2'), file).getMeta().sheets[0].name;
  assert.equal(name(`${'C'.repeat(120)}.csv`), 'C'.repeat(31));
  assert.equal(name('C:\\data\\data [v1].csv'), 'data (v1)');
  assert.equal(name('a.b.c.csv'), 'a.b.c');
});

test('detectSsvDelimiter picks ";" or runs of spaces', () => {
  assert.equal(S.detectSsvDelimiter('name;price;qty\nx;1,5;3\ny;2,25;4\n'), ';');
  assert.equal(S.detectSsvDelimiter('a; b; c\n1; 2; 3\n'), ';');
  assert.equal(S.detectSsvDelimiter('id name note\n1 x a;b\n2 y c\n3 z d\n4 w e\n'), ' ');
  assert.equal(S.detectSsvDelimiter('id note\n1 "a;b"\n2 "c;d"\n3 "e;f"\n'), ' ', 'quoted ";" does not count');
  assert.equal(S.detectSsvDelimiter(''), ' ');
});
