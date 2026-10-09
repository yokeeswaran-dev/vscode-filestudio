// Spreadsheet model (src/renderers/sheet.ts): styled.xlsx -> WorkbookModel, and CSV / TSV / PSV / SSV parsing.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

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

test('detectSsvDelimiter picks ";" or runs of spaces', () => {
  assert.equal(S.detectSsvDelimiter('name;price;qty\nx;1,5;3\ny;2,25;4\n'), ';');
  assert.equal(S.detectSsvDelimiter('a; b; c\n1; 2; 3\n'), ';');
  assert.equal(S.detectSsvDelimiter('id name note\n1 x a;b\n2 y c\n3 z d\n4 w e\n'), ' ');
  assert.equal(S.detectSsvDelimiter('id note\n1 "a;b"\n2 "c;d"\n3 "e;f"\n'), ' ', 'quoted ";" does not count');
  assert.equal(S.detectSsvDelimiter(''), ' ');
});
