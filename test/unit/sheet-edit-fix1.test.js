// Regression tests for the P2 review fixes (round 1): xlsx core editing.
//   - Delete / clear skips rows hidden by an autofilter (Excel keeps their data)
//   - a date cell's edit text reads back as the same date in every locale (ISO text, not toLocaleDateString)
//   - automatic percent entry: 5 typed into a 0.0% cell is 5%
//   - an edit that grows the workbook past the 100,000-cell edit limit is refused (open-time size rule)
//   - copy / paste inside FileStudio keeps exact numbers and text that looks like a value
//   - one cell twice in one setCells: the last value wins; _xlfn. / _xlws. prefixes do not skip the spill block
// Run with `npm test` (it bundles the renderers and media/viewer.js into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');

const S = require('../.out/sheet.js');

globalThis.acquireVsCodeApi ??= () => ({ postMessage() {}, getState: () => undefined, setState() {} });
globalThis.window ??= { addEventListener() {} };
globalThis.document ??= { getElementById: () => null };
globalThis.navigator ??= { platform: '', userAgent: '' };
const V = require('../.out/viewer.js');

async function builtModel(build) {
  const wb = new ExcelJS.Workbook();
  build(wb);
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  return S.workbookToModel(await S.loadWorkbook(bytes));
}

const setCells = (sheetId, ...cells) => ({ kind: 'setCells', sheetId, cells });
const plannedCells = (plan) => plan.after.ops[plan.after.ops.length - 1].cells;

// ----- autofilter: Delete keeps filtered-out rows -----

test('fix1: clear inside an autofilter skips the rows the filter hides (Excel Clear Contents)', async () => {
  const model = await builtModel((wb) => {
    const ws = wb.addWorksheet('S');
    ws.addRow(['name', 'n']);
    for (const [k, v] of [['a', 1], ['b', 2], ['c', 3], ['d', 4], ['e', 5], ['f', 6]]) ws.addRow([k, v]);
    ws.autoFilter = 'A1:B7';
    for (const r of [3, 4, 5]) ws.getRow(r).hidden = true;
    ws.getCell('D4').value = 'hidden row, other column'; // in a filtered-out row: not visible, kept too
    ws.getRow(9).hidden = true;
    ws.getCell('A9').value = 'hidden, no filter';
  });
  const id = model.getMeta().sheets[0].id;
  const plan = S.planEdit(model, { kind: 'clear', sheetId: id, ranges: [{ r0: 1, c0: 0, r1: 8, c1: 3 }], what: 'contents' });
  assert.ok(plan.kind, JSON.stringify(plan));
  const rows = [...new Set(plannedCells(plan).map(([r]) => r))].sort((a, b) => a - b);
  // rows 3-5 (0-based 2-4) hidden by the filter keep b|2, c|3, d|4 and D4 (only visible cells are cleared); row 9 is
  // hidden by hand below the filter range: cleared (Excel does the same without a filter).
  assert.deepEqual(rows, [1, 5, 6, 8]);
  assert.ok(!plannedCells(plan).some(([r, c]) => r >= 2 && r <= 4 && c <= 1), 'no cell of a filtered-out row');
});

// ----- dates: locale-independent edit text -----

test('fix1: a date cell starts editing with ISO text that parseCellInput reads back as the same date', () => {
  V.__setGrid({ kind: 'sheet', meta: { date1904: false, sheets: [] } });
  const cases = [
    [45296, '2024-01-05'], // 5 Jan 2024: en-GB showed 05/01/2024, read back as May 1
    [45651, '2024-12-25'], // en-GB 25/12/2024 was stored as text
    [45296.5, '2024-01-05 12:00:00'],
    [0.75, '18:00:00'],
    [45296 + (13 * 3600 + 5 * 60 + 7.25) / 86400, '2024-01-05 13:05:07.250'],
  ];
  for (const [serial, want] of cases) {
    const text = V.rawValue({ t: 'd', v: serial, w: 'whatever the locale shows' });
    assert.equal(text, want, `serial ${serial}`);
    const back = S.parseCellInput(text, { numFmt: 'dd/mm/yyyy' });
    assert.equal(back.kind, 'number', text);
    assert.ok(Math.abs(back.value - serial) < 1e-8, `${text} -> ${back.value}, want ${serial}`);
  }
  V.__setGrid({ kind: 'sheet', meta: { date1904: true, sheets: [] } });
  assert.equal(V.rawValue({ t: 'd', v: 43834 }), '2024-01-05'); // 1904 system: serial 43834 = 2024-01-05
  assert.equal(S.parseCellInput('2024-01-05', { date1904: true }).value, 43834);
  V.__setGrid(null);
});

// ----- automatic percent entry -----

test('fix1: parseCellInput: a typed plain number of 1 or more in a percent cell is a percent (Excel automatic percent entry)', () => {
  const typed = (t, numFmt) => S.parseCellInput(t, { numFmt, autoPercent: true });
  assert.deepEqual(typed('5', '0.0%'), { kind: 'number', value: 0.05 });
  assert.deepEqual(typed('50', '0%'), { kind: 'number', value: 0.5 });
  assert.deepEqual(typed('-5', '0.00%'), { kind: 'number', value: -0.05 });
  assert.deepEqual(typed('12.5', '0.0%'), { kind: 'number', value: 0.125 });
  assert.deepEqual(typed('1,250', '0%'), { kind: 'number', value: 12.5 });
  assert.equal(typed('0.25', '0%').value, 0.25, 'below 1: already a fraction');
  assert.equal(typed('5%', '0.0%').value, 0.05, 'a typed % is not divided again');
  assert.equal(typed('5', 'General').value, 5);
  assert.equal(typed('5', '0.0"%"').value, 5, 'a quoted % is text in the format, not a percent');
  assert.equal(S.parseCellInput('5', { numFmt: '0.0%' }).value, 5, 'not typed (paste): as is');
  assert.equal(typed('2024-01-05', '0%').value, 45296, 'dates are not percents');
  assert.equal(S.isPercentFormat('0.00%;[Red]-0.00%'), true);
  assert.equal(S.isPercentFormat('#,##0'), false);
});

test('fix1: typing 5 into a 0.0% cell stores 0.05; a pasted block keeps the typed numbers', async () => {
  const model = await builtModel((wb) => {
    const ws = wb.addWorksheet('S');
    for (const ref of ['A1', 'A2', 'A3']) {
      ws.getCell(ref).value = 0.07;
      ws.getCell(ref).numFmt = '0.0%';
    }
  });
  const id = model.getMeta().sheets[0].id;
  const single = S.planEdit(model, setCells(id, [0, 0, '5']));
  assert.deepEqual(plannedCells(single)[0][2].value, { t: 'n', v: 0.05 });
  const block = S.planEdit(model, setCells(id, [1, 0, '5'], [2, 0, '6']));
  assert.deepEqual(plannedCells(block).map(([, , p]) => p.value.v), [5, 6]);
});

// ----- size rule for edits -----

test('fix1: an edit that grows a small workbook past 100,000 cells is refused; one that fits is allowed', async () => {
  const model = await builtModel((wb) => {
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.getCell('C2').value = 2;
  });
  const id = model.getMeta().sheets[0].id;
  const far = S.planEdit(model, setCells(id, [4999, 25, 'x'])); // Z5000: 5000 x 26 = 130,000 cells
  assert.ok(far.kind);
  assert.equal(S.editGrowsTooLarge(model, far), S.XLSX_EDIT_GROWS_TOO_LARGE_TEXT);
  const cells = [];
  for (let r = 0; r < 317; r++) for (let c = 0; c < 317; c++) cells.push([r, c, '1']);
  assert.equal(S.editGrowsTooLarge(model, S.planEdit(model, { kind: 'setCells', sheetId: id, cells })), S.XLSX_EDIT_GROWS_TOO_LARGE_TEXT, '317 x 317');
  const fits = [];
  for (let r = 0; r < 316; r++) for (let c = 0; c < 316; c++) fits.push([r, c, '1']);
  assert.equal(S.editGrowsTooLarge(model, S.planEdit(model, { kind: 'setCells', sheetId: id, cells: fits })), undefined, '316 x 316');
  assert.equal(S.editGrowsTooLarge(model, S.planEdit(model, setCells(id, [0, 0, '9']))), undefined, 'inside the used range');
});

// ----- copy / paste inside FileStudio -----

function copyValues(cellsByRef, range) {
  const byRow = new Map();
  for (const [ref, cell] of Object.entries(cellsByRef)) {
    const c = ref.charCodeAt(0) - 65;
    const r = Number(ref.slice(1)) - 1;
    if (!byRow.has(r)) byRow.set(r, { r, cells: [] });
    byRow.get(r).cells.push({ c, ...cell });
  }
  for (const row of byRow.values()) row.cells.sort((a, b) => a.c - b.c);
  const meta = { cols: {}, rows: {}, merges: [] };
  const view = { index: 0, meta, mergeIndex: new Map(), sel: { ranges: [range], active: { r: range.r0, c: range.c0 } } };
  V.__setGrid({ kind: 'sheet', meta: { date1904: false, sheets: [meta] }, views: new Map([[0, view]]), sheet: 0 });
  const mode = range.r0 === range.r1 && range.c0 === range.c1 ? 'cell' : 'rows';
  const payload = V.buildCopyPayload(view, [range], mode, (r) => byRow.get(r) ?? null);
  V.__setGrid(null);
  return payload;
}

test('fix1: copy keeps the exact value for paste inside FileStudio (rounded numbers, text that reads as a value)', () => {
  const p = copyValues(
    {
      A1: { t: 'n', v: 1234.5678, w: '$1,234.57' },
      A2: { t: 'n', v: 123456789.12345, w: '123456789.1' },
      A3: { t: 'n', v: 0.00001234, w: '0.0000123' },
      A4: { t: 'n', v: 5, w: '5.00 kg' },
      A5: { t: 's', v: '007', w: '007' },
      A6: { t: 's', v: '=1+1', w: '=1+1' },
      A7: { t: 's', v: 'true', w: 'true' },
      A8: { t: 's', v: ' lead', w: ' lead' },
      A9: { t: 'n', v: 0.05123, w: '5.1%' },
      A10: { t: 'n', v: 42, w: '42' },
      A11: { t: 'n', v: 3, w: '3', f: 'A10/14' },
    },
    { r0: 0, c0: 0, r1: 10, c1: 0 },
  );
  assert.equal(p.text.split('\r\n')[0], '$1,234.57', 'the clipboard text is still the display text');
  assert.deepEqual(p.values.map((r) => r[0]), ['$1234.5678', '123456789.12345', '0.00001234', '5', "'007", "'=1+1", "'true", ' lead', '5.123%', '42', '3']);
  // What the host reads from those texts: the original values (General target cells).
  const read = p.values.map(([t]) => S.parseCellInput(t));
  assert.equal(read[0].value, 1234.5678);
  assert.equal(read[1].value, 123456789.12345);
  assert.equal(read[2].value, 0.00001234);
  assert.equal(read[3].value, 5);
  assert.deepEqual(read[4], { kind: 'text', text: '007', quotePrefix: true });
  assert.deepEqual(read[5], { kind: 'text', text: '=1+1', quotePrefix: true });
  assert.deepEqual(read[6], { kind: 'text', text: 'true', quotePrefix: true });
  assert.deepEqual(read[7], { kind: 'text', text: ' lead', quotePrefix: false });
  assert.equal(read[8].value, 0.05123);
  const one = copyValues({ A1: { t: 's', v: '007', w: '007' } }, { r0: 0, c0: 0, r1: 0, c1: 0 });
  assert.deepEqual(one.values, [["'007"]], 'single-cell copy too');
});

// ----- small planner fixes -----

test('fix1: the same cell twice in one setCells: the last value wins (also when the last is the old value)', async () => {
  const model = await builtModel((wb) => {
    wb.addWorksheet('S').getCell('A1').value = 5;
  });
  const id = model.getMeta().sheets[0].id;
  assert.deepEqual(S.planEdit(model, setCells(id, [0, 0, '7'], [0, 0, '5'])), { noop: true });
  const plan = S.planEdit(model, setCells(id, [0, 0, '5'], [0, 1, 'x'], [0, 0, '8']));
  assert.deepEqual(plannedCells(plan).map(([r, c, p]) => [r, c, p.value]), [[0, 0, { t: 'n', v: 8 }], [0, 1, { t: 's', v: 'x' }]]);
});

test('fix1: _xlfn. / _xlws. prefixes do not get a spilling function past the block', () => {
  for (const f of ['_xlfn.SEQUENCE(3)', '_xlws.SORT(A1:A3)', '_xlfn._xlws.FILTER(A1:A3,A1:A3>1)']) {
    const check = S.checkFormulaInput(f, ['Sheet1']);
    assert.equal(check.ok, false, f);
    assert.match(check.block.message, /spilling/, f);
  }
  assert.equal(S.checkFormulaInput('_xlfn.IFNA(1,2)', ['Sheet1']).ok, true);
});
