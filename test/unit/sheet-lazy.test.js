// 0.2.1 big sheets: a worksheet above the size threshold loads row by row ("lazy"). These tests lower the threshold so that
// small files exercise the same code: window == full load, shared formulas across windows, eviction, the guard for rows
// that are not loaded, edits on a lazy sheet, fail-closed eligibility, and the byte rule of edits (B1).
// Run with `npm test`.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const ROWS = 1200;

async function bigBytes(build) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Big');
  ws.columns = [{ width: 14 }, { width: 20 }, { width: 12 }];
  ws.getRow(1).values = ['id', 'text', 'sum'];
  ws.getRow(1).height = 24;
  ws.getRow(1).font = { bold: true };
  for (let r = 2; r <= ROWS; r++) {
    const row = ws.getRow(r);
    row.getCell(1).value = r - 1;
    row.getCell(2).value = r % 7 === 0 ? `Text with <angle> & "quotes" ${r}` : `row ${r}`;
    row.getCell(3).value = { formula: `A${r}*2`, result: (r - 1) * 2 };
    if (r % 50 === 0) row.hidden = true;
    if (r % 11 === 0) row.getCell(4).value = new Date(Date.UTC(2024, 0, r % 28 + 1));
    if (r % 13 === 0) row.getCell(1).numFmt = '0.00';
  }
  ws.addWorksheet; // (keeps the linter quiet)
  const small = wb.addWorksheet('Small');
  small.getCell('A1').value = 'hello';
  build?.(wb, ws);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** The same sheet with the formulas of column C as one shared formula (master in C2, copies below). */
async function sharedBytes(build) {
  const zip = await JSZip.loadAsync(await bigBytes(build));
  const name = Object.keys(zip.files).find((n) => /worksheets\/sheet1\.xml$/.test(n));
  let xml = await zip.file(name).async('string');
  xml = xml.replace(/<c r="C(\d+)"([^>]*)><f>[^<]*<\/f>/g, (_m, r, attrs) =>
    Number(r) === 2 ? `<c r="C2"${attrs}><f t="shared" ref="C2:C${ROWS}" si="0">A2*2</f>` : `<c r="C${r}"${attrs}><f t="shared" si="0"/>`,
  );
  zip.file(name, xml);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

async function bothModels(bytes, tuning = { sheetBytes: 1000, chunkRows: 40, keepRows: 150, maxRequestRows: 2000, sstBytes: 2 * 1024 * 1024, scanBytes: 4 << 20 }) {
  Object.assign(S.lazyTuning, tuning);
  const reasons = [];
  const lazyWb = await S.loadWorkbookLazy(bytes, {}, undefined, (part, reason) => reasons.push(reason));
  return { lazy: S.workbookToModel(lazyWb), full: S.workbookToModel(await S.loadWorkbook(bytes)), lazyWb, reasons };
}

const norm = (rr) => {
  const st = new Map(rr.styles);
  return JSON.stringify(rr.rows.map((r) => ({ r: r.r, cells: r.cells.map((c) => ({ ...c, s: c.s ? st.get(c.s) : 0 })) })));
};

test.afterEach(() => Object.assign(S.lazyTuning, { sheetBytes: 3 * 1024 * 1024, chunkRows: 500, keepRows: 6000, maxRequestRows: 10_000, sstBytes: 2 * 1024 * 1024, scanBytes: 4 << 20 }));

test('lazy: a big sheet is lazy, a small one and a sheet under the threshold are not', async () => {
  const bytes = await bigBytes();
  const { lazy, lazyWb, reasons } = await bothModels(bytes);
  assert.deepEqual(reasons, []);
  assert.equal(S.lazySheetCount(lazyWb), 1);
  assert.equal(lazy.isLazyIndex(0), true);
  assert.equal(lazy.isLazyIndex(1), false);
  const off = await bothModels(bytes, { sheetBytes: 50_000_000, chunkRows: 500, keepRows: 6000, maxRequestRows: 10_000, sstBytes: 2 * 1024 * 1024, scanBytes: 4 << 20 });
  assert.equal(S.lazySheetCount(off.lazyWb), 0);
});

test('lazy: meta (used range, row and column properties) equals the full load', async () => {
  const { lazy, full } = await bothModels(await bigBytes());
  const a = full.getMeta().sheets;
  const b = lazy.getMeta().sheets;
  for (const k of Object.keys(a[0])) assert.deepEqual(b[0][k], a[0][k], `meta.${k}`);
  assert.deepEqual(b[1], a[1]);
});

test('lazy: every window equals the full load (values, formulas, dates, styles, hidden rows)', async () => {
  const { lazy, full } = await bothModels(await bigBytes());
  for (const [a, b] of [[0, 30], [1, 2], [45, 60], [39, 81], [1190, 1210], [1199, 1200], [0, 1200], [500, 520], [7, 9]]) {
    await lazy.ensureRows(0, [{ r0: a, r1: b }]);
    assert.equal(norm(lazy.getRows(0, a, b, 0, 10)), norm(full.getRows(0, a, b, 0, 10)), `rows ${a}-${b}`);
  }
});

test('lazy: shared formulas whose first cell is in another window show the same formulas', async () => {
  const { lazy, full, reasons } = await bothModels(await sharedBytes());
  assert.deepEqual(reasons, []);
  for (const [a, b] of [[900, 940], [2, 5], [1100, 1200], [3, 4]]) {
    await lazy.ensureRows(0, [{ r0: a, r1: b }]);
    assert.equal(norm(lazy.getRows(0, a, b, 0, 10)), norm(full.getRows(0, a, b, 0, 10)), `rows ${a}-${b}`);
  }
  const rows = lazy.getRows(0, 900, 901, 0, 10).rows;
  assert.match(JSON.stringify(rows), /A901\*2/);
});

test('lazy: reading a row that is not loaded fails (never shows it empty); rows are evicted and reloaded', async () => {
  const { lazy, full } = await bothModels(await bigBytes());
  assert.throws(() => lazy.getRows(0, 600, 610, 0, 10), /not loaded/);
  for (let a = 0; a < ROWS; a += 100) {
    await lazy.ensureRows(0, [{ r0: a, r1: a + 99 }]);
    assert.equal(norm(lazy.getRows(0, a, a + 100, 0, 10)), norm(full.getRows(0, a, a + 100, 0, 10)));
  }
  // keepRows is 150: rows not used for 32 requests are dropped, and must be asked for again
  for (let i = 0; i < 40; i++) await lazy.ensureRows(0, [{ r0: 1100 + i, r1: 1100 + i }]);
  assert.throws(() => lazy.getRows(0, 0, 10, 0, 10), /not loaded/);
  await lazy.ensureRows(0, [{ r0: 0, r1: 20 }]);
  assert.equal(norm(lazy.getRows(0, 0, 21, 0, 10)), norm(full.getRows(0, 0, 21, 0, 10)));
});

test('lazy: pinned rows stay loaded; a request for too many rows is refused with a text', async () => {
  const { lazy } = await bothModels(await bigBytes());
  await lazy.ensureRows(0, [{ r0: 10, r1: 12 }], true);
  for (let a = 100; a < ROWS; a += 100) await lazy.ensureRows(0, [{ r0: a, r1: a + 99 }]);
  assert.doesNotThrow(() => lazy.getRows(0, 10, 13, 0, 10));
  S.lazyTuning.maxRequestRows = 300;
  await assert.rejects(() => lazy.ensureRows(0, [{ r0: 0, r1: 800 }]), (e) => e instanceof S.LazyLimitError && e.message === S.LAZY_TOO_MANY_ROWS_TEXT);
});

test('lazy: an edit on a lazy sheet plans, applies, reads back and undoes like on the full model', async () => {
  const bytes = await bigBytes();
  const { lazy, full } = await bothModels(bytes);
  const id = lazy.getMeta().sheets[0].id;
  const intent = { kind: 'setCells', sheetId: id, cells: [[500, 1, 'edited <&>'], [500, 0, '42'], [501, 2, '=A2+1']] };
  await lazy.ensureRows(id, [{ r0: 500, r1: 501 }], true, true);
  const plan = S.planEdit(lazy, intent);
  assert.ok(plan.kind, JSON.stringify(plan));
  const fullPlan = S.planEdit(full, intent);
  assert.deepEqual(plan.after.ops, fullPlan.after.ops, 'same package ops as the full model');
  await lazy.ensureRows(0, [{ r0: 497, r1: 504 }]);
  const before = norm(lazy.getRows(0, 498, 504, 0, 10));
  S.applyEdit(lazy, plan);
  S.applyEdit(full, fullPlan);
  assert.equal(norm(lazy.getRows(0, 498, 504, 0, 10)), norm(full.getRows(0, 498, 504, 0, 10)));
  // other rows get evicted meanwhile; the edited ones stay
  for (let a = 600; a < ROWS; a += 100) await lazy.ensureRows(0, [{ r0: a, r1: a + 99 }]);
  S.invertEdit(lazy, plan);
  assert.equal(norm(lazy.getRows(0, 498, 504, 0, 10)), before, 'undo restores the display');
  // the file: built from the ORIGINAL bytes + the ops, and read back with the full loader
  const built = await S.buildXlsx(bytes, plan.after.ops);
  const back = S.workbookToModel(await S.loadWorkbook(built.bytes));
  assert.match(norm(back.getRows(0, 500, 502, 0, 10)), /edited <&>/);
});

test('lazy: an edit that reads a shared formula master loads its copies, and a huge group is refused', async () => {
  const { lazy } = await bothModels(await sharedBytes());
  const id = lazy.getMeta().sheets[0].id;
  await lazy.ensureRows(id, [{ r0: 1, r1: 1 }], true, true); // C2 = the master; its group reaches row 1200
  const plan = S.planEdit(lazy, { kind: 'setCells', sheetId: id, cells: [[1, 2, '7']] });
  assert.ok(plan.kind, JSON.stringify(plan));
  S.lazyTuning.maxRequestRows = 100; // a group that reaches more rows than a request may load
  await assert.rejects(() => lazy.ensureRows(id, [{ r0: 1, r1: 1 }], true, true), (e) => e instanceof S.LazyLimitError && e.message === S.LAZY_SHARED_TEXT);
});

test('lazy: fail closed - a sheet with relations or comments in its XML is not lazy', async () => {
  const withNote = await bigBytes((_wb, ws) => {
    ws.getCell('B3').note = 'a note';
  });
  const a = await bothModels(withNote);
  assert.equal(S.lazySheetCount(a.lazyWb), 0);
  assert.match(a.reasons.join('|'), /relations/);
  const zip = await JSZip.loadAsync(await bigBytes());
  const name = Object.keys(zip.files).find((n) => /worksheets\/sheet1\.xml$/.test(n));
  zip.file(name, (await zip.file(name).async('string')).replace('<row r="5"', '<!-- x --><row r="5"'));
  const b = await bothModels(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })));
  assert.equal(S.lazySheetCount(b.lazyWb), 0);
  assert.match(b.reasons.join('|'), /<!--|between its rows/);
});

test('size rules: lazy sheets are not counted in cells; the byte limit is XLSX_LAZY_EDIT_BYTES (25 MB) with a lazy sheet', async () => {
  const meta = { sheets: [{ rowCount: 100_000, colCount: 50 }, { rowCount: 10, colCount: 10 }] };
  assert.equal(S.xlsxTooLargeReason(meta, 10), S.XLSX_TOO_LARGE_TEXT);
  assert.equal(S.xlsxTooLargeReason(meta, 10, (i) => i === 0), undefined);
  assert.equal(S.xlsxTooLargeReason(meta, S.XLSX_LAZY_EDIT_BYTES, (i) => i === 0), undefined);
  assert.equal(S.xlsxTooLargeReason(meta, S.XLSX_LAZY_EDIT_BYTES + 1, (i) => i === 0), S.XLSX_LAZY_TOO_LARGE_TEXT);
  assert.equal(S.xlsxTooLargeReason({ sheets: [{ rowCount: 10, colCount: 10 }] }, 26 * 1024 * 1024), S.XLSX_TOO_LARGE_TEXT);
});

test('size rules (B1): the unpacked-bytes rule counts what applied edits add', () => {
  const big = 'x'.repeat(1_000_000);
  const plan = { after: { ops: [{ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: big } }]] }] } };
  assert.ok(S.opsGrowthBytes(plan.after.ops) >= 2_000_000);
  const limit = 25 * 1024 * 1024;
  assert.equal(S.editGrowsBytesTooLarge(limit - 3_000_000, 0, plan, false), undefined);
  assert.equal(S.editGrowsBytesTooLarge(limit - 3_000_000, 2_000_000, plan, false), S.XLSX_EDIT_BYTES_TOO_LARGE_TEXT, 'applied edits count');
  assert.equal(S.editGrowsBytesTooLarge(limit - 3_000_000, 2_000_000, plan, true), S.XLSX_EDIT_BYTES_TOO_LARGE_TEXT, 'a lazy workbook has the limit XLSX_LAZY_EDIT_BYTES');
  assert.equal(S.editGrowsBytesTooLarge(0, 0, plan, true), undefined);
});

test('lazy: fail closed - a big shared string table (every window would parse it again) loads the sheet as a whole', async () => {
  const a = await bothModels(await bigBytes(), { sheetBytes: 1000, chunkRows: 40, keepRows: 150, maxRequestRows: 2000, sstBytes: 10 });
  assert.equal(S.lazySheetCount(a.lazyWb), 0);
  assert.match(a.reasons.join('|'), /shared strings are large/);
});

test('lossy scan: a big worksheet part (unpacked with zlib, searched in slices) reports the same features as the small-part path', async () => {
  const zip = await JSZip.loadAsync(await bigBytes());
  const name = Object.keys(zip.files).find((n) => /worksheets\/sheet1\.xml$/.test(n));
  const xml = await zip.file(name).async('string');
  const ext = '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines/></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>';
  zip.file(name, xml.replace('</worksheet>', ext + '</worksheet>'));
  const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  const small = await S.detectLossyFeatures(bytes);
  S.lazyTuning.scanBytes = 100;
  const big = await S.detectLossyFeatures(bytes);
  assert.deepEqual(big, small);
  assert.ok(big.some((l) => /sparkline/i.test(l)), JSON.stringify(big));
});

// ----- fix round 1 -----

test('fix r1: editing another cell of a shared formula master row does not load (or refuse) the group; the master cell does', async () => {
  const { lazy } = await bothModels(await sharedBytes());
  const id = lazy.getMeta().sheets[0].id;
  S.lazyTuning.maxRequestRows = 100; // the group (C2:C1200) is longer than a request may load
  const cell = (r, c) => [{ r0: r, r1: r, c0: c, c1: c }];
  await lazy.ensureRows(id, [{ r0: 1, r1: 1 }], true, true, cell(1, 0)); // A2: not the formula cell
  await lazy.ensureRows(id, [{ r0: 1, r1: 1 }], true, true, cell(1, 1)); // B2
  await assert.rejects(() => lazy.ensureRows(id, [{ r0: 1, r1: 1 }], true, true, cell(1, 2)), (e) => e instanceof S.LazyLimitError && e.message === S.LAZY_SHARED_TEXT);
});

test('fix r1: a clear adds no bytes to the growth estimate', () => {
  const cells = [];
  for (let i = 0; i < 250_000; i++) cells.push([i, 0, { value: null }]);
  assert.equal(S.opsGrowthBytes([{ kind: 'cells', sheet: 1, cells }]), 0);
});

// ----- fix round 2 -----

test('fix r2: editing a shared formula copy works after its master row was evicted', async () => {
  const { lazy } = await bothModels(await sharedBytes((wb, ws) => { for (let r = ROWS + 1; r <= 2400; r++) ws.getCell('A' + r).value = r; }), { sheetBytes: 1000, chunkRows: 40, keepRows: 150, maxRequestRows: 2000, sstBytes: 2 * 1024 * 1024, scanBytes: 4 << 20 });
  const id = lazy.getMeta().sheets[0].id;
  await lazy.ensureRows(0, [{ r0: 0, r1: 5 }]); // the master row C2 is loaded as a normal row
  await lazy.ensureRows(0, [{ r0: 600, r1: 610 }]);
  for (let a = 1300; a < 2300; a += 10) {
    await lazy.ensureRows(0, [{ r0: a, r1: a + 9 }]);
    await lazy.ensureRows(0, [{ r0: 600, r1: 610 }]);
  }
  await lazy.ensureRows(id, [{ r0: 605, r1: 605 }], true, true, [{ r0: 605, r1: 605, c0: 2, c1: 2 }]);
  assert.doesNotThrow(() => S.planEdit(lazy, { kind: 'setCells', sheetId: id, cells: [[605, 2, '5']] }));
});

test('fix r2: a trailing row with only default row attributes does not extend the lazy used range (same as the full load)', async () => {
  const cases = [
    ['s="0" customFormat="0"', ''],
    ['hidden="false" customFormat="false" customHeight="false" outlineLevel="0"', ''],
    ['ht="15" customHeight="1"', ''],
    ['hidden="1"', ''],
  ];
  for (const [attrs] of cases) {
    const zip = await JSZip.loadAsync(await bigBytes());
    const name = Object.keys(zip.files).find((n) => /worksheets\/sheet1\.xml$/.test(n));
    const xml = (await zip.file(name).async('string')).replace('</sheetData>', `<row r="3000" ${attrs}/></sheetData>`);
    zip.file(name, xml);
    const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    const { lazy, full } = await bothModels(bytes);
    assert.equal(lazy.getMeta().sheets[0].rowCount, full.getMeta().sheets[0].rowCount, `rowCount for ${attrs}`);
  }
});

test('fix r2: the byte guard counts XML escapes and multi-byte text', () => {
  const growth = (text) => S.opsGrowthBytes([{ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: text } }]] }]);
  const n = 30000;
  assert.ok(growth('&'.repeat(n)) >= n * 5, 'ampersand is 5 bytes (&amp;)');
  assert.ok(growth('<'.repeat(n)) >= n * 4, 'less-than is 4 bytes (&lt;)');
  assert.ok(growth('日'.repeat(n)) >= n * 3, 'CJK is 3 bytes');
  assert.ok(growth('a'.repeat(n)) >= n, 'ascii');
});
