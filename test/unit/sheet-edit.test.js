// xlsx cell editing model (src/renderers/sheet.ts, EDITING + ADDRESSES `formula syntax` + WORKBOOK MODEL views, P2):
//   - sheet ids, IndexMap, sheet views (identity fast path), reading through the overlay
//   - parseCellInput (what a typed text means), checkFormulaInput (formula input, file syntax)
//   - planEdit / applyEdit / invertEdit for setCells and clear contents: identity of apply + invert on every sample,
//     blocks (arrays, merges, tables, protected sheets, host facts, spilling formulas), shared formulas, limits
//   - the edit-model fuzz against a naive model, with a mutation check (a broken inverse must be caught)
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const SAMPLES = path.join(__dirname, '..', 'samples');
const sample = (name) => new Uint8Array(fs.readFileSync(path.join(SAMPLES, name)));
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/** A fresh model of a sample (edit tests never share one: the model keeps the edits). */
async function sampleModel(name = 'styled.xlsx') {
  return S.workbookToModel(await S.loadWorkbook(sample(name)));
}

/** A workbook built with ExcelJS, written to bytes and loaded again like a file (so loader side channels run). */
async function builtModel(build) {
  const wb = new ExcelJS.Workbook();
  build(wb);
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  return { model: S.workbookToModel(await S.loadWorkbook(bytes)), bytes };
}

/**
 * All rows of a sheet (by position) as { 'r,c': cell }. A style index is only meaningful inside one model (the table is
 * filled as cells are first read), so cells carry the resolved style in `style` instead of `s`.
 */
function grid(model, sheet = 0) {
  const meta = model.getMeta().sheets[sheet];
  const out = new Map();
  const rows = model.getRows(sheet, 0, meta.rowCount + 50).rows;
  const styles = model.getMeta().styles;
  for (const row of rows) {
    for (const cell of row.cells) {
      const { s, ...rest } = cell;
      out.set(`${row.r},${cell.c}`, s ? { ...rest, style: styles[s] } : rest);
    }
  }
  return out;
}

const cellOf = (model, sheet, r, c) => grid(model, sheet).get(`${r},${c}`);
const setCells = (sheetId, ...cells) => ({ kind: 'setCells', sheetId, cells });
const clearRange = (sheetId, r0, c0, r1, c1) => ({ kind: 'clear', sheetId, ranges: [{ r0, c0, r1, c1 }], what: 'contents' });

/** plan -> apply, returns the edit; fails the test when the plan is not an Edit. */
function doEdit(model, intent) {
  const plan = S.planEdit(model, intent);
  assert.ok(plan && plan.kind, `expected an Edit, got ${JSON.stringify(plan)}`);
  S.applyEdit(model, plan);
  return plan;
}

// ----- sheet ids -----

test('sheet ids: "s" + the sheetId attribute, unique, stable', async () => {
  const model = await sampleModel();
  const { sheets } = model.getMeta();
  assert.deepEqual(sheets.map((s) => s.id), ['s1', 's2', 's3', 's4', 's5', 's6', 's7']);
  assert.equal(model.sheetIndexOf('s3'), 2);
  assert.equal(model.sheetIndexOf('s99'), -1);
  assert.equal(model.sheetIdAt(0), 's1');
  assert.equal(model.idProblem, undefined);
});

test('sheet ids follow the sheetId attribute, not the position (a swapped file)', async () => {
  const { bytes } = await builtModel((wb) => {
    wb.addWorksheet('First').getCell('A1').value = 'one';
    wb.addWorksheet('Second').getCell('A1').value = 'two';
  });
  const zip = await JSZip.loadAsync(bytes);
  let xml = await zip.file('xl/workbook.xml').async('string');
  xml = xml.replace(/sheetId="1"/, 'sheetId="X"').replace(/sheetId="2"/, 'sheetId="1"').replace(/sheetId="X"/, 'sheetId="2"');
  zip.file('xl/workbook.xml', xml);
  const model = S.workbookToModel(await S.loadWorkbook(new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))));
  const { sheets } = model.getMeta();
  assert.deepEqual(sheets.map((s) => [s.name, s.id]), [['First', 's2'], ['Second', 's1']]);
  // Renaming does not change the id; the id still finds the sheet.
  sheets[0].name = 'Renamed';
  assert.equal(model.sheetIndexOf('s2'), 0);
});

test('two sheets with the same sheetId: fresh ids, and every edit is refused', async () => {
  // ExcelJS keeps one sheet per id when it loads a file, so the situation is built in memory.
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('A').getCell('A1').value = 1;
  wb.addWorksheet('B').getCell('A1').value = 2;
  wb.worksheets[1].id = wb.worksheets[0].id;
  const model = S.workbookToModel(wb);
  const ids = model.getMeta().sheets.map((s) => s.id);
  assert.equal(new Set(ids).size, 2, 'ids are unique');
  assert.match(model.idProblem, /same internal id/);
  const plan = S.planEdit(model, setCells(ids[0], [0, 0, 'x']));
  assert.equal(plan.blocked?.[0].code, 'unsupported');
});

// ----- IndexMap -----

test('IndexMap: identity, splice, undo gives the identity piece back', () => {
  const m = new S.IndexMap(10);
  assert.equal(m.identity, true);
  assert.equal(m.get(7), 7);
  const removed = m.splice(3, 2, []); // delete positions 3..4
  assert.equal(m.len, 8);
  assert.equal(m.identity, false);
  assert.deepEqual([m.get(2), m.get(3), m.get(7)], [2, 5, 9]);
  m.splice(3, 0, removed); // undo
  assert.equal(m.len, 10);
  assert.equal(m.identity, true, 'one piece again');
  assert.deepEqual(m.pieces, [{ s: 0, n: 10 }]);
  const ins = m.splice(5, 0, [{ s: 100, n: 3 }]); // insert 3 new ids
  assert.deepEqual(ins, []);
  assert.deepEqual([m.get(4), m.get(5), m.get(7), m.get(8)], [4, 100, 102, 5]);
  assert.throws(() => m.get(13), RangeError);
});

test('IndexMap fuzz against an array (splice, ids, get)', () => {
  for (const seed of [1, 2, 3]) {
    let x = seed * 7919;
    const rnd = (n) => ((x = (x * 1103515245 + 12345) & 0x7fffffff) % n);
    const n = 60;
    const m = new S.IndexMap(n);
    let naive = Array.from({ length: n }, (_, i) => i);
    let next = 1000;
    for (let step = 0; step < 400; step++) {
      const at = rnd(naive.length + 1);
      const count = rnd(Math.min(5, naive.length - at) + 1);
      const ins = rnd(4);
      const insert = ins ? (rnd(2) ? [{ s: next, n: ins }] : [{ ids: Int32Array.from({ length: ins }, (_, i) => next + i * 2) }]) : [];
      const expect = insert.length ? (insert[0].ids ? [...insert[0].ids] : Array.from({ length: ins }, (_, i) => next + i)) : [];
      next += 10;
      const removed = m.splice(at, count, insert);
      assert.equal(removed.reduce((a, p) => a + (p.ids ? p.ids.length : p.n), 0), count);
      naive.splice(at, count, ...expect);
      assert.equal(m.len, naive.length);
      for (let p = 0; p < naive.length; p++) assert.equal(m.get(p), naive[p], `seed ${seed} step ${step} pos ${p}`);
    }
  }
});

// ----- parseCellInput -----

test('parseCellInput: numbers, percent, currency, grouping', () => {
  const num = (t, fmt) => S.parseCellInput(t, { numFmt: fmt });
  assert.deepEqual(num('42'), { kind: 'number', value: 42 });
  assert.deepEqual(num('-3.5'), { kind: 'number', value: -3.5 });
  assert.deepEqual(num('+7'), { kind: 'number', value: 7 });
  assert.deepEqual(num('(12)'), { kind: 'number', value: -12 });
  assert.deepEqual(num('.5'), { kind: 'number', value: 0.5 });
  assert.deepEqual(num('5.'), { kind: 'number', value: 5 });
  assert.deepEqual(num('1e3'), { kind: 'number', value: 1000 });
  assert.deepEqual(num('2.5E-2'), { kind: 'number', value: 0.025 });
  assert.deepEqual(num('  12  '), { kind: 'number', value: 12 });
  assert.deepEqual(num('007'), { kind: 'number', value: 7 }, 'like Excel: leading zeros go');
  assert.deepEqual(num('50%'), { kind: 'number', value: 0.5, numFmt: '0%' });
  assert.deepEqual(num('5.5%'), { kind: 'number', value: 0.055, numFmt: '0.00%' });
  assert.deepEqual(num('-12.25%'), { kind: 'number', value: -0.1225, numFmt: '0.00%' });
  assert.deepEqual(num('$5'), { kind: 'number', value: 5, numFmt: '"$"#,##0' });
  assert.deepEqual(num('$1,234.50'), { kind: 'number', value: 1234.5, numFmt: '"$"#,##0.00' });
  assert.deepEqual(num('-$5'), { kind: 'number', value: -5, numFmt: '"$"#,##0' });
  assert.deepEqual(num('$-5'), { kind: 'number', value: -5, numFmt: '"$"#,##0' });
  assert.deepEqual(num('($5.00)'), { kind: 'number', value: -5, numFmt: '"$"#,##0.00' });
  assert.deepEqual(num('1,234'), { kind: 'number', value: 1234, numFmt: '#,##0' });
  assert.deepEqual(num('1,234,567.8'), { kind: 'number', value: 1234567.8, numFmt: '#,##0.00' });
  // not numbers
  for (const t of ['1,23', '1,2345', '12abc', '1.2.3', '--5', '$', '%', 'Infinity', 'NaN', '0x10', '1e', '5 5', '1/2', '(5', '5)', '-(5)']) {
    assert.equal(S.parseCellInput(t).kind, 'text', t);
  }
  // 16 or more significant digits stay text (Excel would round them silently); trailing zeros are not significant
  assert.equal(S.parseCellInput('1234567890123456').kind, 'text');
  assert.equal(S.parseCellInput('0.1234567890123456').kind, 'text');
  assert.equal(S.parseCellInput('123456789012345').kind, 'number');
  assert.equal(S.parseCellInput('1000000000000000').kind, 'number');
  assert.equal(Object.is(S.parseCellInput('-0').value, 0), true, 'no negative zero');
});

test('parseCellInput: dates and times (1900 and 1904 systems)', () => {
  const when = (t, date1904 = false, numFmt) => S.parseCellInput(t, { date1904, numFmt });
  assert.deepEqual(when('2024-03-05'), { kind: 'number', value: 45356, numFmt: 'm/d/yyyy' });
  assert.deepEqual(when('3/5/2024'), { kind: 'number', value: 45356, numFmt: 'm/d/yyyy' });
  assert.deepEqual(when('03/05/24'), { kind: 'number', value: 45356, numFmt: 'm/d/yyyy' }, 'two-digit year: 00-29 -> 2000s');
  assert.equal(when('1/1/99').value, 36161, '30-99 -> 1900s');
  assert.equal(when('1900-01-01').value, 1);
  assert.equal(when('1900-02-28').value, 59);
  assert.equal(when('1900-02-29').value, 60, "Excel's day 60");
  assert.equal(when('1900-03-01').value, 61);
  assert.equal(when('9999-12-31').value, 2958465);
  assert.equal(when('1904-01-01', true).value, 0);
  assert.equal(when('2024-03-05', true).value, 45356 - 1462);
  assert.equal(when('1903-12-31', true).kind, 'text', 'before 1904 in the 1904 system');
  for (const t of ['2023-02-29', '2024-13-01', '2024-00-10', '1899-12-31', '10000-01-01', '2024-1-32', '13/1/2024', '1/1/20245']) {
    assert.equal(when(t).kind, 'text', t);
  }
  assert.deepEqual(when('12:30'), { kind: 'number', value: 0.5208333333333334, numFmt: 'h:mm' });
  assert.deepEqual(when('14:30:15'), { kind: 'number', value: (14 * 3600 + 30 * 60 + 15) / 86400, numFmt: 'h:mm:ss' });
  assert.deepEqual(when('1:30 PM'), { kind: 'number', value: 13.5 / 24, numFmt: 'h:mm AM/PM' });
  assert.deepEqual(when('12:00 AM'), { kind: 'number', value: 0, numFmt: 'h:mm AM/PM' });
  assert.deepEqual(when('2024-03-05 12:00'), { kind: 'number', value: 45356.5, numFmt: 'm/d/yyyy h:mm' });
  assert.deepEqual(when('3/5/2024 6:00 PM'), { kind: 'number', value: 45356.75, numFmt: 'm/d/yyyy h:mm' });
  assert.deepEqual(when('2024-03-05T06:00'), { kind: 'number', value: 45356.25, numFmt: 'm/d/yyyy h:mm' });
  for (const t of ['24:00', '12:60', '13:00 PM', '0:30 AM', '12:30:61', '1:2']) assert.equal(when(t).kind, 'text', t);
});

test('parseCellInput: booleans, errors, apostrophe, text format, formulas, empty', () => {
  const p = (t, numFmt) => S.parseCellInput(t, { numFmt });
  assert.deepEqual(p('true'), { kind: 'boolean', value: true });
  assert.deepEqual(p(' FALSE '), { kind: 'boolean', value: false });
  assert.deepEqual(p('#n/a'), { kind: 'error', value: '#N/A' });
  assert.deepEqual(p('#DIV/0!'), { kind: 'error', value: '#DIV/0!' });
  assert.equal(p('#SPILL!').kind, 'text', 'only the classic error literals');
  assert.deepEqual(p("'007"), { kind: 'text', text: '007', quotePrefix: true });
  assert.deepEqual(p("'=1+1"), { kind: 'text', text: '=1+1', quotePrefix: true });
  assert.deepEqual(p("''x"), { kind: 'text', text: "'x", quotePrefix: true });
  assert.deepEqual(p("'"), { kind: 'empty' });
  assert.deepEqual(p(''), { kind: 'empty' });
  assert.deepEqual(p('007', '@'), { kind: 'text', text: '007', quotePrefix: false }, 'text format: nothing is converted');
  assert.deepEqual(p('=1+1', '@'), { kind: 'text', text: '=1+1', quotePrefix: false });
  assert.deepEqual(p("'x", '@'), { kind: 'text', text: "'x", quotePrefix: false }, 'in a text cell the apostrophe is part of the text');
  assert.deepEqual(p('=SUM(A1:A3)'), { kind: 'formula', body: 'SUM(A1:A3)' });
  assert.equal(p('=').kind, 'text', 'a lone = is text');
  assert.deepEqual(p('   '), { kind: 'text', text: '   ', quotePrefix: false });
  assert.deepEqual(p(' hello '), { kind: 'text', text: ' hello ', quotePrefix: false }, 'text keeps its spaces');
  assert.equal(p('hello').kind, 'text');
  // a number typed into a formatted cell keeps its format (no preset applied by parseCellInput: the planner decides)
  assert.deepEqual(p('5%', '0.00'), { kind: 'number', value: 0.05, numFmt: '0%' });
});

// ----- formula input -----

test('checkFormulaInput: file syntax gets the _xlfn. prefixes, the editor text does not', () => {
  const ok = (body, names = ['Sheet1']) => {
    const r = S.checkFormulaInput(body, names);
    assert.equal(r.ok, true, `${body}: ${JSON.stringify(r)}`);
    return r;
  };
  assert.deepEqual(ok('SUM(A1:A3)'), { ok: true, file: 'SUM(A1:A3)', editor: 'SUM(A1:A3)' });
  const concat = ok('concat(A1,B1)&TEXTJOIN(",",TRUE,C1:C3)');
  assert.equal(concat.file, '_xlfn.CONCAT(A1,B1)&_xlfn.TEXTJOIN(",",TRUE,C1:C3)');
  assert.equal(concat.editor, 'CONCAT(A1,B1)&TEXTJOIN(",",TRUE,C1:C3)');
  assert.equal(ok('IFS(A1>1,"a",TRUE,"b")').file, '_xlfn.IFS(A1>1,"a",TRUE,"b")');
  assert.equal(ok('XLOOKUP(A1,B:B,C:C)').file, '_xlfn.XLOOKUP(A1,B:B,C:C)');
  assert.equal(ok('STDEV.S(A1:A9)').file, '_xlfn.STDEV.S(A1:A9)');
  assert.equal(ok('DAYS(A1,B1)').file, '_xlfn.DAYS(A1,B1)');
  assert.equal(ok('MAXIFS(A:A,B:B,1)').file, '_xlfn.MAXIFS(A:A,B:B,1)');
  assert.equal(ok('IFERROR(A1,0)').file, 'IFERROR(A1,0)', 'old functions stay plain');
  assert.equal(ok('_xlfn.CONCAT(A1)').file, '_xlfn.CONCAT(A1)', 'a typed prefix is kept, not doubled');
  assert.equal(ok('sum(a1:a3)').file, 'SUM(a1:a3)', 'function names go upper case, references keep their spelling');
  assert.equal(ok("'Sheet1'!A1+Sheet1!B2").file, "'Sheet1'!A1+Sheet1!B2");
  assert.equal(ok('foo(1)').file, 'FOO(1)', 'unknown functions are allowed (#NAME? in Excel)');
  assert.equal(ok('MyName*2').file, 'MyName*2', 'unknown names are allowed');
  assert.equal(ok('"a ""quoted"" text"&A1').file, '"a ""quoted"" text"&A1');
  assert.equal(ok('1+\n2').file, '1+\n2', 'a line break inside a formula is kept');
  assert.equal(ok('{1,2;3,4}').file, '{1,2;3,4}');
  assert.equal(ok('SUM(Sheet1:Sheet1!A1)').ok, true);
  // round trip: the file text of a cell shows without prefixes and goes back to the same file text
  const back = ok(concat.editor);
  assert.equal(back.file, concat.file);
});

test('checkFormulaInput: invalid formulas are refused with the standard message', () => {
  for (const body of [
    '1+', 'SUM(A1', '1+2)', '"abc', '(1+2', 'A1:', 'SUM(A1;A2)', '', '   ', '=A1', '@A1', 'SUM(A1:XFE5)', 'A1048577+1', 'XFE1', 'ZZZ99',
    "'No Such Sheet'!A1", 'Nope!A1', 'Sheet1:Nope!A1', 'A1 +* 2', 'IF(,', '{1,2', '1E999', 'a\u0001b', 'a\ud800b', '#', '1 2 3 +',
  ]) {
    const r = S.checkFormulaInput(body, ['Sheet1']);
    assert.deepEqual(r, { ok: false, reason: 'invalid', message: "There's a problem with this formula." }, JSON.stringify(body));
  }
  // limits
  assert.equal(S.checkFormulaInput('1+'.repeat(5000) + '1', ['S']).ok, false, 'longer than 8,192 characters');
  assert.equal(S.checkFormulaInput('('.repeat(65) + '1' + ')'.repeat(65), ['S']).ok, false, 'nested 65 levels');
  assert.equal(S.checkFormulaInput('('.repeat(64) + '1' + ')'.repeat(64), ['S']).ok, true, 'nested 64 levels');
  assert.equal(S.checkFormulaInput(`SUM(${Array(256).fill('1').join(',')})`, ['S']).ok, false, '256 arguments');
  assert.equal(S.checkFormulaInput(`SUM(${Array(255).fill('1').join(',')})`, ['S']).ok, true, '255 arguments');
  assert.equal(S.checkFormulaInput(`"${'x'.repeat(256)}"`, ['S']).ok, false, 'text constant of 256 characters');
  assert.equal(S.checkFormulaInput(`"${'x'.repeat(255)}"`, ['S']).ok, true);
  assert.equal(S.checkFormulaInput(`SUM({1,2,3,4},1)`, ['S']).ok, true, 'commas of an array constant are not arguments');
});

test('checkFormulaInput: what FileStudio cannot write yet is blocked, with a message that says so', () => {
  const blocked = (body) => {
    const r = S.checkFormulaInput(body, ['Sheet1']);
    assert.equal(r.ok, false, body);
    assert.equal(r.reason, 'blocked', `${body}: ${JSON.stringify(r)}`);
    assert.equal(r.block.code, 'unsupported');
    return r.block.message;
  };
  for (const f of ['SEQUENCE(3)', 'filter(A1:A9,B1:B9>1)', 'SORT(A1:A9)', 'UNIQUE(A:A)', 'TRANSPOSE(A1:C1)', 'TEXTSPLIT(A1,",")', 'VSTACK(A1:A2,B1:B2)', 'A1#', 'SUM(A1#)']) {
    assert.match(blocked(f), /spilling \(dynamic array\) formulas/, f);
  }
  assert.match(blocked('LET(x,1,x+1)'), /LET or LAMBDA/);
  assert.match(blocked('LAMBDA(x,x)(1)'), /LET or LAMBDA/);
  assert.match(blocked('[1]Sheet1!A1+1'), /other workbooks/);
  assert.match(blocked("'[1]Sheet one'!A1"), /other workbooks/);
  assert.match(blocked('Table1[Col]'), /table references/);
  assert.match(blocked('SUM(Table1[[#This Row],[Col]])'), /table references/);
});

// ----- reading through views -----

test('a sheet with an overlay reads exactly like the plain loops (general path == identity fast path)', async () => {
  const pristine = await sampleModel();
  const model = await sampleModel();
  const id = model.getMeta().sheets[0].id;
  // an edit far from the data, then undo not applied: the view is no longer plain
  doEdit(model, setCells(id, [400, 20, 'far away']));
  assert.equal(model.viewOf(id).over.size, 1);
  for (let sheet = 0; sheet < pristine.getMeta().sheets.length; sheet++) {
    const meta = pristine.getMeta().sheets[sheet];
    const a = pristine.getRows(sheet, 0, meta.rowCount + 5);
    const b = model.getRows(sheet, 0, meta.rowCount + 5);
    if (sheet === 0) b.rows = b.rows.filter((r) => r.r !== 400);
    assert.deepEqual(b.rows, a.rows, `sheet ${sheet} rows`);
    const ranges = [{ r0: 0, c0: 0, r1: meta.rowCount, c1: meta.colCount }];
    assert.deepEqual(model.getStats(sheet, ranges), pristine.getStats(sheet, ranges), `sheet ${sheet} stats`);
  }
  // column windows too
  assert.deepEqual(model.getRows(0, 3, 12, 2, 5).rows, pristine.getRows(0, 3, 12, 2, 5).rows);
});

test('reading through moved maps: rows and columns come from the ids (the base P4 builds on)', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[1].id; // Numbers
  const pristine = await sampleModel();
  const before = grid(pristine, 1);
  const view = model.viewOf(id);
  // content rows 2..4 reversed (a sort), column 1 deleted
  const ids = view.contentMap.ids(2, 4);
  view.contentMap.splice(2, 3, [{ ids: Int32Array.from([ids[2], ids[1], ids[0]]) }]);
  view.colMap.splice(1, 1, []);
  const after = grid(model, 1);
  // the cell at old (r, c) is now at (newPos(r), c - 1 when c > 1); column 1 is gone
  const newPos = (r) => (r >= 2 && r <= 4 ? 6 - r : r);
  let compared = 0;
  for (const [key, cell] of before) {
    const [r, c] = key.split(',').map(Number);
    if (c === 1) continue;
    const found = after.get(`${newPos(r)},${c > 1 ? c - 1 : c}`);
    assert.ok(found, `cell ${key} is somewhere`);
    assert.deepEqual({ ...found, c: cell.c }, cell, `cell ${key}`);
    compared++;
  }
  assert.ok(compared > 100);
  assert.equal(after.size, [...before.keys()].filter((k) => k.split(',')[1] !== '1').length, 'no cell appeared or vanished');
  assert.equal(model.getMeta().sheets[1].id, id);
});

// ----- apply / invert -----

test('setCells: values, text, formulas, presets show up in getRows; undo restores everything', async () => {
  const model = await sampleModel();
  const pristine = await sampleModel();
  const id = model.getMeta().sheets[1].id;
  // rows 100+ of the Numbers sheet are empty General cells
  const edit = doEdit(model, setCells(id, [100, 0, 'hello'], [101, 1, '42'], [102, 2, '=1+2'], [103, 3, '50%'], [104, 4, 'TRUE'], [105, 5, '#N/A'], [106, 0, '2024-03-05'], [107, 0, "'007"]));
  const g = grid(model, 1);
  assert.deepEqual([g.get('100,0').t, g.get('100,0').v], ['s', 'hello']);
  assert.deepEqual([g.get('101,1').t, g.get('101,1').v, g.get('101,1').w], ['n', 42, '42']);
  assert.equal(g.get('102,2').f, '1+2');
  assert.equal(g.get('102,2').calc, 'stale');
  assert.deepEqual([g.get('103,3').v, g.get('103,3').w], [0.5, '50%']);
  assert.deepEqual([g.get('104,4').t, g.get('104,4').v, g.get('104,4').w], ['b', true, 'TRUE']);
  assert.deepEqual([g.get('105,5').t, g.get('105,5').w], ['e', '#N/A']);
  assert.deepEqual([g.get('106,0').t, g.get('106,0').w], ['d', '3/5/2024']);
  assert.deepEqual([g.get('107,0').t, g.get('107,0').v], ['s', '007']);
  assert.equal(edit.label, 'Edit 8 cells');
  assert.deepEqual(edit.range, { r0: 100, c0: 0, r1: 107, c1: 5 });
  assert.equal(edit.after.ops.length, 1);
  assert.equal(edit.after.ops[0].kind, 'cells');
  assert.equal(edit.after.ops[0].sheet, 2);
  S.invertEdit(model, edit);
  assert.deepEqual(grid(model, 1), grid(pristine, 1));
  assert.equal(model.viewOf(id).over.size, 0);
  S.applyEdit(model, edit); // redo
  assert.equal(grid(model, 1).get('101,1').v, 42);
  S.invertEdit(model, edit);
  assert.deepEqual(grid(model, 1), grid(pristine, 1));
});

test('file operations: positions, values, presets and the cached value of formulas', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[0].id;
  const edit = S.planEdit(model, setCells(id, [30, 7, '5%'], [31, 7, 'a\nb'], [32, 7, "'12"], [33, 7, '=CONCAT("a","b")'], [34, 7, '']));
  const op = edit.after.ops[edit.after.ops.length - 1];
  const patches = Object.fromEntries(op.cells.map(([r, c, p]) => [`${r},${c}`, p]));
  assert.deepEqual(patches['30,7'], { value: { t: 'n', v: 0.05 }, style: [{ numFmt: '0%' }] });
  assert.deepEqual(patches['31,7'], { value: { t: 's', v: 'a\nb' }, style: [{ wrap: true }] });
  assert.deepEqual(patches['32,7'], { value: { t: 's', v: '12' }, style: [{ quotePrefix: true }] });
  assert.deepEqual(patches['33,7'], { formula: '_xlfn.CONCAT("a","b")', cached: null });
  assert.equal(patches['34,7'], undefined, 'clearing an empty cell is no change');
  // positions are sorted row by row
  assert.deepEqual(op.cells.map(([r]) => r), [30, 31, 32, 33]);
});

test('unchanged input is no edit; a changed preset format is one', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[1].id;
  const g = grid(model, 1);
  const textCell = [...g.entries()].find(([, c]) => c.t === 's' && c.v && !c.f);
  const [r, c] = textCell[0].split(',').map(Number);
  assert.deepEqual(S.planEdit(model, setCells(id, [r, c, textCell[1].v])), { noop: true });
  const numCell = [...g.entries()].find(([, c]) => c.t === 'n' && c.w === String(c.v) && !c.f);
  if (numCell) {
    const [nr, nc] = numCell[0].split(',').map(Number);
    assert.deepEqual(S.planEdit(model, setCells(id, [nr, nc, String(numCell[1].v)])), { noop: true });
  }
  assert.deepEqual(S.planEdit(model, setCells(id, [200, 3, ''])), { noop: true }, 'clearing an empty cell');
  assert.deepEqual(S.planEdit(model, setCells(id)), { noop: true }, 'no cells');
  assert.deepEqual(S.planEdit(model, clearRange(id, 100, 0, 150, 10)), { noop: true });
});

test('apply + invert is the identity on every sheet of every sample workbook', async () => {
  for (const name of ['styled.xlsx']) {
    const pristine = await sampleModel(name);
    const model = await sampleModel(name);
    const sheets = model.getMeta().sheets;
    for (let s = 0; s < sheets.length; s++) {
      const id = sheets[s].id;
      const before = grid(model, s);
      const stats = model.getStats(s, [{ r0: 0, c0: 0, r1: 200, c1: 40 }]);
      const edits = [];
      edits.push(doEdit(model, setCells(id, [0, 0, 'x'], [1, 1, '3.5'], [2, 2, '=A1'], [sheets[s].rowCount + 3, sheets[s].colCount + 2, 'beyond'])));
      const clear = S.planEdit(model, clearRange(id, 0, 0, Math.min(15, sheets[s].rowCount), Math.min(8, sheets[s].colCount)));
      if (clear.kind) {
        S.applyEdit(model, clear);
        edits.push(clear);
      }
      assert.notDeepEqual(grid(model, s), before, `sheet ${s}: the edits changed something`);
      for (const e of edits.reverse()) S.invertEdit(model, e);
      assert.deepEqual(grid(model, s), before, `sheet ${s}: grid after undo`);
      assert.deepEqual(model.getStats(s, [{ r0: 0, c0: 0, r1: 200, c1: 40 }]), stats, `sheet ${s}: stats after undo`);
      assert.equal(model.viewOf(id).over.size, 0, `sheet ${s}: no overlay entries left`);
      assert.deepEqual(model.getMeta().sheets[s].rowCount, pristine.getMeta().sheets[s].rowCount);
      assert.deepEqual(model.getMeta().sheets[s].colCount, pristine.getMeta().sheets[s].colCount);
    }
  }
});

test('an edit beyond the used range grows it (meta change reported) and undo shrinks it back', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[3].id; // Hidden, 9 x 1
  const rows = model.getMeta().sheets[3].rowCount;
  const e = S.planEdit(model, setCells(id, [rows + 10, 4, 'far']));
  const change = S.applyEdit(model, e);
  assert.equal(change.meta, true);
  assert.deepEqual(change.rows, [rows + 10, rows + 11]);
  assert.deepEqual([model.sheetMeta(id).rowCount, model.sheetMeta(id).colCount], [rows + 11, 5]);
  assert.equal(cellOf(model, 3, rows + 10, 4).v, 'far');
  const back = S.invertEdit(model, e);
  assert.equal(back.meta, true);
  assert.deepEqual([model.sheetMeta(id).rowCount, model.sheetMeta(id).colCount], [rows, 1]);
  // an edit inside the used range changes no meta
  assert.equal(S.applyEdit(model, S.planEdit(model, setCells(id, [0, 0, 'in']))).meta, false);
  // refreshSheet recomputes the same meta
  assert.deepEqual(model.refreshSheet(id).rowCount, rows);
});

test('clear contents: values and formulas go, styles stay; notes and links of other cells are untouched', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[0].id;
  const before = grid(model, 0);
  const filled = [...before.entries()].filter(([, c]) => c.t !== 'z' && c.t !== undefined).slice(0, 5);
  assert.ok(filled.length >= 3);
  const [r, c] = filled[0][0].split(',').map(Number);
  const edit = doEdit(model, clearRange(id, r, c, r, c));
  const after = grid(model, 0).get(`${r},${c}`);
  assert.ok(!after || after.t === 'z', 'content is gone');
  assert.deepEqual(after?.style, before.get(`${r},${c}`).style, 'style is kept');
  assert.equal(edit.label, `Clear ${String.fromCharCode(65 + c)}${r + 1}`);
  assert.deepEqual(edit.after.ops[0].cells, [[r, c, { value: null }]]);
  S.invertEdit(model, edit);
  assert.deepEqual(grid(model, 0), before);
  assert.deepEqual(S.planEdit(model, { kind: 'clear', sheetId: id, ranges: [{ r0: 0, c0: 0, r1: 1, c1: 1 }], what: 'formats' }).blocked[0].message, 'This action is not available yet.');
});

// ----- blocks -----

test('blocks: array formulas, merged cells, tables, protected sheets, host facts', async () => {
  const { model } = await builtModel((wb) => {
    const ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 1;
    ws.getCell('B2').value = { formula: 'A1*2', shareType: 'array', ref: 'B2:B4', result: 2 };
    ws.mergeCells('D2:E3');
    ws.getCell('D2').value = 'merged';
    ws.addTable({ name: 'Sales', ref: 'G1', headerRow: true, totalsRow: true, columns: [{ name: 'Item', totalsRowLabel: 'Total' }, { name: 'Qty', totalsRowFunction: 'sum' }], rows: [['a', 1], ['b', 2]] });
    const locked = wb.addWorksheet('Locked');
    locked.getCell('A1').value = 'locked';
    locked.getCell('B1').value = 'free';
    locked.getCell('B1').protection = { locked: false };
    locked.protect('', {});
  });
  const id = 's1';
  // array range (multi-cell array B2:B4; its other cells hold cached values)
  let r = S.planEdit(model, setCells(id, [2, 1, 'x']));
  assert.equal(r.blocked[0].code, 'array');
  assert.match(r.blocked[0].message, /array formula \(B2:B4\)/);
  assert.deepEqual(r.blocked[0].where, ['Data!B3']);
  r = S.planEdit(model, clearRange(id, 1, 1, 1, 1));
  assert.equal(r.blocked[0].code, 'array');
  // merged: a single cell goes to the anchor; in a block empty text is dropped and a value blocks
  const single = S.planEdit(model, setCells(id, [2, 4, 'to anchor'])); // E3 is part of D2:E3
  assert.deepEqual(single.after.ops[0].cells.map(([rr, cc]) => [rr, cc]), [[1, 3]], 'single cell edit goes to D2');
  const block = S.planEdit(model, setCells(id, [1, 3, 'a'], [1, 4, ''], [2, 3, ''], [2, 4, '']));
  assert.deepEqual(block.after.ops[0].cells.map(([rr, cc]) => [rr, cc]), [[1, 3]], 'empty text of the other merged cells is dropped');
  r = S.planEdit(model, setCells(id, [1, 3, 'a'], [1, 4, 'b']));
  assert.equal(r.blocked[0].code, 'merge');
  assert.match(r.blocked[0].message, /D2:E3 is a merged cell/);
  // table header and totals row
  r = S.planEdit(model, setCells(id, [0, 6, 'Renamed']));
  assert.equal(r.blocked[0].code, 'table');
  assert.match(r.blocked[0].message, /header row of table “Sales”/);
  r = S.planEdit(model, setCells(id, [3, 7, '5']));
  assert.match(r.blocked[0].message, /totals row of table “Sales”/);
  assert.ok(S.planEdit(model, setCells(id, [1, 6, 'body cell'])).kind, 'a body cell of the table is fine');
  // protected sheet: locked cells are refused, unlocked ones are fine
  r = S.planEdit(model, setCells('s2', [0, 0, 'x']));
  assert.equal(r.blocked[0].code, 'protected');
  assert.equal(r.blocked[0].message, 'The sheet is protected in Excel.');
  assert.ok(S.planEdit(model, setCells('s2', [0, 1, 'x'])).kind, 'the unlocked cell can be edited');
  // facts only the host knows
  model.extraCellBlock = (sheet, rr, cc) => (sheet === 's1' && rr === 0 && cc === 0 ? 'This cell holds a picture or linked data type. Edit it in Excel.' : undefined);
  r = S.planEdit(model, setCells(id, [0, 0, 'x']));
  assert.equal(r.blocked[0].message, 'This cell holds a picture or linked data type. Edit it in Excel.');
  assert.ok(S.planEdit(model, setCells(id, [0, 2, 'x'])).kind);
  // a blocked edit changes nothing
  assert.equal(model.viewOf(id).over.size, 0);
  // blocks are collected for the whole edit: all offending cells named, once per code
  r = S.planEdit(model, setCells(id, [1, 1, 'a'], [2, 1, 'b'], [3, 1, 'c'], [0, 0, 'd']));
  assert.deepEqual(r.blocked.map((b) => b.code).sort(), ['array', 'unsupported']);
  assert.deepEqual(r.blocked.find((b) => b.code === 'array').where, ['Data!B2', 'Data!B3', 'Data!B4']);
});

test('invalid input: unknown sheet, bad shapes, limits, lone surrogates, formulas', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[0].id;
  assert.deepEqual(S.planEdit(model, setCells('s99', [0, 0, 'x'])), { invalid: 'This sheet is no longer in the workbook.' });
  for (const bad of [[[-1, 0, 'x']], [[0, MAX_COLS, 'x']], [[MAX_ROWS, 0, 'x']], [[0.5, 0, 'x']], [[0, 0, 5]], [[0, 0]], ['x']]) {
    assert.deepEqual(S.planEdit(model, setCells(id, ...bad)), { invalid: 'This change is not valid.' }, JSON.stringify(bad));
  }
  assert.deepEqual(S.planEdit(model, setCells(id, [0, 0, 'a\ud800b'])), { invalid: 'The text contains a character Excel cannot store.' });
  assert.match(S.planEdit(model, setCells(id, [0, 0, 'x'.repeat(32768)])).invalid, /too long for a cell \(max 32,767 characters\)/);
  assert.ok(S.planEdit(model, setCells(id, [0, 0, 'x'.repeat(32767)])).kind);
  const many = Array.from({ length: 1_000_001 }, (_, i) => [i % 1000, Math.floor(i / 1000) % 16000, 'x']);
  assert.deepEqual(S.planEdit(model, { kind: 'setCells', sheetId: id, cells: many }), { invalid: 'This change is too large (max 1,000,000 cells).' });
  assert.deepEqual(S.planEdit(model, setCells(id, [0, 0, '=1+'])), { invalid: "There's a problem with this formula." });
  assert.deepEqual(S.planEdit(model, setCells(id, [0, 0, '=Nope!A1'])), { invalid: "There's a problem with this formula." });
  assert.match(S.planEdit(model, setCells(id, [0, 0, '=SEQUENCE(3)'])).blocked[0].message, /spilling/);
  // kinds P2 does not do yet
  assert.equal(S.planEdit(model, { kind: 'insertRows', sheetId: id, at: 0, count: 1 }).blocked[0].message, 'This action is not available yet.');
  assert.equal(S.planEdit(model, { kind: 'mdText', edits: [] }).blocked[0].code, 'unsupported');
  // CRLF from a paste is stored as LF
  const e = S.planEdit(model, setCells(id, [0, 0, 'a\r\nb\rc']));
  assert.equal(e.after.ops[0].cells[0][2].value.v, 'a\nb\nc');
});

// ----- shared formulas, hyperlinks, rich text -----

/** A workbook whose B1:B3 are one shared formula (master B1) over A1:A3, built by rewriting the sheet XML. */
async function sharedModel() {
  const { bytes } = await builtModel((wb) => {
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 3; r++) {
      ws.getCell(`A${r}`).value = r * 10;
      ws.getCell(`B${r}`).value = { formula: `A${r}*2`, result: r * 20 };
    }
    ws.getCell('C1').value = { text: 'link', hyperlink: 'https://example.com/' };
    ws.getCell('C2').value = { richText: [{ text: 'Rich ', font: { bold: true } }, { text: 'text' }] };
  });
  const zip = await JSZip.loadAsync(bytes);
  let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  xml = xml.replace(/<f>A1\*2<\/f>/, '<f t="shared" ref="B1:B3" si="0">A1*2</f>').replace(/<f>A2\*2<\/f>/, '<f t="shared" si="0"/>').replace(/<f>A3\*2<\/f>/, '<f t="shared" si="0"/>');
  assert.match(xml, /t="shared" ref="B1:B3"/);
  zip.file('xl/worksheets/sheet1.xml', xml);
  return S.workbookToModel(await S.loadWorkbook(new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))));
}

test('shared formula: the clones show their formulas; a change of the master separates the group first', async () => {
  const model = await sharedModel();
  assert.deepEqual([2, 3].map((r) => cellOf(model, 0, r - 1, 1).f), ['A2*2', 'A3*2']);
  const state = model.cellState('s1', 0, 1);
  assert.equal(state.sharedMaster, true);
  assert.equal(model.cellState('s1', 1, 1).sharedMaster, false);
  assert.equal(model.cellState('s1', 1, 1).formula, 'A2*2');
  // a value over the master
  const edit = doEdit(model, setCells('s1', [0, 1, '7']));
  assert.equal(edit.after.ops.length, 2, 'file-only fix first, the user cells last');
  assert.deepEqual(edit.after.ops[0].cells, [
    [1, 1, { formula: 'A2*2', cached: { t: 'n', v: 40 } }],
    [2, 1, { formula: 'A3*2', cached: { t: 'n', v: 60 } }],
  ]);
  assert.deepEqual(edit.after.ops[1].cells, [[0, 1, { value: { t: 'n', v: 7 } }]]);
  // the display: master changed, clones unchanged and still showing their results
  assert.equal(cellOf(model, 0, 0, 1).v, 7);
  assert.deepEqual([1, 2].map((r) => [cellOf(model, 0, r, 1).f, cellOf(model, 0, r, 1).v]), [['A2*2', 40], ['A3*2', 60]]);
  // a second change of the same master needs no new separation
  const again = S.planEdit(model, setCells('s1', [0, 1, '8']));
  assert.equal(again.after.ops.length, 1);
  // undo puts the group back
  S.invertEdit(model, edit);
  assert.equal(model.cellState('s1', 0, 1).sharedMaster, true);
  assert.equal(cellOf(model, 0, 0, 1).f, 'A1*2');
});

test('shared formula: clear and new formula over the master separate too; a clone edited alone does not', async () => {
  for (const intent of [clearRange('s1', 0, 1, 0, 1), setCells('s1', [0, 1, '=A1+1'])]) {
    const model = await sharedModel();
    const e = S.planEdit(model, intent);
    assert.equal(e.after.ops.length, 2);
    assert.deepEqual(e.after.ops[0].cells.map(([r, c]) => [r, c]), [[1, 1], [2, 1]]);
  }
  const model = await sharedModel();
  const e = S.planEdit(model, setCells('s1', [1, 1, '5']));
  assert.equal(e.after.ops.length, 1, 'editing a clone changes only that cell');
  S.applyEdit(model, e);
  // now the master changes: the clone already edited is not given a formula again
  const e2 = S.planEdit(model, setCells('s1', [0, 1, '1']));
  assert.deepEqual(e2.after.ops[0].cells.map(([r, c]) => [r, c]), [[2, 1]]);
  // an edit that covers master and a clone: the user's cell wins
  const model3 = await sharedModel();
  const e3 = S.planEdit(model3, setCells('s1', [0, 1, '1'], [1, 1, '2']));
  assert.deepEqual(e3.after.ops[0].cells.map(([r, c]) => [r, c]), [[2, 1]]);
  assert.deepEqual(e3.after.ops[1].cells.map(([r, c]) => [r, c]), [[0, 1], [1, 1]]);
});

test('hyperlink cell: the text is replaced, the link stays; rich text is read as its plain text', async () => {
  const model = await sharedModel();
  const link = cellOf(model, 0, 0, 2);
  assert.equal(link.link, 'https://example.com/');
  doEdit(model, setCells('s1', [0, 2, 'new text']));
  const after = cellOf(model, 0, 0, 2);
  assert.deepEqual([after.v, after.link], ['new text', 'https://example.com/']);
  assert.equal(model.cellState('s1', 1, 2).value.v, 'Rich text');
  // rich text typed back unchanged is no edit
  assert.deepEqual(S.planEdit(model, setCells('s1', [1, 2, 'Rich text'])), { noop: true });
  // a cleared hyperlink cell shows nothing
  doEdit(model, clearRange('s1', 0, 2, 0, 2));
  assert.equal(cellOf(model, 0, 0, 2)?.link, undefined);
});

// ----- the edit-model fuzz -----

/** Small deterministic random numbers. */
function rng(seed) {
  let x = seed * 2654435761 + 12345;
  return (n) => ((x = (x * 1103515245 + 12345) & 0x7fffffff) % n);
}

/**
 * Random setCells / clear / undo / redo on the sheets of styled.xlsx without merges, checked against a naive model:
 * the loaded cells (a second, never edited model) plus a map of what the user typed, kept as immutable snapshots per
 * edit. Every few steps the WHOLE grid of every sheet must equal the naive grid; undoing everything must give back
 * the loaded workbook with an empty overlay. `invert` can be replaced (mutation check). Returns the number of failures.
 */
async function runFuzz(seed, steps, { invert = S.invertEdit } = {}) {
  const model = await sampleModel();
  const pristine = await sampleModel();
  const rnd = rng(seed);
  const sheets = model.getMeta().sheets.map((s, i) => ({ ...s, i })).filter((s) => s.merges.length === 0 && s.state !== 'veryHidden');
  let naive = new Map(); // `${sheet},${r},${c}` -> content | null (cleared)
  const done = [];
  const redo = [];
  let failures = 0;
  const fail = (why) => {
    failures++;
    if (process.env.FS_FUZZ_VERBOSE) console.log(`fuzz seed ${seed}: ${why}`);
  };
  const norm = (cell) => ({ t: cell.t === 'd' ? 'n' : cell.t, v: cell.v, f: cell.f });
  const expectedGrid = (nv, s) => {
    const exp = new Map();
    for (const [key, cell] of grid(pristine, s)) if (cell.t !== 'z') exp.set(key, norm(cell));
    for (const [key, content] of nv) {
      const [ks, r, c] = key.split(',');
      if (+ks !== s) continue;
      if (content === null) exp.delete(`${r},${c}`);
      else exp.set(`${r},${c}`, content);
    }
    return exp;
  };
  const sameGrid = (a, b) => a.size === b.size && [...b].every(([k, e]) => a.has(k) && JSON.stringify(norm(a.get(k))) === JSON.stringify(norm(e)));
  const check = (label) => {
    for (const { i } of sheets) {
      const actual = new Map();
      for (const [key, cell] of grid(model, i)) if (cell.t !== 'z') actual.set(key, norm(cell));
      if (!sameGrid(actual, expectedGrid(naive, i))) {
        fail(`${label}: grid of sheet ${i} differs`);
        return;
      }
    }
  };
  const word = () => ['a', 'bb', 'ccc', 'Hello World', 'x y', 'ü'][rnd(6)];
  /** A position whose number format is General (a typed number stays a number). */
  const pick = (s) => {
    for (let tries = 0; tries < 20; tries++) {
      const r = rnd(4) ? rnd(30) : rnd(300);
      const c = rnd(4) ? rnd(8) : rnd(40);
      if (pristine.cellState(sheets[s].id, r, c).numFmt === 'General') return [r, c];
    }
    return undefined;
  };
  for (let step = 0; step < steps; step++) {
    const op = rnd(10);
    const s = rnd(sheets.length);
    const id = sheets[s].id;
    const sheetIndex = sheets[s].i;
    if (op < 6) {
      const cells = [];
      const after = new Map(naive);
      for (let n = 1 + rnd(6); n > 0; n--) {
        const rc = pick(s);
        if (!rc) continue;
        const [r, c] = rc;
        const kind = rnd(5);
        let text;
        let content;
        if (kind === 0) {
          text = String(rnd(1000));
          content = { t: 'n', v: Number(text), f: undefined };
        } else if (kind === 1) {
          text = word();
          content = { t: 's', v: text, f: undefined };
        } else if (kind === 2) {
          text = '';
          content = null;
        } else if (kind === 3) {
          const k = rnd(9);
          text = `=A1+${k}`;
          content = { t: 's', v: '', f: `A1+${k}` };
        } else {
          text = rnd(2) ? 'TRUE' : 'FALSE';
          content = { t: 'b', v: text === 'TRUE', f: undefined };
        }
        cells.push([r, c, text]);
        after.set(`${sheetIndex},${r},${c}`, content);
      }
      const plan = S.planEdit(model, setCells(id, ...cells));
      if (plan.noop) {
        if (!sameGrid(expectedGrid(after, sheetIndex), expectedGrid(naive, sheetIndex))) fail('a no-op plan for an edit that changes the grid');
        continue;
      }
      if (!plan.kind) {
        fail(`plan refused: ${JSON.stringify(plan)}`);
        continue;
      }
      S.applyEdit(model, plan);
      done.push({ edit: plan, before: naive, after });
      naive = after;
      redo.length = 0;
    } else if (op < 8) {
      const r0 = rnd(4) ? rnd(30) : rnd(300);
      const c0 = rnd(4) ? rnd(8) : rnd(40);
      const rg = { r0, c0, r1: r0 + rnd(6), c1: c0 + rnd(4) };
      const after = new Map(naive);
      for (const key of expectedGrid(naive, sheetIndex).keys()) {
        const [r, c] = key.split(',').map(Number);
        if (r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1) after.set(`${sheetIndex},${r},${c}`, null);
      }
      const plan = S.planEdit(model, { kind: 'clear', sheetId: id, ranges: [rg], what: 'contents' });
      if (plan.noop) {
        if (!sameGrid(expectedGrid(after, sheetIndex), expectedGrid(naive, sheetIndex))) fail('a no-op clear that would change the grid');
        continue;
      }
      if (!plan.kind) {
        fail(`clear refused: ${JSON.stringify(plan)}`);
        continue;
      }
      S.applyEdit(model, plan);
      done.push({ edit: plan, before: naive, after });
      naive = after;
      redo.length = 0;
    } else if (op === 8) {
      const top = done.pop();
      if (!top) continue;
      invert(model, top.edit);
      naive = top.before;
      redo.push(top);
    } else {
      const top = redo.pop();
      if (!top) continue;
      S.applyEdit(model, top.edit);
      naive = top.after;
      done.push(top);
    }
    if (step % 4 === 0) check(`step ${step}`);
  }
  check('final');
  // undo everything: the loaded workbook, with nothing left in the overlay
  while (done.length) invert(model, done.pop().edit);
  for (const { i, id } of sheets) {
    if (!sameGrid(grid(model, i), grid(pristine, i)) || JSON.stringify([...grid(model, i)]) !== JSON.stringify([...grid(pristine, i)])) fail(`sheet ${i} is not back to the loaded state`);
    if (model.viewOf(id).over.size !== 0) fail(`sheet ${i} keeps overlay entries`);
  }
  return failures;
}

test('edit-model fuzz: random setCells / clear / undo / redo against a naive model (3 seeds)', async () => {
  for (const seed of [1, 2, 3]) {
    const failures = await runFuzz(seed, 250);
    assert.equal(failures, 0, `seed ${seed}`);
  }
});

test('edit-model fuzz bites: a broken inverse is caught (mutation check)', async () => {
  // The inverse forgets to restore the last entry it should: the full-grid check must fail.
  const broken = (m, e) => {
    const saved = e.before.cells;
    e.before = { cells: saved.slice(1) };
    try {
      return S.invertEdit(m, e);
    } finally {
      e.before = { cells: saved };
    }
  };
  let caught = 0;
  for (const seed of [1, 2, 3]) caught += (await runFuzz(seed, 250, { invert: broken })) > 0 ? 1 : 0;
  assert.equal(caught, 3, 'every seed notices the broken inverse');
});

// ----- sizes -----

test('a 100,000-cell paste: plan + apply + undo are fast and undo data stays small', async () => {
  const model = await sampleModel();
  const id = model.getMeta().sheets[1].id;
  const cells = [];
  for (let r = 0; r < 1000; r++) for (let c = 0; c < 100; c++) cells.push([r + 100, c, String(r * 100 + c)]);
  const heapBefore = (() => {
    global.gc?.();
    return process.memoryUsage().heapUsed;
  })();
  const t0 = performance.now();
  const edit = S.planEdit(model, setCells(id, ...cells));
  assert.ok(edit.kind);
  S.applyEdit(model, edit);
  const t1 = performance.now();
  assert.equal(cellOf(model, 1, 100 + 999, 99).v, 999 * 100 + 99);
  assert.equal(model.viewOf(id).over.size, 100000);
  S.invertEdit(model, edit);
  const t2 = performance.now();
  assert.equal(model.viewOf(id).over.size, 0);
  const heapAfter = (() => {
    global.gc?.();
    return process.memoryUsage().heapUsed;
  })();
  // generous bounds (a CI machine can be slow): the point is "not seconds"
  assert.ok(t1 - t0 < 4000, `plan + apply took ${Math.round(t1 - t0)} ms`);
  assert.ok(t2 - t1 < 2000, `undo took ${Math.round(t2 - t1)} ms`);
  if (global.gc) assert.ok(heapAfter - heapBefore < 20 * 1024 * 1024 + cells.length * 100, `heap grew ${(heapAfter - heapBefore) / 1048576} MB (edit kept alive: plan ops + undo data)`);
});

test('a table written without headerRowCount (OOXML default: one header row) still blocks its header cells (ExcelJS reads it as "no header")', async () => {
  const wb = await S.loadWorkbook(new Uint8Array(fs.readFileSync(path.join(SAMPLES, 'xlsx-save', 'objects.xlsx'))));
  const model = S.workbookToModel(wb);
  assert.deepEqual(model.tableCellAt('s1', 0, 0), { name: 'SalesTable', part: 'header' });
  const plan = S.planEdit(model, setCells('s1', [0, 0, 'Renamed']));
  assert.ok('blocked' in plan && plan.blocked[0].code === 'table', JSON.stringify(plan));
  assert.ok(!('blocked' in S.planEdit(model, setCells('s1', [1, 0, 'East']))), 'a body cell of the table is fine');
});
