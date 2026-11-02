// 0.2.1 host wiring of big sheets: a sheet above the size threshold loads row by row (rows are asked for with getRows),
// the 100,000-cell read-only limit does not apply to it, edits / undo / save work, and the unpacked-bytes rule (B1) counts
// what applied edits add. The threshold is lowered with host.lazyTuning so a small file takes the same path.
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ExcelJS = require('exceljs');

const { T, host, extensionUri, token, writeTmp, lastOf, waitFor } = require('./load-host');

const DEFAULTS = { ...host.lazyTuning };
beforeEach(() => {
  T.reset();
  Object.assign(host.lazyTuning, DEFAULTS, { sheetBytes: 500, chunkRows: 40, keepRows: 150 });
});
afterEach(() => Object.assign(host.lazyTuning, DEFAULTS));

let counter = 0;

async function bookBytes(rows, cols, far) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  for (let r = 1; r <= rows; r++) sheet.getCell(r, 1).value = r === 1 ? 'head' : `v${r}`;
  if (far) sheet.getCell(far.r, far.c).value = 1;
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function openBook(bytes) {
  const uri = writeTmp(`lz${++counter}-book.xlsx`, bytes);
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  const events = [];
  doc.onDidChange((e) => {
    if (typeof e.undo === 'function') events.push(e);
  });
  const panel = T.createPanel();
  provider.resolveCustomEditor(doc, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'init');
  return { uri, doc, events, panel, init };
}

let editId = 500;
async function sendEdit(o, intent) {
  const id = ++editId;
  o.panel.posted.length = 0;
  o.panel.send({ type: 'edit', editId: id, baseVersion: o.doc.versions.version, intent });
  return waitFor(() => o.panel.posted.find((m) => m.type === 'editResult' && m.editId === id), 8000, 'editResult');
}

let reqId = 900;
async function getRows(o, start, end) {
  const id = ++reqId;
  o.panel.posted.length = 0;
  o.panel.send({ type: 'getRows', reqId: id, sheet: 0, sheetId: 's1', start, end });
  return waitFor(() => o.panel.posted.find((m) => m.type === 'rows' && m.reqId === id), 8000, 'rows');
}
const textAt = (reply, r, c) => reply.rows.find((row) => row.r === r)?.cells.find((cell) => cell.c === c)?.v;

test('lazy host: a workbook with 100,000+ cells is editable when its sheet loads row by row; rows come with getRows', async () => {
  const o = await openBook(await bookBytes(300, 1, { r: 1001, c: 100 })); // used range 1001 x 100 = 100,100 cells
  assert.equal(o.init.readOnly, false);
  assert.equal(o.init.meta.sheets[0].rowCount, 1001);
  assert.equal(o.doc.gateState.block, undefined);
  assert.equal(textAt(await getRows(o, 0, 5), 2, 0), 'v3');
  assert.equal(textAt(await getRows(o, 290, 300), 295, 0), 'v296');
  const res = await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: [[1000, 99, 'far'], [3, 0, 'near']] });
  assert.equal(res.accepted, true, JSON.stringify(res) + (T.logLines ? T.logLines.join("|") : ""));
  assert.equal(textAt(await getRows(o, 1000, 1001), 1000, 99), 'far');
  assert.equal(textAt(await getRows(o, 3, 4), 3, 0), 'near');
});

test('lazy host: edit, save, reopen: the file holds the edit and the other rows are unchanged', async () => {
  const bytes = await bookBytes(400, 1);
  const o = await openBook(bytes);
  assert.equal((await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: [[250, 0, 'changed']] })).accepted, true);
  await o.doc.save(token);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(o.uri.fsPath);
  const sheet = wb.getWorksheet('Data');
  assert.equal(sheet.getCell('A251').value, 'changed');
  assert.equal(sheet.getCell('A250').value, 'v250');
  assert.equal(sheet.getCell('A400').value, 'v400');
});

test('lazy host: edit then undo = not dirty (nothing to write); redo brings it back', async () => {
  const o = await openBook(await bookBytes(400, 1));
  await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: [[10, 0, 'x']] });
  assert.equal(o.doc.hasUnsavedChanges, true);
  // scroll far away: other rows are loaded and the old ones may be evicted, the edited row stays
  for (let a = 100; a < 400; a += 20) await getRows(o, a, a + 20);
  await o.events[0].undo();
  assert.equal(o.doc.hasUnsavedChanges, false);
  assert.equal(textAt(await getRows(o, 10, 11), 10, 0), 'v11');
  await o.events[0].redo();
  assert.equal(textAt(await getRows(o, 10, 11), 10, 0), 'x');
});

test('lazy host: a request for far too many rows is refused with a text and the document stays editable', async () => {
  host.lazyTuning.maxRequestRows = 50;
  const o = await openBook(await bookBytes(400, 1));
  const res = await sendEdit(o, { kind: 'clear', sheetId: 's1', ranges: [{ r0: 0, c0: 0, r1: 300, c1: 0 }], what: 'contents' });
  assert.equal(res.accepted, false);
  assert.match(res.reason, /more than 10,000 rows|Change fewer rows/);
  assert.equal((await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: [[1, 0, 'ok']] })).accepted, true);
});

test('B1: the unpacked-bytes rule counts what applied edits add, and undo gives it back', async () => {
  Object.assign(host.lazyTuning, DEFAULTS); // a small workbook: the 25 MB rule
  const o = await openBook(await bookBytes(5, 1));
  const big = 'x'.repeat(32_000);
  const cells = (from) => Array.from({ length: 250 }, (_, i) => [from + i, 1, big]); // 8 million characters
  assert.equal((await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: cells(10) })).accepted, true);
  const refused = await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: cells(300) });
  assert.equal(refused.accepted, false);
  assert.match(refused.reason, /too large to edit safely/);
  await o.events[0].undo();
  assert.equal((await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: cells(300) })).accepted, true, 'after undo there is room again');
});

test('fix r1: an edit inside a tall merged area works when the first row of the merge is not loaded', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  for (let r = 1; r <= 400; r++) sheet.getCell(r, 1).value = `v${r}`;
  sheet.getCell(100, 5).value = 'master';
  sheet.mergeCells('E100:F300');
  const o = await openBook(Buffer.from(await workbook.xlsx.writeBuffer()));
  const res = await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells: [[200, 4, 'in merge']] });
  assert.equal(res.accepted, true, JSON.stringify(res));
  assert.equal(textAt(await getRows(o, 99, 100), 99, 4), 'in merge');
});

// ----- fix round 2 -----

test('fix r2: a getStats probe over too many rows of a big sheet is marked unknown (not "0 cells")', async () => {
  host.lazyTuning.maxRequestRows = 50;
  const o = await openBook(await bookBytes(400, 1));
  o.panel.posted.length = 0;
  o.panel.send({ type: 'getStats', reqId: 4001, sheet: 0, sheetId: 's1', ranges: [{ r0: 1, c0: 0, r1: 399, c1: 0 }] });
  const res = await waitFor(() => o.panel.posted.find((m) => m.type === 'stats' && m.reqId === 4001), 8000, 'stats');
  assert.equal(res.stats.unknown, true);
  o.panel.posted.length = 0;
  o.panel.send({ type: 'getStats', reqId: 4002, sheet: 0, sheetId: 's1', ranges: [{ r0: 1, c0: 0, r1: 20, c1: 0 }] });
  const ok = await waitFor(() => o.panel.posted.find((m) => m.type === 'stats' && m.reqId === 4002), 8000, 'stats');
  assert.equal(ok.stats.unknown, undefined);
  assert.equal(ok.stats.count, 20);
});

test('fix r2: a row window parses at most 100 rows per chunk and the host can answer other work between chunks', async () => {
  assert.ok(DEFAULTS.chunkRows <= 100, `chunkRows is ${DEFAULTS.chunkRows} (a 60-row request must not parse 500 rows in one block)`);
  const o = await openBook(await bookBytes(400, 1));
  let ticks = 0;
  let stop = false;
  const spin = () => {
    if (stop) return;
    ticks++;
    setImmediate(spin);
  };
  spin();
  await getRows(o, 0, 300); // 300 rows = 8 chunks of 40 (the test tuning)
  stop = true;
  assert.ok(ticks >= 7, `only ${ticks} event loop turns during a window of 8 chunks`);
});
