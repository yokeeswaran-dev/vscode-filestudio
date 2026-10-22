// Fix round 2 for F: regression tests (each failed before its fix).
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { vscode, T, host, extensionUri, token, writeTmp, lastOf, waitFor, tick } = require('./load-host');

beforeEach(() => T.reset());
const { parseEditMessage } = host;

const edit = (intent) => ({ type: 'edit', editId: 7, baseVersion: 3, intent });
const range = { r0: 0, c0: 0, r1: 1, c1: 1 };
const query = { scope: 's1', text: 'a', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' };
const intents = (sheetId) => ({
  setCells: { kind: 'setCells', sheetId, cells: [[0, 0, 'a']] },
  clear: { kind: 'clear', sheetId, ranges: [range], what: 'contents' },
  insertRows: { kind: 'insertRows', sheetId, at: 0, count: 1 },
  deleteRows: { kind: 'deleteRows', sheetId, at: 0, count: 1 },
  insertCols: { kind: 'insertCols', sheetId, at: 0, count: 1 },
  deleteCols: { kind: 'deleteCols', sheetId, at: 0, count: 1 },
  sort: { kind: 'sort', sheetId, range, key: 0, desc: false, header: false },
  replaceAll: { kind: 'replaceAll', query: { ...query, scope: sheetId }, replace: 'b' },
  style: { kind: 'style', sheetId, ranges: [range], delta: { bold: true } },
  merge: { kind: 'merge', sheetId, range },
  unmerge: { kind: 'unmerge', sheetId, range },
  resize: { kind: 'resize', sheetId, axis: 'col', index: [0], size: 40 },
  sheetAdd: { kind: 'sheetAdd', index: 0 },
  sheetRename: { kind: 'sheetRename', sheetId, name: 'x' },
  sheetDelete: { kind: 'sheetDelete', sheetId },
  sheetDuplicate: { kind: 'sheetDuplicate', sheetId },
  sheetMove: { kind: 'sheetMove', sheetId, index: 0 },
  sheetHide: { kind: 'sheetHide', sheetId, hidden: true },
  tabColor: { kind: 'tabColor', sheetId, argb: null },
  mdText: { kind: 'mdText', edits: [{ start: 0, end: 0, oldText: '', text: 'x' }] },
});

test('parseEditMessage ties the intent kind to the view kind', () => {
  const verdict = (name, view, sheetId) => parseEditMessage(edit(intents(sheetId)[name]), view);
  const csvOk = ['setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll'];
  for (const name of Object.keys(intents('s1'))) {
    const isMd = name === 'mdText';
    assert.equal(verdict(name, 'sheet', 's1').ok, !isMd, `${name} from sheet`);
    assert.equal(verdict(name, 'csv', 'csv').ok, csvOk.includes(name), `${name} from csv`);
    assert.equal(verdict(name, 'markdown', 's1').ok, isMd, `${name} from markdown`);
    for (const view of ['docx', 'pdf', 'pptx']) assert.equal(verdict(name, view, 's1').ok, false, `${name} from ${view}`);
    // no view kind given (shape check only): everything is allowed as before
    assert.equal(verdict(name, undefined, 's1').ok, true, `${name} without a view kind`);
  }
});

test('CSV insert / delete / resize counts are not capped at the Excel grid', () => {
  const del = (count, view) => parseEditMessage(edit({ kind: 'deleteRows', sheetId: view === 'csv' ? 'csv' : 's1', at: 0, count }), view);
  assert.equal(del(2_000_000, 'csv').ok, true, 'csv: more than 1,048,576 rows');
  assert.equal(del(1_048_576, 'sheet').ok, true);
  assert.equal(del(1_048_577, 'sheet').ok, false, 'xlsx keeps the Excel limit');
  const resize = (n, view) =>
    parseEditMessage(edit({ kind: 'resize', sheetId: view === 'csv' ? 'csv' : 's1', axis: 'row', index: Array.from({ length: n }, (_, i) => i), size: null }), view);
  assert.equal(resize(1_048_577, 'sheet').ok, false, 'xlsx keeps the Excel limit');
  const csvResize = parseEditMessage(
    edit({ kind: 'resize', sheetId: 'csv', axis: 'row', index: Array.from({ length: 1_048_577 }, (_, i) => i), size: null }),
    'csv',
  );
  assert.equal(csvResize.ok, false, 'resize is not a CSV edit at all');
});

// ===== a malformed read is answered =====

async function openSheet(name) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = 'a';
  const uri = writeTmp(name, Buffer.from(await workbook.xlsx.writeBuffer()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const document = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(document, panel, token);
  panel.send({ type: 'ready' });
  await waitFor(() => lastOf(panel, 'init'), 5000, 'sheet init');
  panel.posted.length = 0;
  return panel;
}

test('a find / filter / measure with a bad shape but a valid reqId is answered with an error, not dropped', async () => {
  const panel = await openSheet('r2-reads.xlsx');
  panel.send({ type: 'find', reqId: 11, query: { ...query, text: 'x'.repeat(1001) }, direction: 'all' });
  const found = await waitFor(() => lastOf(panel, 'findResult'), 3000, 'findResult');
  assert.equal(found.reqId, 11);
  assert.deepEqual([found.total, found.matches, found.truncated], [0, [], false]);
  assert.equal(typeof found.error, 'string');
  panel.send({ type: 'filter', reqId: 12, sheetId: 's1', col: -1 });
  const filtered = await waitFor(() => lastOf(panel, 'filterResult'), 3000, 'filterResult');
  assert.equal(filtered.reqId, 12);
  assert.equal(typeof filtered.error, 'string');
  panel.send({ type: 'measure', reqId: 13, sheetId: 's1', axis: 'diagonal', index: 0 });
  const measured = await waitFor(() => lastOf(panel, 'columnTexts'), 3000, 'columnTexts');
  assert.equal(measured.reqId, 13);
  assert.deepEqual(measured.texts, []);
  assert.equal(typeof measured.error, 'string');
  // no reqId at all: nothing to answer
  panel.posted.length = 0;
  panel.send({ type: 'find', query, direction: 'all' });
  await tick(50);
  assert.deepEqual(panel.posted, []);
});

// ===== checkTextFidelity =====

const codec = (bom) => ({
  encode: async (text) => Buffer.from((bom ? '﻿' : '') + text, 'utf8'),
  decode: async (bytes) => new TextDecoder('utf-8').decode(bytes),
});
function fidelityDoc(name, bytes, text, encoding) {
  const uri = writeTmp(name, bytes);
  const document = new T.FakeTextDocument(uri, text);
  Object.assign(document, { eol: vscode.EndOfLine.LF, encoding, isDirty: false });
  return document;
}

test('checkTextFidelity: a file with two byte order marks is not reported (VS Code keeps both when it saves)', async () => {
  T.config.clear();
  Object.assign(vscode.workspace, codec(true));
  const bom = Buffer.from([0xef, 0xbb, 0xbf]);
  const bytes = Buffer.concat([bom, bom, Buffer.from('# t\n', 'utf8')]);
  const r = await host.checkTextFidelity(fidelityDoc('r2-dbom.md', bytes, '# t\n', 'utf8bom'));
  assert.deepEqual(r, { ok: true });
});

test('checkTextFidelity: no per-line arrays on a big file (no split / match over the whole text)', async () => {
  T.config.clear();
  T.config.set('files.trimTrailingWhitespace', true);
  T.config.set('files.trimFinalNewlines', true);
  Object.assign(vscode.workspace, codec(false));
  const text = 'a \n'.repeat(400_000) + '\n\n';
  const big = (s) => typeof s === 'string' && s.length > 1_000_000;
  const origSplit = String.prototype.split;
  const origMatch = String.prototype.match;
  let calls = 0;
  String.prototype.split = function (...args) {
    if (big(String(this))) calls++;
    return origSplit.apply(this, args);
  };
  String.prototype.match = function (...args) {
    if (big(String(this))) calls++;
    return origMatch.apply(this, args);
  };
  let r;
  try {
    r = await host.checkTextFidelity(fidelityDoc('r2-big.md', Buffer.from(text), text, 'utf8'));
  } finally {
    String.prototype.split = origSplit;
    String.prototype.match = origMatch;
  }
  assert.equal(calls, 0, 'split / match ran over the whole text');
  assert.equal(r.ok, false);
  const details = r.canEditAnyway.details.join('|');
  assert.match(details, /400,000 lines|400000 lines/);
  assert.match(details, /remove 2 empty lines/);
});
