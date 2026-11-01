// xlsx cell editing of 0.2.0 P2 in a REAL VS Code (contract section 14, P2 acceptance, integration part). The edits are the
// webview's `edit` messages delivered through the test API (`post`); saving, undo, redo and revert go through VS Code's own
// commands. Every test works on a temporary copy of a sample. Prompts are answered with `api.answer(id, button)`.
// Set FS_P2_OUT to a folder to also keep a copy of every saved workbook there (the Excel check reads them).
'use strict';
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const vscode = require('vscode');
const h = require('./helpers');

const sha = (uri) => crypto.createHash('sha256').update(fs.readFileSync(uri.fsPath)).digest('hex');
const TOUCHED = ['fileStudio.editing.enabled', 'fileStudio.xlsx.editing', 'files.readonlyInclude'];
const settings = () => vscode.workspace.getConfiguration();

let api;
let counter = 0;
let editId = 1000;

const safeState = (uri) => {
  try {
    return api.state(uri.toString());
  } catch {
    return undefined;
  }
};

async function openCopy(sample = 'styled.xlsx') {
  const name = `it-edit-${++counter}-${sample.split('/').pop()}`;
  const uri = h.copyFile(sample, name);
  await h.openWith(name);
  await h.waitFor(() => safeState(uri)?.editState, { message: `the edit state of ${name}` });
  // the webview has said ready (a message delivered before that is not handled)
  await h.waitFor(async () => (await api.post(uri.toString(), { type: 'getRows', reqId: ++editId, sheet: 0, sheetId: 's1', start: 0, end: 1 })).some((m) => m.type === 'rows'), { message: 'the webview to be ready' });
  return { uri, key: uri.toString(), name };
}

/** Delivers one `edit` as the first panel would; resolves with the messages the host posted back. */
async function post(ctx, intent, extra = {}) {
  const id = ++editId;
  const messages = await api.post(ctx.key, { type: 'edit', editId: id, baseVersion: api.state(ctx.key).version, intent, ...extra });
  const result = messages.find((m) => m.type === 'editResult' && m.editId === id);
  assert.ok(result, 'exactly one editResult came back');
  assert.equal(messages.filter((m) => m.type === 'editResult').length, 1);
  return { result, messages };
}

const setCell = (r, c, text, sheetId = 's1') => ({ kind: 'setCells', sheetId, cells: [[r, c, text]] });
const clean = (ctx, message = 'the tab to be clean') => h.waitFor(() => !h.isDirty(ctx.uri), { message });
const sheet1Xml = async (uri) => (await JSZip.loadAsync(fs.readFileSync(uri.fsPath))).file('xl/worksheets/sheet1.xml').async('string');

function keepForExcel(ctx, label) {
  const out = process.env.FS_P2_OUT;
  if (!out) return;
  fs.mkdirSync(out, { recursive: true });
  fs.copyFileSync(ctx.uri.fsPath, path.join(out, `${label}-${path.basename(ctx.uri.fsPath)}`));
}

async function readRows(ctx, sheetId = 's1') {
  const messages = await api.post(ctx.key, { type: 'getRows', reqId: ++editId, sheet: 0, sheetId, start: 0, end: 130 });
  return messages.find((m) => m.type === 'rows');
}
const textAt = (rows, r, c) => rows.rows.find((row) => row.r === r)?.cells.find((cell) => cell.c === c)?.v;

suite('xlsx cell editing (P2)', () => {
  suiteSetup(async () => {
    const ext = h.extension();
    api = ext.isActive ? ext.exports : await ext.activate();
    assert.ok(api && api.version === 1, 'the test API');
    await h.closeAll();
  });
  teardown(async () => {
    for (const key of TOUCHED) await settings().update(key, undefined, vscode.ConfigurationTarget.Global);
    await h.closeAll();
  });

  test('init says editable; edit -> patch rows -> Save -> the value is in the file and shows after reopening', async () => {
    const ctx = await openCopy();
    assert.deepEqual(safeState(ctx.uri).editState, { editable: true, kinds: ['setCells', 'clear'] });
    const { result } = await post(ctx, setCell(0, 0, 'typed in the grid'));
    assert.equal(result.accepted, true);
    assert.equal(textAt({ rows: result.patches[0].rows }, 0, 0), 'typed in the grid');
    await h.waitFor(() => h.isDirty(ctx.uri), { message: 'the tab to become dirty' });
    await h.save();
    await clean(ctx);
    keepForExcel(ctx, 'value');
    assert.match(await sheet1Xml(ctx.uri), /<c r="A1"[^>]*t="s"/);
    await h.closeAll();
    await h.openWith(ctx.name);
    await h.waitFor(() => safeState(ctx.uri)?.editState, { message: 'reopened' });
    const rows = await readRows(ctx);
    assert.equal(textAt(rows, 0, 0), 'typed in the grid', 'reopened file shows the saved value');
  });

  test('a number and a formula with a _xlfn. function are saved with the _xlfn. prefix Excel needs', async () => {
    const ctx = await openCopy();
    assert.equal((await post(ctx, setCell(100, 20, '1234.5'))).result.accepted, true);
    assert.equal((await post(ctx, setCell(101, 20, '=DAYS(5,3)'))).result.accepted, true, 'typed without the prefix');
    assert.equal((await post(ctx, setCell(102, 20, '=_xlfn.IFNA(U101,0)'))).result.accepted, true);
    await h.save();
    await clean(ctx);
    keepForExcel(ctx, 'formula');
    const xml = await sheet1Xml(ctx.uri);
    assert.match(xml, /<c r="U101"[^>]*><v>1234.5<\/v>/);
    assert.match(xml, /<f>_xlfn\.DAYS\(5,3\)<\/f>/);
    assert.match(xml, /<f>_xlfn\.IFNA\(U101,0\)<\/f>/);
  });

  test('undo and redo (VS Code commands): edit + undo = clean, Save writes nothing; redo = dirty again', async () => {
    const ctx = await openCopy();
    const before = sha(ctx.uri);
    await post(ctx, setCell(0, 0, 'edit one'));
    await h.waitFor(() => h.isDirty(ctx.uri), { message: 'dirty' });
    await h.undo();
    await h.waitFor(() => api.state(ctx.key).dirty === false, { message: 'FileStudio clean after undo' });
    await h.save();
    await clean(ctx);
    await h.sleep(300);
    assert.equal(sha(ctx.uri), before, 'nothing written');
    await h.redo();
    await h.waitFor(() => api.state(ctx.key).dirty === true, { message: 'dirty after redo' });
    assert.equal(textAt(await readRows(ctx), 0, 0), 'edit one');
  });

  test('one Clear and one paste-sized setCells are each one undo step', async () => {
    const ctx = await openCopy();
    const cells = [];
    for (let r = 100; r < 120; r++) for (let c = 20; c < 25; c++) cells.push([r, c, `v${r}-${c}`]);
    const pasted = (await post(ctx, { kind: 'setCells', sheetId: 's1', cells })).result;
    assert.equal(pasted.accepted, true, JSON.stringify(pasted));
    assert.equal((await post(ctx, { kind: 'clear', sheetId: 's1', ranges: [{ r0: 100, c0: 20, r1: 119, c1: 24 }], what: 'contents' })).result.accepted, true);
    await h.undo(); // the clear
    assert.equal(textAt(await readRows(ctx), 103, 23), 'v103-23');
    await h.undo(); // the paste
    await h.waitFor(() => api.state(ctx.key).dirty === false, { message: 'two undos = clean' });
  });

  test('kill switch: toggled while open -> editState disabled and an edit is refused; unsaved edits can still be saved (D-075)', async () => {
    const ctx = await openCopy();
    await post(ctx, setCell(0, 0, 'before the switch'));
    await settings().update('fileStudio.xlsx.editing', false, vscode.ConfigurationTarget.Global);
    await h.waitFor(() => safeState(ctx.uri).editState.code === 'disabled', { message: 'editState disabled' });
    const { result } = await post(ctx, setCell(0, 0, 'refused'));
    assert.deepEqual([result.accepted, result.code], [false, 'disabled']);
    await h.save();
    await clean(ctx, 'saved although editing is off');
    assert.match(await sheet1Xml(ctx.uri), /<c r="A1"/);
    await settings().update('fileStudio.xlsx.editing', undefined, vscode.ConfigurationTarget.Global);
    await h.waitFor(() => safeState(ctx.uri).editState.editable === true, { message: 'editable again' });
  });

  test('files.readonlyInclude -> readOnly state and a refused edit', async () => {
    const ctx = await openCopy();
    await settings().update('files.readonlyInclude', { '**/it-edit-*': true }, vscode.ConfigurationTarget.Global);
    await h.waitFor(() => safeState(ctx.uri).editState.code === 'readOnly', { message: 'editState readOnly' });
    const { result } = await post(ctx, setCell(0, 0, 'x'));
    assert.deepEqual([result.accepted, result.code], [false, 'readOnly']);
    assert.equal(h.isDirty(ctx.uri), false);
  });

  test('a lossy workbook: edit refused `lossy` until Edit anyway is answered, then accepted and saved', async () => {
    const ctx = await openCopy('xlsx-save/objects.xlsx');
    assert.equal(safeState(ctx.uri).editState.code, 'lossy');
    const refused = await post(ctx, setCell(30, 10, 'x'));
    assert.deepEqual([refused.result.accepted, refused.result.code], [false, 'lossy']);
    await api.post(ctx.key, { type: 'editAnyway' }); // no queued answer = Cancel
    assert.equal(safeState(ctx.uri).editState.code, 'lossy');
    api.answer('editAnyway', 'Edit anyway');
    await api.post(ctx.key, { type: 'editAnyway' });
    await h.waitFor(() => safeState(ctx.uri).editState.editable === true, { message: 'editable after Edit anyway' });
    assert.equal((await post(ctx, setCell(30, 10, 'lossy but kept'))).result.accepted, true);
    await h.save();
    await clean(ctx);
    keepForExcel(ctx, 'lossy');
  });

  test('split editor: both panels see the same version; one edit is one undo step for the document', async () => {
    const ctx = await openCopy();
    await vscode.commands.executeCommand('workbench.action.splitEditor');
    await h.waitFor(() => h.allTabs().filter((t) => t.input?.uri?.toString() === ctx.key).length >= 2, { message: 'two tabs of the file' });
    await h.sleep(800);
    const { result } = await post(ctx, setCell(0, 0, 'from the first panel'));
    assert.equal(result.accepted, true);
    assert.equal(api.state(ctx.key).version, result.version, 'one host version per document');
    const tabs = h.allTabs().filter((t) => t.input?.uri?.toString() === ctx.key);
    await h.waitFor(() => tabs.every((t) => t.isDirty), { message: 'both tabs dirty' });
    await h.undo();
    await h.waitFor(() => api.state(ctx.key).dirty === false, { message: 'clean after one undo' });
  });

  test('Revert: edits gone, the view refreshed (editState back), an old undo step does nothing', async () => {
    const ctx = await openCopy();
    const before = sha(ctx.uri);
    await post(ctx, setCell(0, 0, 'to be reverted'));
    await h.revert();
    await clean(ctx);
    assert.equal(sha(ctx.uri), before);
    assert.equal(textAt(await readRows(ctx), 0, 0) === 'to be reverted', false);
    assert.equal(safeState(ctx.uri).editState.editable, true);
  });
});
