// Undo / redo pressed while a save of the same xlsx runs (P8 re-check). VS Code sets its save point when the save call has
// returned to it; the undo/redo step runs in FileStudio's queue right after the save. The tab must still show dirty, and
// the tab dirty state is what makes VS Code ask on close, or the undone/redone edit is lost. data/medium-save.xlsx is big enough that a save takes about a second.
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vscode = require('vscode');
const JSZip = require('jszip');
const h = require('./helpers');

const MEDIUM = path.join(__dirname, 'data', 'medium-save.xlsx');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let api;
let id = 50000;
let n = 0;

async function openMedium() {
  const name = `saveundo-${++n}.xlsx`;
  fs.copyFileSync(MEDIUM, h.fileUri(name).fsPath);
  const uri = h.fileUri(name);
  await h.openWith(name);
  await h.waitFor(() => { try { return api.state(uri.toString()).editState; } catch { return undefined; } }, { timeout: 60000, message: 'edit state' });
  await h.waitFor(async () => (await api.post(uri.toString(), { type: 'getRows', reqId: ++id, sheet: 0, sheetId: 's1', start: 0, end: 1 })).some((m) => m.type === 'rows'), { timeout: 60000, message: 'ready' });
  await sleep(3000);
  return { uri, key: uri.toString() };
}
async function edit(ctx, r, c, text) {
  const editId = ++id;
  const v0 = api.state(ctx.key).version;
  const msgs = await api.post(ctx.key, { type: 'edit', editId, baseVersion: v0, intent: { kind: 'setCells', sheetId: 's1', cells: [[r, c, text]] } });
  const res = msgs.find((m) => m.type === 'editResult' && m.editId === editId);
  if (!res) await h.waitFor(() => api.state(ctx.key).version > v0, { message: 'edit applied' });
  else assert.ok(res.accepted, 'edit accepted: ' + JSON.stringify(res));
}
async function fileHas(uri, text) {
  const zip = await JSZip.loadAsync(fs.readFileSync(uri.fsPath));
  const sst = zip.file('xl/sharedStrings.xml') ? await zip.file('xl/sharedStrings.xml').async('string') : '';
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  return sst.includes(`>${text}<`) || sheet.includes(`>${text}<`);
}
/** The document has unsaved changes, so the tab must be dirty and closing must prompt (the editor stays open). */
async function assertDirtyAndPrompts(ctx, what) {
  const st = api.state(ctx.key);
  assert.equal(st.dirty, true, `${what}: document dirty`);
  assert.equal(h.isDirty(ctx.uri), true, `${what}: tab dirty`);
  // VS Code asks before closing exactly when the tab is dirty (a modal prompt cannot be answered from a test)
}

suite('save-undo-redo', () => {
  suiteSetup(async () => {
    const ext = h.extension();
    api = ext.isActive ? ext.exports : await ext.activate();
    await h.closeAll();
  });
  teardown(async () => {
    await h.closeAll();
  });

  test('redo pressed while the save runs keeps the tab dirty', async () => {
    const ctx = await openMedium();
    await edit(ctx, 0, 0, 'one');
    await edit(ctx, 1, 0, 'two');
    await h.undo();
    await h.waitFor(() => api.state(ctx.key).version >= 3, { message: 'undo done' });
    const saving = h.save();
    await sleep(150);
    const redoing = h.redo();
    await saving;
    await redoing;
    await sleep(2500);
    assert.equal(await fileHas(ctx.uri, 'two'), false, 'the save did not include the redo');
    await assertDirtyAndPrompts(ctx, 'redo during save');
  });

  test('undo pressed while the save runs keeps the tab dirty', async () => {
    const ctx = await openMedium();
    await edit(ctx, 0, 0, 'one');
    await edit(ctx, 1, 0, 'two');
    const saving = h.save();
    await sleep(150);
    const undoing = h.undo();
    await saving;
    await undoing;
    await sleep(2500);
    await assertDirtyAndPrompts(ctx, 'undo during save');
  });

  test('autoSave afterDelay: redo while the auto save runs keeps the tab dirty', async () => {
    const files = vscode.workspace.getConfiguration('files');
    await files.update('autoSave', 'afterDelay', vscode.ConfigurationTarget.Workspace);
    await files.update('autoSaveDelay', 300, vscode.ConfigurationTarget.Workspace);
    try {
      const ctx = await openMedium();
      await edit(ctx, 0, 0, 'one');
      await edit(ctx, 1, 0, 'two');
      await h.waitFor(() => !h.isDirty(ctx.uri) && !api.state(ctx.key).dirty, { timeout: 30000, message: 'auto saved two' });
      await h.undo();
      await sleep(600);
      await h.redo();
      await sleep(1000);
      // the redo ended up after an auto save started; the next auto save may still write it, so only the end state matters:
      // either it is saved (clean tab and clean document) or both are dirty. Never "tab clean, document dirty".
      await h.waitFor(() => api.state(ctx.key).dirty === h.isDirty(ctx.uri), { timeout: 20000, message: 'tab and document agree' });
      await sleep(6000);
      const st = api.state(ctx.key);
      assert.equal(h.isDirty(ctx.uri), st.dirty, 'tab and document agree');
      if (!st.dirty) assert.equal(await fileHas(ctx.uri, 'two'), true, 'clean means the redo is in the file');
    } finally {
      await files.update('autoSave', undefined, vscode.ConfigurationTarget.Workspace);
      await files.update('autoSaveDelay', undefined, vscode.ConfigurationTarget.Workspace);
    }
  });
});
