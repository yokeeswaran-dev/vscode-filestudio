// P2 host wiring of xlsx cell editing (src/viewerProvider.ts SheetEditorProvider + SheetDocument; contract §3.2, §3.3,
// §3.5, §3.6, §14 P2): edit -> gate -> version rule -> plan -> apply -> journal -> editResult / patch, editAnyway,
// sendInit editState / version, getRows / getStats by sheetId, the kill switch and D-075, undo / redo broadcast, the V3
// check against the display model. No VS Code window: the host runs against test/host/fake-vscode.js.
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, extensionUri, token, writeTmp, lastOf, waitFor, tick } = require('./load-host');

const api = host.createTestApi();
const EDIT_TEXT = host.EDIT_TEXT;

beforeEach(() => T.reset());
afterEach(() => {
  delete host.sheetEditHooks.afterApply;
});

let counter = 0;
const unique = (name) => `se${++counter}-${name}`;

async function bytesOf(fill, patchXml) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('B1').value = 5;
  workbook.addWorksheet('Two').getCell('A1').value = 'two';
  if (fill) fill(workbook);
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  if (patchXml) await patchXml(zip);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** A document with `panels` webview panels, each initialised (ready -> init). */
async function openBook(bytes, panels = 1) {
  const uri = writeTmp(unique('book.xlsx'), bytes ?? (await bytesOf()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  const events = [];
  doc.onDidChange((e) => {
    if (typeof e.undo === 'function') events.push(e);
  });
  const list = [];
  for (let i = 0; i < panels; i++) {
    const panel = T.createPanel();
    provider.resolveCustomEditor(doc, panel, token);
    panel.send({ type: 'ready' });
    const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'init');
    list.push({ panel, init });
  }
  return { uri, provider, doc, events, panel: list[0].panel, init: list[0].init, panels: list };
}

let editId = 100;
const sendEdit = async (open, intent, extra = {}, panel = open.panel) => {
  const id = ++editId;
  panel.posted.length = 0;
  panel.send({ type: 'edit', editId: id, baseVersion: open.doc.versions.version, intent, ...extra });
  return waitFor(() => panel.posted.find((m) => m.type === 'editResult' && m.editId === id), 5000, 'editResult');
};
const setCell = (r, c, text, sheetId = 's1') => ({ kind: 'setCells', sheetId, cells: [[r, c, text]] });

const cellText = (rows, r, c) => rows.find((row) => row.r === r)?.cells.find((cell) => cell.c === c)?.v;

async function sheetXml(file, n = 1) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  return zip.file(`xl/worksheets/sheet${n}.xml`).async('string');
}

// ===== init, getRows, getStats =====

test('init carries version, readOnly = !editable and editState with the implemented kinds', async () => {
  const o = await openBook();
  assert.equal(o.init.readOnly, false);
  assert.equal(o.init.version, 0);
  assert.deepEqual(o.init.editState, { editable: true, kinds: ['setCells', 'clear'] });
  assert.equal(o.init.meta.sheets[0].id, 's1');
});

test('getRows and getStats resolve the sheet by sheetId and answer with version and sheetId', async () => {
  const o = await openBook();
  o.panel.posted.length = 0;
  o.panel.send({ type: 'getRows', reqId: 1, sheet: 0, sheetId: 's2', start: 0, end: 5 });
  const rows = await waitFor(() => lastOf(o.panel, 'rows'), 3000, 'rows');
  assert.equal(cellText(rows.rows, 0, 0), 'two', 'sheet "Two" was found by its id, not by position 0');
  assert.equal(rows.sheetId, 's2');
  assert.equal(rows.version, 0);
  o.panel.send({ type: 'getRows', reqId: 2, sheet: 0, sheetId: 's9', start: 0, end: 5 });
  const gone = await waitFor(() => o.panel.posted.find((m) => m.type === 'rows' && m.reqId === 2), 3000, 'rows of a missing sheet');
  assert.deepEqual(gone.rows, [], 'a sheet that is gone answers empty, no error view');
  assert.equal(lastOf(o.panel, 'error'), undefined);
  o.panel.send({ type: 'getStats', reqId: 3, sheet: 0, sheetId: 's1', ranges: [{ r0: 0, c0: 1, r1: 0, c1: 1 }] });
  const stats = await waitFor(() => lastOf(o.panel, 'stats'), 3000, 'stats');
  assert.equal(stats.stats.sum, 5);
  // 0.1.2 requests (no sheetId) keep working
  o.panel.send({ type: 'getRows', reqId: 4, sheet: 1, start: 0, end: 5 });
  const old = await waitFor(() => o.panel.posted.find((m) => m.type === 'rows' && m.reqId === 4), 3000, 'rows by index');
  assert.equal(cellText(old.rows, 0, 0), 'two');
});

// ===== the edit path =====

test('an accepted edit: one editResult with version and the changed rows; the journal, dirty flag and display follow', async () => {
  const o = await openBook();
  const result = await sendEdit(o, setCell(0, 0, 'changed'));
  assert.equal(result.accepted, true);
  assert.equal(result.version, 1);
  assert.equal(result.patches.length, 1);
  assert.equal(result.patches[0].sheetId, 's1');
  assert.equal(cellText(result.patches[0].rows, 0, 0), 'changed');
  assert.equal(o.doc.hasUnsavedChanges, true);
  assert.equal(o.events.length, 1, 'one VS Code undo step');
  assert.equal(o.events[0].label, 'Edit A1');
  assert.equal(o.panel.posted.filter((m) => m.type === 'editResult').length, 1, 'exactly one answer');
  assert.equal(o.panel.posted.find((m) => m.type === 'patch'), undefined, 'the sender gets no patch');
  // reads see the new value
  o.panel.send({ type: 'getRows', reqId: 50, sheet: 0, sheetId: 's1', start: 0, end: 1 });
  const rows = await waitFor(() => o.panel.posted.find((m) => m.type === 'rows' && m.reqId === 50), 3000, 'rows');
  assert.equal(cellText(rows.rows, 0, 0), 'changed');
});

test('typed numbers become numbers; a formula is journaled with its file syntax; a change of nothing is accepted without a version bump', async () => {
  const o = await openBook();
  const num = await sendEdit(o, setCell(0, 1, '42'));
  assert.equal(num.accepted, true);
  const same = await sendEdit(o, setCell(0, 1, '42'));
  assert.deepEqual([same.accepted, same.version, same.patches], [true, 1, undefined], 'no change: no new version, no undo step');
  assert.equal(o.events.length, 1);
  const formula = await sendEdit(o, setCell(1, 0, '=B1*2'));
  assert.equal(formula.accepted, true);
  const saved = await o.doc.save(token).then(() => sheetXml(o.uri.fsPath));
  assert.match(saved, /<c r="B1"[^>]*><v>42<\/v>/);
  assert.match(saved, /<f>B1\*2<\/f>/);
});

test('every failure path answers exactly once: invalid, blocked, not yet implemented, unknown sheet', async () => {
  const o = await openBook();
  const checks = [
    [setCell(0, 0, '=1+', 's1'), 'invalid'],
    [setCell(0, 0, '=SUM(A1:A2)+[1]Sheet1!A1', 's1'), undefined], // external reference: blocked or invalid, never accepted
    [{ kind: 'clear', sheetId: 's1', ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], what: 'formats' }, 'blocked'],
    [setCell(0, 0, 'x', 's77'), 'invalid'],
  ];
  for (const [intent, code] of checks) {
    const r = await sendEdit(o, intent);
    assert.equal(r.accepted, false, JSON.stringify(intent));
    if (code) assert.equal(r.code, code, JSON.stringify(intent));
    assert.equal(o.panel.posted.filter((m) => m.type === 'editResult').length, 1);
  }
  assert.equal(o.doc.hasUnsavedChanges, false, 'nothing was journaled');
  assert.equal(o.events.length, 0);
});

test('fix1: an edit that would grow the workbook past the 100,000-cell edit limit is refused `blocked`; the document stays editable', async () => {
  const o = await openBook();
  const far = await sendEdit(o, setCell(4999, 25, 'x')); // Z5000: 5000 x 26 cells in use
  assert.deepEqual([far.accepted, far.code], [false, 'blocked']);
  assert.match(far.reason, /too large to edit safely/);
  const cells = [];
  for (let r = 0; r < 317; r++) for (let c = 0; c < 317; c++) cells.push([r, c, '1']);
  const paste = await sendEdit(o, { kind: 'setCells', sheetId: 's1', cells });
  assert.deepEqual([paste.accepted, paste.code], [false, 'blocked'], '317 x 317 paste');
  assert.equal(o.doc.hasUnsavedChanges, false);
  const ok = await sendEdit(o, setCell(2, 2, 'fits'));
  assert.equal(ok.accepted, true);
});

test('stale rule: a value edit with an older baseVersion is fine; one older than structureVersion is refused and the view refreshed', async () => {
  const o = await openBook();
  await sendEdit(o, setCell(0, 0, 'a'));
  await sendEdit(o, setCell(0, 0, 'b'));
  const old = await sendEdit(o, setCell(0, 0, 'c'), { baseVersion: 0 });
  assert.equal(old.accepted, true, 'value edits on unmoved cells: the last write wins');
  // a structural edit (as P4 will make) raises structureVersion
  await o.doc.queue.run(() => o.doc.commit({ label: 'Merge', structural: true, after: { ops: [], expect: [] } }));
  const { structureVersion, version } = o.doc.versions;
  assert.equal(structureVersion, version);
  o.panel.posted.length = 0;
  o.panel.send({ type: 'edit', editId: 900, baseVersion: structureVersion - 1, intent: setCell(0, 0, 'late') });
  const stale = await waitFor(() => o.panel.posted.find((m) => m.type === 'editResult'), 3000, 'editResult');
  assert.deepEqual([stale.accepted, stale.code, stale.reason], [false, 'stale', EDIT_TEXT.stale]);
  const refreshed = await waitFor(() => lastOf(o.panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(refreshed.version, version);
  const fine = await sendEdit(o, setCell(0, 0, 'ok'), { baseVersion: structureVersion });
  assert.equal(fine.accepted, true);
});

test('then: "save" saves through VS Code after the edit when this editor is the active tab', async () => {
  const o = await openBook();
  T.addTab({ input: new vscode.TabInputCustom(o.uri, host.SheetEditorProvider.viewType), label: 'book.xlsx', isActive: true });
  T.calls.length = 0;
  await sendEdit(o, setCell(0, 0, 'saved'), { then: 'save' });
  await waitFor(() => T.calls.find((c) => c[0] === 'executeCommand' && c[1] === 'workbench.action.files.save'), 3000, 'save command');
});

// ===== two panels =====

test('split editor: the other panel gets a patch with the same version; the sender gets only its editResult', async () => {
  const o = await openBook(undefined, 2);
  const [a, b] = o.panels.map((p) => p.panel);
  b.posted.length = 0;
  const result = await sendEdit(o, setCell(2, 2, 'from A'), {}, a);
  const patch = await waitFor(() => lastOf(b, 'patch'), 3000, 'patch');
  assert.equal(patch.version, result.version);
  assert.equal(patch.patches[0].sheetId, 's1');
  assert.equal(cellText(patch.patches[0].rows, 2, 2), 'from A');
  assert.equal(lastOf(a, 'patch'), undefined);
  // the used range grew (row 3, column C): both get the new sheet meta
  assert.ok(patch.sheets?.[0] && patch.sheets[0].id === 's1', 'sheet meta with the new used range');
  assert.ok(result.sheets?.[0] && result.sheets[0].id === 's1');
  assert.deepEqual(a.posted.filter((m) => m.type === 'editResult').length, 1);
});

test('undo and redo (VS Code edit events) patch every panel with the new version; edit + undo = not dirty, save writes nothing', async () => {
  const o = await openBook(undefined, 2);
  const [a, b] = o.panels.map((p) => p.panel);
  const before = fs.readFileSync(o.uri.fsPath);
  await sendEdit(o, setCell(0, 0, 'edited'), {}, a);
  b.posted.length = 0;
  a.posted.length = 0;
  await o.events[0].undo();
  assert.equal(o.doc.hasUnsavedChanges, false);
  for (const p of [a, b]) {
    const patch = await waitFor(() => lastOf(p, 'patch'), 3000, 'patch after undo');
    assert.equal(patch.version, o.doc.versions.version);
    assert.equal(cellText(patch.patches[0].rows, 0, 0), 'hello');
  }
  await o.doc.save(token);
  assert.deepEqual(fs.readFileSync(o.uri.fsPath), before, 'nothing written');
  b.posted.length = 0;
  await o.events[0].redo();
  const redo = await waitFor(() => lastOf(b, 'patch'), 3000, 'patch after redo');
  assert.equal(cellText(redo.patches[0].rows, 0, 0), 'edited');
  assert.equal(o.doc.hasUnsavedChanges, true);
});

// ===== gate: kill switch, read-only, lossy, poison =====

test('kill switch: editState goes to every panel, an edit is refused `disabled`; a dirty document still saves and undoes (D-075)', async () => {
  const o = await openBook(undefined, 2);
  const [a, b] = o.panels.map((p) => p.panel);
  await sendEdit(o, setCell(0, 0, 'keep me'), {}, a);
  b.posted.length = 0;
  T.config.set('fileStudio.xlsx.editing', false);
  T.fireConfigChange(['fileStudio.xlsx.editing']);
  const state = await waitFor(() => lastOf(b, 'editState'), 3000, 'editState');
  assert.deepEqual([state.state.editable, state.state.code], [false, 'disabled']);
  const refused = await sendEdit(o, setCell(0, 0, 'no'), {}, a);
  assert.deepEqual([refused.accepted, refused.code], [false, 'disabled']);
  assert.match(refused.reason, /fileStudio\.xlsx\.editing/);
  // save and undo of the edit made before still work
  assert.equal(o.doc.hasUnsavedChanges, true);
  await o.doc.save(token);
  assert.equal(o.doc.hasUnsavedChanges, false);
  assert.match(await sheetXml(o.uri.fsPath), /keep me|<c r="A1"/);
  await o.events[0].undo();
  assert.equal(o.doc.hasUnsavedChanges, true, 'undo after the save is a change again');
  T.config.set('fileStudio.xlsx.editing', true);
  T.fireConfigChange(['fileStudio.xlsx.editing']);
  await waitFor(() => lastOf(a, 'editState')?.state.editable === true, 3000, 'editable again');
});

test('files.readonlyInclude: init is read-only and an edit is refused `readOnly`', async () => {
  const name = `locked/${unique('book.xlsx')}`;
  T.config.set('files.readonlyInclude', { '**/locked/**': true });
  const uri = writeTmp(name, await bytesOf());
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(doc, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'init');
  assert.equal(init.readOnly, true);
  assert.deepEqual([init.editState.editable, init.editState.code], [false, 'readOnly']);
  const o = { doc, panel };
  const r = await sendEdit(o, setCell(0, 0, 'x'));
  assert.deepEqual([r.accepted, r.code], [false, 'readOnly']);
  assert.equal(doc.hasUnsavedChanges, false);
});

async function lossyBook() {
  // A chart part makes the workbook lossy (content FileStudio does not display).
  return bytesOf(undefined, async (zip) => {
    zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
    zip.file('xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>');
    const ct = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'));
  });
}

test('lossy workbook: edit refused `lossy` until Edit anyway is answered in the host modal, then accepted', async () => {
  const o = await openBook(await lossyBook());
  assert.ok(o.doc.lossyFeatures.length > 0, 'the fixture is lossy');
  assert.equal(o.init.editState.editable, false);
  assert.equal(o.init.editState.code, 'lossy');
  assert.ok(o.init.editState.canEditAnyway);
  const refused = await sendEdit(o, setCell(0, 0, 'x'));
  assert.deepEqual([refused.accepted, refused.code], [false, 'lossy']);
  // the message alone is never consent: no queued answer = Cancel
  o.panel.posted.length = 0;
  o.panel.send({ type: 'editAnyway' });
  await tick(80);
  assert.equal(o.doc.canAcceptLossy, true);
  api.answer('editAnyway', 'Edit anyway');
  o.panel.send({ type: 'editAnyway' });
  const state = await waitFor(() => [...o.panel.posted].reverse().find((m) => m.type === 'editState' && m.state.editable), 3000, 'editable');
  assert.deepEqual(state.state.kinds, ['setCells', 'clear']);
  const ok = await sendEdit(o, setCell(0, 0, 'x'));
  assert.equal(ok.accepted, true);
});

test('poison: a poisoned document refuses new edits `internal` with the poison text and says so in editState', async () => {
  const o = await openBook();
  await sendEdit(o, setCell(0, 0, 'a'));
  o.doc.poison(new Error('test poison'));
  const refused = await sendEdit(o, setCell(0, 0, 'b'));
  assert.deepEqual([refused.accepted, refused.code, refused.reason], [false, 'internal', EDIT_TEXT.poisoned]);
  const state = await waitFor(() => lastOf(o.panel, 'editState'), 3000, 'editState');
  assert.deepEqual([state.state.editable, state.state.code], [false, 'internal']);
  // save still writes the last good state
  await o.doc.save(token);
  assert.equal(o.doc.hasUnsavedChanges, false);
});

// ===== file contents =====

test('value and clear on a shared-formula master: the group is separated, the clones keep their results in the saved file', async () => {
  const bytes = await bytesOf(
    (wb) => {
      const ws = wb.addWorksheet('S');
      for (let r = 1; r <= 3; r++) {
        ws.getCell(`A${r}`).value = r * 10;
        ws.getCell(`B${r}`).value = { formula: `A${r}*2`, result: r * 20 };
      }
    },
    async (zip) => {
      let xml = await zip.file('xl/worksheets/sheet3.xml').async('string');
      xml = xml.replace(/<f>A1\*2<\/f>/, '<f t="shared" ref="B1:B3" si="0">A1*2</f>').replace(/<f>A2\*2<\/f>/, '<f t="shared" si="0"/>').replace(/<f>A3\*2<\/f>/, '<f t="shared" si="0"/>');
      zip.file('xl/worksheets/sheet3.xml', xml);
    },
  );
  for (const intent of [setCell(0, 1, '7', 's3'), { kind: 'clear', sheetId: 's3', ranges: [{ r0: 0, c0: 1, r1: 0, c1: 1 }], what: 'contents' }]) {
    const o = await openBook(bytes);
    const r = await sendEdit(o, intent);
    assert.equal(r.accepted, true, JSON.stringify(intent));
    await o.doc.save(token);
    const xml = await sheetXml(o.uri.fsPath, 3);
    assert.doesNotMatch(xml, /t="shared"/, 'no orphan clone is left');
    assert.match(xml, /<f>A2\*2<\/f><v>40<\/v>/);
    assert.match(xml, /<f>A3\*2<\/f><v>60<\/v>/);
  }
});

test('merged cells: a single edit of a non-anchor goes to the anchor; a block with a non-empty non-anchor is blocked', async () => {
  const o = await openBook(
    await bytesOf((wb) => {
      const ws = wb.addWorksheet('M');
      ws.getCell('A1').value = 'anchor';
      ws.mergeCells('A1:B2');
    }),
  );
  const single = await sendEdit(o, setCell(1, 1, 'into anchor', 's3'));
  assert.equal(single.accepted, true);
  assert.equal(cellText(single.patches[0].rows, 0, 0), 'into anchor');
  const block = await sendEdit(o, { kind: 'setCells', sheetId: 's3', cells: [[0, 0, 'p'], [0, 1, 'q']] });
  assert.deepEqual([block.accepted, block.code], [false, 'blocked']);
  assert.ok(block.where?.length);
  const dropped = await sendEdit(o, { kind: 'setCells', sheetId: 's3', cells: [[0, 0, 'p2'], [0, 1, '']] });
  assert.equal(dropped.accepted, true, 'an empty non-anchor value is dropped');
});

test('V3: a display / file mismatch (corrupted ops) makes the save fail and writes nothing', async () => {
  const o = await openBook();
  const before = fs.readFileSync(o.uri.fsPath);
  host.sheetEditHooks.afterApply = (plan) => {
    const op = plan.after.ops[plan.after.ops.length - 1];
    op.cells[0][2] = { value: { t: 's', v: 'not what the grid shows' } };
  };
  await sendEdit(o, setCell(0, 0, 'what the grid shows'));
  await assert.rejects(o.doc.save(token), /could not save/i);
  assert.deepEqual(fs.readFileSync(o.uri.fsPath), before, 'nothing written');
  assert.equal(o.doc.hasUnsavedChanges, true, 'the edit is kept');
});

test('lone surrogate and oversize text are `invalid` (the display model never sees them)', async () => {
  const o = await openBook();
  const r = await sendEdit(o, setCell(0, 0, 'bad \uD83D'));
  assert.deepEqual([r.accepted, r.code], [false, 'invalid']);
  assert.equal(o.doc.hasUnsavedChanges, false);
});

test('revert: invalidate with the new version and editState; the old undo step is inert', async () => {
  const o = await openBook();
  await sendEdit(o, setCell(0, 0, 'edited'));
  o.panel.posted.length = 0;
  await o.provider.revertCustomDocument(o.doc, token);
  const inv = await waitFor(() => lastOf(o.panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(inv.version, o.doc.versions.version);
  assert.deepEqual(inv.editState, { editable: true, kinds: ['setCells', 'clear'] });
  assert.equal(o.doc.hasUnsavedChanges, false);
  await o.events[0].undo();
  assert.equal(o.doc.hasUnsavedChanges, false, 'the old step belongs to the version before the revert');
});

test('command: undo / redo / save run behind the queued edits, only for the active tab', async () => {
  const o = await openBook();
  T.calls.length = 0;
  o.panel.send({ type: 'command', id: 'undo' });
  await tick(60);
  assert.deepEqual(T.calls.filter((c) => c[0] === 'executeCommand'), [], 'not the active tab: nothing runs');
  T.addTab({ input: new vscode.TabInputCustom(o.uri, host.SheetEditorProvider.viewType), label: 'book.xlsx', isActive: true });
  o.panel.send({ type: 'command', id: 'undo' });
  await waitFor(() => T.calls.find((c) => c[0] === 'executeCommand' && c[1] === 'undo'), 3000, 'undo command');
  o.panel.send({ type: 'command', id: 'save' });
  await waitFor(() => T.calls.find((c) => c[0] === 'executeCommand' && c[1] === 'workbench.action.files.save'), 3000, 'save command');
});
