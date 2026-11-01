// P1 fix round 1: regression tests for the confirmed review findings of the xlsx save path (SheetDocument).
// Same setup as sheet-document.test.js: host code against test/host/fake-vscode.js, prompts through the test API.
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, extensionUri, token, writeTmp } = require('./load-host');

const api = host.createTestApi();
const cells = (list, sheet = 1) => ({ kind: 'cells', sheet, cells: list });
const A1 = (text) => cells([[0, 0, { value: { t: 's', v: text } }]]);

beforeEach(() => T.reset());
afterEach(() => {
  host.sheetSaveHooks.fs = fsp;
  delete host.sheetSaveHooks.afterVerify;
});

let counter = 0;
async function xlsxBytes() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('B1').value = 5;
  workbook.addWorksheet('Two').getCell('A1').value = 'two';
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
async function open(name = 'book.xlsx') {
  const uri = writeTmp(`fr1${++counter}-${name}`, await xlsxBytes());
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  const events = [];
  doc.onDidChange((e) => events.push(e));
  await doc.loading;
  return { uri, provider, doc, events };
}
const isContentEvent = (e) => typeof e.undo !== 'function' && typeof e.redo !== 'function';

test('F1: an edit also fires a content-change event (sticky changed flag), so save -> undo -> new edit cannot look clean in VS Code', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('A')]);
  assert.equal(events.filter((e) => typeof e.undo === 'function').length, 1, 'one undo step');
  assert.equal(events.filter(isContentEvent).length, 1, 'one content-change event');
  assert.equal(doc.hasUnsavedChanges, true);
});

test('F2: an edit queued before the reload step of reloadIfChangedOnDisk is kept (no silent discard)', async () => {
  const { uri, doc } = await open();
  let release;
  const gate = new Promise((r) => (release = r));
  const busy = doc.queue.run(() => gate); // the queue is busy (a save, a backup, an earlier edit)
  const changed = await xlsxBytes();
  const zip = await JSZip.loadAsync(changed);
  zip.file('docProps/x.txt', 'external');
  fs.writeFileSync(uri.fsPath, await zip.generateAsync({ type: 'nodebuffer' }));
  api.answer('externalChange', 'Keep My Edits');
  const reload = doc.reloadIfChangedOnDisk();
  const edit = api.pushPackageOps(uri.toString(), [A1('mine')]);
  await new Promise((r) => setTimeout(r, 50));
  release();
  await busy;
  await edit;
  await reload;
  assert.equal(doc.hasUnsavedChanges, true, 'the edit is still unsaved work');
  assert.equal(doc.testState.journalLength, 1);
});

test('F3: a short lock on the target (EPERM on rename) is retried and the save succeeds', { skip: process.platform !== 'win32' }, async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  let failures = 2;
  host.sheetSaveHooks.fs = {
    ...fsp,
    rename: async (a, b) => {
      if (failures-- > 0) throw Object.assign(new Error('EPERM: operation not permitted, rename'), { code: 'EPERM' });
      return fsp.rename(a, b);
    },
  };
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
  const zip = await JSZip.loadAsync(fs.readFileSync(uri.fsPath));
  assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'), /mine|<v>\d+<\/v>/);
});

test('F10: the temp file name is short even for a very long file name', async () => {
  const { uri, doc } = await open(`${'n'.repeat(120)}.xlsx`);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const tempNames = [];
  host.sheetSaveHooks.fs = {
    ...fsp,
    open: async (p, flags, mode) => {
      if (flags === 'wx') tempNames.push(path.basename(p));
      return fsp.open(p, flags, mode);
    },
  };
  await doc.save(token);
  assert.ok(tempNames.length >= 1);
  assert.ok(tempNames.every((n) => n.length <= 80), `temp names: ${tempNames.map((n) => n.length)}`);
});

test('F11: EACCES on the write check says "no write permission", not "open in another program"', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  host.sheetSaveHooks.fs = {
    ...fsp,
    open: async (p, flags, mode) => {
      if (flags === 'r+') throw Object.assign(new Error('denied'), { code: 'EACCES' });
      return fsp.open(p, flags, mode);
    },
  };
  await assert.rejects(doc.save(token), (err) => /no write permission/.test(err.message) && !/another program/.test(err.message));
});

test('F12: Save As onto the source with another spelling (case) is a plain Save', { skip: process.platform === 'linux' }, async () => {
  const { uri, doc } = await open('Same.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const upper = vscode.Uri.file(path.join(path.dirname(uri.fsPath), path.basename(uri.fsPath).toUpperCase()));
  await doc.saveAs(upper, token);
  assert.equal(doc.hasUnsavedChanges, false, 'the document is clean again');
  assert.ok(fs.existsSync(`${uri.fsPath}.bak`), 'the .bak of a normal save was made');
});

test('poison: commit() refuses new edits on a poisoned document (contract 3.5)', async () => {
  const { doc } = await open();
  doc.poison(new Error('test'));
  assert.throws(() => doc.commit({ label: 'x', after: { ops: [], expect: [] } }), /internal error|poison|Revert/i);
  assert.equal(doc.testState.journalLength, 0);
});

test('onYield: the cancel hook of a save also reaches the verify step (V1-V3 are cooperative)', async () => {
  // A tall sheet: the row walks of build and verify yield many times. The verify step must poll the cancel token too.
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  for (let r = 1; r <= 40000; r++) sheet.addRow([r, 'row ' + r]);
  const uri = writeTmp('fr1-tall.xlsx', Buffer.from(await workbook.xlsx.writeBuffer()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  await doc.loading;
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  let polls = 0;
  const counting = { get isCancellationRequested() { polls++; return false; }, onCancellationRequested: () => ({ dispose() {} }) };
  await doc.save(counting);
  // Build alone polls about 7 times on this sheet; with the verify step (V1-V3) the save polls about 30 times.
  assert.ok(polls >= 20, `the token was polled only ${polls} times: verify does not yield`);
});
