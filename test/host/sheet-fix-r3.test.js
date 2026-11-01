// P1 fix round 3: regression tests for the confirmed review findings of the xlsx save path (SheetDocument).
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const { execFileSync } = require('node:child_process');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, tmp, extensionUri, token, writeTmp, waitFor } = require('./load-host');

const api = host.createTestApi();
const A1 = (text) => ({ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: text } }]] });
const win = { skip: process.platform !== 'win32' };

beforeEach(() => T.reset());
afterEach(() => {
  host.sheetSaveHooks.fs = fsp;
  delete host.sheetSaveHooks.afterVerify;
});

let counter = 0;
async function xlsxBytes(patch) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('B1').value = 5;
  sheet.getCell('A2').value = 7;
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  if (!patch) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  await patch(zip);
  return zip.generateAsync({ type: 'nodebuffer' });
}
/** A cell (B2) stored inside row 1: valid for the loader, not for Excel (the scan says "unpatchable"). */
const wrongRowBytes = () =>
  xlsxBytes(async (zip) => {
    const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    zip.file('xl/worksheets/sheet1.xml', xml.replace('</row>', '<c r="B2" t="n"><v>1</v></c></row>'));
  });
async function open(name = 'book.xlsx', bytes) {
  const uri = writeTmp(`fr3${++counter}-${name}`, bytes ?? (await xlsxBytes()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  await doc.loading;
  return { uri, provider, doc };
}
const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

test('R3-1a: an edit on a file the scan finds unpatchable is refused (the scan is awaited first)', async () => {
  const { uri, doc } = await open('wrongrow.xlsx', await wrongRowBytes());
  await assert.rejects(api.pushPackageOps(uri.toString(), [A1('mine')]), /cannot edit this file safely/);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.journalLength, 0);
});

test('R3-1b: an edit that was journaled before the scan ended cannot be saved or backed up', async () => {
  const { uri, doc } = await open('wrongrow2.xlsx', await wrongRowBytes());
  const before = sha(uri.fsPath);
  await doc.queue.run(() => doc.commit({ label: 'e', after: { ops: [A1('mine')], expect: [] } }));
  await assert.rejects(doc.save(token), /cannot edit this file safely/);
  assert.equal(sha(uri.fsPath), before, 'the file is not written');
  await assert.rejects(doc.backup(vscode.Uri.file(path.join(tmp, `fr3-bk-${++counter}.xlsx`)), token), /cannot edit this file safely/);
  // After the scan result is known, commit() itself refuses.
  assert.throws(() => doc.commit({ label: 'e2', after: { ops: [A1('more')], expect: [] } }), /cannot edit this file safely/);
});

test('R3-2a: "Keep My Edits" does not put an undo step on top of the real edits', async () => {
  const { uri, doc } = await open('keep.xlsx');
  const raw = [];
  doc.onDidChange((e) => raw.push(e));
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await api.pushPackageOps(uri.toString(), [A1('two')]);
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/k.txt', 'k')));
  const count = raw.length;
  api.answer('externalChange', 'Keep My Edits');
  await doc.reloadIfChangedOnDisk();
  await waitFor(() => doc.testState.diverged, 2000, 'Keep My Edits');
  const added = raw.slice(count);
  assert.ok(added.length >= 1, 'VS Code is told the document changed (stays dirty)');
  assert.equal(added.filter((e) => typeof e.undo === 'function').length, 0, 'no undo / redo step on top of the real edits');
  assert.equal(doc.hasUnsavedChanges, true);
});

test('R3-2b: a refused (stale) undo does not put an undo step on top of the real edits', async () => {
  const { uri, doc } = await open('stale.xlsx');
  const raw = [];
  doc.onDidChange((e) => raw.push(e));
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const stale = raw.find((e) => typeof e.undo === 'function');
  await doc.revert();
  await api.pushPackageOps(uri.toString(), [A1('two')]);
  const count = raw.length;
  await stale.undo();
  assert.equal(raw.slice(count).filter((e) => typeof e.undo === 'function').length, 0);
  assert.equal(doc.hasUnsavedChanges, true);
});

test('R3-3: a custom Windows ACL survives Save when the file name contains "(I)"', win, async () => {
  const { uri, doc } = await open('Budget (I).xlsx');
  execFileSync('icacls', [uri.fsPath, '/inheritance:r']);
  execFileSync('icacls', [uri.fsPath, '/grant:r', `${process.env.USERNAME}:(F)`]);
  const aclOf = (file) =>
    execFileSync('icacls', [file], { encoding: 'utf8' }).split('\n').filter((l) => l.includes(':(')).map((l) => l.replace(file, '').trim()).sort().join('|');
  const before = aclOf(uri.fsPath);
  assert.ok(!/\(I\)/.test(before), `precondition: ${before}`);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  assert.equal(aclOf(uri.fsPath), before);
});

test('R3-5a: a small workbook with big pictures is not locked read-only', async () => {
  const bytes = await xlsxBytes((zip) => zip.file('xl/media/image1.png', crypto.randomBytes(6 * 1024 * 1024), { compression: 'STORE' }));
  assert.ok(bytes.length > 5 * 1024 * 1024, `fixture is ${bytes.length} bytes`);
  const { uri, doc } = await open('pictures.xlsx', bytes);
  await doc.ensureScan();
  assert.equal(doc.gateState.block, undefined);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('R3-5b: a small zip that unpacks to a huge sheet is read-only', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  for (let r = 1; r <= 1000; r++) sheet.getCell(r, 1).value = `${r}-${'x'.repeat(30000)}`; // about 30 MB of text, compresses to almost nothing
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  assert.ok(bytes.length < 1024 * 1024, `fixture is ${bytes.length} bytes`);
  const { doc } = await open('bomb.xlsx', bytes);
  assert.match(doc.gateState.block, /too large to edit safely/);
});

test('R3-7: a load error names the pre-image an interrupted in-place write left next to the file', async () => {
  const { uri } = await open('crash.xlsx');
  const prev = path.join(path.dirname(uri.fsPath), `.~${path.basename(uri.fsPath)}.fs-abcdef123456.prev`);
  fs.writeFileSync(prev, fs.readFileSync(uri.fsPath));
  fs.writeFileSync(uri.fsPath, fs.readFileSync(uri.fsPath).subarray(0, 100));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  await doc.loading;
  assert.ok(doc.loadError, 'precondition: the load failed');
  assert.ok(String(doc.loadError.message ?? doc.loadError).includes(prev), String(doc.loadError.message ?? doc.loadError));
});

test('R3-8: a late answer to an old "changed on disk" message is ignored when the document is clean again', async () => {
  const { uri, doc } = await open('late.xlsx');
  const raw = [];
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  doc.onDidChange((e) => raw.push(e));
  api.answer('externalChange', 'Keep My Edits');
  const asking = doc.askExternalChange();
  doc.savedStateId = doc.stateId; // the user saved (or reverted) while the message was open
  await asking;
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.diverged, false);
  assert.equal(raw.length, 0);
});

test('R3-9: stale temp files of this file are removed before a save; fresh ones and other files are kept', async () => {
  const { uri, doc } = await open('orphan.xlsx');
  const dir = path.dirname(uri.fsPath);
  const base = path.basename(uri.fsPath);
  const stale = path.join(dir, `.~${base}.fs-aaaaaaaaaaaa.tmp`);
  const fresh = path.join(dir, `.~${base}.fs-bbbbbbbbbbbb.tmp`);
  const other = path.join(dir, `.~other-${base}.fs-cccccccccccc.tmp`);
  for (const f of [stale, fresh, other]) fs.writeFileSync(f, 'x');
  const old = new Date(Date.now() - 3600 * 1000);
  fs.utimesSync(stale, old, old);
  fs.utimesSync(other, old, old);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  assert.equal(fs.existsSync(stale), false, 'stale temp removed');
  assert.equal(fs.existsSync(fresh), true, 'fresh temp kept (maybe another save is running)');
  assert.equal(fs.existsSync(other), true, "another file's temp kept");
});
