// P1 fix round 2: regression tests for the confirmed review findings of the xlsx save path (SheetDocument).
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, tmp, extensionUri, token, writeTmp } = require('./load-host');

const api = host.createTestApi();
const A1 = (text) => ({ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: text } }]] });

beforeEach(() => T.reset());
afterEach(() => {
  host.sheetSaveHooks.fs = fsp;
  delete host.sheetSaveHooks.afterVerify;
});

let counter = 0;
async function xlsxBytes(extra) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('B1').value = 5;
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  if (!extra) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  zip.file('docProps/x.txt', extra);
  return zip.generateAsync({ type: 'nodebuffer' });
}
async function open(name = 'book.xlsx') {
  const uri = writeTmp(`fr2${++counter}-${name}`, await xlsxBytes());
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  await doc.loading;
  return { uri, provider, doc };
}
const settles = (promise, ms) =>
  Promise.race([promise.then(() => 'done', () => 'done'), new Promise((r) => setTimeout(() => r('HANG'), ms))]);

test('R2-1: a read-back mismatch does not point to a .bak that was never made', async () => {
  T.config.set('fileStudio.xlsx.backupOnSave', false);
  const { uri, doc } = await open('rb.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  let armed = false;
  host.sheetSaveHooks.afterVerify = () => {
    armed = true;
  };
  host.sheetSaveHooks.fs = {
    ...fsp,
    readFile: async (p, ...rest) => {
      const data = await fsp.readFile(p, ...rest);
      if (armed && p === uri.fsPath) {
        const copy = Buffer.from(data);
        copy[copy.length - 1] ^= 0xff;
        return copy;
      }
      return data;
    },
  };
  await assert.rejects(doc.save(token), (err) => {
    assert.ok(!fs.existsSync(`${uri.fsPath}.bak`), 'precondition: no .bak');
    assert.doesNotMatch(err.message, /\.bak/, err.message);
    return /does not match/.test(err.message);
  });
});

test('R2-2: a save queued before Revert does not deadlock the queue', async () => {
  const { uri, doc } = await open('dead.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const busy = doc.queue.run(() => new Promise((r) => setTimeout(r, 300)));
  const save = doc.save(token);
  const revert = doc.revert();
  assert.equal(await settles(Promise.all([busy, save, revert]), 3000), 'done', 'save and revert finish');
  const later = api.pushPackageOps(uri.toString(), [A1('later')]);
  assert.equal(await settles(later, 3000), 'done', 'the queue is alive');
});

test('R2-3: a restored backup whose sidecar has sourceStamp null still asks before overwriting a changed file', async () => {
  const { uri, provider, doc } = await open('nullstamp.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const dest = vscode.Uri.file(path.join(tmp, `fr2-null-${++counter}.xlsx`));
  const backup = await doc.backup(dest, token);
  const sidecarPath = `${dest.fsPath}.json`;
  const sidecar = JSON.parse(fs.readFileSync(sidecarPath, 'utf8'));
  sidecar.sourceStamp = null;
  fs.writeFileSync(sidecarPath, JSON.stringify(sidecar));
  fs.writeFileSync(uri.fsPath, await xlsxBytes('external')); // another program, while VS Code was closed
  const external = fs.readFileSync(uri.fsPath);
  const restored = provider.openCustomDocument(uri, { backupId: backup.id }, token);
  await restored.loading;
  await assert.rejects(restored.save(token), (err) => err instanceof vscode.CancellationError);
  assert.deepEqual(fs.readFileSync(uri.fsPath), external, 'the external version is untouched');
});

test('R2-4: the patch check does not run when a workbook is opened, only when it is edited, saved or backed up', async () => {
  const { uri, doc } = await open('lazy.xlsx');
  await new Promise((r) => setTimeout(r, 50));
  assert.equal(doc.scan, undefined, 'no scan at open');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  assert.notEqual(doc.scan, undefined, 'scan after the first edit');
});

test('R2-5: a hot-exit backup does not keep the built workbook in memory', async () => {
  const { uri, doc } = await open('bk.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.backup(vscode.Uri.file(path.join(tmp, `fr2-bk-${++counter}.xlsx`)), token);
  assert.equal(doc.lastBackup, undefined);
});

test('R2-6: EPERM on the final rename gives the "open in another program" text, not the raw system error', { skip: process.platform !== 'win32' }, async () => {
  const { uri, doc } = await open('eperm.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  host.sheetSaveHooks.fs = {
    ...fsp,
    rename: async () => {
      throw Object.assign(new Error("EPERM: operation not permitted, rename 'a' -> 'b'"), { code: 'EPERM' });
    },
  };
  await assert.rejects(doc.save(token), (err) => /open in another program/.test(err.message) && !/EPERM/.test(err.message));
  assert.equal(doc.hasUnsavedChanges, true);
});

test('R2-7: a file deleted outside asks "recreate" instead of the backup question', async () => {
  const { uri, doc } = await open('deleted.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.unlinkSync(uri.fsPath);
  api.answer('diskChangedOnSave', 'Recreate');
  await doc.save(token); // a bakFailed question would be unanswered = Cancel = reject
  assert.ok(fs.existsSync(uri.fsPath));
  assert.equal(doc.hasUnsavedChanges, false);
});

const win = { skip: process.platform !== 'win32' };
const aclOf = (file) =>
  execFileSync('icacls', [file], { encoding: 'utf8' })
    .split('\n')
    .filter((l) => l.includes(':('))
    .map((l) => l.replace(file, '').trim())
    .sort()
    .join('|');

test('R2-8: a custom Windows ACL (not inherited) survives Save', win, async () => {
  const { uri, doc } = await open('acl.xlsx');
  execFileSync('icacls', [uri.fsPath, '/inheritance:r']);
  execFileSync('icacls', [uri.fsPath, '/grant:r', `${process.env.USERNAME}:(F)`]);
  const before = aclOf(uri.fsPath);
  assert.ok(!/\(I\)/.test(before), `precondition: ${before}`);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  assert.equal(aclOf(uri.fsPath), before);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('R2-9: the Hidden attribute survives Save', win, async () => {
  const { uri, doc } = await open('hidden.xlsx');
  execFileSync('attrib', ['+h', uri.fsPath]);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  assert.match(execFileSync('attrib', [uri.fsPath], { encoding: 'utf8' }).split(uri.fsPath)[0], /H/);
  execFileSync('attrib', ['-h', uri.fsPath]);
});
