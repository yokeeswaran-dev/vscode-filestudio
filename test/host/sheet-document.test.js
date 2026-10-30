// SheetDocument (src/viewerProvider.ts, SHEET): the xlsx state machine and save path of 0.2.0 P1 (contract §4.1-§4.6):
// journal, state ids, generation, poison, save / save as / revert, safe write, .bak, external change, hot-exit backup
// and restore, and the test API hooks (pushPackageOps, lastSaveReport). No VS Code window: the host code runs against
// test/host/fake-vscode.js; prompts are answered through the test API queue (ask), never shown.
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, tmp, extensionUri, token, writeTmp, callsOf, waitFor, tick } = require('./load-host');

const api = host.createTestApi(); // ask() now answers from the queue; no answer = Cancel

const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const cells = (list, sheet = 1) => ({ kind: 'cells', sheet, cells: list });
const str = (v) => ({ t: 's', v });
const num = (v) => ({ t: 'n', v });

beforeEach(() => {
  T.reset();
});
afterEach(() => {
  host.sheetSaveHooks.fs = fsp;
  delete host.sheetSaveHooks.afterVerify;
  delete host.sheetSaveHooks.failBackupVerify;
});

let counter = 0;
const unique = (name) => `sd${++counter}-${name}`;

/** A small workbook written by ExcelJS: Data!A1 = 'hello', B1 = 5; second sheet 'Two'. */
async function xlsxBytes(extra) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('B1').value = 5;
  workbook.addWorksheet('Two').getCell('A1').value = 'two';
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  if (extra) await extra(zip);
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** Opens a workbook document the way the provider does; `events` collects the VS Code edit events. */
async function open(name = 'book.xlsx', bytes, openContext = {}) {
  const uri = writeTmp(unique(name), bytes ?? (await xlsxBytes()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, openContext, token);
  const events = [];
  // Edit events (with undo / redo) go to `events`; the content-change event every edit also fires (D-F1) is not an undo step.
  doc.onDidChange((e) => {
    if (typeof e.undo === 'function') events.push(e);
  });
  await doc.loading;
  return { uri, provider, doc, events };
}

/** Reads A1 of the first sheet from the file's XML (text of the cell, shared strings resolved). */
async function sheet1Xml(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  return zip.file('xl/worksheets/sheet1.xml').async('string');
}

const A1 = (text) => cells([[0, 0, { value: str(text) }]]);

// ===== journal, state ids, generation =====

test('edit -> undo -> save writes nothing (file bytes and mtime stay)', async () => {
  const { uri, doc, events } = await open();
  const before = sha(uri.fsPath);
  const mtime = fs.statSync(uri.fsPath).mtimeMs;
  await api.pushPackageOps(uri.toString(), [A1('changed')]);
  assert.equal(doc.hasUnsavedChanges, true);
  assert.equal(events.length, 1);
  await events[0].undo();
  assert.equal(doc.hasUnsavedChanges, false, 'back on the saved state');
  T.calls.length = 0;
  await doc.save(token);
  assert.equal(callsOf('fs.writeFile').length, 0);
  assert.equal(sha(uri.fsPath), before);
  assert.equal(fs.statSync(uri.fsPath).mtimeMs, mtime);
  assert.equal(fs.existsSync(`${uri.fsPath}.bak`), false, 'no .bak when nothing was written');
});

test('edit -> save writes the patched file; undo -> save writes the original cell XML again', async () => {
  const { uri, doc, events } = await open();
  const originalXml = await sheet1Xml(uri.fsPath);
  await api.pushPackageOps(uri.toString(), [A1('changed')]);
  await doc.save(token);
  assert.notEqual(await sheet1Xml(uri.fsPath), originalXml);
  assert.match(await sheet1Xml(uri.fsPath), /<c r="A1"[^>]*t="s"/);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(api.lastSaveReport(uri.toString()).changedParts.includes('xl/worksheets/sheet1.xml'), true);
  await events[0].undo();
  assert.equal(doc.hasUnsavedChanges, true, 'undo after a save is a change again');
  await doc.save(token);
  assert.equal(await sheet1Xml(uri.fsPath), originalXml, 'the original cell XML is back');
});

test('a 1,000-cell op is one journal entry and one undo step', async () => {
  const { uri, doc, events } = await open();
  const list = [];
  for (let r = 0; r < 1000; r++) list.push([r, 3, { value: num(r) }]);
  await api.pushPackageOps(uri.toString(), [cells(list)]);
  assert.equal(doc.testState.journalLength, 1);
  assert.equal(events.length, 1);
  await doc.save(token);
  const zip = await JSZip.loadAsync(fs.readFileSync(uri.fsPath));
  assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'), /<c r="D1000"[^>]*><v>999<\/v>/);
});

test('state ids: unique per edit, undo/redo move pos, a new edit truncates the redo tail', async () => {
  const { uri, doc, events } = await open();
  const s0 = api.state(uri.toString());
  assert.equal(s0.dirty, false);
  assert.equal(s0.generation, 0);
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await api.pushPackageOps(uri.toString(), [A1('two')]);
  const s2 = api.state(uri.toString());
  assert.equal(s2.stateId, 2);
  assert.equal(s2.savedStateId, 0);
  await events[1].undo();
  assert.equal(api.state(uri.toString()).stateId, 1);
  await events[1].redo();
  assert.equal(api.state(uri.toString()).stateId, 2);
  await events[1].undo();
  await api.pushPackageOps(uri.toString(), [A1('three')]);
  assert.equal(doc.testState.journalLength, 2, 'the undone edit was dropped');
  assert.equal(api.state(uri.toString()).stateId, 3);
});

test('the save point is a state id: edit, save, undo, redo = clean again', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
  await events[0].undo();
  assert.equal(doc.hasUnsavedChanges, true);
  await events[0].redo();
  assert.equal(doc.hasUnsavedChanges, false, 'redo returns to the saved state');
});

test('revert: generation + 1, journal cleared, an old undo closure does nothing and says so', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const before = api.state(uri.toString()).generation;
  await doc.revert();
  const s = api.state(uri.toString());
  assert.equal(s.generation, before + 1);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.journalLength, 0);
  const stampSha = sha(uri.fsPath);
  await events[0].undo(); // stale closure
  assert.equal(api.state(uri.toString()).stateId, s.stateId, 'nothing changed');
  assert.ok(callsOf('info').some((c) => /belongs to an earlier version/.test(c[1])));
  assert.equal(sha(uri.fsPath), stampSha);
});

test('an undo closure that does nothing while dirty adds no undo step on top of the real edits (round 3; D-067 for the empty journal)', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const stale = events[0];
  await api.pushPackageOps(uri.toString(), [A1('two')]); // journal: [one, two]
  const count = events.length;
  await stale.undo(); // 'one' is not the last applied edit: refused
  assert.equal(doc.testState.pos, 2);
  assert.equal(events.length, count, 'no no-op edit on top of the real edits: Ctrl+Z reaches them');
  assert.equal(doc.hasUnsavedChanges, true);
});

test('poison: a failing undo step poisons the document; save and backup still write the last good state', async () => {
  const { uri, doc, events } = await open();
  doc.commit({ label: 'bad', after: { ops: [A1('good')], expect: [] }, invert: () => { throw new Error('boom'); } });
  assert.equal(doc.testState.pos, 1);
  await events.at(-1).undo();
  assert.equal(doc.testState.poisoned, true);
  assert.equal(doc.testState.pos, 1, 'pos does not move');
  assert.equal(api.state(uri.toString()).poisoned, true);
  const state = doc.computeEditState();
  assert.equal(state.editable, false);
  assert.equal(state.code, 'internal');
  assert.match(state.reason, /Use Revert File/);
  await doc.save(token); // still allowed: the file is built from the journal
  assert.match(await sheet1Xml(uri.fsPath), /<c r="A1"[^>]*t="s"/);
  await doc.revert();
  assert.equal(doc.testState.poisoned, false, 'revert clears the poison');
});

// ===== save pipeline =====

test('save: a changed disk stamp asks diskChangedOnSave; Cancel keeps the document dirty and the file untouched', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.writeFileSync(uri.fsPath, Buffer.concat([fs.readFileSync(uri.fsPath), Buffer.from('x')])); // another program
  const other = sha(uri.fsPath);
  await assert.rejects(doc.save(token), (err) => err instanceof vscode.CancellationError);
  assert.equal(sha(uri.fsPath), other);
  assert.equal(doc.hasUnsavedChanges, true);
  api.answer('diskChangedOnSave', 'Overwrite');
  fs.writeFileSync(uri.fsPath, fs.readFileSync(uri.fsPath));
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('save: an external change written between verify and write is caught (step 4b)', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  host.sheetSaveHooks.afterVerify = () => {
    fs.appendFileSync(uri.fsPath, 'zzz'); // an external save in the build window
    delete host.sheetSaveHooks.afterVerify;
  };
  const other = Buffer.concat([fs.readFileSync(uri.fsPath), Buffer.from('zzz')]);
  await assert.rejects(doc.save(token), (err) => err instanceof vscode.CancellationError);
  assert.deepEqual(fs.readFileSync(uri.fsPath), other, 'nothing was written');
  assert.equal(doc.hasUnsavedChanges, true);
  assert.ok(T.logLines.some((l) => l.includes("ask('diskChangedOnSave')")));
});

test('save: an owner file ~$x.xlsx asks ownerFile; Cancel writes nothing', async () => {
  const name = unique('owned.xlsx');
  const { uri, doc } = await open(name);
  const owner = path.join(path.dirname(uri.fsPath), `~$${path.basename(uri.fsPath)}`);
  fs.writeFileSync(owner, 'x');
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const before = sha(uri.fsPath);
  await assert.rejects(doc.save(token), (err) => err instanceof vscode.CancellationError);
  assert.equal(sha(uri.fsPath), before);
  api.answer('ownerFile', 'Save');
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('save: a file open in another program (EBUSY on r+) is an error; nothing written, still dirty', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const before = sha(uri.fsPath);
  host.sheetSaveHooks.fs = {
    ...fsp,
    open: async (p, flags, mode) => {
      if (flags === 'r+') throw Object.assign(new Error('busy'), { code: 'EBUSY' });
      return fsp.open(p, flags, mode);
    },
  };
  await assert.rejects(doc.save(token), /open in another program/);
  assert.equal(sha(uri.fsPath), before);
  assert.equal(doc.hasUnsavedChanges, true);
});

test('save: files.readonlyInclude set while dirty is an error that points to Save As; still dirty', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  T.config.set('files.readonlyInclude', { '**/*.xlsx': true });
  await assert.rejects(doc.save(token), /read-only \(files\.readonlyInclude\)\. Use Save As/);
  assert.equal(doc.hasUnsavedChanges, true);
});

test('save: a failing verify writes nothing (SaveVerifyError text, edits kept)', async () => {
  const { uri, doc } = await open();
  // A cell the patcher refuses to patch in place: a patch for a sheet id that does not exist.
  await api.pushPackageOps(uri.toString(), [cells([[0, 0, { value: str('x') }]], 99)]);
  const before = sha(uri.fsPath);
  await assert.rejects(doc.save(token), /could not save .* safely/);
  assert.equal(sha(uri.fsPath), before);
  assert.equal(doc.hasUnsavedChanges, true);
});

test('save: the default is temp + rename (no temp file left), the real file is replaced', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  await doc.save(token);
  const left = fs.readdirSync(path.dirname(uri.fsPath)).filter((n) => n.startsWith('.~') && n.includes(path.basename(uri.fsPath)));
  assert.deepEqual(left, []);
});

test('save: post-write mismatch is an error, the document stays dirty, diskStamp follows the file', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  host.sheetSaveHooks.fs = {
    ...fsp,
    rename: async (from, to) => {
      await fsp.rename(from, to);
      fs.appendFileSync(to, 'x'); // the "disk" returns different bytes
    },
  };
  await assert.rejects(doc.save(token), /does not match what was written/);
  assert.equal(doc.hasUnsavedChanges, true);
  host.sheetSaveHooks.fs = fsp;
  // A retry must not ask about an external change: the stamp is the one now on disk.
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('write in place (hard link): a pre-image first; ENOSPC after half the bytes -> the pre-image is written back', async () => {
  const { uri, doc } = await open();
  fs.linkSync(uri.fsPath, `${uri.fsPath}.link`);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const before = sha(uri.fsPath);
  host.sheetSaveHooks.fs = {
    ...fsp,
    open: async (p, flags, mode) => {
      const handle = await fsp.open(p, flags, mode);
      if (flags !== 'r+' || path.basename(p) !== path.basename(uri.fsPath)) return handle;
      return new Proxy(handle, {
        get(target, key) {
          if (key === 'write') {
            return async (buf, off, len, pos) => {
              await target.write(buf, off, Math.floor(len / 2), pos);
              throw Object.assign(new Error('no space left'), { code: 'ENOSPC' });
            };
          }
          const value = target[key];
          return typeof value === 'function' ? value.bind(target) : value;
        },
      });
    },
  };
  await assert.rejects(doc.save(token), /no space left/);
  assert.equal(sha(uri.fsPath), before, 'file SHA-256 = before');
  assert.equal(doc.hasUnsavedChanges, true);
  host.sheetSaveHooks.fs = fsp;
  await doc.save(token); // a healthy retry works and updates both names
  assert.equal(sha(uri.fsPath), sha(`${uri.fsPath}.link`));
  assert.notEqual(sha(uri.fsPath), before);
  const stray = fs.readdirSync(path.dirname(uri.fsPath)).filter((n) => n.endsWith('.prev'));
  assert.deepEqual(stray, [], 'the pre-image is deleted after a good save');
});

test('write in place without any pre-image possible asks inPlaceNoBackup; Cancel writes nothing', async () => {
  const { uri } = await open();
  fs.linkSync(uri.fsPath, `${uri.fsPath}.link`);
  const before = sha(uri.fsPath);
  host.sheetSaveHooks.fs = { ...fsp, copyFile: async () => { throw Object.assign(new Error('denied'), { code: 'EACCES' }); } };
  await assert.rejects(host.writeFileSafely(uri, Buffer.from('new bytes')), (err) => err instanceof vscode.CancellationError);
  assert.equal(sha(uri.fsPath), before);
});

test('write: a missing target is created with temp + rename', async () => {
  const target = vscode.Uri.file(path.join(tmp, unique('new.bin')));
  await host.writeFileSafely(target, Buffer.from('abc'));
  assert.equal(fs.readFileSync(target.fsPath, 'utf8'), 'abc');
});

test('write: a 0600 file keeps its mode after temp + rename (POSIX)', { skip: process.platform === 'win32' }, async () => {
  const target = writeTmp(unique('mode.bin'), 'old');
  fs.chmodSync(target.fsPath, 0o600);
  await host.writeFileSafely(target, Buffer.from('new'));
  assert.equal(fs.statSync(target.fsPath).mode & 0o777, 0o600);
});

test('write: a symlink target is updated and the link stays a link (POSIX)', { skip: process.platform === 'win32' }, async () => {
  const real = writeTmp(unique('real.bin'), 'old');
  const link = path.join(tmp, unique('link.bin'));
  fs.symlinkSync(real.fsPath, link);
  await host.writeFileSafely(vscode.Uri.file(link), Buffer.from('new'));
  assert.equal(fs.lstatSync(link).isSymbolicLink(), true);
  assert.equal(fs.readFileSync(real.fsPath, 'utf8'), 'new');
});

// ===== .bak =====

test('.bak: made once before the first save, an existing .bak is never overwritten', async () => {
  const { uri, doc } = await open();
  const original = fs.readFileSync(uri.fsPath);
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await doc.save(token);
  assert.deepEqual(fs.readFileSync(`${uri.fsPath}.bak`), original);
  await api.pushPackageOps(uri.toString(), [A1('two')]);
  await doc.save(token);
  assert.deepEqual(fs.readFileSync(`${uri.fsPath}.bak`), original, 'still the oldest original');

  const second = await open('second.xlsx');
  fs.writeFileSync(`${second.uri.fsPath}.bak`, 'older');
  await api.pushPackageOps(second.uri.toString(), [A1('x')]);
  await second.doc.save(token);
  assert.equal(fs.readFileSync(`${second.uri.fsPath}.bak`, 'utf8'), 'older', 'an existing .bak is untouched');
});

test('.bak: backupOnSave = false makes no .bak', async () => {
  T.config.set('fileStudio.xlsx.backupOnSave', false);
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await doc.save(token);
  assert.equal(fs.existsSync(`${uri.fsPath}.bak`), false);
});

test('.bak: a folder named x.xlsx.bak asks bakFailed; Cancel writes nothing, Save goes on without a backup', async () => {
  const { uri, doc } = await open();
  fs.mkdirSync(`${uri.fsPath}.bak`);
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const before = sha(uri.fsPath);
  await assert.rejects(doc.save(token), (err) => err instanceof vscode.CancellationError);
  assert.equal(sha(uri.fsPath), before);
  assert.equal(doc.hasUnsavedChanges, true);
  api.answer('bakFailed', 'Save');
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.ok(fs.statSync(`${uri.fsPath}.bak`).isDirectory());
});

// ===== save as =====

test('save as: clean = exactly the base bytes; dirty = patched; the own file and the dirty state stay', async () => {
  const { uri, doc } = await open();
  const base = fs.readFileSync(uri.fsPath);
  const clean = vscode.Uri.file(path.join(tmp, unique('copy-clean.xlsx')));
  await doc.saveAs(clean, token);
  assert.deepEqual(fs.readFileSync(clean.fsPath), base);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const dirty = vscode.Uri.file(path.join(tmp, unique('copy-dirty.xlsx')));
  await doc.saveAs(dirty, token);
  assert.match(await sheet1Xml(dirty.fsPath), /<c r="A1"[^>]*t="s"/);
  assert.deepEqual(fs.readFileSync(uri.fsPath), base, 'the own file was not written');
  assert.equal(doc.hasUnsavedChanges, true);
  assert.equal(fs.existsSync(`${dirty.fsPath}.bak`), false, 'no .bak for a Save As target');
});

test('save as onto a locked target is an error and the source stays dirty', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const target = writeTmp(unique('locked.xlsx'), 'busy');
  host.sheetSaveHooks.fs = {
    ...fsp,
    open: async (p, flags, mode) => {
      if (flags === 'r+' && p === target.fsPath) throw Object.assign(new Error('busy'), { code: 'EBUSY' });
      return fsp.open(p, flags, mode);
    },
  };
  await assert.rejects(doc.saveAs(target, token), /open in another program/);
  assert.equal(fs.readFileSync(target.fsPath, 'utf8'), 'busy');
  assert.equal(doc.hasUnsavedChanges, true);
});

// ===== revert and external change =====

test('revert of a half-written file changes nothing (state, model, ids)', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const state = doc.testState;
  const model = doc.model;
  fs.writeFileSync(uri.fsPath, fs.readFileSync(uri.fsPath).subarray(0, 100)); // truncated zip
  await assert.rejects(doc.revert());
  assert.deepEqual(doc.testState, state);
  assert.equal(doc.model, model);
  assert.equal(doc.loadError, undefined);
  assert.equal(doc.hasUnsavedChanges, true);
});

test('external change while clean reloads; while dirty it warns once and keeps the edits', async () => {
  const { uri, provider, doc } = await open();
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/x.txt', 'changed')));
  assert.equal(await doc.reloadIfChangedOnDisk(), true);
  assert.equal(doc.testState.generation, 1);

  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/y.txt', 'again')));
  api.answer('externalChange', 'Keep My Edits');
  assert.equal(await doc.reloadIfChangedOnDisk(), false);
  await waitFor(() => doc.testState.diverged, 2000, 'Keep My Edits');
  assert.equal(doc.hasUnsavedChanges, true);
  assert.ok(T.logLines.some((l) => l.includes("ask('externalChange')") && l.includes('changed on disk by another program')));
  // The next save goes through the disk-changed prompt.
  await assert.rejects(doc.save(token), (err) => err instanceof vscode.CancellationError);
  assert.equal(typeof provider, 'object');
});

test('external change while dirty: "Reload from Disk" discards the edits and refreshes the panels', async () => {
  const { uri, provider, doc } = await open();
  let refreshed = 0;
  doc.onReloaded = () => refreshed++;
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/z.txt', 'z')));
  api.answer('externalChange', 'Reload from Disk (discard my edits)');
  await doc.reloadIfChangedOnDisk();
  await waitFor(() => refreshed === 1, 2000, 'reload');
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.journalLength, 0);
  assert.equal(typeof provider, 'object');
});

/** Runs the "Reload from Disk" path with one fake panel; `run(doc)` is what the fake Revert File command does. */
async function reloadWithPanel(panel, run) {
  const { uri, doc } = await open();
  doc.panels.add(panel);
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/p.txt', 'p')));
  api.answer('externalChange', 'Reload from Disk (discard my edits)');
  const original = vscode.commands.executeCommand;
  vscode.commands.executeCommand = async (id, ...args) => {
    vscode.__test.calls.push(['executeCommand', id, ...args]);
    if (id === 'workbench.action.files.revert') await run?.(doc);
  };
  try {
    await doc.reloadIfChangedOnDisk();
    await waitFor(() => doc.testState.generation >= 1 && !doc.hasUnsavedChanges, 3000, 'reload');
    await tick(150); // a second (wrong) revert would show here
  } finally {
    vscode.commands.executeCommand = original;
  }
  return { uri, doc };
}

test('"Reload from Disk" with the panel active: VS Code Revert File runs (it clears the tab\'s dirty mark) and the document is reverted once', async () => {
  const panel = { active: true, viewColumn: 1, revealed: 0, reveal() { this.revealed++; } };
  const { doc } = await reloadWithPanel(panel, (d) => d.revert());
  assert.equal(callsOf('executeCommand').filter((c) => c[1] === 'workbench.action.files.revert').length, 1);
  assert.equal(panel.revealed, 0);
  assert.equal(doc.testState.generation, 1, 'one reload by the command, none by the fallback');
  assert.equal(doc.hasUnsavedChanges, false);
});

test('"Reload from Disk" with another tab active: the panel is revealed first, then Revert File runs', async () => {
  const panel = { active: false, viewColumn: 1, revealed: 0, reveal() { this.revealed++; this.active = true; } };
  const { doc } = await reloadWithPanel(panel, (d) => d.revert());
  assert.equal(panel.revealed, 1);
  assert.equal(callsOf('executeCommand').filter((c) => c[1] === 'workbench.action.files.revert').length, 1);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('"Reload from Disk" when the panel never becomes active: no command is sent (it could revert another file), the document reverts itself', async () => {
  const panel = { active: false, viewColumn: 1, revealed: 0, reveal() { this.revealed++; } };
  const { doc } = await reloadWithPanel(panel, undefined);
  assert.equal(panel.revealed, 1);
  assert.equal(callsOf('executeCommand').filter((c) => c[1] === 'workbench.action.files.revert').length, 0);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.generation, 1);
});

test('"Reload from Disk" when the Revert File command does nothing: the document reverts itself', async () => {
  const panel = { active: true, viewColumn: 1, reveal() {} };
  const { doc } = await reloadWithPanel(panel, async () => undefined);
  assert.equal(callsOf('executeCommand').filter((c) => c[1] === 'workbench.action.files.revert').length, 1);
  assert.equal(doc.hasUnsavedChanges, false);
  assert.equal(doc.testState.generation, 1);
});

// ===== hot exit =====

test('backup writes a verified workbook + sidecar; restore opens it dirty with lossyAccepted and fires one content event', async () => {
  const chart = await xlsxBytes((zip) => zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'));
  const { uri, provider, doc } = await open('chart.xlsx', chart);
  assert.deepEqual([...doc.lossyFeatures], ['Charts']);
  doc.acceptLossy();
  await api.pushPackageOps(uri.toString(), [A1('unsaved')]);
  const dest = vscode.Uri.file(path.join(tmp, unique('backup.xlsx')));
  const backup = await doc.backup(dest, token);
  const sidecar = JSON.parse(fs.readFileSync(`${dest.fsPath}.json`, 'utf8'));
  assert.equal(sidecar.v, 1);
  assert.equal(sidecar.lossyAccepted, true);
  assert.equal(sidecar.divergedFromDisk, false);
  assert.ok(sidecar.sourceStamp.size > 0);
  assert.match(await sheet1Xml(dest.fsPath), /<c r="A1"[^>]*t="s"/);

  const restoredEvents = [];
  const restored = provider.openCustomDocument(uri, { backupId: backup.id }, token);
  restored.onDidChange((e) => {
    if (typeof e.undo === 'function') restoredEvents.push(e);
  });
  await restored.loading;
  await tick(5);
  assert.equal(restored.hasUnsavedChanges, true, 'dirty (diverged from disk)');
  assert.equal(restored.testState.diverged, true);
  assert.equal(restored.testState.journalLength, 0);
  assert.equal(restoredEvents.length, 1, 'one content event keeps VS Code dirty');
  assert.equal(restoredEvents[0].label, 'Unsaved changes');
  assert.equal(restored.computeEditState().code, undefined, 'lossy accepted from the sidecar');

  // restore -> edit -> undo -> the no-op's undo re-fires, so VS Code cannot reach its "clean" point
  await restored.queue.run(() => restored.commit({ label: 'e', after: { ops: [A1('more')], expect: [] } }));
  const edit = restoredEvents.at(-1);
  await edit.undo();
  await restoredEvents[0].undo();
  assert.ok(restoredEvents.length >= 3, 'the no-op undo fired another no-op');

  await backup.delete();
  assert.equal(fs.existsSync(dest.fsPath), false);
  assert.equal(fs.existsSync(`${dest.fsPath}.json`), false);
});

test('backup: unchanged state reuses the bytes; a changed state builds again', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const a = vscode.Uri.file(path.join(tmp, unique('b1.xlsx')));
  const b = vscode.Uri.file(path.join(tmp, unique('b2.xlsx')));
  await doc.backup(a, token);
  await doc.backup(b, token);
  assert.deepEqual(fs.readFileSync(a.fsPath), fs.readFileSync(b.fsPath));
});

test('backup whose verify fails (test hook) writes no file and throws', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  host.sheetSaveHooks.failBackupVerify = true;
  const dest = vscode.Uri.file(path.join(tmp, unique('bad-backup.xlsx')));
  await assert.rejects(doc.backup(dest, token), /could not save .* safely/);
  assert.equal(fs.existsSync(dest.fsPath), false);
  assert.equal(fs.existsSync(`${dest.fsPath}.json`), false);
  assert.ok(T.logLines.some((l) => /Could not build/.test(l)));
});

test('an edit pushed while a backup builds is not held up by it (the snapshot is the only queue step)', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  const dest = vscode.Uri.file(path.join(tmp, unique('b3.xlsx')));
  const pending = doc.backup(dest, token);
  await api.pushPackageOps(uri.toString(), [A1('two')]);
  await pending;
  assert.equal(doc.testState.journalLength, 2);
});

test('restore with a missing sidecar: lossy not accepted, and the first save asks diskChangedOnSave', async () => {
  const { uri, provider, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('unsaved')]);
  const dest = vscode.Uri.file(path.join(tmp, unique('nosidecar.xlsx')));
  const backup = await doc.backup(dest, token);
  fs.rmSync(`${dest.fsPath}.json`);
  const restored = provider.openCustomDocument(uri, { backupId: backup.id }, token);
  await restored.loading;
  assert.equal(restored.hasUnsavedChanges, true);
  const before = sha(uri.fsPath);
  await assert.rejects(restored.save(token), (err) => err instanceof vscode.CancellationError);
  assert.equal(sha(uri.fsPath), before);
  api.answer('diskChangedOnSave', 'Overwrite');
  await restored.save(token);
  assert.equal(restored.hasUnsavedChanges, false);
  assert.match(await sheet1Xml(uri.fsPath), /<c r="A1"[^>]*t="s"/);
});

test('restore of corrupt backup bytes keeps a copy under restore-failed/, opens the file from disk and tells the user', async () => {
  const { uri, provider } = await open();
  const dest = writeTmp(unique('corrupt.xlsx'), 'PK this is not a workbook');
  const storage = path.join(tmp, unique('storage'));
  const restored = provider.openCustomDocument(uri, { backupId: dest.toString() }, token);
  // The provider has no storage folder here: the copy goes to the temp folder; use a provider with one instead.
  await restored.loading;
  const withStorage = new host.SheetEditorProvider(extensionUri, vscode.Uri.file(storage));
  T.calls.length = 0;
  const doc = withStorage.openCustomDocument(uri, { backupId: dest.toString() }, token);
  await doc.loading;
  assert.equal(doc.hasUnsavedChanges, false, 'opened from disk, clean');
  assert.equal(doc.loadError, undefined);
  const copies = fs.readdirSync(path.join(storage, 'restore-failed'));
  assert.equal(copies.length, 1);
  assert.equal(fs.readFileSync(path.join(storage, 'restore-failed', copies[0]), 'utf8'), 'PK this is not a workbook');
  assert.ok(callsOf('warning').some((c) => /could not restore the unsaved changes/.test(c[1]) && c[1].includes('restore-failed')));
});

// ===== hard read-only scan, edit state =====

test('a sheet other than the edited one with rows out of order makes the document hard read-only', async () => {
  const bytes = await xlsxBytes(async (zip) => {
    const rows = '<row r="2"><c r="A2"><v>1</v></c></row><row r="1"><c r="A1"><v>2</v></c></row>';
    const text = await zip.file('xl/worksheets/sheet2.xml').async('string');
    zip.file('xl/worksheets/sheet2.xml', text.replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${rows}</sheetData>`));
  });
  const { doc } = await open('order.xlsx', bytes);
  await doc.ensureScan();
  assert.match(doc.gateState.block, /FileStudio cannot edit this file safely: .*rows/i);
  const state = doc.computeEditState();
  assert.equal(state.editable, false);
  assert.equal(state.code, 'blocked');
  // Save As of the clean, hard read-only file still writes the original bytes.
  const copy = vscode.Uri.file(path.join(tmp, unique('order-copy.xlsx')));
  await doc.saveAs(copy, token);
  assert.deepEqual(fs.readFileSync(copy.fsPath), bytes);
});

test('lossy workbook: editState code lossy with canEditAnyway; acceptLossy clears it; revert keeps it only for the same list', async () => {
  const chart = await xlsxBytes((zip) => zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'));
  const { uri, doc } = await open('lossy.xlsx', chart);
  let state = doc.computeEditState();
  assert.equal(state.code, 'lossy');
  assert.ok(state.canEditAnyway.details.length > 0);
  doc.acceptLossy();
  assert.equal(doc.gateState.lossy.accepted, true);
  await doc.revert();
  assert.equal(doc.gateState.lossy.accepted, true, 'same list: kept');
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('xl/embeddings/o.bin', 'x')));
  await doc.revert();
  assert.equal(doc.gateState.lossy.accepted, false, 'another list: asked again');
  state = doc.computeEditState();
  assert.equal(state.code, 'lossy');
});

test('a plain workbook stays view-only (xlsx editing waits for P2) with no banner or message', async () => {
  const { doc } = await open();
  assert.deepEqual(doc.computeEditState(), { editable: false, kinds: [] });
});

test('the version rule bookkeeping: structural edits raise structureVersion, value edits do not', async () => {
  const { uri, doc } = await open();
  await api.pushPackageOps(uri.toString(), [A1('v')]);
  assert.deepEqual(doc.versions, { version: 1, structureVersion: 0 });
  await doc.queue.run(() => doc.commit({ label: 's', structural: true, after: { ops: [], expect: [] } }));
  assert.deepEqual(doc.versions, { version: 2, structureVersion: 2 });
  await doc.revert();
  assert.equal(doc.versions.structureVersion, doc.versions.version);
});

test('kill switch off while dirty: save, undo and backup still work (D-075)', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  T.config.set('fileStudio.editing.enabled', false);
  await events[0].undo();
  assert.equal(doc.hasUnsavedChanges, false);
  await events[0].redo();
  assert.equal(doc.hasUnsavedChanges, true);
  await doc.backup(vscode.Uri.file(path.join(tmp, unique('ks.xlsx'))), token);
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
});

test('test API: pushPackageOps and lastSaveReport exist; state() works for a document without a panel', async () => {
  const { uri, doc } = await open();
  assert.equal(typeof api.pushPackageOps, 'function');
  assert.equal(api.lastSaveReport(uri.toString()), undefined);
  await api.pushPackageOps(uri.toString(), [A1('x')], 'My label');
  assert.equal(api.state(uri.toString()).dirty, true);
  await doc.save(token);
  assert.ok(api.lastSaveReport(uri.toString()).timings.total >= 0);
  assert.equal(api.state(uri.toString()).dirty, false);
});

// ===== stage b4: autosave, restore then save, "Keep My Edits" event =====

/** The shared strings of the file, in order (for the cells a test wrote). */
async function sharedStringsOf(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file('xl/sharedStrings.xml').async('string');
  return [...xml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((m) => m[1]);
}

test('autosave pattern: 20 cells typed while saves and backups run; none is lost and the file ends with all 20', async () => {
  const { uri, doc } = await open();
  const dest = vscode.Uri.file(path.join(tmp, unique('auto.xlsx')));
  const running = [];
  for (let i = 0; i < 20; i++) {
    await api.pushPackageOps(uri.toString(), [cells([[20 + i, 0, { value: str(`typed ${i}`) }]])]); // accepted at once
    if (i % 3 === 0) running.push(doc.save(token)); // files.autoSave: afterDelay fires in the middle of the typing
    if (i % 4 === 1) running.push(doc.backup(dest, token).then(() => undefined)); // and so does the hot-exit backup
    await tick(2);
  }
  await Promise.all(running);
  assert.equal(doc.testState.journalLength, 20, 'every edit was accepted');
  await doc.save(token); // the last autosave
  assert.equal(doc.hasUnsavedChanges, false);
  const strings = await sharedStringsOf(uri.fsPath);
  for (let i = 0; i < 20; i++) assert.ok(strings.includes(`typed ${i}`), `typed ${i} is in the saved file`);
  assert.ok(fs.existsSync(`${uri.fsPath}.bak`), 'the .bak was made once, before the first save');
  assert.equal(callsOf('error').length, 0, 'no error shown');
});

test('restore -> new edit -> save: the restored cell and the new cell are both written; the original goes to .bak; no prompt', async () => {
  const original = await xlsxBytes();
  const { uri, provider, doc } = await open('restore-save.xlsx', original);
  await api.pushPackageOps(uri.toString(), [A1('restored cell')]);
  const dest = vscode.Uri.file(path.join(tmp, unique('rs-backup.xlsx')));
  const backup = await doc.backup(dest, token);
  const restored = provider.openCustomDocument(uri, { backupId: backup.id }, token);
  await restored.loading;
  assert.equal(sha(uri.fsPath), crypto.createHash('sha256').update(original).digest('hex'), 'the file on disk was not touched');
  await restored.queue.run(() => restored.commit({ label: 'new', after: { ops: [cells([[5, 0, { value: str('new cell') }]])], expect: [] } }));
  T.calls.length = 0;
  await restored.save(token); // no answer queued: a prompt would be a Cancel and reject
  assert.equal(restored.hasUnsavedChanges, false);
  const strings = await sharedStringsOf(uri.fsPath);
  assert.ok(strings.includes('restored cell') && strings.includes('new cell'));
  assert.deepEqual(fs.readFileSync(`${uri.fsPath}.bak`), original, 'the .bak holds the file as it was before the first save');
});

test('a second hot exit: the backup of a restored document (no new edits) still holds the unsaved cell and stays diverged', async () => {
  const { uri, provider, doc } = await open('twice.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('survives two restarts')]);
  const first = vscode.Uri.file(path.join(tmp, unique('twice-1.xlsx')));
  const backup1 = await doc.backup(first, token);
  const restored1 = provider.openCustomDocument(uri, { backupId: backup1.id }, token);
  await restored1.loading;
  const second = vscode.Uri.file(path.join(tmp, unique('twice-2.xlsx')));
  const backup2 = await restored1.backup(second, token);
  assert.equal(JSON.parse(fs.readFileSync(`${second.fsPath}.json`, 'utf8')).divergedFromDisk, true);
  const restored2 = provider.openCustomDocument(uri, { backupId: backup2.id }, token);
  await restored2.loading;
  assert.equal(restored2.hasUnsavedChanges, true);
  await restored2.save(token);
  assert.ok((await sharedStringsOf(uri.fsPath)).includes('survives two restarts'));
});

test('"Keep My Edits" keeps VS Code dirty without an undo step on top of the real edits (round 3)', async () => {
  const { uri, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  fs.writeFileSync(uri.fsPath, await xlsxBytes((zip) => zip.file('docProps/k.txt', 'k')));
  const count = events.length;
  api.answer('externalChange', 'Keep My Edits');
  await doc.reloadIfChangedOnDisk();
  await waitFor(() => doc.testState.diverged, 2000, 'Keep My Edits');
  assert.equal(events.length, count, 'no undo / redo step is added');
  assert.equal(doc.hasUnsavedChanges, true);
  assert.equal(doc.testState.journalLength, 1, 'the journal did not change');
});

test('a failed Revert File (half-written file) keeps VS Code dirty: one no-op edit while the document still has its edit', async () => {
  const { uri, provider, doc, events } = await open();
  await api.pushPackageOps(uri.toString(), [A1('mine')]);
  const state = doc.testState;
  fs.writeFileSync(uri.fsPath, fs.readFileSync(uri.fsPath).subarray(0, 100)); // truncated zip
  const count = events.length;
  await assert.rejects(provider.revertCustomDocument(doc, token), /not a valid \.xlsx workbook/);
  assert.deepEqual(doc.testState, state, 'state unchanged');
  assert.equal(events.length, count + 1, 'the tab stays dirty (VS Code clears its mark when Revert File was asked)');
  assert.equal(events.at(-1).label, 'Unsaved changes');
  // A clean document has nothing to keep: no event.
  const clean = await open('clean-half.xlsx');
  fs.writeFileSync(clean.uri.fsPath, fs.readFileSync(clean.uri.fsPath).subarray(0, 100));
  await assert.rejects(clean.provider.revertCustomDocument(clean.doc, token));
  assert.equal(clean.events.length, 0);
});

// ===== size limit (0.2.0: editing very large workbooks comes in a later version) =====

/** A workbook whose used range is `rows` x `cols` cells (one far corner cell; small file). */
async function sizedXlsx(rows, cols) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell(rows, cols).value = 1;
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

test('size limit: a workbook just under the limit still edits and saves', async () => {
  const { uri, doc } = await open('under.xlsx', await sizedXlsx(1000, 100)); // 100,000 cells
  await doc.ensureScan();
  assert.equal(doc.gateState.block, undefined);
  await api.pushPackageOps(uri.toString(), [cells([[0, 0, { value: str('changed') }]])]);
  await doc.save(token);
  assert.equal(doc.hasUnsavedChanges, false);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(uri.fsPath);
  assert.equal(wb.getWorksheet('Data').getCell('A1').value, 'changed');
});

test('size limit: a workbook over the limit is hard read-only and nothing can write it', async () => {
  const bytes = await sizedXlsx(1001, 100); // 100,100 cells
  const { uri, doc, events } = await open('over.xlsx', bytes);
  assert.match(doc.gateState.block, /too large to edit safely in this version/);
  assert.match(doc.gateState.block, /later version/);
  const state = doc.computeEditState();
  assert.equal(state.editable, false);
  assert.equal(state.code, 'blocked');
  await assert.rejects(() => api.pushPackageOps(uri.toString(), [cells([[0, 0, { value: str('x') }]])]), /too large to edit safely/);
  assert.equal(events.length, 0);
  assert.equal(doc.hasUnsavedChanges, false);
  await doc.save(token); // nothing to save: the file is not written
  assert.deepEqual(fs.readFileSync(uri.fsPath), bytes);
  await doc.revert(); // a reload keeps it read-only
  assert.match(doc.gateState.block, /too large/);
});

test('size limit: a small zip that unpacks to a huge sheet is read-only even with few cells (round 3: counted unpacked)', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  for (let r = 1; r <= 1000; r++) sheet.getCell(r, 1).value = `${r}-${'x'.repeat(30000)}`; // about 30 MB of text, compresses to almost nothing
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  assert.ok(bytes.length < 1024 * 1024, `fixture is ${bytes.length} bytes`);
  const { doc } = await open('bytes.xlsx', bytes);
  assert.match(doc.gateState.block, /too large to edit safely/);
});

test('F-R3: a too-large workbook with a chart is blocked, with no "Edit anyway" offer', async () => {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(await sizedXlsx(1001, 100));
  zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
  const { doc } = await open('big-chart.xlsx', await zip.generateAsync({ type: 'nodebuffer' }));
  assert.ok(doc.lossyFeatures.length > 0, 'fixture is lossy');
  const state = doc.computeEditState();
  assert.equal(state.code, 'blocked');
  assert.equal(state.canEditAnyway, undefined);
});
