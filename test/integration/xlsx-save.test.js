// xlsx save path of 0.2.0 P1 in a REAL VS Code (contract section 14, P1 acceptance, integration part). Edits are pushed with
// the test API (`pushPackageOps`: one journal entry, no edit UI yet); saving, undo, revert and the prompts then go through
// VS Code's own commands. Every test works on a temporary copy of a sample (the workspace itself is a temp copy of
// test/samples). Prompts are answered with `api.answer(id, button)`; no answer = Cancel.
// Not here (host tests in test/host/sheet-document.test.js cover them with a fake fs and hooks, the gate runs in the workshop):
// Save As (opens a file dialog), hot-exit RESTORE over a restart (a test window has no persistent backups: VS Code uses an
// in-memory backup service there, probed in stage b4 with a two-run harness: Backups/ stayed empty; the backup FILES that
// VS Code asks for are checked below), test hooks that live inside the bundle (pause after verify, failing backup verify),
// ENOSPC, large.xlsx perf, Excel locks.
'use strict';
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const vscode = require('vscode');
const h = require('./helpers');

const sha = (uri) => crypto.createHash('sha256').update(fs.readFileSync(uri.fsPath)).digest('hex');
const A1 = (text) => ({ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: text } }]] });
const sheet1Xml = async (uri) => (await JSZip.loadAsync(fs.readFileSync(uri.fsPath))).file('xl/worksheets/sheet1.xml').async('string');

let api;
let counter = 0;

function safeState(uri) {
  try {
    return api.state(uri.toString());
  } catch {
    return undefined;
  }
}

/** Copies a sample to a new unique name, opens it with FileStudio and returns { uri, key, name }. */
async function openCopy(sample = 'styled.xlsx') {
  const name = `it-save-${++counter}-${sample.split('/').pop()}`;
  const uri = h.copyFile(sample, name);
  await h.openWith(name);
  await h.waitFor(() => safeState(uri), { message: `the document ${name}` });
  return { uri, key: uri.toString(), name };
}

/** Pushes an edit and waits until VS Code shows the tab as dirty. */
async function edit(ctx, text = 'edited by test') {
  await api.pushPackageOps(ctx.key, [A1(text)], 'Test edit');
  await h.waitFor(() => h.isDirty(ctx.uri), { message: 'the tab to become dirty' });
}

/**
 * The hot-exit backup VS Code asked FileStudio for (backupCustomDocument): a workbook plus a '.json' sidecar in the
 * extension's workspace storage. Resolves with { file, sidecar } of the one whose shared strings contain `text`, or undefined.
 */
async function findBackup(text) {
  const root = path.join(process.env.FS_IT_USER_DATA_DIR, 'User', 'workspaceStorage');
  const sidecars = [];
  try {
    for (const space of fs.readdirSync(root)) {
      const dir = path.join(root, space, 'Yokeeswaran.filestudio');
      if (fs.existsSync(dir)) sidecars.push(...fs.readdirSync(dir).filter((n) => n.endsWith('.json')).map((n) => path.join(dir, n)));
    }
  } catch {
    return undefined;
  }
  for (const sidecar of sidecars) {
    const file = sidecar.slice(0, -'.json'.length);
    try {
      const zip = await JSZip.loadAsync(fs.readFileSync(file));
      const strings = zip.file('xl/sharedStrings.xml') ? await zip.file('xl/sharedStrings.xml').async('string') : '';
      if (strings.includes(`>${text}<`)) return { file, sidecar };
    } catch {
      /* not there yet, half-written, or another document's backup */
    }
  }
  return undefined;
}

const clean = (ctx, message = 'the tab to be clean') => h.waitFor(() => !h.isDirty(ctx.uri), { message });

suite('xlsx save path (P1)', () => {
  suiteSetup(async () => {
    const ext = h.extension();
    assert.ok(ext, `extension ${h.EXTENSION_ID} is not loaded`);
    api = await ext.activate();
    assert.ok(api && typeof api.pushPackageOps === 'function', 'the test API (pushPackageOps) is available in test mode');
    await h.closeAll();
  });
  teardown(() => h.closeAll());

  test('pushPackageOps -> Save: the cell is patched, the tab is clean, a .bak holds the original, the report names the part', async () => {
    const ctx = await openCopy();
    const original = fs.readFileSync(ctx.uri.fsPath);
    const xmlBefore = await sheet1Xml(ctx.uri);
    await edit(ctx);
    assert.equal(api.state(ctx.key).dirty, true);
    await h.save();
    await clean(ctx);
    assert.notEqual(await sheet1Xml(ctx.uri), xmlBefore);
    assert.match(await sheet1Xml(ctx.uri), /<c r="A1"[^>]*t="s"/);
    assert.deepEqual(fs.readFileSync(`${ctx.uri.fsPath}.bak`), original, '.bak = the original bytes');
    const report = api.lastSaveReport(ctx.key);
    assert.ok(report && report.changedParts.includes('xl/worksheets/sheet1.xml'), 'lastSaveReport lists sheet1.xml');
    const state = api.state(ctx.key);
    assert.equal(state.stateId, state.savedStateId);
  });

  test('edit + undo + Save: nothing is written (SHA and mtime unchanged, no .bak)', async () => {
    const ctx = await openCopy();
    const before = sha(ctx.uri);
    const mtime = fs.statSync(ctx.uri.fsPath).mtimeMs;
    await edit(ctx);
    await h.undo();
    // FileStudio is back on the saved state; the tab stays dirty on purpose (the sticky changed flag of D-F1: VS Code cannot
    // tell the saved state from a new edit after save -> undo -> edit). Save then writes nothing and cleans the tab.
    await h.waitFor(() => api.state(ctx.key).dirty === false, { message: 'FileStudio clean after undo' });
    await h.save();
    await clean(ctx, 'the tab to be clean after the (empty) save');
    await h.sleep(500);
    assert.equal(sha(ctx.uri), before);
    assert.equal(fs.statSync(ctx.uri.fsPath).mtimeMs, mtime);
    assert.equal(fs.existsSync(`${ctx.uri.fsPath}.bak`), false);
  });

  test('F1: edit A, Save, Undo, edit B: the tab is dirty (closing it asks), Save writes B', async () => {
    const ctx = await openCopy();
    await edit(ctx, 'A edit');
    await h.save();
    await clean(ctx, 'clean after save');
    await h.undo();
    await api.pushPackageOps(ctx.key, [A1('B edit')], 'Test edit B');
    await h.sleep(500);
    assert.equal(api.state(ctx.key).dirty, true, 'FileStudio has unsaved work');
    assert.equal(h.isDirty(ctx.uri), true, 'the tab says so too (no silent loss on close)');
    await h.save();
    await clean(ctx, 'clean after saving B');
    const sst = await (await JSZip.loadAsync(fs.readFileSync(ctx.uri.fsPath))).file('xl/sharedStrings.xml').async('string');
    assert.ok(sst.includes('>B edit<'), 'edit B is in the file');
  });

  test('edit + Save + undo + Save: the original cell XML is back', async () => {
    const ctx = await openCopy();
    const xmlBefore = await sheet1Xml(ctx.uri);
    await edit(ctx);
    await h.save();
    await clean(ctx, 'clean after save');
    await h.undo();
    await h.waitFor(() => h.isDirty(ctx.uri), { message: 'dirty after undo of a saved edit' });
    await h.save();
    await clean(ctx, 'clean after the second save');
    assert.equal(await sheet1Xml(ctx.uri), xmlBefore);
  });

  test('Revert: the edits are gone, the file is untouched, an old undo step does nothing', async () => {
    const ctx = await openCopy();
    const before = sha(ctx.uri);
    await edit(ctx);
    const generation = api.state(ctx.key).generation;
    await h.revert();
    await clean(ctx, 'the tab to be clean after revert');
    assert.equal(sha(ctx.uri), before);
    await h.waitFor(() => api.state(ctx.key).generation > generation, { message: 'generation + 1 after revert (state ' + JSON.stringify(api.state(ctx.key)) + ')' });
    await h.undo(); // the closure of the discarded edit: a no-op that says so
    await h.sleep(300);
    assert.equal(sha(ctx.uri), before);
    assert.equal(api.state(ctx.key).stateId, api.state(ctx.key).savedStateId);
  });

  test('an existing .bak is never overwritten', async () => {
    const ctx = await openCopy();
    fs.writeFileSync(`${ctx.uri.fsPath}.bak`, 'older backup');
    await edit(ctx);
    await h.save();
    await clean(ctx, 'clean after save');
    assert.equal(fs.readFileSync(`${ctx.uri.fsPath}.bak`, 'utf8'), 'older backup');
  });

  test('a folder named x.xlsx.bak: the prompt appears, Cancel writes nothing and the tab stays dirty', async () => {
    const ctx = await openCopy();
    fs.mkdirSync(`${ctx.uri.fsPath}.bak`);
    const before = sha(ctx.uri);
    await edit(ctx);
    await h.save(); // no answer = Cancel
    await h.sleep(800);
    assert.equal(sha(ctx.uri), before);
    assert.equal(h.isDirty(ctx.uri), true);
    api.answer('bakFailed', 'Save');
    await h.save();
    await clean(ctx, 'clean after Save without a backup');
    assert.notEqual(sha(ctx.uri), before);
  });

  test('the disk changed after opening: Save asks, Cancel keeps the tab dirty and the file as the other program left it', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    fs.writeFileSync(ctx.uri.fsPath, Buffer.concat([fs.readFileSync(ctx.uri.fsPath), Buffer.from('x')]));
    const other = sha(ctx.uri);
    await h.save(); // Cancel
    await h.sleep(800);
    assert.equal(sha(ctx.uri), other);
    assert.equal(h.isDirty(ctx.uri), true);
  });

  test('an owner file ~$x.xlsx next to the workbook: Save asks, Cancel writes nothing', async () => {
    const ctx = await openCopy();
    const owner = path.join(path.dirname(ctx.uri.fsPath), `~$${path.basename(ctx.uri.fsPath)}`);
    fs.writeFileSync(owner, 'x');
    try {
      const before = sha(ctx.uri);
      await edit(ctx);
      await h.save(); // Cancel
      await h.sleep(800);
      assert.equal(sha(ctx.uri), before);
      assert.equal(h.isDirty(ctx.uri), true);
    } finally {
      fs.rmSync(owner, { force: true });
    }
  });

  test('a read-only file: Save is an error, the file is unchanged and the tab stays dirty', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    const before = sha(ctx.uri);
    fs.chmodSync(ctx.uri.fsPath, 0o444);
    try {
      await h.save().then(undefined, () => undefined); // the error is shown to the user; the command may reject
      await h.sleep(800);
      assert.equal(sha(ctx.uri), before);
      assert.equal(h.isDirty(ctx.uri), true);
    } finally {
      fs.chmodSync(ctx.uri.fsPath, 0o666);
    }
  });

  test('a hard link: the file is written in place, both names show the new bytes', async () => {
    const ctx = await openCopy();
    const link = h.fileUri(`it-save-link-${counter}.xlsx`);
    try {
      fs.linkSync(ctx.uri.fsPath, link.fsPath);
    } catch {
      return; // the file system has no hard links: nothing to test
    }
    try {
      await edit(ctx);
      await h.save();
      await clean(ctx, 'clean after save');
      assert.equal(sha(link), sha(ctx.uri), 'both names have the same new bytes');
      assert.match(await sheet1Xml(link), /<c r="A1"[^>]*t="s"/);
    } finally {
      fs.rmSync(link.fsPath, { force: true });
    }
  });

  test('a symlink: the target is updated and the link stays a link', async () => {
    const ctx = await openCopy();
    await h.closeAll();
    const linkName = `it-save-sym-${counter}.xlsx`;
    const link = h.fileUri(linkName);
    try {
      fs.symlinkSync(ctx.uri.fsPath, link.fsPath);
    } catch {
      return; // no permission to make symlinks (normal on Windows without developer mode)
    }
    try {
      await h.openWith(linkName);
      await h.waitFor(() => safeState(link), { message: 'the document of the link' });
      await api.pushPackageOps(link.toString(), [A1('through link')]);
      await h.waitFor(() => h.isDirty(link), { message: 'dirty' });
      await h.save();
      await h.waitFor(() => !h.isDirty(link), { message: 'clean' });
      assert.ok(fs.lstatSync(link.fsPath).isSymbolicLink(), 'the link is still a link');
      assert.match(await sheet1Xml(ctx.uri), /<c r="A1"[^>]*t="s"/);
    } finally {
      fs.rmSync(link.fsPath, { force: true });
    }
  });

  test('generated fixtures: one edit + Save changes only the expected parts and logs no problem', { timeout: 180_000 }, async () => {
    const fixtures = fs.readdirSync(h.fileUri('xlsx-save').fsPath).filter((n) => n.endsWith('.xlsx'));
    assert.ok(fixtures.length >= 4, 'the generated fixtures are in the workspace');
    for (const fixture of fixtures) {
      const ctx = await openCopy(`xlsx-save/${fixture}`);
      const before = await JSZip.loadAsync(fs.readFileSync(ctx.uri.fsPath));
      const mark = h.logMark();
      // objects.xlsx has a table at A1:D9 (F4: a header cell edit is refused), so its edit goes below the table.
      await api.pushPackageOps(ctx.key, [fixture === 'objects.xlsx' ? { kind: 'cells', sheet: 1, cells: [[20, 0, { value: { t: 's', v: 'edited by test' } }]] } : A1('edited by test')], 'Test edit');
      await h.waitFor(() => h.isDirty(ctx.uri), { message: 'the tab to become dirty' });
      await h.save();
      await clean(ctx, `clean after saving ${fixture}`);
      const after = await JSZip.loadAsync(fs.readFileSync(ctx.uri.fsPath));
      const differing = [];
      for (const name of Object.keys(before.files)) {
        if (before.files[name].dir) continue;
        const a = await before.file(name).async('nodebuffer');
        const b = after.file(name) ? await after.file(name).async('nodebuffer') : undefined;
        if (!b || !a.equals(b)) differing.push(name);
      }
      assert.ok(differing.includes('xl/worksheets/sheet1.xml'), `${fixture}: sheet1.xml changed`);
      const allowed = new Set(['xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml', 'xl/workbook.xml', 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']);
      assert.deepEqual(differing.filter((n) => !allowed.has(n)), [], `${fixture}: only the expected parts differ`);
      h.assertNoLogProblems(h.logSince(mark), fixture);
      await h.closeAll();
    }
  });

  // Contract 4.2: VS Code's Revert File ignores a Uri for a custom editor (probe, decisions.md 29) and works on the ACTIVE
  // editor. "Reload from Disk" therefore reveals the panel first; another dirty tab must never lose its edits.
  test('"Reload from Disk" of a background tab reverts only that document; the active dirty tab keeps its edits', async () => {
    const a = await openCopy();
    await edit(a);
    const generationA = api.state(a.key).generation;
    const b = await openCopy(); // b is now the active tab
    await edit(b, 'edit in the other tab');
    const stateB = api.state(b.key);
    const other = fs.readFileSync(h.fileUri('xlsx-save/formulas.xlsx').fsPath);
    api.answer('externalChange', 'Reload from Disk (discard my edits)');
    fs.writeFileSync(a.uri.fsPath, other);
    await h.waitFor(() => api.state(a.key).generation > generationA, { timeout: 20_000, message: 'the reload of the first document' });
    await clean(a, 'the first tab to be clean after "Reload from Disk"');
    assert.equal(h.isDirty(b.uri), true, 'the other tab is still dirty');
    assert.equal(api.state(b.key).generation, stateB.generation);
    assert.equal(api.state(b.key).stateId, stateB.stateId, 'the other document keeps its edit');
    assert.equal(sha(a.uri), crypto.createHash('sha256').update(other).digest('hex'));
  });

  test('the file changed on disk while dirty: "Reload from Disk" discards the edits, reloads, and the tab becomes clean', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    const generation = api.state(ctx.key).generation;
    const other = fs.readFileSync(h.fileUri('xlsx-save/formulas.xlsx').fsPath); // a different, valid workbook
    api.answer('externalChange', 'Reload from Disk (discard my edits)');
    fs.writeFileSync(ctx.uri.fsPath, other);
    await h.waitFor(() => api.state(ctx.key).generation > generation, { timeout: 20_000, message: 'the reload after the external change' });
    assert.equal(api.state(ctx.key).stateId, api.state(ctx.key).savedStateId);
    assert.equal(sha(ctx.uri), crypto.createHash('sha256').update(other).digest('hex'), 'the other program\'s file is untouched');
    await clean(ctx, 'the tab to be clean after "Reload from Disk"');
  });

  test('the file changed on disk while dirty: "Keep My Edits" keeps the tab dirty; Save then asks and Overwrite writes the edits', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    const other = fs.readFileSync(h.fileUri('xlsx-save/formulas.xlsx').fsPath);
    api.answer('externalChange', 'Keep My Edits');
    fs.writeFileSync(ctx.uri.fsPath, other);
    await h.sleep(2500); // the watcher is debounced; the warning then answers "Keep My Edits"
    assert.equal(h.isDirty(ctx.uri), true);
    assert.deepEqual(fs.readFileSync(ctx.uri.fsPath), other, 'nothing was written yet');
    api.answer('diskChangedOnSave', 'Overwrite');
    await h.save();
    await clean(ctx, 'the tab to be clean after Overwrite');
    assert.match(await sheet1Xml(ctx.uri), /<c r="A1"[^>]*t="s"/);
  });

  test('R3: after "Keep My Edits" Ctrl+Z still undoes the real edits (no no-op step on top)', async () => {
    const ctx = await openCopy();
    await api.pushPackageOps(ctx.key, [A1('one')], 'e1');
    await api.pushPackageOps(ctx.key, [{ kind: 'cells', sheet: 1, cells: [[0, 1, { value: { t: 's', v: 'two' } }]] }], 'e2');
    await h.waitFor(() => h.isDirty(ctx.uri), { message: 'dirty' });
    api.answer('externalChange', 'Keep My Edits');
    fs.writeFileSync(ctx.uri.fsPath, fs.readFileSync(h.fileUri('xlsx-save/formulas.xlsx').fsPath));
    await h.sleep(2500);
    const before = api.state(ctx.key).stateId;
    await h.undo();
    await h.waitFor(() => api.state(ctx.key).stateId !== before, { timeout: 5000, message: 'the undo to reach the real edits' });
    assert.equal(h.isDirty(ctx.uri), true, 'the tab is still dirty (one edit left, and the file differs)');
  });

  // ----- stage b4: settings that change while a document is dirty, a damaged file, autosave -----

  test('files.readonlyInclude set while dirty: Save is an error, the file is unchanged and the tab stays dirty', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    const before = sha(ctx.uri);
    const files = vscode.workspace.getConfiguration('files');
    await files.update('readonlyInclude', { ['**/' + ctx.name]: true }, vscode.ConfigurationTarget.Workspace);
    try {
      // VS Code shows the error (FileStudio's text points to Save As); whether VS Code or FileStudio refuses first is not visible here.
      await h.save().then(undefined, () => undefined);
      await h.sleep(800);
      assert.equal(sha(ctx.uri), before, 'nothing was written');
      assert.equal(h.isDirty(ctx.uri), true, 'the tab is still dirty');
      assert.equal(api.state(ctx.key).stateId > api.state(ctx.key).savedStateId, true, 'the document still has its edit');
    } finally {
      await files.update('readonlyInclude', undefined, vscode.ConfigurationTarget.Workspace);
    }
  });

  test('the file on disk is a half-written zip while dirty: Revert fails, the document keeps its state and its edit', async () => {
    const ctx = await openCopy();
    await edit(ctx);
    const state = api.state(ctx.key);
    const good = fs.readFileSync(ctx.uri.fsPath);
    const mark = h.logMark();
    api.answer('externalChange', 'Keep My Edits'); // the watcher sees the damaged file first
    fs.writeFileSync(ctx.uri.fsPath, good.subarray(0, 100));
    try {
      await h.waitFor(() => h.logSince(mark).some((e) => e.text.includes('externalChange')), { timeout: 20_000, message: 'the external-change prompt' });
      await h.revert().then(undefined, () => undefined); // the error is shown to the user
      await h.sleep(800);
      const after = api.state(ctx.key);
      assert.equal(after.generation, state.generation, 'no reload happened');
      assert.equal(after.stateId, state.stateId, 'the edit is still applied');
      assert.equal(h.isDirty(ctx.uri), true, 'the tab is still dirty (VS Code cleared its mark; the no-op edit of D-067 sets it again)');
      // Undo and redo of that no-op step keep the tab dirty: VS Code cannot reach its "clean" point while the document has an edit.
      await h.undo();
      await h.sleep(500);
      assert.equal(h.isDirty(ctx.uri), true, 'dirty after Undo');
      await h.redo();
      await h.sleep(500);
      assert.equal(h.isDirty(ctx.uri), true, 'dirty after Redo');
      assert.equal(api.state(ctx.key).stateId, state.stateId, 'the real edit is untouched by those steps');
      // The other program finishes writing: Revert now works and the tab becomes clean.
      fs.writeFileSync(ctx.uri.fsPath, good);
      await h.revert();
      await clean(ctx, 'the tab to be clean after Revert of a complete file');
      await h.waitFor(() => api.state(ctx.key).generation === state.generation + 1, { message: 'the document to reload' });
      assert.equal(api.state(ctx.key).stateId, api.state(ctx.key).savedStateId, 'the document is clean again');
    } finally {
      fs.writeFileSync(ctx.uri.fsPath, good); // the teardown reverts the tab, which needs a readable file
    }
  });

  test('files.autoSave afterDelay: 20 cells typed one after another are all accepted and saved, none is lost', { timeout: 120_000 }, async () => {
    const ctx = await openCopy();
    const files = vscode.workspace.getConfiguration('files');
    await files.update('autoSave', 'afterDelay', vscode.ConfigurationTarget.Workspace);
    await files.update('autoSaveDelay', 100, vscode.ConfigurationTarget.Workspace);
    try {
      const savedIds = new Set([api.state(ctx.key).savedStateId]);
      for (let i = 0; i < 20; i++) {
        await api.pushPackageOps(ctx.key, [{ kind: 'cells', sheet: 1, cells: [[400 + i, 0, { value: { t: 's', v: `auto ${i}` } }]] }], `Typed ${i}`);
        await h.sleep(40 + (i % 5) * 60); // autosave fires somewhere in the middle of the typing
        savedIds.add(api.state(ctx.key).savedStateId);
      }
      assert.equal(savedIds.size >= 2, true, 'autosave saved at least once while the cells were still being typed');
      await h.waitFor(() => { const s = api.state(ctx.key); return s.stateId === s.savedStateId && !h.isDirty(ctx.uri); }, { timeout: 20_000, message: 'autosave to catch up ' + JSON.stringify({ st: api.state(ctx.key), tabDirty: h.isDirty(ctx.uri) }) });
      const zip = await JSZip.loadAsync(fs.readFileSync(ctx.uri.fsPath));
      const strings = await zip.file('xl/sharedStrings.xml').async('string');
      for (let i = 0; i < 20; i++) assert.ok(strings.includes(`>auto ${i}<`), `auto ${i} is in the saved file`);
      assert.ok(fs.existsSync(`${ctx.uri.fsPath}.bak`), 'the .bak was made once');
    } finally {
      await files.update('autoSave', undefined, vscode.ConfigurationTarget.Workspace);
      await files.update('autoSaveDelay', undefined, vscode.ConfigurationTarget.Workspace);
    }
  });

  test('VS Code backs a dirty workbook up on its own: a verified workbook + sidecar appear, and both are removed after Save', { timeout: 90_000 }, async () => {
    const ctx = await openCopy();
    const text = `backup probe ${ctx.name}`;
    await edit(ctx, text);
    const found = await h.waitFor(() => findBackup(text), { timeout: 30_000, interval: 500, message: 'VS Code to back the dirty tab up' });
    const sidecar = JSON.parse(fs.readFileSync(found.sidecar, 'utf8'));
    assert.equal(sidecar.v, 1);
    assert.equal(sidecar.lossyAccepted, false);
    assert.equal(sidecar.divergedFromDisk, false);
    assert.equal(sidecar.sourceStamp.size, fs.statSync(ctx.uri.fsPath).size, 'the sidecar names the file the backup was made from');
    assert.notEqual(sha({ fsPath: found.file }), sha(ctx.uri), 'the backup holds the unsaved edit, the file on disk does not');
    const onDisk = await (await JSZip.loadAsync(fs.readFileSync(ctx.uri.fsPath))).file('xl/sharedStrings.xml').async('string');
    assert.equal(onDisk.includes(`>${text}<`), false, 'the file on disk does not have the edit yet');
    await h.save();
    await clean(ctx, 'clean after Save');
    await h.waitFor(() => !fs.existsSync(found.file) && !fs.existsSync(found.sidecar), { timeout: 20_000, message: 'the backup files to be removed after Save' });
  });
});
