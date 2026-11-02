// CSV / TSV / PSV / SSV editing in a real VS Code (v0.2.0 P3, contract §8 and §14 P3 "Integration"). Drives the host
// through the test API (the value activate() returns in ExtensionMode.Test): `post` delivers a webview message to the
// first panel of a document and returns what the host posted back. No grid UI is needed. Files are temporary copies in
// the workspace (zz-*), removed after every test.
'use strict';
const assert = require('node:assert/strict');
const fs = require('fs');
const vscode = require('vscode');
const h = require('./helpers');

const CSV = 'fileStudio.csv';
const CHECKING = 'Checking the file…';
const KINDS = ['setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll'];

/** Settings this suite changes (Global), put back after every test. */
const TOUCHED = [
  'fileStudio.editing.enabled',
  'fileStudio.csv.editing',
  'files.readonlyInclude',
  'files.trimFinalNewlines',
  'files.trimTrailingWhitespace',
  '[plaintext]',
];
const settings = () => vscode.workspace.getConfiguration();
async function resetSettings() {
  for (const key of TOUCHED) await settings().update(key, undefined, vscode.ConfigurationTarget.Global);
}

async function testApi() {
  const ext = h.extension();
  const api = ext.isActive ? ext.exports : await ext.activate();
  assert.ok(api && api.version === 1, 'activate() returned the test API (ExtensionMode.Test)');
  return api;
}

let editCounter = 0;
const created = [];

suite('CSV editing (P3)', () => {
  let api;
  suiteSetup(async () => {
    api = await testApi();
    await h.closeAll();
    await resetSettings();
  });
  teardown(async () => {
    delete process.env.FILESTUDIO_TEST_FIDELITY_DELAY_MS;
    await h.closeAll();
    await resetSettings();
    for (const name of created.splice(0)) h.deleteFile(name);
  });

  /** Writes `content` to a temporary workspace file. */
  function make(name, content) {
    created.push(name);
    return h.writeBytes(name, content);
  }

  const stateOf = (uri) => api.state(uri.toString());
  const editStateOf = (uri) => stateOf(uri).editState;

  /** Waits until the host knows the edit state of `uri` and `check` accepts it. */
  function waitForState(uri, check, message) {
    return h.waitFor(
      () => {
        const state = editStateOf(uri);
        return state && check(state) ? state : undefined;
      },
      { timeout: 15_000, message },
    );
  }

  /** Opens `name` (a workspace file) in FileStudio, waits for its webview ('ready') and for the fidelity check to finish. */
  async function open(name, options = {}) {
    const mark = h.logMark();
    const { uri } = await h.openWith(name, CSV, options.viewColumn);
    await waitForState(uri, (s) => s.reason !== CHECKING, `the edit state of ${name} (check finished)`);
    await webviewReady(name, mark);
    return uri;
  }

  /**
   * Waits until the host logged that the webview of `name` sent 'ready', then until the log is quiet: VS Code may reload a
   * webview when the editor layout changes, and the host drops the answer of an edit that was in flight by design.
   */
  async function webviewReady(name, mark) {
    await h.waitFor(() => h.logSince(mark).some((e) => e.text.includes(`${name}: viewer ready`)), {
      timeout: 15_000,
      message: `the webview of ${name} to report ready`,
    });
    await h.waitForLogQuiet({ min: 300, quiet: 300, max: 3000 });
  }

  /**
   * Sends an `edit` through the test API; resolves with its `editResult`. When the editor layout changes, VS Code may
   * reload the webview while the edit is in flight: the host then drops the answer by design (the 'init' of the new
   * webview shows in the captured messages). If the document did not change either, the edit is sent again.
   */
  async function edit(uri, intent, extra = {}) {
    for (let attempt = 0; ; attempt++) {
      const editId = ++editCounter;
      const baseVersion = extra.baseVersion ?? stateOf(uri).version;
      const versionBefore = stateOf(uri).version;
      const messages = await api.post(uri.toString(), {
        type: 'edit',
        editId,
        baseVersion,
        intent,
        ...(extra.then ? { then: extra.then } : {}),
      });
      const result = messages.find((m) => m.type === 'editResult' && m.editId === editId);
      if (result) return result;
      const reloaded = messages.some((m) => m.type === 'init');
      if (reloaded && stateOf(uri).version === versionBefore + 1) return { type: 'editResult', editId, accepted: true, version: stateOf(uri).version, assumed: true };
      assert.ok(reloaded && stateOf(uri).version === versionBefore && attempt < 3, `an editResult for edit ${editId}: ${JSON.stringify(messages).slice(0, 400)}`);
      await h.sleep(500);
    }
  }
  const setCells = (...cells) => ({ kind: 'setCells', sheetId: 'csv', cells });
  const documentOf = (uri) => vscode.workspace.textDocuments.find((d) => d.uri.toString() === uri.toString());
  const textOf = (uri) => documentOf(uri).getText();

  test('a grid edit writes exactly the changed bytes; undo and redo bring them back; save writes the file', async () => {
    const mark = h.logMark();
    const uri = make('zz-edit.csv', 'name,qty\nx,1\ny,2\n');
    await open('zz-edit.csv');
    const before = stateOf(uri);
    assert.equal(before.editState.editable, true);
    assert.deepEqual(before.editState.kinds, KINDS);

    const result = await edit(uri, setCells([1, 1, '7'], [2, 0, 'p,q']));
    assert.equal(result.accepted, true, result.reason);
    assert.equal(result.version, before.version + 1);
    assert.equal(textOf(uri), 'name,qty\nx,7\n"p,q",2\n');
    assert.equal(h.isDirty(uri), true);
    assert.equal(stateOf(uri).version, before.version + 1, 'our own change is one version step (the echo is not counted again)');

    await h.undo();
    await h.waitFor(() => textOf(uri) === 'name,qty\nx,1\ny,2\n', { message: 'undo to restore the bytes' });
    assert.equal(stateOf(uri).version, before.version + 2, 'undo is a change that is not ours: +1');
    await h.redo();
    await h.waitFor(() => textOf(uri) === 'name,qty\nx,7\n"p,q",2\n', { message: 'redo to write the edit again' });

    await h.save();
    await h.waitFor(() => !h.isDirty(uri), { message: 'the file to be saved' });
    assert.equal(h.readBytes('zz-edit.csv').toString(), 'name,qty\nx,7\n"p,q",2\n');
    await h.waitForLogQuiet({ min: 300 });
    h.assertNoLogProblems(h.logSince(mark), 'edit, undo, redo, save');
  });

  test('the change is visible in a text editor of the same file next to the grid', async () => {
    const uri = make('zz-side.csv', 'a,b\n1,2\n');
    await open('zz-side.csv');
    const editor = await vscode.window.showTextDocument(uri, { viewColumn: vscode.ViewColumn.Beside, preview: false });
    const result = await edit(uri, setCells([1, 0, 'changed']));
    assert.equal(result.accepted, true, result.reason);
    assert.equal(editor.document.getText(), 'a,b\nchanged,2\n');
    assert.equal(editor.document.isDirty, true);
  });

  test('an edit sent with an old version after a text editor change is stale; with the new version it works', async () => {
    const uri = make('zz-stale.csv', 'a,b\n1,2\n');
    await open('zz-stale.csv');
    const oldVersion = stateOf(uri).version;
    const typing = new vscode.WorkspaceEdit();
    typing.insert(uri, new vscode.Position(0, 0), 'x');
    assert.ok(await vscode.workspace.applyEdit(typing), 'the text editor change');
    await h.waitFor(() => stateOf(uri).version === oldVersion + 1, { message: 'the version to follow the text editor change' });
    const refused = await edit(uri, setCells([1, 0, 'y']), { baseVersion: oldVersion });
    assert.equal(refused.accepted, false);
    assert.equal(refused.code, 'stale');
    assert.equal(textOf(uri), 'xa,b\n1,2\n', 'the stale edit changed nothing');
    const ok = await edit(uri, setCells([1, 0, 'y']));
    assert.equal(ok.accepted, true, ok.reason);
    assert.equal(textOf(uri), 'xa,b\ny,2\n');
  });

  test('files.readonlyInclude makes the document read-only (readOnly); removing it makes it editable again', async () => {
    const uri = make('zz-readonly.csv', 'a,b\n1,2\n');
    await open('zz-readonly.csv');
    await settings().update('files.readonlyInclude', { '**/zz-readonly.csv': true }, vscode.ConfigurationTarget.Global);
    const state = await waitForState(uri, (s) => !s.editable, 'editState not editable (readonlyInclude)');
    assert.equal(state.code, 'readOnly');
    const refused = await edit(uri, setCells([1, 0, 'x']));
    assert.equal(refused.accepted, false);
    assert.equal(refused.code, 'readOnly');
    assert.equal(textOf(uri), 'a,b\n1,2\n');
    await settings().update('files.readonlyInclude', undefined, vscode.ConfigurationTarget.Global);
    await waitForState(uri, (s) => s.editable, 'editable again');
  });

  test('turning the kill switch off and on changes the editState; an edit while it is off is refused (disabled)', async () => {
    const uri = make('zz-kill.csv', 'a,b\n1,2\n');
    await open('zz-kill.csv');
    await settings().update('fileStudio.csv.editing', false, vscode.ConfigurationTarget.Global);
    const off = await waitForState(uri, (s) => !s.editable, 'editState not editable (fileStudio.csv.editing)');
    assert.equal(off.code, 'disabled');
    const refused = await edit(uri, setCells([1, 0, 'x']));
    assert.equal(refused.code, 'disabled');
    assert.equal(textOf(uri), 'a,b\n1,2\n');
    await settings().update('fileStudio.csv.editing', true, vscode.ConfigurationTarget.Global);
    await waitForState(uri, (s) => s.editable, 'editable again');
  });

  test('turning editing off while the document is dirty keeps Save and undo (D-075)', async () => {
    const uri = make('zz-dirty-off.csv', 'a,b\n1,2\n');
    await open('zz-dirty-off.csv');
    assert.equal((await edit(uri, setCells([1, 0, 'x']))).accepted, true);
    await settings().update('fileStudio.editing.enabled', false, vscode.ConfigurationTarget.Global);
    await waitForState(uri, (s) => !s.editable, 'editState not editable');
    await h.save();
    await h.waitFor(() => !h.isDirty(uri), { message: 'the edits to be saved' });
    assert.equal(h.readBytes('zz-dirty-off.csv').toString(), 'a,b\nx,2\n');
    await h.undo();
    await h.waitFor(() => textOf(uri) === 'a,b\n1,2\n', { message: 'undo to work while editing is off' });
  });

  test('Change Delimiter re-pins the delimiter: the version moves and an edit sent before is stale', async () => {
    const uri = make('zz-delim.csv', 'a,b|c\n1,2|3\n');
    await open('zz-delim.csv');
    const before = stateOf(uri).version;
    // The quick pick is driven with its own commands; a key sent before it is open would go to the editor, so retry.
    let moved = false;
    for (let attempt = 0; attempt < 4 && !moved; attempt++) {
      const picked = vscode.commands.executeCommand('fileStudio.changeDelimiter');
      await h.sleep(800 + attempt * 400);
      // items: Detect automatically, Comma, Semicolon, Tab, Pipe, Spaces
      for (let i = 0; i < 4; i++) await vscode.commands.executeCommand('workbench.action.quickOpenSelectNext');
      await vscode.commands.executeCommand('workbench.action.acceptSelectedQuickOpenItem');
      await picked;
      moved = await h.waitFor(() => stateOf(uri).version === before + 1, { timeout: 2500, message: 'the version to move' }).then(() => true, () => false);
      if (!moved) await vscode.commands.executeCommand('workbench.action.closeQuickOpen');
    }
    assert.ok(moved, 'the version moved after Change Delimiter');
    const stale = await edit(uri, setCells([1, 0, 'x']), { baseVersion: before });
    assert.equal(stale.code, 'stale');
    const ok = await edit(uri, setCells([1, 1, 'x']));
    assert.equal(ok.accepted, true, ok.reason);
    assert.equal(textOf(uri), 'a,b|c\n1,2|x\n', 'written with the delimiter on screen (|)');
  });

  test('the Excel sample.csv (a line break inside a cell, CRLF rows) is not editable (fidelity) until "Edit anyway"', async () => {
    const uri = h.copyFile('sample.csv', 'zz-sample.csv');
    created.push('zz-sample.csv');
    await open('zz-sample.csv');
    const state = editStateOf(uri);
    assert.equal(state.editable, false);
    assert.equal(state.code, 'fidelity');
    assert.ok(state.canEditAnyway && state.canEditAnyway.details.length > 0, 'the banner lists what VS Code would change');
    const refused = await edit(uri, setCells([1, 1, 'x']));
    assert.equal(refused.code, 'fidelity');
    assert.equal(h.isDirty(uri), false);
    api.answer('fidelity', 'Edit anyway');
    await api.post(uri.toString(), { type: 'editAnyway' });
    await waitForState(uri, (s) => s.editable, 'editable after Edit anyway');
    const ok = await edit(uri, setCells([1, 1, 'x']));
    assert.equal(ok.accepted, true, ok.reason);
  });

  test('a Latin-1 file opened as UTF-8 is blocked (wrong encoding), without Edit anyway', async () => {
    const uri = make('zz-latin1.csv', Buffer.from([0x6e, 0x61, 0x6d, 0x65, 0x2c, 0x63, 0x69, 0x74, 0x79, 0x0a, 0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x2c, 0x4b, 0xf6, 0x6c, 0x6e, 0x0a]));
    await open('zz-latin1.csv');
    const state = editStateOf(uri);
    assert.equal(state.editable, false);
    assert.equal(state.code, 'blocked');
    assert.equal(state.canEditAnyway, undefined);
    const refused = await edit(uri, setCells([1, 0, 'x']));
    assert.equal(refused.code, 'blocked');
    assert.equal(h.isDirty(uri), false);
  });

  test('a quote that is never closed blocks editing (Row N opens a quote)', async () => {
    const uri = make('zz-unclosed.csv', 'a,b\n1,"never closed\n3,4\n');
    await open('zz-unclosed.csv');
    const state = editStateOf(uri);
    assert.equal(state.code, 'blocked');
    assert.equal(state.reason, 'Row 2 opens a quote that is never closed; editing could change the rest of the file.');
  });

  test('two panels of one file share one version; an edit in one is one version step', async () => {
    const uri = make('zz-two.csv', 'a,b\n1,2\n');
    await open('zz-two.csv');
    await open('zz-two.csv', { viewColumn: vscode.ViewColumn.Beside });
    const before = stateOf(uri).version;
    const result = await edit(uri, setCells([1, 0, '9']));
    assert.equal(result.accepted, true, result.reason);
    assert.equal(stateOf(uri).version, before + 1);
    assert.equal(textOf(uri), 'a,b\n9,2\n');
  });

  test('a CSV over 50 MB opens in VS Code\'s text editor (0.1.2 behaviour), not in FileStudio', async () => {
    const name = 'zz-big.csv';
    created.push(name);
    const fd = fs.openSync(h.fileUri(name).fsPath, 'w');
    try {
      fs.writeSync(fd, 'a,b\n');
      fs.ftruncateSync(fd, 51 * 1024 * 1024);
    } finally {
      fs.closeSync(fd);
    }
    await vscode.commands.executeCommand('vscode.openWith', h.fileUri(name), CSV);
    const tab = await h.waitFor(() => {
      const t = h.findTab(name);
      return t && t.input instanceof vscode.TabInputText ? t : undefined;
    }, { timeout: 20_000, message: 'the text editor tab of the big file' });
    assert.ok(tab.input instanceof vscode.TabInputText);
  });

  // ----- 0.2.1 (P3 WP4): what the grid UI sends, through the same host -----

  test('a paste of 1,000 x 10 cells is one edit and one undo step; typed values stay text', async () => {
    const uri = make('zz-paste.csv', 'a,b\n1,2\n');
    await open('zz-paste.csv');
    const cells = [];
    for (let r = 1; r <= 1000; r++) for (let c = 0; c < 10; c++) cells.push([r, c, r === 1 && c === 0 ? '=1+1' : r === 1 && c === 1 ? '007' : `v${r}_${c}`]);
    const before = stateOf(uri).version;
    const result = await edit(uri, { kind: 'setCells', sheetId: 'csv', cells });
    assert.equal(result.accepted, true, result.reason);
    assert.equal(stateOf(uri).version, before + 1, 'one version step');
    const lines = textOf(uri).split('\n');
    assert.equal(lines.length, 1002, 'a,b + 1000 rows + the final line break');
    assert.equal(lines[1], '=1+1,007,v1_2,v1_3,v1_4,v1_5,v1_6,v1_7,v1_8,v1_9');
    await h.undo();
    await h.waitFor(() => textOf(uri) === 'a,b\n1,2\n', { message: 'ONE undo to bring the whole paste back' });
  });

  test('sort and Replace All write the expected bytes, one undo step each; row 1 (the header) never moves by sort', async () => {
    const uri = make('zz-sortrep.csv', 'name,qty\nb,2\na,10\nc,1\n');
    await open('zz-sortrep.csv');
    const range = { r0: 0, c0: 0, r1: 3, c1: 1 };
    let result = await edit(uri, { kind: 'sort', sheetId: 'csv', range, key: 1, desc: false, header: true });
    assert.equal(result.accepted, true, result.reason);
    assert.equal(textOf(uri), 'name,qty\nc,1\nb,2\na,10\n', 'numbers sort as numbers, the header stays');
    await h.undo();
    await h.waitFor(() => textOf(uri) === 'name,qty\nb,2\na,10\nc,1\n', { message: 'undo of the sort' });
    const query = { scope: 'csv', text: 'A', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' };
    result = await edit(uri, { kind: 'replaceAll', query, replace: 'x' });
    assert.equal(result.accepted, true, result.reason);
    assert.equal(textOf(uri), 'nxme,qty\nb,2\nx,10\nc,1\n', 'Replace All includes row 1 like Find does');
    await h.undo();
    await h.waitFor(() => textOf(uri) === 'name,qty\nb,2\na,10\nc,1\n', { message: 'undo of Replace All' });
  });

  test('a 48 MB CSV: one cell edit is answered fast (the host and VS Code together)', async () => {
    const name = 'zz-48mb.csv';
    created.push(name);
    const row = 'id,"Name 1",City1,12345.67,2026-10-01,"note, with comma",C1,7,3.14,TRUE,user@example.com,"he said ""hi"""\r\n';
    const head = 'id,name,city,amount,date,note,code,qty,price,flag,email,comment\r\n';
    const n = Math.floor((48 * 1000 * 1000) / row.length);
    fs.writeFileSync(h.fileUri(name).fsPath, Buffer.from(head + row.repeat(n)));
    const uri = await open(name);
    assert.equal(editStateOf(uri).editable, true, 'a file under 50 MB is editable');
    // Warm up once (JIT, the lazy quote policy), then measure the round trip of one cell.
    let result = await edit(uri, setCells([n >> 1, 5, 'warm up']));
    assert.equal(result.accepted, true, result.reason);
    const t0 = process.hrtime.bigint();
    result = await edit(uri, setCells([n >> 1, 5, 'measured, value']));
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    assert.equal(result.accepted, true, result.reason);
    console.log(`[perf] 48 MB CSV, one-cell edit, host round trip: ${ms.toFixed(1)} ms`);
    assert.ok(ms < 100, `one cell edit on a 48 MB CSV took ${ms.toFixed(1)} ms (limit 100 ms)`);
  });

  // ----- critic additions (D-072) -----

  test('an edit sent before the fidelity check finished is refused (fidelity); after it, the same edit works', async () => {
    process.env.FILESTUDIO_TEST_FIDELITY_DELAY_MS = '2500';
    const uri = make('zz-pending.csv', 'a,b\n1,2\n');
    await h.openWith('zz-pending.csv', CSV);
    const pending = await waitForState(uri, () => true, 'the first edit state');
    assert.equal(pending.editable, false);
    assert.equal(pending.reason, CHECKING);
    const refused = await edit(uri, setCells([1, 0, 'x']), { baseVersion: stateOf(uri).version });
    assert.equal(refused.accepted, false);
    assert.equal(refused.code, 'fidelity');
    assert.equal(textOf(uri), 'a,b\n1,2\n');
    await waitForState(uri, (s) => s.editable, 'the check to finish');
    const ok = await edit(uri, setCells([1, 0, 'x']));
    assert.equal(ok.accepted, true, ok.reason);
  });

  test('a clean CSV replaced on disk by a mixed-EOL version is reloaded by VS Code and checked again (not editable)', async () => {
    const uri = make('zz-reload.csv', 'a,b\n1,2\n');
    await open('zz-reload.csv');
    assert.equal(editStateOf(uri).editable, true);
    fs.writeFileSync(h.fileUri('zz-reload.csv').fsPath, 'a,b\r\n1,2\n3,4\r\n');
    const state = await waitForState(uri, (s) => !s.editable && s.code === 'fidelity' && s.reason !== CHECKING, 'the check after the reload');
    assert.ok(state.canEditAnyway, 'the mixed line breaks are listed');
    assert.match(state.canEditAnyway.details.join(' '), /line break/);
  });

  test('files.trimFinalNewlines counts trailing blank records in the fidelity message', async () => {
    await settings().update('files.trimFinalNewlines', true, vscode.ConfigurationTarget.Global);
    const uri = make('zz-trim.csv', 'a,b\n1,2\n\n\n');
    await open('zz-trim.csv');
    const state = editStateOf(uri);
    assert.equal(state.editable, false);
    assert.equal(state.code, 'fidelity');
    assert.match(state.canEditAnyway.details.join(' '), /empty line.*end of the file.*files\.trimFinalNewlines/);
  });

  test('a language override of files.trimTrailingWhitespace is honoured (".csv" is plaintext in the test VS Code)', async () => {
    const uri = make('zz-trailing.csv', 'a,b  \n1,2\n');
    await open('zz-trailing.csv');
    assert.equal(editStateOf(uri).editable, true, 'no setting: nothing would change');
    await settings().update('[plaintext]', { 'files.trimTrailingWhitespace': true }, vscode.ConfigurationTarget.Global);
    // The setting is read per document language when the check runs: open the file again.
    await h.closeAll();
    await open('zz-trailing.csv');
    const state = editStateOf(uri);
    assert.equal(state.code, 'fidelity');
    assert.match(state.canEditAnyway.details.join(' '), /remove spaces at the end of 1 line.*files\.trimTrailingWhitespace/);
  });

  test('an untitled CSV is editable (there are no bytes on disk to change)', async () => {
    const uri = vscode.Uri.parse('untitled:zz-new.csv');
    await vscode.commands.executeCommand('vscode.openWith', uri, CSV);
    await waitForState(uri, (s) => s.reason !== CHECKING, 'the edit state of the untitled CSV');
    assert.equal(editStateOf(uri).editable, true);
    const result = await edit(uri, setCells([0, 0, 'x']));
    assert.equal(result.accepted, true, result.reason);
    assert.equal(textOf(uri), 'x');
  });
});
