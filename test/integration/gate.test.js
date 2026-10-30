// The edit gate in a real VS Code (v0.2.0 foundation): the settings kill switch and files.readonlyInclude change the
// editState of an open Markdown file, a task toggle is refused when editing is off or the file is read-only and applied
// when it is on, and `languages.match` works with a minimal { uri, languageId } object for an xlsx Uri (contract §3.2).
// Drives the host through the test API (the value activate() returns in ExtensionMode.Test, contract §13.3).
'use strict';
const assert = require('node:assert/strict');
const vscode = require('vscode');
const h = require('./helpers');

const TASKS = 'gate-tasks.md';
const TASKS_TEXT = '# Tasks\n\n- [ ] first\n- [x] second\n';
/** Line index (0-based) of "- [ ] first". */
const FIRST_LINE = 2;

const settings = () => vscode.workspace.getConfiguration();
/** Settings this suite changes, put back after every test. */
const TOUCHED = ['fileStudio.editing.enabled', 'fileStudio.markdown.editing', 'files.readonlyInclude'];

async function resetSettings() {
  for (const key of TOUCHED) await settings().update(key, undefined, vscode.ConfigurationTarget.Global);
}

/** The test API, or fails with a clear message when the extension did not return it. */
async function testApi() {
  const ext = h.extension();
  const api = ext.isActive ? ext.exports : await ext.activate();
  assert.ok(api && api.version === 1, 'activate() returned the test API (ExtensionMode.Test)');
  return api;
}

/** Waits until the host's editState of `uri` matches `check`. */
function waitForEditState(api, uri, check, message) {
  return h.waitFor(
    () => {
      const state = api.state(uri.toString()).editState;
      return state && check(state) ? state : undefined;
    },
    { timeout: 15_000, message },
  );
}

suite('Edit gate', () => {
  let api;
  suiteSetup(async () => {
    api = await testApi();
    await h.closeAll();
    await resetSettings();
  });
  teardown(async () => {
    await h.closeAll();
    await resetSettings();
    h.deleteFile(TASKS);
  });

  async function openTasks() {
    h.writeBytes(TASKS, TASKS_TEXT);
    const { uri } = await h.openWith(TASKS, 'fileStudio.markdown');
    await waitForEditState(api, uri, () => true, 'the first editState of the Markdown document');
    // The panel is rendered when it sent 'ready' and the host answered: the version is the document's.
    await h.waitFor(() => api.state(uri.toString()).version > 0, { message: 'the document version' });
    return uri;
  }

  const toggle = async (uri, checked) => {
    const messages = await api.post(uri.toString(), {
      type: 'toggleTask',
      line: FIRST_LINE,
      checked,
      version: api.state(uri.toString()).version,
    });
    return messages.find((m) => m.type === 'toggleTaskResult');
  };

  test('the test API exists in the test host and reports an editable Markdown document', async () => {
    const uri = await openTasks();
    const state = api.state(uri.toString()).editState;
    assert.equal(state.editable, true);
  });

  test('switching fileStudio.editing.enabled off posts an editState that is not editable, and on makes it editable again', async () => {
    const uri = await openTasks();
    await settings().update('fileStudio.editing.enabled', false, vscode.ConfigurationTarget.Global);
    const off = await waitForEditState(api, uri, (s) => !s.editable, 'editState not editable (kill switch)');
    assert.equal(off.code, 'disabled');
    await settings().update('fileStudio.editing.enabled', true, vscode.ConfigurationTarget.Global);
    await waitForEditState(api, uri, (s) => s.editable, 'editState editable again');
  });

  test('switching fileStudio.markdown.editing off changes the editState of the Markdown document', async () => {
    const uri = await openTasks();
    await settings().update('fileStudio.markdown.editing', false, vscode.ConfigurationTarget.Global);
    const off = await waitForEditState(api, uri, (s) => !s.editable, 'editState not editable (markdown.editing)');
    assert.equal(off.code, 'disabled');
  });

  test('a task toggle is refused when editing is off, and the file is not changed', async () => {
    const uri = await openTasks();
    await settings().update('fileStudio.markdown.editing', false, vscode.ConfigurationTarget.Global);
    await waitForEditState(api, uri, (s) => !s.editable, 'editState not editable');
    const result = await toggle(uri, true);
    assert.ok(result, 'a toggleTaskResult was posted');
    assert.equal(result.applied, false);
    assert.equal(h.isDirty(uri), false, 'the refused toggle did not change the document');
    assert.equal(h.readBytes(TASKS).toString(), TASKS_TEXT);
  });

  test('a task toggle is applied when editing is on (0.1.2 behaviour)', async () => {
    const uri = await openTasks();
    // The first rendering may still be on its way (the host knows the task markers only after it rendered the version).
    const result = await h.waitFor(
      async () => {
        const r = await toggle(uri, true);
        return r && (r.applied || r.reason !== 'no task checkbox on that line') ? r : undefined;
      },
      { timeout: 15_000, message: 'the task toggle to be applied' },
    );
    assert.equal(result.applied, true, result.reason);
    const doc = vscode.workspace.textDocuments.find((d) => d.uri.toString() === uri.toString());
    assert.ok(doc, 'the text document is open');
    assert.equal(doc.lineAt(FIRST_LINE).text, '- [x] first');
    assert.equal(doc.lineAt(FIRST_LINE + 1).text, '- [x] second', 'only the requested task changed');
    // The same toggle again: already in that state, still applied, no second change.
    const again = await toggle(uri, true);
    assert.equal(again.applied, true);
    assert.equal(doc.lineAt(FIRST_LINE).text, '- [x] first');
  });

  test('files.readonlyInclude makes the document read-only: editState and task toggle refuse', async () => {
    const uri = await openTasks();
    await settings().update('files.readonlyInclude', { '**/gate-tasks.md': true }, vscode.ConfigurationTarget.Global);
    const state = await waitForEditState(api, uri, (s) => !s.editable, 'editState not editable (readonlyInclude)');
    assert.equal(state.code, 'readOnly');
    const result = await toggle(uri, true);
    assert.ok(result, 'a toggleTaskResult was posted');
    assert.equal(result.applied, false);
    assert.equal(h.readBytes(TASKS).toString(), TASKS_TEXT);
    await settings().update('files.readonlyInclude', undefined, vscode.ConfigurationTarget.Global);
    await waitForEditState(api, uri, (s) => s.editable, 'editable again after readonlyInclude is removed');
  });

  test('languages.match accepts a minimal { uri, languageId } object for an xlsx Uri (contract §3.2 step 2)', () => {
    const uri = h.fileUri('styled.xlsx');
    const doc = { uri, languageId: 'plaintext' };
    const hit = vscode.languages.match({ pattern: '**/*.xlsx' }, doc);
    const miss = vscode.languages.match({ pattern: '**/*.csv' }, doc);
    const bySuffix = vscode.languages.match({ pattern: '**/styled.xlsx' }, doc);
    console.log(`[gate] languages.match({pattern:'**/*.xlsx'}, {uri, languageId}) = ${hit}; csv pattern = ${miss}; name pattern = ${bySuffix}`);
    assert.ok(hit > 0, 'a matching glob scores above 0');
    assert.ok(bySuffix > 0, 'a file-name glob scores above 0');
    assert.equal(miss, 0, 'a non-matching glob scores 0');
  });
});
