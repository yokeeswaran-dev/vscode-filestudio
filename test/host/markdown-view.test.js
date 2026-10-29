// Host tests for the Markdown view (v0.2.0, P6): modes, Split scroll sync, WYSIWYG edits (applyMdEdit), the keys
// (command undo / redo / save), the flush before a save, the save-fidelity gate (D-072) and the size limits.
// The host code runs against test/host/fake-vscode.js. Run with `npm test`.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');

const { vscode, T, host, extensionUri, token, tick, writeTmp, lastOf, callsOf, waitFor } = require('./load-host');

beforeEach(() => T.reset());

/**
 * Opens `text` as a Markdown panel (the file on disk holds `disk`, default = the same text) and waits for the init and
 * for the save check to finish (editState.kinds has 'mdText' or the check said something else).
 */
async function openMd(name, text, { disk = text, settle = true } = {}) {
  const uri = writeTmp(name, disk);
  const document = new T.FakeTextDocument(uri, text);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  if (settle) await settled(panel);
  return { uri, document, panel, init };
}

/** Waits until the latest editState a panel got is not "Checking the file…". */
async function settled(panel) {
  return waitFor(
    () => {
      const state = [...panel.posted].reverse().find((m) => m.editState)?.editState;
      return state && state.reason !== 'Checking the file…' ? state : undefined;
    },
    3000,
    'the save check',
  );
}

let nextEditId = 1;
/** Sends an `edit` (mdText) that replaces `oldText` (first occurrence) with `text`; resolves with the editResult. */
async function sendEdit(opened, oldText, text, extra = {}) {
  const { document, panel } = opened;
  const start = document.getText().indexOf(oldText);
  assert.ok(start >= 0, `"${oldText}" is in the document`);
  const editId = nextEditId++;
  panel.posted.length = 0;
  panel.send({
    type: 'edit',
    editId,
    baseVersion: extra.baseVersion ?? document.version,
    intent: { kind: 'mdText', edits: [{ start, end: start + oldText.length, oldText, text }] },
    ...(extra.then ? { then: extra.then } : {}),
  });
  return waitFor(() => panel.posted.find((m) => m.type === 'editResult' && m.editId === editId), 3000, 'editResult');
}

/** Switches the panel to a mode and waits for the update the host answers with. */
async function setMode(opened, mode) {
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'setMode', mode });
  return waitFor(() => lastOf(opened.panel, 'markdownUpdate'), 3000, 'markdownUpdate after setMode');
}

const DOC = '# Title\n\nhello world\n\n- item\n';

// ===== init, modes, limits =====

test('init carries eol, limits, resourceBase; the save check ends with kinds [mdText]; Preview has html and no blocks', async () => {
  const { init, panel } = await openMd('init.md', DOC);
  assert.equal(init.kind, 'markdown');
  assert.equal(init.eol, '\n');
  assert.deepEqual(init.limits, { wysiwyg: true });
  assert.match(init.resourceBase, /\/$/);
  assert.ok(init.html.includes('<h1'), 'html for Preview');
  assert.equal(init.blocks, undefined, 'blocks only go to a WYSIWYG panel');
  assert.equal(init.editState.editable, true);
  // The first state says "checking" (WYSIWYG not usable yet, the checkboxes are); the update after the check allows it.
  assert.equal(init.editState.code, 'fidelity');
  assert.equal(init.editState.reason, 'Checking the file…');
  assert.deepEqual(init.editState.kinds, []);
  const done = await settled(panel);
  assert.deepEqual(done, { editable: true, kinds: ['mdText'] });
});

test('CRLF document: eol "\\r\\n"; WYSIWYG mode sends blocks and no html', async () => {
  const text = '# T\r\n\r\npara\r\n';
  const opened = await openMd('crlf.md', text);
  assert.equal(opened.init.eol, '\r\n');
  const update = await setMode(opened, 'wysiwyg');
  assert.equal(update.html, '');
  assert.deepEqual(update.blocks.map((b) => [b.kind, b.start, b.end]).filter((b) => b[0] !== 'gap').length, 2);
  assert.equal(update.source, text);
  assert.deepEqual(update.toc.map((t) => t.text), ['T']);
  assert.equal(update.eol, '\r\n');
});

test('back to Preview sends the html again', async () => {
  const opened = await openMd('back.md', DOC);
  await setMode(opened, 'wysiwyg');
  const update = await setMode(opened, 'preview');
  assert.ok(update.html.includes('<h1'));
  assert.equal(update.blocks, undefined);
});

test('a file over 200,000 characters: limits.wysiwyg false; WYSIWYG falls back to Preview with the html', async () => {
  const big = '# Big\n\n' + 'word '.repeat(41_000) + '\n';
  assert.ok(big.length > 200_000);
  const opened = await openMd('big.md', big);
  assert.deepEqual(opened.init.limits, {
    wysiwyg: false,
    reason: 'This file is too large for WYSIWYG editing. Use Split to edit the Markdown source.',
  });
  const update = await setMode(opened, 'wysiwyg');
  assert.ok(update.html.includes('<h1'), 'Preview html');
  assert.equal(update.blocks, undefined);
  assert.equal(update.limits.wysiwyg, false);
});

test('more than 3,000 blocks: limits.wysiwyg false', async () => {
  const text = Array.from({ length: 3001 }, (_, i) => `p${i}`).join('\n\n') + '\n';
  assert.ok(text.length < 200_000);
  const opened = await openMd('many.md', text);
  assert.equal(opened.init.limits.wysiwyg, false);
  assert.match(opened.init.limits.reason, /too large for WYSIWYG/);
  const edit = await sendEdit(opened, 'p1', 'q1');
  assert.deepEqual([edit.accepted, edit.code], [false, 'blocked']);
  assert.equal(opened.document.getText(), text);
});

test('defaultMode wysiwyg: the init already carries the blocks and no html', async () => {
  T.config.set('fileStudio.markdown.defaultMode', 'wysiwyg');
  const uri = writeTmp('dm.md', DOC);
  const document = new T.FakeTextDocument(uri, DOC);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  // The check is still pending: WYSIWYG opens read-only ("Checking the file…").
  assert.equal(init.mode, 'wysiwyg');
  assert.equal(init.html, '');
  assert.ok(init.blocks.length > 0);
  assert.equal(init.editState.reason, 'Checking the file…');
});

// ===== kill switch, read-only =====

test('kill switch off: WYSIWYG falls back to Preview, edits are refused (disabled), Split still opens the editor', async () => {
  const opened = await openMd('kill.md', DOC);
  T.config.set('fileStudio.markdown.editing', false);
  const update0 = await setMode(opened, 'wysiwyg');
  assert.ok(update0.html.includes('<h1'), 'Preview: the kill switch is read at once');
  assert.equal(update0.editState.editable, false);
  const refused = await sendEdit(opened, 'hello', 'bye');
  assert.deepEqual([refused.accepted, refused.code], [false, 'disabled']);
  assert.match(refused.reason, /fileStudio\.markdown\.editing/);
  assert.equal(opened.document.getText(), DOC);
  assert.deepEqual(callsOf('applyEdit'), []);
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await waitFor(() => callsOf('showTextDocument').length === 1, 3000, 'showTextDocument');
});

test('read-only document (files.readonlyInclude): edits are refused (readOnly), WYSIWYG is not offered', async () => {
  T.config.set('files.readonlyInclude', { '**/locked/**': true });
  const opened = await openMd('locked/ro.md', DOC);
  assert.equal(opened.init.readOnly, true);
  const update = await setMode(opened, 'wysiwyg');
  assert.ok(update.html.includes('<h1'));
  const refused = await sendEdit(opened, 'hello', 'bye');
  assert.deepEqual([refused.accepted, refused.code], [false, 'readOnly']);
  assert.match(refused.reason, /files\.readonlyInclude/);
  assert.deepEqual(callsOf('applyEdit'), []);
});

// ===== applyMdEdit =====

test('an edit changes only the typed word: one applyEdit, the answer carries the new version, the update follows', async () => {
  const opened = await openMd('edit.md', DOC);
  await setMode(opened, 'wysiwyg');
  T.calls.length = 0;
  const result = await sendEdit(opened, 'world', 'there');
  assert.equal(result.accepted, true);
  assert.equal(opened.document.getText(), '# Title\n\nhello there\n\n- item\n');
  assert.equal(result.version, opened.document.version);
  assert.equal(callsOf('applyEdit').length, 1);
  assert.deepEqual(callsOf('applyEdit')[0][1], [[2, 6, 2, 11, 'there']]);
  // The answer is posted before the update with the new source and blocks.
  const types = opened.panel.posted.map((m) => m.type);
  assert.ok(types.indexOf('editResult') < types.indexOf('markdownUpdate'), types.join());
  const update = lastOf(opened.panel, 'markdownUpdate');
  assert.equal(update.version, opened.document.version);
  assert.equal(update.source, opened.document.getText());
  assert.ok(update.blocks.length > 0);
  assert.equal(update.html, '');
});

test('several edits in one message are one WorkspaceEdit (one undo step)', async () => {
  const opened = await openMd('multi.md', DOC);
  await setMode(opened, 'wysiwyg');
  T.calls.length = 0;
  const text = opened.document.getText();
  const a = text.indexOf('Title');
  const b = text.indexOf('item');
  opened.panel.posted.length = 0;
  opened.panel.send({
    type: 'edit',
    editId: 900,
    baseVersion: opened.document.version,
    intent: {
      kind: 'mdText',
      edits: [
        { start: a, end: a + 5, oldText: 'Title', text: 'Head' },
        { start: b, end: b + 4, oldText: 'item', text: 'thing' },
      ],
    },
  });
  const result = await waitFor(() => lastOf(opened.panel, 'editResult'), 3000, 'editResult');
  assert.equal(result.accepted, true);
  assert.equal(callsOf('applyEdit').length, 1, 'one WorkspaceEdit');
  assert.equal(opened.document.getText(), '# Head\n\nhello world\n\n- thing\n');
});

test('stale baseVersion: refused (stale) and the view gets a fresh update; nothing changed', async () => {
  const opened = await openMd('stale.md', DOC);
  await setMode(opened, 'wysiwyg');
  const refused = await sendEdit(opened, 'world', 'there', { baseVersion: opened.document.version - 1 });
  assert.deepEqual([refused.accepted, refused.code], [false, 'stale']);
  assert.match(refused.reason, /The file changed while you were editing/);
  assert.equal(opened.document.getText(), DOC);
  assert.ok(lastOf(opened.panel, 'markdownUpdate'), 'reload from the source');
  assert.deepEqual(callsOf('applyEdit'), []);
});

test('oldText that is not in the document at that place: refused (invalid)', async () => {
  const opened = await openMd('invalid.md', DOC);
  await setMode(opened, 'wysiwyg');
  const start = DOC.indexOf('world');
  opened.panel.posted.length = 0;
  opened.panel.send({
    type: 'edit',
    editId: 901,
    baseVersion: opened.document.version,
    intent: { kind: 'mdText', edits: [{ start, end: start + 5, oldText: 'wxrld', text: 'there' }] },
  });
  const refused = await waitFor(() => lastOf(opened.panel, 'editResult'), 3000, 'editResult');
  assert.deepEqual([refused.accepted, refused.code], [false, 'invalid']);
  assert.equal(opened.document.getText(), DOC);
  assert.deepEqual(callsOf('applyEdit'), []);
});

test('an edit that would change another block (paragraph merge): refused (blocked) with the message, text unchanged', async () => {
  const src = 'Intro.\n\nA paragraph\n- item\n';
  const opened = await openMd('merge.md', src);
  await setMode(opened, 'wysiwyg');
  // The list item loses its marker: "item" would continue the untouched paragraph above it.
  const refused = await sendEdit(opened, '- item', 'item');
  assert.deepEqual([refused.accepted, refused.code], [false, 'blocked']);
  assert.equal(refused.reason, 'This change would also change other parts of the file; it was not applied.');
  assert.equal(opened.document.getText(), src);
  assert.deepEqual(callsOf('applyEdit'), []);
});

test('new text gets the document line break: LF text in a CRLF document is written as CRLF', async () => {
  const text = '# T\r\n\r\nline\r\n';
  const opened = await openMd('eol.md', text);
  await setMode(opened, 'wysiwyg');
  const result = await sendEdit(opened, 'line', 'line one\nline two\rline three');
  assert.equal(result.accepted, true);
  assert.equal(opened.document.getText(), '# T\r\n\r\nline one\r\nline two\r\nline three\r\n');
  assert.equal(callsOf('applyEdit')[0][1][0][4], 'line one\r\nline two\r\nline three');
});

test('applyEdit false: refused (failed) with the VS Code message; the view is refreshed', async () => {
  const opened = await openMd('failed.md', DOC);
  await setMode(opened, 'wysiwyg');
  const original = vscode.workspace.applyEdit;
  vscode.workspace.applyEdit = async () => false;
  try {
    const refused = await sendEdit(opened, 'world', 'there');
    assert.deepEqual([refused.accepted, refused.code], [false, 'failed']);
    assert.equal(refused.reason, 'VS Code did not apply the change. The file was not changed.');
    assert.equal(opened.document.getText(), DOC);
  } finally {
    vscode.workspace.applyEdit = original;
  }
});

test('a non-mdText intent is refused (invalid); a malformed edit never reaches the document', async () => {
  const opened = await openMd('kind.md', DOC);
  await setMode(opened, 'wysiwyg');
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'edit', editId: 902, baseVersion: 1, intent: { kind: 'setCells', sheetId: 's', cells: [] } });
  const refused = await waitFor(() => lastOf(opened.panel, 'editResult'), 3000, 'editResult');
  assert.deepEqual([refused.accepted, refused.code], [false, 'invalid']);
  assert.deepEqual(callsOf('applyEdit'), []);
});

test('then: "save" saves the document after the edit', async () => {
  const opened = await openMd('thensave.md', DOC);
  await setMode(opened, 'wysiwyg');
  const result = await sendEdit(opened, 'world', 'there', { then: 'save' });
  assert.equal(result.accepted, true);
  await waitFor(() => fs.readFileSync(opened.uri.fsPath, 'utf8') === '# Title\n\nhello there\n\n- item\n', 3000, 'the saved file');
  assert.equal(callsOf('save').length, 1);
});

test('two panels of one document: an edit in A reaches B as a markdownUpdate; B has the same queue and state', async () => {
  const a = await openMd('two.md', DOC);
  const panelB = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(a.document, panelB, token);
  panelB.send({ type: 'ready' });
  await waitFor(() => lastOf(panelB, 'init'), 3000, 'init B');
  await settled(panelB);
  await setMode(a, 'wysiwyg');
  panelB.posted.length = 0;
  const result = await sendEdit(a, 'world', 'there');
  assert.equal(result.accepted, true);
  const update = await waitFor(() => lastOf(panelB, 'markdownUpdate'), 3000, 'update in B');
  assert.equal(update.source, '# Title\n\nhello there\n\n- item\n');
});

// ===== command: undo / redo / save, flush =====

function activateTab(uri) {
  return T.addTab({ input: new vscode.TabInputCustom(uri, 'fileStudio.markdown'), label: 'x.md', isActive: true });
}

test('a command runs behind the queued edit: edit, then undo (never reordered)', async () => {
  const opened = await openMd('order.md', DOC);
  await setMode(opened, 'wysiwyg');
  activateTab(opened.uri);
  T.calls.length = 0;
  opened.panel.posted.length = 0;
  const start = DOC.indexOf('world');
  opened.panel.send({
    type: 'edit',
    editId: 903,
    baseVersion: opened.document.version,
    intent: { kind: 'mdText', edits: [{ start, end: start + 5, oldText: 'world', text: 'there' }] },
  });
  opened.panel.send({ type: 'command', id: 'undo' });
  await waitFor(() => callsOf('executeCommand').length === 1, 3000, 'undo');
  assert.deepEqual(T.calls.map((c) => c[0]).filter((n) => n === 'applyEdit' || n === 'executeCommand'), ['applyEdit', 'executeCommand']);
  assert.deepEqual(callsOf('executeCommand')[0].slice(1), ['undo']);
});

test('redo and save commands run on the active tab; save uses document.save()', async () => {
  const opened = await openMd('keys.md', DOC);
  await setMode(opened, 'wysiwyg');
  activateTab(opened.uri);
  T.calls.length = 0;
  opened.panel.send({ type: 'command', id: 'redo' });
  opened.panel.send({ type: 'command', id: 'save' });
  await waitFor(() => callsOf('save').length === 1 && callsOf('executeCommand').length === 1, 3000, 'redo + save');
  assert.deepEqual(callsOf('executeCommand')[0].slice(1), ['redo']);
});

test('a command while another tab is active does nothing and says why (status bar)', async () => {
  const opened = await openMd('inactive.md', DOC);
  await setMode(opened, 'wysiwyg');
  const other = writeTmp('other.md', 'x');
  activateTab(other);
  T.calls.length = 0;
  opened.panel.send({ type: 'command', id: 'undo' });
  await waitFor(() => callsOf('statusBar').length === 1, 3000, 'status message');
  assert.deepEqual(callsOf('executeCommand'), []);
  assert.deepEqual(callsOf('save'), []);
  assert.match(callsOf('statusBar')[0][1], /not the active tab/);
});

test('undo / redo / save commands still work while editing is turned off (nothing is stranded, D-075)', async () => {
  const opened = await openMd('stranded.md', DOC);
  await setMode(opened, 'wysiwyg');
  activateTab(opened.uri);
  T.config.set('fileStudio.editing.enabled', false);
  T.calls.length = 0;
  opened.panel.send({ type: 'command', id: 'undo' });
  opened.panel.send({ type: 'command', id: 'save' });
  await waitFor(() => callsOf('save').length === 1 && callsOf('executeCommand').length === 1, 3000, 'undo + save');
});

test('Save All: the host asks a WYSIWYG panel to flush and the save waits for its "flushed"', async () => {
  const opened = await openMd('flush.md', DOC);
  await setMode(opened, 'wysiwyg');
  opened.panel.posted.length = 0;
  let released = false;
  const saving = T.fireWillSave(opened.document).then(() => {
    released = true;
  });
  const request = await waitFor(() => lastOf(opened.panel, 'flushRequest'), 3000, 'flushRequest');
  await tick(30);
  assert.equal(released, false, 'the save waits');
  // The editor's pending typing arrives as an edit, then "flushed".
  const start = DOC.indexOf('world');
  opened.panel.send({
    type: 'edit',
    editId: 904,
    baseVersion: opened.document.version,
    intent: { kind: 'mdText', edits: [{ start, end: start + 5, oldText: 'world', text: 'there' }] },
  });
  opened.panel.send({ type: 'flushed', id: request.id });
  await saving;
  assert.equal(released, true);
  assert.equal(opened.document.getText(), '# Title\n\nhello there\n\n- item\n', 'the typing was applied before the save went on');
});

test('Save All without an answer: the save goes on after 1 second and the status bar says the typing was not saved yet', async () => {
  const opened = await openMd('noflush.md', DOC);
  await setMode(opened, 'wysiwyg');
  const t0 = Date.now();
  await T.fireWillSave(opened.document);
  const took = Date.now() - t0;
  assert.ok(took >= 900 && took < 2500, `waited ${took} ms`);
  assert.match(callsOf('statusBar').at(-1)[1], /last typing was not saved yet/);
});

test('Preview mode: a save does not wait for any flush', async () => {
  const opened = await openMd('previewsave.md', DOC);
  opened.panel.posted.length = 0;
  const t0 = Date.now();
  await T.fireWillSave(opened.document);
  assert.ok(Date.now() - t0 < 200);
  assert.equal(lastOf(opened.panel, 'flushRequest'), undefined);
});

// ===== Split =====

test('Split opens the source beside the preview once; a visible editor is reused', async () => {
  const opened = await openMd('split.md', DOC);
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await waitFor(() => callsOf('showTextDocument').length === 1, 3000, 'showTextDocument');
  assert.deepEqual(callsOf('showTextDocument')[0].slice(2), [vscode.ViewColumn.Beside, true, false]);
  opened.panel.send({ type: 'setMode', mode: 'preview' });
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await tick(50);
  assert.equal(callsOf('showTextDocument').length, 1, 'the editor that is open is reused');
});

test('Split: scrolling the editor reveals the line in the preview; the preview scrolling reveals it in the editor, echo dropped', async () => {
  const opened = await openMd('sync.md', Array.from({ length: 50 }, (_, i) => `line ${i}\n`).join('\n'));
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await waitFor(() => callsOf('showTextDocument').length === 1, 3000, 'showTextDocument');
  const editor = vscode.window.visibleTextEditors[0];
  opened.panel.posted.length = 0;
  T.fireVisibleRanges(editor, 12);
  assert.deepEqual(lastOf(opened.panel, 'revealLine'), { type: 'revealLine', line: 12 });
  // Preview -> editor.
  opened.panel.send({ type: 'previewScrolled', line: 30.6 });
  await waitFor(() => callsOf('revealRange').length === 1, 3000, 'revealRange');
  assert.deepEqual(callsOf('revealRange')[0].slice(1), [30, vscode.TextEditorRevealType.AtTop]);
  // The editor's own scroll event within 150 ms is the echo: not sent back.
  opened.panel.posted.length = 0;
  T.fireVisibleRanges(editor, 30);
  assert.equal(lastOf(opened.panel, 'revealLine'), undefined, 'echo dropped');
  await tick(200);
  T.fireVisibleRanges(editor, 33);
  assert.equal(lastOf(opened.panel, 'revealLine').line, 33);
});

test('Split: scroll events of another document and of a panel in Preview mode are ignored', async () => {
  const opened = await openMd('only.md', DOC);
  const otherDoc = new T.FakeTextDocument(writeTmp('x.md', 'x'), 'x');
  const otherEditor = T.makeEditor(otherDoc);
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await waitFor(() => callsOf('showTextDocument').length === 1, 3000, 'showTextDocument');
  opened.panel.posted.length = 0;
  T.fireVisibleRanges(otherEditor, 5);
  assert.equal(lastOf(opened.panel, 'revealLine'), undefined);
  opened.panel.send({ type: 'setMode', mode: 'preview' });
  await tick(20);
  opened.panel.posted.length = 0;
  T.fireVisibleRanges(vscode.window.visibleTextEditors[0], 7);
  assert.equal(lastOf(opened.panel, 'revealLine'), undefined, 'not in Split mode');
});

test('Split works for a read-only document (the text editor is VS Code\'s own)', async () => {
  T.config.set('files.readonlyInclude', { '**/locked/**': true });
  const opened = await openMd('locked/split.md', DOC);
  opened.panel.send({ type: 'setMode', mode: 'split' });
  await waitFor(() => callsOf('showTextDocument').length === 1, 3000, 'showTextDocument');
});

// ===== fidelity (contract §8.6, D-072) =====

test('an edit before the save check has finished is refused (fidelity) and never applied', async () => {
  const uri = writeTmp('pending.md', DOC);
  const document = new T.FakeTextDocument(uri, DOC);
  T.textDocuments.set(uri.toString(), document);
  const original = vscode.workspace.fs.readFile;
  let release = () => undefined;
  const gate = new Promise((resolve) => (release = resolve));
  vscode.workspace.fs.readFile = async (u) => {
    await gate;
    return original(u);
  };
  try {
    const panel = T.createPanel();
    new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
    panel.send({ type: 'ready' });
    await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
    const opened = { uri, document, panel };
    const refused = await sendEdit(opened, 'world', 'there');
    assert.deepEqual([refused.accepted, refused.code], [false, 'fidelity']);
    assert.equal(refused.reason, 'FileStudio is still checking this file; try again in a moment.');
    assert.equal(document.getText(), DOC);
    assert.deepEqual(callsOf('applyEdit'), []);
    // A task toggle is exempt from the check (0.1.2 behaviour).
    release();
    await settled(panel);
    const ok = await sendEdit(opened, 'world', 'there');
    assert.equal(ok.accepted, true);
  } finally {
    release();
    vscode.workspace.fs.readFile = original;
  }
});

test('task toggles are exempt from the save check (a file with mixed line breaks still toggles)', async () => {
  const text = '- [ ] a\r\n- [ ] b\r\n';
  const opened = await openMd('mixed-task.md', text, { disk: '- [ ] a\r\n- [ ] b\n' });
  assert.equal(opened.init.editState.editable, true);
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'toggleTask', line: 0, checked: true, version: opened.document.version });
  const result = await waitFor(() => lastOf(opened.panel, 'toggleTaskResult'), 3000, 'toggleTaskResult');
  assert.equal(result.applied, true);
  assert.equal(opened.document.getText(), '- [x] a\r\n- [ ] b\r\n');
});

test('a file VS Code would save with other line breaks: WYSIWYG read-only with "Edit anyway"; the host asks; accepted once per session', async () => {
  const text = 'a\r\n\r\nb\r\n';
  const opened = await openMd('mixed.md', text, { disk: 'a\r\n\nb\r\n' });
  const state = await settled(opened.panel);
  assert.equal(state.editable, true, 'the checkboxes work');
  assert.deepEqual(state.kinds, []);
  assert.equal(state.code, 'fidelity');
  assert.ok(state.canEditAnyway.details.some((d) => /1 line break \(LF → CRLF\)/.test(d)), state.canEditAnyway.details.join('|'));
  await setMode(opened, 'wysiwyg'); // read-only WYSIWYG is shown (the webview sets it read-only)
  assert.equal(lastOf(opened.panel, 'markdownUpdate').html, '');
  const refused = await sendEdit(opened, 'b', 'bb');
  assert.deepEqual([refused.accepted, refused.code], [false, 'fidelity']);
  assert.match(refused.reason, /Choose "Edit anyway"/);
  // "Edit anyway" without an answer in the test queue = Cancel: still refused.
  opened.panel.send({ type: 'editAnyway' });
  await tick(30);
  assert.equal(opened.document.getText(), text);
  const stillRefused = await sendEdit(opened, 'b', 'bb');
  assert.equal(stillRefused.code, 'fidelity');
  // The user accepts: edits work (the host asked itself; the message alone was never consent).
  const api = host.createTestApi();
  api.answer('fidelity', 'Edit anyway');
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'editAnyway' });
  const accepted = await waitFor(() => [...opened.panel.posted].reverse().find((m) => m.editState && m.editState.kinds?.includes('mdText')), 3000, 'accepted state');
  assert.deepEqual(accepted.editState, { editable: true, kinds: ['mdText'] });
  const ok = await sendEdit(opened, 'b', 'bb');
  assert.equal(ok.accepted, true);
  assert.equal(opened.document.getText(), 'a\r\n\r\nbb\r\n');
});

test('wrong encoding (U+FFFD that is not in the bytes): WYSIWYG is blocked, no "Edit anyway"', async () => {
  const opened = await openMd('enc.md', 'caf�\n', { disk: Buffer.from([0x63, 0x61, 0x66, 0xe9, 0x0a]) });
  const state = await settled(opened.panel);
  assert.equal(state.editable, true);
  assert.deepEqual(state.kinds, []);
  assert.equal(state.code, 'blocked');
  assert.match(state.reason, /not valid utf8 text/);
  assert.equal(state.canEditAnyway, undefined);
  const refused = await sendEdit(opened, 'caf', 'cafe');
  assert.deepEqual([refused.accepted, refused.code], [false, 'blocked']);
});

test('re-check after a reload from disk (clean document): the old verdict and "Edit anyway" are dropped when the result differs (D-072)', async () => {
  const opened = await openMd('reload.md', 'a\n\nb\n');
  assert.deepEqual((await settled(opened.panel)).kinds, ['mdText']);
  // VS Code reloads the file from disk (the document stays clean) and it now has mixed line breaks.
  fs.writeFileSync(opened.uri.fsPath, 'a\r\n\nb\n');
  opened.document.eol = vscode.EndOfLine.CRLF;
  opened.panel.posted.length = 0;
  opened.document.setText('a\r\n\r\nb\r\n', false);
  const state = await waitFor(
    () => [...opened.panel.posted].reverse().find((m) => m.editState?.code === 'fidelity')?.editState,
    3000,
    'fidelity state after the reload',
  );
  assert.equal(state.editable, true);
  assert.deepEqual(state.kinds, []);
  const refused = await sendEdit(opened, 'b', 'bb');
  assert.deepEqual([refused.accepted, refused.code], [false, 'fidelity']);
});

test('re-check: an edit that arrives while the re-check runs waits for it and is judged by the new verdict', async () => {
  const opened = await openMd('wait.md', 'a\n\nb\n');
  await setMode(opened, 'wysiwyg');
  fs.writeFileSync(opened.uri.fsPath, 'a\r\n\nb\n');
  opened.document.eol = vscode.EndOfLine.CRLF;
  opened.document.setText('a\r\n\r\nb\r\n', false); // reload from disk; the re-check is scheduled
  const refused = await sendEdit(opened, 'b', 'bb'); // sent at once: must not be judged by the old "ok"
  assert.deepEqual([refused.accepted, refused.code], [false, 'fidelity']);
  assert.equal(opened.document.getText(), 'a\r\n\r\nb\r\n');
});

test('re-check with the same verdict keeps "Edit anyway"', async () => {
  const opened = await openMd('keep.md', 'a\r\n\r\nb\r\n', { disk: 'a\r\n\nb\r\n' });
  await settled(opened.panel);
  const api = host.createTestApi();
  api.answer('fidelity', 'Edit anyway');
  opened.panel.send({ type: 'editAnyway' });
  await waitFor(() => [...opened.panel.posted].reverse().find((m) => m.editState?.kinds?.includes('mdText')), 3000, 'accepted');
  // A reload that leaves the same situation on disk.
  opened.document.setText('a\r\n\r\nb\r\n', false);
  await tick(700);
  const ok = await sendEdit(opened, 'b', 'bb');
  assert.equal(ok.accepted, true);
});

// ===== own edits and toggles =====

test('a toggle in Split/Preview still pushes the new rendering at once (own change recognised by its text)', async () => {
  const opened = await openMd('toggle.md', '- [ ] a\n- [ ] b\n');
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'toggleTask', line: 0, checked: true, version: opened.document.version });
  await waitFor(() => lastOf(opened.panel, 'toggleTaskResult'), 3000, 'result');
  const update = lastOf(opened.panel, 'markdownUpdate');
  assert.ok(update, 'pushed at once');
  assert.equal(update.version, opened.document.version);
  assert.equal(update.source, '- [x] a\n- [ ] b\n');
});
