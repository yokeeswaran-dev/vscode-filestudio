// Markdown modes in a REAL VS Code (v0.2.0, P6): the host side with VS Code's own APIs (text codec for the save check,
// applyEdit, save participants, undo, text editors beside). The webview is the real one; its typing cannot be simulated,
// so edits are delivered through the test API (`post`) as if the webview had sent them.
'use strict';
const assert = require('node:assert/strict');
const vscode = require('vscode');
const h = require('./helpers');

const DOC = '# Title\n\nhello world\n\n- item\n';

async function api() {
  const ext = h.extension();
  assert.ok(ext, 'extension loaded');
  const exp = await ext.activate();
  const value = exp ?? ext.exports;
  assert.ok(value && typeof value.post === 'function', 'the test API is available');
  return value;
}

/**
 * Waits until the real webview has reported 'ready'. A 'ready' starts a new life of the panel (the answers to edits that
 * were sent before it are dropped on purpose), so a test must not post before it. The host answers a mode request with a
 * rendering only once the webview is ready.
 */
async function webviewReady(t, uri) {
  await h.waitFor(
    async () => (await t.post(uri.toString(), { type: 'setMode', mode: 'preview' })).some((m) => m.type === 'markdownUpdate' || m.type === 'init'),
    { timeout: 30_000, message: 'the webview to be ready' },
  );
}

/** A copy of a Markdown text in the workspace, opened in FileStudio; returns the document and the test API. */
async function openCopy(name, bytes) {
  const uri = h.writeBytes(name, bytes);
  await h.openWith(uri);
  const document = await vscode.workspace.openTextDocument(uri);
  const t = await api();
  await h.waitFor(() => {
    try {
      return t.state(uri.toString()).editState && t.state(uri.toString()).editState.kinds?.includes('mdText');
    } catch {
      return false;
    }
  }, { timeout: 15_000, message: 'the save check to finish (kinds has mdText)' });
  await webviewReady(t, uri);
  return { uri, document, t };
}

let editId = 1000;
function edit(document, oldText, text) {
  const start = document.getText().indexOf(oldText);
  assert.ok(start >= 0);
  return { type: 'edit', editId: editId++, baseVersion: document.version, intent: { kind: 'mdText', edits: [{ start, end: start + oldText.length, oldText, text }] } };
}
const resultOf = (messages) => messages.find((m) => m.type === 'editResult');

suite('Markdown WYSIWYG host (real VS Code)', () => {
  teardown(() => h.closeAll());

  test('open -> every mode -> save: the bytes are identical (LF, CRLF, BOM, no final newline, tabs, trailing spaces)', async () => {
    const samples = {
      'zz-lf.md': Buffer.from(DOC),
      'zz-crlf.md': Buffer.from(DOC.replace(/\n/g, '\r\n')),
      'zz-bom.md': Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(DOC)]),
      'zz-nofinal.md': Buffer.from('# T\n\nabc'),
      'zz-tabs.md': Buffer.from('# T\n\n\tindented\ttab  \n\ntrailing spaces   \n'),
    };
    for (const [name, bytes] of Object.entries(samples)) {
      const { uri, document, t } = await openCopy(name, bytes);
      try {
        for (const mode of ['split', 'wysiwyg', 'preview']) {
          await t.post(uri.toString(), { type: 'setMode', mode });
        }
        await vscode.commands.executeCommand('workbench.action.files.save');
        await h.sleep(500);
        assert.ok(h.readBytes(uri).equals(bytes), `${name}: bytes changed`);
        assert.equal(document.isDirty, false, `${name}: dirty`);
      } finally {
        await h.closeAll();
        h.deleteFile(uri);
      }
    }
  });

  test('WYSIWYG edit through the host: only the typed word changes; undo restores the bytes; save writes it', async () => {
    const { uri, document, t } = await openCopy('zz-edit.md', Buffer.from(DOC));
    try {
      await t.post(uri.toString(), { type: 'setMode', mode: 'wysiwyg' });
      const result = resultOf(await t.post(uri.toString(), edit(document, 'world', 'there')));
      assert.equal(result.accepted, true);
      assert.equal(document.getText(), DOC.replace('world', 'there'));
      assert.equal(document.isDirty, true);
      await h.focusTab(h.findTab(uri));
      await t.post(uri.toString(), { type: 'command', id: 'undo' });
      await h.waitFor(() => document.getText() === DOC, { message: 'undo' });
      await h.waitFor(() => !document.isDirty, { message: 'clean after undo' });
      await t.post(uri.toString(), edit(document, 'world', 'again'));
      await t.post(uri.toString(), { type: 'command', id: 'save' });
      await h.waitFor(() => h.readBytes(uri).toString() === DOC.replace('world', 'again'), { message: 'the saved file' });
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('a stale edit and an edit that would change another block are refused; nothing changes', async () => {
    const { uri, document, t } = await openCopy('zz-refuse.md', Buffer.from('Intro.\n\nA paragraph\n- item\n'));
    try {
      const stale = { ...edit(document, 'Intro', 'Outro'), baseVersion: document.version - 1 };
      assert.equal(resultOf(await t.post(uri.toString(), stale)).code, 'stale');
      const blocked = resultOf(await t.post(uri.toString(), edit(document, '- item', 'item')));
      assert.deepEqual([blocked.accepted, blocked.code], [false, 'blocked']);
      assert.equal(document.getText(), 'Intro.\n\nA paragraph\n- item\n');
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('mixed line breaks: the save check (real VS Code text codec) offers "Edit anyway"; accepted, the edit works', async () => {
    const uri = h.writeBytes('zz-mixed.md', Buffer.from('a\r\n\nb\r\n'));
    await h.openWith(uri);
    const document = await vscode.workspace.openTextDocument(uri);
    const t = await api();
    try {
      const state = await h.waitFor(() => {
        try {
          const s = t.state(uri.toString()).editState;
          return s && s.code === 'fidelity' && s.canEditAnyway ? s : undefined;
        } catch {
          return undefined;
        }
      }, { timeout: 15_000, message: 'the fidelity state' });
      assert.deepEqual(state.kinds, []);
      await webviewReady(t, uri);
      const refused = resultOf(await t.post(uri.toString(), edit(document, 'b', 'bb')));
      assert.equal(refused.code, 'fidelity');
      t.answer('fidelity', 'Edit anyway');
      await t.post(uri.toString(), { type: 'editAnyway' });
      await h.waitFor(() => t.state(uri.toString()).editState?.kinds?.includes('mdText'), { message: 'accepted' });
      const answered = await t.post(uri.toString(), edit(document, 'b', 'bb'));
      assert.equal(resultOf(answered)?.accepted, true, JSON.stringify(answered).slice(0, 600));
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('Split opens the source in a text editor beside and reuses it', async () => {
    const { uri, t } = await openCopy('zz-split.md', Buffer.from(DOC));
    try {
      await t.post(uri.toString(), { type: 'setMode', mode: 'split' });
      await h.waitFor(() => vscode.window.visibleTextEditors.some((e) => e.document.uri.toString() === uri.toString()), { message: 'the text editor' });
      const count = () => vscode.window.visibleTextEditors.filter((e) => e.document.uri.toString() === uri.toString()).length;
      assert.equal(count(), 1);
      await t.post(uri.toString(), { type: 'setMode', mode: 'preview' });
      await t.post(uri.toString(), { type: 'setMode', mode: 'split' });
      await h.sleep(500);
      assert.equal(count(), 1, 'reused');
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('Save All with a WYSIWYG panel: the real webview answers the flush request and the save goes on', async () => {
    const { uri, document, t } = await openCopy('zz-flush.md', Buffer.from(DOC));
    try {
      await t.post(uri.toString(), { type: 'setMode', mode: 'wysiwyg' });
      await t.post(uri.toString(), edit(document, 'world', 'there'));
      const started = Date.now();
      await vscode.workspace.saveAll();
      await h.waitFor(() => h.readBytes(uri).toString() === DOC.replace('world', 'there'), { message: 'saved' });
      assert.ok(Date.now() - started < 5000);
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('Save All while an edit is still on its way: the save waits for the edit (the typing is in the saved file)', async () => {
    const { uri, document, t } = await openCopy('zz-flush2.md', Buffer.from(DOC));
    try {
      await t.post(uri.toString(), { type: 'setMode', mode: 'wysiwyg' });
      // (Save All only saves dirty documents: a first, finished edit makes it dirty)
      assert.equal(resultOf(await t.post(uri.toString(), edit(document, 'world', 'there'))).accepted, true);
      // the second edit is sent but not awaited: it is queued when VS Code starts the save; the real webview's `flushed`
      // reply must not release the save before that edit is applied (contract §9.3)
      const posted = t.post(uri.toString(), edit(document, 'item', 'thing'));
      await vscode.workspace.saveAll();
      assert.equal(resultOf(await posted).accepted, true);
      const wanted = DOC.replace('world', 'there').replace('item', 'thing');
      await h.waitFor(() => h.readBytes(uri).toString() === wanted, { message: 'the saved file has both edits' });
      assert.equal(document.isDirty, false);
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('undo is only run when the FileStudio tab is the active one (two editor groups)', async () => {
    const { uri, document, t } = await openCopy('zz-guard.md', Buffer.from(DOC));
    const other = h.writeBytes('zz-guard-other.txt', Buffer.from('another file\n'));
    try {
      await t.post(uri.toString(), { type: 'setMode', mode: 'wysiwyg' });
      assert.equal(resultOf(await t.post(uri.toString(), edit(document, 'world', 'there'))).accepted, true);
      // another editor in the group beside takes the focus: VS Code's undo would act on THAT editor
      await vscode.window.showTextDocument(other, { viewColumn: vscode.ViewColumn.Beside, preview: false });
      const active = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
      assert.ok(active instanceof vscode.TabInputText, 'the other file is active');
      await t.post(uri.toString(), { type: 'command', id: 'undo' });
      await h.sleep(600);
      assert.equal(document.getText(), DOC.replace('world', 'there'), 'nothing was undone');
      assert.equal(await vscode.workspace.openTextDocument(other).then((d) => d.getText()), 'another file\n', 'the other file is untouched');
      // the other editor has an edit of its own to undo: VS Code's undo must not reach it from the FileStudio tab
      const otherDoc = await vscode.workspace.openTextDocument(other);
      const typed = new vscode.WorkspaceEdit();
      typed.insert(other, new vscode.Position(0, 0), 'X');
      assert.equal(await vscode.workspace.applyEdit(typed), true);
      await h.focusTab(h.findTab(uri));
      await h.waitFor(() => {
        const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
        return input instanceof vscode.TabInputCustom && input.uri.toString() === uri.toString();
      }, { message: 'the FileStudio tab to be the active one' });
      // (in real use the user's key press is in the webview, which has the keyboard focus; here VS Code moves the focus
      // a moment after the tab is shown, and the undo command goes to whatever has the focus)
      await h.sleep(1500);
      await t.post(uri.toString(), { type: 'command', id: 'undo' });
      await h.sleep(1500);
      const now = JSON.stringify({ md: document.getText(), other: otherDoc.getText() });
      assert.equal(document.getText(), DOC, 'undo on the active FileStudio tab: ' + now);
      assert.equal(otherDoc.getText(), 'Xanother file\n', 'the other file keeps its edit (the undo did not reach it): ' + now);
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
      h.deleteFile(other);
    }
  });

  test('read-only (files.readonlyInclude): edits are refused and WYSIWYG is not offered', async () => {
    const config = vscode.workspace.getConfiguration('files');
    const { uri, document, t } = await openCopy('zz-ro.md', Buffer.from(DOC));
    try {
      await config.update('readonlyInclude', { '**/zz-ro.md': true }, vscode.ConfigurationTarget.Workspace);
      await h.sleep(500);
      const refused = resultOf(await t.post(uri.toString(), edit(document, 'world', 'there')));
      assert.deepEqual([refused.accepted, refused.code], [false, 'readOnly']);
    } finally {
      await config.update('readonlyInclude', undefined, vscode.ConfigurationTarget.Workspace);
      await h.closeAll();
      h.deleteFile(uri);
    }
  });

  test('kill switch: WYSIWYG edits are refused (disabled); Split still opens the editor', async () => {
    const config = vscode.workspace.getConfiguration('fileStudio');
    const { uri, document, t } = await openCopy('zz-kill.md', Buffer.from(DOC));
    try {
      await config.update('markdown.editing', false, vscode.ConfigurationTarget.Workspace);
      await h.sleep(500);
      const refused = resultOf(await t.post(uri.toString(), edit(document, 'world', 'there')));
      assert.deepEqual([refused.accepted, refused.code], [false, 'disabled']);
      await t.post(uri.toString(), { type: 'setMode', mode: 'split' });
      await h.waitFor(() => vscode.window.visibleTextEditors.some((e) => e.document.uri.toString() === uri.toString()), { message: 'the text editor' });
    } finally {
      await config.update('markdown.editing', undefined, vscode.ConfigurationTarget.Workspace);
      await h.closeAll();
      h.deleteFile(uri);
    }
  });
});
