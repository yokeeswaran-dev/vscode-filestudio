// Track F (contract §3, §8.6, §13.3, §14 F WP7): SHARED: EDIT GATE and TEST API against the fake `vscode`.
// Check order (kill switch → read-only → document state), the kill switch read on every call, SerialQueue (FIFO, a
// failing job does not stop it), ask() answered by the test API (none queued = Cancel), checkTextFidelity (clean,
// mixed EOL, CR-only, double BOM, trailing whitespace, final newlines, language-scoped settings, untitled:, unreadable
// bytes, wrong encoding, dirty at open) and the `pending` state; Markdown task toggles behind the gate in one queue.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');

const { vscode, T, host, extensionUri, token, tick, writeTmp, lastOf, callsOf, waitFor } = require('./load-host');

const { EDIT_TEXT } = host;
const api = host.createTestApi(); // test mode for this file: ask() takes queued answers

/** The fake has no text codec: a UTF-8 one like VS Code's (decode strips one BOM; invalid bytes become U+FFFD). */
const codec = {
  encode: async (text, { encoding }) => {
    const body = Buffer.from(text, 'utf8');
    return new Uint8Array(encoding === 'utf8bom' ? Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), body]) : body);
  },
  decode: async (bytes) => new TextDecoder('utf-8').decode(bytes),
};
const originalGetConfiguration = vscode.workspace.getConfiguration;

beforeEach(() => {
  T.reset();
  vscode.workspace.encode = codec.encode;
  vscode.workspace.decode = codec.decode;
  vscode.workspace.getConfiguration = originalGetConfiguration;
});

let fileCounter = 0;
const unique = (name) => `${++fileCounter}-${name}`;

// ===== SERIAL QUEUE =====

test('SerialQueue runs jobs one at a time in FIFO order; a failing job does not stop it', async () => {
  const queue = new host.SerialQueue();
  const log = [];
  const job = (name, ms) => async () => {
    log.push(`start ${name}`);
    await tick(ms);
    log.push(`end ${name}`);
    return name;
  };
  const a = queue.run(job('a', 30));
  const b = queue.run(() => {
    throw new Error('sync failure');
  });
  const c = queue.run(async () => {
    await tick(5);
    throw new Error('async failure');
  });
  const d = queue.run(job('d', 1));
  const e = queue.run(() => 'e');
  assert.equal(queue.size, 5);
  assert.equal(await a, 'a');
  await assert.rejects(b, /sync failure/);
  await assert.rejects(c, /async failure/);
  assert.equal(await d, 'd');
  assert.equal(await e, 'e');
  assert.deepEqual(log, ['start a', 'end a', 'start d', 'end d']);
  assert.equal(queue.size, 0);
});

// ===== KILL SWITCH =====

test('the kill switch is read on every call: master switch first, then the format switch', () => {
  const uri = vscode.Uri.file('/w/a.csv');
  assert.equal(host.editingEnabled('csv', uri), true, 'default: on');
  T.config.set('fileStudio.csv.editing', false);
  assert.equal(host.editingEnabled('csv', uri), false);
  assert.equal(host.disabledEditSetting('csv', uri), 'fileStudio.csv.editing');
  assert.equal(host.editingEnabled('xlsx', uri), true, 'only the CSV switch is off');
  T.config.set('fileStudio.editing.enabled', false);
  assert.equal(host.disabledEditSetting('csv', uri), 'fileStudio.editing.enabled', 'the master switch is named first');
  assert.equal(host.editingEnabled('markdown', uri), false);
  T.config.set('fileStudio.editing.enabled', true);
  T.config.set('fileStudio.csv.editing', true);
  assert.equal(host.editingEnabled('csv', uri), true, 'no cache');
});

// ===== GATE ORDER =====

test('checkEditable order: kill switch, then read-only (with the reason), then the document state', async () => {
  const uri = writeTmp(`locked/${unique('book.xlsx')}`, 'x');
  T.config.set('files.readonlyInclude', { '**/locked/**': true });
  T.config.set('fileStudio.xlsx.editing', false);
  assert.deepEqual(await host.checkEditable('xlsx', uri, {}), {
    ok: false, code: 'disabled', reason: 'Editing is turned off (setting fileStudio.xlsx.editing).',
  });
  T.config.set('fileStudio.xlsx.editing', true);
  // An xlsx has no TextDocument: the glob check uses a {uri, languageId} stand-in.
  assert.equal(await host.isReadonlyResource(uri), true);
  assert.deepEqual(await host.checkEditable('xlsx', uri, { poisoned: true }), {
    ok: false, code: 'readOnly', reason: 'This file is read-only (files.readonlyInclude).',
  });
  T.config.set('files.readonlyExclude', { '**/*book.xlsx': true });
  assert.deepEqual(await host.checkEditable('xlsx', uri, {}), { ok: true });
  // Document state, in order: poisoned, lossy, fidelity (pending first), hard block.
  const gate = (s) => host.checkEditable('xlsx', uri, s);
  assert.deepEqual(await gate({ poisoned: true, lossy: { accepted: false } }), { ok: false, code: 'internal', reason: EDIT_TEXT.poisoned });
  assert.deepEqual(await gate({ lossy: { accepted: false }, block: 'x' }), { ok: false, code: 'lossy', reason: EDIT_TEXT.lossy });
  assert.deepEqual(await gate({ lossy: { accepted: true } }), { ok: true });
  const pending = { pending: true, ok: false, accepted: false, blocked: 'never' };
  assert.deepEqual(await gate({ fidelity: pending }), { ok: false, code: 'fidelity', reason: EDIT_TEXT.fidelityPending });
  assert.deepEqual(await gate({ fidelity: { ok: false, accepted: true, blocked: 'Wrong encoding.' } }), { ok: false, code: 'blocked', reason: 'Wrong encoding.' });
  assert.deepEqual(await gate({ fidelity: { ok: false, accepted: false } }), { ok: false, code: 'fidelity', reason: EDIT_TEXT.fidelity });
  assert.deepEqual(await gate({ fidelity: { ok: false, accepted: true } }), { ok: true });
  assert.deepEqual(await gate({ block: 'Unclosed quote.' }), { ok: false, code: 'blocked', reason: 'Unclosed quote.' });
});

test('read-only reasons: a Git version and a file without write permission', async (t) => {
  const git = vscode.Uri.parse('git:/repo/a.md?%7B%7D');
  assert.deepEqual(await host.checkEditable('markdown', git, {}), { ok: false, code: 'readOnly', reason: 'This file is read-only (Git version).' });
  const uri = writeTmp(unique('locked.csv'), 'a\n');
  fs.chmodSync(uri.fsPath, 0o444);
  t.after(() => fs.chmodSync(uri.fsPath, 0o666));
  assert.deepEqual(await host.checkEditable('csv', uri, {}), { ok: true }, 'files.readonlyFromPermissions is off');
  T.config.set('files.readonlyFromPermissions', true);
  assert.deepEqual(await host.checkEditable('csv', uri, {}), { ok: false, code: 'readOnly', reason: 'This file is read-only (no write permission).' });
});

// ===== ASK (TEST API ANSWERS) =====

test('ask() takes the queued test answers in order; none queued (or not a button) = Cancel, never a dialog', async () => {
  api.answer('editAnyway', 'Edit Anyway');
  api.answer('editAnyway', undefined);
  api.answer('editAnyway', 'Delete Everything');
  api.answer('bakFailed', 'Save Anyway');
  const ask = (id) => host.ask(id, 'Question?', { modal: true }, 'Edit Anyway', 'Save Anyway');
  assert.equal(await ask('editAnyway'), 'Edit Anyway');
  assert.equal(await ask('bakFailed'), 'Save Anyway', 'queues are per prompt id');
  assert.equal(await ask('editAnyway'), undefined, 'queued Cancel');
  assert.equal(await ask('editAnyway'), undefined, 'not one of the buttons = Cancel');
  assert.equal(await ask('editAnyway'), undefined, 'nothing queued = Cancel');
  assert.deepEqual(callsOf('warning'), [], 'no dialog was shown');
});

// ===== TEXT FIDELITY =====

/**
 * A document as VS Code would load `bytes`: every leading BOM stripped, line breaks normalized to `eol` (s2-protocol
 * E5), `encoding` as detected. `text` overrides the model text (a dirty document).
 */
function fidelityDoc(name, bytes, { eol = 'LF', encoding = 'utf8', dirty = false, text } = {}) {
  const uri = writeTmp(unique(name), Buffer.from(bytes));
  let model = text ?? new TextDecoder('utf-8', { ignoreBOM: true }).decode(Buffer.from(bytes)).replace(/^﻿+/, '');
  if (text === undefined) model = model.replace(/\r\n|\r|\n/g, eol === 'CRLF' ? '\r\n' : '\n');
  const document = new T.FakeTextDocument(uri, model);
  Object.assign(document, { eol: eol === 'CRLF' ? vscode.EndOfLine.CRLF : vscode.EndOfLine.LF, encoding, isDirty: dirty });
  return document;
}
const bom = [0xef, 0xbb, 0xbf];
const bytesOf = (text) => [...Buffer.from(text, 'utf8')];
const changes = (...details) => ({ ok: false, canEditAnyway: { title: EDIT_TEXT.fidelityTitle, details } });

test('checkTextFidelity: files VS Code writes back unchanged are ok', async () => {
  assert.deepEqual(await host.checkTextFidelity(fidelityDoc('lf.csv', bytesOf('a,b\n1,2\n'))), { ok: true });
  assert.deepEqual(await host.checkTextFidelity(fidelityDoc('crlf.csv', bytesOf('a\r\nb\r\n'), { eol: 'CRLF' })), { ok: true });
  assert.deepEqual(await host.checkTextFidelity(fidelityDoc('bom.csv', [...bom, ...bytesOf('a\n')], { encoding: 'utf8bom' })), { ok: true });
  // A real U+FFFD stored in the file is not a decoding error.
  assert.deepEqual(await host.checkTextFidelity(fidelityDoc('fffd.md', bytesOf('a � b\n'))), { ok: true });
  // untitled: nothing on disk to change (and nothing is read).
  const untitled = new T.FakeTextDocument(vscode.Uri.parse('untitled:Untitled-1'), 'a\r\nb\n');
  assert.deepEqual(await host.checkTextFidelity(untitled), { ok: true });
});

test('checkTextFidelity: mixed line breaks and CR-only are counted (Edit anyway); a double BOM is not a change', async () => {
  assert.deepEqual(
    await host.checkTextFidelity(fidelityDoc('mixed.csv', bytesOf('a\r\n"x\ny"\r\nc\r\n'), { eol: 'CRLF' })),
    changes('VS Code will change 1 line break (LF → CRLF).'),
  );
  assert.deepEqual(
    await host.checkTextFidelity(fidelityDoc('cr.md', bytesOf('a\rb\r\nc\r'), { eol: 'LF' })),
    changes('VS Code will change 1 line break (CRLF → LF).', 'VS Code will change 2 line breaks (CR → LF).'),
  );
  assert.deepEqual(
    await host.checkTextFidelity(fidelityDoc('bom2.csv', [...bom, ...bom, ...bytesOf('a\n')], { encoding: 'utf8bom' })),
    { ok: true }, // (fix round 2: real VS Code 1.142 writes both BOMs back unchanged)
  );
});

test('checkTextFidelity: save settings are counted, read for the document (language overrides)', async () => {
  const trailing = fidelityDoc('trail.csv', bytesOf('a  \nb\t\nc\n\n\n'));
  assert.deepEqual(await host.checkTextFidelity(trailing), { ok: true }, 'settings off');
  T.config.set('files.trimTrailingWhitespace', true);
  T.config.set('files.trimFinalNewlines', true);
  assert.deepEqual(
    await host.checkTextFidelity(trailing),
    changes(
      'VS Code will remove spaces at the end of 2 lines (setting files.trimTrailingWhitespace).',
      'VS Code will remove 2 empty lines at the end of the file (setting files.trimFinalNewlines).',
    ),
  );
  T.config.clear();
  T.config.set('files.insertFinalNewline', true);
  assert.deepEqual(
    await host.checkTextFidelity(fidelityDoc('nofinal.md', bytesOf('a\nb'))),
    changes('VS Code will add a line break at the end of the file (setting files.insertFinalNewline).'),
  );
  T.config.clear();
  // A [markdown] override: the settings are read with the document as scope.
  const doc = fidelityDoc('override.md', bytesOf('a  \n'));
  const scopes = [];
  vscode.workspace.getConfiguration = (section, scope) => {
    if (section === 'files') scopes.push(scope);
    return section === 'files' && scope === doc ? { get: (key) => key === 'trimTrailingWhitespace' } : originalGetConfiguration(section, scope);
  };
  assert.deepEqual(await host.checkTextFidelity(doc), changes('VS Code will remove spaces at the end of 1 line (setting files.trimTrailingWhitespace).'));
  assert.ok(scopes.includes(doc), 'getConfiguration("files", document)');
});

test('checkTextFidelity: wrong encoding is blocked; unreadable bytes and a missing API offer Edit anyway', async () => {
  const latin1 = fidelityDoc('latin1.csv', [0x63, 0x61, 0x66, 0xe9, 0x0a]); // "café" in Latin-1, read as UTF-8
  assert.match(latin1.getText(), /�/);
  const blocked = await host.checkTextFidelity(latin1);
  assert.equal(blocked.ok, false);
  assert.match(blocked.blocked, /^FileStudio cannot edit this file safely: it is not valid utf8 text\./);
  assert.equal(blocked.canEditAnyway, undefined, 'no Edit anyway');

  const gone = fidelityDoc('gone.csv', bytesOf('a\n'));
  fs.rmSync(gone.uri.fsPath);
  assert.deepEqual(await host.checkTextFidelity(gone), changes(EDIT_TEXT.fidelityUnreadable));

  delete vscode.workspace.encode;
  assert.deepEqual(await host.checkTextFidelity(fidelityDoc('noapi.csv', bytesOf('a\n'))), changes(EDIT_TEXT.fidelityUnreadable));
});

test('checkTextFidelity: a document dirty at open is judged by the line breaks on disk', async () => {
  const mixed = fidelityDoc('dirty-mixed.csv', bytesOf('a\r\nb\nc\r\n'), { eol: 'CRLF', dirty: true, text: 'edited\r\n' });
  assert.deepEqual(await host.checkTextFidelity(mixed), changes('VS Code will change 1 line break (LF → CRLF).'));
  const clean = fidelityDoc('dirty-clean.csv', bytesOf('a\nb\n'), { dirty: true, text: 'a\nedited\n' });
  assert.deepEqual(await host.checkTextFidelity(clean), { ok: true });
});

// ===== MARKDOWN TASK TOGGLES BEHIND THE GATE =====

function openMarkdown(name, text, column = 1) {
  const uri = writeTmp(name, text);
  const document = T.textDocuments.get(uri.toString()) ?? new T.FakeTextDocument(uri, text);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel(column);
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
  panel.send({ type: 'ready' });
  return { uri, document, panel };
}

async function toggle(panel, line, checked, version) {
  panel.posted.length = 0;
  panel.send({ type: 'toggleTask', line, checked, version });
  return waitFor(() => lastOf(panel, 'toggleTaskResult'), 3000, 'toggleTaskResult');
}

test('kill switch: a task toggle is refused (with the reason) when Markdown editing is off, applied when on', async (t) => {
  const { document, panel } = openMarkdown(unique('tasks.md'), '- [ ] a\n- [ ] b\n');
  t.after(() => panel.dispose());
  const init = await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  assert.equal(init.readOnly, false);

  // Read at the click (no settings event needed): refused, and the view gets the new state.
  T.config.set('fileStudio.markdown.editing', false);
  const refused = await toggle(panel, 0, true, 1);
  assert.deepEqual([refused.applied, refused.reason], [false, 'Editing is turned off (setting fileStudio.markdown.editing).']);
  assert.equal(document.getText(), '- [ ] a\n- [ ] b\n');
  assert.deepEqual(callsOf('applyEdit'), []);
  const update = lastOf(panel, 'markdownUpdate');
  assert.equal(update.readOnly, true);
  assert.deepEqual(update.editState, {
    editable: false, code: 'disabled', reason: 'Editing is turned off (setting fileStudio.markdown.editing).', kinds: [],
  });
  assert.ok(!panel.posted.some((m) => m.type === 'error'), 'never the error view');

  // Turned on again: the settings listener sends the new state at once; the toggle works.
  panel.posted.length = 0;
  T.config.set('fileStudio.markdown.editing', true);
  T.fireConfigChange(['fileStudio.markdown.editing']);
  const enabled = await waitFor(() => lastOf(panel, 'markdownUpdate'), 3000, 'markdownUpdate');
  assert.deepEqual([enabled.readOnly, enabled.editState], [false, { editable: true, kinds: ['mdText'] }]); // P6: WYSIWYG edits are allowed
  const applied = await toggle(panel, 0, true, 1);
  assert.equal(applied.applied, true);
  assert.equal(document.getText(), '- [x] a\n- [ ] b\n');

  // The master switch turns it off too; another format's switch does not.
  panel.posted.length = 0;
  T.config.set('fileStudio.xlsx.editing', false);
  T.fireConfigChange(['fileStudio.xlsx.editing']);
  await tick(50);
  assert.equal(lastOf(panel, 'markdownUpdate'), undefined, 'not affected by the xlsx switch');
  T.config.set('fileStudio.editing.enabled', false);
  T.fireConfigChange(['fileStudio.editing.enabled']);
  const off = await waitFor(() => lastOf(panel, 'markdownUpdate'), 3000, 'markdownUpdate');
  assert.equal(off.readOnly, true);
  assert.equal(off.editState.reason, 'Editing is turned off (setting fileStudio.editing.enabled).');
});

test('task toggles of all panels of one document run one at a time in the document queue', async (t) => {
  const name = unique('shared.md');
  const a = openMarkdown(name, '- [ ] a\n- [ ] b\n', 1);
  const b = openMarkdown(name, '- [ ] a\n- [ ] b\n', 2);
  t.after(() => {
    a.panel.dispose();
    b.panel.dispose();
  });
  assert.equal(a.document, b.document);
  await waitFor(() => lastOf(a.panel, 'init') && lastOf(b.panel, 'init'), 3000, 'both inits');
  a.panel.posted.length = 0;
  b.panel.posted.length = 0;
  // Both clicked on version 1 at the same time: the second one runs after the first and sees the change.
  a.panel.send({ type: 'toggleTask', line: 0, checked: true, version: 1 });
  b.panel.send({ type: 'toggleTask', line: 1, checked: true, version: 1 });
  const first = await waitFor(() => lastOf(a.panel, 'toggleTaskResult'), 3000, 'first result');
  const second = await waitFor(() => lastOf(b.panel, 'toggleTaskResult'), 3000, 'second result');
  assert.equal(first.applied, true);
  assert.deepEqual([second.applied, second.reason], [false, 'the document changed since the click']);
  assert.equal(a.document.getText(), '- [x] a\n- [ ] b\n');
  assert.equal(callsOf('applyEdit').length, 1);
});

// ===== TEST API =====

test('test API: post() delivers a webview message and returns what was posted back; state() reads the document', async (t) => {
  const { uri, document, panel } = openMarkdown(unique('api.md'), '- [ ] a\n');
  t.after(() => panel.dispose());
  await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  assert.deepEqual(api.state(uri.toString()), { version: 1, dirty: false, editState: { editable: true, kinds: ['mdText'] } }); // P6
  const posted = await api.post(uri.toString(), { type: 'toggleTask', line: 0, checked: true, version: 1 });
  assert.deepEqual(
    posted.filter((m) => m.type === 'toggleTaskResult'),
    [{ type: 'toggleTaskResult', line: 0, version: 1, applied: true, documentVersion: 2 }],
  );
  assert.equal(document.getText(), '- [x] a\n');
  assert.equal(api.state(uri.toString()).version, 2);
  const answer = await api.post(uri.toString(), { type: 'edit', editId: 1, baseVersion: 2, intent: { kind: 'mdText', edits: [] } });
  assert.deepEqual(answer, [{ type: 'editResult', editId: 1, accepted: false, code: 'invalid', reason: EDIT_TEXT.invalid, version: 2 }]);
  await assert.rejects(api.post('file:///nowhere/x.md', { type: 'ready' }), /No FileStudio panel shows/);
  assert.throws(() => api.state('file:///nowhere/x.md'), /no open document/);
});
