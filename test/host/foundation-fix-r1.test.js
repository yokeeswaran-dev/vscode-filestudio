// Fix round 1 for F: regression tests (each failed before its fix).
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const { vscode, T, host, extensionUri, writeTmp, tick } = require('./load-host');
const internal = require('../.out/host.js');

beforeEach(() => T.reset());
const { parseEditMessage } = host;

const codec = {
  encode: async (text) => new Uint8Array(Buffer.from(text, 'utf8')),
  decode: async (bytes) => new TextDecoder('utf-8').decode(bytes),
};
let counter = 0;
function fidelityDoc(name, text) {
  const uri = writeTmp(`r1-${++counter}-${name}`, Buffer.from(text, 'utf8'));
  const document = new T.FakeTextDocument(uri, text);
  Object.assign(document, { eol: vscode.EndOfLine.LF, encoding: 'utf8', isDirty: false });
  return document;
}
async function detailsOf(name, text, settings) {
  T.config.clear();
  for (const [k, v] of Object.entries(settings)) T.config.set(`files.${k}`, v);
  vscode.workspace.encode = codec.encode;
  vscode.workspace.decode = codec.decode;
  const r = await host.checkTextFidelity(fidelityDoc(name, text));
  return r.ok ? [] : r.canEditAnyway.details;
}

// ===== save settings (real VS Code 1.142 results) =====

test('trimFinalNewlines: same decisions as VS Code (whitespace-only lines stay, line-break-only file becomes empty)', async () => {
  const s = { trimFinalNewlines: true };
  for (const text of ['\n', '\n\n', '\r\n\r\n']) {
    assert.equal((await detailsOf('a.md', text, s)).length, 1, `warn for ${JSON.stringify(text)}`);
  }
  assert.equal((await detailsOf('a.md', '\n \n\n', s)).length, 1, '\\n \\n\\n loses one line');
  for (const text of ['a\n \n', '\na\n \n', 'a\n', 'a', ' \n']) {
    assert.deepEqual(await detailsOf('a.md', text, s), [], `no warning for ${JSON.stringify(text)}`);
  }
  const two = await detailsOf('a.md', 'a\n \n\n', s);
  assert.match(two[0], /remove 1 empty line /, 'only one line is removed');
});

test('insertFinalNewline: no warning when the last line is only spaces or tabs', async () => {
  const s = { insertFinalNewline: true };
  for (const text of ['a\n ', ' ', 'a\n\t']) assert.deepEqual(await detailsOf('b.md', text, s), [], JSON.stringify(text));
  assert.equal((await detailsOf('b.md', 'a\nb', s)).length, 1);
});

test('trimTrailingWhitespace check is not quadratic on a long run of spaces', async () => {
  const text = ' '.repeat(80_000) + 'x';
  const t0 = Date.now();
  assert.deepEqual(await detailsOf('c.md', text, { trimTrailingWhitespace: true }), []);
  assert.ok(Date.now() - t0 < 1000, `took ${Date.now() - t0} ms`);
  assert.match((await detailsOf('c.md', 'a  \nb\t\nc \r\nd  ', { trimTrailingWhitespace: true }))[0], /end of 4 lines/);
});

// ===== kill switch =====

test('kill switch: any value other than true or unset turns editing off', () => {
  const uri = vscode.Uri.file('/w/k.csv');
  for (const v of ['false', 0, 'off', null, false, 'no']) {
    T.config.clear();
    T.config.set('fileStudio.editing.enabled', v);
    assert.equal(host.editingEnabled('csv', uri), false, `editing.enabled = ${JSON.stringify(v)}`);
    T.config.clear();
    T.config.set('fileStudio.csv.editing', v);
    assert.equal(host.editingEnabled('csv', uri), false, `csv.editing = ${JSON.stringify(v)}`);
  }
  T.config.clear();
  assert.equal(host.editingEnabled('csv', uri), true, 'unset (the package default is true)');
  T.config.set('fileStudio.editing.enabled', true);
  T.config.set('fileStudio.csv.editing', true);
  assert.equal(host.editingEnabled('csv', uri), true);
});

// ===== toggleTask parser =====

test('toggleTask: a negative, fractional or too large line or version is refused, not clamped', () => {
  const ok = { type: 'toggleTask', line: 3, checked: true, version: 2 };
  assert.deepEqual(host.parseWebviewMessage(ok), ok);
  for (const bad of [{ line: -1 }, { line: 0.5 }, { line: 1e12 }, { version: -1 }, { version: 1.5 }, { line: NaN }]) {
    assert.equal(host.parseWebviewMessage({ ...ok, ...bad }), undefined, JSON.stringify(bad));
  }
});

// ===== XML-invalid characters =====

test('sheet names and number formats refuse XML-invalid characters', () => {
  const edit = (intent) => ({ type: 'edit', editId: 1, baseVersion: 0, intent });
  const range = { r0: 0, c0: 0, r1: 0, c1: 0 };
  for (const bad of ['a￾b', 'a￿']) {
    assert.equal(parseEditMessage(edit({ kind: 'sheetRename', sheetId: 's1', name: bad }), 'sheet').ok, false, 'name');
  }
  for (const bad of ['0\u0001', '0\u0000', 'a￾', '0\n0', '0￿']) {
    assert.equal(parseEditMessage(edit({ kind: 'style', sheetId: 's1', ranges: [range], delta: { numFmt: bad } }), 'sheet').ok, false, JSON.stringify(bad));
  }
  // Markup characters are valid text (the patcher escapes every attribute): accepted unchanged.
  const fmt = '0"/><numFmt numFmtId="9" formatCode="x';
  const r = parseEditMessage(edit({ kind: 'style', sheetId: 's1', ranges: [range], delta: { numFmt: fmt } }), 'sheet');
  assert.equal(r.ok, true);
  assert.equal(r.msg.intent.delta.numFmt, fmt);
});

// ===== sessions: editId reuse =====

const intent = { kind: 'setCells', sheetId: 'csv', cells: [[0, 0, 'x']] };
const results = (panel) => panel.posted.filter((m) => m.type === 'editResult');

test('a second edit with an editId that is still open is refused and does not take the first edit\'s answer', async () => {
  const panel = T.createPanel();
  const session = new internal.ViewerSession(panel, vscode.Uri.file('/w/d1.csv'), 'csv', () => undefined);
  let release;
  const gate = new Promise((r) => (release = r));
  const handled = [];
  session.editHandlers = {
    version: () => 1,
    onEdit: async (msg) => {
      handled.push(msg.intent.cells[0][2]);
      await gate;
      session.post({ type: 'editResult', editId: msg.editId, accepted: true, version: 2 });
    },
  };
  panel.send({ type: 'ready' });
  panel.send({ type: 'edit', editId: 7, baseVersion: 1, intent });
  await tick(5);
  panel.send({ type: 'edit', editId: 7, baseVersion: 1, intent: { ...intent, cells: 'bad' } });
  await tick(5);
  release();
  await tick(20);
  assert.deepEqual(handled, ['x']);
  const r = results(panel);
  assert.equal(r.length, 1, JSON.stringify(r));
  assert.equal(r[0].accepted, true, 'the first edit keeps its own answer');
  session.dispose();
});

test('after the webview reloads, an old edit\'s answer is not given to the new edit with the same editId', async () => {
  const panel = T.createPanel();
  const session = new internal.ViewerSession(panel, vscode.Uri.file('/w/d2.csv'), 'csv', () => undefined);
  const queue = new host.SerialQueue();
  let release;
  const gate = new Promise((r) => (release = r));
  let call = 0;
  session.editHandlers = {
    version: () => 7,
    onEdit: (msg) =>
      queue.run(async () => {
        const mine = ++call;
        if (mine === 1) await gate;
        session.post({ type: 'editResult', editId: msg.editId, accepted: mine === 2, version: 8, ...(mine === 1 ? { code: 'stale', reason: 'old' } : {}) });
      }),
  };
  panel.send({ type: 'ready' });
  panel.send({ type: 'edit', editId: 1, baseVersion: 7, intent });
  await tick(5);
  panel.send({ type: 'ready' }); // reload: editIds restart
  panel.send({ type: 'edit', editId: 1, baseVersion: 7, intent });
  await tick(5);
  release();
  await tick(30);
  const r = results(panel);
  assert.equal(call, 2, 'both edits were handled');
  assert.equal(r.length, 1, JSON.stringify(r));
  assert.equal(r[0].accepted, true, 'the answer the new webview gets is the answer of its own edit');
  session.dispose();
});

test('a handler that returns its queue job promise is answered once, even when it is slow', async () => {
  const panel = T.createPanel();
  const session = new internal.ViewerSession(panel, vscode.Uri.file('/w/d3.csv'), 'csv', () => undefined);
  const queue = new host.SerialQueue();
  session.editHandlers = {
    onEdit: (msg) =>
      queue.run(async () => {
        await tick(20);
        session.post({ type: 'editResult', editId: msg.editId, accepted: true, version: 2 });
      }),
  };
  panel.send({ type: 'edit', editId: 5, baseVersion: 1, intent });
  await tick(60);
  const r = results(panel);
  assert.equal(r.length, 1);
  assert.equal(r[0].accepted, true);
  session.dispose();
});

// ===== activate with a fake vscode that has no ExtensionMode =====

test('activate() does not throw when vscode.ExtensionMode is missing (older fake)', () => {
  const saved = vscode.ExtensionMode;
  delete vscode.ExtensionMode;
  try {
    host.extension.activate({ extensionUri, subscriptions: [] });
  } finally {
    vscode.ExtensionMode = saved;
  }
});
