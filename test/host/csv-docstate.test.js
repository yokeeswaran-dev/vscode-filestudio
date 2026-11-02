// Host tests for CsvDocState (P3 WP2, contract §8): one state per open CSV / TSV / PSV / SSV document, shared by all its
// panels: pinned delimiter, versions and the echo rule, edits through workspace.applyEdit (exact bytes), patches to the
// other panels, structural edits, whole-file progress, Change Delimiter, Find, the kill switch / read-only gate and the
// save-fidelity check (pending = not editable, "Edit anyway"). The host code runs against test/host/fake-vscode.js.
// Run with `npm test`.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');

const { vscode, T, host, extensionUri, token, tick, writeTmp, lastOf, callsOf, waitFor } = require('./load-host');

beforeEach(() => {
  T.reset();
  delete process.env.FILESTUDIO_TEST_FIDELITY_DELAY_MS;
});

let fileCounter = 0;
let editCounter = 1000;

/** A text document VS Code has synced, shown in a FileStudio csv panel (not ready yet). */
function openCsv(name, text, options = {}) {
  const uri = writeTmp(`csv-${++fileCounter}-${name}`, options.disk ?? text);
  const document = new T.FakeTextDocument(uri, text);
  if (options.crlf) document.eol = 2;
  if (options.encoding) document.encoding = options.encoding;
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'csv').resolveCustomTextEditor(document, panel, token);
  return { uri, document, panel };
}

/** The document's edit state as the host holds it (the test API reads the edit gate entry). */
const stateOf = (opened) => host.createTestApi().state(opened.uri.toString()).editState;

/** Opens a panel and waits for its init and, unless `settle: false`, for the fidelity check to have finished. */
async function openReady(name, text, options = {}) {
  const opened = openCsv(name, text, options);
  opened.panel.send({ type: 'ready' });
  opened.init = await waitFor(() => lastOf(opened.panel, 'init'), 3000, 'init');
  if (options.settle !== false) {
    await waitFor(() => {
      const state = stateOf(opened);
      return state && state.reason !== 'Checking the file…';
    }, 3000, 'the fidelity check');
  }
  return opened;
}

/** A second panel for the same document, ready. */
async function addPanel(opened) {
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'csv').resolveCustomTextEditor(opened.document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 3000, 'second init');
  return { panel, init };
}

/** Sends an `edit` and returns the editResult. */
async function edit(panel, baseVersion, intent, extra = {}) {
  const editId = ++editCounter;
  panel.posted.length = 0;
  panel.send({ type: 'edit', editId, baseVersion, intent, ...extra });
  const result = await waitFor(() => panel.posted.find((m) => m.type === 'editResult' && m.editId === editId), 3000, `editResult ${editId}`);
  return result;
}

/** A change made in a text editor next to the grid (the document becomes dirty, so no fidelity re-check). */
const typeInto = (document, text) => {
  document.isDirty = true;
  document.setText(text);
};
const setCells = (...cells) => ({ kind: 'setCells', sheetId: 'csv', cells });
const rowsOf = async (panel, end = 50) => {
  panel.posted.length = 0;
  panel.send({ type: 'getRows', reqId: 77, sheet: 0, start: 0, end });
  const msg = await waitFor(() => lastOf(panel, 'rows'), 3000, 'rows');
  const out = [];
  for (const row of msg.rows) out[row.r] = row.cells.map((c) => c.v);
  return { msg, rows: out };
};

// ===== OPEN =====

test('init carries the version, the pinned delimiter, id "csv" and an editable state (fidelity ok)', async () => {
  const opened = await openReady('open.csv', 'a,b\n1,2\n');
  const { panel, init } = opened;
  assert.equal(init.kind, 'csv');
  assert.equal(init.version, 1);
  assert.deepEqual(init.delimited, { format: 'csv', delimiter: ',' });
  assert.equal(init.meta.sheets[0].id, 'csv');
  assert.equal(init.readOnly, !init.editState.editable, 'readOnly stays the opposite of editState.editable');
  const state = stateOf(opened);
  assert.equal(state.editable, true);
  assert.deepEqual(state.kinds, ['setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll']);
  const { msg } = await rowsOf(panel);
  assert.equal(msg.version, 1, 'rows carry the host version');
  assert.equal(msg.sheetId, 'csv');
});

test('before the fidelity check has finished the document is not editable and an edit is refused (D-072)', async () => {
  host.createTestApi(); // the delay hook exists only in test mode
  process.env.FILESTUDIO_TEST_FIDELITY_DELAY_MS = '400';
  const opened = await openReady('pending.csv', 'a,b\n1,2\n', { settle: false });
  const { document, panel, init } = opened;
  assert.equal(init.editState.editable, false);
  assert.equal(init.editState.code, 'fidelity');
  assert.equal(init.editState.reason, 'Checking the file…');
  assert.equal(init.readOnly, true);
  const refused = await edit(panel, init.version, setCells([1, 1, 'x']));
  assert.equal(refused.accepted, false);
  assert.equal(refused.code, 'fidelity');
  assert.equal(refused.reason, 'FileStudio is still checking this file; try again in a moment.');
  assert.equal(document.getText(), 'a,b\n1,2\n');
  assert.deepEqual(callsOf('applyEdit'), []);
  await waitFor(() => stateOf(opened)?.editable, 3000, 'editable');
  assert.equal(panel.posted.filter((m) => m.type === 'editState').pop()?.state.editable, true, 'the panel was told');
  const ok = await edit(panel, init.version, setCells([1, 1, 'x']));
  assert.equal(ok.accepted, true);
  assert.equal(document.getText(), 'a,b\n1,x\n');
});

// ===== EDITS =====

test('setCells writes only the changed field; the answer has the new version and the changed rows', async () => {
  const { document, panel, init } = await openReady('set.csv', 'a,b\n1,2\n3,4\n');
  const result = await edit(panel, init.version, setCells([1, 1, 'x'], [2, 0, 'p,q']));
  assert.equal(result.accepted, true);
  assert.equal(result.version, init.version + 1);
  assert.equal(document.getText(), 'a,b\n1,x\n"p,q",4\n');
  assert.equal(callsOf('applyEdit').length, 1, 'one WorkspaceEdit = one undo step');
  assert.deepEqual(result.patches, [{ sheetId: 'csv', start: 1, end: 3, rows: [
    { r: 1, cells: [{ c: 0, t: 'n', v: 1, w: '1' }, { c: 1, t: 's', v: 'x', w: 'x' }] },
    { r: 2, cells: [{ c: 0, t: 's', v: 'p,q', w: 'p,q' }, { c: 1, t: 'n', v: 4, w: '4' }] },
  ], styles: [] }]);
  // The echo of our own change is not an external change: no invalidate follows, the version moved by exactly one.
  await tick(400);
  assert.equal(panel.posted.some((m) => m.type === 'invalidate'), false);
  const { rows, msg } = await rowsOf(panel);
  assert.equal(msg.version, init.version + 1);
  assert.deepEqual(rows[1], [1, 'x']);
  assert.equal(host.createTestApi().state(document.uri.toString()).version, init.version + 1);
});

test('an edit with an old baseVersion is stale: nothing is written and the view is refreshed', async () => {
  const { document, panel, init } = await openReady('stale.csv', 'a,b\n1,2\n');
  typeInto(document, 'a,b\n9,9\n'); // a text editor change (not ours)
  const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(refused.accepted, false);
  assert.equal(refused.code, 'stale');
  assert.equal(refused.version, init.version + 1, 'the answer names the current version');
  assert.equal(document.getText(), 'a,b\n9,9\n');
  assert.deepEqual(callsOf('applyEdit'), []);
  const invalidate = await waitFor(() => lastOf(panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(invalidate.version, init.version + 1);
});

test('a change that is not ours (text editor, undo, reload) bumps the version at once and is re-indexed on demand', async () => {
  const { document, panel, init } = await openReady('ext.csv', 'a,b\n1,2\n');
  panel.posted.length = 0;
  typeInto(document, 'a,b\nx,y\n');
  const api = host.createTestApi();
  assert.equal(api.state(document.uri.toString()).version, init.version + 1, 'the version moved before the 150 ms debounce');
  // An edit planned for the new version works at once (ensureIndexed covers the debounce gap).
  const ok = await edit(panel, init.version + 1, setCells([1, 1, 'z']));
  assert.equal(ok.accepted, true);
  assert.equal(document.getText(), 'a,b\nx,z\n');
});

test('the delimiter is pinned at the first parse: a text change does not re-detect it', async () => {
  const { document, panel, init } = await openReady('pin.csv', 'a;b\n1;2\n');
  assert.deepEqual(init.delimited, { format: 'csv', delimiter: ';' });
  typeInto(document, 'a,b\n1,2\n'); // now commas, but the view keeps ';'
  panel.posted.length = 0;
  const invalidate = await waitFor(() => lastOf(panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(lastOf(panel, 'init'), undefined, 'no new init: the delimiter is the same');
  assert.equal(invalidate.version, init.version + 1);
  const { rows } = await rowsOf(panel);
  assert.deepEqual(rows, [['a,b'], [1.2]], "';' stays the delimiter, so 1,2 is the number 1.2");
});

test('Change Delimiter re-pins: new version, new init, an edit planned before is stale', async () => {
  const { uri, document, panel, init } = await openReady('repin.csv', 'a,b|c\n1,2|3\n');
  panel.posted.length = 0;
  T.quickPick.answer = (items) => items.find((item) => item.delimiter === '|');
  await host.changeDelimiter(uri);
  const second = await waitFor(() => lastOf(panel, 'init'), 3000, 'init after Change Delimiter');
  assert.deepEqual(second.delimited, { format: 'csv', delimiter: '|' });
  assert.equal(second.version, init.version + 1);
  const stale = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(stale.code, 'stale');
  const ok = await edit(panel, second.version, setCells([1, 1, 'x']));
  assert.equal(ok.accepted, true);
  assert.equal(document.getText(), 'a,b|c\n1,2|x\n', 'written with the delimiter on screen');
  T.quickPick.answer = (items) => items[0];
  await host.changeDelimiter(uri); // back to "Detect automatically"
});

test('two panels of one document share the version; the other panel gets a patch', async () => {
  const first = await openReady('two.csv', 'a,b\n1,2\n');
  const second = await addPanel(first);
  assert.equal(second.init.version, first.init.version, 'same version');
  second.panel.posted.length = 0;
  const result = await edit(first.panel, first.init.version, setCells([1, 0, '7']));
  assert.equal(result.accepted, true);
  const patch = await waitFor(() => lastOf(second.panel, 'patch'), 3000, 'patch in the second panel');
  assert.equal(patch.version, result.version);
  assert.equal(patch.patches[0].rows[0].cells[0].v, 7);
});

test('a row appended below the data widens the used range: the sender gets a meta patch', async () => {
  const { document, panel, init } = await openReady('append.csv', 'a,b\n1,2\n');
  const result = await edit(panel, init.version, setCells([3, 0, 'new']));
  assert.equal(result.accepted, true);
  assert.equal(document.getText(), 'a,b\n1,2\n,\nnew,\n');
  const meta = await waitFor(() => lastOf(panel, 'patch'), 3000, 'meta patch');
  assert.equal(meta.sheets[0].rowCount, 4);
});

test('a no-op edit is accepted without touching the document', async () => {
  const { document, panel, init } = await openReady('noop.csv', 'a,b\n1,2\n');
  const result = await edit(panel, init.version, setCells([9, 9, '']));
  assert.equal(result.accepted, true);
  assert.equal(result.version, init.version);
  assert.deepEqual(callsOf('applyEdit'), []);
  assert.equal(document.getText(), 'a,b\n1,2\n');
});

test('then: save saves the document after the edit was accepted', async () => {
  const { uri, panel, init } = await openReady('save.csv', 'a,b\n1,2\n');
  const result = await edit(panel, init.version, setCells([1, 0, '5']), { then: 'save' });
  assert.equal(result.accepted, true);
  await waitFor(() => callsOf('save').length === 1, 3000, 'save');
  assert.equal(fs.readFileSync(uri.fsPath, 'utf8'), 'a,b\n5,2\n');
});

test('CRLF document: new line breaks inside a value are written as CRLF', async () => {
  const { document, panel, init } = await openReady('crlf.csv', 'a,b\r\n1,2\r\n', { crlf: true });
  const result = await edit(panel, init.version, setCells([1, 1, 'x\ny']));
  assert.equal(result.accepted, true);
  assert.equal(document.getText(), 'a,b\r\n1,"x\r\ny"\r\n');
});

// ===== STRUCTURE AND WHOLE-FILE OPERATIONS =====

test('insertRows rebuilds the index: editResult without patches, then an invalidate with the new version', async () => {
  const { document, panel, init } = await openReady('rows.csv', 'a,b\n1,2\n3,4\n');
  const result = await edit(panel, init.version, { kind: 'insertRows', sheetId: 'csv', at: 1, count: 1 });
  assert.equal(result.accepted, true);
  assert.equal(result.patches, undefined);
  assert.equal(document.getText(), 'a,b\n,\n1,2\n3,4\n');
  const invalidate = await waitFor(() => lastOf(panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(invalidate.version, result.version);
  assert.equal(invalidate.meta.sheets[0].rowCount, 4);
  const ok = await edit(panel, result.version, setCells([2, 0, 'x']));
  assert.equal(ok.accepted, true);
  assert.equal(document.getText(), 'a,b\n,\nx,2\n3,4\n');
});

test('sort runs under a progress title; the header does not move', async () => {
  const { document, panel, init } = await openReady('sort.csv', 'k,v\nb,1\na,2\nc,3\n');
  const result = await edit(panel, init.version, { kind: 'sort', sheetId: 'csv', range: { r0: 0, c0: 0, r1: 3, c1: 1 }, key: 0, desc: false, header: false });
  assert.equal(result.accepted, true);
  assert.equal(document.getText(), 'k,v\na,2\nb,1\nc,3\n');
  assert.deepEqual(callsOf('withProgress').map((c) => c[1]), ['Sorting 4 rows…']);
  assert.equal(callsOf('applyEdit').length, 1);
});

test('blocked cases: columns in a space-separated file, a quote that never closes', async () => {
  const ssv = await openReady('cols.ssv', 'a b\n1 2\n');
  const refused = await edit(ssv.panel, ssv.init.version, { kind: 'insertCols', sheetId: 'csv', at: 0, count: 1 });
  assert.equal(refused.accepted, false);
  assert.equal(refused.code, 'blocked');
  assert.equal(refused.reason, 'Columns cannot be inserted or deleted in a space-separated file.');
  assert.equal(ssv.document.getText(), 'a b\n1 2\n');

  const open = await openReady('open-quote.csv', 'a,b\n1,"never closed\n3,4\n');
  const quoteState = stateOf(open);
  assert.equal(quoteState.editable, false);
  assert.equal(quoteState.code, 'blocked');
  assert.equal(quoteState.reason, 'Row 2 opens a quote that is never closed; editing could change the rest of the file.');
  const blocked = await edit(open.panel, open.init.version, setCells([0, 0, 'x']));
  assert.equal(blocked.code, 'blocked');
  assert.equal(open.document.getText(), 'a,b\n1,"never closed\n3,4\n');
});

test('an intent the planner rejects as invalid (bad regex) is answered invalid, nothing changes', async () => {
  const { document, panel, init } = await openReady('badre.csv', 'a,b\n1,2\n');
  const result = await edit(panel, init.version, { kind: 'replaceAll', query: { scope: 'csv', text: '(x+)+', matchCase: false, regex: true, wholeCell: false, lookIn: 'values' }, replace: 'y' });
  assert.equal(result.accepted, false);
  assert.equal(result.code, 'invalid');
  assert.equal(result.reason, 'This pattern could take too long; simplify it.');
  assert.equal(document.getText(), 'a,b\n1,2\n');
});

// ===== FIND =====

test('find: all matches in reading order, next / prev with wrap-around, an invalid pattern answers error', async () => {
  const { panel, init } = await openReady('find.csv', 'name,note\nfoo,bar\nbaz,Foo\nqux,foo foo\n');
  const query = { scope: 'csv', text: 'foo', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' };
  const ask = async (msg) => {
    panel.posted.length = 0;
    panel.send({ type: 'find', reqId: 5, query, direction: 'all', ...msg });
    return waitFor(() => lastOf(panel, 'findResult'), 3000, 'findResult');
  };
  const all = await ask({});
  assert.deepEqual(all.matches.map((m) => [m.r, m.c]), [[1, 0], [2, 1], [3, 1]]);
  assert.equal(all.total, 3);
  assert.equal(all.truncated, false);
  assert.equal(all.version, init.version);
  assert.equal(all.matches[0].sheetId, 'csv');
  const next = await ask({ direction: 'next', from: { sheetId: 'csv', r: 1, c: 0 } });
  assert.deepEqual(next.matches.map((m) => [m.r, m.c]), [[2, 1]]);
  const wrap = await ask({ direction: 'next', from: { sheetId: 'csv', r: 3, c: 1 } });
  assert.deepEqual(wrap.matches.map((m) => [m.r, m.c]), [[1, 0]]);
  const prev = await ask({ direction: 'prev', from: { sheetId: 'csv', r: 1, c: 0 } });
  assert.deepEqual(prev.matches.map((m) => [m.r, m.c]), [[3, 1]]);
  const whole = await ask({ query: { ...query, wholeCell: true, matchCase: true, text: 'foo' } });
  assert.deepEqual(whole.matches.map((m) => [m.r, m.c]), [[1, 0]]);
  const bad = await ask({ query: { ...query, regex: true, text: '(a+)+' } });
  assert.equal(bad.error, 'This pattern could take too long; simplify it.');
  assert.deepEqual(bad.matches, []);
});

test('find is limited to 10,000 matches (truncated) and sees an edit made a moment ago', async () => {
  const lines = ['h'];
  for (let i = 0; i < 12000; i++) lines.push('x');
  const { panel, init } = await openReady('many.csv', `${lines.join('\n')}\n`);
  panel.posted.length = 0;
  panel.send({ type: 'find', reqId: 6, query: { scope: 'csv', text: 'x', matchCase: false, regex: false, wholeCell: true, lookIn: 'values' }, direction: 'all' });
  const found = await waitFor(() => lastOf(panel, 'findResult'), 5000, 'findResult');
  assert.equal(found.matches.length, 10000);
  assert.equal(found.total, 12000);
  assert.equal(found.truncated, true);
  const result = await edit(panel, init.version, setCells([1, 0, 'nope']));
  assert.equal(result.accepted, true);
  panel.posted.length = 0;
  panel.send({ type: 'find', reqId: 7, query: { scope: 'csv', text: 'nope', matchCase: false, regex: false, wholeCell: true, lookIn: 'values' }, direction: 'all' });
  const again = await waitFor(() => lastOf(panel, 'findResult'), 3000, 'findResult');
  assert.deepEqual(again.matches.map((m) => [m.r, m.c]), [[1, 0]]);
  assert.equal(again.version, result.version);
});

// ===== THE GATE: SETTINGS, READ-ONLY, FIDELITY =====

test('the kill switch refuses edits and the panels follow a changed setting', async () => {
  const { document, panel, init } = await openReady('kill.csv', 'a,b\n1,2\n');
  panel.posted.length = 0;
  T.config.set('fileStudio.csv.editing', false);
  T.fireConfigChange(['fileStudio.csv.editing']);
  const off = await waitFor(() => panel.posted.filter((m) => m.type === 'editState').pop(), 3000, 'editState');
  assert.equal(off.state.editable, false);
  assert.equal(off.state.code, 'disabled');
  const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(refused.code, 'disabled');
  assert.equal(refused.reason, 'Editing is turned off (setting fileStudio.csv.editing).');
  assert.equal(document.getText(), 'a,b\n1,2\n');
  T.config.set('fileStudio.csv.editing', true);
  T.fireConfigChange(['fileStudio.csv.editing']);
  await waitFor(() => panel.posted.filter((m) => m.type === 'editState').pop()?.state.editable, 3000, 'editable again');
  const ok = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(ok.accepted, true);
});

test('files.readonlyInclude: the init is read-only and an edit is refused (readOnly)', async () => {
  T.config.set('files.readonlyInclude', { '**/*-ro.csv': true });
  const { document, panel, init } = await openReady('ro.csv', 'a,b\n1,2\n');
  assert.equal(init.readOnly, true);
  assert.equal(init.editState.code, 'readOnly');
  const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(refused.code, 'readOnly');
  assert.equal(refused.reason, 'This file is read-only (files.readonlyInclude).');
  assert.equal(document.getText(), 'a,b\n1,2\n');
});

test('applyEdit answering false is "failed" and nothing changes', async () => {
  const { document, panel, init } = await openReady('failed.csv', 'a,b\n1,2\n');
  const original = vscode.workspace.applyEdit;
  vscode.workspace.applyEdit = async () => false;
  try {
    const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
    assert.equal(refused.code, 'failed');
    assert.equal(refused.reason, 'VS Code did not apply the change. The file was not changed.');
  } finally {
    vscode.workspace.applyEdit = original;
  }
  assert.equal(document.getText(), 'a,b\n1,2\n');
});

test('a file VS Code would save differently is not editable until "Edit anyway" is confirmed', async () => {
  const api = host.createTestApi();
  // The disk has CRLF line breaks, VS Code's document (LF) would write LF back: two line breaks change.
  const opened = await openReady('fidelity.csv', 'a,b\n1,2\n', { disk: 'a,b\r\n1,2\r\n' });
  const { document, panel, init } = opened;
  const state = stateOf(opened);
  assert.equal(state.editable, false);
  assert.equal(state.code, 'fidelity');
  assert.deepEqual(state.canEditAnyway.details, ['VS Code will change 2 line breaks (CRLF → LF).']);
  const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(refused.code, 'fidelity');
  assert.equal(document.getText(), 'a,b\n1,2\n');

  // Cancel (no queued answer): still not editable. Confirm: editable.
  panel.send({ type: 'editAnyway' });
  await tick(100);
  assert.equal(host.createTestApi().state(document.uri.toString()).editState.editable, false);
  api.answer('fidelity', 'Edit anyway');
  panel.posted.length = 0;
  panel.send({ type: 'editAnyway' });
  const editable = await waitFor(() => panel.posted.filter((m) => m.type === 'editState').pop()?.state.editable, 3000, 'editable');
  assert.ok(editable);
  const ok = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(ok.accepted, true);
  assert.equal(document.getText(), 'a,b\nx,2\n');
});

test('wrong encoding (replacement characters that are not in the bytes) is blocked, with no "Edit anyway"', async () => {
  const opened = await openReady('badenc.csv', 'a,b' + String.fromCharCode(10) + String.fromCharCode(0xfffd) + ',2' + String.fromCharCode(10), { disk: Buffer.from([0x61, 0x2c, 0x62, 0x0a, 0xe9, 0x2c, 0x32, 0x0a]) }); // Latin-1 bytes decoded as UTF-8
  const { panel, init } = opened;
  const state = stateOf(opened);
  assert.equal(state.code, 'blocked');
  assert.equal(state.canEditAnyway, undefined);
  const refused = await edit(panel, init.version, setCells([1, 0, 'x']));
  assert.equal(refused.code, 'blocked');
});

test('a change that leaves the document clean (reload, Revert) checks the file again', async () => {
  const { uri, document, panel, init } = await openReady('recheck.csv', 'a,b\n1,2\n');
  panel.posted.length = 0;
  // The file on disk now has CRLF line breaks; VS Code reloads the clean document (its text stays LF-normalised).
  fs.writeFileSync(uri.fsPath, 'a,b\r\n3,4\r\n');
  document.setText('a,b\n3,4\n', false);
  const state = await waitFor(() => {
    const s = panel.posted.filter((m) => m.type === 'editState').pop()?.state;
    return s && s.code === 'fidelity' && s.canEditAnyway ? s : undefined;
  }, 3000, 'fidelity verdict after the reload');
  assert.deepEqual(state.canEditAnyway.details, ['VS Code will change 2 line breaks (CRLF → LF).']);
  const refused = await edit(panel, init.version + 1, setCells([1, 0, 'x']));
  assert.equal(refused.code, 'fidelity');
});

// ===== COMMANDS FROM THE GRID =====

test('command save saves; command undo runs VS Code undo only when this editor is the active tab', async () => {
  const { panel } = await openReady('command.csv', 'a,b\n1,2\n');
  T.calls.length = 0;
  panel.send({ type: 'command', id: 'save' });
  await waitFor(() => callsOf('save').length === 1, 3000, 'save');
  panel.send({ type: 'command', id: 'undo' });
  await waitFor(() => callsOf('executeCommand').some((c) => c[1] === 'undo'), 3000, 'undo');
  panel.active = false;
  panel.send({ type: 'command', id: 'redo' });
  await tick(100);
  assert.equal(callsOf('executeCommand').some((c) => c[1] === 'redo'), false);
});


// ===== FIX ROUND 1 =====

test('a hidden panel does not parse on an external change; it parses when shown again', async () => {
  const { document, panel, init } = await openReady('hidden.csv', 'a,b\n1,2\n');
  panel.posted.length = 0;
  panel.visible = false;
  let reads = 0;
  const getText = document.getText.bind(document);
  document.getText = (...args) => {
    reads++;
    return getText(...args);
  };
  typeInto(document, 'a,b\n3,4\n5,6\n');
  reads = 0;
  await new Promise((r) => setTimeout(r, 400)); // past the 150 ms debounce
  assert.equal(reads, 0, 'nothing read or parsed while every panel is hidden');
  assert.equal(lastOf(panel, 'invalidate'), undefined, 'and nothing sent');
  panel.visible = true;
  panel.fireViewState();
  const invalidate = await waitFor(() => lastOf(panel, 'invalidate'), 3000, 'invalidate after show');
  assert.equal(invalidate.version, init.version + 1);
  assert.ok(reads > 0, 'parsed on show');
  const { rows } = await rowsOf(panel);
  assert.deepEqual(rows, [['a', 'b'], [3, 4], [5, 6]]);
});

test('an empty file opened in the grid detects the delimiter again when text is typed in', async () => {
  const { document, panel, init } = await openReady('empty.csv', '');
  assert.deepEqual(init.delimited, { format: 'csv', delimiter: ',' });
  panel.posted.length = 0;
  typeInto(document, 'a;b;c\n1;2;3\n');
  const again = await waitFor(() => lastOf(panel, 'init'), 3000, 'a new init');
  assert.deepEqual(again.delimited, { format: 'csv', delimiter: ';' });
  const { rows } = await rowsOf(panel);
  assert.deepEqual(rows, [['a', 'b', 'c'], [1, 2, 3]]);
});

test('a file with real columns keeps its pinned delimiter (no re-detect)', async () => {
  const { document, panel } = await openReady('keep.csv', 'a;b\n1;2\n');
  panel.posted.length = 0;
  typeInto(document, 'a,b\n1,2\n');
  await waitFor(() => lastOf(panel, 'invalidate'), 3000, 'invalidate');
  assert.equal(lastOf(panel, 'init'), undefined);
});

test('find: a regex with several repeating parts skips a very long value instead of hanging; next does not wrap after a stopped scan', async () => {
  const { panel } = await openReady('slow.csv', `k\n${'a'.repeat(30000)}c\nab\n`);
  const query = { scope: 'csv', text: 'a*a*a*b', matchCase: false, regex: true, wholeCell: false, lookIn: 'values' };
  panel.posted.length = 0;
  const t0 = Date.now();
  panel.send({ type: 'find', reqId: 9, query, direction: 'all' });
  const res = await waitFor(() => lastOf(panel, 'findResult'), 5000, 'findResult');
  assert.ok(Date.now() - t0 < 3000, 'returns quickly');
  assert.deepEqual(res.matches.map((m) => [m.r, m.c]), [[2, 0]], 'the short value still matches');
  assert.equal(res.truncated, true);
  assert.match(res.error, /too long for this regular expression/);
});

// ===== FIX ROUND 2 =====

test('R2 a character the file encoding cannot hold is refused (it would be saved as ?); edits that fit still work', async () => {
  const opened = openCsv('w1252.csv', 'name,city\nAnna,Zürich\n', { disk: Buffer.from('name,city\nAnna,Z\xfcrich\n', 'latin1'), encoding: 'windows1252' });
  opened.panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(opened.panel, 'init'), 3000, 'init');
  await waitFor(() => stateOf(opened) && stateOf(opened).reason !== 'Checking the file…', 3000, 'the fidelity check');
  assert.equal(stateOf(opened).editable, true, JSON.stringify(stateOf(opened)));
  const refused = await edit(opened.panel, init.version, setCells([1, 0, 'Tokyo 東京 😀']));
  assert.equal(refused.accepted, false);
  assert.equal(refused.code, 'blocked');
  assert.match(refused.reason, /encoding/i);
  assert.equal(opened.document.getText(), 'name,city\nAnna,Zürich\n');
  const ok = await edit(opened.panel, init.version, setCells([1, 0, 'Émile']));
  assert.equal(ok.accepted, true, 'a value the encoding holds is fine');
});
