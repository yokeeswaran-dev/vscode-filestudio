// Pure helpers of the grid cell editor in media/viewer.js (P2): reading pasted TSV, deciding when an edit text gets an
// apostrophe, which key starts an edit, and the overlay of pending cell texts. (The editor itself is covered by
// test/webview/grid-edit.webview.test.js.) media/viewer.js is bundled to test/.out/viewer.js by test/build-tests.js.
const { test } = require('node:test');
const assert = require('node:assert/strict');

globalThis.acquireVsCodeApi = () => ({ postMessage() {}, getState: () => undefined, setState() {} });
globalThis.window = { addEventListener() {} };
globalThis.document = { getElementById: () => null };
globalThis.navigator ??= { platform: '', userAgent: '' };

const V = require('../.out/viewer.js');

test('parseTsv: rows and tabs, the final line break ends the last row', () => {
  assert.deepEqual(V.parseTsv('a\tb\nc\td'), [['a', 'b'], ['c', 'd']]);
  assert.deepEqual(V.parseTsv('a\tb\r\nc\td\r\n'), [['a', 'b'], ['c', 'd']]);
  assert.deepEqual(V.parseTsv('a'), [['a']]);
  assert.deepEqual(V.parseTsv(''), []);
  assert.deepEqual(V.parseTsv('a\t'), [['a', '']], 'a trailing tab is an empty last cell');
  assert.deepEqual(V.parseTsv('\tb'), [['', 'b']]);
  assert.deepEqual(V.parseTsv('a\t\tb\n\nc'), [['a', '', 'b'], [''], ['c']], 'empty fields and an empty line');
});

test('parseTsv: a quoted field keeps tabs, line breaks and quotes; quotes inside a field are text', () => {
  assert.deepEqual(V.parseTsv('"a\tb"\t"line1\r\nline2"\n"say ""hi"""'), [['a\tb', 'line1\nline2'], ['say "hi"']]);
  assert.deepEqual(V.parseTsv('5" pipe\tHe said "no"'), [['5" pipe', 'He said "no"']], 'a quote that does not start a field is plain text');
  assert.deepEqual(V.parseTsv('"unclosed\ttext'), [['unclosed\ttext']], 'an unclosed quote runs to the end');
  assert.deepEqual(V.parseTsv('"a"b\tc'), [['ab', 'c']], 'text after a closing quote is kept');
});

test('textReadsAsValue: only text that would change meaning when typed back gets an apostrophe', () => {
  for (const text of ['007', '12.5', '-3', '(4)', '$5', '1,234', '50%', '1e5', '=A1+B1', '=1+1', "'x", 'TRUE', 'false', '#N/A', '#DIV/0!', '2024-01-31', '3/4/2020', '12:30', '9:05 PM']) {
    assert.equal(V.textReadsAsValue(text), true, text);
  }
  for (const text of ['', 'Short', '3 apples', '1st', 'abc123', 'e5', 'a=b', 'hello world', 'N/A', 'true story', '1-2', '#tag']) {
    assert.equal(V.textReadsAsValue(text), false, text);
  }
});

test('editStartKey: printable characters (incl. AltGr and Space) start an edit; IME and dead keys start an empty one; shortcuts do not', () => {
  const key = (props) => ({ key: '', ctrlKey: false, altKey: false, metaKey: false, shiftKey: false, keyCode: 0, getModifierState: (m) => m === 'AltGraph' && !!props.altGr, ...props });
  assert.equal(V.editStartKey(key({ key: 'a' })), 'a');
  assert.equal(V.editStartKey(key({ key: 'A', shiftKey: true })), 'A');
  assert.equal(V.editStartKey(key({ key: '7' })), '7');
  assert.equal(V.editStartKey(key({ key: ' ' })), ' ');
  assert.equal(V.editStartKey(key({ key: '\u{1F600}' })), '\u{1F600}', 'a character outside the BMP is one key');
  assert.equal(V.editStartKey(key({ key: '@', ctrlKey: true, altKey: true, altGr: true })), '@', 'AltGr = Ctrl+Alt on Windows');
  assert.equal(V.editStartKey(key({ key: 'Process', keyCode: 229 })), '');
  assert.equal(V.editStartKey(key({ key: 'Dead' })), '');
  assert.equal(V.editStartKey(key({ key: 'x', keyCode: 229 })), '');
  for (const p of [{ key: ' ', shiftKey: true }, { key: 'c', ctrlKey: true }, { key: 'c', metaKey: true }, { key: 'f', altKey: true }, { key: 'Enter' }, { key: 'Tab' }, { key: 'F2' }, { key: 'ArrowDown' }, { key: 'Delete' }, { key: 'Shift' }, { key: 'Unidentified' }]) {
    assert.equal(V.editStartKey(key(p)), null, JSON.stringify(p));
  }
});

test('overlay: pending texts by sheet, row and column; a drop only removes the text it was made for; the cached row is never changed', () => {
  V.__setGrid({ kind: 'sheet', overlay: new Map(), dataVersion: 1 });
  assert.equal(V.getCellOverlay(0, 3, 2), undefined);
  V.setCellOverlay(0, 3, 2, 'a');
  V.setCellOverlay(0, 3, 5, 'b');
  V.setCellOverlay(1, 3, 2, 'other sheet');
  assert.equal(V.getCellOverlay(0, 3, 2), 'a');
  assert.equal(V.getCellOverlay(1, 3, 2), 'other sheet');
  V.dropCellOverlay(0, 3, 2, 'stale text');
  assert.equal(V.getCellOverlay(0, 3, 2), 'a', 'a newer pending text of the same cell stays');
  V.dropCellOverlay(0, 3, 2, 'a');
  assert.equal(V.getCellOverlay(0, 3, 2), undefined);
  assert.equal(V.getCellOverlay(0, 3, 5), 'b');

  const cached = [{ c: 1, t: 's', v: 'x' }, { c: 5, t: 's', v: 'old' }];
  const merged = V.withOverlayCells(cached, new Map([[2, 'new'], [5, 'over']]));
  assert.deepEqual(merged.map((c) => [c.c, c.v]), [[1, 'x'], [2, 'new'], [5, 'old']], 'a text cell for every overlay column without a cell, sorted; existing cells stay');
  assert.equal(cached.length, 2, 'the cached row is not changed');
  assert.equal(V.withOverlayCells(cached, new Map([[1, 'z']])), cached, 'nothing to add: the same array');
});
