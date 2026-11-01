// Fix round 3 for F: regression tests (each failed before its fix).
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const { T, host } = require('./load-host');

beforeEach(() => T.reset());
const { parseEditMessage } = host;
const edit = (intent) => ({ type: 'edit', editId: 7, baseVersion: 3, intent });
const q = (text, regex, scope = 's1') => ({ scope, text, matchCase: false, regex, wholeCell: false, lookIn: 'values' });

test('replaceAll accepts regex patterns that are valid without the u flag', () => {
  for (const p of ['\d{3}\-\d{4}', 'user\@x', 'a\_b', '\#\d', '\:', 'x{', 'a]', '[a-z]{2,}', '\p{L}+', '\w+@\w+']) {
    const r = parseEditMessage(edit({ kind: 'replaceAll', query: q(p, true), replace: 'b' }), 'sheet');
    assert.equal(r.ok, true, p);
  }
  for (const p of ['(', '[a-', '*a', '\\']) {
    assert.equal(parseEditMessage(edit({ kind: 'replaceAll', query: q(p, true), replace: 'b' }), 'sheet').ok, false, p);
  }
});

test('sheetId and scope are tied to the view kind', () => {
  const set = (sheetId) => edit({ kind: 'setCells', sheetId, cells: [[0, 0, 'a']] });
  assert.equal(parseEditMessage(set('s1'), 'csv').ok, false);
  assert.equal(parseEditMessage(set('csv'), 'sheet').ok, false);
  assert.equal(parseEditMessage(set('csv'), 'csv').ok, true);
  assert.equal(parseEditMessage(set('s1'), 'sheet').ok, true);
  assert.equal(parseEditMessage(edit({ kind: 'replaceAll', query: q('a', false, 's1'), replace: 'b' }), 'csv').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'replaceAll', query: q('a', false, 'csv'), replace: 'b' }), 'sheet').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'replaceAll', query: q('a', false, 'workbook'), replace: 'b' }), 'csv').ok, true);
  assert.equal(parseEditMessage(edit({ kind: 'replaceAll', query: q('a', false, 'csv'), replace: 'b' }), 'csv').ok, true);
});

test('a lone surrogate is refused for csv and markdown text too', () => {
  assert.equal(parseEditMessage(edit({ kind: 'setCells', sheetId: 'csv', cells: [[0, 0, '\ud800']] }), 'csv').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'mdText', edits: [{ start: 0, end: 0, oldText: '', text: '\ud800' }] }), 'markdown').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'setCells', sheetId: 'csv', cells: [[0, 0, '😀']] }), 'csv').ok, true);
});
