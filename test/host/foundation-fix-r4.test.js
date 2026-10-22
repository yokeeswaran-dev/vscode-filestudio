// Fix round 4 for F: regression tests (each failed before its fix).
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const { T, host } = require('./load-host');

beforeEach(() => T.reset());
const { parseEditMessage } = host;
const edit = (intent) => ({ type: 'edit', editId: 7, baseVersion: 3, intent });
const range = { r0: 0, c0: 0, r1: 0, c1: 0 };

test('sparse arrays (holes) are refused for ranges and resize.index', () => {
  const ranges = new Array(3);
  ranges[1] = range;
  assert.equal(parseEditMessage(edit({ kind: 'clear', sheetId: 's1', ranges, what: 'all' }), 'sheet').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'style', sheetId: 's1', ranges, delta: {} }), 'sheet').ok, false);
  const index = new Array(2);
  index[1] = 5;
  assert.equal(parseEditMessage(edit({ kind: 'resize', sheetId: 's1', axis: 'row', index, size: 20 }), 'sheet').ok, false);
  // a dense array still works
  assert.equal(parseEditMessage(edit({ kind: 'resize', sheetId: 's1', axis: 'row', index: [1, 5], size: 20 }), 'sheet').ok, true);
});

test('insert/delete at is bounded by the Excel grid for xlsx (append position allowed), not for csv', () => {
  const ins = (kind, at, count = 1) => edit({ kind, sheetId: 's1', at, count });
  assert.equal(parseEditMessage(ins('insertRows', 2000000000, 1), 'sheet').ok, false);
  assert.equal(parseEditMessage(ins('deleteCols', 1e9, 5), 'sheet').ok, false);
  assert.equal(parseEditMessage(ins('insertCols', 16385), 'sheet').ok, false);
  assert.equal(parseEditMessage(ins('insertCols', 16384), 'sheet').ok, true, 'append position');
  assert.equal(parseEditMessage(ins('insertRows', 1048576), 'sheet').ok, true, 'append position');
  assert.equal(parseEditMessage(ins('insertRows', 1048577), 'sheet').ok, false);
  assert.equal(parseEditMessage(edit({ kind: 'insertRows', sheetId: 'csv', at: 2000000, count: 1 }), 'csv').ok, true);
});

test('lone surrogate refusal does not mention Excel for csv and markdown', () => {
  const a = parseEditMessage(edit({ kind: 'setCells', sheetId: 'csv', cells: [[0, 0, 'a\uD800']] }), 'csv');
  const b = parseEditMessage(edit({ kind: 'mdText', edits: [{ start: 0, end: 0, oldText: '', text: '\uDC00' }] }), 'markdown');
  for (const r of [a, b]) {
    assert.equal(r.ok, false);
    assert.doesNotMatch(r.reason, /Excel/);
  }
});
