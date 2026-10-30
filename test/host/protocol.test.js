// Track F (contract §2.7, §14 F WP7): the editing protocol. parseWebviewMessage / parseEditMessage accept every new
// message in its exact shape, refuse every limit with `invalid` (an `edit` is answered, never dropped), refuse unknown
// kinds, and leave all 0.1.2 messages unchanged. ViewerSession routes the new messages to optional view handlers; a
// view without one answers an `edit` with `blocked` "This action is not available yet.". xlsx and CSV stay view-only.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');

const { T, host, extensionUri, token, writeTmp, lastOf, waitFor, tick } = require('./load-host');

beforeEach(() => T.reset());

const { parseWebviewMessage: parse, parseEditMessage, EDIT_LIMITS, EDIT_TEXT } = host;

const MI = 1024 * 1024;
const edit = (intent, extra = {}) => ({ type: 'edit', editId: 7, baseVersion: 3, intent, ...extra });

/** `raw` is refused: parseEditMessage says why (with the editId, so the session answers it) and parse drops it. */
function refused(raw, kind, reason) {
  const result = parseEditMessage(raw, kind);
  assert.equal(result.ok, false, `accepted: ${JSON.stringify(raw).slice(0, 160)}`);
  assert.equal(result.editId, typeof raw.editId === 'number' && Number.isInteger(raw.editId) && raw.editId >= 0 ? raw.editId : undefined);
  if (reason !== undefined) assert.equal(result.reason, reason);
  assert.equal(parse(raw, kind), undefined);
}

/** `raw` is accepted unchanged (parse returns a clean copy equal to it). */
function accepted(raw, kind) {
  const result = parseEditMessage(raw, kind);
  assert.equal(result.ok, true, `refused (${result.reason}): ${JSON.stringify(raw).slice(0, 160)}`);
  assert.deepEqual(result.msg, raw);
  assert.deepEqual(parse(raw, kind), raw);
}

// ===== 0.1.2 MESSAGES =====

test('all 0.1.2 webview messages parse exactly as before (clamping, limits, unknown types)', () => {
  for (const type of ['ready', 'reopenAsText', 'revealInExplorer']) assert.deepEqual(parse({ type, extra: 1 }), { type });
  assert.deepEqual(parse({ type: 'getRows', reqId: 1.7, sheet: -2, start: 5, end: 99_999, c0: 9, c1: 3 }), {
    type: 'getRows', reqId: 1, sheet: 0, start: 5, end: 10_005, c0: 3, c1: 9,
  });
  assert.deepEqual(parse({ type: 'getRows', reqId: 1, sheet: 0, start: 4, end: 2, c0: 1 }), { type: 'getRows', reqId: 1, sheet: 0, start: 4, end: 4 });
  assert.equal(parse({ type: 'getRows', reqId: 'x', sheet: 0, start: 0 }), undefined);
  const ranges = Array.from({ length: 1_100 }, () => ({ r0: 5, c0: 4, r1: 1, c1: 2 }));
  const stats = parse({ type: 'getStats', reqId: 2, sheet: 1, ranges: [...ranges, { r0: 'a' }] });
  assert.equal(stats.ranges.length, 1_024);
  assert.deepEqual(stats.ranges[0], { r0: 1, c0: 2, r1: 5, c1: 4 });
  assert.deepEqual(parse({ type: 'openLink', href: 'https://x' }), { type: 'openLink', href: 'https://x' });
  assert.equal(parse({ type: 'openLink', href: '' }), undefined);
  assert.equal(parse({ type: 'openLink', href: 'x'.repeat(8_193) }), undefined);
  assert.deepEqual(parse({ type: 'toggleTask', line: 2, checked: true, version: 4 }), { type: 'toggleTask', line: 2, checked: true, version: 4 });
  assert.equal(parse({ type: 'toggleTask', line: 2, checked: 'yes', version: 4 }), undefined);
  assert.deepEqual(parse({ type: 'getSlide', reqId: 3, index: -1 }), { type: 'getSlide', reqId: 3, index: 0 });
  assert.deepEqual(parse({ type: 'log', level: 'bad', message: 'm'.repeat(20_000) }), { type: 'log', level: 'info', message: 'm'.repeat(10_000) });
  assert.equal(parse({ type: 'nope' }), undefined);
  assert.equal(parse(null), undefined);
  assert.equal(parse(['ready']), undefined);
});

// ===== NEW READ / MODE MESSAGES =====

test('new non-edit messages: valid shapes accepted, bad shapes dropped', () => {
  const query = { scope: 's1', text: 'a(', matchCase: false, regex: true, wholeCell: false, lookIn: 'values' };
  const ok = [
    { type: 'getRows', reqId: 1, sheet: 0, start: 0, end: 5, sheetId: 's3' },
    { type: 'getStats', reqId: 1, sheet: 0, ranges: [], sheetId: 'csv' },
    { type: 'editAnyway' },
    // A find with a bad pattern is passed on: its handler answers findResult.error (a read is never left unanswered).
    { type: 'find', reqId: 4, query, direction: 'next', from: { sheetId: 's1', r: 2, c: 3 } },
    { type: 'find', reqId: 4, query: { ...query, scope: 'workbook', regex: false }, direction: 'all' },
    { type: 'filter', reqId: 5, sheetId: 's1', col: 2 },
    { type: 'filter', reqId: 5, sheetId: 's1', col: 2, keep: null },
    { type: 'filter', reqId: 5, sheetId: 's1', col: 2, keep: ['a', ''] },
    { type: 'measure', reqId: 6, sheetId: 's2', axis: 'col', index: 4 },
    { type: 'command', id: 'undo' },
    { type: 'command', id: 'save' },
    { type: 'setMode', mode: 'wysiwyg' },
    { type: 'previewScrolled', line: 12.5 },
    { type: 'flushed', id: 9 },
  ];
  for (const msg of ok) assert.deepEqual(parse(msg), msg, JSON.stringify(msg));
  const bad = [
    { type: 'getRows', reqId: 1, sheet: 0, start: 0, sheetId: 'Sheet1' },
    { type: 'getStats', reqId: 1, sheet: 0, ranges: [], sheetId: 3 },
    { type: 'find', reqId: 4, direction: 'next' },
    { type: 'find', reqId: 4, query, direction: 'up' },
    { type: 'find', reqId: 4, query: { ...query, text: 'x'.repeat(1_001) }, direction: 'next' },
    { type: 'find', reqId: 4, query, direction: 'next', from: { sheetId: 's1', r: -1, c: 0 } },
    { type: 'filter', reqId: 5, sheetId: 's1', col: 2, keep: [1] },
    { type: 'filter', reqId: 5, sheetId: 's1', col: 1.5 },
    { type: 'measure', reqId: 6, sheetId: 's2', axis: 'z', index: 4 },
    { type: 'command', id: 'format' },
    { type: 'setMode', mode: 'edit' },
    { type: 'previewScrolled', line: -1 },
    { type: 'previewScrolled', line: Number.NaN },
    { type: 'flushed', id: -1 },
  ];
  for (const msg of bad) assert.equal(parse(msg), undefined, JSON.stringify(msg));
});

// ===== EDIT: VALID SHAPES =====

test('every edit kind is accepted in its exact shape (contract §2.2)', () => {
  const query = { scope: 'workbook', text: 'a+', matchCase: true, regex: true, wholeCell: false, lookIn: 'formulas' };
  const intents = [
    { kind: 'setCells', sheetId: 's1', cells: [[0, 0, 'x'], [1_048_575, 16_383, '=A1']] },
    { kind: 'clear', sheetId: 's1', ranges: [{ r0: 0, c0: 0, r1: 2, c1: 2 }], what: 'all' },
    { kind: 'insertRows', sheetId: 's1', at: 3, count: 2 },
    { kind: 'deleteRows', sheetId: 's1', at: 3, count: 1 },
    { kind: 'insertCols', sheetId: 's1', at: 0, count: 1_048_576 },
    { kind: 'deleteCols', sheetId: 's1', at: 1, count: 1 },
    { kind: 'sort', sheetId: 's1', range: { r0: 0, c0: 0, r1: 9, c1: 3 }, key: 2, desc: true, header: false },
    { kind: 'replaceAll', query, replace: 'b' },
    {
      kind: 'style',
      sheetId: 's1',
      ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }],
      delta: {
        bold: true, italic: false, strike: true, wrap: true, quotePrefix: false, underline: null, fontColor: 'FF112233', fill: null,
        hAlign: 'justify', vAlign: null, numFmt: '0.00%', border: { top: { style: 'thin', color: 'FF000000' }, left: null },
      },
    },
    { kind: 'merge', sheetId: 's1', range: { r0: 0, c0: 0, r1: 1, c1: 1 } },
    { kind: 'unmerge', sheetId: 's1', range: { r0: 0, c0: 0, r1: 1, c1: 1 } },
    { kind: 'resize', sheetId: 's1', axis: 'row', index: [0, 5], size: 24 },
    { kind: 'resize', sheetId: 's1', axis: 'col', index: [3], size: null },
    { kind: 'sheetAdd', index: 2 },
    { kind: 'sheetAdd', index: 0, name: 'New sheet' },
    { kind: 'sheetRename', sheetId: 's2', name: 'Ünïcode 名前' },
    { kind: 'sheetDelete', sheetId: 's2' },
    { kind: 'sheetDuplicate', sheetId: 's2' },
    { kind: 'sheetMove', sheetId: 's2', index: 0 },
    { kind: 'sheetHide', sheetId: 's2', hidden: true },
    { kind: 'tabColor', sheetId: 's2', argb: 'ff00aa00' },
    { kind: 'tabColor', sheetId: 's2', argb: null },
  ];
  for (const intent of intents) accepted(edit(intent), 'sheet');
  accepted(edit({ kind: 'setCells', sheetId: 'csv', cells: [[5_000_000, 0, 'x']] }), 'csv');
  accepted(edit({ kind: 'clear', sheetId: 'csv', ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], what: 'contents' }), 'csv');
  accepted(edit({ kind: 'mdText', edits: [{ start: 0, end: 2, oldText: 'ab', text: 'x' }, { start: 2, end: 2, oldText: '', text: 'y' }] }), 'markdown');
  accepted(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 0, 'x']] }, { then: 'save' }), 'sheet');
  // Unknown properties are not passed on.
  const parsed = parse(edit({ kind: 'sheetDelete', sheetId: 's2', extra: 1 }, { extra: 2 }), 'sheet');
  assert.deepEqual(parsed, edit({ kind: 'sheetDelete', sheetId: 's2' }));
  // Ranges are normalized like the 0.1.2 getStats ranges.
  assert.deepEqual(parse(edit({ kind: 'merge', sheetId: 's1', range: { r0: 3, c0: 4, r1: 1, c1: 2 } }), 'sheet').intent.range, { r0: 1, c0: 2, r1: 3, c1: 4 });
});

// ===== EDIT: EVERY LIMIT IS `invalid` =====

test('every §2.7 limit makes an edit invalid (answered with the reason, never clamped)', () => {
  const cells = (n, text = 'x') => ({ kind: 'setCells', sheetId: 's1', cells: Array.from({ length: n }, (_, i) => [i, 0, text]) });
  const range = { r0: 0, c0: 0, r1: 0, c1: 0 };
  // setCells: > 1,000,000 cells (checked before the cells are read), empty, malformed cells.
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: new Array(EDIT_LIMITS.cells + 1) }), 'sheet', EDIT_TEXT.tooLarge('1,000,000 cells'));
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [] }), 'sheet');
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 0]] }), 'sheet');
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 0, 5]] }), 'sheet');
  // One value: 32,767 characters for xlsx, 1 Mi for csv.
  accepted(edit(cells(1, 'x'.repeat(32_767))), 'sheet');
  refused(edit(cells(1, 'x'.repeat(32_768))), 'sheet', EDIT_TEXT.tooLarge('32,767 characters in one cell'));
  accepted(edit({ ...cells(1, 'x'.repeat(32_768)), sheetId: 'csv' }), 'csv');
  refused(edit({ ...cells(1, 'x'.repeat(MI + 1)), sheetId: 'csv' }), 'csv', EDIT_TEXT.tooLarge('1,048,576 characters in one cell'));
  // All text together: 16 Mi characters.
  refused(edit({ ...cells(17, 'x'.repeat(MI)), sheetId: 'csv' }), 'csv', EDIT_TEXT.tooLarge('16 million characters of text'));
  // Positions: never clamped or truncated; xlsx is bounded by Excel's grid.
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[-1, 0, 'x']] }), 'sheet');
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[1.5, 0, 'x']] }), 'sheet');
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[1_048_576, 0, 'x']] }), 'sheet');
  refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 16_384, 'x']] }), 'sheet');
  refused(edit({ kind: 'merge', sheetId: 's1', range: { ...range, r1: -3 } }), 'sheet');
  // ranges: 1..1,000.
  refused(edit({ kind: 'clear', sheetId: 's1', ranges: new Array(1_001).fill(range), what: 'contents' }), 'sheet', EDIT_TEXT.tooLarge('1,000 ranges'));
  refused(edit({ kind: 'clear', sheetId: 's1', ranges: [], what: 'contents' }), 'sheet');
  refused(edit({ kind: 'clear', sheetId: 'csv', ranges: [range], what: 'formats' }), 'csv'); // a CSV cell has no format
  refused(edit({ kind: 'clear', sheetId: 's1', ranges: [range], what: 'values' }), 'sheet');
  // count: 1..1,048,576.
  refused(edit({ kind: 'insertRows', sheetId: 's1', at: 0, count: 1_048_577 }), 'sheet');
  refused(edit({ kind: 'deleteCols', sheetId: 's1', at: 0, count: 0 }), 'sheet');
  // Find / replace text: 1,000 characters; a regex must compile.
  const query = { scope: 's1', text: 'a', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' };
  refused(edit({ kind: 'replaceAll', query: { ...query, text: 'x'.repeat(1_001) }, replace: '' }), 'sheet', EDIT_TEXT.tooLarge('1,000 characters to find'));
  refused(edit({ kind: 'replaceAll', query, replace: 'x'.repeat(1_001) }), 'sheet', EDIT_TEXT.tooLarge('1,000 characters to replace with'));
  refused(edit({ kind: 'replaceAll', query: { ...query, regex: true, text: 'a(' }, replace: '' }), 'sheet', EDIT_TEXT.badPattern);
  refused(edit({ kind: 'replaceAll', query: { ...query, lookIn: 'comments' }, replace: '' }), 'sheet');
  // mdText: ≤ 256 edits, sorted, not overlapping, oldText as long as its range, ≤ 4 Mi characters.
  const md = (edits) => edit({ kind: 'mdText', edits });
  const one = (start, end, oldText, text = '') => ({ start, end, oldText, text });
  refused(md(Array.from({ length: 257 }, (_, i) => one(i, i, ''))), 'markdown', EDIT_TEXT.tooLarge('256 text changes at once'));
  refused(md([]), 'markdown');
  refused(md([one(5, 6, 'a'), one(0, 1, 'b')]), 'markdown'); // not sorted
  refused(md([one(0, 3, 'abc'), one(2, 4, 'cd')]), 'markdown'); // overlapping
  refused(md([one(2, 2, '', 'x'), one(2, 2, '', 'y')]), 'markdown'); // two insertions at one offset
  refused(md([one(0, 3, 'ab')]), 'markdown'); // oldText length
  refused(md([one(4, 2, '')]), 'markdown');
  refused(md([one(0, 0, '', 'x'.repeat(4 * MI + 1))]), 'markdown', EDIT_TEXT.tooLarge('4 million characters'));
  // Sheet names: 1..31 characters, no control characters.
  refused(edit({ kind: 'sheetRename', sheetId: 's1', name: 'x'.repeat(32) }), 'sheet', EDIT_TEXT.tooLarge('31 characters in a sheet name'));
  refused(edit({ kind: 'sheetRename', sheetId: 's1', name: '' }), 'sheet', EDIT_TEXT.badSheetName);
  refused(edit({ kind: 'sheetRename', sheetId: 's1', name: 'a\tb' }), 'sheet', EDIT_TEXT.badSheetName);
  refused(edit({ kind: 'sheetAdd', index: 0, name: 'a\u0085b' }), 'sheet', EDIT_TEXT.badSheetName);
  // argb: 8 hex digits.
  refused(edit({ kind: 'tabColor', sheetId: 's1', argb: '#ff0000' }), 'sheet');
  refused(edit({ kind: 'style', sheetId: 's1', ranges: [range], delta: { fill: 'FF00000' } }), 'sheet');
  refused(edit({ kind: 'style', sheetId: 's1', ranges: [range], delta: { border: { top: { style: 'wavy' } } } }), 'sheet');
  refused(edit({ kind: 'style', sheetId: 's1', ranges: [range], delta: { numFmt: '0'.repeat(256) } }), 'sheet');
  refused(edit({ kind: 'resize', sheetId: 's1', axis: 'row', index: [0], size: -1 }), 'sheet');
  refused(edit({ kind: 'resize', sheetId: 's1', axis: 'row', index: [], size: 10 }), 'sheet');
  // Sheet ids: 's' + digits or 'csv'.
  refused(edit({ kind: 'sheetDelete', sheetId: 'Sheet1' }), 'sheet');
  // Unknown kind, bad envelope.
  refused(edit({ kind: 'formatPainter', sheetId: 's1' }), 'sheet', EDIT_TEXT.invalid);
  refused(edit(null), 'sheet');
  refused(edit({ kind: 'sheetDelete', sheetId: 's1' }, { then: 'close' }), 'sheet');
  refused({ type: 'edit', editId: 7, intent: { kind: 'sheetDelete', sheetId: 's1' } }, 'sheet'); // no baseVersion
  refused({ type: 'edit', editId: -1, baseVersion: 0, intent: { kind: 'sheetDelete', sheetId: 's1' } }, 'sheet');
});

test('xlsx text with a lone surrogate is invalid (UTF-8 would store U+FFFD)', () => {
  const lone = ['a\uD800b', 'end\uD83D', '\uDC00start'];
  for (const text of lone) {
    refused(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 0, text]] }), 'sheet', EDIT_TEXT.loneSurrogate);
    refused(edit({ kind: 'sheetRename', sheetId: 's1', name: text }), 'sheet', EDIT_TEXT.loneSurrogate);
    const query = { scope: 's1', text: 'a', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' };
    refused(edit({ kind: 'replaceAll', query, replace: text }), 'sheet', EDIT_TEXT.loneSurrogate);
    refused(edit({ kind: 'setCells', sheetId: 'csv', cells: [[0, 0, text]] }), 'csv', EDIT_TEXT.loneSurrogate); // round 3: every kind
  }
  accepted(edit({ kind: 'setCells', sheetId: 's1', cells: [[0, 0, 'emoji 😀 ok']] }), 'sheet');
});

// ===== SESSION ROUTING =====

async function xlsxFile(name) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = 'a';
  return writeTmp(name, Buffer.from(await workbook.xlsx.writeBuffer()));
}

async function openSheet(name) {
  const uri = await xlsxFile(name);
  const provider = new host.SheetEditorProvider(extensionUri);
  const document = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'sheet init');
  return { uri, document, panel, init };
}

function openText(name, text, kind) {
  const uri = writeTmp(name, text);
  const document = new T.FakeTextDocument(uri, text);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, kind).resolveCustomTextEditor(document, panel, token);
  return { uri, document, panel };
}

test('xlsx cell editing is on (P2); an intent that is not implemented yet is answered blocked, a malformed one invalid; CSV stays view-only', async () => {
  const sheet = await openSheet('view-only.xlsx');
  assert.equal(sheet.init.readOnly, false);
  assert.deepEqual(sheet.init.editState, { editable: true, kinds: ['setCells', 'clear'] });
  assert.equal(sheet.init.version, 0);
  sheet.panel.posted.length = 0;
  sheet.panel.send(edit({ kind: 'clear', sheetId: 's1', ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], what: 'formats' }));
  const blocked = await waitFor(() => lastOf(sheet.panel, 'editResult'), 3000, 'editResult');
  assert.deepEqual(blocked, { type: 'editResult', editId: 7, accepted: false, code: 'blocked', reason: EDIT_TEXT.notAvailable, version: 0, where: [] });
  sheet.panel.posted.length = 0;
  sheet.panel.send({ type: 'edit', editId: 8, baseVersion: 0, intent: { kind: 'setCells', sheetId: 's1', cells: [[-1, 0, 'b']] } });
  const bad = await waitFor(() => lastOf(sheet.panel, 'editResult'), 3000, 'editResult');
  assert.deepEqual(bad, { type: 'editResult', editId: 8, accepted: false, code: 'invalid', reason: EDIT_TEXT.invalid, version: 0 });
  // Reads without a handler are logged and dropped (no answer, no error view).
  sheet.panel.posted.length = 0;
  sheet.panel.send({ type: 'find', reqId: 1, query: { scope: 's1', text: 'a', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' }, direction: 'all' });
  sheet.panel.send({ type: 'command', id: 'undo' });
  await tick(50);
  assert.deepEqual(sheet.panel.posted, []);

  const csv = openText('view-only.csv', 'a,b\n1,2\n', 'csv');
  csv.panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(csv.panel, 'init'), 3000, 'csv init');
  assert.equal(init.readOnly, true);
  assert.deepEqual(init.editState, { editable: false, kinds: [] });
  csv.panel.send(edit({ kind: 'setCells', sheetId: 'csv', cells: [[0, 0, 'b']] }));
  const csvResult = await waitFor(() => lastOf(csv.panel, 'editResult'), 3000, 'editResult');
  assert.equal(csvResult.code, 'blocked');
  assert.equal(csv.document.getText(), 'a,b\n1,2\n');
  assert.deepEqual(T.calls.filter((c) => c[0] === 'applyEdit'), []);
});

test('a Markdown init carries editState (editable) and readOnly = !editable', async () => {
  const md = openText('state.md', '- [ ] a\n', 'markdown');
  md.panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(md.panel, 'init'), 3000, 'markdown init');
  assert.equal(init.readOnly, false);
  // P6 (D-072): the first state says "checking" (task boxes work, WYSIWYG edits wait for the save check) ...
  assert.deepEqual(init.editState, { editable: true, kinds: [], code: 'fidelity', reason: 'Checking the file…' });
  // ... and the state after the check allows mdText.
  const done = await waitFor(
    () => [...md.panel.posted].reverse().find((m) => m.editState && m.editState.reason !== 'Checking the file…')?.editState,
    3000,
    'the save check',
  );
  assert.deepEqual(done, { editable: true, kinds: ['mdText'] });
  // A malformed mdText edit is refused as invalid, answered with the document version.
  md.panel.send(edit({ kind: 'mdText', edits: [] }));
  const result = await waitFor(() => lastOf(md.panel, 'editResult'), 3000, 'editResult');
  assert.deepEqual([result.accepted, result.code, result.version], [false, 'invalid', 1]);
});
