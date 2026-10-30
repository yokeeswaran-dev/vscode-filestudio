// Regression tests for fixed 0.1.x bugs in pure functions of the webview script (media/viewer.js): formula bar, status
// bar, copy payload and slide geometry. test/build-tests.js bundles viewer.js to test/.out/viewer.js with these
// functions exported; the few browser globals the script touches when it loads are stubbed below. Each test names the
// bug id from the 0.1.x bug hunts. Run with `npm test`.
const { test } = require('node:test');
const assert = require('node:assert/strict');

// ----- browser globals viewer.js uses at load time -----
globalThis.acquireVsCodeApi = () => ({ postMessage() {}, getState: () => undefined, setState() {} });
globalThis.window = { addEventListener() {} };
globalThis.document = { getElementById: () => null };
globalThis.navigator ??= { platform: '', userAgent: '' }; // Node 21+ has its own navigator; Node 20 has none

const V = require('../.out/viewer.js');

/** Grid state (`G` in viewer.js) of a csv or sheet view: only `kind` matters to these functions. */
const setKind = (kind) => V.__setGrid({ kind });

/**
 * A SheetView (viewer.js typedef) with rows `cells` ({ 'B2': cell }) and the selection `ranges` (active = first cell of
 * the first range). `rows` / `cols` = SheetMeta row / column info (hidden flags).
 */
function sheetView({ cells = {}, ranges, merges = [], rows = {}, cols = {} }) {
  const byRow = new Map();
  for (const [address, cell] of Object.entries(cells)) {
    const m = /^([A-Z])(\d+)$/.exec(address);
    const r = Number(m[2]) - 1;
    const c = m[1].charCodeAt(0) - 65;
    if (!byRow.has(r)) byRow.set(r, { r, cells: [] });
    byRow.get(r).cells.push({ c, ...cell });
  }
  for (const row of byRow.values()) row.cells.sort((a, b) => a.c - b.c);
  const active = { r: ranges[0].r0, c: ranges[0].c0 };
  const view = {
    index: 0,
    meta: { merges, rows, cols, rowCount: 100, colCount: 26 },
    mergeIndex: V.buildMergeIndex(merges),
    maxR: 1048576,
    maxC: 16384,
    sel: { ranges, active, anchor: active, cursor: active },
  };
  return { view, rowOf: (r) => byRow.get(r) ?? null };
}

/** planCopy + buildCopyPayload, as the copy handler runs them. */
function copy(spec) {
  const { view, rowOf } = sheetView(spec);
  const plan = V.planCopy(view);
  assert.equal(typeof plan, 'object', `copy refused: ${plan}`);
  // `values` (what a paste inside FileStudio uses) is checked in sheet-edit-fix1.test.js.
  const { values: _values, ...payload } = V.buildCopyPayload(plan.view, plan.ranges, plan.mode, rowOf) ?? {};
  return { mode: plan.mode, ...payload };
}

const range = (r0, c0, r1 = r0, c1 = c0) => ({ r0, c0, r1, c1 });

// ===== FORMULA BAR =====

test('H5-F16: the formula bar writes exponents like Excel (1E+21), never JavaScript style (1e+21)', () => {
  const cases = [
    [1e21, '1E+21'],
    [-1e21, '-1E+21'],
    [1e-20, '1E-20'],
    [1.23456789012345e-12, '1.23456789012345E-12'],
    [1e20, '100000000000000000000'],
    [1e-12, '0.000000000001'],
    [123456789012345680, '123456789012346000'],
  ];
  assert.deepEqual(
    cases.map(([n]) => `${n} -> ${V.formulaBarNumber(n)}`),
    cases.map(([n, want]) => `${n} -> ${want}`),
  );
  setKind('sheet');
  assert.equal(V.rawValue({ t: 'n', v: 1e21, w: '1E+21' }), '1E+21');
});

test('R3-024: a CSV number shows its field text in the formula bar (007 stays 007), an xlsx number its value', () => {
  setKind('csv');
  assert.equal(V.rawValue({ t: 'n', v: 7, w: '007' }), '007');
  assert.equal(V.rawValue({ t: 'n', v: 100000, w: '1e5' }), '1e5');
  assert.equal(V.rawValue({ t: 'n', v: 0.5, w: '50%' }), '50%');
  setKind('sheet');
  assert.equal(V.rawValue({ t: 'n', v: 7, w: '007' }), '7', 'xlsx: the number, not its format');
  assert.equal(V.rawValue({ t: 'n', v: 3, w: '3', f: 'A1+2' }), '=A1+2');
});

// ===== STATUS BAR =====

test('H5-F8: the status bar shows very small and very large numbers in General notation (not 0)', () => {
  assert.equal(V.formatStat(1e-12), '1E-12');
  assert.equal(V.formatStat(-2.5e-10), '-2.5E-10');
  assert.equal(V.formatStat(1e21), '1E+21');
  assert.equal(V.formatStat(0), '0');
});

// ===== COPY =====

test('U2: copying one cell or one merged area gives its text as is (no TSV quotes, no tabs)', () => {
  setKind('sheet');
  assert.deepEqual(copy({ cells: { A1: { t: 's', v: 'a\tb', w: 'a\tb' } }, ranges: [range(0, 0)] }), {
    mode: 'cell',
    text: 'a\tb',
    html: '',
    cells: 1,
  });
  const merged = copy({
    cells: { B2: { t: 's', v: 'multi\nline "x"', w: 'multi\nline "x"' } },
    merges: [range(1, 1, 2, 2)],
    ranges: [range(1, 1, 2, 2)],
  });
  assert.deepEqual({ mode: merged.mode, text: merged.text }, { mode: 'cell', text: 'multi\nline "x"' });
  // Two cells are a table: TSV quoting applies there.
  const two = copy({ cells: { A1: { t: 's', v: 'a\tb', w: 'a\tb' }, B1: { t: 's', v: 'c', w: 'c' } }, ranges: [range(0, 0, 0, 1)] });
  assert.equal(two.text, '"a\tb"\tc');
});

test('R2-F1: copy skips hidden rows and columns', () => {
  setKind('sheet');
  const cells = {};
  for (const r of [30, 31, 32]) for (const col of ['B', 'C', 'D']) cells[`${col}${r}`] = { t: 's', v: `${col}${r}`, w: `${col}${r}` };
  const { text } = copy({ cells, ranges: [range(29, 1, 31, 3)], rows: { 30: { hidden: true } }, cols: { 2: { hidden: true } } });
  assert.equal(text, 'B30\tD30\r\nB32\tD32');
});

test('R2-F5: text cells are marked as text in the copied HTML (Excel keeps 007 as text)', () => {
  setKind('sheet');
  const sheetCopy = copy({ cells: { A1: { t: 's', v: '007', w: '007' }, B1: { t: 'n', v: 7, w: '7' } }, ranges: [range(0, 0, 0, 1)] });
  assert.equal(sheetCopy.html, `<meta charset="utf-8"><table><tr><td style="mso-number-format:'\\@'">007</td><td>7</td></tr></table>`);
  setKind('csv');
  const csvCopy = copy({ cells: { A1: { t: 's', v: '007', w: '007' }, B1: { t: 'n', v: 7, w: '7' } }, ranges: [range(0, 0, 0, 1)] });
  assert.ok(!csvCopy.html.includes('mso-number-format'), 'csv: no text marks');
});

// ===== STYLES =====

test('H9-F10: Markdown links follow accessibility.underlineLinks (var(--text-link-decoration)), not a fixed style', () => {
  const css = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', '..', 'media', 'viewer.css'), 'utf8');
  /** Declarations of the rule whose selector list is exactly `selectors`. */
  const ruleOf = (selectors) => {
    const escaped = selectors.map((s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join(',\\s*');
    return new RegExp(`(?:^|\\})\\s*${escaped}\\s*\\{([^}]*)\\}`, 'm').exec(css)?.[1];
  };
  for (const selectors of [['.md-view .md-article a'], ['.md-view .md-article .footnote-ref a', '.md-view .md-article .footnote-backref']]) {
    const body = ruleOf(selectors);
    assert.ok(body, `rule ${selectors.join(', ')} exists`);
    assert.match(body, /text-decoration: var\(--text-link-decoration, none\);/, selectors.join(', '));
  }
});

// ===== SLIDES (PPTX) =====

/** Numbers of an SVG path, rounded to 2 decimals. */
const numbersOf = (d) => (d.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number);

test('R3-016: wedgeEllipseCallout is drawn as an ellipse bubble with a pointer (not pptxtojson\'s tiny triangle)', () => {
  const pptxtojsonPath = 'M 0 0 L 9 0 L 4.5 4.5 Z';
  const d = V.pptxShapePathData({ path: pptxtojsonPath, pathViewBox: { x: 0, y: 0, width: 90, height: 90 } }, 'wedgeEllipseCallout', 90, 90);
  // Default adjustments (-20833, 62500): pointer tip at (45 - 18.75, 45 + 56.25); the arc is the large one, radius 45.
  assert.match(d, /^M 26\.25 101\.25 L [\d.]+ [\d.]+ A 45 45 0 1 1 [\d.]+ [\d.]+ Z$/);
  const [, , x1, y1, , , , , , x2, y2] = numbersOf(d);
  for (const [x, y] of [[x1, y1], [x2, y2]]) assert.ok(Math.abs(Math.hypot(x - 45, y - 45) - 45) < 0.05, `(${x}, ${y}) is on the ellipse`);
});

test('R3-043: block arrow heads are as long as PowerPoint draws them (half the shorter side by default)', () => {
  const wrong = { path: 'M 0 12.5 L 100 12.5 L 100 0 L 200 25 L 100 50 L 100 37.5 L 0 37.5 Z', pathViewBox: { x: 0, y: 0, width: 200, height: 50 } };
  assert.equal(V.pptxShapePathData(wrong, 'rightArrow', 200, 50), 'M 0 12.5 L 175 12.5 L 175 0 L 200 25 L 175 50 L 175 37.5 L 0 37.5 Z');
  assert.equal(V.pptxShapePathData({}, 'downArrow', 50, 200), 'M 12.5 0 L 12.5 175 L 0 175 L 25 200 L 50 175 L 37.5 175 L 37.5 0 Z');
  assert.equal(
    V.pptxShapePathData({ keypoints: { adj2: 0.5 } }, 'leftArrow', 200, 50),
    'M 200 12.5 L 12.5 12.5 L 12.5 0 L 0 25 L 12.5 50 L 12.5 37.5 L 200 37.5 Z',
    'adj2 (keypoint 0.5 = 25000) = a quarter of the shorter side',
  );
});

test('R3-044: text of an ellipse / arrow sits in the preset text rectangle, not the whole box', () => {
  const round = (rect) => Object.fromEntries(Object.entries(rect).map(([k, v]) => [k, Math.round(v * 100) / 100]));
  assert.deepEqual(round(V.pptxTextRect('ellipse', 200, 100)), { l: 29.29, t: 14.64, r: 170.71, b: 85.36 });
  assert.deepEqual(round(V.pptxTextRect('wedgeEllipseCallout', 200, 100)), { l: 29.29, t: 14.64, r: 170.71, b: 85.36 });
  assert.deepEqual(round(V.pptxTextRect('rightArrow', 200, 50)), { l: 0, t: 12.5, r: 187.5, b: 37.5 });
  assert.equal(V.pptxTextRect('rect', 200, 50), null, 'rect: the whole box');
});

test('R3-047: gradients blend in linear light like PowerPoint (red -> blue midpoint #BC00BC, not #800080)', () => {
  const css = V.pptxCssGradient({ colors: [{ pos: '0%', color: '#FF0000' }, { pos: '100%', color: '#0000FF' }], rot: 0 });
  assert.match(css, /^linear-gradient\(90deg, #FF0000 0%, /);
  assert.ok(css.includes('rgba(188, 0, 188, 1) 50%'), css);
  assert.match(css, /, #0000FF 100%\)$/);
});
