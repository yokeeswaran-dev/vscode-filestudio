// CSV / TSV / PSV / SSV editing (src/renderers/sheet.ts `// ----- CSV EDITS -----`, contract §2.4 + §8.3): record
// index, quote policy, minimal edit planners and index shifting. Ports the s2-protocol prototype tests (the generated
// tricky cases, the quote-policy table and the seeded fuzz). Run with `npm test` (it bundles the renderers first).
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const S = require('../.out/sheet.js');

const SAMPLES = path.join(__dirname, '..', 'samples');

// ----- helpers -----

/** VS Code's TextModel load rule: pick CRLF when CR-based breaks are more than half of all breaks, then unify. */
function vscodeNormalize(text) {
  let cr = 0;
  let lf = 0;
  let crlf = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    if (ch === 13) {
      if (text.charCodeAt(i + 1) === 10) {
        crlf++;
        i++;
      } else cr++;
    } else if (ch === 10) lf++;
  }
  const total = cr + lf + crlf;
  const eol = total === 0 ? '\n' : cr + crlf > total / 2 ? '\r\n' : '\n';
  return { text: total ? text.replace(/\r\n|\r|\n/g, eol) : text, eol };
}

/** document.getText() of a file's bytes: UTF-8 (invalid bytes -> U+FFFD), one BOM removed, line breaks unified. */
function documentText(bytes) {
  let text = Buffer.from(bytes).toString('utf8');
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  return vscodeNormalize(text);
}

/** The pinned delimiter of a file type (= VP parseDelimited without a chosen delimiter). */
function pinOf(text, name) {
  const ext = path.extname(name).slice(1).toLowerCase();
  if (ext === 'tsv') return { delimiter: '\t', collapse: false };
  if (ext === 'psv') return { delimiter: '|', collapse: false };
  if (ext === 'ssv') return S.detectSsvDelimiter(text) === ';' ? { delimiter: ';', collapse: false } : { delimiter: ' ', collapse: true };
  return { delimiter: S.parseCsv(text, { forView: true }).delimiter, collapse: false };
}

/** The rows the grid shows for this pin (= VP parseDelimited with that delimiter). */
function viewRows(text, pin) {
  return S.parseCsv(text, pin.collapse ? { delimiter: ' ', collapseSpaces: true, forView: true } : { delimiter: pin.delimiter, forView: true }).rows;
}

function indexRows(text, idx) {
  const rows = [];
  for (let r = 0; r < idx.count; r++) rows.push(S.decodeCsvRecord(text, idx, r));
  return rows;
}

function applyEdits(text, edits) {
  let out = text;
  const sorted = [...edits].sort((a, b) => b.start - a.start);
  for (let i = 1; i < sorted.length; i++) assert.ok(sorted[i].end <= sorted[i - 1].start, 'edits overlap');
  for (const e of sorted) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return out;
}

/** Plans one intent on a fresh index + policy and applies it. */
function edit(text, intent, opts = {}) {
  const delimiter = opts.delimiter ?? ',';
  const collapse = !!opts.collapse;
  const eol = opts.eol ?? '\n';
  const idx = S.indexCsv(text, delimiter, collapse);
  const policy = S.csvQuotePolicy(text, idx);
  const edits = S.planCsvEdit(text, idx, policy, eol, intent);
  return { edits, out: applyEdits(text, edits), idx };
}

const set = (cells) => ({ kind: 'setCells', sheetId: 'csv', cells });
const rowsOp = (kind, at, count) => ({ kind, sheetId: 'csv', at, count });
const sortOp = (r0, r1, key, desc = false, header = false) => ({ kind: 'sort', sheetId: 'csv', range: { r0, c0: 0, r1, c1: 50 }, key, desc, header });
const clearOp = (...ranges) => ({ kind: 'clear', sheetId: 'csv', ranges, what: 'contents' });
const replaceOp = (text, replace, o = {}) => ({
  kind: 'replaceAll',
  query: { scope: 'csv', text, matchCase: !!o.matchCase, regex: !!o.regex, wholeCell: !!o.wholeCell, lookIn: 'values' },
  replace,
});

/** True when offset k of s lies inside a CRLF or a surrogate pair. */
function splitsPair(s, k) {
  if (k <= 0 || k >= s.length) return false;
  const a = s.charCodeAt(k - 1);
  const b = s.charCodeAt(k);
  return (a === 13 && b === 10) || (a >= 0xd800 && a <= 0xdbff && b >= 0xdc00 && b <= 0xdfff);
}

function sameIndex(a, b) {
  const head = (x) => ({ count: x.count, sepLine: x.sepLine, trailingNewline: x.trailingNewline, unclosed: x.unclosedQuoteRecord });
  if (JSON.stringify(head(a)) !== JSON.stringify(head(b))) return false;
  for (let r = 0; r < a.count; r++) {
    if (a.start[r] !== b.start[r] || a.contentEnd[r] !== b.contentEnd[r] || a.end[r] !== b.end[r]) return false;
  }
  return true;
}

// The tricky cases of the s2-protocol prototype (raw file text; a leading \uFEFF is a BOM byte sequence).
const GENERATED = {
  'g-quoted-nl-crlf.csv': 'id,text,n\r\n1,"line1\r\nline2",3\r\n2,"a ""q"" b",4\r\n3,plain,5\r\n',
  'g-sep-bom.csv': '\uFEFFsep=;\r\nA;B;C\r\n1,5;x;"y;z"\r\n2,25;;\r\n',
  'g-ragged.csv': 'a,b,c,d\n1,2\n1,2,3,4,5,6\n\n7\n,,\n',
  'g-blank-trailing.csv': 'h1,h2\nx,y\n\n\n',
  'g-no-final-nl.csv': 'h1,h2\nx,y\nlast,row',
  'g-unclosed.csv': 'h1,h2\nx,"never closes\nmore,text\n',
  'g-excel-junk.csv': 'name,size\n"12" monitor,5\n"Smith" ,x\n"b"c"d",9\n',
  'g-quote-all.csv': '"id","name","amount"\r\n"1","Ann","10.5"\r\n"2","Bob","7"\r\n"3","Cy","8"\r\n',
  'g-mixed-style.csv': '"name",age,"city"\n"Ann",31,"Oslo"\n"Bob",40,"Rome"\n"Cy",22,"Lima"\n"Di",50,"Kyiv"\n',
  'g-emoji.csv': 'k,v\n😀,"🎉, party"\n𝒳,y\n',
  'g-double-bom.csv': '\uFEFF\uFEFFa,b\n1,2\n',
  'g-empty.csv': '',
  'g-header-only.csv': 'a,b,c',
  'g-one-nl.csv': '\r\n',
  'g-tsv-quotes.tsv': 'a\tb\n"x\ty"\t2\nq"uo\t3\n',
  'g-space.ssv': 'name   age  city\nAnn    31   Oslo\n"Bo B" 40   "New York"\n  Cy   22   Lima  \n\n',
  'g-semicolon.ssv': 'a;b;c\n1,5;2;3\n4;5,25;6\n',
  'g-pipe.psv': 'a|b|c\n1|"x|y"|3\n|||\n',
  'g-cr-only.csv': 'a,b\r1,2\r3,4\r',
  'g-mixed-eol.csv': 'a,b\r\n1,"x\ny"\n3,4\r\n5,6\r\n',
  'g-single-col.csv': 'v\nx;y\np;q\n',
};

/** Every delimited sample (top level + csv-edit/) and generated case: [name, bytes]. */
function allCases() {
  const out = [];
  const add = (dir, prefix) => {
    for (const name of fs.readdirSync(dir).sort()) {
      if (/\.(csv|tsv|psv|ssv)$/i.test(name)) out.push([prefix + name, fs.readFileSync(path.join(dir, name))]);
    }
  };
  add(SAMPLES, '');
  add(path.join(SAMPLES, 'csv-edit'), 'csv-edit/');
  for (const [name, text] of Object.entries(GENERATED)) out.push([name, Buffer.from(text, 'utf8')]);
  return out;
}

// ----- index -----

test('the record index decodes every sample, fixture and generated case exactly like the grid', () => {
  const cases = allCases();
  // 6 samples + 12 csv-edit fixtures + 21 generated cases
  assert.ok(cases.length >= 39, `found only ${cases.length} cases (are the csv-edit fixtures generated?)`);
  for (const [name, bytes] of cases) {
    const { text } = documentText(bytes);
    const pin = pinOf(text, name);
    const idx = S.indexCsv(text, pin.delimiter, pin.collapse);
    // the records cover the text after the BOM / sep= line exactly
    let joined = idx.count ? text.slice(0, idx.start[0]) : '';
    for (let r = 0; r < idx.count; r++) {
      assert.equal(idx.start[r], r ? idx.end[r - 1] : idx.start[0], `${name}: record ${r} does not follow the previous one`);
      joined += text.slice(idx.start[r], idx.end[r]);
    }
    if (idx.count) assert.equal(joined, text, `${name}: spans do not cover the text`);
    assert.deepEqual(indexRows(text, idx), viewRows(text, pin), `${name}: decode differs from the view`);
  }
});

test('index facts of the csv-edit fixtures', () => {
  const load = (name) => {
    const { text, eol } = documentText(fs.readFileSync(path.join(SAMPLES, 'csv-edit', name)));
    const pin = pinOf(text, name);
    return { text, eol, idx: S.indexCsv(text, pin.delimiter, pin.collapse) };
  };
  const ragged = load('ragged.csv');
  assert.equal(ragged.idx.count, 9);
  assert.equal(ragged.idx.trailingNewline, true);
  assert.equal(load('no-final-newline.csv').idx.trailingNewline, false);
  assert.equal(load('sep-line.csv').idx.sepLine, true);
  assert.equal(load('sep-line.csv').idx.delimiter, ';');
  assert.equal(load('unclosed-quote.csv').idx.unclosedQuoteRecord, 1);
  assert.equal(load('ragged.csv').idx.unclosedQuoteRecord, -1);
  const doubleBom = load('double-bom.csv');
  // VS Code removes the first BOM; the second one stays in the text, before the first record (the grid skips it, as
  // parseCsv does), so edits never touch it
  assert.equal(doubleBom.text.charCodeAt(0), 0xfeff);
  assert.equal(doubleBom.idx.start[0], 1);
  assert.equal(S.decodeCsvRecord(doubleBom.text, doubleBom.idx, 0)[0], 'a');
  assert.ok(edit(doubleBom.text, set([[0, 0, 'z']])).out.startsWith('\uFEFFz,b'));
  assert.equal(load('mixed-eol.csv').eol, '\r\n');
  assert.equal(load('cr-only.csv').eol, '\r\n');
  const latin = load('latin1.csv');
  assert.ok(latin.text.includes('�'), 'Latin-1 bytes decode to U+FFFD as UTF-8');
  const space = load('space.ssv');
  assert.equal(space.idx.collapse, true);
  assert.deepEqual(S.decodeCsvRecord(space.text, space.idx, 2), ['Bo B', '40', 'New York']);
  assert.deepEqual(S.decodeCsvRecord(space.text, space.idx, 5), ['']);
});

// ----- quote policy -----

test('quote-policy table: new values are quoted the way the file quotes that column / class', () => {
  const t = (label, text, delimiter, r, c, v, want) => {
    assert.equal(edit(text, set([[r, c, v]]), { delimiter }).out, want, label);
  };
  t('quote-all text', '"a","b"\n"1","x"\n"2","y"\n"3","z"\n', ',', 2, 1, 'new', '"a","b"\n"1","x"\n"2","new"\n"3","z"\n');
  t('quote-all number', '"a","b"\n"1","x"\n"2","y"\n"3","z"\n', ',', 2, 0, '9', '"a","b"\n"1","x"\n"9","y"\n"3","z"\n');
  t('minimal', 'a,b\n1,x\n2,y\n3,z\n', ',', 2, 1, 'new', 'a,b\n1,x\n2,new\n3,z\n');
  t('minimal forced', 'a,b\n1,x\n2,y\n3,z\n', ',', 2, 1, 'n,ew', 'a,b\n1,x\n2,"n,ew"\n3,z\n');
  t('text quoted, numbers not', '"n",v\n"a",1\n"b",2\n"c",3\n', ',', 2, 0, 'q', '"n",v\n"a",1\n"q",2\n"c",3\n');
  t('text quoted, numbers not (number)', '"n",v\n"a",1\n"b",2\n"c",3\n', ',', 2, 1, '7', '"n",v\n"a",1\n"b",7\n"c",3\n');
  t('semicolon decimal comma', 'a;b\n1,5;x\n2,5;y\n3,5;z\n', ';', 2, 0, '4,25', 'a;b\n1,5;x\n4,25;y\n3,5;z\n');
  t('neighbour quoting kept', 'a,b,c\n" x ",2,"y"\n', ',', 1, 1, '3', 'a,b,c\n" x ",3,"y"\n');
  t('ragged pad', 'a,b,c\n1\n', ',', 1, 2, 'z', 'a,b,c\n1,,z\n');
  t('append without a final line break', 'a,b\n1,2', ',', 2, 0, 'n', 'a,b\n1,2\nn,');
  t('append with a final line break', 'a,b\n1,2\n', ',', 2, 1, 'n', 'a,b\n1,2\n,n\n');
  // the nonNumeric fix: a ';' file with text quoted and decimal-comma numbers unquoted
  const fr = '"P";"Prix"\n"Pain";1,25\n"Lait";0,99\n"Sel";2\n';
  t('decimal comma number stays unquoted', fr, ';', 2, 1, '4,5', '"P";"Prix"\n"Pain";1,25\n"Lait";4,5\n"Sel";2\n');
  t('text in a quoted-text file is quoted', fr, ';', 2, 0, 'Thé', '"P";"Prix"\n"Pain";1,25\n"Thé";0,99\n"Sel";2\n');
  t('a value with a quote is always quoted', 'a,b\n1,2\n', ',', 1, 0, 'say "hi"', 'a,b\n"say ""hi""",2\n');
  t('a value starting with a quote is quoted', 'a,b\n1,2\n', ',', 1, 1, '"', 'a,b\n1,""""\n');
});

// ----- trimEdit -----

test('trimEdit keeps the common prefix / suffix and never splits a surrogate pair or a CRLF', () => {
  assert.equal(S.trimEdit('abc', 0, 3, 'abc'), null);
  assert.deepEqual(S.trimEdit('a😀b', 0, 4, 'a😁b'), { start: 1, end: 3, text: '😁' });
  assert.deepEqual(S.trimEdit('x\r\ny', 0, 4, 'x\rz'), { start: 1, end: 4, text: '\rz' });
  assert.deepEqual(S.trimEdit('a\r\nb', 0, 4, 'c\nb'), { start: 0, end: 3, text: 'c\n' });
  // the common prefix would end between the halves of the second emoji: it stops before that pair
  assert.deepEqual(S.trimEdit('😀😀', 0, 4, '😀😁😀'), { start: 2, end: 2, text: '😁' });
  // a CRLF in the middle of a shared suffix stays whole
  assert.deepEqual(S.trimEdit('a\r\nb', 0, 4, 'x\r\nb'), { start: 0, end: 1, text: 'x' });
  // property: applying the trimmed edit gives the full replacement; no boundary inside a pair (old or new text)
  const tokens = ['a', 'b', '\r\n', '\r', '\n', '😀', '😁', '"'];
  let seed = 11;
  const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  const word = () => Array.from({ length: Math.floor(rnd() * 6) }, () => tokens[Math.floor(rnd() * tokens.length)]).join('');
  for (let k = 0; k < 3000; k++) {
    const before = word();
    const mid = word();
    const after = word();
    const insert = word();
    const text = before + mid + after;
    // the planners only pass record / field boundaries, which never lie inside a pair
    if (splitsPair(text, before.length) || splitsPair(text, before.length + mid.length)) continue;
    const e = S.trimEdit(text, before.length, before.length + mid.length, insert);
    if (!e) {
      assert.equal(mid, insert);
      continue;
    }
    assert.equal(text.slice(0, e.start) + e.text + text.slice(e.end), before + insert + after);
    assert.ok(!splitsPair(text, e.start) && !splitsPair(text, e.end), `boundary inside a pair: ${JSON.stringify([text, e])}`);
  }
});

// ----- planner rules -----

test('more than 2,000 changed records become one edit; fewer stay one edit per record', () => {
  const big = 'k,v\n' + Array.from({ length: 3000 }, (_, i) => `${i},x${i}`).join('\n') + '\n';
  const cells = Array.from({ length: 2500 }, (_, i) => [i + 1, 1, `y${i}`]);
  const one = edit(big, set(cells));
  assert.equal(one.edits.length, 1);
  const rows = viewRows(one.out, { delimiter: ',', collapse: false });
  assert.equal(rows[1][1], 'y0');
  assert.equal(rows[2500][1], 'y2499');
  assert.equal(rows[2501][1], 'x2500');
  assert.equal(rows.length, 3001);
  const few = edit(big, set(cells.slice(0, 1500)));
  assert.equal(few.edits.length, 1500);
  const cleared = edit(big, clearOp({ r0: 0, c0: 1, r1: 1048575, c1: 1 }));
  assert.equal(cleared.edits.length, 1);
  assert.ok(viewRows(cleared.out, { delimiter: ',', collapse: false }).every((row) => row[1] === ''));
});

test('a quote that never closes blocks every edit kind with the row number', () => {
  const text = 'h1,h2\nx,"never closes\nmore,text\n';
  const intents = [
    set([[0, 0, 'q']]),
    clearOp({ r0: 0, c0: 0, r1: 0, c1: 0 }),
    rowsOp('insertRows', 0, 1),
    rowsOp('deleteRows', 0, 1),
    rowsOp('insertCols', 0, 1),
    rowsOp('deleteCols', 0, 1),
    sortOp(1, 5, 0),
    replaceOp('h', 'x'),
  ];
  for (const intent of intents) {
    assert.throws(
      () => edit(text, intent),
      (error) =>
        error instanceof S.CsvEditBlocked &&
        error.code === 'unclosedQuote' &&
        error.message === 'Row 2 opens a quote that is never closed; editing could change the rest of the file.',
      intent.kind,
    );
  }
});

test('column insert / delete is blocked in space-separated text; other edits work there', () => {
  const text = 'name   age  city\nAnn    31   Oslo\n"Bo B" 40   "New York"\n\n';
  const opts = { delimiter: ' ', collapse: true };
  for (const kind of ['insertCols', 'deleteCols']) {
    assert.throws(
      () => edit(text, rowsOp(kind, 1, 1), opts),
      (error) => error instanceof S.CsvEditBlocked && error.code === 'spaceColumns' && error.message === 'Columns cannot be inserted or deleted in a space-separated file.',
    );
  }
  const typed = edit(text, set([[1, 2, 'Los Angeles'], [3, 1, '7'], [1, 0, '']]), opts).out;
  assert.equal(typed, 'name   age  city\n""    31   "Los Angeles"\n"Bo B" 40   "New York"\n"" 7\n');
  assert.deepEqual(viewRows(typed, opts), [['name', 'age', 'city'], ['', '31', 'Los Angeles'], ['Bo B', '40', 'New York'], ['', '7']]);
});

test('column insert / delete leaves blank records (empty line or "") byte for byte', () => {
  const text = 'a,b\n""\n\nx,y\n';
  assert.equal(edit(text, rowsOp('deleteCols', 0, 1)).out, 'b\n""\n\ny\n');
  assert.equal(edit(text, rowsOp('deleteCols', 0, 2)).out, '\n""\n\n\n');
  assert.equal(edit(text, rowsOp('insertCols', 0, 1)).out, ',a,b\n""\n\n,x,y\n');
});

test('sort never moves record 0 or the sep= line, keeps trailing blank records, moves records byte for byte', () => {
  const text = 'sep=,\nname,n\n"Cy",10\nann,9\n"Bob",\nDi,abc\n,x\n\n\n';
  // the range includes row 1 (r0 = 0): clamped to start at row 2
  const { out } = edit(text, sortOp(0, 1048575, 1));
  assert.equal(out, 'sep=,\nname,n\nann,9\n"Cy",10\nDi,abc\n,x\n"Bob",\n\n\n', 'numbers, then text, then blank keys; trailing blanks stay');
  const desc = edit(text, sortOp(0, 1048575, 1, true)).out;
  assert.equal(desc, 'sep=,\nname,n\n,x\nDi,abc\n"Cy",10\nann,9\n"Bob",\n\n\n', 'descending: text before numbers, blanks still last');
  const byName = edit(text, sortOp(1, 5, 0)).out;
  assert.equal(byName, 'sep=,\nname,n\nann,9\n"Bob",\n"Cy",10\nDi,abc\n,x\n\n\n', 'case-insensitive text, blank key last, header row stays');
  // header flag: the first row of the range is a header too
  const headed = edit('h\n3\n2\n1\n', sortOp(1, 3, 0, false, true)).out;
  assert.equal(headed, 'h\n3\n1\n2\n');
  // a ';' file sorts decimal-comma numbers as numbers (as the grid reads them)
  assert.equal(edit('a;b\n10,5;x\n9,75;y\n', sortOp(1, 2, 0), { delimiter: ';' }).out, 'a;b\n9,75;y\n10,5;x\n');
  // stable for equal keys
  assert.equal(edit('k,v\n1,a\n1,b\n0,c\n1,d\n', sortOp(1, 4, 0)).out, 'k,v\n0,c\n1,a\n1,b\n1,d\n');
  // nothing to sort
  assert.deepEqual(edit('k\n1\n', sortOp(0, 5, 0)).edits, []);
});

test('values are stored exactly as typed; line breaks inside a value use the document EOL', () => {
  assert.equal(edit('a,b\r\n1,2\r\n', set([[1, 0, '=1+1']]), { eol: '\r\n' }).out, 'a,b\r\n=1+1,2\r\n');
  assert.equal(edit('a,b\r\n1,2\r\n', set([[1, 1, 'x\ny']]), { eol: '\r\n' }).out, 'a,b\r\n1,"x\r\ny"\r\n');
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, 'x\r\ny\rz']]), { eol: '\n' }).out, 'a,b\n1,"x\ny\nz"\n');
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, '  pad  ']])).out, 'a,b\n1,"  pad  "\n', 'trailing spaces: quoted (R2)');
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, '😀']])).out, 'a,b\n1,😀\n');
  // '' on a missing cell is a no-op; the same value is a no-op
  assert.deepEqual(edit('a,b\n1\n', set([[1, 5, ''], [7, 0, ''], [1, 0, '1']])).edits, []);
});

test('appending below the data: rows as wide as the header, gaps empty, BOM / sep= line kept', () => {
  assert.equal(edit('a,b,c\n1,2,3\n', set([[3, 1, 'x']])).out, 'a,b,c\n1,2,3\n,,\n,x,\n');
  assert.equal(edit('', set([[0, 0, 'x']])).out, 'x');
  assert.equal(edit('\uFEFF', set([[0, 1, 'x']])).out, '\uFEFF,x');
  // a 'sep=' line without a line break would swallow the first record
  assert.equal(edit('sep=;', set([[0, 0, 'x']]), { delimiter: ';' }).out, 'sep=;\nx');
  assert.equal(edit('sep=;\r\n', set([[0, 0, 'x']]), { delimiter: ';', eol: '\r\n' }).out, 'sep=;\r\nx');
  // insert rows below the data does nothing (virtual rows)
  assert.deepEqual(edit('a\n1\n', rowsOp('insertRows', 2, 3)).edits, []);
});

test('the last record of a file without a final line break never vanishes', () => {
  // clearing its only field writes "" (an empty last record without a line break cannot exist)
  assert.equal(edit('h\nx', set([[1, 0, '']])).out, 'h\n""');
  assert.equal(edit('a,b\nx,y', rowsOp('deleteCols', 0, 2)).out, '\n""');
  // deleting the last record keeps "no final line break"...
  assert.equal(edit('h\na\nb', rowsOp('deleteRows', 2, 1)).out, 'h\na');
  assert.equal(edit('h\na\nb', rowsOp('deleteRows', 0, 3)).out, '');
  assert.equal(edit('sep=,\nh\na', rowsOp('deleteRows', 0, 2)).out, 'sep=,\n');
  // ...unless the new last record is empty: then its line break stays
  const kept = edit('h\n\nx', rowsOp('deleteRows', 2, 1)).out;
  assert.equal(kept, 'h\n\n');
  assert.deepEqual(viewRows(kept, { delimiter: ',' }), [['h'], ['']]);
  // sort that moves an empty record to the very end keeps it as a line
  const sorted = edit('k\nb\n\na', sortOp(1, 3, 0)).out;
  assert.deepEqual(viewRows(sorted, { delimiter: ',' }), [['k'], ['a'], ['b'], ['']]);
});

test('results over 50 MB are blocked before big strings are built', () => {
  const tooLarge = (error) => error instanceof S.CsvEditBlocked && error.code === 'tooLarge' && /larger than 50 MB/.test(error.message);
  const text = 'a,b,c\n1,2,3\n';
  assert.throws(() => edit(text, set([[30_000_000, 0, 'x']])), tooLarge);
  assert.throws(() => edit(text, set([[1, 60_000_000, 'x']])), tooLarge);
  assert.throws(() => edit(text, rowsOp('insertRows', 1, 20_000_000)), tooLarge);
  assert.throws(() => edit(text, rowsOp('insertCols', 0, 60_000_000)), tooLarge);
  assert.throws(() => edit(text, set([[1, 1, 'x'.repeat(52_500_000)]])), tooLarge);
});

test('kinds that are not CSV edits and bad numbers are refused', () => {
  assert.throws(() => edit('a\n', { kind: 'merge', sheetId: 'csv', range: { r0: 0, c0: 0, r1: 1, c1: 1 } }), /not available/);
  assert.throws(() => edit('a\n', set([[-1, 0, 'x']])), /invalid/);
  assert.throws(() => edit('a\n', set([[0.5, 0, 'x']])), /invalid/);
  assert.throws(() => edit('a\n', rowsOp('deleteRows', 0, Number.NaN)), /invalid/);
  assert.deepEqual(edit('a\n', { kind: 'clear', sheetId: 'csv', ranges: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], what: 'formats' }).edits, []);
});

test('Replace All: plain text, case, whole cell, regex groups; empty cells untouched; risky patterns refused', () => {
  const text = 'name,city\nAnn,Oslo\nanna,OSLO\n,x\n';
  assert.equal(edit(text, replaceOp('ann', 'Jo')).out, 'name,city\nJo,Oslo\nJoa,OSLO\n,x\n');
  assert.equal(edit(text, replaceOp('ann', 'Jo', { matchCase: true })).out, 'name,city\nAnn,Oslo\nJoa,OSLO\n,x\n');
  assert.equal(edit(text, replaceOp('oslo', 'Rome', { wholeCell: true })).out, 'name,city\nAnn,Rome\nanna,Rome\n,x\n');
  assert.equal(edit(text, replaceOp('(A)(nn)', '$2$1', { regex: true, matchCase: true })).out, 'name,city\nnnA,Oslo\nanna,OSLO\n,x\n');
  assert.equal(edit(text, replaceOp('$1', 'a,b')).edits.length, 0, 'plain text: $ is literal, nothing matches');
  assert.equal(edit('k\n$1\n', replaceOp('$1', 'a,b')).out, 'k\n"a,b"\n', 'plain text replacement is literal; quoted when needed');
  assert.equal(edit(text, replaceOp('^$', 'EMPTY', { regex: true })).edits.length, 0, 'empty values are never changed');
  // Replace All covers every row the grid shows, row 1 (the header) too, like Excel
  assert.equal(edit('k\n😀x\n', replaceOp('.', '-', { regex: true })).out, '-\n--\n', 'a match never splits a surrogate pair');
  for (const risky of ['(a+)+', '(x*)*', '(?:a|a)+', '(a|ab)*', '((b+))+c', '(\\w+\\s?)+$']) {
    assert.deepEqual(S.csvFindPattern({ scope: 'csv', text: risky, matchCase: false, regex: true, wholeCell: false, lookIn: 'values' }), {
      invalid: 'This pattern could take too long; simplify it.',
    }, risky);
    assert.throws(() => edit(text, replaceOp(risky, 'x', { regex: true })), /take too long/);
  }
  for (const fine of ['(ab)+', '\\d+', '(?:foo|bar)+', '[a+]+', 'a{2,3}', '(a)?b']) {
    assert.ok('re' in S.csvFindPattern({ scope: 'csv', text: fine, matchCase: false, regex: true, wholeCell: false, lookIn: 'values' }), fine);
  }
  assert.ok('invalid' in S.csvFindPattern({ scope: 'csv', text: '(', matchCase: false, regex: true, wholeCell: false, lookIn: 'values' }));
  assert.ok('invalid' in S.csvFindPattern({ scope: 'csv', text: 'x'.repeat(1001), matchCase: false, regex: false, wholeCell: false, lookIn: 'values' }));
  assert.ok('re' in S.csvFindPattern({ scope: 'csv', text: 'a.b(c)[d]{e}|f\\g/h^$*+?', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' }));
});

test('shiftCsvIndex re-scans only what changed and falls back to a full index when records split differently', () => {
  const text = 'a,b\n1,"x\ny"\n3,4\n5,6\n';
  const idx = S.indexCsv(text, ',', false);
  const policy = S.csvQuotePolicy(text, idx);
  const edits = S.planCsvEdit(text, idx, policy, '\n', set([[1, 1, 'z'], [3, 0, 'long value']]));
  const out = applyEdits(text, edits);
  assert.ok(sameIndex(S.shiftCsvIndex(idx, edits, out), S.indexCsv(out, ',', false)));
  // an edit that opens a quote changes how every later record splits: same result as a full index
  const opened = [{ start: idx.start[2], end: idx.start[2], text: '"' }];
  const broken = applyEdits(text, opened);
  assert.ok(sameIndex(S.shiftCsvIndex(idx, opened, broken), S.indexCsv(broken, ',', false)));
  // no edits: the same index
  assert.equal(S.shiftCsvIndex(idx, [], text), idx);
});

// ----- seeded fuzz (s2-protocol E1 port): every delimiter, >= 3 seeds x 2,000 ops each -----

const MODES = [
  { name: 'comma', delimiter: ',', collapse: false },
  { name: 'semicolon', delimiter: ';', collapse: false },
  { name: 'tab', delimiter: '\t', collapse: false },
  { name: 'pipe', delimiter: '|', collapse: false },
  { name: 'ssv-semicolon', delimiter: ';', collapse: false, french: true },
  { name: 'ssv-space', delimiter: ' ', collapse: true },
];
const SEEDS = [1, 7, 2024];
const OPS_PER_SEED = 2000;
const fuzzCollator = new Intl.Collator(undefined, { sensitivity: 'base' });
const VALUES = ['', 'x', 'hello world', '42', '3,5', '1.5', 'a,b', 'a;b', 'p|q', 't\tu', 'say "hi"', 'multi\nline', '=1+1', '  pad  ', '😀', '"', '007', 'Zoë', '-1e3', 'x\r\ny'];

function makeRandom(seed) {
  let s = seed;
  const rnd = () => ((s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  return { rnd, ri: (n) => Math.floor(rnd() * n), pick: (a) => a[Math.floor(rnd() * a.length)] };
}

/** A random file text in VS Code's form (one EOL everywhere) for a mode, written the way real tools quote. */
function randomText(R, mode) {
  const eol = R.rnd() < 0.5 ? '\r\n' : '\n';
  const quoteAll = R.rnd() < 0.2;
  const quoteText = !quoteAll && R.rnd() < 0.25;
  const width = 1 + R.ri(5);
  const records = R.ri(22);
  const pool = mode.french ? [...VALUES, '1 234,56', 'Crème brûlée', '12,5 %'] : VALUES;
  const encode = (v) => {
    const value = v.replace(/\r\n|\r|\n/g, eol);
    const must =
      value.includes(mode.delimiter) || value.includes('"') || value.includes('\n') || value.includes('\r') ||
      (mode.collapse && (value === '' || value.includes(' ') || value.includes('\t'))) || value.endsWith(' ') || value.endsWith('\t');
    const numeric = /^[-+]?\d+([.,]\d+)?$/.test(value);
    return must || quoteAll || (quoteText && !numeric && value !== '') ? `"${value.replace(/"/g, '""')}"` : value;
  };
  const lines = [];
  for (let r = 0; r < records; r++) {
    if (r && R.rnd() < 0.06) {
      lines.push(mode.collapse && R.rnd() < 0.5 ? '   ' : '');
      continue;
    }
    const n = R.rnd() < 0.12 ? 1 + R.ri(width + 2) : width;
    const fields = Array.from({ length: n }, () => encode(R.pick(pool)));
    if (mode.collapse) {
      const gaps = fields.map(() => ' '.repeat(1 + R.ri(3)));
      lines.push((R.rnd() < 0.2 ? '  ' : '') + fields.map((f, i) => (i ? gaps[i] : '') + f).join('') + (R.rnd() < 0.2 ? '  ' : ''));
    } else {
      lines.push(fields.join(mode.delimiter));
    }
  }
  let text = lines.join(eol);
  if (records && R.rnd() < 0.8) text += eol;
  if (records && R.rnd() < 0.1) text += eol.repeat(1 + R.ri(2)); // trailing blank records
  if (!mode.collapse && R.rnd() < 0.12) text = `sep=${mode.delimiter}${eol}${text}`;
  if (R.rnd() < 0.05) text = `\uFEFF${text}`; // a second BOM (the first one is never in the text)
  return { text, eol };
}

const isBlankRow = (row) => row.length === 1 && row[0] === '';

/** Number the grid shows for a value (CsvGridModel), or undefined for text. */
function gridNumber(v, delimiter) {
  const model = new S.CsvGridModel({ delimiter, newline: '\n', hasBom: false, rows: [[v]], trailingNewline: false }, 'x.csv');
  const cell = model.getRows(0, 0, 1).rows[0]?.cells[0];
  return cell && cell.t === 'n' ? cell.v : undefined;
}

/** Expected rows (and the record mapping for byte checks) of one intent, from the rows before it. */
function expectRows(rows, intent, mode, eol) {
  const width = rows.length ? rows[0].length : 1;
  const blank = () => (width <= 1 && mode.collapse ? [''] : new Array(Math.max(1, width)).fill(''));
  const want = rows.map((row) => row.slice());
  const same = (r) => [r, r];
  switch (intent.kind) {
    case 'setCells': {
      const cells = new Map();
      for (const [r, c, v] of intent.cells) cells.set(`${r}:${c}`, [r, c, v.replace(/\r\n|\r|\n/g, eol)]);
      for (const [r, c, v] of cells.values()) {
        if (r >= want.length && v === '') continue;
        while (want.length <= r) want.push(blank());
        const row = want[r];
        if (c >= row.length && v === '') continue;
        while (row.length <= c) row.push('');
        row[c] = v;
      }
      return { want, map: rows.map((_, r) => same(r)) };
    }
    case 'clear': {
      for (const g of intent.ranges) {
        for (let r = g.r0; r <= Math.min(g.r1, want.length - 1); r++) {
          for (let c = g.c0; c <= Math.min(g.c1, want[r].length - 1); c++) want[r][c] = '';
        }
      }
      return { want, map: rows.map((_, r) => same(r)) };
    }
    case 'insertRows': {
      const { at, count } = intent;
      if (at < want.length) want.splice(at, 0, ...Array.from({ length: count }, blank));
      return { want, map: rows.map((_, r) => [r, r < at ? r : r + count]) };
    }
    case 'deleteRows': {
      const { at, count } = intent;
      want.splice(at, count);
      return { want, map: rows.map((_, r) => (r < at ? same(r) : r >= at + count ? [r, r - count] : null)).filter(Boolean) };
    }
    case 'insertCols': {
      const { at, count } = intent;
      const out = want.map((row) => (row.length > at && !isBlankRow(row) ? [...row.slice(0, at), ...new Array(count).fill(''), ...row.slice(at)] : row));
      return { want: out, map: rows.map((_, r) => same(r)) };
    }
    case 'deleteCols': {
      const { at, count } = intent;
      const out = want.map((row) => {
        if (row.length <= at) return row;
        const x = [...row.slice(0, at), ...row.slice(at + count)];
        return x.length ? x : [''];
      });
      return { want: out, map: rows.map((_, r) => same(r)) };
    }
    case 'sort': {
      const { range, key, desc, header } = intent;
      const r0 = Math.max(Math.min(range.r0, range.r1) + (header ? 1 : 0), 1);
      let r1 = Math.min(Math.max(range.r0, range.r1), rows.length - 1);
      while (r1 >= r0 && isBlankRow(rows[r1])) r1--;
      if (r1 - r0 < 1) return { want, map: rows.map((_, r) => same(r)) };
      const keys = [];
      for (let r = r0; r <= r1; r++) {
        const v = rows[r][key] ?? '';
        const t = v.trim();
        keys.push({ r, blank: t === '', num: t === '' ? undefined : gridNumber(v, mode.delimiter), str: t });
      }
      keys.sort((a, b) => {
        if (a.blank !== b.blank) return a.blank ? 1 : -1;
        if (a.blank) return a.r - b.r;
        let d;
        if (a.num !== undefined && b.num !== undefined) d = a.num - b.num;
        else if (a.num !== undefined || b.num !== undefined) d = a.num !== undefined ? -1 : 1;
        else d = fuzzCollator.compare(a.str, b.str);
        return (desc ? -d : d) || a.r - b.r;
      });
      const map = rows.map((_, r) => same(r));
      keys.forEach((k, i) => {
        want[r0 + i] = rows[k.r].slice();
        map[k.r] = [k.r, r0 + i];
      });
      return { want, map };
    }
    case 'replaceAll': {
      const q = intent.query;
      let source = q.regex ? q.text : q.text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
      if (q.wholeCell) source = `^(?:${source})$`;
      const re = new RegExp(source, q.matchCase ? 'gu' : 'giu');
      const rep = intent.replace.replace(/\r\n|\r|\n/g, eol);
      for (const row of want) {
        for (let c = 0; c < row.length; c++) {
          if (row[c] === '') continue;
          const next = q.regex ? row[c].replace(re, rep) : row[c].replace(re, () => rep);
          row[c] = next.replace(/\r\n|\r|\n/g, eol);
        }
      }
      return { want, map: rows.map((_, r) => same(r)) };
    }
    default:
      throw new Error(intent.kind);
  }
}

function randomIntent(R, rows) {
  const count = rows.length;
  const width = count ? rows[0].length : 1;
  const roll = R.rnd();
  const row = () => (R.rnd() < 0.08 ? count + R.ri(3) : R.ri(Math.max(1, count)));
  if (roll < 0.33) {
    const cells = [];
    for (let k = R.rnd() < 0.7 ? 1 : 2 + R.ri(6); k > 0; k--) cells.push([row(), R.ri(width + 2), R.pick(VALUES)]);
    return set(cells);
  }
  if (roll < 0.43) {
    // paste: a rectangle
    const r0 = row();
    const c0 = R.ri(width + 1);
    const cells = [];
    const h = 1 + R.ri(4);
    const w = 1 + R.ri(3);
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) cells.push([r0 + r, c0 + c, R.pick(VALUES)]);
    return set(cells);
  }
  if (roll < 0.52) {
    const range = () => {
      const r0 = R.ri(count + 1);
      const c0 = R.ri(width + 1);
      return { r0, c0, r1: R.rnd() < 0.2 ? 1048575 : r0 + R.ri(4), c1: R.rnd() < 0.2 ? 16383 : c0 + R.ri(3) };
    };
    return R.rnd() < 0.7 ? clearOp(range()) : clearOp(range(), range());
  }
  if (roll < 0.60) return rowsOp('insertRows', R.ri(count + 1), 1 + R.ri(3));
  if (roll < 0.68) return rowsOp('deleteRows', R.ri(Math.max(1, count)), 1 + R.ri(3));
  if (roll < 0.74) return rowsOp('insertCols', R.ri(width + 1), 1 + R.ri(2));
  if (roll < 0.80) return rowsOp('deleteCols', R.ri(width), 1 + R.ri(2));
  if (roll < 0.90) {
    const r0 = R.ri(3);
    return sortOp(r0, R.rnd() < 0.3 ? 1048575 : r0 + R.ri(count + 2), R.ri(width + 1), R.rnd() < 0.5, R.rnd() < 0.2);
  }
  const regex = R.rnd() < 0.4;
  const text = regex ? R.pick(['\\d+', '[aeiou]', '^x$', '(l)(i)', 'o.']) : R.pick(['a', 'x', '1', 'hi', '"', 'line', 'A,B']);
  return replaceOp(text, R.pick(['', 'Q', 'a,b', '$2$1', 'n\nl', '""']), { regex, matchCase: R.rnd() < 0.5, wholeCell: R.rnd() < 0.2 });
}

/**
 * Runs `limit` random intents. Without `start` each round starts from a random text of the mode; with `start`
 * ({ text, eol }, a sample file as document.getText()) every round starts from that file.
 */
function fuzzMode(mode, seed, failures, start, limit = OPS_PER_SEED) {
  const R = makeRandom(seed * 7919 + mode.name.length);
  let ops = 0;
  const pin = { delimiter: mode.delimiter, collapse: mode.collapse };
  let text;
  let eol;
  let idx;
  let policy;
  const reset = () => {
    ({ text, eol } = start ?? randomText(R, mode));
    idx = S.indexCsv(text, pin.delimiter, pin.collapse);
    policy = S.csvQuotePolicy(text, idx); // like CsvDocState: once per pinned document
  };
  reset();
  const fail = (what, extra) => failures.push({ mode: mode.name, seed, op: ops, what, ...extra });
  const maxLength = Math.max(6000, 2 * text.length);
  const maxCount = Math.max(60, 2 * idx.count);
  while (ops < limit) {
    if (text.length > maxLength || idx.count > maxCount) reset();
    ops++;
    const rows = viewRows(text, pin);
    const intent = randomIntent(R, rows);
    const label = JSON.stringify(intent).slice(0, 160);
    let edits;
    try {
      edits = S.planCsvEdit(text, idx, policy, eol, intent);
    } catch (error) {
      if (pin.collapse && (intent.kind === 'insertCols' || intent.kind === 'deleteCols') && error.code === 'spaceColumns') continue;
      fail(`threw: ${error.message}`, { label, text });
      reset();
      continue;
    }
    if (pin.collapse && (intent.kind === 'insertCols' || intent.kind === 'deleteCols')) {
      fail('space-separated column op not blocked', { label });
      continue;
    }
    for (let k = 0; k < edits.length; k++) {
      const e = edits[k];
      if (k && e.start < edits[k - 1].end) fail('edits not sorted / overlapping', { label });
      if (splitsPair(text, e.start) || splitsPair(text, e.end)) fail('edit boundary inside a CRLF / surrogate pair', { label });
    }
    const out = applyEdits(text, edits);
    const nidx = S.indexCsv(out, pin.delimiter, pin.collapse);
    const got = viewRows(out, pin);
    const { want, map } = expectRows(rows, intent, mode, eol);
    if (JSON.stringify(got) !== JSON.stringify(want)) {
      const r = got.findIndex((row, i) => JSON.stringify(row) !== JSON.stringify(want[i]));
      fail('rows differ', { label, row: r, got: JSON.stringify(got[r]), want: JSON.stringify(want[r]), counts: [got.length, want.length], text, out });
      reset();
      continue;
    }
    if (JSON.stringify(indexRows(out, nidx)) !== JSON.stringify(got)) fail('index decode differs from the view after the edit', { label });
    // BOM / sep= line unchanged
    const prefix = idx.count ? text.slice(0, idx.start[0]) : text;
    if (!out.startsWith(prefix)) fail('BOM / sep= line changed', { label });
    // final line break unchanged (except: empty before/after, or an empty last record that needs its line break)
    if (idx.count && nidx.count && nidx.trailingNewline !== idx.trailingNewline && !isBlankRow(got[got.length - 1])) {
      fail('final line break state changed', { label, text, out });
    }
    // untouched records keep their bytes (sort: moved byte for byte)
    for (const pair of map) {
      const [a, b] = pair;
      const moved = intent.kind === 'sort' && a !== b;
      if (!moved && JSON.stringify(want[b]) !== JSON.stringify(rows[a])) continue;
      if (text.slice(idx.start[a], idx.contentEnd[a]) !== out.slice(nidx.start[b], nidx.contentEnd[b])) {
        fail(`record ${a} -> ${b} changed bytes`, { label, before: text.slice(idx.start[a], idx.contentEnd[a]), after: out.slice(nidx.start[b], nidx.contentEnd[b]) });
        break;
      }
    }
    // every line break is the document EOL
    if (eol === '\n' ? out.includes('\r') : /\r(?!\n)|(?<!\r)\n/.test(out)) fail('a line break is not the document EOL', { label });
    // shifting the old index gives the full index of the new text
    const shifted = S.shiftCsvIndex(idx, edits, out);
    if (!sameIndex(shifted, nidx)) fail('shiftCsvIndex differs from indexCsv', { label });
    text = out;
    idx = shifted;
  }
  return ops;
}

for (const mode of MODES) {
  test(`seeded fuzz (${mode.name}): ${SEEDS.length} seeds x ${OPS_PER_SEED} ops, expected rows, untouched bytes, EOL, index shift`, () => {
    const failures = [];
    let ops = 0;
    for (const seed of SEEDS) ops += fuzzMode(mode, seed, failures);
    assert.equal(ops, SEEDS.length * OPS_PER_SEED);
    assert.equal(failures.length, 0, `${failures.length} failures, first: ${JSON.stringify(failures.slice(0, 3), null, 1).slice(0, 3000)}`);
  });
}

// The prototype's per-file run: random intents on every sample, csv-edit fixture and generated case (their real
// quirks: Excel quoting, sep= line, double BOM, CR-only and mixed line breaks after VS Code unified them, ...).
test('seeded fuzz on every sample, fixture and generated case: 300 ops each', () => {
  const failures = [];
  let files = 0;
  for (const [name, bytes] of allCases()) {
    const start = documentText(bytes);
    const pin = pinOf(start.text, name);
    const mode = { name, delimiter: pin.delimiter, collapse: pin.collapse };
    const idx = S.indexCsv(start.text, pin.delimiter, pin.collapse);
    if (idx.unclosedQuoteRecord >= 0) {
      // every edit kind is blocked (nothing is planned)
      assert.throws(() => S.planCsvEdit(start.text, idx, S.csvQuotePolicy(start.text, idx), start.eol, set([[0, 0, 'x']])), S.CsvEditBlocked, name);
      continue;
    }
    fuzzMode(mode, 7, failures, start, 300);
    files++;
  }
  assert.ok(files >= 37, `only ${files} files fuzzed`);
  assert.equal(failures.length, 0, `${failures.length} failures, first: ${JSON.stringify(failures.slice(0, 3), null, 1).slice(0, 3000)}`);
});


// ===== FIX ROUND 1 =====

test('slow regex: too many repeating parts are refused; long values are refused (Replace All) before they can hang', () => {
  const q = (text, o = {}) => ({ scope: 'csv', text, matchCase: false, regex: true, wholeCell: false, lookIn: 'values', ...o });
  for (const p of ['a*a*a*a*a*b', '.*.*.*.*x+', 'a+a+a+a+a{2,}']) assert.ok('invalid' in S.csvFindPattern(q(p)), p);
  assert.equal(S.csvFindPattern(q('a*a*a*b')).maxCell, 100);
  assert.equal(S.csvFindPattern(q('\d+-\d+')).maxCell, 300);
  assert.equal(S.csvFindPattern(q('[*+]x\*')).maxCell, 2000, 'a class and an escape are not repeats');
  assert.equal(S.csvFindPattern(q('a*a*a*b', { regex: false })).maxCell, Infinity, 'plain text has no repeats');
  const text = `k\n${'a'.repeat(30000)}c\n`;
  const t0 = Date.now();
  assert.throws(() => edit(text, replaceOp('a*a*a*b', 'x', { regex: true })), /too long for this regular expression.*nothing was changed/);
  assert.ok(Date.now() - t0 < 2000, 'refused at once');
  assert.equal(edit(text, replaceOp('a', 'b')).edits.length, 1, 'plain text still works on a long value');
});

test("a value 'sep=X' in the first record would become Excel's separator line: blocked", () => {
  const blocked = (fn) => assert.throws(fn, (e) => e instanceof S.CsvEditBlocked && e.code === 'sepLine');
  blocked(() => edit('x\n1\n2\n', set([[0, 0, 'sep=;']])));
  blocked(() => edit('a\n1\n', set([[0, 0, 'sep=,']])));
  blocked(() => edit('xx\n1\n', replaceOp('xx', 'sep=,')));
  blocked(() => edit('h\nsep=,\na\n', rowsOp('deleteRows', 0, 1)));
  assert.equal(edit('x\n1\n2\n', set([[1, 0, 'sep=;']])).out, 'x\nsep=;\n2\n', 'not the first record: fine');
  assert.equal(edit('x,y\n1,2\n', set([[0, 0, 'sep=;']])).out, 'sep=;,y\n1,2\n', 'two fields: not a separator line');
  assert.equal(edit('sep=;\nx;y\n1;2\n', set([[0, 0, 'z']]), { delimiter: ';' }).out, 'sep=;\nz;y\n1;2\n', 'a file that already has the line is edited as before');
});


// ===== FIX ROUND 2 =====

test('R2 exponential regex: a counted repeat of a group with an optional part or an empty alternative is refused', () => {
  const q = (text) => ({ scope: 'csv', text, matchCase: false, regex: true, wholeCell: false, lookIn: 'values' });
  for (const p of ['(a|b?){24}c', '(a?){24}c', '(a|){24}c', '(a?)+c', '(a|b?)*c']) {
    const t0 = Date.now();
    assert.ok('invalid' in S.csvFindPattern(q(p)), p);
    assert.ok(Date.now() - t0 < 500);
  }
  for (const p of ['(Mr|Ms){1}x', '(\d{3})-(\d{4})', 'https?://x', '(ab)+', '(x)?y+']) assert.ok('re' in S.csvFindPattern(q(p)), p);
});

test('R2 a pattern that is valid without the u flag still works (\- \@ \_ literal {)', () => {
  const q = (text) => ({ scope: 'csv', text, matchCase: false, regex: true, wholeCell: false, lookIn: 'values' });
  for (const p of ['\-', '\@', '\#', '\_', 'a{', '{x}']) assert.ok('re' in S.csvFindPattern(q(p)), p);
  assert.equal(edit('k\na-b\n', replaceOp('\-', '+', { regex: true })).out, 'k\na+b\n');
  assert.ok('invalid' in S.csvFindPattern(q('(')));
});

test('R2 sort A to Z puts accented letters with their base letter', () => {
  const out = edit('n\nZoe\nÉmile\nAdam\néa\nzed\n', sortOp(1, 5, 0)).out;
  assert.equal(out, 'n\nAdam\néa\nÉmile\nzed\nZoe\n');
});

test('R2 a value with trailing spaces or tabs is quoted (VS Code trims trailing whitespace on save)', () => {
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, 'pad  ']])).out, 'a,b\n1,"pad  "\n');
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, 'pad\t']])).out, 'a,b\n1,"pad\t"\n');
  assert.equal(edit('a,b\n1,2\n', set([[1, 1, ' pad']])).out, 'a,b\n1, pad\n');
});

test('R2 a value that would start the file with U+FEFF is blocked (it would read back as a BOM)', () => {
  const blocked = (fn) => assert.throws(fn, (e) => e instanceof S.CsvEditBlocked);
  blocked(() => edit('a,b\n1,2\n', set([[0, 0, '﻿head']])));
  blocked(() => edit('a,b\n1,2\n', replaceOp('a', '﻿a')));
  assert.equal(edit('a,b\n1,2\n', set([[1, 0, '﻿x']])).out, 'a,b\n﻿x,2\n', 'not at the file start: fine');
});

test('R2 a huge column index or column count is refused at once', () => {
  const t0 = Date.now();
  assert.throws(() => edit('a\n1\n', set([[0, 1e7, 'x']])), (e) => e instanceof S.CsvEditBlocked && e.code === 'tooLarge');
  assert.throws(() => edit('a\n1\n', rowsOp('insertCols', 0, 2e7)), (e) => e instanceof S.CsvEditBlocked && e.code === 'tooLarge');
  assert.ok(Date.now() - t0 < 500);
});
