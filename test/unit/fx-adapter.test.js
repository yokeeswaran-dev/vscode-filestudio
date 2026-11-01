// fx adapter (src/renderers/sheet.ts, ADDRESSES `fx adapter`): formula tokens over @borgar/fx.
// Checks: joining the tokens gives the formula back (every formula in test/samples, tricky formulas, grammar fuzz),
// references are found with exact positions / coordinates / $ flags / prefixes, Excel's quoting rule for sheet names,
// and fxEmitRef keeps the spelling of every unchanged part.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const SAMPLES = path.join(__dirname, '..', 'samples');
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/** join(tokens) === formula, no gaps / overlaps, and the identity emit changes nothing. */
function assertLossless(formula, where) {
  const tokens = S.fxTokenize(formula);
  assert.equal(S.fxJoin(formula, tokens), formula, `join differs: ${where}`);
  let pos = 0;
  for (const t of tokens) {
    assert.equal(t.start, pos, `gap or overlap at ${pos}: ${where}`);
    assert.ok(t.end > t.start, `empty token: ${where}`);
    pos = t.end;
  }
  assert.equal(pos, formula.length, `tokens end early: ${where}`);
  const same = S.fxEmit(formula, tokens, (t) => (t.type === 'REF' ? S.fxEmitRef(formula, t, { area: { ...t.ref }, prefix: t.ref.prefix }) : null));
  assert.equal(same, formula, `identity emit changed the text: ${where}`);
  return tokens;
}

function decodeXml(text) {
  return text.replace(/&(?:#x([0-9a-f]+)|#(\d+)|(amp|lt|gt|quot|apos));/gi, (_, hex, dec, name) => {
    if (hex) return String.fromCodePoint(parseInt(hex, 16));
    if (dec) return String.fromCodePoint(parseInt(dec, 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[name.toLowerCase()];
  });
}

/** Every formula text of an xlsx package: cell <f>, defined names, CF <formula>, DV <formula1/2>. */
async function formulasOf(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const out = [];
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/(worksheets\/sheet\d+|workbook)\.xml$/i.test(name)) continue;
    const xml = await zip.file(name).async('string');
    const re = /<(f|definedName|formula|formula1|formula2)\b[^>]*>([^<]*)<\/\1>/g;
    for (let m; (m = re.exec(xml)); ) if (m[2]) out.push({ text: decodeXml(m[2]), where: `${name} <${m[1]}>` });
  }
  return out;
}

test('lossless on every formula in test/samples (xlsx up to 10 MB; large.xlsx is generated, not in Git)', async () => {
  let files = 0;
  let count = 0;
  let unknown = 0;
  for (const name of fs.readdirSync(SAMPLES)) {
    if (!name.toLowerCase().endsWith('.xlsx')) continue;
    const file = path.join(SAMPLES, name);
    if (fs.statSync(file).size > 10 * 1024 * 1024 && process.env.FS_FX_LARGE !== '1') continue;
    files++;
    for (const f of await formulasOf(fs.readFileSync(file))) {
      count++;
      const tokens = assertLossless(f.text, `${name} ${f.where}: ${f.text}`);
      unknown += tokens.filter((t) => t.type === 'UNKNOWN').length;
    }
  }
  assert.ok(files >= 1 && count >= 10, `expected formulas in the samples (files ${files}, formulas ${count})`);
  assert.equal(unknown, 0, 'no sample formula may contain unreadable text');
});

test('lossless and exact on tricky formulas', () => {
  const cases = [
    ['SUM(A1:B2,Sheet1!$C$3)', ['A1:B2', 'Sheet1!$C$3']],
    ["'My Data'!A1+[1]Sheet1!B2+'[2]x y'!C:C+Sheet1:Sheet3!A1", ["'My Data'!A1", '[1]Sheet1!B2', "'[2]x y'!C:C", 'Sheet1:Sheet3!A1']],
    ['"A1"&"Sheet1!B2"&"a""b"&A1', ['A1']],
    ["'it''s'!B2+'x!y'!C3", ["'it''s'!B2", "'x!y'!C3"]],
    ['_xlfn.XLOOKUP(1,A:A,B:B)+_xlfn._xlws.SORT($2:$5)', ['A:A', 'B:B', '$2:$5']],
    ['LOG10(ABCD1)+XFE1+A1B', []], // a function name and names that look like cells are not references
    ['R1C1+RC1', ['RC1']], // R1C1 notation is off: R1C1 is a name, RC1 is the A1 cell in column RC
    ['{1,2;3,4}+{"a,b",TRUE;#N/A,-1}', []],
    ['SUM( A1 , 2:3 )\n+B$4', ['A1', '2:3', 'B$4']],
    ['A1#+Sheet1!B2#', ['A1#', 'Sheet1!B2#']],
    ['A1:B2 C3:D4', ['A1:B2', 'C3:D4']],
    ['$A$1:$XFD$1048576+XFD:XFD+1048576:1048576', ['$A$1:$XFD$1048576', 'XFD:XFD', '1048576:1048576']],
    ['"unterminated', []],
    ['Sheet1!#REF!+#REF!+A1', ['A1']],
    ['Table1[[#This Row],[a\'[b\']]]+[@x]+T_2[[#All],[A]:[B]]', []],
    ['=IF(A1>0,"x",B2%)', ['A1', 'B2']],
  ];
  for (const [formula, refs] of cases) {
    const tokens = assertLossless(formula, formula);
    const got = tokens.filter((t) => t.type === 'REF').map((t) => formula.slice(t.start, t.end));
    assert.deepEqual(got, refs, formula);
  }
});

test('token kinds: names, tables, external, 3D, spill, qualified #REF!', () => {
  const kinds = (f) => S.fxTokenize(f).map((t) => [t.type, f.slice(t.start, t.end), t.prefix ?? t.ref?.prefix ?? null]);
  const name = kinds('[1]!Rate+Sheet1!Tax+Total');
  assert.deepEqual(name.filter((k) => k[0] === 'NAME'), [
    ['NAME', '[1]!Rate', { quoted: false, book: '[1]', sheet1: null, sheet2: null }],
    ['NAME', 'Sheet1!Tax', { quoted: false, book: null, sheet1: 'Sheet1', sheet2: null }],
    ['NAME', 'Total', null],
  ]);
  const struct = kinds("'[1]Sh'!Table1[Col]+Table1[Col]");
  assert.deepEqual(struct.filter((k) => k[0] === 'STRUCT'), [
    ['STRUCT', "'[1]Sh'!Table1[Col]", { quoted: true, book: '[1]', sheet1: 'Sh', sheet2: null }],
    ['STRUCT', 'Table1[Col]', null],
  ]);
  const err = S.fxTokenize("'a b'!#REF!");
  assert.equal(err.length, 1);
  assert.equal(err[0].type, 'ERR');
  assert.equal(err[0].refError, true);
  assert.deepEqual(err[0].prefix, { quoted: true, book: null, sheet1: 'a b', sheet2: null });
  const [ref3d] = S.fxTokenize('Sheet1:Sheet3!$B$2:C5');
  assert.deepEqual({ ...ref3d.ref, prefix: undefined }, {
    kind: 'area', r0: 1, c0: 1, r1: 4, c1: 2, $r0: true, $c0: true, $r1: false, $c1: false, spill: false, bodyStart: 14, prefix: undefined,
  });
  assert.deepEqual(ref3d.ref.prefix, { quoted: false, book: null, sheet1: 'Sheet1', sheet2: 'Sheet3' });
  const [cols] = S.fxTokenize('$C:E');
  assert.deepEqual([cols.ref.kind, cols.ref.r0, cols.ref.r1, cols.ref.c0, cols.ref.c1, cols.ref.$c0, cols.ref.$c1], ['cols', 0, MAX_ROWS - 1, 2, 4, true, false]);
  const [rows] = S.fxTokenize('3:$7');
  assert.deepEqual([rows.ref.kind, rows.ref.r0, rows.ref.r1, rows.ref.c0, rows.ref.c1, rows.ref.$r0, rows.ref.$r1], ['rows', 2, 6, 0, MAX_COLS - 1, false, true]);
  const [spill] = S.fxTokenize('A1#');
  assert.equal(spill.ref.spill, true);
  assert.equal(spill.end, 3);
});

// Excel's own output for "=<name>!A1" (s2-editmodel E1, excel/out/result.json: 33 of 33 names).
const EXCEL_QUOTES = [
  ['My Data', "'My Data'!A1"], ['Q1', "'Q1'!A1"], ['R1C1', "'R1C1'!A1"], ['TRUE', "'TRUE'!A1"], ['1abc', "'1abc'!A1"],
  ['abc1', "'abc1'!A1"], ['A.B', 'A.B!A1'], ['Données', 'Données!A1'], ['_x', '_x!A1'], ['R', "'R'!A1"], ['C', "'C'!A1"],
  ['RC', "'RC'!A1"], ['ABCD1', 'ABCD1!A1'], ['XFD1', "'XFD1'!A1"], ['XFE1', 'XFE1!A1'], ['Sheet-1', "'Sheet-1'!A1"],
  ['a b', "'a b'!A1"], ['R2', "'R2'!A1"], ['C3', "'C3'!A1"], ['RC1', "'RC1'!A1"], ['R1C', "'R1C'!A1"], ['Sheet1.', 'Sheet1.!A1'],
  ['.x', "'.x'!A1"], ['Ab_1.2', 'Ab_1.2!A1'], ['x!y', "'x!y'!A1"], ['mañana', 'mañana!A1'], ['Σ1', 'Σ1!A1'], ['A1B2', 'A1B2!A1'],
  ['ZZZZ99', 'ZZZZ99!A1'], ['TRUEX', 'TRUEX!A1'], ['FALSE', "'FALSE'!A1"], ['Q1a', 'Q1a!A1'], ['K', 'K!A1'],
];

test("quoting table: sheet names quoted exactly where Excel quotes them (33 names)", () => {
  for (const [name, excel] of EXCEL_QUOTES) {
    assert.equal(S.quoteSheetForRef(name) + '!A1', excel, name);
    assert.equal(S.sheetNeedsQuotes(name), excel.startsWith("'"), name);
    const prefix = { quoted: false, book: null, sheet1: name, sheet2: null };
    assert.equal(S.fxPrefixText(prefix) + 'A1', excel, `prefix ${name}`);
    // and the tokenizer reads Excel's text back to the same sheet name
    const [t] = S.fxTokenize(excel);
    assert.equal(t.type, 'REF', excel);
    assert.equal(t.ref.prefix.sheet1, name, excel);
  }
  assert.equal(S.quoteSheetForRef("it's"), "'it''s'");
  assert.equal(S.fxPrefixText({ quoted: false, book: '[2]', sheet1: 'x y', sheet2: null }), "'[2]x y'!");
  assert.equal(S.fxPrefixText({ quoted: false, book: null, sheet1: 'S1', sheet2: 'My S' }), "'S1:My S'!");
  assert.equal(S.fxPrefixText({ quoted: true, book: null, sheet1: 'Sheet1', sheet2: null }), "'Sheet1'!", 'written quotes are kept');
});

test('fxEmitRef keeps the spelling of every unchanged part', () => {
  const rewrite = (formula, change) =>
    S.fxEmit(formula, S.fxTokenize(formula), (t) => (t.type === 'REF' ? S.fxEmitRef(formula, t, change(t.ref)) : null));
  // one end changes: the other keeps its lower-case / $ spelling, the quoted prefix is kept
  assert.equal(rewrite("'Sheet1'!a1:b5", (r) => ({ area: { ...r, r1: r.r1 + 1 } })), "'Sheet1'!a1:B6");
  // prefix changes (rename): the body keeps its spelling, quotes follow Excel's rule
  assert.equal(rewrite('Sheet1!$a$1+Sheet1!b2', (r) => ({ prefix: { ...r.prefix, quoted: false, sheet1: 'New Name' } })), "'New Name'!$a$1+'New Name'!b2");
  assert.equal(rewrite("'Old one'!A1", (r) => ({ prefix: { ...r.prefix, quoted: false, sheet1: 'Data' } })), 'Data!A1');
  // nothing changes: a reversed range keeps its text
  assert.equal(rewrite('B5:A1', (r) => ({ area: { ...r }, prefix: r.prefix })), 'B5:A1');
  // a reversed range that changes is written normalized (same meaning)
  assert.equal(rewrite('B5:A1', (r) => ({ area: { ...r, r1: 9 } })), 'A1:B10');
  // whole rows / columns, spill, prefix removed, 3D
  assert.equal(rewrite('$2:7', (r) => ({ area: { ...r, r1: 7 } })), '$2:8');
  assert.equal(rewrite('c:$E', (r) => ({ area: { ...r, c0: 3 } })), 'D:$E');
  assert.equal(rewrite('A1#', (r) => ({ area: { ...r, r0: 1, r1: 1 } })), 'A2#');
  assert.equal(rewrite('Sheet1!A1', () => ({ prefix: null })), 'A1');
  assert.equal(rewrite('Jan:Mar!A1', (r) => ({ prefix: { ...r.prefix, sheet2: 'Q 4' } })), "'Jan:Q 4'!A1");
  // the same sheet named differently (case) counts as a change of the text only when the name differs
  assert.equal(rewrite("'Sheet1'!A1", (r) => ({ prefix: { ...r.prefix } })), "'Sheet1'!A1");
});

// ----- grammar fuzz (port of ROUND\s2-editmodel\proto\gen-fuzz.js, E6) -----

function colName(i) {
  let n = i + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function bodyText(a) {
  const col = (c, abs) => (abs ? '$' : '') + colName(c);
  const row = (r, abs) => (abs ? '$' : '') + (r + 1);
  if (a.kind === 'cols') return `${col(a.c0, a.$c0)}:${col(a.c1, a.$c1)}`;
  if (a.kind === 'rows') return `${row(a.r0, a.$r0)}:${row(a.r1, a.$r1)}`;
  const first = col(a.c0, a.$c0) + row(a.r0, a.$r0);
  return a.kind === 'cell' ? first : `${first}:${col(a.c1, a.$c1)}${row(a.r1, a.$r1)}`;
}

function makeGenerator(seed) {
  let s = seed;
  const rnd = () => {
    s = (s * 1103515245 + 12345) & 0x7fffffff;
    return s / 0x7fffffff;
  };
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const chance = (p) => rnd() < p;
  const SHEETS_PLAIN = ['Sheet1', 'Data', 'A.B', 'Données', '_x', 'ABCD1', 'Q1a', 'K', 'Sheet1.', 'mañana', 'Σ1', 'XFE1'];
  const SHEETS_QUOTED = ['My Data', "it's", 'Q1', 'R1C1', 'TRUE', '1abc', 'x!y', 'a,b', 'a(b)', "''", 'Sales 2024', 'a"b', 'a;b', 'R', 'C3', '#x', 'a{b}', 'a-b', 'a+b'];
  const FUNCS = ['SUM', 'LOG10', '_xlfn.XLOOKUP', 'ATAN2', 'DAYS360', '_xlfn.F.DIST.RT', 'IF', 'INDEX', '_xlfn._xlws.SORT', 'OFFSET', 'T', 'N'];
  const NAMES = ['Tax_Rate', 'Q1_2', 'LOG10x', 'ABCD12', 'XFE5', '_xlpm.x', 'Total.Sales', 'rate', 'A1B', 'x\\y'];
  const STRINGS = ['"A1"', '"Sheet1!B2"', '"a""b"', '""', '"\'Q1\'!A1"', '"[1]x!A1"', '"{1}"', '"#REF!"'];
  const STRUCTS = ['Table1[Col]', "Table1[[#This Row],[a'[b']]]", '[@x]', 'T_2[[#All],[A]:[B]]', 'Tbl[#Totals]'];
  const ARRAYS = ['{1,2;3,4}', '{"a,b",TRUE;#N/A,-1}', '{1}'];
  const ERRS = ['#N/A', '#REF!', '#DIV/0!', '#VALUE!', '#NAME?'];
  const NUMS = ['1', '1.5', '.5', '1E+10', '2.5e-3', '0', '100'];

  function genBody() {
    const k = pick(['cell', 'cell', 'area', 'cols', 'rows']);
    const r = () => (chance(0.1) ? (chance(0.5) ? MAX_ROWS - 1 : 0) : Math.floor(rnd() * 2000));
    const c = () => (chance(0.1) ? (chance(0.5) ? MAX_COLS - 1 : 0) : Math.floor(rnd() * 800));
    const d = () => chance(0.3);
    const a = { kind: k, $r0: d(), $c0: d(), $r1: d(), $c1: d() };
    if (k === 'cell') {
      a.r0 = a.r1 = r();
      a.c0 = a.c1 = c();
      a.$r1 = a.$r0;
      a.$c1 = a.$c0;
    } else if (k === 'area') {
      a.r0 = r(); a.r1 = r(); a.c0 = c(); a.c1 = c();
    } else if (k === 'cols') {
      a.c0 = c(); a.c1 = c(); a.r0 = 0; a.r1 = MAX_ROWS - 1; a.$r0 = a.$r1 = false;
    } else {
      a.r0 = r(); a.r1 = r(); a.c0 = 0; a.c1 = MAX_COLS - 1; a.$c0 = a.$c1 = false;
    }
    // Excel always writes ranges top-left:bottom-right
    if (a.r0 > a.r1) [a.r0, a.r1, a.$r0, a.$r1] = [a.r1, a.r0, a.$r1, a.$r0];
    if (a.c0 > a.c1) [a.c0, a.c1, a.$c0, a.$c1] = [a.c1, a.c0, a.$c1, a.$c0];
    return { area: a, text: bodyText(a) };
  }
  function genPrefix() {
    const x = rnd();
    if (x < 0.35) return { text: '', p: null };
    if (x < 0.55) { const sh = pick(SHEETS_PLAIN); return { text: sh + '!', p: { sheet1: sh } }; }
    if (x < 0.8) { const sh = pick(SHEETS_QUOTED); return { text: `'${sh.replace(/'/g, "''")}'!`, p: { sheet1: sh } }; }
    if (x < 0.88) {
      const a = pick(SHEETS_PLAIN);
      let b = pick(SHEETS_PLAIN);
      while (b === a) b = pick(SHEETS_PLAIN);
      return { text: `${a}:${b}!`, p: { sheet1: a, sheet2: b } };
    }
    const n = 1 + Math.floor(rnd() * 9);
    if (x < 0.94) { const sh = pick(SHEETS_PLAIN); return { text: `[${n}]${sh}!`, p: { book: `[${n}]`, sheet1: sh } }; }
    const sh = pick(SHEETS_QUOTED);
    return { text: `'[${n}]${sh.replace(/'/g, "''")}'!`, p: { book: `[${n}]`, sheet1: sh } };
  }
  function gen(depth, refs, parts) {
    const ws = () => (chance(0.1) ? ' ' : '');
    const term = () => {
      const x = rnd();
      if (x < 0.4 || depth > 3) {
        const pre = genPrefix();
        const body = genBody();
        const text = pre.text + body.text;
        refs.push({ text, ...body.area, prefix: pre.p });
        parts.push(text);
        return;
      }
      if (x < 0.5) return void parts.push(pick(NUMS));
      if (x < 0.57) return void parts.push(pick(STRINGS));
      if (x < 0.62) return void parts.push(pick(['TRUE', 'FALSE']));
      if (x < 0.66) return void parts.push(pick(ERRS));
      if (x < 0.71) return void parts.push(chance(0.3) ? `[1]!${pick(NAMES).replace('_xlpm.', '')}` : pick(NAMES));
      if (x < 0.75) return void parts.push(pick(STRUCTS));
      if (x < 0.78) return void parts.push(pick(ARRAYS));
      if (x < 0.82) {
        parts.push('(');
        gen(depth + 1, refs, parts);
        parts.push(')');
        return;
      }
      if (x < 0.85) {
        parts.push('-');
        term();
        return;
      }
      parts.push(pick(FUNCS) + '(');
      const n = Math.floor(rnd() * 3);
      for (let i = 0; i <= n; i++) {
        if (i) parts.push(',' + ws());
        gen(depth + 1, refs, parts);
      }
      parts.push(')');
    };
    term();
    const k = Math.floor(rnd() * 3);
    for (let i = 0; i < k; i++) {
      const op = pick(['+', '-', '*', '/', '&', '^', '=', '<>', '<=', '+', '*']);
      parts.push(ws() + op + ws());
      term();
    }
    if (chance(0.05)) parts.push('%');
  }
  return () => {
    const refs = [];
    const parts = [];
    gen(0, refs, parts);
    return { formula: parts.join(''), refs };
  };
}

const FUZZ_SEEDS = [1, 7, 20261016];
const FUZZ_COUNT = 10_000;

test(`grammar fuzz: ${FUZZ_SEEDS.length} seeds x ${FUZZ_COUNT} formulas find exactly the generated references`, () => {
  const failures = [];
  let refsSeen = 0;
  for (const seed of FUZZ_SEEDS) {
    const next = makeGenerator(seed);
    for (let i = 0; i < FUZZ_COUNT; i++) {
      const { formula, refs } = next();
      const tokens = S.fxTokenize(formula);
      const problem = (() => {
        if (S.fxJoin(formula, tokens) !== formula) return 'join';
        const got = tokens.filter((t) => t.type === 'REF');
        if (got.length !== refs.length) return `found ${got.length} refs, expected ${refs.length}`;
        for (let j = 0; j < refs.length; j++) {
          const g = got[j].ref;
          const e = refs[j];
          if (formula.slice(got[j].start, got[j].end) !== e.text) return `ref ${j} text`;
          const sameArea = g.kind === e.kind && g.r0 === e.r0 && g.r1 === e.r1 && g.c0 === e.c0 && g.c1 === e.c1 &&
            (e.kind === 'cols' || (g.$r0 === e.$r0 && g.$r1 === e.$r1)) && (e.kind === 'rows' || (g.$c0 === e.$c0 && g.$c1 === e.$c1));
          if (!sameArea) return `ref ${j} area`;
          const gp = g.prefix || {};
          const ep = e.prefix || {};
          if ((gp.sheet1 || null) !== (ep.sheet1 || null) || (gp.sheet2 || null) !== (ep.sheet2 || null) || (gp.book || null) !== (ep.book || null)) {
            return `ref ${j} prefix`;
          }
        }
        // identity emit changes nothing
        const same = S.fxEmit(formula, tokens, (t) => (t.type === 'REF' ? S.fxEmitRef(formula, t, { area: { ...t.ref }, prefix: t.ref.prefix }) : null));
        if (same !== formula) return 'identity emit';
        // moving every reference one row down (where possible) changes only reference text
        const moved = S.fxEmit(formula, tokens, (t) => {
          if (t.type !== 'REF' || t.ref.r1 >= MAX_ROWS - 1 || t.ref.kind === 'cols') return null;
          return S.fxEmitRef(formula, t, { area: { ...t.ref, r0: t.ref.r0 + 1, r1: t.ref.r1 + 1 } });
        });
        const outside = (text) => {
          const ts = S.fxTokenize(text);
          let rest = '';
          let last = 0;
          for (const t of ts) {
            if (t.type !== 'REF') continue;
            rest += text.slice(last, t.start) + '\u0000';
            last = t.end;
          }
          return rest + text.slice(last);
        };
        if (outside(moved) !== outside(formula)) return 'non-reference text changed by a rewrite';
        const movedRefs = S.fxTokenize(moved).filter((t) => t.type === 'REF');
        for (let j = 0; j < got.length; j++) {
          const before = got[j].ref;
          const after = movedRefs[j]?.ref;
          const shift = before.r1 < MAX_ROWS - 1 && before.kind !== 'cols' ? 1 : 0;
          if (!after || after.r0 !== before.r0 + shift || after.r1 !== before.r1 + shift || after.c0 !== before.c0 || after.c1 !== before.c1) {
            return `rewritten ref ${j} reads back wrong`;
          }
        }
        return null;
      })();
      refsSeen += refs.length;
      if (problem && failures.length < 8) failures.push(`seed ${seed} #${i}: ${problem}: ${formula}`);
    }
  }
  assert.deepEqual(failures, []);
  assert.ok(refsSeen > FUZZ_SEEDS.length * FUZZ_COUNT, `the fuzz should generate references (${refsSeen})`);
});

test('fxTokenize never throws and stays lossless on junk', () => {
  const junk = ["'unclosed", 'A1:', ':B2', '[1', 'SUM((', ')))', '#', "!'", 'A1:XFE2', '\u0000￿', '😀+A1', '=', '==A1'];
  for (const text of junk) assertLossless(text, JSON.stringify(text));
  assert.deepEqual(S.fxTokenize(''), []);
});

// ===== length cap (fix round 1: @borgar/fx is quadratic on apostrophe runs) =====

test('fxTokenize: a formula over 8192 characters is one UNKNOWN token, fast (no hang on a crafted file)', () => {
  const evil = '=' + "'".repeat(65_536);
  const t0 = Date.now();
  const tokens = S.fxTokenize(evil);
  const ms = Date.now() - t0;
  assert.deepEqual(tokens, [{ type: 'UNKNOWN', start: 0, end: evil.length }]);
  assert.ok(ms < 500, `took ${ms} ms`);
  assert.equal(S.fxJoin(evil, tokens), evil);
  // Excel's own limit (8192 characters, including the '=') is still tokenized.
  const longRefs = '=' + Array.from({ length: 1000 }, () => 'A1').join('+');
  assert.ok(longRefs.length <= 8192);
  assert.equal(S.fxTokenize(longRefs).filter((t) => t.type === 'REF').length, 1000);
  assert.ok(S.fxTokenize('=' + 'A1+'.repeat(2700) + '1').length > 1);
  const t1 = Date.now();
  S.fxTokenize('=' + "'".repeat(8191));
  assert.ok(Date.now() - t1 < 2000, 'a formula at the cap is still fast');
});

// ===== fix round 2 =====

test('fxTokenize: many quote-heavy formulas share one time budget (no freeze across a whole workbook)', () => {
  const shape = "'".repeat(10) + '+'; // costs tens of ms per 8000-character formula in @borgar/fx
  const evil = '=' + shape.repeat(Math.floor(8000 / shape.length));
  const t0 = Date.now();
  for (let i = 0; i < 100; i++) S.fxTokenize(evil);
  const ms = Date.now() - t0;
  assert.ok(ms < 2500, `100 hostile formulas took ${ms} ms`);
  // a normal formula is still read after the budget ran out
  assert.equal(S.fxTokenize("'My Data'!A1+B2").filter((x) => x.type === 'REF').length, 2);
});

test('fxTokenize: a long run of apostrophes is one UNKNOWN token at once', () => {
  const evil = '=' + "'".repeat(7000);
  const t0 = Date.now();
  const tokens = S.fxTokenize(evil);
  assert.ok(Date.now() - t0 < 30, 'run rejected before fx runs');
  assert.deepEqual(tokens, [{ type: 'UNKNOWN', start: 0, end: evil.length }]);
  const dots = '=' + '1.'.repeat(3500);
  assert.deepEqual(S.fxTokenize(dots), [{ type: 'UNKNOWN', start: 0, end: dots.length }]);
});

test('sheetNeedsQuotes: names with non-ASCII decimal digits are quoted (Excel 16 does)', () => {
  assert.equal(S.sheetNeedsQuotes('x١'), true);
  assert.equal(S.sheetNeedsQuotes('x٤'), true);
  assert.equal(S.quoteSheetForRef('x١'), "'x١'");
  assert.equal(S.sheetNeedsQuotes('ab_1'), false);
  assert.equal(S.sheetNeedsQuotes('Données'), false);
});

test('fxEmitRef: a coordinate outside the grid, NaN or an empty sheet name throws (never writes bad formula text)', () => {
  const f = 'SUM(A1:B2)+Sheet1!C3';
  const toks = S.fxTokenize(f).filter((t) => t.type === 'REF');
  const emit = (i, patch) => S.fxEmitRef(f, toks[i], { area: { ...toks[i].ref, ...patch } });
  for (const patch of [{ r0: -1 }, { r0: 1048576 }, { c0: 16384 }, { c0: -1 }, { r0: NaN }, { r1: 1.5 }, { c1: Infinity }]) {
    assert.throws(() => emit(0, patch), RangeError, JSON.stringify(patch));
  }
  assert.throws(() => S.fxEmitRef(f, toks[1], { prefix: { quoted: false, book: null, sheet1: '', sheet2: null } }), RangeError);
  assert.throws(() => S.fxPrefixText({ quoted: false, book: null, sheet1: '', sheet2: null }), RangeError);
  // edges of the grid are fine, and a swapped corner pair is written normalized
  assert.equal(emit(0, { r0: 0, r1: 1048575, c0: 0, c1: 16383 }), 'A1:XFD1048576');
  assert.equal(emit(0, { r0: 5, r1: 2 }), 'A3:B6');
});

test('fxTokenize: a long run of dots, apostrophes or [ inside a string literal does not make the formula opaque', () => {
  for (const ch of ['.', "'", '[']) {
    const f = `SUM(A1:B2)&"${ch.repeat(200)}"`;
    const toks = S.fxTokenize(f);
    assert.equal(toks.map((t) => f.slice(t.start, t.end)).join(''), f);
    assert.ok(toks.some((t) => t.type === 'REF'), `ref found with a run of ${ch}`);
    assert.ok(!toks.some((t) => t.type === 'UNKNOWN'), `no UNKNOWN with a run of ${ch}`);
  }
  // a hostile run outside a string is still refused
  assert.equal(S.fxTokenize('A1+' + '.'.repeat(100))[0].type, 'UNKNOWN');
});
