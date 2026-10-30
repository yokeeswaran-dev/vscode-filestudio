// Markdown WYSIWYG edit blocks (src/renderers/markdown.ts, section EDIT BLOCKS): markdownBlocks, applyMdTextEdits,
// verifyMdEdit. Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { markdownBlocks, applyMdTextEdits, verifyMdEdit } = require('../.out/markdown.js');

const ROOT = path.join(__dirname, '..', '..');

/** Blocks as [kind, text, fixed] so a test reads like the document. */
const view = (src, opts) => markdownBlocks(src, opts).map((b) => [b.kind, src.slice(b.start, b.end), b.fixed]);

/** An edit of `src` that replaces [start, end) with `text` (oldText taken from `src`). */
const edit = (src, start, end, text) => ({ start, end, oldText: src.slice(start, end), text });

/** An edit that replaces the first `find` in `src` with `text`. */
const replace = (src, find, text) => {
  const at = src.indexOf(find);
  assert.ok(at >= 0, `not found: ${find}`);
  return edit(src, at, at + find.length, text);
};

/** Applies `edits` and runs the host check on the result. */
const check = (src, edits) => verifyMdEdit(src, edits, applyMdTextEdits(src, edits));

// ----- markdownBlocks: offsets, lines, kinds -----

test('markdownBlocks: LF offsets, lines and kinds; blank lines belong to no block', () => {
  const src = '# Title\n\nSome *text*\nsecond line\n\n- a\n- b\n\n```js\nx = 1;\n```\n\n---\n';
  assert.deepEqual(markdownBlocks(src), [
    { start: 0, end: 7, line: 0, endLine: 0, kind: 'heading', fixed: false },
    { start: 9, end: 32, line: 2, endLine: 3, kind: 'paragraph', fixed: false },
    { start: 34, end: 41, line: 5, endLine: 6, kind: 'bullet_list', fixed: false },
    { start: 43, end: 59, line: 8, endLine: 10, kind: 'fence', fixed: false },
    { start: 61, end: 64, line: 12, endLine: 12, kind: 'hr', fixed: false },
  ]);
  for (const b of markdownBlocks(src)) assert.ok(!src.slice(b.start, b.end).endsWith('\n'));
});

test('markdownBlocks: CRLF and lone CR give the same blocks; offsets exclude the line break', () => {
  const lf = '# Title\n\nSome text\nmore\n\n1. one\n2. two\n';
  const lfBlocks = markdownBlocks(lf);
  for (const eol of ['\r\n', '\r']) {
    const src = lf.replace(/\n/g, eol);
    const blocks = markdownBlocks(src);
    assert.deepEqual(blocks.map((b) => [b.kind, b.line, b.endLine, b.fixed]), lfBlocks.map((b) => [b.kind, b.line, b.endLine, b.fixed]), JSON.stringify(eol));
    assert.deepEqual(blocks.map((b) => src.slice(b.start, b.end)), lfBlocks.map((b) => lf.slice(b.start, b.end).replace(/\n/g, eol)));
    for (const b of blocks) assert.ok(!/[\r\n]$/.test(src.slice(b.start, b.end)), `block on line ${b.line} ends with a line break`);
  }
  assert.deepEqual(markdownBlocks('# T\r\n\r\npara\r\n')[1], { start: 7, end: 11, line: 2, endLine: 2, kind: 'paragraph', fixed: false });
  assert.deepEqual(markdownBlocks('# T\r\rpara\r')[1], { start: 5, end: 9, line: 2, endLine: 2, kind: 'paragraph', fixed: false });
});

test('markdownBlocks: a BOM is counted in the offsets but not in the lines', () => {
  const src = '﻿# T\n\nabc';
  assert.deepEqual(markdownBlocks(src), [
    { start: 1, end: 4, line: 0, endLine: 0, kind: 'heading', fixed: false },
    { start: 6, end: 9, line: 2, endLine: 2, kind: 'paragraph', fixed: false },
  ]);
});

test('markdownBlocks: no final newline, tabs and trailing spaces stay inside the block text', () => {
  assert.deepEqual(view('para one\n\nlast line'), [['paragraph', 'para one', false], ['paragraph', 'last line', false]]);
  const src = '- a  \n\tb\n\npara\t \n\n\t\n';
  assert.deepEqual(view(src), [['bullet_list', '- a  \n\tb', false], ['paragraph', 'para\t ', false]]);
  assert.deepEqual(markdownBlocks(''), []);
  assert.deepEqual(markdownBlocks('\n\n  \n'), []);
});

test('markdownBlocks: text the preview does not show becomes a fixed gap block', () => {
  const src = 'see [x][r]\n\n[r]: https://example.com\n[s]: https://example.org\n\n# Next\n';
  assert.deepEqual(view(src), [
    ['paragraph', 'see [x][r]', false],
    ['gap', '[r]: https://example.com\n[s]: https://example.org', true],
    ['heading', '# Next', false],
  ]);
  const gap = markdownBlocks(src)[1];
  assert.match(gap.reason, /link reference definitions/);
});

test('markdownBlocks: footnote definitions are fixed blocks one by one; an unused one is a gap', () => {
  const src = 'Text[^1] and[^n].\n\n[^1]: one\n    more\n\n[^n]: two\n\n[^u]: unused\n';
  assert.deepEqual(view(src), [
    ['paragraph', 'Text[^1] and[^n].', false],
    ['footnote', '[^1]: one\n    more', true],
    ['footnote', '[^n]: two', true],
    ['gap', '[^u]: unused', true],
  ]);
  // a footnote definition inside a list overlaps the list: one fixed block
  assert.deepEqual(view('- a[^1]\n\n  [^1]: x\n'), [['overlap', '- a[^1]\n\n  [^1]: x', true]]);
});

test('markdownBlocks: alerts, front matter, $$ math, HTML and indented code are fixed; quotes and fences are not', () => {
  assert.deepEqual(view('---\ntitle: x\n---\n\n# H\n'), [['front_matter', '---\ntitle: x\n---', true], ['heading', '# H', false]]);
  assert.deepEqual(view('> [!NOTE]\n> hi\n\n> quote\n'), [['alert', '> [!NOTE]\n> hi', true], ['blockquote', '> quote', false]]);
  assert.deepEqual(view('$$\nx^2\n$$\n\npara\n'), [['math_block', '$$\nx^2\n$$', true], ['paragraph', 'para', false]]);
  assert.deepEqual(view('<div>\nhi\n</div>\n\npara\n'), [['html_block', '<div>\nhi\n</div>', true], ['paragraph', 'para', false]]);
  assert.deepEqual(view('para\n\n    code\n\n```js\nx\n```\n'), [
    ['paragraph', 'para', false],
    ['code_block', '    code', true],
    ['fence', '```js\nx\n```', false],
  ]);
  for (const b of markdownBlocks('---\ntitle: x\n---\n\n> [!WARNING]\n> w\n')) assert.equal(typeof b.reason, 'string');
});

test('markdownBlocks: contextCheck fixes blocks that render differently alone', () => {
  const src = 'see [x][r] here\n\nnote[^1]\n\n# Same\n\n# Same\n\nplain **bold**\n\n[r]: https://example.com\n\n[^1]: the note\n';
  const plain = markdownBlocks(src);
  const ctx = markdownBlocks(src, { contextCheck: true });
  assert.deepEqual(ctx.map((b) => [b.start, b.end, b.kind]), plain.map((b) => [b.start, b.end, b.kind]));
  const fixedOf = (blocks) => blocks.map((b) => [src.slice(b.start, b.end), b.fixed]);
  assert.deepEqual(fixedOf(ctx), [
    ['see [x][r] here', true], // reference link
    ['note[^1]', true], // footnote reference
    ['# Same', false], // duplicate heading slugs do not count (ids are position data)
    ['# Same', false],
    ['plain **bold**', false],
    ['[r]: https://example.com', true],
    ['[^1]: the note', true],
  ]);
  assert.equal(ctx[0].reason, 'depends on other parts of the document');
  assert.equal(fixedOf(plain)[0][1], false);
});

// ----- markdownBlocks on real documents -----

/** REPO Markdown files with LF line breaks (the working copy may have CRLF). */
const realDocs = ['test/samples/kitchen-sink.md', 'test/samples/other.md', 'README.md', 'CHANGELOG.md', 'CONTRIBUTING.md']
  .filter((f) => fs.existsSync(path.join(ROOT, f)))
  .map((f) => [f, fs.readFileSync(path.join(ROOT, f), 'utf8').replace(/\r\n?/g, '\n')]);

test('markdownBlocks on real documents: ordered, on line bounds, every non-blank line in exactly one block', () => {
  assert.ok(realDocs.length >= 3);
  for (const [name, lf] of realDocs) {
    const variants = { LF: lf, CRLF: lf.replace(/\n/g, '\r\n'), BOM: `﻿${lf}`, 'no final newline': lf.replace(/\n+$/, '') };
    const base = markdownBlocks(lf, { contextCheck: true });
    for (const [label, src] of Object.entries(variants)) {
      const blocks = markdownBlocks(src, { contextCheck: true });
      const where = `${name} (${label})`;
      assert.deepEqual(blocks.map((b) => [b.kind, b.line, b.endLine, b.fixed]), base.map((b) => [b.kind, b.line, b.endLine, b.fixed]), where);
      const lines = (src.charCodeAt(0) === 0xfeff ? src.slice(1) : src).split(/\r\n|\n|\r/);
      const owner = new Array(lines.length).fill(-1);
      let prevEnd = -1;
      blocks.forEach((b, i) => {
        assert.ok(b.start > prevEnd && b.start < b.end && b.end <= src.length, `${where}: block ${i} bounds`);
        prevEnd = b.end;
        const text = src.slice(b.start, b.end);
        assert.equal(text.split(/\r\n|\n|\r/).join('\n'), lines.slice(b.line, b.endLine + 1).join('\n'), `${where}: block ${i} text = its lines`);
        assert.ok(!/^[ \t]*$/.test(lines[b.line]) && !/^[ \t]*$/.test(lines[b.endLine]), `${where}: block ${i} starts/ends on a blank line`);
        for (let l = b.line; l <= b.endLine; l++) owner[l] = i;
      });
      lines.forEach((line, l) => {
        if (!/^[ \t]*$/.test(line)) assert.ok(owner[l] >= 0, `${where}: line ${l} is in no block: ${line}`);
      });
      // a no-op check passes on every document
      assert.deepEqual(verifyMdEdit(src, [], src), { ok: true, bad: [] }, where);
    }
  }
});

test('markdownBlocks: kitchen-sink.md has every fixed kind and stays fast at the WYSIWYG limit', () => {
  const ks = realDocs.find(([n]) => n.endsWith('kitchen-sink.md'))[1];
  const kinds = new Set(markdownBlocks(ks).map((b) => b.kind));
  for (const k of ['front_matter', 'heading', 'paragraph', 'bullet_list', 'ordered_list', 'fence', 'table', 'alert', 'math_block', 'html_block', 'footnote']) {
    assert.ok(kinds.has(k), `kitchen-sink.md has no ${k} block`);
  }
  const parts = [];
  for (let i = 0; i < 1000; i++) parts.push(`## Heading ${i}`, `Paragraph ${i} with **bold** and a [link](https://example.com/${i}).`, `- item ${i}\n- [ ] task`);
  const big = `${parts.join('\n\n')}\n`;
  const t0 = Date.now();
  const blocks = markdownBlocks(big, { contextCheck: true });
  const ms = Date.now() - t0;
  assert.equal(blocks.length, 3000);
  assert.ok(ms < 5000, `markdownBlocks took ${ms} ms for 3,000 blocks`);
});

// ----- applyMdTextEdits -----

test('applyMdTextEdits: applies edits in any order; no edits = same text', () => {
  const src = 'abc\ndef\nghi';
  const out = applyMdTextEdits(src, [edit(src, 8, 9, 'G'), edit(src, 0, 0, '>'), edit(src, 4, 7, '')]);
  assert.equal(out, '>abc\n\nGhi');
  assert.equal(applyMdTextEdits(src, []), src);
  assert.equal(applyMdTextEdits('ab😀cd', [edit('ab😀cd', 2, 4, '😃😃')]), 'ab😃😃cd');
});

test('applyMdTextEdits: throws on oldText mismatch, overlap and malformed edits', () => {
  const src = 'abcdef';
  assert.throws(() => applyMdTextEdits(src, [{ start: 0, end: 2, oldText: 'xx', text: 'y' }]), /not the text the edit expects/);
  assert.throws(() => applyMdTextEdits(src, [edit(src, 0, 3, 'x'), edit(src, 2, 4, 'y')]), /Overlapping/);
  assert.throws(() => applyMdTextEdits(src, [edit(src, 2, 2, 'x'), edit(src, 2, 4, 'y')]), /Overlapping/); // same start
  assert.throws(() => applyMdTextEdits(src, [edit(src, 1, 1, 'x'), edit(src, 1, 1, 'y')]), /Overlapping/);
  for (const bad of [
    { start: -1, end: 0, oldText: '', text: '' },
    { start: 3, end: 2, oldText: '', text: '' },
    { start: 0, end: 7, oldText: 'abcdef', text: '' },
    { start: 0.5, end: 1, oldText: 'a', text: '' },
    { start: 0, end: 1, oldText: 'a' },
    null,
  ]) {
    assert.throws(() => applyMdTextEdits(src, [bad]), /Invalid Markdown edit/, JSON.stringify(bad));
  }
  assert.throws(() => applyMdTextEdits(src, null), /Invalid Markdown edit/);
  // touching but not overlapping edits are fine
  assert.equal(applyMdTextEdits(src, [edit(src, 0, 2, 'X'), edit(src, 2, 4, 'Y')]), 'XYef');
});

test('applyMdTextEdits: never splits or creates a CRLF pair or a surrogate pair', () => {
  const crlf = '# T\r\n\r\nabc\r\n';
  assert.throws(() => applyMdTextEdits(crlf, [edit(crlf, 4, 5, '')]), /line break/); // removes the \n of a \r\n
  assert.throws(() => applyMdTextEdits(crlf, [edit(crlf, 3, 4, '')]), /line break/); // removes the \r of a \r\n
  assert.throws(() => applyMdTextEdits('# T\rx', [edit('# T\rx', 4, 4, '\n')]), /line break/); // lone \r + new \n
  assert.throws(() => applyMdTextEdits('a\nb', [edit('a\nb', 1, 1, 'x\r')]), /line break/); // new \r + old \n
  assert.equal(applyMdTextEdits(crlf, [edit(crlf, 3, 5, '\r\n\r\nnew')]), '# T\r\n\r\nnew\r\nabc\r\n');
  const emoji = 'ab😀cd';
  assert.throws(() => applyMdTextEdits(emoji, [edit(emoji, 3, 4, 'x')]), /split a character/);
  assert.throws(() => applyMdTextEdits(emoji, [edit(emoji, 2, 3, 'x')]), /split a character/);
  assert.throws(() => applyMdTextEdits(emoji, [edit(emoji, 0, 0, '\ud800')]), /split a character/);
  assert.throws(() => applyMdTextEdits(emoji, [edit(emoji, 0, 0, 'x\udc00')]), /split a character/);
});

// ----- verifyMdEdit -----

const doc = [
  '---',
  'title: Demo',
  '---',
  '',
  '# Heading',
  '',
  'Hello world, this is a paragraph.',
  '',
  '- apple',
  '- [ ] task',
  '',
  '| a | b |',
  '| - | - |',
  '| 1 | 2 |',
  '',
  '> [!NOTE]',
  '> An alert.',
  '',
  '```js',
  'x = 1;',
  '```',
  '',
].join('\n');

test('verifyMdEdit: accepts edits inside a block (paragraph, heading, list, table, code, fixed raw text)', () => {
  for (const [find, text] of [
    ['Hello world', 'Hello big world'],
    ['# Heading', '# Heading 2'],
    ['- apple', '- apple pie'],
    ['| 1 |', '| 10 |'],
    ['x = 1;', 'x = 2;'],
    ['title: Demo', 'title: Other'],
    ['> An alert.', '> An alert, edited.'],
  ]) {
    assert.deepEqual(check(doc, [replace(doc, find, text)]), { ok: true, bad: [] }, find);
  }
  // typing at the very start / end of a block
  const at = doc.indexOf('Hello');
  assert.deepEqual(check(doc, [edit(doc, at, at, 'Oh. ')]), { ok: true, bad: [] });
  const end = doc.indexOf('paragraph.') + 'paragraph.'.length;
  assert.deepEqual(check(doc, [edit(doc, end, end, ' More.')]), { ok: true, bad: [] });
  // two edits in two blocks at once
  assert.deepEqual(check(doc, [replace(doc, 'Heading', 'Title'), replace(doc, 'apple', 'pear')]), { ok: true, bad: [] });
});

test('verifyMdEdit: accepts new, deleted, split and joined blocks', () => {
  const src = '# T\n\nPara one.\n\nPara two.\n';
  const afterOne = src.indexOf('Para one.') + 'Para one.'.length;
  assert.deepEqual(check(src, [edit(src, afterOne, afterOne, '\n\nNew para.')]), { ok: true, bad: [] }); // insert
  assert.deepEqual(check(src, [edit(src, 3, afterOne, '')]), { ok: true, bad: [] }); // delete "Para one."
  assert.deepEqual(check(src, [replace(src, 'Para one.', 'Para\n\none.')]), { ok: true, bad: [] }); // Enter in the middle
  assert.deepEqual(check(src, [edit(src, afterOne, afterOne + 2, '')]), { ok: true, bad: [] }); // Backspace joins two paragraphs
});

test('verifyMdEdit: CRLF and BOM documents; the BOM must stay', () => {
  const crlf = doc.replace(/\n/g, '\r\n');
  assert.deepEqual(check(crlf, [replace(crlf, 'Hello world', 'Hi world')]), { ok: true, bad: [] });
  assert.deepEqual(check(crlf, [replace(crlf, '- apple', '- apple\r\n- banana')]), { ok: true, bad: [] });
  const bom = `﻿${doc}`;
  assert.deepEqual(check(bom, [replace(bom, 'Hello', 'Hey')]), { ok: true, bad: [] });
  assert.deepEqual(check(bom, [edit(bom, 0, 1, '')]), { ok: false, bad: [{ line: 0, kind: 'bom' }] });
  assert.deepEqual(check(doc, [edit(doc, 0, 0, '﻿')]).bad[0], { line: 0, kind: 'bom' });
});

test('verifyMdEdit: rejects an edit that merges an untouched paragraph (lazy continuation)', () => {
  const src = 'Intro.\n\nA paragraph\n- item\n';
  // the list item loses its marker: "item" now continues the untouched paragraph above it
  assert.deepEqual(check(src, [replace(src, '- item', 'item')]), { ok: false, bad: [{ line: 2, kind: 'paragraph' }] });
  // same for a quote marker removed under a paragraph
  const q = 'A paragraph\n> quoted\n';
  assert.equal(check(q, [replace(q, '> quoted', 'quoted')]).ok, false);
});

test('verifyMdEdit: rejects a "---" that turns the paragraph above into a setext heading', () => {
  const src = 'Title\n* * *\n\nmore\n';
  assert.deepEqual(check(src, [replace(src, '* * *', '---')]), { ok: false, bad: [{ line: 0, kind: 'paragraph' }] });
  const eq = 'Title\n\n***\n';
  // with a blank line in between it stays a paragraph + rule: allowed
  assert.deepEqual(check(eq, [replace(eq, '***', '---')]), { ok: true, bad: [] });
});

test('verifyMdEdit: rejects an edit that lets a list swallow the next untouched block', () => {
  const src = '- a\n- b\n\nAfter the list.\n\n# H\n';
  // indenting the paragraph puts it inside the (untouched) list
  assert.deepEqual(check(src, [replace(src, 'After', '  After')]).bad, [{ line: 0, kind: 'bullet_list' }]);
  // a new list item glued to an untouched list below merges both lists
  const two = 'Para\n\n- x\n';
  assert.equal(check(two, [replace(two, 'Para', '- new')]).ok, false);
});

test('verifyMdEdit: rejects edits that do not produce newText or are invalid', () => {
  const src = 'abc\n\ndef\n';
  assert.deepEqual(verifyMdEdit(src, [edit(src, 0, 3, 'x')], 'y\n\ndef\n'), { ok: false, bad: [{ line: -1, kind: 'edits' }] });
  assert.deepEqual(verifyMdEdit(src, [{ start: 0, end: 3, oldText: 'zzz', text: 'x' }], 'x\n\ndef\n'), { ok: false, bad: [{ line: -1, kind: 'edits' }] });
  assert.deepEqual(verifyMdEdit(src, [], 'abc\n\ndeX\n'), { ok: false, bad: [{ line: -1, kind: 'edits' }] });
});
