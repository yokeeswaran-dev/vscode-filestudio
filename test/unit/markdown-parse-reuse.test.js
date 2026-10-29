// Markdown edit blocks: a document is parsed once, not once per question (fix round 1, P6). The host asks for the blocks of
// the same text several times per WYSIWYG edit (verifyMdEdit on the old and the new text, then the rendering with the
// context check, then the headings); on a 3,000-block document each parse costs 60-110 ms on the extension host thread.
// test/build-tests.js counts the parses: `globalThis.__fsMdParses` holds the source length of every markdown-it parse.
// Run with `npm test`.
const { test } = require('node:test');
const assert = require('node:assert/strict');

const { markdownBlocks, markdownHeadings, renderMarkdown, applyMdTextEdits, verifyMdEdit } = require('../.out/markdown.js');

/** A document of `sections` x (heading, paragraph with a reference link, list) + a footnote; `tag` makes its text unique. */
function doc(tag, sections = 30) {
  const parts = [`# ${tag}`];
  for (let i = 0; i < sections; i++) {
    parts.push(`## Section ${i}`, `Paragraph ${i} of ${tag}: **bold**, a [reference link][ref] and a footnote.[^n]`, `- item ${i}a\n- item ${i}b`);
  }
  parts.push('[ref]: https://example.com/ref', '[^n]: The footnote.');
  return parts.join('\n\n') + '\n';
}

/** Runs `fn` and returns [its result, how many markdown-it parses of (about) the whole of `text` it made]. */
function parsesOfWhole(text, fn) {
  const log = globalThis.__fsMdParses;
  assert.ok(Array.isArray(log), 'test/build-tests.js counts the markdown-it parses');
  log.length = 0;
  const result = fn();
  return [result, log.filter((length) => length >= text.length * 0.9).length];
}

test('markdownBlocks on the same text a second time does not parse it again', () => {
  const text = doc('reuse-1');
  const [first, n1] = parsesOfWhole(text, () => markdownBlocks(text));
  const [second, n2] = parsesOfWhole(text, () => markdownBlocks(text));
  assert.equal(n1, 1, 'the first call parses');
  assert.equal(n2, 0, 'the second call uses the parse of the first');
  assert.deepEqual(second, first);
});

test('verifyMdEdit after the blocks of the old text were asked for parses only the new text', () => {
  const old = doc('reuse-2');
  markdownBlocks(old, { contextCheck: true }); // what the rendering before the edit did
  const at = old.indexOf('item 3a');
  const edits = [{ start: at, end: at + 4, oldText: 'item', text: 'entry' }];
  const newText = applyMdTextEdits(old, edits);
  const [verdict, n] = parsesOfWhole(old, () => verifyMdEdit(old, edits, newText));
  assert.deepEqual(verdict, { ok: true, bad: [] });
  assert.equal(n, 1, 'only the new text is parsed');
  // ... and the rendering of the new text right after it (context check, headings) parses nothing more
  const [, more] = parsesOfWhole(old, () => {
    markdownBlocks(newText, { contextCheck: true });
    markdownHeadings(newText);
  });
  assert.equal(more, 0);
});

test('markdownHeadings of a text whose blocks were just parsed does not parse again, and gives the preview\'s toc', () => {
  const text = doc('reuse-3');
  markdownBlocks(text, { contextCheck: true });
  const [toc, n] = parsesOfWhole(text, () => markdownHeadings(text));
  assert.equal(n, 0);
  assert.deepEqual(toc, renderMarkdown(text).toc);
  assert.ok(toc.length > 30);
});

test('markdownHeadings of an unknown text still parses it (other documents are looked up for links)', () => {
  const text = doc('reuse-4');
  const [toc, n] = parsesOfWhole(text, () => markdownHeadings(text));
  assert.equal(n, 1);
  assert.deepEqual(toc, renderMarkdown(text).toc);
});

test('the context check gives the same blocks the second time (reference links and footnotes stay fixed)', () => {
  const text = doc('reuse-5');
  const first = markdownBlocks(text, { contextCheck: true });
  const second = markdownBlocks(text, { contextCheck: true });
  const third = markdownBlocks(text, { contextCheck: true });
  assert.deepEqual(second, first);
  assert.deepEqual(third, first);
  const fixed = first.filter((b) => b.fixed && b.reason === 'depends on other parts of the document');
  assert.ok(fixed.length >= 30, `the paragraphs with a reference link are fixed (${fixed.length})`);
  // a call without the context check after one with it is not affected by it
  assert.equal(markdownBlocks(text).filter((b) => b.reason === 'depends on other parts of the document').length, 0);
});

test('the results are copies: changing one does not change what the next call returns', () => {
  const text = doc('reuse-6');
  const blocks = markdownBlocks(text);
  const expected = JSON.parse(JSON.stringify(blocks));
  blocks[0].fixed = true;
  blocks[1].start = -5;
  blocks.pop();
  assert.deepEqual(markdownBlocks(text), expected);
  const toc = markdownHeadings(text);
  const tocExpected = JSON.parse(JSON.stringify(toc));
  toc[0].text = 'changed';
  toc.pop();
  assert.deepEqual(markdownHeadings(text), tocExpected);
});

test('two texts of the same length never share a result', () => {
  const a = '# Alpha\n\nsome text here\n';
  const b = '# Bravo\n\nsome text here\n';
  assert.equal(a.length, b.length);
  markdownBlocks(a);
  markdownBlocks(b);
  assert.deepEqual(markdownHeadings(a).map((t) => t.text), ['Alpha']);
  assert.deepEqual(markdownHeadings(b).map((t) => t.text), ['Bravo']);
  assert.deepEqual(markdownHeadings(a).map((t) => t.slug), ['alpha']);
});

test('a very large text is not kept (memory): it is parsed again', () => {
  const text = '# Huge\n\n' + 'word '.repeat(90_000) + '\n'; // about 450,000 characters, 2 blocks
  const [, n1] = parsesOfWhole(text, () => markdownBlocks(text));
  const [, n2] = parsesOfWhole(text, () => markdownBlocks(text));
  assert.equal(n1, 1);
  assert.equal(n2, 1);
});
