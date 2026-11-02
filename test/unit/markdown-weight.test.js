// WYSIWYG block weight (v0.2.1, B2; src/renderers/markdown.ts, section EDIT BLOCKS): markdownWeight(source, blocks) says how
// costly the editor finds a document: tables and lists cost most, protected (raw) blocks little.
const { test } = require('node:test');
const assert = require('node:assert/strict');

const { markdownBlocks, markdownWeight, MD_WEIGHT_LIMIT } = require('../.out/markdown.js');

const weigh = (src) => markdownWeight(src, markdownBlocks(src));

test('weight: a plain block costs about 1; a blank document costs 0', () => {
  assert.equal(weigh(''), 0);
  assert.equal(weigh('Just a paragraph.\n'), 1);
  assert.equal(weigh('# Heading\n'), 1);
});

test('weight: a table or a list costs more than a paragraph, and grows with its rows or items', () => {
  const table = (rows) => `| A | B |\n| --- | --- |\n${Array.from({ length: rows }, (_, i) => `| ${i} | x |`).join('\n')}\n`;
  const list = (items) => Array.from({ length: items }, (_, i) => `- item ${i}`).join('\n') + '\n';
  assert.ok(weigh(table(2)) >= 4 * weigh('Para.\n'), 'a small table is at least 4 paragraphs');
  assert.ok(weigh(table(10)) > weigh(table(2)));
  assert.ok(weigh(list(5)) >= 4 * weigh('Para.\n'));
  assert.ok(weigh(list(20)) > weigh(list(5)));
});

test('weight: protected blocks (front matter, HTML) cost less than a paragraph; the sum is over all blocks', () => {
  assert.ok(weigh('<div>x</div>\n\n'.repeat(6)) < weigh('Para.\n\n'.repeat(6)), 'protected blocks cost less');
  const one = weigh('Para.\n');
  assert.equal(weigh('Para.\n\nPara.\n\nPara.\n'), 3 * one);
});

test('weight: the limit is a number a 3,000-block mixed document is over and a 300-block one is far under', () => {
  assert.equal(typeof MD_WEIGHT_LIMIT, 'number');
  const doc = (n) => Array.from({ length: n }, (_, i) => (i % 3 === 0 ? `| a | b |\n| - | - |\n| ${i} | y |` : i % 3 === 1 ? `- a ${i}\n- b ${i}\n- c ${i}` : `Para ${i}`)).join('\n\n') + '\n';
  assert.ok(weigh(doc(3000)) > MD_WEIGHT_LIMIT, `3000 blocks: ${weigh(doc(3000))}`);
  assert.ok(weigh(doc(300)) < MD_WEIGHT_LIMIT / 4);
});
