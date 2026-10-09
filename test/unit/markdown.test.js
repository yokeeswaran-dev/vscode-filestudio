// Markdown preview renderer (src/renderers/markdown.ts) on test/samples/kitchen-sink.md.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { renderMarkdown, markdownHeadings } = require('../.out/markdown.js');

const source = fs.readFileSync(path.join(__dirname, '..', 'samples', 'kitchen-sink.md'), 'utf8');
const lines = source.split(/\r?\n/);
const { html, toc, tasks } = renderMarkdown(source);

/** How often `re` (a global RegExp) matches the rendered HTML. */
const count = (re) => (html.match(re) ?? []).length;

test('blocks carry data-source-line (scroll sync); front matter becomes a table', () => {
  assert.ok(count(/data-source-line="\d+"/g) > 100);
  assert.match(html, /<table class="front-matter" data-source-line="0">/);
  assert.match(html, /<h1 id="kitchen-sink" data-source-line="22">Kitchen Sink<\/h1>/);
});

test('task list checkboxes have data-line and match the source markers', () => {
  assert.equal(tasks.length, 8);
  for (const task of tasks) {
    const marker = lines[task.line][task.column];
    assert.equal(marker === 'x' || marker === 'X', task.checked, `task on line ${task.line}`);
    assert.match(html, new RegExp(`data-line="${task.line}"${task.checked ? ' checked' : '(?! checked)'}`));
  }
  assert.deepEqual(tasks[0], { line: 134, column: 3, checked: true });
});

test('GitHub alerts render as titled callouts', () => {
  for (const kind of ['note', 'tip', 'important', 'warning', 'caution']) {
    assert.match(html, new RegExp(`<div class="markdown-alert markdown-alert-${kind}" data-source-line="\\d+">`), kind);
  }
  assert.match(html, /<p class="markdown-alert-title" data-source-line="\d+">Note<\/p>/);
});

test('math is rendered by KaTeX (inline and display)', () => {
  assert.ok(count(/<span class="katex">/g) >= 3);
  assert.ok(count(/class="katex-display"/g) >= 1);
});

test('mermaid fences become escaped <div class="mermaid"> blocks', () => {
  assert.equal(count(/<div class="mermaid" data-source-line="\d+">/g), 3);
  assert.match(html, /<div class="mermaid" data-source-line="\d+">flowchart LR\n {2}A\[Open file\] --&gt; B/);
});

test('headings get GitHub-style ids and a table of contents', () => {
  assert.deepEqual(toc[0], { level: 1, text: 'Kitchen Sink', slug: 'kitchen-sink', line: 22 });
  const duplicates = toc.filter((h) => h.text === 'Duplicate heading').map((h) => h.slug);
  assert.deepEqual(duplicates, ['duplicate-heading', 'duplicate-heading-1']);
  for (const { slug } of toc) assert.ok(html.includes(` id="${slug}"`), `heading id ${slug}`);
  assert.deepEqual(markdownHeadings(source), toc, 'markdownHeadings agrees with the rendered toc');
});

test('footnotes: references and the footnote list', () => {
  assert.match(html, /<section data-source-line="\d+" class="footnotes">/);
  assert.equal(count(/class="footnote-item"/g), 3);
});

test('resolveResource rewrites relative image paths', () => {
  const seen = [];
  const result = renderMarkdown(source, {
    resolveResource: (href) => (seen.push(href), href.startsWith('images/') ? `https://res.test/${href}` : undefined),
  });
  assert.match(result.html, /<img src="https:\/\/res\.test\/images\/sample\.png" alt="Sample image"/);
  assert.match(result.html, /<img src="https:\/\/img\.shields\.io\//, 'remote images are left alone');
});

test('never throws on odd input', () => {
  const odd = ['', '﻿# BOM', '$$', '```mermaid', '[^missing]', '> [!BOGUS]\n> text', '\u0000\u0001', '#'.repeat(10000), '- [ ]', '$\\frac{$', null, undefined, 42];
  for (const input of odd) {
    assert.equal(typeof renderMarkdown(input).html, 'string', JSON.stringify(input));
  }
  assert.match(renderMarkdown('﻿# BOM').html, /^<h1 id="bom" data-source-line="0">BOM<\/h1>/);
});
