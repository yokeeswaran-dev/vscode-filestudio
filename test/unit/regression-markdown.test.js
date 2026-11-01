// Regression tests for fixed 0.1.x Markdown preview bugs (src/renderers/markdown.ts). Each test names the bug id from
// the 0.1.x bug hunts. Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');

const { renderMarkdown } = require('../.out/markdown.js');

const md = (source, opts) => renderMarkdown(source, opts).html;
/** Text of the KaTeX <annotation> (the TeX source KaTeX rendered), or undefined when no math was rendered. */
const texOf = (html) => /<annotation encoding="application\/x-tex">([^<]*)<\/annotation>/.exec(html)?.[1];

// ===== TASK LISTS (the model for every 0.2.0 text edit) =====

test('H6-F7: a task marker after a nested list marker on the same line is found (column of the box character)', () => {
  const source = '- - [ ] nested\n- [ ] plain\n1. - [x] numbered\n';
  const { tasks } = renderMarkdown(source);
  assert.deepEqual(tasks, [
    { line: 0, column: 5, checked: false },
    { line: 1, column: 3, checked: false },
    { line: 2, column: 6, checked: true },
  ]);
  const lines = source.split('\n');
  for (const task of tasks) assert.match(lines[task.line][task.column], /[ xX]/, `line ${task.line}: column points into the brackets`);
});

test('H9-F8: task checkboxes have an accessible name (the item text)', () => {
  const html = md('- [ ] buy milk\n- [x] **bold** task\n');
  const inputs = html.match(/<input[^>]*>/g) ?? [];
  assert.equal(inputs.length, 2);
  assert.match(inputs[0], /aria-label="buy milk"/);
  assert.match(inputs[1], /aria-label="bold task"/);
});

// ===== MATH =====

test('H6-F9: a paragraph that starts with "$$" does not swallow the rest of the document as math', () => {
  const html = md('$$ expands to the PID\n\n```\necho $$\n```\n\nafter');
  assert.ok(!/katex/.test(html), 'no math rendered');
  assert.match(html, /<p data-source-line="0">\$\$ expands to the PID<\/p>/);
  assert.match(html, /<pre data-source-line="2"><code[^>]*>echo \$\$\n<\/code><\/pre>/, 'code block kept');
  assert.match(html, /<p data-source-line="6">after<\/p>/);
});

test('H6-F10: code spans win over inline math ("$5 ... `$HOME`")', () => {
  const html = md('It costs $5. Use `$HOME` here.');
  assert.match(html, /<code>\$HOME<\/code>/);
  assert.ok(!/katex/.test(html), 'no math rendered');
});

test('H6-F13: dollar-backtick math $`...`$ renders as math without the backticks', () => {
  const html = md('x $`\\sqrt{3}`$ y');
  assert.match(html, /class="math math-inline"/);
  assert.equal(texOf(html), '\\sqrt{3}');
  assert.ok(!html.includes('`'), 'no backtick left in the output');
});

// ===== GITHUB COMPATIBILITY =====

test('H6-F8: a heading with an emoji shortcode gets the GitHub slug', () => {
  const { toc } = renderMarkdown('## :rocket: Intro\n\n## Intro\n');
  assert.deepEqual(
    toc.map((entry) => entry.slug),
    ['rocket-intro', 'intro'],
  );
});

test('H6-F11: alerts only where GitHub makes them (not in list items, nested quotes or without a body)', () => {
  const html = md('- > [!NOTE]\n  > text\n\n> > [!TIP]\n> > x\n\n> [!NOTE]\n\n> [!WARNING]\n> real\n');
  const alerts = html.match(/class="markdown-alert [^"]*"/g) ?? [];
  assert.deepEqual(alerts, ['class="markdown-alert markdown-alert-warning"'], 'only the top-level alert with a body');
  assert.match(html, /<p data-source-line="0">\[!NOTE\]\ntext<\/p>/, 'list item: plain blockquote');
});

test('H6-F12: repeated references to one footnote all show its number', () => {
  const html = md('a[^1] b[^1]\n\n[^1]: note\n');
  const labels = [...html.matchAll(/<sup class="footnote-ref"><a [^>]*>([^<]*)<\/a>/g)].map((m) => m[1]);
  assert.deepEqual(labels, ['1', '1']);
});

test('H6-F14: single-tilde strikethrough renders like GitHub', () => {
  assert.match(md('a ~single~ b ~~double~~ c'), /a <s>single<\/s> b <s>double<\/s> c/);
});

test('H6-F17: linkify does not link "//host/path" or "ftp://" (GitHub does not)', () => {
  const html = md('see //host/path and ftp://x.example/f and https://ok.example/x');
  const links = [...html.matchAll(/<a href="([^"]*)"/g)].map((m) => m[1]);
  assert.deepEqual(links, ['https://ok.example/x']);
});

// ===== RESOURCES =====

test('H6-F1: srcset of <img> and <picture><source> are rewritten through resolveResource', () => {
  const html = md('<img srcset="img/a.png 1x, img/a@2x.png 2x" src="img/a.png">\n\n<picture><source srcset="img/b.png 1x"><img src="img/c.png"></picture>', {
    resolveResource: (href) => `https://res.test/${href}`,
  });
  assert.match(html, /<img data-source-line="0" srcset="https:\/\/res\.test\/img\/a\.png 1x, https:\/\/res\.test\/img\/a@2x\.png 2x" src="https:\/\/res\.test\/img\/a\.png">/);
  assert.match(html, /<source srcset="https:\/\/res\.test\/img\/b\.png 1x">/);
});
