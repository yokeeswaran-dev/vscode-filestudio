// Word renderer (src/renderers/docx.ts) on test/samples/sample.docx.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { renderDocx } = require('../.out/docx.js');

const sample = (name) => new Uint8Array(fs.readFileSync(path.join(__dirname, '..', 'samples', name)));

let resultPromise;
const rendered = () => (resultPromise ??= renderDocx(sample('sample.docx')));

test('sample.docx: converts without warnings', async () => {
  const { html, warnings } = await rendered();
  assert.deepEqual(warnings, []);
  assert.match(html, /<p class="doc-subtitle">/);
});

test('sample.docx: title and headings', async () => {
  const { html } = await rendered();
  assert.match(html, /^<h1 class="doc-title">FileStudio Sample Document<\/h1>/);
  assert.match(html, /<h1>Introduction<\/h1>/);
  assert.match(html, /<h2>Numbered list<\/h2><ol><li>Step one<\/li>/);
});

test('sample.docx: inline formatting, links, footnote, special characters', async () => {
  const { html } = await rendered();
  assert.match(html, /<strong>bold<\/strong>/);
  assert.match(html, /<a href="https:\/\/code\.visualstudio\.com\/">Visual Studio Code website<\/a>/);
  assert.match(html, /<sup><a href="#footnote-1" id="footnote-ref-1">\[1\]<\/a><\/sup>/);
  assert.match(html, /café, naïve, Straße/);
});

test('sample.docx: a table and an embedded image as a data: URI', async () => {
  const { html } = await rendered();
  assert.match(html, /<table>[\s\S]*<td>[\s\S]*<\/table>/);
  assert.match(html, /<img alt="Bar chart generated as a PNG" src="data:image\/png;base64,iVBORw0KGgo/);
});

test('a 0-byte .docx is an empty document, as in Word (not an error)', async () => {
  assert.deepEqual(await renderDocx(new Uint8Array(0)), { html: '<p class="doc-empty"><em>This document is empty.</em></p>', warnings: [] });
});

test('files that are not .docx are rejected with a friendly Error', async () => {
  await assert.rejects(renderDocx(new TextEncoder().encode('hello world')), /not a valid \.docx document \(it is not a ZIP package\)/);
  await assert.rejects(
    renderDocx(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0])),
    /password-protected\/encrypted or a legacy Word 97-2003/,
  );
  await assert.rejects(renderDocx(sample('styled.xlsx')), (err) => {
    assert.match(err.message, /has no Word document body/);
    assert.ok(err.cause, 'the library error is kept as cause');
    return true;
  });
});
