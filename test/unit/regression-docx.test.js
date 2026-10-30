// Regression tests for fixed 0.1.x Word preview bugs (src/renderers/docx.ts). Fixtures are test/samples/sample.docx with
// zip parts replaced in memory (the sample itself is never written). Each test names the bug id from the 0.1.x bug hunts.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

const { renderDocx } = require('../.out/docx.js');

const SAMPLE = fs.readFileSync(path.join(__dirname, '..', 'samples', 'sample.docx'));
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** sample.docx with zip parts replaced: `parts` maps a part name to (old text or undefined) -> new text. */
async function patchedDocx(parts) {
  const zip = await JSZip.loadAsync(SAMPLE);
  for (const [name, edit] of Object.entries(parts)) {
    const old = zip.file(name) ? await zip.file(name).async('string') : undefined;
    const next = edit(old);
    assert.notEqual(next, old, `fixture: ${name} changed`);
    zip.file(name, next);
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

/** A document.xml body with just these paragraphs (raw <w:p> XML). */
const bodyOf = (paragraphs) => (xml) => xml.replace(/<w:body>[\s\S]*?(<w:sectPr>)/, `<w:body>${paragraphs}$1`);

test('H7-F1: a style map stored inside the .docx (mammoth/style-map) is ignored', async () => {
  const bytes = await patchedDocx({
    'mammoth/style-map': () => "p => h6.injected-by-docx:fresh\nr => script\np[style-name='Heading 1'] => style",
  });
  const { html } = await renderDocx(bytes);
  assert.ok(!/injected-by-docx|<h6|<script|<style/.test(html), 'no element from the embedded map');
  assert.match(html, /<h1>Introduction<\/h1>/, 'our own style map still applies');
});

test('H7-F5: page and column breaks keep the words on either side apart', async () => {
  const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const bytes = await patchedDocx({
    'word/document.xml': bodyOf(
      `<w:p>${run('Before page')}<w:r><w:br w:type="page"/></w:r>${run('After page')}</w:p>` +
        `<w:p>${run('Before column')}<w:r><w:br w:type="column"/></w:r>${run('After column')}</w:p>`,
    ),
  });
  const { html } = await renderDocx(bytes);
  assert.match(html, /Before page<br \/>After page/);
  assert.match(html, /Before column<br \/>After column/);
});

/** A 1x1 24-bit BMP (red pixel). */
function bmpBytes() {
  const header = Buffer.alloc(54);
  header.write('BM', 0, 'latin1');
  header.writeUInt32LE(58, 2); // file size
  header.writeUInt32LE(54, 10); // pixel data offset
  header.writeUInt32LE(40, 14); // info header size
  header.writeInt32LE(1, 18); // width
  header.writeInt32LE(1, 22); // height
  header.writeUInt16LE(1, 26); // planes
  header.writeUInt16LE(24, 28); // bits per pixel
  header.writeUInt32LE(4, 34); // image size
  return Buffer.concat([header, Buffer.from([0x00, 0x00, 0xff, 0x00])]);
}

test('H7-F6: a BMP picture is shown, with no "unlikely to display" warning', async () => {
  const zip = await JSZip.loadAsync(SAMPLE);
  zip.remove('word/media/image1.png');
  zip.file('word/media/image1.bmp', bmpBytes());
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('media/image1.png', 'media/image1.bmp'));
  const types = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', types.replace('<Default Extension="xml"', '<Default Extension="bmp" ContentType="image/bmp"/><Default Extension="xml"'));
  const { html, warnings } = await renderDocx(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })));
  assert.deepEqual(warnings, []);
  assert.match(html, /<img [^>]*src="data:image\/bmp;base64,Qk06/);
});

test('H7-F7: a linked (not embedded) picture gets a placeholder and a friendly warning', async () => {
  const bytes = await patchedDocx({
    'word/document.xml': (xml) => xml.replace('<a:blip r:embed="rId6"/>', '<a:blip r:link="rId90"/>'),
    'word/_rels/document.xml.rels': (xml) =>
      xml.replace('</Relationships>', `<Relationship Id="rId90" Type="${REL}/image" Target="file:///C:/pictures/photo.png" TargetMode="External"/></Relationships>`),
  });
  const { html, warnings } = await renderDocx(bytes);
  assert.equal(warnings.length, 1, JSON.stringify(warnings));
  assert.match(warnings[0], /^A linked image \(file:\/\/\/C:\/pictures\/photo\.png\) is not stored in the document, so it is not shown; a placeholder marks its position\.$/);
  const img = /<img [^>]*>/.exec(html)?.[0] ?? '';
  assert.match(img, /src="data:image\/svg\+xml;base64,/, 'placeholder image in place of the picture');
  assert.ok(Buffer.from(/base64,([^"]*)"/.exec(img)[1], 'base64').toString('utf8').includes('Linked image not shown'));
});

test('K3: a damaged zip gets a readable message (no zip library text)', async () => {
  const truncated = new TextEncoder().encode('PK\u0003\u0004 truncated');
  const half = new Uint8Array(SAMPLE.subarray(0, SAMPLE.length >> 1));
  for (const [label, bytes] of [['truncated', truncated], ['half a document', half]]) {
    await assert.rejects(renderDocx(bytes), (err) => {
      assert.match(err.message, /^This file is not a valid \.docx document: its ZIP package is damaged/, label);
      assert.doesNotMatch(err.message, /central directory|Corrupted zip|end of data/i, `${label}: library text not in the message`);
      return true;
    });
  }
});
