// Presentation renderer (src/renderers/pptx.ts) on test/samples/sample.pptx.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { renderPptx } = require('../.out/pptx.js');

const sample = (name) => new Uint8Array(fs.readFileSync(path.join(__dirname, '..', 'samples', name)));

let deckPromise;
const deck = () => (deckPromise ??= renderPptx(sample('sample.pptx')));

/** Every element of a slide, including the children of groups. */
function allElements(elements) {
  return elements.flatMap((e) => [e, ...(Array.isArray(e.elements) ? allElements(e.elements) : [])]);
}

test('sample.pptx: deck outline (size, slides, titles, hidden slide, notes flags)', async () => {
  const { meta } = await deck();
  assert.deepEqual([meta.width, meta.height, meta.slideCount], [960, 540, 8]);
  assert.equal(meta.slides[0].title, 'FileStudio Sample Presentation');
  assert.deepEqual(
    meta.slides.filter((s) => s.hidden).map((s) => [s.index, s.title]),
    [[6, 'Hidden slide']],
  );
  assert.deepEqual(meta.slides.map((s) => s.hasNotes), [true, true, false, false, false, true, false, true]);
  assert.ok(meta.fonts.includes('Calibri'));
});

test('sample.pptx: getSlide returns notes, layout elements and the hidden flag', async () => {
  const doc = await deck();
  const first = doc.getSlide(0);
  assert.match(first.notes, /^Speaker notes for the title slide\./);
  assert.ok(first.elements.some((e) => e.layout === true), 'layout / master elements are included');
  assert.equal(doc.getSlide(6).hidden, true);
});

test('sample.pptx: charts, pictures and tables', async () => {
  const doc = await deck();
  const charts = allElements(doc.getSlide(5).elements).filter((e) => e.type === 'chart');
  assert.deepEqual(charts.map((c) => c.chartType), ['barChart', 'pieChart']);
  assert.deepEqual(charts[0].data.map((series) => series.key), ['2023', '2024']);
  const images = allElements(doc.getSlide(3).elements).filter((e) => e.type === 'image');
  assert.equal(images.length, 2);
  for (const image of images) {
    assert.match(image.base64, /^data:image\/png;base64,iVBORw0KGgo/);
    assert.equal(image.blob, undefined);
  }
  assert.ok(doc.getSlide(4).elements.some((e) => e.type === 'table'));
});

test('getSlide rejects an index outside the deck', async () => {
  const doc = await deck();
  assert.throws(() => doc.getSlide(8), RangeError);
  assert.throws(() => doc.getSlide(-1), RangeError);
});

test('files that are not .pptx are rejected with a friendly Error', async () => {
  await assert.rejects(renderPptx(new Uint8Array(0)), /file is empty/);
  await assert.rejects(renderPptx(new TextEncoder().encode('hello world')), /not a valid \.pptx presentation \(it is not a ZIP package\)/);
  await assert.rejects(renderPptx(sample('sample.docx')), /holds a Word document/);
  await assert.rejects(renderPptx(sample('styled.xlsx')), /holds an Excel workbook/);
});
