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

// ----- 0.1.2: decks changed in memory from sample.pptx (jszip) -----

const JSZip = require('jszip');

/** sample.pptx with `edit(zip)` applied, rendered. */
async function patchedDeck(edit) {
  const zip = await JSZip.loadAsync(sample('sample.pptx'));
  await edit(zip);
  return renderPptx(await zip.generateAsync({ type: 'uint8array' }));
}

/** Replaces `from` (must occur) in a part of the zip. */
async function replaceIn(zip, part, from, to) {
  const xml = await zip.file(part).async('string');
  assert.ok(xml.includes(from), `${part} contains ${from}`);
  zip.file(part, xml.split(from).join(to));
}

test('a missing picture or a damaged slide no longer fails the whole deck (R3-008)', async () => {
  const doc = await patchedDeck(async (zip) => {
    zip.remove('ppt/media/image1.png');
    zip.file('ppt/slides/slide3.xml', 'this is not XML');
  });
  assert.equal(doc.meta.slideCount, 8);
  assert.ok(doc.meta.lossy.includes('Pictures missing from the file'), doc.meta.lossy.join());
  assert.ok(doc.meta.lossy.includes('Damaged slides (shown empty)'), doc.meta.lossy.join());
  const pictures = doc.getSlide(3).elements.filter((e) => e.type === 'image' && !e.layout);
  // Picture 20 (image1.png, missing) is a placeholder; Picture 21 (image2.png) is still shown.
  assert.deepEqual(
    pictures.map((p) => [Math.round(p.left), p.base64.slice(0, 26)]),
    [
      [480, 'data:image/svg+xml;base64,'],
      [768, 'data:image/png;base64,iVBO'],
    ],
  );
  assert.ok(doc.getSlide(2).elements.some((e) => e.name === 'This slide is damaged and cannot be shown'));
  assert.equal(doc.getSlide(3).notes, '');
  assert.equal(doc.getSlide(4).elements.some((e) => e.type === 'table'), true, 'the other slides are drawn');
});

test('charts: combo plots, date categories, decoded names, title, legend, formats, labels (R3-012/013/014/051)', async () => {
  const part = 'ppt/charts/chart1.xml';
  const doc = await patchedDeck(async (zip) => {
    let xml = await zip.file(part).async('string');
    // Series 2 moves to a line plot on a secondary axis; categories become dates; series 1 gets value labels.
    const series = [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0]);
    assert.equal(series.length, 2);
    const dates = '<c:cat><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f><c:numCache><c:formatCode>mmm-yy</c:formatCode><c:ptCount val="4"/>' +
      [37377, 37408, 37438, 37469].map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') + '</c:numCache></c:numRef></c:cat>';
    const withDates = (ser) => ser.replace(/<c:cat>[\s\S]*?<\/c:cat>/, dates);
    const first = withDates(series[0])
      .replace('<c:v>2023</c:v>', '<c:v>R&amp;D</c:v>')
      .replace('<c:invertIfNegative val="0"/>', '<c:invertIfNegative val="0"/><c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>');
    xml = xml.replace(series[0], first).replace(series[1], '');
    xml = xml.replace(
      '</c:barChart>',
      `</c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${withDates(series[1]).replace('<c:invertIfNegative val="0"/>', '')}` +
        '<c:marker val="1"/><c:axId val="600000001"/><c:axId val="600000002"/></c:lineChart>',
    );
    xml = xml.replace(
      '</c:plotArea>',
      '<c:catAx><c:axId val="600000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="600000002"/></c:catAx>' +
        '<c:valAx><c:axId val="600000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:crossAx val="600000001"/><c:crosses val="max"/></c:valAx></c:plotArea>',
    );
    xml = xml.replace(/(<c:valAx><c:axId val="500000002"\/>[\s\S]*?)<c:numFmt formatCode="General" sourceLinked="1"\/>/, (_all, head) => `${head}<c:numFmt formatCode="&quot;$&quot;#,##0" sourceLinked="0"/>`);
    xml = xml.replace(/<c:legendPos val="\w+"\/>/, '<c:legendPos val="t"/>');
    zip.file(part, xml);
  });
  const chart = doc.getSlide(5).elements.find((e) => e.type === 'chart' && e.chartType === 'barChart');
  assert.ok(chart, 'the combo chart keeps its first plot type');
  assert.deepEqual(chart.data.map((s) => s.key), ['R&D', '2024']);
  assert.deepEqual(chart.seriesTypes, ['barChart', 'lineChart']);
  assert.deepEqual(chart.secondary, [false, true]);
  assert.deepEqual(Object.values(chart.data[0].xlabels), ['May-02', 'Jun-02', 'Jul-02', 'Aug-02']);
  assert.equal(chart.title, 'Units sold per quarter');
  assert.equal(chart.titleSize, 16);
  assert.equal(chart.legend, 't');
  assert.equal(chart.valueFormat, '"$"#,##0');
  assert.equal(chart.valueFormat2, '0%');
  assert.equal(chart.gridlines, true);
  assert.equal(chart.fontSize, 12);
  assert.deepEqual(chart.dataLabels, [{ value: true, percent: false, category: false, series: false }, null]);
  const pie = doc.getSlide(5).elements.find((e) => e.type === 'chart' && e.chartType === 'pieChart');
  assert.equal(pie.title, 'Sessions by device (%)');
  assert.equal(pie.legend, 'b');
  assert.deepEqual(pie.dataLabels, [{ value: false, percent: true, category: false, series: false }]);
});

test('SVG-only pictures, slide jump actions and linked pictures (R3-054/050/053)', async () => {
  const slide = 'ppt/slides/slide4.xml';
  const rels = 'ppt/slides/_rels/slide4.xml.rels';
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>';
  const doc = await patchedDeck(async (zip) => {
    // Picture 21: only the SVG extension (no r:embed on a:blip).
    await replaceIn(
      zip,
      slide,
      '<a:blip r:embed="rId3"/>',
      '<a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId9"/></a:ext></a:extLst></a:blip>',
    );
    // Picture 20: linked to a file, not stored in the package.
    await replaceIn(zip, slide, '<a:blip r:embed="rId2"/>', '<a:blip r:link="rId2"/>');
    await replaceIn(zip, rels, 'Target="../media/image1.png"/>', 'Target="file:///C:/pictures/photo.png" TargetMode="External"/>');
    // The link button jumps to the last slide; a caption run jumps to the next slide.
    await replaceIn(zip, slide, '<a:hlinkClick r:id="rId4"/>', '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=lastslide"/>');
    await replaceIn(
      zip,
      slide,
      '<a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr><a:t>Pictures:',
      '<a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></a:rPr><a:t>Pictures:',
    );
    await replaceIn(zip, rels, '</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.svg"/></Relationships>');
    await replaceIn(zip, '[Content_Types].xml', '<Default Extension="png"', '<Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="png"');
    zip.file('ppt/media/image9.svg', svg);
  });
  const elements = doc.getSlide(3).elements.filter((e) => !e.layout);
  const at = (left) => elements.find((e) => e.type === 'image' && Math.round(e.left) === left);
  assert.equal(at(768).base64, `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`);
  assert.equal(at(480).name, 'Linked picture not loaded');
  assert.match(at(480).base64, /^data:image\/svg\+xml;base64,/);
  assert.ok(doc.meta.lossy.includes('Linked pictures'));
  assert.ok(elements.some((e) => e.link === 'ppaction://hlinkshowjump?jump=lastslide'), 'shape action link');
  assert.ok(
    elements.some((e) => typeof e.content === 'string' && e.content.includes('href="#ppaction://hlinkshowjump?jump=nextslide"')),
    'text run action link',
  );
});

/** pptx.ts marker characters (private use U+E000 to U+E009): none may reach the webview. */
const MARKER_CHARS = new RegExp(`[${String.fromCharCode(0xe000)}-${String.fromCharCode(0xe009)}]`);
const ZERO_WIDTH_SPACE = String.fromCharCode(0x200b);

/** DrawingML tint / shade: blend toward white / black in linear RGB (independent of pptx.ts). */
function linearBlend(hex, amount, toWhite) {
  const lin = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const srgb = (c) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
  return `#${[0, 2, 4]
    .map((i) => lin(parseInt(hex.slice(i, i + 2), 16) / 255) * amount + (toWhite ? 1 - amount : 0))
    .map((c) => Math.round(srgb(c) * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;
}

test('text: inherited bullets, numbering schemes, soft breaks, tabs, highlight, underlines, link colour, vertical text, line spacing (R3-009/010/011/045/046/052)', async () => {
  const slide = 'ppt/slides/slide2.xml';
  const bullet = '<a:pPr><a:buFont typeface="Arial" panose="020B0604020202020204"/><a:buChar char="•"/></a:pPr>';
  const doc = await patchedDeck(async (zip) => {
    // Level-one paragraphs without their own bullet: the master bodyStyle bullet (•) applies.
    await replaceIn(zip, slide, bullet, '');
    await replaceIn(zip, slide, '<a:buAutoNum type="arabicPeriod"/>', '<a:buAutoNum type="alphaUcPeriod" startAt="3"/>');
    await replaceIn(zip, slide, '<a:t>Second level bullet</a:t>', '<a:t>Second level</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US" dirty="0"/><a:t>bullet</a:t>');
    await replaceIn(zip, slide, '<a:t>Third level bullet</a:t>', '<a:t>Third\tlevel</a:t>');
    await replaceIn(zip, slide, '<a:pPr lvl="2">', '<a:pPr lvl="2"><a:lnSpc><a:spcPct val="200000"/></a:lnSpc>');
    await replaceIn(
      zip,
      slide,
      '<a:rPr lang="en-US" b="1" dirty="0"/><a:t>bold',
      '<a:rPr lang="en-US" b="1" u="dbl" dirty="0"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>bold',
    );
    await replaceIn(zip, slide, '<a:t>Numbered item two</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>', '<a:t>Numbered item two</a:t></a:r></a:p><a:p><a:endParaRPr lang="en-US" sz="800"/></a:p>');
    await replaceIn(zip, slide, '<p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>', '<p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr vert="vert270"/>');
  });
  const [title, body] = doc.getSlide(1).elements.filter((e) => !e.layout);
  const markers = [...body.content.matchAll(/<span class="pptx-bullet"[^>]*>([^<]*)<\/span>/g)].map((m) => m[1]);
  assert.deepEqual(markers, ['•', '–', '▪', '•', 'C.', 'D.']);
  assert.ok(!/<[uo]l>/.test(body.content), 'bullets are markers, not lists');
  assert.match(body.content, /Second&nbsp;level<br>bullet/);
  assert.match(body.content, /Third\tlevel/);
  assert.match(body.content, /line-height: 2\.4;/, 'a 200% line is 2 x 1.2 x the font size');
  assert.match(body.content, /<span style="background-color:#FFFF00;text-decoration-line:underline;text-decoration-style:double">bold<\/span>/);
  assert.match(body.content, /<span style="color: #0563C1;[^"]*"><a href="https:\/\/code\.visualstudio\.com\/"/, 'theme hyperlink colour');
  assert.ok(body.content.includes(`font-size: 8pt;font-family: Calibri;">${ZERO_WIDTH_SPACE}</span>`), 'an empty paragraph has its end-of-paragraph size');
  assert.ok(!MARKER_CHARS.test(body.content), 'no marker characters left');
  assert.equal(title.isVertical, true);
  assert.equal(title.vert270, true);
  assert.equal(body.vert270, undefined);
});

test('colours and outlines: tint / shade in linear RGB, outline without a width, table cell margins (R3-048/042/007)', async () => {
  const doc = await patchedDeck(async (zip) => {
    await replaceIn(
      zip,
      'ppt/slides/slide3.xml',
      '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700">',
      '<a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:ln>',
    );
    await replaceIn(
      zip,
      'ppt/slides/slide5.xml',
      '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:tcPr>',
      '<a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:tcPr>',
    );
  });
  const rect = doc.getSlide(2).elements.find((e) => e.name === 'Rectangle 10');
  assert.equal(rect.fill.value.toUpperCase(), linearBlend('4472C4', 0.5, false));
  assert.equal(rect.borderWidth, 0.75, 'an outline without w is 0.75 pt');
  assert.equal(doc.getSlide(2).elements.find((e) => e.name === 'Flipped Triangle 34').borderWidth, 0, 'a:ln with noFill stays without outline');
  const cell = doc.getSlide(4).elements.find((e) => e.type === 'table').data[0][0];
  assert.equal(cell.fillColor.toUpperCase(), linearBlend('4472C4', 0.4, true));
  assert.deepEqual(cell.margin, { t: 3.6, r: 7, b: 3.6, l: 7 });
  assert.ok(!MARKER_CHARS.test(cell.text));
});
