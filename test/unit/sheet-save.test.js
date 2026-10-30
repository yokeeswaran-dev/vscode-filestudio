// Surgical xlsx save (src/renderers/sheet.ts, sections XLSX PATCH and XLSX SAVE; contract §2.3, §4.2, §4.3, §4.7,
// §14 P1 WP2-WP3): cell / shared-string / style / workbook patchers, buildXlsx, verifyXlsx (V1-V4), checkPatchable.
// Workbooks are built here from raw XML parts (so every byte is known) and loaded fresh in every test.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const sample = (name) => fs.readFileSync(path.join(__dirname, '..', 'samples', name));

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_STRICT = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const STYLES = (P = '') =>
  `${DECL}<${P}styleSheet xmlns${P ? `:${P.slice(0, -1)}` : ''}="${NS_MAIN}">` +
  `<${P}numFmts count="1"><${P}numFmt numFmtId="164" formatCode="0.000"/></${P}numFmts>` +
  `<${P}fonts count="2"><${P}font><${P}sz val="11"/><${P}name val="Calibri"/></${P}font>` +
  `<${P}font><${P}b val="0"/><${P}sz val="11"/><${P}name val="Calibri"/></${P}font></${P}fonts>` +
  `<${P}fills count="2"><${P}fill><${P}patternFill patternType="none"/></${P}fill><${P}fill><${P}patternFill patternType="gray125"/></${P}fill></${P}fills>` +
  `<${P}borders count="1"><${P}border><${P}left/><${P}right/><${P}top/><${P}bottom/><${P}diagonal/></${P}border></${P}borders>` +
  `<${P}cellStyleXfs count="1"><${P}xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></${P}cellStyleXfs>` +
  `<${P}cellXfs count="3"><${P}xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<${P}xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<${P}xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></${P}cellXfs>` +
  `<${P}cellStyles count="1"><${P}cellStyle name="Normal" xfId="0" builtinId="0"/></${P}cellStyles></${P}styleSheet>`;

/** A worksheet part: `rows` is the inside of <sheetData> (null = <sheetData/>); `head` goes before <sheetData>. */
function sheetXml(rows, { P = '', head = '', tail = '', ns = NS_MAIN } = {}) {
  const xmlns = P ? `xmlns:${P.slice(0, -1)}="${ns}"` : `xmlns="${ns}"`;
  const data = rows === null ? `<${P}sheetData/>` : `<${P}sheetData>${rows}</${P}sheetData>`;
  return `${DECL}<${P}worksheet ${xmlns} xmlns:r="${NS_R}">${head}${data}${tail}</${P}worksheet>`;
}

/** Shared strings part with these (plain) strings. */
function sstXml(strings, P = '') {
  const xmlns = P ? `xmlns:${P.slice(0, -1)}="${NS_MAIN}"` : `xmlns="${NS_MAIN}"`;
  return `${DECL}<${P}sst ${xmlns} count="${strings.length + 1}" uniqueCount="${strings.length}">` +
    strings.map((s) => `<${P}si><${P}t>${s}</${P}t></${P}si>`).join('') + `</${P}sst>`;
}

/**
 * A small xlsx from raw parts. o: { sheets: [{ name, xml, sheetId? }], sst?: string | null (null = no sharedStrings
 * part), styles?: string, P?: '' | 'x:', calcPr?: string (inside workbook, default '<calcPr calcId="191029"/>'),
 * calcChain?: boolean, ns?: workbook namespace }.
 */
async function makeXlsx(o) {
  const P = o.P ?? '';
  const zip = new JSZip();
  const sheets = o.sheets;
  const overrides = [`<Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>`];
  const rels = [];
  sheets.forEach((sheet, i) => {
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheet.xml);
    overrides.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT}.worksheet+xml"/>`);
    rels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
  });
  zip.file('xl/styles.xml', o.styles ?? STYLES(P));
  overrides.push(`<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/>`);
  rels.push(`<Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>`);
  if (o.sst !== null) {
    zip.file('xl/sharedStrings.xml', o.sst ?? sstXml(['zero', 'one'], P));
    overrides.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.sharedStrings+xml"/>`);
    rels.push(`<Relationship Id="rIdT" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`);
  }
  if (o.calcChain) {
    zip.file('xl/calcChain.xml', `${DECL}<calcChain xmlns="${NS_MAIN}"><c r="A1" i="1"/></calcChain>`);
    overrides.push(`<Override PartName="/xl/calcChain.xml" ContentType="${CT}.calcChain+xml"/>`);
    rels.push(`<Relationship Id="rIdC" Type="${REL}/calcChain" Target="calcChain.xml"/>`);
  }
  zip.file('[Content_Types].xml', `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    `<Default Extension="xml" ContentType="application/xml"/>${overrides.join('')}</Types>`);
  zip.file('_rels/.rels', `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/_rels/workbook.xml.rels', `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
  const xmlns = P ? `xmlns:${P.slice(0, -1)}="${o.ns ?? NS_MAIN}"` : `xmlns="${o.ns ?? NS_MAIN}"`;
  zip.file('xl/workbook.xml', `${DECL}<${P}workbook ${xmlns} xmlns:r="${NS_R}"><${P}sheets>` +
    sheets.map((s, i) => `<${P}sheet name="${s.name}" sheetId="${s.sheetId ?? i + 1}" r:id="rId${i + 1}"/>`).join('') +
    `</${P}sheets>${o.calcPr ?? `<${P}calcPr calcId="191029"/>`}</${P}workbook>`);
  zip.file('docProps/app.xml', `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Test</Application></Properties>`);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/** Text of one part. */
async function part(bytes, name) {
  const zip = S.parseZip(bytes);
  const entry = zip.get(name);
  return entry ? (await S.readEntry(zip, entry)).toString('utf8') : undefined;
}

/** <sheetData> inside of sheet1. */
async function sheetData(bytes, n = 1) {
  const text = await part(bytes, `xl/worksheets/sheet${n}.xml`);
  return /<(?:x:)?sheetData>([\s\S]*)<\/(?:x:)?sheetData>/.exec(text)?.[1] ?? '';
}

/** Names of the entries whose raw bytes differ (or that were added / removed). */
function changedEntries(a, b) {
  const za = S.parseZip(a);
  const zb = S.parseZip(b);
  const out = [];
  for (const e of za.entries) {
    const n = zb.get(e.name);
    if (!n) out.push(`-${e.name}`);
    else if (!za.buf.subarray(e.dataStart, e.dataStart + e.csize).equals(zb.buf.subarray(n.dataStart, n.dataStart + n.csize))) out.push(e.name);
  }
  for (const e of zb.entries) if (!za.get(e.name)) out.push(`+${e.name}`);
  return out.sort();
}

const cells = (sheet, list) => ({ kind: 'cells', sheet, cells: list });
const num = (v) => ({ t: 'n', v });
const str = (v) => ({ t: 's', v });

/** buildXlsx + verifyXlsx with the given expectations (verify must pass). */
async function save(base, ops, expect = [], o) {
  const { bytes, report } = await S.buildXlsx(base, ops);
  await S.verifyXlsx(base, bytes, ops, expect, o);
  return { bytes, report };
}

// ----- sheet cells -----

test('cells: an existing cell gets a new value and keeps its style; other cells and parts stay byte-identical', async () => {
  const rows = '<row r="1" spans="1:2"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }, { name: 'S2', xml: sheetXml('<row r="1"><c r="A1"><v>9</v></c></row>') }] });
  const { bytes, report } = await save(base, [cells(1, [[0, 0, { value: num(5) }]])], [{ sheet: 1, r: 0, c: 0, value: num(5), styleIndex: 1 }]);
  assert.equal(await sheetData(bytes), '<row r="1" spans="1:2"><c r="A1" s="1"><v>5</v></c><c r="B1"><v>2</v></c></row>');
  assert.deepEqual(changedEntries(base, bytes), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  assert.deepEqual(report.changedParts.sort(), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  assert.match(await part(bytes, 'xl/workbook.xml'), /<calcPr calcId="191029" fullCalcOnLoad="1"\/>/);
});

test('cells: a missing cell is inserted in column order; new rows go to the start, middle and end', async () => {
  const rows = '<row r="2"><c r="A2"><v>1</v></c><c r="C2"><v>3</v></c></row><row r="4"><c r="A4"><v>4</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const ops = [cells(1, [[1, 1, { value: num(2) }], [0, 0, { value: num(10) }], [2, 0, { value: num(30) }], [5, 3, { value: num(60) }], [1, 4, { value: num(5) }]])];
  const expect = [[1, 1, 2], [0, 0, 10], [2, 0, 30], [5, 3, 60], [1, 4, 5]].map(([r, c, v]) => ({ sheet: 1, r, c, value: num(v) }));
  const { bytes } = await save(base, ops, expect);
  assert.equal(await sheetData(bytes),
    '<row r="1"><c r="A1"><v>10</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="E2"><v>5</v></c></row>' +
    '<row r="3"><c r="A3"><v>30</v></c></row>' +
    '<row r="4"><c r="A4"><v>4</v></c></row>' +
    '<row r="6"><c r="D6"><v>60</v></c></row>');
});

test('cells: a self-closing <row/> and an empty <sheetData/> take new cells', async () => {
  const a = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="2" spans="1:3" ht="20" customHeight="1"/>') }] });
  const { bytes } = await save(a, [cells(1, [[1, 1, { value: num(7) }]])], [{ sheet: 1, r: 1, c: 1, value: num(7) }]);
  assert.equal(await sheetData(bytes), '<row r="2" spans="1:3" ht="20" customHeight="1"><c r="B2"><v>7</v></c></row>');
  const b = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(null) }] });
  const res = await save(b, [cells(1, [[0, 0, { value: num(1) }]])], [{ sheet: 1, r: 0, c: 0, value: num(1) }]);
  assert.match(await part(res.bytes, 'xl/worksheets/sheet1.xml'), /<sheetData><row r="1"><c r="A1"><v>1<\/v><\/c><\/row><\/sheetData>/);
});

test('cells: x:-prefixed parts are patched with the same prefix (cells, shared strings, styles)', async () => {
  const P = 'x:';
  const base = await makeXlsx({ P, sheets: [{ name: 'S1', xml: sheetXml('<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>', { P }) }] });
  const ops = [cells(1, [[0, 0, { value: str('text') }], [0, 1, { formula: 'A1&"!"', cached: str('text!') }], [0, 2, { value: num(1), style: [{ bold: true }] }]])];
  const { bytes } = await save(base, ops, [
    { sheet: 1, r: 0, c: 0, value: str('text') },
    { sheet: 1, r: 0, c: 1, formula: 'A1&"!"', value: str('text!') },
    { sheet: 1, r: 0, c: 2, value: num(1), styleIndex: 3 },
  ], { deep: false });
  assert.equal(await sheetData(bytes),
    '<x:row r="1"><x:c r="A1" t="s"><x:v>2</x:v></x:c><x:c r="B1" t="str"><x:f>A1&amp;"!"</x:f><x:v>text!</x:v></x:c>' +
    '<x:c r="C1" s="3"><x:v>1</x:v></x:c></x:row>');
  assert.match(await part(bytes, 'xl/sharedStrings.xml'), /<x:si><x:t>text<\/x:t><\/x:si><\/x:sst>$/);
  assert.match(await part(bytes, 'xl/styles.xml'), /<x:font><x:b\/><x:sz val="11"\/><x:name val="Calibri"\/><\/x:font><\/x:fonts>/);
});

test('cells: clear keeps the style and drops value and formula; an empty unstyled cell is dropped', async () => {
  const rows = '<row r="1"><c r="A1" s="1"><f>B1</f><v>3</v></c><c r="B1" t="s"><v>0</v></c><c r="C1"><v>1</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const ops = [cells(1, [[0, 0, { value: null }], [0, 1, { value: null }]])];
  const { bytes } = await save(base, ops, [{ sheet: 1, r: 0, c: 0, value: null, styleIndex: 1 }, { sheet: 1, r: 0, c: 1, value: null }]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1" s="1"/><c r="C1"><v>1</v></c></row>');
});

test('cells: formulas are written in file syntax with their cached value (number, text, error, boolean, none)', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const ops = [cells(1, [
    [1, 0, { formula: 'SUM(A1:A1)', cached: num(1) }],
    [1, 1, { formula: 'IF(A1<2,"a<b","c")', cached: str('a<b') }],
    [1, 2, { formula: '1/0', cached: { t: 'e', v: '#DIV/0!' } }],
    [1, 3, { formula: 'A1=1', cached: { t: 'b', v: true } }],
    [1, 4, { formula: '_xlfn.CONCAT(A1,"x")', cached: null }],
  ])];
  const { bytes } = await save(base, ops, [
    { sheet: 1, r: 1, c: 0, formula: 'SUM(A1:A1)', value: num(1) },
    { sheet: 1, r: 1, c: 1, formula: 'IF(A1<2,"a<b","c")', value: str('a<b') },
    { sheet: 1, r: 1, c: 2, formula: '1/0', value: { t: 'e', v: '#DIV/0!' } },
    { sheet: 1, r: 1, c: 3, formula: 'A1=1', value: { t: 'b', v: true } },
    { sheet: 1, r: 1, c: 4, formula: '_xlfn.CONCAT(A1,"x")', value: null },
  ]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1"><v>1</v></c></row><row r="2">' +
    '<c r="A2"><f>SUM(A1:A1)</f><v>1</v></c>' +
    '<c r="B2" t="str"><f>IF(A1&lt;2,"a&lt;b","c")</f><v>a&lt;b</v></c>' +
    '<c r="C2" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
    '<c r="D2" t="b"><f>A1=1</f><v>1</v></c>' +
    '<c r="E2"><f>_xlfn.CONCAT(A1,"x")</f></c></row>');
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { formula: '=A1' }]])]), S.PatchError);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { formula: 'A1', value: num(1) }]])]), S.PatchError);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: num(Infinity) }]])]), S.PatchError);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: { t: 'e', v: '#NOPE!' } }]])]), S.PatchError);
});

test('cells: a cached-only patch keeps the <f> element bytes (shared formula clone attributes too)', async () => {
  const rows = '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>1</v></c></row><row r="2"><c r="A2" s="1"><f t="shared" si="0"/><v>2</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const { bytes } = await save(base, [cells(1, [[0, 0, { cached: num(9) }], [1, 0, { cached: str('x') }]])], [
    { sheet: 1, r: 0, c: 0, formula: 'B1+1', value: num(9) },
    { sheet: 1, r: 1, c: 0, formula: 'B2+1', value: str('x'), styleIndex: 1 },
  ]);
  assert.equal(await sheetData(bytes),
    '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>9</v></c></row><row r="2"><c r="A2" s="1" t="str"><f t="shared" si="0"/><v>x</v></c></row>');
  // A clone has no formula text of its own (V3 cannot compare it): V4 compares the expanded formula after reload.
  const ops = [cells(1, [[1, 0, { cached: str('x') }]])];
  const wrong = [{ sheet: 1, r: 1, c: 0, formula: 'B9+1', value: str('x') }];
  const built = await S.buildXlsx(base, ops);
  await S.verifyXlsx(base, built.bytes, ops, wrong, { deep: false });
  await assert.rejects(S.verifyXlsx(base, built.bytes, ops, wrong, { deep: true }), (err) => err instanceof S.SaveVerifyError && err.step === 'V4');
  await assert.rejects(S.buildXlsx(base, [cells(1, [[2, 0, { cached: num(1) }]])]), /no formula/);
});

test('cells: rows without a number, array formula ranges and cells with pictures (vm) are refused', async () => {
  const noR = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row><c r="A1"><v>1</v></c></row>') }] });
  await assert.rejects(S.buildXlsx(noR, [cells(1, [[0, 0, { value: num(2) }]])]), /without a row number/);
  const array = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>4</v></c></row>') }] });
  await assert.rejects(S.buildXlsx(array, [cells(1, [[1, 0, { value: num(1) }]])]), /array formula \(A1:A2\)/);
  await assert.rejects(S.buildXlsx(array, [cells(1, [[0, 0, { value: null }]])]), /array formula/);
  await save(array, [cells(1, [[1, 0, { style: [{ bold: true }] }]])], [{ sheet: 1, r: 1, c: 0, styleIndex: 3 }]);
  const spill = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c></row>') }] });
  await assert.rejects(S.buildXlsx(spill, [cells(1, [[2, 0, { value: num(1) }]])]), /array formula/);
  const vm = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c></row>') }] });
  await assert.rejects(S.buildXlsx(vm, [cells(1, [[0, 0, { value: num(1) }]])]), /picture or linked data type/);
});

test('cells: a shared-formula master change is refused unless every clone gets its own content (D-068)', async () => {
  const rows = '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row>' +
    '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>6</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: num(1) }]])]), /shared formula that A2 still uses/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { formula: 'B1*3' }], [1, 0, { formula: 'B2*2' }]])]), /A3 still uses/);
  const ops = [cells(1, [[0, 0, { value: num(1) }], [1, 0, { formula: 'B2*2' }], [2, 0, { formula: 'B3*2' }]])];
  const { bytes } = await save(base, ops, [
    { sheet: 1, r: 0, c: 0, value: num(1) },
    { sheet: 1, r: 1, c: 0, formula: 'B2*2' },
    { sheet: 1, r: 2, c: 0, formula: 'B3*2' },
  ]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><f>B2*2</f></c></row><row r="3"><c r="A3"><f>B3*2</f></c></row>');
});

test('cells: patches of one cell apply in order; positions outside the sheet are refused', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const ops = [cells(1, [[0, 0, { value: num(2) }]]), cells(1, [[0, 0, { style: [{ italic: true }] }]]), cells(1, [[0, 0, { value: num(3) }]])];
  const { bytes } = await save(base, ops, [{ sheet: 1, r: 0, c: 0, value: num(3), styleIndex: 3 }]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1" s="3"><v>3</v></c></row>');
  await assert.rejects(S.buildXlsx(base, [cells(1, [[1048576, 0, { value: num(1) }]])]), /invalid cell position/);
  await assert.rejects(S.buildXlsx(base, [cells(9, [[0, 0, { value: num(1) }]])]), /no sheet with id 9/);
  await assert.rejects(S.buildXlsx(base, [{ kind: 'insertRows', sheet: 1, at: 0, count: 1 }]), /not supported yet/);
});

test('cells: a new cell takes the style of its row (customFormat) or column', async () => {
  const xml = sheetXml('<row r="1" s="2" customFormat="1"><c r="A1"><v>1</v></c></row>', { head: '<cols><col min="2" max="3" width="9" style="1" customWidth="1"/></cols>' });
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml }] });
  const { bytes } = await save(base, [cells(1, [[0, 4, { value: num(5) }], [1, 1, { value: num(6) }]])], [
    { sheet: 1, r: 0, c: 4, value: num(5), styleIndex: 2 },
    { sheet: 1, r: 1, c: 1, value: num(6), styleIndex: 1 },
  ]);
  assert.equal(await sheetData(bytes), '<row r="1" s="2" customFormat="1"><c r="A1"><v>1</v></c><c r="E1" s="2"><v>5</v></c></row><row r="2"><c r="B2" s="1"><v>6</v></c></row>');
});

// ----- shared strings -----

test('sharedStrings: escapes, xml:space, _xHHHH_ (control characters, U+FFFE, literal escapes), counts, dedup', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" t="s"><v>0</v></c></row>') }] });
  const texts = ['<&>"', ' lead', 'trail ', 'a\tb', 'line1\nline2', '\u0001ctl', 'lit _x0041_ x', 'non￾char', 'cr\rx', 'ümlaut 😀', '<&>"'];
  const ops = [cells(1, texts.map((t, i) => [1, i, { value: str(t) }]))];
  const { bytes } = await save(base, ops, texts.map((t, i) => ({ sheet: 1, r: 1, c: i, value: str(t) })), { deep: true });
  const sst = await part(bytes, 'xl/sharedStrings.xml');
  assert.match(sst, /count="13" uniqueCount="12"/);
  const added = sst.slice(sst.indexOf('<si><t>one</t></si>') + '<si><t>one</t></si>'.length, sst.indexOf('</sst>'));
  assert.equal(added,
    '<si><t>&lt;&amp;&gt;"</t></si><si><t xml:space="preserve"> lead</t></si><si><t xml:space="preserve">trail </t></si>' +
    '<si><t xml:space="preserve">a\tb</t></si><si><t xml:space="preserve">line1\nline2</t></si><si><t>_x0001_ctl</t></si>' +
    '<si><t>lit _x005F_x0041_ x</t></si><si><t>non_xFFFE_char</t></si><si><t>cr_x000D_x</t></si><si><t>ümlaut 😀</t></si>');
  assert.match(await sheetData(bytes), /<c r="A2" t="s"><v>2<\/v><\/c>.*<c r="K2" t="s"><v>2<\/v><\/c>/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: str('lone \uD800 surrogate') }]])]), /invalid character/);
});

test('sharedStrings: a workbook without a sharedStrings part gets inline strings', async () => {
  const base = await makeXlsx({ sst: null, sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const { bytes } = await save(base, [cells(1, [[0, 1, { value: str(' inline <x> ') }]])], [{ sheet: 1, r: 0, c: 1, value: str(' inline <x> ') }]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t xml:space="preserve"> inline &lt;x&gt; </t></is></c></row>');
  assert.deepEqual(changedEntries(base, bytes), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
});

// ----- styles -----

test('styles: derived xfs are appended and deduplicated; existing xfs, fonts and fills stay byte-identical', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1"><v>2</v></c></row>') }] });
  const ops = [cells(1, [
    [0, 0, { style: [{ bold: true }] }], // xf 2 has <b val="0"/>: the derived font has <b/> instead
    [0, 1, { style: [{ bold: true }] }],
    [0, 2, { value: num(1), style: [{ bold: true }] }], // same base (0) + same delta as B1: same xf
    [0, 3, { value: num(1), style: [{ numFmt: '0.0000' }] }],
    [0, 4, { value: num(1), style: [{ numFmt: '0.000' }] }],
    [0, 5, { value: num(1), style: [{ numFmt: '0.00' }] }],
    [0, 6, { value: num(1), style: [{ fill: 'FFFFFF00', fontColor: 'FFFF0000' }, { fill: null }] }],
    [0, 7, { value: num(1), style: [{ border: { top: { style: 'thin', color: 'FF0000FF' }, left: null }, hAlign: 'center', wrap: true, quotePrefix: true }] }],
  ])];
  // A1 (xf 2, font <b val="0"/>) + bold and B1/C1 (xf 0) + bold give the same xf 3; '0.000' is the existing xf 1.
  const { bytes } = await save(base, ops, [
    { sheet: 1, r: 0, c: 0, styleIndex: 3 }, { sheet: 1, r: 0, c: 1, styleIndex: 3 }, { sheet: 1, r: 0, c: 2, styleIndex: 3 },
    { sheet: 1, r: 0, c: 3, styleIndex: 4 }, { sheet: 1, r: 0, c: 4, styleIndex: 1 }, { sheet: 1, r: 0, c: 5, styleIndex: 5 },
    { sheet: 1, r: 0, c: 6, styleIndex: 6 }, { sheet: 1, r: 0, c: 7, styleIndex: 7 },
  ]);
  const before = await part(base, 'xl/styles.xml');
  const after = await part(bytes, 'xl/styles.xml');
  const section = (text, name) => /<(\w+) count="\d+">([\s\S]*?)<\/\1>/g.exec(text.slice(text.indexOf(`<${name} `)))?.[2] ?? '';
  for (const name of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs']) {
    assert.ok(section(after, name).startsWith(section(before, name)), `${name}: existing items unchanged`);
  }
  assert.match(after, /<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"\/><numFmt numFmtId="165" formatCode="0.0000"\/><\/numFmts>/);
  assert.match(after, /<fonts count="4">/);
  const fonts = section(after, 'fonts');
  assert.ok(fonts.endsWith('<font><b/><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>'), fonts);
  const xfs = section(after, 'cellXfs');
  assert.ok(xfs.includes('<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'), xfs);
  assert.ok(xfs.includes('<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'), xfs);
  assert.ok(xfs.includes('<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'), xfs);
  assert.ok(xfs.includes('<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'), xfs);
  assert.match(after, /<borders count="2">.*<border><left\/><right\/><top style="thin"><color rgb="FF0000FF"\/><\/top><bottom\/><diagonal\/><\/border><\/borders>/s);
  assert.match(xfs, /quotePrefix="1"><alignment horizontal="center" wrapText="1"\/><\/xf>/);
  assert.match(after, /<cellXfs count="8">/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { style: [{ fill: 'red' }] }]])]), /invalid fill/);
});

test('styles: an empty self-closing <numFmts/> and a styles part without <numFmts> take a new number format', async () => {
  const rows = '<row r="1"><c r="A1"><v>1</v></c></row>';
  const op = [cells(1, [[0, 0, { style: [{ numFmt: '0.0%' }] }]])];
  // xf 1 still names numFmtId 164 without a <numFmt> for it: the new format must not reuse 164 (those cells would change).
  const selfClosing = await makeXlsx({ styles: STYLES().replace(/<numFmts count="1">.*<\/numFmts>/, '<numFmts count="0"/>'), sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const a = await save(selfClosing, op, [{ sheet: 1, r: 0, c: 0, styleIndex: 3 }]);
  assert.match(await part(a.bytes, 'xl/styles.xml'), /<numFmts count="1"><numFmt numFmtId="165" formatCode="0.0%"\/><\/numFmts><fonts/);
  const none = await makeXlsx({ styles: STYLES().replace(/<numFmts count="1">.*<\/numFmts>/, '').replace('numFmtId="164"', 'numFmtId="0"'), sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const b = await save(none, op, [{ sheet: 1, r: 0, c: 0, styleIndex: 3 }]);
  assert.match(await part(b.bytes, 'xl/styles.xml'), /<styleSheet xmlns="[^"]+"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"\/><\/numFmts><fonts/);
});

// ----- workbook -----

test('workbook: fullCalcOnLoad is added (other calcPr attributes kept, or calcPr inserted in schema order)', async () => {
  const rows = '<row r="1"><c r="A1"><v>1</v></c></row>';
  const kept = await makeXlsx({ calcPr: '<calcPr calcId="191029" iterate="1" iterateCount="50"/>', sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const a = await save(kept, [cells(1, [[0, 0, { value: num(2) }]])]);
  assert.match(await part(a.bytes, 'xl/workbook.xml'), /<calcPr calcId="191029" iterate="1" iterateCount="50" fullCalcOnLoad="1"\/>/);
  const off = await makeXlsx({ calcPr: '<calcPr fullCalcOnLoad="0"/>', sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  assert.match(await part((await save(off, [cells(1, [[0, 0, { value: num(2) }]])])).bytes, 'xl/workbook.xml'), /<calcPr fullCalcOnLoad="1"\/>/);
  const missing = await makeXlsx({ calcPr: '<extLst><ext uri="{x}"/></extLst>', sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  assert.match(await part((await save(missing, [cells(1, [[0, 0, { value: num(2) }]])])).bytes, 'xl/workbook.xml'), /<\/sheets><calcPr fullCalcOnLoad="1"\/><extLst>/);
});

test('workbook: calcChain is removed with its relationship and content type', async () => {
  const base = await makeXlsx({ calcChain: true, sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>') }] });
  const { bytes, report } = await save(base, [cells(1, [[0, 0, { value: num(5) }]])], [{ sheet: 1, r: 0, c: 0, value: num(5) }]);
  assert.deepEqual(report.removedParts, ['xl/calcChain.xml']);
  assert.deepEqual(changedEntries(base, bytes), ['-xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  assert.doesNotMatch(await part(bytes, '[Content_Types].xml'), /calcChain/);
  assert.doesNotMatch(await part(bytes, 'xl/_rels/workbook.xml.rels'), /calcChain/);
  assert.match(await part(bytes, 'xl/_rels/workbook.xml.rels'), /rIdS.*rIdT/);
});

// ----- buildXlsx: no-op -----

test('buildXlsx: no ops, or ops that change nothing, return the original bytes', async () => {
  const base = await makeXlsx({ calcChain: true, sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const none = await S.buildXlsx(base, []);
  assert.equal(none.bytes, base);
  assert.deepEqual(none.report.changedParts, []);
  const same = await save(base, [cells(1, [[0, 0, { value: num(1) }]]), cells(1, [[4, 4, { value: null }]])]);
  assert.ok(Buffer.from(same.bytes).equals(base));
  const notZip = Buffer.from('not a zip');
  assert.equal((await S.buildXlsx(notZip, [])).bytes, notZip, 'Save As of a clean unpatchable file writes it as it is');
  const style = sample('styled.xlsx');
  assert.ok(Buffer.from((await S.buildXlsx(style, [])).bytes).equals(style));
});

test('buildXlsx: the yield hook runs and can cancel', async () => {
  const rows = Array.from({ length: 50 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join('');
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  let calls = 0;
  await S.buildXlsx(base, [cells(1, [[60, 0, { value: num(1) }]])], { yieldEvery: 5, onYield: () => calls++ });
  assert.ok(calls >= 10, `yielded ${calls} times`);
  const cancel = new Error('cancelled');
  await assert.rejects(S.buildXlsx(base, [cells(1, [[60, 0, { value: num(1) }]])], { yieldEvery: 5, onYield: () => { throw cancel; } }), (err) => err === cancel);
});

// ----- verifyXlsx -----

test('verifyXlsx: a flipped byte, a dropped part, a broken fragment, a stale reference or a wrong cell is caught', async () => {
  const base = await makeXlsx({ calcChain: true, sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const ops = [cells(1, [[0, 0, { value: str('new') }]])];
  const { bytes } = await S.buildXlsx(base, ops);
  const good = [{ sheet: 1, r: 0, c: 0, value: str('new') }];
  await S.verifyXlsx(base, bytes, ops, good);
  const step = (name) => (err) => err instanceof S.SaveVerifyError && err.step === name;
  const zip = S.parseZip(bytes);
  // V1: a flipped byte in a rewritten part (inflate / CRC) and in an unchanged part
  for (const name of ['xl/worksheets/sheet1.xml', 'docProps/app.xml']) {
    const bad = Buffer.from(bytes);
    const e = zip.get(name);
    bad[e.dataStart + Math.floor(e.csize / 2)] ^= 0x55;
    await assert.rejects(S.verifyXlsx(base, bad, ops, good), step('V1'), name);
  }
  // V1: a dropped part; an extra part
  await assert.rejects(S.verifyXlsx(base, (await S.writeZip(zip, new Map([['docProps/app.xml', null]]))).bytes, ops, good), step('V1'));
  await assert.rejects(S.verifyXlsx(base, (await S.writeZip(zip, new Map(), { added: [{ name: 'x.xml', parts: [Buffer.from('<x/>')] }] })).bytes, ops, good), step('V1'));
  // V1: the removed calcChain is still referenced
  const ct = await part(base, '[Content_Types].xml');
  await assert.rejects(S.verifyXlsx(base, (await S.writeZip(zip, new Map([['[Content_Types].xml', [Buffer.from(ct)]]]))).bytes, ops, good), /still has a content type/);
  // V2: a broken fragment in the sheet part
  const sheet = await part(bytes, 'xl/worksheets/sheet1.xml');
  const broken = (await S.writeZip(zip, new Map([['xl/worksheets/sheet1.xml', [Buffer.from(sheet.replace('</c>', '</x>'))]]]))).bytes;
  await assert.rejects(S.verifyXlsx(base, broken, ops, good), step('V2'));
  // V3: a cell that does not hold what the display model shows
  await assert.rejects(S.verifyXlsx(base, bytes, ops, [{ sheet: 1, r: 0, c: 0, value: str('other') }]), step('V3'));
  await assert.rejects(S.verifyXlsx(base, bytes, ops, [{ sheet: 1, r: 0, c: 0, value: str('new'), styleIndex: 1 }]), step('V3'));
  await assert.rejects(S.verifyXlsx(base, bytes, ops, [{ sheet: 1, r: 0, c: 0, formula: 'A2' }]), step('V3'));
  await assert.rejects(S.verifyXlsx(base, bytes, ops, [{ sheet: 1, r: 5, c: 0, value: num(1) }]), step('V3'));
  // A part that the ops cannot change must stay byte-identical even when it is valid XML
  const styles = await part(bytes, 'xl/styles.xml');
  const restyled = (await S.writeZip(zip, new Map([['docProps/app.xml', [Buffer.from(styles)]]]))).bytes;
  await assert.rejects(S.verifyXlsx(base, restyled, ops, good), /must stay byte-identical/);
  // Unchanged bytes for no ops pass; any change fails
  await S.verifyXlsx(base, base, [], []);
  await assert.rejects(S.verifyXlsx(base, bytes, [], []), step('V1'));
});

test('verifyXlsx: V4 reloads the file with the loader and compares the edited cells (also dates, formulas, booleans)', async () => {
  const xml = sheetXml('<row r="1"><c r="A1" s="1"><v>1</v></c></row>');
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml }] });
  const ops = [cells(1, [[0, 0, { value: num(45000.123456789) }], [0, 1, { formula: '_xlfn.CONCAT("a","b")', cached: str('ab') }], [0, 2, { value: { t: 'b', v: false } }], [0, 3, { value: { t: 'e', v: '#N/A' } }]])];
  const expect = [
    { sheet: 1, r: 0, c: 0, value: num(45000.123456789) },
    { sheet: 1, r: 0, c: 1, formula: '_xlfn.CONCAT("a","b")', value: str('ab') },
    { sheet: 1, r: 0, c: 2, value: { t: 'b', v: false } },
    { sheet: 1, r: 0, c: 3, value: { t: 'e', v: '#N/A' } },
  ];
  const { bytes } = await S.buildXlsx(base, ops);
  await S.verifyXlsx(base, bytes, ops, expect, { deep: true });
});

test('verifyXlsx: a big cell-patched sheet part is checked row by row; a change outside the edited rows is caught', async () => {
  const N = 450000;
  const chunks = [];
  for (let r = 1; r <= N; r++) chunks.push(`<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}" t="s"><v>0</v></c></row>`);
  const big = Buffer.from(sheetXml(chunks.join('')), 'utf8');
  assert.ok(big.length > 32 << 20, `the sheet part has ${big.length} bytes`);
  const small = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const base = (await S.writeZip(S.parseZip(small), new Map([['xl/worksheets/sheet1.xml', [big]]]))).bytes;
  const ops = [cells(1, [[9, 0, { value: num(-1) }], [N - 1, 2, { value: str('end') }], [N + 1, 0, { value: num(7) }]])];
  const expect = [{ sheet: 1, r: 9, c: 0, value: num(-1) }, { sheet: 1, r: N - 1, c: 2, value: str('end') }, { sheet: 1, r: N + 1, c: 0, value: num(7) }];
  const { bytes } = await S.buildXlsx(base, ops);
  await S.verifyXlsx(base, bytes, ops, expect, { deep: false });
  const zip = S.parseZip(bytes);
  const text = (await S.readEntry(zip, zip.get('xl/worksheets/sheet1.xml'))).toString('utf8');
  const tampered = text.replace('<c r="A20"><v>20</v>', '<c r="A20"><v>21</v>');
  assert.notEqual(tampered, text);
  const bad = (await S.writeZip(zip, new Map([['xl/worksheets/sheet1.xml', [Buffer.from(tampered)]]]))).bytes;
  await assert.rejects(S.verifyXlsx(base, bad, ops, expect, { deep: false }), /changed outside the edited rows/);
  const brokenRow = text.replace('<c r="A10"><v>-1</v></c>', '<c r="A10"><v>-1</v>');
  const bad2 = (await S.writeZip(zip, new Map([['xl/worksheets/sheet1.xml', [Buffer.from(brokenRow)]]]))).bytes;
  await assert.rejects(S.verifyXlsx(base, bad2, ops, expect, { deep: false }), (err) => err instanceof S.SaveVerifyError && err.step === 'V2');
});

// ----- checkPatchable -----

test('checkPatchable: hard read-only reasons (zip, Strict, rows out of order on ANY sheet, cells outside their row)', async () => {
  assert.deepEqual(await S.checkPatchable(sample('styled.xlsx')), { ok: true });
  const ok = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
  const check = async (xml2, o) => S.checkPatchable(await makeXlsx({ sheets: [{ name: 'S1', xml: ok }, { name: 'S2', xml: xml2 }], ...o }));
  assert.deepEqual(await check(ok), { ok: true });
  const outOfOrder = await check(sheetXml('<row r="3"><c r="A3"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c></row>'));
  assert.equal(outOfOrder.ok, false);
  assert.match(outOfOrder.reason, /the sheet "S2" has rows out of order \(row 2\)/);
  assert.match((await check(sheetXml('<row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>1</v></c></row>'))).reason, /cells out of order/);
  assert.match((await check(sheetXml('<row r="2"><c r="B3"><v>1</v></c></row>'))).reason, /cell B3 inside row 2/);
  assert.match((await check(sheetXml('<row><c r="A1"><v>1</v></c></row>'))).reason, /without a row number/);
  assert.match((await check(sheetXml('<c r="A1"><v>1</v></c><row r="2"/>'))).reason, /cells outside a row/);
  assert.match((await check(ok, { ns: NS_STRICT })).reason, /Strict Open XML/);
  assert.deepEqual(await check(sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>')), { ok: true });
  assert.deepEqual(await S.checkPatchable(await makeXlsx({ sheets: [{ name: 'S1', xml: ok }, { name: 'S2', xml: sheetXml('<row r="3"/><row r="2"/>') }] }), { scanSheets: false }), { ok: true });
  const notZip = await S.checkPatchable(Buffer.from('not a zip package'));
  assert.equal(notZip.ok, false);
  const cancel = new Error('cancelled');
  await assert.rejects(S.checkPatchable(sample('styled.xlsx'), { onYield: () => { throw cancel; } }), (err) => err === cancel);
});

// ----- whole workbooks -----

/** Every cell of every sheet as 'Sheet!r,c' → JSON of its CellData + resolved style, and each sheet's meta. */
async function snapshot(bytes) {
  const model = S.workbookToModel(await S.loadWorkbook(new Uint8Array(bytes)));
  const meta = model.getMeta();
  const styles = new Map(meta.styles.map((s, i) => [i, s]));
  const cellsOut = new Map();
  meta.sheets.forEach((sheet, index) => {
    const { rows, styles: added } = model.getRows(index, 0, sheet.rowCount);
    for (const [i, s] of added) styles.set(i, s);
    for (const row of rows) {
      for (const cell of row.cells) {
        const { s, ...rest } = cell;
        cellsOut.set(`${sheet.name}!${row.r},${cell.c}`, JSON.stringify({ ...rest, style: styles.get(s ?? 0) ?? {} }));
      }
    }
  });
  const sheets = meta.sheets.map(({ rowCount, colCount, rows, ...rest }) => JSON.stringify(rest));
  return { cells: cellsOut, sheets };
}

test('styled.xlsx: one value, one formula and one style edit per sheet change only the expected parts; the loader sees no other difference', async () => {
  const base = sample('styled.xlsx');
  const zip = S.parseZip(base);
  const wb = (await S.readEntry(zip, zip.get('xl/workbook.xml'))).toString();
  const ids = [...wb.matchAll(/<sheet sheetId="(\d+)" name="([^"]+)"/g)].map((m) => ({ id: Number(m[1]), name: m[2] }));
  assert.equal(ids.length, 7);
  const ops = [];
  const expect = [];
  const edited = new Set();
  for (const { id, name } of ids) {
    ops.push(cells(id, [
      [300, 0, { value: str(`edited ${name}`) }],
      [301, 1, { formula: 'A301&"!"', cached: str(`edited ${name}!`) }],
      [302, 2, { value: num(42), style: [{ bold: true, fill: 'FFFFFF00' }] }],
    ]));
    expect.push({ sheet: id, r: 300, c: 0, value: str(`edited ${name}`) }, { sheet: id, r: 301, c: 1, formula: 'A301&"!"', value: str(`edited ${name}!`) }, { sheet: id, r: 302, c: 2, value: num(42) });
    for (const [r, c] of [[300, 0], [301, 1], [302, 2]]) edited.add(`${name}!${r},${c}`);
  }
  const { bytes, report } = await save(base, ops, expect, { deep: true });
  assert.deepEqual(changedEntries(base, bytes), [
    'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/workbook.xml',
    ...[1, 2, 3, 4, 5, 6, 7].map((n) => `xl/worksheets/sheet${n}.xml`),
  ].sort());
  assert.equal(report.identicalParts, S.parseZip(base).entries.length - 10);
  const a = await snapshot(base);
  const b = await snapshot(bytes);
  assert.deepEqual(b.sheets, a.sheets);
  const diffs = [];
  for (const key of new Set([...a.cells.keys(), ...b.cells.keys()])) {
    if (edited.has(key)) continue;
    if (a.cells.get(key) !== b.cells.get(key)) diffs.push(`${key}: ${a.cells.get(key)} -> ${b.cells.get(key)}`);
  }
  assert.deepEqual(diffs.slice(0, 5), []);
  for (const key of edited) assert.ok(b.cells.has(key), `${key} is shown after reload`);
});

test('a generated workbook: edit + clear back to the original value saves the original cell XML again', async () => {
  const rows = '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" t="s"><v>1</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  const first = await save(base, [cells(1, [[0, 0, { value: num(2) }]])]);
  assert.notDeepEqual(changedEntries(base, first.bytes), []);
  const back = await save(base, [cells(1, [[0, 0, { value: num(2) }]]), cells(1, [[0, 0, { value: num(1) }]])], [{ sheet: 1, r: 0, c: 0, value: num(1), styleIndex: 1 }]);
  assert.equal(await sheetData(back.bytes), rows);
});

// ----- loader display fixes and lossy v2 (WP4) -----

/** Adds (or replaces) raw parts in a package; `null` removes a part. */
async function withParts(bytes, parts) {
  const zip = await JSZip.loadAsync(bytes);
  for (const [name, content] of Object.entries(parts)) {
    if (content === null) zip.remove(name);
    else zip.file(name, content);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/** Cells of one sheet as 'A1' → CellData + resolved style. */
async function shownCells(bytes, sheet = 0) {
  const model = S.workbookToModel(await S.loadWorkbook(new Uint8Array(bytes)));
  const meta = model.getMeta();
  const { rows, styles } = model.getRows(sheet, 0, meta.sheets[sheet].rowCount);
  const styleOf = new Map([...meta.styles.map((s, i) => [i, s]), ...styles]);
  const out = {};
  for (const row of rows) {
    for (const cell of row.cells) out[`${String.fromCharCode(65 + cell.c)}${row.r + 1}`] = { ...cell, style: styleOf.get(cell.s ?? 0) ?? {} };
  }
  return out;
}

const DISPLAY_STYLES =
  `${DECL}<styleSheet xmlns="${NS_MAIN}"><numFmts count="5">` +
  '<numFmt numFmtId="164" formatCode="#,##0.00\\ \\m"/><numFmt numFmtId="165" formatCode="0.0\\ \\k\\m"/>' +
  '<numFmt numFmtId="166" formatCode="0&quot; min&quot;"/><numFmt numFmtId="167" formatCode="yyyy\\-mm\\-dd"/>' +
  '<numFmt numFmtId="168" formatCode="[h]:mm"/></numFmts>' +
  '<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>' +
  '<font><b val="0"/><i val="false"/><strike val="0"/><sz val="11"/><name val="Calibri"/></font>' +
  '<font><b/><i val="1"/><strike val="true"/><sz val="11"/><name val="Calibri"/></font>' +
  '<font><b val="off"/><i val="FALSE"/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  [164, 165, 166, 167, 168].map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('') +
  [1, 2, 3].map((id) => `<xf numFmtId="0" fontId="${id}" fillId="0" borderId="0" xfId="0" applyFont="1"/>`).join('') +
  '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

test('loader (D-055): numbers in unit formats with a literal "m" stay exact numbers; real date formats stay dates', async () => {
  const rows =
    '<row r="1"><c r="A1" s="1"><v>123456789.5</v></c><c r="B1" s="1"><f>A1*2</f><v>246913579</v></c></row>' +
    '<row r="2"><c r="A2" s="1"><v>2.123456789</v></c><c r="B2" s="2"><v>2.5</v></c></row>' +
    '<row r="3"><c r="A3" s="3"><v>90</v></c><c r="B3" s="4"><v>45000</v></c></row>' +
    '<row r="4"><c r="A4" s="5"><v>1.5</v></c><c r="B4" s="1"><v>1E+300</v></c></row>';
  const bytes = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }], styles: DISPLAY_STYLES });
  const c = await shownCells(bytes);
  assert.deepEqual([c.A1.t, c.A1.v, c.A1.w], ['n', 123456789.5, '123,456,789.50 m']);
  assert.deepEqual([c.B1.t, c.B1.v, c.B1.w, c.B1.f], ['n', 246913579, '246,913,579.00 m', 'A1*2']);
  assert.deepEqual([c.A2.t, c.A2.v, c.A2.w], ['n', 2.123456789, '2.12 m'], 'no Date round trip (precision kept)');
  assert.deepEqual([c.B2.t, c.B2.v, c.B2.w], ['n', 2.5, '2.5 km']);
  assert.deepEqual([c.A3.t, c.A3.v, c.A3.w], ['n', 90, '90 min']);
  assert.deepEqual([c.B4.t, c.B4.v], ['n', 1e300], 'a value past the JS Date range is no error');
  // genuine date / time formats are unchanged
  assert.deepEqual([c.B3.t, c.B3.v, c.B3.w], ['d', 45000, '2023-03-15']);
  assert.deepEqual([c.A4.t, c.A4.v, c.A4.w], ['d', 1.5, '36:00']);
});

test('loader (D-055): <b val="0"/>, <i val="false"/>, <strike val="0"/> and "off" are not set; <b/> and val="1"/"true" are', async () => {
  const sst = `${DECL}<sst xmlns="${NS_MAIN}" count="2" uniqueCount="2"><si><t>plain</t></si>` +
    '<si><r><rPr><b val="0"/><sz val="11"/></rPr><t xml:space="preserve">not bold </t></r><r><rPr><b/><sz val="11"/></rPr><t>bold</t></r></si></sst>';
  const rows = '<row r="1"><c r="A1" s="6" t="s"><v>0</v></c><c r="B1" s="7" t="s"><v>0</v></c><c r="C1" s="8" t="s"><v>0</v></c>' +
    '<c r="D1" t="s"><v>1</v></c></row>';
  const bytes = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }], styles: DISPLAY_STYLES, sst });
  const c = await shownCells(bytes);
  assert.deepEqual(c.A1.style.font ?? {}, {}, 'all flags off');
  assert.deepEqual(c.B1.style.font, { bold: true, italic: true, strike: true });
  assert.deepEqual(c.C1.style.font ?? {}, {}, '"off" and "FALSE" read as off');
  assert.deepEqual(c.D1.rt.map((run) => [run.text, run.font?.bold ?? false]), [['not bold ', false], ['bold', true]]);
});

test('loader: loadedBytes is the very object passed in, unchanged', async () => {
  const data = new Uint8Array(sample('styled.xlsx'));
  const before = Buffer.from(data).toString('base64');
  const wb = await S.loadWorkbook(data);
  assert.equal(S.loadedBytes(wb), data);
  assert.equal(Buffer.from(data).toString('base64'), before);
});

test('lossy v2: data tables, query tables, connections, data model, rich values and OLE objects are reported', async () => {
  const plain = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  assert.deepEqual(await S.detectLossyFeatures(plain), []);
  const sheet = (rows, tail = '') => ({ 'xl/worksheets/sheet1.xml': sheetXml(rows, { tail }) });
  const cases = [
    [sheet('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="dataTable" ref="B1:B3" dt2D="0" dtr="0" r1="A1"/><v>2</v></c></row>'), ['What-if data tables']],
    [{ 'xl/queryTables/queryTable1.xml': '<queryTable/>' }, ['Query tables / data connections']],
    [{ 'xl/connections.xml': '<connections/>' }, ['Query tables / data connections']],
    [{ 'xl/model/item.data': 'x' }, ['Data model (Power Pivot)']],
    [sheet('<row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c></row>'), ['Pictures in cells / linked data types']],
    [{ 'xl/richData/rdrichvalue.xml': '<rvData/>' }, ['Pictures in cells / linked data types']],
    [sheet('<row r="1"><c r="A1"><v>1</v></c></row>', '<oleObjects><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId9"/></oleObjects>'), ['Embedded objects (OLE)']],
    [{ 'xl/embeddings/oleObject1.bin': 'x' }, ['Embedded objects (OLE)']],
    // not triggers: the words inside a value, and cm (a dynamic array) is not vm
    [sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>&lt;f t="dataTable"&gt; vm="1"</t></is></c><c r="B1" cm="1"><f t="array" ref="B1">A1</f><v>1</v></c></row>'), []],
  ];
  for (const [parts, expected] of cases) {
    assert.deepEqual(await S.detectLossyFeatures(await withParts(plain, parts)), expected, JSON.stringify(Object.keys(parts)));
  }
  // x:-prefixed sheet XML, several at once, and the fixed order (0.1.x labels first)
  const all = await withParts(plain, {
    'xl/charts/chart1.xml': '<c:chartSpace/>', 'xl/model/item.data': 'x', 'xl/connections.xml': '<connections/>',
    'xl/worksheets/sheet1.xml': sheetXml('<x:row r="1"><x:c r="A1" vm="2"><x:f t="dataTable" ref="A1:A2" r1="B1"/><x:v>1</x:v></x:c></x:row>', { P: 'x:' }),
  });
  assert.deepEqual(await S.detectLossyFeatures(all), ['Charts', 'What-if data tables', 'Query tables / data connections', 'Data model (Power Pivot)', 'Pictures in cells / linked data types']);
});

test('lossy v2 texts: banner, "Edit anyway" details per kind of content, hard read-only message', async () => {
  assert.equal(S.lossyBannerText([]), undefined);
  assert.equal(S.lossyBannerText(['Charts', 'Pivot tables']),
    'Nothing in this file will be removed. FileStudio does not show these objects and cannot update them: Charts, Pivot tables.');
  const ask = S.lossyEditAnyway(['Charts', 'Pivot tables', 'Macros (VBA project)', 'External links', 'Sparklines', 'What-if data tables', 'Shapes / text boxes', 'Pictures in cells / linked data types']);
  assert.equal(ask.title, 'This workbook has content that FileStudio cannot update. Edit it anyway?');
  assert.deepEqual(ask.details, [
    'Charts, Sparklines: show the old numbers until the file is opened in Excel.',
    'Pivot tables: show the old numbers until you refresh them in Excel.',
    'External links: linked values stay as they were last saved until Excel updates the links.',
    'What-if data tables: results stay old until Excel recalculates or refreshes them.',
    'Pictures in cells / linked data types: kept as they are; the cells that hold them cannot be changed in FileStudio.',
    'Macros (VBA project), Shapes / text boxes: kept as they are; FileStudio does not update them.',
    'Inserting or deleting rows and columns, sorting and sheet changes stay blocked for this file.',
    'Nothing in this file will be removed.',
  ]);
  const bad = sheetXml('<row r="2"><c r="A2"><v>1</v></c></row><row r="1"><c r="A1"><v>1</v></c></row>');
  const res = await S.checkPatchable(await makeXlsx({ sheets: [{ name: 'S1', xml: bad }] }));
  assert.equal(res.ok, false);
  assert.equal(S.hardReadOnlyText(res.reason), 'FileStudio cannot edit this file safely: the sheet "S1" has rows out of order (row 1). It is open read-only.');
});

// ----- generated and sample workbooks (WP7) -----
// test/samples/xlsx-save/*.xlsx are written by test/generate-samples.js (`// ----- xlsx save fixtures -----`) as raw
// Excel-style parts: objects (table, chart, picture, pivot, external link, sheetIds out of tab order), formulas (1904,
// shared / array / dynamic-array formulas, names, calcChain), comments (threaded + persons, notes, rich text, links),
// formatting (unit formats, <b val="0"/>, custom palette, CF, > 10k-cell validations, page breaks, row/column styles).

const SAMPLES = path.join(__dirname, '..', 'samples');
const SAVE_FIXTURES = ['objects.xlsx', 'formulas.xlsx', 'comments.xlsx', 'formatting.xlsx'].map((name) => `xlsx-save/${name}`);

/** Every .xlsx under test/samples except large.xlsx (generated on demand, not in Git). */
function sampleWorkbooks() {
  const out = [];
  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (/\.xlsx$/i.test(entry.name) && entry.name !== 'large.xlsx') out.push(path.relative(SAMPLES, full).split(path.sep).join('/'));
    }
  };
  walk(SAMPLES);
  return out.sort();
}

/** Attributes of every `<tag ...>` in a part ('r:id' keeps its prefix). */
function tagsOf(xml, tag) {
  return [...xml.matchAll(new RegExp(`<${tag}\\b([^>]*?)/?>`, 'g'))].map((m) =>
    Object.fromEntries([...m[1].matchAll(/([\w:]+)="([^"]*)"/g)].map((a) => [a[1], a[2]])));
}

/** Worksheets (sheetId, name, part), the styles, sharedStrings and calcChain parts, and the number of entries. */
async function packageParts(bytes) {
  const zip = S.parseZip(bytes);
  const text = async (name) => (await S.readEntry(zip, zip.get(name))).toString('utf8');
  const rels = tagsOf(await text('xl/_rels/workbook.xml.rels'), 'Relationship');
  const target = (rel) => (rel.Target.startsWith('/') ? rel.Target.slice(1) : `xl/${rel.Target}`);
  const sheets = tagsOf(await text('xl/workbook.xml'), 'sheet').map((s) => {
    const rel = rels.find((r) => r.Id === s['r:id']);
    return { id: Number(s.sheetId), name: s.name, part: rel && /\/worksheet$/.test(rel.Type) ? target(rel) : undefined };
  });
  const one = (type) => rels.filter((r) => r.Type.endsWith(`/${type}`)).map(target)[0];
  return { sheets, sst: one('sharedStrings'), styles: one('styles'), calcChain: one('calcChain'), count: zip.entries.length };
}

for (const name of sampleWorkbooks()) {
  test(`${name}: no-op save = same bytes; one value, one formula and one style edit per sheet change only the expected parts; the loader sees no other difference`, async () => {
    const base = sample(name);
    const parts = await packageParts(base);
    const worksheets = parts.sheets.filter((s) => s.part);
    assert.ok(worksheets.length > 0);
    assert.deepEqual(await S.checkPatchable(base), { ok: true });
    // no-op: nothing to save, and an edit that is undone (value, then cleared again) gives the original bytes
    assert.equal((await S.buildXlsx(base, [])).bytes, base);
    const a = await snapshot(base);
    const model = S.workbookToModel(await S.loadWorkbook(new Uint8Array(base)));
    const metaOf = new Map(model.workbook.worksheets.map((ws, i) => [ws.id, model.getMeta().sheets[i]]));
    const undo = worksheets.flatMap((s) => [
      cells(s.id, [[metaOf.get(s.id).rowCount, 0, { value: str('x') }]]),
      cells(s.id, [[metaOf.get(s.id).rowCount, 0, { value: null }]]),
    ]);
    assert.ok(Buffer.from((await S.buildXlsx(base, undo)).bytes).equals(base), 'edit + clear again = the original bytes');

    const ops = [];
    const expect = [];
    const edited = new Set();
    for (const s of worksheets) {
      const meta = metaOf.get(s.id);
      const r0 = meta.rowCount + 1;
      const text = `edited <${s.name}> & "x"`;
      ops.push(cells(s.id, [
        [r0, 0, { value: str(text) }],
        [r0 + 1, 1, { formula: 'SUM(1,2)', cached: num(3) }],
        [r0 + 2, 2, { value: num(42.5), style: [{ bold: true, fill: 'FFFFFF00' }] }],
        [0, 0, { style: [{ italic: true }] }],
      ]));
      expect.push(
        { sheet: s.id, r: r0, c: 0, value: str(text) },
        { sheet: s.id, r: r0 + 1, c: 1, formula: 'SUM(1,2)', value: num(3) },
        { sheet: s.id, r: r0 + 2, c: 2, value: num(42.5) },
      );
      for (const [r, c] of [[r0, 0], [r0 + 1, 1], [r0 + 2, 2], [0, 0]]) edited.add(`${meta.name}!${r},${c}`);
    }
    const { bytes, report } = await save(base, ops, expect, { deep: true });
    const expected = [...worksheets.map((s) => s.part), parts.styles, 'xl/workbook.xml'];
    if (parts.sst) expected.push(parts.sst);
    if (parts.calcChain) expected.push(`-${parts.calcChain}`, '[Content_Types].xml', 'xl/_rels/workbook.xml.rels');
    assert.deepEqual(changedEntries(base, bytes), expected.sort());
    assert.equal(report.identicalParts, parts.count - expected.length);
    const b = await snapshot(bytes);
    assert.deepEqual(b.sheets, a.sheets);
    const diffs = [];
    for (const key of new Set([...a.cells.keys(), ...b.cells.keys()])) {
      if (edited.has(key)) continue;
      if (a.cells.get(key) !== b.cells.get(key)) diffs.push(`${key}: ${a.cells.get(key)} -> ${b.cells.get(key)}`);
    }
    assert.deepEqual(diffs.slice(0, 5), []);
    for (const key of edited) assert.ok(b.cells.has(key), `${key} is shown after reload`);
  });
}

test('xlsx-save fixtures: all generated workbooks are there and report the expected lossy content', async () => {
  assert.deepEqual(sampleWorkbooks().filter((n) => n.startsWith('xlsx-save/')), [...SAVE_FIXTURES].sort());
  const lossy = {};
  for (const name of SAVE_FIXTURES) lossy[name] = await S.detectLossyFeatures(sample(name));
  assert.deepEqual(lossy, {
    'xlsx-save/objects.xlsx': ['Charts', 'Pivot tables', 'External links'],
    'xlsx-save/formulas.xlsx': [],
    'xlsx-save/comments.xlsx': [],
    'xlsx-save/formatting.xlsx': [],
  });
});

test('objects.xlsx: a value under a table, chart and pivot changes only its sheet and workbook.xml; sheetIds map to parts', async () => {
  const base = sample('xlsx-save/objects.xlsx');
  // Sales = sheetId 1 (sheet1.xml), Summary = sheetId 3 (sheet2.xml), Pivot = sheetId 2 (sheet3.xml)
  const ops = [cells(1, [[2, 2, { value: num(99) }]]), cells(3, [[7, 0, { value: num(7) }]])];
  const { bytes } = await save(base, ops, [{ sheet: 1, r: 2, c: 2, value: num(99) }, { sheet: 3, r: 7, c: 0, value: num(7) }], { deep: true });
  assert.deepEqual(changedEntries(base, bytes), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
  const wbBefore = await part(base, 'xl/workbook.xml');
  const wbAfter = await part(bytes, 'xl/workbook.xml');
  assert.equal(wbAfter, wbBefore.replace('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>'));
  const sales = await shownCells(bytes, 0);
  assert.equal(sales.C3.v, 99);
  assert.equal(sales.D3.f, 'C3*12.5', 'the shared-formula clone keeps its formula');
  assert.equal(sales.D3.v, 1187.5, 'and its cached value until Excel recalculates');
  assert.equal((await shownCells(bytes, 1)).A8.v, 7);
  assert.equal(await sheetData(bytes, 3), await sheetData(base, 3));
});

test('formulas.xlsx: an input edit drops calcChain only; array and dynamic-array ranges are refused; 1904 dates stay 1904', async () => {
  const base = sample('xlsx-save/formulas.xlsx');
  const { bytes } = await save(base, [cells(1, [[2, 0, { value: num(7) }]])], [{ sheet: 1, r: 2, c: 0, value: num(7) }], { deep: true });
  assert.deepEqual(changedEntries(base, bytes), ['-xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  assert.ok(!(await part(bytes, '[Content_Types].xml')).includes('calcChain'));
  assert.ok(!(await part(bytes, 'xl/_rels/workbook.xml.rels')).includes('calcChain'));
  const wb = await part(bytes, 'xl/workbook.xml');
  assert.ok(wb.includes('<workbookPr date1904="1"'));
  assert.ok(wb.includes('<definedName name="Rate" localSheetId="1">Dates!$B$7</definedName>'));
  const data = await sheetData(bytes, 1);
  assert.ok(data.includes('<c r="G2" cm="1"><f t="array" ref="G2:G6" aca="false">_xlfn._xlws.SORT(A2:A6,1,-1)</f><v>15</v></c>'));
  assert.ok(data.includes('<c r="A3"><v>7</v></c>'));
  for (const [r, c] of [[2, 4], [3, 6], [1, 4], [1, 6]]) {
    await assert.rejects(S.buildXlsx(base, [cells(1, [[r, c, { value: num(1) }]])]), (err) => err.name === 'PatchError', `${r},${c} is refused`);
  }
  // 1904: serial 43905 = 2024-03-16
  const before = await shownCells(base, 1);
  const dated = await save(base, [cells(2, [[1, 1, { value: num(43905) }]])], [{ sheet: 2, r: 1, c: 1, value: num(43905) }], { deep: true });
  const after = await shownCells(dated.bytes, 1);
  assert.equal(after.B2.t, 'd');
  assert.match(before.B2.w, /15/);
  assert.match(after.B2.w, /16/);
  assert.equal(after.B5.w, before.B5.w, 'the 1904 epoch cell is unchanged');
});

test('comments.xlsx: edits of cells with threaded comments and notes keep comments, persons, VML and links byte-identical', async () => {
  const base = sample('xlsx-save/comments.xlsx');
  const ops = [cells(1, [[1, 1, { value: str('Carla') }], [2, 3, { value: str('Done in April') }], [6, 0, { style: [{ bold: true }] }]])];
  const { bytes } = await save(base, ops, [{ sheet: 1, r: 1, c: 1, value: str('Carla') }, { sheet: 1, r: 2, c: 3, value: str('Done in April') }], { deep: true });
  assert.deepEqual(changedEntries(base, bytes), ['xl/sharedStrings.xml', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  const before = await shownCells(base, 0);
  const after = await shownCells(bytes, 0);
  assert.equal(after.B2.v, 'Carla');
  assert.equal(after.B2.note, before.B2.note);
  assert.match(after.B2.note, /^Ana Silva: Is this the right owner\?/);
  assert.equal(after.D3.v, 'Done in April');
  assert.equal(after.D3.note, before.D3.note);
  assert.equal(after.A7.link, before.A7.link);
  assert.equal(after.A7.style.font.bold, true);
  assert.equal(after.A8.link, before.A8.link);
});

test('formatting.xlsx: unit-format numbers stay numbers, <b val="0"/> stays not bold, new cells take row/column styles, rules and validations stay byte-identical', async () => {
  const base = sample('xlsx-save/formatting.xlsx');
  const ops = [cells(1, [
    [2, 1, { value: num(2500.25) }],
    [3, 1, { value: num(0.5) }],
    [8, 0, { style: [{ fill: 'FFFFFF00' }] }],
    [4, 4, { value: num(1) }],
    [13, 2, { value: str('row styled') }],
  ])];
  const { bytes } = await save(base, ops, [
    { sheet: 1, r: 2, c: 1, value: num(2500.25), styleIndex: 3 },
    { sheet: 1, r: 4, c: 4, value: num(1), styleIndex: 14 },
    { sheet: 1, r: 13, c: 2, value: str('row styled'), styleIndex: 13 },
  ], { deep: true });
  assert.deepEqual(changedEntries(base, bytes), ['xl/sharedStrings.xml', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  const after = await shownCells(bytes, 0);
  assert.deepEqual([after.B3.t, after.B3.v, after.B3.w], ['n', 2500.25, '2,500.25 m']);
  assert.deepEqual([after.B4.t, after.B4.v, after.B4.w], ['n', 0.5, '0.5 km']);
  assert.equal(after.A9.style.fill, '#FFFF00');
  assert.ok(!after.A9.style.font?.bold, '<b val="0"/> base font stays not bold');
  assert.ok(after.E5.style.border, 'a new cell takes its column style');
  assert.equal(after.C14.style.fill, after.A14.style.fill, 'a new cell takes its row style');
  // the custom palette, the "off" font flags and every existing xf are kept
  const styles = await part(bytes, 'xl/styles.xml');
  const stylesBase = await part(base, 'xl/styles.xml');
  const block = (xml, tag) => xml.slice(xml.indexOf(`<${tag}`), xml.indexOf(`</${tag}>`));
  assert.equal(block(styles, 'colors'), block(stylesBase, 'colors'));
  assert.ok(styles.includes('<b val="0"/>') && styles.includes('<i val="false"/>') && styles.includes('<strike val="0"/>'));
  const xfs = (xml) => block(xml, 'cellXfs').replace(/^<cellXfs count="\d+">/, '');
  assert.ok(xfs(styles).startsWith(xfs(stylesBase)), 'existing xfs are byte-identical');
  // merges, conditional formats, the > 10k-cell validations and the page breaks are copied as they were
  const tail = (xml) => xml.slice(xml.indexOf('</sheetData>'));
  assert.equal(tail(await part(bytes, 'xl/worksheets/sheet1.xml')), tail(await part(base, 'xl/worksheets/sheet1.xml')));
});

// ===== fix round 1 (review findings F4-F9, F13 and the shared-formula whitespace case) =====

/** makeXlsx + one table part on sheet 1 (ref `ref`, optional headerRowCount). */
async function withTable(rows, ref, headerRowCount) {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows, { tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' }) }] });
  const zip = await JSZip.loadAsync(base);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`);
  zip.file('xl/tables/table1.xml', `${DECL}<table xmlns="${NS_MAIN}" id="1" name="T1" displayName="T1" ref="${ref}"${headerRowCount === undefined ? '' : ` headerRowCount="${headerRowCount}"`}><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/></tableColumns></table>`);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
const TABLE_ROWS = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c><c r="B2"><v>1</v></c></row>';

test('F4: a content patch on a table header cell is refused; body cells, style-only patches and header-less tables are fine', async () => {
  const base = await withTable(TABLE_ROWS, 'A1:B2');
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: str('Renamed') }]])]), /table header cell/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 1, { value: num(3) }]])]), /table header cell/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: null }]])]), /table header cell/);
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 1, { formula: '1+1' }]])]), /table header cell/);
  await save(base, [cells(1, [[1, 1, { value: num(9) }]])], [{ sheet: 1, r: 1, c: 1, value: num(9) }]);
  await save(base, [cells(1, [[0, 0, { style: [{ bold: true }] }]])]);
  const noHeader = await withTable(TABLE_ROWS, 'A1:B2', 0);
  await save(noHeader, [cells(1, [[0, 0, { value: num(1) }]])], [{ sheet: 1, r: 0, c: 0, value: num(1) }]);
});

test('F5: V4 (synchronous reload) is decided by the unpacked sharedStrings size too, not only by the zip size', async () => {
  const small = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  assert.equal(S.deepVerifyAllowed(small), true);
  const zip = await JSZip.loadAsync(small);
  zip.file('xl/sharedStrings.xml', Buffer.alloc(40 << 20, 0x20)); // 40 MB of spaces: a tiny zip
  const huge = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  assert.ok(huge.length < S.VERIFY_DEEP_MAX_BYTES, 'the zip itself is small');
  assert.equal(S.deepVerifyAllowed(huge), false);
});

test('F6: the text "_x005F_x0041_" (looks like an escape of an escape) round-trips through the shared strings', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const text = '_x005F_x0041_';
  const { bytes } = await save(base, [cells(1, [[0, 0, { value: str(text) }]])], [{ sheet: 1, r: 0, c: 0, value: str(text) }]);
  assert.match(await part(bytes, 'xl/sharedStrings.xml'), /<t>_x005F_x005F_x005F_x0041_<\/t>/);
  const two = '_x0041_ and _x005F_';
  await save(base, [cells(1, [[0, 0, { value: str(two) }]])], [{ sheet: 1, r: 0, c: 0, value: str(two) }]);
});

test('F7: <dimension ref> grows when cells are written outside it (and is left alone when they fit)', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>', { head: '<dimension ref="A1:A3"/>' }) }] });
  const { bytes } = await save(base, [cells(1, [[4, 4, { value: num(5) }], [99, 25, { value: num(6) }]])], [{ sheet: 1, r: 4, c: 4, value: num(5) }, { sheet: 1, r: 99, c: 25, value: num(6) }]);
  assert.match(await part(bytes, 'xl/worksheets/sheet1.xml'), /<dimension ref="A1:Z100"\/>/);
  const inside = await save(base, [cells(1, [[1, 0, { value: num(2) }]])], [{ sheet: 1, r: 1, c: 0, value: num(2) }]);
  assert.match(await part(inside.bytes, 'xl/worksheets/sheet1.xml'), /<dimension ref="A1:A3"\/>/);
});

test('F8: V4 does not refuse a correct save for cells the loader shows differently (merged non-anchor cell, empty string)', async () => {
  const xml = sheetXml('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>', { tail: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' });
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml }] });
  await save(base, [cells(1, [[0, 1, { value: num(7) }]])], [{ sheet: 1, r: 0, c: 1, value: num(7) }], { deep: true });
  const plain = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  await save(plain, [cells(1, [[0, 0, { value: str('') }]])], [{ sheet: 1, r: 0, c: 0, value: str('') }], { deep: true });
});

test('F9: an XML comment or CDATA section inside a row, or in sharedStrings, makes the file hard read-only (checkPatchable)', async () => {
  const withComment = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><!-- x --><v>1</v></c></row>') }] });
  const a = await S.checkPatchable(withComment);
  assert.equal(a.ok, false);
  assert.match(a.reason, /comment or CDATA/);
  const withCdata = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" t="inlineStr"><is><t><![CDATA[<c r="B1"></c></row>]]></t></is></c></row>') }] });
  assert.equal((await S.checkPatchable(withCdata)).ok, false);
  const sst = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }], sst: sstXml(['a']).replace('<si>', '<!-- <si> --><si>') });
  const c = await S.checkPatchable(sst);
  assert.equal(c.ok, false);
  assert.match(c.reason, /shared strings/);
  const clean = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  assert.deepEqual(await S.checkPatchable(clean), { ok: true });
});

test('shared formula clones written as si = "0" (spaces around =) are found: the de-share check (D-068) still works', async () => {
  const rows = '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si = "0">B1*2</f><v>2</v></c></row>' +
    '<row r="2"><c r="A2"><f t="shared" si = "0"/><v>4</v></c></row>';
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml(rows) }] });
  await assert.rejects(S.buildXlsx(base, [cells(1, [[0, 0, { value: num(1) }]])]), /shared formula that A2 still uses/);
});

test('F13: every relationship id used by a generated fixture sheet (legacyDrawing, drawing, tablePart) exists in its rels with the right type', async () => {
  for (const name of ['comments.xlsx', 'objects.xlsx', 'formulas.xlsx', 'formatting.xlsx']) {
    const zip = await JSZip.loadAsync(sample(`xlsx-save/${name}`));
    for (const file of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      const xml = await zip.file(file).async('string');
      const rels = await zip.file(file.replace('worksheets/', 'worksheets/_rels/') + '.rels')?.async('string');
      for (const m of xml.matchAll(/<(legacyDrawing|drawing|tablePart)\b[^>]*r:id="([^"]+)"/g)) {
        const want = { legacyDrawing: 'vmlDrawing', drawing: 'drawing', tablePart: 'table' }[m[1]];
        const rel = new RegExp(`<Relationship\\b[^>]*Id="${m[2]}"[^>]*>`).exec(rels ?? '')?.[0];
        assert.ok(rel, `${name} ${file}: ${m[1]} ${m[2]} has no relationship`);
        assert.ok(rel.includes(`/${want}"`), `${name} ${file}: ${m[1]} ${m[2]} points to the wrong kind: ${rel}`);
      }
    }
  }
});

// ----- fix round 2: limits Excel enforces, and '>' inside an attribute value -----

test('R2: a formula Excel would remove (too long, unbalanced, unterminated string) and text over 32767 characters are refused', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }] });
  const refused = (patch) => assert.rejects(S.buildXlsx(base, [cells(1, [[0, 1, patch]])]), S.PatchError, JSON.stringify(patch).slice(0, 40));
  await refused({ formula: 'SUM(A1' });
  await refused({ formula: 'SUM(A1))' });
  await refused({ formula: '"abc' });
  await refused({ formula: `1${'+1'.repeat(4100)}` }); // 8201 characters
  await refused({ value: str('x'.repeat(32768)) });
  await refused({ value: { t: 's', v: 'y'.repeat(40000) } });
  await refused({ value: num(5e-324) });
  await refused({ formula: 'A1', value: str('z'.repeat(32768)) });
  // Still accepted: strings with brackets / parentheses inside quotes, the exact limits.
  await S.buildXlsx(base, [cells(1, [[0, 1, { formula: 'IF(A1>1,"(",")")' }]])]);
  await S.buildXlsx(base, [cells(1, [[0, 1, { formula: `1${'+1'.repeat(4095)}` }]])]); // 8191 characters
  await S.buildXlsx(base, [cells(1, [[0, 1, { value: str('x'.repeat(32767)) }]])]);
  await S.buildXlsx(base, [cells(1, [[0, 1, { value: num(2.2250738585072014e-308) }]])]);
  await S.buildXlsx(base, [cells(1, [[0, 1, { value: num(0) }]])]);
});

test('R2: a cell whose start tag has ">" inside an attribute value is patched without damage', async () => {
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" foo="a>b"><v>1</v></c></row>') }] });
  const { bytes } = await S.buildXlsx(base, [cells(1, [[0, 0, { value: num(7) }]])]);
  assert.equal(await sheetData(bytes), '<row r="1"><c r="A1" foo="a>b"><v>7</v></c></row>');
});

// ===== fix round 3 =====

test('R3-6: new shared strings go before a trailing <extLst> of the sst', async () => {
  const sst = `${DECL}<sst xmlns="${NS_MAIN}" count="1" uniqueCount="1"><si><t>a</t></si><extLst><ext uri="x"/></extLst></sst>`;
  const base = await makeXlsx({ sheets: [{ name: 'S1', xml: sheetXml('<row r="1"><c r="A1" t="s"><v>0</v></c></row>') }], sst });
  const { bytes } = await save(base, [cells(1, [[0, 0, { value: str('new') }]])], [{ sheet: 1, r: 0, c: 0, value: str('new') }], { deep: false }); // the ExcelJS loader (V4) cannot read an sst with extLst at all
  assert.match(await part(bytes, 'xl/sharedStrings.xml'), /<si><t>a<\/t><\/si><si><t>new<\/t><\/si><extLst>/);
});

test('R3-10: a part that declares more than 512 MB unpacked is refused by parseZip', async () => {
  const zip = new JSZip();
  zip.file('a.xml', '<a/>');
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  const p = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  buf.writeUInt32LE(600 * 1024 * 1024, p + 24);
  assert.throws(() => S.parseZip(buf), (err) => err instanceof S.ZipError && /too large/.test(err.message));
});
