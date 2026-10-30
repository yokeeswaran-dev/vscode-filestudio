// Regression tests for fixed 0.1.x spreadsheet bugs (src/renderers/sheet.ts): xlsx number formats, colours, parts that
// ExcelJS reads wrong, lossy-feature detection and CSV parsing. Each test names the bug id from the 0.1.x bug hunts.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const S = require('../.out/sheet.js');

const sample = (name) => fs.readFileSync(path.join(__dirname, '..', 'samples', name));
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Column letters + row number -> { r, c } (0-based). */
function parseAddress(address) {
  const m = /^([A-Z]+)(\d+)$/.exec(address);
  const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { r: Number(m[2]) - 1, c };
}

/**
 * Builds a workbook with ExcelJS (`build` gets the Workbook), then replaces zip parts: `parts` maps a part name to
 * (old text or undefined) -> new text. Returns the .xlsx bytes.
 */
async function patchedXlsx(build, parts = {}) {
  const workbook = new ExcelJS.Workbook();
  await build(workbook);
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  for (const [name, edit] of Object.entries(parts)) {
    const old = zip.file(name) ? await zip.file(name).async('string') : undefined;
    zip.file(name, edit(old));
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
}

/** Loads fresh (no cache: edit-safe) and returns the model plus a cell lookup for sheet `sheet`. */
async function load(bytes, options, sheet = 0) {
  const workbook = await S.loadWorkbook(bytes, options);
  const model = S.workbookToModel(workbook);
  const { rows, styles } = model.getRows(sheet, 0, 200);
  const styleMap = new Map(model.getMeta().styles.map((style, i) => [i, style]));
  for (const [i, style] of styles) styleMap.set(i, style);
  const cell = (address) => {
    const { r, c } = parseAddress(address);
    return rows.find((row) => row.r === r)?.cells.find((x) => x.c === c);
  };
  const styleOf = (address) => {
    const found = cell(address);
    return found && found.s !== undefined ? styleMap.get(found.s) ?? {} : {};
  };
  return { workbook, model, cell, styleOf };
}

/** Display text (cell.w) of each [format, value] in a fresh one-column workbook (1900 or 1904 date system). */
async function displayed(cases, date1904 = false) {
  const bytes = await patchedXlsx((wb) => {
    if (date1904) wb.properties.date1904 = true;
    const ws = wb.addWorksheet('N');
    cases.forEach(([format, value], i) => {
      ws.getCell(i + 1, 1).value = value;
      if (format !== 'General') ws.getCell(i + 1, 1).numFmt = format;
    });
  });
  const { cell } = await load(bytes);
  return cases.map((_, i) => cell(`A${i + 1}`)?.w);
}

/** Asserts that every [format, value, excelText] case shows the text Excel 16 shows. */
async function assertExcelText(cases, date1904 = false) {
  const got = await displayed(cases, date1904);
  assert.deepEqual(
    cases.map(([format, value], i) => `${format} | ${value} -> ${got[i]}`),
    cases.map(([format, value, want]) => `${format} | ${value} -> ${want}`),
  );
}

// ===== NUMBER FORMATS (expected text = real Excel 16, from the H1 hunt oracle) =====

test('H1-F-NF1: time rounding to whole seconds carries into minutes / hours / days', async () => {
  await assertExcelText([
    ['h:mm:ss', 0.43124537037037036, '10:21:00'],
    ['h:mm', 0.999999, '0:00'],
    ['mm:ss', 0.00069, '01:00'],
    ['yyyy-mm-dd hh:mm:ss', 45000.99999999, '2023-03-16 00:00:00'],
    ['[h]:mm:ss', 100.99999999, '2424:00:00'],
  ]);
});

test('H1-F-NF2: no minus sign when every shown digit is 0 (no "-0")', async () => {
  await assertExcelText([
    ['0', -0.4, '0'],
    ['0.00', -0.004, '0.00'],
    ['0.0%', -0.0004, '0.0%'],
    ['$#,##0', -0.125, '$0'],
    ['0.00', -1e-15, '0.00'],
  ]);
});

test('H1-F-NF3: formats the ssf library rejects are formatted, not shown as General', async () => {
  await assertExcelText([
    ['#,##0.0#', 1234.5, '1,234.5'],
    ['#,##0.##', 1234.5, '1,234.5'],
    ['#,###.0', 1234.5678, '1,234.6'],
    ['0.??', 1234.5678, '1234.57'],
    ['??0.0?', 1234.5678, '1234.57'],
  ]);
});

test('H1-F-NF4: values are rounded to 15 significant digits before formatting', async () => {
  await assertExcelText([
    ['0.00 "kg"', 1.005, '1.01 kg'],
    ['0.00', 123456789012345.67, '123456789012346.00'],
    ['0%', 0.145, '15%'],
  ]);
});

test('H1-F-NF5: scientific formats match Excel (mantissa digits, exponent sign and width)', async () => {
  await assertExcelText([
    ['0E+0', 1234, '1E+3'],
    ['0.00E-00', 0.000123, '1.23E-04'],
    ['0.00E-00', 12345, '1.23E04'],
    ['##0.0E+0', 12345, '12.3E+3'],
    ['0.00E+00', 9.999, '1.00E+01'],
  ]);
});

test('H1-F-NF7: General rounds with carry (99999999999.5 -> 1E+11); "?" pads a trailing space', async () => {
  await assertExcelText([
    ['General', 99999999999.5, '1E+11'],
    ['General', 9999999999.95, '10000000000'],
    ['?.??', 1.5, '1.5 '],
  ]);
});

test('H1-F-NF8: "?" decimals are optional digits padded with spaces', async () => {
  await assertExcelText([
    ['0.0?', 1234.5678, '1234.57'],
    ['0.0?', -0.5, '-0.5 '],
  ]);
});

test('H1-F-NF9: negative halves round away from zero', async () => {
  await assertExcelText([
    ['0" x"', -2.5, '-3 x'],
    ['0.0" kg"', -1.25, '-1.3 kg'],
    ['0.00 "kg"', -0.125, '-0.13 kg'],
  ]);
});

test('H1-F-NF10: 1904 date system shows negative times and dates like Excel (not ########)', async () => {
  await assertExcelText(
    [
      ['[h]:mm', -0.5, '-12:00'],
      ['h:mm', -0.25, '-6:00'],
      ['m/d/yyyy', -1, '-1/2/1904'],
      ['yyyy-mm-dd', -400, '-1905-02-04'],
    ],
    true,
  );
});

test('H1-F-NF11: 1904 date system: past 31 Dec 9999 shows ######## (not 1/1/0000)', async () => {
  await assertExcelText(
    [
      ['m/d/yyyy', 2957003.5, '12/31/9999'],
      ['m/d/yyyy', 2957004, '########'],
    ],
    true,
  );
});

test('H1-F-NF6 / K1: accounting "* " formats get a fill layout ($ left, number right), not one glued string', async () => {
  const acct = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
  const bytes = await patchedXlsx((wb) => {
    const ws = wb.addWorksheet('F');
    const cases = [[1234.567, acct], [-1234.567, acct], ['abc', '@*-'], [1234.567, '#,##0.00_);(#,##0.00)']];
    cases.forEach(([value, format], i) => {
      ws.getCell(i + 1, 1).value = value;
      ws.getCell(i + 1, 1).numFmt = format;
    });
  });
  const { cell } = await load(bytes);
  const pick = (a) => ({ w: cell(a).w, fill: cell(a).fill });
  assert.deepEqual(pick('A1'), { w: ' $1,234.57 ', fill: { at: 2, pads: [[0, '('], [10, ')']] } });
  assert.deepEqual(pick('A2'), { w: ' $(1,234.57)', fill: { at: 2, pads: [[0, '(']] } });
  assert.deepEqual(pick('A3'), { w: 'abc', fill: { at: 3, char: '-' } });
  assert.deepEqual(pick('A4'), { w: '1,234.57 ', fill: undefined }, 'no "*" in the format: no fill layout');
});

// ===== COLOURS AND FILLS =====

test('H1-F-C2: indexed colours 66..81 are the Windows system colours Excel paints', () => {
  const theme = S.parseThemeXml(undefined);
  assert.equal(S.resolveColor({ indexed: 66 }, theme), '#646464');
  assert.equal(S.resolveColor({ indexed: 71 }, theme), '#0078D7');
  assert.equal(S.resolveColor({ indexed: 80 }, theme), '#FFFFE1');
  assert.equal(S.resolveColor({ indexed: 82 }, theme), undefined, 'past the table: no colour');
});

test('H1-F-C3: solid patternFill with no fgColor paints black; fgColor auto="1" paints nothing', async () => {
  const marks = ['FF100001', 'FF100002', 'FF100003'];
  const bytes = await patchedXlsx(
    (wb) => {
      const ws = wb.addWorksheet('F');
      marks.forEach((argb, i) => {
        ws.getCell(i + 1, 1).value = `fill ${i}`;
        ws.getCell(i + 1, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
      });
      ws.getCell('A4').value = 'fg 64';
      ws.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { indexed: 64 } };
    },
    {
      'xl/styles.xml': (xml) => {
        const out = xml
          .replace(`<patternFill patternType="solid"><fgColor rgb="${marks[0]}"/></patternFill>`, '<patternFill patternType="solid"/>')
          .replace(`<fgColor rgb="${marks[1]}"/>`, '<bgColor indexed="10"/>')
          .replace(`<fgColor rgb="${marks[2]}"/>`, '<fgColor auto="1"/>');
        assert.ok(out.includes('<patternFill patternType="solid"/>') && out.includes('<fgColor auto="1"/>'), 'fixture patched');
        return out;
      },
    },
  );
  const { styleOf } = await load(bytes);
  assert.deepEqual(
    ['A1', 'A2', 'A3', 'A4'].map((a) => styleOf(a).fill),
    ['#000000', '#000000', undefined, '#000000'],
    'no fgColor / bgColor only = black, fg auto = no fill, indexed 64 = black',
  );
});

// ===== PARTS EXCELJS READS WRONG (patched at load) =====

test('H2-F1: hyperlink with r:id + location keeps the #fragment; a dangling r:id does not fail the load', async () => {
  const bytes = await patchedXlsx(
    (wb) => {
      const ws = wb.addWorksheet('L');
      ws.getCell('A1').value = { text: 'page', hyperlink: 'https://example.com/page' };
      ws.getCell('A2').value = { text: 'file', hyperlink: 'other.xlsx' };
      ws.getCell('A3').value = { text: 'gone', hyperlink: 'https://example.com/x' };
      ws.getCell('A4').value = 'after';
    },
    {
      'xl/worksheets/sheet1.xml': (xml) => {
        const out = xml
          .replace('<hyperlink ref="A1" r:id="rId1"/>', '<hyperlink ref="A1" r:id="rId1" location="section-2"/>')
          .replace('<hyperlink ref="A2" r:id="rId2"/>', '<hyperlink ref="A2" r:id="rId2" location="Sheet1!A1"/>')
          .replace('<hyperlink ref="A3" r:id="rId3"/>', '<hyperlink ref="A3" r:id="rId99"/>');
        assert.ok(out.includes('rId99'), 'fixture patched');
        return out;
      },
    },
  );
  const { cell } = await load(bytes);
  assert.equal(cell('A1').link, 'https://example.com/page#section-2');
  assert.equal(cell('A2').link, 'other.xlsx#Sheet1!A1');
  assert.equal(cell('A3').link, undefined, 'dangling r:id: link dropped');
  assert.equal(cell('A4').w, 'after', 'the rest of the sheet loads');
});

/** Sheet with list validations written with whole-column / whole-row / openpyxl-style sqrefs. */
function validationXlsx() {
  return patchedXlsx(
    (wb) => {
      const ws = wb.addWorksheet('V');
      ws.getCell('A1').value = 'src';
      ws.getCell('A1').dataValidation = { type: 'list', formulae: ['"tmp"'] };
    },
    {
      'xl/worksheets/sheet1.xml': (xml) => {
        const dv =
          '<dataValidations count="3">' +
          '<dataValidation type="list" sqref="B:B 2:2"><formula1>"Yes,No"</formula1></dataValidation>' +
          '<dataValidation type="list" sqref="D1:D3 D5:D1048576"><formula1>$A$1</formula1></dataValidation>' +
          '<dataValidation type="list" sqref="A10:XFD200"><formula1>"p,q"</formula1></dataValidation></dataValidations>';
        const out = xml.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/, dv);
        assert.ok(out.includes('A10:XFD200'), 'fixture patched');
        return out;
      },
    },
  );
}

test('H2-F2 / K7: whole-column list validation loads fast, travels once in SheetMeta, and saves with its sqref', async () => {
  const bytes = await validationXlsx();
  const t0 = Date.now();
  const workbook = await S.loadWorkbook(bytes);
  const model = S.workbookToModel(workbook);
  const reply = model.getRows(0, 0, 256);
  const ms = Date.now() - t0;
  assert.ok(ms < 5000, `load + model + getRows took ${ms} ms (was > 5 s: millions of expanded validation keys)`);
  const meta = model.getMeta().sheets[0];
  assert.equal(meta.validations.length, 3, 'one entry per list validation');
  assert.ok(reply.rows.every((row) => row.cells.every((c) => c.list === undefined)), 'no per-cell list in getRows');
  assert.ok(JSON.stringify(reply).length < 5000, 'getRows reply stays small');
  const keys = Object.keys(workbook.worksheets[0].dataValidations.model);
  assert.ok(keys.length < 20 && keys.includes('B1:B1048576') && keys.includes('A10:XFD200'), `ranges stay one key each (${keys.length} keys)`);
  // ExcelJS' writer threw on whole-row validations before the fix; a save must keep the sqrefs.
  const saved = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const xml = await saved.file('xl/worksheets/sheet1.xml').async('string');
  const sqrefs = [...xml.matchAll(/<dataValidation\b[^>]*\bsqref="([^"]*)"/g)].map((m) => m[1]).sort();
  assert.deepEqual(sqrefs, ['A10:XFD200', 'A2:XFD2', 'B1:B1048576', 'D1:D3', 'D5:D1048576'].sort());
});

test('H2-F3: openpyxl-style sqref "B:B" / "2:2" is the whole column / row, not A1 / A2', async () => {
  const { model } = await load(await validationXlsx());
  const [first, second, third] = model.getMeta().sheets[0].validations;
  assert.deepEqual(first, { ranges: [{ r0: 0, c0: 1, r1: 1048575, c1: 1 }, { r0: 1, c0: 0, r1: 1, c1: 16383 }], values: ['Yes', 'No'] });
  assert.deepEqual(second, { ranges: [{ r0: 0, c0: 3, r1: 2, c1: 3 }, { r0: 4, c0: 3, r1: 1048575, c1: 3 }], values: ['src'] });
  assert.deepEqual(third, { ranges: [{ r0: 9, c0: 0, r1: 199, c1: 16383 }], values: ['p', 'q'] });
});

test('H2-F7: a sheet named "History" (refused by the ExcelJS API) loads with its name', async () => {
  const bytes = await patchedXlsx(
    (wb) => {
      for (const name of ['Data', 'Hist']) wb.addWorksheet(name).getCell('A1').value = name;
    },
    { 'xl/workbook.xml': (xml) => xml.replace('name="Hist"', 'name="History"') },
  );
  const { model, cell } = await load(bytes, undefined, 1);
  assert.deepEqual(model.getMeta().sheets.map((s) => s.name), ['Data', 'History']);
  assert.equal(cell('A1').w, 'Hist', 'rows of the "History" sheet are there');
});

test('U4a: a hidden active sheet stays the active sheet (Excel opens on it); a very hidden one gives way', async () => {
  const activeWith = async (activeTab) => {
    const bytes = await patchedXlsx((wb) => {
      wb.addWorksheet('Shown').getCell('A1').value = 1;
      wb.addWorksheet('Hid', { state: 'hidden' }).getCell('A1').value = 2;
      wb.addWorksheet('VeryHid', { state: 'veryHidden' }).getCell('A1').value = 3;
      wb.addWorksheet('Last').getCell('A1').value = 4;
      wb.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab, visibility: 'visible' }];
    });
    const { model } = await load(bytes);
    const meta = model.getMeta();
    assert.deepEqual(meta.sheets.map((s) => s.state), ['visible', 'hidden', 'veryHidden', 'visible']);
    return meta.activeSheet;
  };
  assert.equal(await activeWith(1), 1, 'hidden active sheet: kept');
  assert.equal(await activeWith(2), 0, 'very hidden active sheet: first visible sheet');
  assert.equal(await activeWith(3), 3, 'visible active sheet: kept');
});

test('H2-F8 / H4-F7: notes written as <text><t> are kept; a note on a cell absent from sheetData is shown', async () => {
  const bytes = await patchedXlsx(
    (wb) => {
      const ws = wb.addWorksheet('N');
      for (const [address, text] of [['A1', 'plain'], ['A2', 'runs'], ['A3', 'orphan']]) {
        ws.getCell(address).value = address;
        ws.getCell(address).note = text;
      }
    },
    {
      'xl/comments1.xml': (xml) => {
        const out = xml.replace('<text><r><t>plain</t></r></text>', '<text><t>plain t</t></text>').replace('<comment ref="A3"', '<comment ref="D9"');
        assert.ok(out.includes('<text><t>plain t</t></text>') && out.includes('ref="D9"'), 'fixture patched');
        return out;
      },
    },
  );
  const { model, cell } = await load(bytes);
  assert.equal(cell('A1').note, 'plain t', 'H2-F8: note with no runs');
  assert.equal(cell('A2').note, 'runs');
  assert.equal(cell('D9')?.note, 'orphan', 'H4-F7: note on a cell with no <c> element');
  const sheet = model.getMeta().sheets[0];
  assert.deepEqual([sheet.rowCount, sheet.colCount], [9, 4], 'H4-F7: the note cell is in the used range');
});

test('H2-F10: external workbook reference shows path + [file] like Excel (LoadOptions.folder), absPath wins', async () => {
  const extLink = {
    'xl/externalLinks/externalLink1.xml': () =>
      `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook xmlns:r="${REL}" r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>`,
    'xl/externalLinks/_rels/externalLink1.xml.rels': () =>
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="ext-source.xlsx" TargetMode="External"/></Relationships>`,
    'xl/_rels/workbook.xml.rels': (xml) => xml.replace('</Relationships>', `<Relationship Id="rIdE1" Type="${REL}/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`),
  };
  const build = (wb) => {
    wb.addWorksheet('X').getCell('A1').value = { formula: '[1]Sheet1!B2*10', result: 10 };
  };
  const refs = (xml) => xml.replace('</sheets>', '</sheets><externalReferences><externalReference r:id="rIdE1"/></externalReferences>');
  const noAbsPath = await patchedXlsx(build, { ...extLink, 'xl/workbook.xml': refs });
  const withAbsPath = await patchedXlsx(build, {
    ...extLink,
    'xl/workbook.xml': (xml) =>
      refs(xml).replace(
        '<sheets>',
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="D:\\Elsewhere\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><sheets>',
      ),
  });
  assert.equal((await load(noAbsPath, { folder: 'C:\\Data\\Books' })).cell('A1').f, "'C:\\Data\\Books\\[ext-source.xlsx]Sheet1'!B2*10");
  assert.equal((await load(withAbsPath, { folder: 'C:\\Data\\Books' })).cell('A1').f, "'D:\\Elsewhere\\[ext-source.xlsx]Sheet1'!B2*10");
  assert.equal((await load(noAbsPath)).cell('A1').f, "'[ext-source.xlsx]Sheet1'!B2*10", 'no folder given');
});

test('K6: a legacy array formula (<f t="array">) keeps its array flag', async () => {
  const bytes = await patchedXlsx(
    (wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 2;
      ws.getCell('C1').value = 2;
    },
    {
      'xl/worksheets/sheet1.xml': (xml) => {
        const out = xml.replace(/<c r="C1"([^>]*)><v>2<\/v><\/c>/, '<c r="C1"$1><f t="array" ref="C1">SUM(A1:B1*1)</f><v>2</v></c>');
        assert.ok(out.includes('t="array"'), 'fixture patched');
        return out;
      },
    },
  );
  const { cell } = await load(bytes);
  assert.equal(cell('C1').f, 'SUM(A1:B1*1)');
  assert.equal(cell('C1').array, true);
});

test('H2-F9: chartex / slicer fallback shapes and form-control twin shapes are not "Shapes / text boxes"', async () => {
  const base = sample('styled.xlsx');
  const zip = await JSZip.loadAsync(base);
  zip.file(
    'xl/drawings/drawing8.xml',
    '<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:mc="m"><xdr:twoCellAnchor><mc:AlternateContent><mc:Choice Requires="cx1"><xdr:graphicFrame macro=""/></mc:Choice>' +
      '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Chart 1"/></xdr:nvSpPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>' +
      '<mc:AlternateContent><mc:Choice Requires="a14"><xdr:twoCellAnchor><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1025" name="Button 1" hidden="1"><a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1025"/></a:ext></a:extLst></xdr:cNvPr></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></xdr:wsDr>',
  );
  zip.file('xl/ctrlProps/ctrlProp1.xml', '<formControlPr objectType="Button"/>');
  const before = await S.detectLossyFeatures(new Uint8Array(base));
  const features = await S.detectLossyFeatures(await zip.generateAsync({ type: 'uint8array' }));
  assert.ok(!features.includes('Shapes / text boxes'), `got ${JSON.stringify(features)}`);
  assert.ok(features.includes('Form controls / ActiveX'));
  // a real text box is still reported
  zip.file('xl/drawings/drawing7.xml', '<xdr:wsDr xmlns:xdr="x"><xdr:twoCellAnchor><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="TextBox 1"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>');
  const withTextBox = await S.detectLossyFeatures(await zip.generateAsync({ type: 'uint8array' }));
  assert.ok(withTextBox.includes('Shapes / text boxes'), `got ${JSON.stringify(withTextBox)} (styled.xlsx alone: ${JSON.stringify(before)})`);
});

test('K3: a damaged zip gets a readable message (library text kept in cause)', async () => {
  const truncated = new TextEncoder().encode('PK\u0003\u0004 truncated');
  const half = new Uint8Array(sample('styled.xlsx').subarray(0, sample('styled.xlsx').length >> 1));
  for (const [label, bytes] of [['truncated', truncated], ['half a workbook', half]]) {
    await assert.rejects(S.loadWorkbook(bytes), (err) => {
      assert.match(err.message, /^This file is not a valid \.xlsx workbook: its zip package is damaged/, label);
      assert.doesNotMatch(err.message, /central directory|Corrupted zip/i, `${label}: library text not in the message`);
      return true;
    });
  }
  await assert.rejects(S.detectLossyFeatures(truncated), /zip package is damaged/);
});

// ===== CSV =====

test('H2-F4: text after a closing quote stays in the same field (no row split); round trip is byte-identical', () => {
  const text = 'x,"12" monitor,y\na,b,c\n';
  const model = S.parseCsv(text);
  assert.equal(model.rows.length, 2);
  assert.deepEqual(model.rows[0], ['x', '12 monitor', 'y']);
  assert.equal(S.serializeCsv(model), text);
});

test('H2-F5: a second BOM is data (kept in A1); round trip is byte-identical', () => {
  const text = '\uFEFF\uFEFFa,b\n1,2\n';
  const model = S.parseCsv(text);
  assert.equal(model.rows[0][0], '\uFEFFa');
  assert.equal(S.serializeCsv(model), text);
});

test('H2-F6 / K5: CSV select-all past the 2,000,000-cell cap gives partial stats and says so', () => {
  const line = new Array(50).fill('1').join(',');
  const grid = new S.CsvGridModel(S.parseCsv(Array.from({ length: 60000 }, () => line).join('\n') + '\n'), 'big.csv');
  const capped = grid.getStats(0, [{ r0: 0, c0: 0, r1: 59999, c1: 49 }]);
  assert.deepEqual(
    { count: capped.count, sum: capped.sum, partial: capped.partial, scanned: capped.scanned },
    { count: 2000000, sum: 2000000, partial: true, scanned: 2000000 },
  );
  const exact = grid.getStats(0, [{ r0: 0, c0: 0, r1: 39999, c1: 49 }]);
  assert.equal(exact.count, 2000000);
  assert.equal(exact.partial, undefined, 'exactly 2,000,000 cells is complete');
});
