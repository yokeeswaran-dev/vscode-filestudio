#!/usr/bin/env node
/**
 * Sample-file generator for FileStudio.
 *
 * Builds the files in test/samples/ that exercise every renderer of the extension:
 *
 *   styled.xlsx           multi-sheet ExcelJS showcase (fonts, theme/indexed colors, fills, borders, alignment,
 *                         merges, rich text, links, notes, validation, number formats, formulas, conditional
 *                         formats, hidden sheets, images, frozen panes, zoom, RTL)
 *   large.xlsx            100,000 x 50 streaming workbook (skipped when present unless --force)
 *   sample.csv            CRLF, quoted commas/quotes/newlines, unicode, empty fields, numbers
 *   sample-semicolon.csv  ';' delimiter, LF, UTF-8 BOM, decimal commas
 *   sample.tsv            tab delimited, LF
 *   sample.psv            '|' delimited, LF, quoted fields containing pipes, quotes and a line break
 *   sample.ssv            ';' delimited (French style), decimal commas, spaces inside fields, UTF-8 BOM, CRLF
 *   sample-space.ssv      space separated, aligned columns (runs of spaces), "quoted fields" with spaces, LF
 *   sample.docx           hand-built OOXML (styles, numbering, table, hyperlink, image, footnote)
 *   sample.pdf            hand-written PDF 1.7: 4 pages (one landscape), selectable text, outline, links, image
 *   sample.pptx           hand-built 16:9 PresentationML: layouts, bullets, preset/custom shapes, group, pictures,
 *                         table with merged cells, bar + pie charts, speaker notes, hidden slide, gradient background
 *   images/sample.png     small RGBA PNG referenced by kitchen-sink.md
 *
 * kitchen-sink.md and other.md live in the same folder but are written by hand, never by this script.
 *
 * Usage: node test/generate-samples.js [--force] [--large-rows N] [--large-cols N] [--skip-large]
 *
 * The script is idempotent: every file except large.xlsx is rebuilt in memory and only written when its bytes
 * changed (all timestamps inside the zip containers are fixed), so re-running it leaves the folder untouched.
 * Only project dependencies (exceljs, jszip) and Node built-ins are used.
 */
'use strict';

// ===== IMPORTS & CONSTANTS =====

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const SAMPLES_DIR = path.join(__dirname, 'samples');
const IMAGES_DIR = path.join(SAMPLES_DIR, 'images');

/** Fixed timestamp for document properties and zip entries, so regenerated files are byte-identical. */
const FIXED_DATE = new Date(Date.UTC(2024, 0, 15, 9, 30, 0));
const GENERATOR_NAME = 'FileStudio sample generator';

const DEFAULT_LARGE_ROWS = 100000;
const DEFAULT_LARGE_COLS = 50;
const EXCEL_MAX_ROWS = 1048576;
const EXCEL_MAX_COLS = 16384;

/** Office 2013+ default theme ("Office") — ExcelJS ships the 2007 palette, real files use this one. */
const MODERN_THEME_CLR_SCHEME =
  '<a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme>';

/** Accent colors of the theme above (used for the generated images so they match the workbook). */
const ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

/** Placeholder formula replaced by a real `duplicateValues` rule after ExcelJS wrote the file (ExcelJS cannot). */
const DUPLICATE_VALUES_SENTINEL = '__FILE_VIEWER_DUPLICATE_VALUES__';

// ===== CLI =====

const USAGE = `Usage: node test/generate-samples.js [options]

Generates the sample files in test/samples/.

Options:
  --force           Regenerate large.xlsx even if it exists and rewrite files whose content is unchanged
  --skip-large      Do not generate large.xlsx
  --large-rows N    Data rows in large.xlsx, header not included (default ${DEFAULT_LARGE_ROWS})
  --large-cols N    Columns in large.xlsx (default ${DEFAULT_LARGE_COLS})
  -h, --help        Show this help`;

class UsageError extends Error {}

/**
 * @param {string[]} argv
 * @returns {{ force: boolean, skipLarge: boolean, largeRows: number, largeCols: number, help: boolean }}
 */
function parseArgs(argv) {
  const opts = {
    force: false,
    skipLarge: false,
    largeRows: DEFAULT_LARGE_ROWS,
    largeCols: DEFAULT_LARGE_COLS,
    help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    let name = argv[i];
    let inlineValue;
    const eq = name.indexOf('=');
    if (name.startsWith('--') && eq > 0) {
      inlineValue = name.slice(eq + 1);
      name = name.slice(0, eq);
    }
    const takeValue = () => {
      if (inlineValue !== undefined) return inlineValue;
      if (i + 1 >= argv.length) throw new UsageError(`${name} needs a value`);
      return argv[++i];
    };
    switch (name) {
      case '--force':
        opts.force = true;
        break;
      case '--skip-large':
        opts.skipLarge = true;
        break;
      case '--large-rows':
        opts.largeRows = parseCount(name, takeValue(), EXCEL_MAX_ROWS - 1);
        break;
      case '--large-cols':
        opts.largeCols = parseCount(name, takeValue(), EXCEL_MAX_COLS);
        break;
      case '-h':
      case '--help':
        opts.help = true;
        break;
      default:
        throw new UsageError(`Unknown option: ${argv[i]}`);
    }
  }
  return opts;
}

/** Parses a positive integer option value ("100000", "100_000" and "100,000" are accepted). */
function parseCount(name, raw, max) {
  const text = String(raw).replace(/[_,]/g, '');
  if (!/^\d+$/.test(text)) throw new UsageError(`${name} expects a positive integer, got "${raw}"`);
  const value = Number(text);
  if (value < 1 || value > max) throw new UsageError(`${name} must be between 1 and ${max}, got ${value}`);
  return value;
}

// ===== UTILITIES =====

/** Deterministic PRNG (mulberry32) so every run produces the same data. */
function createRandom(seed) {
  let state = seed >>> 0;
  const next = () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  return {
    next,
    /** Integer in [min, max]. */
    int: (min, max) => min + Math.floor(next() * (max - min + 1)),
    /** Number in [min, max) rounded to `decimals`. */
    float: (min, max, decimals) => Number((min + next() * (max - min)).toFixed(decimals)),
    pick: list => list[Math.floor(next() * list.length)],
    chance: probability => next() < probability,
  };
}

/** Writes `data` only when it differs from the file on disk. Returns the resulting status. */
function writeIfChanged(file, data, force) {
  const exists = fs.existsSync(file);
  if (exists && !force && fs.readFileSync(file).equals(data)) return 'unchanged';
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, data);
  return exists ? 'updated' : 'created';
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatMs(ms) {
  return ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(2)} s`;
}

function nowMs() {
  return Number(process.hrtime.bigint()) / 1e6;
}

function xmlEscape(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 1-based column number -> letters (1 -> A, 27 -> AA). */
function columnLetter(column) {
  let letters = '';
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

/** 1-based row/column -> 'B7'. */
function cellAddress(row, column) {
  return `${columnLetter(column)}${row}`;
}

/** JS Date (UTC wall clock) -> Excel 1900 serial number. */
function excelSerial(date) {
  return date.getTime() / 86400000 + 25569;
}

/** Rounds to 15 significant digits, like the value Excel stores for a computed result. */
function excelRound(value) {
  return Number(value.toPrecision(15));
}

/** Folds accents for ASCII e-mail addresses ("Tomás" -> "tomas"). */
function asciiFold(text) {
  return text.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9.]/g, '').toLowerCase();
}

/**
 * Re-packs a zip with fixed entry timestamps (and an optional per-entry transform) so the output is
 * byte-identical between runs. `[Content_Types].xml` and `_rels/.rels` are stored first, like Office does.
 * @param {Buffer} buffer
 * @param {(name: string, content: Buffer) => Buffer | string | undefined} [transform]
 */
async function repackZip(buffer, transform) {
  const source = await JSZip.loadAsync(buffer);
  const names = [];
  source.forEach((name, entry) => {
    if (!entry.dir) names.push(name);
  });
  const priority = name => (name === '[Content_Types].xml' ? 0 : name === '_rels/.rels' ? 1 : 2);
  names.sort((a, b) => priority(a) - priority(b));

  const target = new JSZip();
  for (const name of names) {
    let content = await source.file(name).async('nodebuffer');
    if (transform) {
      const replaced = transform(name, content);
      if (replaced !== undefined) content = Buffer.isBuffer(replaced) ? replaced : Buffer.from(replaced, 'utf8');
    }
    target.file(name, content, { date: FIXED_DATE, createFolders: false });
  }
  return target.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    platform: 'DOS',
  });
}

// ===== PNG ENCODER =====

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let crc = 0xffffffff;
  for (let i = 0; i < buffer.length; i++) crc = CRC_TABLE[(crc ^ buffer[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const typeAndData = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typeAndData), 0);
  return Buffer.concat([length, typeAndData, crc]);
}

/**
 * Encodes an 8-bit truecolor PNG (RGB, or RGBA when `alpha` is set) from a pixel callback.
 * @param {number} width
 * @param {number} height
 * @param {(x: number, y: number) => number[]} pixel returns [r, g, b] or [r, g, b, a] (0-255)
 * @param {{ alpha?: boolean }} [options]
 */
function encodePng(width, height, pixel, options = {}) {
  const channels = options.alpha ? 4 : 3;
  const stride = width * channels + 1;
  const raw = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    let offset = y * stride;
    raw[offset++] = 0; // filter type: None
    for (let x = 0; x < width; x++) {
      const rgba = pixel(x, y);
      for (let c = 0; c < channels; c++) {
        const value = c < 3 ? rgba[c] : rgba[3] === undefined ? 255 : rgba[3];
        raw[offset++] = Math.max(0, Math.min(255, Math.round(value)));
      }
    }
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8; // bit depth
  header[9] = options.alpha ? 6 : 2; // color type: RGBA / RGB
  header[10] = 0; // compression
  header[11] = 0; // filter method
  header[12] = 0; // no interlace
  const physical = Buffer.alloc(9); // 96 DPI = 3780 pixels per metre
  physical.writeUInt32BE(3780, 0);
  physical.writeUInt32BE(3780, 4);
  physical[8] = 1;
  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', header),
    pngChunk('pHYs', physical),
    pngChunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** h in degrees, s and l in [0, 1] -> [r, g, b]. */
function hslToRgb(h, s, l) {
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const hp = (((h % 360) + 360) % 360) / 60;
  const x = chroma * (1 - Math.abs((hp % 2) - 1));
  const [r, g, b] =
    hp < 1 ? [chroma, x, 0] : hp < 2 ? [x, chroma, 0] : hp < 3 ? [0, chroma, x]
      : hp < 4 ? [0, x, chroma] : hp < 5 ? [x, 0, chroma] : [chroma, 0, x];
  const m = l - chroma / 2;
  return [(r + m) * 255, (g + m) * 255, (b + m) * 255];
}

function hexToRgb(hex) {
  const value = parseInt(hex, 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function mix(a, b, t) {
  return a.map((v, i) => v + (b[i] - v) * t);
}

/** Pixel callback of the hue/lightness gradient: dark frame and a white top-left marker (shows orientation). */
function gradientPixel(width, height) {
  return (x, y) => {
    if (x === 0 || y === 0 || x === width - 1 || y === height - 1) return [51, 51, 51];
    if (x >= 4 && x < 12 && y >= 4 && y < 12) return [255, 255, 255];
    return hslToRgb((330 * x) / (width - 1), 0.75, 0.78 - (0.45 * y) / (height - 1));
  };
}

/** 120x80 hue/lightness gradient with a dark frame and a white top-left marker (shows orientation). */
function makeGradientPng(width = 120, height = 80) {
  return encodePng(width, height, gradientPixel(width, height));
}

/** 240x144 bar chart picture in the theme accent colors (embedded in the docx and the Images sheet). */
function makeChartPng(width = 240, height = 144) {
  const bars = [62, 95, 41, 112, 78, 88];
  const baseline = height - 14;
  const barWidth = 26;
  const gap = (width - 20 - bars.length * barWidth) / (bars.length - 1);
  const top = [240, 246, 255];
  const bottom = [255, 255, 255];
  return encodePng(width, height, (x, y) => {
    if (x === 0 || y === 0 || x === width - 1 || y === height - 1) return [140, 140, 140];
    for (let i = 0; i < bars.length; i++) {
      const left = 10 + i * (barWidth + gap);
      if (x >= left && x < left + barWidth && y < baseline && y >= baseline - bars[i]) {
        const shade = 1 - 0.18 * ((x - left) / barWidth);
        return hexToRgb(ACCENTS[i]).map(v => v * shade);
      }
    }
    if (y === baseline) return [89, 89, 89];
    if (y < baseline && (baseline - y) % 24 === 0) return [217, 217, 217];
    return mix(top, bottom, y / height);
  });
}

/** Distance from point (px, py) to the segment (ax, ay)-(bx, by). */
function segmentDistance(px, py, ax, ay, bx, by) {
  const dx = bx - ax;
  const dy = by - ay;
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

/** 64x64 RGBA badge: anti-aliased blue disc with a white check mark on a transparent background. */
function makeBadgePng(size = 64) {
  const center = (size - 1) / 2;
  const radius = size / 2 - 2;
  const light = hexToRgb('8FAADC');
  const dark = hexToRgb('2F5597');
  const s = size / 64;
  return encodePng(
    size,
    size,
    (x, y) => {
      const distance = Math.hypot(x - center, y - center);
      const coverage = Math.max(0, Math.min(1, radius - distance + 0.5));
      if (coverage === 0) return [0, 0, 0, 0];
      const check = Math.min(
        segmentDistance(x, y, 18 * s, 33 * s, 28 * s, 43 * s),
        segmentDistance(x, y, 28 * s, 43 * s, 46 * s, 22 * s),
      );
      const ink = Math.max(0, Math.min(1, 3.2 * s - check + 0.5));
      const base = mix(light, dark, Math.min(1, Math.hypot(x - center * 0.6, y - center * 0.6) / (size * 0.9)));
      const rgb = mix(base, [255, 255, 255], ink);
      return [rgb[0], rgb[1], rgb[2], coverage * 255];
    },
    { alpha: true },
  );
}

// ===== STYLED WORKBOOK: STYLE HELPERS =====

const ARGB = hex => ({ argb: `FF${hex.toUpperCase()}` });
const solidFill = color => ({ type: 'pattern', pattern: 'solid', fgColor: color, bgColor: { indexed: 64 } });
const edge = (style, color = ARGB('595959')) => ({ style, color });
const box = (style, color) => ({ top: edge(style, color), left: edge(style, color), bottom: edge(style, color), right: edge(style, color) });
/** Full font definition (what Excel itself writes): Calibri 11 plus overrides. */
const font = (overrides = {}) => ({ name: 'Calibri', family: 2, size: 11, ...overrides });

const CENTER = { horizontal: 'center', vertical: 'middle' };
const LABEL_FONT = font({ bold: true, color: { theme: 1, tint: 0.25 } });
const SECTION_FONT = font({ bold: true, size: 12, color: { theme: 4, tint: -0.25 } });
const SECTION_FILL = solidFill({ theme: 4, tint: 0.8 });
const LINK_FONT = font({ underline: true, color: { theme: 10 } });
const NOTE_FONT = font({ italic: true, size: 9, color: ARGB('7F7F7F') });
const HEADER_STYLE = {
  font: font({ bold: true, color: { theme: 1 } }),
  fill: solidFill({ theme: 0, tint: -0.15 }),
  border: { bottom: edge('medium', { theme: 4 }) },
  alignment: { vertical: 'middle' },
};

/** Applies a partial style ({ font, fill, border, alignment, numFmt }) to a cell. */
function applyStyle(cell, style = {}) {
  if (style.font) cell.font = style.font;
  if (style.fill) cell.fill = style.fill;
  if (style.border) cell.border = style.border;
  if (style.alignment) cell.alignment = style.alignment;
  if (style.numFmt) cell.numFmt = style.numFmt;
}

/**
 * Sets value, style and extras ({ note, dataValidation }) of one cell. `value === undefined` leaves the value
 * alone, which creates an empty cell that only carries a style.
 */
function put(ws, address, value, style, extra = {}) {
  const cell = ws.getCell(address);
  if (value !== undefined) cell.value = value;
  applyStyle(cell, style);
  if (extra.note) cell.note = extra.note;
  if (extra.dataValidation) cell.dataValidation = extra.dataValidation;
  return cell;
}

// ===== STYLED WORKBOOK: FORMATTING SHEET =====

/**
 * Sheet 1 — every cell-formatting feature: fonts, colors (argb / theme+tint / indexed), fills, borders, alignment,
 * merges, sizes, hidden rows/columns, rich text, hyperlinks, notes, validation, row/column styles. Frozen 2 rows x 1 col.
 */
function addFormattingSheet(wb) {
  const ws = wb.addWorksheet('Formatting', {
    properties: { tabColor: ARGB('4472C4') },
    views: [{ state: 'frozen', xSplit: 1, ySplit: 2, topLeftCell: 'B3', activeCell: 'B3' }],
  });
  const LAST_COLUMN = 13; // M
  const widths = { A: 30, B: 18, C: 16, D: 16, E: 16, F: 16, G: 16, H: 18, I: 12, J: 20, K: 14, L: 16, M: 16 };
  for (const [letter, width] of Object.entries(widths)) ws.getColumn(letter).width = width;

  // Column default style (written as <col style=…>): set before any cell exists so new cells inherit it.
  const styledColumn = ws.getColumn('M');
  styledColumn.fill = solidFill(ARGB('FFF2CC'));
  styledColumn.font = font({ color: ARGB('7F6000') });
  ws.getColumn('K').hidden = true;

  // --- Title (merged header row) and frozen column-header row ---
  put(ws, 'A1', 'FileStudio — styled.xlsx formatting showcase', {
    font: font({ name: 'Calibri Light', family: 2, size: 18, bold: true, color: { theme: 0 } }),
    fill: solidFill({ theme: 4 }),
    alignment: CENTER,
  });
  ws.mergeCells(`A1:${columnLetter(LAST_COLUMN)}1`);
  ws.getRow(1).height = 36;

  put(ws, 'A2', 'Feature', HEADER_STYLE);
  put(ws, 'B2', 'Examples (frozen header, merged B2:J2)', { ...HEADER_STYLE, alignment: CENTER });
  ws.mergeCells('B2:J2');
  put(ws, 'K2', 'Hidden column K', HEADER_STYLE);
  put(ws, 'L2', 'List source', HEADER_STYLE);
  put(ws, 'M2', 'Column style', HEADER_STYLE);
  ws.getRow(2).height = 20;
  ['Apple', 'Banana', 'Cherry', 'Date', 'Elderberry', 'Fig'].forEach((fruit, i) => put(ws, `L${3 + i}`, fruit));

  // --- Row cursor helpers ---
  let r = 3;
  const section = title => {
    r += 1; // spacer row before every section
    const row = ws.getRow(r);
    row.height = 21;
    for (let c = 1; c <= LAST_COLUMN; c++) row.getCell(c).fill = SECTION_FILL;
    put(ws, `A${r}`, title, { font: SECTION_FONT, alignment: { vertical: 'middle' } });
    r += 1;
  };
  const label = text => put(ws, `A${r}`, text, { font: LABEL_FONT, alignment: { vertical: 'middle' } });
  /**
   * One feature row: label in A, then cells from column B. Each spec is [value, style?, extra?] or null (skip).
   */
  const line = (text, specs = [], { height, startColumn = 2 } = {}) => {
    if (text) label(text);
    specs.forEach((spec, i) => {
      if (spec) put(ws, cellAddress(r, startColumn + i), spec[0], spec[1], spec[2]);
    });
    if (height) ws.getRow(r).height = height;
    return r++;
  };

  // --- Fonts ---
  section('Fonts');
  line(
    'Font families',
    ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana', 'Consolas'].map(name => [
      name,
      { font: font({ name, family: undefined }) },
    ]),
  );
  line(
    'Font sizes (pt)',
    [8, 10, 11, 14, 18, 24, 36].map(size => [`${size} pt`, { font: font({ size }), alignment: { vertical: 'bottom' } }]),
    { height: 48 },
  );
  line('Bold / italic / underline / strike', [
    ['Bold', { font: font({ bold: true }) }],
    ['Italic', { font: font({ italic: true }) }],
    ['Underline', { font: font({ underline: true }) }],
    ['Double underline', { font: font({ underline: 'double' }) }],
    ['Strikethrough', { font: font({ strike: true }) }],
    ['All four', { font: font({ bold: true, italic: true, underline: true, strike: true }) }],
    ['Accounting u/l', { font: font({ underline: 'singleAccounting' }) }],
  ]);
  line('Superscript / subscript', [
    [{ richText: [{ text: 'E = mc', font: font() }, { text: '2', font: font({ vertAlign: 'superscript' }) }] }],
    [{ richText: [{ text: 'H', font: font() }, { text: '2', font: font({ vertAlign: 'subscript' }) }, { text: 'O', font: font() }] }],
    ['Superscript cell', { font: font({ vertAlign: 'superscript' }) }],
    ['Subscript cell', { font: font({ vertAlign: 'subscript' }) }],
  ]);
  line(
    'Font colors (ARGB)',
    [['Red', 'FF0000'], ['Green', '00B050'], ['Blue', '0070C0'], ['Orange', 'FFC000'], ['Purple', '7030A0'], ['Gray', '808080']].map(
      ([name, hex]) => [name, { font: font({ bold: true, color: ARGB(hex) }) }],
    ),
  );
  line(
    'Font colors (theme + tint)',
    [
      ['theme 1 (Text 1)', { theme: 1 }],
      ['theme 4 (Accent 1)', { theme: 4 }],
      ['theme 4, tint 0.4', { theme: 4, tint: 0.4 }],
      ['theme 5, tint -0.25', { theme: 5, tint: -0.25 }],
      ['theme 9 (Accent 6)', { theme: 9 }],
      ['theme 3 (Text 2)', { theme: 3 }],
      ['theme 10 (Link)', { theme: 10 }],
    ].map(([text, color]) => [text, { font: font({ bold: true, color }) }]),
  );
  line(
    'Font colors (indexed palette)',
    [10, 12, 17, 20, 53, 64].map(index => [
      index === 64 ? 'indexed 64 (system)' : `indexed ${index}`,
      { font: font({ bold: true, color: { indexed: index } }) },
    ]),
  );

  // --- Fills ---
  section('Fills');
  line(
    'Solid fills (ARGB)',
    [['FFC7CE', 'Light red'], ['C6EFCE', 'Light green'], ['FFEB9C', 'Light yellow'], ['BDD7EE', 'Light blue'], ['000000', 'Black']].map(
      ([hex, text]) => [text, { fill: solidFill(ARGB(hex)), font: font({ color: hex === '000000' ? ARGB('FFFFFF') : undefined }) }],
    ),
  );
  line(
    'Solid fills (theme + tint)',
    [
      ['theme 4, tint 0.8', { theme: 4, tint: 0.8 }, false],
      ['theme 4, tint 0.4', { theme: 4, tint: 0.4 }, false],
      ['theme 4', { theme: 4 }, true],
      ['theme 4, tint -0.5', { theme: 4, tint: -0.5 }, true],
      ['theme 5, tint 0.6', { theme: 5, tint: 0.6 }, false],
      ['theme 9, tint 0.4', { theme: 9, tint: 0.4 }, false],
      ['theme 0, tint -0.15', { theme: 0, tint: -0.15 }, false],
    ].map(([text, color, dark]) => [text, { fill: solidFill(color), font: dark ? font({ color: { theme: 0 } }) : undefined }]),
  );
  const PATTERNS = [
    'darkGrid', 'lightTrellis', 'gray125', 'gray0625', 'darkVertical', 'lightHorizontal', 'darkDown',
    'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkUp', 'darkTrellis', 'lightVertical',
    'lightDown', 'lightUp', 'lightGrid',
  ];
  for (let i = 0; i < PATTERNS.length; i += 7) {
    const themed = i >= 14; // last row uses theme colors for the pattern fg/bg
    line(
      `Pattern fills (${i / 7 + 1}/3)${themed ? ' — theme colors' : ''}`,
      PATTERNS.slice(i, i + 7).map(pattern => [
        pattern,
        {
          fill: themed
            ? { type: 'pattern', pattern, fgColor: { theme: 5 }, bgColor: { theme: 0 } }
            : { type: 'pattern', pattern, fgColor: ARGB('2F5597'), bgColor: ARGB('FFF2CC') },
          font: font({ bold: true, size: 9 }),
          alignment: CENTER,
        },
      ]),
      { height: 24 },
    );
  }
  line(
    'Gradient fills',
    [
      ['Linear 0°', { type: 'gradient', gradient: 'angle', degree: 0, stops: [{ position: 0, color: ARGB('FFFFFF') }, { position: 1, color: ARGB('4472C4') }] }],
      ['Linear 90°, 3 stops', { type: 'gradient', gradient: 'angle', degree: 90, stops: [{ position: 0, color: ARGB('C00000') }, { position: 0.5, color: ARGB('FFFFFF') }, { position: 1, color: ARGB('00B050') }] }],
      ['Linear 45°', { type: 'gradient', gradient: 'angle', degree: 45, stops: [{ position: 0, color: ARGB('FFC000') }, { position: 1, color: ARGB('7030A0') }] }],
      ['Path (center)', { type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops: [{ position: 0, color: ARGB('FFFFFF') }, { position: 1, color: { theme: 5 } }] }],
    ].map(([text, fill]) => [text, { fill, alignment: CENTER, font: font({ bold: true }) }]),
    { height: 32 },
  );

  // --- Borders ---
  section('Borders');
  const BORDER_STYLES = [
    'thin', 'medium', 'thick', 'double', 'hair', 'dotted', 'dashed', 'mediumDashed',
    'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
  ];
  for (let i = 0; i < BORDER_STYLES.length; i += 4) {
    if (i === 0) label('Border line styles (all sides)');
    BORDER_STYLES.slice(i, i + 4).forEach((style, j) =>
      put(ws, cellAddress(r, 2 + j * 2), style, { border: box(style, ARGB('1F3864')), alignment: CENTER }),
    );
    ws.getRow(r).height = 24;
    r++;
    if (i + 4 < BORDER_STYLES.length) {
      ws.getRow(r).height = 8; // spacer so each sample stands alone
      r++;
    }
  }
  label('Adjacent cells (thin grid, medium outline)');
  for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
      put(ws, cellAddress(r + i, 2 + j), i * 3 + j + 1, {
        alignment: CENTER,
        border: {
          top: edge(i === 0 ? 'medium' : 'thin'),
          bottom: edge(i === 2 ? 'medium' : 'thin'),
          left: edge(j === 0 ? 'medium' : 'thin'),
          right: edge(j === 2 ? 'medium' : 'thin'),
        },
      });
    }
  }
  put(ws, `F${r}`, 'Bottom borders only', { border: { bottom: edge('thin') } });
  put(ws, `G${r}`, 'share an edge', { border: { bottom: edge('thin') } });
  put(ws, `F${r + 1}`, 'Top thick / bottom double', { border: { top: edge('thick'), bottom: edge('double') } });
  r += 3;
  line(
    'Colored borders',
    [
      ['Mixed sides', { alignment: CENTER, border: { top: edge('thick', ARGB('FF0000')), right: edge('medium', ARGB('00B050')), bottom: edge('double', ARGB('0070C0')), left: edge('dashed', ARGB('FFC000')) } }],
      null,
      ['Theme color', { alignment: CENTER, border: box('thick', { theme: 5 }) }],
      null,
      ['Indexed color', { alignment: CENTER, border: box('medium', { indexed: 10 }) }],
    ],
    { height: 30 },
  );
  line(
    'Diagonal borders',
    [
      ['Diagonal up', { alignment: CENTER, border: { diagonal: { up: true, style: 'thin', color: ARGB('C00000') } } }],
      null,
      ['Diagonal down', { alignment: CENTER, border: { diagonal: { down: true, style: 'medium', color: ARGB('0070C0') } } }],
      null,
      ['Both + box', { alignment: CENTER, border: { ...box('thin'), diagonal: { up: true, down: true, style: 'dashed', color: ARGB('7F7F7F') } } }],
    ],
    { height: 36 },
  );

  // --- Alignment ---
  section('Alignment');
  ['top', 'middle', 'bottom'].forEach((vertical, i) => {
    if (i === 0) label('Horizontal × vertical');
    ['left', 'center', 'right'].forEach((horizontal, j) =>
      put(ws, cellAddress(r, 2 + j), `${horizontal} / ${vertical}`, {
        alignment: { horizontal, vertical },
        border: box('hair', ARGB('A6A6A6')),
      }),
    );
    ws.getRow(r).height = 36;
    r++;
  });
  line(
    'Other horizontal modes',
    [
      ['ab', { alignment: { horizontal: 'fill' } }],
      ['Justified text that wraps over several lines.', { alignment: { horizontal: 'justify', vertical: 'top' } }],
      ['Distributed words here', { alignment: { horizontal: 'distributed', vertical: 'middle' } }],
      ['Centered across D:E', { alignment: { horizontal: 'centerContinuous', vertical: 'middle' } }],
      [undefined, { alignment: { horizontal: 'centerContinuous', vertical: 'middle' } }],
    ],
    { height: 48 },
  );
  line(
    'Vertical justify / distributed',
    [
      ['Vertically justified text in a tall cell', { alignment: { vertical: 'justify', wrapText: true } }],
      ['Vertically distributed text lines', { alignment: { vertical: 'distributed', wrapText: true } }],
    ],
    { height: 64 },
  );
  line(
    'Wrap text',
    [
      ['This long sentence wraps inside the cell because wrap text is on.', { alignment: { wrapText: true, vertical: 'top' } }],
      ['Line one\nLine two\nLine three', { alignment: { wrapText: true, vertical: 'top' } }],
      ['Not wrapped: clipped by the next cell', {}],
      ['(next cell)', {}],
    ],
    { height: 60 },
  );
  line('Text overflow', [
    ['Long left-aligned text spills into the empty cells on its right', {}],
    null,
    null,
    ['Stops here ->', { font: font({ italic: true }) }],
    null,
    null,
    ['Right-aligned text spills into the empty cells on its left', { alignment: { horizontal: 'right' } }],
  ]);
  line('Indent', [
    ['Indent 1', { alignment: { horizontal: 'left', indent: 1 } }],
    ['Indent 2', { alignment: { horizontal: 'left', indent: 2 } }],
    ['Indent 3', { alignment: { horizontal: 'left', indent: 3 } }],
    ['Right, indent 2', { alignment: { horizontal: 'right', indent: 2 } }],
  ]);
  line(
    'Text rotation',
    [
      ['45°', { alignment: { ...CENTER, textRotation: 45 } }],
      ['-45°', { alignment: { ...CENTER, textRotation: -45 } }],
      ['90°', { alignment: { ...CENTER, textRotation: 90 } }],
      ['-90°', { alignment: { ...CENTER, textRotation: -90 } }],
      ['Vertical', { alignment: { ...CENTER, textRotation: 'vertical' } }],
    ],
    { height: 72 },
  );
  line('Shrink to fit', [
    ['This text shrinks to fit the cell width', { alignment: { shrinkToFit: true } }],
    [1234567.891, { alignment: { shrinkToFit: true }, numFmt: '#,##0.000' }],
  ]);

  // --- Merged cells ---
  section('Merged cells');
  const mergeTop = r;
  label('Horizontal merge (B:D)');
  r++;
  label('Block merge (B:D × 2 rows)');
  r++;
  label('Vertical (F) and block (H:J) merges');
  r++;
  put(ws, `B${mergeTop}`, 'Merged B:D', { alignment: CENTER, fill: solidFill({ theme: 4, tint: 0.8 }), border: box('thin') });
  ws.mergeCells(`B${mergeTop}:D${mergeTop}`);
  put(ws, `B${mergeTop + 1}`, 'Block B:D × 2 rows', { alignment: CENTER, fill: solidFill({ theme: 9, tint: 0.8 }), border: box('thin') });
  ws.mergeCells(`B${mergeTop + 1}:D${mergeTop + 2}`);
  put(ws, `F${mergeTop}`, 'Vertical F × 3 rows', { alignment: { ...CENTER, wrapText: true }, fill: solidFill({ theme: 5, tint: 0.8 }), border: box('thin') });
  ws.mergeCells(`F${mergeTop}:F${mergeTop + 2}`);
  put(ws, `H${mergeTop}`, 'Block H:J × 3 rows with a medium outline', {
    alignment: { ...CENTER, wrapText: true },
    fill: solidFill({ theme: 7, tint: 0.6 }),
    border: box('medium', { theme: 1 }),
  });
  ws.mergeCells(`H${mergeTop}:J${mergeTop + 2}`);

  // --- Rows & columns ---
  section('Row heights, column widths, hidden rows/columns');
  line('Custom row height 32 pt', [['Tall row', { alignment: { vertical: 'middle' } }]], { height: 32 });
  line('Custom row height 10 pt', [['Short row', { font: font({ size: 7 }) }]], { height: 10 });
  line('Next row is hidden ↓');
  const hiddenRow = line('Hidden row', [['This row is hidden', {}]]);
  ws.getRow(hiddenRow).hidden = true;
  line('Previous row is hidden ↑');
  const widthRow = line(
    'Column widths (characters)',
    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].map(letter => [`${letter} = ${widths[letter]}`, { font: NOTE_FONT, alignment: { horizontal: 'center' } }]),
  );
  put(ws, `K${widthRow}`, 'K is hidden', { font: NOTE_FONT });

  // --- Rich text, hyperlinks, notes, validation ---
  section('Rich text, hyperlinks, notes, data validation');
  line('Rich text (multiple runs)', [
    [
      {
        richText: [
          { text: 'Rich ', font: font({ bold: true }) },
          { text: 'text ', font: font({ italic: true, color: ARGB('C00000') }) },
          { text: 'with ', font: font({ underline: true }) },
          { text: 'many ', font: font({ size: 14, color: { theme: 4 } }) },
          { text: 'runs', font: { strike: true, name: 'Consolas' } },
        ],
      },
    ],
  ]);
  line('Hyperlinks', [
    [{ text: 'VS Code website', hyperlink: 'https://code.visualstudio.com/', tooltip: 'https://code.visualstudio.com/' }, { font: LINK_FONT }],
    [{ text: 'E-mail link', hyperlink: 'mailto:someone@example.com' }, { font: LINK_FONT }],
    [{ text: 'Go to Numbers!A1', hyperlink: 'Numbers!A1' }, { font: LINK_FONT }],
    [{ text: 'Go to CondFormat!B2', hyperlink: '#CondFormat!B2' }, { font: LINK_FONT }],
    [{ formula: 'HYPERLINK("https://example.com/","HYPERLINK()")', result: 'HYPERLINK()' }, { font: LINK_FONT }],
  ]);
  line('Link kinds ↑', [
    ['external https', { font: NOTE_FONT }],
    ['mailto', { font: NOTE_FONT }],
    ['internal (location)', { font: NOTE_FONT }],
    ['internal (#target)', { font: NOTE_FONT }],
    ['HYPERLINK() formula', { font: NOTE_FONT }],
  ]);
  line('Cell notes (comments)', [
    ['Plain note', {}, { note: 'A plain-text note.\nIt has two lines.' }],
    [
      'Rich note',
      {},
      {
        note: {
          texts: [
            { font: { bold: true, size: 9, name: 'Tahoma', family: 2 }, text: 'Reviewer:\n' },
            { font: { size: 9, name: 'Tahoma', family: 2 }, text: 'Rich note with a bold author line.' },
          ],
        },
      },
    ],
    [42, { alignment: { horizontal: 'center' } }, { note: 'Notes can sit on numbers too.' }],
  ]);
  line('Data validation', [
    ['Green', {}, { dataValidation: { type: 'list', allowBlank: true, formulae: ['"Red,Green,Blue"'], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid color', error: 'Choose Red, Green or Blue.' } }],
    ['Banana', {}, { dataValidation: { type: 'list', allowBlank: true, formulae: ['$L$3:$L$8'], showErrorMessage: true } }],
    ['Mango', {}, { dataValidation: { type: 'list', allowBlank: true, formulae: ['FruitList'], showErrorMessage: true } }],
    ['Lime', {}, { dataValidation: { type: 'list', allowBlank: true, formulae: ['Hidden!$A$2:$A$7'], showErrorMessage: true } }],
    [50, { alignment: { horizontal: 'center' } }, { dataValidation: { type: 'whole', operator: 'between', allowBlank: true, formulae: [1, 100], showInputMessage: true, promptTitle: 'Whole number', prompt: 'Enter a whole number between 1 and 100.', showErrorMessage: true } }],
  ]);
  line('Validation kinds ↑', [
    ['list: literal', { font: NOTE_FONT }],
    ['list: range L3:L8', { font: NOTE_FONT }],
    ['list: name FruitList', { font: NOTE_FONT }],
    ['list: Hidden!A2:A7', { font: NOTE_FONT }],
    ['whole 1..100', { font: NOTE_FONT }],
  ]);

  // --- Row / column default styles ---
  section('Row and column default styles');
  const styledRow = ws.getRow(r);
  styledRow.fill = solidFill({ theme: 7, tint: 0.8 }); // written as <row s=… customFormat="1">
  styledRow.font = font({ italic: true, color: { theme: 7, tint: -0.5 } });
  line('Row style (whole row)', [['Every cell of this row, even empty ones, uses the row style']]);
  line('Column style (column M)', [['Column M has a default fill — see the yellow column ->', { font: NOTE_FONT }]]);

  return ws;
}

// ===== STYLED WORKBOOK: NUMBERS SHEET =====

const ACCOUNTING_FORMAT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const SAMPLE_DATE = new Date(Date.UTC(2024, 2, 15)); // Friday 15 March 2024
const SAMPLE_DATE_TIME = new Date(Date.UTC(2024, 2, 15, 14, 30));

/**
 * Number-format samples: [description, format code, value, what Excel displays].
 * The last column was captured from Excel 365 (en-US conventions) so the viewer's output can be checked by eye.
 * Locale-dependent parts: the "/" date separator and the digit grouping of "#,##0" follow the OS settings in
 * Excel (en-IN shows 12,34,568 and 3-15-2024).
 */
const NUMBER_FORMATS = [
  ['General', 'General', 1234.5678, '1234.5678'],
  ['General (large number)', 'General', 123456789012, '1.23457E+11'],
  ['General (small number)', 'General', 0.000012345, '0.000012345'],
  ['Fixed, 2 decimals', '0.00', 3.14159, '3.14'],
  ['Thousands separator', '#,##0', 1234567.891, '1,234,568'],
  ['Thousands, 2 decimals (negative)', '#,##0.00', -9876.5, '-9,876.50'],
  ['Currency', '$#,##0.00;[Red]-$#,##0.00', 1234.5, '$1,234.50'],
  ['Currency (negative, red)', '$#,##0.00;[Red]-$#,##0.00', -1234.5, '-$1,234.50 (red)'],
  ['Accounting', ACCOUNTING_FORMAT, 1234.5, '"$" at the left edge, "1,234.50 " at the right'],
  ['Accounting (negative)', ACCOUNTING_FORMAT, -1234.5, '"$" at the left edge, "(1,234.50)" at the right'],
  ['Accounting (zero)', ACCOUNTING_FORMAT, 0, '"$" at the left edge, "-" near the right'],
  ['Percent', '0%', 0.256, '26%'],
  ['Percent, 2 decimals', '0.00%', 0.0525, '5.25%'],
  ['Percent (negative, red)', '0.0%;[Red]-0.0%', -0.034, '-3.4% (red)'],
  ['Scientific', '0.00E+00', 12345.678, '1.23E+04'],
  ['Scientific (small)', '0.00E+00', 0.000123, '1.23E-04'],
  ['Fraction', '# ?/?', 1.25, '1 1/4'],
  ['Fraction, 2 digits (space padded)', '# ??/??', 3.14159, '3  1/7 '],
  ['Fraction in eighths', '# ?/8', 2.375, '2 3/8'],
  ['Date m/d/yyyy', 'm/d/yyyy', SAMPLE_DATE, '3/15/2024'],
  ['Date d-mmm-yy', 'd-mmm-yy', SAMPLE_DATE, '15-Mar-24'],
  ['Date + time', 'yyyy-mm-dd hh:mm', SAMPLE_DATE_TIME, '2024-03-15 14:30'],
  ['Long date', 'dddd, mmmm d, yyyy', SAMPLE_DATE, 'Friday, March 15, 2024'],
  ['Date with CJK literals', 'yyyy"年"m"月"d"日"', SAMPLE_DATE, '2024年3月15日'],
  ['Time 12-hour', 'h:mm AM/PM', 14.5 / 24, '2:30 PM'],
  ['Time with seconds', 'h:mm:ss', 18 / 24 + 15 / 86400, '18:00:15'],
  ['Elapsed hours', '[h]:mm:ss', 1.5, '36:00:00'],
  ['Text (string)', '@', '00123', '00123'],
  ['Text (number)', '@', 42, '42'],
  ['Custom colors (positive)', '[Blue]0;[Red]-0;"zero"', 5, '5 (blue)'],
  ['Custom colors (negative)', '[Blue]0;[Red]-0;"zero"', -5, '-5 (red)'],
  ['Custom colors (zero)', '[Blue]0;[Red]-0;"zero"', 0, 'zero'],
  ['Unit suffix', '0.0 "kg"', 72.456, '72.5 kg'],
  ['Leading zeros', '00000', 501, '00501'],
  ['Digit groups', '000-00-0000', 123456789, '123-45-6789'],
  ['Thousands scaling', '#,##0,"K"', 1234567, '1,235K'],
  ['Conditional sections (millions)', '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', 2500000, '2.5M'],
  ['Conditional sections (thousands)', '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', 4200, '4.2K'],
  ['Negative in parentheses', '#,##0;(#,##0)', -4321, '(4,321)'],
];

/** Sheet 2 — number formats with expected output, then formulas with cached results (shared formula column). */
function addNumbersSheet(wb) {
  const ws = wb.addWorksheet('Numbers', {
    properties: { tabColor: ARGB('70AD47') },
    views: [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }],
  });
  [34, 46, 24, 16, 44, 14].forEach((width, i) => (ws.getColumn(i + 1).width = width));

  ['Description', 'Format code', 'Formatted value', 'Raw value', 'Excel displays (en-US)'].forEach((text, i) =>
    put(ws, cellAddress(1, i + 1), text, HEADER_STYLE),
  );
  ws.getRow(1).height = 20;

  let r = 2;
  for (const [description, format, value, expected] of NUMBER_FORMATS) {
    put(ws, `A${r}`, description);
    put(ws, `B${r}`, format, { numFmt: '@', font: font({ name: 'Consolas', family: 3, size: 10 }) });
    put(ws, `C${r}`, value, { numFmt: format });
    put(ws, `D${r}`, value instanceof Date ? excelSerial(value) : value, typeof value === 'string' ? { numFmt: '@' } : {});
    put(ws, `E${r}`, expected, { numFmt: '@', font: font({ color: ARGB('595959') }) });
    r++;
  }

  // --- Formulas: a small table with shared formulas, then single formulas with cached results ---
  r += 1;
  put(ws, `A${r}`, 'Formulas (cached results)', { font: SECTION_FONT, fill: SECTION_FILL });
  for (let c = 2; c <= 6; c++) put(ws, cellAddress(r, c), undefined, { fill: SECTION_FILL });
  r += 1;
  const tableHeader = r;
  ['Product', 'Qty', 'Unit price', 'Total (shared formula)', 'Share (shared, $abs$)', 'In stock'].forEach((text, i) =>
    put(ws, cellAddress(tableHeader, i + 1), text, HEADER_STYLE),
  );
  const products = [
    ['Apples', 10, 0.5, true],
    ['Bananas', 24, 0.25, true],
    ['Cherries', 5, 3.2, false],
    ['Dates', 12, 1.75, true],
    ['Elderberries', 3, 4.5, false],
    ['Figs', 8, 2.1, true],
  ];
  const first = tableHeader + 1;
  const last = tableHeader + products.length;
  const totalRow = last + 1;
  const totals = products.map(([, qty, price]) => excelRound(qty * price));
  const grandTotal = excelRound(totals.reduce((a, b) => a + b, 0));
  const totalQty = products.reduce((a, p) => a + p[1], 0);
  products.forEach(([name, qty, price, inStock], i) => {
    const row = first + i;
    put(ws, `A${row}`, name);
    put(ws, `B${row}`, qty);
    put(ws, `C${row}`, price, { numFmt: '$#,##0.00' });
    // D/E: one master formula each, the other rows are shared-formula clones (<f t="shared" si=…/>).
    put(ws, `D${row}`, i === 0 ? { formula: `B${row}*C${row}`, result: totals[i] } : { sharedFormula: `D${first}`, result: totals[i] }, { numFmt: '$#,##0.00' });
    put(ws, `E${row}`, i === 0 ? { formula: `D${row}/$D$${totalRow}`, result: excelRound(totals[i] / grandTotal) } : { sharedFormula: `E${first}`, result: excelRound(totals[i] / grandTotal) }, { numFmt: '0.0%' });
    put(ws, `F${row}`, inStock, { alignment: { horizontal: 'center' } });
  });
  const totalStyle = { font: font({ bold: true }), border: { top: edge('thin'), bottom: edge('double') } };
  put(ws, `A${totalRow}`, 'Total', totalStyle);
  put(ws, `B${totalRow}`, { formula: `SUM(B${first}:B${last})`, result: totalQty }, totalStyle);
  put(ws, `C${totalRow}`, undefined, totalStyle);
  put(ws, `D${totalRow}`, { formula: `SUM(D${first}:D${last})`, result: grandTotal }, { ...totalStyle, numFmt: '$#,##0.00' });
  put(ws, `E${totalRow}`, { formula: `SUM(E${first}:E${last})`, result: 1 }, { ...totalStyle, numFmt: '0.0%' });
  put(ws, `F${totalRow}`, undefined, totalStyle);

  r = totalRow + 2;
  ['Example', 'Formula', 'Result'].forEach((text, i) => put(ws, cellAddress(r, i + 1), text, HEADER_STYLE));
  r += 1;
  const averagePrice = excelRound(products.reduce((a, p) => a + p[2], 0) / products.length);
  const formulaRows = [
    ['SUM', `SUM(B${first}:B${last})`, totalQty],
    ['AVERAGE', `AVERAGE(C${first}:C${last})`, averagePrice, '0.00'],
    ['IF (text result)', `IF(D${totalRow}>50,"Large order","Small order")`, grandTotal > 50 ? 'Large order' : 'Small order'],
    ['VLOOKUP', `VLOOKUP("Cherries",A${first}:D${last},4,FALSE)`, totals[2], '$#,##0.00'],
    ['CONCAT (Excel 2016+, _xlfn prefix)', `_xlfn.CONCAT(A${first}," x ",B${first})`, `${products[0][0]} x ${products[0][1]}`],
    ['& operator with TEXT()', `A${first + 1}&" @ "&TEXT(C${first + 1},"$0.00")`, `${products[1][0]} @ $${products[1][2].toFixed(2)}`],
    ['CONCATENATE', `CONCATENATE("Qty: ",B${first + 2})`, `Qty: ${products[2][1]}`],
    ['Division by zero', `B${first}/0`, { error: '#DIV/0!' }],
    ['Lookup miss', `VLOOKUP("Kiwi",A${first}:D${last},4,FALSE)`, { error: '#N/A' }],
    ['Boolean result', `B${first}>5`, products[0][1] > 5],
    ['Date result', 'DATE(2024,3,15)+30', new Date(Date.UTC(2024, 3, 14)), 'd-mmm-yy'],
    ['ROUND(PI(),4)', 'ROUND(PI(),4)', 3.1416, '0.0000'],
    ['No cached result', '1+1', undefined, undefined, 'No <v> in the file: blank until recalculated (Excel shows 2)'],
  ];
  for (const [description, formula, result, numFmt, note] of formulaRows) {
    put(ws, `A${r}`, description);
    put(ws, `B${r}`, `=${formula}`, { numFmt: '@', font: font({ name: 'Consolas', family: 3, size: 10 }) });
    put(ws, `C${r}`, result === undefined ? { formula } : { formula, result }, numFmt ? { numFmt } : {});
    if (note) put(ws, `D${r}`, note, { font: NOTE_FONT });
    r++;
  }
  put(ws, `A${r}`, 'Boolean values');
  put(ws, `B${r}`, 'TRUE / FALSE literals', { font: NOTE_FONT });
  put(ws, `C${r}`, true);
  put(ws, `D${r}`, false);
  r++;
  put(ws, `A${r}`, 'Error value (literal)');
  put(ws, `B${r}`, '#N/A stored as a value', { font: NOTE_FONT });
  put(ws, `C${r}`, { error: '#N/A' });
  return ws;
}

// ===== STYLED WORKBOOK: CONDITIONAL FORMATS SHEET =====

/** Sheet 3 — a data table with one conditional-format rule type per column, plus an auto filter. */
function addCondFormatSheet(wb) {
  const ws = wb.addWorksheet('CondFormat', {
    properties: { tabColor: { theme: 7 }, defaultColWidth: 12 },
    views: [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'B2' }],
  });
  const headers = ['Name', 'Score', 'Item', 'Sales', 'Code', 'Temp °C', 'Growth', 'Progress', 'Trend', 'vs Avg'];
  headers.forEach((text, i) => put(ws, cellAddress(1, i + 1), text, { ...HEADER_STYLE, alignment: CENTER }));
  ws.getColumn(1).width = 14;
  ws.getColumn(3).width = 14;
  ws.getColumn(12).width = 58;

  const random = createRandom(20240315);
  const names = ['Alice', 'Bob', 'Chandra', 'Dmitri', 'Elena', 'Farah', 'Goran', 'Hiro', 'Ines', 'Jamal', 'Kirsten', 'Luis', 'Mei', 'Nadia', 'Oskar', 'Priya', 'Quentin', 'Rosa', 'Sven', 'Tariq'];
  const items = ['Apple', 'Pineapple', 'Banana', 'Apple pie', 'Grape', 'Crabapple', 'Mango', 'Kiwi', 'Pear', 'Applesauce'];
  const codes = ['A1', 'B2', 'C3', 'D4', 'E5', 'F6', 'G7', 'H8', 'J9', 'K10', 'L11', 'M12', 'N13', 'P14'];
  const first = 2;
  const last = first + names.length - 1;
  names.forEach((name, i) => {
    const row = first + i;
    put(ws, `A${row}`, name);
    put(ws, `B${row}`, random.int(1, 100), { alignment: { horizontal: 'center' } });
    put(ws, `C${row}`, items[i % items.length]);
    put(ws, `D${row}`, random.int(100, 10000), { numFmt: '#,##0' });
    put(ws, `E${row}`, codes[random.int(0, codes.length - 1)], { alignment: { horizontal: 'center' } });
    put(ws, `F${row}`, random.float(-10, 40, 1), { numFmt: '0.0' });
    put(ws, `G${row}`, random.float(-0.5, 0.5, 3), { numFmt: '0.0%' });
    put(ws, `H${row}`, random.float(0, 1, 2), { numFmt: '0%' });
    put(ws, `I${row}`, random.int(-100, 100));
    put(ws, `J${row}`, random.int(0, 200));
  });
  // Guarantee at least one duplicate pair in the Code column whatever the PRNG produced.
  ws.getCell(`E${last}`).value = ws.getCell(`E${first}`).value;

  const range = column => `${column}${first}:${column}${last}`;
  const fillStyle = (fillHex, fontHex) => ({
    fill: { type: 'pattern', pattern: 'solid', bgColor: ARGB(fillHex) },
    font: { color: ARGB(fontHex) },
  });
  ws.addConditionalFormatting({
    ref: range('A'),
    rules: [{ type: 'expression', formulae: [`$B${first}>90`], style: { font: { bold: true, italic: true, color: ARGB('7030A0') } } }],
  });
  ws.addConditionalFormatting({
    ref: range('B'),
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: [75], style: fillStyle('FFC7CE', '9C0006') },
      { type: 'cellIs', operator: 'between', formulae: [40, 60], style: fillStyle('FFEB9C', '9C5700') },
    ],
  });
  ws.addConditionalFormatting({
    ref: range('C'),
    rules: [{ type: 'containsText', operator: 'containsText', text: 'apple', style: fillStyle('C6EFCE', '006100') }],
  });
  ws.addConditionalFormatting({
    ref: range('D'),
    rules: [
      { type: 'top10', rank: 3, percent: false, bottom: false, style: { font: { bold: true, color: ARGB('1F4E78') }, fill: { type: 'pattern', pattern: 'solid', bgColor: ARGB('BDD7EE') } } },
      { type: 'top10', rank: 3, percent: false, bottom: true, style: fillStyle('FCE4D6', '833C0B') },
    ],
  });
  ws.addConditionalFormatting({
    ref: range('E'),
    // Rewritten to <cfRule type="duplicateValues"> by patchStyledWorkbook (ExcelJS cannot write that type).
    rules: [{ type: 'expression', formulae: [DUPLICATE_VALUES_SENTINEL], style: fillStyle('FFC7CE', '9C0006') }],
  });
  ws.addConditionalFormatting({
    ref: range('F'),
    rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], color: [ARGB('FFFFFF'), ARGB('63BE7B')] }],
  });
  ws.addConditionalFormatting({
    ref: range('G'),
    rules: [{
      type: 'colorScale',
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [ARGB('F8696B'), ARGB('FFEB84'), ARGB('63BE7B')],
    }],
  });
  ws.addConditionalFormatting({
    ref: range('H'),
    rules: [{ type: 'dataBar', gradient: true, cfvo: [{ type: 'min' }, { type: 'max' }], color: ARGB('638EC6') }],
  });
  ws.addConditionalFormatting({
    ref: range('I'),
    rules: [{
      type: 'iconSet',
      iconSet: '3Arrows',
      showValue: true,
      cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }],
    }],
  });
  ws.addConditionalFormatting({
    ref: range('J'),
    rules: [{ type: 'aboveAverage', aboveAverage: true, style: fillStyle('E2EFDA', '375623') }],
  });
  ws.autoFilter = `A1:J${last}`;

  const legend = [
    'Conditional formatting rules',
    'A: expression $B2>90 → bold purple italic',
    'B: cellIs greaterThan 75 → red; between 40 and 60 → yellow',
    'C: containsText "apple" → green',
    'D: top 3 → blue; bottom 3 → orange',
    'E: duplicateValues → red',
    'F: 2-color scale white → green',
    'G: 3-color scale red → yellow (50th percentile) → green',
    'H: data bar (gradient, blue)',
    'I: icon set 3Arrows (0 / 33 / 67 percent)',
    'J: aboveAverage → green',
    'Row 1 has an auto filter (A1:J21).',
  ];
  legend.forEach((text, i) => put(ws, `L${i + 1}`, text, i === 0 ? { font: font({ bold: true }) } : { font: NOTE_FONT }));
  return ws;
}

// ===== STYLED WORKBOOK: HIDDEN, IMAGES, RTL SHEETS =====

/** Sheets 4 and 7 — a hidden sheet (source of the FruitList name) and a veryHidden one. */
function addHiddenSheet(wb) {
  const ws = wb.addWorksheet('Hidden', { state: 'hidden', properties: { tabColor: ARGB('7F7F7F') } });
  ws.getColumn(1).width = 44;
  put(ws, 'A1', 'Fruits (source of the FruitList defined name)', { font: font({ bold: true }) });
  ['Lemon', 'Lime', 'Mango', 'Orange', 'Peach', 'Pear'].forEach((fruit, i) => put(ws, `A${2 + i}`, fruit));
  put(ws, 'A9', 'This sheet has state="hidden": not shown as a tab, listed by Excel\'s Unhide dialog.', { font: NOTE_FONT });
  wb.definedNames.add('Hidden!$A$2:$A$7', 'FruitList');
  return ws;
}

function addVeryHiddenSheet(wb) {
  const ws = wb.addWorksheet('VeryHidden', { state: 'veryHidden' });
  ws.getColumn(1).width = 70;
  put(ws, 'A1', 'This sheet has state="veryHidden": Excel\'s Unhide dialog does not list it.', { font: NOTE_FONT });
  return ws;
}

/** Sheet 5 — embedded PNGs (one-cell, two-cell and offset anchors), gridlines off, zoom 125 %. */
function addImagesSheet(wb, images) {
  const ws = wb.addWorksheet('Images', {
    properties: { tabColor: ARGB('ED7D31') },
    views: [{ showGridLines: false, zoomScale: 125, zoomScaleNormal: 125 }],
  });
  ws.getColumn(1).width = 4;
  put(ws, 'B1', 'Embedded images (gridlines off, zoom 125 %)', { font: font({ bold: true, size: 14, color: { theme: 4, tint: -0.25 } }) });
  put(ws, 'B2', 'One-cell anchor at B4, 120 × 80 px', { font: NOTE_FONT });
  put(ws, 'E2', 'Two-cell anchor E4:H12 (same picture, stretched)', { font: NOTE_FONT });
  put(ws, 'B14', 'Anchor B15 + offset, 240 × 144 px', { font: NOTE_FONT });

  const gradientId = wb.addImage({ buffer: images.gradient, extension: 'png' });
  const chartId = wb.addImage({ buffer: images.chart, extension: 'png' });
  ws.addImage(gradientId, { tl: { col: 1, row: 3 }, ext: { width: 120, height: 80 }, editAs: 'oneCell' });
  ws.addImage(gradientId, { tl: { col: 4, row: 3 }, br: { col: 8, row: 12 }, editAs: 'twoCell' });
  ws.addImage(chartId, { tl: { col: 1.5, row: 14.25 }, ext: { width: 240, height: 144 } });
  return ws;
}

/** Sheet 6 — a right-to-left sheet with Arabic and Hebrew text. */
function addRtlSheet(wb) {
  const ws = wb.addWorksheet('RTL', {
    properties: { tabColor: { theme: 5, tint: -0.25 } },
    views: [{ rightToLeft: true }],
  });
  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 26;
  ws.getColumn(3).width = 16;
  [['اللغة', 'النص', 'الرقم'], ['العربية', 'مرحبا بالعالم', 1234.5], ['עברית', 'שלום עולם', -42], ['English', 'Hello, world', 0.5]].forEach(
    (values, i) =>
      values.forEach((value, j) =>
        put(ws, cellAddress(i + 1, j + 1), value, i === 0 ? HEADER_STYLE : typeof value === 'number' ? { numFmt: '#,##0.00' } : {}),
      ),
  );
  put(ws, 'A6', 'Sheet view rightToLeft="1": column A is on the right.', { font: NOTE_FONT });
  return ws;
}

// ===== STYLED WORKBOOK: BUILD & PATCH =====

/** Builds the whole showcase workbook in memory (ExcelJS model). */
function buildStyledWorkbook(images) {
  const wb = new ExcelJS.Workbook();
  wb.creator = GENERATOR_NAME;
  wb.lastModifiedBy = GENERATOR_NAME;
  wb.created = FIXED_DATE;
  wb.modified = FIXED_DATE;
  wb.title = 'FileStudio styled showcase';
  wb.subject = 'Formatting, number formats, formulas, conditional formats, images';
  wb.keywords = 'file-viewer sample exceljs';
  wb.views = [{ x: 0, y: 0, width: 28800, height: 17000, firstSheet: 0, activeTab: 0, visibility: 'visible' }];

  addFormattingSheet(wb); // 1
  addNumbersSheet(wb); // 2
  addCondFormatSheet(wb); // 3
  addHiddenSheet(wb); // 4
  addImagesSheet(wb, images); // 5
  addRtlSheet(wb); // 6
  addVeryHiddenSheet(wb); // 7
  return wb;
}

/** Parses the attributes of one XML start tag into an ordered list of [name, value]. */
function parseAttributes(tag) {
  const attributes = [];
  const pattern = /([\w:.-]+)="([^"]*)"/g;
  let match;
  while ((match = pattern.exec(tag))) attributes.push([match[1], match[2]]);
  return attributes;
}

/**
 * Post-processes the ExcelJS output so it matches what Excel itself writes:
 *  - modern Office theme colors (accent1 4472C4 …) instead of ExcelJS's 2007 palette,
 *  - a real `duplicateValues` conditional-format rule (ExcelJS has no writer for that type),
 *  - no dangling x14 extension reference on the (classic) data bar rule, no editAs on one-cell anchors,
 *  - internal hyperlinks: `location="Sheet!A1"` only (ExcelJS also adds a bogus external relationship);
 *    links whose target starts with '#' keep the relationship form `Target="#Sheet!A1"` instead,
 *  - fixed zip timestamps (byte-identical output).
 */
async function patchStyledWorkbook(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const sheetRelsToDrop = new Map(); // rels path -> Set of rIds to remove
  const sheetXml = new Map();
  let duplicateRules = 0;

  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
    let xml = await zip.file(name).async('string');
    xml = xml.replace(
      /<cfRule type="expression" dxfId="(\d+)" priority="(\d+)"><formula>([^<]*)<\/formula><\/cfRule>/g,
      (whole, dxfId, priority, formula) => {
        if (formula !== DUPLICATE_VALUES_SENTINEL) return whole;
        duplicateRules++;
        return `<cfRule type="duplicateValues" dxfId="${dxfId}" priority="${priority}"/>`;
      },
    );
    // ExcelJS always adds an x14 extension reference to dataBar rules, even when it writes no x14 rule for it.
    xml = xml.replace(/<extLst><ext uri="\{B025F937-C7B1-47D3-B67F-A62EFF666E3E\}"[^>]*><x14:id\/><\/ext><\/extLst>/g, '');
    const relsPath = name.replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/') + '.rels';
    xml = xml.replace(/<hyperlink\b[^>]*\/>/g, tag => {
      const attributes = parseAttributes(tag);
      const location = attributes.find(([key]) => key === 'location');
      if (!location) return tag;
      const relId = attributes.find(([key]) => key === 'r:id');
      let kept;
      if (location[1].startsWith('#')) {
        kept = attributes.filter(([key]) => key !== 'location'); // relationship form, Target="#Sheet!A1"
      } else {
        kept = attributes.filter(([key]) => key !== 'r:id'); // Excel's native form
        if (relId) {
          if (!sheetRelsToDrop.has(relsPath)) sheetRelsToDrop.set(relsPath, new Set());
          sheetRelsToDrop.get(relsPath).add(relId[1]);
        }
      }
      return `<hyperlink ${kept.map(([key, value]) => `${key}="${value}"`).join(' ')}/>`;
    });
    sheetXml.set(name, xml);
  }
  if (duplicateRules !== 1) {
    throw new Error(`Expected to rewrite 1 duplicateValues rule, rewrote ${duplicateRules} (ExcelJS output changed?)`);
  }

  return repackZip(buffer, (name, content) => {
    if (sheetXml.has(name)) return sheetXml.get(name);
    if (/^xl\/drawings\/drawing\d+\.xml$/.test(name)) {
      // CT_OneCellAnchor has no editAs attribute (only twoCellAnchor does); ExcelJS writes one anyway.
      return content.toString('utf8').replace(/<xdr:oneCellAnchor editAs="[^"]*">/g, '<xdr:oneCellAnchor>');
    }
    if (name === 'xl/theme/theme1.xml') {
      const xml = content.toString('utf8');
      const patched = xml
        .replace(/<a:clrScheme name="Office">[\s\S]*?<\/a:clrScheme>/, MODERN_THEME_CLR_SCHEME)
        .replace('<a:latin typeface="Cambria"/>', '<a:latin typeface="Calibri Light"/>');
      if (!patched.includes('<a:srgbClr val="4472C4"/>')) throw new Error('Could not patch the theme color scheme');
      return patched;
    }
    if (sheetRelsToDrop.has(name)) {
      const drop = sheetRelsToDrop.get(name);
      return content
        .toString('utf8')
        .replace(/<Relationship\b[^>]*\/>/g, rel => (drop.has((/\bId="([^"]+)"/.exec(rel) || [])[1]) ? '' : rel));
    }
    return undefined;
  });
}

async function generateStyledXlsx(images) {
  const wb = buildStyledWorkbook(images);
  const raw = Buffer.from(await wb.xlsx.writeBuffer());
  return patchStyledWorkbook(raw);
}

// ===== LARGE WORKBOOK (STREAMING) =====

const LARGE_KINDS = [
  { kind: 'text', label: 'Text', width: 18 },
  { kind: 'int', label: 'Integer', width: 11, numFmt: '#,##0' },
  { kind: 'decimal', label: 'Decimal', width: 12, numFmt: '#,##0.00' },
  { kind: 'date', label: 'Date', width: 12, numFmt: 'yyyy-mm-dd' },
  { kind: 'category', label: 'Category', width: 11 },
  { kind: 'currency', label: 'Amount', width: 13, numFmt: '$#,##0.00;[Red]-$#,##0.00' },
  { kind: 'percent', label: 'Percent', width: 10, numFmt: '0.0%' },
  { kind: 'bool', label: 'Flag', width: 8 },
];
const LARGE_WORDS = ['alpha', 'bravo', 'charlie', 'delta', 'echo', 'foxtrot', 'golf', 'hotel', 'india', 'juliet', 'kilo', 'lima', 'mike', 'november', 'oscar', 'papa'];
const LARGE_CATEGORIES = ['North', 'South', 'East', 'West', 'Central'];
const LARGE_DATE_BASE = Date.UTC(2020, 0, 1);

/**
 * Streams `rows` x `cols` cells to `file` with ExcelJS's WorkbookWriter (styles on, inline strings).
 * The file is written to a temporary name first so an interrupted run never leaves a truncated large.xlsx.
 */
async function generateLargeXlsx(file, rows, cols, onProgress) {
  const temp = `${file}.tmp-${process.pid}`;
  try {
    await writeLargeXlsx(temp, rows, cols, onProgress);
    fs.renameSync(temp, file);
  } catch (error) {
    fs.rmSync(temp, { force: true });
    throw error;
  }
}

/** Writes the streaming workbook body; see generateLargeXlsx. */
async function writeLargeXlsx(temp, rows, cols, onProgress) {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: temp, useStyles: true, useSharedStrings: false });
  wb.creator = GENERATOR_NAME;
  wb.created = FIXED_DATE;
  wb.modified = FIXED_DATE;
  const ws = wb.addWorksheet('Data', {
    views: [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }],
  });

  const columns = [];
  for (let c = 0; c < cols; c++) {
    const spec = c === 0 ? { kind: 'id', label: 'ID', width: 9, numFmt: '0' } : LARGE_KINDS[(c - 1) % LARGE_KINDS.length];
    columns.push({ ...spec, header: c === 0 ? 'ID' : `${spec.label} ${c}` });
  }
  ws.columns = columns.map(spec => ({ width: spec.width, style: spec.numFmt ? { numFmt: spec.numFmt } : {} }));

  const headerStyle = {
    font: font({ bold: true, color: { theme: 0 } }),
    fill: solidFill({ theme: 4 }),
    alignment: { vertical: 'middle' },
    border: { bottom: edge('thin', { theme: 4, tint: -0.5 }) },
  };
  const header = ws.addRow(columns.map(spec => spec.header));
  header.eachCell(cell => (cell.style = { ...headerStyle, numFmt: 'General' }));
  header.height = 20;
  header.commit();

  const random = createRandom(42);
  const progressStep = Math.max(1, Math.floor(rows / 10));
  const values = new Array(cols);
  for (let r = 1; r <= rows; r++) {
    for (let c = 0; c < cols; c++) {
      switch (columns[c].kind) {
        case 'id':
          values[c] = r;
          break;
        case 'text':
          values[c] = `${LARGE_WORDS[(r + c) % LARGE_WORDS.length]} ${LARGE_WORDS[random.int(0, LARGE_WORDS.length - 1)]} ${r}`;
          break;
        case 'int':
          values[c] = random.int(-5000, 50000);
          break;
        case 'decimal':
          values[c] = random.float(-1000, 100000, 2);
          break;
        case 'date':
          values[c] = new Date(LARGE_DATE_BASE + ((r * 7 + c * 13) % 2500) * 86400000);
          break;
        case 'category':
          values[c] = LARGE_CATEGORIES[(r * 3 + c) % LARGE_CATEGORIES.length];
          break;
        case 'currency':
          values[c] = random.float(-2500, 25000, 2);
          break;
        case 'percent':
          values[c] = random.float(-1, 1, 4);
          break;
        case 'bool':
          values[c] = (r + c) % 3 === 0;
          break;
      }
    }
    ws.addRow(values).commit();
    if (onProgress && r % progressStep === 0) onProgress(r, rows);
  }
  ws.commit();
  await wb.commit();
}

// ===== CSV / TSV / PSV / SSV =====

/** Quotes a field when needed (delimiter, quote, CR/LF, surrounding whitespace) or when forced. */
function delimitedField(value, delimiter, forceQuote) {
  const text = value === null || value === undefined ? '' : String(value);
  const needsQuotes = forceQuote || text.includes(delimiter) || /["\r\n]/.test(text) || /^\s|\s$/.test(text);
  return needsQuotes ? `"${text.replace(/"/g, '""')}"` : text;
}

/** Rows are arrays of values; a value wrapped as { q: text } is always quoted. */
function toDelimited(rows, { delimiter, newline, bom = false }) {
  const lines = rows.map(row =>
    row
      .map(value => (value && typeof value === 'object' ? delimitedField(value.q, delimiter, true) : delimitedField(value, delimiter, false)))
      .join(delimiter),
  );
  return Buffer.from((bom ? '﻿' : '') + lines.join(newline) + newline, 'utf8');
}

/** sample.csv — CRLF, header + 50 rows of deliberately awkward data. */
function buildSampleCsv() {
  const header = ['id', 'name', 'email', 'city', 'country', 'amount', 'change_pct', 'joined', 'active', 'notes'];
  const longNote = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.';
  const q = text => ({ q: text });
  const rows = [
    [1, 'Smith, John', 'john.smith@example.com', 'London', 'UK', '1234.56', '-2.5', '2024-01-15', 'true', 'Comma inside the name'],
    [2, 'Zoë Müller', 'zoe.muller@example.com', 'Zürich', 'Switzerland', '-42.50', '0.75', '2023-11-02', 'false', 'She said "hello" twice'],
    [3, "O'Brien, Seán", 'sean.obrien@example.com', 'Dublin', 'Ireland', '0', '0', '2022-06-30', 'true', 'Line one\nLine two (embedded LF)'],
    [4, '李小龍', 'lee@example.com', '香港', 'China', '98765.4321', '12.345', '2021-02-03', 'true', 'CJK characters'],
    [5, 'Αλέξανδρος', 'alex@example.com', 'Αθήνα', 'Greece', '', '', '2020-08-08', '', 'Empty amount, change and active fields'],
    [6, 'مريم', 'maryam@example.com', 'القاهرة', 'Egypt', '1e3', '-0.001', '2019-12-31', 'false', 'Right-to-left text and scientific notation'],
    [7, '  Padded  ', 'pad@example.com', 'New York, NY', 'USA', '1,234.56', '5%', '2018-04-01', 'true', 'Quoted thousands separator; leading/trailing spaces kept'],
    [8, 'Emoji 🎉', 'party@example.com', 'São Paulo', 'Brazil', '-0.01', '-100', '2017-07-07', 'true', 'Multi-line\r\nwith CRLF inside quotes'],
    [9, '"Quoted" name', 'quote@example.com', 'Paris', 'France', '3.14159', '3.14159', '2016-02-29', 'false', 'Field that starts with a quote'],
    [10, 'Trailing Empty', 'te@example.com', 'Berlin', 'Germany', '100', '', '2015-05-05', 'true', ''],
    [11, 'Zip Code', 'zip@example.com', 'Boston', 'USA', '00123', '0.5', '2014-10-10', 'true', 'Leading zeros (numeric-looking text)'],
    [12, q('Tab\tInside'), 'tab@example.com', 'Oslo', 'Norway', '-1234567.891', '-99.99', '2013-03-03', 'false', 'Tab character inside a quoted field'],
    [13, q('Plain quoted'), q('plain@example.com'), q('Rome'), q('Italy'), q('42'), q('1.5'), q('2012-12-12'), q('true'), q('Every field quoted even when not needed')],
    [14, '=SUM(A1:A2)', 'formula@example.com', 'Madrid', 'Spain', '7', '7', '2011-11-11', 'true', 'Formula-like text must stay text'],
    [15, 'Long Note', 'long@example.com', 'Sydney', 'Australia', '1', '1', '2010-10-10', 'true', longNote],
  ];
  const firstNames = ['Olivia', 'Liam', 'Emma', 'Noah', 'Ava', 'Mateo', 'Sofia', 'Lucas', 'Mia', 'Ethan', 'Amara', 'Kenji', 'Priya', 'Diego', 'Ingrid', 'Tomás', 'Chloé', 'Björn', 'Aisha', 'Wei'];
  const lastNames = ['Johnson', 'García', 'Nakamura', 'Okafor', 'Schmidt', 'Rossi', 'Kowalski', 'Andersson', 'Patel', 'Dubois', 'Novák', 'Silva', 'Kim', 'Murphy', 'Haddad'];
  const places = [['Toronto', 'Canada'], ['Mumbai', 'India'], ['Tokyo', 'Japan'], ['Lagos', 'Nigeria'], ['München', 'Germany'], ['Milan', 'Italy'], ['Kraków', 'Poland'], ['Stockholm', 'Sweden'], ['Lyon', 'France'], ['Prague', 'Czechia'], ['Lisbon', 'Portugal'], ['Seoul', 'South Korea'], ['Cork', 'Ireland'], ['Beirut', 'Lebanon'], ['Austin', 'USA'], ['Montréal', 'Canada']];
  const notes = ['', '', 'Follow up, high priority', 'VIP', 'Prefers e-mail', 'Paid in advance', 'Requested "express" shipping', ''];
  const random = createRandom(7);
  for (let id = rows.length + 1; id <= 50; id++) {
    const firstName = random.pick(firstNames);
    const lastName = random.pick(lastNames);
    const [city, country] = random.pick(places);
    const joined = new Date(Date.UTC(2015, 0, 1) + random.int(0, 3500) * 86400000).toISOString().slice(0, 10);
    rows.push([
      id,
      `${firstName} ${lastName}`,
      `${asciiFold(firstName)}.${asciiFold(lastName)}@example.com`,
      city,
      country,
      random.float(-5000, 50000, 2).toFixed(2),
      random.float(-50, 50, random.chance(0.5) ? 1 : 3),
      joined,
      random.chance(0.7) ? 'true' : 'false',
      random.pick(notes),
    ]);
  }
  return toDelimited([header, ...rows], { delimiter: ',', newline: '\r\n' });
}

/** sample-semicolon.csv — European style: ';' delimiter, decimal commas, LF, UTF-8 BOM. */
function buildSemicolonCsv() {
  const rows = [
    ['Artikel', 'Beschreibung', 'Menge', 'Preis (€)', 'Datum', 'Lager'],
    ['A-1001', 'Kaffeebohnen, gemahlen', '12', '8,99', '15.03.2024', 'Hamburg'],
    ['A-1002', 'Grüner Tee; lose', '5', '4,50', '16.03.2024', 'München'],
    ['A-1003', 'Schokolade "Zartbitter"', '40', '1,29', '17.03.2024', 'Köln'],
    ['A-1004', 'Brötchen', '120', '0,35', '18.03.2024', 'Düsseldorf'],
    ['A-1005', 'Käse (Gouda)', '7', '12,75', '19.03.2024', 'Bremen'],
    ['A-1006', 'Äpfel', '-3', '2,10', '20.03.2024', 'Hamburg'],
    ['A-1007', 'Weißwein 0,75 l', '24', '6,49', '21.03.2024', 'Mainz'],
    ['A-1008', 'Mehl Type 405', '0', '0,89', '22.03.2024', 'Stuttgart'],
    ['A-1009', '', '15', '', '23.03.2024', 'Leipzig'],
    ['A-1010', 'Olivenöl\nextra nativ', '9', '9,95', '24.03.2024', 'Dresden'],
    ['A-1011', 'Müsli', '18', '3,49', '25.03.2024', 'Nürnberg'],
    ['A-1012', 'Joghurt 1,5 %', '60', '0,59', '26.03.2024', 'Hannover'],
    ['A-1013', 'Senf, mittelscharf', '22', '1,15', '27.03.2024', 'Düsseldorf'],
    ['A-1014', 'Spätzle', '14', '1,99', '28.03.2024', 'Stuttgart'],
    ['A-1015', 'Gesamtsumme', '1.234', '1.234,56', '29.03.2024', 'Zentrale'],
  ];
  return toDelimited(rows, { delimiter: ';', newline: '\n', bom: true });
}

/** sample.tsv — tab separated, LF. */
function buildSampleTsv() {
  const header = ['id', 'product', 'category', 'price', 'qty', 'updated', 'comment'];
  const products = ['Widget', 'Gadget', 'Doohickey', 'Gizmo', 'Thingamajig', 'Whatsit', 'Sprocket', 'Flange'];
  const categories = ['Tools', 'Toys', 'Garden', 'Kitchen', 'Office'];
  const comments = ['', 'Back-ordered', 'Contains, a comma', 'Size: 10" x 12"', 'Ships in 2–3 days', '日本語のコメント', ''];
  const random = createRandom(11);
  const rows = [header];
  for (let id = 1; id <= 24; id++) {
    rows.push([
      id,
      `${random.pick(products)} ${columnLetter(id)}`,
      random.pick(categories),
      random.float(0.5, 250, 2).toFixed(2),
      random.int(-5, 500),
      new Date(Date.UTC(2024, 0, 1) + random.int(0, 300) * 86400000).toISOString().slice(0, 10),
      random.pick(comments),
    ]);
  }
  rows.push([25, { q: 'Tab\tseparated name' }, 'Edge cases', '0', '0', '2024-12-31', 'Quoted field containing a tab']);
  return toDelimited(rows, { delimiter: '\t', newline: '\n' });
}

/** sample.psv — pipe separated, LF, header + 30 rows; fields with pipes, quotes, a line break or padding are quoted. */
function buildSamplePsv() {
  const header = ['id', 'service', 'endpoint', 'method', 'status', 'latency_ms', 'checked_at', 'region', 'notes'];
  const rows = [
    [1, 'search', '/api/search?q=a|b', 'GET', 200, '12.5', '2024-03-15T09:30:00Z', 'eu-west', 'Query string contains a pipe'],
    [2, 'auth', '/api/login', 'POST', 401, '48.0', '2024-03-15T09:30:04Z', 'us-east', 'Rejected: "bad password"'],
    [3, 'billing', '/api/invoices', 'GET', 500, '1203.75', '2024-03-15T09:30:09Z', 'us-east', 'TimeoutError\nretried twice (embedded LF)'],
    [4, 'reports', '/api/export', 'GET', 200, '', '2024-03-15T09:30:15Z', 'ap-south', ''],
    [5, 'gateway | edge', '/', 'HEAD', 204, '0.9', '2024-03-15T09:30:16Z', 'eu-west', 'Service name contains a pipe'],
    [6, 'données', '/api/données/été', 'GET', 200, '33.3', '2024-03-15T09:30:21Z', 'eu-central', 'Unicode path: café, 東京'],
    [7, '  padded  ', '/api/padded', 'GET', 200, '7', '2024-03-15T09:30:25Z', 'eu-west', 'Leading/trailing spaces kept (quoted)'],
    [8, 'regex', '/api/match', 'POST', 422, '15.25', '2024-03-15T09:30:31Z', 'us-west', 'Pattern "^(a|b|c)$" failed'],
    [9, 'cache', '/api/cache', 'DELETE', 202, '-1', '2024-03-15T09:30:38Z', 'us-west', 'Negative latency marks a skipped check'],
    [10, 'metrics', '/metrics', 'GET', 200, '1e2', '2024-03-15T09:30:44Z', 'ap-south', { q: 'Quoted although nothing needs quoting' }],
  ];
  const services = ['search', 'auth', 'billing', 'reports', 'cache', 'metrics', 'storage', 'mail'];
  const methods = ['GET', 'GET', 'GET', 'POST', 'PUT', 'DELETE'];
  const statuses = [200, 200, 200, 200, 201, 204, 301, 404, 500, 503];
  const regions = ['eu-west', 'eu-central', 'us-east', 'us-west', 'ap-south'];
  const notes = ['', '', '', 'OK', 'Slow response', 'Retry 1|2|3', 'Timeout after "30s"', 'Cache hit'];
  const random = createRandom(23);
  let time = Date.UTC(2024, 2, 15, 9, 30, 44);
  for (let id = rows.length + 1; id <= 30; id++) {
    const service = random.pick(services);
    time += random.int(1, 9) * 1000;
    rows.push([
      id,
      service,
      `/api/${service}/${random.int(1, 999)}`,
      random.pick(methods),
      random.pick(statuses),
      random.float(0.5, 900, random.chance(0.5) ? 1 : 2).toFixed(2),
      new Date(time).toISOString().replace('.000Z', 'Z'),
      random.pick(regions),
      random.pick(notes),
    ]);
  }
  return toDelimited([header, ...rows], { delimiter: '|', newline: '\n' });
}

/** French number format with a plain space as thousands separator: 1274.48 -> "1 274,48". */
function frenchNumber(value, decimals) {
  const [whole, fraction] = Math.abs(value).toFixed(decimals).split('.');
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return `${value < 0 ? '-' : ''}${grouped}${fraction ? `,${fraction}` : ''}`;
}

/**
 * sample.ssv — ';' delimited, French conventions: decimal commas, a plain space as thousands separator (so many
 * fields contain spaces and the ';' must still win the delimiter detection), quoted ';' / quotes / CRLF in fields,
 * UTF-8 BOM, CRLF line endings. Amounts are quantity × unit price and the last row holds the totals.
 */
function buildSampleSsv() {
  const items = [
    // [reference, label, quantity, unit price, date, warehouse, remark]
    ['FR-001', 'Café moulu 250 g', 12, 4.95, '15/03/2024', 'Paris Nord', ''],
    ['FR-002', 'Thé vert; vrac', 5, 7.2, '16/03/2024', 'Lyon', 'Point-virgule dans le libellé'],
    ['FR-003', 'Chocolat "Noir 70 %"', 40, 2.15, '17/03/2024', 'Lille', 'Guillemets dans le libellé'],
    ['FR-004', 'Baguette tradition', 1200, 1.3, '18/03/2024', 'Paris Sud', 'Grosse commande'],
    ['FR-005', 'Fromage (Comté 18 mois)', 7, 24.9, '19/03/2024', 'Besançon', ''],
    ['FR-006', 'Pommes Golden', -3, 2.49, '20/03/2024', 'Rouen', 'Retour client'],
    ['FR-007', 'Vin rouge 0,75 l', 24, 8.75, '21/03/2024', 'Bordeaux', ''],
    ['FR-008', 'Farine T55', 0, 1.09, '22/03/2024', 'Toulouse', 'Rupture de stock'],
    ['FR-009', '', 15, null, '23/03/2024', 'Nantes', 'Libellé et prix manquants'],
    ['FR-010', 'Huile d’olive\r\nvierge extra', 9, 11.5, '24/03/2024', 'Marseille', 'Retour à la ligne (CRLF) dans le libellé'],
    ['FR-011', 'Confiture abricot', 18, 3.6, '25/03/2024', 'Avignon', ''],
    ['FR-012', 'Yaourt nature x 4', 60, 1.85, '26/03/2024', 'Rennes', ''],
    ['FR-013', 'Moutarde de Dijon', 22, 2.3, '27/03/2024', 'Dijon', ''],
    ['FR-014', 'Crème fraîche 30 %', 14, 2.05, '28/03/2024', 'Caen', ''],
  ];
  const rows = [['Référence', 'Libellé', 'Quantité', 'Prix unitaire (€)', 'Montant (€)', 'Date', 'Entrepôt', 'Remarque']];
  let totalQuantity = 0;
  let totalCents = 0;
  for (const [reference, label, quantity, price, date, warehouse, remark] of items) {
    const cents = price === null ? 0 : Math.round(quantity * price * 100);
    totalQuantity += quantity;
    totalCents += cents;
    rows.push([
      reference,
      label,
      frenchNumber(quantity, 0),
      price === null ? '' : frenchNumber(price, 2),
      price === null ? '' : frenchNumber(cents / 100, 2),
      date,
      warehouse,
      remark,
    ]);
  }
  rows.push(['TOTAL', 'Total général', frenchNumber(totalQuantity, 0), '', frenchNumber(totalCents / 100, 2), '29/03/2024', 'Siège', 'Séparateur de milliers : espace']);
  return toDelimited(rows, { delimiter: ';', newline: '\r\n', bom: true });
}

/**
 * Space-separated, column-aligned text: every column is padded to its widest value (two spaces between columns,
 * so runs of spaces separate fields and most lines end with spaces); right-aligned columns give leading spaces.
 * Fields that are empty or contain a space, quote, tab or line break are "quoted" ("" for an empty field).
 */
function toAligned(rows, { align, newline }) {
  const cells = rows.map(row =>
    row.map(value => {
      const text = String(value);
      return text === '' || /[\s"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
    }),
  );
  const widths = cells[0].map((_, c) => Math.max(...cells.map(row => row[c].length)));
  const lines = cells.map(row => row.map((text, c) => (align[c] === 'right' ? text.padStart(widths[c]) : text.padEnd(widths[c]))).join('  '));
  return Buffer.from(lines.join(newline) + newline, 'utf8');
}

/** sample-space.ssv — space separated, aligned columns, quoted "two words" fields, LF, no BOM. */
function buildSpaceSsv() {
  const rows = [
    ['rank', 'name', 'born', 'city', 'field', 'score', 'note'],
    [1, 'Ada Lovelace', '1815-12-10', 'London', 'Mathematics', '98.5', 'First published algorithm'],
    [2, 'Alan Turing', '1912-06-23', 'London', 'Computer science', '97.25', ''],
    [3, 'Grace Hopper', '1906-12-09', 'New York', 'Compilers', '96', 'Popularised "debugging"'],
    [4, 'Katherine Johnson', '1918-08-26', 'White Sulphur Springs', 'Mathematics', '95.75', 'Orbital mechanics'],
    [5, 'Edsger Dijkstra', '1930-05-11', 'Rotterdam', 'Algorithms', '94', 'Shortest paths; semaphores'],
    [6, 'Barbara Liskov', '1939-11-07', 'Los Angeles', 'Programming languages', '93.5', 'Substitution'],
    [7, 'Donald Knuth', '1938-01-10', 'Milwaukee', 'Algorithms', '93', 'TeX'],
    [8, 'Margaret Hamilton', '1936-08-17', 'Paoli', 'Software engineering', '92.25', 'Apollo'],
    [9, 'John McCarthy', '1927-09-04', 'Boston', 'Artificial intelligence', '91', 'Lisp'],
    [10, 'Kurt Gödel', '1906-04-28', 'Brno', 'Logic', '90.5', 'Incompleteness theorems'],
    [11, 'Ole-Johan Dahl', '1931-10-12', 'Mandal', 'Programming languages', '', 'Simula (no score)'],
    [12, 'Euclid', '', 'Alexandria', 'Geometry', '89', 'Birth date unknown'],
  ];
  return toAligned(rows, { align: ['right', 'left', 'left', 'left', 'left', 'right', 'left'], newline: '\n' });
}

// ===== DOCX: XML HELPERS =====

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};
const EMU_PER_PX = 9525; // 914400 EMU per inch / 96 px per inch

/** Run properties in schema order (rStyle, rFonts, b, i, strike, color, sz, highlight, u, vertAlign). */
function runProperties(p = {}) {
  let xml = '';
  if (p.style) xml += `<w:rStyle w:val="${p.style}"/>`;
  if (p.font) xml += `<w:rFonts w:ascii="${p.font}" w:hAnsi="${p.font}" w:cs="${p.font}"/>`;
  if (p.b) xml += '<w:b/><w:bCs/>';
  if (p.i) xml += '<w:i/><w:iCs/>';
  if (p.strike) xml += '<w:strike/>';
  if (p.color) xml += `<w:color w:val="${p.color}"/>`;
  if (p.sz) xml += `<w:sz w:val="${p.sz}"/><w:szCs w:val="${p.sz}"/>`;
  if (p.highlight) xml += `<w:highlight w:val="${p.highlight}"/>`;
  if (p.u) xml += `<w:u w:val="${p.u === true ? 'single' : p.u}"/>`;
  if (p.vertAlign) xml += `<w:vertAlign w:val="${p.vertAlign}"/>`;
  return xml ? `<w:rPr>${xml}</w:rPr>` : '';
}

/** A text run; '\t' becomes <w:tab/> and '\n' a line break. */
function run(text, props) {
  const body = String(text)
    .split(/(\t|\n)/)
    .map(part => (part === '\t' ? '<w:tab/>' : part === '\n' ? '<w:br/>' : part ? `<w:t xml:space="preserve">${xmlEscape(part)}</w:t>` : ''))
    .join('');
  return `<w:r>${runProperties(props)}${body}</w:r>`;
}

/** A paragraph. `content` is run XML (string or array); options map to pPr children in schema order. */
function paragraph(content, options = {}) {
  let pPr = '';
  if (options.style) pPr += `<w:pStyle w:val="${options.style}"/>`;
  if (options.keepNext) pPr += '<w:keepNext/>';
  if (options.numId !== undefined) pPr += `<w:numPr><w:ilvl w:val="${options.ilvl || 0}"/><w:numId w:val="${options.numId}"/></w:numPr>`;
  if (options.spacingAfter !== undefined) pPr += `<w:spacing w:after="${options.spacingAfter}"/>`;
  if (options.jc) pPr += `<w:jc w:val="${options.jc}"/>`;
  const body = Array.isArray(content) ? content.join('') : content || '';
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}</w:p>`;
}

function hyperlink(runs, { relId, anchor }) {
  const target = relId ? `r:id="${relId}"` : `w:anchor="${anchor}"`;
  return `<w:hyperlink ${target} w:history="1">${Array.isArray(runs) ? runs.join('') : runs}</w:hyperlink>`;
}

function inlineImage({ relId, id, name, description, widthPx, heightPx }) {
  const cx = widthPx * EMU_PER_PX;
  const cy = heightPx * EMU_PER_PX;
  return (
    '<w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${id}" name="Picture ${id}" descr="${xmlEscape(description)}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic><a:graphicData uri="${NS.pic}">` +
    '<pic:pic>' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${xmlEscape(name)}" descr="${xmlEscape(description)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic>' +
    '</a:graphicData></a:graphic>' +
    '</wp:inline>' +
    '</w:drawing></w:r>'
  );
}

/** A table: first row is a repeating header row (w:tblHeader); `align` gives per-column paragraph alignment. */
function table(rows, { style, columnWidths, align = [] }) {
  const grid = columnWidths.map(width => `<w:gridCol w:w="${width}"/>`).join('');
  const body = rows
    .map((cells, rowIndex) => {
      const trPr = rowIndex === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
      const tcs = cells
        .map((text, c) => {
          const tcPr = `<w:tcPr><w:tcW w:w="${columnWidths[c]}" w:type="dxa"/></w:tcPr>`;
          return `<w:tc>${tcPr}${paragraph(run(text), { jc: rowIndex > 0 ? align[c] : undefined })}</w:tc>`;
        })
        .join('');
      return `<w:tr>${trPr}${tcs}</w:tr>`;
    })
    .join('');
  const tblPr =
    `<w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="0" w:type="auto"/>` +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
  return `<w:tbl>${tblPr}<w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

// ===== DOCX: PARTS =====

const DOCX_REL = { styles: 'rId1', numbering: 'rId2', settings: 'rId3', fontTable: 'rId4', footnotes: 'rId5', image: 'rId6', link: 'rId7' };

function docxContentTypes() {
  return (
    XML_DECLARATION +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'
  );
}

function docxPackageRels() {
  return (
    XML_DECLARATION +
    `<Relationships xmlns="${NS.rel}">` +
    `<Relationship Id="rId1" Type="${NS.relType}/officeDocument" Target="word/document.xml"/>` +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    `<Relationship Id="rId3" Type="${NS.relType}/extended-properties" Target="docProps/app.xml"/>` +
    '</Relationships>'
  );
}

function docxCoreProps() {
  const stamp = FIXED_DATE.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return (
    XML_DECLARATION +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>FileStudio Sample Document</dc:title>' +
    '<dc:subject>Hand-built OOXML used to test the .docx preview</dc:subject>' +
    `<dc:creator>${GENERATOR_NAME}</dc:creator>` +
    '<cp:keywords>file-viewer, sample, docx</cp:keywords>' +
    `<cp:lastModifiedBy>${GENERATOR_NAME}</cp:lastModifiedBy>` +
    '<cp:revision>1</cp:revision>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '</cp:coreProperties>'
  );
}

function docxAppProps() {
  return (
    XML_DECLARATION +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Template>Normal.dotm</Template><TotalTime>0</TotalTime>' +
    `<Application>${GENERATOR_NAME}</Application>` +
    '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate>' +
    '<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
    '</Properties>'
  );
}

function docxDocumentRels() {
  const rel = (id, type, target, external) =>
    `<Relationship Id="${id}" Type="${NS.relType}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
  return (
    XML_DECLARATION +
    `<Relationships xmlns="${NS.rel}">` +
    rel(DOCX_REL.styles, 'styles', 'styles.xml') +
    rel(DOCX_REL.numbering, 'numbering', 'numbering.xml') +
    rel(DOCX_REL.settings, 'settings', 'settings.xml') +
    rel(DOCX_REL.fontTable, 'fontTable', 'fontTable.xml') +
    rel(DOCX_REL.footnotes, 'footnotes', 'footnotes.xml') +
    rel(DOCX_REL.image, 'image', 'media/image1.png') +
    rel(DOCX_REL.link, 'hyperlink', 'https://code.visualstudio.com/', true) +
    '</Relationships>'
  );
}

/** Paragraph/character/table styles with Word's built-in names (e.g. "heading 1", "caption"). */
function docxStyles() {
  const headingFont = '<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>';
  const border = side => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/>`;
  const heading = (level, size, color, before) =>
    `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/>` +
    `<w:next w:val="Normal"/><w:uiPriority w:val="9"/>${level > 1 ? '<w:unhideWhenUsed/>' : ''}<w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="80"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr>${headingFont}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
  return (
    XML_DECLARATION +
    `<w:styles xmlns:w="${NS.w}">` +
    '<w:docDefaults>' +
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>' +
    '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>' +
    '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>' +
    '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/>' +
    '<w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
    '<w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' +
    `<w:rPr>${headingFont}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
    '<w:uiPriority w:val="11"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr>' +
    '<w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
    heading(1, 32, '2F5496', 360) +
    heading(2, 26, '2F5496', 200) +
    heading(3, 24, '1F3763', 160) +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
    '<w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/>' +
    '<w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Normal"/>' +
    '<w:next w:val="Normal"/><w:uiPriority w:val="30"/><w:qFormat/><w:pPr>' +
    '<w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4"/></w:pBdr>' +
    '<w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>' +
    '<w:rPr><w:i/><w:iCs/><w:color w:val="4472C4"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
    '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
    '<w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="SourceCode"><w:name w:val="Source Code"/><w:basedOn w:val="Normal"/>' +
    '<w:uiPriority w:val="99"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' +
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>' +
    '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>' +
    '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' +
    '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>' +
    '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:basedOn w:val="DefaultParagraphFont"/>' +
    '<w:uiPriority w:val="22"/><w:qFormat/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="VerbatimChar"><w:name w:val="Verbatim Char"/><w:basedOn w:val="DefaultParagraphFont"/>' +
    '<w:uiPriority w:val="99"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/>' +
    '<w:szCs w:val="20"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:styleId="SampleTable"><w:name w:val="Sample Table"/><w:basedOn w:val="TableNormal"/>' +
    '<w:uiPriority w:val="49"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    `<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr>` +
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr>' +
    '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>' +
    '</w:style>' +
    '</w:styles>'
  );
}

/** numId 1 = bullets (abstract 0), numId 2 = decimal / lowerLetter / lowerRoman (abstract 1). */
function docxNumbering() {
  const level = (ilvl, numFmt, text, font) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${text}"/>` +
    `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr>` +
    (font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : '') +
    '</w:lvl>';
  return (
    XML_DECLARATION +
    `<w:numbering xmlns:w="${NS.w}">` +
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
    level(0, 'bullet', '•', 'Calibri') +
    level(1, 'bullet', '◦', 'Calibri') +
    level(2, 'bullet', '▪', 'Calibri') +
    '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' +
    level(0, 'decimal', '%1.') +
    level(1, 'lowerLetter', '%2.') +
    level(2, 'lowerRoman', '%3.') +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>'
  );
}

function docxSettings() {
  return (
    XML_DECLARATION +
    `<w:settings xmlns:w="${NS.w}">` +
    '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>' +
    '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
    '</w:settings>'
  );
}

function docxFontTable() {
  const fontEntry = (name, panose, family, pitch) =>
    `<w:font w:name="${name}"><w:panose1 w:val="${panose}"/><w:charset w:val="00"/><w:family w:val="${family}"/><w:pitch w:val="${pitch}"/></w:font>`;
  return (
    XML_DECLARATION +
    `<w:fonts xmlns:w="${NS.w}">` +
    fontEntry('Calibri', '020F0502020204030204', 'swiss', 'variable') +
    fontEntry('Calibri Light', '020F0302020204030204', 'swiss', 'variable') +
    fontEntry('Times New Roman', '02020603050405020304', 'roman', 'variable') +
    fontEntry('Consolas', '020B0609020204030204', 'modern', 'fixed') +
    '</w:fonts>'
  );
}

function docxFootnotes() {
  const separator = (type, id, element) =>
    `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r>${element}</w:r></w:p></w:footnote>`;
  return (
    XML_DECLARATION +
    `<w:footnotes xmlns:w="${NS.w}" xmlns:r="${NS.r}">` +
    separator('separator', -1, '<w:separator/>') +
    separator('continuationSeparator', 0, '<w:continuationSeparator/>') +
    '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
    '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
    run(' Footnotes are optional in the spec; mammoth renders them as a numbered list at the end.') +
    '</w:p></w:footnote>' +
    '</w:footnotes>'
  );
}

/** The document body: title, headings, inline formatting, links, lists, table, quotes, code, image, footnote. */
function docxDocument(image) {
  const BULLETS = 1;
  const NUMBERS = 2;
  const footnoteRef = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
  const body = [
    paragraph(run('FileStudio Sample Document'), { style: 'Title' }),
    paragraph(run('A hand-built OOXML file for testing the .docx preview'), { style: 'Subtitle' }),

    paragraph(run('Introduction'), { style: 'Heading1' }),
    paragraph([
      run('This document was generated by '),
      run('test/generate-samples.js', { style: 'VerbatimChar' }),
      run(' without Word. It exercises '),
      run('bold', { b: true }),
      run(', '),
      run('italic', { i: true }),
      run(', '),
      run('underline', { u: true }),
      run(', '),
      run('bold italic', { b: true, i: true }),
      run(', '),
      run('strikethrough', { strike: true }),
      run(', '),
      run('colored text', { color: 'C00000' }),
      run(', '),
      run('highlighted text', { highlight: 'yellow' }),
      run(', a '),
      run('Strong', { style: 'Strong' }),
      run(' character style, H'),
      run('2', { vertAlign: 'subscript' }),
      run('O and E = mc'),
      run('2', { vertAlign: 'superscript' }),
      run('.'),
    ]),
    paragraph([
      run('Visit the '),
      hyperlink(run('Visual Studio Code website', { style: 'Hyperlink' }), { relId: DOCX_REL.link }),
      run(' or jump to the '),
      hyperlink(run('table section', { style: 'Hyperlink' }), { anchor: 'table_section' }),
      run('. This sentence carries a footnote.'),
      footnoteRef,
    ]),
    paragraph(run('Unicode and special characters'), { style: 'Heading2' }),
    paragraph(run('Accents: café, naïve, Straße · CJK: 日本語, 中文 · Greek: Ελληνικά · Hebrew: עברית · Emoji: 🎉 ✔ · XML specials: <tag> & "quotes" \'apostrophes\'.')),
    paragraph(run('A line with a tab\there, and a manual line break\nright before this text.')),

    paragraph(run('Lists'), { style: 'Heading1' }),
    paragraph(run('Bulleted list'), { style: 'Heading2' }),
    paragraph(run('First bullet'), { style: 'ListParagraph', numId: BULLETS, ilvl: 0 }),
    paragraph([run('Second bullet with '), run('bold', { b: true }), run(' text')], { style: 'ListParagraph', numId: BULLETS, ilvl: 0 }),
    paragraph(run('Nested bullet (level 2)'), { style: 'ListParagraph', numId: BULLETS, ilvl: 1 }),
    paragraph(run('Deeper bullet (level 3)'), { style: 'ListParagraph', numId: BULLETS, ilvl: 2 }),
    paragraph(run('Back to level 2'), { style: 'ListParagraph', numId: BULLETS, ilvl: 1 }),
    paragraph(run('Third bullet'), { style: 'ListParagraph', numId: BULLETS, ilvl: 0 }),
    paragraph(run('Numbered list'), { style: 'Heading2' }),
    paragraph(run('Step one'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 0 }),
    paragraph(run('Step two'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 0 }),
    paragraph(run('Sub-step a'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 1 }),
    paragraph(run('Sub-step b'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 1 }),
    paragraph(run('Sub-sub-step i'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 2 }),
    paragraph(run('Step three'), { style: 'ListParagraph', numId: NUMBERS, ilvl: 0 }),

    paragraph(['<w:bookmarkStart w:id="0" w:name="table_section"/>', run('Table'), '<w:bookmarkEnd w:id="0"/>'], { style: 'Heading1' }),
    paragraph(run('Quarterly figures (3 columns × 4 rows, the first row is a repeating header row):')),
    table(
      [
        ['Quarter', 'Revenue', 'Growth'],
        ['Q1 2024', '$1,200', '+4.0%'],
        ['Q2 2024', '$1,450', '+20.8%'],
        ['Q3 2024', '$1,380', '−4.8%'],
      ],
      { style: 'SampleTable', columnWidths: [3120, 3120, 3120], align: ['left', 'right', 'right'] },
    ),
    paragraph(run('Quotes'), { style: 'Heading2' }),
    paragraph(run('“Simplicity is prerequisite for reliability.” — Edsger W. Dijkstra'), { style: 'Quote' }),
    paragraph(run('Intense quote: measure twice, cut once.'), { style: 'IntenseQuote' }),
    paragraph(run('Code'), { style: 'Heading3' }),
    paragraph(run('const viewer = await openWith(uri);'), { style: 'SourceCode' }),
    paragraph(run('viewer.render({ theme: "dark" });'), { style: 'SourceCode' }),

    paragraph(run('Image'), { style: 'Heading1' }),
    paragraph(inlineImage({ relId: DOCX_REL.image, id: 1, name: 'image1.png', description: 'Bar chart generated as a PNG', widthPx: image.width, heightPx: image.height }), { jc: 'center', spacingAfter: 0 }),
    paragraph(run('Figure 1: a PNG generated by the sample script and embedded inline.'), { style: 'Caption', jc: 'center' }),
    paragraph(run('End of document.')),
  ];
  const sectPr =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>';
  return (
    XML_DECLARATION +
    `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}">` +
    `<w:body>${body.join('')}${sectPr}</w:body></w:document>`
  );
}

/** Assembles sample.docx with fixed timestamps. */
async function generateDocx(images) {
  const zip = new JSZip();
  const add = (name, content) => zip.file(name, content, { date: FIXED_DATE, createFolders: false });
  add('[Content_Types].xml', docxContentTypes());
  add('_rels/.rels', docxPackageRels());
  add('docProps/core.xml', docxCoreProps());
  add('docProps/app.xml', docxAppProps());
  add('word/document.xml', docxDocument({ width: 240, height: 144 }));
  add('word/_rels/document.xml.rels', docxDocumentRels());
  add('word/styles.xml', docxStyles());
  add('word/numbering.xml', docxNumbering());
  add('word/settings.xml', docxSettings());
  add('word/fontTable.xml', docxFontTable());
  add('word/footnotes.xml', docxFootnotes());
  add('word/media/image1.png', images.chart);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, platform: 'DOS' });
}

// ===== PDF: WRITER =====

/** Unicode characters of WinAnsiEncoding's 0x80-0x9F block (0x20-0x7E and 0xA0-0xFF are the same as Latin-1). */
const WIN_ANSI_HIGH = new Map([
  ['€', 0x80], ['‚', 0x82], ['ƒ', 0x83], ['„', 0x84], ['…', 0x85], ['†', 0x86], ['‡', 0x87], ['ˆ', 0x88],
  ['‰', 0x89], ['Š', 0x8a], ['‹', 0x8b], ['Œ', 0x8c], ['Ž', 0x8e], ['‘', 0x91], ['’', 0x92], ['“', 0x93],
  ['”', 0x94], ['•', 0x95], ['–', 0x96], ['—', 0x97], ['˜', 0x98], ['™', 0x99], ['š', 0x9a], ['›', 0x9b],
  ['œ', 0x9c], ['ž', 0x9e], ['Ÿ', 0x9f],
]);

/** Text -> WinAnsiEncoding bytes as a binary ('latin1') string. Throws on characters the encoding does not have. */
function winAnsi(text) {
  let bytes = '';
  for (const ch of text) {
    const code = ch.codePointAt(0);
    const byte = WIN_ANSI_HIGH.get(ch) ?? ((code >= 0x20 && code <= 0x7e) || (code >= 0xa0 && code <= 0xff) ? code : undefined);
    if (byte === undefined) {
      throw new Error(`PDF: "${ch}" (U+${code.toString(16).toUpperCase().padStart(4, '0')}) is not in WinAnsiEncoding`);
    }
    bytes += String.fromCharCode(byte);
  }
  return bytes;
}

/** Literal string for a content stream: WinAnsi bytes with \ ( ) escaped. */
function pdfLiteral(text) {
  return `(${winAnsi(text).replace(/[\\()]/g, '\\$&')})`;
}

/** Text string outside content streams (outline titles, document info): an ASCII literal, else UTF-16BE with a BOM. */
function pdfTextString(text) {
  if (/^[\x20-\x7e]*$/.test(text)) return `(${text.replace(/[\\()]/g, '\\$&')})`;
  return `<FEFF${Buffer.from(text, 'utf16le').swap16().toString('hex').toUpperCase()}>`;
}

/** Number for content streams and dictionaries: at most 3 decimals, no trailing zeros, never "-0". */
function pdfNumber(value) {
  const rounded = Math.round(value * 1000) / 1000;
  return String(rounded === 0 ? 0 : rounded);
}

/** 'RRGGBB' -> "r g b" (0..1) for the rg / RG operators and /C arrays. */
function pdfColor(hex) {
  return hexToRgb(hex).map(value => pdfNumber(value / 255)).join(' ');
}

/** PDF date string of FIXED_DATE: D:YYYYMMDDHHmmSSZ. */
function pdfDate(date) {
  return `D:${date.toISOString().replace(/[-:T]/g, '').slice(0, 14)}Z`;
}

/** Raw 8-bit RGB samples (no alpha, no row filter) for an image XObject. */
function rawRgb(width, height, pixel) {
  const data = Buffer.alloc(width * height * 3);
  let offset = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      for (const value of pixel(x, y).slice(0, 3)) data[offset++] = Math.max(0, Math.min(255, Math.round(value)));
    }
  }
  return data;
}

/**
 * Minimal PDF file writer: numbered indirect objects, a classic cross-reference table and a trailer.
 * Object bodies are binary ('latin1') strings or Buffers; reserve() hands out numbers for forward references.
 */
class PdfWriter {
  constructor() {
    this.objects = [];
  }

  reserve() {
    this.objects.push(null);
    return this.objects.length;
  }

  set(id, body) {
    this.objects[id - 1] = Buffer.isBuffer(body) ? body : Buffer.from(body, 'latin1');
    return id;
  }

  add(body) {
    return this.set(this.reserve(), body);
  }

  /** Adds a stream object (`entries` = extra dictionary entries), FlateDecode-compressed. */
  addStream(entries, data) {
    const payload = zlib.deflateSync(data, { level: 9 });
    const head = `<< ${entries ? `${entries} ` : ''}/Filter /FlateDecode /Length ${payload.length} >>\nstream\n`;
    return this.add(Buffer.concat([Buffer.from(head, 'latin1'), payload, Buffer.from('\nendstream', 'latin1')]));
  }

  /** Header (with the binary-marker comment), objects, xref table and trailer. */
  toBuffer({ root, info }) {
    const header = Buffer.from('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n', 'latin1');
    const chunks = [header];
    const offsets = [];
    let offset = header.length;
    this.objects.forEach((body, index) => {
      if (!body) throw new Error(`PDF object ${index + 1} was reserved but never written`);
      const chunk = Buffer.concat([Buffer.from(`${index + 1} 0 obj\n`, 'latin1'), body, Buffer.from('\nendobj\n', 'latin1')]);
      offsets.push(offset);
      chunks.push(chunk);
      offset += chunk.length;
    });
    // File identifier: a digest of the objects, stable between runs and different for different content.
    const id = crypto.createHash('md5').update(Buffer.concat(chunks)).digest('hex').toUpperCase();
    const size = this.objects.length + 1;
    // Every xref entry is exactly 20 bytes: 10-digit offset, 5-digit generation, n/f, CR LF.
    const xref = `xref\n0 ${size}\n0000000000 65535 f\r\n${offsets.map(o => `${String(o).padStart(10, '0')} 00000 n\r\n`).join('')}`;
    const trailer = `trailer\n<< /Size ${size} /Root ${root} 0 R /Info ${info} 0 R /ID [<${id}> <${id}>] >>\nstartxref\n${offset}\n%%EOF\n`;
    chunks.push(Buffer.from(xref + trailer, 'latin1'));
    return Buffer.concat(chunks);
  }
}

// ===== PDF: TEXT LAYOUT =====

/** Advance widths (1/1000 em) of Helvetica and Helvetica-Bold for ' '..'~' (Adobe AFM metrics, WinAnsi codes). */
const HELVETICA_ASCII =
  '278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 ' +
  '584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 ' +
  '667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 ' +
  '278 556 500 722 500 500 500 334 260 334 584';
const HELVETICA_BOLD_ASCII =
  '278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 333 333 ' +
  '584 584 584 611 975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 667 778 722 667 611 722 667 944 ' +
  '667 667 611 333 278 333 584 556 333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 611 611 389 556 ' +
  '333 611 556 778 556 556 500 389 280 389 584';
/** Widths of the non-ASCII characters the sample uses (character immediately followed by its width). */
const HELVETICA_EXTRA =
  'é556 ï278 ß611 Æ1000 ø611 ã556 ©737 ®737 °400 €556 £556 ±584 ×584 ½834 •350 “333 ”333 ‘222 ’222 –556 —1000 ' +
  '…1000 ·278 ü556 ö556 ä556 É667 ç500 è556 á556 ñ556 å556 ó556 í278 ú556 Ø778 µ556 §556 Ñ722 Å667';
const HELVETICA_BOLD_EXTRA =
  'é556 ï278 ß611 Æ1000 ø611 ã556 ©737 ®737 °400 €556 £556 ±584 ×584 ½834 •350 “500 ”500 ‘278 ’278 –556 —1000 ' +
  '…1000 ·278 ü611 ö611 ä556 É667 ç556 è556 á556 ñ611 å556 ó611 í278 ú611 Ø778 µ611 §556 Ñ722 Å722';

function widthTable(ascii, extra) {
  const table = new Map();
  ascii.split(' ').forEach((width, i) => table.set(String.fromCharCode(32 + i), Number(width)));
  for (const entry of extra.split(' ')) table.set(entry[0], Number(entry.slice(1)));
  return table;
}

const HELVETICA_WIDTHS = widthTable(HELVETICA_ASCII, HELVETICA_EXTRA);

/** The four standard (non-embedded) fonts of the sample; Courier is monospaced (600 units per glyph). */
const PDF_FONTS = {
  regular: { name: 'F1', baseFont: 'Helvetica', widths: HELVETICA_WIDTHS },
  bold: { name: 'F2', baseFont: 'Helvetica-Bold', widths: widthTable(HELVETICA_BOLD_ASCII, HELVETICA_BOLD_EXTRA) },
  italic: { name: 'F3', baseFont: 'Helvetica-Oblique', widths: HELVETICA_WIDTHS },
  mono: { name: 'F4', baseFont: 'Courier', widths: null },
};

/** Width in points of `text` set in `font` (a PDF_FONTS key) at `size`. */
function textWidth(text, font, size) {
  const { widths, baseFont } = PDF_FONTS[font];
  let units = 0;
  for (const ch of text) {
    const width = widths ? widths.get(ch) : 600;
    if (width === undefined) throw new Error(`PDF: no ${baseFont} width for "${ch}" — add it to the width table`);
    units += width;
  }
  return (units * size) / 1000;
}

/** Text styles (points, RRGGBB). `before` / `after` are paragraph spacing. */
const PDF_TEXT = {
  title: { font: 'bold', size: 26, leading: 32, color: '1F3864', after: 2 },
  subtitle: { font: 'italic', size: 13, leading: 18, color: '595959', after: 8 },
  h1: { font: 'bold', size: 18, leading: 22, color: '2F5496', before: 14, after: 6 },
  h2: { font: 'bold', size: 13, leading: 17, color: '1F3864', before: 8, after: 4 },
  body: { font: 'regular', size: 11, leading: 15, color: '262626', after: 8 },
  caption: { font: 'italic', size: 9, leading: 12, color: '595959', after: 10 },
  code: { font: 'mono', size: 9.5, leading: 13, color: '1F1F1F', after: 10 },
  small: { font: 'regular', size: 8, leading: 10, color: '404040', after: 0 },
  footer: { font: 'regular', size: 8.5, leading: 10, color: '7F7F7F', after: 0 },
};
const PDF_LINK_COLOR = '0563C1';
const PDF_MARGIN = 72;

/**
 * Greedy line breaking over styled runs ({ text, font, color, link }). Lines only break at spaces (a word may span
 * runs, e.g. a link followed by a period); spaces at a break are dropped. Returns lines of segments, where a
 * segment is the consecutive text of one run.
 */
function wrapRuns(runs, size, maxWidth) {
  // Tokens: maximal space / non-space pieces of each run.
  const tokens = [];
  runs.forEach((run, runIndex) => {
    for (const [text] of run.text.matchAll(/\s+|\S+/g)) {
      tokens.push({ run, runIndex, text, space: /^\s/.test(text), width: textWidth(text, run.font, size) });
    }
  });
  // Units: words (runs of non-space tokens) and the spaces between them.
  const units = [];
  for (const token of tokens) {
    const last = units[units.length - 1];
    if (last && last.space === token.space) {
      last.tokens.push(token);
      last.width += token.width;
    } else {
      units.push({ space: token.space, tokens: [token], width: token.width });
    }
  }
  const lines = [];
  let line = { tokens: [], width: 0 };
  let pending = null; // spaces waiting for the next word on the same line
  for (const unit of units) {
    if (unit.space) {
      if (line.tokens.length) pending = unit;
      continue;
    }
    const spaceWidth = pending ? pending.width : 0;
    if (line.tokens.length && line.width + spaceWidth + unit.width > maxWidth) {
      lines.push(line);
      line = { tokens: [], width: 0 };
    } else if (pending) {
      line.tokens.push(...pending.tokens);
      line.width += spaceWidth;
    }
    pending = null;
    line.tokens.push(...unit.tokens);
    line.width += unit.width;
  }
  if (line.tokens.length) lines.push(line);
  return lines.map(({ tokens: lineTokens, width }) => {
    const segments = [];
    for (const token of lineTokens) {
      const last = segments[segments.length - 1];
      if (last && last.runIndex === token.runIndex) last.text += token.text;
      else segments.push({ runIndex: token.runIndex, text: token.text, font: token.run.font, color: token.run.color, link: token.run.link });
    }
    return { segments, width };
  });
}

/**
 * One page laid out top-down: `y` is the distance from the top edge, content operators use PDF's bottom-up
 * coordinates. Collects the content operators, link annotations and anchors (PDF y of headings, for the
 * outline and internal links).
 */
class PdfPage {
  constructor(width, height) {
    this.width = width;
    this.height = height;
    this.y = PDF_MARGIN;
    this.ops = [];
    this.links = []; // { rect: [x1, y1, x2, y2], target: { uri } | { page, fit } | { name } }
    this.anchors = {};
  }

  get left() {
    return PDF_MARGIN;
  }

  get contentWidth() {
    return this.width - 2 * PDF_MARGIN;
  }

  /** PDF y coordinate of a point `top` points below the top edge. */
  pdfY(top) {
    return this.height - top;
  }

  /** Baseline (from the top) of a line of `style` that starts at the current position. */
  baseline(style) {
    return this.y + (style.leading - style.size) / 2 + style.size * 0.8;
  }

  drawText(text, x, baseline, { font, size, color }) {
    this.ops.push(
      `BT /${PDF_FONTS[font].name} ${pdfNumber(size)} Tf ${pdfColor(color)} rg ` +
        `${pdfNumber(x)} ${pdfNumber(this.pdfY(baseline))} Td ${pdfLiteral(text)} Tj ET`,
    );
  }

  /** Filled rectangle; x / top / width / height in top-down page units. */
  fillRect(x, top, width, height, color) {
    this.ops.push(`${pdfColor(color)} rg ${pdfNumber(x)} ${pdfNumber(this.pdfY(top + height))} ${pdfNumber(width)} ${pdfNumber(height)} re f`);
  }

  /** Straight line between two top-down points. */
  line(x1, top1, x2, top2, color, width) {
    this.ops.push(`${pdfColor(color)} RG ${pdfNumber(width)} w ${pdfNumber(x1)} ${pdfNumber(this.pdfY(top1))} m ${pdfNumber(x2)} ${pdfNumber(this.pdfY(top2))} l S`);
  }

  /** Remembers the current position under `name` (a little headroom above the line). */
  anchor(name) {
    this.anchors[name] = pdfNumber(this.pdfY(this.y) + 4);
  }

  /**
   * Wraps and draws a paragraph. `runs` is a string or a list of strings / { text, font?, color?, link? };
   * links are blue and underlined and get one link annotation per line they occupy.
   */
  paragraph(runs, style, { indent = 0, after = style.after, align = 'left' } = {}) {
    const list = (Array.isArray(runs) ? runs : [runs]).map(run => {
      const spec = typeof run === 'string' ? { text: run } : run;
      return {
        text: spec.text,
        font: spec.font || style.font,
        color: spec.link ? PDF_LINK_COLOR : spec.color || style.color,
        link: spec.link,
      };
    });
    const maxWidth = this.contentWidth - indent;
    for (const { segments, width } of wrapRuns(list, style.size, maxWidth)) {
      const baseline = this.baseline(style);
      let x = this.left + indent + (align === 'center' ? (maxWidth - width) / 2 : align === 'right' ? maxWidth - width : 0);
      for (const segment of segments) {
        const segmentWidth = textWidth(segment.text, segment.font, style.size);
        this.drawText(segment.text, x, baseline, { font: segment.font, size: style.size, color: segment.color });
        if (segment.link) {
          this.fillRect(x, baseline + style.size * 0.1, segmentWidth, style.size * 0.06, segment.color); // underline
          this.links.push({
            rect: [x, this.pdfY(baseline + style.size * 0.25), x + segmentWidth, this.pdfY(baseline - style.size * 0.85)],
            target: segment.link,
          });
        }
        x += segmentWidth;
      }
      this.y += style.leading;
    }
    this.y += after;
  }

  /** Heading with space before it (except at the top of the page); records an anchor for the outline. */
  heading(text, style, anchorName) {
    if (this.y > PDF_MARGIN) this.y += style.before;
    if (anchorName) this.anchor(anchorName);
    this.paragraph(text, style);
  }

  /** Bulleted ('•') or numbered list; wrapped lines align with the item text. */
  list(items, style, { numbered = false } = {}) {
    items.forEach((runs, i) => {
      this.drawText(numbered ? `${i + 1}.` : '•', this.left + 12, this.baseline(style), style);
      this.paragraph(runs, style, { indent: 28, after: 3 });
    });
    this.y += style.after - 3;
  }

  /** Monospaced lines on a light grey panel with a darker left edge. */
  codeBlock(lines, style) {
    const padding = 6;
    const height = lines.length * style.leading + 2 * padding;
    this.fillRect(this.left, this.y, this.contentWidth, height, 'F2F2F2');
    this.fillRect(this.left, this.y, 2.5, height, 'A6A6A6');
    this.y += padding;
    for (const line of lines) {
      if (textWidth(line, style.font, style.size) > this.contentWidth - 20) throw new Error(`PDF: code line too long: ${line}`);
      this.drawText(line, this.left + 12, this.baseline(style), style);
      this.y += style.leading;
    }
    this.y += padding + style.after;
  }

  /** Image XObject `name` centred in the content area. */
  image(name, width, height) {
    const x = this.left + (this.contentWidth - width) / 2;
    this.ops.push(`q ${pdfNumber(width)} 0 0 ${pdfNumber(height)} ${pdfNumber(x)} ${pdfNumber(this.pdfY(this.y + height))} cm /${name} Do Q`);
    this.y += height;
  }

  /** Centred page footer 36 pt above the bottom edge. */
  footer(text) {
    const style = PDF_TEXT.footer;
    const width = textWidth(text, style.font, style.size);
    this.drawText(text, (this.width - width) / 2, this.height - 36, style);
  }

  /** Throws when the flowing content ran into the bottom margin (the sample must stay hand-checked, not paginated). */
  checkFits(label) {
    if (this.y > this.height - PDF_MARGIN + 12) throw new Error(`PDF: ${label} overflows the page (y = ${this.y.toFixed(1)})`);
  }
}

// ===== PDF: DOCUMENT =====

const PDF_IMAGE = { name: 'Im1', width: 192, height: 120 };
const PDF_PAGE_COUNT = 4;

function pdfFooter(page, number) {
  page.footer(`FileStudio sample PDF  ·  Page ${number} of ${PDF_PAGE_COUNT}`);
}

/** Page 1 (portrait): title, introduction, external / e-mail / internal links, special characters, code. */
function layoutPdfPage1() {
  const page = new PdfPage(612, 792);
  page.anchor('top');
  page.paragraph('FileStudio Sample PDF', PDF_TEXT.title);
  page.paragraph('A hand-written PDF 1.7 file for testing the PDF preview', PDF_TEXT.subtitle);
  page.fillRect(page.left, page.y, page.contentWidth, 1.5, '4472C4');
  page.y += 6;

  page.heading('1  Introduction', PDF_TEXT.h1, 'intro');
  page.paragraph(
    'This document was written byte by byte by test/generate-samples.js, without a PDF library. It has four pages, ' +
      'a document outline (bookmarks), external and internal links, an embedded image, vector graphics and one ' +
      'landscape page. All text uses the standard fonts Helvetica and Courier, so it can be selected, copied and searched.',
    PDF_TEXT.body,
  );
  page.paragraph(
    'The file has a classic cross-reference table, compressed content streams (FlateDecode) and a document ' +
      'information dictionary with a title, so a viewer can show “FileStudio Sample PDF” as the document name.',
    PDF_TEXT.body,
  );

  page.heading('1.1  Links', PDF_TEXT.h2, 'links');
  page.paragraph(
    [
      'External link: ',
      { text: 'https://example.com', link: { uri: 'https://example.com/' } },
      '. E-mail link: ',
      { text: 'someone@example.com', link: { uri: 'mailto:someone@example.com' } },
      '. Internal link: ',
      { text: 'jump to page 3 (the landscape page)', link: { page: 2, fit: true } },
      '. Each link is a separate link annotation over the underlined text.',
    ],
    PDF_TEXT.body,
  );

  page.heading('1.2  Special characters', PDF_TEXT.h2, 'special');
  page.paragraph(
    'Accents: café, naïve, Straße, Ærøskøbing, São Paulo, Ñandú, Århus. Symbols: © ® ° € £ ± × ½ µ §. ' +
      'Typography: “double” and ‘single’ quotes, en – and em — dashes, a bullet • and an ellipsis… ' +
      'Characters that PDF strings escape: (parentheses) and a back\\slash.',
    PDF_TEXT.body,
  );

  page.heading('1.3  Code', PDF_TEXT.h2, 'code');
  page.paragraph('Courier text on a shaded panel (each line is one text object):', PDF_TEXT.body, { after: 4 });
  page.codeBlock(
    [
      "const pdf = await pdfjsLib.getDocument({ url: 'sample.pdf' }).promise;",
      'const page = await pdf.getPage(3);',
      'console.log(page.view); // [0, 0, 792, 612] -> landscape',
      'const text = await page.getTextContent();',
    ],
    PDF_TEXT.code,
  );
  page.checkFits('page 1');
  pdfFooter(page, 1);
  return page;
}

/** Page 2 (portrait): embedded RGB image and vector graphics (rectangles, circles, transparency, a bar chart). */
function layoutPdfPage2() {
  const page = new PdfPage(612, 792);
  page.heading('2  Image', PDF_TEXT.h1, 'image');
  page.paragraph(
    `The picture below is a ${PDF_IMAGE.width} × ${PDF_IMAGE.height} pixel image XObject: 8-bit DeviceRGB samples ` +
      'compressed with FlateDecode, no alpha channel. It is drawn at 288 × 180 points, so it is scaled up by 1.5.',
    PDF_TEXT.body,
  );
  page.y += 4;
  page.image(PDF_IMAGE.name, 288, 180);
  page.y += 6;
  page.paragraph(
    'Figure 1: hue/lightness gradient with a dark frame; the white square marks the top-left corner.',
    PDF_TEXT.caption,
    { align: 'center' },
  );

  page.heading('2.1  Vector graphics', PDF_TEXT.h2, 'vector');
  page.paragraph(
    'Path operators: a filled rectangle, a dashed outline, a circle built from four Bézier curves, a second ' +
      'circle at 60 % opacity (ExtGState), a triangle, and a small bar chart with a polyline on top.',
    PDF_TEXT.body,
  );
  drawPdfShapes(page);
  page.checkFits('page 2');
  pdfFooter(page, 2);
  return page;
}

/** Circle as four cubic Bézier curves around (cx, cy) in PDF coordinates. */
function pdfCircle(cx, cy, r) {
  const k = 0.5523 * r;
  const n = pdfNumber;
  return (
    `${n(cx + r)} ${n(cy)} m ` +
    `${n(cx + r)} ${n(cy + k)} ${n(cx + k)} ${n(cy + r)} ${n(cx)} ${n(cy + r)} c ` +
    `${n(cx - k)} ${n(cy + r)} ${n(cx - r)} ${n(cy + k)} ${n(cx - r)} ${n(cy)} c ` +
    `${n(cx - r)} ${n(cy - k)} ${n(cx - k)} ${n(cy - r)} ${n(cx)} ${n(cy - r)} c ` +
    `${n(cx + k)} ${n(cy - r)} ${n(cx + r)} ${n(cy - k)} ${n(cx + r)} ${n(cy)} c`
  );
}

/** The vector-graphics showcase of page 2 (shapes row, then a bar chart with value labels and a trend line). */
function drawPdfShapes(page) {
  const n = pdfNumber;
  const top = page.y + 6;
  const y = value => n(page.pdfY(value));
  const ops = page.ops;
  ops.push('q');
  page.fillRect(72, top, 90, 60, '4472C4');
  ops.push(`${pdfColor('C00000')} RG 1.5 w [4 3] 0 d 182 ${y(top + 60)} 90 60 re S [] 0 d`);
  ops.push(`${pdfColor('ED7D31')} rg ${pdfColor('843C0C')} RG 1 w ${pdfCircle(325, page.pdfY(top + 30), 30)} B`);
  ops.push(`q /GS1 gs ${pdfColor('5B9BD5')} rg ${pdfCircle(355, page.pdfY(top + 42), 26)} f Q`);
  ops.push(`${pdfColor('70AD47')} rg 405 ${y(top + 60)} m 445 ${y(top)} l 485 ${y(top + 60)} l h f`);
  ops.push(`${pdfColor('7F7F7F')} RG 3 w 1 J 505 ${y(top + 55)} m 540 ${y(top + 5)} l S 0 J`);

  // Bar chart: axis, bars in the theme accent colors, value labels above, month labels below, trend polyline.
  const values = [42, 65, 37, 74, 58, 50];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  const chartTop = top + 80;
  const baseline = chartTop + 90;
  const left = 96;
  const step = 52;
  const barWidth = 32;
  page.drawText('Monthly totals', 72, chartTop + 2, { font: 'bold', size: 9, color: '262626' });
  page.line(left - 8, chartTop + 8, left - 8, baseline, '595959', 0.75);
  page.line(left - 8, baseline, left + step * values.length, baseline, '595959', 0.75);
  const points = [];
  values.forEach((value, i) => {
    const x = left + i * step;
    page.fillRect(x, baseline - value, barWidth, value, ACCENTS[i]);
    const label = String(value);
    page.drawText(label, x + (barWidth - textWidth(label, 'regular', 8)) / 2, baseline - value - 3, PDF_TEXT.small);
    page.drawText(months[i], x + (barWidth - textWidth(months[i], 'regular', 8)) / 2, baseline + 10, PDF_TEXT.small);
    points.push([x + barWidth / 2, baseline - value - 14]);
  });
  ops.push(`${pdfColor('1F3864')} RG 1.25 w 1 j ${points.map(([px, py], i) => `${n(px)} ${y(py)} ${i ? 'l' : 'm'}`).join(' ')} S 0 j`);
  ops.push('Q');
  page.y = baseline + 22;
}

/** Regional figures for the landscape table (current quarters and the previous year's total). */
const PDF_TABLE_ROWS = [
  ['North America', [1250, 1310, 1405, 1520], 5155],
  ['Europe', [980, 1015, 990, 1102], 3964],
  ['Asia Pacific', [1430, 1505, 1610, 1755], 5738],
  ['Latin America', [410, 395, 430, 468], 1724],
  ['Middle East & Africa', [300, 322, 341, 360], 1264],
];

function groupedNumber(value) {
  return value.toLocaleString('en-US');
}

function percentChange(current, previous) {
  const change = ((current - previous) / previous) * 100;
  return `${change >= 0 ? '+' : '-'}${Math.abs(change).toFixed(1)} %`;
}

/** Page 3 (landscape 792 x 612): target of the internal link; a ruled table with right-aligned numbers. */
function layoutPdfPage3() {
  const page = new PdfPage(792, 612);
  page.heading('3  Landscape page', PDF_TEXT.h1, 'landscape');
  page.paragraph(
    'This page is 11 × 8.5 inches (792 × 612 points): its MediaBox is wider than it is tall. The internal link on ' +
      'page 1 and the outline entry “3  Landscape page” both lead here.',
    PDF_TEXT.body,
  );
  page.anchor('table');
  page.y += 4;
  const columns = [
    ['Region', 180, 'left'],
    ['Q1', 72, 'right'],
    ['Q2', 72, 'right'],
    ['Q3', 72, 'right'],
    ['Q4', 72, 'right'],
    ['Total', 90, 'right'],
    ['vs. last year', 90, 'right'],
  ];
  const totals = [0, 0, 0, 0];
  let total = 0;
  let previousTotal = 0;
  const rows = PDF_TABLE_ROWS.map(([region, quarters, previous]) => {
    const sum = quarters.reduce((a, b) => a + b, 0);
    quarters.forEach((value, i) => (totals[i] += value));
    total += sum;
    previousTotal += previous;
    return [region, ...quarters.map(groupedNumber), groupedNumber(sum), percentChange(sum, previous)];
  });
  rows.push(['Total', ...totals.map(groupedNumber), groupedNumber(total), percentChange(total, previousTotal)]);

  const rowHeight = 22;
  const padding = 8;
  const tableWidth = columns.reduce((sum, [, width]) => sum + width, 0);
  const drawRow = (cells, top, { font, color, fill }) => {
    if (fill) page.fillRect(page.left, top, tableWidth, rowHeight, fill);
    let x = page.left;
    cells.forEach((text, c) => {
      const [, width, align] = columns[c];
      const textX = align === 'right' ? x + width - padding - textWidth(text, font, 10) : x + padding;
      page.drawText(text, textX, top + 15, { font, size: 10, color });
      x += width;
    });
  };
  const tableTop = page.y;
  drawRow(columns.map(([title]) => title), tableTop, { font: 'bold', color: 'FFFFFF', fill: '4472C4' });
  rows.forEach((cells, r) => {
    const last = r === rows.length - 1;
    drawRow(cells, tableTop + (r + 1) * rowHeight, {
      font: last ? 'bold' : 'regular',
      color: '262626',
      fill: last ? 'D9E2F3' : r % 2 ? 'F2F2F2' : null,
    });
  });
  const tableBottom = tableTop + (rows.length + 1) * rowHeight;
  for (let r = 1; r <= rows.length; r++) {
    const lineTop = tableTop + r * rowHeight;
    page.line(page.left, lineTop, page.left + tableWidth, lineTop, r === rows.length ? '404040' : 'BFBFBF', r === rows.length ? 1 : 0.5);
  }
  let x = page.left;
  for (const [, width] of columns) {
    page.line(x, tableTop, x, tableBottom, 'BFBFBF', 0.5);
    x += width;
  }
  page.line(x, tableTop, x, tableBottom, 'BFBFBF', 0.5);
  page.line(page.left, tableBottom, page.left + tableWidth, tableBottom, '404040', 1);
  page.y = tableBottom + 10;
  page.paragraph(
    'Table 1: revenue by region and quarter (thousand €). Numbers are right-aligned; the header row has a filled ' +
      'background, every second row is shaded and the total row is bold.',
    PDF_TEXT.caption,
  );
  page.checkFits('page 3');
  pdfFooter(page, 3);
  return page;
}

/** Page 4 (portrait): find test, lists, a long paragraph and named-destination links back. */
function layoutPdfPage4() {
  const page = new PdfPage(612, 792);
  page.heading('4  Searching and selecting text', PDF_TEXT.h1, 'search');
  page.paragraph(
    'Every word in this file is real text drawn with standard fonts, so it can be selected with the mouse, copied ' +
      'with Ctrl+C and found with the search box.',
    PDF_TEXT.body,
  );
  page.paragraph(
    'Find test: the word on this line comes in three spellings — needle, Needle and NEEDLE. A search without ' +
      'Match case finds all three; with Match case on, a search for the lowercase word finds one.',
    PDF_TEXT.body,
  );

  page.heading('4.1  Lists', PDF_TEXT.h2, 'lists');
  page.list(
    [
      'Zoom with Ctrl+Plus and Ctrl+Minus, or fit the page width to the window.',
      'The outline panel lists the headings of all four pages; the entry “2  Image” starts collapsed.',
      'Page 3 uses landscape orientation (11 × 8.5 inches); the other pages are US Letter portrait.',
    ],
    PDF_TEXT.body,
  );
  page.list(
    [
      'Open the file in VS Code.',
      'Use Page Up and Page Down to move between the pages.',
      'Click a link: external links open in the browser, internal links jump inside the document.',
    ],
    PDF_TEXT.body,
    { numbered: true },
  );

  page.heading('4.2  A long paragraph', PDF_TEXT.h2, 'long');
  page.paragraph(
    'A PDF file is a set of numbered objects: dictionaries, arrays, strings, numbers and streams. The catalog points ' +
      'to the page tree, the outline and the named destinations; every page points to its resources (fonts, images ' +
      'and graphics states), its content stream and its annotations. The cross-reference table at the end of the ' +
      'file stores the byte offset of every object, which is why a viewer can open page 4 without reading pages 1 ' +
      'to 3 first. Text is placed glyph by glyph with explicit coordinates, so a viewer that wants to select or ' +
      'search text has to rebuild words and lines from those positions — this paragraph is long enough to check ' +
      'that wrapped lines come back in the right order.',
    PDF_TEXT.body,
  );
  page.paragraph(
    [
      { text: 'Back to the first page', link: { name: 'page1' } },
      ' or to ',
      { text: 'the landscape table', link: { name: 'landscape-table' } },
      ' (both links use named destinations).',
    ],
    PDF_TEXT.body,
  );
  page.paragraph('End of document.', { ...PDF_TEXT.body, font: 'italic' });
  page.checkFits('page 4');
  pdfFooter(page, 4);
  return page;
}

/** Outline tree: page index, anchor name (null = fit the whole page), bold / colour / closed flags, children. */
const PDF_OUTLINE = [
  {
    title: '1  Introduction',
    page: 0,
    anchor: 'intro',
    bold: true,
    children: [
      { title: '1.1  Links', page: 0, anchor: 'links' },
      { title: '1.2  Special characters — café, €, ©', page: 0, anchor: 'special' },
      { title: '1.3  Code', page: 0, anchor: 'code' },
    ],
  },
  {
    title: '2  Image',
    page: 1,
    anchor: 'image',
    bold: true,
    closed: true,
    children: [{ title: '2.1  Vector graphics', page: 1, anchor: 'vector' }],
  },
  { title: '3  Landscape page', page: 2, anchor: null, bold: true, color: 'C00000' },
  {
    title: '4  Searching and selecting text',
    page: 3,
    anchor: 'search',
    bold: true,
    children: [
      { title: '4.1  Lists', page: 3, anchor: 'lists' },
      { title: '4.2  A long paragraph', page: 3, anchor: 'long' },
    ],
  },
];

/**
 * Writes outline items (recursively) under `parent`. Returns first/last ids and the number of items that are
 * visible when the parent is open (closed items hide their descendants and get a negative /Count).
 */
function writePdfOutline(pdf, items, parent, destination) {
  const ids = items.map(() => pdf.reserve());
  let visible = 0;
  items.forEach((item, i) => {
    let entries = `/Title ${pdfTextString(item.title)} /Parent ${parent} 0 R /Dest ${destination(item.page, item.anchor)}`;
    if (i > 0) entries += ` /Prev ${ids[i - 1]} 0 R`;
    if (i < ids.length - 1) entries += ` /Next ${ids[i + 1]} 0 R`;
    if (item.children) {
      const child = writePdfOutline(pdf, item.children, ids[i], destination);
      entries += ` /First ${child.first} 0 R /Last ${child.last} 0 R /Count ${item.closed ? -child.visible : child.visible}`;
      if (!item.closed) visible += child.visible;
    }
    if (item.color) entries += ` /C [${pdfColor(item.color)}]`;
    if (item.bold) entries += ' /F 2';
    pdf.set(ids[i], `<< ${entries} >>`);
    visible++;
  });
  return { first: ids[0], last: ids[ids.length - 1], visible };
}

/** Assembles sample.pdf: fonts, image, pages with annotations, outline, named destinations, catalog, info. */
function generatePdf() {
  const pdf = new PdfWriter();
  const catalogId = pdf.reserve();
  const pagesId = pdf.reserve();
  const pageIds = Array.from({ length: PDF_PAGE_COUNT }, () => pdf.reserve());

  const fontEntries = Object.values(PDF_FONTS)
    .map(font => `/${font.name} ${pdf.add(`<< /Type /Font /Subtype /Type1 /BaseFont /${font.baseFont} /Encoding /WinAnsiEncoding >>`)} 0 R`)
    .join(' ');
  const imageId = pdf.addStream(
    `/Type /XObject /Subtype /Image /Width ${PDF_IMAGE.width} /Height ${PDF_IMAGE.height} /ColorSpace /DeviceRGB /BitsPerComponent 8`,
    rawRgb(PDF_IMAGE.width, PDF_IMAGE.height, gradientPixel(PDF_IMAGE.width, PDF_IMAGE.height)),
  );
  const graphicsStateId = pdf.add('<< /Type /ExtGState /ca 0.6 /CA 0.6 >>');
  const resourcesId = pdf.add(
    `<< /Font << ${fontEntries} >> /XObject << /${PDF_IMAGE.name} ${imageId} 0 R >> ` +
      `/ExtGState << /GS1 ${graphicsStateId} 0 R >> /ProcSet [/PDF /Text /ImageC] >>`,
  );

  const pages = [layoutPdfPage1(), layoutPdfPage2(), layoutPdfPage3(), layoutPdfPage4()];
  /** Explicit destination: /XYZ at an anchor (zoom unchanged) or the whole page (/Fit). */
  const destination = (pageIndex, anchorName) => {
    if (!anchorName) return `[${pageIds[pageIndex]} 0 R /Fit]`;
    const top = pages[pageIndex].anchors[anchorName];
    if (top === undefined) throw new Error(`PDF: unknown anchor ${anchorName} on page ${pageIndex + 1}`);
    return `[${pageIds[pageIndex]} 0 R /XYZ ${PDF_MARGIN} ${top} null]`;
  };
  const namedDestinations = { 'landscape-table': destination(2, 'table'), page1: destination(0, 'top') };

  pages.forEach((page, index) => {
    const contentId = pdf.addStream('', Buffer.from(page.ops.join('\n'), 'latin1'));
    const annotIds = page.links.map(({ rect, target }) => {
      const action = target.uri
        ? `/A << /Type /Action /S /URI /URI ${pdfTextString(target.uri)} >>`
        : target.name
          ? `/A << /Type /Action /S /GoTo /D ${pdfTextString(target.name)} >>`
          : `/Dest ${destination(target.page, target.fit ? null : target.anchor)}`;
      return pdf.add(
        `<< /Type /Annot /Subtype /Link /Rect [${rect.map(pdfNumber).join(' ')}] /Border [0 0 0] /F 4 ` +
          `/P ${pageIds[index]} 0 R ${action} >>`,
      );
    });
    pdf.set(
      pageIds[index],
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${page.width} ${page.height}] /Resources ${resourcesId} 0 R ` +
        `/Contents ${contentId} 0 R${annotIds.length ? ` /Annots [${annotIds.map(id => `${id} 0 R`).join(' ')}]` : ''} >>`,
    );
  });
  pdf.set(pagesId, `<< /Type /Pages /Kids [${pageIds.map(id => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`);

  const outlinesId = pdf.reserve();
  const outline = writePdfOutline(pdf, PDF_OUTLINE, outlinesId, destination);
  pdf.set(outlinesId, `<< /Type /Outlines /First ${outline.first} 0 R /Last ${outline.last} 0 R /Count ${outline.visible} >>`);

  // Name tree leaf: keys must be sorted (byte order).
  const names = Object.keys(namedDestinations).sort();
  const destsId = pdf.add(`<< /Names [${names.map(name => `${pdfTextString(name)} ${namedDestinations[name]}`).join(' ')}] >>`);

  pdf.set(
    catalogId,
    `<< /Type /Catalog /Pages ${pagesId} 0 R /Outlines ${outlinesId} 0 R /Names << /Dests ${destsId} 0 R >> ` +
      '/PageMode /UseOutlines /PageLayout /OneColumn /ViewerPreferences << /DisplayDocTitle true >> /Lang (en-US) >>',
  );
  const stamp = pdfDate(FIXED_DATE);
  const infoId = pdf.add(
    '<< /Title (FileStudio Sample PDF) /Subject (Hand-written PDF used to test the PDF preview) ' +
      `/Keywords (file-viewer, sample, pdf) /Author (${GENERATOR_NAME}) /Creator (test/generate-samples.js) ` +
      `/Producer (${GENERATOR_NAME}) /CreationDate (${stamp}) /ModDate (${stamp}) >>`,
  );
  return pdf.toBuffer({ root: catalogId, info: infoId });
}

// ===== PPTX: XML HELPERS =====

const PPTX_NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  table: 'http://schemas.openxmlformats.org/drawingml/2006/table',
};
/** Namespace declarations of every PresentationML part root. */
const PML_ROOT = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${PPTX_NS.p}"`;
const PML_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml';
const EMU_PER_PT = 12700; // 914400 EMU per inch / 72 pt per inch
/** 16:9 slide in points (13.333 x 7.5 in = 12192000 x 6858000 EMU); the notes page is 7.5 x 10 in. */
const PPTX_SLIDE = { width: 960, height: 540 };
const PPTX_NOTES_PAGE = { width: 540, height: 720 };

/** Points -> EMU. */
function emu(points) {
  return Math.round(points * EMU_PER_PT);
}

/** Points -> hundredths of a point (the unit of `sz` attributes). */
function centiPoints(points) {
  return Math.round(points * 100);
}

/** A relationships part. `type` is a short officeDocument relationship name ('slide', 'image', …) or a full URI. */
function pptxRels(list) {
  const rel = ({ id, type, target, external }) =>
    `<Relationship Id="${id}" Type="${type.includes('://') ? type : `${NS.relType}/${type}`}" Target="${xmlEscape(target)}"` +
    `${external ? ' TargetMode="External"' : ''}/>`;
  return XML_DECLARATION + `<Relationships xmlns="${NS.rel}">${list.map(rel).join('')}</Relationships>`;
}

/**
 * DrawingML color: '4472C4' -> srgbClr, 'accent1' / 'tx1' / 'hlink' -> schemeClr,
 * { scheme | srgb, lumMod, lumOff, tint, shade, alpha } -> the color with transforms.
 */
function drawingColor(color) {
  if (typeof color === 'string') {
    return /^[0-9A-F]{6}$/i.test(color) ? `<a:srgbClr val="${color}"/>` : `<a:schemeClr val="${color}"/>`;
  }
  const tag = color.srgb ? 'srgbClr' : 'schemeClr';
  const transforms = ['tint', 'shade', 'lumMod', 'lumOff', 'alpha']
    .filter(name => color[name] !== undefined)
    .map(name => `<a:${name} val="${color[name]}"/>`)
    .join('');
  return `<a:${tag} val="${color.srgb || color.scheme}">${transforms}</a:${tag}>`;
}

/**
 * Fill: { none: true }, { solid: color }, { gradient: { angle (deg), stops: [[pos %, color], …] } },
 * { pattern: { prst, fg, bg } } or { image: relId }.
 */
function drawingFill(fill) {
  if (!fill) return '';
  if (fill.none) return '<a:noFill/>';
  if (fill.solid) return `<a:solidFill>${drawingColor(fill.solid)}</a:solidFill>`;
  if (fill.gradient) {
    const { angle = 90, stops } = fill.gradient;
    const gs = stops.map(([pos, color]) => `<a:gs pos="${pos * 1000}">${drawingColor(color)}</a:gs>`).join('');
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst><a:lin ang="${angle * 60000}" scaled="0"/></a:gradFill>`;
  }
  if (fill.pattern) {
    const { prst, fg, bg } = fill.pattern;
    return `<a:pattFill prst="${prst}"><a:fgClr>${drawingColor(fg)}</a:fgClr><a:bgClr>${drawingColor(bg)}</a:bgClr></a:pattFill>`;
  }
  if (fill.image) return `<a:blipFill rotWithShape="1"><a:blip r:embed="${fill.image}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
  throw new Error(`PPTX: unknown fill ${JSON.stringify(fill)}`);
}

/** Outline: { none: true } or { width (pt), color, dash (prstDash), head, tail (line-end types) }. */
function drawingLine(line) {
  if (!line) return '';
  if (line.none) return '<a:ln><a:noFill/></a:ln>';
  return (
    `<a:ln w="${emu(line.width || 1)}"><a:solidFill>${drawingColor(line.color || 'tx1')}</a:solidFill>` +
    (line.dash ? `<a:prstDash val="${line.dash}"/>` : '') +
    (line.head ? `<a:headEnd type="${line.head}" w="med" len="med"/>` : '') +
    (line.tail ? `<a:tailEnd type="${line.tail}" w="med" len="med"/>` : '') +
    '</a:ln>'
  );
}

/** Position { x, y, w, h (pt), rot (deg), flipH, flipV } -> <a:xfrm> (or `tag`, e.g. p:xfrm); `extra` = chOff/chExt. */
function drawingXfrm({ x, y, w, h, rot, flipH, flipV }, tag = 'a:xfrm', extra = '') {
  const attrs = (rot ? ` rot="${Math.round(rot * 60000)}"` : '') + (flipH ? ' flipH="1"' : '') + (flipV ? ' flipV="1"' : '');
  return `<${tag}${attrs}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/>${extra}</${tag}>`;
}

function presetGeometry(prst, adjust = {}) {
  const gd = Object.entries(adjust).map(([name, value]) => `<a:gd name="${name}" fmla="val ${value}"/>`).join('');
  return `<a:prstGeom prst="${prst}">${gd ? `<a:avLst>${gd}</a:avLst>` : '<a:avLst/>'}</a:prstGeom>`;
}

/** Freeform geometry: `path` is [['M', x, y], ['L', x, y], ['C', x1, y1, x2, y2, x, y], ['Z']] in pt inside the box. */
function customGeometry(path, box) {
  const point = (x, y) => `<a:pt x="${emu(x)}" y="${emu(y)}"/>`;
  const commands = path
    .map(([command, ...v]) => {
      switch (command) {
        case 'M': return `<a:moveTo>${point(v[0], v[1])}</a:moveTo>`;
        case 'L': return `<a:lnTo>${point(v[0], v[1])}</a:lnTo>`;
        case 'C': return `<a:cubicBezTo>${point(v[0], v[1])}${point(v[2], v[3])}${point(v[4], v[5])}</a:cubicBezTo>`;
        case 'Z': return '<a:close/>';
        default: throw new Error(`PPTX: unknown path command ${command}`);
      }
    })
    .join('');
  return (
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
    `<a:pathLst><a:path w="${emu(box.w)}" h="${emu(box.h)}">${commands}</a:path></a:pathLst></a:custGeom>`
  );
}

/** A text run. props: sz (pt), b, i, u, strike, baseline (e.g. 30000 = superscript), color, font, link (relId). */
function textRun(text, props = {}) {
  let attrs = ' lang="en-US"';
  if (props.sz) attrs += ` sz="${centiPoints(props.sz)}"`;
  if (props.b) attrs += ' b="1"';
  if (props.i) attrs += ' i="1"';
  if (props.u) attrs += ' u="sng"';
  if (props.strike) attrs += ' strike="sngStrike"';
  if (props.baseline) attrs += ` baseline="${props.baseline}"`;
  attrs += ' dirty="0"';
  let children = '';
  if (props.color) children += `<a:solidFill>${drawingColor(props.color)}</a:solidFill>`;
  if (props.font) children += `<a:latin typeface="${props.font}"/><a:cs typeface="${props.font}"/>`;
  if (props.link) children += `<a:hlinkClick r:id="${props.link}"/>`;
  return `<a:r><a:rPr${attrs}${children ? `>${children}</a:rPr>` : '/>'}<a:t>${xmlEscape(text)}</a:t></a:r>`;
}

/**
 * A paragraph. `runs` is plain text (one run) or an array of run XML. options: lvl (0-based), align ('l' | 'ctr' |
 * 'r'), bullet (a character, 'num' for 1. 2. 3., 'none'), marL / indent (pt), sz (pt, plain text).
 */
function textParagraph(runs, options = {}) {
  const content = typeof runs === 'string' ? textRun(runs, { sz: options.sz }) : runs.join('');
  let attrs = '';
  if (options.marL !== undefined) attrs += ` marL="${emu(options.marL)}"`;
  if (options.lvl) attrs += ` lvl="${options.lvl}"`;
  if (options.indent !== undefined) attrs += ` indent="${emu(options.indent)}"`;
  if (options.align) attrs += ` algn="${options.align}"`;
  let children = '';
  if (options.bullet === 'num') children = '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
  else if (options.bullet === 'none') children = '<a:buNone/>';
  else if (options.bullet) children = `<a:buFont typeface="Arial" panose="020B0604020202020204"/><a:buChar char="${xmlEscape(options.bullet)}"/>`;
  const pPr = attrs || children ? `<a:pPr${attrs}${children ? `>${children}</a:pPr>` : '/>'}` : '';
  return `<a:p>${pPr}${content}<a:endParaRPr lang="en-US"${options.sz ? ` sz="${centiPoints(options.sz)}"` : ''} dirty="0"/></a:p>`;
}

/**
 * A text body (`p:txBody`, or `a:txBody` in table cells). options: anchor ('t' | 'ctr' | 'b'), wrap ('square' |
 * 'none'), inset (pt, all sides), autofit ('shape' | 'normal'), lstStyle (level-style XML).
 */
function textBody(paragraphs, options = {}, tag = 'p:txBody') {
  let attrs = '';
  if (options.wrap) attrs += ` wrap="${options.wrap}"`;
  if (options.inset !== undefined) {
    const inset = emu(options.inset);
    attrs += ` lIns="${inset}" tIns="${inset}" rIns="${inset}" bIns="${inset}"`;
  }
  if (options.anchor) attrs += ` anchor="${options.anchor}"`;
  const fit = options.autofit === 'shape' ? '<a:spAutoFit/>' : options.autofit === 'normal' ? '<a:normAutofit/>' : '';
  const bodyPr = fit ? `<a:bodyPr${attrs}>${fit}</a:bodyPr>` : `<a:bodyPr${attrs}/>`;
  const lstStyle = options.lstStyle ? `<a:lstStyle>${options.lstStyle}</a:lstStyle>` : '<a:lstStyle/>';
  return `<${tag}>${bodyPr}${lstStyle}${paragraphs.join('')}</${tag}>`;
}

/**
 * One level of a text style (<a:lvlNpPr>, level 1-9). sz (pt), marL / indent (pt), bullet (a character, 'none',
 * or undefined = inherit), lnSpc (%), spcBef (pt), color, font ('+mj-lt' / '+mn-lt'), align.
 */
function levelStyle(level, { sz, marL = 0, indent = 0, bullet, lnSpc, spcBef, color = 'tx1', font = '+mn-lt', align = 'l' }) {
  const spacing =
    (lnSpc ? `<a:lnSpc><a:spcPct val="${lnSpc * 1000}"/></a:lnSpc>` : '') +
    (spcBef !== undefined ? `<a:spcBef><a:spcPts val="${centiPoints(spcBef)}"/></a:spcBef>` : '');
  const bulletXml =
    bullet === undefined ? '' : bullet === 'none' ? '<a:buNone/>'
      : `<a:buFont typeface="Arial" panose="020B0604020202020204"/><a:buChar char="${xmlEscape(bullet)}"/>`;
  const fonts = `<a:latin typeface="${font}"/><a:ea typeface="${font.replace('-lt', '-ea')}"/><a:cs typeface="${font.replace('-lt', '-cs')}"/>`;
  return (
    `<a:lvl${level}pPr marL="${emu(marL)}" indent="${emu(indent)}" algn="${align}" defTabSz="914400" rtl="0" eaLnBrk="1" ` +
    `latinLnBrk="0" hangingPunct="1">${spacing}${bulletXml}` +
    `<a:defRPr sz="${centiPoints(sz)}" kern="1200"><a:solidFill>${drawingColor(color)}</a:solidFill>${fonts}</a:defRPr></a:lvl${level}pPr>`
  );
}

/** <p:cNvPr> with an optional description and click hyperlink (relId). */
function nonVisualProps({ id, name, descr, link }) {
  const attrs = `id="${id}" name="${xmlEscape(name)}"${descr ? ` descr="${xmlEscape(descr)}"` : ''}`;
  return link ? `<p:cNvPr ${attrs}><a:hlinkClick r:id="${link}"/></p:cNvPr>` : `<p:cNvPr ${attrs}/>`;
}

const SLIDE_IMAGE_LOCKS = 'noGrp="1" noRot="1" noChangeAspect="1"';

/**
 * A shape (<p:sp>). spec: id, name, ph ({ type, idx }), box (see drawingXfrm; omitted for placeholders that inherit
 * their position), geom (preset name — default 'rect' for non-placeholders — or { path } for a freeform), adjust,
 * fill, line, text (paragraph XML array), body (textBody options), txBox, link (relId), locks (spLocks attributes).
 */
function pptxShape(spec) {
  const locks = spec.locks || (spec.ph ? 'noGrp="1"' : '');
  const cNvSpPr = `<p:cNvSpPr${spec.txBox ? ' txBox="1"' : ''}${locks ? `><a:spLocks ${locks}/></p:cNvSpPr>` : '/>'}`;
  const ph = spec.ph
    ? `<p:ph${spec.ph.type ? ` type="${spec.ph.type}"` : ''}${spec.ph.idx !== undefined ? ` idx="${spec.ph.idx}"` : ''}/>`
    : '';
  let spPr = spec.box ? drawingXfrm(spec.box) : '';
  if (spec.geom && typeof spec.geom === 'object') spPr += customGeometry(spec.geom.path, spec.box);
  else if (spec.geom || !spec.ph) spPr += presetGeometry(spec.geom || 'rect', spec.adjust);
  spPr += drawingFill(spec.fill) + drawingLine(spec.line);
  return (
    `<p:sp><p:nvSpPr>${nonVisualProps(spec)}${cNvSpPr}<p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    (spPr ? `<p:spPr>${spPr}</p:spPr>` : '<p:spPr/>') +
    (spec.text ? textBody(spec.text, spec.body) : '') +
    '</p:sp>'
  );
}

/** A connector (<p:cxnSp>): straightConnector1 / bentConnector3 / … with line ends. */
function pptxConnector({ id, name, box, geom = 'straightConnector1', line }) {
  return (
    `<p:cxnSp><p:nvCxnSpPr>${nonVisualProps({ id, name })}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr>${drawingXfrm(box)}${presetGeometry(geom)}${drawingLine(line)}</p:spPr></p:cxnSp>`
  );
}

/** A picture (<p:pic>) stretched over `box`. */
function pptxPicture({ id, name, descr, relId, box, line, link }) {
  return (
    `<p:pic><p:nvPicPr>${nonVisualProps({ id, name, descr, link })}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${drawingXfrm(box)}${presetGeometry('rect')}${drawingLine(line)}</p:spPr></p:pic>`
  );
}

/** A group (<p:grpSp>): `children` are positioned in the child space `child` ({ x, y, w, h } pt), mapped onto `box`. */
function pptxGroup({ id, name, box, child = box, children }) {
  const childSpace = `<a:chOff x="${emu(child.x)}" y="${emu(child.y)}"/><a:chExt cx="${emu(child.w)}" cy="${emu(child.h)}"/>`;
  return (
    `<p:grpSp><p:nvGrpSpPr>${nonVisualProps({ id, name })}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr>${drawingXfrm(box, 'a:xfrm', childSpace)}</p:grpSpPr>${children.join('')}</p:grpSp>`
  );
}

/** A graphic frame (table or chart) with its <a:graphicData> payload. */
function pptxGraphicFrame({ id, name, box, uri, data, locks = '' }) {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${nonVisualProps({ id, name })}` +
    `<p:cNvGraphicFramePr>${locks ? `<a:graphicFrameLocks ${locks}/>` : ''}</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `${drawingXfrm(box, 'p:xfrm')}<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** The shape tree of a slide, layout, master or notes page. */
function shapeTree(children) {
  return (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    `${children.join('')}</p:spTree>`
  );
}

/**
 * A DrawingML table. `rows` is the full grid: a cell is a string or { text, colSpan, rowSpan, fill, color, bold, align };
 * positions covered by a merge are null and are written as hMerge / vMerge continuation cells.
 */
function pptxTable(rows, { columnWidths, rowHeight }) {
  const covered = rows.map(row => row.map(() => null));
  rows.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (!cell || typeof cell !== 'object') return;
      for (let dr = 0; dr < (cell.rowSpan || 1); dr++) {
        for (let dc = 0; dc < (cell.colSpan || 1); dc++) {
          if (dr || dc) covered[r + dr][c + dc] = { h: dc > 0, v: dr > 0 };
        }
      }
    }),
  );
  const borders = ['lnL', 'lnR', 'lnT', 'lnB']
    .map(side => `<a:${side} w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:${side}>`)
    .join('');
  const cellXml = (cell, r, c) => {
    const merge = covered[r][c];
    if (merge) {
      if (cell !== null) throw new Error(`PPTX table: cell ${r},${c} is covered by a merge and must be null`);
      const attrs = (merge.h ? ' hMerge="1"' : '') + (merge.v ? ' vMerge="1"' : '');
      return `<a:tc${attrs}>${textBody([textParagraph([])], {}, 'a:txBody')}<a:tcPr>${borders}</a:tcPr></a:tc>`;
    }
    if (cell === null) throw new Error(`PPTX table: cell ${r},${c} is null but not covered by a merge`);
    const spec = typeof cell === 'string' ? { text: cell } : cell;
    const attrs = (spec.colSpan > 1 ? ` gridSpan="${spec.colSpan}"` : '') + (spec.rowSpan > 1 ? ` rowSpan="${spec.rowSpan}"` : '');
    const text = textParagraph([textRun(spec.text, { sz: 14, b: spec.bold, color: spec.color })], { align: spec.align || 'l' });
    return (
      `<a:tc${attrs}>${textBody([text], {}, 'a:txBody')}` +
      `<a:tcPr marL="${emu(7)}" marR="${emu(7)}" anchor="ctr">${borders}${spec.fill ? drawingFill({ solid: spec.fill }) : ''}</a:tcPr></a:tc>`
    );
  };
  const grid = columnWidths.map(width => `<a:gridCol w="${emu(width)}"/>`).join('');
  const body = rows.map((row, r) => `<a:tr h="${emu(rowHeight)}">${row.map((cell, c) => cellXml(cell, r, c)).join('')}</a:tr>`).join('');
  return `<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${body}</a:tbl>`;
}

// ===== PPTX: PARTS =====

const PPTX_LAYOUT = { title: 1, content: 2, titleOnly: 3 };

function pptxContentTypes({ slideCount, notesCount, chartCount, layoutCount }) {
  const override = (part, type) => `<Override PartName="${part}" ContentType="${type}"/>`;
  const range = (count, make) => Array.from({ length: count }, (_, i) => make(i + 1)).join('');
  const theme = 'application/vnd.openxmlformats-officedocument.theme+xml';
  return (
    XML_DECLARATION +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    override('/ppt/presentation.xml', `${PML_TYPE}.presentation.main+xml`) +
    override('/ppt/slideMasters/slideMaster1.xml', `${PML_TYPE}.slideMaster+xml`) +
    range(layoutCount, n => override(`/ppt/slideLayouts/slideLayout${n}.xml`, `${PML_TYPE}.slideLayout+xml`)) +
    range(slideCount, n => override(`/ppt/slides/slide${n}.xml`, `${PML_TYPE}.slide+xml`)) +
    override('/ppt/notesMasters/notesMaster1.xml', `${PML_TYPE}.notesMaster+xml`) +
    range(notesCount, n => override(`/ppt/notesSlides/notesSlide${n}.xml`, `${PML_TYPE}.notesSlide+xml`)) +
    range(chartCount, n => override(`/ppt/charts/chart${n}.xml`, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')) +
    override('/ppt/theme/theme1.xml', theme) +
    override('/ppt/theme/theme2.xml', theme) +
    override('/ppt/presProps.xml', `${PML_TYPE}.presProps+xml`) +
    override('/ppt/viewProps.xml', `${PML_TYPE}.viewProps+xml`) +
    override('/ppt/tableStyles.xml', `${PML_TYPE}.tableStyles+xml`) +
    override('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml') +
    override('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml') +
    '</Types>'
  );
}

function pptxCoreProps() {
  const stamp = FIXED_DATE.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return (
    XML_DECLARATION +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>FileStudio Sample Presentation</dc:title>' +
    '<dc:subject>Hand-built PresentationML used to test the .pptx preview</dc:subject>' +
    `<dc:creator>${GENERATOR_NAME}</dc:creator>` +
    '<cp:keywords>file-viewer, sample, pptx</cp:keywords>' +
    `<cp:lastModifiedBy>${GENERATOR_NAME}</cp:lastModifiedBy>` +
    '<cp:revision>1</cp:revision>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '</cp:coreProperties>'
  );
}

function pptxAppProps({ slideCount, notesCount, hiddenCount }) {
  return (
    XML_DECLARATION +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<TotalTime>0</TotalTime><Application>${GENERATOR_NAME}</Application><PresentationFormat>Widescreen</PresentationFormat>` +
    `<Slides>${slideCount}</Slides><Notes>${notesCount}</Notes><HiddenSlides>${hiddenCount}</HiddenSlides><MMClips>0</MMClips>` +
    '<ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>' +
    '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion>' +
    '</Properties>'
  );
}

/** presentation.xml.rels: master (rId1), notes master (rId2), slides (rId3…), then props, theme and table styles. */
function pptxPresentationRels(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, i) => ({ id: `rId${i + 3}`, type: 'slide', target: `slides/slide${i + 1}.xml` }));
  const next = slideCount + 3;
  return pptxRels([
    { id: 'rId1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' },
    { id: 'rId2', type: 'notesMaster', target: 'notesMasters/notesMaster1.xml' },
    ...slides,
    { id: `rId${next}`, type: 'presProps', target: 'presProps.xml' },
    { id: `rId${next + 1}`, type: 'viewProps', target: 'viewProps.xml' },
    { id: `rId${next + 2}`, type: 'theme', target: 'theme/theme1.xml' },
    { id: `rId${next + 3}`, type: 'tableStyles', target: 'tableStyles.xml' },
  ]);
}

function pptxPresentation(slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`).join('');
  const levels = [1, 2, 3, 4, 5].map(level => levelStyle(level, { sz: 18, marL: 36 * (level - 1) })).join('');
  return (
    XML_DECLARATION +
    `<p:presentation ${PML_ROOT} saveSubsetFonts="1">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    '<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>' +
    `<p:sldIdLst>${slideIds}</p:sldIdLst>` +
    `<p:sldSz cx="${emu(PPTX_SLIDE.width)}" cy="${emu(PPTX_SLIDE.height)}"/>` +
    `<p:notesSz cx="${emu(PPTX_NOTES_PAGE.width)}" cy="${emu(PPTX_NOTES_PAGE.height)}"/>` +
    `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${levels}</p:defaultTextStyle>` +
    '</p:presentation>'
  );
}

function pptxPresProps() {
  return XML_DECLARATION + `<p:presentationPr ${PML_ROOT}/>`;
}

function pptxViewProps() {
  return (
    XML_DECLARATION +
    `<p:viewPr ${PML_ROOT}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>` +
    '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
  );
}

/** No custom table styles: the sample table is formatted cell by cell. */
function pptxTableStyles() {
  return XML_DECLARATION + `<a:tblStyleLst xmlns:a="${NS.a}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

/** The Office 2013+ theme: color scheme, Calibri Light / Calibri fonts, and the format scheme (fills, lines, effects). */
function pptxTheme() {
  const phClr = (transforms = []) =>
    `<a:schemeClr val="phClr">${transforms.map(([name, value]) => `<a:${name} val="${value}"/>`).join('')}</a:schemeClr>`;
  const gradient = stops =>
    `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([pos, transforms]) => `<a:gs pos="${pos}">${phClr(transforms)}</a:gs>`).join('')}` +
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>';
  const line = width =>
    `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${phClr()}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  return (
    XML_DECLARATION +
    `<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
    MODERN_THEME_CLR_SCHEME +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    '<a:fmtScheme name="Office"><a:fillStyleLst>' +
    `<a:solidFill>${phClr()}</a:solidFill>` +
    gradient([
      [0, [['lumMod', 110000], ['satMod', 105000], ['tint', 67000]]],
      [50000, [['lumMod', 105000], ['satMod', 103000], ['tint', 73000]]],
      [100000, [['lumMod', 105000], ['satMod', 109000], ['tint', 81000]]],
    ]) +
    gradient([
      [0, [['satMod', 103000], ['lumMod', 102000], ['tint', 94000]]],
      [50000, [['satMod', 110000], ['lumMod', 100000], ['shade', 100000]]],
      [100000, [['lumMod', 99000], ['satMod', 120000], ['shade', 78000]]],
    ]) +
    `</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">' +
    '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst><a:solidFill>${phClr()}</a:solidFill>` +
    `<a:solidFill>${phClr([['tint', 95000], ['satMod', 170000]])}</a:solidFill>` +
    gradient([
      [0, [['tint', 93000], ['satMod', 150000], ['shade', 98000], ['lumMod', 102000]]],
      [50000, [['tint', 98000], ['satMod', 130000], ['shade', 90000], ['lumMod', 103000]]],
      [100000, [['shade', 63000], ['satMod', 120000]]],
    ]) +
    '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
  );
}

const PPTX_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" ' +
  'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';

/** The master: title + body placeholders, an accent bar and a footer text box (drawn on every slide but the title slide). */
function pptxSlideMaster(layoutCount) {
  const bullets = ['•', '–', '•', '–', '»'];
  const bodyLevels = [24, 20, 18, 16, 16]
    .map((sz, i) => levelStyle(i + 1, { sz, marL: 18 + 36 * i, indent: -18, bullet: bullets[i], lnSpc: 90, spcBef: i ? 5 : 10 }))
    .join('');
  const otherLevels = [1, 2, 3, 4, 5].map(level => levelStyle(level, { sz: 18, marL: 36 * (level - 1) })).join('');
  const layoutIds = Array.from({ length: layoutCount }, (_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join('');
  const shapes = [
    pptxShape({
      id: 2, name: 'Title Placeholder 1', ph: { type: 'title' }, geom: 'rect', box: { x: 48, y: 24, w: 864, h: 84 },
      body: { anchor: 'ctr', autofit: 'normal' }, text: [textParagraph('Click to edit Master title style')],
    }),
    pptxShape({
      id: 3, name: 'Text Placeholder 2', ph: { type: 'body', idx: 1 }, geom: 'rect', box: { x: 48, y: 120, w: 864, h: 370 },
      body: { autofit: 'normal' },
      text: ['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level'].map((text, lvl) => textParagraph(text, { lvl })),
    }),
    pptxShape({ id: 7, name: 'Accent Bar', box: { x: 0, y: 528, w: 960, h: 12 }, fill: { solid: 'accent1' }, line: { none: true } }),
    pptxShape({
      id: 8, name: 'Deck Footer', txBox: true, box: { x: 48, y: 502, w: 420, h: 22 }, fill: { none: true },
      body: { wrap: 'square', inset: 0, anchor: 'ctr' }, text: [textParagraph([textRun('FileStudio sample deck', { sz: 10, color: '7F7F7F' })])],
    }),
  ];
  return (
    XML_DECLARATION +
    `<p:sldMaster ${PML_ROOT}>` +
    `<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${shapeTree(shapes)}</p:cSld>` +
    `<p:clrMap ${PPTX_CLR_MAP}/>` +
    `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>` +
    '<p:txStyles>' +
    `<p:titleStyle>${levelStyle(1, { sz: 40, bullet: 'none', lnSpc: 90, font: '+mj-lt' })}</p:titleStyle>` +
    `<p:bodyStyle>${bodyLevels}</p:bodyStyle>` +
    `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${otherLevels}</p:otherStyle>` +
    '</p:txStyles></p:sldMaster>'
  );
}

/** Title Slide (own decoration, master shapes hidden), Title and Content, Title Only. */
function pptxSlideLayouts() {
  const title = (box, extra = {}) =>
    pptxShape({ id: 2, name: 'Title 1', ph: { type: extra.type || 'title' }, box, body: extra.body, text: [textParagraph('Click to edit Master title style')] });
  return [
    {
      name: 'Title Slide', type: 'title', showMasterSp: false,
      shapes: [
        title({ x: 80, y: 110, w: 800, h: 170 }, {
          type: 'ctrTitle',
          body: { anchor: 'b', lstStyle: '<a:lvl1pPr algn="ctr"><a:defRPr sz="4800"/></a:lvl1pPr>' },
        }),
        pptxShape({
          id: 3, name: 'Subtitle 2', ph: { type: 'subTitle', idx: 1 }, box: { x: 120, y: 296, w: 720, h: 110 },
          body: { lstStyle: '<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:defRPr></a:lvl1pPr>' },
          text: [textParagraph('Click to edit Master subtitle style')],
        }),
        pptxShape({ id: 7, name: 'Title Band', box: { x: 0, y: 450, w: 960, h: 90 }, fill: { gradient: { angle: 0, stops: [[0, 'accent1'], [100, 'accent5']] } }, line: { none: true } }),
        pptxShape({ id: 8, name: 'Title Band Edge', box: { x: 0, y: 444, w: 960, h: 6 }, fill: { solid: 'accent2' }, line: { none: true } }),
      ],
    },
    {
      name: 'Title and Content', type: 'obj',
      shapes: [
        title(),
        pptxShape({ id: 3, name: 'Content Placeholder 2', ph: { idx: 1 }, text: ['Click to edit Master text styles', 'Second level', 'Third level'].map((text, lvl) => textParagraph(text, { lvl })) }),
      ],
    },
    { name: 'Title Only', type: 'titleOnly', shapes: [title()] },
  ];
}

function pptxSlideLayout(layout) {
  return (
    XML_DECLARATION +
    `<p:sldLayout ${PML_ROOT}${layout.showMasterSp === false ? ' showMasterSp="0"' : ''} type="${layout.type}" preserve="1">` +
    `<p:cSld name="${xmlEscape(layout.name)}">${shapeTree(layout.shapes)}</p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
  );
}

/** Notes master: slide image above, notes body below (7.5 x 10 in page). */
function pptxNotesMaster() {
  const notesLevel = levelStyle(1, { sz: 12 });
  const shapes = [
    pptxShape({
      id: 2, name: 'Slide Image Placeholder 1', ph: { type: 'sldImg', idx: 2 }, locks: SLIDE_IMAGE_LOCKS, geom: 'rect',
      box: { x: 60, y: 54, w: 420, h: 236.25 }, fill: { none: true }, line: { width: 1, color: '000000' },
    }),
    pptxShape({
      id: 3, name: 'Notes Placeholder 2', ph: { type: 'body', idx: 3 }, geom: 'rect', box: { x: 54, y: 324, w: 432, h: 324 },
      text: [textParagraph('Click to edit Master text styles')],
    }),
  ];
  return (
    XML_DECLARATION +
    `<p:notesMaster ${PML_ROOT}>` +
    `<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${shapeTree(shapes)}</p:cSld>` +
    `<p:clrMap ${PPTX_CLR_MAP}/><p:notesStyle>${notesLevel}</p:notesStyle></p:notesMaster>`
  );
}

/** A notes page: slide image placeholder + the notes text (one paragraph per string). */
function pptxNotesSlide(paragraphs) {
  const shapes = [
    pptxShape({ id: 2, name: 'Slide Image Placeholder 1', ph: { type: 'sldImg' }, locks: SLIDE_IMAGE_LOCKS }),
    pptxShape({ id: 3, name: 'Notes Placeholder 2', ph: { type: 'body', idx: 1 }, text: paragraphs.map(text => textParagraph(text)) }),
  ];
  return (
    XML_DECLARATION +
    `<p:notes ${PML_ROOT}><p:cSld>${shapeTree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
  );
}

function pptxSlide(slide, rels) {
  const background = slide.background ? `<p:bg><p:bgPr>${drawingFill(slide.background)}<a:effectLst/></p:bgPr></p:bg>` : '';
  return (
    XML_DECLARATION +
    `<p:sld ${PML_ROOT}${slide.hidden ? ' show="0"' : ''}>` +
    `<p:cSld>${background}${shapeTree(slide.shapes(rels))}</p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  );
}

// ===== PPTX: SLIDES =====

/** Centered bold label inside a shape. */
function shapeLabel(text, color = 'FFFFFF', sz = 12) {
  return { text: [textParagraph([textRun(text, { sz, b: true, color })], { align: 'ctr' })], body: { anchor: 'ctr', inset: 4 } };
}

/** Preset shapes of slide 3: [preset, label, fill]; the callout gets a pattern fill. */
const PPTX_PRESET_SHAPES = [
  ['rect', 'Rectangle', 'accent1'],
  ['roundRect', 'Rounded rectangle', 'accent2'],
  ['ellipse', 'Ellipse', 'accent5'],
  ['triangle', 'Triangle', 'accent4'],
  ['diamond', 'Diamond', 'accent6'],
  ['hexagon', 'Hexagon', 'accent3'],
  ['rightArrow', 'Right arrow', 'accent1'],
  ['chevron', 'Chevron', 'accent2'],
  ['star5', 'Star', 'accent4'],
  ['parallelogram', 'Parallelogram', 'accent5'],
  ['homePlate', 'Pentagon', 'accent6'],
  ['wedgeRectCallout', 'Callout', null],
];

/** Units sold per region and half-year (slide 5). */
const PPTX_TABLE_ROWS = [
  ['North', 1204, 1310],
  ['South', 986, 1042],
  ['East', 1433, 1502],
  ['West', 877, 951],
];

function pptxTableSlideRows() {
  const head = { fill: 'accent1', color: 'FFFFFF', bold: true, align: 'ctr' };
  const number = (value, extra) => ({ text: groupedNumber(value), align: 'r', ...extra });
  const rows = [
    [{ text: 'Region', rowSpan: 2, ...head }, { text: 'Units sold', colSpan: 2, ...head }, null, { text: 'Total', rowSpan: 2, ...head }],
    [null, { text: 'H1 2024', ...head }, { text: 'H2 2024', ...head }, null],
  ];
  PPTX_TABLE_ROWS.forEach(([region, h1, h2], i) => {
    const band = i % 2 === 0 ? { fill: 'DEEBF7' } : {};
    rows.push([{ text: region, ...band }, number(h1, band), number(h2, band), number(h1 + h2, { bold: true, ...band })]);
  });
  const sum = column => PPTX_TABLE_ROWS.reduce((total, row) => total + row[column], 0);
  const total = { bold: true, fill: 'E7E6E6' };
  rows.push([{ text: 'All regions', ...total }, number(sum(1), total), number(sum(2), total), number(sum(1) + sum(2), total)]);
  return rows;
}

/**
 * The deck. Each slide: layout, hidden, background, notes (paragraphs), rels ({ name: { type, target, external } };
 * rId1 = layout, rId2 = notes, then these in order) and shapes(rels) -> shape XML, where rels maps names to rIds.
 */
function pptxSlideDefinitions() {
  const title = (text, color) => pptxShape({ id: 2, name: 'Title 1', ph: { type: 'title' }, text: [textParagraph([textRun(text, { color })])] });
  const caption = (id, box, text, extra = {}) =>
    pptxShape({
      id, name: `Caption ${id}`, txBox: true, box, fill: { none: true }, body: { wrap: 'square', anchor: 'ctr' },
      text: [textParagraph([textRun(text, { sz: 14, color: '595959', ...extra })], { align: extra.align })],
    });
  return [
    {
      layout: PPTX_LAYOUT.title,
      notes: [
        'Speaker notes for the title slide. The viewer shows them in the notes pane.',
        'Second paragraph with special characters: <tag> & "quotes" — café.',
      ],
      shapes: () => [
        pptxShape({ id: 2, name: 'Title 1', ph: { type: 'ctrTitle' }, text: [textParagraph('FileStudio Sample Presentation')] }),
        pptxShape({
          id: 3, name: 'Subtitle 2', ph: { type: 'subTitle', idx: 1 },
          text: [
            textParagraph('Hand-built PresentationML for the .pptx preview'),
            textParagraph([textRun('Generated by '), textRun('test/generate-samples.js', { font: 'Consolas' })]),
          ],
        }),
      ],
    },
    {
      layout: PPTX_LAYOUT.content,
      rels: { website: { type: 'hyperlink', target: 'https://code.visualstudio.com/', external: true } },
      notes: [
        'Bullets use explicit a:buChar / a:buAutoNum paragraph properties at three levels.',
        'The link opens through the extension (http, https and mailto only).',
      ],
      shapes: rels => [
        title('Bullets and text formatting'),
        pptxShape({
          id: 3, name: 'Content Placeholder 2', ph: { idx: 1 },
          text: [
            textParagraph(
              [
                textRun('First level with '), textRun('bold', { b: true }), textRun(', '), textRun('italic', { i: true }), textRun(', '),
                textRun('underlined', { u: true }), textRun(', '), textRun('struck-through', { strike: true }), textRun(' and '),
                textRun('colored', { color: 'C00000' }), textRun(' text'),
              ],
              { bullet: '•' },
            ),
            textParagraph('Second level bullet', { lvl: 1, bullet: '–' }),
            textParagraph('Third level bullet', { lvl: 2, bullet: '▪' }),
            textParagraph([textRun('A hyperlink to '), textRun('the Visual Studio Code website', { link: rels.website }), textRun('.')], { bullet: '•' }),
            textParagraph('Numbered item one', { bullet: 'num' }),
            textParagraph('Numbered item two', { bullet: 'num' }),
            textParagraph(
              [textRun('No bullet: café, naïve, 日本語, Ελληνικά, 🎉 · <tag> & "quotes" · E = mc'), textRun('2', { baseline: 30000 })],
              { bullet: 'none', marL: 0, indent: 0 },
            ),
          ],
        }),
      ],
    },
    {
      layout: PPTX_LAYOUT.titleOnly,
      shapes: () => {
        const presets = PPTX_PRESET_SHAPES.map(([geom, label, color], i) => {
          const box = { x: 60 + (i % 6) * 144, y: 124 + Math.floor(i / 6) * 112, w: 120, h: 90 };
          const fill = color ? { solid: color } : { pattern: { prst: 'ltUpDiag', fg: 'accent1', bg: 'FFFFFF' } };
          const line = { width: 1, color: color ? { scheme: color, lumMod: 75000 } : 'accent1' };
          return pptxShape({ id: 10 + i, name: `${label} ${10 + i}`, geom, box, fill, line, ...shapeLabel(label, color ? 'FFFFFF' : '1F3864') });
        });
        const wave = [['M', 0, 25], ['C', 30, 0, 90, 50, 120, 25], ['L', 120, 65], ['C', 90, 90, 30, 40, 0, 65], ['Z']];
        return [
          title('Preset shapes, connectors and transforms'),
          ...presets,
          pptxConnector({ id: 30, name: 'Straight Arrow Connector 30', box: { x: 60, y: 362, w: 120, h: 66 }, line: { width: 2, color: 'accent1', head: 'oval', tail: 'triangle' } }),
          pptxConnector({ id: 31, name: 'Elbow Connector 31', geom: 'bentConnector3', box: { x: 204, y: 362, w: 120, h: 66 }, line: { width: 2, color: 'accent2', dash: 'dash', tail: 'arrow' } }),
          pptxShape({ id: 32, name: 'Freeform 32', geom: { path: wave }, box: { x: 348, y: 350, w: 120, h: 90 }, fill: { solid: 'accent5' }, line: { width: 1, color: '1F3864' }, ...shapeLabel('Freeform') }),
          pptxShape({ id: 33, name: 'Rotated Rectangle 33', box: { x: 492, y: 355, w: 120, h: 80, rot: 30 }, fill: { none: true }, line: { width: 2, color: 'accent2', dash: 'dash' }, ...shapeLabel('Rotated 30°', 'C55A11') }),
          pptxShape({ id: 34, name: 'Flipped Triangle 34', geom: 'rtTriangle', box: { x: 636, y: 350, w: 120, h: 90, flipH: true }, fill: { solid: 'accent6' }, line: { none: true }, ...shapeLabel('Flipped') }),
          pptxShape({ id: 35, name: 'Gradient Shape 35', geom: 'roundRect', box: { x: 780, y: 350, w: 120, h: 90 }, fill: { gradient: { angle: 90, stops: [[0, 'FFC000'], [100, 'ED7D31']] } }, line: { none: true }, ...shapeLabel('Gradient') }),
          caption(36, { x: 60, y: 448, w: 840, h: 30 }, 'Connectors with line ends, a freeform path, rotation, a horizontal flip, a gradient and a pattern fill.'),
        ];
      },
    },
    {
      layout: PPTX_LAYOUT.titleOnly,
      rels: {
        gradient: { type: 'image', target: '../media/image1.png' },
        badge: { type: 'image', target: '../media/image2.png' },
        docs: { type: 'hyperlink', target: 'https://code.visualstudio.com/docs', external: true },
      },
      shapes: rels => [
        title('Group, pictures and a linked shape'),
        pptxGroup({
          id: 10, name: 'Group 10', box: { x: 60, y: 130, w: 360, h: 220 }, child: { x: 0, y: 0, w: 720, h: 440 },
          children: [
            pptxShape({ id: 11, name: 'Group Frame 11', geom: 'roundRect', box: { x: 0, y: 0, w: 720, h: 440 }, fill: { solid: 'F2F2F2' }, line: { width: 1, color: 'A5A5A5' } }),
            pptxShape({ id: 12, name: 'Oval 12', geom: 'ellipse', box: { x: 40, y: 40, w: 200, h: 200 }, fill: { solid: 'accent1' }, ...shapeLabel('A', 'FFFFFF', 28) }),
            pptxShape({ id: 13, name: 'Rectangle 13', box: { x: 260, y: 40, w: 200, h: 200 }, fill: { solid: 'accent2' }, ...shapeLabel('B', 'FFFFFF', 28) }),
            pptxShape({ id: 14, name: 'Triangle 14', geom: 'triangle', box: { x: 480, y: 40, w: 200, h: 200 }, fill: { solid: 'accent6' }, ...shapeLabel('C', 'FFFFFF', 28) }),
            pptxShape({
              id: 15, name: 'Group Text 15', txBox: true, box: { x: 40, y: 280, w: 640, h: 120 }, fill: { none: true }, body: { wrap: 'square', anchor: 'ctr' },
              text: [textParagraph([textRun('Three shapes and a text box in one group', { sz: 14, color: '404040' })], { align: 'ctr' })],
            }),
          ],
        }),
        caption(16, { x: 60, y: 356, w: 360, h: 26 }, 'Group: child space 720 × 440 pt shown at 50 %', { sz: 12, align: 'ctr' }),
        pptxPicture({ id: 20, name: 'Picture 20', descr: 'Gradient picture (opaque PNG)', relId: rels.gradient, box: { x: 480, y: 130, w: 240, h: 160 }, line: { width: 1, color: 'A5A5A5' } }),
        pptxPicture({ id: 21, name: 'Picture 21', descr: 'Badge (PNG with transparency), rotated 15°', relId: rels.badge, box: { x: 768, y: 150, w: 96, h: 96, rot: 15 } }),
        caption(22, { x: 480, y: 296, w: 420, h: 26 }, 'Pictures: an opaque gradient and a rotated transparent badge', { sz: 12 }),
        pptxShape({ id: 23, name: 'Link Button 23', geom: 'roundRect', link: rels.docs, box: { x: 480, y: 340, w: 260, h: 48 }, fill: { solid: 'accent1' }, line: { none: true }, ...shapeLabel('Open the VS Code docs', 'FFFFFF', 16) }),
      ],
    },
    {
      layout: PPTX_LAYOUT.titleOnly,
      shapes: () => [
        title('Table with merged cells'),
        pptxGraphicFrame({
          id: 10, name: 'Table 10', box: { x: 90, y: 130, w: 780, h: 238 }, uri: PPTX_NS.table, locks: 'noGrp="1"',
          data: pptxTable(pptxTableSlideRows(), { columnWidths: [240, 180, 180, 180], rowHeight: 34 }),
        }),
        caption(11, { x: 90, y: 390, w: 780, h: 40 }, '"Units sold" spans two columns; "Region" and "Total" span two header rows.'),
      ],
    },
    {
      layout: PPTX_LAYOUT.titleOnly,
      rels: {
        bar: { type: 'chart', target: '../charts/chart1.xml' },
        pie: { type: 'chart', target: '../charts/chart2.xml' },
      },
      notes: [
        `Bar chart: clustered columns, ${BAR_CHART_SERIES.length} series over ${CHART_QUARTERS.length} quarters (cached values only, no embedded workbook).`,
        'Pie chart: one series with per-slice colors and percentage labels.',
      ],
      shapes: rels => [
        title('Charts'),
        pptxGraphicFrame({ id: 10, name: 'Bar Chart 10', box: { x: 40, y: 116, w: 440, h: 376 }, uri: PPTX_NS.c, data: `<c:chart xmlns:c="${PPTX_NS.c}" r:id="${rels.bar}"/>` }),
        pptxGraphicFrame({ id: 11, name: 'Pie Chart 11', box: { x: 500, y: 116, w: 420, h: 376 }, uri: PPTX_NS.c, data: `<c:chart xmlns:c="${PPTX_NS.c}" r:id="${rels.pie}"/>` }),
      ],
    },
    {
      layout: PPTX_LAYOUT.content,
      hidden: true,
      shapes: () => [
        title('Hidden slide'),
        pptxShape({
          id: 3, name: 'Content Placeholder 2', ph: { idx: 1 },
          text: [
            textParagraph('This slide has show="0": PowerPoint skips it in a slide show.', { bullet: '•' }),
            textParagraph('The viewer still lists it, marked as hidden.', { bullet: '•' }),
          ],
        }),
      ],
    },
    {
      layout: PPTX_LAYOUT.titleOnly,
      background: { gradient: { angle: 45, stops: [[0, '1F3864'], [55, '4472C4'], [100, '9DC3E6']] } },
      notes: ['Last slide: a linear gradient background (45°, three stops) with white text on top.'],
      shapes: () => [
        title('Gradient background', 'FFFFFF'),
        pptxShape({
          id: 10, name: 'Closing Text 10', txBox: true, box: { x: 120, y: 190, w: 720, h: 160 }, fill: { none: true }, body: { wrap: 'square', anchor: 'ctr' },
          text: [
            textParagraph([textRun('Thank you!', { sz: 54, b: true, color: 'FFFFFF' })], { align: 'ctr' }),
            textParagraph([textRun('Background: linear gradient, three stops, 45°', { sz: 18, color: 'FFFFFF' })], { align: 'ctr' }),
          ],
        }),
      ],
    },
  ];
}

// ===== PPTX: CHARTS =====

const CHART_QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];
const BAR_CHART_SERIES = [
  { name: '2023', color: 'accent1', values: [120, 135, 128, 150] },
  { name: '2024', color: 'accent2', values: [132, 148, 141, 167] },
];
const PIE_CHART_SLICES = [
  { name: 'Desktop', color: 'accent1', value: 46 },
  { name: 'Mobile', color: 'accent2', value: 38 },
  { name: 'Tablet', color: 'accent6', value: 12 },
  { name: 'Other', color: 'accent3', value: 4 },
];
const CHART_AXIS_IDS = [500000001, 500000002];

/** Cached string reference (series name or categories) into the chart's notional data sheet. */
function chartStrRef(formula, values) {
  const points = values.map((value, i) => `<c:pt idx="${i}"><c:v>${xmlEscape(value)}</c:v></c:pt>`).join('');
  return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

/** Cached number reference (series values). */
function chartNumRef(formula, values) {
  const points = values.map((value, i) => `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`).join('');
  return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

/** Text properties for chart titles, axes, legends and data labels. */
function chartTextProps(sz, color = { scheme: 'tx1', lumMod: 65000, lumOff: 35000 }, bold = false) {
  return (
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${centiPoints(sz)}" b="${bold ? 1 : 0}">` +
    `<a:solidFill>${drawingColor(color)}</a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr>` +
    '<a:endParaRPr lang="en-US"/></a:p></c:txPr>'
  );
}

/** A chart part: title, plot area, legend at the bottom, no frame, no embedded workbook (cached values only). */
function chartSpace(title, plotArea) {
  const titleXml =
    '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1600" b="0"/></a:pPr>' +
    `<a:r><a:rPr lang="en-US" sz="1600" b="0"/><a:t>${xmlEscape(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  return (
    XML_DECLARATION +
    `<c:chartSpace xmlns:c="${PPTX_NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>' +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plotArea}</c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${chartTextProps(12)}</c:legend>` +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${chartTextProps(12)}` +
    '</c:chartSpace>'
  );
}

/** chart1.xml: clustered column chart, two series (sheet columns B, C) over four quarters (column A). */
function pptxBarChart() {
  const last = CHART_QUARTERS.length + 1;
  const series = BAR_CHART_SERIES.map((s, i) => {
    const column = columnLetter(i + 2);
    return (
      `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${chartStrRef(`Sheet1!$${column}$1`, [s.name])}</c:tx>` +
      `<c:spPr><a:solidFill>${drawingColor(s.color)}</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
      `<c:cat>${chartStrRef(`Sheet1!$A$2:$A$${last}`, CHART_QUARTERS)}</c:cat>` +
      `<c:val>${chartNumRef(`Sheet1!$${column}$2:$${column}$${last}`, s.values)}</c:val></c:ser>`
    );
  }).join('');
  const [catId, valId] = CHART_AXIS_IDS;
  const axisLine = color => `<c:spPr><a:ln w="9525">${color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : '<a:noFill/>'}</a:ln></c:spPr>`;
  const plotArea =
    `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}` +
    `<c:gapWidth val="150"/><c:overlap val="-20"/><c:axId val="${catId}"/><c:axId val="${valId}"/></c:barChart>` +
    `<c:catAx><c:axId val="${catId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    `${axisLine('BFBFBF')}${chartTextProps(12)}<c:crossAx val="${valId}"/><c:crosses val="autoZero"/><c:auto val="1"/>` +
    '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    `<c:valAx><c:axId val="${valId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
    `<c:majorGridlines>${axisLine('D9D9D9')}</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/>` +
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    `${axisLine()}${chartTextProps(12)}<c:crossAx val="${catId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  return chartSpace('Units sold per quarter', plotArea);
}

/** chart2.xml: pie chart, one series with a color per slice (c:dPt) and percentage labels. */
function pptxPieChart() {
  const last = PIE_CHART_SLICES.length + 1;
  const points = PIE_CHART_SLICES.map(
    (slice, i) =>
      `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${drawingColor(slice.color)}</a:solidFill>` +
      '<a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></c:spPr></c:dPt>',
  ).join('');
  const labels =
    `<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${chartTextProps(12, 'FFFFFF', true)}<c:dLblPos val="ctr"/>` +
    '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/>' +
    '<c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>';
  const series =
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${chartStrRef('Sheet1!$B$1', ['Share of sessions'])}</c:tx>${points}${labels}` +
    `<c:cat>${chartStrRef(`Sheet1!$A$2:$A$${last}`, PIE_CHART_SLICES.map(slice => slice.name))}</c:cat>` +
    `<c:val>${chartNumRef(`Sheet1!$B$2:$B$${last}`, PIE_CHART_SLICES.map(slice => slice.value))}</c:val></c:ser>`;
  return chartSpace('Sessions by device (%)', `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`);
}

/** Assembles sample.pptx with fixed timestamps. */
async function generatePptx() {
  const slides = pptxSlideDefinitions();
  const layouts = pptxSlideLayouts();
  const charts = [pptxBarChart(), pptxPieChart()];
  const counts = {
    slideCount: slides.length,
    notesCount: slides.filter(slide => slide.notes).length,
    hiddenCount: slides.filter(slide => slide.hidden).length,
    chartCount: charts.length,
    layoutCount: layouts.length,
  };

  const zip = new JSZip();
  const add = (name, content) => zip.file(name, content, { date: FIXED_DATE, createFolders: false });
  add('[Content_Types].xml', pptxContentTypes(counts));
  add('_rels/.rels', pptxRels([
    { id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' },
    { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
    { id: 'rId3', type: 'extended-properties', target: 'docProps/app.xml' },
  ]));
  add('docProps/core.xml', pptxCoreProps());
  add('docProps/app.xml', pptxAppProps(counts));
  add('ppt/presentation.xml', pptxPresentation(slides.length));
  add('ppt/_rels/presentation.xml.rels', pptxPresentationRels(slides.length));
  add('ppt/presProps.xml', pptxPresProps());
  add('ppt/viewProps.xml', pptxViewProps());
  add('ppt/tableStyles.xml', pptxTableStyles());
  add('ppt/theme/theme1.xml', pptxTheme());
  add('ppt/theme/theme2.xml', pptxTheme());

  add('ppt/slideMasters/slideMaster1.xml', pptxSlideMaster(layouts.length));
  add('ppt/slideMasters/_rels/slideMaster1.xml.rels', pptxRels([
    ...layouts.map((_, i) => ({ id: `rId${i + 1}`, type: 'slideLayout', target: `../slideLayouts/slideLayout${i + 1}.xml` })),
    { id: `rId${layouts.length + 1}`, type: 'theme', target: '../theme/theme1.xml' },
  ]));
  layouts.forEach((layout, i) => {
    add(`ppt/slideLayouts/slideLayout${i + 1}.xml`, pptxSlideLayout(layout));
    add(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, pptxRels([{ id: 'rId1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]));
  });
  add('ppt/notesMasters/notesMaster1.xml', pptxNotesMaster());
  add('ppt/notesMasters/_rels/notesMaster1.xml.rels', pptxRels([{ id: 'rId1', type: 'theme', target: '../theme/theme2.xml' }]));

  let notesNumber = 0;
  slides.forEach((slide, i) => {
    const number = i + 1;
    const rels = [{ id: 'rId1', type: 'slideLayout', target: `../slideLayouts/slideLayout${slide.layout}.xml` }];
    if (slide.notes) {
      notesNumber++;
      rels.push({ id: 'rId2', type: 'notesSlide', target: `../notesSlides/notesSlide${notesNumber}.xml` });
      add(`ppt/notesSlides/notesSlide${notesNumber}.xml`, pptxNotesSlide(slide.notes));
      add(`ppt/notesSlides/_rels/notesSlide${notesNumber}.xml.rels`, pptxRels([
        { id: 'rId1', type: 'notesMaster', target: '../notesMasters/notesMaster1.xml' },
        { id: 'rId2', type: 'slide', target: `../slides/slide${number}.xml` },
      ]));
    }
    const ids = {};
    for (const [name, rel] of Object.entries(slide.rels || {})) {
      ids[name] = `rId${rels.length + 1}`;
      rels.push({ id: ids[name], ...rel });
    }
    add(`ppt/slides/slide${number}.xml`, pptxSlide(slide, ids));
    add(`ppt/slides/_rels/slide${number}.xml.rels`, pptxRels(rels));
  });

  charts.forEach((chart, i) => add(`ppt/charts/chart${i + 1}.xml`, chart));
  add('ppt/media/image1.png', makeGradientPng(240, 160));
  add('ppt/media/image2.png', makeBadgePng(128));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, platform: 'DOS' });
}

// ===== MAIN =====

async function main() {
  let opts;
  try {
    opts = parseArgs(process.argv.slice(2));
  } catch (error) {
    if (error instanceof UsageError) {
      console.error(`${error.message}\n\n${USAGE}`);
      process.exitCode = 2;
      return;
    }
    throw error;
  }
  if (opts.help) {
    console.log(USAGE);
    return;
  }

  fs.mkdirSync(IMAGES_DIR, { recursive: true });
  const started = nowMs();
  const results = [];
  /** Builds one file in memory, writes it if changed and records the outcome. */
  const step = async (relativePath, build) => {
    const t0 = nowMs();
    const data = await build();
    const file = path.join(SAMPLES_DIR, relativePath);
    const status = writeIfChanged(file, data, opts.force);
    results.push({ file: relativePath, size: data.length, status, ms: nowMs() - t0 });
  };

  const images = {
    gradient: makeGradientPng(120, 80),
    chart: makeChartPng(240, 144),
  };
  await step('images/sample.png', () => makeBadgePng(64));
  await step('styled.xlsx', () => generateStyledXlsx(images));
  await step('sample.csv', buildSampleCsv);
  await step('sample-semicolon.csv', buildSemicolonCsv);
  await step('sample.tsv', buildSampleTsv);
  await step('sample.psv', buildSamplePsv);
  await step('sample.ssv', buildSampleSsv);
  await step('sample-space.ssv', buildSpaceSsv);
  await step('sample.docx', () => generateDocx(images));
  await step('sample.pdf', generatePdf);
  await step('sample.pptx', generatePptx);

  const largeFile = path.join(SAMPLES_DIR, 'large.xlsx');
  if (opts.skipLarge) {
    results.push({ file: 'large.xlsx', size: fs.existsSync(largeFile) ? fs.statSync(largeFile).size : 0, status: 'skipped (--skip-large)', ms: 0 });
  } else if (fs.existsSync(largeFile) && !opts.force) {
    results.push({ file: 'large.xlsx', size: fs.statSync(largeFile).size, status: 'exists (use --force to rebuild)', ms: 0 });
  } else {
    const existed = fs.existsSync(largeFile);
    const cells = opts.largeRows * opts.largeCols;
    console.log(`large.xlsx: streaming ${opts.largeRows.toLocaleString('en-US')} rows x ${opts.largeCols} columns (${cells.toLocaleString('en-US')} cells)...`);
    const t0 = nowMs();
    await generateLargeXlsx(largeFile, opts.largeRows, opts.largeCols, (done, total) => {
      const elapsed = (nowMs() - t0) / 1000;
      console.log(`  ${String(Math.round((done / total) * 100)).padStart(3)}%  ${done.toLocaleString('en-US')} rows  ${elapsed.toFixed(1)} s`);
    });
    const ms = nowMs() - t0;
    console.log(`large.xlsx: done in ${formatMs(ms)} (${Math.round(cells / (ms / 1000)).toLocaleString('en-US')} cells/s)`);
    results.push({ file: 'large.xlsx', size: fs.statSync(largeFile).size, status: existed ? 'updated' : 'created', ms });
  }

  const width = Math.max(...results.map(result => result.file.length));
  console.log(`\nSamples in ${path.relative(process.cwd(), SAMPLES_DIR) || SAMPLES_DIR}:`);
  for (const result of results) {
    console.log(
      `  ${result.file.padEnd(width)}  ${formatBytes(result.size).padStart(9)}  ${formatMs(result.ms).padStart(8)}  ${result.status}`,
    );
  }
  console.log(`Total ${formatMs(nowMs() - started)}`);
}

main().catch(error => {
  console.error(error && error.stack ? error.stack : error);
  process.exitCode = 1;
});
