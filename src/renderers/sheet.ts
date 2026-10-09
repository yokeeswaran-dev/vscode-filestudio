// Spreadsheet model: xlsx/csv <-> JSON model sent to the webview.
//
// The extension host keeps the ExcelJS workbook (or parsed CSV) in memory and
// the webview only ever sees a WorkbookMeta (sheet structure + style table)
// plus windows of rows that it requests while scrolling. This keeps memory and
// postMessage payloads bounded for very large sheets.
//
// Conventions used by every type below:
//   - row / column indexes are 0-based (Excel row 1 = r 0, column A = c 0)
//   - sizes are CSS pixels at 100% zoom
//   - colors are '#RRGGBB' (alpha dropped); theme/indexed colors are resolved here

// ===== MODEL TYPES (shared with media/viewer.js — keep in sync) =====

/** Inclusive cell range. */
export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface FontStyle {
  name?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean;
  color?: string;
  vertAlign?: 'superscript' | 'subscript';
}

export type BorderLineStyle =
  | 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';

export interface BorderEdge {
  style: BorderLineStyle;
  color?: string;
}

export interface CellStyle {
  font?: FontStyle;
  /** Background color (solid fill, or the foreground color of a pattern fill). */
  fill?: string;
  /** Non-solid pattern fills: pattern name + colors, rendered approximately. */
  pattern?: { type: string; fg?: string; bg?: string };
  /** Gradient fills: rendered as a CSS linear-gradient. */
  gradient?: { angle: number; stops: { position: number; color: string }[] };
  border?: {
    top?: BorderEdge;
    right?: BorderEdge;
    bottom?: BorderEdge;
    left?: BorderEdge;
    /** Diagonal lines (rendered with an SVG background when present). */
    diagonal?: BorderEdge & { up?: boolean; down?: boolean };
  };
  align?: {
    h?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
    v?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
    wrap?: boolean;
    shrink?: boolean;
    /** Indent level (1 level = 3 characters, ~9px... rendered as level * 9px). */
    indent?: number;
    /** Degrees, -90..90, or 'vertical' (stacked text, Excel's 255). */
    rotation?: number | 'vertical';
  };
  /** Excel number format code, e.g. '0.00%' ('General' when absent). */
  numFmt?: string;
}

/** One run of a rich-text cell. */
export interface RichRun {
  text: string;
  font?: FontStyle;
}

export type CellType =
  | 'n' // number
  | 's' // string
  | 'b' // boolean
  | 'd' // date (v is the Excel serial number)
  | 'e' // error, v is the error text e.g. '#DIV/0!'
  | 'z'; // empty cell that only carries a style

export interface CellData {
  /** Column index. */
  c: number;
  t: CellType;
  /** Raw value. Dates are Excel serial numbers (respecting the 1904 flag). */
  v?: string | number | boolean | null;
  /** Display text after applying the number format (what the grid shows). */
  w?: string;
  /** Formula text WITHOUT the leading '=' (shared formulas expanded). */
  f?: string;
  /**
   * `f` is a legacy array formula (entered with Ctrl+Shift+Enter): the formula bar shows it as {=f}. Every cell of a
   * multi-cell array range carries the range's formula with this flag. Dynamic-array formulas (Excel 365, spilling)
   * are shown without braces and do not set it.
   */
  array?: boolean;
  /** Index into WorkbookMeta.styles. Omitted = 0 (default style). */
  s?: number;
  /** Rich text runs (w/v still hold the plain text). */
  rt?: RichRun[];
  /** The number format repeats a character to fill the cell (accounting formats): how to lay `w` out. */
  fill?: FillLayout;
  /** Hyperlink target (external URL, or '#Sheet2!A1' for internal links). */
  link?: string;
  /** Plain text of the cell comment / note. */
  note?: string;
  /**
   * @deprecated No longer sent: per-cell copies made every reply huge (a whole-column
   * validation repeated its list in every cell). Lists are in SheetMeta.validations.
   */
  list?: string[];
}

/**
 * Layout of a value whose number format repeats a character to fill the cell (`*x`: the accounting format
 * `_("$"* #,##0.00_)`, the text format `@*-`). Like Excel, the grid shows `w.slice(0, at)` at the left edge and
 * `w.slice(at)` at the right edge, with `char` repeated in between; `w` itself holds no fill characters (copy, stats).
 */
export interface FillLayout {
  /** Position in `w` where the fill goes. */
  at: number;
  /** The character repeated; absent = a space (the usual accounting fill). */
  char?: string;
  /** '_x' paddings: [position in `w` (a space there), x] - the blank is as wide as x, e.g. ')' in `_)`. */
  pads?: [number, string][];
}

export interface RowData {
  r: number;
  /** Only non-empty (value or non-default style) cells, ascending by c. */
  cells: CellData[];
}

export interface ColInfo {
  /** Width in px. */
  w?: number;
  hidden?: boolean;
  /** Column default style index. */
  s?: number;
}

export interface RowInfo {
  /** Height in px. */
  h?: number;
  hidden?: boolean;
  /** Row default style index. */
  s?: number;
}

export interface CfValueObject {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax';
  value?: number | string;
}

/** Conditional-format rule, normalized from ExcelJS. Evaluated in the webview (phase 3). */
export interface CfRule {
  ranges: Range[];
  priority: number;
  type:
    | 'cellIs' | 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
    | 'top10' | 'aboveAverage' | 'duplicateValues' | 'uniqueValues'
    | 'colorScale' | 'dataBar' | 'iconSet' | 'expression' | 'containsBlanks' | 'notContainsBlanks'
    | 'containsErrors' | 'notContainsErrors' | 'timePeriod' | string;
  operator?: string;
  formulae?: string[];
  text?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  stopIfTrue?: boolean;
  /** Differential style applied when the rule matches (colors resolved). */
  style?: CellStyle;
  colorScale?: { cfvo: CfValueObject[]; colors: string[] };
  dataBar?: { cfvo: CfValueObject[]; color: string; gradient?: boolean; showValue?: boolean };
  iconSet?: { name: string; cfvo: CfValueObject[]; reverse?: boolean; showValue?: boolean };
}

export interface ImageAnchor {
  /** Top-left anchor cell plus offset in px inside that cell. */
  from: { r: number; c: number; dx: number; dy: number };
  /** Bottom-right anchor (two-cell anchor), if known. */
  to?: { r: number; c: number; dx: number; dy: number };
  /** Size in px (one-cell anchor), if known. */
  ext?: { w: number; h: number };
  /** data: URI of the image. */
  src: string;
}

/** A list (dropdown) data validation: the cells it covers and its items. */
export interface ListValidation {
  /** Cells with the validation (a whole column / row runs to r1 1048575 / c1 16383). */
  ranges: Range[];
  /** Display texts of the items (a literal "a,b,c" list or the cells of the source range, capped). */
  values: string[];
}

export interface SheetMeta {
  /** Position in WorkbookMeta.sheets; the identifier used in messages. */
  index: number;
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  tabColor?: string;
  /** Number of rows / columns in the used range (at least 1). */
  rowCount: number;
  colCount: number;
  defaultColWidth: number;
  defaultRowHeight: number;
  /** Sparse per-column / per-row info keyed by index. */
  cols: Record<number, ColInfo>;
  rows: Record<number, RowInfo>;
  merges: Range[];
  /** Frozen rows at the top / columns at the left (0 = none). */
  frozen: { rows: number; cols: number };
  showGridLines: boolean;
  /** Percent, 100 = normal. */
  zoom: number;
  rightToLeft?: boolean;
  conditionalFormats: CfRule[];
  images: ImageAnchor[];
  autoFilter?: Range;
  /** List data validations, one entry per distinct list (absent when the sheet has none). */
  validations?: ListValidation[];
}

export interface WorkbookMeta {
  sheets: SheetMeta[];
  /** Deduplicated style table; styles[0] is the workbook default style. */
  styles: CellStyle[];
  /**
   * Index of the sheet that was active when the file was saved. It can be a hidden sheet (Excel opens the file on it,
   * unhiding it; the webview shows it revealed), never a very hidden one (then the first visible sheet, or the first
   * sheet when none is visible).
   */
  activeSheet: number;
  /** Workbook default font (Normal style), used for the grid base font. */
  defaultFont: { name: string; size: number };
  date1904: boolean;
}

/** Aggregates for the status bar. */
export interface SelectionStats {
  /** Non-empty cells. */
  count: number;
  /** Numeric cells. */
  numCount: number;
  sum: number;
  avg?: number;
  min?: number;
  max?: number;
  /**
   * Cells holding an error value (#DIV/0!, #N/A ...); they are counted in `count` too. Like Excel, the status bar then
   * shows only Count. Absent when there is none (always for CSV, whose fields are text).
   */
  errors?: number;
  /**
   * Sum / Average / Min / Max as Excel's status bar shows them: in the number format of the first numeric cell in the
   * order of `ranges` (range by range, each row by row; Excel 16 does the same: B9 0.0% then B10 #,##0.00 shows both
   * as percentages, 'B10,B9' as numbers), e.g. a date, a percentage or a currency. Absent when that cell has the
   * General format (the viewer formats plain numbers itself) and for CSV.
   */
  text?: { sum: string; avg?: string; min?: string; max?: string };
  /**
   * True when the selection holds more cells than getStats reads (its work cap): the values above cover only the
   * first `scanned` cells, row by row, and are lower bounds / partial aggregates. Absent when complete.
   */
  partial?: boolean;
  /** With `partial`: how many cells (present in the sheet) were read before stopping. */
  scanned?: number;
}

/** Result of parseCsv(). */
export interface CsvModel {
  delimiter: string;
  /** '\r\n' | '\n' | '\r' — the dominant line ending of the source. */
  newline: string;
  hasBom: boolean;
  /** Rows of raw field text (unquoted). */
  rows: string[][];
  /** True when the file ends with a line break. */
  trailingNewline: boolean;
  /**
   * (v0.2.0, additive) Space-separated text (delimiter ' '): a run of spaces is one separator and leading /
   * trailing spaces are not fields. Absent for every other delimiter. serializeCsv writes edited records with
   * single spaces (empty fields as "").
   */
  collapseSpaces?: boolean;
  /**
   * (v0.1.1, additive) Excel's delimiter directive when the text starts with one (after the BOM): the line as
   * written, e.g. 'sep=;\r\n'. It sets `delimiter` (unless parseCsv was given one) and is not a row; serializeCsv
   * writes it back. Absent = no directive.
   */
  sepLine?: string;
}

// ===== IMPLEMENTATION =====
//
// Public API (phase-1 contract): loadWorkbook, GridSource, WorkbookModel /
// workbookToModel, CsvGridModel, parseCsv / serializeCsv, ThemePalette /
// parseThemeXml / resolveColor, detectLossyFeatures; v0.2.0 adds
// detectSsvDelimiter. Pure Node code - this module must never import `vscode`
// (it is unit-tested in plain Node).
//
// Sections:
//   LIBRARIES          third-party imports, local typing for untyped ssf
//   EXCELJS INTERNALS  the (undocumented) slice of ExcelJS' object model we read
//   CONSTANTS          units, limits and builtin tables
//   COLORS             theme palette, tint maths, indexed palette, resolveColor
//   ADDRESSES          A1 helpers, range parsing, shared-formula translation
//   NUMBER FORMATS     ssf wrapper with fast paths and caches
//   STYLES             ExcelJS style -> CellStyle conversion + dedupe table
//   LOADING            loadWorkbook (raw styles.xml side channel, internal links)
//   GRID SOURCE        interface implemented by both grid models
//   WORKBOOK MODEL     WorkbookModel / workbookToModel (meta, rows, stats)
//   SELECTION STATS    range-union walker + accumulator (xlsx and csv)
//   CSV                parseCsv / serializeCsv / detectSsvDelimiter / CsvGridModel
//   LOSSY FEATURES     detectLossyFeatures

// ===== LIBRARIES =====

import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import Papa from 'papaparse';

/** The part of ssf (SheetJS number formatter, untyped CommonJS) used here. */
interface SsfLib {
  format(fmt: string, value: unknown, opts?: { date1904?: boolean }): string;
  is_date(fmt: string): boolean;
}
const SSF = require('ssf') as SsfLib;

/** ExcelJS' <dataValidations> parser (an internal module; see patchValidationParser). */
interface XDataValidationsXformClass {
  prototype: { parseClose(this: { model?: object; _address?: unknown; _dataValidation?: XDataValidation }, name: string): boolean };
}
const DataValidationsXform = require('exceljs/lib/xlsx/xform/sheet/data-validations-xform') as XDataValidationsXformClass | undefined;

/** ExcelJS' Worksheet class (an internal module; see patchSheetNameCheck). */
interface XWorksheetClass {
  prototype: object;
}
const WorksheetClass = require('exceljs/lib/doc/worksheet') as XWorksheetClass | undefined;

/** ExcelJS' parser of one <comment> and its <t> parser (internal modules; see patchCommentParser). */
interface XParser {
  parser?: XParser;
  model?: unknown;
  parseOpen(node: { name: string }): boolean;
  parseText(text: string): void;
  parseClose(name: string): boolean;
}
interface XCommentParser extends XParser {
  model?: { note?: { texts?: XRichRun[] } };
  /** Set while inside <rPh> (phonetic text, not shown). */
  fvPhonetic?: boolean;
  /** The <t> parser of a note written without runs. */
  fvPlainText?: XParser;
}
const CommentXform = require('exceljs/lib/xlsx/xform/comment/comment-xform') as { prototype: XCommentParser } | undefined;
const TextXform = require('exceljs/lib/xlsx/xform/strings/text-xform') as (new () => XParser) | undefined;

/** ExcelJS' parser of one <c> cell (an internal module; see patchCellParser). */
interface XCellParser {
  parser?: unknown;
  model?: XCellLoadModel;
  /** The <c t="..."> attribute of the cell being parsed ('s', 'str', 'inlineStr', ...). */
  t?: string;
  parseOpen(node: { name: string; attributes?: Record<string, string | undefined> }): boolean;
  parseClose(name: string): boolean;
}
const CellXform = require('exceljs/lib/xlsx/xform/sheet/cell-xform') as { prototype: XCellParser } | undefined;

// ===== EXCELJS INTERNALS =====
// ExcelJS 4.4 keeps its document in plain objects that are mostly absent from
// its public typings. Every property below was checked against
// node_modules/exceljs/lib: doc/cell.js (Cell._value, Value#type/model,
// Cell._comment), doc/row.js (_cells, height, hidden, style), doc/column.js,
// doc/worksheet.js (_rows, _columns, _merges, views, properties, getImages,
// dataValidations.model, conditionalFormattings), doc/anchor.js (native*),
// doc/workbook.js (_themes, media, definedNames), xlsx/xlsx.js (reconcile) and
// the xlsx/xform/** parsers that define the shape of style / rule models
// (sheet/data-validations-xform.js, comment/comment-xform.js and the
// doc/worksheet.js name setter are also wrapped while loading).

/** {argb} | {theme, tint} | {indexed}; `undefined` = automatic. */
interface XColor {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
  auto?: boolean;
}

interface XFont {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** `true` for <u/>, else 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none'. */
  underline?: boolean | string;
  strike?: boolean;
  color?: XColor;
  vertAlign?: string;
}

interface XFill {
  type?: string; // 'pattern' | 'gradient'
  pattern?: string;
  fgColor?: XColor;
  bgColor?: XColor;
  gradient?: string; // 'angle' | 'path'
  degree?: number;
  stops?: { position?: number; color?: XColor }[];
}

interface XBorderEdge {
  style?: string;
  color?: XColor;
}

interface XBorder {
  top?: XBorderEdge;
  right?: XBorderEdge;
  bottom?: XBorderEdge;
  left?: XBorderEdge;
  diagonal?: XBorderEdge & { up?: boolean; down?: boolean };
  /** Border-wide default colour (style models built through the ExcelJS API). */
  color?: XColor;
}

interface XAlignment {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number | string;
}

interface XStyle {
  /** Cell styles carry a string; dxf styles carry {id, formatCode}. */
  numFmt?: string | { formatCode?: string };
  font?: XFont;
  fill?: XFill;
  border?: XBorder;
  alignment?: XAlignment;
}

interface XRichRun {
  text?: unknown;
  font?: XFont;
}

interface XValueModel {
  address?: string;
  value?: unknown;
  text?: unknown;
  hyperlink?: string;
  formula?: string;
  sharedFormula?: string;
  result?: unknown;
  /** 'array' for an array formula (with `ref`, its range), 'shared' for a shared-formula master. */
  shareType?: string;
  ref?: string;
}

/** A cell as parsed from sheet XML, before ExcelJS turns it into a Cell (see patchCellParser). */
interface XCellLoadModel {
  address?: string;
  shareType?: string;
  ref?: string;
  /** The <c cm="n"> attribute (ExcelJS drops it): 1-based cell-metadata record, e.g. a dynamic array. */
  cm?: string;
  /** Cell value (a string for t="str" / t="inlineStr", {richText} for inline runs). */
  value?: unknown;
  /** Cached result of a formula (a string for t="str"). */
  result?: unknown;
}

/** Cell value holder (NumberValue, FormulaValue, ...): `type` is a getter. */
interface XValue {
  readonly type: number;
  model: XValueModel;
}

interface XCell {
  _value: XValue;
  style?: XStyle;
  _comment?: { note?: string | { texts?: XRichRun[] } };
}

interface XRow {
  _cells: (XCell | undefined)[];
  style?: XStyle;
  height?: number;
  hidden?: boolean;
}

interface XColumn {
  width?: number;
  hidden?: boolean;
  style?: XStyle;
  readonly isCustomWidth?: boolean;
}

interface XAnchor {
  nativeCol?: number;
  nativeColOff?: number;
  nativeRow?: number;
  nativeRowOff?: number;
}

interface XImage {
  imageId?: number | string;
  range?: { tl?: XAnchor; br?: XAnchor; ext?: { width?: number; height?: number } };
}

interface XMedia {
  type?: string;
  extension?: string;
  buffer?: Uint8Array;
  base64?: string;
}

interface XView {
  state?: string;
  xSplit?: number;
  ySplit?: number;
  showGridLines?: boolean;
  zoomScale?: number;
  rightToLeft?: boolean;
}

interface XDataValidation {
  type?: string;
  operator?: string;
  formulae?: unknown[];
}

interface XCfvo {
  type?: string;
  value?: unknown;
}

interface XCfRule {
  type?: string;
  operator?: string;
  priority?: number;
  formulae?: unknown[];
  text?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  stopIfTrue?: boolean;
  timePeriod?: string;
  style?: XStyle;
  cfvo?: XCfvo[];
  color?: XColor | XColor[];
  gradient?: boolean;
  showValue?: boolean;
  iconSet?: string;
  reverse?: boolean;
}

interface XWorksheet {
  id?: number;
  name: string;
  state?: string;
  orderNo?: number;
  _rows: (XRow | undefined)[];
  _columns?: (XColumn | undefined)[] | null;
  _merges?: Record<string, { model?: { top?: number; left?: number; bottom?: number; right?: number } } | undefined>;
  properties?: { defaultRowHeight?: number; defaultColWidth?: number; tabColor?: XColor };
  views?: XView[];
  autoFilter?: unknown;
  dataValidations?: { model?: Record<string, XDataValidation | undefined> };
  conditionalFormattings?: { ref?: string; rules?: XCfRule[] }[];
  getImages?(): XImage[];
}

interface XWorkbook {
  readonly worksheets: XWorksheet[];
  views?: { activeTab?: number }[];
  properties?: { date1904?: boolean };
  _themes?: Record<string, string>;
  media?: XMedia[];
  definedNames?: { getRanges(name: string): { ranges?: string[] } };
}

/** ExcelJS Enums.ValueType. */
const VT = {
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
} as const;

// ===== CONSTANTS =====

const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
/** Longest sheet name Excel allows. */
const SHEET_NAME_MAX = 31;
/** Maximum digit width (px) of the default font (Calibri 11): Excel's column-width unit. */
const MDW = 7;
const DEFAULT_COL_WIDTH_PX = 64;
const DEFAULT_ROW_HEIGHT_PT = 15;
const EMU_PER_PX = 9525;
const MS_PER_DAY = 86_400_000;
/** getStats never visits more cells than this. */
const STATS_CELL_CAP = 2_000_000;
/** 10000-01-01: the first serial past the last date Excel can display (1900 system). */
const DATE_LIMIT_SERIAL = 2_958_466;
/** Serial of 1904-01-01 in the 1900 date system (1904 serial = 1900 serial - this). */
const DATE1904_OFFSET = 1462;
/** What Excel shows for a negative or too large date / time. */
const DATE_OVERFLOW_TEXT = '########';
/** Formatted date / time strings cached per format (the cache is reset when full). */
const DATE_CACHE_LIMIT = 8192;
/** Data-validation list values resolved from a range are capped to this many entries. */
const LIST_VALUES_CAP = 1000;
/** Validation keys read per sheet from workbooks built in memory (loaded ones keep their sqref). */
const VALIDATION_CELLS_CAP = 100_000;
/** A loaded validation range bigger than this stays one ExcelJS model key instead of one per cell. */
const VALIDATION_EXPAND_CAP = 10_000;
/** A hyperlink with a range ref is expanded to at most this many cells. */
const HYPERLINK_EXPAND_CAP = 10_000;
/** Column default styles are sent for columns up to colCount + this margin. */
const COL_STYLE_MARGIN = 256;

const ptToPx = (pt: number): number => Math.round((pt * 4) / 3);
const widthToPx = (chars: number): number => Math.round(chars * MDW);

/** Excel's builtin number formats (en-US), by numFmtId. */
const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

/** ExcelJS' own builtin-format table has a few codes that differ from Excel's display. */
const EXCELJS_FORMAT_FIXES: Record<string, string> = {
  'mm-dd-yy': BUILTIN_NUMFMTS[14],
  'm/d/yy "h":mm': BUILTIN_NUMFMTS[22],
  'mmss.0': BUILTIN_NUMFMTS[47],
};

const BORDER_STYLES: ReadonlySet<string> = new Set<BorderLineStyle>([
  'thin', 'medium', 'thick', 'dotted', 'dashed', 'double', 'hair',
  'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
]);

const H_ALIGN: ReadonlySet<string> = new Set(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']);
const V_ALIGN: ReadonlySet<string> = new Set(['top', 'middle', 'bottom', 'distributed', 'justify']);

/** Approximate ink coverage of Excel's pattern fills (fraction of the foreground colour). */
const PATTERN_DENSITY: Record<string, number> = {
  gray0625: 0.0625,
  gray125: 0.125,
  lightGray: 0.25,
  mediumGray: 0.5,
  darkGray: 0.75,
  lightHorizontal: 0.25,
  lightVertical: 0.25,
  lightDown: 0.25,
  lightUp: 0.25,
  lightGrid: 0.4375,
  lightTrellis: 0.375,
  darkHorizontal: 0.5,
  darkVertical: 0.5,
  darkDown: 0.5,
  darkUp: 0.5,
  darkGrid: 0.75,
  darkTrellis: 0.75,
};

// ===== COLORS =====

export interface ThemePalette {
  /** 12 entries: lt1 dk1 lt2 dk2 accent1..6 hlink folHlink in Excel index order ('#RRGGBB'). */
  colors: string[];
  /** Workbook-specific indexed palette (styles.xml <indexedColors>), when present. */
  indexed?: string[];
}

/** Theme slots in Excel's colour-index order (the theme XML lists dk1/lt1 and dk2/lt2 swapped). */
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** Office 2013-2022 theme, used when a workbook carries no theme part. */
const DEFAULT_THEME_COLORS = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72',
];

/** Excel's legacy 64-entry indexed palette (indexes 64 and up are SYSTEM_COLORS). */
const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
].map((hex) => `#${hex}`);

/**
 * Indexes 64..81 are Windows system colours (64 = window text, 65 = window, then
 * button face / shadow / highlight / info background ...). Values are what Excel
 * paints with the default Windows 10/11 colours (checked against Excel 16).
 */
const SYSTEM_COLORS = [
  '000000', 'FFFFFF', '646464', 'F0F0F0', '000000', 'FFFFFF', 'A0A0A0', '0078D7', '000000',
  'C8C8C8', '373737', 'FFFFFF', '646464', '000000', 'FFFFFF', '000000', 'FFFFE1', '000000',
].map((hex) => `#${hex}`);

/** Parses the <a:clrScheme> of a theme part. Missing slots keep the Office defaults. */
export function parseThemeXml(xml: string | undefined): ThemePalette {
  const colors = DEFAULT_THEME_COLORS.slice();
  const scheme = xml ? xmlSection(xml, 'clrScheme') : undefined;
  if (scheme === undefined) {
    return { colors };
  }
  THEME_SLOTS.forEach((slot, index) => {
    const body = xmlSection(scheme, slot);
    if (body === undefined) {
      return;
    }
    const srgb = /<(?:\w+:)?srgbClr\b[^>]*?\bval\s*=\s*["']([0-9A-Fa-f]{6})["']/.exec(body);
    if (srgb) {
      colors[index] = `#${srgb[1].toUpperCase()}`;
      return;
    }
    const sys = /<(?:\w+:)?sysClr\b([^>]*)>/.exec(body);
    if (sys) {
      const attrs = parseXmlAttributes(sys[1]);
      const last = attrs.lastClr && /^[0-9A-Fa-f]{6}$/.test(attrs.lastClr) ? attrs.lastClr : undefined;
      if (last) {
        colors[index] = `#${last.toUpperCase()}`;
      } else if (attrs.val === 'windowText') {
        colors[index] = '#000000';
      } else if (attrs.val === 'window') {
        colors[index] = '#FFFFFF';
      }
    }
  });
  return { colors };
}

/**
 * Theme + tint / indexed / argb colour -> '#RRGGBB' (alpha dropped). Returns
 * undefined for automatic or unknown colours. Exported for unit tests.
 */
export function resolveColor(color: unknown, theme: ThemePalette): string | undefined {
  if (!color || typeof color !== 'object') {
    return undefined;
  }
  const c = color as XColor & { rgb?: unknown };
  let hex: string | undefined;
  if (typeof c.argb === 'string') {
    hex = normalizeHex(c.argb);
  } else if (typeof c.rgb === 'string') {
    hex = normalizeHex(c.rgb);
  } else if (typeof c.theme === 'number') {
    const value = theme.colors[c.theme];
    hex = value === undefined ? undefined : normalizeHex(value);
  } else if (typeof c.indexed === 'number') {
    hex = indexedColor(c.indexed, theme.indexed);
  }
  if (!hex) {
    return undefined;
  }
  return typeof c.tint === 'number' && c.tint !== 0 && Number.isFinite(c.tint) ? applyTint(hex, c.tint) : hex;
}

/** 'FFRRGGBB' | 'RRGGBB' | '#RRGGBB' -> '#RRGGBB'. */
function normalizeHex(value: string): string | undefined {
  let hex = value.trim().replace(/^#/, '');
  if (hex.length === 8) {
    hex = hex.slice(2);
  }
  return /^[0-9A-Fa-f]{6}$/.test(hex) ? `#${hex.toUpperCase()}` : undefined;
}

function indexedColor(index: number, custom: string[] | undefined): string | undefined {
  if (custom && index >= 0 && index < custom.length && custom[index]) {
    return custom[index];
  }
  if (index >= 0 && index < INDEXED_COLORS.length) {
    return INDEXED_COLORS[index];
  }
  return SYSTEM_COLORS[index - INDEXED_COLORS.length];
}

/** Scale of Windows' integer HLS (ColorRGBToHLS / ColorHLSToRGB), which Excel's tint maths runs in. */
const HLSMAX = 240;

/**
 * ECMA-376 tint, computed the way Excel does it: the colour is converted to
 * Windows' integer HLS (0..240), the luminance moves - tint < 0: L * (1 + tint);
 * tint > 0: L * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint)), every product
 * truncated to an integer - and the result is converted back with integer
 * maths. This reproduces all of 720 theme + tint fills resolved by real Excel 16
 * (exact tints such as 0.4 and the k/32767 values Excel writes, 0.39997558...).
 */
function applyTint(hex: string, tint: number): string {
  const t = Math.max(-1, Math.min(1, tint));
  const [h, l, s] = rgbToHls(parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16));
  // The epsilon keeps float noise (240 * 0.95 = 227.99999999999997) on the right side of the truncation.
  const floor = (x: number): number => Math.floor(x + 1e-9);
  const lum = t < 0 ? floor(l * (1 + t)) : floor(l * (1 - t)) + HLSMAX - floor(HLSMAX * (1 - t));
  const [r, g, b] = hlsToRgb(h, Math.max(0, Math.min(HLSMAX, lum)), s);
  return `#${toHex2(r / 255)}${toHex2(g / 255)}${toHex2(b / 255)}`;
}

/** Integer division as in C (all operands here are non-negative). */
const idiv = (a: number, b: number): number => Math.trunc(a / b);

/** Windows ColorRGBToHLS: 0..255 channels -> [hue, luminance, saturation], each 0..240. */
function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = idiv((max + min) * HLSMAX + 255, 510);
  if (max === min) {
    return [idiv(HLSMAX * 2, 3), l, 0]; // achromatic: the hue Windows reports
  }
  const d = max - min;
  const s = l <= HLSMAX / 2 ? idiv(idiv(max + min, 2) + d * HLSMAX, max + min) : idiv(idiv(510 - max - min, 2) + d * HLSMAX, 510 - max - min);
  const norm = (x: number): number => idiv(idiv(d, 2) + (max - x) * (HLSMAX / 6), d);
  let h = r === max ? norm(b) - norm(g) : g === max ? HLSMAX / 3 + norm(r) - norm(b) : (HLSMAX * 2) / 3 + norm(g) - norm(r);
  if (h < 0) h += HLSMAX;
  else if (h >= HLSMAX) h -= HLSMAX;
  return [h, l, s];
}

/** Windows ColorHLSToRGB: [hue, luminance, saturation] (0..240) -> 0..255 channels. */
function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  const toRgb = (v: number): number => idiv(v * 255 + HLSMAX / 2, HLSMAX);
  if (!s) {
    const grey = toRgb(l);
    return [grey, grey, grey];
  }
  const mid2 = l > HLSMAX / 2 ? s + l - idiv(s * l + HLSMAX / 2, HLSMAX) : idiv((s + HLSMAX) * l + HLSMAX / 2, HLSMAX);
  const mid1 = l * 2 - mid2;
  const hue = (x: number): number => {
    let v = x > HLSMAX ? x - HLSMAX : x < 0 ? x + HLSMAX : x;
    if (v > (HLSMAX * 2) / 3) return mid1;
    if (v > HLSMAX / 2) v = (HLSMAX * 2) / 3 - v;
    else if (v > HLSMAX / 6) return mid2;
    return idiv(v * (mid2 - mid1) + HLSMAX / 12, HLSMAX / 6) + mid1;
  };
  return [toRgb(hue(h + HLSMAX / 3)), toRgb(hue(h)), toRgb(hue(h - HLSMAX / 3))];
}

function toHex2(unit: number): string {
  return Math.max(0, Math.min(255, Math.round(unit * 255))).toString(16).padStart(2, '0').toUpperCase();
}

/** Mixes `fg` over `bg` with the given coverage (0..1). */
function blendColors(fg: string, bg: string, coverage: number): string {
  const channel = (i: number): string => {
    const a = parseInt(fg.slice(i, i + 2), 16);
    const b = parseInt(bg.slice(i, i + 2), 16);
    return toHex2((a * coverage + b * (1 - coverage)) / 255);
  };
  return `#${channel(1)}${channel(3)}${channel(5)}`;
}

/** True for colours Excel shows as "Automatic" text (none, auto, indexed 64). */
function isAutomaticColor(color: XColor | undefined): boolean {
  return !color || color.auto === true || color.indexed === 64 || (color.argb === undefined && color.theme === undefined && color.indexed === undefined);
}

// ===== ADDRESSES =====

/** 'A' -> 0, 'XFD' -> 16383; -1 when not letters. */
function colToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i) & ~32; // upper-case
    if (code < 65 || code > 90) {
      return -1;
    }
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

/** 0 -> 'A', 16383 -> 'XFD'. */
function indexToCol(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

const CELL_REF_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COL_REF_RE = /^\$?([A-Za-z]{1,3})$/;
const ROW_REF_RE = /^\$?(\d{1,7})$/;

/** 'B3' / '$B$3' -> {r: 2, c: 1}. */
function decodeCell(ref: string): { r: number; c: number } | undefined {
  const m = CELL_REF_RE.exec(ref.trim());
  if (!m) {
    return undefined;
  }
  const c = colToIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  return c >= 0 && c < MAX_COLS && r >= 0 && r < MAX_ROWS ? { r, c } : undefined;
}

/** 'A1', 'A1:C5', 'A:C', '2:4' (with or without $) -> inclusive 0-based Range. */
function decodeRange(ref: string): Range | undefined {
  const parts = ref.trim().split(':');
  if (parts.length === 1) {
    const cell = decodeCell(parts[0]);
    return cell ? { r0: cell.r, c0: cell.c, r1: cell.r, c1: cell.c } : undefined;
  }
  if (parts.length !== 2) {
    return undefined;
  }
  const a = decodeCell(parts[0]);
  const b = decodeCell(parts[1]);
  if (a && b) {
    return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
  }
  const ca = COL_REF_RE.exec(parts[0].trim());
  const cb = COL_REF_RE.exec(parts[1].trim());
  if (ca && cb) {
    const c0 = colToIndex(ca[1]);
    const c1 = colToIndex(cb[1]);
    if (c0 < 0 || c1 < 0 || c0 >= MAX_COLS || c1 >= MAX_COLS) return undefined;
    return { r0: 0, c0: Math.min(c0, c1), r1: MAX_ROWS - 1, c1: Math.max(c0, c1) };
  }
  const ra = ROW_REF_RE.exec(parts[0].trim());
  const rb = ROW_REF_RE.exec(parts[1].trim());
  if (ra && rb) {
    const r0 = parseInt(ra[1], 10) - 1;
    const r1 = parseInt(rb[1], 10) - 1;
    if (r0 < 0 || r1 < 0 || r0 >= MAX_ROWS || r1 >= MAX_ROWS) return undefined;
    return { r0: Math.min(r0, r1), c0: 0, r1: Math.max(r0, r1), c1: MAX_COLS - 1 };
  }
  return undefined;
}

/** Space-separated list of refs (sqref) -> ranges. */
function parseSqref(sqref: unknown): Range[] {
  if (typeof sqref !== 'string') {
    return [];
  }
  const out: Range[] = [];
  for (const part of sqref.split(/\s+/)) {
    const range = part ? decodeRange(part) : undefined;
    if (range) {
      out.push(range);
    }
  }
  return out;
}

/** "'My Sheet'!$A$1:$B$2" -> {sheet: 'My Sheet', ref: '$A$1:$B$2'}. */
function splitSheetRef(text: string): { sheet?: string; ref: string } {
  const bang = text.lastIndexOf('!');
  if (bang < 0) {
    return { ref: text.trim() };
  }
  let sheet = text.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, ref: text.slice(bang + 1).trim() };
}

/** Quotes a sheet name for use in a reference when Excel would. */
function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !CELL_REF_RE.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** 'B2:C3' -> ['B2', 'B3', 'C2', 'C3'] (ExcelJS address style), capped. */
function expandAddresses(ref: string, cap: number): string[] {
  const range = decodeRange(ref);
  if (!range) {
    return [];
  }
  const out: string[] = [];
  for (let c = range.c0; c <= range.c1; c++) {
    const letters = indexToCol(c);
    for (let r = range.r0; r <= range.r1; r++) {
      if (out.length >= cap) {
        return out;
      }
      out.push(`${letters}${r + 1}`);
    }
  }
  return out;
}

/**
 * Formula tokens relevant to relative-reference translation. String literals,
 * quoted sheet names and [bracketed] parts (structured / external references)
 * are matched first so references inside them are left alone. The look-arounds
 * keep function names (LOG10), defined names and numbers (1E10) untouched.
 */
const FORMULA_REF_RE =
  /"(?:[^"]|"")*"|'(?:[^']|'')*'|\[[^\]]*\]|(?<![\p{L}\p{N}_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\p{L}\p{N}_(!.[])|(?<![\p{L}\p{N}_.$])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\p{L}\p{N}_(!.[])|(?<![\p{L}\p{N}_.$:])(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\p{L}\p{N}_(!.[:])/gu;

/**
 * Moves every relative reference of `formula` by (dr, dc) - how Excel derives
 * the formula of a cell that shares the formula of a master cell.
 */
function translateFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) {
    return formula;
  }
  const moveCol = (letters: string, absolute: string): string | undefined => {
    const index = colToIndex(letters);
    const moved = absolute ? index : index + dc;
    return moved >= 0 && moved < MAX_COLS ? `${absolute}${indexToCol(moved)}` : undefined;
  };
  const moveRow = (digits: string, absolute: string): string | undefined => {
    const index = parseInt(digits, 10) - 1;
    const moved = absolute ? index : index + dr;
    return moved >= 0 && moved < MAX_ROWS ? `${absolute}${moved + 1}` : undefined;
  };
  return formula.replace(
    FORMULA_REF_RE,
    (match: string, cAbs?: string, col?: string, rAbs?: string, row?: string, caAbs?: string, colA?: string, cbAbs?: string, colB?: string, raAbs?: string, rowA?: string, rbAbs?: string, rowB?: string): string => {
      if (col !== undefined && row !== undefined) {
        if (colToIndex(col) >= MAX_COLS) return match; // e.g. a name like "ABCD1" - not a reference
        const c = moveCol(col, cAbs ?? '');
        const r = moveRow(row, rAbs ?? '');
        return c && r ? c + r : '#REF!';
      }
      if (colA !== undefined && colB !== undefined) {
        if (colToIndex(colA) >= MAX_COLS || colToIndex(colB) >= MAX_COLS) return match;
        const a = moveCol(colA, caAbs ?? '');
        const b = moveCol(colB, cbAbs ?? '');
        return a && b ? `${a}:${b}` : '#REF!';
      }
      if (rowA !== undefined && rowB !== undefined) {
        const a = moveRow(rowA, raAbs ?? '');
        const b = moveRow(rowB, rbAbs ?? '');
        return a && b ? `${a}:${b}` : '#REF!';
      }
      return match;
    },
  );
}

/** Formula text for display: no leading '=', no _xlfn./_xlws./_xlpm. storage prefixes. */
function displayFormula(formula: string): string {
  const text = formula.startsWith('=') ? formula.slice(1) : formula;
  return text.includes('_xl') ? text.replace(/_xl(?:fn|ws|pm)\./g, '') : text;
}

/** String literal | '[n]Sheet name'! | [n]Sheet1! or [n]! (a workbook-level name follows). */
const EXTERNAL_REF_RE = /"(?:[^"]|"")*"|'\[(\d+)\]((?:[^']|'')*)'!|\[(\d+)\]([\p{L}\p{N}_.]*)!/gu;

/**
 * Files store references to other workbooks by position in <externalReferences>
 * ([1]Sheet1!B2, '[1]My Sheet'!B2, [1]!Name); Excel's formula bar shows the
 * workbook instead: 'C:\Data\[Book.xlsx]Sheet1'!B2, 'C:\Data\Book.xlsx'!Name
 * ('[Book.xlsx]Sheet1'!B2 while its folder is unknown).
 */
function showExternalRefs(formula: string, books: readonly (ExternalBook | undefined)[]): string {
  if (!formula.includes('[')) {
    return formula;
  }
  return formula.replace(EXTERNAL_REF_RE, (match: string, quotedIndex?: string, quotedSheet?: string, index?: string, sheet?: string): string => {
    const book = books[Number(quotedIndex ?? index) - 1]; // a string literal has neither: NaN
    if (!book) {
      return match;
    }
    const name = quotedSheet !== undefined ? quotedSheet.replace(/''/g, "'") : sheet ?? '';
    const text = name ? `${book.dir}[${book.file}]${name}` : `${book.dir}${book.file}`;
    const plain = !book.dir && quoteSheetName(book.file) === book.file && (!name || quoteSheetName(name) === name);
    return plain ? `${text}!` : `'${text.replace(/'/g, "''")}'!`;
  });
}

// ===== NUMBER FORMATS =====

/** Splits a format code into its ';' sections, honouring quotes, escapes and [..]. */
function splitFormatSections(fmt: string): string[] {
  const out: string[] = [];
  let start = 0;
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (quoted) {
      if (ch === '"') quoted = false;
    } else if (bracket) {
      if (ch === ']') bracket = false;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      i++; // the next character is a literal / padding character
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === '[') {
      bracket = true;
    } else if (ch === ';') {
      out.push(fmt.slice(start, i));
      start = i + 1;
    }
  }
  out.push(fmt.slice(start));
  return out;
}

/** Does the format change how text is displayed (a text section or an '@')? */
function formatAffectsText(fmt: string): boolean {
  const sections = splitFormatSections(fmt);
  if (sections.length >= 4) {
    return true;
  }
  return sections.some((section) => section.replace(/"[^"]*"|\\./g, '').includes('@'));
}

/**
 * Excel "General" for numbers: up to 11 characters (a minus sign not counted), else 6-digit scientific. Digits are
 * rounded half away from zero on the 15-digit decimal value, as Excel does (8238230.9475 -> '8238230.948', where
 * ssf's binary rounding gave '8238230.947'). Below 0.0001 a value is shown in full when 9 decimals hold it exactly
 * ('0.000012345'), else in scientific ('1.23456E-05'; 0.000099999999995 -> '1E-04').
 */
function formatGeneral(value: number): string {
  const abs = Math.abs(value);
  if (Number.isInteger(value) && abs < 1e11) {
    return String(value);
  }
  const dec = toDecimal15(abs);
  if (!dec.digits) {
    return '0';
  }
  const sign = value < 0 ? '-' : '';
  const e = dec.point - 1; // abs = d.ddd x 10^e
  if (e < 11 && (e >= -4 || dec.digits.length - dec.point <= 9)) {
    // 10 significant digits fit (9 decimals below 1); 10 or 11 integer digits leave no room for decimals, and
    // only a value that rounds to 12 integer digits goes scientific.
    const r = roundDecimal(dec, 0, Math.max(0, 9 - Math.max(e, 0)));
    if (r.int.length <= 11) {
      const frac = r.frac.replace(/0+$/, '');
      return `${sign}${r.int || '0'}${frac ? `.${frac}` : ''}`;
    }
  }
  let exp = e;
  let r = roundDecimal(dec, -exp, 5);
  if (r.int.length > 1) {
    exp++; // 9.999995E+20 -> 1E+21
    r = roundDecimal(dec, -exp, 5);
  }
  const frac = r.frac.replace(/0+$/, '');
  return `${sign}${r.int}${frac ? `.${frac}` : ''}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

// ----- numeric sections -----
//
// Number sections (everything but dates, times, fractions and text) are laid out
// here rather than by ssf, which loses the sign rules, Excel's 15-digit rounding,
// '?' and '#' decimals, grouping with optional decimals, scientific exponents and
// conditions. Rules were checked against real Excel 16 (fix-A nf-*.js oracle,
// ~3,500 format / value pairs).

/** Smallest normal double: Excel has no denormals and shows them as 0. */
const MIN_NORMAL = 2.2250738585072014e-308;

type DigitPlaceholder = '0' | '#' | '?';

/** One token of a number section; digit placeholders refer to their position in ints / decs / exps. */
type NumToken =
  | { k: 'lit'; v: string } // literal text ("..", \x, '%', [$€-407] currency ...)
  | { k: 'pad'; c: string } // '_x': a blank as wide as x (a space in the text, see FillLayout.pads)
  | { k: 'fill'; c: string } // '*x': x repeated to fill the cell (nothing in the text, see FillLayout)
  | { k: 'int'; i: number }
  | { k: 'dec'; i: number }
  | { k: 'exp'; i: number }
  | { k: 'dot' }
  | { k: 'e' } // the E+ / E- marker
  | { k: 'general' };

/** Collects the CellData.fill layout while a value is formatted (only for formats with a '*' fill). */
interface LayoutOut {
  fill?: FillLayout;
}

interface NumCondition {
  op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  value: number;
}

interface NumSection {
  tokens: NumToken[];
  cond?: NumCondition;
  ints: DigitPlaceholder[];
  decs: DigitPlaceholder[];
  exps: DigitPlaceholder[];
  /** Thousands separators (a comma between integer placeholders). */
  group: boolean;
  /** Power of ten applied before display: +2 per '%', -3 per scaling comma ('0,' or '0.0,,'). */
  shift: number;
  /** Scientific: does the exponent always show its sign (E+) or only when negative (E-)? */
  sci?: { plus: boolean };
  /** Contains the 'General' keyword. */
  general: boolean;
  /** Has a '*x' repeat-to-fill (only the first one counts, as in Excel). */
  fill: boolean;
}

interface NumFormat {
  /** The number sections (1..3; the text section is not included). */
  sections: NumSection[];
  /** Some section has a '*x' fill: values get a FillLayout. */
  fill: boolean;
  /**
   * Per section: can values >= 0 end up there? A section that only ever gets
   * negative values (the 2nd of '0;(0)', or '[<0]..') shows |value| without a
   * minus; any other section shows negative values with a leading '-'.
   */
  signed: boolean[];
}

/** A number as Excel holds it: 15 significant decimal digits (no leading / trailing zeros; '' = 0). */
interface Decimal15 {
  digits: string;
  /** Number of digits before the decimal point (may be <= 0 or exceed digits.length). */
  point: number;
}

function toDecimal15(abs: number): Decimal15 {
  if (!(abs >= MIN_NORMAL) || !Number.isFinite(abs)) {
    return { digits: '', point: 0 };
  }
  const m = /^(\d+)(?:\.(\d*))?(?:e([+-]\d+))?$/.exec(abs.toPrecision(15));
  if (!m) {
    return { digits: '', point: 0 };
  }
  const all = m[1] + (m[2] ?? '');
  const lead = all.length - all.replace(/^0+/, '').length;
  return { digits: all.slice(lead).replace(/0+$/, ''), point: m[1].length + (m[3] ? parseInt(m[3], 10) : 0) - lead };
}

/**
 * Rounds `dec` * 10^shift half away from zero to `decimals` places, in decimal
 * (1.005 -> '1.01', 0.145 * 100 -> '15'). Returns the integer digits ('' for
 * zero) and exactly `decimals` fraction digits.
 */
function roundDecimal(dec: Decimal15, shift: number, decimals: number): { int: string; frac: string } {
  let digits = dec.digits;
  let point = dec.point + shift;
  const keep = point + decimals;
  if (!digits || keep < 0) {
    return { int: '', frac: '0'.repeat(decimals) };
  }
  if (keep < digits.length) {
    const up = digits.charCodeAt(keep) >= 53; // '5'
    digits = digits.slice(0, keep);
    if (up) {
      let i = digits.length - 1;
      while (i >= 0 && digits[i] === '9') i--;
      if (i < 0) {
        digits = `1${'0'.repeat(digits.length)}`;
        point++;
      } else {
        digits = `${digits.slice(0, i)}${String.fromCharCode(digits.charCodeAt(i) + 1)}${'0'.repeat(digits.length - i - 1)}`;
      }
    }
  }
  const int = point > 0 ? digits.slice(0, point).padEnd(point, '0') : '';
  const frac = (point >= 0 ? digits.slice(point) : '0'.repeat(-point) + digits).slice(0, decimals).padEnd(decimals, '0');
  return { int: /^0*$/.test(int) ? '' : int, frac };
}

/**
 * roundDecimal(toDecimal15(abs), shift, decimals) with a cheap float path: while
 * the result has at most 15 digits and the value is not within float noise of a
 * rounding boundary, Math.round gives the same digits.
 */
function roundFixed(abs: number, shift: number, decimals: number): { int: string; frac: string } {
  const scaled = abs * 10 ** (shift + decimals);
  if (scaled < 1e15) {
    const floor = Math.floor(scaled);
    const half = scaled - floor - 0.5;
    if (Math.abs(half) > 1e-9 + scaled * 1e-14) {
      const n = half > 0 ? floor + 1 : floor;
      if (!decimals) {
        return { int: n ? String(n) : '', frac: '' };
      }
      const text = String(n).padStart(decimals + 1, '0');
      const int = text.slice(0, -decimals);
      return { int: int === '0' ? '' : int, frac: text.slice(-decimals) };
    }
  }
  return roundDecimal(toDecimal15(abs), shift, decimals);
}

const NUM_CONDITION_RE = /^(<=|>=|<>|<|>|=)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)$/;
const NUM_COLOR_RE = /^(?:black|white|red|green|blue|yellow|magenta|cyan|color\s*\d+)$/i;

/** Parses one number section; undefined for anything else (dates, fractions, '@', unknown codes, letters Excel rejects). */
function compileNumberSection(src: string): NumSection | undefined {
  const s: NumSection = { tokens: [], ints: [], decs: [], exps: [], group: false, shift: 0, general: false, fill: false };
  const commas: { at: number; zone: 'int' | 'dec' | 'exp' }[] = [];
  let zone: 'int' | 'dec' | 'exp' = 'int';
  const lit = (v: string): void => {
    const last = s.tokens[s.tokens.length - 1];
    if (last?.k === 'lit') last.v += v;
    else if (v) s.tokens.push({ k: 'lit', v });
  };
  for (let i = 0; i < src.length; ) {
    const ch = src[i];
    switch (ch) {
      case '"': {
        const end = src.indexOf('"', i + 1);
        if (end < 0) return undefined;
        lit(src.slice(i + 1, end));
        i = end + 1;
        break;
      }
      case '\\':
        lit(src[i + 1] ?? '');
        i += 2;
        break;
      case '_': // reserves the width of the next character
        s.tokens.push({ k: 'pad', c: src[i + 1] ?? ' ' });
        i += 2;
        break;
      case '*': // repeats the next character to fill the cell (laid out by the grid: CellData.fill)
        if (!s.fill) s.tokens.push({ k: 'fill', c: src[i + 1] ?? ' ' });
        s.fill = true;
        i += 2;
        break;
      case '[': {
        const end = src.indexOf(']', i);
        if (end < 0) return undefined;
        const body = src.slice(i + 1, end).trim();
        i = end + 1;
        const cond = NUM_CONDITION_RE.exec(body);
        if (cond) {
          s.cond = { op: cond[1] as NumCondition['op'], value: parseFloat(cond[2]) };
        } else if (body.startsWith('$')) {
          lit(body.slice(1).replace(/-[0-9A-Fa-f]*$/, '')); // [$€-407] -> '€', [$-409] -> ''
        } else if (!NUM_COLOR_RE.test(body)) {
          return undefined; // [h] elapsed time, [DBNum1] ...
        }
        break;
      }
      case '0':
      case '#':
      case '?': {
        const list = zone === 'int' ? s.ints : zone === 'dec' ? s.decs : s.exps;
        s.tokens.push({ k: zone, i: list.length });
        list.push(ch);
        i++;
        break;
      }
      case '.':
        if (zone === 'int') {
          s.tokens.push({ k: 'dot' });
          zone = 'dec';
        } else {
          lit('.');
        }
        i++;
        break;
      case ',':
        commas.push({ at: s.tokens.length, zone });
        i++;
        break;
      case '%':
        lit('%');
        s.shift += 2;
        i++;
        break;
      case 'E':
      case 'e':
        if ((src[i + 1] !== '+' && src[i + 1] !== '-') || s.sci) return undefined;
        s.sci = { plus: src[i + 1] === '+' };
        s.tokens.push({ k: 'e' });
        zone = 'exp';
        i += 2;
        break;
      default:
        if (/^general/i.test(src.slice(i, i + 7))) {
          s.tokens.push({ k: 'general' });
          s.general = true;
          i += 7;
          break;
        }
        // Date / time codes (y m d h s, e g b era years) and 'n' do not go with digits; '@' text, '/' fractions. Excel
        // takes any other letter as a literal ('0.0x', '0.0 °C', '#,##0 K'; it escapes them only when it saves).
        if (/[BDEGHMNSYbdeghmnsy@/]/.test(ch)) return undefined;
        lit(ch);
        i++;
        break;
    }
  }
  // A comma followed by more integer placeholders groups thousands; one after the
  // last placeholder of the number scales by 1000.
  for (const comma of commas) {
    const rest = s.tokens.slice(comma.at);
    if (comma.zone === 'int' && rest.some((t) => t.k === 'int')) s.group = true;
    else if (comma.zone !== 'exp' && !rest.some((t) => t.k === 'int' || t.k === 'dec')) s.shift -= 3;
  }
  if (s.general && s.ints.length + s.decs.length + s.exps.length) {
    return undefined;
  }
  return s;
}

/** True when the section holds a '@' text placeholder. */
function isTextSection(section: string): boolean {
  return section.replace(/"[^"]*"|\\.|_.|\*./g, '').includes('@');
}

/**
 * Text through a text section with a '*' fill (`@*-`, `* @`, `_(* @_)`): '@' is the value, quoted / escaped
 * characters are literals, '_x' a padding blank, [colour] / [$€-407] tags as in number sections. The fill character
 * itself is not part of the text; `layout` gets its position (ssf would keep one copy or drop it).
 */
function renderTextSection(section: string, value: string, layout?: LayoutOut): string {
  const fill: FillLayout | undefined = layout ? { at: 0 } : undefined;
  let text = '';
  let filled = false;
  for (let i = 0; i < section.length; ) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      text += section.slice(i + 1, end < 0 ? section.length : end);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '\\') {
      text += section[i + 1] ?? '';
      i += 2;
    } else if (ch === '_' || ch === '*') {
      if (ch === '_' || !filled) text += layoutToken({ k: ch === '_' ? 'pad' : 'fill', c: section[i + 1] ?? ' ' }, text.length, fill);
      filled ||= ch === '*';
      i += 2;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const body = end < 0 ? '' : section.slice(i + 1, end);
      if (body.startsWith('$')) text += body.slice(1).replace(/-[0-9A-Fa-f]*$/, '');
      i = end < 0 ? section.length : end + 1;
    } else {
      text += ch === '@' ? value : ch;
      i++;
    }
  }
  if (fill && layout && filled) layout.fill = fill;
  return text;
}

/** Compiles a number format; undefined when a number section is not plain numeric (left to ssf). */
function compileNumberFormat(fmt: string): NumFormat | undefined {
  const raw = splitFormatSections(fmt);
  if (raw.length > 4) {
    return undefined;
  }
  // A last section with '@' (or the 4th) is the text section. A format that is only a text section shows numbers as General.
  const numeric = raw.length === 4 ? raw.slice(0, 3) : isTextSection(raw[raw.length - 1]) ? raw.slice(0, -1) : raw;
  const sections: NumSection[] = [];
  for (const src of numeric.length ? numeric : ['General']) {
    const section = compileNumberSection(src);
    if (!section) return undefined;
    sections.push(section);
  }
  const f: NumFormat = { sections, signed: [], fill: sections.some((section) => section.fill) };
  const probes = [0, 1, 1e300];
  for (const section of sections) {
    const t = section.cond?.value;
    if (t !== undefined) {
      const d = Math.max(1, Math.abs(t)) * 1e-9;
      probes.push(t, t - d, t + d);
    }
  }
  f.signed = sections.map((_, i) => probes.some((p) => p >= 0 && selectNumberSection(f, p) === i));
  return f;
}

function testCondition(c: NumCondition, v: number): boolean {
  switch (c.op) {
    case '<': return v < c.value;
    case '<=': return v <= c.value;
    case '>': return v > c.value;
    case '>=': return v >= c.value;
    case '=': return v === c.value;
    default: return v !== c.value;
  }
}

/** Index of the section that formats `v` (Excel's / ssf's rules, conditions included). */
function selectNumberSection(f: NumFormat, v: number): number {
  const n = f.sections.length;
  const c0 = f.sections[0].cond;
  const c1 = n > 1 ? f.sections[1].cond : undefined;
  if (c0 || c1) {
    if (c0 && testCondition(c0, v)) return 0;
    if (c1 && testCondition(c1, v)) return 1;
    return c0 && c1 ? (n > 2 ? 2 : 0) : n > 1 ? 1 : 0;
  }
  if (v < 0) return n > 1 ? 1 : 0;
  return v === 0 && n > 2 ? 2 : 0;
}

/** `layout` (formats with a '*' fill) receives where the fill and the '_x' paddings are in the returned text. */
function renderNumberFormat(f: NumFormat, value: number, layout?: LayoutOut): string {
  const index = selectNumberSection(f, value);
  const section = f.sections[index];
  const fill: FillLayout | undefined = layout && section.fill ? { at: 0 } : undefined;
  const { text, zero } = renderNumberSection(section, Math.abs(value), fill);
  // Like Excel, a value that displays as zero ('0' for -0.4) gets no minus.
  const minus = value < 0 && f.signed[index] && !zero;
  if (fill && layout) layout.fill = minus ? shiftFillLayout(fill, 1) : fill;
  return minus ? `-${text}` : text;
}

/** The layout of a text with `n` more characters in front of it. */
function shiftFillLayout(fill: FillLayout, n: number): FillLayout {
  return { ...fill, at: fill.at + n, ...(fill.pads ? { pads: fill.pads.map(([at, c]): [number, string] => [at + n, c]) } : {}) };
}

/** Text of a fill / padding token; records its position in `fill` (when given) as the text is built. */
function layoutToken(t: NumToken & { k: 'pad' | 'fill' }, at: number, fill: FillLayout | undefined): string {
  if (t.k === 'fill') {
    if (fill) {
      fill.at = at;
      if (t.c !== ' ') fill.char = t.c;
    }
    return '';
  }
  if (fill) (fill.pads ??= []).push([at, t.c]);
  return ' ';
}

/** Lays out |value| in one section; `zero` = every displayed digit is 0. */
function renderNumberSection(s: NumSection, abs: number, fill?: FillLayout): { text: string; zero: boolean } {
  if (s.general || !(s.ints.length + s.decs.length + s.exps.length)) {
    let text = '';
    for (const t of s.tokens) {
      if (t.k === 'pad' || t.k === 'fill') text += layoutToken(t, text.length, fill);
      else text += t.k === 'lit' ? t.v : t.k === 'general' ? formatGeneral(abs) : '';
    }
    return { text, zero: false };
  }
  let r: { int: string; frac: string };
  let expText = '';
  let zeroMantissa = false;
  if (s.sci) {
    // Scientific: the exponent is a multiple of the number of integer placeholders
    // (##0.0E+0 is engineering notation); a mantissa rounded past them moves up.
    const dec = toDecimal15(abs);
    const m = Math.max(1, s.ints.length);
    let e = 0;
    if (dec.digits) {
      e = Math.floor((dec.point + s.shift - 1) / m) * m;
      r = roundDecimal(dec, s.shift - e, s.decs.length);
      if (r.int.length > m) {
        e += m;
        r = roundDecimal(dec, s.shift - e, s.decs.length);
      }
    } else {
      r = { int: '', frac: '0'.repeat(s.decs.length) };
      zeroMantissa = true; // Excel shows 0 as 000.0E+0 for ##0.0E+0: every placeholder a '0'
    }
    expText = `${e < 0 ? '-' : s.sci.plus ? '+' : ''}${String(Math.abs(e)).padStart(s.exps.length, '0')}`;
  } else {
    r = roundFixed(abs, s.shift, s.decs.length);
  }
  // Integer digits fill the placeholders from the right; extra digits go before the first one.
  const n = s.ints.length;
  const slots = new Array<string>(n);
  for (let j = n - 1, d = r.int.length - 1; j >= 0; j--, d--) {
    const ph = s.ints[j];
    slots[j] = d >= 0 ? r.int[d] : zeroMantissa || ph === '0' ? '0' : ph === '?' ? ' ' : '';
  }
  let lead = r.int.length > n ? r.int.slice(0, r.int.length - n) : '';
  const seps = new Array<string>(n).fill('');
  if (s.group) {
    // Separator every 3 shown digits from the right; next to a '?' blank it is a blank too.
    const units: number[] = []; // lead digit k -> -1 - k, else the slot index
    for (let k = 0; k < lead.length; k++) units.push(-1 - k);
    for (let j = 0; j < n; j++) if (slots[j]) units.push(j);
    const charAt = (u: number): string => (u < 0 ? lead[-1 - u] : slots[u]);
    let grouped = '';
    units.forEach((u, idx) => {
      const sep = idx > 0 && (units.length - idx) % 3 === 0 ? (charAt(units[idx - 1]) === ' ' ? ' ' : ',') : '';
      if (u < 0) grouped += sep + charAt(u);
      else seps[u] = sep;
    });
    lead = grouped;
  }
  // Trailing zero decimals: '#' drops them, '?' turns them into blanks.
  const decChars = r.frac.split('');
  for (let j = decChars.length - 1; j >= 0 && decChars[j] === '0' && s.decs[j] !== '0'; j--) {
    decChars[j] = s.decs[j] === '?' ? ' ' : '';
  }
  let text = '';
  let leadDone = false;
  for (const t of s.tokens) {
    switch (t.k) {
      case 'lit':
        text += t.v;
        break;
      case 'pad':
      case 'fill':
        text += layoutToken(t, text.length, fill);
        break;
      case 'int':
        if (!leadDone) text += lead;
        leadDone = true;
        text += seps[t.i] + slots[t.i];
        break;
      case 'dot':
      case 'e':
        if (!leadDone) text += lead;
        leadDone = true;
        text += t.k === 'dot' ? '.' : `E${expText}`;
        break;
      case 'dec':
        text += decChars[t.i];
        break;
      default:
        break; // exponent digits are part of expText
    }
  }
  return { text, zero: !s.sci && !r.int && !/[1-9]/.test(r.frac) };
}

// ----- date / time sections -----

/** Characters ssf reads in a date section as codes ('a' only in AM/PM and A/P) or as plain literals. */
const SSF_DATE_CHARS: ReadonlySet<string> = new Set('yYmMdDhHsSeEbBaA0123456789#?@ ,$-+/():!^&\'~{}<>=€cfijklopqrtuvwxzP');

/**
 * A date / time format as ssf can read it. Excel takes every character that is not a code as a literal (and escapes
 * it when it saves the file: 'dd\.mm\.yyyy'), but ExcelJS, openpyxl and pandas store the code as it was typed
 * ('dd.mm.yyyy', 'yyyy-mm-ddThh:mm:ss', 'yyyy年m月d日'), and ssf throws on such literals ("bad second format: .").
 * In each date section, a '.' that does not start sub-seconds ('ss.00') and every character ssf knows neither as a
 * code nor as a literal (T, CJK, '|', ...) are quoted; the era code g / ggg is dropped (empty for the Gregorian
 * calendar, as Excel 16 shows it). Number and text sections are left as they are.
 */
function normalizeDateFormat(fmt: string): string {
  return splitFormatSections(fmt)
    .map((section) => {
      let date = false;
      try {
        date = !isTextSection(section) && SSF.is_date(section);
      } catch {
        // not a date section
      }
      return date ? normalizeDateSection(section) : section;
    })
    .join(';');
}

function normalizeDateSection(section: string): string {
  let out = '';
  let literal = '';
  const flush = (): void => {
    if (literal) out += `"${literal}"`;
    literal = '';
  };
  for (let i = 0; i < section.length; ) {
    const ch = section[i];
    if (ch === '"' || ch === '[') {
      const end = section.indexOf(ch === '"' ? '"' : ']', i + 1);
      const stop = end < 0 ? section.length : end + 1;
      flush();
      out += section.slice(i, stop);
      i = stop;
    } else if (ch === '\\' || ch === '_') {
      // '\x' literal, '_x' blank: ssf takes one UTF-16 unit for x, so a character outside the BMP is quoted / padded as a space
      const next = String.fromCodePoint(section.codePointAt(i + 1) ?? 32);
      flush();
      out += next.length === 1 ? section.slice(i, i + 2) : ch === '\\' ? `"${next}"` : '_ ';
      i += 1 + next.length;
    } else if (ch === '*') {
      // '*x' fill: dates get no fill layout, so x shows once, as ssf shows most fill characters ('* ' and '**' none)
      const next = i + 1 < section.length ? String.fromCodePoint(section.codePointAt(i + 1) ?? 32) : '';
      if (next === ' ' || next === '*' || next === '"') {
        flush();
        out += next === '"' ? '\\"' : `*${next}`; // kept: dropping it could join two codes ('dddd* d')
      } else {
        literal += next;
      }
      i += 1 + next.length;
    } else if (ch === '.' && section[i + 1] !== '0') {
      literal += '.';
      i++;
    } else if (/^general/i.test(section.slice(i, i + 7)) || section.startsWith('上午/下午', i)) {
      const n = ch === '上' ? 5 : 7;
      flush();
      out += section.slice(i, i + n);
      i += n;
    } else if (ch === 'g' || ch === 'G') {
      i++; // era
    } else if (ch === '.' || SSF_DATE_CHARS.has(ch)) {
      flush();
      out += ch;
      i++;
    } else {
      const cp = String.fromCodePoint(section.codePointAt(i) ?? 32);
      literal += cp;
      i += cp.length;
    }
  }
  flush();
  return out;
}

/** ssf's last serial (9999-12-31 0:00); Excel's last day goes on to 23:59:59.999. */
const SSF_LAST_SERIAL = 2_958_465;
/** Days in 400 Gregorian years: after them, months, days and weekdays repeat. */
const GREGORIAN_CYCLE_DAYS = 146_097;
/** NumberFormatter.lastDay: the marks put around [h] / [m] / [s] (control characters no format shows) and their units per day. */
const ELAPSED_MARKS = { h: '\u0001', m: '\u0002', s: '\u0003' };
const ELAPSED_PER_DAY = [24, 1_440, 86_400];

/**
 * Number formatting: number sections are laid out by compileNumberFormat /
 * renderNumberFormat, dates / times / fractions / text sections by ssf (dates
 * rounded and range-checked like Excel first). Compiled formats and formatted
 * dates are cached per format; formats ssf cannot parse fall back to General.
 */
class NumberFormatter {
  private readonly opts: { date1904: boolean };
  private readonly dateFlags = new Map<string, boolean>();
  private readonly textFlags = new Map<string, boolean>();
  private readonly compiled = new Map<string, NumFormat | null>();
  private readonly textFills = new Map<string, string | null>();
  /** Date formats as ssf reads them (normalizeDateFormat) + their sub-second digits. */
  private readonly dateFormats = new Map<string, { code: string; digits: number }>();
  private readonly dateCache = new Map<string, Map<number, string>>();
  private readonly broken = new Set<string>();

  constructor(date1904: boolean) {
    this.opts = { date1904 };
  }

  /**
   * True when the format displays a date and/or time. A format that compiles as a number format never does, even
   * when ssf thinks so: ssf reads the 'E' of an exponent after a literal ('0.00\ E+00', as Excel saves it) as a date code.
   */
  isDate(fmt: string): boolean {
    let flag = this.dateFlags.get(fmt);
    if (flag === undefined) {
      try {
        flag = fmt !== 'General' && !this.compile(fmt) && SSF.is_date(fmt);
      } catch {
        flag = false;
      }
      this.dateFlags.set(fmt, flag);
    }
    return flag;
  }

  /** Formats a number; with `layout`, a format with a '*' fill also reports where the fill goes (CellData.fill). */
  number(fmt: string, value: number, layout?: LayoutOut): string {
    if (fmt === 'General' || this.broken.has(fmt)) {
      return formatGeneral(value);
    }
    if (this.isDate(fmt)) {
      return this.date(fmt, value);
    }
    const compiled = this.compile(fmt);
    if (compiled) {
      return renderNumberFormat(compiled, value, layout);
    }
    try {
      return SSF.format(fmt, value, this.opts); // fractions
    } catch {
      this.checkBroken(fmt);
      return formatGeneral(value);
    }
  }

  /**
   * Date / time formats. Like Excel, the serial is first rounded to the precision
   * on display (whole seconds, or the decimals of 'ss.00') - so 23:59:59.6 shows
   * as the next day 0:00:00, a carry ssf drops. Negative serials exist only in the
   * 1904 date system ('-' + the positive value); past 9999-12-31 is '########'.
   * Values repeat a lot (one date per day), so results are cached per format.
   */
  private date(fmt: string, value: number): string {
    const { code, digits } = this.dateFormat(fmt);
    if (value < 0) {
      return this.opts.date1904 ? `-${this.date(fmt, -value)}` : DATE_OVERFLOW_TEXT;
    }
    const unit = 86_400 * 10 ** digits;
    const serial = Math.round(value * unit) / unit;
    if (!(serial < (this.opts.date1904 ? DATE_LIMIT_SERIAL - DATE1904_OFFSET : DATE_LIMIT_SERIAL))) {
      return DATE_OVERFLOW_TEXT;
    }
    let cache = this.dateCache.get(fmt);
    if (!cache) {
      this.dateCache.set(fmt, (cache = new Map()));
    }
    let text = cache.get(serial);
    if (text === undefined) {
      try {
        text = serial > SSF_LAST_SERIAL ? this.lastDay(code, serial) : SSF.format(code, serial, this.opts);
      } catch {
        this.checkBroken(fmt);
        return formatGeneral(value);
      }
      if (cache.size >= DATE_CACHE_LIMIT) cache.clear();
      cache.set(serial, text);
    }
    return text;
  }

  /** The format as ssf reads it (normalizeDateFormat) and its sub-second digits ('ss.00' -> 2). */
  private dateFormat(fmt: string): { code: string; digits: number } {
    let info = this.dateFormats.get(fmt);
    if (!info) {
      const code = normalizeDateFormat(fmt);
      const m = /\.(0+)/.exec(code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, ''));
      info = { code, digits: m ? Math.min(3, m[1].length) : 0 };
      this.dateFormats.set(fmt, info);
    }
    return info;
  }

  /**
   * A time on 9999-12-31 (1900 system), past ssf's last serial (ssf returns ''). The value is formatted 400 and 800
   * years earlier - same month, day, weekday and time - and the characters that differ between the two texts (the
   * year digits 9599 / 9199) are the 9s of 9999. Elapsed times ([h], [m], [s]) count from serial 0, so each one is
   * formatted between two marks (quoted, so ssf reads the format the same way) and gets the hours / minutes /
   * seconds of the shift back.
   */
  private lastDay(code: string, serial: number): string {
    const marked = code.replace(/"[^"]*"|\\.|\[(h+|m+|s+)\]/gi, (all: string, unit?: string) => {
      const mark = unit ? ELAPSED_MARKS[unit.charAt(0).toLowerCase() as keyof typeof ELAPSED_MARKS] : '';
      return mark ? `"${mark}"${all}"${mark}"` : all;
    });
    const shifted = (days: number): string =>
      SSF.format(marked, serial - days, this.opts).replace(/([\u0001-\u0003])(\d+)\1/g, (_all, mark: string, count: string) =>
        String(Number(count) + days * ELAPSED_PER_DAY[mark.charCodeAt(0) - 1]),
      );
    const a = shifted(GREGORIAN_CYCLE_DAYS);
    const b = shifted(2 * GREGORIAN_CYCLE_DAYS);
    if (a.length !== b.length) {
      return a;
    }
    let text = '';
    for (let i = 0; i < a.length; i++) text += a[i] === b[i] ? a[i] : '9';
    return text;
  }

  /** Does a number shown with this format get a FillLayout ('*' fill in its number sections)? Dates do not. */
  numberHasFill(fmt: string): boolean {
    return fmt !== 'General' && !!this.compile(fmt)?.fill;
  }

  /** The text section of the format when it has a '*' fill (laid out here, not by ssf), else null. */
  textFillSection(fmt: string): string | null {
    let section = this.textFills.get(fmt);
    if (section === undefined) {
      const sections = splitFormatSections(fmt);
      const text = sections.length >= 4 ? sections[3] : isTextSection(sections[sections.length - 1]) ? sections[sections.length - 1] : undefined;
      section = text !== undefined && /\*./.test(text.replace(/"[^"]*"|\\.|_./g, '')) ? text : null;
      this.textFills.set(fmt, section);
    }
    return section;
  }

  private compile(fmt: string): NumFormat | null {
    let compiled = this.compiled.get(fmt);
    if (compiled === undefined) {
      compiled = compileNumberFormat(fmt) ?? null;
      this.compiled.set(fmt, compiled);
    }
    return compiled;
  }

  /** Formats text; with `layout`, a text section with a '*' fill also reports where the fill goes (CellData.fill). */
  text(fmt: string, value: string, layout?: LayoutOut): string {
    if (fmt === 'General' || fmt === '@' || this.broken.has(fmt)) {
      return value;
    }
    let affects = this.textFlags.get(fmt);
    if (affects === undefined) {
      affects = formatAffectsText(fmt);
      this.textFlags.set(fmt, affects);
    }
    if (!affects) {
      return value;
    }
    const fillSection = this.textFillSection(fmt);
    if (fillSection !== null) {
      return renderTextSection(fillSection, value, layout);
    }
    try {
      return SSF.format(fmt, value, this.opts);
    } catch {
      this.checkBroken(fmt);
      return value;
    }
  }

  /** After a failure: formats ssf cannot parse at all fall back to General from now on. */
  private checkBroken(fmt: string): void {
    try {
      SSF.format(this.isDate(fmt) ? this.dateFormat(fmt).code : fmt, 1, this.opts);
    } catch {
      this.broken.add(fmt);
    }
  }

}

// ===== STYLES =====

function isStyleEmpty(style: XStyle | undefined): boolean {
  return !style || (!style.numFmt && !style.font && !style.fill && !style.border && !style.alignment);
}

/** Deduplicated style table: identity cache first, JSON key second. styles[0] = default. */
class StyleTable {
  readonly styles: CellStyle[] = [{}];
  private readonly byKey = new Map<string, number>([['{}', 0]]);
  private readonly byObject = new Map<object, number>();

  constructor(private readonly convert: (style: XStyle) => CellStyle) {}

  /** Style-table index of an ExcelJS style object (0 = default). */
  indexOf(style: XStyle | undefined): number {
    if (!style) {
      return 0;
    }
    const known = this.byObject.get(style);
    if (known !== undefined) {
      return known;
    }
    if (isStyleEmpty(style)) {
      return 0; // empty objects are common (merged cells) and not worth caching
    }
    const index = this.add(this.convert(style));
    this.byObject.set(style, index);
    return index;
  }

  add(style: CellStyle): number {
    const key = JSON.stringify(style);
    let index = this.byKey.get(key);
    if (index === undefined) {
      index = this.styles.length;
      this.styles.push(style);
      this.byKey.set(key, index);
    }
    return index;
  }

  /**
   * Forgets the object-identity cache (indices stay valid). ExcelJS shares one
   * style object between cells and mutates it in place when a cell's font /
   * fill / numFmt is assigned, so cached identities can go stale after edits.
   */
  forgetObjects(): void {
    this.byObject.clear();
  }

  /** [index, style] pairs for the given indices (ascending, default excluded). */
  entries(indices: Iterable<number>): [number, CellStyle][] {
    return [...indices]
      .filter((i) => i > 0 && i < this.styles.length)
      .sort((a, b) => a - b)
      .map((i): [number, CellStyle] => [i, this.styles[i]]);
  }
}

/** Converts ExcelJS style models (cells, rows, columns, dxf) into CellStyle. */
class StyleConverter {
  constructor(
    private readonly palette: ThemePalette,
    private readonly defaultFont: { name: string; size: number },
    private readonly extras: WorkbookExtras | undefined,
  ) {}

  color(color: XColor | undefined): string | undefined {
    return resolveColor(color, this.palette);
  }

  /** Cell / row / column style. Values equal to the workbook default are omitted. */
  cellStyle(style: XStyle): CellStyle {
    const out: CellStyle = {};
    const font = style.font ? this.cellFont(style.font) : undefined;
    if (font) out.font = font;
    this.applyFill(out, style.fill, false);
    const border = this.border(style.border);
    if (border) out.border = border;
    const align = this.alignment(style.alignment);
    if (align) out.align = align;
    const numFmt = this.numFmt(style);
    if (numFmt !== 'General') out.numFmt = numFmt;
    return out;
  }

  /** Differential (conditional-format) style: only what the dxf specifies, colours kept as-is. */
  dxfStyle(style: XStyle): CellStyle | undefined {
    const out: CellStyle = {};
    if (style.font) {
      const f = style.font;
      const font: FontStyle = {};
      if (f.name) font.name = f.name;
      if (typeof f.size === 'number' && f.size > 0) font.size = f.size;
      if (f.bold) font.bold = true;
      if (f.italic) font.italic = true;
      const underline = normalizeUnderline(f.underline);
      if (underline) font.underline = underline;
      if (f.strike) font.strike = true;
      const color = this.color(f.color);
      if (color) font.color = color;
      if (f.vertAlign === 'superscript' || f.vertAlign === 'subscript') font.vertAlign = f.vertAlign;
      if (Object.keys(font).length) out.font = font;
    }
    this.applyFill(out, style.fill, true);
    const border = this.border(style.border);
    if (border) out.border = border;
    const align = this.alignment(style.alignment);
    if (align) out.align = align;
    const code = typeof style.numFmt === 'string' ? style.numFmt : style.numFmt?.formatCode;
    if (code && code !== 'General') out.numFmt = code;
    return Object.keys(out).length ? out : undefined;
  }

  /**
   * Font of a rich-text run, relative to the cell's own font: only properties that
   * differ are returned (the webview layers the run over the cell font).
   */
  runFont(run: XFont, cell: XFont | undefined): FontStyle | undefined {
    const out: FontStyle = {};
    const name = run.name || this.defaultFont.name;
    if (name !== (cell?.name || this.defaultFont.name)) out.name = name;
    const size = this.fontSize(run) ?? this.defaultFont.size;
    if (size !== ((cell && this.fontSize(cell)) ?? this.defaultFont.size)) out.size = size;
    if (!!run.bold !== !!cell?.bold) out.bold = !!run.bold;
    if (!!run.italic !== !!cell?.italic) out.italic = !!run.italic;
    if (!!run.strike !== !!cell?.strike) out.strike = !!run.strike;
    const underline = normalizeUnderline(run.underline);
    if (underline && underline !== normalizeUnderline(cell?.underline)) out.underline = underline;
    const vert = run.vertAlign === 'superscript' || run.vertAlign === 'subscript' ? run.vertAlign : undefined;
    if (vert && vert !== cell?.vertAlign) out.vertAlign = vert;
    const color = this.textColor(run.color);
    const cellColor = this.textColor(cell?.color);
    if (color !== cellColor) out.color = color ?? '#000000';
    return Object.keys(out).length ? out : undefined;
  }

  /** Text colour, undefined for "Automatic" (and plain black, which is what Automatic shows). */
  private textColor(color: XColor | undefined): string | undefined {
    if (isAutomaticColor(color)) {
      return undefined;
    }
    const hex = this.color(color);
    return hex === '#000000' ? undefined : hex;
  }

  /**
   * Exact font size: ExcelJS truncates sz to an integer, the raw styles.xml does
   * not. The raw size is used only while the ExcelJS font still has its loaded size.
   */
  private fontSize(font: XFont): number | undefined {
    const loaded = this.extras?.fontIndex.get(font);
    const raw = loaded && loaded.size === font.size ? this.extras?.raw?.fonts[loaded.index]?.size : undefined;
    const size = raw ?? font.size;
    return typeof size === 'number' && size > 0 ? size : undefined;
  }

  private cellFont(f: XFont): FontStyle | undefined {
    const out: FontStyle = {};
    if (f.name && f.name !== this.defaultFont.name) out.name = f.name;
    const size = this.fontSize(f);
    if (size !== undefined && size !== this.defaultFont.size) out.size = size;
    if (f.bold) out.bold = true;
    if (f.italic) out.italic = true;
    const underline = normalizeUnderline(f.underline);
    if (underline) out.underline = underline;
    if (f.strike) out.strike = true;
    const color = this.textColor(f.color);
    if (color) out.color = color;
    if (f.vertAlign === 'superscript' || f.vertAlign === 'subscript') out.vertAlign = f.vertAlign;
    return Object.keys(out).length ? out : undefined;
  }

  /**
   * Solid fills -> `fill`. Patterns -> `pattern` plus `fill` = the perceived colour
   * (fg blended over bg by pattern density). Gradients -> `gradient` (angle in
   * Excel degrees: 0 = left to right, 90 = top to bottom). In dxf fills the
   * colour of a solid fill lives in bgColor; a cell's solid fill paints fgColor
   * only, and like Excel a missing fgColor is the default foreground (black).
   */
  private applyFill(out: CellStyle, fill: XFill | undefined, dxf: boolean): void {
    if (!fill) {
      return;
    }
    if (fill.type === 'gradient') {
      const stops = (fill.stops ?? [])
        .map((stop) => ({
          position: Math.max(0, Math.min(1, typeof stop.position === 'number' ? stop.position : 0)),
          color: this.color(stop.color) ?? '#FFFFFF',
        }))
        .sort((a, b) => a.position - b.position);
      if (stops.length) {
        out.gradient = { angle: fill.gradient === 'path' ? 0 : fill.degree ?? 0, stops };
      }
      return;
    }
    const pattern = fill.pattern ?? (dxf ? 'solid' : 'none');
    if (pattern === 'none') {
      return;
    }
    const fg = isAutomaticColor(fill.fgColor) ? undefined : this.color(fill.fgColor);
    const bg = isAutomaticColor(fill.bgColor) ? undefined : this.color(fill.bgColor);
    if (pattern === 'solid') {
      // Cell fill: indexed 64 is black, 65 white, auto the window colour (= no fill); bgColor is ignored.
      // ExcelJS drops <fgColor auto="1"/>, so the raw styles.xml tells auto apart from a missing fgColor.
      let color: string | undefined;
      if (dxf) {
        color = bg ?? fg;
      } else if (fill.fgColor) {
        color = this.color(fill.fgColor);
      } else if (!this.hasAutoForeground(fill)) {
        color = SYSTEM_COLORS[0];
      }
      if (color) out.fill = color;
      return;
    }
    out.fill = blendColors(fg ?? '#000000', bg ?? '#FFFFFF', PATTERN_DENSITY[pattern] ?? 0.5);
    out.pattern = { type: pattern };
    if (fg) out.pattern.fg = fg;
    if (bg) out.pattern.bg = bg;
  }

  /** Was this (loaded, still colourless) fill written with <fgColor auto="1"/>? */
  private hasAutoForeground(fill: XFill): boolean {
    const index = this.extras?.fillIndex.get(fill);
    return index !== undefined && !!this.extras?.raw?.autoFgFills.has(index);
  }

  private border(border: XBorder | undefined): CellStyle['border'] | undefined {
    if (!border) {
      return undefined;
    }
    const out: NonNullable<CellStyle['border']> = {};
    const top = this.edge(border.top, border.color);
    if (top) out.top = top;
    const right = this.edge(border.right, border.color);
    if (right) out.right = right;
    const bottom = this.edge(border.bottom, border.color);
    if (bottom) out.bottom = bottom;
    const left = this.edge(border.left, border.color);
    if (left) out.left = left;
    const diagonal = border.diagonal;
    if (diagonal && (diagonal.up || diagonal.down)) {
      const edge = this.edge(diagonal, border.color);
      if (edge) {
        out.diagonal = { ...edge };
        if (diagonal.up) out.diagonal.up = true;
        if (diagonal.down) out.diagonal.down = true;
      }
    }
    return Object.keys(out).length ? out : undefined;
  }

  private edge(edge: XBorderEdge | undefined, fallback: XColor | undefined): BorderEdge | undefined {
    if (!edge || !edge.style || !BORDER_STYLES.has(edge.style)) {
      return undefined;
    }
    const out: BorderEdge = { style: edge.style as BorderLineStyle };
    const raw = edge.color ?? fallback;
    const color = isAutomaticColor(raw) ? undefined : this.color(raw);
    if (color) out.color = color;
    return out;
  }

  private alignment(a: XAlignment | undefined): CellStyle['align'] | undefined {
    if (!a) {
      return undefined;
    }
    const out: NonNullable<CellStyle['align']> = {};
    if (a.horizontal && H_ALIGN.has(a.horizontal)) out.h = a.horizontal as NonNullable<CellStyle['align']>['h'];
    const vertical = a.vertical === 'center' ? 'middle' : a.vertical;
    if (vertical && V_ALIGN.has(vertical)) out.v = vertical as NonNullable<CellStyle['align']>['v'];
    if (a.wrapText) out.wrap = true;
    if (a.shrinkToFit) out.shrink = true;
    if (typeof a.indent === 'number' && a.indent > 0) out.indent = Math.round(a.indent);
    if (a.textRotation === 'vertical' || a.textRotation === 255) {
      out.rotation = 'vertical';
    } else if (typeof a.textRotation === 'number' && a.textRotation !== 0) {
      const deg = a.textRotation > 90 && a.textRotation <= 180 ? 90 - a.textRotation : a.textRotation;
      if (deg >= -90 && deg <= 90 && deg !== 0) out.rotation = deg;
    }
    return Object.keys(out).length ? out : undefined;
  }

  /**
   * Effective number format of a cell style. Uses the original code from
   * styles.xml when the style object can be traced to its xf (ExcelJS strips
   * backslash escapes and maps builtin ids to non-Excel codes).
   */
  numFmt(style: XStyle): string {
    const extras = this.extras;
    const raw = extras?.raw;
    if (raw) {
      const loaded = extras.xfByStyle.get(style);
      // Only while the shared style object still carries the format it was loaded with.
      const id = loaded && loaded.numFmt === style.numFmt ? raw.xfNumFmtIds[loaded.xf] : undefined;
      if (id !== undefined) {
        const code = raw.numFmts.get(id) ?? BUILTIN_NUMFMTS[id];
        if (code) {
          return code;
        }
      }
    }
    const code = typeof style.numFmt === 'string' ? style.numFmt : style.numFmt?.formatCode;
    if (!code) {
      return 'General';
    }
    return extras?.strippedFormats.get(code) ?? EXCELJS_FORMAT_FIXES[code] ?? code;
  }
}

function normalizeUnderline(value: boolean | string | undefined): FontStyle['underline'] | undefined {
  if (value === true || value === 'single') return 'single';
  if (value === 'double' || value === 'singleAccounting' || value === 'doubleAccounting') return value;
  return undefined;
}

// ===== LOADING =====

/** Parts of xl/styles.xml that ExcelJS loses or alters while loading. */
interface RawStyles {
  /** Explicit <numFmt> codes by id (XML-decoded, escapes intact). */
  numFmts: Map<number, string>;
  /** numFmtId of each cellXfs entry. */
  xfNumFmtIds: number[];
  /** <fonts> entries (exact sizes; fonts[0] is the Normal style font). */
  fonts: { name?: string; size?: number }[];
  /** Indexes of <fills> entries with <fgColor auto="1"/> (ExcelJS drops the auto colour). */
  autoFgFills: Set<number>;
  /** Custom <indexedColors> palette. */
  indexedColors?: string[];
}

/**
 * A workbook that formulas reference as [n] (xl/externalLinks): where Excel's
 * formula bar says it is, e.g. dir 'C:\Data\' + file 'Book.xlsx'.
 */
interface ExternalBook {
  /** Folder or URL prefix with its trailing separator; '' when unknown. */
  dir: string;
  file: string;
}

/** An array formula of a loaded sheet: master cell, range, and whether it is an Excel 365 dynamic array. */
interface ArrayFormulaInfo {
  r: number;
  c: number;
  range: Range;
  dynamic: boolean;
}

/** Side-channel information captured by loadWorkbook for workbookToModel. */
interface WorkbookExtras {
  raw?: RawStyles;
  /** External workbooks by formula index - 1 (undefined: DDE / OLE links, unreadable parts). */
  externalBooks?: (ExternalBook | undefined)[];
  /** Array formulas per worksheet id (absent: none in the file). */
  arrayFormulas?: Map<number, ArrayFormulaInfo[]>;
  /** Threaded-comment text per worksheet id and cell (row * MAX_COLS + col), shown instead of the legacy note. */
  threadedNotes?: Map<number, Map<number, string>>;
  /** ExcelJS cell-style object -> cellXfs index (objects are shared per xf) + numFmt at load time. */
  xfByStyle: Map<object, { xf: number; numFmt: unknown }>;
  /** ExcelJS font object -> <fonts> index + size at load time. */
  fontIndex: Map<object, { index: number; size: unknown }>;
  /** ExcelJS fill object (shared per <fills> entry) -> its index. */
  fillIndex: Map<object, number>;
  /** ExcelJS' backslash-stripped format code -> original code. */
  strippedFormats: Map<string, string>;
}

const workbookExtras = new WeakMap<object, WorkbookExtras>();

/** Shape of the intermediate model handed to XLSX#reconcile (xlsx/xlsx.js). */
interface XLoadModel {
  styles?: unknown;
  worksheets?: { id?: number; sheetNo?: string | number; hyperlinks?: XHyperlinkModel[]; rows?: { cells?: XCellLoadModel[] }[] }[];
  worksheetRels?: Record<string, XRelationship[] | undefined>;
  /** Parsed comments parts by relationship target ('../comments1.xml'). */
  comments?: Record<string, { comments?: XCommentModel[] } | undefined>;
}

interface XCommentModel {
  ref?: string;
  note?: unknown;
}

interface XHyperlinkModel {
  address?: string;
  rId?: string;
  /** The `location` attribute (internal target), see hyperlink-xform.js. */
  target?: string;
  tooltip?: string;
}

interface XRelationship {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

const HYPERLINK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const THREADED_COMMENT_REL_TYPE = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const PERSON_REL_TYPE = 'http://schemas.microsoft.com/office/2017/10/relationships/person';

/** Optional facts about the file loadWorkbook cannot know from its bytes. */
export interface LoadOptions {
  /**
   * Folder the file is in (e.g. 'C:\Data'): formulas that reference another workbook by a relative path show it in
   * this folder, as Excel does when the file does not record where it was saved (x15ac:absPath).
   */
  folder?: string;
}

/**
 * Loads an .xlsx with ExcelJS. Things ExcelJS drops are recovered on the way:
 * internal hyperlinks (<hyperlink location="Sheet2!A1"> without r:id become
 * '#Sheet2!A1' links, the same form ExcelJS writes), the raw styles.xml data
 * needed for exact number formats, font sizes and custom indexed palettes, and
 * the original sqref of data validations. Sheet names ExcelJS' API refuses
 * ("History", see patchSheetNameCheck) are taken as they are in the file,
 * notes are kept when written without runs or on cells absent from <sheetData>,
 * threaded comments are read (readThreadedComments), and the _xHHHH_ escapes
 * of inline strings and formula results are decoded (patchCellParser).
 */
export async function loadWorkbook(data: Uint8Array, options?: LoadOptions): Promise<ExcelJS.Workbook> {
  const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  assertZipContainer(buffer);
  patchValidationParser();
  patchSheetNameCheck();
  patchCommentParser();
  patchCellParser();
  const rawParts = readRawParts(buffer, options?.folder);
  const workbook = new ExcelJS.Workbook();
  const captured: CapturedLoad = { arraysParsedBefore: arrayFormulasParsed };
  const restore = hookReconcile(workbook, captured);
  loadingWorkbooks.add(workbook);
  try {
    // ExcelJS types its input as an ArrayBuffer-like "Buffer"; at runtime it takes a Node Buffer.
    await workbook.xlsx.load(buffer as unknown as Parameters<typeof workbook.xlsx.load>[0]);
  } catch (err) {
    throw workbookReadError(err);
  } finally {
    loadingWorkbooks.delete(workbook);
    restore();
  }
  const { styles, externalBooks, hasWorkbookPart, dynamicArrayCells, threadedNotes } = await rawParts;
  if (hasWorkbookPart === false) throw new Error(NO_WORKBOOK_PART_MESSAGE);
  try {
    attachOrphanNotes(workbook, captured.comments ?? []);
  } catch {
    // best effort only: the workbook itself loaded
  }
  const arrayFormulas = captured.arrays ? arrayFormulasById(captured.arrays, dynamicArrayCells) : undefined;
  workbookExtras.set(workbook, { ...buildExtras(styles, captured.styles), externalBooks, arrayFormulas, threadedNotes });
  return workbook;
}

/** What the reconcile hook captures from ExcelJS' intermediate load model. */
interface CapturedLoad {
  /** StylesXform (see buildExtras). */
  styles?: unknown;
  /**
   * Each worksheet's parsed notes, whether or not ExcelJS attaches them to a cell,
   * with the worksheet's load model (its `id` is only assigned during reconcile).
   */
  comments?: { sheet: { id?: number }; comments: XCommentModel[] }[];
  /** arrayFormulasParsed when this load started (unchanged at reconcile: the file has no array formula). */
  arraysParsedBefore: number;
  /** Each worksheet's array-formula cells (only collected when the file has some). */
  arrays?: { sheet: { id?: number }; cells: XCellLoadModel[] }[];
}

/** Array formulas (<f t="array">) parsed so far by any load (see patchCellParser). */
let arrayFormulasParsed = 0;
let cellParserPatched = false;

/**
 * ExcelJS drops the `cm` attribute of <c> (the cell-metadata record that marks an Excel 365 dynamic-array formula,
 * which the formula bar shows without the {} of a legacy array formula), and it decodes the _xHHHH_ escapes of
 * cell text (see decodeXstring) only in shared strings and rich-text runs. Its cell parser is wrapped once, for
 * good, to keep `cm` on the parsed cell model, to count array formulas (so the load only looks for them when there
 * are), and to decode the escapes of plain inline strings and string formula results (t="inlineStr" / t="str":
 * Excel stores ="x"&CHAR(13)&CHAR(10)&"y" as <v>x_x000D_\ny</v>).
 */
function patchCellParser(): void {
  const proto = CellXform?.prototype;
  const parseOpen = proto?.parseOpen;
  const parseClose = proto?.parseClose;
  if (cellParserPatched || !proto || typeof parseOpen !== 'function' || typeof parseClose !== 'function') {
    return;
  }
  cellParserPatched = true;
  proto.parseClose = function parseCloseDecodingText(this: XCellParser, name: string): boolean {
    const handled = parseClose.call(this, name);
    const model = this.model;
    if (name === 'c' && model && (this.t === 'str' || this.t === 'inlineStr')) {
      if (typeof model.value === 'string') model.value = decodeXstring(model.value);
      if (typeof model.result === 'string') model.result = decodeXstring(model.result);
    }
    return handled;
  };
  proto.parseOpen = function parseOpenWithMetadata(this: XCellParser, node): boolean {
    const nested = this.parser !== undefined;
    const handled = parseOpen.call(this, node);
    if (!nested && this.model) {
      if (node.name === 'c' && node.attributes?.cm) {
        this.model.cm = node.attributes.cm;
      } else if (node.name === 'f' && node.attributes?.t === 'array') {
        arrayFormulasParsed++;
      }
    }
    return handled;
  };
}

/**
 * Cell text as Excel reads it from the file (an ST_Xstring): each _xHHHH_ escape is the UTF-16 code unit HHHH, in
 * one left-to-right pass ('_x005F_x000D_' is the literal text '_x000D_'). Excel 16 takes the hex digits in either
 * case; ExcelJS' shared-string reader only upper case.
 */
function decodeXstring(text: string): string {
  return text.includes('_x') ? text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16))) : text;
}

/** The array-formula cells of every worksheet of the load model (a walk over all parsed cells). */
function collectArrayFormulas(model: XLoadModel): NonNullable<CapturedLoad['arrays']> {
  const out: NonNullable<CapturedLoad['arrays']> = [];
  for (const sheet of model.worksheets ?? []) {
    const cells: XCellLoadModel[] = [];
    for (const row of sheet?.rows ?? []) {
      for (const cell of row?.cells ?? []) {
        if (cell?.shareType === 'array' && typeof cell.address === 'string') cells.push(cell);
      }
    }
    if (cells.length && sheet) out.push({ sheet, cells });
  }
  return out;
}

/** Captured array-formula cells -> ArrayFormulaInfo per worksheet id; `cm` in `dynamicCells` = dynamic array. */
function arrayFormulasById(
  sheets: NonNullable<CapturedLoad['arrays']>,
  dynamicCells: ReadonlySet<number> | undefined,
): Map<number, ArrayFormulaInfo[]> {
  const out = new Map<number, ArrayFormulaInfo[]>();
  for (const { sheet, cells } of sheets) {
    if (typeof sheet.id !== 'number') continue;
    const list: ArrayFormulaInfo[] = [];
    for (const cell of cells) {
      const at = decodeCell(cell.address ?? '');
      if (!at) continue;
      const range = (cell.ref ? decodeRange(cell.ref) : undefined) ?? { r0: at.r, c0: at.c, r1: at.r, c1: at.c };
      list.push({ r: at.r, c: at.c, range, dynamic: cell.cm !== undefined && !!dynamicCells?.has(Number(cell.cm)) });
    }
    if (list.length) out.set(sheet.id, list);
  }
  return out;
}

/**
 * The notes of every worksheet, found the way ExcelJS' worksheet reconcile finds
 * them (comments relationship -> parsed part). The note objects are shared, so
 * they also carry what reconcile merges in from the VML drawing.
 */
function sheetComments(model: XLoadModel): NonNullable<CapturedLoad['comments']> {
  const out: NonNullable<CapturedLoad['comments']> = [];
  for (const sheet of model.worksheets ?? []) {
    const rel = model.worksheetRels?.[String(sheet?.sheetNo)]?.find((r) => r?.Type === COMMENTS_REL_TYPE);
    const comments = rel ? model.comments?.[rel.Target]?.comments : undefined;
    if (Array.isArray(comments) && comments.length) out.push({ sheet, comments });
  }
  return out;
}

/**
 * ExcelJS only attaches a note to a cell present in <sheetData>; Excel keeps a
 * note on an empty, unstyled cell without writing the cell. Such notes get their
 * cell (an empty one, as Excel shows it) after loading.
 */
function attachOrphanNotes(workbook: ExcelJS.Workbook, sheets: NonNullable<CapturedLoad['comments']>): void {
  for (const { sheet, comments } of sheets) {
    const id = sheet.id;
    const ws = typeof id === 'number' ? (workbook.getWorksheet(id) as unknown as (XWorksheet & { getCell(ref: string): { note: unknown } }) | undefined) : undefined;
    if (!ws) {
      continue;
    }
    for (const comment of comments) {
      const at = typeof comment?.ref === 'string' ? decodeCell(comment.ref) : undefined;
      if (at && comment.note && !ws._rows[at.r]?._cells[at.c]?._comment) {
        ws.getCell(`${indexToCol(at.c)}${at.r + 1}`).note = comment.note;
      }
    }
  }
}

let commentParserPatched = false;

/**
 * ExcelJS reads a note's text only from rich-text runs (<text><r><t>...); a
 * note written as a plain <text><t>...</t></text> - openpyxl's notes, Excel's
 * threaded-comment fallbacks - came out empty and was dropped. The <comment>
 * parser is wrapped once, for good, to read such a <t> as one plain run
 * (phonetic <rPh> text stays ignored, as before).
 */
function patchCommentParser(): void {
  const proto = CommentXform?.prototype;
  const { parseOpen, parseClose } = proto ?? {};
  if (commentParserPatched || !proto || !TextXform || typeof parseOpen !== 'function' || typeof parseClose !== 'function') {
    return;
  }
  commentParserPatched = true;
  const PlainText = TextXform;
  proto.parseOpen = function parseOpenPlainText(this: XCommentParser, node: { name: string }): boolean {
    if (!this.parser) {
      if (node.name === 'comment') {
        this.fvPhonetic = false;
      } else if (node.name === 'rPh') {
        this.fvPhonetic = true;
      } else if (node.name === 't' && !this.fvPhonetic && Array.isArray(this.model?.note?.texts)) {
        this.parser = this.fvPlainText ??= new PlainText();
        this.parser.parseOpen(node);
        return true;
      }
    }
    return parseOpen.call(this, node);
  };
  proto.parseClose = function parseClosePlainText(this: XCommentParser, name: string): boolean {
    if (name === 't' && this.parser && this.parser === this.fvPlainText) {
      this.model?.note?.texts?.push({ text: this.parser.model });
      this.parser = undefined;
      return true;
    }
    if (name === 'rPh' && !this.parser) {
      this.fvPhonetic = false;
    }
    return parseClose.call(this, name);
  };
}

/** Turns "not a zip" into a readable error (encrypted workbooks and .xls are OLE2 files). */
function assertZipContainer(buffer: Uint8Array): void {
  if (buffer.length >= 4 && buffer[0] === 0xd0 && buffer[1] === 0xcf && buffer[2] === 0x11 && buffer[3] === 0xe0) {
    throw new Error('This file is an OLE2 compound document: a password-protected workbook or a legacy .xls file. Only unencrypted .xlsx files can be opened.');
  }
  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    throw new Error('This file is not a valid .xlsx workbook (it is not a zip package).');
  }
}

/** jszip (and the inflate code it uses) on a truncated or damaged package. */
const ZIP_DAMAGED_RE =
  /corrupted zip|central dir|end of data reached|invalid signature|size mismatch|crc32|missing \d+ bytes|is this a zip file|invalid (?:stored block|block type|distance|code|literal)|incorrect header check|unexpected end of (?:file|data)/i;
const ZIP_ENCRYPTED_RE = /encrypted zip/i;
const ZIP_COMPRESSION_RE = /compression/i;
/** sax (ExcelJS' XML parser) reports "<reason>\nLine: n\nColumn: n\nChar: c". */
const XML_ERROR_RE = /\nLine: -?\d+\s*\nColumn: -?\d+|unclosed root tag|non-whitespace before first tag|text data outside of root node/i;

/**
 * A zip package that starts like an .xlsx but cannot be read gets a readable reason, like non-zip input does
 * (assertZipContainer); the library's own text ("Corrupted zip: can't find end of central directory", "Cannot read
 * properties of undefined (reading 'sheets')") stays available as `cause` for the error details. Errors that are
 * already readable (ours, ExcelJS' explicit checks) are returned unchanged.
 */
function workbookReadError(err: unknown): unknown {
  const text = err instanceof Error ? err.message : String(err);
  let reason: string | undefined;
  if (ZIP_ENCRYPTED_RE.test(text)) {
    reason = 'This file is not a valid .xlsx workbook: its zip package is encrypted.';
  } else if (ZIP_DAMAGED_RE.test(text)) {
    reason = 'This file is not a valid .xlsx workbook: its zip package is damaged (truncated or corrupted), so it cannot be read.';
  } else if (ZIP_COMPRESSION_RE.test(text)) {
    reason = 'This file is not a valid .xlsx workbook: its zip package uses a compression method that cannot be read.';
  } else if (XML_ERROR_RE.test(text) || err instanceof TypeError || err instanceof RangeError) {
    // ExcelJS has no checks of its own: damaged XML surfaces as a parser error or as a TypeError further on.
    reason = 'This workbook could not be read: part of its content is damaged or is not valid spreadsheet XML.';
  }
  return reason === undefined ? err : new Error(reason, { cause: err });
}

/** A readable zip without xl/workbook.xml: some other zip-based file (e.g. a .docx) renamed to .xlsx. */
const NO_WORKBOOK_PART_MESSAGE =
  'This file is not a valid .xlsx workbook: its zip package has no workbook part (it may be another kind of file, such as a .docx, renamed to .xlsx).';

/** A loaded <dataValidation> rule with the sqref it was read from. */
interface LoadedValidation {
  sqref: string;
  rule: XDataValidation;
}

/** Sheet dataValidations.model (the object ExcelJS keeps per sheet) -> its rules in file order. */
const loadedValidations = new WeakMap<object, LoadedValidation[]>();
let validationParserPatched = false;

/**
 * ExcelJS expands every <dataValidation sqref> into one model key per cell (a
 * whole column is 1,048,576 keys: ~0.8 s and ~75 MB while loading) and reads
 * 'B:B' / '2:2' as 'A1' / 'A2'. Its parser is wrapped once, for good, to record
 * each rule with its original sqref (the grid model never walks the keys) and to
 * keep a part bigger than VALIDATION_EXPAND_CAP cells as ONE range key such as
 * 'B1:B1048576' - the form ExcelJS' writer saves back as that sqref. Smaller
 * parts are still expanded per cell by ExcelJS.
 */
function patchValidationParser(): void {
  const proto = DataValidationsXform?.prototype;
  const original = proto?.parseClose;
  if (validationParserPatched || !proto || typeof original !== 'function') {
    return;
  }
  validationParserPatched = true;
  proto.parseClose = function parseCloseKeepingSqref(name: string): boolean {
    if (name !== 'dataValidation') {
      return original.call(this, name);
    }
    try {
      const { model, _address: sqref, _dataValidation: rule } = this;
      if (model && rule && typeof sqref === 'string') {
        let rules = loadedValidations.get(model);
        if (!rules) loadedValidations.set(model, (rules = []));
        rules.push({ sqref, rule });
        const keys = model as Record<string, XDataValidation>;
        const expand: string[] = [];
        for (const part of sqref.split(/\s+/)) {
          const range = part ? decodeRange(part) : undefined;
          if (range && (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1) > VALIDATION_EXPAND_CAP) {
            keys[`${indexToCol(range.c0)}${range.r0 + 1}:${indexToCol(range.c1)}${range.r1 + 1}`] = rule;
          } else if (part) {
            expand.push(part);
          }
        }
        if (!expand.length) {
          // what ExcelJS' parseClose does besides adding the keys
          if (!Array.isArray(rule.formulae) || !rule.formulae.length) {
            delete rule.formulae;
            delete rule.operator;
          }
          return true;
        }
        this._address = expand.join(' ');
      }
    } catch {
      // best effort only: never break loading
    }
    return original.call(this, name);
  };
}

/** Workbooks that loadWorkbook is loading right now (see patchSheetNameCheck). */
const loadingWorkbooks = new WeakSet<object>();
let sheetNameCheckPatched = false;

/**
 * ExcelJS' Worksheet name setter throws on names that are in real files: the
 * reserved "History" (pandas / openpyxl write it; Excel opens such a file
 * without complaint), names over 31 characters and case-insensitive duplicates
 * (Excel repairs those) - and the throw fails the whole workbook. The setter is
 * wrapped once, for good: while loadWorkbook loads a workbook, a non-empty name
 * from the file is kept as it is; every other call keeps ExcelJS' checks.
 */
function patchSheetNameCheck(): void {
  const proto = WorksheetClass?.prototype;
  const descriptor = proto ? Object.getOwnPropertyDescriptor(proto, 'name') : undefined;
  const original = descriptor?.set;
  if (sheetNameCheckPatched || !proto || !descriptor || typeof original !== 'function') {
    return;
  }
  sheetNameCheckPatched = true;
  Object.defineProperty(proto, 'name', {
    ...descriptor,
    set(this: { _workbook?: object; _name?: string }, name: unknown): void {
      if (typeof name === 'string' && name && this._workbook && loadingWorkbooks.has(this._workbook)) {
        this._name = name;
      } else {
        original.call(this, name);
      }
    },
  });
}

/** Wraps XLSX#reconcile on this workbook's XLSX instance; returns the undo. */
function hookReconcile(workbook: ExcelJS.Workbook, captured: CapturedLoad): () => void {
  const xlsx = workbook.xlsx as unknown as { reconcile?: (model: XLoadModel, options: unknown) => void };
  const original = xlsx.reconcile;
  if (typeof original !== 'function') {
    return () => undefined;
  }
  const ownProperty = Object.prototype.hasOwnProperty.call(xlsx, 'reconcile');
  xlsx.reconcile = function reconcileWithSideChannel(this: unknown, model: XLoadModel, options: unknown): void {
    try {
      captured.styles = model.styles;
      captured.comments = sheetComments(model);
      injectInternalHyperlinks(model);
      if (arrayFormulasParsed !== captured.arraysParsedBefore) captured.arrays = collectArrayFormulas(model);
    } catch {
      // best effort only: never break loading
    }
    original.call(this, model, options);
  };
  return () => {
    if (ownProperty) {
      xlsx.reconcile = original;
    } else {
      delete xlsx.reconcile; // back to the prototype method
    }
  };
}

/**
 * ExcelJS only maps hyperlinks that have an r:id, to the relationship target
 * alone. A `location` (what follows '#': a sheet / cell, or the fragment Excel
 * splits off a URL such as https://example.com/page#section-2) gets a synthetic
 * relationship: '#<location>' without r:id, '<target>#<location>' with one.
 * Hyperlinks on ranges are expanded; ones whose r:id has no relationship are
 * dropped (ExcelJS would throw on them and fail the whole workbook).
 */
function injectInternalHyperlinks(model: XLoadModel): void {
  let counter = 0;
  for (const sheet of model.worksheets ?? []) {
    const links = sheet?.hyperlinks;
    if (!Array.isArray(links) || !links.length) {
      continue;
    }
    const key = String(sheet.sheetNo);
    const rels = model.worksheetRels ? (model.worksheetRels[key] ?? (model.worksheetRels[key] = [])) : undefined;
    const out: XHyperlinkModel[] = [];
    for (const link of links) {
      if (!link || typeof link.address !== 'string') {
        continue;
      }
      let rId = link.rId;
      const base = rId ? rels?.find((rel) => rel?.Id === rId) : undefined;
      if (rId && rels && !base) {
        continue;
      }
      // ExcelJS itself writes internal links as Target="#Sheet!A1" plus location="#Sheet!A1": already complete
      const location = typeof link.target === 'string' ? link.target.replace(/^#/, '') : '';
      if (location && rels && !base?.Target?.includes('#')) {
        rId = `rIdFileViewerLocation${++counter}`;
        rels.push({ Id: rId, Type: HYPERLINK_REL_TYPE, Target: `${base?.Target ?? ''}#${location}`, TargetMode: base ? base.TargetMode : 'Internal' });
      }
      if (!rId) {
        continue;
      }
      const addresses = link.address.includes(':') ? expandAddresses(link.address, HYPERLINK_EXPAND_CAP) : [link.address];
      for (const address of addresses) {
        out.push({ ...link, address, rId });
      }
    }
    sheet.hyperlinks = out;
  }
}

/**
 * Package parts read straight from the zip: styles.xml details, the external workbooks, the cell-metadata records
 * of dynamic arrays and the threaded comments. `hasWorkbookPart` is undefined when the zip itself cannot be read
 * (ExcelJS reports that).
 */
async function readRawParts(buffer: Uint8Array, folder?: string): Promise<{
  styles?: RawStyles;
  externalBooks?: (ExternalBook | undefined)[];
  dynamicArrayCells?: Set<number>;
  threadedNotes?: Map<number, Map<number, string>>;
  hasWorkbookPart?: boolean;
}> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return {};
  }
  // ExcelJS reads the workbook from this part name only (a leading '/' is tolerated).
  if (!zip.file('xl/workbook.xml') && !zip.file('/xl/workbook.xml')) return { hasWorkbookPart: false };
  const read = async <T>(parse: () => Promise<T>): Promise<T | undefined> => {
    try {
      return await parse();
    } catch {
      return undefined; // best effort only: ExcelJS reports broken packages itself
    }
  };
  const [styles, externalBooks, dynamicArrayCells, threadedNotes] = await Promise.all([
    read(async () => {
      const file = zip.file('xl/styles.xml');
      return file ? parseRawStyles(await file.async('string')) : undefined;
    }),
    read(() => readExternalBooks(zip, folder)),
    read(async () => {
      const file = zip.file('xl/metadata.xml');
      return file ? dynamicArrayMetadata(await file.async('string')) : undefined;
    }),
    read(() => readThreadedComments(zip)),
  ]);
  return { styles, externalBooks, dynamicArrayCells, threadedNotes, hasWorkbookPart: true };
}

/**
 * xl/metadata.xml -> the 1-based <cellMetadata> records (`cm` of a cell) that mark a dynamic-array formula: a record
 * whose <rc t v> points at the XLDAPR metadata type and at a futureMetadata entry with fDynamic="1".
 */
function dynamicArrayMetadata(xml: string): Set<number> {
  const out = new Set<number>();
  const types = [...(xmlSection(xml, 'metadataTypes') ?? '').matchAll(/<(?:\w+:)?metadataType\b([^>]*?)\/?>/g)].map(
    (m) => parseXmlAttributes(m[1]).name,
  );
  const blocks = (section: string): string[] => [...section.matchAll(/<(?:\w+:)?bk\b[^>]*>([\s\S]*?)<\/(?:\w+:)?bk>/g)].map((m) => m[1]);
  const future = /<(?:\w+:)?futureMetadata\b[^>]*\bname\s*=\s*["']XLDAPR["'][^>]*>([\s\S]*?)<\/(?:\w+:)?futureMetadata>/.exec(xml);
  const dynamic = future ? blocks(future[1]).map((bk) => /\bfDynamic\s*=\s*["'](?:1|true)["']/.test(bk)) : [];
  blocks(xmlSection(xml, 'cellMetadata') ?? '').forEach((bk, i) => {
    for (const rc of bk.matchAll(/<(?:\w+:)?rc\b([^>]*?)\/?>/g)) {
      const { t, v } = parseXmlAttributes(rc[1]);
      if (types[Number(t) - 1] === 'XLDAPR' && dynamic[Number(v)]) out.add(i + 1);
    }
  });
  return out;
}

/**
 * The workbooks behind formula references like [1]Sheet1!B2, in
 * <externalReferences> order. Like Excel, the location is the externalBook's
 * absoluteUrl when present (Excel 2021+), else its target - a relative one taken
 * from the folder the workbook was saved in (x15ac:absPath), when recorded, else
 * from the folder it is in now (`folder`, LoadOptions).
 */
async function readExternalBooks(zip: JSZip, folder?: string): Promise<(ExternalBook | undefined)[] | undefined> {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const refs = workbookXml ? xmlSection(workbookXml, 'externalReferences') : undefined;
  if (!workbookXml || refs === undefined) {
    return undefined;
  }
  const absPathTag = /<(?:\w+:)?absPath\b([^>]*?)\/?>/.exec(workbookXml);
  const absPath = (absPathTag ? parseXmlAttributes(absPathTag[1]).url || undefined : undefined) ?? (folder || undefined);
  const workbookRels = await readRelationships(zip, 'xl/workbook.xml');
  const out: (ExternalBook | undefined)[] = [];
  for (const m of refs.matchAll(/<(?:\w+:)?externalReference\b([^>]*?)\/?>/g)) {
    const part = workbookRels.get(relationshipId(m[1]) ?? '');
    const path = part ? resolvePartPath('xl/workbook.xml', part) : undefined;
    const xml = path ? await zip.file(path)?.async('string') : undefined;
    const book = xml ? /<(?:\w+:)?externalBook\b([^>]*)>/.exec(xml) : undefined; // absent for DDE / OLE links
    if (!path || !xml || !book) {
      out.push(undefined);
      continue;
    }
    const rels = await readRelationships(zip, path);
    const absolute = /<(?:\w+:)?absoluteUrl\b([^>]*?)\/?>/.exec(xml);
    const target = (absolute ? rels.get(relationshipId(absolute[1]) ?? '') : undefined) ?? rels.get(relationshipId(book[1]) ?? '');
    out.push(target ? externalBookLocation(target, absPath) : undefined);
  }
  return out;
}

/** The r:id attribute (any prefix) of an element's attribute text. */
function relationshipId(attributes: string): string | undefined {
  const attrs = parseXmlAttributes(attributes);
  const key = Object.keys(attrs).find((name) => /^\w+:id$/.test(name));
  return key ? attrs[key] : undefined;
}

/** Relationship id -> target of a part's .rels (targets as written); only relationships of `type` when given. */
async function readRelationships(zip: JSZip, partPath: string, type?: string): Promise<Map<string, string>> {
  const slash = partPath.lastIndexOf('/');
  const xml = await zip.file(`${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`)?.async('string');
  const out = new Map<string, string>();
  for (const m of xml?.matchAll(/<(?:\w+:)?Relationship\b([^>]*?)\/?>/g) ?? []) {
    const attrs = parseXmlAttributes(m[1]);
    if (attrs.Id && attrs.Target !== undefined && (type === undefined || attrs.Type === type)) out.set(attrs.Id, attrs.Target);
  }
  return out;
}

/**
 * Threaded comments (Excel 365) as note text, per worksheet id (sheetId) and cell (row * MAX_COLS + col): one
 * 'Author: text' line per comment, the thread's first comment, then its replies. Excel also writes a legacy note for
 * each thread whose text only tells older versions about it ("[Threaded comment] Your version of Excel allows you to
 * read this threaded comment; ..."); like Excel 365, the grid shows the thread instead.
 */
async function readThreadedComments(zip: JSZip): Promise<Map<number, Map<number, string>> | undefined> {
  if (!Object.keys(zip.files).some((name) => /^\/?xl\/threadedComments\//i.test(name))) {
    return undefined;
  }
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!workbookXml) {
    return undefined;
  }
  const persons = new Map<string, string>();
  for (const target of (await readRelationships(zip, 'xl/workbook.xml', PERSON_REL_TYPE)).values()) {
    const xml = await zip.file(resolvePartPath('xl/workbook.xml', target))?.async('string');
    for (const m of xml?.matchAll(/<(?:\w+:)?person\b([^>]*?)\/?>/g) ?? []) {
      const { id, displayName } = parseXmlAttributes(m[1]);
      if (id && displayName) persons.set(id, displayName);
    }
  }
  const workbookRels = await readRelationships(zip, 'xl/workbook.xml');
  const out = new Map<number, Map<number, string>>();
  for (const m of (xmlSection(workbookXml, 'sheets') ?? '').matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)) {
    const sheetId = parseInt(parseXmlAttributes(m[1]).sheetId ?? '', 10);
    const target = workbookRels.get(relationshipId(m[1]) ?? '');
    if (!Number.isFinite(sheetId) || !target) continue;
    const sheetPath = resolvePartPath('xl/workbook.xml', target);
    const threads = new Map<number, { root: boolean; line: string }[]>();
    for (const part of (await readRelationships(zip, sheetPath, THREADED_COMMENT_REL_TYPE)).values()) {
      const xml = await zip.file(resolvePartPath(sheetPath, part))?.async('string');
      for (const c of xml?.matchAll(/<(?:\w+:)?threadedComment\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?threadedComment>)/g) ?? []) {
        const attrs = parseXmlAttributes(c[1]);
        const at = decodeCell(attrs.ref ?? '');
        if (!at) continue;
        // line breaks normalized as an XML parser does (Excel writes CR LF inside the text)
        const raw = /<(?:\w+:)?text\b[^>]*>([\s\S]*?)<\/(?:\w+:)?text>/.exec(c[2] ?? '')?.[1] ?? '';
        const text = decodeXmlEntities(raw.replace(/\r\n?/g, '\n'));
        const author = persons.get(attrs.personId ?? '');
        const key = at.r * MAX_COLS + at.c;
        let thread = threads.get(key);
        if (!thread) threads.set(key, (thread = []));
        thread.push({ root: !attrs.parentId, line: author ? `${author}: ${text}` : text });
      }
    }
    if (threads.size) {
      // the first comment of the thread first (Excel writes it first anyway), replies in file order
      const notes = new Map<number, string>();
      for (const [key, thread] of threads) {
        notes.set(key, [...thread.filter((c) => c.root), ...thread.filter((c) => !c.root)].map((c) => c.line).join('\n'));
      }
      out.set(sheetId, notes);
    }
  }
  return out.size ? out : undefined;
}

/** Zip path of a relationship target ('externalLinks/externalLink1.xml' from 'xl/workbook.xml'). */
function resolvePartPath(from: string, target: string): string {
  const parts = target.startsWith('/') ? [] : from.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else if (segment && segment !== '.') parts.push(segment);
  }
  return parts.join('/');
}

/**
 * External workbook target -> the folder and file name Excel shows: URLs
 * percent-decoded ('https://host/Shared Documents/'), file: URIs and relative
 * targets as Windows paths ('C:\Data\', '\\server\share\'), '..' resolved.
 */
function externalBookLocation(target: string, absPath: string | undefined): ExternalBook {
  let path = target;
  try {
    path = decodeURI(target);
  } catch {
    // keep malformed escapes as written
  }
  // file:///C:/x.xlsx -> C:/x.xlsx; file://server/share/x.xlsx -> //server/share/x.xlsx
  path = path.replace(/^file:\/\/\/(?=[A-Za-z]:)/i, '').replace(/^file:(?=\/\/)/i, '');
  if (!/^[A-Za-z][\w+.-]*:\/\//.test(path)) {
    if (absPath && !/^(?:[A-Za-z]:|[\\/])/.test(path)) {
      const sep = absPath.includes('\\') ? '\\' : '/';
      path = `${absPath.replace(/[\\/]+$/, '')}${sep}${path}`; // relative to the folder the workbook was saved in
    }
    if (/^(?:[A-Za-z]:|[\\/]{2})|\\/.test(path)) path = windowsPath(path);
  }
  const cut = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));
  return { dir: path.slice(0, cut + 1), file: path.slice(cut + 1) };
}

/**
 * Windows path with '\' separators and '.' / '..' segments resolved (a UNC '\\'
 * prefix kept; 'D:x' read as 'D:\x', as Excel does).
 */
function windowsPath(path: string): string {
  const out: string[] = [];
  for (const segment of path.replace(/^[\\/]+/, '').replace(/^([A-Za-z]:)(?![\\/])/, '$1\\').split(/[\\/]/)) {
    if (segment === '..') {
      if (out.length > 1) out.pop();
    } else if (segment && segment !== '.') {
      out.push(segment);
    }
  }
  return (/^[\\/]{2}/.test(path) ? '\\\\' : /^[\\/]/.test(path) ? '\\' : '') + out.join('\\');
}

function parseRawStyles(xml: string): RawStyles {
  const numFmts = new Map<number, string>();
  for (const m of xml.matchAll(/<(?:\w+:)?numFmt\b([^>]*?)\/?>/g)) {
    const attrs = parseXmlAttributes(m[1]);
    const id = parseInt(attrs.numFmtId ?? '', 10);
    if (Number.isFinite(id) && attrs.formatCode !== undefined) {
      numFmts.set(id, attrs.formatCode);
    }
  }
  const xfNumFmtIds: number[] = [];
  const cellXfs = xmlSection(xml, 'cellXfs');
  if (cellXfs !== undefined) {
    for (const m of cellXfs.matchAll(/<(?:\w+:)?xf\b([^>]*?)\/?>/g)) {
      xfNumFmtIds.push(parseInt(parseXmlAttributes(m[1]).numFmtId ?? '0', 10) || 0);
    }
  }
  const fonts: RawStyles['fonts'] = [];
  const fontsSection = xmlSection(xml, 'fonts');
  if (fontsSection !== undefined) {
    for (const m of fontsSection.matchAll(/<(?:\w+:)?font\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?font>)/g)) {
      const body = m[1] ?? '';
      const sz = /<(?:\w+:)?sz\b[^>]*?\bval\s*=\s*["']([^"']*)["']/.exec(body);
      const name = /<(?:\w+:)?name\b[^>]*?\bval\s*=\s*["']([^"']*)["']/.exec(body);
      const size = sz ? parseFloat(sz[1]) : NaN;
      fonts.push({ size: Number.isFinite(size) && size > 0 ? size : undefined, name: name ? decodeXmlEntities(name[1]) : undefined });
    }
  }
  const autoFgFills = new Set<number>();
  const fillsSection = xmlSection(xml, 'fills');
  if (fillsSection !== undefined) {
    let index = 0;
    for (const m of fillsSection.matchAll(/<(?:\w+:)?fill\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?fill>)/g)) {
      if (/<(?:\w+:)?fgColor\b[^>]*?\bauto\s*=\s*["'](?:1|true)["']/.test(m[1] ?? '')) autoFgFills.add(index);
      index++;
    }
  }
  let indexedColors: string[] | undefined;
  const indexed = xmlSection(xml, 'indexedColors');
  if (indexed !== undefined) {
    indexedColors = [];
    for (const m of indexed.matchAll(/<(?:\w+:)?rgbColor\b([^>]*?)\/?>/g)) {
      indexedColors.push(normalizeHex(parseXmlAttributes(m[1]).rgb ?? '') ?? '#000000');
    }
  }
  return { numFmts, xfNumFmtIds, fonts, autoFgFills, indexedColors };
}

function buildExtras(raw: RawStyles | undefined, stylesXform: unknown): WorkbookExtras {
  const extras: WorkbookExtras = { raw, xfByStyle: new Map(), fontIndex: new Map(), fillIndex: new Map(), strippedFormats: new Map() };
  // StylesXform (xlsx/xform/style/styles-xform.js): index.model[xfId] is the style
  // object shared by all cells with that xf; model.fonts[i] / model.fills[i] the
  // parsed fonts / fills (shared by every style that uses them).
  const sx = stylesXform as { index?: { model?: unknown }; model?: { fonts?: unknown; fills?: unknown } } | undefined;
  const models = sx?.index?.model;
  if (Array.isArray(models)) {
    models.forEach((model, xf) => {
      if (model && typeof model === 'object') extras.xfByStyle.set(model, { xf, numFmt: (model as XStyle).numFmt });
    });
  }
  const fonts = sx?.model?.fonts;
  if (Array.isArray(fonts)) {
    fonts.forEach((font, index) => {
      if (font && typeof font === 'object') extras.fontIndex.set(font, { index, size: (font as XFont).size });
    });
  }
  const fills = sx?.model?.fills;
  if (Array.isArray(fills)) {
    fills.forEach((fill, index) => {
      if (fill && typeof fill === 'object') extras.fillIndex.set(fill, index);
    });
  }
  raw?.numFmts.forEach((code) => {
    const stripped = code.replace(/\\(.)/g, '$1');
    if (!extras.strippedFormats.has(stripped)) extras.strippedFormats.set(stripped, code);
  });
  return extras;
}

/** Inner XML of the first <tag ...>...</tag> (any namespace prefix). */
function xmlSection(xml: string, tag: string): string | undefined {
  const m = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*?(?<!/)>([\\s\\S]*?)</(?:\\w+:)?${tag}>`).exec(xml);
  return m ? m[1] : undefined;
}

function parseXmlAttributes(source: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of source.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    out[m[1]] = decodeXmlEntities(m[2] ?? m[3] ?? '');
  }
  return out;
}

function decodeXmlEntities(text: string): string {
  if (!text.includes('&')) {
    return text;
  }
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, entity: string) => {
    switch (entity) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default:
        return String.fromCodePoint(entity[1] === 'x' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10));
    }
  });
}

// ===== GRID SOURCE =====

/** Anything the grid can display. Implemented by WorkbookModel (xlsx) and CsvGridModel (csv). */
export interface GridSource {
  getMeta(): WorkbookMeta;
  /** Rows [start, end) of sheet `sheet`. Rows with no cells may be omitted. New style-table entries discovered
   *  while converting are returned in `styles` (index, style) - the webview merges them into its table.
   *  Optional `c0` / `c1`: only the cells of columns c0..c1 (inclusive; e.g. copying one column of a large sheet). */
  getRows(sheet: number, start: number, end: number, c0?: number, c1?: number): { rows: RowData[]; styles: [number, CellStyle][] };
  getStats(sheet: number, ranges: Range[]): SelectionStats;
}

/** getRows' optional column window as [first, last] (inclusive, last may be Infinity); the whole row when absent. */
function columnWindow(c0: number | undefined, c1: number | undefined): [number, number] {
  const first = typeof c0 === 'number' && Number.isFinite(c0) ? Math.max(0, Math.floor(c0)) : 0;
  const last = typeof c1 === 'number' && Number.isFinite(c1) ? Math.floor(c1) : Infinity;
  return [first, last];
}

// ===== WORKBOOK MODEL =====

/** Per-sheet state of a WorkbookModel. */
interface SheetContext {
  ws: XWorksheet;
  index: number;
  /** Resolved list values per validation formula. */
  lists: Map<string, string[] | undefined>;
  /** Masters of the sheet's dynamic-array formulas (row * MAX_COLS + col), shown without braces. */
  dynamicArrays: Set<number>;
  /** Legacy array formulas spanning several cells, by first row: their other cells show the formula too. */
  arrayRanges: ArrayFormulaInfo[];
  /** Threaded comments by cell (row * MAX_COLS + col): shown instead of the cell's legacy note. */
  threadedNotes?: ReadonlyMap<number, string>;
}

/**
 * ExcelJS workbook -> grid model. Sheet structure (meta) is computed eagerly;
 * cells are converted lazily, one requested window at a time.
 *
 * Style table: entries are appended as rows are converted. getMeta() returns
 * everything known so far; every getRows reply carries all style entries its
 * rows reference, so a reply is self-contained even if the webview dropped an
 * earlier one.
 */
export class WorkbookModel implements GridSource {
  readonly workbook: ExcelJS.Workbook;
  private readonly wb: XWorkbook;
  private readonly palette: ThemePalette;
  private readonly date1904: boolean;
  private readonly defaultFont: { name: string; size: number };
  private readonly formatter: NumberFormatter;
  private readonly converter: StyleConverter;
  private readonly styleTable: StyleTable;
  private readonly sharedFormulas = new WeakMap<object, string>();
  private readonly imageUris = new Map<string, string | undefined>();
  private readonly links = new Map<string, string>();
  private readonly externalBooks: readonly (ExternalBook | undefined)[];
  /** Array formulas found while loading, per worksheet id (see loadWorkbook). */
  private readonly arrayFormulas: ReadonlyMap<number, ArrayFormulaInfo[]> | undefined;
  /** Threaded comments found while loading, per worksheet id (see readThreadedComments). */
  private readonly threadedNotes: ReadonlyMap<number, ReadonlyMap<number, string>> | undefined;
  private sheets: SheetContext[] = [];
  private meta: Omit<WorkbookMeta, 'styles'> = { sheets: [], activeSheet: 0, defaultFont: { name: 'Calibri', size: 11 }, date1904: false };

  constructor(workbook: ExcelJS.Workbook) {
    this.workbook = workbook;
    this.wb = workbook as unknown as XWorkbook;
    const extras = workbookExtras.get(workbook);
    this.date1904 = !!this.wb.properties?.date1904;
    this.palette = { ...parseThemeXml(pickTheme(this.wb._themes)), indexed: extras?.raw?.indexedColors };
    const normal = extras?.raw?.fonts[0];
    this.defaultFont = { name: normal?.name || 'Calibri', size: normal?.size || 11 };
    this.formatter = new NumberFormatter(this.date1904);
    this.externalBooks = extras?.externalBooks ?? [];
    this.arrayFormulas = extras?.arrayFormulas;
    this.threadedNotes = extras?.threadedNotes;
    this.converter = new StyleConverter(this.palette, this.defaultFont, extras);
    this.styleTable = new StyleTable((style) => this.converter.cellStyle(style));
    this.refresh();
  }

  getMeta(): WorkbookMeta {
    return {
      sheets: this.meta.sheets,
      styles: this.styleTable.styles.slice(),
      activeSheet: this.meta.activeSheet,
      defaultFont: { ...this.meta.defaultFont },
      date1904: this.meta.date1904,
    };
  }

  /** Recomputes the sheet structure from the workbook (after edits / sheet changes). */
  refresh(): WorkbookMeta {
    this.styleTable.forgetObjects();
    this.sheets = this.wb.worksheets.map((ws, index) => {
      const arrays = (typeof ws.id === 'number' ? this.arrayFormulas?.get(ws.id) : undefined) ?? [];
      return {
        ws,
        index,
        lists: new Map(),
        dynamicArrays: new Set(arrays.filter((a) => a.dynamic).map((a) => a.r * MAX_COLS + a.c)),
        arrayRanges: arrays
          .filter((a) => !a.dynamic && (a.range.r0 !== a.range.r1 || a.range.c0 !== a.range.c1))
          .sort((a, b) => a.range.r0 - b.range.r0),
        threadedNotes: typeof ws.id === 'number' ? this.threadedNotes?.get(ws.id) : undefined,
      };
    });
    const sheets = this.sheets.map((ctx) => this.safeSheetMeta(ctx));
    this.meta = { sheets, activeSheet: this.findActiveSheet(sheets), defaultFont: { ...this.defaultFont }, date1904: this.date1904 };
    return this.getMeta();
  }

  getRows(sheet: number, start: number, end: number, c0?: number, c1?: number): { rows: RowData[]; styles: [number, CellStyle][] } {
    const ctx = this.sheets[sheet];
    if (!ctx) {
      return { rows: [], styles: [] };
    }
    const xrows = ctx.ws._rows;
    const lo = Math.max(0, Math.floor(start) || 0);
    const hi = Math.min(Math.floor(end) || 0, xrows.length, MAX_ROWS);
    const [colLo, colHi] = columnWindow(c0, c1);
    // Multi-cell legacy array formulas reaching into this window (their other cells show the formula too).
    const arrays = ctx.arrayRanges.length ? ctx.arrayRanges.filter((a) => a.range.r0 < hi && a.range.r1 >= lo) : [];
    const rows: RowData[] = [];
    const used = new Set<number>();
    for (let r = lo; r < hi; r++) {
      const row = xrows[r];
      if (!row) {
        continue;
      }
      const cells: CellData[] = [];
      const xcells = row._cells;
      const last = Math.min(xcells.length - 1, colHi);
      for (let c = colLo; c <= last; c++) {
        const cell = xcells[c];
        if (!cell) {
          continue;
        }
        const data = this.safeCellData(cell, r, c, ctx);
        if (data) {
          if (arrays.length && data.f === undefined) this.markArrayMember(data, r, c, arrays, ctx);
          cells.push(data);
          if (data.s) used.add(data.s);
        }
      }
      if (cells.length) {
        rows.push({ r, cells });
      }
    }
    return { rows, styles: this.styleTable.entries(used) };
  }

  /** A cell inside a multi-cell legacy array formula shows the range's formula, as Excel's formula bar does. */
  private markArrayMember(data: CellData, r: number, c: number, arrays: readonly ArrayFormulaInfo[], ctx: SheetContext): void {
    for (const a of arrays) {
      const g = a.range;
      if (r < g.r0 || r > g.r1 || c < g.c0 || c > g.c1 || (r === a.r && c === a.c)) continue;
      const master = ctx.ws._rows[a.r]?._cells[a.c]?._value;
      if (master?.model?.shareType !== 'array') return; // the master is gone (an edit since loading)
      const f = this.formulaText(master, a.r, a.c, ctx);
      if (f) {
        data.f = f;
        data.array = true;
      }
      return;
    }
  }

  getStats(sheet: number, ranges: Range[]): SelectionStats {
    const acc = new StatsAccumulator();
    const ctx = this.sheets[sheet];
    if (!ctx) {
      return acc.result();
    }
    const xrows = ctx.ws._rows;
    forEachRangeRow(ranges, xrows.length, MAX_COLS, (r, c0, c1) => {
      const row = xrows[r];
      if (!row) {
        return true;
      }
      const cells = row._cells;
      const hi = Math.min(c1, cells.length - 1);
      for (let c = c0; c <= hi; c++) {
        if (!acc.tick()) {
          return false;
        }
        const cell = cells[c];
        if (cell) {
          acc.add(this.statValue(cell));
        }
      }
      return true;
    });
    if (acc.hasNumbers) {
      acc.numFmt = this.firstNumberFormat(ctx, ranges);
    }
    return acc.result((fmt, value) => this.formatter.number(fmt, value));
  }

  /**
   * Number format of the first numeric cell in the selection's own order (range by range, each one row by row), in
   * which Excel 16 shows Sum / Average / Min / Max: 'D1,D4' with a date in D1 sums as a date, 'D4,D1' as D4's General.
   * Undefined when the search passes the stats' work cap.
   */
  private firstNumberFormat(ctx: SheetContext, ranges: readonly Range[]): string | undefined {
    const xrows = ctx.ws._rows;
    let work = 0;
    for (const g of ranges) {
      if (!g || ![g.r0, g.r1, g.c0, g.c1].every((n) => typeof n === 'number' && Number.isFinite(n))) {
        continue;
      }
      const r1 = Math.min(xrows.length - 1, Math.floor(Math.max(g.r0, g.r1)));
      const c0 = Math.max(0, Math.floor(Math.min(g.c0, g.c1)));
      const c1 = Math.min(MAX_COLS - 1, Math.floor(Math.max(g.c0, g.c1)));
      for (let r = Math.max(0, Math.floor(Math.min(g.r0, g.r1))); r <= r1; r++) {
        if (++work > STATS_CELL_CAP) {
          return undefined;
        }
        const cells = xrows[r]?._cells;
        if (!cells) {
          continue;
        }
        const hi = Math.min(c1, cells.length - 1);
        for (let c = c0; c <= hi; c++) {
          const cell = cells[c];
          if (++work > STATS_CELL_CAP) {
            return undefined;
          }
          if (cell && typeof this.statValue(cell) === 'number') {
            return this.statNumFmt(cell);
          }
        }
      }
    }
    return undefined;
  }

  // ----- sheet meta -----

  /** Sheet meta; a sheet with unexpected data degrades to a plain grid instead of failing the workbook. */
  private safeSheetMeta(ctx: SheetContext): SheetMeta {
    try {
      return this.buildSheetMeta(ctx);
    } catch {
      const ws = ctx.ws;
      let colCount = 1;
      ws._rows?.forEach((row) => {
        if (row?._cells && row._cells.length > colCount) colCount = row._cells.length;
      });
      return {
        index: ctx.index,
        name: String(ws.name ?? `Sheet${ctx.index + 1}`),
        state: ws.state === 'hidden' || ws.state === 'veryHidden' ? ws.state : 'visible',
        rowCount: Math.max(1, ws._rows?.length ?? 0),
        colCount: Math.min(MAX_COLS, colCount),
        defaultColWidth: DEFAULT_COL_WIDTH_PX,
        defaultRowHeight: ptToPx(DEFAULT_ROW_HEIGHT_PT),
        cols: {},
        rows: {},
        merges: [],
        frozen: { rows: 0, cols: 0 },
        showGridLines: true,
        zoom: 100,
        conditionalFormats: [],
        images: [],
      };
    }
  }

  private buildSheetMeta(ctx: SheetContext): SheetMeta {
    const ws = ctx.ws;
    const props = ws.properties ?? {};
    const defaultRowHeight = typeof props.defaultRowHeight === 'number' && props.defaultRowHeight > 0 ? ptToPx(props.defaultRowHeight) : ptToPx(DEFAULT_ROW_HEIGHT_PT);
    const defaultColWidth = typeof props.defaultColWidth === 'number' && props.defaultColWidth > 0 ? widthToPx(props.defaultColWidth) : DEFAULT_COL_WIDTH_PX;
    const view = ws.views?.[0];
    const merges = this.sheetMerges(ws);
    const images = this.sheetImages(ws);
    const autoFilter = decodeAutoFilter(ws.autoFilter);
    const { rowCount, colCount } = this.usedRange(ws, merges, images, autoFilter, defaultColWidth, defaultRowHeight);
    const meta: SheetMeta = {
      index: ctx.index,
      name: ws.name,
      state: ws.state === 'hidden' || ws.state === 'veryHidden' ? ws.state : 'visible',
      rowCount,
      colCount,
      defaultColWidth,
      defaultRowHeight,
      cols: this.colInfos(ws, colCount, defaultColWidth),
      rows: this.rowInfos(ws, defaultRowHeight),
      merges,
      frozen:
        view?.state === 'frozen'
          ? { rows: clampInt(view.ySplit, 0, MAX_ROWS), cols: clampInt(view.xSplit, 0, MAX_COLS) }
          : { rows: 0, cols: 0 },
      showGridLines: view?.showGridLines !== false,
      zoom: typeof view?.zoomScale === 'number' && view.zoomScale > 0 ? Math.max(10, Math.min(400, view.zoomScale)) : 100,
      conditionalFormats: this.sheetConditionalFormats(ws),
      images,
    };
    const tabColor = resolveColor(props.tabColor, this.palette);
    if (tabColor) meta.tabColor = tabColor;
    if (view?.rightToLeft) meta.rightToLeft = true;
    if (autoFilter) meta.autoFilter = autoFilter;
    let validations: ListValidation[] = [];
    try {
      validations = this.sheetValidations(ctx);
    } catch {
      // unexpected validation data must not cost the sheet its structure
    }
    if (validations.length) meta.validations = validations;
    return meta;
  }

  private findActiveSheet(sheets: SheetMeta[]): number {
    const tab = this.wb.views?.[0]?.activeTab;
    let index = -1;
    if (typeof tab === 'number') {
      // workbook.xml activeTab counts every <sheet> (chartsheets too) - ExcelJS' orderNo does as well.
      index = this.sheets.findIndex((ctx) => ctx.ws.orderNo === tab);
      if (index < 0 && tab >= 0 && tab < sheets.length) index = tab;
    }
    // Excel opens a workbook whose active tab is a hidden sheet ON that sheet (it unhides it); a very hidden one never
    // (then the first visible sheet).
    if (index < 0 || (sheets[index]?.state !== 'visible' && sheets[index]?.state !== 'hidden')) {
      index = sheets.findIndex((s) => s.state === 'visible');
    }
    return Math.max(0, index);
  }

  /**
   * Used range: last row/column with a value, a non-default style, a comment or
   * row formatting, extended by merges, images and the auto-filter. (Worksheet's
   * actualRowCount counts rows rather than giving the last index.)
   */
  private usedRange(ws: XWorksheet, merges: Range[], images: ImageAnchor[], autoFilter: Range | undefined, colPx: number, rowPx: number): { rowCount: number; colCount: number } {
    let maxRow = 0;
    let maxCol = 0;
    ws._rows.forEach((row, r) => {
      if (!row) {
        return;
      }
      const last = lastUsedCell(row);
      if (last >= 0) {
        maxRow = Math.max(maxRow, r + 1);
        maxCol = Math.max(maxCol, last + 1);
      } else if (row.height !== undefined || row.hidden || !isStyleEmpty(row.style)) {
        maxRow = Math.max(maxRow, r + 1);
      }
    });
    for (const m of merges) {
      maxRow = Math.max(maxRow, m.r1 + 1);
      maxCol = Math.max(maxCol, m.c1 + 1);
    }
    for (const img of images) {
      const end = img.to ?? {
        r: img.from.r + Math.max(0, Math.ceil((img.from.dy + (img.ext?.h ?? 0)) / rowPx) - 1),
        c: img.from.c + Math.max(0, Math.ceil((img.from.dx + (img.ext?.w ?? 0)) / colPx) - 1),
      };
      maxRow = Math.max(maxRow, Math.min(MAX_ROWS, end.r + 1));
      maxCol = Math.max(maxCol, Math.min(MAX_COLS, end.c + 1));
    }
    if (autoFilter) {
      maxRow = Math.max(maxRow, autoFilter.r1 + 1);
      maxCol = Math.max(maxCol, autoFilter.c1 + 1);
    }
    return { rowCount: Math.max(1, maxRow), colCount: Math.max(1, maxCol) };
  }

  private colInfos(ws: XWorksheet, colCount: number, defaultPx: number): Record<number, ColInfo> {
    const out: Record<number, ColInfo> = {};
    const columns = ws._columns ?? [];
    const n = Math.min(columns.length, MAX_COLS);
    for (let c = 0; c < n; c++) {
      const column = columns[c];
      if (!column) {
        continue;
      }
      const info: ColInfo = {};
      const width = column.width;
      if (column.isCustomWidth && typeof width === 'number' && Number.isFinite(width)) {
        const px = Math.max(0, widthToPx(width));
        if (px !== defaultPx) info.w = px;
        if (px === 0) info.hidden = true;
      }
      if (column.hidden) info.hidden = true;
      if (c < colCount + COL_STYLE_MARGIN && !isStyleEmpty(column.style)) {
        const s = this.styleTable.indexOf(column.style);
        if (s) info.s = s;
      }
      if (info.w !== undefined || info.hidden || info.s) {
        out[c] = info;
      }
    }
    return out;
  }

  private rowInfos(ws: XWorksheet, defaultPx: number): Record<number, RowInfo> {
    const out: Record<number, RowInfo> = {};
    ws._rows.forEach((row, r) => {
      if (!row) {
        return;
      }
      const info: RowInfo = {};
      if (typeof row.height === 'number' && Number.isFinite(row.height)) {
        const px = Math.max(0, ptToPx(row.height));
        if (px !== defaultPx) info.h = px;
        if (px === 0) info.hidden = true;
      }
      if (row.hidden) info.hidden = true;
      if (!isStyleEmpty(row.style)) {
        const s = this.styleTable.indexOf(row.style);
        if (s) info.s = s;
      }
      if (info.h !== undefined || info.hidden || info.s) {
        out[r] = info;
      }
    });
    return out;
  }

  private sheetMerges(ws: XWorksheet): Range[] {
    const out: Range[] = [];
    for (const merge of Object.values(ws._merges ?? {})) {
      const m = merge?.model;
      if (!m || !m.top || !m.left || !m.bottom || !m.right) {
        continue;
      }
      out.push({ r0: m.top - 1, c0: m.left - 1, r1: m.bottom - 1, c1: m.right - 1 });
    }
    return out.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
  }

  private sheetImages(ws: XWorksheet): ImageAnchor[] {
    let images: XImage[] = [];
    try {
      images = ws.getImages?.() ?? [];
    } catch {
      return [];
    }
    const out: ImageAnchor[] = [];
    for (const image of images) {
      const range = image?.range;
      if (!range?.tl) {
        continue;
      }
      const src = this.imageUri(image.imageId);
      if (!src) {
        continue;
      }
      const anchor: ImageAnchor = { from: anchorPoint(range.tl), src };
      if (range.br) anchor.to = anchorPoint(range.br);
      const w = range.ext?.width;
      const h = range.ext?.height;
      if (typeof w === 'number' && typeof h === 'number' && (w > 0 || h > 0)) {
        anchor.ext = { w: Math.round(w), h: Math.round(h) };
      }
      out.push(anchor);
    }
    return out;
  }

  /** data: URI of a workbook medium (cached; formats browsers cannot show are skipped). */
  private imageUri(imageId: number | string | undefined): string | undefined {
    const key = String(imageId);
    if (this.imageUris.has(key)) {
      return this.imageUris.get(key);
    }
    let uri: string | undefined;
    const media = this.wb.media?.[Number(imageId)];
    const mime = media ? IMAGE_MIME[(media.extension ?? '').toLowerCase()] : undefined;
    if (media && mime) {
      if (media.buffer && media.buffer.length) {
        uri = `data:${mime};base64,${Buffer.from(media.buffer.buffer, media.buffer.byteOffset, media.buffer.byteLength).toString('base64')}`;
      } else if (typeof media.base64 === 'string' && media.base64) {
        uri = media.base64.startsWith('data:') ? media.base64 : `data:${mime};base64,${media.base64}`;
      }
    }
    this.imageUris.set(key, uri);
    return uri;
  }

  private sheetConditionalFormats(ws: XWorksheet): CfRule[] {
    const out: CfRule[] = [];
    let order = 0;
    for (const cf of ws.conditionalFormattings ?? []) {
      const ranges = parseSqref(cf?.ref);
      if (!ranges.length) {
        continue;
      }
      for (const rule of cf.rules ?? []) {
        if (rule) {
          out.push(this.cfRule(rule, ranges, ++order));
        }
      }
    }
    return out.sort((a, b) => a.priority - b.priority);
  }

  /**
   * ExcelJS rule -> CfRule. ExcelJS folds containsBlanks / notContainsBlanks /
   * containsErrors / notContainsErrors into type 'containsText' + operator; they
   * are unfolded again. For timePeriod rules `operator` holds the period
   * ('today', 'last7Days', ...).
   */
  private cfRule(rule: XCfRule, ranges: Range[], order: number): CfRule {
    let type: string = rule.type ?? 'expression';
    let operator = rule.operator;
    if (type === 'containsText' && operator && operator !== 'containsText') {
      type = operator;
      operator = undefined;
    }
    if (type === 'timePeriod') {
      operator = rule.timePeriod;
    }
    const out: CfRule = { ranges, priority: typeof rule.priority === 'number' && Number.isFinite(rule.priority) ? rule.priority : order, type };
    if (operator) out.operator = operator;
    const formulae = (rule.formulae ?? []).filter((f) => f !== undefined && f !== null).map((f) => String(f));
    if (formulae.length) out.formulae = formulae;
    if (type === 'containsText' || type === 'notContainsText' || type === 'beginsWith' || type === 'endsWith') {
      const text = typeof rule.text === 'string' ? rule.text : firstStringLiteral(formulae[0]);
      if (text !== undefined) out.text = text;
    }
    if (type === 'top10') {
      out.rank = typeof rule.rank === 'number' && Number.isFinite(rule.rank) ? rule.rank : 10;
      out.percent = !!rule.percent;
      out.bottom = !!rule.bottom;
    }
    if (type === 'aboveAverage') out.aboveAverage = rule.aboveAverage !== false;
    if (rule.stopIfTrue) out.stopIfTrue = true;
    if (rule.style) {
      const style = this.converter.dxfStyle(rule.style);
      if (style) out.style = style;
    }
    const cfvo = (rule.cfvo ?? []).map(normalizeCfvo);
    const colors = Array.isArray(rule.color) ? rule.color : rule.color ? [rule.color] : [];
    if (type === 'colorScale') {
      out.colorScale = { cfvo, colors: colors.map((c) => this.converter.color(c) ?? '#FFFFFF') };
    } else if (type === 'dataBar') {
      out.dataBar = {
        cfvo,
        color: this.converter.color(colors[0]) ?? '#638EC6',
        gradient: rule.gradient !== false,
        showValue: rule.showValue !== false,
      };
    } else if (type === 'iconSet') {
      out.iconSet = { name: rule.iconSet || '3TrafficLights', cfvo, reverse: !!rule.reverse, showValue: rule.showValue !== false };
    }
    return out;
  }

  // ----- cells -----

  private safeCellData(cell: XCell, r: number, c: number, ctx: SheetContext): CellData | undefined {
    try {
      return this.cellData(cell, r, c, ctx);
    } catch {
      return undefined; // a malformed cell must not break the whole window
    }
  }

  /** One ExcelJS cell -> CellData; undefined when the cell carries nothing to show. */
  private cellData(cell: XCell, r: number, c: number, ctx: SheetContext): CellData | undefined {
    const s = this.styleTable.indexOf(cell.style);
    const fmt = this.styleTable.styles[s].numFmt ?? 'General';
    const font = cell.style?.font;
    const xv = cell._value;
    const m = xv.model;
    let out: CellData;
    switch (xv.type) {
      case VT.Null:
      case VT.Merge:
        out = { c, t: 'z' };
        break;
      case VT.Formula: {
        const result = m.result;
        out = result === undefined || result === null ? { c, t: 's', v: '', w: '' } : this.valueCell(c, result, fmt, font);
        const f = this.formulaText(xv, r, c, ctx);
        if (f) {
          out.f = f;
          if (m.shareType === 'array' && !ctx.dynamicArrays.has(r * MAX_COLS + c)) out.array = true;
          const link = this.normalizeLink(hyperlinkFormulaTarget(f));
          if (link) out.link = link;
        }
        break;
      }
      case VT.Hyperlink: {
        out = this.valueCell(c, m.text ?? '', fmt, font);
        const link = this.normalizeLink(m.hyperlink);
        if (link) out.link = link;
        break;
      }
      default:
        out = this.valueCell(c, m.value, fmt, font);
        break;
    }
    if (s) out.s = s;
    const note = ctx.threadedNotes?.get(r * MAX_COLS + c) ?? noteText(cell);
    if (note) out.note = note;
    if (out.t === 'z' && !out.s && !out.note) {
      return undefined;
    }
    return out;
  }

  /** Raw ExcelJS value (number, string, boolean, Date, {error}, {richText}, ...) -> CellData. */
  private valueCell(c: number, raw: unknown, fmt: string, font: XFont | undefined): CellData {
    switch (typeof raw) {
      case 'number':
        return this.numberCell(c, raw, fmt);
      case 'string': {
        if (fmt === 'General' || fmt === '@' || this.formatter.textFillSection(fmt) === null) {
          return { c, t: 's', v: raw, w: this.formatter.text(fmt, raw) };
        }
        const layout: LayoutOut = {};
        const cell: CellData = { c, t: 's', v: raw, w: this.formatter.text(fmt, raw, layout) };
        if (layout.fill) cell.fill = layout.fill;
        return cell;
      }
      case 'boolean':
        return { c, t: 'b', v: raw, w: raw ? 'TRUE' : 'FALSE' };
      case 'undefined':
        return { c, t: 'z' };
      default:
        break;
    }
    if (raw === null) {
      return { c, t: 'z' };
    }
    if (raw instanceof Date) {
      return this.dateCell(c, raw, fmt);
    }
    if (typeof raw === 'object') {
      const o = raw as { error?: unknown; richText?: unknown; text?: unknown; result?: unknown };
      if (o.error !== undefined) {
        const error = String(o.error);
        return { c, t: 'e', v: error, w: error };
      }
      if (Array.isArray(o.richText)) {
        return this.richCell(c, o.richText as XRichRun[], font);
      }
      if ('text' in o) {
        return this.valueCell(c, o.text, fmt, font); // nested hyperlink value
      }
      if ('result' in o) {
        return this.valueCell(c, o.result, fmt, font); // formula value object
      }
    }
    const text = String(raw);
    return { c, t: 's', v: text, w: text };
  }

  private numberCell(c: number, value: number, fmt: string): CellData {
    if (!Number.isFinite(value)) {
      return { c, t: 'e', v: '#NUM!', w: '#NUM!' };
    }
    if (!this.formatter.numberHasFill(fmt)) {
      return { c, t: this.formatter.isDate(fmt) ? 'd' : 'n', v: value, w: this.formatter.number(fmt, value) };
    }
    // Accounting-style format: the grid lays the text out around the repeated fill character.
    const layout: LayoutOut = {};
    const cell: CellData = { c, t: 'n', v: value, w: this.formatter.number(fmt, value, layout) };
    if (layout.fill) cell.fill = layout.fill;
    return cell;
  }

  /** ExcelJS turns date-formatted numbers into Dates (UTC wall clock); back to a serial. */
  private dateCell(c: number, date: Date, fmt: string): CellData {
    const serial = this.dateSerial(date);
    if (serial === undefined) {
      return { c, t: 'e', v: '#VALUE!', w: '#VALUE!' };
    }
    if (fmt === 'General') {
      const dateFmt = Number.isInteger(serial) ? BUILTIN_NUMFMTS[14] : BUILTIN_NUMFMTS[22];
      return { c, t: 'd', v: serial, w: this.formatter.number(dateFmt, serial) };
    }
    return this.numberCell(c, serial, fmt);
  }

  /**
   * Date -> Excel serial. ExcelJS rounded the serial to whole milliseconds, so the
   * epoch offset is added in integer milliseconds before the single division: that
   * gives back the closest double (0.6041666666666666 for 14:30, not ...78793).
   */
  private dateSerial(date: Date): number | undefined {
    const ms = date.getTime();
    if (!Number.isFinite(ms)) {
      return undefined;
    }
    const epochDays = this.date1904 ? 25569 - DATE1904_OFFSET : 25569;
    return (ms + epochDays * MS_PER_DAY) / MS_PER_DAY;
  }

  private richCell(c: number, runs: XRichRun[], cellFont: XFont | undefined): CellData {
    let text = '';
    let styled = false;
    const rt: RichRun[] = [];
    for (const run of runs) {
      const part = typeof run?.text === 'string' ? run.text : run?.text === undefined || run?.text === null ? '' : String(run.text);
      text += part;
      const font = run?.font ? this.converter.runFont(run.font, cellFont) : undefined;
      if (font) {
        styled = true;
        rt.push({ text: part, font });
      } else {
        rt.push({ text: part });
      }
    }
    const out: CellData = { c, t: 's', v: text, w: text };
    if (styled) out.rt = rt;
    return out;
  }

  /** Formula text; shared-formula dependents are translated from their master. */
  private formulaText(xv: XValue, r: number, c: number, ctx: SheetContext): string | undefined {
    const m = xv.model;
    if (typeof m.formula === 'string' && m.formula) {
      return showExternalRefs(displayFormula(m.formula), this.externalBooks);
    }
    if (typeof m.sharedFormula !== 'string') {
      return undefined;
    }
    const cached = this.sharedFormulas.get(xv);
    if (cached !== undefined) {
      return cached;
    }
    const master = decodeCell(m.sharedFormula);
    const masterFormula = master ? ctx.ws._rows[master.r]?._cells[master.c]?._value?.model?.formula : undefined;
    if (!master || typeof masterFormula !== 'string') {
      return undefined;
    }
    const text = showExternalRefs(displayFormula(translateFormula(masterFormula, r - master.r, c - master.c)), this.externalBooks);
    this.sharedFormulas.set(xv, text);
    return text;
  }

  /** External links as-is; internal ones as '#Sheet!A1' (defined names resolved, '$' dropped). */
  private normalizeLink(target: unknown): string | undefined {
    if (typeof target !== 'string' || !target) {
      return undefined;
    }
    if (!target.startsWith('#')) {
      return target;
    }
    const cached = this.links.get(target);
    if (cached !== undefined) {
      return cached;
    }
    let location = target.slice(1).trim();
    if (!location.includes('!')) {
      const named = this.definedNameRange(location);
      if (named) location = named;
    }
    let out = target;
    if (location.includes('!')) {
      const { sheet, ref } = splitSheetRef(location);
      out = `#${sheet === undefined ? '' : `${quoteSheetName(sheet)}!`}${ref.replace(/\$/g, '')}`;
    }
    this.links.set(target, out);
    return out;
  }

  private definedNameRange(name: string): string | undefined {
    try {
      const ranges = this.wb.definedNames?.getRanges(name)?.ranges;
      return ranges && ranges.length ? ranges[0] : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * Stats contribution: undefined = empty, null = non-numeric value, 'error' = a cell the grid shows as an error
   * (#N/A, #DIV/0! ..., also #NUM! for a non-finite number and #VALUE! for an unreadable date), number = numeric.
   */
  private statValue(cell: XCell): StatValue {
    const xv = cell._value;
    const m = xv.model;
    switch (xv.type) {
      case VT.Null:
      case VT.Merge:
        return undefined;
      case VT.Number:
        return typeof m.value === 'number' ? (Number.isFinite(m.value) ? m.value : 'error') : null;
      case VT.Date:
        return m.value instanceof Date ? this.dateSerial(m.value) ?? 'error' : null;
      case VT.Error:
        return 'error';
      case VT.Formula: {
        const result = m.result;
        if (typeof result === 'number') return Number.isFinite(result) ? result : 'error';
        if (result instanceof Date) return this.dateSerial(result) ?? 'error';
        if (isErrorValue(result)) return 'error';
        return null;
      }
      case VT.Hyperlink:
        return typeof m.text === 'number' && Number.isFinite(m.text) ? m.text : null;
      default:
        return null;
    }
  }

  /** Number format a numeric cell is shown in (as cellData: a date without a format shows as a date). */
  private statNumFmt(cell: XCell): string {
    const fmt = this.styleTable.styles[this.styleTable.indexOf(cell.style)].numFmt ?? 'General';
    if (fmt !== 'General') return fmt;
    const xv = cell._value;
    const date = xv.type === VT.Date ? xv.model.value : xv.type === VT.Formula ? xv.model.result : undefined;
    if (!(date instanceof Date)) return fmt;
    return Number.isInteger(this.dateSerial(date)) ? BUILTIN_NUMFMTS[14] : BUILTIN_NUMFMTS[22];
  }

  // ----- data validation -----

  /**
   * List validations of a sheet, one entry per distinct list. Loaded workbooks use
   * the rules recorded with their sqref while parsing (a whole-column rule is one
   * range, not a million cells); workbooks built in memory fall back to ExcelJS'
   * per-cell keys, merged back into ranges. Lists that cannot be resolved
   * (INDIRECT, external workbooks ...) are left out.
   */
  private sheetValidations(ctx: SheetContext): ListValidation[] {
    const model = ctx.ws.dataValidations?.model;
    if (!model) {
      return [];
    }
    const byFormula = new Map<string, Range[]>();
    const formulaOf = (rule: XDataValidation | undefined): string | undefined => {
      const f = rule?.type === 'list' ? rule.formulae?.[0] : undefined;
      return f === undefined || f === null ? undefined : String(f);
    };
    const add = (formula: string, ranges: Range[]): void => {
      const list = byFormula.get(formula);
      if (list) list.push(...ranges);
      else byFormula.set(formula, ranges);
    };
    const loaded = loadedValidations.get(model);
    if (loaded) {
      for (const { sqref, rule } of loaded) {
        const formula = formulaOf(rule);
        const ranges = formula === undefined ? [] : parseSqref(sqref.replace(/^range:/, ''));
        if (formula !== undefined && ranges.length) add(formula, ranges);
      }
    } else {
      const cells = new Map<string, { r: number; c: number }[]>();
      let budget = VALIDATION_CELLS_CAP;
      for (const key in model) {
        if (--budget < 0) break;
        const formula = formulaOf(model[key]);
        const range = formula === undefined ? undefined : decodeRange(key.replace(/^range:/, ''));
        if (!range || formula === undefined) continue;
        if (range.r0 === range.r1 && range.c0 === range.c1) {
          const list = cells.get(formula);
          if (list) list.push({ r: range.r0, c: range.c0 });
          else cells.set(formula, [{ r: range.r0, c: range.c0 }]);
        } else {
          add(formula, [range]);
        }
      }
      cells.forEach((list, formula) => add(formula, mergeCellsIntoRanges(list)));
    }
    const out: ListValidation[] = [];
    const byValues = new Map<string, ListValidation>();
    byFormula.forEach((ranges, formula) => {
      const values = this.listValues(ctx, formula);
      if (!values) return;
      const key = JSON.stringify(values);
      const same = byValues.get(key);
      if (same) {
        same.ranges.push(...ranges);
      } else {
        const entry = { ranges, values };
        byValues.set(key, entry);
        out.push(entry);
      }
    });
    return out;
  }

  /**
   * Values of a list validation: a literal ("a,b,c"), a range (Sheet2!$A$1:$A$5,
   * same-sheet $A$1:$A$5) or a defined name. Other formulas (INDIRECT...) -> undefined.
   */
  private listValues(ctx: SheetContext, formula: string): string[] | undefined {
    if (ctx.lists.has(formula)) {
      return ctx.lists.get(formula);
    }
    let result: string[] | undefined;
    try {
      let f = formula.trim();
      if (f.startsWith('=')) f = f.slice(1).trim();
      if (f.startsWith('"')) {
        const literal = firstStringLiteral(f);
        result = literal === undefined ? undefined : literal.split(',').map((item) => item.trim());
      } else {
        result = this.rangeValues(ctx, f) ?? this.rangeValues(ctx, this.definedNameRange(f) ?? '');
      }
    } catch {
      result = undefined;
    }
    ctx.lists.set(formula, result);
    return result;
  }

  /** Display texts of a (sheet-qualified) range, row-major, empty cells skipped. */
  private rangeValues(ctx: SheetContext, text: string): string[] | undefined {
    if (!text) {
      return undefined;
    }
    const { sheet, ref } = splitSheetRef(text);
    if (sheet !== undefined && /^\[/.test(sheet)) {
      return undefined; // external workbook
    }
    const range = decodeRange(ref);
    if (!range) {
      return undefined;
    }
    const target = sheet === undefined ? ctx : this.sheets.find((s) => s.ws.name.toLowerCase() === sheet.toLowerCase());
    if (!target) {
      return undefined;
    }
    const values: string[] = [];
    const xrows = target.ws._rows;
    const lastRow = Math.min(range.r1, xrows.length - 1);
    for (let r = range.r0; r <= lastRow && values.length < LIST_VALUES_CAP; r++) {
      const cells = xrows[r]?._cells;
      if (!cells) continue;
      const lastCol = Math.min(range.c1, cells.length - 1);
      for (let c = range.c0; c <= lastCol && values.length < LIST_VALUES_CAP; c++) {
        const cell = cells[c];
        const data = cell ? this.safeCellData(cell, r, c, target) : undefined;
        const display = data && data.t !== 'z' ? data.w ?? String(data.v ?? '') : '';
        if (display !== '') values.push(display);
      }
    }
    return values;
  }
}

/** ExcelJS workbook -> model (meta computed eagerly, rows converted lazily per window). */
export function workbookToModel(workbook: ExcelJS.Workbook): WorkbookModel {
  return new WorkbookModel(workbook);
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  ico: 'image/x-icon',
};

function pickTheme(themes: Record<string, string> | undefined): string | undefined {
  if (!themes) {
    return undefined;
  }
  return themes.theme1 ?? Object.values(themes).find((xml) => typeof xml === 'string');
}

/** Anchor in native coordinates (0-based cell + EMU offset) -> px offsets. */
function anchorPoint(anchor: XAnchor): { r: number; c: number; dx: number; dy: number } {
  return {
    r: Math.max(0, Math.floor(anchor.nativeRow ?? 0)),
    c: Math.max(0, Math.floor(anchor.nativeCol ?? 0)),
    dx: Math.round((anchor.nativeColOff ?? 0) / EMU_PER_PX),
    dy: Math.round((anchor.nativeRowOff ?? 0) / EMU_PER_PX),
  };
}

/** worksheet.autoFilter: 'A1:D10' (loaded) or {from, to} (API). */
function decodeAutoFilter(filter: unknown): Range | undefined {
  if (typeof filter === 'string') {
    return decodeRange(filter);
  }
  if (filter && typeof filter === 'object') {
    const f = filter as { from?: unknown; to?: unknown };
    const point = (p: unknown): { r: number; c: number } | undefined => {
      if (typeof p === 'string') return decodeCell(p);
      if (p && typeof p === 'object') {
        const q = p as { row?: number; column?: number };
        if (typeof q.row === 'number' && typeof q.column === 'number') return { r: q.row - 1, c: q.column - 1 };
      }
      return undefined;
    };
    const a = point(f.from);
    const b = point(f.to);
    if (a && b) {
      return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
    }
  }
  return undefined;
}

/** Single cells -> rectangles: vertical runs per column, then equal runs of adjacent columns joined. */
function mergeCellsIntoRanges(cells: { r: number; c: number }[]): Range[] {
  const runs: Range[] = [];
  for (const { r, c } of cells.slice().sort((a, b) => a.c - b.c || a.r - b.r)) {
    const last = runs[runs.length - 1];
    if (last && last.c0 === c && r <= last.r1 + 1) last.r1 = Math.max(last.r1, r);
    else runs.push({ r0: r, c0: c, r1: r, c1: c });
  }
  const out: Range[] = [];
  const open = new Map<string, Range>(); // "r0:r1" -> the range that ends in the previous column
  for (const run of runs) {
    const key = `${run.r0}:${run.r1}`;
    const prev = open.get(key);
    if (prev && prev.c1 === run.c0 - 1) {
      prev.c1 = run.c0;
    } else {
      const range = { ...run };
      out.push(range);
      open.set(key, range);
    }
  }
  return out;
}

/** Index of the last cell carrying a value, a style or a comment; -1 if none. */
function lastUsedCell(row: XRow): number {
  const cells = row._cells;
  for (let c = cells.length - 1; c >= 0; c--) {
    const cell = cells[c];
    if (cell && (cell._value.type !== VT.Null || !isStyleEmpty(cell.style) || cell._comment)) {
      return c;
    }
  }
  return -1;
}

/** Plain text of a cell note (string or rich {texts}). */
function noteText(cell: XCell): string | undefined {
  const note = cell._comment?.note;
  if (typeof note === 'string') {
    return note || undefined;
  }
  if (note && Array.isArray(note.texts)) {
    const text = note.texts.map((t) => (typeof t?.text === 'string' ? t.text : '')).join('');
    return text || undefined;
  }
  return undefined;
}

/** Target of a `HYPERLINK("url", ...)` formula when the target is a string literal. */
function hyperlinkFormulaTarget(formula: string): string | undefined {
  const m = /^\s*HYPERLINK\(\s*"((?:[^"]|"")*)"\s*[,)]/i.exec(formula);
  return m ? m[1].replace(/""/g, '"') : undefined;
}

/** First "..." literal of a formula, unescaped. */
function firstStringLiteral(formula: string | undefined): string | undefined {
  if (!formula) {
    return undefined;
  }
  const m = /"((?:[^"]|"")*)"/.exec(formula);
  return m ? m[1].replace(/""/g, '"') : undefined;
}

const CFVO_TYPES: ReadonlySet<string> = new Set(['min', 'max', 'num', 'percent', 'percentile', 'formula', 'autoMin', 'autoMax']);

function normalizeCfvo(cfvo: XCfvo | undefined): CfValueObject {
  const type = cfvo?.type && CFVO_TYPES.has(cfvo.type) ? (cfvo.type as CfValueObject['type']) : 'num';
  const out: CfValueObject = { type };
  const value = cfvo?.value;
  if (typeof value === 'number' && Number.isFinite(value)) {
    out.value = value;
  } else if (typeof value === 'string' && value) {
    out.value = value;
  }
  return out;
}

function clampInt(value: unknown, min: number, max: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? Math.max(min, Math.min(max, Math.round(value))) : min;
}

// ===== SELECTION STATS =====

/** One cell's part in the stats: undefined = empty, null = not numeric, 'error' = an error value, number = numeric. */
type StatValue = number | null | undefined | 'error';

/** An ExcelJS error value ({ error: '#N/A' }), e.g. a formula's cached result. */
function isErrorValue(value: unknown): boolean {
  return typeof value === 'object' && value !== null && (value as { error?: unknown }).error !== undefined;
}

/** Aggregates Count / Sum / Average / Min / Max with a work cap (a capped result says so: `partial`). */
class StatsAccumulator {
  /** Number format of the selection's first numeric cell (set by the caller); Sum / Average / Min / Max are shown in it. */
  numFmt: string | undefined;
  private count = 0;
  private numCount = 0;
  private errors = 0;
  private sum = 0;
  private min = Infinity;
  private max = -Infinity;
  private work = 0;
  private stopped = false;

  /** Counts one visited cell slot; false once the cap is exceeded (the result is then partial). */
  tick(): boolean {
    if (++this.work <= STATS_CELL_CAP) return true;
    this.stopped = true;
    return false;
  }

  get hasNumbers(): boolean {
    return this.numCount > 0;
  }

  /** undefined = empty (ignored), null = non-empty but not numeric, 'error' = an error value, number = numeric. */
  add(value: StatValue): void {
    if (value === undefined) return;
    this.count++;
    if (value === null) return;
    if (value === 'error') {
      this.errors++;
      return;
    }
    this.numCount++;
    this.sum += value;
    if (value < this.min) this.min = value;
    if (value > this.max) this.max = value;
  }

  /** `format` (workbooks): formats a value in `numFmt` for SelectionStats.text. */
  result(format?: (fmt: string, value: number) => string): SelectionStats {
    const out: SelectionStats = { count: this.count, numCount: this.numCount, sum: this.sum };
    if (this.numCount) {
      out.avg = this.sum / this.numCount;
      out.min = this.min;
      out.max = this.max;
      const fmt = this.numFmt;
      if (format && fmt !== undefined && fmt !== 'General') {
        try {
          // Without the layout the grid uses: the padding of '_)' / '* ' formats is not wanted in the status bar.
          const text = (value: number): string => format(fmt, value).trim();
          out.text = { sum: text(this.sum), avg: text(out.avg), min: text(this.min), max: text(this.max) };
        } catch {
          // The numbers are still shown, in the viewer's own format.
        }
      }
    }
    if (this.errors) {
      out.errors = this.errors;
    }
    if (this.stopped) {
      out.partial = true;
      out.scanned = STATS_CELL_CAP;
    }
    return out;
  }
}

/**
 * Visits the union of `ranges` row by row as merged column spans, so overlapping
 * ranges are counted once. Ranges are clamped to [0, rowLimit) x [0, colLimit).
 * `visit` returns false to stop.
 */
function forEachRangeRow(ranges: readonly Range[] | undefined, rowLimit: number, colLimit: number, visit: (r: number, c0: number, c1: number) => boolean): void {
  const clean: Range[] = [];
  for (const g of ranges ?? []) {
    if (!g || ![g.r0, g.r1, g.c0, g.c1].every((n) => typeof n === 'number' && Number.isFinite(n))) {
      continue;
    }
    const r0 = Math.max(0, Math.floor(Math.min(g.r0, g.r1)));
    const r1 = Math.min(rowLimit - 1, Math.floor(Math.max(g.r0, g.r1)));
    const c0 = Math.max(0, Math.floor(Math.min(g.c0, g.c1)));
    const c1 = Math.min(colLimit - 1, Math.floor(Math.max(g.c0, g.c1)));
    if (r0 <= r1 && c0 <= c1) {
      clean.push({ r0, c0, r1, c1 });
    }
  }
  if (!clean.length) {
    return;
  }
  // Split the rows into bands where the set of covering ranges is constant.
  const cuts = [...new Set(clean.flatMap((g) => [g.r0, g.r1 + 1]))].sort((a, b) => a - b);
  for (let i = 0; i < cuts.length - 1; i++) {
    const top = cuts[i];
    const bottom = cuts[i + 1] - 1;
    const spans = clean
      .filter((g) => g.r0 <= top && g.r1 >= bottom)
      .map((g) => [g.c0, g.c1] as [number, number])
      .sort((a, b) => a[0] - b[0]);
    if (!spans.length) {
      continue;
    }
    const merged: [number, number][] = [];
    for (const span of spans) {
      const last = merged[merged.length - 1];
      if (last && span[0] <= last[1] + 1) {
        last[1] = Math.max(last[1], span[1]);
      } else {
        merged.push([span[0], span[1]]);
      }
    }
    for (let r = top; r <= bottom; r++) {
      for (const [c0, c1] of merged) {
        if (!visit(r, c0, c1)) {
          return;
        }
      }
    }
  }
}

// ===== CSV =====

/**
 * How the source quoted fields that did not need quoting: minimal quoting, or
 * one of the common "quote everything / every non-empty / every non-numeric
 * field" styles. Detected from the first records.
 */
type QuotePolicy = 'minimal' | 'all' | 'nonEmpty' | 'nonNumeric';

/**
 * What serializeCsv needs to reproduce the source byte for byte: the quoting
 * policy, plus the original text of every record that the policy does not
 * reproduce (e.g. one field quoted only because it has padding spaces). Raw
 * texts are keyed on the row array, so inserted / deleted rows do not shift
 * them, and are only used while they still decode to the row's current fields.
 */
interface CsvSource {
  policy: QuotePolicy;
  raw?: WeakMap<string[], string>;
}

/** Keyed on both the parsed model and its rows array (edits may copy one or the other). */
const csvSources = new WeakMap<object, CsvSource>();
const CSV_DELIMITERS = [',', ';', '\t', '|'];
/** Tie-break order when several delimiters split the sample equally well. */
const CSV_DELIMITER_PRIORITY = ['\t', ';', '|', ','];
const CSV_SCAN_LIMIT = 1 << 20;
const CSV_SAMPLE_RECORDS = 200;
const CSV_POLICY_SAMPLE_ROWS = 200;
const QUOTE_POLICIES: readonly QuotePolicy[] = ['minimal', 'all', 'nonNumeric', 'nonEmpty'];
const NUMERIC_TEXT_RE = /^[ \t]*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?[ \t]*$/;
const GROUPED_NUMBER_RE = /^[ \t]*[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?[ \t]*$/;
const DECIMAL_COMMA_RE = /^[ \t]*[-+]?\d+,\d+[ \t]*$/;
const PERCENT_TEXT_RE = /^[ \t]*[-+]?(?:\d+\.?\d*|\.\d+)%[ \t]*$/;
/**
 * Excel's delimiter directive: a first line 'sep=' + one character (any case, spaces around it and surrounding quotes
 * allowed, also as the whole file), checked against Excel 16. Group 1 = the delimiter.
 */
const CSV_SEP_LINE_RE = /^[ \t]*"?sep=([^\r\n])"?[ \t]*(?:\r\n|\n|\r|$)/i;

/**
 * Parses CSV/TSV text with papaparse (Excel's rules where papaparse reads
 * malformed quoting differently, see readCsvRecords). Detects BOM, delimiter
 * (Excel's 'sep=' first line, else , ; tab |), dominant line ending and
 * trailing newline so serializeCsv can reproduce the file. `delimiter` fixes
 * the delimiter (PSV: '|'); a 'sep=' line is then still not a row.
 *
 * `delimiter: ' '` with `collapseSpaces` reads space-separated text like
 * Excel's "treat consecutive delimiters as one" (checked against Excel 16
 * OpenText): a run of spaces is one separator, spaces after the last field
 * are not a field, tabs are field text, and fields may be quoted (CSV rules:
 * `""` is an empty field, `"New York"` keeps its space). Two deliberate
 * differences from Excel's text import: spaces before the first field are not
 * an empty first field (right-aligned columns stay aligned), and a quoted
 * field may span lines like in Excel's CSV reader (so serializeCsv output
 * always reads back).
 */
export function parseCsv(text: string, opts?: { delimiter?: string; collapseSpaces?: boolean }): CsvModel {
  let body = text;
  let hasBom = false;
  if (body.charCodeAt(0) === 0xfeff) {
    hasBom = true;
    body = body.slice(1);
  }
  // Excel's 'sep=;' first line names the delimiter and is not data (see CSV_SEP_LINE_RE).
  const sep = CSV_SEP_LINE_RE.exec(body);
  if (sep) {
    body = body.slice(sep[0].length);
  }
  const collapse = opts?.delimiter === ' ' && !!opts.collapseSpaces;
  const scan = scanCsv(body, collapse);
  const delimiter = opts?.delimiter || sep?.[1] || scan.delimiter;
  const newline = scan.newline;
  // Whether the last record is terminated by a line break (decided after parsing: an
  // unterminated quote can swallow the final line break into the last field).
  let trailingNewline = false;
  if (hasOtherLineBreaks(body, newline)) {
    // papaparse splits records on a single newline sequence: unify the stray ones first.
    body = unifyCsvNewlines(body, delimiter, newline);
  }
  let rows: string[][] = [];
  /** Offset just past each record (including its line break). */
  let ends: number[] = [];
  if (body.length) {
    ({ rows, ends } = readCsvRecords(body, delimiter, newline, collapse));
    const last = rows[rows.length - 1];
    if ((body.endsWith('\n') || body.endsWith('\r')) && last && last.length === 1 && last[0] === '') {
      rows.pop(); // the final line break reads as an empty record
      ends.pop();
      trailingNewline = true;
    }
  }
  const model: CsvModel = { delimiter, newline, hasBom, rows, trailingNewline };
  if (collapse) {
    model.collapseSpaces = true;
  }
  if (sep) {
    model.sepLine = sep[0];
  }
  const source = describeCsvSource(model, body, ends);
  if (source) {
    csvSources.set(model, source);
    csvSources.set(rows, source);
  }
  return model;
}

/** CsvModel -> text, keeping delimiter, BOM, newline style, quoting style and trailing newline. */
export function serializeCsv(model: CsvModel): string {
  const source = csvSources.get(model) ?? csvSources.get(model.rows);
  const policy = source?.policy ?? 'minimal';
  const delimiter = model.delimiter || ',';
  const collapse = delimiter === ' ' && !!model.collapseSpaces;
  const newline = model.newline || '\n';
  const lines = new Array<string>(model.rows.length);
  for (let i = 0; i < model.rows.length; i++) {
    const row = model.rows[i] ?? [];
    const raw = source?.raw?.get(row);
    lines[i] =
      raw !== undefined && rawRecordMatches(raw, row, delimiter, newline, collapse) ? raw : serializeCsvRow(row, delimiter, policy, collapse);
  }
  let out = (model.sepLine ?? '') + lines.join(newline);
  if (model.trailingNewline && model.rows.length) {
    out += newline;
  }
  return model.hasBom ? `\uFEFF${out}` : out;
}

/**
 * Delimiter of a .ssv file, from the first MB / 200 non-blank records: ';' when
 * semicolons split the sample consistently, else ' ' (space-separated: parse
 * with `collapseSpaces`). Consistently = most records hold the same number of
 * them (the rule parseCsv uses for .csv), or, in a ragged file, most records
 * hold some while runs of spaces do not split the records consistently.
 * Semicolons inside quoted fields do not count; a quote opens a field at the
 * start of a record or after ';' or a space. An empty sample is space-separated.
 */
export function detectSsvDelimiter(text: string): ';' | ' ' {
  const limit = Math.min(text.length, CSV_SCAN_LIMIT);
  /** Per sample record: ';' count, and separators when runs of spaces separate. */
  const semicolons: number[] = [];
  const spaceRuns: number[] = [];
  let semi = 0;
  let fields = 0;
  let inField = false;
  let inQuotes = false;
  let fieldStart = true;
  const endRecord = (): void => {
    if (fields && semicolons.length < CSV_SAMPLE_RECORDS) {
      semicolons.push(semi);
      spaceRuns.push(fields - 1);
    }
    semi = 0;
    fields = 0;
    inField = false;
    fieldStart = true;
  };
  for (let i = text.charCodeAt(0) === 0xfeff ? 1 : 0; i < limit && semicolons.length < CSV_SAMPLE_RECORDS; i++) {
    const ch = text.charCodeAt(i);
    if (inQuotes) {
      if (ch === 34) {
        if (text.charCodeAt(i + 1) === 34) i++;
        else inQuotes = false;
      }
      continue;
    }
    if (ch === 13 || ch === 10) {
      endRecord(); // the LF of a CRLF ends an empty record, which is not sampled
      continue;
    }
    if (ch === 32) {
      inField = false;
      fieldStart = true;
      continue;
    }
    if (!inField) {
      fields++;
      inField = true;
    }
    if (ch === 34 && fieldStart) {
      inQuotes = true;
    } else if (ch === 59) {
      semi++;
    }
    fieldStart = ch === 59;
  }
  endRecord();
  const records = semicolons.length;
  if (!records) {
    return ' ';
  }
  const required = requiredConsistency(records);
  if (splitConsistency(semicolons) >= required) {
    return ';';
  }
  const withSemicolons = semicolons.reduce((sum, n) => sum + (n > 0 ? 1 : 0), 0);
  return withSemicolons / records >= required && splitConsistency(spaceRuns) < required ? ';' : ' ';
}

function papaOptions(delimiter: string, newline: string): Papa.ParseConfig<string[]> {
  return {
    delimiter,
    newline: newline as '\r\n' | '\n' | '\r',
    quoteChar: '"',
    escapeChar: '"',
    header: false,
    dynamicTyping: false,
    skipEmptyLines: false,
  };
}

/** One record; space-separated (`collapse`) records use single spaces and a lone empty field is a blank line. */
function serializeCsvRow(row: readonly string[], delimiter: string, policy: QuotePolicy, collapse = false): string {
  if (collapse && row.length <= 1 && !row[0]) {
    return '';
  }
  let line = '';
  for (let j = 0; j < row.length; j++) {
    if (j) line += delimiter;
    line += quoteCsvField(row[j] ?? '', delimiter, policy, collapse);
  }
  return line;
}

/**
 * Quotes a field when it contains the delimiter, a quote, CR or LF (or when the policy asks). Space-separated
 * (`collapse`) text also quotes empty fields: unquoted, they would vanish into the run of spaces.
 */
function quoteCsvField(field: string, delimiter: string, policy: QuotePolicy, collapse = false): string {
  let quote =
    field.includes(delimiter) || field.includes('"') || field.includes('\n') || field.includes('\r') || (collapse && field === '');
  if (!quote && policy !== 'minimal') {
    quote = policy === 'all' || (policy === 'nonEmpty' ? field !== '' : !NUMERIC_TEXT_RE.test(field));
  }
  return quote ? `"${field.replace(/"/g, '""')}"` : field;
}

/**
 * Picks the quoting policy that reproduces most of the first records (minimal on
 * ties) and remembers the original text of every record it does not reproduce
 * (in space-separated text also every record aligned with extra spaces).
 * Returns undefined when plain minimal quoting reproduces the whole file.
 */
function describeCsvSource(model: CsvModel, body: string, ends: readonly number[]): CsvSource | undefined {
  const { rows, delimiter, newline } = model;
  const collapse = !!model.collapseSpaces;
  if (!rows.length || (!collapse && !body.includes('"'))) {
    return undefined; // nothing is quoted: minimal quoting reproduces every record
  }
  const lastIndex = rows.length - 1;
  const recordStart = (i: number): number => (i ? ends[i - 1] : 0);
  /** End of record i without its line break (every record but an unterminated last one has one). */
  const recordEnd = (i: number): number => {
    const start = recordStart(i);
    const end = ends[i];
    const terminated = i < lastIndex || model.trailingNewline;
    return terminated && end - newline.length >= start && body.startsWith(newline, end - newline.length) ? end - newline.length : end;
  };
  /** Compares without slicing: does `text` equal the source text of record i? */
  const reproduces = (i: number, text: string): boolean => {
    const start = recordStart(i);
    return recordEnd(i) - start === text.length && body.startsWith(text, start);
  };
  let policy: QuotePolicy = 'minimal';
  let best = -1;
  const sample = Math.min(rows.length, CSV_POLICY_SAMPLE_ROWS);
  for (const candidate of QUOTE_POLICIES) {
    let matches = 0;
    for (let i = 0; i < sample; i++) {
      if (reproduces(i, serializeCsvRow(rows[i], delimiter, candidate, collapse))) matches++;
    }
    if (matches > best) {
      best = matches;
      policy = candidate;
    }
    if (matches === sample) break;
  }
  let raw: WeakMap<string[], string> | undefined;
  for (let i = 0; i < rows.length; i++) {
    if (!reproduces(i, serializeCsvRow(rows[i], delimiter, policy, collapse))) {
      (raw ??= new WeakMap()).set(rows[i], body.slice(recordStart(i), recordEnd(i)));
    }
  }
  return policy === 'minimal' && !raw ? undefined : { policy, raw };
}

/** Does the original record text still decode to exactly these fields? */
function rawRecordMatches(raw: string, row: readonly string[], delimiter: string, newline: string, collapse: boolean): boolean {
  if (raw === '') {
    return row.length === 1 && row[0] === '';
  }
  const parsed = readCsvRecords(raw, delimiter, newline, collapse).rows;
  const fields = parsed.length === 1 ? parsed[0] : undefined;
  return !!fields && fields.length === row.length && fields.every((value, j) => value === row[j]);
}

/**
 * CSV text -> records, plus the offset just past each record (including its
 * line break). papaparse reads well-formed text; text papaparse reads unlike
 * Excel (see readsUnlikeExcel) and space-separated text go through
 * readCsvLikeExcel.
 */
function readCsvRecords(text: string, delimiter: string, newline: string, collapse: boolean): { rows: string[][]; ends: number[] } {
  if (collapse || readsUnlikeExcel(text, delimiter)) {
    return readCsvLikeExcel(text, delimiter, newline, collapse);
  }
  const rows: string[][] = [];
  const ends: number[] = [];
  Papa.parse<string[]>(text, {
    ...papaOptions(delimiter, newline),
    step: (result) => {
      rows.push(result.data);
      ends.push(result.meta.cursor);
    },
  });
  return { rows, ends };
}

/**
 * True when papaparse would read the text unlike Excel: a closing quote followed
 * by more text before the delimiter (`"12" monitor`, `"Smith" ,`) - papaparse
 * swallows the rest of the FILE into that field, or trims the spaces -; a quote
 * that never closes (papaparse keeps its '""' escapes); a leading byte-order mark
 * (one more than the stripped one: papaparse drops it, Excel keeps it in A1).
 * Unquoted line breaks are all `newline` here (see unifyCsvNewlines).
 */
function readsUnlikeExcel(text: string, delimiter: string): boolean {
  if (text.charCodeAt(0) === 0xfeff) {
    return true;
  }
  // Jumps from quote to quote: outside quoted fields every delimiter / line break is
  // a real one, so a quote opens a field exactly when one of them precedes it.
  for (let i = text.indexOf('"'); i >= 0; i = text.indexOf('"', i + 1)) {
    const prev = text.charCodeAt(i - 1);
    if (i && prev !== 13 && prev !== 10 && !text.startsWith(delimiter, i - delimiter.length)) {
      continue; // a quote inside unquoted text is a plain character (for both)
    }
    let close = text.indexOf('"', i + 1);
    while (close >= 0 && text.charCodeAt(close + 1) === 34) {
      close = text.indexOf('"', close + 2);
    }
    if (close < 0) {
      return true;
    }
    const next = text.charCodeAt(close + 1);
    if (close + 1 < text.length && next !== 13 && next !== 10 && !text.startsWith(delimiter, close + 1)) {
      return true;
    }
    i = close;
  }
  return false;
}

/**
 * Excel's CSV rules (checked against Excel 16): a field is quoted only when it
 * starts with '"'; inside, '""' is a quote and the next lone '"' ends the
 * quoting; any text after it up to the delimiter is kept as-is (`"12" monitor`
 * -> 12 monitor, `"b"c"d"` -> bc"d"); a quote that never closes takes the rest
 * of the text, line breaks included. Unquoted line breaks are all `newline`.
 * With `collapse` (space-separated, delimiter ' ') a run of spaces is one
 * separator and spaces at the start / end of a record separate nothing: a
 * blank or all-space record is one empty field, like an empty line.
 */
function readCsvLikeExcel(text: string, delimiter: string, newline: string, collapse = false): { rows: string[][]; ends: number[] } {
  const rows: string[][] = [];
  const ends: number[] = [];
  const n = text.length;
  const d0 = delimiter.charCodeAt(0);
  const skipSpaces = (from: number): number => {
    while (text.charCodeAt(from) === 32) from++;
    return from;
  };
  let fields: string[] = [];
  let i = collapse ? skipSpaces(0) : 0;
  for (;;) {
    let value = '';
    let j = i;
    if (text.charCodeAt(i) === 34) {
      let from = i + 1;
      for (;;) {
        const q = text.indexOf('"', from);
        if (q < 0) {
          value += text.slice(from);
          j = n;
          break;
        }
        value += text.slice(from, q);
        if (text.charCodeAt(q + 1) === 34) {
          value += '"';
          from = q + 2;
          continue;
        }
        j = q + 1;
        break;
      }
    }
    let k = j;
    for (; k < n; k++) {
      const ch = text.charCodeAt(k);
      if (ch === 13 || ch === 10 || (ch === d0 && text.startsWith(delimiter, k))) break;
    }
    fields.push(j === k ? value : value + text.slice(j, k));
    if (k < n && text.charCodeAt(k) !== 13 && text.charCodeAt(k) !== 10) {
      if (!collapse) {
        i = k + delimiter.length; // next field of the same record
        continue;
      }
      k = skipSpaces(k);
      if (k < n && text.charCodeAt(k) !== 13 && text.charCodeAt(k) !== 10) {
        i = k; // next field after the run of spaces
        continue;
      }
    }
    const end = k < n ? k + (text.startsWith(newline, k) ? newline.length : 1) : n;
    rows.push(fields);
    ends.push(end);
    if (end >= n) {
      if (k < n) {
        rows.push(['']); // like papaparse: a final line break is followed by an empty record
        ends.push(n);
      }
      break;
    }
    fields = [];
    i = collapse ? skipSpaces(end) : end;
  }
  return { rows, ends };
}

/**
 * One pass over the first MB: line-ending counts and per-record delimiter counts
 * (outside quotes). The delimiter is the candidate that splits the most records
 * consistently into >= 2 fields; ties prefer tab, ';', '|', then ','. With
 * `spaceSeparated` a quote opens a field after a space instead (only the line
 * endings are wanted then).
 */
function scanCsv(text: string, spaceSeparated = false): { delimiter: string; newline: string } {
  const limit = Math.min(text.length, CSV_SCAN_LIMIT);
  const codes = CSV_DELIMITERS.map((d) => d.charCodeAt(0));
  const perRecord: number[][] = CSV_DELIMITERS.map(() => []);
  let current = CSV_DELIMITERS.map(() => 0);
  let crlf = 0;
  let lf = 0;
  let cr = 0;
  let inQuotes = false;
  let fieldStart = true;
  let recordHasContent = false;
  const endRecord = (): void => {
    if (recordHasContent && perRecord[0].length < CSV_SAMPLE_RECORDS) {
      current.forEach((n, k) => perRecord[k].push(n));
    }
    current = CSV_DELIMITERS.map(() => 0);
    recordHasContent = false;
    fieldStart = true;
  };
  for (let i = 0; i < limit; i++) {
    const ch = text.charCodeAt(i);
    if (inQuotes) {
      if (ch === 34) {
        if (text.charCodeAt(i + 1) === 34) i++;
        else inQuotes = false;
      }
      continue;
    }
    if (ch === 13 || ch === 10) {
      if (ch === 13 && text.charCodeAt(i + 1) === 10) {
        crlf++;
        i++;
      } else if (ch === 10) {
        lf++;
      } else {
        cr++;
      }
      endRecord();
      continue;
    }
    recordHasContent = true;
    if (ch === 34 && fieldStart) {
      inQuotes = true;
      fieldStart = false;
      continue;
    }
    const k = codes.indexOf(ch);
    if (k >= 0) {
      current[k]++;
    }
    fieldStart = spaceSeparated ? ch === 32 : k >= 0;
  }
  endRecord();

  let newline = '\n';
  if (crlf >= lf && crlf >= cr && crlf > 0) newline = '\r\n';
  else if (lf >= cr && lf > 0) newline = '\n';
  else if (cr > 0) newline = '\r';

  let best: { delimiter: string; score: number } | undefined;
  for (const delimiter of CSV_DELIMITER_PRIORITY) {
    const counts = perRecord[CSV_DELIMITERS.indexOf(delimiter)];
    const consistency = splitConsistency(counts);
    if (counts.length && consistency >= requiredConsistency(counts.length) && (!best || consistency > best.score + 1e-9)) {
      best = { delimiter, score: consistency };
    }
  }
  return { delimiter: best?.delimiter ?? ',', newline };
}

/**
 * How consistently a delimiter splits sample records, from its count in each:
 * the share of records holding its most common non-zero count (0 when no
 * record holds it).
 */
function splitConsistency(counts: readonly number[]): number {
  const freq = new Map<number, number>();
  for (const n of counts) freq.set(n, (freq.get(n) ?? 0) + 1);
  let modeFreq = 0;
  for (const [n, f] of freq) {
    if (n > 0 && f > modeFreq) modeFreq = f;
  }
  return counts.length ? modeFreq / counts.length : 0;
}

/** The splitConsistency a delimiter needs over `records` sample records (all of them for one or two). */
function requiredConsistency(records: number): number {
  return records <= 2 ? 1 : 0.8;
}

/** Cheap native check: does the text contain line breaks other than `newline` (quoted or not)? */
function hasOtherLineBreaks(text: string, newline: string): boolean {
  if (newline === '\r\n') {
    return /(?:^|[^\r])\n|\r(?!\n)/.test(text);
  }
  return newline === '\n' ? text.includes('\r') : text.includes('\n');
}

/**
 * Rewrites every unquoted line break (CRLF, LF or CR) to `newline`. Line breaks
 * inside quoted fields are part of the value and are left untouched. Used for
 * files with mixed line endings, which then serialize with the dominant one.
 */
function unifyCsvNewlines(text: string, delimiter: string, newline: string): string {
  const parts: string[] = [];
  let start = 0;
  let inQuotes = false;
  let fieldStart = true;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    if (inQuotes) {
      if (ch === 34) {
        if (text.charCodeAt(i + 1) === 34) i++;
        else inQuotes = false;
      }
      continue;
    }
    if (ch === 13 || ch === 10) {
      parts.push(text.slice(start, i), newline);
      if (ch === 13 && text.charCodeAt(i + 1) === 10) i++;
      start = i + 1;
      fieldStart = true;
      continue;
    }
    if (ch === 34 && fieldStart) {
      inQuotes = true;
      fieldStart = false;
      continue;
    }
    fieldStart = delimiter !== '' && text.startsWith(delimiter, i);
    if (fieldStart) i += delimiter.length - 1;
  }
  parts.push(text.slice(start));
  return parts.join('');
}

/** Numeric value of numeric-looking CSV text (Excel-like), else undefined. */
function csvNumber(text: string, delimiter: string): number | undefined {
  if (!text || text.length > 64) {
    return undefined;
  }
  const first = text.charCodeAt(0);
  if (!((first >= 48 && first <= 57) || first === 45 || first === 43 || first === 46 || first === 32 || first === 9)) {
    return undefined;
  }
  let value: number | undefined;
  if (NUMERIC_TEXT_RE.test(text)) {
    value = Number(text.trim());
  } else if (PERCENT_TEXT_RE.test(text)) {
    value = Number(text.trim().slice(0, -1)) / 100;
  } else if (delimiter === ';' && DECIMAL_COMMA_RE.test(text)) {
    value = Number(text.trim().replace(',', '.'));
  } else if (delimiter !== ',' || GROUPED_NUMBER_RE.test(text)) {
    value = GROUPED_NUMBER_RE.test(text) && delimiter !== ';' ? Number(text.trim().replace(/,/g, '')) : undefined;
  }
  return value !== undefined && Number.isFinite(value) ? value : undefined;
}

/**
 * The sheet name Excel 16 gives a CSV file: its base name without the extension, '[' and ']' (not allowed in sheet
 * names) turned into '(' and ')', cut to 31 characters (never inside a surrogate pair).
 */
function csvSheetName(fileName: string): string {
  const base = fileName.replace(/^.*[\\/]/, '').replace(/\.[^.]*$/, '').replace(/\[/g, '(').replace(/\]/g, ')');
  const cut = base.length > SHEET_NAME_MAX && /[\uD800-\uDBFF]/.test(base[SHEET_NAME_MAX - 1]) ? SHEET_NAME_MAX - 1 : SHEET_NAME_MAX;
  return base.slice(0, cut) || 'Sheet1';
}

/**
 * CSV rows as a one-sheet grid. All cells are text, except numeric-looking
 * fields which are numbers (v) that keep their original text (w). With
 * `hasHeader` the first row stays text and is frozen. Column widths are sized
 * to the content of the first rows.
 */
export class CsvGridModel implements GridSource {
  readonly csv: CsvModel;
  readonly fileName: string;
  readonly hasHeader: boolean;
  private readonly meta: WorkbookMeta;

  constructor(csv: CsvModel, fileName: string, hasHeader = false) {
    this.csv = csv;
    this.fileName = fileName;
    this.hasHeader = hasHeader;
    const rows = csv.rows;
    let colCount = 0;
    for (const row of rows) {
      if (row.length > colCount) colCount = row.length;
    }
    const sheet: SheetMeta = {
      index: 0,
      name: csvSheetName(fileName),
      state: 'visible',
      rowCount: Math.max(1, rows.length),
      colCount: Math.max(1, colCount),
      defaultColWidth: DEFAULT_COL_WIDTH_PX,
      defaultRowHeight: ptToPx(DEFAULT_ROW_HEIGHT_PT),
      cols: csvColumnWidths(rows, colCount),
      rows: {},
      merges: [],
      frozen: { rows: hasHeader && rows.length > 1 ? 1 : 0, cols: 0 },
      showGridLines: true,
      zoom: 100,
      conditionalFormats: [],
      images: [],
    };
    this.meta = { sheets: [sheet], styles: [{}], activeSheet: 0, defaultFont: { name: 'Calibri', size: 11 }, date1904: false };
  }

  getMeta(): WorkbookMeta {
    return this.meta;
  }

  getRows(sheet: number, start: number, end: number, c0?: number, c1?: number): { rows: RowData[]; styles: [number, CellStyle][] } {
    const out: RowData[] = [];
    if (sheet !== 0) {
      return { rows: out, styles: [] };
    }
    const rows = this.csv.rows;
    const hi = Math.min(Math.floor(end) || 0, rows.length);
    const [colLo, colHi] = columnWindow(c0, c1);
    for (let r = Math.max(0, Math.floor(start) || 0); r < hi; r++) {
      const row = rows[r];
      if (!row) continue;
      const cells: CellData[] = [];
      const header = this.hasHeader && r === 0;
      const last = Math.min(row.length - 1, colHi);
      for (let c = colLo; c <= last; c++) {
        const text = row[c];
        if (text === undefined || text === null || text === '') continue;
        const n = header ? undefined : csvNumber(text, this.csv.delimiter);
        cells.push(n === undefined ? { c, t: 's', v: text, w: text } : { c, t: 'n', v: n, w: text });
      }
      if (cells.length) out.push({ r, cells });
    }
    return { rows: out, styles: [] };
  }

  getStats(sheet: number, ranges: Range[]): SelectionStats {
    const acc = new StatsAccumulator();
    if (sheet !== 0) {
      return acc.result();
    }
    const rows = this.csv.rows;
    forEachRangeRow(ranges, rows.length, MAX_COLS, (r, c0, c1) => {
      const row = rows[r];
      if (!row) return true;
      const header = this.hasHeader && r === 0;
      const hi = Math.min(c1, row.length - 1);
      for (let c = c0; c <= hi; c++) {
        if (!acc.tick()) return false;
        const text = row[c];
        if (text === undefined || text === null || text === '') continue;
        const n = header ? undefined : csvNumber(text, this.csv.delimiter);
        acc.add(n === undefined ? null : n);
      }
      return true;
    });
    return acc.result();
  }
}

/** Column widths from the first rows' content (px, default width when narrower). */
function csvColumnWidths(rows: readonly string[][], colCount: number): Record<number, ColInfo> {
  const longest = new Array<number>(colCount).fill(0);
  const sample = Math.min(rows.length, 1000);
  for (let r = 0; r < sample; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const text = row[c] ?? '';
      const nl = text.indexOf('\n');
      const len = Math.min(nl >= 0 ? nl : text.length, 60);
      if (len > longest[c]) longest[c] = len;
    }
  }
  const out: Record<number, ColInfo> = {};
  longest.forEach((len, c) => {
    const px = Math.min(480, Math.round(len * MDW + 10));
    if (px > DEFAULT_COL_WIDTH_PX) out[c] = { w: px };
  });
  return out;
}

// ===== LOSSY FEATURES =====

const LOSSY = {
  charts: 'Charts',
  dialogSheets: 'Dialog sheets',
  macroSheets: 'Macro sheets (Excel 4.0)',
  pivots: 'Pivot tables',
  macros: 'Macros (VBA project)',
  slicers: 'Slicers / timelines',
  externalLinks: 'External links',
  threadedComments: 'Threaded comments',
  shapes: 'Shapes / text boxes',
  sparklines: 'Sparklines',
  x14Cf: 'Extended conditional formats (x14)',
  controls: 'Form controls / ActiveX',
} as const;

const LOSSY_ORDER: string[] = Object.values(LOSSY);

/**
 * Features ExcelJS cannot round-trip (they would be lost or degraded on save).
 * Reads the raw zip: part names first, then drawings and (streamed) sheet XML.
 */
export async function detectLossyFeatures(data: Uint8Array): Promise<string[]> {
  try {
    return await scanLossyFeatures(data);
  } catch (err) {
    throw workbookReadError(err);
  }
}

async function scanLossyFeatures(data: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(data);
  const found = new Set<string>();
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const has = (re: RegExp): boolean => names.some((name) => re.test(name.replace(/^\//, '')));
  if (has(/^xl\/charts\/chart[^/]*\.xml$/i) || has(/^xl\/chartsheets\//i)) found.add(LOSSY.charts);
  // sheets that are not worksheets: ExcelJS loads neither, so they get no tab
  if (has(/^xl\/dialogsheets\//i)) found.add(LOSSY.dialogSheets);
  if (has(/^xl\/macrosheets\//i)) found.add(LOSSY.macroSheets);
  if (has(/^xl\/pivot(Tables|Cache)\//i)) found.add(LOSSY.pivots);
  if (has(/^xl\/vbaProject\.bin$/i)) found.add(LOSSY.macros);
  if (has(/^xl\/(slicers|slicerCaches|timelines|timelineCaches)\//i)) found.add(LOSSY.slicers);
  if (has(/^xl\/externalLinks\//i)) found.add(LOSSY.externalLinks);
  if (has(/^xl\/threadedComments\//i)) found.add(LOSSY.threadedComments);
  if (has(/^xl\/(ctrlProps|activeX)\//i)) found.add(LOSSY.controls);

  for (const name of names) {
    const path = name.replace(/^\//, '');
    if (/^xl\/drawings\/[^/]+\.xml$/i.test(path) && !found.has(LOSSY.shapes)) {
      if (drawingHasShapes(await zip.files[name].async('string'))) found.add(LOSSY.shapes);
    } else if (/^xl\/worksheets\/[^/]+\.xml$/i.test(path)) {
      await scanZipEntry(zip.files[name], (text) => {
        if (text.includes('sparklineGroup')) found.add(LOSSY.sparklines);
        if (text.includes('x14:cfRule')) {
          for (const m of text.matchAll(/<x14:cfRule\b[^>]*?\btype\s*=\s*["']([^"']+)["']/g)) {
            if (!X14_CF_ROUND_TRIPPED.has(m[1])) found.add(LOSSY.x14Cf);
          }
        }
        if (/<(?:\w+:)?controls[\s>]/.test(text)) found.add(LOSSY.controls);
        return found.has(LOSSY.sparklines) && found.has(LOSSY.x14Cf) && found.has(LOSSY.controls);
      });
    }
  }
  return LOSSY_ORDER.filter((label) => found.has(label));
}

/**
 * Does a drawing part hold shapes, text boxes, groups or connectors? Not counted:
 * the <mc:Fallback> stand-in shape Excel writes for older versions next to a
 * chartex chart or a slicer, and the twin shape of a form / ActiveX control
 * (marked with a14:compatExt, the id of its legacy VML shape) - those features
 * are reported as such.
 */
function drawingHasShapes(xml: string): boolean {
  const body = xml.replace(/<(?:\w+:)?Fallback\b(?:[^>]*\/>|[\s\S]*?<\/(?:\w+:)?Fallback>)/g, '');
  if (/<(?:\w+:)?(?:grpSp|cxnSp)\b/.test(body)) {
    return true;
  }
  for (const m of body.matchAll(/<(?:\w+:)?sp\b(?:[^>]*\/>|[\s\S]*?<\/(?:\w+:)?sp>)/g)) {
    if (!/<(?:\w+:)?compatExt\b/.test(m[0])) return true;
  }
  return false;
}

/** x14 rule types ExcelJS reads and writes itself (xlsx/xform/sheet/cf-ext): data bars and icon sets. */
const X14_CF_ROUND_TRIPPED: ReadonlySet<string> = new Set(['dataBar', 'iconSet']);

/**
 * Streams a zip entry as text without building one big string. `inspect` sees
 * every chunk prefixed with the tail of the previous one (so short tags split
 * across chunks are still seen) and returns true to stop early.
 */
function scanZipEntry(file: JSZip.JSZipObject, inspect: (text: string) => boolean): Promise<void> {
  return new Promise((resolve, reject) => {
    let tail = '';
    let done = false;
    // ZipObject#internalStream (jszip/lib/zipObject.js) is public at runtime but missing from the typings.
    const stream = (file as unknown as { internalStream(type: 'string'): JSZip.JSZipStreamHelper<string> }).internalStream('string');
    stream
      .on('data', (chunk: string) => {
        if (done) return;
        const text = tail + chunk;
        tail = text.slice(-256);
        if (inspect(text)) {
          done = true;
          stream.pause();
          resolve();
        }
      })
      .on('error', (error: Error) => {
        if (!done) {
          done = true;
          reject(error);
        }
      })
      .on('end', () => {
        if (!done) {
          done = true;
          resolve();
        }
      })
      .resume();
  });
}
