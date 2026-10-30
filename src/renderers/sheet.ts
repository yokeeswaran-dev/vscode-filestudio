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
  /** (v0.2.0, P5) Recalc state of a formula cell; absent = value is current (computed or cached and not dirty). */
  calc?: 'stale' | 'unsupported' | 'circular';
  /** (v0.2.0, P5) Conditional-format paint computed by the host for this cell (absent = no rule applies). */
  cf?: CfPaint;
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
  /** (v0.2.0, P2) Stable id; survives rename, reorder, undo. Optional only for type compatibility; the host always sets it. */
  id?: SheetId;
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

/** (v0.2.0) Stable sheet id: xlsx = 's' + the <sheet sheetId="n"> attribute of workbook.xml (= ExcelJS ws.id); csv = 'csv'. */
export type SheetId = string;

/** (v0.2.0, P5) Result of the conditional formats of one cell. `s` replaces CellData.s for painting (cell style + dxf merged). */
export interface CfPaint {
  s?: number;
  /** Colour-scale fill, '#rrggbb'. */
  fill?: string;
  bar?: { pct: number; color: string; negative?: boolean; gradient: boolean; showValue: boolean };
  icon?: { set: string; index: number; showValue: boolean };
}

// ----- edit intents -----
// (v0.2.0) What the webview asks the host to change. Declared once here (protocol frozen after track F, D-062) and
// shared with media/viewer.js as JSDoc typedefs. Shapes and limits are validated in parseWebviewMessage
// (src/viewerProvider.ts); the meaning is checked again by the phase's planner (planEdit / planCsvEdit / md checks).

/** What the webview asks for. Validated in parseWebviewMessage (shape + limits) and again by planEdit (meaning). */
export type EditIntent =
  // ---- values (xlsx P2, csv P3) ----
  | { kind: 'setCells'; sheetId: SheetId; cells: [r: number, c: number, text: string][] } // typing, paste, fill, replace one
  | { kind: 'clear'; sheetId: SheetId; ranges: Range[]; what: 'contents' | 'formats' | 'all' } // csv: 'contents' only
  // ---- structure (csv P3, xlsx P4) ----
  | { kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; sheetId: SheetId; at: number; count: number }
  | { kind: 'sort'; sheetId: SheetId; range: Range; key: number; desc: boolean; header: boolean } // key = absolute column
  | { kind: 'replaceAll'; query: FindQuery; replace: string } // scope inside the query
  // ---- xlsx only (P4) ----
  | { kind: 'style'; sheetId: SheetId; ranges: Range[]; delta: StyleDelta }
  | { kind: 'merge' | 'unmerge'; sheetId: SheetId; range: Range }
  | { kind: 'resize'; sheetId: SheetId; axis: 'row' | 'col'; index: number[]; size: number | null } // px; null = default/auto height
  | { kind: 'sheetAdd'; index: number; name?: string }
  | { kind: 'sheetRename'; sheetId: SheetId; name: string }
  | { kind: 'sheetDelete' | 'sheetDuplicate'; sheetId: SheetId }
  | { kind: 'sheetMove'; sheetId: SheetId; index: number }
  | { kind: 'sheetHide'; sheetId: SheetId; hidden: boolean }
  | { kind: 'tabColor'; sheetId: SheetId; argb: string | null }
  // ---- markdown (P6) ----
  | { kind: 'mdText'; edits: MdTextEdit[] };

export type EditKind = EditIntent['kind'];

export interface FindQuery {
  /** csv: always 'csv'. */
  scope: SheetId | 'workbook';
  /** At most 1,000 characters. */
  text: string;
  matchCase: boolean;
  regex: boolean;
  wholeCell: boolean;
  /** csv: 'values'. */
  lookIn: 'values' | 'formulas';
}

/** Applied in order on the cell's ORIGINAL xf (never on an ExcelJS style object). null = remove that property. */
export interface StyleDelta {
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double' | null;
  strike?: boolean;
  /** 'FFRRGGBB'. */
  fontColor?: string | null;
  /** 'FFRRGGBB'. */
  fill?: string | null;
  border?: { top?: BorderSide; bottom?: BorderSide; left?: BorderSide; right?: BorderSide };
  hAlign?: 'left' | 'center' | 'right' | 'justify' | null;
  vAlign?: 'top' | 'center' | 'bottom' | null;
  wrap?: boolean;
  /** OOXML number format code, already escaped. */
  numFmt?: string;
  /** Set by typing a leading apostrophe. */
  quotePrefix?: boolean;
}

export type BorderSide = { style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'; color?: string } | null;

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
//   EDIT TYPES         (v0.2.0) file-side edit types: FileValue, CellPatch, PackageOp, BlockReason (host only)
//   EXCELJS INTERNALS  the (undocumented) slice of ExcelJS' object model we read
//   CONSTANTS          units, limits and builtin tables
//   COLORS             theme palette, tint maths, indexed palette, resolveColor
//   ADDRESSES          A1 helpers, range parsing, shared-formula translation; (v0.2.0) fx formula tokenizer adapter
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
// Excel formula tokenizer, xlsx flavour (file syntax: [n] external workbook indexes, _xlfn. prefixes). See ADDRESSES
// `fx adapter`.
import { parseA1Ref as fxParseA1Ref, tokenize as fxTokenizeRaw, tokenTypes as FX_TOKEN } from '@borgar/fx/xlsx';
import type { MdTextEdit } from './markdown';

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

// ===== EDIT TYPES =====
// (v0.2.0) What an edit writes into the file. Host only (never sent to the webview). The save path (XLSX SAVE, P1)
// patches the loaded package with these operations; every phase that changes a workbook produces them.

/** A cell value as the file stores it. */
export type FileValue = { t: 'n'; v: number } | { t: 's'; v: string } | { t: 'b'; v: boolean } | { t: 'e'; v: string };

export interface CellPatch {
  /** null = clear contents (keep s). */
  value?: FileValue | null;
  /** FILE syntax (_xlfn. prefixes, English names, ','), without '='. */
  formula?: string;
  /** <v> of a formula cell; null = write no <v> (unsupported function, D-021). */
  cached?: FileValue | null;
  /** Derived xf appended to styles.xml (existing xfs never change). */
  style?: StyleDelta[];
}

/** `sheet` = the numeric <sheet sheetId> attribute (SheetId = 's' + sheet). Positions are those at edit time. */
export type PackageOp =
  | { kind: 'cells'; sheet: number; cells: [r: number, c: number, patch: CellPatch][] }
  | { kind: 'merge' | 'unmerge'; sheet: number; range: Range }
  | { kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; sheet: number; at: number; count: number }
  | { kind: 'sort'; sheet: number; range: Range; byRow: Int32Array } // new row range.r0+i <- old row byRow[i]
  | { kind: 'rowProps'; sheet: number; rows: number[]; height?: number | null; style?: StyleDelta[] } // points
  | { kind: 'colProps'; sheet: number; cols: number[]; width?: number | null; style?: StyleDelta[] } // Excel char units
  | { kind: 'sheetAdd'; sheet: number; name: string; index: number }
  | { kind: 'sheetRename'; sheet: number; name: string }
  | { kind: 'sheetDelete'; sheet: number }
  | { kind: 'sheetDuplicate'; sheet: number; newSheet: number; name: string; index: number }
  | { kind: 'sheetMove'; sheet: number; index: number }
  | { kind: 'sheetState'; sheet: number; state: 'visible' | 'hidden' }
  | { kind: 'tabColor'; sheet: number; argb: string | null };

/** Why an edit is refused (editResult code 'blocked'); used by every phase's block rules. */
export interface BlockReason {
  code: 'ref' | 'array' | 'merge' | 'table' | 'rule' | 'name' | 'lossy' | 'protected' | 'unsupported' | 'sheet3d' | 'limit';
  /** Full user text (contract §3.4 style). */
  message: string;
  /** E.g. ['Data!H1', 'F!A2'] (first 12, then "+N more"). */
  where: string[];
}

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

// ----- fx adapter -----
// (v0.2.0) Formula text <-> tokens through @borgar/fx (xlsx flavour = file syntax: A1 references only, R1C1 off,
// [n] external workbook indexes, _xlfn. prefixes). Used by formula input (P2), reference shifting (P4, display and
// save, one grammar) and recalc (P5). Rules:
//   - the tokens cover the formula without gaps or overlaps: joining their text always gives the formula back
//     (fxJoin); text fx cannot read becomes an 'UNKNOWN' token (callers that rewrite a formula must block on it);
//   - a reference is re-spelled only where it changes (fxEmitRef): an unchanged prefix keeps its text (quotes,
//     case), an unchanged range end keeps its text ('a1', '$b$2'), an unchanged reference keeps all of it (B5:A1).
// Ported from ROUND\s2-editmodel\proto\lex-fx.js (E6 fuzz 90,000/90,000, E7 corpus 532/532).

/** A sheet / workbook prefix of a reference ('My Data'!, [1]Sheet1!, Sheet1:Sheet3!, [1]!), unescaped. */
export interface FxPrefix {
  /** The prefix was written in quotes ('Sheet1'!A1). */
  quoted: boolean;
  /** External workbook as written: '[1]' (file syntax: index into <externalReferences>), else null. */
  book: string | null;
  /** First (or only) sheet; null for a workbook-level name ([1]!Name). */
  sheet1: string | null;
  /** Last sheet of a 3D reference (Sheet1:Sheet3!A1), else null. */
  sheet2: string | null;
}

/**
 * The rectangle a reference names: 0-based, inclusive. Whole columns ('cols', A:C) run over every row and whole rows
 * ('rows', 2:4) over every column; their unused `$` flags are false.
 */
export interface FxArea {
  kind: 'cell' | 'area' | 'rows' | 'cols';
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  $r0: boolean;
  $c0: boolean;
  $r1: boolean;
  $c1: boolean;
}

export interface FxRef extends FxArea {
  prefix: FxPrefix | null;
  /** Followed by the spill operator (A1#); the '#' is part of the token. */
  spill: boolean;
  /** Offset in the formula where the A1 part starts (after the prefix and its '!'). */
  bodyStart: number;
}

/**
 * One formula token: `formula.slice(start, end)` is its text. `type`: 'REF' (cell, range, whole rows / columns, with
 * `ref`), 'NAME' (defined name, `prefix` when qualified), 'STRUCT' (table reference, `prefix` when qualified), 'ERR'
 * (a qualified #REF!: Sheet1!#REF!, with `prefix`), 'UNKNOWN' (unreadable text), or fx's own token type for the rest
 * ('func', 'operator', 'number', 'string', 'bool', 'error', 'whitespace', 'newline', 'fx_prefix').
 */
export interface FxToken {
  type: string;
  start: number;
  end: number;
  ref?: FxRef;
  prefix?: FxPrefix | null;
  refError?: boolean;
}

const FX_TOKENIZE_OPTIONS = { withLocation: true, mergeRefs: true, negativeNumbers: false, allowTernary: false, r1c1: false };

/** "'[1]My ''x'''" -> {quoted, book '[1]', sheet1 "My 'x'"}; 'Sheet1:Sheet3' -> 3D. Text before the '!'. */
function fxParsePrefix(text: string): FxPrefix {
  const quoted = text.length >= 2 && text.startsWith("'") && text.endsWith("'");
  let inner = quoted ? text.slice(1, -1).replace(/''/g, "'") : text;
  let book: string | null = null;
  if (inner.startsWith('[')) {
    const close = inner.indexOf(']');
    if (close > 0) {
      book = inner.slice(0, close + 1);
      inner = inner.slice(close + 1);
    }
  }
  // Excel forbids ':' in sheet names, so a ':' separates the two ends of a 3D reference.
  const colon = inner.indexOf(':');
  const sheet1 = (colon >= 0 ? inner.slice(0, colon) : inner) || null;
  const sheet2 = colon >= 0 ? inner.slice(colon + 1) || null : null;
  return { quoted, book, sheet1, sheet2 };
}

/** fx range of a reference -> FxArea of `kind`. */
function fxToArea(
  range: { top: number | null; left: number | null; bottom?: number | null; right?: number | null; $top?: boolean | null; $left?: boolean | null; $bottom?: boolean | null; $right?: boolean | null },
  kind: FxArea['kind'],
): FxArea | undefined {
  const r0 = range.top ?? 0;
  const c0 = range.left ?? 0;
  const r1 = range.bottom ?? range.top ?? 0;
  const c1 = range.right ?? range.left ?? 0;
  if (kind === 'cols') {
    if (range.left == null) return undefined;
    return { kind, r0: 0, r1: MAX_ROWS - 1, c0, c1, $r0: false, $r1: false, $c0: !!range.$left, $c1: !!range.$right };
  }
  if (kind === 'rows') {
    if (range.top == null) return undefined;
    return { kind, r0, r1, c0: 0, c1: MAX_COLS - 1, $r0: !!range.$top, $r1: !!range.$bottom, $c0: false, $c1: false };
  }
  if (range.top == null || range.left == null) return undefined;
  return {
    kind,
    r0,
    c0,
    r1,
    c1,
    $r0: !!range.$top,
    $c0: !!range.$left,
    $r1: !!(range.$bottom ?? range.$top),
    $c1: !!(range.$right ?? range.$left),
  };
}

/**
 * Tokenizes a formula (with or without the leading '='). Never throws: if fx fails, or its tokens do not cover the
 * text exactly, the result is one 'UNKNOWN' token for the whole formula.
 */
/** Longest formula fxTokenize reads (Excel's limit). @borgar/fx is quadratic on some input (apostrophe runs), and a file's formula text is not limited. */
export const FX_MAX_FORMULA_CHARS = 8192;

/** Runs no real formula has (a sheet name is at most 31 characters): they make @borgar/fx quadratic. */
const FX_HOSTILE_RUN = /['[]{64,}|\.{64,}|(?:\d\.){64,}/;
/** Time budget shared by all calls: slow quote-heavy formulas add their time as debt; the debt shrinks with the clock. */
const FX_DEBT_LIMIT_MS = 1000;
const FX_SLOW_CALL_MS = 15;
let fxDebtMs = 0;
let fxDebtAt = 0;
function fxNow(): number {
  return typeof performance !== 'undefined' ? performance.now() : Date.now();
}

/** The formula with the body of every "text" literal blanked (same length): a dotted leader in a string is harmless. */
function fxMaskStrings(formula: string): string {
  if (formula.indexOf('"') < 0) return formula;
  return formula.replace(/"(?:[^"]|"")*"?/g, (m) => '"' + ' '.repeat(Math.max(0, m.length - 2)) + (m.length > 1 ? '"' : ''));
}

export function fxTokenize(formula: string): FxToken[] {
  if (!formula) return [];
  const whole: FxToken[] = [{ type: 'UNKNOWN', start: 0, end: formula.length }];
  if (formula.length > FX_MAX_FORMULA_CHARS || FX_HOSTILE_RUN.test(fxMaskStrings(formula))) return whole;
  // Shared budget: the adapter itself stops a workbook of crafted formulas from freezing the host. Only quote-heavy
  // text can be slow; once the debt is over the limit such a formula is opaque (one UNKNOWN token).
  const t0 = fxNow();
  fxDebtMs = Math.max(0, fxDebtMs - (t0 - fxDebtAt));
  fxDebtAt = t0;
  const quoteHeavy = formula.length > 256 && (formula.match(/'/g)?.length ?? 0) > 8;
  if (quoteHeavy && fxDebtMs > FX_DEBT_LIMIT_MS) return whole;
  let raw: { type: string; value: string; loc?: number[] }[];
  try {
    raw = fxTokenizeRaw(formula, FX_TOKENIZE_OPTIONS);
  } catch {
    return whole;
  } finally {
    const spent = fxNow() - t0;
    if (quoteHeavy && spent > FX_SLOW_CALL_MS) fxDebtMs += spent;
    fxDebtAt = fxNow();
  }
  const out: FxToken[] = [];
  let pos = 0;
  for (let i = 0; i < raw.length; i++) {
    const t = raw[i];
    const start = t.loc?.[0];
    const end = t.loc?.[1];
    if (start !== pos || end === undefined || end <= start || formula.slice(start, end) !== t.value) return whole;
    pos = end;
    const value = t.value;
    if (t.type === FX_TOKEN.REF_RANGE || t.type === FX_TOKEN.REF_BEAM) {
      const ref = fxParseA1Ref(value, { allowNamed: false, allowTernary: false });
      const range = ref && 'range' in ref ? ref.range : undefined;
      const bang = value.lastIndexOf('!'); // the A1 part never contains '!'; a quoted sheet name may
      const body = value.slice(bang + 1);
      const kind: FxArea['kind'] =
        t.type === FX_TOKEN.REF_BEAM ? (range?.top == null ? 'cols' : 'rows') : body.includes(':') ? 'area' : 'cell';
      const area = range ? fxToArea(range, kind) : undefined;
      if (!area) {
        out.push({ type: 'UNKNOWN', start, end });
        continue;
      }
      let tokenEnd = end;
      let spill = false;
      const next = raw[i + 1];
      if (next && next.type === FX_TOKEN.OPERATOR && next.value === '#' && next.loc?.[0] === end && next.loc[1] === end + 1) {
        spill = true;
        tokenEnd = end + 1;
        pos = tokenEnd;
        i++;
      }
      const prefix = bang >= 0 ? fxParsePrefix(value.slice(0, bang)) : null;
      out.push({ type: 'REF', start, end: tokenEnd, ref: { ...area, prefix, spill, bodyStart: start + bang + 1 } });
    } else if (t.type === FX_TOKEN.REF_NAMED || t.type === FX_TOKEN.REF_STRUCT) {
      // A table reference's own brackets come after the prefix: split at the last '!' before the first '[' of the
      // table part ('[1]Sh'!Table1[Col], [@x]); a leading '[n]' is an external book.
      const bang = fxPrefixEnd(value);
      out.push({
        type: t.type === FX_TOKEN.REF_NAMED ? 'NAME' : 'STRUCT',
        start,
        end,
        prefix: bang > 0 ? fxParsePrefix(value.slice(0, bang)) : null,
      });
    } else if (
      (t.type === FX_TOKEN.CONTEXT || t.type === FX_TOKEN.CONTEXT_QUOTE) &&
      raw[i + 1]?.type === FX_TOKEN.OPERATOR &&
      raw[i + 1].value === '!' &&
      raw[i + 2]?.type === FX_TOKEN.ERROR &&
      raw[i + 2].value.toUpperCase() === '#REF!' &&
      raw[i + 1].loc?.[0] === end &&
      raw[i + 2].loc?.[0] === end + 1
    ) {
      // Sheet1!#REF! (a reference Excel already broke): one token, so a sheet rename can still rewrite its prefix.
      const errEnd = end + 1 + raw[i + 2].value.length;
      if (formula.slice(end + 1, errEnd) !== raw[i + 2].value) return whole;
      out.push({ type: 'ERR', start, end: errEnd, prefix: fxParsePrefix(value), refError: true });
      pos = errEnd;
      i += 2;
    } else if (t.type === FX_TOKEN.UNKNOWN) {
      out.push({ type: 'UNKNOWN', start, end });
    } else {
      out.push({ type: t.type, start, end });
    }
  }
  if (pos !== formula.length) return whole;
  return out;
}

/** Unquoted prefix of a name / table token: optional '[n]' book, sheet text, '!' ('[1]!Name', 'Sheet1!Name'). */
const FX_PLAIN_PREFIX_RE = /^(?:\[[^\]]*\])?[^[\]'!]*!/;

/**
 * Offset of the '!' that ends the prefix of a name / table token, or -1 when it has none. A table reference's own
 * brackets ('Table1[Col]', '[@x]', '[[#This Row],[a]]') never count as a prefix, nor does a '!' inside them.
 */
function fxPrefixEnd(value: string): number {
  if (value.startsWith("'")) {
    for (let i = 1; i < value.length; i++) {
      if (value[i] !== "'") continue;
      if (value[i + 1] === "'") {
        i++;
        continue;
      }
      return value[i + 1] === '!' ? i + 1 : -1;
    }
    return -1;
  }
  const m = FX_PLAIN_PREFIX_RE.exec(value);
  return m ? m[0].length - 1 : -1;
}

/** Joins the tokens' text: always equals the formula they came from. */
export function fxJoin(formula: string, tokens: readonly FxToken[]): string {
  let out = '';
  for (const t of tokens) out += formula.slice(t.start, t.end);
  return out;
}

/** Excel's rule for sheet names in references (checked against Excel's own output: 33 names, s2-editmodel E1). */
export function sheetNeedsQuotes(name: string): boolean {
  if (!/^[\p{L}_\\][\p{L}0-9_.\\]*$/u.test(name)) return true; // spaces, punctuation, leading digit or dot, digits of other scripts (Excel quotes x٤)
  if (/^(TRUE|FALSE)$/i.test(name)) return true;
  const a1 = /^([A-Za-z]{1,3})([0-9]+)$/.exec(name); // looks like a cell (column <= XFD)
  if (a1 && colToIndex(a1[1]) < MAX_COLS && +a1[2] >= 1 && +a1[2] <= MAX_ROWS) return true;
  if (/^(R|C|RC|R[0-9]+|C[0-9]+|R[0-9]*C[0-9]*)$/i.test(name)) return true; // looks like R1C1
  return false;
}

/** A sheet name as a reference prefix writes it (without the '!'): quoted only when Excel would. */
export function quoteSheetForRef(name: string): string {
  return sheetNeedsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

/** Prefix text including the '!'. `quoted` keeps quotes that were written; otherwise they follow sheetNeedsQuotes. */
export function fxPrefixText(p: FxPrefix): string {
  if (p.sheet1 === '' || p.sheet2 === '' || p.book === '') throw new RangeError('fx: empty sheet name in a reference prefix');
  const inner = (p.book ?? '') + (p.sheet1 ?? '') + (p.sheet2 ? ':' + p.sheet2 : '');
  const needs =
    p.quoted ||
    (p.book !== null && /[\\/ ']/.test(p.book)) ||
    (p.sheet1 !== null && sheetNeedsQuotes(p.sheet1)) ||
    (p.sheet2 !== null && sheetNeedsQuotes(p.sheet2));
  return (needs ? `'${inner.replace(/'/g, "''")}'` : inner) + '!';
}

/** Throws a RangeError for a coordinate that is not an integer inside the grid (the caller must write #REF! itself); swapped corners are put in order. */
function fxCheckedArea(a: FxArea): FxArea {
  const ok = (n: number, max: number) => Number.isInteger(n) && n >= 0 && n < max;
  const usesRow = a.kind !== 'cols';
  const usesCol = a.kind !== 'rows';
  const two = a.kind !== 'cell';
  if (
    (usesRow && !(ok(a.r0, MAX_ROWS) && (!two || ok(a.r1, MAX_ROWS)))) ||
    (usesCol && !(ok(a.c0, MAX_COLS) && (!two || ok(a.c1, MAX_COLS))))
  ) {
    throw new RangeError('fx: reference outside the grid');
  }
  const swapR = two && usesRow && a.r0 > a.r1;
  const swapC = two && usesCol && a.c0 > a.c1;
  if (!swapR && !swapC) return a;
  return {
    ...a,
    r0: swapR ? a.r1 : a.r0, r1: swapR ? a.r0 : a.r1, $r0: swapR ? a.$r1 : a.$r0, $r1: swapR ? a.$r0 : a.$r1,
    c0: swapC ? a.c1 : a.c0, c1: swapC ? a.c0 : a.c1, $c0: swapC ? a.$c1 : a.$c0, $c1: swapC ? a.$c0 : a.$c1,
  };
}

/** One end of a reference: 'B3', '$B$3', 'B' (cols), '3' (rows). */
function fxEndText(kind: FxArea['kind'], r: number, c: number, $r: boolean, $c: boolean): string {
  const col = (abs: boolean) => (abs ? '$' : '') + indexToCol(c);
  const row = (abs: boolean) => (abs ? '$' : '') + (r + 1);
  if (kind === 'cols') return col($c);
  if (kind === 'rows') return row($r);
  return col($c) + row($r);
}

/** Reference text of an area without prefix: 'A1', 'A1:B2', 'A:C', '2:4' (normalized, upper-case). */
export function fxAreaText(a: FxArea): string {
  const first = fxEndText(a.kind, a.r0, a.c0, a.$r0, a.$c0);
  return a.kind === 'cell' ? first : `${first}:${fxEndText(a.kind, a.r1, a.c1, a.$r1, a.$c1)}`;
}

const FX_END_RE = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d{1,7})?$/;

function samePrefix(a: FxPrefix | null, b: FxPrefix | null): boolean {
  if (a === null || b === null) return a === b;
  return a.book === b.book && a.sheet1 === b.sheet1 && a.sheet2 === b.sheet2;
}

function sameArea(a: FxArea, b: FxArea): boolean {
  return (
    a.kind === b.kind && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1 &&
    a.$r0 === b.$r0 && a.$c0 === b.$c0 && a.$r1 === b.$r1 && a.$c1 === b.$c1
  );
}

/**
 * New text of a 'REF' token. `next.prefix` (null = no prefix) and `next.area` replace the token's prefix and range;
 * undefined keeps them. Unchanged parts keep their original spelling: the whole prefix text when the new prefix names
 * the same book and sheets, the whole range text when the area is unchanged, and otherwise each end whose row, column
 * and `$` flags are unchanged. The spill '#' is kept.
 */
export function fxEmitRef(formula: string, token: FxToken, next: { prefix?: FxPrefix | null; area?: FxArea }): string {
  const ref = token.ref;
  if (!ref) return formula.slice(token.start, token.end);
  const bodyEnd = token.end - (ref.spill ? 1 : 0);
  const oldPrefixText = formula.slice(token.start, ref.bodyStart);
  const oldBody = formula.slice(ref.bodyStart, bodyEnd);
  let prefixText = oldPrefixText;
  if (next.prefix !== undefined && !samePrefix(next.prefix, ref.prefix)) {
    prefixText = next.prefix === null ? '' : fxPrefixText(next.prefix);
  }
  let body = oldBody;
  const area = next.area && fxCheckedArea(next.area);
  if (area && !sameArea(area, ref)) {
    const ends = oldBody.split(':');
    const sameShape = area.kind === ref.kind && ends.length === (area.kind === 'cell' ? 1 : 2);
    /** The old text of end `k` when it already says (r, c, $r, $c), else the new text. */
    const keep = (k: 0 | 1, r: number, c: number, $r: boolean, $c: boolean): string => {
      const m = sameShape ? FX_END_RE.exec(ends[k]) : null;
      if (m) {
        const hasCol = m[2] !== undefined;
        const hasRow = m[4] !== undefined;
        const colAbs = hasCol && m[1] === '$';
        const rowAbs = hasCol ? m[3] === '$' : m[1] === '$' || m[3] === '$';
        const colOk = area.kind === 'rows' ? !hasCol : hasCol && colToIndex(m[2]) === c && colAbs === $c;
        const rowOk = area.kind === 'cols' ? !hasRow : hasRow && parseInt(m[4], 10) - 1 === r && rowAbs === $r;
        if (colOk && rowOk) return ends[k];
      }
      return fxEndText(area.kind, r, c, $r, $c);
    };
    const first = keep(0, area.r0, area.c0, area.$r0, area.$c0);
    body = area.kind === 'cell' ? first : `${first}:${keep(1, area.r1, area.c1, area.$r1, area.$c1)}`;
  }
  return prefixText + body + (ref.spill ? '#' : '');
}

/** Rebuilds a formula from its tokens: `replace` returns new text for a token, or null to keep its text. */
export function fxEmit(formula: string, tokens: readonly FxToken[], replace: (token: FxToken) => string | null): string {
  let out = '';
  for (const t of tokens) out += replace(t) ?? formula.slice(t.start, t.end);
  return out;
}

// ----- formula syntax -----
// (v0.2.0, P2) Formula INPUT: what the user types in the cell editor (editor syntax: English function names, ',' as
// separator, no _xlfn. prefixes) -> what the file stores (file syntax: the same text with the _xlfn. prefixes Excel
// writes for functions newer than Excel 2007). Built on the fx adapter above. A formula that fx cannot parse is refused
// BEFORE it can reach the file: Excel would "repair" a workbook with a broken <f>.
// (The import sits here, not in LIBRARIES, so this section never touches lines other tracks edit.)
import { parse as fxParseAst } from '@borgar/fx/xlsx';

/** Functions Excel stores with a prefix (names newer than Excel 2007): NAME -> prefix. Checked against Excel's own files (see test). */
const FILE_FUNCTION_PREFIX: ReadonlyMap<string, string> = new Map(
  (
    'ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR ' +
    'BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST COMBINA CONCAT ' +
    'CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL ENCODEURL ERF.PRECISE ERFC.PRECISE ' +
    'EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT ' +
    'FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS ' +
    'HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV ' +
    'MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE ' +
    'PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC ' +
    'QUARTILE.INC QUERYSTRING RANK.AVG RANK.EQ RRI SEC SECH SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT ' +
    'T.INV T.INV.2T T.TEST TEXTAFTER TEXTBEFORE TEXTJOIN UNICHAR UNICODE VALUETOTEXT VAR.P VAR.S WEBSERVICE WEIBULL.DIST ' +
    'WORKDAY.INTL XLOOKUP XMATCH XOR Z.TEST'
  )
    .split(' ')
    .map((name): [string, string] => [name, '_xlfn.']),
);
// Functions the file stores with the worksheet-function namespace too (used by the spilling functions, which are refused
// for input below; kept for the tests that read Excel's own formulas back).
const FILE_FUNCTION_PREFIX_WS: ReadonlyMap<string, string> = new Map([['FILTER', '_xlfn._xlws.'], ['SORT', '_xlfn._xlws.']]);

/** Functions whose result spills into neighbouring cells (Excel 365): refused in new formulas (§5.3). */
const SPILLING_FUNCTIONS: ReadonlySet<string> = new Set(
  'SEQUENCE FILTER SORT SORTBY UNIQUE RANDARRAY TEXTSPLIT VSTACK HSTACK TOCOL TOROW TAKE DROP EXPAND WRAPROWS WRAPCOLS CHOOSEROWS CHOOSECOLS TRANSPOSE'.split(' '),
);
/** Functions whose parameters are written with special file prefixes (_xlpm.): refused in new formulas. */
const LAMBDA_FUNCTIONS: ReadonlySet<string> = new Set('LET LAMBDA MAP REDUCE SCAN BYROW BYCOL MAKEARRAY ISOMITTED'.split(' '));

export const FORMULA_PROBLEM_TEXT = "There's a problem with this formula.";
const FORMULA_LIMITS = { depth: 64, args: 255, text: 255 } as const;

export type FormulaCheck =
  | { ok: true; /** File syntax, without '='. */ file: string; /** Editor / display syntax, without '='. */ editor: string }
  | { ok: false; reason: 'invalid'; message: string }
  | { ok: false; reason: 'blocked'; block: BlockReason };

/** A name that looks like a cell reference outside the grid (A1048577, XFE1): refused, as Excel's formula bar would. */
const OUT_OF_GRID_NAME_RE = /^\$?[A-Za-z]{1,3}\$?\d{1,8}$/;

function hasLoneSurrogate(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const d = text.charCodeAt(i + 1);
      if (d >= 0xdc00 && d <= 0xdfff) i++;
      else return true;
    } else if (c >= 0xdc00 && c <= 0xdfff) return true;
  }
  return false;
}

/**
 * Checks a formula typed into a cell (`body` = text after the '=') and converts it to file syntax. Refused as
 * 'invalid' (cell keeps its content): not parseable, unbalanced quotes / parentheses, a reference outside the grid or
 * to an unknown sheet, over Excel's limits (8,192 characters, 64 nesting levels, 255 arguments, 255 characters per
 * text constant). Refused as 'blocked': what FileStudio cannot write safely yet (spilling functions, LET / LAMBDA,
 * references to other workbooks, table references, spill references `A1#`). Unknown function names are allowed (the
 * cell shows #NAME? like in Excel). `sheetNames`: every sheet of the workbook (compared ignoring case).
 */
export function checkFormulaInput(body: string, sheetNames: readonly string[]): FormulaCheck {
  const invalid = (): FormulaCheck => ({ ok: false, reason: 'invalid', message: FORMULA_PROBLEM_TEXT });
  const blocked = (message: string): FormulaCheck => ({ ok: false, reason: 'blocked', block: { code: 'unsupported', message, where: [] } });
  if (!body.trim() || body.length > FX_MAX_FORMULA_CHARS || body.trimStart().startsWith('=')) return invalid();
  if (/[\u0000-\u0008\u000b-\u001f￾￿]/.test(body) || hasLoneSurrogate(body)) return invalid();
  const tokens = fxTokenize(body);
  if (tokens.some((t) => t.type === 'UNKNOWN')) return invalid();
  const known = new Set(sheetNames.map((n) => n.toLowerCase()));
  const text = (t: FxToken): string => body.slice(t.start, t.end);
  // fx's grammar is lenient at the ends: a formula starts with an operand, '(', '{' or a sign, and does not end on a binary operator.
  const significant = tokens.filter((t) => t.type !== 'whitespace' && t.type !== 'newline');
  const first = significant[0];
  const last = significant[significant.length - 1];
  if (first.type === 'operator' && !'({+-'.includes(text(first))) return invalid();
  if (last.type === 'operator' && !')}%#'.includes(text(last))) return invalid();
  const bounds = (a: FxArea): boolean => a.r0 >= 0 && a.c0 >= 0 && a.r1 < MAX_ROWS && a.c1 < MAX_COLS;
  // Meaning of every token (before the grammar check, so a blocked formula says why).
  for (const t of tokens) {
    if (t.type === 'func') {
      // _xlfn. / _xlws. are the file prefixes of newer functions: typed by hand they must not get past the checks.
      const name = text(t).toUpperCase().replace(/^(?:_XLFN\.|_XLWS\.)+/, '');
      if (SPILLING_FUNCTIONS.has(name)) return blocked("FileStudio can't create spilling (dynamic array) formulas yet. Enter this formula in Excel.");
      if (LAMBDA_FUNCTIONS.has(name)) return blocked("FileStudio can't create formulas with LET or LAMBDA yet. Enter this formula in Excel.");
    } else if (t.type === 'operator' && text(t) === '@') {
      return invalid();
    } else if (t.type === 'REF' && t.ref) {
      if (t.ref.spill) return blocked("FileStudio can't create spilling (dynamic array) formulas yet. Enter this formula in Excel.");
      if (!bounds(t.ref)) return invalid();
    } else if (t.type === 'STRUCT') {
      return blocked("FileStudio can't create formulas with table references yet. Enter this formula in Excel.");
    } else if (t.type === 'NAME' && !t.prefix && OUT_OF_GRID_NAME_RE.test(text(t))) {
      return invalid();
    } else if (t.type === 'string') {
      if (text(t).length - 2 - (text(t).match(/""/g)?.length ?? 0) > FORMULA_LIMITS.text) return invalid();
    } else if (t.type === 'number') {
      if (!Number.isFinite(Number(text(t)))) return invalid();
    }
    const prefix = t.ref ? t.ref.prefix : t.prefix;
    if ((t.type === 'REF' || t.type === 'NAME' || t.type === 'STRUCT' || t.type === 'ERR') && prefix) {
      if (prefix.book !== null) return blocked("FileStudio can't create formulas that refer to other workbooks yet. Enter this formula in Excel.");
      for (const sheet of [prefix.sheet1, prefix.sheet2]) if (sheet !== null && !known.has(sheet.toLowerCase())) return invalid();
    }
  }
  // Grammar (fx parser) and Excel's nesting limits.
  try {
    // The parser wants merged references and negative numbers as one token (unlike the adapter's own tokens).
    const raw = fxTokenizeRaw(body, { mergeRefs: true, negativeNumbers: true, allowTernary: false, r1c1: false });
    fxParseAst(raw);
    const stack: { func: boolean; args: number }[] = [];
    let prev = '';
    for (const r of raw) {
      if (r.type === 'operator') {
        if (r.value === '(' || r.value === '{') {
          stack.push({ func: r.value === '(' && prev === 'func', args: 1 });
          if (stack.length > FORMULA_LIMITS.depth) return invalid();
        } else if (r.value === ')' || r.value === '}') {
          stack.pop();
        } else if (r.value === ',' && stack.length) {
          const top = stack[stack.length - 1];
          if (++top.args > FORMULA_LIMITS.args && top.func) return invalid();
        }
      }
      if (r.type !== 'whitespace' && r.type !== 'newline') prev = r.type;
    }
  } catch {
    return invalid();
  }
  const file = fxEmit(body, tokens, (t) => {
    if (t.type !== 'func') return null;
    const name = text(t);
    if (/^_xl/i.test(name)) return null; // typed with its prefix: kept
    const upper = name.toUpperCase();
    return (FILE_FUNCTION_PREFIX.get(upper) ?? FILE_FUNCTION_PREFIX_WS.get(upper) ?? '') + upper;
  });
  return { ok: true, file, editor: file.includes('_xl') ? file.replace(/_xl(?:fn|ws|pm)\./g, '') : file };
}

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
  /** The bytes passed to loadWorkbook (see loadedBytes). */
  bytes?: Uint8Array;
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
 * Numbers in unit formats with a literal 'm' stay numbers and <b val="0"/> is
 * read as not bold (patchDisplayFixes); loadedBytes(workbook) is `data` itself.
 */
export async function loadWorkbook(data: Uint8Array, options?: LoadOptions): Promise<ExcelJS.Workbook> {
  const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  assertZipContainer(buffer);
  patchValidationParser();
  patchSheetNameCheck();
  patchCommentParser();
  patchCellParser();
  patchDisplayFixes();
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
  workbookExtras.set(workbook, { ...buildExtras(styles, captured.styles), externalBooks, arrayFormulas, threadedNotes, bytes: data });
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
  const blocks = (section: string): string[] => [...xmlElements(section, 'bk')].map((m) => m.inner ?? '');
  const future = [...xmlElements(xml, 'futureMetadata')].find((m) => parseXmlAttributes(m.attrs).name === 'XLDAPR')?.inner;
  const dynamic = future !== undefined ? blocks(future).map((bk) => /\bfDynamic\s*=\s*["'](?:1|true)["']/.test(bk)) : [];
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
    for (const m of xmlElements(xml ?? '', 'person')) {
      const { id, displayName } = parseXmlAttributes(m.attrs);
      if (id && displayName) persons.set(id, displayName);
    }
  }
  const workbookRels = await readRelationships(zip, 'xl/workbook.xml');
  const out = new Map<number, Map<number, string>>();
  for (const m of xmlElements(xmlSection(workbookXml, 'sheets') ?? '', 'sheet')) {
    const sheetId = parseInt(parseXmlAttributes(m.attrs).sheetId ?? '', 10);
    const target = workbookRels.get(relationshipId(m.attrs) ?? '');
    if (!Number.isFinite(sheetId) || !target) continue;
    const sheetPath = resolvePartPath('xl/workbook.xml', target);
    const threads = new Map<number, { root: boolean; line: string }[]>();
    for (const part of (await readRelationships(zip, sheetPath, THREADED_COMMENT_REL_TYPE)).values()) {
      const xml = await zip.file(resolvePartPath(sheetPath, part))?.async('string');
      for (const c of xmlElements(xml ?? '', 'threadedComment')) {
        const attrs = parseXmlAttributes(c.attrs);
        const at = decodeCell(attrs.ref ?? '');
        if (!at) continue;
        // line breaks normalized as an XML parser does (Excel writes CR LF inside the text)
        const raw = xmlSection(c.inner ?? '', 'text') ?? '';
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
  for (const element of xmlElements(xml, tag)) {
    if (element.inner !== undefined) return element.inner;
  }
  return undefined;
}

/**
 * The <tag ...>...</tag> and <tag .../> elements (any namespace prefix) in file order: the text between the tag name
 * and '>' (attributes), and the inner XML (undefined when self-closing). Parts come from the file, so this scans
 * forward in linear time (a lazy regex to the end tag is quadratic on many start tags without one) and stops at the
 * first start tag without its '>' or end tag, as an XML parser stops at malformed XML.
 */
function* xmlElements(xml: string, tag: string): Generator<{ attrs: string; inner: string | undefined }> {
  const start = new RegExp(`<(?:\\w+:)?${tag}(?=[\\s/>])`, 'g');
  const end = new RegExp(`</(?:\\w+:)?${tag}\\s*>`, 'g');
  for (let m = start.exec(xml); m; m = start.exec(xml)) {
    const close = xml.indexOf('>', start.lastIndex);
    if (close < 0) return;
    if (xml[close - 1] === '/') {
      yield { attrs: xml.slice(start.lastIndex, close - 1), inner: undefined };
      start.lastIndex = close + 1;
      continue;
    }
    end.lastIndex = close + 1;
    const e = end.exec(xml);
    if (!e) return;
    yield { attrs: xml.slice(start.lastIndex, close), inner: xml.slice(close + 1, e.index) };
    start.lastIndex = e.index + e[0].length;
  }
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

// ----- display fixes (D-055) -----
// Two ExcelJS 4.4 misreads that matter once a cell can be edited (the editor starts from the shown value):
//  - a number whose custom format has a literal 'm' ('#,##0.00\ \m', '0\ \k\m') became a JS Date: ExcelJS strips the
//    backslashes (style/numfmt-xform.js) before utils.isDateFmt looks for date letters (sheet/cell-xform.js
//    reconcile). The serial lost precision, and values past the year 275760 became '#VALUE!'.
//  - <b val="0"/>, <i val="false"/>, <strike val="0"/> (any ST_OnOff "off" value) were read as set: BooleanXform
//    (simple/boolean-xform.js) ignores the attribute.
// The parsers are wrapped once, for good, like the other load patches.

/** ExcelJS' parser of <b>, <i>, <strike> ... (an internal module; see patchDisplayFixes). */
interface XBooleanParser {
  tag?: string;
  attr?: string;
  model?: unknown;
  parseOpen(node: { name: string; attributes?: Record<string, string | undefined> }): unknown;
}
const BooleanXform = require('exceljs/lib/xlsx/xform/simple/boolean-xform') as { prototype: XBooleanParser } | undefined;
/** ExcelJS' parser of one <numFmt> (an internal module; see patchDisplayFixes). */
interface XNumFmtParser {
  model?: { id?: number; formatCode?: string };
  parseOpen(node: { name: string; attributes?: Record<string, string | undefined> }): boolean;
}
const NumFmtXform = require('exceljs/lib/xlsx/xform/style/numfmt-xform') as { prototype: XNumFmtParser } | undefined;

/** The cell model and options ExcelJS' CellXform#reconcile works on (sheet/cell-xform.js). */
interface XReconcileCell {
  type?: number;
  value?: unknown;
  result?: unknown;
  styleId?: number;
}
interface XStylesForReconcile {
  model?: { styles?: ({ numFmtId?: number } | undefined)[]; numFmts?: ({ id?: number } | undefined)[] };
}
type XReconcile = (this: unknown, model: XReconcileCell, options: { styles?: XStylesForReconcile } | undefined) => void;

/** ST_OnOff "off" values (the attribute is case-sensitive in the schema; read leniently). */
const ON_OFF_FALSE_RE = /^(?:0|false|off)$/i;
/** Each parsed <numFmt> model -> its formatCode as written in the file (escapes intact). */
const rawFormatCodes = new WeakMap<object, string>();
/** Per StylesXform of a load: cellXfs index -> "keep the number" (its format is not a date format for FileStudio). */
const numberStyles = new WeakMap<object, { formatter: NumberFormatter; keep: Map<number, boolean> }>();
let displayFixesPatched = false;

/**
 * Wraps BooleanXform#parseOpen (an "off" val gives false, which FontXform then leaves unset), NumFmtXform#parseOpen
 * (records the file's format code) and CellXform#reconcile (a number or formula result keeps its numeric value when
 * ExcelJS would turn it into a Date but FileStudio's own formatter does not show that format as a date; the display
 * then formats the exact number, see StyleConverter numFmt and NumberFormatter.isDate).
 */
function patchDisplayFixes(): void {
  if (displayFixesPatched) return;
  displayFixesPatched = true;
  const bool = BooleanXform?.prototype;
  const boolOpen = bool?.parseOpen;
  if (bool && typeof boolOpen === 'function') {
    bool.parseOpen = function parseOpenOnOff(this: XBooleanParser, node): unknown {
      const handled = boolOpen.call(this, node);
      if (node.name === this.tag && ON_OFF_FALSE_RE.test(node.attributes?.[this.attr ?? 'val'] ?? '')) this.model = false;
      return handled;
    };
  }
  const numFmt = NumFmtXform?.prototype;
  const numFmtOpen = numFmt?.parseOpen;
  if (numFmt && typeof numFmtOpen === 'function') {
    numFmt.parseOpen = function parseOpenKeepingCode(this: XNumFmtParser, node): boolean {
      const handled = numFmtOpen.call(this, node);
      const code = node.attributes?.formatCode;
      if (handled && this.model && typeof this.model === 'object' && typeof code === 'string') rawFormatCodes.set(this.model, code);
      return handled;
    };
  }
  const cell = CellXform?.prototype as unknown as { reconcile?: XReconcile } | undefined;
  const reconcile = cell?.reconcile;
  if (cell && typeof reconcile === 'function') {
    cell.reconcile = function reconcileKeepingNumbers(this: unknown, model, options): void {
      const type = model?.type;
      const value = model?.value;
      const result = model?.result;
      const keep =
        ((type === VT.Number && typeof value === 'number') || (type === VT.Formula && typeof result === 'number')) &&
        typeof model.styleId === 'number' &&
        styleKeepsNumber(options?.styles, model.styleId);
      reconcile.call(this, model, options);
      if (!keep) return;
      if (type === VT.Number) {
        model.type = VT.Number;
        model.value = value;
      } else {
        model.result = result;
      }
    };
  }
}

/** Whether cells with this cellXfs index keep their number: the xf's custom format is not a date for FileStudio. */
function styleKeepsNumber(styles: XStylesForReconcile | undefined, styleId: number): boolean {
  if (!styles || typeof styles !== 'object') return false;
  let cache = numberStyles.get(styles);
  if (!cache) {
    cache = { formatter: new NumberFormatter(false), keep: new Map() };
    numberStyles.set(styles, cache);
  }
  let keep = cache.keep.get(styleId);
  if (keep === undefined) {
    keep = false;
    const id = styles.model?.styles?.[styleId]?.numFmtId;
    const list = styles.model?.numFmts;
    if (typeof id === 'number' && Array.isArray(list)) {
      // the last <numFmt> with this id wins, as in StylesXform's index
      let code: string | undefined;
      for (const entry of list) {
        if (entry && entry.id === id) code = rawFormatCodes.get(entry) ?? code;
      }
      if (code !== undefined) keep = !cache.formatter.isDate(code);
    }
    cache.keep.set(styleId, keep);
  }
  return keep;
}

/**
 * The bytes a workbook was loaded from: the very object passed to loadWorkbook (never a copy, never changed), so the
 * host keeps exactly what the display model shows as the base of every save (contract §4.1 `base`, D-006).
 */
export function loadedBytes(workbook: ExcelJS.Workbook): Uint8Array | undefined {
  return workbookExtras.get(workbook)?.bytes;
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
  /** (v0.2.0, P2) Stable sheet id ('s' + the <sheet sheetId> attribute); the position in `index` can change, this cannot. */
  id: SheetId;
  /** (v0.2.0, P2) Position -> id maps and the cell overlay of the sheet (EDITING); every read goes through it. */
  view: SheetView;
  /** (v0.2.0, P2) Every array / spill range of the sheet as loaded (dynamic ones too): value edits inside are blocked. */
  arrays: ArrayFormulaInfo[];
  /** (v0.2.0, P2) Lazy lookups of the loaded merges and tables, and the used range of the loaded sheet. */
  mergeIndex?: RangeIndex;
  tables?: LoadedTable[];
  loadedExtent?: { rowCount: number; colCount: number };
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
  /** (v0.2.0, P2) Sheet position by id. */
  private sheetById = new Map<SheetId, number>();
  /**
   * (v0.2.0, P2) Set when the workbook's sheet ids cannot identify its sheets (two sheets with the same <sheet sheetId>, or
   * none): the model then refuses every edit (planEdit) and says why.
   */
  idProblem: string | undefined;
  /**
   * (v0.2.0, P2) Set by the host for facts only the package knows (a cell with `vm` = in-cell picture / linked data type,
   * a pivot range): returns the message of the block, or undefined. Asked for every cell a value edit changes.
   */
  extraCellBlock: ((sheet: SheetId, r: number, c: number) => string | undefined) | undefined;

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
    // Stable ids (EDITING): 's' + the sheetId attribute. A missing or repeated attribute cannot identify a sheet: such a
    // sheet gets a fresh id and the model refuses edits (idProblem) instead of addressing the wrong sheet.
    const ids = assignSheetIds(this.wb.worksheets);
    this.idProblem = ids.problem;
    this.sheetById = new Map(ids.ids.map((id, index) => [id, index]));
    this.sheets = this.wb.worksheets.map((ws, index) => {
      const arrays = (typeof ws.id === 'number' ? this.arrayFormulas?.get(ws.id) : undefined) ?? [];
      return {
        ws,
        index,
        id: ids.ids[index],
        view: newSheetView(ids.ids[index], ws),
        arrays,
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

  // ----- editing support (v0.2.0): the EDITING section (planEdit / applyEdit / invertEdit) reads and changes the model through these -----

  /** Position of the sheet with this id, -1 when there is none. */
  sheetIndexOf(id: SheetId): number {
    return this.sheetById.get(id) ?? -1;
  }

  /** Id of the sheet at a position. */
  sheetIdAt(index: number): SheetId | undefined {
    return this.sheets[index]?.id;
  }

  /** Names of all sheets (hidden ones too), in tab order. */
  sheetNames(): string[] {
    return this.sheets.map((s) => String(s.ws.name));
  }

  get usesDate1904(): boolean {
    return this.date1904;
  }

  /** The view (maps + overlay) of a sheet. */
  viewOf(id: SheetId): SheetView | undefined {
    return this.sheets[this.sheetById.get(id) ?? -1]?.view;
  }

  /** Meta of one sheet as the webview needs it now (after the used range changed). */
  sheetMeta(id: SheetId): SheetMeta | undefined {
    return this.meta.sheets[this.sheetById.get(id) ?? -1];
  }

  /**
   * Recomputes the meta of one sheet (used range, ...) after edits and returns it (undefined: no such sheet). The base
   * workbook never changes, so this is only needed when the view changed something the meta shows.
   */
  refreshSheet(id: SheetId): SheetMeta | undefined {
    const index = this.sheetById.get(id);
    const ctx = index === undefined ? undefined : this.sheets[index];
    if (!ctx || index === undefined) return undefined;
    ctx.lists.clear();
    const meta = this.safeSheetMeta(ctx);
    this.meta.sheets[index] = meta;
    return meta;
  }

  /** The cached lists of data validations read from cells: dropped after any edit (the next read looks again). */
  dropListCaches(): void {
    for (const ctx of this.sheets) ctx.lists.clear();
  }

  /** What planEdit needs to know about the cell at logical (r, c) of a sheet. */
  cellState(id: SheetId, r: number, c: number): CellState | undefined {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    if (!ctx) return undefined;
    const view = ctx.view;
    const rowId = view.contentMap.get(r);
    const colId = view.colMap.get(c);
    const key = cellKey(rowId, colId);
    const over = view.over.get(key);
    const brow = rowId < NEW_ROW_BASE ? ctx.ws._rows[rowId] : undefined;
    const base = colId < NEW_COL_BASE ? brow?._cells[colId] : undefined;
    const style = over?.s ?? this.baseStyleIndex(ctx, base, rowId, colId);
    const st = this.styleTable.styles[style];
    let value: FileValue | undefined;
    let formula: string | undefined;
    let sharedMaster = false;
    if (over) {
      if (over.f !== undefined) formula = over.f;
      if (over.t === 'e') value = typeof over.v === 'string' ? { t: 'e', v: over.v } : undefined;
      else if (over.t === 'z' || over.v === undefined || over.v === null) value = undefined;
      else if (typeof over.v === 'number') value = { t: 'n', v: over.v };
      else if (typeof over.v === 'boolean') value = { t: 'b', v: over.v };
      else value = { t: 's', v: over.v };
    } else if (base) {
      const xv = base._value;
      const m = xv.model;
      switch (xv.type) {
        case VT.Number:
          if (typeof m.value === 'number') value = { t: 'n', v: m.value };
          break;
        case VT.Date: {
          const serial = m.value instanceof Date ? this.dateSerial(m.value) : undefined;
          if (serial !== undefined) value = { t: 'n', v: serial };
          break;
        }
        case VT.String:
        case VT.SharedString:
          if (typeof m.value === 'string') value = { t: 's', v: m.value };
          break;
        case VT.RichText:
          value = { t: 's', v: richTextPlain(m.value) };
          break;
        case VT.Boolean:
          if (typeof m.value === 'boolean') value = { t: 'b', v: m.value };
          break;
        case VT.Error:
          value = { t: 'e', v: String((m.value as { error?: unknown } | undefined)?.error ?? '#VALUE!') };
          break;
        case VT.Hyperlink:
          value = typeof m.text === 'string' ? { t: 's', v: m.text } : typeof m.text === 'number' ? { t: 'n', v: m.text } : undefined;
          break;
        case VT.Formula: {
          value = this.resultValue(m.result);
          if (typeof m.formula === 'string' && m.formula) {
            formula = m.formula.startsWith('=') ? m.formula.slice(1) : m.formula;
            sharedMaster = m.shareType === 'shared';
          } else if (typeof m.sharedFormula === 'string') {
            const master = decodeCell(m.sharedFormula);
            const masterFormula = master ? ctx.ws._rows[master.r]?._cells[master.c]?._value?.model?.formula : undefined;
            if (master && typeof masterFormula === 'string') {
              formula = translateFormula(masterFormula.startsWith('=') ? masterFormula.slice(1) : masterFormula, rowId - master.r, colId - master.c);
            }
          }
          break;
        }
        default:
          break;
      }
    }
    return {
      rowId,
      colId,
      key,
      override: over,
      value,
      formula,
      style,
      numFmt: st.numFmt ?? 'General',
      wrap: !!st.align?.wrap,
      locked: this.cellLocked(ctx, base, rowId, colId),
      sharedMaster: sharedMaster && !over,
      hasHyperlink: !!base && base._value.type === VT.Hyperlink,
    };
  }

  /** A formula's cached result (ExcelJS value) as the file stores it; undefined when there is none. */
  private resultValue(result: unknown): FileValue | undefined {
    if (typeof result === 'number') return Number.isFinite(result) ? { t: 'n', v: result } : undefined;
    if (typeof result === 'string') return { t: 's', v: result };
    if (typeof result === 'boolean') return { t: 'b', v: result };
    if (result instanceof Date) {
      const serial = this.dateSerial(result);
      return serial === undefined ? undefined : { t: 'n', v: serial };
    }
    if (isErrorValue(result)) return { t: 'e', v: String((result as { error: unknown }).error) };
    return undefined;
  }

  /** Visits the cells inside `range` that exist (loaded or in the overlay), row by row, with their logical position. */
  forEachExistingCell(id: SheetId, range: Range, visit: (r: number, c: number) => void): void {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    if (!ctx) return;
    const view = ctx.view;
    const last = Math.min(range.r1, viewRowLimit(view) - 1);
    for (let r = range.r0; r <= last; r++) {
      this.visitRowCells(ctx, view.contentMap.get(r), range.c0, range.c1, (c) => {
        visit(r, c);
        return true;
      });
    }
  }

  /** Is the cell locked (the default) in the file's own style? Only matters on a protected sheet. */
  private cellLocked(ctx: SheetContext, base: XCell | undefined, rowId: number, colId: number): boolean {
    const own = (st: XStyle | undefined): boolean | undefined => (st as { protection?: { locked?: boolean } } | undefined)?.protection?.locked;
    const row = rowId < NEW_ROW_BASE ? ctx.ws._rows[rowId] : undefined;
    const column = colId < NEW_COL_BASE ? ctx.ws._columns?.[colId] : undefined;
    return (own(base?.style) ?? (base ? undefined : own(row?.style) ?? own(column?.style))) !== false;
  }

  /** The array / spill range (legacy or dynamic) of the loaded sheet that holds the base cell (rowId, colId). */
  arrayRangeAt(id: SheetId, rowId: number, colId: number): Range | undefined {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    return ctx?.arrays.find((a) => rowId >= a.range.r0 && rowId <= a.range.r1 && colId >= a.range.c0 && colId <= a.range.c1)?.range;
  }

  /** The merged area (loaded merges) that holds logical (r, c), if any. */
  mergeAt(id: SheetId, r: number, c: number): Range | undefined {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    if (!ctx) return undefined;
    ctx.mergeIndex ??= buildRangeIndex(this.sheetMerges(ctx.ws));
    return findInRangeIndex(ctx.mergeIndex, r, c);
  }

  /** The table whose header or totals row holds logical (r, c), if any. */
  tableCellAt(id: SheetId, r: number, c: number): { name: string; part: 'header' | 'totals' } | undefined {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    if (!ctx) return undefined;
    ctx.tables ??= loadedTables(ctx.ws);
    for (const t of ctx.tables) {
      if (c < t.range.c0 || c > t.range.c1) continue;
      if (t.header && r === t.range.r0) return { name: t.name, part: 'header' };
      if (t.totals && r === t.range.r1) return { name: t.name, part: 'totals' };
    }
    return undefined;
  }

  /** The sheet has "Protect Sheet" on (cells with the default locked state cannot be changed in Excel). */
  sheetProtected(id: SheetId): boolean {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    return !!(ctx?.ws as { sheetProtection?: { sheet?: boolean } | null } | undefined)?.sheetProtection?.sheet;
  }

  /**
   * Is logical row r a data row of the sheet's autofilter that is hidden (filtered out)? Excel's Clear Contents skips
   * such rows, so the user never loses data they cannot see. Rows added by edits are never hidden.
   */
  filterHidesRow(id: SheetId, r: number): boolean {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    if (!ctx) return false;
    const filter = decodeAutoFilter(ctx.ws.autoFilter);
    if (!filter || r <= filter.r0 || r > filter.r1) return false;
    const rowId = ctx.view.contentMap.get(r);
    const row = rowId < NEW_ROW_BASE ? ctx.ws._rows[rowId] : undefined;
    return !!(row as { hidden?: boolean } | undefined)?.hidden;
  }

  /**
   * The other cells of the shared formula whose master is at base cell (rowId, colId): [r, c, formula (file syntax), cached
   * value]. 'large' when the group is too big to separate in one edit. Cells already replaced by the overlay are left out.
   */
  sharedClones(id: SheetId, rowId: number, colId: number): { r: number; c: number; formula: string; cached: FileValue | null }[] | 'large' | undefined {
    const ctx = this.sheets[this.sheetById.get(id) ?? -1];
    const master = ctx?.ws._rows[rowId]?._cells[colId]?._value;
    if (!ctx || !master || master.type !== VT.Formula || master.model.shareType !== 'shared' || typeof master.model.formula !== 'string') return undefined;
    const ref = typeof master.model.ref === 'string' ? decodeRange(master.model.ref) : undefined;
    if (!ref) return undefined;
    if ((ref.r1 - ref.r0 + 1) * (ref.c1 - ref.c0 + 1) > SHARED_GROUP_CAP) return 'large';
    const address = `${indexToCol(colId)}${rowId + 1}`;
    const masterText = master.model.formula.startsWith('=') ? master.model.formula.slice(1) : master.model.formula;
    const view = ctx.view;
    const out: { r: number; c: number; formula: string; cached: FileValue | null }[] = [];
    for (let r = ref.r0; r <= Math.min(ref.r1, ctx.ws._rows.length - 1); r++) {
      const cells = ctx.ws._rows[r]?._cells;
      if (!cells) continue;
      for (let c = ref.c0; c <= Math.min(ref.c1, cells.length - 1); c++) {
        const xv = cells[c]?._value;
        if (!xv || xv.type !== VT.Formula || xv.model.sharedFormula !== address) continue;
        const own = view.over.get(cellKey(r, c));
        if (own && (own.f !== undefined || own.v !== undefined || own.t !== undefined)) continue; // already its own cell
        out.push({ r, c, formula: translateFormula(masterText, r - rowId, c - colId), cached: this.resultValue(xv.model.result) ?? null });
      }
    }
    return out;
  }

  /** Style-table index of `base` with number format / wrap changes applied (display only; the file derives its own xf). */
  deriveStyle(base: number, deltas: readonly StyleDelta[]): number {
    const from = this.styleTable.styles[base] ?? {};
    let next: CellStyle | undefined;
    for (const d of deltas) {
      if (d.numFmt !== undefined && (d.numFmt === 'General' ? from.numFmt !== undefined : d.numFmt !== from.numFmt)) {
        next = { ...(next ?? from) };
        if (d.numFmt === 'General') delete next.numFmt;
        else next.numFmt = d.numFmt;
      }
      if (d.wrap !== undefined && !!(next ?? from).align?.wrap !== d.wrap) {
        next = { ...(next ?? from), align: { ...(next ?? from).align, wrap: d.wrap } };
      }
    }
    return next ? this.styleTable.add(next) : base;
  }

  /** Recomputes how far the overlay reaches and, when that changed the used range of the sheet, its meta. true = meta changed. */
  updateExtent(id: SheetId): boolean {
    const index = this.sheetById.get(id);
    const ctx = index === undefined ? undefined : this.sheets[index];
    const meta = index === undefined ? undefined : this.meta.sheets[index];
    if (!ctx || !meta) return false;
    const extent = viewExtent(ctx.view);
    // The used range grows with edits and goes back when the edit is undone; it never drops below the loaded one.
    const loaded = ctx.loadedExtent ?? (ctx.loadedExtent = this.loadedUsedRange(ctx));
    const rowCount = Math.max(loaded.rowCount, extent.rows);
    const colCount = Math.max(loaded.colCount, extent.cols);
    if (rowCount === meta.rowCount && colCount === meta.colCount) return false;
    this.meta.sheets[index as number] = { ...meta, rowCount, colCount };
    return true;
  }

  private loadedUsedRange(ctx: SheetContext): { rowCount: number; colCount: number } {
    const ws = ctx.ws;
    const props = ws.properties ?? {};
    const rowPx = typeof props.defaultRowHeight === 'number' && props.defaultRowHeight > 0 ? ptToPx(props.defaultRowHeight) : ptToPx(DEFAULT_ROW_HEIGHT_PT);
    const colPx = typeof props.defaultColWidth === 'number' && props.defaultColWidth > 0 ? widthToPx(props.defaultColWidth) : DEFAULT_COL_WIDTH_PX;
    try {
      return this.usedRange(ws, this.sheetMerges(ws), this.sheetImages(ws), decodeAutoFilter(ws.autoFilter), colPx, rowPx);
    } catch {
      return { rowCount: 1, colCount: 1 };
    }
  }

  getRows(sheet: number, start: number, end: number, c0?: number, c1?: number): { rows: RowData[]; styles: [number, CellStyle][] } {
    const ctx = this.sheets[sheet];
    if (!ctx) {
      return { rows: [], styles: [] };
    }
    // (v0.2.0) A sheet with edits reads through its view; an unedited one keeps the loop below (identity fast path).
    if (!viewIsPlain(ctx.view)) return this.getRowsThroughView(ctx, start, end, c0, c1);
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

  // ----- reading through a sheet view (v0.2.0, EDITING) -----

  /** getRows for a sheet with edits: positions go through the view's maps, the overlay replaces and adds cells. */
  private getRowsThroughView(ctx: SheetContext, start: number, end: number, c0?: number, c1?: number): { rows: RowData[]; styles: [number, CellStyle][] } {
    const view = ctx.view;
    const lo = Math.max(0, Math.floor(start) || 0);
    const hi = Math.min(Math.floor(end) || 0, viewRowLimit(view), MAX_ROWS);
    const [colLo, colHi] = columnWindow(c0, c1);
    // Multi-cell legacy array formulas of the (unmoved) base positions reaching into this window.
    const arrays = view.contentMap.identity && ctx.arrayRanges.length ? ctx.arrayRanges.filter((a) => a.range.r0 < hi && a.range.r1 >= lo) : [];
    const rows: RowData[] = [];
    const used = new Set<number>();
    for (let r = lo; r < hi; r++) {
      const rowId = view.contentMap.get(r);
      const cells: CellData[] = [];
      this.visitRowCells(ctx, rowId, colLo, colHi, (c, cell, colId, o) => {
        const data = o ? this.overrideCellData(ctx, cell, rowId, colId, c, o) : cell ? this.safeCellData(cell, rowId, colId, ctx) : undefined;
        if (data) {
          data.c = c;
          if (!o && arrays.length && data.f === undefined) this.markArrayMember(data, r, c, arrays, ctx);
          cells.push(data);
          if (data.s) used.add(data.s);
        }
        return true;
      });
      if (cells.length) rows.push({ r, cells });
    }
    return { rows, styles: this.styleTable.entries(used) };
  }

  /**
   * Visits the cells of one content row that exist (in the base workbook or in the overlay) with logical column in
   * [colLo, colHi], ascending. `visit(c, base cell, column id, override)` returns false to stop.
   */
  private visitRowCells(
    ctx: SheetContext,
    rowId: number,
    colLo: number,
    colHi: number,
    visit: (c: number, cell: XCell | undefined, colId: number, o: CellOverride | undefined) => boolean,
  ): void {
    const view = ctx.view;
    const brow = rowId < NEW_ROW_BASE ? ctx.ws._rows[rowId] : undefined;
    const extra = view.overRows.get(rowId);
    if (!brow && !extra) return;
    const cells = brow?._cells ?? [];
    if (view.colMap.identity) {
      // Logical column = column id: the base cells in order, the overlay-only cells merged in.
      const ex = extra ? [...extra].filter((id) => id >= colLo && id <= colHi && !cells[id]).sort((a, b) => a - b) : [];
      let k = 0;
      const last = Math.min(cells.length - 1, colHi);
      for (let c = colLo; c <= last; c++) {
        while (k < ex.length && ex[k] < c) {
          if (!visit(ex[k], undefined, ex[k], view.over.get(cellKey(rowId, ex[k])))) return;
          k++;
        }
        const cell = cells[c];
        if (!cell) continue; // an overlay-only cell is visited through `ex`
        const o = extra?.has(c) ? view.over.get(cellKey(rowId, c)) : undefined;
        if (!visit(c, cell, c, o)) return;
      }
      for (; k < ex.length; k++) if (!visit(ex[k], undefined, ex[k], view.over.get(cellKey(rowId, ex[k])))) return;
      return;
    }
    // Moved columns: walk the logical columns (P4 may add an inverse map; only edit-model tests get here before that).
    const lim = Math.min(colHi, view.colMap.len - 1);
    for (let c = colLo; c <= lim; c++) {
      const colId = view.colMap.get(c);
      const cell = colId < NEW_COL_BASE ? cells[colId] : undefined;
      const o = extra?.has(colId) ? view.over.get(cellKey(rowId, colId)) : undefined;
      if ((cell || o) && !visit(c, cell, colId, o)) return;
    }
  }

  /** Style index a cell has before any override: its own style, else the row's / column's default for a missing cell. */
  private baseStyleIndex(ctx: SheetContext, cell: XCell | undefined, rowId: number, colId: number): number {
    if (cell) return this.styleTable.indexOf(cell.style);
    const row = rowId < NEW_ROW_BASE ? ctx.ws._rows[rowId] : undefined;
    if (row && !isStyleEmpty(row.style)) return this.styleTable.indexOf(row.style);
    const column = colId < NEW_COL_BASE ? ctx.ws._columns?.[colId] : undefined;
    return column && !isStyleEmpty(column.style) ? this.styleTable.indexOf(column.style) : 0;
  }

  /** A cell the overlay replaces: its content comes from the override, its style / note from the override / the base cell. */
  private overrideCellData(ctx: SheetContext, cell: XCell | undefined, rowId: number, colId: number, c: number, o: CellOverride): CellData | undefined {
    const s = o.s ?? this.baseStyleIndex(ctx, cell, rowId, colId);
    const fmt = this.styleTable.styles[s].numFmt ?? 'General';
    let out: CellData;
    if (o.f !== undefined) {
      out = o.v === undefined || o.v === null ? { c, t: 's', v: '', w: '' } : this.overrideValueCell(c, o, fmt);
      // Overrides keep the FILE syntax (_xlfn. prefixes); the grid shows the formula without them, as for loaded cells.
      out.f = displayFormula(o.f);
      if (o.calc) out.calc = o.calc;
      const link = this.normalizeLink(hyperlinkFormulaTarget(out.f));
      if (link) out.link = link;
    } else if (o.t === 'z' || o.v === undefined || o.v === null) {
      out = { c, t: 'z' };
    } else {
      out = this.overrideValueCell(c, o, fmt);
      // A hyperlink stays on its cell when the text is replaced (the link is a property of the cell).
      const link = cell && cell._value.type === VT.Hyperlink ? this.normalizeLink(cell._value.model.hyperlink) : undefined;
      if (link) out.link = link;
    }
    if (s) out.s = s;
    const note = ctx.threadedNotes?.get(rowId * MAX_COLS + colId) ?? (cell ? noteText(cell) : undefined);
    if (note) out.note = note;
    if (out.t === 'z' && !out.s && !out.note) return undefined;
    return out;
  }

  private overrideValueCell(c: number, o: CellOverride, fmt: string): CellData {
    if (o.t === 'e') return { c, t: 'e', v: String(o.v), w: String(o.v) };
    return this.valueCell(c, o.v, fmt, undefined);
  }

  /** Stats contribution of an overridden cell (same meaning as statValue). */
  private overrideStat(o: CellOverride): StatValue {
    if (o.f !== undefined) {
      if (o.t === 'e') return 'error';
      return typeof o.v === 'number' ? (Number.isFinite(o.v) ? o.v : 'error') : null;
    }
    if (o.t === 'z' || o.v === undefined || o.v === null) return undefined;
    if (o.t === 'e') return 'error';
    return typeof o.v === 'number' ? (Number.isFinite(o.v) ? o.v : 'error') : null;
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
    if (!viewIsPlain(ctx.view)) return this.getStatsThroughView(ctx, ranges, acc);
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

  /** getStats for a sheet with edits (see getRowsThroughView). */
  private getStatsThroughView(ctx: SheetContext, ranges: Range[], acc: StatsAccumulator): SelectionStats {
    const view = ctx.view;
    forEachRangeRow(ranges, viewRowLimit(view), MAX_COLS, (r, c0, c1) => {
      let go = true;
      this.visitRowCells(ctx, view.contentMap.get(r), c0, c1, (_c, cell, _colId, o) => {
        if (!acc.tick()) return (go = false);
        acc.add(o ? this.overrideStat(o) : cell ? this.statValue(cell) : undefined);
        return true;
      });
      return go;
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
    if (!viewIsPlain(ctx.view)) return this.firstNumberFormatThroughView(ctx, ranges);
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

  private firstNumberFormatThroughView(ctx: SheetContext, ranges: readonly Range[]): string | undefined {
    const view = ctx.view;
    const limit = viewRowLimit(view);
    let work = 0;
    for (const g of ranges) {
      if (!g || ![g.r0, g.r1, g.c0, g.c1].every((n) => typeof n === 'number' && Number.isFinite(n))) {
        continue;
      }
      const r1 = Math.min(limit - 1, Math.floor(Math.max(g.r0, g.r1)));
      const c0 = Math.max(0, Math.floor(Math.min(g.c0, g.c1)));
      const c1 = Math.min(MAX_COLS - 1, Math.floor(Math.max(g.c0, g.c1)));
      for (let r = Math.max(0, Math.floor(Math.min(g.r0, g.r1))); r <= r1; r++) {
        if (++work > STATS_CELL_CAP) {
          return undefined;
        }
        const rowId = view.contentMap.get(r);
        let found: string | undefined;
        this.visitRowCells(ctx, rowId, c0, c1, (_c, cell, colId, o) => {
          if (++work > STATS_CELL_CAP) return false;
          if (o) {
            if (this.overrideStat(o) !== undefined && typeof this.overrideStat(o) === 'number') {
              found = this.styleTable.styles[o.s ?? this.baseStyleIndex(ctx, cell, rowId, colId)].numFmt ?? 'General';
              return false;
            }
          } else if (cell && typeof this.statValue(cell) === 'number') {
            found = this.statNumFmt(cell);
            return false;
          }
          return true;
        });
        if (found !== undefined) return found;
        if (work > STATS_CELL_CAP) return undefined;
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
        id: ctx.id,
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
    const used = this.usedRange(ws, merges, images, autoFilter, defaultColWidth, defaultRowHeight);
    ctx.loadedExtent = used;
    // Edited cells beyond the loaded used range extend it (the overlay of EDITING).
    const extent = viewExtent(ctx.view);
    const rowCount = Math.max(used.rowCount, extent.rows);
    const colCount = Math.max(used.colCount, extent.cols);
    const meta: SheetMeta = {
      index: ctx.index,
      id: ctx.id,
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
    if (!viewIsPlain(target.view)) {
      // A source sheet with edits: the cells as the grid shows them now.
      const tv = target.view;
      const limit = Math.min(range.r1 + 1, viewRowLimit(tv));
      for (let r = range.r0; r < limit && values.length < LIST_VALUES_CAP; r++) {
        const rowId = tv.contentMap.get(r);
        this.visitRowCells(target, rowId, range.c0, range.c1, (c, cell, colId, o) => {
          const data = o ? this.overrideCellData(target, cell, rowId, colId, c, o) : cell ? this.safeCellData(cell, rowId, colId, target) : undefined;
          const display = data && data.t !== 'z' ? data.w ?? String(data.v ?? '') : '';
          if (display !== '') values.push(display);
          return values.length < LIST_VALUES_CAP;
        });
      }
      return values;
    }
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

// ===== EDITING =====
// (v0.2.0, P2) Cell editing of xlsx workbooks in the extension host: the model, not the file. An edit is planned
// (planEdit: checks, one Edit with the file operations it needs), applied to the display model (applyEdit, also redo)
// and taken back (invertEdit). The ExcelJS model is never changed: every sheet has a VIEW over it - position -> id maps
// and a cell overlay - so undo restores exactly what an edit replaced (ids and overlay entries), nothing is copied.
// Sections: ids and index maps, sheet views, cell input (what a typed text means), plan, apply / invert.
// P4 adds structure, sort, styles and sheet operations as sub-sections at the end; P5 recalculation results.
// Not built here: SheetView.objects / .formulas and UndoData.formulas / .objects / .sheets (P4 adds them, additive).

// ----- ids and index maps -----

/** Row ids of rows that were not in the loaded sheet (inserted rows) start here; loaded rows keep their index as id. */
const NEW_ROW_BASE = 2 ** 21;
/** Column ids of inserted columns start here (a sheet has at most 16,384 columns). */
const NEW_COL_BASE = MAX_COLS;
const CELL_KEY_COLS = 2 ** 18;

/** Overlay key of a cell: rowId * 2**18 + colId (below 2**53; for a loaded cell (rowId, colId) is its base position). */
export function cellKey(rowId: number, colId: number): number {
  return rowId * CELL_KEY_COLS + colId;
}

/** A run of consecutive ids ({s, n}: s, s+1, ...) or an explicit list. */
export type IndexPiece = { s: number; n: number } | { ids: Int32Array };

const pieceLen = (pc: IndexPiece): number => ('ids' in pc ? pc.ids.length : pc.n);

/**
 * Logical position -> stable id, as pieces (identity runs or explicit id lists): moving rows costs a few pieces, not a
 * million numbers. splice returns the removed pieces: that is the whole undo data of an insert / delete / sort.
 */
export class IndexMap {
  pieces: IndexPiece[];
  /** Number of positions. */
  len: number;
  private pre: Float64Array | null = null;

  constructor(n: number) {
    this.pieces = n > 0 ? [{ s: 0, n }] : [];
    this.len = n;
  }

  /** Position p holds id p everywhere (nothing moved): readers use their plain loops. */
  get identity(): boolean {
    const pc = this.pieces.length === 1 ? this.pieces[0] : undefined;
    return !!pc && !('ids' in pc) && pc.s === 0 && pc.n === this.len;
  }

  private prefix(): Float64Array {
    if (this.pre) return this.pre;
    const pre = new Float64Array(this.pieces.length + 1);
    for (let i = 0; i < this.pieces.length; i++) pre[i + 1] = pre[i] + pieceLen(this.pieces[i]);
    return (this.pre = pre);
  }

  /** Id at position p. */
  get(p: number): number {
    if (!Number.isInteger(p) || p < 0 || p >= this.len) throw new RangeError(`IndexMap: position ${p} outside 0..${this.len - 1}`);
    if (this.pieces.length === 1) {
      const only = this.pieces[0];
      return 'ids' in only ? only.ids[p] : only.s + p;
    }
    const pre = this.prefix();
    let lo = 0;
    let hi = this.pieces.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (pre[mid] <= p) lo = mid;
      else hi = mid - 1;
    }
    const pc = this.pieces[lo];
    const off = p - pre[lo];
    return 'ids' in pc ? pc.ids[off] : pc.s + off;
  }

  /** Splits so that a piece boundary exists at p; returns the index of the piece starting at p. */
  private cut(p: number): number {
    const pre = this.prefix();
    let i = 0;
    while (i < this.pieces.length && pre[i + 1] <= p) i++;
    if (i >= this.pieces.length || pre[i] === p) return i;
    const pc = this.pieces[i];
    const off = p - pre[i];
    const a: IndexPiece = 'ids' in pc ? { ids: pc.ids.subarray(0, off) } : { s: pc.s, n: off };
    const b: IndexPiece = 'ids' in pc ? { ids: pc.ids.subarray(off) } : { s: pc.s + off, n: pc.n - off };
    this.pieces.splice(i, 1, a, b);
    this.pre = null;
    return i + 1;
  }

  /** Removes [p, p + count) and puts `insert` there. Returns the removed pieces (the undo data). */
  splice(p: number, count: number, insert: readonly IndexPiece[]): IndexPiece[] {
    if (!Number.isInteger(p) || !Number.isInteger(count) || p < 0 || count < 0 || p + count > this.len) throw new RangeError(`IndexMap: splice(${p}, ${count}) outside 0..${this.len}`);
    const i = this.cut(p);
    const j = this.cut(p + count);
    const removed = this.pieces.splice(i, j - i, ...insert);
    this.len += insert.reduce((n, pc) => n + pieceLen(pc), 0) - count;
    this.pre = null;
    this.coalesce(Math.max(0, i - 1), i + insert.length + 1);
    return removed;
  }

  /** Joins neighbouring identity runs ({s, n} + {s + n, m}) in [from, to): undoing an insert gives back one piece. */
  private coalesce(from: number, to: number): void {
    for (let i = Math.min(to, this.pieces.length - 1); i > from; i--) {
      const a = this.pieces[i - 1];
      const b = this.pieces[i];
      if (!('ids' in a) && !('ids' in b) && a.s + a.n === b.s) {
        this.pieces.splice(i - 1, 2, { s: a.s, n: a.n + b.n });
        this.pre = null;
      } else if (pieceLen(b) === 0) {
        this.pieces.splice(i, 1);
        this.pre = null;
      }
    }
  }

  /** Explicit ids of positions [p0, p1]. */
  ids(p0: number, p1: number): Int32Array {
    const out = new Int32Array(p1 - p0 + 1);
    for (let p = p0; p <= p1; p++) out[p - p0] = this.get(p);
    return out;
  }
}

// ----- sheet views -----

/** What the overlay says about one cell: replaces the loaded cell's content (and style, when `s` is set). */
export interface CellOverride {
  /** Value (a number, text, boolean; for t 'e' the error text). Absent with `f`: the formula's result is not known. */
  v?: string | number | boolean | null;
  /** 'z' = cleared (contents gone, style kept); 'e' = error value; 'n' | 's' | 'b' = the type of `v`. */
  t?: CellType;
  /** Formula in FILE syntax (_xlfn. prefixes), without '='. */
  f?: string;
  /** Style-table index when the edit changed the style (a number format preset); absent = the cell's own style. */
  s?: number;
  /** Recalc state of a formula cell (P5): until a result is calculated a new formula is 'stale'. */
  calc?: 'stale' | 'unsupported' | 'circular';
}

/** One sheet of a WorkbookModel: ids of what is where, and the cells the user changed. */
export interface SheetView {
  id: SheetId;
  ws: XWorksheet;
  /** Row position -> row id (row properties follow this). */
  rowMap: IndexMap;
  /** Row position -> content row id (the cells of a row follow this one; a sort changes only this). */
  contentMap: IndexMap;
  /** Column position -> column id. */
  colMap: IndexMap;
  /** Cells changed by edits, by cellKey(rowId, colId). Entries are replaced, never changed in place. */
  over: Map<number, CellOverride>;
  /** rowId -> column ids with an override (derived from `over`; setOverride keeps it in step). */
  overRows: Map<number, Set<number>>;
  /** Largest row id ever in `over`: how far reads must look beyond the loaded rows. */
  maxOverRowId: number;
}

function newSheetView(id: SheetId, ws: XWorksheet): SheetView {
  return { id, ws, rowMap: new IndexMap(MAX_ROWS), contentMap: new IndexMap(MAX_ROWS), colMap: new IndexMap(MAX_COLS), over: new Map(), overRows: new Map(), maxOverRowId: -1 };
}

/** Nothing moved, nothing replaced: the loaded loops are exact (identity fast path of getRows / getStats). */
function viewIsPlain(view: SheetView): boolean {
  return view.over.size === 0 && view.rowMap.identity && view.contentMap.identity && view.colMap.identity;
}

/** Rows a reader has to look at: the loaded rows, and further down the rows of overlay cells. */
function viewRowLimit(view: SheetView): number {
  return view.contentMap.identity ? Math.min(MAX_ROWS, Math.max(view.ws._rows.length, view.maxOverRowId + 1)) : MAX_ROWS;
}

/** The only place that changes `over` (and its row index): set an entry, or remove it (undefined). */
function setOverride(view: SheetView, rowId: number, colId: number, o: CellOverride | undefined): void {
  const key = cellKey(rowId, colId);
  if (o === undefined) {
    if (!view.over.delete(key)) return;
    const cols = view.overRows.get(rowId);
    if (cols) {
      cols.delete(colId);
      if (!cols.size) view.overRows.delete(rowId);
    }
    return;
  }
  view.over.set(key, o);
  let cols = view.overRows.get(rowId);
  if (!cols) view.overRows.set(rowId, (cols = new Set()));
  cols.add(colId);
  if (rowId > view.maxOverRowId) view.maxOverRowId = rowId;
}

/** Used range the overlay adds: rows / columns up to the last overlay cell that has content or a style (loaded ids = positions). */
function viewExtent(view: SheetView): { rows: number; cols: number } {
  let rows = 0;
  let cols = 0;
  if (!view.contentMap.identity || !view.colMap.identity) return { rows, cols };
  for (const [key, o] of view.over) {
    const filled = (o.t !== undefined && o.t !== 'z' && o.v !== undefined && o.v !== null) || o.f !== undefined || !!o.s;
    if (!filled) continue;
    const rowId = Math.floor(key / CELL_KEY_COLS);
    const colId = key - rowId * CELL_KEY_COLS;
    if (rowId + 1 > rows) rows = rowId + 1;
    if (colId + 1 > cols) cols = colId + 1;
  }
  return { rows, cols };
}

/**
 * 's' + the sheetId attribute (ExcelJS: worksheet.id). A missing / repeated / invalid id gets a fresh one so every sheet
 * still has a unique id, and `problem` says why edits are refused.
 */
function assignSheetIds(worksheets: readonly XWorksheet[]): { ids: SheetId[]; problem: string | undefined } {
  const seen = new Set<number>();
  let max = 0;
  for (const ws of worksheets) if (typeof ws.id === 'number' && Number.isInteger(ws.id) && ws.id > max) max = ws.id;
  let problem: string | undefined;
  const ids = worksheets.map((ws) => {
    let n = ws.id;
    if (typeof n !== 'number' || !Number.isInteger(n) || n < 1 || seen.has(n)) {
      problem = 'Two sheets of this workbook have the same internal id, so FileStudio cannot tell them apart safely. The workbook is read-only.';
      n = ++max;
    }
    seen.add(n);
    return `s${n}`;
  });
  return { ids, problem };
}

/** Rectangles with a lookup by cell (loaded merges): 256-row buckets, rectangles taller than 4,096 rows on a side list. */
interface RangeIndex {
  buckets: Map<number, Range[]>;
  tall: Range[];
}

function buildRangeIndex(list: readonly Range[]): RangeIndex {
  const idx: RangeIndex = { buckets: new Map(), tall: [] };
  for (const g of list) {
    if (g.r1 - g.r0 > 4096) {
      idx.tall.push(g);
      continue;
    }
    for (let b = g.r0 >> 8; b <= g.r1 >> 8; b++) {
      const bucket = idx.buckets.get(b);
      if (bucket) bucket.push(g);
      else idx.buckets.set(b, [g]);
    }
  }
  return idx;
}

function findInRangeIndex(idx: RangeIndex, r: number, c: number): Range | undefined {
  const inside = (g: Range): boolean => r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1;
  return idx.buckets.get(r >> 8)?.find(inside) ?? idx.tall.find(inside);
}

/** A table of the loaded sheet (ExcelJS keeps its model in `table`). */
interface LoadedTable {
  name: string;
  range: Range;
  header: boolean;
  totals: boolean;
}

function loadedTables(ws: XWorksheet): LoadedTable[] {
  const out: LoadedTable[] = [];
  const tables = (ws as { tables?: Record<string, { table?: { name?: string; tableRef?: string; ref?: string; headerRow?: boolean; totalsRow?: boolean } } | undefined> }).tables;
  for (const name in tables ?? {}) {
    const t = tables?.[name]?.table;
    const range = typeof t?.tableRef === 'string' ? decodeRange(t.tableRef) : typeof t?.ref === 'string' ? decodeRange(t.ref) : undefined;
    // ExcelJS sets headerRow = (headerRowCount === '1'), so a table without the attribute (OOXML default: 1 header row) reads as
    // "no header". The header row is always treated as there (safe: a header-less table only loses edits of its first row).
    if (t && range) out.push({ name: t.name ?? name, range, header: true, totals: t.totalsRow === true });
  }
  return out;
}

/** Plain text of an ExcelJS rich-text value. */
function richTextPlain(value: unknown): string {
  const runs = (value as { richText?: { text?: unknown }[] } | undefined)?.richText;
  return Array.isArray(runs) ? runs.map((r) => (typeof r?.text === 'string' ? r.text : '')).join('') : '';
}

/** Most cells a shared-formula group may span to be separated in one edit (a bigger group is refused). */
const SHARED_GROUP_CAP = 2_000_000;

// ----- types of edits -----

/** What planEdit knows about one cell (positions are logical; ids are what the overlay uses). */
export interface CellState {
  rowId: number;
  colId: number;
  key: number;
  override: CellOverride | undefined;
  /** Effective content (undefined = empty); a formula cell: its cached result. */
  value: FileValue | undefined;
  /** Effective formula in FILE syntax (a shared formula's clones expanded), undefined = no formula. */
  formula: string | undefined;
  /** Effective style-table index and its number format ('General' when none) / wrap flag. */
  style: number;
  numFmt: string;
  wrap: boolean;
  /** The cell is locked (matters on a protected sheet). */
  locked: boolean;
  /** The cell is the master of a shared formula that is still shared (no overlay entry yet). */
  sharedMaster: boolean;
  hasHyperlink: boolean;
}

/** One user action = one Edit = one journal entry = one VS Code undo step (contract §2.2). */
export interface Edit {
  /** The state id this edit produces; the document assigns it when it commits (planEdit leaves 0). */
  id: number;
  kind: EditKind;
  sheetId: SheetId | null;
  /** The cells the edit touches (logical positions at edit time). */
  range: Range | null;
  /** VS Code undo label, e.g. "Edit B4". */
  label: string;
  /** Positions moved (rows / columns / sheets / sort) or merge / unmerge: bumps structureVersion. false for cell edits. */
  structural: boolean;
  /**
   * What the file needs, positions at edit time. Convention for `cells` edits: operations before the last one are
   * file-only fixes (separating a shared formula); the LAST operation holds the cells the user changed, and the display
   * (applyEdit) is derived from it.
   */
  after: { ops: PackageOp[]; newIds?: { rows?: [number, number]; cols?: [number, number]; sheet?: number } };
  /** Display-only restore data, filled by the first applyEdit; never written to the file. */
  before?: UndoData;
}

/** Only what an edit changed (compact). P4 adds formulas / objects / sheets. */
export interface UndoData {
  /** Overlay entries before the edit (undefined = there was none). */
  cells?: [key: number, prev: CellOverride | undefined][];
  /** Index map pieces an edit replaced (structure, sort: P4). */
  maps?: { axis: 'rows' | 'content' | 'cols'; at: number; removed: IndexPiece[]; insertedCount: number }[];
}

/** The rows whose display changed: logical rows [start, end) of the sheet; `meta`: the sheet's meta changed (used range). */
export interface RowsChange {
  sheetId: SheetId | null;
  rows?: [start: number, end: number];
  structural: boolean;
  meta: boolean;
}

/** What planEdit answers: an Edit, or why there is none. */
export type PlanResult = Edit | { blocked: BlockReason[] } | { invalid: string } | { noop: true };

export const EDIT_TOO_LARGE_TEXT = 'This change is too large (max 1,000,000 cells).';
export const EDIT_NOT_AVAILABLE_TEXT = 'This action is not available yet.';
const EDIT_UNKNOWN_SHEET_TEXT = 'This sheet is no longer in the workbook.';
const EDIT_BAD_REQUEST_TEXT = 'This change is not valid.';
const EDIT_BAD_TEXT = 'The text contains a character Excel cannot store.';
const MAX_EDIT_CELLS = 1_000_000;
const MAX_CELL_TEXT = 32_767;
const MAX_EDIT_CHARS = 16 * 1024 * 1024;
const MAX_CLEAR_RANGES = 1000;

// ----- cell input -----
// What a typed text means (en-US, contract §5.2). Anything not recognised stays text: a text is never changed into a
// number by a guess, and a number with more than 15 significant digits stays text (Excel would round it silently).

/** Result of parseCellInput. */
export type CellInput =
  | { kind: 'empty' }
  /** `quotePrefix`: typed with a leading apostrophe (stored text, the apostrophe is not part of it). */
  | { kind: 'text'; text: string; quotePrefix: boolean }
  /** `numFmt`: preset applied when the cell is General (percent, currency, grouping, date, time). */
  | { kind: 'number'; value: number; numFmt?: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }
  /** '=' + more: `body` is what follows the '='; checkFormulaInput decides whether it is acceptable. */
  | { kind: 'formula'; body: string };

const ERROR_LITERALS: ReadonlySet<string> = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);
const NUMBER_RE = /^(\()?([+-])?(\$)?([+-])?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)([eE][+-]?\d+)?(%)?(\))?$/;
const ISO_DATE_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const US_DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/;
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2})(\.\d{1,3})?)?(?:\s?([AaPp][Mm]))?$/;

/** Does the number format show a percent (a % outside quotes, escapes and [..] codes, first section)? */
export function isPercentFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  const first = fmt.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '').split(';')[0];
  return first.includes('%');
}

/**
 * Cell text -> what to store. `cell.numFmt` is the number format of the cell ('General' / undefined = General, '@' = Text).
 * `cell.autoPercent`: typed into a cell (not pasted); Excel's automatic percent entry then reads a plain number of 1 or
 * more in a percent-formatted cell as a percent (5 -> 5%, 0.25 stays 25%).
 * Pure; no knowledge of the workbook beyond that and the date system.
 */
export function parseCellInput(text: string, cell: { numFmt?: string; date1904?: boolean; autoPercent?: boolean } = {}): CellInput {
  if (text === '') return { kind: 'empty' };
  const fmt = cell.numFmt ?? 'General';
  if (fmt === '@') return { kind: 'text', text, quotePrefix: false };
  if (text.startsWith("'")) return text.length === 1 ? { kind: 'empty' } : { kind: 'text', text: text.slice(1), quotePrefix: true };
  if (text.startsWith('=') && text.length > 1) return { kind: 'formula', body: text.slice(1) };
  const t = text.trim();
  if (t === '') return { kind: 'text', text, quotePrefix: false };
  if (/^(true|false)$/i.test(t)) return { kind: 'boolean', value: t.toLowerCase() === 'true' };
  const upper = t.toUpperCase();
  if (upper.startsWith('#') && ERROR_LITERALS.has(upper)) return { kind: 'error', value: upper };
  const num = parseNumberText(t);
  if (num && cell.autoPercent && isPercentFormat(fmt) && (!num.numFmt || num.numFmt.startsWith('#,##0')) && Math.abs(num.value) >= 1) {
    const scaled = Number(`${num.value}e-2`);
    return { kind: 'number', value: Number.isFinite(scaled) ? scaled : num.value / 100 };
  }
  if (num) return { kind: 'number', value: num.value, ...(num.numFmt ? { numFmt: num.numFmt } : {}) };
  const when = parseDateTimeText(t, !!cell.date1904);
  if (when) return { kind: 'number', value: when.value, numFmt: when.numFmt };
  return { kind: 'text', text, quotePrefix: false };
}

function parseNumberText(t: string): { value: number; numFmt?: string } | undefined {
  const m = NUMBER_RE.exec(t);
  if (!m) return undefined;
  const [, open, sign1, currency, sign2, body, exp, percent, close] = m;
  if (!!open !== !!close || (sign1 && sign2) || (open && (sign1 || sign2))) return undefined;
  if (currency && exp) return undefined;
  const grouped = body.includes(',');
  if (grouped && exp) return undefined;
  const mantissa = body.replace(/,/g, '');
  // Excel keeps 15 significant digits and writes zeros for the rest: such a number stays text here, so no digit is lost.
  if (mantissa.replace('.', '').replace(/^0+/, '').replace(/0+$/, '').length > 15) return undefined;
  const power = exp ? parseInt(exp.slice(1), 10) : 0;
  let value = Number(`${mantissa}e${power - (percent ? 2 : 0)}`);
  if (!Number.isFinite(value)) return undefined;
  if (open || sign1 === '-' || sign2 === '-') value = -value;
  if (value === 0) value = 0; // no negative zero
  const decimals = body.includes('.') ? body.length - body.indexOf('.') - 1 : 0;
  let numFmt: string | undefined;
  if (percent) numFmt = decimals ? '0.00%' : '0%';
  else if (currency) numFmt = decimals ? '"$"#,##0.00' : '"$"#,##0';
  else if (grouped) numFmt = decimals ? '#,##0.00' : '#,##0';
  return { value, ...(numFmt ? { numFmt } : {}) };
}

/** Days from 1899-12-30 for a valid calendar date; undefined when the date does not exist. 1900-02-29 is Excel's fake day 60. */
function excelDateSerial(y: number, mo: number, d: number, date1904: boolean): number | undefined {
  if (y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1 || d > 31) return undefined;
  if (y === 1900 && mo === 2 && d === 29) return date1904 ? undefined : 60;
  const ms = Date.UTC(2000, mo - 1, d);
  const date = new Date(ms);
  date.setUTCFullYear(y);
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== mo - 1 || date.getUTCDate() !== d) return undefined;
  const days = Math.round(date.getTime() / MS_PER_DAY) + 25569; // 1970-01-01 = 25569
  const serial = days < 61 ? days - 1 : days; // before 1900-03-01 Excel's 1900 system skips the fake Feb 29
  if (date1904) return serial - DATE1904_OFFSET >= 0 ? serial - DATE1904_OFFSET : undefined;
  return serial;
}

function parseDateTimeText(t: string, date1904: boolean): { value: number; numFmt: string } | undefined {
  let datePart = t;
  let timePart: string | undefined;
  const split = /^(\S+?)(?:[ T](\d{1,2}:\d{2}.*))?$/.exec(t);
  if (split) {
    datePart = split[1];
    timePart = split[2];
  }
  let serial: number | undefined;
  if (timePart === undefined && TIME_RE.test(t)) {
    const time = parseTimeText(t);
    return time ? { value: time.fraction, numFmt: time.numFmt } : undefined;
  }
  const iso = ISO_DATE_RE.exec(datePart);
  const us = iso ? null : US_DATE_RE.exec(datePart);
  if (iso) serial = excelDateSerial(+iso[1], +iso[2], +iso[3], date1904);
  else if (us) {
    let y = +us[3];
    if (us[3].length === 2) y += y < 30 ? 2000 : 1900;
    serial = excelDateSerial(y, +us[1], +us[2], date1904);
  }
  if (serial === undefined) return undefined;
  if (timePart === undefined) return { value: serial, numFmt: BUILTIN_NUMFMTS[14] };
  const time = parseTimeText(timePart);
  return time ? { value: serial + time.fraction, numFmt: BUILTIN_NUMFMTS[22] } : undefined;
}

function parseTimeText(t: string): { fraction: number; numFmt: string } | undefined {
  const m = TIME_RE.exec(t);
  if (!m) return undefined;
  let hour = +m[1];
  const minute = +m[2];
  const second = m[3] === undefined ? 0 : +m[3] + (m[4] ? Number(m[4]) : 0);
  const ampm = m[5]?.toUpperCase();
  if (ampm) {
    if (hour < 1 || hour > 12) return undefined;
    hour = (hour % 12) + (ampm === 'PM' ? 12 : 0);
  }
  if (hour > 23 || minute > 59 || second >= 60) return undefined;
  const numFmt = ampm ? (m[3] === undefined ? BUILTIN_NUMFMTS[18] : BUILTIN_NUMFMTS[19]) : m[3] === undefined ? BUILTIN_NUMFMTS[20] : BUILTIN_NUMFMTS[21];
  return { fraction: (hour * 3600 + minute * 60 + second) / 86400, numFmt };
}

// ----- plan -----

/** Where the blocks of one edit are collected: one BlockReason per code and message, `where` capped at 12 (+ "+N more"). */
class BlockCollector {
  private readonly byKey = new Map<string, { code: BlockReason['code']; message: string; where: string[]; more: number }>();

  add(code: BlockReason['code'], message: string, where: string): void {
    const key = `${code}\n${message}`;
    let entry = this.byKey.get(key);
    if (!entry) this.byKey.set(key, (entry = { code, message, where: [], more: 0 }));
    if (entry.where.length < 12) entry.where.push(where);
    else entry.more++;
  }

  get any(): boolean {
    return this.byKey.size > 0;
  }

  result(): BlockReason[] {
    return [...this.byKey.values()].map((e) => ({ code: e.code, message: e.message, where: e.more ? [...e.where, `+${e.more} more`] : e.where }));
  }
}

const rangeText = (g: Range): string => (g.r0 === g.r1 && g.c0 === g.c1 ? `${indexToCol(g.c0)}${g.r0 + 1}` : `${indexToCol(g.c0)}${g.r0 + 1}:${indexToCol(g.c1)}${g.r1 + 1}`);

/** Plans one edit; changes nothing (except that the style table may learn a derived style). See PlanResult. */
export function planEdit(m: WorkbookModel, i: EditIntent): PlanResult {
  if (i.kind !== 'setCells' && i.kind !== 'clear') return { blocked: [{ code: 'unsupported', message: EDIT_NOT_AVAILABLE_TEXT, where: [] }] };
  if (m.sheetIndexOf(i.sheetId) < 0) return { invalid: EDIT_UNKNOWN_SHEET_TEXT };
  if (m.idProblem) return { blocked: [{ code: 'unsupported', message: m.idProblem, where: [] }] };
  return i.kind === 'setCells' ? planSetCells(m, i) : planClear(m, i);
}

/** The blocks that apply to a value edit of one cell (array, table, protected sheet, host facts), added to `blocks`; true = the cell is blocked. */
function blockCell(m: WorkbookModel, sheetId: SheetId, sheetName: string, protectedSheet: boolean, r: number, c: number, st: CellState, blocks: BlockCollector): boolean {
  let hit = false;
  const where = (): string => `${quoteSheetForRef(sheetName)}!${indexToCol(c)}${r + 1}`;
  const array = m.arrayRangeAt(sheetId, st.rowId, st.colId);
  if (array) {
    blocks.add('array', `This cell is part of an array formula (${rangeText(array)}). Edit it in Excel.`, where());
    hit = true;
  }
  const table = m.tableCellAt(sheetId, r, c);
  if (table) {
    blocks.add('table', `${indexToCol(c)}${r + 1} is part of the ${table.part === 'header' ? 'header row' : 'totals row'} of table “${table.name}”. Change it in Excel.`, where());
    hit = true;
  }
  if (protectedSheet && st.locked) {
    blocks.add('protected', 'The sheet is protected in Excel.', where());
    hit = true;
  }
  const extra = m.extraCellBlock?.(sheetId, r, c);
  if (extra) {
    blocks.add('unsupported', extra, where());
    hit = true;
  }
  return hit;
}

/** File operations that separate a shared formula whose master is about to change: every other cell gets its own <f>. */
function separateSharedFormula(m: WorkbookModel, sheetId: SheetId, st: CellState, r: number, c: number, blocks: BlockCollector, into: Map<number, [number, number, CellPatch]>): void {
  const clones = m.sharedClones(sheetId, st.rowId, st.colId);
  if (clones === undefined) return;
  if (clones === 'large') {
    blocks.add('unsupported', 'This shared formula covers too many cells to change safely here. Edit it in Excel.', `${indexToCol(c)}${r + 1}`);
    return;
  }
  for (const k of clones) {
    // A translation that leaves the grid (#REF!) is not what Excel computes for the clone: refuse instead of writing it.
    if (k.formula.includes('#REF!')) {
      blocks.add('unsupported', 'A formula copied from this one cannot be separated safely here. Edit it in Excel.', `${indexToCol(k.c)}${k.r + 1}`);
      return;
    }
    into.set(cellKey(k.r, k.c), [k.r, k.c, { formula: k.formula, cached: k.cached }]);
  }
}

function sheetNumber(id: SheetId): number {
  return Number(id.slice(1));
}

function boxOf(cells: readonly [number, number, unknown][]): Range {
  let r0 = Infinity;
  let r1 = -1;
  let c0 = Infinity;
  let c1 = -1;
  for (const [r, c] of cells) {
    if (r < r0) r0 = r;
    if (r > r1) r1 = r;
    if (c < c0) c0 = c;
    if (c > c1) c1 = c;
  }
  return { r0, c0, r1, c1 };
}

function sortedCells<T extends [number, number, unknown]>(cells: Iterable<T>): T[] {
  return [...cells].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

function planSetCells(m: WorkbookModel, i: Extract<EditIntent, { kind: 'setCells' }>): PlanResult {
  const sheetId = i.sheetId;
  const cells = i.cells;
  if (!Array.isArray(cells)) return { invalid: EDIT_BAD_REQUEST_TEXT };
  if (cells.length === 0) return { noop: true };
  if (cells.length > MAX_EDIT_CELLS) return { invalid: EDIT_TOO_LARGE_TEXT };
  const multi = cells.length > 1;
  const date1904 = m.usesDate1904;
  const names = m.sheetNames();
  const sheetName = names[m.sheetIndexOf(sheetId)];
  const protectedSheet = m.sheetProtected(sheetId);
  const blocks = new BlockCollector();
  const planned = new Map<number, [number, number, CellPatch]>();
  const separated = new Map<number, [number, number, CellPatch]>();
  const separatedMasters = new Set<number>();
  const formulas = new Map<string, FormulaCheck>();
  let chars = 0;
  for (const entry of multi ? lastWriteWins(cells) : cells) {
    if (!Array.isArray(entry) || entry.length !== 3 || typeof entry[2] !== 'string') return { invalid: EDIT_BAD_REQUEST_TEXT };
    let [r, c] = entry;
    const text = entry[2].includes('\r') ? entry[2].replace(/\r\n?/g, '\n') : entry[2];
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return { invalid: EDIT_BAD_REQUEST_TEXT };
    if (text.length > MAX_CELL_TEXT) return { invalid: `This text is too long for a cell (max ${MAX_CELL_TEXT.toLocaleString('en-US')} characters).` };
    if ((chars += text.length) > MAX_EDIT_CHARS) return { invalid: EDIT_TOO_LARGE_TEXT };
    if (hasLoneSurrogate(text)) return { invalid: EDIT_BAD_TEXT };
    // A merged cell holds its value in the top-left cell: a single edit goes there (Excel); in a block the other cells of the merge are skipped when empty and refuse the edit otherwise.
    const merge = m.mergeAt(sheetId, r, c);
    if (merge && (r !== merge.r0 || c !== merge.c0)) {
      if (!multi) {
        r = merge.r0;
        c = merge.c0;
      } else {
        if (text !== '') blocks.add('merge', `${rangeText(merge)} is a merged cell. Paste a block that matches the merged cells, or unmerge first.`, `${quoteSheetForRef(sheetName)}!${indexToCol(c)}${r + 1}`);
        continue;
      }
    }
    const st = m.cellState(sheetId, r, c);
    if (!st) return { invalid: EDIT_UNKNOWN_SHEET_TEXT };
    const input = parseCellInput(text, { numFmt: st.numFmt, date1904, autoPercent: !multi });
    const patch: CellPatch = {};
    const deltas: StyleDelta[] = [];
    switch (input.kind) {
      case 'empty':
        if (st.value === undefined && st.formula === undefined) continue;
        patch.value = null;
        break;
      case 'text': {
        const wantWrap = input.text.includes('\n') && !st.wrap;
        if (st.formula === undefined && st.value?.t === 's' && st.value.v === input.text && !wantWrap) continue;
        patch.value = { t: 's', v: input.text };
        if (wantWrap) deltas.push({ wrap: true });
        if (input.quotePrefix) deltas.push({ quotePrefix: true });
        break;
      }
      case 'number': {
        const preset = st.numFmt === 'General' && input.numFmt ? input.numFmt : undefined;
        if (st.formula === undefined && st.value?.t === 'n' && st.value.v === input.value && !preset) continue;
        patch.value = { t: 'n', v: input.value };
        if (preset) deltas.push({ numFmt: preset });
        break;
      }
      case 'boolean':
        if (st.formula === undefined && st.value?.t === 'b' && st.value.v === input.value) continue;
        patch.value = { t: 'b', v: input.value };
        break;
      case 'error':
        if (st.formula === undefined && st.value?.t === 'e' && st.value.v === input.value) continue;
        patch.value = { t: 'e', v: input.value };
        break;
      case 'formula': {
        let check = formulas.get(input.body);
        if (!check) formulas.set(input.body, (check = checkFormulaInput(input.body, names)));
        if (!check.ok) {
          if (check.reason === 'invalid') return { invalid: check.message };
          blocks.add(check.block.code, check.block.message, `${quoteSheetForRef(sheetName)}!${indexToCol(c)}${r + 1}`);
          continue;
        }
        if (st.formula === check.file) continue;
        patch.formula = check.file;
        patch.cached = null; // no <v> until the formula is calculated (P5): the cell shows 'stale'
        break;
      }
    }
    blockCell(m, sheetId, sheetName, protectedSheet, r, c, st, blocks);
    if (deltas.length) patch.style = deltas;
    if (st.sharedMaster && !separatedMasters.has(st.key)) {
      separatedMasters.add(st.key);
      separateSharedFormula(m, sheetId, st, r, c, blocks, separated);
    }
    planned.set(st.key, [r, c, patch]);
  }
  if (blocks.any) return { blocked: blocks.result() };
  if (!planned.size) return { noop: true };
  return buildCellsEdit(sheetId, 'setCells', planned, separated, planned.size === 1 ? `Edit ${rangeText(boxOf([...planned.values()]))}` : `Edit ${planned.size.toLocaleString('en-US')} cells`);
}

/** One entry per cell, the last one sent for it (the order of the rest is kept); malformed entries stay for the checks. */
function lastWriteWins<T>(cells: readonly T[]): T[] {
  const seen = new Set<number>();
  const out: T[] = [];
  for (let k = cells.length - 1; k >= 0; k--) {
    const e = cells[k] as unknown;
    if (Array.isArray(e) && Number.isInteger(e[0]) && Number.isInteger(e[1])) {
      const key = (e[0] as number) * MAX_COLS + (e[1] as number);
      if (seen.has(key)) continue;
      seen.add(key);
    }
    out.push(cells[k]);
  }
  return out.reverse();
}

function planClear(m: WorkbookModel, i: Extract<EditIntent, { kind: 'clear' }>): PlanResult {
  const sheetId = i.sheetId;
  if (i.what !== 'contents') return { blocked: [{ code: 'unsupported', message: EDIT_NOT_AVAILABLE_TEXT, where: [] }] };
  if (!Array.isArray(i.ranges)) return { invalid: EDIT_BAD_REQUEST_TEXT };
  if (i.ranges.length > MAX_CLEAR_RANGES) return { invalid: EDIT_TOO_LARGE_TEXT };
  const sheetName = m.sheetNames()[m.sheetIndexOf(sheetId)];
  const protectedSheet = m.sheetProtected(sheetId);
  const blocks = new BlockCollector();
  const planned = new Map<number, [number, number, CellPatch]>();
  const separated = new Map<number, [number, number, CellPatch]>();
  const separatedMasters = new Set<number>();
  for (const g of i.ranges) {
    if (!g || ![g.r0, g.c0, g.r1, g.c1].every((n) => Number.isInteger(n) && n >= 0) || g.r0 > g.r1 || g.c0 > g.c1 || g.r1 >= MAX_ROWS || g.c1 >= MAX_COLS) return { invalid: EDIT_BAD_REQUEST_TEXT };
    m.forEachExistingCell(sheetId, g, (r, c) => {
      if (m.filterHidesRow(sheetId, r)) return; // Excel: rows hidden by a filter keep their data
      const st = m.cellState(sheetId, r, c);
      if (!st || (st.value === undefined && st.formula === undefined) || planned.has(st.key)) return;
      if (blockCell(m, sheetId, sheetName, protectedSheet, r, c, st, blocks)) return;
      if (st.sharedMaster && !separatedMasters.has(st.key)) {
        separatedMasters.add(st.key);
        separateSharedFormula(m, sheetId, st, r, c, blocks, separated);
      }
      planned.set(st.key, [r, c, { value: null }]);
      if (planned.size > MAX_EDIT_CELLS) blocks.add('limit', EDIT_TOO_LARGE_TEXT, `${quoteSheetForRef(sheetName)}!${indexToCol(c)}${r + 1}`);
    });
  }
  if (blocks.any) return { blocked: blocks.result() };
  if (!planned.size) return { noop: true };
  const box = boxOf([...planned.values()]);
  const label = i.ranges.length === 1 ? `Clear ${rangeText(i.ranges[0])}` : 'Clear contents';
  return buildCellsEdit(sheetId, 'clear', planned, separated, planned.size === 1 ? `Clear ${rangeText(box)}` : label);
}

function buildCellsEdit(
  sheetId: SheetId,
  kind: 'setCells' | 'clear',
  planned: Map<number, [number, number, CellPatch]>,
  separated: Map<number, [number, number, CellPatch]>,
  label: string,
): Edit {
  // A cell the user changes needs no separate fix.
  for (const key of planned.keys()) separated.delete(key);
  const sheet = sheetNumber(sheetId);
  const ops: PackageOp[] = [];
  if (separated.size) ops.push({ kind: 'cells', sheet, cells: sortedCells(separated.values()) });
  ops.push({ kind: 'cells', sheet, cells: sortedCells(planned.values()) });
  return { id: 0, kind, sheetId, range: boxOf([...planned.values()]), label, structural: false, after: { ops } };
}

// ----- apply / invert -----

/** The overlay entry a cell patch stands for (the display side of the file operation), given the cell as it is now. */
function overrideOfPatch(m: WorkbookModel, st: CellState, patch: CellPatch): CellOverride {
  const o: CellOverride = {};
  const set = (v: FileValue): void => {
    o.t = v.t === 'n' ? 'n' : v.t === 's' ? 's' : v.t === 'b' ? 'b' : 'e';
    o.v = v.v;
  };
  if (patch.formula !== undefined) {
    o.f = patch.formula;
    if (patch.cached) set(patch.cached);
    else o.calc = 'stale';
  } else if (patch.value === null) {
    o.t = 'z';
  } else if (patch.value) {
    set(patch.value);
  } else {
    throw new Error('applyEdit: a cell patch without a value or formula is not supported yet');
  }
  if (patch.style?.length) {
    const s = m.deriveStyle(st.style, patch.style);
    if (s !== st.style || st.override?.s !== undefined) o.s = s;
  } else if (st.override?.s !== undefined) {
    o.s = st.override.s;
  }
  return o;
}

function cellsOpOf(e: Edit): Extract<PackageOp, { kind: 'cells' }> {
  const op = e.after.ops[e.after.ops.length - 1];
  if (!op || op.kind !== 'cells') throw new Error('applyEdit: the edit has no cell operation');
  return op;
}

function changeOf(m: WorkbookModel, e: Edit): RowsChange {
  const meta = e.sheetId ? m.updateExtent(e.sheetId) : false;
  m.dropListCaches();
  return { sheetId: e.sheetId, rows: e.range ? [e.range.r0, e.range.r1 + 1] : undefined, structural: false, meta };
}

/**
 * Applies an edit to the display model (the first time and for redo). Everything is validated before the first change:
 * a throw leaves the model as it was. The first run records the overlay entries it replaced in `e.before`.
 */
export function applyEdit(m: WorkbookModel, e: Edit): RowsChange {
  if (e.kind !== 'setCells' && e.kind !== 'clear') throw new Error(`applyEdit: unsupported edit kind '${e.kind}'`);
  const view = e.sheetId ? m.viewOf(e.sheetId) : undefined;
  if (!e.sheetId || !view) throw new Error('applyEdit: unknown sheet');
  const op = cellsOpOf(e);
  const writes: [rowId: number, colId: number, o: CellOverride][] = [];
  for (const [r, c, patch] of op.cells) {
    const st = m.cellState(e.sheetId, r, c);
    if (!st) throw new Error('applyEdit: unknown cell');
    writes.push([st.rowId, st.colId, overrideOfPatch(m, st, patch)]);
  }
  const record = e.before === undefined;
  const before: [number, CellOverride | undefined][] = [];
  for (const [rowId, colId, o] of writes) {
    if (record) before.push([cellKey(rowId, colId), view.over.get(cellKey(rowId, colId))]);
    setOverride(view, rowId, colId, o);
  }
  if (record) e.before = { cells: before };
  return changeOf(m, e);
}

/** Takes an applied edit back: the overlay entries it replaced come back exactly (e.before). */
export function invertEdit(m: WorkbookModel, e: Edit): RowsChange {
  const view = e.sheetId ? m.viewOf(e.sheetId) : undefined;
  if (!e.sheetId || !view) throw new Error('invertEdit: unknown sheet');
  const cells = e.before?.cells;
  if (!cells) throw new Error('invertEdit: the edit was never applied');
  for (let k = cells.length - 1; k >= 0; k--) {
    const [key, prev] = cells[k];
    const rowId = Math.floor(key / CELL_KEY_COLS);
    setOverride(view, rowId, key - rowId * CELL_KEY_COLS, prev);
  }
  return changeOf(m, e);
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
/**
 * Numbers in ';'-separated text, which comes from decimal-comma locales (Excel's list separator is ';' there): ','
 * is the decimal point and '.' groups thousands, for the whole file ('1.234,56', '999,00', '1.000' = 1000). A '.'
 * that is not a thousands group ('1.5') is text, as in those locales.
 */
const DECIMAL_COMMA_RE = /^[ \t]*[-+]?(?:\d+(?:,\d*)?|,\d+)(?:[eE][-+]?\d+)?[ \t]*$/;
const DOT_GROUPED_NUMBER_RE = /^[ \t]*[-+]?\d{1,3}(?:\.\d{3})+(?:,\d+)?[ \t]*$/;
const DECIMAL_COMMA_PERCENT_RE = /^[ \t]*[-+]?(?:\d+(?:,\d*)?|,\d+)%[ \t]*$/;
const PERCENT_TEXT_RE = /^[ \t]*[-+]?(?:\d+\.?\d*|\.\d+)%[ \t]*$/;
/**
 * Excel's delimiter directive: a first line 'sep=' + one character (any case, spaces around it and surrounding quotes
 * allowed, also as the whole file), checked against Excel 16. Group 1 = the delimiter; Excel skips spaces after '='
 * ('sep= ;' is ';'), and a space ('sep= ') hides the line but is not used as the delimiter.
 */
const CSV_SEP_LINE_RE = /^[ \t]*"?sep= *([^\r\n])"?[ \t]*(?:\r\n|\n|\r|$)/i;

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
 *
 * `forView` skips what only serializeCsv needs (the quoting policy and the
 * original text of records it does not reproduce): a read-only view of a big
 * file parses faster.
 */
export function parseCsv(text: string, opts?: { delimiter?: string; collapseSpaces?: boolean; forView?: boolean }): CsvModel {
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
  const delimiter = opts?.delimiter || (sep?.[1] !== ' ' && sep?.[1]) || scan.delimiter;
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
  const source = opts?.forView ? undefined : describeCsvSource(model, body, ends);
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
  if (delimiter === ';') {
    // decimal comma, '.' thousands groups (see DECIMAL_COMMA_RE)
    if (DECIMAL_COMMA_RE.test(text) || DOT_GROUPED_NUMBER_RE.test(text)) {
      value = Number(text.trim().replace(/\./g, '').replace(',', '.'));
    } else if (DECIMAL_COMMA_PERCENT_RE.test(text)) {
      value = Number(text.trim().slice(0, -1).replace(',', '.')) / 100;
    }
  } else if (NUMERIC_TEXT_RE.test(text)) {
    value = Number(text.trim());
  } else if (PERCENT_TEXT_RE.test(text)) {
    value = Number(text.trim().slice(0, -1)) / 100;
  } else if (GROUPED_NUMBER_RE.test(text)) {
    value = Number(text.trim().replace(/,/g, ''));
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
    // Blank lines (records with no text) at the end are not part of the used range, like Excel.
    let usedRows = rows.length;
    while (usedRows > 0 && rows[usedRows - 1].every((field) => !field)) usedRows--;
    const sheet: SheetMeta = {
      index: 0,
      name: csvSheetName(fileName),
      state: 'visible',
      rowCount: Math.max(1, usedRows),
      colCount: Math.max(1, colCount),
      defaultColWidth: DEFAULT_COL_WIDTH_PX,
      defaultRowHeight: ptToPx(DEFAULT_ROW_HEIGHT_PT),
      cols: csvColumnWidths(rows, colCount),
      rows: {},
      merges: [],
      frozen: { rows: hasHeader && usedRows > 1 ? 1 : 0, cols: 0 },
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
    // Every column the grid shows (a CSV file may be wider than Excel's 16,384 columns).
    forEachRangeRow(ranges, rows.length, this.meta.sheets[0].colCount, (r, c0, c1) => {
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

/**
 * Column widths from the first rows' content (px, default width when narrower). Calibri 11 digits are about 7.43 px
 * (MDW is 7): half a pixel of margin per character keeps long numbers (IDs) from turning into 1.23457E+13.
 */
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
    const px = Math.min(480, Math.ceil(len * (MDW + 0.5) + 10));
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
  // lossy v2 (D-012): kept byte for byte by the surgical save, but their results go stale when cells change
  oleObjects: 'Embedded objects (OLE)',
  dataTables: 'What-if data tables',
  queryTables: 'Query tables / data connections',
  dataModel: 'Data model (Power Pivot)',
  richValues: 'Pictures in cells / linked data types',
} as const;

const LOSSY_ORDER: string[] = Object.values(LOSSY);

/**
 * Content FileStudio does not show and cannot update (lossy v2, contract §4.7, D-012): a workbook with any of it
 * opens read-only until the user accepts "Edit anyway". Nothing of it is removed on save (the surgical save copies
 * every part it does not edit byte for byte), but charts, pivots, links, data tables ... keep stale results.
 * Reads the raw zip: part names first, then drawings and (streamed) sheet XML. Returns the labels in a fixed order.
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
  // shown as the cell's note when the threads can be read (readThreadedComments)
  if (has(/^xl\/threadedComments\//i) && !(await readThreadedComments(zip).catch(() => undefined))) {
    found.add(LOSSY.threadedComments);
  }
  if (has(/^xl\/(ctrlProps|activeX)\//i)) found.add(LOSSY.controls);
  if (has(/^xl\/embeddings\//i)) found.add(LOSSY.oleObjects);
  if (has(/^xl\/(queryTables\/|connections\.xml$)/i)) found.add(LOSSY.queryTables);
  if (has(/^xl\/model\//i)) found.add(LOSSY.dataModel);
  if (has(/^xl\/richData\//i)) found.add(LOSSY.richValues);
  const sheetFlags = [LOSSY.sparklines, LOSSY.x14Cf, LOSSY.controls, LOSSY.oleObjects, LOSSY.dataTables, LOSSY.richValues];

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
        if (/<(?:\w+:)?oleObjects[\s>]/.test(text)) found.add(LOSSY.oleObjects);
        if (text.includes('dataTable') && /<(?:\w+:)?f\s[^>]*?\bt\s*=\s*["']dataTable["']/.test(text)) found.add(LOSSY.dataTables);
        // vm = value metadata: a picture in a cell or a linked data type (rich value)
        if (text.includes('vm=') && /<(?:\w+:)?c\s[^>]*?\bvm\s*=/.test(text)) found.add(LOSSY.richValues);
        return sheetFlags.every((label) => found.has(label));
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

// ----- lossy v2 texts -----

/** What happens to each kind of lossy content when cells change (the "Edit anyway" details, contract §4.7). */
const LOSSY_GROUPS: { labels: readonly string[]; effect: string }[] = [
  { labels: [LOSSY.charts, LOSSY.sparklines], effect: 'show the old numbers until the file is opened in Excel' },
  { labels: [LOSSY.pivots, LOSSY.slicers], effect: 'show the old numbers until you refresh them in Excel' },
  { labels: [LOSSY.externalLinks], effect: 'linked values stay as they were last saved until Excel updates the links' },
  { labels: [LOSSY.dataTables, LOSSY.queryTables, LOSSY.dataModel], effect: 'results stay old until Excel recalculates or refreshes them' },
  { labels: [LOSSY.richValues], effect: 'kept as they are; the cells that hold them cannot be changed in FileStudio' },
];

/** Banner of a workbook with lossy content (undefined: none). Replaces the 0.1.x "does not display" text. */
export function lossyBannerText(features: readonly string[]): string | undefined {
  if (features.length === 0) return undefined;
  return `Nothing in this file will be removed. FileStudio does not show these objects and cannot update them: ${features.join(', ')}.`;
}

/**
 * The "Edit anyway" question for a workbook with lossy content (EditState.canEditAnyway; the host shows `title` as a
 * modal with `details` as its detail text): what goes stale, per kind of content, and what stays blocked (§4.7, §5.5).
 */
export function lossyEditAnyway(features: readonly string[]): { title: string; details: string[] } {
  const details: string[] = [];
  const grouped = new Set<string>();
  for (const group of LOSSY_GROUPS) {
    const present = group.labels.filter((label) => features.includes(label));
    for (const label of present) grouped.add(label);
    if (present.length) details.push(`${present.join(', ')}: ${group.effect}.`);
  }
  const rest = features.filter((label) => !grouped.has(label));
  if (rest.length) details.push(`${rest.join(', ')}: kept as they are; FileStudio does not update them.`);
  details.push('Inserting or deleting rows and columns, sorting and sheet changes stay blocked for this file.');
  details.push('Nothing in this file will be removed.');
  return { title: 'This workbook has content that FileStudio cannot update. Edit it anyway?', details };
}

/**
 * Size limit for editing xlsx (0.2.0, user decision 2026-10-30). A save in the real extension host stays inside its
 * budget (save < 5 s, longest block < 250 ms) up to about 150,000 cells (measured: 150,000 cells 0.6 s / 157 ms;
 * 200,000: 0.7 s / 217-228 ms; 250,000: 1.0 s / 284 ms; 500,000: block 785 ms). The limit is set lower, to 100,000 cells or
 * 25 MB of unpacked XML in the parts a save rewrites (worksheets, shared strings, styles; see xlsxEditCostBytes), for margin
 * (see p1-save/decisions.md). Pictures and other parts a save copies raw do not count, and a small zip that unpacks to
 * hundreds of MB of text does. Above the limit the workbook is hard read-only; the lazy / streaming save is a 0.2.1 item.
 */
export const XLSX_EDIT_LIMITS = { cells: 100_000, bytes: 25 * 1024 * 1024 } as const;

/** Unpacked size of the parts a save rewrites (worksheets, shared strings, styles). A zip that cannot be read counts as its own size. */
export function xlsxEditCostBytes(base: Uint8Array): number {
  try {
    let n = 0;
    for (const e of parseZip(base).entries) if (/^xl\/(worksheets\/|sharedStrings|styles)/i.test(e.name)) n += e.usize;
    return n;
  } catch {
    return base.length;
  }
}

export const XLSX_TOO_LARGE_TEXT =
  'This file is too large to edit safely in this version. You can view, scroll and copy, but not edit. Editing large files comes in a later version. It is open read-only.';

/** Reason text when the workbook is over XLSX_EDIT_LIMITS, or undefined. `cells` = sum of used rows x columns of all sheets. */
export function xlsxTooLargeReason(meta: WorkbookMeta, bytes: number): string | undefined {
  let cells = 0;
  for (const sheet of meta.sheets) cells += sheet.rowCount * sheet.colCount;
  if (cells <= XLSX_EDIT_LIMITS.cells && bytes <= XLSX_EDIT_LIMITS.bytes) return undefined;
  return XLSX_TOO_LARGE_TEXT;
}

export const XLSX_EDIT_GROWS_TOO_LARGE_TEXT =
  'This change would make the workbook too large to edit safely in this version (more than 100,000 cells in use). It was not applied.';

/**
 * The size rule of xlsxTooLargeReason, checked for a planned edit: undefined when the workbook stays within
 * XLSX_EDIT_LIMITS.cells after the edit (or the edit does not grow the used range), else the reason to refuse it.
 * Without it one paste could grow a small workbook past the limit (slow saves, a file that reopens read-only).
 */
export function editGrowsTooLarge(m: WorkbookModel, plan: Edit): string | undefined {
  const sheet = plan.sheetId === null ? undefined : m.sheetMeta(plan.sheetId);
  if (!sheet || !plan.range || plan.kind === 'clear') return undefined;
  const rows = Math.max(sheet.rowCount, plan.range.r1 + 1);
  const cols = Math.max(sheet.colCount, plan.range.c1 + 1);
  if (rows === sheet.rowCount && cols === sheet.colCount) return undefined;
  let cells = rows * cols;
  for (let i = 0, id = m.sheetIdAt(0); id !== undefined; id = m.sheetIdAt(++i)) {
    const other = id === plan.sheetId ? undefined : m.sheetMeta(id);
    if (other) cells += other.rowCount * other.colCount;
  }
  return cells > XLSX_EDIT_LIMITS.cells ? XLSX_EDIT_GROWS_TOO_LARGE_TEXT : undefined;
}

/** Text for a workbook the save path cannot patch (`reason` from checkPatchable); there is no "Edit anyway". */
export function hardReadOnlyText(reason: string): string {
  if (reason === XLSX_TOO_LARGE_TEXT) return reason;
  return `FileStudio cannot edit this file safely: ${reason}. It is open read-only.`;
}

// ===== XLSX PACKAGE =====
//
// Surgical xlsx save, zip layer (P1, D-006; contract §2.3, §4.2). The bytes a document loaded stay the base of every
// save: parseZip reads the central directory, readEntry unpacks one part (size and CRC checked), and writeZip writes a
// new package in which every unchanged entry is a raw copy of its original bytes (local header, data, data
// descriptor; same order, same zip comment) and only rewritten parts are deflated again (zlib runs on the libuv
// thread pool; CRCs of big buffers are computed in 8 MB slices with yields). Anything that could not be written back
// byte-exactly (ZIP64, encryption, compression other than store/deflate, duplicate names, damage) is a ZipError:
// such a file is hard read-only (§4.7). The imports below belong to the save sections only.

import * as zlib from 'zlib';
import { SaxesParser } from 'saxes';

const ZIP_LOCAL_SIG = 0x04034b50;
const ZIP_CENTRAL_SIG = 0x02014b50;
const ZIP_END_SIG = 0x06054b50;
const ZIP64_LOCATOR_SIG = 0x07064b50;
/** Parts larger than this are deflated with level 1 (about 10x faster; the size Excel writes), others with level 6. */
const ZIP_FAST_DEFLATE_BYTES = 4 << 20;
/** CRCs of big buffers are computed in slices of this size, with a yield to the event loop between slices. */
const ZIP_CRC_SLICE = 8 << 20;
/** DOS date 1980-01-01 for parts a save adds (deterministic output). */
const ZIP_DOS_DATE_1980 = 0x21;

/** Called at every cooperative yield of the save code; it may throw to cancel (for example a cancelled backup). */
export type SaveYieldHook = () => void;

/** The package cannot be written back byte-exactly (→ hard read-only). The message is a short reason in plain English. */
/** A part that declares more than this unpacked is refused (a small file must not allocate gigabytes). */
const ZIP_MAX_PART_BYTES = 512 * 1024 * 1024;

export class ZipError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ZipError';
  }
}

/** One entry of the central directory. Offsets point into ZipPackage.buf. */
export interface ZipEntry {
  /** Part name (UTF-8 when flag bit 11 is set, else read as Latin-1; used for lookups only). */
  name: string;
  /** The name bytes as stored (a rewritten entry keeps them, so non-UTF-8 names stay byte-exact). */
  nameBytes: Buffer;
  flags: number;
  method: number;
  crc: number;
  csize: number;
  usize: number;
  /** Offset of the local header. */
  lho: number;
  /** Offset of the compressed data. */
  dataStart: number;
  /** End of the entry's raw span: the next local header or the central directory (data descriptor included). */
  spanEnd: number;
  /** The whole central-directory record (46 bytes + name + extra field + comment). */
  rec: Buffer;
}

export interface ZipPackage {
  buf: Buffer;
  /** Central-directory order. */
  entries: ZipEntry[];
  /** Local-header (file) order. */
  byOffset: ZipEntry[];
  cdOffset: number;
  /** Bytes before the first local header (normally none). */
  prefixEnd: number;
  comment: Buffer;
  /** Entry by part name, ignoring case (OPC part names are case-insensitive). */
  get(name: string): ZipEntry | undefined;
}

/** A Node Buffer view of the same bytes (no copy). */
function asNodeBuffer(data: Uint8Array): Buffer {
  return Buffer.isBuffer(data) ? data : Buffer.from(data.buffer, data.byteOffset, data.byteLength);
}

/** Short text of any thrown value (for messages and logs). */
function saveErrorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

const saveTick = (): Promise<void> => new Promise((resolve) => setImmediate(resolve));

/** Lets the event loop run once, then calls the hook (which may throw to cancel). */
async function saveYield(onYield?: SaveYieldHook): Promise<void> {
  await saveTick();
  onYield?.();
}

/** Longest stretch of save work between two yields (the budget is a 250 ms block at most, contract §4.8). */
const SAVE_SLICE_MS = 30;

/**
 * A step counter for long loops: `step()` returns a promise to await when the loop should yield — every `every`
 * steps, or earlier when SAVE_SLICE_MS passed since the last yield (the clock is read every 64 steps) — else undefined.
 */
function saveSlicer(onYield: SaveYieldHook | undefined, every: number): () => Promise<void> | undefined {
  let n = 0;
  let last = performance.now();
  const pause = async (): Promise<void> => {
    await saveYield(onYield);
    last = performance.now();
  };
  return (): Promise<void> | undefined => {
    n++;
    if (n % every === 0) return pause();
    if ((n & 63) === 0 && performance.now() - last >= SAVE_SLICE_MS) return pause();
    return undefined;
  };
}

let crcTable: Int32Array | undefined;

/**
 * CRC-32 of `data`, continuing from `prev`. Uses zlib.crc32 when the runtime has it (Node ≥ 20.15 / 22.2), else a
 * table (o.table forces the table, for tests).
 */
export function crc32(data: Uint8Array, prev = 0, o?: { table?: boolean }): number {
  if (!o?.table && typeof zlib.crc32 === 'function') {
    return zlib.crc32(data, prev) >>> 0;
  }
  if (!crcTable) {
    crcTable = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c;
    }
  }
  let c = ~prev;
  for (let i = 0; i < data.length; i++) c = crcTable[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return ~c >>> 0;
}

/** crc32 over a big buffer in 8 MB slices, yielding to the event loop between slices. */
async function crc32Cooperative(data: Uint8Array, prev: number, onYield?: SaveYieldHook): Promise<number> {
  if (data.length <= ZIP_CRC_SLICE) {
    return crc32(data, prev);
  }
  for (let i = 0; i < data.length; i += ZIP_CRC_SLICE) {
    prev = crc32(data.subarray(i, Math.min(data.length, i + ZIP_CRC_SLICE)), prev);
    await saveYield(onYield);
  }
  return prev;
}

/**
 * Reads the central directory of an xlsx package. Throws ZipError for everything a save could not copy back
 * byte-exactly: ZIP64, multi-part zips, encrypted entries, compression other than store (0) / deflate (8),
 * duplicate part names (ignoring case), local headers that do not match the directory, overlapping entries.
 */
export function parseZip(data: Uint8Array): ZipPackage {
  const buf = asNodeBuffer(data);
  if (buf.length < 22) {
    throw new ZipError('it is not a zip package');
  }
  let eocd = -1;
  for (let i = buf.length - 22, min = Math.max(0, buf.length - 22 - 0xffff); i >= min; i--) {
    if (buf.readUInt32LE(i) === ZIP_END_SIG && i + 22 + buf.readUInt16LE(i + 20) === buf.length) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) {
    throw new ZipError('it is not a zip package (no end record)');
  }
  if (eocd >= 20 && buf.readUInt32LE(eocd - 20) === ZIP64_LOCATOR_SIG) {
    throw new ZipError('it is a ZIP64 package');
  }
  const disk = buf.readUInt16LE(eocd + 4);
  const cdDisk = buf.readUInt16LE(eocd + 6);
  const onDisk = buf.readUInt16LE(eocd + 8);
  const total = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  if (total === 0xffff || onDisk === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    throw new ZipError('it is a ZIP64 package');
  }
  if (disk !== 0 || cdDisk !== 0 || onDisk !== total) {
    throw new ZipError('it is a multi-part zip package');
  }
  const cdEnd = cdOffset + cdSize;
  if (cdEnd > eocd) {
    throw new ZipError('its zip directory is damaged');
  }
  const entries: ZipEntry[] = [];
  const byName = new Map<string, ZipEntry>();
  let p = cdOffset;
  for (let i = 0; i < total; i++) {
    if (p + 46 > cdEnd || buf.readUInt32LE(p) !== ZIP_CENTRAL_SIG) {
      throw new ZipError('its zip directory is damaged');
    }
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const crc = buf.readUInt32LE(p + 16);
    const csize = buf.readUInt32LE(p + 20);
    const usize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const lho = buf.readUInt32LE(p + 42);
    const recEnd = p + 46 + nameLen + extraLen + commentLen;
    if (recEnd > cdEnd) {
      throw new ZipError('its zip directory is damaged');
    }
    const nameBytes = buf.subarray(p + 46, p + 46 + nameLen);
    const name = nameBytes.toString(flags & 0x800 ? 'utf8' : 'latin1');
    if (usize > ZIP_MAX_PART_BYTES && usize !== 0xffffffff) {
      throw new ZipError(`the part ${name} is too large to read safely (${Math.round(usize / 1048576)} MB unpacked)`);
    }
    if (flags & 0x41) {
      throw new ZipError(`the part ${name} is encrypted`);
    }
    if (csize === 0xffffffff || usize === 0xffffffff || lho === 0xffffffff) {
      throw new ZipError('it is a ZIP64 package');
    }
    if (method !== 0 && method !== 8) {
      throw new ZipError(`the part ${name} uses zip compression method ${method}`);
    }
    const key = name.toLowerCase();
    if (byName.has(key)) {
      throw new ZipError(`the part ${name} is in the package twice`);
    }
    if (lho + 30 > cdOffset || buf.readUInt32LE(lho) !== ZIP_LOCAL_SIG) {
      throw new ZipError(`the zip header of ${name} is damaged`);
    }
    const localNameLen = buf.readUInt16LE(lho + 26);
    const localExtraLen = buf.readUInt16LE(lho + 28);
    if (localNameLen !== nameLen || buf.compare(nameBytes, 0, nameLen, lho + 30, lho + 30 + nameLen) !== 0) {
      throw new ZipError(`the zip header of ${name} does not match the zip directory`);
    }
    const entry: ZipEntry = {
      name, nameBytes, flags, method, crc, csize, usize, lho,
      dataStart: lho + 30 + localNameLen + localExtraLen,
      spanEnd: cdOffset,
      rec: buf.subarray(p, recEnd),
    };
    entries.push(entry);
    byName.set(key, entry);
    p = recEnd;
  }
  const byOffset = entries.slice().sort((a, b) => a.lho - b.lho);
  for (let i = 0; i < byOffset.length; i++) {
    const entry = byOffset[i];
    entry.spanEnd = i + 1 < byOffset.length ? byOffset[i + 1].lho : cdOffset;
    if (entry.dataStart + entry.csize > entry.spanEnd) {
      throw new ZipError(`the part ${entry.name} overlaps the next part`);
    }
  }
  return {
    buf,
    entries,
    byOffset,
    cdOffset,
    prefixEnd: byOffset.length ? byOffset[0].lho : cdOffset,
    comment: buf.subarray(eocd + 22),
    get: (name: string) => byName.get(name.toLowerCase()),
  };
}

/**
 * Inflates one entry (thread pool) straight into a buffer of its declared size, computing the CRC per output chunk:
 * no big copy on the JS thread (a 200 MB sheet part would block it for a long time with a single concat).
 */
function inflateEntry(raw: Uint8Array, entry: ZipEntry): Promise<{ data: Buffer; crc: number }> {
  return new Promise((resolve, reject) => {
    // Unpacked into one buffer of the size the central directory gives (no 2x peak and no big Buffer.concat on the
    // event loop, as zlib.inflateRaw would do); more or fewer bytes than that size → ZipError.
    let out: Buffer;
    try {
      out = Buffer.allocUnsafe(entry.usize);
    } catch {
      reject(new ZipError(`the part ${entry.name} is too big`));
      return;
    }
    const inflate = zlib.createInflateRaw({ chunkSize: 1 << 16 });
    let filled = 0;
    let crc = 0;
    let done = false;
    const fail = (message: string): void => {
      if (done) return;
      done = true;
      inflate.destroy();
      reject(new ZipError(`the part ${entry.name} ${message}`));
    };
    inflate.on('data', (chunk: Buffer) => {
      if (done) return;
      if (filled + chunk.length > out.length) {
        fail('has the wrong size');
        return;
      }
      chunk.copy(out, filled);
      filled += chunk.length;
      crc = crc32(chunk, crc);
    });
    inflate.on('error', (err: Error) => fail(`cannot be unpacked (${err.message})`));
    inflate.on('end', () => {
      if (done) return;
      if (filled !== out.length) {
        fail('has the wrong size');
        return;
      }
      done = true;
      resolve({ data: out, crc });
    });
    const SLICE = 1 << 20;
    let pos = 0;
    const pump = (): void => {
      while (pos < raw.length && !done) {
        const slice = raw.subarray(pos, Math.min(raw.length, pos + SLICE));
        pos += slice.length;
        if (!inflate.write(slice)) {
          inflate.once('drain', pump);
          return;
        }
      }
      if (!done) inflate.end();
    };
    pump();
  });
}

/**
 * The unpacked bytes of one entry (inflate runs on the thread pool); size and CRC are checked (ZipError on a
 * mismatch). A stored entry is returned as a view of the package bytes: callers never modify what they read.
 */
export async function readEntry(zip: ZipPackage, entry: ZipEntry, onYield?: SaveYieldHook): Promise<Buffer> {
  const raw = zip.buf.subarray(entry.dataStart, entry.dataStart + entry.csize);
  let out: Buffer;
  let crc: number;
  if (entry.method === 0) {
    out = raw;
    if (out.length !== entry.usize) {
      throw new ZipError(`the part ${entry.name} has the wrong size`);
    }
    crc = await crc32Cooperative(out, 0, onYield);
  } else {
    ({ data: out, crc } = await inflateEntry(raw, entry));
    if (out.length > ZIP_CRC_SLICE) await saveYield(onYield);
  }
  if (crc !== entry.crc) {
    throw new ZipError(`the part ${entry.name} is damaged (CRC)`);
  }
  return out;
}

/** Deflates a list of buffers as one part (streaming, thread pool). Level 6 up to 4 MB, level 1 above. */
function deflateParts(parts: readonly Uint8Array[], onYield?: SaveYieldHook): Promise<{ data: Buffer; crc: number; usize: number }> {
  const usize = parts.reduce((n, part) => n + part.length, 0);
  const level = usize > ZIP_FAST_DEFLATE_BYTES ? 1 : 6;
  return new Promise((resolve, reject) => {
    const deflate = zlib.createDeflateRaw({ level, memLevel: 8 });
    const chunks: Buffer[] = [];
    let crc = 0;
    deflate.on('data', (chunk: Buffer) => chunks.push(chunk));
    deflate.on('error', reject);
    deflate.on('end', () => resolve({ data: Buffer.concat(chunks), crc, usize }));
    (async () => {
      for (const part of parts) {
        crc = await crc32Cooperative(part, crc, onYield);
        if (!deflate.write(part)) {
          await new Promise<void>((done) => deflate.once('drain', () => done()));
        }
      }
      deflate.end();
    })().catch((err: unknown) => {
      deflate.destroy();
      reject(err);
    });
  });
}

export interface ZipWriteResult {
  bytes: Buffer;
  /** Rewritten and added part names. */
  changed: string[];
  removed: string[];
  /** Entries copied as raw bytes. */
  identical: number;
}

/**
 * Writes a new package from `zip`. `changes`: part name → new content (a list of buffers, concatenated) or null
 * (remove the part). `o.added`: new parts, appended at the end. Every other entry is copied byte for byte (its local
 * header, data and data descriptor), in the original order; the central directory keeps its order and the zip
 * comment is kept. A rewritten entry keeps its name bytes, time, attributes and comment; it loses its data
 * descriptor and extra field (sizes and CRC are in the headers).
 */
export async function writeZip(
  zip: ZipPackage,
  changes: ReadonlyMap<string, readonly Uint8Array[] | null>,
  o?: { added?: readonly { name: string; parts: readonly Uint8Array[] }[]; onYield?: SaveYieldHook },
): Promise<ZipWriteResult> {
  const out: Buffer[] = [];
  let offset = 0;
  const push = (chunk: Buffer): void => {
    out.push(chunk);
    offset += chunk.length;
  };
  const changeOf = new Map<ZipEntry, readonly Uint8Array[] | null>();
  for (const [name, parts] of changes) {
    const entry = zip.get(name);
    if (!entry) throw new Error(`writeZip: the package has no part ${name}`);
    changeOf.set(entry, parts);
  }
  const writeNew = async (nameBytes: Buffer, parts: readonly Uint8Array[], base: ZipEntry | undefined): Promise<Buffer> => {
    const { data, crc, usize } = await deflateParts(parts, o?.onYield);
    const utf8 = base ? base.flags & 0x800 : /[^\x00-\x7f]/.test(nameBytes.toString('latin1')) ? 0x800 : 0;
    const flags = utf8 | (base ? base.flags & 0x0006 : 0); // keep the deflate-option bits; no data descriptor
    const local = Buffer.alloc(30 + nameBytes.length);
    local.writeUInt32LE(ZIP_LOCAL_SIG, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(base ? base.rec.readUInt16LE(12) : 0, 10);
    local.writeUInt16LE(base ? base.rec.readUInt16LE(14) : ZIP_DOS_DATE_1980, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(usize, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    nameBytes.copy(local, 30);
    const lho = offset;
    push(local);
    push(data);
    const comment = base ? base.rec.subarray(46 + base.rec.readUInt16LE(28) + base.rec.readUInt16LE(30)) : Buffer.alloc(0);
    const rec = Buffer.alloc(46 + nameBytes.length + comment.length);
    rec.writeUInt32LE(ZIP_CENTRAL_SIG, 0);
    rec.writeUInt16LE(base ? base.rec.readUInt16LE(4) : 20, 4); // version made by
    rec.writeUInt16LE(20, 6);
    local.copy(rec, 8, 6, 28); // flags, method, time, date, crc, sizes, name length
    rec.writeUInt16LE(0, 30);
    rec.writeUInt16LE(comment.length, 32);
    rec.writeUInt16LE(0, 34);
    rec.writeUInt16LE(base ? base.rec.readUInt16LE(36) : 0, 36);
    rec.writeUInt32LE(base ? base.rec.readUInt32LE(38) : 0, 38);
    rec.writeUInt32LE(lho, 42);
    nameBytes.copy(rec, 46);
    comment.copy(rec, 46 + nameBytes.length);
    return rec;
  };
  const changed: string[] = [];
  const removed: string[] = [];
  let identical = 0;
  if (zip.prefixEnd > 0) push(zip.buf.subarray(0, zip.prefixEnd));
  const central = new Map<ZipEntry, Buffer>();
  for (const entry of zip.byOffset) {
    const parts = changeOf.get(entry);
    if (parts === null) {
      removed.push(entry.name);
    } else if (parts) {
      central.set(entry, await writeNew(entry.nameBytes, parts, entry));
      changed.push(entry.name);
    } else {
      const lho = offset;
      push(zip.buf.subarray(entry.lho, entry.spanEnd));
      const rec = Buffer.from(entry.rec);
      rec.writeUInt32LE(lho, 42);
      central.set(entry, rec);
      identical++;
    }
    if (offset > 0xffffffff) throw new ZipError('the saved package would need ZIP64');
  }
  const addedRecs: Buffer[] = [];
  const addedNames = new Set<string>();
  for (const add of o?.added ?? []) {
    const key = add.name.toLowerCase();
    if (zip.get(add.name) || addedNames.has(key)) throw new Error(`writeZip: the part ${add.name} exists already`);
    addedNames.add(key);
    addedRecs.push(await writeNew(Buffer.from(add.name, 'utf8'), add.parts, undefined));
    changed.push(add.name);
  }
  const count = central.size + addedRecs.length;
  const cdStart = offset;
  for (const entry of zip.entries) {
    const rec = central.get(entry);
    if (rec) push(rec);
  }
  for (const rec of addedRecs) push(rec);
  const cdSize = offset - cdStart;
  if (count >= 0xffff || offset > 0xffffffff) {
    throw new ZipError('the saved package would need ZIP64');
  }
  const end = Buffer.alloc(22);
  end.writeUInt32LE(ZIP_END_SIG, 0);
  end.writeUInt16LE(count, 8);
  end.writeUInt16LE(count, 10);
  end.writeUInt32LE(cdSize, 12);
  end.writeUInt32LE(cdStart, 16);
  end.writeUInt16LE(zip.comment.length, 20);
  push(end);
  push(zip.comment);
  return { bytes: Buffer.concat(out, offset), changed, removed, identical };
}

// ===== XLSX PATCH =====
//
// Byte-level patchers for the parts an edit touches (P1, D-006; rules in contract §4.3). Every byte outside the edited
// elements stays as it was: sheet XML is walked as a Buffer (an ASCII tag search is safe on UTF-8) and only the
// patched rows are decoded and re-emitted; sharedStrings and styles.xml are append-only (existing items never
// change); workbook.xml only gets calcPr fullCalcOnLoad="1"; calcChain is dropped with its relationship and content
// type. Whatever a patcher cannot do safely throws PatchError: the save fails and nothing is written.

/** A patcher refuses a change (unsafe or not supported). The save fails and nothing is written. */
export class PatchError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PatchError';
  }
}

// ----- xml helpers -----

/** One attribute as written in the file: the value is still XML-escaped. */
interface XmlAttr {
  value: string;
  quote: '"' | "'";
}

/** A child element: local name, its start tag and its whole text. */
interface XmlChild {
  local: string;
  open: string;
  xml: string;
}

interface XmlElement {
  qname: string;
  attrs: Map<string, XmlAttr>;
  children: XmlChild[];
}

const XML_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const XML_ENTITY_RE = /&(?:(lt|gt|amp|quot|apos)|#(\d+)|#x([0-9A-Fa-f]+));/g;
const XML_NAMED_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
/** Characters XML 1.0 does not allow (C0 controls except tab, LF, CR; U+FFFE; U+FFFF) and lone surrogates. */
const XML_INVALID_CHAR_RE = /[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;
const LONE_SURROGATE_RE = /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;
const UTF8_STRICT = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });

function regexEscape(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Attributes of a start tag (`<name a="1" b='2'>`), in file order. */
function parseXmlAttrs(tag: string): Map<string, XmlAttr> {
  const attrs = new Map<string, XmlAttr>();
  const nameEnd = tag.search(/[\s/>]/);
  for (const m of tag.slice(nameEnd < 0 ? tag.length : nameEnd).matchAll(XML_ATTR_RE)) {
    attrs.set(m[1], m[2] !== undefined ? { value: m[2], quote: '"' } : { value: m[3] ?? '', quote: "'" });
  }
  return attrs;
}

/** Index of the ">" that ends the start tag at the beginning of `xml` (a ">" inside a quoted attribute value is skipped). */
function startTagEnd(xml: string): number {
  let quote = '';
  for (let i = 0; i < xml.length; i++) {
    const ch = xml[i];
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return i;
    }
  }
  return -1;
}

function emitXmlAttrs(attrs: ReadonlyMap<string, XmlAttr>): string {
  let out = '';
  for (const [name, attr] of attrs) out += ` ${name}=${attr.quote}${attr.value}${attr.quote}`;
  return out;
}

/** Sets an attribute (`value` already escaped); an existing attribute keeps its place and quote style. */
function setXmlAttr(attrs: Map<string, XmlAttr>, name: string, value: string): void {
  const quote = attrs.get(name)?.quote ?? '"';
  attrs.set(name, { value: quote === "'" ? value.replace(/'/g, '&apos;') : value, quote });
}

function xmlAttrText(attrs: ReadonlyMap<string, XmlAttr>, name: string): string | undefined {
  const attr = attrs.get(name);
  return attr === undefined ? undefined : xmlUnescape(attr.value);
}

function xmlUnescape(text: string): string {
  if (!text.includes('&')) return text;
  return text.replace(XML_ENTITY_RE, (match: string, named?: string, dec?: string, hex?: string): string => {
    if (named) return XML_NAMED_ENTITIES[named];
    const code = dec !== undefined ? parseInt(dec, 10) : parseInt(hex ?? '', 16);
    try {
      return String.fromCodePoint(code);
    } catch {
      return match;
    }
  });
}

/** Text content: `& < >`, and CR as a character reference (a raw CR would be read back as LF). */
function xmlEscapeText(text: string): string {
  return text.replace(/[&<>\r]/g, (ch) => (ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : '&#13;'));
}

/** Attribute value in double quotes (tab, LF, CR as character references: attribute values are normalised). */
function xmlEscapeAttr(text: string): string {
  return text.replace(/[&<>"\t\n\r]/g, (ch) =>
    ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : `&#${ch.charCodeAt(0)};`);
}

/**
 * Cell text as an OOXML ST_Xstring, XML-escaped: a literal `_xHHHH_` is protected as `_x005F_xHHHH_`, control
 * characters (all C0 except tab and LF, so CR too) and U+FFFE / U+FFFF become `_xHHHH_` (upper-case hex, as Excel
 * writes and as decodeXstring reads them back). A lone surrogate is refused (it cannot be stored).
 */
function escapeXstring(text: string, where: string): string {
  if (LONE_SURROGATE_RE.test(text)) {
    throw new PatchError(`${where}: the text contains an invalid character`);
  }
  const escaped = text
    .replace(/_(?=x[0-9A-Fa-f]{4}_)/g, '_x005F_') // lookahead: '_x005F_x0041_' must not lose its shared underscore
    .replace(/[\x00-\x08\x0b-\x1f\ufffe\uffff]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
  return xmlEscapeText(escaped);
}

/** xml:space="preserve" is needed for leading/trailing white space, tabs and line breaks. */
function needsSpacePreserve(text: string): boolean {
  return /^\s|\s$|[\n\t]/.test(text);
}

function isXmlNameEnd(code: number): boolean {
  return code === 0x20 || code === 0x3e || code === 0x2f || code === 0x09 || code === 0x0a || code === 0x0d;
}

/** Root element prefix ('' or 'x:') of a part. */
function xmlRootPrefix(head: string, local: string, where: string): string {
  const m = new RegExp(`<([A-Za-z_][\\w.-]*:)?${local}[\\s>/]`).exec(head);
  if (!m) throw new PatchError(`${where} has no <${local}> element`);
  return m[1] ?? '';
}

/** Throws when a part is not UTF-8 (UTF-16 byte order mark or another declared encoding). */
function assertUtf8Head(data: Uint8Array, where: string): void {
  if (data.length >= 2 && ((data[0] === 0xff && data[1] === 0xfe) || (data[0] === 0xfe && data[1] === 0xff))) {
    throw new PatchError(`${where} is not UTF-8`);
  }
  const head = asNodeBuffer(data).toString('latin1', 0, Math.min(data.length, 256));
  const decl = /^(?:\xef\xbb\xbf)?<\?xml[^>]*\bencoding\s*=\s*["']([^"']+)["']/.exec(head);
  if (decl && !/^utf-?8$/i.test(decl[1])) {
    throw new PatchError(`${where} is not UTF-8 (${decl[1]})`);
  }
}

/** A whole small part as text (strict UTF-8; a byte order mark is kept so it is written back). */
function partText(data: Uint8Array, where: string): string {
  assertUtf8Head(data, where);
  try {
    return UTF8_STRICT.decode(data);
  } catch {
    throw new PatchError(`${where} is not valid UTF-8`);
  }
}

/** The child elements of an element body. Text between children must be white space. */
function splitXmlChildren(inner: string, where: string): XmlChild[] {
  const out: XmlChild[] = [];
  let p = 0;
  while (p < inner.length) {
    const lt = inner.indexOf('<', p);
    if (lt < 0 || inner.slice(p, lt).trim()) {
      if (inner.slice(p, lt < 0 ? inner.length : lt).trim()) throw new PatchError(`${where} has unexpected text`);
      if (lt < 0) break;
    }
    const gt = inner.indexOf('>', lt);
    const name = /^<([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/.exec(inner.slice(lt, gt + 1));
    if (gt < 0 || !name) throw new PatchError(`${where} is damaged`);
    let end = gt + 1;
    if (inner[gt - 1] !== '/') {
      const close = `</${name[1] ?? ''}${name[2]}>`;
      const at = inner.indexOf(close, gt);
      if (at < 0) throw new PatchError(`${where} is damaged`);
      end = at + close.length;
    }
    out.push({ local: name[2], open: inner.slice(lt, gt + 1), xml: inner.slice(lt, end) });
    p = end;
  }
  return out;
}

function parseXmlElement(xml: string, where: string): XmlElement {
  const gt = xml.indexOf('>');
  const open = xml.slice(0, gt + 1);
  const qname = /^<([^\s/>]+)/.exec(open)?.[1];
  if (gt < 0 || !qname) throw new PatchError(`${where} is damaged`);
  const inner = xml[gt - 1] === '/' ? '' : xml.slice(gt + 1, xml.lastIndexOf('</'));
  return { qname, attrs: parseXmlAttrs(open), children: splitXmlChildren(inner, where) };
}

function buildXmlElement(el: XmlElement): string {
  const body = el.children.map((child) => child.xml).join('');
  return body ? `<${el.qname}${emitXmlAttrs(el.attrs)}>${body}</${el.qname}>` : `<${el.qname}${emitXmlAttrs(el.attrs)}/>`;
}

/** Replaces (or removes, xml = null) the child `local`; a new child goes to its place in the schema order. */
function setXmlChild(el: XmlElement, local: string, xml: string | null, order: readonly string[]): void {
  const at = el.children.findIndex((child) => child.local === local);
  el.children = el.children.filter((child) => child.local !== local);
  if (xml === null) return;
  const child: XmlChild = { local, open: xml, xml };
  if (at >= 0) {
    el.children.splice(Math.min(at, el.children.length), 0, child);
    return;
  }
  const rank = order.indexOf(local);
  const before = el.children.findIndex((other) => order.indexOf(other.local) > rank);
  el.children.splice(before < 0 ? el.children.length : before, 0, child);
}

/**
 * Calls `found` for every occurrence of `needle` in buf[from, to), searching 8 MB windows with a yield between them
 * (one Buffer.indexOf over a 200 MB part would block the event loop too long).
 */
async function forEachOccurrence(buf: Buffer, needle: Buffer, from: number, to: number, onYield: SaveYieldHook | undefined, found: (at: number) => void): Promise<void> {
  const WINDOW = 8 << 20;
  for (let pos = from; pos < to; pos += WINDOW) {
    const window = buf.subarray(pos, Math.min(to, pos + WINDOW + needle.length - 1));
    for (let i = window.indexOf(needle); i >= 0 && i < WINDOW; i = window.indexOf(needle, i + 1)) found(pos + i);
    if (to - from > WINDOW) await saveYield(onYield);
  }
}

// ----- package map -----

interface XlsxRel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

interface XlsxSheetRef {
  name: string;
  /** The <sheet sheetId> attribute (SheetId = 's' + sheetId). */
  sheetId: number;
  /** Zip entry name of the sheet part (undefined when the relationship or the part is missing). */
  part?: string;
  /** Relationship type (…/worksheet, …/chartsheet, …). */
  type: string;
}

interface XlsxPackage {
  zip: ZipPackage;
  wbPath: string;
  wbXml: string;
  wbRelsPath: string;
  wbRelsXml: string;
  ctXml: string;
  sheets: XlsxSheetRef[];
  sstPath?: string;
  stylesPath?: string;
  /** calcChain parts: the workbook relationship target and any part with the calcChain content type. */
  calcChainPaths: string[];
  /** Strict Open XML (purl.oclc.org namespaces): not patchable. */
  strict: boolean;
}

const CALC_CHAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';

function relsPartOf(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

/** Zip entry name of a relationship target (relative to `fromPart`, or absolute). */
function resolvePartTarget(fromPart: string, target: string): string {
  const joined = target.startsWith('/') ? target.slice(1) : fromPart.slice(0, fromPart.lastIndexOf('/') + 1) + target;
  const out: string[] = [];
  for (const seg of joined.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return out.join('/');
}

/** The exact entry name for a part name (case-insensitive; percent-encoded names tried decoded too). */
function entryNameOf(zip: ZipPackage, name: string): string | undefined {
  const entry = zip.get(name);
  if (entry) return entry.name;
  try {
    return zip.get(decodeURIComponent(name))?.name;
  } catch {
    return undefined;
  }
}

function parseRels(text: string): XlsxRel[] {
  const rels: XlsxRel[] = [];
  for (const m of text.matchAll(/<(?:[\w.-]+:)?Relationship\b[^>]*>/g)) {
    const attrs = parseXmlAttrs(m[0]);
    rels.push({
      id: xmlAttrText(attrs, 'Id') ?? '',
      type: xmlAttrText(attrs, 'Type') ?? '',
      target: xmlAttrText(attrs, 'Target') ?? '',
      external: xmlAttrText(attrs, 'TargetMode') === 'External',
    });
  }
  return rels;
}

async function readPartText(zip: ZipPackage, name: string, onYield?: SaveYieldHook): Promise<string | undefined> {
  const entry = zip.get(name);
  return entry ? partText(await readEntry(zip, entry, onYield), name) : undefined;
}

/** Reads the package map: workbook, its relationships, sheets (sheetId → part), sharedStrings, styles, calcChain. */
async function openXlsxPackage(zip: ZipPackage, onYield?: SaveYieldHook): Promise<XlsxPackage> {
  const rootRels = await readPartText(zip, '_rels/.rels', onYield);
  if (rootRels === undefined) throw new PatchError('it has no package relationships (_rels/.rels)');
  const office = parseRels(rootRels).find((rel) => !rel.external && /\/officeDocument$/.test(rel.type));
  const wbPath = office && entryNameOf(zip, resolvePartTarget('', office.target));
  const wbXml = wbPath && (await readPartText(zip, wbPath, onYield));
  if (!wbPath || !wbXml) throw new PatchError('it has no workbook part');
  const wbRelsPath = entryNameOf(zip, relsPartOf(wbPath));
  const wbRelsXml = wbRelsPath && (await readPartText(zip, wbRelsPath, onYield));
  if (!wbRelsPath || !wbRelsXml) throw new PatchError('it has no workbook relationships');
  const ctXml = await readPartText(zip, '[Content_Types].xml', onYield);
  if (ctXml === undefined) throw new PatchError('it has no content types part');
  const P = xmlRootPrefix(wbXml, 'workbook', wbPath);
  const rootTag = new RegExp(`<${regexEscape(P)}workbook\\b[^>]*>`).exec(wbXml)?.[0] ?? '';
  const rels = parseRels(wbRelsXml).filter((rel) => !rel.external);
  const sheets: XlsxSheetRef[] = [];
  for (const m of wbXml.matchAll(new RegExp(`<${regexEscape(P)}sheet\\b[^>]*>`, 'g'))) {
    const attrs = parseXmlAttrs(m[0]);
    const ridName = [...attrs.keys()].find((key) => /(^|:)id$/.test(key));
    const rid = ridName ? xmlAttrText(attrs, ridName) : undefined;
    const rel = rid !== undefined ? rels.find((x) => x.id === rid) : undefined;
    sheets.push({
      name: xmlAttrText(attrs, 'name') ?? '',
      sheetId: Number(xmlAttrText(attrs, 'sheetId')),
      part: rel ? entryNameOf(zip, resolvePartTarget(wbPath, rel.target)) : undefined,
      type: rel?.type ?? '',
    });
  }
  const relPart = (suffix: string): string | undefined => {
    const rel = rels.find((x) => x.type.endsWith(suffix));
    return rel ? entryNameOf(zip, resolvePartTarget(wbPath, rel.target)) : undefined;
  };
  const calcChainPaths = new Set<string>();
  const chainRel = relPart('/calcChain');
  if (chainRel) calcChainPaths.add(chainRel);
  for (const m of ctXml.matchAll(/<(?:[\w.-]+:)?Override\b[^>]*>/g)) {
    const attrs = parseXmlAttrs(m[0]);
    const name = xmlAttrText(attrs, 'PartName');
    const entry = name && xmlAttrText(attrs, 'ContentType') === CALC_CHAIN_CONTENT_TYPE ? entryNameOf(zip, resolvePartTarget('', name)) : undefined;
    if (entry) calcChainPaths.add(entry);
  }
  return {
    zip, wbPath, wbXml, wbRelsPath, wbRelsXml, ctXml, sheets,
    sstPath: relPart('/sharedStrings'),
    stylesPath: relPart('/styles'),
    calcChainPaths: [...calcChainPaths],
    strict: /purl\.oclc\.org\/ooxml\/spreadsheetml\/main/.test(rootTag),
  };
}

// ----- sheet cells -----

/** Where <sheetData> is in a sheet part, plus the tag bytes used to walk it. */
interface SheetXmlLayout {
  P: string;
  where: string;
  sdOpen: number;
  sdOpenEnd: number;
  sdSelfClosing: boolean;
  /** Offset of </sheetData> (= sdOpenEnd when self-closing). */
  sdClose: number;
  rowTag: Buffer;
  rowEndTag: Buffer;
}

interface RowSpan {
  start: number;
  /** Offset just after the start tag. */
  openEnd: number;
  end: number;
  /** 0-based row index. */
  r: number;
  selfClosing: boolean;
  /** The start tag (ASCII). */
  tag: string;
}

interface CellSpan {
  c: number;
  start: number;
  end: number;
}

const ROW_NUMBER_RE = /\sr\s*=\s*["'](\d+)["']/;
const CELL_REF_ATTR_RE = /\sr\s*=\s*["']([A-Za-z]{1,3})(\d{1,7})["']/;

/** Offset of the next start tag `tag` (bytes of `<P:name`) in buf[from, end), or -1. */
function findStartTag(buf: Buffer, tag: Buffer, from: number, end: number): number {
  for (let pos = from; ;) {
    const window = buf.subarray(pos, end);
    const i = window.indexOf(tag);
    if (i < 0 || pos + i + tag.length >= end) return -1;
    if (isXmlNameEnd(buf[pos + i + tag.length])) return pos + i;
    pos += i + tag.length;
  }
}

/** Offset of the first `needle` in buf[from, end), or -1; big ranges are searched in 8 MB windows with a yield between. */
async function indexOfCooperative(buf: Buffer, needle: Buffer, from: number, onYield?: SaveYieldHook): Promise<number> {
  const WINDOW = 8 << 20;
  for (let pos = from; pos < buf.length; pos += WINDOW) {
    const i = buf.subarray(pos, Math.min(buf.length, pos + WINDOW + needle.length - 1)).indexOf(needle);
    if (i >= 0) return pos + i;
    if (buf.length - from > WINDOW) await saveYield(onYield);
  }
  return -1;
}

/**
 * Where <sheetData> is in a sheet part. The search for </sheetData> runs in windows with yields (one indexOf over a
 * 200 MB part blocks the event loop for 100-200 ms).
 */
async function sheetLayout(buf: Buffer, where: string, onYield?: SaveYieldHook): Promise<SheetXmlLayout> {
  assertUtf8Head(buf, where);
  const P = xmlRootPrefix(buf.toString('latin1', 0, Math.min(buf.length, 16384)), 'worksheet', where);
  const sdOpen = findStartTag(buf, Buffer.from(`<${P}sheetData`), 0, buf.length);
  if (sdOpen < 0) throw new PatchError(`${where} has no cell data (<sheetData>)`);
  const gt = buf.indexOf(0x3e, sdOpen);
  if (gt < 0) throw new PatchError(`${where} is damaged`);
  const sdSelfClosing = buf[gt - 1] === 0x2f;
  const sdClose = sdSelfClosing ? gt + 1 : await indexOfCooperative(buf, Buffer.from(`</${P}sheetData>`), gt, onYield);
  if (sdClose < 0) throw new PatchError(`${where} is damaged (no </sheetData>)`);
  return { P, where, sdOpen, sdOpenEnd: gt + 1, sdSelfClosing, sdClose, rowTag: Buffer.from(`<${P}row`), rowEndTag: Buffer.from(`</${P}row>`) };
}

/** The row starting at `rs`. Rows without a row number are refused (§4.7: hard read-only). */
function readRowAt(buf: Buffer, L: SheetXmlLayout, rs: number): RowSpan {
  const gt = buf.indexOf(0x3e, rs);
  if (gt < 0 || gt >= L.sdClose) throw new PatchError(`${L.where} is damaged`);
  const selfClosing = buf[gt - 1] === 0x2f;
  let end = gt + 1;
  if (!selfClosing) {
    const close = buf.indexOf(L.rowEndTag, gt);
    if (close < 0 || close >= L.sdClose) throw new PatchError(`${L.where} is damaged (a row is not closed)`);
    end = close + L.rowEndTag.length;
  }
  const tag = buf.toString('latin1', rs, gt + 1);
  const m = ROW_NUMBER_RE.exec(tag);
  if (!m) throw new PatchError(`${L.where} has rows without a row number`);
  const r = parseInt(m[1], 10) - 1;
  if (!(r >= 0 && r < MAX_ROWS)) throw new PatchError(`${L.where} has an invalid row number (${m[1]})`);
  return { start: rs, openEnd: gt + 1, end, r, selfClosing, tag };
}

/** The cells of a row body (text between the row tags): column, offsets. Checks order and row numbers. */
function parseRowCells(body: string, P: string, r: number, where: string): CellSpan[] {
  const cells: CellSpan[] = [];
  const cTag = `<${P}c`;
  const cEnd = `</${P}c>`;
  let p = 0;
  let prevC = -1;
  for (;;) {
    const i = body.indexOf(cTag, p);
    if (i < 0) break;
    if (!isXmlNameEnd(body.charCodeAt(i + cTag.length))) {
      p = i + cTag.length;
      continue;
    }
    const gt = body.indexOf('>', i);
    if (gt < 0) throw new PatchError(`${where} is damaged (row ${r + 1})`);
    let end = gt + 1;
    if (body[gt - 1] !== '/') {
      const close = body.indexOf(cEnd, gt);
      if (close < 0) throw new PatchError(`${where} is damaged (row ${r + 1})`);
      end = close + cEnd.length;
    }
    const ref = CELL_REF_ATTR_RE.exec(body.slice(i, gt + 1));
    let c = prevC + 1;
    if (ref) {
      c = colToIndex(ref[1]);
      if (parseInt(ref[2], 10) - 1 !== r) throw new PatchError(`${where} has cell ${ref[1]}${ref[2]} inside row ${r + 1}`);
      if (c < 0 || c >= MAX_COLS) throw new PatchError(`${where} has an invalid cell reference (${ref[1]}${ref[2]})`);
    }
    if (c <= prevC || c >= MAX_COLS) throw new PatchError(`${where} has cells out of order in row ${r + 1}`);
    prevC = c;
    cells.push({ c, start: i, end });
    p = end;
  }
  return cells;
}

/** Style index of each column from <cols> (0 = none). */
function columnStyles(buf: Buffer, L: SheetXmlLayout): (c: number) => number {
  const head = buf.toString('latin1', 0, L.sdOpen);
  const cols: { min: number; max: number; style: number }[] = [];
  for (const m of head.matchAll(new RegExp(`<${regexEscape(L.P)}col\\b[^>]*>`, 'g'))) {
    const attrs = parseXmlAttrs(m[0]);
    const style = Number(xmlAttrText(attrs, 'style') ?? 0);
    if (style > 0) cols.push({ min: Number(xmlAttrText(attrs, 'min')), max: Number(xmlAttrText(attrs, 'max')), style });
  }
  return (c: number): number => cols.find((col) => c + 1 >= col.min && c + 1 <= col.max)?.style ?? 0;
}

function decodeRowText(buf: Buffer, row: RowSpan, where: string): string {
  try {
    return UTF8_STRICT.decode(buf.subarray(row.start, row.end));
  } catch {
    throw new PatchError(`${where} row ${row.r + 1} is not valid UTF-8`);
  }
}

/**
 * Builds the new XML of one cell from its old XML (undefined = no cell yet) and the cell's style inherited from its
 * row or column (used for a new cell). null = the cell is removed.
 */
type CellPatchFn = (oldXml: string | undefined, r: number, c: number, inheritS: number) => string | null;

/**
 * Patches the cells of one sheet part. `rows`: 0-based row → (0-based column → patch function). Returns the new
 * part as slices (unchanged regions are views of `buf`, no copy). All rows are walked (yielding every `yieldEvery`
 * rows or SAVE_SLICE_MS) so rows out of order anywhere in the part are refused.
 */
async function patchSheetCells(
  buf: Buffer,
  L: SheetXmlLayout,
  rows: ReadonlyMap<number, ReadonlyMap<number, CellPatchFn>>,
  o: { yieldEvery: number; onYield?: SaveYieldHook },
): Promise<{ parts: Buffer[]; changed: boolean }> {
  const P = L.P;
  const colStyle = columnStyles(buf, L);
  const wanted = [...rows.keys()].sort((a, b) => a - b);
  const newCells = (r: number, inherit: (c: number) => number): string => {
    const fns = rows.get(r);
    if (!fns) return '';
    let cells = '';
    for (const c of [...fns.keys()].sort((a, b) => a - b)) cells += fns.get(c)?.(undefined, r, c, inherit(c)) ?? '';
    return cells;
  };
  const newRow = (r: number): string => {
    const cells = newCells(r, colStyle);
    return cells ? `<${P}row r="${r + 1}">${cells}</${P}row>` : '';
  };
  if (L.sdSelfClosing) {
    const rowsXml = wanted.map(newRow).join('');
    if (!rowsXml) return { parts: [buf], changed: false };
    return { parts: [buf.subarray(0, L.sdOpen), Buffer.from(`<${P}sheetData>${rowsXml}</${P}sheetData>`, 'utf8'), buf.subarray(L.sdOpenEnd)], changed: true };
  }
  const out: Buffer[] = [];
  let changed = false;
  let copyFrom = 0;
  let pos = L.sdOpenEnd;
  let prevR = -1;
  let wi = 0;
  const step = saveSlicer(o.onYield, o.yieldEvery);
  const emitNew = (r: number, at: number): void => {
    const xml = newRow(r);
    if (!xml) return;
    out.push(buf.subarray(copyFrom, at));
    copyFrom = at;
    out.push(Buffer.from(xml, 'utf8'));
    changed = true;
  };
  for (;;) {
    const rs = findStartTag(buf, L.rowTag, pos, L.sdClose);
    if (rs < 0) break;
    const row = readRowAt(buf, L, rs);
    if (row.r <= prevR) throw new PatchError(`${L.where} has rows out of order (row ${row.r + 1})`);
    prevR = row.r;
    while (wi < wanted.length && wanted[wi] < row.r) emitNew(wanted[wi++], rs);
    if (wi < wanted.length && wanted[wi] === row.r) {
      const fns = rows.get(row.r);
      const oldXml = decodeRowText(buf, row, L.where);
      const newXml = fns ? patchRowXml(oldXml, row, fns, P, colStyle, L.where) : oldXml;
      if (newXml !== oldXml) {
        out.push(buf.subarray(copyFrom, rs));
        out.push(Buffer.from(newXml, 'utf8'));
        copyFrom = row.end;
        changed = true;
      }
      wi++;
    }
    pos = row.end;
    const pause = step();
    if (pause) await pause;
  }
  while (wi < wanted.length) emitNew(wanted[wi++], L.sdClose);
  if (!changed) return { parts: [buf], changed: false };
  out.push(buf.subarray(copyFrom));
  return { parts: out, changed };
}

/** Re-emits one row: unchanged cells keep their exact text; patched cells are replaced or inserted in column order. */
function patchRowXml(rowXml: string, row: RowSpan, fns: ReadonlyMap<number, CellPatchFn>, P: string, colStyle: (c: number) => number, where: string): string {
  const rowAttrs = parseXmlAttrs(row.tag);
  const custom = xmlAttrText(rowAttrs, 'customFormat');
  const rowStyle = custom === '1' || custom === 'true' ? Number(xmlAttrText(rowAttrs, 's') ?? 0) : undefined;
  const inherit = (c: number): number => rowStyle ?? colStyle(c);
  const pending = [...fns.keys()].sort((a, b) => a - b);
  const run = (c: number, oldXml: string | undefined): string => fns.get(c)?.(oldXml, row.r, c, inherit(c)) ?? '';
  if (row.selfClosing) {
    const cells = pending.map((c) => run(c, undefined)).join('');
    return cells ? `${row.tag.slice(0, -2).replace(/\s+$/, '')}>${cells}</${P}row>` : rowXml;
  }
  const open = row.tag.length;
  const close = rowXml.lastIndexOf(`</${P}row>`);
  const body = rowXml.slice(open, close);
  const cells = parseRowCells(body, P, row.r, where);
  let res = rowXml.slice(0, open);
  let from = 0;
  let k = 0;
  for (const cell of cells) {
    while (k < pending.length && pending[k] < cell.c) {
      res += body.slice(from, cell.start) + run(pending[k++], undefined);
      from = cell.start;
    }
    if (k < pending.length && pending[k] === cell.c) {
      res += body.slice(from, cell.start) + run(pending[k++], body.slice(cell.start, cell.end));
      from = cell.end;
    }
  }
  if (k < pending.length) {
    const tail = Math.max(from, cells.length ? cells[cells.length - 1].end : 0);
    res += body.slice(from, tail);
    from = tail;
    while (k < pending.length) res += run(pending[k++], undefined);
  }
  return res + body.slice(from) + rowXml.slice(close);
}

/** Multi-cell array or data-table formula ranges of a sheet part (their cells cannot be changed one by one). */
async function findArrayRanges(buf: Buffer, L: SheetXmlLayout, onYield?: SaveYieldHook): Promise<{ range: Range; ref: string }[]> {
  const out: { range: Range; ref: string }[] = [];
  const fTag = new RegExp(`^<${regexEscape(L.P)}f[\\s/>]`);
  for (const word of ['array', 'dataTable']) {
    await forEachOccurrence(buf, Buffer.from(word), L.sdOpenEnd, L.sdClose, onYield, (at) => {
      const lt = buf.lastIndexOf(0x3c, at);
      const gt = buf.indexOf(0x3e, at);
      if (lt < 0 || gt < 0 || gt - lt > 4096) return;
      const tag = buf.toString('utf8', lt, gt + 1);
      if (!fTag.test(tag)) return;
      const attrs = parseXmlAttrs(tag);
      const ref = xmlAttrText(attrs, 'ref');
      if (xmlAttrText(attrs, 't') !== word || !ref) return;
      const range = decodeRange(ref);
      if (!range) throw new PatchError(`${L.where} has an array formula with an invalid range (${ref})`);
      if (range.r0 !== range.r1 || range.c0 !== range.c1) out.push({ range, ref });
    });
  }
  return out;
}

/** Cells of the shared formula `si`: the master (has `ref`) and its clones. */
async function findSharedFormulaCells(buf: Buffer, L: SheetXmlLayout, si: string, onYield?: SaveYieldHook): Promise<{ r: number; c: number; master: boolean; formula: string }[]> {
  const out: { r: number; c: number; master: boolean; formula: string }[] = [];
  const fTag = new RegExp(`^<${regexEscape(L.P)}f[\\s/>]`);
  const cTag = Buffer.from(`<${L.P}c`);
  const seen = new Set<number>();
  // The needle is the attribute NAME only: XML allows spaces around '=' (`si = "0"`), parseXmlAttrs reads the value.
  for (const needle of ['si']) {
    await forEachOccurrence(buf, Buffer.from(needle), L.sdOpenEnd, L.sdClose, onYield, (at) => {
      const lt = buf.lastIndexOf(0x3c, at);
      if (seen.has(lt)) return;
      const gt = buf.indexOf(0x3e, at);
      if (lt < 0 || gt < 0 || !isXmlNameEnd(buf[at - 1])) return;
      const tag = buf.toString('utf8', lt, gt + 1);
      if (!fTag.test(tag)) return;
      const attrs = parseXmlAttrs(tag);
      if (xmlAttrText(attrs, 't') !== 'shared' || xmlAttrText(attrs, 'si') !== si) return;
      seen.add(lt);
      let cs = buf.lastIndexOf(cTag, lt);
      while (cs >= 0 && !isXmlNameEnd(buf[cs + cTag.length])) cs = buf.lastIndexOf(cTag, cs - 1);
      const cellTag = cs >= 0 ? buf.toString('latin1', cs, buf.indexOf(0x3e, cs) + 1) : '';
      const ref = CELL_REF_ATTR_RE.exec(cellTag);
      if (!ref) throw new PatchError(`${L.where} has a shared formula in a cell without a reference`);
      let formula = '';
      if (buf[gt - 1] !== 0x2f) {
        const close = buf.indexOf(`</${L.P}f>`, gt);
        formula = xmlUnescape(buf.toString('utf8', gt + 1, close));
      }
      out.push({ r: parseInt(ref[2], 10) - 1, c: colToIndex(ref[1]), master: attrs.has('ref'), formula });
    });
  }
  return out;
}

const VALUE_ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
const CACHED_ERRORS = new Set([...VALUE_ERRORS, '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#EXTERNAL!', '#PYTHON!']);

/** What the cell patchers of one sheet share: prefix, lazy sharedStrings / styles, array ranges, de-share bookkeeping. */
interface CellPatchContext {
  P: string;
  where: string;
  /** undefined = the file has no sharedStrings part: strings are written inline. */
  sst(): SharedStringsPatch | undefined;
  styles(): StylesPatch;
  arrays: readonly { range: Range; ref: string }[];
  /** Shared-formula masters whose content a patch replaced (checked against their clones after the walk). */
  masters: { si: string; addr: string }[];
}

function numberText(value: number, where: string): string {
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new PatchError(`${where}: not a finite number`);
  // Excel stores 0 for anything below its smallest normal number: refuse instead of changing the value silently.
  if (value !== 0 && Math.abs(value) < 2.2250738585072014e-308) throw new PatchError(`${where}: the number is too small for Excel`);
  return String(value);
}

/** Longest text of one Excel cell. */
const XLSX_MAX_TEXT_CHARS = 32_767;

/**
 * Why Excel would remove or repair this formula (file syntax), or undefined. Checks only what is certain: the length
 * limit, an unterminated quote and unbalanced parentheses (quotes and [brackets] are skipped).
 */
function formulaSyntaxProblem(formula: string): string | undefined {
  if (formula.length > FX_MAX_FORMULA_CHARS) return `the formula is longer than ${FX_MAX_FORMULA_CHARS} characters, the limit of Excel`;
  let depth = 0;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      for (; j < formula.length; j++) {
        if (formula[j] !== ch) continue;
        if (formula[j + 1] === ch) j++; // a doubled quote is a quote character
        else break;
      }
      if (j >= formula.length) return 'the formula has a quote that is not closed';
      i = j;
    } else if (ch === '[') {
      let level = 1;
      i++;
      for (; i < formula.length && level > 0; i++) {
        if (formula[i] === '[') level++;
        else if (formula[i] === ']') level--;
      }
      i--;
    } else if (ch === '(') {
      depth++;
    } else if (ch === ')' && --depth < 0) {
      return 'the formula has a closing parenthesis without an opening one';
    }
  }
  return depth > 0 ? 'the formula has an opening parenthesis that is not closed' : undefined;
}

/** <v> (and the t attribute) for a value, or for the cached result of a formula (`cached` = true). */
function valueXml(value: FileValue, attrs: Map<string, XmlAttr>, ctx: CellPatchContext, where: string, cached: boolean): string {
  const P = ctx.P;
  switch (value.t) {
    case 'n':
      return `<${P}v>${numberText(value.v, where)}</${P}v>`;
    case 'b':
      if (typeof value.v !== 'boolean') throw new PatchError(`${where}: not a boolean`);
      setXmlAttr(attrs, 't', 'b');
      return `<${P}v>${value.v ? 1 : 0}</${P}v>`;
    case 'e':
      if (!(cached ? CACHED_ERRORS : VALUE_ERRORS).has(value.v)) throw new PatchError(`${where}: unknown error value ${String(value.v)}`);
      setXmlAttr(attrs, 't', 'e');
      return `<${P}v>${xmlEscapeText(value.v)}</${P}v>`;
    case 's': {
      if (typeof value.v !== 'string') throw new PatchError(`${where}: not a text`);
      if (value.v.length > XLSX_MAX_TEXT_CHARS) throw new PatchError(`${where}: the text is longer than ${XLSX_MAX_TEXT_CHARS} characters, the limit of one Excel cell`);
      if (cached) {
        setXmlAttr(attrs, 't', 'str');
        return `<${P}v>${escapeXstring(value.v, where)}</${P}v>`;
      }
      const sst = ctx.sst();
      if (sst) {
        setXmlAttr(attrs, 't', 's');
        return `<${P}v>${sst.add(value.v, where)}</${P}v>`;
      }
      setXmlAttr(attrs, 't', 'inlineStr');
      return `<${P}is><${P}t${needsSpacePreserve(value.v) ? ' xml:space="preserve"' : ''}>${escapeXstring(value.v, where)}</${P}t></${P}is>`;
    }
  }
  throw new PatchError(`${where}: unknown value type`);
}

/**
 * Applies one CellPatch to a cell (§4.3): value / clear keep `s`; a formula is written in file syntax with its
 * cached <v>; a cached-only patch keeps the <f> element bytes; style deltas derive a new xf from the cell's xf (a
 * new cell starts from its row/column style). Refused: content changes of a cell with `vm` (picture / linked data
 * type) or inside a multi-cell array / data-table range. A shared-formula master is recorded for the de-share check.
 */
function applyCellPatch(oldXml: string | undefined, patch: CellPatch, r: number, c: number, inheritS: number, ctx: CellPatchContext): string | null {
  const P = ctx.P;
  const addr = indexToCol(c) + (r + 1);
  const where = `${ctx.where}!${addr}`;
  if (patch.value !== undefined && patch.formula !== undefined) throw new PatchError(`${where}: a value and a formula in one change`);
  const content = patch.value !== undefined || patch.formula !== undefined;
  let attrs = new Map<string, XmlAttr>();
  let inner = '';
  let children: XmlChild[] = [];
  if (oldXml !== undefined) {
    const gt = startTagEnd(oldXml);
    attrs = parseXmlAttrs(oldXml.slice(0, gt + 1));
    if (oldXml[gt - 1] !== '/') {
      inner = oldXml.slice(gt + 1, oldXml.length - `</${P}c>`.length);
      children = splitXmlChildren(inner, where);
    }
  }
  setXmlAttr(attrs, 'r', addr);
  const setStyle = (xf: number): void => {
    if (xf === 0) attrs.delete('s');
    else setXmlAttr(attrs, 's', String(xf));
  };
  if (patch.style?.length) {
    const base = oldXml !== undefined ? Number(xmlAttrText(attrs, 's') ?? 0) : inheritS;
    setStyle(ctx.styles().deriveXf(base, patch.style, where));
  } else if (oldXml === undefined && inheritS > 0) {
    setStyle(inheritS);
  }
  const formulaChild = children.find((child) => child.local === 'f');
  const others = children.filter((child) => child.local !== 'f' && child.local !== 'v' && child.local !== 'is').map((child) => child.xml).join('');
  let body: string;
  if (content) {
    if (attrs.has('vm')) throw new PatchError(`${where} holds a picture or linked data type. Edit it in Excel.`);
    const array = ctx.arrays.find((a) => r >= a.range.r0 && r <= a.range.r1 && c >= a.range.c0 && c <= a.range.c1);
    if (array) throw new PatchError(`${where} is part of an array formula (${array.ref}). Edit it in Excel.`);
    if (formulaChild) {
      const fAttrs = parseXmlAttrs(formulaChild.open);
      if (xmlAttrText(fAttrs, 't') === 'shared' && fAttrs.has('ref')) ctx.masters.push({ si: xmlAttrText(fAttrs, 'si') ?? '', addr });
    }
    attrs.delete('t');
    attrs.delete('cm');
    attrs.delete('vm');
    if (patch.formula !== undefined) {
      const formula = patch.formula;
      if (typeof formula !== 'string' || !formula || formula.startsWith('=') || XML_INVALID_CHAR_RE.test(formula)) {
        throw new PatchError(`${where}: the formula must be in file syntax, without "=", and without control characters`);
      }
      const formulaProblem = formulaSyntaxProblem(formula);
      if (formulaProblem) throw new PatchError(`${where}: ${formulaProblem}`);
      body = `<${P}f>${xmlEscapeText(formula)}</${P}f>` + (patch.cached ? valueXml(patch.cached, attrs, ctx, where, true) : '');
    } else {
      body = patch.value ? valueXml(patch.value, attrs, ctx, where, false) : '';
    }
    body += others;
  } else if (patch.cached !== undefined) {
    if (!formulaChild) throw new PatchError(`${where} has no formula, so it cannot take a calculated value`);
    attrs.delete('t');
    body = formulaChild.xml + (patch.cached ? valueXml(patch.cached, attrs, ctx, where, true) : '') + others;
  } else {
    body = inner;
  }
  if (!body && !attrs.has('s') && inheritS === 0) return null; // an empty, unstyled cell is dropped
  const tag = `<${P}c${emitXmlAttrs(attrs)}`;
  return body ? `${tag}>${body}</${P}c>` : `${tag}/>`;
}

// ----- shared strings -----

/** sharedStrings.xml, append-only: new strings get new indexes; count / uniqueCount are updated. */
class SharedStringsPatch {
  private readonly P: string;
  private readonly baseCount: number;
  private readonly added: string[] = [];
  private readonly index = new Map<string, number>();

  constructor(private readonly text: string, private readonly where: string) {
    this.P = xmlRootPrefix(text, 'sst', where);
    let n = 0;
    const tag = `<${this.P}si`;
    for (let p = text.indexOf(tag); p >= 0; p = text.indexOf(tag, p + tag.length)) {
      if (isXmlNameEnd(text.charCodeAt(p + tag.length))) n++;
    }
    this.baseCount = n;
  }

  add(value: string, where: string): number {
    const known = this.index.get(value);
    if (known !== undefined) return known;
    const P = this.P;
    const i = this.baseCount + this.added.length;
    this.added.push(`<${P}si><${P}t${needsSpacePreserve(value) ? ' xml:space="preserve"' : ''}>${escapeXstring(value, where)}</${P}t></${P}si>`);
    this.index.set(value, i);
    return i;
  }

  get changed(): boolean {
    return this.added.length > 0;
  }

  toString(): string {
    const P = this.P;
    const m = new RegExp(`<${regexEscape(P)}sst\\b[^>]*>`).exec(this.text);
    if (!m) throw new PatchError(`${this.where} is damaged`);
    let open = m[0];
    open = open.replace(/(\suniqueCount\s*=\s*)(["'])\d*\2/, (_, pre: string, q: string) => `${pre}${q}${this.baseCount + this.added.length}${q}`);
    open = open.replace(/(\scount\s*=\s*)(["'])(\d*)\2/, (_, pre: string, q: string, n: string) => `${pre}${q}${Number(n) + this.added.length}${q}`);
    const after = m.index + m[0].length;
    if (open.endsWith('/>')) {
      return this.text.slice(0, m.index) + open.slice(0, -2).replace(/\s+$/, '') + '>' + this.added.join('') + `</${P}sst>` + this.text.slice(after);
    }
    const close = this.text.lastIndexOf(`</${P}sst>`);
    if (close < after) throw new PatchError(`${this.where} is damaged`);
    // New <si> go after the last existing <si>: the schema puts <extLst> last.
    const ext = this.text.slice(after, close).search(new RegExp(`<${regexEscape(P)}extLst\\b`));
    const at = ext >= 0 ? after + ext : close;
    return this.text.slice(0, m.index) + open + this.text.slice(after, at) + this.added.join('') + this.text.slice(at);
  }
}

// ----- styles -----

type StyleSectionName = 'numFmts' | 'fonts' | 'fills' | 'borders' | 'cellXfs';

interface StyleSection {
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  open: string;
  items: string[];
  selfClosing: boolean;
}

const FONT_CHILD_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const BORDER_CHILD_ORDER = ['start', 'end', 'left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];
const XF_CHILD_ORDER = ['alignment', 'protection', 'extLst'];
const DELTA_BORDER_STYLES: ReadonlySet<string> = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair']);
const ARGB_RE = /^[0-9A-Fa-f]{8}$/;
/** Built-in number formats that may be used by id (locale-independent ones only). */
const BUILTIN_NUMFMT_IDS: Record<string, number> = { General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13, '@': 49 };
/** Excel refuses more cell formats than this. */
const MAX_CELL_XFS = 64000;

/** The children of a container element of styles.xml (for example the <font> items of <fonts>). */
function styleSection(text: string, P: string, container: StyleSectionName, child: string, where: string): StyleSection | undefined {
  const m = new RegExp(`<${regexEscape(P)}${container}(?:\\s[^>]*?)?(/?)>`).exec(text);
  if (!m) return undefined;
  const start = m.index;
  const openEnd = start + m[0].length;
  if (m[1] === '/') return { start, openEnd, closeStart: openEnd, end: openEnd, open: m[0], items: [], selfClosing: true };
  const closeTag = `</${P}${container}>`;
  const closeStart = text.indexOf(closeTag, openEnd);
  if (closeStart < 0) throw new PatchError(`${where} is damaged (<${container}>)`);
  const body = text.slice(openEnd, closeStart);
  const items: string[] = [];
  const tag = `<${P}${child}`;
  const endTag = `</${P}${child}>`;
  for (let p = 0; ;) {
    const i = body.indexOf(tag, p);
    if (i < 0) break;
    if (!isXmlNameEnd(body.charCodeAt(i + tag.length))) {
      p = i + tag.length;
      continue;
    }
    const gt = body.indexOf('>', i);
    const end = body[gt - 1] === '/' ? gt + 1 : body.indexOf(endTag, gt) + endTag.length;
    if (gt < 0 || end < gt) throw new PatchError(`${where} is damaged (<${child}>)`);
    items.push(body.slice(i, end));
    p = end;
  }
  return { start, openEnd, closeStart, end: closeStart + closeTag.length, open: m[0], items, selfClosing: false };
}

/** Later deltas override earlier ones property by property (every StyleDelta property is an absolute "set"). */
function mergeStyleDeltas(deltas: readonly StyleDelta[]): StyleDelta {
  const out: StyleDelta = {};
  for (const delta of deltas) {
    const border = delta.border ? { ...out.border, ...delta.border } : out.border;
    Object.assign(out, delta);
    if (border) out.border = border;
  }
  return out;
}

/** styles.xml, append-only: derived fonts / fills / borders / number formats / cell xfs are added (deduplicated). */
class StylesPatch {
  private readonly P: string;
  private readonly sec: Record<StyleSectionName, StyleSection | undefined>;
  private readonly added: Record<StyleSectionName, string[]> = { numFmts: [], fonts: [], fills: [], borders: [], cellXfs: [] };
  private readonly memo = new Map<string, number>();
  /** Highest numFmtId used anywhere in the part (numFmts, xfs, dxfs): a new custom format gets a higher id. */
  private readonly maxNumFmtId: number;

  constructor(private readonly text: string, private readonly where: string) {
    let maxId = 163;
    for (const m of text.matchAll(/\snumFmtId\s*=\s*["'](\d+)["']/g)) maxId = Math.max(maxId, Number(m[1]));
    this.maxNumFmtId = maxId;
    const P = (this.P = xmlRootPrefix(text, 'styleSheet', where));
    this.sec = {
      numFmts: styleSection(text, P, 'numFmts', 'numFmt', where),
      fonts: styleSection(text, P, 'fonts', 'font', where),
      fills: styleSection(text, P, 'fills', 'fill', where),
      borders: styleSection(text, P, 'borders', 'border', where),
      cellXfs: styleSection(text, P, 'cellXfs', 'xf', where),
    };
    if (!this.sec.fonts || !this.sec.fills || !this.sec.cellXfs) throw new PatchError(`${where} has no fonts, fills or cell formats`);
  }

  private count(name: StyleSectionName): number {
    return (this.sec[name]?.items.length ?? 0) + this.added[name].length;
  }

  private item(name: StyleSectionName, i: number): string | undefined {
    const base = this.sec[name]?.items ?? [];
    return i < base.length ? base[i] : this.added[name][i - base.length];
  }

  /** Index of an item with exactly this XML, appending it when there is none. */
  private push(name: StyleSectionName, xml: string): number {
    const base = this.sec[name]?.items ?? [];
    const found = base.indexOf(xml);
    if (found >= 0) return found;
    const added = this.added[name].indexOf(xml);
    if (added >= 0) return base.length + added;
    this.added[name].push(xml);
    return this.count(name) - 1;
  }

  /** The xf that is xf `base` with the deltas applied (in order). Existing xfs never change. */
  deriveXf(base: number, deltas: readonly StyleDelta[], where: string): number {
    const key = `${base}|${JSON.stringify(deltas)}`;
    const memo = this.memo.get(key);
    if (memo !== undefined) return memo;
    const P = this.P;
    const xfXml = Number.isInteger(base) && base >= 0 ? this.item('cellXfs', base) : undefined;
    if (xfXml === undefined) throw new PatchError(`${where}: the cell format ${base} does not exist`);
    const d = mergeStyleDeltas(deltas);
    const xf = parseXmlElement(xfXml, this.where);
    const idOf = (name: string): number => Number(xmlAttrText(xf.attrs, name) ?? 0);
    if (d.bold !== undefined || d.italic !== undefined || d.underline !== undefined || d.strike !== undefined || d.fontColor !== undefined) {
      const fontXml = this.item('fonts', idOf('fontId'));
      if (fontXml === undefined) throw new PatchError(`${where}: the font ${idOf('fontId')} does not exist`);
      const font = parseXmlElement(fontXml, this.where);
      const flag = (local: string, on: boolean | undefined): void => {
        if (on !== undefined) setXmlChild(font, local, on ? `<${P}${local}/>` : null, FONT_CHILD_ORDER);
      };
      flag('b', d.bold);
      flag('i', d.italic);
      flag('strike', d.strike);
      if (d.underline !== undefined) {
        if (d.underline !== null && d.underline !== 'single' && d.underline !== 'double') throw new PatchError(`${where}: unknown underline`);
        setXmlChild(font, 'u', d.underline === null ? null : d.underline === 'double' ? `<${P}u val="double"/>` : `<${P}u/>`, FONT_CHILD_ORDER);
      }
      if (d.fontColor !== undefined) {
        if (d.fontColor !== null && !ARGB_RE.test(d.fontColor)) throw new PatchError(`${where}: invalid font colour`);
        setXmlChild(font, 'color', d.fontColor === null ? null : `<${P}color rgb="${d.fontColor.toUpperCase()}"/>`, FONT_CHILD_ORDER);
      }
      setXmlAttr(xf.attrs, 'fontId', String(this.push('fonts', buildXmlElement(font))));
      setXmlAttr(xf.attrs, 'applyFont', '1');
    }
    if (d.fill !== undefined) {
      if (d.fill !== null && !ARGB_RE.test(d.fill)) throw new PatchError(`${where}: invalid fill colour`);
      let fillId: number;
      if (d.fill === null) {
        const first = this.item('fills', 0) ?? '';
        fillId = /patternType\s*=\s*["']none["']/.test(first) ? 0 : this.push('fills', `<${P}fill><${P}patternFill patternType="none"/></${P}fill>`);
      } else {
        fillId = this.push('fills', `<${P}fill><${P}patternFill patternType="solid"><${P}fgColor rgb="${d.fill.toUpperCase()}"/><${P}bgColor indexed="64"/></${P}patternFill></${P}fill>`);
      }
      setXmlAttr(xf.attrs, 'fillId', String(fillId));
      setXmlAttr(xf.attrs, 'applyFill', '1');
    }
    if (d.border) {
      const borderXml = this.item('borders', idOf('borderId'));
      if (borderXml === undefined) throw new PatchError(`${where}: the border ${idOf('borderId')} does not exist`);
      const border = parseXmlElement(borderXml, this.where);
      for (const side of ['left', 'right', 'top', 'bottom'] as const) {
        const value = d.border[side];
        if (value === undefined) continue;
        if (value !== null && (!DELTA_BORDER_STYLES.has(value.style) || (value.color !== undefined && !ARGB_RE.test(value.color)))) {
          throw new PatchError(`${where}: invalid border`);
        }
        const xml = value === null
          ? `<${P}${side}/>`
          : value.color
            ? `<${P}${side} style="${value.style}"><${P}color rgb="${value.color.toUpperCase()}"/></${P}${side}>`
            : `<${P}${side} style="${value.style}"/>`;
        setXmlChild(border, side, xml, BORDER_CHILD_ORDER);
      }
      setXmlAttr(xf.attrs, 'borderId', String(this.push('borders', buildXmlElement(border))));
      setXmlAttr(xf.attrs, 'applyBorder', '1');
    }
    if (d.hAlign !== undefined || d.vAlign !== undefined || d.wrap !== undefined) {
      const current = xf.children.find((child) => child.local === 'alignment');
      const align = current ? parseXmlAttrs(current.open) : new Map<string, XmlAttr>();
      if (d.hAlign !== undefined) {
        if (d.hAlign !== null && !['left', 'center', 'right', 'justify'].includes(d.hAlign)) throw new PatchError(`${where}: invalid alignment`);
        if (d.hAlign === null) align.delete('horizontal');
        else setXmlAttr(align, 'horizontal', d.hAlign);
      }
      if (d.vAlign !== undefined) {
        if (d.vAlign !== null && !['top', 'center', 'bottom'].includes(d.vAlign)) throw new PatchError(`${where}: invalid alignment`);
        if (d.vAlign === null) align.delete('vertical');
        else setXmlAttr(align, 'vertical', d.vAlign);
      }
      if (d.wrap !== undefined) {
        if (d.wrap) setXmlAttr(align, 'wrapText', '1');
        else align.delete('wrapText');
      }
      setXmlChild(xf, 'alignment', align.size ? `<${P}alignment${emitXmlAttrs(align)}/>` : null, XF_CHILD_ORDER);
      setXmlAttr(xf.attrs, 'applyAlignment', '1');
    }
    if (d.numFmt !== undefined) {
      setXmlAttr(xf.attrs, 'numFmtId', String(this.numFmtId(d.numFmt, where)));
      setXmlAttr(xf.attrs, 'applyNumberFormat', '1');
    }
    if (d.quotePrefix !== undefined) {
      if (d.quotePrefix) setXmlAttr(xf.attrs, 'quotePrefix', '1');
      else xf.attrs.delete('quotePrefix');
    }
    const index = this.push('cellXfs', buildXmlElement(xf));
    if (index >= MAX_CELL_XFS) throw new PatchError(`${where}: the file has too many cell formats`);
    this.memo.set(key, index);
    return index;
  }

  /**
   * Id of a number format code: an existing <numFmt> with that code, a built-in id, or a new id ≥ 164 that no part of
   * styles.xml uses yet (an xf may name a custom id that has no <numFmt>; reusing it would change those cells).
   */
  private numFmtId(code: string, where: string): number {
    if (typeof code !== 'string' || !code || code.length > 255 || XML_INVALID_CHAR_RE.test(code)) throw new PatchError(`${where}: invalid number format`);
    const items = (this.sec.numFmts?.items ?? []).concat(this.added.numFmts);
    const ids = new Set<number>();
    for (const item of items) {
      const attrs = parseXmlAttrs(item);
      const id = Number(xmlAttrText(attrs, 'numFmtId'));
      if (xmlAttrText(attrs, 'formatCode') === code && Number.isInteger(id)) return id;
      if (Number.isInteger(id)) ids.add(id);
    }
    const builtin = Object.hasOwn(BUILTIN_NUMFMT_IDS, code) ? BUILTIN_NUMFMT_IDS[code] : undefined;
    if (builtin !== undefined && !ids.has(builtin)) return builtin;
    const id = Math.max(this.maxNumFmtId, ...ids) + 1;
    this.added.numFmts.push(`<${this.P}numFmt numFmtId="${id}" formatCode="${xmlEscapeAttr(code)}"/>`);
    return id;
  }

  get changed(): boolean {
    return Object.values(this.added).some((list) => list.length > 0);
  }

  toString(): string {
    const P = this.P;
    let text = this.text;
    const withCount = (open: string, total: number): string =>
      /\scount\s*=/.test(open) ? open.replace(/(\scount\s*=\s*)(["'])\d*\2/, (_, pre: string, q: string) => `${pre}${q}${total}${q}`) : open.replace(/\s*(\/?)>$/, ` count="${total}"$1>`);
    // Apply from the end of the text backwards, so the offsets of earlier sections stay valid.
    const names = (Object.keys(this.sec) as StyleSectionName[]).filter((name) => this.sec[name] && this.added[name].length);
    names.sort((a, b) => (this.sec[b]?.start ?? 0) - (this.sec[a]?.start ?? 0));
    for (const name of names) {
      const s = this.sec[name];
      if (!s) continue;
      const open = withCount(s.open, s.items.length + this.added[name].length);
      const items = this.added[name].join('');
      text = s.selfClosing
        ? text.slice(0, s.start) + open.slice(0, -2).replace(/\s+$/, '') + '>' + items + `</${P}${name}>` + text.slice(s.end)
        : text.slice(0, s.start) + open + text.slice(s.openEnd, s.closeStart) + items + text.slice(s.closeStart);
    }
    if (!this.sec.numFmts && this.added.numFmts.length) {
      const m = new RegExp(`<${regexEscape(P)}styleSheet\\b[^>]*>`).exec(text);
      if (!m) throw new PatchError(`${this.where} is damaged`);
      const at = m.index + m[0].length;
      text = text.slice(0, at) + `<${P}numFmts count="${this.added.numFmts.length}">${this.added.numFmts.join('')}</${P}numFmts>` + text.slice(at);
    }
    return text;
  }
}

// ----- workbook -----

/** CT_Workbook children that come after <calcPr> (where a missing calcPr is inserted). */
const AFTER_CALC_PR = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

/** <calcPr fullCalcOnLoad="1"> with every other calc setting kept; calcPr is inserted in schema order when missing. */
function setFullCalcOnLoad(text: string, where: string): string {
  const P = xmlRootPrefix(text, 'workbook', where);
  const m = new RegExp(`<${regexEscape(P)}calcPr\\b[^>]*>`).exec(text);
  if (m) {
    const tag = m[0];
    const flag = /(\sfullCalcOnLoad\s*=\s*)(["'])([^"']*)\2/.exec(tag);
    if (flag && (flag[3] === '1' || flag[3] === 'true')) return text;
    const next = flag ? tag.replace(flag[0], `${flag[1]}${flag[2]}1${flag[2]}`) : tag.replace(/\s*(\/?)>$/, ' fullCalcOnLoad="1"$1>');
    return text.slice(0, m.index) + next + text.slice(m.index + tag.length);
  }
  let at = text.lastIndexOf(`</${P}workbook>`);
  if (at < 0) throw new PatchError(`${where} is damaged`);
  for (const name of AFTER_CALC_PR) {
    const i = text.search(new RegExp(`<${regexEscape(P)}${name}[\\s/>]`));
    if (i >= 0 && i < at) at = i;
  }
  return text.slice(0, at) + `<${P}calcPr fullCalcOnLoad="1"/>` + text.slice(at);
}

/** Removes the elements `<Name …/>` / `<Name …>…</Name>` (any prefix) for which `drop(attrs)` is true. */
function removeXmlElements(text: string, name: string, drop: (attrs: Map<string, XmlAttr>) => boolean): string {
  const re = new RegExp(`<((?:[\\w.-]+:)?)${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</\\1${name}>)`, 'g');
  return text.replace(re, (element: string) => (drop(parseXmlAttrs(element.slice(0, element.indexOf('>') + 1))) ? '' : element));
}

/** [Content_Types].xml and workbook rels without the calcChain parts. */
function dropCalcChainRefs(ctXml: string, wbRelsXml: string, wbPath: string, chainPaths: readonly string[]): { ctXml: string; wbRelsXml: string } {
  const names = new Set(chainPaths.map((path) => path.toLowerCase()));
  return {
    ctXml: removeXmlElements(ctXml, 'Override', (attrs) => names.has(resolvePartTarget('', xmlAttrText(attrs, 'PartName') ?? '').toLowerCase())),
    wbRelsXml: removeXmlElements(wbRelsXml, 'Relationship', (attrs) =>
      xmlAttrText(attrs, 'TargetMode') !== 'External' && names.has(resolvePartTarget(wbPath, xmlAttrText(attrs, 'Target') ?? '').toLowerCase())),
  };
}

// ===== XLSX SAVE =====
//
// Surgical xlsx save (P1, D-006; contract §2.3, §4.2). buildXlsx = the loaded bytes (`base`) + the PackageOps of the
// journal: only the touched parts are rewritten, every other zip entry is copied byte for byte. verifyXlsx checks the
// new bytes before anything is written (V1 zip, V2 XML, V3 read-back, V4 reload with the loader); checkPatchable tells
// at load whether a file can be saved this way at all (else it is hard read-only, §4.7). ExcelJS never writes a file.


export interface SaveReport {
  changedParts: string[];
  removedParts: string[];
  identicalParts: number;
  timings: Record<string, number>;
}

/**
 * What one saved cell must hold (V3/V4). The expected value and formula come from the display model (§4.2 step 4).
 * Content is checked when `value` or `formula` is given: `formula` absent = the cell has no formula; `value` with a
 * formula = its cached value (absent = not checked); `value: null` = no value. `styleIndex` = the xf index (s).
 */
export interface CellExpectation {
  sheet: number;
  r: number;
  c: number;
  value?: FileValue | null;
  formula?: string;
  styleIndex?: number;
}

/** A check of the new bytes failed (§4.2 step 4): nothing is written, the edits are kept. */
export class SaveVerifyError extends Error {
  constructor(readonly step: 'V1' | 'V2' | 'V3' | 'V4', message: string) {
    super(`${step}: ${message}`);
    this.name = 'SaveVerifyError';
  }
}

/** Rewritten parts up to this size are parsed whole in V2; bigger cell-patched sheet parts are checked row by row. (An object so tests can lower it.) */
export const verifyTuning = { fullParseBytes: 32 << 20 };
/** V4 (reload with the loader) runs for files up to this size unless the caller says otherwise (D-008). */
export const VERIFY_DEEP_MAX_BYTES = 10 << 20;
/** V4 also needs a modest sharedStrings part: text-heavy workbooks compress 25:1, a 3.7 MB file can hold 95 MB of strings. */
export const VERIFY_DEEP_MAX_SST_BYTES = 32 << 20;

/** Whether V4 (a synchronous ExcelJS reload) may run: the zip is small AND its parts are small once unpacked. */
export function deepVerifyAllowed(bytes: Uint8Array): boolean {
  if (bytes.length > VERIFY_DEEP_MAX_BYTES) return false;
  try {
    return !parseZip(bytes).entries.some((e) => /sharedStrings[^/]*$/i.test(e.name) && e.usize > VERIFY_DEEP_MAX_SST_BYTES);
  } catch {
    return false; // V1 reports a damaged zip
  }
}
/** Most rows walked between two yields to the event loop (a walk also yields after SAVE_SLICE_MS, see saveSlicer). */
const SAVE_YIELD_ROWS = 10_000;
/** Text decoded and parsed between two yields in V2. */
const VERIFY_PARSE_CHUNK = 256 << 10;

const WORKSHEET_REL_RE = /\/worksheet$/;

function saveNow(): number {
  return performance.now();
}

/** Cell content patch (value or formula), as opposed to style-only or cached-only patches. */
function isContentPatch(patch: CellPatch): boolean {
  return patch.value !== undefined || patch.formula !== undefined;
}

/**
 * Header cells of the tables of a sheet as {r0, r1, c0, c1} boxes (0-based, inclusive). Excel needs the header text to
 * equal the tableColumn name and repairs the file otherwise, so a content patch there is refused (defense in depth,
 * the edit layer blocks these cells too).
 */
async function tableHeaderBoxes(pkg: XlsxPackage, sheetPart: string, onYield?: SaveYieldHook): Promise<{ r0: number; r1: number; c0: number; c1: number }[]> {
  const out: { r0: number; r1: number; c0: number; c1: number }[] = [];
  const relsPath = entryNameOf(pkg.zip, relsPartOf(sheetPart));
  const relsXml = relsPath ? await readPartText(pkg.zip, relsPath, onYield) : undefined;
  if (!relsXml) return out;
  for (const rel of parseRels(relsXml)) {
    if (rel.external || !/\/table$/.test(rel.type)) continue;
    const tablePath = entryNameOf(pkg.zip, resolvePartTarget(sheetPart, rel.target));
    const xml = tablePath ? await readPartText(pkg.zip, tablePath, onYield) : undefined;
    const tag = xml ? /<(?:[\w.-]+:)?table\b[^>]*>/.exec(xml)?.[0] : undefined;
    if (!tag) continue;
    const attrs = parseXmlAttrs(tag);
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec((xmlAttrText(attrs, 'ref') ?? '').trim());
    if (!m) continue;
    const headerRows = xmlAttrText(attrs, 'headerRowCount') === undefined ? 1 : Number(xmlAttrText(attrs, 'headerRowCount'));
    if (!(headerRows > 0)) continue;
    const r0 = parseInt(m[2], 10) - 1;
    out.push({ r0, r1: r0 + headerRows - 1, c0: colToIndex(m[1].toUpperCase()), c1: colToIndex((m[3] ?? m[1]).toUpperCase()) });
  }
  return out;
}

/**
 * Cells written outside the old <dimension ref>: widen it (Excel does not care, readers that trust the dimension do).
 * The head of the part (everything before <sheetData>) is always inside parts[0].
 */
function widenDimension(parts: Buffer[], L: SheetXmlLayout, rows: Map<number, Map<number, CellPatch[]>>): void {
  let r0 = Infinity;
  let r1 = -1;
  let c0 = Infinity;
  let c1 = -1;
  for (const [r, cols] of rows) {
    for (const [c, list] of cols) {
      if (!list.some((p) => isContentPatch(p) && p.value !== null)) continue;
      r0 = Math.min(r0, r);
      r1 = Math.max(r1, r);
      c0 = Math.min(c0, c);
      c1 = Math.max(c1, c);
    }
  }
  if (r1 < 0 || parts[0].length < L.sdOpen) return;
  const head = parts[0].toString('utf8', 0, L.sdOpen);
  const re = new RegExp(`<${regexEscape(L.P)}dimension\\b[^>]*>`);
  const tag = re.exec(head)?.[0];
  if (!tag) return;
  const attrs = parseXmlAttrs(tag);
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec((xmlAttrText(attrs, 'ref') ?? '').trim());
  if (!m) return;
  const a0 = colToIndex(m[1].toUpperCase());
  const a1 = colToIndex((m[3] ?? m[1]).toUpperCase());
  const b0 = parseInt(m[2], 10) - 1;
  const b1 = parseInt(m[4] ?? m[2], 10) - 1;
  const n0 = Math.min(a0, c0);
  const n1 = Math.max(a1, c1);
  const q0 = Math.min(b0, r0);
  const q1 = Math.max(b1, r1);
  if (n0 === a0 && n1 === a1 && q0 === b0 && q1 === b1) return;
  const ref = `${indexToCol(n0)}${q0 + 1}:${indexToCol(n1)}${q1 + 1}`;
  const newTag = tag.replace(/(\sref\s*=\s*)("[^"]*"|'[^']*')/, `$1"${ref}"`);
  parts[0] = Buffer.concat([Buffer.from(head.replace(tag, () => newTag), 'utf8'), parts[0].subarray(L.sdOpen)]);
}

function sheetLabel(sheet: XlsxSheetRef): string {
  return sheet.name || `sheet ${sheet.sheetId}`;
}

/** The worksheet with this sheetId (PatchError when there is none or it is not a worksheet, e.g. a chart sheet). */
function worksheetOf(pkg: XlsxPackage, sheetId: number): XlsxSheetRef & { part: string } {
  const sheet = pkg.sheets.find((s) => s.sheetId === sheetId);
  if (!sheet) throw new PatchError(`the workbook has no sheet with id ${sheetId}`);
  if (!WORKSHEET_REL_RE.test(sheet.type)) throw new PatchError(`the sheet "${sheetLabel(sheet)}" is not a worksheet`);
  if (!sheet.part || !pkg.zip.get(sheet.part)) throw new PatchError(`the part of the sheet "${sheetLabel(sheet)}" is missing`);
  return sheet as XlsxSheetRef & { part: string };
}

/**
 * The new package for `base` with `ops` applied (§4.2 step 3). Only `cells` ops exist in P1 (others throw PatchError
 * until P4). No ops, or ops that change nothing, return `base` itself: an unmodified workbook is never re-serialized.
 * Throws ZipError / PatchError when the package cannot be patched safely (nothing is written). Cooperative: yields
 * every `yieldEvery` rows, 8 MB of CRC and between parts; `onYield` runs at every yield and may throw to cancel.
 */
export async function buildXlsx(
  base: Uint8Array,
  ops: readonly PackageOp[],
  o?: { yieldEvery?: number; onYield?: SaveYieldHook },
): Promise<{ bytes: Uint8Array; report: SaveReport }> {
  const t0 = saveNow();
  const timings: Record<string, number> = {};
  const lap = (name: string, from: number): number => {
    const now = saveNow();
    timings[name] = (timings[name] ?? 0) + (now - from);
    return now;
  };
  const unchanged = (identical: number): { bytes: Uint8Array; report: SaveReport } => {
    timings.total = saveNow() - t0;
    return { bytes: base, report: { changedParts: [], removedParts: [], identicalParts: identical, timings } };
  };
  if (!ops.length) {
    let identical = 0;
    try {
      identical = parseZip(base).entries.length;
    } catch {
      // Save As of a clean, unpatchable (hard read-only) file: base is written as it is.
    }
    return unchanged(identical);
  }
  const onYield = o?.onYield;
  const yieldEvery = Math.max(1, Math.floor(o?.yieldEvery ?? SAVE_YIELD_ROWS));
  let t = t0;
  const zip = parseZip(base);
  const pkg = await openXlsxPackage(zip, onYield);
  if (pkg.strict) throw new PatchError('it uses Strict Open XML');
  // Group the cell patches per sheet and cell, in op order (later patches apply on top of earlier ones).
  const bySheet = new Map<number, Map<number, Map<number, CellPatch[]>>>();
  let needSst = false;
  let needStyles = false;
  for (const op of ops) {
    if (op.kind !== 'cells') throw new PatchError(`saving "${op.kind}" changes is not supported yet`);
    let rows = bySheet.get(op.sheet);
    if (!rows) bySheet.set(op.sheet, (rows = new Map()));
    for (const [r, c, patch] of op.cells) {
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) {
        throw new PatchError(`a change has an invalid cell position (${r}, ${c})`);
      }
      let cols = rows.get(r);
      if (!cols) rows.set(r, (cols = new Map()));
      let list = cols.get(c);
      if (!list) cols.set(c, (list = []));
      list.push(patch);
      if (patch.value && patch.value.t === 's') needSst = true;
      if (patch.style?.length) needStyles = true;
    }
  }
  let sst: SharedStringsPatch | undefined;
  if (needSst && pkg.sstPath) {
    sst = new SharedStringsPatch((await readPartText(zip, pkg.sstPath, onYield)) ?? '', pkg.sstPath);
  }
  let styles: StylesPatch | undefined;
  if (needStyles) {
    if (!pkg.stylesPath) throw new PatchError('it has no styles part');
    styles = new StylesPatch((await readPartText(zip, pkg.stylesPath, onYield)) ?? '', pkg.stylesPath);
  }
  t = lap('open', t);
  const changes = new Map<string, readonly Uint8Array[] | null>();
  for (const [sheetId, rows] of bySheet) {
    const sheet = worksheetOf(pkg, sheetId);
    const entry = zip.get(sheet.part);
    if (!entry) throw new PatchError(`the part of the sheet "${sheetLabel(sheet)}" is missing`);
    const buf = await readEntry(zip, entry, onYield);
    t = lap('inflate', t);
    const L = await sheetLayout(buf, `the sheet "${sheetLabel(sheet)}"`, onYield);
    const contentCells = new Set<number>();
    for (const [r, cols] of rows) {
      for (const [c, list] of cols) if (list.some(isContentPatch)) contentCells.add(r * MAX_COLS + c);
    }
    if (contentCells.size) {
      for (const box of await tableHeaderBoxes(pkg, sheet.part, onYield)) {
        for (const key of contentCells) {
          const r = Math.floor(key / MAX_COLS);
          const c = key % MAX_COLS;
          if (r >= box.r0 && r <= box.r1 && c >= box.c0 && c <= box.c1) {
            throw new PatchError(`${sheetLabel(sheet)}!${indexToCol(c)}${r + 1} is a table header cell (Excel needs it to match the table column name)`);
          }
        }
      }
    }
    const ctx: CellPatchContext = {
      P: L.P,
      where: sheetLabel(sheet),
      sst: () => sst,
      styles: () => {
        if (!styles) throw new PatchError('the style table was not loaded');
        return styles;
      },
      arrays: contentCells.size ? await findArrayRanges(buf, L, onYield) : [],
      masters: [],
    };
    const fns = new Map<number, Map<number, CellPatchFn>>();
    for (const [r, cols] of rows) {
      const rowFns = new Map<number, CellPatchFn>();
      for (const [c, list] of cols) {
        rowFns.set(c, (oldXml, rr, cc, inheritS) => {
          let xml = oldXml;
          for (const patch of list) xml = applyCellPatch(xml, patch, rr, cc, inheritS, ctx) ?? undefined;
          return xml ?? null;
        });
      }
      fns.set(r, rowFns);
    }
    const res = await patchSheetCells(buf, L, fns, { yieldEvery, onYield });
    // Defense in depth (D-068): a shared-formula master whose content changed must have no clone left that still
    // points at it (P2 de-shares at plan time; a clone without its own new content would lose its formula).
    for (const master of ctx.masters) {
      for (const cell of await findSharedFormulaCells(buf, L, master.si, onYield)) {
        if (!cell.master && !contentCells.has(cell.r * MAX_COLS + cell.c)) {
          throw new PatchError(
            `${ctx.where}!${master.addr} is the first cell of a shared formula that ${indexToCol(cell.c)}${cell.r + 1} still uses`,
          );
        }
      }
    }
    if (res.changed) {
      widenDimension(res.parts, L, rows);
      changes.set(entry.name, res.parts);
    }
    t = lap('sheets', t);
    await saveYield(onYield);
  }
  if (!changes.size) return unchanged(zip.entries.length);
  if (sst?.changed && pkg.sstPath) changes.set(pkg.sstPath, [Buffer.from(sst.toString(), 'utf8')]);
  if (styles?.changed && pkg.stylesPath) changes.set(pkg.stylesPath, [Buffer.from(styles.toString(), 'utf8')]);
  t = lap('sharedStringsStyles', t);
  // workbook: full calculation on load (cached values may be stale); the calc chain is dropped with its relationship
  // and content type (a stale chain makes Excel repair the file; s2-save negative control).
  const wbXml = setFullCalcOnLoad(pkg.wbXml, pkg.wbPath);
  if (wbXml !== pkg.wbXml) changes.set(pkg.wbPath, [Buffer.from(wbXml, 'utf8')]);
  if (pkg.calcChainPaths.length) {
    const dropped = dropCalcChainRefs(pkg.ctXml, pkg.wbRelsXml, pkg.wbPath, pkg.calcChainPaths);
    if (dropped.ctXml !== pkg.ctXml) changes.set('[Content_Types].xml', [Buffer.from(dropped.ctXml, 'utf8')]);
    if (dropped.wbRelsXml !== pkg.wbRelsXml) changes.set(pkg.wbRelsPath, [Buffer.from(dropped.wbRelsXml, 'utf8')]);
    for (const path of pkg.calcChainPaths) changes.set(path, null);
  }
  t = lap('workbook', t);
  const written = await writeZip(zip, changes, { onYield });
  lap('zip', t);
  timings.total = saveNow() - t0;
  return {
    bytes: written.bytes,
    report: { changedParts: written.changed, removedParts: written.removed, identicalParts: written.identical, timings },
  };
}

// ----- verify -----

/** V2: the whole part is well-formed UTF-8 XML (saxes, fed in 256 KB pieces with yields in between). */
async function verifyXmlWellFormed(data: Uint8Array, where: string, onYield?: SaveYieldHook): Promise<void> {
  const parser = new SaxesParser({ xmlns: false });
  let error: string | undefined;
  parser.on('error', (err: Error) => {
    error ??= err.message;
  });
  const decoder = new TextDecoder('utf-8', { fatal: true });
  for (let i = 0; i < data.length && error === undefined; i += VERIFY_PARSE_CHUNK) {
    try {
      parser.write(decoder.decode(data.subarray(i, Math.min(data.length, i + VERIFY_PARSE_CHUNK)), { stream: true }));
    } catch (err) {
      error ??= saveErrorText(err);
    }
    if (data.length > VERIFY_PARSE_CHUNK) await saveYield(onYield);
  }
  if (error === undefined) {
    try {
      const rest = decoder.decode();
      if (rest) parser.write(rest);
      parser.close();
    } catch (err) {
      error ??= saveErrorText(err);
    }
  }
  if (error !== undefined) throw new SaveVerifyError('V2', `${where} is not well-formed XML (${error})`);
}

/** Row spans of a sheet part's <sheetData> (walks the whole part; yields every SAVE_YIELD_ROWS rows or SAVE_SLICE_MS). */
async function sheetRowSpans(buf: Buffer, L: SheetXmlLayout, onYield?: SaveYieldHook): Promise<RowSpan[]> {
  const rows: RowSpan[] = [];
  if (L.sdSelfClosing) return rows;
  const step = saveSlicer(onYield, SAVE_YIELD_ROWS);
  for (let pos = L.sdOpenEnd; ;) {
    const rs = findStartTag(buf, L.rowTag, pos, L.sdClose);
    if (rs < 0) break;
    const row = readRowAt(buf, L, rs);
    rows.push(row);
    pos = row.end;
    const pause = step();
    if (pause) await pause;
  }
  return rows;
}

/**
 * V2 for a big sheet part that only had cell patches: every patched row is parsed alone (well-formed), the rows are
 * in order, and the part without the patched rows is byte-identical to the base part without them (so every splice
 * point is a row element boundary and nothing else moved). Self-closing <sheetData/> in the base → full parse.
 */
async function verifySheetRows(baseBuf: Buffer, newBuf: Buffer, patchedRows: ReadonlySet<number>, where: string, onYield?: SaveYieldHook): Promise<void> {
  let baseRows: RowSpan[];
  let newRows: RowSpan[];
  let LB: SheetXmlLayout;
  let LN: SheetXmlLayout;
  try {
    LB = await sheetLayout(baseBuf, where, onYield);
    if (LB.sdSelfClosing) {
      await verifyXmlWellFormed(newBuf, where, onYield);
      return;
    }
    baseRows = await sheetRowSpans(baseBuf, LB, onYield);
    LN = await sheetLayout(newBuf, where, onYield);
    newRows = await sheetRowSpans(newBuf, LN, onYield);
  } catch (err) {
    if (err instanceof PatchError) throw new SaveVerifyError('V2', saveErrorText(err));
    throw err;
  }
  // The <dimension ref> in the head may be widened by the patch (widenDimension): only its ref may differ, and the new
  // ref must be a plain range. Everything else of the head is compared byte by byte below.
  const dimOf = (buf: Buffer, L: SheetXmlLayout): { start: number; end: number; tag: string } | undefined => {
    const head = buf.toString('latin1', 0, L.sdOpen);
    const m = new RegExp(`<${regexEscape(L.P)}dimension\\b[^>]*>`).exec(head);
    return m ? { start: m.index, end: m.index + m[0].length, tag: m[0] } : undefined;
  };
  const dimB = dimOf(baseBuf, LB);
  const dimN = dimOf(newBuf, LN);
  let dimSkip = false;
  if (dimB && dimN && dimB.tag !== dimN.tag) {
    const withoutRef = (tag: string): string => tag.replace(/(\sref\s*=\s*)("[^"]*"|'[^']*')/, '$1""');
    const ref = /\sref\s*=\s*"([^"]*)"/.exec(dimN.tag)?.[1] ?? '';
    if (withoutRef(dimB.tag) !== withoutRef(dimN.tag) || !/^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$/.test(ref)) {
      throw new SaveVerifyError('V2', `${where} changed outside the edited rows`);
    }
    dimSkip = true;
  }
  let prevR = -1;
  for (const row of newRows) {
    if (row.r <= prevR) throw new SaveVerifyError('V2', `${where} has rows out of order (row ${row.r + 1})`);
    prevR = row.r;
    if (patchedRows.has(row.r)) await verifyXmlWellFormed(newBuf.subarray(row.start, row.end), `${where} row ${row.r + 1}`, onYield);
  }
  /** The part as a list of [start, end) spans without the patched rows. */
  const kept = (rows: readonly RowSpan[], size: number, dim: { start: number; end: number } | undefined): [number, number][] => {
    const out: [number, number][] = [];
    let from = 0;
    if (dimSkip && dim) {
      out.push([0, dim.start]);
      from = dim.end;
    }
    for (const row of rows) {
      if (!patchedRows.has(row.r)) continue;
      out.push([from, row.start]);
      from = row.end;
    }
    out.push([from, size]);
    return out;
  };
  const a = kept(baseRows, baseBuf.length, dimB);
  const b = kept(newRows, newBuf.length, dimN);
  // Compare the two span lists as one byte stream each.
  let ia = 0;
  let ib = 0;
  let pa = a[0][0];
  let pb = b[0][0];
  let compared = 0;
  for (;;) {
    while (ia < a.length && pa >= a[ia][1]) pa = ++ia < a.length ? a[ia][0] : 0;
    while (ib < b.length && pb >= b[ib][1]) pb = ++ib < b.length ? b[ib][0] : 0;
    if (ia >= a.length || ib >= b.length) {
      if (ia < a.length || ib < b.length) throw new SaveVerifyError('V2', `${where} changed outside the edited rows`);
      return;
    }
    const n = Math.min(a[ia][1] - pa, b[ib][1] - pb, 8 << 20);
    if (baseBuf.compare(newBuf, pb, pb + n, pa, pa + n) !== 0) throw new SaveVerifyError('V2', `${where} changed outside the edited rows`);
    pa += n;
    pb += n;
    compared += n;
    if (compared >= 8 << 20) {
      compared = 0;
      await saveYield(onYield);
    }
  }
}

/** The text of an <si> / <is> body: all <t> runs, without phonetic (<rPh>) text, unescaped and _xHHHH_-decoded. */
function xstringItemText(xml: string, P: string): string {
  const body = xml.replace(new RegExp(`<${regexEscape(P)}rPh\\b[\\s\\S]*?</${regexEscape(P)}rPh>`, 'g'), '');
  let text = '';
  for (const m of body.matchAll(new RegExp(`<${regexEscape(P)}t(?:\\s[^>]*)?(?:/>|>([\\s\\S]*?)</${regexEscape(P)}t>)`, 'g'))) {
    text += m[1] ?? '';
  }
  return decodeXstring(xmlUnescape(text));
}

/** One cell as found in the saved XML (V3). */
interface SavedCell {
  s: number;
  formula?: string;
  /** The <f> is a clone of a shared formula (<f t="shared" si="n"/>, no text of its own; kept byte for byte). */
  sharedClone?: boolean;
  /** undefined = no value. */
  value?: FileValue;
}

/** Reads the <c> element text of a saved cell (strings through `sstText`). */
function readSavedCell(xml: string, P: string, where: string, sstText: (i: number) => string): SavedCell {
  const gt = xml.indexOf('>');
  const attrs = parseXmlAttrs(xml.slice(0, gt + 1));
  const out: SavedCell = { s: Number(xmlAttrText(attrs, 's') ?? 0) };
  if (xml[gt - 1] === '/') return out;
  const children = splitXmlChildren(xml.slice(gt + 1, xml.length - `</${P}c>`.length), where);
  const t = xmlAttrText(attrs, 't') ?? 'n';
  const f = children.find((child) => child.local === 'f');
  if (f) {
    const fAttrs = parseXmlAttrs(f.open);
    out.formula = f.open.endsWith('/>') ? '' : xmlUnescape(f.xml.slice(f.open.length, f.xml.length - `</${P}f>`.length));
    if (xmlAttrText(fAttrs, 't') === 'shared' && !fAttrs.has('ref')) out.sharedClone = true;
  }
  const is = children.find((child) => child.local === 'is');
  const v = children.find((child) => child.local === 'v');
  if (t === 'inlineStr' && is) {
    out.value = { t: 's', v: xstringItemText(is.xml, P) };
  } else if (v) {
    const raw = v.open.endsWith('/>') ? '' : xmlUnescape(v.xml.slice(v.open.length, v.xml.length - `</${P}v>`.length));
    switch (t) {
      case 's':
        out.value = { t: 's', v: sstText(Number(raw)) };
        break;
      case 'str':
        out.value = { t: 's', v: decodeXstring(raw) };
        break;
      case 'b':
        out.value = { t: 'b', v: raw.trim() === '1' || raw.trim() === 'true' };
        break;
      case 'e':
        out.value = { t: 'e', v: raw };
        break;
      case 'n':
        out.value = { t: 'n', v: raw.trim() === '' ? NaN : Number(raw) };
        break;
      default:
        throw new SaveVerifyError('V3', `${where} has the unexpected cell type "${t}"`);
    }
  }
  return out;
}

function fileValueText(value: FileValue | null | undefined): string {
  return value === undefined ? 'nothing' : value === null ? 'no value' : `${value.t}:${JSON.stringify(value.v)}`;
}

function sameFileValue(a: FileValue | null | undefined, b: FileValue | null | undefined): boolean {
  if (!a || !b) return !a && !b;
  return a.t === b.t && (a.t === 'n' ? a.v === b.v || (Number.isNaN(a.v) && Number.isNaN(b.v as number)) : a.v === b.v);
}

/** Shared strings of the saved package, read once and indexed lazily (V3). */
function sharedStringReader(text: string | undefined, where: string): (i: number) => string {
  let items: string[] | undefined;
  let P = '';
  return (i: number): string => {
    if (text === undefined) throw new SaveVerifyError('V3', `${where}: a cell uses a shared string but the file has none`);
    if (!items) {
      P = xmlRootPrefix(text, 'sst', 'sharedStrings');
      items = [];
      const tag = `<${P}si`;
      const end = `</${P}si>`;
      for (let p = text.indexOf(tag); p >= 0; p = text.indexOf(tag, p + tag.length)) {
        if (!isXmlNameEnd(text.charCodeAt(p + tag.length))) continue;
        const gt = text.indexOf('>', p);
        const close = text[gt - 1] === '/' ? gt + 1 : text.indexOf(end, gt) + end.length;
        items.push(text.slice(p, close));
        p = close - tag.length;
      }
    }
    const item = Number.isInteger(i) ? items[i] : undefined;
    if (item === undefined) throw new SaveVerifyError('V3', `${where}: shared string ${i} does not exist`);
    return xstringItemText(item, P);
  };
}

/** V3: every expected cell is in the new sheet XML with the expected value / formula / style index. */
async function verifyReadBack(zip: ZipPackage, expect: readonly CellExpectation[], cache: ReadonlyMap<string, Buffer>, onYield?: SaveYieldHook): Promise<void> {
  if (!expect.length) return;
  let pkg: XlsxPackage;
  try {
    pkg = await openXlsxPackage(zip, onYield);
  } catch (err) {
    if (err instanceof PatchError || err instanceof ZipError) throw new SaveVerifyError('V3', saveErrorText(err));
    throw err;
  }
  let sstText: string | undefined;
  if (pkg.sstPath) sstText = cache.get(pkg.sstPath)?.toString('utf8') ?? (await readPartText(zip, pkg.sstPath, onYield));
  const sstAt = sharedStringReader(sstText, 'sharedStrings');
  const bySheet = new Map<number, CellExpectation[]>();
  for (const e of expect) {
    const list = bySheet.get(e.sheet);
    if (list) list.push(e);
    else bySheet.set(e.sheet, [e]);
  }
  for (const [sheetId, list] of bySheet) {
    let sheet: XlsxSheetRef & { part: string };
    let L: SheetXmlLayout;
    let buf: Buffer;
    const wanted = new Map<number, Map<number, string>>();
    for (const e of list) wanted.set(e.r, new Map());
    try {
      sheet = worksheetOf(pkg, sheetId);
      const entry = zip.get(sheet.part);
      if (!entry) throw new PatchError(`the part of sheet ${sheetId} is missing`);
      buf = cache.get(entry.name) ?? (await readEntry(zip, entry, onYield));
      L = await sheetLayout(buf, `the sheet "${sheetLabel(sheet)}"`, onYield);
      let last = -1;
      for (const r of wanted.keys()) last = Math.max(last, r);
      const step = saveSlicer(onYield, SAVE_YIELD_ROWS);
      for (let pos = L.sdOpenEnd; !L.sdSelfClosing;) {
        const rs = findStartTag(buf, L.rowTag, pos, L.sdClose);
        if (rs < 0) break;
        const row = readRowAt(buf, L, rs);
        const cells = wanted.get(row.r);
        if (cells && !row.selfClosing) {
          const text = decodeRowText(buf, row, L.where);
          const body = text.slice(row.tag.length, text.lastIndexOf(`</${L.P}row>`));
          for (const span of parseRowCells(body, L.P, row.r, L.where)) cells.set(span.c, body.slice(span.start, span.end));
        }
        if (row.r >= last) break;
        pos = row.end;
        const pause = step();
        if (pause) await pause;
      }
    } catch (err) {
      if (err instanceof PatchError || err instanceof ZipError) throw new SaveVerifyError('V3', saveErrorText(err));
      throw err;
    }
    for (const e of list) {
      const where = `${sheetLabel(sheet)}!${indexToCol(e.c)}${e.r + 1}`;
      const xml = wanted.get(e.r)?.get(e.c);
      const got: SavedCell = xml === undefined ? { s: 0 } : readSavedCell(xml, L.P, where, sstAt);
      if (e.styleIndex !== undefined && got.s !== e.styleIndex) {
        throw new SaveVerifyError('V3', `${where} has style ${got.s}, expected ${e.styleIndex}`);
      }
      if (e.value === undefined && e.formula === undefined) continue;
      // A shared-formula clone has no formula text of its own: only a cached-value patch can leave it in place (its <f>
      // bytes are kept); the expanded formula is compared by V4.
      const formulaOk = got.sharedClone ? e.formula !== undefined : (got.formula ?? undefined) === e.formula;
      if (!formulaOk) {
        throw new SaveVerifyError('V3', `${where} has the formula ${JSON.stringify(got.formula ?? null)}, expected ${JSON.stringify(e.formula ?? null)}`);
      }
      if (e.value !== undefined && !sameFileValue(got.value, e.value)) {
        throw new SaveVerifyError('V3', `${where} holds ${fileValueText(got.value ?? null)}, expected ${fileValueText(e.value)}`);
      }
    }
  }
}

/** V4: the new bytes load with the loader and every expected cell shows the expected value / formula (D-008). */
async function verifyDeep(bytes: Uint8Array, expect: readonly CellExpectation[]): Promise<void> {
  if (!expect.length) return;
  let model: WorkbookModel;
  try {
    model = workbookToModel(await loadWorkbook(bytes));
  } catch (err) {
    throw new SaveVerifyError('V4', `the saved file cannot be read back (${saveErrorText(err)})`);
  }
  const worksheets = model.workbook.worksheets;
  for (const e of expect) {
    const index = worksheets.findIndex((ws) => ws.id === e.sheet);
    const where = `${index >= 0 ? worksheets[index].name : `sheet ${e.sheet}`}!${indexToCol(e.c)}${e.r + 1}`;
    if (index < 0) throw new SaveVerifyError('V4', `${where}: the sheet is missing after reload`);
    if (e.value === undefined && e.formula === undefined) continue;
    const cell = model.getRows(index, e.r, e.r + 1, e.c, e.c).rows.find((row) => row.r === e.r)?.cells.find((x) => x.c === e.c);
    // Cases where the loader shows something else for correct XML (V3 still checks the XML): not comparable here.
    if (cell?.link !== undefined) continue; // hyperlink cells: the loader drops the formula / shows the link text
    if (model.getMeta().sheets[index]?.merges.some((m) => e.r >= m.r0 && e.r <= m.r1 && e.c >= m.c0 && e.c <= m.c1 && !(e.r === m.r0 && e.c === m.c0))) continue; // merged non-anchor cell
    if (e.value && e.value.t === 'n' && cell?.t === 'e' && cell.v === '#VALUE!') continue; // a huge number in a date format
    if (e.value && e.value.t === 's' && e.value.v === '') continue; // an empty string value reads back as an empty cell
    const formula = e.formula === undefined ? undefined : displayFormula(e.formula);
    if ((cell?.f || undefined) !== (formula || undefined)) {
      throw new SaveVerifyError('V4', `${where} shows the formula ${JSON.stringify(cell?.f ?? null)} after reload, expected ${JSON.stringify(formula ?? null)}`);
    }
    if (e.value === undefined) continue;
    const empty = !cell || cell.t === 'z' || (formula !== undefined && cell.t === 's' && cell.v === '');
    let same: boolean;
    if (e.value === null) {
      same = empty;
    } else if (e.value.t === 'n') {
      const v = cell?.v;
      same = typeof v === 'number' && (cell?.t === 'n' ? v === e.value.v : cell?.t === 'd' && Math.abs(v - e.value.v) <= 1e-8 * Math.max(1, Math.abs(v)));
    } else if (e.value.t === 's') {
      same = cell?.t === 's' && cell.v === e.value.v;
    } else {
      same = cell?.t === e.value.t && cell.v === e.value.v;
    }
    if (!same) {
      throw new SaveVerifyError('V4', `${where} shows ${cell ? `${cell.t}:${JSON.stringify(cell.v ?? null)}` : 'nothing'} after reload, expected ${fileValueText(e.value)}`);
    }
  }
}

/**
 * Checks the bytes buildXlsx made from `base` and `ops` before they are written (§4.2 step 4). Throws SaveVerifyError
 * (nothing may be written) when: V1 the zip does not parse, an entry is missing / added / reordered, an entry that
 * `ops` cannot change is not byte-identical, a rewritten entry does not inflate to its size + CRC, or a removed part
 * is still referenced; V2 a rewritten XML part is not well-formed (big cell-patched sheet parts: each patched row is
 * parsed and the rest must equal the base); V3 an expected cell is not in the new XML as expected; V4 (`o.deep`,
 * default: file ≤ 10 MB) the loader does not show the expected cells.
 */
export async function verifyXlsx(
  base: Uint8Array,
  bytes: Uint8Array,
  ops: readonly PackageOp[],
  expect: CellExpectation[],
  o?: { deep: boolean; onYield?: SaveYieldHook },
): Promise<void> {
  const onYield = o?.onYield;
  let baseZip: ZipPackage;
  let zip: ZipPackage;
  try {
    baseZip = parseZip(base);
  } catch (err) {
    throw new SaveVerifyError('V1', `the original file is not a valid zip package (${saveErrorText(err)})`);
  }
  try {
    zip = parseZip(bytes);
  } catch (err) {
    throw new SaveVerifyError('V1', `the new file is not a valid zip package (${saveErrorText(err)})`);
  }
  // ----- V1: entry set, order, raw bytes -----
  const allowedChanged = new Set<string>();
  const allowedRemoved = new Set<string>();
  const patchedRows = new Map<string, Set<number>>();
  let basePkg: XlsxPackage | undefined;
  if (ops.length) {
    try {
      basePkg = await openXlsxPackage(baseZip, onYield);
    } catch (err) {
      if (err instanceof PatchError || err instanceof ZipError) throw new SaveVerifyError('V1', saveErrorText(err));
      throw err;
    }
    for (const op of ops) {
      if (op.kind !== 'cells') throw new SaveVerifyError('V1', `"${op.kind}" changes cannot be checked yet`);
      let part: string;
      try {
        part = worksheetOf(basePkg, op.sheet).part.toLowerCase();
      } catch (err) {
        throw new SaveVerifyError('V1', saveErrorText(err));
      }
      allowedChanged.add(part);
      let rows = patchedRows.get(part);
      if (!rows) patchedRows.set(part, (rows = new Set()));
      for (const [r] of op.cells) rows.add(r);
    }
    for (const path of [basePkg.sstPath, basePkg.stylesPath, basePkg.wbPath, basePkg.wbRelsPath, '[Content_Types].xml']) {
      if (path) allowedChanged.add(path.toLowerCase());
    }
    for (const path of basePkg.calcChainPaths) allowedRemoved.add(path.toLowerCase());
  }
  if (baseZip.comment.compare(zip.comment) !== 0) throw new SaveVerifyError('V1', 'the zip comment changed');
  if (baseZip.buf.compare(zip.buf, 0, zip.prefixEnd, 0, baseZip.prefixEnd) !== 0) {
    throw new SaveVerifyError('V1', 'the bytes before the first part changed');
  }
  const rewritten = new Map<string, Buffer>();
  const removed: string[] = [];
  for (const e of baseZip.entries) {
    const n = zip.get(e.name);
    if (!n) {
      if (!allowedRemoved.has(e.name.toLowerCase())) throw new SaveVerifyError('V1', `the part ${e.name} is missing`);
      removed.push(e.name);
      continue;
    }
    const same = n.method === e.method && n.crc === e.crc && n.usize === e.usize && n.csize === e.csize &&
      n.spanEnd - n.lho === e.spanEnd - e.lho && baseZip.buf.compare(zip.buf, n.lho, n.spanEnd, e.lho, e.spanEnd) === 0 &&
      n.nameBytes.equals(e.nameBytes);
    if (same) continue;
    if (!allowedChanged.has(e.name.toLowerCase())) throw new SaveVerifyError('V1', `the part ${e.name} changed but must stay byte-identical`);
    try {
      rewritten.set(n.name, await readEntry(zip, n, onYield));
    } catch (err) {
      if (err instanceof ZipError) throw new SaveVerifyError('V1', saveErrorText(err));
      throw err; // cancelled (onYield)
    }
  }
  const expectedOrder = baseZip.entries.filter((e) => zip.get(e.name)).map((e) => e.name);
  if (zip.entries.length !== expectedOrder.length) throw new SaveVerifyError('V1', 'the new file has parts the original does not have');
  if (zip.entries.some((e, i) => e.name !== expectedOrder[i])) throw new SaveVerifyError('V1', 'the order of the zip directory changed');
  const expectedLocal = baseZip.byOffset.filter((e) => zip.get(e.name)).map((e) => e.name);
  if (zip.byOffset.some((e, i) => e.name !== expectedLocal[i])) throw new SaveVerifyError('V1', 'the order of the parts changed');
  if (removed.length && basePkg) {
    // A removed part must not be referenced any more (content type override, workbook relationship).
    const ct = rewritten.get('[Content_Types].xml')?.toString('utf8') ?? (await readPartText(zip, '[Content_Types].xml', onYield)) ?? '';
    const relsEntry = zip.get(basePkg.wbRelsPath);
    const rels = (relsEntry && (rewritten.get(relsEntry.name)?.toString('utf8') ?? (await readPartText(zip, relsEntry.name, onYield)))) ?? '';
    const gone = new Set(removed.map((name) => name.toLowerCase()));
    for (const m of ct.matchAll(/<(?:[\w.-]+:)?Override\b[^>]*>/g)) {
      const name = xmlAttrText(parseXmlAttrs(m[0]), 'PartName');
      if (name && gone.has(resolvePartTarget('', name).toLowerCase())) throw new SaveVerifyError('V1', `the removed part ${name} still has a content type`);
    }
    for (const rel of parseRels(rels)) {
      if (!rel.external && gone.has(resolvePartTarget(basePkg.wbPath, rel.target).toLowerCase())) {
        throw new SaveVerifyError('V1', `the removed part ${rel.target} is still referenced by the workbook`);
      }
    }
  }
  // ----- V2: rewritten XML parts are well-formed -----
  for (const [name, data] of rewritten) {
    if (!/\.(xml|rels)$/i.test(name)) continue;
    const rows = patchedRows.get(name.toLowerCase());
    if (data.length > verifyTuning.fullParseBytes && rows) {
      const baseEntry = baseZip.get(name);
      if (!baseEntry) throw new SaveVerifyError('V2', `${name} has no original`);
      await verifySheetRows(await readEntry(baseZip, baseEntry, onYield), data, rows, name, onYield);
    } else {
      await verifyXmlWellFormed(data, name, onYield);
    }
  }
  // ----- V3: read back the edited cells -----
  await verifyReadBack(zip, expect, rewritten, onYield);
  // ----- V4: reload with the loader (small files) -----
  if (o?.deep ?? deepVerifyAllowed(bytes)) await verifyDeep(bytes, expect);
}

// ----- patchable check -----

/** Hard read-only scan of one worksheet part (§4.7): rows with numbers and in order, cells inside their row and in order. */
async function scanSheetPart(buf: Buffer, where: string, yieldEvery: number, onYield?: SaveYieldHook): Promise<void> {
  const L = await sheetLayout(buf, where, onYield);
  if (L.sdSelfClosing) return;
  const gapOk = (from: number, to: number): boolean => {
    if (to <= from) return true;
    return /^(?:\s|<!--[\s\S]*?-->)*$/.test(buf.toString('latin1', from, to));
  };
  let pos = L.sdOpenEnd;
  let prevR = -1;
  const step = saveSlicer(onYield, yieldEvery);
  for (;;) {
    const rs = findStartTag(buf, L.rowTag, pos, L.sdClose);
    if (!gapOk(pos, rs < 0 ? L.sdClose : rs)) throw new PatchError(`${where} has cells outside a row`);
    if (rs < 0) break;
    const row = readRowAt(buf, L, rs);
    if (row.r <= prevR) throw new PatchError(`${where} has rows out of order (row ${row.r + 1})`);
    prevR = row.r;
    if (!row.selfClosing) {
      const inner = buf.toString('latin1', row.openEnd, row.end - L.rowEndTag.length);
      // The patcher walks bytes, not an XML tree: a comment or CDATA section inside a row could hide a fake <c> or </row>.
      if (inner.includes('<!--') || inner.includes('<![CDATA[')) throw new PatchError(`${where} has an XML comment or CDATA section inside a row (row ${row.r + 1})`);
      parseRowCells(inner, L.P, row.r, where);
    }
    pos = row.end;
    const pause = step();
    if (pause) await pause;
  }
}

/**
 * Whether the save path can patch this package (§4.7 hard read-only; reason = plain text for "FileStudio cannot
 * edit this file safely: <reason>. It is open read-only."): ZIP64, encrypted entries, compression other than
 * store/deflate, duplicate part names, damaged zip, Strict Open XML, parts that are not UTF-8, and — unless
 * `scanSheets` is false — every worksheet part is scanned for rows without numbers, rows out of order and cells
 * outside their row or out of order. Cooperative (for the background check after load); `onYield` may throw to cancel.
 */
export async function checkPatchable(
  base: Uint8Array,
  o?: { scanSheets?: boolean; yieldEvery?: number; onYield?: SaveYieldHook },
): Promise<{ ok: true } | { ok: false; reason: string }> {
  const state = { cancelled: false };
  const userYield = o?.onYield;
  const onYield: SaveYieldHook | undefined = userYield
    ? (): void => {
      try {
        userYield();
      } catch (err) {
        state.cancelled = true;
        throw err;
      }
    }
    : undefined;
  const yieldEvery = Math.max(1, Math.floor(o?.yieldEvery ?? SAVE_YIELD_ROWS));
  try {
    const zip = parseZip(base);
    const pkg = await openXlsxPackage(zip, onYield);
    if (pkg.strict) return { ok: false, reason: 'it uses Strict Open XML' };
    for (const path of [pkg.sstPath, pkg.stylesPath]) {
      const entry = path ? zip.get(path) : undefined;
      if (entry) {
        const text = partText(await readEntry(zip, entry, onYield), entry.name);
        if (path === pkg.sstPath && (text.includes('<!--') || text.includes('<![CDATA['))) {
          return { ok: false, reason: 'its shared strings contain an XML comment or CDATA section' };
        }
      }
    }
    if (o?.scanSheets !== false) {
      for (const sheet of pkg.sheets) {
        if (!WORKSHEET_REL_RE.test(sheet.type)) continue;
        const entry = sheet.part ? zip.get(sheet.part) : undefined;
        if (!entry) return { ok: false, reason: `the part of the sheet "${sheetLabel(sheet)}" is missing` };
        await scanSheetPart(await readEntry(zip, entry, onYield), `the sheet "${sheetLabel(sheet)}"`, yieldEvery, onYield);
        await saveYield(onYield);
      }
    }
    return { ok: true };
  } catch (err) {
    if (state.cancelled) throw err;
    if (err instanceof ZipError || err instanceof PatchError) return { ok: false, reason: err.message };
    return { ok: false, reason: `it could not be checked (${saveErrorText(err)})` };
  }
}
