// Custom editor providers for FileStudio.
//
//   SheetEditorProvider  - .xlsx  (CustomEditorProvider, ExcelJS workbook + edit stack)
//   TextViewerProvider   - .csv/.tsv/.psv/.ssv/.md (CustomTextEditorProvider, TextDocument is the source of truth)
//   DocxViewerProvider   - .docx  (CustomReadonlyEditorProvider)
//   PdfViewerProvider    - .pdf   (CustomReadonlyEditorProvider; pdf.js renders in the webview)
//   PptxViewerProvider   - .pptx  (CustomReadonlyEditorProvider; slides are sent one at a time)
//
// All of them share getHtml() (strict CSP + nonce) and the message protocol below.

import type { CellStyle, Range, RowData, SelectionStats, WorkbookMeta } from './renderers/sheet';
import type { TaskMarker, TocEntry } from './renderers/markdown';
import type { PptxDeckMeta, PptxSlide } from './renderers/pptx';
import type { BorderSide, EditIntent, EditKind, FindQuery, SheetId, SheetMeta, StyleDelta } from './renderers/sheet';
import type { MdBlock, MdTextEdit } from './renderers/markdown';

import { execFile } from 'child_process';
import { randomBytes } from 'crypto';
import { constants as fsConstants } from 'fs';
import { access, open as openFileHandle, stat as statPath } from 'fs/promises';
import * as path from 'path';
import * as vscode from 'vscode';
import {
  CsvGridModel,
  detectLossyFeatures,
  detectSsvDelimiter,
  loadWorkbook,
  parseCsv,
  workbookToModel,
  type CsvModel,
  type GridSource,
  type WorkbookModel,
} from './renderers/sheet';
import { markdownHeadings, renderMarkdown, type RenderResult } from './renderers/markdown';
import { renderDocx, type DocxResult } from './renderers/docx';
import { checkPdf, PDF_HEADER_SEARCH_BYTES } from './renderers/pdf';
import { renderPptx, type PptxDocument as PptxDeck } from './renderers/pptx';

// ===== MESSAGE PROTOCOL (shared with media/viewer.js — keep in sync) =====

export type ViewKind = 'sheet' | 'csv' | 'markdown' | 'docx' | 'pdf' | 'pptx';

/** Delimited-text flavour of a 'csv' view, chosen by file extension (.csv auto-detects its delimiter). */
export type DelimitedFormat = 'csv' | 'tsv' | 'psv' | 'ssv';
export type MarkdownMode = 'preview' | 'split' | 'wysiwyg';

// ----- editing (v0.2.0) -----
// Declared once by track F (contract §2.7, D-062); later phases add handlers only. The host is the authority: every
// `edit` is checked again on the host (SHARED: EDIT GATE) and answered with exactly one `editResult`.

/** Why an edit was not applied (or why a document is not editable). Texts: EDIT_TEXT (SHARED: EDIT GATE). */
export type EditRejectCode = 'stale' | 'readOnly' | 'disabled' | 'lossy' | 'fidelity' | 'blocked' | 'invalid' | 'failed' | 'internal';

/** Whether a document can be edited now; sent in `init` / `invalidate` / `markdownUpdate` and as `editState`. */
export interface EditState {
  editable: boolean;
  /** Why not editable. */
  code?: EditRejectCode;
  /** User text, shown in the banner / status line. */
  reason?: string;
  /** lossy (xlsx) / fidelity (csv, Markdown WYSIWYG): the host offers "Edit anyway" (a modal it shows itself). */
  canEditAnyway?: { title: string; details: string[] };
  /** Intents this document supports now (the webview shows only these actions). */
  kinds?: EditKind[];
}

/** Replaces rows [start, end) of a sheet in the webview's row cache. */
export interface RowsPatch {
  sheetId: SheetId;
  start: number;
  end: number;
  rows: RowData[];
  styles: [number, CellStyle][];
}

/** Every modal question the host asks (`ask()`); integration tests answer them through the test API. */
export type AskId =
  | 'editAnyway'
  | 'bakFailed'
  | 'diskChangedOnSave'
  | 'mergeClears'
  | 'sheetDelete'
  | 'fidelity'
  | 'ownerFile'
  | 'inPlaceNoBackup';

/** Extension -> webview. */
export type HostMessage =
  | {
      type: 'init';
      kind: 'sheet' | 'csv';
      fileName: string;
      meta: WorkbookMeta;
      readOnly: boolean;
      /** Banner text shown above the grid (e.g. lossy features), if any. */
      banner?: string;
      /** kind 'csv' only: which delimited format and the delimiter actually used (' ' = runs of spaces). */
      delimited?: { format: DelimitedFormat; delimiter: string };
      /** (v0.2.0) Host version of the document (one counter per document; +1 on every content change). */
      version?: number;
      /** (v0.2.0) Whether the grid may be edited; `readOnly` stays `!editState.editable`. */
      editState?: EditState;
    }
  | {
      type: 'init';
      kind: 'markdown';
      fileName: string;
      html: string;
      toc: TocEntry[];
      source: string;
      mode: MarkdownMode;
      /** TextDocument.version the html was rendered from. */
      version: number;
      /**
       * VS Code treats the document as read-only (a git: version, files.readonlyInclude, files.readonlyFromPermissions
       * with a read-only file): task checkboxes are disabled, and a toggleTask is refused. Also in markdownUpdate, which
       * follows a change of this state while the file is open (a settings change).
       */
      readOnly?: boolean;
      /** (v0.2.0) Whether FileStudio may change the document (task toggles, WYSIWYG); `readOnly` = `!editState.editable`. */
      editState?: EditState;
      /** (v0.2.0, P6) Top-level blocks of `source` for the WYSIWYG editor. */
      blocks?: MdBlock[];
      /** (v0.2.0, P6) Line break of the document. */
      eol?: '\n' | '\r\n';
      /** (v0.2.0, P6) Whether WYSIWYG can be used for this document (and why not). */
      limits?: { wysiwyg: boolean; reason?: string };
      /** (v0.2.0, P6) Webview URI base for relative resources in WYSIWYG mode. */
      resourceBase?: string;
    }
  | {
      type: 'init';
      kind: 'docx';
      fileName: string;
      html: string;
      /** Conversion warnings from mammoth (shown collapsed). */
      warnings: string[];
    }
  | {
      /**
       * PDF view. `source.uri` is a webview URI of the file (pdf.js range-loads it; preferred for file: documents);
       * `source.data` carries the bytes when the document cannot be served as a webview resource (git:, virtual FS).
       */
      type: 'init';
      kind: 'pdf';
      fileName: string;
      source: { uri: string } | { data: Uint8Array };
      fileSize: number;
      banner?: string;
    }
  | {
      /** PPTX view: deck outline only; slides are fetched with getSlide. */
      type: 'init';
      kind: 'pptx';
      fileName: string;
      deck: PptxDeckMeta;
      banner?: string;
    }
  | {
      /** Reply to getSlide (`reqId` / `index` echo the request). */
      type: 'slide';
      reqId: number;
      index: number;
      slide: PptxSlide;
    }
  | {
      /**
       * Reply to getRows. `styles` carries style-table entries first used by these rows. With `c0` / `c1` (echoed
       * from the request) the rows hold only the cells of those columns: not for the grid's row cache.
       */
      type: 'rows';
      reqId: number;
      sheet: number;
      start: number;
      end: number;
      c0?: number;
      c1?: number;
      rows: RowData[];
      styles: [number, CellStyle][];
      /** (v0.2.0) Document version the rows were read from (the webview drops older ones). */
      version?: number;
      /** (v0.2.0) Sheet the rows belong to (the webview drops rows of a sheet it did not ask for). */
      sheetId?: SheetId;
    }
  /**
   * Reply to getStats. Besides the numbers, `stats.errors` counts error cells (then the status bar shows only Count,
   * like Excel) and `stats.text` has Sum / Average / Min / Max in the number format of the first numeric cell (dates,
   * percentages, currency; absent for General). See SelectionStats in renderers/sheet.ts.
   */
  | { type: 'stats'; reqId: number; stats: SelectionStats }
  | {
      /** Data changed (CSV text edited elsewhere, file reloaded): drop the row cache, re-request visible rows. */
      type: 'invalidate';
      meta?: WorkbookMeta;
      /**
       * xlsx reload: the lossy-content banner of the reloaded file (null = none), so it follows content added or
       * removed on disk (a chart). Absent = keep the banner as it is (CSV).
       */
      banner?: string | null;
      /** (v0.2.0) Host version of the document after the change. */
      version?: number;
      /** (v0.2.0) Edit state after the change. */
      editState?: EditState;
    }
  /**
   * `readOnly`: as in the markdown init (sent again with every update, so a change of the state reaches the view).
   * (v0.2.0) `editState`, `blocks`, `eol`: as in the markdown init; `html` may be '' while the panel is in WYSIWYG.
   */
  | {
      type: 'markdownUpdate';
      html: string;
      toc: TocEntry[];
      source: string;
      version: number;
      readOnly?: boolean;
      editState?: EditState;
      blocks?: MdBlock[];
      eol?: '\n' | '\r\n';
    }
  /**
   * A link to the document itself was followed (`readme.md#usage` in readme.md, `file:///C:/x/self.docx#intro` in
   * self.docx): scroll the markdown / docx view to `fragment` (decoded; '' = the top), like an in-page `#usage` link.
   */
  | { type: 'scrollToFragment'; fragment: string }
  | {
      /**
       * Answer to every toggleTask (`line` / `version` echo the request). `applied`: the document now has the requested
       * state (the edit was made, or it already had it); its rendering (markdownUpdate) is sent or follows. Otherwise
       * the toggle was refused (`reason`, e.g. the document changed since the click, or it is read-only) and the
       * checkbox should show the document's state again. `documentVersion`: TextDocument.version after handling the toggle.
       */
      type: 'toggleTaskResult';
      line: number;
      version: number;
      applied: boolean;
      documentVersion: number;
      reason?: string;
    }
  | {
      type: 'error';
      message: string;
      detail?: string;
      /** csv/md: offer "Reopen as Text". xlsx/docx/pdf/pptx: offer "Reveal in Explorer" instead. */
      canReopenAsText: boolean;
    }
  // ---------- (v0.2.0) editing: declared by track F, handled by P1-P6 ----------
  /** Answer to one `edit` (exactly one per edit): applied, with the new version and the changed rows. */
  | { type: 'editResult'; editId: number; accepted: true; version: number; patches?: RowsPatch[]; select?: { sheetId: SheetId; range: Range } }
  /** Answer to one `edit`: not applied; the webview rolls back its overlay and shows `reason` in the status line. */
  | { type: 'editResult'; editId: number; accepted: false; code: EditRejectCode; reason: string; version: number; where?: string[] }
  /**
   * Content changed by another panel, undo/redo or recalc. `sheets` replaces those sheets' meta (by id) without
   * dropping the row cache.
   */
  | { type: 'patch'; version: number; patches: RowsPatch[]; sheets?: SheetMeta[]; styles?: [number, CellStyle][] }
  | { type: 'editState'; state: EditState }
  /** P5: recalculation progress. */
  | { type: 'recalcState'; busy: boolean; done?: number; total?: number; circular?: string }
  /** P3/P4: answer to `find`. */
  | {
      type: 'findResult';
      reqId: number;
      version: number;
      matches: { sheetId: SheetId; r: number; c: number }[];
      total: number;
      truncated: boolean;
      error?: string;
    }
  /** P4: answer to `filter`. */
  | {
      type: 'filterResult';
      reqId: number;
      version: number;
      sheetId: SheetId;
      values?: { text: string; count: number }[];
      hidden?: [start: number, end: number][];
      truncated?: boolean;
      /** The request could not be read (a read is never left unanswered). */
      error?: string;
    }
  /** P4 auto-fit: answer to `measure` (the longest texts of a row / column). */
  | { type: 'columnTexts'; reqId: number; version: number; sheetId: SheetId; axis: 'row' | 'col'; index: number; texts: { text: string; s?: number }[]; error?: string }
  /** P6 split sync: scroll the preview to this (fractional) source line. */
  | { type: 'revealLine'; line: number }
  /** P6 WYSIWYG: send pending typing now (answered by `edit` messages, then `flushed`). */
  | { type: 'flushRequest'; id: number };

/** Webview -> extension. */
export type WebviewMessage =
  | { type: 'ready' }
  /**
   * Rows [start, end) of a sheet. Optional `c0` / `c1`: only the cells of columns c0..c1 (inclusive), e.g. to copy
   * a few columns of a large sheet without converting every cell of its rows.
   */
  | { type: 'getRows'; reqId: number; sheet: number; start: number; end: number; c0?: number; c1?: number; sheetId?: SheetId }
  /** (v0.2.0) `sheetId` (getRows too): the host resolves the sheet by id when given; `sheet` stays for 0.1.2. */
  | { type: 'getStats'; reqId: number; sheet: number; ranges: Range[]; sheetId?: SheetId }
  /** External URL (http/https/mailto), a path relative to the document, or a workspace file. */
  | { type: 'openLink'; href: string }
  /**
   * Markdown task checkbox toggled; `line` is the 0-based source line from data-line, `version` the TextDocument.version
   * of the rendering it was clicked in. Always answered with toggleTaskResult.
   */
  | { type: 'toggleTask'; line: number; checked: boolean; version: number }
  /** PPTX: one slide of the deck (0-based). Answered with 'slide' (or 'error' for a bad index / parse failure). */
  | { type: 'getSlide'; reqId: number; index: number }
  | { type: 'reopenAsText' }
  | { type: 'revealInExplorer' }
  | { type: 'log'; level: 'info' | 'warn' | 'error'; message: string }
  // ---------- (v0.2.0) editing: declared by track F, handled by P1-P6 ----------
  /**
   * One user action = one edit = one VS Code undo step. Always answered with exactly one `editResult` (same `editId`).
   * `then: 'save'`: save the document after the edit was applied (Ctrl+S while the cell editor is open).
   */
  | { type: 'edit'; editId: number; baseVersion: number; intent: EditIntent; then?: 'save' }
  /** "Edit anyway" was clicked: the host shows its own modal (lossy xlsx / fidelity csv, md); this is never consent. */
  | { type: 'editAnyway' }
  /** P3/P4: find cells (answered with `findResult`). */
  | { type: 'find'; reqId: number; query: FindQuery; from?: { sheetId: SheetId; r: number; c: number }; direction: 'next' | 'prev' | 'all' }
  /** P4: list a column's values (`keep` absent), filter to `keep`, or clear the filter (`keep` null). */
  | { type: 'filter'; reqId: number; sheetId: SheetId; col: number; keep?: string[] | null }
  /** P4 auto-fit: the longest texts of a row / column (answered with `columnTexts`). */
  | { type: 'measure'; reqId: number; sheetId: SheetId; axis: 'row' | 'col'; index: number }
  /** P6 WYSIWYG; P2 grid while an edit is in flight (D-067). The host runs it after the queued edits. */
  | { type: 'command'; id: 'undo' | 'redo' | 'save' }
  /** P6: switch the Markdown mode of this panel. */
  | { type: 'setMode'; mode: MarkdownMode }
  /** P6 split sync: the preview scrolled to this (fractional) source line. */
  | { type: 'previewScrolled'; line: number }
  /** P6 WYSIWYG: answer to `flushRequest`, sent after the flush's edit message. */
  | { type: 'flushed'; id: number };

// ===== IMPLEMENTATION =====

// ===== SHARED: VIEW TYPES & COMMANDS =====

/** viewType of every custom editor contributed in package.json. */
export const VIEW_TYPES = {
  sheet: 'fileStudio.sheet',
  csv: 'fileStudio.csv',
  markdown: 'fileStudio.markdown',
  docx: 'fileStudio.docx',
  pdf: 'fileStudio.pdf',
  pptx: 'fileStudio.pptx',
} as const satisfies Record<ViewKind, string>;

export type ViewType = (typeof VIEW_TYPES)[ViewKind];

const KIND_BY_EXTENSION: ReadonlyMap<string, ViewKind> = new Map<string, ViewKind>([
  ['.xlsx', 'sheet'],
  ['.csv', 'csv'],
  ['.tsv', 'csv'],
  ['.psv', 'csv'],
  ['.ssv', 'csv'],
  ['.md', 'markdown'],
  ['.markdown', 'markdown'],
  ['.docx', 'docx'],
  ['.pdf', 'pdf'],
  ['.pptx', 'pptx'],
  ['.pptm', 'pptx'],
  ['.ppsx', 'pptx'],
  ['.ppsm', 'pptx'],
  ['.potx', 'pptx'],
  ['.potm', 'pptx'],
]);

/** File extension of a resource, lower-case with the dot ('' when it has none). */
function extensionOf(uri: vscode.Uri): string {
  return path.posix.extname(uri.path).toLowerCase();
}

/** Which viewer handles a file, judged by its extension (case-insensitive). */
export function viewKindForUri(uri: vscode.Uri): ViewKind | undefined {
  return KIND_BY_EXTENSION.get(extensionOf(uri));
}

/** Custom editor viewType for a file (used by the `fileStudio.openWith` command), or undefined if unsupported. */
export function viewTypeForUri(uri: vscode.Uri): ViewType | undefined {
  const kind = viewKindForUri(uri);
  return kind === undefined ? undefined : VIEW_TYPES[kind];
}

/** Display name of a resource (last path segment). */
export function fileNameOf(uri: vscode.Uri): string {
  return path.posix.basename(uri.path) || uri.toString(true);
}

function isTextViewerTab(tab: vscode.Tab, uri: vscode.Uri): boolean {
  const input = tab.input;
  return (
    input instanceof vscode.TabInputCustom &&
    (input.viewType === VIEW_TYPES.csv || input.viewType === VIEW_TYPES.markdown) &&
    input.uri.toString() === uri.toString()
  );
}

function tabGroupFor(viewColumn: vscode.ViewColumn | undefined): vscode.TabGroup {
  const groups = vscode.window.tabGroups;
  return (viewColumn !== undefined && groups.all.find((g) => g.viewColumn === viewColumn)) || groups.activeTabGroup;
}

/**
 * Replaces the FileStudio editor of a CSV/TSV/PSV/SSV/Markdown file with VS Code's text editor (in the same editor
 * group). The TextDocument is shared, so unsaved changes carry over. Binary formats (xlsx/docx/pdf/pptx) are refused.
 */
export async function reopenAsText(uri: vscode.Uri, viewColumn?: vscode.ViewColumn): Promise<void> {
  const kind = viewKindForUri(uri);
  if (kind === 'sheet' || kind === 'docx' || kind === 'pdf' || kind === 'pptx') {
    void vscode.window.showInformationMessage(
      `"${fileNameOf(uri)}" is a binary file. Reopen as Text is available for CSV, TSV, PSV, SSV and Markdown files.`,
    );
    return;
  }
  const column = tabGroupFor(viewColumn).viewColumn;
  await vscode.commands.executeCommand('vscode.openWith', uri, 'default', column);
  // vscode.openWith adds a second tab next to ours; close ours so this behaves like "Reopen Editor With".
  const stale = tabGroupFor(column).tabs.filter((tab) => isTextViewerTab(tab, uri));
  if (stale.length > 0) {
    try {
      await vscode.window.tabGroups.close(stale, true);
    } catch (err) {
      logError(`Could not close the FileStudio tab of ${uri.toString(true)}`, err);
    }
  }
}

/**
 * Reveals a document in the OS file manager (falls back to the VS Code explorer for non-file schemes). A Git version
 * of a file (`git:` URI, e.g. the left side of "Open Changes") reveals its working file.
 */
export async function revealInOS(uri: vscode.Uri): Promise<void> {
  const target = gitVersionOf(uri)?.workingFile ?? uri;
  try {
    await vscode.commands.executeCommand(target.scheme === 'file' ? 'revealFileInOS' : 'revealInExplorer', target);
  } catch (err) {
    logError(`Could not reveal ${target.toString(true)}`, err);
  }
}

// ===== SHARED: HTML SHELL & CSP =====

/** Random nonce for the script tag (base64 of 16 random bytes). */
export function createNonce(): string {
  return randomBytes(16).toString('base64');
}

/**
 * The webview Content-Security-Policy. 'unsafe-inline' styles are required by KaTeX, mermaid and inline cell styles.
 * Remote (https:) images and media are allowed in the markdown preview only (badges, hosted pictures); workbook,
 * Word and slide content is embedded as data: URIs. The PDF view also runs the pdf.js worker, which the webview
 * creates from a blob: URL of the worker script shipped in dist/ (`worker-src blob:`; the file itself is fetched
 * through connect-src). Frames, form submission and <base> are refused by the policy itself, not only by the
 * sanitizer and the webview's event handlers.
 */
export function buildCsp(cspSource: string, nonce: string, kind: ViewKind): string {
  const remote = kind === 'markdown' ? ' https:' : '';
  const workers = kind === 'pdf' ? 'worker-src blob:; ' : '';
  return (
    `default-src 'none'; img-src ${cspSource} data: blob:${remote}; media-src ${cspSource} data:${remote}; ` +
    `style-src ${cspSource} 'unsafe-inline'; font-src ${cspSource} data:; script-src 'nonce-${nonce}' ${cspSource}; ` +
    `connect-src ${cspSource}; ${workers}frame-src 'none'; child-src 'none'; form-action 'none'; base-uri 'none';`
  );
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Minimal escaping for a double-quoted attribute value (keeps the CSP's single quotes readable). */
function escapeAttribute(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

/** Folders every webview may load from: the bundled viewer (dist/) and our stylesheet (media/). */
function baseResourceRoots(extensionUri: vscode.Uri): vscode.Uri[] {
  return [vscode.Uri.joinPath(extensionUri, 'dist'), vscode.Uri.joinPath(extensionUri, 'media')];
}

function webviewOptions(localResourceRoots: vscode.Uri[]): vscode.WebviewOptions {
  return { enableScripts: true, enableCommandUris: false, localResourceRoots };
}

/**
 * The static page every viewer starts from. `#app` shows a loading message until media/viewer.js takes over; if the
 * script never runs, a hint appears after a few seconds (pure CSS) so the panel is never silently blank. The hint is
 * visibility-hidden until then, so screen readers do not announce the failure text on every open.
 */
export function getHtml(webview: vscode.Webview, extensionUri: vscode.Uri, kind: ViewKind, fileName = ''): string {
  const nonce = createNonce();
  const csp = buildCsp(webview.cspSource, nonce, kind);
  const asset = (...segments: string[]): string =>
    escapeAttribute(webview.asWebviewUri(vscode.Uri.joinPath(extensionUri, ...segments)).toString());
  const name = escapeHtml(fileName || 'file');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${escapeAttribute(csp)}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${name}</title>
<link rel="stylesheet" href="${asset('dist', 'viewer.css')}">
<link rel="stylesheet" href="${asset('media', 'viewer.css')}">
<style>
  .fv-boot { padding: 24px; font-family: var(--vscode-font-family); color: var(--vscode-descriptionForeground); }
  .fv-boot-hint { visibility: hidden; animation: fv-boot-reveal 0s linear 10s forwards; }
  @keyframes fv-boot-reveal { to { visibility: visible; } }
</style>
</head>
<body data-kind="${kind}">
<div id="app">
  <div class="fv-boot" role="status" aria-live="polite">
    <p class="fv-boot-message">Loading ${name}…</p>
    <p class="fv-boot-hint">Still loading. Large files can take a while; if nothing appears, the viewer could not start. See the "FileStudio" output channel (View &gt; Output) for details, or reopen the file with another editor.</p>
  </div>
</div>
<noscript><p class="fv-boot">FileStudio needs JavaScript to display ${name}.</p></noscript>
<script type="module" nonce="${nonce}" src="${asset('dist', 'viewer.js')}"></script>
</body>
</html>`;
}

// ===== SHARED: MESSAGING & LINKS =====

/** If the webview has not sent 'ready' this long after becoming visible, the viewer script most likely failed. */
const READY_TIMEOUT_MS = 10_000;
/** Index bound for rows/columns/sheets in requests (CSV files may exceed Excel's 1,048,576 rows). */
const MAX_INDEX = 0x7fffffff;
const MAX_ROWS_PER_REQUEST = 10_000;
const MAX_STATS_RANGES = 1_024;
const MAX_HREF_LENGTH = 8_192;
const MAX_LOG_LENGTH = 10_000;

type MessageOf<T extends WebviewMessage['type']> = Extract<WebviewMessage, { type: T }>;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** Integer in [min, max] (truncated and clamped), or undefined when `value` is not a finite number. */
function toInt(value: unknown, min: number, max: number): number | undefined {
  if (typeof value !== 'number' || !Number.isFinite(value)) return undefined;
  return Math.min(max, Math.max(min, Math.trunc(value)));
}

function toRange(value: unknown): Range | undefined {
  if (!isRecord(value)) return undefined;
  const r0 = toInt(value.r0, 0, MAX_INDEX);
  const c0 = toInt(value.c0, 0, MAX_INDEX);
  const r1 = toInt(value.r1, 0, MAX_INDEX);
  const c1 = toInt(value.c1, 0, MAX_INDEX);
  if (r0 === undefined || c0 === undefined || r1 === undefined || c1 === undefined) return undefined;
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

// ----- edit messages (v0.2.0, contract §2.7) -----
// Shape and limits of every editing message. Unlike the read requests above, numbers in an `edit` are never clamped
// or truncated (row -1 must not become row 0): a value out of range makes the whole edit `invalid`. The meaning is
// checked again by the phase's planner (planEdit / planCsvEdit / Markdown checks).

/** Limits of `parseWebviewMessage` for editing messages (contract §2.7). */
export const EDIT_LIMITS = {
  /** Cells in one setCells. */
  cells: 1_000_000,
  /** Characters of one cell value: xlsx (Excel's limit) and csv. */
  xlsxValueChars: 32_767,
  csvValueChars: 1024 * 1024,
  /** All text of one edit together. */
  totalChars: 16 * 1024 * 1024,
  ranges: 1_000,
  /** Rows / columns inserted or deleted at once (Excel's row count). */
  count: 1_048_576,
  /** FindQuery.text and the replace text. */
  findChars: 1_000,
  /** Markdown text edits in one mdText, and their characters (oldText + text). */
  mdEdits: 256,
  mdChars: 4 * 1024 * 1024,
  sheetNameChars: 31,
  /** OOXML number format codes (Excel's limit). */
  numFmtChars: 255,
  /** Values in one filter `keep` list. */
  filterValues: 1_000_000,
  /** Largest row height / column width in px. */
  sizePx: 65_535,
} as const;

/** Excel's last row / column (0-based): a bound for xlsx (kind 'sheet') positions. */
const XLSX_MAX_ROW = 1_048_575;
const XLSX_MAX_COL = 16_383;
/** xlsx: 's' + the <sheet sheetId> attribute; csv: 'csv' (SheetId, contract §2.1). */
const SHEET_ID_RE = /^(?:s[0-9]{1,10}|csv)$/;
const ARGB_RE = /^[0-9A-F]{8}$/i;
/** C0 and C1 control characters and DEL (not allowed in sheet names). */
const CONTROL_CHAR_RE = /[\u0000-\u001F\u007F-\u009F\uFFFE\uFFFF]/;
/** C0 controls and U+FFFE / U+FFFF: not allowed in XML text (numFmt is written into styles.xml). */
const XML_INVALID_RE = /[\u0000-\u001F\uFFFE\uFFFF]/;
/** A UTF-16 surrogate without its partner: UTF-8 would silently store U+FFFD instead (D-071). */
const LONE_SURROGATE_RE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const BORDER_STYLES = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'] as const;

/** A refused editing message: `message` is the user text of the `editResult` (code 'invalid'). */
class EditShapeError extends Error {}

function badEdit(reason: string = EDIT_TEXT.invalid): never {
  throw new EditShapeError(reason);
}

/** An integer in [0, max] exactly as sent (no truncation, no clamping), or undefined. */
function toIndex(value: unknown, max: number): number | undefined {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 && value <= max ? value : undefined;
}

function toSheetId(value: unknown): SheetId | undefined {
  return typeof value === 'string' && SHEET_ID_RE.test(value) ? value : undefined;
}

function editIndex(value: unknown, max: number): number {
  return toIndex(value, max) ?? badEdit();
}

function editBool(value: unknown): boolean {
  return typeof value === 'boolean' ? value : badEdit();
}

/** A sheet id; when the view kind is known, 'csv' for a CSV view and s<n> for a workbook view (nothing else). */
function editSheetId(value: unknown, kind?: ViewKind): SheetId {
  const id = toSheetId(value) ?? badEdit();
  if (kind === 'csv' ? id !== 'csv' : kind === 'sheet' && id === 'csv') badEdit();
  return id;
}

function editEnum<T extends string>(value: unknown, values: readonly T[]): T {
  return typeof value === 'string' && (values as readonly string[]).includes(value) ? (value as T) : badEdit();
}

function editArgb(value: unknown): string {
  return typeof value === 'string' && ARGB_RE.test(value) ? value : badEdit();
}

/** Text budget of one message (EDIT_LIMITS.totalChars) and whether its text goes into an xlsx file. */
interface EditTextBudget {
  chars: number;
  readonly xlsx: boolean;
}

/** A string of at most `max` characters, counted against the message budget; text written to a file has no lone surrogate (it would become U+FFFD). */
function editText(value: unknown, max: number, budget: EditTextBudget, tooLong: string): string {
  if (typeof value !== 'string') badEdit();
  if (value.length > max) badEdit(EDIT_TEXT.tooLarge(tooLong));
  budget.chars += value.length;
  if (budget.chars > EDIT_LIMITS.totalChars) badEdit(EDIT_TEXT.tooLarge('16 million characters of text'));
  if (LONE_SURROGATE_RE.test(value)) badEdit(EDIT_TEXT.loneSurrogate);
  return value;
}

/** Sheet name: 1-31 characters, no control characters (Excel's other name rules are the planner's). */
function editSheetName(value: unknown, budget: EditTextBudget): string {
  const name = editText(value, EDIT_LIMITS.sheetNameChars, budget, '31 characters in a sheet name');
  if (name.length === 0 || CONTROL_CHAR_RE.test(name)) badEdit(EDIT_TEXT.badSheetName);
  return name;
}

/** Row/column bounds of the document kind: Excel's grid for xlsx, MAX_INDEX otherwise (CSV may have more rows). */
function gridBounds(kind: ViewKind | undefined): { maxRow: number; maxCol: number } {
  return kind === 'sheet' ? { maxRow: XLSX_MAX_ROW, maxCol: XLSX_MAX_COL } : { maxRow: MAX_INDEX, maxCol: MAX_INDEX };
}

function editRange(value: unknown, kind: ViewKind | undefined): Range {
  if (!isRecord(value)) badEdit();
  const { maxRow, maxCol } = gridBounds(kind);
  const r0 = editIndex(value.r0, maxRow);
  const c0 = editIndex(value.c0, maxCol);
  const r1 = editIndex(value.r1, maxRow);
  const c1 = editIndex(value.c1, maxCol);
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

function editRanges(value: unknown, kind: ViewKind | undefined): Range[] {
  if (!Array.isArray(value) || value.length === 0) badEdit();
  if (value.length > EDIT_LIMITS.ranges) badEdit(EDIT_TEXT.tooLarge('1,000 ranges'));
  return Array.from(value, (item) => editRange(item, kind)); // (Array.from visits holes; map would skip them)
}

/**
 * FindQuery shape. `compileRegex`: a regular expression must compile (with the 'u' flag, else without it) — for replaceAll; a `find`
 * with a bad pattern is passed on, so its handler can answer `findResult.error` (a read is never left unanswered).
 */
function editFindQuery(value: unknown, budget: EditTextBudget, compileRegex: boolean, kind?: ViewKind): FindQuery {
  if (!isRecord(value)) badEdit();
  const scope = value.scope === 'workbook' ? 'workbook' : editSheetId(value.scope, kind);
  const text = editText(value.text, EDIT_LIMITS.findChars, budget, '1,000 characters to find');
  const query: FindQuery = {
    scope,
    text,
    matchCase: editBool(value.matchCase),
    regex: editBool(value.regex),
    wholeCell: editBool(value.wholeCell),
    lookIn: editEnum(value.lookIn, ['values', 'formulas'] as const),
  };
  if (compileRegex && query.regex) {
    try {
      try {
        new RegExp(text, 'u');
      } catch {
        new RegExp(text); // ordinary patterns (-, x{, a]) are valid only without 'u'; the matcher uses the flags that compile
      }
    } catch {
      badEdit(EDIT_TEXT.badPattern);
    }
  }
  return query;
}

function editBorderSide(value: unknown): BorderSide {
  if (value === null) return null;
  if (!isRecord(value)) badEdit();
  const side: NonNullable<BorderSide> = { style: editEnum(value.style, BORDER_STYLES) };
  if (value.color !== undefined) side.color = editArgb(value.color);
  return side;
}

/** StyleDelta with only the known properties (unknown ones are dropped), each of the right type. */
function editStyleDelta(value: unknown, budget: EditTextBudget): StyleDelta {
  if (!isRecord(value)) badEdit();
  const delta: StyleDelta = {};
  for (const key of ['bold', 'italic', 'strike', 'wrap', 'quotePrefix'] as const) {
    if (value[key] !== undefined) delta[key] = editBool(value[key]);
  }
  if (value.underline !== undefined) delta.underline = value.underline === null ? null : editEnum(value.underline, ['single', 'double'] as const);
  if (value.fontColor !== undefined) delta.fontColor = value.fontColor === null ? null : editArgb(value.fontColor);
  if (value.fill !== undefined) delta.fill = value.fill === null ? null : editArgb(value.fill);
  if (value.hAlign !== undefined) {
    delta.hAlign = value.hAlign === null ? null : editEnum(value.hAlign, ['left', 'center', 'right', 'justify'] as const);
  }
  if (value.vAlign !== undefined) delta.vAlign = value.vAlign === null ? null : editEnum(value.vAlign, ['top', 'center', 'bottom'] as const);
  if (value.numFmt !== undefined) {
    delta.numFmt = editText(value.numFmt, EDIT_LIMITS.numFmtChars, budget, '255 characters in a number format');
    if (XML_INVALID_RE.test(delta.numFmt)) badEdit();
  }
  if (value.border !== undefined) {
    if (!isRecord(value.border)) badEdit();
    const border: NonNullable<StyleDelta['border']> = {};
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      if (value.border[side] !== undefined) border[side] = editBorderSide(value.border[side]);
    }
    delta.border = border;
  }
  return delta;
}

/** mdText edits: sorted by offset, not overlapping, `oldText` as long as its range, at most 256 edits / 4 Mi chars. */
function editMdTextEdits(value: unknown, budget: EditTextBudget): MdTextEdit[] {
  if (!Array.isArray(value) || value.length === 0) badEdit();
  if (value.length > EDIT_LIMITS.mdEdits) badEdit(EDIT_TEXT.tooLarge('256 text changes at once'));
  const edits: MdTextEdit[] = [];
  let chars = 0;
  for (const item of value) {
    if (!isRecord(item)) badEdit();
    const start = editIndex(item.start, MAX_INDEX);
    const end = editIndex(item.end, MAX_INDEX);
    if (end < start || typeof item.oldText !== 'string' || typeof item.text !== 'string') badEdit();
    if (item.oldText.length !== end - start) badEdit();
    if (LONE_SURROGATE_RE.test(item.text)) badEdit(EDIT_TEXT.loneSurrogate);
    chars += item.oldText.length + item.text.length;
    if (chars > EDIT_LIMITS.mdChars) badEdit(EDIT_TEXT.tooLarge('4 million characters'));
    const prev = edits[edits.length - 1];
    // Sorted and not overlapping; two insertions at the same offset have no defined order.
    if (prev && (start < prev.end || (start === prev.end && prev.start === prev.end && start === end))) badEdit();
    budget.chars += item.oldText.length + item.text.length;
    if (budget.chars > EDIT_LIMITS.totalChars) badEdit(EDIT_TEXT.tooLarge('16 million characters of text'));
    edits.push({ start, end, oldText: item.oldText, text: item.text });
  }
  return edits;
}

/** Intent kinds each view can send (a message for another document kind is `invalid`). Views not listed: none. */
const SHEET_INTENTS: ReadonlySet<string> = new Set([
  'setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll', 'style', 'merge', 'unmerge',
  'resize', 'sheetAdd', 'sheetRename', 'sheetDelete', 'sheetDuplicate', 'sheetMove', 'sheetHide', 'tabColor',
]);
const CSV_INTENTS: ReadonlySet<string> = new Set(['setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll']);
const MD_INTENTS: ReadonlySet<string> = new Set(['mdText']);
function intentAllowedIn(kind: ViewKind | undefined, intent: string): boolean {
  if (kind === undefined) return true; // shape check only
  const allowed = kind === 'sheet' ? SHEET_INTENTS : kind === 'csv' ? CSV_INTENTS : kind === 'markdown' ? MD_INTENTS : undefined;
  return allowed?.has(intent) ?? false;
}

/** One EditIntent (contract §2.2) with every §2.7 limit; throws EditShapeError. */
function editIntentOf(value: unknown, kind: ViewKind | undefined, budget: EditTextBudget): EditIntent {
  if (!isRecord(value) || typeof value.kind !== 'string') badEdit();
  if (!intentAllowedIn(kind, value.kind)) badEdit();
  const { maxRow, maxCol } = gridBounds(kind);
  switch (value.kind) {
    case 'setCells': {
      const sheetId = editSheetId(value.sheetId, kind);
      if (!Array.isArray(value.cells) || value.cells.length === 0) badEdit();
      if (value.cells.length > EDIT_LIMITS.cells) badEdit(EDIT_TEXT.tooLarge('1,000,000 cells'));
      const [maxChars, tooLong] = budget.xlsx
        ? [EDIT_LIMITS.xlsxValueChars, '32,767 characters in one cell']
        : [EDIT_LIMITS.csvValueChars, '1,048,576 characters in one cell'];
      const cells: [r: number, c: number, text: string][] = new Array(value.cells.length);
      for (let i = 0; i < value.cells.length; i++) {
        const cell: unknown = value.cells[i];
        if (!Array.isArray(cell) || cell.length !== 3) badEdit();
        cells[i] = [editIndex(cell[0], maxRow), editIndex(cell[1], maxCol), editText(cell[2], maxChars, budget, tooLong)];
      }
      return { kind: 'setCells', sheetId, cells };
    }
    case 'clear': {
      const sheetId = editSheetId(value.sheetId, kind);
      const ranges = editRanges(value.ranges, kind);
      const what = editEnum(value.what, ['contents', 'formats', 'all'] as const);
      if (kind === 'csv' && what !== 'contents') badEdit(); // a CSV cell has no format
      return { kind: 'clear', sheetId, ranges, what };
    }
    case 'insertRows':
    case 'deleteRows':
    case 'insertCols':
    case 'deleteCols': {
      const count = editIndex(value.count, kind === 'csv' ? MAX_INDEX : EDIT_LIMITS.count); // (a CSV may have more rows than Excel)
      if (count === 0) badEdit();
      return { kind: value.kind, sheetId: editSheetId(value.sheetId, kind), at: editIndex(value.at, kind === 'sheet' ? (value.kind.endsWith('Rows') ? maxRow : maxCol) + 1 : MAX_INDEX), count };
    }
    case 'sort':
      return {
        kind: 'sort',
        sheetId: editSheetId(value.sheetId, kind),
        range: editRange(value.range, kind),
        key: editIndex(value.key, maxCol),
        desc: editBool(value.desc),
        header: editBool(value.header),
      };
    case 'replaceAll': {
      const query = editFindQuery(value.query, budget, true, kind);
      return { kind: 'replaceAll', query, replace: editText(value.replace, EDIT_LIMITS.findChars, budget, '1,000 characters to replace with') };
    }
    case 'style':
      return {
        kind: 'style',
        sheetId: editSheetId(value.sheetId, kind),
        ranges: editRanges(value.ranges, kind),
        delta: editStyleDelta(value.delta, budget),
      };
    case 'merge':
    case 'unmerge':
      return { kind: value.kind, sheetId: editSheetId(value.sheetId, kind), range: editRange(value.range, kind) };
    case 'resize': {
      const sheetId = editSheetId(value.sheetId, kind);
      const axis = editEnum(value.axis, ['row', 'col'] as const);
      if (!Array.isArray(value.index) || value.index.length === 0) badEdit();
      if (value.index.length > EDIT_LIMITS.count) badEdit(EDIT_TEXT.tooLarge('1,048,576 rows or columns'));
      const max = axis === 'row' ? maxRow : maxCol;
      const index = Array.from(value.index as unknown[], (item) => editIndex(item, max));
      let size: number | null = null;
      if (value.size !== null) {
        if (typeof value.size !== 'number' || !Number.isFinite(value.size) || value.size < 0 || value.size > EDIT_LIMITS.sizePx) badEdit();
        size = value.size;
      }
      return { kind: 'resize', sheetId, axis, index, size };
    }
    case 'sheetAdd': {
      const intent: Extract<EditIntent, { kind: 'sheetAdd' }> = { kind: 'sheetAdd', index: editIndex(value.index, MAX_INDEX) };
      if (value.name !== undefined) intent.name = editSheetName(value.name, budget);
      return intent;
    }
    case 'sheetRename':
      return { kind: 'sheetRename', sheetId: editSheetId(value.sheetId, kind), name: editSheetName(value.name, budget) };
    case 'sheetDelete':
    case 'sheetDuplicate':
      return { kind: value.kind, sheetId: editSheetId(value.sheetId, kind) };
    case 'sheetMove':
      return { kind: 'sheetMove', sheetId: editSheetId(value.sheetId, kind), index: editIndex(value.index, MAX_INDEX) };
    case 'sheetHide':
      return { kind: 'sheetHide', sheetId: editSheetId(value.sheetId, kind), hidden: editBool(value.hidden) };
    case 'tabColor':
      return { kind: 'tabColor', sheetId: editSheetId(value.sheetId, kind), argb: value.argb === null ? null : editArgb(value.argb) };
    case 'mdText':
      return { kind: 'mdText', edits: editMdTextEdits(value.edits, budget) };
    default:
      return badEdit(); // unknown kind
  }
}

/** parseEditMessage: the edit, or why it was refused (`editId` when the message names one; it is answered `invalid`). */
export type EditParseResult = { ok: true; msg: MessageOf<'edit'> } | { ok: false; reason: string; editId?: number };

/**
 * Validates an `edit` message (shape + every §2.7 limit). `kind` is the view the message came from: xlsx ('sheet')
 * positions are bounded by Excel's grid and xlsx text may not contain a lone surrogate. Never throws.
 */
export function parseEditMessage(raw: unknown, kind?: ViewKind): EditParseResult {
  const editId = isRecord(raw) ? toIndex(raw.editId, Number.MAX_SAFE_INTEGER) : undefined;
  try {
    if (!isRecord(raw) || raw.type !== 'edit' || editId === undefined) return badEdit();
    const baseVersion = editIndex(raw.baseVersion, Number.MAX_SAFE_INTEGER);
    if (raw.then !== undefined && raw.then !== 'save') badEdit();
    const intent = editIntentOf(raw.intent, kind, { chars: 0, xlsx: kind === 'sheet' });
    const msg: MessageOf<'edit'> = { type: 'edit', editId, baseVersion, intent };
    if (raw.then === 'save') msg.then = 'save';
    return { ok: true, msg };
  } catch (err) {
    if (err instanceof EditShapeError) return { ok: false, reason: err.message, editId };
    // A bug in the validation itself (not a bad message): refuse the edit, never let it through.
    return { ok: false, reason: EDIT_TEXT.internal, editId };
  }
}

function parseFindMessage(raw: Record<string, unknown>): MessageOf<'find'> | undefined {
  const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
  if (reqId === undefined) return undefined;
  try {
    const query = editFindQuery(raw.query, { chars: 0, xlsx: false }, false);
    const direction = editEnum(raw.direction, ['next', 'prev', 'all'] as const);
    const msg: MessageOf<'find'> = { type: 'find', reqId, query, direction };
    if (raw.from !== undefined) {
      if (!isRecord(raw.from)) return undefined;
      msg.from = { sheetId: editSheetId(raw.from.sheetId), r: editIndex(raw.from.r, MAX_INDEX), c: editIndex(raw.from.c, MAX_INDEX) };
    }
    return msg;
  } catch {
    return undefined;
  }
}

function parseFilterMessage(raw: Record<string, unknown>): MessageOf<'filter'> | undefined {
  const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
  const sheetId = toSheetId(raw.sheetId);
  const col = toIndex(raw.col, MAX_INDEX);
  if (reqId === undefined || sheetId === undefined || col === undefined) return undefined;
  const msg: MessageOf<'filter'> = { type: 'filter', reqId, sheetId, col };
  if (raw.keep === null) msg.keep = null;
  else if (raw.keep !== undefined) {
    if (!Array.isArray(raw.keep) || raw.keep.length > EDIT_LIMITS.filterValues) return undefined;
    let chars = 0;
    for (const item of raw.keep) {
      if (typeof item !== 'string') return undefined;
      chars += item.length;
      if (chars > EDIT_LIMITS.totalChars) return undefined;
    }
    msg.keep = raw.keep.slice() as string[];
  }
  return msg;
}

/**
 * Validates a message posted by the webview. Unknown types and malformed payloads yield undefined; numbers are
 * truncated and clamped, row windows and range lists are bounded, strings are length-limited. (v0.2.0) Editing
 * messages are validated exactly (contract §2.7, EDIT_LIMITS); `kind` is the view the message came from (xlsx
 * limits apply to 'sheet'). A refused `edit` is answered `invalid` by the session (parseEditMessage tells why).
 */
export function parseWebviewMessage(raw: unknown, kind?: ViewKind): WebviewMessage | undefined {
  if (!isRecord(raw) || typeof raw.type !== 'string') return undefined;
  switch (raw.type) {
    case 'ready':
    case 'reopenAsText':
    case 'revealInExplorer':
      return { type: raw.type };
    case 'getRows': {
      const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
      const sheet = toInt(raw.sheet, 0, MAX_INDEX);
      const start = toInt(raw.start, 0, MAX_INDEX);
      if (reqId === undefined || sheet === undefined || start === undefined) return undefined;
      const end = Math.min(Math.max(toInt(raw.end, 0, MAX_INDEX) ?? start, start), start + MAX_ROWS_PER_REQUEST);
      const msg: MessageOf<'getRows'> = { type: 'getRows', reqId, sheet, start, end };
      // Optional column window: both bounds or none (a malformed window falls back to whole rows).
      const c0 = toInt(raw.c0, 0, MAX_INDEX);
      const c1 = toInt(raw.c1, 0, MAX_INDEX);
      if (c0 !== undefined && c1 !== undefined) {
        msg.c0 = Math.min(c0, c1);
        msg.c1 = Math.max(c0, c1);
      }
      if (raw.sheetId !== undefined) {
        const sheetId = toSheetId(raw.sheetId);
        if (sheetId === undefined) return undefined; // a malformed id: dropped (the session logs it)
        msg.sheetId = sheetId;
      }
      return msg;
    }
    case 'getStats': {
      const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
      const sheet = toInt(raw.sheet, 0, MAX_INDEX);
      if (reqId === undefined || sheet === undefined || !Array.isArray(raw.ranges)) return undefined;
      const ranges: Range[] = [];
      for (const item of raw.ranges.slice(0, MAX_STATS_RANGES)) {
        const range = toRange(item);
        if (range) ranges.push(range);
      }
      const msg: MessageOf<'getStats'> = { type: 'getStats', reqId, sheet, ranges };
      if (raw.sheetId !== undefined) {
        const sheetId = toSheetId(raw.sheetId);
        if (sheetId === undefined) return undefined;
        msg.sheetId = sheetId;
      }
      return msg;
    }
    case 'openLink':
      return typeof raw.href === 'string' && raw.href.length > 0 && raw.href.length <= MAX_HREF_LENGTH
        ? { type: 'openLink', href: raw.href }
        : undefined;
    case 'toggleTask': {
      // Exact integers only (a message that changes a file is never clamped or truncated).
      const line = toIndex(raw.line, MAX_INDEX);
      const version = toIndex(raw.version, Number.MAX_SAFE_INTEGER);
      if (line === undefined || version === undefined || typeof raw.checked !== 'boolean') return undefined;
      return { type: 'toggleTask', line, checked: raw.checked, version };
    }
    case 'getSlide': {
      // Like row indexes: truncated and clamped to >= 0. An index past the deck is answered by the provider.
      const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
      const index = toInt(raw.index, 0, MAX_INDEX);
      if (reqId === undefined || index === undefined) return undefined;
      return { type: 'getSlide', reqId, index };
    }
    case 'log': {
      if (typeof raw.message !== 'string') return undefined;
      const level = raw.level === 'warn' || raw.level === 'error' ? raw.level : 'info';
      return { type: 'log', level, message: raw.message.slice(0, MAX_LOG_LENGTH) };
    }
    // ----- (v0.2.0) editing messages: declared by track F (contract §2.7); handlers come with P1-P6 -----
    case 'edit': {
      const parsed = parseEditMessage(raw, kind);
      return parsed.ok ? parsed.msg : undefined;
    }
    case 'editAnyway':
      return { type: 'editAnyway' };
    case 'find':
      return parseFindMessage(raw);
    case 'filter':
      return parseFilterMessage(raw);
    case 'measure': {
      const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
      const sheetId = toSheetId(raw.sheetId);
      const index = toIndex(raw.index, MAX_INDEX);
      if (reqId === undefined || sheetId === undefined || index === undefined) return undefined;
      return raw.axis === 'row' || raw.axis === 'col' ? { type: 'measure', reqId, sheetId, axis: raw.axis, index } : undefined;
    }
    case 'command':
      return raw.id === 'undo' || raw.id === 'redo' || raw.id === 'save' ? { type: 'command', id: raw.id } : undefined;
    case 'setMode':
      return raw.mode === 'preview' || raw.mode === 'split' || raw.mode === 'wysiwyg' ? { type: 'setMode', mode: raw.mode } : undefined;
    case 'previewScrolled':
      // A fractional source line (the scroll position inside a block).
      return typeof raw.line === 'number' && Number.isFinite(raw.line) && raw.line >= 0
        ? { type: 'previewScrolled', line: Math.min(raw.line, MAX_INDEX) }
        : undefined;
    case 'flushed': {
      const id = toIndex(raw.id, Number.MAX_SAFE_INTEGER);
      return id === undefined ? undefined : { type: 'flushed', id };
    }
    default:
      return undefined;
  }
}

/** One error line for users: the error message without the stack. */
function errorMessageOf(err: unknown): string {
  if (err instanceof Error) return err.message || err.name;
  return typeof err === 'string' ? err : String(err);
}

/**
 * Full error text for the collapsible detail / logs (stack when available), followed by its causes: the renderers
 * throw readable messages and keep the library's technical error as `cause`.
 */
function errorDetailOf(err: unknown): string {
  let detail = '';
  let current: unknown = err;
  for (let depth = 0; depth < 4 && current !== undefined; depth++) {
    const text = current instanceof Error ? current.stack || `${current.name}: ${current.message}` : errorMessageOf(current);
    detail += depth === 0 ? text : `\nCaused by: ${text}`;
    current = current instanceof Error ? current.cause : undefined;
  }
  return detail;
}

/**
 * (v0.2.0) Optional handlers of a view for the editing messages of contract §2.7. A view sets the ones its phase
 * implements (`session.editHandlers = { onEdit: … }`, a one-line hook); a missing handler means "not available yet":
 * an `edit` is answered `blocked` "This action is not available yet.", any other message is logged and dropped.
 * `onEdit` must answer with exactly one `editResult` before the promise it returns settles (the session answers
 * `internal` for an edit that is still unanswered then, and drops a second answer).
 */
export interface ViewEditHandlers {
  /** The document's current host version (used when the session itself rejects an edit). Default 0. */
  version?(): number;
  onEdit?(msg: MessageOf<'edit'>): unknown;
  onEditAnyway?(msg: MessageOf<'editAnyway'>): unknown;
  onFind?(msg: MessageOf<'find'>): unknown;
  onFilter?(msg: MessageOf<'filter'>): unknown;
  onMeasure?(msg: MessageOf<'measure'>): unknown;
  onCommand?(msg: MessageOf<'command'>): unknown;
  onSetMode?(msg: MessageOf<'setMode'>): unknown;
  onPreviewScrolled?(msg: MessageOf<'previewScrolled'>): unknown;
  onFlushed?(msg: MessageOf<'flushed'>): unknown;
}

/** Every open session (the test API finds a document's first panel here). */
const liveSessions = new Set<ViewerSession>();

/**
 * One webview panel showing one document. Validates incoming messages, tracks the ready handshake (and logs when it
 * never happens), remembers what the webview currently shows and ties listeners to the panel's lifetime.
 */
class ViewerSession implements vscode.Disposable {
  readonly fileName: string;
  /** What the webview displays: nothing yet (loading), a successful init, or the error view. */
  shown: 'nothing' | 'content' | 'error' = 'nothing';
  /** (v0.2.0) The view's handlers for the editing messages (see ViewEditHandlers). */
  editHandlers: ViewEditHandlers = {};
  /**
   * (v0.2.0) Edits received and not answered yet (each gets exactly one editResult). `epoch` = the webview life
   * (counted up by every 'ready'): a reloaded webview restarts its editIds, so an answer from an older life is dropped
   * and a new edit with the same id waits until the old one is closed. `closed` settles when the entry is removed.
   */
  private readonly openEdits = new Map<number, { epoch: number; closed: Promise<void>; close: () => void }>();
  private epoch = 0;
  /** (v0.2.0, test API) Lists that collect every message posted while a test message is handled. */
  private readonly captures = new Set<HostMessage[]>();

  private ready = false;
  private disposed = false;
  private readyTimer: ReturnType<typeof setTimeout> | undefined;
  private waitingForVisibility = false;
  private readonly disposables: vscode.Disposable[] = [];

  constructor(
    readonly panel: vscode.WebviewPanel,
    readonly uri: vscode.Uri,
    readonly kind: ViewKind,
    private readonly handler: (msg: WebviewMessage) => unknown,
  ) {
    this.fileName = fileNameOf(uri);
    this.disposables.push(
      panel.webview.onDidReceiveMessage((raw: unknown) => this.receive(raw)),
      panel.onDidChangeViewState(() => {
        if (this.waitingForVisibility && panel.visible && !this.ready) {
          this.waitingForVisibility = false;
          this.armReadyTimer();
        }
      }),
      panel.onDidDispose(() => this.dispose()),
    );
    liveSessions.add(this);
    this.armReadyTimer();
  }

  get webview(): vscode.Webview {
    return this.panel.webview;
  }

  /** True once the webview script has installed its listeners and reported 'ready'. */
  get isReady(): boolean {
    return this.ready && !this.disposed;
  }

  get isDisposed(): boolean {
    return this.disposed;
  }

  /** Disposes `disposable` together with the panel. */
  onDispose(disposable: vscode.Disposable): void {
    if (this.disposed) disposable.dispose();
    else this.disposables.push(disposable);
  }

  post(msg: HostMessage): void {
    if (this.disposed) return;
    if (msg.type === 'editResult' && !this.settleEdit(msg)) return;
    for (const capture of this.captures) capture.push(msg);
    if (msg.type === 'init') this.shown = 'content';
    else if (msg.type === 'error') this.shown = 'error';
    try {
      Promise.resolve(this.panel.webview.postMessage(msg)).then(
        (delivered) => {
          if (!delivered && !this.disposed) getLog().debug(`[${this.kind}] '${msg.type}' not delivered to ${this.fileName}`);
        },
        (err: unknown) => logError(`[${this.kind}] Could not post '${msg.type}' to ${this.fileName}`, err),
      );
    } catch (err) {
      logError(`[${this.kind}] Could not post '${msg.type}' to ${this.fileName}`, err);
    }
  }

  /**
   * Shows the error view (with "Reopen as Text" for csv/md, "Reveal in Explorer" for xlsx/docx/pdf/pptx) and logs
   * `err` with its detail. `alreadyLogged`: `err` is a load / parse failure that its owner logged when it happened, so
   * showing it (again, e.g. after a webview reload or for a row request) adds no second error entry for one failure.
   */
  postError(summary: string, err?: unknown, alreadyLogged = false): void {
    const where = `[${this.kind}] ${summary} (${this.uri.toString(true)})`;
    if (err !== undefined && alreadyLogged) getLog().debug(`${where}: shown in the viewer (logged when it occurred)`);
    else if (err !== undefined) logError(where, err);
    else getLog().error(where);
    this.post({
      type: 'error',
      message: err === undefined ? summary : `${summary}: ${errorMessageOf(err)}`,
      detail: err === undefined ? undefined : errorDetailOf(err),
      canReopenAsText: this.kind === 'csv' || this.kind === 'markdown',
    });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    liveSessions.delete(this);
    this.clearReadyTimer();
    for (const disposable of this.disposables.splice(0)) {
      try {
        disposable.dispose();
      } catch (err) {
        logError(`[${this.kind}] Error while disposing the viewer of ${this.fileName}`, err);
      }
    }
  }

  /**
   * (v0.2.0, test API) Handles `raw` as if the webview had posted it; resolves with the messages posted to this panel
   * while it was handled.
   */
  async deliverForTest(raw: unknown): Promise<HostMessage[]> {
    const captured: HostMessage[] = [];
    this.captures.add(captured);
    try {
      await this.receive(raw);
    } finally {
      this.captures.delete(captured);
    }
    return captured;
  }

  /** Answers an edit with a rejection (the status line shows `reason`; never the error view, contract §0 rule 6). */
  rejectEdit(editId: number, code: EditRejectCode, reason: string, where?: string[]): void {
    const version = this.editHandlers.version?.() ?? 0;
    this.post({ type: 'editResult', editId, accepted: false, code, reason, version, ...(where ? { where } : {}) });
  }

  private receive(raw: unknown): Promise<void> {
    if (this.disposed) return Promise.resolve();
    let msg: WebviewMessage | undefined;
    if (isRecord(raw) && raw.type === 'edit') {
      const openId = toIndex(raw.editId, Number.MAX_SAFE_INTEGER);
      const open = openId === undefined ? undefined : this.openEdits.get(openId);
      if (open) {
        if (open.epoch !== this.epoch) return open.closed.then(() => this.receive(raw)); // the old webview's edit
        // A buggy webview reused an id: the new message must not take the first edit's one answer.
        getLog().error(`[${this.kind}] Ignored an 'edit' with editId ${openId} that is still open from the viewer of ${this.fileName}`);
        return Promise.resolve();
      }
      // An edit is never dropped silently: a malformed one is answered 'invalid' (the webview rolls back its overlay).
      const parsed = parseEditMessage(raw, this.kind);
      if (!parsed.ok) {
        if (parsed.editId === undefined) {
          getLog().info(`[${this.kind}] Ignored an 'edit' message without a valid editId from the viewer of ${this.fileName}`);
        } else {
          this.openEdit(parsed.editId);
          this.rejectEdit(parsed.editId, 'invalid', parsed.reason);
        }
        return Promise.resolve();
      }
      msg = parsed.msg;
    } else {
      msg = parseWebviewMessage(raw, this.kind);
    }
    if (!msg) {
      this.answerBadRead(raw);
      const type = isRecord(raw) && typeof raw.type === 'string' ? `'${raw.type.slice(0, 64)}'` : 'malformed';
      getLog().debug(`[${this.kind}] Ignored ${type} message from the viewer of ${this.fileName}`);
      return Promise.resolve();
    }
    if (msg.type === 'ready') {
      // A (re)loaded webview starts from the static loading page again.
      this.ready = true;
      this.epoch++;
      this.shown = 'nothing';
      this.waitingForVisibility = false;
      this.clearReadyTimer();
    }
    const message = msg;
    const epoch = this.epoch;
    if (message.type === 'edit') this.openEdit(message.editId);
    return Promise.resolve()
      .then(() => this.route(message))
      .catch((err: unknown) => {
        if (this.disposed) return;
        if (message.type === 'ready') this.postError(`Could not open "${this.fileName}"`, err);
        else logError(`[${this.kind}] Handling '${message.type}' failed for ${this.fileName}`, err);
      })
      .then(() => {
        // Every edit gets exactly one answer, also when its handler failed or did not answer.
        if (message.type !== 'edit' || this.disposed) return;
        const open = this.openEdits.get(message.editId);
        if (!open || open.epoch !== epoch) return;
        if (epoch !== this.epoch) {
          this.closeEdit(message.editId); // the webview that sent it is gone: nobody to tell
          return;
        }
        getLog().error(`[${this.kind}] Edit ${message.editId} of ${this.fileName} was not answered by its handler`);
        this.rejectEdit(message.editId, 'internal', EDIT_TEXT.internal);
      });
  }

  /** (v0.2.0) Editing messages go to the view's optional handlers (editHandlers); everything else to `handler`. */
  private route(msg: WebviewMessage): unknown {
    const h = this.editHandlers;
    switch (msg.type) {
      case 'edit':
        if (h.onEdit) return h.onEdit(msg);
        this.rejectEdit(msg.editId, 'blocked', EDIT_TEXT.notAvailable);
        return undefined;
      case 'editAnyway':
        return h.onEditAnyway ? h.onEditAnyway(msg) : this.notAvailable(msg);
      case 'find':
        return h.onFind ? h.onFind(msg) : this.notAvailable(msg);
      case 'filter':
        return h.onFilter ? h.onFilter(msg) : this.notAvailable(msg);
      case 'measure':
        return h.onMeasure ? h.onMeasure(msg) : this.notAvailable(msg);
      case 'command':
        return h.onCommand ? h.onCommand(msg) : this.notAvailable(msg);
      case 'setMode':
        return h.onSetMode ? h.onSetMode(msg) : this.notAvailable(msg);
      case 'previewScrolled':
        return h.onPreviewScrolled ? h.onPreviewScrolled(msg) : this.notAvailable(msg);
      case 'flushed':
        return h.onFlushed ? h.onFlushed(msg) : this.notAvailable(msg);
      default:
        return this.handler(msg);
    }
  }

  /** A find / filter / measure that cannot be read but names its reqId is answered with an error: the webview is waiting for it. */
  private answerBadRead(raw: unknown): void {
    if (!isRecord(raw)) return;
    const reqId = toInt(raw.reqId, 0, Number.MAX_SAFE_INTEGER);
    if (reqId === undefined) return;
    const version = this.editHandlers.version?.() ?? 0;
    const error = EDIT_TEXT.badRequest;
    const sheetId = toSheetId(raw.sheetId) ?? '';
    if (raw.type === 'find') this.post({ type: 'findResult', reqId, version, matches: [], total: 0, truncated: false, error });
    else if (raw.type === 'filter') this.post({ type: 'filterResult', reqId, version, sheetId, error });
    else if (raw.type === 'measure') {
      this.post({ type: 'columnTexts', reqId, version, sheetId, axis: raw.axis === 'row' ? 'row' : 'col', index: 0, texts: [], error });
    }
  }

  private notAvailable(msg: WebviewMessage): void {
    getLog().info(`[${this.kind}] Ignored '${msg.type}' from the viewer of ${this.fileName}: not available in this view yet`);
  }

  private openEdit(editId: number): void {
    let close = (): void => undefined;
    const closed = new Promise<void>((resolve) => (close = resolve));
    this.openEdits.set(editId, { epoch: this.epoch, closed, close });
  }

  private closeEdit(editId: number): void {
    const open = this.openEdits.get(editId);
    this.openEdits.delete(editId);
    open?.close();
  }

  /** Called for every editResult before it is posted: false (dropped, logged) for a second answer to one edit. */
  private settleEdit(msg: Extract<HostMessage, { type: 'editResult' }>): boolean {
    const open = this.openEdits.get(msg.editId);
    if (!open) {
      getLog().error(`[${this.kind}] Dropped a second answer to edit ${msg.editId} of ${this.fileName}`);
      return false;
    }
    this.closeEdit(msg.editId);
    if (open.epoch !== this.epoch) {
      getLog().info(`[${this.kind}] Dropped the answer to edit ${msg.editId} of ${this.fileName}: the viewer was reloaded since`);
      return false;
    }
    if (!msg.accepted) {
      // Every rejection is logged (contract §3.4); the user sees it in the status line.
      getLog().info(`[${this.kind}] Edit ${msg.editId} of ${this.fileName} not applied (${msg.code}): ${msg.reason}`);
    }
    return true;
  }

  private armReadyTimer(): void {
    this.clearReadyTimer();
    this.readyTimer = setTimeout(() => {
      this.readyTimer = undefined;
      if (this.ready || this.disposed) return;
      if (!this.panel.visible) {
        // Hidden webviews may not load until shown; restart the clock when the panel becomes visible.
        this.waitingForVisibility = true;
        return;
      }
      getLog().warn(
        `[${this.kind}] The viewer of ${this.uri.toString(true)} did not report 'ready' within ` +
          `${READY_TIMEOUT_MS / 1000} s; dist/viewer.js probably failed to load or threw during start-up ` +
          `(run "Developer: Open Webview Developer Tools" for details).`,
      );
    }, READY_TIMEOUT_MS);
  }

  private clearReadyTimer(): void {
    if (this.readyTimer !== undefined) clearTimeout(this.readyTimer);
    this.readyTimer = undefined;
  }
}

/** Messages every viewer understands the same way. */
async function handleCommonMessage(
  session: ViewerSession,
  msg: WebviewMessage,
  policy: Omit<LinkPolicy, 'documentUri'>,
): Promise<void> {
  switch (msg.type) {
    case 'openLink':
      await openLink(msg.href, { ...policy, documentUri: session.uri });
      return;
    case 'reopenAsText':
      if (session.kind === 'csv' || session.kind === 'markdown') await reopenAsText(session.uri, session.panel.viewColumn);
      else getLog().warn(`[${session.kind}] Reopen as Text is not offered for ${session.fileName}; request ignored`);
      return;
    case 'revealInExplorer':
      await revealInOS(session.uri);
      return;
    case 'log':
      getLog()[msg.level](`[webview:${session.kind}] ${session.fileName}: ${msg.message}`);
      return;
    default:
      getLog().debug(`[${session.kind}] Ignored '${msg.type}' message (not applicable to this viewer)`);
  }
}

// ----- links -----

const EXTERNAL_LINK_SCHEMES: ReadonlySet<string> = new Set(['http', 'https', 'mailto']);
const SCHEME_RE = /^([a-zA-Z][a-zA-Z0-9+.-]*):/;
/** Drive letter + separator. The separator may still be encoded: markdown-it writes `C:\docs` as `C:%5Cdocs`. */
const WINDOWS_ABSOLUTE_RE = /^[a-zA-Z]:(?:[\\/]|%5c)/i;

/** Lower-case URI scheme of a link, or undefined for relative / absolute file-system paths. */
function linkScheme(href: string): string | undefined {
  if (WINDOWS_ABSOLUTE_RE.test(href)) return undefined;
  const match = SCHEME_RE.exec(href);
  return match ? match[1].toLowerCase() : undefined;
}

/**
 * A `file:` link as a Uri. Word stores local hyperlinks as `file:///C:\dir\Other.docx` (Windows separators, spaces
 * often unencoded): backslashes in the path are separators, not part of a name.
 */
function fileLinkUri(link: string): vscode.Uri {
  const hashAt = link.indexOf('#');
  const main = hashAt >= 0 ? link.slice(0, hashAt) : link;
  return vscode.Uri.parse(main.replace(/\\/g, '/') + (hashAt >= 0 ? link.slice(hashAt) : ''), true);
}

/** Splits `path?query#fragment` (all parts still encoded). */
function splitHref(href: string): { path: string; query: string; fragment: string } {
  const hashAt = href.indexOf('#');
  const fragment = hashAt >= 0 ? href.slice(hashAt + 1) : '';
  const rest = hashAt >= 0 ? href.slice(0, hashAt) : href;
  const queryAt = rest.indexOf('?');
  return {
    path: queryAt >= 0 ? rest.slice(0, queryAt) : rest,
    query: queryAt >= 0 ? rest.slice(queryAt + 1) : '',
    fragment,
  };
}

function safeDecode(text: string): string {
  try {
    return decodeURIComponent(text);
  } catch {
    return text;
  }
}

/**
 * Resolves a link/image path (as written in the document, still URI-encoded, without query/fragment) to a URI:
 * relative paths against the document's folder, `/x` against the document's workspace folder, Windows absolute
 * and UNC paths as files. Backslashes are accepted as separators (Excel links often use them).
 */
function resolveLocalPath(rawPath: string, documentUri: vscode.Uri): vscode.Uri | undefined {
  if (!rawPath) return undefined;
  const unescaped = safeDecode(rawPath);
  const isUnc = unescaped.startsWith('\\\\');
  const decoded = unescaped.replace(/\\/g, '/');
  if (WINDOWS_ABSOLUTE_RE.test(decoded) || isUnc) return vscode.Uri.file(decoded);
  if (decoded.startsWith('//')) return undefined; // protocol-relative URL: not a local path
  const firstFolder = vscode.workspace.workspaceFolders?.[0]?.uri;
  if (documentUri.scheme === 'untitled') {
    // An unsaved document has no folder: resolve against the first workspace folder, if any.
    return firstFolder ? vscode.Uri.joinPath(firstFolder, decoded) : undefined;
  }
  const base = documentUri.with({ query: '', fragment: '' });
  if (decoded.startsWith('/')) {
    const folder = vscode.workspace.getWorkspaceFolder(base);
    return folder ? vscode.Uri.joinPath(folder.uri, decoded) : base.with({ path: decoded });
  }
  return vscode.Uri.joinPath(base, '..', decoded);
}

/**
 * Line fragments as VS Code's Markdown links accept them -> editor selection (1-based): `#L10` or `#10` (line 10),
 * `#L10,5` (line 10, column 5), `#L10-L20`, `#L10,5-L20,2`. A reversed range is swapped; line 0 is not a line. A range
 * end without a column includes that whole line (`#L10-L20` selects lines 10 to 20, like GitHub's highlight).
 */
function selectionFromFragment(fragment: string): vscode.Range | undefined {
  const match = /^L?(\d+)(?:,(\d+))?(?:-L?(\d+)(?:,(\d+))?)?$/i.exec(fragment);
  if (!match || Number(match[1]) === 0 || Number(match[3]) === 0) return undefined;
  const point = (line: string, column: string | undefined) => ({
    line: Number(line) - 1,
    column: column === undefined ? undefined : Math.max(0, Number(column) - 1),
  });
  let start = point(match[1], match[2]);
  if (match[3] === undefined) return new vscode.Range(start.line, start.column ?? 0, start.line, start.column ?? 0);
  let end = point(match[3], match[4]);
  if (end.line < start.line || (end.line === start.line && (end.column ?? 0) < (start.column ?? 0))) {
    [start, end] = [end, start];
  }
  return end.column === undefined
    ? new vscode.Range(start.line, start.column ?? 0, end.line + 1, 0)
    : new vscode.Range(start.line, start.column ?? 0, end.line, end.column);
}

function displayPath(uri: vscode.Uri): string {
  return uri.scheme === 'file' ? uri.fsPath : uri.toString(true);
}

/**
 * A heading fragment of a link to a Markdown file (`other.md#install-steps`) -> the heading's line, found with the
 * renderer's slugs (the preview's ids: exact, then case-insensitive, also with GitHub's `user-content-` prefix).
 */
async function headingFromFragment(target: vscode.Uri, fragment: string): Promise<vscode.Range | undefined> {
  if (!fragment || viewKindForUri(target) !== 'markdown') return undefined;
  let text: string;
  try {
    text = (await vscode.workspace.openTextDocument(target)).getText(); // the open (maybe unsaved) text if there is one
  } catch {
    return undefined; // unreadable / too large for the extension host: open it at the top
  }
  const toc = markdownHeadings(text);
  const ids = [fragment, fragment.replace(/^user-content-/i, '')];
  const entry =
    toc.find((t) => ids.includes(t.slug)) ?? toc.find((t) => ids.some((id) => id.toLowerCase() === t.slug.toLowerCase()));
  return entry ? new vscode.Range(entry.line, 0, entry.line, 0) : undefined;
}

/** A missing file. Other stat failures are not "not found", e.g. a UNC host refused by `security.allowedUNCHosts`. */
function isFileNotFound(err: unknown): boolean {
  return err instanceof vscode.FileSystemError && err.code === 'FileNotFound';
}

/**
 * Opens a local target: directories are revealed in the explorer, files open in their default editor. Like VS Code's
 * Markdown links, a fragment is first the heading it names in a Markdown file (`#l2` -> `## L2`), else a line
 * (`#L10`, `#10`, `#L10,5`, `#L10-L20`).
 */
async function openLocalTarget(target: vscode.Uri, fragment: string): Promise<void> {
  let stat: vscode.FileStat;
  try {
    stat = await vscode.workspace.fs.stat(target);
  } catch (err) {
    if (isFileNotFound(err)) {
      void vscode.window.showWarningMessage(`FileStudio could not find "${displayPath(target)}".`);
      return;
    }
    // VS Code's own reason, e.g. "UNC host 'server' access is not allowed. Please update the
    // 'security.allowedUNCHosts' setting if you want to allow this host." (the file may well exist).
    getLog().warn(`Could not open link target ${target.toString(true)}: ${errorMessageOf(err)}`);
    void vscode.window.showWarningMessage(`FileStudio could not open "${displayPath(target)}": ${errorMessageOf(err)}`);
    return;
  }
  if (stat.type & vscode.FileType.Directory) {
    await vscode.commands.executeCommand('revealInExplorer', target);
    return;
  }
  const selection = (await headingFromFragment(target, fragment)) ?? selectionFromFragment(fragment);
  await vscode.commands.executeCommand('vscode.open', target, selection ? { selection } : undefined);
}

function reportBlockedLink(href: string, reason: string): void {
  getLog().warn(`Blocked link "${href.slice(0, 200)}": ${reason}`);
  void vscode.window.showWarningMessage(`FileStudio did not open the link: ${reason}.`);
}

/**
 * Opens an http/https/mailto URL in the browser / mail client exactly as written. It is normalized the way a browser
 * normalizes an href (WHATWG URL: spaces and non-ASCII percent-encoded, existing escapes kept) and handed to VS Code
 * as a string. A Uri would be opened as encodeURI(uri.toString(true)), which decodes and re-encodes every escape
 * (`?q=C%23` -> `C%2523`, `a%26b` -> `a&b`, `C%2B%2B` -> `C++`).
 */
async function openExternalUrl(link: string): Promise<void> {
  let url: string;
  try {
    url = new URL(link).href;
  } catch {
    reportBlockedLink(link, 'it is not a valid web address');
    return;
  }
  // Typed as Uri, but the extension host accepts a string and the opener then uses it as is.
  await vscode.env.openExternal(url as unknown as vscode.Uri);
}

/** How a viewer follows links (see openLink). */
interface LinkPolicy {
  documentUri: vscode.Uri;
  /** `file:` URIs and relative / absolute paths open inside VS Code (else only http/https/mailto). */
  allowLocal: boolean;
  /** A local link to the document itself (`readme.md#usage` in readme.md): handled by the viewer, not opened again. */
  onSameDocument?: (fragment: string) => void;
}

/**
 * Follows a link clicked in a viewer. http/https/mailto open externally; with `allowLocal`, `file:` URIs and
 * relative/absolute paths open inside VS Code (a link to the document itself goes to `onSameDocument` when given).
 * Every other scheme (notably `command:`) is refused.
 */
export async function openLink(href: string, policy: LinkPolicy): Promise<void> {
  const link = href.trim();
  if (!link || link.startsWith('#')) return; // in-page anchors are handled by the webview
  try {
    const scheme = linkScheme(link);
    let target: vscode.Uri | undefined;
    let fragment: string;
    if (scheme !== undefined) {
      if (EXTERNAL_LINK_SCHEMES.has(scheme)) {
        await openExternalUrl(link);
        return;
      }
      if (scheme !== 'file' || !policy.allowLocal) {
        reportBlockedLink(link, `"${scheme}:" links are not allowed`);
        return;
      }
      const uri = fileLinkUri(link);
      target = uri.with({ query: '', fragment: '' });
      fragment = uri.fragment;
    } else {
      if (!policy.allowLocal) {
        reportBlockedLink(link, 'only http, https and mailto links can be opened from this document');
        return;
      }
      const parts = splitHref(link);
      target = resolveLocalPath(parts.path, policy.documentUri);
      fragment = safeDecode(parts.fragment);
      if (!target) {
        reportBlockedLink(link, 'the link target is not a local path');
        return;
      }
    }
    if (policy.onSameDocument && isSameFile(target, policy.documentUri)) {
      policy.onSameDocument(fragment);
      return;
    }
    await openLocalTarget(target, fragment);
  } catch (err) {
    logError(`Could not open link "${link.slice(0, 200)}"`, err);
    void vscode.window.showErrorMessage(`FileStudio could not open the link: ${errorMessageOf(err)}`);
  }
}

// ===== SHARED: EDIT GATE =====
// (v0.2.0, track F; contract §3) The extension host is the authority for every change FileStudio makes to a file.
// Building blocks for every editable view: one serial queue per document, the kill switch (4 settings, read on every
// call), the read-only check, checkEditable (§3.2 steps 1-3), the reject texts (§3.4), ask() for every modal prompt
// (§3.6; answered by the test API in tests), checkTextFidelity (§8.6) and a registry of open documents whose edit
// state follows the settings (one config listener). The document-specific steps (version rule, planning, applying)
// belong to the phase that owns the document (P1/P2 xlsx, P3 csv, P6 Markdown WYSIWYG). Turning editing off stops
// only NEW edits: save, revert, undo/redo and backup of a dirty document never check the kill switch (D-075).

// ----- texts (contract §3.4) -----

/** User texts of rejected edits and of not-editable states (status line, aria-live; every rejection is also logged). */
export const EDIT_TEXT = {
  disabled: (setting: string): string => `Editing is turned off (setting ${setting}).`,
  /** `why`: 'files.readonlyInclude', 'Git version', 'no write permission', 'read-only file system'. */
  readOnly: (why: string): string => `This file is read-only (${why}).`,
  failed: 'VS Code did not apply the change. The file was not changed.',
  stale: 'The file changed while you were editing. Your change was not applied; the view was refreshed.',
  lossy: 'This file has content FileStudio cannot show or update. Choose "Edit anyway" to edit it.',
  fidelity: 'Saving in VS Code would also change parts of this file you did not edit. Choose "Edit anyway" to edit it.',
  fidelityTitle: 'Saving in VS Code changes parts of this file you did not edit.',
  fidelityPending: 'FileStudio is still checking this file; try again in a moment.',
  checking: 'Checking the file…',
  fidelityUnreadable: 'FileStudio could not check how VS Code will save this file.',
  invalid: 'FileStudio could not read this change. Nothing was changed.',
  tooLarge: (limit: string): string => `This change is too large (max ${limit}).`,
  badFormula: "There's a problem with this formula.",
  badPattern: 'There is a problem with this search pattern.',
  badRequest: 'FileStudio could not read this request.',
  badSheetName: 'This sheet name cannot be used.',
  loneSurrogate: 'The text has an unpaired surrogate character and cannot be saved.',
  internal: 'Internal check failed. Nothing was changed.',
  poisoned: 'An internal error happened. Your file on disk was not changed. Use Revert File to reload.',
  notAvailable: 'This action is not available yet.',
} as const;

// ----- serial queue (contract §3.1) -----

/**
 * Runs jobs one at a time, in the order they were added (FIFO). A job that throws or rejects does not stop the
 * queue: its own promise rejects and the next job starts. One queue per document runs edits, VS Code undo/redo
 * callbacks, save, revert, disk reloads and Markdown task toggles (contract §3.1); reads do not wait.
 */
export class SerialQueue {
  private tail: Promise<void> = Promise.resolve();
  private waiting = 0;

  /** Jobs added and not finished yet (running included). */
  get size(): number {
    return this.waiting;
  }

  run<T>(job: () => T | Promise<T>): Promise<T> {
    this.waiting++;
    const result = this.tail.then(job);
    const done = (): void => {
      this.waiting--;
    };
    this.tail = result.then(done, done);
    return result;
  }
}

// ----- prompts (contract §3.6) -----

/**
 * Every modal question FileStudio asks goes through here (contract §0 rule 7). In ExtensionMode.Test the answer comes
 * from the test API queue (TEST API `answer`); with no queued answer it is undefined (= Cancel), so a test never waits
 * on a dialog. A queued answer that is not one of `buttons` also counts as Cancel. Every question is logged.
 */
export async function ask(
  id: AskId,
  message: string,
  options: { modal: true; detail?: string },
  ...buttons: string[]
): Promise<string | undefined> {
  if (testApiEnabled) {
    const queued = testAnswers.get(id);
    const hasAnswer = queued !== undefined && queued.length > 0;
    const answer = hasAnswer ? queued.shift() : undefined;
    const chosen = answer !== undefined && buttons.includes(answer) ? answer : undefined;
    getLog().info(
      `[test] ask('${id}') "${message}" -> ${chosen === undefined ? 'Cancel' : `"${chosen}"`}` +
        (hasAnswer ? '' : ' (no queued answer)'),
    );
    return chosen;
  }
  const answer = await vscode.window.showWarningMessage(message, options, ...buttons);
  getLog().info(`Asked (${id}) "${message}" -> ${answer === undefined ? 'Cancel' : `"${answer}"`}`);
  return answer;
}

// ----- kill switch and read-only (contract §3.2 steps 1-2) -----

/** The document formats FileStudio can change. */
export type EditFormat = 'xlsx' | 'csv' | 'markdown';

/** The four editing settings (contract §11): the master switch and one per format. */
export const EDIT_SETTINGS = [
  'fileStudio.editing.enabled',
  'fileStudio.xlsx.editing',
  'fileStudio.csv.editing',
  'fileStudio.markdown.editing',
] as const;

/**
 * The setting that turns editing of `f` off for `uri` (the master switch first), or undefined when editing is on.
 * Read on every call (no cache). Only an explicit `false` turns editing off (the default is true).
 */
export function disabledEditSetting(f: EditFormat, uri: vscode.Uri): string | undefined {
  const config = vscode.workspace.getConfiguration('fileStudio', uri);
  // Off for any value but true: a hand-edited "false", 0 or null must not leave editing on. Unset (undefined) = on.
  const off = (v: unknown): boolean => v !== undefined && v !== true;
  if (off(config.get<unknown>('editing.enabled'))) return 'fileStudio.editing.enabled';
  if (off(config.get<unknown>(`${f}.editing`))) return `fileStudio.${f}.editing`;
  return undefined;
}

/** Kill switch (contract §3.1): `fileStudio.editing.enabled` && `fileStudio.<f>.editing`, read on every call. */
export function editingEnabled(f: EditFormat, uri: vscode.Uri): boolean {
  return disabledEditSetting(f, uri) === undefined;
}

/**
 * Why VS Code's own editor treats `uri` as read-only (filesConfigurationService.isReadonly), or undefined: a read-only
 * file system ('Git version' for git:), then `files.readonlyInclude`, or `files.readonlyFromPermissions` with a file
 * that cannot be written ('no write permission'), both unless `files.readonlyExclude` matches. Like VS Code, a glob
 * matches the path relative to the workspace folder or the absolute path. "Set Active Editor Read-only in Session" is
 * not visible to extensions (an edit is then refused by VS Code: EDIT_TEXT.failed). `target`: the TextDocument when
 * there is one; otherwise (xlsx) a minimal {uri, languageId} object stands in for it in languages.match (F WP7 checks
 * this in a real VS Code).
 */
export async function readonlyReasonOf(uri: vscode.Uri, target?: vscode.TextDocument): Promise<string | undefined> {
  if (vscode.workspace.fs.isWritableFileSystem(uri.scheme) === false) {
    return uri.scheme === 'git' ? 'Git version' : 'read-only file system';
  }
  const document = target ?? ({ uri, languageId: '' } as unknown as vscode.TextDocument);
  const files = vscode.workspace.getConfiguration('files', uri);
  const folder = vscode.workspace.getWorkspaceFolder(uri);
  const matches = (setting: string): boolean =>
    Object.entries(files.get<Record<string, boolean>>(setting) ?? {}).some(
      ([glob, on]) =>
        on === true &&
        (vscode.languages.match({ pattern: glob }, document) > 0 ||
          (folder !== undefined && vscode.languages.match({ pattern: new vscode.RelativePattern(folder, glob) }, document) > 0)),
    );
  if (matches('readonlyInclude')) return matches('readonlyExclude') ? undefined : 'files.readonlyInclude';
  // files.readonlyFromPermissions is window-scoped: read it without a resource (as VS Code does), or the extension
  // host logs "Accessing a window scoped configuration for a resource is not expected".
  if (uri.scheme === 'file' && vscode.workspace.getConfiguration('files').get<boolean>('readonlyFromPermissions') === true) {
    try {
      await access(uri.fsPath, fsConstants.W_OK);
    } catch (err) {
      // A file that is gone is not read-only: VS Code shows it as deleted, and a save writes it again.
      if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') return matches('readonlyExclude') ? undefined : 'no write permission';
    }
  }
  return undefined;
}

/** `isReadonlyDocument` generalised to a Uri (contract §3.1). */
export async function isReadonlyResource(uri: vscode.Uri): Promise<boolean> {
  return (await readonlyReasonOf(uri)) !== undefined;
}

// ----- gate (contract §3.1, §3.2) -----

export type GateResult = { ok: true } | { ok: false; code: EditRejectCode; reason: string };

/** The document's own state for the gate (contract §3.1), kept by the phase that owns the document. */
export interface DocGateState {
  /** xlsx: an undo/redo/rollback step failed (contract §3.5). */
  poisoned?: boolean;
  /** xlsx: lossy features, "Edit anyway" accepted or not. */
  lossy?: { accepted: boolean };
  /** csv / Markdown WYSIWYG: checkTextFidelity result (pending = not finished yet = not editable, D-072). */
  fidelity?: { pending?: boolean; ok: boolean; accepted: boolean; blocked?: string };
  /** A hard block (unpatchable package, unclosed quote, …): user text. */
  block?: string;
}

/** §3.2 step 3 (synchronous): document state → poisoned, lossy, fidelity (pending first), hard blocks. */
export function documentGate(s: DocGateState): GateResult {
  if (s.poisoned) return { ok: false, code: 'internal', reason: EDIT_TEXT.poisoned };
  if (s.lossy && !s.lossy.accepted) return { ok: false, code: 'lossy', reason: EDIT_TEXT.lossy };
  if (s.fidelity) {
    if (s.fidelity.pending) return { ok: false, code: 'fidelity', reason: EDIT_TEXT.fidelityPending };
    if (s.fidelity.blocked !== undefined) return { ok: false, code: 'blocked', reason: s.fidelity.blocked };
    if (!s.fidelity.ok && !s.fidelity.accepted) return { ok: false, code: 'fidelity', reason: EDIT_TEXT.fidelity };
  }
  if (s.block !== undefined) return { ok: false, code: 'blocked', reason: s.block };
  return { ok: true };
}

/**
 * §3.2 steps 1-3 for one edit or toggle (call it inside the document queue): kill switch (sync) → read-only (async)
 * → document state. The kill switch is read again after the read-only await, so a setting turned off meanwhile still
 * wins. The caller then runs step 4 (version rule → plan → apply) synchronously, with no await in between.
 * Never rejects: a check that fails unexpectedly refuses the edit ('internal').
 */
export async function checkEditable(
  f: EditFormat,
  uri: vscode.Uri,
  s: DocGateState = {},
  target?: vscode.TextDocument,
): Promise<GateResult> {
  try {
    let off = disabledEditSetting(f, uri);
    if (off !== undefined) return { ok: false, code: 'disabled', reason: EDIT_TEXT.disabled(off) };
    const readOnly = await readonlyReasonOf(uri, target);
    off = disabledEditSetting(f, uri);
    if (off !== undefined) return { ok: false, code: 'disabled', reason: EDIT_TEXT.disabled(off) };
    if (readOnly !== undefined) return { ok: false, code: 'readOnly', reason: EDIT_TEXT.readOnly(readOnly) };
    return documentGate(s);
  } catch (err) {
    logError(`Could not check whether ${uri.toString(true)} can be edited`, err);
    return { ok: false, code: 'internal', reason: EDIT_TEXT.internal };
  }
}

/** EditState for a gate result: editable with `kinds`, or not editable with the code and reason (and no kinds). */
export function editStateOf(result: GateResult, kinds: EditKind[]): EditState {
  return result.ok ? { editable: true, kinds } : { editable: false, code: result.code, reason: result.reason, kinds: [] };
}

/** xlsx and CSV until their editing phases land (contract §14 F): view-only exactly as 0.1.2, no banner or message. */
export function viewOnlyEditState(): EditState {
  return { editable: false, kinds: [] };
}

function sameEditState(a: EditState | undefined, b: EditState | undefined): boolean {
  return a !== undefined && b !== undefined && JSON.stringify(a) === JSON.stringify(b);
}

// ----- text fidelity (contract §8.6, D-035, D-072) -----

/** Largest file checkTextFidelity reads (VS Code gives extensions at most 50 MB of text anyway). */
const FIDELITY_MAX_BYTES = 50 * 1024 * 1024;

/** Result of checkTextFidelity (the caller stores it in DocGateState.fidelity and its EditState). */
export interface TextFidelity {
  /** Saving in VS Code writes back exactly the bytes on disk (or there is nothing on disk to change: untitled). */
  ok: boolean;
  /** Never editable: the text was decoded with the wrong encoding (user text, no "Edit anyway"). */
  blocked?: string;
  /** Not ok and not blocked: what VS Code's save would change, for the "Edit anyway" question and the banner. */
  canEditAnyway?: { title: string; details: string[] };
}

/** VS Code API newer than @types/vscode 1.90 (engine ^1.123): workspace.encode / decode and TextDocument.encoding. */
interface TextCodecApi {
  encode(text: string, options: { encoding: string }): Thenable<Uint8Array>;
  decode(content: Uint8Array, options: { encoding: string }): Thenable<string>;
}

function fidelityChanges(details: string[]): TextFidelity {
  return { ok: false, canEditAnyway: { title: EDIT_TEXT.fidelityTitle, details } };
}

function plural(n: number, word: string): string {
  return `${n.toLocaleString('en-US')} ${word}${n === 1 ? '' : 's'}`;
}

/** Line breaks in `text` by kind. */
function countLineBreaks(text: string): { crlf: number; lf: number; cr: number } {
  let crlf = 0;
  let lf = 0;
  let cr = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    if (ch === 13) {
      if (text.charCodeAt(i + 1) === 10) {
        crlf++;
        i++;
      } else cr++;
    } else if (ch === 10) lf++;
  }
  return { crlf, lf, cr };
}

function countChar(text: string, code: number): number {
  let n = 0;
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) === code) n++;
  return n;
}

/** Byte sequence of U+FFFD / of a BOM in `encoding` (undefined: the encoding cannot store it). */
function encodedMark(encoding: string, mark: 'replacement' | 'bom'): number[] | undefined {
  const enc = encoding.toLowerCase();
  if (enc === 'utf8' || enc === 'utf8bom') return mark === 'replacement' ? [0xef, 0xbf, 0xbd] : [0xef, 0xbb, 0xbf];
  if (enc === 'utf16le') return mark === 'replacement' ? [0xfd, 0xff] : [0xff, 0xfe];
  if (enc === 'utf16be') return mark === 'replacement' ? [0xff, 0xfd] : [0xfe, 0xff];
  return undefined;
}

function startsWithAt(bytes: Uint8Array, at: number, seq: number[]): boolean {
  for (let k = 0; k < seq.length; k++) if (bytes[at + k] !== seq[k]) return false;
  return true;
}

/** U+FFFD characters really stored in the bytes (UTF-16: at character boundaries only). */
function countReplacementBytes(bytes: Uint8Array, encoding: string): number {
  const seq = encodedMark(encoding, 'replacement');
  if (!seq) return 0;
  const step = seq.length === 2 ? 2 : 1;
  let n = 0;
  for (let i = 0; i + seq.length <= bytes.length; i += step) if (startsWithAt(bytes, i, seq)) n++;
  return n;
}

/** Byte order marks at the very start of the bytes. */
function countLeadingBoms(bytes: Uint8Array, encoding: string): number {
  const seq = encodedMark(encoding, 'bom');
  if (!seq) return 0;
  let n = 0;
  while (startsWithAt(bytes, n * seq.length, seq)) n++;
  return n;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Lines that end in spaces or tabs (VS Code's trimTrailingWhitespace). A plain loop: no array of matches, linear. */
function countTrailingWhitespaceLines(text: string): number {
  let n = 0;
  let inRun = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    if (ch === 32 || ch === 9) inRun = true;
    else {
      if (inRun && (ch === 10 || ch === 13)) n++;
      inRun = false;
    }
  }
  return inRun ? n + 1 : n;
}

/**
 * What VS Code's save participants will change in `text` with the user's settings (read for the document, so the
 * `[csv]` / `[markdown]` language overrides apply): trailing whitespace, a missing final newline, extra final newlines.
 */
function saveSettingChanges(document: vscode.TextDocument, text: string): string[] {
  const files = vscode.workspace.getConfiguration('files', document);
  const details: string[] = [];
  if (files.get<unknown>('trimTrailingWhitespace') === true) {
    const lines = countTrailingWhitespaceLines(text);
    if (lines > 0) {
      details.push(`VS Code will remove spaces at the end of ${plural(lines, 'line')} (setting files.trimTrailingWhitespace).`);
    }
  }
  if (files.get<unknown>('insertFinalNewline') === true && text.length > 0 && !/[\r\n]$/.test(text) && !/(?:^|[\r\n])[ \t]+$/.test(text)) {
    // (VS Code adds nothing when the last line is only spaces or tabs.)
    details.push('VS Code will add a line break at the end of the file (setting files.insertFinalNewline).');
  }
  if (files.get<unknown>('trimFinalNewlines') === true) {
    // VS Code's rule (checked in 1.142): only truly empty lines are trimmed (a line of spaces stays and stops the
    // trimming); a file of line breaks only becomes empty. The last "line" after the final line break is not extra.
    // Scans from the end: no array of lines for a big file.
    let blank = 0;
    let e = text.length;
    for (;;) {
      const prev = e > 0 ? text.charCodeAt(e - 1) : -1;
      if (e > 0 && prev !== 10 && prev !== 13) break; // a line with text
      blank++; // the line that ends at e is empty
      if (e === 0) break;
      e -= prev === 10 && e >= 2 && text.charCodeAt(e - 2) === 13 ? 2 : 1;
    }
    const extra = blank - 1;
    if (extra > 0) {
      details.push(`VS Code will remove ${plural(extra, 'empty line')} at the end of the file (setting files.trimFinalNewlines).`);
    }
  }
  return details;
}

/**
 * Checks whether saving `document` in VS Code writes back the bytes on disk unchanged (contract §8.6, D-035, D-072):
 * CSV and Markdown WYSIWYG call it (P3, P6) once per open and again after a non-own change that leaves the document
 * clean; Markdown task toggles are exempt. `untitled:` → ok. The file's bytes are compared with
 * `workspace.encode(getText(), {encoding: document.encoding})`; a difference is explained (mixed / CR-only line
 * breaks with counts; several leading BOMs are the same bytes), and VS Code's save settings are counted (trim trailing whitespace,
 * insert / trim final newlines). Replacement characters that are not in the bytes (wrong encoding) → `blocked`.
 * Unreadable bytes, a file over 50 MB or a VS Code without the encoding API → not ok with "Edit anyway"
 * (EDIT_TEXT.fidelityUnreadable). A document already dirty is judged by the line breaks on disk only.
 */
export async function checkTextFidelity(document: vscode.TextDocument): Promise<TextFidelity> {
  const uri = document.uri;
  if (uri.scheme === 'untitled') return { ok: true };
  const text = document.getText();
  let details: string[];
  try {
    details = saveSettingChanges(document, text);
  } catch (err) {
    logError(`[fidelity] Could not read the save settings for ${uri.toString(true)}`, err);
    return fidelityChanges([EDIT_TEXT.fidelityUnreadable]);
  }
  const codec = vscode.workspace as unknown as Partial<TextCodecApi>;
  const encoding = (document as unknown as { encoding?: unknown }).encoding;
  let disk: Uint8Array;
  let diskText: string;
  try {
    if (typeof codec.encode !== 'function' || typeof codec.decode !== 'function' || typeof encoding !== 'string') {
      throw new Error('this VS Code has no text encoding API');
    }
    const stat = await vscode.workspace.fs.stat(uri);
    if (stat.size > FIDELITY_MAX_BYTES) throw new Error('the file is larger than 50 MB');
    disk = await vscode.workspace.fs.readFile(uri);
    if (!document.isDirty) {
      const encoded = await codec.encode(text, { encoding });
      // VS Code strips every leading BOM from the text but writes them all back on save (checked in 1.142): a file with
      // several BOMs is the same bytes when only the extra BOMs differ.
      const boms = countLeadingBoms(disk, encoding);
      const bomLength = encodedMark(encoding, 'bom')?.length ?? 0;
      if (sameBytes(encoded, disk) || (boms > 1 && sameBytes(encoded, disk.subarray((boms - 1) * bomLength)))) {
        return details.length === 0 ? { ok: true } : fidelityChanges(details);
      }
    }
    diskText = await codec.decode(disk, { encoding });
  } catch (err) {
    getLog().info(`[fidelity] Could not check how VS Code will save ${uri.toString(true)}: ${errorMessageOf(err)}`);
    return fidelityChanges([EDIT_TEXT.fidelityUnreadable, ...details]);
  }
  // Wrong encoding: decoding put replacement characters where the bytes have none; a save would write U+FFFD there.
  if (countChar(diskText, 0xfffd) > countReplacementBytes(disk, encoding)) {
    return {
      ok: false,
      blocked:
        `FileStudio cannot edit this file safely: it is not valid ${encoding} text. ` +
        'Use "Reopen with Encoding" in VS Code to open it with the right encoding.',
    };
  }
  const explained: string[] = [];
  const breaks = countLineBreaks(diskText);
  const toCrlf = document.eol === vscode.EndOfLine.CRLF;
  const target = toCrlf ? 'CRLF' : 'LF';
  const changes: [number, string][] = toCrlf
    ? [
        [breaks.lf, 'LF'],
        [breaks.cr, 'CR'],
      ]
    : [
        [breaks.crlf, 'CRLF'],
        [breaks.cr, 'CR'],
      ];
  for (const [count, from] of changes) {
    if (count > 0) explained.push(`VS Code will change ${plural(count, 'line break')} (${from} → ${target}).`);
  }
  // A clean document whose bytes differ for no reason found above: still not the same bytes.
  if (!document.isDirty && explained.length === 0) {
    explained.push(`VS Code will save some characters differently (encoding ${encoding}).`);
  }
  const all = [...explained, ...details];
  return all.length === 0 ? { ok: true } : fidelityChanges(all);
}

// ----- open documents and the config listener (contract §3.6) -----

/** What a view registers for its document (the first panel's spec is used; all panels of a document share it). */
export interface EditGateDocSpec {
  format: EditFormat;
  uri: vscode.Uri;
  /** The document's edit state now (gate checks + the document's own state). */
  computeEditState(): EditState | Promise<EditState>;
  /** For the test API `state()`: the document's host version and dirty flag (0 / false when absent). */
  version?(): number;
  dirty?(): boolean;
}

/** One open document known to the edit gate, shared by all its panels. */
export class EditGateDoc {
  /** The document's queue for F's Markdown task toggles (xlsx and CSV keep their queue in SheetDocument / CsvDocState). */
  readonly queue = new SerialQueue();
  /** The last edit state computed (undefined before the first refresh). */
  state: EditState | undefined;
  /** One callback per attached panel. */
  readonly listeners = new Set<(state: EditState) => void>();
  private refreshes = 0;

  constructor(
    readonly key: string,
    readonly spec: EditGateDocSpec,
  ) {}

  /**
   * Computes the edit state again (settings and read-only state are read now) and calls every panel's callback when
   * it changed (not for the first state: each view sends that in its init). Resolves with the fresh state; never rejects (a failing computation gives a not-editable 'internal'
   * state, so nothing is changed by mistake).
   */
  async refresh(): Promise<EditState> {
    const seq = ++this.refreshes;
    let state: EditState;
    try {
      state = await this.spec.computeEditState();
    } catch (err) {
      logError(`Could not compute the edit state of ${this.spec.uri.toString(true)}`, err);
      state = { editable: false, code: 'internal', reason: EDIT_TEXT.internal, kinds: [] };
    }
    if (seq === this.refreshes && !sameEditState(state, this.state)) {
      const first = this.state === undefined; // the views send the first state in their init
      this.state = state;
      for (const listener of first ? [] : [...this.listeners]) {
        try {
          listener(state);
        } catch (err) {
          logError(`Could not send the edit state of ${this.spec.uri.toString(true)}`, err);
        }
      }
    }
    return state;
  }
}

/** A panel's attachment to its document's EditGateDoc; dispose it with the panel. */
export interface EditGateHandle extends vscode.Disposable {
  readonly doc: EditGateDoc;
}

const editGateDocs = new Map<string, EditGateDoc>();
let editConfigListener: vscode.Disposable | undefined;

/**
 * Attaches one panel to the EditGateDoc of its document (made by the first panel, removed with the last one).
 * `onState` runs whenever the document's edit state changes (a setting, the read-only state, the document's own
 * state) and sends it to the panel (`editState`, or the view's own update message).
 */
export function attachEditGateDoc(spec: EditGateDocSpec, onState: (state: EditState) => void): EditGateHandle {
  const key = `${spec.format}:${spec.uri.toString()}`;
  let doc = editGateDocs.get(key);
  if (!doc) {
    doc = new EditGateDoc(key, spec);
    editGateDocs.set(key, doc);
    void doc.refresh();
  }
  doc.listeners.add(onState);
  watchEditSettings();
  const attached = doc;
  let disposed = false;
  return {
    doc: attached,
    dispose: () => {
      if (disposed) return;
      disposed = true;
      attached.listeners.delete(onState);
      if (attached.listeners.size === 0 && editGateDocs.get(key) === attached) editGateDocs.delete(key);
      if (editGateDocs.size === 0) {
        editConfigListener?.dispose();
        editConfigListener = undefined;
      }
    },
  };
}

/**
 * The one configuration listener (contract §3.6): a change of an editing setting or a read-only setting recomputes
 * the edit state of every open document it affects and sends it to the document's panels. Exists while documents are
 * attached. The settings are still read again for every edit.
 */
function watchEditSettings(): void {
  if (editConfigListener) return;
  editConfigListener = vscode.workspace.onDidChangeConfiguration((e) => {
    const settings: readonly string[] = [...EDIT_SETTINGS, ...READONLY_SETTINGS];
    for (const doc of [...editGateDocs.values()]) {
      if (settings.some((setting) => e.affectsConfiguration(setting, doc.spec.uri))) void doc.refresh();
    }
  });
}

// ===== SHARED: DEBOUNCE & FILE WATCHING =====

interface Debounced extends vscode.Disposable {
  schedule(): void;
  cancel(): void;
}

function debounce(fn: () => void, delayMs: number): Debounced {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const cancel = (): void => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
  };
  return {
    schedule() {
      cancel();
      timer = setTimeout(() => {
        timer = undefined;
        fn();
      }, delayMs);
    },
    cancel,
    dispose: cancel,
  };
}

/** Writers emit several events per save (truncate, write, rename): reload once they have settled. */
const DISK_RELOAD_DEBOUNCE_MS = 300;

/** Size + modification time of a file: the state a document was loaded from or last saved as. */
interface DiskStamp {
  readonly size: number;
  readonly mtime: number;
}

/** The file's current stamp, or undefined when it cannot be read (deleted, not a file system resource). */
async function diskStampOf(uri: vscode.Uri): Promise<DiskStamp | undefined> {
  try {
    const stat = await vscode.workspace.fs.stat(uri);
    return { size: stat.size, mtime: stat.mtime };
  } catch {
    return undefined;
  }
}

/** Nothing to reload: the file is gone (keep the last state) or still has the known stamp (own write, repeated event). */
function isUnchanged(current: DiskStamp | undefined, known: DiskStamp | undefined): boolean {
  return current === undefined || (known !== undefined && current.size === known.size && current.mtime === known.mtime);
}

/** Same resource; paths compare case-insensitively on Windows and macOS (file events may differ in case). */
function isSameFile(a: vscode.Uri, b: vscode.Uri): boolean {
  if (a.scheme !== b.scheme || a.authority.toLowerCase() !== b.authority.toLowerCase()) return false;
  if (a.path === b.path) return true;
  return a.scheme === 'file' && process.platform !== 'linux' && a.path.toLowerCase() === b.path.toLowerCase();
}

/**
 * Calls `reload` when another program creates or changes the file: VS Code reloads TextDocuments itself, but not the
 * binary custom documents (xlsx/docx). Debounced, and never run twice at the same time. The folder is watched without
 * recursion and events are matched by path, so file names with glob characters (`[1].xlsx`) need no escaping.
 * Deletion is ignored: the viewer keeps showing the last state, and a re-created file arrives as a create event.
 */
function watchForDiskChanges(uri: vscode.Uri, reload: () => Promise<unknown>): vscode.Disposable {
  if (uri.scheme === 'untitled') return { dispose: () => undefined };
  let running: Promise<unknown> = Promise.resolve();
  const scheduled = debounce(() => {
    running = running
      .then(reload)
      .catch((err: unknown) => logError(`Could not reload ${uri.toString(true)} after it changed on disk`, err));
  }, DISK_RELOAD_DEBOUNCE_MS);
  let watcher: vscode.FileSystemWatcher;
  try {
    const folder = vscode.Uri.joinPath(uri, '..');
    watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(folder, '*'), false, false, true);
  } catch (err) {
    logError(`Could not watch ${uri.toString(true)} for changes on disk`, err);
    return scheduled;
  }
  const onEvent = (changed: vscode.Uri): void => {
    if (isSameFile(changed, uri)) scheduled.schedule();
  };
  const listeners = [watcher.onDidCreate(onEvent), watcher.onDidChange(onEvent)];
  return {
    dispose: () => {
      scheduled.dispose();
      for (const listener of listeners) listener.dispose();
      watcher.dispose();
    },
  };
}

// ===== SHARED: READING DOCUMENT FILES =====

/** Every .xlsx / .docx / .pptx is a ZIP package: it starts with a local file header ("PK\x03\x04"). */
function isZipPackage(data: Uint8Array): boolean {
  return data.length >= 4 && data[0] === 0x50 && data[1] === 0x4b && data[2] === 0x03 && data[3] === 0x04;
}

/** A Git version of a file: what VS Code's Git extension writes into a `git:` URI query (toGitUri: path + ref). */
interface GitVersion {
  /** The file in the working tree. */
  readonly workingFile: vscode.Uri;
  /** 'HEAD', a commit, '~' (the staged version), '~1'..'~3' (merge stages) or '' (the index). */
  readonly ref: string;
}

function gitVersionOf(uri: vscode.Uri): GitVersion | undefined {
  if (uri.scheme !== 'git') return undefined;
  try {
    const query: unknown = JSON.parse(uri.query);
    if (!isRecord(query) || typeof query.path !== 'string' || typeof query.ref !== 'string') return undefined;
    if (query.submoduleOf !== undefined) return undefined; // a submodule diff, not a file
    return { workingFile: vscode.Uri.file(query.path), ref: query.ref };
  } catch {
    return undefined;
  }
}

/** The part of VS Code's Git extension API used here (`vscode.git` exports). */
interface GitExtensionExports {
  getAPI(version: 1): { readonly git: { readonly path: string } };
}

/** Upper bound for a document read from Git. */
const MAX_GIT_BLOB_BYTES = 1024 * 1024 * 1024;

function execFileBytes(file: string, args: string[], cwd: string): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    execFile(file, args, { cwd, encoding: 'buffer', maxBuffer: MAX_GIT_BLOB_BYTES, windowsHide: true }, (err, stdout) =>
      err ? reject(err) : resolve(stdout),
    );
  });
}

async function isDirectory(fsPath: string): Promise<boolean> {
  try {
    return (await statPath(fsPath)).isDirectory();
  } catch {
    return false;
  }
}

/**
 * The committed bytes of a Git version, read with `git cat-file blob` (which never converts content) by the git
 * executable of VS Code's Git extension. Undefined when that extension is not available, and in Restricted Mode (the
 * Git extension does not run there either, and git must not run in an untrusted repository).
 */
async function readGitBlob(version: GitVersion): Promise<Uint8Array | undefined> {
  if (!vscode.workspace.isTrusted) return undefined;
  const gitExtension = vscode.extensions.getExtension<GitExtensionExports>('vscode.git');
  let gitPath: string | undefined;
  try {
    gitPath = gitExtension?.isActive ? gitExtension.exports.getAPI(1).git.path : undefined;
  } catch {
    return undefined; // Git is disabled (git.enabled: false)
  }
  if (!gitPath) return undefined;
  // The Git extension's refs: '~' = the index if the file is staged, else HEAD (equal to the index then); '~N' = merge
  // stage N; '' = the index.
  const ref = version.ref;
  const revisions = ref === '~' ? ['', 'HEAD'] : /^~\d$/.test(ref) ? [`:${ref[1]}`] : [ref];
  if (revisions.some((revision) => revision.startsWith('-'))) return undefined; // a ref must never become an option
  // git runs in the nearest folder that still exists: the file's folder may be deleted or renamed in the working tree
  // (Open Changes / Timeline of a deleted file); the path is then relative to that folder, with '/' separators.
  let cwd = path.dirname(version.workingFile.fsPath);
  while (!(await isDirectory(cwd))) {
    const parent = path.dirname(cwd);
    if (parent === cwd) throw new Error(`no folder of ${version.workingFile.fsPath} exists`);
    cwd = parent;
  }
  const name = path.relative(cwd, version.workingFile.fsPath).split(path.sep).join('/');
  let lastError: unknown;
  for (const revision of revisions) {
    try {
      return await execFileBytes(gitPath, ['cat-file', 'blob', `${revision}:./${name}`], cwd);
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError;
}

/**
 * Reads a workbook / document / presentation / PDF; `isExpected` checks that the bytes are of that format. A Git
 * version ("Open Changes", Timeline, "Open File (HEAD)") comes from VS Code's Git extension, which runs
 * `git show --textconv`: when a diff driver converts the file type (Git for Windows maps *.docx and *.pdf to
 * `astextplain`), that is plain text, not the file. Such a version is read as committed instead.
 */
async function readDocumentBytes(uri: vscode.Uri, isExpected: (data: Uint8Array) => boolean): Promise<Uint8Array> {
  const data = await vscode.workspace.fs.readFile(uri);
  const version = gitVersionOf(uri);
  if (!version || isExpected(data)) return data;
  try {
    const blob = await readGitBlob(version);
    if (blob) {
      getLog().info(`${uri.toString(true)}: Git converted the file to text (textconv); read the committed file instead`);
      return blob;
    }
  } catch (err) {
    getLog().warn(`Could not read ${version.workingFile.fsPath} (ref "${version.ref}") from Git: ${errorMessageOf(err)}`);
  }
  return data;
}

// ===== SHEET (XLSX) =====

type Workbook = Awaited<ReturnType<typeof loadWorkbook>>;

const EMPTY_STATS: Readonly<SelectionStats> = Object.freeze({ count: 0, numCount: 0, sum: 0 });

function toUint8Array(data: ArrayBuffer | ArrayBufferView): Uint8Array {
  return ArrayBuffer.isView(data)
    ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
    : new Uint8Array(data);
}

function sameUri(a: vscode.Uri, b: vscode.Uri): boolean {
  return a.toString() === b.toString();
}

function lossyBanner(features: readonly string[], what: 'workbook' | 'presentation' = 'workbook'): string | undefined {
  if (features.length === 0) return undefined;
  return `This ${what} contains content that FileStudio does not display: ${features.join(', ')}.`;
}

/**
 * An .xlsx file opened in the spreadsheet editor. Loading starts on construction and never rejects: failures are
 * kept in `loadError` so the webview can show them. Phase 1 is read-only; phase 2 records edits through `pushEdit`.
 */
export class SheetDocument implements vscode.CustomDocument {
  /** Every webview panel currently showing this document. */
  readonly panels = new Set<vscode.WebviewPanel>();

  private readonly _onDidChange = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SheetDocument>>();
  /** Edit events, forwarded to VS Code via SheetEditorProvider.onDidChangeCustomDocument. */
  readonly onDidChange = this._onDidChange.event;
  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  readonly onDidDispose = this._onDidDispose.event;

  private _loading: Promise<void>;
  private _workbook: Workbook | undefined;
  private _model: WorkbookModel | undefined;
  private _loadError: unknown;
  private _lossyFeatures: string[] = [];

  /** Bumped by every edit / undo / redo; equals savedVersion when the in-memory workbook matches the last save. */
  private changeVersion = 0;
  private savedVersion = 0;
  /** Loaded from a backup or untitled data, so the bytes on disk are not what the workbook holds. */
  private divergedFromDisk = false;
  /** The file as last loaded or saved by us; a watcher event with the same stamp is not an external change. */
  private diskStamp: DiskStamp | undefined;
  /** Our own write to the file, if one is in progress (its watcher events must not trigger a reload). */
  private ownWrite: Promise<void> = Promise.resolve();
  private bakAttempted = false;
  private disposed = false;

  /** Opens `uri`, restoring from the hot-exit backup (falling back to the file) or untitled data when given. */
  static open(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): SheetDocument {
    if (openContext.untitledDocumentData) return new SheetDocument(uri, [openContext.untitledDocumentData]);
    if (openContext.backupId) return new SheetDocument(uri, [vscode.Uri.parse(openContext.backupId), uri]);
    return new SheetDocument(uri, [uri]);
  }

  private constructor(
    readonly uri: vscode.Uri,
    sources: (vscode.Uri | Uint8Array)[],
  ) {
    this._loading = this.load(sources);
  }

  get fileName(): string {
    return fileNameOf(this.uri);
  }

  /** Settles (never rejects) when the current load / reload has finished. */
  get loading(): Promise<void> {
    return this._loading;
  }

  get workbook(): Workbook | undefined {
    return this._workbook;
  }

  get model(): WorkbookModel | undefined {
    return this._model;
  }

  /** Why the workbook could not be loaded (undefined after a successful load). */
  get loadError(): unknown {
    return this._loadError;
  }

  get lossyFeatures(): readonly string[] {
    return this._lossyFeatures;
  }

  get hasUnsavedChanges(): boolean {
    return this.divergedFromDisk || this.changeVersion !== this.savedVersion;
  }

  /**
   * Phase 2 entry point: puts an edit on VS Code's undo stack (marking the document dirty). `undo`/`redo` must
   * update the workbook and notify the webviews themselves.
   */
  pushEdit(edit: { label: string; undo: () => void | Thenable<void>; redo: () => void | Thenable<void> }): void {
    this.changeVersion++;
    this._onDidChange.fire({
      document: this,
      label: edit.label,
      undo: async () => {
        this.changeVersion++;
        await edit.undo();
      },
      redo: async () => {
        this.changeVersion++;
        await edit.redo();
      },
    });
  }

  /** Writes the workbook to its own file. An unmodified workbook is never re-serialized (ExcelJS is lossy). */
  async save(token: vscode.CancellationToken): Promise<void> {
    if (!this.hasUnsavedChanges) {
      getLog().info(`[sheet] Save of ${this.fileName} skipped: no changes, the file on disk is left untouched`);
      return;
    }
    await this.writeBakOnce();
    const write = this.writeWorkbookTo(this.uri, token).then(async () => {
      this.diskStamp = await diskStampOf(this.uri);
    });
    this.ownWrite = write.catch(() => undefined);
    await write;
    this.markSaved();
  }

  async saveAs(target: vscode.Uri, token: vscode.CancellationToken): Promise<void> {
    if (sameUri(target, this.uri)) {
      await this.save(token);
      return;
    }
    await this.writeTo(target, token);
  }

  /** Discards in-memory state and reloads the file from disk. */
  async revert(): Promise<void> {
    this._loading = this.load([this.uri]);
    await this._loading;
    this.changeVersion = 0;
    this.savedVersion = 0;
  }

  /**
   * Called (debounced) by the file watcher: reloads when the file differs from what was last loaded or saved.
   * Unsaved changes are never discarded (phase 2: VS Code reports the conflict on save). Returns whether it reloaded.
   */
  async reloadIfChangedOnDisk(): Promise<boolean> {
    await this._loading;
    await this.ownWrite;
    if (this.disposed || isUnchanged(await diskStampOf(this.uri), this.diskStamp)) return false;
    if (this.hasUnsavedChanges) {
      getLog().warn(`[sheet] ${this.fileName} changed on disk; not reloaded because it has unsaved changes`);
      return false;
    }
    getLog().info(`[sheet] ${this.fileName} changed on disk; reloading`);
    await this.revert();
    return !this.disposed;
  }

  async backup(destination: vscode.Uri, token: vscode.CancellationToken): Promise<vscode.CustomDocumentBackup> {
    await this.writeTo(destination, token);
    return {
      id: destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(destination);
        } catch {
          // Already gone.
        }
      },
    };
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this._onDidDispose.fire();
    this._onDidDispose.dispose();
    this._onDidChange.dispose();
    this.panels.clear();
  }

  private async load(sources: (vscode.Uri | Uint8Array)[]): Promise<void> {
    // Stamp before reading: a write racing the read then shows up as a change at the next watcher event.
    this.diskStamp = await diskStampOf(this.uri);
    try {
      const { data, source } = await readFirstAvailable(sources);
      this.divergedFromDisk = !(source instanceof vscode.Uri && sameUri(source, this.uri));
      let lossyError: unknown;
      // Relative links to other workbooks are shown in the file's folder (Excel's formula bar does the same).
      const folder = this.uri.scheme === 'file' ? path.dirname(this.uri.fsPath) : undefined;
      const [workbook, lossy] = await Promise.all([
        loadWorkbook(data, { folder }),
        detectLossyFeatures(data).catch((err: unknown): string[] => {
          lossyError = err ?? new Error('Unknown error');
          return [];
        }),
      ]);
      // Only worth a log line when the workbook itself loaded (otherwise the load error says it all).
      if (lossyError !== undefined) logError(`[sheet] Could not inspect ${this.fileName} for unsupported features`, lossyError);
      const model = workbookToModel(workbook);
      this._workbook = workbook;
      this._model = model;
      this._lossyFeatures = lossy;
      this._loadError = undefined;
      getLog().info(`[sheet] Loaded ${this.uri.toString(true)}${lossy.length ? ` (not displayed: ${lossy.join(', ')})` : ''}`);
    } catch (err) {
      this._workbook = undefined;
      this._model = undefined;
      this._lossyFeatures = [];
      this._loadError = err ?? new Error('Unknown error');
      logError(`[sheet] Could not load ${this.uri.toString(true)}`, err);
    }
  }

  /** Workbook bytes to `target`: re-serialized when modified, otherwise the original file copied byte for byte. */
  private async writeTo(target: vscode.Uri, token: vscode.CancellationToken): Promise<void> {
    if (this.hasUnsavedChanges) {
      await this.writeWorkbookTo(target, token);
      return;
    }
    const original = await vscode.workspace.fs.readFile(this.uri);
    if (token.isCancellationRequested) throw new vscode.CancellationError();
    await vscode.workspace.fs.writeFile(target, original);
  }

  private async writeWorkbookTo(target: vscode.Uri, token: vscode.CancellationToken): Promise<void> {
    await this._loading;
    const workbook = this._workbook;
    if (!workbook) {
      throw new Error(`"${this.fileName}" could not be loaded, so it cannot be saved: ${errorMessageOf(this._loadError)}`);
    }
    const buffer = await workbook.xlsx.writeBuffer();
    if (token.isCancellationRequested) throw new vscode.CancellationError();
    await vscode.workspace.fs.writeFile(target, toUint8Array(buffer as ArrayBuffer | ArrayBufferView));
  }

  private markSaved(): void {
    this.savedVersion = this.changeVersion;
    this.divergedFromDisk = false;
  }

  /** `fileStudio.xlsx.backupOnSave`: copy the original to `<name>.xlsx.bak` before the first save (never overwritten). */
  private async writeBakOnce(): Promise<void> {
    if (this.bakAttempted) return;
    this.bakAttempted = true;
    if (!vscode.workspace.getConfiguration('fileStudio', this.uri).get<boolean>('xlsx.backupOnSave', true)) return;
    const bak = this.uri.with({ path: `${this.uri.path}.bak` });
    try {
      await vscode.workspace.fs.stat(bak);
      getLog().info(`[sheet] ${displayPath(bak)} already exists; keeping it`);
      return;
    } catch {
      // No backup yet.
    }
    try {
      await vscode.workspace.fs.copy(this.uri, bak, { overwrite: false });
      getLog().info(`[sheet] Wrote backup ${displayPath(bak)}`);
    } catch (err) {
      logError(`[sheet] Could not write backup ${displayPath(bak)}; saving anyway`, err);
    }
  }
}

/** Reads the first source that can be read (e.g. a backup, then the original file). */
async function readFirstAvailable(
  sources: (vscode.Uri | Uint8Array)[],
): Promise<{ data: Uint8Array; source: vscode.Uri | Uint8Array }> {
  let lastError: unknown = new Error('No source to read from');
  for (const source of sources) {
    if (source instanceof Uint8Array) return { data: source, source };
    try {
      return { data: await readDocumentBytes(source, isZipPackage), source };
    } catch (err) {
      lastError = err;
      if (source !== sources[sources.length - 1]) logError(`Could not read ${source.toString(true)}; trying the next source`, err);
    }
  }
  throw lastError;
}

// ----- grid requests (shared by xlsx and csv) -----

/** `loadError`: why there is no `source` (logged by the document / view that failed to load or parse it). */
function answerGetRows(
  session: ViewerSession,
  source: GridSource | undefined,
  msg: MessageOf<'getRows'>,
  loadError?: unknown,
): void {
  if (!source) {
    session.postError(`"${session.fileName}" could not be displayed`, loadError, true);
    return;
  }
  const columns = msg.c0 === undefined ? {} : { c0: msg.c0, c1: msg.c1 };
  const reply = (rows: RowData[], styles: [number, CellStyle][]): void =>
    session.post({ type: 'rows', reqId: msg.reqId, sheet: msg.sheet, start: msg.start, end: msg.end, ...columns, rows, styles });
  try {
    const { rows, styles } = source.getRows(msg.sheet, msg.start, msg.end, msg.c0, msg.c1);
    reply(rows, styles);
  } catch (err) {
    if (isStaleSheetRequest(source, msg.sheet)) {
      // In flight across a reload that removed the sheet: the webview already has the new meta; answer empty.
      getLog().debug(`[${session.kind}] getRows for missing sheet ${msg.sheet} of ${session.fileName} answered empty`);
      reply([], []);
      return;
    }
    session.postError(`Could not read rows ${msg.start + 1}-${msg.end} of sheet ${msg.sheet + 1}`, err);
  }
}

/** Only consulted after a failure, so getMeta() stays off the hot path. */
function isStaleSheetRequest(source: GridSource, sheet: number): boolean {
  try {
    return sheet >= source.getMeta().sheets.length;
  } catch {
    return false;
  }
}

/** Status-bar aggregates. A failure only empties the status bar (logged), it does not replace the grid. */
function answerGetStats(session: ViewerSession, source: GridSource | undefined, msg: MessageOf<'getStats'>): void {
  let stats: SelectionStats = { ...EMPTY_STATS };
  if (source) {
    try {
      stats = source.getStats(msg.sheet, msg.ranges);
    } catch (err) {
      logError(`[${session.kind}] Could not compute selection statistics for ${session.fileName}`, err);
    }
  }
  session.post({ type: 'stats', reqId: msg.reqId, stats });
}

export class SheetEditorProvider implements vscode.CustomEditorProvider<SheetDocument> {
  static readonly viewType = VIEW_TYPES.sheet;

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<SheetDocument>
  >();
  readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  private readonly sessions = new Map<vscode.WebviewPanel, ViewerSession>();

  constructor(private readonly extensionUri: vscode.Uri) {}

  openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken,
  ): SheetDocument {
    const document = SheetDocument.open(uri, openContext);
    const forward = document.onDidChange((e) => this._onDidChangeCustomDocument.fire(e));
    const watcher = watchForDiskChanges(uri, async () => {
      if (await document.reloadIfChangedOnDisk()) this.refreshPanels(document);
    });
    document.onDidDispose(() => {
      forward.dispose();
      watcher.dispose();
    });
    return document;
  }

  resolveCustomEditor(document: SheetDocument, panel: vscode.WebviewPanel, _token: vscode.CancellationToken): void {
    panel.webview.options = webviewOptions(baseResourceRoots(this.extensionUri));
    panel.webview.html = getHtml(panel.webview, this.extensionUri, 'sheet', document.fileName);

    const session: ViewerSession = new ViewerSession(panel, document.uri, 'sheet', (msg) =>
      this.onMessage(document, session, msg),
    );
    document.panels.add(panel);
    this.sessions.set(panel, session);
    // (v0.2.0, F) Edit gate registry: view-only until the xlsx editing phases (P1/P2) give it a real edit state.
    const gate = { format: 'xlsx', uri: document.uri, computeEditState: viewOnlyEditState, dirty: () => document.hasUnsavedChanges } as const;
    session.onDispose(attachEditGateDoc(gate, (state) => session.post({ type: 'editState', state })));
    session.onDispose({
      dispose: () => {
        document.panels.delete(panel);
        this.sessions.delete(panel);
      },
    });
  }

  saveCustomDocument(document: SheetDocument, cancellation: vscode.CancellationToken): Promise<void> {
    return document.save(cancellation);
  }

  saveCustomDocumentAs(
    document: SheetDocument,
    destination: vscode.Uri,
    cancellation: vscode.CancellationToken,
  ): Promise<void> {
    return document.saveAs(destination, cancellation);
  }

  async revertCustomDocument(document: SheetDocument, _cancellation: vscode.CancellationToken): Promise<void> {
    await document.revert();
    this.refreshPanels(document);
    if (document.loadError !== undefined) {
      throw new Error(`Could not reload "${document.fileName}": ${errorMessageOf(document.loadError)}`);
    }
  }

  backupCustomDocument(
    document: SheetDocument,
    context: vscode.CustomDocumentBackupContext,
    cancellation: vscode.CancellationToken,
  ): Promise<vscode.CustomDocumentBackup> {
    return document.backup(context.destination, cancellation);
  }

  private async onMessage(document: SheetDocument, session: ViewerSession, msg: WebviewMessage): Promise<void> {
    switch (msg.type) {
      case 'ready':
        // 'ready' may arrive before or after the load finishes: waiting on the load promise covers both orders.
        await document.loading;
        this.sendInit(document, session);
        return;
      case 'getRows':
        await document.loading;
        answerGetRows(session, document.model, msg, document.loadError);
        return;
      case 'getStats':
        await document.loading;
        answerGetStats(session, document.model, msg);
        return;
      default:
        await handleCommonMessage(session, msg, { allowLocal: true });
    }
  }

  private sendInit(document: SheetDocument, session: ViewerSession): void {
    const model = document.model;
    if (!model) {
      // SheetDocument.load logged the failure once; every panel and every 'ready' only shows it.
      session.postError(`Could not open "${document.fileName}"`, document.loadError, true);
      return;
    }
    try {
      session.post({
        type: 'init',
        kind: 'sheet',
        fileName: document.fileName,
        meta: model.getMeta(),
        readOnly: true,
        banner: lossyBanner(document.lossyFeatures),
        editState: viewOnlyEditState(),
      });
    } catch (err) {
      session.postError(`Could not open "${document.fileName}"`, err);
    }
  }

  /** After a revert or a change on disk: every panel of the document shows the reloaded workbook. */
  private refreshPanels(document: SheetDocument): void {
    for (const panel of document.panels) {
      const session = this.sessions.get(panel);
      if (session) this.refreshAfterReload(document, session);
    }
  }

  /**
   * Grids get 'invalidate' with the new meta (sheet, selection and scroll are kept) and the reloaded file's banner;
   * error views get a fresh init.
   */
  private refreshAfterReload(document: SheetDocument, session: ViewerSession): void {
    if (!session.isReady) return; // the pending 'ready' will send init from the reloaded workbook
    const model = document.model;
    if (!model || session.shown !== 'content') {
      this.sendInit(document, session);
      return;
    }
    try {
      session.post({
        type: 'invalidate',
        meta: model.getMeta(),
        banner: lossyBanner(document.lossyFeatures) ?? null,
        editState: viewOnlyEditState(),
      });
    } catch (err) {
      session.postError(`Could not reload "${document.fileName}"`, err);
    }
  }
}

// ===== TEXT (CSV / MARKDOWN) =====

/** Re-parse / re-render delay after the TextDocument changes. */
const TEXT_DEBOUNCE_MS = 150;

/** Extensions whose relative references in markdown are loaded through the webview (images, media). */
const MARKDOWN_MEDIA_EXTENSIONS: ReadonlySet<string> = new Set([
  '.png', '.jpg', '.jpeg', '.jpe', '.jfif', '.gif', '.svg', '.webp', '.bmp', '.ico', '.avif', '.apng', '.tif', '.tiff',
  '.mp4', '.webm', '.ogg', '.ogv', '.mp3', '.wav', '.m4a', '.mov', '.vtt',
]);

/** Content changes of one TextDocument (dirty-flag-only events, e.g. after save, are skipped). */
function onTextDocumentChanged(document: vscode.TextDocument, listener: () => void): vscode.Disposable {
  const key = document.uri.toString();
  return vscode.workspace.onDidChangeTextDocument((e) => {
    if (e.contentChanges.length > 0 && e.document.uri.toString() === key) listener();
  });
}

function textKindFor(uri: vscode.Uri, fallback: 'csv' | 'markdown'): 'csv' | 'markdown' {
  const kind = viewKindForUri(uri);
  return kind === 'csv' || kind === 'markdown' ? kind : fallback;
}

/**
 * Serves `.csv`/`.tsv`/`.psv`/`.ssv` (kind 'csv') and `.md` (kind 'markdown') files; the kind follows the file
 * extension, with the kind this instance was registered for as fallback. The TextDocument stays the source of truth.
 */
export class TextViewerProvider implements vscode.CustomTextEditorProvider {
  constructor(
    private readonly extensionUri: vscode.Uri,
    private readonly defaultKind: 'csv' | 'markdown',
  ) {}

  resolveCustomTextEditor(
    document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
    _token: vscode.CancellationToken,
  ): void {
    const kind = textKindFor(document.uri, this.defaultKind);
    const roots =
      kind === 'markdown' ? markdownResourceRoots(this.extensionUri, document.uri) : baseResourceRoots(this.extensionUri);
    panel.webview.options = webviewOptions(roots);
    panel.webview.html = getHtml(panel.webview, this.extensionUri, kind, fileNameOf(document.uri));
    if (kind === 'csv') new CsvView(document, panel);
    else new MarkdownView(document, panel);
  }
}

// ----- text files VS Code does not pass to a custom text editor -----

/**
 * VS Code passes a text document to extensions only up to 50 MB of text (TextModel._MODEL_SYNC_LIMIT, in UTF-16 code
 * units). On a larger file a custom text editor is never resolved and VS Code shows "The editor could not be opened
 * due to an unexpected error". UTF-8 needs at least one byte per code unit, so smaller files are always synced.
 */
const EXTENSION_SYNC_LIMIT = 50 * 1024 * 1024;

/**
 * Replaces FileStudio tabs (CSV/TSV/PSV/SSV/Markdown) whose file VS Code cannot give the extension as a text document
 * with VS Code's text editor, and says why: files over the 50 MB extension limit, and files VS Code considers binary
 * (a NUL byte near the start, e.g. a workbook saved as .csv). Our tab would only show VS Code's "The editor could not
 * be opened due to an unexpected error"; the text editor explains the file and offers "Open Anyway". Also checks the
 * tabs already open at activation (opening such a tab is what activates us).
 */
export function watchUnreadableTextTabs(): vscode.Disposable {
  /** Files being checked -> whether another event for them came meanwhile (a click: "opened", then "changed" active). */
  const checking = new Map<string, boolean>();
  const check = (tab: vscode.Tab): void => {
    const input = tab.input;
    if (!(input instanceof vscode.TabInputCustom)) return;
    if (input.viewType !== VIEW_TYPES.csv && input.viewType !== VIEW_TYPES.markdown) return;
    const key = input.uri.toString();
    if (checking.has(key)) {
      checking.set(key, true);
      return;
    }
    checking.set(key, false);
    void reopenIfUnreadable(input.uri, tab).finally(() => {
      const again = checking.get(key);
      checking.delete(key);
      if (again && tab.group.tabs.includes(tab)) check(tab);
    });
  };
  for (const group of vscode.window.tabGroups.all) for (const tab of group.tabs) check(tab);
  return vscode.window.tabGroups.onDidChangeTabs((e) => {
    for (const tab of [...e.opened, ...e.changed]) check(tab);
  });
}

/** VS Code's reason from an openTextDocument failure ("cannot open <uri>. Detail: <reason>"). */
function openFailureReason(err: unknown): string {
  const message = errorMessageOf(err);
  return /\bDetail: (.+)$/s.exec(message)?.[1].trim() || message;
}

async function reopenIfUnreadable(uri: vscode.Uri, tab: vscode.Tab): Promise<void> {
  const key = uri.toString();
  if (vscode.workspace.textDocuments.some((d) => d.uri.toString() === key)) return; // synced: our editor works
  const stamp = await diskStampOf(uri);
  if (!stamp) return;
  const tooLarge = stamp.size > EXTENSION_SYNC_LIMIT;
  // VS Code resolves only the editor a group shows, so only that tab is probed (again when a tab becomes active);
  // a tab of a file over the size limit always is. Tab properties are live: read after the await.
  if (!tooLarge && !tab.isActive) return;
  const size = `${(stamp.size / (1024 * 1024)).toFixed(1)} MB`;
  let reason: string;
  try {
    // Throws exactly when the custom editor cannot work: "Files above 50MB cannot be synchronized with extensions",
    // "File seems to be binary and cannot be opened as text".
    await vscode.workspace.openTextDocument(uri);
    return; // a text document (also: over 50 MB on disk, but fewer than 50 M characters once decoded)
  } catch (err) {
    reason = openFailureReason(err);
    getLog().warn(`${uri.toString(true)} (${size}) cannot be shown by FileStudio: ${errorMessageOf(err)}`);
  }
  await reopenAsText(uri, tab.group.viewColumn);
  void vscode.window.showWarningMessage(
    tooLarge
      ? `"${fileNameOf(uri)}" (${size}) is too large for FileStudio: VS Code passes text files of up to 50 MB to ` +
          'extensions. It was opened in the text editor instead.'
      : `FileStudio cannot show "${fileNameOf(uri)}": VS Code does not open it as text (${reason}). It was opened ` +
          'in the text editor instead.',
  );
}

// ----- csv / tsv / psv / ssv -----

/** Delimited-text format of a file, by extension (any other file served as kind 'csv' is read as CSV). */
function delimitedFormatOf(uri: vscode.Uri): DelimitedFormat {
  switch (extensionOf(uri)) {
    case '.tsv':
      return 'tsv';
    case '.psv':
      return 'psv';
    case '.ssv':
      return 'ssv';
    default:
      return 'csv';
  }
}

/**
 * Parses delimited text: TSV on tabs, PSV on '|', SSV on ';' or on runs of spaces (whichever detectSsvDelimiter finds
 * in the text, so an edit can switch it), CSV on the delimiter the parser detects. A `chosen` delimiter (Change
 * Delimiter) wins over all of these.
 */
function parseDelimited(text: string, format: DelimitedFormat, chosen?: string): CsvModel {
  // forView: the view never writes the text back, so the parser skips what only saving needs
  if (chosen !== undefined) {
    return parseCsv(text, chosen === ' ' ? { delimiter: ' ', collapseSpaces: true, forView: true } : { delimiter: chosen, forView: true });
  }
  switch (format) {
    case 'tsv':
      return parseCsv(text, { delimiter: '\t', forView: true });
    case 'psv':
      return parseCsv(text, { delimiter: '|', forView: true });
    case 'ssv':
      return detectSsvDelimiter(text) === ';'
        ? parseCsv(text, { delimiter: ';', forView: true })
        : parseCsv(text, { delimiter: ' ', collapseSpaces: true, forView: true });
    default:
      return parseCsv(text, { forView: true });
  }
}

/** Delimiters chosen with `fileStudio.changeDelimiter`, by document URI, for this session (none = detect / by format). */
const chosenDelimiters = new Map<string, string>();
/** CsvViews to re-parse when a delimiter is chosen (called with the document URI string). */
const chosenDelimiterListeners = new Set<(key: string) => void>();

const DELIMITER_CHOICES: readonly { label: string; description: string; delimiter: string | undefined }[] = [
  { label: 'Detect automatically', description: 'by file type and content', delimiter: undefined },
  { label: 'Comma', description: ',', delimiter: ',' },
  { label: 'Semicolon', description: ';', delimiter: ';' },
  { label: 'Tab', description: '\\t', delimiter: '\t' },
  { label: 'Pipe', description: '|', delimiter: '|' },
  { label: 'Spaces', description: 'runs of spaces', delimiter: ' ' },
];

/**
 * `fileStudio.changeDelimiter`: a quick pick of delimiters; every FileStudio view of the file re-parses with the
 * choice. It lasts until VS Code is closed; the file itself is not changed.
 */
export async function changeDelimiter(uri: vscode.Uri): Promise<void> {
  const key = uri.toString();
  const current = chosenDelimiters.get(key);
  const pick = await vscode.window.showQuickPick(
    DELIMITER_CHOICES.map((choice) => ({
      ...choice,
      detail: choice.delimiter === current ? 'Current choice' : undefined,
    })),
    { placeHolder: `Split the columns of "${fileNameOf(uri)}" on` },
  );
  if (!pick || pick.delimiter === current) return;
  if (pick.delimiter === undefined) chosenDelimiters.delete(key);
  else chosenDelimiters.set(key, pick.delimiter);
  for (const listener of chosenDelimiterListeners) listener(key);
}

/**
 * One CSV/TSV/PSV/SSV panel: parses on open, re-parses (debounced) when the text changes, answers grid requests. The
 * delimiter is chosen again on every parse (CSV and SSV detect it); when it changes, the webview gets a fresh init.
 */
class CsvView {
  private readonly session: ViewerSession;
  private readonly reparse: Debounced;
  private readonly format: DelimitedFormat;
  private model: CsvGridModel | undefined;
  /** Delimiter of `model` (' ' = runs of spaces). */
  private delimiter = '';
  /** Delimiter announced in the last init: the webview shows it, so a different one needs a new init. */
  private initDelimiter: string | undefined;
  private parseError: unknown;
  private parsedVersion = -1;
  /** The text changed while the panel was hidden: re-parsed when it is shown again. */
  private stale = false;

  constructor(
    private readonly document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
  ) {
    this.format = delimitedFormatOf(document.uri);
    this.session = new ViewerSession(panel, document.uri, 'csv', (msg) => this.onMessage(msg));
    // (v0.2.0, F) Edit gate registry: view-only until P3 (CsvDocState) gives it a real edit state.
    const gate = { format: 'csv', uri: document.uri, computeEditState: viewOnlyEditState, dirty: () => document.isDirty } as const;
    this.session.onDispose(attachEditGateDoc(gate, (state) => this.session.post({ type: 'editState', state })));
    this.reparse = debounce(() => this.refresh(), TEXT_DEBOUNCE_MS);
    this.session.onDispose(this.reparse);
    this.session.onDispose(onTextDocumentChanged(document, () => this.reparse.schedule()));
    // A big file blocks the extension host for seconds per parse: a hidden panel parses when it is shown again.
    this.session.onDispose(
      panel.onDidChangeViewState(() => {
        if (panel.visible && this.stale) this.refresh();
      }),
    );
    // "Change Delimiter" for this file: parse again with the chosen delimiter.
    const onChosenDelimiter = (key: string): void => {
      if (key !== document.uri.toString()) return;
      this.parsedVersion = -1;
      this.refresh();
    };
    chosenDelimiterListeners.add(onChosenDelimiter);
    this.session.onDispose({ dispose: () => chosenDelimiterListeners.delete(onChosenDelimiter) });
    // Parse right after the HTML is handed over, so the model is usually ready when the webview asks for it.
    const initialParse = setTimeout(() => this.ensureParsed(), 0);
    this.session.onDispose({ dispose: () => clearTimeout(initialParse) });
  }

  private ensureParsed(): void {
    if (this.session.isDisposed || this.parsedVersion === this.document.version) return;
    this.parsedVersion = this.document.version;
    try {
      const csv = parseDelimited(this.document.getText(), this.format, chosenDelimiters.get(this.document.uri.toString()));
      this.model = new CsvGridModel(csv, this.session.fileName);
      this.delimiter = csv.delimiter;
      this.parseError = undefined;
    } catch (err) {
      this.model = undefined;
      this.parseError = err ?? new Error('Unknown error');
      logError(`[csv] Could not parse ${this.document.uri.toString(true)}`, err);
    }
  }

  private onMessage(msg: WebviewMessage): Promise<void> | void {
    switch (msg.type) {
      case 'ready':
        this.ensureParsed();
        this.sendInit();
        return;
      case 'getRows':
        // Served from the current model; a pending re-parse is followed by 'invalidate'.
        answerGetRows(this.session, this.model, msg, this.parseError);
        return;
      case 'getStats':
        answerGetStats(this.session, this.model, msg);
        return;
      default:
        return handleCommonMessage(this.session, msg, { allowLocal: true });
    }
  }

  private sendInit(): void {
    if (!this.model) {
      this.session.postError(`Could not parse "${this.session.fileName}"`, this.parseError, true); // logged by ensureParsed
      return;
    }
    try {
      this.session.post({
        type: 'init',
        kind: 'csv',
        fileName: this.session.fileName,
        meta: this.model.getMeta(),
        readOnly: true,
        delimited: { format: this.format, delimiter: this.delimiter },
        editState: viewOnlyEditState(),
      });
      this.initDelimiter = this.delimiter;
    } catch (err) {
      this.session.postError(`Could not display "${this.session.fileName}"`, err);
    }
  }

  /** Debounced: the text changed (typing in a text editor, undo, external reload). */
  private refresh(): void {
    this.stale = !this.session.panel.visible;
    if (this.stale) return;
    this.ensureParsed();
    if (!this.session.isReady) return; // 'ready' will send init from the current text
    if (!this.model || this.session.shown !== 'content' || this.delimiter !== this.initDelimiter) {
      this.sendInit();
      return;
    }
    try {
      this.session.post({ type: 'invalidate', meta: this.model.getMeta(), editState: viewOnlyEditState() });
    } catch (err) {
      this.session.postError(`Could not display "${this.session.fileName}"`, err);
    }
  }
}

// ----- markdown -----

/** A task marker as the renderer records its position (TaskMarker.column is the character between the brackets). */
const TASK_MARKER_TEXT_RE = /^\[[ xX]\]$/;

function markdownMode(uri: vscode.Uri): MarkdownMode {
  const mode = vscode.workspace.getConfiguration('fileStudio', uri).get<string>('markdown.defaultMode');
  return mode === 'split' || mode === 'wysiwyg' ? mode : 'preview';
}

/** Settings that make VS Code treat a file as read-only (re-checked when one of them changes). */
const READONLY_SETTINGS = ['files.readonlyInclude', 'files.readonlyExclude', 'files.readonlyFromPermissions'];

/** dist + media, plus the document's folder and all workspace folders (for relative images). */
function markdownResourceRoots(extensionUri: vscode.Uri, documentUri: vscode.Uri): vscode.Uri[] {
  const roots = baseResourceRoots(extensionUri);
  if (documentUri.scheme !== 'untitled') {
    roots.push(vscode.Uri.joinPath(documentUri.with({ query: '', fragment: '' }), '..'));
  }
  for (const folder of vscode.workspace.workspaceFolders ?? []) roots.push(folder.uri);
  return roots;
}

/**
 * RenderOptions.resolveResource for markdown. Images and other media (judged by extension) become webview URIs
 * (query/fragment dropped, path decoded, resolved against the document). Everything else — links to documents,
 * folders, other files — is left unchanged (undefined): the webview posts 'openLink' with the original href.
 */
function resolveMarkdownResource(href: string, documentUri: vscode.Uri, webview: vscode.Webview): string | undefined {
  const scheme = linkScheme(href);
  if (scheme === 'file') {
    try {
      const uri = vscode.Uri.parse(href, true).with({ query: '', fragment: '' });
      return MARKDOWN_MEDIA_EXTENSIONS.has(path.posix.extname(uri.path).toLowerCase())
        ? webview.asWebviewUri(uri).toString()
        : undefined;
    } catch {
      return undefined;
    }
  }
  if (scheme !== undefined) return undefined; // absolute URL / data: / other schemes are not ours to rewrite
  const { path: rawPath } = splitHref(href);
  if (!MARKDOWN_MEDIA_EXTENSIONS.has(path.posix.extname(safeDecode(rawPath)).toLowerCase())) return undefined;
  const target = resolveLocalPath(rawPath, documentUri);
  return target ? webview.asWebviewUri(target).toString() : undefined;
}

/** One Markdown preview panel: renders on 'ready', re-renders (debounced) on change, applies task toggles. */
class MarkdownView {
  private readonly session: ViewerSession;
  private readonly rerender: Debounced;
  /** Document version our own pending task toggle will produce (that change is pushed at once, never re-inited). */
  private selfEditVersion: number | undefined;
  /**
   * Task checkboxes of the last rendering sent to the webview, by line, and the version it was made from. The webview
   * only knows versions it got from us, so a toggle for the current version is checked against exactly what it showed.
   */
  private rendered: { version: number; tasks: Map<number, TaskMarker> } | undefined;
  /**
   * (v0.2.0) Whether FileStudio may change the document (SHARED: EDIT GATE: kill switch, VS Code's read-only state).
   * Not editable: task toggles are refused and the webview disables them (`readOnly` = !editable).
   */
  private editState: EditState = { editable: true, kinds: [] };
  /** The document's edit gate entry, shared by all its panels (queue, state, settings listener). */
  private readonly gate: EditGateHandle;
  /** The first edit-state check; 'ready' waits for it, so the init already says whether the checkboxes work. */
  private readonly firstEditStateCheck: Promise<EditState>;
  /** Set when the first check is done: later changes of the state go to the webview at once. */
  private editStateKnown = false;

  constructor(
    private readonly document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
  ) {
    this.session = new ViewerSession(panel, document.uri, 'markdown', (msg) => this.onMessage(msg));
    this.rerender = debounce(() => this.pushUpdate(), TEXT_DEBOUNCE_MS);
    this.session.onDispose(this.rerender);
    this.session.onDispose(onTextDocumentChanged(document, () => this.onDocumentChanged()));
    // (v0.2.0) Kill switch + read-only through the edit gate; its one settings listener re-checks every open document.
    this.gate = attachEditGateDoc(
      {
        format: 'markdown',
        uri: document.uri,
        computeEditState: async () => editStateOf(await checkEditable('markdown', document.uri, {}, document), []),
        version: () => document.version,
        dirty: () => document.isDirty,
      },
      (state) => this.setEditState(state, this.editStateKnown),
    );
    this.session.onDispose(this.gate);
    this.session.editHandlers = { version: () => document.version };
    this.firstEditStateCheck = this.refreshEditState(false);
  }

  private onMessage(msg: WebviewMessage): Promise<void> | void {
    switch (msg.type) {
      case 'ready':
        return this.firstEditStateCheck.then(() => this.sendInit());
      case 'toggleTask':
        return this.toggleTask(msg);
      default:
        // A link to this very document (readme.md#usage) scrolls the preview instead of opening the source.
        return handleCommonMessage(this.session, msg, {
          allowLocal: true,
          onSameDocument: (fragment) => this.session.post({ type: 'scrollToFragment', fragment }),
        });
    }
  }

  private render(): (RenderResult & { source: string; version: number }) | undefined {
    const source = this.document.getText();
    const version = this.document.version;
    try {
      const result = renderMarkdown(source, {
        resolveResource: (href: string) => resolveMarkdownResource(href, this.document.uri, this.session.webview),
      });
      this.rendered = { version, tasks: new Map((result.tasks ?? []).map((task) => [task.line, task])) };
      return { html: result.html, toc: result.toc, source, version };
    } catch (err) {
      this.session.postError(`Could not render "${this.session.fileName}"`, err);
      return undefined;
    }
  }

  private sendInit(): void {
    const rendered = this.render();
    if (!rendered) return;
    this.session.post({
      type: 'init',
      kind: 'markdown',
      fileName: this.session.fileName,
      html: rendered.html,
      toc: rendered.toc,
      source: rendered.source,
      mode: markdownMode(this.document.uri),
      version: rendered.version,
      readOnly: !this.editState.editable,
      editState: this.editState,
    });
  }

  /**
   * Checks again whether FileStudio may change the document (kill switch, read-only; SHARED: EDIT GATE). With `push`,
   * a change goes to the webview at once (a settings change, a file made read-only since the last check), so its task
   * checkboxes follow. Resolves with the state just computed; never rejects.
   */
  private async refreshEditState(push: boolean): Promise<EditState> {
    const state = await this.gate.doc.refresh();
    this.setEditState(this.gate.doc.state ?? state, push);
    this.editStateKnown = true;
    return state;
  }

  private setEditState(state: EditState, push: boolean): void {
    if (JSON.stringify(state) === JSON.stringify(this.editState)) return;
    const wasEditable = this.editState.editable;
    this.editState = state;
    if (state.editable !== wasEditable) {
      const name = this.session.fileName;
      getLog().info(
        state.editable
          ? `[markdown] ${name} can be changed again (no longer read-only)`
          : state.code === 'readOnly'
            ? `[markdown] ${name} is read-only: its task checkboxes are disabled`
            : `[markdown] ${name}: ${state.reason ?? 'not editable'} Its task checkboxes are disabled.`,
      );
    }
    if (push) this.pushUpdate();
  }

  /** Sends the current rendering: a cheap markdownUpdate when content is shown, otherwise a full init. */
  private pushUpdate(): void {
    if (!this.session.isReady) return; // 'ready' will render the current text
    if (this.session.shown !== 'content') {
      this.sendInit();
      return;
    }
    const rendered = this.render();
    if (!rendered) return;
    this.session.post({
      type: 'markdownUpdate',
      html: rendered.html,
      toc: rendered.toc,
      source: rendered.source,
      version: rendered.version,
      readOnly: !this.editState.editable,
      editState: this.editState,
    });
  }

  private onDocumentChanged(): void {
    if (this.selfEditVersion !== undefined && this.document.version === this.selfEditVersion) {
      // Our own toggle: the checkbox is already in its new state; send the new version right away so a quick second
      // toggle is not rejected as stale.
      this.selfEditVersion = undefined;
      this.rerender.cancel();
      this.pushUpdate();
      return;
    }
    this.rerender.schedule();
  }

  /**
   * Applies a task toggle and always answers with toggleTaskResult. Only a checkbox the renderer produced for the
   * version on screen can edit the document, and only the marker character the renderer located for it: never a
   * line guessed by a pattern, nor a checkbox written in raw HTML.
   */
  private toggleTask(msg: MessageOf<'toggleTask'>): Promise<void> {
    // (v0.2.0) In the document's queue: toggles from all its panels run one at a time (contract §3.1).
    return this.gate.doc.queue.run(() => this.applyToggle(msg));
  }

  private async applyToggle(msg: MessageOf<'toggleTask'>): Promise<void> {
    const document = this.document;
    const where = `line ${msg.line + 1} of ${this.session.fileName}`;
    // Checked again at the click (SHARED: EDIT GATE): the kill switch, then VS Code's read-only state (the file may
    // have been made read-only since: VS Code's own editor refuses edits to a read-only document, and an edit applied
    // anyway would change it in memory only, not dirty, never saved). A refusal answers here, never with the error view.
    const state = await this.refreshEditState(true);
    if (!state.editable) {
      const reason = state.code === 'readOnly' ? 'the document is read-only' : (state.reason ?? EDIT_TEXT.internal);
      getLog().info(`[markdown] Task toggle on ${where} refused: ${reason}`);
      this.answerToggle(msg, false, reason);
      return;
    }
    // No await from here to the applyEdit call: the version check must hold when the edit is made (contract §3.2).
    if (msg.version !== document.version) {
      // The webview clicked an older rendering; the change already scheduled the one it is missing.
      getLog().info(`[markdown] Task toggle on ${where} ignored: the document changed (v${msg.version} -> v${document.version})`);
      this.answerToggle(msg, false, 'the document changed since the click');
      return;
    }
    const task = this.rendered?.version === msg.version ? this.rendered.tasks.get(msg.line) : undefined;
    const text = task && msg.line < document.lineCount ? document.lineAt(msg.line).text : '';
    if (!task || !TASK_MARKER_TEXT_RE.test(text.slice(task.column - 1, task.column + 2))) {
      getLog().warn(`[markdown] Task toggle on ${where} ignored: the preview shows no task there`);
      this.answerToggle(msg, false, 'no task checkbox on that line');
      return;
    }
    if ((text[task.column] !== ' ') === msg.checked) {
      this.answerToggle(msg, true); // already in the requested state
      return;
    }
    const edit = new vscode.WorkspaceEdit();
    edit.replace(document.uri, new vscode.Range(msg.line, task.column, msg.line, task.column + 1), msg.checked ? 'x' : ' ');
    this.selfEditVersion = document.version + 1;
    let applied = false;
    try {
      applied = await vscode.workspace.applyEdit(edit);
    } catch (err) {
      logError(`[markdown] Could not toggle the task on ${where}`, err);
    }
    if (!applied) {
      this.selfEditVersion = undefined;
      getLog().warn(`[markdown] Task toggle on ${where} was not applied`);
      this.answerToggle(msg, false, 'the edit could not be applied');
      return;
    }
    this.answerToggle(msg, true); // the new rendering went out with the document change (onDocumentChanged)
  }

  private answerToggle(msg: MessageOf<'toggleTask'>, applied: boolean, reason?: string): void {
    this.session.post({
      type: 'toggleTaskResult',
      line: msg.line,
      version: msg.version,
      applied,
      documentVersion: this.document.version,
      ...(reason === undefined ? {} : { reason }),
    });
  }
}

// ===== DOCX =====

type DocxOutcome = { ok: true; result: DocxResult } | { ok: false; error: unknown };

async function renderDocxFile(uri: vscode.Uri): Promise<DocxOutcome> {
  try {
    const data = await readDocumentBytes(uri, isZipPackage);
    const result = await renderDocx(data);
    if (data.length === 0) {
      getLog().info(`[docx] ${uri.toString(true)} is empty (0 bytes): shown as an empty document, as Word opens it`);
    }
    if (result.warnings.length > 0) {
      getLog().info(`[docx] ${result.warnings.length} conversion warning(s) for ${uri.toString(true)}`);
    }
    return { ok: true, result };
  } catch (err) {
    return { ok: false, error: err ?? new Error('Unknown error') };
  }
}

/**
 * A .docx opened in the preview. The file is converted once and shared by all its panels, and converted again when
 * another program changes it (the panels then get a fresh init).
 */
export class DocxDocument implements vscode.CustomDocument {
  private _rendering: Promise<DocxOutcome>;
  /** The file as last converted; a watcher event with the same stamp is not a change. */
  private diskStamp: DiskStamp | undefined;
  private disposed = false;
  private readonly watcher: vscode.Disposable;
  private readonly _onDidReload = new vscode.EventEmitter<void>();
  /** Fired when `rendering` holds a new conversion after a change on disk. */
  readonly onDidReload = this._onDidReload.event;

  constructor(readonly uri: vscode.Uri) {
    this._rendering = this.render();
    this.watcher = watchForDiskChanges(uri, () => this.reloadIfChangedOnDisk());
  }

  /** The current conversion (settles, never rejects). */
  get rendering(): Promise<DocxOutcome> {
    return this._rendering;
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.watcher.dispose();
    this._onDidReload.dispose();
  }

  private async render(): Promise<DocxOutcome> {
    this.diskStamp = await diskStampOf(this.uri); // before reading, like SheetDocument.load
    return renderDocxFile(this.uri);
  }

  private async reloadIfChangedOnDisk(): Promise<void> {
    await this._rendering;
    if (this.disposed || isUnchanged(await diskStampOf(this.uri), this.diskStamp)) return;
    getLog().info(`[docx] ${fileNameOf(this.uri)} changed on disk; reloading`);
    this._rendering = this.render();
    await this._rendering;
    if (!this.disposed) this._onDidReload.fire();
  }
}

/** Read-only .docx preview: the document's conversion is sent when the webview is ready, and again after a reload. */
export class DocxViewerProvider implements vscode.CustomReadonlyEditorProvider<DocxDocument> {
  static readonly viewType = VIEW_TYPES.docx;

  constructor(private readonly extensionUri: vscode.Uri) {}

  openCustomDocument(uri: vscode.Uri, _openContext: vscode.CustomDocumentOpenContext): DocxDocument {
    return new DocxDocument(uri);
  }

  resolveCustomEditor(document: DocxDocument, panel: vscode.WebviewPanel, _token: vscode.CancellationToken): void {
    const fileName = fileNameOf(document.uri);
    panel.webview.options = webviewOptions(baseResourceRoots(this.extensionUri));
    panel.webview.html = getHtml(panel.webview, this.extensionUri, 'docx', fileName);

    const sendContent = async (): Promise<void> => {
      const rendering = document.rendering;
      const outcome = await rendering;
      if (rendering !== document.rendering) return; // superseded by a newer conversion, which is sent instead
      if (outcome.ok) {
        session.post({ type: 'init', kind: 'docx', fileName, html: outcome.result.html, warnings: outcome.result.warnings });
      } else {
        session.postError(`Could not open "${fileName}"`, outcome.error);
      }
    };
    // Word documents link to other files with absolute (file:///C:\...) or relative paths: they open in VS Code like
    // Markdown links (same checks: no command:/other schemes, directories revealed); a link to the document itself
    // (self.docx#bookmark) scrolls the page.
    const session: ViewerSession = new ViewerSession(panel, document.uri, 'docx', async (msg) => {
      if (msg.type === 'ready') await sendContent();
      else {
        await handleCommonMessage(session, msg, {
          allowLocal: true,
          onSameDocument: (fragment) => session.post({ type: 'scrollToFragment', fragment }),
        });
      }
    });
    session.onDispose(
      document.onDidReload(() => {
        if (session.isReady) void sendContent(); // otherwise the pending 'ready' sends the new conversion
      }),
    );
  }
}

// ===== READ-ONLY FILE DOCUMENTS (PDF / PPTX) =====

/**
 * A read-only file shown by PdfViewerProvider / PptxViewerProvider. `loader` turns the file into what the panels show
 * and must not reject (failures are part of its result). The result is shared by all panels of the document and made
 * again when another program changes the file (onDidReload; the panels then get a fresh init), like DocxDocument.
 */
export class ReadonlyFileDocument<T> implements vscode.CustomDocument {
  private _loading: Promise<T>;
  /** The file as last loaded; a watcher event with the same stamp is not a change. */
  private diskStamp: DiskStamp | undefined;
  private disposed = false;
  private readonly watcher: vscode.Disposable;
  private readonly _onDidReload = new vscode.EventEmitter<void>();
  /** Fired when `loading` holds a new result after a change on disk. */
  readonly onDidReload = this._onDidReload.event;

  constructor(
    readonly uri: vscode.Uri,
    private readonly kind: ViewKind,
    private readonly loader: (uri: vscode.Uri) => Promise<T>,
  ) {
    this._loading = this.load();
    this.watcher = watchForDiskChanges(uri, () => this.reloadIfChangedOnDisk());
  }

  /** The current load (settles, never rejects). */
  get loading(): Promise<T> {
    return this._loading;
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.watcher.dispose();
    this._onDidReload.dispose();
  }

  private async load(): Promise<T> {
    this.diskStamp = await diskStampOf(this.uri); // before reading, like SheetDocument.load
    return this.loader(this.uri);
  }

  private async reloadIfChangedOnDisk(): Promise<void> {
    await this._loading;
    if (this.disposed || isUnchanged(await diskStampOf(this.uri), this.diskStamp)) return;
    getLog().info(`[${this.kind}] ${fileNameOf(this.uri)} changed on disk; reloading`);
    this._loading = this.load();
    await this._loading;
    if (!this.disposed) this._onDidReload.fire();
  }
}

/**
 * Sends a read-only document's current result when the webview is ready and again after each reload: `send` gets the
 * result of the newest load only (a result superseded while it was awaited is dropped; the newer one follows).
 */
function sendOnReadyAndReload<T>(
  session: ViewerSession,
  document: ReadonlyFileDocument<T>,
  send: (result: T) => void,
): () => Promise<void> {
  const sendContent = async (): Promise<void> => {
    const loading = document.loading;
    const result = await loading;
    if (loading === document.loading) send(result);
  };
  session.onDispose(
    document.onDidReload(() => {
      if (session.isReady) void sendContent(); // otherwise the pending 'ready' sends the new result
    }),
  );
  return sendContent;
}

// ===== PDF =====

/** What a PDF panel is initialised from, or why the file cannot be shown. */
type PdfOutcome =
  | {
      ok: true;
      fileSize: number;
      /** PDF version from the header ('1.7'), if readable. */
      version: string | undefined;
      /** The file's bytes, for resources that are not local files (git:, virtual file systems). */
      data: Uint8Array | undefined;
    }
  | {
      ok: false;
      /** Why the bytes are not a PDF (checkPdf), shown as is; otherwise `error` (the file could not be read). */
      reason?: string;
      error?: unknown;
    };

/** The first `length` bytes of a local file (fewer when it is shorter) and its size, without reading the rest. */
async function readFileHead(fsPath: string, length: number): Promise<{ head: Uint8Array; size: number }> {
  const handle = await openFileHandle(fsPath, 'r');
  try {
    const { size } = await handle.stat();
    const buffer = new Uint8Array(Math.min(length, size));
    const { bytesRead } = await handle.read(buffer, 0, buffer.length, 0);
    return { head: buffer.subarray(0, bytesRead), size };
  } finally {
    await handle.close();
  }
}

/**
 * Checks a PDF's header. A local file is only opened for its first bytes: pdf.js fetches it from its webview URI with
 * range requests, so a large file is never read whole by the extension host. Other resources are read whole (their
 * bytes go to the webview).
 */
async function loadPdfFile(uri: vscode.Uri): Promise<PdfOutcome> {
  try {
    let head: Uint8Array;
    let fileSize: number;
    let data: Uint8Array | undefined;
    if (uri.scheme === 'file') {
      ({ head, size: fileSize } = await readFileHead(uri.fsPath, PDF_HEADER_SEARCH_BYTES));
    } else {
      const isPdf = (bytes: Uint8Array): boolean => checkPdf(bytes.subarray(0, PDF_HEADER_SEARCH_BYTES)).error === undefined;
      data = await readDocumentBytes(uri, isPdf);
      // The bytes go to the webview in init. A Node Buffer (what git, and so the Git file system, returns) would arrive
      // as { type: 'Buffer', data: [...] }: VS Code's postMessage serializer meets Buffer.toJSON before its typed-array
      // handling. A plain Uint8Array (a copy) is transferred as bytes.
      if (Buffer.isBuffer(data)) data = new Uint8Array(data);
      head = data.subarray(0, PDF_HEADER_SEARCH_BYTES);
      fileSize = data.byteLength;
    }
    const check = checkPdf(head);
    if (check.error !== undefined) return { ok: false, reason: check.error };
    getLog().info(`[pdf] Loaded ${uri.toString(true)} (PDF ${check.version ?? 'version unknown'}, ${fileSize} bytes)`);
    return { ok: true, fileSize, version: check.version, data };
  } catch (err) {
    return { ok: false, error: err ?? new Error('Unknown error') };
  }
}

/**
 * Read-only PDF view. The host checks the header; pdf.js loads and renders the file in the webview, from its webview
 * URI for local files (the file's folder is a resource root of that panel), otherwise from the bytes sent in init.
 * Links in the PDF that leave the document arrive as openLink (the docx policy: web/mail links and local files).
 */
export class PdfViewerProvider implements vscode.CustomReadonlyEditorProvider<ReadonlyFileDocument<PdfOutcome>> {
  static readonly viewType = VIEW_TYPES.pdf;

  constructor(private readonly extensionUri: vscode.Uri) {}

  openCustomDocument(uri: vscode.Uri, _openContext: vscode.CustomDocumentOpenContext): ReadonlyFileDocument<PdfOutcome> {
    return new ReadonlyFileDocument(uri, 'pdf', loadPdfFile);
  }

  resolveCustomEditor(
    document: ReadonlyFileDocument<PdfOutcome>,
    panel: vscode.WebviewPanel,
    _token: vscode.CancellationToken,
  ): void {
    const fileName = fileNameOf(document.uri);
    const roots = baseResourceRoots(this.extensionUri);
    if (document.uri.scheme === 'file') roots.push(vscode.Uri.joinPath(document.uri.with({ query: '', fragment: '' }), '..'));
    panel.webview.options = webviewOptions(roots);
    panel.webview.html = getHtml(panel.webview, this.extensionUri, 'pdf', fileName);

    const session: ViewerSession = new ViewerSession(panel, document.uri, 'pdf', async (msg) => {
      if (msg.type === 'ready') await sendContent();
      else await handleCommonMessage(session, msg, { allowLocal: true });
    });
    const sendContent = sendOnReadyAndReload(session, document, (outcome) => {
      if (!outcome.ok) {
        if (outcome.reason !== undefined) session.postError(`Could not open "${fileName}": ${outcome.reason}`);
        else session.postError(`Could not open "${fileName}"`, outcome.error);
        return;
      }
      const source = outcome.data ? { data: outcome.data } : { uri: panel.webview.asWebviewUri(document.uri).toString() };
      session.post({ type: 'init', kind: 'pdf', fileName, source, fileSize: outcome.fileSize });
    });
  }
}

// ===== PPTX =====

type PptxOutcome = { ok: true; deck: PptxDeck } | { ok: false; error: unknown };

/** Parses the deck once (renderPptx throws readable messages, keeping the library's error as `cause`). */
async function loadPptxFile(uri: vscode.Uri): Promise<PptxOutcome> {
  try {
    const deck = await renderPptx(await readDocumentBytes(uri, isZipPackage));
    const { slideCount, lossy } = deck.meta;
    getLog().info(
      `[pptx] Loaded ${uri.toString(true)} (${slideCount} slide${slideCount === 1 ? '' : 's'}` +
        `${lossy.length ? `; not displayed: ${lossy.join(', ')}` : ''})`,
    );
    return { ok: true, deck };
  } catch (err) {
    return { ok: false, error: err ?? new Error('Unknown error') };
  }
}

/**
 * Read-only presentation view. init carries the deck outline (slide size, titles, hidden flags); the webview asks for
 * each slide it draws with getSlide (thumbnails included), so a deck with many large pictures is not sent at once.
 */
export class PptxViewerProvider implements vscode.CustomReadonlyEditorProvider<ReadonlyFileDocument<PptxOutcome>> {
  static readonly viewType = VIEW_TYPES.pptx;

  constructor(private readonly extensionUri: vscode.Uri) {}

  openCustomDocument(uri: vscode.Uri, _openContext: vscode.CustomDocumentOpenContext): ReadonlyFileDocument<PptxOutcome> {
    return new ReadonlyFileDocument(uri, 'pptx', loadPptxFile);
  }

  resolveCustomEditor(
    document: ReadonlyFileDocument<PptxOutcome>,
    panel: vscode.WebviewPanel,
    _token: vscode.CancellationToken,
  ): void {
    const fileName = fileNameOf(document.uri);
    panel.webview.options = webviewOptions(baseResourceRoots(this.extensionUri));
    panel.webview.html = getHtml(panel.webview, this.extensionUri, 'pptx', fileName);
    /**
     * Most slides of any deck sent to this panel. A getSlide past the current deck but below this was sent for a deck
     * the panel showed before a reload removed slides; the new deck's init replaces it, so it gets no answer (an error
     * would replace the new deck). Beyond every deck the panel was shown, the index is simply invalid.
     */
    let slidesShown = 0;

    const sendSlide = async (msg: MessageOf<'getSlide'>): Promise<void> => {
      const outcome = await document.loading;
      if (!outcome.ok) {
        session.postError(`Could not open "${fileName}"`, outcome.error);
        return;
      }
      const { slideCount } = outcome.deck.meta;
      if (msg.index >= slideCount) {
        if (msg.index < slidesShown) {
          getLog().debug(`[pptx] getSlide ${msg.index + 1} of ${fileName} is for the deck before a reload; not answered`);
          return;
        }
        session.postError(
          `Could not display slide ${msg.index + 1}: "${fileName}" has ${slideCount} slide${slideCount === 1 ? '' : 's'}`,
        );
        return;
      }
      let slide: PptxSlide;
      try {
        slide = outcome.deck.getSlide(msg.index);
      } catch (err) {
        session.postError(`Could not display slide ${msg.index + 1} of "${fileName}"`, err);
        return;
      }
      session.post({ type: 'slide', reqId: msg.reqId, index: msg.index, slide });
    };

    const session: ViewerSession = new ViewerSession(panel, document.uri, 'pptx', async (msg) => {
      switch (msg.type) {
        case 'ready':
          await sendContent();
          return;
        case 'getSlide':
          await sendSlide(msg);
          return;
        default:
          await handleCommonMessage(session, msg, { allowLocal: true });
      }
    });
    const sendContent = sendOnReadyAndReload(session, document, (outcome) => {
      if (!outcome.ok) {
        session.postError(`Could not open "${fileName}"`, outcome.error);
        return;
      }
      const deck = outcome.deck.meta;
      session.post({ type: 'init', kind: 'pptx', fileName, deck, banner: lossyBanner(deck.lossy, 'presentation') });
      slidesShown = Math.max(slidesShown, deck.slideCount);
    });
  }
}

// ===== TEST API =====
// (v0.2.0, track F; contract §13.3) What integration tests use to drive FileStudio: deliver webview messages, read a
// document's state and answer modal prompts (ask). activate() returns it only when VS Code runs extension tests
// (ExtensionMode.Test); in normal use it does not exist and ask() shows real dialogs.

export interface FileStudioTestApi {
  version: 1;
  /** Deliver a WebviewMessage as if from the first panel of that document; resolves with the messages posted back to it while handling. */
  post(uri: string, msg: WebviewMessage): Promise<HostMessage[]>;
  /** Host state of an open document (`version` 0 / `dirty` false until the owning phase registers them). */
  state(uri: string): {
    version: number;
    dirty: boolean;
    editState?: EditState;
    stateId?: number;
    savedStateId?: number;
    generation?: number;
    poisoned?: boolean;
  };
  /** Queue answers for the next ask() prompts (button text, or undefined = Cancel). */
  answer(id: AskId, button: string | undefined): void;
}

/** True once activate() created the test API (ExtensionMode.Test): ask() then answers from `testAnswers`. */
let testApiEnabled = false;
/** Answers queued by FileStudioTestApi.answer, per prompt id, used in order. */
const testAnswers = new Map<AskId, (string | undefined)[]>();

/** Creates the test API (only for ExtensionMode.Test; see extension.ts). */
export function createTestApi(): FileStudioTestApi {
  testApiEnabled = true;
  const key = (uri: string): string => vscode.Uri.parse(uri, true).toString();
  return {
    version: 1,
    post(uri: string, msg: WebviewMessage): Promise<HostMessage[]> {
      const target = key(uri);
      const session = [...liveSessions].find((s) => !s.isDisposed && s.uri.toString() === target);
      if (!session) return Promise.reject(new Error(`No FileStudio panel shows ${uri}`));
      return session.deliverForTest(msg);
    },
    state(uri: string) {
      const target = key(uri);
      const doc = [...editGateDocs.values()].find((d) => d.spec.uri.toString() === target);
      if (!doc) throw new Error(`FileStudio has no open document ${uri}`);
      return {
        version: doc.spec.version?.() ?? 0,
        dirty: doc.spec.dirty?.() ?? false,
        ...(doc.state ? { editState: doc.state } : {}),
      };
    },
    answer(id: AskId, button: string | undefined): void {
      const queue = testAnswers.get(id) ?? [];
      queue.push(button);
      testAnswers.set(id, queue);
    },
  };
}

// ===== LOGGING =====

let outputChannel: vscode.LogOutputChannel | undefined;

/** The shared "FileStudio" output channel (created on first use). */
export function getLog(): vscode.LogOutputChannel {
  if (!outputChannel) outputChannel = vscode.window.createOutputChannel('FileStudio', { log: true });
  return outputChannel;
}

export function disposeLog(): void {
  outputChannel?.dispose();
  outputChannel = undefined;
}

function logError(message: string, err: unknown): void {
  getLog().error(`${message}: ${errorDetailOf(err)}`);
}
