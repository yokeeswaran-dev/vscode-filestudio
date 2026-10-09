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

import { execFile } from 'child_process';
import { randomBytes } from 'crypto';
import { constants as fsConstants } from 'fs';
import { access, open as openFileHandle } from 'fs/promises';
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
    }
  /** `readOnly`: as in the markdown init (sent again with every update, so a change of the state reaches the view). */
  | { type: 'markdownUpdate'; html: string; toc: TocEntry[]; source: string; version: number; readOnly?: boolean }
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
    };

/** Webview -> extension. */
export type WebviewMessage =
  | { type: 'ready' }
  /**
   * Rows [start, end) of a sheet. Optional `c0` / `c1`: only the cells of columns c0..c1 (inclusive), e.g. to copy
   * a few columns of a large sheet without converting every cell of its rows.
   */
  | { type: 'getRows'; reqId: number; sheet: number; start: number; end: number; c0?: number; c1?: number }
  | { type: 'getStats'; reqId: number; sheet: number; ranges: Range[] }
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
  | { type: 'log'; level: 'info' | 'warn' | 'error'; message: string };

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

/**
 * Validates a message posted by the webview. Unknown types and malformed payloads yield undefined; numbers are
 * truncated and clamped, row windows and range lists are bounded, strings are length-limited.
 */
export function parseWebviewMessage(raw: unknown): WebviewMessage | undefined {
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
      return { type: 'getStats', reqId, sheet, ranges };
    }
    case 'openLink':
      return typeof raw.href === 'string' && raw.href.length > 0 && raw.href.length <= MAX_HREF_LENGTH
        ? { type: 'openLink', href: raw.href }
        : undefined;
    case 'toggleTask': {
      const line = toInt(raw.line, 0, MAX_INDEX);
      const version = toInt(raw.version, 0, Number.MAX_SAFE_INTEGER);
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
 * One webview panel showing one document. Validates incoming messages, tracks the ready handshake (and logs when it
 * never happens), remembers what the webview currently shows and ties listeners to the panel's lifetime.
 */
class ViewerSession implements vscode.Disposable {
  readonly fileName: string;
  /** What the webview displays: nothing yet (loading), a successful init, or the error view. */
  shown: 'nothing' | 'content' | 'error' = 'nothing';

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
    this.clearReadyTimer();
    for (const disposable of this.disposables.splice(0)) {
      try {
        disposable.dispose();
      } catch (err) {
        logError(`[${this.kind}] Error while disposing the viewer of ${this.fileName}`, err);
      }
    }
  }

  private receive(raw: unknown): void {
    if (this.disposed) return;
    const msg = parseWebviewMessage(raw);
    if (!msg) {
      const type = isRecord(raw) && typeof raw.type === 'string' ? `'${raw.type.slice(0, 64)}'` : 'malformed';
      getLog().debug(`[${this.kind}] Ignored ${type} message from the viewer of ${this.fileName}`);
      return;
    }
    if (msg.type === 'ready') {
      // A (re)loaded webview starts from the static loading page again.
      this.ready = true;
      this.shown = 'nothing';
      this.waitingForVisibility = false;
      this.clearReadyTimer();
    }
    Promise.resolve()
      .then(() => this.handler(msg))
      .catch((err: unknown) => {
        if (this.disposed) return;
        if (msg.type === 'ready') this.postError(`Could not open "${this.fileName}"`, err);
        else logError(`[${this.kind}] Handling '${msg.type}' failed for ${this.fileName}`, err);
      });
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
  const cwd = path.dirname(version.workingFile.fsPath);
  const name = path.basename(version.workingFile.fsPath);
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
      session.post({ type: 'invalidate', meta: model.getMeta(), banner: lossyBanner(document.lossyFeatures) ?? null });
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
 * in the text, so an edit can switch it), CSV on the delimiter the parser detects.
 */
function parseDelimited(text: string, format: DelimitedFormat): CsvModel {
  switch (format) {
    case 'tsv':
      return parseCsv(text, { delimiter: '\t' });
    case 'psv':
      return parseCsv(text, { delimiter: '|' });
    case 'ssv':
      return detectSsvDelimiter(text) === ';'
        ? parseCsv(text, { delimiter: ';' })
        : parseCsv(text, { delimiter: ' ', collapseSpaces: true });
    default:
      return parseCsv(text);
  }
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

  constructor(
    private readonly document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
  ) {
    this.format = delimitedFormatOf(document.uri);
    this.session = new ViewerSession(panel, document.uri, 'csv', (msg) => this.onMessage(msg));
    this.reparse = debounce(() => this.refresh(), TEXT_DEBOUNCE_MS);
    this.session.onDispose(this.reparse);
    this.session.onDispose(onTextDocumentChanged(document, () => this.reparse.schedule()));
    // Parse right after the HTML is handed over, so the model is usually ready when the webview asks for it.
    const initialParse = setTimeout(() => this.ensureParsed(), 0);
    this.session.onDispose({ dispose: () => clearTimeout(initialParse) });
  }

  private ensureParsed(): void {
    if (this.session.isDisposed || this.parsedVersion === this.document.version) return;
    this.parsedVersion = this.document.version;
    try {
      const csv = parseDelimited(this.document.getText(), this.format);
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
      });
      this.initDelimiter = this.delimiter;
    } catch (err) {
      this.session.postError(`Could not display "${this.session.fileName}"`, err);
    }
  }

  /** Debounced: the text changed (typing in a text editor, undo, external reload). */
  private refresh(): void {
    this.ensureParsed();
    if (!this.session.isReady) return; // 'ready' will send init from the current text
    if (!this.model || this.session.shown !== 'content' || this.delimiter !== this.initDelimiter) {
      this.sendInit();
      return;
    }
    try {
      this.session.post({ type: 'invalidate', meta: this.model.getMeta() });
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

/**
 * Whether VS Code's text editor treats the document as read-only (filesConfigurationService.isReadonly), so a task
 * checkbox must not change it either: a read-only file system (git:, a commit), then `files.readonlyInclude`, or
 * `files.readonlyFromPermissions` with a file that cannot be written, both unless `files.readonlyExclude` matches.
 * Like VS Code, a glob matches the path relative to the workspace folder or the absolute path. "Set Active Editor
 * Read-only in Session" is not visible to extensions.
 */
async function isReadonlyDocument(document: vscode.TextDocument): Promise<boolean> {
  const uri = document.uri;
  if (vscode.workspace.fs.isWritableFileSystem(uri.scheme) === false) return true;
  const files = vscode.workspace.getConfiguration('files', uri);
  const folder = vscode.workspace.getWorkspaceFolder(uri);
  const matches = (setting: string): boolean =>
    Object.entries(files.get<Record<string, boolean>>(setting) ?? {}).some(
      ([glob, on]) =>
        on === true &&
        (vscode.languages.match({ pattern: glob }, document) > 0 ||
          (folder !== undefined && vscode.languages.match({ pattern: new vscode.RelativePattern(folder, glob) }, document) > 0)),
    );
  if (matches('readonlyInclude')) return !matches('readonlyExclude');
  // files.readonlyFromPermissions is window-scoped: read it without a resource (as VS Code does), or the extension
  // host logs "Accessing a window scoped configuration for a resource is not expected".
  if (uri.scheme === 'file' && vscode.workspace.getConfiguration('files').get<boolean>('readonlyFromPermissions') === true) {
    try {
      await access(uri.fsPath, fsConstants.W_OK);
    } catch (err) {
      // A file that is gone is not read-only: VS Code shows it as deleted, and a save writes it again.
      if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') return !matches('readonlyExclude');
    }
  }
  return false;
}

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
  /** VS Code treats the document as read-only (isReadonlyDocument): task toggles are refused, the webview disables them. */
  private readOnly = false;
  /** The first read-only check; 'ready' waits for it, so the init already says whether the checkboxes work. */
  private readonly firstReadOnlyCheck: Promise<boolean>;

  constructor(
    private readonly document: vscode.TextDocument,
    panel: vscode.WebviewPanel,
  ) {
    this.session = new ViewerSession(panel, document.uri, 'markdown', (msg) => this.onMessage(msg));
    this.rerender = debounce(() => this.pushUpdate(), TEXT_DEBOUNCE_MS);
    this.session.onDispose(this.rerender);
    this.session.onDispose(onTextDocumentChanged(document, () => this.onDocumentChanged()));
    this.firstReadOnlyCheck = this.refreshReadOnly(false);
    this.session.onDispose(
      vscode.workspace.onDidChangeConfiguration((e) => {
        if (READONLY_SETTINGS.some((setting) => e.affectsConfiguration(setting, document.uri))) void this.refreshReadOnly(true);
      }),
    );
  }

  private onMessage(msg: WebviewMessage): Promise<void> | void {
    switch (msg.type) {
      case 'ready':
        return this.firstReadOnlyCheck.then(() => this.sendInit());
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
      readOnly: this.readOnly,
    });
  }

  /**
   * Checks again whether the document is read-only. With `push`, a change goes to the webview at once (a settings
   * change, a file made read-only since the last check), so its task checkboxes follow. Never rejects.
   */
  private async refreshReadOnly(push: boolean): Promise<boolean> {
    let readOnly = false;
    try {
      readOnly = await isReadonlyDocument(this.document);
    } catch (err) {
      logError(`[markdown] Could not check whether ${this.session.fileName} is read-only`, err);
    }
    if (readOnly !== this.readOnly) {
      this.readOnly = readOnly;
      getLog().info(`[markdown] ${this.session.fileName} is ${readOnly ? 'read-only: its task checkboxes are disabled' : 'no longer read-only'}`);
      if (push) this.pushUpdate();
    }
    return readOnly;
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
      readOnly: this.readOnly,
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
  private async toggleTask(msg: MessageOf<'toggleTask'>): Promise<void> {
    const document = this.document;
    const where = `line ${msg.line + 1} of ${this.session.fileName}`;
    // Checked again at the click (the file may have been made read-only since): VS Code's own editor refuses edits to
    // a read-only document, and an edit applied anyway would change it in memory only (not dirty, never saved).
    if (await this.refreshReadOnly(true)) {
      getLog().info(`[markdown] Task toggle on ${where} refused: the document is read-only`);
      this.answerToggle(msg, false, 'the document is read-only');
      return;
    }
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
