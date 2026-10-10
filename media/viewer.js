// @ts-check
/*
 * FileStudio — webview script, bundled by esbuild into dist/viewer.js (ESM, code splitting).
 *
 * One script serves every view kind. <body data-kind> says which view is loading and the extension's
 * `init` message carries the data:
 *   sheet / csv  → virtual spreadsheet grid  (GRID, SHEET TABS, FORMULA BAR, STATUS BAR)
 *   markdown     → GitHub-like preview        (MARKDOWN)
 *   docx         → page-like document view    (DOCX)
 *   pdf          → pdf.js page viewer         (PDF)
 *   pptx         → presentation view          (PPTX)
 *
 * Plain browser JS with JSDoc types. The message protocol and the grid model are the TypeScript
 * interfaces in src/viewerProvider.ts and src/renderers/sheet.ts (type-only imports below).
 */

// @ts-ignore -- side-effect CSS import: esbuild emits it as dist/viewer.css (KaTeX styles for the Markdown view).
import 'katex/dist/katex.min.css';
// @ts-ignore -- side-effect CSS import: esbuild adds the pdf.js viewer styles (page, text and annotation layers) to dist/viewer.css.
import 'pdfjs-dist/web/pdf_viewer.css';
import { describePdfFailure, missingPdfEngineFeatures, PDF_ENGINE_UNSUPPORTED } from '../src/renderers/pdf';
import DOMPurify from 'dompurify';

/** @typedef {import('../src/renderers/sheet').Range} Range */
/** @typedef {import('../src/renderers/sheet').CellStyle} CellStyle */
/** @typedef {import('../src/renderers/sheet').FontStyle} FontStyle */
/** @typedef {import('../src/renderers/sheet').BorderEdge} BorderEdge */
/** @typedef {import('../src/renderers/sheet').RichRun} RichRun */
/** @typedef {import('../src/renderers/sheet').CellData} CellData */
/** @typedef {import('../src/renderers/sheet').FillLayout} FillLayout */
/** @typedef {import('../src/renderers/sheet').RowData} RowData */
/** @typedef {import('../src/renderers/sheet').SheetMeta} SheetMeta */
/** @typedef {import('../src/renderers/sheet').WorkbookMeta} WorkbookMeta */
/** @typedef {import('../src/renderers/sheet').SelectionStats} SelectionStats */
/** @typedef {import('../src/viewerProvider').HostMessage} HostMessage */
/** @typedef {import('../src/viewerProvider').WebviewMessage} WebviewMessage */
/** @typedef {Extract<HostMessage, { type: 'init', kind: 'sheet' | 'csv' }>} GridInitMessage */
/** @typedef {Extract<HostMessage, { type: 'rows' }>} RowsMessage */
/** @typedef {Extract<HostMessage, { type: 'stats' }>} StatsMessage */
/** @typedef {Extract<HostMessage, { type: 'invalidate' }>} InvalidateMessage */
/** @typedef {{ message: string, detail?: string, canReopenAsText?: boolean }} ErrorInfo */

// ===== BOOTSTRAP & MESSAGING =====

/**
 * @typedef {object} VsCodeApi
 * @property {(msg: WebviewMessage) => void} postMessage
 * @property {() => any} getState
 * @property {(state: any) => void} setState
 */

/** VS Code webview API (can only be acquired once per page). */
const vscode = /** @type {VsCodeApi} */ (/** @type {any} */ (globalThis).acquireVsCodeApi());
/** Root container of every view: `<div id="app">` from the extension's HTML. */
const root = /** @type {HTMLElement} */ (document.getElementById('app'));

/** @param {WebviewMessage} msg */
function post(msg) {
  vscode.postMessage(msg);
}

/** View kinds of `init` (HostMessage), i.e. the views this script can show. */
const VIEW_KINDS = ['sheet', 'csv', 'markdown', 'docx', 'pdf', 'pptx'];

/**
 * Routes one extension → webview message to the section that handles it. View-specific messages are only accepted
 * by the view they belong to (`rows` / `stats` / `invalidate`: grid, `markdownUpdate`: markdown preview); a stray one
 * (host bug, late reply) is logged and dropped instead of replacing or corrupting the view on screen.
 * @param {HostMessage} msg
 */
function dispatch(msg) {
  switch (msg.type) {
    case 'init':
      if (!VIEW_KINDS.includes(msg.kind)) {
        logError(`Ignored an 'init' message of unknown kind '${String(msg.kind)}'`);
        showError({ message: `This viewer cannot display content of kind '${String(msg.kind)}'.` });
        break;
      }
      disposeView();
      document.body.dataset.kind = msg.kind;
      if (msg.kind === 'sheet' || msg.kind === 'csv') showGrid(msg);
      else if (msg.kind === 'markdown') showMarkdown(msg);
      else if (msg.kind === 'docx') showDocx(msg);
      else if (msg.kind === 'pdf') showPdf(msg);
      else if (msg.kind === 'pptx') showPptx(msg);
      break;
    case 'rows':
      if (acceptsViewMessage(msg.type, ['sheet', 'csv'])) onRows(msg);
      break;
    case 'stats':
      if (acceptsViewMessage(msg.type, ['sheet', 'csv'])) onStats(msg);
      break;
    case 'invalidate':
      if (acceptsViewMessage(msg.type, ['sheet', 'csv'])) onInvalidate(msg);
      break;
    case 'markdownUpdate':
      if (acceptsViewMessage(msg.type, ['markdown'])) updateMarkdown(msg);
      break;
    case 'toggleTaskResult':
      if (acceptsViewMessage(msg.type, ['markdown'])) mdOnToggleResult(msg);
      break;
    case 'scrollToFragment':
      if (acceptsViewMessage(msg.type, ['markdown', 'docx'])) mdOnScrollToFragment(msg);
      break;
    case 'slide':
      if (acceptsViewMessage(msg.type, ['pptx'])) onSlide(msg);
      break;
    case 'error':
      showError(msg);
      break;
  }
}

/**
 * @param {string} type
 * @param {string[]} kinds view kinds the message applies to
 * @returns {boolean} false (and a log line) when the webview shows another kind of view
 */
function acceptsViewMessage(type, kinds) {
  const kind = currentKind();
  if (kinds.includes(kind)) return true;
  log('warn', `Ignored a '${type}' message: this webview shows a ${kind || 'not yet initialized'} view.`);
  return false;
}

window.addEventListener('message', (event) => {
  const msg = event.data;
  if (!msg || typeof msg !== 'object' || typeof msg.type !== 'string') return;
  try {
    dispatch(/** @type {HostMessage} */ (msg));
  } catch (err) {
    logError(`Failed to handle the '${msg.type}' message`, err);
    if (msg.type === 'init' || msg.type === 'invalidate' || msg.type === 'markdownUpdate') {
      showError({ message: 'This file could not be rendered.', detail: errorDetail(err), canReopenAsText: canReopenAsText() });
    }
  }
});
window.addEventListener('error', (event) => {
  // Benign browser notice (layout settled on the next frame), not a script failure.
  if (/ResizeObserver loop/.test(String(event.message))) return;
  logError('Uncaught error', event.error ?? event.message);
});
window.addEventListener('unhandledrejection', (event) => logError('Unhandled promise rejection', event.reason));

// Sent once the whole module has evaluated, so every section's top-level state is initialised.
queueMicrotask(() => post({ type: 'ready' }));

// ===== UTILITIES =====

/** Cleanup callbacks of the current view; run when it is replaced (new init, error view). */
/** @type {(() => void)[]} */
const viewDisposers = [];

/** @param {() => void} fn */
function onDispose(fn) {
  viewDisposers.push(fn);
}

function disposeView() {
  while (viewDisposers.length) {
    const fn = viewDisposers.pop();
    try {
      fn?.();
    } catch (err) {
      logError('View cleanup failed', err);
    }
  }
}

/**
 * addEventListener that is undone automatically when the view is disposed.
 * @param {EventTarget} target
 * @param {string} type
 * @param {(event: any) => void} fn
 * @param {AddEventListenerOptions | boolean} [options]
 */
function listen(target, type, fn, options) {
  target.addEventListener(type, fn, options);
  onDispose(() => target.removeEventListener(type, fn, options));
}

/**
 * Document views (Markdown, Word, PDF): when the focus falls to <body> (a round trip through the workbench, e.g. the
 * Explorer and back, or the focused element went away), the view's scroller takes it back as the frame gets the focus
 * again or a key is pressed, so arrow keys, Space and Page Up / Down go on scrolling (like the grid's restoreLostFocus).
 * A key pressed on <body> scrolls the scroller natively once it has the focus.
 * @param {HTMLElement} scroller
 */
function keepScrollerFocus(scroller) {
  const take = () => {
    const focus = document.activeElement;
    if (scroller.isConnected && (!focus || focus === document.body)) scroller.focus({ preventScroll: true });
  };
  listen(window, 'focus', take);
  listen(document, 'keydown', (/** @type {KeyboardEvent} */ e) => {
    if (e.target === document.body) take();
  }, true);
}

/** @param {unknown} err */
function errorDetail(err) {
  if (err instanceof Error) return err.stack || `${err.name}: ${err.message}`;
  return String(err);
}

/**
 * @param {'info' | 'warn' | 'error'} level
 * @param {string} message
 */
function log(level, message) {
  try {
    post({ type: 'log', level, message });
  } catch {
    // The host is gone (panel closing); nothing useful left to do.
  }
}

/**
 * @param {string} message
 * @param {unknown} [err]
 */
function logError(message, err) {
  const text = err === undefined ? message : `${message}: ${errorDetail(err)}`;
  console.error(text);
  log('error', text);
}

function currentKind() {
  return document.body.dataset.kind || '';
}

function canReopenAsText() {
  const kind = currentKind();
  return kind === 'csv' || kind === 'markdown';
}

/** @type {Record<string, string>} */
const HTML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

/** @param {unknown} s */
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (ch) => HTML_ESCAPES[ch]);
}

/**
 * Base DOMPurify configuration: data-* attributes, KaTeX MathML, style attributes, checkboxes, target.
 * Like GitHub's sanitizer, document content can neither restyle nor navigate the webview: no <style> (global CSS,
 * CSS-initiated requests), <link> / <meta> / <base>, no <form> (submitting navigates the frame) and no image maps
 * (<area href> navigates the frame past the link routing of the preview).
 */
const SANITIZE_BASE = Object.freeze({
  ADD_TAGS: ['semantics', 'annotation'],
  ADD_ATTR: ['target', 'encoding'],
  FORBID_TAGS: ['style', 'link', 'meta', 'base', 'form', 'map', 'area'],
  ALLOW_DATA_ATTR: true,
});
/** Config keys whose arrays are concatenated with the base ones instead of replacing them. */
const SANITIZE_LIST_KEYS = /** @type {const} */ (['ADD_TAGS', 'ADD_ATTR', 'FORBID_TAGS']);

/**
 * Data attributes VS Code's webview host reads from the page (`data-vscode-context`: context-menu keys and
 * `preventDefaultContextMenuItems`): document content must not set them.
 */
const RESERVED_DATA_ATTR_RE = /^data-vscode-/i;

DOMPurify.addHook('uponSanitizeAttribute', (_node, data) => {
  if (RESERVED_DATA_ATTR_RE.test(data.attrName)) data.keepAttr = false;
});

/**
 * Sanitizes untrusted HTML (rendered Markdown / DOCX). `extraConfig` is merged into the base
 * DOMPurify config; ADD_TAGS / ADD_ATTR / FORBID_TAGS arrays are concatenated instead of replaced.
 * @param {string} html
 * @param {Record<string, any>} [extraConfig]
 * @returns {string}
 */
function sanitize(html, extraConfig) {
  /** @type {Record<string, any>} */
  const config = { ...SANITIZE_BASE, ...extraConfig };
  for (const key of SANITIZE_LIST_KEYS) {
    if (extraConfig && Array.isArray(extraConfig[key])) config[key] = [...SANITIZE_BASE[key], ...extraConfig[key]];
    else config[key] = [...SANITIZE_BASE[key]];
  }
  return String(DOMPurify.sanitize(html, config));
}

/**
 * Small DOM builder: h('div', { class: 'x', onclick: fn, dataset: { a: '1' }, style: 'color:red' }, child, ...).
 * Attributes: class/className, style (string or object), dataset, text, on* (listeners), true → empty
 * attribute, null/undefined/false → skipped. Children: nodes, strings, numbers, nested arrays; nullish and
 * booleans are skipped.
 * @param {string} tag
 * @param {Record<string, any> | null} [attrs]
 * @param {...any} children
 * @returns {HTMLElement}
 */
function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [key, value] of Object.entries(attrs)) {
      if (value === null || value === undefined || value === false) continue;
      if (key === 'class' || key === 'className') el.className = String(value);
      else if (key === 'style') {
        if (typeof value === 'string') el.style.cssText = value;
        else Object.assign(el.style, value);
      } else if (key === 'dataset') Object.assign(el.dataset, value);
      else if (key === 'text') el.textContent = String(value);
      else if (key.startsWith('on') && typeof value === 'function') el.addEventListener(key.slice(2).toLowerCase(), value);
      else el.setAttribute(key, value === true ? '' : String(value));
    }
  }
  appendChildren(el, children);
  return el;
}

/**
 * @param {Node} parent
 * @param {any[]} children
 */
function appendChildren(parent, children) {
  for (const child of children) {
    if (child === null || child === undefined || typeof child === 'boolean') continue;
    if (Array.isArray(child)) appendChildren(parent, child);
    else if (child instanceof Node) parent.appendChild(child);
    else parent.appendChild(document.createTextNode(String(child)));
  }
}

/**
 * Element from a trusted, static SVG/HTML string (icons defined in this file only).
 * @param {string} markup
 * @param {string} [className]
 */
function icon(markup, className = 'fv-icon') {
  const span = document.createElement('span');
  span.className = className;
  span.setAttribute('aria-hidden', 'true');
  span.innerHTML = markup;
  return span;
}

const ICONS = {
  chevronLeft: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M10 3.5 5.5 8l4.5 4.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  chevronRight: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M6 3.5 10.5 8 6 12.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  eyeClosed: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M1.5 6.5C3 8.6 5.3 10 8 10s5-1.4 6.5-3.5M4 9l-1.3 1.8M8 10v2.2M12 9l1.3 1.8" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>',
  info: '<svg viewBox="0 0 16 16" width="16" height="16"><circle cx="8" cy="8" r="6.3" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M8 7v4.2" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><circle cx="8" cy="4.9" r=".9" fill="currentColor"/></svg>',
  error: '<svg viewBox="0 0 16 16" width="32" height="32"><circle cx="8" cy="8" r="6.3" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M5.6 5.6l4.8 4.8M10.4 5.6l-4.8 4.8" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg>',
  close: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>',
};

/**
 * Trailing-edge debounce with cancel() and flush().
 * @template {any[]} A
 * @param {(...args: A) => void} fn
 * @param {number} ms
 * @returns {((...args: A) => void) & { cancel(): void, flush(): void }}
 */
function debounce(fn, ms) {
  /** @type {ReturnType<typeof setTimeout> | undefined} */
  let timer;
  /** @type {A | null} */
  let pending = null;
  const run = () => {
    timer = undefined;
    const args = pending;
    pending = null;
    if (args) fn(...args);
  };
  /** @param {A} args */
  const debounced = (...args) => {
    pending = args;
    if (timer !== undefined) clearTimeout(timer);
    timer = setTimeout(run, ms);
  };
  debounced.cancel = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
    pending = null;
  };
  debounced.flush = () => {
    if (timer !== undefined) {
      clearTimeout(timer);
      run();
    }
  };
  return debounced;
}

/**
 * @param {number} value
 * @param {number} lo
 * @param {number} hi
 */
function clamp(value, lo, hi) {
  return value < lo ? lo : value > hi ? hi : value;
}

/** Reads one top-level key of the persisted webview state. @param {string} key */
function getStateKey(key) {
  const state = vscode.getState();
  return state && typeof state === 'object' ? state[key] : undefined;
}

/**
 * Persists one top-level key of the webview state, keeping the other keys.
 * @param {string} key
 * @param {unknown} value
 */
function setStateKey(key, value) {
  const state = vscode.getState();
  vscode.setState({ ...(state && typeof state === 'object' ? state : {}), [key]: value });
}

const IS_MAC = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);

// ===== ERROR VIEW =====

/**
 * Replaces the whole view with an error panel (never a blank page). Accepts the host's `error` message
 * or a plain string. csv/md offer "Reopen as Text", xlsx/docx "Reveal in Explorer".
 * @param {string | ErrorInfo} msg
 */
function showError(msg) {
  /** @type {ErrorInfo} */
  const info = typeof msg === 'string' ? { message: msg } : msg;
  disposeView();
  const textAction = info.canReopenAsText ?? canReopenAsText();
  const action = textAction
    ? h('button', { class: 'fv-btn', type: 'button', onclick: () => post({ type: 'reopenAsText' }) }, 'Reopen as Text')
    : h('button', { class: 'fv-btn', type: 'button', onclick: () => post({ type: 'revealInExplorer' }) }, IS_MAC ? 'Reveal in Finder' : 'Reveal in Explorer');
  const view = h(
    'div',
    { class: 'fv-error', role: 'alert' },
    icon(ICONS.error, 'fv-error-icon'),
    h('h2', { class: 'fv-error-title' }, 'Unable to display this file'),
    h('p', { class: 'fv-error-message' }, info.message || 'An unknown error occurred.'),
    info.detail ? h('details', { class: 'fv-error-detail' }, h('summary', null, 'Details'), h('pre', null, info.detail)) : null,
    h('div', { class: 'fv-error-actions' }, action),
  );
  root.className = 'fv-error-app';
  root.replaceChildren(view);
  onDispose(() => root.classList.remove('fv-error-app'));
  action.focus();
}

// ===== GRID =====
//
// Virtual spreadsheet grid (xlsx read-only in phase 1, csv/tsv).
//
// Geometry: a single native scroll container (.fv-scroller) holds a content box as large as the sheet.
// Four panes keep the headers and frozen rows/columns in place with `position: sticky`, so the browser's
// compositor moves them in lock-step with the scroll position (no scroll-linked JS jitter):
//
//   corner band (sticky top+left): select-all corner, frozen column/row headers, frozen×frozen cells
//   top band    (sticky top):      column headers + frozen-row cells of the scrolling columns
//   left band   (sticky left):     row headers + frozen-column cells of the scrolling rows
//   main layer  (scrolls):         every other cell
//
// All panes position their children in "content coordinates" — x = rowHeaderWidth + colLeft[c],
// y = colHeaderHeight + rowTop[r] — minus the pane origin, so one set of layout math serves all four.
// Only cells in the visible window (+ overscan) exist in the DOM; elements are keyed per cell and recycled
// while scrolling, and everything is (re)painted from a requestAnimationFrame callback. Sizes are CSS px
// already multiplied by the sheet zoom. Very tall sheets (big CSVs) map the native scroll range onto the
// virtual height (Chromium caps element sizes at ~33M px). Right-to-left sheets mirror the whole scroller
// and flip only glyph runs back, so every geometric rule above holds unchanged (pointer x and the
// left/right arrow keys go through scrollerX / L.rtl). No DOM measuring: text widths (overflow, ###,
// shrink-to-fit) come from a cached canvas context, and only when a cheap bound says it matters.

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const BLOCK_ROWS = 256;
const MAX_CACHED_BLOCKS = 160;
const MAX_INFLIGHT = 3;
const REQUEST_TIMEOUT_MS = 20000;
const DIRECT_CHUNK_ROWS = 4096;
/** Most rows of one column-limited copy request (fetchRows; the host serves at most 10,000 rows per getRows). */
const DIRECT_CHUNK_MAX_ROWS = 8192;
const OVERSCAN_X = 160;
const OVERSCAN_Y = 240;
const PAD_ROWS = 200;
const PAD_COLS = 30;
/** Navigating this close to the end of the layout grows it (growLayout). */
const GROW_MARGIN_ROWS = 10;
const GROW_MARGIN_COLS = 3;
const MIN_ZOOM = 10;
const MAX_ZOOM = 400;
/** Zoom change per Ctrl+wheel notch, as in Excel (the status-bar − / + buttons step by 10). */
const WHEEL_ZOOM_STEP = 15;
const MAX_SCROLL_HEIGHT = 10_000_000;
const REBASE_PX = 1_000_000;
const REGION_SNAP = 512;
const GRID_SNAP_ROWS = 64;
const GRID_SNAP_COLS = 16;
const STATS_DEBOUNCE_MS = 100;
const COPY_CELL_LIMIT = 2_000_000;
/** Cells one countCells request may cover (the host stops counting at 2,000,000). */
const PROBE_CELLS = 1_500_000;
const TOOLTIP_DELAY_MS = 350;
/** @type {CellData[]} */
const NO_CELLS = [];

/**
 * @typedef {object} Layout
 * @property {number} z          zoom factor (1 = 100%)
 * @property {number} nRows      rows in the layout (used range + padding)
 * @property {number} nCols
 * @property {Float64Array} rowTop   prefix sums of row heights, length nRows + 1 (hidden rows = 0)
 * @property {Float64Array} colLeft  prefix sums of column widths, length nCols + 1
 * @property {number} fr         frozen rows
 * @property {number} fc         frozen columns
 * @property {number} hdrW       row header width
 * @property {number} hdrH       column header height
 * @property {number} A          top band height (hdrH + frozen rows)
 * @property {number} B          left band width (hdrW + frozen cols)
 * @property {number} W          content width
 * @property {number} H          virtual content height
 * @property {number} Hc         real (capped) content height
 * @property {boolean} scaled    H > Hc: scroll positions are mapped onto the virtual height
 * @property {Int32Array} colStyle   column default style per column (0 = none)
 * @property {boolean} colStyleVisual  some column default style paints (fill/border)
 * @property {boolean} showGrid
 * @property {boolean} rtl       right-to-left sheet (the scroller is mirrored)
 * @property {number} hdrFont    header font size, px
 */

/**
 * @typedef {object} SheetView
 * @property {number} index
 * @property {SheetMeta} meta
 * @property {number} zoom            percent
 * @property {Layout | null} L
 * @property {number} maxR            Excel row limit (or rowCount for bigger CSVs)
 * @property {number} maxC
 * @property {Map<number, number[]>} mergeIndex   256-row bucket → indexes into meta.merges
 * @property {Cell} ext               farthest cell navigated to: the layout reaches it (Excel's scroll range grows the same way)
 * @property {Selection} sel
 * @property {number} scrollTop       virtual scroll position
 * @property {number} scrollLeft
 * @property {{ x: number, y: number, w: number, h: number, src: string }[]} images  content-coordinate rects
 */

/** @typedef {{ r: number, c: number }} Cell */
/**
 * `anchor` = fixed corner and `cursor` = moving corner of the last range (Shift+Arrow / Shift+click extend it). After a
 * plain move onto a merged cell, `active` is the merge's top-left cell and `cursor` the position inside the merge the
 * move arrived at, so the next arrow key keeps the row / column like Excel.
 * @typedef {{ ranges: Range[], active: Cell, anchor: Cell, cursor: Cell }} Selection
 */

/**
 * `_xk`: keys of extra elements a cell element owns in the same recycler (its overflow mask), kept alive with it.
 * @typedef {HTMLElement & { _f?: number, _dv?: number, _css?: string, _cls?: string, _html?: string, _xk?: string[] }} RcEl
 * @typedef {{ parent: HTMLElement, tag: string, map: Map<number | string, RcEl>, pool: RcEl[], frame: number }} Recycler
 */

/**
 * @typedef {object} Pane
 * @property {string} name
 * @property {HTMLElement} el    layer element; children use content coordinates minus (ox, oy)
 * @property {number} ox
 * @property {number} oy
 * @property {number} base       re-base offset for very tall sheets
 * @property {Recycler} grid     gridlines + loading placeholders
 * @property {Recycler} cells
 * @property {Recycler} imgs
 * @property {Recycler} sel
 * @property {Recycler} hdrs
 * @property {string} sigCells
 * @property {string} sigSel
 * @property {string} sigHdr
 * @property {string} sigImg
 */

/**
 * @typedef {object} StyleInfo
 * @property {string} css        fill, font, colour declarations
 * @property {string} vcls       vertical alignment class suffix
 * @property {boolean} hasFill
 * @property {string} bg         background colour of the fill ('' = none, or a gradient only)
 * @property {boolean} visual    paints even when empty (fill, border or diagonal)
 * @property {'left' | 'center' | 'right' | 'justify' | 'distributed' | 'fill' | 'centerContinuous' | undefined} h
 *   Excel's horizontal alignment; undefined = General
 * @property {boolean} wrap      wrap text (also implied by justify / distributed alignment)
 * @property {boolean} shrink
 * @property {number} indent     px
 * @property {number | 'vertical'} rot
 * @property {Edge | null} top
 * @property {Edge | null} right
 * @property {Edge | null} bottom
 * @property {Edge | null} left
 * @property {string} font       canvas font shorthand (text measuring)
 * @property {number} fontPx
 * @property {string} family     CSS font-family list
 * @property {boolean} bold
 * @property {boolean} italic
 * @property {boolean} autoColor  font colour is "Automatic" (black / unset)
 * @property {string | undefined} fontColor  the workbook's font colour (unadjusted)
 * @property {boolean} strikeIn  strikethrough drawn on an inner span (the cell's own text-decoration is a double underline)
 * @property {'superscript' | 'subscript' | undefined} va
 */
/** @typedef {{ w: number, css: string, rank: number }} Edge  rank: Excel's precedence on a shared edge (BORDER_RANK) */

/**
 * @typedef {object} GridDom
 * @property {HTMLElement} scroller
 * @property {HTMLElement} content
 * @property {HTMLElement} bandTop
 * @property {HTMLElement} bandCorner
 * @property {HTMLElement} bandLeft
 * @property {{ main: Pane, top: Pane, left: Pane, corner: Pane }} panes
 * @property {HTMLElement[]} frozenH   frozen-row separator lines (top band, corner)
 * @property {HTMLElement[]} frozenV   frozen-column separator lines (left band, corner)
 * @property {HTMLElement} cornerBtn
 * @property {HTMLInputElement} nameBox
 * @property {HTMLElement} formula
 * @property {HTMLElement} tabStrip
 * @property {HTMLElement} tabNav
 * @property {HTMLButtonElement} tabPrev
 * @property {HTMLButtonElement} tabNext
 * @property {HTMLButtonElement} hiddenBtn
 * @property {HTMLElement} statusMsg
 * @property {HTMLElement} statusStats
 * @property {HTMLElement} statusInfo
 * @property {HTMLButtonElement} zoomLabel
 * @property {HTMLElement} tooltip
 * @property {HTMLElement} ariaRow    screen-reader structure of the grid: role=row holding the active cell
 * @property {HTMLElement} ariaCell   role=gridcell, the scroller's aria-activedescendant
 * @property {HTMLElement} ariaDesc   description of the active cell (note, link target)
 * @property {HTMLElement | null} menu
 * @property {HTMLElement | null} banner
 * @property {HTMLElement} formulaBar
 * @property {HTMLElement} tabBar
 * @property {HTMLElement} statusBar
 */

/**
 * @typedef {object} GridState
 * @property {'sheet' | 'csv'} kind
 * @property {string} fileName
 * @property {boolean} readOnly
 * @property {WorkbookMeta} meta
 * @property {CellStyle[]} styles      style table (grows with `rows` replies)
 * @property {number} sheet            active sheet index (-1 while the workbook has no worksheets)
 * @property {Map<number, SheetView>} views
 * @property {Set<number>} revealed    hidden sheets shown for viewing
 * @property {Map<number, StyleInfo>} styleCache
 * @property {number} styleZoom        zoom the style cache was built for
 * @property {boolean} dark
 * @property {string | undefined} defaultFontColor
 * @property {number} dataVersion
 * @property {number} selVersion
 * @property {number} layoutVersion
 * @property {number} viewW
 * @property {number} viewH
 * @property {any} saved               persisted view state
 * @property {GridDom} dom
 * @property {boolean} wired            the grid's events are wired (first mountGrid)
 * @property {boolean} renderQueued
 * @property {number} statsSeq
 * @property {StatsRequest | null} statsReq  selection statistics being gathered (scheduleStats)
 * @property {{ mode: 'cells' | 'rows' | 'cols' | 'deselect', x: number, y: number, raf: number, anchor: number, last: number, frozenR: boolean, frozenC: boolean, deselect?: { from: Cell, base: Selection, to: Cell } } | null} drag
 *   frozenR / frozenC: the drag started in the frozen rows / columns (else it auto-scrolls instead of entering them);
 *   deselect: Ctrl+drag from a selected cell (deselectCells) - the cell, the selection before, the cell last hovered
 * @property {{ key: string, timer: ReturnType<typeof setTimeout> | undefined }} hover
 * @property {boolean} keyNav           the last selection change came from the keyboard (shows the active cell's note / link)
 * @property {{ rows: [number, number, boolean][], cols: [number, number, boolean][] }} hdrSel
 * @property {number} zoomWheelAt
 * @property {CellBox | null} cellBox   read-only cell text box opened by a double-click (CELL TEXT BOX)
 * @property {HTMLElement | null} lostFocus  the grid or the cell text box whose focus went to <body> (see wireGridEvents)
 * @property {string | undefined} bannerText  lossy-content banner the host sent last (init / invalidate), even if dismissed
 */

/** @type {GridState | null} */
let G = null;

/** Active sheet view. Only called while a grid is shown. */
function activeView() {
  const g = /** @type {GridState} */ (G);
  return /** @type {SheetView} */ (g.views.get(g.sheet));
}

/** @param {GridInitMessage} msg */
function showGrid(msg) {
  const meta = msg.meta;
  const saved = getStateKey('grid');
  const dom = buildGridDom(msg);
  G = {
    kind: msg.kind,
    fileName: msg.fileName,
    readOnly: msg.readOnly !== false,
    meta,
    styles: meta.styles.slice(),
    sheet: -1,
    views: new Map(),
    revealed: new Set(),
    styleCache: new Map(),
    styleZoom: 0,
    dark: isDarkTheme(),
    defaultFontColor: normColor(meta.styles[0]?.font?.color),
    dataVersion: 1,
    selVersion: 1,
    layoutVersion: 1,
    viewW: 0,
    viewH: 0,
    saved: saved && typeof saved === 'object' && saved.file === msg.fileName ? saved : null,
    dom,
    wired: false,
    renderQueued: false,
    statsSeq: 0,
    statsReq: null,
    drag: null,
    hover: { key: '', timer: undefined },
    keyNav: false,
    hdrSel: { rows: [], cols: [] },
    zoomWheelAt: 0,
    cellBox: null,
    lostFocus: null,
    bannerText: msg.banner,
  };
  resetRowCache();
  rowCache.metaFloor = rowCache.reqSeq;
  lastRegion = { sheet: -1, version: -1 };
  onDispose(() => {
    closeCellBox(false);
    hideTooltip();
    closeMenu();
    saveGridState.flush();
    scheduleStats.cancel();
    announceActive.cancel();
    resetRowCache();
    G = null;
    root.classList.remove('fv-grid-app');
  });

  if (!meta.sheets.length) {
    showNoSheets(msg.banner);
    return;
  }

  mountGrid(msg.banner);

  // Hidden sheets this panel had revealed for viewing stay revealed (the active one may be among them).
  if (G.saved && Array.isArray(G.saved.revealed)) {
    for (const name of G.saved.revealed) {
      const index = meta.sheets.findIndex((s) => s.name === name);
      if (index >= 0 && meta.sheets[index].state !== 'visible') G.revealed.add(index);
    }
  }
  // Initial sheet: the one persisted for this panel, else the workbook's active sheet. A hidden active sheet is shown
  // revealed (Excel opens the file on it, unhiding it); any other hidden one gives way to the first visible sheet.
  let first = -1;
  if (G.saved && typeof G.saved.activeSheet === 'string') first = meta.sheets.findIndex((s) => s.name === G?.saved.activeSheet);
  if (first < 0) {
    first = clamp(meta.activeSheet | 0, 0, meta.sheets.length - 1);
    if (meta.sheets[first].state === 'hidden') G.revealed.add(first);
  }
  if (meta.sheets[first].state !== 'visible' && !G.revealed.has(first)) {
    const visible = meta.sheets.findIndex((s) => s.state === 'visible');
    if (visible >= 0) first = visible;
    else G.revealed.add(first);
  }
  activateSheet(first, true);
  dom.scroller.focus({ preventScroll: true });
}

/**
 * A workbook without worksheets: ExcelJS reads only worksheets, so these are files whose sheets are all chart
 * sheets (or dialog / macro sheets). The lossy-content banner still says what is not shown.
 * @param {string | undefined} banner
 */
function showNoSheets(banner) {
  root.className = 'fv-grid-app';
  root.replaceChildren(
    ...(banner ? [buildBanner(banner, null)] : []),
    h('div', { class: 'fv-empty', role: 'status' }, 'This workbook has no worksheets to show. Its sheets are chart sheets (or other sheet types) that FileStudio does not display.'),
  );
}

/**
 * Puts the grid (banner, formula bar, scroller, tabs, status bar) into #app; the first time also wires its events.
 * Also brings the grid back when worksheets return after showNoSheets (invalidate).
 * @param {string | undefined} banner
 */
function mountGrid(banner) {
  const g = /** @type {GridState} */ (G);
  root.className = 'fv-grid-app';
  root.replaceChildren(...buildGridApp({ banner }, g.dom));
  g.viewW = g.dom.scroller.clientWidth;
  g.viewH = g.dom.scroller.clientHeight;
  if (!g.wired) {
    g.wired = true;
    wireGridEvents();
  }
}

/**
 * Builds the grid's DOM skeleton (no data yet).
 * @param {GridInitMessage} msg
 * @returns {GridDom}
 */
function buildGridDom(msg) {
  const panes = { main: makePane('main'), top: makePane('top'), left: makePane('left'), corner: makePane('corner') };
  const cornerBtn = h('div', { class: 'fv-corner', title: 'Select all', role: 'button', 'aria-label': 'Select all cells' });
  panes.corner.el.append(cornerBtn);
  const frozenH = [h('div', { class: 'fv-frz fv-frz-h' }), h('div', { class: 'fv-frz fv-frz-h' })];
  const frozenV = [h('div', { class: 'fv-frz fv-frz-v' }), h('div', { class: 'fv-frz fv-frz-v' })];
  const bandCorner = h('div', { class: 'fv-band fv-band-corner' }, panes.corner.el, frozenH[1], frozenV[1]);
  const bandTop = h('div', { class: 'fv-band fv-band-top' }, panes.top.el, frozenH[0], bandCorner);
  const bandLeft = h('div', { class: 'fv-band fv-band-left' }, panes.left.el, frozenV[0]);
  // The painted panes are presentation only. Screen readers get a grid whose single row holds the active cell
  // (aria-activedescendant, updated by announceActive): a virtual grid cannot expose every cell.
  const content = h('div', { class: 'fv-content', 'aria-hidden': 'true' }, panes.main.el, bandTop, bandLeft);
  const ariaCell = h('div', { role: 'gridcell', id: 'fv-grid-cell-0', 'aria-selected': 'true', 'aria-describedby': 'fv-grid-desc' });
  const ariaRow = h('div', { class: 'fv-sr-only', role: 'row' }, ariaCell);
  const scroller = h(
    'div',
    {
      class: 'fv-scroller',
      tabindex: '0',
      role: 'grid',
      'aria-readonly': 'true',
      'aria-multiselectable': 'true',
      'aria-label': `${msg.fileName} spreadsheet`,
      'aria-activedescendant': ariaCell.id,
    },
    ariaRow,
    content,
  );
  const fb = buildFormulaBar();
  const tabs = buildTabBar();
  const status = buildStatusBar(msg);
  return {
    scroller,
    content,
    bandTop,
    bandCorner,
    bandLeft,
    panes,
    frozenH,
    frozenV,
    cornerBtn,
    nameBox: fb.nameBox,
    formula: fb.formula,
    tabStrip: tabs.strip,
    tabNav: tabs.nav,
    tabPrev: tabs.prev,
    tabNext: tabs.next,
    hiddenBtn: tabs.hiddenBtn,
    statusMsg: status.msg,
    statusStats: status.stats,
    statusInfo: status.info,
    zoomLabel: status.zoomLabel,
    tooltip: h('div', { class: 'fv-tooltip', role: 'tooltip', hidden: true }),
    ariaRow,
    ariaCell,
    ariaDesc: h('div', { class: 'fv-sr-only', id: 'fv-grid-desc' }),
    menu: null,
    banner: null,
    formulaBar: fb.bar,
    tabBar: tabs.bar,
    statusBar: status.bar,
  };
}

/**
 * Top-level children of #app for the grid view.
 * @param {{ banner?: string }} msg  lossy-content banner text (GridInitMessage.banner)
 * @param {GridDom} dom
 */
function buildGridApp(msg, dom) {
  /** @type {HTMLElement[]} */
  const parts = [];
  if (msg.banner) parts.push(buildGridBanner(msg.banner, dom));
  parts.push(dom.formulaBar, h('div', { class: 'fv-grid-host' }, dom.scroller), dom.tabBar, dom.statusBar, dom.tooltip, dom.ariaDesc);
  return parts;
}

/**
 * The grid's lossy-content banner (kept in dom.banner).
 * @param {string} text @param {GridDom} dom
 */
function buildGridBanner(text, dom) {
  dom.banner = buildBanner(text, () => {
    dom.banner = null;
    // The Dismiss button had focus: give it back to the grid instead of losing it to <body>.
    G?.dom.scroller.focus({ preventScroll: true });
  });
  return dom.banner;
}

/**
 * After a reload (InvalidateMessage.banner): shows, replaces or removes the lossy-content banner, on the grid or on the
 * "no worksheets" page, so it says what the reloaded file holds (a chart added or removed on disk). A banner the user
 * dismissed stays dismissed while the host sends the same text.
 * @param {string | undefined} text
 */
function updateBanner(text) {
  const g = G;
  if (!g || text === g.bannerText) return;
  g.bannerText = text;
  const focused = document.activeElement;
  root.querySelector(':scope > .fv-banner')?.remove();
  g.dom.banner = null;
  if (text) root.prepend(g.dom.scroller.isConnected ? buildGridBanner(text, g.dom) : buildBanner(text, null));
  // Removing a banner whose Dismiss button had the focus would drop it to <body>.
  if (focused instanceof HTMLElement && !focused.isConnected && g.dom.scroller.isConnected) g.dom.scroller.focus({ preventScroll: true });
}

/**
 * Lossy-content banner with a Dismiss button.
 * @param {string} text @param {(() => void) | null} onDismiss called after the banner is removed
 */
function buildBanner(text, onDismiss) {
  const banner = h(
    'div',
    { class: 'fv-banner', role: 'status' },
    icon(ICONS.info),
    h('span', { class: 'fv-banner-text', title: text }, text),
    h(
      'button',
      {
        class: 'fv-icon-btn fv-banner-close',
        type: 'button',
        title: 'Dismiss',
        'aria-label': 'Dismiss',
        onclick: () => {
          banner.remove();
          onDismiss?.();
        },
      },
      icon(ICONS.close),
    ),
  );
  return banner;
}

function wireGridEvents() {
  const g = /** @type {GridState} */ (G);
  const sc = g.dom.scroller;
  listen(sc, 'scroll', onGridScroll, { passive: true });
  listen(sc, 'mousedown', onGridMouseDown);
  listen(sc, 'mousemove', onGridHover);
  listen(sc, 'mouseleave', () => hoverReset());
  listen(sc, 'keydown', onGridKeyDown);
  listen(sc, 'wheel', onGridWheel, { passive: false });
  listen(sc, 'dblclick', onGridDblClick);
  listen(window, 'mousemove', onWindowMouseMove);
  listen(window, 'mouseup', onWindowMouseUp);
  listen(window, 'keydown', (e) => trackModifier(e, true));
  listen(window, 'keyup', (e) => trackModifier(e, false));
  listen(window, 'blur', () => {
    trackModifier(null, false);
    onWindowMouseUp();
  });
  // The webview document going away (reload, editor moved to another window) keeps the last change: the state save
  // is debounced and would otherwise be dropped.
  listen(window, 'pagehide', () => saveGridState.flush());
  listen(document, 'copy', onCopyEvent);
  // VS Code's context menu (also a click in the workbench) takes the focus out of the webview's frame, which drops it
  // to <body>: the grid / cell text box that had it is remembered, so that the menu's Copy copies from it and the
  // focus goes back to it when the frame gets the focus again or a key is pressed (onLostFocusKeyDown).
  listen(document, 'focusout', (/** @type {FocusEvent} */ e) => {
    if (G && !e.relatedTarget && (e.target === G.dom.scroller || e.target === G.cellBox?.el)) G.lostFocus = /** @type {HTMLElement} */ (e.target);
  });
  listen(document, 'focusin', () => {
    if (G) G.lostFocus = null;
  });
  listen(window, 'focus', () => restoreLostFocus());
  listen(document, 'keydown', onLostFocusKeyDown);
  listen(document, 'mousedown', (e) => {
    // A click anywhere but the cell text box closes it (the focus goes back to the grid unless the click focuses
    // something else, such as the name box).
    if (G?.cellBox && !G.cellBox.el.contains(/** @type {Node} */ (e.target))) closeCellBox();
    const menu = G?.dom.menu;
    if (menu && !menu.contains(/** @type {Node} */ (e.target)) && e.target !== G?.dom.hiddenBtn && !G?.dom.hiddenBtn.contains(/** @type {Node} */ (e.target))) closeMenu();
  });
  listen(root, 'keydown', onGridAppKeyDown);
  // Clicking the bars (their buttons, labels, empty space) keeps the keyboard focus where it is, so arrow keys go on
  // driving the grid like in Excel. The name box and the formula text still take focus (typing / selecting text).
  // Delegated from #app: mountGrid builds a new banner when the grid comes back after showNoSheets.
  listen(root, 'mousedown', (/** @type {MouseEvent} */ e) => {
    const target = /** @type {HTMLElement} */ (e.target);
    if (e.button === 0 && target.closest('.fv-banner, .fv-formula-bar, .fv-tabbar, .fv-statusbar') && !target.closest('.fv-namebox, .fv-formula')) e.preventDefault();
    // A press elsewhere ends a text selection in the formula bar, like leaving Excel's formula bar (the grid's press
    // is default-prevented, which would keep it highlighted).
    const s = window.getSelection();
    if (s && !s.isCollapsed && !target.closest('.fv-formula') && G?.dom.formula.contains(s.anchorNode)) s.removeAllRanges();
  });

  // Tab arrows are updated on the next frame: showing them resizes the observed tab strip, which inside the
  // callback would trigger the browser's "ResizeObserver loop" error.
  let arrowsRaf = 0;
  const ro = new ResizeObserver(() => {
    if (!G) return;
    // The cell text box is placed for the old size.
    if (G.cellBox && (sc.clientWidth !== G.viewW || sc.clientHeight !== G.viewH)) closeCellBox();
    G.viewW = sc.clientWidth;
    G.viewH = sc.clientHeight;
    scheduleRender();
    if (!arrowsRaf) {
      arrowsRaf = requestAnimationFrame(() => {
        arrowsRaf = 0;
        updateTabArrows();
      });
    }
  });
  ro.observe(sc);
  ro.observe(g.dom.tabStrip);
  onDispose(() => {
    ro.disconnect();
    if (arrowsRaf) cancelAnimationFrame(arrowsRaf);
  });

  // Theme switches change auto colours baked into diagonal-border images and the explicit colours adjusted for the
  // editor background (readableOnDark): rebuild the style table. VS Code switches a theme by replacing the
  // variables on <html> and the body classes, so both are observed and the cache key includes the background —
  // two dark (or two light) themes share the body class.
  let themeKey = gridThemeKey();
  const mo = new MutationObserver(() => {
    if (!G) return;
    const key = gridThemeKey();
    if (key === themeKey) return;
    themeKey = key;
    closeCellBox(); // its colours come from the old theme
    G.dark = isDarkTheme();
    darkColorCache.clear();
    darkBgLuminance = -1;
    G.styleCache.clear();
    G.layoutVersion++;
    for (const p of Object.values(G.dom.panes)) clearPane(p);
    scheduleRender();
  });
  mo.observe(document.body, { attributes: true, attributeFilter: ['class'] });
  mo.observe(document.documentElement, { attributes: true, attributeFilter: ['style'] });
  onDispose(() => mo.disconnect());
}

function isDarkTheme() {
  const cl = document.body.classList;
  return cl.contains('vscode-dark') || (cl.contains('vscode-high-contrast') && !cl.contains('vscode-high-contrast-light'));
}

/** What the grid's theme-dependent caches depend on: dark or light, and the editor background. */
function gridThemeKey() {
  return `${isDarkTheme()}|${getComputedStyle(document.documentElement).getPropertyValue('--vscode-editor-background').trim()}`;
}

/**
 * Switches the visible sheet (also used for the first display and after invalidate).
 * @param {number} index
 * @param {boolean} [force]
 */
function activateSheet(index, force = false) {
  if (!G) return;
  if (index < 0 || index >= G.meta.sheets.length) return;
  if (index === G.sheet && !force) return;
  closeCellBox();
  hideTooltip();
  closeMenu();
  if (G.meta.sheets[index].state !== 'visible') G.revealed.add(index);
  G.sheet = index;
  const view = getView(index);
  applyLayout(view);
  setVirtualScroll(view, view.scrollTop, view.scrollLeft);
  G.dom.scroller.setAttribute('aria-rowcount', String(view.maxR));
  G.dom.scroller.setAttribute('aria-colcount', String(view.maxC));
  G.selVersion++;
  computeHeaderSel(view);
  renderTabs();
  updateFormulaBar();
  updateStatusInfo();
  updateZoomLabel();
  scheduleStats();
  saveGridState();
  announceActive();
  scheduleRender();
}

/**
 * View state for a sheet, created on first use (restores persisted scroll/selection/zoom).
 * @param {number} index
 * @returns {SheetView}
 */
function getView(index) {
  const g = /** @type {GridState} */ (G);
  const existing = g.views.get(index);
  if (existing) return existing;
  const sm = g.meta.sheets[index];
  const saved = g.saved?.sheets?.[sm.name];
  /** @type {SheetView} */
  const view = {
    index,
    meta: sm,
    zoom: clamp(Math.round(Number(saved?.z) || sm.zoom || 100), MIN_ZOOM, MAX_ZOOM),
    L: null,
    maxR: Math.max(MAX_ROWS, sm.rowCount),
    maxC: Math.max(MAX_COLS, sm.colCount),
    mergeIndex: buildMergeIndex(sm.merges),
    ext: { r: 0, c: 0 },
    sel: /** @type {Selection} */ (/** @type {unknown} */ (null)),
    scrollTop: Math.max(0, Number(saved?.st) || 0),
    scrollLeft: Math.max(0, Number(saved?.sl) || 0),
    images: [],
  };
  view.sel = restoreSelection(view, saved?.sel) ?? initialSelection(view);
  // A restored selection beyond the used range keeps its cells inside the layout.
  for (const p of [view.sel.active, view.sel.anchor, view.sel.cursor]) {
    view.ext = { r: Math.max(view.ext.r, p.r), c: Math.max(view.ext.c, p.c) };
  }
  g.views.set(index, view);
  return view;
}

// ----- MODEL CACHE -----
// Rows live in the extension host. The grid asks for 256-row blocks (getRows) and keeps a bounded LRU of
// them. A block is requested at most once at a time (in-flight map), at most MAX_INFLIGHT requests are
// outstanding (visible blocks first, so fast scrolling only fetches what is still on screen), and replies
// whose reqId no longer matches the in-flight entry (cache invalidated, request timed out) are dropped.
// Copy of large selections uses "direct" requests that bypass the cache. Keyboard searches over rows that are not
// loaded (Ctrl+Arrow, Ctrl+A) ask the host to count non-empty cells instead (getStats "probes", countCells).

const rowCache = {
  /** @type {Map<number, (RowData | undefined)[]>} */
  blocks: new Map(),
  /** @type {Map<number, { reqId: number, sent: number }>} */
  inflight: new Map(),
  /** Blocks wanted by the current frame, most important first. @type {number[]} */
  wanted: [],
  /** @type {Map<number, { resolve: (msg: RowsMessage) => void, reject: (err: Error) => void }>} */
  direct: new Map(),
  /** getStats requests of countCells by reqId (ids from G.statsSeq, shared with the status bar). */
  /** @type {Map<number, { resolve: (count: number) => void, reject: (err: Error) => void }>} */
  probes: new Map(),
  reqSeq: 0,
  /** Replies to requests up to this id predate the current WorkbookMeta (style indexes may differ). */
  metaFloor: 0,
};

/**
 * @param {number} sheet
 * @param {number} block
 */
function blockKey(sheet, block) {
  return sheet * 1_048_576 + block;
}

function resetRowCache() {
  rowCache.blocks.clear();
  rowCache.inflight.clear();
  rowCache.wanted = [];
  for (const pending of rowCache.direct.values()) pending.reject(new Error('The data changed while it was being read.'));
  rowCache.direct.clear();
  for (const pending of rowCache.probes.values()) pending.reject(new Error('The data changed while it was being read.'));
  rowCache.probes.clear();
}

/**
 * Row data: RowData, null = loaded and empty, undefined = not loaded yet.
 * @param {number} sheet
 * @param {number} r
 * @returns {RowData | null | undefined}
 */
function getRow(sheet, r) {
  const g = /** @type {GridState} */ (G);
  const sm = g.meta.sheets[sheet];
  if (!sm || r < 0 || r >= sm.rowCount) return null;
  const block = rowCache.blocks.get(blockKey(sheet, Math.floor(r / BLOCK_ROWS)));
  if (!block) return undefined;
  return block[r % BLOCK_ROWS] || null;
}

/**
 * Cell at (r, c) of the active sheet; undefined when its row is not loaded yet.
 * @param {SheetView} view
 * @param {number} r
 * @param {number} c
 * @returns {CellData | null | undefined}
 */
function cellAt(view, r, c) {
  const row = getRow(view.index, r);
  if (row === undefined) return undefined;
  return row ? findCell(row.cells, c) : null;
}

/** Sends getRows for wanted blocks that are neither cached nor in flight. */
function pumpRequests() {
  if (!G) return;
  const now = performance.now();
  for (const [key, inf] of rowCache.inflight) if (now - inf.sent > REQUEST_TIMEOUT_MS) rowCache.inflight.delete(key);
  for (const key of rowCache.wanted) {
    if (rowCache.inflight.size >= MAX_INFLIGHT) break;
    if (rowCache.blocks.has(key) || rowCache.inflight.has(key)) continue;
    const sheet = Math.floor(key / 1_048_576);
    const block = key % 1_048_576;
    const sm = G.meta.sheets[sheet];
    if (!sm) continue;
    const start = block * BLOCK_ROWS;
    const end = Math.min(start + BLOCK_ROWS, sm.rowCount);
    if (start >= end) continue;
    const reqId = ++rowCache.reqSeq;
    rowCache.inflight.set(key, { reqId, sent: now });
    post({ type: 'getRows', reqId, sheet, start, end });
  }
}

/**
 * Records the blocks the current frame needs (priority order) and refreshes their LRU position.
 * @param {number[]} keys
 */
function setWantedBlocks(keys) {
  rowCache.wanted = keys;
  for (const key of keys) {
    const block = rowCache.blocks.get(key);
    if (block) {
      rowCache.blocks.delete(key);
      rowCache.blocks.set(key, block);
    }
  }
}

function evictBlocks() {
  if (rowCache.blocks.size <= MAX_CACHED_BLOCKS) return;
  const keep = new Set(rowCache.wanted);
  for (const key of rowCache.blocks.keys()) {
    if (rowCache.blocks.size <= MAX_CACHED_BLOCKS) break;
    if (!keep.has(key)) rowCache.blocks.delete(key);
  }
}

/**
 * Adds style-table entries discovered by the host. Indexes are stable for one WorkbookMeta, so entries that
 * are already known are kept (their cached rendering info stays valid).
 * @param {[number, CellStyle][]} entries
 */
function mergeStyles(entries) {
  if (!G || !Array.isArray(entries)) return;
  for (const entry of entries) {
    if (!Array.isArray(entry)) continue;
    const [index, style] = entry;
    if (!Number.isInteger(index) || index <= 0 || G.styles[index] !== undefined) continue;
    G.styles[index] = style || {};
    G.styleCache.delete(index);
  }
}

/** @param {RowsMessage} msg */
function onRows(msg) {
  if (!G) return;
  // Style entries are valid for the current meta even when the rows themselves arrive too late.
  if (msg.reqId > rowCache.metaFloor) mergeStyles(msg.styles);
  const direct = rowCache.direct.get(msg.reqId);
  if (direct) {
    rowCache.direct.delete(msg.reqId);
    direct.resolve(msg);
    return;
  }
  if (msg.c0 !== undefined) return; // column-limited rows (copy) never go into the render cache
  const key = blockKey(msg.sheet, Math.floor(msg.start / BLOCK_ROWS));
  const inflight = rowCache.inflight.get(key);
  if (!inflight || inflight.reqId !== msg.reqId) return; // stale reply
  rowCache.inflight.delete(key);
  /** @type {(RowData | undefined)[]} */
  const block = new Array(BLOCK_ROWS);
  for (const row of msg.rows) {
    const i = row.r - msg.start;
    if (i >= 0 && i < BLOCK_ROWS) block[i] = row;
  }
  rowCache.blocks.set(key, block);
  evictBlocks();
  G.dataVersion++;
  if (msg.sheet === G.sheet) {
    scheduleRender();
    const active = activeView().sel.active;
    if (active.r >= msg.start && active.r < msg.end) {
      updateFormulaBar();
      // Announced before its row was loaded: say it again with the text.
      if (G.dom.ariaCell.dataset.loaded === 'false') announceActive();
    }
    // A cell text box opened before its row was loaded gets its text now.
    const box = G.cellBox;
    if (box && !box.loaded && box.sheet === msg.sheet && box.r >= msg.start && box.r < msg.end) openCellBox(activeView(), box.r, box.c);
  }
  pumpRequests();
}

/**
 * Reads rows [r0, r1] of a sheet for copy: cached blocks are used directly, the rest is fetched with
 * uncached getRows requests. With `cols` the fetched rows hold only the cells of columns cols.c0..cols.c1
 * (getRows c0 / c1): copying one column of a wide sheet does not convert (or transfer) all the others, and the
 * requests can then cover more rows each. Such rows are partial: they never go into the render cache.
 * @param {number} sheet
 * @param {number} r0
 * @param {number} r1
 * @param {{ c0: number, c1: number }} [cols]
 * @returns {Promise<Map<number, RowData>>}
 */
async function fetchRows(sheet, r0, r1, cols) {
  const g = /** @type {GridState} */ (G);
  const sm = g.meta.sheets[sheet];
  /** @type {Map<number, RowData>} */
  const rows = new Map();
  r1 = Math.min(r1, sm.rowCount - 1);
  // About DIRECT_CHUNK_ROWS full rows of cells per request (at most DIRECT_CHUNK_MAX_ROWS rows; the host caps 10,000).
  const chunkRows = cols
    ? clamp(Math.floor((DIRECT_CHUNK_ROWS * Math.max(sm.colCount, 1)) / (cols.c1 - cols.c0 + 1)), DIRECT_CHUNK_ROWS, DIRECT_CHUNK_MAX_ROWS)
    : DIRECT_CHUNK_ROWS;
  /** @type {[number, number][]} */
  const missing = [];
  for (let b = Math.floor(r0 / BLOCK_ROWS); b <= Math.floor(r1 / BLOCK_ROWS); b++) {
    const block = rowCache.blocks.get(blockKey(sheet, b));
    const start = Math.max(r0, b * BLOCK_ROWS);
    const end = Math.min(r1, b * BLOCK_ROWS + BLOCK_ROWS - 1);
    if (block) {
      for (let r = start; r <= end; r++) {
        const row = block[r % BLOCK_ROWS];
        if (row) rows.set(r, row);
      }
    } else if (missing.length && missing[missing.length - 1][1] === start - 1 && end + 1 - missing[missing.length - 1][0] <= chunkRows) {
      missing[missing.length - 1][1] = end;
    } else missing.push([start, end]);
  }
  for (const [start, end] of missing) {
    const reqId = ++rowCache.reqSeq;
    /** @type {RowsMessage} */
    const reply = await new Promise((resolve, reject) => {
      rowCache.direct.set(reqId, { resolve, reject });
      post(cols ? { type: 'getRows', reqId, sheet, start, end: end + 1, c0: cols.c0, c1: cols.c1 } : { type: 'getRows', reqId, sheet, start, end: end + 1 });
    });
    for (const row of reply.rows) if (row.r >= r0 && row.r <= r1) rows.set(row.r, row);
  }
  return rows;
}

/**
 * Non-empty cells in the ranges of a sheet, counted by the host (getStats: same notion of "empty" as the status bar
 * Count). Lets Ctrl+Arrow / Ctrl+A search rows that are not loaded without fetching them.
 * @param {number} sheet @param {Range[]} ranges
 * @returns {Promise<number>}
 */
function countCells(sheet, ranges) {
  const g = /** @type {GridState} */ (G);
  const reqId = ++g.statsSeq;
  return new Promise((resolve, reject) => {
    rowCache.probes.set(reqId, { resolve, reject });
    post({ type: 'getStats', reqId, sheet, ranges });
  });
}

/** @param {InvalidateMessage} msg */
function onInvalidate(msg) {
  if (!G) return;
  closeCellBox(); // its text may have changed
  resetRowCache();
  hideTooltip();
  if (msg.banner !== undefined) updateBanner(msg.banner ?? undefined);
  if (msg.meta) {
    rowCache.metaFloor = rowCache.reqSeq;
    G.saved = snapshotState();
    const activeName = G.saved.activeSheet;
    const revealedNames = new Set(Array.isArray(G.saved.revealed) ? G.saved.revealed : []);
    G.meta = msg.meta;
    G.styles = msg.meta.styles.slice();
    G.styleCache.clear();
    G.defaultFontColor = normColor(msg.meta.styles[0]?.font?.color);
    G.views = new Map();
    G.revealed = new Set();
    msg.meta.sheets.forEach((s, i) => {
      if (revealedNames.has(s.name)) G?.revealed.add(i);
    });
    if (!msg.meta.sheets.length) {
      // No active sheet until worksheets come back: queued renders, stats and announcements stop.
      const banner = G.dom.banner?.isConnected ? G.dom.banner.querySelector('.fv-banner-text')?.textContent ?? undefined : undefined;
      G.sheet = -1;
      closeMenu();
      onWindowMouseUp();
      showNoSheets(banner);
      return;
    }
    // Worksheets again after the "no worksheets" page (or a file that opened without any): the grid comes back.
    if (!G.dom.scroller.isConnected) {
      mountGrid(root.querySelector('.fv-banner-text')?.textContent ?? undefined);
      if (document.activeElement === document.body) G.dom.scroller.focus({ preventScroll: true });
    }
    let index = msg.meta.sheets.findIndex((s) => s.name === activeName);
    if (index < 0) {
      // The active sheet is gone: same position, else the workbook's active sheet (revealed when it is hidden, as on
      // first display); another hidden one only if revealed.
      index = clamp(G.sheet >= 0 ? G.sheet : msg.meta.activeSheet | 0, 0, msg.meta.sheets.length - 1);
      if (G.sheet < 0 && msg.meta.sheets[index].state === 'hidden') G.revealed.add(index);
      const visible = msg.meta.sheets.findIndex((s) => s.state === 'visible');
      if (msg.meta.sheets[index].state !== 'visible' && !G.revealed.has(index) && visible >= 0) index = visible;
    }
    G.sheet = -1;
    activateSheet(index, true);
  } else {
    G.dataVersion++;
    scheduleRender();
    updateFormulaBar();
    scheduleStats();
  }
}

// ----- LAYOUT -----

/**
 * Builds the prefix-sum layout of a sheet at its zoom.
 * @param {SheetView} view
 * @returns {Layout}
 */
function buildLayout(view) {
  const g = /** @type {GridState} */ (G);
  const sm = view.meta;
  const z = view.zoom / 100;
  let usedR = Math.max(1, sm.rowCount);
  let usedC = Math.max(1, sm.colCount);
  for (const m of sm.merges) {
    usedR = Math.max(usedR, m.r1 + 1);
    usedC = Math.max(usedC, m.c1 + 1);
  }
  for (const im of sm.images) {
    const to = im.to ?? im.from;
    usedR = Math.max(usedR, to.r + 1);
    usedC = Math.max(usedC, to.c + 1);
  }
  usedR = Math.max(usedR, sm.frozen.rows + 1, view.ext.r + 1);
  usedC = Math.max(usedC, sm.frozen.cols + 1, view.ext.c + 1);
  const nRows = Math.min(view.maxR, usedR + Math.ceil(PAD_ROWS / z));
  const nCols = Math.min(view.maxC, Math.max(usedC + Math.ceil(PAD_COLS / z), Math.ceil(26 / z)));

  const defH = Math.max(0, Math.round((sm.defaultRowHeight || 20) * z));
  const rowTop = new Float64Array(nRows + 1);
  rowTop.fill(defH, 1);
  for (const key in sm.rows) {
    const r = Number(key);
    if (!(r >= 0 && r < nRows)) continue;
    const info = sm.rows[r];
    rowTop[r + 1] = info.hidden ? 0 : info.h !== undefined && info.h !== null ? Math.max(0, Math.round(info.h * z)) : defH;
  }
  for (let r = 1; r <= nRows; r++) rowTop[r] += rowTop[r - 1];

  const defW = Math.max(0, Math.round((sm.defaultColWidth || 64) * z));
  const colLeft = new Float64Array(nCols + 1);
  colLeft.fill(defW, 1);
  const colStyle = new Int32Array(nCols);
  let colStyleVisual = false;
  for (const key in sm.cols) {
    const c = Number(key);
    if (!(c >= 0 && c < nCols)) continue;
    const info = sm.cols[c];
    colLeft[c + 1] = info.hidden ? 0 : info.w !== undefined && info.w !== null ? Math.max(0, Math.round(info.w * z)) : defW;
    if (info.s) {
      colStyle[c] = info.s;
      if (styleInfoRaw(g.styles[info.s], z).visual) colStyleVisual = true;
    }
  }
  for (let c = 1; c <= nCols; c++) colLeft[c] += colLeft[c - 1];

  const fr = clamp(sm.frozen.rows | 0, 0, nRows - 1);
  const fc = clamp(sm.frozen.cols | 0, 0, nCols - 1);
  const digits = Math.max(3, String(nRows).length);
  const hdrH = Math.max(12, Math.round(20 * z));
  const hdrW = Math.max(20, Math.round((digits * 7 + 12) * z));
  const H = hdrH + rowTop[nRows];
  const Hc = Math.min(H, MAX_SCROLL_HEIGHT);
  return {
    z,
    nRows,
    nCols,
    rowTop,
    colLeft,
    fr,
    fc,
    hdrW,
    hdrH,
    A: hdrH + rowTop[fr],
    B: hdrW + colLeft[fc],
    W: hdrW + colLeft[nCols],
    H,
    Hc,
    scaled: H > Hc,
    colStyle,
    colStyleVisual,
    showGrid: sm.showGridLines !== false,
    rtl: !!sm.rightToLeft,
    hdrFont: Math.max(6, Math.round(11 * z * 10) / 10),
  };
}

/**
 * Largest index i in [0, n-1] with prefix[i] <= pos (binary search). Hidden (zero-size) entries share
 * the prefix value of the next visible one, so the visible entry is returned.
 * @param {Float64Array} prefix
 * @param {number} n
 * @param {number} pos
 */
function findIndex(prefix, n, pos) {
  if (pos <= 0) {
    let i = 0;
    while (i < n - 1 && prefix[i + 1] === 0) i++;
    return pos < 0 ? 0 : i;
  }
  if (pos >= prefix[n]) return n - 1;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (prefix[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** @param {Layout} L @param {number} r */
function rowHeight(L, r) {
  return L.rowTop[r + 1] - L.rowTop[r];
}

/** @param {Layout} L @param {number} c */
function colWidth(L, c) {
  return L.colLeft[c + 1] - L.colLeft[c];
}

/**
 * Next visible row from r in direction dir (±1); r itself when there is none.
 * @param {Layout} L @param {number} r @param {number} dir
 */
function stepRow(L, r, dir) {
  let x = r + dir;
  while (x >= 0 && x < L.nRows && rowHeight(L, x) === 0) x += dir;
  return x < 0 || x >= L.nRows ? r : x;
}

/** @param {Layout} L @param {number} c @param {number} dir */
function stepCol(L, c, dir) {
  let x = c + dir;
  while (x >= 0 && x < L.nCols && colWidth(L, x) === 0) x += dir;
  return x < 0 || x >= L.nCols ? c : x;
}

/** First visible row at or after r (or before, when none after). @param {Layout} L @param {number} r */
function visibleRow(L, r) {
  r = clamp(r, 0, L.nRows - 1);
  if (rowHeight(L, r) > 0) return r;
  const next = stepRow(L, r, 1);
  return next !== r ? next : stepRow(L, r, -1);
}

/** @param {Layout} L @param {number} c */
function visibleCol(L, c) {
  c = clamp(c, 0, L.nCols - 1);
  if (colWidth(L, c) > 0) return c;
  const next = stepCol(L, c, 1);
  return next !== c ? next : stepCol(L, c, -1);
}

/**
 * Buckets merges by 256-row blocks for O(1)-ish lookup.
 * @param {Range[]} merges
 */
function buildMergeIndex(merges) {
  /** @type {Map<number, number[]>} */
  const index = new Map();
  merges.forEach((m, i) => {
    for (let b = Math.floor(m.r0 / BLOCK_ROWS); b <= Math.floor(m.r1 / BLOCK_ROWS); b++) {
      const list = index.get(b);
      if (list) list.push(i);
      else index.set(b, [i]);
    }
  });
  return index;
}

/**
 * Index of the merge covering (r, c), or -1.
 * @param {SheetView} view @param {number} r @param {number} c
 */
function mergeAt(view, r, c) {
  if (!view.mergeIndex.size) return -1;
  const list = view.mergeIndex.get(Math.floor(r / BLOCK_ROWS));
  if (!list) return -1;
  const merges = view.meta.merges;
  for (const i of list) {
    const m = merges[i];
    if (r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) return i;
  }
  return -1;
}

/**
 * Indexes of merges that may intersect rows [r0, r1] (deduplicated).
 * @param {SheetView} view @param {number} r0 @param {number} r1
 */
function mergesInRows(view, r0, r1) {
  const merges = view.meta.merges;
  if (!merges.length) return [];
  const b0 = Math.floor(r0 / BLOCK_ROWS);
  const b1 = Math.floor(r1 / BLOCK_ROWS);
  if (b1 - b0 > merges.length) return merges.map((_, i) => i);
  /** @type {Set<number>} */
  const out = new Set();
  for (let b = b0; b <= b1; b++) {
    const list = view.mergeIndex.get(b);
    if (list) for (const i of list) out.add(i);
  }
  return [...out];
}

/** Top-left cell of the merge containing (r, c), or the cell itself. @param {SheetView} view @param {number} r @param {number} c */
function masterOf(view, r, c) {
  const mi = mergeAt(view, r, c);
  if (mi < 0) return { r, c };
  const m = view.meta.merges[mi];
  return { r: m.r0, c: m.c0 };
}

/** Area of the cell (its merge when merged). @param {SheetView} view @param {number} r @param {number} c @returns {Range} */
function cellArea(view, r, c) {
  const mi = mergeAt(view, r, c);
  if (mi < 0) return { r0: r, c0: c, r1: r, c1: c };
  const m = view.meta.merges[mi];
  return { r0: m.r0, c0: m.c0, r1: m.r1, c1: m.c1 };
}

/** Vertical scroll factor for very tall sheets (1 otherwise). @param {Layout} L */
function yScale(L) {
  if (!L.scaled || !G) return 1;
  const vh = G.viewH;
  return Math.max(1, (L.H - vh) / Math.max(1, L.Hc - vh));
}

/** Virtual (layout) scroll position of the scroller. */
function virtualScrollTop() {
  const g = /** @type {GridState} */ (G);
  const L = /** @type {Layout} */ (activeView().L);
  const st = g.dom.scroller.scrollTop;
  return L.scaled ? st * yScale(L) : st;
}

/**
 * Scrolls to a virtual position (undefined keeps an axis).
 * @param {SheetView} view
 * @param {number | undefined} top
 * @param {number | undefined} left
 */
function setVirtualScroll(view, top, left) {
  const g = /** @type {GridState} */ (G);
  const L = /** @type {Layout} */ (view.L);
  const sc = g.dom.scroller;
  if (top !== undefined) {
    const real = L.scaled ? top / yScale(L) : top;
    if (Math.abs(sc.scrollTop - real) >= 0.5) sc.scrollTop = real;
  }
  if (left !== undefined && Math.abs(sc.scrollLeft - left) >= 0.5) sc.scrollLeft = left;
}

/**
 * Extends the layout so that it reaches (r, c) with room to spare: navigating past the used range (arrow keys,
 * Ctrl+Arrow to the sheet end, name box, links) grows the scroll range like Excel's. The scroll position is kept.
 * @param {SheetView} view @param {number} r @param {number} c
 */
function growLayout(view, r, c) {
  const L = view.L;
  if (!G || !L) return;
  const growR = r >= L.nRows - GROW_MARGIN_ROWS && L.nRows < view.maxR;
  const growC = c >= L.nCols - GROW_MARGIN_COLS && L.nCols < view.maxC;
  if (!growR && !growC) return;
  view.ext = { r: Math.max(view.ext.r, Math.min(r, view.maxR - 1)), c: Math.max(view.ext.c, Math.min(c, view.maxC - 1)) };
  if (view.index !== G.sheet) {
    view.L = null; // rebuilt when the sheet is shown
    return;
  }
  const top = virtualScrollTop();
  const left = G.dom.scroller.scrollLeft;
  view.L = null;
  applyLayout(view);
  setVirtualScroll(view, top, left);
  G.selVersion++;
  scheduleRender();
}

/**
 * (Re)applies the active view's layout to the DOM: content size, bands, fonts, frozen lines; clears panes.
 * @param {SheetView} view
 */
function applyLayout(view) {
  const g = /** @type {GridState} */ (G);
  if (g.styleZoom !== view.zoom) {
    g.styleCache.clear();
    g.styleZoom = view.zoom;
  }
  if (!view.L || view.L.z !== view.zoom / 100) view.L = buildLayout(view);
  const L = view.L;
  g.layoutVersion++;
  const d = g.dom;
  const def = g.meta.defaultFont || { name: 'Calibri', size: 11 };
  d.content.style.cssText = `width:${L.W}px;height:${L.Hc}px;font-family:${fontFamilyCss(def.name || 'Calibri')};font-size:${ptToPx(def.size || 11, L.z)}px;`;
  d.scroller.style.setProperty('--fv-hdr-font', `${L.hdrFont}px`);
  d.scroller.classList.toggle('fv-nogrid', !L.showGrid);
  // Right-to-left sheets: the scroller (panes, borders, overflow, selection, scrollbar) is mirrored as a
  // whole and only glyph runs, images and rotated text are flipped back (see .fv-rtl in viewer.css).
  d.scroller.classList.toggle('fv-rtl', L.rtl);
  d.bandTop.style.cssText = `width:${L.W}px;height:${L.A}px;`;
  d.bandCorner.style.cssText = `width:${L.B}px;height:${L.A}px;`;
  d.bandLeft.style.cssText = `width:${L.B}px;height:${Math.max(0, L.Hc - L.A)}px;`;
  d.cornerBtn.style.cssText = `width:${L.hdrW}px;height:${L.hdrH}px;`;
  const P = d.panes;
  P.left.oy = L.A;
  P.main.oy = 0;
  P.main.base = 0;
  P.left.base = 0;
  P.main.el.style.top = '0px';
  P.left.el.style.top = '0px';
  for (const p of Object.values(P)) clearPane(p);
  // Frozen separators: [0] lives in the scrolling band, [1] in the corner band.
  d.frozenH[0].style.cssText = L.fr ? `top:${L.A - 1}px;left:0;width:${L.W}px;` : 'display:none';
  d.frozenH[1].style.cssText = L.fr ? `top:${L.A - 1}px;left:0;width:${L.B}px;` : 'display:none';
  d.frozenV[0].style.cssText = L.fc ? `left:${L.B - 1}px;top:0;height:${Math.max(0, L.Hc - L.A)}px;` : 'display:none';
  d.frozenV[1].style.cssText = L.fc ? `left:${L.B - 1}px;top:0;height:${L.A}px;` : 'display:none';
  view.images = computeImageRects(view, L);
}

/**
 * Content-coordinate rectangles of the sheet's images.
 * @param {SheetView} view @param {Layout} L
 */
function computeImageRects(view, L) {
  const out = [];
  for (const im of view.meta.images) {
    if (!im || !im.src || !im.from) continue;
    const fc = clamp(im.from.c, 0, L.nCols);
    const fr = clamp(im.from.r, 0, L.nRows);
    const x = L.hdrW + L.colLeft[fc] + (im.from.dx || 0) * L.z;
    const y = L.hdrH + L.rowTop[fr] + (im.from.dy || 0) * L.z;
    let w = 0;
    let hh = 0;
    if (im.to) {
      w = L.hdrW + L.colLeft[clamp(im.to.c, 0, L.nCols)] + (im.to.dx || 0) * L.z - x;
      hh = L.hdrH + L.rowTop[clamp(im.to.r, 0, L.nRows)] + (im.to.dy || 0) * L.z - y;
    }
    if ((w <= 0 || hh <= 0) && im.ext) {
      w = im.ext.w * L.z;
      hh = im.ext.h * L.z;
    }
    if (w > 0 && hh > 0) out.push({ x, y, w, h: hh, src: im.src });
  }
  return out;
}

// ----- RENDER -----

/** @param {HTMLElement} parent @param {string} [tag] */
function makeRecycler(parent, tag = 'div') {
  return /** @type {Recycler} */ ({ parent, tag, map: new Map(), pool: [], frame: 0 });
}

/** @param {Recycler} rc */
function rcBegin(rc) {
  rc.frame++;
}

/** Detached elements kept per recycler for reuse (measured: cheaper than keeping them attached but hidden). */
const RC_POOL_MAX = 400;

/**
 * Element for a key in this frame (reused when it already exists, else taken from the pool).
 * @param {Recycler} rc
 * @param {number | string} key
 * @returns {RcEl}
 */
function rcGet(rc, key) {
  let el = rc.map.get(key);
  if (el) {
    el._f = rc.frame;
    return el;
  }
  el = rc.pool.pop() || /** @type {RcEl} */ (document.createElement(rc.tag));
  el._f = rc.frame;
  el._dv = -1;
  el._css = undefined;
  el._cls = undefined;
  el._html = undefined;
  rc.map.set(key, el);
  rc.parent.appendChild(el);
  return el;
}

/** Detaches elements not used in this frame and pools them. @param {Recycler} rc */
function rcEnd(rc) {
  for (const [key, el] of rc.map) {
    if (el._f === rc.frame) continue;
    rc.map.delete(key);
    el.remove();
    if (rc.pool.length < RC_POOL_MAX) rc.pool.push(el);
  }
}

/** @param {Recycler} rc */
function rcClear(rc) {
  rc.frame++;
  rcEnd(rc);
}

/** @param {RcEl} el @param {string} css */
function setCss(el, css) {
  if (el._css !== css) {
    el._css = css;
    el.style.cssText = css;
  }
}

/** @param {RcEl} el @param {string} cls */
function setCls(el, cls) {
  if (el._cls !== cls) {
    el._cls = cls;
    el.className = cls;
  }
}

/** @param {RcEl} el @param {string} text */
function setText(el, text) {
  if (el._html !== text) {
    el._html = text;
    el.textContent = text;
  }
}

/** Trusted markup built by this file (all cell text is escaped). @param {RcEl} el @param {string} html */
function setHtml(el, html) {
  const key = '\u0001' + html;
  if (el._html !== key) {
    el._html = key;
    el.innerHTML = html;
  }
}

/** @param {string} name @returns {Pane} */
function makePane(name) {
  const el = h('div', { class: `fv-layer fv-layer-${name}` });
  const sub = (/** @type {string} */ cls) => {
    const s = h('div', { class: `fv-sub ${cls}` });
    el.appendChild(s);
    return s;
  };
  const gridEl = sub('fv-sub-grid');
  const cellsEl = sub('fv-sub-cells');
  const imgsEl = sub('fv-sub-imgs');
  const selEl = sub('fv-sub-sel');
  const hdrsEl = sub('fv-sub-hdrs');
  return {
    name,
    el,
    ox: 0,
    oy: 0,
    base: 0,
    grid: makeRecycler(gridEl),
    cells: makeRecycler(cellsEl),
    imgs: makeRecycler(imgsEl, 'img'),
    sel: makeRecycler(selEl),
    hdrs: makeRecycler(hdrsEl),
    sigCells: '',
    sigSel: '',
    sigHdr: '',
    sigImg: '',
  };
}

/** @param {Pane} p */
function clearPane(p) {
  rcClear(p.grid);
  rcClear(p.cells);
  rcClear(p.imgs);
  rcClear(p.sel);
  rcClear(p.hdrs);
  p.sigCells = p.sigSel = p.sigHdr = p.sigImg = '';
}

function scheduleRender() {
  if (!G || G.renderQueued) return;
  G.renderQueued = true;
  requestAnimationFrame(renderFrame);
}

function renderFrame() {
  if (!G) return;
  G.renderQueued = false;
  if (G.sheet < 0) return; // no worksheets (showNoSheets)
  try {
    render();
  } catch (err) {
    logError('Grid render failed', err);
  }
}

function onGridScroll() {
  if (!G) return;
  // Scrolling closes the cell text box (not the scroll event of the position it opened at).
  const box = G.cellBox;
  if (box && (Math.abs(G.dom.scroller.scrollTop - box.st) >= 1 || Math.abs(G.dom.scroller.scrollLeft - box.sl) >= 1)) closeCellBox();
  scheduleRender();
  hoverReset();
  saveGridState();
}

/** Renders the visible window of the active sheet into the four panes. */
function render() {
  const g = /** @type {GridState} */ (G);
  const view = activeView();
  const L = view.L;
  if (!L) return;
  const sc = g.dom.scroller;
  const st = sc.scrollTop;
  const sl = sc.scrollLeft;
  const vy = L.scaled ? st * yScale(L) : st;
  view.scrollTop = vy;
  view.scrollLeft = sl;
  const vw = g.viewW;
  const vh = g.viewH;
  const P = g.dom.panes;

  // Very tall sheets: main/left layers are shifted by the difference between virtual and real scroll.
  if (L.scaled) {
    const dyn = vy - st;
    const base = Math.floor(dyn / REBASE_PX) * REBASE_PX;
    for (const p of [P.main, P.left]) {
      if (p.base !== base) {
        p.base = base;
        p.oy = (p === P.left ? L.A : 0) + base;
        clearPane(p);
      }
      p.el.style.top = `${-(dyn - base)}px`;
    }
  }

  // Visible windows (layout indexes).
  const fr = L.fr;
  const fc = L.fc;
  const frEnd = fr > 0 ? Math.min(fr - 1, findIndex(L.rowTop, L.nRows, Math.max(0, vh - L.hdrH))) : -1;
  const fcEnd = fc > 0 ? Math.min(fc - 1, findIndex(L.colLeft, L.nCols, Math.max(0, vw - L.hdrW))) : -1;
  const y0 = L.rowTop[fr] + vy;
  const x0 = L.colLeft[fc] + sl;
  const r0 = Math.max(fr, findIndex(L.rowTop, L.nRows, y0 - OVERSCAN_Y));
  const r1 = Math.max(r0, findIndex(L.rowTop, L.nRows, y0 + Math.max(0, vh - L.A) + OVERSCAN_Y));
  const c0 = Math.max(fc, findIndex(L.colLeft, L.nCols, x0 - OVERSCAN_X));
  const c1 = Math.max(c0, findIndex(L.colLeft, L.nCols, x0 + Math.max(0, vw - L.B) + OVERSCAN_X));

  // Data: request the blocks of everything on screen first, then merge masters, the active row, prefetch.
  /** @type {number[]} */
  const keys = [];
  const rowCount = view.meta.rowCount;
  const want = (/** @type {number} */ r) => {
    if (r < 0 || r >= rowCount) return;
    const key = blockKey(view.index, Math.floor(r / BLOCK_ROWS));
    if (!keys.includes(key)) keys.push(key);
  };
  const visTop = findIndex(L.rowTop, L.nRows, y0);
  const visBot = findIndex(L.rowTop, L.nRows, y0 + Math.max(0, vh - L.A));
  for (let r = visTop; r <= visBot; r += BLOCK_ROWS) want(r);
  want(visBot);
  for (let r = 0; r <= frEnd; r += BLOCK_ROWS) want(r);
  if (frEnd >= 0) want(frEnd);
  want(r0);
  want(r1);
  for (const mi of mergesInRows(view, r0, r1)) want(view.meta.merges[mi].r0);
  for (const mi of frEnd >= 0 ? mergesInRows(view, 0, frEnd) : []) want(view.meta.merges[mi].r0);
  want(view.sel.active.r);
  want(r1 + BLOCK_ROWS);
  want(r0 - BLOCK_ROWS);
  setWantedBlocks(keys);
  pumpRequests();

  // Clip regions (content coordinates) of each pane's visible area, for selection and images. They are
  // snapped outward to REGION_SNAP steps so that the overlays are not rebuilt on every scrolled pixel; what
  // lies beyond the viewport is off-screen or under the opaque header/frozen bands.
  const lo = (/** @type {number} */ v) => Math.floor(v / REGION_SNAP) * REGION_SNAP;
  const hi = (/** @type {number} */ v) => Math.ceil(v / REGION_SNAP) * REGION_SNAP;
  const regions = {
    corner: { x0: 0, y0: 0, x1: L.B, y1: L.A },
    top: { x0: lo(sl + L.B), y0: 0, x1: hi(sl + vw), y1: L.A },
    left: { x0: 0, y0: lo(vy + L.A), x1: L.B, y1: hi(vy + vh) },
    main: { x0: lo(sl + L.B), y0: lo(vy + L.A), x1: hi(sl + vw), y1: hi(vy + vh) },
  };

  renderPane(P.corner, view, 0, frEnd, 0, fcEnd, regions.corner);
  renderPane(P.top, view, 0, frEnd, c0, c1, regions.top);
  renderPane(P.left, view, r0, r1, 0, fcEnd, regions.left);
  renderPane(P.main, view, r0, r1, c0, c1, regions.main);
  renderHeaders(P.corner, L, fcEnd >= 0 ? [0, fcEnd] : null, frEnd >= 0 ? [0, frEnd] : null);
  renderHeaders(P.top, L, [c0, c1], null);
  renderHeaders(P.left, L, null, [r0, r1]);
}

/**
 * @param {Pane} p
 * @param {SheetView} view
 * @param {number} ra @param {number} rb @param {number} ca @param {number} cb
 * @param {{ x0: number, y0: number, x1: number, y1: number }} region
 */
function renderPane(p, view, ra, rb, ca, cb, region) {
  const g = /** @type {GridState} */ (G);
  const L = /** @type {Layout} */ (view.L);
  const win = `${ra},${rb},${ca},${cb},${g.layoutVersion}`;
  const cellSig = `${win},${g.dataVersion}`;
  if (p.sigCells !== cellSig) {
    p.sigCells = cellSig;
    renderPaneCells(p, view, L, ra, rb, ca, cb);
    renderPaneGrid(p, view, L, ra, rb, ca, cb);
  }
  const regionKey = `${Math.round(region.x0)},${Math.round(region.y0)},${Math.round(region.x1)},${Math.round(region.y1)}`;
  const selSig = `${win},${g.selVersion},${regionKey}`;
  if (p.sigSel !== selSig) {
    p.sigSel = selSig;
    renderPaneSelection(p, view, L, region);
  }
  if (view.images.length) {
    const imgSig = `${g.layoutVersion},${regionKey}`;
    if (p.sigImg !== imgSig) {
      p.sigImg = imgSig;
      renderPaneImages(p, view, L, region);
    }
  }
}

/**
 * Cells of rows [ra, rb] × cols [ca, cb] of one pane, plus overflow sources just outside the window and
 * merged areas intersecting it.
 * @param {Pane} p @param {SheetView} view @param {Layout} L
 * @param {number} ra @param {number} rb @param {number} ca @param {number} cb
 */
function renderPaneCells(p, view, L, ra, rb, ca, cb) {
  const rc = p.cells;
  rcBegin(rc);
  if (ra <= rb && ca <= cb) {
    const sm = view.meta;
    // Columns this pane owns (overflow sources outside the window must still belong to the pane).
    const ownC0 = p.name === 'main' || p.name === 'top' ? L.fc : 0;
    const ownC1 = p.name === 'main' || p.name === 'top' ? L.nCols - 1 : L.fc - 1;
    for (let r = ra; r <= rb; r++) {
      if (rowHeight(L, r) === 0) continue;
      const row = getRow(view.index, r);
      if (row === undefined) continue; // placeholder drawn by renderPaneGrid
      const cells = row ? row.cells : NO_CELLS;
      const rowS = rowStyleIndex(sm, r);
      let i = lowerBound(cells, ca);
      // Text overflowing into the window from the left (a formula returning "" blocks overflow, see overflowTarget).
      for (let j = i - 1, n = 0; j >= 0 && n < 64; j--, n++) {
        if (isEmptyCell(cells[j]) && !cells[j].f) continue;
        if (cells[j].c >= ownC0) renderCell(p, view, L, r, cells[j], cells, rowS);
        break;
      }
      const fillGaps = (rowS !== 0 && styleInfo(rowS).visual) || L.colStyleVisual;
      if (fillGaps) {
        for (let c = ca; c <= cb; c++) {
          if (i < cells.length && cells[i].c === c) renderCell(p, view, L, r, cells[i++], cells, rowS);
          else renderStyleOnly(p, view, L, r, c, cells, rowS);
        }
      } else {
        for (; i < cells.length && cells[i].c <= cb; i++) renderCell(p, view, L, r, cells[i], cells, rowS);
      }
      // Right-aligned / centred text overflowing into the window from the right.
      for (let j = lowerBound(cells, cb + 1), n = 0; j < cells.length && n < 64; j++, n++) {
        if (isEmptyCell(cells[j]) && !cells[j].f) continue;
        if (cells[j].c <= ownC1) renderCell(p, view, L, r, cells[j], cells, rowS);
        break;
      }
    }
    for (const mi of mergesInRows(view, ra, rb)) {
      const m = sm.merges[mi];
      if (m.c1 < ca || m.c0 > cb || m.r1 < ra || m.r0 > rb) continue;
      renderMerge(p, view, L, mi, m);
    }
  }
  rcEnd(rc);
}

/**
 * @param {Pane} p @param {SheetView} view @param {Layout} L @param {number} r
 * @param {CellData} cell @param {CellData[]} cells @param {number} rowS
 */
function renderCell(p, view, L, r, cell, cells, rowS) {
  const c = cell.c;
  if (c >= L.nCols || colWidth(L, c) === 0) return;
  if (view.mergeIndex.size && mergeAt(view, r, c) >= 0) return;
  const s = cell.s || 0;
  if (isEmptyCell(cell) && !cell.note && !styleInfo(s).visual) {
    // An explicitly unstyled cell inside a painted row/column band hides the band.
    if (!(rowS && styleInfo(rowS).visual) && !(L.colStyle[c] && styleInfo(L.colStyle[c]).visual)) return;
  }
  paintCell(p, view, L, r * MAX_COLS + c, null, r, c, cell, s, cells);
}

/**
 * Empty cell painted by its row/column default style.
 * @param {Pane} p @param {SheetView} view @param {Layout} L @param {number} r @param {number} c
 * @param {CellData[]} cells @param {number} rowS
 */
function renderStyleOnly(p, view, L, r, c, cells, rowS) {
  if (colWidth(L, c) === 0) return;
  const s = rowS || L.colStyle[c] || 0;
  if (!s || !styleInfo(s).visual) return;
  if (view.mergeIndex.size && mergeAt(view, r, c) >= 0) return;
  paintCell(p, view, L, r * MAX_COLS + c, null, r, c, null, s, cells);
}

/**
 * @param {Pane} p @param {SheetView} view @param {Layout} L @param {number} mi @param {Range} m
 */
function renderMerge(p, view, L, mi, m) {
  if (L.colLeft[Math.min(m.c1 + 1, L.nCols)] - L.colLeft[m.c0] <= 0) return;
  if (L.rowTop[Math.min(m.r1 + 1, L.nRows)] - L.rowTop[m.r0] <= 0) return;
  const row = getRow(view.index, m.r0);
  if (row === undefined) return;
  const cells = row ? row.cells : NO_CELLS;
  const cell = findCell(cells, m.c0);
  const s = cell ? cell.s || 0 : rowStyleIndex(view.meta, m.r0) || L.colStyle[m.c0] || 0;
  paintCell(p, view, L, 'm' + mi, m, m.r0, m.c0, cell, s, cells);
}

/**
 * Paints one cell or merged area: borders (cellBorders / mergeBorders), fill and font (style table), and the text —
 * fitted (numbers), repeated (fill alignment), shrunk, rotated, or spilling into empty neighbours (overflow).
 * @param {Pane} p @param {SheetView} view @param {Layout} L
 * @param {number | string} key
 * @param {Range | null} m
 * @param {number} r @param {number} c
 * @param {CellData | null} cell
 * @param {number} s
 * @param {CellData[]} rowCells
 */
function paintCell(p, view, L, key, m, r, c, cell, s, rowCells) {
  const g = /** @type {GridState} */ (G);
  const el = rcGet(p.cells, key);
  if (el._dv === g.dataVersion) {
    if (el._xk) for (const k of el._xk) rcGet(p.cells, k);
    return;
  }
  el._dv = g.dataVersion;

  const S = styleInfo(s);
  const rEnd = m ? Math.min(m.r1, L.nRows - 1) : r;
  const cEnd = m ? Math.min(m.c1, L.nCols - 1) : c;
  const x = L.hdrW + L.colLeft[c] - p.ox;
  const y = L.hdrH + L.rowTop[r] - p.oy;
  const w = L.colLeft[cEnd + 1] - L.colLeft[c];
  const hh = L.rowTop[rEnd + 1] - L.rowTop[r];
  const B = m ? mergeBorders(view, L, m, rEnd, cEnd, S, rowCells) : cellBorders(view, L, r, c, S, rowCells);
  const { left, top, extL, extT } = B;

  // Text. Numbers, dates, booleans and errors are values: never wrapped and never cut — they lose decimals
  // (General) or show ### like Excel.
  const value = !!cell && (cell.t === 'n' || cell.t === 'd' || cell.t === 'b' || cell.t === 'e');
  const wrap = S.wrap && !value;
  const raw = cell ? displayText(cell) : '';
  let text = !wrap && raw.includes('\n') ? oneLine(raw) : raw;
  /** @type {RichRun[] | null} */
  let rich = cell && cell.rt && cell.rt.length ? cell.rt : null;
  // Fitted values always fit their cell (no clipping check needed below).
  const fitted = !!text && value && !S.rot && !S.shrink && !rich;
  if (fitted) text = fitNumber(cell, text, S, w);
  // Repeat-to-fill formats (accounting `_("$"* #,##0.00_)`, `@*-`) span the whole cell whatever its alignment (fillHtml);
  // a value that did not fit (### / rounded) is shown as such. Not with wrapping, rotation, shrinking or rich text.
  const fillFmt = cell && cell.fill && text === raw && !wrap && !S.rot && !S.shrink && !rich && S.h !== 'fill' ? cell.fill : null;
  // Right-to-left sheets mirror the scroller as a whole, but Excel lays a cell's text out as on a left-to-right sheet
  // (alignment, indent side, overflow direction): the layout below uses the mirrored alignment. ±90° and stacked text
  // are the exception — Excel mirrors their position with the sheet.
  const hText = cellAlign(S, cell);
  const hEff = L.rtl && S.rot !== 'vertical' && S.rot !== 90 && S.rot !== -90 ? mirrorAlign(hText) : hText;
  const indentL = hEff === 'left' || hEff === 'distributed' ? S.indent : 0;
  const indentR = hEff === 'right' || hEff === 'distributed' ? S.indent : 0;
  // Fill alignment repeats the content across the cell (as many whole copies as fit) and never overflows.
  if (S.h === 'fill' && text && !wrap && !S.rot) {
    const one = rich ? richTextWidth(rich, S) : measureText(S.font, text);
    const copies = one > 0 ? Math.floor((w - 4 - S.indent) / one) : 0;
    if (copies > 1) {
      text = text.repeat(copies);
      if (rich) rich = /** @type {RichRun[]} */ (Array(copies).fill(rich).flat());
    }
  }
  /** @type {Overflow | null} */
  let ov = null;
  if (!m && cell && text && !wrap && !S.shrink && !S.rot && S.h !== 'fill' && !fillFmt) {
    if (S.h === 'centerContinuous') ov = centerAcross(view, L, r, c, w, S, rowCells);
    if (!ov && cell.t === 's') ov = computeOverflow(view, L, r, c, w, text, S, hEff, rowCells, rich);
  }
  let shrink = 1;
  if (S.shrink && !wrap && text) {
    const tw = (rich ? richTextWidth(rich, S) : measureText(S.font, text)) + 4 + S.indent;
    if (tw > w && w > 0) shrink = Math.max(0.2, w / tw);
  }

  // Unpainted merged areas cover inner gridlines but leave their outer right/bottom gridline visible.
  const trimW = m && !S.hasFill && L.showGrid && !B.right && !B.rMixed ? 1 : 0;
  const trimH = m && !S.hasFill && L.showGrid && !B.bottom && !B.bMixed ? 1 : 0;
  let css = `left:${x - extL}px;top:${y - extT}px;width:${w + extL - trimW}px;height:${hh + extT - trimH}px;${S.css}`;
  if (m && !S.hasFill) css += 'background-color:var(--fv-cell-bg);';
  /** @type {string | undefined} */
  let fmtColor;
  if (cell && cell.t !== 'z' && !rich) {
    fmtColor = numFmtColor(g.styles[s]?.numFmt, cell);
    if (fmtColor) css += `color:${cellFontColor(fmtColor, S.hasFill) || 'var(--fv-cell-fg)'};`;
  }
  if (left) css += `border-left:${left.css};`;
  if (top) css += `border-top:${top.css};`;
  if (B.right) css += `border-right:${B.right.css};`;
  if (B.bottom) css += `border-bottom:${B.bottom.css};`;
  const pl = Math.max(0, 2 + extL - (left ? left.w : 0)) + indentL;
  const pt = Math.max(0, extT - (top ? top.w : 0));
  const pr = 2 + indentR;
  if (pl !== 2 || pr !== 2 || pt !== 0) css += `padding:${pt}px ${pr}px 0 ${pl}px;`;
  if (!ov && hEff === 'distributed') {
    // Every line is spread over the width (the last one too); a single word is centred.
    css += /\s/.test(text.trim()) ? 'text-align:justify;text-align-last:justify;' : 'text-align:center;';
  } else if (!ov && hEff !== 'left') css += `text-align:${hEff};`;
  if (shrink < 1) css += `font-size:${Math.max(1, Math.round(S.fontPx * shrink * 100) / 100)}px;`;

  let cls = 'fv-c' + S.vcls;
  if (wrap) cls += ' fv-wrap';
  if (ov) cls += ' fv-ovc';
  if (m) cls += ' fv-m';
  if (cell && cell.note) cls += ' fv-note';
  if (cell && cell.link) cls += ' fv-link';

  /** @type {string | null} */
  let html = null;
  if (text) {
    // Right-to-left: glyph runs sit in shrink-wrapped spans that flip back inside the mirrored scroller. Wrapped lines
    // inside a span are aligned by the unmirrored alignment (the span flips back as a whole).
    const spanCss = hEff !== hText ? ` style="text-align:${hText}"` : '';
    const glyphs = L.rtl ? (/** @type {string} */ s) => `<span class="fv-g"${spanCss}>${s}</span>` : (/** @type {string} */ s) => s;
    /** Text markup with the cell's (or, given hasFill, a background's) colours. @param {boolean} hasFill */
    const body = (hasFill) => {
      let t = rich ? richHtml(rich, hasFill, !wrap, shrink) : escapeHtml(text);
      if (S.va && !rich) t = `<span class="fv-va-${S.va === 'superscript' ? 'sup' : 'sub'}">${t}</span>`;
      return S.strikeIn ? `<span class="fv-strike">${t}</span>` : t;
    };
    if (fillFmt) {
      html = fillHtml(text, fillFmt, S, w - 4 - S.indent);
    } else if (ov) {
      html = overflowHtml(ov, w, hh, hEff, S, overflowLayers(ov, w, S, fmtColor, body, glyphs), extL - (left ? left.w : 0), extT - (top ? top.w : 0));
    } else if (S.rot) {
      const tw = S.rot === 'vertical' || Math.abs(S.rot) === 90 ? 0 : rich ? richTextWidth(rich, S) : measureText(S.font, text);
      html = rotatedHtml(S.rot, hEff, body(S.hasFill), L.rtl, tw, S.fontPx * 1.2);
    } else {
      // Text wider than a right-aligned / centred cell (blocked, merged, no room) keeps its end / middle visible:
      // the shrink-wrapped block overflows its flex container at the start side (text-align cannot do that).
      const clip = !fitted && !wrap && shrink === 1 && (hEff === 'right' || hEff === 'center') && !textFits(text, rich, S, w - 4 - S.indent);
      const inner = body(S.hasFill);
      if (clip) html = `<div class="fv-tx fv-clip" style="align-self:${hEff === 'right' ? 'flex-end' : 'center'}">${glyphs(inner)}</div>`;
      else if (rich || S.va || S.strikeIn || L.rtl) html = `<div class="fv-tx">${glyphs(inner)}</div>`;
    }
  }
  if (B.segs) html = (html ?? escapeHtml(text)) + B.segs;
  setCss(el, css);
  setCls(el, cls);
  if (html !== null) setHtml(el, html);
  else setText(el, text);

  // Gridlines are hidden under overflowing text (Excel): a cell-background mask BELOW every cell (z-index -1 in the
  // cells layer), so neighbours' fills and borders stay visible over it.
  if (ov && L.showGrid) {
    const mk = 'v' + key;
    const mask = rcGet(p.cells, mk);
    setCss(mask, `left:${x - ov.l}px;top:${y}px;width:${ov.l + w + ov.r - 1}px;height:${hh - 1}px;`);
    setCls(mask, 'fv-ovmask');
    setText(mask, ''); // a pooled element may still hold a former cell's text
    el._xk = [mk];
  } else el._xk = undefined;
}

/**
 * Effective horizontal alignment of a cell's text. General: numbers and dates right, booleans and errors centred,
 * text left (right when it starts with a right-to-left script) — rotated text right for downward (negative) angles
 * and 90°, stacked text centred (Excel).
 * @param {StyleInfo} S @param {CellData | null} cell
 * @returns {'left' | 'center' | 'right' | 'justify' | 'distributed'}
 */
function cellAlign(S, cell) {
  switch (S.h) {
    case 'fill':
      return 'left';
    case 'centerContinuous':
      return 'center';
    case undefined:
      if (S.rot && (!cell || cell.t === 's')) return S.rot === 'vertical' ? 'center' : S.rot === 90 || (S.rot < 0 && S.rot !== -90) ? 'right' : 'left';
      return generalAlign(cell);
    default:
      return S.h;
  }
}

/** Left ↔ right (layout inside the mirrored scroller of a right-to-left sheet). @param {ReturnType<typeof cellAlign>} h */
function mirrorAlign(h) {
  return h === 'left' ? 'right' : h === 'right' ? 'left' : h;
}

/** Single-line text fits `avail` px (cheap upper bound first, then the cached canvas measure). */
function textFits(/** @type {string} */ text, /** @type {RichRun[] | null} */ rich, /** @type {StyleInfo} */ S, /** @type {number} */ avail) {
  if (!rich && text.length * S.fontPx * 1.2 <= avail) return true;
  return (rich ? richTextWidth(rich, S) : measureText(S.font, text)) <= avail;
}

/** Excel's precedence of line styles on a shared edge, lowest first (measured in Excel 16; ties: right/lower cell). */
const BORDER_RANK = ['hair', 'dashDotDot', 'dashDot', 'dotted', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];

/**
 * @typedef {object} Borders
 * @property {Edge | null} left
 * @property {Edge | null} top
 * @property {Edge | null} right
 * @property {Edge | null} bottom
 * @property {number} extL   the box extends one pixel left, over the neighbour's gridline
 * @property {number} extT
 * @property {string} [segs] markup of per-cell border segments (merged areas whose sides are not uniform)
 * @property {boolean} [rMixed] the right side is drawn as segments
 * @property {boolean} [bMixed] the bottom side is drawn as segments
 */

/**
 * Borders a cell draws. A shared edge is drawn by exactly one of its two cells: Excel shows the heavier line
 * style (BORDER_RANK), on a tie the right/lower cell's left/top border. This also holds across the frozen-pane
 * edge, where the two cells live in different panes. Left/top edges (and fills, like Excel) extend one pixel over
 * the neighbour's gridline, except across the frozen-pane edge.
 * @param {SheetView} view @param {Layout} L @param {number} r @param {number} c @param {StyleInfo} S
 * @param {CellData[]} rowCells
 * @returns {Borders}
 */
function cellBorders(view, L, r, c, S, rowCells) {
  /** @type {Borders} */
  const b = { left: null, top: null, right: S.right, bottom: S.bottom, extL: 0, extT: 0 };
  if (S.left || S.hasFill) {
    const pc = stepCol(L, c, -1);
    const nb = pc !== c ? edgeAt(view, L, r, pc, 'right', rowCells) : null;
    if (S.left && !(nb && nb.rank > S.left.rank)) b.left = S.left;
    if (pc !== c && !(c >= L.fc && pc < L.fc) && (b.left || (S.hasFill && !nb))) b.extL = 1;
  }
  if (S.top || S.hasFill) {
    const pr = stepRow(L, r, -1);
    const nb = pr !== r ? edgeAt(view, L, pr, c, 'bottom', null) : null;
    if (S.top && !(nb && nb.rank > S.top.rank)) b.top = S.top;
    if (pr !== r && !(r >= L.fr && pr < L.fr) && (b.top || (S.hasFill && !nb))) b.extT = 1;
  }
  if (b.right) {
    const nc = stepCol(L, c, 1);
    const nb = nc !== c ? edgeAt(view, L, r, nc, 'left', rowCells) : null;
    if (nb && nb.rank >= b.right.rank) b.right = null;
  }
  if (b.bottom) {
    const nr = stepRow(L, r, 1);
    const nb = nr !== r ? edgeAt(view, L, nr, c, 'top', null) : null;
    if (nb && nb.rank >= b.bottom.rank) b.bottom = null;
  }
  return b;
}

/** Perimeter cells examined per side of a merged area (unexamined ones are assumed to match the side). */
const MERGE_SIDE_CELLS = 1024;

/**
 * Outer borders of a merged area. Excel draws them per cell along the perimeter — the top edge from the cells of
 * the first row, the right edge from the last column, and so on — so a box applied in Excel (each edge cell carries
 * its side) is closed, while a box stored on the master only (other generators) shows just the master's top and
 * left segments. A side whose segments agree is one CSS border; otherwise the segments (and, on an unfilled area,
 * the gridline where a segment has no border) are drawn as children of the area's box. Shared edges with the
 * neighbours resolve like cellBorders.
 * @param {SheetView} view @param {Layout} L @param {Range} m @param {number} rEnd @param {number} cEnd
 * @param {StyleInfo} S master style (fill) @param {CellData[]} rowCells cells of row m.r0
 * @returns {Borders}
 */
function mergeBorders(view, L, m, rEnd, cEnd, S, rowCells) {
  const { r0, c0 } = m;
  // Hidden rows/columns draw nothing: the sides come from the first/last VISIBLE row and column of the area.
  const rTop = rowHeight(L, r0) ? r0 : Math.min(rEnd, stepRow(L, r0, 1));
  const rBot = rowHeight(L, rEnd) ? rEnd : Math.max(r0, stepRow(L, rEnd, -1));
  const cLeft = colWidth(L, c0) ? c0 : Math.min(cEnd, stepCol(L, c0, 1));
  const cRight = colWidth(L, cEnd) ? cEnd : Math.max(c0, stepCol(L, cEnd, -1));
  const pr = stepRow(L, rTop, -1);
  const nr = stepRow(L, rBot, 1);
  const pc = stepCol(L, cLeft, -1);
  const nc = stepCol(L, cRight, 1);
  /** @type {Record<string, CellData[] | null>} */
  const rows = {};
  /** Cells of row r, null when the row is not loaded. @param {number} r */
  const rowOf = (r) => {
    if (r === r0) return rowCells;
    if (!(r in rows)) {
      const row = getRow(view.index, r);
      rows[r] = row === undefined ? null : row ? row.cells : NO_CELLS;
    }
    return rows[r];
  };
  /**
   * Segments of one side: [index, own edge (undefined = row not loaded)] after the shared-edge rule.
   * @param {'top' | 'right' | 'bottom' | 'left'} side
   */
  const side = (side) => {
    const vertical = side === 'left' || side === 'right';
    const [a, z] = vertical ? [r0, rEnd] : [c0, cEnd];
    /** @type {[number, Edge | null | undefined][]} */
    const segs = [];
    for (let i = a, n = 0; i <= z && n < MERGE_SIDE_CELLS; i++, n++) {
      if (vertical ? !rowHeight(L, i) : !colWidth(L, i)) continue;
      const r = vertical ? i : side === 'top' ? rTop : rBot;
      const c = vertical ? (side === 'left' ? cLeft : cRight) : i;
      const cells = rowOf(r);
      /** @type {Edge | null | undefined} */
      let e = cells ? edgeAt(view, L, r, c, side, cells) : undefined;
      if (e) {
        // The neighbour across this segment: who wins the shared edge?
        const [nr2, nc2, other, later] =
          side === 'top' ? [pr, c, 'bottom', true] : side === 'bottom' ? [nr, c, 'top', false] : side === 'left' ? [r, pc, 'right', true] : [r, nc, 'left', false];
        if (nr2 !== r || nc2 !== c) {
          const nCells = nr2 === r ? cells : rowOf(nr2);
          const nb = nCells ? edgeAt(view, L, nr2, nc2, /** @type {'top' | 'right' | 'bottom' | 'left'} */ (other), nCells) : null;
          if (nb && (later ? nb.rank > e.rank : nb.rank >= e.rank)) e = null;
        }
      }
      segs.push([i, e]);
    }
    let uniform = true;
    /** @type {Edge | null | undefined} */
    let first;
    for (const [, e] of segs) {
      if (e === undefined) continue;
      if (first === undefined) first = e;
      else if ((e ? e.css : '') !== (first ? first.css : '')) uniform = false;
    }
    return { segs, uniform, edge: first ?? null };
  };
  const sides = { top: side('top'), right: side('right'), bottom: side('bottom'), left: side('left') };
  /** @type {Borders} */
  const b = {
    left: sides.left.uniform ? sides.left.edge : null,
    top: sides.top.uniform ? sides.top.edge : null,
    right: sides.right.uniform ? sides.right.edge : null,
    bottom: sides.bottom.uniform ? sides.bottom.edge : null,
    extL: 0,
    extT: 0,
  };
  const leftMixed = !sides.left.uniform;
  const topMixed = !sides.top.uniform;
  if (pc !== cLeft && !(cLeft >= L.fc && pc < L.fc)) {
    const cells = rowOf(rTop);
    if (b.left || leftMixed || (S.hasFill && !(cells && edgeAt(view, L, rTop, pc, 'right', cells)))) b.extL = 1;
  }
  if (pr !== rTop && !(rTop >= L.fr && pr < L.fr)) {
    const above = rowOf(pr);
    if (b.top || topMixed || (S.hasFill && !(above && edgeAt(view, L, pr, cLeft, 'bottom', above)))) b.extT = 1;
  }
  b.rMixed = !sides.right.uniform;
  b.bMixed = !sides.bottom.uniform;
  if (leftMixed || topMixed || b.rMixed || b.bMixed) {
    // Children of the area's box are positioned against its padding box (inside the uniform CSS borders). Where a
    // segment has no border, an unfilled area shows the gridline (left/top: only when the box covers it).
    const ox = b.left ? b.left.w : 0;
    const oy = b.top ? b.top.w : 0;
    const gaps = L.showGrid && !S.hasFill;
    let out = '';
    /**
     * @param {string} pos @param {Edge | null | undefined} e @param {string} prop border-* property
     * @param {boolean} across horizontal segment @param {boolean} gl draw the gridline when there is no border
     */
    const seg = (pos, e, prop, across, gl) => {
      if (e) out += `<div class="fv-bseg" style="${pos}${across ? 'height' : 'width'}:${e.w}px;${prop}:${e.css}"></div>`;
      else if (gl) out += `<div class="fv-bseg fv-bseg-gl" style="${pos}${across ? 'height' : 'width'}:1px"></div>`;
    };
    const atX = (/** @type {number} */ i) => `left:${b.extL + L.colLeft[i] - L.colLeft[c0] - ox}px;width:${colWidth(L, i)}px;`;
    const atY = (/** @type {number} */ i) => `top:${b.extT + L.rowTop[i] - L.rowTop[r0] - oy}px;height:${rowHeight(L, i)}px;`;
    if (topMixed) for (const [i, e] of sides.top.segs) seg(`top:0;${atX(i)}`, e, 'border-top', true, gaps && !!b.extT);
    if (leftMixed) for (const [i, e] of sides.left.segs) seg(`left:0;${atY(i)}`, e, 'border-left', false, gaps && !!b.extL);
    if (b.bMixed) for (const [i, e] of sides.bottom.segs) seg(`bottom:0;${atX(i)}`, e, 'border-bottom', true, gaps);
    if (b.rMixed) for (const [i, e] of sides.right.segs) seg(`right:0;${atY(i)}`, e, 'border-right', false, gaps);
    b.segs = out;
  }
  return b;
}

/**
 * The border a cell draws on one side by its own style — merged areas are NOT resolved to their master (Excel draws
 * a merge's outer edges from its edge cells). null = none, undefined = the row is not loaded yet.
 * @param {SheetView} view @param {Layout} L @param {number} r @param {number} c
 * @param {'top' | 'right' | 'bottom' | 'left'} side @param {CellData[] | null} rowCells cells of row r, if at hand
 * @returns {Edge | null | undefined}
 */
function edgeAt(view, L, r, c, side, rowCells) {
  let cells = rowCells;
  if (!cells) {
    const row = getRow(view.index, r);
    if (row === undefined) return undefined;
    cells = row ? row.cells : NO_CELLS;
  }
  const cell = findCell(cells, c);
  const s = cell ? cell.s || 0 : rowStyleIndex(view.meta, r) || L.colStyle[c] || 0;
  return s ? styleInfo(s)[side] : null;
}

/**
 * Text spilling out of its cell: px into the empty neighbours on the left/right, `fills` = cell-relative x ranges
 * [x0, x1) of the source and the spanned neighbours that have a fill, `across` = Center Across Selection box.
 * @typedef {{ l: number, r: number, fills: [number, number][], across?: boolean }} Overflow
 */

/**
 * How far non-wrapped text spills into empty neighbours (Excel: until the next non-empty cell).
 * Text width is measured with a cached canvas context — no DOM measuring.
 * @param {SheetView} view @param {Layout} L @param {number} r @param {number} c @param {number} w
 * @param {string} text single-line display text @param {StyleInfo} S @param {string} hEff @param {CellData[]} rowCells
 * @param {RichRun[] | null} rich
 * @returns {Overflow | null}
 */
function computeOverflow(view, L, r, c, w, text, S, hEff, rowCells, rich) {
  if (hEff !== 'left' && hEff !== 'center' && hEff !== 'right') return null;
  // Cheap upper bound first (no glyph of a common font is wider than ~1.2em), then measure.
  if (!rich && text.length * S.fontPx * 1.2 + 4 + S.indent <= w) return null;
  const tw = (rich ? richTextWidth(rich, S) : measureText(S.font, text)) + 5 + S.indent;
  if (tw <= w) return null;
  const need = hEff === 'center' ? (tw - w) / 2 : tw - w;
  const lastOwn = c < L.fc ? L.fc - 1 : L.nCols - 1;
  const firstOwn = c < L.fc ? 0 : L.fc;
  /** @type {[number, number][]} */
  const fills = S.hasFill ? [[0, w]] : [];
  let right = 0;
  let leftExt = 0;
  if (hEff === 'left' || hEff === 'center') {
    for (let cc = c + 1, n = 0; cc <= lastOwn && right < need && n < 256; cc++, n++) {
      const cw = colWidth(L, cc);
      if (!cw) continue;
      if (!overflowTarget(view, L, r, cc, rowCells)) break;
      if (filledAt(view, L, r, cc, rowCells)) fills.push([w + right, w + right + cw]);
      right += cw;
    }
  }
  if (hEff === 'right' || hEff === 'center') {
    for (let cc = c - 1, n = 0; cc >= firstOwn && leftExt < need && n < 256; cc--, n++) {
      const cw = colWidth(L, cc);
      if (!cw) continue;
      if (!overflowTarget(view, L, r, cc, rowCells)) break;
      leftExt += cw;
      if (filledAt(view, L, r, cc, rowCells)) fills.push([-leftExt, -leftExt + cw]);
    }
  }
  if (!right && !leftExt) return null;
  return { l: leftExt, r: right, fills: mergeRanges(fills) };
}

/**
 * Center Across Selection: the text is centred over the cell and the following empty cells that carry the same
 * alignment (null when there are none — the cell then behaves like centred text).
 * @param {SheetView} view @param {Layout} L @param {number} r @param {number} c @param {number} w
 * @param {StyleInfo} S @param {CellData[]} rowCells
 * @returns {Overflow | null}
 */
function centerAcross(view, L, r, c, w, S, rowCells) {
  const lastOwn = c < L.fc ? L.fc - 1 : L.nCols - 1;
  /** @type {[number, number][]} */
  const fills = S.hasFill ? [[0, w]] : [];
  let right = 0;
  for (let cc = c + 1, n = 0; cc <= lastOwn && n < 256; cc++, n++) {
    const cw = colWidth(L, cc);
    if (!cw) continue;
    if (!overflowTarget(view, L, r, cc, rowCells)) break;
    const cell = findCell(rowCells, cc);
    const s = cell ? cell.s || 0 : rowStyleIndex(view.meta, r) || L.colStyle[cc] || 0;
    if (!s || styleInfo(s).h !== 'centerContinuous') break;
    if (styleInfo(s).hasFill) fills.push([w + right, w + right + cw]);
    right += cw;
  }
  return right ? { l: 0, r: right, fills: mergeRanges(fills), across: true } : null;
}

/** Empty, unmerged cell (a formula returning "" is not empty: it blocks overflow like in Excel). @param {SheetView} view @param {Layout} L @param {number} r @param {number} c @param {CellData[]} rowCells */
function overflowTarget(view, L, r, c, rowCells) {
  const cell = findCell(rowCells, c);
  if (cell && (!isEmptyCell(cell) || cell.f)) return false;
  return !(view.mergeIndex.size && mergeAt(view, r, c) >= 0);
}

/** Cell paints a fill of its own (or of its row/column). @param {SheetView} view @param {Layout} L @param {number} r @param {number} c @param {CellData[]} rowCells */
function filledAt(view, L, r, c, rowCells) {
  const cell = findCell(rowCells, c);
  const s = cell ? cell.s || 0 : rowStyleIndex(view.meta, r) || L.colStyle[c] || 0;
  return s !== 0 && styleInfo(s).hasFill;
}

/** Sorted union of [x0, x1) ranges. @param {[number, number][]} ranges @returns {[number, number][]} */
function mergeRanges(ranges) {
  ranges.sort((a, b) => a[0] - b[0]);
  /** @type {[number, number][]} */
  const out = [];
  for (const [a, b] of ranges) {
    const last = out[out.length - 1];
    if (last && a <= last[1]) last[1] = Math.max(last[1], b);
    else out.push([a, b]);
  }
  return out;
}

/**
 * Text layers of an overflow box. In dark themes the theme colours of unfilled cells (light automatic text,
 * lightened dark colours) would be unreadable over a light fill, and the workbook colours over the editor
 * background: when the spanned cells mix filled and unfilled ones, the text is drawn twice in the same place,
 * each copy clipped to the cells it suits (workbook colours over fills, theme colours elsewhere).
 * @param {Overflow} ov @param {number} w @param {StyleInfo} S @param {string | undefined} fmtColor number-format colour
 * @param {(hasFill: boolean) => string} body text markup for a background @param {(s: string) => string} glyphs
 * @returns {{ html: string, color?: string, clip?: [number, number][] }[]}
 */
function overflowLayers(ov, w, S, fmtColor, body, glyphs) {
  const g = /** @type {GridState} */ (G);
  const span = /** @type {[number, number]} */ ([-ov.l, w + ov.r]);
  const allFilled = ov.fills.length === 1 && ov.fills[0][0] <= span[0] && ov.fills[0][1] >= span[1];
  if (g.dark && ov.fills.length && !allFilled) {
    const color = fmtColor || S.fontColor;
    const onFill = cellFontColor(color, true);
    const offFill = cellFontColor(color, false) || 'var(--fv-cell-fg)';
    const plain = complementRanges(ov.fills);
    return [
      { html: glyphs(body(true)), color: onFill, clip: ov.fills },
      { html: glyphs(body(false)), color: offFill, clip: plain },
    ];
  }
  return [{ html: glyphs(body(S.hasFill)) }];
}

/** Complement of sorted, disjoint ranges over the whole line. @param {[number, number][]} ranges */
function complementRanges(ranges) {
  const E = 99999;
  /** @type {[number, number][]} */
  const out = [];
  let x = -E;
  for (const [a, b] of ranges) {
    if (a > x) out.push([x, a]);
    x = Math.max(x, b);
  }
  if (x < E) out.push([x, E]);
  return out;
}

/** clip-path keeping only the given x ranges of a box (vertically unbounded). @param {[number, number][]} ranges */
function clipXRanges(ranges) {
  const E = 99999;
  const pts = [];
  for (const [a, b] of ranges) pts.push(`${a}px ${-E}px`, `${b}px ${-E}px`, `${b}px ${E}px`, `${a}px ${E}px`, `${a}px ${-E}px`);
  return `polygon(evenodd,${pts.join(',')})`;
}

/**
 * Markup of overflowing text: an absolutely positioned box spanning the neighbours, positioned relative to the
 * cell's padding box, holding one or more text layers. Centred text is centred on the cell (on the whole box for
 * Center Across Selection); text longer than the box keeps its end (right) or middle (centre) visible because the
 * shrink-wrapped text block overflows its flex layer at the start side.
 * @param {Overflow} ov @param {number} w @param {number} hh
 * @param {string} hEff @param {StyleInfo} S
 * @param {{ html: string, color?: string, clip?: [number, number][] }[]} layers clip ranges are cell-relative
 * @param {number} dx offset of the cell's padding box from its left edge @param {number} dy
 */
function overflowHtml(ov, w, hh, hEff, S, layers, dx, dy) {
  const total = ov.l + w + ov.r;
  let iw = total;
  let ml = 0;
  if (hEff === 'center' && !ov.across) {
    const half = Math.max(ov.l, ov.r);
    iw = w + 2 * half;
    ml = ov.l - half;
  }
  const pad = hEff === 'left' ? `padding-left:${2 + S.indent}px;` : hEff === 'right' ? `padding-right:${2 + S.indent}px;` : '';
  const self = hEff === 'right' ? 'flex-end' : hEff === 'center' ? 'center' : 'flex-start';
  const box = `left:${ml}px;width:${iw}px;${pad}`;
  const shift = ov.l - ml; // cell-relative x -> layer x
  let out = '';
  for (const layer of layers) {
    const clip = layer.clip ? `clip-path:${clipXRanges(layer.clip.map(([a, b]) => [a + shift, b + shift]))};` : '';
    const color = layer.color ? `color:${layer.color};` : '';
    out += `<div class="fv-ovi" style="${box}${color}${clip}"><div class="fv-tx fv-clip" style="align-self:${self}">${layer.html}</div></div>`;
  }
  return `<div class="fv-ov" style="left:${dx - ov.l}px;top:${dy}px;width:${total}px;height:${hh}px">${out}</div>`;
}

/**
 * Rotated / stacked text. Angles other than ±90° lay out the rotated text's bounding box (tw × lh rotated) in the
 * cell like any text block, so the rotated line stays inside it (Excel). On mirrored (right-to-left) sheets the
 * text flips back so glyphs stay readable; like in Excel, ±90° text keeps its reading direction and other angles
 * are mirrored (30° descends to the right).
 * @param {number | 'vertical'} rot @param {string} hEff @param {string} inner @param {boolean} rtl
 * @param {number} tw single-line text width @param {number} lh line height
 */
function rotatedHtml(rot, hEff, inner, rtl, tw, lh) {
  const self = hEff === 'center' || hEff === 'justify' || hEff === 'distributed' ? 'center' : hEff === 'right' ? 'flex-end' : 'flex-start';
  if (rot === 'vertical') return `<div class="fv-tx fv-stack" style="align-self:${self}">${inner}</div>`;
  if (rot === 90) return `<div class="fv-tx fv-r90" style="align-self:${self}">${inner}</div>`;
  if (rot === -90) return `<div class="fv-tx fv-r-90" style="align-self:${self}">${inner}</div>`;
  const a = (Math.abs(rot) * Math.PI) / 180;
  const bw = Math.ceil(tw * Math.cos(a) + lh * Math.sin(a));
  const bh = Math.ceil(tw * Math.sin(a) + lh * Math.cos(a));
  return (
    `<div class="fv-tx fv-rotw" style="align-self:${self};width:${bw}px;height:${bh}px">` +
    `<div class="fv-rot" style="transform:translate(-50%,-50%) ${rtl ? `scaleX(-1) rotate(${rot}deg)` : `rotate(${-rot}deg)`}">${inner}</div></div>`
  );
}

/**
 * Glyph sets of formatted numbers: plain ones (digits, separators, minus, spaces) and all of them (plus sign, percent,
 * parentheses, exponent, time / fraction separators). The set's widest glyph times the length bounds a number's
 * width without measuring the text (most cells hold plain numbers, bounded by the digit width).
 */
const DIGIT_CHARS = '0123456789.,- \u00a0';
const NUMBER_CHARS = DIGIT_CHARS + ':%+()Ee/';
const DIGIT_TEXT_RE = /^[\d.,\- \u00a0]*$/;
const NUMBER_TEXT_RE = /^[\d\s.,:%+\-()E/]*$/i;

/** Widest glyph of DIGIT_CHARS / NUMBER_CHARS per canvas font. @type {Map<string, number>[]} */
const maxCharCaches = [new Map(), new Map()];

/** Width of the widest glyph of a set (0 = DIGIT_CHARS, 1 = NUMBER_CHARS) in a canvas font, cached per font. */
function maxCharWidth(/** @type {string} */ font, /** @type {0 | 1} */ set) {
  const cache = maxCharCaches[set];
  let w = cache.get(font);
  if (w === undefined) {
    w = 0;
    for (const ch of set ? NUMBER_CHARS : DIGIT_CHARS) w = Math.max(w, measureText(font, ch));
    if (cache.size > 500) cache.clear();
    cache.set(font, w);
  }
  return w;
}

/**
 * Display text of a value (number, date, boolean, error) that fits the cell width, like Excel: General-format
 * numbers are rounded to fewer decimals, then shown in scientific notation; anything else that does not fit
 * becomes '###…'. Measured with the cached canvas only when a safe upper bound (the font's widest number glyph
 * times the length) says the text might not fit.
 * @param {CellData} cell @param {string} text @param {StyleInfo} S @param {number} w cell width
 */
function fitNumber(cell, text, S, w) {
  const avail = w - 4 - S.indent;
  const charW = DIGIT_TEXT_RE.test(text) ? maxCharWidth(S.font, 0) : NUMBER_TEXT_RE.test(text) ? maxCharWidth(S.font, 1) : S.fontPx * 1.25;
  if (text.length * charW <= avail) return text;
  if (avail <= 0) return '';
  const fits = (/** @type {string} */ s) => measureText(S.font, s) <= avail;
  if (fits(text)) return text;
  const g = /** @type {GridState} */ (G);
  const style = g.styles[cell.s || 0];
  const general = !style || !style.numFmt || style.numFmt === 'General';
  if (general && cell.t === 'n' && typeof cell.v === 'number' && Number.isFinite(cell.v)) {
    const v = cell.v;
    const intDigits = Math.abs(v) >= 1 ? Math.floor(Math.log10(Math.abs(v))) + 1 : 1;
    if (intDigits <= 11) {
      for (let d = Math.min(10, Math.max(0, 11 - intDigits)); d >= 0; d--) {
        const s = formatGeneral(Number(v.toFixed(d)));
        if (fits(s)) return s;
      }
    }
    for (let k = 5; k >= 0; k--) {
      const s = v
        .toExponential(k)
        .replace(/\.?0+e/, 'e')
        .replace(/e([+-])(\d)$/, 'e$10$2')
        .toUpperCase();
      if (fits(s)) return s;
    }
  }
  const hash = measureText(S.font, '#');
  return '#'.repeat(Math.max(1, Math.floor(avail / Math.max(1, hash))));
}

/**
 * A value whose number format repeats a character to fill the cell (CellData.fill), laid out like Excel: the text
 * before the fill at the left edge, the text after it at the right edge, as many whole copies of the fill character
 * as fit in between. '_x' paddings (spaces in the text) become invisible copies of x, so they are exactly as wide as
 * x (`_)` lines positive numbers up with the ')' of negative ones).
 * @param {string} text CellData.w @param {FillLayout} fill @param {StyleInfo} S @param {number} avail content width (px)
 */
function fillHtml(text, fill, S, avail) {
  const at = clamp(fill.at, 0, text.length);
  const pads = Array.isArray(fill.pads) ? fill.pads.filter((p) => Array.isArray(p) && text[p[0]] === ' ' && typeof p[1] === 'string' && p[1]) : [];
  let shown = '';
  /** Markup of text[from, to) with its paddings. @param {number} from @param {number} to */
  const part = (from, to) => {
    let out = '';
    let i = from;
    for (const [p, ch] of pads) {
      if (p < from || p >= to) continue;
      out += escapeHtml(text.slice(i, p)) + `<span class="fv-pad">${escapeHtml(ch)}</span>`;
      shown += text.slice(i, p) + ch;
      i = p + 1;
    }
    shown += text.slice(i, to);
    return out + escapeHtml(text.slice(i, to));
  };
  const before = part(0, at);
  const after = part(at, text.length);
  const ch = typeof fill.char === 'string' && fill.char ? fill.char : ' ';
  const room = avail - measureText(S.font, shown);
  const one = ch === ' ' ? 0 : measureText(S.font, ch);
  const copies = one > 0 && room > 0 ? Math.min(2000, Math.floor(room / one)) : 0;
  return `<div class="fv-tx fv-fill"><span>${before}</span><span class="fv-fillc">${escapeHtml(ch.repeat(copies))}</span><span>${after}</span></div>`;
}

/** Line breaks → spaces (non-wrapped cells show their text on one line). @param {string} s */
function oneLine(s) {
  return s.replace(/\r\n|[\r\n]/g, ' ');
}

/**
 * Width of rich text, each run measured with its own font layered over the cell font.
 * @param {RichRun[]} runs @param {StyleInfo} S
 */
function richTextWidth(runs, S) {
  const g = /** @type {GridState} */ (G);
  const z = g.styleZoom / 100;
  let w = 0;
  for (const run of runs) {
    const f = run.font;
    let font = S.font;
    if (f && (f.name || f.size || f.bold !== undefined || f.italic !== undefined || f.vertAlign)) {
      const italic = f.italic ?? S.italic;
      const bold = f.bold ?? S.bold;
      let px = f.size ? ptToPx(f.size, z) : S.fontPx;
      if (f.vertAlign) px *= 0.75;
      font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${px}px ${f.name ? fontFamilyCss(f.name) : S.family}`;
    }
    w += measureText(font, oneLine(run.text || ''));
  }
  return w;
}

/**
 * @param {RichRun[]} runs @param {boolean} hasFill @param {boolean} [singleLine] line breaks shown as spaces
 * @param {number} [scale] shrink-to-fit factor for runs with their own size
 */
function richHtml(runs, hasFill, singleLine = false, scale = 1) {
  const g = /** @type {GridState} */ (G);
  const z = g.styleZoom / 100;
  let out = '';
  for (const run of runs) {
    const f = run.font;
    let css = '';
    if (f) {
      if (f.name) css += `font-family:${fontFamilyCss(f.name)};`;
      // Super/subscript runs are 3/4 of their own size (of the cell's when they have none).
      const px = f.size ? Math.round(ptToPx(f.size, z) * scale * (f.vertAlign ? 0.75 : 1) * 100) / 100 : 0;
      if (px) css += `font-size:${px}px;`;
      else if (f.vertAlign) css += 'font-size:0.75em;';
      if (f.bold) css += 'font-weight:700;';
      else if (f.bold === false) css += 'font-weight:400;';
      if (f.italic) css += 'font-style:italic;';
      else if (f.italic === false) css += 'font-style:normal;';
      css += textDecorationCss(f);
      if (f.color) {
        const color = cellFontColor(f.color, hasFill);
        css += color ? `color:${color};` : 'color:var(--fv-cell-fg);';
      }
      if (f.vertAlign === 'superscript') css += 'vertical-align:super;';
      else if (f.vertAlign === 'subscript') css += 'vertical-align:sub;';
    }
    let t = escapeHtml(singleLine ? oneLine(run.text || '') : run.text);
    if (f && strikeInside(f)) t = `<span class="fv-strike">${t}</span>`;
    out += css ? `<span style="${escapeHtml(css)}">${t}</span>` : t;
  }
  return out;
}

/**
 * Gridlines and "loading" placeholders for unloaded rows of one pane.
 * @param {Pane} p @param {SheetView} view @param {Layout} L
 * @param {number} ra @param {number} rb @param {number} ca @param {number} cb
 */
function renderPaneGrid(p, view, L, ra, rb, ca, cb) {
  const rc = p.grid;
  rcBegin(rc);
  if (ra <= rb && ca <= cb) {
    const x0 = L.hdrW + L.colLeft[ca] - p.ox;
    const x1 = L.hdrW + L.colLeft[cb + 1] - p.ox;
    if (L.showGrid) {
      // Line lengths are snapped to row/column chunks so that scrolling does not restyle every line; the
      // overhang lies outside the viewport or under the opaque bands (which clip their own panes).
      const gx0 = L.hdrW + L.colLeft[Math.floor(ca / GRID_SNAP_COLS) * GRID_SNAP_COLS] - p.ox;
      const gx1 = L.hdrW + L.colLeft[Math.min(L.nCols, Math.ceil((cb + 1) / GRID_SNAP_COLS) * GRID_SNAP_COLS)] - p.ox;
      const gy0 = L.hdrH + L.rowTop[Math.floor(ra / GRID_SNAP_ROWS) * GRID_SNAP_ROWS] - p.oy;
      const gy1 = L.hdrH + L.rowTop[Math.min(L.nRows, Math.ceil((rb + 1) / GRID_SNAP_ROWS) * GRID_SNAP_ROWS)] - p.oy;
      for (let r = ra; r <= rb; r++) {
        if (!rowHeight(L, r)) continue;
        const el = rcGet(rc, 'h' + r);
        setCss(el, `left:${gx0}px;top:${L.hdrH + L.rowTop[r + 1] - 1 - p.oy}px;width:${gx1 - gx0}px;height:1px;`);
        setCls(el, 'fv-gl');
      }
      for (let c = ca; c <= cb; c++) {
        if (!colWidth(L, c)) continue;
        const el = rcGet(rc, 'v' + c);
        setCss(el, `left:${L.hdrW + L.colLeft[c + 1] - 1 - p.ox}px;top:${gy0}px;width:1px;height:${gy1 - gy0}px;`);
        setCls(el, 'fv-gl');
      }
    }
    let seg = -1;
    for (let r = ra; r <= rb + 1; r++) {
      const loading = r <= rb && getRow(view.index, r) === undefined;
      if (loading && seg < 0) seg = r;
      else if (!loading && seg >= 0) {
        const el = rcGet(rc, 'L' + seg);
        setCss(el, `left:${x0}px;top:${L.hdrH + L.rowTop[seg] - p.oy}px;width:${x1 - x0}px;height:${L.rowTop[r] - L.rowTop[seg]}px;`);
        setCls(el, 'fv-loading-rows');
        seg = -1;
      }
    }
  }
  rcEnd(rc);
}

/**
 * Column/row headers of a pane.
 * @param {Pane} p @param {Layout} L
 * @param {[number, number] | null} cols @param {[number, number] | null} rows
 */
function renderHeaders(p, L, cols, rows) {
  const g = /** @type {GridState} */ (G);
  const sig = `${cols},${rows},${g.layoutVersion},${g.selVersion}`;
  if (p.sigHdr === sig) return;
  p.sigHdr = sig;
  const rc = p.hdrs;
  rcBegin(rc);
  if (cols) {
    for (let c = cols[0]; c <= cols[1]; c++) {
      const w = colWidth(L, c);
      if (!w) continue;
      const el = rcGet(rc, 'c' + c);
      setCss(el, `left:${L.hdrW + L.colLeft[c] - p.ox}px;top:${-p.oy}px;width:${w}px;height:${L.hdrH}px;`);
      setCls(el, 'fv-ch' + headerClass(g.hdrSel.cols, c));
      if (L.rtl) setHtml(el, `<span class="fv-g">${colName(c)}</span>`);
      else setText(el, colName(c));
    }
  }
  if (rows) {
    for (let r = rows[0]; r <= rows[1]; r++) {
      const hh = rowHeight(L, r);
      if (!hh) continue;
      const el = rcGet(rc, 'r' + r);
      setCss(el, `left:${-p.ox}px;top:${L.hdrH + L.rowTop[r] - p.oy}px;width:${L.hdrW}px;height:${hh}px;`);
      setCls(el, 'fv-rh' + headerClass(g.hdrSel.rows, r));
      if (L.rtl) setHtml(el, `<span class="fv-g">${r + 1}</span>`);
      else setText(el, String(r + 1));
    }
  }
  rcEnd(rc);
}

/** @param {[number, number, boolean][]} spans @param {number} i */
function headerClass(spans, i) {
  let cls = '';
  for (const [a, b, full] of spans) {
    if (i < a || i > b) continue;
    if (full) return ' fv-hfull';
    cls = ' fv-hsel';
  }
  return cls;
}

/** Row/column spans of the selection for header highlighting. @param {SheetView} view */
function computeHeaderSel(view) {
  const g = /** @type {GridState} */ (G);
  /** @type {[number, number, boolean][]} */
  const rows = [];
  /** @type {[number, number, boolean][]} */
  const cols = [];
  for (const rg of view.sel.ranges) {
    rows.push([rg.r0, rg.r1, isFullRows(view, rg)]);
    cols.push([rg.c0, rg.c1, isFullCols(view, rg)]);
  }
  g.hdrSel = { rows, cols };
}

/**
 * Selection overlay of one pane: translucent fill (the active cell stays clear), a 2px border around a
 * single range, and the active-cell outline.
 * @param {Pane} p @param {SheetView} view @param {Layout} L
 * @param {{ x0: number, y0: number, x1: number, y1: number }} region
 */
function renderPaneSelection(p, view, L, region) {
  const rc = p.sel;
  rcBegin(rc);
  const sel = view.sel;
  const act = rangeRect(L, cellArea(view, sel.active.r, sel.active.c));
  const last = sel.ranges.length - 1;
  let k = 0;
  sel.ranges.forEach((rg, i) => {
    const rr = rangeRect(L, rg);
    if (!rr) return;
    const pieces = i === last && act ? subtractRect(rr, act) : [rr];
    for (const piece of pieces) drawSelRect(p, rc, 'f' + k++, piece, region, 'fv-sel-fill', false);
    if (last === 0 && act && !(rr.x0 === act.x0 && rr.x1 === act.x1 && rr.y0 === act.y0 && rr.y1 === act.y1)) {
      drawSelRect(p, rc, 'b', rr, region, 'fv-sel-border', true);
    }
  });
  if (act) {
    const single = last === 0 && sel.ranges[0] && sameRect(rangeRect(L, sel.ranges[0]), act);
    drawSelRect(p, rc, 'a', act, region, single ? 'fv-sel-border' : 'fv-sel-active', true);
  }
  rcEnd(rc);
}

/**
 * @param {{ x0: number, y0: number, x1: number, y1: number } | null} a
 * @param {{ x0: number, y0: number, x1: number, y1: number } | null} b
 */
function sameRect(a, b) {
  return !!a && !!b && a.x0 === b.x0 && a.y0 === b.y0 && a.x1 === b.x1 && a.y1 === b.y1;
}

/**
 * Content-coordinate rectangle of a range, clamped to the layout; null when nothing is visible.
 * @param {Layout} L @param {Range} rg
 */
function rangeRect(L, rg) {
  const r0 = clamp(rg.r0, 0, L.nRows - 1);
  const r1 = clamp(rg.r1, 0, L.nRows - 1);
  const c0 = clamp(rg.c0, 0, L.nCols - 1);
  const c1 = clamp(rg.c1, 0, L.nCols - 1);
  const rect = { x0: L.hdrW + L.colLeft[c0], y0: L.hdrH + L.rowTop[r0], x1: L.hdrW + L.colLeft[c1 + 1], y1: L.hdrH + L.rowTop[r1 + 1] };
  return rect.x1 > rect.x0 && rect.y1 > rect.y0 ? rect : null;
}

/**
 * r minus hole (up to four rectangles).
 * @param {{ x0: number, y0: number, x1: number, y1: number }} r
 * @param {{ x0: number, y0: number, x1: number, y1: number }} hole
 */
function subtractRect(r, hole) {
  if (hole.x1 <= r.x0 || hole.x0 >= r.x1 || hole.y1 <= r.y0 || hole.y0 >= r.y1) return [r];
  const out = [];
  if (hole.y0 > r.y0) out.push({ x0: r.x0, y0: r.y0, x1: r.x1, y1: hole.y0 });
  if (hole.y1 < r.y1) out.push({ x0: r.x0, y0: hole.y1, x1: r.x1, y1: r.y1 });
  const y0 = Math.max(r.y0, hole.y0);
  const y1 = Math.min(r.y1, hole.y1);
  if (hole.x0 > r.x0) out.push({ x0: r.x0, y0, x1: hole.x0, y1 });
  if (hole.x1 < r.x1) out.push({ x0: hole.x1, y0, x1: r.x1, y1 });
  return out;
}

/**
 * Draws a fill or a border box clipped to the pane's visible region (clipped sides lose their border).
 * Borders straddle the gridlines: the 2px range border runs from one pixel outside the range to its last
 * pixel (the range's own right/bottom gridline); the 1px active-cell outline sits exactly on the gridlines.
 * @param {Pane} p @param {Recycler} rc @param {string} key
 * @param {{ x0: number, y0: number, x1: number, y1: number }} rect
 * @param {{ x0: number, y0: number, x1: number, y1: number }} region
 * @param {string} cls @param {boolean} border
 */
function drawSelRect(p, rc, key, rect, region, cls, border) {
  const bw = border ? (cls === 'fv-sel-active' ? 1 : 2) : 0;
  let x0 = rect.x0 - bw;
  let y0 = rect.y0 - bw;
  let x1 = rect.x1 + (bw === 2 ? 1 : 0);
  let y1 = rect.y1 + (bw === 2 ? 1 : 0);
  const cl = x0 < region.x0;
  const ct = y0 < region.y0;
  const cr = x1 > region.x1;
  const cb = y1 > region.y1;
  x0 = Math.max(x0, region.x0);
  y0 = Math.max(y0, region.y0);
  x1 = Math.min(x1, region.x1);
  y1 = Math.min(y1, region.y1);
  if (x1 <= x0 || y1 <= y0) return;
  const el = rcGet(rc, key);
  let css = `left:${x0 - p.ox}px;top:${y0 - p.oy}px;width:${x1 - x0}px;height:${y1 - y0}px;`;
  if (bw && (cl || ct || cr || cb)) css += `border-width:${ct ? 0 : bw}px ${cr ? 0 : bw}px ${cb ? 0 : bw}px ${cl ? 0 : bw}px;`;
  setCss(el, css);
  setCls(el, cls);
}

/**
 * Images anchored to cells; each pane draws those intersecting its static area and visible region.
 * @param {Pane} p @param {SheetView} view @param {Layout} L
 * @param {{ x0: number, y0: number, x1: number, y1: number }} region
 */
function renderPaneImages(p, view, L, region) {
  const rc = p.imgs;
  rcBegin(rc);
  const right = p.name === 'main' || p.name === 'top';
  const bottom = p.name === 'main' || p.name === 'left';
  view.images.forEach((im, i) => {
    if (right ? im.x + im.w <= L.B : im.x >= L.B) return;
    if (bottom ? im.y + im.h <= L.A : im.y >= L.A) return;
    if (im.x + im.w < region.x0 || im.x > region.x1 || im.y + im.h < region.y0 || im.y > region.y1) return;
    const el = /** @type {RcEl & HTMLImageElement} */ (rcGet(rc, i));
    if (el._html !== im.src) {
      el._html = im.src;
      el.src = im.src;
      el.alt = '';
      el.draggable = false;
    }
    setCss(el, `left:${im.x - p.ox}px;top:${im.y - p.oy}px;width:${im.w}px;height:${im.h}px;`);
    setCls(el, 'fv-img');
  });
  rcEnd(rc);
}

// ----- RENDER: style table → CSS -----

/** Excel border line styles → [width px, CSS style]. 'hair' is a faint 1px line. */
/** @type {Record<string, [number, string]>} */
const BORDER_STYLES = {
  thin: [1, 'solid'],
  medium: [2, 'solid'],
  thick: [3, 'solid'],
  dotted: [1, 'dotted'],
  dashed: [1, 'dashed'],
  double: [3, 'double'],
  hair: [1, 'hair'],
  mediumDashed: [2, 'dashed'],
  dashDot: [1, 'dashed'],
  mediumDashDot: [2, 'dashed'],
  dashDotDot: [1, 'dotted'],
  mediumDashDotDot: [2, 'dashed'],
  slantDashDot: [2, 'dashed'],
};

/** Percentage of the pattern colour in Excel's gray patterns. */
/** @type {Record<string, number>} */
const GRAY_PATTERNS = { darkGray: 75, mediumGray: 50, lightGray: 25, gray125: 12.5, gray0625: 6.25 };

/** Metric-compatible fallbacks for common Office fonts. */
/** @type {Record<string, string[]>} */
const FONT_ALIASES = {
  calibri: ['Carlito'],
  'calibri light': ['Carlito'],
  cambria: ['Caladea'],
  arial: ['Liberation Sans', 'Arimo', 'Helvetica'],
  'times new roman': ['Liberation Serif', 'Tinos', 'Times'],
  'courier new': ['Liberation Mono', 'Cousine', 'Courier'],
  aptos: ['Segoe UI', 'Calibri', 'Carlito'],
  'aptos narrow': ['Arial Narrow', 'Calibri', 'Carlito'],
};

/** @type {Map<string, string>} */
const fontFamilyCache = new Map();

/** CSS font-family list for an Office font name (single-quoted, safe inside style attributes). @param {string} name */
function fontFamilyCss(name) {
  let css = fontFamilyCache.get(name);
  if (css) return css;
  const clean = name.replace(/['"\\;{}<>]/g, '').trim() || 'Calibri';
  const lower = clean.toLowerCase();
  const generic = /mono|courier|consolas|menlo|lucida console|fixedsys/.test(lower)
    ? 'monospace'
    : /times|cambria|georgia|garamond|palatino|book antiqua|bookman|century schoolbook|serif|mincho|song|ming/.test(lower) && !lower.includes('sans')
      ? 'serif'
      : 'sans-serif';
  css = [clean, ...(FONT_ALIASES[lower] || [])].map((f) => `'${f}'`).concat(generic).join(', ');
  fontFamilyCache.set(name, css);
  return css;
}

/** @param {number} pt @param {number} z */
function ptToPx(pt, z) {
  return Math.round(((pt * 4) / 3) * z * 100) / 100;
}

/** @param {string | undefined} color */
function normColor(color) {
  return typeof color === 'string' && color ? color.trim().toUpperCase() : undefined;
}

/**
 * Text colour: explicit colours are honoured; on an unfilled cell, black (the model's spelling of
 * "Automatic" for rich-text runs) and the workbook's default font colour count as automatic, which follows
 * the theme foreground so text stays readable in dark themes. Automatic text on a fill is black (Excel look).
 * @param {string | undefined} color @param {boolean} hasFill
 */
function cellFontColor(color, hasFill) {
  const g = /** @type {GridState} */ (G);
  const norm = normColor(color);
  if (norm && hasFill) return /** @type {string} */ (color);
  if (norm && norm !== '#000000' && norm !== g.defaultFontColor) return g.dark ? readableOnDark(/** @type {string} */ (color)) : /** @type {string} */ (color);
  return hasFill ? '#000000' : '';
}

/** Explicit colours already adjusted for the current dark editor background. @type {Map<string, string>} */
const darkColorCache = new Map();
/** Relative luminance of the editor background the cache was built for (-1 = not measured yet). */
let darkBgLuminance = -1;

/**
 * Dark themes paint unfilled cells with the editor background, so dark explicit colours (e.g. '#404040' labels,
 * navy borders, pure blue) would vanish. Colours with less than 3:1 contrast against the background get their
 * lightness mirrored (hue kept, like Office's dark-cell mode) and raised until they reach 4.5:1. Filled cells keep
 * the workbook's colours exactly.
 * @param {string} color
 */
function readableOnDark(color) {
  const norm = normColor(color);
  if (!norm || !/^#[0-9A-F]{6}$/.test(norm)) return color;
  let out = darkColorCache.get(norm);
  if (out) return out;
  if (darkBgLuminance < 0) darkBgLuminance = cssColorLuminance(getComputedStyle(document.documentElement).getPropertyValue('--vscode-editor-background')) ?? 0.012;
  const rgb = [1, 3, 5].map((i) => parseInt(norm.slice(i, i + 2), 16) / 255);
  const contrast = (/** @type {number} */ lum) => (Math.max(lum, darkBgLuminance) + 0.05) / (Math.min(lum, darkBgLuminance) + 0.05);
  if (contrast(relLuminance(rgb)) >= 3) out = color;
  else {
    const [hue, sat, light] = rgbToHsl(rgb);
    let l = Math.max(light, 1 - light);
    let next = hslToRgb(hue, sat, l);
    while (contrast(relLuminance(next)) < 4.5 && l < 0.97) {
      l = Math.min(0.97, l + 0.03);
      next = hslToRgb(hue, sat, l);
    }
    out = '#' + next.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  darkColorCache.set(norm, out);
  return out;
}

/** @param {number[]} rgb channels 0..1 */
function relLuminance(rgb) {
  const [r, g, b] = rgb.map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Luminance of a '#rgb' / '#rrggbb[aa]' / 'rgb(...)' CSS colour, or undefined. @param {string} css */
function cssColorLuminance(css) {
  const s = css.trim();
  let m = /^#([0-9a-f]{3,8})$/i.exec(s);
  if (m) {
    let hex = m[1];
    if (hex.length <= 4) hex = [...hex].map((ch) => ch + ch).join('');
    return relLuminance([0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255));
  }
  m = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(s);
  return m ? relLuminance([m[1], m[2], m[3]].map((v) => Number(v) / 255)) : undefined;
}

/** @param {number[]} rgb @returns {[number, number, number]} */
function rgbToHsl(rgb) {
  const [r, g, b] = rgb;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const hue = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [hue / 6, s, l];
}

/** @param {number} hue @param {number} s @param {number} l @returns {number[]} */
function hslToRgb(hue, s, l) {
  if (!s) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (/** @type {number} */ t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)];
}

/**
 * Underline / strikethrough. One text-decoration has one line style, so with a double underline the (always
 * single) strikethrough is drawn by an inner span instead (strikeInside, `.fv-strike`).
 * @param {FontStyle} f
 */
function textDecorationCss(f) {
  const lines = [];
  const double = f.underline === 'double' || f.underline === 'doubleAccounting';
  if (f.underline) lines.push('underline');
  if (f.strike && !double) lines.push('line-through');
  if (!lines.length) return '';
  return `text-decoration:${lines.join(' ')}${double ? ' double' : ''};`;
}

/** Strikethrough that must be drawn on an inner span (see textDecorationCss). @param {FontStyle} f */
function strikeInside(f) {
  return !!f.strike && (f.underline === 'double' || f.underline === 'doubleAccounting');
}

/**
 * @param {BorderEdge | undefined} e @param {boolean} hasFill
 * @returns {Edge | null}
 */
function edgeInfo(e, hasFill) {
  if (!e || !e.style) return null;
  const spec = BORDER_STYLES[e.style];
  if (!spec) return null;
  const g = /** @type {GridState} */ (G);
  const color = e.color ? (!hasFill && g.dark ? readableOnDark(e.color) : e.color) : hasFill ? '#000000' : 'var(--fv-auto-border)';
  const [w, kind] = spec;
  const rank = BORDER_RANK.indexOf(e.style);
  if (kind === 'hair') return { w, css: `1px solid color-mix(in srgb, ${color} 55%, transparent)`, rank };
  return { w, css: `${w}px ${kind} ${color}`, rank };
}

/**
 * SVG background for diagonal borders.
 * @param {BorderEdge & { up?: boolean, down?: boolean }} d @param {boolean} hasFill
 */
function diagonalImage(d, hasFill) {
  const g = /** @type {GridState} */ (G);
  const spec = BORDER_STYLES[d.style] || BORDER_STYLES.thin;
  const explicit = d.color && !hasFill && g.dark ? readableOnDark(d.color) : d.color;
  const color = (explicit || (hasFill || !g.dark ? '#000000' : '#C8C8C8')).replace(/[^#\w]/g, '');
  const dash = spec[1] === 'dashed' ? " stroke-dasharray='4 2'" : spec[1] === 'dotted' ? " stroke-dasharray='1 2'" : '';
  const stroke = `stroke='${color}' stroke-width='${spec[0]}'${dash} vector-effect='non-scaling-stroke'`;
  let lines = '';
  if (d.down) lines += `<line x1='0' y1='0' x2='100' y2='100' ${stroke}/>`;
  if (d.up) lines += `<line x1='0' y1='100' x2='100' y2='0' ${stroke}/>`;
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100' preserveAspectRatio='none'>${lines}</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

/** @param {{ type: string, fg?: string, bg?: string }} p */
function patternCss(p) {
  const fg = p.fg || '#000000';
  const bg = p.bg || '#FFFFFF';
  const gray = GRAY_PATTERNS[p.type];
  if (gray !== undefined) return { color: `color-mix(in srgb, ${fg} ${gray}%, ${bg})`, image: '' };
  const t = p.type.startsWith('dark') ? 2 : 1;
  const line = (/** @type {number} */ deg) => `repeating-linear-gradient(${deg}deg, ${fg} 0 ${t}px, transparent ${t}px 4px)`;
  /** @type {string[]} */
  let images = [];
  switch (p.type.replace(/^(dark|light)/, '')) {
    case 'Horizontal':
      images = [line(0)];
      break;
    case 'Vertical':
      images = [line(90)];
      break;
    case 'Down':
      images = [line(45)];
      break;
    case 'Up':
      images = [line(-45)];
      break;
    case 'Grid':
      images = [line(0), line(90)];
      break;
    case 'Trellis':
      images = [line(45), line(-45)];
      break;
    default:
      return { color: `color-mix(in srgb, ${fg} 50%, ${bg})`, image: '' };
  }
  return { color: bg, image: images.join(',') };
}

/** @param {{ angle: number, stops: { position: number, color: string }[] }} gr */
function gradientCss(gr) {
  const stops = gr.stops.map((s) => `${s.color} ${Math.round(clamp(s.position, 0, 1) * 1000) / 10}%`).join(', ');
  return `linear-gradient(${90 + (gr.angle || 0)}deg, ${stops})`;
}

/**
 * Cached, zoom-dependent rendering info of a style index.
 * @param {number} s
 * @returns {StyleInfo}
 */
function styleInfo(s) {
  const g = /** @type {GridState} */ (G);
  let info = g.styleCache.get(s);
  if (!info) {
    info = styleInfoRaw(g.styles[s], g.styleZoom / 100);
    g.styleCache.set(s, info);
  }
  return info;
}

/**
 * @param {CellStyle | undefined} st @param {number} z
 * @returns {StyleInfo}
 */
function styleInfoRaw(st, z) {
  const g = /** @type {GridState} */ (G);
  st = st || {};
  const def = g.meta.defaultFont || { name: 'Calibri', size: 11 };
  // Fill
  /** @type {string[]} */
  const images = [];
  let bgColor = '';
  const pat = st.pattern;
  if (st.gradient && st.gradient.stops && st.gradient.stops.length) images.push(gradientCss(st.gradient));
  else if (pat && pat.type && pat.type !== 'none' && pat.type !== 'solid') {
    const pc = patternCss(pat);
    bgColor = pc.color;
    if (pc.image) images.push(pc.image);
  } else if (st.fill) bgColor = st.fill;
  else if (pat && pat.type === 'solid' && pat.fg) bgColor = pat.fg;
  const hasFill = !!bgColor || images.length > 0;
  // Font
  const f = st.font || {};
  const family = f.name || def.name || 'Calibri';
  const fontPx = ptToPx(f.size || def.size || 11, z);
  let css = '';
  if (f.name && f.name !== def.name) css += `font-family:${fontFamilyCss(f.name)};`;
  if (f.size && f.size !== def.size) css += `font-size:${fontPx}px;`;
  if (f.bold) css += 'font-weight:700;';
  if (f.italic) css += 'font-style:italic;';
  css += textDecorationCss(f);
  const color = cellFontColor(f.color, hasFill);
  if (color) css += `color:${color};`;
  // Borders
  const b = st.border || {};
  const top = edgeInfo(b.top, hasFill);
  const right = edgeInfo(b.right, hasFill);
  const bottom = edgeInfo(b.bottom, hasFill);
  const left = edgeInfo(b.left, hasFill);
  const dg = b.diagonal;
  const diag = dg && dg.style && (dg.up || dg.down) ? diagonalImage(dg, hasFill) : '';
  if (diag) images.unshift(diag);
  if (bgColor) css += `background-color:${bgColor};`;
  if (images.length) css += `background-image:${images.join(',')};background-size:100% 100%;background-repeat:no-repeat;`;
  if (diag && images.length > 1) css = css.replace('background-size:100% 100%;background-repeat:no-repeat;', `background-size:100% 100%${',auto'.repeat(images.length - 1)};background-repeat:no-repeat${',repeat'.repeat(images.length - 1)};`);
  else if (!diag && images.length) css = css.replace('background-size:100% 100%;background-repeat:no-repeat;', '');
  // Alignment
  const a = st.align || {};
  /** @type {StyleInfo['h']} */
  let hAlign;
  switch (a.h) {
    case 'left':
    case 'center':
    case 'right':
    case 'justify':
    case 'distributed':
    case 'fill':
    case 'centerContinuous':
      hAlign = a.h;
      break;
    default:
      hAlign = undefined;
  }
  const wrap = !!a.wrap || hAlign === 'justify' || hAlign === 'distributed' || a.v === 'justify' || a.v === 'distributed';
  const vcls = a.v === 'top' ? ' fv-vt' : a.v === 'middle' || a.v === 'distributed' || a.v === 'justify' ? ' fv-vm' : '';
  let rot = a.rotation ?? 0;
  if (typeof rot === 'number') {
    if (rot === 255) rot = 'vertical';
    else if (rot > 90 && rot <= 180) rot = 90 - rot;
    else rot = clamp(Math.round(rot), -90, 90);
  } else if (rot !== 'vertical') rot = 0;
  return {
    css,
    vcls,
    hasFill,
    bg: bgColor,
    visual: hasFill || !!(top || right || bottom || left || diag),
    h: hAlign,
    wrap,
    shrink: !!a.shrink && !wrap,
    indent: Math.max(0, a.indent || 0) * 9 * z,
    rot,
    top,
    right,
    bottom,
    left,
    font: `${f.italic ? 'italic ' : ''}${f.bold ? 'bold ' : ''}${fontPx}px ${fontFamilyCss(family)}`,
    fontPx,
    family: fontFamilyCss(family),
    bold: !!f.bold,
    italic: !!f.italic,
    autoColor: !normColor(f.color) || normColor(f.color) === '#000000' || normColor(f.color) === g.defaultFontColor,
    fontColor: f.color,
    strikeIn: strikeInside(f),
    va: f.vertAlign,
  };
}

// ----- RENDER: number-format colours -----
// `[Red]0.00;[Blue]-0.00` etc.: Excel colours the displayed value by the format section it uses, and the
// format colour wins over the font colour. The host formats the text; the colour is derived here from the
// style's format code.

/** @typedef {{ op: string, value: number } | null} FmtCond */
/** @typedef {{ color: string | undefined, cond: FmtCond, text: boolean }} FmtSection */

/** Excel's eight named format colours. @type {Record<string, string>} */
const FMT_COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };

/** [Color1]..[Color56] = legacy indexed palette entries 8..63. */
const FMT_PALETTE = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 ' +
  '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF ' +
  '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 ' +
  '003366 339966 003300 333300 993300 993366 333399 333333'
)
  .split(' ')
  .map((hex) => '#' + hex);

/** @type {Map<string, FmtSection[] | null>} */
const fmtCache = new Map();

/**
 * Sections of a format code that carry a colour (null when none does).
 * @param {string} fmt
 * @returns {FmtSection[] | null}
 */
function parseFmtColors(fmt) {
  let parsed = fmtCache.get(fmt);
  if (parsed !== undefined) return parsed;
  /** @type {FmtSection[]} */
  const sections = [];
  /** @type {string | undefined} */
  let color;
  /** @type {FmtCond} */
  let cond = null;
  let text = false;
  const flush = () => {
    sections.push({ color, cond, text });
    color = undefined;
    cond = null;
    text = false;
  };
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') {
      const end = fmt.indexOf('"', i + 1);
      i = end < 0 ? fmt.length : end;
    } else if (ch === '\\') i++;
    else if (ch === '@') text = true;
    else if (ch === ';') flush();
    else if (ch === '[') {
      const end = fmt.indexOf(']', i + 1);
      const token = fmt.slice(i + 1, end < 0 ? fmt.length : end).trim();
      i = end < 0 ? fmt.length : end;
      const lower = token.toLowerCase();
      const indexed = /^color\s*(\d{1,2})$/.exec(lower);
      const c = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/.exec(token);
      if (FMT_COLORS[lower]) color = FMT_COLORS[lower];
      else if (indexed && Number(indexed[1]) >= 1 && Number(indexed[1]) <= 56) color = FMT_PALETTE[Number(indexed[1]) - 1];
      else if (c) cond = { op: c[1], value: Number(c[2]) };
    }
  }
  flush();
  parsed = sections.some((s) => s.color) ? sections : null;
  if (fmtCache.size > 2000) fmtCache.clear();
  fmtCache.set(fmt, parsed);
  return parsed;
}

/** @param {FmtCond} cond @param {number} v */
function fmtCondMatches(cond, v) {
  if (!cond) return false;
  switch (cond.op) {
    case '<':
      return v < cond.value;
    case '<=':
      return v <= cond.value;
    case '>':
      return v > cond.value;
    case '>=':
      return v >= cond.value;
    case '=':
      return v === cond.value;
    default:
      return v !== cond.value;
  }
}

/**
 * Colour of a value under a format code, or undefined.
 * @param {string | undefined} fmt @param {CellData} cell
 */
function numFmtColor(fmt, cell) {
  if (!fmt || !fmt.includes('[')) return undefined;
  const sections = parseFmtColors(fmt);
  if (!sections) return undefined;
  const v = cell.v;
  if (typeof v !== 'number') {
    if (cell.t !== 's') return undefined;
    // Text: the 4th section, or a lone section containing @.
    const textSection = sections[3] ?? (sections.length === 1 && sections[0].text ? sections[0] : undefined);
    return textSection?.color;
  }
  /** @type {FmtSection | undefined} */
  let section;
  if (sections[0].cond || sections[1]?.cond) {
    if (fmtCondMatches(sections[0].cond, v)) section = sections[0];
    else if (sections[1] && (sections[1].cond ? fmtCondMatches(sections[1].cond, v) : true)) section = sections[1];
    else section = sections[2] ?? sections[1];
  } else {
    const numeric = sections.filter((s, i) => i < 3);
    if (numeric.length === 1) section = numeric[0];
    else if (numeric.length === 2) section = v < 0 ? numeric[1] : numeric[0];
    else section = v > 0 ? numeric[0] : v < 0 ? numeric[1] : numeric[2];
  }
  return section?.color;
}

/** @type {CanvasRenderingContext2D | null} */
let measureCtx = null;
/** The font last assigned to measureCtx. */
let measureFont = '';
/** @type {Map<string, number>} */
const measureCache = new Map();

/** Text width in px for a canvas font (cached; never touches the DOM layout). @param {string} font @param {string} text */
function measureText(font, text) {
  const key = font + '\u0000' + text;
  let w = measureCache.get(key);
  if (w !== undefined) return w;
  if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
  if (!measureCtx) return text.length * 7;
  if (measureCache.size > 20000) measureCache.clear();
  // Assigning ctx.font parses the font string: only when it changes.
  if (measureFont !== font) {
    measureCtx.font = font;
    measureFont = font;
  }
  w = measureCtx.measureText(text).width;
  measureCache.set(key, w);
  return w;
}

// ----- RENDER: cell values -----

/** @param {CellData | null | undefined} cell */
function isEmptyCell(cell) {
  if (!cell || cell.t === 'z') return true;
  return (cell.v === null || cell.v === undefined || cell.v === '') && !cell.w;
}

/** Text the grid shows for a cell (number format already applied by the host). @param {CellData | null | undefined} cell */
function displayText(cell) {
  if (!cell || cell.t === 'z') return '';
  if (cell.w !== undefined && cell.w !== null) return cell.w;
  const v = cell.v;
  if (v === null || v === undefined) return '';
  switch (cell.t) {
    case 'n':
      return typeof v === 'number' ? formatGeneral(v) : String(v);
    case 'b':
      return v ? 'TRUE' : 'FALSE';
    case 'd':
      return typeof v === 'number' ? serialToDateString(v) : String(v);
    default:
      return String(v);
  }
}

/** Text whose first letter (strong directional character) is from a right-to-left script. */
const RTL_TEXT_RE = /^\P{L}*(?=\p{L})[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u;

/** Excel "General" alignment. @param {CellData | null | undefined} cell */
function generalAlign(cell) {
  if (!cell) return 'left';
  if (cell.t === 'n' || cell.t === 'd') return 'right';
  if (cell.t === 'b' || cell.t === 'e') return 'center';
  // Excel's "Context" reading order: text whose first letter is from a right-to-left script aligns right.
  return RTL_TEXT_RE.test(displayText(cell)) ? 'right' : 'left';
}

/** Excel "General" number format (up to ~11 significant characters). @param {number} n */
function formatGeneral(n) {
  if (!Number.isFinite(n)) return '#NUM!';
  if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
  const a = Math.abs(n);
  if (a >= 1e11 || a < 1e-9) {
    return n
      .toExponential(5)
      .replace(/\.?0+e/, 'e')
      .replace(/e([+-])(\d)$/, 'e$10$2')
      .toUpperCase();
  }
  let s = n.toPrecision(10);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

/**
 * Excel serial → locale date/time text (formula bar / fallback display).
 * @param {number} serial
 */
function serialToDateString(serial) {
  const g = /** @type {GridState} */ (G);
  let days = serial + (g.meta.date1904 ? 1462 : 0);
  if (!g.meta.date1904 && days < 61) days += 1; // Excel's fictitious 1900-02-29
  const ms = Math.round((days - 25569) * 86400000);
  const date = new Date(ms);
  if (Number.isNaN(date.getTime())) return String(serial);
  const hasTime = Math.abs(serial - Math.floor(serial)) > 1e-9;
  const opts = { timeZone: 'UTC' };
  if (serial < 1 && serial >= 0) return date.toLocaleTimeString(undefined, opts);
  return hasTime ? `${date.toLocaleDateString(undefined, opts)} ${date.toLocaleTimeString(undefined, opts)}` : date.toLocaleDateString(undefined, opts);
}

/** Value shown in the formula bar: `=formula` (`{=formula}`), or the raw value. @param {CellData | null | undefined} cell */
function rawValue(cell) {
  if (!cell || cell.t === 'z') return '';
  // Legacy (Ctrl+Shift+Enter) array formulas in braces, like Excel's formula bar; dynamic arrays have none.
  if (cell.f) return cell.array ? `{=${cell.f}}` : '=' + cell.f;
  // A delimited file's number shows its field text (007, 1e5, 50%), the same as the cell and the copy.
  if (G?.kind === 'csv' && cell.t === 'n' && typeof cell.w === 'string') return cell.w;
  const v = cell.v;
  if (v === null || v === undefined) return cell.w ?? '';
  switch (cell.t) {
    case 'n':
      return typeof v === 'number' ? formulaBarNumber(v) : String(v);
    case 'd':
      return typeof v === 'number' ? serialToDateString(v) : String(v);
    case 'b':
      return v ? 'TRUE' : 'FALSE';
    default:
      return String(v);
  }
}

/**
 * A number as Excel's formula bar shows it (measured with UI Automation): 15 significant digits, written out in full
 * while that takes at most 21 characters (100000000000000000000, 0.000000000001), else scientific (1E+21, 1E-20,
 * 1.23456789012345E-12).
 * @param {number} n
 */
function formulaBarNumber(n) {
  if (!Number.isFinite(n)) return String(n);
  if (n === 0) return '0';
  const [mantissa, e] = Number(n.toPrecision(15)).toExponential().split('e');
  const exp = Number(e);
  const digits = mantissa.replace(/^-/, '').replace('.', '');
  const sign = n < 0 ? '-' : '';
  let full;
  if (exp < 0) full = `0.${'0'.repeat(-exp - 1)}${digits}`;
  else if (digits.length <= exp + 1) full = digits + '0'.repeat(exp + 1 - digits.length);
  else full = `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`;
  if (full.length <= 21) return sign + full;
  return `${mantissa}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/** @param {SheetMeta} sm @param {number} r */
function rowStyleIndex(sm, r) {
  const info = sm.rows[r];
  return info && info.s ? info.s : 0;
}

/** Binary search of a cell in a row's ascending cell list. @param {CellData[]} cells @param {number} c */
function findCell(cells, c) {
  let lo = 0;
  let hi = cells.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const cc = cells[mid].c;
    if (cc === c) return cells[mid];
    if (cc < c) lo = mid + 1;
    else hi = mid - 1;
  }
  return null;
}

/** First index with cells[i].c >= c. @param {CellData[]} cells @param {number} c */
function lowerBound(cells, c) {
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (cells[mid].c < c) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** 0 → A, 25 → Z, 26 → AA. @param {number} c */
function colName(c) {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

/** 'A' → 0, 'AA' → 26; -1 when invalid. @param {string} letters */
function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0) - 64;
    if (code < 1 || code > 26) return -1;
    n = n * 26 + code;
  }
  return n - 1;
}

// ----- SELECTION -----

/** @param {SheetView} view @param {Range} rg */
function isFullRows(view, rg) {
  return rg.c0 === 0 && rg.c1 >= view.maxC - 1;
}

/** @param {SheetView} view @param {Range} rg */
function isFullCols(view, rg) {
  return rg.r0 === 0 && rg.r1 >= view.maxR - 1;
}

/** @param {Cell} a @param {Cell} b @returns {Range} */
function rangeOf(a, b) {
  return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
}

/**
 * Grows a range until it contains every merge it touches.
 * @param {SheetView} view @param {Range} rg
 */
function snapRange(view, rg) {
  const merges = view.meta.merges;
  if (!merges.length) return rg;
  let { r0, c0, r1, c1 } = rg;
  for (let guard = 0, changed = true; changed && guard < 64; guard++) {
    changed = false;
    for (const mi of mergesInRows(view, r0, r1)) {
      const m = merges[mi];
      if (m.r1 < r0 || m.r0 > r1 || m.c1 < c0 || m.c0 > c1) continue;
      if (m.r0 < r0) (r0 = m.r0), (changed = true);
      if (m.c0 < c0) (c0 = m.c0), (changed = true);
      if (m.r1 > r1) (r1 = m.r1), (changed = true);
      if (m.c1 > c1) (c1 = m.c1), (changed = true);
    }
  }
  return { r0, c0, r1, c1 };
}

/** @param {SheetView} view */
function initialSelection(view) {
  const L = view.L;
  const r = L ? visibleRow(L, 0) : 0;
  const c = L ? visibleCol(L, 0) : 0;
  const m = masterOf(view, r, c);
  return { ranges: [cellArea(view, m.r, m.c)], active: m, anchor: m, cursor: m };
}

/**
 * Validates a persisted selection.
 * @param {SheetView} view @param {any} saved
 * @returns {Selection | null}
 */
function restoreSelection(view, saved) {
  if (!saved || !Array.isArray(saved.ranges) || !saved.ranges.length || !saved.active) return null;
  const okInt = (/** @type {any} */ n, /** @type {number} */ max) => Number.isInteger(n) && n >= 0 && n < max;
  /** @type {Range[]} */
  const ranges = [];
  for (const rg of saved.ranges.slice(0, 256)) {
    if (!rg || !okInt(rg.r0, view.maxR) || !okInt(rg.r1, view.maxR) || !okInt(rg.c0, view.maxC) || !okInt(rg.c1, view.maxC)) return null;
    if (rg.r0 > rg.r1 || rg.c0 > rg.c1) return null;
    ranges.push({ r0: rg.r0, c0: rg.c0, r1: rg.r1, c1: rg.c1 });
  }
  const cellOk = (/** @type {any} */ p) => p && okInt(p.r, view.maxR) && okInt(p.c, view.maxC);
  if (!cellOk(saved.active)) return null;
  const active = { r: saved.active.r, c: saved.active.c };
  const anchor = cellOk(saved.anchor) ? { r: saved.anchor.r, c: saved.anchor.c } : active;
  const cursor = cellOk(saved.cursor) ? { r: saved.cursor.r, c: saved.cursor.c } : anchor;
  return { ranges, active, anchor, cursor };
}

/**
 * Replaces the selection.
 * @param {SheetView} view @param {Range[]} ranges @param {Cell} active
 * @param {{ anchor?: Cell, cursor?: Cell, reveal?: Cell | null }} [opts]
 */
function setSelection(view, ranges, active, opts = {}) {
  view.sel = { ranges, active, anchor: opts.anchor ?? active, cursor: opts.cursor ?? opts.anchor ?? active };
  growLayout(view, Math.max(active.r, view.sel.cursor.r), Math.max(active.c, view.sel.cursor.c));
  selectionChanged(view, opts.reveal === undefined ? active : opts.reveal);
}

/**
 * The range from the anchor to `cursor` (grown over merges). While the last range is whole rows (columns), as after
 * Shift+Space or a header click, keyboard extension keeps whole rows (columns) like Excel.
 * @param {SheetView} view @param {Cell} cursor
 */
function extendedRange(view, cursor) {
  const sel = view.sel;
  const last = sel.ranges[sel.ranges.length - 1];
  const rg = rangeOf(sel.anchor, cursor);
  if (last && isFullRows(view, last)) (rg.c0 = 0), (rg.c1 = view.maxC - 1);
  if (last && isFullCols(view, last)) (rg.r0 = 0), (rg.r1 = view.maxR - 1);
  return snapRange(view, rg);
}

/**
 * @param {SheetView} view @param {Cell | null} reveal
 */
function selectionChanged(view, reveal) {
  if (!G) return;
  G.selVersion++;
  computeHeaderSel(view);
  if (reveal) scrollCellIntoView(view, reveal.r, reveal.c);
  scheduleRender();
  updateFormulaBar();
  scheduleStats();
  saveGridState();
  announceActive();
}

/** Single cell (snapped to its merge). @param {SheetView} view @param {number} r @param {number} c @param {boolean} add */
function selectCell(view, r, c, add = false) {
  const m = masterOf(view, r, c);
  const area = cellArea(view, m.r, m.c);
  const ranges = add ? [...view.sel.ranges, area] : [area];
  setSelection(view, ranges, m);
}

/**
 * Ctrl+click (Ctrl+drag) on selected cells deselects them, like Excel 365: every range of `base` is split around the
 * hole (measured in Excel: the part below it, then right of it, left of it and above it, in place of the range), and
 * the active cell becomes the top-left cell of the last piece of the last range that was split. When nothing would
 * stay selected, the selection is `base` (Excel keeps a single cell that is Ctrl+clicked).
 * @param {SheetView} view @param {Selection} base  the selection before the click @param {Range} hole
 */
function deselectCells(view, base, hole) {
  /** @type {Range[]} */
  const ranges = [];
  /** @type {Range | undefined} */
  let lastPiece;
  let split = false;
  for (const rg of base.ranges) {
    if (hole.r1 < rg.r0 || hole.r0 > rg.r1 || hole.c1 < rg.c0 || hole.c0 > rg.c1) {
      ranges.push(rg);
      continue;
    }
    split = true;
    const r0 = Math.max(rg.r0, hole.r0);
    const r1 = Math.min(rg.r1, hole.r1);
    /** @type {Range[]} */
    const pieces = [];
    if (hole.r1 < rg.r1) pieces.push({ r0: hole.r1 + 1, c0: rg.c0, r1: rg.r1, c1: rg.c1 });
    if (hole.c1 < rg.c1) pieces.push({ r0, c0: hole.c1 + 1, r1, c1: rg.c1 });
    if (hole.c0 > rg.c0) pieces.push({ r0, c0: rg.c0, r1, c1: hole.c0 - 1 });
    if (hole.r0 > rg.r0) pieces.push({ r0: rg.r0, c0: rg.c0, r1: hole.r0 - 1, c1: rg.c1 });
    if (pieces.length) lastPiece = pieces[pieces.length - 1];
    ranges.push(...pieces);
  }
  if (!split || !ranges.length) {
    if (view.sel !== base) {
      view.sel = base;
      selectionChanged(view, null);
    }
    return;
  }
  // A range removed whole leaves no piece: the active cell stays when still selected, else the last range's first cell.
  const last = ranges[ranges.length - 1];
  const p = lastPiece ? { r: lastPiece.r0, c: lastPiece.c0 } : selectionContainsIn(ranges, base.active) ? base.active : { r: last.r0, c: last.c0 };
  setSelection(view, ranges, masterOf(view, p.r, p.c), { reveal: null });
}

/** Extends the last range from the anchor to (r, c). @param {SheetView} view @param {number} r @param {number} c */
function extendSelection(view, r, c, reveal = true) {
  const sel = view.sel;
  const rg = snapRange(view, rangeOf(sel.anchor, { r, c }));
  const ranges = sel.ranges.slice(0, -1).concat([rg]);
  setSelection(view, ranges, sel.active, { anchor: sel.anchor, cursor: { r, c }, reveal: reveal ? { r, c } : null });
}

/**
 * Whole rows a..b (or columns), from a header click / drag. The active cell is the first cell of row (column)
 * `activeIdx` shown in the window, like Excel, so selecting a header never scrolls and the next arrow key starts
 * where the user is looking. Merged cells do not widen the range: Excel's header click on column D selects D:D even
 * where B35:D35 is merged (the name box and Ctrl/Shift+Space do grow over merges, see snapRange).
 * @param {SheetView} view @param {'rows' | 'cols'} mode @param {number} a @param {number} b
 * @param {boolean | 'last'} add  true = new range, false = replace all, 'last' = replace the last range (Shift+click)
 * @param {number} [activeIdx]
 */
function selectLines(view, mode, a, b, add, activeIdx = a) {
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  /** @type {Range} */
  const rg = mode === 'rows' ? { r0: lo, r1: hi, c0: 0, c1: view.maxC - 1 } : { r0: 0, r1: view.maxR - 1, c0: lo, c1: hi };
  const shown = firstShownCell(view);
  const active = mode === 'rows' ? masterOf(view, activeIdx, shown.c) : masterOf(view, shown.r, activeIdx);
  const keep = add === 'last' ? view.sel.ranges.slice(0, -1) : add ? view.sel.ranges : [];
  const anchor = mode === 'rows' ? { r: activeIdx, c: active.c } : { r: active.r, c: activeIdx };
  // Shift+click keeps the active cell of the range being extended.
  const keepActive = add === 'last' && selectionContainsIn([rg], view.sel.active);
  setSelection(view, [...keep, rg], keepActive ? view.sel.active : active, { anchor, cursor: mode === 'rows' ? { r: b, c: active.c } : { r: active.r, c: b }, reveal: null });
}

/** @param {Range[]} ranges @param {Cell} p */
function selectionContainsIn(ranges, p) {
  return ranges.some((rg) => p.r >= rg.r0 && p.r <= rg.r1 && p.c >= rg.c0 && p.c <= rg.c1);
}

/**
 * Top-left cell shown in the window: the first row / column of the frozen panes, else the first fully visible
 * scrolled one.
 * @param {SheetView} view
 * @returns {Cell}
 */
function firstShownCell(view) {
  const g = /** @type {GridState} */ (G);
  const L = /** @type {Layout} */ (view.L);
  let r = 0;
  let c = 0;
  if (L.fr === 0 && view.index === g.sheet) {
    const y = virtualScrollTop();
    r = findIndex(L.rowTop, L.nRows, y);
    if (L.rowTop[r] < y - 0.5) r = stepRow(L, r, 1);
  }
  if (L.fc === 0 && view.index === g.sheet) {
    const x = g.dom.scroller.scrollLeft;
    c = findIndex(L.colLeft, L.nCols, x);
    if (L.colLeft[c] < x - 0.5) c = stepCol(L, c, 1);
  }
  return { r: visibleRow(L, r), c: visibleCol(L, c) };
}

/**
 * Shift+Space / Ctrl+Space: every row (column) the selection touches, keeping the active cell (Excel's EntireRow /
 * EntireColumn of all ranges).
 * @param {SheetView} view @param {'rows' | 'cols'} mode
 */
function selectEntire(view, mode) {
  const sel = view.sel;
  const ranges = sel.ranges.map((rg) => snapRange(view, mode === 'rows' ? { r0: rg.r0, r1: rg.r1, c0: 0, c1: view.maxC - 1 } : { r0: 0, r1: view.maxR - 1, c0: rg.c0, c1: rg.c1 }));
  setSelection(view, ranges, sel.active, { anchor: sel.anchor, cursor: sel.cursor, reveal: null });
}

/** The whole sheet; the active cell stays where it is (Excel). @param {SheetView} view */
function selectAll(view) {
  const active = view.sel.active;
  setSelection(view, [{ r0: 0, c0: 0, r1: view.maxR - 1, c1: view.maxC - 1 }], active, { anchor: active, cursor: active, reveal: null });
}

/** @param {SheetView} view @param {number} r @param {number} c */
function selectionContains(view, r, c) {
  return view.sel.ranges.some((rg) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1);
}

/** Address text for the name box. @param {SheetView} view */
function selectionLabel(view) {
  const sel = view.sel;
  const a = sel.active;
  const activeAddr = colName(a.c) + (a.r + 1);
  if (sel.ranges.length !== 1) return activeAddr;
  const rg = sel.ranges[0];
  const fullRows = isFullRows(view, rg);
  const fullCols = isFullCols(view, rg);
  if (fullRows && fullCols) return activeAddr;
  if (fullRows) return `${rg.r0 + 1}:${rg.r1 + 1}`;
  if (fullCols) return `${colName(rg.c0)}:${colName(rg.c1)}`;
  const area = cellArea(view, a.r, a.c);
  if (rg.r0 === area.r0 && rg.r1 === area.r1 && rg.c0 === area.c0 && rg.c1 === area.c1) return activeAddr;
  return `${colName(rg.c0)}${rg.r0 + 1}:${colName(rg.c1)}${rg.r1 + 1}`;
}

/**
 * Scrolls so that the cell (its merge) is visible below/right of the frozen panes.
 * @param {SheetView} view @param {number} r @param {number} c
 */
function scrollCellIntoView(view, r, c) {
  const g = /** @type {GridState} */ (G);
  const L = view.L;
  if (!L) return;
  const area = cellArea(view, clamp(r, 0, L.nRows - 1), clamp(c, 0, L.nCols - 1));
  const vy = virtualScrollTop();
  const sl = g.dom.scroller.scrollLeft;
  let top;
  let left;
  if (area.r1 >= L.fr) {
    const r0 = Math.max(area.r0, L.fr);
    const y0 = L.rowTop[r0] - L.rowTop[L.fr];
    const y1 = L.rowTop[Math.min(area.r1, L.nRows - 1) + 1] - L.rowTop[L.fr];
    const visH = Math.max(1, g.viewH - L.A);
    if (y0 < vy) top = y0;
    else if (y1 > vy + visH) top = Math.min(y0, y1 - visH);
  }
  if (area.c1 >= L.fc) {
    const c0 = Math.max(area.c0, L.fc);
    const x0 = L.colLeft[c0] - L.colLeft[L.fc];
    const x1 = L.colLeft[Math.min(area.c1, L.nCols - 1) + 1] - L.colLeft[L.fc];
    const visW = Math.max(1, g.viewW - L.B);
    if (x0 < sl) left = x0;
    else if (x1 > sl + visW) left = Math.min(x0, x1 - visW);
  }
  if (top !== undefined || left !== undefined) setVirtualScroll(view, top, left);
}

/** Right-to-left sheet: the whole scroller is mirrored (see applyLayout). */
function isRtl() {
  return !!(G && G.sheet >= 0 && activeView().L?.rtl);
}

/**
 * Scroller-relative layout x of a viewport x (mirrored for right-to-left sheets).
 * @param {DOMRect} rect scroller bounding rect @param {number} clientX
 */
function scrollerX(rect, clientX) {
  return isRtl() ? rect.right - clientX : clientX - rect.left;
}

/**
 * Viewport x of a scroller-relative layout x (inverse of scrollerX).
 * @param {DOMRect} rect @param {number} x
 */
function clientXOf(rect, x) {
  return isRtl() ? rect.right - x : rect.left + x;
}

/**
 * Converts a point in scroller coordinates to a grid position.
 * @param {number} vx @param {number} vy
 */
function hitTest(vx, vy) {
  const g = /** @type {GridState} */ (G);
  const view = activeView();
  const L = /** @type {Layout} */ (view.L);
  const cx = vx < L.B ? vx : vx + g.dom.scroller.scrollLeft;
  const cy = vy < L.A ? vy : vy + virtualScrollTop();
  const c = findIndex(L.colLeft, L.nCols, cx - L.hdrW);
  const r = findIndex(L.rowTop, L.nRows, cy - L.hdrH);
  const inColHdr = vy < L.hdrH;
  const inRowHdr = vx < L.hdrW;
  /** @type {'corner' | 'colhdr' | 'rowhdr' | 'cell'} */
  const area = inColHdr && inRowHdr ? 'corner' : inColHdr ? 'colhdr' : inRowHdr ? 'rowhdr' : 'cell';
  return { area, r: visibleRow(L, r), c: visibleCol(L, c) };
}

/** @param {MouseEvent} e */
function onGridMouseDown(e) {
  if (!G) return;
  const sc = G.dom.scroller;
  const rect = sc.getBoundingClientRect();
  const vx = scrollerX(rect, e.clientX);
  const vy = e.clientY - rect.top;
  if (vx >= sc.clientWidth || vy >= sc.clientHeight || vx < 0 || vy < 0) return; // scrollbars
  hoverReset();
  closeMenu();
  G.keyNav = false;
  const view = activeView();
  const hit = hitTest(vx, vy);
  if (e.button === 2) {
    // Right click keeps a selection that contains the cell (context menu acts on it), else selects the cell.
    if (hit.area === 'cell' && !selectionContains(view, hit.r, hit.c)) selectCell(view, hit.r, hit.c);
    sc.focus({ preventScroll: true });
    return;
  }
  if (e.button !== 0) return;
  e.preventDefault();
  sc.focus({ preventScroll: true });
  const additive = e.ctrlKey || e.metaKey;
  switch (hit.area) {
    case 'corner':
      selectAll(view);
      return;
    // Shift+click extends the last range only (earlier Ctrl+click ranges stay, like Excel).
    case 'colhdr':
      if (e.shiftKey) selectLines(view, 'cols', view.sel.anchor.c, hit.c, 'last', view.sel.anchor.c);
      else selectLines(view, 'cols', hit.c, hit.c, additive);
      startDrag('cols', e, e.shiftKey ? view.sel.anchor.c : hit.c);
      return;
    case 'rowhdr':
      if (e.shiftKey) selectLines(view, 'rows', view.sel.anchor.r, hit.r, 'last', view.sel.anchor.r);
      else selectLines(view, 'rows', hit.r, hit.r, additive);
      startDrag('rows', e, e.shiftKey ? view.sel.anchor.r : hit.r);
      return;
    default: {
      const m = masterOf(view, hit.r, hit.c);
      if (additive && !e.shiftKey) {
        const cell = cellAt(view, m.r, m.c);
        if (cell && cell.link) {
          // The second press of a Ctrl+double-click does not follow the link again.
          if (e.detail < 2) openLink(cell.link);
          return;
        }
        if (selectionContains(view, hit.r, hit.c)) {
          // Ctrl+click on a selected cell deselects it; dragging on deselects the cells passed over (Excel 365).
          const base = view.sel;
          startDrag('deselect', e, 0, { from: m, base });
          deselectCells(view, base, cellArea(view, m.r, m.c));
          return;
        }
      }
      if (e.shiftKey) extendSelection(view, hit.r, hit.c, false);
      else selectCell(view, hit.r, hit.c, additive);
      startDrag('cells', e, 0);
    }
  }
}

/**
 * @param {'cells' | 'rows' | 'cols' | 'deselect'} mode @param {MouseEvent} e @param {number} anchor
 * @param {{ from: Cell, base: Selection }} [deselect]  mode 'deselect': the Ctrl+clicked cell and the selection before
 */
function startDrag(mode, e, anchor, deselect) {
  if (!G) return;
  // `last` = header index already applied (the mousedown handler selected up to the clicked header).
  const view = activeView();
  const L = /** @type {Layout} */ (view.L);
  const last = mode === 'rows' ? view.sel.cursor.r : mode === 'cols' ? view.sel.cursor.c : -1;
  const from = deselect ? deselect.from : view.sel.anchor;
  const frozenR = mode === 'cols' || (mode === 'rows' ? anchor : from.r) < L.fr;
  const frozenC = mode === 'rows' || (mode === 'cols' ? anchor : from.c) < L.fc;
  G.drag = { mode, x: e.clientX, y: e.clientY, raf: 0, anchor, last, frozenR, frozenC, deselect: deselect ? { ...deselect, to: deselect.from } : undefined };
}

/**
 * Upper-left limits of the pointer while dragging: a drag that started in the scrolling area stops at the frozen
 * panes' edge while the sheet can still scroll back (autoScrollStep scrolls instead), like Excel.
 * @param {NonNullable<GridState['drag']>} d @param {Layout} L
 */
function dragMin(d, L) {
  const sc = /** @type {GridState} */ (G).dom.scroller;
  return { x: !d.frozenC && sc.scrollLeft > 0 ? L.B : L.hdrW, y: !d.frozenR && sc.scrollTop > 0 ? L.A : L.hdrH };
}

/** @param {MouseEvent} e */
function onWindowMouseMove(e) {
  if (!G || !G.drag) return;
  if (!(e.buttons & 1)) {
    onWindowMouseUp();
    return;
  }
  G.drag.x = e.clientX;
  G.drag.y = e.clientY;
  dragUpdate();
  if (!G.drag.raf) G.drag.raf = requestAnimationFrame(autoScrollStep);
}

function onWindowMouseUp() {
  if (!G || !G.drag) return;
  if (G.drag.raf) cancelAnimationFrame(G.drag.raf);
  G.drag = null;
}

/** Updates the dragged selection from the last pointer position (clamped into the grid). */
function dragUpdate() {
  const g = /** @type {GridState} */ (G);
  const d = g.drag;
  if (!d) return;
  const sc = g.dom.scroller;
  const view = activeView();
  const L = /** @type {Layout} */ (view.L);
  const rect = sc.getBoundingClientRect();
  const min = dragMin(d, L);
  const vx = clamp(scrollerX(rect, d.x), min.x, Math.max(min.x, sc.clientWidth - 1));
  const vy = clamp(d.y - rect.top, min.y, Math.max(min.y, sc.clientHeight - 1));
  const hit = hitTest(vx, vy);
  if (d.mode === 'rows' || d.mode === 'cols') {
    // Whole rows / columns from the drag anchor to the hovered header (updated only when it changes); merged cells do
    // not widen them (selectLines).
    const rows = d.mode === 'rows';
    const at = rows ? hit.r : hit.c;
    if (at === d.last) return;
    d.last = at;
    const lo = Math.min(d.anchor, at);
    const hi = Math.max(d.anchor, at);
    /** @type {Range} */
    const rg = rows ? { r0: lo, r1: hi, c0: 0, c1: view.maxC - 1 } : { r0: 0, r1: view.maxR - 1, c0: lo, c1: hi };
    const sel = view.sel;
    view.sel = { ranges: [...sel.ranges.slice(0, -1), rg], active: sel.active, anchor: sel.anchor, cursor: rows ? { r: at, c: 0 } : { r: 0, c: at } };
    selectionChanged(view, null);
  } else if (d.deselect) {
    // Ctrl+drag from a selected cell: the rectangle from it to the hovered cell is taken out of the first selection.
    const ds = d.deselect;
    if (hit.r === ds.to.r && hit.c === ds.to.c) return;
    ds.to = { r: hit.r, c: hit.c };
    deselectCells(view, ds.base, snapRange(view, rangeOf(ds.from, hit)));
  } else if (hit.r !== view.sel.cursor.r || hit.c !== view.sel.cursor.c) {
    extendSelection(view, hit.r, hit.c, false);
  }
}

/** Scrolls while the pointer is dragged beyond the grid edges. */
function autoScrollStep() {
  const g = G;
  if (!g || !g.drag) return;
  const d = g.drag;
  d.raf = 0;
  const sc = g.dom.scroller;
  const L = /** @type {Layout} */ (activeView().L);
  const rect = sc.getBoundingClientRect();
  const vx = scrollerX(rect, d.x);
  const vy = d.y - rect.top;
  const min = dragMin(d, L);
  let dx = 0;
  let dy = 0;
  if (d.mode !== 'rows') {
    if (vx > sc.clientWidth) dx = vx - sc.clientWidth;
    else if (vx < min.x && sc.scrollLeft > 0) dx = vx - min.x;
  }
  if (d.mode !== 'cols') {
    if (vy > sc.clientHeight) dy = vy - sc.clientHeight;
    else if (vy < min.y && sc.scrollTop > 0) dy = vy - min.y;
  }
  if (!dx && !dy) return;
  const speed = (/** @type {number} */ v) => Math.sign(v) * Math.min(80, 4 + Math.abs(v) * 0.6);
  if (dx) sc.scrollLeft += speed(dx);
  if (dy) sc.scrollTop += speed(dy);
  dragUpdate();
  d.raf = requestAnimationFrame(autoScrollStep);
}

/** Tooltip for comments and hyperlinks. @param {MouseEvent} e */
function onGridHover(e) {
  if (!G || G.drag || e.buttons) return;
  const sc = G.dom.scroller;
  const rect = sc.getBoundingClientRect();
  const vx = scrollerX(rect, e.clientX);
  const vy = e.clientY - rect.top;
  const L = /** @type {Layout} */ (activeView().L);
  if (vx < L.hdrW || vy < L.hdrH || vx >= sc.clientWidth || vy >= sc.clientHeight) {
    hoverReset();
    return;
  }
  const view = activeView();
  const hit = hitTest(vx, vy);
  const m = masterOf(view, hit.r, hit.c);
  const key = `${view.index}:${m.r}:${m.c}`;
  if (key === G.hover.key) return;
  hoverReset();
  G.hover.key = key;
  const cell = cellAt(view, m.r, m.c);
  if (!cell || (!cell.note && !cell.link)) return;
  G.hover.timer = setTimeout(() => showCellTooltip(view, m.r, m.c, cell), TOOLTIP_DELAY_MS);
}

function hoverReset() {
  if (!G) return;
  if (G.hover.timer !== undefined) clearTimeout(G.hover.timer);
  G.hover.timer = undefined;
  G.hover.key = '';
  hideTooltip();
}

/** @param {SheetView} view @param {number} r @param {number} c @param {CellData} cell */
function showCellTooltip(view, r, c, cell) {
  if (!G || G.sheet !== view.index) return;
  const L = /** @type {Layout} */ (view.L);
  const tip = G.dom.tooltip;
  /** @type {HTMLElement[]} */
  const parts = [];
  if (cell.note) parts.push(h('div', { class: 'fv-tip-note' }, cell.note));
  if (cell.link) {
    const mod = IS_MAC ? 'Cmd' : 'Ctrl';
    parts.push(
      h(
        'div',
        { class: 'fv-tip-link' },
        h('span', { class: 'fv-tip-href' }, linkTarget(cell.link)),
        h('span', { class: 'fv-tip-hint' }, `${mod}+Click or ${mod}+Enter to follow link`),
      ),
    );
  }
  tip.replaceChildren(...parts);
  tip.hidden = false;
  const area = cellArea(view, r, c);
  const sc = G.dom.scroller;
  const rect = sc.getBoundingClientRect();
  const x0 = L.hdrW + L.colLeft[area.c0] - (area.c0 >= L.fc ? sc.scrollLeft : 0);
  const x1 = L.hdrW + L.colLeft[Math.min(area.c1, L.nCols - 1) + 1] - (area.c0 >= L.fc ? sc.scrollLeft : 0);
  const y0 = L.hdrH + L.rowTop[area.r0] - (area.r0 >= L.fr ? virtualScrollTop() : 0);
  const tw = tip.offsetWidth;
  const th = tip.offsetHeight;
  // Visual edges of the cell (the scroller is mirrored for right-to-left sheets).
  const a = clientXOf(rect, x0);
  const b = clientXOf(rect, x1);
  let left = Math.max(a, b) + 6;
  if (left + tw > window.innerWidth - 4) left = Math.max(4, Math.min(a, b) - tw - 6);
  const top = clamp(rect.top + y0, 4, Math.max(4, window.innerHeight - th - 4));
  tip.style.left = `${left}px`;
  tip.style.top = `${top}px`;
}

function hideTooltip() {
  if (G && !G.dom.tooltip.hidden) G.dom.tooltip.hidden = true;
}

/** Ctrl/Cmd held: hyperlinks show a pointer cursor. @param {KeyboardEvent | null} e @param {boolean} down */
function trackModifier(e, down) {
  if (!G) return;
  const on = e ? (e.key === 'Control' || e.key === 'Meta' ? down : e.ctrlKey || e.metaKey) : false;
  root.classList.toggle('fv-ctrl', on);
}

/**
 * Ctrl+wheel zooms the sheet. On mirrored (right-to-left) sheets horizontal wheel/trackpad deltas are
 * applied by hand so content follows the gesture instead of moving the opposite way.
 * @param {WheelEvent} e
 */
function onGridWheel(e) {
  if (!G) return;
  if (!e.ctrlKey) {
    if (isRtl() && e.deltaX) {
      e.preventDefault();
      const unit = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? G.viewH : 1;
      const sc = G.dom.scroller;
      sc.scrollLeft -= e.deltaX * unit;
      sc.scrollTop += e.deltaY * unit;
    }
    return;
  }
  e.preventDefault();
  const now = performance.now();
  if (now - G.zoomWheelAt < 60) return;
  G.zoomWheelAt = now;
  const view = activeView();
  setZoom(view.zoom + (e.deltaY < 0 ? WHEEL_ZOOM_STEP : -WHEEL_ZOOM_STEP));
}

/**
 * Changes the active sheet's zoom keeping the top-left visible cell in place.
 * @param {number} zoom
 */
function setZoom(zoom) {
  if (!G) return;
  const view = activeView();
  zoom = clamp(Math.round(zoom), MIN_ZOOM, MAX_ZOOM);
  if (zoom === view.zoom || !view.L) return;
  closeCellBox();
  const L = view.L;
  const vy = virtualScrollTop();
  const sl = G.dom.scroller.scrollLeft;
  const topRow = findIndex(L.rowTop, L.nRows, L.rowTop[L.fr] + vy);
  const leftCol = findIndex(L.colLeft, L.nCols, L.colLeft[L.fc] + sl);
  view.zoom = zoom;
  applyLayout(view);
  const L2 = /** @type {Layout} */ (view.L);
  setVirtualScroll(view, Math.max(0, L2.rowTop[Math.min(topRow, L2.nRows - 1)] - L2.rowTop[L2.fr]), Math.max(0, L2.colLeft[Math.min(leftCol, L2.nCols - 1)] - L2.colLeft[L2.fc]));
  G.selVersion++;
  updateZoomLabel();
  saveGridState();
  scheduleRender();
}

/**
 * Screen-reader view of the active cell: the gridcell named by the grid's aria-activedescendant gets the address,
 * text and selection, its description the note and link target. Its id alternates so that every move is announced.
 * After keyboard navigation the note / link tooltip of the active cell is shown too (the mouse gets it on hover).
 */
const announceActive = debounce(() => {
  if (!G || G.sheet < 0) return;
  const view = activeView();
  const d = G.dom;
  const a = view.sel.active;
  const cell = cellAt(view, a.r, a.c);
  const text = displayText(cell);
  const addr = colName(a.c) + (a.r + 1);
  const label = selectionLabel(view);
  /** @type {string[]} */
  const desc = [];
  if (cell?.note) desc.push(`Note: ${cell.note}`);
  if (cell?.link) desc.push(`Link: ${linkTarget(cell.link)}, ${IS_MAC ? 'Cmd' : 'Ctrl'}+Enter to follow`);
  d.ariaDesc.textContent = desc.join('. ');
  d.ariaRow.setAttribute('aria-rowindex', String(a.r + 1));
  d.ariaCell.setAttribute('aria-colindex', String(a.c + 1));
  d.ariaCell.setAttribute('aria-label', `${addr}${text ? ', ' + text : ''}${label !== addr ? `, selected ${label}` : ''}`);
  d.ariaCell.textContent = text;
  d.ariaCell.dataset.loaded = String(cell !== undefined);
  d.ariaCell.id = d.ariaCell.id === 'fv-grid-cell-0' ? 'fv-grid-cell-1' : 'fv-grid-cell-0';
  d.scroller.setAttribute('aria-activedescendant', d.ariaCell.id);
  if (G.keyNav && cell && (cell.note || cell.link) && document.activeElement === d.scroller) showCellTooltip(view, a.r, a.c, cell);
}, 150);

// ----- KEYBOARD -----

/** @param {KeyboardEvent} e */
function onGridKeyDown(e) {
  if (!G || e.isComposing) return;
  // From <body>: onLostFocusKeyDown gave the grid its focus back.
  if (e.target !== G.dom.scroller && !(e.target === document.body && document.activeElement === G.dom.scroller)) return;
  const view = activeView();
  const L = view.L;
  if (!L) return;
  const ctrl = e.ctrlKey || e.metaKey;
  const shift = e.shiftKey;
  switch (e.key) {
    case 'ArrowUp':
      moveSelection(view, -1, 0, shift, ctrl);
      break;
    case 'ArrowDown':
      moveSelection(view, 1, 0, shift, ctrl);
      break;
    // Left/right are visual: on a right-to-left sheet column A is on the right.
    case 'ArrowLeft':
      moveSelection(view, 0, L.rtl ? 1 : -1, shift, ctrl);
      break;
    case 'ArrowRight':
      moveSelection(view, 0, L.rtl ? -1 : 1, shift, ctrl);
      break;
    case 'Tab':
      if (ctrl || e.altKey) return;
      cycleInSelection(view, 'row', shift ? -1 : 1);
      break;
    case 'Enter': {
      if (e.altKey) return;
      if (!ctrl) {
        cycleInSelection(view, 'col', shift ? -1 : 1);
        break;
      }
      // Ctrl+Enter follows the active cell's hyperlink: the keyboard twin of Ctrl+Click.
      const a = view.sel.active;
      const cell = cellAt(view, a.r, a.c);
      if (shift || !cell || !cell.link) return;
      openLink(cell.link);
      break;
    }
    case 'Home': {
      const target = ctrl ? { r: visibleRow(L, L.fr), c: visibleCol(L, L.fc) } : { r: (shift ? view.sel.cursor : view.sel.active).r, c: visibleCol(L, L.fc) };
      jumpTo(view, target, shift);
      break;
    }
    case 'End':
      if (!ctrl) return;
      jumpTo(view, { r: visibleRow(L, Math.max(0, view.meta.rowCount - 1)), c: visibleCol(L, Math.max(0, view.meta.colCount - 1)) }, shift);
      break;
    case 'PageDown':
    case 'PageUp':
      if (ctrl) {
        if (!switchSheet(e.key === 'PageDown' ? 1 : -1)) return;
        break;
      }
      pageMove(view, e.key === 'PageDown' ? 1 : -1, shift, e.altKey);
      break;
    case ' ':
      if (shift && !ctrl) selectEntire(view, 'rows');
      else if (ctrl && !shift) selectEntire(view, 'cols');
      else if (ctrl && shift) selectRegionOrAll(view);
      else return;
      break;
    case 'Escape':
      hoverReset();
      closeMenu();
      return;
    case 'F2': {
      // Excel's F2 edits the active cell in place: here the read-only cell text box, caret at the end of the text.
      if (ctrl || shift || e.altKey) return;
      const m = masterOf(view, view.sel.active.r, view.sel.active.c);
      openCellBox(view, m.r, m.c, true);
      break;
    }
    case 'Backspace':
      // Ctrl+Backspace scrolls the active cell into view; Shift+Backspace keeps only the active cell selected (Excel).
      if (e.altKey || ctrl === shift) return;
      if (ctrl) scrollCellIntoView(view, view.sel.active.r, view.sel.active.c);
      else selectCell(view, view.sel.active.r, view.sel.active.c);
      break;
    default: {
      const k = e.key.toLowerCase();
      if (ctrl && !e.altKey && k === 'a') selectRegionOrAll(view);
      else if (ctrl && !e.altKey && !shift && k === 'c') copySelectionFromKeyboard();
      else if (ctrl && !e.altKey && !shift && e.key === '.') nextCorner(view);
      else return;
    }
  }
  // The previous cell's tooltip goes; announceActive shows the new active cell's note / link (keyboard users get
  // no hover).
  G.keyNav = true;
  hoverReset();
  e.preventDefault();
  e.stopPropagation();
}

/**
 * Keys that work wherever the focus is in the grid view (the grid's own handler runs first and stops what it
 * handles): F6 / Shift+F6 move between the view's regions like Excel's F6 (past the first / last region the key is
 * left to VS Code, whose F6 moves on to the next workbench part), Ctrl+G goes to the name box (Excel's Go To),
 * Ctrl+PageUp / PageDown switch sheets, Ctrl+A selects no page text.
 * @param {KeyboardEvent} e
 */
function onGridAppKeyDown(e) {
  if (!G || G.sheet < 0 || e.defaultPrevented || e.isComposing) return;
  const ctrl = e.ctrlKey || e.metaKey;
  if (e.key === 'F6' && !ctrl && !e.altKey) {
    if (!focusRegion(e.shiftKey ? -1 : 1)) return;
  } else if (ctrl && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'g') {
    closeMenu();
    G.dom.nameBox.focus();
  } else if (ctrl && !e.altKey && !e.shiftKey && (e.key === 'PageDown' || e.key === 'PageUp')) {
    const focus = document.activeElement;
    const inTabs = !!focus && G.dom.tabBar.contains(focus);
    if (!switchSheet(e.key === 'PageDown' ? 1 : -1)) return;
    // The tab strip was rebuilt: keep the focus on the (new) active tab; from the menu, go back to the grid.
    if (inTabs) /** @type {HTMLElement | null} */ (G.dom.tabStrip.querySelector('.fv-tab-active'))?.focus();
    else if (!focus || focus === document.body || !focus.isConnected) G.dom.scroller.focus({ preventScroll: true });
  } else if (ctrl && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'a') {
    // On a tab, a bar button or the menu, select-all would highlight the whole view's text (the browser's, or VS
    // Code's run on the page): nothing to select there. The name box and the cell text box select their own text.
    const t = /** @type {HTMLElement} */ (e.target);
    if (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA') return;
  } else return;
  e.preventDefault();
  e.stopPropagation();
}

/**
 * A key pressed while the focus is lost on <body> (VS Code's context menu or the workbench took it, see
 * wireGridEvents; or the focused element went away): the grid takes the focus back and handles the key, so the
 * keyboard goes on working where it left off. When the cell text box had the focus, the box takes it back.
 * @param {KeyboardEvent} e
 */
function onLostFocusKeyDown(e) {
  if (!G || G.sheet < 0 || e.target !== document.body || e.defaultPrevented) return;
  const box = G.cellBox;
  if (box && G.lostFocus === box.el) {
    box.menu = false;
    box.el.focus({ preventScroll: true });
    return;
  }
  G.dom.scroller.focus({ preventScroll: true });
  onGridKeyDown(e);
  if (!e.defaultPrevented) onGridAppKeyDown(e);
}

/** The frame got the focus back: the grid / cell text box that lost it to <body> takes it again. */
function restoreLostFocus() {
  const el = G?.lostFocus;
  const focus = document.activeElement;
  if (!G || !el || !el.isConnected || (focus && focus !== document.body)) return;
  if (G.cellBox?.el === el) G.cellBox.menu = false;
  el.focus({ preventScroll: true });
}

/**
 * Moves the focus to the next / previous region of the grid view: banner, formula bar (name box), grid, sheet tabs,
 * status bar. Lost focus (<body>) counts as the grid. Returns false at either end.
 * @param {number} dir
 */
function focusRegion(dir) {
  const d = /** @type {GridState} */ (G).dom;
  /** @type {[HTMLElement | null, HTMLElement | null][]} region container, element that takes the focus */
  const all = [
    [d.banner?.isConnected ? d.banner : null, d.banner?.querySelector('button') ?? null],
    [d.formulaBar, d.nameBox],
    [d.scroller, d.scroller],
    [d.tabBar, d.tabStrip.querySelector('.fv-tab-active')],
    [d.statusBar, d.statusBar.querySelector('button')],
  ];
  const regions = all.filter((x) => x[0] && x[1]);
  // The cell text box belongs to the grid region.
  const cellBox = /** @type {GridState} */ (G).cellBox;
  const focus = cellBox && document.activeElement === cellBox.el ? d.scroller : document.activeElement;
  let at = regions.findIndex(([box]) => !!focus && !!box && box.contains(focus));
  if (at < 0 && d.menu && focus && d.menu.contains(focus)) at = regions.findIndex(([box]) => box === d.tabBar);
  const next = at < 0 ? regions.findIndex(([box]) => box === d.scroller) : at + dir;
  if (next < 0 || next >= regions.length) return false;
  closeMenu();
  /** @type {HTMLElement} */ (regions[next][1]).focus({ preventScroll: true });
  return true;
}

/**
 * Arrow-key movement; shift extends from the anchor, ctrl jumps to the data edge (asking the host when the rows
 * on the way are not loaded).
 * @param {SheetView} view @param {number} dr @param {number} dc @param {boolean} extend @param {boolean} jump
 */
function moveSelection(view, dr, dc, extend, jump) {
  if (jump) {
    // The origin is read when the search runs: a queued Ctrl+Arrow continues from where the previous one landed.
    searchThen(
      () => {
        const from = moveOrigin(view, dr, dc, extend);
        return dataEdge(view, from.r, from.c, dr, dc);
      },
      (alive) => {
        const from = moveOrigin(view, dr, dc, extend);
        return dataEdgeAsync(view, from.r, from.c, dr, dc, alive);
      },
      (target) => moveTo(view, target, dr, dc, extend),
    );
    return;
  }
  const L = /** @type {Layout} */ (view.L);
  const from = moveOrigin(view, dr, dc, extend);
  moveTo(view, { r: dr ? stepRow(L, from.r, dr) : from.r, c: dc ? stepCol(L, from.c, dc) : from.c }, dr, dc, extend);
}

/**
 * Where an arrow key starts. Extending: the moving corner, or the far edge of the anchor's merged cell while the
 * range does not reach past it (else the first Shift+Arrow from a merged cell would land inside it). Moving: the
 * active cell, leaving a merged cell from its far edge on the row / column the cursor entered it (Excel).
 * @param {SheetView} view @param {number} dr @param {number} dc @param {boolean} extend
 * @returns {Cell}
 */
function moveOrigin(view, dr, dc, extend) {
  const sel = view.sel;
  if (extend) {
    const a = cellArea(view, sel.anchor.r, sel.anchor.c);
    let { r, c } = sel.cursor;
    if (dr && r >= a.r0 && r <= a.r1) r = dr > 0 ? a.r1 : a.r0;
    if (dc && c >= a.c0 && c <= a.c1) c = dc > 0 ? a.c1 : a.c0;
    return { r, c };
  }
  const area = cellArea(view, sel.active.r, sel.active.c);
  const at = selectionContainsIn([area], sel.cursor) ? sel.cursor : sel.active;
  return { r: dr > 0 ? area.r1 : dr < 0 ? area.r0 : at.r, c: dc > 0 ? area.c1 : dc < 0 ? area.c0 : at.c };
}

/**
 * Completes an arrow-key move to `target`.
 * @param {SheetView} view @param {Cell} target @param {number} dr @param {number} dc @param {boolean} extend
 */
function moveTo(view, target, dr, dc, extend) {
  const sel = view.sel;
  if (extend) {
    let rg = extendedRange(view, target);
    // When a merge grew the range past the cursor, continue from the range edge next time.
    const cursor = { r: target.r, c: target.c };
    if (dr > 0 && rg.r1 > cursor.r && cursor.r >= sel.anchor.r) cursor.r = rg.r1;
    if (dr < 0 && rg.r0 < cursor.r && cursor.r <= sel.anchor.r) cursor.r = rg.r0;
    if (dc > 0 && rg.c1 > cursor.c && cursor.c >= sel.anchor.c) cursor.c = rg.c1;
    if (dc < 0 && rg.c0 < cursor.c && cursor.c <= sel.anchor.c) cursor.c = rg.c0;
    rg = extendedRange(view, cursor);
    setSelection(view, sel.ranges.slice(0, -1).concat([rg]), sel.active, { anchor: sel.anchor, cursor, reveal: cursor });
  } else {
    // A merged cell becomes active through its top-left cell; the cursor remembers where the move arrived.
    const m = masterOf(view, target.r, target.c);
    setSelection(view, [cellArea(view, m.r, m.c)], m, { cursor: { r: target.r, c: target.c } });
  }
}

/**
 * Moves (or extends) to a cell.
 * @param {SheetView} view @param {Cell} target @param {boolean} extend
 */
function jumpTo(view, target, extend) {
  if (extend) {
    const sel = view.sel;
    setSelection(view, sel.ranges.slice(0, -1).concat([extendedRange(view, target)]), sel.active, { anchor: sel.anchor, cursor: target, reveal: target });
  } else selectCell(view, target.r, target.c);
}

// Keyboard searches that may need the host (Ctrl+Arrow, Ctrl+A over rows that are not loaded) run one at a time in
// key order, so pressing Ctrl+Down twice hops twice. A result is dropped when the selection changed meanwhile.
const searches = { chain: /** @type {Promise<void>} */ (Promise.resolve()), pending: 0 };
const SEARCHING = 'Searching…';
/** Thrown by a search whose result is no longer wanted (the selection changed): it stops asking the host. */
const SEARCH_STALE = new Error('The selection changed during the search.');

/**
 * Runs a keyboard search: `quick` answers from loaded rows (null = rows missing), else `slow` asks the host and
 * calls `alive()` before every request (it throws SEARCH_STALE once the result would be dropped).
 * @template T
 * @param {() => T | null} quick @param {(alive: () => void) => Promise<T>} slow @param {(result: T) => void} apply
 */
function searchThen(quick, slow, apply) {
  const grid = G;
  if (!grid) return;
  if (!searches.pending) {
    const result = quick();
    if (result !== null) {
      apply(result);
      return;
    }
  }
  searches.pending++;
  searches.chain = searches.chain
    .then(async () => {
      if (G !== grid) return;
      const seq = grid.selVersion;
      const sheet = grid.sheet;
      const current = () => G === grid && grid.selVersion === seq && grid.sheet === sheet;
      const alive = () => {
        if (!current()) throw SEARCH_STALE;
      };
      const busy = setTimeout(() => G === grid && setStatusMessage(SEARCHING, 'info', 0), 400);
      try {
        const result = quick() ?? (await slow(alive));
        if (current()) apply(result);
      } finally {
        clearTimeout(busy);
        if (G === grid && grid.dom.statusMsg.textContent === SEARCHING) setStatusMessage('', 'info');
      }
    })
    // Superseded, or the data changed under the search (reload, edit): the key press is dropped.
    .catch((err) => {
      if (err !== SEARCH_STALE) log('warn', `Keyboard search stopped: ${errorDetail(err)}`);
    })
    .finally(() => {
      searches.pending--;
    });
}

/** Non-empty for Ctrl+Arrow / Ctrl+A: any value, or a formula even when its result is "" (Excel). @param {CellData | null | undefined} cell */
function isFilled(cell) {
  return !!cell && (!isEmptyCell(cell) || (!!cell.f && cell.t !== 'z'));
}

/** @param {Range} a @param {Range} b */
function sameRange(a, b) {
  return a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;
}

/**
 * Excel's Ctrl+Arrow (Range.End): from a non-empty cell whose neighbour is non-empty → the last non-empty cell of the
 * run; otherwise → the next non-empty cell; none → the edge of the sheet (row 1,048,576 / column XFD). Hidden rows
 * and columns are skipped. Answers from loaded rows; null when a row it needs is not loaded (dataEdgeAsync).
 * @param {SheetView} view @param {number} r @param {number} c @param {number} dr @param {number} dc
 * @param {(r: number) => RowData | null | undefined} [rowOf]  row source (default: the row cache)
 * @returns {Cell | null}
 */
function dataEdge(view, r, c, dr, dc, rowOf = (x) => getRow(view.index, x)) {
  const L = /** @type {Layout} */ (view.L);
  const sm = view.meta;
  const step = (/** @type {Cell} */ p) => ({ r: dr ? stepRow(L, p.r, dr) : p.r, c: dc ? stepCol(L, p.c, dc) : p.c });
  const same = (/** @type {Cell} */ a, /** @type {Cell} */ b) => a.r === b.r && a.c === b.c;
  /** @returns {boolean | undefined} */
  const filled = (/** @type {Cell} */ p) => {
    if (p.r >= sm.rowCount || p.c >= sm.colCount) return false;
    const row = rowOf(p.r);
    if (row === undefined) return undefined;
    return isFilled(row ? findCell(row.cells, p.c) : null);
  };
  const edge = () => sheetEdge(view, r, c, dr, dc);
  // Nothing beyond the used range: straight to the edge of the sheet.
  if ((dr > 0 && r >= sm.rowCount - 1) || (dc > 0 && c >= sm.colCount - 1)) return edge();
  const cur = { r, c };
  const next = step(cur);
  if (same(next, cur)) return edge();
  const curFilled = filled(cur);
  const nextFilled = filled(next);
  if (curFilled === undefined || nextFilled === undefined) return null;
  if (curFilled && nextFilled) {
    for (let p = next; ; ) {
      const n2 = step(p);
      if (same(n2, p)) return p;
      const f = filled(n2);
      if (f === undefined) return null;
      if (!f) return p;
      p = n2;
    }
  }
  for (let p = next; ; ) {
    const f = filled(p);
    if (f === undefined) return null;
    if (f) return p;
    if ((dr > 0 && p.r >= sm.rowCount) || (dc > 0 && p.c >= sm.colCount)) return edge();
    const n2 = step(p);
    if (same(n2, p)) return edge();
    p = n2;
  }
}

/**
 * Last row / column of the sheet in the direction (first visible one going up / left).
 * @param {SheetView} view @param {number} r @param {number} c @param {number} dr @param {number} dc
 * @returns {Cell}
 */
function sheetEdge(view, r, c, dr, dc) {
  const L = /** @type {Layout} */ (view.L);
  if (dr > 0) return { r: view.maxR - 1, c };
  if (dr < 0) return { r: visibleRow(L, 0), c };
  if (dc > 0) return { r, c: view.maxC - 1 };
  return { r, c: visibleCol(L, 0) };
}

/**
 * dataEdge when rows on the way are not loaded. Left / right needs one row (fetched); up / down counts the column's
 * non-empty cells with the host and binary-searches the edge (about 20 requests for a million rows) instead of
 * loading the rows. Hidden rows are left out of the counted ranges, so they are skipped as in dataEdge.
 * @param {SheetView} view @param {number} r @param {number} c @param {number} dr @param {number} dc
 * @param {() => void} alive  throws when the search is superseded (checked before every request)
 * @returns {Promise<Cell>}
 */
async function dataEdgeAsync(view, r, c, dr, dc, alive) {
  if (dc) {
    alive();
    const rows = await fetchRows(view.index, r, r);
    return dataEdge(view, r, c, dr, dc, (x) => (x === r ? (rows.get(r) ?? null) : getRow(view.index, x))) ?? sheetEdge(view, r, c, dr, dc);
  }
  const L = /** @type {Layout} */ (view.L);
  const last = view.meta.rowCount - 1; // no data below
  const hidden = hiddenRows(view);
  /** Non-empty cells / visible rows of column c in rows [a, b]. */
  const span = (/** @type {number} */ a, /** @type {number} */ b) => columnSpan(hidden, a, b, c);
  const count = async (/** @type {number} */ a, /** @type {number} */ b) => {
    const s = span(a, b);
    if (!s.ranges.length) return 0;
    alive();
    return countCells(view.index, s.ranges);
  };
  const filled = async (/** @type {number} */ x) => x <= last && (await count(x, x)) > 0;
  const edge = sheetEdge(view, r, c, dr, 0);
  if (dr > 0) {
    if (r >= last) return edge;
    const next = stepRow(L, r, 1);
    if ((await filled(r)) && (await filled(next))) {
      // The run ends before the first visible empty row after `next` (or at the last visible row with data).
      if ((await count(next + 1, last)) === span(next + 1, last).visible) return { r: stepRow(L, last + 1, -1), c };
      let lo = next + 1;
      let hi = last;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if ((await count(next + 1, mid)) < span(next + 1, mid).visible) hi = mid;
        else lo = mid + 1;
      }
      return { r: stepRow(L, lo, -1), c };
    }
    if ((await count(r + 1, last)) === 0) return edge;
    let lo = r + 1;
    let hi = last;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if ((await count(r + 1, mid)) > 0) hi = mid;
      else lo = mid + 1;
    }
    return { r: lo, c };
  }
  const prev = stepRow(L, r, -1);
  if (prev === r) return edge;
  if ((await filled(r)) && (await filled(prev))) {
    // The run starts after the last visible empty row before `prev` (or at the first visible row).
    if (prev === 0 || (await count(0, prev - 1)) === span(0, prev - 1).visible) return edge;
    let lo = 0;
    let hi = prev - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if ((await count(mid, prev - 1)) < span(mid, prev - 1).visible) lo = mid;
      else hi = mid - 1;
    }
    return { r: stepRow(L, lo, 1), c };
  }
  if ((await count(0, r - 1)) === 0) return edge;
  let lo = 0;
  let hi = r - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if ((await count(mid, r - 1)) > 0) lo = mid;
    else hi = mid - 1;
  }
  return { r: lo, c };
}

/** Hidden (zero-height) rows of the view's layout, ascending. @param {SheetView} view */
function hiddenRows(view) {
  const L = /** @type {Layout} */ (view.L);
  /** @type {number[]} */
  const out = [];
  for (const key in view.meta.rows) {
    const r = Number(key);
    if (r >= 0 && r < L.nRows && rowHeight(L, r) === 0) out.push(r);
  }
  return out.sort((a, b) => a - b);
}

/**
 * Column c, rows [a, b] without the hidden rows: the ranges to count and the number of visible rows. Very many
 * hidden rows (more than the host takes ranges) fall back to the whole span.
 * @param {number[]} hidden ascending @param {number} a @param {number} b @param {number} c
 * @returns {{ ranges: Range[], visible: number }}
 */
function columnSpan(hidden, a, b, c) {
  if (a > b) return { ranges: [], visible: 0 };
  /** @type {Range[]} */
  const ranges = [];
  let start = a;
  let lo = 0;
  let hi = hidden.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (hidden[mid] < a) lo = mid + 1;
    else hi = mid;
  }
  let skipped = 0;
  for (let i = lo; i < hidden.length && hidden[i] <= b; i++) {
    if (hidden[i] > start) ranges.push({ r0: start, r1: hidden[i] - 1, c0: c, c1: c });
    start = hidden[i] + 1;
    skipped++;
  }
  if (start <= b) ranges.push({ r0: start, r1: b, c0: c, c1: c });
  if (ranges.length > 1000) return { ranges: [{ r0: a, r1: b, c0: c, c1: c }], visible: b - a + 1 };
  return { ranges, visible: b - a + 1 - skipped };
}

/**
 * Tab / Enter inside a selection of more than one cell: the active cell walks through the range (row by row for
 * Tab, column by column for Enter, backwards with Shift) and on into the next range of a multi-range selection,
 * like Excel. A single cell moves right / down instead.
 * @param {SheetView} view @param {'row' | 'col'} major @param {number} delta
 */
function cycleInSelection(view, major, delta) {
  const L = /** @type {Layout} */ (view.L);
  const sel = view.sel;
  const area = cellArea(view, sel.active.r, sel.active.c);
  if (sel.ranges.length === 1 && sameRange(sel.ranges[0], area)) {
    moveSelection(view, major === 'col' ? delta : 0, major === 'row' ? delta : 0, false, false);
    return;
  }
  const n = sel.ranges.length;
  const clip = (/** @type {Range} */ rg) => ({ r0: rg.r0, c0: rg.c0, r1: Math.min(rg.r1, L.nRows - 1), c1: Math.min(rg.c1, L.nCols - 1) });
  // The range holding the active cell (the latest one that does).
  let k = n - 1;
  while (k > 0 && !selectionContainsIn([sel.ranges[k]], sel.active)) k--;
  let rg = clip(sel.ranges[k]);
  let { r, c } = sel.active;
  for (let guard = 0; guard < 100000; guard++) {
    let wrapped = false;
    if (major === 'row') {
      c += delta;
      if (c > rg.c1) (c = rg.c0), (r += 1), (wrapped = r > rg.r1);
      else if (c < rg.c0) (c = rg.c1), (r -= 1), (wrapped = r < rg.r0);
    } else {
      r += delta;
      if (r > rg.r1) (r = rg.r0), (c += 1), (wrapped = c > rg.c1);
      else if (r < rg.r0) (r = rg.r1), (c -= 1), (wrapped = c < rg.c0);
    }
    if (wrapped) {
      // Past the end (start) of this range: the first (last) cell of the next (previous) range.
      k = (k + delta + n) % n;
      rg = clip(sel.ranges[k]);
      r = delta > 0 ? rg.r0 : rg.r1;
      c = delta > 0 ? rg.c0 : rg.c1;
    }
    if (rowHeight(L, r) === 0 || colWidth(L, c) === 0) continue;
    const m = masterOf(view, r, c);
    if (m.r !== r || m.c !== c) continue;
    break;
  }
  if (k === n - 1) view.sel = { ...sel, active: { r, c } };
  else {
    // The range holding the active cell becomes the last one (the one Shift+Arrow / Shift+click extend).
    const ranges = [...sel.ranges.slice(k + 1), ...sel.ranges.slice(0, k + 1)];
    const cur = ranges[n - 1];
    view.sel = { ranges, active: { r, c }, anchor: { r: cur.r0, c: cur.c0 }, cursor: { r: cur.r1, c: cur.c1 } };
  }
  selectionChanged(view, { r, c });
}

/**
 * Ctrl+. (Excel): the active cell moves clockwise to the next corner of its range (top-left, top-right, bottom-right,
 * bottom-left; from a cell that is no corner, to the top-left one), and the range is now extended from that corner, so
 * Shift+Arrow moves the opposite corner. The selection itself does not change.
 * @param {SheetView} view
 */
function nextCorner(view) {
  const sel = view.sel;
  const n = sel.ranges.length;
  // The range holding the active cell (the latest one that does).
  let k = n - 1;
  while (k > 0 && !selectionContainsIn([sel.ranges[k]], sel.active)) k--;
  const rg = sel.ranges[k];
  /** @type {Cell[]} */
  const corners = [
    { r: rg.r0, c: rg.c0 },
    { r: rg.r0, c: rg.c1 },
    { r: rg.r1, c: rg.c1 },
    { r: rg.r1, c: rg.c0 },
  ];
  // A merged active cell is at every corner it covers: the next stop is the first corner clockwise after them.
  const area = cellArea(view, sel.active.r, sel.active.c);
  const inArea = corners.map((p) => selectionContainsIn([area], p));
  if (inArea.every(Boolean)) return; // one cell or one merged area: no other corner
  const next = inArea.some(Boolean) ? [0, 1, 2, 3].find((j) => !inArea[j] && inArea[(j + 3) % 4]) ?? 0 : 0;
  const corner = corners[next];
  const opposite = corners[(next + 2) % 4];
  const active = masterOf(view, corner.r, corner.c);
  // The range holding the active cell becomes the last one (the one Shift+Arrow / Shift+click extend).
  const ranges = k === n - 1 ? sel.ranges : [...sel.ranges.slice(k + 1), ...sel.ranges.slice(0, k + 1)];
  setSelection(view, ranges, active, { anchor: corner, cursor: opposite });
}

/**
 * PageUp / PageDown (Alt: left / right) like Excel: scrolls by the rows (columns) that fit on the screen, so every
 * page starts with a whole row, and moves the active cell by the same number of rows. Up undoes down.
 * @param {SheetView} view @param {number} dir @param {boolean} extend @param {boolean} horizontal
 */
function pageMove(view, dir, extend, horizontal) {
  const g = /** @type {GridState} */ (G);
  const L = /** @type {Layout} */ (view.L);
  const from = extend ? view.sel.cursor : view.sel.active;
  const lines = horizontal ? L.colLeft : L.rowTop;
  const n = horizontal ? L.nCols : L.nRows;
  const frozen = horizontal ? L.fc : L.fr;
  const room = Math.max(1, horizontal ? g.viewW - L.B : g.viewH - L.A);
  const pos = lines[frozen] + (horizontal ? g.dom.scroller.scrollLeft : virtualScrollTop());
  // First line fully shown in the scrolling area.
  let top = findIndex(lines, n, pos);
  if (lines[top] < pos - 0.5) top = Math.min(n - 1, top + 1);
  // First line not fully shown below the window = one page further on.
  const pageEnd = Math.max(top + 1, findIndex(lines, n, lines[top] + room));
  let next;
  if (dir > 0) next = Math.min(n - 1, pageEnd);
  else {
    next = findIndex(lines, n, lines[top] - room);
    if (lines[next] < lines[top] - room - 0.5) next++;
    next = clamp(next, frozen, top);
  }
  // At the top (end) the window cannot move: the active cell still moves by a page.
  const delta = next !== top ? next - top : dir * (pageEnd - top);
  const at = clamp((horizontal ? from.c : from.r) + delta, 0, n - 1);
  const offset = lines[next] - lines[frozen];
  if (horizontal) setVirtualScroll(view, undefined, offset);
  else setVirtualScroll(view, offset, undefined);
  jumpTo(view, horizontal ? { r: from.r, c: visibleCol(L, at) } : { r: visibleRow(L, at), c: from.c }, extend);
}

/**
 * Ctrl+PageUp/PageDown: previous/next visible sheet. Returns false when there is none (a one-sheet file, CSV, or the
 * first / last sheet): the key is then left to VS Code, whose Ctrl+PageUp/PageDown goes to the previous/next editor.
 * @param {number} dir
 */
function switchSheet(dir) {
  if (!G) return false;
  const order = G.meta.sheets.filter((s) => s.state === 'visible' || G?.revealed.has(s.index)).map((s) => s.index);
  const pos = order.indexOf(G.sheet);
  const next = order[pos + dir];
  if (next === undefined) return false;
  activateSheet(next);
  return true;
}

/**
 * Ctrl+A / Ctrl+Shift+Space like Excel: the current region around the active cell first; the whole sheet when that
 * region is already selected or the active cell has no data around it. The active cell stays.
 * @param {SheetView} view
 */
function selectRegionOrAll(view) {
  // Pressed again right after it selected the region: the whole sheet, without searching again.
  if (G && lastRegion.sheet === G.sheet && lastRegion.version === G.selVersion) {
    selectAll(view);
    return;
  }
  searchThen(
    () => null,
    (alive) => currentRegion(view, alive),
    (rg) => {
      const sel = view.sel;
      const area = cellArea(view, sel.active.r, sel.active.c);
      if (sameRange(rg, area) || (sel.ranges.length === 1 && sameRange(sel.ranges[0], rg))) selectAll(view);
      else {
        setSelection(view, [rg], sel.active, { anchor: { r: rg.r0, c: rg.c0 }, cursor: { r: rg.r1, c: rg.c1 }, reveal: null });
        if (G) lastRegion = { sheet: view.index, version: G.selVersion };
      }
    },
  );
}

/** The selection Ctrl+A made last (sheet, G.selVersion right after it). */
let lastRegion = { sheet: -1, version: -1 };

/**
 * Excel's CurrentRegion of the active cell: grown while a row or column next to it (diagonals included) has a
 * non-empty cell, and over merged cells. Counting with the host skips whole blocks of rows / columns while the
 * region's cells there are all filled; otherwise columns are checked one count at a time and rows are read row by
 * row (one 256-row fetch per block), so big tables do not have to be loaded.
 * @param {SheetView} view
 * @param {() => void} alive  throws when the search is superseded (checked before every request)
 * @returns {Promise<Range>}
 */
async function currentRegion(view, alive) {
  const sheet = view.index;
  const lastR = view.meta.rowCount - 1;
  const lastC = view.meta.colCount - 1;
  const a = view.sel.active;
  let rg = cellArea(view, a.r, a.c);
  const count = (/** @type {Range} */ r) => {
    alive();
    return countCells(sheet, [r]);
  };
  /** Rows fetched for this search. @type {Map<number, RowData | null>} */
  const fetched = new Map();
  const rowHas = async (/** @type {number} */ r, /** @type {number} */ c0, /** @type {number} */ c1) => {
    if (r < 0 || r > lastR) return false;
    let row = getRow(sheet, r);
    if (row === undefined) {
      if (!fetched.has(r)) {
        const b0 = Math.floor(r / BLOCK_ROWS) * BLOCK_ROWS;
        const b1 = Math.min(lastR, b0 + BLOCK_ROWS - 1);
        alive();
        const rows = await fetchRows(sheet, b0, b1);
        for (let x = b0; x <= b1; x++) fetched.set(x, rows.get(x) ?? null);
      }
      row = fetched.get(r) ?? null;
    }
    if (!row) return false;
    for (let i = lowerBound(row.cells, Math.max(0, c0)); i < row.cells.length && row.cells[i].c <= c1; i++) if (isFilled(row.cells[i])) return true;
    return false;
  };
  const growRows = async (/** @type {number} */ dir) => {
    let grew = false;
    let jump = BLOCK_ROWS;
    for (;;) {
      const edge = dir > 0 ? rg.r1 : rg.r0;
      if (dir > 0 ? edge >= lastR : edge <= 0) return grew;
      const first = edge + dir;
      const far = dir > 0 ? Math.min(lastR, edge + jump) : Math.max(0, edge - jump);
      const lo = Math.min(first, far);
      const hi = Math.max(first, far);
      const width = rg.c1 - rg.c0 + 1;
      if (hi > lo && (await count({ r0: lo, r1: hi, c0: rg.c0, c1: rg.c1 })) === (hi - lo + 1) * width) {
        rg = dir > 0 ? { ...rg, r1: hi } : { ...rg, r0: lo };
        grew = true;
        jump = Math.min(jump * 2, Math.max(BLOCK_ROWS, Math.floor(PROBE_CELLS / width)));
        continue;
      }
      jump = BLOCK_ROWS;
      // Row by row to the end of this block of rows.
      const blockEnd = dir > 0 ? Math.min(lastR, (Math.floor(first / BLOCK_ROWS) + 1) * BLOCK_ROWS - 1) : Math.floor(first / BLOCK_ROWS) * BLOCK_ROWS;
      let r = first;
      while ((dir > 0 ? r <= blockEnd : r >= blockEnd) && (await rowHas(r, rg.c0 - 1, rg.c1 + 1))) r += dir;
      if (r !== first) {
        rg = dir > 0 ? { ...rg, r1: r - 1 } : { ...rg, r0: r + 1 };
        grew = true;
      }
      if (dir > 0 ? r <= blockEnd : r >= blockEnd) return grew; // stopped at an empty row
    }
  };
  const growCols = async (/** @type {number} */ dir) => {
    let grew = false;
    let jump = 8;
    for (;;) {
      const edge = dir > 0 ? rg.c1 : rg.c0;
      if (dir > 0 ? edge >= lastC : edge <= 0) return grew;
      const first = edge + dir;
      const far = dir > 0 ? Math.min(lastC, edge + jump) : Math.max(0, edge - jump);
      const lo = Math.min(first, far);
      const hi = Math.max(first, far);
      const height = rg.r1 - rg.r0 + 1;
      if (hi > lo && (await count({ r0: rg.r0, r1: rg.r1, c0: lo, c1: hi })) === (hi - lo + 1) * height) {
        rg = dir > 0 ? { ...rg, c1: hi } : { ...rg, c0: lo };
        grew = true;
        jump = Math.min(jump * 2, Math.max(1, Math.floor(PROBE_CELLS / height)));
        continue;
      }
      jump = 8;
      if (!(await count({ r0: Math.max(0, rg.r0 - 1), r1: Math.min(lastR, rg.r1 + 1), c0: first, c1: first }))) return grew;
      rg = dir > 0 ? { ...rg, c1: first } : { ...rg, c0: first };
      grew = true;
    }
  };
  for (let round = 0; round < 64; round++) {
    let grew = false;
    for (const dir of [-1, 1]) if (await growRows(dir)) grew = true;
    for (const dir of [-1, 1]) if (await growCols(dir)) grew = true;
    const snapped = snapRange(view, rg);
    if (!sameRange(snapped, rg)) {
      rg = snapped;
      grew = true;
    }
    if (!grew) break;
  }
  return rg;
}

// ----- CLIPBOARD -----
// Ctrl+C copies the selection as TSV (display text; fields with tab, newline or quote are quoted) plus a
// simple HTML table. Hidden rows and columns (hidden in the file or by a filter) are left out, as in the text and
// HTML Excel puts on the clipboard; a selection of hidden cells only keeps them. Text cells of a workbook are marked as
// text in the HTML (mso-number-format, as Excel marks them), so pasting into Excel keeps '007', '1-2' or '=1+1' as text
// instead of turning them into numbers, dates or formulas. A CSV file has no cell types: its fields are left to Excel's
// paste, which converts them as Excel does when it opens the file. One cell (or one merged area) is copied as its
// display text alone: plain text, never quoted, no table and no line break added. Rows that are not cached are fetched
// first, then copied the same way (or written as plain text with the async clipboard API when the browser refuses a
// copy event).

/**
 * 'cell' = one cell or one merged area (plain text only), 'single' = one range, 'rows' / 'cols' = several ranges
 * stacked vertically / side by side.
 * @typedef {'cell' | 'single' | 'rows' | 'cols'} CopyMode
 */

/** Text prepared for the next `copy` event (set by the keyboard handler); html '' = plain text only. @type {{ text: string, html: string, cells: number } | null} */
let pendingCopy = null;

function copySelectionFromKeyboard() {
  if (!G) return;
  const plan = planCopy(activeView());
  if (typeof plan === 'string') {
    setStatusMessage(plan, 'warn');
    return;
  }
  const payload = buildCopyPayload(plan.view, plan.ranges, plan.mode, (r) => getRow(plan.view.index, r));
  if (!payload) {
    copyAsync(plan);
    return;
  }
  copyPayload(payload);
}

/**
 * Puts a payload on the clipboard through a copy event (text + HTML), else (execCommand refused, or no copy event
 * reached onCopyEvent) as plain text with the async clipboard API.
 * @param {{ text: string, html: string, cells: number }} payload
 */
function copyPayload(payload) {
  pendingCopy = payload;
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  if (!ok || pendingCopy) {
    pendingCopy = null;
    writeClipboard(payload.text, payload.cells);
  }
}

/**
 * Native copy (VS Code's context menu or Edit menu, Cmd+C forwarded by VS Code) while the grid has the focus, or
 * while the focus is lost on <body> after the grid or the cell text box had it: VS Code's context menu takes the
 * focus away before its Copy runs (see wireGridEvents). The box then gives its selected text.
 * @param {ClipboardEvent} e
 */
function onCopyEvent(e) {
  if (!G || !e.clipboardData) return;
  const focus = document.activeElement;
  const lost = !focus || focus === document.body ? G.lostFocus : null;
  const box = G.cellBox;
  if (box && lost === box.el) {
    const text = box.el.value.slice(box.el.selectionStart, box.el.selectionEnd);
    if (text) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', text);
    }
    return;
  }
  if (focus !== G.dom.scroller && lost !== G.dom.scroller) return;
  let payload = pendingCopy;
  pendingCopy = null;
  if (!payload) {
    const plan = planCopy(activeView());
    if (typeof plan === 'string') {
      e.preventDefault();
      setStatusMessage(plan, 'warn');
      return;
    }
    payload = buildCopyPayload(plan.view, plan.ranges, plan.mode, (r) => getRow(plan.view.index, r));
    if (!payload) {
      e.preventDefault();
      copyAsync(plan);
      return;
    }
  }
  e.preventDefault();
  e.clipboardData.setData('text/plain', payload.text);
  if (payload.html) e.clipboardData.setData('text/html', payload.html);
  setStatusMessage(payload.cells ? `Copied ${formatCount(payload.cells)} cell${payload.cells === 1 ? '' : 's'}` : 'Copied', 'info');
}

/**
 * Validates the selection for copying (Excel rules for multiple ranges).
 * @param {SheetView} view
 * @returns {{ view: SheetView, ranges: Range[], mode: CopyMode, cells: number } | string}
 */
function planCopy(view) {
  const sm = view.meta;
  const sel = view.sel;
  const area = cellArea(view, sel.active.r, sel.active.c);
  if (sel.ranges.length === 1 && sameRange(sel.ranges[0], area)) return { view, ranges: [area], mode: 'cell', cells: 1 };
  const ranges = view.sel.ranges.map((rg) => ({
    r0: rg.r0,
    c0: rg.c0,
    r1: isFullCols(view, rg) ? Math.max(rg.r0, sm.rowCount - 1) : rg.r1,
    c1: isFullRows(view, rg) ? Math.max(rg.c0, sm.colCount - 1) : rg.c1,
  }));
  /** @type {CopyMode} */
  let mode = 'single';
  if (ranges.length > 1) {
    const sameRows = ranges.every((rg) => rg.r0 === ranges[0].r0 && rg.r1 === ranges[0].r1);
    const sameCols = ranges.every((rg) => rg.c0 === ranges[0].c0 && rg.c1 === ranges[0].c1);
    if (sameRows) {
      mode = 'cols';
      ranges.sort((a, b) => a.c0 - b.c0);
    } else if (sameCols) {
      mode = 'rows';
      ranges.sort((a, b) => a.r0 - b.r0);
    } else return "This action won't work on multiple selections.";
  }
  let cells = 0;
  for (const rg of ranges) cells += (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1);
  if (cells > COPY_CELL_LIMIT) return `The selection is too large to copy (${formatCount(cells)} cells; the limit is ${formatCount(COPY_CELL_LIMIT)}).`;
  return { view, ranges, mode, cells };
}

/**
 * Builds TSV + HTML for the ranges (mode 'cell': the cell's display text, html ''); null when a needed row is not
 * available.
 * @param {SheetView} view @param {Range[]} ranges @param {CopyMode} mode
 * @param {(r: number) => RowData | null | undefined} rowOf
 * @returns {{ text: string, html: string, cells: number } | null}
 */
function buildCopyPayload(view, ranges, mode, rowOf) {
  if (mode === 'cell') {
    const { r0, c0 } = ranges[0];
    const row = rowOf(r0);
    if (row === undefined) return null;
    return { text: displayText(row ? findCell(row.cells, c0) : null), html: '', cells: 1 };
  }
  const sm = view.meta;
  /** Rows (columns) of [a, b] that are not hidden; all of them when every one is hidden. */
  const shown = (/** @type {number} */ a, /** @type {number} */ b, /** @type {Record<number, { hidden?: boolean }>} */ info) => {
    /** @type {number[]} */
    const out = [];
    for (let i = a; i <= b; i++) if (info[i]?.hidden !== true) out.push(i);
    if (out.length) return out;
    for (let i = a; i <= b; i++) out.push(i);
    return out;
  };
  /** @type {string[][]} */
  const table = [];
  /** Per table cell: a text cell (marked as text in the HTML). @type {boolean[][]} */
  const isText = [];
  const markText = G?.kind === 'sheet';
  /** @param {number} r @param {number[]} cols ascending @param {string[]} out @param {boolean[]} outText */
  const appendRow = (r, cols, out, outText) => {
    const row = rowOf(r);
    if (row === undefined) return false;
    const cells = row ? row.cells : NO_CELLS;
    let i = cols.length ? lowerBound(cells, cols[0]) : 0;
    for (const c of cols) {
      while (i < cells.length && cells[i].c < c) i++;
      let text = '';
      let str = false;
      if (i < cells.length && cells[i].c === c) {
        const cell = cells[i++];
        const covered = view.mergeIndex.size && mergeAt(view, r, c) >= 0 && !(masterOf(view, r, c).r === r && masterOf(view, r, c).c === c);
        if (!covered) {
          text = displayText(cell);
          str = markText && cell.t === 's' && text !== '';
        }
      }
      out.push(text);
      outText.push(str);
    }
    return true;
  };
  if (mode === 'cols') {
    const { r0, r1 } = ranges[0];
    const colsOf = ranges.map((rg) => shown(rg.c0, rg.c1, sm.cols));
    for (const r of shown(r0, r1, sm.rows)) {
      /** @type {string[]} */
      const out = [];
      /** @type {boolean[]} */
      const outText = [];
      for (const cols of colsOf) if (!appendRow(r, cols, out, outText)) return null;
      table.push(out);
      isText.push(outText);
    }
  } else {
    for (const rg of ranges) {
      const cols = shown(rg.c0, rg.c1, sm.cols);
      for (const r of shown(rg.r0, rg.r1, sm.rows)) {
        /** @type {string[]} */
        const out = [];
        /** @type {boolean[]} */
        const outText = [];
        if (!appendRow(r, cols, out, outText)) return null;
        table.push(out);
        isText.push(outText);
      }
    }
  }
  const quote = (/** @type {string} */ s) => (/[\t\r\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const text = table.map((row) => row.map(quote).join('\t')).join('\r\n');
  const td = (/** @type {string} */ s, /** @type {boolean} */ str) => `<td${str ? ` style="mso-number-format:'\\@'"` : ''}>${escapeHtml(s).replace(/\r?\n/g, '<br>')}</td>`;
  const html = '<meta charset="utf-8"><table>' + table.map((row, i) => '<tr>' + row.map((s, j) => td(s, isText[i][j])).join('') + '</tr>').join('') + '</table>';
  return { text, html, cells: table.length * (table[0]?.length ?? 0) };
}

/**
 * Fetches the rows of a large selection, then writes the clipboard asynchronously.
 * @param {{ view: SheetView, ranges: Range[], mode: CopyMode, cells: number }} plan
 */
async function copyAsync(plan) {
  if (!G) return;
  const grid = G;
  setStatusMessage('Copying…', 'info', 0);
  try {
    let r0 = Infinity;
    let r1 = -1;
    let c0 = Infinity;
    let c1 = -1;
    for (const rg of plan.ranges) {
      r0 = Math.min(r0, rg.r0);
      r1 = Math.max(r1, rg.r1);
      c0 = Math.min(c0, rg.c0);
      c1 = Math.max(c1, rg.c1);
    }
    // Only the copied columns are fetched (merged-cell coverage comes from the sheet meta, not from the rows).
    const rows = await fetchRows(plan.view.index, r0, r1, { c0, c1 });
    if (G !== grid) return;
    const payload = buildCopyPayload(plan.view, plan.ranges, plan.mode, (r) => rows.get(r) ?? null);
    if (payload) copyPayload(payload);
  } catch (err) {
    setStatusMessage(`Copy failed: ${err instanceof Error ? err.message : String(err)}`, 'warn');
  }
}

/** @param {string} text @param {number} cells */
async function writeClipboard(text, cells) {
  try {
    await navigator.clipboard.writeText(text);
    setStatusMessage(`Copied ${formatCount(cells)} cell${cells === 1 ? '' : 's'}`, 'info');
  } catch (err) {
    setStatusMessage('Copy failed: the clipboard is not available.', 'warn');
    logError('Clipboard write failed', err);
  }
}

// ----- CELL TEXT BOX -----
// Double-clicking a cell opens Excel's in-cell edit box, read-only: a textarea that refuses every edit, over the cell
// (or merged area), with the cell's display text, all of it selected. Part of it can be selected with the mouse or keyboard and
// copied: the textarea's own copy (plain text); the grid's copy handler only acts while the grid has the focus.
// Nothing can be typed, pasted, cut or dropped, and nothing is sent to the host. It lies outside the scroller, so a
// mirrored right-to-left sheet does not mirror its text, and rotated text is shown unrotated. It closes on Escape,
// Enter or Tab, a click elsewhere, scrolling, a sheet switch, zoom, invalidate, a theme change or a resize; the focus
// then goes back to the grid, whose selection is unchanged. (A double-click on a header border, Excel's autofit, is
// phase 3: headers ignore double-clicks.)

/**
 * @typedef {object} CellBox
 * @property {HTMLTextAreaElement} el
 * @property {number} sheet
 * @property {number} r          top-left cell of the cell / merged area
 * @property {number} c
 * @property {number} st         scroller position when it opened: a scroll event at that position keeps it open
 * @property {number} sl
 * @property {boolean} loaded    the cell's row was loaded (else onRows fills the text in when it arrives)
 * @property {boolean} menu      a context menu was opened on it and the focus has not come back since: VS Code's menu
 *   takes the focus away (to <body>), which must not close the box (its Copy copies the box's selected text)
 */

/** @param {MouseEvent} e */
function onGridDblClick(e) {
  e.preventDefault();
  if (!G || e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;
  const sc = G.dom.scroller;
  const rect = sc.getBoundingClientRect();
  const vx = scrollerX(rect, e.clientX);
  const vy = e.clientY - rect.top;
  if (vx >= sc.clientWidth || vy >= sc.clientHeight || vx < 0 || vy < 0) return; // scrollbars
  const hit = hitTest(vx, vy);
  if (hit.area !== 'cell') return;
  const view = activeView();
  const m = masterOf(view, hit.r, hit.c);
  openCellBox(view, m.r, m.c);
}

/**
 * Opens (or re-opens) the read-only text box of a cell: all of its text selected (double-click), or the caret at the
 * end of the text (F2, as Excel's in-cell edit puts it).
 * @param {SheetView} view @param {number} r @param {number} c  top-left cell of the cell / merged area
 * @param {boolean} [caretAtEnd]
 */
function openCellBox(view, r, c, caretAtEnd = false) {
  const g = /** @type {GridState} */ (G);
  closeCellBox(false);
  const host = g.dom.scroller.parentElement;
  if (!view.L || !host) return;
  hoverReset();
  closeMenu();
  scrollCellIntoView(view, r, c);
  const cell = cellAt(view, r, c);
  const text = cell ? displayText(cell) : '';
  // Not the readonly attribute: Chrome then ignores the keys that move the caret (Shift+Arrow, Home, End), so the
  // text could not be selected from the keyboard. Every edit is cancelled instead (beforeinput, paste, cut, drop); an
  // IME composition, which cannot be cancelled, is undone by the input handler.
  const el = /** @type {HTMLTextAreaElement} */ (
    h('textarea', {
      class: 'fv-cellbox',
      'aria-readonly': 'true',
      spellcheck: 'false',
      autocomplete: 'off',
      'aria-label': `${colName(c)}${r + 1} cell text, read-only`,
      onkeydown: (/** @type {KeyboardEvent} */ e) => {
        if (G?.cellBox?.el === el) G.cellBox.menu = false;
        if (e.isComposing || (e.key !== 'Escape' && e.key !== 'Enter' && e.key !== 'Tab')) return;
        e.preventDefault();
        e.stopPropagation();
        closeCellBox();
      },
      // Focus moving elsewhere in the view (F6, Ctrl+G, a click on the name box) closes it; switching to another
      // window keeps it, and the browser gives it the focus back. So does VS Code's context menu on it.
      onblur: () => {
        if (G?.cellBox?.el === el && document.hasFocus() && !G.cellBox.menu) closeCellBox(false);
      },
      oncontextmenu: () => {
        if (G?.cellBox?.el === el) G.cellBox.menu = true;
      },
      onwheel: (/** @type {WheelEvent} */ e) => onCellBoxWheel(e, el),
      onbeforeinput: (/** @type {Event} */ e) => e.preventDefault(),
      oninput: () => {
        if (el.value === text) return;
        const at = Math.min(el.selectionStart, text.length);
        el.value = text;
        el.setSelectionRange(at, at);
      },
      onpaste: (/** @type {Event} */ e) => e.preventDefault(),
      oncut: (/** @type {Event} */ e) => e.preventDefault(),
      ondrop: (/** @type {Event} */ e) => e.preventDefault(),
      ondragstart: (/** @type {Event} */ e) => e.preventDefault(),
    })
  );
  el.value = text;
  const sc = g.dom.scroller;
  g.cellBox = { el, sheet: view.index, r, c, st: sc.scrollTop, sl: sc.scrollLeft, loaded: cell !== undefined, menu: false };
  host.appendChild(el);
  layoutCellBox(view, el, r, c, cell);
  el.focus({ preventScroll: true });
  el.setSelectionRange(caretAtEnd ? el.value.length : 0, el.value.length);
  el.scrollTop = caretAtEnd ? el.scrollHeight : 0;
}

/**
 * Closes the cell text box. With `focusGrid` the grid takes the focus back when the box had it (or the focus was lost).
 * @param {boolean} [focusGrid]
 */
function closeCellBox(focusGrid = true) {
  const box = G?.cellBox;
  if (!G || !box) return;
  // Cleared first: removing the focused box fires its blur handler, which must find it closed.
  G.cellBox = null;
  const hadFocus = document.activeElement === box.el;
  box.el.remove();
  if (focusGrid && (hadFocus || document.activeElement === document.body) && G.dom.scroller.isConnected) G.dom.scroller.focus({ preventScroll: true });
}

/**
 * Places the box over the cell in the cell's font, colours and alignment, grown to fit the text like Excel's edit box:
 * wider in the direction the text runs (to the right for left-aligned text, to the left for right-aligned, both ways
 * for centred; a wrapped cell keeps its width), taller for more lines, and never past the visible grid area (there
 * the text wraps and the box scrolls).
 * @param {SheetView} view @param {HTMLTextAreaElement} el @param {number} r @param {number} c
 * @param {CellData | null | undefined} cell
 */
function layoutCellBox(view, el, r, c, cell) {
  const g = /** @type {GridState} */ (G);
  const text = el.value;
  const L = /** @type {Layout} */ (view.L);
  const sc = g.dom.scroller;
  const host = /** @type {HTMLElement} */ (sc.parentElement);
  const s = cell ? cell.s || 0 : rowStyleIndex(view.meta, r) || L.colStyle[c] || 0;
  const S = styleInfo(s);
  const st = g.styles[s] || {};
  const fmtColor = cell && cell.t !== 'z' && !(cell.rt && cell.rt.length) ? numFmtColor(st.numFmt, cell) : undefined;
  const color = cellFontColor(fmtColor ?? st.font?.color, S.hasFill);
  // Alignment of the unrotated text (justify / distributed / fill read from the left).
  const hText = cellAlign(S.rot ? { ...S, rot: 0 } : S, cell ?? null);
  const align = hText === 'center' || hText === 'right' ? hText : 'left';

  // Scroller layout coordinates (before right-to-left mirroring) of the cell and of the visible area of its pane.
  const area = cellArea(view, r, c);
  const cEnd = Math.min(area.c1, L.nCols - 1);
  const rEnd = Math.min(area.r1, L.nRows - 1);
  const dx = area.c0 >= L.fc ? sc.scrollLeft : 0;
  const dy = area.r0 >= L.fr ? virtualScrollTop() : 0;
  const minX = area.c0 >= L.fc ? L.B : L.hdrW;
  const minY = area.r0 >= L.fr ? L.A : L.hdrH;
  const maxX = Math.max(minX, sc.clientWidth);
  const maxY = Math.max(minY, sc.clientHeight);
  const x0 = clamp(L.hdrW + L.colLeft[area.c0] - dx, minX, maxX);
  const x1 = clamp(L.hdrW + L.colLeft[cEnd + 1] - dx, x0, maxX);
  const y0 = clamp(L.hdrH + L.rowTop[area.r0] - dy, minY, maxY);
  const y1 = clamp(L.hdrH + L.rowTop[rEnd + 1] - dy, y0, maxY);
  // Host-relative screen coordinates.
  const rect = sc.getBoundingClientRect();
  const hr = host.getBoundingClientRect();
  const rtl = L.rtl;
  const vis = (/** @type {number} */ x) => (rtl ? rect.right - hr.left - x : rect.left - hr.left + x);
  const cellL = Math.min(vis(x0), vis(x1));
  const cellR = Math.max(vis(x0), vis(x1));
  const areaL = Math.min(vis(minX), vis(maxX));
  const areaR = Math.max(vis(minX), vis(maxX));
  const cellTop = rect.top - hr.top + y0;
  const areaT = rect.top - hr.top + minY;
  const areaB = rect.top - hr.top + maxY;
  const cellW = cellR - cellL;

  // 2px padding on each side (the accent frame is an outline outside the box); the widest line is measured like the
  // grid measures text (canvas), plus 2px for rounding.
  const frame = 4;
  // At least room for two characters (narrow columns).
  let width = Math.max(cellW, frame + Math.ceil(S.fontPx * 2));
  if (!S.wrap && text) {
    let widest = 0;
    for (const line of text.split('\n')) widest = Math.max(widest, measureText(S.font, line));
    const room = align === 'left' ? areaR - cellL : align === 'right' ? cellR - areaL : areaR - areaL;
    width = Math.min(Math.max(width, Math.ceil(widest) + frame + 2), Math.max(width, room));
  }
  const left = align === 'left' ? cellL : align === 'right' ? cellR - width : clamp((cellL + cellR - width) / 2, areaL, Math.max(areaL, areaR - width));
  const f = st.font || {};
  // Measured without a scrollbar (it would narrow the text while the height is 0).
  el.style.cssText =
    `left:${left}px;top:${cellTop}px;width:${width}px;height:0;overflow:hidden;` +
    `font:${S.italic ? 'italic ' : ''}${S.bold ? '700 ' : ''}${S.fontPx}px ${S.family};line-height:1.2;text-align:${align};` +
    (color ? `color:${color};` : '') +
    (S.bg ? `background-color:${S.bg};` : '') +
    textDecorationCss(f);
  // Taller for more lines (one layout read): down from the cell, moved up when there is no room below, never past the
  // visible grid (then the box scrolls).
  const want = Math.max(el.scrollHeight, y1 - y0, Math.ceil(S.fontPx * 1.2));
  const top = cellTop + want > areaB ? Math.max(areaT, areaB - want) : cellTop;
  el.style.top = `${top}px`;
  el.style.height = `${Math.max(y1 - y0, Math.min(want, areaB - top))}px`;
  el.style.overflow = '';
}

/**
 * Wheel over the cell text box: Ctrl zooms like over the grid; a wheel the box cannot use for its own text scrolls
 * the grid (which closes the box).
 * @param {WheelEvent} e @param {HTMLTextAreaElement} el
 */
function onCellBoxWheel(e, el) {
  if (!G) return;
  if (e.ctrlKey) {
    onGridWheel(e);
    return;
  }
  const down = e.deltaY > 0;
  const own = e.deltaY !== 0 && (down ? el.scrollTop + el.clientHeight < el.scrollHeight - 1 : el.scrollTop > 0);
  if (own) return;
  e.preventDefault();
  const sc = G.dom.scroller;
  const unit = e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? G.viewH : 1;
  closeCellBox();
  sc.scrollTop += e.deltaY * unit;
  sc.scrollLeft += (isRtl() ? -1 : 1) * e.deltaX * unit;
}

// ----- LINKS & NAVIGATION -----

/**
 * Internal link target (`#Sheet!A1`, `#'My Sheet'!A1:B2`, `Sheet2!B3`) or null for external links. Links into
 * another workbook (`other.xlsx#Sheet1!A1`, HYPERLINK's `[other.xlsx]Sheet1!A1`) are external: the host opens the file.
 * @param {string} href
 * @returns {{ sheet?: string, ref: string } | null}
 */
function parseInternalLink(href) {
  let s = href.trim();
  if (s.startsWith('#')) s = s.slice(1);
  else if (/^[a-z][a-z0-9+.-]*:/i.test(s) || s.includes('/') || s.includes('\\') || s.includes('#') || workbookLink(s)) return null;
  const bang = s.lastIndexOf('!');
  if (bang < 0) return /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(s) ? { ref: s } : href.startsWith('#') ? { ref: s } : null;
  let sheet = s.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, ref: s.slice(bang + 1) };
}

/**
 * Excel's workbook reference in a HYPERLINK target, `[Book.xlsx]Sheet1!A1` or `'[Book.xlsx]My Sheet'!A1`, as the
 * file path and the location inside it (`Sheet1!A1`); null for anything else.
 * @param {string} s
 */
function workbookLink(s) {
  const m = /^'?\[([^\]]+)\](.*)$/.exec(s.trim());
  if (!m) return null;
  let location = m[2];
  // '[Book]Sheet name'!A1 -> 'Sheet name'!A1
  if (s.trim().startsWith("'")) location = location.replace(/^(.*)'(!.*)?$/, (_, name, ref) => `'${name}'${ref ?? ''}`);
  return { file: m[1], location };
}

/** Ctrl/Cmd+Click (Ctrl+Enter) on a hyperlink. @param {string} href */
function openLink(href) {
  const internal = parseInternalLink(href);
  if (internal) {
    navigateTo(internal.sheet, internal.ref);
    return;
  }
  // The host resolves `file#location` against the document folder and opens the file.
  const book = workbookLink(href);
  post({ type: 'openLink', href: book ? `${book.file}${book.location ? '#' + book.location : ''}` : href });
}

/** What a hyperlink leads to, for the tooltip and screen readers. @param {string} href */
function linkTarget(href) {
  return parseInternalLink(href) ? `Go to ${href.trim().replace(/^#/, '')}` : href;
}

/**
 * Selects a reference, optionally on another sheet (hidden sheets are revealed for viewing).
 * @param {string | undefined} sheetName @param {string} ref
 * @returns {boolean}
 */
function navigateTo(sheetName, ref) {
  if (!G) return false;
  let index = G.sheet;
  if (sheetName !== undefined) {
    const wanted = sheetName.toLowerCase();
    index = G.meta.sheets.findIndex((s) => s.name.toLowerCase() === wanted);
    if (index < 0) {
      setStatusMessage(`Sheet "${sheetName}" was not found.`, 'warn');
      return false;
    }
  }
  const target = getView(index);
  const rg = parseRangeRef(ref, target);
  if (!rg) {
    setStatusMessage(`"${ref}" is not a valid reference.`, 'warn');
    return false;
  }
  // Excel refuses references into hidden sheets ("Reference isn't valid"); this viewer shows the sheet read-only, as
  // its hidden-sheets menu does, and says so.
  const sm = G.meta.sheets[index];
  const reveals = sm.state !== 'visible' && !G.revealed.has(index);
  if (index !== G.sheet) activateSheet(index);
  const view = activeView();
  // A reference past the used range grows the layout to it (instead of stopping at its last row).
  growLayout(view, rg.r0, rg.c0);
  const L = /** @type {Layout} */ (view.L);
  const active = masterOf(view, visibleRow(L, Math.min(rg.r0, L.nRows - 1)), visibleCol(L, Math.min(rg.c0, L.nCols - 1)));
  const snapped = snapRange(view, rg);
  setSelection(view, [snapped], active);
  if (reveals) setStatusMessage(`Showing the ${hiddenSheetLabel(sm)} "${sm.name}" (Excel does not go to hidden sheets).`, 'info');
  G.dom.scroller.focus({ preventScroll: true });
  return true;
}

/**
 * Parses A1, A1:C5, $A$1, A:C, 3:5 into a range.
 * @param {string} ref @param {SheetView} view
 * @returns {Range | null}
 */
function parseRangeRef(ref, view) {
  const s = ref.replace(/\$/g, '').trim().toUpperCase();
  let m = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(s);
  if (m) {
    const c0 = colIndex(m[1]);
    const r0 = Number(m[2]) - 1;
    const c1 = m[3] ? colIndex(m[3]) : c0;
    const r1 = m[4] ? Number(m[4]) - 1 : r0;
    if (c0 < 0 || c1 < 0 || r0 < 0 || r1 < 0 || Math.max(r0, r1) >= view.maxR || Math.max(c0, c1) >= view.maxC) return null;
    return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
  }
  m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s);
  if (m) {
    const c0 = colIndex(m[1]);
    const c1 = colIndex(m[2]);
    if (c0 < 0 || c1 < 0 || Math.max(c0, c1) >= view.maxC) return null;
    return { r0: 0, r1: view.maxR - 1, c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
  }
  m = /^(\d{1,7}):(\d{1,7})$/.exec(s);
  if (m) {
    const r0 = Number(m[1]) - 1;
    const r1 = Number(m[2]) - 1;
    if (r0 < 0 || r1 < 0 || Math.max(r0, r1) >= view.maxR) return null;
    return { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: 0, c1: view.maxC - 1 };
  }
  return null;
}

// ----- VIEW STATE -----

/** Snapshot of per-sheet view state (persisted with vscode.setState). */
function snapshotState() {
  const g = /** @type {GridState} */ (G);
  /** @type {Record<string, any>} */
  const sheets = { ...(g.saved?.sheets || {}) };
  for (const v of g.views.values()) {
    sheets[v.meta.name] = {
      st: Math.round(v.scrollTop),
      sl: Math.round(v.scrollLeft),
      z: v.zoom,
      sel: { ranges: v.sel.ranges.slice(0, 64), active: v.sel.active, anchor: v.sel.anchor, cursor: v.sel.cursor },
    };
  }
  // Hidden sheets revealed for viewing (one of them may be the active sheet). While the workbook has no worksheets
  // (sheet -1 after an invalidate), the last active / revealed sheets are kept for when worksheets come back.
  if (g.sheet < 0 && g.saved) return { file: g.fileName, activeSheet: g.saved.activeSheet, revealed: g.saved.revealed, sheets };
  const revealed = [...g.revealed].map((i) => g.meta.sheets[i]?.name).filter((name) => name !== undefined);
  return { file: g.fileName, activeSheet: g.meta.sheets[g.sheet]?.name, revealed, sheets };
}

const saveGridState = debounce(() => {
  if (!G || G.sheet < 0) return;
  G.saved = snapshotState();
  setStateKey('grid', G.saved);
}, 400);

/** Phase 3 hook: sheet tab context menu (rename, insert, delete, move, colour). @param {number} _index @param {MouseEvent} _e */
function onTabContextMenu(_index, _e) {
  // Read-only in phase 1: the default webview context menu is kept.
}

// ===== SHEET TABS =====

function buildTabBar() {
  const prev = /** @type {HTMLButtonElement} */ (
    h('button', { class: 'fv-icon-btn fv-tab-scroll', type: 'button', title: 'Scroll tabs left', 'aria-label': 'Scroll tabs left', onclick: () => scrollTabs(-1) }, icon(ICONS.chevronLeft))
  );
  const next = /** @type {HTMLButtonElement} */ (
    h('button', { class: 'fv-icon-btn fv-tab-scroll', type: 'button', title: 'Scroll tabs right', 'aria-label': 'Scroll tabs right', onclick: () => scrollTabs(1) }, icon(ICONS.chevronRight))
  );
  const nav = h('div', { class: 'fv-tab-nav' }, prev, next);
  const strip = h('div', { class: 'fv-tabs', role: 'tablist', 'aria-label': 'Sheets' });
  strip.addEventListener('scroll', () => updateTabArrows(), { passive: true });
  strip.addEventListener(
    'wheel',
    (e) => {
      // A vertical wheel scrolls towards the later tabs (to the left in a right-to-left bar).
      const delta = Math.abs(e.deltaX) > Math.abs(e.deltaY) ? e.deltaX : tabsRtl() ? -e.deltaY : e.deltaY;
      if (!delta || strip.scrollWidth <= strip.clientWidth) return;
      e.preventDefault();
      strip.scrollLeft += delta;
    },
    { passive: false },
  );
  strip.addEventListener('keydown', onTabKeyDown);
  const hiddenBtn = /** @type {HTMLButtonElement} */ (
    h(
      'button',
      {
        class: 'fv-icon-btn fv-hidden-btn',
        type: 'button',
        title: 'Hidden sheets',
        'aria-label': 'Hidden sheets',
        'aria-haspopup': 'menu',
        'aria-expanded': 'false',
        hidden: true,
        onclick: () => toggleHiddenMenu(),
        // Menu button keys: Down / Up open the menu on its first / last item.
        onkeydown: (/** @type {KeyboardEvent} */ e) => {
          if ((e.key !== 'ArrowDown' && e.key !== 'ArrowUp') || e.ctrlKey || e.altKey || e.metaKey || G?.dom.menu) return;
          e.preventDefault();
          toggleHiddenMenu(e.key === 'ArrowUp');
        },
      },
      icon(ICONS.eyeClosed),
    )
  );
  const bar = h('div', { class: 'fv-tabbar' }, nav, strip, hiddenBtn);
  return { bar, nav, strip, prev, next, hiddenBtn };
}

/** Rebuilds the tab strip (visible sheets + hidden sheets revealed for viewing). */
function renderTabs() {
  if (!G) return;
  const d = G.dom;
  /** @type {HTMLElement[]} */
  const tabs = [];
  let hiddenCount = 0;
  for (const sm of G.meta.sheets) {
    if (sm.state !== 'visible') hiddenCount++;
    const revealed = sm.state !== 'visible' && G.revealed.has(sm.index);
    if (sm.state !== 'visible' && !revealed) continue;
    const active = sm.index === G.sheet;
    const tab = h(
      'button',
      {
        class: `fv-tab${active ? ' fv-tab-active' : ''}${revealed ? ' fv-tab-hidden' : ''}${sm.tabColor ? ' fv-tab-colored' : ''}`,
        type: 'button',
        role: 'tab',
        'aria-selected': String(active),
        tabindex: active ? '0' : '-1',
        title: revealed ? `${sm.name} (${hiddenSheetLabel(sm)})` : sm.name,
        // Screen readers hear what the italic label and the dashed mark show.
        'aria-label': revealed ? `${sm.name}, ${hiddenSheetLabel(sm)}` : null,
        dataset: { index: String(sm.index) },
        onclick: () => {
          activateSheet(sm.index);
          G?.dom.scroller.focus({ preventScroll: true });
        },
        oncontextmenu: (/** @type {MouseEvent} */ e) => onTabContextMenu(sm.index, e),
      },
      h('span', { class: 'fv-tab-label' }, sm.name),
    );
    if (sm.tabColor) tab.style.setProperty('--fv-tab-color', sm.tabColor);
    tabs.push(tab);
  }
  d.tabStrip.replaceChildren(...tabs);
  // Like Excel, a right-to-left sheet mirrors the tab bar: tabs start at the right (the formula bar stays as is).
  d.tabBar.dir = isRtl() ? 'rtl' : 'ltr';
  // Every hidden sheet of the workbook counts, also those shown as tabs: the menu lists them all.
  d.hiddenBtn.hidden = hiddenCount === 0;
  d.hiddenBtn.title = `Hidden sheets (${hiddenCount})`;
  d.hiddenBtn.setAttribute('aria-label', d.hiddenBtn.title);
  updateTabArrows();
  const activeTab = /** @type {HTMLElement | null} */ (d.tabStrip.querySelector('.fv-tab-active'));
  if (activeTab) ensureTabVisible(activeTab);
}

/** 'hidden sheet' / 'very hidden sheet' (tab, menu and status labels). @param {SheetMeta} sm */
function hiddenSheetLabel(sm) {
  return sm.state === 'veryHidden' ? 'very hidden sheet' : 'hidden sheet';
}

/** The tab bar is mirrored (right-to-left sheet active, see renderTabs). */
function tabsRtl() {
  return !!G && G.dom.tabBar.dir === 'rtl';
}

/** Scrolls the strip so the tab is fully visible (screen geometry: works in both directions). @param {HTMLElement} tab */
function ensureTabVisible(tab) {
  if (!G) return;
  const strip = G.dom.tabStrip;
  const t = tab.getBoundingClientRect();
  const s = strip.getBoundingClientRect();
  if (t.left < s.left) strip.scrollLeft -= s.left - t.left;
  else if (t.right > s.right) strip.scrollLeft += t.right - s.right;
}

/** @param {number} dir */
function scrollTabs(dir) {
  if (!G) return;
  const strip = G.dom.tabStrip;
  strip.scrollBy({ left: dir * Math.max(80, strip.clientWidth * 0.6), behavior: reduceMotion() ? 'auto' : 'smooth' });
}

/**
 * Animations are off: the OS setting (media query) or VS Code's "Reduce Motion" (workbench.reduceMotion), which
 * reaches the webview only as the body class vscode-reduce-motion.
 */
function reduceMotion() {
  return document.body.classList.contains('vscode-reduce-motion') || window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

function updateTabArrows() {
  if (!G) return;
  const d = G.dom;
  const strip = d.tabStrip;
  const overflow = strip.scrollWidth > strip.clientWidth + 1;
  d.tabNav.classList.toggle('fv-tab-nav-active', overflow);
  // x = distance scrolled from the far left (scrollLeft runs from 0 to negative values in a right-to-left strip);
  // 1px tolerance at both ends for fractional tab widths.
  const max = strip.scrollWidth - strip.clientWidth;
  const x = tabsRtl() ? strip.scrollLeft + max : strip.scrollLeft;
  d.tabPrev.disabled = !overflow || x <= 1;
  d.tabNext.disabled = !overflow || x >= max - 1;
}

/** Arrow keys move between tabs (roving tabindex). @param {KeyboardEvent} e */
function onTabKeyDown(e) {
  if (!G) return;
  const tabs = /** @type {HTMLElement[]} */ ([...G.dom.tabStrip.querySelectorAll('.fv-tab')]);
  const pos = tabs.indexOf(/** @type {HTMLElement} */ (document.activeElement));
  if (pos < 0) return;
  let next = -1;
  const rtl = tabsRtl();
  if (e.key === (rtl ? 'ArrowLeft' : 'ArrowRight')) next = Math.min(tabs.length - 1, pos + 1);
  else if (e.key === (rtl ? 'ArrowRight' : 'ArrowLeft')) next = Math.max(0, pos - 1);
  else if (e.key === 'Home') next = 0;
  else if (e.key === 'End') next = tabs.length - 1;
  if (next < 0) return;
  e.preventDefault();
  activateSheet(Number(tabs[next].dataset.index));
  const tab = /** @type {HTMLElement | null} */ (G.dom.tabStrip.querySelector('.fv-tab-active'));
  tab?.focus();
}

/**
 * Menu listing every hidden sheet of the workbook; choosing one shows it read-only as a tab (revealed for viewing, the
 * file is not changed). Sheets already shown say so, and a last item hides them again. Very hidden sheets (Excel's
 * Unhide dialog leaves them out: only VBA shows them) are labelled.
 * @param {boolean} [focusLast] focus the last item (Up key on the button) instead of the first
 */
function toggleHiddenMenu(focusLast = false) {
  if (!G) return;
  if (G.dom.menu) {
    closeMenu();
    return;
  }
  const g = G;
  const hidden = g.meta.sheets.filter((s) => s.state !== 'visible');
  if (!hidden.length) return;
  const items = hidden.map((s) => {
    const shown = g.revealed.has(s.index);
    const very = s.state === 'veryHidden';
    const hint = [very ? 'very hidden' : '', shown ? 'shown' : ''].filter(Boolean).join(' · ');
    return h(
      'button',
      {
        class: 'fv-menu-item',
        type: 'button',
        role: 'menuitem',
        'aria-label': `${s.name}, ${hiddenSheetLabel(s)}${very ? ' (Excel shows it only through VBA)' : ''}${shown ? ', shown' : ''}`,
        title: very ? `${s.name}: very hidden sheet. Excel's Unhide dialog does not list it; only VBA can show it.` : `${s.name}: hidden sheet`,
        onclick: () => {
          closeMenu();
          activateSheet(s.index);
          G?.dom.scroller.focus({ preventScroll: true });
        },
      },
      h('span', { class: 'fv-menu-label' }, s.name),
      hint ? h('span', { class: 'fv-menu-hint' }, hint) : null,
    );
  });
  if (hidden.some((s) => g.revealed.has(s.index))) {
    items.push(
      h('div', { class: 'fv-menu-sep', role: 'separator' }),
      h(
        'button',
        {
          class: 'fv-menu-item',
          type: 'button',
          role: 'menuitem',
          onclick: () => {
            closeMenu();
            hideRevealedSheets();
            G?.dom.scroller.focus({ preventScroll: true });
          },
        },
        h('span', { class: 'fv-menu-label' }, 'Hide the shown sheets again'),
      ),
    );
  }
  // The title repeats the menu's name: hidden from screen readers (a menu holds menu items only).
  const menu = h('div', { class: 'fv-menu', role: 'menu', 'aria-label': `Hidden sheets (${hidden.length})` }, h('div', { class: 'fv-menu-title', 'aria-hidden': 'true' }, 'Hidden sheets'), ...items);
  menu.addEventListener('keydown', (e) => {
    const list = /** @type {HTMLElement[]} */ ([...menu.querySelectorAll('.fv-menu-item')]);
    const pos = list.indexOf(/** @type {HTMLElement} */ (document.activeElement));
    if (e.key === 'Escape' || e.key === 'Tab') {
      // The menu is one stop: Escape and Tab close it and leave the focus on its button.
      e.preventDefault();
      closeMenu();
      G?.dom.hiddenBtn.focus();
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      list[(pos + (e.key === 'ArrowDown' ? 1 : list.length - 1)) % list.length]?.focus();
    } else if (e.key === 'Home' || e.key === 'End') {
      e.preventDefault();
      list[e.key === 'Home' ? 0 : list.length - 1]?.focus();
    }
  });
  // Focus moving anywhere else (a click, the grid taking focus) closes the menu.
  menu.addEventListener('focusout', (e) => {
    const to = /** @type {Node | null} */ (e.relatedTarget);
    if (G?.dom.menu === menu && !(to && menu.contains(to))) closeMenu();
  });
  G.dom.menu = menu;
  root.appendChild(menu);
  const btn = G.dom.hiddenBtn.getBoundingClientRect();
  const mw = menu.offsetWidth;
  menu.style.left = `${clamp(tabsRtl() ? btn.left : btn.right - mw, 4, Math.max(4, window.innerWidth - mw - 4))}px`;
  menu.style.bottom = `${Math.max(4, window.innerHeight - btn.top + 2)}px`;
  G.dom.hiddenBtn.setAttribute('aria-expanded', 'true');
  const list = menu.querySelectorAll('.fv-menu-item');
  /** @type {HTMLElement | undefined} */ (list[focusLast ? list.length - 1 : 0])?.focus();
}

/**
 * Hides the sheets shown for viewing again (their tabs go). When the active sheet is one of them, the next visible
 * sheet becomes active, else the previous one (Excel's Hide); a workbook without visible sheets keeps its sheet.
 */
function hideRevealedSheets() {
  if (!G) return;
  const sheets = G.meta.sheets;
  const active = G.sheet;
  G.revealed.clear();
  if (sheets[active].state !== 'visible') {
    const next = sheets.find((s) => s.index > active && s.state === 'visible') ?? [...sheets].reverse().find((s) => s.state === 'visible');
    if (next) {
      activateSheet(next.index);
      return;
    }
    G.revealed.add(active);
  }
  renderTabs();
  saveGridState();
}

function closeMenu() {
  if (!G || !G.dom.menu) return;
  // Cleared first: removing the focused menu fires its focusout handler, which must find the menu closed.
  const menu = G.dom.menu;
  G.dom.menu = null;
  G.dom.hiddenBtn.setAttribute('aria-expanded', 'false');
  menu.remove();
}

// ===== FORMULA BAR =====

function buildFormulaBar() {
  const nameBox = /** @type {HTMLInputElement} */ (
    h('input', { class: 'fv-namebox', type: 'text', spellcheck: 'false', autocomplete: 'off', 'aria-label': 'Name Box', title: 'Name Box: type a reference such as B12, A1:C5 or Sheet2!A1 and press Enter' })
  );
  nameBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const text = nameBox.value.trim();
      const bang = text.lastIndexOf('!');
      let sheet;
      let ref = text;
      if (bang >= 0) {
        sheet = text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
        ref = text.slice(bang + 1);
      }
      if (!navigateTo(sheet, ref)) {
        nameBox.classList.add('fv-invalid');
        setTimeout(() => nameBox.classList.remove('fv-invalid'), 600);
        nameBox.select();
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      updateFormulaBar();
      G?.dom.scroller.focus({ preventScroll: true });
    }
  });
  nameBox.addEventListener('focus', () => nameBox.select());
  nameBox.addEventListener('blur', () => updateFormulaBar());
  // Read-only text box: focusable so keyboard and screen-reader users can read (and select) the whole value;
  // Escape / Enter go back to the grid like Excel.
  const formula = h('div', { class: 'fv-formula', role: 'textbox', tabindex: '0', 'aria-readonly': 'true', 'aria-label': 'Formula Bar' });
  formula.addEventListener('keydown', (e) => {
    if ((e.key === 'Escape' || e.key === 'Enter') && !e.ctrlKey && !e.altKey && !e.metaKey) {
      e.preventDefault();
      G?.dom.scroller.focus({ preventScroll: true });
    } else if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'a') {
      // Select all = the formula text, as in a text box (the page's select-all would take the whole view). Not
      // passed on: VS Code would run its own select-all on the page.
      e.preventDefault();
      e.stopPropagation();
      window.getSelection()?.selectAllChildren(formula);
    }
  });
  const bar = h(
    'div',
    { class: 'fv-formula-bar' },
    nameBox,
    h('div', { class: 'fv-fx-sep', 'aria-hidden': 'true' }),
    h('span', { class: 'fv-fx', 'aria-hidden': 'true', title: 'Cell value or formula (read-only)' }, 'fx'),
    formula,
  );
  return { bar, nameBox, formula };
}

/** Name box = address of the selection, formula bar = raw value or =formula of the active cell. */
function updateFormulaBar() {
  if (!G || G.sheet < 0) return;
  const view = activeView();
  const d = G.dom;
  if (document.activeElement !== d.nameBox) d.nameBox.value = selectionLabel(view);
  const a = view.sel.active;
  const cell = cellAt(view, a.r, a.c);
  const text = cell === undefined ? '' : rawValue(cell);
  if (d.formula.textContent !== text) {
    d.formula.textContent = text;
    d.formula.title = text.length > 60 || text.includes('\n') ? text : '';
  }
  d.formula.classList.toggle('fv-formula-f', !!(cell && cell.f));
}

// ===== STATUS BAR =====

/** @param {GridInitMessage} msg */
function buildStatusBar(msg) {
  const info = h('span', { class: 'fv-status-item fv-status-dims' });
  const mode = h('span', { class: 'fv-status-item fv-status-mode', title: msg.readOnly !== false ? 'This view is read-only' : '' }, msg.readOnly !== false ? 'Read-only' : 'Ready');
  const msgEl = h('span', { class: 'fv-status-item fv-status-msg', role: 'status', 'aria-live': 'polite' });
  const stats = h('span', { class: 'fv-status-item fv-status-stats', 'aria-live': 'polite' });
  const zoomLabel = /** @type {HTMLButtonElement} */ (h('button', { class: 'fv-zoom-label', type: 'button', title: 'Reset zoom to the sheet setting', onclick: () => resetZoom() }, '100%'));
  const zoom = h(
    'span',
    { class: 'fv-status-item fv-zoom', role: 'group', 'aria-label': 'Zoom' },
    h('button', { class: 'fv-icon-btn fv-zoom-btn', type: 'button', title: 'Zoom out (Ctrl+Wheel)', 'aria-label': 'Zoom out', onclick: () => G && setZoom(activeView().zoom - 10) }, '−'),
    zoomLabel,
    h('button', { class: 'fv-icon-btn fv-zoom-btn', type: 'button', title: 'Zoom in (Ctrl+Wheel)', 'aria-label': 'Zoom in', onclick: () => G && setZoom(activeView().zoom + 10) }, '+'),
  );
  const delimiter = msg.kind === 'csv' && msg.delimited ? buildDelimiterStatus(msg.delimited) : null;
  const bar = h('div', { class: 'fv-statusbar' }, mode, info, ...(delimiter ? [delimiter] : []), msgEl, h('span', { class: 'fv-status-spacer' }), stats, zoom);
  return { bar, msg: msgEl, stats, info, zoomLabel };
}

/** Names of the delimiters parseDelimited uses (' ' = runs of spaces). */
const DELIMITER_NAMES = /** @type {Record<string, string>} */ ({ ',': 'Comma', ';': 'Semicolon', '\t': 'Tab', '|': 'Pipe', ' ': 'Spaces' });

/**
 * Delimited text: the format and the delimiter the columns were split on ("CSV · Semicolon"), so a wrong guess (a tie,
 * a .csv that is really tab-separated) can be seen.
 * @param {{ format: string, delimiter: string }} delimited
 */
function buildDelimiterStatus(delimited) {
  const d = delimited.delimiter;
  const name = DELIMITER_NAMES[d] ?? `“${d}”`;
  const shown = d === ' ' ? 'runs of spaces' : d === '\t' ? 'tabs' : `“${d}”`;
  return h(
    'span',
    { class: 'fv-status-item fv-status-delim', title: `${delimited.format.toUpperCase()} file: the columns are split on ${shown} (FileStudio: Change Delimiter to choose another)` },
    `${delimited.format.toUpperCase()} · ${name}`,
  );
}

function updateStatusInfo() {
  if (!G) return;
  const sm = G.meta.sheets[G.sheet];
  const hidden = sm.state !== 'visible' ? ` · ${hiddenSheetLabel(sm)}` : '';
  G.dom.statusInfo.textContent = `${formatCount(sm.rowCount)} × ${formatCount(sm.colCount)}${hidden}`;
  G.dom.statusInfo.title = `${sm.name}: ${formatCount(sm.rowCount)} rows × ${formatCount(sm.colCount)} columns in the used range`;
}

function updateZoomLabel() {
  if (!G) return;
  G.dom.zoomLabel.textContent = `${activeView().zoom}%`;
}

function resetZoom() {
  if (!G) return;
  const view = activeView();
  setZoom(clamp(view.meta.zoom || 100, MIN_ZOOM, MAX_ZOOM));
}

/** @type {ReturnType<typeof setTimeout> | undefined} */
let statusMsgTimer;

/**
 * Transient message in the status bar (copy feedback, navigation errors).
 * @param {string} text @param {'info' | 'warn'} kind @param {number} [ms] 0 = until replaced
 */
function setStatusMessage(text, kind, ms = 4000) {
  if (!G) return;
  const el = G.dom.statusMsg;
  el.textContent = text;
  el.className = `fv-status-item fv-status-msg fv-status-${kind}`;
  if (statusMsgTimer !== undefined) clearTimeout(statusMsgTimer);
  statusMsgTimer = ms ? setTimeout(() => G && (G.dom.statusMsg.textContent = ''), ms) : undefined;
}

/**
 * Selection statistics being gathered: the getStats requests not answered yet, the sum of the answers so far, how many
 * requests were sent and whether ranges were left out (more hidden-row gaps than the requests carry).
 * @typedef {{ ids: Set<number>, stats: SelectionStats, requests: number, cut: boolean }} StatsRequest
 */

/** Most ranges one getStats request carries (the host reads the first 1,024: MAX_STATS_RANGES in viewerProvider.ts). */
const STATS_MAX_RANGES = 1024;
/** Most getStats requests for one selection; a selection split into more ranges is aggregated only partly. */
const STATS_MAX_REQUESTS = 16;

/**
 * Selection statistics via getStats, debounced. Single cells show nothing (like Excel). Hidden rows (hidden in the
 * file or by a filter) are left out like in Excel's status bar (hidden columns are counted there): the ranges are
 * split around them, over several requests when there are many gaps.
 */
const scheduleStats = debounce(() => {
  if (!G || G.sheet < 0) return;
  const view = activeView();
  const sm = view.meta;
  const sel = view.sel;
  const area = cellArea(view, sel.active.r, sel.active.c);
  const r = sel.ranges[0];
  const single = sel.ranges.length === 1 && r.r0 === area.r0 && r.r1 === area.r1 && r.c0 === area.c0 && r.c1 === area.c1;
  if (single) {
    G.statsReq = null;
    clearStats();
    return;
  }
  /** @type {Range[]} */
  const ranges = [];
  for (const rg of sel.ranges) {
    const r1 = Math.min(rg.r1, sm.rowCount - 1);
    const c1 = Math.min(rg.c1, sm.colCount - 1);
    if (rg.r0 <= r1 && rg.c0 <= c1) ranges.push({ r0: rg.r0, c0: rg.c0, r1, c1 });
  }
  const parts = withoutHiddenRows(hiddenRowsOf(sm), ranges);
  if (!parts.length) {
    G.statsReq = null;
    clearStats();
    return;
  }
  const sent = Math.min(parts.length, STATS_MAX_RANGES * STATS_MAX_REQUESTS);
  /** @type {StatsRequest} */
  const req = { ids: new Set(), stats: { count: 0, numCount: 0, sum: 0 }, requests: 0, cut: sent < parts.length };
  G.statsReq = req;
  G.dom.statusStats.classList.add('fv-pending');
  for (let i = 0; i < sent; i += STATS_MAX_RANGES) {
    const reqId = ++G.statsSeq;
    req.ids.add(reqId);
    req.requests++;
    post({ type: 'getStats', reqId, sheet: view.index, ranges: parts.slice(i, Math.min(sent, i + STATS_MAX_RANGES)) });
  }
}, STATS_DEBOUNCE_MS);

/** Hidden rows of a sheet (the host marks zero-height rows hidden too), ascending. @type {WeakMap<SheetMeta, number[]>} */
const hiddenRowCache = new WeakMap();

/** @param {SheetMeta} sm */
function hiddenRowsOf(sm) {
  let out = hiddenRowCache.get(sm);
  if (!out) {
    out = [];
    for (const key in sm.rows) if (sm.rows[key].hidden === true) out.push(Number(key));
    out.sort((a, b) => a - b);
    hiddenRowCache.set(sm, out);
  }
  return out;
}

/**
 * The ranges without their hidden rows (the same array when none of them holds one), else as disjoint ranges ascending
 * by row: like the host's forEachRangeRow, the rows are cut into bands covered by the same ranges and each band's
 * column spans are merged, so cells of overlapping ranges still count once when the ranges go out in several requests.
 * @param {number[]} hidden ascending @param {Range[]} ranges
 * @returns {Range[]}
 */
function withoutHiddenRows(hidden, ranges) {
  /** Index of the first hidden row at or after row r. */
  const from = (/** @type {number} */ r) => {
    let lo = 0;
    let hi = hidden.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (hidden[mid] < r) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const holds = (/** @type {Range} */ rg) => {
    const i = from(rg.r0);
    return i < hidden.length && hidden[i] <= rg.r1;
  };
  if (!ranges.some(holds)) return ranges;
  /** @type {Range[]} */
  const out = [];
  const cuts = [...new Set(ranges.flatMap((rg) => [rg.r0, rg.r1 + 1]))].sort((a, b) => a - b);
  for (let k = 0; k + 1 < cuts.length; k++) {
    const top = cuts[k];
    const bottom = cuts[k + 1] - 1;
    /** @type {[number, number][]} */
    const spans = [];
    for (const rg of ranges.filter((g) => g.r0 <= top && g.r1 >= bottom).sort((a, b) => a.c0 - b.c0)) {
      const last = spans[spans.length - 1];
      if (last && rg.c0 <= last[1] + 1) last[1] = Math.max(last[1], rg.c1);
      else spans.push([rg.c0, rg.c1]);
    }
    // Runs of visible rows between the band's hidden rows.
    for (let i = from(top), start = top; start <= bottom; i++) {
      const end = i < hidden.length && hidden[i] <= bottom ? hidden[i] - 1 : bottom;
      if (start <= end) for (const [c0, c1] of spans) out.push({ r0: start, r1: end, c0, c1 });
      start = end + 2;
    }
  }
  return out;
}

/** Empties the selection statistics (no stats for a single cell or an empty selection). */
function clearStats() {
  if (!G) return;
  const el = G.dom.statusStats;
  el.replaceChildren();
  el.classList.remove('fv-stats-partial');
  el.title = '';
}

/** @param {StatsMessage} msg */
function onStats(msg) {
  if (!G) return;
  const probe = rowCache.probes.get(msg.reqId);
  if (probe) {
    rowCache.probes.delete(msg.reqId);
    probe.resolve(msg.stats.count);
    return;
  }
  const req = G.statsReq;
  if (!req || !req.ids.delete(msg.reqId)) return;
  // The answers of a selection split over several requests add up.
  const s = req.stats;
  const m = msg.stats;
  // Sum / Average / Min / Max in the first numeric cell's format (SelectionStats.text) come formatted from the host for
  // the cells of one request: kept when a single answer holds numbers, else the numbers are shown in our own format.
  if (m.numCount) s.text = s.numCount ? undefined : m.text;
  s.count += m.count;
  s.numCount += m.numCount;
  s.sum += m.sum;
  if (m.errors) s.errors = (s.errors ?? 0) + m.errors;
  if (m.min !== undefined) s.min = s.min === undefined ? m.min : Math.min(s.min, m.min);
  if (m.max !== undefined) s.max = s.max === undefined ? m.max : Math.max(s.max, m.max);
  if (m.partial) (s.partial = true), (s.scanned = (s.scanned ?? 0) + (m.scanned ?? 0));
  if (req.ids.size) return;
  G.statsReq = null;
  if (s.numCount) s.avg = s.sum / s.numCount;
  const el = G.dom.statusStats;
  el.classList.remove('fv-pending');
  // A selection larger than the host's work cap (SelectionStats.partial) or split into more ranges than the requests
  // carry is aggregated only partly: say so.
  const partial = s.partial === true || req.cut;
  const mark = partial ? ' (partial)' : '';
  /** @type {[string, string][]} */
  const items = [];
  if (s.count > 0) items.push([`Count${mark}`, formatCount(s.count)]);
  // Like Excel, a selection with an error value (#DIV/0!, #N/A ...) shows only Count.
  if (s.numCount > 0 && !s.errors) {
    const text = s.text;
    items.push([`Sum${mark}`, text?.sum ?? formatStat(s.sum)]);
    if (s.avg !== undefined) items.push([`Average${mark}`, text?.avg ?? formatStat(s.avg)]);
    if (s.min !== undefined) items.push([`Min${mark}`, text?.min ?? formatStat(s.min)]);
    if (s.max !== undefined) items.push([`Max${mark}`, text?.max ?? formatStat(s.max)]);
  }
  el.classList.toggle('fv-stats-partial', partial);
  el.title = !partial
    ? ''
    : req.requests === 1 && !req.cut
      ? `The selection is too large to aggregate completely: these values cover only its first ${formatCount(s.scanned ?? 0)} cells (row by row).`
      : 'The selection is too large to aggregate completely: these values cover only part of it.';
  el.replaceChildren(...items.map(([label, value]) => h('span', { class: 'fv-stat' }, h('span', { class: 'fv-stat-label' }, `${label}:`), ' ', h('span', { class: 'fv-stat-value' }, value))));
}

/**
 * Status-bar number. Very small and very large magnitudes use Excel's General notation (1E-12, 1E+21): fixed
 * notation would show 0 or a 22-digit number.
 * @param {number} n
 */
function formatStat(n) {
  if (!Number.isFinite(n)) return String(n);
  const a = Math.abs(n);
  if (a !== 0 && (a < 1e-9 || a >= 1e15)) return formatGeneral(n);
  return Number(n.toPrecision(12)).toLocaleString(undefined, { maximumFractionDigits: 10 });
}

/** @param {number} n */
function formatCount(n) {
  return n.toLocaleString();
}

// ===== MARKDOWN =====
// Markdown preview (phase 1 = preview only). The host renders markdown -> HTML (src/renderers/markdown.ts); this
// section sanitizes it into a GitHub-like article and adds the interactive parts:
//   - mermaid diagrams: lazy `import('mermaid')`, rendered when they come near the viewport (IntersectionObserver),
//     one at a time with event-loop yields, SVG cached per source+theme; the theme is derived from the --vscode-*
//     colours and diagrams are re-themed whenever they change; until rendered a diagram is an empty sized placeholder
//   - copy buttons on every `pre > code` block, copying the code exactly as shown (result announced through a polite
//     live region)
//   - task checkboxes -> `toggleTask`, answered by `toggleTaskResult` (serialized: one edit per document version,
//     quick clicks are queued; only the renderer's own checkboxes count)
//   - link routing: `#anchor` scrolls inside the article and moves the focus there, anything else -> `openLink`
//     (the renderer's `data-href` when it belongs to the href); clicks never reach VS Code's own link handler
//   - `markdownUpdate`: the new HTML is parsed into an inert document and diffed block by block (top-level nodes,
//     line-number attributes ignored) against the article on screen; only new/changed blocks are sanitized and
//     built off-DOM, then swapped in (changed on-screen diagrams are rendered before the swap). Unchanged blocks
//     keep their DOM (loaded images, rendered diagrams, open <details>) and get the shifted line numbers. The block
//     at the top of the view stays where it was (no flicker, no jump), and so does the keyboard focus.
//   - very large documents (.md-large) use `content-visibility: auto` blocks so off-screen content costs nothing, and
//     are built progressively: the first screens at once, the rest in idle time (MD_FIRST_PAINT_CHARS).
// Uses the BOOTSTRAP / UTILITIES globals: post, sanitize, h, showError, root, onDispose, disposeView, debounce,
// getStateKey, setStateKey, log, logError, errorDetail.
// Helpers of this section are prefixed `md` (shared with DOCX below); DOCX-only helpers are prefixed `docx`.

// ----- MARKDOWN: TYPES, CONSTANTS & STATE -----

/**
 * @typedef {{ level: number, text: string, slug: string, line: number }} MdTocEntry
 * @typedef {{ type?: string, kind?: string, fileName?: string, html: string, toc?: MdTocEntry[], source?: string,
 *   mode?: string, version: number, readOnly?: boolean }} MdMessage  Fields shared by `init` (kind 'markdown') and
 *   `markdownUpdate` (`readOnly`: the document cannot be edited, task checkboxes are disabled; an update without it
 *   keeps the current state).
 * @typedef {{ key: string, dark: boolean, variables: Record<string, string | boolean> }} MdMermaidTheme
 *   Mermaid 'base' theme whose themeVariables are derived from the VS Code theme's colours (see mdMermaidTheme);
 *   `key` identifies it in the SVG cache and tells when diagrams must be re-themed.
 * @typedef {{ svg?: string, id?: string, error?: string, transient?: boolean }} MdDiagramResult
 *   Rendered diagram (`svg`, its ids prefixed with the render `id`, see mdFreshSvgIds) or the error to show.
 * @typedef {{ line: number, checked: boolean, version: number }} MdToggle
 *   A task toggle: source line, requested state, version of the article it applies to.
 * @typedef {Extract<HostMessage, { type: 'toggleTaskResult' }>} MdToggleResult
 * @typedef {{ anchor: Element | null, offset: number, top: number | null, until: number }} MdScrollKeeper
 *   Scroll position held through late layout changes (images, diagrams) until the user scrolls himself: `anchor`
 *   stays `offset` px below the top of the view; without an anchor, scrollTop stays `top` (null: nothing held). Until
 *   `until` (performance.now()), any size change of the content puts it back too (see mdHoldOnResize).
 * @typedef {{ line: number | null, offset: number, top: number }} MdSavedScroll
 *   Persisted view position: top-level block starting at source `line`, `offset` px from the top of the view.
 */

/**
 * Extra DOMPurify options for preview content: keep every `id`/`name` (GitHub-style `user-content-` prefix) instead
 * of dropping the ones that collide with DOM properties (`#title`, `#images`, ...). Anchor lookup accepts both forms.
 */
const MD_SANITIZE_OPTIONS = { SANITIZE_NAMED_PROPS: true };
const MD_ID_PREFIX = 'user-content-';
/** vscode.setState keys (merged into the shared state object). */
const MD_STATE_SCROLL_KEY = 'markdownScroll';
const DOCX_STATE_SCROLL_KEY = 'docxScrollTop';
const MD_SCROLL_SAVE_MS = 200;
/**
 * How long after a position is set (restore, update, in-page jump) a size change of the content puts it back: blocks of a
 * large document get their real height (content-visibility) only in the frames after they come near the view.
 */
const MD_SCROLL_SETTLE_MS = 1000;
/** On update, how long to wait for changed on-screen diagrams before swapping anyway (they then finish in place). */
const MD_DIAGRAM_SWAP_WAIT_MS = 800;
/** At most this many changed diagrams are rendered before an update is swapped in; the rest render lazily. */
const MD_DIAGRAM_PRERENDER_MAX = 4;
/**
 * Diagrams are rendered lazily, once they come within this distance of the viewport (IntersectionObserver
 * rootMargin, % of the view height): a document with hundreds of diagrams stays responsive.
 */
const MD_DIAGRAM_ROOT_MARGIN = '150% 0px';
/** Placeholder height of a diagram not rendered yet (px): estimate bounds (see mdEstimateDiagramHeight). */
const MD_DIAGRAM_MIN_HEIGHT = 120;
const MD_DIAGRAM_MAX_ESTIMATE = 520;
/**
 * HTML size from which the article uses `content-visibility: auto` blocks (.md-large): layout, paint and the
 * accessibility tree then only cover the blocks near the viewport (a 5 MB article otherwise takes seconds, and
 * tens of seconds with a screen reader attached).
 */
const MD_LARGE_DOCUMENT_CHARS = 300_000;
/**
 * A large document is built progressively: on `init`, blocks worth MD_FIRST_PAINT_CHARS of HTML (several screens,
 * counted from the restored position) are sanitized and shown at once; the others follow in document order in idle
 * time (requestIdleCallback; while the view is busy, a MD_DEFERRED_BUSY_SLICE_MS slice every MD_DEFERRED_TIMEOUT_MS),
 * about MD_DEFERRED_CHUNK_CHARS of HTML per sanitize call. Every block is still sanitized before it is inserted, and
 * find-in-page sees the text as it arrives.
 */
const MD_FIRST_PAINT_CHARS = 100_000;
const MD_DEFERRED_CHUNK_CHARS = 25_000;
const MD_DEFERRED_MIN_IDLE_MS = 8;
const MD_DEFERRED_TIMEOUT_MS = 50;
const MD_DEFERRED_BUSY_SLICE_MS = 16;
/**
 * Line-number attributes; ignored when comparing blocks (an edit above shifts every line number below it). Only
 * removed inside start tags (MD_START_TAG_RE): the same text in the content (prose, code) is content.
 */
const MD_LINE_ATTRS_RE = / data-(?:source-)?line="[^"]*"/g;
/**
 * A start tag of serialized HTML (outerHTML) up to its last attribute: the serializer writes every attribute as
 * ` name="value"` with `"` escaped, and text never contains a raw `<` (raw-text elements such as <script> and
 * <style> are removed by the sanitizer anyway).
 */
const MD_START_TAG_RE = /<[a-zA-Z][^\s/>]*(?: [^\s"'>/=]+="[^"]*")*/g;
const MD_LINE_SELECTOR = '[data-source-line],[data-line]';
const MD_COPY_FEEDBACK_MS = 1500;
/** A toggleTask the host has not answered (toggleTaskResult) after this long is given up (the checkbox reverts). */
const MD_TOGGLE_TIMEOUT_MS = 5000;
const MD_MERMAID_CACHE_MAX = 64;

const mdState = {
  /** Scroll container (`.md-view`); null until the first init. @type {HTMLElement | null} */
  view: null,
  /** Article on screen (patched in place by updates). @type {HTMLElement | null} */
  article: null,
  /** Detached holder of an update's new blocks while their diagrams render (not yet inserted). @type {HTMLElement | null} */
  pending: null,
  /** Polite live region of the view (copy feedback for screen readers). @type {HTMLElement | null} */
  status: null,
  /** Blocks of a large document not built yet (see MD_FIRST_PAINT_CHARS); null when the article is complete. */
  deferred: /** @type {MdDeferredBlocks | null} */ (null),
  fileName: '',
  mode: 'preview',
  source: '',
  /** HTML of the latest requested render (identical updates skip the DOM work). */
  html: '',
  /** @type {MdTocEntry[]} */
  toc: [],
  /** Latest TextDocument.version received from the host. */
  version: 0,
  /** The document is read-only (a git: version, a file VS Code shows read-only): task checkboxes are disabled. */
  readOnly: false,
  /** Bumped on every render; async work belonging to an older render is dropped. */
  generation: 0,
  /** Mermaid theme of the diagrams on screen (set by every render). @type {MdMermaidTheme} */
  theme: { key: '', dark: false, variables: {} },
  /** @type {MdScrollKeeper} */
  scroll: { anchor: null, offset: 0, top: null, until: 0 },
  /** Watches the diagrams of the current article; renders them when they get near the viewport. */
  diagramObserver: /** @type {IntersectionObserver | null} */ (null),
  /** Measures rendered diagrams (placeholder height of their next render). */
  diagramSizes: /** @type {ResizeObserver | null} */ (null),
  /** Diagrams of the current article that are within MD_DIAGRAM_ROOT_MARGIN of the viewport. @type {Set<Element>} */
  nearby: new Set(),
};

/**
 * The article is a sequence of blocks: one per top-level node of the rendered HTML. `sig` is the raw markup of that
 * node with the line-number attributes removed (an edit above shifts every line number below it); `nodes` are the
 * sanitized, interactive DOM nodes that show it (usually one; none when sanitizing removed everything).
 * @typedef {{ sig: string, nodes: Node[] }} MdBlock
 * @typedef {{ node: Node, markup: string, sig: string }} MdRawBlock
 *   Top-level node of a fresh (unsanitized, inert) render: `markup` is its serialized HTML ('' for whitespace).
 * @typedef {{ raws: MdRawBlock[], blocks: MdBlock[], next: number, handle: number }} MdDeferredBlocks
 *   The tail of a large document still to be built: `blocks[i]` (already in mdBlocks, no nodes yet) shows `raws[i]`;
 *   `next` is the first one not built, `handle` the scheduled idle callback.
 * @typedef {{ order: ({ old: MdBlock, raw: MdRawBlock } | { raw: MdRawBlock })[], removed: MdBlock[], start: number, end: number }} MdBlockPlan
 *   Result of diffing the blocks on screen against a fresh render: `order` is the new block sequence (an `old`
 *   block reused for an identical node, else the `raw` block to materialize), `removed` the old blocks that go,
 *   [start, end) the changed range of the old block list.
 */

/** Blocks of the article on screen, in order (mdState.article's children are their nodes). @type {MdBlock[]} */
let mdBlocks = [];

// ----- MARKDOWN: VIEW (init / update) -----

/**
 * `init` for kind 'markdown': builds the preview from scratch (also used to recover when an update arrives while
 * no preview is on screen, e.g. after an error page).
 * @param {MdMessage} msg
 */
function showMarkdown(msg) {
  try {
    mdResetTasks();
    const mode = msg.mode || 'preview';
    const view = h('div', {
      class: 'md-view',
      tabindex: '-1', // focusable so keyboard scrolling works right away
      role: 'document',
      'aria-label': msg.fileName || null,
      dataset: { mode },
      onclick: mdOnClick,
      onauxclick: mdOnClick,
      onchange: mdOnChange,
      onsubmit: mdPreventDefault, // raw-HTML forms must never navigate the webview
    });

    const status = h('div', { class: 'fv-sr-only', role: 'status', 'aria-live': 'polite' });
    view.append(status);

    mdState.view = view;
    mdState.article = null;
    mdState.pending = null;
    mdState.status = status;
    mdState.fileName = msg.fileName || '';
    mdState.mode = mode;
    mdState.generation++;
    mdMermaid.current = null; // the theme may have changed while no preview was watching it
    mdState.theme = mdMermaidTheme();
    mdState.scroll = mdCreateScrollKeeper(view, () => mdSaveScroll(view));
    mdAcceptMessage(msg);
    mdObserveTheme();
    mdObserveDiagrams(view);
    onDispose(mdDispose);

    const html = typeof msg.html === 'string' ? msg.html : '';
    mdState.html = html;
    const large = html.length > MD_LARGE_DOCUMENT_CHARS;
    const saved = mdReadSavedScroll();
    const raws = mdParseBlocks(html);
    // Large documents: the first blocks now (through the restored position), the rest progressively (MD_FIRST_PAINT_CHARS).
    const eager = large ? mdEagerBlockCount(raws, saved) : raws.length;
    const blocks = mdMaterializeBlocks(raws.slice(0, eager), mdState.theme, null);
    const deferred = raws.slice(eager).map((raw) => ({ sig: raw.sig, nodes: /** @type {Node[]} */ ([]) }));
    const placeholder = deferred.length ? null : mdEmptyPlaceholder(blocks);
    if (placeholder) blocks.push(placeholder);
    const article = h('article', {
      class: large ? 'md-article md-large' : 'md-article',
      dataset: { version: String(mdState.version) },
    });
    for (const block of blocks) article.append(...block.nodes);
    view.append(article);
    mdState.article = article;
    mdBlocks = [...blocks, ...deferred];
    root.replaceChildren(view);

    mdHoldOnResize(view, article, mdState.scroll);
    mdRestoreScroll(view, article, saved);
    mdWatchDiagrams([article], []);
    mdTasksAfterRender(article);
    if (deferred.length) mdStartDeferredBlocks(raws.slice(eager), deferred);
    keepScrollerFocus(view);
    view.focus({ preventScroll: true });
  } catch (err) {
    mdFail(err, 'Could not display the Markdown preview.', true);
  }
}

/** View disposal (new `init`, error page): stop async work and observers that belong to the old view. */
function mdDispose() {
  mdState.generation++; // drops updates still waiting for diagrams
  mdCancelDeferredBlocks();
  mdState.view = null;
  mdState.article = null;
  mdState.pending = null;
  mdState.status = null;
  mdState.nearby.clear();
  mdBlocks = [];
  mdResetTasks();
}

/**
 * `markdownUpdate`: show the new content, keep the scroll position.
 * @param {MdMessage} msg
 */
function updateMarkdown(msg) {
  if (!mdState.view || !mdState.view.isConnected || !mdState.article) {
    // No preview on screen (e.g. the error page): rebuild it like an `init` would, disposing the current view first.
    disposeView();
    showMarkdown({
      ...msg,
      type: 'init',
      kind: 'markdown',
      fileName: msg.fileName || mdState.fileName,
      mode: msg.mode || mdState.mode,
      readOnly: typeof msg.readOnly === 'boolean' ? msg.readOnly : mdState.readOnly,
    });
    return;
  }
  try {
    const html = typeof msg.html === 'string' ? msg.html : '';
    mdAcceptMessage(msg);
    if (html === mdState.html) {
      // Same output (the change did not affect the rendering): only the document version moved on. A pending
      // update for this html stamps the latest version itself when it lands.
      if (!mdState.pending) {
        mdState.article.dataset.version = String(mdState.version);
        mdTasksAfterRender(mdState.article);
      }
      return;
    }
    mdState.html = html;
    mdPatchArticle(html).catch((err) => mdFail(err, 'Could not update the Markdown preview.', true));
  } catch (err) {
    mdFail(err, 'Could not update the Markdown preview.', true);
  }
}

/**
 * Records the per-message fields (version, toc, source, read-only) of an init/update.
 * @param {MdMessage} msg
 */
function mdAcceptMessage(msg) {
  if (msg.type === 'init') mdState.readOnly = msg.readOnly === true;
  else if (typeof msg.readOnly === 'boolean') mdState.readOnly = msg.readOnly;
  mdState.version = Number.isFinite(msg.version) ? msg.version : mdState.version;
  mdState.toc = Array.isArray(msg.toc) ? msg.toc : [];
  mdState.source = typeof msg.source === 'string' ? msg.source : '';
}

// ----- MARKDOWN: BLOCKS -----

/**
 * Parses rendered HTML into an inert document (no script runs, no resource loads) and returns its top-level
 * element and text nodes, the units the preview is built, diffed and patched in, each serialized once.
 * @param {string} html
 * @returns {MdRawBlock[]}
 */
function mdParseBlocks(html) {
  // A leading <body> keeps leading <style>/<meta>/... in the body, like DOMPurify's own parse.
  const doc = new DOMParser().parseFromString('<!DOCTYPE html><body>' + html, 'text/html');
  /** @type {MdRawBlock[]} */
  const raws = [];
  doc.body.childNodes.forEach((node) => {
    if (node instanceof Element) {
      const markup = node.outerHTML;
      raws.push({ node, markup, sig: mdBlockSignature(markup) });
    } else if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || '';
      raws.push({ node, markup: text.trim() ? escapeHtml(text) : '', sig: '#text:' + text });
    }
  });
  return raws;
}

/**
 * Block signature: its markup without the line-number attributes of its tags.
 * @param {string} markup outerHTML of a top-level node
 */
function mdBlockSignature(markup) {
  return markup.replace(MD_START_TAG_RE, (tag) => (tag.includes(' data-') ? tag.replace(MD_LINE_ATTRS_RE, '') : tag));
}

/**
 * Sanitizes raw blocks into DOM nodes with their interactive parts (task checkboxes, copy buttons, diagrams from
 * the cache). All blocks go through ONE sanitize call (per-call setup is expensive), each wrapped in a marker
 * <div data-md-block="i"> that is unwrapped afterwards, so every block knows exactly which nodes show it.
 * Built off-DOM, in a holder appended to `staging` when given (diagrams being rendered there still count as live,
 * see mdIsLiveDiagram).
 * @param {MdRawBlock[]} raws
 * @param {MdMermaidTheme} theme
 * @param {HTMLElement | null} staging
 * @returns {MdBlock[]}
 */
function mdMaterializeBlocks(raws, theme, staging) {
  /** @type {MdBlock[]} */
  const blocks = raws.map((raw) => ({ sig: raw.sig, nodes: [] }));
  /** @type {string[]} */
  const markup = [];
  raws.forEach((raw, i) => {
    if (raw.markup) markup.push(`<div data-md-block="${i}">${raw.markup}</div>`);
    else blocks[i].nodes.push(document.createTextNode(raw.node.textContent || '')); // whitespace: inert
  });
  if (!markup.length) return blocks;
  const holder = document.createElement('div');
  if (staging) staging.append(holder);
  holder.innerHTML = sanitize(markup.join(''), MD_SANITIZE_OPTIONS);
  mdHardenCodeBlocks(holder);
  mdAddCopyButtons(holder);
  mdSetTasksReadOnly(holder);
  for (const div of mdCollectDiagrams(holder)) mdShowCachedDiagram(div, theme);
  for (const wrapper of Array.from(holder.children)) {
    const block = blocks[Number(wrapper.getAttribute('data-md-block'))];
    if (block) block.nodes = Array.from(wrapper.childNodes);
  }
  return blocks;
}

/**
 * The "empty document" placeholder block when no block shows anything, else null.
 * @param {MdBlock[]} blocks
 * @returns {MdBlock | null}
 */
function mdEmptyPlaceholder(blocks) {
  const visible = blocks.some((block) =>
    block.nodes.some((node) => node instanceof Element || (node.textContent || '').trim() !== ''),
  );
  return visible ? null : { sig: '#placeholder', nodes: [h('p', { class: 'md-empty' }, 'This document is empty.')] };
}

/**
 * How many leading blocks of a large document `init` builds at once: those through the block a persisted position
 * restores to, then MD_FIRST_PAINT_CHARS of HTML more (all of them for a pixel-only position).
 * @param {MdRawBlock[]} raws
 * @param {MdSavedScroll | null} saved
 */
function mdEagerBlockCount(raws, saved) {
  let i = 0;
  if (saved && saved.line === null && saved.top > 0) return raws.length;
  if (saved && saved.line !== null) {
    for (; i < raws.length; i++) {
      const node = raws[i].node;
      const line = node instanceof Element ? mdBlockLine(node) : null;
      if (line !== null && line > saved.line) break;
    }
  }
  for (let chars = 0; i < raws.length && chars < MD_FIRST_PAINT_CHARS; i++) chars += raws[i].markup.length;
  return i;
}

/**
 * Builds the rest of a large document in idle time, chunk by chunk in document order (appended to the article).
 * @param {MdRawBlock[]} raws
 * @param {MdBlock[]} blocks their (empty) blocks, already in mdBlocks
 */
function mdStartDeferredBlocks(raws, blocks) {
  /** @type {MdDeferredBlocks} */
  const deferred = { raws, blocks, next: 0, handle: 0 };
  mdState.deferred = deferred;
  /** @param {IdleDeadline} deadline */
  const step = (deadline) => {
    if (mdState.deferred !== deferred) return;
    // Idle: as long as the browser has idle time left; busy (the timeout fired): a short slice of work anyway.
    const busyEnd = performance.now() + MD_DEFERRED_BUSY_SLICE_MS;
    const more = () => (deadline.didTimeout ? performance.now() < busyEnd : deadline.timeRemaining() > MD_DEFERRED_MIN_IDLE_MS);
    try {
      do mdBuildDeferredChunk(deferred, MD_DEFERRED_CHUNK_CHARS);
      while (deferred.next < raws.length && more());
    } catch (err) {
      mdState.deferred = null;
      mdFail(err, 'Could not display the Markdown preview.', true);
      return;
    }
    if (deferred.next < raws.length) deferred.handle = requestIdleCallback(step, { timeout: MD_DEFERRED_TIMEOUT_MS });
    else mdState.deferred = null;
  };
  deferred.handle = requestIdleCallback(step, { timeout: MD_DEFERRED_TIMEOUT_MS });
}

/**
 * Builds whatever is still deferred, now: before an update is diffed against the blocks, before an in-page link looks
 * for its target.
 */
function mdFinishDeferredBlocks() {
  const deferred = mdState.deferred;
  if (!deferred) return;
  cancelIdleCallback(deferred.handle);
  mdState.deferred = null;
  mdBuildDeferredChunk(deferred, Infinity);
}

function mdCancelDeferredBlocks() {
  if (mdState.deferred) cancelIdleCallback(mdState.deferred.handle);
  mdState.deferred = null;
}

/**
 * Sanitizes and appends the next deferred blocks (at least one, about `maxChars` of HTML).
 * @param {MdDeferredBlocks} deferred
 * @param {number} maxChars
 */
function mdBuildDeferredChunk(deferred, maxChars) {
  const article = mdState.article;
  if (!article) return;
  const start = deferred.next;
  let end = start;
  for (let chars = 0; end < deferred.raws.length && (end === start || chars < maxChars); end++) chars += deferred.raws[end].markup.length;
  const built = mdMaterializeBlocks(deferred.raws.slice(start, end), mdState.theme, null);
  /** @type {Node[]} */
  const nodes = [];
  built.forEach((block, k) => {
    deferred.blocks[start + k].nodes = block.nodes;
    nodes.push(...block.nodes);
  });
  article.append(...nodes);
  deferred.next = end;
  mdWatchDiagrams(nodes, []);
  if (end >= deferred.raws.length) {
    const placeholder = mdEmptyPlaceholder(mdBlocks); // everything sanitized away
    if (placeholder) {
      mdBlocks.push(placeholder);
      article.append(...placeholder.nodes);
    }
  }
}

/**
 * Elements carrying line-number attributes, in document order (including `node` itself).
 * @param {Node} node
 * @returns {Element[]}
 */
function mdLineTargets(node) {
  if (!(node instanceof Element)) return [];
  const targets = Array.from(node.querySelectorAll(MD_LINE_SELECTOR));
  if (node.matches(MD_LINE_SELECTOR)) targets.unshift(node);
  return targets;
}

/**
 * A reused block has the same markup as the fresh render, but its source lines may have moved (also inside it,
 * e.g. a blank line added between two paragraphs of a list item): every `data-source-line` / `data-line` value
 * is copied over when the update lands (task checkboxes post them back to the host).
 * @param {MdBlock} block
 * @param {Node} raw
 * @returns {(() => void) | null} applies the new line numbers; null when the block's structure does not line up
 *   with the raw node (then it is rebuilt)
 */
function mdLineSync(block, raw) {
  const sources = mdLineTargets(raw);
  const targets = block.nodes.flatMap(mdLineTargets);
  if (sources.length !== targets.length) return null;
  return () => {
    for (let i = 0; i < targets.length; i++) {
      for (const name of ['data-source-line', 'data-line']) {
        const value = sources[i].getAttribute(name);
        if (value === null) targets[i].removeAttribute(name);
        else if (targets[i].getAttribute(name) !== value) targets[i].setAttribute(name, value);
      }
    }
  };
}

// ----- MARKDOWN: UPDATE (block patching) -----

/**
 * Applies a new render to the article on screen: identical blocks keep their DOM, the others are sanitized and
 * materialized off-DOM and swapped in. Changed diagrams near the viewport get up to MD_DIAGRAM_SWAP_WAIT_MS to
 * render first, so the page does not jump twice. The block at the top of the view stays where it was.
 * @param {string} html
 */
async function mdPatchArticle(html) {
  const view = mdState.view;
  const article = mdState.article;
  if (!view || !article) return;
  const generation = ++mdState.generation;
  const theme = mdMermaidTheme();
  mdState.theme = theme;

  const raws = mdParseBlocks(html);
  let plan = mdDiffBlocks(mdBlocks, raws);
  // A large document still being built (mdStartDeferredBlocks): the blocks not built yet are the tail of mdBlocks. An
  // edit above them leaves them deferred (they are built from the new render later); one reaching into them builds
  // the rest of the document first.
  let deferred = mdState.deferred;
  if (deferred && plan.end > mdBlocks.length - (deferred.raws.length - deferred.next)) {
    mdFinishDeferredBlocks();
    deferred = null;
    plan = mdDiffBlocks(mdBlocks, raws);
  }
  const unbuilt = new Set(deferred ? deferred.blocks.slice(deferred.next) : []);
  const builtAtPlan = deferred ? deferred.next : 0;

  // Reused blocks take the new line numbers when the update lands; new blocks are built off-DOM, in a staging element.
  /** @type {(MdBlock | null)[]} */
  const slots = [];
  /** @type {(() => void)[]} */
  const lineSyncs = [];
  /** @type {MdRawBlock[]} */
  const toBuild = [];
  /** @type {number[]} */
  const buildSlots = [];
  const removed = plan.removed.slice();
  for (const entry of plan.order) {
    if ('old' in entry && unbuilt.has(entry.old)) {
      slots.push(entry.old); // not built yet: stays deferred
      continue;
    }
    const sync = 'old' in entry ? mdLineSync(entry.old, entry.raw.node) : null;
    if ('old' in entry && sync) {
      slots.push(entry.old);
      lineSyncs.push(sync);
      continue;
    }
    if ('old' in entry) removed.push(entry.old); // structure did not line up: rebuild it
    buildSlots.push(slots.length);
    toBuild.push(entry.raw);
    slots.push(null);
  }
  const staging = h('div');
  const built = mdMaterializeBlocks(toBuild, theme, staging);
  buildSlots.forEach((slot, i) => (slots[slot] = built[i]));
  const next = /** @type {MdBlock[]} */ (slots);
  const added = new Set(built);
  const placeholder = unbuilt.size ? null : mdEmptyPlaceholder(next);
  if (placeholder) {
    next.push(placeholder);
    added.add(placeholder);
  }

  const waiting = mdChangedDiagramsOnScreen(view, plan, staging);
  if (waiting.length) {
    mdState.pending = staging;
    await Promise.race([Promise.all(waiting.map((div) => mdRenderDiagram(div, theme))), mdDelay(MD_DIAGRAM_SWAP_WAIT_MS)]);
    if (generation !== mdState.generation || view !== mdState.view) return; // superseded by a newer render / view
  }

  // Swap: remember which surviving block is at the top of the view and where keyboard focus is, patch, put the block
  // back where it was and the focus on the element that replaced the focused one.
  const surviving = new Set(next.filter((block) => !added.has(block)).flatMap((block) => block.nodes));
  const anchor = mdPickScrollAnchor(view, article, (el) => surviving.has(el));
  const focus = mdCaptureFocus(article, plan, next);
  const removedNodes = removed.flatMap((block) => block.nodes);
  for (const node of removedNodes) node.parentNode?.removeChild(node);
  let cursor = article.firstChild;
  for (const node of next.flatMap((block) => block.nodes)) {
    if (node === cursor) cursor = cursor.nextSibling;
    else article.insertBefore(node, cursor);
  }
  for (const sync of lineSyncs) sync();
  mdBlocks = next;
  if (deferred) {
    // The deferred tail now builds from the new render (shifted line numbers). Blocks of it built from the old render
    // while this update waited for its diagrams take the new line numbers like any reused block.
    const tailBlocks = next.slice(next.length - unbuilt.size);
    const tailRaws = raws.slice(raws.length - unbuilt.size);
    const stillDeferred = mdState.deferred === deferred;
    const builtSince = stillDeferred ? deferred.next - builtAtPlan : unbuilt.size;
    for (let i = 0; i < builtSince; i++) {
      const sync = mdLineSync(tailBlocks[i], tailRaws[i].node);
      if (sync) sync();
    }
    if (stillDeferred) Object.assign(deferred, { raws: tailRaws, blocks: tailBlocks, next: builtSince });
  }
  article.classList.toggle('md-large', html.length > MD_LARGE_DOCUMENT_CHARS);
  article.dataset.version = String(mdState.version); // the latest version that rendered to this html
  mdState.pending = null;

  mdKeepScroll(view, mdState.scroll, anchor ? anchor.anchor : null, anchor ? anchor.offset : 0, null);
  mdRestoreFocus(focus, view);
  mdWatchDiagrams(Array.from(added, (block) => block.nodes).flat(), removedNodes);
  mdTasksAfterRender(article);
  if (mdMermaidTheme().key !== theme.key) mdRethemeDiagrams(); // theme changed while this update was waiting
}

/**
 * @typedef {{ el: Element, target: MdBlock | null, selector: string, nth: number }} MdFocusMemo
 *   Keyboard focus inside a block an update replaces: the `nth` element matching `selector` of that block, to be
 *   focused in `target` (the block that takes its place) once the update is on screen.
 */

/**
 * Before an update is swapped in: where keyboard focus is, when the block holding it is about to be replaced (e.g. a
 * task toggled with Space re-renders its list). Without this the focus falls back to <body>, and the next Tab / Space
 * starts over at the top of the list: a keyboard user would toggle the wrong task.
 * @param {HTMLElement} article
 * @param {MdBlockPlan} plan
 * @param {MdBlock[]} next new block list (`plan.order` order; a placeholder may follow)
 * @returns {MdFocusMemo | null}
 */
function mdCaptureFocus(article, plan, next) {
  const active = document.activeElement;
  if (!active || !article.contains(active)) return null;
  const index = mdBlocks.findIndex((block) => block.nodes.some((node) => node.contains(active)));
  if (index < 0 || next.includes(mdBlocks[index])) return null; // kept as it is: so is the focus
  const block = mdBlocks[index];
  // Its replacement: the block rebuilt in its place, else the new block at the same index of the changed range.
  let slot = plan.order.findIndex((entry) => 'old' in entry && entry.old === block);
  if (slot < 0) {
    const newEnd = plan.order.length - (mdBlocks.length - plan.end);
    slot = newEnd > plan.start ? clamp(index, plan.start, newEnd - 1) : -1;
  }
  const selector = active instanceof HTMLInputElement ? `input[type="${CSS.escape(active.type)}"]` : active.localName;
  return { el: active, target: slot >= 0 ? next[slot] : null, selector, nth: mdBlockElements(block, selector).indexOf(active) };
}

/**
 * After the swap: focuses the counterpart of the element that had focus (same kind, same position in its block), or
 * the view when there is none. Does nothing when focus is still (or again) somewhere else.
 * @param {MdFocusMemo | null} memo
 * @param {HTMLElement} view
 */
function mdRestoreFocus(memo, view) {
  if (!memo || memo.el.isConnected) return;
  const active = document.activeElement;
  if (active && active !== document.body) return;
  const candidates = memo.target ? mdBlockElements(memo.target, memo.selector) : [];
  const el = memo.nth >= 0 ? candidates[Math.min(memo.nth, candidates.length - 1)] : undefined;
  if (el instanceof HTMLElement || el instanceof SVGElement) el.focus({ preventScroll: true });
  if (document.activeElement !== el) view.focus({ preventScroll: true }); // no counterpart, or not focusable
}

/**
 * Elements of a block matching `selector`, in document order.
 * @param {MdBlock} block
 * @param {string} selector
 * @returns {Element[]}
 */
function mdBlockElements(block, selector) {
  return block.nodes.flatMap((node) =>
    node instanceof Element ? [...(node.matches(selector) ? [node] : []), ...Array.from(node.querySelectorAll(selector))] : [],
  );
}

/**
 * Diffs the blocks on screen against the raw nodes of a fresh render (by signature): common prefix and suffix are
 * reused as they are; inside the changed middle, identical blocks are matched greedily in order (so an untouched
 * block between two edits is reused too).
 * @param {MdBlock[]} oldBlocks
 * @param {MdRawBlock[]} raws
 * @returns {MdBlockPlan}
 */
function mdDiffBlocks(oldBlocks, raws) {
  let start = 0;
  const max = Math.min(oldBlocks.length, raws.length);
  while (start < max && oldBlocks[start].sig === raws[start].sig) start++;
  let oldEnd = oldBlocks.length;
  let newEnd = raws.length;
  while (oldEnd > start && newEnd > start && oldBlocks[oldEnd - 1].sig === raws[newEnd - 1].sig) {
    oldEnd--;
    newEnd--;
  }

  /** @type {MdBlockPlan['order']} */
  const order = [];
  for (let i = 0; i < start; i++) order.push({ old: oldBlocks[i], raw: raws[i] });
  /** @type {Map<string, { list: number[], next: number }>} old middle blocks by signature, in increasing order */
  const candidates = new Map();
  for (let i = start; i < oldEnd; i++) {
    const entry = candidates.get(oldBlocks[i].sig);
    if (entry) entry.list.push(i);
    else candidates.set(oldBlocks[i].sig, { list: [i], next: 0 });
  }
  const reused = new Set();
  let last = start - 1;
  for (let k = start; k < newEnd; k++) {
    const entry = candidates.get(raws[k].sig);
    if (entry) while (entry.next < entry.list.length && entry.list[entry.next] <= last) entry.next++;
    const i = entry && entry.next < entry.list.length ? entry.list[entry.next++] : undefined;
    if (i === undefined) {
      order.push({ raw: raws[k] });
      continue;
    }
    reused.add(i);
    last = i;
    order.push({ old: oldBlocks[i], raw: raws[k] });
  }
  for (let k = newEnd; k < raws.length; k++) order.push({ old: oldBlocks[oldEnd + (k - newEnd)], raw: raws[k] });
  const removed = oldBlocks.slice(start, oldEnd).filter((_, j) => !reused.has(start + j));
  return { order, removed, start, end: oldEnd };
}

/**
 * New diagrams of an update (not in the cache) when the changed region is on or near the screen: they are
 * rendered before the swap (at most MD_DIAGRAM_PRERENDER_MAX; the rest render lazily).
 * @param {HTMLElement} view
 * @param {MdBlockPlan} plan
 * @param {HTMLElement} staging
 * @returns {HTMLElement[]}
 */
function mdChangedDiagramsOnScreen(view, plan, staging) {
  const fresh = Array.from(staging.querySelectorAll('.mermaid.md-mermaid-pending')).map(
    (el) => /** @type {HTMLElement} */ (el),
  );
  if (!fresh.length) return [];
  // Changed region of the old article: elements of blocks [start, end), or the boundary at `start`.
  const firstEl = mdNearestElement(plan.start, 1) || mdNearestElement(plan.start - 1, -1);
  const lastEl = mdNearestElement(plan.end - 1, -1) || firstEl;
  if (!firstEl || !lastEl) return fresh.slice(0, MD_DIAGRAM_PRERENDER_MAX);
  const viewRect = view.getBoundingClientRect();
  const margin = view.clientHeight * 1.5;
  const a = firstEl.getBoundingClientRect();
  const b = lastEl.getBoundingClientRect();
  const near = Math.max(a.bottom, b.bottom) >= viewRect.top - margin && Math.min(a.top, b.top) <= viewRect.bottom + margin;
  return near ? fresh.slice(0, MD_DIAGRAM_PRERENDER_MAX) : [];
}

/**
 * First element node of the blocks on screen, scanning from block `index` in `step` direction.
 * @param {number} index
 * @param {1 | -1} step
 * @returns {Element | null}
 */
function mdNearestElement(index, step) {
  for (let i = index; i >= 0 && i < mdBlocks.length; i += step) {
    const nodes = mdBlocks[i].nodes;
    for (let j = step > 0 ? 0 : nodes.length - 1; j >= 0 && j < nodes.length; j += step) {
      const node = nodes[j];
      if (node instanceof Element) return node;
    }
  }
  return null;
}

// ----- MARKDOWN: SCROLL -----

/**
 * Holds a scroll position through late layout changes and persists the position (debounced, flushed when the view
 * is disposed or the webview document unloads: reload, editor moved to another window). The held position is dropped
 * as soon as the user scrolls himself.
 * @param {HTMLElement} view
 * @param {() => void} saveState
 * @returns {MdScrollKeeper}
 */
function mdCreateScrollKeeper(view, saveState) {
  /** @type {MdScrollKeeper} */
  const keeper = { anchor: null, offset: 0, top: null, until: 0 };
  const release = () => {
    keeper.anchor = null;
    keeper.top = null;
  };
  for (const type of ['wheel', 'touchstart', 'pointerdown', 'keydown']) {
    view.addEventListener(type, release, { passive: true });
  }
  const save = debounce(() => {
    try {
      if (view.isConnected) saveState();
    } catch {
      // View state is a convenience only.
    }
  }, MD_SCROLL_SAVE_MS);
  view.addEventListener('scroll', () => save(), { passive: true });
  onDispose(() => save.flush());
  listen(window, 'pagehide', () => save.flush());
  // Images that finish loading late change the height of the content above: put the held position back.
  view.addEventListener('load', () => mdHoldScroll(view, keeper), true);
  return keeper;
}

/**
 * Holds a new position (`anchor` at `offset` px below the top of the view, or scrollTop `top` without an anchor) and
 * applies it now.
 * @param {HTMLElement} view
 * @param {MdScrollKeeper} keeper
 * @param {Element | null} anchor
 * @param {number} offset
 * @param {number | null} top
 */
function mdKeepScroll(view, keeper, anchor, offset, top) {
  keeper.anchor = anchor;
  keeper.offset = offset;
  keeper.top = anchor ? null : top;
  keeper.until = performance.now() + MD_SCROLL_SETTLE_MS;
  mdHoldScroll(view, keeper);
}

/**
 * Puts a held position back when the content changes size within MD_SCROLL_SETTLE_MS of being set. In a large document
 * the blocks are laid out with an estimated height (content-visibility: auto) until they come near the view; when they
 * get their real height, the held block would otherwise move away (a reload restored the view sections away from where
 * the user was, a jump landed lower than the scroll padding).
 * @param {HTMLElement} view
 * @param {Element} content the article / page (its height changes with any block's)
 * @param {MdScrollKeeper} keeper
 */
function mdHoldOnResize(view, content, keeper) {
  const observer = new ResizeObserver(() => {
    if (performance.now() <= keeper.until) mdHoldScroll(view, keeper);
  });
  observer.observe(content);
  onDispose(() => observer.disconnect());
}

/**
 * Re-applies the held scroll position, if any.
 * @param {HTMLElement} view
 * @param {MdScrollKeeper} keeper
 */
function mdHoldScroll(view, keeper) {
  if (!view.isConnected) return;
  if (keeper.anchor && keeper.anchor.isConnected) {
    const delta = keeper.anchor.getBoundingClientRect().top - view.getBoundingClientRect().top - keeper.offset;
    if (Math.abs(delta) >= 0.5) view.scrollTop += delta;
  } else if (keeper.top !== null) {
    view.scrollTop = keeper.top;
  }
}

/**
 * Index of the first element in `blocks` whose bottom edge is below the top of the view (binary search; blocks
 * are in document order). `blocks.length` when everything is above.
 * @param {HTMLElement} view
 * @param {Element[]} blocks
 */
function mdFirstVisibleBlock(view, blocks) {
  const top = view.getBoundingClientRect().top;
  let lo = 0;
  let hi = blocks.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (blocks[mid].getBoundingClientRect().bottom <= top) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * Before an update: the block to keep in place. The block at the top of the view if it survives the update, else
 * the nearest surviving block above it (the edited content then grows downwards), else below it. Nothing at the
 * very top of the document (new content inserted at the top stays visible).
 * @param {HTMLElement} view
 * @param {HTMLElement} article
 * @param {(el: Element) => boolean} survives
 * @returns {{ anchor: Element, offset: number } | null}
 */
function mdPickScrollAnchor(view, article, survives) {
  if (view.scrollTop <= 0) return null;
  const blocks = Array.from(article.children);
  if (!blocks.length) return null;
  const first = Math.min(mdFirstVisibleBlock(view, blocks), blocks.length - 1);
  let index = -1;
  for (let i = first; i >= 0 && index < 0; i--) if (survives(blocks[i])) index = i;
  for (let i = first + 1; i < blocks.length && index < 0; i++) if (survives(blocks[i])) index = i;
  if (index < 0) return null;
  const anchor = blocks[index];
  return { anchor, offset: anchor.getBoundingClientRect().top - view.getBoundingClientRect().top };
}

/**
 * Source line of a top-level block (its own `data-source-line`, else its first descendant's).
 * @param {Element} block
 * @returns {number | null}
 */
function mdBlockLine(block) {
  const el = block.hasAttribute('data-source-line') ? block : block.querySelector('[data-source-line]');
  const line = el ? Number(el.getAttribute('data-source-line')) : NaN;
  return Number.isInteger(line) && line >= 0 ? line : null;
}

/**
 * Persists the view position as "block at source line N, offset px from the top" (survives edits and reloads
 * better than a pixel offset, and works with content-visibility estimates). The block is the first one that starts
 * inside the view (it is laid out at once when the position is restored), else the one the view starts in.
 * @param {HTMLElement} view
 */
function mdSaveScroll(view) {
  const article = mdState.article;
  if (!article || view !== mdState.view) return;
  const blocks = Array.from(article.children);
  const viewRect = view.getBoundingClientRect();
  /** @type {MdSavedScroll} */
  const saved = { line: null, offset: 0, top: Math.round(view.scrollTop) };
  const first = Math.min(mdFirstVisibleBlock(view, blocks), blocks.length - 1);
  /** @param {number} i */
  const pick = (i) => {
    const line = mdBlockLine(blocks[i]);
    if (line === null) return false;
    saved.line = line;
    saved.offset = Math.round(blocks[i].getBoundingClientRect().top - viewRect.top);
    return true;
  };
  let found = false;
  for (let i = first; i >= 0 && i < blocks.length && !found; i++) {
    const top = blocks[i].getBoundingClientRect().top;
    if (top >= viewRect.bottom) break;
    if (top >= viewRect.top) found = pick(i);
  }
  for (let i = first; i >= 0 && !found; i--) found = pick(i);
  setStateKey(MD_STATE_SCROLL_KEY, saved);
}

/** @returns {MdSavedScroll | null} the position persisted by a previous view of this webview, if any */
function mdReadSavedScroll() {
  try {
    const value = /** @type {unknown} */ (getStateKey(MD_STATE_SCROLL_KEY));
    if (!value || typeof value !== 'object') return null;
    const { line, offset, top } = /** @type {Record<string, unknown>} */ (value);
    return {
      line: typeof line === 'number' && Number.isInteger(line) && line >= 0 ? line : null,
      offset: typeof offset === 'number' && Number.isFinite(offset) ? offset : 0,
      top: typeof top === 'number' && Number.isFinite(top) && top >= 0 ? top : 0,
    };
  } catch {
    return null;
  }
}

/**
 * Init: scroll to the persisted position and hold it while images and diagrams settle.
 * @param {HTMLElement} view
 * @param {HTMLElement} article
 * @param {MdSavedScroll | null} saved
 */
function mdRestoreScroll(view, article, saved) {
  if (!saved || (saved.line === null && saved.top <= 0)) return;
  /** @type {Element | null} */
  let anchor = null;
  if (saved.line !== null) {
    for (const block of Array.from(article.children)) {
      const line = mdBlockLine(block);
      if (line === null) continue;
      if (line > saved.line) break;
      anchor = block;
    }
  }
  if (anchor instanceof HTMLElement && article.classList.contains('md-large')) {
    // Laid out for real now, not with the content-visibility estimate: even a block taller than the view, which starts
    // above it, lands where it was (an estimated one could stay out of view and never get its real height).
    const block = anchor;
    block.style.contentVisibility = 'visible';
    window.setTimeout(() => block.style.removeProperty('content-visibility'), MD_SCROLL_SETTLE_MS);
  }
  mdKeepScroll(view, mdState.scroll, anchor, saved.offset, saved.top);
}

// ----- MARKDOWN: MERMAID -----

const mdMermaid = {
  /** @type {Promise<typeof import('mermaid').default> | null} */
  loading: null,
  /** Theme the library is currently initialized with (initialized on first use, again only on theme change). */
  initialized: /** @type {MdMermaidTheme | null} */ (null),
  /** Theme derived from the current VS Code theme; null when the theme changed since (recomputed on demand). */
  current: /** @type {MdMermaidTheme | null} */ (null),
  /** source -> last rendered height in px, LRU: placeholder height of the next render. @type {Map<string, number>} */
  heights: new Map(),
  /** Serializes initialize+render so every SVG really has the theme it is cached under. */
  queue: /** @type {Promise<unknown>} */ (Promise.resolve()),
  /** theme + source -> result, LRU (Map keeps insertion order). @type {Map<string, MdDiagramResult>} */
  cache: new Map(),
  /** Diagram element -> cache key it should show (latest request wins). @type {WeakMap<Element, string>} */
  wanted: new WeakMap(),
  /** Diagram element -> cache key it currently shows. @type {WeakMap<Element, string>} */
  shown: new WeakMap(),
  nextId: 0,
  /** Number of real mermaid.render() calls (diagnostics). */
  renderCount: 0,
  loadErrorLogged: false,
};

/**
 * Mermaid theme for the current VS Code theme (cached until the theme changes, see mdObserveTheme).
 * @returns {MdMermaidTheme}
 */
function mdMermaidTheme() {
  if (!mdMermaid.current) mdMermaid.current = mdComputeMermaidTheme();
  return mdMermaid.current;
}

/**
 * Mermaid's 'base' theme with themeVariables taken from the VS Code theme: editor background / foreground, the focus
 * colour as accent (contrast border in high contrast), the chart colours for multi-colour diagrams (pie, mindmap,
 * timeline), the UI font. So a diagram matches any theme, not only a generic light or dark look. Colours are resolved
 * to opaque #rrggbb (mermaid derives further shades from them with its own colour library).
 * @returns {MdMermaidTheme}
 */
function mdComputeMermaidTheme() {
  const classes = document.body.classList;
  const hc = classes.contains('vscode-high-contrast') || classes.contains('vscode-high-contrast-light');
  const darkClass = !classes.contains('vscode-high-contrast-light') && (classes.contains('vscode-dark') || classes.contains('vscode-high-contrast'));
  const probe = document.createElement('span');
  probe.style.display = 'none';
  document.body.append(probe);
  /** @param {string} value CSS colour (var(--vscode-…, fallback)) @param {number[]} under */
  const color = (value, under) => {
    probe.style.color = '';
    probe.style.color = value;
    return mdOpaqueRgb(getComputedStyle(probe).color, under);
  };
  const bg = color(`var(--vscode-editor-background, ${darkClass ? '#1e1e1e' : '#ffffff'})`, darkClass ? [30, 30, 30] : [255, 255, 255]);
  const fg = color(`var(--vscode-editor-foreground, ${darkClass ? '#cccccc' : '#333333'})`, bg);
  const accent = color(hc ? 'var(--vscode-contrastBorder, var(--vscode-focusBorder, #0078d4))' : 'var(--vscode-focusBorder, #0078d4)', bg);
  const chart = (/** @type {string} */ name, /** @type {string} */ fallback) => color(`var(--vscode-charts-${name}, ${fallback})`, bg);
  const palette = [chart('blue', '#3794ff'), chart('orange', '#d18616'), chart('green', '#89d185'), chart('purple', '#b180d7'), chart('red', '#f14c4c'), chart('yellow', '#cca700')];
  probe.style.fontFamily = 'var(--vscode-font-family, system-ui, sans-serif)';
  const fontFamily = getComputedStyle(probe).fontFamily;
  probe.style.color = 'var(--vscode-widget-shadow, transparent)';
  const shadow = getComputedStyle(probe).color; // rgba() kept as it is: a translucent shadow
  probe.remove();

  const hex = (/** @type {number[]} */ c) => '#' + c.map((v) => Math.round(clamp(v, 0, 255)).toString(16).padStart(2, '0')).join('');
  /** `a` over `b` with weight t */
  const mix = (/** @type {number[]} */ a, /** @type {number[]} */ b, /** @type {number} */ t) => a.map((v, i) => v * t + b[i] * (1 - t));
  const [, orange, , , red, yellow] = palette;
  const dark = mdRelativeLuminance(bg) < 0.4;
  /** @type {Record<string, string | boolean>} */
  const variables = {
    darkMode: dark,
    background: hex(bg),
    fontFamily,
    // flat borders in the node colour (not mermaid's blue-to-orange gradient), the theme's widget shadow (none in HC)
    useGradient: false,
    dropShadow: /^rgba\(.*,\s*0\)$|^transparent$/.test(shadow) ? 'none' : `drop-shadow(1px 2px 2px ${shadow})`,
    primaryColor: hex(mix(accent, bg, 0.16)),
    primaryTextColor: hex(fg),
    primaryBorderColor: hex(hc ? accent : mix(accent, bg, 0.75)), // high contrast: the theme's border colour as is
    secondaryColor: hex(mix(orange, bg, 0.16)),
    secondaryTextColor: hex(fg),
    secondaryBorderColor: hex(mix(orange, bg, 0.7)),
    tertiaryColor: hex(mix(fg, bg, 0.06)),
    tertiaryTextColor: hex(fg),
    tertiaryBorderColor: hex(mix(fg, bg, 0.35)),
    lineColor: hex(hc ? fg : mix(fg, bg, 0.75)),
    textColor: hex(fg),
    nodeTextColor: hex(fg),
    titleColor: hex(fg),
    edgeLabelBackground: hex(bg),
    noteBkgColor: hex(mix(yellow, bg, 0.18)),
    noteTextColor: hex(fg),
    noteBorderColor: hex(mix(yellow, bg, 0.65)),
    errorBkgColor: hex(mix(red, bg, 0.2)),
    errorTextColor: hex(fg),
    pieStrokeColor: hex(bg),
    pieOuterStrokeColor: hex(mix(fg, bg, 0.4)),
    pieTitleTextColor: hex(fg),
    pieSectionTextColor: hex(dark ? bg : fg), // dark text reads best on the saturated chart colours
    pieLegendTextColor: hex(fg),
  };
  for (let i = 0; i < 12; i++) {
    const c = i < palette.length ? palette[i] : mix(palette[i - palette.length], bg, 0.55);
    variables[`pie${i + 1}`] = hex(c);
    variables[`cScale${i}`] = hex(mix(c, bg, 0.35));
    variables[`cScaleLabel${i}`] = hex(fg);
  }
  return { key: JSON.stringify(variables), dark, variables };
}

/**
 * Parses a computed colour (rgb() / rgba() / color(srgb …)) and composites it over `under`.
 * @param {string} css
 * @param {number[]} under opaque [r, g, b]
 * @returns {number[]} opaque [r, g, b]
 */
function mdOpaqueRgb(css, under) {
  let rgba = null;
  const rgb = /rgba?\(([^)]+)\)/.exec(css);
  if (rgb) {
    const p = rgb[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    rgba = [p[0], p[1], p[2], p[3] === undefined ? 1 : p[3]];
  } else {
    const srgb = /color\(srgb ([\d.e-]+) ([\d.e-]+) ([\d.e-]+)(?: \/ ([\d.e-]+))?\)/.exec(css);
    if (srgb) rgba = [Number(srgb[1]) * 255, Number(srgb[2]) * 255, Number(srgb[3]) * 255, srgb[4] === undefined ? 1 : Number(srgb[4])];
  }
  if (!rgba || rgba.some((v) => !Number.isFinite(v))) return under.slice();
  const a = clamp(rgba[3], 0, 1);
  return [0, 1, 2].map((i) => rgba[i] * a + under[i] * (1 - a));
}

/**
 * WCAG relative luminance of an sRGB colour.
 * @param {number[]} rgb
 */
function mdRelativeLuminance(rgb) {
  const [r, g, b] = rgb.map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * Watches the theme: VS Code swaps the body class (vscode-light / vscode-dark / vscode-high-contrast) and rewrites the
 * --vscode-* variables on <html> (only the latter when switching between two themes of the same kind). When the
 * derived mermaid theme changes, the diagrams are re-rendered. Lives as long as the current view.
 */
function mdObserveTheme() {
  const observer = new MutationObserver(() => {
    mdMermaid.current = null;
    if (mdState.view && mdMermaidTheme().key !== mdState.theme.key) mdRethemeDiagrams();
  });
  observer.observe(document.body, { attributes: true, attributeFilter: ['class', 'style'] });
  observer.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'style'] });
  onDispose(() => observer.disconnect());
}

/**
 * Theme changed: diagrams near the viewport are re-rendered now (or taken from the cache), the others keep their
 * old SVG until they come near the viewport (no layout jumps far away from what the user is reading).
 */
function mdRethemeDiagrams() {
  const theme = mdMermaidTheme();
  mdState.theme = theme;
  const article = mdState.article;
  if (!article) return;
  article.querySelectorAll('.mermaid').forEach((el) => {
    const div = /** @type {HTMLElement} */ (el);
    if (div.dataset.mermaidSource === undefined) return;
    if (!mdShowCachedDiagram(div, theme) && mdState.nearby.has(div)) void mdRenderDiagram(div, theme);
  });
}

/**
 * Creates the per-view IntersectionObserver that renders diagrams lazily as they approach the viewport.
 * @param {HTMLElement} view scroll container
 */
function mdObserveDiagrams(view) {
  mdState.nearby.clear();
  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        const div = /** @type {HTMLElement} */ (entry.target);
        if (!entry.isIntersecting) {
          mdState.nearby.delete(div);
          continue;
        }
        mdState.nearby.add(div);
        if (mdState.article && mdState.article.contains(div)) void mdRenderDiagram(div, mdMermaidTheme());
      }
    },
    { root: view, rootMargin: MD_DIAGRAM_ROOT_MARGIN },
  );
  const sizes = new ResizeObserver(mdOnDiagramResize);
  mdState.diagramObserver = observer;
  mdState.diagramSizes = sizes;
  onDispose(() => {
    observer.disconnect();
    sizes.disconnect();
    if (mdState.diagramObserver === observer) mdState.diagramObserver = null;
    if (mdState.diagramSizes === sizes) mdState.diagramSizes = null;
  });
}

/**
 * Keeps the lazy-render observer in sync with the article: watches the diagrams of inserted blocks (it reports
 * each one once right away) and forgets those of removed blocks.
 * @param {Node[]} added
 * @param {Node[]} removed
 */
function mdWatchDiagrams(added, removed) {
  const observer = mdState.diagramObserver;
  if (!observer) return;
  for (const node of removed) {
    for (const div of mdDiagramsIn(node)) {
      observer.unobserve(div);
      if (mdState.diagramSizes) mdState.diagramSizes.unobserve(div);
      mdState.nearby.delete(div);
    }
  }
  for (const node of added) for (const div of mdDiagramsIn(node)) observer.observe(div);
}

/**
 * @param {Node} node
 * @returns {Element[]} `.mermaid` elements in `node`, including `node` itself
 */
function mdDiagramsIn(node) {
  if (!(node instanceof Element)) return [];
  const list = Array.from(node.querySelectorAll('.mermaid'));
  if (node.classList.contains('mermaid')) list.unshift(node);
  return list;
}

/**
 * True while `div` belongs to the article on screen or to the update being prepared.
 * @param {Element} div
 */
function mdIsLiveDiagram(div) {
  return !!((mdState.article && mdState.article.contains(div)) || (mdState.pending && mdState.pending.contains(div)));
}

/**
 * Records the source of every diagram of a freshly sanitized container (`data-mermaid-source`), so later updates
 * and theme changes can re-render after the content was replaced by the SVG. Until its SVG is ready a diagram is an
 * empty placeholder ("Rendering diagram…" is CSS generated content): its source is neither shown nor matched by
 * find-in-page, and the placeholder is as tall as the diagram was last time (else an estimate from its number of
 * lines), so the page moves as little as possible when it renders.
 * @param {ParentNode} container
 * @returns {HTMLElement[]}
 */
function mdCollectDiagrams(container) {
  /** @type {HTMLElement[]} */
  const diagrams = [];
  container.querySelectorAll('.mermaid').forEach((el) => {
    const div = /** @type {HTMLElement} */ (el);
    const source = div.textContent || '';
    div.dataset.mermaidSource = source;
    div.textContent = '';
    div.classList.add('md-mermaid-pending');
    div.style.height = `${mdMermaid.heights.get(source) ?? mdEstimateDiagramHeight(source)}px`;
    div.setAttribute('aria-busy', 'true');
    diagrams.push(div);
  });
  return diagrams;
}

/**
 * Placeholder height of a diagram never rendered so far: grows with its number of lines (nodes, messages).
 * @param {string} source
 */
function mdEstimateDiagramHeight(source) {
  const lines = source.split('\n').filter((line) => line.trim() !== '').length;
  return clamp(40 + lines * 36, MD_DIAGRAM_MIN_HEIGHT, MD_DIAGRAM_MAX_ESTIMATE);
}

/**
 * Shows the diagram from the cache if possible.
 * @param {HTMLElement} div
 * @param {MdMermaidTheme} theme
 * @returns {boolean} false when it still has to be rendered
 */
function mdShowCachedDiagram(div, theme) {
  const source = div.dataset.mermaidSource || '';
  const key = mdMermaidKey(source, theme);
  mdMermaid.wanted.set(div, key);
  if (mdMermaid.shown.get(div) === key) return true;
  const hit = mdCacheGet(key);
  if (!hit) return false;
  mdShowDiagramResult(div, key, source, hit);
  return true;
}

/**
 * Renders one diagram for `theme` (cache first). Never throws: failures become an error box in the diagram.
 * @param {HTMLElement} div
 * @param {MdMermaidTheme} theme
 * @returns {Promise<void>}
 */
async function mdRenderDiagram(div, theme) {
  if (mdShowCachedDiagram(div, theme)) return;
  const source = div.dataset.mermaidSource || '';
  const key = mdMermaidKey(source, theme);
  // Superseded (theme switched again, newer request) or thrown away (newer update, view disposed): skip the work.
  const stillWanted = () => mdMermaid.wanted.get(div) === key && mdIsLiveDiagram(div);
  const result = await mdRenderMermaidSource(source, theme, stillWanted);
  if (!result || !stillWanted()) return;
  mdShowDiagramResult(div, key, source, result);
}

/**
 * @param {string} source
 * @param {MdMermaidTheme} theme
 */
function mdMermaidKey(source, theme) {
  return theme.key + '\n' + source;
}

/** @returns {Promise<typeof import('mermaid').default>} */
function mdLoadMermaid() {
  if (!mdMermaid.loading) {
    mdMermaid.loading = import('mermaid').then(
      (mod) => mod.default,
      (err) => {
        mdMermaid.loading = null; // allow a retry on the next render
        throw err;
      },
    );
  }
  return mdMermaid.loading;
}

/**
 * Renders a mermaid source to SVG through the serialized queue; results (including syntax errors) are cached.
 * Yields to the event loop before each render so input and painting are never blocked for long.
 * @param {string} source
 * @param {MdMermaidTheme} theme
 * @param {() => boolean} stillWanted checked when the job's turn comes; false skips it (resolves null)
 * @returns {Promise<MdDiagramResult | null>}
 */
function mdRenderMermaidSource(source, theme, stillWanted) {
  const key = mdMermaidKey(source, theme);
  const hit = mdCacheGet(key);
  if (hit) return Promise.resolve(hit);
  const task = mdMermaid.queue.then(async () => {
    const queuedHit = mdCacheGet(key); // an identical diagram may have been rendered while this one waited
    if (queuedHit) return queuedHit;
    if (!stillWanted()) return null;
    await mdDelay(0); // macrotask boundary between diagrams (mermaid's own awaits are microtasks)
    if (!stillWanted()) return null;
    let mermaid;
    try {
      mermaid = await mdLoadMermaid();
    } catch (err) {
      const message = 'Mermaid could not be loaded: ' + mdErrorMessage(err);
      if (!mdMermaid.loadErrorLogged) {
        mdMermaid.loadErrorLogged = true;
        log('error', message);
      }
      return /** @type {MdDiagramResult} */ ({ error: message, transient: true }); // not cached: retried later
    }
    if (!mdMermaid.initialized || mdMermaid.initialized.key !== theme.key) {
      mermaid.initialize({ startOnLoad: false, securityLevel: 'strict', theme: 'base', themeVariables: theme.variables, suppressErrorRendering: true });
      mdMermaid.initialized = theme;
    }
    const id = 'md-mermaid-' + ++mdMermaid.nextId;
    mdMermaid.renderCount++;
    /** @type {MdDiagramResult} */
    let result;
    try {
      const { svg } = await mermaid.render(id, source);
      result = { svg, id };
    } catch (err) {
      mdRemoveMermaidLeftovers(id);
      result = { error: mdErrorMessage(err) };
    }
    mdCacheSet(key, result);
    return result;
  });
  mdMermaid.queue = task.catch(() => undefined);
  return task;
}

/**
 * Removes temporary nodes mermaid may leave in <body> when a render fails.
 * @param {string} id
 */
function mdRemoveMermaidLeftovers(id) {
  for (const leftover of ['d' + id, 'i' + id, id]) {
    const el = document.getElementById(leftover);
    if (el && !(mdState.view && mdState.view.contains(el))) el.remove();
  }
}

/**
 * @param {HTMLElement} div
 * @param {string} key
 * @param {string} source
 * @param {MdDiagramResult} result
 */
function mdShowDiagramResult(div, key, source, result) {
  div.classList.remove('md-mermaid-pending');
  div.style.removeProperty('height'); // the placeholder height (mdCollectDiagrams)
  div.removeAttribute('aria-busy');
  if (typeof result.svg === 'string') {
    // Parse inertly, strip anything executable (defense in depth on top of securityLevel 'strict' + CSP), insert.
    const template = document.createElement('template');
    template.innerHTML = mdFreshSvgIds(result.svg, result.id);
    mdHardenSvg(template.content);
    div.classList.remove('md-mermaid-failed');
    div.replaceChildren(template.content);
  } else {
    div.classList.add('md-mermaid-failed');
    div.replaceChildren(mdDiagramErrorBox(result.error || 'Unknown error.', source));
  }
  // A transient failure (library could not be loaded) is retried the next time the diagram comes into view.
  if (result.transient) mdMermaid.shown.delete(div);
  else mdMermaid.shown.set(div, key);
  // Its rendered height becomes the placeholder height of the next render of this source (measured after layout).
  if (!result.transient && mdState.diagramSizes) mdState.diagramSizes.observe(div);
  // The diagram changed height: keep a held scroll position where it was.
  if (mdState.view && mdState.article && mdState.article.contains(div)) mdHoldScroll(mdState.view, mdState.scroll);
}

/**
 * A cached SVG can be on screen several times (identical diagrams, a block rebuilt by an update), and mermaid prefixes
 * every id in it (markers, gradients, the svg's own id its <style> is scoped to) with the id it was rendered under.
 * Each copy gets a fresh prefix, so ids stay unique in the page and `url(#…)` references stay inside their diagram.
 * @param {string} svg
 * @param {string | undefined} id render id of `svg`
 * @returns {string}
 */
function mdFreshSvgIds(svg, id) {
  if (!id) return svg;
  const fresh = 'md-mermaid-' + ++mdMermaid.nextId;
  // The render id ends in digits: a following digit belongs to another id (md-mermaid-1 vs md-mermaid-12).
  return svg.replace(new RegExp(id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![0-9])', 'g'), fresh);
}

/**
 * ResizeObserver callback: remembers the height of rendered diagrams (LRU by source), see mdCollectDiagrams.
 * @param {ResizeObserverEntry[]} entries
 */
function mdOnDiagramResize(entries) {
  const heights = mdMermaid.heights;
  for (const entry of entries) {
    const div = /** @type {HTMLElement} */ (entry.target);
    const source = div.dataset.mermaidSource;
    const box = entry.borderBoxSize && entry.borderBoxSize[0];
    const height = Math.round(box ? box.blockSize : entry.contentRect.height);
    if (source === undefined || height <= 0 || div.classList.contains('md-mermaid-pending')) continue;
    heights.delete(source); // refresh LRU position
    heights.set(source, height);
  }
  while (heights.size > MD_MERMAID_CACHE_MAX * 4) {
    const oldest = heights.keys().next();
    if (oldest.done) break;
    heights.delete(oldest.value);
  }
}

/**
 * @param {DocumentFragment} fragment
 */
function mdHardenSvg(fragment) {
  fragment.querySelectorAll('script').forEach((el) => el.remove());
  fragment.querySelectorAll('*').forEach((el) => {
    for (const attr of Array.from(el.attributes)) {
      const name = attr.name.toLowerCase();
      const isLink = name === 'href' || name === 'xlink:href' || name === 'src';
      if (name.startsWith('on') || (isLink && /^\s*(javascript|vbscript|data:text\/html)/i.test(attr.value))) {
        el.removeAttribute(attr.name);
      }
    }
  });
}

/**
 * Readable per-diagram error: what went wrong plus the diagram source. Never breaks the rest of the page.
 * @param {string} message
 * @param {string} source
 * @returns {HTMLElement}
 */
function mdDiagramErrorBox(message, source) {
  const box = h(
    'div',
    { class: 'md-mermaid-error', role: 'note', 'aria-label': 'Mermaid diagram error' },
    h('div', { class: 'md-mermaid-error-title' }, 'Mermaid diagram could not be rendered'),
    h('pre', { class: 'md-mermaid-error-message' }, message.trim()),
    h('div', { class: 'md-mermaid-error-label' }, 'Diagram source'),
    h('pre', { class: 'md-mermaid-error-source' }, h('code', null, source.replace(/\n$/, ''))),
  );
  mdAddCopyButtons(box); // the source block can be copied like any other code block
  return box;
}

/**
 * @param {string} key
 * @returns {MdDiagramResult | undefined}
 */
function mdCacheGet(key) {
  const value = mdMermaid.cache.get(key);
  if (value) {
    mdMermaid.cache.delete(key); // refresh LRU position
    mdMermaid.cache.set(key, value);
  }
  return value;
}

/**
 * @param {string} key
 * @param {MdDiagramResult} value
 */
function mdCacheSet(key, value) {
  mdMermaid.cache.set(key, value);
  while (mdMermaid.cache.size > MD_MERMAID_CACHE_MAX) {
    const oldest = mdMermaid.cache.keys().next();
    if (oldest.done) break;
    mdMermaid.cache.delete(oldest.value);
  }
}

// ----- MARKDOWN: CODE COPY -----
// A copy button copies exactly the code that is shown (no pastejacking): only the buttons this section made count, each
// copies its own <code> as rendered (innerText: no display:none / hidden / closed <details> text), and a code block of
// the document keeps no attribute or class that could hide, shrink, recolour or move part of its text while it is still
// rendered (GitHub's sanitizer removes these too).

/** @type {WeakMap<HTMLButtonElement, number>} */
const mdCopyTimers = new WeakMap();
/** Copy buttons made by mdAddCopyButtons -> the <code> each one copies. @type {WeakMap<HTMLButtonElement, HTMLElement>} */
const mdCopySources = new WeakMap();
/** Wrappers made by mdAddCopyButtons. @type {WeakSet<Element>} */
const mdCodeBlocks = new WeakSet();
/** Attributes kept on a code block (the <pre> and everything in it); `class` is filtered with MD_CODE_CLASS_RE. */
const MD_CODE_ATTR_RE = /^(?:href|title|lang|id|name|open|hidden|data-[^\s=]+)$/i;
/** highlight.js classes (tokens, sub-scopes such as `title function_`, the language). */
const MD_CODE_CLASS_RE = /^(?:hljs(?:-[\w-]+)?|language-\S+|[a-z]+_+)$/;

/**
 * Sanitized document content: code blocks that get a copy button keep only MD_CODE_ATTR_RE attributes and
 * highlight.js classes (no `style`, `color`, `size`, stylesheet classes such as fv-sr-only), and buttons of the document
 * lose the copy button look (they copy nothing, see mdOnClick).
 * @param {ParentNode} container
 */
function mdHardenCodeBlocks(container) {
  container.querySelectorAll('pre').forEach((pre) => {
    if (!pre.querySelector(':scope > code')) return;
    for (const el of [pre, ...Array.from(pre.querySelectorAll('*'))]) {
      for (const attr of Array.from(el.attributes)) {
        if (attr.name === 'class') {
          const kept = Array.from(el.classList).filter((name) => MD_CODE_CLASS_RE.test(name));
          if (kept.length) el.setAttribute('class', kept.join(' '));
          else el.removeAttribute('class');
        } else if (!MD_CODE_ATTR_RE.test(attr.name)) {
          el.removeAttribute(attr.name);
        }
      }
    }
  });
  container.querySelectorAll('.md-copy-button').forEach((el) => el.classList.remove('md-copy-button'));
}

/**
 * Copy button states: visible label (`data-label`, drawn by CSS as generated content, so it is never part of a text
 * selection, a copy or find-in-page), accessible name, and what the live region announces.
 */
const MD_COPY_STATES = {
  idle: { label: 'Copy', name: 'Copy code', announce: '' },
  copied: { label: 'Copied', name: 'Copied', announce: 'Code copied to the clipboard.' },
  failed: { label: 'Copy failed', name: 'Copy failed', announce: 'The code could not be copied.' },
};

/**
 * Wraps every `pre > code` block and adds a copy button next to it (outside the horizontally scrolling <pre>).
 * @param {ParentNode} container
 */
function mdAddCopyButtons(container) {
  container.querySelectorAll('pre > code').forEach((el) => {
    const code = /** @type {HTMLElement} */ (el);
    const pre = code.parentElement;
    const parent = pre && pre.parentElement;
    if (!pre || !parent || mdCodeBlocks.has(parent)) return; // the first <code> of a <pre> only
    if (pre.closest('table.front-matter')) return; // raw YAML values: no copy affordance inside the metadata table
    const block = h('div', { class: 'md-code-block' });
    mdCodeBlocks.add(block);
    parent.insertBefore(block, pre);
    const button = /** @type {HTMLButtonElement} */ (h('button', { class: 'md-copy-button', type: 'button', title: 'Copy code' }));
    mdSetCopyState(button, MD_COPY_STATES.idle);
    mdCopySources.set(button, code);
    block.append(pre, button);
  });
}

/**
 * @param {HTMLButtonElement} button
 * @param {{ label: string, name: string }} state
 */
function mdSetCopyState(button, state) {
  button.dataset.label = state.label;
  button.setAttribute('aria-label', state.name);
}

/**
 * @param {HTMLButtonElement} button
 */
async function mdCopyCode(button) {
  const code = mdCopySources.get(button);
  if (!code) return;
  let ok = true;
  try {
    // innerText of an element that is not rendered is its whole textContent: copy nothing then.
    if (!code.getClientRects().length) throw new Error('The code block is not shown.');
    await mdWriteClipboard(code.innerText.replace(/\n$/, ''));
  } catch {
    ok = false;
  }
  const state = ok ? MD_COPY_STATES.copied : MD_COPY_STATES.failed;
  mdSetCopyState(button, state);
  mdAnnounce(state.announce);
  button.classList.toggle('is-copied', ok);
  button.classList.toggle('is-failed', !ok);
  window.clearTimeout(mdCopyTimers.get(button));
  mdCopyTimers.set(
    button,
    window.setTimeout(() => {
      mdSetCopyState(button, MD_COPY_STATES.idle);
      button.classList.remove('is-copied', 'is-failed');
    }, MD_COPY_FEEDBACK_MS),
  );
}

/**
 * Announces a short status to screen readers through the view's polite live region (cleared first, so the same
 * message is announced again on the next copy).
 * @param {string} text
 */
function mdAnnounce(text) {
  const status = mdState.status;
  if (!status || !text) return;
  status.textContent = '';
  window.setTimeout(() => {
    if (status.isConnected) status.textContent = text;
  }, 50);
}

/**
 * navigator.clipboard first; falls back to execCommand('copy') when the async API is unavailable or refused.
 * @param {string} text
 */
async function mdWriteClipboard(text) {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch {
      // fall through
    }
  }
  const active = document.activeElement;
  const area = document.createElement('textarea');
  area.value = text;
  area.setAttribute('readonly', '');
  area.style.position = 'fixed';
  area.style.top = '0';
  area.style.left = '0';
  area.style.opacity = '0';
  document.body.append(area);
  area.select();
  const ok = document.execCommand('copy');
  area.remove();
  if (active instanceof HTMLElement) active.focus({ preventScroll: true });
  if (!ok) throw new Error('Copy command was rejected.');
}

// ----- MARKDOWN: TASK LISTS -----
// A task checkbox edits the source: `toggleTask` {line, checked, version of the article clicked in}, answered by the
// host with `toggleTaskResult` (applied or refused, and the document version after it). One toggle is in flight at a
// time; clicks meanwhile are queued (one per line, latest wins) and sent once the article on screen is as new as the
// document. A queued click is only replayed on a newer article when every version step in between came from our own
// applied toggles (they move no line); after any other change it is dropped, since its line may now be another task.
// Whatever happens, every checkbox shows the document's state (defaultChecked) plus the clicks still pending.
// Only the renderer's task checkboxes count: raw-HTML inputs arrive disabled (markdown.ts) and inputs inside a
// diagram (mermaid HTML labels) are ignored.

const mdTasks = {
  /** Toggle posted to the host, waiting for its toggleTaskResult. @type {MdToggle | null} */
  inflight: null,
  /** Toggles clicked while another was pending; one entry per line, latest click wins. @type {MdToggle[]} */
  queue: [],
  /** Applied toggles whose rendering has not arrived yet (still shown as clicked). @type {MdToggle[]} */
  confirmed: [],
  /** The host's document version after the last answered toggle: queued toggles wait for an article this new. */
  awaitVersion: 0,
  /** Document versions made by our own applied toggles (a checkbox edit moves no line). @type {Set<number>} */
  ownVersions: new Set(),
  timer: 0,
};

/** Task checkboxes disabled because the document is read-only (mdSetTasksReadOnly). @type {WeakSet<HTMLInputElement>} */
const mdReadOnlyTasks = new WeakSet();
const MD_READ_ONLY_TASK_TITLE = 'Read-only document';

/**
 * A checkbox the renderer produced for a task item (not raw HTML: disabled; not part of a diagram). A read-only
 * document's task checkboxes are disabled too, but still count (they show the document's state).
 * @param {Element} el
 * @returns {el is HTMLInputElement}
 */
function mdIsTaskCheckbox(el) {
  return (
    el instanceof HTMLInputElement &&
    el.classList.contains('task-list-item-checkbox') &&
    el.hasAttribute('data-line') &&
    (!el.disabled || mdReadOnlyTasks.has(el)) &&
    !el.closest('.mermaid, svg')
  );
}

/**
 * A read-only document (mdState.readOnly) shows its task checkboxes disabled, like GitHub's read-only task lists: a
 * click cannot change it. They are enabled again when an update says the document can be edited.
 * @param {ParentNode} container
 */
function mdSetTasksReadOnly(container) {
  const readOnly = mdState.readOnly;
  for (const input of mdTaskCheckboxes(container)) {
    if (mdReadOnlyTasks.has(input) === readOnly) continue;
    input.disabled = readOnly;
    if (readOnly) {
      mdReadOnlyTasks.add(input);
      input.title = MD_READ_ONLY_TASK_TITLE;
    } else {
      mdReadOnlyTasks.delete(input);
      input.removeAttribute('title');
    }
  }
}

/**
 * @param {Event} event
 */
function mdOnChange(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !mdIsTaskCheckbox(input)) return;
  const article = mdState.article;
  if (!article || !article.contains(input)) return;
  if (mdState.readOnly) {
    input.checked = input.defaultChecked; // read-only document: nothing to edit
    return;
  }
  const line = mdTaskLine(input);
  if (line === null) {
    input.checked = input.defaultChecked; // not mapped to a source line: nothing to edit
    return;
  }
  // The version of the article the checkbox belongs to (its line numbers), not the latest received one.
  mdRequestToggle({ line, checked: input.checked, version: Number(article.dataset.version) });
}

/**
 * The checkbox's own data-line (the renderer always writes it).
 * @param {HTMLInputElement} input
 * @returns {number | null}
 */
function mdTaskLine(input) {
  const raw = input.getAttribute('data-line');
  const line = raw === null || raw.trim() === '' ? NaN : Number(raw);
  return Number.isInteger(line) && line >= 0 ? line : null;
}

/**
 * Sends the toggle now, or queues it while another one is pending or the article on screen is older than the document.
 * @param {MdToggle} toggle
 */
function mdRequestToggle(toggle) {
  if (!mdTasks.inflight && toggle.version >= mdTasks.awaitVersion) {
    mdPostToggle(toggle);
    return;
  }
  mdTasks.queue = mdTasks.queue.filter((t) => t.line !== toggle.line);
  mdTasks.queue.push(toggle);
}

/**
 * @param {MdToggle} toggle
 */
function mdPostToggle(toggle) {
  mdTasks.inflight = toggle;
  post({ type: 'toggleTask', line: toggle.line, checked: toggle.checked, version: toggle.version });
  window.clearTimeout(mdTasks.timer);
  mdTasks.timer = window.setTimeout(mdOnToggleTimeout, MD_TOGGLE_TIMEOUT_MS);
}

/**
 * `toggleTaskResult`: settles the in-flight toggle. Applied: shown as clicked until its rendering arrives (it may
 * already have: VS Code reports the document change before the edit's completion). Refused: the checkbox shows the
 * document's state again at once; the click is not retried (the document changed under it, or it was not a task).
 * @param {MdToggleResult} msg
 */
function mdOnToggleResult(msg) {
  const toggle = mdTasks.inflight;
  if (!toggle || toggle.line !== msg.line || toggle.version !== msg.version) return; // late answer (timed out)
  window.clearTimeout(mdTasks.timer);
  mdTasks.inflight = null;
  const documentVersion = Number.isFinite(msg.documentVersion) ? msg.documentVersion : toggle.version;
  if (msg.applied) {
    if (documentVersion > toggle.version) mdTasks.ownVersions.add(toggle.version + 1);
    mdTasks.confirmed.push({ ...toggle, version: documentVersion });
  } else {
    log('info', `toggleTask (line ${msg.line}, version ${msg.version}) was not applied: ${msg.reason || 'refused'}.`);
  }
  mdTasks.awaitVersion = Math.max(mdTasks.awaitVersion, documentVersion);
  if (mdState.article) mdTasksAfterRender(mdState.article);
}

/**
 * Called whenever an article becomes current and after every toggle answer: sends the next queued toggle when the
 * article is as new as the document, then makes every checkbox show the document's state plus the pending clicks.
 * @param {HTMLElement} article
 */
function mdTasksAfterRender(article) {
  const version = Number(article.dataset.version);
  mdSetTasksReadOnly(article);
  if (mdState.readOnly) mdTasks.queue = []; // the document became read-only: queued clicks cannot be applied
  mdTasks.confirmed = mdTasks.confirmed.filter((t) => t.version > version); // rendered now: the source says so
  if (!mdTasks.inflight && version >= mdTasks.awaitVersion) {
    while (!mdTasks.inflight && mdTasks.queue.length) {
      const next = /** @type {MdToggle} */ (mdTasks.queue.shift());
      if (!mdTasksLinesKept(next.version, version)) {
        log('info', `toggleTask (line ${next.line}) dropped: the document changed since the click.`);
        continue;
      }
      const input = mdFindTaskCheckbox(article, next.line);
      if (!input || input.defaultChecked === next.checked) continue; // gone, or the source already says so
      mdPostToggle({ ...next, version });
    }
    if (!mdTasks.inflight && !mdTasks.queue.length && !mdTasks.confirmed.length) mdTasks.ownVersions.clear();
  }
  for (const input of mdTaskCheckboxes(article)) input.checked = input.defaultChecked;
  const pending = [...mdTasks.confirmed, ...(mdTasks.inflight ? [mdTasks.inflight] : []), ...mdTasks.queue];
  for (const toggle of pending) {
    const input = mdFindTaskCheckbox(article, toggle.line);
    if (input) input.checked = toggle.checked;
  }
}

/**
 * True when line numbers seen at version `from` are still valid at version `to`: every version in between was made
 * by one of our own task toggles.
 * @param {number} from
 * @param {number} to
 */
function mdTasksLinesKept(from, to) {
  for (let v = from + 1; v <= to; v++) if (!mdTasks.ownVersions.has(v)) return false;
  return true;
}

/** No answer for the posted toggle (host gone or too old to answer): show the document's state again. */
function mdOnToggleTimeout() {
  const lost = mdTasks.inflight;
  mdResetTasks();
  if (mdState.article) mdTasksAfterRender(mdState.article);
  if (lost) {
    log('warn', `toggleTask (line ${lost.line}, version ${lost.version}) was not answered; reverted.`);
  }
}

function mdResetTasks() {
  window.clearTimeout(mdTasks.timer);
  mdTasks.inflight = null;
  mdTasks.queue = [];
  mdTasks.confirmed = [];
  mdTasks.awaitVersion = 0;
  mdTasks.ownVersions.clear();
}

/**
 * The renderer's task checkboxes in `container`.
 * @param {ParentNode} container
 * @returns {HTMLInputElement[]}
 */
function mdTaskCheckboxes(container) {
  return Array.from(container.querySelectorAll('input.task-list-item-checkbox')).filter(mdIsTaskCheckbox);
}

/**
 * @param {ParentNode} container
 * @param {number} line
 * @returns {HTMLInputElement | null}
 */
function mdFindTaskCheckbox(container, line) {
  for (const el of container.querySelectorAll(`input.task-list-item-checkbox[data-line="${line}"]`)) {
    if (mdIsTaskCheckbox(el)) return el;
  }
  return null;
}

// ----- MARKDOWN: LINKS & ANCHORS -----

/** Elements whose activation navigates: HTML and SVG links, image-map areas (the sanitizer also removes those). */
const MD_LINK_SELECTOR = 'a, area';
const MD_XLINK_NS = 'http://www.w3.org/1999/xlink';
/** Same rule as markdown.ts: a reference with a scheme (or a drive letter) is not relative. */
const MD_SCHEME_RE = /^[a-z][a-z0-9+.-]*:/i;
/**
 * Where the renderer's resolved media links point: webview resource URIs (asWebviewUri), served from the origin of
 * this very script, or from another *.vscode-resource.vscode-cdn.net host (e.g. a remote authority).
 */
const MD_RESOURCE_ORIGIN = mdScriptOrigin();
const MD_RESOURCE_HOST_RE = /\.vscode-resource\.vscode-cdn\.net$/i;

/** Origin this script was loaded from ('' if unknown or opaque: then only the resource host pattern applies). */
function mdScriptOrigin() {
  try {
    const origin = new URL(import.meta.url).origin;
    return origin === 'null' ? '' : origin;
  } catch {
    return '';
  }
}

/**
 * Delegated click/auxclick handler of the markdown view: copy buttons and links.
 * @param {MouseEvent} event
 */
function mdOnClick(event) {
  const target = event.target instanceof Element ? event.target : null;
  if (!target) return;
  if (event.type === 'click') {
    const button = target.closest('.md-copy-button');
    if (button instanceof HTMLButtonElement && mdCopySources.has(button) && mdState.view && mdState.view.contains(button)) {
      event.preventDefault();
      void mdCopyCode(button);
      return;
    }
  } else if (event.button !== 1) {
    return; // auxclick: only the middle button opens links
  }
  const link = target.closest(MD_LINK_SELECTOR);
  const article = mdState.article;
  if (link && article && article.contains(link)) mdRouteLink(event, link, article, mdState.toc, mdState.scroll);
}

/**
 * `#fragment` scrolls inside `scope`; anything else goes to the host as `openLink`, which applies the scheme policy.
 * The click is consumed: the webview itself never navigates, and VS Code's own link handler (on the content window,
 * it ignores defaultPrevented) must not open the link a second time, possibly at a different URL.
 * @param {MouseEvent} event
 * @param {Element} link `<a>` (HTML or SVG) or `<area>`
 * @param {HTMLElement} scope
 * @param {MdTocEntry[] | null} toc
 * @param {MdScrollKeeper} keeper
 */
function mdRouteLink(event, link, scope, toc, keeper) {
  event.preventDefault();
  event.stopPropagation();
  // SVG links (e.g. inside an inline <svg>) may still use the legacy xlink:href attribute.
  const raw = link.getAttribute('href') ?? link.getAttributeNS(MD_XLINK_NS, 'href');
  if (raw === null) return; // plain named anchor
  const href = raw.trim();
  if (href.startsWith('#')) {
    mdScrollToFragment(scope, href.slice(1), toc, keeper);
    return;
  }
  const target = mdOriginalHref(link, href) ?? href;
  if (target) post({ type: 'openLink', href: target });
}

/**
 * The link's `data-href` when it is the renderer's: markdown.ts keeps the relative path of a link in `data-href` and
 * leaves the href as that path or, for media files, makes it the webview resource URI of the path. Any other
 * combination (e.g. raw HTML `<a href="https://a.example/" data-href="vscode:…">`) is ignored: a link always opens
 * what its href shows.
 * @param {Element} link
 * @param {string} href trimmed href
 * @returns {string | null}
 */
function mdOriginalHref(link, href) {
  const original = (link.getAttribute('data-href') || '').trim();
  if (!original || !mdIsRelativeReference(original)) return null;
  return original === href || mdIsResourceUri(href) ? original : null;
}

/**
 * Mirrors markdown.ts isRelativeReference: no scheme, not protocol-relative, not a fragment / query only.
 * @param {string} href
 */
function mdIsRelativeReference(href) {
  return !/^[#?]/.test(href) && !href.startsWith('//') && !href.startsWith('\\\\') && !MD_SCHEME_RE.test(href);
}

/**
 * True for an absolute webview resource URI (what resolveResource makes of a relative media path).
 * @param {string} href
 */
function mdIsResourceUri(href) {
  if (!MD_SCHEME_RE.test(href)) return false;
  try {
    const url = new URL(href);
    return (MD_RESOURCE_ORIGIN !== '' && url.origin === MD_RESOURCE_ORIGIN) || MD_RESOURCE_HOST_RE.test(url.hostname);
  } catch {
    return false;
  }
}

/**
 * `scrollToFragment`: the host found that a clicked link points to this very document (`readme.md#usage`,
 * `self.docx#bookmark`): behave like an in-page `#usage` link.
 * @param {Extract<HostMessage, { type: 'scrollToFragment' }>} msg
 */
function mdOnScrollToFragment(msg) {
  const fragment = encodeURIComponent(typeof msg.fragment === 'string' ? msg.fragment : '');
  if (currentKind() === 'docx') {
    if (docxState.page) mdScrollToFragment(docxState.page, fragment, null, docxState.scroll);
  } else if (mdState.article) {
    mdScrollToFragment(mdState.article, fragment, mdState.toc, mdState.scroll);
  }
}

/**
 * @param {HTMLElement} scope
 * @param {string} fragment raw (still URI-encoded) fragment without '#'
 * @param {MdTocEntry[] | null} toc
 * @param {MdScrollKeeper} keeper
 */
function mdScrollToFragment(scope, fragment, toc, keeper) {
  let id = fragment;
  try {
    id = decodeURIComponent(fragment);
  } catch {
    // malformed escape: use the raw text
  }
  keeper.anchor = null; // an explicit jump replaces any held position
  keeper.top = null;
  let target = id ? mdFindAnchorTarget(scope, id, toc) : null;
  if (!target && id && scope === mdState.article && mdState.deferred) {
    mdFinishDeferredBlocks(); // the target may be in a block not built yet
    target = mdFindAnchorTarget(scope, id, toc);
  }
  const view = scope.closest('.md-view, .docx-view');
  if (!target) {
    if (view instanceof HTMLElement && (!id || id.toLowerCase() === 'top')) {
      view.scrollTop = 0;
      view.focus({ preventScroll: true }); // the next Tab starts at the top, like the browser's own "#top"
    }
    return;
  }
  const marked = mdMarkedElement(target, scope);
  // The anchor's own line when it is inside the marked block (an anchor in the middle of a paragraph), else that block.
  const shown = marked !== target && marked.contains(target) && target.getClientRects().length ? target : marked;
  shown.scrollIntoView({ block: 'start', inline: 'nearest' });
  // Held where it landed while the blocks around it get their real height (a large document's blocks just built).
  if (view instanceof HTMLElement) {
    mdKeepScroll(view, keeper, shown, shown.getBoundingClientRect().top - view.getBoundingClientRect().top, null);
  }
  mdFlashTarget(marked);
  mdFocusTarget(marked);
}

/** Blocks an empty anchor stands for (see mdMarkedElement). */
const MD_MARKED_BLOCKS = 'h1, h2, h3, h4, h5, h6, p, li, dt, dd, td, th, caption, figcaption, summary, blockquote, pre';

/**
 * The element an in-page jump highlights and focuses. Usually the target itself; an empty named anchor (`<a name>`,
 * `<a id>` of a Word bookmark or TOC entry) has no area to highlight and cannot take focus, so the block it is in stands
 * for it (`<h1><a id="_Toc1"></a>Chapter 2</h1>`), or, when that block shows nothing else (an anchor alone on its line
 * before a heading), the next element that does.
 * @param {Element} target
 * @param {HTMLElement} scope
 * @returns {Element}
 */
function mdMarkedElement(target, scope) {
  /** @param {Element} el */
  const hasArea = (el) => {
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  };
  if (hasArea(target) || !target.getClientRects().length) return target; // shown, or not rendered at all
  const block = target.parentElement && target.parentElement.closest(MD_MARKED_BLOCKS);
  const start = block && scope.contains(block) ? block : target;
  if (start !== target && hasArea(start)) return start;
  for (let el = /** @type {Element | null} */ (start); el && el !== scope; el = el.parentElement) {
    for (let next = el.nextElementSibling; next; next = next.nextElementSibling) if (hasArea(next)) return next;
  }
  return start;
}

/**
 * Moves keyboard focus to the element an in-page link jumped to, as the browser's own fragment navigation moves the
 * sequential focus starting point: the next Tab continues after the target (e.g. inside the footnote) instead of
 * after the link, and screen readers read the target. A target that is not focusable gets tabindex="-1" while it has
 * focus (class md-focus-target: no focus ring, the flash highlight marks it).
 * @param {Element} el
 */
function mdFocusTarget(el) {
  if (!(el instanceof HTMLElement || el instanceof SVGElement)) return;
  if (el.tabIndex < 0 && !el.hasAttribute('tabindex')) {
    el.setAttribute('tabindex', '-1');
    el.classList.add('md-focus-target');
    el.addEventListener(
      'blur',
      () => {
        el.removeAttribute('tabindex');
        el.classList.remove('md-focus-target');
      },
      { once: true },
    );
  }
  el.focus({ preventScroll: true });
}

/**
 * Finds the element a fragment points to: `id` / `name` (with or without the sanitizer's `user-content-` prefix,
 * exact then lower-case), then the TOC (heading whose id did not survive sanitizing) by its source line.
 * @param {HTMLElement} scope
 * @param {string} id decoded fragment
 * @param {MdTocEntry[] | null} toc
 * @returns {Element | null}
 */
function mdFindAnchorTarget(scope, id, toc) {
  const variants = id === id.toLowerCase() ? [id] : [id, id.toLowerCase()];
  for (const variant of variants) {
    for (const candidate of [MD_ID_PREFIX + variant, variant]) {
      const escaped = CSS.escape(candidate);
      const el = scope.querySelector(`#${escaped}`) || scope.querySelector(`[name="${escaped}"]`);
      if (el) return el;
    }
  }
  if (toc && toc.length) {
    const lower = id.toLowerCase();
    const entry = toc.find((t) => t.slug === id) || toc.find((t) => t.slug.toLowerCase() === lower);
    if (entry) {
      const heading = scope.querySelector(
        ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].map((h) => `${h}[data-source-line="${entry.line}"]`).join(','),
      );
      if (heading) return heading;
    }
  }
  return null;
}

/**
 * Briefly highlights the element an in-page link jumped to.
 * @param {Element} el
 */
function mdFlashTarget(el) {
  el.classList.remove('md-target-flash');
  void (/** @type {HTMLElement} */ (el).offsetWidth); // restart the animation
  el.classList.add('md-target-flash');
  el.addEventListener('animationend', () => el.classList.remove('md-target-flash'), { once: true });
}

// ----- MARKDOWN: HELPERS -----

/**
 * @param {number} ms
 * @returns {Promise<void>}
 */
function mdDelay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

/**
 * @param {unknown} err
 * @returns {string}
 */
function mdErrorMessage(err) {
  if (err instanceof Error) return err.message || String(err);
  if (typeof err === 'string') return err;
  if (err && typeof err === 'object' && 'message' in err) return String(/** @type {{ message: unknown }} */ (err).message);
  return String(err);
}

/**
 * @param {Event} event
 */
function mdPreventDefault(event) {
  event.preventDefault();
}

/**
 * Unexpected failure while building a view: log it and show the error page (never a blank panel).
 * @param {unknown} err
 * @param {string} message
 * @param {boolean} canReopenAsText
 */
function mdFail(err, message, canReopenAsText) {
  logError(message, err);
  showError({ message, detail: errorDetail(err), canReopenAsText });
}

// ===== DOCX =====
// Read-only .docx preview: mammoth HTML (sanitized) on a centered, page-like sheet; conversion warnings collapsed
// above it; links routed like the markdown preview (`#bookmark` in-page, everything else -> openLink); form fields
// (check boxes) show their state but are disabled.

/**
 * Word stores hyperlinks to local files as `file:///C:\dir\Other.docx` (or relative paths). DOMPurify's default URI
 * policy drops `file:` hrefs, which left such links as dead text; here they are kept (DOMPurify's default list plus
 * `file:`) so a click reaches the host's openLink, which applies the same checks as for Markdown links. Nothing is
 * loaded from them: the page never navigates and the CSP allows no file: images.
 */
const DOCX_SANITIZE_OPTIONS = {
  ...MD_SANITIZE_OPTIONS,
  // Also drive-letter paths (`C:\dir\file`, `C:/dir/file`: Word stores these as they were typed), which DOMPurify would
  // read as an unknown `c:` scheme; the host opens them like file: links.
  ALLOWED_URI_REGEXP: /^(?:(?:(?:f|ht)tps?|mailto|tel|callto|sms|cid|xmpp|matrix|file):|[a-z]:[\\/]|[^a-z]|[a-z+.-]+(?:[^a-z+.\-:]|$))/i,
};
/** Text blocks that get their own direction (docxSetDirections; the CSS keeps their tabs and spaces). */
const DOCX_TEXT_BLOCKS = 'p, li, td, th, h1, h2, h3, h4, h5, h6, dt, dd, caption';
/** A letter of a right-to-left script. */
const DOCX_RTL_LETTER = '(?=\\p{L})[\\p{Script=Hebrew}\\p{Script=Arabic}\\p{Script=Syriac}\\p{Script=Thaana}\\p{Script=Nko}\\p{Script=Samaritan}\\p{Script=Mandaic}\\p{Script=Adlam}\\p{Script=Hanifi_Rohingya}]';
const DOCX_RTL_LETTER_RE = new RegExp(DOCX_RTL_LETTER, 'u');
const DOCX_RTL_LETTERS_RE = new RegExp(DOCX_RTL_LETTER, 'gu');
const DOCX_LETTERS_RE = /\p{L}/gu;
/** Words that may push a table past the page (shorter ones are not measured, see docxFitTables). */
const DOCX_LONG_WORD_RE = /\S{8,}/g;

const docxState = {
  /** @type {HTMLElement | null} */
  view: null,
  /** @type {HTMLElement | null} */
  page: null,
  /** @type {MdScrollKeeper} */
  scroll: { anchor: null, offset: 0, top: null, until: 0 },
};

/**
 * `init` for kind 'docx'.
 * @param {{ type?: string, kind?: string, fileName?: string, html: string, warnings?: string[] }} msg
 */
function showDocx(msg) {
  try {
    const page = h('article', { class: 'docx-page' });
    page.innerHTML = sanitize(typeof msg.html === 'string' ? msg.html : '', DOCX_SANITIZE_OPTIONS);
    // Read-only document: form fields (Word check boxes, mammoth renders them as <input type=checkbox>) show their
    // state but cannot be changed — like GitHub's task lists in a view the reader cannot edit.
    page.querySelectorAll('input, select, textarea, button').forEach((el) => {
      /** @type {HTMLInputElement} */ (el).disabled = true;
    });
    if (page.childElementCount === 0 && !(page.textContent || '').trim()) {
      page.replaceChildren(h('p', { class: 'docx-empty' }, 'This document has no displayable content.'));
    }
    const view = h(
      'div',
      {
        class: 'docx-view',
        tabindex: '-1',
        role: 'document',
        'aria-label': msg.fileName || null,
        onclick: docxOnClick,
        onauxclick: docxOnClick,
        onsubmit: mdPreventDefault,
      },
      h('div', { class: 'docx-stack' }, docxBuildWarnings(msg.warnings), page),
    );

    docxState.view = view;
    docxState.page = page;
    docxState.scroll = mdCreateScrollKeeper(view, () => setStateKey(DOCX_STATE_SCROLL_KEY, Math.round(view.scrollTop)));
    onDispose(() => {
      docxState.view = null;
      docxState.page = null;
    });
    docxSetDirections(page);
    root.replaceChildren(view);
    docxWatchTables(page);
    mdHoldOnResize(view, page, docxState.scroll);

    const saved = docxReadScrollTop();
    // images (data: URIs) decode late: hold the offset until the user scrolls
    if (saved !== undefined) mdKeepScroll(view, docxState.scroll, null, 0, saved);
    keepScrollerFocus(view);
    view.focus({ preventScroll: true });
  } catch (err) {
    mdFail(err, 'Could not display the Word document.', false);
  }
}

/**
 * Pixel scroll offset persisted by a previous docx view (vscode.getState), if any.
 * @returns {number | undefined}
 */
function docxReadScrollTop() {
  try {
    const value = getStateKey(DOCX_STATE_SCROLL_KEY);
    return typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Paragraph direction. Word takes it from the paragraph's own right-to-left flag (w:bidi), which mammoth drops, and
 * never from the first letter (unicode-bidi: plaintext would right-align an English paragraph that starts with a Hebrew
 * name). The closest guess: a text block is right-to-left when most of its letters are of a right-to-left script (a tie
 * goes to its first letter); a block without letters keeps the direction of the block around it.
 * @param {HTMLElement} page
 */
function docxSetDirections(page) {
  if (!DOCX_RTL_LETTER_RE.test(page.textContent || '')) return; // everything left-to-right, as the page is
  /** @type {Map<Element, string>} */
  const dirs = new Map();
  page.querySelectorAll(DOCX_TEXT_BLOCKS).forEach((el) => {
    const outer = el.parentElement && el.parentElement.closest(DOCX_TEXT_BLOCKS);
    const inherited = (outer && dirs.get(outer)) || 'ltr';
    const own = docxTextDirection(el.textContent || '') || inherited;
    dirs.set(el, own);
    if (own !== inherited) /** @type {HTMLElement} */ (el).dir = own;
  });
}

/**
 * @param {string} text
 * @returns {'ltr' | 'rtl' | null} null when the text has no letters
 */
function docxTextDirection(text) {
  const letters = (text.match(DOCX_LETTERS_RE) || []).length;
  if (!letters) return null;
  const rtl = (text.match(DOCX_RTL_LETTERS_RE) || []).length;
  if (rtl * 2 !== letters) return rtl * 2 > letters ? 'rtl' : 'ltr';
  const first = /\p{L}/u.exec(text);
  return first && DOCX_RTL_LETTER_RE.test(first[0]) ? 'rtl' : 'ltr';
}

/**
 * Keeps tables on the page (docxFitTables) now and whenever the page width changes (window, editor split).
 * @param {HTMLElement} page
 */
function docxWatchTables(page) {
  if (!page.querySelector('table')) return;
  let width = -1;
  const observer = new ResizeObserver(() => {
    if (page.clientWidth === width) return;
    width = page.clientWidth;
    docxFitTables(page);
  });
  observer.observe(page);
  onDispose(() => observer.disconnect());
}

/**
 * Like Word, a word too long for the page (a URL, a path) wraps inside its cell instead of pushing the table past the
 * page. Only where needed: in a table wider than the page, the cells with the longest words may break them anywhere
 * (class docx-cell-wrap), longest first, until the table fits. overflow-wrap: anywhere on every cell would let the
 * table layout squeeze the other columns below their own words ('Qua/ntit/y', '14,2/50.0/0').
 * @param {HTMLElement} page
 */
function docxFitTables(page) {
  const style = getComputedStyle(page);
  const available = page.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight) + 0.5;
  const range = document.createRange();
  page.querySelectorAll('table').forEach((table) => {
    if (table.parentElement && table.parentElement.closest('table')) return; // nested: fitted with its outer table
    if (table.getBoundingClientRect().width <= available) return;
    /** @type {{ cell: Element, widest: number }[]} */
    const cells = [];
    table.querySelectorAll('td, th').forEach((cell) => {
      if (cell.classList.contains('docx-cell-wrap')) return;
      let widest = 0;
      const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        for (const match of (node.nodeValue || '').matchAll(DOCX_LONG_WORD_RE)) {
          range.setStart(node, match.index);
          range.setEnd(node, match.index + match[0].length);
          widest = Math.max(widest, range.getBoundingClientRect().width);
        }
      }
      if (widest > 0) cells.push({ cell, widest });
    });
    cells.sort((a, b) => b.widest - a.widest);
    // A batch of similar widths per layout pass: few passes even for a column full of long links.
    let next = 0;
    while (next < cells.length && table.getBoundingClientRect().width > available) {
      const limit = cells[next].widest * 0.75;
      while (next < cells.length && cells[next].widest >= limit) cells[next++].cell.classList.add('docx-cell-wrap');
    }
    // Still too wide (many columns of short words): every cell wraps, as Word does in columns narrower than a word.
    if (table.getBoundingClientRect().width > available) {
      table.querySelectorAll('td, th').forEach((cell) => cell.classList.add('docx-cell-wrap'));
    }
  });
}

/**
 * @param {MouseEvent} event
 */
function docxOnClick(event) {
  if (event.type === 'auxclick' && event.button !== 1) return;
  const target = event.target instanceof Element ? event.target : null;
  const link = target && target.closest(MD_LINK_SELECTOR);
  const page = docxState.page;
  if (link && page && page.contains(link)) mdRouteLink(event, link, page, null, docxState.scroll);
}

/**
 * Collapsible list of mammoth warnings, identical messages grouped with a count.
 * @param {unknown} warnings
 * @returns {HTMLElement | null}
 */
function docxBuildWarnings(warnings) {
  if (!Array.isArray(warnings) || warnings.length === 0) return null;
  /** @type {Map<string, number>} */
  const counts = new Map();
  for (const warning of warnings) {
    const text = String(warning == null ? '' : warning).trim();
    if (text) counts.set(text, (counts.get(text) || 0) + 1);
  }
  if (counts.size === 0) return null;
  let total = 0;
  counts.forEach((n) => (total += n));

  return h(
    'details',
    { class: 'docx-warnings' },
    h(
      'summary',
      null,
      h('span', { class: 'docx-warnings-title' }, `${total} conversion warning${total === 1 ? '' : 's'}`),
      h('span', { class: 'docx-warnings-hint' }, 'Some formatting may differ from Word.'),
    ),
    h(
      'ul',
      { class: 'docx-warnings-list' },
      Array.from(counts, ([text, n]) =>
        h('li', null, text, n > 1 ? h('span', { class: 'docx-warning-count' }, `×${n}`) : null),
      ),
    ),
  );
}

// ===== PDF =====
// Read-only PDF view. pdf.js (loaded lazily) draws the pages with its viewer component (PDFViewer: canvas, text layer
// for selection / find, annotation layer for links) inside a VS Code-styled frame: toolbar (outline, page n / N,
// zoom, fit width / page, find), outline sidebar, floating find bar, password prompt. Pages stay white in every theme;
// the chrome follows the theme. Internal links and outline entries jump inside the view; every other link goes to the
// host as openLink (links under the PDF's folder as relative paths, like the docx policy).
//
// Webview constraints (see esbuild.js): the worker runs from a blob: URL of dist/pdf.worker.min.mjs (a webview cannot
// start a worker from its resource URL) and fetches nothing itself (`useWorkerFetch: false`: CMaps and standard fonts
// are fetched on this thread). The CSP allows no WebAssembly (`useWasm: false`): JBIG2 / CCITT / JPEG 2000 images are
// decoded by the asm.js fallbacks bundled into the worker.

/** @typedef {Extract<HostMessage, { type: 'init', kind: 'pdf' }>} PdfInitMessage */
/** @typedef {typeof import('pdfjs-dist')} PdfjsLib */
/** @typedef {{ lib: PdfjsLib, viewerLib: any }} PdfLibraries */

/**
 * Where pdf.js finds the files esbuild.js ships next to dist/viewer.js: the worker script and the data the PDF view
 * loads at run time (CMaps, standard fonts, annotation icons).
 */
const PDF_ASSETS = {
  worker: 'pdf.worker.min.mjs',
  cMaps: 'pdfjs/cmaps/',
  standardFonts: 'pdfjs/standard_fonts/',
  images: 'pdfjs/images/',
};
/** Webview state key: `{ page, zoom, left, top, outline }` of the last position (PdfSavedView). */
const PDF_STATE_KEY = 'pdfView';
/** Zoom names PDFViewer.currentScaleValue understands besides a scale factor. */
const PDF_ZOOM_PRESETS = ['auto', 'page-width', 'page-fit', 'page-actual'];
/** The presets that fit pages to the view (pdfApplyZoom keeps every page inside the width). */
const PDF_FIT_PRESETS = ['auto', 'page-width', 'page-fit'];
/** Room pdf.js leaves next to a page fitted to the width (its SCROLLBAR_PADDING: page margins and borders). */
const PDF_FIT_PADDING = 40;
/** Zoom of a PDF opened for the first time, and of Ctrl+0 (fit width, at most 125 %; like Firefox). */
const PDF_DEFAULT_ZOOM = 'auto';
const PDF_SAVE_DELAY_MS = 250;
/** Canvas redraw delay while zooming with Ctrl+Wheel / pinch (pages are scaled by CSS meanwhile). */
const PDF_WHEEL_ZOOM_DELAY_MS = 400;
/** Page Up / Page Down scroll by this share of the view's height, or by its height less the overlap if that is more. */
const PDF_PAGE_KEY_STEP = 0.875;
const PDF_PAGE_KEY_OVERLAP = 40;
/** Outline levels shown (deeper entries are left out; real outlines are far shallower). */
const PDF_MAX_OUTLINE_DEPTH = 32;
/** pdf.js FindState (web/pdf_find_controller.js). */
const PDF_FIND_STATE = { FOUND: 0, NOT_FOUND: 1, WRAPPED: 2, PENDING: 3 };

const PDF_ICONS = {
  outline: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M2.5 4h1M2.5 8h1M2.5 12h1M6 4h7.5M6 8h7.5M6 12h7.5" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>',
  chevronUp: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M3.5 10 8 5.5l4.5 4.5" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  chevronDown: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M3.5 6 8 10.5 12.5 6" fill="none" stroke="currentColor" stroke-width="1.4"/></svg>',
  fitWidth: '<svg viewBox="0 0 16 16" width="16" height="16"><rect x="1.5" y="3.5" width="13" height="9" rx="1" fill="none" stroke="currentColor"/><path d="M4 8h8M4 8l1.8-1.8M4 8l1.8 1.8M12 8l-1.8-1.8M12 8l-1.8 1.8" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  fitPage: '<svg viewBox="0 0 16 16" width="16" height="16"><rect x="3.5" y="1.5" width="9" height="13" rx="1" fill="none" stroke="currentColor"/><path d="M8 4v8M8 4 6.2 5.8M8 4l1.8 1.8M8 12l-1.8-1.8M8 12l1.8-1.8" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  search: '<svg viewBox="0 0 16 16" width="16" height="16"><circle cx="6.8" cy="6.8" r="4.3" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="m10 10 4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
};

/**
 * @typedef {object} PdfSavedView  position persisted in the webview state (restored on reopen and after a reload)
 * @property {number} page 1-based page at the top of the view
 * @property {string} zoom PDFViewer.currentScaleValue: a preset name or a scale factor ('1.25')
 * @property {number | null} left PDF-space point at the top-left of the view (pdf.js location), if known
 * @property {number | null} top
 * @property {boolean} outline the outline sidebar is open
 */

/**
 * @typedef {object} PdfDom
 * @property {HTMLElement} app
 * @property {HTMLButtonElement} outlineBtn
 * @property {HTMLButtonElement} prevBtn
 * @property {HTMLButtonElement} nextBtn
 * @property {HTMLInputElement} pageInput
 * @property {HTMLElement} pageCount
 * @property {HTMLButtonElement} zoomOutBtn
 * @property {HTMLButtonElement} zoomLabel
 * @property {HTMLButtonElement} zoomInBtn
 * @property {HTMLButtonElement} fitWidthBtn
 * @property {HTMLButtonElement} fitPageBtn
 * @property {HTMLButtonElement} findBtn
 * @property {HTMLElement} findBar
 * @property {HTMLInputElement} findInput
 * @property {HTMLButtonElement} findCaseBtn
 * @property {HTMLElement} findStatus
 * @property {HTMLButtonElement} findPrevBtn
 * @property {HTMLButtonElement} findNextBtn
 * @property {HTMLElement} sidebar
 * @property {HTMLElement} outline
 * @property {HTMLElement} container  PDFViewer container (absolutely positioned scroller)
 * @property {HTMLElement} viewer     PDFViewer viewer element (.pdfViewer)
 * @property {HTMLElement} overlay    loading message / password prompt over the pages
 */

/**
 * @typedef {object} PdfView
 * @property {boolean} disposed set when the view is replaced: pending pdf.js work then stops touching it
 * @property {string} fileName
 * @property {string | null} baseUrl the PDF's webview URI (relative links resolve against it); null for bytes
 * @property {PdfSavedView | null} saved position to restore once the pages are laid out
 * @property {PdfDom} dom
 * @property {AbortController} abort stops the viewer's own listeners (scroll, resize)
 * @property {PdfjsLib | null} lib
 * @property {any} loadingTask
 * @property {any} pdfDocument
 * @property {any} viewer PDFViewer
 * @property {any} eventBus
 * @property {any} linkService
 * @property {number} pagesCount
 * @property {{ pageNumber: number, scale: number | string, left: number, top: number } | null} location last updateviewarea
 * @property {string | null} zoomPreset the PDF_ZOOM_PRESETS zoom shown (also when pdfApplyZoom lowered a fit zoom to
 *   a scale factor), null for a zoom factor
 * @property {boolean} fitting set while pdfApplyZoom lowers a fit zoom (the scale change keeps zoomPreset)
 * @property {boolean} outlineOpen
 * @property {boolean} matchCase
 * @property {number} findState PDF_FIND_STATE of the last find
 * @property {{ pageNumber: number, destArray: any[] } | null} restoreAgain restored position to set once more on
 *   'pagesloaded' (every page has page 1's size until then); dropped when the user moves first
 * @property {(() => void) & { cancel(): void, flush(): void }} saveSoon
 */

const pdfState = {
  /** The PDF view on screen; null when another view (or the error view) is shown. @type {PdfView | null} */
  view: null,
};

/** pdf.js and its viewer component, loaded once per webview. @type {Promise<PdfLibraries> | null} */
let pdfLibraries = null;

/**
 * `init` for kind 'pdf': builds the frame at once, then loads pdf.js and the document (async; failures end in the
 * error view, a password-protected file asks for its password).
 * @param {PdfInitMessage} msg
 */
function showPdf(msg) {
  const fileName = typeof msg.fileName === 'string' && msg.fileName ? msg.fileName : 'file.pdf';
  const source = msg.source && typeof msg.source === 'object' ? msg.source : null;
  const baseUrl = source && 'uri' in source && typeof source.uri === 'string' ? source.uri : null;
  /** @type {PdfView} */
  const view = {
    disposed: false,
    fileName,
    baseUrl,
    saved: pdfReadState(),
    dom: pdfBuildDom(fileName),
    abort: new AbortController(),
    lib: null,
    loadingTask: null,
    pdfDocument: null,
    viewer: null,
    eventBus: null,
    linkService: null,
    pagesCount: 0,
    location: null,
    zoomPreset: null,
    fitting: false,
    outlineOpen: false,
    matchCase: false,
    findState: PDF_FIND_STATE.FOUND,
    restoreAgain: null,
    saveSoon: debounce(() => pdfSaveView(view), PDF_SAVE_DELAY_MS),
  };
  pdfState.view = view;
  onDispose(() => pdfDispose(view));

  root.className = 'pdf-app';
  onDispose(() => root.classList.remove('pdf-app'));
  root.replaceChildren(...(msg.banner ? [buildBanner(msg.banner, null)] : []), ...pdfBuildFrame(view));
  pdfWireFrame(view);
  pdfShowLoading(view, null);
  keepScrollerFocus(view.dom.container);
  view.dom.container.focus({ preventScroll: true });
  void pdfOpen(view, msg);
}

/** @param {PdfView} view */
function pdfDispose(view) {
  view.disposed = true;
  view.saveSoon.flush();
  if (pdfState.view === view) pdfState.view = null;
  view.abort.abort();
  try {
    view.viewer?.setDocument(null);
    view.linkService?.setDocument(null);
  } catch (err) {
    logError('PDF view cleanup failed', err);
  }
  // Destroys the document and its worker; a load still in progress rejects (ignored: the view is disposed).
  view.loadingTask?.destroy().catch(() => {});
}

// ----- PDF: LOADING -----

/** @returns {Promise<PdfLibraries>} */
function pdfLoadLibraries() {
  if (!pdfLibraries) {
    pdfLibraries = (async () => {
      const lib = await import('pdfjs-dist');
      // pdf_viewer.mjs reads the API from globalThis.pdfjsLib (set by pdfjs-dist) while it evaluates: import it second.
      const viewerLib = await import('pdfjs-dist/web/pdf_viewer.mjs');
      lib.GlobalWorkerOptions.workerSrc = await pdfWorkerUrl();
      return { lib, viewerLib };
    })();
    // A failed load (e.g. the worker file could not be read) is tried again by the next init.
    pdfLibraries.catch(() => {
      pdfLibraries = null;
    });
  }
  return pdfLibraries;
}

/**
 * blob: URL of the worker script. The webview cannot start a worker from its resource URL, and a worker cannot fetch
 * webview resources, so the script is fetched here and started from memory (CSP `worker-src blob:`).
 * @returns {Promise<string>}
 */
async function pdfWorkerUrl() {
  const url = new URL(PDF_ASSETS.worker, import.meta.url);
  const response = await fetch(url);
  if (!response.ok) throw new Error(`The PDF worker failed to load (HTTP ${response.status}) from ${url.href}`);
  const blob = new Blob([await response.arrayBuffer()], { type: 'text/javascript' });
  return URL.createObjectURL(blob);
}

/**
 * getDocument source from the init message: the webview URI (range-loaded by pdf.js; also the base of relative links)
 * or the bytes the host read (non-file documents). null when the message carries neither.
 * @param {PdfInitMessage} msg
 * @returns {{ url: string, docBaseUrl: string } | { data: Uint8Array } | null}
 */
function pdfDocumentSource(msg) {
  const source = /** @type {any} */ (msg.source);
  if (!source || typeof source !== 'object') return null;
  if (typeof source.uri === 'string' && source.uri) return { url: source.uri, docBaseUrl: source.uri };
  const data = source.data;
  if (data instanceof Uint8Array) return { data };
  if (data instanceof ArrayBuffer) return { data: new Uint8Array(data) };
  if (Array.isArray(data)) return { data: Uint8Array.from(data) };
  return null;
}

/**
 * Loads pdf.js and the document, then shows it. Ends in the error view on failure (unless the view was replaced).
 * @param {PdfView} view
 * @param {PdfInitMessage} msg
 */
async function pdfOpen(view, msg) {
  try {
    const source = pdfDocumentSource(msg);
    if (!source) throw new Error('The PDF view received no file to display.');
    // pdf.js would fail later with "x is not a function" (or draw part of a page); say plainly that VS Code is too old.
    const missing = missingPdfEngineFeatures();
    if (missing.length > 0) {
      throw Object.assign(new Error(`This web view lacks: ${missing.join(', ')}`), { name: PDF_ENGINE_UNSUPPORTED });
    }
    const { lib, viewerLib } = await pdfLoadLibraries();
    if (view.disposed) return;
    view.lib = lib;
    const task = lib.getDocument({
      ...source,
      cMapUrl: new URL(PDF_ASSETS.cMaps, import.meta.url).href,
      cMapPacked: true,
      standardFontDataUrl: new URL(PDF_ASSETS.standardFonts, import.meta.url).href,
      useWorkerFetch: false,
      useWasm: false,
      // pdf.js warnings (one "Failed to compile PostScript function to wasm" per Type 4 function, as wasm is off on
      // purpose; font quirks) are noise in the console: real failures go through pdfFail and the 'pagerendered' log.
      verbosity: lib.VerbosityLevel.ERRORS,
    });
    view.loadingTask = task;
    task.onPassword = (/** @type {(password: string) => void} */ update, /** @type {number} */ reason) => {
      if (!view.disposed) pdfShowPasswordPrompt(view, reason, update);
    };
    task.onProgress = (/** @type {{ loaded: number, total: number }} */ progress) => {
      if (!view.disposed && !view.pdfDocument && !view.dom.overlay.querySelector('form')) {
        const { loaded, total } = progress;
        pdfShowLoading(view, total > 0 ? clamp(Math.floor((loaded / total) * 100), 0, 100) : null);
      }
    };
    const pdfDocument = await task.promise;
    if (view.disposed) return;
    pdfShowDocument(view, viewerLib, pdfDocument);
  } catch (err) {
    if (view.disposed) return;
    pdfFail(view, err);
  }
}

/**
 * @param {PdfView} view
 * @param {unknown} err
 */
function pdfFail(view, err) {
  const message = `Could not open "${view.fileName}": ${describePdfFailure(err)}`;
  logError(message, err);
  showError({ message, detail: pdfErrorDetail(err) });
}

/**
 * errorDetail for pdf.js errors, which may also arrive as plain `{ name, message }` objects from the worker.
 * @param {unknown} err
 */
function pdfErrorDetail(err) {
  if (!(err instanceof Error) && err && typeof err === 'object' && typeof (/** @type {any} */ (err).message) === 'string') {
    const e = /** @type {{ name?: unknown, message: string }} */ (err);
    return `${typeof e.name === 'string' && e.name ? e.name : 'Error'}: ${e.message}`;
  }
  return errorDetail(err);
}

/**
 * Loading message over the (still empty) page area.
 * @param {PdfView} view
 * @param {number | null} percent
 */
function pdfShowLoading(view, percent) {
  const { overlay } = view.dom;
  const text = percent === null ? 'Loading PDF…' : `Loading PDF… ${percent}%`;
  const current = overlay.querySelector('.pdf-loading');
  if (current) current.textContent = text;
  else overlay.replaceChildren(h('p', { class: 'fv-boot-message pdf-loading', role: 'status', 'aria-live': 'polite' }, text));
  overlay.hidden = false;
}

/**
 * Asks for the password of an encrypted PDF (again, after a wrong one). The password goes to pdf.js only: it is not
 * logged or kept in the webview state.
 * @param {PdfView} view
 * @param {number} reason pdf.js PasswordResponses: 1 = needed, 2 = incorrect
 * @param {(password: string) => void} update
 */
function pdfShowPasswordPrompt(view, reason, update) {
  const incorrect = reason === (view.lib?.PasswordResponses?.INCORRECT_PASSWORD ?? 2);
  const input = /** @type {HTMLInputElement} */ (
    h('input', {
      class: 'pdf-password-input',
      type: 'password',
      autocomplete: 'off',
      spellcheck: 'false',
      'aria-label': 'Password',
      'aria-describedby': 'pdf-password-message',
      'aria-invalid': incorrect ? 'true' : null,
    })
  );
  const form = h(
    'form',
    {
      class: 'pdf-password',
      onsubmit: (/** @type {SubmitEvent} */ event) => {
        event.preventDefault();
        if (view.disposed) return;
        const password = input.value;
        pdfShowLoading(view, null);
        update(password);
      },
    },
    h('h2', { class: 'pdf-password-title' }, 'Password required'),
    h('p', { class: 'pdf-password-message', id: 'pdf-password-message' }, describePdfFailure({ name: 'PasswordException', code: reason })),
    h('div', { class: 'pdf-password-row' }, input, h('button', { class: 'fv-btn', type: 'submit' }, 'Open')),
  );
  view.dom.overlay.replaceChildren(form);
  view.dom.overlay.hidden = false;
  input.focus();
}

// ----- PDF: DOCUMENT -----

/**
 * Puts the loaded document into a PDFViewer and connects the toolbar, find bar and outline to it.
 * @param {PdfView} view
 * @param {any} viewerLib pdfjs-dist/web/pdf_viewer.mjs
 * @param {any} pdfDocument PDFDocumentProxy
 */
function pdfShowDocument(view, viewerLib, pdfDocument) {
  const lib = /** @type {PdfjsLib} */ (view.lib);
  const { dom } = view;
  const eventBus = new viewerLib.EventBus();
  const linkService = new viewerLib.PDFLinkService({ eventBus });
  const findController = new viewerLib.PDFFindController({ eventBus, linkService });
  const viewer = new viewerLib.PDFViewer({
    container: dom.container,
    viewer: dom.viewer,
    eventBus,
    linkService,
    findController,
    // View-only: annotations (links, form fields) are drawn but forms are not editable; no annotation editing.
    annotationMode: lib.AnnotationMode.ENABLE,
    annotationEditorMode: lib.AnnotationEditorType.DISABLE,
    imageResourcesPath: new URL(PDF_ASSETS.images, import.meta.url).href,
    abortSignal: view.abort.signal,
  });
  linkService.setViewer(viewer);
  Object.assign(view, { pdfDocument, viewer, eventBus, linkService, pagesCount: pdfDocument.numPages });

  eventBus.on('pagesinit', () => {
    if (!view.disposed) pdfRestoreView(view);
  });
  // Until now every page had the size of the first one: a fit zoom is lowered if a page turned out wider. (Not the
  // whole preset again: pdf.js's 'auto' depends on the current page, which may now be a restored landscape page.)
  eventBus.on('pagesloaded', () => {
    if (view.disposed) return;
    if (view.zoomPreset && PDF_FIT_PRESETS.includes(view.zoomPreset)) pdfFitWidestPage(view);
    // The restored point was placed with page 1's size for every page: on a page of another size it landed on the
    // wrong spot (even the previous page). Every page has its real size now.
    const again = view.restoreAgain;
    view.restoreAgain = null;
    if (again) viewer.scrollPageIntoView({ ...again, allowNegativeOffset: true });
  });
  eventBus.on('pagechanging', (/** @type {{ pageNumber: number }} */ e) => pdfUpdatePageControls(view, e.pageNumber));
  eventBus.on('scalechanging', (/** @type {{ scale: number, presetValue?: string }} */ e) => {
    if (e.presetValue) view.zoomPreset = e.presetValue;
    else if (!view.fitting) view.zoomPreset = null;
    pdfUpdateZoomControls(view, e.scale);
  });
  eventBus.on('updateviewarea', (/** @type {{ location: PdfView['location'] }} */ e) => {
    if (view.disposed || !e.location) return;
    view.location = e.location;
    view.saveSoon();
  });
  eventBus.on('updatefindcontrolstate', (/** @type {{ state: number, matchesCount: { current: number, total: number } }} */ e) =>
    pdfUpdateFindStatus(view, e.state, e.matchesCount),
  );
  eventBus.on('updatefindmatchescount', (/** @type {{ matchesCount: { current: number, total: number } }} */ e) =>
    pdfUpdateFindStatus(view, null, e.matchesCount),
  );
  /** @type {Set<number>} */
  const failedPages = new Set();
  eventBus.on('pagerendered', (/** @type {{ pageNumber: number, error?: any }} */ e) => {
    if (view.disposed || !e.error || e.error?.name === 'RenderingCancelledException' || failedPages.has(e.pageNumber)) return;
    failedPages.add(e.pageNumber);
    log('warn', `Page ${e.pageNumber} of ${view.fileName} could not be rendered completely: ${pdfErrorDetail(e.error)}`);
  });

  // A document without pages, or whose first page cannot be read (broken page tree or cross-reference table), would
  // leave the frame empty: PDFViewer needs page 1 to lay out every page and only logs to the console when it fails.
  if (view.pagesCount < 1) {
    pdfFail(view, new Error('This PDF has no pages.'));
    return;
  }
  viewer.setDocument(pdfDocument);
  linkService.setDocument(pdfDocument, null);
  viewer.pagesPromise?.catch((/** @type {unknown} */ err) => {
    if (!view.disposed && view.viewer === viewer) pdfFail(view, err);
  });

  dom.overlay.hidden = true;
  dom.overlay.replaceChildren();
  dom.pageCount.textContent = `of ${view.pagesCount}`;
  dom.pageInput.disabled = false;
  for (const btn of [dom.zoomOutBtn, dom.zoomLabel, dom.zoomInBtn, dom.fitWidthBtn, dom.fitPageBtn, dom.findBtn]) btn.disabled = false;
  pdfUpdatePageControls(view, 1);
  if (dom.app.contains(document.activeElement) || document.activeElement === document.body) dom.container.focus({ preventScroll: true });
  void pdfLoadOutline(view);
}

/**
 * pagesinit: zoom and position of the last visit (or the default zoom at the top).
 * @param {PdfView} view
 */
function pdfRestoreView(view) {
  const { viewer } = view;
  const saved = view.saved;
  pdfApplyZoom(view, saved ? saved.zoom : PDF_DEFAULT_ZOOM);
  if (!saved) return;
  const pageNumber = clamp(saved.page, 1, view.pagesCount);
  if (pageNumber === saved.page && saved.left !== null && saved.top !== null) {
    const destArray = [null, { name: 'XYZ' }, saved.left, saved.top, null];
    viewer.scrollPageIntoView({ pageNumber, destArray, allowNegativeOffset: true });
    // Set again on 'pagesloaded', when the pages above and this one have their real sizes (unless the user moved).
    if (pageNumber > 1) view.restoreAgain = { pageNumber, destArray };
  } else if (pageNumber > 1) {
    viewer.currentPageNumber = pageNumber;
  }
}

/** @returns {PdfSavedView | null} */
function pdfReadState() {
  try {
    const value = /** @type {any} */ (getStateKey(PDF_STATE_KEY));
    if (!value || typeof value !== 'object') return null;
    /** @param {unknown} n */
    const coord = (n) => (typeof n === 'number' && Number.isFinite(n) ? n : null);
    return {
      page: Number.isInteger(value.page) && value.page >= 1 ? value.page : 1,
      zoom: pdfZoomValue(value.zoom),
      left: coord(value.left),
      top: coord(value.top),
      outline: value.outline === true,
    };
  } catch {
    return null;
  }
}

/**
 * PDFViewer.currentScaleValue for a zoom from the state or a pdf.js location: a preset name, a percentage (125)
 * as a scale factor ('1.25'), or a scale factor saved by pdfSaveView ('1.6').
 * @param {unknown} zoom
 */
function pdfZoomValue(zoom) {
  if (typeof zoom === 'string' && PDF_ZOOM_PRESETS.includes(zoom)) return zoom;
  if (typeof zoom === 'number' && Number.isFinite(zoom) && zoom > 0) return String(clamp(zoom, 10, 2500) / 100);
  if (typeof zoom === 'string' && /^\d*\.?\d+$/.test(zoom) && Number(zoom) > 0) return String(clamp(Number(zoom), 0.1, 25));
  return PDF_DEFAULT_ZOOM;
}

/** @param {PdfView} view */
function pdfSaveView(view) {
  const location = view.location;
  if (!location) return; // nothing shown yet: keep the previous position
  /** @type {PdfSavedView} */
  const saved = {
    page: location.pageNumber,
    zoom: view.zoomPreset ?? pdfZoomValue(location.scale),
    left: Math.round(location.left),
    top: Math.round(location.top),
    outline: view.outlineOpen,
  };
  view.saved = saved;
  setStateKey(PDF_STATE_KEY, saved);
}

// ----- PDF: FRAME (toolbar, find bar, sidebar, pages) -----

/**
 * @param {string} fileName
 * @returns {PdfDom}
 */
function pdfBuildDom(fileName) {
  /**
   * @param {string} title tooltip (with the shortcut)
   * @param {string} label accessible name
   * @param {Node | string} content
   * @param {Record<string, any>} [attrs]
   */
  const button = (title, label, content, attrs) =>
    /** @type {HTMLButtonElement} */ (h('button', { class: 'fv-icon-btn', type: 'button', title, 'aria-label': label, disabled: true, ...attrs }, content));
  const viewer = h('div', { class: 'pdfViewer' });
  return {
    app: root,
    outlineBtn: button('Outline', 'Outline', icon(PDF_ICONS.outline), { 'aria-pressed': 'false', 'aria-controls': 'pdf-sidebar' }),
    prevBtn: button('Previous page', 'Previous page', icon(PDF_ICONS.chevronUp)),
    nextBtn: button('Next page', 'Next page', icon(PDF_ICONS.chevronDown)),
    pageInput: /** @type {HTMLInputElement} */ (
      h('input', { class: 'pdf-page-input', type: 'text', inputmode: 'numeric', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Page number', title: 'Page number', disabled: true })
    ),
    pageCount: h('span', { class: 'pdf-page-count' }),
    zoomOutBtn: button('Zoom out (Ctrl+-)', 'Zoom out', '−', { class: 'fv-icon-btn fv-zoom-btn' }),
    zoomLabel: /** @type {HTMLButtonElement} */ (h('button', { class: 'fv-zoom-label pdf-zoom-label', type: 'button', title: 'Reset zoom (Ctrl+0)', 'aria-label': 'Reset zoom', disabled: true }, '100%')),
    zoomInBtn: button('Zoom in (Ctrl+=)', 'Zoom in', '+', { class: 'fv-icon-btn fv-zoom-btn' }),
    fitWidthBtn: button('Fit width', 'Fit width', icon(PDF_ICONS.fitWidth), { 'aria-pressed': 'false' }),
    fitPageBtn: button('Fit page', 'Fit page', icon(PDF_ICONS.fitPage), { 'aria-pressed': 'false' }),
    findBtn: button('Find (Ctrl+F)', 'Find', icon(PDF_ICONS.search), { 'aria-expanded': 'false', 'aria-controls': 'pdf-findbar' }),
    findBar: h('div', { class: 'pdf-findbar', id: 'pdf-findbar', role: 'search', hidden: true }),
    findInput: /** @type {HTMLInputElement} */ (
      h('input', { class: 'pdf-find-input', type: 'text', placeholder: 'Find', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Find in document' })
    ),
    findCaseBtn: /** @type {HTMLButtonElement} */ (h('button', { class: 'pdf-find-option', type: 'button', title: 'Match Case', 'aria-label': 'Match case', 'aria-pressed': 'false' }, 'Aa')),
    findStatus: h('span', { class: 'pdf-find-status', role: 'status', 'aria-live': 'polite' }),
    findPrevBtn: button('Previous match (Shift+Enter)', 'Previous match', icon(PDF_ICONS.chevronUp), { disabled: false }),
    findNextBtn: button('Next match (Enter)', 'Next match', icon(PDF_ICONS.chevronDown), { disabled: false }),
    sidebar: h('nav', { class: 'pdf-sidebar', id: 'pdf-sidebar', 'aria-label': 'Outline', hidden: true }),
    outline: h('div', { class: 'pdf-outline' }),
    container: h('div', { class: 'pdf-container', tabindex: '0', role: 'document', 'aria-label': fileName }, viewer),
    viewer,
    overlay: h('div', { class: 'pdf-overlay' }),
  };
}

/**
 * The view's top-level parts (after the optional banner): toolbar and body (sidebar + pages, find bar floating).
 * @param {PdfView} view
 * @returns {HTMLElement[]}
 */
function pdfBuildFrame(view) {
  const d = view.dom;
  const closeFind = h(
    'button',
    { class: 'fv-icon-btn', type: 'button', title: 'Close (Escape)', 'aria-label': 'Close find', onclick: () => pdfCloseFind(view) },
    icon(ICONS.close),
  );
  d.findBar.append(d.findInput, d.findCaseBtn, d.findStatus, d.findPrevBtn, d.findNextBtn, closeFind);
  d.sidebar.append(d.outline);

  const toolbar = h(
    'div',
    { class: 'pdf-toolbar', role: 'toolbar', 'aria-label': 'PDF toolbar' },
    d.outlineBtn,
    h('span', { class: 'pdf-sep' }),
    h('span', { class: 'pdf-group', role: 'group', 'aria-label': 'Pages' }, d.prevBtn, d.nextBtn, d.pageInput, d.pageCount),
    h('span', { class: 'pdf-sep' }),
    h('span', { class: 'pdf-group fv-zoom', role: 'group', 'aria-label': 'Zoom' }, d.zoomOutBtn, d.zoomLabel, d.zoomInBtn, d.fitWidthBtn, d.fitPageBtn),
    h('span', { class: 'pdf-spacer' }),
    d.findBtn,
  );
  const main = h('div', { class: 'pdf-main' }, d.container, d.overlay, d.findBar);
  return [toolbar, h('div', { class: 'pdf-body' }, d.sidebar, main)];
}

/** @param {PdfView} view */
function pdfWireFrame(view) {
  const d = view.dom;
  d.outlineBtn.addEventListener('click', () => pdfToggleOutline(view));
  d.outline.addEventListener('keydown', pdfOnOutlineKeyDown);
  d.prevBtn.addEventListener('click', () => view.viewer?.previousPage());
  d.nextBtn.addEventListener('click', () => view.viewer?.nextPage());
  d.zoomOutBtn.addEventListener('click', () => pdfZoomBy(view, -1));
  d.zoomInBtn.addEventListener('click', () => pdfZoomBy(view, 1));
  d.zoomLabel.addEventListener('click', () => pdfSetZoom(view, PDF_DEFAULT_ZOOM));
  d.fitWidthBtn.addEventListener('click', () => pdfSetZoom(view, 'page-width'));
  d.fitPageBtn.addEventListener('click', () => pdfSetZoom(view, 'page-fit'));
  d.findBtn.addEventListener('click', () => (d.findBar.hidden ? pdfOpenFind(view) : pdfCloseFind(view)));

  d.pageInput.addEventListener('focus', () => d.pageInput.select());
  d.pageInput.addEventListener('change', () => pdfGoToTypedPage(view));
  d.pageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') pdfGoToTypedPage(view);
    else if (e.key === 'Escape') {
      pdfUpdatePageControls(view, view.viewer?.currentPageNumber ?? 1);
      d.container.focus({ preventScroll: true });
    }
  });
  d.pageInput.addEventListener('blur', () => {
    if (view.viewer) d.pageInput.value = String(view.viewer.currentPageNumber);
  });

  d.findInput.addEventListener('input', () => pdfFind(view, '', false));
  d.findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.isComposing) {
      e.preventDefault();
      pdfFind(view, 'again', e.shiftKey);
    }
  });
  d.findCaseBtn.addEventListener('click', () => {
    view.matchCase = !view.matchCase;
    d.findCaseBtn.setAttribute('aria-pressed', String(view.matchCase));
    pdfFind(view, 'casesensitivitychange', false);
  });
  d.findPrevBtn.addEventListener('click', () => pdfFind(view, 'again', true));
  d.findNextBtn.addEventListener('click', () => pdfFind(view, 'again', false));

  // Links in the pages: pdf.js follows internal ones itself (and cancels the click); the rest go to the host. The
  // click must not reach VS Code's own link handler on the window either (it would open the href directly).
  listen(d.container, 'click', (e) => pdfOnLinkClick(view, e));
  listen(d.container, 'auxclick', (e) => pdfOnLinkClick(view, e));
  listen(d.container, 'wheel', (e) => pdfOnWheel(view, e), { passive: false });
  // The user moves before the pages are laid out: the restored position is not set again over theirs.
  const keepUserPosition = () => (view.restoreAgain = null);
  for (const type of ['wheel', 'touchstart']) listen(d.container, type, keepUserPosition, { passive: true });
  listen(d.app, 'pointerdown', keepUserPosition, true);
  listen(window, 'keydown', keepUserPosition, true);
  // Capture phase: runs before VS Code's keydown listener on the window, which forwards keys to the workbench (where
  // Ctrl+= would zoom the whole window); handled keys stop there.
  listen(window, 'keydown', (e) => pdfOnKeyDown(view, e), true);

  // PDFViewer does not refit 'auto' / 'page-width' / 'page-fit' when its container is resized (window, sidebar,
  // dismissed banner): the pdf.js app does that itself, and so does this view.
  let lastSize = '';
  let frame = 0;
  const observer = new ResizeObserver(() => {
    cancelAnimationFrame(frame);
    frame = requestAnimationFrame(() => {
      const size = `${d.container.clientWidth}x${d.container.clientHeight}`;
      if (size === lastSize || view.disposed) return;
      lastSize = size;
      if (view.zoomPreset) pdfApplyZoom(view, view.zoomPreset);
    });
  });
  observer.observe(d.container);
  onDispose(() => {
    observer.disconnect();
    cancelAnimationFrame(frame);
  });
}

/**
 * @param {PdfView} view
 * @param {number} pageNumber
 */
function pdfUpdatePageControls(view, pageNumber) {
  if (view.disposed) return;
  const d = view.dom;
  if (document.activeElement !== d.pageInput) d.pageInput.value = String(pageNumber);
  d.prevBtn.disabled = !view.pdfDocument || pageNumber <= 1;
  d.nextBtn.disabled = !view.pdfDocument || pageNumber >= view.pagesCount;
}

/**
 * @param {PdfView} view
 * @param {number} scale
 */
function pdfUpdateZoomControls(view, scale) {
  if (view.disposed) return;
  const d = view.dom;
  d.zoomLabel.textContent = `${Math.round(scale * 100)}%`;
  // The text alone would be read as "125% button": say what the button does.
  d.zoomLabel.setAttribute('aria-label', `Reset zoom (current ${Math.round(scale * 100)}%)`);
  d.fitWidthBtn.setAttribute('aria-pressed', String(view.zoomPreset === 'page-width'));
  d.fitPageBtn.setAttribute('aria-pressed', String(view.zoomPreset === 'page-fit'));
}

/** @param {PdfView} view */
function pdfGoToTypedPage(view) {
  const d = view.dom;
  if (!view.viewer) return;
  const n = Number.parseInt(d.pageInput.value.trim(), 10);
  if (Number.isInteger(n)) pdfGoToPage(view, n);
  d.pageInput.value = String(view.viewer.currentPageNumber);
  d.pageInput.select();
}

/**
 * @param {PdfView} view
 * @param {number} pageNumber clamped to the document
 */
function pdfGoToPage(view, pageNumber) {
  if (!view.viewer || view.pagesCount < 1) return;
  view.viewer.currentPageNumber = clamp(Math.trunc(pageNumber), 1, view.pagesCount);
}

/**
 * @param {PdfView} view
 * @param {string} value PDFViewer.currentScaleValue
 */
function pdfSetZoom(view, value) {
  if (view.viewer && view.pdfDocument) pdfApplyZoom(view, value);
}

/**
 * Sets the zoom: a scale factor or a PDF_ZOOM_PRESETS name. pdf.js fits 'auto' / 'page-width' / 'page-fit' to the
 * CURRENT page only, so a wider page (a landscape page in a portrait document) stuck out of the view: a horizontal
 * scroll bar at "Fit width", and pdf.js's find, which scrolls each match to the horizontal centre, moved the pages left
 * and right. Like the browser's PDF viewer, a fit zoom here also keeps the widest page inside the view.
 * @param {PdfView} view
 * @param {string} value
 */
function pdfApplyZoom(view, value) {
  const { viewer } = view;
  if (!viewer) return;
  viewer.currentScaleValue = value;
  view.zoomPreset = PDF_ZOOM_PRESETS.includes(value) ? value : null;
  if (PDF_FIT_PRESETS.includes(value)) pdfFitWidestPage(view);
}

/**
 * Lowers the zoom of a fit preset so that the widest page fits the view's width too (no horizontal scroll bar; find,
 * which scrolls a match to the horizontal centre, then cannot move the pages sideways). Never raises it.
 * @param {PdfView} view
 */
function pdfFitWidestPage(view) {
  const { viewer } = view;
  if (!viewer) return;
  const width = view.dom.container.clientWidth - PDF_FIT_PADDING;
  let fit = Infinity;
  for (let i = 0; i < view.pagesCount; i++) {
    const page = viewer.getPageView(i);
    if (page && page.width > 0) fit = Math.min(fit, (width / page.width) * page.scale);
  }
  if (!(fit > 0 && fit < viewer.currentScale - 1e-4)) return;
  view.fitting = true;
  try {
    viewer.currentScale = fit;
  } finally {
    view.fitting = false;
  }
}

/**
 * @param {PdfView} view
 * @param {number} steps > 0 zooms in
 */
function pdfZoomBy(view, steps) {
  if (view.viewer && view.pdfDocument) view.viewer.updateScale({ steps });
}

/**
 * Ctrl+Wheel (and touchpad pinch, which arrives as Ctrl+Wheel) zooms around the pointer.
 * @param {PdfView} view
 * @param {WheelEvent} e
 */
function pdfOnWheel(view, e) {
  if (!(e.ctrlKey || e.metaKey)) return;
  e.preventDefault();
  const { viewer } = view;
  if (!viewer || !view.pdfDocument || e.deltaY === 0) return;
  const c = view.dom.container;
  const rect = c.getBoundingClientRect();
  // pdf.js subtracts the container's offsetLeft / offsetTop from `origin`.
  const origin = [e.clientX - rect.left + c.offsetLeft, e.clientY - rect.top + c.offsetTop];
  const pinch = e.deltaMode === WheelEvent.DOM_DELTA_PIXEL && Math.abs(e.deltaY) < 50;
  if (pinch) viewer.updateScale({ scaleFactor: Math.exp(-e.deltaY / 100), origin, drawingDelay: PDF_WHEEL_ZOOM_DELAY_MS });
  else viewer.updateScale({ steps: e.deltaY < 0 ? 1 : -1, origin, drawingDelay: PDF_WHEEL_ZOOM_DELAY_MS });
}

/**
 * True when typing keys belong to the focused element (page number, find text, password).
 * @param {EventTarget | null} target
 */
function pdfIsTextField(target) {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT') return true;
  return target instanceof HTMLInputElement && !['checkbox', 'radio', 'button', 'submit', 'reset'].includes(target.type);
}

/**
 * View shortcuts: Ctrl+F find, F3 / Shift+F3 and Ctrl+G / Ctrl+Shift+G next / previous match, Alt+C match case,
 * Ctrl+= / Ctrl+- / Ctrl+0 zoom, Page Up / Page Down one screen up / down (a page at "Fit page"), Home / End top /
 * bottom of the document, Escape closes the find bar. Arrow keys and Space scroll the focused pages natively.
 * @param {PdfView} view
 * @param {KeyboardEvent} e
 */
function pdfOnKeyDown(view, e) {
  if (e.isComposing || view.disposed) return;
  const ctrl = e.ctrlKey || e.metaKey;
  const key = e.key;
  const d = view.dom;
  // Next / previous match work while the find bar is open with a query (else Ctrl+G stays VS Code's).
  const findQuery = !d.findBar.hidden && d.findInput.value !== '';
  /** @type {(() => void) | null} */
  let action = null;
  if (ctrl && !e.altKey) {
    if (!e.shiftKey && (key === 'f' || key === 'F')) action = () => pdfOpenFind(view);
    else if (key === '+' || key === '=' || e.code === 'NumpadAdd') action = () => pdfZoomBy(view, 1);
    else if (key === '-' || key === '_' || e.code === 'NumpadSubtract') action = () => pdfZoomBy(view, -1);
    else if (key === '0' || e.code === 'Numpad0') action = () => pdfSetZoom(view, PDF_DEFAULT_ZOOM);
    else if ((key === 'g' || key === 'G') && findQuery) action = () => pdfFind(view, 'again', e.shiftKey);
  }
  // F3 / Shift+F3 like VS Code's find widget (and the browser's PDF viewer); with no query, F3 opens the find bar.
  if (!action && key === 'F3' && !ctrl && !e.altKey && view.viewer) {
    action = findQuery ? () => pdfFind(view, 'again', e.shiftKey) : () => pdfOpenFind(view);
  }
  if (!action && e.altKey && !ctrl && !e.shiftKey && e.code === 'KeyC' && !d.findBar.hidden) action = () => d.findCaseBtn.click();
  // Home / End in the outline move in the outline (pdfOnOutlineKeyDown), not in the pages.
  const inOutline = e.target instanceof Node && d.sidebar.contains(e.target);
  if (!action && !e.altKey && !e.shiftKey && !pdfIsTextField(e.target)) {
    if (!ctrl && (key === 'PageDown' || key === 'PageUp')) action = () => pdfPageKey(view, key === 'PageDown' ? 1 : -1);
    else if ((key === 'Home' || key === 'End') && !inOutline) action = () => pdfScrollToEnd(view, key === 'End');
  }
  if (
    !action &&
    key === 'Escape' &&
    !ctrl &&
    !e.altKey &&
    !e.shiftKey &&
    !view.dom.findBar.hidden &&
    (e.target === view.dom.findInput || !pdfIsTextField(e.target))
  ) {
    action = () => pdfCloseFind(view);
  }
  if (!action) return;
  e.preventDefault();
  e.stopPropagation();
  action();
}

/**
 * Page Up / Page Down: one screen up / down, less a little overlap (like the browser's own paging), so no part of a
 * page is skipped; at "Fit page", where a page fills the screen, the previous / next page (like pdf.js's app).
 * @param {PdfView} view
 * @param {number} dir 1 = down
 */
function pdfPageKey(view, dir) {
  const { viewer } = view;
  if (!viewer || !view.pdfDocument) return;
  if (view.zoomPreset === 'page-fit') {
    if (dir > 0) viewer.nextPage();
    else viewer.previousPage();
    return;
  }
  const c = view.dom.container;
  c.scrollTop += dir * Math.max(c.clientHeight * PDF_PAGE_KEY_STEP, c.clientHeight - PDF_PAGE_KEY_OVERLAP);
}

/**
 * Home / End (also with Ctrl): the top / the bottom of the document (not the top of the last page).
 * @param {PdfView} view
 * @param {boolean} end
 */
function pdfScrollToEnd(view, end) {
  if (!view.viewer || !view.pdfDocument) return;
  const c = view.dom.container;
  c.scrollTop = end ? c.scrollHeight : 0;
}

// ----- PDF: FIND -----

/** @param {PdfView} view */
function pdfOpenFind(view) {
  if (!view.viewer) return;
  const d = view.dom;
  // Reopened with the last query: its matches are highlighted again (closing removed them), so the "n of m" shown
  // still matches the page, like VS Code's find widget.
  if (d.findBar.hidden && d.findInput.value) pdfFind(view, 'highlightallchange', false);
  d.findBar.hidden = false;
  d.findBtn.setAttribute('aria-expanded', 'true');
  d.findInput.focus();
  d.findInput.select();
}

/** @param {PdfView} view */
function pdfCloseFind(view) {
  const d = view.dom;
  if (d.findBar.hidden) return;
  d.findBar.hidden = true;
  d.findBtn.setAttribute('aria-expanded', 'false');
  view.eventBus?.dispatch('findbarclose', { source: view });
  d.container.focus({ preventScroll: true });
}

/**
 * Runs a search in the PDFFindController ('' = the query changed, 'again' = next / previous match,
 * 'highlightallchange' = highlight the current matches again without moving).
 * @param {PdfView} view
 * @param {'' | 'again' | 'casesensitivitychange' | 'highlightallchange'} type
 * @param {boolean} findPrevious
 */
function pdfFind(view, type, findPrevious) {
  if (!view.eventBus) return;
  view.eventBus.dispatch('find', {
    source: view,
    type,
    query: view.dom.findInput.value,
    caseSensitive: view.matchCase,
    entireWord: false,
    highlightAll: true,
    findPrevious,
    matchDiacritics: false,
  });
}

/**
 * "3 of 12" / "No results" next to the find field.
 * @param {PdfView} view
 * @param {number | null} state PDF_FIND_STATE, or null for a count update only
 * @param {{ current: number, total: number } | undefined} matchesCount
 */
function pdfUpdateFindStatus(view, state, matchesCount) {
  if (view.disposed) return;
  const d = view.dom;
  if (state !== null) view.findState = state;
  const total = matchesCount?.total ?? 0;
  const current = matchesCount?.current ?? 0;
  let text = '';
  if (d.findInput.value) {
    if (total > 0) text = current > 0 ? `${current} of ${total}` : `${total} found`;
    else if (view.findState === PDF_FIND_STATE.PENDING) text = 'Searching…';
    else if (view.findState === PDF_FIND_STATE.NOT_FOUND) text = 'No results';
  }
  d.findStatus.textContent = text;
  d.findInput.classList.toggle('pdf-find-none', text === 'No results');
}

// ----- PDF: OUTLINE -----

/** @param {PdfView} view */
async function pdfLoadOutline(view) {
  /** @type {any[] | null} */
  let outline = null;
  try {
    outline = await view.pdfDocument.getOutline();
  } catch (err) {
    log('warn', `The outline of ${view.fileName} could not be read: ${pdfErrorDetail(err)}`);
  }
  if (view.disposed) return;
  const d = view.dom;
  if (!Array.isArray(outline) || outline.length === 0) {
    d.outlineBtn.disabled = true;
    d.outlineBtn.title = 'Outline (this PDF has none)';
    return;
  }
  d.outline.replaceChildren(pdfBuildOutlineList(view, outline, 0));
  d.outlineBtn.disabled = false;
  if (view.saved?.outline) pdfToggleOutline(view, true);
}

/**
 * Nested list of outline entries; an entry with children gets an expand / collapse button (collapsed when the PDF
 * says so: a negative /Count).
 * @param {PdfView} view
 * @param {any[]} items pdf.js outline nodes
 * @param {number} depth
 * @returns {HTMLElement}
 */
function pdfBuildOutlineList(view, items, depth) {
  const list = h('ul', { class: 'pdf-outline-list', role: depth === 0 ? null : 'group' });
  for (const item of items) {
    if (!item || typeof item !== 'object') continue;
    const title = String(item.title ?? '').replace(/\0/g, '').trim() || 'Untitled';
    const classes = ['pdf-outline-link', item.bold ? 'pdf-outline-bold' : '', item.italic ? 'pdf-outline-italic' : ''].filter(Boolean).join(' ');
    const link = h('button', { class: classes, type: 'button', title, onclick: () => pdfFollowOutlineItem(view, item) }, title);
    const children = Array.isArray(item.items) && item.items.length > 0 && depth + 1 < PDF_MAX_OUTLINE_DEPTH ? item.items : null;
    const li = h('li', { class: 'pdf-outline-item' });
    if (children) {
      const sub = pdfBuildOutlineList(view, children, depth + 1);
      const expanded = !(typeof item.count === 'number' && item.count < 0);
      sub.hidden = !expanded;
      const toggle = h(
        'button',
        {
          class: 'fv-icon-btn pdf-outline-toggle',
          type: 'button',
          // Its own name: the entry's button next to it (which goes to the entry) is named after the title.
          'aria-label': `${expanded ? 'Collapse' : 'Expand'} ${title}`,
          'aria-expanded': String(expanded),
          onclick: () => {
            sub.hidden = !sub.hidden;
            toggle.setAttribute('aria-expanded', String(!sub.hidden));
            toggle.setAttribute('aria-label', `${sub.hidden ? 'Expand' : 'Collapse'} ${title}`);
          },
        },
        icon(ICONS.chevronRight),
      );
      li.append(h('div', { class: 'pdf-outline-row' }, toggle, link), sub);
    } else {
      li.append(h('div', { class: 'pdf-outline-row' }, h('span', { class: 'pdf-outline-indent' }), link));
    }
    list.append(li);
  }
  return list;
}

/**
 * Tree keys in the outline, like a VS Code tree: Up / Down previous / next entry shown, Right expands (or goes to the
 * first child), Left collapses (or goes to the parent), Home / End first / last entry shown.
 * @param {KeyboardEvent} e
 */
function pdfOnOutlineKeyDown(e) {
  if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || e.isComposing) return;
  const outline = e.currentTarget instanceof HTMLElement ? e.currentTarget : null;
  const li = e.target instanceof HTMLElement ? e.target.closest('.pdf-outline-item') : null;
  if (!outline || !li) return;
  /** @param {Element | null | undefined} item */
  const linkOf = (item) => /** @type {HTMLElement | null} */ (item?.querySelector(':scope > .pdf-outline-row > .pdf-outline-link') ?? null);
  // Entries inside a collapsed (hidden) list have no layout box.
  const links = [.../** @type {NodeListOf<HTMLElement>} */ (outline.querySelectorAll('.pdf-outline-link'))].filter((el) => el.getClientRects().length > 0);
  const i = links.indexOf(/** @type {HTMLElement} */ (linkOf(li)));
  const toggle = /** @type {HTMLElement | null} */ (li.querySelector(':scope > .pdf-outline-row > .pdf-outline-toggle'));
  const sub = /** @type {HTMLElement | null} */ (li.querySelector(':scope > ul'));
  /** @type {HTMLElement | null | undefined} */
  let next = null;
  switch (e.key) {
    case 'ArrowDown':
      next = links[i + 1];
      break;
    case 'ArrowUp':
      next = i > 0 ? links[i - 1] : null;
      break;
    case 'Home':
      next = links[0];
      break;
    case 'End':
      next = links[links.length - 1];
      break;
    case 'ArrowRight':
      if (sub && toggle && sub.hidden) toggle.click();
      else if (sub) next = linkOf(sub.querySelector(':scope > .pdf-outline-item'));
      break;
    case 'ArrowLeft':
      if (sub && toggle && !sub.hidden) toggle.click();
      else next = linkOf(li.parentElement?.closest('.pdf-outline-item'));
      break;
    default:
      return;
  }
  e.preventDefault();
  e.stopPropagation();
  next?.focus();
}

/**
 * @param {PdfView} view
 * @param {any} item pdf.js outline node: dest (in-document), url (external), action (named) or setOCGState
 */
function pdfFollowOutlineItem(view, item) {
  const ls = view.linkService;
  if (!ls || view.disposed) return;
  if (item.dest) {
    Promise.resolve(ls.goToDestination(item.dest)).catch((/** @type {unknown} */ err) =>
      log('warn', `Outline entry "${String(item.title ?? '')}" points to a missing destination: ${pdfErrorDetail(err)}`),
    );
  } else if (typeof item.url === 'string' && item.url) {
    post({ type: 'openLink', href: pdfLinkHref(item.url, view.baseUrl) });
  } else if (item.action) {
    ls.executeNamedAction(item.action);
  } else if (item.setOCGState) {
    void ls.executeSetOCGState(item.setOCGState);
  }
}

/**
 * @param {PdfView} view
 * @param {boolean} [open] default: toggle
 */
function pdfToggleOutline(view, open) {
  const d = view.dom;
  const show = open ?? Boolean(d.sidebar.hidden);
  if (show && d.outlineBtn.disabled) return;
  d.sidebar.hidden = !show;
  d.outlineBtn.setAttribute('aria-pressed', String(show));
  view.outlineOpen = show;
  view.saveSoon();
}

// ----- PDF: LINKS -----

/**
 * @param {PdfView} view
 * @param {MouseEvent} event click / auxclick in the pages
 */
function pdfOnLinkClick(view, event) {
  if (event.type === 'auxclick' && event.button !== 1) return;
  const target = event.target instanceof Element ? event.target : null;
  const link = target?.closest('a');
  if (!(link instanceof HTMLAnchorElement) || !view.dom.container.contains(link)) return;
  event.stopPropagation();
  if (event.defaultPrevented) return; // pdf.js followed it (destination, named action)
  event.preventDefault();
  const href = (link.getAttribute('href') || '').trim();
  if (!href || href.startsWith('#')) return;
  post({ type: 'openLink', href: pdfLinkHref(link.href, view.baseUrl) });
}

/**
 * openLink href for a link pdf.js resolved against the PDF's webview URI: a target in the same resource origin (a
 * relative link to a file next to the PDF, a GoToR action) becomes a path relative to the PDF's folder again, which
 * the host resolves and opens in VS Code; other links (http, https, mailto) are sent as they are.
 * @param {string} href absolute URL
 * @param {string | null} baseUrl
 * @returns {string}
 */
function pdfLinkHref(href, baseUrl) {
  if (!baseUrl) return href;
  let url;
  let base;
  try {
    url = new URL(href);
    base = new URL(baseUrl);
  } catch {
    return href;
  }
  if (url.origin === 'null' || url.origin !== base.origin) return href;
  const from = base.pathname.split('/').slice(0, -1);
  const to = url.pathname.split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
  const relative = [...Array(from.length - common).fill('..'), ...to.slice(common)].join('/');
  return relative + url.hash;
}

// ===== PPTX =====
// Read-only presentation view: thumbnail strip (virtualised; hidden slides marked) + the current slide scaled to fit
// (zoom −/+/fit) + speaker notes. The host sends the deck outline at init (PptxDeckMeta) and each slide on request
// (getSlide → 'slide'), so a deck with many large pictures is never sent at once.
//
// A slide is built once in POINTS (CSS `pt`: pptxtojson's text HTML uses pt font sizes) as absolutely positioned
// elements, then scaled with a CSS transform: zooming only changes the scale factor. Thumbnails use the same builder.
//   - text / shape HTML: sanitize() (DOCX policy: file: links kept for openLink), &nbsp; → spaces so text wraps;
//   - shapes: inline SVG in points; pptxtojson's preset / custom `path` (scaled from `pathViewBox`), own geometry for
//     the common presets when it is missing, rect otherwise; lines and connectors with head / tail ends;
//   - fills: colour, gradient, picture, pattern (≈ a colour); borders with width and dash; rotation and flips;
//   - images (crop, shape clip, border), tables (merges, borders, fills), groups, SmartArt drawings, equations;
//   - charts: lazy import('echarts') (bar/column, line, area, pie, doughnut, scatter); other kinds → a labelled
//     placeholder with the data; media → a placeholder (never played or embedded).
// Links: a link to another slide of the deck (`ppt/slides/slideN.xml`) jumps there; anything else → openLink.

/** @typedef {import('../src/renderers/pptx').PptxDeckMeta} PptxDeckMeta */
/** @typedef {import('../src/renderers/pptx').PptxSlideInfo} PptxSlideInfo */
/** @typedef {import('../src/renderers/pptx').PptxSlide} PptxSlide */
/** @typedef {Extract<HostMessage, { type: 'init', kind: 'pptx' }>} PptxInitMessage */
/** @typedef {Extract<HostMessage, { type: 'slide' }>} PptxSlideMessage */
/**
 * A slide element as the host sends it (PptxElement: pptxtojson objects, normalised). Untrusted data: every field is
 * checked where it is used.
 * @typedef {Record<string, any>} PptxEl
 */
/** Box of an element in its parent's coordinates (points), after the parent group's flips. */
/** @typedef {{ left: number, top: number, width: number, height: number, rotate: number, flipH: boolean, flipV: boolean }} PptxPlace */
/** A flipped group mirrors its children's boxes (their text is not mirrored, as in PowerPoint). */
/** @typedef {{ w: number, h: number, flipH: boolean, flipV: boolean }} PptxMirror */
/** @typedef {{ el: PptxEl, host: HTMLElement, kind: string, width: number, height: number, dark: boolean }} PptxChartJob */
/**
 * @typedef {object} PptxBuildContext
 * @property {number} index slide index (for log messages)
 * @property {boolean} thumb building a thumbnail: no charts, no links, lighter placeholders
 * @property {boolean} dark the slide background is dark (chart text colours)
 * @property {PptxChartJob[]} charts charts to draw once the slide is in the document
 * @property {boolean} failed an element could not be drawn (logged once per slide)
 */
/**
 * @typedef {object} PptxDom
 * @property {HTMLElement} app
 * @property {HTMLInputElement} pageInput
 * @property {HTMLElement} pageTotal
 * @property {HTMLElement} title
 * @property {HTMLElement} hiddenBadge
 * @property {HTMLButtonElement} thumbsBtn
 * @property {HTMLButtonElement} prevBtn
 * @property {HTMLButtonElement} nextBtn
 * @property {HTMLButtonElement} zoomOutBtn
 * @property {HTMLButtonElement} zoomInBtn
 * @property {HTMLElement} zoomLabel
 * @property {HTMLButtonElement} fitBtn
 * @property {HTMLButtonElement} notesBtn
 * @property {HTMLElement} thumbs scroller of the thumbnail strip (listbox)
 * @property {HTMLElement} thumbsInner full-height sizer the visible thumbnails are positioned in
 * @property {HTMLElement} stage scroller of the main slide
 * @property {HTMLElement} frame box of the scaled slide
 * @property {HTMLElement | null} slideEl the slide in the frame (or its loading placeholder)
 * @property {number} slideIndex index the frame shows (-1: loading placeholder)
 * @property {HTMLElement} notes
 * @property {HTMLElement} notesBody
 * @property {HTMLElement} live polite announcement of the current slide
 */

/** Persisted view state (vscode.setState): { slide, zoom: 'fit' | number, notes, thumbs }. */
const PPTX_STATE_KEY = 'pptxView';
/** CSS pixels per point (1pt = 1/72in, 1px = 1/96in). */
const PPTX_PX_PER_PT = 4 / 3;
const PPTX_ZOOM_STEPS = [0.1, 0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
const PPTX_MIN_ZOOM = PPTX_ZOOM_STEPS[0];
const PPTX_MAX_ZOOM = PPTX_ZOOM_STEPS[PPTX_ZOOM_STEPS.length - 1];
/** Width of a thumbnail picture (px) and the vertical space around it in the strip. */
const PPTX_THUMB_WIDTH = 152;
const PPTX_THUMB_GAP = 12;
const PPTX_THUMB_OVERSCAN = 3;
/** getSlide requests in flight at once (the host answers in order; the current slide may jump the queue). */
const PPTX_MAX_IN_FLIGHT = 3;
/** Slides kept in memory (pictures are data: URIs, so a slide can be large). */
const PPTX_SLIDE_CACHE_SIZE = 24;
/** Default text insets of PowerPoint (0.1in left/right, 0.05in top/bottom), in points. */
const PPTX_DEFAULT_INSET = Object.freeze({ l: 7.2, t: 3.6, r: 7.2, b: 3.6 });
const PPTX_SVG_NS = 'http://www.w3.org/2000/svg';
/** Office 2023 theme accents: chart colours when neither the chart nor the deck has any. */
const PPTX_DEFAULT_PALETTE = ['#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E'];
const PPTX_SANS_FALLBACK = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';
const PPTX_SERIF_FALLBACK = 'Cambria, Caladea, "Times New Roman", serif';
const PPTX_MONO_FALLBACK = 'Consolas, "Courier New", monospace';

/** Presets whose holes are sub-paths drawn in the same direction as the outline (filled even-odd). */
const PPTX_EVENODD_TYPES = new Set(['donut', 'noSmoking', 'frame']);

/** Presets drawn with the viewer's own geometry (pptxPresetPath) rather than pptxtojson's path. */
const PPTX_OWN_PRESETS = new Set([
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'rtTriangle',
  'diamond',
  'parallelogram',
  'trapezoid',
  'pentagon',
  'hexagon',
  'octagon',
  'chevron',
  'homePlate',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'wedgeRectCallout',
  'wedgeRoundRectCallout',
  'wedgeEllipseCallout',
]);

/** In-between stops per gradient segment, so the browser's sRGB blend follows PowerPoint's linear-light blend. */
const PPTX_GRADIENT_STEPS = 8;

/** Line-like presets: drawn as open paths with their own geometry (head / tail ends, never filled). */
const PPTX_LINE_TYPES = new Set([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
]);

/**
 * pptxtojson's dash patterns (one per DrawingML prstDash) → PowerPoint's own pattern in multiples of the line width.
 * Unknown patterns are used as they are, also scaled by the width.
 * @type {Record<string, number[]>}
 */
const PPTX_DASHES = {
  '5': [4, 3], // dash
  '5, 5, 1, 5': [4, 3, 1, 3], // dashDot
  '1, 5': [1, 3], // dot
  '10, 5': [8, 3], // lgDash
  '10, 5, 1, 5': [8, 3, 1, 3], // lgDashDot
  '10, 5, 1, 5, 1, 5': [8, 3, 1, 3, 1, 3], // lgDashDotDot
  '5, 2': [3, 1], // sysDash
  '5, 2, 1, 5': [3, 1, 1, 1], // sysDashDot
  '5, 2, 1, 5, 1, 5': [3, 1, 1, 1, 1, 1], // sysDashDotDot
  '2, 5': [1, 1], // sysDot
};

/** Arrowhead length / width in multiples of the line width (DrawingML sm / med / lg). */
/** @type {Record<string, number>} */
const PPTX_END_SIZES = { sm: 2, med: 3, lg: 5 };

/** @type {Record<string, string>} */
const PPTX_CHART_NAMES = {
  barChart: 'Bar chart',
  bar3DChart: '3-D bar chart',
  lineChart: 'Line chart',
  line3DChart: '3-D line chart',
  areaChart: 'Area chart',
  area3DChart: '3-D area chart',
  pieChart: 'Pie chart',
  pie3DChart: '3-D pie chart',
  doughnutChart: 'Doughnut chart',
  scatterChart: 'Scatter chart',
  bubbleChart: 'Bubble chart',
  radarChart: 'Radar chart',
  surfaceChart: 'Surface chart',
  surface3DChart: '3-D surface chart',
  stockChart: 'Stock chart',
};

/** Icons of this view (trusted, static markup). */
const PPTX_ICONS = {
  thumbnails:
    '<svg viewBox="0 0 16 16" width="16" height="16"><rect x="1.5" y="2.5" width="13" height="11" rx="1" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M5.5 2.5v11" stroke="currentColor" stroke-width="1.2"/><path d="M2.8 4.8h1.5M2.8 7h1.5M2.8 9.2h1.5" stroke="currentColor" stroke-width="1.1"/></svg>',
  minus: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M3.5 8h9" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
  plus: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M3.5 8h9M8 3.5v9" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
  fit: '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M2 5.5V2h3.5M10.5 2H14v3.5M14 10.5V14h-3.5M5.5 14H2v-3.5" fill="none" stroke="currentColor" stroke-width="1.3"/><rect x="5" y="5.5" width="6" height="5" fill="none" stroke="currentColor" stroke-width="1.1"/></svg>',
  notes:
    '<svg viewBox="0 0 16 16" width="16" height="16"><rect x="2.5" y="1.5" width="11" height="13" rx="1" fill="none" stroke="currentColor" stroke-width="1.2"/><path d="M5 5h6M5 7.8h6M5 10.6h4" stroke="currentColor" stroke-width="1.1" stroke-linecap="round"/></svg>',
  film: '<svg viewBox="0 0 16 16" width="16" height="16"><rect x="1.5" y="3" width="13" height="10" rx="1.2" fill="none" stroke="currentColor" stroke-width="1.1"/><path d="M6.5 5.8v4.4L10.3 8z" fill="currentColor"/></svg>',
  speaker:
    '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M2.5 6h2.3L8 3.3v9.4L4.8 10H2.5z" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round"/><path d="M10.4 5.6a3.4 3.4 0 0 1 0 4.8M12 4a5.6 5.6 0 0 1 0 8" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linecap="round"/></svg>',
  chart:
    '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M2 14h12" stroke="currentColor" stroke-width="1.1"/><path d="M3.5 13V8.5h2.2V13M7 13V4h2.2v9M10.5 13V6.5h2.2V13" fill="none" stroke="currentColor" stroke-width="1.1"/></svg>',
};

/** Ids of getSlide requests: never reused, so a late reply for an earlier deck (before a reload) matches nothing. */
let pptxNextReqId = 1;
/** Counter for ids inside the slide SVGs (gradients, clip paths, markers): unique in the whole page. */
let pptxUid = 0;
/** @type {Promise<typeof import('echarts')> | null} */
let pptxEchartsPromise = null;

const pptxState = {
  /** @type {PptxDeckMeta | null} */
  deck: null,
  fileName: '',
  current: 0,
  /** @type {'fit' | number} */
  zoom: /** @type {'fit' | number} */ ('fit'),
  /** Scale actually applied (fit resolved), 1 = 100 %. */
  scale: 1,
  notesOpen: false,
  thumbsOpen: true,
  /** @type {Map<number, number>} reqId → slide index of the getSlide requests in flight */
  inFlight: new Map(),
  /** @type {Set<number>} slide indexes in flight */
  requested: new Set(),
  /** @type {number[]} slides to fetch first (the current one) */
  urgent: [],
  /** @type {number[]} slides to fetch when nothing is urgent (neighbours, visible thumbnails) */
  background: [],
  /** @type {Map<number, PptxSlide>} LRU: the most recently used slide last */
  cache: new Map(),
  /** @type {PptxDom | null} */
  dom: null,
  /** @type {import('echarts').EChartsType[]} charts drawn on the main slide */
  charts: [],
  /** Increments with every main-slide build; a chart job of an older build is dropped. */
  renderToken: 0,
  /** @type {Map<number, HTMLElement>} thumbnail items currently in the strip */
  thumbItems: new Map(),
  thumbItemHeight: 100,
  thumbScale: 0.1,
};

/**
 * `init` for kind 'pptx'.
 * @param {PptxInitMessage} msg
 */
function showPptx(msg) {
  try {
    const deck = pptxCheckDeck(msg.deck);
    const saved = pptxReadState();
    const s = pptxState;
    s.deck = deck;
    s.fileName = typeof msg.fileName === 'string' ? msg.fileName : '';
    s.current = clamp(saved.slide ?? 0, 0, Math.max(0, deck.slideCount - 1));
    s.zoom = saved.zoom ?? 'fit';
    s.notesOpen = saved.notes ?? deck.slides.some((info) => info.hasNotes);
    s.thumbsOpen = saved.thumbs ?? (root.clientWidth === 0 || root.clientWidth >= 560);
    s.inFlight = new Map();
    s.requested = new Set();
    s.urgent = [];
    s.background = [];
    s.cache = new Map();
    s.charts = [];
    s.thumbItems = new Map();
    s.renderToken++;

    const dom = pptxBuildApp(msg, deck);
    s.dom = dom;
    onDispose(() => {
      pptxDisposeCharts();
      s.renderToken++;
      s.deck = null;
      s.dom = null;
      s.cache = new Map();
      s.thumbItems = new Map();
      s.inFlight = new Map();
      s.requested = new Set();
      s.urgent = [];
      s.background = [];
    });
    root.replaceChildren(dom.app);
    pptxWireEvents(dom);

    if (deck.slideCount === 0) {
      dom.stage.replaceChildren(h('p', { class: 'fv-empty pptx-empty' }, 'This presentation has no slides.'));
      pptxUpdateChrome();
      return;
    }
    pptxApplyZoom();
    pptxRenderThumbs();
    pptxShowSlide(s.current, false);
    dom.app.focus({ preventScroll: true });
  } catch (err) {
    mdFail(err, 'Could not display the presentation.', false);
  }
}

/**
 * Reply to getSlide. Replies the view no longer waits for (another deck, a disposed view) are dropped.
 * @param {PptxSlideMessage} msg
 */
function onSlide(msg) {
  const s = pptxState;
  const index = s.inFlight.get(msg.reqId);
  if (index === undefined || !s.deck || !s.dom) {
    log('info', `Ignored a stale 'slide' reply (request ${String(msg.reqId)}).`);
    return;
  }
  s.inFlight.delete(msg.reqId);
  s.requested.delete(index);
  const slide = msg.slide;
  if (msg.index !== index || !slide || typeof slide !== 'object' || !Array.isArray(slide.elements)) {
    logError(`Ignored a malformed 'slide' reply for slide ${index + 1}`);
    pptxPump();
    return;
  }
  pptxCacheSlide(index, slide);
  if (index === s.current) pptxRenderMain();
  const item = s.thumbItems.get(index);
  if (item && !item.dataset.loaded) pptxFillThumb(item, slide, index);
  pptxPump();
}

// ----- PPTX: DECK & STATE -----

/**
 * The deck outline with defaults for anything missing or malformed (sizes in points).
 * @param {unknown} value
 * @returns {PptxDeckMeta}
 */
function pptxCheckDeck(value) {
  const deck = /** @type {Record<string, any>} */ (value && typeof value === 'object' ? value : {});
  let width = pptxNum(deck.width);
  let height = pptxNum(deck.height);
  if (!(width > 0 && height > 0)) {
    log('warn', 'The presentation has no valid slide size; using 16:9.');
    width = 960;
    height = 540;
  }
  const count = Number.isInteger(deck.slideCount) && deck.slideCount >= 0 ? deck.slideCount : Array.isArray(deck.slides) ? deck.slides.length : 0;
  const infos = Array.isArray(deck.slides) ? deck.slides : [];
  /** @type {PptxSlideInfo[]} */
  const slides = [];
  for (let i = 0; i < count; i++) {
    const info = infos[i] && typeof infos[i] === 'object' ? infos[i] : {};
    slides.push({
      index: i,
      title: typeof info.title === 'string' ? info.title : '',
      hidden: info.hidden === true,
      hasNotes: info.hasNotes === true,
      ...(typeof info.part === 'string' ? { part: info.part } : {}),
    });
  }
  return {
    width,
    height,
    slideCount: count,
    slides,
    themeColors: Array.isArray(deck.themeColors) ? deck.themeColors.filter((c) => typeof c === 'string') : [],
    fonts: Array.isArray(deck.fonts) ? deck.fonts.filter((f) => typeof f === 'string') : [],
    lossy: Array.isArray(deck.lossy) ? deck.lossy.filter((f) => typeof f === 'string') : [],
  };
}

/**
 * The persisted view state, each field checked.
 * @returns {{ slide?: number, zoom?: 'fit' | number, notes?: boolean, thumbs?: boolean }}
 */
function pptxReadState() {
  try {
    const value = getStateKey(PPTX_STATE_KEY);
    if (!value || typeof value !== 'object') return {};
    /** @type {{ slide?: number, zoom?: 'fit' | number, notes?: boolean, thumbs?: boolean }} */
    const out = {};
    if (Number.isInteger(value.slide) && value.slide >= 0) out.slide = value.slide;
    if (value.zoom === 'fit') out.zoom = 'fit';
    else if (typeof value.zoom === 'number' && Number.isFinite(value.zoom)) out.zoom = clamp(value.zoom, PPTX_MIN_ZOOM, PPTX_MAX_ZOOM);
    if (typeof value.notes === 'boolean') out.notes = value.notes;
    if (typeof value.thumbs === 'boolean') out.thumbs = value.thumbs;
    return out;
  } catch {
    return {};
  }
}

function pptxSaveState() {
  const s = pptxState;
  if (!s.deck) return;
  try {
    setStateKey(PPTX_STATE_KEY, { slide: s.current, zoom: s.zoom, notes: s.notesOpen, thumbs: s.thumbsOpen });
  } catch (err) {
    logError('Could not save the presentation view state', err);
  }
}

// ----- PPTX: SLIDE REQUESTS -----

/**
 * Asks the host for a slide unless it is cached or already on its way.
 * @param {number} index
 * @param {boolean} urgent the slide on screen: fetched before thumbnails and neighbours
 */
function pptxRequest(index, urgent) {
  const s = pptxState;
  if (!s.deck || !Number.isInteger(index) || index < 0 || index >= s.deck.slideCount) return;
  if (s.cache.has(index) || s.requested.has(index)) return;
  if (urgent) {
    s.urgent = [index, ...s.urgent.filter((i) => i !== index)];
    s.background = s.background.filter((i) => i !== index);
  } else if (!s.urgent.includes(index) && !s.background.includes(index)) {
    s.background.push(index);
  }
  pptxPump();
}

/**
 * Replaces the background queue: the current slide's neighbours, then the visible thumbnails still missing.
 * @param {number[]} thumbs
 */
function pptxWantBackground(thumbs) {
  const s = pptxState;
  if (!s.deck) return;
  const want = [s.current + 1, s.current - 1, ...thumbs];
  /** @type {number[]} */
  const queue = [];
  for (const i of want) {
    if (i < 0 || i >= s.deck.slideCount || s.cache.has(i) || s.requested.has(i) || queue.includes(i) || s.urgent.includes(i)) continue;
    queue.push(i);
  }
  s.background = queue;
  pptxPump();
}

function pptxPump() {
  const s = pptxState;
  if (!s.deck) return;
  // The slide on screen may use two extra slots, so it never waits behind a screenful of thumbnails.
  while (s.inFlight.size < PPTX_MAX_IN_FLIGHT || (s.urgent.length > 0 && s.inFlight.size < PPTX_MAX_IN_FLIGHT + 2)) {
    const index = s.urgent.length > 0 ? s.urgent.shift() : s.background.shift();
    if (index === undefined) break;
    if (s.cache.has(index) || s.requested.has(index)) continue;
    const reqId = pptxNextReqId++;
    s.inFlight.set(reqId, index);
    s.requested.add(index);
    post({ type: 'getSlide', reqId, index });
  }
}

/**
 * @param {number} index
 * @param {PptxSlide} slide
 */
function pptxCacheSlide(index, slide) {
  const cache = pptxState.cache;
  cache.delete(index);
  cache.set(index, slide);
  for (const key of cache.keys()) {
    if (cache.size <= PPTX_SLIDE_CACHE_SIZE) break;
    if (key !== pptxState.current) cache.delete(key);
  }
}

/**
 * Cached slide, marked as recently used.
 * @param {number} index
 * @returns {PptxSlide | undefined}
 */
function pptxCachedSlide(index) {
  const slide = pptxState.cache.get(index);
  if (slide) pptxCacheSlide(index, slide);
  return slide;
}

// ----- PPTX: VIEW CHROME -----

/**
 * @param {PptxInitMessage} msg
 * @param {PptxDeckMeta} deck
 * @returns {PptxDom}
 */
function pptxBuildApp(msg, deck) {
  const s = pptxState;
  /**
   * @param {string} markup
   * @param {string} label
   * @param {() => void} onclick
   */
  const button = (markup, label, onclick) =>
    /** @type {HTMLButtonElement} */ (h('button', { class: 'fv-icon-btn', type: 'button', title: label, 'aria-label': label, onclick }, icon(markup)));

  const thumbsBtn = button(PPTX_ICONS.thumbnails, 'Slide thumbnails', () => pptxToggleThumbs());
  const prevBtn = button(ICONS.chevronLeft, 'Previous slide (Left Arrow / Page Up)', () => pptxGoTo(pptxState.current - 1));
  const nextBtn = button(ICONS.chevronRight, 'Next slide (Right Arrow / Page Down)', () => pptxGoTo(pptxState.current + 1));
  const pageInput = /** @type {HTMLInputElement} */ (
    h('input', {
      class: 'pptx-page-input',
      type: 'text',
      inputmode: 'numeric',
      autocomplete: 'off',
      spellcheck: 'false',
      'aria-label': 'Slide number',
      title: 'Go to slide (Enter)',
    })
  );
  const pageTotal = h('span', { class: 'pptx-page-total' }, `/ ${deck.slideCount}`);
  const title = h('span', { class: 'pptx-title' });
  const hiddenBadge = h('span', { class: 'pptx-hidden-badge', hidden: true, title: 'This slide is hidden in the slide show' }, icon(ICONS.eyeClosed), 'Hidden');
  const zoomOutBtn = button(PPTX_ICONS.minus, `Zoom out (${IS_MAC ? 'Cmd' : 'Ctrl'}+-)`, () => pptxZoomBy(-1));
  const zoomInBtn = button(PPTX_ICONS.plus, `Zoom in (${IS_MAC ? 'Cmd' : 'Ctrl'}+=)`, () => pptxZoomBy(1));
  const zoomLabel = h('span', { class: 'pptx-zoom-label', title: 'Zoom' }, '100%');
  const fitBtn = button(PPTX_ICONS.fit, `Fit slide to window (${IS_MAC ? 'Cmd' : 'Ctrl'}+0)`, () => pptxSetZoom('fit'));
  const notesBtn = button(PPTX_ICONS.notes, 'Speaker notes', () => pptxToggleNotes());

  const toolbar = h(
    'div',
    { class: 'pptx-toolbar', role: 'toolbar', 'aria-label': 'Presentation' },
    thumbsBtn,
    h('span', { class: 'pptx-sep', 'aria-hidden': 'true' }),
    prevBtn,
    h('span', { class: 'pptx-page' }, pageInput, pageTotal),
    nextBtn,
    h('span', { class: 'pptx-title-wrap' }, title, hiddenBadge),
    zoomOutBtn,
    zoomLabel,
    zoomInBtn,
    fitBtn,
    h('span', { class: 'pptx-sep', 'aria-hidden': 'true' }),
    notesBtn,
  );

  const thumbsInner = h('div', { class: 'pptx-thumbs-inner' });
  const thumbs = h('div', { class: 'pptx-thumbs', role: 'listbox', tabindex: '0', 'aria-label': 'Slides' }, thumbsInner);
  const frame = h('div', { class: 'pptx-frame' });
  const stage = h('div', { class: 'pptx-stage', role: 'region', 'aria-label': 'Slide' }, frame);
  const notesBody = h('div', { class: 'pptx-notes-body', tabindex: '0', 'aria-label': 'Speaker notes' });
  const notes = h('div', { class: 'pptx-notes', role: 'region', 'aria-label': 'Speaker notes' }, h('div', { class: 'pptx-notes-title' }, 'Notes'), notesBody);
  const live = h('span', { class: 'fv-sr-only', 'aria-live': 'polite' });
  const app = h(
    'div',
    { class: 'pptx-app', tabindex: '-1', 'aria-label': msg.fileName || null },
    msg.banner ? buildBanner(String(msg.banner), null) : null,
    toolbar,
    h('div', { class: 'pptx-body' }, thumbs, h('div', { class: 'pptx-main' }, stage, notes)),
    live,
  );

  // Thumbnail metrics: the picture keeps the slide's aspect ratio (very tall or wide slides are bounded).
  const thumbHeight = clamp(Math.round((PPTX_THUMB_WIDTH * deck.height) / deck.width), 24, 400);
  s.thumbScale = Math.min(PPTX_THUMB_WIDTH / (deck.width * PPTX_PX_PER_PT), thumbHeight / (deck.height * PPTX_PX_PER_PT));
  s.thumbItemHeight = thumbHeight + PPTX_THUMB_GAP;
  thumbsInner.style.height = `${deck.slideCount * s.thumbItemHeight + PPTX_THUMB_GAP}px`;
  thumbs.style.setProperty('--pptx-thumb-w', `${PPTX_THUMB_WIDTH}px`);
  thumbs.style.setProperty('--pptx-thumb-h', `${thumbHeight}px`);
  thumbs.hidden = !s.thumbsOpen;
  notes.hidden = !s.notesOpen;

  return {
    app,
    pageInput,
    pageTotal,
    title,
    hiddenBadge,
    thumbsBtn,
    prevBtn,
    nextBtn,
    zoomOutBtn,
    zoomInBtn,
    zoomLabel,
    fitBtn,
    notesBtn,
    thumbs,
    thumbsInner,
    stage,
    frame,
    slideEl: null,
    slideIndex: -1,
    notes,
    notesBody,
    live,
  };
}

/** @param {PptxDom} dom */
function pptxWireEvents(dom) {
  listen(document, 'keydown', pptxOnKeyDown);
  listen(dom.stage, 'click', pptxOnStageClick);
  listen(dom.stage, 'auxclick', pptxOnStageClick);
  listen(dom.stage, 'wheel', pptxOnWheel, { passive: false });
  listen(dom.thumbs, 'click', pptxOnThumbClick);
  let thumbFrame = 0;
  listen(dom.thumbs, 'scroll', () => {
    if (thumbFrame) return;
    thumbFrame = requestAnimationFrame(() => {
      thumbFrame = 0;
      pptxRenderThumbs();
    });
  });
  onDispose(() => cancelAnimationFrame(thumbFrame));

  listen(dom.pageInput, 'keydown', (/** @type {KeyboardEvent} */ event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      const n = Number.parseInt(dom.pageInput.value.trim(), 10);
      if (Number.isInteger(n) && pptxState.deck && n >= 1 && n <= pptxState.deck.slideCount) pptxGoTo(n - 1);
      else pptxUpdateChrome();
      dom.pageInput.select();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      pptxUpdateChrome();
      dom.app.focus({ preventScroll: true });
    }
  });
  listen(dom.pageInput, 'focus', () => dom.pageInput.select());
  listen(dom.pageInput, 'blur', () => pptxUpdateChrome());

  // Fit follows the panel size; the thumbnail strip renders what became visible.
  let resizeFrame = 0;
  const observer = new ResizeObserver(() => {
    if (resizeFrame) return;
    resizeFrame = requestAnimationFrame(() => {
      resizeFrame = 0;
      if (pptxState.zoom === 'fit') pptxApplyZoom();
      pptxRenderThumbs();
    });
  });
  observer.observe(dom.stage);
  observer.observe(dom.thumbs);
  onDispose(() => {
    observer.disconnect();
    cancelAnimationFrame(resizeFrame);
  });
}

/** Toolbar, notes and announcements for the current slide. */
function pptxUpdateChrome() {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  const count = s.deck.slideCount;
  const info = s.deck.slides[s.current];
  if (document.activeElement !== dom.pageInput) dom.pageInput.value = count ? String(s.current + 1) : '0';
  dom.pageInput.size = Math.max(2, String(count).length);
  dom.prevBtn.disabled = count === 0 || s.current <= 0;
  dom.nextBtn.disabled = count === 0 || s.current >= count - 1;
  const title = info ? info.title.trim() || `Slide ${s.current + 1}` : '';
  dom.title.textContent = title;
  dom.title.title = title;
  dom.hiddenBadge.hidden = !info?.hidden;
  dom.thumbsBtn.setAttribute('aria-pressed', String(s.thumbsOpen));
  dom.notesBtn.setAttribute('aria-pressed', String(s.notesOpen));
  dom.stage.setAttribute('aria-label', count ? `Slide ${s.current + 1} of ${count}` : 'Slide');
}

/**
 * Speaker notes of the current slide ('' while it loads).
 * @param {PptxSlide | undefined} slide
 */
function pptxUpdateNotes(slide) {
  const dom = pptxState.dom;
  if (!dom) return;
  const text = slide && typeof slide.notes === 'string' ? slide.notes.replace(/\s+$/, '') : '';
  dom.notesBody.classList.toggle('pptx-notes-empty', !!slide && !text);
  dom.notesBody.textContent = slide ? text || 'No speaker notes.' : '';
}

/**
 * @param {number} index
 */
function pptxGoTo(index) {
  const s = pptxState;
  if (!s.deck || s.deck.slideCount === 0) return;
  const target = clamp(Math.trunc(index), 0, s.deck.slideCount - 1);
  if (target === s.current && s.dom?.slideIndex === target) return;
  pptxShowSlide(target, true);
}

/**
 * Makes `index` the current slide: toolbar, main slide (requested if needed), thumbnail selection, state.
 * @param {number} index
 * @param {boolean} announce tell screen readers (user navigation)
 */
function pptxShowSlide(index, announce) {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  s.current = index;
  pptxUpdateChrome();
  pptxSaveState();
  pptxRenderMain();
  pptxSyncThumbSelection(true);
  if (announce) {
    const info = s.deck.slides[index];
    dom.live.textContent = `Slide ${index + 1} of ${s.deck.slideCount}${info?.title ? `: ${info.title}` : ''}${info?.hidden ? ' (hidden)' : ''}`;
  }
}

/** Draws the current slide into the frame, or a loading page while it is fetched. */
function pptxRenderMain() {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  const index = s.current;
  const slide = pptxCachedSlide(index);
  if (!slide) {
    if (dom.slideIndex !== -1 || !dom.slideEl) {
      pptxDisposeCharts();
      s.renderToken++;
      const blank = pptxBlankSlide(s.deck);
      blank.classList.add('pptx-loading');
      blank.setAttribute('aria-busy', 'true');
      dom.frame.replaceChildren(blank);
      dom.slideEl = blank;
      dom.slideIndex = -1;
      pptxApplyZoom();
    }
    pptxUpdateNotes(undefined);
    pptxRequest(index, true);
    return;
  }
  if (dom.slideIndex === index && dom.slideEl) return;

  pptxDisposeCharts();
  const token = ++s.renderToken;
  /** @type {PptxBuildContext} */
  const ctx = { index, thumb: false, dark: pptxIsDarkFill(slide.background), charts: [], failed: false };
  const el = pptxBuildSlide(slide, ctx);
  dom.frame.replaceChildren(el);
  dom.slideEl = el;
  dom.slideIndex = index;
  pptxApplyZoom();
  pptxUpdateNotes(slide);
  if (ctx.charts.length) void pptxRunCharts(ctx.charts, () => token === s.renderToken);
  pptxWantBackground(pptxMissingThumbs());
}

function pptxDisposeCharts() {
  const charts = pptxState.charts;
  pptxState.charts = [];
  for (const chart of charts) {
    try {
      chart.dispose();
    } catch (err) {
      logError('Could not dispose a chart', err);
    }
  }
}

function pptxToggleNotes() {
  const s = pptxState;
  if (!s.dom) return;
  s.notesOpen = !s.notesOpen;
  s.dom.notes.hidden = !s.notesOpen;
  pptxUpdateChrome();
  pptxSaveState();
  if (s.zoom === 'fit') pptxApplyZoom();
}

function pptxToggleThumbs() {
  const s = pptxState;
  if (!s.dom) return;
  s.thumbsOpen = !s.thumbsOpen;
  s.dom.thumbs.hidden = !s.thumbsOpen;
  pptxUpdateChrome();
  pptxSaveState();
  if (s.zoom === 'fit') pptxApplyZoom();
  pptxRenderThumbs();
  pptxSyncThumbSelection(true);
}

// ----- PPTX: ZOOM -----

/** Sizes the frame and scales the slide for the current zoom ('fit': as large as the stage allows). */
function pptxApplyZoom() {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  const slideW = s.deck.width * PPTX_PX_PER_PT;
  const slideH = s.deck.height * PPTX_PX_PER_PT;
  let scale;
  if (s.zoom === 'fit') {
    const style = getComputedStyle(dom.stage);
    const availW = dom.stage.clientWidth - (parseFloat(style.paddingLeft) || 0) - (parseFloat(style.paddingRight) || 0);
    const availH = dom.stage.clientHeight - (parseFloat(style.paddingTop) || 0) - (parseFloat(style.paddingBottom) || 0);
    scale = Math.min(availW / slideW, availH / slideH);
    if (!(scale > 0)) scale = s.scale > 0 ? s.scale : 1; // not laid out yet (hidden panel)
    scale = clamp(scale, 0.02, PPTX_MAX_ZOOM);
  } else {
    scale = clamp(s.zoom, PPTX_MIN_ZOOM, PPTX_MAX_ZOOM);
  }
  s.scale = scale;
  // Floor: a fitted slide must never be a fraction of a pixel too large (scroll bars would appear and re-fit).
  dom.frame.style.width = `${Math.max(1, Math.floor(slideW * scale))}px`;
  dom.frame.style.height = `${Math.max(1, Math.floor(slideH * scale))}px`;
  if (dom.slideEl) dom.slideEl.style.transform = `scale(${scale})`;
  dom.zoomLabel.textContent = `${Math.round(scale * 100)}%`;
  dom.fitBtn.setAttribute('aria-pressed', String(s.zoom === 'fit'));
  dom.zoomOutBtn.disabled = s.deck.slideCount === 0 || scale <= PPTX_MIN_ZOOM + 1e-6;
  dom.zoomInBtn.disabled = s.deck.slideCount === 0 || scale >= PPTX_MAX_ZOOM - 1e-6;
}

/**
 * Sets the zoom, keeping the point at the centre of the stage in place.
 * @param {'fit' | number} zoom
 */
function pptxSetZoom(zoom) {
  const s = pptxState;
  const dom = s.dom;
  if (!dom) return;
  const stage = dom.stage;
  const cx = stage.scrollWidth ? (stage.scrollLeft + stage.clientWidth / 2) / stage.scrollWidth : 0.5;
  const cy = stage.scrollHeight ? (stage.scrollTop + stage.clientHeight / 2) / stage.scrollHeight : 0.5;
  s.zoom = zoom === 'fit' ? 'fit' : clamp(zoom, PPTX_MIN_ZOOM, PPTX_MAX_ZOOM);
  pptxApplyZoom();
  stage.scrollLeft = cx * stage.scrollWidth - stage.clientWidth / 2;
  stage.scrollTop = cy * stage.scrollHeight - stage.clientHeight / 2;
  pptxSaveState();
}

/** @param {number} direction 1: in, -1: out */
function pptxZoomBy(direction) {
  const current = pptxState.scale;
  const next =
    direction > 0
      ? PPTX_ZOOM_STEPS.find((z) => z > current * 1.001) ?? PPTX_MAX_ZOOM
      : [...PPTX_ZOOM_STEPS].reverse().find((z) => z < current * 0.999) ?? PPTX_MIN_ZOOM;
  pptxSetZoom(next);
}

/** @param {WheelEvent} event */
function pptxOnWheel(event) {
  if (!(IS_MAC ? event.metaKey : event.ctrlKey) || event.deltaY === 0) return;
  event.preventDefault();
  pptxZoomBy(event.deltaY < 0 ? 1 : -1);
}

// ----- PPTX: KEYBOARD & LINKS -----

/** @param {KeyboardEvent} event */
function pptxOnKeyDown(event) {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck || event.defaultPrevented) return;
  const target = event.target instanceof Element ? event.target : null;
  if (target && !dom.app.contains(target) && target !== document.body) return;
  const editable = !!target && !!target.closest('input, textarea, select, [contenteditable="true"]');
  const mod = IS_MAC ? event.metaKey : event.ctrlKey;

  if (mod && !event.altKey) {
    if (event.key === '+' || event.key === '=') pptxZoomBy(1);
    else if (event.key === '-' || event.key === '_') pptxZoomBy(-1);
    else if (event.key === '0') pptxSetZoom('fit');
    else if ((event.key === 'a' || event.key === 'A') && !editable && !event.shiftKey) {
      // Select all = the text of the current slide (or of the notes when they have the focus), not the toolbar.
      const scope = target && dom.notes.contains(target) ? dom.notesBody : dom.slideEl;
      if (!scope) return;
      window.getSelection()?.selectAllChildren(scope);
    } else return;
    event.preventDefault();
    return;
  }
  if (event.altKey || event.ctrlKey || event.metaKey || editable || s.deck.slideCount === 0) return;

  // Arrows scroll a zoomed-in slide or the notes; Page Up / Page Down always change the slide.
  const inNotes = !!target && dom.notes.contains(target);
  const inThumbs = !!target && dom.thumbs.contains(target);
  const stage = dom.stage;
  const scrollsX = !inThumbs && stage.scrollWidth > stage.clientWidth + 1;
  const scrollsY = !inThumbs && stage.scrollHeight > stage.clientHeight + 1;
  let next;
  switch (event.key) {
    case 'ArrowLeft':
      if (inNotes || scrollsX) return;
      next = s.current - 1;
      break;
    case 'ArrowRight':
      if (inNotes || scrollsX) return;
      next = s.current + 1;
      break;
    case 'ArrowUp':
      if (inNotes || scrollsY) return;
      next = s.current - 1;
      break;
    case 'ArrowDown':
      if (inNotes || scrollsY) return;
      next = s.current + 1;
      break;
    case 'PageUp':
      next = s.current - 1;
      break;
    case 'PageDown':
      next = s.current + 1;
      break;
    case 'Home':
      if (inNotes) return;
      next = 0;
      break;
    case 'End':
      if (inNotes) return;
      next = s.deck.slideCount - 1;
      break;
    default:
      return;
  }
  event.preventDefault();
  pptxGoTo(next);
}

/**
 * Links on the main slide: text links (`<a href>`) and shapes / pictures with a hyperlink.
 * @param {MouseEvent} event
 */
function pptxOnStageClick(event) {
  if (event.type === 'auxclick' && event.button !== 1) return;
  const dom = pptxState.dom;
  const target = event.target instanceof Element ? event.target : null;
  const link = target?.closest('a, [data-pptx-link]');
  if (!dom || !link || !dom.slideEl || !dom.slideEl.contains(link)) return;
  event.preventDefault();
  event.stopPropagation();
  // A drag that selected text inside the link is a selection, not a click.
  const selection = window.getSelection();
  if (selection && !selection.isCollapsed && selection.anchorNode && link.contains(selection.anchorNode)) return;
  const href = (link.getAttribute('href') ?? link.getAttribute('data-pptx-link') ?? '').trim();
  pptxFollowLink(href);
}

/**
 * `slide8.xml` / `ppt/slides/slide8.xml` (a link to another slide: its package part) → that slide;
 * `#…` → nothing; anything else → the host (openLink checks the scheme and local files).
 * @param {string} href
 */
function pptxFollowLink(href) {
  const s = pptxState;
  // A slide jump action of a text run comes as '#ppaction://hlinkshowjump?jump=…' (the sanitizer keeps '#' links).
  if (/^#ppaction:/i.test(href)) href = href.slice(1);
  if (!href || href.startsWith('#') || !s.deck) return;
  const slideRef = /(?:^|[\\/])(slide(\d+)\.xml)$/i.exec(href);
  if (slideRef && !/^[a-z][a-z0-9+.-]*:/i.test(href)) {
    const index = pptxSlideIndexOfPart(s.deck, slideRef[1], Number(slideRef[2]) - 1);
    if (index >= 0 && index < s.deck.slideCount) pptxGoTo(index);
    return;
  }
  const jump = /^ppaction:\/\/hlinkshowjump\?jump=(\w+)/i.exec(href);
  if (jump) {
    const to = jump[1].toLowerCase();
    if (to === 'nextslide') pptxGoTo(s.current + 1);
    else if (to === 'previousslide') pptxGoTo(s.current - 1);
    else if (to === 'firstslide') pptxGoTo(0);
    else if (to === 'lastslide') pptxGoTo(s.deck.slideCount - 1);
    return;
  }
  if (/^ppaction:/i.test(href)) return;
  post({ type: 'openLink', href });
}

/**
 * Deck position of the slide stored in the part named `fileName` ('slide8.xml'). The host gives each slide's part
 * (PptxSlideInfo.part) because the deck order (presentation.xml) can differ from the part numbers; without it the
 * slides are in part-number order, so `fallback` (the number - 1) is used.
 * @param {PptxDeckMeta} deck
 * @param {string} fileName
 * @param {number} fallback
 * @returns {number} -1 when no slide of the deck is stored in that part
 */
function pptxSlideIndexOfPart(deck, fileName, fallback) {
  if (!deck.slides.some((info) => typeof info.part === 'string')) return fallback;
  const name = fileName.toLowerCase();
  return deck.slides.findIndex((info) => typeof info.part === 'string' && info.part.toLowerCase().split('/').pop() === name);
}

// ----- PPTX: THUMBNAILS -----

/** Renders the thumbnails in (and near) the visible part of the strip; requests the slides they still need. */
function pptxRenderThumbs() {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  if (!s.thumbsOpen || s.deck.slideCount === 0) {
    for (const item of s.thumbItems.values()) item.remove();
    s.thumbItems.clear();
    pptxWantBackground([]);
    return;
  }
  const { first, last } = pptxVisibleThumbRange(PPTX_THUMB_OVERSCAN);
  for (const [i, item] of s.thumbItems) {
    if (i < first || i > last) {
      item.remove();
      s.thumbItems.delete(i);
    }
  }
  for (let i = first; i <= last; i++) {
    let item = s.thumbItems.get(i);
    if (!item) {
      item = pptxBuildThumbItem(i);
      s.thumbItems.set(i, item);
      dom.thumbsInner.appendChild(item);
    }
    if (!item.dataset.loaded) {
      const slide = s.cache.get(i);
      if (slide) pptxFillThumb(item, slide, i);
    }
  }
  pptxWantBackground(pptxMissingThumbs());
}

/**
 * @param {number} overscan extra items above and below the viewport
 * @returns {{ first: number, last: number }}
 */
function pptxVisibleThumbRange(overscan) {
  const s = pptxState;
  const dom = s.dom;
  const count = s.deck ? s.deck.slideCount : 0;
  if (!dom || count === 0) return { first: 0, last: -1 };
  const itemH = s.thumbItemHeight;
  const viewH = dom.thumbs.clientHeight || 600;
  const top = dom.thumbs.scrollTop;
  return {
    first: clamp(Math.floor(top / itemH) - overscan, 0, count - 1),
    last: clamp(Math.ceil((top + viewH) / itemH) + overscan, 0, count - 1),
  };
}

/** Visible thumbnails whose slide is not loaded yet, in screen order. */
function pptxMissingThumbs() {
  const s = pptxState;
  if (!s.thumbsOpen) return [];
  const { first, last } = pptxVisibleThumbRange(1);
  /** @type {number[]} */
  const out = [];
  for (let i = first; i <= last; i++) if (!s.cache.has(i) && !s.thumbItems.get(i)?.dataset.loaded) out.push(i);
  return out;
}

/**
 * @param {number} index
 * @returns {HTMLElement}
 */
function pptxBuildThumbItem(index) {
  const s = pptxState;
  const info = s.deck?.slides[index];
  const hidden = !!info?.hidden;
  const label = `Slide ${index + 1}${info?.title ? `: ${info.title}` : ''}${hidden ? ' (hidden)' : ''}`;
  const item = h(
    'div',
    {
      class: `pptx-thumb${hidden ? ' pptx-thumb-hidden' : ''}${index === s.current ? ' pptx-thumb-current' : ''}`,
      id: `pptx-thumb-${index}`,
      role: 'option',
      'aria-selected': String(index === s.current),
      // Only the visible part of the list is in the DOM: give each option its real position in the deck.
      'aria-posinset': String(index + 1),
      'aria-setsize': String(s.deck?.slideCount ?? 0),
      'aria-label': label,
      title: label,
      dataset: { index: String(index) },
    },
    h('span', { class: 'pptx-thumb-num', 'aria-hidden': 'true' }, String(index + 1), hidden ? icon(ICONS.eyeClosed) : null),
    h('div', { class: 'pptx-thumb-pic' }),
  );
  item.style.top = `${PPTX_THUMB_GAP + index * s.thumbItemHeight}px`;
  return item;
}

/**
 * Puts the miniature of `slide` into a thumbnail item.
 * @param {HTMLElement} item
 * @param {PptxSlide} slide
 * @param {number} index
 */
function pptxFillThumb(item, slide, index) {
  const pic = item.querySelector('.pptx-thumb-pic');
  if (!pic) return;
  /** @type {PptxBuildContext} */
  const ctx = { index, thumb: true, dark: pptxIsDarkFill(slide.background), charts: [], failed: false };
  const mini = pptxBuildSlide(slide, ctx);
  mini.setAttribute('inert', '');
  mini.setAttribute('aria-hidden', 'true');
  mini.style.transform = `scale(${pptxState.thumbScale})`;
  pic.replaceChildren(mini);
  item.dataset.loaded = '1';
}

/** @param {MouseEvent} event */
function pptxOnThumbClick(event) {
  const target = event.target instanceof Element ? event.target : null;
  const item = target?.closest('.pptx-thumb');
  if (!(item instanceof HTMLElement)) return;
  const index = Number(item.dataset.index);
  if (Number.isInteger(index)) pptxGoTo(index);
}

/**
 * Marks the current slide in the strip and scrolls it into view.
 * @param {boolean} reveal
 */
function pptxSyncThumbSelection(reveal) {
  const s = pptxState;
  const dom = s.dom;
  if (!dom || !s.deck) return;
  for (const [i, item] of s.thumbItems) {
    const current = i === s.current;
    item.classList.toggle('pptx-thumb-current', current);
    item.setAttribute('aria-selected', String(current));
  }
  if (s.thumbItems.has(s.current)) dom.thumbs.setAttribute('aria-activedescendant', `pptx-thumb-${s.current}`);
  else dom.thumbs.removeAttribute('aria-activedescendant');
  if (!reveal || !s.thumbsOpen) return;
  const top = PPTX_THUMB_GAP + s.current * s.thumbItemHeight;
  const bottom = top + s.thumbItemHeight;
  const view = dom.thumbs;
  if (top - PPTX_THUMB_GAP < view.scrollTop) view.scrollTop = top - PPTX_THUMB_GAP;
  else if (bottom > view.scrollTop + view.clientHeight) view.scrollTop = bottom - view.clientHeight;
  pptxRenderThumbs();
  if (s.thumbItems.has(s.current)) dom.thumbs.setAttribute('aria-activedescendant', `pptx-thumb-${s.current}`);
}

// ----- PPTX: SLIDE -----

/**
 * White page of the deck's size (loading placeholder).
 * @param {PptxDeckMeta} deck
 */
function pptxBlankSlide(deck) {
  const el = h('div', { class: 'pptx-slide' });
  el.style.width = `${deck.width}pt`;
  el.style.height = `${deck.height}pt`;
  return el;
}

/**
 * The slide as absolutely positioned elements in points (unscaled; the caller sets the transform).
 * @param {PptxSlide} slide
 * @param {PptxBuildContext} ctx
 * @returns {HTMLElement}
 */
function pptxBuildSlide(slide, ctx) {
  const deck = /** @type {PptxDeckMeta} */ (pptxState.deck);
  const el = pptxBlankSlide(deck);
  if (ctx.dark) el.classList.add('pptx-slide-dark');
  pptxApplyBackground(el, slide.background);
  for (const item of Array.isArray(slide.elements) ? slide.elements : []) {
    const node = pptxBuildElementSafe(/** @type {PptxEl} */ (item), ctx, null);
    if (node) el.appendChild(node);
  }
  return el;
}

/**
 * @param {unknown} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 * @returns {Element | null}
 */
function pptxBuildElementSafe(el, ctx, mirror) {
  if (!el || typeof el !== 'object') return null;
  try {
    return pptxBuildElement(/** @type {PptxEl} */ (el), ctx, mirror);
  } catch (err) {
    if (!ctx.failed) {
      ctx.failed = true;
      logError(`Could not draw an element of slide ${ctx.index + 1}`, err);
    }
    return null;
  }
}

/**
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 * @returns {Element | null}
 */
function pptxBuildElement(el, ctx, mirror) {
  switch (el.type) {
    case 'shape':
      return pptxBuildShape(el, ctx, mirror, false);
    case 'text':
      return pptxBuildShape(el, ctx, mirror, true);
    case 'image':
      return pptxBuildImage(el, ctx, mirror);
    case 'table':
      return pptxBuildTable(el, mirror);
    case 'chart':
      return pptxBuildChart(el, ctx, mirror);
    case 'group':
    case 'diagram':
      return pptxBuildGroup(el, ctx, mirror);
    case 'media':
      return pptxBuildMedia(el, ctx, mirror);
    case 'math':
      return pptxBuildMath(el, mirror);
    default:
      return null;
  }
}

/**
 * @param {PptxEl} el
 * @param {PptxMirror | null} mirror
 * @returns {PptxPlace}
 */
function pptxPlace(el, mirror) {
  let left = pptxNum(el.left);
  let top = pptxNum(el.top);
  const width = Math.max(0, pptxNum(el.width));
  const height = Math.max(0, pptxNum(el.height));
  let rotate = pptxNum(el.rotate);
  let flipH = el.isFlipH === true;
  let flipV = el.isFlipV === true;
  if (mirror?.flipH) {
    left = mirror.w - left - width;
    rotate = -rotate;
    flipH = !flipH;
  }
  if (mirror?.flipV) {
    top = mirror.h - top - height;
    rotate = -rotate;
    flipV = !flipV;
  }
  return { left, top, width, height, rotate, flipH, flipV };
}

/**
 * Positioned box of an element (rotation about its centre, like PowerPoint).
 * @param {PptxPlace} place
 * @param {string} className
 */
function pptxBox(place, className) {
  const box = h('div', { class: `pptx-el ${className}` });
  box.style.left = `${pptxR(place.left)}pt`;
  box.style.top = `${pptxR(place.top)}pt`;
  box.style.width = `${pptxR(place.width)}pt`;
  box.style.height = `${pptxR(place.height)}pt`;
  const rotate = pptxR(place.rotate % 360);
  if (rotate) box.style.transform = `rotate(${rotate}deg)`;
  return box;
}

/**
 * A shape or text box: SVG geometry (fill, outline, line ends) + its text.
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 * @param {boolean} isText pptxtojson 'text' element (a text box: rectangle geometry)
 */
function pptxBuildShape(el, ctx, mirror, isText) {
  const place = pptxPlace(el, mirror);
  const svg = pptxShapeSvg(el, place, isText);
  const text = pptxTextLayer(el, place);
  if (!svg && !text) return null;
  const box = pptxBox(place, isText ? 'pptx-textbox' : 'pptx-shape');
  if (svg) box.appendChild(svg);
  if (text) box.appendChild(text);
  const shadow = pptxShadow(el.shadow);
  if (shadow) box.style.filter = shadow;
  if (!ctx.thumb) pptxSetLink(box, el.link);
  return box;
}

/**
 * @param {Element} box
 * @param {unknown} link
 */
function pptxSetLink(box, link) {
  if (typeof link !== 'string' || !link.trim()) return;
  const href = link.trim();
  box.setAttribute('data-pptx-link', href);
  box.classList.add('pptx-has-link');
  box.setAttribute('title', href);
  // A link for the keyboard and screen readers too (like the text links): focusable, named, opened with Enter.
  box.setAttribute('role', 'link');
  box.setAttribute('tabindex', '0');
  const jump = /^ppaction:\/\/hlinkshowjump\?jump=(\w+)/i.exec(href)?.[1].toLowerCase();
  const target =
    { nextslide: 'Next slide', previousslide: 'Previous slide', firstslide: 'First slide', lastslide: 'Last slide' }[jump ?? ''] ??
    (/(?:^|[\\/])slide(\d+)\.xml$/i.test(href) ? 'Another slide' : href);
  const text = (box.textContent ?? '').replace(/\s+/g, ' ').trim();
  box.setAttribute('aria-label', text && text !== target ? `${text} (${target})` : target);
  box.addEventListener('keydown', (event) => {
    if (!(event instanceof KeyboardEvent) || event.target !== box || (event.key !== 'Enter' && event.key !== ' ')) return;
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    event.preventDefault();
    event.stopPropagation();
    pptxFollowLink(href);
  });
}

/**
 * @param {unknown} shadow pptxtojson Shadow (offsets and blur in points)
 * @returns {string} CSS filter, '' for none
 */
function pptxShadow(shadow) {
  if (!shadow || typeof shadow !== 'object') return '';
  const sh = /** @type {Record<string, any>} */ (shadow);
  const color = pptxColor(sh.color);
  if (!color) return '';
  return `drop-shadow(${pptxR(pptxNum(sh.h))}pt ${pptxR(pptxNum(sh.v))}pt ${pptxR(Math.max(0, pptxNum(sh.blur)))}pt ${color})`;
}

/**
 * The SVG of a shape (or of a text box's fill / outline), drawn in points; null when nothing would be visible.
 * @param {PptxEl} el
 * @param {PptxPlace} place
 * @param {boolean} isText
 * @returns {SVGSVGElement | null}
 */
function pptxShapeSvg(el, place, isText) {
  const w = place.width;
  const hgt = place.height;
  const shapType = isText ? 'rect' : typeof el.shapType === 'string' && el.shapType ? el.shapType : 'rect';
  const isLine = !isText && PPTX_LINE_TYPES.has(shapType);
  const stroke = pptxStroke(el);
  const defs = pptxSvg('defs');
  const fill = isLine || el.strokeOnly === true ? { paint: 'none', image: '', opacity: 1 } : pptxSvgFill(el.fill, defs);
  if (!stroke && fill.paint === 'none' && !fill.image) return null;

  const g = pptxSvg('g', { transform: `scale(${PPTX_PX_PER_PT})${pptxFlipTransform(place)}` });
  if (isLine) {
    if (!stroke) return null;
    pptxDrawConnector(g, shapType, w, hgt, el, stroke);
  } else {
    const d = isText ? pptxRectPath(w, hgt) : pptxShapePathData(el, shapType, w, hgt);
    if (fill.image) {
      const clipId = pptxNextId();
      defs.appendChild(pptxSvg('clipPath', { id: clipId }, pptxSvg('path', { d })));
      g.appendChild(
        pptxSvg('image', {
          href: fill.image,
          x: 0,
          y: 0,
          width: Math.max(w, 0.01),
          height: Math.max(hgt, 0.01),
          preserveAspectRatio: 'none',
          'clip-path': `url(#${clipId})`,
          opacity: fill.opacity < 1 ? fill.opacity : null,
        }),
      );
    }
    // Presets with holes draw them as same-direction sub-paths; everything else keeps nonzero (custom geometry
    // concatenates separate DrawingML paths, which PowerPoint fills as a union).
    const path = pptxSvg('path', { d, fill: fill.paint, 'fill-rule': PPTX_EVENODD_TYPES.has(shapType) ? 'evenodd' : null });
    if (stroke) pptxApplyStroke(path, stroke);
    g.appendChild(path);
    if (stroke && (el.strokeOnly === true || shapType === 'custom') && (pptxEndType(el.headEnd) || pptxEndType(el.tailEnd))) {
      const head = pptxMarker(el.headEnd, stroke.color, defs, true);
      const tail = pptxMarker(el.tailEnd, stroke.color, defs, false);
      if (head) path.setAttribute('marker-start', `url(#${head})`);
      if (tail) path.setAttribute('marker-end', `url(#${tail})`);
    }
  }
  if (defs.childNodes.length) g.insertBefore(defs, g.firstChild);
  const svg = /** @type {SVGSVGElement} */ (pptxSvg('svg', { class: 'pptx-svg', 'aria-hidden': 'true', focusable: 'false' }, g));
  // An SVG of zero width or height is not rendered at all: keep 1pt (horizontal / vertical lines overflow it).
  svg.style.width = `${pptxR(Math.max(w, 1))}pt`;
  svg.style.height = `${pptxR(Math.max(hgt, 1))}pt`;
  return svg;
}

/**
 * SVG transform (in points) of a flipped element, '' when not flipped.
 * @param {PptxPlace} place
 */
function pptxFlipTransform(place) {
  let t = '';
  if (place.flipH) t += ` translate(${pptxR(place.width)} 0) scale(-1 1)`;
  if (place.flipV) t += ` translate(0 ${pptxR(place.height)}) scale(1 -1)`;
  return t;
}

/**
 * Outline of an element: colour, width (pt) and dash pattern; null for no outline.
 * @param {PptxEl} el
 * @returns {{ color: string, width: number, dash: string } | null}
 */
function pptxStroke(el) {
  const width = pptxNum(el.borderWidth);
  const color = pptxColor(el.borderColor);
  if (!(width > 0) || !color) return null;
  return { color, width, dash: pptxDashArray(el.borderStrokeDasharray, width) };
}

/**
 * @param {SVGElement} node
 * @param {{ color: string, width: number, dash: string }} stroke
 */
function pptxApplyStroke(node, stroke) {
  node.setAttribute('stroke', stroke.color);
  node.setAttribute('stroke-width', String(pptxR(stroke.width)));
  if (stroke.dash) node.setAttribute('stroke-dasharray', stroke.dash);
  node.setAttribute('stroke-miterlimit', '8');
}

/**
 * @param {unknown} value pptxtojson borderStrokeDasharray
 * @param {number} width line width (pt)
 * @returns {string} SVG stroke-dasharray, '' for solid
 */
function pptxDashArray(value, width) {
  const key = String(value ?? '').trim();
  if (!key || key === '0') return '';
  const units =
    PPTX_DASHES[key] ??
    key
      .split(/[\s,]+/)
      .map(Number)
      .filter((n) => Number.isFinite(n) && n >= 0);
  if (!units.length || units.every((n) => n === 0)) return '';
  const w = Math.max(width, 0.75);
  return units.map((n) => pptxR(n * w)).join(' ');
}

/**
 * Geometry of a (non-line) shape in points: pptxtojson's path scaled from its viewBox to the element size, own
 * geometry for the common presets when there is none, otherwise a rectangle.
 * @param {PptxEl} el
 * @param {string} shapType
 * @param {number} w
 * @param {number} h
 */
function pptxShapePathData(el, shapType, w, h) {
  // pptxtojson's geometry of these presets is wrong (arrow heads from the width, callout arcs): own geometry first.
  if (PPTX_OWN_PRESETS.has(shapType)) {
    const own = pptxPresetPath(shapType, w, h, el.keypoints);
    if (own) return own;
  }
  const path = typeof el.path === 'string' ? el.path.trim() : '';
  if (path && /^[MmZzLlHhVvCcSsQqTtAa0-9eE.,+\-\s]+$/.test(path) && /\d/.test(path)) {
    const vb = el.pathViewBox && typeof el.pathViewBox === 'object' ? el.pathViewBox : {};
    const vw = pptxNum(vb.width);
    const vh = pptxNum(vb.height);
    const vx = pptxNum(vb.x);
    const vy = pptxNum(vb.y);
    const sx = vw > 0 ? w / vw : 1;
    const sy = vh > 0 ? h / vh : 1;
    if (Math.abs(sx - 1) < 1e-4 && Math.abs(sy - 1) < 1e-4 && !vx && !vy) return path;
    const scaled = pptxScalePath(path, sx, sy, vx, vy);
    if (scaled) return scaled;
  }
  return pptxPresetPath(shapType, w, h, el.keypoints) || pptxRectPath(w, h);
}

/**
 * Scales SVG path data: x' = (x - ox) * sx, y' = (y - oy) * sy (offsets for absolute commands only).
 * @param {string} d
 * @param {number} sx
 * @param {number} sy
 * @param {number} ox
 * @param {number} oy
 * @returns {string} '' when the data is not understood
 */
function pptxScalePath(d, sx, sy, ox, oy) {
  const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g);
  if (!tokens) return '';
  /** @type {(string | number)[]} */
  const out = [];
  let cmd = '';
  let arg = 0;
  for (const token of tokens) {
    if (/^[a-zA-Z]$/.test(token)) {
      if (!/[MmZzLlHhVvCcSsQqTtAa]/.test(token)) return '';
      cmd = token;
      arg = 0;
      out.push(token);
      continue;
    }
    if (!cmd) return '';
    const abs = cmd === cmd.toUpperCase();
    const c = cmd.toUpperCase();
    let v = Number(token);
    /** @type {'x' | 'y' | 'rx' | 'ry' | ''} */
    let axis;
    if (c === 'H') axis = 'x';
    else if (c === 'V') axis = 'y';
    else if (c === 'A') {
      const k = arg % 7;
      axis = k === 0 ? 'rx' : k === 1 ? 'ry' : k === 5 ? 'x' : k === 6 ? 'y' : '';
    } else axis = arg % 2 === 0 ? 'x' : 'y';
    if (axis === 'x') v = (v - (abs ? ox : 0)) * sx;
    else if (axis === 'y') v = (v - (abs ? oy : 0)) * sy;
    else if (axis === 'rx') v = Math.abs(v * sx);
    else if (axis === 'ry') v = Math.abs(v * sy);
    out.push(pptxR(v, 1000));
    arg++;
  }
  return out.join(' ');
}

/** @param {number} w @param {number} h */
function pptxRectPath(w, h) {
  return `M 0 0 L ${pptxR(w)} 0 L ${pptxR(w)} ${pptxR(h)} L 0 ${pptxR(h)} Z`;
}

/**
 * DrawingML adjust value as a fraction (adjust / 100000). pptxtojson's `keypoints` hold the raw value / 50000.
 * @param {unknown} keypoints
 * @param {string} name
 * @param {number} fallback default adjust value (1/100000 units)
 */
function pptxAdj(keypoints, name, fallback) {
  const k = keypoints && typeof keypoints === 'object' ? Number(/** @type {Record<string, unknown>} */ (keypoints)[name]) : NaN;
  return (Number.isFinite(k) ? k * 50000 : fallback) / 100000;
}

/**
 * Own geometry of the common presets (DrawingML presetShapeDefinitions, default and given adjust values), used
 * when pptxtojson has no path, and to clip pictures. '' for other presets.
 * @param {string} type
 * @param {number} w
 * @param {number} h
 * @param {unknown} kp pptxtojson keypoints
 * @returns {string}
 */
function pptxPresetPath(type, w, h, kp) {
  const ss = Math.min(w, h);
  /** @param {number[][]} points */
  const poly = (points) => `M ${points.map(([x, y]) => `${pptxR(x)} ${pptxR(y)}`).join(' L ')} Z`;
  switch (type) {
    case 'rect':
      return pptxRectPath(w, h);
    case 'roundRect': {
      const r = pptxR(Math.min(ss * pptxAdj(kp, 'adj', 16667), w / 2, h / 2));
      const [W, H] = [pptxR(w), pptxR(h)];
      return `M ${r} 0 L ${pptxR(w - r)} 0 A ${r} ${r} 0 0 1 ${W} ${r} L ${W} ${pptxR(h - r)} A ${r} ${r} 0 0 1 ${pptxR(w - r)} ${H} L ${r} ${H} A ${r} ${r} 0 0 1 0 ${pptxR(h - r)} L 0 ${r} A ${r} ${r} 0 0 1 ${r} 0 Z`;
    }
    case 'ellipse': {
      const [rx, ry] = [pptxR(w / 2), pptxR(h / 2)];
      return `M 0 ${ry} A ${rx} ${ry} 0 1 1 ${pptxR(w)} ${ry} A ${rx} ${ry} 0 1 1 0 ${ry} Z`;
    }
    case 'triangle':
      return poly([[w * pptxAdj(kp, 'adj', 50000), 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 25000), w);
      return poly([[x, 0], [w, 0], [w - x, h], [0, h]]);
    }
    case 'trapezoid': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 25000), w / 2);
      return poly([[0, h], [x, 0], [w - x, 0], [w, h]]);
    }
    case 'pentagon':
      return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'hexagon': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 25000), w / 2);
      return poly([[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]]);
    }
    case 'octagon': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 29289), w / 2, h / 2);
      return poly([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]]);
    }
    case 'chevron': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 50000), w);
      return poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]);
    }
    case 'homePlate': {
      const x = Math.min(ss * pptxAdj(kp, 'adj', 50000), w);
      return poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h]]);
    }
    case 'rightArrow':
    case 'leftArrow': {
      const t = h * pptxAdj(kp, 'adj1', 50000);
      const head = Math.min(ss * pptxAdj(kp, 'adj2', 50000), w);
      const [y1, y2] = [(h - t) / 2, (h + t) / 2];
      return type === 'rightArrow'
        ? poly([[0, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, y2], [0, y2]])
        : poly([[w, y1], [head, y1], [head, 0], [0, h / 2], [head, h], [head, y2], [w, y2]]);
    }
    case 'upArrow':
    case 'downArrow': {
      const t = w * pptxAdj(kp, 'adj1', 50000);
      const head = Math.min(ss * pptxAdj(kp, 'adj2', 50000), h);
      const [x1, x2] = [(w - t) / 2, (w + t) / 2];
      return type === 'upArrow'
        ? poly([[x1, h], [x1, head], [0, head], [w / 2, 0], [w, head], [x2, head], [x2, h]])
        : poly([[x1, 0], [x1, h - head], [0, h - head], [w / 2, h], [w, h - head], [x2, h - head], [x2, 0]]);
    }
    case 'star5': {
      // DrawingML star5: radii stretched by hf / vf so the star touches all four sides of its box.
      const rx = (w / 2) * 1.05146;
      const ry = (h / 2) * 1.10557;
      const cy = ry;
      const inner = pptxAdj(kp, 'adj', 19098) * 2;
      /** @type {number[][]} */
      const points = [];
      for (let i = 0; i < 10; i++) {
        const angle = ((-90 + i * 36) * Math.PI) / 180;
        const r = i % 2 === 0 ? 1 : inner;
        points.push([w / 2 + rx * r * Math.cos(angle), cy + ry * r * Math.sin(angle)]);
      }
      return poly(points);
    }
    case 'wedgeEllipseCallout': {
      // The bubble is the ellipse from 11° before to 11° after the direction of the pointer tip.
      const [rx, ry] = [w / 2, h / 2];
      const dx = w * pptxAdj(kp, 'adj1', -20833);
      const dy = h * pptxAdj(kp, 'adj2', 62500);
      const angle = Math.atan2(dy * w, dx * h);
      const at = (/** @type {number} */ a) => [rx + rx * Math.cos(a), ry + ry * Math.sin(a)];
      const [x1, y1] = at(angle + (11 * Math.PI) / 180);
      const [x2, y2] = at(angle - (11 * Math.PI) / 180);
      return `M ${pptxR(rx + dx)} ${pptxR(ry + dy)} L ${pptxR(x1)} ${pptxR(y1)} A ${pptxR(rx)} ${pptxR(ry)} 0 1 1 ${pptxR(x2)} ${pptxR(y2)} Z`;
    }
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout':
      return pptxWedgeRectPath(type === 'wedgeRoundRectCallout', w, h, kp);
    default:
      // Other callouts are drawn as their rectangle without the pointer.
      if (/Callout|^callout\d$/.test(type)) return pptxRectPath(w, h);
      return '';
  }
}

/**
 * DrawingML wedgeRectCallout / wedgeRoundRectCallout: the pointer leaves the side the tip is beyond (between 2/12 and
 * 5/12, or 7/12 and 10/12, of that side).
 * @param {boolean} round
 * @param {number} w
 * @param {number} h
 * @param {unknown} kp
 */
function pptxWedgeRectPath(round, w, h, kp) {
  const dx = w * pptxAdj(kp, 'adj1', -20833);
  const dy = h * pptxAdj(kp, 'adj2', 62500);
  const [xp, yp] = [w / 2 + dx, h / 2 + dy];
  const vertical = Math.abs(dy) - Math.abs((dx * h) / w) > 0;
  const [x1, x2] = dx > 0 ? [(w * 7) / 12, (w * 10) / 12] : [(w * 2) / 12, (w * 5) / 12];
  const [y1, y2] = dy > 0 ? [(h * 7) / 12, (h * 10) / 12] : [(h * 2) / 12, (h * 5) / 12];
  const r = round ? Math.min(Math.min(w, h) * pptxAdj(kp, 'adj3', 16667), w / 2, h / 2) : 0;
  /** @type {(string | number)[]} */
  const d = ['M', 0, r];
  const corner = (/** @type {number} */ x, /** @type {number} */ y) => {
    if (r > 0) d.push('A', r, r, 0, 0, 1, x, y);
  };
  const line = (/** @type {number} */ x, /** @type {number} */ y) => d.push('L', x, y);
  corner(r, 0);
  if (vertical && dy < 0) [[x1, 0], [xp, yp], [x2, 0]].forEach(([x, y]) => line(x, y));
  line(w - r, 0);
  corner(w, r);
  if (!vertical && dx > 0) [[w, y1], [xp, yp], [w, y2]].forEach(([x, y]) => line(x, y));
  line(w, h - r);
  corner(w - r, h);
  if (vertical && dy > 0) [[x2, h], [xp, yp], [x1, h]].forEach(([x, y]) => line(x, y));
  line(r, h);
  corner(0, h - r);
  if (!vertical && dx <= 0) [[0, y2], [xp, yp], [0, y1]].forEach(([x, y]) => line(x, y));
  d.push('Z');
  return d.map((v) => (typeof v === 'number' ? pptxR(v) : v)).join(' ');
}

/**
 * The text rectangle of a preset (DrawingML presetShapeDefinitions <a:rect>) in points, null for the whole box.
 * @param {string} type
 * @param {number} w
 * @param {number} h
 * @param {unknown} kp pptxtojson keypoints
 * @returns {{ l: number, t: number, r: number, b: number } | null}
 */
function pptxTextRect(type, w, h, kp) {
  const ss = Math.min(w, h);
  if (!(ss > 0)) return null;
  switch (type) {
    case 'ellipse':
    case 'wedgeEllipseCallout': {
      const [dx, dy] = [(w / 2) * (1 - Math.SQRT1_2), (h / 2) * (1 - Math.SQRT1_2)];
      return { l: dx, t: dy, r: w - dx, b: h - dy };
    }
    case 'roundRect':
    case 'wedgeRoundRectCallout': {
      const i = Math.min(ss * pptxAdj(kp, type === 'roundRect' ? 'adj' : 'adj3', 16667), w / 2, h / 2) * 0.29289;
      return { l: i, t: i, r: w - i, b: h - i };
    }
    case 'triangle': {
      const x1 = (w * pptxAdj(kp, 'adj', 50000)) / 2;
      return { l: x1, t: h / 2, r: x1 + w / 2, b: h };
    }
    case 'rtTriangle':
      return { l: w / 12, t: (h * 7) / 12, r: (w * 7) / 12, b: (h * 11) / 12 };
    case 'diamond':
      return { l: w / 4, t: h / 4, r: (w * 3) / 4, b: (h * 3) / 4 };
    case 'parallelogram': {
      const max = w / ss;
      const q = (1 + (5 * Math.min(pptxAdj(kp, 'adj', 25000), max)) / max) / 12;
      return { l: q * w, t: q * h, r: w - q * w, b: h - q * h };
    }
    case 'trapezoid': {
      const max = (0.5 * w) / ss;
      const f = Math.min(pptxAdj(kp, 'adj', 25000), max) / max / 3;
      return { l: f * w, t: f * h, r: w - f * w, b: h };
    }
    case 'pentagon':
      return { l: w * 0.191, t: h * 0.236, r: w * 0.809, b: h };
    case 'octagon': {
      const i = Math.min(ss * pptxAdj(kp, 'adj', 29289), w / 2, h / 2) / 2;
      return { l: i, t: i, r: w - i, b: h - i };
    }
    case 'chevron': {
      const x1 = Math.min(ss * pptxAdj(kp, 'adj', 50000), w);
      return w - x1 > x1 ? { l: x1, t: 0, r: w - x1, b: h } : null;
    }
    case 'homePlate': {
      const x1 = Math.min(ss * pptxAdj(kp, 'adj', 50000), w);
      return { l: 0, t: 0, r: (w - x1 + w) / 2, b: h };
    }
    case 'rightArrow':
    case 'leftArrow': {
      const y1 = h / 2 - (h * pptxAdj(kp, 'adj1', 50000)) / 2;
      const head = Math.min(ss * pptxAdj(kp, 'adj2', 50000), w);
      const back = (y1 * head) / (h / 2);
      return type === 'rightArrow' ? { l: 0, t: y1, r: w - head + back, b: h - y1 } : { l: head - back, t: y1, r: w, b: h - y1 };
    }
    case 'upArrow':
    case 'downArrow': {
      const x1 = w / 2 - (w * pptxAdj(kp, 'adj1', 50000)) / 2;
      const head = Math.min(ss * pptxAdj(kp, 'adj2', 50000), h);
      const back = (x1 * head) / (w / 2);
      return type === 'upArrow' ? { l: x1, t: head - back, r: w - x1, b: h } : { l: x1, t: 0, r: w - x1, b: h - head + back };
    }
    default:
      return null;
  }
}

/**
 * pptxtojson LineEnd type, '' for none.
 * @param {unknown} end
 */
function pptxEndType(end) {
  const type = end && typeof end === 'object' ? /** @type {Record<string, unknown>} */ (end).type : undefined;
  return typeof type === 'string' && type !== 'none' ? type : '';
}

/**
 * Arrowhead outline in line-width units, tip at the origin, pointing along +x; `cut`: how far the line must stop
 * before the tip so it does not show through the head.
 * @param {unknown} end
 * @returns {{ d: string, filled: boolean, cut: number } | null}
 */
function pptxEndShape(end) {
  const type = pptxEndType(end);
  if (!type) return null;
  const e = /** @type {Record<string, unknown>} */ (end);
  const L = PPTX_END_SIZES[String(e.length)] ?? 3;
  const W = PPTX_END_SIZES[String(e.width)] ?? 3;
  const hw = W / 2;
  switch (type) {
    case 'stealth':
      return { d: `M 0 0 L ${-L} ${-hw} L ${-L * 0.7} 0 L ${-L} ${hw} Z`, filled: true, cut: L * 0.5 };
    case 'arrow':
      return { d: `M ${-L} ${-hw} L 0 0 L ${-L} ${hw}`, filled: false, cut: 0.5 };
    case 'diamond':
      return { d: `M ${L / 2} 0 L 0 ${-hw} L ${-L / 2} 0 L 0 ${hw} Z`, filled: true, cut: 0 };
    case 'oval':
      return { d: `M ${L / 2} 0 A ${L / 2} ${hw} 0 1 1 ${-L / 2} 0 A ${L / 2} ${hw} 0 1 1 ${L / 2} 0 Z`, filled: true, cut: 0 };
    default: // triangle (and unknown types)
      return { d: `M 0 0 L ${-L} ${-hw} L ${-L} ${hw} Z`, filled: true, cut: L * 0.6 };
  }
}

/**
 * SVG marker for a line end of a path whose geometry is not ours (arcs, custom lines). Returns its id.
 * @param {unknown} end
 * @param {string} color
 * @param {SVGElement} defs
 * @param {boolean} atStart
 * @returns {string}
 */
function pptxMarker(end, color, defs, atStart) {
  const shape = pptxEndShape(end);
  if (!shape) return '';
  const id = pptxNextId();
  const path = pptxSvg('path', shape.filled ? { d: shape.d, fill: color } : { d: shape.d, fill: 'none', stroke: color, 'stroke-width': 1 });
  defs.appendChild(
    pptxSvg(
      'marker',
      { id, viewBox: '-6 -6 12 12', refX: 0, refY: 0, markerWidth: 12, markerHeight: 12, markerUnits: 'strokeWidth', orient: atStart ? 'auto-start-reverse' : 'auto', overflow: 'visible' },
      path,
    ),
  );
  return id;
}

/**
 * Lines and connectors (DrawingML geometry, adjust values from keypoints): the line, shortened under its
 * arrowheads, and the heads drawn at the true end points.
 * @param {SVGElement} g
 * @param {string} type
 * @param {number} w
 * @param {number} h
 * @param {PptxEl} el
 * @param {{ color: string, width: number, dash: string }} stroke
 */
function pptxDrawConnector(g, type, w, h, el, stroke) {
  const kp = el.keypoints;
  /** @type {number[]} */
  const start = [0, 0];
  /** @type {{ c?: number[][], p: number[] }[]} */
  let segs;
  switch (type) {
    case 'bentConnector2':
      segs = [{ p: [w, 0] }, { p: [w, h] }];
      break;
    case 'bentConnector3': {
      const x1 = w * pptxAdj(kp, 'adj1', 50000);
      segs = [{ p: [x1, 0] }, { p: [x1, h] }, { p: [w, h] }];
      break;
    }
    case 'bentConnector4': {
      const x1 = w * pptxAdj(kp, 'adj1', 50000);
      const y2 = h * pptxAdj(kp, 'adj2', 50000);
      segs = [{ p: [x1, 0] }, { p: [x1, y2] }, { p: [w, y2] }, { p: [w, h] }];
      break;
    }
    case 'bentConnector5': {
      const x1 = w * pptxAdj(kp, 'adj1', 50000);
      const y2 = h * pptxAdj(kp, 'adj2', 50000);
      const x3 = w * pptxAdj(kp, 'adj3', 50000);
      segs = [{ p: [x1, 0] }, { p: [x1, y2] }, { p: [x3, y2] }, { p: [x3, h] }, { p: [w, h] }];
      break;
    }
    case 'curvedConnector2':
      segs = [{ c: [[w / 2, 0], [w, h / 2]], p: [w, h] }];
      break;
    case 'curvedConnector3': {
      const x2 = w * pptxAdj(kp, 'adj1', 50000);
      segs = [
        { c: [[x2 / 2, 0], [x2, h / 4]], p: [x2, h / 2] },
        { c: [[x2, (h * 3) / 4], [(w + x2) / 2, h]], p: [w, h] },
      ];
      break;
    }
    case 'curvedConnector4':
    case 'curvedConnector5': {
      const x2 = w * pptxAdj(kp, 'adj1', 50000);
      const y4 = h * pptxAdj(kp, 'adj2', 50000);
      const x3 = (w + x2) / 2;
      const y1 = y4 / 2;
      segs = [
        { c: [[x2 / 2, 0], [x2, y1 / 2]], p: [x2, y1] },
        { c: [[x2, (y1 + y4) / 2], [(x2 + x3) / 2, y4]], p: [x3, y4] },
        { c: [[(x3 + w) / 2, y4], [w, (h + y4) / 2]], p: [w, h] },
      ];
      break;
    }
    default: // line, straightConnector1
      segs = [{ p: [w, h] }];
  }

  const head = pptxEndShape(el.headEnd);
  const tail = pptxEndShape(el.tailEnd);
  const sw = stroke.width;
  const last = segs[segs.length - 1];
  const beforeEnd = last.c ? last.c[last.c.length - 1] : segs.length > 1 ? segs[segs.length - 2].p : start;
  const afterStart = segs[0].c ? segs[0].c[0] : segs[0].p;
  const from = head ? pptxMoveToward(start, afterStart, head.cut * sw) : start;
  const to = tail ? pptxMoveToward(last.p, beforeEnd, tail.cut * sw) : last.p;

  /** @param {number[]} p */
  const pt = (p) => `${pptxR(p[0])} ${pptxR(p[1])}`;
  let d = `M ${pt(from)}`;
  segs.forEach((seg, i) => {
    const p = i === segs.length - 1 ? to : seg.p;
    d += seg.c ? ` C ${pt(seg.c[0])} ${pt(seg.c[1])} ${pt(p)}` : ` L ${pt(p)}`;
  });
  const line = pptxSvg('path', { d, fill: 'none', 'stroke-linejoin': 'round' });
  pptxApplyStroke(line, stroke);
  g.appendChild(line);
  if (head) g.appendChild(pptxEndPath(head, start, afterStart, sw, stroke.color));
  if (tail) g.appendChild(pptxEndPath(tail, last.p, beforeEnd, sw, stroke.color));
}

/**
 * Point `p` moved towards `q` by `distance` (at most 45 % of the way).
 * @param {number[]} p
 * @param {number[]} q
 * @param {number} distance
 */
function pptxMoveToward(p, q, distance) {
  const dx = q[0] - p[0];
  const dy = q[1] - p[1];
  const len = Math.hypot(dx, dy);
  if (!len || !(distance > 0)) return p;
  const t = Math.min(distance, len * 0.45) / len;
  return [p[0] + dx * t, p[1] + dy * t];
}

/**
 * An arrowhead at `tip`, pointing away from `from`.
 * @param {{ d: string, filled: boolean }} shape
 * @param {number[]} tip
 * @param {number[]} from
 * @param {number} sw line width (pt)
 * @param {string} color
 */
function pptxEndPath(shape, tip, from, sw, color) {
  const dx = tip[0] - from[0];
  const dy = tip[1] - from[1];
  const angle = dx || dy ? (Math.atan2(dy, dx) * 180) / Math.PI : 0;
  return pptxSvg('path', {
    d: shape.d,
    transform: `translate(${pptxR(tip[0])} ${pptxR(tip[1])}) rotate(${pptxR(angle)}) scale(${pptxR(Math.max(sw, 0.5), 1000)})`,
    fill: shape.filled ? color : 'none',
    stroke: shape.filled ? null : color,
    'stroke-width': shape.filled ? null : 1,
    'stroke-linejoin': 'miter',
  });
}

/**
 * Fill of a shape as an SVG paint (gradients go into `defs`) or a picture.
 * @param {unknown} fill pptxtojson Fill
 * @param {SVGElement} defs
 * @returns {{ paint: string, image: string, opacity: number }}
 */
function pptxSvgFill(fill, defs) {
  const none = { paint: 'none', image: '', opacity: 1 };
  if (!fill || typeof fill !== 'object') return none;
  const f = /** @type {Record<string, any>} */ (fill);
  switch (f.type) {
    case 'color':
      return { ...none, paint: pptxColor(f.value) || 'none' };
    case 'gradient': {
      const id = pptxGradientDef(f.value, defs);
      return { ...none, paint: id ? `url(#${id})` : 'none' };
    }
    case 'image': {
      const value = f.value && typeof f.value === 'object' ? f.value : {};
      const image = pptxImageUri(value.base64);
      const opacity = pptxNum(value.opacity, 1);
      return { paint: 'none', image, opacity: opacity > 0 && opacity <= 1 ? opacity : 1 };
    }
    case 'pattern':
      return { ...none, paint: pptxPatternColor(f.value) || 'none' };
    default:
      return none;
  }
}

/**
 * Gradient stops: offsets in percent, valid colours only.
 * @param {unknown} value pptxtojson GradientFill value
 * @returns {{ offset: number, color: string }[]}
 */
function pptxGradientStops(value) {
  const colors = value && typeof value === 'object' ? /** @type {Record<string, any>} */ (value).colors : null;
  if (!Array.isArray(colors)) return [];
  /** @type {{ offset: number, color: string }[]} */
  const stops = [];
  for (const stop of colors) {
    const color = pptxColor(stop?.color);
    if (!color) continue;
    stops.push({ offset: clamp(pptxNum(parseFloat(String(stop.pos))), 0, 100), color });
  }
  return stops.sort((a, b) => a.offset - b.offset);
}

/**
 * Gradient stops with in-between stops blended in linear light, as PowerPoint blends (the browser blends in sRGB,
 * which makes the middle tones too dark). Colours it cannot read are kept as they are.
 * @param {{ offset: number, color: string }[]} stops
 * @returns {{ offset: number, color: string }[]}
 */
function pptxLinearStops(stops) {
  /** @param {string} color */
  const read = (color) => {
    const hex = pptxRgba(color);
    if (hex) return hex;
    const m = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+)(%?))?\s*\)$/i.exec(color);
    if (!m) return null;
    const alpha = m[4] === undefined ? 1 : Number(m[4]) / (m[5] ? 100 : 1);
    return [Number(m[1]), Number(m[2]), Number(m[3]), alpha];
  };
  const toLinear = (/** @type {number} */ c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const fromLinear = (/** @type {number} */ c) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
  /** @type {{ offset: number, color: string }[]} */
  const out = [];
  stops.forEach((stop, i) => {
    const prev = stops[i - 1];
    const a = prev ? read(prev.color) : null;
    const b = read(stop.color);
    if (prev && a && b && stop.offset > prev.offset) {
      for (let k = 1; k < PPTX_GRADIENT_STEPS; k++) {
        const t = k / PPTX_GRADIENT_STEPS;
        const rgb = [0, 1, 2].map((j) => Math.round(fromLinear(toLinear(a[j] / 255) * (1 - t) + toLinear(b[j] / 255) * t) * 255));
        const alpha = pptxR(a[3] * (1 - t) + b[3] * t, 1000);
        out.push({ offset: prev.offset + (stop.offset - prev.offset) * t, color: `rgba(${rgb.join(', ')}, ${alpha})` });
      }
    }
    out.push(stop);
  });
  return out;
}

/**
 * Adds a <linearGradient> / <radialGradient> to `defs`; returns its id ('' without stops).
 * DrawingML angles: 0° = left to right, 90° = top to bottom; path gradients start at the centre.
 * @param {unknown} value
 * @param {SVGElement} defs
 */
function pptxGradientDef(value, defs) {
  const stops = pptxGradientStops(value);
  if (!stops.length) return '';
  const v = /** @type {Record<string, any>} */ (value);
  const id = pptxNextId();
  let grad;
  if (v.path === 'circle' || v.path === 'rect' || v.path === 'shape') {
    grad = pptxSvg('radialGradient', { id, cx: '50%', cy: '50%', r: '50%' });
  } else {
    const rad = (pptxNum(v.rot) * Math.PI) / 180;
    const dx = Math.cos(rad) / 2;
    const dy = Math.sin(rad) / 2;
    grad = pptxSvg('linearGradient', { id, x1: pptxR(0.5 - dx, 1000), y1: pptxR(0.5 - dy, 1000), x2: pptxR(0.5 + dx, 1000), y2: pptxR(0.5 + dy, 1000) });
  }
  for (const stop of pptxLinearStops(stops)) grad.appendChild(pptxSvg('stop', { offset: `${pptxR(stop.offset)}%`, 'stop-color': stop.color }));
  defs.appendChild(grad);
  return id;
}

/**
 * CSS gradient of a fill value ('' without stops).
 * @param {unknown} value
 */
function pptxCssGradient(value) {
  const stops = pptxGradientStops(value);
  if (!stops.length) return '';
  const list = pptxLinearStops(stops)
    .map((s) => `${s.color} ${pptxR(s.offset)}%`)
    .join(', ');
  const v = /** @type {Record<string, any>} */ (value);
  if (v.path === 'circle' || v.path === 'rect' || v.path === 'shape') return `radial-gradient(closest-side, ${list})`;
  return `linear-gradient(${pptxR(pptxNum(v.rot) + 90)}deg, ${list})`;
}

/**
 * A pattern fill drawn as one colour: foreground and background mixed by the pattern's density.
 * @param {unknown} value pptxtojson PatternFill value
 */
function pptxPatternColor(value) {
  if (!value || typeof value !== 'object') return '';
  const v = /** @type {Record<string, any>} */ (value);
  const fg = pptxColor(v.foregroundColor);
  const bg = pptxColor(v.backgroundColor);
  if (!fg || !bg) return fg || bg;
  const type = String(v.type || '');
  const pct = /^pct(\d+)$/.exec(type);
  const amount = pct ? Number(pct[1]) / 100 : /^dk/.test(type) ? 0.7 : /^lt/.test(type) ? 0.3 : 0.5;
  return pptxMixColors(fg, bg, amount);
}

/**
 * Slide background (resolved by the host through layout / master; null = white).
 * @param {HTMLElement} el
 * @param {unknown} fill
 */
function pptxApplyBackground(el, fill) {
  el.style.backgroundColor = '#ffffff';
  if (!fill || typeof fill !== 'object') return;
  const f = /** @type {Record<string, any>} */ (fill);
  switch (f.type) {
    case 'color': {
      const color = pptxColor(f.value);
      if (color) el.style.backgroundColor = color;
      break;
    }
    case 'gradient': {
      const css = pptxCssGradient(f.value);
      if (css) el.style.backgroundImage = css;
      break;
    }
    case 'image': {
      const src = pptxImageUri(f.value && typeof f.value === 'object' ? f.value.base64 : '');
      if (src) {
        el.style.backgroundImage = `url("${src}")`;
        el.style.backgroundSize = '100% 100%';
      }
      break;
    }
    case 'pattern': {
      const color = pptxPatternColor(f.value);
      if (color) el.style.backgroundColor = color;
      break;
    }
  }
}

/**
 * Dark slide backgrounds get light chart text.
 * @param {unknown} fill
 */
function pptxIsDarkFill(fill) {
  if (!fill || typeof fill !== 'object') return false;
  const f = /** @type {Record<string, any>} */ (fill);
  /** @type {string[]} */
  let colors = [];
  if (f.type === 'color') colors = [pptxColor(f.value)];
  else if (f.type === 'gradient') colors = pptxGradientStops(f.value).map((s) => s.color);
  else if (f.type === 'pattern') colors = [pptxPatternColor(f.value)];
  const lums = colors.map(pptxLuminance).filter((l) => l !== null);
  return lums.length > 0 && lums.reduce((a, b) => /** @type {number} */ (a) + /** @type {number} */ (b), 0) / lums.length < 0.4;
}

// ----- PPTX: TEXT -----

/**
 * The text of a shape / text box: sanitized HTML in a box with the element's insets and vertical anchor.
 * @param {PptxEl} el
 * @param {PptxPlace} place
 * @returns {HTMLElement | null}
 */
function pptxTextLayer(el, place) {
  const html = typeof el.content === 'string' ? el.content : '';
  if (!html || !pptxHasText(html)) return null;
  const anchor = el.vAlign === 'mid' ? 'mid' : el.vAlign === 'down' ? 'down' : 'up';
  const layer = h('div', { class: `pptx-text pptx-anchor-${anchor}` });
  const inset = el.textInset && typeof el.textInset === 'object' ? el.textInset : PPTX_DEFAULT_INSET;
  // The text sits in the preset's text rectangle (a triangle's lower middle, an arrow's shaft…), mirrored with a flip.
  const rect = el.type === 'shape' ? pptxTextRect(String(el.shapType || ''), place.width, place.height, el.keypoints) : null;
  /** @type {Record<string, number>} */
  const off = rect ? { l: rect.l, t: rect.t, r: place.width - rect.r, b: place.height - rect.b } : { l: 0, t: 0, r: 0, b: 0 };
  if (place.flipH !== place.flipV) [off.l, off.r] = [off.r, off.l];
  layer.style.padding = ['t', 'r', 'b', 'l']
    .map((k) => `${pptxR(Math.max(0, pptxNum(inset[k], /** @type {any} */ (PPTX_DEFAULT_INSET)[k])) + Math.max(0, off[k]))}pt`)
    .join(' ');
  if (el.wrap === false) layer.classList.add('pptx-nowrap');
  if (el.isVertical === true) layer.classList.add('pptx-vertical');
  // PowerPoint turns the text of a vertically flipped shape upside down (a horizontal flip leaves it readable);
  // 270° vertical text (bodyPr vert270) is vertical text turned upside down.
  if (place.flipV !== (el.vert270 === true)) layer.style.transform = 'rotate(180deg)';
  const body = h('div', { class: 'pptx-text-body' });
  body.innerHTML = sanitize(html, DOCX_SANITIZE_OPTIONS);
  pptxFixText(body);
  const autoFit = el.autoFit && typeof el.autoFit === 'object' ? el.autoFit : null;
  const fontScale = autoFit && autoFit.type === 'text' ? pptxNum(autoFit.fontScale) : 0;
  if (fontScale > 0 && fontScale < 100) body.style.zoom = String(fontScale / 100);
  layer.appendChild(body);
  return layer;
}

/**
 * Whether text HTML has anything but tags and white space.
 * @param {string} html
 */
function pptxHasText(html) {
  // A zero-width space is the text of an empty paragraph (pptx.ts sizes it like PowerPoint).
  return /\S/.test(html.replace(/<[^>]*>/g, '').split(String.fromCharCode(0x200b)).join('').replace(/&nbsp;|&#160;|&#xa0;| /gi, ' '));
}

/**
 * pptxtojson writes every space as &nbsp; (nothing would wrap): back to spaces, kept by `white-space: pre-wrap`.
 * Fonts get fallbacks so a font that is not installed degrades to a similar one, not the browser's serif default.
 * @param {HTMLElement} scope
 */
function pptxFixText(scope) {
  const walker = document.createTreeWalker(scope, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = /** @type {Text} */ (node);
    if (text.data.includes(' ')) text.data = text.data.replace(/ /g, ' ');
  }
  scope.querySelectorAll('[style]').forEach((el) => {
    const style = /** @type {HTMLElement} */ (el).style;
    if (style && style.fontFamily) style.fontFamily = pptxFontStack(style.fontFamily);
  });
  scope.querySelectorAll('p').forEach((p) => {
    // The line box of a paragraph is as tall as its largest run (pptxtojson sizes the runs only; the slide's 18pt
    // would otherwise be the smallest line height).
    let size = 0;
    p.querySelectorAll('span').forEach((span) => {
      size = Math.max(size, pptxPt(/** @type {HTMLElement} */ (span).style.fontSize));
    });
    if (size > 0) p.style.fontSize = `${pptxR(size)}pt`;
    // The bullet hangs at marL + indent (never left of the text box) and the text starts at marL.
    const bullet = p.querySelector('.pptx-bullet');
    if (bullet instanceof HTMLElement) {
      const indent = pptxPt(p.style.textIndent);
      if (indent < 0 && pptxPt(p.style.marginLeft) + indent < 0) p.style.marginLeft = `${pptxR(-indent)}pt`;
      if (indent < 0) bullet.style.minWidth = `${pptxR(-indent)}pt`;
      else bullet.style.paddingRight = '0.3em';
    }
  });
}

/**
 * A CSS length of pptxtojson's HTML in points (pt or px; 0 for anything else).
 * @param {string} value
 */
function pptxPt(value) {
  const m = /^(-?[\d.]+)(pt|px)$/.exec(String(value || '').trim());
  if (!m) return 0;
  const n = Number(m[1]);
  return Number.isFinite(n) ? (m[2] === 'px' ? n * 0.75 : n) : 0;
}

/**
 * @param {string} family CSS font-family value of a text run
 */
function pptxFontStack(family) {
  if (/\b(?:sans-serif|serif|monospace|system-ui)\s*$/i.test(family)) return family;
  if (/mono|consol|courier|code/i.test(family)) return `${family}, ${PPTX_MONO_FALLBACK}`;
  if (/times|georgia|cambria|garamond|palatino|book antiqua|baskerville|minion|serif/i.test(family) && !/sans/i.test(family)) {
    return `${family}, ${PPTX_SERIF_FALLBACK}`;
  }
  return `${family}, ${PPTX_SANS_FALLBACK}`;
}

// ----- PPTX: PICTURES, TABLES, GROUPS, MEDIA -----

/**
 * A picture: cropped (srcRect, percent), clipped to its shape, flipped, with its outline.
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 */
function pptxBuildImage(el, ctx, mirror) {
  const place = pptxPlace(el, mirror);
  const box = pptxBox(place, 'pptx-image');
  if (!ctx.thumb) pptxSetLink(box, el.link);
  const src = pptxImageUri(el.base64);
  if (!src) {
    box.classList.add('pptx-image-missing');
    return box;
  }
  const w = place.width;
  const hgt = place.height;
  const crop = el.rect && typeof el.rect === 'object' ? el.rect : {};
  const [cl, ct, cr, cb] = ['l', 't', 'r', 'b'].map((k) => clamp(pptxNum(crop[k]) / 100, -10, 0.99));
  const visibleW = Math.max(0.01, 1 - cl - cr);
  const visibleH = Math.max(0.01, 1 - ct - cb);
  const geom = typeof el.geom === 'string' && el.geom ? el.geom : 'rect';
  const clipPath = geom === 'rect' ? '' : pptxPresetPath(geom, w, hgt, null);
  const stroke = pptxStroke(el);

  if (!clipPath && !stroke) {
    const img = h('img', { class: 'pptx-img', src, alt: '', draggable: 'false' });
    img.style.width = `${pptxR(100 / visibleW, 1000)}%`;
    img.style.height = `${pptxR(100 / visibleH, 1000)}%`;
    img.style.left = `${pptxR((-cl / visibleW) * 100, 1000)}%`;
    img.style.top = `${pptxR((-ct / visibleH) * 100, 1000)}%`;
    const clip = h('div', { class: 'pptx-img-clip' }, img);
    if (place.flipH || place.flipV) clip.style.transform = `scale(${place.flipH ? -1 : 1}, ${place.flipV ? -1 : 1})`;
    box.appendChild(clip);
    return box;
  }

  // Shaped or outlined picture: SVG (clip path in points, outline on the same geometry).
  const d = clipPath || pptxRectPath(w, hgt);
  const clipId = pptxNextId();
  const g = pptxSvg(
    'g',
    { transform: `scale(${PPTX_PX_PER_PT})${pptxFlipTransform(place)}` },
    pptxSvg('defs', null, pptxSvg('clipPath', { id: clipId }, pptxSvg('path', { d }))),
    pptxSvg('image', {
      href: src,
      x: pptxR((-cl / visibleW) * w),
      y: pptxR((-ct / visibleH) * hgt),
      width: pptxR(Math.max(w, 0.01) / visibleW),
      height: pptxR(Math.max(hgt, 0.01) / visibleH),
      preserveAspectRatio: 'none',
      'clip-path': `url(#${clipId})`,
    }),
  );
  if (stroke) {
    const outline = pptxSvg('path', { d, fill: 'none' });
    pptxApplyStroke(outline, stroke);
    g.appendChild(outline);
  }
  const svg = pptxSvg('svg', { class: 'pptx-svg', role: 'img', 'aria-label': 'Picture', focusable: 'false' }, g);
  /** @type {SVGSVGElement} */ (svg).style.width = `${pptxR(Math.max(w, 1))}pt`;
  /** @type {SVGSVGElement} */ (svg).style.height = `${pptxR(Math.max(hgt, 1))}pt`;
  box.appendChild(svg);
  return box;
}

/**
 * An equation: its fallback picture, or its text.
 * @param {PptxEl} el
 * @param {PptxMirror | null} mirror
 */
function pptxBuildMath(el, mirror) {
  const box = pptxBox(pptxPlace(el, mirror), 'pptx-math');
  const src = pptxImageUri(el.picBase64);
  const text = typeof el.text === 'string' && el.text ? el.text : typeof el.latex === 'string' ? el.latex : '';
  if (src) box.appendChild(h('img', { class: 'pptx-math-img', src, alt: text || 'Equation', draggable: 'false' }));
  else if (text) box.appendChild(h('span', { class: 'pptx-math-text' }, text));
  return box;
}

/**
 * A table: column widths and row heights in points; merged cells span; cell fills, text colours and borders.
 * @param {PptxEl} el
 * @param {PptxMirror | null} mirror
 */
function pptxBuildTable(el, mirror) {
  const place = pptxPlace(el, mirror);
  const box = pptxBox(place, 'pptx-table-box');
  const rows = Array.isArray(el.data) ? el.data : [];
  const colWidths = Array.isArray(el.colWidths) ? el.colWidths.map((w) => Math.max(0, pptxNum(w))) : [];
  const rowHeights = Array.isArray(el.rowHeights) ? el.rowHeights : [];
  const colCount = Math.max(colWidths.length, ...rows.map((r) => (Array.isArray(r) ? r.length : 0)));
  const outer = el.borders && typeof el.borders === 'object' ? el.borders : {};
  const table = h('table', { class: 'pptx-table' });
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  table.style.width = `${pptxR(totalWidth > 0 ? totalWidth : place.width)}pt`;
  if (colWidths.length) {
    table.appendChild(
      h(
        'colgroup',
        null,
        colWidths.map((w) => {
          const col = h('col');
          col.style.width = `${pptxR(w)}pt`;
          return col;
        }),
      ),
    );
  }
  const tbody = h('tbody');
  rows.forEach((row, i) => {
    const tr = h('tr');
    const height = pptxNum(rowHeights[i]);
    if (height > 0) tr.style.height = `${pptxR(height)}pt`;
    (Array.isArray(row) ? row : []).forEach((cell, j) => {
      if (!cell || typeof cell !== 'object' || cell.hMerge || cell.vMerge) return;
      const td = /** @type {HTMLTableCellElement} */ (h('td'));
      const rowSpan = Number.isInteger(cell.rowSpan) && cell.rowSpan > 1 ? cell.rowSpan : 1;
      const colSpan = Number.isInteger(cell.colSpan) && cell.colSpan > 1 ? cell.colSpan : 1;
      if (rowSpan > 1) td.rowSpan = rowSpan;
      if (colSpan > 1) td.colSpan = colSpan;
      const fillColor = pptxColor(cell.fillColor);
      if (fillColor) td.style.backgroundColor = fillColor;
      const fontColor = pptxColor(cell.fontColor);
      if (fontColor) td.style.color = fontColor;
      if (cell.fontBold === true) td.style.fontWeight = 'bold';
      td.style.verticalAlign = cell.vAlign === 'mid' ? 'middle' : cell.vAlign === 'down' ? 'bottom' : 'top';
      const borders = cell.borders && typeof cell.borders === 'object' ? cell.borders : {};
      /** @type {[string, unknown][]} */
      const sides = [
        ['borderTop', borders.top ?? (i === 0 ? outer.top : undefined)],
        ['borderBottom', borders.bottom ?? (i + rowSpan >= rows.length ? outer.bottom : undefined)],
        ['borderLeft', borders.left ?? (j === 0 ? outer.left : undefined)],
        ['borderRight', borders.right ?? (j + colSpan >= colCount ? outer.right : undefined)],
      ];
      for (const [prop, border] of sides) {
        const css = pptxCssBorder(border);
        if (css) /** @type {any} */ (td.style)[prop] = css;
      }
      const margin = cell.margin && typeof cell.margin === 'object' ? cell.margin : null;
      if (margin) td.style.padding = ['t', 'r', 'b', 'l'].map((k) => `${pptxR(Math.max(0, pptxNum(margin[k])))}pt`).join(' ');
      const text = h('div', { class: 'pptx-cell-text' });
      text.innerHTML = sanitize(typeof cell.text === 'string' ? cell.text : '', DOCX_SANITIZE_OPTIONS);
      pptxFixText(text);
      td.appendChild(text);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  box.appendChild(table);
  return box;
}

/**
 * CSS border of a pptxtojson Border ('' for none).
 * @param {unknown} border
 */
function pptxCssBorder(border) {
  if (!border || typeof border !== 'object') return '';
  const b = /** @type {Record<string, any>} */ (border);
  const width = pptxNum(b.borderWidth);
  const color = pptxColor(b.borderColor);
  if (!(width > 0) || !color) return '';
  const style = b.borderType === 'dashed' || b.borderType === 'dotted' ? b.borderType : 'solid';
  return `${pptxR(width)}pt ${style} ${color}`;
}

/**
 * A group (or SmartArt drawing): children in the group's coordinates; a flipped group mirrors their boxes.
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 */
function pptxBuildGroup(el, ctx, mirror) {
  const place = pptxPlace(el, mirror);
  const box = pptxBox(place, el.type === 'diagram' ? 'pptx-group pptx-diagram' : 'pptx-group');
  const childMirror = place.flipH || place.flipV ? { w: place.width, h: place.height, flipH: place.flipH, flipV: place.flipV } : null;
  for (const child of Array.isArray(el.elements) ? el.elements : []) {
    const node = pptxBuildElementSafe(child, ctx, childMirror);
    if (node) box.appendChild(node);
  }
  return box.childNodes.length ? box : null;
}

/**
 * Video / audio: a placeholder (media is never played or embedded).
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 */
function pptxBuildMedia(el, ctx, mirror) {
  const place = pptxPlace(el, mirror);
  const box = pptxBox(place, 'pptx-media');
  const video = el.mediaType !== 'audio';
  const label = video ? 'Video' : 'Audio';
  box.style.fontSize = `${pptxR(clamp(Math.min(place.width / 10, place.height / 4), 5, 18))}pt`;
  box.setAttribute('role', 'img');
  box.setAttribute('aria-label', `${label} (not played in FileStudio)`);
  box.title = `${label}: not played in FileStudio`;
  appendChildren(box, [
    icon(video ? PPTX_ICONS.film : PPTX_ICONS.speaker, 'fv-icon pptx-media-icon'),
    h('span', { class: 'pptx-media-label' }, label),
    ctx.thumb ? null : h('span', { class: 'pptx-media-note' }, 'not played in FileStudio'),
  ]);
  return box;
}

// ----- PPTX: CHARTS -----

/**
 * @param {PptxEl} el
 * @returns {string} 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter', '' when not drawn
 */
function pptxChartKind(el) {
  switch (el.chartType) {
    case 'barChart':
    case 'bar3DChart':
      return 'bar';
    case 'lineChart':
    case 'line3DChart':
      return 'line';
    case 'areaChart':
    case 'area3DChart':
      return 'area';
    case 'pieChart':
    case 'pie3DChart':
      return 'pie';
    case 'doughnutChart':
      return 'doughnut';
    case 'scatterChart':
      return 'scatter';
    default:
      return '';
  }
}

/**
 * @param {PptxEl} el
 */
function pptxChartName(el) {
  if ((el.chartType === 'barChart' || el.chartType === 'bar3DChart') && el.barDir !== 'bar') return el.chartType === 'barChart' ? 'Column chart' : '3-D column chart';
  return PPTX_CHART_NAMES[String(el.chartType)] || 'Chart';
}

/**
 * A chart: drawn by echarts once the slide is on screen (thumbnails show a light placeholder).
 * @param {PptxEl} el
 * @param {PptxBuildContext} ctx
 * @param {PptxMirror | null} mirror
 */
function pptxBuildChart(el, ctx, mirror) {
  const place = pptxPlace(el, mirror);
  const box = pptxBox(place, 'pptx-chart');
  const kind = pptxChartKind(el);
  if (!kind || ctx.thumb) {
    box.appendChild(pptxChartPlaceholder(el, !ctx.thumb, ctx.thumb));
    return box;
  }
  const chartTitle = typeof el.title === 'string' ? el.title.replace(/\s+/g, ' ').trim() : '';
  const host = h('div', { class: 'pptx-chart-host', role: 'img', 'aria-label': chartTitle ? `${pptxChartName(el)}: ${chartTitle}` : pptxChartName(el) }, h('span', { class: 'pptx-chart-loading' }, 'Loading chart…'));
  box.appendChild(host);
  ctx.charts.push({ el, host, kind, width: place.width * PPTX_PX_PER_PT, height: place.height * PPTX_PX_PER_PT, dark: ctx.dark });
  return box;
}

/**
 * Labelled box standing for a chart; with `withData`, a table of the chart's values.
 * @param {PptxEl} el
 * @param {boolean} withData
 * @param {boolean} thumb
 */
function pptxChartPlaceholder(el, withData, thumb) {
  const name = pptxChartName(el);
  return h(
    'div',
    { class: `pptx-chart-fallback${thumb ? ' pptx-chart-thumb' : ''}`, role: 'img', 'aria-label': name },
    h('div', { class: 'pptx-chart-fallback-title' }, icon(PPTX_ICONS.chart), thumb ? name : `${name} (not drawn by FileStudio)`),
    withData ? pptxChartDataTable(el) : null,
  );
}

/**
 * The chart's values as a small table (series × categories; scatter: X and Y columns).
 * @param {PptxEl} el
 */
function pptxChartDataTable(el) {
  const data = Array.isArray(el.data) ? el.data : [];
  /** @param {unknown} v */
  const cell = (v) => h('td', null, typeof v === 'number' ? (Number.isFinite(v) ? String(v) : '') : String(v ?? ''));
  const table = h('table', { class: 'pptx-chart-data' });
  if (data.length && Array.isArray(data[0])) {
    // Scatter / bubble: [xs, ys1, ys2, …]
    const columns = data.map((_, i) => (i === 0 ? 'X' : data.length > 2 ? `Y${i}` : 'Y'));
    table.appendChild(h('thead', null, h('tr', null, columns.map((c) => h('th', null, c)))));
    const length = Math.min(200, Math.max(...data.map((col) => (Array.isArray(col) ? col.length : 0))));
    const body = h('tbody');
    for (let r = 0; r < length; r++) body.appendChild(h('tr', null, data.map((col) => cell(Array.isArray(col) ? col[r] : ''))));
    table.appendChild(body);
    return h('div', { class: 'pptx-chart-data-wrap' }, table);
  }
  const categories = pptxChartCategories(data);
  table.appendChild(h('thead', null, h('tr', null, h('th', null, ''), categories.slice(0, 50).map((c) => h('th', null, c)))));
  const body = h('tbody');
  data.slice(0, 50).forEach((series, i) => {
    const values = Array.isArray(series?.values) ? series.values : [];
    body.appendChild(h('tr', null, h('th', null, pptxSeriesName(series, i)), values.slice(0, 50).map((/** @type {any} */ v) => cell(v?.y))));
  });
  table.appendChild(body);
  return h('div', { class: 'pptx-chart-data-wrap' }, table);
}

/**
 * Category labels of a chart, from its first series.
 * @param {any[]} data pptxtojson ChartItem[]
 * @returns {string[]}
 */
function pptxChartCategories(data) {
  const first = data.find((s) => Array.isArray(s?.values) && s.values.length);
  if (!first) return [];
  const labels = first.xlabels && typeof first.xlabels === 'object' ? first.xlabels : {};
  return first.values.map((/** @type {any} */ v, /** @type {number} */ i) => {
    const label = labels[v?.x] ?? labels[i];
    return label !== undefined && label !== null ? String(label) : String(i + 1);
  });
}

/**
 * @param {any} series
 * @param {number} i
 */
function pptxSeriesName(series, i) {
  const key = series?.key;
  return typeof key === 'string' && key.trim() ? key : `Series ${i + 1}`;
}

/** @returns {Promise<typeof import('echarts')>} */
function pptxLoadEcharts() {
  if (!pptxEchartsPromise) {
    pptxEchartsPromise = import('echarts').catch((err) => {
      pptxEchartsPromise = null; // try again with the next slide
      throw err;
    });
  }
  return pptxEchartsPromise;
}

/**
 * Draws the charts of the main slide (svg renderer: stays sharp under the slide's CSS scale).
 * @param {PptxChartJob[]} jobs
 * @param {() => boolean} isCurrent false once another slide replaced this one
 */
async function pptxRunCharts(jobs, isCurrent) {
  /** @type {typeof import('echarts')} */
  let echarts;
  try {
    echarts = await pptxLoadEcharts();
  } catch (err) {
    logError('Could not load the chart library', err);
    if (isCurrent()) for (const job of jobs) job.host.replaceChildren(pptxChartPlaceholder(job.el, true, false));
    return;
  }
  if (!isCurrent()) return;
  for (const job of jobs) {
    try {
      job.host.replaceChildren();
      const chart = echarts.init(job.host, null, {
        renderer: 'svg',
        width: Math.max(1, Math.round(job.width)),
        height: Math.max(1, Math.round(job.height)),
      });
      pptxState.charts.push(chart);
      chart.setOption(pptxChartOption(job));
    } catch (err) {
      logError(`Could not draw a chart on slide ${pptxState.current + 1}`, err);
      job.host.replaceChildren(pptxChartPlaceholder(job.el, true, false));
    }
  }
}

/**
 * echarts option of a chart (Office-like look: grey text, light grid lines, legend at the bottom).
 * @param {PptxChartJob} job
 * @returns {any}
 */
function pptxChartOption(job) {
  const el = job.el;
  const fg = job.dark ? '#D9D9D9' : '#595959';
  const gridColor = job.dark ? 'rgba(255, 255, 255, 0.18)' : '#D9D9D9';
  const axisColor = job.dark ? '#A6A6A6' : '#BFBFBF';
  const deckPalette = (pptxState.deck?.themeColors ?? []).map(pptxColor).filter(Boolean);
  const basePalette = deckPalette.length ? deckPalette : PPTX_DEFAULT_PALETTE;
  const chartColors = Array.isArray(el.colors) ? el.colors : [];
  /** @param {number} i */
  const colorAt = (i) => pptxColor(chartColors[i]) || basePalette[i % basePalette.length];
  const data = Array.isArray(el.data) ? el.data : [];
  // Chart settings read from the chart part (PptxChartExtras); a chart without them keeps the defaults.
  const fontPt = pptxNum(el.fontSize);
  const fontPx = fontPt > 0 ? pptxR(clamp(fontPt, 4, 72) * PPTX_PX_PER_PT) : 0;
  const textStyle = { color: fg, fontSize: fontPx || 13, fontFamily: PPTX_SANS_FALLBACK };
  const option = /** @type {Record<string, any>} */ ({
    animation: false,
    textStyle,
    tooltip: { trigger: 'item', confine: true },
  });
  const title = typeof el.title === 'string' ? el.title.trim() : '';
  const titlePt = pptxNum(el.titleSize);
  const titlePx = titlePt > 0 ? pptxR(clamp(titlePt, 4, 96) * PPTX_PX_PER_PT) : pptxR((fontPx || 13) * 1.2);
  const titleH = title ? Math.ceil(titlePx * 1.3 * title.split('\n').length) + 10 : 0;
  if (title) option.title = { text: title, left: 'center', top: 6, textStyle: { color: fg, fontSize: titlePx, fontWeight: 'normal' } };
  const legendPos = typeof el.legend === 'string' ? el.legend : '';
  const legendFont = fontPx || 12;
  const legend = { textStyle: { color: fg, fontSize: legendFont }, icon: 'rect', itemWidth: 10, itemHeight: 10 };
  /**
   * Legend of the chart (null: none) and the room it takes: PowerPoint's position when the chart part gives one,
   * else at the bottom when there is more than one series.
   * @param {string[]} names
   * @param {boolean} many
   */
  const legendFor = (names, many) => {
    if (legendPos === 'none' || (!legendPos && !many)) return null;
    const pos = legendPos || 'b';
    const vertical = pos === 'l' || pos === 'r' || pos === 'tr';
    const longest = Math.max(0, ...names.map((n) => String(n).length));
    const width = Math.ceil(26 + longest * legendFont * 0.55);
    /** @type {Record<string, any>} */
    const box = { ...legend, data: names };
    if (pos === 't') Object.assign(box, { top: titleH + 4 });
    else if (pos === 'b') Object.assign(box, { bottom: 4 });
    else Object.assign(box, { orient: 'vertical', top: pos === 'tr' ? titleH + 4 : 'middle', [pos === 'l' ? 'left' : 'right']: 8 });
    const room = vertical ? { side: pos === 'l' ? 'left' : 'right', size: width + 8 } : { side: pos === 't' ? 'top' : 'bottom', size: Math.ceil(legendFont * 1.4) + 14 };
    return { box, room };
  };
  /** @param {{ box: any, room: { side: string, size: number } } | null} withLegend */
  const gridFor = (withLegend) => {
    /** @type {Record<string, any>} */
    const g = { left: 12, right: 16, top: 16 + titleH, bottom: 12, containLabel: true };
    if (withLegend) g[withLegend.room.side] += withLegend.room.size;
    return g;
  };
  const gridlines = el.gridlines !== false;
  /**
   * @param {boolean} percent
   * @param {unknown} [format] Excel number format code of the axis
   */
  const valueAxis = (percent, format) => ({
    type: 'value',
    max: percent ? 100 : undefined,
    axisLabel: { color: fg, fontSize: fontPx || undefined, formatter: percent ? '{value}%' : pptxNumberFormatter(format) },
    axisLine: { show: false },
    splitLine: { show: gridlines, lineStyle: { color: gridColor } },
  });
  const dataLabels = Array.isArray(el.dataLabels) ? el.dataLabels : [];
  /**
   * Data labels of series `i` (PptxDataLabels), as an echarts label; undefined when the series shows none.
   * @param {number} i
   * @param {string} position
   * @param {(p: any) => number} valueOf the point's own value (before percent stacking)
   * @param {string} seriesName
   */
  const labelFor = (i, position, valueOf, seriesName) => {
    const spec = dataLabels[i];
    if (!spec || typeof spec !== 'object') return undefined;
    const format = pptxNumberFormatter(spec.format) ?? ((/** @type {number} */ v) => String(pptxR(v, 1e6)));
    return {
      show: true,
      position,
      color: fg,
      fontSize: fontPx || 12,
      formatter: (/** @type {any} */ p) => {
        const parts = [];
        if (spec.series) parts.push(seriesName);
        if (spec.category) parts.push(String(p.name ?? ''));
        if (spec.value) {
          const v = valueOf(p);
          if (Number.isFinite(v)) parts.push(format(v));
        }
        if (spec.percent && Number.isFinite(p.percent)) parts.push(`${Math.round(p.percent)}%`);
        return parts.join(', ');
      },
    };
  };

  if (job.kind === 'pie' || job.kind === 'doughnut') {
    const series = data[0];
    const categories = pptxChartCategories(data);
    const values = Array.isArray(series?.values) ? series.values : [];
    const hole = clamp(parseFloat(String(el.holeSize)) || 50, 10, 90);
    const pieLegend = legendFor(categories, true);
    if (pieLegend) option.legend = pieLegend.box;
    const side = pieLegend?.room.side;
    const radius = title || side === 'top' ? 62 : 70;
    const name = pptxSeriesName(series, 0);
    option.series = [
      {
        type: 'pie',
        name,
        radius: job.kind === 'doughnut' ? [`${pptxR((radius * hole) / 100)}%`, `${radius}%`] : `${radius}%`,
        center: [side === 'right' ? '42%' : side === 'left' ? '58%' : '50%', side === 'top' || title ? '54%' : '46%'],
        label: labelFor(0, 'inside', (p) => Number(p.value), name) ?? { show: false },
        itemStyle: { borderColor: job.dark ? '#262626' : '#FFFFFF', borderWidth: 1 },
        data: values.map((/** @type {any} */ v, /** @type {number} */ i) => ({
          name: categories[i] ?? String(i + 1),
          value: Number.isFinite(v?.y) ? v.y : 0,
          itemStyle: { color: colorAt(i) },
        })),
      },
    ];
    return option;
  }

  if (job.kind === 'scatter') {
    const xs = Array.isArray(data[0]) ? data[0] : [];
    const names = Array.isArray(el.seriesNames) ? el.seriesNames : [];
    const series = data.slice(1).map((ys, i) => {
      const name = typeof names[i] === 'string' && names[i].trim() ? names[i] : data.length > 2 ? `Series ${i + 1}` : 'Series 1';
      return {
        type: 'scatter',
        name,
        symbolSize: 8,
        itemStyle: { color: colorAt(i) },
        label: labelFor(i, 'right', (p) => Number(Array.isArray(p.value) ? p.value[1] : p.value), name),
        data: xs.map((/** @type {number} */ x, /** @type {number} */ j) => [x, Array.isArray(ys) ? ys[j] : null]),
      };
    });
    const scatterLegend = legendFor(
      series.map((s) => s.name),
      series.length > 1,
    );
    if (scatterLegend) option.legend = scatterLegend.box;
    option.grid = gridFor(scatterLegend);
    option.xAxis = { ...valueAxis(false), axisLine: { show: true, lineStyle: { color: axisColor } } };
    option.yAxis = valueAxis(false, el.valueFormat);
    option.series = series;
    return option;
  }

  // bar / line / area: categories × series
  const categories = pptxChartCategories(data);
  const grouping = String(el.grouping || '');
  const stacked = grouping === 'stacked' || grouping === 'percentStacked';
  const percent = grouping === 'percentStacked';
  /** @type {(number | null)[][]} */
  const rows = data.map((series) =>
    (Array.isArray(series?.values) ? series.values : []).map((/** @type {any} */ v) => (Number.isFinite(v?.y) ? v.y : null)),
  );
  const raw = rows.map((row) => row.slice());
  if (percent) {
    const totals = categories.map((_, c) => rows.reduce((sum, row) => sum + Math.abs(row[c] ?? 0), 0));
    rows.forEach((row) => row.forEach((v, c) => (row[c] = v === null || !totals[c] ? v : pptxR((v / totals[c]) * 100, 100))));
  }
  const horizontal = job.kind === 'bar' && el.barDir === 'bar';
  // A combo chart (PptxChartExtras.seriesTypes): each series keeps its own plot type and value axis.
  const seriesTypes = Array.isArray(el.seriesTypes) ? el.seriesTypes : [];
  const secondaryOf = Array.isArray(el.secondary) ? el.secondary : [];
  const kindOf = (/** @type {number} */ i) => (typeof seriesTypes[i] === 'string' ? pptxChartKind({ chartType: seriesTypes[i] }) : '') || job.kind;
  const kinds = data.map((_, i) => kindOf(i));
  const hasSecondary = data.some((_, i) => secondaryOf[i] === true);
  // Like PowerPoint: bar and line points sit between the tick marks, area charts reach the plot edges.
  const edgeToEdge = kinds.length > 0 && kinds.every((k) => k === 'area');
  const categoryAxis = {
    type: 'category',
    data: categories,
    boundaryGap: !edgeToEdge,
    axisLine: { lineStyle: { color: axisColor } },
    axisTick: { show: false },
    axisLabel: { color: fg, fontSize: fontPx || undefined },
  };
  const names = data.map((series, i) => pptxSeriesName(series, i));
  // Horizontal bars: PowerPoint lists the legend in reverse too (Series 3, 2, 1), like the bars from the top.
  const barLegend = legendFor(horizontal ? [...names].reverse() : names, data.length > 1);
  if (barLegend) option.legend = barLegend.box;
  // The last category label of an edge-to-edge axis is centred on the plot's right edge: leave room for it.
  const plotGrid = gridFor(barLegend);
  if (edgeToEdge && !(barLegend && barLegend.room.side === 'right')) plotGrid.right = 40;
  option.grid = plotGrid;
  const valueAxes = [valueAxis(percent, el.valueFormat)];
  if (hasSecondary) valueAxes.push({ ...valueAxis(false, el.valueFormat2), splitLine: { show: false, lineStyle: { color: gridColor } } });
  option.xAxis = horizontal ? valueAxes : categoryAxis;
  option.yAxis = horizontal ? categoryAxis : valueAxes;
  const series = data.map((_, i) => {
    const kind = kinds[i];
    const secondary = secondaryOf[i] === true;
    const label = labelFor(
      i,
      kind === 'bar' ? (stacked ? 'inside' : horizontal ? 'right' : 'top') : 'top',
      (p) => Number(raw[i]?.[p.dataIndex]),
      names[i],
    );
    const base = {
      name: names[i],
      data: rows[i],
      stack: stacked ? `total-${kind}-${secondary ? 2 : 1}` : undefined,
      itemStyle: { color: colorAt(i) },
      label,
      [horizontal ? 'xAxisIndex' : 'yAxisIndex']: secondary ? 1 : 0,
    };
    if (kind === 'bar') return { ...base, type: 'bar', barMaxWidth: 64 };
    if (kind === 'area') {
      return { ...base, type: 'line', showSymbol: false, lineStyle: { width: 1, color: colorAt(i) }, areaStyle: { color: colorAt(i), opacity: stacked ? 1 : 0.85 } };
    }
    return { ...base, type: 'line', showSymbol: el.marker === true || !!label, symbolSize: 7, lineStyle: { width: 3, color: colorAt(i) } };
  });
  // Horizontal bars: PowerPoint puts series 1 nearest the category axis (lowest in each group); echarts draws the
  // series of a vertical category axis top-down, so they are given in reverse.
  option.series = horizontal ? series.reverse() : series;
  return option;
}

/**
 * Formatter for an Excel number format code ('"₹"#,##0.00', '0%', '0.0,"K"', …) as chart axes and data labels use
 * them: literal text, currency, percent, decimals, thousands separators and scaling. Dates and other codes fall back
 * to the plain number. Undefined for General / no format.
 * @param {unknown} code
 * @returns {((value: number) => string) | undefined}
 */
function pptxNumberFormatter(code) {
  if (typeof code !== 'string' || !code.trim() || /^general$/i.test(code.trim())) return undefined;
  /** @type {string[]} */
  const sections = [''];
  for (const m of code.matchAll(/"[^"]*"|\\.|\[[^\]]*\]|;|[^"\\[;]+/g)) {
    if (m[0] === ';') sections.push('');
    else sections[sections.length - 1] += m[0];
  }
  /**
   * @param {string} section
   * @returns {((abs: number) => string) | null}
   */
  const compile = (section) => {
    let prefix = '';
    let suffix = '';
    let integer = '';
    let fraction = '';
    let inFraction = false;
    let seenDigit = false;
    let percent = 0;
    let scale = 0;
    let date = false;
    for (const m of section.matchAll(/"([^"]*)"|\\(.)|\[\$([^\]-]*)[^\]]*\]|\[[^\]]*\]|_.|\*.|([0#?])|(\.)|(,)|(%)|([eE][+-])|(.)/gu)) {
      const [, quoted, escaped, currency, digit, dot, comma, pct, exp, other] = m;
      /** @type {string | undefined} */
      let literal = quoted ?? escaped ?? currency;
      if (digit) {
        if (inFraction) fraction += digit;
        else integer += digit;
        seenDigit = true;
        suffix = '';
        scale = 0;
        continue;
      }
      if (dot && !inFraction) {
        inFraction = true;
        seenDigit = true;
        continue;
      }
      if (comma) {
        // Between digits: thousands separator; after the last digit: divides by 1000 (counted in `scale`).
        if (seenDigit && !inFraction) integer += ',';
        if (seenDigit) scale++;
        continue;
      }
      if (pct) {
        percent++;
        literal = '%';
      }
      if (exp) return null;
      if (other !== undefined) {
        if (/[ymdhsAa]/.test(other)) date = true;
        literal = other;
      }
      if (m[0].startsWith('_')) literal = ' ';
      if (literal === undefined) continue;
      if (seenDigit) suffix += literal;
      else prefix += literal;
    }
    if (date || (!integer && !fraction && !seenDigit)) return date ? null : () => `${prefix}${suffix}`;
    const grouped = /[0#?],[0#?]/.test(integer);
    const minInt = (integer.replace(/,/g, '').match(/0/g) ?? []).length;
    const maxDec = fraction.length;
    const minDec = (fraction.match(/0/g) ?? []).length;
    return (abs) => {
      const value = (abs * 100 ** percent) / 1000 ** scale;
      let [int, dec = ''] = value.toFixed(maxDec).split('.');
      while (dec.length > minDec && dec.endsWith('0')) dec = dec.slice(0, -1);
      if (int === '0' && minInt === 0) int = '';
      else if (int.length < minInt) int = int.padStart(minInt, '0');
      if (grouped) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      return `${prefix}${int}${dec ? `.${dec}` : ''}${suffix}`;
    };
  };
  const positive = compile(sections[0]);
  const negative = sections.length > 1 && sections[1] !== '' ? compile(sections[1]) : null;
  const zero = sections.length > 2 && sections[2] !== '' ? compile(sections[2]) : null;
  if (!positive) return undefined;
  return (value) => {
    if (!Number.isFinite(value)) return '';
    if (value === 0 && zero) return zero(0);
    if (value < 0) return negative ? negative(-value) : `-${positive(-value)}`;
    return positive(value);
  };
}

// ----- PPTX: VALUES -----

/**
 * Finite number or `fallback`.
 * @param {unknown} value
 * @param {number} [fallback]
 */
function pptxNum(value, fallback = 0) {
  const n = typeof value === 'number' ? value : typeof value === 'string' && value.trim() ? Number(value) : NaN;
  return Number.isFinite(n) ? n : fallback;
}

/**
 * Rounds for CSS / SVG output.
 * @param {number} value
 * @param {number} [precision] 100 = two decimals
 */
function pptxR(value, precision = 100) {
  const n = Math.round(value * precision) / precision;
  return Object.is(n, -0) ? 0 : n;
}

function pptxNextId() {
  pptxUid += 1;
  return `pptx-${pptxUid}`;
}

const PPTX_COLOR_RE = /^(?:#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})|(?:rgba?|hsla?)\([\d\s.,%+-]*\)|[a-z]{3,24})$/i;

/**
 * A colour from the deck if it is a plain colour value (hex, rgb[a](), hsl[a](), a colour name), else ''.
 * Never a url(…) or a CSS keyword that inherits.
 * @param {unknown} value
 */
function pptxColor(value) {
  const s = typeof value === 'string' ? value.trim() : '';
  if (!s || !PPTX_COLOR_RE.test(s) || /^(?:inherit|initial|unset|revert|currentcolor|none)$/i.test(s)) return '';
  return s;
}

/**
 * [r, g, b, a] (0–255, a 0–1) of a hex colour, null for other notations.
 * @param {string} color
 * @returns {number[] | null}
 */
function pptxRgba(color) {
  const m = /^#([0-9a-f]+)$/i.exec(color);
  if (!m) return null;
  let hex = m[1];
  if (hex.length === 3 || hex.length === 4) hex = hex.replace(/./g, (c) => c + c);
  if (hex.length !== 6 && hex.length !== 8) return null;
  const n = (/** @type {number} */ i) => Number.parseInt(hex.slice(i, i + 2), 16);
  return [n(0), n(2), n(4), hex.length === 8 ? n(6) / 255 : 1];
}

/**
 * `a` × amount + `b` × (1 - amount) (hex colours; otherwise `a`).
 * @param {string} a
 * @param {string} b
 * @param {number} amount
 */
function pptxMixColors(a, b, amount) {
  const ca = pptxRgba(a);
  const cb = pptxRgba(b);
  if (!ca || !cb) return a;
  const mix = (/** @type {number} */ i) => Math.round(ca[i] * amount + cb[i] * (1 - amount));
  return `rgb(${mix(0)}, ${mix(1)}, ${mix(2)})`;
}

/**
 * Approximate luminance 0–1 of a hex colour (null if unknown or mostly transparent).
 * @param {string} color
 * @returns {number | null}
 */
function pptxLuminance(color) {
  const c = color ? pptxRgba(color) : null;
  if (!c || c[3] < 0.5) return null;
  return (0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]) / 255;
}

/**
 * The value if it is a base64 data: URI of an image (pictures are never loaded from anywhere else), else ''.
 * @param {unknown} value
 */
function pptxImageUri(value) {
  if (typeof value !== 'string' || value.length < 24) return '';
  return /^data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/=\r\n]*$/i.test(value) ? value : '';
}

/**
 * SVG element builder (attributes with null / undefined / false are skipped).
 * @param {string} tag
 * @param {Record<string, string | number | boolean | null | undefined> | null} [attrs]
 * @param {...(Node | null)} children
 * @returns {SVGElement}
 */
function pptxSvg(tag, attrs, ...children) {
  const el = /** @type {SVGElement} */ (document.createElementNS(PPTX_SVG_NS, tag));
  if (attrs) {
    for (const [key, value] of Object.entries(attrs)) {
      if (value === null || value === undefined || value === false) continue;
      el.setAttribute(key, String(value));
    }
  }
  for (const child of children) if (child) el.appendChild(child);
  return el;
}
