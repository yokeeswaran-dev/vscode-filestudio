// Webview regression harness: the REAL built webview (dist/viewer.js + dist/viewer.css + media/viewer.css) in headless
// Chrome / Edge, behind a mock extension host that answers like src/viewerProvider.ts with the REAL renderers
// (test/.out/*.js, bundled by test/build-tests.js).
//
// - The page is built like getHtml() in viewerProvider.ts: same CSP (taken from buildCsp in the source), same nonce
//   rule, same <link> / <script> order, body.vscode-<theme> class and real VS Code theme variables (themes.json).
// - acquireVsCodeApi() is a stub: postMessage goes to the mock host in Node (page.exposeFunction), the replies come
//   back as window 'message' events in order, like VS Code's message channel. getState / setState use sessionStorage,
//   so a page reload keeps the view state like a webview reload does.
// - No browser download: the browser comes from FS_CHROME (only that one when set), else PUPPETEER_EXECUTABLE_PATH,
//   CHROME_BIN (set on GitHub's runners) or CHROME_PATH, else the usual Chrome / Edge / Chromium install paths. Without
//   one the suites are SKIPPED with a message on a developer PC, but FAIL on CI (CI set) or with FS_WEBVIEW_REQUIRE_BROWSER=1.
//
// Environment (all optional): FS_CHROME (browser executable), FS_DIST (another built dist/ folder, for example a build
// with a fix reverted), FS_WEBVIEW_REQUIRE_BROWSER=1|0 (fail | skip without a browser; default: fail only on CI),
// FS_WEBVIEW_HEADFUL=1 (show the browser), FS_WEBVIEW_SLOWMO=<ms>.
'use strict';
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.join(__dirname, '..', '..');
const DIST = path.resolve(process.env.FS_DIST || path.join(ROOT, 'dist'));
const MEDIA = path.join(ROOT, 'media');
const SAMPLES = path.join(ROOT, 'test', 'samples');
const OUT = path.join(ROOT, 'test', '.out');
const THEMES = require('./themes.json');
const THEME_CLASS = { light: 'vscode-light', dark: 'vscode-dark', hc: 'vscode-high-contrast', hcl: 'vscode-high-contrast-light vscode-high-contrast' };
const TYPES = {
  '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.ttf': 'font/ttf', '.woff': 'font/woff',
  '.woff2': 'font/woff2', '.svg': 'image/svg+xml', '.png': 'image/png', '.gif': 'image/gif', '.json': 'application/json', '.map': 'application/json',
  '.pdf': 'application/pdf', '.md': 'text/markdown',
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ===== BROWSER =====

/** First existing browser executable, or undefined. FS_CHROME, when set, is the only candidate. */
function findBrowser() {
  if (process.env.FS_CHROME) return fs.existsSync(process.env.FS_CHROME) ? process.env.FS_CHROME : undefined;
  for (const name of ['PUPPETEER_EXECUTABLE_PATH', 'CHROME_BIN', 'CHROME_PATH']) {
    const value = process.env[name];
    if (value && fs.existsSync(value)) return value;
  }
  const list = [];
  for (const root of [process.env.PROGRAMFILES, process.env['PROGRAMFILES(X86)'], process.env.LOCALAPPDATA].filter(Boolean)) {
    list.push(path.join(root, 'Google', 'Chrome', 'Application', 'chrome.exe'), path.join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe'));
  }
  list.push(
    '/usr/bin/google-chrome', '/usr/bin/google-chrome-stable', '/usr/bin/chromium', '/usr/bin/chromium-browser', '/usr/bin/microsoft-edge',
    '/snap/bin/chromium', '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
  );
  return list.find((p) => fs.existsSync(p));
}

const BROWSER = findBrowser();
/** Reason to skip every webview suite (no browser), or false. node:test prints it next to each skipped suite. */
const SKIP = BROWSER
  ? false
  : process.env.FS_CHROME
    ? `FS_CHROME=${process.env.FS_CHROME} does not exist: no browser for the webview tests`
    : 'no Chrome / Edge / Chromium found: set FS_CHROME=<path of chrome or msedge> to run the webview tests';
/**
 * A missing browser is a FAILURE (not a skip) when FS_WEBVIEW_REQUIRE_BROWSER=1, or on CI (CI is set) unless
 * FS_WEBVIEW_REQUIRE_BROWSER=0: a runner image without Chrome must not turn the regression net off silently.
 */
const REQUIRE_BROWSER = process.env.FS_WEBVIEW_REQUIRE_BROWSER === '1' || (!!process.env.CI && process.env.FS_WEBVIEW_REQUIRE_BROWSER !== '0');
if (SKIP && !globalThis.__fsSkipNoted) {
  globalThis.__fsSkipNoted = true;
  console.log(REQUIRE_BROWSER ? `[webview] FAILED: ${SKIP} (a browser is required here: CI or FS_WEBVIEW_REQUIRE_BROWSER=1)` : `[webview] SKIPPED: ${SKIP}`);
}

// ===== RENDERERS AND CSP =====

/** A renderer bundle from test/.out (made by `node test/build-tests.js`, which `npm run test:webview` runs first). */
function renderer(name) {
  const file = path.join(OUT, `${name}.js`);
  if (!fs.existsSync(file)) throw new Error(`${file} is missing: run "node test/build-tests.js" first (npm run test:webview does).`);
  return require(file);
}

let cspBuilder;
/**
 * The webview CSP exactly as viewerProvider.ts builds it: the buildCsp function is taken from the source (it is pure),
 * so the harness can never drift from the extension's policy.
 */
function buildCsp(cspSource, nonce, kind) {
  if (!cspBuilder) {
    const src = fs.readFileSync(path.join(ROOT, 'src', 'viewerProvider.ts'), 'utf8');
    const m = /export function buildCsp\(cspSource: string, nonce: string, kind: ViewKind\): string \{\r?\n([\s\S]*?)\r?\n\}\r?\n/.exec(src);
    if (!m) throw new Error('test/webview/harness.js: buildCsp not found in src/viewerProvider.ts (update the harness)');
    cspBuilder = new Function('cspSource', 'nonce', 'kind', m[1]);
  }
  return cspBuilder(cspSource, nonce, kind);
}

// ===== HTTP SERVER =====

/** Extra files served at /res/<name> (PDF bytes, markdown images): name -> Buffer. */
const resources = new Map();

function page(kind, fileName, theme, origin) {
  const nonce = crypto.randomBytes(16).toString('base64');
  const csp = buildCsp(origin, nonce, kind).replace(/"/g, '&quot;');
  const vars = Object.entries(THEMES[theme] || THEMES.light).map(([k, v]) => `${k}: ${v};`).join(' ').replace(/"/g, '&quot;');
  const cls = THEME_CLASS[theme] || THEME_CLASS.light;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  // The stub runs before viewer.js (inline + nonce, like VS Code's own preload). Replies are delivered in order.
  return `<!DOCTYPE html>
<html lang="en" style="${vars}">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(fileName)}</title>
<style id="_defaultStyles">
  html { scrollbar-color: var(--vscode-scrollbarSlider-background) var(--vscode-editor-background); }
  body { background-color: transparent; color: var(--vscode-editor-foreground); font-family: var(--vscode-font-family); font-weight: var(--vscode-font-weight); font-size: var(--vscode-font-size); margin: 0; padding: 0 20px; }
  img, video { max-width: 100%; max-height: 100%; }
  a, a code { color: var(--vscode-textLink-foreground); }
  code { font-family: var(--monaco-monospace-font); color: var(--vscode-textPreformat-foreground); }
</style>
<script nonce="${nonce}">
(function () {
  const key = 'fs-state:' + ${JSON.stringify(fileName)};
  window.__posted = [];
  window.__received = [];
  // __fsIdle(): no message to the mock host is waiting for its reply, and the webview has handled every delivered
  // reply. The counting listener is added at the first delivery, so after viewer.js's own listener: it runs once the
  // webview handled the message (and a follow-up request, for example the next step of a Ctrl+Arrow search, is sent).
  let hostBusy = 0;
  let delivered = 0;
  let handled = 0;
  let counting = false;
  let chain = Promise.resolve();
  const deliver = (msgs) => {
    if (!counting) {
      counting = true;
      window.addEventListener('message', (e) => { if (e.data && typeof e.data.type === 'string') handled++; });
    }
    for (const m of msgs) { window.__received.push(m); delivered++; window.postMessage(m, '*'); }
  };
  window.__fsDeliver = deliver;
  window.__fsIdle = () => hostBusy === 0 && handled >= delivered;
  window.acquireVsCodeApi = function () {
    return {
      postMessage(msg) {
        const json = JSON.stringify(msg);
        window.__posted.push(JSON.parse(json));
        hostBusy++;
        chain = chain
          .then(() => window.__fsHost(json))
          .then((r) => deliver(JSON.parse(r)))
          .catch((e) => console.error('mock host failure', e))
          .finally(() => hostBusy--);
      },
      getState() { try { return JSON.parse(sessionStorage.getItem(key) || 'null'); } catch { return null; } },
      setState(s) { sessionStorage.setItem(key, JSON.stringify(s)); },
    };
  };
})();
</script>
<link rel="stylesheet" href="/dist/viewer.css">
<link rel="stylesheet" href="/media/viewer.css">
</head>
<body data-kind="${kind}" class="${cls}" data-vscode-theme-kind="${cls.split(' ')[0]}">
<div id="app">
  <div class="fv-boot" role="status" aria-live="polite"><p class="fv-boot-message">Loading ${esc(fileName)}…</p></div>
</div>
<script type="module" nonce="${nonce}" src="/dist/viewer.js"></script>
</body>
</html>`;
}

function startServer() {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      const url = new URL(req.url, 'http://x');
      const p = decodeURIComponent(url.pathname);
      if (p === '/view') {
        let html;
        try {
          html = page(url.searchParams.get('kind'), url.searchParams.get('file'), url.searchParams.get('theme') || 'light', `http://${req.headers.host}`);
        } catch (err) {
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end(String(err && err.stack));
          return;
        }
        res.writeHead(200, { 'Content-Type': 'text/html', 'Cache-Control': 'no-store' });
        res.end(html);
        return;
      }
      if (p.startsWith('/res/') && resources.has(p.slice(5))) {
        const name = p.slice(5);
        res.writeHead(200, { 'Content-Type': TYPES[path.extname(name).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-store' });
        res.end(resources.get(name));
        return;
      }
      const roots = { '/dist/': DIST, '/media/': MEDIA, '/res/': SAMPLES };
      const prefix = Object.keys(roots).find((k) => p.startsWith(k));
      const file = prefix && path.join(roots[prefix], p.slice(prefix.length));
      if (!file || !file.startsWith(roots[prefix] + path.sep)) {
        res.writeHead(404);
        res.end();
        return;
      }
      fs.readFile(file, (err, data) => {
        if (err) {
          res.writeHead(404);
          res.end('not found');
          return;
        }
        res.writeHead(200, { 'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-store' });
        res.end(data);
      });
    });
    server.listen(0, '127.0.0.1', () => resolve(server));
  });
}

// ===== HARNESS =====

/**
 * One browser + one server for a test file. `open(host, opts)` shows a view of the mock `host` in a new page.
 * Usage (node:test): const h = await startHarness(); ... await h.close();
 */
async function startHarness() {
  if (!fs.existsSync(path.join(DIST, 'viewer.js'))) throw new Error(`${path.join(DIST, 'viewer.js')} is missing: run "npm run build" first.`);
  const puppeteer = require('puppeteer-core');
  const server = await startServer();
  const base = `http://127.0.0.1:${server.address().port}`;
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'fs-webview-'));
  const browser = await puppeteer.launch({
    executablePath: BROWSER,
    headless: process.env.FS_WEBVIEW_HEADFUL ? false : true,
    slowMo: Number(process.env.FS_WEBVIEW_SLOWMO) || 0,
    userDataDir: profile,
    args: ['--no-sandbox', '--font-render-hinting=none', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions'],
  });
  await browser.defaultBrowserContext().overridePermissions(base, ['clipboard-read', 'clipboard-write', 'clipboard-sanitized-write']);

  /**
   * @param {MockHost} host
   * @param {{ theme?: string, width?: number, height?: number, allowErrors?: boolean, ignoreErrors?: RegExp, before?: (page: any) => Promise<void> }} [opts]
   *   allowErrors: page errors do not fail the test; ignoreErrors: only the errors matching it are expected
   */
  async function open(host, opts = {}) {
    const { theme = 'light', width = 1280, height = 800 } = opts;
    const pg = await browser.newPage();
    await pg.setViewport({ width, height, deviceScaleFactor: 1 });
    const errors = [];
    pg.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
    pg.on('console', (m) => {
      // A failed load is reported with its URL by the 'response' listener below.
      if (m.type() === 'error' && !/^Failed to load resource/.test(m.text())) errors.push(`console.error: ${m.text()}`);
    });
    pg.on('response', (r) => {
      if (r.status() >= 400 && !/\/favicon\.ico$/.test(r.url())) errors.push(`HTTP ${r.status()}: ${r.url()}`);
    });
    await pg.exposeFunction('__fsHost', async (json) => JSON.stringify(await host.receive(JSON.parse(json))));
    await pg.evaluateOnNewDocument(() => {
      window.__csp = [];
      document.addEventListener('securitypolicyviolation', (e) => window.__csp.push(`${e.violatedDirective} ${e.blockedURI}`));
      // Every clipboard write of a copy event: { 'text/plain': ..., 'text/html': ... } of the last copy.
      window.__copied = {};
      document.addEventListener(
        'copy',
        (e) => {
          const dt = e.clipboardData;
          if (!dt) return;
          const set = dt.setData.bind(dt);
          dt.setData = (type, data) => {
            window.__copied[type] = data;
            set(type, data);
          };
        },
        true,
      );
    });
    if ('origin' in host) host.origin = base;
    if (opts.before) await opts.before(pg);
    const url = `${base}/view?kind=${host.kind}&file=${encodeURIComponent(host.fileName)}&theme=${theme}`;
    const view = new View(pg, host, url, errors, opts);
    await view.load();
    return view;
  }

  async function close() {
    await browser.close().catch(() => {});
    server.close();
    try {
      fs.rmSync(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
    } catch {
      // the browser still holds a file: left in the temp folder
    }
  }
  /** Opens a view, runs fn(view), closes it (close() fails the test on page errors, unless fn already failed). */
  async function use(host, opts, fn) {
    const view = await open(host, opts);
    try {
      await fn(view);
    } catch (err) {
      await view.page.close().catch(() => {});
      throw err;
    }
    await view.close();
  }
  return { open, use, close, browser, base };
}

/**
 * A node:test suite with one browser for all its tests. Without a browser it is one skipped test (with the reason,
 * counted in the summary), or one failing test when a browser is required (REQUIRE_BROWSER).
 * body({ test, ctx }): test(title, fn(h), { timeout }) where h = { open, use, ... } of startHarness().
 */
function suite(name, body) {
  const { describe, it, before, after } = require('node:test');
  if (SKIP) {
    describe(name, () => {
      if (REQUIRE_BROWSER) {
        it('a browser for the webview tests', () => {
          throw new Error(`${SKIP}. A browser is required here (CI or FS_WEBVIEW_REQUIRE_BROWSER=1); set FS_WEBVIEW_REQUIRE_BROWSER=0 to skip instead.`);
        });
      } else it('all tests of this suite', { skip: SKIP }, () => {});
    });
    return;
  }
  describe(name, () => {
    const ctx = { h: undefined };
    before(async () => {
      ctx.h = await startHarness();
    });
    after(async () => {
      if (ctx.h) await ctx.h.close();
    });
    const test = (title, fn, opts = {}) => it(title, { timeout: opts.timeout || 90000 }, () => fn(ctx.h));
    body({ test, ctx });
  });
}

/** A page showing one view. */
class View {
  constructor(page, host, url, errors, opts) {
    this.page = page;
    this.host = host;
    this.url = url;
    this.errors = errors;
    this.opts = opts;
  }

  async load(nav = () => this.page.goto(this.url, { waitUntil: 'load' })) {
    await nav();
    await this.page.waitForFunction(() => document.querySelector('.fv-scroller, .fv-empty, .md-view, .docx-view, .pdf-view, .fv-error, .pptx-view, [class*="pdf"], [class*="pptx"]') && !document.querySelector('.fv-boot'), { timeout: 60000, polling: 50 });
    await this.settle();
  }

  /** Reload like a webview reload (view state kept through getState / setState). */
  async reload() {
    await this.load(() => this.page.reload({ waitUntil: 'load' }));
  }

  /**
   * Waits until no rows are loading, the mock host is idle (no open request, every reply handled: __fsIdle) and two
   * frames have been drawn. So it also waits for the async keyboard searches (Ctrl+Arrow, Ctrl+A over rows that are
   * not loaded), which ask the host for rows, step by step.
   */
  async settle(ms = 6000) {
    await this.page.evaluate(async (ms) => {
      const t0 = performance.now();
      const raf = () => new Promise((r) => requestAnimationFrame(() => r()));
      for (;;) {
        await raf();
        await new Promise((r) => setTimeout(r, 20));
        const idle = !document.querySelectorAll('.fv-loading-rows').length && (!window.__fsIdle || window.__fsIdle());
        if (idle || performance.now() - t0 > ms) break;
      }
      await raf();
      await raf();
    }, ms);
  }

  /** Every webview -> host message since the page loaded (or since `clearPosted`). */
  posted(type) {
    return this.page.evaluate((type) => window.__posted.filter((m) => !type || m.type === type), type);
  }

  clearPosted() {
    return this.page.evaluate(() => (window.__posted = []));
  }

  /** Host -> webview messages (an external change, a late reply ...), delivered like VS Code does. */
  push(...msgs) {
    return this.page.evaluate((msgs) => window.__fsDeliver(msgs), msgs);
  }

  eval(fn, ...args) {
    return this.page.evaluate(fn, ...args);
  }

  /** Closes the page; fails when the page threw, logged a console error or hit the CSP (unless allowErrors). */
  async close() {
    const csp = await this.page.evaluate(() => window.__csp).catch(() => []);
    await this.page.close().catch(() => {});
    const ignore = this.opts.ignoreErrors;
    const problems = [...this.errors, ...csp.map((c) => `CSP violation: ${c}`)].filter((p) => !(ignore && ignore.test(p)));
    if (problems.length && !this.opts.allowErrors) throw new Error(`webview errors:\n${problems.join('\n')}`);
  }
}

// ===== MOCK HOSTS =====
// A host has `kind`, `fileName` and `receive(msg) -> HostMessage[]` (the replies to one webview message). Each answers
// like the matching provider in viewerProvider.ts. `log` keeps the 'log' messages the webview sent.

const sheetModels = new Map();

/**
 * xlsx view: `build` returns the workbook bytes (an ExcelJS buffer from fixtures.js); loaded once per fileName with
 * the real loadWorkbook / workbookToModel / detectLossyFeatures.
 * @param {string} fileName
 * @param {() => Promise<Uint8Array>} build
 * @param {{ meta?: (meta: any) => any, banner?: string }} [opts] meta: changes the meta sent in init (e.g. no worksheets);
 *   banner: the lossy banner text to send instead of the one detectLossyFeatures gives
 */
function sheetHost(fileName, build, opts = {}) {
  const S = renderer('sheet');
  if (!sheetModels.has(fileName)) {
    sheetModels.set(
      fileName,
      (async () => {
        const data = new Uint8Array(await build());
        const [wb, lossy] = await Promise.all([S.loadWorkbook(data), S.detectLossyFeatures(data).catch(() => [])]);
        return { model: S.workbookToModel(wb), lossy };
      })(),
    );
  }
  return gridHost('sheet', fileName, () => sheetModels.get(fileName), opts);
}

/** csv / tsv view of `text` (delimiter by extension, like viewerProvider.ts). */
function csvHost(fileName, text) {
  const S = renderer('sheet');
  const ext = path.extname(fileName).toLowerCase();
  const delimiter = ext === '.tsv' ? '\t' : ext === '.psv' ? '|' : ',';
  const csv = S.parseCsv(text, ext === '.csv' ? undefined : { delimiter });
  const loaded = Promise.resolve({ model: new S.CsvGridModel(csv, fileName), lossy: [], delimited: { format: ext.slice(1), delimiter: csv.delimiter } });
  return gridHost('csv', fileName, () => loaded, {});
}

function gridHost(kind, fileName, load, opts) {
  const host = {
    kind,
    fileName,
    log: [],
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      const { model, lossy, delimited } = await load();
      if (msg.type === 'ready') {
        const meta = opts.meta ? opts.meta(model.getMeta()) : model.getMeta();
        const banner = opts.banner !== undefined ? opts.banner : lossy.length ? `This workbook contains content that FileStudio does not display: ${lossy.join(', ')}.` : undefined;
        return [{ type: 'init', kind, fileName, meta, readOnly: true, banner, ...(delimited ? { delimited } : {}) }];
      }
      if (msg.type === 'getRows') {
        const { rows, styles } = model.getRows(msg.sheet, msg.start, msg.end, msg.c0, msg.c1);
        const cols = msg.c0 === undefined ? {} : { c0: msg.c0, c1: msg.c1 };
        return [{ type: 'rows', reqId: msg.reqId, sheet: msg.sheet, start: msg.start, end: msg.end, ...cols, rows, styles }];
      }
      if (msg.type === 'getStats') return [{ type: 'stats', reqId: msg.reqId, stats: model.getStats(msg.sheet, msg.ranges) }];
      return [];
    },
    /** The model (for invalidate tests). */
    model: async () => (await load()).model,
  };
  return host;
}

/**
 * Markdown view of `text`, with MarkdownView.toggleTask's rules: a toggle of an older version is refused
 * (toggleTaskResult applied:false), an applied one edits the marker, re-renders (markdownUpdate) and answers.
 * `host.text` / `host.version` can be changed by a test (an edit in the text editor).
 */
function markdownHost(fileName, text) {
  const M = renderer('markdown');
  const host = {
    kind: 'markdown',
    fileName,
    text,
    version: 1,
    log: [],
    rendered: undefined,
    render() {
      const res = M.renderMarkdown(host.text, {
        resolveResource: (href) => (/^[a-z][a-z0-9+.-]*:/i.test(href) || href.startsWith('//') ? undefined : `/res/${href.split(/[?#]/)[0]}`),
      });
      host.rendered = { version: host.version, tasks: new Map((res.tasks || []).map((t) => [t.line, t])) };
      return res;
    },
    /** The markdownUpdate the extension would send for the current text. */
    update() {
      const res = host.render();
      return { type: 'markdownUpdate', html: res.html, toc: res.toc, source: host.text, version: host.version };
    },
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      if (msg.type === 'ready') {
        const res = host.render();
        return [{ type: 'init', kind: 'markdown', fileName, html: res.html, toc: res.toc, source: host.text, mode: 'preview', version: host.version }];
      }
      if (msg.type === 'toggleTask') {
        const answer = (applied, reason) => ({ type: 'toggleTaskResult', line: msg.line, version: msg.version, applied, documentVersion: host.version, ...(reason ? { reason } : {}) });
        if (msg.version !== host.version) return [answer(false, 'the document changed since the click')];
        const task = host.rendered && host.rendered.version === msg.version ? host.rendered.tasks.get(msg.line) : undefined;
        const lines = host.text.split('\n');
        const line = task && lines[msg.line] !== undefined ? lines[msg.line] : '';
        if (!task || !/^\[[ xX]\]$/.test(line.slice(task.column - 1, task.column + 2))) return [answer(false, 'no task checkbox on that line')];
        if ((line[task.column] !== ' ') === msg.checked) return [answer(true)];
        lines[msg.line] = line.slice(0, task.column) + (msg.checked ? 'x' : ' ') + line.slice(task.column + 1);
        host.text = lines.join('\n');
        host.version++;
        return [host.update(), answer(true)];
      }
      if (msg.type === 'openLink') {
        // Like openLink's onSameDocument: a link to this very file scrolls the view.
        const m = /^(?:\.\/)?([^#?]*)(?:\?[^#]*)?(?:#(.*))?$/.exec(String(msg.href).trim());
        if (m && m[1] && decodeURIComponent(m[1]).replace(/^.*[\\/]/, '') === fileName) return [{ type: 'scrollToFragment', fragment: decodeURIComponent(m[2] || '') }];
      }
      return [];
    },
  };
  return host;
}

/** Word view: `build` returns the .docx bytes; rendered with the real renderDocx. */
function docxHost(fileName, build) {
  const D = renderer('docx');
  const host = {
    kind: 'docx',
    fileName,
    log: [],
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      if (msg.type === 'ready') {
        const res = await D.renderDocx(new Uint8Array(await build()));
        return [{ type: 'init', kind: 'docx', fileName, html: res.html, warnings: res.warnings }];
      }
      return [];
    },
  };
  return host;
}

/** PDF view: header check like PdfViewerProvider, then the bytes served at /res/<fileName> (pdf.js loads them). */
function pdfHost(fileName, build) {
  const P = renderer('pdf');
  const host = {
    kind: 'pdf',
    fileName,
    log: [],
    origin: '',
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      if (msg.type === 'ready') {
        const data = Buffer.from(await build());
        const check = P.checkPdf(new Uint8Array(data.subarray(0, 1024)));
        if (check.error !== undefined) return [{ type: 'error', message: `Could not open "${fileName}": ${check.error}`, canReopenAsText: false }];
        resources.set(fileName, data);
        return [{ type: 'init', kind: 'pdf', fileName, source: { uri: `${host.origin}/res/${encodeURIComponent(fileName)}` }, fileSize: data.byteLength }];
      }
      return [];
    },
  };
  return host;
}

/** A sample file of test/samples (bytes). */
const sample = (name) => fs.readFileSync(path.join(SAMPLES, name));

// ===== GRID HELPERS =====
// Cells are addressed like the name box ("C5") or by 0-based (row, col). All helpers wait for the grid to settle.

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

/** In-page helpers (window.__g). Installed by `grid(view)`. */
function installGridHelpers() {
  if (window.__g) return;
  const vis = (sel) => [...document.querySelectorAll(sel)].filter((e) => e.style.display !== 'none' && e.offsetParent !== null);
  const colName = (c) => {
    let s = '';
    let n = c + 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = (n - m - 1) / 26;
    }
    return s;
  };
  const mid = (r) => ({ x: r.left + r.width / 2, y: r.top + r.height / 2 });
  // VS Code's webview host forwards every keydown that reaches window (bubble phase) to the workbench.
  window.__fwd = [];
  window.addEventListener('keydown', (e) => window.__fwd.push(`${e.ctrlKey ? 'Ctrl+' : ''}${e.shiftKey ? 'Shift+' : ''}${e.key}`));
  window.__g = {
    colName,
    hdr(kind, label) {
      return vis(kind === 'row' ? '.fv-rh' : '.fv-ch').find((e) => e.textContent === label) || null;
    },
    /** Page rect of a cell from its row and column headers. */
    cellRect(r, c) {
      const ch = this.hdr('col', colName(c));
      const rh = this.hdr('row', String(r + 1));
      if (!ch || !rh) return null;
      const a = ch.getBoundingClientRect();
      const b = rh.getBoundingClientRect();
      return { x: a.left, y: b.top, w: a.width, h: b.height, r: a.right, b: b.bottom, cx: a.left + a.width / 2, cy: b.top + b.height / 2 };
    },
    hdrXY(kind, label) {
      const e = this.hdr(kind, label);
      return e ? mid(e.getBoundingClientRect()) : null;
    },
    /** Labels of the visible whole-row / whole-column headers. */
    full(kind) {
      return vis(kind === 'row' ? '.fv-rh.fv-hfull' : '.fv-ch.fv-hfull').map((e) => e.textContent);
    },
    allFull() {
      return vis('.fv-ch').every((e) => e.classList.contains('fv-hfull')) && vis('.fv-rh').every((e) => e.classList.contains('fv-hfull'));
    },
    /** Active cell address from the ARIA grid cell. */
    active() {
      const g = document.querySelector('.fv-scroller');
      const cell = g && document.getElementById(g.getAttribute('aria-activedescendant') || '');
      const label = cell && cell.getAttribute('aria-label');
      return label ? label.split(',')[0] : null;
    },
    focus() {
      const el = document.activeElement;
      if (!el || el === document.body) return 'BODY';
      return (typeof el.className === 'string' && el.className ? el.className.split(' ')[0] : el.tagName.toLowerCase()) + (el.getAttribute('aria-label') ? `[${el.getAttribute('aria-label')}]` : '');
    },
    /** The cell element whose text includes `text`. */
    cellEl(text) {
      return [...document.querySelectorAll('.fv-sub-cells > .fv-c, .fv-c')].find((e) => e.textContent.includes(text)) || null;
    },
    /** Union of the client rects of the text nodes inside an element. */
    textBox(el) {
      const tw = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let box = null;
      for (let n = tw.nextNode(); n; n = tw.nextNode()) {
        if (!n.textContent.trim()) continue;
        const range = document.createRange();
        range.selectNodeContents(n);
        for (const r of range.getClientRects()) {
          box = box ? { x: Math.min(box.x, r.left), y: Math.min(box.y, r.top), r: Math.max(box.r, r.right), b: Math.max(box.b, r.bottom) } : { x: r.left, y: r.top, r: r.right, b: r.bottom };
        }
      }
      return box;
    },
    /** Is the first scrolling row flush with the bottom of the top band (headers + `fr` frozen rows)? */
    rowAligned(fr) {
      const rh = vis('.fv-rh');
      const band = fr ? rh.find((e) => e.textContent === String(fr)).getBoundingClientRect().bottom : Math.max(...vis('.fv-ch').map((e) => e.getBoundingClientRect().bottom));
      const edge = rh.filter((e) => Number(e.textContent) > fr).map((e) => e.getBoundingClientRect()).find((r) => r.top <= band + 0.5 && r.bottom > band + 0.5);
      return !!edge && Math.abs(edge.top - band) < 1;
    },
    tabs() {
      return [...document.querySelectorAll('.fv-tab')].map((e) => e.textContent + (e.classList.contains('fv-tab-hidden') ? '[h]' : '') + (e.getAttribute('aria-selected') === 'true' ? '*' : ''));
    },
  };
}

/** Grid actions and readings for a sheet / csv view. */
function grid(view) {
  const p = view.page;
  const ready = () => p.evaluate(installGridHelpers);
  const withMods = async (mods, fn) => {
    for (const m of mods) await p.keyboard.down(m);
    try {
      await fn();
    } finally {
      for (const m of mods.slice().reverse()) await p.keyboard.up(m);
    }
  };
  const g = {
    view,
    page: p,
    ready,
    async press(key, mods = [], times = 1) {
      for (let i = 0; i < times; i++) await withMods(mods, () => p.keyboard.press(key));
      await view.settle();
      await ready();
    },
    /** Types a reference into the name box + Enter (Excel's Go To). */
    async goto(ref) {
      await p.click('.fv-namebox');
      await withMods(['Control'], () => p.keyboard.press('a'));
      await p.keyboard.type(ref);
      await p.keyboard.press('Enter');
      await view.settle();
      await ready();
    },
    async clickAt(xy, mods = [], clickOpts = {}) {
      if (!xy) throw new Error('click target not visible');
      await withMods(mods, () => p.mouse.click(xy.x, xy.y, clickOpts));
      await view.settle();
    },
    async cellXY(r, c) {
      await ready();
      const rect = await p.evaluate((r, c) => window.__g.cellRect(r, c), r, c);
      if (!rect) throw new Error(`cell ${colName(c)}${r + 1} not visible`);
      return { x: rect.cx, y: rect.cy };
    },
    async clickCell(r, c, mods = [], clickOpts = {}) {
      await g.clickAt(await g.cellXY(r, c), mods, clickOpts);
    },
    async clickHdr(kind, label, mods = []) {
      await ready();
      await g.clickAt(await p.evaluate((k, l) => window.__g.hdrXY(k, l), kind, label), mods);
    },
    nameBox: () => p.$eval('.fv-namebox', (e) => e.value),
    formula: () => p.$eval('.fv-formula', (e) => e.textContent),
    focus: async () => (await ready(), p.evaluate(() => window.__g.focus())),
    active: async () => (await ready(), p.evaluate(() => window.__g.active())),
    full: async (kind) => (await ready(), p.evaluate((k) => window.__g.full(k), kind)),
    tabs: async () => (await ready(), p.evaluate(() => window.__g.tabs())),
    activeTab: () => p.$eval('.fv-tab-active', (e) => e.textContent),
    async clickTab(name) {
      await p.evaluate((name) => [...document.querySelectorAll('.fv-tab')].find((t) => t.textContent === name).click(), name);
      await view.settle();
      await ready();
    },
    /** UI state after the debounced updates (ARIA announcement 150 ms, stats). */
    async ui() {
      await sleep(200);
      await view.settle();
      await ready();
      return p.evaluate(() => ({
        name: document.querySelector('.fv-namebox')?.value,
        formula: document.querySelector('.fv-formula')?.textContent,
        focus: window.__g.focus(),
        tab: document.querySelector('.fv-tab-active')?.textContent,
        zoom: document.querySelector('.fv-zoom-label')?.textContent,
        st: Math.round(document.querySelector('.fv-scroller')?.scrollTop ?? -1),
        active: window.__g.active(),
      }));
    },
    /** Status-bar stats text once the host answered. */
    async stats() {
      await sleep(250);
      await p.waitForFunction(() => !document.querySelector('.fv-status-stats').classList.contains('fv-pending'), { timeout: 8000, polling: 50 }).catch(() => {});
      return p.$$eval('.fv-status-stats .fv-stat', (els) => els.map((e) => e.textContent.replace(/\s+/g, ' ').trim()).join(' '));
    },
    /** Ctrl+C on the grid: { text, html } written by the copy event (or the async clipboard path). */
    async copy() {
      await p.evaluate(() => {
        window.__copied = {};
        const m = document.querySelector('.fv-status-msg');
        if (m) m.textContent = '';
      });
      await withMods(['Control'], () => p.keyboard.press('KeyC'));
      await p.waitForFunction(() => window.__copied['text/plain'] !== undefined || /^Copied/.test(document.querySelector('.fv-status-msg')?.textContent || ''), { timeout: 10000, polling: 50 }).catch(() => {});
      return p.evaluate(() => ({ text: window.__copied['text/plain'], html: window.__copied['text/html'] }));
    },
    /** The read-only cell text box (double-click / F2), or null. */
    cellBox: () =>
      p.evaluate(() => {
        const el = document.querySelector('.fv-cellbox');
        if (!el) return null;
        return { value: el.value, start: el.selectionStart, end: el.selectionEnd, focused: document.activeElement === el, label: el.getAttribute('aria-label'), ro: el.getAttribute('aria-readonly') };
      }),
  };
  return g;
}

module.exports = {
  ROOT, DIST, SAMPLES, BROWSER, SKIP, REQUIRE_BROWSER, sleep, colName,
  renderer, buildCsp, startHarness, suite, sheetHost, csvHost, markdownHost, docxHost, pdfHost, sample, grid,
};
