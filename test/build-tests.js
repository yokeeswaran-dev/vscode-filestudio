// Bundles code for the plain-Node tests (run by `npm test` before `node --test`):
//   test/.out/<name>.js  - the pure renderer modules (src/renderers/*.ts, which never import `vscode`) for test/unit/:
//                          sheet, markdown, docx, pdf and pptx.
//   test/.out/viewer.js  - media/viewer.js for the webview unit tests (test/unit/regression-webview.test.js): the
//                          functions in VIEWER_EXPORTS are exported (viewer.js itself exports nothing), CSS imports are
//                          empty, DOMPurify is a stub (it needs a real DOM) and the lazily imported libraries (mermaid,
//                          pdf.js, ECharts) stay external. The test provides the few browser globals it needs at load.
//   test/.out/host.js    - src/extension.ts + src/viewerProvider.ts for test/host/ (`vscode` external: the tests map
//                          it to test/host/fake-vscode.js). ViewerSession is exported too (in memory only).
// Usage: node test/build-tests.js
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(__dirname, '.out');
const RENDERERS = ['sheet', 'markdown', 'docx', 'pdf', 'pptx'];

/** media/viewer.js functions the webview unit tests call. */
const VIEWER_EXPORTS = [
  'rawValue',
  'formulaBarNumber',
  'formatStat',
  'displayText',
  'planCopy',
  'buildCopyPayload',
  'buildMergeIndex',
  'pptxShapePathData',
  'pptxTextRect',
  'pptxCssGradient',
];

/** Appends the test exports (and a setter for the module's grid state `G`) to media/viewer.js; stubs DOMPurify. */
const viewerTestPlugin = {
  name: 'viewer-test-exports',
  setup(build) {
    build.onLoad({ filter: /[\\/]media[\\/]viewer\.js$/ }, async (args) => {
      const source = await fs.promises.readFile(args.path, 'utf8');
      const hooks = `\nexport { ${VIEWER_EXPORTS.join(', ')} };\n/** Test hook: sets the grid state (kind, meta, ...). */\nexport function __setGrid(grid) { G = grid; }\n`;
      return { contents: source + hooks, loader: 'js', resolveDir: path.dirname(args.path) };
    });
    // Lazily imported libraries stay external (their CSS imports, e.g. pdfjs-dist/web/pdf_viewer.css, are bundled empty).
    build.onResolve({ filter: /^(mermaid|echarts|pdfjs-dist|pdfjs-dist\/web\/pdf_viewer\.mjs)$/ }, (args) => ({ path: args.path, external: true }));
    build.onResolve({ filter: /^dompurify$/ }, () => ({ path: 'dompurify', namespace: 'viewer-test-stub' }));
    build.onLoad({ filter: /.*/, namespace: 'viewer-test-stub' }, () => ({
      contents:
        'export default { isSupported: false, addHook() {}, removeHook() {}, sanitize() { throw new Error("DOMPurify is not available in the unit-test bundle"); } };',
      loader: 'js',
    }));
  },
};

// Exports the module-private ViewerSession class in the host bundle (in memory; src/ is not changed), so host tests can
// install their own edit handlers the way the edit phases do.
const exportSession = {
  name: 'export-session',
  setup(build) {
    build.onLoad({ filter: /[\\/]viewerProvider\.ts$/ }, async (args) => {
      const src = await fs.promises.readFile(args.path, 'utf8');
      const out = src.replace('\nclass ViewerSession implements', '\nexport class ViewerSession implements');
      if (out === src) throw new Error('ViewerSession declaration not found');
      return { contents: out, loader: 'ts' };
    });
  },
};

// Counts the Markdown parses (in memory; src/ is not changed): every `MarkdownIt.prototype.parse` call pushes the length of
// its source to `globalThis.__fsMdParses`. Tests that guard the cost of a rendering (a document is parsed once, not
// three times) clear the array, run the code, and count the entries that are as long as the whole document.
const countMarkdownParses = {
  name: 'count-markdown-parses',
  setup(build) {
    build.onLoad({ filter: /[\\/]renderers[\\/]markdown\.ts$/ }, async (args) => {
      const src = await fs.promises.readFile(args.path, 'utf8');
      const hook =
        '\n;(() => {\n' +
        '  const g = globalThis as unknown as { __fsMdParses?: number[] };\n' +
        '  g.__fsMdParses ??= [];\n' +
        '  const proto = (MarkdownIt as unknown as { prototype: { parse: (s: string, e: unknown) => unknown; __fsCounted?: boolean } }).prototype;\n' +
        '  if (proto.__fsCounted) return;\n' +
        '  proto.__fsCounted = true;\n' +
        '  const original = proto.parse;\n' +
        '  proto.parse = function (this: unknown, s: string, e: unknown) {\n' +
        '    g.__fsMdParses!.push(String(s).length);\n' +
        '    return original.call(this, s, e);\n' +
        '  };\n' +
        '})();\n';
      return { contents: src + hook, loader: 'ts' };
    });
  },
};

const common = { absWorkingDir: ROOT, outdir: OUT, bundle: true, platform: 'node', format: 'cjs', target: 'node18', logLevel: 'warning' };

Promise.all([
  esbuild.build({
    ...common,
    entryPoints: Object.fromEntries(RENDERERS.map((name) => [name, path.join(ROOT, 'src', 'renderers', `${name}.ts`)])),
    plugins: [countMarkdownParses],
  }),
  esbuild.build({
    ...common,
    entryPoints: { viewer: path.join(ROOT, 'media', 'viewer.js') },
    loader: { '.css': 'empty' },
    // `import.meta.url` (Markdown resource origin) is empty in a CommonJS bundle; viewer.js handles that.
    logOverride: { 'empty-import-meta': 'silent' },
    plugins: [viewerTestPlugin],
  }),
  esbuild.build({
    ...common,
    outdir: undefined,
    outfile: path.join(OUT, 'host.js'),
    stdin: {
      contents: "export * from './src/viewerProvider';\nexport * as extension from './src/extension';\n",
      resolveDir: ROOT,
      sourcefile: 'host-entry.ts',
      loader: 'ts',
    },
    external: ['vscode'],
    plugins: [exportSession, countMarkdownParses],
  }),
]).catch(() => process.exit(1));
