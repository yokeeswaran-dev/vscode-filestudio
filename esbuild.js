// Builds the three bundles:
//   dist/extension.js        - Node (extension host), CommonJS, `vscode` external
//   dist/viewer.js           - browser (webview), ESM with code splitting so heavy
//                              libraries (mermaid, pdf.js, ...) load lazily as dist/chunks/*.
//                              CSS imported from npm packages (KaTeX, the pdf.js viewer) is
//                              emitted as dist/viewer.css, its fonts and images as dist/fonts/*.
//   dist/pdf.worker.min.mjs  - the pdf.js worker; the PDF view fetches it and starts it as a
//                              blob: worker (a webview cannot start a worker from its resource URL).
// and copies the pdf.js data files the PDF view loads at run time into dist/pdfjs/.
// Usage: node esbuild.js [--watch] [--production]
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/** Prints errors in a format the $esbuild-watch problem matcher understands. */
const problemMatcherPlugin = {
  name: 'problem-matcher',
  setup(build) {
    build.onStart(() => console.log(`[watch] build started (${build.initialOptions.outfile || build.initialOptions.outdir})`));
    build.onEnd((result) => {
      for (const { text, location } of result.errors) {
        console.error(`✘ [ERROR] ${text}`);
        if (location) console.error(`    ${location.file}:${location.line}:${location.column}:`);
      }
      console.log('[watch] build finished');
    });
  },
};

// ----- excluded third-party modules -----

/**
 * Modules some dependencies load that FileStudio never uses, replaced by small stubs so their code (and licences) are
 * not shipped:
 *  - `unzipper` (with binary/buffers/chainsaw/bluebird): only ExcelJS's streaming WorkbookReader uses it; FileStudio
 *    loads workbooks with workbook.xlsx.load (JSZip). `buffers` has no licence at all.
 *  - `elkjs` (EPL-2.0): only Mermaid's optional "elk" layout uses it; such diagrams show Mermaid's error box instead.
 */
const EXCLUDED_MODULES = {
  unzipper:
    "const unavailable = () => { throw new Error('FileStudio does not include the streaming .xlsx reader (unzipper).'); };\n" +
    'module.exports = { Parse: unavailable, Open: { file: unavailable, buffer: unavailable } };',
  elkjs:
    'export default class ELK {\n' +
    "  constructor() { throw new Error('The ELK layout is not available in FileStudio; use the default Mermaid layout.'); }\n" +
    '}',
};

/**
 * Mermaid registers the ELK layout (elkjs) next to dagre and uses it by default. Without the registration Mermaid
 * falls back to dagre for every diagram, including ones that ask for `layout: elk` (getRegisteredLayoutAlgorithm),
 * and the ELK chunk is never bundled.
 */
const MERMAID_ELK_LOADERS = '...elkLayoutLoaders()';

const excludeModulesPlugin = {
  name: 'exclude-modules',
  setup(build) {
    let elkUnregistered = 0;
    build.onStart(() => {
      elkUnregistered = 0;
    });
    build.onLoad({ filter: /[\\/]mermaid[\\/]dist[\\/]chunks[\\/]mermaid\.core[\\/]chunk-[\w-]+\.mjs$/ }, async (args) => {
      const source = await fs.promises.readFile(args.path, 'utf8');
      if (!source.includes(MERMAID_ELK_LOADERS)) return undefined;
      elkUnregistered++;
      return { contents: source.replace(MERMAID_ELK_LOADERS, ''), loader: 'js', resolveDir: path.dirname(args.path) };
    });
    const isViewerBuild = JSON.stringify(build.initialOptions.entryPoints ?? '').includes('viewer.js');
    build.onEnd((result) => {
      if (isViewerBuild && !watch && !result.errors.length && elkUnregistered !== 1) {
        result.errors.push({ text: `mermaid: expected 1 ELK layout registration to remove, found ${elkUnregistered}. Update excludeModulesPlugin in esbuild.js.` });
      }
    });
    build.onResolve({ filter: /^(unzipper|elkjs)(\/.*)?$/ }, (args) => ({
      path: args.path.split('/')[0],
      namespace: 'excluded-module',
    }));
    build.onLoad({ filter: /.*/, namespace: 'excluded-module' }, (args) => ({
      contents: EXCLUDED_MODULES[args.path],
      loader: 'js',
    }));
  },
};

const common = {
  bundle: true,
  minify: production,
  sourcemap: !production,
  sourcesContent: false,
  logLevel: 'silent',
  plugins: [problemMatcherPlugin, excludeModulesPlugin],
};

// ----- pdf.js -----

const PDFJS_ROOT = path.dirname(require.resolve('pdfjs-dist/package.json'));

/**
 * pdf.js 6 decodes JBIG2, CCITT fax and JPEG 2000 images with WebAssembly, or else with asm.js modules that the
 * worker loads by `import(`${wasmUrl}<name>_nowasm_fallback.js`)`. Inside a VS Code webview the worker can do
 * neither: the CSP has no 'wasm-unsafe-eval', and the webview's service worker does not answer resource requests
 * made by a worker. So the two fallback modules are bundled into the worker and that import is pointed at them
 * (the PDF view passes `useWasm: false`). Without them scanned (CCITT / JBIG2) and JPEG 2000 images stay blank.
 */
const pdfWorkerDecodersPlugin = {
  name: 'pdf-worker-decoders',
  setup(build) {
    build.onLoad({ filter: /[\\/]pdfjs-dist[\\/]build[\\/]pdf\.worker\.min\.mjs$/ }, async (args) => {
      const source = await fs.promises.readFile(args.path, 'utf8');
      const importRe = /import\(\s*(?:\/\*[^*]*\*\/\s*)*`\$\{WasmImage\.#[\w$]+\}\$\{this\._noWasmFilename\}`\s*\)/g;
      const found = source.match(importRe)?.length ?? 0;
      if (found !== 1) {
        throw new Error(
          `pdf.worker.min.mjs: expected 1 dynamic import of the nowasm image decoders, found ${found}. ` +
            'pdfjs-dist changed: update pdfWorkerDecodersPlugin in esbuild.js.',
        );
      }
      const decoders =
        "{ 'jbig2_nowasm_fallback.js': fvJbig2Decoder, 'openjpeg_nowasm_fallback.js': fvOpenJpegDecoder }";
      return {
        contents:
          "import fvJbig2Decoder from 'pdfjs-dist/wasm/jbig2_nowasm_fallback.js';\n" +
          "import fvOpenJpegDecoder from 'pdfjs-dist/wasm/openjpeg_nowasm_fallback.js';\n" +
          source.replace(importRe, `Promise.resolve({ default: ${decoders}[this._noWasmFilename] })`),
        loader: 'js',
        resolveDir: path.dirname(args.path),
      };
    });
  },
};

/**
 * Data files the PDF view loads at run time, relative to dist/viewer.js (see the PDF section of media/viewer.js):
 * CMaps (text of PDFs with non-embedded CJK fonts), the standard fonts (PDFs that do not embed Helvetica, Times, ...),
 * the annotation icons (sticky notes without an appearance) and the licences of what is shipped.
 */
function copyPdfjsData() {
  const out = path.join(__dirname, 'dist', 'pdfjs');
  const copy = (from, to, filter) =>
    fs.cpSync(path.join(PDFJS_ROOT, from), path.join(out, to), { recursive: true, filter: (src) => !filter || fs.statSync(src).isDirectory() || filter(path.basename(src)) });
  copy('cmaps', 'cmaps');
  // Not the Liberation Sans fonts (GPL-2.0 with a font exception; FileStudio ships only MIT/Apache/BSD-style
  // licences). pdf.js then draws a non-embedded Helvetica/Arial with the system's sans-serif font.
  copy('standard_fonts', 'standard_fonts', (name) => !/^(LiberationSans-.*\.ttf|LICENSE_LIBERATION)$/.test(name));
  copy('web/images', 'images', (name) => /^annotation-.*\.svg$/.test(name));
  copy('LICENSE', 'LICENSE');
  // The JBIG2 / CCITT and OpenJPEG decoders bundled into dist/pdf.worker.min.mjs.
  copy('wasm', '.', (name) => /^LICENSE_(PDFJS_)?(JBIG2|OPENJPEG)$/.test(name));
}

async function main() {
  const extensionCtx = await esbuild.context({
    ...common,
    entryPoints: ['src/extension.ts'],
    outfile: 'dist/extension.js',
    platform: 'node',
    format: 'cjs',
    target: 'node18',
    external: ['vscode'],
  });

  const viewerCtx = await esbuild.context({
    ...common,
    entryPoints: { viewer: 'media/viewer.js' },
    outdir: 'dist',
    platform: 'browser',
    format: 'esm',
    splitting: true,
    target: 'es2022',
    entryNames: '[name]',
    chunkNames: 'chunks/[name]-[hash]',
    assetNames: 'fonts/[name]-[hash]',
    loader: {
      '.woff': 'file',
      '.woff2': 'file',
      '.ttf': 'file',
      '.eot': 'file',
      '.svg': 'file',
      '.png': 'file',
      '.gif': 'file',
    },
    define: { 'process.env.NODE_ENV': production ? '"production"' : '"development"' },
  });

  const pdfWorkerCtx = await esbuild.context({
    ...common,
    entryPoints: [path.join(PDFJS_ROOT, 'build', 'pdf.worker.min.mjs')],
    outfile: 'dist/pdf.worker.min.mjs',
    platform: 'browser',
    format: 'esm',
    target: 'es2022',
    // Started from a blob: URL, where a source map could not be found anyway.
    sourcemap: false,
    plugins: [...common.plugins, pdfWorkerDecodersPlugin],
  });

  copyPdfjsData();
  const contexts = [extensionCtx, viewerCtx, pdfWorkerCtx];
  if (watch) {
    await Promise.all(contexts.map((ctx) => ctx.watch()));
  } else {
    await Promise.all(contexts.map((ctx) => ctx.rebuild()));
    await Promise.all(contexts.map((ctx) => ctx.dispose()));
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
