// Bundles the pure renderer modules (src/renderers/*.ts, which never import `vscode`) for plain Node, so the unit
// tests in test/unit/ can require them: test/.out/<name>.js for sheet, markdown, docx, pdf and pptx.
// Usage: node test/build-tests.js   (run by `npm test` before `node --test test/unit/`)
const esbuild = require('esbuild');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const RENDERERS = ['sheet', 'markdown', 'docx', 'pdf', 'pptx'];

esbuild
  .build({
    absWorkingDir: ROOT,
    entryPoints: Object.fromEntries(RENDERERS.map((name) => [name, path.join(ROOT, 'src', 'renderers', `${name}.ts`)])),
    outdir: path.join(__dirname, '.out'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node18',
    logLevel: 'warning',
  })
  .catch(() => process.exit(1));
