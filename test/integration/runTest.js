// Integration tests in a REAL VS Code: starts VS Code with FileStudio loaded from this repo (`dist/`) and runs the tests
// in test/integration/*.test.js inside it with a small runner (see suite/index.js). Build first: `npm run build` (the
// runner stops when dist/ is older than src/ or media/, so the tests never pass against an old build).
//
// Usage: npm run test:integration          (Linux without a display: xvfb-run -a npm run test:integration)
//
// Environment variables (all optional):
//   FS_VSCODE_EXE      path of an installed VS Code executable to use (Code.exe / code). Without it, VS Code is downloaded
//                      once into .vscode-test/ (git-ignored) and reused.
//   FS_VSCODE_VERSION  version to download when FS_VSCODE_EXE is not set: 'stable' (default), 'insiders' or e.g. '1.123.0'.
//   FS_IT_TMPDIR       folder for the run folder (profile, logs, workspace copy). Default: the OS temp folder.
//   FS_IT_KEEP=1       keep the run folder after a passing run (a failing run always keeps it, for its logs).
//   FS_IT_GREP         run only the tests whose full title contains this text.
//
// Every run is isolated: its own --user-data-dir and --extensions-dir, all other extensions disabled, and the workspace
// is a fresh temporary copy of test/samples (without large.xlsx). Tests never write into test/samples.
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const { runTests } = require('@vscode/test-electron');

const ROOT = path.resolve(__dirname, '..', '..');
const SAMPLES = path.join(ROOT, 'test', 'samples');
/** Top-level samples not copied into the workspace: large.xlsx (44 MB, not in Git) takes about a minute to open. */
const SKIPPED_SAMPLES = new Set(['large.xlsx']);

/** User settings of the test profile: no welcome page, no update checks, no Git scanning, no AI features. */
const SETTINGS = {
  'workbench.startupEditor': 'none',
  'security.workspace.trust.enabled': false,
  'update.mode': 'none',
  'telemetry.telemetryLevel': 'off',
  'extensions.autoUpdate': false,
  'extensions.autoCheckUpdates': false,
  'extensions.ignoreRecommendations': true,
  'workbench.enableExperiments': false,
  'workbench.tips.enabled': false,
  'workbench.editor.enablePreview': false,
  'window.restoreWindows': 'none',
  'chat.disableAIFeatures': true,
  'git.enabled': false,
};

/** The newest file under `dir` (recursive): { file, mtimeMs }, or undefined when there is none. */
function newestFile(dir) {
  let newest;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    let found;
    if (entry.isDirectory()) found = newestFile(p);
    else if (entry.isFile()) found = { file: p, mtimeMs: fs.statSync(p).mtimeMs };
    if (found && (!newest || found.mtimeMs > newest.mtimeMs)) newest = found;
  }
  return newest;
}

/** An error text when dist/ is missing, or older than a file in src/ or media/ (changed after the last build). */
function staleBuild() {
  const built = ['extension.js', 'viewer.js'].map((name) => path.join(ROOT, 'dist', name));
  const missing = built.find((file) => !fs.existsSync(file));
  if (missing) return `${path.relative(ROOT, missing)} is missing. Run \`npm run build\` first.`;
  const buildTime = Math.min(...built.map((file) => fs.statSync(file).mtimeMs));
  for (const dir of ['src', 'media']) {
    const newest = newestFile(path.join(ROOT, dir));
    if (newest && newest.mtimeMs > buildTime) {
      return `${path.relative(ROOT, newest.file)} changed after the last build (dist/ is older). Run \`npm run build\` first.`;
    }
  }
  return undefined;
}

async function main() {
  const stale = staleBuild();
  if (stale) {
    console.error(`[integration] ${stale}`);
    return 1;
  }

  const base = process.env.FS_IT_TMPDIR ? path.resolve(process.env.FS_IT_TMPDIR) : os.tmpdir();
  fs.mkdirSync(base, { recursive: true });
  const runDir = fs.mkdtempSync(path.join(base, 'filestudio-it-'));
  const userDataDir = path.join(runDir, 'user-data');
  const extensionsDir = path.join(runDir, 'extensions');
  const workspaceDir = path.join(runDir, 'workspace');
  fs.mkdirSync(path.join(userDataDir, 'User'), { recursive: true });
  fs.mkdirSync(extensionsDir);
  fs.writeFileSync(path.join(userDataDir, 'User', 'settings.json'), JSON.stringify(SETTINGS, null, 2) + '\n');
  fs.cpSync(SAMPLES, workspaceDir, {
    recursive: true,
    filter: (src) => !(path.dirname(src) === SAMPLES && SKIPPED_SAMPLES.has(path.basename(src))),
  });

  // Variables inherited from a VS Code terminal (ELECTRON_RUN_AS_NODE, VSCODE_*) would make the test instance start as
  // Node or talk to the VS Code that owns the terminal. The test instance gets a clean environment.
  for (const key of Object.keys(process.env)) {
    if (/^(ELECTRON_RUN_AS_NODE|VSCODE_)/i.test(key)) delete process.env[key];
  }

  const options = {
    extensionDevelopmentPath: ROOT,
    extensionTestsPath: path.join(__dirname, 'suite', 'index.js'),
    cachePath: path.join(ROOT, '.vscode-test'),
    // runTests adds --disable-workspace-trust, --skip-welcome, --skip-release-notes, --disable-updates, --no-sandbox.
    launchArgs: [
      workspaceDir,
      `--user-data-dir=${userDataDir}`,
      `--extensions-dir=${extensionsDir}`,
      '--disable-extensions', // all extensions except the one under test
    ],
    extensionTestsEnv: {
      FS_IT_USER_DATA_DIR: userDataDir, // the tests read the FileStudio output-channel log from here
      FS_IT_GREP: process.env.FS_IT_GREP || '',
    },
  };
  if (process.env.FS_VSCODE_EXE) options.vscodeExecutablePath = path.resolve(process.env.FS_VSCODE_EXE);
  else options.version = process.env.FS_VSCODE_VERSION || 'stable';

  console.log(`[integration] VS Code: ${options.vscodeExecutablePath || options.version}; run folder: ${runDir}`);
  const started = Date.now();
  let code = 1;
  try {
    code = await runTests(options);
  } catch (err) {
    console.error(`[integration] ${(err && err.message) || err}`);
    code = 1;
  }
  console.log(`[integration] finished in ${Math.round((Date.now() - started) / 1000)} s with exit code ${code}`);
  if (code === 0 && process.env.FS_IT_KEEP !== '1') {
    fs.rmSync(runDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
  } else {
    console.log(`[integration] kept the run folder (profile, logs, workspace): ${runDir}`);
  }
  return code;
}

main().then(
  (code) => process.exit(code),
  (err) => {
    console.error(err);
    process.exit(1);
  },
);
