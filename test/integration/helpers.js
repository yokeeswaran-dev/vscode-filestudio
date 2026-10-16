// Shared helpers for the integration tests. They run INSIDE the test VS Code (started by runTest.js), so they can use
// the `vscode` API and Node's `fs`. The workspace is a temporary copy of test/samples: tests may change files there,
// but should work on a copy (`copyFile`) when other tests need the original bytes.
'use strict';
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { pathToFileURL } = require('url');
const vscode = require('vscode');

const EXTENSION_ID = 'Yokeeswaran.filestudio';

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** Polls `check` (may be async) until it returns a truthy value, which is returned. Throws after `timeout` ms. */
async function waitFor(check, { timeout = 10_000, interval = 100, message = 'a condition' } = {}) {
  const end = Date.now() + timeout;
  for (;;) {
    const value = await check();
    if (value) return value;
    if (Date.now() >= end) throw new Error(`Timed out after ${timeout} ms waiting for ${message}`);
    await sleep(interval);
  }
}

// ----- workspace files -----

function workspaceUri() {
  const folder = vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders[0];
  if (!folder) throw new Error('The test VS Code has no workspace folder');
  return folder.uri;
}

/** A workspace file: a path relative to the workspace folder ('a.csv', 'images/x.png') or a Uri (returned as is). */
function fileUri(file) {
  return file instanceof vscode.Uri ? file : vscode.Uri.joinPath(workspaceUri(), ...file.split('/'));
}

function readBytes(file) {
  return fs.readFileSync(fileUri(file).fsPath);
}

function writeBytes(file, data) {
  const uri = fileUri(file);
  fs.writeFileSync(uri.fsPath, data);
  return uri;
}

/** Copies a workspace file (for example a sample) to a new name, so a test can change it; returns the new Uri. */
function copyFile(from, to) {
  const uri = fileUri(to);
  fs.copyFileSync(fileUri(from).fsPath, uri.fsPath);
  return uri;
}

function deleteFile(file) {
  fs.rmSync(fileUri(file).fsPath, { force: true });
}

// ----- FileStudio editors and tabs -----

function extension() {
  return vscode.extensions.getExtension(EXTENSION_ID);
}

/** The FileStudio custom editor viewType for a file name (from package.json `customEditors`), or undefined. */
function viewTypeFor(fileName) {
  const ext = extension();
  const editors = (ext && ext.packageJSON.contributes && ext.packageJSON.contributes.customEditors) || [];
  const name = path.basename(String(fileName)).toLowerCase();
  for (const editor of editors) {
    for (const selector of editor.selector || []) {
      const pattern = String(selector.filenamePattern || '').toLowerCase();
      if (pattern.startsWith('*.') && name.endsWith(pattern.slice(1))) return editor.viewType;
    }
  }
  return undefined;
}

/** Every viewType FileStudio registers, sorted. */
function allViewTypes() {
  const ext = extension();
  const editors = (ext && ext.packageJSON.contributes && ext.packageJSON.contributes.customEditors) || [];
  return editors.map((editor) => editor.viewType).sort();
}

/** Names of the top-level workspace files that FileStudio opens, sorted. */
function sampleFiles() {
  return fs
    .readdirSync(workspaceUri().fsPath, { withFileTypes: true })
    .filter((entry) => entry.isFile() && viewTypeFor(entry.name))
    .map((entry) => entry.name)
    .sort();
}

function allTabs() {
  return vscode.window.tabGroups.all.flatMap((group) => group.tabs);
}

/** The open tab of a file in any editor group; with `viewType`, only a FileStudio (custom editor) tab of that type. */
function findTab(file, viewType) {
  const key = fileUri(file).toString();
  return allTabs().find((tab) => {
    const input = tab.input;
    if (input instanceof vscode.TabInputCustom) return input.uri.toString() === key && (!viewType || input.viewType === viewType);
    if (input instanceof vscode.TabInputText) return !viewType && input.uri.toString() === key;
    return false;
  });
}

/** Short text for messages: 'custom:<viewType>', 'text', 'other:<class>' or 'none'. */
function describeTab(tab) {
  if (!tab) return 'none';
  const input = tab.input;
  if (input instanceof vscode.TabInputCustom) return 'custom:' + input.viewType;
  if (input instanceof vscode.TabInputText) return 'text';
  return 'other:' + (input && input.constructor ? input.constructor.name : String(input));
}

/**
 * Opens a workspace file with a FileStudio editor (default: the one package.json assigns to its extension) and waits
 * for its tab. `options` = a ViewColumn or { viewColumn, preserveFocus, preview } as for `vscode.openWith`.
 */
async function openWith(file, viewType = viewTypeFor(fileUri(file).path), options) {
  const uri = fileUri(file);
  if (!viewType) throw new Error(`FileStudio has no editor for ${uri.fsPath}`);
  await vscode.commands.executeCommand('vscode.openWith', uri, viewType, options);
  const tab = await waitFor(() => findTab(uri, viewType), {
    timeout: 15_000,
    message: `the ${viewType} tab of ${path.basename(uri.fsPath)}`,
  });
  return { uri, viewType, tab };
}

/** Makes this tab the active editor of its group. */
async function focusTab(tab) {
  const input = tab.input;
  if (input instanceof vscode.TabInputCustom) {
    await vscode.commands.executeCommand('vscode.openWith', input.uri, input.viewType, tab.group.viewColumn);
  } else if (input instanceof vscode.TabInputText) {
    await vscode.window.showTextDocument(input.uri, { viewColumn: tab.group.viewColumn, preview: false });
  }
}

/** True when the file's tab shows unsaved changes, or its text document (csv, Markdown) is dirty. */
function isDirty(file) {
  const key = fileUri(file).toString();
  const tabDirty = allTabs().some((tab) => tab.isDirty && tab.input && tab.input.uri && tab.input.uri.toString() === key);
  const doc = vscode.workspace.textDocuments.find((d) => d.uri.toString() === key);
  return tabDirty || Boolean(doc && doc.isDirty);
}

/**
 * Closes every editor. Dirty editors are REVERTED first (never saved): closing a dirty editor would open the modal
 * "Do you want to save?" dialog, which blocks the test run.
 */
async function closeAll() {
  for (let i = 0; i < 50; i++) {
    const dirty = allTabs().find((tab) => tab.isDirty);
    if (!dirty) break;
    await focusTab(dirty);
    await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
  }
  await vscode.commands.executeCommand('workbench.action.closeAllEditors');
  await waitFor(() => allTabs().length === 0, { timeout: 10_000, message: 'all editors to close' });
}

// Commands on the ACTIVE editor (open or focus it first). Edit tests: edit -> undo -> redo -> save / revert -> isDirty.
const save = () => vscode.commands.executeCommand('workbench.action.files.save');
const revert = () => vscode.commands.executeCommand('workbench.action.files.revert');
const undo = () => vscode.commands.executeCommand('undo');
const redo = () => vscode.commands.executeCommand('redo');

// ----- FileStudio output-channel log -----
// The "FileStudio" channel is a LogOutputChannel. VS Code writes it to a file in the test profile:
// <user-data-dir>/logs/<session>/window1/exthost/Yokeeswaran.filestudio/FileStudio.log, one entry per
// "<date> <time> [level] text" line (stack traces continue on the next lines). Activation failures go to exthost.log,
// and so do errors of FileStudio's host code that do not pass its log channel (a listener or timer callback that
// throws, an unhandled promise rejection): exthost.log names the extension or has a stack into its dist/ (or src/).

const LOG_LINE = /^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\.\d+ \[(\w+)\] (.*)$/;

function logFiles() {
  const userDataDir = process.env.FS_IT_USER_DATA_DIR;
  if (!userDataDir) throw new Error('FS_IT_USER_DATA_DIR is not set: run the tests with `npm run test:integration`');
  const files = [];
  const walk = (dir) => {
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(p);
      else if (/filestudio.*\.log$/i.test(entry.name) || entry.name === 'exthost.log') files.push(p);
    }
  };
  walk(path.join(userDataDir, 'logs'));
  return files;
}

/** The current end of every log file. Pass it to `logSince` to get only the entries written after this point. */
function logMark() {
  const mark = new Map();
  for (const file of logFiles()) mark.set(file, fs.statSync(file).size);
  return mark;
}

/**
 * Lower-case forms of FileStudio's code folders (dist/ and src/ of the extension) as they appear in stack traces:
 * 'c:\x\dist', 'c:/x/dist', 'file:///c:/x/dist' and the encoded 'file:///c%3a/x/dist'.
 */
function extensionCodePaths() {
  const ext = extension();
  if (!ext) return [];
  const forms = new Set();
  for (const dir of ['dist', 'src']) {
    const p = path.join(ext.extensionPath, dir);
    const variants = [p.replace(/\//g, '\\'), p.replace(/\\/g, '/'), pathToFileURL(p).href, vscode.Uri.file(p).toString()];
    for (const form of variants) forms.add(form.toLowerCase());
  }
  return [...forms];
}

/**
 * Log entries written after `mark` (default: all): { log: 'FileStudio' | 'exthost', level, text }.
 * exthost.log entries are included when they name the extension (for example a failed activation), and errors when
 * their text (with the stack) points into FileStudio's code (dist/ or src/ of the extension folder).
 */
function logSince(mark = new Map()) {
  const entries = [];
  for (const file of logFiles()) {
    const log = path.basename(file) === 'exthost.log' ? 'exthost' : 'FileStudio';
    const text = fs.readFileSync(file).subarray(mark.get(file) || 0).toString('utf8');
    let current;
    for (const line of text.split(/\r?\n/)) {
      const m = LOG_LINE.exec(line);
      if (m) {
        current = { log, level: m[1], text: m[2] };
        entries.push(current);
      } else if (current && line) {
        current.text += '\n' + line;
      }
    }
  }
  const id = EXTENSION_ID.toLowerCase();
  const codePaths = extensionCodePaths();
  return entries.filter((e) => {
    if (e.log === 'FileStudio') return true;
    const text = e.text.toLowerCase();
    return text.includes(id) || (e.level === 'error' && codePaths.some((p) => text.includes(p)));
  });
}

function logSize() {
  let size = 0;
  for (const file of logFiles()) size += fs.statSync(file).size;
  return size;
}

/** Waits at least `min` ms, then until the log files have not grown for `quiet` ms (at most `max` ms in total). */
async function waitForLogQuiet({ min = 1000, quiet = 500, max = 10_000 } = {}) {
  const start = Date.now();
  await sleep(min);
  let size = logSize();
  let since = Date.now();
  while (Date.now() - since < quiet && Date.now() - start < max) {
    await sleep(100);
    const now = logSize();
    if (now !== size) {
      size = now;
      since = Date.now();
    }
  }
}

/** Fails, listing the entries, when `entries` contain an error, or a warning in the FileStudio channel. */
function assertNoLogProblems(entries, what) {
  const problems = entries
    .filter((e) => e.level === 'error' || (e.log === 'FileStudio' && e.level === 'warning'))
    .map((e) => `[${e.log} ${e.level}] ${e.text.split('\n').slice(0, 4).join(' | ')}`);
  assert.deepEqual(problems, [], `${what}: errors or warnings were logged`);
}

module.exports = {
  EXTENSION_ID,
  sleep,
  waitFor,
  workspaceUri,
  fileUri,
  readBytes,
  writeBytes,
  copyFile,
  deleteFile,
  extension,
  viewTypeFor,
  allViewTypes,
  sampleFiles,
  allTabs,
  findTab,
  describeTab,
  openWith,
  focusTab,
  isDirty,
  closeAll,
  save,
  revert,
  undo,
  redo,
  logFiles,
  logMark,
  logSince,
  extensionCodePaths,
  waitForLogQuiet,
  assertNoLogProblems,
};
