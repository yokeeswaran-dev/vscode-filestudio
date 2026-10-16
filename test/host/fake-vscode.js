// A small in-memory stand-in for the `vscode` module, so src/extension.ts and src/viewerProvider.ts (bundled to
// test/.out/host.js by test/build-tests.js) run in plain Node. It covers only what the host code uses; behaviour that
// matters to a test copies VS Code's (comments say which). Tests load it through test/host/load-host.js and drive it
// with the `__test` helpers (not part of the real API).
'use strict';
const fs = require('node:fs');
const nodePath = require('node:path');

const isWindows = process.platform === 'win32';

/** Every side effect worth asserting, as [name, ...args]. */
const calls = [];
const record = (name, ...args) => calls.push([name, ...args]);

// ===== URI (the parts of vscode-uri the host code uses) =====

const URI_RE = /^(([^:/?#]+?):)?(\/\/([^/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/;

function percentDecode(text) {
  return text.replace(/(%[0-9A-Fa-f]{2})+/g, (run) => {
    try {
      return decodeURIComponent(run);
    } catch {
      return run;
    }
  });
}

/** vscode-uri's encodeURIComponentFast: unreserved characters (and '/' in a path) kept, the rest %XX. */
function encodeFull(text, isPath) {
  let out = '';
  for (const ch of text) {
    if (/[A-Za-z0-9\-._~]/.test(ch) || (isPath && ch === '/')) out += ch;
    else out += encodeURIComponent(ch).replace(/[!'()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
  }
  return out;
}
const encodeMinimal = (text) => text.replace(/[#?]/g, (c) => (c === '#' ? '%23' : '%3F'));

class Uri {
  constructor(scheme, authority, path, query, fragment) {
    this.scheme = scheme;
    this.authority = authority;
    // vscode-uri _referenceResolution: http(s) and file paths are absolute
    if ((scheme === 'file' || scheme === 'http' || scheme === 'https') && path[0] !== '/') path = `/${path}`;
    this.path = path;
    this.query = query;
    this.fragment = fragment;
  }

  static parse(value, strict = false) {
    const m = URI_RE.exec(value);
    if (strict && !m[2]) throw new Error(`[UriError]: Scheme is missing: {scheme: "", authority: "", path: "${value}", query: "", fragment: ""}`);
    return new Uri(m[2] || 'file', percentDecode(m[4] ?? ''), percentDecode(m[5] ?? ''), percentDecode(m[7] ?? ''), percentDecode(m[9] ?? ''));
  }

  static file(path) {
    let authority = '';
    if (isWindows) path = path.replace(/\\/g, '/');
    if (path.startsWith('//')) {
      const end = path.indexOf('/', 2);
      authority = end < 0 ? path.slice(2) : path.slice(2, end);
      path = end < 0 ? '/' : path.slice(end) || '/';
    }
    return new Uri('file', authority, path, '', '');
  }

  static joinPath(base, ...segments) {
    return base.with({ path: nodePath.posix.join(base.path || '/', ...segments) });
  }

  with(change) {
    const pick = (key) => (change[key] === undefined ? this[key] : change[key] === null ? '' : change[key]);
    return new Uri(pick('scheme'), pick('authority'), pick('path'), pick('query'), pick('fragment'));
  }

  /** Like vscode-uri: `//host/path` for UNC, lower-case drive letter, backslashes on Windows. */
  get fsPath() {
    let value;
    if (this.authority && this.path.length > 1 && this.scheme === 'file') value = `//${this.authority}${this.path}`;
    else if (/^\/[A-Za-z]:/.test(this.path)) value = this.path[1].toLowerCase() + this.path.slice(2);
    else value = this.path;
    return isWindows ? value.replace(/\//g, '\\') : value;
  }

  toString(skipEncoding = false) {
    const encode = skipEncoding ? encodeMinimal : (text, isPath) => encodeFull(text, isPath);
    let out = this.scheme ? `${this.scheme}:` : '';
    if (this.authority || this.scheme === 'file') out += '//';
    if (this.authority) out += encode(this.authority.toLowerCase(), false);
    if (this.path) out += encode(this.path.replace(/^\/([A-Z]):/, (_, d) => `/${d.toLowerCase()}:`), true);
    if (this.query) out += `?${encode(this.query, false)}`;
    if (this.fragment) out += `#${skipEncoding ? this.fragment : encodeFull(this.fragment, false)}`;
    return out;
  }

  toJSON() {
    return { scheme: this.scheme, authority: this.authority, path: this.path, query: this.query, fragment: this.fragment };
  }
}

// ===== PRIMITIVES =====

class EventEmitter {
  constructor() {
    this.listeners = new Set();
    this.event = (fn) => {
      this.listeners.add(fn);
      return { dispose: () => this.listeners.delete(fn) };
    };
  }
  fire(e) {
    for (const fn of [...this.listeners]) fn(e);
  }
  dispose() {
    this.listeners.clear();
  }
}
class Position {
  constructor(line, character) {
    this.line = line;
    this.character = character;
  }
}
class Range {
  constructor(startLine, startCharacter, endLine, endCharacter) {
    this.start = new Position(startLine, startCharacter);
    this.end = new Position(endLine, endCharacter);
  }
}
class WorkspaceEdit {
  constructor() {
    this.edits = [];
  }
  replace(uri, range, text) {
    this.edits.push({ uri, range, text });
  }
}
class CancellationError extends Error {
  constructor() {
    super('Canceled');
    this.name = 'Canceled';
  }
}
class FileSystemError extends Error {
  constructor(message, code) {
    super(message);
    this.code = code;
  }
  static FileNotFound(uri) {
    return new FileSystemError(`Unable to read file '${uri}' (Error: Unable to resolve nonexistent file '${uri}')`, 'FileNotFound');
  }
}
const FileType = { Unknown: 0, File: 1, Directory: 2, SymbolicLink: 64 };
const ViewColumn = { Active: -1, Beside: -2, One: 1, Two: 2, Three: 3 };
class TabInputText {
  constructor(uri) {
    this.uri = uri;
  }
}
class TabInputTextDiff {
  constructor(original, modified) {
    this.original = original;
    this.modified = modified;
  }
}
class TabInputCustom {
  constructor(uri, viewType) {
    this.uri = uri;
    this.viewType = viewType;
  }
}
class TabInputNotebook {
  constructor(uri, notebookType) {
    this.uri = uri;
    this.notebookType = notebookType;
  }
}
class RelativePattern {
  constructor(base, pattern) {
    this.baseUri = typeof base === 'string' ? Uri.file(base) : base.uri ? base.uri : base;
    this.base = this.baseUri.fsPath;
    this.pattern = pattern;
  }
}

// ===== OUTPUT CHANNEL =====

/** Lines of the "FileStudio" log channel as "[level] message". */
const logLines = [];
function createLogChannel(name) {
  const channel = { name, dispose() {} };
  for (const level of ['trace', 'debug', 'info', 'warn', 'error']) channel[level] = (message) => logLines.push(`[${level}] ${message}`);
  return channel;
}

// ===== FILE SYSTEM (the real disk for file: URIs) =====

/**
 * VS Code with its defaults (security.restrictUNCAccess on, no allowed hosts) refuses a UNC host without touching the
 * network.
 */
function refuseUncHost(uri) {
  if (uri.scheme === 'file' && uri.authority) {
    throw new FileSystemError(
      `UNC host '${uri.authority}' access is not allowed. Please update the 'security.allowedUNCHosts' setting if you want to allow this host.`,
      'Unknown',
    );
  }
}
const fsApi = {
  // Like VS Code: false for a read-only provider (the Git extension's git:), undefined for an unknown scheme.
  isWritableFileSystem: (scheme) => (scheme === 'git' ? false : scheme === 'file' ? true : undefined),
  async readFile(uri) {
    refuseUncHost(uri);
    try {
      return new Uint8Array(fs.readFileSync(uri.fsPath));
    } catch {
      throw FileSystemError.FileNotFound(uri);
    }
  },
  async writeFile(uri, content) {
    fs.writeFileSync(uri.fsPath, content);
    record('fs.writeFile', uri.fsPath);
  },
  async stat(uri) {
    refuseUncHost(uri);
    let st;
    try {
      st = fs.statSync(uri.fsPath);
    } catch {
      throw FileSystemError.FileNotFound(uri);
    }
    return { type: st.isDirectory() ? FileType.Directory : FileType.File, size: st.size, ctime: st.ctimeMs, mtime: st.mtimeMs };
  },
  async copy(source, target) {
    fs.copyFileSync(source.fsPath, target.fsPath);
    record('fs.copy', source.fsPath, target.fsPath);
  },
  async delete(uri) {
    fs.rmSync(uri.fsPath, { force: true });
    record('fs.delete', uri.fsPath);
  },
};

// ===== FILE WATCHERS (events are fired by tests: __test.fireFileEvent) =====

const watchers = new Set();
function createFileSystemWatcher(pattern, ignoreCreate = false, ignoreChange = false, ignoreDelete = false) {
  if (!(pattern instanceof RelativePattern)) throw new Error('fake vscode: only RelativePattern watchers are supported');
  const emitters = { create: new EventEmitter(), change: new EventEmitter(), delete: new EventEmitter() };
  const ignored = { create: ignoreCreate, change: ignoreChange, delete: ignoreDelete };
  const watcher = {
    onDidCreate: emitters.create.event,
    onDidChange: emitters.change.event,
    onDidDelete: emitters.delete.event,
    dispose() {
      watchers.delete(watcher);
    },
    /** Non-recursive '*' or an exact name, case-insensitive (Windows). */
    fire(kind, uri) {
      const dir = nodePath.dirname(uri.fsPath).toLowerCase();
      const name = nodePath.basename(uri.fsPath).toLowerCase();
      const matches = dir === pattern.base.toLowerCase() && (pattern.pattern === '*' || pattern.pattern.toLowerCase() === name);
      if (matches && !ignored[kind]) emitters[kind].fire(uri);
    },
  };
  watchers.add(watcher);
  return watcher;
}

// ===== TEXT DOCUMENTS =====

const onDidChangeTextDocument = new EventEmitter();
class FakeTextDocument {
  constructor(uri, text) {
    this.uri = uri;
    this.text = text;
    this.version = 1;
  }
  getText() {
    return this.text;
  }
  get lineCount() {
    return this.text.split('\n').length;
  }
  lineAt(line) {
    return { text: this.text.split(/\r?\n/)[line] };
  }
  offsetAt(position) {
    const lines = this.text.split('\n');
    let offset = 0;
    for (let i = 0; i < position.line; i++) offset += lines[i].length + 1;
    return offset + position.character;
  }
  /** Replaces the text (an edit, or another program changing the file) and fires onDidChangeTextDocument. */
  setText(text) {
    this.text = text;
    this.version++;
    onDidChangeTextDocument.fire({ document: this, contentChanges: [{ text }], reason: undefined });
  }
}
/** Documents VS Code has synced to the extension host (uri.toString() -> FakeTextDocument). */
const textDocuments = new Map();

/** VS Code passes text documents to extensions only up to 50 MB, and never files it considers binary. */
const SYNC_LIMIT = 50 * 1024 * 1024;
async function openTextDocument(uri) {
  const known = textDocuments.get(uri.toString());
  if (known) return known;
  let size;
  try {
    size = fs.statSync(uri.fsPath).size;
  } catch {
    throw new Error(`cannot open ${uri.toString()}. Detail: Unable to read file '${uri.fsPath}'`);
  }
  if (size > SYNC_LIMIT) throw new Error(`cannot open ${uri.toString()}. Detail: Files above 50MB cannot be synchronized with extensions.`);
  const bytes = fs.readFileSync(uri.fsPath);
  const utf16Bom = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff);
  if (!utf16Bom && bytes.subarray(0, 512).includes(0)) {
    throw new Error(`cannot open ${uri.toString()}. Detail: File seems to be binary and cannot be opened as text`);
  }
  return new FakeTextDocument(uri, bytes.toString('utf8'));
}

// ===== CONFIGURATION =====

/** 'section.key' -> value (e.g. 'files.readonlyInclude'). */
const config = new Map();
const onDidChangeConfiguration = new EventEmitter();
function getConfiguration(section) {
  return { get: (key, fallback) => (config.has(`${section}.${key}`) ? config.get(`${section}.${key}`) : fallback) };
}
/** A string glob (absolute path) or RelativePattern (path under its base) against the document, case-insensitive. */
function globRegExp(glob) {
  const map = { '**/': '(?:.*[\\\\/])?', '**': '.*', '*': '[^\\\\/]*', '?': '[^\\\\/]', '/': '[\\\\/]' };
  const body = glob.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*\*\/?|\*|\?|\//g, (token) => map[token]);
  return new RegExp(`^${body}$`, 'i');
}
const languages = {
  match(selector, document) {
    const pattern = selector.pattern;
    if (pattern === undefined) return 0;
    const file = document.uri.fsPath;
    if (typeof pattern === 'string') return globRegExp(pattern).test(file) ? 5 : 0;
    const base = pattern.base.replace(/[\\/]+$/, '');
    if (!file.toLowerCase().startsWith(base.toLowerCase() + nodePath.sep)) return 0;
    return globRegExp(pattern.pattern).test(file.slice(base.length + 1)) ? 5 : 0;
  },
};

// ===== TABS =====

const tabGroup = { viewColumn: 1, isActive: true, tabs: [], get activeTab() { return this.tabs.find((t) => t.isActive); } };
const onDidChangeTabs = new EventEmitter();
const tabGroups = {
  all: [tabGroup],
  get activeTabGroup() {
    return tabGroup;
  },
  onDidChangeTabs: onDidChangeTabs.event,
  async close(tabs) {
    const list = Array.isArray(tabs) ? tabs : [tabs];
    record('tabGroups.close', list.map((t) => t.label));
    tabGroup.tabs = tabGroup.tabs.filter((t) => !list.includes(t));
    return true;
  },
};
/** Adds a tab to the (only) group: { input, label, isActive, isDirty }. */
function addTab(tab) {
  const full = { isActive: false, isDirty: false, isPinned: false, isPreview: false, ...tab, group: tabGroup };
  if (full.isActive) for (const other of tabGroup.tabs) other.isActive = false;
  tabGroup.tabs.push(full);
  return full;
}

// ===== WEBVIEW PANELS =====

function createPanel(viewColumn = 1) {
  const onMessage = new EventEmitter();
  const onDispose = new EventEmitter();
  const onViewState = new EventEmitter();
  const posted = [];
  const panel = {
    visible: true,
    viewColumn,
    posted,
    webview: {
      options: undefined,
      html: '',
      cspSource: 'https://*.vscode-cdn.net',
      asWebviewUri: (uri) => Uri.parse(`https://file+.vscode-resource.vscode-cdn.net${uri.path}`),
      postMessage: async (msg) => {
        posted.push(msg);
        return true;
      },
      onDidReceiveMessage: onMessage.event,
    },
    onDidDispose: onDispose.event,
    onDidChangeViewState: onViewState.event,
    /** Test helper: a message from the webview. */
    send: (msg) => onMessage.fire(msg),
    dispose: () => onDispose.fire(),
  };
  return panel;
}

// ===== NAMESPACES =====

const registrations = { providers: [], commands: new Map() };
/** Answer of the next showQuickPick: a function (items) -> item. */
const quickPick = { answer: (items) => items[0] };
const workspaceFolders = [];

const window = {
  createOutputChannel: (name, options) => createLogChannel(name, options),
  showInformationMessage: async (message) => record('info', message),
  showWarningMessage: async (message) => record('warning', message),
  showErrorMessage: async (message) => record('errorMessage', message),
  showQuickPick: async (items, options) => {
    record('quickPick', options?.placeHolder);
    return quickPick.answer(await items);
  },
  registerCustomEditorProvider: (viewType, provider, options) => {
    registrations.providers.push({ viewType, provider, options });
    return { dispose() {} };
  },
  tabGroups,
  activeTextEditor: undefined,
};

const workspace = {
  fs: fsApi,
  createFileSystemWatcher,
  getConfiguration,
  onDidChangeConfiguration: onDidChangeConfiguration.event,
  onDidChangeTextDocument: onDidChangeTextDocument.event,
  get textDocuments() {
    return [...textDocuments.values()];
  },
  openTextDocument,
  isTrusted: true,
  workspaceFolders,
  getWorkspaceFolder: (uri) => workspaceFolders.find((f) => uri.path.toLowerCase().startsWith(`${f.uri.path.toLowerCase()}/`)),
  applyEdit: async (edit) => {
    record('applyEdit', edit.edits.map((e) => [e.range.start.line, e.range.start.character, e.range.end.line, e.range.end.character, e.text]));
    for (const e of edit.edits) {
      const document = textDocuments.get(e.uri.toString());
      const text = document.getText();
      const from = document.offsetAt(e.range.start);
      const to = document.offsetAt(e.range.end);
      document.setText(text.slice(0, from) + e.text + text.slice(to));
    }
    return true;
  },
};

const commands = {
  executeCommand: async (id, ...args) => record('executeCommand', id, ...args),
  registerCommand: (id, fn) => {
    registrations.commands.set(id, fn);
    return { dispose() {} };
  },
};

// Records what VS Code hands to the browser: a string target as is, a Uri as encodeURI(uri.toString(true)).
const env = {
  openExternal: async (target) => {
    record('openExternal', typeof target === 'string' ? target : encodeURI(target.toString(true)));
    return true;
  },
};

const extensions = { getExtension: () => undefined };

/** Clears recorded calls, logs, tabs, documents, settings and registrations between tests. */
function reset() {
  calls.length = 0;
  logLines.length = 0;
  tabGroup.tabs = [];
  textDocuments.clear();
  config.clear();
  workspaceFolders.length = 0;
  registrations.providers.length = 0;
  registrations.commands.clear();
  quickPick.answer = (items) => items[0];
}

module.exports = {
  Uri,
  EventEmitter,
  Position,
  Range,
  WorkspaceEdit,
  CancellationError,
  FileSystemError,
  FileType,
  ViewColumn,
  TabInputText,
  TabInputTextDiff,
  TabInputCustom,
  TabInputNotebook,
  RelativePattern,
  window,
  workspace,
  commands,
  env,
  extensions,
  languages,
  __test: {
    calls,
    logLines,
    config,
    textDocuments,
    registrations,
    tabGroup,
    workspaceFolders,
    quickPick,
    FakeTextDocument,
    createPanel,
    addTab,
    reset,
    /** Delivers a file event ('create' | 'change' | 'delete') to every matching watcher. */
    fireFileEvent: (kind, uri) => {
      for (const watcher of [...watchers]) watcher.fire(kind, uri);
    },
    /** Fires onDidChangeConfiguration for these setting keys. */
    fireConfigChange: (keys) =>
      onDidChangeConfiguration.fire({ affectsConfiguration: (section) => keys.some((k) => k === section || k.startsWith(`${section}.`)) }),
    /** Fires tabGroups.onDidChangeTabs. */
    fireTabsChanged: (e) => onDidChangeTabs.fire({ opened: [], changed: [], closed: [], ...e }),
  },
};
