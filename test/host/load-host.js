// Loads the extension host code (test/.out/host.js, built by test/build-tests.js) with `require('vscode')` mapped to
// test/host/fake-vscode.js, plus small helpers for the host tests.
'use strict';
const Module = require('node:module');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const FAKE = path.join(__dirname, 'fake-vscode.js');
const resolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  return request === 'vscode' ? FAKE : resolve.call(this, request, ...rest);
};

const vscode = require('vscode');
const host = require('../.out/host.js');

/** A temp folder for this test file, removed when the process exits. */
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'fs-host-'));
process.on('exit', () => fs.rmSync(tmp, { recursive: true, force: true }));

module.exports = {
  vscode,
  T: vscode.__test,
  host,
  tmp,
  extensionUri: vscode.Uri.file(path.join(tmp, 'extension')),
  token: { isCancellationRequested: false, onCancellationRequested: () => ({ dispose() {} }) },
  tick: (ms = 0) => new Promise((resolveTick) => setTimeout(resolveTick, ms)),
  /** Writes `content` to `name` in the temp folder; returns its Uri. */
  writeTmp(name, content) {
    const file = path.join(tmp, name);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, content);
    return vscode.Uri.file(file);
  },
  /** The last message of `type` a panel received. */
  lastOf: (panel, type) => [...panel.posted].reverse().find((m) => m.type === type),
  /** Recorded calls named `name` (see fake-vscode.js). */
  callsOf: (name) => vscode.__test.calls.filter((c) => c[0] === name),
  /** Polls `check` until it returns a truthy value (or fails after `ms`). */
  async waitFor(check, ms = 3000, what = 'condition') {
    const end = Date.now() + ms;
    for (;;) {
      const value = await check();
      if (value) return value;
      if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
      await new Promise((r) => setTimeout(r, 10));
    }
  },
};
