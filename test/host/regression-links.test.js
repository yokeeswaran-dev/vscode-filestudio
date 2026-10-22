// Regression tests for fixed 0.1.x extension-host bugs in links, the webview page (CSP, loading hint), error logging and
// file-type routing. The host code runs against test/host/fake-vscode.js. Each test names the bug id from the 0.1.x bug
// hunts. Run with `npm test`.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { vscode, T, host, tmp, extensionUri, token, writeTmp, lastOf, callsOf, waitFor } = require('./load-host');

beforeEach(() => T.reset());

const isWindows = process.platform === 'win32';
const docUri = () => writeTmp('doc.md', '# Doc\n\n## Usage\n');
const follow = (href, documentUri = docUri()) => host.openLink(href, { documentUri, allowLocal: true });
/** The vscode.open calls as [fsPath, selection as [line, character] or undefined]. */
const opened = () =>
  callsOf('executeCommand')
    .filter((c) => c[1] === 'vscode.open')
    .map(([, , uri, options]) => [uri.fsPath, options?.selection ? [options.selection.start.line, options.selection.start.character] : undefined]);
const warnings = () => callsOf('warning').map((c) => c[1]);

// ===== EXTERNAL LINKS =====

test('H3-F1: web links open exactly as written (escapes such as %23 / %26 / %2B are not decoded or doubled)', async () => {
  await follow('https://example.com/search?q=C%23&a=a%26b&lang=C%2B%2B#frag%20x');
  await follow('https://example.com/a b?x=é');
  assert.deepEqual(callsOf('openExternal').map((c) => c[1]), [
    'https://example.com/search?q=C%23&a=a%26b&lang=C%2B%2B#frag%20x',
    'https://example.com/a%20b?x=%C3%A9',
  ]);
});

// ===== LOCAL LINKS =====

const OTHER = ['# Other', '', '## Install steps', '', 'text', '', '## L2', '', 'a', 'b', 'c'].join('\n');

test('H3-F7: a link to a heading of another Markdown file opens it at that heading; a link into the same file scrolls', async () => {
  const other = writeTmp('other.md', OTHER);
  await follow('other.md#install-steps');
  await follow('other.md#Install-Steps');
  await follow('other.md#user-content-install-steps');
  assert.deepEqual(opened(), [
    [other.fsPath, [2, 0]],
    [other.fsPath, [2, 0]],
    [other.fsPath, [2, 0]],
  ]);
  // In a Markdown view, a link to the document itself scrolls the preview instead of opening the file again.
  const uri = docUri();
  const document = new T.FakeTextDocument(uri, '# Doc\n\n## Usage\n');
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
  panel.send({ type: 'ready' });
  await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  panel.send({ type: 'openLink', href: 'doc.md#usage' });
  assert.deepEqual(await waitFor(() => lastOf(panel, 'scrollToFragment'), 3000, 'scrollToFragment'), { type: 'scrollToFragment', fragment: 'usage' });
  assert.equal(opened().length, 3, 'not opened again');
});

test('R2H2-HFRAG: "#l2" is the heading "L2" of a Markdown file; "#L5,3", "#10" and "#L9-L10" are VS Code line links', async () => {
  const other = writeTmp('other.md', OTHER);
  await follow('other.md#l2');
  await follow('other.md#L5,3');
  await follow('other.md#10');
  await follow('other.md#L9-L10');
  assert.deepEqual(opened(), [
    [other.fsPath, [6, 0]],
    [other.fsPath, [4, 2]],
    [other.fsPath, [9, 0]],
    [other.fsPath, [8, 0]],
  ]);
  const range = callsOf('executeCommand').filter((c) => c[1] === 'vscode.open')[3][3].selection;
  assert.deepEqual([range.end.line, range.end.character], [10, 0], '#L9-L10 selects both lines');
});

test('H3-F8: a Windows absolute path with encoded backslashes (C:%5Cdir%5Cfile.md) is a file path, not a "c:" link', async () => {
  await follow('C:%5Cfs-no-such-dir%5Cother.md');
  assert.equal(warnings().length, 1);
  assert.doesNotMatch(warnings()[0], /links are not allowed/);
  assert.match(warnings()[0], /^FileStudio could not find "c:[\\/]fs-no-such-dir[\\/]other\.md"\.$/);
  if (isWindows) {
    T.reset();
    const target = writeTmp('win path.md', '# x\n');
    await follow(target.fsPath.replace(/\\/g, '%5C').replace(/ /g, '%20'));
    assert.deepEqual(opened(), [[target.fsPath, undefined]]);
    assert.deepEqual(warnings(), []);
  }
});

test('R2H3-C: a refused UNC host shows VS Code\'s reason, not "could not find"', async () => {
  await follow('\\\\server\\share\\book.xlsx');
  assert.equal(warnings().length, 1);
  assert.match(warnings()[0], /^FileStudio could not open ".*server.share.book\.xlsx": UNC host 'server' access is not allowed/);
  assert.doesNotMatch(warnings()[0], /could not find/);
  assert.deepEqual(opened(), []);
});

test('H7-F8 (host side): links in a Word document to local files (file:///C:\\...) and relative paths open in VS Code', async () => {
  const docx = writeTmp('linked.docx', fs.readFileSync(path.join(__dirname, '..', 'samples', 'sample.docx')));
  const target = writeTmp('target.docx', 'x');
  const panel = T.createPanel();
  const provider = new host.DocxViewerProvider(extensionUri);
  provider.resolveCustomEditor(provider.openCustomDocument(docx, {}), panel, token);
  panel.send({ type: 'ready' });
  await waitFor(() => lastOf(panel, 'init'), 5000, 'docx init');
  // Word writes local hyperlinks as file:///C:\dir\Other.docx (Windows separators).
  const fileHref = isWindows ? `file:///${target.fsPath}` : `file://${target.fsPath}`;
  panel.send({ type: 'openLink', href: fileHref });
  panel.send({ type: 'openLink', href: 'target.docx' });
  panel.send({ type: 'openLink', href: 'command:workbench.action.reloadWindow' });
  await waitFor(() => opened().length === 2 && warnings().length === 1, 3000, 'links handled');
  assert.deepEqual(opened(), [
    [target.fsPath, undefined],
    [target.fsPath, undefined],
  ]);
  assert.deepEqual(warnings(), ['FileStudio did not open the link: "command:" links are not allowed.']);
});

// ===== WEBVIEW PAGE =====

test('H8-F5: the page CSP refuses forms, <base> and frames; remote images only in the Markdown preview', () => {
  for (const kind of ['sheet', 'csv', 'markdown', 'docx', 'pdf', 'pptx']) {
    const csp = host.buildCsp('https://*.vscode-cdn.net', 'NONCE', kind);
    for (const directive of ["form-action 'none'", "base-uri 'none'", "frame-src 'none'", "child-src 'none'", "default-src 'none'"]) {
      assert.ok(csp.includes(directive), `${kind}: ${directive}`);
    }
    assert.equal(/\bhttps:(?!\/\/)/.test(csp), kind === 'markdown', `${kind}: remote https: sources`);
    assert.doesNotMatch(csp, /unsafe-eval/);
  }
});

test('H7-F11: the "still loading" hint stays hidden (also from screen readers) until it is revealed after 10 s', () => {
  const html = host.getHtml(T.createPanel().webview, extensionUri, 'docx', 'a.docx');
  assert.match(html, /\.fv-boot-hint \{ visibility: hidden; animation: fv-boot-reveal 0s linear 10s forwards; \}/);
  assert.match(html, /@keyframes fv-boot-reveal \{ to \{ visibility: visible; \} \}/);
  assert.doesNotMatch(html, /\.fv-boot-hint \{[^}]*(opacity|display)/, 'hidden by visibility, not just transparent');
  assert.match(html, /<p class="fv-boot-hint">Still loading\./);
});

// ===== ERRORS AND FILE TYPES =====

test('R2H2-C1: a corrupt workbook is logged once per load (not again per ready / row request), once more per reload', async () => {
  const uri = writeTmp('corrupt.xlsx', Buffer.from('PK\u0003\u0004 not really a zip'));
  const errors = () => T.logLines.filter((line) => line.startsWith('[error]'));
  const provider = new host.SheetEditorProvider(extensionUri);
  const document = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(document, panel, token);
  panel.send({ type: 'ready' });
  await waitFor(() => lastOf(panel, 'error'), 5000, 'error view');
  panel.send({ type: 'ready' }); // webview reloaded
  panel.send({ type: 'getRows', reqId: 1, sheet: 0, start: 0, end: 10 });
  await waitFor(() => panel.posted.filter((m) => m.type === 'error').length >= 2, 3000, 'second error view');
  assert.equal(errors().length, 1, errors().join('\n'));
  assert.match(errors()[0], /\[sheet\] Could not load .*corrupt\.xlsx/);
  // Changed on disk, still corrupt: one more entry for the new load.
  fs.writeFileSync(uri.fsPath, Buffer.from('PK\u0003\u0004 still not a zip, longer'));
  T.fireFileEvent('change', uri);
  await waitFor(() => errors().length === 2, 5000, 'second load logged');
  panel.send({ type: 'ready' });
  await waitFor(() => panel.posted.filter((m) => m.type === 'error').length >= 4, 3000, 'error view after the reload');
  assert.equal(errors().length, 2);
  document.dispose();
});

test('R3-033 (host side): pptm / ppsx / ppsm / potx / potm files open in the presentation viewer', async () => {
  for (const ext of ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'POTM']) {
    assert.equal(host.viewTypeForUri(vscode.Uri.file(path.join(tmp, `deck.${ext}`))), 'fileStudio.pptx', ext);
  }
  assert.equal(host.viewTypeForUri(vscode.Uri.file(path.join(tmp, 'deck.ppt'))), undefined);
  host.extension.activate({ extensionUri, subscriptions: [] });
  await T.registrations.commands.get('fileStudio.openWith')(vscode.Uri.file(path.join(tmp, 'Slides.potm')));
  const call = callsOf('executeCommand').find((c) => c[1] === 'vscode.openWith');
  assert.equal(call[3], 'fileStudio.pptx');
  assert.deepEqual(warnings(), []);
});
