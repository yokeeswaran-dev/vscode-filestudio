// Regression tests for fixed 0.1.x extension-host bugs on the paths 0.2.0 editing builds on: reload after a change on
// disk, Change Delimiter, Markdown task toggles (the model for every text edit), read-only documents, files VS Code
// cannot pass to a custom text editor, and Open with on a text tab. The host code runs against test/host/fake-vscode.js.
// Each test names the bug id from the 0.1.x bug hunts. Run with `npm test`.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const { vscode, T, host, tmp, extensionUri, token, tick, writeTmp, lastOf, callsOf, waitFor } = require('./load-host');

const pkg = require('../../package.json');

beforeEach(() => T.reset());

/** .xlsx bytes with A1 = `text`; `chart` adds a chart part (a feature the viewer does not show). */
async function xlsxBytes(text, chart = false) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = text;
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  if (chart) zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** A text document VS Code has synced (FakeTextDocument), shown in a FileStudio csv / markdown panel. */
function openText(name, text, kind) {
  const uri = writeTmp(name, text);
  const document = new T.FakeTextDocument(uri, text);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, kind).resolveCustomTextEditor(document, panel, token);
  return { uri, document, panel };
}

// ===== SHEET: RELOAD AFTER A CHANGE ON DISK =====

test('H3-F6 / R2H2-BANNER: a workbook changed on disk is reloaded; the "not displayed" banner follows the new file', async () => {
  const uri = writeTmp('reload.xlsx', await xlsxBytes('first'));
  const provider = new host.SheetEditorProvider(extensionUri);
  const document = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'init');
  assert.equal(init.banner, undefined, 'no banner for a plain workbook');

  // Another program adds a chart: the grid is refreshed in place with the new banner.
  fs.writeFileSync(uri.fsPath, await xlsxBytes('second, with a chart', true));
  T.fireFileEvent('change', uri);
  const withChart = await waitFor(() => lastOf(panel, 'invalidate'), 5000, 'invalidate after the change');
  assert.match(withChart.banner, /^This workbook contains content that FileStudio does not display: Charts/);
  panel.send({ type: 'getRows', reqId: 1, sheet: 0, start: 0, end: 10 });
  const rows = await waitFor(() => lastOf(panel, 'rows'), 5000, 'rows');
  assert.equal(rows.rows[0].cells[0].v, 'second, with a chart', 'rows come from the reloaded file');

  // The chart is removed again: the banner is cleared (null, not left as it was).
  panel.posted.length = 0;
  fs.writeFileSync(uri.fsPath, await xlsxBytes('third'));
  T.fireFileEvent('change', uri);
  const cleared = await waitFor(() => lastOf(panel, 'invalidate'), 5000, 'invalidate after the second change');
  assert.equal(cleared.banner, null);
  assert.equal(lastOf(panel, 'init'), undefined, 'refreshed in place, not re-initialised');
  document.dispose();
});

// ===== CSV: CHANGE DELIMITER =====

test('R3-055: Change Delimiter re-parses the view with the chosen delimiter and leaves the document text alone', async () => {
  const text = 'name,note\nx,a|b\ny,c|d\n';
  const { uri, document, panel } = openText('delim.csv', text, 'csv');
  panel.send({ type: 'ready' });
  const first = await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  assert.deepEqual(first.delimited, { format: 'csv', delimiter: ',' }, 'detected');

  panel.posted.length = 0;
  T.quickPick.answer = (items) => items.find((item) => item.delimiter === '|');
  await host.changeDelimiter(uri);
  const second = await waitFor(() => lastOf(panel, 'init'), 3000, 'init after Change Delimiter');
  assert.deepEqual(second.delimited, { format: 'csv', delimiter: '|' }, 'the status bar shows the choice');
  panel.send({ type: 'getRows', reqId: 2, sheet: 0, start: 0, end: 5 });
  const rows = await waitFor(() => lastOf(panel, 'rows'), 3000, 'rows');
  assert.deepEqual(
    rows.rows.map((row) => row.cells.map((c) => c.v)),
    [['name,note'], ['x,a', 'b'], ['y,c', 'd']],
  );
  assert.equal(document.getText(), text, 'the file text is unchanged');
  assert.equal(document.version, 1);
  assert.deepEqual(callsOf('applyEdit'), [], 'no edit');
  T.quickPick.answer = (items) => items[0];
  await host.changeDelimiter(uri); // back to "Detect automatically" (session state is shared by the whole file)
});

// ===== MARKDOWN TASK TOGGLES =====

/** Opens markdown `text`, waits for init; returns helpers to toggle a task and read the answer. */
async function openMarkdown(name, text) {
  const opened = openText(name, text, 'markdown');
  opened.panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(opened.panel, 'init'), 3000, 'markdown init');
  const toggle = async (line, checked, version) => {
    opened.panel.posted.length = 0;
    opened.panel.send({ type: 'toggleTask', line, checked, version });
    return waitFor(() => lastOf(opened.panel, 'toggleTaskResult'), 3000, 'toggleTaskResult');
  };
  return { ...opened, init, toggle };
}

test('H3-F5 / H6-F7: a task after a nested list marker on the same line can be toggled; a stale click is refused', async () => {
  const { document, toggle } = await openMarkdown('nested.md', '- - [ ] a\n- [ ] b\n');
  const done = await toggle(0, true, 1);
  assert.equal(done.applied, true);
  assert.equal(document.getText(), '- - [x] a\n- [ ] b\n', 'only the nested marker changed');
  assert.deepEqual(callsOf('applyEdit'), [[ 'applyEdit', [[0, 5, 0, 6, 'x']] ]]);
  const stale = await toggle(1, true, 1);
  assert.deepEqual([stale.applied, stale.reason, stale.documentVersion], [false, 'the document changed since the click', 2]);
  assert.equal(document.getText(), '- - [x] a\n- [ ] b\n');
});

test('H8-F3: a checkbox the renderer did not make (raw HTML) cannot toggle a line', async () => {
  // The attack: a raw-HTML checkbox with the task class and the data-line of a REAL task (line 2). The host cannot tell
  // its click from a click on the real checkbox, so the renderer must make every raw <input> disabled (the webview
  // ignores disabled checkboxes).
  const text = '<input type="checkbox" class="task-list-item-checkbox" data-line="2"> [ ] not a task\n\n- [ ] real task\n';
  const { document, init, toggle } = await openMarkdown('raw.md', text);
  assert.match(init.html, /<input disabled type="checkbox" class="task-list-item-checkbox" data-line="2">/, 'raw checkbox is inert');
  const checkboxes = init.html.match(/<input\b[^>]*>/g) ?? [];
  const enabled = checkboxes.filter((tag) => !/\sdisabled\b/.test(tag) && /class="task-list-item-checkbox"/.test(tag));
  assert.equal(checkboxes.length, 2);
  assert.equal(enabled.length, 1, 'only the renderer-made checkbox is enabled');
  assert.match(enabled[0], /data-line="2"/, 'the real task is there');
  const refused = await toggle(0, true, 1);
  assert.deepEqual([refused.applied, refused.reason], [false, 'no task checkbox on that line']);
  assert.equal(document.getText(), text);
  assert.deepEqual(callsOf('applyEdit'), []);
});

test('R2H4-2: a read-only document (files.readonlyInclude) refuses task toggles and says so in init', async () => {
  T.config.set('files.readonlyInclude', { '**/locked/**': true });
  const { document, init, toggle } = await openMarkdown('locked/tasks.md', '- [ ] a\n');
  assert.equal(init.readOnly, true);
  const refused = await toggle(0, true, 1);
  assert.deepEqual([refused.applied, refused.reason], [false, 'the document is read-only']);
  assert.equal(document.getText(), '- [ ] a\n');
  assert.deepEqual(callsOf('applyEdit'), []);
  // readonlyExclude wins over readonlyInclude, re-checked at the click.
  T.config.set('files.readonlyExclude', { '**/tasks.md': true });
  const allowed = await toggle(0, true, 1);
  assert.equal(allowed.applied, true);
  assert.equal(document.getText(), '- [x] a\n');
});

// ===== TEXT FILES VS CODE DOES NOT PASS TO EXTENSIONS =====

/** Runs the unreadable-tab watcher on a FileStudio tab of `uri`; returns the recorded executeCommand + warning calls. */
async function watchTab(uri, viewType) {
  const tab = T.addTab({ input: new vscode.TabInputCustom(uri, viewType), label: path.basename(uri.fsPath), isActive: true });
  const watcher = host.watchUnreadableTextTabs();
  try {
    await waitFor(() => callsOf('warning').length > 0, 3000, 'warning');
  } finally {
    watcher.dispose();
  }
  return { tab, open: callsOf('executeCommand').filter((c) => c[1] === 'vscode.openWith'), warnings: callsOf('warning').map((c) => c[1]) };
}

test('H3-F2: a CSV over the 50 MB extension limit reopens in the text editor with a message (no generic error)', async () => {
  const file = path.join(tmp, 'big.csv');
  const fd = fs.openSync(file, 'w');
  try {
    fs.writeSync(fd, 'a,b\n');
    fs.ftruncateSync(fd, 51 * 1024 * 1024);
  } finally {
    fs.closeSync(fd);
  }
  const uri = vscode.Uri.file(file);
  try {
    const { tab, open, warnings } = await watchTab(uri, 'fileStudio.csv');
    assert.equal(open.length, 1);
    assert.equal(open[0][2].fsPath, uri.fsPath);
    assert.equal(open[0][3], 'default', 'the text editor');
    assert.ok(!T.tabGroup.tabs.includes(tab), 'the FileStudio tab is closed');
    assert.equal(warnings.length, 1);
    assert.match(warnings[0], /^"big\.csv" \(51\.0 MB\) is too large for FileStudio: VS Code passes text files of up to 50 MB/);
  } finally {
    fs.rmSync(file, { force: true });
  }
});

test('R2H4-1: a .csv / .md that VS Code considers binary reopens in the text editor with VS Code\'s reason', async () => {
  for (const [name, viewType] of [['binary.csv', 'fileStudio.csv'], ['binary.md', 'fileStudio.markdown']]) {
    T.reset();
    const uri = writeTmp(name, Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x08, 0x00]));
    const { open, warnings } = await watchTab(uri, viewType);
    assert.equal(open.length, 1, name);
    assert.deepEqual(warnings, [
      `FileStudio cannot show "${name}": VS Code does not open it as text (File seems to be binary and cannot be opened as text). It was opened in the text editor instead.`,
    ]);
  }
});

// ===== OPEN WITH FILESTUDIO / REGISTRATION =====

test('R3-028: every custom editor is registered with supportsMultipleEditorsPerDocument (split editor, open to the side)', () => {
  host.extension.activate({ extensionUri, subscriptions: [] });
  const providers = T.registrations.providers;
  assert.deepEqual(providers.map((p) => p.viewType).sort(), pkg.contributes.customEditors.map((e) => e.viewType).sort());
  for (const { viewType, options } of providers) assert.equal(options?.supportsMultipleEditorsPerDocument, true, viewType);
});

test('R3-027: Open with FileStudio on a text tab replaces it; a tab with unsaved changes stays open', async () => {
  host.extension.activate({ extensionUri, subscriptions: [] });
  const openWith = T.registrations.commands.get('fileStudio.openWith');
  for (const isDirty of [false, true]) {
    T.calls.length = 0;
    T.tabGroup.tabs = [];
    const uri = writeTmp('table.csv', 'a,b\n');
    const textTab = T.addTab({ input: new vscode.TabInputText(uri), label: 'table.csv', isActive: true, isDirty });
    await openWith(); // command palette / editor title: the active tab's resource
    const opened = callsOf('executeCommand').filter((c) => c[1] === 'vscode.openWith');
    assert.equal(opened.length, 1);
    assert.equal(opened[0][2].fsPath, uri.fsPath);
    assert.equal(opened[0][3], 'fileStudio.csv');
    assert.equal(T.tabGroup.tabs.includes(textTab), isDirty, isDirty ? 'dirty text tab kept' : 'clean text tab closed');
  }
});

test('size limit: a workbook over the limit shows the "too large to edit" banner in init', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.getCell(1001, 100).value = 1; // 100,100 cells
  const uri = writeTmp('too-large.xlsx', Buffer.from(await workbook.xlsx.writeBuffer()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const document = provider.openCustomDocument(uri, {}, token);
  const panel = T.createPanel();
  provider.resolveCustomEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 5000, 'init');
  assert.match(init.banner, /too large to edit safely in this version/);
  assert.match(init.banner, /later version/);
  assert.equal(init.readOnly, true);
});
