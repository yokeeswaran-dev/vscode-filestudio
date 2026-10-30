// Opens every sample (in the temporary workspace copy of test/samples) with its FileStudio editor in a real VS Code.
// - One test per sample: the tab is the FileStudio custom editor and active, nothing is logged as an error or warning
//   in the first ~1.5 s, the document is not dirty, and the file on disk has the same bytes after closing. It does NOT
//   check that the viewer started: the host only warns about a missing 'ready' after 10 s.
// - "every viewer starts": ALL samples again, in batches of at most 9 visible editors side by side, kept open for
//   READY_TIMEOUT_MS + 2 s, so a viewer that never reports 'ready', or fails later (an uncaught webview error), is
//   logged and fails the test.
'use strict';
const assert = require('node:assert/strict');
const vscode = require('vscode');
const h = require('./helpers');

/** The host logs a warning when a VISIBLE viewer has not sent 'ready' after this long (READY_TIMEOUT_MS, viewerProvider.ts). */
const READY_TIMEOUT_MS = 10_000;
/** Editors open at the same time in the 'ready' check: a 3 x 3 grid of editor groups (ViewColumn.One .. Nine). */
const MAX_SIDE_BY_SIDE = 9;

/** Splits `list` into the fewest batches of at most `max` items, with about the same size. */
function batches(list, max) {
  const count = Math.ceil(list.length / max);
  const size = Math.ceil(list.length / count);
  const result = [];
  for (let i = 0; i < list.length; i += size) result.push(list.slice(i, i + size));
  return result;
}

suite('Open every sample with its FileStudio editor', () => {
  suiteSetup(async () => {
    assert.ok(h.extension(), `extension ${h.EXTENSION_ID} is not loaded`);
    await h.closeAll();
  });
  teardown(() => h.closeAll());

  test('the workspace has a sample for every FileStudio editor', () => {
    const covered = new Set(h.sampleFiles().map(h.viewTypeFor));
    assert.deepEqual([...covered].sort(), h.allViewTypes());
  });

  for (const name of h.sampleFiles()) {
    test(name, async () => {
      const before = h.readBytes(name);
      const mark = h.logMark();
      const { uri, viewType, tab } = await h.openWith(name);

      assert.ok(tab.input instanceof vscode.TabInputCustom, `the tab is ${h.describeTab(tab)}, expected custom:${viewType}`);
      assert.equal(tab.input.viewType, viewType);
      assert.ok(tab.isActive, 'the FileStudio tab is the active tab');
      assert.ok(h.extension().isActive, 'FileStudio is activated');

      await h.waitForLogQuiet();
      h.assertNoLogProblems(h.logSince(mark), name);
      assert.equal(h.isDirty(uri), false, `${name} is dirty after opening`);

      await h.closeAll();
      assert.ok(h.readBytes(name).equals(before), `${name} changed on disk after opening and closing it`);
    });
  }

  const readyBatches = batches(h.sampleFiles(), MAX_SIDE_BY_SIDE);
  readyBatches.forEach((names, index) => {
    const title = `every viewer starts (batch ${index + 1}/${readyBatches.length}): ${names.join(', ')} side by side report 'ready'`;
    test(title, { timeout: 90_000 }, async () => {
      const mark = h.logMark();
      // 3 columns x 3 rows: every viewer is visible at the same time, so each one's 'ready' timer runs.
      const column = { groups: [{}, {}, {}] };
      await vscode.commands.executeCommand('vscode.setEditorLayout', { orientation: 0, groups: [column, column, column] });
      try {
        let viewColumn = vscode.ViewColumn.One;
        for (const name of names) {
          await h.openWith(name, undefined, { viewColumn: viewColumn++, preserveFocus: true });
        }
        await h.sleep(READY_TIMEOUT_MS + 2_000);
        for (const name of names) {
          const tab = h.findTab(name, h.viewTypeFor(name));
          assert.ok(tab, `${name} is still open`);
          // The host's 'ready' timer only warns for a VISIBLE panel: each tab must be the shown tab of its own group.
          assert.ok(tab.isActive, `${name} is the active tab of its group (visible)`);
          assert.equal(tab.group.tabs.length, 1, `${name} has its editor group to itself`);
          assert.equal(h.isDirty(name), false, `${name} is dirty`);
        }
        await h.waitForLogQuiet({ min: 0 });
        h.assertNoLogProblems(h.logSince(mark), `viewers side by side (${names.join(', ')})`);
      } finally {
        await h.closeAll();
        await vscode.commands.executeCommand('workbench.action.editorLayoutSingle');
      }
    });
  });

  test('negative control: a broken .xlsx is logged as an error (proves the log check sees errors)', async () => {
    const name = 'zz-not-a-zip.xlsx';
    h.writeBytes(name, Buffer.from('this is not a zip file\n'));
    try {
      const mark = h.logMark();
      const { tab } = await h.openWith(name);
      assert.ok(tab.input instanceof vscode.TabInputCustom, `the tab is ${h.describeTab(tab)}`);
      await h.waitFor(() => h.logSince(mark).some((e) => e.level === 'error' && e.text.includes(name)), {
        timeout: 15_000,
        message: `an [error] entry for ${name} in the FileStudio log`,
      });
      assert.equal(h.isDirty(name), false, `${name} is dirty`);
    } finally {
      await h.closeAll();
      h.deleteFile(name);
    }
  });
});

suite('Test helpers', () => {
  teardown(() => h.closeAll());

  test('a text edit makes a FileStudio CSV editor dirty; undo, redo and closeAll (revert) work', async () => {
    const uri = h.copyFile('sample.csv', 'zz-helpers.csv');
    try {
      const before = h.readBytes(uri);
      await h.openWith(uri);
      const edit = new vscode.WorkspaceEdit();
      edit.insert(uri, new vscode.Position(0, 0), 'x');
      assert.ok(await vscode.workspace.applyEdit(edit), 'applyEdit');
      await h.waitFor(() => h.isDirty(uri), { message: 'the edited CSV to be dirty' });

      await h.undo();
      await h.waitFor(() => !h.isDirty(uri), { message: 'undo to clear the dirty state' });
      await h.redo();
      await h.waitFor(() => h.isDirty(uri), { message: 'redo to make it dirty again' });

      await h.closeAll(); // reverts, never saves, never shows the save dialog
      assert.equal(h.isDirty(uri), false, 'not dirty after closeAll');
      assert.ok(h.readBytes(uri).equals(before), 'the file on disk is unchanged');
    } finally {
      await h.closeAll();
      h.deleteFile(uri);
    }
  });
});
