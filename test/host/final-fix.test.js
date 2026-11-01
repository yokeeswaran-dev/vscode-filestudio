// Final release review, fix round: regression tests for the confirmed findings (host side).
// Same setup as sheet-fix-r1.test.js: host code against test/host/fake-vscode.js.
'use strict';
const { test, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');

const { T, host, extensionUri, token, writeTmp } = require('./load-host');

const api = host.createTestApi();
const A1 = (text) => ({ kind: 'cells', sheet: 1, cells: [[0, 0, { value: { t: 's', v: text } }]] });

beforeEach(() => T.reset());
afterEach(() => {
  delete host.sheetSaveHooks.afterVerify;
});

let counter = 0;
async function open(name = 'book.xlsx') {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = 'hello';
  const uri = writeTmp(`ff${++counter}-${name}`, Buffer.from(await workbook.xlsx.writeBuffer()));
  const provider = new host.SheetEditorProvider(extensionUri);
  const doc = provider.openCustomDocument(uri, {}, token);
  const events = [];
  doc.onDidChange((e) => events.push(e));
  await doc.loading;
  return { uri, doc, events };
}
const isStep = (e) => typeof e.undo === 'function';

for (const direction of ['undo', 'redo']) {
  test(`${direction} pressed during a save: the document stays dirty for VS Code (a content-change event follows the step)`, async () => {
    const { uri, doc, events } = await open(`${direction}.xlsx`);
    await api.pushPackageOps(uri.toString(), [A1('one')]);
    await api.pushPackageOps(uri.toString(), [A1('two')]);
    const steps = events.filter(isStep);
    assert.equal(steps.length, 2);
    // undo: saving "two" while the user presses undo. redo: "two" was undone, saving "one" while the user presses redo.
    if (direction === 'redo') await steps[1].undo();
    let step;
    host.sheetSaveHooks.afterVerify = () => {
      delete host.sheetSaveHooks.afterVerify;
      step = direction === 'undo' ? steps[1].undo() : steps[1].redo(); // queued behind the save
    };
    const before = events.length;
    await doc.save(token);
    await step;
    const early = events.length;
    await new Promise((r) => setTimeout(r, 400)); // the mark comes after VS Code's save point, a moment after the step
    assert.ok(events.length > early, 'a content-change event comes after the step ended, not only during it');
    assert.equal(doc.hasUnsavedChanges, true, 'the document moved on after the save point');
    const after = events.slice(before);
    assert.ok(after.length > 0 && !isStep(after[after.length - 1]), 'the last event is a content-change event, so the tab is dirty');
  });
}

test('undo with no save around: no extra event (VS Code already counts the step; the tab shows dirty by itself)', async () => {
  const { uri, doc, events } = await open('plain-undo.xlsx');
  await api.pushPackageOps(uri.toString(), [A1('one')]);
  await doc.save(token);
  const before = events.length;
  await events.filter(isStep)[0].undo();
  assert.equal(doc.hasUnsavedChanges, true);
  assert.equal(events.length, before, 'nothing is fired by the step');
});
