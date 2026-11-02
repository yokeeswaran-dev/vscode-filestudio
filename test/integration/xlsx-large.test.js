// 0.2.1 big sheets in a REAL VS Code, on files made from large.xlsx (100,000 x 50 cells, 44 MB; not in Git).
//   case "20 MB sheet": the first 10,000 rows (19 MB of XML, 500,000 cells; editable): open, scroll, edit, undo, save, against
//     the budgets of contract 4.8 (open < 5 s, save < 5 s, no block > 250 ms, memory growth < 1 GB). Enforced.
//   case "large.xlsx": 191 MB of XML: opens fast (rows load by window) but stays read-only (over XLSX_LAZY_EDIT_BYTES);
//     measured and reported (FS_LARGE_EDIT=1 also tries edit + save: use it to decide whether the limit can be raised).
// Numbers go to FS_LARGE_OUT (default: the OS temp folder, large-timing.json). Runs only with FS_LARGE=1 and
// test/samples/large.xlsx (`npm run samples` makes it). Start it through the workshop limiter (heavy.js), see p9-bigsave\run-large.js.
'use strict';
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { monitorEventLoopDelay, performance } = require('node:perf_hooks');
const JSZip = require('jszip');
const h = require('./helpers');

const LARGE = path.resolve(__dirname, '..', 'samples', 'large.xlsx');
const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const mb = (n) => Math.round(n / 1048576);
const BUDGET = { openMs: 5000, saveMs: 5000, blockMs: 250, growthMb: 1024 };

/** large.xlsx cut to its first `rows` data rows. */
async function cutLarge(rows) {
  const zip = await JSZip.loadAsync(fs.readFileSync(LARGE));
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  let p = xml.indexOf('<sheetData>') + 11;
  for (let i = 0; i < rows; i++) p = xml.indexOf('</row>', p) + 6;
  const cut = (xml.slice(0, p) + xml.slice(xml.lastIndexOf('</sheetData>'))).replace(/<dimension ref="[^"]*"/, `<dimension ref="A1:AX${rows}"`);
  zip.file('xl/worksheets/sheet1.xml', cut);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function measure(api, label, bytes, { edit, rows }) {
  await h.closeAll();
  const name = `it-${label}.xlsx`;
  const uri = h.writeBytes(name, bytes);
  const key = uri.toString();
  const original = sha(uri.fsPath);
  const lag = monitorEventLoopDelay({ resolution: 10 });
  const numbers = { label, zipBytes: bytes.length };
  const rss0 = process.memoryUsage().rss;
  let id = 7000;
  const getRows = async (start, end) => {
    const t = performance.now();
    const reply = (await api.post(key, { type: 'getRows', reqId: ++id, sheet: 0, sheetId: 's1', start, end })).find((m) => m.type === 'rows');
    assert.ok(reply, 'rows reply');
    return { reply, ms: Math.round(performance.now() - t) };
  };
  const setCell = async (r, c, text) => {
    const t = performance.now();
    const editId = ++id;
    const messages = await api.post(key, { type: 'edit', editId, baseVersion: api.state(key).version, intent: { kind: 'setCells', sheetId: 's1', cells: [[r, c, text]] } });
    const result = messages.find((m) => m.type === 'editResult' && m.editId === editId);
    assert.ok(result, `editResult (got: ${JSON.stringify(messages).slice(0, 400)})`);
    return { result, ms: Math.round(performance.now() - t) };
  };

  lag.enable();
  let t = performance.now();
  await h.openWith(name);
  await h.waitFor(() => { try { return api.state(key).editState; } catch { return undefined; } }, { timeout: 180_000, message: `the edit state of ${name}` });
  const first = await getRows(0, 60);
  numbers.openMs = Math.round(performance.now() - t);
  numbers.openMaxBlockMs = Math.round(lag.max / 1e6);
  assert.equal(first.reply.rows.length, 60);
  lag.reset();
  numbers.scrollMs = [];
  for (const start of [Math.floor(rows / 2), rows - 60, Math.floor(rows / 10), Math.floor(rows * 0.7)]) numbers.scrollMs.push((await getRows(start, start + 60)).ms);
  numbers.scrollMaxBlockMs = Math.round(lag.max / 1e6);
  lag.reset();
  const editable = api.state(key).editState.editable === true;
  numbers.editable = editable;

  if (editable && edit) {
    const e1 = await setCell(0, 0, 'edited first cell');
    const e2 = await setCell(Math.floor(rows / 2), 3, '12345.5');
    const e3 = await setCell(rows - 1, 49, 'last cell');
    assert.ok([e1, e2, e3].every((e) => e.result.accepted === true), JSON.stringify([e1.result, e2.result, e3.result]));
    numbers.editMs = [e1.ms, e2.ms, e3.ms];
    numbers.editMaxBlockMs = Math.round(lag.max / 1e6);
    lag.reset();
    await h.undo();
    await h.undo();
    await h.undo();
    await h.waitFor(() => api.state(key).dirty === false, { message: 'clean after undo' });
    numbers.undoMaxBlockMs = Math.round(lag.max / 1e6);
    await h.save();
    assert.equal(sha(uri.fsPath), original, 'edit + undo writes nothing: the file is byte for byte the same');
    await h.redo();
    await h.redo();
    await h.redo();
    await h.waitFor(() => h.isDirty(uri), { message: 'dirty after redo' });
    lag.reset();
    t = performance.now();
    await h.save();
    await h.waitFor(() => !h.isDirty(uri), { timeout: 300_000, message: 'the tab to be clean after save' });
    numbers.saveMs = Math.round(performance.now() - t);
    numbers.saveMaxBlockMs = Math.round(lag.max / 1e6);
    const zip = await JSZip.loadAsync(fs.readFileSync(uri.fsPath));
    const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    assert.match(xml, /<c r="A1"[^>]*>(?:(?!<\/c>).)*edited first cell/);
    assert.match(xml, new RegExp(`<c r="D${Math.floor(rows / 2) + 1}"[^>]*>(?:(?!</c>).)*12345\\.5`));
    assert.match(xml, new RegExp(`<c r="AX${rows}"[^>]*>(?:(?!</c>).)*last cell`));
  }
  numbers.rssGrowthMb = mb(process.memoryUsage().rss - rss0);
  numbers.peakRssMb = mb(process.memoryUsage().rss);
  lag.disable();
  await h.closeAll();
  fs.rmSync(uri.fsPath, { force: true });
  return numbers;
}

function report(all) {
  fs.writeFileSync(process.env.FS_LARGE_OUT || path.join(os.tmpdir(), 'large-timing.json'), JSON.stringify(all, null, 2));
  console.log(`[large] ${JSON.stringify(all)}`);
}

suite('xlsx big sheets (0.2.1)', () => {
  const results = [];
  let api;
  suiteSetup(async () => {
    const ext = h.extension();
    api = ext.isActive ? ext.exports : await ext.activate();
  });

  test('a 20 MB sheet (500,000 cells): open, scroll, edit, undo, save within the budgets', { timeout: 600_000 }, async () => {
    if (process.env.FS_LARGE !== '1' || !fs.existsSync(LARGE)) return; // not asked for (or no file): skipped
    const n = await measure(api, 'sheet20mb', await cutLarge(10_000), { edit: true, rows: 10_000 });
    results.push(n);
    report(results);
    assert.equal(n.editable, true, 'a 20 MB sheet is editable in 0.2.1 (the 100,000-cell limit is gone)');
    assert.ok(n.openMs < BUDGET.openMs, `open ${n.openMs} ms (budget ${BUDGET.openMs})`);
    assert.ok(n.saveMs < BUDGET.saveMs, `save ${n.saveMs} ms (budget ${BUDGET.saveMs})`);
    for (const k of ['openMaxBlockMs', 'scrollMaxBlockMs', 'editMaxBlockMs', 'undoMaxBlockMs', 'saveMaxBlockMs']) assert.ok(n[k] < BUDGET.blockMs, `${k} ${n[k]} ms (budget ${BUDGET.blockMs})`);
    assert.ok(n.rssGrowthMb < BUDGET.growthMb, `memory growth ${n.rssGrowthMb} MB (budget ${BUDGET.growthMb})`);
  });

  test('large.xlsx (191 MB of XML): opens by window and scrolls; read-only over the limit', { timeout: 900_000 }, async () => {
    if (process.env.FS_LARGE !== '1' || !fs.existsSync(LARGE)) return;
    const n = await measure(api, 'large', fs.readFileSync(LARGE), { edit: process.env.FS_LARGE_EDIT === '1', rows: 100_000 });
    results.push(n);
    report(results);
    // The numbers of this case are reported, not asserted: open / scroll of a 191 MB sheet depend on the machine (other jobs, free RAM).
    if (process.env.FS_LARGE_EDIT !== '1') assert.equal(n.editable, false, 'large.xlsx stays read-only while XLSX_LAZY_EDIT_BYTES is 25 MB');
  });
});
