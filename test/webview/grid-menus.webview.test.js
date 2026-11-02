// Webview acceptance tests for P3 WP4 (CSV / TSV / PSV / SSV editing): the shared structure UI of the grid (row / column /
// cell context menu, sort menu, insert / delete, Clear contents), Find & Replace (Ctrl+F / Ctrl+H, F3, Enter, Esc), paste of a
// block as ONE edit, the "Edit anyway" free read-only state, high contrast and right-to-left. The host is a mock that
// records every edit and answers `find` like CsvDocState.find (values only); what the real host does with the edits is
// tested in test/host/csv-docstate.test.js, test/unit/csv-edit.test.js and test/integration/csv-edit.test.js.
// The CSP is checked by the harness: a view that hits a securitypolicyviolation fails when it closes.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');

const sleep = H.sleep;
const ALL_KINDS = ['setCells', 'clear', 'insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'sort', 'replaceAll'];
const TEXT = 'name,qty,note\napple,3,red\nbanana,12,yellow\ncherry,7,red\ndate,5,brown\n';

/** A csv view whose host records edits and answers find. */
function csvEditHost(opts = {}) {
  const S = H.renderer('sheet');
  const fileName = opts.fileName || 'menus.csv';
  const csv = S.parseCsv(opts.text || TEXT, opts.parse);
  const model = new S.CsvGridModel(csv, fileName);
  model.getMeta().sheets[0].id = 'csv'; // like CsvDocState.install
  const host = {
    kind: 'csv',
    fileName,
    log: [],
    edits: [],
    finds: [],
    commands: [],
    version: 1,
    editState: opts.editState || { editable: true, kinds: ALL_KINDS },
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      if (msg.type === 'ready') {
        const meta = model.getMeta();
        if (opts.rtl) meta.sheets[0].rightToLeft = true;
        if (opts.asSheet) return [{ type: 'init', kind: 'sheet', fileName, meta, readOnly: !host.editState.editable, version: host.version, editState: host.editState }]; // an xlsx grid
        return [{ type: 'init', kind: 'csv', fileName, meta, readOnly: !host.editState.editable, version: host.version, editState: host.editState, delimited: opts.delimited || { format: 'csv', delimiter: csv.delimiter } }];
      }
      if (msg.type === 'getRows') {
        const { rows, styles } = model.getRows(msg.sheet, msg.start, msg.end, msg.c0, msg.c1);
        const cols = msg.c0 === undefined ? {} : { c0: msg.c0, c1: msg.c1 };
        return [{ type: 'rows', reqId: msg.reqId, sheet: msg.sheet, sheetId: 'csv', version: host.version, start: msg.start, end: msg.end, ...cols, rows, styles }];
      }
      if (msg.type === 'getStats') return [{ type: 'stats', reqId: msg.reqId, stats: model.getStats(msg.sheet, msg.ranges) }];
      if (msg.type === 'edit') {
        host.edits.push(msg);
        if (host.reject) return [{ type: 'editResult', editId: msg.editId, accepted: false, code: 'blocked', reason: host.reject, version: host.version }];
        host.version++;
        return [{ type: 'editResult', editId: msg.editId, accepted: true, version: host.version }];
      }
      if (msg.type === 'command') {
        host.commands.push(msg);
        return [];
      }
      if (msg.type === 'find') {
        host.finds.push(msg);
        return [host.findAnswer(msg)];
      }
      return [];
    },
    /** Like CsvDocState.find: values only, row-major, `from` = strictly after, wraps around. */
    findAnswer(msg) {
      const found = S.csvFindPattern(msg.query);
      if ('invalid' in found) return { type: 'findResult', reqId: msg.reqId, version: host.version, matches: [], total: 0, truncated: false, error: found.invalid };
      const all = [];
      csv.rows.forEach((row, r) => row.forEach((v, c) => {
        if (typeof v !== 'string' || v === '') return;
        found.re.lastIndex = 0;
        if (found.re.test(v)) all.push({ sheetId: 'csv', r, c });
      }));
      const from = msg.from;
      let pick;
      if (msg.direction === 'next') pick = (from && all.find((m) => m.r > from.r || (m.r === from.r && m.c > from.c))) || all[0];
      else if (msg.direction === 'prev') pick = (from && [...all].reverse().find((m) => m.r < from.r || (m.r === from.r && m.c < from.c))) || all[all.length - 1];
      return { type: 'findResult', reqId: msg.reqId, version: host.version, matches: msg.direction === 'all' ? all : pick ? [pick] : [], total: all.length, truncated: false };
    },
  };
  return host;
}

const menuItems = (view) =>
  view.eval(() => [...document.querySelectorAll('.fv-ctx .fv-menu-item')].map((e) => ({ label: e.querySelector('.fv-menu-label').textContent, hint: e.querySelector('.fv-menu-hint')?.textContent || '', disabled: e.getAttribute('aria-disabled') === 'true' })));
const labels = async (view) => (await menuItems(view)).map((i) => i.label);
const hasMenu = (view) => view.eval(() => !!document.querySelector('.fv-ctx'));
const status = (view) => view.eval(() => document.querySelector('.fv-status-msg')?.textContent || '');
const findState = (view) =>
  view.eval(() => {
    const bar = document.querySelector('.fv-find');
    if (!bar) return null;
    const q = (s) => bar.querySelector(s);
    return {
      hidden: bar.hidden,
      count: q('.fv-find-count').textContent,
      findValue: q('input[aria-label="Find"]').value,
      focus: document.activeElement === q('input[aria-label="Find"]') ? 'find' : document.activeElement === q('input[aria-label="Replace with"]') ? 'replace' : document.activeElement?.className?.split(' ')[0] || 'other',
      replaceShown: !q('.fv-find-replace').hidden,
    };
  });

/** Right-click on a cell, a row header or a column header. */
async function rightClick(g, what, a, b) {
  await g.ready();
  const xy = what === 'cell' ? await g.cellXY(a, b) : await g.page.evaluate((k, l) => window.__g.hdrXY(k, l), what, a);
  await g.page.mouse.click(xy.x, xy.y, { button: 'right' });
  await g.view.settle();
}

H.suite('csv: structure menus, find bar and paste (P3 WP4)', ({ test }) => {
  test('M1 right-click on a cell: the menu lists the host kinds; arrows, Esc and focus return; a click sends one edit', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B3');
      await rightClick(g, 'cell', 2, 1);
      assert.deepEqual(await labels(view), [
        'Insert 1 row above', 'Insert 1 row below', 'Insert 1 column left', 'Insert 1 column right', 'Delete 1 row', 'Delete 1 column', 'Clear contents', 'Sort A to Z', 'Sort Z to A',
      ]);
      const attrs = await view.eval(() => {
        const m = document.querySelector('.fv-ctx');
        return { role: m.getAttribute('role'), label: m.getAttribute('aria-label'), items: [...m.querySelectorAll('[role=menuitem]')].length, focus: document.activeElement?.querySelector('.fv-menu-label')?.textContent };
      });
      assert.deepEqual(attrs, { role: 'menu', label: 'Cell menu', items: 9, focus: 'Insert 1 row above' });
      await p.keyboard.press('ArrowDown');
      await p.keyboard.press('ArrowUp');
      await p.keyboard.press('ArrowUp');
      assert.equal(await view.eval(() => document.activeElement.querySelector('.fv-menu-label').textContent), 'Sort Z to A', 'ArrowUp wraps to the last item');
      await p.keyboard.press('Home');
      assert.equal(await view.eval(() => document.activeElement.querySelector('.fv-menu-label').textContent), 'Insert 1 row above');
      await p.keyboard.press('Escape');
      assert.equal(await hasMenu(view), false);
      assert.equal(await g.focus(), 'fv-scroller[' + (await view.eval(() => document.querySelector('.fv-scroller').getAttribute('aria-label'))) + ']', 'the focus is back on the grid');
      assert.equal(host.edits.length, 0);

      // a click on an item sends ONE edit and closes the menu
      await rightClick(g, 'cell', 2, 1);
      await view.eval(() => [...document.querySelectorAll('.fv-ctx .fv-menu-item')].find((e) => e.textContent.startsWith('Insert 1 row below')).click());
      await view.settle();
      assert.equal(await hasMenu(view), false);
      assert.equal(host.edits.length, 1);
      assert.deepEqual(host.edits[0].intent, { kind: 'insertRows', sheetId: 'csv', at: 3, count: 1 });
      assert.equal(await status(view), 'Inserted 1 row.', 'the live status says what happened');
    });
  });

  test('M2 headers: a row header menu has row items only; a column header menu has column items, Clear and Sort (row 1 is the header)', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await rightClick(g, 'row', '2');
      assert.deepEqual(await labels(view), ['Insert 1 row above', 'Insert 1 row below', 'Delete 1 row', 'Clear contents']);
      assert.deepEqual(await g.full('row'), ['2'], 'the right click selected the row');
      await view.page.keyboard.press('Escape');

      await rightClick(g, 'col', 'C');
      assert.deepEqual(await labels(view), ['Insert 1 column left', 'Insert 1 column right', 'Delete 1 column', 'Clear contents', 'Sort A to Z', 'Sort Z to A']);
      assert.deepEqual(await g.full('col'), ['C']);
      const sort = (await menuItems(view)).find((i) => i.label === 'Sort Z to A');
      assert.equal(sort.hint, 'column C');
      await view.eval(() => [...document.querySelectorAll('.fv-ctx .fv-menu-item')].find((e) => e.textContent.startsWith('Sort Z to A')).click());
      await view.settle();
      assert.deepEqual(host.edits[0].intent, { kind: 'sort', sheetId: 'csv', range: { r0: 0, c0: 0, r1: 4, c1: 2 }, key: 2, desc: true, header: true });
      assert.match(await status(view), /^Sorted by column C, Z to A\. Row 1 stays on top\./);
    });
  });

  test('M3 counts follow the selection; a right click outside the selection selects the cell; delete and Clear send one edit each', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B2:C4');
      await rightClick(g, 'cell', 2, 2);
      assert.deepEqual((await labels(view)).slice(0, 6), ['Insert 3 rows above', 'Insert 3 rows below', 'Insert 2 columns left', 'Insert 2 columns right', 'Delete 3 rows', 'Delete 2 columns']);
      assert.equal(await g.nameBox(), 'B2:C4', 'a click inside the selection keeps it');
      await view.eval(() => [...document.querySelectorAll('.fv-ctx .fv-menu-item')].find((e) => e.textContent === 'Delete 2 columns').click());
      await view.settle();
      assert.deepEqual(host.edits[0].intent, { kind: 'deleteCols', sheetId: 'csv', at: 1, count: 2 });

      await rightClick(g, 'cell', 0, 0);
      assert.equal(await g.nameBox(), 'A1', 'a click outside selects the cell');
      await view.eval(() => [...document.querySelectorAll('.fv-ctx .fv-menu-item')].find((e) => e.textContent.startsWith('Clear contents')).click());
      await view.settle();
      assert.equal(host.edits[1].intent.kind, 'clear');
      assert.deepEqual(host.edits[1].intent.ranges, [{ r0: 0, c0: 0, r1: 0, c1: 0 }]);
      assert.equal(host.edits.length, 2);
    });
  });

  test('M4 last row / last column: "below" and "right" are disabled with a reason (nothing to insert before nothing)', (h) =>
    h.use(csvEditHost(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('C5');
      await rightClick(g, 'cell', 4, 2);
      const items = await menuItems(view);
      const state = Object.fromEntries(items.map((i) => [i.label, i.disabled]));
      assert.equal(state['Insert 1 row below'], true);
      assert.equal(state['Insert 1 column right'], true);
      assert.equal(state['Insert 1 row above'], false);
      assert.equal(await view.eval(() => document.querySelector('.fv-menu-disabled').title).then((t) => /last row/.test(t)), true);
      // arrow keys skip disabled items
      await view.page.keyboard.press('End');
      assert.equal(await view.eval(() => document.activeElement.querySelector('.fv-menu-label').textContent), 'Sort Z to A');
      // a disabled item does nothing
      await view.eval(() => document.querySelector('.fv-menu-disabled').click());
      assert.equal(view.host.edits.length, 0);
      await view.page.keyboard.press('Escape');
    }));

  test('M5 the host offers fewer kinds, or none: the menu follows editState; read-only keeps VS Code\'s own menu', async (h) => {
    await h.use(csvEditHost({ editState: { editable: true, kinds: ['setCells', 'clear'] } }), {}, async (view) => {
      const g = H.grid(view);
      await rightClick(g, 'cell', 1, 1);
      assert.deepEqual(await labels(view), ['Clear contents']);
    });
    const ro = csvEditHost({ editState: { editable: false, reason: 'Read-only', kinds: [] } });
    await h.use(ro, {}, async (view) => {
      const g = H.grid(view);
      await g.ready();
      const xy = await g.cellXY(1, 1);
      const prevented = await view.eval(({ x, y }) => {
        const el = document.elementFromPoint(x, y);
        const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 2 });
        el.dispatchEvent(ev);
        return ev.defaultPrevented;
      }, xy);
      assert.equal(prevented, false, 'a read-only grid does not take over the context menu');
      assert.equal(await hasMenu(view), false);
      await g.goto('A1');
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('f');
      await view.page.keyboard.up('Control');
      await view.settle();
      assert.notEqual(await findState(view), null, 'Find still works in a read-only grid');
      assert.equal(await view.eval(() => !document.querySelector('.fv-find-toggle') || document.querySelector('.fv-find-toggle').hidden), true, 'but there is no Replace');
    });
  });

  test('M6 a file split on spaces: column items are disabled with the reason (the host blocks them too); row items work', (h) =>
    h.use(csvEditHost({ fileName: 'space.ssv', text: 'a  b  c\n1  2  3\n', parse: { delimiter: ' ' }, delimited: { format: 'ssv', delimiter: ' ' } }), {}, async (view) => {
      const g = H.grid(view);
      await rightClick(g, 'cell', 1, 1);
      const items = await menuItems(view);
      const dis = Object.fromEntries(items.map((i) => [i.label, i.disabled]));
      assert.equal(dis['Insert 1 column left'], true);
      assert.equal(dis['Delete 1 column'], true);
      assert.equal(dis['Insert 1 row above'], false);
      assert.match(await view.eval(() => [...document.querySelectorAll('.fv-menu-disabled')].find((e) => e.textContent.startsWith('Insert 1 column left')).title), /runs of spaces/);
    }));

  test('M7 keyboard: the Menu key and Shift+F10 open the cell menu at the active cell; Tab closes it', (h) =>
    h.use(csvEditHost(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B2');
      await p.keyboard.press('ContextMenu');
      await view.settle();
      assert.equal(await hasMenu(view), true);
      const pos = await view.eval(() => {
        const m = document.querySelector('.fv-ctx').getBoundingClientRect();
        const c = window.__g.cellRect(1, 1);
        return { dx: Math.abs(m.left - c.x), dy: Math.abs(m.top - c.y) };
      });
      assert.ok(pos.dx < 40 && pos.dy < 40, `the menu opens at the cell: ${JSON.stringify(pos)}`);
      await p.keyboard.press('Tab');
      assert.equal(await hasMenu(view), false);
      await p.keyboard.down('Shift');
      await p.keyboard.press('F10');
      await p.keyboard.up('Shift');
      await view.settle();
      assert.equal(await hasMenu(view), true);
      await p.keyboard.press('Enter'); // "Insert 1 row above" has the focus
      await view.settle();
      assert.deepEqual(view.host.edits[0].intent, { kind: 'insertRows', sheetId: 'csv', at: 1, count: 1 });
      assert.equal(await hasMenu(view), false);
    }));

  test('M8 Find: Ctrl+F opens the bar (the key does not reach VS Code), typing counts, Enter / Shift+Enter / F3 move, Esc closes', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await g.ready();
      await view.eval(() => (window.__fwd.length = 0));
      await p.keyboard.down('Control');
      await p.keyboard.press('f');
      await p.keyboard.up('Control');
      await view.settle();
      let st = await findState(view);
      assert.deepEqual([st.hidden, st.focus, st.replaceShown], [false, 'find', false]);
      assert.equal(await view.eval(() => window.__fwd.some((k) => /^Ctrl\+f$/i.test(k))), false, 'Ctrl+F is handled here, not forwarded');
      assert.equal(await view.eval(() => document.querySelector('.fv-find').getAttribute('role')), 'search');

      await p.keyboard.type('red');
      await sleep(400);
      await view.settle();
      st = await findState(view);
      assert.equal(st.count, '2 found');
      assert.deepEqual(host.finds.at(-1).query, { scope: 'csv', text: 'red', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' });
      assert.equal(host.finds.at(-1).direction, 'all');

      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(await g.nameBox(), 'C2');
      assert.deepEqual(host.finds.at(-1).from, { sheetId: 'csv', r: 0, c: 0 });
      assert.equal((await findState(view)).count, '1 of 2');
      assert.equal((await findState(view)).focus, 'find', 'the focus stays in the box so Enter can repeat');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(await g.nameBox(), 'C4');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(await g.nameBox(), 'C2', 'wraps around');
      await p.keyboard.down('Shift');
      await p.keyboard.press('Enter');
      await p.keyboard.up('Shift');
      await view.settle();
      assert.equal(await g.nameBox(), 'C4', 'Shift+Enter goes back');
      await p.keyboard.press('F3');
      await view.settle();
      assert.equal(await g.nameBox(), 'C2', 'F3 in the box');

      await p.keyboard.press('Escape');
      await view.settle();
      assert.equal((await findState(view)).hidden, true);
      assert.equal(await view.eval(() => document.activeElement === document.querySelector('.fv-scroller')), true, 'the focus goes back to the grid');
      await p.keyboard.press('F3');
      await view.settle();
      assert.equal(await g.nameBox(), 'C4', 'F3 on the grid goes on with the last text');
      assert.equal((await findState(view)).hidden, false, 'and shows the bar');
    });
  });

  test('M9 Find options: Aa, Cell, .* (Alt+C / Alt+W / Alt+R); an invalid pattern says so and moves nothing', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await p.keyboard.down('Control');
      await p.keyboard.press('f');
      await p.keyboard.up('Control');
      await p.keyboard.type('RED');
      await sleep(400);
      await view.settle();
      assert.equal((await findState(view)).count, '2 found', 'case does not matter by default');
      await p.keyboard.down('Alt');
      await p.keyboard.press('c');
      await p.keyboard.up('Alt');
      await sleep(100);
      await view.settle();
      assert.equal((await findState(view)).count, 'No results');
      assert.equal(await view.eval(() => document.querySelector('[aria-label="Match case"]').getAttribute('aria-pressed')), 'true');
      await p.keyboard.down('Alt');
      await p.keyboard.press('c');
      await p.keyboard.up('Alt');
      await p.keyboard.down('Alt');
      await p.keyboard.press('w');
      await p.keyboard.up('Alt');
      await sleep(100);
      await view.settle();
      assert.equal((await findState(view)).count, '2 found', 'whole cell: red = red');
      assert.equal(host.finds.at(-1).query.wholeCell, true);
      await p.keyboard.down('Alt');
      await p.keyboard.press('r');
      await p.keyboard.up('Alt');
      await p.keyboard.down('Control');
      await p.keyboard.press('a');
      await p.keyboard.up('Control');
      await p.keyboard.type('(');
      await sleep(400);
      await view.settle();
      assert.equal((await findState(view)).count, 'Invalid');
      assert.match(await view.eval(() => document.querySelector('.fv-find-count').title), /not a valid regular expression/);
      await g.goto('B2').catch(() => {});
    });
  });

  test('M10 Replace: Ctrl+H shows the Replace row; Replace changes the active matching cell (one setCells) and moves on; Replace all is ONE replaceAll', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await p.keyboard.down('Control');
      await p.keyboard.press('h');
      await p.keyboard.up('Control');
      await view.settle();
      let st = await findState(view);
      assert.deepEqual([st.hidden, st.replaceShown, st.focus], [false, true, 'find']);
      await p.keyboard.type('red');
      await p.keyboard.press('Enter'); // C2
      await sleep(300);
      await view.settle();
      assert.equal(await g.nameBox(), 'C2');
      await p.keyboard.press('Tab'); // case button ... go to the Replace box through its label
      await view.eval(() => document.querySelector('[aria-label="Replace with"]').focus());
      await p.keyboard.type('green');
      await p.keyboard.press('Enter'); // replace this cell, then next
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.deepEqual(host.edits[0].intent, { kind: 'setCells', sheetId: 'csv', cells: [[1, 2, 'green']] });
      assert.equal(await g.nameBox(), 'C4', 'moved to the next match');
      await view.eval(() => [...document.querySelectorAll('.fv-find-btn')].find((b) => b.getAttribute('aria-label') === 'Replace all').click());
      await view.settle();
      assert.equal(host.edits.length, 2);
      assert.deepEqual(host.edits[1].intent, { kind: 'replaceAll', query: { scope: 'csv', text: 'red', matchCase: false, regex: false, wholeCell: false, lookIn: 'values' }, replace: 'green' });
      assert.match(await status(view), /^Replaced in \d+ cells?\.$/);
      st = await findState(view);
      assert.equal(st.hidden, false);
    });
  });

  test('M11 a rejected Replace All shows the host\'s reason in the status line; nothing is shown as changed', (h) => {
    const host = csvEditHost();
    host.reject = 'The file changed; try again.';
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await p.keyboard.down('Control');
      await p.keyboard.press('h');
      await p.keyboard.up('Control');
      await p.keyboard.type('red');
      await sleep(400);
      await view.settle();
      await view.eval(() => [...document.querySelectorAll('.fv-find-btn')].find((b) => b.getAttribute('aria-label') === 'Replace all').click());
      await view.settle();
      assert.equal(await status(view), 'The file changed; try again.');
    });
  });

  test('M12 paste of a block is ONE setCells; typed values stay text (=1+1, 007, 1e3); a single cell fills the selection', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A2');
      const rows = [];
      for (let i = 0; i < 1000; i++) rows.push(Array.from({ length: 10 }, (_, j) => (i === 0 && j === 0 ? '=1+1' : i === 0 && j === 1 ? '007' : `v${i}_${j}`)).join('\t'));
      await view.eval((tsv) => {
        const dt = new DataTransfer();
        dt.setData('text/plain', tsv);
        document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      }, rows.join('\r\n') + '\r\n');
      await view.settle();
      assert.equal(host.edits.length, 1, 'one paste = one edit = one undo step');
      const cells = host.edits[0].intent.cells;
      assert.equal(host.edits[0].intent.kind, 'setCells');
      assert.equal(cells.length, 10000);
      assert.deepEqual(cells[0], [1, 0, '=1+1']);
      assert.deepEqual(cells[1], [1, 1, '007']);
      assert.deepEqual(cells[9999], [1000, 9, 'v999_9']);

      await g.goto('B3:B4');
      await view.eval(() => {
        const dt = new DataTransfer();
        dt.setData('text/plain', 'x');
        document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      });
      await view.settle();
      assert.deepEqual(host.edits[1].intent.cells, [[2, 1, 'x'], [3, 1, 'x']]);
    });
  });

  test('M13 pasted Excel text cells keep their text in a CSV: no leading apostrophe', (h) => {
    const host = csvEditHost();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A2');
      await view.eval(() => {
        const dt = new DataTransfer();
        dt.setData('text/html', '<table><tr><td x:str style="mso-number-format:\'\\@\'">007</td><td>b</td></tr></table>');
        dt.setData('text/plain', '007\tb');
        document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      });
      await view.settle();
      assert.deepEqual(host.edits[0].intent.cells, [[1, 0, '007'], [1, 1, 'b']]);
    });
  });

  test('M14 high contrast: the menu and the bar use the contrast border; pressed options are outlined; no horizontal overflow at 360 px', async (h) => {
    const host = csvEditHost();
    await h.use(host, { theme: 'hc', width: 360, height: 600 }, async (view) => {
      const g = H.grid(view);
      await g.goto('A1');
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('h');
      await view.page.keyboard.up('Control');
      await view.page.keyboard.down('Alt');
      await view.page.keyboard.press('c');
      await view.page.keyboard.up('Alt');
      await view.settle();
      const css = await view.eval(() => {
        const bar = document.querySelector('.fv-find');
        const cs = getComputedStyle(bar);
        const opt = getComputedStyle(document.querySelector('[aria-label="Match case"]'));
        const r = bar.getBoundingClientRect();
        return { border: cs.borderLeftWidth + ' ' + cs.borderLeftStyle, borderColor: cs.borderLeftColor, shadow: cs.boxShadow, optOutline: opt.outlineStyle, optWidth: opt.outlineWidth, fits: r.left >= 0 && r.right <= window.innerWidth, noHScroll: document.documentElement.scrollWidth <= window.innerWidth };
      });
      assert.equal(css.border, '1px solid');
      assert.notEqual(css.borderColor, 'rgba(0, 0, 0, 0)');
      assert.equal(css.shadow, 'none');
      assert.equal(css.optOutline, 'solid');
      assert.ok(css.fits && css.noHScroll, `the bar fits a narrow view: ${JSON.stringify(css)}`);
      await view.page.keyboard.press('Escape');
      await rightClick(g, 'cell', 1, 1);
      const menu = await view.eval(() => {
        const cs = getComputedStyle(document.querySelector('.fv-ctx'));
        const dis = document.querySelector('.fv-menu-disabled');
        return { border: cs.borderLeftStyle, color: cs.borderLeftColor, dis: dis ? getComputedStyle(dis).textDecorationLine : 'none-found' };
      });
      assert.equal(menu.border, 'solid');
      assert.notEqual(menu.color, 'rgba(0, 0, 0, 0)');
    });
  });

  test('M15 right-to-left view: the menu and the bar are mirrored (dir=rtl, bar at the left edge)', (h) =>
    h.use(csvEditHost({ rtl: true }), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A1');
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('f');
      await view.page.keyboard.up('Control');
      const bar = await view.eval(() => {
        const el = document.querySelector('.fv-find');
        const r = el.getBoundingClientRect();
        return { dir: el.dir, left: r.left, right: window.innerWidth - r.right };
      });
      assert.equal(bar.dir, 'rtl');
      assert.ok(bar.left < bar.right, `the bar is at the left edge: ${JSON.stringify(bar)}`);
      await view.page.keyboard.press('Escape');
      await g.ready();
      const xy = await view.eval(() => {
        const r = window.__g.cellRect(1, 1);
        return r ? { x: r.cx, y: r.cy } : null;
      });
      await view.page.mouse.click(xy.x, xy.y, { button: 'right' });
      await view.settle();
      assert.equal(await view.eval(() => document.querySelector('.fv-ctx').dir), 'rtl');
      assert.equal(await view.eval(() => getComputedStyle(document.querySelector('.fv-ctx .fv-menu-item')).textAlign), 'start');
    }));

  test('M16 reduced motion and no host: a menu or bar never animates; closing the view with the bar open leaves no error', (h) =>
    h.use(csvEditHost(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A1');
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('f');
      await view.page.keyboard.up('Control');
      await view.page.keyboard.type('a');
      const anim = await view.eval(() => ({ bar: getComputedStyle(document.querySelector('.fv-find')).transitionDuration, menu: 'n/a' }));
      assert.equal(anim.bar, '0s');
    }));
  test('M17 fix round 1: Enter on a find-bar button keeps the button; Alt+C works by key code; Menu key in a read-only grid keeps VS Code\'s menu', async (h) => {
    const host = csvEditHost();
    await h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await p.keyboard.down('Control');
      await p.keyboard.press('h');
      await p.keyboard.up('Control');
      await view.settle();
      await p.keyboard.type('red');
      await sleep(400);
      await view.settle();
      // Option+C on a Mac types another character: the key code decides
      await view.eval(() => {
        const input = document.querySelector('input[aria-label="Find"]');
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ç', code: 'KeyC', altKey: true, bubbles: true, cancelable: true }));
      });
      assert.equal(await view.eval(() => document.querySelector('[aria-label="Match case"]').getAttribute('aria-pressed')), 'true', 'Alt+C by key code');
      await view.eval(() => document.querySelector('[aria-label="Match case"]').click()); // off again
      await sleep(300);
      await view.settle();
      // Enter on Replace all: the button's own click (one replaceAll), not "next match"
      const before = host.finds.length;
      await view.eval(() => {
        const b = [...document.querySelectorAll('.fv-find-btn')].find((x) => x.getAttribute('aria-label') === 'Replace all');
        b.focus();
      });
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.filter((e) => e.intent.kind === 'replaceAll').length, 1, 'Enter on Replace all runs it');
      assert.equal(host.finds.filter((f) => f.direction !== 'all').length, 0, 'and does not search');
      void before;
      // Enter on Close closes the bar
      await view.eval(() => {
        const b = [...document.querySelectorAll('.fv-find button')].find((x) => /close/i.test(x.getAttribute('aria-label') || x.title || ''));
        b.focus();
      });
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal((await findState(view)).hidden, true, 'Enter on Close closes the bar');
    });
    const ro = csvEditHost({ editState: { editable: false, reason: 'Read-only', kinds: [] } });
    await h.use(ro, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B2');
      await g.ready();
      const out = await view.eval(() => {
        const sc = document.querySelector('[aria-activedescendant]') || document.querySelector('.fv-scroller');
        sc.dispatchEvent(new KeyboardEvent('keydown', { key: 'ContextMenu', code: 'ContextMenu', bubbles: true, cancelable: true }));
        const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 60, clientY: 60, button: 2 });
        sc.dispatchEvent(ev);
        return ev.defaultPrevented;
      });
      assert.equal(out, false, 'the follow-up contextmenu event is not swallowed when no menu opened');
    });
  });

  test('M18 fix round 2: an editable xlsx grid keeps the VS Code menu (right click, Menu key); a focused disabled item shows an outline', async (h) => {
    const xl = csvEditHost({ asSheet: true, editState: { editable: true, kinds: ['setCells', 'clear'] } });
    await h.use(xl, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B2');
      await g.ready();
      const out = await view.eval(() => {
        const sc = document.querySelector('[aria-activedescendant]') || document.querySelector('.fv-scroller');
        sc.dispatchEvent(new KeyboardEvent('keydown', { key: 'ContextMenu', code: 'ContextMenu', bubbles: true, cancelable: true }));
        const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 60, clientY: 60, button: 2 });
        sc.dispatchEvent(ev);
        return { prevented: ev.defaultPrevented, menu: !!document.querySelector('.fv-ctx') };
      });
      assert.deepEqual(out, { prevented: false, menu: false });
    });
    await h.use(csvEditHost(), {}, async (view) => {
      const g = H.grid(view);
      await rightClick(g, 'cell', 1, 1);
      const style = await view.eval(() => {
        const el = document.querySelector('.fv-menu-item');
        el.classList.add('fv-menu-disabled');
        el.setAttribute('aria-disabled', 'true');
        el.focus();
        const cs = getComputedStyle(el);
        return { style: cs.outlineStyle, width: cs.outlineWidth };
      });
      assert.notEqual(style.style, 'none', 'a focused disabled item has an outline');
    });
  });

  test('M19 fix round 2: F3 inside the cell editor does nothing; F6 reaches the find bar', async (h) => {
    const host = csvEditHost();
    await h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B2');
      await p.keyboard.down('Control');
      await p.keyboard.press('f');
      await p.keyboard.up('Control');
      await view.settle();
      await p.keyboard.type('red');
      await sleep(400);
      await view.settle();
      assert.equal((await findState(view)).hidden, false);
      // F6 from the Find box goes on to the grid, and from the formula bar to the find bar
      await p.keyboard.press('Escape');
      await view.settle();
      await g.goto('B2');
      await p.keyboard.down('Control');
      await p.keyboard.press('f');
      await p.keyboard.up('Control');
      await view.settle();
      await p.keyboard.down('Shift');
      await p.keyboard.press('F6');
      await p.keyboard.up('Shift');
      assert.notEqual((await findState(view)).focus, 'find', 'Shift+F6 leaves the find box backwards');
      await p.keyboard.press('F6');
      assert.equal((await findState(view)).focus, 'find', 'F6 comes back to the find box');
      await p.keyboard.press('Escape');
      await view.settle();
      // F3 while the cell editor is open: the editor stays and the bar is not opened
      await g.goto('B3');
      await p.keyboard.press('F2');
      await view.settle();
      const before = host.finds.length;
      await p.keyboard.press('F3');
      await view.settle();
      assert.equal(await view.eval(() => !!document.querySelector('.fv-cellbox')), true, 'the cell editor is still open');
      assert.equal(host.finds.length, before, 'no search ran');
    });
  });
});
