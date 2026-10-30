// Webview acceptance tests for P2 (xlsx core editing), stage b1: the grid cell editor against a mock host that plans and
// applies edits with the real edit model (planEdit / applyEdit). Covers: editor modes and every close path, keys
// (Enter / Tab / Esc / Delete / Ctrl+Z / Ctrl+S), IME, the formula bar, paste, the overlay, the send queue (one in flight,
// re-base, drop on a foreign patch), D-067 (undo held back while an edit is in flight), read-only state, tabs by id,
// the status bar and the "Edit anyway" banner, and that copy still works.
// The CSP is checked by the harness: a view that hits a securitypolicyviolation fails when it closes.
'use strict';
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const H = require('./harness');
const F = require('./fixtures');

const sleep = H.sleep;

/** cells.xlsx plus a formula, a number with a format, and a second plain sheet. */
async function editBook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let c = 1; c <= 8; c++) ws.getColumn(c).width = 14;
  ws.getCell('A1').value = 1;
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = { formula: 'A1+A2', result: 3 };
  ws.getCell('B1').value = 'Short';
  ws.getCell('B2').value = '007'; // text that would read as a number
  ws.getCell('B3').value = '=1+1'; // text that would read as a formula
  ws.getCell('B4').value = 12.5;
  ws.getCell('C1').value = 'one';
  ws.getCell('C2').value = 'two';
  ws.getCell('C3').value = 'three';
  ws.getCell('D5').value = 'last';
  const other = wb.addWorksheet('Other');
  other.getCell('A1').value = 'other sheet';
  const rtl = wb.addWorksheet('RTL', { views: [{ rightToLeft: true }] });
  rtl.getCell('B2').value = 'rtl';
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/**
 * An xlsx view that answers like the planned host: edits go through planEdit / applyEdit of the real edit model and come
 * back as editResult with the changed rows. `host.hold` keeps replies back until `host.release(view)`; `host.reject`
 * refuses the next edit. Everything the webview sent is in `host.edits` / `host.commands` / `host.requests`.
 */
function editHost(fileName, build, opts = {}) {
  const S = H.renderer('sheet');
  let loading;
  const load = () => (loading ||= (async () => S.workbookToModel(await S.loadWorkbook(new Uint8Array(await build()))))());
  const host = {
    kind: 'sheet',
    fileName,
    log: [],
    edits: [],
    commands: [],
    requests: [],
    version: 1,
    hold: false,
    held: [],
    reject: null,
    editState: opts.editState || { editable: true, kinds: ['setCells', 'clear'] },
    banner: opts.banner,
    anyway: 0,
    async receive(msg) {
      if (msg.type === 'log') {
        host.log.push(msg);
        return [];
      }
      const model = await load();
      if (msg.type === 'ready') {
        return [{ type: 'init', kind: 'sheet', fileName, meta: model.getMeta(), readOnly: !host.editState.editable, version: host.version, editState: host.editState, banner: host.banner }];
      }
      if (msg.type === 'getRows') {
        host.requests.push(msg);
        const { rows, styles } = model.getRows(msg.sheet, msg.start, msg.end, msg.c0, msg.c1);
        const cols = msg.c0 === undefined ? {} : { c0: msg.c0, c1: msg.c1 };
        return [{ type: 'rows', reqId: msg.reqId, sheet: msg.sheet, sheetId: model.sheetIdAt(msg.sheet), version: host.version, start: msg.start, end: msg.end, ...cols, rows, styles }];
      }
      if (msg.type === 'getStats') return [{ type: 'stats', reqId: msg.reqId, stats: model.getStats(msg.sheet, msg.ranges) }];
      if (msg.type === 'edit') {
        host.edits.push(msg);
        if (host.hold) {
          host.held.push(msg);
          return [];
        }
        return [host.answer(model, msg)];
      }
      if (msg.type === 'command') {
        host.commands.push(msg);
        return [];
      }
      if (msg.type === 'editAnyway') {
        host.anyway++;
        host.editState = { editable: true, kinds: ['setCells', 'clear'], canEditAnyway: undefined };
        return [{ type: 'editState', state: host.editState }];
      }
      return [];
    },
    /** One editResult for an edit (plans and applies it on the real model). */
    answer(model, msg) {
      if (host.reject) {
        const r = host.reject;
        host.reject = null;
        return { type: 'editResult', editId: msg.editId, accepted: false, code: r.code || 'blocked', reason: r.reason, version: host.version };
      }
      const plan = S.planEdit(model, msg.intent);
      if ('blocked' in plan) return { type: 'editResult', editId: msg.editId, accepted: false, code: 'blocked', reason: plan.blocked[0].message, version: host.version };
      if ('invalid' in plan) return { type: 'editResult', editId: msg.editId, accepted: false, code: 'invalid', reason: plan.invalid, version: host.version };
      if ('noop' in plan) return { type: 'editResult', editId: msg.editId, accepted: true, version: host.version };
      const change = S.applyEdit(model, plan);
      host.version++;
      const sheet = model.sheetIndexOf(plan.sheetId);
      const [start, end] = change.rows || [0, 0];
      const { rows, styles } = model.getRows(sheet, start, end);
      return { type: 'editResult', editId: msg.editId, accepted: true, version: host.version, patches: [{ sheetId: plan.sheetId, start, end, rows, styles }] };
    },
    /** Delivers the replies of the held edits, in order. */
    async release(view) {
      const model = await load();
      const held = host.held.splice(0);
      host.hold = false;
      for (const msg of held) await view.push(host.answer(model, msg));
      await view.settle();
    },
    model: () => load(),
  };
  return host;
}

const book = (opts) => editHost('edit.xlsx', editBook, opts);

/** Text shown by the grid for a cell (from its element), '' when none. */
const cellText = (view, text) => view.eval((t) => [...document.querySelectorAll('.fv-c')].some((e) => e.textContent === t), text);
const status = (view) => view.eval(() => document.querySelector('.fv-status-msg')?.textContent || '');
const mode = (view) => view.eval(() => document.querySelector('.fv-status-mode')?.textContent);

H.suite('sheet: grid cell editor (P2 core)', ({ test }) => {
  // ----- modes and start text -----

  test('E1 modes: F2 and double-click edit with the caret at the end, typing replaces; read-only keeps the 0.1.2 box', (h) =>
    h.use(book(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B1');
      await p.keyboard.press('F2');
      let b = await g.cellBox();
      assert.deepEqual([b.ro, b.value, b.start, b.end, b.focused], ['false', 'Short', 5, 5, true]);
      assert.equal(b.label, 'B1 cell text');
      await p.keyboard.press('Escape');
      assert.equal(await g.cellBox(), null);

      await g.clickCell(0, 1, [], { count: 2 });
      b = await g.cellBox();
      assert.deepEqual([b.ro, b.value, b.start, b.end], ['false', 'Short', 5, 5]);
      await p.keyboard.press('Escape');

      await g.goto('B1');
      await p.keyboard.type('X');
      b = await g.cellBox();
      assert.deepEqual([b.ro, b.value, b.start, b.end, b.focused], ['false', 'X', 1, 1, true]);
      assert.equal(await g.formula(), 'X', 'the formula bar follows the box');
      await p.keyboard.press('Escape');
      assert.equal(await g.formula(), 'Short');

      // a read-only grid: the box of 0.1.2 (display text, all selected, nothing can be typed)
      const ro = book({ editState: { editable: false, reason: 'Read-only' } });
      await h.use(ro, {}, async (rv) => {
        const rg = H.grid(rv);
        await rg.goto('B1');
        await rv.page.keyboard.type('X');
        assert.equal(await rg.cellBox(), null, 'typing does not open an editor');
        await rv.page.keyboard.press('F2');
        const rb = await rg.cellBox();
        assert.deepEqual([rb.ro, rb.value, rb.start, rb.end], ['true', 'Short', 5, 5]);
        await rv.page.keyboard.press('Escape');
        await rg.clickCell(0, 1, [], { count: 2 });
        const rb2 = await rg.cellBox();
        assert.deepEqual([rb2.ro, rb2.start, rb2.end], ['true', 0, 5], 'double-click selects all in the read-only box');
        await rv.page.keyboard.press('Escape');
        assert.equal(ro.edits.length, 0);
      });
    }));

  test('E2 edit start text: raw value, =formula, and an apostrophe for text that would read as a value (xlsx)', (h) =>
    h.use(book(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const start = async (ref) => {
        await g.goto(ref);
        await p.keyboard.press('F2');
        const b = await g.cellBox();
        await p.keyboard.press('Escape');
        return b.value;
      };
      assert.equal(await start('A3'), '=A1+A2', 'a formula cell starts as =formula, never its result');
      assert.equal(await start('B4'), '12.5');
      assert.equal(await start('B2'), "'007", 'text 007 would become the number 7');
      assert.equal(await start('B3'), "'=1+1", 'text =1+1 would become a formula');
      assert.equal(await start('B1'), 'Short', 'ordinary text has no apostrophe');
      // the formula bar shows the same text while the cell is only selected
      await g.goto('B2');
      assert.equal(await g.formula(), "'007");
      // unchanged text sends nothing
      await p.keyboard.press('F2');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(view.host.edits.length, 0, 'Enter without a change sends nothing');
    }));

  test('E3 commit keys: Enter down, Tab right, Shift+Enter up, Shift+Tab left; one char is a real edit; moves inside a selection', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const key = async (k, shift) => {
        if (shift) await p.keyboard.down('Shift');
        await p.keyboard.press(k);
        if (shift) await p.keyboard.up('Shift');
        await view.settle();
      };
      await g.goto('E3');
      await p.keyboard.type('x');
      await key('Enter');
      assert.equal(await g.nameBox(), 'E4');
      assert.equal(host.edits.length, 1, 'a single typed character is an edit');
      assert.deepEqual(host.edits[0].intent, { kind: 'setCells', sheetId: host.edits[0].intent.sheetId, cells: [[2, 4, 'x']] });

      await p.keyboard.type('y');
      await key('Tab');
      assert.equal(await g.nameBox(), 'F4');
      await p.keyboard.type('z');
      await key('Enter', true);
      assert.equal(await g.nameBox(), 'F3');
      await p.keyboard.type('w');
      await key('Tab', true);
      assert.equal(await g.nameBox(), 'E3');
      assert.deepEqual(host.edits.map((e) => e.intent.cells[0][2]), ['x', 'y', 'z', 'w']);

      // inside a selection of several cells Enter walks through it, like the grid
      await g.goto('H1:H3');
      await p.keyboard.type('a');
      await key('Enter');
      assert.equal(await g.nameBox(), 'H1:H3');
      await sleep(300); // the ARIA announcement of the active cell is debounced
      assert.equal(await g.view.eval(() => document.querySelector('.fv-scroller').getAttribute('aria-activedescendant') && document.getElementById(document.querySelector('.fv-scroller').getAttribute('aria-activedescendant')).getAttribute('aria-label').split(',')[0]), 'H2');
    });
  });

  test('E3b right-to-left sheet: Tab moves to the next column (drawn to the left), Enter down', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.clickTab('RTL');
      await g.goto('B2');
      await p.keyboard.type('m');
      await p.keyboard.press('Tab');
      await view.settle();
      assert.equal(await g.nameBox(), 'C2');
      const box = await g.cellBox();
      assert.equal(box, null);
      await p.keyboard.type('n');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(await g.nameBox(), 'C3');
      assert.deepEqual(host.edits.map((e) => e.intent.cells[0].slice(0, 2)), [[1, 1], [1, 2]]);
      // the editor opens mirrored: on the RTL sheet its box is placed from the right edge of the cell
      await g.goto('B2');
      await p.keyboard.press('F2');
      const where = await view.eval(() => {
        const cell = [...document.querySelectorAll('.fv-c')].find((e) => e.textContent === 'm');
        const box = document.querySelector('.fv-cellbox');
        return { cell: cell.getBoundingClientRect().right, box: box.getBoundingClientRect().right };
      });
      assert.ok(Math.abs(where.cell - where.box) < 3 || where.box >= where.cell - 3, 'box sits over the cell');
      await p.keyboard.press('Escape');
    });
  });

  test('E4 Esc cancels, Alt+Enter is a line break, Ctrl+S commits with then:save (unchanged: only a save command)', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B1');
      await p.keyboard.press('F2');
      await p.keyboard.type(' more');
      await p.keyboard.press('Escape');
      await view.settle();
      assert.equal(await g.cellBox(), null);
      assert.equal(host.edits.length, 0, 'Esc sends nothing');
      assert.equal(await g.formula(), 'Short');
      assert.equal(await cellText(view, 'Short'), true, 'the cell shows its old text');
      assert.equal(await g.view.eval(() => document.activeElement === document.querySelector('.fv-scroller')), true, 'the grid has the focus again');

      await p.keyboard.press('F2');
      await p.keyboard.down('Alt');
      await p.keyboard.press('Enter');
      await p.keyboard.up('Alt');
      await p.keyboard.type('line2');
      assert.equal((await g.cellBox()).value, 'Short\nline2');
      assert.equal(await g.formula(), 'Short\nline2');

      await p.keyboard.down('Control');
      await p.keyboard.press('KeyS');
      await p.keyboard.up('Control');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.equal(host.edits[0].then, 'save');
      assert.equal(host.edits[0].intent.cells[0][2], 'Short\nline2');
      const fwd = await view.eval(() => window.__fwd);
      assert.ok(!fwd.includes('Ctrl+s'), 'Ctrl+S in the box does not reach VS Code (the webview asks for the save itself)');

      // unchanged: only a save command
      await g.goto('C1');
      await p.keyboard.press('F2');
      await p.keyboard.down('Control');
      await p.keyboard.press('KeyS');
      await p.keyboard.up('Control');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.deepEqual(host.commands.map((c) => c.id), ['save']);
    });
  });

  test('E5 Delete / Backspace clear the contents of the selection: one `clear` edit for all ranges; Ctrl+X does nothing', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('C1:C2');
      await g.clickCell(2, 2, ['Control']); // a second range: C3
      await p.keyboard.press('Delete');
      await view.settle();
      assert.equal(host.edits.length, 1);
      const i = host.edits[0].intent;
      assert.equal(i.kind, 'clear');
      assert.equal(i.what, 'contents');
      assert.deepEqual(i.ranges, [{ r0: 0, c0: 2, r1: 1, c1: 2 }, { r0: 2, c0: 2, r1: 2, c1: 2 }]);
      assert.equal(await cellText(view, 'one'), false, 'cleared at once');
      assert.equal(await cellText(view, 'three'), false);
      await g.goto('D5');
      await p.keyboard.press('Backspace');
      await view.settle();
      assert.equal(host.edits.length, 2);
      assert.equal(host.edits[1].intent.kind, 'clear');
      assert.equal(await cellText(view, 'last'), false);

      // Ctrl+X: not in 0.2.0 (D-044)
      await g.goto('B1');
      await p.keyboard.down('Control');
      await p.keyboard.press('KeyX');
      await p.keyboard.up('Control');
      await view.settle();
      assert.match(await status(view), /Cut is not available yet/);
      assert.equal(host.edits.length, 2);
      assert.equal(await cellText(view, 'Short'), true);
    });
  });

  test('E5b typing: Space starts an edit, Shift+Space still selects the row, an AltGr character types, Ctrl+key does not', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E1');
      await p.keyboard.down('Shift');
      await p.keyboard.press('Space');
      await p.keyboard.up('Shift');
      await view.settle();
      assert.equal(await g.cellBox(), null, 'Shift+Space is not typing');
      assert.deepEqual(await g.full('row'), ['1']);
      await g.goto('E1');
      await p.keyboard.press('Space');
      assert.equal((await g.cellBox())?.value, ' ');
      await p.keyboard.press('Escape');
      // AltGr (Ctrl+Alt on Windows) produces a character
      await view.eval(() => {
        const sc = document.querySelector('.fv-scroller');
        window.__alt = sc.dispatchEvent(new KeyboardEvent('keydown', { key: '@', ctrlKey: true, altKey: true, bubbles: true, cancelable: true, modifierAltGraph: true }));
      });
      assert.equal((await g.cellBox())?.value, '@');
      await p.keyboard.press('Escape');
      // Ctrl+key and Alt+key are shortcuts
      await view.eval(() => {
        const sc = document.querySelector('.fv-scroller');
        sc.dispatchEvent(new KeyboardEvent('keydown', { key: 'q', ctrlKey: true, bubbles: true, cancelable: true }));
        sc.dispatchEvent(new KeyboardEvent('keydown', { key: 'q', altKey: true, bubbles: true, cancelable: true }));
      });
      assert.equal(await g.cellBox(), null);
      assert.equal(host.edits.length, 0);
    });
  });

  // ----- paste -----

  test('E6 paste: TSV (quoted fields), Excel HTML (inert, spans), a single cell fills the selection, limits', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const paste = async (data) => {
        await view.eval((data) => {
          const dt = new DataTransfer();
          for (const [type, value] of Object.entries(data)) dt.setData(type, value);
          document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
        }, data);
        await view.settle();
      };
      await g.goto('F1');
      await paste({ 'text/plain': '10\t20\r\n30\t"4\t0 ""q"""\r\n' });
      let i = host.edits[0].intent;
      assert.equal(i.kind, 'setCells');
      assert.deepEqual(i.cells.map((c) => c.slice(0, 2)), [[0, 5], [0, 6], [1, 5], [1, 6]]);
      assert.deepEqual(i.cells.map((c) => c[2]), ['10', '20', '30', '4\t0 "q"'], 'quoted field: tab and quotes kept; the final line break ends the last row');
      assert.equal(await g.nameBox(), 'F1:G2', 'the pasted block is selected');
      assert.equal(await cellText(view, '10'), true, 'shown at once');

      // 1x1 fills the selection
      await g.goto('F4:G5');
      await paste({ 'text/plain': 'Fill' });
      i = host.edits[1].intent;
      assert.equal(i.cells.length, 4);
      assert.ok(i.cells.every((c) => c[2] === 'Fill'));

      // HTML: a table (spans expanded); a script and an image in it are never run or inserted
      await g.goto('A7');
      await paste({
        'text/html': '<html><body><img src="x" onerror="window.__pwned=1"><script>window.__pwned=2<\/script><table><tr><td colspan="2">wide</td><td>c</td></tr><tr><td>a&nbsp;b</td><td>x<br>y</td><td rowspan="2">r</td></tr><tr><td>1</td><td>2</td></tr></table></body></html>',
        'text/plain': 'ignored',
      });
      i = host.edits[2].intent;
      assert.deepEqual(i.cells.map((c) => [c[0], c[1], c[2]]), [[6, 0, 'wide'], [6, 1, ''], [6, 2, 'c'], [7, 0, 'a b'], [7, 1, 'x\ny'], [7, 2, 'r'], [8, 0, '1'], [8, 1, '2'], [8, 2, '']]);
      assert.equal(await view.eval(() => window.__pwned), undefined, 'pasted HTML is inert');
      assert.equal(await view.eval(() => document.querySelectorAll('.fv-grid-app table, #app img').length), 0, 'pasted HTML is never inserted');

      // too far / too large: nothing is sent
      await g.goto('XFD1');
      await paste({ 'text/plain': 'a\tb' });
      assert.match(await status(view), /past the end of the sheet/);
      assert.equal(host.edits.length, 3);
    });
  });

  test('E6b paste is not handled while an editor is open (the textarea pastes natively) and not on a read-only grid', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B1');
      await view.page.keyboard.press('F2');
      const prevented = await view.eval(() => {
        const dt = new DataTransfer();
        dt.setData('text/plain', 'zz');
        const ev = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
        document.querySelector('.fv-cellbox').dispatchEvent(ev);
        return ev.defaultPrevented;
      });
      assert.equal(prevented, false, 'the box keeps its own paste');
      await view.page.keyboard.press('Escape');
      assert.equal(host.edits.length, 0);
      const ro = book({ editState: { editable: false, reason: 'x' } });
      await h.use(ro, {}, async (rv) => {
        await rv.eval(() => {
          const dt = new DataTransfer();
          dt.setData('text/plain', 'zz');
          document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
        });
        await rv.settle();
        assert.equal(ro.edits.length, 0);
      });
    });
  });

  // ----- overlay, queue, results -----

  test('E7 overlay: shown at once (also in an empty cell), replaced by the host rows on accept, dropped with the reason on reject', (h) => {
    const host = book();
    host.hold = true;
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E5');
      await p.keyboard.type('hello');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.equal(await cellText(view, 'hello'), true, 'the typed text is drawn before the host answers (an empty cell too)');
      await host.release(view);
      assert.equal(await cellText(view, 'hello'), true, 'then the host row shows it');
      assert.equal(await view.eval(() => document.querySelector('.fv-status-msg').textContent), '');

      // reject: the old content comes back, the reason is in the status line (never the error page)
      host.reject = { code: 'blocked', reason: 'This cell is part of an array formula (K5:K7). Edit it in Excel.' };
      host.hold = true;
      await g.goto('B1');
      await p.keyboard.type('NO');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(await cellText(view, 'NO'), true);
      host.hold = false;
      await host.release(view);
      // the held edit was answered with the rejection
      assert.equal(await cellText(view, 'NO'), false, 'overlay dropped');
      assert.equal(await cellText(view, 'Short'), true);
      assert.match(await status(view), /array formula/);
      assert.equal(await view.eval(() => !!document.querySelector('.fv-error')), false);
      // review fix round 1: a refused editor text opens the editor again on its cell (Excel keeps the editor open)
      assert.equal((await g.cellBox())?.value, 'NO', 'the editor is open again with the refused text');
      assert.equal(await g.nameBox(), 'B1');
      await p.keyboard.press('Escape');
    });
  });

  test('E8 send queue: one in flight, the next is re-based on the accepted version, a foreign patch drops queued edits', (h) => {
    const host = book();
    host.hold = true;
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const typeIn = async (ref, text) => {
        await g.goto(ref);
        await p.keyboard.type(text);
        await p.keyboard.press('Enter');
        await view.settle();
      };
      await typeIn('E1', 'A');
      await typeIn('E2', 'B');
      assert.equal(host.edits.length, 1, 'the second edit waits');
      assert.equal(host.edits[0].baseVersion, 1);
      host.hold = false;
      await host.release(view);
      assert.equal(host.edits.length, 2, 'sent after the editResult');
      assert.equal(host.edits[1].baseVersion, host.version - 1, 're-based on the version the first edit produced');

      // a patch from another source while an edit is in flight: the queued edit is never sent
      host.hold = true;
      await typeIn('E3', 'C');
      await typeIn('E4', 'D');
      assert.equal(host.edits.length, 3);
      await view.push({ type: 'patch', version: host.version + 5, patches: [] });
      await view.settle();
      assert.match(await status(view), /The file changed; your last change was not applied/);
      assert.equal(await cellText(view, 'D'), false, 'the queued edit left no overlay');
      host.hold = false;
      await host.release(view);
      assert.equal(host.edits.length, 3, 'the dropped edit was not sent');
    });
  });

  test('E9 D-067: Ctrl+Z / Ctrl+Y while an edit is in flight are held back and sent as one command after the editResult, in order', (h) => {
    const host = book();
    host.hold = true;
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const ctrl = async (k, shift) => {
        await p.keyboard.down('Control');
        if (shift) await p.keyboard.down('Shift');
        await p.keyboard.press(k);
        if (shift) await p.keyboard.up('Shift');
        await p.keyboard.up('Control');
        await view.settle();
      };
      await g.goto('E1');
      await view.eval(() => (window.__fwd = []));
      await p.keyboard.type('A');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1);
      await ctrl('KeyZ');
      await ctrl('KeyY');
      await ctrl('KeyZ', true);
      assert.deepEqual(host.commands, [], 'nothing sent while the edit is in flight');
      assert.equal((await view.eval(() => window.__fwd)).filter((k) => /^Ctrl\+(Shift\+)?[zy]$/i.test(k)).length, 0, 'the keys did not reach VS Code');
      host.hold = false;
      await host.release(view);
      assert.deepEqual(host.commands.map((c) => c.id), ['undo', 'redo', 'redo'], 'Ctrl+Z undo, Ctrl+Y redo, Ctrl+Shift+Z redo');

      // nothing in flight: the keys go to VS Code (document undo)
      await ctrl('KeyZ');
      assert.equal((await view.eval(() => window.__fwd)).filter((k) => /^Ctrl\+z$/i.test(k)).length, 1);
      assert.equal(host.commands.length, 3);

      // an edit queued behind the one in flight: the command comes after BOTH
      host.hold = true;
      await g.goto('E2');
      await p.keyboard.type('B');
      await p.keyboard.press('Enter');
      await g.goto('E3');
      await p.keyboard.type('C');
      await p.keyboard.press('Enter');
      await view.settle();
      await ctrl('KeyZ');
      host.hold = false;
      host.commands.length = 0;
      const sentBefore = host.edits.length;
      await host.release(view);
      assert.equal(host.edits.length, sentBefore + 1, 'the queued edit was sent');
      assert.deepEqual(host.commands.map((c) => c.id), ['undo']);
      // Ctrl+S in flight: held back too (the file must get the edits first)
      host.hold = true;
      await g.goto('E4');
      await p.keyboard.type('D');
      await p.keyboard.press('Enter');
      await view.settle();
      host.commands.length = 0;
      await ctrl('KeyS');
      assert.deepEqual(host.commands, []);
      assert.equal((await view.eval(() => window.__fwd)).filter((k) => /^Ctrl\+s$/i.test(k)).length, 0);
      host.hold = false;
      await host.release(view);
      assert.deepEqual(host.commands.map((c) => c.id), ['save']);
    });
  });

  test('E9b Ctrl+Z / Ctrl+Y inside the editor box and the formula bar: not default-prevented (the browser undoes the typing) and not forwarded to VS Code', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      // (headless Chrome has no editing command for Ctrl+Z: the test checks what the handler does with the key. D-026: the
      // key is NOT default-prevented, so the browser can undo the typing, and it does not bubble to VS Code.)
      await view.eval(() => {
        window.__fwd = [];
        window.__keys = [];
        window.addEventListener('keydown', (e) => /^[zy]$/i.test(e.key) && e.ctrlKey && window.__keys.push(e), true);
      });
      const ctrlKey = async (k) => {
        await p.keyboard.down('Control');
        await p.keyboard.press(k);
        await p.keyboard.up('Control');
      };
      const seen = () => view.eval(() => ({ forwarded: window.__fwd.filter((k) => /^Ctrl\+[zy]$/i.test(k)).length, prevented: window.__keys.map((e) => e.defaultPrevented) }));
      await g.goto('B1');
      await p.keyboard.press('F2');
      await p.keyboard.type('abc');
      await ctrlKey('KeyZ');
      await ctrlKey('KeyY');
      assert.deepEqual(await seen(), { forwarded: 0, prevented: [false, false] }, 'editor box');
      assert.equal(host.commands.length, 0);
      await p.keyboard.press('Escape');
      // the formula bar
      await p.click('.fv-formula');
      await p.keyboard.type('xyz');
      await ctrlKey('KeyZ');
      assert.deepEqual(await seen(), { forwarded: 0, prevented: [false, false, false] }, 'formula bar');
      await p.keyboard.press('Escape');
      assert.equal(host.edits.length, 0);
    });
  });

  // ----- close paths -----

  test('E10 every close path of the editor: commit on blur / sheet switch / hidden, cancel on Esc / invalidate / not editable, keep on scroll / zoom / resize / theme', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const open = async (ref, text) => {
        await g.goto(ref);
        await p.keyboard.press('F2');
        await p.keyboard.type(text);
      };
      // commit: a click on another cell (blur)
      await open('E1', 'a');
      await g.clickCell(7, 0);
      assert.equal(await g.cellBox(), null);
      assert.equal(host.edits.length, 1, 'blur commits');
      assert.deepEqual(host.edits[0].intent.cells[0], [0, 4, 'a']);
      // commit: sheet switch
      await open('E2', 'b');
      await g.clickTab('Other');
      assert.equal(host.edits.length, 2, 'a sheet switch commits');
      assert.deepEqual(host.edits[1].intent.cells[0], [1, 4, 'b']);
      await g.clickTab('Data');
      // commit: the panel is hidden
      await open('E3', 'c');
      await view.eval(() => {
        Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => 'hidden' });
        document.dispatchEvent(new Event('visibilitychange'));
      });
      await view.settle();
      assert.equal(host.edits.length, 3, 'a hidden panel commits');
      await view.eval(() => {
        Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => 'visible' });
      });
      // cancel: Esc, invalidate, not editable
      await open('E4', 'd');
      await p.keyboard.press('Escape');
      assert.equal(host.edits.length, 3);
      await open('E4', 'd');
      await view.push({ type: 'invalidate', version: host.version + 1 });
      await view.settle();
      assert.equal(await g.cellBox(), null, 'invalidate closes the editor');
      assert.equal(host.edits.length, 3, 'and cancels it');
      await open('E4', 'e');
      host.editState = { editable: false, reason: 'Editing is turned off (setting fileStudio.editing.enabled).' };
      await view.push({ type: 'editState', state: host.editState });
      await p.waitForFunction(() => /Editing is turned off/.test(document.querySelector('.fv-status-msg').textContent), { timeout: 3000 });
      assert.equal(await g.cellBox(), null);
      assert.equal(host.edits.length, 3, 'not editable cancels');
      assert.equal(await mode(view), 'Read-only');
      assert.equal(await view.eval(() => document.querySelector('.fv-formula').getAttribute('aria-readonly')), 'true');
      assert.equal(await view.eval(() => document.querySelector('.fv-formula').hasAttribute('contenteditable')), false);
      host.editState = { editable: true, kinds: ['setCells', 'clear'] };
      await view.push({ type: 'editState', state: host.editState });
      await view.settle();
      assert.equal(await mode(view), 'Ready');

      // keep: scroll, resize, Ctrl+wheel zoom, theme change: the editor stays open with its text
      await open('E4', 'f');
      await view.eval(() => (document.querySelector('.fv-scroller').scrollTop = 40));
      await sleep(120);
      assert.equal((await g.cellBox())?.value, 'f', 'scroll keeps the editor');
      await p.setViewport({ width: 1100, height: 700, deviceScaleFactor: 1 });
      await sleep(250);
      assert.equal((await g.cellBox())?.value, 'f', 'resize keeps the editor');
      await view.eval(() => (document.querySelector('.fv-scroller').scrollTop = 0));
      await sleep(100);
      const zoomBefore = await view.eval(() => document.querySelector('.fv-zoom-label').textContent);
      const rect = await view.eval(() => {
        const r = document.querySelector('.fv-cellbox').getBoundingClientRect();
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
      });
      await p.mouse.move(rect.x, rect.y);
      await p.keyboard.down('Control');
      await p.mouse.wheel({ deltaY: -100 });
      await p.keyboard.up('Control');
      await sleep(250);
      assert.notEqual(await view.eval(() => document.querySelector('.fv-zoom-label').textContent), zoomBefore, 'the wheel zoomed');
      assert.equal((await g.cellBox())?.value, 'f', 'zoom keeps the editor');
      await view.eval(() => {
        document.body.classList.remove('vscode-light');
        document.body.classList.add('vscode-dark');
        document.body.setAttribute('data-vscode-theme-kind', 'vscode-dark');
      });
      await sleep(250);
      assert.equal((await g.cellBox())?.value, 'f', 'a theme change keeps the editor');
      await p.keyboard.press('Escape');
      assert.equal(host.edits.length, 3);
    });
  });

  test('E11 IME: the first key opens the box without losing the composition; Enter inside a composition never commits', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E1');
      // the first key of an IME composition: keyCode 229, key "Process"; the grid does NOT take the key (no preventDefault)
      const first = await view.eval(() => {
        const sc = document.querySelector('.fv-scroller');
        const ev = new KeyboardEvent('keydown', { key: 'Process', keyCode: 229, bubbles: true, cancelable: true });
        sc.dispatchEvent(ev);
        const b = document.querySelector('.fv-cellbox');
        return { prevented: ev.defaultPrevented, open: !!b, focused: document.activeElement === b, value: b && b.value, ro: b && b.getAttribute('aria-readonly') };
      });
      assert.deepEqual(first, { prevented: false, open: true, focused: true, value: '', ro: 'false' });
      // the composition runs in the box: Enter during it (isComposing) must not commit
      const during = await view.eval(() => {
        const b = document.querySelector('.fv-cellbox');
        b.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        b.value = 'に';
        b.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: 'に', isComposing: true }));
        const ev = new KeyboardEvent('keydown', { key: 'Enter', keyCode: 229, isComposing: true, bubbles: true, cancelable: true });
        b.dispatchEvent(ev);
        return { open: !!document.querySelector('.fv-cellbox'), prevented: ev.defaultPrevented };
      });
      assert.deepEqual(during, { open: true, prevented: false });
      assert.equal(host.edits.length, 0);
      // after compositionend the text stays; now Enter commits it
      await view.eval(() => {
        const b = document.querySelector('.fv-cellbox');
        b.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'に' }));
      });
      assert.equal(await g.formula(), 'に');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.equal(host.edits[0].intent.cells[0][2], 'に');

      // a dead key opens the box too
      await g.goto('E3');
      const dead = await view.eval(() => {
        const ev = new KeyboardEvent('keydown', { key: 'Dead', keyCode: 229, bubbles: true, cancelable: true });
        document.querySelector('.fv-scroller').dispatchEvent(ev);
        return { prevented: ev.defaultPrevented, open: !!document.querySelector('.fv-cellbox') };
      });
      assert.deepEqual(dead, { prevented: false, open: true });
      await p.keyboard.press('Escape');
    });
  });

  test('E11b a real IME sequence through the DevTools protocol: first key, composition, commit', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E1');
      const client = await p.createCDPSession();
      await client.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', windowsVirtualKeyCode: 229, key: 'Process', code: 'KeyN' });
      await client.send('Input.imeSetComposition', { text: 'に', selectionStart: 1, selectionEnd: 1 });
      await client.send('Input.insertText', { text: '日本' });
      await client.send('Input.dispatchKeyEvent', { type: 'keyUp', windowsVirtualKeyCode: 78, key: 'n', code: 'KeyN' });
      const b = await g.cellBox();
      assert.ok(b, 'the editor is open');
      assert.equal(b.ro, 'false');
      assert.equal(b.value, '日本', 'the first key of the composition was not lost');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.equal(host.edits[0].intent.cells[0][2], '日本');
      await client.detach();
    });
  });

  // ----- formula bar -----

  test('E12 formula bar: editable, a second editor of the same cell (focus opens it, typing syncs both ways); Enter / Tab / Esc / blur; one edit only', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const attrs = await view.eval(() => {
        const f = document.querySelector('.fv-formula');
        return [f.getAttribute('contenteditable'), f.getAttribute('aria-readonly')];
      });
      assert.deepEqual(attrs, ['plaintext-only', 'false']);
      await g.goto('B1');
      await p.click('.fv-formula');
      const b = await g.cellBox();
      assert.ok(b, 'the editor box opened over the cell');
      assert.equal(b.focused, false, 'the formula bar keeps the focus');
      await p.keyboard.type('!');
      assert.equal((await g.cellBox()).value, 'Short!', 'formula bar to box');
      // box to formula bar: click into the box, type; the formula bar follows
      await view.eval(() => document.querySelector('.fv-cellbox').focus());
      await p.keyboard.type('?');
      assert.equal(await g.formula(), 'Short!?', 'box to formula bar');
      assert.equal((await g.cellBox()).value, 'Short!?');
      await p.keyboard.press('Escape');
      assert.equal(await g.cellBox(), null);
      assert.equal(await g.formula(), 'Short');
      assert.equal(host.edits.length, 0);

      // Enter in the formula bar: one edit, then the cell below
      await p.click('.fv-formula');
      await p.keyboard.type('+');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1, 'exactly one edit');
      assert.equal(host.edits[0].intent.cells[0][2], 'Short+');
      assert.equal(await g.nameBox(), 'B2');
      assert.equal(await g.cellBox(), null);
      assert.equal(await view.eval(() => document.activeElement === document.querySelector('.fv-scroller')), true);

      // blur: a click on another cell commits (the cell that was being edited)
      await g.goto('C1');
      await p.click('.fv-formula');
      await p.keyboard.type('X');
      await g.clickCell(3, 3);
      assert.equal(host.edits.length, 2);
      assert.deepEqual(host.edits[1].intent.cells[0], [0, 2, 'oneX']);
      assert.equal(await g.nameBox(), 'D4');
      assert.equal(await g.formula(), '', 'the bar shows the new active cell');

      // Tab moves right; Ctrl+S commits with then:save; Alt+Enter is a line break
      await g.goto('C2');
      await p.click('.fv-formula');
      await p.keyboard.down('Alt');
      await p.keyboard.press('Enter');
      await p.keyboard.up('Alt');
      await p.keyboard.type('z');
      assert.equal((await g.cellBox()).value, 'two\nz');
      await p.keyboard.down('Control');
      await p.keyboard.press('KeyS');
      await p.keyboard.up('Control');
      await view.settle();
      assert.equal(host.edits.length, 3);
      assert.equal(host.edits[2].then, 'save');
      await g.goto('C3');
      await p.click('.fv-formula');
      await p.keyboard.type('t');
      await p.keyboard.press('Tab');
      await view.settle();
      assert.equal(await g.nameBox(), 'D3');
      assert.equal(host.edits.length, 4);
    });
  });

  test('E12b the formula bar of a read-only grid is read-only (0.1.2): no editor, Enter / Esc go back to the grid', (h) => {
    const ro = book({ editState: { editable: false, reason: 'This file is read-only.' } });
    return h.use(ro, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('B1');
      await p.click('.fv-formula');
      assert.equal(await g.cellBox(), null);
      assert.equal(await view.eval(() => document.querySelector('.fv-formula').hasAttribute('contenteditable')), false);
      await p.keyboard.press('Enter');
      assert.equal(await view.eval(() => document.activeElement === document.querySelector('.fv-scroller')), true);
      assert.equal(await mode(view), 'Read-only');
      assert.match(await view.eval(() => document.querySelector('.fv-status-mode').title), /This file is read-only/);
    });
  });

  // ----- model cache, tabs, status, banner -----

  test('E13 model cache: host rows replace [start, end) (an emptied row too), stale rows are dropped and asked again, tabs and requests carry ids', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      // ids on tabs and requests
      const tabs = await view.eval(() => [...document.querySelectorAll('.fv-tab')].map((t) => t.dataset.id));
      assert.deepEqual(tabs.length, 3);
      assert.ok(tabs.every((id) => /^s\d+$/.test(id)), `tab ids ${tabs}`);
      assert.ok(host.requests.length > 0 && host.requests.every((r) => /^s\d+$/.test(r.sheetId)), 'getRows carries sheetId');
      const idData = tabs[0];
      assert.equal(await cellText(view, 'Short'), true);

      // a patch from another panel: rows [0,3) replaced; row 2 (index 1 'B2/C2') becomes empty because it is not in `rows`
      const model = await host.model();
      const S = H.renderer('sheet');
      const edit = S.planEdit(model, { kind: 'clear', sheetId: idData, ranges: [{ r0: 1, c0: 0, r1: 1, c1: 7 }], what: 'contents' });
      S.applyEdit(model, edit);
      host.version++;
      const { rows, styles } = model.getRows(0, 0, 3);
      await view.push({ type: 'patch', version: host.version, patches: [{ sheetId: idData, start: 0, end: 3, rows, styles }] });
      await view.settle();
      assert.equal(await cellText(view, 'two'), false, 'the emptied row is gone');
      assert.equal(await cellText(view, 'one'), true);
      assert.equal(await cellText(view, '007'), false);

      // stale rows (an older version) are dropped: the block is asked for again, with the new version
      const before = host.requests.length;
      await view.push({ type: 'patch', version: host.version + 1, patches: [] });
      await view.eval(() => {
        // a late reply for a request made before: its version is lower than the grid's
        window.postMessage({ type: 'rows', reqId: 9999, sheet: 0, sheetId: 's1', version: 1, start: 0, end: 256, rows: [], styles: [] }, '*');
      });
      await view.settle();
      assert.equal(await cellText(view, 'one'), true, 'a stale reply does not empty the grid');
      assert.ok(host.requests.length >= before);

      // a patch with new sheet meta (the used range grew): replaced by id, the grid keeps working
      const meta = model.getMeta();
      const grown = { ...meta.sheets[0], rowCount: meta.sheets[0].rowCount + 5 };
      await view.push({ type: 'patch', version: host.version + 2, patches: [], sheets: [grown] });
      await view.settle();
      assert.equal(await cellText(view, 'one'), true);
      await g.goto('A9');
      assert.equal(await g.nameBox(), 'A9');
    });
  });

  test('E13b the active sheet and view state are kept by sheet id: reload restores the sheet', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.clickTab('Other');
      await sleep(600); // the state save is debounced
      const key = await view.eval(() => {
        const raw = sessionStorage.getItem('fs-state:edit.xlsx');
        const s = raw && JSON.parse(raw);
        return s && s.grid && s.grid.activeSheet;
      });
      assert.match(key, /^id:s\d+$/, `saved by id: ${key}`);
      await view.reload();
      assert.equal(await g.activeTab(), 'Other');
    });
  });

  test('E14 status bar and banner: Ready / Read-only with the reason, "Edit anyway" posts editAnyway, no Dismiss while it is offered', (h) => {
    const host = book({ editState: { editable: false, code: 'lossy', reason: 'This workbook has content FileStudio cannot keep when saving.', canEditAnyway: { title: 'x', details: ['a chart'] } }, banner: 'This workbook has a chart that would be lost.' });
    return h.use(host, {}, async (view) => {
      const p = view.page;
      assert.equal(await mode(view), 'Read-only');
      assert.match(await view.eval(() => document.querySelector('.fv-status-mode').title), /cannot keep when saving/);
      assert.equal(await view.eval(() => !!document.querySelector('.fv-banner .fv-banner-close')), false, 'no Dismiss next to the button');
      assert.equal(await view.eval(() => document.querySelector('.fv-banner-action').textContent), 'Edit anyway');
      await p.click('.fv-banner-action');
      await view.settle();
      assert.equal(host.anyway, 1, 'editAnyway message sent (the host asks the user; the message is no consent)');
      assert.equal(await mode(view), 'Ready');
      assert.equal(await view.eval(() => !!document.querySelector('.fv-banner-action')), false, 'the button goes when the host allows editing');
      assert.equal(await view.eval(() => document.querySelector('.fv-banner .fv-banner-text').textContent), 'This workbook has a chart that would be lost.');
      assert.equal(await view.eval(() => !!document.querySelector('.fv-banner .fv-banner-close')), true, 'the banner can be dismissed again');
      // the same text with a changed action is rebuilt (updateBanner does not return early)
      await view.push({ type: 'editState', state: { editable: false, canEditAnyway: { title: 'x', details: [] } } });
      await view.settle();
      assert.equal(await view.eval(() => !!document.querySelector('.fv-banner-action')), true);
      // typing opens nothing while it is read-only
      await H.grid(view).goto('B1');
      await p.keyboard.type('q');
      assert.equal(await H.grid(view).cellBox(), null);
      assert.equal(host.edits.length, 0);
    });
  });

  test('E15 copy keeps working: Ctrl+C of a range gives the same TSV and HTML as a read-only grid; a pending overlay is what is copied', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B1:C3');
      const editable = await g.copy();
      let readOnly;
      await h.use(book({ editState: { editable: false } }), {}, async (rv) => {
        const rg = H.grid(rv);
        await rg.goto('B1:C3');
        readOnly = await rg.copy();
      });
      assert.deepEqual(editable, readOnly);
      assert.match(editable.text, /Short\tone/);
      assert.match(editable.html || '', /<table>/);

      // with the editor closed and an overlay value pending, the copy has the overlay text (what the grid shows)
      host.hold = true;
      await g.goto('C1');
      await view.page.keyboard.type('pending');
      await view.page.keyboard.press('Enter');
      await view.settle();
      await g.goto('B1:C3');
      const pending = await g.copy();
      assert.equal(pending.text.split('\r\n')[0], 'Short\tpending');
      assert.match(pending.html || '', /pending/);
      await g.goto('C1');
      const single = await g.copy();
      assert.equal(single.text, 'pending');
      host.hold = false;
      await host.release(view);
    });
  });

  test('E16 high-contrast and dark themes: the editor frame and the banner button exist and the editor works', async (h) => {
    for (const theme of ['hc', 'dark']) {
      const host = book({ editState: { editable: true, kinds: ['setCells', 'clear'] } });
      await h.use(host, { theme }, async (view) => {
        const g = H.grid(view);
        await g.goto('B1');
        await view.page.keyboard.press('F2');
        const style = await view.eval(() => {
          const b = document.querySelector('.fv-cellbox');
          const cs = getComputedStyle(b);
          return { outline: cs.outlineStyle, width: parseFloat(cs.outlineWidth), caret: cs.caretColor };
        });
        assert.equal(style.outline, 'solid', theme);
        assert.ok(style.width >= 2, theme);
        await view.page.keyboard.type('!');
        await view.page.keyboard.press('Enter');
        await view.settle();
        assert.equal(host.edits.length, 1, theme);
      });
    }
  });

  test('E17 a large paste (20,000 cells): one edit, answered, grid responsive; limits are checked before sending', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('J1');
      const tsv = Array.from({ length: 400 }, (_, r) => Array.from({ length: 50 }, (_, c) => `${r}x${c}`).join('\t')).join('\n');
      const t0 = Date.now();
      await view.eval((tsv) => {
        const dt = new DataTransfer();
        dt.setData('text/plain', tsv);
        document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      }, tsv);
      await view.settle(15000);
      assert.equal(host.edits.length, 1);
      assert.equal(host.edits[0].intent.cells.length, 20000);
      assert.ok(Date.now() - t0 < 15000);
      assert.equal(await g.nameBox(), 'J1:BG400');
      // a block over the sheet's end is refused in the webview
      await g.goto('XFD1048576');
      await view.eval(() => {
        const dt = new DataTransfer();
        dt.setData('text/plain', 'a\nb');
        document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      });
      await view.settle();
      assert.equal(host.edits.length, 1);
    });
  });
});

H.suite('csv: the same editor (format-agnostic, P3 reuses it)', ({ test }) => {
  test('E18 a csv panel with editing on: the editor works, no apostrophe hint and no formula look for text such as =1+1', (h) => {
    const inner = H.csvHost('c.csv', '007,=1+1\n1,2\n');
    const edits = [];
    const host = {
      kind: 'csv',
      fileName: 'c.csv',
      log: [],
      async receive(msg) {
        if (msg.type === 'edit') {
          edits.push(msg);
          return [{ type: 'editResult', editId: msg.editId, accepted: true, version: 2 }];
        }
        const replies = await inner.receive(msg);
        return replies.map((m) => {
          if (m.type === 'init') return { ...m, readOnly: false, version: 1, editState: { editable: true, kinds: ['setCells', 'clear'] }, meta: { ...m.meta, sheets: m.meta.sheets.map((sh) => ({ ...sh, id: 'csv' })) } };
          if (m.type === 'rows') return { ...m, version: 1, sheetId: 'csv' };
          return m;
        });
      },
    };
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      await p.keyboard.press('F2');
      assert.equal((await g.cellBox()).value, '007', 'the field text, no apostrophe');
      await p.keyboard.press('Escape');
      await g.goto('B1');
      assert.equal(await g.formula(), '=1+1');
      assert.equal(await view.eval(() => document.querySelector('.fv-formula').classList.contains('fv-formula-f')), false, 'csv text is never a formula');
      await p.keyboard.type('x');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(edits.length, 1);
      assert.deepEqual(edits[0].intent, { kind: 'setCells', sheetId: 'csv', cells: [[0, 1, 'x']] });
      assert.equal(await mode(view), 'Ready');
    });
  });

  // ----- review fixes, round 1 -----

  /** A synthetic paste on the grid. */
  const pasteData = async (view, data) => {
    await view.eval((data) => {
      const dt = new DataTransfer();
      for (const [type, value] of Object.entries(data)) dt.setData(type, value);
      document.querySelector('.fv-scroller').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    }, data);
    await view.settle();
  };

  test('X1 HTML paste keeps text markers (x:str, mso-number-format "\\@" inline or by class) and exact x:num values', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      await H.grid(view).goto('F1');
      const html =
        '<html><head><style>.xl65{mso-number-format:"\\@";}</style></head><body><table><tr>' +
        '<td class=xl65>007</td><td x:str>=1+1</td><td style="mso-number-format:\'\\@\'">true</td><td x:str> lead</td>' +
        '<td x:num="3.14159">3.14</td><td x:num="0.05123">5.1%</td><td x:num="45296">1/5/2024</td><td x:num>9</td><td>plain</td>' +
        '</tr></table></body></html>';
      await pasteData(view, { 'text/html': html, 'text/plain': 'x' });
      assert.equal(host.edits.length, 1);
      assert.deepEqual(host.edits[0].intent.cells.map((c) => c[2]), ["'007", "'=1+1", "'true", ' lead', '3.14159', '5.123%', '1/5/2024', '9', 'plain']);
      const model = await host.model();
      const row = model.getRows(0, 0, 1).rows[0];
      const at = (c) => row.cells.find((x) => x.c === c);
      assert.deepEqual([at(5).t, at(5).v], ['s', '007'], 'text stays text');
      assert.deepEqual([at(6).t, at(6).v, at(6).f], ['s', '=1+1', undefined], 'no live formula');
      assert.deepEqual([at(7).t, at(7).v], ['s', 'true']);
      assert.equal(at(9).v, 3.14159, 'the exact value, not the display text');
    });
  });

  test('X2 copy then paste inside FileStudio pastes the values (text that reads as a value stays text)', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B2:B4');
      const copied = await g.copy();
      assert.equal(copied.text, '007\r\n=1+1\r\n12.5');
      await g.goto('H1');
      await pasteData(view, { 'text/plain': copied.text }); // the plain text only: the copy is matched by its text
      assert.deepEqual(host.edits[0].intent.cells.map((c) => c[2]), ["'007", "'=1+1", '12.5']);
      await g.goto('H5');
      await pasteData(view, { 'text/plain': 'other text' });
      assert.deepEqual(host.edits[1].intent.cells.map((c) => c[2]), ['other text'], 'another clipboard text is read as it is');
    });
  });

  test('X3 a refused editor text opens the editor again with the text (Excel keeps the editor open)', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E5');
      await p.keyboard.type('=A1+');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 1);
      assert.match(await status(view), /problem with this formula/);
      const box = await g.cellBox();
      assert.ok(box, 'the editor is open again');
      assert.equal(box.value, '=A1+');
      assert.equal(await g.nameBox(), 'E5');
      await p.keyboard.type('1');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.equal(host.edits.length, 2);
      assert.deepEqual(host.edits[1].intent.cells, [[4, 4, '=A1+1']]);
      assert.equal(await g.cellBox(), null);
    });
  });

  test('X4 read-only: typing and Delete say why nothing happens; "Edit anyway" that turns editing on is announced', (h) => {
    const host = book({ editState: { editable: false, code: 'lossy', reason: 'This workbook has content FileStudio cannot keep when saving.', canEditAnyway: { title: 'x', details: ['a chart'] } }, banner: 'This workbook has a chart that would be lost.' });
    return h.use(host, {}, async (view) => {
      const p = view.page;
      await H.grid(view).goto('B1');
      await p.keyboard.type('q');
      assert.match(await status(view), /^This view is read-only: This workbook has content/);
      assert.equal(host.edits.length, 0);
      await p.click('.fv-banner-action');
      await view.settle();
      assert.equal(await mode(view), 'Ready');
      assert.equal(await status(view), 'Editing is on.');
      assert.equal(await view.eval(() => document.querySelector('.fv-status-msg').getAttribute('aria-live') || document.querySelector('.fv-status-msg').closest('[aria-live]')?.getAttribute('aria-live')), 'polite');
      assert.equal(await view.eval(() => document.querySelector('.fv-formula').getAttribute('aria-multiline')), 'true');
    });
  });

  test('X5 Ctrl+Enter commits and stays in the cell; accepted Delete and paste are announced', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('E3');
      await p.keyboard.type('abc');
      await p.keyboard.down('Control');
      await p.keyboard.press('Enter');
      await p.keyboard.up('Control');
      await view.settle();
      assert.equal(await g.cellBox(), null);
      assert.deepEqual(host.edits[0].intent.cells, [[2, 4, 'abc']]);
      assert.equal(await g.nameBox(), 'E3', 'stays in the cell');
      await p.keyboard.press('Delete');
      await view.settle();
      assert.equal(await status(view), 'Cleared.');
      await g.goto('F1');
      await pasteData(view, { 'text/plain': 'a\tb' });
      assert.equal(await status(view), 'Changed 2 cells.');
    });
  });

  test('X6 an HTML table whose spans add up to more than 1,000,000 cells is refused before it is expanded', (h) => {
    const host = book();
    return h.use(host, {}, async (view) => {
      await H.grid(view).goto('A10');
      const td = '<td rowspan=1000 colspan=1000>x</td>';
      const t0 = Date.now();
      await pasteData(view, { 'text/html': `<table><tr>${td.repeat(30)}</tr></table>`, 'text/plain': 'x' });
      assert.ok(Date.now() - t0 < 3000, `fast (${Date.now() - t0} ms)`);
      assert.match(await status(view), /too large/);
      assert.equal(host.edits.length, 0);
    });
  });
});

// keep the old fixtures module referenced (its builders are used by the other webview suites)
void F;
