// Regression tests for fixed 0.1.x grid bugs around copy, the read-only cell text box, header clicks and hidden
// sheets. These paths turn into 0.2.0 editing (cell editor, paste, cut, sheet hide/unhide): they must keep working.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const cells = () => H.sheetHost('cells.xlsx', F.cells);

H.suite('sheet: copy, cell text box, headers, hidden sheets (0.1.x regressions)', ({ test }) => {
  test('U1 headers show select-arrow cursors, not resize cursors (also mirrored on a right-to-left sheet)', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const cursors = () =>
        view.eval(() => {
          const one = (sel) => getComputedStyle([...document.querySelectorAll(sel)].find((e) => e.offsetParent !== null && e.style.display !== 'none')).cursor;
          return { ch: one('.fv-ch'), rh: one('.fv-rh'), corner: one('.fv-corner') };
        });
      const cur = await cursors();
      assert.match(cur.ch, /^url\("data:image\/svg\+xml,.+"\) 7 12, default$/);
      assert.match(cur.rh, /^url\("data:image\/svg\+xml,.+"\) 12 7, default$/);
      assert.equal(cur.corner, 'default');
      await g.clickTab('RTL');
      assert.match((await cursors()).rh, /^url\("data:image\/svg\+xml,.+"\) 2 7, default$/);
    }));

  test('U2 Ctrl+C on one cell / one merged area copies its text only (no quotes, no tabs, no HTML table)', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const copy = async () => {
        const c = await g.copy();
        return [c.text, c.html ?? 'no html'];
      };
      await g.goto('C6');
      assert.deepEqual(await copy(), ['Line one\nLine two\nLine three', 'no html']);
      await g.goto('C5');
      assert.deepEqual(await copy(), ['a\tb', 'no html']);
      await g.goto('B2');
      assert.deepEqual(await copy(), ['Merged title B2:C2', 'no html']);
      await g.goto('B6:C6');
      const several = await g.copy();
      assert.equal(several.text, 'Short\t"Line one\nLine two\nLine three"');
      assert.match(several.html || '', /<table>/);
    }));

  test('U3 double-click opens a read-only cell text box with the display text; typing changes and sends nothing', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('C6');
      await g.clickCell(5, 2, [], { count: 2 });
      await H.sleep(100);
      const posted = (await view.posted()).length;
      const b = await g.cellBox();
      assert.ok(b, 'cell text box open');
      assert.equal(b.value, 'Line one\nLine two\nLine three');
      assert.deepEqual([b.start, b.end, b.focused, b.ro, b.label], [0, b.value.length, true, 'true', 'C6 cell text, read-only']);
      assert.equal(await g.formula(), 'Line one\nLine two\nLine three');
      await p.keyboard.type('typed');
      await p.keyboard.press('Backspace');
      await p.keyboard.press('Delete');
      await p.keyboard.down('Control');
      await p.keyboard.press('KeyV');
      await p.keyboard.press('KeyX');
      await p.keyboard.up('Control');
      assert.equal((await g.cellBox()).value, 'Line one\nLine two\nLine three');
      assert.equal((await view.posted()).length, posted, 'nothing sent to the host');
      await p.keyboard.press('Escape');
      await H.sleep(80);
      assert.equal(await g.cellBox(), null);
      assert.equal(await g.nameBox(), 'C6');
      assert.match(await g.focus(), /^fv-scroller/);
      // A hyperlink cell: the box opens, no navigation.
      await view.clearPosted();
      await g.clickCell(7, 4, [], { count: 2 });
      await H.sleep(100);
      assert.equal((await g.cellBox())?.value, 'Example site');
      assert.deepEqual(await view.posted('openLink'), []);
      await p.keyboard.press('Escape');
      // Right-to-left sheet: the box sits over the cell and its text is not mirrored.
      await g.clickTab('RTL');
      await g.clickCell(1, 1, [], { count: 2 });
      await H.sleep(100);
      const rtl = await view.eval(() => {
        const el = document.querySelector('.fv-cellbox');
        const r = el.getBoundingClientRect();
        const ch = window.__g.hdr('col', 'B').getBoundingClientRect();
        return { value: el.value, transform: getComputedStyle(el).transform, dl: Math.abs(r.left - ch.left), dr: Math.abs(r.right - ch.right) };
      });
      assert.equal(rtl.value, 'مرحبا بالعالم');
      assert.equal(rtl.transform, 'none');
      assert.ok(rtl.dl <= 1 && rtl.dr <= 1, JSON.stringify(rtl));
    }));

  test('R2-F7 F2 opens the cell text box with the caret at the end; Escape gives the focus back to the grid', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('C6');
      await g.press('F2');
      const b = await g.cellBox();
      assert.ok(b && b.focused, 'box focused');
      assert.equal(b.value, 'Line one\nLine two\nLine three');
      assert.deepEqual([b.start, b.end], [b.value.length, b.value.length]);
      await g.press('Escape');
      assert.equal(await g.cellBox(), null);
      assert.match(await g.focus(), /^fv-scroller/);
    }));

  test('R2-F1 / R2-F8 Ctrl+C and the status bar leave hidden rows out; hidden columns are not copied', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('B30:B32');
      assert.equal((await g.copy()).text, 'r30\r\nr32');
      assert.equal(await g.stats(), 'Count: 2');
      await g.goto('F30:H30');
      assert.equal((await g.copy()).text, 'f30\th30');
      assert.equal(await g.stats(), 'Count: 3', 'a hidden column still counts (Excel)');
    }));

  test('R2-F5 text cells are marked as text in the copied HTML (Excel pastes 007 as 007), numbers are not', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A10:A13');
      const c = await g.copy();
      assert.equal(c.text, '007\r\n1-2\r\n=1+1\r\n3/4');
      const tds = (c.html || '').match(/<td[^>]*>/g) || [];
      assert.equal(tds.length, 4);
      assert.ok(tds.every((t) => t === `<td style="mso-number-format:'\\@'">`), c.html);
      await g.goto('A13:A14');
      const n = await g.copy();
      assert.match(n.html || '', /<td>12\.5<\/td>/);
    }));

  test('R2-F6 Ctrl+A selects only the formula-bar text in the formula bar, and no page text on a tab or button', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const pageSel = () =>
        view.eval(() => {
          const s = window.getSelection();
          const f = document.querySelector('.fv-formula');
          return { text: s.toString(), inFormula: !s.isCollapsed && f.contains(s.anchorNode) && f.contains(s.focusNode) };
        });
      await g.goto('C6');
      const ftext = await g.formula();
      await p.click('.fv-formula');
      await g.press('KeyA', ['Control']);
      const s = await pageSel();
      assert.equal(s.text, ftext);
      assert.ok(s.inFormula);
      await view.eval(() => window.getSelection().removeAllRanges());
      await p.focus('.fv-tab-active');
      await g.press('KeyA', ['Control']);
      assert.equal((await pageSel()).text, '');
      await p.focus('.fv-zoom-btn');
      await g.press('KeyA', ['Control']);
      assert.equal((await pageSel()).text, '');
      // clicking a cell ends a text selection in the formula bar
      const fb = await view.eval(() => {
        const r = document.querySelector('.fv-formula').getBoundingClientRect();
        return { x: r.left + 20, y: r.top + 8 };
      });
      await p.mouse.click(fb.x, fb.y, { count: 2 });
      assert.ok((await pageSel()).text.length > 0);
      await g.clickCell(0, 0);
      assert.equal((await pageSel()).text, '');
    }));

  test('R2-F11 the context-menu Copy (copy event on <body>) copies the selection; keys work after the menu', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const menu = (selector) =>
        view.eval((selector) => {
          const el = document.querySelector(selector);
          el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2 }));
          el.blur(); // VS Code's menu takes the focus out of the webview frame
          return document.activeElement === document.body;
        }, selector);
      const menuCopy = () =>
        view.eval(() => {
          const dt = new DataTransfer();
          document.body.dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true }));
          return { text: dt.getData('text/plain'), html: dt.getData('text/html') };
        });
      await g.goto('B6:C6');
      assert.ok(await menu('.fv-scroller'));
      const c = await menuCopy();
      assert.equal(c.text, 'Short\t"Line one\nLine two\nLine three"');
      assert.match(c.html, /<table>/);
      await view.eval(() => window.dispatchEvent(new Event('focus')));
      assert.match(await g.focus(), /^fv-scroller/);
      await menu('.fv-scroller');
      await g.press('ArrowDown');
      assert.equal(await g.nameBox(), 'B7');
      assert.match(await g.focus(), /^fv-scroller/);
      // The cell text box stays open under the menu; its Copy copies the selected text.
      await g.goto('C6');
      await g.press('F2');
      await view.eval(() => document.querySelector('.fv-cellbox').setSelectionRange(9, 17));
      assert.ok(await menu('.fv-cellbox'));
      assert.ok(await g.cellBox(), 'box still open');
      assert.deepEqual(await menuCopy(), { text: 'Line two', html: '' });
    }));

  test('R2-F12 a header click next to merges selects only that column / row (Excel)', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      for (const [label, want] of [['D', 'D:D'], ['C', 'C:C']]) {
        await g.clickHdr('col', label);
        assert.equal(await g.nameBox(), want, `header ${label}`);
      }
      await g.goto('A37');
      await g.clickHdr('row', '37');
      assert.equal(await g.nameBox(), '37:37');
    }));

  test('R2-F15 Ctrl+click inside the selection deselects that cell (Excel 365)', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('D1:D4');
      await g.clickCell(1, 3, ['Control']); // D2
      assert.equal(await g.stats(), 'Count: 3 Sum: 5 Average: 2.5 Min: 1 Max: 4');
      assert.equal(await g.nameBox(), 'D1');
      const hole = await view.eval(() => {
        const r = window.__g.cellRect(1, 3);
        return document.elementsFromPoint(r.cx, r.cy).some((e) => e.classList.contains('fv-sel-fill'));
      });
      assert.equal(hole, false, 'D2 is not shown selected');
    }));

  test('U4a a workbook whose active tab is a hidden sheet opens on that sheet, shown (Excel 16)', (h) =>
    h.use(H.sheetHost('hidden-active.xlsx', F.hiddenActive), {}, async (view) => {
      assert.deepEqual(await H.grid(view).tabs(), ['One', 'Two[h]*', 'Three']);
    }));

  test('U4b the hidden-sheets button and menu (count, very hidden label); "Hide again" removes the shown tab', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      const btn = await p.$eval('.fv-hidden-btn', (e) => [e.getAttribute('aria-label'), e.getAttribute('aria-expanded')]);
      assert.deepEqual(btn, ['Hidden sheets (2)', 'false']);
      await p.focus('.fv-hidden-btn');
      await p.keyboard.press('ArrowUp');
      await H.sleep(100);
      const items = await p.$$eval('.fv-menu-item', (els) => els.map((e) => [e.querySelector('.fv-menu-label').textContent, e.querySelector('.fv-menu-hint')?.textContent ?? '']));
      assert.deepEqual(items, [['Hid', ''], ['VeryHid', 'very hidden']]);
      assert.equal(await view.eval(() => document.activeElement.querySelector('.fv-menu-label')?.textContent), 'VeryHid');
      await p.keyboard.press('Enter');
      await view.settle();
      assert.ok((await g.tabs()).includes('VeryHid[h]*'));
      assert.match(await p.$eval('.fv-status-dims', (e) => e.textContent), /very hidden sheet$/);
      await p.click('.fv-hidden-btn');
      await H.sleep(100);
      await view.eval(() => [...document.querySelectorAll('.fv-menu-item')].find((e) => e.textContent.startsWith('Hide the shown')).click());
      await view.settle();
      assert.deepEqual(await g.tabs(), ['Copy', 'RTL*']);
      assert.ok(await p.$('.fv-hidden-btn'), 'the menu button is still there');
    }));

  test('U4c a link into a hidden sheet shows that sheet with a message (Excel refuses silently)', (h) =>
    h.use(H.sheetHost('hidden-link.xlsx', F.hiddenLink), {}, async (view) => {
      const g = H.grid(view);
      await g.clickCell(1, 0, ['Control']);
      assert.ok((await g.tabs()).includes('Two[h]*'));
      assert.equal(await g.nameBox(), 'B2');
      assert.equal(await view.eval(() => document.querySelector('.fv-status-msg').textContent), 'Showing the hidden sheet "Two" (Excel does not go to hidden sheets).');
    }));

  test('R3-024 CSV: the formula bar shows the field text (007), not the converted number', (h) =>
    h.use(H.csvHost('codes.csv', 'Code,Name\n007,Bond\n0042,Answer\n'), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A2');
      assert.equal(await g.formula(), '007');
      await g.goto('A3');
      assert.equal(await g.formula(), '0042');
    }));

  test('R3-058 CSV: Ctrl+PageDown / Ctrl+PageUp are left to VS Code (switch editor tabs)', (h) =>
    h.use(H.csvHost('keys.csv', 'a,b\n1,2\n3,4\n'), {}, async (view) => {
      const g = H.grid(view);
      await g.ready();
      const prevented = await view.eval(() => {
        const out = [];
        window.addEventListener('keydown', (e) => out.push(`${e.key}:${e.defaultPrevented}`));
        window.__keyOut = out;
        return true;
      });
      assert.ok(prevented);
      await view.page.focus('.fv-scroller');
      await g.press('PageDown', ['Control']);
      await g.press('PageUp', ['Control']);
      const out = await view.eval(() => window.__keyOut.filter((k) => k.startsWith('Page')));
      assert.deepEqual(out, ['PageDown:false', 'PageUp:false']);
    }));
});
