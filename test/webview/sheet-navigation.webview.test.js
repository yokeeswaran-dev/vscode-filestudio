// Regression tests for fixed 0.1.x grid bugs: selection, keyboard and mouse navigation, view state.
// Each test names the bug id of the 0.1.x hunts (see CHANGELOG 0.1.1 / 0.1.2). Expected values are Excel's.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const nav = () => H.sheetHost('nav.xlsx', F.nav);
const merges = () => H.sheetHost('merges.xlsx', F.merges);

H.suite('sheet: selection and navigation (0.1.x regressions)', ({ test }) => {
  test('H5-F1 Shift+Space / Ctrl+Space keep the active cell and select every row / column of the selection', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('C25');
      await g.press(' ', ['Shift']);
      const a = await g.ui();
      assert.equal(a.name, '25:25');
      assert.equal(a.active, 'C25');
      await g.goto('B3:C6'); // snaps to B2:C6 (merge B2:C3)
      await g.clickCell(9, 4, ['Control']); // + E10
      await g.press(' ', ['Shift']);
      const b = await g.ui();
      assert.equal((await g.full('row')).join(','), '2,3,4,5,6,10');
      assert.equal(b.active, 'E10');
      await g.goto('H25');
      await g.press(' ', ['Control']);
      const c = await g.ui();
      assert.equal(c.name, 'H:H');
      assert.equal(c.active, 'H25');
      await g.press('ArrowRight', ['Shift']);
      assert.equal((await g.ui()).name, 'H:I');
    }));

  test('H5-F2 Ctrl+Arrow finds data in rows that are not loaded yet and stops at the sheet edge', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A1');
      const down = [];
      for (let i = 0; i < 7; i++) {
        await g.press('ArrowDown', ['Control']);
        await H.sleep(100);
        down.push((await g.ui()).name);
      }
      assert.equal(down.join(' '), 'A10 A15 A20 A600 A1000 A1500 A1048576');
      await g.goto('K40');
      await g.press('ArrowRight', ['Control']);
      await H.sleep(100);
      assert.equal((await g.ui()).name, 'XFD40');
      await g.goto('A1500');
      await g.press('ArrowUp', ['Control']);
      await H.sleep(100);
      assert.equal((await g.ui()).name, 'A1000');
    }));

  test('H5-F3 name box / internal link past the used range select and show that cell', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A5000');
      const u = await g.ui();
      assert.equal(u.name, 'A5000');
      assert.equal(u.active, 'A5000');
      assert.ok(await view.eval(() => !!window.__g.hdr('row', '5000')), 'row 5000 is shown');
      await g.goto('A1');
      await g.clickCell(5, 6, ['Control']); // G6 -> #Data!A5000
      assert.equal((await g.ui()).name, 'A5000');
    }));

  test('H5-F4 the extension end of a selection survives a webview reload', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('D10');
      await g.press('ArrowDown', ['Shift'], 3);
      await g.press('ArrowRight', ['Shift'], 2);
      assert.equal((await g.ui()).name, 'D10:F13');
      await H.sleep(700); // the state save is debounced
      await view.reload();
      await g.press('ArrowDown', ['Shift']);
      assert.equal((await g.ui()).name, 'D10:F14');
    }));

  test('H5-F5 Tab / Shift+Enter cycle through the ranges of a multi-range selection (selection kept)', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A1:A2');
      await g.clickCell(0, 3, ['Control']); // + D1
      const seq = [];
      for (let i = 0; i < 4; i++) {
        await g.press('Tab');
        seq.push((await g.ui()).active);
      }
      await g.press('Enter', ['Shift']);
      seq.push((await g.ui()).active);
      assert.equal(seq.join(' '), 'A1 A2 D1 A1 D1');
    }));

  test('H5-F6 the first Shift+Arrow from a merged cell grows the range', (h) =>
    h.use(merges(), {}, async (view) => {
      const g = H.grid(view);
      await g.clickCell(6, 2); // C7 -> C6:C9
      await g.press('ArrowDown', ['Shift']);
      assert.equal((await g.ui()).name, 'C6:C10');
      await g.clickCell(2, 1); // B3 -> B3:D4
      await g.press('ArrowRight', ['Shift']);
      assert.equal((await g.ui()).name, 'B3:E4');
      await g.clickCell(6, 2);
      await g.press('ArrowUp', ['Shift']);
      assert.equal((await g.ui()).name, 'C5:C9');
    }));

  test("H5-F7 HYPERLINK('other.xlsx#Sheet1!A1') / [other.xlsx] links go to the host as file links", (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await view.clearPosted();
      await g.clickCell(2, 6, ['Control']); // G3
      await g.clickCell(4, 6, ['Control']); // G5
      await H.sleep(200);
      const links = (await view.posted('openLink')).map((m) => m.href);
      assert.deepEqual(links, ['other.xlsx#Sheet1!A1', 'other.xlsx#Sheet1!A1']);
    }));

  test('H5-F8 the status bar shows tiny / huge numbers like Excel (1E-12, 1E+21), not 0', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('H8:H9');
      const a = await g.stats();
      assert.match(a, /Sum: 1E-12/);
      assert.match(a, /Min: 1E-12/);
      await g.goto('H7:H8');
      assert.match(await g.stats(), /Max: 1E\+21/);
    }));

  test('H5-F9 Shift+click on a header after Ctrl+click extends only the last range', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.clickHdr('row', '5');
      await g.clickHdr('row', '12', ['Shift']);
      await g.clickHdr('row', '20', ['Control']);
      await g.clickHdr('row', '22', ['Shift']);
      assert.equal((await g.full('row')).join(','), '5,6,7,8,9,10,11,12,20,21,22');
    }));

  test('H5-F10 / H9-F4 Tab in the hidden-sheets menu closes it and focuses its button', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await view.page.click('.fv-hidden-btn');
      await H.sleep(100);
      await view.page.keyboard.press('Tab');
      await H.sleep(100);
      assert.equal(await view.eval(() => !!document.querySelector('.fv-menu')), false, 'menu closed');
      assert.match(await g.focus(), /Hidden sheets/);
    }));

  test('H5-F11 PageDown / PageUp page by whole rows and undo each other (also below frozen rows)', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      for (const [sheet, start, fr] of [['Data', 'C4', 0], ['Frozen', 'C66', 3]]) {
        await g.goto(`${sheet}!${start}`);
        const seq = [];
        for (const k of ['PageDown', 'PageDown', 'PageDown', 'PageUp', 'PageUp', 'PageUp']) {
          await g.press(k);
          const u = await g.ui();
          seq.push(u.name + ((await view.eval((fr) => window.__g.rowAligned(fr), fr)) ? '' : '(cut)'));
        }
        assert.equal(seq[5], start, `${sheet}: ${seq.join(' ')}`);
        assert.notEqual(seq[2], start, `${sheet}: ${seq.join(' ')}`);
        assert.ok(!seq.some((x) => x.includes('cut')), `${sheet}: top row cut: ${seq.join(' ')}`);
      }
    }));

  test('H5-F13 a drag from the scrolling rows into the frozen rows auto-scrolls instead of selecting frozen rows', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('Frozen!E120');
      const st0 = (await g.ui()).st;
      const xy = await g.cellXY(119, 4);
      const band = await view.eval(() => window.__g.hdr('row', '3').getBoundingClientRect().bottom);
      await p.mouse.move(xy.x, xy.y);
      await p.mouse.down();
      await p.mouse.move(xy.x, band - 10, { steps: 5 });
      await H.sleep(250);
      const mid = await view.eval(() => ({ st: Math.round(document.querySelector('.fv-scroller').scrollTop), name: document.querySelector('.fv-namebox').value }));
      await p.mouse.up();
      const top = Number(/^[A-Z]+(\d+)/.exec(mid.name)?.[1] ?? 0);
      assert.ok(mid.st < st0, `scrollTop ${st0} -> ${mid.st}`);
      assert.ok(top >= 4, `range ${mid.name} reaches into the frozen rows`);
    }));

  test('H5-F14 arrow keys through a merge keep the column / row', (h) =>
    h.use(merges(), {}, async (view) => {
      const g = H.grid(view);
      await g.clickCell(3, 0); // A4
      await g.press('ArrowRight');
      assert.equal((await g.ui()).name, 'B3'); // the merge B3:D4 (name box shows its first cell, like Excel)
      await g.press('ArrowRight');
      assert.equal((await g.ui()).name, 'E4');
      await g.clickCell(9, 2); // C10
      await g.press('ArrowUp');
      assert.equal((await g.ui()).name, 'C6'); // the merge C6:C9
      await g.press('ArrowUp');
      assert.equal((await g.ui()).name, 'C5');
      await g.goto('C1');
      await g.press('ArrowDown');
      await g.press('ArrowDown');
      await g.press('ArrowDown');
      assert.equal((await g.ui()).name, 'C5', 'C1 Down x3 across B3:D4 lands in column C');
    }));

  test('H5-F15 Ctrl+A selects the current region first (Excel CurrentRegion), then the whole sheet', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('A3');
      await g.press('a', ['Control']);
      const a = await g.ui();
      assert.equal(a.name, 'A1:H13');
      assert.equal(a.active, 'A3');
      await g.press('a', ['Control']);
      assert.ok(await view.eval(() => window.__g.allFull()), 'second Ctrl+A selects all');
      await g.goto('M200');
      await g.press('a', ['Control']);
      assert.ok(await view.eval(() => window.__g.allFull()), 'Ctrl+A on an empty cell selects all');
    }));

  test("H5-F16 the formula bar shows numbers like Excel ('1E+21', not '1e+21')", (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('H7');
      assert.equal((await g.ui()).formula, '1E+21');
      await g.goto('H8');
      assert.equal((await g.ui()).formula, '0.000000000001');
    }));

  test('H5-F17 zoom button / zoom label / status bar clicks keep keyboard focus in the grid', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.goto('A1');
      const z0 = (await g.ui()).zoom;
      await p.click('.fv-zoom-btn:last-child');
      await view.settle();
      await g.press('ArrowDown');
      await g.press('Enter');
      const a = await g.ui();
      assert.equal(a.name, 'A3');
      assert.equal(a.zoom, `${parseInt(z0, 10) + 10}%`);
      await p.click('.fv-zoom-label');
      await view.settle();
      const sb = await view.eval(() => {
        const r = document.querySelector('.fv-status-spacer').getBoundingClientRect();
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
      });
      await p.mouse.click(sb.x, sb.y);
      await g.press('ArrowDown');
      const b = await g.ui();
      assert.equal(b.name, 'A4');
      assert.match(b.focus, /^fv-scroller/);
    }));

  test('H5-F18 Ctrl+PageDown from a sheet tab / the name box switches sheets and is not forwarded to VS Code', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.ready();
      await p.focus('.fv-tab-active');
      await view.eval(() => window.__fwd.splice(0));
      await g.press('PageDown', ['Control']);
      assert.equal((await g.ui()).tab, 'Other Sheet');
      await p.focus('.fv-namebox');
      await g.press('PageDown', ['Control']);
      assert.equal((await g.ui()).tab, 'Frozen');
      const leaked = (await view.eval(() => window.__fwd.splice(0))).filter((k) => k.includes('PageDown'));
      assert.deepEqual(leaked, []);
    }));

  test('H5-F19 Escape in the grid closes the hidden-sheets menu', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await view.page.click('.fv-hidden-btn');
      await H.sleep(100);
      await view.page.focus('.fv-scroller');
      await g.press('Escape');
      assert.equal(await view.eval(() => !!document.querySelector('.fv-menu')), false);
    }));

  test('H5-F20 a revealed hidden sheet that is active is restored after a reload', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await view.page.click('.fv-hidden-btn');
      await H.sleep(100);
      await view.page.click('.fv-menu-item');
      await view.settle();
      assert.equal((await g.ui()).tab, 'Secret');
      await H.sleep(700);
      await view.reload();
      assert.equal((await g.ui()).tab, 'Secret');
    }));

  test('H5-F21 a header click while scrolled keeps the view (active cell = first shown cell)', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('M200');
      const st0 = (await g.ui()).st;
      await g.clickHdr('col', 'N');
      const a = await g.ui();
      assert.equal(a.name, 'N:N');
      assert.ok(Number(/^N(\d+)$/.exec(a.active || '')?.[1] ?? 0) > 100, `active ${a.active}`);
      await g.press('ArrowDown');
      const b = await g.ui();
      assert.ok(Math.abs(b.st - st0) < 40, `scrollTop ${st0} -> ${b.st}`);
    }));

  test('R2-F10 Ctrl+. walks the corners; Ctrl+Backspace shows the active cell; Shift+Backspace collapses', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await g.goto('L40:N44');
      const corners = [];
      for (let i = 0; i < 4; i++) {
        await g.press('Period', ['Control']);
        corners.push((await g.ui()).active);
      }
      assert.deepEqual(corners, ['N40', 'N44', 'L44', 'L40']);
      assert.equal(await g.nameBox(), 'L40:N44');
      await g.press('Backspace', ['Shift']);
      assert.equal(await g.nameBox(), 'L40');
      await view.eval(() => (document.querySelector('.fv-scroller').scrollTop = 8000));
      await H.sleep(150);
      await g.press('Backspace', ['Control']);
      await H.sleep(150);
      const top = await view.eval(() => document.querySelector('.fv-scroller').scrollTop);
      assert.ok(top > 100 && top < 1200, `scrollTop ${top}: row 40 back in view`);
      assert.ok(await view.eval(() => !!window.__g.hdr('row', '40')), 'row 40 shown');
      assert.equal(await g.nameBox(), 'L40');
    }));

  test('D2-VS1 a change right before the webview unloads survives the reload (state save flushed)', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      await H.sleep(600);
      await g.clickTab('Other Sheet');
      await view.page.focus('.fv-scroller');
      await view.page.keyboard.press('ArrowDown');
      await view.page.keyboard.press('ArrowDown');
      await H.sleep(50);
      await view.reload();
      const u = await g.ui();
      assert.equal(u.tab, 'Other Sheet');
      assert.equal(u.name, 'A3');
    }));

  test('D2-INV1 worksheets that disappear (invalidate) and come back restore the grid; so does an init without any', (h) =>
    h.use(nav(), {}, async (view) => {
      const g = H.grid(view);
      const host = view.host;
      const meta = (await host.model()).getMeta();
      await g.goto('C7');
      await view.page.keyboard.press('ArrowDown');
      await view.push({ type: 'invalidate', meta: { ...meta, sheets: [] } });
      await H.sleep(300);
      assert.match(await view.eval(() => document.querySelector('.fv-empty')?.textContent || ''), /no worksheets/);
      await view.push({ type: 'invalidate', meta });
      await H.sleep(300);
      await view.settle();
      const back = await g.ui();
      assert.equal(back.name, 'C8');
      assert.match(back.focus, /^fv-scroller/);
      await g.press('ArrowDown');
      assert.equal((await g.ui()).name, 'C9');
      await view.push({ type: 'init', kind: 'sheet', fileName: 'nav.xlsx', meta: { ...meta, sheets: [] }, readOnly: true, banner: 'Not shown: Charts.' });
      await H.sleep(300);
      await view.push({ type: 'invalidate', meta });
      await H.sleep(300);
      await view.settle();
      const b = await view.eval(() => ({ grid: !!document.querySelector('#app .fv-scroller'), tabs: document.querySelectorAll('#app .fv-tab').length, banner: document.querySelector('.fv-banner-text')?.textContent || '' }));
      assert.deepEqual(b, { grid: true, tabs: 3, banner: 'Not shown: Charts.' });
    }));

  test('D2-SB1 in a narrow status bar the stats give way: zoom controls stay visible, no stat cut in half', async (h) => {
    for (const width of [560, 400]) {
      await h.use(nav(), { width, height: 500 }, async (view) => {
        const g = H.grid(view);
        await g.goto('H4:H13');
        await g.stats();
        const m = await view.eval(() => {
          const bar = document.querySelector('.fv-statusbar').getBoundingClientRect();
          const box = document.querySelector('.fv-status-stats').getBoundingClientRect();
          const onLine = [...document.querySelectorAll('.fv-stat')].map((e) => e.getBoundingClientRect()).filter((r) => r.top >= bar.top - 0.5 && r.bottom <= bar.bottom + 0.5);
          return { zoom: document.querySelector('.fv-zoom').getBoundingClientRect().right <= bar.right + 0.5, shown: onLine.filter((r) => r.right <= box.right + 0.5).length, cut: onLine.filter((r) => r.left < box.right - 0.5 && r.right > box.right + 0.5).length };
        });
        assert.ok(m.zoom && m.shown >= 1 && m.cut === 0, `@${width}: ${JSON.stringify(m)}`);
      });
    }
  });

  test('H2-F11 a workbook without worksheets (chart sheets only) explains itself and keeps the banner', (h) =>
    h.use(H.sheetHost('nav.xlsx', F.nav, { meta: (m) => ({ ...m, sheets: [], activeSheet: 0 }), banner: 'This workbook contains content that FileStudio does not display: Charts.' }), {}, async (view) => {
      const out = await view.eval(() => ({ text: document.querySelector('.fv-empty')?.textContent || '', banner: document.querySelector('.fv-banner-text')?.textContent || '' }));
      assert.match(out.text, /chart sheets/i);
      assert.match(out.banner, /Charts/);
    }));
});
