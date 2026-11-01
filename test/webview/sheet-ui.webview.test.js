// Regression tests for fixed 0.1.x grid UI bugs: sheet tabs, zoom, status bar, Reduce Motion, keyboard regions,
// focus, ARIA grid, notes and links for keyboard users, contrast in the real VS Code themes.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const styled = () => H.sheetHost('styled.xlsx', () => H.sample('styled.xlsx'));
const cells = (opts) => H.sheetHost('cells.xlsx', F.cells, opts);
const tabs = () => H.sheetHost('tabs.xlsx', F.tabs);
const BANNER = 'This workbook contains content that FileStudio does not display: Charts.';

/** Lowest WCAG contrast of the visible text of `sel` against its (composited) background. */
function contrast(sel) {
  const parse = (s) => {
    const m = /rgba?\(([^)]+)\)/.exec(s);
    if (m) {
      const p = m[1].split(/[\s,/]+/).filter(Boolean).map(Number);
      return [p[0], p[1], p[2], p[3] === undefined ? 1 : p[3]];
    }
    const c = /color\(srgb ([\d.]+) ([\d.]+) ([\d.]+)(?: \/ ([\d.]+))?\)/.exec(s);
    return c ? [c[1] * 255, c[2] * 255, c[3] * 255, c[4] === undefined ? 1 : Number(c[4])] : null;
  };
  const over = (t, b) => [0, 1, 2].map((i) => t[i] * t[3] + b[i] * (1 - t[3])).concat(1);
  const lum = (c) => {
    const f = (v) => ((v /= 255) <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4);
    return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2]);
  };
  const ratio = (a, b) => (Math.max(lum(a), lum(b)) + 0.05) / (Math.min(lum(a), lum(b)) + 0.05);
  const pageBg = getComputedStyle(document.documentElement).getPropertyValue('--vscode-editor-background').trim();
  let min = 99;
  let n = 0;
  for (const el of document.querySelectorAll(sel)) {
    const r = el.getBoundingClientRect();
    if (!el.textContent.trim() || r.width < 1 || r.bottom < 0 || r.top > innerHeight || getComputedStyle(el).display === 'none') continue;
    const chain = [];
    for (let e = el; e; e = e.parentElement) {
      const c = parse(getComputedStyle(e).backgroundColor);
      if (c && c[3] > 0) {
        chain.push(c);
        if (c[3] >= 1) break;
      }
    }
    let bg = [1, 3, 5].map((i) => parseInt(pageBg.slice(i, i + 2), 16)).concat(1);
    for (let i = chain.length - 1; i >= 0; i--) bg = over(chain[i], bg);
    let op = 1;
    for (let e = el; e; e = e.parentElement) op *= Number(getComputedStyle(e).opacity);
    const fg = parse(getComputedStyle(el).color);
    min = Math.min(min, Math.round(ratio(over([fg[0], fg[1], fg[2], fg[3] * op], bg), bg) * 100) / 100);
    n++;
  }
  return { min, n };
}

/** Screenshot colour 'r,g,b' of one page pixel. */
async function pixel(page, x, y) {
  const buf = await page.screenshot({ clip: { x, y, width: 1, height: 1 }, encoding: 'base64' });
  return page.evaluate(async (b64) => {
    const img = new Image();
    img.src = 'data:image/png;base64,' + b64;
    await img.decode();
    const cv = document.createElement('canvas');
    cv.width = cv.height = 1;
    const ctx = cv.getContext('2d');
    ctx.drawImage(img, 0, 0);
    return Array.from(ctx.getImageData(0, 0, 1, 1).data.slice(0, 3)).join(',');
  }, buf);
}

H.suite('sheet: tabs, zoom, status bar, focus, accessibility (0.1.x regressions)', ({ test }) => {
  test('H5-F12 a click on a tab-scroll arrow keeps keyboard focus in the grid', (h) =>
    h.use(tabs(), { width: 900, height: 600 }, async (view) => {
      const g = H.grid(view);
      const btns = await view.page.$$('.fv-tab-scroll');
      await btns[1].click();
      await H.sleep(300);
      await g.press('ArrowDown');
      const u = await g.ui();
      assert.equal(u.name, 'A2');
      assert.match(u.focus, /^fv-scroller/);
    }));

  test('R2-F13 a click on the disabled "previous sheets" arrow keeps the grid focused (not <body>)', (h) =>
    h.use(tabs(), { width: 900, height: 600 }, async (view) => {
      const g = H.grid(view);
      await g.goto('B2');
      await view.eval(() => (document.querySelector('.fv-tabs').scrollLeft = 0));
      await H.sleep(300);
      const prev = await view.page.$eval('.fv-tab-scroll', (e) => {
        const r = e.getBoundingClientRect();
        return { disabled: e.disabled, x: r.left + r.width / 2, y: r.top + r.height / 2 };
      });
      assert.ok(prev.disabled, 'the previous arrow is disabled at the start');
      await view.page.mouse.click(prev.x, prev.y);
      await H.sleep(100);
      await g.press('ArrowDown');
      await g.press('ArrowDown');
      assert.match(await g.focus(), /^fv-scroller/);
      assert.equal(await g.nameBox(), 'B4');
    }));

  test('R2-F16 Ctrl+wheel zooms in 15% steps like Excel', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const xy = await g.cellXY(3, 2);
      await view.page.mouse.move(xy.x, xy.y);
      const zooms = [];
      for (let i = 0; i < 2; i++) {
        await view.page.keyboard.down('Control');
        await view.page.mouse.wheel({ deltaY: -100 });
        await view.page.keyboard.up('Control');
        await H.sleep(200);
        zooms.push(await view.page.$eval('.fv-zoom-label', (e) => e.textContent));
      }
      assert.deepEqual(zooms, ['115%', '130%']);
    }));

  test('R2H4-4 at 220px the zoom -, % and + stay inside the status bar and clickable', (h) =>
    h.use(cells(), { width: 220, height: 400 }, async (view) => {
      const g = H.grid(view);
      await g.goto('D1:D4');
      await g.stats();
      const bar = await view.eval(() => {
        const sb = document.querySelector('.fv-statusbar').getBoundingClientRect();
        const btns = [...document.querySelectorAll('.fv-zoom-btn, .fv-zoom-label')].map((e) => {
          const r = e.getBoundingClientRect();
          const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
          return { r: Math.round(r.right), hit: !!hit && (hit === e || e.contains(hit)) };
        });
        return { right: Math.round(sb.right), btns };
      });
      assert.equal(bar.btns.length, 3);
      assert.ok(bar.btns.every((x) => x.hit && x.r <= bar.right), JSON.stringify(bar));
    }));

  test('R2H4-5 Reduce Motion (body.vscode-reduce-motion) stops the loading shimmer, spinner and stats fade', (h) =>
    h.use(cells(), {}, async (view) => {
      const motion = await view.eval(() => {
        const host = document.querySelector('.fv-content');
        const rows = document.createElement('div');
        rows.className = 'fv-loading-rows';
        host.appendChild(rows);
        const boot = document.createElement('div');
        boot.className = 'fv-boot-message';
        document.body.appendChild(boot);
        const read = () => [getComputedStyle(rows).animationName, getComputedStyle(boot, '::before').animationName, getComputedStyle(document.querySelector('.fv-status-stats')).transitionDuration];
        document.body.classList.add('vscode-reduce-motion');
        const on = read();
        document.body.classList.remove('vscode-reduce-motion');
        const off = read();
        rows.remove();
        boot.remove();
        return { on, off };
      });
      assert.deepEqual(motion.on, ['none', 'none', '0s']);
      assert.deepEqual(motion.off, ['fv-shimmer', 'fv-spin', '0.12s']);
    }));

  test('H9-F2 F6 / Shift+F6 / Ctrl+G reach every region (banner, name box, grid, tabs, status bar); Tab stays a cell move', (h) =>
    h.use(cells({ banner: BANNER }), {}, async (view) => {
      const g = H.grid(view);
      const p = view.page;
      await g.ready();
      await p.focus('.fv-scroller');
      const walk = [];
      for (const k of ['F6', 'F6', 'F6', 'Shift+F6', 'Shift+F6', 'Shift+F6', 'Shift+F6']) {
        const parts = k.split('+');
        if (parts.length > 1) await p.keyboard.down('Shift');
        await p.keyboard.press(parts[parts.length - 1]);
        if (parts.length > 1) await p.keyboard.up('Shift');
        await H.sleep(60);
        walk.push(await g.focus());
      }
      assert.match(walk[0], /^fv-tab/, walk.join(' > '));
      assert.match(walk[1], /Zoom out/, walk.join(' > '));
      assert.match(walk[3], /^fv-tab/, walk.join(' > '));
      assert.match(walk[4], /^fv-scroller/, walk.join(' > '));
      assert.match(walk[5], /^fv-namebox/, walk.join(' > '));
      assert.match(walk[6], /Dismiss/, walk.join(' > '));
      await p.focus('.fv-scroller');
      await p.keyboard.down('Control');
      await p.keyboard.press('g');
      await p.keyboard.up('Control');
      await H.sleep(60);
      assert.match(await g.focus(), /^fv-namebox/, 'Ctrl+G');
      await p.focus('.fv-scroller');
      await p.keyboard.press('Tab');
      await H.sleep(60);
      assert.match(await g.focus(), /^fv-scroller/, 'Tab in the grid');
    }));

  test('H9-F3 the banner Dismiss gives keyboard focus back to the grid (arrows work)', (h) =>
    h.use(cells({ banner: BANNER }), {}, async (view) => {
      const g = H.grid(view);
      const before = (await g.ui()).name;
      await view.page.click('.fv-banner-close');
      await H.sleep(100);
      assert.match(await g.focus(), /^fv-scroller/);
      await g.press('ArrowDown', [], 2);
      assert.notEqual((await g.ui()).name, before);
    }));

  test('H9-F9 ARIA grid: row / column count, the active descendant is a gridcell inside a row', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      await view.page.focus('.fv-scroller');
      await g.press('ArrowDown');
      await H.sleep(220);
      const aria = await view.eval(() => {
        const grid = document.querySelector('[role=grid]');
        const cell = grid && document.getElementById(grid.getAttribute('aria-activedescendant') || '');
        return { rowcount: grid?.getAttribute('aria-rowcount'), colcount: grid?.getAttribute('aria-colcount'), cell: cell ? `${cell.getAttribute('role')} in ${cell.parentElement.getAttribute('role')} ${cell.parentElement.getAttribute('aria-rowindex')}/${cell.getAttribute('aria-colindex')}` : null };
      });
      assert.deepEqual(aria, { rowcount: '1048576', colcount: '16384', cell: 'gridcell in row 2/1' });
    }));

  test('H9-F12 cell notes and link targets reach keyboard users (description, tooltip, Ctrl+Enter)', (h) =>
    h.use(cells(), {}, async (view) => {
      const g = H.grid(view);
      const desc = () =>
        view.eval(() => {
          const grid = document.querySelector('.fv-scroller');
          const cell = document.getElementById(grid.getAttribute('aria-activedescendant') || '');
          const d = cell ? (cell.getAttribute('aria-describedby') || '').split(' ').map((id) => document.getElementById(id)?.textContent || '').join(' ') : '';
          const tip = document.querySelector('.fv-tooltip');
          return { desc: d, tip: !tip || tip.hidden ? '' : tip.textContent };
        });
      await g.goto('E9');
      await g.press('ArrowDown'); // E10 (note)
      await H.sleep(600);
      const note = await desc();
      assert.match(note.desc, /A cell note/);
      assert.ok(note.tip.length > 0, 'tooltip shown');
      await g.goto('E7');
      await g.press('ArrowDown'); // E8 (link)
      await H.sleep(600);
      const link = await desc();
      assert.match(link.desc, /https:\/\/example\.com/);
      await view.clearPosted();
      await g.press('Enter', ['Control']);
      await H.sleep(200);
      assert.deepEqual((await view.posted('openLink')).map((m) => m.href), ['https://example.com/']);
    }));

  test('H9-C3 header labels reach 4.5:1 contrast (light, dark, high contrast)', async (h) => {
    for (const theme of ['light', 'dark', 'hc']) {
      await h.use(styled(), { theme }, async (view) => {
        const r = await view.eval(contrast, '.fv-rh:not(.fv-hsel), .fv-ch:not(.fv-hsel)');
        assert.ok(r.n > 0 && r.min >= 4.5, `${theme}: ${r.min}:1 over ${r.n} labels`);
      });
    }
  });

  test('H9-C4 inactive coloured sheet tabs keep a 4.5:1 label (light, dark, high contrast)', async (h) => {
    for (const theme of ['light', 'dark', 'hc']) {
      await h.use(styled(), { theme }, async (view) => {
        const r = await view.eval(contrast, '.fv-tab.fv-tab-colored:not(.fv-tab-active) .fv-tab-label');
        assert.ok(r.n > 0 && r.min >= 4.5, `${theme}: ${r.min}:1 over ${r.n} tabs`);
      });
    }
  });

  test('H9-C6 keyboard focus on the grid shows the focus-border frame (not after a mouse click)', (h) =>
    h.use(styled(), {}, async (view) => {
      const p = view.page;
      const r = await view.eval(() => {
        const b = document.querySelector('.fv-scroller').getBoundingClientRect();
        return { x: b.left, y: b.top, w: b.width, h: b.height, fb: getComputedStyle(document.documentElement).getPropertyValue('--vscode-focusBorder').trim() };
      });
      const want = [1, 3, 5].map((i) => parseInt(r.fb.slice(i, i + 2), 16)).join(',');
      const framed = async () => {
        const px = [await pixel(p, r.x, r.y + 100), await pixel(p, r.x + 300, r.y), await pixel(p, r.x + r.w - 1, r.y + 100), await pixel(p, r.x + 300, r.y + r.h - 1)];
        return px.every((c) => c === want) ? 'frame' : px.some((c) => c === want) ? `partial ${px.join('|')}` : 'none';
      };
      await p.focus('.fv-scroller');
      assert.equal(await framed(), 'frame', 'keyboard focus');
      await p.click('.fv-namebox');
      await H.sleep(50);
      assert.equal(await framed(), 'none', 'focus in the name box');
      await p.mouse.click(r.x + 200, r.y + 100);
      await H.sleep(80);
      assert.equal(await framed(), 'none', 'after a mouse click');
    }));
});
