// Regression tests for fixed 0.1.x grid RENDERING bugs (H4 overflow / clipping / alignment / borders / decorations,
// K8 right-to-left). Geometry is read from the DOM (text node rects vs. the cell rect from the headers), so the
// checks do not depend on screenshots or on the fonts installed.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const render = () => H.sheetHost('render.xlsx', F.render);

/** Text box and cell rect of the cell whose text includes `text` (cell at r, c0..c1). */
const geo = (view, text, r, c0, c1 = c0) =>
  view.eval(
    (text, r, c0, c1) => {
      const t = window.__g.textBox(window.__g.cellEl(text));
      const a = window.__g.cellRect(r, c0);
      const b = window.__g.cellRect(r, c1);
      return { textL: t.x, textR: t.r, textT: t.y, textB: t.b, cellL: a.x, cellR: b.r, cellT: a.y, cellB: a.b };
    },
    text,
    r,
    c0,
    c1,
  );

H.suite('sheet: rendering (0.1.x regressions)', ({ test }) => {
  test('H4-F1 clipped right-aligned / centred text shows its end / middle, not its start', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const right = await geo(view, 'START merged right', 2, 1, 2);
      assert.ok(Math.abs(right.textR - (right.cellR - 2)) <= 3, `right-aligned end at the right edge: ${JSON.stringify(right)}`);
      assert.ok(right.textL < right.cellL, 'its start is clipped');
      const mid = await geo(view, 'START merged centred', 4, 1, 2);
      assert.ok(Math.abs((mid.textL + mid.textR) / 2 - (mid.cellL + mid.cellR) / 2) <= 3, `centred: ${JSON.stringify(mid)}`);
      assert.ok(mid.textL < mid.cellL);
    }));

  test('H4-F2 a formula returning "" blocks the text overflow of its left neighbour', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const out = await view.eval(() => {
        const el = window.__g.cellEl('This is a long piece');
        const ov = el.querySelector('.fv-ov');
        const t = window.__g.textBox(ov || el);
        const vis = (ov || el).getBoundingClientRect();
        return { overflowBox: !!ov, visibleRight: Math.min(vis.right, t.r), cellRight: window.__g.cellRect(0, 0).r };
      });
      assert.equal(out.overflowBox, false);
      assert.ok(out.visibleRight <= out.cellRight + 0.5, JSON.stringify(out));
    }));

  test('H4-F4 Center Across Selection centres the text over the following empty cells', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const g = await geo(view, 'Title', 6, 0, 4);
      assert.ok(Math.abs((g.textL + g.textR) / 2 - (g.cellL + g.cellR) / 2) <= 3, JSON.stringify(g));
    }));

  test("H4-F5 'fill' alignment repeats the content across the cell and never overflows", (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const out = await view.eval(() => {
        const cells = [...document.querySelectorAll('.fv-c')];
        const a = cells.find((e) => /^-+$/.test(e.textContent));
        const b = cells.find((e) => /^(ab)+$/.test(e.textContent));
        const d = cells.find((e) => e.textContent.startsWith('fill-aligned'));
        return { a: a && a.textContent, b: b && b.textContent, dOverflows: !!d.querySelector('.fv-ov'), dRight: window.__g.textBox(d).r, cellR: window.__g.cellRect(8, 3).r };
      });
      assert.ok(out.a && out.a.length > 5, `'-' repeated: ${out.a}`);
      assert.ok(out.b && out.b.length >= 6, `'ab' repeated: ${out.b}`);
      assert.equal(out.dOverflows, false);
    }));

  test("H4-F6 'distributed' alignment spreads the text with the indent on both sides", (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const g = await geo(view, 'distrib', 10, 3);
      assert.ok(g.textL >= g.cellL + 9, `indent on the left: ${JSON.stringify(g)}`);
      assert.ok(Math.abs(g.textR - (g.cellR - 2 - 9)) <= 3, `indent on the right: ${JSON.stringify(g)}`);
    }));

  test('H4-F8 merged areas take their outer borders from the edge cells, not all from the master', (h) =>
    h.use(render(), {}, async (view) => {
      const grid = H.grid(view);
      await grid.clickTab('Merges');
      await grid.ready();
      const css = (text) =>
        view.eval((text) => {
          const el = window.__g.cellEl(text);
          const cs = getComputedStyle(el);
          const segs = [...el.querySelectorAll('.fv-bseg:not(.fv-bseg-gl)')].map((s) => ['Top', 'Right', 'Bottom', 'Left'].filter((k) => parseFloat(getComputedStyle(s)[`border${k}Width`]) > 0).join(''));
          return { t: cs.borderTopWidth, r: cs.borderRightWidth, b: cs.borderBottomWidth, l: cs.borderLeftWidth, segs };
        }, text);
      const m = await css('master-only box');
      assert.deepEqual([m.r, m.b, m.segs.join()], ['0px', '0px', 'Top,Left'], `master-only box: ${JSON.stringify(m)}`);
      const e = await css('edge cells');
      assert.deepEqual([e.t, e.r, e.b, e.l], ['1px', '1px', '1px', '1px'], `edge cells: ${JSON.stringify(e)}`);
    }));

  test('H4-F9 a shared edge with conflicting borders is drawn once (Excel precedence), also at the frozen-pane edge', (h) =>
    h.use(render(), {}, async (view) => {
      const grid = H.grid(view);
      await grid.clickTab('Frozen');
      await grid.ready();
      const out = await view.eval(() => {
        const w = (text, side) => getComputedStyle(window.__g.cellEl(text))[`border${side}Width`];
        return { B10R: w('B10 thick red', 'Right'), C10L: w('C10 thick blue', 'Left'), G10R: w('G10 thick red', 'Right'), H10L: w('H10 thick blue', 'Left'), F3B: w('F3 bottom', 'Bottom'), F4T: w('F4 top', 'Top') };
      });
      assert.deepEqual(out, { B10R: '0px', C10L: '3px', G10R: '0px', H10L: '3px', F3B: '0px', F4T: '3px' });
    }));

  test('H4-F10 shrink-to-fit also shrinks rich-text runs that have their own size', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const out = await view.eval(() => {
        const el = window.__g.cellEl('Rich ');
        const run = [...el.querySelectorAll('span')].find((s) => s.textContent.startsWith('shrink run'));
        return { runPx: parseFloat(getComputedStyle(run).fontSize), textR: window.__g.textBox(el).r, cellR: window.__g.cellRect(14, 1).r };
      });
      assert.ok(out.runPx < 18.6, `14 pt run shrunk: ${JSON.stringify(out)}`);
      assert.ok(out.textR <= out.cellR + 0.5, `text inside the cell: ${JSON.stringify(out)}`);
    }));

  test('H4-F11 text rotated by other angles than +-90 stays inside its cell', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const out = await view.eval(() => {
        const res = {};
        ['rot 45', 'rot -45', 'rot 30'].forEach((text, i) => {
          const el = [...document.querySelectorAll('.fv-c')].find((e) => e.textContent === text);
          const t = window.__g.textBox(el);
          const c = window.__g.cellRect(19, i);
          res[text] = t.x >= c.x - 0.5 && t.r <= c.r + 0.5 && t.y >= c.y - 0.5 && t.b <= c.b + 0.5;
        });
        return res;
      });
      assert.deepEqual(out, { 'rot 45': true, 'rot -45': true, 'rot 30': true });
    }));

  test('H4-F13 double underline + strikethrough keeps a single strikethrough', (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const d = await view.eval(() => {
        const el = window.__g.cellEl('dbl+strike');
        return [...new Set([el, ...el.querySelectorAll('*')].map((e) => getComputedStyle(e)).map((cs) => `${cs.textDecorationLine}|${cs.textDecorationStyle}`))];
      });
      assert.ok(!d.some((x) => x.includes('line-through') && x.endsWith('|double')), `no double strikethrough: ${d}`);
      assert.ok(d.some((x) => x.includes('line-through')), `a strikethrough: ${d}`);
      assert.ok(d.some((x) => x.startsWith('underline') && x.endsWith('|double')), `a double underline: ${d}`);
    }));

  test("H4-F14 a rich-text superscript run is 3/4 of the run's own size", (h) =>
    h.use(render(), {}, async (view) => {
      await H.grid(view).ready();
      const px = await view.eval(() => parseFloat(getComputedStyle([...window.__g.cellEl('big sup').querySelectorAll('span')].find((s) => s.textContent === '2')).fontSize));
      assert.ok(Math.abs(px - 24) < 0.6, `${px}px (24 pt run = 32px x 0.75 = 24px)`);
    }));

  test('H4-F15 numbers in wrap-text cells stay on one line: ### instead of digits broken across lines', (h) =>
    h.use(render(), {}, async (view) => {
      const grid = H.grid(view);
      await grid.goto('L17');
      const out = await view.eval(() => {
        const res = [];
        for (const r of [16, 17]) {
          const c = window.__g.cellRect(r, 11);
          const el = [...document.querySelectorAll('.fv-c')].find((e) => Math.abs(e.getBoundingClientRect().top - c.y) < 2 && Math.abs(e.getBoundingClientRect().left - c.x) < 2);
          const t = window.__g.textBox(el);
          res.push({ text: el.textContent, lines: Math.round((t.b - t.y) / parseFloat(getComputedStyle(el).fontSize) / 1.1) });
        }
        return res;
      });
      assert.match(out[0].text, /^#+$/, JSON.stringify(out));
      assert.match(out[1].text, /^3\.14/, JSON.stringify(out));
      assert.ok(out.every((x) => x.lines <= 1), JSON.stringify(out));
    }));

  test('K8 right-to-left sheet: tab bar mirrored, rotation angle mirrored, formula bar stays left-to-right', (h) =>
    h.use(H.sheetHost('rtl.xlsx', F.rtl), {}, async (view) => {
      const grid = H.grid(view);
      await grid.ready();
      const probe = () =>
        view.eval(() => {
          const rot = [...document.querySelectorAll('.fv-c')].find((e) => e.textContent === 'rot 30');
          const node = document.createTreeWalker(rot, NodeFilter.SHOW_TEXT).nextNode();
          const glyph = (i) => {
            const q = document.createRange();
            q.setStart(node, i);
            q.setEnd(node, i + 1);
            const b = q.getBoundingClientRect();
            return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
          };
          const tabs = Object.fromEntries([...document.querySelectorAll('.fv-tab')].map((t) => [t.textContent, t.getBoundingClientRect().left]));
          return {
            descends: glyph(node.length - 1).y > glyph(0).y,
            rtlTabRight: tabs.RTL > tabs.LTR,
            formulaBarLtr: document.querySelector('.fv-namebox').getBoundingClientRect().left < document.querySelector('.fv-formula').getBoundingClientRect().left,
          };
        });
      const rtl = await probe();
      assert.deepEqual(rtl, { descends: true, rtlTabRight: true, formulaBarLtr: true }, 'RTL sheet');
      await view.eval(() => document.querySelector('.fv-tab-active').focus());
      await view.page.keyboard.press('ArrowLeft'); // mirrored bar: the next tab is on the left
      await H.sleep(100);
      await view.settle();
      assert.equal(await grid.activeTab(), 'LTR');
      const ltr = await probe();
      assert.deepEqual(ltr, { descends: false, rtlTabRight: false, formulaBarLtr: true }, 'LTR sheet');
    }));
});
