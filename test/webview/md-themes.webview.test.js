// Markdown modes in the four themes (v0.2.0, P6): light, dark, high contrast (dark) and high contrast light. The toolbar, the
// table of contents and the WYSIWYG editor take their colours from the VS Code theme variables; here they are checked for
// readable text (WCAG contrast), a visible "chosen" mode button and the HC borders. With FS_SHOTS=<folder> a screenshot of
// every theme x mode is written there for a look by eye (the tests do not need it).
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const H = require('./harness');
const { mdHost } = require('./md-host');

const KS = fs.readFileSync(path.join(__dirname, '..', 'samples', 'kitchen-sink.md'), 'utf8').replace(/\r\n/g, '\n');
const SHOTS = process.env.FS_SHOTS;
const THEMES = [
  { theme: 'light', min: 4.5 },
  { theme: 'dark', min: 4.5 },
  { theme: 'hc', min: 7 },
  { theme: 'hcl', min: 7 },
];
/** The images in the sample are relative; they do not exist in the harness (404 is expected). */
const OPTS = (theme) => ({ theme, ignoreErrors: /HTTP 404: .*\.(png|jpe?g|gif|svg)/ });

/** Page code: the WCAG contrast of the text colour of the first element matching `selector` against what is behind it. */
async function contrastOf(view, selector) {
  return view.eval((selector) => {
    const el = [...document.querySelectorAll(selector)].find((e) => e.getClientRects().length);
    if (!el) return { missing: selector };
    const ctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
    const rgba = (c) => {
      ctx.clearRect(0, 0, 1, 1);
      ctx.fillStyle = '#000000';
      ctx.fillStyle = c;
      ctx.fillRect(0, 0, 1, 1);
      const d = ctx.getImageData(0, 0, 1, 1).data;
      return [d[0], d[1], d[2], d[3] / 255];
    };
    const parse = (c) => {
      // fillStyle -> pixel loses the alpha of a transparent colour on a cleared canvas: read it from the string instead
      const m = /rgba?\(([^)]+)\)/.exec(c);
      if (m) {
        const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number);
        return [p[0], p[1], p[2], p.length > 3 ? p[3] : 1];
      }
      return rgba(c);
    };
    // the layers from the element up: backgrounds (translucent ones are composited over what is below)
    const layers = [];
    for (let e = el; e; e = e.parentElement) {
      const bg = parse(getComputedStyle(e).backgroundColor);
      if (bg && bg[3] > 0) layers.push(bg);
      if (bg && bg[3] === 1) break;
    }
    let base = [255, 255, 255];
    for (const [r, g, b, a] of layers.reverse()) base = [r * a + base[0] * (1 - a), g * a + base[1] * (1 - a), b * a + base[2] * (1 - a)];
    const fg = parse(getComputedStyle(el).color);
    const text = [0, 1, 2].map((i) => fg[i] * fg[3] + base[i] * (1 - fg[3]));
    const lum = ([r, g, b]) => [r, g, b].map((v) => ((v /= 255) <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)).reduce((s, v, i) => s + v * [0.2126, 0.7152, 0.0722][i], 0);
    const [hi, lo] = [lum(text), lum(base)].sort((a, b) => b - a);
    return { ratio: Math.round(((hi + 0.05) / (lo + 0.05)) * 100) / 100, text: getComputedStyle(el).color, back: base.map(Math.round).join(',') };
  }, selector);
}

async function shot(view, theme, name) {
  if (!SHOTS) return;
  fs.mkdirSync(SHOTS, { recursive: true });
  await view.page.screenshot({ path: path.join(SHOTS, `${theme}-${name}.png`) });
}

const waitEditor = (view) => view.page.waitForFunction(() => document.querySelector('.md-wysiwyg:not([hidden]) .milkdown .ProseMirror'), { timeout: 20000, polling: 50 });

H.suite('markdown modes in four themes (P6)', ({ test }) => {
  for (const { theme, min } of THEMES) {
    test(`${theme}: toolbar, TOC and preview are readable; the chosen mode is visible`, async (h) => {
      await h.use(mdHost('themes.md', KS), OPTS(theme), async (view) => {
        await view.page.click('.md-toc-btn');
        await view.settle();
        for (const selector of ['.md-article p', '.md-article h2', '.md-mode-btn[aria-checked="true"]', '.md-mode-btn[aria-checked="false"]', '.md-toc-row']) {
          const c = await contrastOf(view, selector);
          assert.ok(c.ratio >= min, `${theme} ${selector}: contrast ${c.ratio} < ${min} (${JSON.stringify(c)})`);
        }
        // the chosen button differs from the others in more than text (background, border or outline)
        const look = await view.eval(() => {
          const style = (b) => {
            const s = getComputedStyle(b);
            return [s.backgroundColor, s.borderTopColor, s.borderTopWidth, s.outlineStyle, s.outlineColor, s.fontWeight].join('|');
          };
          const [on, off] = [document.querySelector('.md-mode-btn[aria-checked="true"]'), document.querySelector('.md-mode-btn[aria-checked="false"]')];
          return { on: style(on), off: style(off) };
        });
        assert.notEqual(look.on, look.off, `${theme}: the chosen mode looks like the others`);
        if (theme === 'hc' || theme === 'hcl') assert.match(look.on, /solid/, `${theme}: the chosen mode has an outline in high contrast`);
        await shot(view, theme, 'preview-toc');
        // the mouse stays on the button that was chosen: it must stay readable (HC light once lost its background there)
        await view.page.click('.md-mode-btn[data-mode="split"]');
        await view.settle();
        const hovered = await contrastOf(view, '.md-mode-btn[aria-checked="true"]:hover');
        assert.ok(hovered.ratio >= min, `${theme}: the chosen mode under the mouse: contrast ${hovered.ratio} (${JSON.stringify(hovered)})`);
        await shot(view, theme, 'split');
      });
    });

    test(`${theme}: the WYSIWYG editor is readable and its protected blocks are marked`, async (h) => {
      await h.use(mdHost('themes.md', KS), OPTS(theme), async (view) => {
        await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
        await waitEditor(view);
        await view.settle();
        for (const selector of ['.md-wysiwyg .ProseMirror p', '.md-wysiwyg .ProseMirror h2', '.md-wysiwyg .ProseMirror a', '.md-wysiwyg .ProseMirror li', '.md-wysiwyg .ProseMirror li .label-wrapper', '.md-wysiwyg .ProseMirror pre.fs-raw']) {
          const c = await contrastOf(view, selector);
          assert.ok(c.ratio >= (selector.endsWith('a') ? Math.min(min, 4.5) : min), `${theme} ${selector}: contrast ${c.ratio} (${JSON.stringify(c)})`);
        }
        const raw = await view.eval(() => {
          const pre = document.querySelector('.md-wysiwyg .ProseMirror pre.fs-raw');
          const s = getComputedStyle(pre);
          return { label: getComputedStyle(pre, '::before').content, border: s.borderLeftWidth, line: s.borderLeftStyle, other: s.borderTopStyle };
        });
        assert.match(raw.label, /Markdown source: /);
        assert.equal(raw.border, '3px');
        if (theme === 'hc' || theme === 'hcl') assert.equal(raw.other, 'solid', `${theme}: a protected block has a full border in high contrast`);
        // nothing is sent just by looking at it
        assert.deepEqual(await view.posted('edit'), []);
        await shot(view, theme, 'wysiwyg');
      });
    });
  }
});
