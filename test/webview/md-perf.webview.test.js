// WYSIWYG size behaviour (v0.2.1, B2): typing speed at 1,000 blocks, a Loading state with progress while a big document is
// built (the page stays alive), and the lighter editor above a weight (MD_WYSIWYG_LITE_WEIGHT).
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const { mdHost } = require('./md-host');
const M = H.renderer('markdown');

const hooks = { before: async (pg) => pg.evaluateOnNewDocument(() => { window.__fsTestHooks = {}; }), ignoreErrors: /HTTP 404: .*\.png/ };

/** A document of about `n` top-level blocks: mixed kinds (a table and a list cost most). */
function makeDoc(n) {
  const out = [];
  for (let i = 0; out.length < n; i++) {
    switch (i % 6) {
      case 0: out.push(`## Section ${i}`); break;
      case 1: out.push(`Paragraph ${i} with **bold**, *italic* and a [link](https://example.com/${i}).`); break;
      case 2: out.push(`- item a ${i}\n- item b ${i}\n- item c ${i}`); break;
      case 3: out.push(`| A | B | C |\n| --- | --- | --- |\n| ${i} | x | y |\n| z | w | v |`); break;
      case 4: out.push(`> quote ${i}`); break;
      default: out.push('```js\nconst a = ' + i + ';\n```'); break;
    }
  }
  return out.join('\n\n') + '\n';
}

/** The host of the mock with the weight the real host would send (limits.weight). */
function heavyHost(text) {
  const weight = M.markdownWeight(text, M.markdownBlocks(text));
  return mdHost('big.md', text, { limits: { wysiwyg: true, weight } });
}

H.suite('markdown WYSIWYG size (B2)', ({ test }) => {
  test('typing at 1,000 blocks: a key (transaction + layout) takes under 30 ms; the editor is the lighter one', async (h) => {
    const text = makeDoc(1000);
    const host = heavyHost(text);
    await h.use(host, hooks, async (view) => {
      const t0 = Date.now();
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      await view.page.waitForFunction(() => window.__fsTestHooks.mdw && window.__fsTestHooks.mdw.state.ready, { timeout: 90000, polling: 50 });
      const entry = Date.now() - t0;
      await view.settle();
      const stats = await view.eval(() => {
        const t = window.__fsTestHooks.mdw;
        const v = t.editor.action((ctx) => ctx.get(t.mw.core.editorViewCtx));
        let pos = -1;
        v.state.doc.forEach((n, off) => { if (pos < 0 && n.type.name === 'paragraph' && n.textContent.startsWith('Paragraph 499 ')) pos = off; });
        const times = [];
        for (let i = 0; i < 30; i++) {
          const s = performance.now();
          v.dispatch(v.state.tr.insertText('x', pos + 2));
          void document.body.offsetHeight; // the layout a key causes
          times.push(performance.now() - s);
        }
        times.sort((a, b) => a - b);
        return { median: times[15], max: times[29], lite: t.state.lite };
      });
      console.log(`1000 blocks: entry ${entry} ms, key median ${stats.median.toFixed(1)} ms, max ${stats.max.toFixed(1)} ms, lite ${stats.lite}`);
      assert.equal(stats.lite, true, 'a heavy document gets the lighter editor');
      assert.ok(stats.median < 30, `key median ${stats.median.toFixed(1)} ms`);
    });
  }, { timeout: 180000 });

  test('a big document shows a Loading state with progress, and the page stays alive while it is built', async (h) => {
    const text = makeDoc(1500);
    const host = heavyHost(text);
    await h.use(host, hooks, async (view) => {
      await view.eval(() => {
        const log = (window.__loadLog = { texts: [], gaps: [], on: false, last: 0 });
        new MutationObserver(() => {
          const el = document.querySelector('.md-status');
          const text = el ? el.textContent : '';
          if (text && log.texts[log.texts.length - 1] !== text) log.texts.push(text);
          if (/%/.test(text) && !log.on) { log.on = true; log.last = performance.now(); }
        }).observe(document.body, { subtree: true, childList: true, characterData: true });
        setInterval(() => {
          const now = performance.now();
          if (log.on) { log.gaps.push(now - log.last); }
          log.last = now;
        }, 10);
      });
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      await view.page.waitForFunction(() => window.__fsTestHooks.mdw && window.__fsTestHooks.mdw.state.ready, { timeout: 90000, polling: 50 });
      await view.settle();
      const log = await view.eval(() => ({ texts: window.__loadLog.texts, maxGap: Math.max(0, ...window.__loadLog.gaps), ticks: window.__loadLog.gaps.length }));
      console.log(`1500 blocks: status texts ${JSON.stringify(log.texts.slice(0, 4))}, longest stall ${log.maxGap.toFixed(0)} ms in ${log.ticks} ticks`);
      assert.ok(log.texts.some((t) => /^Loading the WYSIWYG editor… \d+% \([\d,]+ of 1,\d{3} blocks\)$/.test(t)), `a progress text: ${JSON.stringify(log.texts)}`);
      assert.ok(log.ticks >= 5, `the page ran while loading (${log.ticks} ticks)`);
      assert.ok(log.maxGap < 800, `longest stall ${log.maxGap.toFixed(0)} ms`);
      assert.deepEqual(await view.eval(() => document.querySelector('.md-status').textContent), 'The WYSIWYG editor is ready.', 'the live region says it is ready');
      assert.ok(!log.texts.some((t) => /^Loading the WYSIWYG editor… 100%/.test(t)), 'never 100% while the nodes are inserted');
    });
  }, { timeout: 180000 });

  test('host updates while a big document loads do not restart the load', async (h) => {
    const host = heavyHost(makeDoc(800));
    await h.use(host, hooks, async (view) => {
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      await view.page.waitForFunction(() => window.__fsTestHooks.mdw && window.__fsTestHooks.mdw.state.progress, { timeout: 60000, polling: 10 });
      const session = await view.eval(() => window.__fsTestHooks.mdw.state.session);
      for (let i = 0; i < 3; i++) {
        host.version++; // the same text, a new version (formatter, autosave, other editor)
        await view.push(host.update());
        await new Promise((r) => setTimeout(r, 150));
      }
      await view.page.waitForFunction(() => window.__fsTestHooks.mdw.state.ready, { timeout: 90000, polling: 50 });
      await view.settle();
      assert.equal(await view.eval(() => window.__fsTestHooks.mdw.state.session), session, 'the load was not restarted');
      assert.equal(await view.eval(() => window.__fsTestHooks.mdw.state.version), host.version, 'the version is the newest');
    });
  }, { timeout: 180000 });

  test('a small document: the full editor (virtual cursor), no progress step is needed', async (h) => {
    const host = heavyHost(makeDoc(60));
    await h.use(host, hooks, async (view) => {
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      await view.page.waitForFunction(() => window.__fsTestHooks.mdw && window.__fsTestHooks.mdw.state.ready, { timeout: 60000, polling: 50 });
      await view.settle();
      assert.equal(await view.eval(() => window.__fsTestHooks.mdw.state.lite), false);
      assert.equal(await view.eval(() => Boolean(document.querySelector('.ProseMirror.virtual-cursor-enabled'))), true);
    });
  }, { timeout: 120000 });
});
