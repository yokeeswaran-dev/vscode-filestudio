// Regression tests for fixed 0.1.2 PDF view bugs (R3 hunt, E1), on small PDFs built in memory (fixtures.pdf) and
// shown by the real pdf.js of dist/ (worker from a blob: URL, as under the extension's CSP).
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const page = (n, extra = {}) => ({ lines: [`Page ${n} needle`, 'some text', `needle again on page ${n}`], ...extra });
const findPdf = () => H.pdfHost('find.pdf', () => F.pdf([page(1), page(2), page(3), page(4)]));

/** Waits until pdf.js shows a rendered page (or an error view). */
async function waitPdf(view, timeout = 20000) {
  await view.page.waitForFunction(() => document.querySelector('.fv-error') || (document.querySelector('.pdfViewer .page canvas') && document.querySelector('.pdfViewer .page[data-loaded="true"]')), { timeout, polling: 100 });
  await H.sleep(300);
}
const ui = (view) =>
  view.eval(() => ({
    zoom: document.querySelector('.pdf-zoom-label')?.textContent,
    find: document.querySelector('.pdf-find-status')?.textContent,
    hl: document.querySelectorAll('.textLayer .highlight').length,
    error: document.querySelector('.fv-error')?.textContent ?? null,
  }));
const key = async (view, k, mods = []) => {
  for (const m of mods) await view.page.keyboard.down(m);
  await view.page.keyboard.press(k);
  for (const m of [...mods].reverse()) await view.page.keyboard.up(m);
  await H.sleep(350);
};

H.suite('pdf view (0.1.2 regressions)', ({ test }) => {
  test('R3-006 a PDF with no pages shows an error view and logs it (not a blank grey view)', (h) =>
    h.use(H.pdfHost('zero.pdf', () => F.pdf([])), { ignoreErrors: /no pages/ }, async (view) => {
      await waitPdf(view);
      const s = await ui(view);
      assert.match(s.error || '', /has no pages/);
      const logs = await view.posted('log');
      assert.ok(logs.some((m) => m.level === 'error' && /no pages/.test(m.message)), JSON.stringify(logs));
    }));

  test('R3-037 a zoom that is not a preset (160%) is restored after a reload', (h) =>
    h.use(findPdf(), {}, async (view) => {
      await waitPdf(view);
      await view.page.focus('.pdf-container');
      await key(view, 'Equal', ['Control']);
      await key(view, 'Equal', ['Control']);
      await H.sleep(500);
      const a = (await ui(view)).zoom;
      assert.notEqual(a, '100%');
      await H.sleep(700);
      await view.reload();
      await waitPdf(view);
      await H.sleep(800);
      assert.equal((await ui(view)).zoom, a);
    }));

  test('R3-038 a position inside a tall page is restored on that page (not the previous one)', (h) =>
    h.use(H.pdfHost('tall.pdf', () => F.pdf([page(1), page(2), page(3), page(4, { h: 2400 }), page(5)])), {}, async (view) => {
      await waitPdf(view);
      const where = () =>
        view.eval(() => {
          const c = document.querySelector('.pdf-container');
          const ct = c.getBoundingClientRect().top;
          const at = [...document.querySelectorAll('.pdfViewer .page')].find((e) => e.getBoundingClientRect().bottom > ct + 1);
          return { topPage: at && at.dataset.pageNumber, offset: at && Math.round(ct - at.getBoundingClientRect().top) };
        });
      await view.eval(() => {
        const c = document.querySelector('.pdf-container');
        c.scrollTop = document.querySelector('.pdfViewer .page[data-page-number="4"]').offsetTop + 601;
      });
      await H.sleep(1200);
      const a = await where();
      assert.equal(a.topPage, '4');
      await view.reload();
      await waitPdf(view);
      await H.sleep(1500);
      const b = await where();
      assert.equal(b.topPage, '4', JSON.stringify({ a, b }));
      assert.ok(Math.abs(b.offset - a.offset) <= 30, JSON.stringify({ a, b }));
    }));

  test('R3-039 / R3-040 find: F3 / Shift+F3 / Ctrl+G move through the matches; Ctrl+F after Escape shows highlights again', (h) =>
    h.use(findPdf(), {}, async (view) => {
      await waitPdf(view);
      await view.page.focus('.pdf-container');
      await key(view, 'KeyF', ['Control']);
      await view.page.keyboard.type('needle');
      await H.sleep(1000);
      const a = (await ui(view)).find;
      assert.match(a, /^1 of 8/);
      await key(view, 'F3');
      assert.match((await ui(view)).find, /^2 of 8/);
      await key(view, 'F3', ['Shift']);
      assert.match((await ui(view)).find, /^1 of 8/);
      await view.page.focus('.pdf-container');
      await key(view, 'F3');
      assert.match((await ui(view)).find, /^2 of 8/, 'F3 with the pages focused');
      await key(view, 'KeyG', ['Control']);
      assert.match((await ui(view)).find, /^3 of 8/, 'Ctrl+G');
      await key(view, 'Escape');
      await key(view, 'KeyF', ['Control']);
      await H.sleep(500);
      const s = await ui(view);
      assert.ok(!s.find || s.hl > 0, `status "${s.find}" with ${s.hl} highlights`);
    }));

  test('R3-041 PageDown at a high zoom scrolls one screen (no content skipped)', (h) =>
    h.use(findPdf(), {}, async (view) => {
      await waitPdf(view);
      await view.page.focus('.pdf-container');
      for (let i = 0; i < 3; i++) await key(view, 'Equal', ['Control']);
      await view.eval(() => (document.querySelector('.pdf-container').scrollTop = 0));
      await H.sleep(600);
      const m = () => view.eval(() => ({ top: Math.round(document.querySelector('.pdf-container').scrollTop), view: document.querySelector('.pdf-container').clientHeight }));
      const a = await m();
      await key(view, 'PageDown');
      await H.sleep(300);
      const b = await m();
      assert.ok(b.top > a.top, JSON.stringify({ a, b }));
      assert.ok(b.top - (a.top + a.view) <= 20, `content never shown: ${JSON.stringify({ a, b })}`);
    }));

  test('R3-059 a narrow editor (260 / 300 px) clips no toolbar control', async (h) => {
    for (const width of [260, 300]) {
      await h.use(findPdf(), { width, height: 600 }, async (view) => {
        await waitPdf(view);
        const clipped = await view.eval(() => {
          const tb = document.querySelector('.pdf-toolbar').getBoundingClientRect();
          return [...document.querySelectorAll('.pdf-toolbar button, .pdf-toolbar input')]
            .filter((b) => getComputedStyle(b).display !== 'none' && b.offsetParent !== null)
            .filter((b) => {
              const r = b.getBoundingClientRect();
              return r.right > tb.right + 0.5 || r.width < 4;
            })
            .map((b) => b.getAttribute('aria-label') || b.title);
        });
        assert.deepEqual(clipped, [], `@${width}`);
      });
    }
  });

  test('R3-060 the toolbar is a toolbar for screen readers; its buttons have unique names', (h) =>
    h.use(findPdf(), {}, async (view) => {
      await waitPdf(view);
      const tb = await view.page.$eval('.pdf-toolbar', (e) => ({ role: e.getAttribute('role'), label: e.getAttribute('aria-label') }));
      assert.deepEqual(tb, { role: 'toolbar', label: 'PDF toolbar' });
      const names = await view.eval(() =>
        [...document.querySelectorAll('.pdf-toolbar button')].filter((b) => b.offsetParent !== null).map((b) => (b.getAttribute('aria-label') || b.textContent || '').trim()),
      );
      assert.ok(names.every(Boolean), JSON.stringify(names));
      assert.deepEqual(names.filter((n, i) => names.indexOf(n) !== i), [], JSON.stringify(names));
      assert.match(await view.page.$eval('.pdf-zoom-label', (e) => e.getAttribute('aria-label')), /^Reset zoom/);
    }));

  test('E1-PDF-SHIFT fit width keeps a landscape page inside the view (no sideways scrolling)', (h) =>
    h.use(H.pdfHost('landscape.pdf', () => F.pdf([page(1), page(2), page(3, { w: 1100, h: 612 }), page(4)])), {}, async (view) => {
      await waitPdf(view);
      await view.page.click('.pdf-toolbar button[aria-label="Fit width"]');
      await H.sleep(800);
      const s = await view.eval(() => {
        const c = document.querySelector('.pdf-container');
        return { sw: c.scrollWidth, cw: c.clientWidth, left: c.scrollLeft };
      });
      assert.ok(s.sw <= s.cw + 1, JSON.stringify(s));
      assert.equal(s.left, 0);
    }));

  test('R3-070 a damaged PDF that pdf.js can repair opens without pdf.js warnings in the console', (h) => {
    // Every xref offset is 3 bytes wrong: pdf.js re-indexes the file (and, at WARNINGS verbosity, logs "Warning: ...").
    const damaged = async () => {
      const s = (await F.pdf([{ lines: ['Repaired xref'] }])).toString('latin1');
      return Buffer.from(s.replace(/(\d{10}) 00000 n /g, (m, o) => `${String(Number(o) + 3).padStart(10, '0')} 00000 n `), 'latin1');
    };
    const warnings = [];
    const before = async (pg) => pg.on('console', (m) => /^warn/.test(m.type()) && warnings.push(m.text()));
    return h.use(H.pdfHost('damaged-xref.pdf', damaged), { before }, async (view) => {
      await waitPdf(view);
      await H.sleep(500);
      const s = await ui(view);
      assert.equal(s.error, null);
      assert.match(await view.eval(() => document.querySelector('.textLayer')?.textContent ?? ''), /Repaired xref/);
      assert.deepEqual(warnings, []);
    });
  });
});
