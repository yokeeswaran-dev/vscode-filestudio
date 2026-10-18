// Regression tests for fixed 0.1.x Word (docx) view bugs, on small .docx files built in memory (fixtures.docx) and
// rendered by the real renderDocx (mammoth). Also H7-F12: messages for another view kind are dropped.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const F = require('./fixtures');

const LONG_URL = 'https://example.com/' + 'averyveryverylongpathsegment'.repeat(12);
const WORDS = ['Quantity', 'Description', 'Unit price', 'Discount', 'Shipping', 'Warehouse', 'Supplier', 'Category'];
const cell = (text) => `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells) => `<w:tr>${cells.map(cell).join('')}</w:tr>`;

const BODY = [
  F.para('Chapter 1\t1'),
  '<w:p><w:r><w:t>Name</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">Value  two</w:t></w:r></w:p>',
  F.para('שלום עולם, זה טקסט בעברית לגמרי.', '<w:bidi/>'),
  F.para('مرحبا بالعالم، هذا نص عربي.', '<w:bidi/>'),
  F.para('שלום is a Hebrew word at the start of an English sentence.'),
  F.para('مرحبا starts this English paragraph with an Arabic word.'),
  F.para('An ordinary English paragraph.'),
  '<w:p><w:hyperlink r:id="rIdL1"><w:r><w:t>C:\\Windows\\win.ini</w:t></w:r></w:hyperlink></w:p>',
  '<w:p><w:hyperlink r:id="rIdL2"><w:r><w:t>C:/Windows/win.ini</w:t></w:r></w:hyperlink></w:p>',
  '<w:p><w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> done</w:t></w:r></w:p>',
  '<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>highlighted text</w:t></w:r></w:p>',
  '<w:p><w:hyperlink w:anchor="_Toc100"><w:r><w:t>Chapter 2 (TOC entry)</w:t></w:r></w:hyperlink></w:p>',
  ...Array.from({ length: 30 }, (_, i) => F.para(`Filler paragraph ${i}.`)),
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc100"/><w:r><w:t>Chapter 2</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>',
  '<w:p><w:hyperlink r:id="rIdL3"><w:r><w:t>link after target</w:t></w:r></w:hyperlink></w:p>',
  ...Array.from({ length: 30 }, (_, i) => F.para(`Tail paragraph ${i}.`)),
].join('');
const LINKS = [
  { id: 'rIdL1', target: 'C:\\Windows\\win.ini' },
  { id: 'rIdL2', target: 'C:/Windows/win.ini' },
  { id: 'rIdL3', target: 'https://ok.example/after' },
];
const doc = () => H.docxHost('regressions.docx', () => F.docx(BODY, LINKS));
/** An 8-column table of ordinary words, and a table with one long URL. */
const WORDS_TABLE = `<w:tbl>${row(WORDS)}${row(WORDS.map((w) => `${w} value`))}</w:tbl>`;
const URL_TABLE = `<w:tbl>${row(['Link', LONG_URL])}${row(['Note', 'short words only'])}</w:tbl>`;

/** Paragraph text directions and the gaps between text and paragraph edges. */
function directions() {
  return [...document.querySelectorAll('.docx-page p')]
    .filter((p) => p.textContent.trim())
    .map((p) => {
      const rg = document.createRange();
      rg.selectNodeContents(p);
      const tr = rg.getBoundingClientRect();
      const pr = p.getBoundingClientRect();
      const cs = getComputedStyle(p);
      return { text: p.textContent.slice(0, 20), dir: cs.direction, gapL: Math.round(tr.left - pr.left - parseFloat(cs.paddingLeft)), gapR: Math.round(pr.right - parseFloat(cs.paddingRight) - tr.right) };
    });
}

/** Words split over two lines inside table cells; tables wider than the page; horizontal scrolling. */
function tableLayout() {
  const split = [];
  document.querySelectorAll('.docx-page td, .docx-page th').forEach((td) => {
    const walker = document.createTreeWalker(td, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      for (const m of n.textContent.matchAll(/\S+/g)) {
        const r = document.createRange();
        r.setStart(n, m.index);
        r.setEnd(n, m.index + m[0].length);
        if (new Set(Array.from(r.getClientRects()).map((x) => Math.round(x.top))).size > 1) split.push(m[0].slice(0, 30));
      }
    }
  });
  const pg = document.querySelector('.docx-page');
  const v = document.querySelector('.docx-view');
  const right = pg.getBoundingClientRect().right - parseFloat(getComputedStyle(pg).paddingRight) + 1;
  return { split, over: [...pg.querySelectorAll('table')].filter((t) => t.getBoundingClientRect().right > right).length, hScroll: v.scrollWidth > v.clientWidth };
}

H.suite('docx view (0.1.x regressions)', ({ test }) => {
  test('H7-F2 a table with a long URL stays on the page: the URL wraps inside its cell, nothing else does', (h) =>
    h.use(H.docxHost('url-table.docx', () => F.docx(F.para('Table:') + URL_TABLE)), { width: 600 }, async (view) => {
      const t = await view.eval(tableLayout);
      assert.equal(t.over, 0, JSON.stringify(t));
      assert.equal(t.hScroll, false, JSON.stringify(t));
      assert.ok(t.split.length > 0 && t.split.every((w) => w.startsWith('https://example.com/')), `words split: ${JSON.stringify(t.split)}`);
    }));

  test("R2H3-G words in table cells are not split mid-word ('Qua/ntit/y'), the table stays on the page", (h) =>
    h.use(H.docxHost('words-table.docx', () => F.docx(F.para('Table:') + WORDS_TABLE)), { width: 1000 }, async (view) => {
      const t = await view.eval(tableLayout);
      assert.deepEqual(t, { split: [], over: 0, hScroll: false });
    }));

  test('H7-F3 / R2H3-H right-to-left paragraphs are RTL; English paragraphs starting with a Hebrew / Arabic word stay LTR', (h) =>
    h.use(doc(), { width: 1000 }, async (view) => {
      const d = await view.eval(directions);
      const rtl = d.filter((p) => /^[\u0590-\u08ff]/.test(p.text) && !/[A-Za-z]/.test(p.text));
      const mixed = d.filter((p) => /^[\u0590-\u08ff]/.test(p.text) && /[A-Za-z]/.test(p.text));
      assert.equal(rtl.length, 2, JSON.stringify(d));
      assert.ok(rtl.every((p) => p.dir === 'rtl' && p.gapR <= 2 && p.gapL > 20), JSON.stringify(rtl));
      assert.equal(mixed.length, 2, JSON.stringify(d));
      assert.ok(mixed.every((p) => p.dir === 'ltr' && p.gapL <= 2), JSON.stringify(mixed));
    }));

  test('H7-F4 tabs and runs of spaces are kept (white-space: pre-wrap)', (h) =>
    h.use(doc(), {}, async (view) => {
      const p = await view.eval(() => {
        const el = [...document.querySelectorAll('.docx-page p')].find((x) => x.textContent.startsWith('Name'));
        return { text: el.textContent, ws: getComputedStyle(el).whiteSpace };
      });
      assert.equal(p.text, 'Name\tValue  two');
      assert.match(p.ws, /^pre-wrap$|^break-spaces$/);
    }));

  test('H7-F9 a Word check box cannot be toggled in the read-only view', (h) =>
    h.use(doc(), {}, async (view) => {
      await view.eval(() => document.querySelector('.docx-page input[type="checkbox"]').scrollIntoView({ block: 'center' }));
      await view.page.click('.docx-page input[type="checkbox"]');
      assert.equal(await view.eval(() => document.querySelector('.docx-page input[type="checkbox"]').checked), true);
    }));

  test('H7-F10 high contrast: Word highlights are outlined, the jump flash has a visible outline colour', (h) =>
    h.use(doc(), { theme: 'hc', height: 600 }, async (view) => {
      const mark = await view.eval(() => {
        const cs = getComputedStyle(document.querySelector('.docx-page mark'));
        return { style: cs.outlineStyle, width: cs.outlineWidth, color: cs.outlineColor };
      });
      assert.equal(mark.style, 'dashed');
      assert.equal(mark.width, '1px');
      assert.notEqual(mark.color, 'rgba(0, 0, 0, 0)');
      await view.eval(() => [...document.querySelectorAll('.docx-page a')].find((a) => a.textContent === 'Chapter 2 (TOC entry)').click());
      await view.page.waitForSelector('.md-target-flash', { timeout: 5000 });
      // The flash animation is stopped at its first frame: the outline must really be drawn there (colour, style and
      // width), not only the --md-flash-outline custom property be set.
      const flash = await view.eval(() => {
        const el = document.querySelector('.md-target-flash');
        const anims = el.getAnimations();
        for (const a of anims) {
          a.pause();
          a.currentTime = 0;
        }
        const cs = getComputedStyle(el);
        return { anims: anims.length, color: cs.outlineColor, style: cs.outlineStyle, width: parseFloat(cs.outlineWidth) };
      });
      assert.ok(flash.anims >= 1, 'the jump target flashes');
      assert.notEqual(flash.color, 'rgba(0, 0, 0, 0)', `flash outline colour ${flash.color}`);
      assert.notEqual(flash.style, 'none');
      assert.ok(flash.width >= 1, `flash outline width ${flash.width}`);
    }));

  test('R2H3-D Word links to drive-letter paths (C:\\..., C:/...) stay links and go to the host', (h) =>
    h.use(doc(), {}, async (view) => {
      await view.clearPosted();
      for (const text of ['C:\\Windows\\win.ini', 'C:/Windows/win.ini']) {
        await view.eval((t) => {
          const a = [...document.querySelectorAll('.docx-page a')].find((x) => x.textContent === t);
          a.scrollIntoView({ block: 'center' });
          a.click();
        }, text);
        await H.sleep(100);
      }
      assert.deepEqual((await view.posted('openLink')).map((m) => m.href), ['C:\\Windows\\win.ini', 'C:/Windows/win.ini']);
    }));

  test('R2H3-E a Word TOC jump highlights the bookmarked heading and moves the focus there; Tab continues after it', (h) =>
    h.use(doc(), { height: 600 }, async (view) => {
      await view.eval(() => [...document.querySelectorAll('.docx-page a')].find((a) => a.textContent === 'Chapter 2 (TOC entry)').focus());
      await view.page.keyboard.press('Enter');
      await H.sleep(300);
      const flash = await view.eval(() => {
        const f = document.querySelector('.md-target-flash');
        const r = f && f.getBoundingClientRect();
        return { el: f ? f.localName : null, visible: !!r && r.width > 0 && r.height > 0 && r.top >= -2 && r.bottom <= innerHeight };
      });
      assert.deepEqual(flash, { el: 'h1', visible: true });
      await view.page.keyboard.press('Tab');
      await H.sleep(100);
      assert.equal(await view.eval(() => document.activeElement.textContent.slice(0, 40)), 'link after target');
    }));

  test('R2H4-5 Reduce Motion: no docx warnings chevron animation', (h) =>
    h.use(doc(), {}, async (view) => {
      const motion = await view.eval(() => {
        const s = document.querySelector('.docx-warnings summary');
        const off = getComputedStyle(s, '::after').transitionDuration;
        document.body.classList.add('vscode-reduce-motion');
        return { off, on: getComputedStyle(s, '::after').transitionDuration };
      });
      assert.deepEqual(motion, { off: '0.12s', on: '0s' });
    }));

  test('H7-F12 a markdownUpdate posted to a docx view is ignored; an init of an unknown kind shows an error', (h) =>
    h.use(doc(), { ignoreErrors: /unknown kind 'gslides'/ }, async (view) => {
      await view.push({ type: 'markdownUpdate', html: '<h1>MD!</h1>', toc: [], source: '# MD!', version: 99 });
      await H.sleep(150);
      const state = await view.eval(() => ({ docx: !!document.querySelector('.docx-view'), md: !!document.querySelector('.md-view'), kind: document.body.dataset.kind }));
      assert.deepEqual(state, { docx: true, md: false, kind: 'docx' });
      const logged = await view.posted('log');
      assert.ok(logged.some((m) => /Ignored a 'markdownUpdate' message/.test(m.message)), JSON.stringify(logged));
      await view.push({ type: 'init', kind: 'gslides', fileName: 'x.gslides' });
      await H.sleep(150);
      const unknown = await view.eval(() => ({ error: (document.querySelector('.fv-error-message') || {}).textContent || '', kind: document.body.dataset.kind, stale: !!document.querySelector('.docx-view') }));
      assert.match(unknown.error, /gslides/);
      assert.equal(unknown.kind, 'docx');
      assert.equal(unknown.stale, false);
    }));
});
