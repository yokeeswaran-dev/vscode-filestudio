// Markdown modes (v0.2.0, P6): toolbar (radiogroup, keys, disabled WYSIWYG), TOC tree, Split scroll sync, the WYSIWYG
// editor (0 edits after load, typing -> one edit, keys -> commands, flush, protected blocks, read-only, fallbacks).
// A mock host answers like MarkdownView in src/viewerProvider.ts, with the real renderers (markdownBlocks,
// applyMdTextEdits, verifyMdEdit, renderMarkdown).
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const H = require('./harness');
const { mdHost } = require('./md-host');

const ROOT = path.join(__dirname, '..', '..');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const DOC = [
  '# Title',
  '',
  'First paragraph with **bold** and a [link](https://example.com/).',
  '',
  '## Section A',
  '',
  '- item one',
  '- item two',
  '',
  '### Sub A1',
  '',
  'Text under A1.',
  '',
  '## Section B',
  '',
  '```js',
  'const a = 1;',
  '```',
  '',
  '<div class="raw">raw html block</div>',
  '',
  'Last paragraph.',
  '',
].join('\n');

const radios = (view) => view.eval(() => [...document.querySelectorAll('.md-mode-btn')].map((b) => ({ mode: b.dataset.mode, checked: b.getAttribute('aria-checked'), disabled: b.getAttribute('aria-disabled'), tab: b.tabIndex, title: b.title })));
const status = (view) => view.eval(() => document.querySelector('.md-status')?.textContent ?? '');
const waitEditor = (view) => view.page.waitForFunction(() => document.querySelector('.md-wysiwyg:not([hidden]) .milkdown .ProseMirror'), { timeout: 20000, polling: 50 });
const toWysiwyg = async (view) => {
  await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
  await waitEditor(view);
  await view.settle();
};
const sendKey = async (view, ...keys) => {
  for (const k of keys) await view.page.keyboard.press(k);
};
/** Caret at the end of the first block containing `text` inside the editor. */
async function caretIn(view, text, selector = 'p') {
  await view.page.evaluate(
    (text, selector) => {
      const el = [...document.querySelectorAll(`.md-wysiwyg .ProseMirror ${selector}`)].find((e) => e.textContent.includes(text));
      if (!el) throw new Error(`no ${selector} with ${text}`);
      const range = document.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      document.querySelector('.md-wysiwyg .ProseMirror').focus();
    },
    text,
    selector,
  );
}

H.suite('markdown modes (P6)', ({ test }) => {
  // ----- toolbar -----
  test('toolbar: a radiogroup Preview | Split | WYSIWYG, Preview checked, one tab stop; the TOC button is a toggle', async (h) =>
    h.use(mdHost('modes.md', DOC), {}, async (view) => {
      const group = await view.eval(() => {
        const g = document.querySelector('.md-modes');
        return { role: g.getAttribute('role'), label: g.getAttribute('aria-label'), n: g.querySelectorAll('[role=radio]').length };
      });
      assert.deepEqual(group, { role: 'radiogroup', label: 'Markdown view', n: 3 });
      assert.deepEqual((await radios(view)).map((r) => [r.mode, r.checked, r.tab]), [['preview', 'true', 0], ['split', 'false', -1], ['wysiwyg', 'false', -1]]);
      const toc = await view.eval(() => ({ pressed: document.querySelector('.md-toc-btn').getAttribute('aria-pressed'), hidden: document.querySelector('.md-toc').hidden, controls: document.querySelector('.md-toc-btn').getAttribute('aria-controls') }));
      assert.deepEqual(toc, { pressed: 'false', hidden: true, controls: 'md-toc' });
      assert.equal(await view.eval(() => getComputedStyle(document.querySelector('.md-view')).position), 'relative');
      assert.equal(await view.eval(() => getComputedStyle(document.querySelector('.md-app')).position), 'fixed');
    }));

  test('toolbar keys: arrows move the focus and choose the mode (roving tabindex), Home / End jump; the host is told', async (h) => {
    const host = mdHost('keys.md', DOC);
    return h.use(host, {}, async (view) => {
      await view.page.focus('.md-mode-btn[data-mode="preview"]');
      await sendKey(view, 'ArrowRight');
      await view.settle();
      assert.deepEqual((await radios(view)).map((r) => [r.mode, r.checked, r.tab]), [['preview', 'false', -1], ['split', 'true', 0], ['wysiwyg', 'false', -1]]);
      assert.equal(await view.eval(() => document.activeElement.dataset.mode), 'split');
      assert.deepEqual(host.modes.slice(-1), ['split']);
      await sendKey(view, 'Home');
      await view.settle();
      assert.equal((await radios(view))[0].checked, 'true');
      assert.deepEqual(host.modes.slice(-1), ['preview']);
      await sendKey(view, 'ArrowLeft'); // wraps to WYSIWYG
      await waitEditor(view);
      assert.equal((await radios(view))[2].checked, 'true');
      assert.equal(await view.eval(() => document.querySelector('.md-app').dataset.mode), 'wysiwyg');
    });
  });

  test('a disabled WYSIWYG (read-only file) stays focusable with aria-disabled and says why; nothing is sent', async (h) => {
    const reason = 'This file is read-only (files.readonlyInclude).';
    const host = mdHost('ro.md', DOC, { editState: { editable: false, code: 'readOnly', reason, kinds: [] } });
    return h.use(host, {}, async (view) => {
      const wys = (await radios(view))[2];
      assert.equal(wys.disabled, 'true');
      assert.match(wys.title, /read-only/);
      await view.page.focus('.md-mode-btn[data-mode="wysiwyg"]');
      assert.equal(await view.eval(() => document.activeElement.dataset.mode), 'wysiwyg');
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      assert.equal(await status(view), reason);
      assert.equal((await radios(view))[0].checked, 'true', 'still Preview');
      assert.deepEqual(host.modes, []);
      assert.equal(await view.eval(() => document.querySelector('.md-status').getAttribute('aria-live')), 'polite');
    });
  });

  test('too large for WYSIWYG: the button is off with the host\'s reason; Split still works', async (h) => {
    const reason = 'This file is too large for WYSIWYG editing. Use Split to edit the Markdown source.';
    const host = mdHost('big.md', DOC, { limits: { wysiwyg: false, reason } });
    return h.use(host, {}, async (view) => {
      assert.equal((await radios(view))[2].disabled, 'true');
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      assert.equal(await status(view), reason);
      await view.page.click('.md-mode-btn[data-mode="split"]');
      await view.settle();
      assert.deepEqual(host.modes, ['split']);
    });
  });

  test('the mode chosen on the toolbar is kept in the webview state and used after a reload', async (h) => {
    const host = mdHost('state.md', DOC);
    return h.use(host, {}, async (view) => {
      await view.page.click('.md-mode-btn[data-mode="split"]');
      await view.settle();
      host.mode = 'preview'; // a new panel of the host starts in its default mode again
      host.modes.length = 0;
      await view.reload();
      assert.equal((await radios(view))[1].checked, 'true');
      assert.deepEqual(host.modes, ['split'], 'the host is told about the mode the panel is really in');
    });
  });

  // ----- TOC -----
  test('TOC: a tree of the headings (skipped levels nest), keys, collapse state kept in the state; click scrolls', async (h) => {
    const text = '# One\n\ntext\n\n### Deep (skips h2)\n\ntext\n\n## Two\n\n' + 'filler\n\n'.repeat(60) + '## Three\n\n' + 'tail\n\n'.repeat(60);
    const host = mdHost('toc.md', text);
    return h.use(host, {}, async (view) => {
      await view.page.click('.md-toc-btn');
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), false);
      assert.equal(await view.eval(() => document.querySelector('.md-toc-btn').getAttribute('aria-pressed')), 'true');
      const tree = await view.eval(() =>
        [...document.querySelectorAll('.md-toc-row')].map((r) => [r.textContent, r.getAttribute('aria-level'), r.getAttribute('aria-expanded')]),
      );
      assert.deepEqual(tree, [['One', '1', 'true'], ['Deep (skips h2)', '2', null], ['Two', '2', null], ['Three', '2', null]]);
      assert.equal(await view.eval(() => document.querySelector('.md-toc-list').getAttribute('role')), 'tree');
      // keys
      await view.page.focus('.md-toc-row[data-key^="one@"]');
      await sendKey(view, 'ArrowDown');
      assert.match(await view.eval(() => document.activeElement.textContent), /Deep/);
      await sendKey(view, 'ArrowLeft'); // to the parent
      assert.equal(await view.eval(() => document.activeElement.textContent), 'One');
      await sendKey(view, 'ArrowLeft'); // collapses it
      assert.equal(await view.eval(() => document.activeElement.getAttribute('aria-expanded')), 'false');
      assert.equal(await view.eval(() => [...document.querySelectorAll('.md-toc-row')].filter((r) => r.getClientRects().length).length), 1, 'the children are hidden');
      await sendKey(view, 'ArrowRight'); // expands it again
      assert.equal(await view.eval(() => document.activeElement.getAttribute('aria-expanded')), 'true');
      await sendKey(view, 'End');
      assert.equal(await view.eval(() => document.activeElement.textContent), 'Three');
      // Enter scrolls the preview to the heading and moves the highlight
      await sendKey(view, 'Enter');
      await sleep(200);
      assert.ok((await view.eval(() => document.querySelector('.md-view').scrollTop)) > 300);
      await view.page.waitForFunction(() => document.querySelector('.md-toc-row[aria-current="location"]')?.textContent === 'Three', { timeout: 4000 });
      // collapse One again: the saved state survives a rebuild (an update) and a reload
      await sendKey(view, 'Home');
      await sendKey(view, 'ArrowLeft');
      assert.equal(await view.eval(() => document.activeElement.getAttribute('aria-expanded')), 'false');
      host.text = host.text + '\nmore\n';
      host.version++;
      await view.push(host.update());
      await view.settle();
      assert.equal(await view.eval(() => document.querySelector('.md-toc-row[data-key^="one@"]').getAttribute('aria-expanded')), 'false');
      await view.reload();
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), false, 'open state kept');
      assert.equal(await view.eval(() => document.querySelector('.md-toc-row[data-key^="one@"]').getAttribute('aria-expanded')), 'false');
    });
  });

  test('TOC: Escape on an entry closes the panel (state saved) and puts the focus back on the TOC button', async (h) => {
    const host = mdHost('toc-esc.md', DOC);
    return h.use(host, {}, async (view) => {
      await view.page.click('.md-toc-btn');
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), false);
      await view.page.focus('.md-toc-row[data-key^="section-a@"]');
      await sendKey(view, 'ArrowDown'); // any entry
      assert.equal(await view.eval(() => document.activeElement.classList.contains('md-toc-row')), true, 'the focus is on an entry');
      await sendKey(view, 'Escape');
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), true, 'closed');
      assert.equal(await view.eval(() => document.querySelector('.md-toc-btn').getAttribute('aria-pressed')), 'false');
      assert.equal(await view.eval(() => document.activeElement === document.querySelector('.md-toc-btn')), true, 'focus on the toggle');
      await view.reload();
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), true, 'closed state kept');
      // Escape elsewhere (the toolbar) does nothing to the TOC
      await view.page.click('.md-toc-btn');
      await view.page.focus('.md-mode-btn[data-mode="preview"]');
      await sendKey(view, 'Escape');
      assert.equal(await view.eval(() => document.querySelector('.md-toc').hidden), false);
    });
  });

  // ----- Split -----
  test('Split: previewScrolled follows the preview scroll; revealLine scrolls the preview to the line (+-1 line); no echo', async (h) => {
    const text = Array.from({ length: 120 }, (_, i) => `Paragraph ${i} of the long document.`).join('\n\n') + '\n';
    const host = mdHost('split.md', text);
    return h.use(host, {}, async (view) => {
      await view.page.click('.md-mode-btn[data-mode="split"]');
      await view.settle();
      assert.deepEqual(host.modes, ['split']);
      // editor -> preview
      await view.push({ type: 'revealLine', line: 100 });
      await sleep(250);
      const topLine = await view.eval(() => {
        const view = document.querySelector('.md-view');
        const top = view.getBoundingClientRect().top;
        const blocks = [...document.querySelectorAll('.md-article > [data-source-line]')];
        const first = blocks.find((b) => b.getBoundingClientRect().bottom > top + 1);
        return Number(first.getAttribute('data-source-line'));
      });
      assert.ok(Math.abs(topLine - 100) <= 1, `the preview shows line ${topLine}`);
      await sleep(250);
      assert.deepEqual(host.previews, [], 'our own scroll is not sent back (no echo)');
      // preview -> editor
      await view.eval(() => (document.querySelector('.md-view').scrollTop = 1200));
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'previewScrolled'), { timeout: 4000, polling: 30 });
      const sent = (await view.posted('previewScrolled')).at(-1).line;
      const expected = await view.eval(() => {
        const view = document.querySelector('.md-view');
        const top = view.getBoundingClientRect().top;
        const blocks = [...document.querySelectorAll('.md-article > [data-source-line]')];
        const first = blocks.find((b) => b.getBoundingClientRect().bottom > top + 1);
        return Number(first.getAttribute('data-source-line'));
      });
      assert.ok(Math.abs(sent - expected) <= 1, `sent ${sent}, the preview shows line ${expected}`);
      // outside Split nothing is sent
      await view.page.click('.md-mode-btn[data-mode="preview"]');
      await view.settle();
      await view.clearPosted();
      await view.eval(() => (document.querySelector('.md-view').scrollTop = 200));
      await sleep(200);
      assert.deepEqual(await view.posted('previewScrolled'), []);
    });
  });

  test('Split on a large document whose blocks are still being built: revealLine builds the rest first and lands within 1 line, also after an earlier scroll; no echo', async (h) => {
    const text = Array.from({ length: 3500 }, (_, i) => `Paragraph ${i} lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.`).join('\n\n') + '\n';
    const host = mdHost('big-split.md', text, { mode: 'split' });
    return h.use(host, {}, async (view) => {
      await view.page.waitForFunction(() => document.querySelectorAll('.md-article > [data-source-line]').length > 5, { timeout: 20000, polling: 20 });
      const built = () => view.eval(() => document.querySelectorAll('.md-article > p[data-source-line]').length);
      // a scroll of the preview fills the element list with the blocks that exist so far
      await view.eval(() => (document.querySelector('.md-view').scrollTop = 40));
      await sleep(90);
      const before = await built();
      host.previews.length = 0; // (the scroll above is the user's: the host got it)
      const topLine = () =>
        view.eval(() => {
          const view = document.querySelector('.md-view');
          const top = view.getBoundingClientRect().top;
          const first = [...document.querySelectorAll('.md-article > [data-source-line]')].find((b) => b.getBoundingClientRect().bottom > top + 1);
          return Number(first.getAttribute('data-source-line'));
        });
      // the paragraphs are on the source lines 2 * i (a blank line between them); 6001 is in the middle of paragraph 3000
      await view.push({ type: 'revealLine', line: 6001 });
      await sleep(300);
      const after = await built();
      console.log(`large Split: ${before} of ${after} paragraphs built when revealLine came`);
      assert.equal(after, 3500, 'every block is built');
      const shown = await topLine();
      assert.ok(Math.abs(shown - 6001) <= 1, `the preview shows line ${shown} (wanted 6001)`);
      await sleep(250);
      assert.deepEqual(host.previews, [], 'our own scroll is not sent back (no echo)');
      // a line between two blocks well before the end (the end of a page cannot be scrolled to the top) and the very first line
      await view.push({ type: 'revealLine', line: 6801 });
      await sleep(250);
      assert.ok(Math.abs((await topLine()) - 6801) <= 1);
      await view.push({ type: 'revealLine', line: 0 });
      await sleep(250);
      assert.ok((await topLine()) <= 1);
      assert.deepEqual(host.previews, [], 'no echo after the jumps either');
      // preview -> editor on the same document: the line sent is the line at the top
      await view.clearPosted();
      await view.eval(() => (document.querySelector('.md-view').scrollTop = 300000));
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'previewScrolled'), { timeout: 4000, polling: 30 });
      const sent = (await view.posted('previewScrolled')).at(-1).line;
      const expected = await topLine();
      assert.ok(Math.abs(sent - expected) <= 1, `sent ${sent}, the preview shows line ${expected}`);
    });
  }, { timeout: 120000 });

  // ----- WYSIWYG -----
  test('WYSIWYG: the editor opens, protected blocks are raw text, nothing is sent (0 edits after load), back to Preview', async (h) => {
    const host = mdHost('wys.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      assert.deepEqual(host.modes, ['wysiwyg']);
      const info = await view.eval(() => ({
        h1: document.querySelector('.md-wysiwyg .ProseMirror h1')?.textContent,
        raw: [...document.querySelectorAll('.md-wysiwyg pre.fs-raw')].map((p) => [p.dataset.kind, p.textContent]),
        previewHidden: document.querySelector('.md-view').hidden,
        editable: document.querySelector('.md-wysiwyg .ProseMirror').getAttribute('contenteditable'),
      }));
      assert.equal(info.h1, 'Title');
      assert.deepEqual(info.raw, [['html_block', '<div class="raw">raw html block</div>']]);
      assert.equal(info.previewHidden, true);
      assert.equal(info.editable, 'true');
      await sleep(900);
      assert.deepEqual(await view.posted('edit'), [], '0 edits after load');
      assert.equal(host.text, DOC);
      await view.page.click('.md-mode-btn[data-mode="preview"]');
      await view.page.waitForFunction(() => !document.querySelector('.md-view').hidden && document.querySelector('.md-article h1'), { timeout: 5000 });
      assert.equal(await view.eval(() => document.querySelector('.md-wysiwyg').hidden), true);
      assert.equal(host.text, DOC);
    });
  });

  test('WYSIWYG: typing a word is ONE edit that changes only that word; the answer and the update keep the editor in step', async (h) => {
    const host = mdHost('type.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await view.page.keyboard.type(' typed');
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await view.settle();
      assert.equal(host.text, DOC.replace('Last paragraph.', 'Last paragraph. typed'));
      assert.equal(host.edits.length, 1);
      assert.deepEqual(host.edits[0].map((e) => e.text), [' typed']);
      assert.deepEqual(host.refused, []);
      // more typing continues from the accepted state
      await view.page.keyboard.type(' more');
      await view.page.waitForFunction(() => window.__posted.filter((m) => m.type === 'edit').length >= 2, { timeout: 4000, polling: 30 });
      await view.settle();
      assert.equal(host.text, DOC.replace('Last paragraph.', 'Last paragraph. typed more'));
    });
  });

  test('WYSIWYG: Enter sends at once; the text of a protected block is edited as plain text', async (h) => {
    const host = mdHost('enter.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'raw html block', 'pre.fs-raw');
      await view.page.keyboard.type('X');
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await view.settle();
      assert.equal(host.text, DOC.replace('raw html block</div>', 'raw html block</div>X'));
    });
  });

  test('WYSIWYG: Escape in the text puts the focus on the mode button (a way out of the editor); nothing is sent', async (h) => {
    const host = mdHost('esc.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await view.clearPosted();
      const where = () => view.eval(() => ({ mode: document.activeElement?.dataset?.mode ?? null, inEditor: Boolean(document.activeElement?.closest('.ProseMirror')) }));
      assert.deepEqual(await where(), { mode: null, inEditor: true });
      await sendKey(view, 'Escape');
      assert.deepEqual(await where(), { mode: 'wysiwyg', inEditor: false });
      await sleep(600);
      assert.deepEqual(await view.posted('edit'), [], 'no edit');
      assert.equal(host.text, DOC);
      // Tab from the mode button is the way back (roving tabindex: the chosen mode is the tab stop)
      assert.equal(await view.eval(() => document.activeElement.tabIndex), 0);
    });
  });

  test('WYSIWYG: Escape that something else in the editor already used (a menu, a link box) is left alone', async (h) => {
    const host = mdHost('esc2.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      // like the block menu of the editor: a window listener in the capture phase takes the Escape
      await view.eval(() => window.addEventListener('keydown', (e) => { if (e.key === 'Escape') e.preventDefault(); }, true));
      await sendKey(view, 'Escape');
      assert.equal(await view.eval(() => Boolean(document.activeElement?.closest('.ProseMirror'))), true, 'the caret stays in the text');
    });
  });

  test('WYSIWYG: Escape closes the block menu ("/") and keeps the caret in the text; the next Escape leaves the text', async (h) => {
    const host = mdHost('esc3.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await sendKey(view, 'Enter'); // a new empty paragraph
      await view.page.keyboard.type('/');
      const menuShown = () => view.eval(() => [...document.querySelectorAll('.milkdown [data-show="true"]')].some((el) => el.getClientRects().length > 0));
      await view.page.waitForFunction(() => [...document.querySelectorAll('.milkdown [data-show="true"]')].some((el) => el.getClientRects().length > 0), { timeout: 4000, polling: 50 });
      assert.equal(await menuShown(), true, 'the block menu is open');
      await sendKey(view, 'Escape');
      await sleep(200);
      assert.equal(await menuShown(), false, 'the menu is closed');
      assert.equal(await view.eval(() => Boolean(document.activeElement?.closest('.ProseMirror'))), true, 'the caret stays in the text');
      await sendKey(view, 'Escape');
      assert.equal(await view.eval(() => document.activeElement?.dataset?.mode ?? null), 'wysiwyg', 'the next Escape leaves the text');
    });
  });

  test('WYSIWYG keys: Ctrl+Z, Ctrl+Y, Ctrl+Shift+Z and Ctrl+S flush, then ask the host (command); the editor has no undo of its own', async (h) => {
    const host = mdHost('keys2.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await view.page.keyboard.type(' typed');
      // Ctrl+S right after typing: the edit is sent first, then the command.
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('s');
      await view.page.keyboard.up('Control');
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'command'), { timeout: 4000, polling: 30 });
      const order = (await view.posted()).filter((m) => m.type === 'edit' || m.type === 'command').map((m) => m.type + (m.id ? `:${m.id}` : ''));
      assert.deepEqual(order, ['edit', 'command:save']);
      assert.equal(host.text, DOC.replace('Last paragraph.', 'Last paragraph. typed'));
      // Ctrl+Z: the host undoes (the editor does not)
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('z');
      await view.page.keyboard.up('Control');
      await view.page.waitForFunction(() => window.__posted.filter((m) => m.type === 'command').length >= 2, { timeout: 4000, polling: 30 });
      await view.settle();
      assert.deepEqual(host.commands, ['save', 'undo']);
      assert.equal(host.text, DOC, 'undone by the host');
      await view.page.waitForFunction(() => !document.querySelector('.md-wysiwyg .ProseMirror').textContent.includes('typed'), { timeout: 4000 });
      // Ctrl+Shift+Z and Ctrl+Y redo
      await view.page.keyboard.down('Control');
      await view.page.keyboard.down('Shift');
      await view.page.keyboard.press('z');
      await view.page.keyboard.up('Shift');
      await view.page.keyboard.up('Control');
      await view.page.waitForFunction(() => window.__posted.filter((m) => m.type === 'command').length >= 3, { timeout: 4000, polling: 30 });
      await view.page.keyboard.down('Control');
      await view.page.keyboard.press('y');
      await view.page.keyboard.up('Control');
      await view.page.waitForFunction(() => window.__posted.filter((m) => m.type === 'command').length >= 4, { timeout: 4000, polling: 30 });
      assert.deepEqual(host.commands, ['save', 'undo', 'redo', 'redo']);
    });
  });

  test('WYSIWYG: flushRequest (Save All) sends the pending typing, then "flushed"; with nothing pending only "flushed"', async (h) => {
    const host = mdHost('flush.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await view.clearPosted();
      await view.push({ type: 'flushRequest', id: 7 });
      await view.settle();
      assert.deepEqual((await view.posted()).map((m) => m.type), ['flushed']);
      await caretIn(view, 'Last paragraph');
      await view.page.keyboard.type('!');
      await view.clearPosted();
      await view.push({ type: 'flushRequest', id: 8 });
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'flushed'), { timeout: 4000, polling: 30 });
      assert.deepEqual((await view.posted()).map((m) => m.type), ['edit', 'flushed']);
      assert.equal(host.text, DOC.replace('Last paragraph.', 'Last paragraph.!'));
    });
  });

  test('WYSIWYG: an edit the host refuses shows its reason and the editor shows the file again', async (h) => {
    const host = mdHost('refuse.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      host.version += 5; // the host's document moved on: the edit is stale
      await caretIn(view, 'Last paragraph');
      await view.page.keyboard.type('Q');
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await view.settle();
      assert.equal(host.refused.length, 1);
      assert.match(await status(view), /The file changed while you were editing/);
      assert.equal(host.text, DOC);
      assert.ok(!(await view.eval(() => document.querySelector('.md-wysiwyg .ProseMirror').textContent)).includes('Q'), 'the typed text is gone: the editor shows the file');
    });
  });

  test('WYSIWYG: a change of the file by someone else (another panel, undo) reloads the editor', async (h) => {
    const host = mdHost('external.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      host.text = DOC.replace('Last paragraph.', 'Changed elsewhere.');
      host.version++;
      await view.push(host.update());
      await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg .ProseMirror').textContent.includes('Changed elsewhere.'), { timeout: 4000 });
      assert.deepEqual(await view.posted('edit'), []);
    });
  });

  test('WYSIWYG read-only while the save check is open ("Edit anyway"): not typable, the button asks the host', async (h) => {
    const canEditAnyway = { title: 'Saving in VS Code changes parts of this file you did not edit.', details: ['VS Code will change 1 line break (LF → CRLF).'] };
    const host = mdHost('fidelity.md', DOC, { editState: { editable: true, kinds: [], code: 'fidelity', reason: 'Saving in VS Code would also change parts of this file you did not edit. Choose "Edit anyway" to edit it.', canEditAnyway } });
    return h.use(host, {}, async (view) => {
      assert.notEqual((await radios(view))[2].disabled, 'true', 'WYSIWYG is shown, read-only');
      await toWysiwyg(view);
      assert.equal(await view.eval(() => document.querySelector('.md-wysiwyg .ProseMirror').getAttribute('contenteditable')), 'false');
      assert.equal(await view.eval(() => document.querySelector('.md-anyway').hidden), false);
      assert.match(await status(view), /Read-only/);
      await view.page.click('.md-anyway');
      await view.settle();
      assert.equal(host.anyway, 1);
      // the host accepts: editable now
      host.editState = { editable: true, kinds: ['mdText'] };
      await view.push({ type: 'editState', state: host.editState });
      await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg .ProseMirror').getAttribute('contenteditable') === 'true', { timeout: 4000 });
      assert.equal(await view.eval(() => document.querySelector('.md-anyway').hidden), true);
    });
  });

  test('WYSIWYG: turning editing off while the editor is open sends it back to Preview with the reason', async (h) => {
    const host = mdHost('killswitch.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      const reason = 'Editing is turned off (setting fileStudio.markdown.editing).';
      host.editState = { editable: false, code: 'disabled', reason, kinds: [] };
      await view.push({ type: 'editState', state: host.editState });
      await view.page.waitForFunction(() => !document.querySelector('.md-view').hidden, { timeout: 4000 });
      assert.equal(await status(view), reason);
      assert.equal((await radios(view))[0].checked, 'true');
      assert.equal((await radios(view))[2].disabled, 'true');
    });
  });

  test('WYSIWYG: Ctrl+click on a link asks the host to open it; a plain click only places the caret', async (h) => {
    const host = mdHost('links.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await view.clearPosted();
      const link = await view.eval(() => {
        const r = document.querySelector('.md-wysiwyg .ProseMirror a').getBoundingClientRect();
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
      });
      await view.page.mouse.click(link.x, link.y);
      await sleep(200);
      assert.deepEqual(await view.posted('openLink'), []);
      await view.page.keyboard.down('Control');
      await view.page.mouse.click(link.x, link.y);
      await view.page.keyboard.up('Control');
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'openLink'), { timeout: 4000, polling: 30 });
      assert.deepEqual((await view.posted('openLink')).map((m) => m.href), ['https://example.com/']);
    });
  });

  test('WYSIWYG: pasted HTML is sanitized; pasted files are refused with a message', async (h) => {
    const host = mdHost('paste.md', DOC);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await view.eval(() => {
        const dt = new DataTransfer();
        dt.setData('text/html', '<p>safe <b onclick="window.__pwned=1">bold</b><script>window.__pwned=2</script>text</p>');
        dt.setData('text/plain', 'safe text');
        const event = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
        document.querySelector('.md-wysiwyg .ProseMirror').dispatchEvent(event);
      });
      await sleep(300);
      assert.equal(await view.eval(() => window.__pwned), undefined);
      await view.eval(() => {
        const dt = new DataTransfer();
        dt.items.add(new File(['x'], 'a.png', { type: 'image/png' }));
        document.querySelector('.md-wysiwyg .ProseMirror').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      });
      await sleep(100);
      assert.match(await status(view), /Pasting files is not supported here; add the image as a Markdown link\./);
    });
  });

  test('WYSIWYG: a relative image is shown through the resource base; the Markdown keeps the path as written', async (h) => {
    const text = '# T\n\n![alt text](images/sample.png "A sample")\n\nafter\n';
    const host = mdHost('img.md', text);
    return h.use(host, {}, async (view) => {
      await toWysiwyg(view);
      await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg .ProseMirror img:not(.ProseMirror-separator)')?.complete, { timeout: 5000 });
      const img = await view.eval(() => {
        const i = document.querySelector('.md-wysiwyg .ProseMirror img:not(.ProseMirror-separator)');
        return { src: i.getAttribute('src'), alt: i.getAttribute('alt'), title: i.getAttribute('title'), loaded: i.naturalWidth > 0 };
      });
      assert.match(img.src, /^http:\/\/127\.0\.0\.1:\d+\/res\/images\/sample\.png$/);
      assert.deepEqual([img.alt, img.title, img.loaded], ['alt text', 'A sample', true]);
      await sleep(600);
      assert.deepEqual(await view.posted('edit'), []);
      assert.equal(host.text, text);
    });
  });

  test('WYSIWYG: no CSP violation and no eval while the editor opens and is used', async (h) =>
    h.use(mdHost('csp.md', DOC), {}, async (view) => {
      await toWysiwyg(view);
      await caretIn(view, 'Last paragraph');
      await view.page.keyboard.type(' x');
      await sleep(700);
      assert.deepEqual(await view.eval(() => window.__csp), []);
    }));

  test('the host answering a WYSIWYG request with the Preview (not available): back to Preview with the reason', async (h) => {
    const host = mdHost('fallback.md', DOC);
    const reason = 'This file is too large for WYSIWYG editing. Use Split to edit the Markdown source.';
    return h.use(host, {}, async (view) => {
      host.limits = { wysiwyg: false, reason }; // changed on the host side after the init: the next update says so
      await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
      await view.page.waitForFunction(() => document.querySelector('.md-mode-btn[data-mode="preview"]').getAttribute('aria-checked') === 'true', { timeout: 6000 });
      assert.equal(await status(view), reason);
      assert.equal(await view.eval(() => document.querySelector('.md-view').hidden), false);
    });
  });

  // ----- the REPO documents: 0 edits after load, protected share -----
  test('the REPO documents open in WYSIWYG with 0 edits (LF and CRLF); the protected share stays small', async (h) => {
    const files = [
      path.join(ROOT, 'test', 'samples', 'kitchen-sink.md'),
      path.join(ROOT, 'test', 'samples', 'other.md'),
      path.join(ROOT, 'README.md'),
      path.join(ROOT, 'CHANGELOG.md'),
      path.join(ROOT, 'CONTRIBUTING.md'),
    ].filter((f) => fs.existsSync(f));
    assert.ok(files.length >= 3, 'documents found');
    let blocks = 0;
    let raw = 0;
    for (const file of files) {
      const lf = fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
      for (const text of [lf, lf.replace(/\n/g, '\r\n')]) {
        const host = mdHost(path.basename(file), text);
        // (a raw <img> of kitchen-sink.md is a relative path the mock page cannot serve: the Preview of 0.1.2 shows it the same way)
        await h.use(host, { ignoreErrors: /HTTP 404: .*\.png/ }, async (view) => {
          await toWysiwyg(view);
          await sleep(300);
          assert.equal(await view.eval(() => document.querySelector('.md-mode-btn[data-mode="wysiwyg"]').getAttribute('aria-checked')), 'true', `${path.basename(file)} stays in WYSIWYG (self-check passed)`);
          assert.deepEqual(await view.posted('edit'), [], `${path.basename(file)}: 0 edits`);
          assert.equal(host.text, text);
          const logged = host.log.map((m) => /WYSIWYG: (\d+) editable and (\d+) protected blocks of (\d+)/.exec(m.message)).find(Boolean);
          assert.ok(logged, 'the load statistics are logged');
          if (text === lf) {
            blocks += Number(logged[3]);
            raw += Number(logged[2]);
          }
        });
      }
    }
    const share = raw / blocks;
    console.log(`protected share over ${files.length} documents: ${(share * 100).toFixed(1)} % (${raw} of ${blocks} blocks)`);
    assert.ok(share <= 0.3, `protected share ${(share * 100).toFixed(1)} %`);
  }, { timeout: 240000 });
});
