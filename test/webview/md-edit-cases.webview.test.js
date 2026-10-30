// WYSIWYG edit cases (v0.2.0, P6; ported from the s2-markdown spike): ProseMirror transactions like a user makes, then
// the edits the editor computes, then the host's checks. For every case: (1) applying the edits to the old source gives the
// text the editor expects; (2) every block that was NOT touched keeps its exact text and kind (the host's verifyMdEdit, the
// real one); (3) the change is small; (4) sent to the mock host the edit is accepted and the host text is that text.
// Plus the click-through: the caret goes into every block of a document (and a sample of table cells and list items):
// 0 edits.
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const H = require('./harness');
const { mdHost } = require('./md-host');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const KS = fs.readFileSync(path.join(__dirname, '..', 'samples', 'kitchen-sink.md'), 'utf8').replace(/\r\n/g, '\n');
const M = H.renderer('markdown');

/** The test seam of the editor (media/viewer.js: globalThis.__fsTestHooks) must exist before the script runs. */
const hooks = { before: async (pg) => pg.evaluateOnNewDocument(() => { window.__fsTestHooks = {}; }), ignoreErrors: /HTTP 404: .*\.png/ };

const PRELUDE = `const t = window.__fsTestHooks.mdw; const v = t.editor.action((ctx) => ctx.get(t.mw.core.editorViewCtx)); const sch = v.state.schema;
const find = (txt) => { let r = null; v.state.doc.forEach((n, off) => { if (!r && n.textContent.includes(txt)) r = { node: n, pos: off }; }); if (!r) throw new Error('not found ' + txt); return r; };
const textAt = (p, needle, after) => { let at = -1; v.state.doc.nodesBetween(p.pos, p.pos + p.node.nodeSize, (n, pos) => { if (n.isText && n.text.includes(needle) && at < 0) at = pos + (after ? n.text.indexOf(needle) + needle.length : 0); }); return at; };`;

/** [name, source, body (page code using v, sch, find, textAt), max bytes changed] */
const CASES = [
  ['type in paragraph', KS, `const p = find('This document exercises'); v.dispatch(v.state.tr.insertText(' EDITED', p.pos + 1 + 5));`, 60],
  ['type in heading', KS, `const p = find('Heading level 3'); v.dispatch(v.state.tr.insertText('X', p.pos + p.node.nodeSize - 1));`, 60],
  ['type in a list item (only the item changes)', KS, `const p = find('Granny Smith'); v.dispatch(v.state.tr.insertText(' apple', textAt(p, 'Granny', true)));`, 80],
  ['type in a table cell', KS, `const p = find('Left aligned'); v.dispatch(v.state.tr.insertText('!', textAt(p, 'Left aligned', true)));`, 3000],
  ['edit a protected block (front matter) as text', KS, `const p = find('title: Kitchen Sink'); v.dispatch(v.state.tr.insertText('Z', p.pos + 1 + 3));`, 60],
  ['new paragraph after a heading (Enter at its end)', KS, `const p = find('Heading level 1'); v.dispatch(v.state.tr.insert(p.pos + p.node.nodeSize, sch.nodes.paragraph.create(null, sch.text('new para'))));`, 80],
  ['new paragraph between two blocks without a gap', KS, `const p = find('Heading level 5'); v.dispatch(v.state.tr.insert(p.pos + p.node.nodeSize, sch.nodes.paragraph.create(null, sch.text('inserted'))));`, 80],
  ['delete a block', KS, `const p = find('Ordered, nested, starting at 3'); v.dispatch(v.state.tr.delete(p.pos, p.pos + p.node.nodeSize));`, 2000],
  ['delete the first block', '# Title\n\nPara one.\n\nPara two.\n', `const p = find('Title'); v.dispatch(v.state.tr.delete(p.pos, p.pos + p.node.nodeSize));`, 60],
  ['delete the last block', '# Title\n\nPara one.\n\nPara two.\n', `const p = find('Para two'); v.dispatch(v.state.tr.delete(p.pos, p.pos + p.node.nodeSize));`, 60],
  ['move a block (drag = delete + insert)', KS, `const p = find('Plain text with'); const n = p.node; let tr = v.state.tr.delete(p.pos, p.pos + n.nodeSize); let target = -1; tr.doc.forEach((c, off) => { if (target < 0 && c.textContent.includes('Links and images')) target = off; }); tr = tr.insert(target, n); v.dispatch(tr);`, 3000],
  ['split a paragraph (Enter in the middle)', KS, `const p = find('A soft line break'); v.dispatch(v.state.tr.split(p.pos + 1 + 6));`, 200],
  ['CRLF: type in a paragraph', KS.replace(/\n/g, '\r\n'), `const p = find('This document exercises'); v.dispatch(v.state.tr.insertText(' EDITED', p.pos + 1 + 5));`, 60],
  ['CRLF: split a paragraph', KS.replace(/\n/g, '\r\n'), `const p = find('A soft line break'); v.dispatch(v.state.tr.split(p.pos + 1 + 6));`, 200],
  ['BOM + no final newline: type', '﻿# T\n\nabc', `const p = find('abc'); v.dispatch(v.state.tr.insertText('d', p.pos + 4));`, 40],
  ['emoji (surrogate pairs): type next to them', '# T\n\nab\u{1F600}\u{1F600}cd\n', `const p = find('ab'); v.dispatch(v.state.tr.insertText('\u{1F600}', p.pos + 1 + 4));`, 40],
  ['typing at the end of a paragraph that is followed by a list (no merge)', 'Para one.\n\n- item\n', `const p = find('Para one'); v.dispatch(v.state.tr.insertText(' end', p.pos + p.node.nodeSize - 1));`, 60],
  ['toggle a task item', KS, `let at = -1; let node = null; v.state.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'list_item' && n.attrs.checked === false) { at = pos; node = n; } }); if (at >= 0) v.dispatch(v.state.tr.setNodeMarkup(at, undefined, { ...node.attrs, checked: true }));`, 400],
];

H.suite('markdown WYSIWYG edit cases (P6)', ({ test }) => {
  for (const [name, source, body, maxBytes] of CASES) {
    test(`edit case: ${name}`, async (h) => {
      const eol = source.includes('\r\n') ? '\r\n' : '\n';
      const host = mdHost('case.md', source, { eol });
      await h.use(host, hooks, async (view) => {
        await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
        await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg:not([hidden]) .ProseMirror'), { timeout: 20000, polling: 50 });
        await view.settle();
        assert.deepEqual(await view.posted('edit'), [], 'no edit after the load');
        const result = await view.eval(`(() => { ${PRELUDE} ${body}; const res = t.computeEdits(); return { edits: res.edits, nextSource: res.next.source }; })()`);
        // (1) the edits give the text the editor expects
        let applied = source;
        for (const e of [...result.edits].sort((a, b) => b.start - a.start)) {
          assert.equal(applied.slice(e.start, e.end), e.oldText, 'oldText matches');
          applied = applied.slice(0, e.start) + e.text + applied.slice(e.end);
        }
        assert.equal(applied, result.nextSource, 'bookkeeping');
        // (2) every untouched block keeps its text and kind (the host's real check)
        assert.deepEqual(M.verifyMdEdit(source, result.edits, applied), { ok: true, bad: [] }, 'untouched blocks');
        // (3) small
        const changed = result.edits.reduce((n, e) => n + e.oldText.length + e.text.length, 0);
        assert.ok(changed <= maxBytes, `${changed} characters changed (max ${maxBytes})`);
        if (eol === '\r\n') assert.ok(!result.edits.some((e) => /(^|[^\r])\n/.test(e.text)), 'no bare LF in a CRLF document');
        if (source.charCodeAt(0) === 0xfeff) assert.equal(applied.charCodeAt(0), 0xfeff, 'BOM kept');
        // (4) sent: the mock host accepts it (it runs the same checks) and has exactly this text
        await sleep(900);
        await view.settle();
        assert.deepEqual(host.refused, [], 'not refused');
        assert.equal(host.text, result.edits.length ? result.nextSource : source);
        assert.equal(host.edits.length, result.edits.length ? 1 : 0);
      });
    });
  }

  test('click-through: the caret in the blocks of five documents (LF and CRLF, 240+ clicks): 0 edits', async (h) => {
    let clicks = 0;
    // the kitchen sink (LF and CRLF), then the first blocks of three real documents (their line breaks as they are)
    const repoDoc = (name) => fs.readFileSync(path.join(__dirname, '..', '..', name), 'utf8');
    const documents = [[KS, Infinity], [KS.replace(/\n/g, '\r\n'), Infinity], [repoDoc('README.md'), 40], [repoDoc('CHANGELOG.md'), 40], [repoDoc('CONTRIBUTING.md'), 40]];
    for (const [text, cap] of documents) {
      const host = mdHost('click.md', text);
      await h.use(host, hooks, async (view) => {
        await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
        await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg:not([hidden]) .ProseMirror'), { timeout: 20000, polling: 50 });
        await view.settle();
        const count = await view.eval(() => document.querySelector('.md-wysiwyg .ProseMirror').children.length);
        for (let i = 0; i < Math.min(count, cap); i++) {
          const pt = await view.eval((i) => {
            const el = document.querySelector('.md-wysiwyg .ProseMirror').children[i];
            if (!el || el.classList.contains('ProseMirror-widget')) return null;
            el.scrollIntoView({ block: 'center' });
            const r = el.getBoundingClientRect();
            return { x: r.left + Math.min(r.width / 2, 200), y: r.top + Math.min(8, r.height / 2) };
          }, i);
          if (!pt) continue;
          await view.page.mouse.click(pt.x, pt.y);
          clicks++;
        }
        await sleep(700);
        await view.settle();
        const edits = await view.eval(() => window.__fsTestHooks.mdw.computeEdits().edits);
        assert.deepEqual(edits, [], 'the editor has nothing to send');
        assert.deepEqual(await view.posted('edit'), [], 'nothing was sent');
        assert.equal(host.text, text);
      });
    }
    console.log(`click-through: ${clicks} clicks, 0 edits`);
    assert.ok(clicks >= 240, `${clicks} clicks`);
  }, { timeout: 600000 });
});
