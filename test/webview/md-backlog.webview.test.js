// 0.2.1 backlog in WYSIWYG: B3 (typing right after an undo is kept when it can be merged safely) and B7 (the writer keeps
// the author's separators: no blank line added before/after an ATX heading, a setext underline keeps its length).
'use strict';
const assert = require('node:assert/strict');
const H = require('./harness');
const { mdHost } = require('./md-host');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const hooks = { before: async (pg) => pg.evaluateOnNewDocument(() => { window.__fsTestHooks = {}; }), ignoreErrors: /HTTP 404: .*\.png/ };
const PRELUDE = `const t = window.__fsTestHooks.mdw; const v = t.editor.action((ctx) => ctx.get(t.mw.core.editorViewCtx)); const sch = v.state.schema;
const find = (txt) => { let r = null; v.state.doc.forEach((n, off) => { if (!r && n.textContent.includes(txt)) r = { node: n, pos: off }; }); if (!r) throw new Error('not found ' + txt); return r; };`;

async function toWysiwyg(view) {
  await view.page.click('.md-mode-btn[data-mode="wysiwyg"]');
  await view.page.waitForFunction(() => window.__fsTestHooks.mdw && window.__fsTestHooks.mdw.state.ready, { timeout: 30000, polling: 50 });
  await view.settle();
}

const DOC = '# Title\n\nFirst paragraph.\n\nSecond paragraph.\n\nLast paragraph.\n';

H.suite('markdown WYSIWYG backlog 0.2.1 (B3, B7)', ({ test }) => {
  test('B3: typing in the window after an undo is merged onto the new source when it does not touch the undone text', async (h) => {
    const host = mdHost('undo-merge.md', DOC);
    await h.use(host, hooks, async (view) => {
      await toWysiwyg(view);
      await view.eval(`(() => { ${PRELUDE} const p = find('First paragraph'); v.dispatch(v.state.tr.insertText(' one', p.pos + 1 + 'First paragraph.'.length)); })()`);
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await view.settle();
      assert.equal(host.edits.length, 1);
      // Ctrl+Z happens in the host (the answer is on its way) ...
      host.text = DOC;
      host.version++;
      // ... and the user types in another paragraph before the answer arrives
      await view.eval(`(() => { ${PRELUDE} const p = find('Last paragraph'); v.dispatch(v.state.tr.insertText('X', p.pos + 1 + 'Last paragraph.'.length)); })()`);
      await view.push(host.update());
      await view.page.waitForFunction(() => document.querySelector('.md-wysiwyg .ProseMirror').textContent.includes('Last paragraph.X'), { timeout: 6000, polling: 50 });
      await sleep(900);
      await view.settle();
      assert.deepEqual(host.refused, [], 'nothing refused');
      assert.equal(host.text, DOC.replace('Last paragraph.', 'Last paragraph.X'), 'undone, and the typing kept');
      const shown = await view.eval(() => document.querySelector('.md-wysiwyg .ProseMirror').textContent);
      assert.ok(!shown.includes('First paragraph. one'), 'the undo shows');
    });
  }, { timeout: 90000 });

  test('B3: typing that touches the undone text is dropped with the message (never wrong text)', async (h) => {
    const host = mdHost('undo-overlap.md', DOC);
    await h.use(host, hooks, async (view) => {
      await toWysiwyg(view);
      await view.eval(`(() => { ${PRELUDE} const p = find('First paragraph'); v.dispatch(v.state.tr.insertText(' one', p.pos + 1 + 'First paragraph.'.length)); })()`);
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await view.settle();
      host.text = DOC;
      host.version++;
      await view.eval(`(() => { ${PRELUDE} const p = find('First paragraph'); v.dispatch(v.state.tr.insertText('Q', p.pos + 1 + 'First paragraph.'.length + 2)); })()`);
      await view.push(host.update());
      await sleep(1200);
      await view.settle();
      assert.equal(host.text, DOC, 'the file is the undone text');
      assert.deepEqual(host.refused, []);
      assert.match(await view.eval(() => document.querySelector('.md-status').textContent), /The file changed|not applied|^$/);
    });
  }, { timeout: 90000 });


  test('B3 race: typing after a moved edit, before the host update arrives, is written at the right place', async (h) => {
    const host = mdHost('race.md', DOC);
    await h.use(host, hooks, async (view) => {
      await toWysiwyg(view);
      // hold back the host's markdownUpdate that follows an accepted edit
      const held = [];
      const orig = host.receive;
      host.receive = async (msg) => {
        const out = await orig(msg);
        if (msg.type === 'edit' && host.hold) return out.filter((m) => (m.type === 'markdownUpdate' ? (held.push(m), false) : true));
        return out;
      };
      await view.eval(`(() => { ${PRELUDE} const p = find('First paragraph'); v.dispatch(v.state.tr.insertText('X', p.pos + 1 + 'First paragraph.'.length)); })()`);
      // the file changes elsewhere (the title) before the typing is sent: the typing is moved onto the new text
      host.text = DOC.replace('# Title', '# Title!');
      host.version++;
      host.hold = true;
      await view.push(host.update());
      await view.page.waitForFunction(() => window.__posted.some((m) => m.type === 'edit'), { timeout: 4000, polling: 30 });
      await sleep(300);
      assert.equal(host.edits.length, 1, 'the moved typing was accepted');
      // typing before the host update arrives
      await view.eval(`(() => { ${PRELUDE} const p = find('Second paragraph'); v.dispatch(v.state.tr.insertText(' Y', p.pos + 1 + 'Second paragraph.'.length)); })()`);
      await sleep(1200);
      host.hold = false;
      assert.equal(held.length, 1);
      await view.push(held[0]);
      await sleep(1500);
      await view.settle();
      assert.deepEqual(host.refused, []);
      assert.equal(host.text, '# Title!\n\nFirst paragraph.X\n\nSecond paragraph. Y\n\nLast paragraph.\n');
    });
  }, { timeout: 90000 });

  test('B3: a failed edit computation is not sent as an empty edit list', async (h) => {
    const host = mdHost('empty.md', DOC);
    await h.use(host, hooks, async (view) => {
      await toWysiwyg(view);
      const r = await view.eval(() => window.__fsTestHooks.mdw.rebase('a\n\nb', 'a\n\nc', []));
      assert.equal(r, null);
    });
  }, { timeout: 90000 });

  /** [name, source, page code (v, sch, find), expected text] */
  const SEP = [
    ['typing before an ATX heading without a blank line adds no blank line', 'Intro para\n# Heading\n\nNext\n', `const p = find('Intro para'); v.dispatch(v.state.tr.insertText(' more', p.pos + 1 + 10));`, 'Intro para more\n# Heading\n\nNext\n'],
    ['typing in an ATX heading right after a paragraph adds no blank line', 'Intro para\n# Heading\n\nNext\n', `const p = find('Heading'); v.dispatch(v.state.tr.insertText('X', p.pos + 1 + 7));`, 'Intro para\n# HeadingX\n\nNext\n'],
    ['typing before a fenced code block without a blank line adds no blank line', 'Intro para\n```js\nlet a;\n```\n', `const p = find('Intro para'); v.dispatch(v.state.tr.insertText('!', p.pos + 1 + 10));`, 'Intro para!\n```js\nlet a;\n```\n'],
    ['a heading turned into a paragraph keeps a blank line before the next paragraph (no merge)', '# Head\nnext para\n', `const p = find('Head'); v.dispatch(v.state.tr.setNodeMarkup(p.pos, sch.nodes.paragraph));`, 'Head\n\nnext para\n'],
    ['setext heading (=): typing keeps the underline', 'Title\n=====\n\nBody\n', `const p = find('Title'); v.dispatch(v.state.tr.insertText('X', p.pos + 1 + 5));`, 'TitleX\n=====\n\nBody\n'],
    ['setext heading (-): typing a longer text keeps the underline', 'Sub\n---\n\nBody\n', `const p = find('Sub'); v.dispatch(v.state.tr.insertText(' heading is long', p.pos + 1 + 3));`, 'Sub heading is long\n---\n\nBody\n'],
    ['setext heading (=) with a long underline: deleting text keeps the underline', 'Title here\n==========================\n\nBody\n', `const p = find('Title here'); v.dispatch(v.state.tr.delete(p.pos + 1 + 5, p.pos + 1 + 10));`, 'Title\n==========================\n\nBody\n'],
  ];
  for (const [name, source, body, expected] of SEP) {
    test(`B7: ${name}`, async (h) => {
      const host = mdHost('sep.md', source);
      await h.use(host, hooks, async (view) => {
        await toWysiwyg(view);
        const res = await view.eval(`(() => { ${PRELUDE} ${body}; const r = t.computeEdits(); return r.next.source; })()`);
        assert.equal(res, expected);
        await sleep(900);
        await view.settle();
        assert.deepEqual(host.refused, []);
        assert.equal(host.text, expected, 'the host has the same text');
      });
    }, { timeout: 90000 });
  }
});
