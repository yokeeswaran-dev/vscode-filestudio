// Host tests for the cost of the Markdown view on large documents (fix round 1, P6). Parsing a 3,000-block document takes
// 60-110 ms on the extension host thread, so the host must not parse it more often than it needs to:
//   - opening a panel in Preview parses it once (the rendering), not once more for a block count nobody asked for
//   - opening WYSIWYG parses it once (blocks + headings come from the same parse)
//   - one accepted WYSIWYG edit parses the new text once (verify + rendering share it), and the event loop gets a turn
//     between the answer to the edit and the rendering that follows
// test/build-tests.js counts the parses: `globalThis.__fsMdParses` holds the source length of every markdown-it parse.
// Run with `npm test`.
'use strict';
const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');

const { T, host, extensionUri, token, writeTmp, lastOf, waitFor } = require('./load-host');

beforeEach(() => T.reset());

/** A document of `sections` x (heading, paragraph, list); `tag` makes its text unique (no result is shared between tests). */
function doc(tag, sections = 40) {
  const parts = [`# ${tag}`];
  for (let i = 0; i < sections; i++) {
    parts.push(`## Section ${i}`, `Paragraph ${i} of ${tag} with **bold** text and a [link](https://example.com/${i}).`, `- item ${i}a\n- item ${i}b`);
  }
  return parts.join('\n\n') + '\n';
}

/** Parses of (about) the whole of `text` since the log was cleared. */
function wholeParses(text) {
  return globalThis.__fsMdParses.filter((length) => length >= text.length * 0.9).length;
}

async function openMd(name, text) {
  const uri = writeTmp(name, text);
  const document = new T.FakeTextDocument(uri, text);
  T.textDocuments.set(uri.toString(), document);
  const panel = T.createPanel();
  new host.TextViewerProvider(extensionUri, 'markdown').resolveCustomTextEditor(document, panel, token);
  panel.send({ type: 'ready' });
  const init = await waitFor(() => lastOf(panel, 'init'), 3000, 'init');
  // the save check ends with an editState that is not "Checking the file…"
  await waitFor(() => {
    const state = [...panel.posted].reverse().find((m) => m.editState)?.editState;
    return state && state.reason !== 'Checking the file…';
  }, 3000, 'the save check');
  return { uri, document, panel, init };
}

async function setMode(opened, mode) {
  opened.panel.posted.length = 0;
  opened.panel.send({ type: 'setMode', mode });
  return waitFor(() => lastOf(opened.panel, 'markdownUpdate'), 3000, 'markdownUpdate after setMode');
}

let nextEditId = 1;
async function sendEdit(opened, oldText, text) {
  const { document, panel } = opened;
  const start = document.getText().indexOf(oldText);
  assert.ok(start >= 0, `"${oldText}" is in the document`);
  const editId = nextEditId++;
  panel.send({
    type: 'edit',
    editId,
    baseVersion: document.version,
    intent: { kind: 'mdText', edits: [{ start, end: start + oldText.length, oldText, text }] },
  });
  return waitFor(() => panel.posted.find((m) => m.type === 'editResult' && m.editId === editId), 3000, 'editResult');
}

test('opening a panel in Preview parses the document once per rendering sent, not again for a block count', async () => {
  const text = doc('cost-preview');
  globalThis.__fsMdParses.length = 0;
  const { init, panel } = await openMd('cost-preview.md', text);
  assert.ok(init.html.includes('<h1'));
  assert.deepEqual(init.limits, { wysiwyg: true });
  // The init and the update that follows the save check each carry one rendering = one parse of the preview.
  const renderings = panel.posted.filter((m) => (m.type === 'init' || m.type === 'markdownUpdate') && m.html).length;
  assert.ok(renderings >= 1);
  assert.equal(wholeParses(text), renderings);
});

test('a document with more lines than the block limit still gets its blocks counted at the first rendering', async () => {
  const text = Array.from({ length: 3001 }, (_, i) => `p${i}`).join('\n\n') + '\n'; // 3,001 blocks, 6,002 lines
  const { init } = await openMd('cost-lines.md', text);
  assert.equal(init.limits.wysiwyg, false);
  assert.match(init.limits.reason, /too large for WYSIWYG/);
});

test('switching to WYSIWYG parses the document once: the blocks and the headings come from the same parse', async () => {
  const text = doc('cost-wysiwyg');
  const opened = await openMd('cost-wysiwyg.md', text);
  globalThis.__fsMdParses.length = 0;
  const update = await setMode(opened, 'wysiwyg');
  assert.ok(update.blocks.length > 100);
  assert.equal(update.toc.length, 41);
  assert.equal(wholeParses(text), 1);
});

test('one WYSIWYG edit parses the new text once (verify and rendering share it) and answers it first', async () => {
  const text = doc('cost-edit');
  const opened = await openMd('cost-edit.md', text);
  await setMode(opened, 'wysiwyg');
  globalThis.__fsMdParses.length = 0;
  const result = await sendEdit(opened, 'Paragraph 3 of', 'Paragraph three of');
  assert.equal(result.accepted, true);
  const update = await waitFor(() => lastOf(opened.panel, 'markdownUpdate'), 3000, 'markdownUpdate after the edit');
  assert.ok(update.source.includes('Paragraph three of'));
  assert.equal(opened.document.getText(), text.replace('Paragraph 3 of', 'Paragraph three of'));
  assert.equal(wholeParses(text), 1);
});

test('after the answer to a WYSIWYG edit the event loop gets a turn before the rendering is made', async () => {
  const text = doc('cost-yield');
  const opened = await openMd('cost-yield.md', text);
  await setMode(opened, 'wysiwyg');
  const order = [];
  const post = opened.panel.webview.postMessage;
  opened.panel.webview.postMessage = (msg) => {
    if (msg.type === 'editResult') {
      order.push('editResult');
      setImmediate(() => order.push('turn')); // something else waiting for its turn (a keystroke, another extension)
    } else if (msg.type === 'markdownUpdate') {
      order.push('markdownUpdate');
    }
    return post(msg);
  };
  const result = await sendEdit(opened, 'Paragraph 5 of', 'Paragraph five of');
  assert.equal(result.accepted, true);
  await waitFor(() => order.includes('markdownUpdate'), 3000, 'markdownUpdate after the edit');
  assert.deepEqual(order, ['editResult', 'turn', 'markdownUpdate']);
});
