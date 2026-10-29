// A mock Markdown host with the editing messages of v0.2.0 (P6) for the webview tests: it answers like MarkdownView in
// src/viewerProvider.ts (init / markdownUpdate with blocks, limits and edit state, setMode, edit with the real
// applyMdTextEdits + verifyMdEdit, command undo / redo with a small undo stack, previewScrolled, flushed, editAnyway).
'use strict';
const H = require('./harness');

/**
 * A Markdown host with the editing messages. `opts`: editState, limits (what init says), mode (the host's default mode),
 * eol. The host keeps `commands`, `previews` (previewScrolled lines), `edits` (accepted), `refused`, `flushed`, `anyway`.
 */
function mdHost(fileName, text, opts = {}) {
  const M = H.renderer('markdown');
  const host = H.markdownHost(fileName, text);
  Object.assign(host, {
    origin: '',
    mode: opts.mode || 'preview',
    editState: opts.editState || { editable: true, kinds: ['mdText'] },
    limits: opts.limits || { wysiwyg: true },
    eol: opts.eol || (text.includes('\r\n') ? '\r\n' : '\n'),
    commands: [],
    previews: [],
    edits: [],
    refused: [],
    flushed: [],
    anyway: 0,
    undoStack: [],
    redoStack: [],
    modes: [],
  });
  const wysiwygOn = () => host.mode === 'wysiwyg' && host.editState.editable && host.limits.wysiwyg && (host.editState.kinds?.includes('mdText') || host.editState.code === 'fidelity');
  host.payload = () => {
    if (host.mode === 'wysiwyg' && !wysiwygOn()) host.mode = 'preview';
    const res = host.mode === 'wysiwyg' ? { html: '', toc: M.markdownHeadings(host.text) } : host.render();
    return {
      html: host.mode === 'wysiwyg' ? '' : res.html,
      toc: res.toc,
      source: host.text,
      version: host.version,
      eol: host.eol,
      limits: host.limits,
      resourceBase: `${host.origin || ''}/res/`,
      readOnly: !host.editState.editable,
      editState: host.editState,
      ...(host.mode === 'wysiwyg' ? { blocks: M.markdownBlocks(host.text, { contextCheck: true }) } : {}),
    };
  };
  host.update = () => ({ type: 'markdownUpdate', ...host.payload() });
  const base = host.receive.bind(host);
  host.receive = async (msg) => {
    switch (msg.type) {
      case 'ready':
        return [{ type: 'init', kind: 'markdown', fileName, mode: host.mode, ...host.payload() }];
      case 'setMode':
        host.modes.push(msg.mode);
        host.mode = msg.mode;
        return [host.update()];
      case 'previewScrolled':
        host.previews.push(msg.line);
        return [];
      case 'editAnyway':
        host.anyway++;
        return [];
      case 'flushed':
        host.flushed.push(msg.id);
        return [];
      case 'command': {
        host.commands.push(msg.id);
        if (msg.id === 'undo' && host.undoStack.length) {
          host.redoStack.push(host.text);
          host.text = host.undoStack.pop();
          host.version++;
          return [host.update()];
        }
        if (msg.id === 'redo' && host.redoStack.length) {
          host.undoStack.push(host.text);
          host.text = host.redoStack.pop();
          host.version++;
          return [host.update()];
        }
        return [];
      }
      case 'edit': {
        const reject = (code, reason) => {
          host.refused.push({ code, reason, editId: msg.editId });
          return [{ type: 'editResult', editId: msg.editId, accepted: false, code, reason, version: host.version }, host.update()];
        };
        if (msg.baseVersion !== host.version) return reject('stale', 'The file changed while you were editing. Your change was not applied; the view was refreshed.');
        const edits = msg.intent.edits;
        for (const e of edits) if (host.text.slice(e.start, e.end) !== e.oldText) return reject('invalid', 'FileStudio could not read this change. Nothing was changed.');
        const fixed = edits.map((e) => ({ ...e, text: e.text.replace(/\r\n|\r|\n/g, host.eol) }));
        let next;
        try {
          next = M.applyMdTextEdits(host.text, fixed);
        } catch {
          return reject('invalid', 'FileStudio could not read this change. Nothing was changed.');
        }
        if (!M.verifyMdEdit(host.text, fixed, next).ok) return reject('blocked', 'This change would also change other parts of the file; it was not applied.');
        host.undoStack.push(host.text);
        host.redoStack.length = 0;
        host.edits.push(fixed);
        host.text = next;
        host.version++;
        return [{ type: 'editResult', editId: msg.editId, accepted: true, version: host.version }, host.update()];
      }
      default:
        return base(msg);
    }
  };
  return host;
}

module.exports = { mdHost };
