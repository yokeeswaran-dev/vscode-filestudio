// FileStudio extension entry point: registers the six custom editors, the two commands and the fallback for
// text files VS Code cannot pass to a custom editor (too large, binary). All viewer logic lives in viewerProvider.ts.

import * as vscode from 'vscode';
import {
  DocxViewerProvider,
  PdfViewerProvider,
  PptxViewerProvider,
  SheetEditorProvider,
  TextViewerProvider,
  VIEW_TYPES,
  disposeLog,
  fileNameOf,
  getLog,
  reopenAsText,
  viewTypeForUri,
  watchUnreadableTextTabs,
} from './viewerProvider';

// ===== ACTIVATION =====

export function activate(context: vscode.ExtensionContext): void {
  const webviewOptions = { retainContextWhenHidden: true };
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPES.sheet, new SheetEditorProvider(context.extensionUri), {
      webviewOptions,
      supportsMultipleEditorsPerDocument: false,
    }),
    vscode.window.registerCustomEditorProvider(VIEW_TYPES.csv, new TextViewerProvider(context.extensionUri, 'csv'), {
      webviewOptions,
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.window.registerCustomEditorProvider(
      VIEW_TYPES.markdown,
      new TextViewerProvider(context.extensionUri, 'markdown'),
      { webviewOptions, supportsMultipleEditorsPerDocument: true },
    ),
    vscode.window.registerCustomEditorProvider(VIEW_TYPES.docx, new DocxViewerProvider(context.extensionUri), {
      webviewOptions,
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.window.registerCustomEditorProvider(VIEW_TYPES.pdf, new PdfViewerProvider(context.extensionUri), {
      webviewOptions,
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.window.registerCustomEditorProvider(VIEW_TYPES.pptx, new PptxViewerProvider(context.extensionUri), {
      webviewOptions,
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand('fileStudio.openWith', openWithFileStudio),
    vscode.commands.registerCommand('fileStudio.reopenAsText', reopenActiveAsText),
    // CSV/Markdown files VS Code does not pass to extensions as text (over its 50 MB limit, or binary) cannot open in a
    // custom text editor: use the text editor.
    watchUnreadableTextTabs(),
    // Last, so anything logged while the registrations above are disposed still has a channel.
    { dispose: disposeLog },
  );
}

export function deactivate(): void {
  // Everything is registered in context.subscriptions and disposed by VS Code.
}

// ===== COMMANDS =====

const SUPPORTED_FILES = '.xlsx, .csv, .tsv, .psv, .ssv, .md, .markdown, .docx, .pdf and .pptx';

/** URI of the resource in the active editor tab (text, custom, notebook or the modified side of a diff). */
function activeResourceUri(): vscode.Uri | undefined {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  if (
    input instanceof vscode.TabInputText ||
    input instanceof vscode.TabInputCustom ||
    input instanceof vscode.TabInputNotebook
  ) {
    return input.uri;
  }
  if (input instanceof vscode.TabInputTextDiff) return input.modified;
  return vscode.window.activeTextEditor?.document.uri;
}

/**
 * `fileStudio.openWith`. From the explorer the arguments are (clicked uri, all selected uris); from the editor title
 * the resource uri; from the command palette nothing (the active editor's resource is used).
 */
async function openWithFileStudio(arg?: unknown, selection?: unknown): Promise<void> {
  let uris: vscode.Uri[] = [];
  if (Array.isArray(selection) && selection.length > 0) {
    uris = selection.filter((item): item is vscode.Uri => item instanceof vscode.Uri);
  }
  if (uris.length === 0 && arg instanceof vscode.Uri) uris = [arg];
  if (uris.length === 0) {
    const active = activeResourceUri();
    if (active) uris = [active];
  }
  if (uris.length === 0) {
    void vscode.window.showInformationMessage(`Open or select a ${SUPPORTED_FILES} file first.`);
    return;
  }

  for (const uri of uris) {
    const viewType = viewTypeForUri(uri);
    if (!viewType) {
      void vscode.window.showWarningMessage(`FileStudio cannot open "${fileNameOf(uri)}" (supported: ${SUPPORTED_FILES}).`);
      continue;
    }
    try {
      // Several files: open them as regular (non-preview) tabs so they do not replace each other.
      await vscode.commands.executeCommand('vscode.openWith', uri, viewType, uris.length > 1 ? { preview: false } : undefined);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      getLog().error(`Could not open ${uri.toString(true)} with ${viewType}: ${message}`);
      void vscode.window.showErrorMessage(`FileStudio could not open "${fileNameOf(uri)}": ${message}`);
    }
  }
}

/** `fileStudio.reopenAsText`: the uri comes from the editor title button, else from the active custom editor tab. */
async function reopenActiveAsText(arg?: unknown): Promise<void> {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  const uri = arg instanceof vscode.Uri ? arg : input instanceof vscode.TabInputCustom ? input.uri : activeResourceUri();
  if (!uri) {
    void vscode.window.showInformationMessage('Open a CSV, TSV, PSV, SSV or Markdown file in FileStudio first.');
    return;
  }
  try {
    await reopenAsText(uri);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    getLog().error(`Could not reopen ${uri.toString(true)} as text: ${message}`);
    void vscode.window.showErrorMessage(`FileStudio could not reopen "${fileNameOf(uri)}" as text: ${message}`);
  }
}
