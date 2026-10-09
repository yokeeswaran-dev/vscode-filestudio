// PDF checks for the read-only PDF view. Rendering itself happens in the webview with pdf.js; the extension host only
// validates the file and turns failures into friendly messages (same style as docx.ts).
//
// Pure module (no `vscode`, no Node APIs): the host checks the file header with checkPdf, and the webview bundle may
// import describePdfFailure for the errors pdf.js raises while loading or rendering.

// ===== TYPES =====

export interface PdfCheck {
  /** undefined when the bytes look like a PDF; otherwise a friendly reason it cannot be shown. */
  error?: string;
  /** PDF version from the header, e.g. '1.7'. */
  version?: string;
}

// ===== IMPLEMENTATION =====

/** pdf.js (like Acrobat) looks for the `%PDF-` header within the first 1024 bytes; anything before it is ignored. */
export const PDF_HEADER_SEARCH_BYTES = 1024;

const PDF_HEADER = [0x25, 0x50, 0x44, 0x46, 0x2d]; // %PDF-
const ZIP_SIGNATURE = [0x50, 0x4b, 0x03, 0x04];
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const POSTSCRIPT_SIGNATURE = [0x25, 0x21, 0x50, 0x53]; // %!PS
/** Start of an HTML page after white space (a download that saved an error or login page). */
const HTML_START_RE = /^\s*(?:<!doctype\s+html|<html[\s>]|<head[\s>]|<body[\s>]|<\?xml[^>]*>\s*(?:<!doctype\s+html|<html[\s>]))/i;
const VERSION_RE = /^(\d{1,2}\.\d{1,2})/;

const EMPTY_FILE = 'The file is empty (0 bytes), so there is nothing to display.';

function startsWith(data: Uint8Array, sig: readonly number[], at = 0): boolean {
  if (data.length < at + sig.length) return false;
  return sig.every((b, i) => data[at + i] === b);
}

/** Offset of the `%PDF-` header in `head`, or -1. */
function findHeader(head: Uint8Array): number {
  const last = Math.min(head.length, PDF_HEADER_SEARCH_BYTES) - PDF_HEADER.length;
  for (let i = 0; i <= last; i++) {
    if (head[i] === PDF_HEADER[0] && startsWith(head, PDF_HEADER, i)) return i;
  }
  return -1;
}

/** Latin-1 text of a few bytes (enough to read a version number or recognise markup). */
function latin1(bytes: Uint8Array): string {
  let text = '';
  for (let i = 0; i < bytes.length; i++) text += String.fromCharCode(bytes[i]);
  return text;
}

/** What a file that is not a PDF most likely is, worded for users. */
function describeNonPdf(head: Uint8Array): string {
  if (startsWith(head, ZIP_SIGNATURE)) {
    return 'This file is not a PDF document: it is a ZIP package (for example a Word, Excel or PowerPoint file) renamed to .pdf.';
  }
  if (startsWith(head, OLE_SIGNATURE)) {
    return (
      'This file is not a PDF document: it is an Office compound file (a legacy .doc/.xls/.ppt file, or a ' +
      'password-protected Office document) renamed to .pdf.'
    );
  }
  if (startsWith(head, POSTSCRIPT_SIGNATURE)) {
    return 'This file is not a PDF document: it is a PostScript file. Convert it to PDF to view it here.';
  }
  // UTF-8 BOM bytes decode to three Latin-1 characters; drop them before looking for markup.
  const text = latin1(head.subarray(0, 512)).replace(/^\xEF\xBB\xBF/, '');
  if (HTML_START_RE.test(text)) {
    return (
      'This file is not a PDF document: it is a web page (HTML). A download may have saved an error or sign-in ' +
      'page instead of the PDF.'
    );
  }
  return 'This file is not a valid PDF document (it has no %PDF header). It may be corrupted or in a different format.';
}

/**
 * Checks the first bytes of a file (pass at least the first PDF_HEADER_SEARCH_BYTES bytes, or the whole file if it is
 * shorter; later bytes are ignored). A PDF has the `%PDF-x.y` header within the first 1024 bytes. Otherwise `error`
 * says why the file cannot be shown: empty, or another kind of file (ZIP/Office, OLE2, HTML, PostScript) renamed to
 * .pdf. Only the header is checked: a damaged body is reported by pdf.js in the webview (describePdfFailure).
 */
export function checkPdf(head: Uint8Array): PdfCheck {
  if (head.length === 0) return { error: EMPTY_FILE };
  const at = findHeader(head);
  if (at < 0) return { error: describeNonPdf(head) };
  const version = VERSION_RE.exec(latin1(head.subarray(at + PDF_HEADER.length, at + PDF_HEADER.length + 5)))?.[1];
  return version === undefined ? {} : { version };
}

// ===== ERRORS =====

/** pdf.js PasswordException codes (PasswordResponses). */
const NEED_PASSWORD = 1;
const INCORRECT_PASSWORD = 2;

/** The worker did not start: its script could not be fetched / evaluated, or pdf.js fell back to a fake worker. */
const WORKER_FAILED_RE = /fake worker|worker.*(?:failed|terminated|not (?:be )?(?:loaded|started))|workerSrc|importScripts|dynamically imported module/i;
/** The file could not be fetched from the webview resource URI. */
const FETCH_FAILED_RE = /failed to fetch|networkerror|network error|load failed|err_file_not_found|err_access_denied/i;
/** Damaged content found while parsing (pdf.js FormatError / XRefParseException / bad streams, wrapped by the worker). */
const DAMAGED_RE = /formaterror|xref|invalid (?:pdf|root|header|stream|object|dictionary)|bad (?:encoding|fcheck|xref)|unexpected (?:end|eof)|missing endstream|trailer|corrupt/i;

interface ErrorLike {
  name: string;
  message: string;
  code?: unknown;
  status?: unknown;
  missing?: unknown;
  details?: unknown;
}

/** Reads the fields pdf.js errors carry. They may arrive as Error subclasses or as plain (structured-cloned) objects. */
function toErrorLike(error: unknown): ErrorLike {
  if (typeof error === 'string') return { name: '', message: error };
  if (typeof error === 'object' && error !== null) {
    const e = error as Record<string, unknown>;
    return {
      name: typeof e.name === 'string' ? e.name : '',
      message: typeof e.message === 'string' ? e.message : '',
      code: e.code,
      status: e.status,
      missing: e.missing,
      details: e.details,
    };
  }
  return { name: '', message: error === undefined || error === null ? '' : String(error) };
}

/**
 * Readable reason for an error pdf.js reported in the webview (getDocument, page rendering, the worker), worded like
 * checkPdf. Library text ("Invalid PDF structure.", "Unexpected server response (404) ...") is not repeated; the
 * caller keeps the raw error for the details / log. Unknown failures keep their text.
 */
export function describePdfFailure(error: unknown): string {
  const e = toErrorLike(error);
  const raw = e.message.replace(/\s+/g, ' ').trim();
  switch (e.name) {
    case 'PasswordException':
      return e.code === INCORRECT_PASSWORD
        ? 'The password is incorrect. Enter the password of this PDF to open it.'
        : e.code === NEED_PASSWORD || e.code === undefined
          ? 'This PDF is password-protected. Enter its password to open it.'
          : 'This PDF is password-protected and could not be opened.';
    case 'InvalidPDFException':
      return /empty|zero bytes/i.test(raw)
        ? EMPTY_FILE
        : 'This file is not a valid PDF document: its structure is damaged (truncated or corrupted), so it cannot be read.';
    case 'ResponseException':
    case 'MissingPDFException':
    case 'UnexpectedResponseException':
      if (e.missing === true || e.status === 404 || e.name === 'MissingPDFException') {
        return 'The PDF file could not be found. It may have been moved, renamed or deleted.';
      }
      return typeof e.status === 'number' && e.status > 0
        ? `The PDF file could not be read (the viewer got status ${e.status} while loading it).`
        : 'The PDF file could not be read. It may have been moved, deleted or locked by another program.';
    case 'AbortException':
      return 'Loading the PDF was cancelled.';
    case 'RenderingCancelledException':
      return 'Rendering was cancelled.';
    default:
      break;
  }
  const details = typeof e.details === 'string' ? e.details : '';
  if (e.name === 'FormatError' || /^FormatError\b/.test(details) || DAMAGED_RE.test(`${raw} ${details}`)) {
    return 'This PDF could not be read: part of its content is damaged or uses a feature the viewer does not support.';
  }
  if (WORKER_FAILED_RE.test(raw)) {
    return 'The PDF engine could not start (its worker failed to load). Close and reopen the file; if it keeps failing, see the "FileStudio" output channel.';
  }
  if (FETCH_FAILED_RE.test(raw)) {
    return 'The PDF file could not be read. It may have been moved, deleted or locked by another program.';
  }
  return `Could not display the PDF: ${raw || 'Unknown error.'}`;
}
