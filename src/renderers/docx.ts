// .docx -> HTML via mammoth (read-only view).
//
// Pure module (no `vscode` import). Images are embedded as data: URIs so the webview
// needs no extra resource roots; formats browsers cannot show (EMF/WMF/TIFF) and linked
// pictures (not stored in the package) become a labelled placeholder image plus a warning.
// Failures throw a descriptive Error that the provider shows in the webview's error view.

// ===== IMPORTS =====

import mammoth from 'mammoth';

// ===== TYPES =====

export interface DocxResult {
  html: string;
  /** Conversion warnings reported by mammoth. */
  warnings: string[];
}

// ===== IMPLEMENTATION =====

// ===== STYLE MAP =====

/**
 * Word styles -> HTML. Matching on style names is case-insensitive in mammoth; entries are
 * tried before mammoth's default map (which still handles lists, footnotes, etc.).
 */
const STYLE_MAP: string[] = [
  // Document title block
  "p[style-name='Title'] => h1.doc-title:fresh",
  "p[style-name='Subtitle'] => p.doc-subtitle:fresh",

  // Headings (style names and the built-in style ids)
  "p[style-name='Heading 1'] => h1:fresh",
  "p[style-name='Heading 2'] => h2:fresh",
  "p[style-name='Heading 3'] => h3:fresh",
  "p[style-name='Heading 4'] => h4:fresh",
  "p[style-name='Heading 5'] => h5:fresh",
  "p[style-name='Heading 6'] => h6:fresh",
  'p.Heading1 => h1:fresh',
  'p.Heading2 => h2:fresh',
  'p.Heading3 => h3:fresh',
  'p.Heading4 => h4:fresh',
  'p.Heading5 => h5:fresh',
  'p.Heading6 => h6:fresh',
  "p[style-name='TOC Heading'] => h2.doc-toc-heading:fresh",

  // Quotes
  "p[style-name='Quote'] => blockquote > p:fresh",
  "p[style-name='Intense Quote'] => blockquote.intense > p:fresh",
  "p[style-name='Block Text'] => blockquote > p:fresh",

  // Code-ish paragraph styles: consecutive paragraphs merge into one <pre>
  "p[style-name='Code'] => pre:separator('\\n')",
  "p[style-name='Code Block'] => pre:separator('\\n')",
  "p[style-name='Source Code'] => pre:separator('\\n')",
  "p[style-name='HTML Preformatted'] => pre:separator('\\n')",
  "p[style-name='Plain Text'] => pre:separator('\\n')",
  "p[style-name='Macro Text'] => pre:separator('\\n')",

  // Captions / misc paragraphs
  "p[style-name='Caption'] => p.doc-caption:fresh",

  // Character styles
  "r[style-name='Strong'] => strong",
  "r[style-name='Emphasis'] => em",
  "r[style-name='Intense Emphasis'] => strong > em",
  "r[style-name='Subtle Emphasis'] => em",
  "r[style-name='Book Title'] => cite",
  "r[style-name='Code Char'] => code",
  "r[style-name='HTML Code'] => code",
  "r[style-name='Source Code Char'] => code",
  "r[style-name='Verbatim Char'] => code",
  "r[style-name='HTML Keyboard'] => kbd",
  "r[style-name='HTML Variable'] => var",

  // Direct formatting mammoth ignores by default
  'u => u',
  'strike => s',
  'highlight => mark',

  // Page and column breaks: mammoth drops them, which glues the words on either side
  // together ("Before break|After" -> "Before breakAfter"); a line break keeps them apart
  "br[type='page'] => br",
  "br[type='column'] => br",

  // Show review comments (anchored as superscript references, listed at the end)
  'comment-reference => sup',
];

// ===== IMAGES =====

/** Image types browsers render natively; anything else becomes a placeholder. */
const WEB_IMAGE_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/pjpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
  'image/svg+xml',
  'image/avif',
  'image/x-icon',
  'image/vnd.microsoft.icon',
]);

function escapeXml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c] ?? c);
}

/** Small grey SVG box saying why the image is not shown (e.g. "Image not shown (image/x-emf)"). */
function placeholderImage(text: string): string {
  const label = escapeXml(text);
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="320" height="80" viewBox="0 0 320 80">` +
    `<rect x="0.5" y="0.5" width="319" height="79" fill="#8882" stroke="#888" stroke-dasharray="4 3"/>` +
    `<text x="160" y="45" font-family="sans-serif" font-size="13" fill="#888" text-anchor="middle">${label}</text>` +
    `</svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`;
}

/** mammoth's rejection for a linked (not embedded) picture: external file access is off. */
const EXTERNAL_IMAGE_RE = /external image '([^']*)'/;

/**
 * Placeholder + our own warning for an image whose bytes cannot be read: a picture linked
 * to a file or URL instead of stored in the package (never fetched — privacy and local file
 * access), or a broken relationship. Without this mammoth silently drops the image and
 * reports its internal error text.
 */
function unreadableImage(err: unknown, warnings: string[]): { src: string } {
  const message = (err instanceof Error ? err.message : String(err)).trim();
  const linked = EXTERNAL_IMAGE_RE.exec(message);
  if (linked) {
    warnings.push(
      `A linked image (${linked[1]}) is not stored in the document, so it is not shown; a placeholder marks its position.`,
    );
    return { src: placeholderImage('Linked image not shown') };
  }
  // A missing/damaged media part: mammoth's message is internal ("Cannot read properties of null").
  warnings.push('An image could not be read from the document (its data is missing or damaged) and was replaced by a placeholder.');
  return { src: placeholderImage('Image could not be read') };
}

// ===== WARNINGS =====

/**
 * mammoth's own note for image types outside its short list. Always dropped: every image
 * goes through convertImage, which either shows it (WEB_IMAGE_TYPES, e.g. BMP renders
 * fine) or replaces it with a placeholder and adds our own warning.
 */
const UNDISPLAYABLE_IMAGE_RE = /^Image of type .+ is unlikely to display in web browsers$/;

/**
 * Flattens mammoth messages to text, de-duplicated (mammoth repeats e.g. the same
 * unrecognised-style warning for every paragraph), followed by our own warnings.
 */
function collectWarnings(messages: ReadonlyArray<{ type: string; message: string }>, extra: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const msg of messages) {
    const text = (msg.type === 'error' ? `Error: ${msg.message}` : msg.message).trim();
    if (UNDISPLAYABLE_IMAGE_RE.test(text)) continue;
    if (text && !seen.has(text)) {
      seen.add(text);
      out.push(text);
    }
  }
  for (const text of extra) {
    if (!seen.has(text)) {
      seen.add(text);
      out.push(text);
    }
  }
  return out;
}

// ===== ERRORS =====

const ZIP_SIGNATURE = [0x50, 0x4b];
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function startsWith(data: Uint8Array, sig: number[]): boolean {
  if (data.length < sig.length) return false;
  return sig.every((b, i) => data[i] === b);
}

/** Up-front check that gives users a clear reason instead of a zip parser error. */
function assertLooksLikeDocx(data: Uint8Array): void {
  if (data.length === 0) throw new Error('The file is empty (0 bytes), so there is nothing to display.');
  if (startsWith(data, OLE_SIGNATURE)) {
    throw new Error(
      'This file is not an Office Open XML (.docx) package. It is either password-protected/encrypted ' +
        'or a legacy Word 97-2003 (.doc) file renamed to .docx. Open it in Word and save it as a .docx document.',
    );
  }
  if (!startsWith(data, ZIP_SIGNATURE)) {
    throw new Error('This file is not a valid .docx document (it is not a ZIP package). It may be corrupted or in a different format.');
  }
}

/** jszip (and the inflate code it uses) on a truncated or damaged package. */
const ZIP_DAMAGED_RE =
  /corrupted zip|central dir|end of data reached|invalid signature|size mismatch|crc32|missing \d+ bytes|is this a zip file|invalid (?:stored block|block type|distance|code|literal)|incorrect header check|unexpected end of (?:file|data)/i;

/**
 * Readable reason for a conversion failure, worded like the non-zip check (assertLooksLikeDocx). The library's own
 * text ("Corrupted zip: can't find end of central directory", "[xmldom warning] unclosed xml attribute") is not
 * repeated here: it stays available as the error's `cause` for the details. Unknown failures keep their text.
 */
function describeFailure(err: unknown): string {
  const raw = (err instanceof Error ? err.message : String(err))
    .replace(/@#\[line:[^\]]*\]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (/encrypted zip/i.test(raw)) {
    return 'This file is not a valid .docx document: its ZIP package is encrypted.';
  }
  if (ZIP_DAMAGED_RE.test(raw)) {
    return 'This file is not a valid .docx document: its ZIP package is damaged (truncated or corrupted), so it cannot be read.';
  }
  if (/word\/document\.xml|main document part|Could not find/i.test(raw)) {
    return 'This file is not a valid .docx document: its ZIP package has no Word document body (it may be another kind of file, such as an .xlsx, renamed to .docx).';
  }
  if (/xmldom|parse error|unclosed|malformed|xml/i.test(raw)) {
    return 'This document could not be read: part of its content is damaged or is not valid Word XML.';
  }
  return `Could not convert the Word document: ${raw || 'Unknown error.'}`;
}

// ===== PUBLIC API =====

/**
 * Converts a .docx file to HTML for the read-only page view. Throws an Error with a
 * user-facing message when the file cannot be read.
 */
export async function renderDocx(data: Uint8Array): Promise<DocxResult> {
  assertLooksLikeDocx(data);

  const extraWarnings: string[] = [];
  const convertImage = mammoth.images.imgElement(async (image) => {
    const contentType = (image.contentType || '').toLowerCase();
    // Read first: a linked picture fails here whatever its type, and gets the matching message.
    let bytes: Buffer;
    try {
      bytes = await image.readAsBuffer();
    } catch (err) {
      return unreadableImage(err, extraWarnings);
    }
    if (!WEB_IMAGE_TYPES.has(contentType)) {
      extraWarnings.push(`An image of type ${contentType || 'unknown'} cannot be displayed and was replaced by a placeholder.`);
      return { src: placeholderImage(`Image not shown (${contentType || 'unknown format'})`) };
    }
    return { src: `data:${contentType};base64,${bytes.toString('base64')}` };
  });

  let result: Awaited<ReturnType<typeof mammoth.convertToHtml>>;
  try {
    const buffer = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
    result = await mammoth.convertToHtml(
      { buffer },
      {
        styleMap: STYLE_MAP,
        includeDefaultStyleMap: true,
        // Never honour a style map shipped inside the document (zip part `mammoth/style-map`):
        // it can emit any element/attribute (<style>, remote <img>, spoofed data-href links).
        includeEmbeddedStyleMap: false,
        convertImage,
        ignoreEmptyParagraphs: true,
        externalFileAccess: false,
      },
    );
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }

  const html = result.value.trim() ? result.value : '<p class="doc-empty"><em>This document has no text content.</em></p>';
  return { html, warnings: collectWarnings(result.messages, extraWarnings) };
}
