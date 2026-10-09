// .pptx -> slide model for the read-only presentation view (pptxtojson, normalised).
//
// The extension host parses the deck once; the webview gets the outline (PptxDeckMeta) at init and asks for slides
// one at a time (getSlide), so big decks with many images stay fast. Units are POINTS (1/72 inch) as pptxtojson
// reports them; the webview scales the slide to the available space.
//
// Pure module (no `vscode` import). pptxtojson draws the slides; what it does not report (deck order, hidden slides,
// titles, plain-text notes, fonts, content it drops) is read straight from the package with jszip. Failures throw an
// Error with a user-facing message (the library's error is kept as `cause`), like docx.ts.

// ===== IMPORTS =====

import JSZip from 'jszip';
import type {
  BaseElement as PjBaseElement,
  Element as PjElement,
  Fill as PjFill,
  Image as PjImage,
  Math as PjMath,
  Shape as PjShape,
  Text as PjText,
} from 'pptxtojson';
// Not the bare 'pptxtojson': its `main` is a UMD build, and as the package is "type": "module" esbuild bundles that
// file as ESM, where the UMD wrapper finds no `module`/`exports` and puts the API on globalThis — `parse` is then
// undefined in dist/extension.js. The ESM build (jszip and txml bundled inside) works; its types are dist/index.d.ts.
import { parse as parseDeck } from 'pptxtojson/dist/index.js';

// ===== TYPES (shared with media/viewer.js — keep in sync) =====

/**
 * One drawable element of a slide. These are pptxtojson Element objects (see node_modules/pptxtojson/dist/index.d.ts)
 * after normalisation:
 *  - `layout: true` marks elements that come from the slide layout / master (drawn first, behind the slide's own);
 *  - video / audio elements are replaced by `PptxMediaPlaceholder` (media is never played or embedded);
 *  - images keep `base64` as a complete data: URI; `blob` fields are removed;
 *  - text / shape `content` is pptxtojson's HTML, NOT sanitised here — the webview sanitises it with DOMPurify;
 *  - groups keep their child elements (recursively normalised the same way).
 */
export type PptxElement = (PjElement | PptxMediaPlaceholder) & { layout?: boolean };

export interface PptxMediaPlaceholder {
  type: 'media';
  mediaType: 'video' | 'audio';
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
  order: number;
}

export type PptxFill = PjFill;

export interface PptxSlideInfo {
  /** 0-based position in the deck. */
  index: number;
  /** Title placeholder text ('' when the slide has no title). */
  title: string;
  /** Slide hidden in the presentation (<p:sld show="0">). Still shown in the viewer, marked as hidden. */
  hidden: boolean;
  hasNotes: boolean;
  /**
   * (v0.2.0 integration, additive) Package part of the slide ('ppt/slides/slide3.xml'). A link to another slide names
   * this part ('slide3.xml'); the part number is not the deck position when slides were reordered. Missing when the
   * deck order could not be read (then the slides are in part-number order).
   */
  part?: string;
}

export interface PptxDeckMeta {
  /** Slide size in points. */
  width: number;
  height: number;
  slideCount: number;
  slides: PptxSlideInfo[];
  themeColors: string[];
  /** Fonts the deck uses (for a "font not installed" hint; never downloaded). */
  fonts: string[];
  /** Content the viewer cannot show (video, audio, OLE objects, SmartArt without fallback, …), for the banner. */
  lossy: string[];
}

export interface PptxSlide {
  index: number;
  hidden: boolean;
  /** Slide background (resolved through layout / master), or null for white. */
  background: PptxFill | null;
  /** Layout/master elements first, then the slide's own, each group in pptxtojson `order`. */
  elements: PptxElement[];
  /** Speaker notes as plain text ('' when none). */
  notes: string;
}

/** A parsed deck (renderPptx). A valid deck may have no slides (slideCount 0). */
export interface PptxDocument {
  meta: PptxDeckMeta;
  /** Slide `index` (0-based, deck order). Throws a RangeError for an index outside the deck. */
  getSlide(index: number): PptxSlide;
}

// ===== IMPLEMENTATION =====

/** Pictures as data: URIs (the webview CSP allows data: images); video and audio are never loaded. */
const PARSE_OPTIONS = { imageMode: 'base64', videoMode: 'none', audioMode: 'none' } as const;

type ParsedDeck = Awaited<ReturnType<typeof parseDeck>>;
type PjSlide = ParsedDeck['slides'][number];

// ===== LOSSY FEATURES =====

/** Banner names, in banner order. */
const LOSSY = {
  video: 'Video',
  audio: 'Audio',
  ole: 'Embedded objects (OLE)',
  controls: 'ActiveX controls',
  smartArt: 'SmartArt without a saved drawing',
  chartEx: 'Newer chart types (waterfall, treemap, …)',
  otherCharts: 'Radar / stock / surface / bubble charts (shown as data)',
  models3d: '3D models',
  ink: 'Ink drawings',
  linkedPictures: 'Linked pictures',
  pictures: 'Pictures a browser cannot show (EMF / WMF / TIFF)',
  macros: 'Macros (VBA project)',
  comments: 'Comments',
} as const;

const LOSSY_ORDER: string[] = Object.values(LOSSY);

/**
 * Markup (in a slide, layout or master) of content pptxtojson drops or the viewer does not play. pptxtojson skips OLE
 * objects, ActiveX controls, ink and every mc:AlternateContent other than equations and groups (3D models, chartex).
 */
const LOSSY_MARKUP: ReadonlyArray<readonly [RegExp, string]> = [
  [/<a:(?:videoFile|quickTimeFile)\b/, LOSSY.video],
  [/<a:(?:audioFile|wavAudioFile|audioCd)\b/, LOSSY.audio],
  [/<p:oleObj\b/, LOSSY.ole],
  [/<p:control\b/, LOSSY.controls],
  [/<p(?:14)?:contentPart\b/, LOSSY.ink],
  [/<\w+:model3d\b/i, LOSSY.models3d],
  [/\/drawing\/2014\/chartex"/, LOSSY.chartEx],
  // A picture linked to a file or URL and not stored in the package (never fetched).
  [/<a:blip\b(?=[^>]*\br:link=)(?![^>]*\br:embed=)/, LOSSY.linkedPictures],
];

/** A SmartArt frame; drawn only from the drawing PowerPoint saves with it (relationship type diagramDrawing). */
const DIAGRAM_FRAME_RE = /<a:graphicData\b[^>]*\buri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/diagram"/;

/** Chart types the webview draws (3D variants flat); the others are shown as a labelled placeholder with the data. */
const DRAWN_CHART_TYPES = new Set([
  'barChart',
  'bar3DChart',
  'lineChart',
  'line3DChart',
  'areaChart',
  'area3DChart',
  'pieChart',
  'pie3DChart',
  'doughnutChart',
  'scatterChart',
]);

// ===== XML HELPERS =====

const XML_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/**
 * XML character data -> text: entities and character references, then the OOXML `_xHHHH_` escapes Office writes for
 * control characters XML cannot hold (a vertical tab, i.e. a soft line break, is `_x000B_`).
 */
function decodeXml(text: string): string {
  if (!text.includes('&') && !text.includes('_x')) return text;
  return text
    .replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|([A-Za-z]+));/g, (all: string, dec?: string, hex?: string, name?: string) => {
      if (name) return XML_ENTITIES[name] ?? all;
      const code = dec !== undefined ? parseInt(dec, 10) : parseInt(hex ?? '', 16);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : all;
    })
    .replace(/_x(00[01][0-9A-Fa-f])_/g, (_all: string, hex: string) => {
      const code = parseInt(hex, 16);
      return code === 0x09 ? '\t' : code === 0x0a || code === 0x0b || code === 0x0d ? '\n' : '';
    });
}

/** Attributes of one start tag (`<a:off x="1" y="2"/>` -> { x: '1', y: '2' }), values decoded. */
function attrsOf(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[m[1]] = decodeXml(m[2] ?? m[3] ?? '');
  return attrs;
}

/** Attributes of the first `<name …>` start tag in `xml` ({} when there is none). */
function firstTagAttrs(xml: string, name: string): Record<string, string> {
  const tag = new RegExp(`<${name}(?=[\\s/>])[^>]*>`).exec(xml);
  return tag ? attrsOf(tag[0]) : {};
}

/** Start and end tags (comments, processing instructions and CDATA are not matched). */
const TAG_RE = /<(\/?)([A-Za-z_][\w.:-]*)(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/g;

/** The direct children of the first `<p:spTree>` (a slide, layout or master's shapes, in document order). */
function spTreeChildren(xml: string): Array<{ name: string; xml: string }> {
  const start = xml.search(/<p:spTree[\s>]/);
  if (start < 0) return [];
  const children: Array<{ name: string; xml: string }> = [];
  const re = new RegExp(TAG_RE.source, 'g');
  re.lastIndex = start;
  let depth = 0;
  let childStart = -1;
  let childName = '';
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    const [tag, closing, name, selfClosing] = m;
    if (closing) {
      depth--;
      if (depth === 1 && childStart >= 0) {
        children.push({ name: childName, xml: xml.slice(childStart, m.index + tag.length) });
        childStart = -1;
      }
      if (depth <= 0) break;
    } else if (selfClosing) {
      if (depth === 1) children.push({ name, xml: tag });
    } else {
      if (depth === 1) {
        childStart = m.index;
        childName = name;
      }
      depth++;
    }
  }
  return children;
}

/** Text of one `<a:p>` body: runs and fields, `<a:br/>` as a line break (paragraph properties skipped). */
function paragraphText(inner: string): string {
  const body = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, '');
  let text = '';
  for (const m of body.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<a:br\b/g)) {
    text += m[1] !== undefined ? decodeXml(m[1]) : '\n';
  }
  return text;
}

/** Paragraph texts of a text body (or any markup holding `<a:p>` elements). */
function paragraphs(xml: string): string[] {
  return [...xml.matchAll(/<a:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:p>)/g)].map((m) => paragraphText(m[1] ?? ''));
}

/** `<p:sp>` shapes whose placeholder type is one of `types` (`<p:sp>` never nests, so a lazy match is exact). */
function placeholderShapes(xml: string, types: ReadonlySet<string>): string[] {
  const shapes: string[] = [];
  for (const m of xml.matchAll(/<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g)) {
    const type = firstTagAttrs(m[1], 'p:ph').type;
    if (type !== undefined && types.has(type)) shapes.push(m[1]);
  }
  return shapes;
}

const TITLE_TYPES: ReadonlySet<string> = new Set(['title', 'ctrTitle']);
const NOTES_BODY_TYPES: ReadonlySet<string> = new Set(['body']);

/** The slide title: text of the first title / centred-title placeholder, on one line. */
function slideTitle(xml: string): string {
  const shape = placeholderShapes(xml, TITLE_TYPES)[0];
  return shape ? paragraphs(shape).join(' ').replace(/\s+/g, ' ').trim() : '';
}

/** Speaker notes as plain text: the notes page's body placeholder, one line per paragraph. */
function notesText(xml: string): string {
  const lines = placeholderShapes(xml, NOTES_BODY_TYPES).flatMap((shape) => paragraphs(shape));
  return lines
    .join('\n')
    .replace(/[ \t]+$/gm, '')
    .replace(/^\n+|\n+$/g, '');
}

/** `<p:sld show="0">`: the slide is skipped in a slide show. */
function isHiddenSlide(xml: string): boolean {
  const show = firstTagAttrs(xml, 'p:sld').show;
  return show === '0' || show === 'false';
}

// ===== PACKAGE =====

interface Relationship {
  id: string;
  /** Last segment of the relationship type URI ('slide', 'slideLayout', 'notesSlide', 'image', …). */
  type: string;
  /** Package path of the target part (or the raw target when external). */
  target: string;
  external: boolean;
}

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

/** Package path of `target` relative to the part `from` ('ppt/slides/slide1.xml' + '../media/a.png'). */
function resolvePart(from: string, target: string): string {
  let path = target.replace(/\\/g, '/');
  try {
    path = decodeURIComponent(path);
  } catch {
    // keep the raw target
  }
  const base = path.startsWith('/') ? '' : from.slice(0, from.lastIndexOf('/') + 1);
  const out: string[] = [];
  for (const segment of (base + path).split('/')) {
    if (segment === '..') out.pop();
    else if (segment !== '.' && segment !== '') out.push(segment);
  }
  return out.join('/');
}

/** Read access to the zip, tolerant of missing or damaged parts (pptxtojson reports those). */
class PptxPackage {
  readonly names: string[];
  private readonly texts = new Map<string, Promise<string | null>>();

  constructor(private readonly zip: JSZip) {
    this.names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  }

  has(path: string): boolean {
    return this.zip.file(path) !== null;
  }

  /** A part as text (cached), or null when it is missing or cannot be inflated. */
  text(path: string): Promise<string | null> {
    let text = this.texts.get(path);
    if (!text) {
      const file = this.zip.file(path);
      text = file ? file.async('string').catch(() => null) : Promise.resolve(null);
      this.texts.set(path, text);
    }
    return text;
  }

  /** Relationships of a part (from `<dir>/_rels/<name>.rels`). */
  async rels(part: string): Promise<Relationship[]> {
    const slash = part.lastIndexOf('/');
    const xml = await this.text(`${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`);
    if (!xml) return [];
    return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => {
      const a = attrsOf(m[0]);
      const external = a.TargetMode === 'External';
      const target = a.Target ?? '';
      return {
        id: a.Id ?? '',
        type: (a.Type ?? '').replace(/^.*\//, ''),
        target: external ? target : resolvePart(part, target),
        external,
      };
    });
  }

  /** Target of the first internal relationship of `type`. */
  async relTarget(part: string, type: string): Promise<string | undefined> {
    return (await this.rels(part)).find((rel) => rel.type === type && !rel.external)?.target;
  }
}

// ===== OUTLINE (read from the package) =====

interface SlidePart {
  hidden: boolean;
  title: string;
  notes: string;
  /** The layout's own (non-placeholder) shapes, each as the ids pptxtojson may give its element. */
  layoutShapes: string[][];
  /** The layout sets showMasterSp="0": pptxtojson's layoutElements then hold only the layout's shapes. */
  masterShapesHidden: boolean;
}

interface LayoutInfo {
  shapes: string[][];
  masterShapesHidden: boolean;
}

interface DeckOutline {
  /** Per pptxtojson slide (its `slides` order); undefined for parts not in the deck. */
  parts: Array<SlidePart | undefined>;
  /** Deck order (presentation.xml sldIdLst) as indexes into pptxtojson's `slides`. */
  order: number[];
  /** Package part of each pptxtojson slide (its `slides` order). */
  names: string[];
  fonts: string[];
  lossy: Set<string>;
}

const NO_PART: SlidePart = { hidden: false, title: '', notes: '', layoutShapes: [], masterShapesHidden: false };

/** Top-level nodes pptxtojson turns into elements. */
const DRAWABLE_NODES = new Set(['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'mc:AlternateContent']);

/**
 * The shapes pptxtojson draws from a layout (its getLayoutElements: every top-level node but `<p:sp>` placeholders),
 * each as the cNvPr ids its element may get (an mc:AlternateContent takes its id from the Choice or the Fallback).
 */
function layoutShapeIds(xml: string): string[][] {
  const firstId = (markup: string): string | undefined => firstTagAttrs(markup, 'p:cNvPr').id;
  const shapes: string[][] = [];
  for (const child of spTreeChildren(xml)) {
    if (!DRAWABLE_NODES.has(child.name)) continue;
    if (child.name === 'p:sp' && /<p:ph[\s/>]/.test(/<p:nvSpPr[\s>][\s\S]*?<\/p:nvSpPr>/.exec(child.xml)?.[0] ?? '')) {
      continue;
    }
    const ids =
      child.name === 'mc:AlternateContent'
        ? [...child.xml.matchAll(/<mc:(?:Choice|Fallback)\b[^>]*>([\s\S]*?)<\/mc:(?:Choice|Fallback)>/g)].map((m) => firstId(m[1]))
        : [firstId(child.xml)];
    shapes.push(ids.filter((id): id is string => id !== undefined));
  }
  return shapes;
}

/** Slide parts in pptxtojson's order: [Content_Types] slide overrides sorted by the number in the part name. */
async function parsedSlideOrder(pkg: PptxPackage): Promise<string[]> {
  const xml = (await pkg.text('[Content_Types].xml')) ?? '';
  const parts: string[] = [];
  for (const m of xml.matchAll(/<Override\b[^>]*>/g)) {
    const a = attrsOf(m[0]);
    if (a.ContentType === SLIDE_CONTENT_TYPE && a.PartName) parts.push(a.PartName.slice(1));
  }
  const num = (part: string): number => Number(/(\d+)\.xml/.exec(part)?.[1] ?? 0);
  return parts.sort((a, b) => num(a) - num(b));
}

/** Slide parts in deck order (presentation.xml sldIdLst). */
async function deckSlideOrder(pkg: PptxPackage): Promise<string[]> {
  const xml = (await pkg.text('ppt/presentation.xml')) ?? '';
  const targets = new Map(
    (await pkg.rels('ppt/presentation.xml')).filter((rel) => rel.type === 'slide' && !rel.external).map((rel) => [rel.id, rel.target]),
  );
  const list = /<p:sldIdLst\b[\s\S]*?<\/p:sldIdLst>/.exec(xml)?.[0] ?? '';
  const order: string[] = [];
  for (const m of list.matchAll(/<p:sldId\b[^>]*>/g)) {
    const target = targets.get(attrsOf(m[0])['r:id'] ?? '');
    if (target) order.push(target);
  }
  return order;
}

class OutlineReader {
  readonly lossy = new Set<string>();
  /** Theme fonts (headings / body) first, then fonts set directly on text. */
  readonly themeFonts = new Set<string>();
  readonly fonts = new Set<string>();
  private readonly layouts = new Map<string, LayoutInfo>();
  private readonly masters = new Set<string>();

  constructor(private readonly pkg: PptxPackage) {}

  async read(): Promise<DeckOutline> {
    const parsedOrder = await parsedSlideOrder(this.pkg);
    const indexOf = new Map(parsedOrder.map((part, i) => [part.toLowerCase(), i]));
    const order: number[] = [];
    const seen = new Set<number>();
    for (const part of await deckSlideOrder(this.pkg)) {
      const i = indexOf.get(part.toLowerCase());
      if (i !== undefined && !seen.has(i)) {
        seen.add(i);
        order.push(i);
      }
    }
    // Without a usable sldIdLst, keep pptxtojson's order.
    if (order.length === 0) parsedOrder.forEach((_part, i) => order.push(i));

    const parts: Array<SlidePart | undefined> = parsedOrder.map(() => undefined);
    for (const i of order) parts[i] = await this.readSlide(parsedOrder[i]);

    if (this.pkg.names.some((name) => /^ppt\/vbaProject\.bin$/i.test(name))) this.lossy.add(LOSSY.macros);
    if (this.pkg.names.some((name) => /^ppt\/comments\/[^/]+\.xml$/i.test(name))) this.lossy.add(LOSSY.comments);
    return { parts, order, names: parsedOrder, fonts: [...new Set([...this.themeFonts, ...this.fonts])], lossy: this.lossy };
  }

  private async readSlide(path: string): Promise<SlidePart> {
    const xml = (await this.pkg.text(path)) ?? '';
    const rels = await this.pkg.rels(path);
    const target = (type: string): string | undefined => rels.find((rel) => rel.type === type && !rel.external)?.target;
    this.scan(xml);
    if (DIAGRAM_FRAME_RE.test(xml) && !rels.some((rel) => rel.type === 'diagramDrawing')) this.lossy.add(LOSSY.smartArt);

    const notesPath = target('notesSlide');
    const notesXml = notesPath ? await this.pkg.text(notesPath) : null;
    const layoutPath = target('slideLayout');
    const layout = layoutPath ? await this.readLayout(layoutPath) : { shapes: [], masterShapesHidden: false };
    return {
      hidden: isHiddenSlide(xml),
      title: slideTitle(xml),
      notes: notesXml ? notesText(notesXml) : '',
      layoutShapes: layout.shapes,
      masterShapesHidden: layout.masterShapesHidden,
    };
  }

  private async readLayout(path: string): Promise<LayoutInfo> {
    let info = this.layouts.get(path);
    if (!info) {
      const xml = (await this.pkg.text(path)) ?? '';
      this.scan(xml);
      const masterPath = await this.pkg.relTarget(path, 'slideMaster');
      if (masterPath) await this.readMaster(masterPath);
      // pptxtojson draws the master's shapes unless showMasterSp is exactly '0'.
      info = { shapes: layoutShapeIds(xml), masterShapesHidden: firstTagAttrs(xml, 'p:sldLayout').showMasterSp === '0' };
      this.layouts.set(path, info);
    }
    return info;
  }

  private async readMaster(path: string): Promise<void> {
    if (this.masters.has(path)) return;
    this.masters.add(path);
    this.scan((await this.pkg.text(path)) ?? '');
    const themePath = await this.pkg.relTarget(path, 'theme');
    const theme = themePath ? await this.pkg.text(themePath) : null;
    const fontScheme = theme ? /<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/.exec(theme)?.[0] : undefined;
    if (fontScheme) collectFonts(fontScheme, this.themeFonts);
  }

  /** Lossy markup and fonts of a slide, layout or master. */
  private scan(xml: string): void {
    for (const [re, feature] of LOSSY_MARKUP) if (!this.lossy.has(feature) && re.test(xml)) this.lossy.add(feature);
    collectFonts(xml, this.fonts);
  }
}

/** Typefaces of `<a:latin>` / `<a:ea>` / `<a:cs>` (theme references such as `+mn-lt` and empty names skipped). */
function collectFonts(xml: string, into: Set<string>): void {
  for (const m of xml.matchAll(/<a:(?:latin|ea|cs)\b[^>]*>/g)) {
    const face = (attrsOf(m[0]).typeface ?? '').trim();
    if (face && !face.startsWith('+')) into.add(face);
  }
}

// ===== PICTURES =====

/** Declared types a browser shows when the bytes match no known signature (pptxtojson names them by extension). */
const BROWSER_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/avif']);

/** Picture format by signature: a browser MIME type, or the name of a format a browser cannot show. */
function sniffPicture(head: Uint8Array): { mime: string } | { format: string } | undefined {
  const bytesAt = (offset: number, sig: number[]): boolean => sig.every((b, i) => head[offset + i] === b);
  const textAt = (offset: number, text: string): boolean => bytesAt(offset, [...text].map((c) => c.charCodeAt(0)));
  if (bytesAt(0, [0x89, 0x50, 0x4e, 0x47])) return { mime: 'image/png' };
  if (bytesAt(0, [0xff, 0xd8, 0xff])) return { mime: 'image/jpeg' };
  if (textAt(0, 'GIF8')) return { mime: 'image/gif' };
  if (textAt(0, 'RIFF') && textAt(8, 'WEBP')) return { mime: 'image/webp' };
  if (textAt(4, 'ftypavif') || textAt(4, 'ftypavis')) return { mime: 'image/avif' };
  if (bytesAt(0, [0x01, 0, 0, 0]) && textAt(40, ' EMF')) return { format: 'EMF' };
  if (bytesAt(0, [0xd7, 0xcd, 0xc6, 0x9a]) || bytesAt(0, [0x01, 0, 0x09, 0]) || bytesAt(0, [0x02, 0, 0x09, 0])) {
    return { format: 'WMF' };
  }
  if (bytesAt(0, [0x49, 0x49, 0x2a, 0]) || bytesAt(0, [0x4d, 0x4d, 0, 0x2a])) return { format: 'TIFF' };
  if (bytesAt(0, [0x1f, 0x8b])) return { format: 'compressed metafile' };
  if (textAt(0, 'BM')) return { mime: 'image/bmp' };
  if (bytesAt(0, [0, 0, 0x01, 0])) return { mime: 'image/x-icon' };
  return undefined;
}

type WebPicture = { uri: string } | { reason: string };

interface NormaliseContext {
  lossy: Set<string>;
  /** Checked pictures by package path: a picture on a master is checked (and rebuilt) once, not per slide. */
  pictures: Map<string, WebPicture>;
}

/**
 * A picture pptxtojson loaded (`data:<type by extension>;base64,…`) as a data: URI the webview can show: the type is
 * taken from the bytes (pptxtojson leaves it empty for .bmp / .webp and trusts the extension), and formats a browser
 * cannot draw (EMF, WMF, TIFF) give the reason for a placeholder instead.
 */
function webPicture(uri: string, ref: string, ctx: NormaliseContext): WebPicture {
  const cached = ref ? ctx.pictures.get(ref) : undefined;
  if (cached) return cached;
  const result = checkPicture(uri, ref, ctx.lossy);
  if (ref) ctx.pictures.set(ref, result);
  return result;
}

function checkPicture(uri: string, ref: string, lossy: Set<string>): WebPicture {
  const prefix = /^data:([^;,]*);base64,/.exec(uri ?? '');
  if (!prefix) return { reason: 'Picture not available' };
  const declared = prefix[1].toLowerCase();
  const head = Buffer.from(uri.slice(prefix[0].length, prefix[0].length + 64), 'base64');
  const kind = sniffPicture(head);
  let mime: string | undefined;
  if (kind && 'mime' in kind) mime = kind.mime;
  else if (!kind && declared === 'image/svg+xml' && /^﻿?\s*</.test(head.toString('utf8'))) mime = declared;
  else if (!kind && BROWSER_IMAGE_TYPES.has(declared)) mime = declared;
  if (mime) return { uri: mime === declared ? uri : `data:${mime};base64,${uri.slice(prefix[0].length)}` };

  lossy.add(LOSSY.pictures);
  const extension = /\.([a-z0-9]+)$/i.exec(ref)?.[1];
  const format = kind && 'format' in kind ? kind.format : extension ? extension.toUpperCase() : declared || 'unknown format';
  return { reason: `Picture not shown (${format})` };
}

function escapeXml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c] ?? c);
}

/** Grey dashed box with the reason, sized like the picture it replaces (e.g. "Picture not shown (EMF)"). */
function placeholderPicture(text: string, width: number, height: number): string {
  const w = Math.max(1, Math.round(Number.isFinite(width) && width > 0 ? width : 160));
  const h = Math.max(1, Math.round(Number.isFinite(height) && height > 0 ? height : 90));
  const size = Math.max(6, Math.min(14, w / 14, h / 3));
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">` +
    `<rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" fill="#8882" stroke="#888" stroke-dasharray="4 3"/>` +
    `<text x="${w / 2}" y="${h / 2}" font-family="sans-serif" font-size="${size.toFixed(1)}" fill="#888" ` +
    `text-anchor="middle" dominant-baseline="middle">${escapeXml(text)}</text>` +
    `</svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`;
}

// ===== NORMALISATION =====

/** Removes a field the webview must not get (pptxtojson `blob:` URLs only work in the page that created them). */
function dropKey(obj: object, key: string): void {
  delete (obj as Record<string, unknown>)[key];
}

/** pptxtojson lists nodes grouped by tag; `order` (document position) restores the drawing order. */
function byOrder<T extends { order?: unknown }>(elements: readonly T[]): T[] {
  const key = (el: T): number => {
    const n = Number(el.order);
    return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
  };
  return [...elements].sort((a, b) => key(a) - key(b));
}

/** A shape / text fill: picture fills checked like pictures (a picture a browser cannot draw becomes a placeholder). */
function normaliseFill(fill: PjFill, width: number, height: number, ctx: NormaliseContext): PjFill {
  if (!fill || fill.type !== 'image' || !fill.value) return fill;
  const picture = webPicture(fill.value.base64, fill.value.ref, ctx);
  const value = { ...fill.value, base64: 'uri' in picture ? picture.uri : placeholderPicture(picture.reason, width, height) };
  dropKey(value, 'blob');
  return { ...fill, value };
}

/** The slide background, or null (white) when there is none or its picture cannot be shown. */
function normaliseBackground(fill: PjFill | undefined, ctx: NormaliseContext): PptxFill | null {
  if (!fill || !fill.value) return null;
  if (fill.type !== 'image') return fill;
  const picture = webPicture(fill.value.base64, fill.value.ref, ctx);
  if (!('uri' in picture)) return null;
  const value = { ...fill.value, base64: picture.uri };
  dropKey(value, 'blob');
  return { ...fill, value };
}

function normaliseElement(el: PjElement, ctx: NormaliseContext): PptxElement {
  switch (el.type) {
    case 'video':
    case 'audio': {
      ctx.lossy.add(el.type === 'video' ? LOSSY.video : LOSSY.audio);
      const { id, left, top, width, height, order } = el;
      return { type: 'media', mediaType: el.type, id: String(id ?? ''), left, top, width, height, order };
    }
    case 'image': {
      const picture = webPicture(el.base64, el.ref, ctx);
      const image: PjImage = { ...el, base64: 'uri' in picture ? picture.uri : placeholderPicture(picture.reason, el.width, el.height) };
      dropKey(image, 'blob');
      return image;
    }
    case 'math': {
      // An equation: LaTeX plus the picture PowerPoint saves for older readers.
      const picture = el.picBase64 ? webPicture(el.picBase64, el.picRef, ctx) : undefined;
      const math: PjMath = { ...el, picBase64: picture && 'uri' in picture ? picture.uri : '' };
      dropKey(math, 'picBlob');
      return math;
    }
    case 'shape':
    case 'text':
      return { ...el, fill: normaliseFill(el.fill, el.width, el.height, ctx) };
    case 'group':
      return { ...el, elements: normaliseElements(el.elements, ctx) as unknown as PjBaseElement[] };
    case 'diagram':
      return { ...el, elements: normaliseElements(el.elements, ctx) as unknown as Array<PjShape | PjText> };
    case 'chart':
      if (!DRAWN_CHART_TYPES.has(el.chartType)) ctx.lossy.add(LOSSY.otherCharts);
      return el;
    default:
      return el;
  }
}

function normaliseElements(elements: readonly PjElement[] | undefined, ctx: NormaliseContext, layout = false): PptxElement[] {
  return byOrder(elements ?? []).map((el) => {
    const out = normaliseElement(el, ctx);
    if (layout) out.layout = true;
    return out;
  });
}

/**
 * pptxtojson's layoutElements are the layout's own shapes followed by the master's; the master's are drawn first.
 * The layout's run is found by the cNvPr ids of the layout's shapes (ids repeat between parts, so only as a prefix).
 */
function splitLayoutElements(elements: readonly PjElement[], part: SlidePart): { master: PjElement[]; layout: PjElement[] } {
  if (part.masterShapesHidden) return { master: [], layout: [...elements] };
  const remaining = part.layoutShapes.slice();
  let count = 0;
  while (count < elements.length) {
    const id = String(elements[count].id ?? '');
    const match = remaining.findIndex((ids) => ids.includes(id));
    if (match < 0) break;
    remaining.splice(match, 1);
    count++;
  }
  return { master: elements.slice(count), layout: elements.slice(0, count) };
}

function normaliseSlide(raw: PjSlide, index: number, part: SlidePart, ctx: NormaliseContext): PptxSlide {
  const { master, layout } = splitLayoutElements(raw.layoutElements ?? [], part);
  return {
    index,
    hidden: part.hidden,
    background: normaliseBackground(raw.fill, ctx),
    elements: [...normaliseElements(master, ctx, true), ...normaliseElements(layout, ctx, true), ...normaliseElements(raw.elements, ctx)],
    notes: part.notes,
  };
}

function buildDocument(outline: DeckOutline, parsed: ParsedDeck): PptxDocument {
  const ctx: NormaliseContext = { lossy: outline.lossy, pictures: new Map() };
  // The outline mirrors pptxtojson's slide list; should they ever disagree, keep pptxtojson's order without outline.
  const aligned = outline.parts.length === parsed.slides.length;
  const order = aligned ? outline.order : parsed.slides.map((_slide, i) => i);

  const slides = order.map((i, index) => normaliseSlide(parsed.slides[i], index, (aligned && outline.parts[i]) || NO_PART, ctx));
  const infos = order.map((i, index): PptxSlideInfo => {
    const part = (aligned && outline.parts[i]) || NO_PART;
    const info: PptxSlideInfo = { index, title: part.title, hidden: part.hidden, hasNotes: part.notes.trim() !== '' };
    if (aligned && outline.names[i]) info.part = outline.names[i];
    return info;
  });

  const size = (value: number, fallback: number): number => (Number.isFinite(value) && value > 0 ? value : fallback);
  const meta: PptxDeckMeta = {
    // 10 x 7.5 inches: PowerPoint's size when presentation.xml gives none.
    width: size(parsed.size?.width, 720),
    height: size(parsed.size?.height, 540),
    slideCount: slides.length,
    slides: infos,
    themeColors: parsed.themeColors ?? [],
    fonts: [...new Set([...outline.fonts, ...(parsed.usedFonts ?? [])])],
    lossy: LOSSY_ORDER.filter((feature) => ctx.lossy.has(feature)),
  };

  return {
    meta,
    getSlide(index: number): PptxSlide {
      const slide = Number.isInteger(index) ? slides[index] : undefined;
      if (!slide) {
        throw new RangeError(`There is no slide ${index + 1}: the presentation has ${slides.length} slide${slides.length === 1 ? '' : 's'}.`);
      }
      return slide;
    },
  };
}

// ===== ERRORS =====

const ZIP_SIGNATURE = [0x50, 0x4b];
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function startsWith(data: Uint8Array, sig: number[]): boolean {
  if (data.length < sig.length) return false;
  return sig.every((b, i) => data[i] === b);
}

/**
 * An OLE2 compound file renamed .pptx: an encrypted (password-protected) Office file keeps its package in an
 * `EncryptionInfo` / `EncryptedPackage` pair of streams; a PowerPoint 97-2003 file has a `PowerPoint Document` stream.
 * Stream names are UTF-16 in the compound file's directory.
 */
function describeCompoundFile(data: Uint8Array): string {
  const bytes = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  if (bytes.includes(Buffer.from('EncryptionInfo', 'utf16le'))) {
    return (
      'This presentation is password-protected (encrypted), so FileStudio cannot open it. Open it in PowerPoint, ' +
      'remove the password (File > Info > Protect Presentation > Encrypt with Password) and save it again.'
    );
  }
  if (bytes.includes(Buffer.from('PowerPoint Document', 'utf16le'))) {
    return (
      'This is a legacy PowerPoint 97-2003 presentation (.ppt) renamed to .pptx, which FileStudio cannot open. ' +
      'Open it in PowerPoint and save it as a .pptx presentation.'
    );
  }
  return (
    'This file is not an Office Open XML (.pptx) package. It is either password-protected/encrypted or a legacy ' +
    'PowerPoint 97-2003 (.ppt) file renamed to .pptx. Open it in PowerPoint and save it as a .pptx presentation.'
  );
}

/** Up-front check that gives users a clear reason instead of a zip parser error. */
function assertLooksLikePptx(data: Uint8Array): void {
  if (data.length === 0) throw new Error('The file is empty (0 bytes), so there is nothing to display.');
  if (startsWith(data, OLE_SIGNATURE)) throw new Error(describeCompoundFile(data));
  if (!startsWith(data, ZIP_SIGNATURE)) {
    throw new Error('This file is not a valid .pptx presentation (it is not a ZIP package). It may be corrupted or in a different format.');
  }
}

/** The ZIP opened: check it holds a presentation pptxtojson can read (it expects these exact part names). */
async function assertPresentationPackage(pkg: PptxPackage): Promise<void> {
  if (!pkg.has('ppt/presentation.xml')) {
    if (pkg.has('word/document.xml')) {
      throw new Error('This file is not a valid .pptx presentation: its ZIP package holds a Word document (it may be a .docx file renamed to .pptx).');
    }
    if (pkg.has('xl/workbook.xml')) {
      throw new Error('This file is not a valid .pptx presentation: its ZIP package holds an Excel workbook (it may be an .xlsx file renamed to .pptx).');
    }
    throw new Error(
      'This file is not a valid .pptx presentation: its ZIP package has no PowerPoint presentation part (it may be another kind of file renamed to .pptx).',
    );
  }
  if (!pkg.has('[Content_Types].xml')) {
    throw new Error('This file is not a valid .pptx presentation: its ZIP package has no content-type list, so it is damaged or incomplete.');
  }
  const presentation = (await pkg.text('ppt/presentation.xml')) ?? '';
  if (/purl\.oclc\.org\/ooxml\/presentationml\/main/.test(presentation.slice(0, 4000))) {
    throw new Error(
      'This presentation is saved in the Strict Open XML format, which FileStudio cannot read. Open it in PowerPoint and ' +
        'save it as a standard PowerPoint Presentation (.pptx).',
    );
  }
}

/** jszip (and the inflate code it uses) on a truncated or damaged package. */
const ZIP_DAMAGED_RE =
  /corrupted zip|central dir|end of data reached|invalid signature|size mismatch|crc32|missing \d+ bytes|is this a zip file|invalid (?:stored block|block type|distance|code|literal)|incorrect header check|unexpected end of (?:file|data)/i;

/**
 * Readable reason for a parse failure, worded like assertLooksLikePptx. The library's own text ("Cannot read
 * properties of undefined (reading 'attrs')") is never shown: it stays available as the error's `cause`.
 */
function describeFailure(err: unknown): string {
  const raw = (err instanceof Error ? err.message : String(err)).replace(/\s+/g, ' ').trim();
  if (/encrypted zip/i.test(raw)) {
    return 'This file is not a valid .pptx presentation: its ZIP package is encrypted.';
  }
  if (ZIP_DAMAGED_RE.test(raw)) {
    return 'This file is not a valid .pptx presentation: its ZIP package is damaged (truncated or corrupted), so it cannot be read.';
  }
  return 'This presentation could not be read: part of its content is missing, damaged or not valid PowerPoint XML.';
}

/** The bytes as an ArrayBuffer for pptxtojson (no copy when `data` spans its whole buffer). */
function toArrayBuffer(data: Uint8Array): ArrayBuffer {
  if (data.byteOffset === 0 && data.byteLength === data.buffer.byteLength && data.buffer instanceof ArrayBuffer) return data.buffer;
  return data.slice().buffer;
}

// ===== PUBLIC API =====

/**
 * Parses a .pptx file once into the deck outline and normalised slides for the read-only presentation view. Throws an
 * Error with a user-facing message when the file cannot be read.
 */
export async function renderPptx(data: Uint8Array): Promise<PptxDocument> {
  assertLooksLikePptx(data);

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }
  const pkg = new PptxPackage(zip);
  await assertPresentationPackage(pkg);

  let outline: DeckOutline;
  let parsed: ParsedDeck;
  try {
    [outline, parsed] = await Promise.all([new OutlineReader(pkg).read(), parseDeck(toArrayBuffer(data), PARSE_OPTIONS)]);
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }
  try {
    return buildDocument(outline, parsed);
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }
}
