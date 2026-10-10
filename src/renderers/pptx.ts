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
 *    (0.1.2) it may hold `<br>`, tabs, styled spans and `<span class="pptx-bullet">` (bullet / number of a paragraph,
 *    lists are no longer `<ul>` / `<ol>`); text / shape elements may have `vert270: true` (vertical text read bottom
 *    to top, with `isVertical`); table cells may have `margin: { t, r, b, l }` (points);
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

/**
 * (0.1.2, additive) Chart settings read from the chart part, added to a `chart` element (pptxtojson reads only the
 * plot data). Series names and categories in `data` are decoded text; date categories are formatted.
 */
export interface PptxChartExtras {
  /** Chart title (missing: no title). */
  title?: string;
  /** Title text size in points. */
  titleSize?: number;
  /** Legend position (c:legendPos), or 'none' when the chart has no legend. */
  legend?: 'b' | 't' | 'l' | 'r' | 'tr' | 'none';
  /** Number format code (Excel) of the value axis / the secondary value axis; missing for General. */
  valueFormat?: string;
  valueFormat2?: string;
  /** Major gridlines on the value axis. */
  gridlines?: boolean;
  /** Text size in points (c:txPr of the chart or its axes). */
  fontSize?: number;
  /** Data labels per series (same order as `data`); null when the series shows none. */
  dataLabels?: Array<PptxDataLabels | null>;
  /** Combo charts: plot type of each series ('barChart', 'lineChart', 'areaChart', …) and its secondary-axis flag. */
  seriesTypes?: string[];
  secondary?: boolean[];
  /** Series names of a scatter chart (its `data` has none). */
  seriesNames?: string[];
}

export interface PptxDataLabels {
  value: boolean;
  percent: boolean;
  category: boolean;
  series: boolean;
  /** Number format code of the values (missing: the series' own format). */
  format?: string;
}

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
const PARSE_OPTIONS = {
  imageMode: 'base64',
  videoMode: 'none',
  audioMode: 'none',
  // PowerPoint's single line is 1.2 x the font size: a 200% line spacing is 2.4 x the font size.
  singleLineSpacingFactor: 1.2,
} as const;

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
  missingPictures: 'Pictures missing from the file',
  damagedSlides: 'Damaged slides (shown empty)',
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

/** A slide part pptxtojson can read: a whole `<p:sld>` with its shape tree (not a missing, cut or non-XML part). */
function isReadableSlide(xml: string): boolean {
  return /<p:sld[\s>]/.test(xml) && /<p:spTree[\s>]/.test(xml) && /<\/p:sld>\s*$/.test(xml);
}

const EMU_PER_PT = 12700;

/** Position and size (points) of a frame or picture from its first `<p:xfrm>` / `<a:xfrm>`. */
function frameBox(xml: string): { left: number; top: number; width: number; height: number } | undefined {
  const xfrm = /<([pa]):xfrm\b[^>]*>([\s\S]*?)<\/\1:xfrm>/.exec(xml)?.[2];
  if (!xfrm) return undefined;
  const off = firstTagAttrs(xfrm, 'a:off');
  const ext = firstTagAttrs(xfrm, 'a:ext');
  const pt = (value: string | undefined): number => (Number(value) || 0) / EMU_PER_PT;
  return { left: pt(off.x), top: pt(off.y), width: pt(ext.cx), height: pt(ext.cy) };
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

  /** A binary part (a picture) as base64, or null when it is missing or cannot be inflated. */
  base64(path: string): Promise<string | null> {
    const file = this.zip.file(path);
    return file ? file.async('base64').catch(() => null) : Promise.resolve(null);
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
  /** The slide part is missing or not readable XML (repairPackage gives pptxtojson an empty slide instead). */
  damaged?: boolean;
  /** Charts of the slide by the cNvPr id of their frame (read from the chart parts). */
  charts?: Map<string, ChartPartInfo>;
  /**
   * Links of shapes and pictures pptxtojson drops, by cNvPr id: a slide jump action ('ppaction://hlinkshowjump?…',
   * followed by the webview) or another slide's package part.
   */
  links?: Map<string, string>;
  /** Frames pptxtojson draws nothing for (OLE objects, linked pictures), shown as a picture or a labelled box. */
  frames?: FramePlaceholder[];
}

/** A frame pptxtojson skips, drawn as its preview picture or as a labelled box. Units are points. */
interface FramePlaceholder {
  id: string;
  label: string;
  left: number;
  top: number;
  width: number;
  height: number;
  /** Preview picture as a data: URI and its package path (checked like any picture). */
  picture?: { uri: string; ref: string };
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

/** `<p:cNvPr …>…</p:cNvPr>` with children (a self-closing one has no link). */
const CNVPR_RE = /<p:cNvPr\b((?:[^>"']|"[^"]*"|'[^']*')*?)(?<!\/)>([\s\S]*?)<\/p:cNvPr>/g;

/**
 * Links of shapes and pictures that pptxtojson drops (it keeps only http(s) hyperlinks): slide jump actions (Next /
 * Previous / First / Last slide) and links to another slide of the deck, by cNvPr id.
 */
function shapeLinks(xml: string, rels: Relationship[]): Map<string, string> {
  const links = new Map<string, string>();
  for (const m of xml.matchAll(CNVPR_RE)) {
    const id = attrsOf(m[1]).id;
    const click = firstTagAttrs(m[2], 'a:hlinkClick');
    if (!id) continue;
    const action = click.action ?? '';
    const rel = click['r:id'] ? rels.find((r) => r.id === click['r:id']) : undefined;
    if (/^ppaction:\/\/hlinkshowjump\?jump=\w+/i.test(action)) links.set(id, action);
    else if (rel && !rel.external && rel.type === 'slide') links.set(id, rel.target);
  }
  return links;
}

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
    const damaged = !isReadableSlide(xml);
    if (damaged) this.lossy.add(LOSSY.damagedSlides);
    return {
      hidden: isHiddenSlide(xml),
      title: slideTitle(xml),
      notes: notesXml ? notesText(notesXml) : '',
      layoutShapes: layout.shapes,
      masterShapesHidden: layout.masterShapesHidden,
      damaged,
      charts: await this.readCharts(xml, rels),
      links: shapeLinks(xml, rels),
      frames: await this.readFrames(xml, rels),
    };
  }

  /** The slide's charts (frames anywhere on the slide, groups included) by frame id. */
  private async readCharts(xml: string, rels: Relationship[]): Promise<Map<string, ChartPartInfo>> {
    const charts = new Map<string, ChartPartInfo>();
    for (const m of xml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
      const rid = firstTagAttrs(m[0], 'c:chart')['r:id'];
      const id = firstTagAttrs(m[0], 'p:cNvPr').id;
      const part = rid ? rels.find((rel) => rel.id === rid && !rel.external)?.target : undefined;
      const chartXml = part ? await this.pkg.text(part) : null;
      if (id && chartXml) charts.set(id, readChartPart(chartXml));
    }
    return charts;
  }

  /** Top-level OLE objects, linked pictures and SmartArt without a drawing (pptxtojson draws nothing for them). */
  private async readFrames(xml: string, rels: Relationship[]): Promise<FramePlaceholder[]> {
    const frames: FramePlaceholder[] = [];
    const noDrawing = !rels.some((rel) => rel.type === 'diagramDrawing');
    for (const child of spTreeChildren(xml)) {
      const ole = /presentationml\/2006\/ole"/.test(child.xml);
      const linked = child.name === 'p:pic' && /<a:blip\b(?=[^>]*\br:link=)(?![^>]*\br:embed=)/.test(child.xml);
      const smartArt = noDrawing && child.name === 'p:graphicFrame' && DIAGRAM_FRAME_RE.test(child.xml);
      if (!ole && !linked && !smartArt) continue;
      const box = frameBox(child.xml);
      const id = firstTagAttrs(child.xml, 'p:cNvPr').id ?? '';
      if (!box) continue;
      if (linked || smartArt) {
        frames.push({ id, label: linked ? 'Linked picture not loaded' : 'SmartArt not shown', ...box });
        continue;
      }
      const app = /^(Excel|Word|PowerPoint|Visio|Acrobat)/i.exec(firstTagAttrs(child.xml, 'p:oleObj').progId ?? '')?.[1];
      const frame: FramePlaceholder = { id, label: app ? `Embedded ${app} object not shown` : 'Embedded object not shown', ...box };
      // The preview picture Office saves with the object (often EMF, which a browser cannot draw).
      const rid = firstTagAttrs(child.xml, 'a:blip')['r:embed'];
      const ref = rid ? rels.find((rel) => rel.id === rid && !rel.external)?.target : undefined;
      const data = ref ? await this.pkg.base64(ref) : null;
      if (ref && data) frame.picture = { uri: `data:${mimeOfPart(ref)};base64,${data}`, ref };
      frames.push(frame);
    }
    return frames;
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

// ===== CHARTS (read from the chart part) =====

/** The part of ssf (SheetJS number formatter, untyped CommonJS) used for chart categories. */
interface SsfLib {
  format(fmt: string, value: number, opts?: { date1904?: boolean }): string;
  is_date(fmt: string): boolean;
}
// eslint-disable-next-line @typescript-eslint/no-require-imports
const SSF = require('ssf') as SsfLib;

/** Plot types drawn on a category axis: a chart mixing them (a combo chart) is drawn as one. */
const CATEGORY_PLOTS: ReadonlySet<string> = new Set(['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'area3DChart']);

/** One series of a chart part, every plot group included (pptxtojson keeps only the last group). */
interface ChartSeriesPart {
  /** Plot group type ('barChart', 'lineChart', …). */
  plot: string;
  /** Plotted against the secondary value axis. */
  secondary: boolean;
  key: string;
  values: Array<{ x: string; y: number }>;
  xlabels: Record<string, string>;
  /** '#RRGGBB', or 'accent1'…'accent6' for a theme colour; '' when not set. */
  color: string;
  labels: PptxDataLabels | null;
}

interface ChartPartInfo {
  series: ChartSeriesPart[];
  /** Plot groups that have series (more than one: a combo chart). */
  plots: number;
  /** grouping / barDir of the first plot group. */
  grouping?: string;
  barDir?: string;
  extras: PptxChartExtras;
}

/** Inner markup of the first `<name>` element (lazy: for elements that do not nest), '' when there is none. */
function elementXml(xml: string, name: string): string {
  const m = new RegExp(`<${name}(?=[\\s/>])[^>]*?(?:/>|>([\\s\\S]*?)</${name}>)`).exec(xml);
  return m ? (m[1] ?? '') : '';
}

function hasElement(xml: string, name: string): boolean {
  return new RegExp(`<${name}(?=[\\s/>])`).test(xml);
}

/** `val` of the first `<name val="…"/>`. */
function valOf(xml: string, name: string): string | undefined {
  return firstTagAttrs(xml, name).val;
}

/** A boolean element (`<c:showVal val="1"/>`; a missing `val` is true). */
function flagOf(xml: string, name: string): boolean {
  if (!hasElement(xml, name)) return false;
  const val = valOf(xml, name);
  return val === undefined || val === '1' || val === 'true';
}

/** Cached points of a `c:strCache` / `c:numCache` / literal: text by idx, plus the cache's number format. */
function cachePoints(xml: string): { points: Map<string, string>; format?: string } {
  const points = new Map<string, string>();
  for (const m of xml.matchAll(/<c:pt\b([^>]*)>([\s\S]*?)<\/c:pt>/g)) {
    const idx = attrsOf(m[1]).idx;
    const v = /<c:v>([\s\S]*?)<\/c:v>/.exec(m[2]);
    if (idx !== undefined && v) points.set(idx, decodeXml(v[1]));
  }
  const format = /<c:formatCode>([\s\S]*?)<\/c:formatCode>/.exec(xml)?.[1];
  return { points, format: format === undefined ? undefined : decodeXml(format) };
}

/** A number format other than General (null / General: none). */
function realFormat(format: string | undefined): string | undefined {
  return format && format.trim() && format.trim().toLowerCase() !== 'general' ? format : undefined;
}

/** Categories with a number format (dates are stored as serial numbers) shown as the format shows them. */
function formatCategory(text: string, format: string | undefined, date1904: boolean): string {
  const value = Number(text);
  if (!format || text.trim() === '' || !Number.isFinite(value)) return text;
  try {
    return SSF.format(format, value, { date1904 });
  } catch {
    return text;
  }
}

/** Data labels of a `c:dLbls` (per-point `c:dLbl` overrides skipped); null when none is shown. */
function dataLabelsOf(xml: string, valueFormat: string | undefined): PptxDataLabels | null {
  const body = xml.replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/g, '');
  if (flagOf(body, 'c:delete')) return null;
  const labels: PptxDataLabels = {
    value: flagOf(body, 'c:showVal'),
    percent: flagOf(body, 'c:showPercent'),
    category: flagOf(body, 'c:showCatName'),
    series: flagOf(body, 'c:showSerName'),
  };
  if (!labels.value && !labels.percent && !labels.category && !labels.series) return null;
  const numFmt = firstTagAttrs(body, 'c:numFmt');
  const format = realFormat(numFmt.sourceLinked === '1' ? undefined : numFmt.formatCode) ?? valueFormat;
  if (format) labels.format = format;
  return labels;
}

/** Colour of a series (`c:spPr`: fill, or the line of a line series): '#RRGGBB' or a theme accent name. */
function seriesColor(serXml: string): string {
  const spPr = elementXml(serXml, 'c:spPr');
  const fill = elementXml(spPr, 'a:solidFill') || elementXml(elementXml(spPr, 'a:ln'), 'a:solidFill');
  const rgb = valOf(fill, 'a:srgbClr');
  if (rgb && /^[0-9a-f]{6}$/i.test(rgb)) return `#${rgb}`;
  const scheme = valOf(fill, 'a:schemeClr');
  return scheme && /^accent[1-6]$/.test(scheme) ? scheme : '';
}

/** Text size (points) of the first `a:defRPr sz` in `xml`. */
function textSize(xml: string): number | undefined {
  const sz = Number(firstTagAttrs(xml, 'a:defRPr').sz);
  return Number.isFinite(sz) && sz > 0 ? sz / 100 : undefined;
}

/**
 * What pptxtojson does not read from a chart part: every plot group (combo charts), decoded series names and
 * categories (dates formatted), title, legend, axis number formats, gridlines, text size and data labels.
 */
function readChartPart(xml: string): ChartPartInfo {
  const date1904 = flagOf(xml, 'c:date1904');
  const chart = elementXml(xml, 'c:chart');
  const plotArea = elementXml(chart, 'c:plotArea');
  const valAxes = [...plotArea.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/g)].map((m) => m[1]);
  const groups = [...plotArea.matchAll(/<c:(\w+Chart)>([\s\S]*?)<\/c:\1>/g)]
    .map((m) => ({ plot: m[1], xml: m[2], axes: [...m[2].matchAll(/<c:axId\b[^>]*>/g)].map((a) => attrsOf(a[0]).val ?? '') }))
    .filter((g) => /<c:ser>/.test(g.xml));
  const first = groups[0];
  const horizontal = first !== undefined && valOf(first.xml, 'c:barDir') === 'bar';
  /** The value axis of a plot group (scatter: its vertical one). */
  const valueAxisOf = (axes: string[]): string | undefined => {
    const own = valAxes.filter((ax) => axes.includes(valOf(ax, 'c:axId') ?? ''));
    const wanted = horizontal ? ['b', 't'] : ['l', 'r'];
    return own.find((ax) => wanted.includes(valOf(ax, 'c:axPos') ?? '')) ?? own[0];
  };

  const series: ChartSeriesPart[] = [];
  const seriesFormats: Array<string | undefined> = [];
  for (const group of groups) {
    const secondary = group !== first && group.axes.join() !== first.axes.join();
    const groupLabels = elementXml(group.xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ''), 'c:dLbls');
    for (const m of group.xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
      const ser = m[1];
      const tx = elementXml(ser, 'c:tx');
      const name = cachePoints(tx).points.get('0') ?? decodeXml(/<c:v>([\s\S]*?)<\/c:v>/.exec(tx)?.[1] ?? '');
      const cat = cachePoints(elementXml(ser, 'c:cat') || elementXml(ser, 'c:xVal'));
      const val = cachePoints(elementXml(ser, 'c:val') || elementXml(ser, 'c:yVal'));
      const catFormat = realFormat(cat.format);
      const xlabels: Record<string, string> = {};
      for (const [idx, text] of cat.points) xlabels[idx] = formatCategory(text, catFormat, date1904);
      const valueFormat = realFormat(val.format);
      seriesFormats.push(valueFormat);
      const own = elementXml(ser, 'c:dLbls');
      series.push({
        plot: group.plot,
        secondary,
        key: name,
        values: [...val.points].map(([idx, text]) => ({ x: idx, y: parseFloat(text) })),
        xlabels,
        color: seriesColor(ser),
        labels: dataLabelsOf(hasElement(ser, 'c:dLbls') ? own : groupLabels, valueFormat),
      });
    }
  }

  const extras: PptxChartExtras = {};
  // The chart title (c:chart/c:title, before the plot area; axis titles are inside the axes).
  const head = chart.slice(0, Math.max(0, chart.search(/<c:plotArea[\s>]/)));
  if (hasElement(head, 'c:title') && !flagOf(head, 'c:autoTitleDeleted')) {
    const title = elementXml(head, 'c:title');
    const text = paragraphs(elementXml(title, 'c:tx')).join('\n').trim();
    // A title without text is PowerPoint's automatic one: the series name of a one-series chart.
    extras.title = text || (series.length === 1 && series[0].key ? series[0].key : 'Chart Title');
    const size = textSize(title);
    if (size) extras.titleSize = size;
  }
  const legend = elementXml(chart, 'c:legend');
  extras.legend = hasElement(chart, 'c:legend') ? ((valOf(legend, 'c:legendPos') ?? 'r') as PptxChartExtras['legend']) : 'none';
  /** Axis number format: its own, or (source-linked) the format of its first series. */
  const axisFormat = (ax: string | undefined, seriesFormat: string | undefined): string | undefined => {
    if (!ax) return undefined;
    const numFmt = firstTagAttrs(ax, 'c:numFmt');
    return numFmt.sourceLinked === '1' || !numFmt.formatCode ? seriesFormat : realFormat(numFmt.formatCode);
  };
  if (first) {
    const primary = valueAxisOf(first.axes);
    const format = axisFormat(primary, seriesFormats[0]);
    if (format) extras.valueFormat = format;
    extras.gridlines = primary ? hasElement(primary, 'c:majorGridlines') : true;
    const index2 = series.findIndex((s) => s.secondary);
    const second = index2 >= 0 ? groups.find((g) => g.plot === series[index2].plot && g !== first) : undefined;
    const format2 = second ? axisFormat(valueAxisOf(second.axes), seriesFormats[index2]) : undefined;
    if (format2) extras.valueFormat2 = format2;
  }
  // Text size: the chart's own text properties, else those of the legend / axes.
  const tail = xml.slice(xml.lastIndexOf('</c:chart>'));
  const size = textSize(elementXml(tail, 'c:txPr')) ?? textSize(plotArea.replace(/<c:\w+Chart>[\s\S]*?<\/c:\w+Chart>/g, '')) ?? textSize(legend);
  if (size) extras.fontSize = size;

  return {
    series,
    plots: groups.length,
    grouping: first ? valOf(first.xml, 'c:grouping') : undefined,
    barDir: first ? valOf(first.xml, 'c:barDir') : undefined,
    extras,
  };
}

// ===== PICTURES =====

/** Declared types a browser shows when the bytes match no known signature (pptxtojson names them by extension). */
const BROWSER_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/avif']);

/** Declared picture type of a package part, by extension (like pptxtojson; checkPicture then reads the bytes). */
function mimeOfPart(path: string): string {
  const ext = (/\.([a-z0-9]+)$/i.exec(path)?.[1] ?? '').toLowerCase();
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'svg') return 'image/svg+xml';
  return ext ? `image/${ext}` : '';
}

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
  /** Pictures missing from the package (repairPackage put an empty part in their place for pptxtojson). */
  missing: ReadonlySet<string>;
  /** Theme accent colours ('#RRGGBB', accent1 first). */
  themeColors: string[];
  /** The slide whose own elements are being normalised (undefined for layout and master elements). */
  part?: SlidePart;
}

/**
 * A picture pptxtojson loaded (`data:<type by extension>;base64,…`) as a data: URI the webview can show: the type is
 * taken from the bytes (pptxtojson leaves it empty for .bmp / .webp and trusts the extension), and formats a browser
 * cannot draw (EMF, WMF, TIFF) give the reason for a placeholder instead.
 */
function webPicture(uri: string, ref: string, ctx: NormaliseContext): WebPicture {
  const cached = ref ? ctx.pictures.get(ref) : undefined;
  if (cached) return cached;
  if (ref && ctx.missing.has(ref)) {
    ctx.lossy.add(LOSSY.missingPictures);
    return { reason: 'Picture missing from the file' };
  }
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

/**
 * A frame pptxtojson skips (OLE object, linked picture, SmartArt without drawing, damaged slide) as a picture: its
 * preview picture when a browser can show it, else a labelled placeholder box.
 */
function frameElement(frame: FramePlaceholder, order: number): PptxElement {
  const picture = frame.picture ? checkPicture(frame.picture.uri, frame.picture.ref, new Set()) : undefined;
  return {
    type: 'image',
    id: frame.id,
    name: frame.label,
    left: frame.left,
    top: frame.top,
    width: frame.width,
    height: frame.height,
    ref: frame.picture?.ref ?? '',
    base64: picture && 'uri' in picture ? picture.uri : placeholderPicture(frame.label, frame.width, frame.height),
    rotate: 0,
    isFlipH: false,
    isFlipV: false,
    order,
    geom: 'rect',
    borderColor: '',
    borderWidth: 0,
    borderType: 'solid',
    borderStrokeDasharray: '',
  } as Omit<PjImage, 'blob'> as PptxElement;
}

/**
 * A chart with what pptxtojson leaves out (PptxChartExtras): decoded names, formatted date categories and, for a combo
 * chart, the series of every plot group (pptxtojson keeps the last group only).
 */
function normaliseChart(el: Extract<PjElement, { type: 'chart' }>, ctx: NormaliseContext): PptxElement {
  if (!DRAWN_CHART_TYPES.has(el.chartType)) ctx.lossy.add(LOSSY.otherCharts);
  const info = ctx.part?.charts?.get(String(el.id ?? ''));
  if (!info) return el;
  const out: Record<string, unknown> = { ...el, ...info.extras };
  if (el.chartType === 'scatterChart' || el.chartType === 'bubbleChart') {
    out.seriesNames = info.series.map((s) => s.key);
    out.dataLabels = info.series.map((s) => s.labels);
    return out as unknown as PptxElement;
  }
  const combo = info.plots > 1 && CATEGORY_PLOTS.has(el.chartType) && info.series.every((s) => CATEGORY_PLOTS.has(s.plot));
  const series = combo ? info.series : info.series.filter((s) => s.plot === el.chartType);
  if (series.length === 0 || (!combo && series.length !== el.data.length)) return out as unknown as PptxElement;
  out.data = series.map((s) => ({ key: s.key, values: s.values, xlabels: s.xlabels }));
  out.dataLabels = series.map((s) => s.labels);
  if (combo) {
    const accent = (color: string): string => (/^accent([1-6])$/.exec(color) ? (ctx.themeColors[Number(color.slice(6)) - 1] ?? '') : color);
    out.chartType = series[0].plot;
    out.grouping = info.grouping;
    out.barDir = info.barDir;
    out.colors = series.map((s) => accent(s.color));
    out.seriesTypes = series.map((s) => s.plot);
    out.secondary = series.map((s) => s.secondary);
  }
  return out as unknown as PptxElement;
}

function normaliseElement(el: PjElement, ctx: NormaliseContext): PptxElement {
  const part = ctx.part;
  const id = String(el.id ?? '');
  const frame = part?.frames?.find((f) => f.id === id && id !== '');
  if (frame && part?.frames) {
    // pptxtojson's own element for the frame (an OLE object's fallback group) is replaced by the frame.
    part.frames = part.frames.filter((f) => f !== frame);
    return frameElement(frame, el.order);
  }
  const link = part?.links?.get(id);
  if (link && (el.type === 'shape' || el.type === 'text' || el.type === 'image') && !el.link) el = { ...el, link };
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
      return normaliseChart(el, ctx);
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

function normaliseSlide(raw: PjSlide, index: number, part: SlidePart, ctx: NormaliseContext, size: { width: number; height: number }): PptxSlide {
  const { master, layout } = splitLayoutElements(raw.layoutElements ?? [], part);
  const layoutElements = [...normaliseElements(master, ctx, true), ...normaliseElements(layout, ctx, true)];
  // The slide's own charts, links and frames apply to its own elements only (ids repeat between parts).
  const own: SlidePart = { ...part, frames: [...(part.frames ?? [])] };
  ctx.part = own;
  const elements = normaliseElements(raw.elements, ctx);
  ctx.part = undefined;
  // Frames pptxtojson drew nothing for go on top, in document order.
  for (const frame of own.frames ?? []) elements.push(frameElement(frame, Number.MAX_SAFE_INTEGER));
  if (part.damaged) {
    elements.push(frameElement({ id: '', label: 'This slide is damaged and cannot be shown', left: 0, top: 0, ...size }, Number.MAX_SAFE_INTEGER));
  }
  return {
    index,
    hidden: part.hidden,
    background: normaliseBackground(raw.fill, ctx),
    elements: [...layoutElements, ...elements],
    notes: part.notes,
  };
}

function buildDocument(outline: DeckOutline, parsed: ParsedDeck, missing: ReadonlySet<string>): PptxDocument {
  const ctx: NormaliseContext = { lossy: outline.lossy, pictures: new Map(), missing, themeColors: parsed.themeColors ?? [] };
  // The outline mirrors pptxtojson's slide list; should they ever disagree, keep pptxtojson's order without outline.
  const aligned = outline.parts.length === parsed.slides.length;
  const order = aligned ? outline.order : parsed.slides.map((_slide, i) => i);

  const size = (value: number, fallback: number): number => (Number.isFinite(value) && value > 0 ? value : fallback);
  // 10 x 7.5 inches: PowerPoint's size when presentation.xml gives none.
  const slideSize = { width: size(parsed.size?.width, 720), height: size(parsed.size?.height, 540) };
  const slides = order.map((i, index) => normaliseSlide(parsed.slides[i], index, (aligned && outline.parts[i]) || NO_PART, ctx, slideSize));
  const infos = order.map((i, index): PptxSlideInfo => {
    const part = (aligned && outline.parts[i]) || NO_PART;
    const info: PptxSlideInfo = { index, title: part.title, hidden: part.hidden, hasNotes: part.notes.trim() !== '' };
    if (aligned && outline.names[i]) info.part = outline.names[i];
    return info;
  });

  const meta: PptxDeckMeta = {
    width: slideSize.width,
    height: slideSize.height,
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

// ===== REPAIR (before pptxtojson) =====

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** A slide with nothing on it: stands in for a slide part that is missing or not readable XML. */
const EMPTY_SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>' +
  '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>';

/**
 * pptxtojson fails on the whole deck for one missing picture or one damaged slide, drops SVG-only pictures and the
 * slide jump actions of text runs. The package is patched for it (only when needed: re-zipping costs time):
 *  - a missing or unreadable slide part becomes an empty slide (the outline marks it damaged);
 *  - a missing picture part becomes an empty part (normalisation shows a "missing" placeholder);
 *  - `<a:blip>` with only the SVG extension gets that SVG as its picture;
 *  - `<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=…">` gets a relationship whose target is the action
 *    after a '#' ('#ppaction://…'), which pptxtojson then writes as the link of the text run (the webview follows it).
 * Returns the bytes for pptxtojson and the package paths of the missing pictures.
 */
async function repairPackage(zip: JSZip, pkg: PptxPackage, data: Uint8Array): Promise<{ data: ArrayBuffer; missing: Set<string> }> {
  const missing = new Set<string>();
  let changed = false;
  const layouts = pkg.names.filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name));
  for (const path of await parsedSlideOrder(pkg)) {
    const slash = path.lastIndexOf('/');
    const relsPath = `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
    let relsXml = await pkg.text(relsPath);
    if (relsXml === null && layouts.length > 0) {
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/slideLayout" Target="../slideLayouts/${layouts[0].slice(layouts[0].lastIndexOf('/') + 1)}"/></Relationships>`;
      zip.file(relsPath, relsXml);
      changed = true;
    }
    const xml = (await pkg.text(path)) ?? '';
    if (!isReadableSlide(xml)) {
      zip.file(path, EMPTY_SLIDE);
      changed = true;
      continue;
    }
    let added = 0;
    const newRels: string[] = [];
    const fixed = xml
      .replace(/<a:blip\b([^>]*?)(\/?)>/g, (tag: string, attrs: string, selfClosing: string, offset: number) => {
        if (selfClosing || /\br:(?:embed|link)=/.test(attrs)) return tag;
        const end = xml.indexOf('</a:blip>', offset);
        const svg = end > 0 ? /<\w+:svgBlip\b[^>]*\br:embed="([^"]+)"/.exec(xml.slice(offset, end)) : null;
        return svg ? `<a:blip r:embed="${svg[1]}"${attrs}>` : tag;
      })
      .replace(/<a:hlinkClick\b([^>]*?)(\/?)>/g, (tag: string, attrs: string, selfClosing: string) => {
        const a = attrsOf(tag);
        if (a['r:id'] || !/^ppaction:\/\/hlinkshowjump\?jump=\w+$/i.test(a.action ?? '')) return tag;
        const id = `rIdFsJump${++added}`;
        // As an in-page '#' link: the webview's sanitizer keeps it (it drops the unknown ppaction: scheme).
        newRels.push(`<Relationship Id="${id}" Type="${REL_NS}/hyperlink" Target="#${escapeXml(a.action)}" TargetMode="External"/>`);
        return `<a:hlinkClick r:id="${id}"${attrs.replace(/\s*\br:id\s*=\s*(?:"[^"]*"|'[^']*')/, '')}${selfClosing}>`;
      });
    if (fixed !== xml) {
      zip.file(path, fixed);
      changed = true;
    }
    if (newRels.length > 0 && relsXml) {
      zip.file(relsPath, relsXml.replace(/<\/Relationships>\s*$/, `${newRels.join('')}</Relationships>`));
    }
  }
  // Pictures of slides, layouts and masters that are not in the package.
  for (const name of pkg.names) {
    const m = /^(ppt\/(?:slides|slideLayouts|slideMasters)\/)_rels\/([^/]+)\.rels$/i.exec(name);
    if (!m) continue;
    for (const rel of await pkg.rels(m[1] + m[2])) {
      if (rel.type !== 'image' || rel.external || !rel.target || pkg.has(rel.target) || missing.has(rel.target)) continue;
      missing.add(rel.target);
      zip.file(rel.target, new Uint8Array(0));
      changed = true;
    }
  }
  return { data: changed ? toArrayBuffer(await zip.generateAsync({ type: 'uint8array' })) : toArrayBuffer(data), missing };
}

// ===== TEXT MARKUP (around pptxtojson) =====
//
// pptxtojson drops soft line breaks (<a:br>), tabs (4 spaces), highlights, underline styles other than single,
// bullets inherited from the layout / master, numbering schemes and start values, the theme hyperlink colour, 90° /
// 270° vertical text and outlines without an explicit width, and it blends tints / shades in HSL (PowerPoint blends in
// linear RGB). Before parsing, the slide, layout, master, diagram-drawing and table-style parts are patched for it
// (patchMarkup); what pptxtojson cannot carry travels as private-use marker characters in the run text and becomes HTML
// afterwards (fixMarkup).

/** Marker characters (Unicode private use; the deck's own characters in this range are removed). */
const MARK = {
  br: '\uE000',
  tab: '\uE001',
  /** open + CSS + sep … close: a styled span (highlight, underline style). */
  open: '\uE002',
  sep: '\uE003',
  close: '\uE004',
  /** cell + 'top,right,bottom,left' (pt) + sep: the margins of a table cell. */
  cell: '\uE005',
  vert: '\uE006',
  vert270: '\uE007',
  /** bullet + CSS + sep + bullet text + bulletEnd: the bullet or number of a paragraph. */
  bullet: '\uE008',
  bulletEnd: '\uE009',
} as const;
const MARK_RE = /[\uE000-\uE009]/g;

const SCHEME_NAMES = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
/** Outline width (EMU) PowerPoint draws for an outline without a width: 0.75 pt. */
const DEFAULT_LINE_EMU = 9525;

interface ThemeInfo {
  /** Scheme colours (dk1 … folHlink) as RRGGBB. */
  colors: Map<string, string>;
  /** Line widths (EMU) of the theme line styles (lnRef idx 1, 2, 3). */
  lineWidths: number[];
}

interface PlaceholderStyle {
  type: string;
  idx: string;
  /** Inner XML of the placeholder's a:lstStyle. */
  lstStyle: string;
}

interface MasterInfo {
  theme: ThemeInfo;
  clrMap: Record<string, string>;
  placeholders: PlaceholderStyle[];
  titleStyle: string;
  bodyStyle: string;
  otherStyle: string;
}

interface PartMarkup {
  master: MasterInfo;
  /** Slides: placeholders of the slide's layout (bullets are inherited through layout and master). */
  layout?: PlaceholderStyle[];
  /** presentation.xml defaultTextStyle (text that is not in a placeholder). */
  defaultStyle: string;
}

/** Bullet / numbering state of one text body: the counter of each level; whether its empty paragraphs are sized. */
type ListState = Array<{ key: string; n: number } | undefined> & { sizeEmpty?: boolean };

function readTheme(xml: string): ThemeInfo {
  const colors = new Map<string, string>();
  for (const name of SCHEME_NAMES) {
    const m = new RegExp(`<a:${name}>\\s*<a:(?:srgbClr\\b[^>]*?\\bval|sysClr\\b[^>]*?\\blastClr)="([0-9A-Fa-f]{6})"`).exec(xml);
    if (m) colors.set(name, m[1].toUpperCase());
  }
  const list = /<a:lnStyleLst>([\s\S]*?)<\/a:lnStyleLst>/.exec(xml)?.[1] ?? '';
  const lineWidths = [...list.matchAll(/<a:ln\b[^>]*>/g)].map((m) => Number(attrsOf(m[0]).w) || DEFAULT_LINE_EMU);
  return { colors, lineWidths };
}

/** Placeholders of a layout or master with their list styles. */
function placeholderStyles(xml: string): PlaceholderStyle[] {
  const out: PlaceholderStyle[] = [];
  for (const m of xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)) {
    const ph = /<p:ph\b[^>]*>/.exec(m[0]);
    if (!ph) continue;
    const a = attrsOf(ph[0]);
    out.push({ type: a.type ?? 'obj', idx: a.idx ?? '', lstStyle: /<a:lstStyle>([\s\S]*?)<\/a:lstStyle>/.exec(m[0])?.[1] ?? '' });
  }
  return out;
}

/** `<a:lvlNpPr>` (N = level + 1) of a list style, '' when there is none. */
function levelStyle(styles: string, lvl: number): string {
  if (!styles) return '';
  const n = lvl + 1;
  return new RegExp(`<a:lvl${n}pPr\\b[^>]*?(?:/>|>[\\s\\S]*?</a:lvl${n}pPr>)`).exec(styles)?.[0] ?? '';
}

/** Scheme colour name -> RRGGBB through the master's colour map (tx1 -> dk1, bg1 -> lt1, …). */
function schemeColor(name: string, markup: PartMarkup): string | undefined {
  const mapped = /^(?:tx|bg)[12]$/.test(name) ? (markup.master.clrMap[name] ?? { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' }[name]) : name;
  return markup.master.theme.colors.get(mapped ?? name);
}

// ----- colours: tint / shade in linear RGB -----

const toLinear = (c: number): number => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const fromLinear = (c: number): number => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
const clamp01 = (c: number): number => Math.min(1, Math.max(0, c));

function rgbToHsl([r, g, b]: number[]): number[] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb([h, s, l]: number[]): number[] {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number): number => {
    const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    return x < 1 / 6 ? p + (q - p) * 6 * x : x < 1 / 2 ? q : x < 2 / 3 ? p + (q - p) * (2 / 3 - x) * 6 : p;
  };
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)];
}

/**
 * A colour element with a tint or shade as an srgbClr (DrawingML: tint / shade blend toward white / black in linear
 * RGB; lum / sat in HSL), keeping its alpha. null when it cannot be worked out here (placeholder colour, other
 * transforms): pptxtojson then reads it as before.
 */
function tintedColor(tag: string, attrs: Record<string, string>, inner: string, markup: PartMarkup): string | null {
  const base = tag === 'srgbClr' ? attrs.val : tag === 'sysClr' ? attrs.lastClr : schemeColor(attrs.val ?? '', markup);
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null;
  let rgb = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16) / 255);
  const kept: string[] = [];
  for (const m of inner.matchAll(/<a:(\w+)\b[^>]*\/>/g)) {
    const raw = attrsOf(m[0]).val ?? '';
    const v = raw.endsWith('%') ? parseFloat(raw) / 100 : Number(raw) / 100000;
    if (!Number.isFinite(v) && !/^alpha/.test(m[1])) return null;
    const hsl = rgbToHsl(rgb);
    switch (m[1]) {
      case 'tint':
        rgb = rgb.map((c) => fromLinear(toLinear(c) * v + 1 - v));
        break;
      case 'shade':
        rgb = rgb.map((c) => fromLinear(toLinear(c) * v));
        break;
      case 'lumMod':
        rgb = hslToRgb([hsl[0], hsl[1], clamp01(hsl[2] * v)]);
        break;
      case 'lumOff':
        rgb = hslToRgb([hsl[0], hsl[1], clamp01(hsl[2] + v)]);
        break;
      case 'satMod':
        rgb = hslToRgb([hsl[0], clamp01(hsl[1] * v), hsl[2]]);
        break;
      case 'satOff':
        rgb = hslToRgb([hsl[0], clamp01(hsl[1] + v), hsl[2]]);
        break;
      case 'alpha':
      case 'alphaMod':
      case 'alphaOff':
        kept.push(m[0]);
        break;
      default:
        return null;
    }
    rgb = rgb.map(clamp01);
  }
  const hex = rgb
    .map((c) => Math.round(c * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
  return kept.length ? `<a:srgbClr val="${hex}">${kept.join('')}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

/** Style references (p:style, background references) are read by pptxtojson by their own paths: left as they are. */
const COLOR_RE =
  /<(a|p):(fillRef|lnRef|effectRef|fontRef|bgRef)\b[^>]*?(?<!\/)>[\s\S]*?<\/\1:\2>|<a:(schemeClr|srgbClr|sysClr)((?:\s[^>]*?)?)(?<!\/)>([\s\S]*?)<\/a:\3>/g;

function fixTints(xml: string, markup: PartMarkup): string {
  if (!/<a:(?:tint|shade)\b/.test(xml)) return xml;
  return xml.replace(COLOR_RE, (all: string, _ns?: string, ref?: string, tag?: string, attrs?: string, inner?: string) => {
    if (ref || !tag || !/<a:(?:tint|shade)\b/.test(inner ?? '') || !/^(?:\s*<a:\w+\b[^>]*\/>)*\s*$/.test(inner ?? '')) return all;
    return tintedColor(tag, attrsOf(attrs ?? ''), inner ?? '', markup) ?? all;
  });
}

// ----- outlines without a width -----

/** `<a:ln>` of a shape / picture without `w`: the theme line style width (style lnRef) or 0.75 pt, when it is drawn. */
function fixOutlines(xml: string, markup: PartMarkup): string {
  return xml.replace(/<p:(sp|cxnSp|pic)\b[^>]*>[\s\S]*?<\/p:\1>/g, (shape: string) => {
    const spPr = /<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/.exec(shape);
    if (!spPr) return shape;
    const ln = /<a:ln\b([^>]*?)(\/?)>/.exec(spPr[0]);
    if (!ln || /\bw\s*=/.test(ln[1])) return shape;
    const lnRef = Number(firstTagAttrs(shape, 'a:lnRef').idx) || 0;
    const end = ln[2] ? -1 : spPr[0].indexOf('</a:ln>', ln.index);
    const content = end < 0 ? '' : spPr[0].slice(ln.index, end);
    const filled = /<a:(?:solidFill|gradFill|pattFill)\b/.test(content);
    if (/<a:noFill\b/.test(content) || (!filled && lnRef === 0)) return shape;
    const width = lnRef > 0 ? (markup.master.theme.lineWidths[lnRef - 1] ?? DEFAULT_LINE_EMU) : DEFAULT_LINE_EMU;
    const fixed = `${spPr[0].slice(0, ln.index)}<a:ln w="${width}"${ln[1]}${ln[2]}>${spPr[0].slice(ln.index + ln[0].length)}`;
    return shape.replace(spPr[0], () => fixed);
  });
}

// ----- text: runs, bullets -----

/** Symbol-font bullet characters (Wingdings / Symbol code points) -> Unicode. */
const SYMBOL_BULLETS: Record<number, string> = {
  0x6c: '\u25CF', // l: black circle
  0x6e: '\u25A0', // n: black square
  0x71: '\u2751', // q: shadowed box
  0x75: '\u25C6', // u: black diamond
  0x76: '\u2756', // v: diamond of four diamonds
  0xa7: '\u25AA', // §: small black square
  0xb7: '\u2022', // ·: bullet (Symbol)
  0xd8: '\u27A2', // Ø: arrowhead
  0xe8: '\u2794', // è: arrow
  0xfc: '\u2714', // ü: check mark
};
const SYMBOL_FONTS = /^(?:wingdings(?: [23])?|webdings|symbol)$/i;

function alphaNumber(n: number): string {
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

function romanNumber(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ];
  let out = '';
  for (const [value, digits] of table) {
    while (n >= value) {
      out += digits;
      n -= value;
    }
  }
  return out;
}

/** The text of number `n` of an a:buAutoNum scheme ('alphaUcPeriod', 3 -> 'C.'). */
function autoNumber(scheme: string, n: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc)(ParenBoth|ParenR|Period|Plain|Minus)?/.exec(scheme);
  const kind = m?.[1] ?? 'arabic';
  let s = kind === 'arabic' ? String(n) : kind.startsWith('alpha') ? alphaNumber(n) : romanNumber(n);
  if (kind.endsWith('Uc')) s = s.toUpperCase();
  switch (m?.[2]) {
    case 'ParenBoth':
      return `(${s})`;
    case 'ParenR':
      return `${s})`;
    case 'Plain':
      return s;
    case 'Minus':
      return `- ${s} -`;
    default:
      return `${s}.`;
  }
}

/** The first bullet element of one property group found in the paragraph's style chain. */
function bulletProp(chain: string[], re: RegExp): string {
  for (const style of chain) {
    const m = re.exec(style);
    if (m) return m[0];
  }
  return '';
}

/**
 * The bullet of a paragraph (resolved through the paragraph, its shape, the layout and master placeholders and the
 * master text styles, each property on its own), as the CSS and text of its marker; numbering advances `state`.
 */
function paragraphBullet(chain: string[], lvl: number, state: ListState): { css: string; text: string; color: string } | undefined {
  const type = bulletProp(chain, /<a:bu(?:None|Char|AutoNum|Blip)\b[^>]*>/);
  state.length = Math.min(state.length, lvl + 1);
  if (!type || type.startsWith('<a:buNone')) {
    state[lvl] = undefined;
    return undefined;
  }
  const a = attrsOf(type);
  let text = '\u2022';
  const css: string[] = [];
  if (type.startsWith('<a:buAutoNum')) {
    const scheme = a.type ?? 'arabicPeriod';
    const start = Math.max(1, Number(a.startAt) || 1);
    const key = `${scheme}:${start}`;
    const prev = state[lvl];
    const n = prev && prev.key === key ? prev.n + 1 : start;
    state[lvl] = { key, n };
    text = autoNumber(scheme, n);
  } else {
    state[lvl] = undefined;
    if (type.startsWith('<a:buChar')) text = a.char || '\u2022';
  }
  const font = bulletProp(chain, /<a:buFont(?:Tx)?\b[^>]*>/);
  const face = font.startsWith('<a:buFontTx') ? '' : (attrsOf(font).typeface ?? '');
  if (face && SYMBOL_FONTS.test(face) && type.startsWith('<a:buChar')) {
    const code = text.codePointAt(0) ?? 0;
    text = SYMBOL_BULLETS[code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code] ?? '\u2022';
  } else if (face && !face.startsWith('+') && !SYMBOL_FONTS.test(face)) {
    css.push(`font-family:'${face.replace(/['"<>&;\\]/g, '')}'`);
  }
  const size = bulletProp(chain, /<a:buSz(?:Tx|Pct|Pts)\b[^>]*>/);
  const sizeVal = attrsOf(size).val ?? '';
  if (size.startsWith('<a:buSzPct')) {
    const pct = sizeVal.endsWith('%') ? parseFloat(sizeVal) : Number(sizeVal) / 1000;
    if (pct > 0 && pct !== 100) css.push(`font-size:${Math.round(pct * 10) / 10}%`);
  } else if (size.startsWith('<a:buSzPts')) {
    const pts = Number(sizeVal) / 100;
    if (pts > 0) css.push(`font-size:${pts}pt`);
  }
  const clr = bulletProp(chain, /<a:buClrTx\b[^>]*>|<a:buClr\b[^>]*>[\s\S]*?<\/a:buClr>/);
  const color = clr.startsWith('<a:buClr>') || clr.startsWith('<a:buClr ') ? clr.replace(/^<a:buClr\b[^>]*>|<\/a:buClr>$/g, '') : '';
  return { css: css.join(';'), text, color };
}

/** A run property element with the given fill (other fills of the run, not of its text outline, removed). */
function withFill(rPr: string, fill: string): string {
  const base = (rPr || '<a:rPr/>').replace(
    /<a:ln\b[^>]*?(?:\/>|>[\s\S]*?<\/a:ln>)|<a:(solidFill|gradFill|pattFill)\b[^>]*>[\s\S]*?<\/a:\1>|<a:(?:noFill|grpFill)\s*\/>/g,
    (m: string) => (m.startsWith('<a:ln') ? m : ''),
  );
  return base.endsWith('/>') ? `${base.slice(0, -2)}>${fill}</a:rPr>` : base.replace(/^<a:rPr\b[^>]*>/, (open: string) => open + fill);
}

/** CSS of a DrawingML underline other than single (pptxtojson draws `sng` itself), '' for none. */
function underlineCss(u: string | undefined): string {
  if (!u || u === 'none' || u === 'sng') return '';
  const style = u === 'dbl' ? 'double' : /^wavy/.test(u) ? 'wavy' : /^dotted/.test(u) ? 'dotted' : /^(?:dash|dotDash|dotDotDash)/.test(u) ? 'dashed' : 'solid';
  return `text-decoration-line:underline;text-decoration-style:${style}${/heavy$/i.test(u) ? ';text-decoration-thickness:0.1em' : ''}`;
}

const RPR_RE = /<a:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>)/;

/** One a:r / a:fld / a:br of a paragraph as an a:r pptxtojson keeps in order, with its markers. */
function patchRun(markup: PartMarkup, tag: string | undefined, inner: string | undefined, brInner: string | undefined): string {
  if (!tag) return `<a:r>${RPR_RE.exec(brInner ?? '')?.[0] ?? ''}<a:t>${MARK.br}</a:t></a:r>`;
  const content = inner ?? '';
  let rPr = RPR_RE.exec(content)?.[0] ?? '';
  const t = /<a:t>([\s\S]*?)<\/a:t>|<a:t\s*\/>/.exec(content);
  const css: string[] = [];
  const highlight = /<a:highlight>\s*<a:(srgbClr|schemeClr)\b[^>]*?\bval="(\w+)"/.exec(rPr);
  const highlightColor = highlight ? (highlight[1] === 'srgbClr' ? highlight[2] : schemeColor(highlight[2], markup)) : undefined;
  if (highlightColor && /^[0-9A-Fa-f]{6}$/.test(highlightColor)) css.push(`background-color:#${highlightColor}`);
  const underline = underlineCss(firstTagAttrs(rPr, 'a:rPr').u);
  if (underline) css.push(underline);
  // Hyperlinks take the theme hyperlink colour, unless PowerPoint's "use the text colour" extension is set.
  if (/<a:hlinkClick\b/.test(rPr) && !/<\w+:hlinkClr\b[^>]*\bval="tx"/.test(rPr)) rPr = withFill(rPr, '<a:solidFill><a:schemeClr val="hlink"/></a:solidFill>');
  if (!t) return `<a:r>${rPr}</a:r>`;
  let text = (t[1] ?? '').replace(MARK_RE, '').replace(/\t|&#0*9;|&#x0*9;/gi, MARK.tab);
  if (css.length && text) text = `${MARK.open}${css.join(';')}${MARK.sep}${text}${MARK.close}`;
  return `<a:r>${rPr}<a:t>${text}</a:t></a:r>`;
}

const ZERO_WIDTH_SPACE = String.fromCharCode(0x200b);
const RUN_RE = /<a:(r|fld)\b[^>]*>([\s\S]*?)<\/a:\1>|<a:br\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:br>)/g;
const BULLET_TYPE_RE = /<a:buChar\b[^>]*\/>|<a:buAutoNum\b[^>]*\/>|<a:buBlip\b[^>]*?(?:\/>|>[\s\S]*?<\/a:buBlip>)/g;

/** One paragraph: runs patched, its bullet (own or inherited) drawn as a marker run instead of a list. */
function patchParagraph(p: string, markup: PartMarkup, chain: (lvl: number) => string[], state: ListState): string {
  const m = /^(<a:p\b[^>]*>)([\s\S]*)<\/a:p>$/.exec(p);
  if (!m) return p;
  const pPrMatch = /^\s*<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/.exec(m[2]);
  const pPr = pPrMatch ? pPrMatch[0].trim() : '';
  const lvl = Math.min(8, Math.max(0, Number(firstTagAttrs(pPr, 'a:pPr').lvl) || 0));
  const rest = pPrMatch ? m[2].slice(pPrMatch[0].length) : m[2];
  let runs = rest.replace(RUN_RE, (_all: string, tag?: string, inner?: string, brInner?: string) => patchRun(markup, tag, inner, brInner));
  // An empty paragraph is a line of its end-of-paragraph size (pptxtojson would give it the inherited size).
  const end = /<a:endParaRPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:endParaRPr>)/.exec(runs);
  if (end && state.sizeEmpty && !/<a:r>/.test(runs)) runs = `<a:r>${end[0].replace(/^<a:endParaRPr\b/, '<a:rPr').replace(/<\/a:endParaRPr>$/, '</a:rPr>')}<a:t>${ZERO_WIDTH_SPACE}</a:t></a:r>${runs}`;
  // Empty paragraphs get no bullet and do not count in a numbered list.
  const bullet = /<a:t>[^<]/.test(rest) ? paragraphBullet([pPr, ...chain(lvl)], lvl, state) : undefined;
  let marker = '';
  if (bullet) {
    const first = /<a:(?:r|fld)\b[^>]*>\s*(<a:rPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:rPr>))?/.exec(rest)?.[1] ?? '<a:rPr/>';
    // The marker takes the size, font and colour of the first run (its text colour, not the hyperlink colour);
    // never its underline, highlight or link.
    let rPr = first
      .replace(/\s(?:u|strike|baseline)="[^"]*"/g, '')
      .replace(/<a:(highlight|hlinkClick|hlinkMouseOver)\b[^>]*?(?:\/>|>[\s\S]*?<\/a:\1>)/g, '');
    if (bullet.color) rPr = withFill(rPr, `<a:solidFill>${bullet.color}</a:solidFill>`);
    marker = `<a:r>${rPr}<a:t>${MARK.bullet}${bullet.css}${MARK.sep}${escapeXml(bullet.text)}${MARK.bulletEnd}</a:t></a:r>`;
  }
  return `${m[1]}${pPr.replace(BULLET_TYPE_RE, '')}${marker}${runs}</a:p>`;
}

const PARAGRAPH_RE = /<a:p\b[^>]*?\/>|<a:p\b[^>]*>[\s\S]*?<\/a:p>/g;

/** A text body (p:txBody, a:txBody of a table cell, dsp:txBody): every paragraph, and the vertical text marker. */
function patchTextBody(body: string, markup: PartMarkup, chain: (lvl: number) => string[], cell = false): string {
  const state: ListState = [];
  // Empty lines only matter between lines of text (a text body without text stays without), and in table cells.
  state.sizeEmpty = cell || /<a:t>[^<]/.test(body);
  let out = body.replace(PARAGRAPH_RE, (p: string) => patchParagraph(p, markup, chain, state));
  const vert = firstTagAttrs(body, 'a:bodyPr').vert;
  const mark = vert === 'vert' ? MARK.vert : vert === 'vert270' ? MARK.vert270 : '';
  if (mark) out = out.replace('<a:t>', `<a:t>${mark}`);
  return out;
}

const TXSTYLE_OF: Record<string, 'titleStyle' | 'bodyStyle' | 'otherStyle'> = {
  title: 'titleStyle',
  ctrTitle: 'titleStyle',
  dt: 'otherStyle',
  ftr: 'otherStyle',
  sldNum: 'otherStyle',
  hdr: 'otherStyle',
};
const MASTER_TYPE: Record<string, string> = { ctrTitle: 'title', subTitle: 'body', obj: 'body' };

/** The style chain of a shape's paragraphs at level `lvl` (its own pPr comes first, added by the caller). */
function shapeChain(shape: string, markup: PartMarkup, lstStyle: string): (lvl: number) => string[] {
  const ph = /<p:ph\b[^>]*>/.exec(shape);
  const layoutPhs = markup.layout;
  if (!layoutPhs) return (lvl) => [levelStyle(lstStyle, lvl)];
  if (!ph) return (lvl) => [levelStyle(lstStyle, lvl), levelStyle(markup.defaultStyle, lvl)];
  const a = attrsOf(ph[0]);
  const type = a.type ?? 'obj';
  const find = (list: PlaceholderStyle[], t: string, idx: string): PlaceholderStyle | undefined =>
    (idx ? list.find((p) => p.idx === idx) : undefined) ?? list.find((p) => p.type === t);
  const layout = find(layoutPhs, type, a.idx ?? '');
  const master = find(markup.master.placeholders, MASTER_TYPE[type] ?? type, '');
  const txStyle = markup.master[TXSTYLE_OF[type] ?? 'bodyStyle'];
  return (lvl) => [levelStyle(lstStyle, lvl), levelStyle(layout?.lstStyle ?? '', lvl), levelStyle(master?.lstStyle ?? '', lvl), levelStyle(txStyle, lvl)];
}

const SHAPE_OR_BODY_RE = /<p:sp\b[^>]*>[\s\S]*?<\/p:sp>|<a:tc\b[^>]*?(?<!\/)>[\s\S]*?<\/a:tc>|<(a|dsp):txBody\b[^>]*>[\s\S]*?<\/\1:txBody>/g;

/** Default table cell margins (EMU): 0.1 in left / right, 0.05 in top / bottom. */
const CELL_MARGINS: Array<[string, number]> = [
  ['marT', 45720],
  ['marR', 91440],
  ['marB', 45720],
  ['marL', 91440],
];

/** A table cell: its text, and its margins (when set) as a marker pptxtojson keeps (fixMarkup -> cell.margin). */
function patchCell(cell: string, markup: PartMarkup): string {
  const out = cell.replace(/<a:txBody\b[^>]*>[\s\S]*?<\/a:txBody>/, (body: string) =>
    patchTextBody(body, markup, (lvl) => [levelStyle(LST_STYLE_RE.exec(body)?.[1] ?? '', lvl)], true),
  );
  const tcPr = firstTagAttrs(out.slice(out.lastIndexOf('<a:tcPr')), 'a:tcPr');
  if (!CELL_MARGINS.some(([name]) => tcPr[name] !== undefined)) return out;
  const margins = CELL_MARGINS.map(([name, def]) => {
    const emu = Number(tcPr[name] ?? def);
    return Math.round(((Number.isFinite(emu) ? emu : def) / EMU_PER_PT) * 100) / 100;
  });
  return out.replace('<a:t>', `<a:t>${MARK.cell}${margins.join(',')}${MARK.sep}`);
}
const LST_STYLE_RE = /<a:lstStyle>([\s\S]*?)<\/a:lstStyle>/;

/** A slide / layout / master / diagram drawing / table styles part patched for pptxtojson (see the section comment). */
function patchPart(xml: string, markup: PartMarkup): string {
  const out = xml.replace(SHAPE_OR_BODY_RE, (block: string) => {
    if (block.startsWith('<a:tc')) return patchCell(block, markup);
    if (!block.startsWith('<p:sp')) return patchTextBody(block, markup, (lvl) => [levelStyle(LST_STYLE_RE.exec(block)?.[1] ?? '', lvl)]);
    return block.replace(/<p:txBody\b[^>]*>[\s\S]*?<\/p:txBody>/, (body: string) =>
      patchTextBody(body, markup, shapeChain(block, markup, LST_STYLE_RE.exec(body)?.[1] ?? '')),
    );
  });
  return fixTints(fixOutlines(out, markup), markup);
}

/**
 * Patches the parts pptxtojson reads (as they are in `zip`, after repairPackage) and returns the package to parse:
 * `data` when nothing changed.
 */
async function patchMarkup(zip: JSZip, data: ArrayBuffer): Promise<ArrayBuffer> {
  const read = async (path: string): Promise<string> => {
    const file = zip.file(path);
    return file ? file.async('string').catch(() => '') : '';
  };
  const relTarget = async (part: string, type: string): Promise<string | undefined> => {
    const slash = part.lastIndexOf('/');
    const rels = await read(`${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`);
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const a = attrsOf(m[0]);
      if ((a.Type ?? '').endsWith(`/${type}`) && a.TargetMode !== 'External' && a.Target) return resolvePart(part, a.Target);
    }
    return undefined;
  };
  const masters = new Map<string, Promise<MasterInfo>>();
  const masterInfo = (path: string): Promise<MasterInfo> => {
    let info = masters.get(path);
    if (!info) {
      info = (async (): Promise<MasterInfo> => {
        const xml = await read(path);
        const themePath = await relTarget(path, 'theme');
        const txStyle = (name: string): string => new RegExp(`<p:${name}>([\\s\\S]*?)</p:${name}>`).exec(xml)?.[1] ?? '';
        return {
          theme: readTheme(themePath ? await read(themePath) : ''),
          clrMap: firstTagAttrs(xml, 'p:clrMap'),
          placeholders: placeholderStyles(xml),
          titleStyle: txStyle('titleStyle'),
          bodyStyle: txStyle('bodyStyle'),
          otherStyle: txStyle('otherStyle'),
        };
      })();
      masters.set(path, info);
    }
    return info;
  };
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const firstMaster = names.filter((n) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(n)).sort()[0];
  const fallback: MasterInfo = firstMaster
    ? await masterInfo(firstMaster)
    : { theme: readTheme(''), clrMap: {}, placeholders: [], titleStyle: '', bodyStyle: '', otherStyle: '' };
  const defaultStyle = /<p:defaultTextStyle>([\s\S]*?)<\/p:defaultTextStyle>/.exec(await read('ppt/presentation.xml'))?.[1] ?? '';
  let changed = false;
  for (const name of names) {
    let markup: PartMarkup;
    if (/^ppt\/slides\/slide\d+\.xml$/i.test(name)) {
      const layoutPath = await relTarget(name, 'slideLayout');
      const masterPath = layoutPath ? await relTarget(layoutPath, 'slideMaster') : undefined;
      const layout = layoutPath ? placeholderStyles(await read(layoutPath)) : [];
      markup = { master: masterPath ? await masterInfo(masterPath) : fallback, layout, defaultStyle };
    } else if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)) {
      const masterPath = await relTarget(name, 'slideMaster');
      markup = { master: masterPath ? await masterInfo(masterPath) : fallback, defaultStyle };
    } else if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)) {
      markup = { master: await masterInfo(name), defaultStyle };
    } else if (/^ppt\/(?:diagrams\/drawing\d*|tableStyles)\.xml$/i.test(name)) {
      markup = { master: fallback, defaultStyle };
    } else {
      continue;
    }
    const xml = await read(name);
    const patched = xml ? patchPart(xml, markup) : xml;
    if (patched !== xml) {
      zip.file(name, patched);
      changed = true;
    }
  }
  return changed ? toArrayBuffer(await zip.generateAsync({ type: 'uint8array' })) : data;
}

/** Markers in pptxtojson's HTML -> HTML; sets the vertical text flags of `el`. */
function fixHtml(html: string, el?: Record<string, unknown>): string {
  if (!/[\uE000-\uE009]/.test(html)) return html;
  if (el && html.includes(MARK.vert270)) {
    el.isVertical = true;
    el.vert270 = true;
  } else if (el && html.includes(MARK.vert)) {
    el.isVertical = true;
  }
  const css = (s: string): string => s.replace(/&nbsp;/g, ' ').replace(/["<>]/g, '');
  return html
    .replace(/[\uE006\uE007]/g, '')
    .replace(/\uE008([^\uE003]*)\uE003([^\uE009]*)\uE009/g, (_m: string, style: string, text: string) => `<span class="pptx-bullet" style="${css(style)}">${text}</span>`)
    .replace(/\uE002([^\uE003]*)\uE003/g, (_m: string, style: string) => `<span style="${css(style)}">`)
    .replace(/\uE004/g, '</span>')
    .replace(/\uE000/g, '<br>')
    .replace(/\uE001/g, '\t')
    .replace(MARK_RE, '');
}

/** fixHtml over every text of a parsed deck (shapes, text boxes, table cells; groups and diagrams recursively). */
function fixMarkup(parsed: ParsedDeck): ParsedDeck {
  const walk = (elements: unknown): void => {
    if (!Array.isArray(elements)) return;
    for (const el of elements as Array<Record<string, unknown>>) {
      if (!el || typeof el !== 'object') continue;
      if (typeof el.content === 'string') el.content = fixHtml(el.content, el);
      if (el.type === 'table' && Array.isArray(el.data)) {
        for (const row of el.data as unknown[]) {
          if (!Array.isArray(row)) continue;
          for (const cell of row as Array<Record<string, unknown>>) {
            if (!cell || typeof cell.text !== 'string') continue;
            let text = cell.text;
            const margin = new RegExp(`${MARK.cell}([\\d.,]*)${MARK.sep}`).exec(text);
            if (margin) {
              const [t, r, b, l] = margin[1].split(',').map(Number);
              if ([t, r, b, l].every((n) => Number.isFinite(n))) cell.margin = { t, r, b, l };
              text = text.replace(margin[0], '');
            }
            cell.text = fixHtml(text);
          }
        }
      }
      walk(el.elements);
    }
  };
  for (const slide of parsed.slides) {
    walk(slide.elements);
    walk(slide.layoutElements);
  }
  return parsed;
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
  let missing: Set<string>;
  try {
    // The outline reads the parts as they are in the file; the repair then patches the zip for pptxtojson.
    outline = await new OutlineReader(pkg).read();
    const repaired = await repairPackage(zip, pkg, data);
    missing = repaired.missing;
    parsed = fixMarkup(await parseDeck(await patchMarkup(zip, repaired.data), PARSE_OPTIONS));
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }
  try {
    return buildDocument(outline, parsed, missing);
  } catch (err) {
    throw new Error(describeFailure(err), { cause: err });
  }
}
