// Markdown -> HTML for the webview preview (markdown-it + plugins, highlight.js, KaTeX).
//
// Pure module (no `vscode` import) so it can be unit-tested in plain Node.
// One markdown-it instance is created lazily and reused for every render; all
// per-render state (slugs, TOC, resource resolver, math closer index / scan) lives in
// the markdown-it `env` object, so nothing leaks between renders.

// ===== IMPORTS =====

import MarkdownIt from 'markdown-it';
import type {
  Delimiter,
  Env,
  MarkdownIt as MarkdownItInstance,
  RendererRule,
  StateBlock,
  StateCore,
  StateInline,
  Token,
} from 'markdown-it';
import hljs from 'highlight.js/lib/core';
import langBash from 'highlight.js/lib/languages/bash';
import langC from 'highlight.js/lib/languages/c';
import langCpp from 'highlight.js/lib/languages/cpp';
import langCsharp from 'highlight.js/lib/languages/csharp';
import langCss from 'highlight.js/lib/languages/css';
import langDart from 'highlight.js/lib/languages/dart';
import langDiff from 'highlight.js/lib/languages/diff';
import langDockerfile from 'highlight.js/lib/languages/dockerfile';
import langGo from 'highlight.js/lib/languages/go';
import langGraphql from 'highlight.js/lib/languages/graphql';
import langIni from 'highlight.js/lib/languages/ini';
import langJava from 'highlight.js/lib/languages/java';
import langJavascript from 'highlight.js/lib/languages/javascript';
import langJson from 'highlight.js/lib/languages/json';
import langKotlin from 'highlight.js/lib/languages/kotlin';
import langLatex from 'highlight.js/lib/languages/latex';
import langLess from 'highlight.js/lib/languages/less';
import langLua from 'highlight.js/lib/languages/lua';
import langMakefile from 'highlight.js/lib/languages/makefile';
import langMarkdown from 'highlight.js/lib/languages/markdown';
import langObjectivec from 'highlight.js/lib/languages/objectivec';
import langPerl from 'highlight.js/lib/languages/perl';
import langPhp from 'highlight.js/lib/languages/php';
import langPlaintext from 'highlight.js/lib/languages/plaintext';
import langPowershell from 'highlight.js/lib/languages/powershell';
import langProtobuf from 'highlight.js/lib/languages/protobuf';
import langPython from 'highlight.js/lib/languages/python';
import langR from 'highlight.js/lib/languages/r';
import langRuby from 'highlight.js/lib/languages/ruby';
import langRust from 'highlight.js/lib/languages/rust';
import langScala from 'highlight.js/lib/languages/scala';
import langScss from 'highlight.js/lib/languages/scss';
import langShell from 'highlight.js/lib/languages/shell';
import langSql from 'highlight.js/lib/languages/sql';
import langSwift from 'highlight.js/lib/languages/swift';
import langTypescript from 'highlight.js/lib/languages/typescript';
import langXml from 'highlight.js/lib/languages/xml';
import langYaml from 'highlight.js/lib/languages/yaml';
import katex from 'katex';

/** Shape of a markdown-it plugin function (the CJS plugins below ship no types). */
type MarkdownItPlugin = (md: MarkdownItInstance, options?: unknown) => void;

const footnotePlugin = require('markdown-it-footnote') as MarkdownItPlugin;
const emojiPlugins = require('markdown-it-emoji') as {
  bare: MarkdownItPlugin;
  light: MarkdownItPlugin;
  full: MarkdownItPlugin;
};

// ===== TYPES (shared with media/viewer.js — keep in sync) =====

export interface TocEntry {
  /** Heading level 1..6. */
  level: number;
  /** Plain heading text. */
  text: string;
  /** id attribute of the heading element (GitHub-style slug, de-duplicated). */
  slug: string;
  /** 0-based source line of the heading. */
  line: number;
}

export interface RenderOptions {
  /**
   * Rewrites a relative image/link path (as written in the markdown) to a URI
   * the webview can load. Return undefined to leave the href unchanged.
   */
  resolveResource?: (href: string) => string | undefined;
}

/**
 * A task checkbox the renderer produced (`<input class="task-list-item-checkbox" data-line="N">`) and where its
 * marker is in the source, so a toggle edits exactly that marker. Raw-HTML checkboxes are never tasks: the renderer
 * makes them inert (disabled).
 */
export interface TaskMarker {
  /** 0-based source line (the checkbox's data-line). */
  line: number;
  /** 0-based column (UTF-16) of the character between the brackets: ' ', 'x' or 'X'. */
  column: number;
  checked: boolean;
}

export interface RenderResult {
  html: string;
  toc: TocEntry[];
  /** Every task checkbox of `html`, in document order: the only lines a task toggle may edit. */
  tasks?: TaskMarker[];
}

// ===== IMPLEMENTATION =====

// ===== RENDER ENVIRONMENT (per-render state) =====

/** Per-render state, passed to markdown-it as `env` (never stored on the shared instance). */
interface RenderEnv extends Env {
  resolveResource?: (href: string) => string | undefined;
  /** github-slugger style occurrence counter. */
  slugs: Map<string, number>;
  toc: TocEntry[];
  tasks: TaskMarker[];
  /** Lazily computed: offset of the first character of every source line (task markers). */
  lineStarts?: number[];
  /** Lazily computed: sorted 0-based line numbers whose trimmed text ends with `$$`. */
  mathClosers?: number[];
  /** Last forward scan for the end of a `$$` block's paragraph (see mathParagraphEnd). */
  mathScan?: MathScan;
}

interface MathScan {
  /** Parser context the scan is valid for (token count, nesting, container bounds). */
  key: string;
  /** Line the scan started from. */
  from: number;
  /** Lines (from, okUntil] continue the paragraph. */
  okUntil: number;
  /** First line after `from` that ends it, or -1 when not reached yet. */
  stopAt: number;
}

function getEnv(env: unknown): RenderEnv {
  return env as RenderEnv;
}

// ===== UTILITIES =====

const HTML_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => HTML_ESCAPES[c] ?? c);
}

const NAMED_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

/** Decodes the few entities that realistically appear inside raw-HTML attribute values. */
function decodeAttributeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-f]{1,6}|#[0-9]{1,7}|[a-z]+);/gi, (whole, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    }
    return NAMED_ENTITIES[body.toLowerCase()] ?? whole;
  });
}

function isWhitespaceCode(code: number): boolean {
  return code === 0x20 || (code >= 0x09 && code <= 0x0d) || code === 0xa0 || code === 0x2028 || code === 0x2029;
}

/** True when the character at `pos` is preceded by an odd number of backslashes. */
function isEscaped(src: string, pos: number): boolean {
  let n = 0;
  for (let i = pos - 1; i >= 0 && src.charCodeAt(i) === 0x5c /* \ */; i--) n++;
  return n % 2 === 1;
}

function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message || err.name;
  return String(err);
}

/** Small LRU memo for pure, expensive conversions (KaTeX, highlight.js). */
class LruCache<V> {
  private readonly map = new Map<string, V>();
  private readonly max: number;

  constructor(max: number) {
    this.max = max;
  }

  get(key: string): V | undefined {
    const value = this.map.get(key);
    if (value !== undefined) {
      this.map.delete(key);
      this.map.set(key, value);
    }
    return value;
  }

  set(key: string, value: V): void {
    this.map.delete(key);
    this.map.set(key, value);
    if (this.map.size > this.max) {
      const oldest = this.map.keys().next();
      if (!oldest.done) this.map.delete(oldest.value);
    }
  }
}

/** Inserts ` data-source-line="N"` right after the tag name of the first matching opening tag(s). */
function injectSourceLine(html: string, line: string | number | null, tagPattern: string, all = false): string {
  if (line === null || line === undefined) return html;
  const re = new RegExp(`<(${tagPattern})(?=[\\s/>])`, all ? 'gi' : 'i');
  return html.replace(re, `<$1 data-source-line="${escapeHtml(String(line))}"`);
}

// ===== SYNTAX HIGHLIGHTING =====

/** Curated language set registered on highlight.js core (keeps the extension bundle lean). */
const HLJS_LANGUAGES: Record<string, Parameters<typeof hljs.registerLanguage>[1]> = {
  bash: langBash,
  c: langC,
  cpp: langCpp,
  csharp: langCsharp,
  css: langCss,
  dart: langDart,
  diff: langDiff,
  dockerfile: langDockerfile,
  go: langGo,
  graphql: langGraphql,
  ini: langIni,
  java: langJava,
  javascript: langJavascript,
  json: langJson,
  kotlin: langKotlin,
  latex: langLatex,
  less: langLess,
  lua: langLua,
  makefile: langMakefile,
  markdown: langMarkdown,
  objectivec: langObjectivec,
  perl: langPerl,
  php: langPhp,
  plaintext: langPlaintext,
  powershell: langPowershell,
  protobuf: langProtobuf,
  python: langPython,
  r: langR,
  ruby: langRuby,
  rust: langRust,
  scala: langScala,
  scss: langScss,
  shell: langShell,
  sql: langSql,
  swift: langSwift,
  typescript: langTypescript,
  xml: langXml,
  yaml: langYaml,
};

/** Extra fence names people commonly use that the language definitions do not alias themselves. */
const HLJS_EXTRA_ALIASES: Record<string, string[]> = {
  json: ['jsonc', 'json5', 'geojson', 'webmanifest'],
  xml: ['vue', 'svelte', 'xaml', 'csproj', 'xsd', 'xsl', 'xslt'],
  ini: ['env', 'dotenv', 'editorconfig', 'properties', 'cfg', 'conf'],
  bash: ['sh', 'zsh', 'shell-script', 'ksh'],
  protobuf: ['proto'],
  plaintext: ['text', 'txt', 'plain', 'none'],
};

let hljsReady = false;

function ensureHighlighter(): void {
  if (hljsReady) return;
  for (const [name, definition] of Object.entries(HLJS_LANGUAGES)) {
    if (!hljs.getLanguage(name)) hljs.registerLanguage(name, definition);
  }
  for (const [languageName, aliases] of Object.entries(HLJS_EXTRA_ALIASES)) {
    hljs.registerAliases(aliases, { languageName });
  }
  hljsReady = true;
}

/** Blocks larger than this are shown escaped but unhighlighted (bounded render time). */
const MAX_HIGHLIGHT_CHARS = 256 * 1024;
/** Only blocks below this size are memoised (bounds the cache to a few MB). */
const MAX_CACHED_HIGHLIGHT_CHARS = 16 * 1024;
const highlightCache = new LruCache<string>(200);

/** Returns highlighted HTML, or undefined when the language is unknown (caller escapes). */
function highlightCode(code: string, lang: string): string | undefined {
  if (!lang || code.length > MAX_HIGHLIGHT_CHARS || !hljs.getLanguage(lang)) return undefined;
  const key = `${lang}\u0000${code}`;
  const cached = highlightCache.get(key);
  if (cached !== undefined) return cached;
  try {
    const html = hljs.highlight(code, { language: lang, ignoreIllegals: true }).value;
    if (code.length <= MAX_CACHED_HIGHLIGHT_CHARS) highlightCache.set(key, html);
    return html;
  } catch {
    return undefined;
  }
}

// ===== MATH (KaTeX) =====

const mathCache = new LruCache<string>(500);
const MAX_CACHED_TEX_CHARS = 4 * 1024;

/** KaTeX -> HTML+MathML. Never throws: parse errors are rendered inline by KaTeX itself. */
function renderMath(tex: string, displayMode: boolean): string {
  const key = (displayMode ? 'D' : 'I') + tex;
  const cached = mathCache.get(key);
  if (cached !== undefined) return cached;
  let html: string;
  try {
    html = katex.renderToString(tex, {
      displayMode,
      throwOnError: false,
      output: 'htmlAndMathml',
      strict: 'ignore',
      trust: false,
    });
  } catch (err) {
    html = `<span class="katex-error" title="${escapeHtml(errorMessage(err))}">${escapeHtml(tex)}</span>`;
  }
  if (tex.length <= MAX_CACHED_TEX_CHARS) mathCache.set(key, html);
  return html;
}

/** Sorted line numbers whose trimmed text ends with `$$` — computed once per render, on first use. */
function getMathClosers(state: StateBlock): number[] {
  const env = getEnv(state.env);
  if (env.mathClosers) return env.mathClosers;
  const closers: number[] = [];
  const src = state.src;
  let line = 0;
  let start = 0;
  while (start <= src.length) {
    let end = src.indexOf('\n', start);
    if (end < 0) end = src.length;
    let e = end;
    while (e > start && isWhitespaceCode(src.charCodeAt(e - 1))) e--;
    if (e - start >= 2 && src.charCodeAt(e - 1) === 0x24 && src.charCodeAt(e - 2) === 0x24) closers.push(line);
    line++;
    start = end + 1;
  }
  env.mathClosers = closers;
  return closers;
}

/** First element of a sorted array that is >= value (index), or array length. */
function lowerBound(sorted: number[], value: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (sorted[mid] < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

type BlockRule = (state: StateBlock, startLine: number, endLine: number, silent: boolean) => boolean;

/** Paragraph terminator rules except block math itself (cached per markdown-it rule list). */
let paragraphTerminators: { source: BlockRule[]; rules: BlockRule[] } | undefined;

/**
 * True when `line` cannot belong to a `$$` block that started above it: a blank line, a
 * line outside the current container (same rule as fences), or a block (fence, heading, list
 * item, blockquote, html …) that interrupts paragraphs — markdown-it's paragraph rule checks.
 */
function endsMathParagraph(state: StateBlock, line: number, endLine: number): boolean {
  if (state.isEmpty(line)) return true;
  if (state.sCount[line] < state.blkIndent) return true;
  // Indented lines continue a paragraph whatever they contain.
  if (state.sCount[line] - state.blkIndent > 3) return false;
  const all = state.md.block.ruler.getRules('paragraph');
  if (paragraphTerminators?.source !== all) {
    paragraphTerminators = { source: all, rules: all.filter((rule) => rule !== mathBlockRule) };
  }
  // Rules check parentType (e.g. a list starting at 2 or an empty item cannot interrupt a paragraph).
  const oldParentType = state.parentType;
  state.parentType = 'paragraph';
  try {
    return paragraphTerminators.rules.some((rule) => rule(state, line, endLine, true));
  } finally {
    state.parentType = oldParentType;
  }
}

/**
 * First line in (startLine, limit] that ends the paragraph a `$$` opener at startLine starts,
 * or -1. The paragraph rule asks again for every following line (block math is a paragraph
 * terminator), so one forward scan is reused while the parser context is unchanged —
 * otherwise a run of `$$` openers would make rendering quadratic.
 */
function mathParagraphEnd(state: StateBlock, startLine: number, limit: number, endLine: number): number {
  const env = getEnv(state.env);
  const key = `${state.tokens.length}|${state.level}|${state.parentType}|${state.blkIndent}|${state.lineMax}|${endLine}`;
  let scan = env.mathScan;
  if (!scan || scan.key !== key || startLine < scan.from || startLine > scan.okUntil) {
    scan = env.mathScan = { key, from: startLine, okUntil: startLine, stopAt: -1 };
  }
  if (scan.stopAt >= 0) return scan.stopAt <= limit ? scan.stopAt : -1;
  for (let line = scan.okUntil + 1; line <= limit; line++) {
    if (endsMathParagraph(state, line, endLine)) {
      scan.stopAt = line;
      return line;
    }
    scan.okUntil = line;
  }
  return -1;
}

/**
 * Block math: `$$ … $$` on one line, or `$$` … `$$` spanning lines (text may follow the
 * opening / precede the closing delimiter). Like GitHub, a multi-line block must close inside
 * the paragraph it starts — no blank line or interrupting block (fence, list, heading …) in
 * between — so "$$ is the PID" prose cannot swallow the document up to a later `echo $$`.
 * An unterminated `$$` is NOT math (stays text).
 */
function mathBlockRule(state: StateBlock, startLine: number, endLine: number, silent: boolean): boolean {
  if (state.sCount[startLine] - state.blkIndent >= 4) return false;
  const src = state.src;
  const pos = state.bMarks[startLine] + state.tShift[startLine];
  const max = state.eMarks[startLine];
  if (pos + 2 > max || src.charCodeAt(pos) !== 0x24 || src.charCodeAt(pos + 1) !== 0x24) return false;

  const firstRest = src.slice(pos + 2, max);
  const firstTrim = firstRest.trim();
  let content: string;
  let closeLine: number;

  if (firstTrim.endsWith('$$')) {
    // Single line: $$ x^2 $$
    content = firstTrim.slice(0, -2);
    if (!content.trim()) return false;
    closeLine = startLine;
  } else if (firstTrim.includes('$$')) {
    // `$$a$$ trailing text` is inline display math inside a paragraph.
    return false;
  } else {
    const closers = getMathClosers(state);
    const idx = lowerBound(closers, startLine + 1);
    if (idx >= closers.length) return false;
    closeLine = closers[idx];
    if (closeLine >= endLine) return false;
    // Lines up to the closer must stay inside the container and the paragraph the opener starts.
    if (mathParagraphEnd(state, startLine, closeLine, endLine) >= 0) return false;
    const closeText = src.slice(state.bMarks[closeLine] + state.tShift[closeLine], state.eMarks[closeLine]).trim();
    const middle = closeLine > startLine + 1 ? state.getLines(startLine + 1, closeLine, state.sCount[startLine], true) : '';
    content = (firstTrim ? firstTrim + '\n' : '') + middle + closeText.slice(0, -2);
    if (!content.trim()) return false;
  }

  if (silent) return true;
  state.line = closeLine + 1;
  const token = state.push('math_block', 'div', 0);
  token.block = true;
  token.content = content.trim();
  token.markup = '$$';
  token.map = [startLine, state.line];
  return true;
}

/**
 * End (exclusive) of the code span whose opening backtick run starts at `pos`, or -1 when
 * the run has no closing run of the same length before `max` (CommonMark code spans).
 */
function codeSpanEnd(src: string, pos: number, max: number): number {
  let runEnd = pos;
  while (runEnd < max && src.charCodeAt(runEnd) === 0x60 /* ` */) runEnd++;
  const length = runEnd - pos;
  let search = runEnd;
  while (search < max) {
    const open = src.indexOf('`', search);
    if (open < 0 || open >= max) return -1;
    let close = open;
    while (close < max && src.charCodeAt(close) === 0x60) close++;
    if (close - open === length) return close;
    search = close;
  }
  return -1;
}

/**
 * Inline math (pandoc rules): `$…$` needs a non-space after the opening `$`, a non-space
 * before the closing `$`, and the closing `$` must not be followed by a digit — so
 * "$5 and $10" stays text. Code spans bind tighter (CommonMark precedence): a `$` inside
 * one never closes math ("costs $5. Use `$HOME`" is text + code). GitHub's `` $`\sqrt{3}`$ ``
 * form is math too. `$$…$$` inside a paragraph renders as display math.
 */
function mathInlineRule(state: StateInline, silent: boolean): boolean {
  const src = state.src;
  const start = state.pos;
  const max = state.posMax;
  if (src.charCodeAt(start) !== 0x24 /* $ */) return false;

  if (start + 1 < max && src.charCodeAt(start + 1) === 0x60 /* ` */) {
    // $`…`$ : the backticks protect the TeX from markdown; anything else falls through.
    const end = codeSpanEnd(src, start + 1, max);
    if (end > 0 && end < max && src.charCodeAt(end) === 0x24) {
      let ticks = 0;
      while (src.charCodeAt(start + 1 + ticks) === 0x60) ticks++;
      const tex = src.slice(start + 1 + ticks, end - ticks).trim();
      if (tex) {
        if (!silent) {
          const token = state.push('math_inline', 'span', 0);
          token.content = tex;
          token.markup = '$`';
        }
        state.pos = end + 1;
        return true;
      }
    }
  }

  if (start + 1 < max && src.charCodeAt(start + 1) === 0x24) {
    const from = start + 2;
    let end = src.indexOf('$$', from);
    while (end >= 0 && end + 2 <= max && isEscaped(src, end)) end = src.indexOf('$$', end + 1);
    if (end < 0 || end + 2 > max || !src.slice(from, end).trim()) {
      // Not math: consume both dollars as text so the second one cannot open `$…$`.
      if (!silent) state.pending += '$$';
      state.pos = start + 2;
      return true;
    }
    if (!silent) {
      const token = state.push('math_inline_display', 'span', 0);
      token.content = src.slice(from, end).trim();
      token.markup = '$$';
    }
    state.pos = end + 2;
    return true;
  }

  if (start + 1 >= max || isWhitespaceCode(src.charCodeAt(start + 1))) return false;
  let end = start + 1;
  for (;;) {
    // Next `$` or backtick: a code span is skipped whole (its dollars are code, not delimiters).
    while (end < max && src.charCodeAt(end) !== 0x24 && src.charCodeAt(end) !== 0x60) end++;
    if (end >= max) return false;
    if (src.charCodeAt(end) === 0x60) {
      const spanEnd = isEscaped(src, end) ? -1 : codeSpanEnd(src, end, max);
      if (spanEnd > 0) {
        end = spanEnd;
      } else {
        while (end < max && src.charCodeAt(end) === 0x60) end++;
      }
      continue;
    }
    if (!isEscaped(src, end)) break;
    end++;
  }
  if (isWhitespaceCode(src.charCodeAt(end - 1))) return false;
  const after = end + 1 < max ? src.charCodeAt(end + 1) : -1;
  if (after >= 0x30 && after <= 0x39) return false;

  if (!silent) {
    const token = state.push('math_inline', 'span', 0);
    token.content = src.slice(start + 1, end);
    token.markup = '$';
  }
  state.pos = end + 1;
  return true;
}

// ===== FRONT MATTER =====

type FrontMatterValue =
  | { kind: 'text'; text: string }
  | { kind: 'list'; items: string[] }
  | { kind: 'raw'; text: string };

interface FrontMatterEntry {
  key: string;
  value: FrontMatterValue;
}

/** `key: value` at column 0 (key may be quoted); value separated by whitespace or end of line. */
const FM_KEY_RE = /^("(?:[^"\\]|\\.)*"|'(?:[^']|'')*'|[^\s#"'\-:[\]{},&*!|>%@`?][^:]*?)[ \t]*:(?:[ \t]+(.*?))?[ \t]*$/;
const FM_ITEM_RE = /^([ \t]*)-(?:[ \t]+(.*?))?[ \t]*$/;

function isBlankOrComment(line: string): boolean {
  const t = line.trim();
  return t === '' || t[0] === '#';
}

function unquoteYaml(value: string): string {
  const s = value.trim();
  if (s.length >= 2 && s[0] === '"' && s[s.length - 1] === '"') {
    const escapes: Record<string, string> = { n: '\n', t: '\t', r: '\r', b: '\b', f: '\f', '0': '\0' };
    return s.slice(1, -1).replace(/\\(.)/g, (_m, c: string) => escapes[c] ?? c);
  }
  if (s.length >= 2 && s[0] === "'" && s[s.length - 1] === "'") return s.slice(1, -1).replace(/''/g, "'");
  return s;
}

/** Removes a trailing `# comment` (YAML: only when preceded by whitespace, never inside quotes). */
function stripYamlComment(value: string): string {
  const s = value.trim();
  const quoted = /^("(?:[^"\\]|\\.)*"|'(?:[^']|'')*')[ \t]*(?:#.*)?$/.exec(s);
  if (quoted) return quoted[1];
  if (s[0] === '#') return '';
  return s.replace(/[ \t]+#.*$/, '');
}

function dedentLines(lines: string[]): string {
  let indent = Infinity;
  for (const line of lines) {
    if (!line.trim()) continue;
    indent = Math.min(indent, /^[ \t]*/.exec(line)![0].length);
  }
  if (!Number.isFinite(indent)) indent = 0;
  return lines
    .map((l) => l.slice(Math.min(indent, /^[ \t]*/.exec(l)![0].length)))
    .join('\n')
    .replace(/\s+$/, '');
}

/** YAML folded scalar (`>`): single newlines become spaces, blank lines become newlines. */
function foldYaml(text: string): string {
  return text
    .split(/\n[ \t]*\n/)
    .map((para) => para.split('\n').map((l) => l.trim()).join(' '))
    .join('\n');
}

/** Splits a simple flow sequence `[a, "b", c]`; undefined for anything nested. */
function parseFlowSequence(value: string): string[] | undefined {
  const m = /^\[(.*)\]$/s.exec(value.trim());
  if (!m) return undefined;
  const inner = m[1];
  if (!inner.trim()) return [];
  const items: string[] = [];
  let current = '';
  let quote = '';
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (quote) {
      current += ch;
      if (ch === '\\' && quote === '"' && i + 1 < inner.length) current += inner[++i];
      else if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      current += ch;
    } else if (ch === '[' || ch === ']' || ch === '{' || ch === '}') {
      return undefined;
    } else if (ch === ',') {
      items.push(unquoteYaml(current));
      current = '';
    } else {
      current += ch;
    }
  }
  if (quote) return undefined;
  if (current.trim() || items.length === 0) items.push(unquoteYaml(current));
  return items;
}

/** `- a` / `- b` lines at one indent with scalar items; undefined for nested/complex content. */
function parseBlockSequence(block: string[]): string[] | undefined {
  let indent = -1;
  const items: string[] = [];
  for (const line of block) {
    if (isBlankOrComment(line)) continue;
    const m = FM_ITEM_RE.exec(line);
    if (!m) return undefined;
    if (indent < 0) indent = m[1].length;
    else if (m[1].length !== indent) return undefined;
    const item = stripYamlComment(m[2] ?? '');
    if (FM_KEY_RE.test(item) || item.startsWith('- ') || /^[|>{[]/.test(item)) return undefined;
    items.push(unquoteYaml(item));
  }
  return items;
}

function interpretYamlValue(inline: string, block: string[]): FrontMatterValue {
  if (/^[|>][-+0-9]*$/.test(inline)) {
    const text = dedentLines(block);
    return { kind: 'text', text: inline[0] === '|' ? text : foldYaml(text) };
  }
  if (inline) {
    const joined = [inline, ...block.map((l) => l.trim())].filter((s) => s).join(' ');
    const flow = parseFlowSequence(joined);
    if (flow) return { kind: 'list', items: flow };
    if (/^\{.*\}$/s.test(joined) || /^[&*!]/.test(joined)) return { kind: 'raw', text: joined };
    return { kind: 'text', text: unquoteYaml(joined) };
  }
  if (!block.some((l) => l.trim())) return { kind: 'text', text: '' };
  const items = parseBlockSequence(block);
  if (items) return { kind: 'list', items };
  return { kind: 'raw', text: dedentLines(block) };
}

/** Minimal YAML reader for flat front matter. Returns null when the block is not a simple mapping. */
function parseFrontMatter(body: string): FrontMatterEntry[] | null {
  const lines = body.split('\n');
  const entries: FrontMatterEntry[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (isBlankOrComment(line)) {
      i++;
      continue;
    }
    if (/^[ \t]/.test(line)) return null;
    const m = FM_KEY_RE.exec(line);
    if (!m) return null;
    const key = unquoteYaml(m[1]);
    const inline = stripYamlComment(m[2] ?? '');
    const block: string[] = [];
    i++;
    while (i < lines.length) {
      const next = lines[i];
      if (next.trim() === '' || /^[ \t]/.test(next) || /^-(?:[ \t]|$)/.test(next)) {
        block.push(next);
        i++;
      } else {
        break;
      }
    }
    while (block.length && !block[block.length - 1].trim()) block.pop();
    entries.push({ key, value: interpretYamlValue(inline, block) });
  }
  return entries;
}

/** Front matter must start the document with `---` and look like a YAML mapping. */
function frontMatterRule(state: StateBlock, startLine: number, endLine: number, silent: boolean): boolean {
  if (startLine !== 0 || state.parentType !== 'root' || state.level !== 0) return false;
  if (state.bMarks[0] !== 0 || state.tShift[0] !== 0) return false;
  const src = state.src;
  if (src.slice(0, state.eMarks[0]).trimEnd() !== '---') return false;

  let closeLine = -1;
  for (let line = 1; line < endLine; line++) {
    const text = src.slice(state.bMarks[line], state.eMarks[line]).trimEnd();
    if (text === '---' || text === '...') {
      closeLine = line;
      break;
    }
  }
  if (closeLine < 0) return false;

  const body = closeLine > 1 ? src.slice(state.bMarks[1], state.eMarks[closeLine - 1]) : '';
  const firstMeaningful = body.split('\n').find((l) => !isBlankOrComment(l));
  if (firstMeaningful !== undefined && !FM_KEY_RE.test(firstMeaningful)) return false;
  if (silent) return true;

  const token = state.push('front_matter', 'table', 0);
  token.block = true;
  token.content = body;
  token.markup = '---';
  token.map = [0, closeLine + 1];
  token.meta = { entries: parseFrontMatter(body) };
  token.attrSet('class', 'front-matter');
  state.line = closeLine + 1;
  return true;
}

function renderFrontMatterValue(value: FrontMatterValue): string {
  switch (value.kind) {
    case 'text':
      return escapeHtml(value.text).replace(/\n/g, '<br>\n');
    case 'list':
      return value.items.length ? `<ul>${value.items.map((it) => `<li>${escapeHtml(it)}</li>`).join('')}</ul>` : '';
    case 'raw':
      return `<pre><code>${escapeHtml(value.text)}</code></pre>`;
  }
}

const renderFrontMatter: RendererRule = (tokens, idx, _options, _env, self) => {
  const token = tokens[idx];
  const entries = (token.meta as { entries: FrontMatterEntry[] | null } | null)?.entries ?? null;
  let rows: string;
  if (entries === null) {
    if (!token.content.trim()) return '';
    rows = `<tr><td colspan="2"><pre><code>${escapeHtml(token.content)}</code></pre></td></tr>\n`;
  } else if (entries.length === 0) {
    return '';
  } else {
    rows = entries
      .map((e) => `<tr><th scope="row">${escapeHtml(e.key)}</th><td>${renderFrontMatterValue(e.value)}</td></tr>\n`)
      .join('');
  }
  return `<table${self.renderAttrs(token)}>\n<tbody>\n${rows}</tbody>\n</table>\n`;
};

// ===== GITHUB ALERTS =====

const ALERT_TITLES: Record<string, string> = {
  note: 'Note',
  tip: 'Tip',
  important: 'Important',
  warning: 'Warning',
  caution: 'Caution',
};
const ALERT_RE = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\][ \t]*(?:\n|$)/i;

/**
 * `> [!NOTE]` blockquotes -> `<div class="markdown-alert markdown-alert-note">` with a
 * `<p class="markdown-alert-title">Note</p>`. Runs after block parsing, before inline parsing,
 * so the marker can be removed from the paragraph source. Like GitHub, only top-level
 * blockquotes with some content after the marker line become alerts (not ones nested in
 * lists, blockquotes or footnotes, and not a lone `> [!WARNING]`).
 */
function alertsRule(state: StateCore): void {
  const tokens = state.tokens;
  for (let i = 0; i < tokens.length; i++) {
    const open = tokens[i];
    if (open.type !== 'blockquote_open' || open.level !== 0) continue;
    const pOpen = tokens[i + 1];
    const inline = tokens[i + 2];
    const pClose = tokens[i + 3];
    if (!pOpen || pOpen.type !== 'paragraph_open' || !inline || inline.type !== 'inline') continue;
    const m = ALERT_RE.exec(inline.content);
    if (!m) continue;

    let closeIdx = -1;
    for (let j = i + 1; j < tokens.length; j++) {
      if (tokens[j].type === 'blockquote_close' && tokens[j].level === open.level) {
        closeIdx = j;
        break;
      }
    }
    if (closeIdx < 0) continue;

    const rest = inline.content.slice(m[0].length);
    const markerOnly = !rest.trim() && pClose && pClose.type === 'paragraph_close';
    if (markerOnly && closeIdx === i + 4) continue; // nothing but the marker: GitHub keeps the blockquote

    const kind = m[1].toLowerCase();
    open.tag = 'div';
    open.attrJoin('class', `markdown-alert markdown-alert-${kind}`);
    open.meta = { ...(open.meta ?? {}), alert: kind };
    tokens[closeIdx].tag = 'div';

    if (markerOnly) {
      tokens.splice(i + 1, 3);
    } else {
      inline.content = rest;
      if (pOpen.map) pOpen.map = [Math.min(pOpen.map[0] + 1, pOpen.map[1]), pOpen.map[1]];
      if (inline.map) inline.map = [Math.min(inline.map[0] + 1, inline.map[1]), inline.map[1]];
    }

    const title = new state.Token('alert_title', 'p', 0);
    title.block = true;
    title.level = open.level + 1;
    title.meta = { kind };
    title.map = open.map ? [open.map[0], open.map[0] + 1] : null;
    title.attrSet('class', 'markdown-alert-title');
    tokens.splice(i + 1, 0, title);
  }
}

const renderAlertTitle: RendererRule = (tokens, idx, _options, _env, self) => {
  const token = tokens[idx];
  const kind = (token.meta as { kind: string }).kind;
  return `<p${self.renderAttrs(token)}>${escapeHtml(ALERT_TITLES[kind] ?? kind)}</p>\n`;
};

// ===== TASK LISTS =====

const TASK_RE = /^\[([ xX])\](?=\s|$)/;

/**
 * `- [ ] item` / `- [x] item` -> `<li class="task-list-item" data-line="N">` with a leading
 * `<input type="checkbox" class="task-list-item-checkbox" data-line="N" [checked] aria-label="item text">`.
 * N is the 0-based source line of the list item (where the `[ ]` marker is). Each checkbox is also recorded with
 * the exact column of its marker (RenderResult.tasks), whatever containers precede it on the line (`- - [ ]`,
 * `1. - [ ]`, `> - [ ]`, `[^1]: - [ ]`): the host edits that column instead of guessing with a pattern.
 */
function taskListPass(state: StateCore): void {
  const env = getEnv(state.env);
  const tokens = state.tokens;
  const listStack: Token[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === 'bullet_list_open' || t.type === 'ordered_list_open') {
      listStack.push(t);
      continue;
    }
    if (t.type === 'bullet_list_close' || t.type === 'ordered_list_close') {
      listStack.pop();
      continue;
    }
    if (t.type !== 'list_item_open' || !t.map) continue;

    const pOpen = tokens[i + 1];
    const inline = tokens[i + 2];
    if (!pOpen || pOpen.type !== 'paragraph_open' || !inline || inline.type !== 'inline') continue;
    if (!pOpen.map || pOpen.map[0] !== t.map[0]) continue;
    const m = TASK_RE.exec(inline.content);
    const first = inline.children?.[0];
    if (!m || !first || first.type !== 'text' || !first.content.startsWith(m[0])) continue;

    const line = t.map[0];
    const checked = m[1] !== ' ';
    // A marker that cannot be located in the source is still shown, but not recorded: toggling it edits nothing.
    const column = taskMarkerColumn(state.src, env, line, inline.content);
    if (column !== undefined) env.tasks.push({ line, column, checked });
    first.content = first.content.slice(m[0].length);
    const checkbox = new state.Token('task_checkbox', 'input', 0);
    // Accessible name: the item's text (screen readers otherwise announce an unlabeled checkbox).
    const label = inlineText(inline.children, true).replace(/\s+/g, ' ').trim();
    checkbox.meta = { checked, line, label };
    inline.children!.unshift(checkbox);

    t.attrJoin('class', 'task-list-item');
    t.attrSet('data-line', String(line));
    const parent = listStack[listStack.length - 1];
    if (parent && !String(parent.attrGet('class') ?? '').split(' ').includes('contains-task-list')) {
      parent.attrJoin('class', 'contains-task-list');
    }
  }
}

/**
 * Column of the character between the brackets of a task marker on source line `line`. The item's paragraph text
 * (`content`, raw source) starts with the marker, and its first line is the end of that source line (markdown-it
 * only strips the container markers and indentation before it, and trailing whitespace).
 */
function taskMarkerColumn(src: string, env: RenderEnv, line: number, content: string): number | undefined {
  const starts = (env.lineStarts ??= lineStartsOf(src));
  if (line >= starts.length) return undefined;
  const end = line + 1 < starts.length ? starts[line + 1] - 1 : src.length;
  const text = src.slice(starts[line], end).trimEnd();
  const first = content.split('\n', 1)[0].trimEnd();
  const at = text.length - first.length;
  if (at < 0 || !text.endsWith(first) || text[at] !== '[' || text[at + 2] !== ']') return undefined;
  return at + 1;
}

/** Offsets of the first character of every line of `src` (markdown-it normalized it to `\n` line breaks). */
function lineStartsOf(src: string): number[] {
  const starts = [0];
  for (let i = src.indexOf('\n'); i >= 0; i = src.indexOf('\n', i + 1)) starts.push(i + 1);
  return starts;
}

/** Longest accessible name taken from the item text (the rest is still read as the item's content). */
const MAX_TASK_LABEL = 200;

const renderTaskCheckbox: RendererRule = (tokens, idx) => {
  const meta = tokens[idx].meta as { checked: boolean; line: number; label: string };
  let label = meta.label || (meta.checked ? 'Completed task' : 'Incomplete task');
  if (label.length > MAX_TASK_LABEL) label = `${label.slice(0, MAX_TASK_LABEL - 1)}…`;
  return (
    `<input type="checkbox" class="task-list-item-checkbox" data-line="${meta.line}"${meta.checked ? ' checked' : ''}` +
    ` aria-label="${escapeHtml(label)}">`
  );
};

// ===== HEADING ANCHORS & TOC =====

/** GitHub slug: lowercase, drop everything except letters, marks, digits, `_`, `-` and spaces; spaces -> `-`. */
function githubSlug(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^\p{L}\p{M}\p{Nd}\p{Pc}\- ]/gu, '')
    .replace(/ /g, '-');
}

/** github-slugger de-duplication: foo, foo-1, foo-2 … */
function uniqueSlug(base: string, occurrences: Map<string, number>): string {
  let slug = base;
  while (occurrences.has(slug)) {
    const n = (occurrences.get(base) ?? 0) + 1;
    occurrences.set(base, n);
    slug = `${base}-${n}`;
  }
  occurrences.set(slug, 0);
  return slug;
}

/**
 * Plain text of inline tokens. `withAlt` (TOC labels, accessible names) includes image alt
 * text and shows emoji as characters; without it the text follows GitHub's slug source:
 * no image alt text, and a `:rocket:` shortcode contributes its name ("rocket-intro"), while
 * a literal 🚀 contributes nothing ("-intro").
 */
function inlineText(tokens: Token[] | null, withAlt: boolean): string {
  if (!tokens) return '';
  let out = '';
  for (const t of tokens) {
    switch (t.type) {
      case 'text':
      case 'code_inline':
      case 'math_inline':
      case 'math_inline_display':
        out += t.content;
        break;
      case 'emoji':
        out += withAlt ? t.content : t.markup;
        break;
      case 'softbreak':
      case 'hardbreak':
        out += ' ';
        break;
      case 'image':
        if (withAlt) out += inlineText(t.children, withAlt);
        break;
      case 'html_inline':
      case 'footnote_ref':
      case 'task_checkbox':
        break;
      default:
        if (t.children) out += inlineText(t.children, withAlt);
        break;
    }
  }
  return out;
}

function headingPass(state: StateCore): void {
  const env = getEnv(state.env);
  const tokens = state.tokens;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type !== 'heading_open') continue;
    const inline = tokens[i + 1];
    const children = inline && inline.type === 'inline' ? inline.children : null;
    // Not trimmed: like GitHub, "🚀 Features" -> "-features" (links written for GitHub keep working).
    const slug = uniqueSlug(githubSlug(inlineText(children, false)) || 'heading', env.slugs);
    t.attrSet('id', slug);
    env.toc.push({
      level: Number(t.tag.slice(1)) || 1,
      text: inlineText(children, true).replace(/\s+/g, ' ').trim(),
      slug,
      line: t.map ? t.map[0] : 0,
    });
  }
}

// ===== LINKS & RESOURCES =====

const SCHEME_RE = /^[a-z][a-z0-9+.-]*:/i;

/** Relative paths only: no scheme (http:, mailto:, data:, vscode-resource:, C:…), no `//host`, no `#anchor`. */
function isRelativeReference(href: string): boolean {
  if (!href) return false;
  const c = href[0];
  if (c === '#' || c === '?') return false;
  if (href.startsWith('//') || href.startsWith('\\\\')) return false;
  return !SCHEME_RE.test(href);
}

/** Runs the resolver on the path part (query/fragment are re-appended). Never throws. */
function resolveRelative(href: string, env: RenderEnv): string | undefined {
  if (!env.resolveResource || !isRelativeReference(href)) return undefined;
  const cut = href.search(/[?#]/);
  const path = cut >= 0 ? href.slice(0, cut) : href;
  const suffix = cut >= 0 ? href.slice(cut) : '';
  if (!path) return undefined;
  try {
    const resolved = env.resolveResource(path);
    return typeof resolved === 'string' && resolved ? resolved + suffix : undefined;
  } catch {
    return undefined;
  }
}

const RAW_TAG_RE = /<(img|a|source|video|audio|track)\b((?:[^>"']|"[^"]*"|'[^']*')*)>/gi;
const RAW_URL_ATTR_RE = /(\s)(src|href|poster|srcset)(\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))/gi;

/**
 * `srcset` (`<img>`, `<picture><source>`): "url [descriptor], url [descriptor]" — each URL runs to
 * the next whitespace (HTML parsing rules); relative ones go through the resolver. Returns
 * undefined when nothing changed.
 */
function rewriteSrcset(value: string, env: RenderEnv, md: MarkdownItInstance): string | undefined {
  let out = '';
  let changed = false;
  let pos = 0;
  const separators = /[\s,]*/y;
  while (pos < value.length) {
    separators.lastIndex = pos;
    const lead = separators.exec(value)![0];
    out += lead;
    pos += lead.length;
    if (pos >= value.length) break;
    let end = pos;
    while (end < value.length && !isWhitespaceCode(value.charCodeAt(end))) end++;
    let url = value.slice(pos, end);
    // A URL that ends in commas has no descriptors; the commas separate candidates.
    const commas = /,+$/.exec(url)?.[0] ?? '';
    if (commas) url = url.slice(0, -commas.length);
    const normalized = md.normalizeLink(url);
    const resolved = isRelativeReference(normalized) ? resolveRelative(normalized, env) : undefined;
    if (resolved) changed = true;
    out += (resolved ?? url) + commas;
    pos = end;
    if (commas) continue;
    // Descriptors run to the next comma outside parentheses.
    let depth = 0;
    let d = pos;
    for (; d < value.length; d++) {
      const c = value[d];
      if (c === '(') depth++;
      else if (c === ')') depth = Math.max(0, depth - 1);
      else if (c === ',' && depth === 0) break;
    }
    out += value.slice(pos, d);
    pos = d;
  }
  return changed ? out : undefined;
}

/**
 * An `<input` start tag as the HTML tokenizer sees it: the name ends at HTML whitespace, `/` or `>` (`<inputx` and
 * `<input ` are other elements).
 */
const RAW_INPUT_RE = /<input(?=[\t\n\f\r />]|$)/gi;

/**
 * Rewrites relative `src`/`href`/`poster`/`srcset` in raw HTML tags; `<a>` keeps the original in data-href. Raw
 * `<input>` elements become disabled: shown, but inert (GitHub drops them), so a checkbox written in HTML - even one
 * with the class and data-line of a task checkbox - can never toggle a task line; only the renderer's own task
 * checkboxes (RenderResult.tasks) can.
 */
function rewriteRawHtml(html: string, env: RenderEnv, md: MarkdownItInstance): string {
  if (html.indexOf('<') < 0) return html;
  html = html.replace(RAW_INPUT_RE, '<input disabled');
  return html.replace(RAW_TAG_RE, (whole, name: string, attrs: string) => {
    const isAnchor = name.toLowerCase() === 'a';
    let original: string | undefined;
    const newAttrs = attrs.replace(
      RAW_URL_ATTR_RE,
      (attrWhole, ws: string, attr: string, _eq: string, dq?: string, sq?: string, uq?: string) => {
        const lower = attr.toLowerCase();
        if (isAnchor ? lower !== 'href' : lower === 'href') return attrWhole;
        if (lower === 'srcset') {
          const srcset = rewriteSrcset(decodeAttributeEntities(dq ?? sq ?? uq ?? ''), env, md);
          return srcset ? `${ws}${attr}="${escapeHtml(srcset)}"` : attrWhole;
        }
        const value = md.normalizeLink(decodeAttributeEntities(dq ?? sq ?? uq ?? '').trim());
        if (!isRelativeReference(value)) return attrWhole;
        if (isAnchor) original = value;
        const resolved = resolveRelative(value, env);
        return resolved ? `${ws}${attr}="${escapeHtml(resolved)}"` : attrWhole;
      },
    );
    if (original === undefined && newAttrs === attrs) return whole;
    const dataHref = original !== undefined && !/\sdata-href\s*=/i.test(attrs) ? ` data-href="${escapeHtml(original)}"` : '';
    return `<${name}${dataHref}${newAttrs}>`;
  });
}

function rewriteInlineChildren(children: Token[], env: RenderEnv, md: MarkdownItInstance): void {
  for (const child of children) {
    if (child.type === 'link_open') {
      const href = child.attrGet('href');
      if (typeof href === 'string' && isRelativeReference(href)) {
        child.attrSet('data-href', href);
        const resolved = resolveRelative(href, env);
        if (resolved) child.attrSet('href', resolved);
      }
    } else if (child.type === 'image') {
      const src = child.attrGet('src');
      if (typeof src === 'string') {
        const resolved = resolveRelative(src, env);
        if (resolved) child.attrSet('src', resolved);
      }
    } else if (child.type === 'html_inline') {
      child.content = rewriteRawHtml(child.content, env, md);
    }
    if (child.children && child.children.length) rewriteInlineChildren(child.children, env, md);
  }
}

function resourcePass(state: StateCore): void {
  const env = getEnv(state.env);
  for (const t of state.tokens) {
    if (t.type === 'inline' && t.children) rewriteInlineChildren(t.children, env, state.md);
    else if (t.type === 'html_block') t.content = rewriteRawHtml(t.content, env, state.md);
  }
}

// ===== SOURCE LINES =====

/**
 * Every block-level opening/self-closing token with a source map gets data-source-line
 * (0-based start line). Footnote containers (built after parsing, without maps) take the
 * line of their first mapped child.
 */
function sourceLinePass(state: StateCore): void {
  const tokens = state.tokens;
  let lastLine = 0;
  for (const t of tokens) {
    if (!t.map || !t.block || t.nesting < 0 || t.type === 'inline') continue;
    t.attrSet('data-source-line', String(t.map[0]));
    if (t.map[0] > lastLine) lastLine = t.map[0];
  }

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type !== 'footnote_block_open' && t.type !== 'footnote_open') continue;
    const closeType = t.type === 'footnote_block_open' ? 'footnote_block_close' : 'footnote_close';
    let line: number | undefined;
    for (let j = i + 1; j < tokens.length && tokens[j].type !== closeType; j++) {
      const map = tokens[j].map;
      if (map) {
        line = map[0];
        break;
      }
    }
    t.attrSet('data-source-line', String(line ?? lastLine));
  }
}

/** Raw HTML blocks: put data-source-line on the block's first opening tag (if it starts with one). */
const renderHtmlBlock: RendererRule = (tokens, idx) => {
  const token = tokens[idx];
  const line = token.attrGet('data-source-line');
  const content = token.content;
  if (line === null) return content;
  const m = /^\s*<([A-Za-z][A-Za-z0-9-]*)((?:[^>"']|"[^"]*"|'[^']*')*)>/.exec(content);
  if (!m || /\sdata-source-line\s*=/i.test(m[2])) return content;
  return injectSourceLine(content, line, m[1].replace(/[^A-Za-z0-9-]/g, ''));
};

// ===== FENCES & CODE =====

/** Language id from a fence info string: "ts title=x" -> "ts", "{.python}" -> "python", "js{1,3}" -> "js". */
function fenceLanguage(md: MarkdownItInstance, info: string): string {
  if (!info) return '';
  const first = md.utils.unescapeAll(info).trim().split(/\s+/)[0] ?? '';
  return first.replace(/^\{\.?|[{}].*$/g, '').toLowerCase();
}

/** Fenced code: mermaid -> div.mermaid, math -> KaTeX block, otherwise highlight.js (unlabeled: plain). */
function createFenceRule(md: MarkdownItInstance): RendererRule {
  return (tokens, idx, _options, _env, self) => {
    const token = tokens[idx];
    const lang = fenceLanguage(md, token.info);
    const attrs = self.renderAttrs(token);

    if (lang === 'mermaid') return `<div class="mermaid"${attrs}>${escapeHtml(token.content)}</div>\n`;
    if (lang === 'math') return `<div class="math math-block"${attrs}>${renderMath(token.content.trim(), true)}</div>\n`;

    const highlighted = lang ? highlightCode(token.content, lang) : undefined;
    const body = highlighted ?? escapeHtml(token.content);
    const cls = lang ? `hljs language-${escapeHtml(lang)}` : 'hljs';
    return `<pre${attrs}><code class="${cls}">${body}</code></pre>\n`;
  };
}

const renderMathBlock: RendererRule = (tokens, idx, _options, _env, self) => {
  const token = tokens[idx];
  return `<div class="math math-block"${self.renderAttrs(token)}>${renderMath(token.content, true)}</div>\n`;
};

const renderMathInline: RendererRule = (tokens, idx) =>
  `<span class="math math-inline">${renderMath(tokens[idx].content, false)}</span>`;

const renderMathInlineDisplay: RendererRule = (tokens, idx) =>
  `<span class="math math-display">${renderMath(tokens[idx].content, true)}</span>`;

// ===== STRIKETHROUGH (single tilde) =====

/** Delimiter id of a single `~`: pairs only with another single `~` (`~a~~` stays text, like GFM). */
const SINGLE_TILDE = 0x1007e;

/**
 * GFM strikethrough accepts one or two tildes (`~text~`); markdown-it only knows `~~`. A run
 * of exactly one `~` becomes a delimiter (left/right-flanking rules as for `~~`).
 */
function singleTildeRule(state: StateInline, silent: boolean): boolean {
  if (silent || state.src.charCodeAt(state.pos) !== 0x7e /* ~ */) return false;
  const scanned = state.scanDelims(state.pos, true);
  if (scanned.length !== 1) return false;
  const token = state.push('text', '', 0);
  token.content = '~';
  state.delimiters.push({
    marker: SINGLE_TILDE,
    length: 0, // no emphasis "rule of 3"
    token: state.tokens.length - 1,
    end: -1,
    open: scanned.can_open,
    close: scanned.can_close,
  });
  state.pos += 1;
  return true;
}

function pairSingleTildes(state: StateInline, delimiters: Delimiter[]): void {
  for (const start of delimiters) {
    if (start.marker !== SINGLE_TILDE || start.end === -1) continue;
    const open = state.tokens[start.token];
    const close = state.tokens[delimiters[start.end].token];
    open.type = 's_open';
    open.tag = 's';
    open.nesting = 1;
    open.markup = '~';
    open.content = '';
    close.type = 's_close';
    close.tag = 's';
    close.nesting = -1;
    close.markup = '~';
    close.content = '';
  }
}

/** Runs after balance_pairs: matched single-tilde delimiters -> `<s>…</s>`. */
function singleTildePostProcess(state: StateInline): void {
  pairSingleTildes(state, state.delimiters);
  for (const meta of state.tokens_meta) {
    if (meta?.delimiters) pairSingleTildes(state, meta.delimiters);
  }
}

// ===== MARKDOWN-IT INSTANCE =====

/**
 * GFM autolinks: `www.example.com` links (without fuzzy bare-domain links like `README.md`),
 * and — like GitHub — no `ftp://…` or protocol-relative `//host/path` links (linkify-it
 * defaults; the host could not open either).
 */
function addWwwLinkify(md: MarkdownItInstance): void {
  let validator: RegExp | undefined;
  md.linkify.set({ fuzzyLink: false, fuzzyEmail: true });
  md.linkify.add('ftp:', null);
  md.linkify.add('//', null);
  md.linkify.add('www.', {
    validate(text, pos, self) {
      validator ??= new RegExp(self.re.get_url_host_port().source + self.re.get_path().source, 'iy');
      validator.lastIndex = pos;
      const m = validator.exec(text);
      return m ? m[0].length : 0;
    },
    normalize(match) {
      match.url = `http://${match.url}`;
    },
  });
}

/** Footnote reference label: like GitHub, the footnote's number for every reference ("1", "1"), not "[1]", "[1:1]". */
const renderFootnoteCaption: RendererRule = (tokens, idx) =>
  String(Number((tokens[idx].meta as { id: number }).id) + 1);

function wrapFootnoteRule(md: MarkdownItInstance, name: string, tags: string): void {
  const original = md.renderer.rules[name];
  if (!original) return;
  md.renderer.rules[name] = (tokens, idx, options, env, self) =>
    injectSourceLine(original(tokens, idx, options, env, self), tokens[idx].attrGet('data-source-line'), tags, true);
}

function createMarkdownIt(): MarkdownItInstance {
  ensureHighlighter();
  const md = new MarkdownIt('default', {
    html: true,
    linkify: true,
    typographer: false,
    breaks: false,
  });
  addWwwLinkify(md);

  md.use(footnotePlugin);
  md.use(emojiPlugins.full, { shortcuts: {} });

  // Block rules
  md.block.ruler.before('table', 'front_matter', frontMatterRule);
  md.block.ruler.after('blockquote', 'math_block', mathBlockRule, {
    alt: ['paragraph', 'reference', 'blockquote', 'list'],
  });

  // Inline rules
  md.inline.ruler.after('escape', 'math_inline', mathInlineRule);
  md.inline.ruler.after('strikethrough', 'strikethrough_single', singleTildeRule);
  md.inline.ruler2.after('strikethrough', 'strikethrough_single', singleTildePostProcess);

  // Core rules
  md.core.ruler.after('block', 'github_alerts', alertsRule);
  md.core.ruler.push('fv_task_lists', taskListPass);
  md.core.ruler.push('fv_headings', headingPass);
  md.core.ruler.push('fv_resources', resourcePass);
  md.core.ruler.push('fv_source_lines', sourceLinePass);

  // Renderer rules
  const rules = md.renderer.rules;
  rules.fence = createFenceRule(md);
  rules.html_block = renderHtmlBlock;
  rules.front_matter = renderFrontMatter;
  rules.alert_title = renderAlertTitle;
  rules.task_checkbox = renderTaskCheckbox;
  rules.math_block = renderMathBlock;
  rules.math_inline = renderMathInline;
  rules.math_inline_display = renderMathInlineDisplay;
  rules.footnote_caption = renderFootnoteCaption;
  wrapFootnoteRule(md, 'footnote_block_open', 'hr|section');
  wrapFootnoteRule(md, 'footnote_open', 'li');

  return md;
}

let sharedMarkdownIt: MarkdownItInstance | undefined;

function getMarkdownIt(): MarkdownItInstance {
  sharedMarkdownIt ??= createMarkdownIt();
  return sharedMarkdownIt;
}

// ===== PUBLIC API =====

/** Largest slice of the source echoed back in the error block. */
const MAX_ERROR_SOURCE = 200_000;

function renderErrorBlock(err: unknown, source: string): string {
  const text = source.length > MAX_ERROR_SOURCE ? `${source.slice(0, MAX_ERROR_SOURCE)}\n…` : source;
  return (
    `<div class="markdown-render-error" data-source-line="0" role="alert">\n` +
    `<p><strong>Markdown preview failed:</strong> ${escapeHtml(errorMessage(err))}</p>\n` +
    `<pre data-source-line="0"><code>${escapeHtml(text)}</code></pre>\n` +
    `</div>\n`
  );
}

/**
 * Renders markdown to HTML for the preview. Never throws: on an internal failure the
 * result is an error block showing the message and the raw source.
 */
export function renderMarkdown(source: string, opts?: RenderOptions): RenderResult {
  const text = typeof source === 'string' ? source : String(source ?? '');
  const env: RenderEnv = {
    resolveResource: opts?.resolveResource,
    slugs: new Map<string, number>(),
    toc: [],
    tasks: [],
  };
  try {
    const md = getMarkdownIt();
    // A leading BOM is not part of the text (VS Code's TextDocument has none either): task columns are counted without it.
    const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
    const tokens = md.parse(src, env);
    const html = md.renderer.render(tokens, md.options, env);
    return { html, toc: env.toc, tasks: env.tasks };
  } catch (err) {
    return { html: renderErrorBlock(err, text), toc: [], tasks: [] };
  }
}

/**
 * The headings of a document exactly as renderMarkdown's `toc` (GitHub slugs, 0-based source lines), parsed without
 * rendering: lets the host resolve a link such as `other.md#install-steps` to the line of that heading.
 */
export function markdownHeadings(source: string): TocEntry[] {
  const text = typeof source === 'string' ? source : String(source ?? '');
  const env: RenderEnv = { slugs: new Map<string, number>(), toc: [], tasks: [] };
  try {
    getMarkdownIt().parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text, env);
    return env.toc;
  } catch {
    return [];
  }
}
