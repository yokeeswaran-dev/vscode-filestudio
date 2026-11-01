// Regression tests for fixed 0.1.x bugs in the extension manifest (package.json) and the build / release scripts.
// Each test names the bug id from the 0.1.x bug hunts. Run with `npm test`.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const ROOT = path.join(__dirname, '..', '..');
const PACKAGE_TEXT = fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8');
const pkg = JSON.parse(PACKAGE_TEXT);
const { customEditors, menus, commands } = pkg.contributes;

const editor = (viewType) => customEditors.find((e) => e.viewType === viewType);
/** File extensions (without the dot) of a custom editor's selector. */
const extensionsOf = (viewType) => editor(viewType).selector.map((s) => /^\*\.(\w+)$/.exec(s.filenamePattern)[1]);
const ALL_EXTENSIONS = customEditors.flatMap((e) => extensionsOf(e.viewType));
/** The `resourceExtname =~ /.../i` regex of a when-clause. */
function extnameRegex(when) {
  const m = /resourceExtname =~ \/(.*?)\/(\w*)/.exec(when);
  assert.ok(m, `when-clause has a resourceExtname test: ${when}`);
  return new RegExp(m[1], m[2]);
}
const openWithMenus = () =>
  Object.entries(menus).flatMap(([menu, items]) => items.filter((i) => i.command === 'fileStudio.openWith').map((i) => ({ menu, when: i.when })));
/** `default` for the text editor and for diffs: a string priority applies to both. */
const priorityFor = (viewType, kind) => {
  const priority = editor(viewType).priority;
  return typeof priority === 'string' ? priority : priority[kind];
};

// ===== CUSTOM EDITORS =====

test('PRE-1: CSV / TSV / PSV / SSV open in FileStudio by default (not as text)', () => {
  assert.deepEqual(extensionsOf('fileStudio.csv'), ['csv', 'tsv', 'psv', 'ssv']);
  assert.equal(priorityFor('fileStudio.csv', 'textEditor'), 'default');
});

test('R3-004: xlsx / docx / pdf / pptx diffs (Open Changes) use FileStudio, not a text diff', () => {
  for (const viewType of ['fileStudio.sheet', 'fileStudio.docx', 'fileStudio.pdf', 'fileStudio.pptx']) {
    assert.equal(priorityFor(viewType, 'textEditor'), 'default', `${viewType}: editor`);
    assert.equal(priorityFor(viewType, 'diffEditor'), 'default', `${viewType}: diff editor`);
  }
});

test('H3-F3: .markdown files are in the Markdown selector and in every Open with FileStudio menu', () => {
  assert.deepEqual(extensionsOf('fileStudio.markdown'), ['md', 'markdown']);
  for (const { menu, when } of openWithMenus()) assert.ok(extnameRegex(when).test('.markdown'), `${menu}: ${when}`);
});

test('R3-033: pptm / ppsx / ppsm / potx / potm are in the presentation selector and the menus', () => {
  assert.deepEqual(extensionsOf('fileStudio.pptx'), ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm']);
  const menuEntries = openWithMenus();
  assert.ok(menuEntries.length >= 3, 'explorer, editor title and command palette entries');
  for (const { menu, when } of menuEntries) {
    const re = extnameRegex(when);
    for (const ext of ALL_EXTENSIONS) assert.ok(re.test(`.${ext}`) && re.test(`.${ext.toUpperCase()}`), `${menu}: .${ext}`);
    assert.ok(!re.test('.txt') && !re.test('.xlsm'), `${menu}: no other extensions`);
  }
});

// ===== MENUS AND COMMANDS =====

test('R3-064: Open with FileStudio is not offered on folders (a folder named x.csv)', () => {
  const explorer = menus['explorer/context'].filter((i) => i.command.startsWith('fileStudio.'));
  assert.ok(explorer.length > 0);
  for (const item of explorer) assert.match(item.when, /(^|&& )!explorerResourceIsFolder( &&|$)/, item.command);
});

test('H3-F4: every editor title-bar button (navigation group) has an icon', () => {
  const navigation = menus['editor/title'].filter((i) => /^navigation/.test(i.group ?? ''));
  assert.ok(navigation.length >= 2);
  for (const item of navigation) {
    const command = commands.find((c) => c.command === item.command);
    assert.ok(command, `${item.command} is declared`);
    assert.match(command.icon ?? '', /^\$\([\w-]+\)$/, `${item.command} has a codicon`);
  }
});

test('R2H4-8: the extension runs in Restricted Mode (untrusted workspaces supported)', () => {
  assert.equal(pkg.capabilities?.untrustedWorkspaces?.supported, true);
});

// ===== BUILD AND RELEASE =====

test('R3-020: package.json is in npm format (npm version must not reformat the whole file)', () => {
  const text = PACKAGE_TEXT.replace(/\r\n/g, '\n');
  assert.equal(text, `${JSON.stringify(pkg, null, 2)}\n`);
});

test('R3-068: only woff2 KaTeX fonts are built and packaged (no woff / ttf fallbacks)', (t) => {
  const ignore = fs.readFileSync(path.join(ROOT, '.vscodeignore'), 'utf8').split(/\r?\n/);
  assert.ok(ignore.includes('dist/fonts/*.woff') && ignore.includes('dist/fonts/*.ttf'), '.vscodeignore drops stale fallbacks');
  const fonts = path.join(ROOT, 'dist', 'fonts');
  if (!fs.existsSync(fonts)) {
    t.skip('dist/ not built (CI runs `npm run build` before `npm test`)');
    return;
  }
  const files = fs.readdirSync(fonts);
  assert.ok(files.some((f) => /^KaTeX_.*\.woff2$/.test(f)), 'KaTeX woff2 fonts are there');
  assert.deepEqual(files.filter((f) => /\.(woff|ttf)$/.test(f)), []);
});

/** Runs .github/scripts/changelog.js on `changelog` in a temp folder; returns { stdout, text } (text = file after). */
function runChangelog(changelog, ...args) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fs-changelog-'));
  try {
    fs.mkdirSync(path.join(dir, '.github', 'scripts'), { recursive: true });
    fs.copyFileSync(path.join(ROOT, '.github', 'scripts', 'changelog.js'), path.join(dir, '.github', 'scripts', 'changelog.js'));
    fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ repository: { url: 'https://github.com/o/r.git' } }));
    fs.writeFileSync(path.join(dir, 'CHANGELOG.md'), changelog);
    const out = [];
    for (const argv of args) out.push(execFileSync(process.execPath, [path.join(dir, '.github', 'scripts', 'changelog.js'), ...argv], { cwd: dir, encoding: 'utf8' }));
    return { stdout: out, text: fs.readFileSync(path.join(dir, 'CHANGELOG.md'), 'utf8') };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

const OLD_RELEASE = '## [1.0.0] - 2026-01-01\n\n### Added\n\n- First release.\n\n';
const LINKS = '[Unreleased]: https://github.com/o/r/compare/v1.0.0...HEAD\n[1.0.0]: https://github.com/o/r/releases/tag/v1.0.0\n';

test('R3-019: release notes keep a section link definition and drop empty headings', () => {
  const changelog =
    '# Changelog\n\n## [Unreleased]\n\n### Added\n\n- New thing, see [#12].\n\n[#12]: https://github.com/o/r/issues/12\n\n### Fixed\n\n#### Sheet\n\n' +
    OLD_RELEASE +
    LINKS;
  const { stdout, text } = runChangelog(changelog, ['release', '1.1.0', '2026-10-16'], ['notes', '1.1.0']);
  assert.equal(stdout[1], '### Added\n\n- New thing, see [#12].\n\n[#12]: https://github.com/o/r/issues/12\n');
  assert.match(text, /## \[Unreleased\]\n\n## \[1\.1\.0\] - 2026-10-16\n\n### Added\n/);
  assert.match(text, /^\[Unreleased\]: https:\/\/github\.com\/o\/r\/compare\/v1\.1\.0\.\.\.HEAD\n\[1\.1\.0\]: https:\/\/github\.com\/o\/r\/compare\/v1\.0\.0\.\.\.v1\.1\.0$/m);
});

test('R3-019: an empty [Unreleased] section gives "Maintenance release" notes and a warning', () => {
  const { stdout } = runChangelog(`# Changelog\n\n## [Unreleased]\n\n### Fixed\n\n${OLD_RELEASE}${LINKS}`, ['release', '1.0.1', '2026-10-16'], ['notes', '1.0.1']);
  assert.match(stdout[0], /^::warning::The \[Unreleased\] section of CHANGELOG\.md has no entries/m);
  assert.equal(stdout[1], '### Changed\n\n- Maintenance release.\n');
});
