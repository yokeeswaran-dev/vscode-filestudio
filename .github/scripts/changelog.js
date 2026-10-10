// CHANGELOG.md helper, used by `node .github/scripts/release.js prepare`. The file follows Keep a Changelog.
//
//   node .github/scripts/changelog.js release <version> <yyyy-mm-dd>
//       Moves the entries under "## [Unreleased]" into a new "## [<version>] - <date>" section and updates the
//       compare links at the bottom of the file.
//   node .github/scripts/changelog.js notes <version>
//       Prints the "## [<version>]" section (for example as GitHub release notes).
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..', '..');
const FILE = path.join(ROOT, 'CHANGELOG.md');
const REPO = require(path.join(ROOT, 'package.json')).repository.url.replace(/\.git$/, '');

const UNRELEASED_HEADER = '## [Unreleased]';
/** Fallback notes when a release is made with an empty [Unreleased] section. */
const EMPTY_NOTES = '### Changed\n\n- Maintenance release.';

function fail(message) {
  console.error(`::error::${message}`);
  process.exit(1);
}

const escapeRe = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Start of the next "## [" heading or of the version compare links ("[Unreleased]: ...", "[1.2.3]: ...") after
 * `from` (the end of a section). Other link definitions (for example "[#12]: ...") belong to the section.
 */
function sectionEnd(text, from) {
  const next = text.slice(from).search(/^## \[|^\[(?:Unreleased|\d+\.\d+\.\d+)\]: /m);
  return next < 0 ? text.length : from + next;
}

/** Removes headings ("### Added", "#### General", ...) with no entries before the next heading of their level. */
function dropEmptyHeadings(notes) {
  const level = (line) => /^(#{3,6}) /.exec(line)?.[1].length ?? 0;
  let lines = notes.split('\n');
  for (let changed = true; changed; ) {
    changed = false;
    lines = lines.filter((line, i) => {
      const own = level(line);
      if (!own) return true;
      let next = i + 1;
      while (next < lines.length && !lines[next].trim()) next++;
      const empty = next === lines.length || (level(lines[next]) > 0 && level(lines[next]) <= own);
      if (empty) changed = true;
      return !empty;
    });
  }
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

function release(version, date) {
  if (!/^\d+\.\d+\.\d+$/.test(version)) fail(`"${version}" is not a version like 1.2.3.`);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) fail(`"${date}" is not a date like 2026-10-09.`);
  let text = fs.readFileSync(FILE, 'utf8').replace(/\r\n/g, '\n');
  if (new RegExp(`^## \\[${escapeRe(version)}\\]`, 'm').test(text)) fail(`CHANGELOG.md already has a [${version}] section.`);

  const start = text.indexOf(UNRELEASED_HEADER);
  if (start < 0) fail(`CHANGELOG.md has no "${UNRELEASED_HEADER}" section.`);
  const bodyStart = start + UNRELEASED_HEADER.length;
  const end = sectionEnd(text, bodyStart);
  let notes = dropEmptyHeadings(text.slice(bodyStart, end).trim());
  if (!notes) {
    console.log('::warning::The [Unreleased] section of CHANGELOG.md has no entries; the release notes say "Maintenance release".');
    notes = EMPTY_NOTES;
  }
  text = `${text.slice(0, start)}${UNRELEASED_HEADER}\n\n## [${version}] - ${date}\n\n${notes}\n\n${text.slice(end).replace(/^\n+/, '')}`;

  // Compare links: [Unreleased] now starts at the new version; the new version compares with the previous one.
  const unreleasedLink = /^\[Unreleased\]: .*\/compare\/v([^.\s]+\.[^.\s]+\.[^.\s]+)\.\.\.HEAD$/m;
  const match = unreleasedLink.exec(text);
  const versionLink = match
    ? `[${version}]: ${REPO}/compare/v${match[1]}...v${version}`
    : `[${version}]: ${REPO}/releases/tag/v${version}`;
  const newUnreleased = `[Unreleased]: ${REPO}/compare/v${version}...HEAD`;
  text = match
    ? text.replace(unreleasedLink, `${newUnreleased}\n${versionLink}`)
    : `${text.trimEnd()}\n\n${newUnreleased}\n${versionLink}\n`;

  fs.writeFileSync(FILE, text.endsWith('\n') ? text : `${text}\n`);
  console.log(`CHANGELOG.md: added the [${version}] - ${date} section.`);
}

function notes(version) {
  const text = fs.readFileSync(FILE, 'utf8').replace(/\r\n/g, '\n');
  const header = new RegExp(`^## \\[${escapeRe(version)}\\][^\\n]*\\n`, 'm').exec(text);
  if (!header) fail(`CHANGELOG.md has no [${version}] section.`);
  const from = header.index + header[0].length;
  process.stdout.write(`${text.slice(from, sectionEnd(text, from)).trim()}\n`);
}

const [command, ...args] = process.argv.slice(2);
if (command === 'release' && args.length === 2) release(args[0], args[1]);
else if (command === 'notes' && args.length === 1) notes(args[0]);
else fail('Usage: changelog.js release <version> <yyyy-mm-dd> | changelog.js notes <version>');
