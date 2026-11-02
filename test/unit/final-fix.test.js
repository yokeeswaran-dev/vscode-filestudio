// Final release review, fix round: the texts a user reads must not promise features that 0.2.0 does not have.
'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

test('settings (0.2.1): the CSV editing setting exists, is on by default and is described like the other format switches', () => {
  const props = JSON.parse(read('package.json')).contributes.configuration.properties;
  const csv = props['fileStudio.csv.editing'];
  assert.ok(csv, 'CSV editing is on in 0.2.1');
  assert.equal(csv.type, 'boolean');
  assert.equal(csv.default, true);
  assert.equal(csv.scope, 'resource');
  assert.match(csv.description, /fileStudio\.editing\.enabled/);
});

test('README: the roadmap names no version for features that are not in 0.2.0', () => {
  const readme = read('README.md');
  assert.doesNotMatch(readme, /Coming in 0\.2\.1/);
  assert.doesNotMatch(readme, /0\.2\.x/);
  assert.match(readme, /Under consideration/);
});

test('texts: no "planned for 0.2.0", no "not available yet", no "unmerge" hint (0.2.0 has no unmerge)', () => {
  for (const f of ['SUPPORT.md', '.github/ISSUE_TEMPLATE/feature_request.md']) assert.doesNotMatch(read(f), /planned for 0\.2\.0/, f);
  assert.match(read('SECURITY.md'), /\| 0\.2\.x \|/);
  const sheet = read('src/renderers/sheet.ts');
  assert.doesNotMatch(sheet, /or unmerge first/);
  assert.doesNotMatch(sheet + read('media/viewer.js'), /not available yet/);
});
