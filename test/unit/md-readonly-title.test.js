'use strict';
// Round 4: the disabled task checkboxes of a read-only Markdown document follow the current reason, also when it changes
// while the document stays read-only. (Runs the real mdSetTasksReadOnly from media/viewer.js with a tiny stub.)
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const src = fs.readFileSync(path.join(__dirname, '..', '..', 'media', 'viewer.js'), 'utf8');
function grab(re) {
  const m = re.exec(src);
  assert.ok(m, `not found: ${re}`);
  return m[0];
}

test('mdSetTasksReadOnly updates the title when the reason changes while still read-only', () => {
  const code = [
    grab(/const mdReadOnlyTasks = new WeakSet\(\);/),
    grab(/const MD_READ_ONLY_TASK_TITLE = [^;]*;/),
    grab(/function mdSetTasksReadOnly\(container\) \{[\s\S]*?\r?\n\}\r?\n/),
  ].join('\n');
  const input = {
    disabled: false,
    title: '',
    removeAttribute(name) {
      if (name === 'title') this.title = '';
    },
  };
  const mdState = { readOnly: true, readOnlyReason: 'Editing is turned off (setting x)' };
  const mdTaskCheckboxes = () => [input];
  const run = new Function('mdState', 'mdTaskCheckboxes', `${code}\nreturn mdSetTasksReadOnly;`)(mdState, mdTaskCheckboxes);
  run({});
  assert.equal(input.title, 'Editing is turned off (setting x)');
  assert.equal(input.disabled, true);
  mdState.readOnlyReason = 'The file is read-only';
  run({});
  assert.equal(input.title, 'The file is read-only');
  mdState.readOnly = false;
  run({});
  assert.equal(input.title, '');
  assert.equal(input.disabled, false);
});
