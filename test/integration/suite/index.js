// Entry point that runs INSIDE the test VS Code (`--extensionTestsPath`, called by @vscode/test-electron). A small test
// runner (no test framework, D-066): it defines the globals suite / test / suiteSetup / suiteTeardown / setup / teardown,
// loads every test/integration/*.test.js, runs the tests one after another and prints one PASS / FAIL line per test.
// The returned promise rejects when a test or hook fails, so `runTests()` exits with code 1.
//
//   suite(title, fn)                 a group of tests; fn runs at once and only declares tests and hooks (suites may nest)
//   test(title, fn)                  a test: fn may be async; it fails when it throws or its promise rejects
//   test(title, { timeout }, fn)     the same with its own time limit in ms (default TEST_TIMEOUT_MS)
//   suiteSetup / suiteTeardown(fn)   run once before / after the tests of the suite (only when one of its tests runs)
//   setup / teardown(fn)             run before / after each test of the suite and of its inner suites
//
// Use node:assert in the tests. FS_IT_GREP = run only the tests whose full title (suite titles + test title, separated
// by spaces) contains this text.
'use strict';
const fs = require('fs');
const path = require('path');

/** Time limit of a test or hook when none is given. */
const TEST_TIMEOUT_MS = 60_000;

/**
 * @typedef {{ title: string, fn: () => unknown, timeout: number }} Runnable
 * @typedef {{ title: string, parent: Suite | undefined, children: (Suite | Runnable)[],
 *   suiteSetup: Runnable[], suiteTeardown: Runnable[], setup: Runnable[], teardown: Runnable[] }} Suite
 */

/** @returns {Suite} */
function newSuite(title, parent) {
  return { title, parent, children: [], suiteSetup: [], suiteTeardown: [], setup: [], teardown: [] };
}

function isSuite(node) {
  return Array.isArray(node.children);
}

function fullTitle(suite, title) {
  const parts = title ? [title] : [];
  for (let s = suite; s; s = s.parent) if (s.title) parts.unshift(s.title);
  return parts.join(' ');
}

/** Checks the arguments of test(title, [options], fn) and hooks; a `done` callback is not supported. */
function runnable(kind, title, optionsOrFn, maybeFn) {
  const fn = typeof optionsOrFn === 'function' ? optionsOrFn : maybeFn;
  const options = typeof optionsOrFn === 'function' ? {} : optionsOrFn || {};
  if (typeof fn !== 'function') throw new Error(`${kind} "${title}": the last argument must be a function`);
  if (fn.length > 0) throw new Error(`${kind} "${title}": callbacks (done) are not supported, return a promise instead`);
  const timeout = options.timeout === undefined ? TEST_TIMEOUT_MS : options.timeout;
  if (!(Number.isFinite(timeout) && timeout > 0)) throw new Error(`${kind} "${title}": timeout must be a positive number`);
  return { title, fn, timeout };
}

/** Runs fn with a time limit. Rejects with the error of fn or with a timeout error. */
function runWithTimeout(item) {
  let timer;
  const timedOut = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`Timed out after ${item.timeout} ms`)), item.timeout);
  });
  const fn = item.fn;
  const run = new Promise((resolve) => resolve(fn())); // a synchronous throw becomes a rejection
  return Promise.race([run, timedOut]).finally(() => clearTimeout(timer));
}

function formatError(err) {
  if (err && err.stack) return String(err.stack);
  return String(err);
}

function seconds(ms) {
  return (ms / 1000).toFixed(1) + ' s';
}

exports.run = async function run() {
  const root = newSuite('', undefined);
  let current = root;

  // ----- globals used by the *.test.js files -----
  global.suite = (title, fn) => {
    if (typeof fn !== 'function') throw new Error(`suite "${title}": the second argument must be a function`);
    const suite = newSuite(String(title), current);
    current.children.push(suite);
    const outer = current;
    current = suite;
    try {
      const result = fn();
      if (result && typeof result.then === 'function') throw new Error(`suite "${title}": the function must not be async`);
    } finally {
      current = outer;
    }
  };
  global.test = (title, optionsOrFn, maybeFn) => {
    current.children.push(runnable('test', String(title), optionsOrFn, maybeFn));
  };
  for (const hook of ['suiteSetup', 'suiteTeardown', 'setup', 'teardown']) {
    global[hook] = (optionsOrFn, maybeFn) => {
      current[hook].push(runnable(hook, `${hook} of "${fullTitle(current, '') || '(root)'}"`, optionsOrFn, maybeFn));
    };
  }

  // ----- load the test files -----
  const dir = path.resolve(__dirname, '..');
  const files = fs
    .readdirSync(dir)
    .filter((name) => name.endsWith('.test.js'))
    .sort();
  for (const name of files) {
    current = root;
    require(path.join(dir, name));
  }

  // ----- run -----
  const grep = process.env.FS_IT_GREP || '';
  const selected = (suite, test) => !grep || fullTitle(suite, test.title).includes(grep);
  const hasSelected = (suite) =>
    suite.children.some((node) => (isSuite(node) ? hasSelected(node) : selected(suite, node)));

  const failures = [];
  let passed = 0;
  const started = Date.now();
  const fail = (title, err) => {
    failures.push({ title, err });
    console.log(`  FAIL ${title}`);
  };

  /** Runs hooks in order; returns the first error (later hooks of the list are skipped), or undefined. */
  async function runHooks(hooks) {
    for (const hook of hooks) {
      try {
        await runWithTimeout(hook);
      } catch (err) {
        return { hook, err };
      }
    }
    return undefined;
  }

  /** The setup hooks from the outermost suite inwards, teardown hooks from the innermost suite outwards. */
  function eachHooks(suite, kind) {
    const list = [];
    for (let s = suite; s; s = s.parent) list.unshift(...s[kind]);
    return kind === 'teardown' ? list.reverse() : list;
  }

  async function runTest(suite, test) {
    const title = fullTitle(suite, test.title);
    const t0 = Date.now();
    let error;
    const setupFailed = await runHooks(eachHooks(suite, 'setup'));
    if (setupFailed) error = setupFailed.err;
    else {
      try {
        await runWithTimeout(test);
      } catch (err) {
        error = err;
      }
    }
    // teardown always runs (it closes editors, deletes temp files), also after a failure.
    for (const hook of eachHooks(suite, 'teardown')) {
      try {
        await runWithTimeout(hook);
      } catch (err) {
        if (!error) error = new Error(`${hook.title} failed: ${formatError(err)}`);
      }
    }
    if (error) fail(`${title} (${seconds(Date.now() - t0)})`, error);
    else {
      passed++;
      console.log(`  PASS ${title} (${seconds(Date.now() - t0)})`);
    }
  }

  async function runSuite(suite) {
    if (!hasSelected(suite)) return;
    if (suite.title) console.log(`\n${fullTitle(suite, '')}`);
    const setupFailed = await runHooks(suite.suiteSetup);
    if (setupFailed) {
      const skipped = suite.children.length;
      fail(`${setupFailed.hook.title} (${skipped} test(s) or suite(s) not run)`, setupFailed.err);
    } else {
      for (const node of suite.children) {
        if (isSuite(node)) await runSuite(node);
        else if (selected(suite, node)) await runTest(suite, node);
      }
    }
    const teardownFailed = await runHooks(suite.suiteTeardown);
    if (teardownFailed) fail(teardownFailed.hook.title, teardownFailed.err);
  }

  await runSuite(root);

  const total = passed + failures.length;
  console.log(`\n${passed} passing, ${failures.length} failing (${seconds(Date.now() - started)})`);
  if (total === 0) throw new Error(grep ? `No test title contains FS_IT_GREP "${grep}"` : 'No integration tests found');
  if (failures.length > 0) {
    failures.forEach((f, i) => console.log(`\n${i + 1}) ${f.title}\n${formatError(f.err)}`));
    throw new Error(`${failures.length} integration test(s) failed`);
  }
};
