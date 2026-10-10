// Release helper for the three-branch flow (dev / hotfix -> main). See CONTRIBUTING.md, "Release steps".
//
//   node .github/scripts/release.js prepare <patch|minor|major>
//       On dev or hotfix, right before the release pull request: bumps the version in package.json and
//       package-lock.json and moves the CHANGELOG.md [Unreleased] notes into a new "## [<version>] - <today>" section.
//   node .github/scripts/release.js check
//       The rules a pull request into main must pass (also run by .github/workflows/release-check.yml):
//       it comes from dev or hotfix, its version is higher than main's, that version is not tagged yet,
//       package-lock.json has the same version, and CHANGELOG.md has a section for it.
'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync, execSync } = require('child_process');

const ROOT = path.join(__dirname, '..', '..');
const SOURCE_BRANCHES = ['dev', 'hotfix'];

const git = (...args) => execFileSync('git', args, { cwd: ROOT, encoding: 'utf8' }).trim();
const readJson = (file) => JSON.parse(fs.readFileSync(path.join(ROOT, file), 'utf8'));

/** "1.2.3" -> [1, 2, 3]; null for anything that is not a plain x.y.z version. */
function parseVersion(text) {
  const m = /^(\d+)\.(\d+)\.(\d+)$/.exec(String(text));
  return m ? m.slice(1).map(Number) : null;
}

/** > 0 when a is higher than b. */
function compareVersions(a, b) {
  for (let i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] - b[i];
  return 0;
}

function prepare(bump) {
  if (!['patch', 'minor', 'major'].includes(bump)) {
    console.error('Usage: node .github/scripts/release.js prepare <patch|minor|major>');
    process.exit(1);
  }
  execSync(`npm version ${bump} --no-git-tag-version`, { cwd: ROOT, stdio: 'inherit' });
  const version = readJson('package.json').version;
  const now = new Date(); // the local date, as the person preparing the release sees it
  const today = [now.getFullYear(), now.getMonth() + 1, now.getDate()].map((n) => String(n).padStart(2, '0')).join('-');
  execFileSync(process.execPath, [path.join(__dirname, 'changelog.js'), 'release', version, today], { cwd: ROOT, stdio: 'inherit' });
  console.log(`\nPrepared ${version}. Next steps:`);
  console.log('  1. Check CHANGELOG.md, then commit: git commit -am "chore(release): v' + version + '"');
  console.log('  2. Push the branch and open the pull request into main (merge it with a merge commit).');
}

function check() {
  const problems = [];
  const branch = process.env.HEAD_REF || git('rev-parse', '--abbrev-ref', 'HEAD');
  if (!SOURCE_BRANCHES.includes(branch)) {
    problems.push(`Only ${SOURCE_BRANCHES.join(' and ')} can be merged into main; this pull request comes from "${branch}".`);
  }
  if (process.env.HEAD_REPO && process.env.BASE_REPO && process.env.HEAD_REPO !== process.env.BASE_REPO) {
    problems.push(`Release pull requests must come from this repository, not from the fork ${process.env.HEAD_REPO}.`);
  }

  const version = readJson('package.json').version;
  const parsed = parseVersion(version);
  let mainVersion = null;
  try {
    mainVersion = JSON.parse(git('show', 'origin/main:package.json')).version;
  } catch {
    problems.push('Could not read package.json on origin/main (run "git fetch origin main" first).');
  }
  if (!parsed) problems.push(`package.json version "${version}" is not a plain x.y.z version.`);
  else if (mainVersion && parseVersion(mainVersion) && compareVersions(parsed, parseVersion(mainVersion)) <= 0) {
    problems.push(
      `package.json version ${version} is not higher than main's ${mainVersion}. ` +
        'On this branch run: node .github/scripts/release.js prepare patch (or minor / major), commit and push.',
    );
  }

  const lock = readJson('package-lock.json');
  if (lock.version !== version || (lock.packages && lock.packages[''] && lock.packages[''].version !== version)) {
    problems.push(`package-lock.json has version ${lock.version}, package.json has ${version}: run "npm install" and commit package-lock.json.`);
  }

  if (git('tag', '-l', `v${version}`)) {
    problems.push(`The tag v${version} already exists: this version was released before. Choose a higher version.`);
  }

  const changelog = fs.readFileSync(path.join(ROOT, 'CHANGELOG.md'), 'utf8');
  const escaped = version.replace(/\./g, '\\.');
  const section = new RegExp(`^## \\[${escaped}\\] - \\d{4}-\\d{2}-\\d{2}\\s*$([\\s\\S]*?)(?=^## \\[|^\\[[^\\]]+\\]: |(?![\\s\\S]))`, 'm').exec(changelog);
  if (!section) problems.push(`CHANGELOG.md has no "## [${version}] - YYYY-MM-DD" section (the prepare command writes it).`);
  else if (!/^- /m.test(section[1])) problems.push(`The CHANGELOG.md section for ${version} has no entries.`);

  if (problems.length) {
    for (const p of problems) console.error(`::error::${p}`);
    process.exit(1);
  }
  console.log(`Release check passed: ${branch} -> main as v${version} (main is ${mainVersion}).`);
}

const [command, arg] = process.argv.slice(2);
if (command === 'prepare') prepare(arg);
else if (command === 'check') check();
else {
  console.error('Usage: node .github/scripts/release.js prepare <patch|minor|major> | check');
  process.exit(1);
}
