# Contributing to FileStudio

Thank you for helping to improve FileStudio! This guide explains how to set up the project, how the code is
organised, and how to send a change.

By taking part in this project you agree to follow our [Code of Conduct](CODE_OF_CONDUCT.md).
To report a security problem, do **not** open a public issue; read [SECURITY.md](SECURITY.md) instead.

## Ways to contribute

- **Report a bug** or **ask for a feature** with the [issue templates](https://github.com/yokeeswaran-dev/vscode-filestudio/issues/new/choose).
- **Fix a bug** or **improve a viewer** with a pull request.
- **Improve the documentation**, for example the README or this guide.

For larger changes, please open an issue first, so we can agree on the approach before you write the code.

## Prerequisites

- [Node.js](https://nodejs.org/) **22 or later** and **npm**
- [VS Code](https://code.visualstudio.com/) **1.123 or later**
- [Git](https://git-scm.com/)

## Set up the project

```bash
git clone https://github.com/yokeeswaran-dev/vscode-filestudio.git
cd vscode-filestudio
npm ci
npm run build
```

`npm ci` installs the exact dependency versions from `package-lock.json`.

## Build, run and test

| Command | What it does |
| --- | --- |
| `npm run build` | Bundles the extension and the webview into `dist/` with esbuild (development build with source maps). |
| `npm run watch` | Same as `build`, and rebuilds on every change. |
| `npm run typecheck` | Type-checks the TypeScript code (`tsc --noEmit`). |
| `npm test` | Bundles the renderers into `test/.out/` and runs the unit tests in `test/unit/` with Node's built-in test runner. |
| `npm run samples` | Regenerates the sample files in `test/samples/` (`node test/generate-samples.js`). |
| `npm run package` | Production build (minified, no source maps). Used when packaging. |

To try your change in VS Code:

1. Open the project folder in VS Code.
2. Press **F5** (launch configuration **Run Extension**). This builds the project and opens a new
   **Extension Development Host** window with the `test/samples` folder.
3. Open any sample file. After a change, run `npm run build` (or keep `npm run watch` running) and reload the
   window (**Developer: Reload Window**).

Useful for debugging:

- **View → Output → FileStudio** shows the extension log.
- **Developer: Open Webview Developer Tools** shows the console of a viewer.

About `npm run samples`: `large.xlsx` (100,000 rows × 50 columns, about 44 MB) is not in Git. The script creates it
when it is missing. Use `node test/generate-samples.js --skip-large` to skip it.

## Project layout

```text
src/
  extension.ts          Entry point: registers the custom editors and the commands
  viewerProvider.ts     Custom editor providers, webview HTML and CSP, message protocol, link handling
  renderers/
    sheet.ts            .xlsx (ExcelJS) and CSV / TSV / PSV / SSV models for the grid
    markdown.ts         Markdown to HTML (markdown-it, KaTeX, highlight.js)
    docx.ts             .docx to HTML (mammoth)
    pdf.ts              PDF header check and error messages
    pptx.ts             .pptx to slide data (pptxtojson)
media/
  viewer.js             Webview script for every view (grid, Markdown, DOCX, PDF, PPTX)
  viewer.css            Webview styles
esbuild.js              Build script (extension, webview, pdf.js worker and data files)
test/
  unit/                 Unit tests (Node's test runner)
  samples/              Sample files used for manual and automated tests
  generate-samples.js   Creates the sample files
  build-tests.js        Bundles the renderers for the unit tests
assets/
  icon.png              Extension icon
  screenshots/          Screenshots used in the README
```

The renderers in `src/renderers/` do not import `vscode`, so they can be tested with plain Node.
The message types shared by the extension and the webview are defined in `src/viewerProvider.ts`
(and `src/renderers/*.ts`). Keep them in sync with `media/viewer.js`.

## Coding style

- Follow the **[.editorconfig](.editorconfig)**: UTF-8, LF line endings, 2 spaces, a final newline.
- **TypeScript strict mode** is on (see `tsconfig.json`). `npm run typecheck` must pass with no errors.
- `media/viewer.js` is **plain JavaScript** with `// @ts-check` and JSDoc types. Keep the types correct.
- Large files are split into sections with comments in this form:

  ```ts
  // ===== SECTION NAME =====
  // ----- sub-section -----
  ```

  Put new code in the matching section, or add a new section.
- Write comments that explain **why**, not only what.
- Keep the security rules: all HTML shown in a webview must go through DOMPurify, do not weaken the
  Content Security Policy, and send links through the existing link handling in `viewerProvider.ts`.
- Do not add telemetry or network requests.

## Dependencies

- Add a new dependency only when it is really needed.
- **Licence rule:** only dependencies with a permissive licence are accepted: **MIT, Apache-2.0, BSD
  (2-Clause or 3-Clause) or ISC**. A dual licence is fine when one of its options is in this list.
  No GPL, LGPL, AGPL or other copyleft-only licences.
- Every new or updated runtime dependency must be added to (or updated in) **[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)**.
- Commit the updated `package-lock.json`.

## Branches and pull requests

| Branch | Purpose |
| --- | --- |
| `main` | Released code only. It changes only through a pull request from `dev` or `hotfix`, and every merge is a release, tagged `vX.Y.Z` automatically. |
| `dev` | Day-to-day development. **All pull requests target `dev`** (it is the default branch). |
| `hotfix` | Urgent fixes for the released version, made by maintainers and released straight into `main`. |
| `fix/…`, `feat/…`, `docs/…` | Short-lived branches for one change each, made from `dev`. |

After every release, `main` goes back into `hotfix` automatically and into `dev` by the maintainer (step 5 of the
release steps), so all three branches start from the released code. Always `git pull` before you start new work.

1. **Fork** the repository and create a branch from `dev` with a short, clear name, for example
   `fix/pdf-find-count` or `feat/xlsx-conditional-formatting`.
2. Make your change. Keep the pull request **small and focused** on one topic.
3. Add or update **unit tests** when you change a renderer.
4. Run the checks:

   ```bash
   npm run typecheck
   npm test
   npm run build
   ```

5. Test the change by hand in the Extension Development Host (F5), in a light and a dark theme.
6. Add a line to the `## [Unreleased]` section of **[CHANGELOG.md](CHANGELOG.md)** when users will notice the change.
7. Open a pull request **into `dev`** and fill in the template. Link the issue it fixes (for example `Fixes #12`).

The CI workflow builds and tests every pull request on Windows and Linux; it must be green before merging.
A maintainer will review your pull request and **squash-merge** it. Please answer review comments by pushing new
commits to the same branch.

### Commit messages

Use short messages in the imperative mood, with an optional type and scope:

```text
<type>(<scope>): <summary>
```

- **type:** `feat`, `fix`, `docs`, `test`, `refactor`, `perf`, `build` or `chore`
- **scope** (optional): `sheet`, `csv`, `markdown`, `docx`, `pdf`, `pptx`, `host` or `build`
- **summary:** at most about 72 characters, no full stop at the end

Examples:

```text
fix(pdf): keep the find count after a zoom change
feat(sheet): show data bars from conditional formatting
docs: explain the Markdown default editor setting
```

### Pull request checklist

- [ ] The change is focused on one topic.
- [ ] `npm run typecheck`, `npm test` and `npm run build` pass.
- [ ] The change was tested in VS Code (light and dark theme).
- [ ] Tests were added or updated where it makes sense.
- [ ] `CHANGELOG.md` is updated (for changes users will notice).
- [ ] New dependencies follow the licence rule and are listed in `THIRD_PARTY_NOTICES.md`.

## Release steps (maintainers)

Versions follow [Semantic Versioning](https://semver.org/): **patch** for fixes (0.1.0 → 0.1.1), **minor** for new
features (0.1.x → 0.2.0), **major** for 1.0.0 and breaking changes.

A release is a pull request from `dev` (or `hotfix`) into `main`:

1. **Bring the branch up to date and set the version** (on `dev`, or on `hotfix` for an urgent fix):

   ```bash
   git checkout dev
   git pull
   git merge origin/main        # only needed when main has commits that dev does not have
   node .github/scripts/release.js prepare patch    # or minor / major
   git commit -am "chore(release): vX.Y.Z"
   git push
   ```

   `prepare` bumps the version in `package.json` / `package-lock.json` and moves the `[Unreleased]` notes of
   `CHANGELOG.md` into a new `## [X.Y.Z] - date` section. Bump the version **only now**, at the end, so a hotfix
   released in the meantime does not take the same number.
2. **Open a pull request from `dev` into `main`** (for example "Release 0.1.1"). It can only be merged when:
   - CI passes on Windows and Linux,
   - the **Release check** passes: the pull request comes from `dev` or `hotfix`, its version is higher than
     `main`'s, that version is not tagged yet, `package-lock.json` matches, and `CHANGELOG.md` has a section for it
     (run `node .github/scripts/release.js check` locally to see the same result),
   - the branch is **up to date with `main`** (otherwise GitHub shows "Update branch"; or merge `origin/main` into it).
3. **Merge with "Create a merge commit"** (squash is turned off for `main`, so `dev` and `main` keep the same history).
4. The **Release** workflow then runs by itself: it tags the merge commit `vX.Y.Z` and merges `main` back into
   `hotfix`.
5. **Bring `main` into `dev`** (the `dev` rules require pull requests, so the workflow cannot push there; as a
   repository admin you can):

   ```bash
   git checkout dev
   git pull
   git merge origin/main
   git push
   ```

6. **Build the VSIX from `main` and publish it by hand:**

   ```bash
   git checkout main
   git pull
   npm ci
   npx @vscode/vsce@4.0.0 package --no-dependencies
   ```

   Upload `filestudio-X.Y.Z.vsix` on the
   [Marketplace publisher page](https://marketplace.visualstudio.com/manage/publishers/yokeeswaran) (**… → Update**).
7. Install the published version from the Marketplace and do a quick check of every format.

**Urgent fix (hotfix):** `git checkout hotfix`, `git pull`, make the fix and commit it, then follow the same steps
from step 1 with `hotfix` in place of `dev` (usually `prepare patch`). Step 5 then brings the fix into `dev` too.

**Branch rules** (Settings → Rules → Rulesets): `main` accepts only pull requests with merge commits, requires the
checks "Build and test (ubuntu-latest)", "Build and test (windows-latest)" and "Release check", and requires the
branch to be up to date. `dev` requires pull requests and CI; repository admins may bypass the rules. `hotfix` only
blocks deleting and force-pushing. No repository secrets are needed.

## Questions

If something in this guide is unclear, please open an issue or see [SUPPORT.md](SUPPORT.md).
