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

- [Node.js](https://nodejs.org/) **20 or later** and **npm**
- [VS Code](https://code.visualstudio.com/) **1.90 or later**
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
| `main` | Released code only. Every release is a tag on `main` (`v0.1.0`, `v0.1.1`, …). |
| `dev` | Day-to-day development. **All pull requests target `dev`** (it is the default branch). |
| `fix/…`, `feat/…`, `docs/…` | Short-lived branches for one change each. |

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

1. Make sure the `[Unreleased]` section of `CHANGELOG.md` on `dev` lists the changes.
2. Open a pull request from `dev` into `main` (for example "Release 0.1.1"), wait for CI, and merge it
   (use a merge commit, not squash, so `dev` and `main` keep the same history).
3. Go to **Actions → Release → Run workflow**, keep the branch `main`, choose **patch / minor / major**, and run it.
   The workflow:
   - bumps the version in `package.json` / `package-lock.json`,
   - moves the `[Unreleased]` notes into a new `## [X.Y.Z] - date` section of `CHANGELOG.md`,
   - type-checks, tests and packages `filestudio-X.Y.Z.vsix`,
   - commits `chore(release): vX.Y.Z`, tags `vX.Y.Z` and pushes to `main`,
   - creates the GitHub release with the `.vsix` and the CHANGELOG notes,
   - publishes to the VS Code Marketplace (needs the `VSCE_PAT` secret),
   - merges `main` back into `dev`.
4. Install the published version from the Marketplace and do a quick check of every format.

**Urgent fix (hotfix):** branch from `main` (`hotfix/…`), open a pull request into `main`, run the Release workflow
with **patch**, then make sure `main` is merged back into `dev`.

**Repository secrets** (Settings → Secrets and variables → Actions):

| Secret | Needed for |
| --- | --- |
| `VSCE_PAT` | Publishing. An Azure DevOps personal access token with the scope **Marketplace → Manage** (organisation: *All accessible organizations*). |
| `RELEASE_TOKEN` | Only when branch rules block the workflow from pushing to `main` / `dev`: a fine-grained personal access token of a maintainer who is on the rules' bypass list, with **Contents: Read and write** on this repository. |

**Manual release (fallback):** `npm version patch --no-git-tag-version`, update `CHANGELOG.md`,
`npx @vscode/vsce package --no-dependencies`, commit, `git tag vX.Y.Z`, `git push origin main --tags`, then upload
the `.vsix` on the [Marketplace publisher page](https://marketplace.visualstudio.com/manage/publishers/yokeeswaran)
or run `npx @vscode/vsce publish --no-dependencies --packagePath filestudio-X.Y.Z.vsix`.

## Questions

If something in this guide is unclear, please open an issue or see [SUPPORT.md](SUPPORT.md).
