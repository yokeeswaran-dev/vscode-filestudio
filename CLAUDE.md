# CLAUDE.md — FileStudio

Read this first. Every session and every subagent follows it. If a rule here conflicts with the user's spec
(`C:\AllDrive\customApp\.scratch\v020\SPEC-0.2.0.md`), the spec wins.

## Resume after a stop or context reset
1. Read this file, then `C:\AllDrive\customApp\.scratch\v020\RUN.md`, `PLAN.md` and `DECISIONS.md` (same folder).
2. Continue from where the work stopped: re-run or resume the saved workflow script
   (`C:\AllDrive\customApp\.scratch\v020\workflows\`). Do not redo finished phases.

## What this project is
- **FileStudio**: a VS Code extension (`package.json` name `filestudio`, publisher `Yokeeswaran`, VS Code `^1.123.0`) that shows
  xlsx, csv/tsv/psv/ssv, md, docx, pdf and pptx files in webviews. Published on the Marketplace: real users have it, so a
  regression or a corrupted file is a real problem.
- 0.1.x is view-only (only edit: Markdown task checkboxes). **0.2.0 adds editing** for xlsx, csv/tsv/psv/ssv and Markdown.
  docx/pdf/pptx stay view-only.

## Structure
- `src/extension.ts` — `activate()`: registers the 6 custom editors and 3 commands.
- `src/viewerProvider.ts` — message protocol types, CSP/HTML, messaging, link policy, file watching, `SheetDocument` +
  `SheetEditorProvider`, `TextViewerProvider` (`CsvView`, `MarkdownView`), docx/pdf/pptx providers, logging.
- `src/renderers/*.ts` (`sheet`, `markdown`, `docx`, `pdf`, `pptx`) — pure Node. **They never import `vscode`** (unit-tested in plain Node).
- `media/viewer.js` — the whole webview UI: **plain JavaScript + `// @ts-check` + JSDoc**, must pass TypeScript strict.
  `media/viewer.css` — UI colours come from `--vscode-*` variables (hex only as `var()` fallbacks or for document
  content such as slides); has a HIGH CONTRAST section.
- `esbuild.js` — builds `dist/` (extension, viewer ESM + lazy chunks, css, fonts, pdf.js). Keep the stubs for `unzipper` and `elkjs`.
- `test/unit/*.test.js` (node --test), `test/host/*.test.js` (host code against the fake `vscode` in `test/host/fake-vscode.js`),
  `test/build-tests.js` (bundles renderers, `viewer.js` with test-only exports, and host code to `test/.out/`),
  `test/generate-samples.js`, `test/samples/` (byte-exact fixtures).
- **No new files in `src/` or `media/`.** Put new code in the matching section of the existing file, using
  `// ===== NAME =====` and `// ----- sub -----`. If a split is truly needed for safety, log the reason in `DECISIONS.md` first.
- Outside `src/` and `media/` you may add: test files, an `@vscode/test-electron` setup, CI changes for it, this file.
- Code map with line anchors: `C:\AllDrive\customApp\.scratch\v020\notes\code-map.md`.

## Commands (run in the repo root)
| Command | What it does |
| --- | --- |
| `npm ci` | install exactly from `package-lock.json` (only when your task says so) |
| `npm run typecheck` | `tsc --noEmit` on `src/**/*.ts` |
| `npm test` | bundles renderers, `media/viewer.js` and host code to `test/.out/`; runs `test/unit` and `test/host` (fake `vscode`) with `node --test` (Node 22+: glob args) |
| `npm run build` | `node esbuild.js` (deletes and rebuilds `dist/`) |
| `npm run package` | production build (minified), as used for the VSIX |
| `npm run samples` | rebuilds `test/samples/*` (writes only files whose bytes changed); `large.xlsx` (not in Git) only if missing |
| `npx @vscode/vsce@4.0.0 package --no-dependencies` | builds the VSIX (CI does this on Linux) |
| F5 | Extension Development Host on `test/samples` (`.vscode/launch.json`) |

- All four checks (`npm ci`, typecheck, test, build) must pass before a commit on `feat/0.2.0`.
- `media/viewer.js` strict check (not in a script yet):
  `node node_modules/typescript/bin/tsc --ignoreConfig --noEmit --allowJs --checkJs --strict --target ES2022 --module preserve --moduleResolution bundler --lib ES2022,DOM,DOM.Iterable --types node --skipLibCheck media/viewer.js`
- Integration tests (`@vscode/test-electron`, planned for 0.2.0, **not there yet**; planned script name
  `npm run test:integration`, Linux CI runs it with `xvfb-run -a`). Until then use the workshop runner:
  `C:\AllDrive\customApp\.scratch\v020\s1-harness-vscode\howto.md`.

## Message protocol rule
- `HostMessage` / `WebviewMessage` live in `src/viewerProvider.ts` (section MESSAGE PROTOCOL); the model types they carry
  live in `src/renderers/sheet.ts`, `markdown.ts` and `pptx.ts`. `media/viewer.js` imports them with JSDoc
  `import(...)` typedefs at the top of the file.
- Keep both sides in sync in the same change. `WebviewMessage`: the TS union + `parseWebviewMessage` (validation + size
  limits) + the `post(...)` call in `viewer.js`. `HostMessage`: the TS union + `dispatch` in `viewer.js`.
  Prefer additive changes (new optional fields, new types).

## Data safety (top priority)
Losing or corrupting a user file is worse than a missing feature.
- **The extension host is the authority.** It checks every edit (read-only document, lossy file not confirmed with
  "Edit anyway", editing turned off by setting, stale version) and rejects it. The webview may show a change at once but must
  roll it back on rejection. Hiding buttons is not enough.
- **Never re-serialize an unmodified file.** xlsx: build the file in memory, **load that buffer again as a check**, and only
  then write (temp file + rename where possible). If the check fails: do not write, show an error, keep the edits.
- `.bak` once per file (setting); never overwrite an existing `.bak`; if it cannot be written, ask before saving.
- External change while the document is dirty → ask the user. Never overwrite newer text or a newer file.
- csv/Markdown change only through **minimal `WorkspaceEdit`s** on the `TextDocument` (only changed records/blocks);
  untouched bytes stay byte for byte. One user action = one undo step. The webview keeps no undo history of its own.
- ExcelJS output is NOT safe by default: a no-edit save already breaks some real files
  (see `C:\AllDrive\customApp\.scratch\v020\notes\gap-analysis.md`).
  Everything `loadWorkbook` patches must be written back or listed as lossy. Nothing is dropped silently.
- When in doubt: **block the action, show a clear message, and log it in `DECISIONS.md`**.

## Security and UI rules
- Strict CSP with a nonce (`buildCsp`). **Never weaken it** (no `unsafe-eval`, no `wasm-unsafe-eval`, no new remote sources).
- All HTML shown in a webview goes through DOMPurify (`sanitize` in `viewer.js`). Pasted HTML is parsed inert, never inserted.
- Links only through the existing link policy (`openLink` in `viewerProvider.ts`). Resources only through `asWebviewUri`.
- No telemetry, no network requests. Errors show in the page with "Reopen as Text" / "Reveal in Explorer", never a blank panel.
- Every new UI works in light, dark and high-contrast themes, with keyboard and screen reader, and respects Reduce Motion.
- Nothing that works in 0.1.2 may break: every CHANGELOG fix is a regression-test candidate.

## Licences and dependencies
- Allowed: **MIT, Apache-2.0, BSD-2-Clause, BSD-3-Clause, ISC** (a dual licence is fine if one option is allowed).
- Forbidden: SheetJS, Handsontable, HyperFormula, any GPL/LGPL/AGPL, paid packages. Check maintenance (recent releases, no
  critical open issues) before adding a dependency.
- Every runtime dependency goes into `THIRD_PARTY_NOTICES.md` with its full licence text. Commit `package-lock.json`.

## Git, commits, CHANGELOG
- `main` = released code only. `dev` = default, but **we cannot push to `dev`: never commit or merge into `dev`**.
- **All 0.2.0 work goes on ONE branch: `feat/0.2.0`.** Commit only there (no other 0.2.0 branches, no merge into `dev`),
  and **push it after every commit**: `git push origin feat/0.2.0`. A pull request `feat/0.2.0` → `dev` is opened by the user.
- **Never** push any other branch, merge into `main`, tag, create a GitHub release, or publish to the Marketplace.
- Only the agent whose task says "you own git for this step" changes git state. Others use read-only git (status, diff, log, show).
- Before a commit: run `git status`; add **explicit paths** only (never `git add -A` / `git add .`); nothing from `.scratch` is committed.
- Commit format `<type>(<scope>): <summary>` — types `feat|fix|docs|test|refactor|perf|build|chore`, scopes
  `sheet|csv|markdown|docx|pdf|pptx|host|build`; imperative, at most ~72 characters, no full stop.
- Every user-visible change gets a line under `## [Unreleased]` in `CHANGELOG.md`: simple English, grouped by format, same
  style as the existing entries.
- Release prep: `node .github/scripts/release.js prepare minor` on `feat/0.2.0`, then commit with explicit paths
  (`package.json package-lock.json CHANGELOG.md`), not the `git commit -am` it prints.

## Tests
- `test/samples/` are byte-exact fixtures (`.gitattributes`: `-text`) and also the F5 workspace. Tests never write into them:
  copy to a temp folder first. Add new samples through `test/generate-samples.js` (deterministic output).
- Load fresh objects in edit tests (`sheet.test.js` caches `styledModel()`).
- Unfinished harnesses live in the workshop; finished tests the product depends on go into `test/`.

## The workshop: `C:\AllDrive\customApp\.scratch`
- One level **above** the repo (outside Git), also listed in `.gitignore` and `.vscodeignore`. **Never committed.**
- This release round: `.scratch\v020\` (`BRIEF.md`, `SPEC-0.2.0.md`, `PLAN.md`, `RUN.md`, `DECISIONS.md`, `notes\`, `workflows\`).
- **Every agent works in its own folder** `.scratch\v020\<agent-name>\` with a `progress.md` (done / left / how to resume) and
  its own builds, logs, test files, Chrome and VS Code profiles (`--user-data-dir`, `--extensions-dir` inside that folder).
- Repo copies exclude `.scratch`, `node_modules`, `.git` (helper: `.scratch\v020\tools\private-build.js`).
- **No git inside the workshop:** never `git init`, `git clone` or `git worktree` there (user rule).
- Kill only processes you started (by PID), never by image name: the user's own VS Code, Chrome and Excel are open.

## How the work runs
- **Heavy work runs in workflows / subagents** (reading code, mapping, implementing, reviewing, testing, fixing).
- The **main session only coordinates**: it starts workflows, reads their short results, decides the next step and updates
  `RUN.md`, `PLAN.md` and `DECISIONS.md`. It does not read large files or long logs itself.
- Every agent returns a compact result: what changed, files touched, test numbers, open issues.
- Agents that edit the same big file (`viewer.js`, `viewerProvider.ts`, `sheet.ts`) run one after another.
