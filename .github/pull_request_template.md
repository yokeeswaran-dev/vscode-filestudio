## Summary

<!-- What does this pull request change, and why? Keep it short. -->

## Linked issue

<!-- For example: Fixes #123 (or "None" for small changes). -->

Fixes #

## Type of change

<!-- Put an x in the boxes that apply: [x] -->

- [ ] Bug fix (non-breaking change that fixes a problem)
- [ ] New feature (non-breaking change that adds something)
- [ ] Breaking change (existing behaviour changes)
- [ ] Performance improvement
- [ ] Refactoring (no change in behaviour)
- [ ] Documentation
- [ ] Tests, build or tooling

## How was this tested?

<!--
Describe the tests you ran: unit tests, manual tests in the Extension Development Host (F5), sample files used,
light / dark theme, operating system.
-->

- [ ] `npm run typecheck`
- [ ] `npm test`
- [ ] `npm run build`
- [ ] Manual test in VS Code (F5) with: <!-- file(s) -->

## Screenshots

<!-- For visual changes: before and after, if possible in a light and a dark theme. Delete this section if not needed. -->

## Checklist

- [ ] My change is focused on one topic.
- [ ] I followed the coding style in [CONTRIBUTING.md](../CONTRIBUTING.md) (`.editorconfig`, section comments, strict TypeScript, `// @ts-check` in `viewer.js`).
- [ ] I added or updated tests where it makes sense.
- [ ] I updated `CHANGELOG.md` (`## [Unreleased]`) for changes users will notice.
- [ ] I updated the documentation (README, settings descriptions) if needed.
- [ ] New or updated dependencies use an allowed licence (MIT, Apache-2.0, BSD or ISC) and are listed in `THIRD_PARTY_NOTICES.md`.
- [ ] HTML shown in a webview is still sanitised, and the Content Security Policy is not weakened.
