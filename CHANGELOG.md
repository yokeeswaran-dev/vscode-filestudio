# Changelog

All notable changes to FileStudio are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-10-08

First public release. This version is **view-only**.

### Added

#### General

- Custom editors for `.xlsx`, `.csv`, `.tsv`, `.psv`, `.ssv`, `.md`, `.markdown`, `.docx`, `.pdf` and `.pptx` files.
  All formats except Markdown open in FileStudio by default; Markdown is available as an option.
- **Open with FileStudio** command in the Explorer context menu, the editor title bar and the Command Palette
  (also for several selected files at once).
- **Reopen as Text** command for CSV, TSV, PSV, SSV and Markdown files.
- Settings `fileStudio.markdown.defaultMode` and `fileStudio.xlsx.backupOnSave` (reserved for editing).
- Views follow the VS Code colour theme, including high contrast themes.
- Strict Content Security Policy for every view and HTML sanitising with DOMPurify.
- Links: `http`, `https` and `mailto` open outside VS Code, local files open in VS Code, other schemes are blocked.
- "FileStudio" output channel for logs, and an error view with **Reopen as Text** or **Reveal in Explorer**.
- No telemetry.

#### Spreadsheets (`.xlsx`)

- Virtual, Excel-like grid that handles large sheets (tested with 100,000 rows × 50 columns).
- Cell styles: fonts, theme and indexed colours, solid, pattern and gradient fills, borders, alignment,
  rotated text and rich text.
- Excel number formats, including dates, currency, percent, fractions and custom formats with colours.
- Merged cells, frozen panes, hidden rows and columns, right-to-left sheets and the saved sheet zoom.
- Embedded images, cell notes on hover, and hyperlinks (Ctrl+Click or Ctrl+Enter), including links to other sheets.
- Sheet tabs with tab colours and a menu to view hidden and very hidden sheets.
- Formula bar with the value or formula of the active cell, and a Name Box to go to a reference (Ctrl+G).
- Status bar with Count, Sum, Average, Min and Max of the selection, the sheet size and zoom controls.
- Copy (Ctrl+C) as tab-separated text plus an HTML table; a single cell is copied as plain text.
- Read-only cell text box on double-click.
- Excel-like keyboard navigation, F6 / Shift+F6 to move between regions, and screen-reader support.
- Banner that lists content the viewer does not show (charts, pivot tables, macros, slicers and more).

#### Delimited text (`.csv`, `.tsv`, `.psv`, `.ssv`)

- Same grid as for workbooks.
- Delimiter detection for `.csv` (comma, semicolon, tab or pipe) and `.ssv` (semicolon or runs of spaces).
- Grid stays in sync with the text document, so edits in the text editor appear right away.
- Files over 50 MB open in the text editor with a message (VS Code limit for extensions).

#### Markdown (`.md`, `.markdown`)

- GitHub-style preview with tables, task lists that can be ticked, footnotes, emoji shortcodes and GitHub alerts.
- YAML front matter shown as a table.
- KaTeX math, Mermaid diagrams (themed) and syntax highlighting with copy buttons on code blocks.
- Heading links, links to other files, local and remote `https:` images.
- Live update while the file is edited, with the scroll position kept.

#### Word documents (`.docx`)

- Page-like view of the document content (headings, formatting, lists, tables, images, links and footnotes),
  converted with mammoth.
- In-document bookmark links and collapsed conversion warnings.

#### PDF (`.pdf`)

- Page view with pdf.js: selectable text, zoom, fit width, fit page and Ctrl+Mouse wheel zoom.
- Page navigation, outline (bookmarks) sidebar and Find (Ctrl+F) with match case.
- Internal and external links, and a password prompt for encrypted files.

#### Presentations (`.pptx`)

- Slide view with zoom and fit, slide thumbnails (hidden slides marked) and a speaker notes pane.
- Text, shapes, connectors, pictures, tables with merged cells, groups, backgrounds and equations.
- Charts with ECharts: bar and column, line, area, pie, doughnut and scatter.
- Links to other slides and to web pages.
- Banner that lists content the viewer does not show (video, audio, embedded objects and more).

[Unreleased]: https://github.com/yokeeswaran-dev/vscode-filestudio/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/yokeeswaran-dev/vscode-filestudio/releases/tag/v0.1.0
