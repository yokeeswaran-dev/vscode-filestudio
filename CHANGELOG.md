# Changelog

All notable changes to FileStudio are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.2] - 2026-10-11

### Fixed

#### General

- **Open Changes** and other diffs of a `.pdf`, `.docx`, `.xlsx` or `.pptx` file now show FileStudio views,
  not a text diff.
- A file that was deleted can now be shown from Git (for example in **Open Changes**).
- **Open with FileStudio** on a text tab now replaces that tab instead of adding a second one.
  A tab with unsaved changes stays open.
- **Split Editor Right** on an `.xlsx` file now shows the workbook in both groups, not an empty group.
- **Open with FileStudio** no longer shows on folders whose name ends in `.csv` or `.md`.
- In PDF, Word and Markdown views, the arrow keys and Space scroll again after you come back from another
  part of VS Code.

#### Spreadsheets (`.xlsx`)

- A damaged threaded-comment part can no longer freeze VS Code.
- The "not displayed" banner no longer lists threaded comments when they are shown.

#### Delimited text (`.csv`, `.tsv`, `.psv`, `.ssv`)

- A first line with a space, such as `sep= ;`, is also understood.
- In `;` files, numbers are read with one rule: `,` is the decimal mark and `.` groups thousands.
- The formula bar shows the number as written in the file (`007`, long numbers, `1e5`, `50%`), the same as
  the cell and the copy.
- Long numbers (14 digits or more) fit their column and no longer show as `1.23457E+13`.
- Status bar statistics now work for columns after `XFD`.
- Empty lines at the end of the file no longer add empty rows.
- Large files react faster to edits, and a hidden view is updated only when it is shown again.
- Ctrl+PageDown and Ctrl+PageUp switch VS Code editors again.
- The status bar shows the format and delimiter in use (for example `CSV · Semicolon`). The new command
  **FileStudio: Change Delimiter** lets you pick another delimiter for this session.

#### PDF (`.pdf`)

- On a VS Code version that is too old for the PDF engine, the view says so, instead of staying blank.
- A PDF with no pages, or whose first page cannot be read, shows an error instead of a blank grey view.
- A custom zoom (for example 160%) and the position inside a tall page come back after a reload.
- Find: F3 / Shift+F3 and Ctrl+G / Ctrl+Shift+G go to the next or previous match, Alt+C turns on match case,
  and reopening Find after Escape highlights the matches again.
- Page Down / Page Up scroll one screen, and End / Ctrl+End go to the end of the document.
- In a narrow view, the toolbar and the find bar no longer hide buttons.
- Better screen-reader and keyboard support: the toolbar, zoom button and outline have clear names, and the
  outline works with the arrow keys, Home and End.
- pdf.js warnings no longer fill the developer console.

#### PowerPoint (`.pptx`)

- `.pptm`, `.ppsx`, `.ppsm`, `.potx` and `.potm` files can now be opened.
- A missing picture or a damaged slide no longer stops the whole deck from opening; a placeholder is shown.
- Charts: combo charts show every series (with a secondary axis), date axes show dates, and the title,
  legend, number format, gridlines, text size and data labels follow the chart. Horizontal bars are in
  PowerPoint's order, and names with `&` or `<` show correctly.
- Text: line spacing and empty lines for small text, bullets and numbering from the layout or master, soft
  line breaks, vertical text, highlight, double/dotted/wavy underlines, tabs and percent line spacing now
  look like PowerPoint.
- Text inside ellipses, arrows, callouts and other shapes uses the shape's text area.
- Block arrows, callouts and other shapes have the right form, and outlines without a set width are drawn.
- Table cell margins are used, and tint/shade colours (for example in table styles) and gradients match
  PowerPoint.
- Hyperlink text uses the theme's link colour.
- Pictures cropped at the top or bottom are no longer squashed, and SVG-only pictures are shown.
- Embedded objects, linked pictures and SmartArt without a drawing show a box or their preview picture.
- Next/Previous/First/Last slide links work, and links on shapes and pictures work with the keyboard and
  screen readers.
- Screen readers announce each slide thumbnail's real position (for example 220 of 220).

### Changed

- FileStudio now needs VS Code 1.123 or newer (the PDF engine needs it).
- New Marketplace name and description: **FileStudio – Excel, CSV, Word, PDF & PowerPoint Viewer**. The
  extension is also listed under the **Visualization** and **Data Science** categories, with more search keywords.
- Setting descriptions now say that `backupOnSave` and the `split` / `wysiwyg` modes are planned for 0.2.0.
- New README: every supported file type at a glance, a section per format, a short FAQ, and how to reopen a
  file as text. It asks you to uninstall the old "File Viewer" extension.
- The package is smaller: only the `woff2` KaTeX fonts are included.
- The echarts NOTICE text is now included in the third-party notices.

## [0.1.1] - 2026-10-11

### Fixed

#### General

- FileStudio now works in Restricted Mode (untrusted workspaces). Before, it was turned off there.
- VS Code's **Reduce Motion** setting now stops the animations in every view (spinners, loading shimmer, fades,
  turning arrows and smooth scrolling).
- A `.csv`, `.tsv` or `.md` file that VS Code considers binary now reopens in the text editor with a message,
  instead of a generic error.
- Files from Git (for example the left side of **Open Changes**) no longer show as corrupt for `.docx`, `.xlsx`,
  `.pptx` and `.pdf`.
- Links with `#heading` open the heading first; VS Code line links (`#10`, `#L5,3`, `#L2-L3`) open at that line.
- A link to a UNC path (`\\server\share\...`) that VS Code does not allow now shows VS Code's reason,
  not "could not find".
- A file that cannot be read now writes one error to the log, not a new one on every refresh.

#### Spreadsheets (`.xlsx`)

- Date and time formats as other apps save them (`dd.mm.yyyy`, `hh.mm`, `yyyy-mm-ddThh:mm:ss`, `yyyy年m月d日`)
  show the date, not the raw number.
- Scientific formats with a space before the exponent (`0.00 E+00`) no longer show as a date.
- Formats with unit letters (`0.0x`, `0.0 °C`) keep the unit and round like Excel.
- The General format rounds like Excel (`8238230.9475` shows `8238230.948`).
- Dates and times on 31 December 9999, also in elapsed formats like `[h]:mm`, no longer show an empty cell.
- `_x000D_` and other `_xHHHH_` codes in text and formula results are decoded, as in Excel.
- Threaded comments show the conversation, not Excel's "Your version of Excel..." note.
- Dialog sheets and Excel 4.0 macro sheets are now listed in the "not displayed" banner.
- The "not displayed" banner now updates when the workbook changes on disk.
- Copy (Ctrl+C) no longer copies hidden or filtered-out rows and columns.
- Text cells pasted into Excel stay text (`007` no longer becomes `7`, `1-2` no longer becomes a date).
- Status bar: hidden and filtered-out rows are no longer counted, a selection with an error value shows only
  Count, and Sum, Average, Min and Max use the cells' number format (dates, percent, currency), as in Excel.
- The context menu **Copy** now works, and the keyboard keeps working after the menu closes.
  A cell text box stays open on right-click.
- Ctrl+A, a triple-click or a drag no longer highlights the whole page (tabs, status bar).
- F2 opens the cell text box. Ctrl+Backspace, Shift+Backspace and Ctrl+. now work as in Excel.
- A click on a column or row header next to merged cells selects only that column or row.
- Ctrl+Click on a selected cell removes it from the selection, as in Excel.
- Ctrl+Mouse wheel zooms in steps of 15%, as in Excel.
- A click on a disabled sheet-scroll arrow no longer stops the arrow keys.
- In a very narrow editor, the zoom buttons stay visible in the status bar.

#### Delimited text (`.csv`, `.tsv`, `.psv`, `.ssv`)

- A first line `sep=;` (as Excel writes it) sets the delimiter and is not shown as a row.
- The sheet tab name is cut to 31 characters and shows `( )` for `[ ]`, as in Excel.

#### Markdown (`.md`, `.markdown`)

- After a reload, a large document reopens at the same place.
- A jump into a part of a large document that was not shown yet lands at the right place.
- In a read-only document (for example a Git version or `files.readonlyInclude`), task checkboxes are disabled
  and can no longer change the document.
- The code block copy button copies only the code you see. Hidden text or a fake copy button in the document
  can no longer put other text on the clipboard.

#### Markdown and Word documents

- A jump to an empty anchor (`<a name>` in Markdown, a table-of-contents or bookmark link in Word) highlights the
  heading or paragraph it marks, and Tab continues from there.

#### Word documents (`.docx`)

- Links to Windows paths (`C:\...` or `C:/...`) open the file.
- Words in tables are no longer split in the middle when the table fits the page.
- Left-to-right paragraphs that start with a Hebrew or Arabic word are no longer aligned to the right.
- An empty (0-byte) file shows an empty document, as in Word, not an error.

#### PDF (`.pdf`)

- Fit width and the automatic zoom keep the widest page (for example a landscape page) inside the view,
  so Find no longer moves the pages sideways when the outline is open.

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

[Unreleased]: https://github.com/yokeeswaran-dev/vscode-filestudio/compare/v0.1.2...HEAD
[0.1.2]: https://github.com/yokeeswaran-dev/vscode-filestudio/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/yokeeswaran-dev/vscode-filestudio/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yokeeswaran-dev/vscode-filestudio/releases/tag/v0.1.0
