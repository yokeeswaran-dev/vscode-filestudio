# Security Policy

FileStudio opens files that may come from anyone, so we take security seriously.
Thank you for helping to keep FileStudio and its users safe.

## Supported versions

Security fixes are made for the latest release line only.

| Version | Supported |
| --- | --- |
| 0.1.x | ✅ Yes |
| < 0.1.0 | ❌ No |

Please update to the latest version before you report a problem.

## Reporting a vulnerability

**Please do not report security problems in public issues, discussions or pull requests.**

Report them privately through GitHub:

1. Go to the **[Security tab](https://github.com/yokeeswaran-dev/vscode-filestudio/security)** of the repository.
2. Click **Report a vulnerability**, or open this link directly:
   <https://github.com/yokeeswaran-dev/vscode-filestudio/security/advisories/new>
3. Fill in the form and submit it. Only the maintainers can see your report.

### What to include

- A short description of the problem and its **impact** (what an attacker can do).
- The **FileStudio version**, **VS Code version** and **operating system**.
- **Steps to reproduce**, and a **sample file** that shows the problem if possible.
  Please remove any private or personal data from the file first.
- A proof of concept, logs or screenshots, if you have them.
- Any idea you have for a fix (optional).

## What to expect

- We will **acknowledge your report within 7 days**.
- We will investigate, keep you informed of our progress, and may ask you for more details.
- When the problem is confirmed, we will work on a fix and publish a new release and a GitHub security advisory.
- We are happy to credit you in the advisory, unless you prefer to stay anonymous.

FileStudio is a small open-source project, so these times are **best effort**. Please give us a reasonable time to
release a fix before you share details of the problem in public.

## Scope

Reports about these areas are especially welcome:

- **Webview isolation:** ways around the Content Security Policy, running scripts or loading content that the
  policy should block.
- **HTML sanitising:** HTML or SVG from a Markdown file, a Word document or a slide that gets past DOMPurify
  (for example script injection or event handlers).
- **Link handling:** links that open something they should not, for example running a `command:` URI,
  bypassing the allowed schemes, or opening unexpected local files without a user click.
- **Parsing untrusted files:** a crafted `.xlsx`, `.csv`, `.md`, `.docx`, `.pdf` or `.pptx` file that leads to code
  execution, reading files it should not, or making network requests.

## Out of scope

- Problems in **VS Code itself** or in other extensions. Please report them to those projects.
- Problems in **third-party libraries** that FileStudio does not expose. Please report them to the library;
  tell us too if FileStudio is affected.
- A file that is **slow to open or uses a lot of memory** without other impact (for example a very large workbook).
  Please report this as a normal bug.
- Content that looks different from the original application. Please report this as a normal bug.
- Remote `https:` images in a Markdown file being loaded. This is expected behaviour (see Privacy & security in the
  [README](README.md#privacy)).
- Attacks that need an attacker who can already change your VS Code settings, extensions or files on your computer.
