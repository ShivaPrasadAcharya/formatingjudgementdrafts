# Judgment Draft Structurer / फैसला मस्यौदा संरचना

Live website:

https://shivaprasadacharya.github.io/formatingjudgementdrafts/

A browser-based formatter for Nepali judgment drafts.

Upload DOCX, DOC, RTF or TXT, or paste text. Preserve original
numbering or create structured automatic Word numbering.

The existing examples, section detection, parties, inline emphasis,
supported colours and highlights, footer metadata, filename generation,
and download options remain available.

## Quick start

1. Upload a document or paste text.
2. Select a numbering mode.
3. Choose fonts and footer behaviour.
4. Press **Structure** after editing input.
5. Review the preview.
6. Download **Output .docx**.

For exact original Word numbering, upload a DOCX file and select
**Keep original numbering intact — मूल नम्बर यथावत्**.

## Numbering modes

| Mode | Behaviour |
| --- | --- |
| Keep original numbering intact | For DOCX, preserves original Word list definitions, numbers, restarts, bullets, hierarchy, tables and section settings. For pasted text, retains typed numbering, punctuation, order and blank lines. |
| Auto-detect | Existing default. Uses section numbering when at least two short subheadings are detected; otherwise paragraph numbering. |
| Sections | Subheadings become १., २. and their paragraphs become १.१, १.२. Text before the first subheading remains unnumbered. |
| Paragraphs only | Numbers paragraphs continuously across facts and findings. Headings remain unnumbered. तपसिल starts a separate list at १. |
| Irrespective | Works without all three main headings. Numbers the body continuously, including तपसिल items. Recognised main headings and closing signatures remain unnumbered. |

### Numbering levels

For regenerated numbering, typed numbers indicate the level,
not the final number.

Examples:

- `१.` — level 1
- `१.१` — level 2
- `१.१.१` — level 3
- `१.१.१.१` — level 4
- `१.१.१.१.१` — level 5

English digits can also be used as input.

A skipped level is reduced to the next valid level.
Party lists use separate numbering sequences.

The **New numbering digits** selector chooses Nepali or English
digits for regenerated lists. It does not alter preserved originals.

### Recognised main headings

- तथ्यगत बेहोरा / तथ्यखण्ड
- यस अदालतको ठहर / ठहरखण्ड
- तपसिल / तपसिलखण्ड

Heading detection is heuristic. Review the preview before downloading.

## Original DOCX mode

The downloaded DOCX is built from the uploaded DOCX package,
not reconstructed from the extracted plain text.

The following remain in place:

- Original numbering definitions.
- Paragraph numbering references.
- Numbering starts and restarts.
- List levels and bullets.
- Tables, images and headers.
- Section and page settings.
- Existing footers, unless removal is selected.

Selected text and numbering fonts are applied during export.

Original bullet symbol fonts are retained so bullet symbols are
not replaced by an incompatible font.

Changing fonts can change line wrapping and pagination even though
page settings and document structure are retained.

### Editing in original mode

The input is read-only when preserving an uploaded DOCX.

Switch to a renumbering mode to edit and restructure text.
Switching back to original mode restores the original imported text.

The font-size selector is disabled in original DOCX mode so
original text sizes remain intact.

The DOC download and “Use output as input” controls are disabled
in original DOCX mode because conversion to RTF cannot faithfully
retain arbitrary Word numbering and section structure.

Use the DOCX download for preservation.

### Preview limitations

The preview displays extracted text and common Word list numbering,
including style-inherited numbering and start overrides.

It is not a page-layout or pagination renderer.

Complex linked numbering styles and advanced Word fields may not
preview exactly. Original-mode DOCX downloads retain the original
numbering definitions independently of the text preview.

## Fonts

Default export fonts:

| Content | Font |
| --- | --- |
| Devanagari text | Kalimati |
| English text | Calibri |
| Paragraph numbering | Fontasy Himali |

DOCX export assigns fonts to script-specific text runs.
This also applies to supported header and footer text.

Fonts are selected, not embedded.

Install Kalimati, Calibri and Fontasy Himali on the computer opening
the downloaded document.

Some legacy Fontasy Himali variants may render Unicode numerals
differently. If necessary, select Kalimati as the number font.

Preserved original characters are not transliterated into legacy
font character codes.

## Footer options

### Keep original / generate for structured output

- Original DOCX mode retains the uploaded document's existing footer.
- Structured output generates the existing case-information and
  page-number footer.
- Preserved pasted text receives the generated page-number footer.

### Remove footer

Removes footer references from all DOCX sections.

This option works independently of the selected numbering mode.

Headers are not removed.

## Recent files

The latest **five** distinct opened files are shown.

- Newest files appear first.
- Reopening a file moves it to the top.
- Matching name, size and last-modified time identify the same file.
- Files are stored in this browser's IndexedDB.
- Recent files survive reloads when browser storage is available.
- Private browsing or clearing site data may remove the history.

The recent-file list belongs to the current browser and device.
It is not synchronized between devices.

## File support

### DOCX

Recommended format.

Supports original-document preservation and extraction of common
automatic numbering and inline text formatting.

### DOC

Legacy binary Word files are read on a best-effort basis.

If extraction is incorrect, open the document in Word and save it
as DOCX before uploading.

### RTF

Supports basic text and supported inline formatting.

Exact preservation of automatic lists is not guaranteed.
Convert to DOCX first when exact numbering matters.

### TXT

Retains visible text and typed numbering.

TXT does not contain automatic Word numbering or character formatting.

### DOC download

The `.doc` download is an RTF compatibility document with a `.doc`
extension, matching the existing application's export approach.

It is not a binary Word DOC file.

DOCX is recommended for the most reliable output.

## Privacy

Documents are processed in the browser.

The application does not send uploaded documents to a
document-processing server.

Recent files are stored locally in browser storage.

External CDN requests load these JavaScript libraries:

- JSZip
- Mammoth
- docx

An internet connection is required to load those libraries unless
they are already cached.

## Repository files

- `index.html`
  - Interface and styles.
  - Existing judgment parsing and structuring.
  - Examples.
  - Recent-file storage.
  - Base import and export functions.

- `upgrade.js`
  - Original DOCX preservation.
  - Word numbering preview extraction.
  - Script-specific DOCX fonts.
  - Footer removal.
  - Numbering-mode guidance.
  - Download state management.

- `README.md`
  - Instructions and limitations.

- `favicon.svg`
  - Existing website icon.

- `.github/workflows/static.yml`
  - Existing GitHub Pages deployment workflow.

Keep `upgrade.js` in the same directory as `index.html`.

## Update through GitHub

1. Open the repository.
2. Replace `index.html`.
3. Add `upgrade.js`.
4. Add or replace `README.md`.
5. Keep the existing favicon and `.github` folder.
6. Commit the changes to `main`.
7. Open the Actions tab.
8. Wait for the GitHub Pages deployment to succeed.
9. Open the live website and refresh with Ctrl+F5.

Suggested commit message:

Improve numbering preservation, export fonts and recent files

## Hosting

The existing GitHub Actions workflow deploys the static repository.

In repository Settings → Pages, use GitHub Actions.

No build step is required.

For local development, serve this directory with any static
HTTP server and open `index.html`.

## Review before use

Compare the downloaded document with the original in Word,
particularly numbering, page breaks and footers.

Original DOCX mode preserves document content rather than generating
or rewriting legal text.
