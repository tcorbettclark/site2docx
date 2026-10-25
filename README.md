# site2docx

One command: a URL or a PDF goes in, a reviewable Word document comes out.

Every page becomes an image capture on its own landscape Word page (colour,
150 DPI, 1pt black border baked into the pixels) in the left column, with
the page's extracted source text beside it in the right column. Numbered
badges baked into the image mark each item at its true position — one badge
per text block or list entry (bulleted and numbered lists get one badge per
item), and one per non-empty cell of a table —
and the same numbers prefix the corresponding right-column paragraphs.
Reviewers use normal Word comments and track changes on the extracted
text. The document opens in Print Layout view.

## Usage

```sh
uv run site2docx URL_OR_PDF [OUTPUT_DOCX]
```

- **URL input** — the page is rendered in headless system Chrome (Playwright),
  all collapsed sections are expanded (toggles clicked via React state, with
  a class-surgery fallback), and the page is printed to a temporary PDF with
  Chrome's default headers/footers (date/title header, URL + "N/M" footer).
  The PDF is discarded after conversion.
- **PDF input** — only the PDF → Word conversion runs.

`OUTPUT_DOCX` is optional: for a URL it defaults to a filename derived from
the page title; for a PDF it defaults to the PDF's name with a `.docx`
extension.

## Example

```sh
uv run site2docx "https://adacomputerscience.org/concepts/coursedata_unit8?topic=proj_data_science_reviewer"
```

## Layout

Each Word page (A4 landscape) is a two-column section:

- **Left (13.6 cm)**: the rendered page image. Numbered badges (white circle,
  black outline) are baked in at each item's true position — just left of
  each block's first character, or just inside the top-right corner of each
  table cell (offset down and left from the corner so the badge clearly
  belongs to its own cell). Blocks hard against the page edge get a
  semi-transparent badge on their top-left corner instead.
- **Right (13.6 cm, 0.5 cm gutter)**: the page's text, reflowed into
  paragraphs at 10pt, each prefixed with its bold badge number. Every
  non-empty table cell is a numbered item of its own.

The whole document is a single two-column section; source pages are separated
by page breaks, so overflowing text flows onto an extra page rather than
being clipped. Chrome's print header/footer bands are filtered out of the
extraction. Badge numbers are drawn *after* text extraction, so they never
leak into the right column. Bullets may be text
glyphs *or* small vector shapes (the Ada CS site draws them as vector
squares); both are treated as list-item markers.

## Modules

- `cli.py` — entry point; dispatches URL vs PDF, handles temp files and naming
- `webpdf.py` — Playwright rendering: load, wait for the React app to settle,
  dismiss cookie banners, scroll to trigger lazy loads, expand collapsed
  sections, print to PDF
- `docxbuild.py` — PDF to Word: landscape two-column pages (image left,
  numbered, reflowed text right), badges baked into the page images

Tunables (DPI, border weight, column widths, print margins) are constants at
the top of `docxbuild.py` and `webpdf.py`.