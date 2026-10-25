# site2docx

One command: a URL or a PDF goes in, a reviewable Word document comes out.

Every page becomes an image capture on its own Word page (colour, 150 DPI,
1pt black border baked into the pixels), so reviewers can't — and won't try
to — edit the text. A "Comments" column sits beside each page image for the
reviewer to type into.

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

Each Word page (A4) is a two-column table:

- **Left (13.6 cm)**: the rendered page image, centred, 1pt black border.
- **Right (5.4 cm)**: a bold "Comments" heading with a rule under it, plus
  blank space.

## Modules

- `cli.py` — entry point; dispatches URL vs PDF, handles temp files and naming
- `webpdf.py` — Playwright rendering: load, wait for the React app to settle,
  dismiss cookie banners, scroll to trigger lazy loads, expand collapsed
  sections, print to PDF
- `docxbuild.py` — PDF to Word: one bordered image + Comments column per page

Tunables (DPI, border weight, column widths, print margins) are constants at
the top of `docxbuild.py` and `webpdf.py`.