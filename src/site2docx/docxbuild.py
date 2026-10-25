"""PDF -> Word: every PDF page becomes one landscape Word page with the page
image in the left column and the page's extracted text in the right column.

Numbered badges are baked into the image at each item's true position — one
badge per text block or list entry, and one per non-empty cell of a table —
and each corresponding paragraph in the right column carries
the same number. Reviewers use normal Word comments and track changes on the
right-column text.

Ordering matters: text is extracted from the PDF *before* badges are drawn,
because badge numbers are PDF text and would otherwise pollute the
extraction.
"""

import re
import sys
from io import BytesIO
from pathlib import Path

import pymupdf as fitz  # PyMuPDF
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
# --- Rendering ---------------------------------------------------------------
DPI = 150
IMAGE_BORDER_PT = 1.0

# --- Word page layout (A4 landscape, two columns), lengths in cm --------------
PAGE_W, PAGE_H = 29.7, 21.0  # A4 landscape
MARGIN = 1.0
GAP = 0.5                    # gutter between the two columns
IMAGE_COL = 13.6             # left column: the page image
IMAGE_W = 13.2               # image width; height 18.66cm fits the 19cm column
TEXT_COL = PAGE_W - 2 * MARGIN - IMAGE_COL - GAP  # right column: source text

# --- Right-column text --------------------------------------------------------
TEXT_PT = 10                 # body text size on the right
TEXT_SPACE_AFTER_PT = 3      # paragraph spacing, kept tight for density

# --- Numbered badges (drawn on the source page, in points) -------------------
BADGE_R_PT = 10.0            # badge circle radius
BADGE_PAD_PT = 4.0           # gap between badge and the block's left edge
BADGE_FONT_PT = 10.0         # badge number size (Helvetica-Bold)
BADGE_MIN_X_PT = 12.0        # never let the badge centre cross the page edge
# Table-cell badges: smaller, centred on the cell's top-left grid
# intersection so they cover border lines rather than content.
TABLE_BADGE_R_PT = 8.0
TABLE_BADGE_FONT_PT = 8.0

# --- Chrome print cruft to strip from extraction (header/footer bands) -------
# Chrome's print margins (see webpdf.PRINT_MARGINS) are 0.4in top / 0.56in
# bottom; header/footer text lives inside those bands. Blocks fully inside a
# band are dropped. Tolerance added for slightly dipping headers.
CRUFT_TOP_PT = 30.0
CRUFT_BOTTOM_PT = 42.0

# --- List items --------------------------------------------------------------
# A new numbered item starts at any line beginning with a bullet or an
# enumeration marker, so each bullet / numbered entry gets its own badge.
# Monospaced text is exempt: '-' and '1.' at line start are usually code.
_BULLETS = "\u2022\u25e6\u25aa\u2023\u25cf\u25cb\u00b7*-"
_LIST_MARKER_RE = re.compile(
    rf"^(?:[{_BULLETS}]|\d{{1,3}}[.)]|[a-z][.)]|[ivx]+[.)]|[A-Z][.)])\s+"
)
_MONO_FONT_HINTS = ("mono", "courier", "consol", "menlo")
# Graphic bullets: some sites (Ada CS included) draw list markers as small
# vector shapes instead of text glyphs, so they never appear in the text
# layer. A line counts as a list item start if such a shape sits just to
# its left, vertically aligned with it.
_MARKER_MAX_DIM_PT = 9.0
_MARKER_MAX_LEFT_PT = 36.0


def _is_monospace(font_name):
    return any(hint in font_name.lower() for hint in _MONO_FONT_HINTS)


def _marker_drawings(page):
    """Small filled shapes that look like graphic list bullets."""
    markers = []
    for d in page.get_drawings():
        r = d["rect"]
        w, h = r.width, r.height
        if d["type"] not in ("f", "fs"):
            continue
        if not (1.5 <= w <= _MARKER_MAX_DIM_PT and 1.5 <= h <= _MARKER_MAX_DIM_PT):
            continue
        if not 0.4 <= w / max(h, 1e-6) <= 2.5:  # roughly square/circular
            continue
        markers.append(fitz.Rect(r))
    return markers


def _has_graphic_marker(bbox, markers):
    """True if a graphic bullet sits just left of this line, aligned with it."""
    for m in markers:
        if bbox.x0 - _MARKER_MAX_LEFT_PT < m.x1 <= bbox.x0:
            m_y_mid = (m.y0 + m.y1) / 2
            if bbox.y0 - 2 <= m_y_mid <= bbox.y1 + 2:
                return True
    return False


def _log(msg):
    print(msg, file=sys.stderr)


def _set_current_compat_mode(doc):
    """Open in Word 2013+ mode, not the template's compatibility mode."""
    compat = doc.settings.element.find(qn("w:compat"))
    for cs in compat.findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")


def _set_print_layout(doc):
    """Open the document in Print Layout view, not Web/Reading Layout."""
    settings = doc.settings.element
    view = OxmlElement("w:view")
    view.set(qn("w:val"), "print")
    zoom = settings.find(qn("w:zoom"))
    if zoom is not None:
        zoom.addprevious(view)
    else:
        settings.insert(0, view)


def _setup_section(section):
    """Landscape A4 with two unequal columns: image left, text right.

    python-docx has no API for unequal section columns, so w:cols is built
    by hand and swapped in for the template's, in its schema-correct
    position (w:cols comes before w:docGrid in w:sectPr).
    """
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Cm(PAGE_W), Cm(PAGE_H)
    section.top_margin = section.bottom_margin = Cm(MARGIN)
    section.left_margin = section.right_margin = Cm(MARGIN)

    sectPr = section._sectPr
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), str(int(GAP * 567)))
    c1 = OxmlElement("w:col")
    c1.set(qn("w:w"), str(int(IMAGE_COL * 567)))
    c1.set(qn("w:space"), str(int(GAP * 567)))
    c2 = OxmlElement("w:col")
    c2.set(qn("w:w"), str(int(TEXT_COL * 567)))
    cols.append(c1)
    cols.append(c2)

    old = sectPr.find(qn("w:cols"))
    if old is not None:
        old.addprevious(cols)
        sectPr.remove(old)
    else:
        anchor = sectPr.find(qn("w:docGrid"))
        if anchor is not None:
            anchor.addprevious(cols)
        else:
            sectPr.append(cols)


def _overlap_fraction(rect, region):
    """Fraction of rect's area that lies inside region."""
    inter = fitz.Rect(rect) & fitz.Rect(region)
    if not inter.is_valid or inter.is_empty:
        return 0.0
    return inter.get_area() / max(fitz.Rect(rect).get_area(), 1e-9)


def _page_items(page):
    """Extract the page's text as (text, rect, kind) triples, reading order.

    kind is "block" for ordinary paragraphs and "cell" for table cells:
    every non-empty cell of every detected table is numbered individually.
    Text blocks inside a detected table are folded into their cell rather
    than numbered twice. Chrome header/footer bands are stripped.
    """
    page_h = page.rect.height
    try:
        tables = list(page.find_tables().tables)
    except Exception:
        tables = []
    regions = [fitz.Rect(t.bbox) for t in tables]
    markers = _marker_drawings(page)

    items = []
    for block in page.get_text("dict", sort=True)["blocks"]:
        if block.get("type") != 0:
            continue
        rect = fitz.Rect(block["bbox"])
        if rect.y1 <= CRUFT_TOP_PT or rect.y0 >= page_h - CRUFT_BOTTOM_PT:
            continue  # header/footer band
        if any(_overlap_fraction(rect, r) > 0.5 for r in regions):
            continue  # numbered as part of its cell instead
        lines = []
        for line in block["lines"]:
            text = "".join(span["text"] for span in line["spans"]).strip()
            if text:
                mono = any(_is_monospace(span["font"]) for span in line["spans"])
                lines.append((text, fitz.Rect(line["bbox"]), mono))
        if not lines:
            continue
        # One item per list entry: a new item starts at any line beginning
        # with a bullet or enumeration marker (but not in monospaced code).
        # Bullets can be text glyphs or small vector shapes (see
        # _marker_drawings), so both are checked.
        segs = []
        for text, bbox, mono in lines:
            is_item_start = (not mono) and (
                _LIST_MARKER_RE.match(text) or _has_graphic_marker(bbox, markers)
            )
            if segs and is_item_start:
                segs.append([])
            if not segs:
                segs.append([])
            segs[-1].append((text, bbox))
        for seg in segs:
            # Reflow: join lines, mending words hyphenated across a break.
            out = seg[0][0]
            for text, _bbox in seg[1:]:
                if out.endswith("-"):
                    out = out[:-1] + text
                else:
                    out += " " + text
            out = out.strip()
            if not out:
                continue
            rect = fitz.Rect(seg[0][1])
            for _text, bbox in seg[1:]:
                rect |= bbox  # union: wrapped lines belong to the same item
            items.append((out, rect, "block"))

    # Every cell of every detected table, row-major; dedupe merged cells.
    seen = set()
    for t in tables:
        for row in t.rows:
            for cell in row.cells:
                if cell is None:
                    continue
                rect = fitz.Rect(cell)
                if rect.y1 <= CRUFT_TOP_PT or rect.y0 >= page_h - CRUFT_BOTTOM_PT:
                    continue
                key = (round(rect.x0, 1), round(rect.y0, 1),
                       round(rect.x1, 1), round(rect.y1, 1))
                if key in seen:
                    continue
                seen.add(key)
                text = " ".join(page.get_text("text", clip=rect).split())
                if text:
                    items.append((text, rect, "cell"))

    items.sort(key=lambda it: (it[1].y0, it[1].x0))
    return items


def _draw_badge(page, n, rect, kind):
    """Bake badge number n into the page.

    Block: just left of the block's first character at the block's own y,
    with a semi-transparent overlay on the block's top-left corner when
    there is no room to the left.

    Cell: smaller, pulled down and left from the cell's top-right grid
    intersection by about a third of the badge size, so the badge sits
    clearly inside its own cell (over border and padding, not content).
    """
    if kind == "cell":
        r, font_pt = TABLE_BADGE_R_PT, TABLE_BADGE_FONT_PT
        inset = 2 * r / 3.0
        cx, cy = rect.x1 - inset, rect.y0 + inset
        overlay = False
    else:
        r, font_pt = BADGE_R_PT, BADGE_FONT_PT
        cx = rect.x0 - r - BADGE_PAD_PT
        overlay = cx < BADGE_MIN_X_PT
        if overlay:
            cx = rect.x0 + r
        cy = rect.y0 + r

    shape = page.new_shape()
    shape.draw_oval(fitz.Rect(cx - r, cy - r, cx + r, cy + r))
    if overlay:
        shape.finish(color=(0, 0, 0), fill=(1, 1, 1), fill_opacity=0.85, width=1.0)
    else:
        shape.finish(color=(0, 0, 0), fill=(1, 1, 1), width=1.0)
    shape.commit()

    label = str(n)
    w = fitz.get_text_length(label, fontname="hebo", fontsize=font_pt)
    page.insert_text((cx - w / 2.0, cy + 0.36 * font_pt), label,
                     fontsize=font_pt, fontname="hebo", color=(0, 0, 0))


def _render_page_png(page, matrix):
    """Render a PDF page to PNG, with a black hairline border drawn just
    inside the page edge and baked into the pixels.

    A border drawn via DrawingML (a:ln) is only advisory — some Word renderers
    drop edges (notably the top) when an image sits flush at the top of its
    line box. Baking it into the image guarantees all four sides appear.
    """
    shape = page.new_shape()
    half = IMAGE_BORDER_PT / 2.0
    inset = page.rect + (half, half, -half, -half)  # keep the full line visible
    shape.draw_rect(inset)
    shape.finish(color=(0, 0, 0), width=IMAGE_BORDER_PT)
    shape.commit()
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return pix.tobytes("png")


def _add_tiny_break_paragraph(doc, break_type):
    """Near-zero-height paragraph whose only job is to carry a break.

    2pt font so the paragraph still fits under the 18.66cm image in
    column 1 (column break) or adds nothing before the next page (page
    break). Word's performance depends on there being a single section, so
    pages are separated by page breaks, not section breaks: a page break in
    a multi-column section lands at the top of column 1 on the next page.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.font.size = Pt(2)
    run.add_break(break_type)


def _add_page(doc, page_index, png_bytes, items, is_last_page):
    """One landscape Word page: image in column 1, numbered text in column 2."""
    # Column 1: the page image with badges and border baked in.
    img_p = doc.add_paragraph()
    img_p.paragraph_format.space_before = img_p.paragraph_format.space_after = Pt(0)
    img_p.add_run().add_picture(BytesIO(png_bytes), width=Cm(IMAGE_W))

    # Tiny column break so the text starts at the top of column 2.
    _add_tiny_break_paragraph(doc, WD_BREAK.COLUMN)

    # Column 2: the page's text, reflowed, numbered to match the badges.
    for n, (text, _rect, _kind) in enumerate(items, 1):
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(TEXT_SPACE_AFTER_PT)
        num = para.add_run(f"{n} ")
        num.bold = True
        num.font.size = Pt(TEXT_PT)
        body = para.add_run(text)
        body.font.size = Pt(TEXT_PT)

    # Next source page starts at the top of column 1 on the next Word page.
    if not is_last_page:
        _add_tiny_break_paragraph(doc, WD_BREAK.PAGE)


def convert(pdf_path, docx_path, dpi=DPI):
    pdf_path, docx_path = Path(pdf_path), Path(docx_path)
    if not pdf_path.is_file():
        raise SystemExit(f"error: no such file: {pdf_path}")

    with fitz.open(pdf_path) as pdf:
        n_pages = pdf.page_count
        _log(f"  pdf: {n_pages} pages, rendering at {dpi} DPI")

        doc = Document()
        _set_current_compat_mode(doc)
        _set_print_layout(doc)
        _setup_section(doc.sections[0])

        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for i, page in enumerate(pdf):
            items = _page_items(page)             # extract text first...
            for n, (_t, rect, kind) in enumerate(items, 1):
                _draw_badge(page, n, rect, kind)  # ...then bake badges in
            png = _render_page_png(page, matrix)
            _add_page(doc, i, png, items, is_last_page=(i == n_pages - 1))
            if (i + 1) % 10 == 0 or (i + 1) == n_pages:
                _log(f"  docx: {i + 1}/{n_pages} pages built")
        _log("")

    doc.save(docx_path)
    size_mb = docx_path.stat().st_size / 1_000_000
    _log(f"  wrote {docx_path} ({size_mb:.1f} MB)")
    return docx_path