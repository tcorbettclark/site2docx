"""PDF -> Word: every PDF page becomes a bordered image on its own Word page,
with a Comments column beside it. Moved here from cli.py for separation of
concerns; see cli.py for the end-to-end URL -> Word flow."""

import sys
from io import BytesIO
from pathlib import Path

import pymupdf as fitz  # PyMuPDF
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# --- Rendering ---------------------------------------------------------------
DPI = 150
IMAGE_BORDER_PT = 1.0

# --- Word page layout (A4 portrait), lengths in cm ---------------------------
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN = 1.0
IMAGE_COL = 13.6                                # left (image) column width
COMMENTS_COL = PAGE_W - 2 * MARGIN - IMAGE_COL  # right column width (5.4)
CELL_PAD = 0.15                                  # table cell inner margin
IMAGE_W = IMAGE_COL - 2 * CELL_PAD                # image width inside the cell

# Schema order for w:tblPr children, used to insert elements in valid sequence.
_TBLPR_ORDER = (
    "w:tblStyle", "w:tblpPr", "w:tblW", "w:jc", "w:tblCellSpacing",
    "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
    "w:tblLook", "w:tblCaption", "w:tblDescription",
)


def _log(msg):
    print(msg, file=sys.stderr)


def _tblpr_insert(tblPr, element):
    """Insert element into w:tblPr at its schema-correct position."""
    order = {qn(tag): i for i, tag in enumerate(_TBLPR_ORDER)}
    pos = order[element.tag]
    for child in tblPr:
        if child.tag in order and order[child.tag] > pos:
            child.addprevious(element)
            return
    tblPr.append(element)


def _set_fixed_layout(table):
    """Force fixed table layout so column widths are honoured exactly."""
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _tblpr_insert(tblPr, layout)
    table.autofit = False


def _set_cell_margins(table, pad_cm):
    """Set uniform inner cell margins (w:dxa = twentieths of a point)."""
    dxa = str(int(round(pad_cm * 567)))
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), "0" if side in ("top", "bottom") else dxa)
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _tblpr_insert(table._tbl.tblPr, mar)


def _keep_row_together(row):
    """Prevent the row from splitting across Word pages."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


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


def _heading_with_rule(paragraph, text):
    """Bold heading with a thin rule underneath (e.g. the Comments header)."""
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(10)
    pPr = paragraph._p.get_or_add_pPr()  # fresh paragraph: pPr is empty
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")  # half-point units: 0.5pt
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_page_table(doc, png_bytes):
    """One two-column page: image left, Comments column right."""
    table = doc.add_table(rows=1, cols=2)
    _set_fixed_layout(table)
    _set_cell_margins(table, CELL_PAD)
    _keep_row_together(table.rows[0])

    table.columns[0].width = Cm(IMAGE_COL)
    table.columns[1].width = Cm(COMMENTS_COL)

    img_cell, com_cell = table.rows[0].cells
    for cell, width in ((img_cell, IMAGE_COL), (com_cell, COMMENTS_COL)):
        cell.width = Cm(width)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Left: the page image, centred, with its 1pt black border baked in.
    para = img_cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(BytesIO(png_bytes), width=Cm(IMAGE_W))

    # Right: Comments heading + blank space for the reviewer.
    _heading_with_rule(com_cell.paragraphs[0], "Comments")
    com_cell.add_paragraph()
    com_cell.add_paragraph()

    return table


def convert(pdf_path, docx_path, dpi=DPI):
    pdf_path, docx_path = Path(pdf_path), Path(docx_path)
    if not pdf_path.is_file():
        raise SystemExit(f"error: no such file: {pdf_path}")

    with fitz.open(pdf_path) as pdf:
        n_pages = pdf.page_count
        _log(f"  pdf: {n_pages} pages, rendering at {dpi} DPI")

        doc = Document()
        section = doc.sections[0]
        section.page_width, section.page_height = Cm(PAGE_W), Cm(PAGE_H)
        section.top_margin = section.bottom_margin = Cm(MARGIN)
        section.left_margin = section.right_margin = Cm(MARGIN)

        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for i, page in enumerate(pdf, start=1):
            _add_page_table(doc, _render_page_png(page, matrix))
            if i < n_pages:
                br = doc.add_paragraph()
                br.paragraph_format.space_before = Pt(0)
                br.paragraph_format.space_after = Pt(0)
                run = br.add_run()
                run.font.size = Pt(2)
                run.add_break(WD_BREAK.PAGE)
            if i % 10 == 0 or i == n_pages:
                _log(f"  docx: {i}/{n_pages} pages built")
        _log("")

    # Word expects a paragraph after a table that ends the document.
    doc.add_paragraph()
    doc.save(docx_path)
    size_mb = docx_path.stat().st_size / 1_000_000
    _log(f"  wrote {docx_path} ({size_mb:.1f} MB)")
    return docx_path