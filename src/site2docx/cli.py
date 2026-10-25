"""site2docx — one command: URL or PDF in, reviewable Word document out.

Usage:
    site2docx URL_OR_PDF [OUTPUT_DOCX]

- If the input is an http(s) URL, the page is rendered in headless system
  Chrome (collapsed sections expanded), printed to a temporary PDF, and the
  PDF is discarded after conversion.
- If the input is a .pdf file, only the PDF -> Word conversion runs.
"""

import re
import sys
import tempfile
import time
from pathlib import Path

from . import docxbuild, webpdf


def _log(msg):
    print(msg, file=sys.stderr)


def _is_url(s):
    return s.startswith(("http://", "https://"))


def _safe_name(title):
    """Turn a page title into a sensible filename (Chrome-style, ':' -> '_')."""
    name = title.strip() or "document"
    name = re.sub(r"[/:]", "_", name)
    return name[:180].rstrip(" .") + ".docx"


def _from_url(url, output=None):
    with tempfile.TemporaryDirectory(prefix="site2docx-") as tmp:
        pdf_path = Path(tmp) / "page.pdf"
        title, stats = webpdf.render(url, pdf_path)
        out = Path(output) if output else Path.cwd() / _safe_name(title)
        docxbuild.convert(pdf_path, out)
    _log(f"intermediate PDF discarded (temporary)")
    return out


def _from_pdf(pdf_arg, output=None):
    pdf_path = Path(pdf_arg)
    out = Path(output) if output else pdf_path.with_suffix(".docx")
    docxbuild.convert(pdf_path, out)
    return out


def main(argv=None):
    args = sys.argv[1:] if argv is None else argv
    if not 1 <= len(args) <= 2:
        raise SystemExit("usage: site2docx URL_OR_PDF [OUTPUT_DOCX]")

    t0 = time.monotonic()
    source = args[0]
    if _is_url(source):
        out = _from_url(source, args[1] if len(args) == 2 else None)
    else:
        out = _from_pdf(source, args[1] if len(args) == 2 else None)

    _log(f"done: {out} ({time.monotonic() - t0:.1f}s total)")


if __name__ == "__main__":
    main()