"""PDF text extraction using opendataloader-pdf (primary) or Tesseract (fallback)."""

import logging
import tempfile
from pathlib import Path

from config import (
    OCR_DPI,
    OCR_LANG,
    POPPLER_PATH,
    TESSERACT_CMD,
    USE_OPENDATALOADER,
    OPENDATALOADER_HYBRID,
    OPENDATALOADER_FORCE_OCR,
)

log = logging.getLogger(__name__)


def _extract_opendataloader(pdf_path: str) -> str:
    """Extract text using opendataloader-pdf."""
    import opendataloader_pdf

    with tempfile.TemporaryDirectory() as tmpdir:
        kwargs = {
            "input_path": [pdf_path],
            "output_dir": tmpdir,
            "format": "text",
        }
        if OPENDATALOADER_HYBRID:
            kwargs["hybrid"] = True
        if OPENDATALOADER_FORCE_OCR:
            kwargs["force_ocr"] = True
            kwargs["ocr_lang"] = OCR_LANG[:2]  # "jpn" -> "ja" for opendataloader

        opendataloader_pdf.convert(**kwargs)

        stem = Path(pdf_path).stem
        for pattern in [f"{stem}.txt", f"{stem}.md", "*.txt", "*.md"]:
            files = list(Path(tmpdir).glob(pattern))
            if files:
                return files[0].read_text(encoding="utf-8")

    raise RuntimeError(f"opendataloader-pdf produced no output for {pdf_path}")


def _extract_tesseract(pdf_path: str, dpi: int, lang: str) -> str:
    """Extract text using Tesseract OCR (legacy fallback)."""
    from pdf2image import convert_from_path
    import pytesseract

    if TESSERACT_CMD:
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

    kwargs = {"dpi": dpi}
    if POPPLER_PATH:
        kwargs["poppler_path"] = POPPLER_PATH
    images = convert_from_path(pdf_path, **kwargs)
    pages = [pytesseract.image_to_string(img, lang=lang) for img in images]
    return "\n".join(pages)


def ocr_pdf(pdf_path: str, dpi: int = OCR_DPI, lang: str = OCR_LANG) -> str:
    """Extract text from a PDF file.

    Uses opendataloader-pdf when available and enabled (USE_OPENDATALOADER=true),
    falls back to Tesseract OCR otherwise.
    """
    if USE_OPENDATALOADER:
        try:
            text = _extract_opendataloader(pdf_path)
            log.info("Extracted text via opendataloader-pdf (%d chars)", len(text))
            return text
        except Exception:
            log.warning(
                "opendataloader-pdf failed, falling back to Tesseract",
                exc_info=True,
            )

    text = _extract_tesseract(pdf_path, dpi, lang)
    log.info("Extracted text via Tesseract (%d chars)", len(text))
    return text
