"""Classifica cada pagina de um PDF como 'digital' (tem camada de texto) ou
'scanned' (sem texto extraivel, provavelmente uma imagem escaneada).
"""

from dataclasses import dataclass

import fitz  # PyMuPDF

MIN_CHARS_PER_PAGE = 20


@dataclass
class PageClassification:
    index: int
    kind: str  # "digital" | "scanned"


def classify_pages(pdf_bytes: bytes) -> list[PageClassification]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        results = []
        for i, page in enumerate(doc):
            text = page.get_text("text").strip()
            kind = "digital" if len(text) >= MIN_CHARS_PER_PAGE else "scanned"
            results.append(PageClassification(index=i, kind=kind))
        return results
    finally:
        doc.close()


def page_count(pdf_bytes: bytes) -> int:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return doc.page_count
    finally:
        doc.close()
