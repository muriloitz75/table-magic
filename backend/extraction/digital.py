"""Extracao de tabelas de paginas de PDF digitais (com camada de texto).

Cascata: pdfplumber (estrategia de linhas, depois texto) -> PyMuPDF find_tables()
como fallback quando pdfplumber nao encontra nada aproveitavel na pagina.
"""

import io

import fitz  # PyMuPDF
import pdfplumber


def _clean(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", " ").strip()


def _rows_have_content(rows: list[list[str]]) -> bool:
    return any(any(cell != "" for cell in row) for row in rows)


def _normalize_raw_table(raw_table: list[list]) -> dict | None:
    if not raw_table or len(raw_table) < 2:
        return None
    headers = [_clean(h) for h in raw_table[0]]
    rows = [[_clean(c) for c in row] for row in raw_table[1:]]
    rows = [r for r in rows if _rows_have_content(r)]
    if not any(h != "" for h in headers) or not rows:
        return None
    return {"tableName": None, "headers": headers, "rows": rows}


def _extract_with_pdfplumber(page) -> list[dict]:
    results = []
    raw_tables = page.extract_tables()
    if not raw_tables:
        raw_tables = page.extract_tables(
            table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
            }
        )
    for raw in raw_tables or []:
        normalized = _normalize_raw_table(raw)
        if normalized:
            results.append(normalized)
    return results


def _extract_with_pymupdf(doc: fitz.Document, page_index: int) -> list[dict]:
    page = doc[page_index]
    finder = page.find_tables()
    results = []
    for table in finder.tables:
        normalized = _normalize_raw_table(table.extract())
        if normalized:
            results.append(normalized)
    return results


def extract_digital_tables(pdf_bytes: bytes, page_indices: list[int]) -> list[dict]:
    """Extrai tabelas das paginas indicadas (indices 0-based) de um PDF digital."""
    tables: list[dict] = []
    fitz_doc: fitz.Document | None = None
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for idx in page_indices:
                if idx >= len(pdf.pages):
                    continue
                page = pdf.pages[idx]
                try:
                    found = _extract_with_pdfplumber(page)
                    if not found:
                        if fitz_doc is None:
                            fitz_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                        found = _extract_with_pymupdf(fitz_doc, idx)
                    tables.extend(found)
                finally:
                    # Libera cache interno do pdfminer/pdfplumber para evitar OOM em PDFs com centenas de paginas.
                    page.flush_cache()
    finally:
        if fitz_doc is not None:
            fitz_doc.close()
    return tables
