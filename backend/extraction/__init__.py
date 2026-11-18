"""Orquestra a extracao de tabelas de um PDF: classifica cada pagina como
digital ou escaneada e despacha para o extrator apropriado, preservando a
ordem das paginas no documento.
"""

from .detect import classify_pages
from .digital import extract_digital_tables
from .ocr import TesseractUnavailableError, extract_ocr_tables
from .report_parser import parse_report


def extract_tables(pdf_bytes: bytes) -> dict:
    # Parser especifico do "Relatorio Sintetico de Processos" (o relatorio real
    # da aplicacao): reconstroi por coordenadas o que a extracao generica de
    # tabelas nao consegue. Se o PDF nao for esse relatorio, retorna None e
    # seguimos para a extracao generica/OCR.
    report = parse_report(pdf_bytes)
    if report is not None:
        return {"tables": [report], "method": "report", "warning": None}

    pages = classify_pages(pdf_bytes)

    tables: list[dict] = []
    methods_used: set[str] = set()
    warning: str | None = None

    i = 0
    while i < len(pages):
        if pages[i].kind == "digital":
            batch_indices: list[int] = []
            while i < len(pages) and pages[i].kind == "digital":
                batch_indices.append(pages[i].index)
                i += 1
            found = extract_digital_tables(pdf_bytes, batch_indices)
            if found:
                tables.extend(found)
                methods_used.add("digital")
        else:
            try:
                found = extract_ocr_tables(pdf_bytes, [pages[i].index])
                if found:
                    tables.extend(found)
                    methods_used.add("ocr")
            except TesseractUnavailableError as exc:
                warning = str(exc)
            i += 1

    if not methods_used:
        method = "none"
    elif len(methods_used) == 1:
        method = methods_used.pop()
    else:
        method = "mixed"

    return {"tables": tables, "method": method, "warning": warning}
