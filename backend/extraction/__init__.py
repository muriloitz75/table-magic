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

    for page in pages:
        if page.kind == "digital":
            found = extract_digital_tables(pdf_bytes, [page.index])
            if found:
                tables.extend(found)
                methods_used.add("digital")
        else:
            try:
                found = extract_ocr_tables(pdf_bytes, [page.index])
                if found:
                    tables.extend(found)
                    methods_used.add("ocr")
            except TesseractUnavailableError as exc:
                warning = str(exc)

    if not methods_used:
        method = "none"
    elif len(methods_used) == 1:
        method = methods_used.pop()
    else:
        method = "mixed"

    return {"tables": tables, "method": method, "warning": warning}
