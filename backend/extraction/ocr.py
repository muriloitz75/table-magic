"""Extracao de tabelas de paginas escaneadas via OCR (Tesseract).

Abordagem heuristica: rasteriza a pagina em alta resolucao, roda o Tesseract
com dados de posicao por palavra, agrupa palavras em linhas (usando a propria
segmentacao de linha do Tesseract) e depois agrupa as posicoes horizontais de
todas as palavras da pagina em "colunas" por proximidade. E mais fragil que a
extracao digital -- paginas resultantes sao marcadas como extraidas via OCR
para o usuario conferir os valores.
"""

import io

import fitz  # PyMuPDF
import numpy as np
import pytesseract
from PIL import Image
from pytesseract import Output

ZOOM = 300 / 72  # renderiza a ~300 DPI
MIN_CONFIDENCE = 40
COLUMN_GAP_RATIO = 0.02  # % da largura da pagina usada como folga entre colunas


class TesseractUnavailableError(RuntimeError):
    pass


def _ensure_tesseract():
    try:
        pytesseract.get_tesseract_version()
    except Exception as exc:  # binario nao encontrado no PATH
        raise TesseractUnavailableError(
            "Tesseract OCR nao esta instalado ou nao foi encontrado no PATH. "
            "Instale o Tesseract (com o pacote de idioma portugues) para "
            "processar PDFs escaneados."
        ) from exc


def _rasterize_page(pdf_bytes: bytes, page_index: int) -> Image.Image:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = doc[page_index]
        matrix = fitz.Matrix(ZOOM, ZOOM)
        pix = page.get_pixmap(matrix=matrix)
        return Image.open(io.BytesIO(pix.tobytes("png")))
    finally:
        doc.close()


def _cluster_columns(lefts: list[int], gap: float) -> list[float]:
    """Agrupa posicoes x proximas em centros de coluna, ordenados."""
    if not lefts:
        return []
    ordered = sorted(lefts)
    clusters: list[list[int]] = [[ordered[0]]]
    for x in ordered[1:]:
        if x - clusters[-1][-1] <= gap:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    return [sum(c) / len(c) for c in clusters]


def _nearest_column(x: int, centers: list[float]) -> int:
    return min(range(len(centers)), key=lambda i: abs(centers[i] - x))


def _extract_lines(data: dict) -> list[list[dict]]:
    lines: dict[tuple, list[dict]] = {}
    n = len(data["text"])
    for i in range(n):
        text = data["text"][i].strip()
        conf = float(data["conf"][i]) if data["conf"][i] not in ("-1", "") else -1
        if not text or conf < MIN_CONFIDENCE:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append(
            {"text": text, "left": data["left"][i], "top": data["top"][i]}
        )
    ordered_keys = sorted(lines.keys())
    return [sorted(lines[k], key=lambda w: w["left"]) for k in ordered_keys]


def _build_grid(lines: list[list[dict]], page_width: int) -> list[list[str]] | None:
    if len(lines) < 2:
        return None

    all_lefts = [w["left"] for line in lines for w in line]
    gap = page_width * COLUMN_GAP_RATIO
    centers = _cluster_columns(all_lefts, gap)
    if len(centers) < 2:
        return None

    grid = []
    for line in lines:
        row = [""] * len(centers)
        for word in line:
            col = _nearest_column(word["left"], centers)
            row[col] = f"{row[col]} {word['text']}".strip() if row[col] else word["text"]
        grid.append(row)
    return grid


def extract_ocr_tables(pdf_bytes: bytes, page_indices: list[int]) -> list[dict]:
    """Extrai tabelas das paginas escaneadas indicadas via OCR."""
    _ensure_tesseract()
    tables: list[dict] = []

    for idx in page_indices:
        image = _rasterize_page(pdf_bytes, idx)
        data = pytesseract.image_to_data(
            image, lang="por+eng", config="--psm 6", output_type=Output.DICT
        )
        lines = _extract_lines(data)
        grid = _build_grid(lines, page_width=np.array(image).shape[1])
        if not grid:
            continue

        headers = grid[0]
        rows = [r for r in grid[1:] if any(c != "" for c in r)]
        if not rows:
            continue

        tables.append({"tableName": None, "headers": headers, "rows": rows})

    return tables
