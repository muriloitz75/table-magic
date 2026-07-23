"""Parser posicional do "Relatorio Sintetico de Processos" (Prefeitura de
Imperatriz / SEFAZGO) — o relatorio real que a aplicacao converte.

Esse relatorio nao tem linhas de grade: cada processo e uma faixa horizontal
de texto ancorada a esquerda pelo numero "NNNNNN - AAAA", e as colunas sao
definidas apenas pela posicao x das palavras. A extracao generica (pdfplumber
find_tables) colapsa cada coluna numa unica celula gigante, entao esse relatorio
precisa de reconstrucao por coordenadas.

O parser espelha o modelo (modelo/Setor21_20260720.xlsx):
  - reconstroi 1 linha por processo (juntando o nome do contribuinte que quebra)
  - separa "Cadastro Abertura" (duas datas iguais) numa unica "Data Abertura"
  - deriva as colunas "Mes" (abrev. pt-BR) e "Ano" a partir da data
  - remove a coluna "Titulo"
  - reordena para a ordem do modelo
  - ordena por "Tipo de Solicitacao" para o agrupamento/subtotais do builder
"""

import io
import re
from collections import defaultdict

import pdfplumber

# Assinatura do cabecalho (tokens que identificam este relatorio)
HEADER_SIGNATURE = ("Proc.", "Contribuinte", "Abertura", "Solicita", "Situa")

# Fronteiras de coluna (x): corredores de espaco em branco entre colunas,
# medidos no template. Escaladas pela largura da pagina. Uma palavra pertence
# a coluna i se BOUNDARIES[i-1] <= x0 < BOUNDARIES[i].
TEMPLATE_WIDTH = 842.0
BOUNDARIES = [82.0, 211.0, 291.0, 387.0, 484.0, 579.0, 674.0, 770.0]
COL_PROC, COL_CONTRIB, COL_CADASTRO, COL_SETORCAD, COL_SITUACAO, \
    COL_SETORATUAL, COL_TIPO, COL_TITULO, COL_DIAS = range(9)
CONTRIB_START = 82.0  # x minimo do contribuinte (para detectar inicio de registro)

# Colunas de saida, na ordem do modelo (Titulo removido; Mes/Ano derivados)
OUTPUT_HEADERS = [
    "Nº Proc. / Ano",
    "Contribuinte",
    "Data Abertura",
    "Mês",
    "Ano",
    "Situação",
    "Tipo de Solicitação",
    "Setor de Cadastro / Usuário",
    "Setor Atual",
    "Dias Atraso",
]

MONTHS_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]

PROC_RE = re.compile(r"^\d{5,8}$")
YEAR_RE = re.compile(r"^\d{4}$")
DATE_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")
# Contribuinte colado na data: "PARTICIPACOE13/07/2026" ou ruido "G15/07/2026"
GLUED_DATE_RE = re.compile(r"^(.*?)(\d{2}/\d{2}/\d{4}.*)$")
FOOTER_RE = re.compile(r"^(PREFEITURA DE IMPERATRIZ|TOTAL DE)\b")


def _is_report(words) -> bool:
    top_texts = " ".join(w["text"] for w in words if w["top"] < 130)
    return all(sig in top_texts for sig in HEADER_SIGNATURE)


def _column_boundaries(page_width: float) -> list[float]:
    scale = page_width / TEMPLATE_WIDTH
    return [b * scale for b in BOUNDARIES]


def _assign_column(x0: float, bounds: list[float]) -> int:
    for i, b in enumerate(bounds):
        if x0 < b:
            return i
    return len(bounds)  # ultima coluna (Dias Atraso)


def _split_glued_words(words: list[dict]) -> list[dict]:
    """Separa tokens onde o texto de uma coluna colou na data da coluna seguinte
    (ex.: 'PARTICIPACOE13/07/2026'). Divide em dois pseudo-tokens com x aproximado
    por interpolacao. Prefixo de 1 letra antes da data e ruido de render e some."""
    out = []
    for w in words:
        text = w["text"]
        m = GLUED_DATE_RE.match(text)
        if m and m.group(1) and not m.group(1).endswith("/"):
            prefix, rest = m.group(1), m.group(2)
            width = w["x1"] - w["x0"]
            split_x = w["x0"] + width * (len(prefix) / len(text))
            if len(prefix.strip()) > 1:  # cauda real do contribuinte
                out.append({**w, "text": prefix, "x1": split_x})
            # prefixo <= 1 char: ruido de render, descartado
            out.append({**w, "text": rest, "x0": split_x})
        else:
            out.append({**w})
    return out


def _cluster_rows(words: list[dict], tol: float = 4.0) -> list[list[dict]]:
    """Agrupa palavras em linhas visuais por proximidade de `top`."""
    rows: list[tuple[float, list[dict]]] = []
    for w in sorted(words, key=lambda k: (k["top"], k["x0"])):
        placed = False
        for i, (top, bucket) in enumerate(rows):
            if abs(w["top"] - top) <= tol:
                bucket.append(w)
                placed = True
                break
        if not placed:
            rows.append((w["top"], [w]))
    return [sorted(bucket, key=lambda k: k["x0"]) for _, bucket in rows]


def _clean_date(raw: str) -> str:
    m = DATE_RE.search(raw)
    return m.group(1) if m else ""


def _month_year(date_str: str) -> tuple[str, int | str]:
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", date_str)
    if not m:
        return "", ""
    month = int(m.group(2))
    year = int(m.group(3))
    return (MONTHS_PT[month - 1] if 1 <= month <= 12 else ""), year


def _finalize_record(cols: dict[int, list[str]]) -> list[str] | None:
    def get(idx: int) -> str:
        return " ".join(cols.get(idx, [])).strip()

    proc = get(COL_PROC)
    if not proc:
        return None
    proc = re.sub(r"\s*-\s*", " - ", proc)

    date = _clean_date(get(COL_CADASTRO))
    month, year = _month_year(date)

    # Contribuinte: remove a data que cola no fim de nomes longos e "/" residual
    contrib = DATE_RE.sub("", get(COL_CONTRIB))
    contrib = re.sub(r"\s*/\s*$", "", contrib)
    contrib = re.sub(r"\s{2,}", " ", contrib).strip()

    # Tipo de Solicitacao: o texto transborda para a zona da coluna "Titulo"
    # (que o modelo funde de volta ao tipo), entao juntamos as duas.
    tipo = re.sub(r"\s{2,}", " ", f"{get(COL_TIPO)} {get(COL_TITULO)}").strip()

    dias = get(COL_DIAS).strip() or "-"

    return [
        proc,
        contrib,
        date,
        month,
        str(year) if year != "" else "",
        get(COL_SITUACAO),
        tipo,
        get(COL_SETORCAD),
        get(COL_SETORATUAL),
        dias,
    ]


def parse_report(pdf_bytes: bytes) -> dict | None:
    """Retorna {tableName, headers, rows} se o PDF for o Relatorio Sintetico de
    Processos; caso contrario None (para o orchestrator cair na extracao generica)."""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            return None
        first_words = pdf.pages[0].extract_words()
        if not _is_report(first_words):
            return None

        records: list[list[str]] = []

        for page in pdf.pages:
            words = page.extract_words()
            bounds = _column_boundaries(page.width)
            rows = _cluster_rows(_split_glued_words(words))

            current: dict[int, list[str]] | None = None
            for row in rows:
                line_text = " ".join(w["text"] for w in row)
                # Pula cabecalho e rodape repetidos
                if any(sig in line_text for sig in ("N° Proc", "Nº Proc", "N Proc")) and "Contribuinte" in line_text:
                    continue
                if FOOTER_RE.match(line_text.strip()):
                    continue
                if not row:
                    continue

                first = row[0]
                scale = page.width / TEMPLATE_WIDTH
                is_new = bool(PROC_RE.match(first["text"])) and first["x0"] < CONTRIB_START * scale

                if is_new:
                    if current is not None:
                        rec = _finalize_record(current)
                        if rec:
                            records.append(rec)
                    current = defaultdict(list)
                    for w in row:
                        current[_assign_column(w["x0"], bounds)].append(w["text"])
                elif current is not None:
                    # Continuacao: nome do contribuinte que quebrou de linha
                    for w in row:
                        col = _assign_column(w["x0"], bounds)
                        if col in (COL_PROC, COL_CONTRIB):
                            current[COL_CONTRIB].append(w["text"])

            if current is not None:
                rec = _finalize_record(current)
                if rec:
                    records.append(rec)
                current = None

    if not records:
        return None

    # Ordena por Tipo de Solicitacao (idx 6) para agrupamento/subtotais contiguos
    tipo_idx = 6
    records.sort(key=lambda r: r[tipo_idx].lower())

    return {
        "tableName": "Relatório Sintético de Processos",
        "headers": list(OUTPUT_HEADERS),
        "rows": records,
        # Relatorio de contagem: os totais devem apenas contar processos por
        # grupo (coluna Quantidade), sem somar Ano/Dias Atraso.
        "count_only": True,
    }
