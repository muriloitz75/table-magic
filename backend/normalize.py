"""Conversao segura de valores de celula para tipos numericos.

So converte valores que inequivocamente seguem o formato numerico brasileiro
(milhar com ponto, decimal com virgula) ou um inteiro puro. Qualquer coisa
ambigua (datas "01.02.2026", codigos com zero a esquerda "00123", floats no
formato americano "1.5") permanece como texto para nao corromper dados.
"""

import re
from typing import Union

_INTEGER = re.compile(r"^-?\d+$")
_DECIMAL_ONLY = re.compile(r"^-?\d+,\d+$")
_THOUSANDS_DECIMAL = re.compile(r"^-?\d{1,3}(\.\d{3})+,\d+$")
_THOUSANDS_INTEGER = re.compile(r"^-?\d{1,3}(\.\d{3})+$")

_CURRENCY_STRIP = re.compile(r"^[R$\s]+|[R$\s]+$")


def parse_ptbr_number(raw: str) -> Union[float, int, None]:
    """Retorna um numero se `raw` casar com um padrao pt-BR inequivoco, senao None."""
    if raw is None:
        return None
    value = _CURRENCY_STRIP.sub("", str(raw).strip())
    if not value:
        return None

    negative = value.startswith("(") and value.endswith(")")
    if negative:
        value = value[1:-1].strip()

    if _INTEGER.match(value):
        num = int(value)
    elif _THOUSANDS_INTEGER.match(value):
        num = int(value.replace(".", ""))
    elif _DECIMAL_ONLY.match(value):
        num = float(value.replace(",", "."))
    elif _THOUSANDS_DECIMAL.match(value):
        num = float(value.replace(".", "").replace(",", "."))
    else:
        return None

    return -num if negative else num


def cell_value(raw):
    """Converte uma celula de tabela para numero quando seguro; senao mantem texto."""
    if raw is None:
        return ""
    text = str(raw).strip()
    if text == "":
        return ""
    num = parse_ptbr_number(text)
    return num if num is not None else text


CURRENCY_KEYWORDS = ("valor", "preço", "preco", "total", "saldo", "pago", "custo")


def is_currency_column(header: str) -> bool:
    lower = header.lower()
    return any(kw in lower for kw in CURRENCY_KEYWORDS)
