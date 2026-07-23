"""Pos-processamento do .xlsx para marcar as formulas UNIQUE/FILTER como
arrays dinamicos (spill) do Excel.

O openpyxl grava formulas de array apenas no formato legado CSE
(`<f t="array" ref="A1">`), que o Excel confina a UMA celula — entao um
UNIQUE/FILTER retorna so o primeiro valor em vez de derramar a lista inteira.
O formato moderno de "dynamic array" exige metadados (`xl/metadata.xml`) e o
atributo `cm="1"` na celula, alem de `aca="1" ca="1"` na formula para forcar o
recalculo na abertura. Este modulo injeta essas pecas, replicando exatamente o
que o Excel 365 grava (e o que o arquivo-modelo do projeto usa).
"""

import re
import zipfile
from io import BytesIO

METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
    'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
    'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    '</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)

METADATA_CONTENT_TYPE = (
    '<Override PartName="/xl/metadata.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
)

METADATA_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"

# <f t="array" ref="A2011"> (openpyxl) -> so os UNIQUE spills usam t="array".
# Qualquer coluna (A, F, ...), pois ha mais de um painel dinamico por aba.
_ARRAY_F_RE = re.compile(r'<f t="array" ref="([A-Z]+\d+)">')


def _patch_sheet(xml: str) -> tuple[str, bool]:
    """Marca as celulas de array dinamico com aca/ca (formula) e cm/t=str (celula)."""
    refs = _ARRAY_F_RE.findall(xml)
    if not refs:
        return xml, False

    # Formula: adiciona aca="1" ... ca="1"
    xml = _ARRAY_F_RE.sub(r'<f t="array" aca="1" ref="\1" ca="1">', xml)

    # Celula: adiciona t="str" cm="1" (cm=1 aponta para o cellMetadata unico)
    for ref in refs:
        cell_re = re.compile(r'(<c r="' + re.escape(ref) + r'")((?: [a-zA-Z]+="[^"]*")*)>')
        xml = cell_re.sub(lambda m: f'{m.group(1)}{m.group(2)} t="str" cm="1">', xml, count=1)

    return xml, True


def _next_rid(rels_xml: str) -> str:
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def inject_dynamic_arrays(xlsx_bytes: bytes) -> bytes:
    """Reescreve o .xlsx convertendo os arrays CSE legados em arrays dinamicos.
    Se nao houver nenhuma formula de array, devolve os bytes originais."""
    zin = zipfile.ZipFile(BytesIO(xlsx_bytes))
    names = zin.namelist()

    patched_any = False
    contents: dict[str, bytes] = {}
    for name in names:
        data = zin.read(name)
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            text = data.decode("utf-8")
            text, changed = _patch_sheet(text)
            if changed:
                patched_any = True
            data = text.encode("utf-8")
        contents[name] = data

    if not patched_any:
        return xlsx_bytes

    # Content types: registra o metadata.xml
    ct = contents["[Content_Types].xml"].decode("utf-8")
    if "metadata.xml" not in ct:
        ct = ct.replace("</Types>", METADATA_CONTENT_TYPE + "</Types>")
        contents["[Content_Types].xml"] = ct.encode("utf-8")

    # Relationships do workbook: aponta para metadata.xml
    rels_name = "xl/_rels/workbook.xml.rels"
    rels = contents[rels_name].decode("utf-8")
    if "metadata.xml" not in rels:
        rid = _next_rid(rels)
        rel = f'<Relationship Id="{rid}" Type="{METADATA_REL_TYPE}" Target="metadata.xml"/>'
        rels = rels.replace("</Relationships>", rel + "</Relationships>")
        contents[rels_name] = rels.encode("utf-8")

    contents["xl/metadata.xml"] = METADATA_XML.encode("utf-8")

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    return out.getvalue()
