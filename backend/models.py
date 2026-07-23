from typing import List, Optional

from pydantic import BaseModel, Field


class PdfFileInput(BaseModel):
    name: str
    tabName: Optional[str] = None
    base64: str
    isPrincipal: bool = False


class ConvertRequest(BaseModel):
    files: List[PdfFileInput] = Field(default_factory=list)
    # Quando True, oculta a coluna "Responsável" na aba principal.
    hideResponsavel: bool = False
    # Quando True, oculta as abas secundárias (mantém apenas a principal visível).
    hideSecondaryTabs: bool = False


class TableData(BaseModel):
    tableName: Optional[str] = None
    headers: List[str]
    rows: List[List[str]]


class FileSummary(BaseModel):
    fileName: str
    tabName: str
    tableCount: int
    rowCount: int
    extractionMethod: str
    warning: Optional[str] = None


class ConvertResponse(BaseModel):
    success: bool = True
    base64: str
    filename: str
    summary: List[FileSummary]
