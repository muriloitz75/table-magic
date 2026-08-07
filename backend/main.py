import base64
import re
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from .excel.builder import build_workbook
from .extraction import extract_tables
from .models import ConvertRequest, ConvertResponse, FileSummary

app = FastAPI(title="Table Magic")


@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"error": exc.detail})


@app.exception_handler(Exception)
async def generic_exception_handler(request, exc: Exception):
    return JSONResponse(status_code=500, content={"error": str(exc) or "Erro desconhecido durante a conversao."})


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/convert", response_model=ConvertResponse)
def convert(payload: ConvertRequest):
    if not payload.files:
        raise HTTPException(status_code=400, detail="Nenhum arquivo enviado.")

    file_results = []
    for f in payload.files:
        if not f.base64:
            raise HTTPException(status_code=400, detail=f'Arquivo "{f.name}" nao contem dados validos.')

        fallback_name = "Principal" if f.isPrincipal else (re.sub(r"\.pdf$", "", f.name, flags=re.IGNORECASE) or "Planilha")
        tab_name = f.tabName or fallback_name

        try:
            pdf_bytes = base64.b64decode(f.base64)
            extraction = extract_tables(pdf_bytes)
            file_results.append(
                {
                    "fileName": f.name,
                    "tabName": tab_name,
                    "tables": extraction["tables"],
                    "extractionMethod": extraction["method"],
                    "warning": extraction["warning"],
                }
            )
        except Exception as exc:
            # Um PDF com problema nao deve derrubar o lote inteiro: registra
            # a falha como aba de aviso e segue processando os demais.
            file_results.append(
                {
                    "fileName": f.name,
                    "tabName": tab_name,
                    "tables": [],
                    "extractionMethod": "erro",
                    "warning": f"Falha ao processar este arquivo: {exc}",
                }
            )

    xlsx_bytes, summary = build_workbook(
        file_results,
        hide_responsavel=payload.hideResponsavel,
        hide_secondary_tabs=payload.hideSecondaryTabs,
        disable_principal_tab=payload.disablePrincipalTab,
    )
    b64_excel = base64.b64encode(xlsx_bytes).decode("ascii")

    return ConvertResponse(
        success=True,
        base64=b64_excel,
        filename="Table_Magic_Combined.xlsx",
        summary=[FileSummary(**s) for s in summary],
    )


# Em producao, o FastAPI tambem serve o build estatico do frontend (dist/).
# Em dev, o Vite roda separado (porta 5173) e faz proxy de /api para ca.
DIST_DIR = Path(__file__).resolve().parent.parent / "dist"
if DIST_DIR.exists():
    app.mount("/assets", StaticFiles(directory=DIST_DIR / "assets"), name="assets")

    @app.get("/{full_path:path}")
    def spa_fallback(full_path: str):
        return FileResponse(DIST_DIR / "index.html")


if __name__ == "__main__":
    import os
    import uvicorn

    port = int(os.getenv("PORT", 3000))
    uvicorn.run("backend.main:app", host="0.0.0.0", port=port)
