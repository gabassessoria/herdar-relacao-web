from io import BytesIO
from pathlib import Path
from urllib.parse import quote
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from core.processing import list_columns, process

app = FastAPI(title="Herdar Relação Anterior", version="1.0.0")
BASE = Path(__file__).resolve().parent.parent
HTML = (BASE / "templates" / "index.html").read_text(encoding="utf-8")

async def read_file(f: UploadFile):
    if not f.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Envie arquivos .xlsx.")
    b = await f.read()
    if not b:
        raise HTTPException(400, "Arquivo vazio.")
    return b

@app.get("/", response_class=HTMLResponse)
async def home():
    return HTML

@app.get("/api/health")
async def health():
    return {"ok": True}

@app.post("/api/colunas")
async def colunas(fonte: UploadFile = File(...)):
    try:
        b = await read_file(fonte)
        return JSONResponse(list_columns(b))
    except ValueError as e:
        raise HTTPException(400, str(e))

@app.post("/api/processar")
async def processar(
    fonte: UploadFile = File(...),
    filha: UploadFile = File(...),
    colunas: str = Form(...),
    sobrescrever: str = Form("true"),
):
    try:
        a = await read_file(fonte)
        b = await read_file(filha)
        selected = json.loads(colunas)
        overwrite = sobrescrever.lower() == "true"
        data, summary = process(a, b, selected, overwrite)
        base = filha.filename.rsplit(".", 1)[0]
        name = f"{base}_COM_HISTORICO.xlsx"
        return StreamingResponse(
            BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}",
                "X-Reapareceram": str(summary["reapareceram"]),
                "X-Com-Heranca": str(summary["com_heranca"]),
                "X-Celulas-Herdadas": str(summary["celulas_herdadas"]),
                "X-Colunas-Criadas": ",".join(summary["colunas_criadas"]),
            },
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
