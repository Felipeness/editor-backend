from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from .schemas import ImportDocxResponse, ExportDocxRequest
from .converters.docx_to_html import docx_to_html_and_meta
from .converters.html_to_docx import html_to_docx
from io import BytesIO

app = FastAPI(title="Docx Backend", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://127.0.0.1:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/docx/import", response_model=ImportDocxResponse)
async def import_docx(file: UploadFile = File(...), request: Request = None):

    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Expected a .docx file")


    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Empty file")


    try:
        fp = BytesIO(data)
        fp.seek(0)
        html, meta = docx_to_html_and_meta(fp)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Import error: {e!s}")


    return ImportDocxResponse(html=html or "<p></p>", metadata=meta or {})

@app.post("/docx/export")
async def export_docx(req: ExportDocxRequest):
    try:
        content = html_to_docx(req.html, req.meta.title, req.meta.author)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export error: {e!s}")

    filename = (req.meta.title or "document").replace('"', "'")
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}.docx"'
    }
    return StreamingResponse(
        BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
