from typing import Optional
import mammoth
from docx import Document as PyDocxDoc
from ..schemas import DocumentMeta

def docx_to_html_and_meta(fp) -> tuple[str, Optional[DocumentMeta]]:
    result = mammoth.convert_to_html(fp)
    html = (result.value or "").strip()

    try:
        fp.seek(0)
        d = PyDocxDoc(fp)
        core = d.core_properties
        meta = DocumentMeta(
            title=core.title or "Untitled",
            author=core.author or "Anonymous",
        )
    except Exception:
        meta = None

    return html, meta
