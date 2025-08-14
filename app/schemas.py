from pydantic import BaseModel, Field

class DocumentMeta(BaseModel):
    title: str = Field(default="Untitled", max_length=200)
    author: str = Field(default="Anonymous", max_length=120)

class ImportDocxResponse(BaseModel):
    html: str
    metadata: DocumentMeta | None = None

class ExportDocxRequest(BaseModel):
    html: str
    meta: DocumentMeta
