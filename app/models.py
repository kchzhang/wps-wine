from pydantic import BaseModel
from typing import Optional, Dict, Any
from enum import Enum

class ConvertFormat(str, Enum):
    PDF = "pdf"
    DOC = "doc"
    DOCX = "docx"
    TXT = "txt"
    HTML = "html"
    RTF = "rtf"

class ConvertRequest(BaseModel):
    format: ConvertFormat = ConvertFormat.PDF
    options: Optional[Dict[str, Any]] = None

class ConvertResponse(BaseModel):
    success: bool
    message: str
    output_file: Optional[str] = None
    file_size: Optional[int] = None
    conversion_time: Optional[float] = None

class HealthResponse(BaseModel):
    status: str
    version: str
    wps_available: bool