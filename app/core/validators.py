# app/core/validators.py
from fastapi import HTTPException, UploadFile
from app.core.config import settings

DOCX_MAGIC_BYTES = b"PK\x03\x04"  # Todo .docx es un ZIP internamente


def validate_docx_file(file: UploadFile, content: bytes) -> None:
    """
    Valida que el archivo sea un .docx real:
    1. Extensión correcta
    2. MIME type permitido
    3. Tamaño dentro del límite
    4. Magic bytes (firma real de ZIP/DOCX)
    """
    # 1. Extensión
    filename = file.filename or ""
    if not filename.lower().endswith(".docx"):
        raise HTTPException(400, "Solo se permiten archivos .docx")

    # 2. MIME type
    if file.content_type not in settings.ALLOWED_MIME_TYPES:
        raise HTTPException(
            400,
            f"Tipo de archivo no permitido: {file.content_type}. "
            f"Se esperaba un documento Word (.docx)"
        )

    # 3. Tamaño
    max_bytes = settings.MAX_FILE_SIZE_MB * 1024 * 1024
    if len(content) > max_bytes:
        raise HTTPException(
            413,
            f"El archivo supera el límite de {settings.MAX_FILE_SIZE_MB}MB"
        )

    # 4. Contenido real (magic bytes)
    if not content.startswith(DOCX_MAGIC_BYTES):
        raise HTTPException(
            400,
            "El archivo no es un documento Word válido "
            "(firma de archivo incorrecta)"
        )