# app/routers/docx.py
from fastapi import APIRouter, UploadFile, File, HTTPException, Query
from fastapi.responses import StreamingResponse
from io import BytesIO

from app.models.schemas import (
    ParseResult, FillRequest, FillResult, DeleteResult, MarkVariablesRequest,
)
from app.core.validators import validate_docx_file
from app.services import docx_service
from app.services.variable_marker import mark_variables, list_bookmarks
from app.utils.minio_client import download_from_minio
from app.utils.section_cache import invalidate_section_cache
from app.utils.docx_cache import get_docx_cached

router = APIRouter(tags=["Documents"])


@router.get("/list-docx")
def list_docx():
    try:
        return docx_service.list_documents()
    except Exception as e:
        raise HTTPException(500, f"Error listando documentos: {e}")


@router.post("/parse-docx", response_model=ParseResult)
async def parse_docx(file: UploadFile = File(...)):
    content = await file.read()
    validate_docx_file(file, content)
    try:
        result = docx_service.upload_and_parse(content, file.filename)
        invalidate_section_cache(result.minio_key)
        return result
    except Exception as e:
        raise HTTPException(500, f"Error procesando el documento: {e}")


@router.get("/parse-docx/{minio_key:path}", response_model=ParseResult)
def get_parse_docx(minio_key: str):
    try:
        return docx_service.reload_document(minio_key)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error recargando documento: {e}")


@router.get("/docx-to-html/{minio_key:path}")
def get_docx_to_html(minio_key: str):
    try:
        return docx_service.get_docx_html(minio_key)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error convirtiendo a HTML: {e}")


@router.post("/fill-docx", response_model=FillResult)
async def fill_docx(request: FillRequest):
    try:
        buffer = download_from_minio(request.minio_key)
    except Exception as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")
    try:
        return docx_service.fill_and_save(
            buffer       = buffer,
            minio_key    = request.minio_key,
            variables    = request.variables,
            tablas       = [t.model_dump() for t in request.tablas],
            imagenes     = [i.model_dump() for i in request.imagenes],
            bloques      = [b.model_dump() for b in request.bloques],
            replacements = [r.model_dump() for r in request.replacements],
        )
    except Exception as e:
        raise HTTPException(500, f"Error procesando: {e}")


@router.post("/fill-docx-download")
async def fill_docx_download(request: FillRequest):
    try:
        buffer = download_from_minio(request.minio_key)
    except Exception as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")
    try:
        result_buffer = docx_service.fill_and_stream(
            buffer       = buffer,
            minio_key    = request.minio_key,
            variables    = request.variables,
            tablas       = [t.model_dump() for t in request.tablas],
            imagenes     = [i.model_dump() for i in request.imagenes],
            bloques      = [b.model_dump() for b in request.bloques],
            replacements = [r.model_dump() for r in request.replacements],
        )
    except Exception as e:
        raise HTTPException(500, f"Error procesando: {e}")

    filename = request.minio_key.split("/")[-1]
    return StreamingResponse(
        BytesIO(result_buffer),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename=generado_{filename}"},
    )


@router.delete("/delete-docx", response_model=DeleteResult)
async def delete_docx(minio_key: str = Query(...)):
    if not minio_key:
        raise HTTPException(400, "Se requiere minio_key")
    try:
        docx_service.delete_document(minio_key)
        return DeleteResult(
            deleted   = True,
            minio_key = minio_key,
            message   = f"Archivo eliminado: {minio_key}",
        )
    except ValueError as e:
        raise HTTPException(403, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error eliminando: {e}")
    
@router.post("/upload-image")
async def upload_image(file: UploadFile = File(...)):
    """
    Sube una imagen a MinIO y retorna su key y URL presignada.
    """
    import time
    from app.utils.minio_client import upload_to_minio, get_presigned_url
 
    if not file.content_type.startswith("image/"):
        raise HTTPException(400, "Solo se permiten archivos de imagen")
 
    content = await file.read()
    if not content:
        raise HTTPException(400, "El archivo está vacío")
 
    timestamp = int(time.time())
    safe_name = (file.filename or "imagen").replace(" ", "_")
    key       = f"docx-images/{timestamp}_{safe_name}"
 
    try:
        upload_to_minio(key, content, content_type=file.content_type)
        url = get_presigned_url(key)
    except Exception as e:
        raise HTTPException(500, f"Error subiendo imagen: {e}")
 
    return {"key": key, "url": url}


@router.get("/extract-text")
def extract_text(key: str = Query(..., description="MinIO key del .docx")):
    try:
        return docx_service.extract_plain_text(key)
    except FileNotFoundError as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")
    except Exception as e:
        raise HTTPException(500, f"Error extrayendo texto: {e}")


@router.get("/docx-bookmarks/{minio_key:path}")
def get_docx_bookmarks(minio_key: str):
    """
    Mapa `variable → marcador` del documento marcado. El editor lo usa para
    saltar a la variable que se elige en el índice lateral.
    """
    try:
        docx_buffer = get_docx_cached(minio_key)
    except Exception as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")

    try:
        return {"minio_key": minio_key, "bookmarks": list_bookmarks(docx_buffer)}
    except Exception as e:
        raise HTTPException(500, f"Error leyendo marcadores: {e}")


@router.post("/docx-marked")
def get_docx_marked(request: MarkVariablesRequest):
    """
    Devuelve el .docx con cada [VARIABLE] resaltada — amarillo si está vacía,
    verde si ya tiene valor. Es lo que se abre en el editor OnlyOffice.
    """
    try:
        docx_buffer = get_docx_cached(request.minio_key)
    except FileNotFoundError as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")
    except Exception as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")

    try:
        marked, total, _, _ = mark_variables(docx_buffer, request.values)
    except Exception as e:
        raise HTTPException(500, f"Error marcando variables: {e}")

    filename = request.minio_key.split("/")[-1]
    return StreamingResponse(
        BytesIO(marked),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Marked-Variables": str(total),
        },
    )