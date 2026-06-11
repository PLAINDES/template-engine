# app/routers/sections.py
from fastapi import APIRouter, HTTPException, Query, BackgroundTasks
from fastapi.responses import Response
from app.models.schemas import HeadingsResult, ExtractSectionsRequest
from app.services.sections.heading_parser import parse_headings
from app.services.sections.section_builder import build_section_document
from app.services.sections.section_html_extractor import (
    get_section_full,
    extract_section_as_html,
    extract_section_structure,
    invalidate_document_cache,
)
from app.utils.docx_cache import get_docx_cached

router = APIRouter(prefix="/sections", tags=["Sections"])


@router.get("/headings/{minio_key:path}", response_model=HeadingsResult)
async def get_headings(minio_key: str):
    try:
        docx_buffer = get_docx_cached(minio_key)
    except Exception as e:
        raise HTTPException(status_code=404, detail=str(e))
    try:
        headings = parse_headings(docx_buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")
    return HeadingsResult(
        minio_key=minio_key,
        filename=minio_key.split("/")[-1],
        headings=headings,
    )


@router.get("/full/{minio_key:path}")
async def get_section_full_endpoint(
    minio_key: str,
    h1_index:  int = Query(..., description="paragraph_index del H1"),
):
    """
    Endpoint combinado — HTML + estructura.
    Primera llamada procesa el doc completo y cachea todo.
    Siguientes llamadas son instantáneas.
    """
    try:
        result = get_section_full(minio_key, h1_index)
        return result
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")


@router.post("/warmup/{minio_key:path}")
async def warmup_cache(
    minio_key:        str,
    background_tasks: BackgroundTasks,
):
    """
    Precalienta el caché del documento en background.
    Llamar justo después de subir/abrir un documento.
    La respuesta es inmediata — el procesamiento ocurre en background.
    """
    def _warmup():
        try:
            from app.utils.full_html_cache import get_full_html_cache
            from app.services.sections.section_html_extractor import _build_full_cache
            if get_full_html_cache(minio_key) is None:
                _build_full_cache(minio_key)
        except Exception as e:
            print(f"[warmup] Error: {e}")

    background_tasks.add_task(_warmup)
    return {"status": "warming_up", "minio_key": minio_key}


@router.get("/html/{minio_key:path}")
async def get_section_html(
    minio_key: str,
    h1_index:  int = Query(..., description="paragraph_index del H1"),
):
    try:
        return extract_section_as_html(minio_key, h1_index)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")


@router.get("/structure/{minio_key:path}")
async def get_section_structure(
    minio_key: str,
    h1_index:  int = Query(..., description="paragraph_index del H1"),
):
    try:
        structure = extract_section_structure(minio_key, h1_index)
        return {"structure": structure}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")


@router.post("/extract")
async def extract_sections(request: ExtractSectionsRequest):
    if not request.selected_indexes:
        raise HTTPException(status_code=400, detail="Selecciona al menos una sección")
    try:
        docx_buffer = get_docx_cached(request.minio_key)
    except Exception as e:
        raise HTTPException(status_code=404, detail=str(e))
    try:
        result_buffer = build_section_document(
            docx_buffer           = docx_buffer,
            selected_para_indexes = request.selected_indexes,
            variables             = request.variables,
            tablas                = request.tablas,
            imagenes              = request.imagenes,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")

    filename = request.minio_key.split("/")[-1]
    invalidate_document_cache(request.minio_key)

    return Response(
        content    = result_buffer,
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers    = {"Content-Disposition": f'attachment; filename="secciones_{filename}"'},
    )