# app/routers/sections.py
from fastapi import APIRouter, HTTPException, Query, BackgroundTasks, Request
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


@router.post("/full-with-replacements")
async def get_section_full_with_replacements(request: Request):
    """
    Igual que /full/ pero aplica reemplazos de texto al HTML resultante.
    """
    import re as re_module

    body = await request.json()
    minio_key    = body.get("minio_key", "")
    h1_index     = body.get("h1_index", 0)
    replacements = body.get("replacements", [])

    print(f"[full-with-replacements] key={minio_key} h1={h1_index} repls={len(replacements)}")

    try:
        result = get_section_full(minio_key, h1_index)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error: {e}")

    if not isinstance(result, dict):
        result = {"html": "", "structure": []}

    html = result.get("html", "")
    if replacements and html:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')

        for repl in replacements:
            if not isinstance(repl, dict):
                continue
            orig = repl.get("originalText", "")
            new_text = repl.get("newText", "")
            if not orig or not new_text:
                continue

            # Normalizar espacios para comparar
            norm = lambda s: re_module.sub(r'\s+', ' ', s).strip()
            norm_orig = norm(orig)

            # Buscar en el texto completo del HTML
            full_text = soup.get_text()
            norm_full = norm(full_text)

            if norm_orig not in norm_full:
                print(f"[full-with-replacements] ✗ No encontrado: {orig[:50]}...")
                continue

            # Buscar en cada text node
            found = False
            for text_node in soup.find_all(string=True):
                node_text = str(text_node)
                if norm(node_text).find(norm_orig) == -1 and norm_orig not in norm(node_text):
                    continue
                # Reemplazar en este nodo
                import re
                pattern = re.sub(r'\s+', r'\\s+', re.escape(norm_orig))
                match = re.search(pattern, node_text)
                if match:
                    styled_tag = soup.new_tag("span")
                    styled_tag["style"] = "background:#ede9fe;border-bottom:2px solid #7c3aed;border-radius:2px;padding:0 2px;"
                    styled_tag.string = new_text
                    before = node_text[:match.start()]
                    after = node_text[match.end():]
                    text_node.replace_with(before)
                    text_node_parent = soup.find(string=before)
                    if text_node_parent:
                        text_node_parent.insert_after(after)
                        text_node_parent.insert_after(styled_tag)
                    found = True
                    print(f"[full-with-replacements] ✓ Reemplazo en nodo: {orig[:50]}...")
                    break

            # Fallback: buscar concatenando nodos contiguos
            if not found:
                # Reemplazo directo en el HTML string como último recurso
                try:
                    words = norm_orig.split(' ')
                    words = [w for w in words if w]
                    pattern = ''.join(
                        re_module.escape(w) + (r'(?:(?:<[^>]*>)*[\s\n\r]*(?:<[^>]*>)*' if i < len(words) - 1 else '')
                        + (')' if i < len(words) - 1 else '')
                        for i, w in enumerate(words)
                    )
                    styled = f'<span style="background:#ede9fe;border-bottom:2px solid #7c3aed;border-radius:2px;padding:0 2px;">{new_text}</span>'
                    new_html = re_module.sub(pattern, styled, str(soup), count=1, flags=re_module.DOTALL)
                    if new_html != str(soup):
                        soup = BeautifulSoup(new_html, 'html.parser')
                        found = True
                        print(f"[full-with-replacements] ✓ Reemplazo regex fallback: {orig[:50]}...")
                except Exception as ex:
                    print(f"[full-with-replacements] Error fallback: {ex}")

            if not found:
                print(f"[full-with-replacements] ✗ No se pudo reemplazar: {orig[:50]}...")

        result["html"] = str(soup)

    return result


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