# app/services/sections/section_html_extractor.py
"""
Enfoque simple y correcto:
- Cada sección se extrae directamente del .docx por rango de párrafos
- El .docx completo se cachea (evita re-descargar de MinIO)
- El HTML de cada sección se cachea (evita re-procesar con mammoth)
- Resultado: primera vez lenta, siguientes instantáneas
"""
import re
from docx import Document
from io import BytesIO
from app.services.sections.heading_parser import _get_heading_level
from app.services.html_converter import docx_to_html
from app.utils.docx_cache import get_docx_cached
from app.utils.full_html_cache import (
    get_full_html_cache,
    set_full_html_cache,
    invalidate_full_html_cache,
)

VAR_RE = re.compile(
    r'\[([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)\]',
    re.UNICODE,
)
TABLE_KEYS = {"CREAR_O_AÑADIR_TABLA","CREAR_O_ANADIR_TABLA","AÑADIR_TABLA",
              "ANADIR_TABLA","TABLA","CREAR_TABLA","INSERTAR_TABLA"}
IMAGE_KEYS = {"IMAGEN","IMAGEN_PEGAR","PEGAR_IMAGEN","INSERTAR_IMAGEN","FOTO"}


def _get_h1_ranges(doc: Document) -> list:
    """
    Retorna lista de { index, start, end } para cada H1.
    start y end son los paragraph_index del rango que le pertenece.
    """
    paragraphs = doc.paragraphs
    total      = len(paragraphs)

    h1_list = [
        idx for idx, para in enumerate(paragraphs)
        if _get_heading_level(para) == 1
    ]

    ranges = []
    for i, h1_idx in enumerate(h1_list):
        end = h1_list[i + 1] if i + 1 < len(h1_list) else total
        ranges.append({
            "index": h1_idx,
            "start": h1_idx,
            "end":   end,
        })

    return ranges


def _extract_docx_section(docx_buffer: bytes, start: int, end: int) -> bytes:
    """
    Crea un nuevo .docx con solo los párrafos del rango [start, end).
    Preserva estilos y formato del original.
    """
    new_doc       = Document(BytesIO(docx_buffer))
    body          = new_doc.element.body
    para_idx      = 0
    last_included = False
    to_remove     = []

    for child in list(body):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            if not (start <= para_idx < end):
                to_remove.append(child)
                last_included = False
            else:
                last_included = True
            para_idx += 1

        elif tag == "tbl":
            if not last_included:
                to_remove.append(child)

    for el in to_remove:
        try:
            body.remove(el)
        except Exception:
            pass

    output = BytesIO()
    new_doc.save(output)
    return output.getvalue()


def _extract_structure(doc: Document, start: int, end: int) -> list:
    paragraphs = doc.paragraphs
    structure  = []
    current_h2 = None

    for idx in range(start, end):
        if idx >= len(paragraphs):
            break
        para  = paragraphs[idx]
        level = _get_heading_level(para)
        text  = para.text.strip()

        if level == 1 and idx == start:
            continue

        if level == 2:
            current_h2 = {
                "level": 2, "text": text,
                "para_idx": idx, "variables": [], "children": [],
            }
            structure.append(current_h2)

        elif level and level >= 3:
            node = {
                "level": level, "text": text,
                "para_idx": idx, "variables": [], "children": [],
            }
            if current_h2:
                current_h2["children"].append(node)
            else:
                structure.append(node)

        else:
            full_text = "".join(run.text for run in para.runs)
            for match in VAR_RE.finditer(full_text):
                key = match.group(1)
                if key.upper() in TABLE_KEYS or key.upper() in IMAGE_KEYS:
                    continue
                if current_h2 and key not in [v["key"] for v in current_h2["variables"]]:
                    current_h2["variables"].append({
                        "key":   key,
                        "label": key.replace("_", " ").title(),
                    })

    return structure


def _build_cache(minio_key: str) -> dict:
    """
    Procesa TODAS las secciones del documento y las cachea.
    Se llama una sola vez por documento.
    """
    docx_buffer = get_docx_cached(minio_key)
    doc         = Document(BytesIO(docx_buffer))
    h1_ranges   = _get_h1_ranges(doc)

    sections_html      = {}
    sections_structure = {}

    for h1 in h1_ranges:
        h1_idx  = h1["index"]
        start   = h1["start"]
        end     = h1["end"]

        # Extraer .docx de la sección y convertir a HTML
        section_buffer          = _extract_docx_section(docx_buffer, start, end)
        html_result             = docx_to_html(section_buffer)
        sections_html[h1_idx]      = html_result["html"]
        sections_structure[h1_idx] = _extract_structure(doc, start, end)

    cache_data = {
        "sections":  sections_html,
        "structure": sections_structure,
    }

    set_full_html_cache(minio_key, cache_data)
    return cache_data


def get_section_full(minio_key: str, h1_index: int) -> dict:
    """
    Retorna { html, structure } para la sección indicada.
    Si el caché no existe lo construye completo.
    Si la sección específica no está (raro), la procesa al vuelo.
    """
    cached = get_full_html_cache(minio_key)

    # Caché no existe → construir todo
    if cached is None:
        cached = _build_cache(minio_key)

    html      = cached["sections"].get(h1_index)
    structure = cached["structure"].get(h1_index, [])

    # Sección no encontrada en caché → procesar solo esa
    if html is None:
        docx_buffer = get_docx_cached(minio_key)
        doc         = Document(BytesIO(docx_buffer))
        h1_ranges   = _get_h1_ranges(doc)

        target = next((r for r in h1_ranges if r["index"] == h1_index), None)
        if target is None:
            raise ValueError(
                f"h1_index={h1_index} no existe en el documento."
            )

        section_buffer = _extract_docx_section(
            docx_buffer, target["start"], target["end"]
        )
        html_result = docx_to_html(section_buffer)
        html        = html_result["html"]
        structure   = _extract_structure(doc, target["start"], target["end"])

        # Actualizar caché
        cached["sections"][h1_index]  = html
        cached["structure"][h1_index] = structure
        set_full_html_cache(minio_key, cached)

    return {
        "html":      html,
        "messages":  [],
        "structure": structure,
    }


def extract_section_as_html(minio_key: str, h1_index: int) -> dict:
    result = get_section_full(minio_key, h1_index)
    return {"html": result["html"], "messages": result["messages"]}


def extract_section_structure(minio_key: str, h1_index: int) -> list:
    return get_section_full(minio_key, h1_index)["structure"]


def invalidate_document_cache(minio_key: str) -> None:
    from app.utils.docx_cache import invalidate_cache
    invalidate_cache(minio_key)
    invalidate_full_html_cache(minio_key)