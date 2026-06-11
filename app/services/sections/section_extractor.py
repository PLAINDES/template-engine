# app/services/sections/section_extractor.py
import copy
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
from typing import List, Set
from app.services.sections.heading_parser import _get_heading_level


def _collect_all_indexes(doc: Document, selected_para_indexes: List[int]) -> Set[int]:
    paragraphs   = doc.paragraphs
    total        = len(paragraphs)
    h1_indexes   = [
        idx for idx, para in enumerate(paragraphs)
        if _get_heading_level(para) == 1
    ]
    selected_set = set(selected_para_indexes)
    included: Set[int] = set()

    for i, h1_idx in enumerate(h1_indexes):
        if h1_idx not in selected_set:
            continue
        next_h1 = h1_indexes[i + 1] if i + 1 < len(h1_indexes) else total
        included.update(range(h1_idx, next_h1))

    return included


def extract_sections(docx_buffer: bytes, selected_para_indexes: List[int]) -> bytes:
    """
    Estrategia: cargar el doc original, clonar su body completo,
    y ELIMINAR los elementos que NO queremos.
    Así se preserva 100% estilos, numbering, fonts, etc.
    """
    doc              = Document(BytesIO(docx_buffer))
    included_indexes = _collect_all_indexes(doc, selected_para_indexes)

    if not included_indexes:
        raise ValueError("Ninguna sección encontrada con los índices proporcionados")

    # Clonar el documento completo desde el buffer original
    new_doc = Document(BytesIO(docx_buffer))

    body          = new_doc.element.body
    para_idx      = 0
    last_included = False
    to_remove     = []

    for child in list(body):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            if para_idx not in included_indexes:
                to_remove.append(child)
                last_included = False
            else:
                last_included = True
            para_idx += 1

        elif tag == "tbl":
            if not last_included:
                to_remove.append(child)

        elif tag == "sectPr":
            # Preservar siempre la configuración de página
            pass

    # Eliminar elementos no deseados
    for el in to_remove:
        body.remove(el)

    output = BytesIO()
    new_doc.save(output)
    return output.getvalue()