# app/services/sections/outline_builder.py
"""
Índice completo del documento: todos los títulos con sus variables y el
marcador que permite saltar a cada uno desde el editor.

A diferencia de `extract_section_structure`, que devuelve el árbol de una sola
sección H1, aquí se recorre el documento entero de una vez — es lo que pinta el
índice lateral de la plataforma.
"""
from io import BytesIO
from typing import Any, Dict, List

from docx import Document

from app.services.parser import VAR_RE, TABLE_KEYS, IMAGE_KEYS
from app.services.sections.heading_parser import _get_heading_level
from app.services.variable_marker import heading_bookmark


def _collect_variables(text: str, node: Dict[str, Any]) -> None:
    """Añade al nodo las variables del texto, sin repetir y sin tablas/imágenes."""
    existing = {v["key"] for v in node["variables"]}
    for match in VAR_RE.finditer(text):
        key = match.group(1)
        upper = key.upper()
        if upper in TABLE_KEYS or upper in IMAGE_KEYS or key in existing:
            continue
        existing.add(key)
        node["variables"].append(
            {"key": key, "label": key.replace("_", " ").title()}
        )


def build_outline(docx_buffer: bytes) -> List[Dict[str, Any]]:
    """
    Árbol de títulos del documento. Cada nodo:
        level, text, para_idx, bookmark, variables[], children[]

    Las variables se cuelgan del título más profundo que esté abierto, igual
    criterio que usa la estructura por sección.
    """
    doc = Document(BytesIO(docx_buffer))

    outline: List[Dict[str, Any]] = []
    stack: List[Dict[str, Any]] = []  # títulos abiertos, por profundidad

    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        level = _get_heading_level(para)

        if level is not None and text.strip():
            clean = VAR_RE.sub("", text).replace("\n", " ").strip() or text.strip()
            node: Dict[str, Any] = {
                "level":     level,
                "text":      clean,
                "para_idx":  idx,
                "bookmark":  heading_bookmark(idx),
                "variables": [],
                "children":  [],
            }

            depth = level - 1
            stack = stack[:depth]

            if stack:
                stack[-1]["children"].append(node)
            else:
                outline.append(node)

            stack.append(node)
            # Algunas plantillas ponen la variable en el propio título
            _collect_variables(text, node)
            continue

        if stack:
            _collect_variables(text, stack[-1])

    return outline
