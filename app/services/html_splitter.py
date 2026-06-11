# app/services/html_splitter.py
"""
Estrategia de caché óptima:
1. Convierte el .docx completo a HTML una sola vez con mammoth
2. Cachea el HTML completo
3. Al pedir una sección, extrae el fragmento del HTML cacheado
   → sin re-procesar mammoth, sin crear .docx temporales
"""
import re
from lxml import etree
from io import StringIO
from typing import Optional

# Clases de heading que mammoth genera
HEADING_CLASSES = {"doc-h1", "doc-h2", "doc-h3", "doc-h4"}
H1_TAGS         = {"h1"}


def split_html_by_h1(full_html: str) -> dict[int, str]:
    """
    Divide el HTML completo en secciones por cada H1.
    Retorna dict: { h1_paragraph_index: html_fragment }

    Nota: el h1_paragraph_index aquí es el índice del H1
    dentro del HTML (orden de aparición), no el del .docx.
    Se mapea luego con el índice real del .docx.
    """
    # Parsear el HTML completo
    parser  = etree.HTMLParser()
    tree    = etree.parse(StringIO(f"<div>{full_html}</div>"), parser)
    body    = tree.find(".//div")

    if body is None:
        return {}

    sections: dict[int, list] = {}
    current_key = None
    h1_order    = 0

    for elem in body:
        tag       = elem.tag.lower() if isinstance(elem.tag, str) else ""
        classes   = (elem.get("class") or "").split()
        is_h1     = tag == "h1" or "doc-h1" in classes

        if is_h1:
            current_key          = h1_order
            sections[current_key] = [elem]
            h1_order             += 1
        elif current_key is not None:
            sections[current_key].append(elem)

    # Serializar cada sección a HTML
    result: dict[int, str] = {}
    for key, elements in sections.items():
        parts = []
        for el in elements:
            parts.append(
                etree.tostring(el, encoding="unicode", method="html")
            )
        result[key] = "".join(parts)

    return result


def extract_section_from_html(
    full_html:  str,
    h1_order:   int,   # índice del H1 en el HTML (0-based)
) -> Optional[str]:
    """
    Extrae el fragmento HTML de una sección específica.
    """
    sections = split_html_by_h1(full_html)
    return sections.get(h1_order)