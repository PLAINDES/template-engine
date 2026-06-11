# app/services/sections/heading_parser.py
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
from typing import List
from app.models.schemas import HeadingItem


HEADING_STYLES = {
    "heading 1": 1, "heading 2": 2, "heading 3": 3,
    "heading 4": 4, "heading 5": 5,
    "título 1":  1, "título 2":  2, "título 3":  3,
    "titulo 1":  1, "titulo 2":  2, "titulo 3":  3,
}


def _get_heading_level(para) -> int | None:
    """
    Detecta si un párrafo es heading y retorna su nivel (1-5).
    Soporta estilos estándar y estilos personalizados con outlineLvl.
    Retorna None si no es heading.
    """
    style_name = (para.style.name or "").lower().strip()

    if style_name in HEADING_STYLES:
        return HEADING_STYLES[style_name]

    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        outlineLvl = pPr.find(qn("w:outlineLvl"))
        if outlineLvl is not None:
            val = outlineLvl.get(qn("w:val"))
            if val is not None:
                return int(val) + 1

    return None


def _flat_headings(doc: Document) -> List[dict]:
    """
    Extrae todos los headings en orden lineal.
    Retorna lista de { index, level, text }.
    """
    result = []
    for idx, para in enumerate(doc.paragraphs):
        level = _get_heading_level(para)
        if level is None:
            continue
        text = para.text.strip()
        if not text:
            continue
        result.append({"index": idx, "level": level, "text": text})
    return result


def _build_tree(flat: List[dict]) -> List[HeadingItem]:
    """
    Convierte la lista plana en árbol jerárquico.
    Cada H1 agrupa sus H2, H3, etc.
    """
    root: List[HeadingItem] = []
    stack: List[HeadingItem] = []

    for item in flat:
        node = HeadingItem(
            index=item["index"],
            level=item["level"],
            text=item["text"],
        )
        while stack and stack[-1].level >= node.level:
            stack.pop()

        if stack:
            stack[-1].children.append(node)
        else:
            root.append(node)

        stack.append(node)

    return root


def parse_headings(docx_buffer: bytes) -> List[HeadingItem]:
    """
    Entry point: retorna el árbol de headings del .docx.
    """
    doc = Document(BytesIO(docx_buffer))
    flat = _flat_headings(doc)
    return _build_tree(flat)