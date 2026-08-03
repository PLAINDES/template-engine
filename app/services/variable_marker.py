# app/services/variable_marker.py
"""
Resalta los placeholders [VARIABLE] dentro del .docx para que se distingan
en el editor OnlyOffice, igual que los spans `doc-var` del visor HTML.

El marcado se hace sobre el documento porque una plantilla real tiene cientos
de variables: hacerlo desde el navegador con la API del editor obligaría a una
búsqueda por variable y sería inviable.
"""
import copy
import re
from io import BytesIO
from typing import Dict, Iterator, List, Tuple

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from app.services.parser import VAR_RE
from app.services.sections.heading_parser import _get_heading_level

# Amarillo = variable sin valor · Verde = variable ya rellenada
COLOR_EMPTY = WD_COLOR_INDEX.YELLOW
COLOR_FILLED = WD_COLOR_INDEX.BRIGHT_GREEN

# Word limita el nombre de un marcador a 40 caracteres sin espacios
BOOKMARK_MAX_LEN = 40
BOOKMARK_PREFIX = "V_"
# Marcador de título: se numera por su posición de párrafo, que es lo que
# identifica al heading en el índice que consume el frontend
HEADING_PREFIX = "H_"


def _iter_cell_paragraphs(cell) -> Iterator:
    for para in cell.paragraphs:
        yield para
    for table in cell.tables:
        yield from _iter_table_paragraphs(table)


def _iter_table_paragraphs(table) -> Iterator:
    for row in table.rows:
        for cell in row.cells:
            yield from _iter_cell_paragraphs(cell)


def _iter_all_paragraphs(doc: Document) -> Iterator:
    """Cuerpo, tablas (incluidas anidadas), encabezados y pies."""
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        yield from _iter_table_paragraphs(table)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for para in container.paragraphs:
                yield para
            for table in container.tables:
                yield from _iter_table_paragraphs(table)


def _split_segments(text: str) -> List[Tuple[str, str]]:
    """Parte el texto en tramos (contenido, key_o_vacío) según los [PLACEHOLDER]."""
    segments: List[Tuple[str, str]] = []
    pos = 0
    for match in VAR_RE.finditer(text):
        if match.start() > pos:
            segments.append((text[pos:match.start()], ""))
        segments.append((match.group(0), match.group(1)))
        pos = match.end()
    if pos < len(text):
        segments.append((text[pos:], ""))
    return segments


def _bookmark_name(key: str, used: set) -> str:
    """Nombre de marcador válido para Word, único dentro del documento."""
    safe = re.sub(r"[^A-Za-z0-9_]", "_", key)[: BOOKMARK_MAX_LEN - len(BOOKMARK_PREFIX)]
    name = f"{BOOKMARK_PREFIX}{safe}"

    if name in used:
        suffix = 2
        while f"{name[: BOOKMARK_MAX_LEN - 2]}_{suffix}" in used:
            suffix += 1
        name = f"{name[: BOOKMARK_MAX_LEN - 2]}_{suffix}"

    used.add(name)
    return name


def _wrap_in_bookmark(run, name: str, bookmark_id: int) -> None:
    """Envuelve el run en un marcador para poder saltar hasta él desde el editor."""
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)

    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))

    run._element.addprevious(start)
    run._element.addnext(end)


def _mark_paragraph(para, values: dict, ctx: dict) -> int:
    """
    Reescribe los runs del párrafo separando cada placeholder en su propio run
    resaltado, y pone un marcador en la primera aparición de cada variable.
    Devuelve cuántas variables marcó.

    Los runs se aplanan al formato del primero — mismo criterio que usa
    `filler._replace_vars_in_paragraph` al rellenar el documento.
    """
    if not para.runs:
        return 0

    full_text = "".join(run.text for run in para.runs)
    if "[" not in full_text:
        return 0

    segments = _split_segments(full_text)
    if not any(key for _, key in segments):
        return 0

    base = para.runs[0]
    base_rpr = base._element.find(qn("w:rPr"))
    base_rpr_copy = copy.deepcopy(base_rpr) if base_rpr is not None else None

    # Vaciar el primer run y eliminar el resto: los segmentos se re-crean debajo
    base.text = ""
    for run in para.runs[1:]:
        run._element.getparent().remove(run._element)

    marked = 0
    for text, key in segments:
        value = (values.get(key) or "").strip() if key else ""
        # Con valor se muestra el valor; sin valor, el propio [PLACEHOLDER]
        run = para.add_run(value if (key and value) else text)
        if base_rpr_copy is not None:
            run._element.insert(0, copy.deepcopy(base_rpr_copy))
        if key:
            run.font.highlight_color = COLOR_FILLED if value else COLOR_EMPTY
            marked += 1

            # Solo la primera aparición lleva marcador: es a donde salta el editor
            if key not in ctx["bookmarks"]:
                name = _bookmark_name(key, ctx["used"])
                ctx["bookmarks"][key] = name
                ctx["next_id"] += 1
                _wrap_in_bookmark(run, name, ctx["next_id"])

    return marked


def heading_bookmark(para_index: int) -> str:
    """Nombre del marcador de un título, derivado de su posición de párrafo."""
    return f"{HEADING_PREFIX}{para_index}"


def _mark_headings(doc: Document, ctx: dict) -> int:
    """
    Pone un marcador en cada título del cuerpo. Sin esto el editor no puede
    saltar a un título: buscarlo por texto cae en la tabla de contenidos, que
    repite los mismos textos al principio del documento.
    """
    marked = 0
    for index, para in enumerate(doc.paragraphs):
        if _get_heading_level(para) is None:
            continue
        if not para.text.strip() or not para.runs:
            continue

        name = heading_bookmark(index)
        ctx["next_id"] += 1
        _wrap_in_bookmark(para.runs[0], name, ctx["next_id"])
        ctx["headings"][index] = name
        marked += 1

    return marked


def mark_variables(
    docx_buffer: bytes, values: dict | None = None
) -> Tuple[bytes, int, Dict[str, str], Dict[int, str]]:
    """
    Devuelve (docx marcado, variables resaltadas, marcadores de variables,
    marcadores de títulos).

    `values` opcional: keys con valor se pintan en verde y el resto en amarillo.
    Los marcadores permiten que el editor salte a una variable o a un título.
    """
    doc = Document(BytesIO(docx_buffer))
    values = values or {}

    # Los ids arrancan altos para no chocar con los marcadores propios del docx
    ctx = {"bookmarks": {}, "headings": {}, "used": set(), "next_id": 90000}

    # Primero los títulos: marcar variables reescribe runs y desplaza posiciones
    _mark_headings(doc, ctx)

    total = 0
    for para in _iter_all_paragraphs(doc):
        total += _mark_paragraph(para, values, ctx)

    out = BytesIO()
    doc.save(out)
    return out.getvalue(), total, ctx["bookmarks"], ctx["headings"]


def list_bookmarks(docx_buffer: bytes) -> Dict[str, str]:
    """Mapa key → nombre de marcador de variable, sin devolver el documento."""
    _, _, bookmarks, _ = mark_variables(docx_buffer)
    return bookmarks
