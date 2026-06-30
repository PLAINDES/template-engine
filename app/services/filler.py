# app/services/filler.py
import re
import copy
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
from typing import Dict, List

VAR_RE = re.compile(
    r'\[([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)\]',
    re.UNICODE,
)

TABLE_KEYS = {
    "CREAR_O_AÑADIR_TABLA", "CREAR_O_ANADIR_TABLA",
    "AÑADIR_TABLA", "ANADIR_TABLA",
    "TABLA", "CREAR_TABLA", "INSERTAR_TABLA",
}

IMAGE_KEYS = {
    "IMAGEN", "IMAGEN_PEGAR", "PEGAR_IMAGEN",
    "INSERTAR_IMAGEN", "FOTO", "ANADIR_IMAGE",
}

IMAGE_PREFIXES = ("IMAGEN_", "FOTO_", "IMG_")


def _is_image_key(upper_key: str) -> bool:
    if upper_key in IMAGE_KEYS:
        return True
    for prefix in IMAGE_PREFIXES:
        if upper_key.startswith(prefix) and upper_key not in IMAGE_KEYS:
            return True
    return False


def fill_document(
    docx_buffer: bytes,
    variables:   Dict[str, str],
    tablas:      List[Dict],
    imagenes:    List[Dict],
    bloques:     List[Dict] = [],
    replacements: List[Dict] = [],
) -> bytes:
    print(f"[filler] tablas recibidas: {len(tablas)}")
    for t in tablas:
        print(f"  tabla: para_idx={t.get('paragraph_index')} headers={t.get('headers')}")
    print(f"[filler] imagenes recibidas: {len(imagenes)}")
    for i in imagenes:
        print(f"  imagen: para_idx={i.get('paragraph_index')} key={i.get('minio_key')}")
    print(f"[filler] bloques recibidos: {len(bloques)}")
    for b in bloques:
        print(f"  bloque tipo={b.get('tipo')} items={len(b.get('items', []))}")
    doc = Document(BytesIO(docx_buffer))

    # 0. Bloques repetibles (antes de variables para que las replique también)
    if bloques:
        _expand_blocks(doc, bloques)

    # 1. Variables de texto
    for para in doc.paragraphs:
        _replace_vars_in_paragraph(para, variables)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_vars_in_paragraph(para, variables)
    for section in doc.sections:
        for para in section.header.paragraphs:
            _replace_vars_in_paragraph(para, variables)
        for para in section.footer.paragraphs:
            _replace_vars_in_paragraph(para, variables)

    # 2. Reemplazos de texto (mejoras de IA)
    if replacements:
        print(f"[filler] reemplazos de texto: {len(replacements)}")
        for repl in replacements:
            orig = repl.get("originalText", "")
            nuevo = repl.get("newText", "")
            if orig and nuevo:
                _apply_text_replacement(doc, orig, nuevo)

    # 3. Tablas
    tablas_sorted = sorted(tablas, key=lambda t: t.get("paragraph_index", 0), reverse=True)
    for tabla_data in tablas_sorted:
        _insert_table_at_placeholder(doc, tabla_data)

    # 3. Imágenes
    imagenes_sorted = sorted(imagenes, key=lambda i: i.get("paragraph_index", 0), reverse=True)
    for imagen_data in imagenes_sorted:
        _insert_image_at_placeholder(doc, imagen_data)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def _replace_vars_in_paragraph(para, variables: Dict[str, str]):
    if not para.runs:
        return
    full_text = "".join(run.text for run in para.runs)
    if not VAR_RE.search(full_text):
        return
    new_text = full_text
    for key, value in variables.items():
        new_text = new_text.replace(f"[{key}]", str(value))
    if new_text == full_text:
        return
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""


def _add_caption_paragraph(doc: Document, target_para, text: str, align: str = "center") -> None:
    """Inserta un párrafo de caption (título o pie) junto al target_para."""
    if not text or not text.strip():
        return
    caption_para = doc.add_paragraph()
    run = caption_para.add_run(text.strip())
    run.font.size    = Pt(9)
    run.font.italic  = True
    run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    alignment_map = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left":   WD_ALIGN_PARAGRAPH.LEFT,
        "right":  WD_ALIGN_PARAGRAPH.RIGHT,
    }
    caption_para.alignment = alignment_map.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    # Mover el párrafo al lugar correcto en el XML
    target_para._element.addnext(caption_para._element)


def _insert_table_at_placeholder(doc: Document, tabla_data: Dict):
    para_idx      = tabla_data.get("paragraph_index", 0)
    headers       = tabla_data.get("headers", [])
    rows_data     = tabla_data.get("rows", [])
    titulo        = tabla_data.get("titulo", "")
    pie           = tabla_data.get("pie", "")

    if not headers and not rows_data:
        return

    paragraphs = doc.paragraphs
    print(f"[insert_table] buscando párrafo {para_idx}")
    print(f"[insert_table] total párrafos en doc: {len(paragraphs)}")
    if para_idx < len(paragraphs):
        text = "".join(run.text for run in paragraphs[para_idx].runs)
        print(f"[insert_table] texto en párrafo {para_idx}: '{text}'")
    if para_idx >= len(paragraphs):
        return

    target_para = paragraphs[para_idx]
    full_text   = "".join(run.text for run in target_para.runs)
    has_placeholder = any(k in full_text for k in TABLE_KEYS)

    if not has_placeholder:
        for para in paragraphs:
            t = "".join(run.text for run in para.runs)
            if any(k in t for k in TABLE_KEYS):
                target_para = para
                break
        else:
            return

    # Limpiar placeholder
    for run in target_para.runs:
        run.text = ""

    num_cols = max(len(headers), max((len(r) for r in rows_data), default=0))
    if num_cols == 0:
        return

    num_rows = len(rows_data) + (1 if headers else 0)
    table    = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = "Table Grid"

    row_offset = 0
    if headers:
        header_row = table.rows[0]
        for ci, header_text in enumerate(headers):
            cell = header_row.cells[ci]
            cell.text = str(header_text)
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
        row_offset = 1

    for ri, row_data in enumerate(rows_data):
        table_row = table.rows[ri + row_offset]
        for ci, cell_val in enumerate(row_data):
            if ci < num_cols:
                table_row.cells[ci].text = str(cell_val)

    # Insertar tabla en el XML
    target_para._element.addnext(table._element)

    # Pie debajo de la tabla
    if pie and pie.strip():
        pie_para = doc.add_paragraph()
        run = pie_para.add_run(pie.strip())
        run.font.size    = Pt(9)
        run.font.italic  = True
        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        pie_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        table._element.addnext(pie_para._element)

    # Título encima de la tabla
    if titulo and titulo.strip():
        titulo_para = doc.add_paragraph()
        run = titulo_para.add_run(titulo.strip())
        run.font.size   = Pt(10)
        run.font.bold   = True
        titulo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        target_para._element.addnext(titulo_para._element)


def _insert_image_at_placeholder(doc: Document, imagen_data: Dict):
    from app.utils.minio_client import download_from_minio

    para_idx     = imagen_data.get("paragraph_index", 0)
    minio_key    = imagen_data.get("minio_key", "")
    var_key      = imagen_data.get("key", "")
    width_inches = imagen_data.get("width_inches", 4.0)
    titulo       = imagen_data.get("titulo", "")
    pie          = imagen_data.get("pie", "")
    descripcion  = imagen_data.get("descripcion", "")

    if not minio_key:
        return

    paragraphs  = doc.paragraphs
    target_para = None

    # Match por nombre de variable (imágenes con nombre único)
    if var_key:
        placeholder = f"[{var_key}]"
        for para in paragraphs:
            t = "".join(run.text for run in para.runs)
            if placeholder in t:
                target_para = para
                break

    # Fallback: match por paragraph_index (imágenes genéricas)
    if not target_para and para_idx < len(paragraphs):
        t = "".join(run.text for run in paragraphs[para_idx].runs)
        if any(k in t.upper() for k in IMAGE_KEYS) or any(t.upper().startswith(p) for p in IMAGE_PREFIXES):
            target_para = paragraphs[para_idx]

    # Fallback: buscar cualquier placeholder de imagen
    if not target_para:
        for para in paragraphs:
            t = "".join(run.text for run in para.runs)
            if _is_image_key(t.strip().strip("[]").upper()):
                target_para = para
                break

    if not target_para:
        return

    # Limpiar placeholder
    for run in target_para.runs:
        run.text = ""

    try:
        image_bytes  = download_from_minio(minio_key)
        image_buffer = BytesIO(image_bytes)
    except Exception as e:
        print(f"Error descargando imagen [{minio_key}]: {e}")
        return

    # Insertar imagen
    run = target_para.add_run()
    run.add_picture(image_buffer, width=Inches(width_inches))
    target_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Pie debajo de la imagen
    if pie and pie.strip():
        pie_para = doc.add_paragraph()
        run = pie_para.add_run(pie.strip())
        run.font.size      = Pt(9)
        run.font.italic    = True
        run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        pie_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        target_para._element.addnext(pie_para._element)

    # Descripción debajo de la imagen (párrafo normal, texto generado por IA)
    if descripcion and descripcion.strip():
        desc_para = doc.add_paragraph()
        run = desc_para.add_run(descripcion.strip())
        run.font.size = Pt(11)
        desc_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Insertar después del pie si existe, si no después de la imagen
        anchor = target_para._element
        if pie and pie.strip():
            # Buscar el pie que acabamos de insertar
            next_elem = target_para._element.getnext()
            if next_elem is not None:
                anchor = next_elem
        anchor.addnext(desc_para._element)

    # Título encima de la imagen — addprevious para que quede antes, no después
    if titulo and titulo.strip():
        titulo_para = doc.add_paragraph()
        run = titulo_para.add_run(titulo.strip())
        run.font.size  = Pt(10)
        run.font.bold  = True
        titulo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        target_para._element.addprevious(titulo_para._element)


def _apply_text_replacement(doc: Document, original: str, replacement: str) -> None:
    """Busca el texto original en los párrafos del documento y lo reemplaza."""
    replaced = False
    for para in doc.paragraphs:
        full_text = "".join(run.text for run in para.runs)
        if original in full_text:
            new_text = full_text.replace(original, replacement)
            if para.runs:
                para.runs[0].text = new_text
                for run in para.runs[1:]:
                    run.text = ""
            replaced = True
            print(f"[filler] ✓ Reemplazo aplicado ({len(original)} → {len(replacement)} chars)")
            break
    if not replaced:
        # Buscar en celdas de tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        full_text = "".join(run.text for run in para.runs)
                        if original in full_text:
                            new_text = full_text.replace(original, replacement)
                            if para.runs:
                                para.runs[0].text = new_text
                                for run in para.runs[1:]:
                                    run.text = ""
                            replaced = True
                            print(f"[filler] ✓ Reemplazo en tabla ({len(original)} → {len(replacement)} chars)")
                            break
                    if replaced:
                        break
                if replaced:
                    break
            if replaced:
                break
    if not replaced:
        print(f"[filler] ✗ No se encontró texto para reemplazar: '{original[:50]}...'")


def _find_preceding_heading_elem(body, start_elem):
    """
    Busca el párrafo con estilo Heading-1 más cercano que precede a start_elem.
    Retorna una copia profunda del elemento XML, o None si no encuentra.
    Detecta estilos en cualquier idioma: Heading1, Ttulo1, h1, H1, etc.
    """
    preceding = []
    for elem in list(body):
        if elem is start_elem:
            break
        preceding.append(elem)

    for elem in reversed(preceding):
        if not elem.tag.endswith('}p'):
            continue
        pPr = elem.find(qn('w:pPr'))
        if pPr is None:
            continue
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            continue
        style_val = pStyle.get(qn('w:val'), '') or ''
        # Captura Heading1, Ttulo1 (ES), berschrift1 (DE), h1, H1, etc.
        if re.search(r'(?i)[a-z]1$', style_val):
            return copy.deepcopy(elem)
    return None


def _expand_blocks(doc: Document, bloques: List[Dict]) -> None:
    """
    Detecta TODOS los pares [BLOQUE_X_START]/[BLOQUE_X_END] en el documento
    en un único escaneo (antes de modificar nada), luego expande cada par.

    Para bloques con más de 1 ítem, replica automáticamente el heading H1 que
    precede al bloque al inicio de cada repetición (a partir de la 2ª), de modo
    que cada empresa/ítem quede bajo su propio encabezado de sección.
    """
    body = doc.element.body

    for bloque in bloques:
        tipo  = bloque.get("tipo", "").upper()
        items = bloque.get("items", [])
        if not items:
            continue

        start_marker = f"[BLOQUE_{tipo}_START]"
        end_marker   = f"[BLOQUE_{tipo}_END]"

        # ── Paso 1: encontrar TODOS los pares usando referencias a elementos ──
        pairs = []          # lista de (start_elem, end_elem, [template_elems])
        current_start = None

        for child in list(body):
            if not child.tag.endswith('}p'):
                continue
            text = "".join(t.text or "" for t in child.iter() if t.tag.endswith('}t'))
            if start_marker in text and current_start is None:
                current_start = child
            elif end_marker in text and current_start is not None:
                # Recolectar los elementos entre start y end
                collecting = False
                template_elems = []
                for elem in list(body):
                    if elem is current_start:
                        collecting = True
                        continue
                    if elem is child:
                        break
                    if collecting:
                        template_elems.append(copy.deepcopy(elem))
                pairs.append((current_start, child, template_elems))
                current_start = None

        if not pairs:
            print(f"[expand_blocks] No se encontraron pares para tipo={tipo}")
            continue

        print(f"[expand_blocks] tipo={tipo} pares={len(pairs)} items={len(items)}")

        # Heading H1 que precede al bloque — se replica para ítems 2, 3, …
        preceding_heading = (
            _find_preceding_heading_elem(body, pairs[0][0])
            if len(items) > 1
            else None
        )
        if preceding_heading is not None:
            heading_preview = "".join(
                t.text or "" for t in preceding_heading.iter() if t.tag.endswith('}t')
            )
            print(f"[expand_blocks] heading previo detectado: '{heading_preview[:80]}'")

        # ── Paso 2: expandir cada par (los pares no se solapan, orden seguro) ──
        for start_elem, end_elem, template_elements in pairs:
            for item_idx, item in enumerate(reversed(items)):
                is_first_item = (item_idx == len(items) - 1)

                # Insertar contenido del bloque en orden correcto (reversed + addnext)
                for elem in reversed(template_elements):
                    new_elem = copy.deepcopy(elem)
                    _replace_vars_in_xml_element(new_elem, item)
                    end_elem.addnext(new_elem)

                # Para ítems 2+: insertar copia del heading justo antes del bloque
                if not is_first_item and preceding_heading is not None:
                    heading_copy = copy.deepcopy(preceding_heading)
                    _replace_vars_in_xml_element(heading_copy, item)
                    end_elem.addnext(heading_copy)

            # Eliminar desde start hasta end inclusive
            to_remove = []
            collecting = False
            for child in list(body):
                if child is start_elem:
                    collecting = True
                if collecting:
                    to_remove.append(child)
                if child is end_elem:
                    break
            for elem in to_remove:
                try:
                    body.remove(elem)
                except Exception:
                    pass


def _replace_vars_in_xml_element(element, variables: Dict[str, str]) -> None:
    """Reemplaza variables [KEY] en nodos <w:t> del elemento XML."""
    for node in element.iter():
        # Solo <w:t> tiene text editable; CT_P y otros tienen text como property read-only
        if node.tag.endswith('}t') and node.text:
            for key, value in variables.items():
                node.text = node.text.replace(f"[{key}]", str(value))