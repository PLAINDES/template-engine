# app/services/sections/section_builder.py
from typing import List, Dict
from app.services.sections.section_extractor import extract_sections
from app.services.filler import fill_document


def build_section_document(
    docx_buffer:          bytes,
    selected_para_indexes: List[int],
    variables:            Dict[str, str],
    tablas:               List[dict],
    imagenes:             List[dict],
) -> bytes:
    """
    Orquesta el proceso completo:
    1. Extrae las secciones seleccionadas del .docx original
    2. Rellena variables, tablas e imágenes en el resultado

    Separamos estos dos pasos para mantener cada servicio
    con una sola responsabilidad.
    """
    # Paso 1: extraer secciones preservando formato
    extracted = extract_sections(docx_buffer, selected_para_indexes)

    # Paso 2: rellenar variables en el doc ya recortado
    if variables or tablas or imagenes:
        extracted = fill_document(extracted, variables, tablas, imagenes)

    return extracted