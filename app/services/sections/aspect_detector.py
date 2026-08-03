# app/services/sections/aspect_detector.py
"""
Deduce la división en aspectos a partir del índice del documento.

Un documento de inversión trae sus capítulos como headings de nivel 1
("Aspectos Generales", "Identificación", "Formulación", "Evaluación"), así que
cada H1 se convierte en un aspecto del editor. Evita tener que mapear los
capítulos a mano en el AspectMapper cada vez que se sube una plantilla.
"""
import re
from typing import Any, Dict, List

from app.models.schemas import HeadingItem

# "1. Aspectos Generales", "1.1 Identificación", "I. Evaluación" → se quita el prefijo
NUMBERING_RE = re.compile(r"^\s*(?:\d+(?:\.\d+)*[.\)]?|[IVXLC]+[.\)])\s+", re.IGNORECASE)

# Colores del stepper, en orden de aparición
PALETTE = ["#843c0c", "#1d4ed8", "#15803d", "#b45309", "#7c3aed", "#be123c"]


def _clean_title(text: str) -> str:
    """Quita la numeración del heading: Word la genera sola en el documento."""
    return NUMBERING_RE.sub("", text).strip() or text.strip()


def _pick_level(headings: List[HeadingItem]) -> List[HeadingItem]:
    """
    Devuelve los headings que definen los capítulos.

    Normalmente son los H1. Si el documento no tiene H1 (o solo tiene uno que
    envuelve todo), se baja un nivel para no proponer un único aspecto.
    """
    level1 = [h for h in headings if h.level == 1]

    if len(level1) > 1:
        return level1

    if len(level1) == 1 and level1[0].children:
        return [c for c in level1[0].children if c.level == 2] or level1

    if not level1:
        return [h for h in headings if h.level == 2]

    return level1


def detect_aspects(headings: List[HeadingItem]) -> List[Dict[str, Any]]:
    """
    Convierte el árbol de headings en la configuración de aspectos que espera
    el backend: un aspecto por capítulo, con el capítulo como única sección.
    """
    chapters = _pick_level(headings)

    aspects: List[Dict[str, Any]] = []
    for order, chapter in enumerate(chapters):
        title = _clean_title(chapter.text)
        aspects.append(
            {
                "name": title,
                "order": order,
                "color": PALETTE[order % len(PALETTE)],
                "sections": [
                    {
                        "h1Index": chapter.index,
                        "h1Text": chapter.text,
                        "order": 0,
                    }
                ],
                # Solo informativo: cuántas subsecciones cuelgan del capítulo
                "subsections": len(chapter.children or []),
            }
        )

    return aspects
