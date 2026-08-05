"""Pone un marcador invisible alrededor del texto de cada comentario.

El problema que resuelve: para corregir el trozo que alguien comentó hay que
poder seleccionarlo, y ONLYOFFICE no da ninguna forma pública de pedirle a un
comentario cuál es su rango. La alternativa era buscar el texto citado, que
falla en cuanto el mismo párrafo aparece dos veces en el informe: no hay manera
de saber cuál de los dos señaló el revisor.

Un marcador sí se puede seleccionar por nombre (`GetBookmarkRange`), y en Word
no se ve. Así que a cada comentario se le pone uno con un identificador propio,
alrededor exactamente del mismo trozo que abarca el comentario.

Es idempotente: si el comentario ya tiene su marcador, se respeta y se devuelve
el mismo nombre. Sellar dos veces el mismo documento no lo llena de marcadores.
"""

from __future__ import annotations

import re
import uuid
from io import BytesIO
from typing import Dict, List, Tuple
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

RUTA_DOCUMENTO = "word/document.xml"

#: Prefijo de los marcadores que pone este módulo. Sirve para reconocerlos y
#: para no tocar los que pone el marcador de variables.
PREFIJO = "cmt-"

_INICIO_COMENTARIO = re.compile(r'<w:commentRangeStart w:id="(\d+)"\s*/>')
_FIN_COMENTARIO = re.compile(r'<w:commentRangeEnd w:id="(\d+)"\s*/>')
_INICIO_MARCADOR = re.compile(r'<w:bookmarkStart[^>]*w:id="(\d+)"[^>]*/>')


def _siguiente_id_de_marcador(xml: str) -> int:
    """Un id de marcador que no choque con los que ya tiene el documento."""
    usados = [int(n) for n in _INICIO_MARCADOR.findall(xml)]
    # Word reserva el 0 para el marcador _GoBack
    return max(usados, default=0) + 1


def _marcador_ya_puesto(xml: str, inicio: int, fin: int) -> str | None:
    """Nombre del marcador de este módulo que ya envuelve el rango, si lo hay.

    Se mira solo el hueco entre el principio y el final del rango del
    comentario: un marcador que empiece ahí dentro es el nuestro de una pasada
    anterior.
    """
    trozo = xml[inicio:fin]
    encontrado = re.search(rf'<w:bookmarkStart[^>]*w:name="({PREFIJO}[^"]+)"', trozo)
    return encontrado.group(1) if encontrado else None


def sellar_comentarios(docx_buffer: bytes) -> Tuple[bytes, Dict[str, str]]:
    """Devuelve el documento con un marcador por comentario, y el mapa id → nombre.

    El mapa va del `w:id` del comentario en este documento al nombre del
    marcador. Ese `w:id` cambia cuando ONLYOFFICE vuelve a guardar; el nombre
    del marcador no, y es lo que hay que usar para volver a encontrar el trozo.
    """
    try:
        entrada = ZipFile(BytesIO(docx_buffer))
    except BadZipFile as e:
        raise ValueError(f"El archivo no es un .docx válido: {e}") from e

    with entrada:
        if RUTA_DOCUMENTO not in entrada.namelist():
            raise ValueError(f"El .docx no tiene {RUTA_DOCUMENTO}")

        xml = entrada.read(RUTA_DOCUMENTO).decode("utf-8")
        nuevo_xml, marcadores = _sellar_xml(xml)

        if not marcadores:
            # Sin comentarios no se reescribe nada: devolver el mismo archivo
            # evita cambiarle el etag y que el editor se recargue sin motivo
            return docx_buffer, {}

        salida = BytesIO()
        with ZipFile(salida, "w", ZIP_DEFLATED) as z:
            for item in entrada.infolist():
                datos = (
                    nuevo_xml.encode("utf-8")
                    if item.filename == RUTA_DOCUMENTO
                    else entrada.read(item.filename)
                )
                z.writestr(item, datos)

        return salida.getvalue(), marcadores


def _sellar_xml(xml: str) -> Tuple[str, Dict[str, str]]:
    """Inserta los marcadores que falten. Devuelve el XML y el mapa id → nombre."""
    marcadores: Dict[str, str] = {}

    # Se recorre de atrás hacia delante: cada inserción desplaza el texto que
    # viene después, y al ir en orden inverso las posiciones ya calculadas
    # siguen siendo válidas.
    trabajos: List[Tuple[int, int, str]] = []

    for inicio in _INICIO_COMENTARIO.finditer(xml):
        ident = inicio.group(1)
        fin = _buscar_fin(xml, ident, inicio.end())
        if fin is None:
            # Un comentario sin cierre es un documento estropeado; se salta en
            # vez de romper el sellado entero de los demás
            continue

        ya_puesto = _marcador_ya_puesto(xml, inicio.end(), fin.start())
        if ya_puesto is not None:
            marcadores[ident] = ya_puesto
            continue

        nombre = f"{PREFIJO}{uuid.uuid4().hex[:16]}"
        marcadores[ident] = nombre
        trabajos.append((inicio.end(), fin.start(), nombre))

    if not trabajos:
        return xml, marcadores

    siguiente_id = _siguiente_id_de_marcador(xml)
    for pos_inicio, pos_fin, nombre in sorted(trabajos, reverse=True):
        id_marcador = siguiente_id
        siguiente_id += 1
        xml = (
            xml[:pos_inicio]
            + f'<w:bookmarkStart w:id="{id_marcador}" w:name="{nombre}"/>'
            + xml[pos_inicio:pos_fin]
            + f'<w:bookmarkEnd w:id="{id_marcador}"/>'
            + xml[pos_fin:]
        )

    return xml, marcadores


def _buscar_fin(xml: str, ident: str, desde: int):
    """El `commentRangeEnd` que cierra este comentario."""
    for fin in _FIN_COMENTARIO.finditer(xml, desde):
        if fin.group(1) == ident:
            return fin
    return None
