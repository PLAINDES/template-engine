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
from hashlib import sha1
from io import BytesIO
from typing import Dict, List, Tuple
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

RUTA_DOCUMENTO = "word/document.xml"
RUTA_COMENTARIOS = "word/comments.xml"

#: Prefijo de los marcadores que pone este módulo. Sirve para reconocerlos y
#: para no tocar los que pone el marcador de variables.
PREFIJO = "cmt-"

_INICIO_COMENTARIO = re.compile(r'<w:commentRangeStart w:id="(\d+)"\s*/>')
_FIN_COMENTARIO = re.compile(r'<w:commentRangeEnd w:id="(\d+)"\s*/>')
_INICIO_MARCADOR = re.compile(r'<w:bookmarkStart[^>]*w:id="(\d+)"[^>]*/>')

# Comentarios de word/comments.xml: cabecera con los atributos y cuerpo con el
# texto. De ahí salen las señas con las que se nombra cada marcador.
_COMENTARIO = re.compile(r"<w:comment\b([^>]*)>(.*?)</w:comment>", re.DOTALL)
_ATRIBUTO_ID = re.compile(r'w:id="(\d+)"')
_ATRIBUTO_AUTOR = re.compile(r'w:author="([^"]*)"')
_ATRIBUTO_FECHA = re.compile(r'w:date="([^"]*)"')
_TEXTO = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.DOTALL)


def _siguiente_id_de_marcador(xml: str) -> int:
    """Un id de marcador que no choque con los que ya tiene el documento."""
    usados = [int(n) for n in _INICIO_MARCADOR.findall(xml)]
    # Word reserva el 0 para el marcador _GoBack
    return max(usados, default=0) + 1


def _marcador_ya_puesto(xml: str, inicio: int, fin: int, nombre: str) -> bool:
    """Si el marcador de ESTE comentario ya está dentro de su rango.

    Se busca por su nombre exacto y no por el prefijo: dos comentarios pueden
    solaparse —tres revisores citando el mismo párrafo—, y con el prefijo a
    secas el rango de uno encontraba el marcador del otro, se daba por sellado
    y se quedaba sin marcador propio.
    """
    trozo = xml[inicio:fin]
    return re.search(rf'<w:bookmarkStart[^>]*w:name="{re.escape(nombre)}"', trozo) is not None


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
        nuevo_xml, marcadores = _sellar_xml(xml, _señas_de_los_comentarios(entrada))

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


def _sellar_xml(xml: str, señas: Dict[str, str]) -> Tuple[str, Dict[str, str]]:
    """Inserta los marcadores que falten. Devuelve el XML y el mapa id → nombre.

    `señas` lleva, por `w:id` de comentario, el autor, la fecha y el texto: es
    de ahí de donde sale el nombre del marcador.
    """
    marcadores: Dict[str, str] = {}
    trabajos: List[Tuple[int, int, str]] = []

    for inicio in _INICIO_COMENTARIO.finditer(xml):
        ident = inicio.group(1)
        fin = _buscar_fin(xml, ident, inicio.end())
        if fin is None:
            # Un comentario sin cierre es un documento estropeado; se salta en
            # vez de romper el sellado entero de los demás
            continue

        nombre = _nombre_del_marcador(ident, señas)
        marcadores[ident] = nombre
        if _marcador_ya_puesto(xml, inicio.end(), fin.start(), nombre):
            continue

        trabajos.append((inicio.end(), fin.start(), nombre))

    if not trabajos:
        return xml, marcadores

    # Cada tag se inserta por su posición en el XML original, de atrás hacia
    # delante. Antes se insertaba cada par de una vez —rebanando el texto del
    # rango entero— y con comentarios solapados las posiciones del siguiente
    # par quedaban desfasadas y las tags caían en medio de otra tag: el
    # documento entero salía ilegible. Los marcadores no exigen anidarse bien
    # en OOXML, así que insertar por posición vale para cualquier solape.
    siguiente_id = _siguiente_id_de_marcador(xml)
    inserciones: List[Tuple[int, str]] = []
    for pos_inicio, pos_fin, nombre in trabajos:
        id_marcador = siguiente_id
        siguiente_id += 1
        inserciones.append(
            (pos_inicio, f'<w:bookmarkStart w:id="{id_marcador}" w:name="{nombre}"/>')
        )
        inserciones.append((pos_fin, f'<w:bookmarkEnd w:id="{id_marcador}"/>'))

    for posicion, tag in sorted(inserciones, key=lambda i: i[0], reverse=True):
        xml = xml[:posicion] + tag + xml[posicion:]

    return xml, marcadores


def _nombre_del_marcador(ident: str, señas: Dict[str, str]) -> str:
    """Nombre estable para el marcador de un comentario.

    Antes era aleatorio, y eso lo hacía inservible: mientras alguien tiene el
    documento abierto, su copia en memoria no lleva el marcador, así que al
    volver a guardar lo borra y el siguiente sellado inventaba otro nombre. El
    que ya se le había entregado al navegador dejaba de existir, y corregir o
    resaltar ese comentario fallaba con "marcador perdido".

    Derivado del autor, la fecha y el texto del comentario, el nombre sale
    siempre igual: se puede volver a sellar las veces que haga falta y el que
    tenga el navegador en la mano sigue valiendo.
    """
    seña = señas.get(ident)
    if seña is None:
        # Sin comments.xml no hay de dónde derivarlo. Pasa con un .docx
        # estropeado, no con uno normal, y sellar con un nombre inventado seria
        # volver al problema de arriba, asi que se dice claro.
        raise ValueError(
            f"El comentario {ident} no aparece en word/comments.xml; "
            "no se le puede poner un marcador estable"
        )
    return f"{PREFIJO}{sha1(seña.encode('utf-8')).hexdigest()[:16]}"


def _señas_de_los_comentarios(entrada: ZipFile) -> Dict[str, str]:
    """Por cada `w:id` de comentario, con qué identificarlo entre guardados.

    El `w:id` no vale: ONLYOFFICE lo renumera al guardar. El autor, la fecha de
    creación y el texto sí se quedan como estaban.
    """
    if RUTA_COMENTARIOS not in entrada.namelist():
        return {}

    xml = entrada.read(RUTA_COMENTARIOS).decode("utf-8")
    señas: Dict[str, str] = {}
    for comentario in _COMENTARIO.finditer(xml):
        cabecera, cuerpo = comentario.group(1), comentario.group(2)
        ident = _ATRIBUTO_ID.search(cabecera)
        if ident is None:
            continue
        autor = _ATRIBUTO_AUTOR.search(cabecera)
        fecha = _ATRIBUTO_FECHA.search(cabecera)
        texto = "".join(_TEXTO.findall(cuerpo))
        señas[ident.group(1)] = "␟".join(
            [
                autor.group(1) if autor else "",
                fecha.group(1) if fecha else "",
                texto,
            ]
        )
    return señas


def _buscar_fin(xml: str, ident: str, desde: int):
    """El `commentRangeEnd` que cierra este comentario."""
    for fin in _FIN_COMENTARIO.finditer(xml, desde):
        if fin.group(1) == ident:
            return fin
    return None
