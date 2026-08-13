"""Lectura de los comentarios que lleva dentro un .docx.

Un comentario de Word no está en un sitio, está repartido en tres:

  word/document.xml          dónde empieza y acaba el trozo comentado, con
                             <w:commentRangeStart> y <w:commentRangeEnd>
  word/comments.xml          el texto del comentario, quién lo escribió y cuándo
  word/commentsExtended.xml  qué comentario responde a cuál, y si está resuelto

python-docx no expone nada de esto, así que se lee el XML directamente. Se hace
aquí, en el servidor y sobre el archivo guardado, en vez de fiarse de la lista
que manda el navegador: el documento es la fuente de verdad de sus propios
comentarios, y lo que llega del cliente puede venir manipulado o incompleto.

El enlace entre `comments.xml` y `commentsExtended.xml` no es el id del
comentario sino el `w14:paraId` del último párrafo del comentario. Es rebuscado,
pero es como lo define OOXML y es lo que escribe tanto Word como OnlyOffice.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from typing import Dict, List
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

RUTA_DOCUMENTO = "word/document.xml"

#: Prefijo de los marcadores que pone `comment_sealer` sobre cada comentario
PREFIJO_DE_MARCADOR = "cmt-"
RUTA_COMENTARIOS = "word/comments.xml"
RUTA_COMENTARIOS_EXT = "word/commentsExtended.xml"


@dataclass
class Comentario:
    """Un comentario del documento, ya resuelto con todas sus partes."""

    id: str
    autor: str
    iniciales: str
    fecha: str
    texto: str
    """Trozo del documento sobre el que se puso el comentario."""
    texto_citado: str
    """El trozo comentado fue borrado con control de cambios: la cita que se
    devuelve es ese texto tachado, no texto visible del documento."""
    cita_eliminada: bool
    """Marcador invisible que envuelve ese trozo, para poder seleccionarlo."""
    marcador: str | None
    """Id del comentario al que responde, o None si abre el hilo."""
    responde_a: str | None
    resuelto: bool
    """Orden de aparición en el documento; sirve para listarlos como se leen."""
    posicion: int
    """Encabezado de nivel 1 bajo el que cae el comentario. Vacío si el
    documento no tiene encabezados antes de ese punto."""
    seccion: str
    """Encabezado más profundo bajo el que cae, dentro de esa sección."""
    subtitulo: str

    def a_dict(self) -> dict:
        return {
            "id": self.id,
            "autor": self.autor,
            "iniciales": self.iniciales,
            "fecha": self.fecha,
            "texto": self.texto,
            "textoCitado": self.texto_citado,
            "citaEliminada": self.cita_eliminada,
            "marcador": self.marcador,
            "respondeA": self.responde_a,
            "resuelto": self.resuelto,
            "posicion": self.posicion,
            "seccion": self.seccion,
            "subtitulo": self.subtitulo,
        }


@dataclass
class _Extendido:
    """Lo que aporta commentsExtended.xml, indexado por paraId."""

    resuelto: bool = False
    padre: str | None = None


@dataclass
class _Rango:
    """Texto comentado, posición, marcador y ubicación, de document.xml."""

    texto: str = ""
    posicion: int = 0
    marcador: str | None = None
    partes: List[str] = field(default_factory=list)
    partes_eliminadas: List[str] = field(default_factory=list)
    cita_eliminada: bool = False
    seccion: str = ""
    subtitulo: str = ""


def _q(nombre: str) -> str:
    """Convierte `w:id` en el nombre con espacio de nombres que usa ElementTree."""
    prefijo, local = nombre.split(":")
    return f"{{{NS[prefijo]}}}{local}"


def _texto_de(elemento: ElementTree.Element) -> str:
    """Todo el texto visible de un elemento, respetando los saltos de párrafo."""
    trozos: List[str] = []
    for hijo in elemento.iter():
        etiqueta = hijo.tag
        if etiqueta == _q("w:t"):
            trozos.append(hijo.text or "")
        elif etiqueta in (_q("w:br"), _q("w:cr")):
            trozos.append("\n")
        elif etiqueta == _q("w:tab"):
            trozos.append("\t")
    return "".join(trozos).strip()


def _nivel_de_encabezado(parrafo: ElementTree.Element) -> int | None:
    """Nivel del encabezado, o None si el párrafo no es un encabezado.

    Se mira primero `w:outlineLvl`, que es explícito y no depende del idioma.
    Cuando no está —lo habitual, porque suele venir en la definición del
    estilo y no en el párrafo— se deduce del nombre del estilo. Word en
    español escribe el styleId sin tildes ("Ttulo1"), OnlyOffice y Word en
    inglés lo escriben "Heading1", y ambos se ven en los documentos que suben
    los revisores.
    """
    propiedades = parrafo.find(_q("w:pPr"))
    if propiedades is None:
        return None

    nivel_esquema = propiedades.find(_q("w:outlineLvl"))
    if nivel_esquema is not None:
        valor = nivel_esquema.get(_q("w:val"))
        if valor is not None and valor.isdigit():
            # outlineLvl cuenta desde 0; los encabezados desde 1
            return int(valor) + 1

    estilo = propiedades.find(_q("w:pStyle"))
    if estilo is None:
        return None

    nombre = (estilo.get(_q("w:val")) or "").lower().replace(" ", "")
    for prefijo in ("heading", "ttulo", "titulo"):
        if nombre.startswith(prefijo):
            resto = nombre[len(prefijo) :]
            if resto.isdigit():
                return int(resto)

    return None


def _leer_extendidos(zipf: ZipFile) -> Dict[str, _Extendido]:
    """Resuelto y jerarquía de respuestas, indexado por paraId."""
    if RUTA_COMENTARIOS_EXT not in zipf.namelist():
        # Un documento sin respuestas ni comentarios resueltos no lleva esta
        # parte. No es un error: simplemente no hay nada que añadir.
        return {}

    raiz = ElementTree.fromstring(zipf.read(RUTA_COMENTARIOS_EXT))
    extendidos: Dict[str, _Extendido] = {}

    for ex in raiz.findall(_q("w15:commentEx")):
        para_id = ex.get(_q("w15:paraId"))
        if para_id is None:
            continue
        extendidos[para_id.upper()] = _Extendido(
            resuelto=ex.get(_q("w15:done")) in ("1", "true"),
            padre=(ex.get(_q("w15:paraIdParent")) or "").upper() or None,
        )

    return extendidos


def _leer_rangos(zipf: ZipFile) -> Dict[str, _Rango]:
    """Texto comentado de cada comentario, en orden de aparición.

    Se recorre el documento de principio a fin llevando la cuenta de qué
    comentarios están abiertos en cada momento. Un mismo trozo de texto puede
    estar dentro de varios comentarios a la vez —comentarios que se solapan—,
    así que el texto se acumula en todos los que estén abiertos.

    En el mismo recorrido se lleva el último encabezado visto, que es lo que
    permite decir en qué parte del informe cae cada comentario. Sale gratis:
    ya se estaban visitando todos los nodos en orden de lectura.
    """
    if RUTA_DOCUMENTO not in zipf.namelist():
        raise ValueError(f"El .docx no tiene {RUTA_DOCUMENTO}")

    raiz = ElementTree.fromstring(zipf.read(RUTA_DOCUMENTO))
    rangos: Dict[str, _Rango] = {}
    abiertos: set[str] = set()
    orden = 0
    seccion = ""
    subtitulo = ""

    for nodo in raiz.iter():
        etiqueta = nodo.tag

        if etiqueta == _q("w:p"):
            nivel = _nivel_de_encabezado(nodo)
            if nivel is not None:
                titulo = _texto_de(nodo)
                if titulo:
                    if nivel == 1:
                        seccion = titulo
                        # Empieza otra sección: el subtítulo de la anterior ya
                        # no describe dónde estamos
                        subtitulo = ""
                    else:
                        subtitulo = titulo

        elif etiqueta == _q("w:commentRangeStart"):
            ident = nodo.get(_q("w:id"))
            if ident is not None:
                orden += 1
                rango = rangos.setdefault(ident, _Rango())
                rango.posicion = orden
                rango.seccion = seccion
                rango.subtitulo = subtitulo
                abiertos.add(ident)

        elif etiqueta == _q("w:commentRangeEnd"):
            ident = nodo.get(_q("w:id"))
            if ident is not None:
                abiertos.discard(ident)

        elif etiqueta == _q("w:bookmarkStart") and abiertos:
            # El marcador que envuelve un comentario lo pone `comment_sealer`;
            # es lo que permite volver a seleccionar ese trozo exacto
            nombre = nodo.get(_q("w:name")) or ""
            if nombre.startswith(PREFIJO_DE_MARCADOR):
                for ident in abiertos:
                    if rangos[ident].marcador is None:
                        rangos[ident].marcador = nombre

        elif etiqueta == _q("w:t") and abiertos:
            texto = nodo.text or ""
            for ident in abiertos:
                rangos[ident].partes.append(texto)

        elif etiqueta == _q("w:delText") and abiertos:
            # Texto que alguien borró con control de cambios y todavía está en
            # el documento, tachado. No cuenta como cita normal, pero si es lo
            # único que abarca el comentario hay que conservarlo: sin él la
            # cita salía vacía y el comentario parecía puesto sobre la nada.
            texto = nodo.text or ""
            for ident in abiertos:
                rangos[ident].partes_eliminadas.append(texto)

    for rango in rangos.values():
        rango.texto = "".join(rango.partes).strip()
        if not rango.texto and rango.partes_eliminadas:
            rango.texto = "".join(rango.partes_eliminadas).strip()
            rango.cita_eliminada = True

    return rangos


def leer_comentarios(docx_buffer: bytes) -> List[dict]:
    """Comentarios del documento, en el orden en que aparecen al leerlo.

    Devuelve lista vacía si el documento no tiene ninguno, que es distinto de
    fallar: un informe sin comentarios es lo normal al empezar.
    """
    try:
        zipf = ZipFile(BytesIO(docx_buffer))
    except BadZipFile as e:
        raise ValueError(f"El archivo no es un .docx válido: {e}") from e

    with zipf:
        if RUTA_COMENTARIOS not in zipf.namelist():
            return []

        raiz = ElementTree.fromstring(zipf.read(RUTA_COMENTARIOS))
        extendidos = _leer_extendidos(zipf)
        rangos = _leer_rangos(zipf)

        # paraId del último párrafo → id del comentario. Es el enlace que usa
        # commentsExtended para decir quién responde a quién.
        id_por_para: Dict[str, str] = {}
        parrafo_final: Dict[str, str] = {}

        for comentario in raiz.findall(_q("w:comment")):
            ident = comentario.get(_q("w:id"))
            if ident is None:
                continue
            parrafos = comentario.findall(_q("w:p"))
            if not parrafos:
                continue
            para_id = parrafos[-1].get(_q("w14:paraId"))
            if para_id is None:
                continue
            id_por_para[para_id.upper()] = ident
            parrafo_final[ident] = para_id.upper()

        comentarios: List[Comentario] = []

        for comentario in raiz.findall(_q("w:comment")):
            ident = comentario.get(_q("w:id"))
            if ident is None:
                continue

            para_id = parrafo_final.get(ident)
            extendido = extendidos.get(para_id or "", _Extendido())
            rango = rangos.get(ident, _Rango())

            comentarios.append(
                Comentario(
                    id=ident,
                    autor=comentario.get(_q("w:author")) or "",
                    iniciales=comentario.get(_q("w:initials")) or "",
                    fecha=comentario.get(_q("w:date")) or "",
                    texto=_texto_de(comentario),
                    texto_citado=rango.texto,
                    cita_eliminada=rango.cita_eliminada,
                    marcador=rango.marcador,
                    responde_a=id_por_para.get(extendido.padre or ""),
                    resuelto=extendido.resuelto,
                    posicion=rango.posicion,
                    seccion=rango.seccion,
                    subtitulo=rango.subtitulo,
                )
            )

        return [c.a_dict() for c in _ordenar_por_hilos(comentarios)]


def _ordenar_por_hilos(comentarios: List[Comentario]) -> List[Comentario]:
    """Deja cada comentario seguido de sus respuestas, en orden de lectura.

    No se puede ordenar solo por posición: una respuesta no abre un rango
    propio en el documento, así que su posición es 0 y acabaría por delante del
    comentario al que contesta.
    """

    def numero(comentario: Comentario) -> int:
        return int(comentario.id) if comentario.id.isdigit() else 0

    respuestas: Dict[str, List[Comentario]] = {}
    raices: List[Comentario] = []

    for comentario in comentarios:
        if comentario.responde_a is None:
            raices.append(comentario)
        else:
            respuestas.setdefault(comentario.responde_a, []).append(comentario)

    # Por dónde aparece en el documento; a igualdad, por orden de creación
    raices.sort(key=lambda c: (c.posicion, numero(c)))

    ordenados: List[Comentario] = []
    vistos: set[str] = set()

    def añadir(comentario: Comentario) -> None:
        # Un ciclo en `responde_a` colgaría la función. No debería pasar, pero
        # el dato viene de un archivo que puede estar mal formado.
        if comentario.id in vistos:
            return
        vistos.add(comentario.id)
        ordenados.append(comentario)
        for hija in sorted(respuestas.get(comentario.id, []), key=numero):
            añadir(hija)

    for raiz in raices:
        añadir(raiz)

    # Respuestas cuyo padre no existe: se han quedado huérfanas al borrar el
    # comentario original. Se devuelven igual, al final, en vez de perderlas.
    for comentario in sorted(comentarios, key=numero):
        if comentario.id not in vistos:
            añadir(comentario)

    return ordenados
