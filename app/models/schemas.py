# app/models/schemas.py
from pydantic import BaseModel
from typing import List, Dict

class VariableInfo(BaseModel):
    key:      str
    label:    str
    value:    str  = ""
    in_table: bool = False
    order:    int  = 0    # orden de aparición en el documento

class TablaPlaceholder(BaseModel):
    index:           int
    paragraph_index: int
    order:           int = 0

class ImagenPlaceholder(BaseModel):
    index:           int
    paragraph_index: int
    order:           int = 0
    key:             str = ""

class ParseResult(BaseModel):
    minio_key:       str
    minio_url:       str
    filename:        str   # nombre original del archivo
    total_variables: int   # cuántas variables se detectaron
    total_tablas:    int   # cuántos placeholders de tabla
    total_imagenes:  int   # cuántos placeholders de imagen
    variables:       List[VariableInfo]
    tablas:          List[TablaPlaceholder]
    imagenes:        List[ImagenPlaceholder]

class TablaData(BaseModel):
    placeholder_index: int
    paragraph_index:   int = 0
    headers:           List[str]
    rows:              List[List[str]]
    titulo:            str = ""
    pie:               str = ""

class ImagenData(BaseModel):
    placeholder_index: int
    paragraph_index:   int = 0
    minio_key:         str
    key:               str = ""
    width_inches:      float = 4.0
    titulo:            str = ""
    pie:               str = ""
    descripcion:       str = ""

class BloqueItemData(BaseModel):
    variables: Dict[str, str] = {}

class BloqueData(BaseModel):
    tipo:  str
    items: List[Dict[str, str]] = []

class TextReplacement(BaseModel):
    originalText: str
    newText:      str
    sectionIndex: int = -1

class FillRequest(BaseModel):
    minio_key:    str
    variables:    Dict[str, str]
    tablas:       List[TablaData]        = []
    imagenes:     List[ImagenData]       = []
    bloques:      List[BloqueData]       = []
    replacements: List[TextReplacement]  = []

class FillResult(BaseModel):
    minio_key: str
    minio_url: str

class DeleteResult(BaseModel):
    deleted:   bool
    minio_key: str
    message:   str
    
class HeadingItem(BaseModel):
    index:    int                    # paragraph_index en el documento
    level:    int                    # 1=H1, 2=H2, 3=H3...
    text:     str                    # texto del heading
    children: List["HeadingItem"] = []  # subsecciones hijas
 
HeadingItem.model_rebuild()          # necesario para el forward reference
 
class HeadingsResult(BaseModel):
    minio_key: str
    filename:  str
    headings:  List[HeadingItem]     # árbol completo
 
class ExtractSectionsRequest(BaseModel):
    minio_key:        str
    selected_indexes: List[int]      # paragraph_index de los H1 seleccionados
    variables:        Dict[str, str] = {}
    tablas:           List[dict]     = []
    imagenes:         List[dict]     = []