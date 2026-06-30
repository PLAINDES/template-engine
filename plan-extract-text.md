# Plan: endpoint `GET /extract-text`

## Objetivo

Exponer un endpoint que extrae todo el texto plano de un `.docx` almacenado en MinIO,
para que el backend NestJS pueda enviárselo a Gemini como contexto del documento
al autocompletar variables del sistema DOCX.

---

## Contrato del endpoint

```
GET /extract-text?key={minioKey}

Response 200:
{
  "texto": "Texto completo del documento separado por saltos de línea..."
}

Response 404: documento no encontrado en MinIO
Response 500: error interno
```

---

## Archivos a modificar

### 1. `app/services/docx_service.py`

Agregar función `extract_plain_text(minio_key: str) -> dict` al final del archivo,
antes de `delete_document`.

**Lógica:**
1. Obtener el buffer del `.docx` desde `docx_cache.get_docx_cached(minio_key)`
   — usa la caché L1 existente, sin descargar dos veces si ya está en memoria.
2. Abrir el buffer con `docx.Document(BytesIO(buffer))`.
3. Recorrer `doc.paragraphs` y filtrar los que tengan texto no vacío.
4. Unir con `\n` y retornar `{"texto": "..."}`.

```python
def extract_plain_text(minio_key: str) -> dict:
    from io import BytesIO
    from docx import Document
    from app.utils.docx_cache import get_docx_cached

    buffer   = get_docx_cached(minio_key)
    doc      = Document(BytesIO(buffer))
    parrafos = [p.text for p in doc.paragraphs if p.text.strip()]
    return {"texto": "\n".join(parrafos)}
```

**Sin librerías nuevas** — `python-docx` ya está en `requirements.txt`.

---

### 2. `app/routers/docx.py`

Agregar endpoint `GET /extract-text` al final del router existente.

**Lógica:**
- Recibe `key` como query param obligatorio.
- Llama a `docx_service.extract_plain_text(key)`.
- Maneja `FileNotFoundError` → 404, cualquier otra excepción → 500.

```python
@router.get("/extract-text")
def extract_text(key: str = Query(..., description="MinIO key del .docx")):
    try:
        return docx_service.extract_plain_text(key)
    except FileNotFoundError as e:
        raise HTTPException(404, f"Documento no encontrado: {e}")
    except Exception as e:
        raise HTTPException(500, f"Error extrayendo texto: {e}")
```

---

## Lo que NO se modifica

| Archivo | Razón |
|---|---|
| `app/models/schemas.py` | La respuesta es un dict simple, no necesita schema Pydantic |
| `app/utils/docx_cache.py` | Se reutiliza tal cual — ningún cambio |
| `app/utils/minio_client.py` | La caché ya lo usa internamente |
| `app/main.py` | El router ya está registrado |
| `requirements.txt` | Sin dependencias nuevas |

---

## Por qué usar `docx_cache` y no `download_from_minio` directo

`download_from_minio` descarga siempre desde MinIO.
`get_docx_cached` devuelve el buffer desde memoria si el documento ya fue
accedido antes (por ejemplo, tras un `parse-docx` o `docx-to-html`).
Esto evita una descarga innecesaria cuando el frontend llama a varios
endpoints seguidos sobre el mismo documento.

---

## Limitación conocida

`doc.paragraphs` no incluye texto dentro de tablas del `.docx`.
Si las variables relevantes para Gemini están en tablas del documento,
se puede agregar en una segunda iteración:

```python
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if p.text.strip():
                    parrafos.append(p.text)
```

Por ahora se implementa solo párrafos — es suficiente para el contexto
narrativo que necesita Gemini.
