# Análisis Técnico — docx-service

## 1. Objetivo general

Microservicio FastAPI especializado en **procesamiento de documentos Word (.docx)**. Permite subir plantillas con variables `[PLACEHOLDER]`, extraer su estructura, rellenarlas con datos dinámicos (texto, tablas, imágenes) y convertirlas a HTML para previsualización. Actúa como backend de documentos para una aplicación mayor (ProProyectApi en NestJS).

---

## 2. Arquitectura

Patrón **Router → Service → Utils**, sin ORM ni base de datos relacional. El almacenamiento es MinIO (S3-compatible). Caché en memoria (3 niveles).

```
Cliente (NestJS)
     │
     ▼
 FastAPI (puerto 8001)
     │
 ┌───┴──────────────┐
 │    Routers       │  ← Validación HTTP, manejo de errores
 └───┬──────────────┘
     │
 ┌───┴──────────────┐
 │    Services      │  ← Lógica de negocio pura
 └───┬──────────────┘
     │
 ┌───┴──────────────┐
 │    Utils         │  ← MinIO client, caché
 └───┬──────────────┘
     │
     ▼
   MinIO (S3)        ← Único almacenamiento persistente
```

---

## 3. Estructura de carpetas

```
docx-service/
├── app/
│   ├── main.py                          # Bootstrap FastAPI, registra routers
│   ├── core/
│   │   ├── config.py                    # Settings (pydantic-settings + .env)
│   │   └── validators.py                # Validación de archivos .docx
│   ├── models/
│   │   └── schemas.py                   # Todos los modelos Pydantic
│   ├── routers/
│   │   ├── health.py                    # GET /health
│   │   ├── docx.py                      # Endpoints de documentos
│   │   └── sections.py                  # Endpoints de secciones
│   ├── services/
│   │   ├── docx_service.py              # Orquestador principal
│   │   ├── parser.py                    # Extrae variables/placeholders del .docx
│   │   ├── filler.py                    # Rellena variables/tablas/imágenes
│   │   ├── html_converter.py            # DOCX → HTML (mammoth)
│   │   ├── html_splitter.py             # Divide HTML por secciones H1
│   │   └── sections/
│   │       ├── heading_parser.py        # Detecta y construye árbol de headings
│   │       ├── section_builder.py       # Orquesta extracción + relleno
│   │       ├── section_extractor.py     # Extrae párrafos por sección
│   │       └── section_html_extractor.py # HTML + estructura + caché
│   └── utils/
│       ├── minio_client.py              # Cliente MinIO (upload/download/presign)
│       ├── docx_cache.py                # Caché nivel 1: buffers .docx
│       ├── full_html_cache.py           # Caché nivel 2: HTML completo
│       └── section_cache.py             # Caché nivel 3: secciones individuales
├── TECHNICAL_CONTEXT.md
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── .env
```

---

## 4. Flujo de una request

```
Request HTTP
    │
    ▼
main.py (CORSMiddleware)
    │
    ▼
Router (validación, HTTPException)
    │
    ├─► validators.py       (extensión, MIME, tamaño, magic bytes)
    │
    ▼
Service
    │
    ├─► docx_cache.py       (¿está en memoria? → sí: devuelve / no: descarga MinIO)
    ├─► minio_client.py     (upload / download / presigned URL)
    ├─► parser.py           (extrae [VARIABLES] con regex)
    ├─► filler.py           (reemplaza variables, inserta tablas/imágenes)
    ├─► html_converter.py   (mammoth → HTML)
    └─► section_html_extractor.py (divide HTML por H1, cachea)
    │
    ▼
Response JSON / archivo .docx / HTML
```

---

## 5. Endpoints completos

### Documents (`/`)

| Método | Ruta | Descripción | Parámetros | Body | Response |
|---|---|---|---|---|---|
| `GET` | `/health` | Estado del servicio | — | — | `{status, service}` |
| `GET` | `/list-docx` | Lista docs en MinIO | — | — | `{total, documents[]}` |
| `POST` | `/parse-docx` | Sube .docx y extrae variables | — | `multipart/file` | `ParseResult` |
| `GET` | `/parse-docx/{minio_key}` | Re-analiza doc existente | `minio_key` path | — | `ParseResult` |
| `GET` | `/docx-to-html/{minio_key}` | Convierte doc a HTML | `minio_key` path | — | `{html, messages}` |
| `POST` | `/fill-docx` | Rellena y guarda en MinIO | — | `FillRequest` | `FillResult` |
| `POST` | `/fill-docx-download` | Rellena y descarga stream | — | `FillRequest` | `.docx` binario |
| `DELETE` | `/delete-docx` | Elimina doc de MinIO | `minio_key` query | — | `DeleteResult` |

### Sections (`/sections`)

| Método | Ruta | Descripción | Parámetros |
|---|---|---|---|
| `GET` | `/sections/headings/{minio_key}` | Árbol jerárquico de headings | `minio_key` path |
| `GET` | `/sections/full/{minio_key}` | HTML + estructura de una sección | `minio_key` path, `h1_index` query |
| `GET` | `/sections/html/{minio_key}` | Solo HTML de sección | `minio_key` path, `h1_index` query |
| `GET` | `/sections/structure/{minio_key}` | Solo estructura JSON de sección | `minio_key` path, `h1_index` query |
| `POST` | `/sections/extract` | Extrae secciones a nuevo .docx | — | `ExtractSectionsRequest` |
| `POST` | `/sections/warmup/{minio_key}` | Precalienta caché en background | `minio_key` path |

---

## 6. Modelos Pydantic (`app/models/schemas.py`)

```
ParseResult
  ├── minio_key, minio_url, filename
  ├── total_variables, total_tablas, total_imagenes
  ├── variables: List[VariableInfo]
  │     └── key, label, value, in_table, order
  ├── tablas: List[TablaPlaceholder]
  │     └── index, paragraph_index, order
  └── imagenes: List[ImagenPlaceholder]
        └── index, paragraph_index, order

FillRequest
  ├── minio_key
  ├── variables: Dict[str, str]
  ├── tablas: List[TablaData]
  │     └── placeholder_index, headers[], rows[][]
  └── imagenes: List[ImagenData]
        └── placeholder_index, minio_key, width_inches

HeadingItem (árbol recursivo)
  ├── index, level, text
  └── children: List[HeadingItem]

ExtractSectionsRequest
  ├── minio_key
  ├── selected_indexes: List[int]   ← paragraph_index de los H1
  ├── variables: Dict[str, str]
  ├── tablas: List[dict]
  └── imagenes: List[dict]
```

---

## 7. Sistema de caché (3 niveles)

```
Request llega con minio_key
         │
         ▼
┌─────────────────────┐
│  docx_cache (L1)    │  Buffer .docx en memoria
│  key: minio_key     │  Hit → buffer / Miss → descarga MinIO
└─────────┬───────────┘
          │
          ▼
┌─────────────────────┐
│ full_html_cache (L2)│  HTML completo + secciones divididas
│  key: minio_key     │  Hit → dict / Miss → mammoth convierte
└─────────┬───────────┘
          │
          ▼
┌─────────────────────┐
│ section_cache (L3)  │  HTML + estructura por sección
│ key: minio_key:idx  │  Hit → dict / Miss → extrae del L2
└─────────────────────┘
```

**Invalidación:**
- `parse-docx` → invalida L3
- `extract` y `fill` → invalida L1 + L2

---

## 8. Árbol de dependencias

```
main.py
  ├── routers/health.py
  ├── routers/docx.py
  │     ├── models/schemas.py
  │     ├── core/validators.py
  │     ├── services/docx_service.py
  │     │     ├── services/parser.py
  │     │     ├── services/filler.py
  │     │     ├── services/html_converter.py
  │     │     └── utils/minio_client.py
  │     │           └── core/config.py
  │     ├── utils/minio_client.py
  │     └── utils/section_cache.py
  └── routers/sections.py
        ├── models/schemas.py
        ├── services/sections/heading_parser.py
        │     └── models/schemas.py
        ├── services/sections/section_builder.py
        │     ├── services/sections/section_extractor.py
        │     │     └── services/sections/heading_parser.py
        │     └── services/filler.py
        ├── services/sections/section_html_extractor.py
        │     ├── services/sections/heading_parser.py
        │     ├── services/html_converter.py
        │     ├── utils/docx_cache.py
        │     └── utils/full_html_cache.py
        └── utils/docx_cache.py
```

---

## 9. Configuración (`app/core/config.py`)

| Variable | Default local | Docker | Descripción |
|---|---|---|---|
| `MINIO_ENDPOINT` | `127.0.0.1` | `minio` | Host de MinIO |
| `MINIO_PORT` | `9100` | `9000` | Puerto |
| `MINIO_USE_SSL` | `false` | `false` | SSL |
| `MINIO_ACCESS_KEY` | `""` | `${MINIO_ACCESS_KEY}` | Credencial |
| `MINIO_SECRET_KEY` | `""` | `${MINIO_SECRET_KEY}` | Credencial |
| `MINIO_BUCKET` | `prosedi` | `prosedi` | Bucket |

---

## 10. Librerías clave

| Librería | Uso |
|---|---|
| `fastapi` | Framework web, routing, validación automática |
| `uvicorn` | Servidor ASGI con hot-reload |
| `python-docx` | Leer/escribir/modificar archivos .docx |
| `mammoth` | Convierte .docx → HTML preservando estilos |
| `minio` | Cliente S3-compatible para MinIO |
| `pydantic` | Validación y serialización de datos |
| `pydantic-settings` | Carga de .env con tipado |
| `lxml` | Copia de estilos XML entre documentos |
| `Pillow` | Soporte de imágenes en python-docx |
| `requests` | Descarga de imágenes externas |

---

## 11. Puntos críticos

| Punto | Riesgo | Descripción |
|---|---|---|
| Caché en memoria | Alto | Se pierde al reiniciar el contenedor — no hay persistencia |
| Sin autenticación | Alto | Todos los endpoints son públicos |
| CORS abierto | Medio | `allow_origins=["*"]` en producción |
| `--reload` en producción | Medio | El Dockerfile usa reload, no apto para prod |
| Regex H1 en HTML | Medio | `.*?` con DOTALL puede ser lento en docs muy grandes |
| SSL MinIO desactivado | Bajo | Solo aceptable en desarrollo |

---

## 12. Cómo levantar el proyecto

### Con Docker (recomendado)

```bash
# 1. Asegurarse que la red Docker existe
docker network create proproyectapi_prosedi_net

# 2. Configurar .env con credenciales de MinIO

# 3. Levantar
docker compose up -d --build

# 4. Verificar
curl http://localhost:8001/health

# 5. Ver logs
docker logs -f prosedi_docx_service
```

### Sin Docker

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8001 --reload
```

### Documentación interactiva

```
http://localhost:8001/docs      # Swagger UI
http://localhost:8001/redoc     # ReDoc
```

---

## 13. Flujo típico de uso

```
1. POST /parse-docx
   └─► Sube .docx, extrae [VARIABLES], guarda en MinIO
       Response: minio_key + lista de variables/tablas/imágenes

2. GET /sections/headings/{minio_key}
   └─► Árbol jerárquico H1 → H2 → H3
       Response: headings[] con index (paragraph_index real)

3. GET /sections/full/{minio_key}?h1_index=X
   └─► HTML de la sección + estructura de variables por subsección
       El h1_index es el paragraph_index del H1, NO un índice secuencial

4. POST /fill-docx  (o /fill-docx-download para descarga directa)
   └─► Envía variables Dict[str,str] + tablas + imágenes
       Response: minio_key del doc generado + URL presignada

5. POST /sections/extract  (opcional)
   └─► Extrae solo secciones seleccionadas a un nuevo .docx
```

---

## 14. Módulos con posible uso no confirmado

| Archivo | Estado |
|---|---|
| `app/services/html_splitter.py` | No aparece importado en routers ni en services visibles — verificar si `docx_service.py` lo usa internamente |

---

## 15. Keys de MinIO por prefijo

| Prefijo | Contenido |
|---|---|
| `docx-uploads/` | Plantillas subidas por el usuario |
| `docx-generated/` | Documentos generados tras relleno |
| `docx-images/` | Imágenes subidas para insertar en documentos |
