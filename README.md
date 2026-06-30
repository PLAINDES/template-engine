# DocX Service - Plataformas Financieras

## Overview

This microservice is the document processing engine for the Plataformas Financieras ecosystem. It handles the full lifecycle of DOCX templates: parsing placeholders, filling variables/tables/images, converting to HTML, and extracting structured sections. Files are stored and retrieved from a MinIO S3-compatible object storage instance.

## Architecture & Tech Stack

### Main Technologies

- **Runtime:** Python 3.11
- **Framework:** FastAPI 0.115 + Uvicorn 0.30
- **Document Processing:** python-docx 1.1.2 + mammoth 1.8.0
- **Storage:** MinIO 7.2.9 (S3-compatible)
- **Data Validation:** Pydantic 2.9 + pydantic-settings 2.5
- **Image Processing:** Pillow 11.1.0

### Directory Structure

- `app/routers/`: HTTP layer. Each file maps to a feature group of endpoints.
- `app/services/`: Business logic. Orchestration services and specialized processors.
  - `sections/`: Sub-package for heading-based section extraction.
- `app/models/`: Pydantic schemas for all request/response contracts.
- `app/core/`: Cross-cutting concerns — settings loader and file validators.
- `app/utils/`: Infrastructure helpers — MinIO client and in-memory caches.

---

## Project Features

### 1. Document Parsing (`POST /parse-docx`)

Uploads a DOCX file to MinIO and extracts its placeholder inventory.

- **Variables:** Detects `[VARIABLE_NAME]` patterns across body, tables, headers, and footers. Supports uppercase, lowercase, accents, and underscores.
- **Tables:** Detects keyword-based placeholders (`[TABLA_...]`, `[CREAR_TABLA_...]`, etc.).
- **Images:** Detects image placeholders (`[IMAGEN_...]`, `[FOTO_...]`, `[FOTOGRAFÍA_...]`, etc.).
- Returns a structured `ParseResult` with all found placeholders and their order of appearance.

### 2. Document Filling (`POST /fill-docx`, `POST /fill-docx-download`)

Fills a previously parsed template with real data and produces a new DOCX.

- **Variable substitution:** Replaces `[KEY]` with its assigned value, preserving font and style.
- **Table injection:** Inserts a full table (headers, rows, title, footer) at the placeholder paragraph.
- **Image injection:** Downloads an image from MinIO by key and inserts it at the placeholder with configurable width and captions.
- **Block repetition:** `[BLOQUE_TYPE_START]...[BLOQUE_TYPE_END]` markers repeat a section of the document for N items, including automatic H1 heading replication.
- `fill-docx` saves the result to MinIO and returns a presigned URL. `fill-docx-download` streams the file directly to the client.

### 3. HTML Conversion (`GET /docx-to-html/{minio_key}`)

Converts a DOCX document into HTML for web rendering.

- Uses **mammoth** with custom style mappings (H1–H4, Normal, Bold, Emphasis).
- Strips the Table of Contents automatically.
- Detects all variable placeholder patterns that Word may fragment across XML runs (e.g., `<strong>[</strong>VAR<strong>]</strong>`) and normalizes them into `<span data-var="KEY" class="doc-var">[KEY]</span>` for frontend interactivity.

### 4. Section Extraction (`/sections/...`)

Navigates a document's heading hierarchy and extracts content by section.

- `GET /sections/headings/{minio_key}` — Returns the full heading tree (H1, H2, H3...).
- `GET /sections/html/{minio_key}?h1_index=N` — Returns the HTML of a specific H1 section.
- `GET /sections/structure/{minio_key}?h1_index=N` — Returns the structural outline of a section.
- `GET /sections/full/{minio_key}?h1_index=N` — Returns HTML + structure in a single call.
- `POST /sections/warmup/{minio_key}` — Pre-caches the full document conversion as a background task.
- `POST /sections/extract` — Extracts multiple sections and assembles them into a new DOCX file.

### 5. Storage Management

All documents are stored in MinIO under a configurable bucket.

- `GET /list-docx` — Lists all uploaded documents with metadata.
- `DELETE /delete-docx` — Deletes a document and invalidates its cache entries.
- `POST /upload-image` — Uploads a standalone image to MinIO for later use in document filling.
- Presigned URLs are generated with a 7-day expiry for secure client downloads.

---

## Development Guide

### Environment Setup

Make sure you have Python 3.11+ and a running MinIO instance.

1. Clone the repository.
2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate      # Windows
   source .venv/bin/activate   # macOS/Linux
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Copy `.env.example` to `.env` and fill in your MinIO credentials.

### Commands

| Command | Description |
| :--- | :--- |
| `uvicorn app.main:app --reload --port 8001` | Starts the development server with hot reload. |
| `uvicorn app.main:app --host 0.0.0.0 --port 8001` | Starts the server bound to all interfaces. |

Interactive API docs are available at `http://localhost:8001/docs` once the server is running.

### Environment Variables

| Variable | Default | Description |
| :--- | :--- | :--- |
| `MINIO_ENDPOINT` | `127.0.0.1` | MinIO host address. |
| `MINIO_PORT` | `9100` | MinIO port. |
| `MINIO_USE_SSL` | `false` | Enable SSL for MinIO connection. |
| `MINIO_ACCESS_KEY` | — | MinIO access key. |
| `MINIO_SECRET_KEY` | — | MinIO secret key. |
| `MINIO_BUCKET` | `prosedi` | Target bucket name. |
| `ALLOWED_ORIGINS` | `["*"]` | CORS allowed origins. Restrict in production. |
| `MAX_FILE_SIZE_MB` | `20` | Maximum upload size in megabytes. |

### Code Formatting & Standards

- **Commits:** We follow the [Conventional Commits](https://www.conventionalcommits.org/) standard:
  - `feat:` New feature or endpoint.
  - `fix:` Bug fix.
  - `refactor:` Code change that neither fixes a bug nor adds a feature.
  - `chore:` Maintenance, dependency updates.

### Development Conventions

- **Routing:** Each router file owns one feature domain. Do not mix document filling logic into the sections router, for example.
- **Service layer:** Routers must not contain business logic. All processing goes in `app/services/`.
- **Caching:** The three cache modules (`docx_cache`, `full_html_cache`, `section_cache`) are thread-safe. Always call the corresponding `invalidate_cache()` when a document is modified or deleted.
- **Placeholders:** Variable names inside brackets must match exactly between the DOCX template and the `FillRequest` payload keys.
- **File validation:** All uploads must go through `app/core/validators.py`. It checks extension, MIME type, file size, and ZIP magic bytes (DOCX signature).

---

## Docker Setup & Deployment Guide

The service is fully containerized. The image is based on `python:3.11-slim` and exposes port **8001**.

### 1. Local Development

Mounts the local directory so code changes reflect without rebuilding.

```bash
docker compose up -d --build
```

### 2. Rebuild After Code Changes

```bash
docker compose up -d --build
```

### 3. Start Without Rebuilding

Use only when the image is already up to date.

```bash
docker compose up -d
```

### 4. Stop Services

```bash
docker compose down
```

> **Note:** The service connects to an external Docker network (`proproyectapi_prosedi_net`). Make sure that network exists before starting the container.
