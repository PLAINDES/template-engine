# app/services/docx_service.py
import time
from app.models.schemas import ParseResult, FillResult
from app.services.parser import extract_variables
from app.services.filler import fill_document
from app.services.html_converter import docx_to_html
from app.utils.minio_client import (
    upload_to_minio,
    get_presigned_url,
    download_from_minio,
    delete_from_minio,
    list_objects_from_minio,
)


def list_documents() -> dict:
    objects = list_objects_from_minio("docx-uploads/")
    result  = []
    for obj in objects:
        try:
            url = get_presigned_url(obj["key"])
        except Exception:
            url = ""
        result.append({
            "minio_key":   obj["key"],
            "filename":    obj["key"].split("/")[-1].split("_", 1)[-1],
            "uploaded_at": obj["last_modified"],
            "size_kb":     round(obj["size"] / 1024, 2),
            "minio_url":   url,
        })
    result.sort(key=lambda x: x["uploaded_at"], reverse=True)
    return {"total": len(result), "documents": result}


def upload_and_parse(buffer: bytes, original_filename: str) -> ParseResult:
    timestamp = int(time.time())
    safe_name = original_filename.replace(" ", "_")
    minio_key = f"docx-uploads/{timestamp}_{safe_name}"

    upload_to_minio(minio_key, buffer)

    try:
        minio_url = get_presigned_url(minio_key)
    except Exception:
        minio_url = ""

    result = extract_variables(buffer)

    return ParseResult(
        minio_key       = minio_key,
        minio_url       = minio_url,
        filename        = original_filename,
        total_variables = result["total_variables"],
        total_tablas    = result["total_tablas"],
        total_imagenes  = result["total_imagenes"],
        variables       = result["variables"],
        tablas          = result["tablas"],
        imagenes        = result["imagenes"],
    )


def reload_document(minio_key: str) -> ParseResult:
    buffer = download_from_minio(minio_key)
    result = extract_variables(buffer)

    try:
        minio_url = get_presigned_url(minio_key)
    except Exception:
        minio_url = ""

    raw_name = minio_key.split("/")[-1]
    filename = raw_name.split("_", 1)[-1] if "_" in raw_name else raw_name

    return ParseResult(
        minio_key       = minio_key,
        minio_url       = minio_url,
        filename        = filename,
        total_variables = result["total_variables"],
        total_tablas    = result["total_tablas"],
        total_imagenes  = result["total_imagenes"],
        variables       = result["variables"],
        tablas          = result["tablas"],
        imagenes        = result["imagenes"],
    )


def get_docx_html(minio_key: str) -> dict:
    buffer = download_from_minio(minio_key)
    return docx_to_html(buffer)


def fill_and_save(buffer: bytes, minio_key: str, variables: dict, tablas: list, imagenes: list, bloques: list = []) -> FillResult:
    result_buffer = fill_document(
        docx_buffer = buffer,
        variables   = variables,
        tablas      = tablas,
        imagenes    = imagenes,
        bloques     = bloques,
    )

    timestamp  = int(time.time())
    original   = minio_key.split("/")[-1]
    output_key = f"docx-generated/{timestamp}_{original}"

    upload_to_minio(output_key, result_buffer)
    output_url = get_presigned_url(output_key)

    return FillResult(minio_key=output_key, minio_url=output_url)


def fill_and_stream(buffer: bytes, minio_key: str, variables: dict, tablas: list, imagenes: list, bloques: list = []) -> bytes:
    return fill_document(
        docx_buffer = buffer,
        variables   = variables,
        tablas      = tablas,
        imagenes    = imagenes,
        bloques     = bloques,
    )


def delete_document(minio_key: str) -> None:
    allowed_prefixes = ("docx-uploads/", "docx-generated/")
    if not any(minio_key.startswith(p) for p in allowed_prefixes):
        raise ValueError(
            f"Solo se pueden eliminar archivos de docx-uploads/ o docx-generated/. "
            f"Key recibida: {minio_key}"
        )
    delete_from_minio(minio_key)