# app/utils/minio_client.py
from minio import Minio
from minio.error import S3Error
from io import BytesIO
from datetime import timedelta
from typing import List, Dict
from app.core.config import settings
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

client = Minio(
    endpoint   = f"{settings.MINIO_ENDPOINT}:{settings.MINIO_PORT}",
    access_key = settings.MINIO_ACCESS_KEY,
    secret_key = settings.MINIO_SECRET_KEY,
    secure     = settings.MINIO_USE_SSL,
)


def download_from_minio(key: str) -> bytes:
    try:
        response = client.get_object(settings.MINIO_BUCKET, key)
        data = response.read()
        response.close()
        response.release_conn()
        return data
    except S3Error as e:
        raise Exception(f"Error descargando [{key}]: {e}")


def upload_to_minio(
    key: str,
    data: bytes,
    content_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
) -> str:
    try:
        client.put_object(
            bucket_name  = settings.MINIO_BUCKET,
            object_name  = key,
            data         = BytesIO(data),
            length       = len(data),
            content_type = content_type,
        )
        return key
    except S3Error as e:
        raise Exception(f"Error subiendo [{key}]: {e}")


def get_presigned_url(key: str, expires_seconds: int = 604800) -> str:
    try:
        url = client.presigned_get_object(
            settings.MINIO_BUCKET,
            key,
            expires=timedelta(seconds=expires_seconds),
        )
        return url
    except S3Error as e:
        raise Exception(f"Error generando URL presignada [{key}]: {e}")


def delete_from_minio(key: str) -> None:
    try:
        client.remove_object(settings.MINIO_BUCKET, key)
    except S3Error as e:
        raise Exception(f"Error eliminando [{key}]: {e}")


def list_objects_from_minio(prefix: str = "") -> List[Dict]:
    """
    Lista todos los objetos en MinIO con un prefijo dado.
    Retorna lista de { key, size, last_modified }
    """
    try:
        objects = client.list_objects(
            settings.MINIO_BUCKET,
            prefix    = prefix,
            recursive = True,
        )
        result = []
        for obj in objects:
            result.append({
                "key":           obj.object_name,
                "size":          obj.size or 0,
                "last_modified": obj.last_modified.isoformat() if obj.last_modified else "",
            })
        return result
    except S3Error as e:
        raise Exception(f"Error listando objetos [{prefix}]: {e}")