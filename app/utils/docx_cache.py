# app/utils/docx_cache.py
import threading
from app.utils.minio_client import download_from_minio

_cache: dict[str, bytes] = {}
_lock  = threading.Lock()


def get_docx_cached(minio_key: str) -> bytes:
    """
    Retorna el buffer del .docx desde caché.
    Si no está en caché, lo descarga de MinIO y lo guarda.
    Thread-safe.
    """
    with _lock:
        if minio_key not in _cache:
            _cache[minio_key] = download_from_minio(minio_key)
        return _cache[minio_key]


def invalidate_cache(minio_key: str) -> None:
    """
    Elimina el .docx del caché — llamar cuando se actualiza el archivo.
    """
    with _lock:
        _cache.pop(minio_key, None)


def clear_cache() -> None:
    with _lock:
        _cache.clear()