# app/utils/section_cache.py
import threading
from typing import Optional

# Caché de secciones: key = "{minio_key}:{h1_index}"
_section_cache: dict[str, dict] = {}
_lock = threading.Lock()


def get_section_cached(minio_key: str, h1_index: int) -> Optional[dict]:
    key = f"{minio_key}:{h1_index}"
    with _lock:
        return _section_cache.get(key)


def set_section_cached(minio_key: str, h1_index: int, data: dict) -> None:
    key = f"{minio_key}:{h1_index}"
    with _lock:
        _section_cache[key] = data


def invalidate_section_cache(minio_key: str) -> None:
    """Elimina todas las secciones cacheadas de un documento."""
    with _lock:
        keys_to_delete = [k for k in _section_cache if k.startswith(f"{minio_key}:")]
        for k in keys_to_delete:
            del _section_cache[k]


def clear_section_cache() -> None:
    with _lock:
        _section_cache.clear()