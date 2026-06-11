# app/utils/full_html_cache.py
"""
Caché del HTML completo del .docx.
Guarda:
  - El HTML completo convertido por mammoth
  - El mapeo de h1_index (.docx) → h1_order (HTML)
  - Las secciones ya divididas por H1
"""
import threading
from typing import Optional

_cache: dict[str, dict] = {}
_lock  = threading.Lock()


def get_full_html_cache(minio_key: str) -> Optional[dict]:
    with _lock:
        return _cache.get(minio_key)


def set_full_html_cache(minio_key: str, data: dict) -> None:
    """
    data debe tener:
    {
        "full_html":    str,              # HTML completo
        "index_map":    dict[int, int],   # h1_index(.docx) → h1_order(HTML)
        "sections":     dict[int, str],   # h1_index(.docx) → html_fragment
        "structure":    dict[int, list],  # h1_index(.docx) → estructura
    }
    """
    with _lock:
        _cache[minio_key] = data


def invalidate_full_html_cache(minio_key: str) -> None:
    with _lock:
        _cache.pop(minio_key, None)


def clear_full_html_cache() -> None:
    with _lock:
        _cache.clear()