"""
data_sources/metabase_client.py
Sin conexion a Metabase - siempre retorna vacio.
Los datos llegan via Excel (POST /api/upload).
"""
import os, time, logging
from config import DATA_CACHE_TTL

logger = logging.getLogger(__name__)

_dc = {"data": None, "fetched_at": 0}


def fetch_orders(force=False, fecha=None, zona=None, cat=None):
    now = time.time()
    if not force and _dc["data"] and (now - _dc["fetched_at"]) < DATA_CACHE_TTL:
        return _dc["data"]
    return _dc["data"] or []


def invalidate_cache():
    _dc.update({"data": None, "fetched_at": 0})


def cache_info():
    now = time.time()
    age = now - _dc["fetched_at"] if _dc["fetched_at"] else None
    return {
        "has_data": bool(_dc["data"]),
        "rows": len(_dc["data"]) if _dc["data"] else 0,
        "age_seconds": round(age, 1) if age else None,
        "ttl_seconds": DATA_CACHE_TTL,
        "fresh": age is not None and age < DATA_CACHE_TTL,
    }
