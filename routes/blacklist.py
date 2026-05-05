"""
routes/blacklist.py
Gestion de lista negra de tecnicos.
Los tecnicos en la lista negra se excluyen de las sugerencias de nivelacion.
"""
import logging
from flask import Blueprint, jsonify, request

logger = logging.getLogger(__name__)
blacklist_bp = Blueprint("blacklist", __name__, url_prefix="/api/blacklist")

_blacklist = set()


def get_blacklist():
    return sorted(_blacklist)


def is_blacklisted(tecnico):
    return tecnico in _blacklist


def filter_suggestions(sugerencias):
    if not _blacklist:
        return sugerencias
    return [s for s in sugerencias
            if s.get("tecnico_actual") not in _blacklist
            and s.get("tecnico_sugerido") not in _blacklist]


@blacklist_bp.get("")
def get_list():
    return jsonify({"status":"ok","blacklist":get_blacklist(),"total":len(_blacklist)})


@blacklist_bp.post("/add")
def add_tech():
    body = request.get_json(silent=True) or {}
    tecnico = str(body.get("tecnico","")).strip()
    if not tecnico:
        return jsonify({"status":"error","message":"Campo tecnico requerido"}),400
    _blacklist.add(tecnico)
    logger.info(f"Lista negra: +'{tecnico}' total={len(_blacklist)}")
    return jsonify({"status":"ok","mensaje":f"'{tecnico}' agregado a lista negra","blacklist":get_blacklist()})


@blacklist_bp.post("/remove")
def remove_tech():
    body = request.get_json(silent=True) or {}
    tecnico = str(body.get("tecnico","")).strip()
    _blacklist.discard(tecnico)
    return jsonify({"status":"ok","mensaje":f"'{tecnico}' removido","blacklist":get_blacklist()})


@blacklist_bp.post("/clear")
def clear_list():
    count = len(_blacklist)
    _blacklist.clear()
    return jsonify({"status":"ok","mensaje":f"{count} tecnicos removidos","blacklist":[]})
