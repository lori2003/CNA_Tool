"""
api/tools_routes.py — Endpoints REST per i tool
=================================================
GET  /api/tools                      → lista tool (JSON)
GET  /api/tools/{uid}/dyninfo?key=   → esegue funzione dynamic_info
POST /api/tools/{uid}/run            → esegue tool, restituisce zip
POST /api/tools/reload               → ricarica tool dal disco
GET  /api/tools/{uid}                → singolo tool (JSON)
GET  /api/config                     → legge config tema + AI
POST /api/config                     → salva config
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
import urllib.parse

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import Response
import io

# Aggiungi la root del progetto a sys.path PRIMA di qualsiasi import locale.
# Questo garantisce che "import streamlit" trovi streamlit.py (lo shim)
# nella root invece del pacchetto vero (non installato).
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
_CORE = _ROOT / "core"
if str(_CORE) not in sys.path:
    sys.path.insert(0, str(_CORE))

from core.discovery import discover_tools, tool_to_json
from core.runner import run_tool
from core.config import load_theme_config, save_theme_config, load_ai_config, save_ai_config

router = APIRouter()

# ── Cache tool in memoria ──────────────────────────────────────
_TOOLS_CACHE: List[Dict[str, Any]] = []
_BASE_DIR: Path = _ROOT


def init_tools(base_dir: Path) -> None:
    global _TOOLS_CACHE, _BASE_DIR
    _BASE_DIR = base_dir
    # streamlit.py è nella root; sys.path.insert(0, root) già fatto sopra.
    # discover_tools importa i tool → "import streamlit as st" trova lo shim.
    _TOOLS_CACHE = discover_tools(base_dir / "tools")


def get_tool_by_uid(uid: str) -> Optional[Dict[str, Any]]:
    uid_decoded = urllib.parse.unquote(uid)
    for t in _TOOLS_CACHE:
        if t.get("uid") == uid_decoded:
            return t
    return None


# ── Routes ────────────────────────────────────────────────────

@router.get("/tools")
def list_tools():
    return [tool_to_json(t) for t in _TOOLS_CACHE]


@router.post("/tools/reload")
def reload_tools():
    global _TOOLS_CACHE
    _TOOLS_CACHE = discover_tools(_BASE_DIR / "tools")
    return {"ok": True, "count": len(_TOOLS_CACHE)}


@router.get("/tools/{uid:path}/dyninfo")
def get_dynamic_info(uid: str, key: str):
    """
    Chiama la funzione dynamic_info di un param e restituisce
    i messaggi catturati dallo shim streamlit (info/warning/error/success).
    """
    t = get_tool_by_uid(uid)
    if not t:
        raise HTTPException(404, f"Tool '{uid}' non trovato.")

    param = next(
        (p for p in t.get("params", []) if p.get("key") == key),
        None,
    )
    if not param or param.get("type") != "dynamic_info":
        raise HTTPException(404, f"Param '{key}' non trovato o non è dynamic_info.")

    mod = t.get("module_obj")
    if not mod:
        return {"messages": [], "text": ""}

    func_name = param.get("function", key)
    func = getattr(mod, func_name, None) or getattr(mod, key, None)
    if not callable(func):
        return {"messages": [], "text": ""}

    # Usa il capture thread-local dello shim streamlit
    import streamlit as _st
    _st._clear_messages()

    try:
        result = func({})
    except Exception as e:
        return {"messages": [{"type": "error", "text": str(e)}], "text": ""}

    captured = list(_st._messages())
    text = str(result).strip() if result and isinstance(result, str) else ""

    # Se la funzione restituisce un testo significativo, aggiungilo ai messaggi
    if text and text not in ("None", ""):
        msg_type = "error" if text.startswith("❌") else "info"
        captured.insert(0, {"type": msg_type, "text": text})

    return {"messages": captured, "text": text}


@router.post("/tools/{uid:path}/run")
async def run_tool_endpoint(uid: str, request: Request):
    t = get_tool_by_uid(uid)
    if not t:
        raise HTTPException(404, f"Tool '{uid}' non trovato.")
    if t.get("import_error") or t.get("runner") is None:
        raise HTTPException(400, "Tool non eseguibile (errore import o manca run()).")

    # Leggi multipart form
    form = await request.form()

    inputs: Dict[str, Any] = {}
    params: Dict[str, Any] = {}

    tool_inputs_keys = {i["key"] for i in t.get("inputs", [])}
    tool_params_keys = {p["key"] for p in t.get("params", [])}
    tool_params_map  = {p["key"]: p for p in t.get("params", [])}

    for field_name, value in form.multi_items():
        if field_name in tool_inputs_keys:
            if hasattr(value, "read"):
                content = await value.read()
                file_data = {"filename": value.filename or field_name, "content": content}
                if field_name in inputs:
                    if isinstance(inputs[field_name], list):
                        inputs[field_name].append(file_data)
                    else:
                        inputs[field_name] = [inputs[field_name], file_data]
                else:
                    inputs[field_name] = file_data
        elif field_name in tool_params_keys:
            p_meta = tool_params_map.get(field_name, {})
            p_type = p_meta.get("type", "text")
            if p_type == "checkbox":
                params[field_name] = value in ("true", "1", "on", True)
            elif p_type == "number":
                try:
                    params[field_name] = float(value) if "." in str(value) else int(value)
                except (ValueError, TypeError):
                    params[field_name] = value
            elif p_type == "multiselect":
                existing = params.get(field_name, [])
                if not isinstance(existing, list):
                    existing = [existing]
                existing.append(value)
                params[field_name] = existing
            else:
                params[field_name] = value

    # Normalizza multi-file
    for inp in t.get("inputs", []):
        k = inp["key"]
        if inp.get("type") in ("txt_multi", "file_multi") and k in inputs:
            if not isinstance(inputs[k], list):
                inputs[k] = [inputs[k]]

    success, message, zip_bytes = run_tool(t, inputs, params)

    if not success:
        raise HTTPException(400, message)

    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=output_{t['id']}.zip"},
    )


@router.get("/tools/{uid:path}")
def get_tool(uid: str):
    t = get_tool_by_uid(uid)
    if not t:
        raise HTTPException(404, f"Tool '{uid}' non trovato.")
    return tool_to_json(t)


@router.post("/chat")
async def chat_endpoint(request: Request):
    """Proxy verso il provider AI configurato (OpenAI-compatible)."""
    try:
        import httpx
    except ImportError:
        raise HTTPException(500, "httpx non installato. Esegui: pip install httpx")

    body = await request.json()
    messages = body.get("messages", [])
    if not messages:
        raise HTTPException(400, "Campo 'messages' mancante o vuoto.")

    ai = load_ai_config(_BASE_DIR)
    model_id = ai.get("model_id", "")
    base_url = ai.get("base_url", "").rstrip("/")
    api_key  = ai.get("api_key", "")

    if not model_id or not base_url:
        raise HTTPException(400, "Configurazione AI incompleta (model_id o base_url mancanti).")

    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    payload = {"model": model_id, "messages": messages}

    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            resp = await client.post(f"{base_url}/chat/completions", json=payload, headers=headers)
        if resp.status_code != 200:
            raise HTTPException(resp.status_code, f"Errore AI: {resp.text[:300]}")
        data = resp.json()
        reply = data["choices"][0]["message"]["content"]
        return {"reply": reply}
    except httpx.RequestError as exc:
        raise HTTPException(502, f"Errore connessione AI: {exc}")


@router.get("/config")
def get_config():
    theme = load_theme_config(_BASE_DIR)
    ai = load_ai_config(_BASE_DIR)
    return {"theme": theme, "ai": ai}


@router.post("/config")
async def post_config(request: Request):
    body = await request.json()
    if "theme" in body:
        save_theme_config(body["theme"], _BASE_DIR)
    if "ai" in body:
        save_ai_config(body["ai"], _BASE_DIR)
    return {"ok": True}
