"""
core/discovery.py — Tool discovery engine
==========================================
Portato da app.py: scansiona tools/ e restituisce la lista dei tool registrati.
I tool devono definire TOOL = {...} e def run(..., out_dir: Path) -> List[Path].
"""
from __future__ import annotations

import ast
import importlib
import importlib.util
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SUPPORTED_INPUT_TYPES = {
    "txt_multi", "txt_single", "xlsx_single",
    "file_multi", "file_single",
    "warning", "info", "error", "success", "markdown",
}
SUPPORTED_PARAM_TYPES = {
    "select", "radio", "checkbox", "number", "text",
    "textarea", "multiselect", "dynamic_info", "folder", "file_path_info",
    "markdown", "info", "warning", "error", "success",
    "action_button",
}


def _slug(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"[^0-9a-z]+", "_", s)
    return s.strip("_") or "x"


def _safe_mod_name(parts: Tuple[str, ...]) -> str:
    raw = "__".join(parts).replace(".py", "")
    raw = re.sub(r"[^0-9A-Za-z_]+", "_", raw)
    return f"tool__{raw}"


def _load_module(mod_name: str, path: Path):
    importlib.invalidate_caches()
    if mod_name in sys.modules:
        del sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(mod_name, str(path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Impossibile creare spec per {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


def discover_tools(tools_dir: Path) -> List[Dict[str, Any]]:
    tools: List[Dict[str, Any]] = []
    if not tools_dir.exists():
        return tools

    for py in sorted(tools_dir.rglob("*.py")):
        if py.name == "__init__.py" or py.name.startswith("_"):
            continue
        if "extension" in py.parts:
            continue

        rel = py.relative_to(tools_dir)
        parts = rel.parts
        region_folder = parts[0] if len(parts) >= 2 else "Generali"
        mod_name = _safe_mod_name(parts)

        try:
            mod = _load_module(mod_name, py)
        except Exception as e:
            tools.append({
                "uid": f"__error__{region_folder}__{py.stem}",
                "id": py.stem,
                "region": region_folder,
                "name": f"❌ Errore import: {py.stem}",
                "description": f"Errore:\n{e}",
                "inputs": [],
                "params": [],
                "runner": None,
                "import_error": True,
                "source_path": str(py),
            })
            continue

        if not hasattr(mod, "TOOL"):
            continue

        tool = dict(getattr(mod, "TOOL"))
        runner = getattr(mod, "run", None)
        dynamic_params = getattr(mod, "get_dynamic_params", None)

        region = tool.get("region") or region_folder
        base_id = tool.get("id") or py.stem
        base_name = tool.get("name") or py.stem

        uid = base_id
        if "/" not in str(uid) and str(region) and str(region) != "Generali":
            uid = f"{region}/{base_id}"

        tool.setdefault("id", base_id)
        tool.setdefault("name", base_name)
        tool.setdefault("description", "")
        tool.setdefault("inputs", [])
        tool.setdefault("params", [])

        tool["uid"] = uid
        tool["region"] = region
        tool["runner"] = runner
        tool["dynamic_params"] = dynamic_params
        tool["module_obj"] = mod
        tool["import_error"] = False
        tool["source_path"] = str(py)

        if not callable(runner):
            tool["import_error"] = True
            tool["runner"] = None
            tool["description"] += "\n\n⚠️ Manca la funzione run(...)."

        for inp in tool.get("inputs", []):
            if inp.get("type") not in SUPPORTED_INPUT_TYPES:
                tool["import_error"] = True
                tool["runner"] = None

        for p in tool.get("params", []):
            if p.get("type") not in SUPPORTED_PARAM_TYPES:
                tool["import_error"] = True
                tool["runner"] = None
            if p.get("type") in ("select", "radio", "multiselect") and not p.get("options"):
                tool["import_error"] = True
                tool["runner"] = None

        tools.append(tool)

    # Deduplicazione per uid — in caso di conflitto vince il primo trovato
    seen: dict = {}
    deduped = []
    for t in tools:
        uid = t.get("uid", "")
        if uid in seen:
            import logging
            logging.getLogger(__name__).warning(
                "Tool uid duplicato '%s': ignorato '%s' (già caricato da '%s')",
                uid, t.get("source_path"), seen[uid],
            )
            continue
        seen[uid] = t.get("source_path", "")
        deduped.append(t)
    tools = deduped

    tools.sort(key=lambda t: (
        bool(t.get("import_error")),
        0 if (t.get("region") or "").lower() == "generali" else 1,
        (t.get("region") or "").lower(),
        (t.get("name") or "").lower(),
    ))
    return tools


def tool_to_json(t: Dict[str, Any]) -> Dict[str, Any]:
    """Serializza un tool (rimuove runner/module, non serializzabili)."""
    return {
        "uid": t.get("uid"),
        "id": t.get("id"),
        "name": t.get("name"),
        "region": t.get("region"),
        "description": t.get("description", ""),
        "inputs": t.get("inputs", []),
        "params": t.get("params", []),
        "import_error": t.get("import_error", False),
        "source_path": t.get("source_path", ""),
        "email_reminder": t.get("email_reminder"),
    }
