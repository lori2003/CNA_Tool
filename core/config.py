"""
core/config.py — Gestione configurazione Toolbox CNA
======================================================
Centralizza load/save per:
  - theme_config.json  (sidebar_lightness, region_order)
  - ai_config.json     (model_id, base_url)

Miglioramenti rispetto alle funzioni inline in app.py:
  - TypedDict per contratto esplicito dei campi
  - Validazione con clamp/default (no silent failures)
  - Logging strutturato invece di bare except: pass
  - Funzioni testabili in isolamento da Streamlit
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import List

log = logging.getLogger(__name__)

# ── TypedDict inline (compatibile Python 3.8+) ─────────────────
# Usiamo dict plain per compatibilità, documentato con commenti.

# ThemeConfig keys: sidebar_lightness (int 10-60), region_order (List[str])
# AIConfig keys: model_id (str), base_url (str)

_THEME_DEFAULTS = {
    "sidebar_lightness": 33,
    "region_order": [],
}

_AI_DEFAULTS = {
    "model_id": "arcee-ai/trinity-large-preview:free",
    "base_url": "https://openrouter.ai/api/v1",
    "api_key": "",
}


# ── Validatori ─────────────────────────────────────────────────

def _validate_theme(raw: dict) -> dict:
    """
    Applica clamp e default ai valori tema.
    Non modifica raw in-place, restituisce un nuovo dict.
    """
    l_raw = raw.get("sidebar_lightness", _THEME_DEFAULTS["sidebar_lightness"])
    if not isinstance(l_raw, (int, float)):
        log.warning(
            "theme_config: sidebar_lightness='%s' non è numerico, uso default %d",
            l_raw, _THEME_DEFAULTS["sidebar_lightness"],
        )
        l_raw = _THEME_DEFAULTS["sidebar_lightness"]
    lightness = max(10, min(60, int(l_raw)))  # clamp 10–60

    order = raw.get("region_order", _THEME_DEFAULTS["region_order"])
    if not isinstance(order, list):
        log.warning("theme_config: region_order non è una lista, uso []")
        order = []
    # Filtro: mantieni solo stringhe
    order = [str(x) for x in order if isinstance(x, str)]

    return {"sidebar_lightness": lightness, "region_order": order}


def _validate_ai(raw: dict) -> dict:
    """
    Applica default e validazione base ai valori AI.
    Non modifica raw in-place, restituisce un nuovo dict.
    """
    model = raw.get("model_id", _AI_DEFAULTS["model_id"])
    if not isinstance(model, str) or not model.strip():
        log.warning("ai_config: model_id non valido ('%s'), uso default", model)
        model = _AI_DEFAULTS["model_id"]

    url = raw.get("base_url", _AI_DEFAULTS["base_url"])
    if not isinstance(url, str) or not url.startswith("http"):
        log.warning("ai_config: base_url non valido ('%s'), uso default", url)
        url = _AI_DEFAULTS["base_url"]

    api_key = raw.get("api_key", "")
    if not isinstance(api_key, str):
        api_key = ""
    # Fallback: se non presente nel JSON, usa la variabile d'ambiente
    if not api_key.strip():
        api_key = os.environ.get("OPENROUTER_API_KEY", "")

    return {"model_id": model.strip(), "base_url": url.strip(), "api_key": api_key.strip()}


# ── Public API ─────────────────────────────────────────────────

def load_theme_config(base_dir: Path) -> dict:
    """
    Carica il tema da <base_dir>/theme_config.json.

    Non lancia mai eccezioni: in caso di errore restituisce i valori default.

    Returns:
        dict con chiavi: sidebar_lightness (int), region_order (list)
    """
    config_file = base_dir / "theme_config.json"
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, dict):
                log.warning("theme_config.json non contiene un oggetto JSON valido")
                return dict(_THEME_DEFAULTS)
            return _validate_theme(raw)
        except Exception as exc:
            log.warning("Impossibile leggere theme_config.json: %s", exc)
    return dict(_THEME_DEFAULTS)


def save_theme_config(config: dict, base_dir: Path) -> bool:
    """
    Salva il tema su <base_dir>/theme_config.json.

    Returns:
        True se salvato con successo, False altrimenti.
    """
    config_file = base_dir / "theme_config.json"
    try:
        validated = _validate_theme(config)
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(validated, f, indent=2, ensure_ascii=False)
        return True
    except Exception as exc:
        log.error("Impossibile salvare theme_config.json: %s", exc)
        return False


def load_ai_config(base_dir: Path) -> dict:
    """
    Carica la config AI da <base_dir>/ai_config.json.

    Non lancia mai eccezioni: in caso di errore restituisce i valori default.

    Returns:
        dict con chiavi: model_id (str), base_url (str)
    """
    config_file = base_dir / "ai_config.json"
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, dict):
                log.warning("ai_config.json non contiene un oggetto JSON valido")
                return dict(_AI_DEFAULTS)
            return _validate_ai(raw)
        except Exception as exc:
            log.warning("Impossibile leggere ai_config.json: %s", exc)
    return dict(_AI_DEFAULTS)


def save_ai_config(config: dict, base_dir: Path) -> bool:
    """
    Salva la config AI su <base_dir>/ai_config.json.

    Returns:
        True se salvato con successo, False altrimenti.
    """
    config_file = base_dir / "ai_config.json"
    try:
        validated = _validate_ai(config)
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(validated, f, indent=2, ensure_ascii=False)
        return True
    except Exception as exc:
        log.error("Impossibile salvare ai_config.json: %s", exc)
        return False
