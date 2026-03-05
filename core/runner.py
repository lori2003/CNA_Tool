"""
core/runner.py — Tool execution engine
=======================================
Esegue run() di un tool in una directory temporanea, raccoglie gli eventi
catturati dallo shim streamlit (error/success/info/warning/log/dataframe/progress)
e restituisce un payload strutturato.

Firma pubblica:
    run_tool(tool, inputs, params) -> (success, message, zip_bytes | None, events)
"""
from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def run_tool(
    tool: Dict[str, Any],
    inputs: Dict[str, Any],
    params: Dict[str, Any],
) -> Tuple[bool, str, Optional[bytes], List[Dict[str, Any]]]:
    """
    Esegue il tool e restituisce (success, message, zip_bytes, events).

    - success: True se tutto ok
    - message: messaggio principale
    - zip_bytes: bytes dello zip con i file output (None se errore)
    - events: lista eventi UI catturati (error/success/info/warning/log/dataframe/progress)
    """
    runner = tool.get("runner")
    if runner is None:
        return False, "Tool non ha una funzione run().", None, []

    # Reset stato streamlit (session_state + messaggi) per ogni run
    try:
        import streamlit as _st
        _st._reset_for_run()
    except Exception:
        pass

    with tempfile.TemporaryDirectory() as tmp:
        out_dir = Path(tmp) / "output"
        out_dir.mkdir()

        # Salva i file di input su disco per passarli a run()
        saved_inputs: Dict[str, Any] = {}
        input_dir = Path(tmp) / "input"
        input_dir.mkdir()

        for key, data in inputs.items():
            if data is None:
                continue
            if isinstance(data, list):
                # file_multi / txt_multi
                paths = []
                for i, file_bytes in enumerate(data):
                    fname = file_bytes.get("filename", f"file_{i}")
                    p = input_dir / fname
                    p.write_bytes(file_bytes["content"])
                    paths.append(p)
                saved_inputs[key] = paths
            else:
                fname = data.get("filename", "file")
                p = input_dir / fname
                p.write_bytes(data["content"])
                saved_inputs[key] = p

        # Merge inputs + params per chiamata run()
        run_kwargs = {**saved_inputs, **params, "out_dir": out_dir}

        try:
            result = runner(**run_kwargs)
        except TypeError:
            try:
                result = runner(**{k: v for k, v in run_kwargs.items()})
            except Exception as e2:
                events = _collect_events()
                events.append({"type": "error", "text": str(e2)})
                return False, f"Errore esecuzione tool: {e2}", None, events
        except Exception as e:
            events = _collect_events()
            events.append({"type": "error", "text": str(e)})
            return False, f"Errore esecuzione tool: {e}", None, events

        # Raccoglie eventi catturati dallo shim
        events = _collect_events()

        # Raccoglie i file output — accetta sia Path che str
        output_files: List[Path] = []
        if isinstance(result, list):
            for item in result:
                if isinstance(item, str):
                    item = Path(item)
                if isinstance(item, Path) and item.exists():
                    output_files.append(item)

        if not output_files:
            output_files = [p for p in out_dir.rglob("*") if p.is_file()]

        if not output_files:
            return False, "Il tool non ha prodotto file di output.", None, events

        # Crea zip in memoria
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in output_files:
                try:
                    arcname = p.relative_to(out_dir)
                except ValueError:
                    arcname = p.name
                zf.write(p, arcname)

        zip_buf.seek(0)
        message = f"Completato. {len(output_files)} file prodotti."
        return True, message, zip_buf.read(), events


def _collect_events() -> List[Dict[str, Any]]:
    """Raccoglie e restituisce gli eventi catturati dallo shim streamlit."""
    try:
        import streamlit as _st
        return list(_st._messages())
    except Exception:
        return []
