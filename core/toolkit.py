"""
core/toolkit.py — Tool SDK nativo per Toolbox CNA
===================================================
Sostituisce lo shim streamlit.py con un'implementazione pulita e senza dipendenze.

Uso nei tool:
    from core.toolkit import ctx
    ctx.error("messaggio")
    ctx.success("ok!")
    ctx.dataframe(df)

Uso interno (runner / routes):
    from core.toolkit import _reset_for_run, _clear_messages, _messages
"""
from __future__ import annotations

import functools
import math
import threading
from typing import Any, Dict, List, Optional

# ── Thread-local storage ───────────────────────────────────────
_tl = threading.local()


def _messages() -> List[dict]:
    if not hasattr(_tl, "msgs"):
        _tl.msgs = []
    return _tl.msgs


def _clear_messages() -> None:
    _tl.msgs = []


def _state() -> dict:
    if not hasattr(_tl, "state"):
        _tl.state = {}
    return _tl.state


def _reset_for_run() -> None:
    """Chiamato dal runner prima di ogni esecuzione."""
    _tl.msgs = []
    _tl.state = {}


# ── Helpers serializzazione ────────────────────────────────────
def _safe_cell(c: Any) -> Any:
    """Converte un valore di cella in tipo JSON-serializzabile nativo."""
    if c is None:
        return None
    if isinstance(c, bool):
        return c
    if hasattr(c, "item"):
        try:
            return c.item()
        except Exception:
            pass
    if isinstance(c, (int, float)):
        try:
            if math.isnan(c) or math.isinf(c):
                return str(c)
        except Exception:
            pass
        return c
    return str(c)


def _df_to_event(df: Any, title: str = "") -> dict:
    """Converte un DataFrame in un evento serializzabile."""
    try:
        if hasattr(df, "columns") and hasattr(df, "head"):
            cols = [str(c) for c in df.columns]
            raw_rows = df.head(200).values.tolist()
            rows = [[_safe_cell(c) for c in r] for r in raw_rows]
            return {"type": "dataframe", "text": title, "data": {"columns": cols, "rows": rows}}
    except Exception:
        pass
    return {"type": "log", "text": str(df)[:500]}


# ── Session state proxy ────────────────────────────────────────
class _SessionState:
    def __getitem__(self, k):         return _state().get(k)
    def __setitem__(self, k, v):      _state()[k] = v
    def __delitem__(self, k):         _state().pop(k, None)
    def __contains__(self, k):        return k in _state()
    def __iter__(self):               return iter(_state())
    def get(self, k, default=None):   return _state().get(k, default)
    def keys(self):                   return _state().keys()
    def values(self):                 return _state().values()
    def items(self):                  return _state().items()
    def setdefault(self, k, default=None):
        if k not in _state():
            _state()[k] = default
        return _state()[k]
    def __getattr__(self, k):
        if k.startswith("_"):
            raise AttributeError(k)
        return _state().get(k)
    def __setattr__(self, k, v):
        if k.startswith("_"):
            super().__setattr__(k, v)
        else:
            _state()[k] = v


# ── No-op layout context (columns, expander, ecc.) ────────────
class _Noop:
    """Context manager no-op per layout containers."""
    def __enter__(self):  return self
    def __exit__(self, *a): pass

    # Forwarda chiamate display al ctx principale
    def error(self, t="", *a, **kw):   ctx.error(t, *a, **kw)
    def success(self, t="", *a, **kw): ctx.success(t, *a, **kw)
    def warning(self, t="", *a, **kw): ctx.warning(t, *a, **kw)
    def info(self, t="", *a, **kw):    ctx.info(t, *a, **kw)
    def write(self, *a, **kw):         ctx.write(*a, **kw)
    def markdown(self, t="", *a, **kw): ctx.markdown(t, *a, **kw)
    def dataframe(self, *a, **kw):     ctx.dataframe(*a, **kw)
    def table(self, *a, **kw):         ctx.table(*a, **kw)
    def progress(self, *a, **kw):      return ctx.progress(*a, **kw)

    # Widget → default
    def text_input(self, label="", value="", **kw):   return value
    def text_area(self, label="", value="", **kw):    return value
    def number_input(self, label="", value=0, **kw):  return value
    def checkbox(self, label="", value=False, **kw):  return value
    def selectbox(self, label="", options=None, index=0, **kw):
        return options[index] if options else None
    def radio(self, label="", options=None, index=0, **kw):
        return options[index] if options else None
    def multiselect(self, label="", options=None, default=None, **kw):
        return default or []
    def slider(self, label="", min_value=0, max_value=100, value=0, **kw): return value
    def button(self, *a, **kw):        return False
    def file_uploader(self, *a, **kw): return None

    # Layout no-ops
    def columns(self, n, **kw):
        count = n if isinstance(n, int) else len(n)
        return [_Noop() for _ in range(count)]
    def container(self, **kw):    return _Noop()
    def expander(self, *a, **kw): return _Noop()
    def empty(self, **kw):        return _Noop()
    def tabs(self, labels):       return [_Noop() for _ in labels]
    def form(self, *a, **kw):     return _Noop()

    # No-op display
    def metric(self, *a, **kw):   pass
    def image(self, *a, **kw):    pass
    def code(self, *a, **kw):     pass
    def caption(self, *a, **kw):  pass
    def subheader(self, *a, **kw): pass
    def header(self, *a, **kw):   pass
    def title(self, *a, **kw):    pass
    def divider(self):            pass
    def update(self, *a, **kw):   return self
    def status(self, *a, **kw):   return _Noop()
    def spinner(self, *a, **kw):  return _Noop()
    def chat_message(self, *a, **kw): return _Noop()


# ── Classe principale ToolContext ──────────────────────────────
class _ToolCtx:
    """
    Proxy thread-local per i tool CNA.
    I tool fanno: from core.toolkit import ctx
    """

    # ── Display — cattura eventi ───────────────────────────────
    def error(self, text="", *a, **kw):
        _messages().append({"type": "error", "text": str(text)})

    def success(self, text="", *a, **kw):
        _messages().append({"type": "success", "text": str(text)})

    def warning(self, text="", *a, **kw):
        _messages().append({"type": "warning", "text": str(text)})

    def info(self, text="", *a, **kw):
        _messages().append({"type": "info", "text": str(text)})

    def write(self, *args, **kw):
        if args:
            _messages().append({"type": "log", "text": str(args[0])})

    def markdown(self, text="", *a, **kw):
        if text:
            _messages().append({"type": "log", "text": str(text)})

    # ── Dataframe / tabella ────────────────────────────────────
    def dataframe(self, df=None, *a, **kw):
        if df is not None:
            _messages().append(_df_to_event(df))

    def table(self, df=None, *a, **kw):
        if df is not None:
            _messages().append(_df_to_event(df))

    # ── Progress ───────────────────────────────────────────────
    def progress(self, value=0, text="", *a, **kw):
        v = float(value) if value is not None else 0.0
        if v > 1:
            v /= 100.0
        _messages().append({"type": "progress", "value": v, "text": str(text) if text else ""})
        return self  # permette bar.progress(n)

    # ── No-op display ─────────────────────────────────────────
    def metric(self, *a, **kw):    pass
    def image(self, *a, **kw):     pass
    def code(self, *a, **kw):      pass
    def caption(self, *a, **kw):   pass
    def subheader(self, *a, **kw): pass
    def header(self, *a, **kw):    pass
    def title(self, *a, **kw):     pass
    def divider(self):             pass
    def exception(self, *a, **kw): pass
    def balloons(self):            pass
    def snow(self):                pass
    def toast(self, *a, **kw):     pass
    def json(self, *a, **kw):      pass
    def pyplot(self, *a, **kw):    pass
    def plotly_chart(self, *a, **kw): pass
    def altair_chart(self, *a, **kw): pass
    def line_chart(self, *a, **kw):   pass
    def bar_chart(self, *a, **kw):    pass
    def area_chart(self, *a, **kw):   pass
    def html(self, *a, **kw):         pass

    # ── Controllo flusso — no-op ───────────────────────────────
    def rerun(self):   pass
    def stop(self):    pass

    # ── Config ────────────────────────────────────────────────
    def set_page_config(self, **kw): pass

    # ── Layout / containers ────────────────────────────────────
    def __enter__(self): return self
    def __exit__(self, *a): pass

    def columns(self, n, **kw) -> list:
        count = n if isinstance(n, int) else len(n)
        return [_Noop() for _ in range(count)]

    def container(self, **kw):    return _Noop()
    def expander(self, *a, **kw): return _Noop()
    def spinner(self, *a, **kw):  return _Noop()
    def status(self, *a, **kw):   return _Noop()
    def form(self, *a, **kw):     return _Noop()
    def empty(self, **kw):        return _Noop()
    def tabs(self, labels):       return [_Noop() for _ in labels]
    def chat_message(self, *a, **kw): return _Noop()
    def sidebar(self):            return _Noop()
    def update(self, *a, **kw):   return self

    # ── Widget input → sempre default ─────────────────────────
    def text_input(self, label="", value="", **kw):   return value
    def text_area(self, label="", value="", **kw):    return value
    def number_input(self, label="", value=0, **kw):  return value
    def checkbox(self, label="", value=False, **kw):  return value
    def color_picker(self, label="", value="#000000", **kw): return value
    def slider(self, label="", min_value=0, max_value=100, value=0, **kw): return value
    def date_input(self, *a, **kw): return None
    def time_input(self, *a, **kw): return None
    def file_uploader(self, *a, **kw): return None
    def chat_input(self, *a, **kw):    return None

    def radio(self, label="", options=None, index=0, **kw):
        return options[index] if options else None

    def selectbox(self, label="", options=None, index=0, **kw):
        return options[index] if options else None

    def multiselect(self, label="", options=None, default=None, **kw):
        return default if default is not None else []

    # ── Button → sempre False ──────────────────────────────────
    def button(self, *a, **kw):              return False
    def form_submit_button(self, *a, **kw):  return False
    def download_button(self, *a, **kw):     return False
    def link_button(self, *a, **kw):         return None

    # ── Caching ────────────────────────────────────────────────
    @staticmethod
    def cache_data(func=None, **kwargs):
        if callable(func):
            try:
                return functools.lru_cache(maxsize=64)(func)
            except TypeError:
                return func
        def decorator(f):
            try:
                return functools.lru_cache(maxsize=64)(f)
            except TypeError:
                return f
        return decorator

    cache_resource = cache_data
    experimental_memo = cache_data
    experimental_singleton = cache_data

    # ── Session state ──────────────────────────────────────────
    @property
    def session_state(self) -> _SessionState:
        return _SESSION_STATE

    # ── Fragment / dialog decorators ───────────────────────────
    @staticmethod
    def fragment(*a, **kw):
        if a and callable(a[0]):
            return a[0]
        return lambda f: f

    @staticmethod
    def dialog(*a, **kw):
        if a and callable(a[0]):
            return a[0]
        return lambda f: f


# ── Singleton pubblici ─────────────────────────────────────────
_SESSION_STATE = _SessionState()
ctx = _ToolCtx()
