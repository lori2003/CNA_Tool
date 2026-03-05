"""
streamlit.py — Shim leggero per Toolbox CNA
============================================
Rimpiazza il pacchetto streamlit (non installato) con implementazioni
compatibili che si integrano col backend FastAPI.

I tool fanno:  import streamlit as st
Python trova questo file (via sys.path root) invece del pacchetto vero.

Caratteristiche:
- Thread-safe: capture messaggi e session_state isolati per richiesta
- st.cache_data  → functools.lru_cache (caching Python nativo)
- st.session_state → dict thread-local
- Funzioni display (info/warning/...) → catturate per dyninfo endpoint
- Widget input → ritornano il valore di default
- Layout/container → context manager no-op
"""
from __future__ import annotations

import functools
import threading
from typing import Any, List, Optional

# ═══════════════════════════════════════════════════════════════
#  THREAD-LOCAL STORAGE
# ═══════════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════════
#  SESSION STATE  (proxy thread-local, API identica a streamlit)
# ═══════════════════════════════════════════════════════════════
class _SessionState:
    def __getitem__(self, k):          return _state().get(k)
    def __setitem__(self, k, v):       _state()[k] = v
    def __delitem__(self, k):          _state().pop(k, None)
    def __contains__(self, k):         return k in _state()
    def __iter__(self):                return iter(_state())
    def get(self, k, default=None):    return _state().get(k, default)
    def keys(self):                    return _state().keys()
    def values(self):                  return _state().values()
    def items(self):                   return _state().items()
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


session_state = _SessionState()


# ═══════════════════════════════════════════════════════════════
#  CONTEXT MANAGER GENERICO
#  Usato da: st.columns(), st.container(), st.expander(), ecc.
# ═══════════════════════════════════════════════════════════════
class _Ctx:
    def __enter__(self):          return self
    def __exit__(self, *a):       pass
    # Metodi display (se usato come "with col1: col1.info(...)")
    def info(self, t='', *a, **kw):    _messages().append({"type": "info",    "text": str(t)})
    def warning(self, t='', *a, **kw): _messages().append({"type": "warning", "text": str(t)})
    def error(self, t='', *a, **kw):   _messages().append({"type": "error",   "text": str(t)})
    def success(self, t='', *a, **kw): _messages().append({"type": "success", "text": str(t)})
    def write(self, *a, **kw):
        if a: _messages().append({"type": "markdown", "text": str(a[0])})
    def markdown(self, t='', *a, **kw):
        if t: _messages().append({"type": "markdown", "text": str(t)})
    # Widget → default
    def text_input(self, label='', value='', **kw):   return value
    def text_area(self, label='', value='', **kw):    return value
    def number_input(self, label='', value=0, **kw):  return value
    def button(self, *a, **kw):                       return False
    def checkbox(self, label='', value=False, **kw):  return value
    def selectbox(self, label='', options=None, index=0, **kw):
        return options[index] if options else None
    def radio(self, label='', options=None, index=0, **kw):
        return options[index] if options else None
    def multiselect(self, label='', options=None, default=None, **kw):
        return default or []
    def slider(self, label='', min_value=0, max_value=100, value=0, **kw): return value
    def file_uploader(self, *a, **kw): return None
    # Layout
    def columns(self, n, **kw):
        count = n if isinstance(n, int) else len(n)
        return [_Ctx() for _ in range(count)]
    def container(self, **kw):     return _Ctx()
    def expander(self, *a, **kw):  return _Ctx()
    def form(self, *a, **kw):      return _Ctx()
    def tabs(self, labels):        return [_Ctx() for _ in labels]
    # No-op display
    def metric(self, *a, **kw):     pass
    def dataframe(self, *a, **kw):  pass
    def table(self, *a, **kw):      pass
    def image(self, *a, **kw):      pass
    def code(self, *a, **kw):       pass
    def json(self, *a, **kw):       pass
    def caption(self, *a, **kw):    pass
    def subheader(self, *a, **kw):  pass
    def header(self, *a, **kw):     pass
    def title(self, *a, **kw):      pass
    def text(self, *a, **kw):       pass
    def divider(self):              pass
    def empty(self, **kw):          return _Ctx()
    def progress(self, *a, **kw):   return _Ctx()
    def update(self, *a, **kw):     return self
    def status(self, *a, **kw):     return _Ctx()
    def spinner(self, *a, **kw):    return _Ctx()
    def chat_message(self, *a, **kw): return _Ctx()


# ═══════════════════════════════════════════════════════════════
#  DISPLAY — cattura messaggi per dyninfo / log
# ═══════════════════════════════════════════════════════════════
def info(text='', *a, **kw):
    _messages().append({"type": "info",    "text": str(text)})

def warning(text='', *a, **kw):
    _messages().append({"type": "warning", "text": str(text)})

def error(text='', *a, **kw):
    _messages().append({"type": "error",   "text": str(text)})

def success(text='', *a, **kw):
    _messages().append({"type": "success", "text": str(text)})

def write(*args, **kw):
    if args:
        _messages().append({"type": "markdown", "text": str(args[0])})

def markdown(text='', *a, **kw):
    if text:
        _messages().append({"type": "markdown", "text": str(text)})

def metric(*a, **kw):    pass
def dataframe(*a, **kw): pass
def table(*a, **kw):     pass
def text(*a, **kw):      pass
def caption(*a, **kw):   pass
def subheader(*a, **kw): pass
def header(*a, **kw):    pass
def title(*a, **kw):     pass
def divider():           pass
def exception(*a, **kw): pass
def balloons():          pass
def snow():              pass
def toast(*a, **kw):     pass
def code(*a, **kw):      pass
def json(*a, **kw):      pass
def image(*a, **kw):     pass
def audio(*a, **kw):     pass
def video(*a, **kw):     pass
def map(*a, **kw):       pass
def html(*a, **kw):      pass
def pyplot(*a, **kw):    pass
def plotly_chart(*a, **kw):      pass
def altair_chart(*a, **kw):      pass
def line_chart(*a, **kw):        pass
def bar_chart(*a, **kw):         pass
def area_chart(*a, **kw):        pass
def vega_lite_chart(*a, **kw):   pass
def graphviz_chart(*a, **kw):    pass


# ═══════════════════════════════════════════════════════════════
#  LAYOUT / CONTAINERS
# ═══════════════════════════════════════════════════════════════
def columns(n, **kw) -> list:
    count = n if isinstance(n, int) else len(n)
    return [_Ctx() for _ in range(count)]

def container(**kw):        return _Ctx()
def expander(*a, **kw):     return _Ctx()
def spinner(*a, **kw):      return _Ctx()
def status(*a, **kw):       return _Ctx()
def form(*a, **kw):         return _Ctx()
def empty(**kw):            return _Ctx()
def tabs(labels):           return [_Ctx() for _ in labels]
def chat_message(*a, **kw): return _Ctx()

sidebar = _Ctx()

# ═══════════════════════════════════════════════════════════════
#  PROGRESS
# ═══════════════════════════════════════════════════════════════
def progress(*a, **kw): return _Ctx()


# ═══════════════════════════════════════════════════════════════
#  INPUT WIDGETS  →  ritornano sempre il default
# ═══════════════════════════════════════════════════════════════
def text_input(label='', value='', **kw):        return value
def text_area(label='', value='', **kw):         return value
def number_input(label='', value=0, **kw):       return value
def slider(label='', min_value=0, max_value=100, value=0, **kw): return value
def checkbox(label='', value=False, **kw):       return value
def color_picker(label='', value='#000000', **kw): return value

def radio(label='', options=None, index=0, **kw):
    return options[index] if options else None

def selectbox(label='', options=None, index=0, **kw):
    return options[index] if options else None

def multiselect(label='', options=None, default=None, **kw):
    return default if default is not None else []

def date_input(*a, **kw):  return None
def time_input(*a, **kw):  return None
def file_uploader(*a, **kw): return None
def chat_input(*a, **kw):  return None


# ═══════════════════════════════════════════════════════════════
#  BOTTONI  →  sempre False / non cliccati
# ═══════════════════════════════════════════════════════════════
def button(*a, **kw):              return False
def form_submit_button(*a, **kw):  return False
def download_button(*a, **kw):     return False
def link_button(*a, **kw):         return None


# ═══════════════════════════════════════════════════════════════
#  CACHING  →  lru_cache Python nativo
# ═══════════════════════════════════════════════════════════════
def cache_data(func=None, **kwargs):
    """
    Rimpiazza @st.cache_data con @functools.lru_cache.
    Supporta sia @st.cache_data che @st.cache_data(show_spinner=...).
    """
    if callable(func):
        # Uso diretto come decorator: @st.cache_data
        try:
            return functools.lru_cache(maxsize=64)(func)
        except TypeError:
            return func
    # Uso come factory: @st.cache_data(show_spinner="...")
    def decorator(f):
        try:
            return functools.lru_cache(maxsize=64)(f)
        except TypeError:
            return f
    return decorator

cache_resource        = cache_data
experimental_memo     = cache_data
experimental_singleton = cache_data


# ═══════════════════════════════════════════════════════════════
#  CONTROLLO FLUSSO
# ═══════════════════════════════════════════════════════════════
def rerun(): pass
def stop():  pass


# ═══════════════════════════════════════════════════════════════
#  CONFIG / DECORATORI AVANZATI
# ═══════════════════════════════════════════════════════════════
def set_page_config(**kw): pass

def fragment(*a, **kw):
    if a and callable(a[0]): return a[0]
    return lambda f: f

def dialog(*a, **kw):
    if a and callable(a[0]): return a[0]
    return lambda f: f
