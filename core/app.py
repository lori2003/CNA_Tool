# ======================================================================================
# 🔒 APP.PY STABILITY CONTRACT & PROTECTION MANIFEST
# ======================================================================================
# Questo file costituisce l'infrastruttura portante (CORE) della Toolbox.
# Per garantire la continuità operativa, ogni AI o sviluppatore deve seguire queste regole:
#
# 1. LOGICA ESISTENTE SACRA: Non modificare, rinominare o eliminare le logiche e le 
#    funzioni che gestiscono il funzionamento degli altri script Python esistenti. 
#    La stabilità dei tool già operativi (Lombardia, Sicilia, ecc.) è la priorità.
#
# 2. EVOLUZIONE PER ADDIZIONE: Le nuove funzionalità, miglioramenti o patch devono essere 
#    implementate in blocchi di codice separati, moduli helper o funzioni aggiuntive. 
#    Evitare di "rimescolare" il codice esistente per minimizzare il rischio di bug.
#
# 3. ISOLAMENTO DELLE INTERFERENZE: Assicurarsi che ogni aggiunta non crei conflitti 
#    con le chiavi di session_state (up_, param_), i nomi delle variabili globali 
#    o il comportamento della UI già consolidato.
#
# 4. ANALISI E DEBUG CHIRURGICO: Il debug è incoraggiato, ma ogni intervento correttivo 
#    sul Core deve essere documentato e testato per non alterare l'output degli altri tool.
#
# ⚠️ REGOLA D'ORO: Prima di toccare il Core, verifica se il risultato può essere 
# ottenuto agendo direttamente sullo script del singolo tool o aggiungendo codice a valle.
# ======================================================================================

# ======================================================================================
# 📖 DOCUMENTATION STANDARD (TOOL DESCRIPTION SCHEMA)
# ======================================================================================
# Ogni script dentro /tools deve contenere nel dizionario TOOL['description'] uno schema
# standardizzato in 4 blocchi per garantire professionalità e uniformità:
#
# #### 📌 1. FINALITÀ DEL TOOL
# [Obiettivo principale e problema risolto]
#
# #### 🚀 2. COME UTILIZZARLO
# [Istruzioni passo-passo per l'utente finale]
#
# #### 🧠 3. LOGICA DI ELABORAZIONE (SPECIFICHE)
# [Dettagli tecnici speculari al codice: es. gestione nomi, formule, Word/Excel engines]
#
# #### 📂 4. RISULTATO FINALE
# [Descrizione dei file o delle cartelle prodotte]
# ======================================================================================

from __future__ import annotations

import ast
import os
import importlib
import importlib.util
import io
import json
import pprint
import re
import sys
import tempfile
import zipfile
import subprocess
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st
try:
    import openai
except ImportError:
    openai = None

st.set_page_config(page_title="Toolbox CNA", page_icon="🧰", layout="wide")

# ------------------------------------------------------------
# Base setup
# ------------------------------------------------------------

# ------------------------------------------------------------
# Base setup
# ------------------------------------------------------------
ENV_BASE = os.getenv("TOOLBOX_HOME")
if ENV_BASE:
    BASE_DIR = Path(ENV_BASE).expanduser().resolve()
else:
    _here = Path(__file__).resolve().parent
    # Se tools/ non è nella stessa cartella (es. app.py è in core/), cerca nel parent
    if not (_here / "tools").exists() and (_here.parent / "tools").exists():
        BASE_DIR = _here.parent
    else:
        BASE_DIR = _here
TOOLS_DIR = BASE_DIR / "tools"
DATA_DIR = BASE_DIR / "data"  # opzionale: data/<Regione>/...

# Ensure imports from project root work even if Streamlit changes working dir
sys.path.insert(0, str(BASE_DIR))

SUPPORTED_INPUT_TYPES = {"txt_multi", "txt_single", "xlsx_single", "file_multi", "file_single", "warning", "info", "error", "success", "markdown"}
SUPPORTED_PARAM_TYPES = {"select", "radio", "checkbox", "number", "text", "textarea", "multiselect", "dynamic_info", "folder", "file_path_info"}

# ------------------------------------------------------------
# Prompt templates (copia & incolla)
# ------------------------------------------------------------
PROMPT_TEMPLATES: Dict[str, str] = {
    "A) Converti un .py in tool Toolbox (TOOL + run + out_dir)": """Trasforma il mio file Python in un *tool* compatibile con la mia Toolbox Streamlit.

REGOLE (OBBLIGATORIE)
1) Il file finale deve contenere:
   - TOOL = {...}
   - def run(..., out_dir: Path) -> List[Path]
2) run() deve:
   - NON usare input() o interazione da terminale
   - leggere eventuali input dai parametri (upload file) e/o dai params (widget)
   - scrivere uno o piu file dentro out_dir
   - ritornare una lista di Path dei file creati
3) Non importare streamlit nel tool.
4) I nomi dei parametri in run(...) devono combaciare con i key dichiarati in TOOL["inputs"] e TOOL["params"].
5) La cartella/regione la scelgo io (tools/<qualcosa>/...). La regione si deduce dalla cartella: non serve nel codice.

SCHEMA TOOL (minimo)
TOOL = {
  "id": "<id_unico>",
  "name": "<nome_mostrato_in_ui>",
  "description": "<breve>",
  "inputs": [ ... ],   # opzionale
  "params": [ ... ],   # opzionale
}

INPUTS SUPPORTATI (upload file)
- txt_multi | txt_single | xlsx_single
Esempio:
{"key":"file_txt","label":"File TXT","type":"txt_single","required":True}

PARAMS SUPPORTATI (widget UI)
- select | radio | checkbox | number | text | textarea
Esempio:
{"key":"azione","label":"Azione","type":"radio","options":["A","B"],"default":"A","required":True}

TASK
Ti incollero un file .py (anche con main()).
Tu devi:
- estrarre la logica in run(...)
- dichiarare TOOL con inputs/params necessari
- eliminare/ignorare input() / argparse / menu da terminale
- garantire output in out_dir e return List[Path]

DATI
Nome file finale: <NOME_FILE>.py
Codice originale: <INCOLLA QUI IL .PY>

OUTPUT
Rispondi con SOLO il contenuto completo del nuovo file <NOME_FILE>.py (niente spiegazioni).""",

    "B) Trasforma scelte (input/argparse/menu) in widget UI (params)": """Rendi il mio script *selezionabile in anticipo* nella UI, convertendo tutte le scelte in TOOL['params'] (widget).

CONTESTO
- In Streamlit non si usa input() da terminale.
- Le scelte/flag vanno dichiarate in TOOL["params"].
- I valori scelti arrivano in run(...) come argomenti normali.
- run(..., out_dir: Path) -> List[Path] deve restare valido e produrre file in out_dir.

TIPI PARAMS SUPPORTATI
- radio / select (scelte discrete) -> richiede options: [...]
- checkbox (booleani)
- number (min/max/step se deducibili)
- text / textarea (testo libero)

TASK
Ti incollero un file .py che contiene scelte tipo:
- input() / menu testuale
- argparse (flag/opzioni)
- if/elif basati su modalita/azione/tipo
- piu funzioni alternative

Tu devi:
1) identificare tutte le scelte possibili
2) creare TOOL["params"] coerente (key, label, type, options/default)
3) aggiornare run(...) includendo quei parametri (stesso nome di key)
4) rimuovere/ignorare input() e parsing da terminale: la decisione dipende SOLO dai params
5) scrivere i risultati in out_dir e ritornare List[Path]

DATI
Nome file finale: <NOME_FILE>.py
Codice originale: <INCOLLA QUI IL .PY>

OUTPUT
Rispondi con SOLO il file .py finale completo, pronto da mettere in tools/<cartella_che_scelgo>/. """,
}

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def _slug(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"[^0-9a-z]+", "_", s)
    return s.strip("_") or "x"

def _minify_code(code: str) -> str:
    """Rimuove commenti, docstring e righe vuote per risparmiare token."""
    # Rimuovi commenti a riga singola
    code = re.sub(r'#.*', '', code)
    # Rimuovi righe vuote
    lines = [line for line in code.splitlines() if line.strip()]
    return "\n".join(lines)


def _key_safe(s: str) -> str:
    return re.sub(r"[^0-9A-Za-z_]+", "_", str(s))[:180]


def _safe_mod_name(parts: Tuple[str, ...]) -> str:
    raw = "__".join(parts)
    raw = raw.replace(".py", "")
    raw = re.sub(r"[^0-9A-Za-z_]+", "_", raw)
    return f"tool__{raw}"


def _load_module_from_path(mod_name: str, path: Path):
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


def _find_tool_dict_span(txt: str) -> Optional[Tuple[int, int]]:
    """
    Trova lo span del dict literal assegnato a TOOL = {...}.
    Ritorna (start_index, end_index) dove end è esclusivo.
    """
    m = re.search(r"^\s*TOOL\s*=\s*\{", txt, flags=re.MULTILINE)
    if not m:
        return None
    start = m.end() - 1  # '{'
    i = start
    depth = 0
    in_str: Optional[str] = None
    esc = False
    while i < len(txt):
        ch = txt[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == in_str:
                in_str = None
        else:
            if ch in ("'", '"'):
                in_str = ch
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return (start, i + 1)
        i += 1
    return None


def _parse_tool_literal(tool_literal: str) -> Dict[str, Any]:
    """
    Parse robusto SOLO per dict literal (niente espressioni Python non-literal).
    Se il TOOL contiene roba non literal, rifiutiamo (così non rompiamo file).
    """
    try:
        obj = ast.literal_eval(tool_literal)
    except Exception as e:
        raise ValueError(f"TOOL non è un dict literal parsabile (ast.literal_eval fallisce): {e}")
    if not isinstance(obj, dict):
        raise ValueError("TOOL parsato ma non è un dict.")
    return obj


def _validate_inputs(inputs: Any) -> Tuple[bool, str]:
    if inputs is None:
        return True, ""
    if not isinstance(inputs, list):
        return False, "inputs deve essere una lista."
    for i, item in enumerate(inputs):
        if not isinstance(item, dict):
            return False, f"inputs[{i}] deve essere un dict."
        if "key" not in item or "type" not in item:
            return False, f"inputs[{i}] richiede almeno 'key' e 'type'."
        if item["type"] not in SUPPORTED_INPUT_TYPES:
            return False, f"inputs[{i}].type non supportato: {item['type']}"
    return True, ""


def _validate_params(params: Any) -> Tuple[bool, str]:
    if params is None:
        return True, ""
    if not isinstance(params, list):
        return False, "params deve essere una lista."
    for i, item in enumerate(params):
        if not isinstance(item, dict):
            return False, f"params[{i}] deve essere un dict."
        if "key" not in item or "type" not in item:
            return False, f"params[{i}] richiede almeno 'key' e 'type'."
        t = item["type"]
        if t not in SUPPORTED_PARAM_TYPES:
            return False, f"params[{i}].type non supportato: {t}"
        if t in ("select", "radio", "multiselect"):
            opts = item.get("options")
            if not isinstance(opts, list) or not opts:
                return False, f"params[{i}] ({t}) richiede options (lista non vuota)."
    return True, ""


def update_tool_fields_in_file(py_path: Path, updates: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Aggiorna in modo sicuro TOOL nel file:
    - legge TOOL dict literal
    - parse con ast.literal_eval (se non è literal -> non tocca)
    - aggiorna campi richiesti
    - riscrive SOLO il blocco TOOL = {...} usando pprint (Python-valid)
    """
    try:
        txt = py_path.read_text(encoding="utf-8")
    except Exception as e:
        return False, f"Non riesco a leggere il file: {e}"

    span = _find_tool_dict_span(txt)
    if not span:
        return False, "Non trovo 'TOOL = {...}' nel file."

    a, b = span
    tool_literal = txt[a:b]

    try:
        tool_obj = _parse_tool_literal(tool_literal)
    except Exception as e:
        return False, str(e)

    for k, v in updates.items():
        tool_obj[k] = v

    ok, msg = _validate_inputs(tool_obj.get("inputs"))
    if not ok:
        return False, msg
    ok, msg = _validate_params(tool_obj.get("params"))
    if not ok:
        return False, msg

    tool_dump = pprint.pformat(tool_obj, width=120, sort_dicts=False)
    txt2 = txt[:a] + tool_dump + txt[b:]

    try:
        py_path.write_text(txt2, encoding="utf-8")
    except Exception as e:
        return False, f"Non riesco a scrivere il file: {e}"

    return True, "TOOL aggiornato."


# ------------------------------------------------------------
# Tool discovery
# ------------------------------------------------------------
def discover_tools() -> List[Dict[str, Any]]:
    tools: List[Dict[str, Any]] = []
    if not TOOLS_DIR.exists():
        return tools

    for py in sorted(TOOLS_DIR.rglob("*.py")):
        if py.name == "__init__.py" or py.name.startswith("_"):
            continue
        
        # Skip files inside 'extension' folders (contain libraries, not tools)
        if "extension" in py.parts:
            continue

        rel = py.relative_to(TOOLS_DIR)
        parts = rel.parts
        region_folder = parts[0] if len(parts) >= 2 else "Generali"

        mod_name = _safe_mod_name(parts)

        try:
            mod = _load_module_from_path(mod_name, py)
        except Exception as e:
            tools.append(
                {
                    "uid": f"__error__{region_folder}__{py.stem}",
                    "id": py.stem,
                    "region": region_folder,
                    "name": f"❌ ERRORE import: {region_folder}/{py.stem}",
                    "description": f"Non riesco a importare:\n{py}\n\nErrore:\n{e}",
                    "inputs": [],
                    "params": [],
                    "runner": None,
                    "import_error": True,
                    "source_path": str(py),
                }
            )
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
        tool.setdefault("exact_txt", None)

        tool["uid"] = uid
        tool["region"] = region
        tool["runner"] = runner
        tool["dynamic_params"] = dynamic_params
        tool["module_obj"] = mod  # Vital for dynamic_info to find functions
        tool["import_error"] = False
        tool["source_path"] = str(py)

        if not callable(runner):
            tool["import_error"] = True
            tool["runner"] = None
            tool["description"] = (tool.get("description") or "") + "\n\n⚠️ Manca la funzione run(...)."

        for inp in tool.get("inputs", []):
            if inp.get("type") not in SUPPORTED_INPUT_TYPES:
                tool["import_error"] = True
                tool["runner"] = None
                tool["description"] = (tool.get("description") or "") + f"\n\n⚠️ Tipo input non supportato: {inp.get('type')}"

        for p in tool.get("params", []):
            if p.get("type") not in SUPPORTED_PARAM_TYPES:
                tool["import_error"] = True
                tool["runner"] = None
                tool["description"] = (tool.get("description") or "") + f"\n\n⚠️ Tipo parametro non supportato: {p.get('type')}"
            if p.get("type") in ("select", "radio", "multiselect") and not p.get("options"):
                tool["import_error"] = True
                tool["runner"] = None
                tool["description"] = (tool.get("description") or "") + f"\n\n⚠️ Parametro '{p.get('key')}' richiede options (lista)."

        tools.append(tool)

    def _sort_key(t: Dict[str, Any]):
        reg = (t.get("region") or "").lower()
        reg_rank = 0 if reg == "generali" else 1
        return (bool(t.get("import_error")), reg_rank, reg, (t.get("name") or "").lower())

    tools.sort(key=_sort_key)
    return tools


@st.cache_resource
def load_tools_cached() -> List[Dict[str, Any]]:
    return discover_tools()

# ------------------------------------------------------------
# Gestione Modalità Assistente Full (Nuova Finestra)
# ------------------------------------------------------------
query_params = st.query_params
if query_params.get("mode") == "assistant":
    st.title("✨ Assistente Intelligente CNA")
    
    if "assistant_chat_history" not in st.session_state:
        st.session_state.assistant_chat_history = []

    # Caricamento Tool per Knowledge Base con CACHE
    tools_list = load_tools_cached()
    
    if "assistant_kb_cache" not in st.session_state:
        kb = []
        for t in tools_list:
            if not t.get("import_error"):
                kb.append(f"- {t.get('name')}: {t.get('description')[:120]}")
        st.session_state.assistant_kb_cache = "\n".join(kb)
    kb_text = st.session_state.assistant_kb_cache
    
    # Visualizzazione Chat Full
    for msg in st.session_state.assistant_chat_history:
        with st.chat_message("user" if msg["role"] == "user" else "assistant"):
            st.write(msg["content"])

    # Input Chat Full
    if prompt := st.chat_input("Come posso aiutarti?"):
        if "assistant_chat_history" not in st.session_state:
            st.session_state.assistant_chat_history = []
        st.session_state.assistant_chat_history.append({"role": "user", "content": prompt})
        
        try:
            # Rilevamento Dinamico del Tool Attivo (Cross-Tab)
            current_tool_uid = None
            try:
                ctx_file = DATA_DIR / "active_context.json"
                if ctx_file.exists():
                    with open(ctx_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        current_tool_uid = data.get("selected_tool_uid")
            except:
                pass
            
            tool_context_info = ""
            if current_tool_uid:
                target = next((t for t in tools_list if t.get("uid") == current_tool_uid), None)
                if target and target.get("source_path"):
                    cache_key = f"src_min_cache_{_key_safe(current_tool_uid)}"
                    if cache_key not in st.session_state:
                        raw_src = Path(target["source_path"]).read_text(encoding="utf-8")
                        st.session_state[cache_key] = _minify_code(raw_src)
                    
                    src_snippet = st.session_state[cache_key]
                    # SMART CONTEXT: Se c'è un tool, non inviamo la KB per risparmiare migliaia di token
                    kb_text = "" 
                    tool_context_info = f"\n\nTOOL: {target['name']}\nCODE:\n```python\n{src_snippet}\n```"

            import openai
            client = openai.OpenAI(
                base_url=st.session_state.get("assistant_base_url", "https://openrouter.ai/api/v1"),
                api_key=st.secrets["openrouter"]["api_key"]
            )
            
            # FLASH PROMPT: Risposte veloci ed essenziali
            instructions = (
                "Sei un assistente tecnico CNA FLASH esperto di Python e Streamlit. Rispondi in modo ESSENZIALE, BREVE e PROFESSIONALE. "
                "Il tuo compito è aiutare l'utente con i tool della Toolbox. Hai accesso alla documentazione e al codice del tool attivo."
                f"\n\nKNOWLEDGE BASE:\n{kb_text}{tool_context_info}"
            )
            
            api_messages = []
            history = st.session_state.assistant_chat_history[-10:]
            
            for i, m in enumerate(history):
                content = m["content"]
                if i == 0: # Iniettiamo le istruzioni nel primo messaggio della finestra di contesto
                    content = f"[ISTRUZIONI DI SISTEMA: {instructions}]\n\nDOMANDA UTENTE: {content}"
                
                role = "assistant" if m["role"] in ["model", "assistant"] else "user"
                api_messages.append({"role": role, "content": content})

            with st.status("🔮 Pensando...", expanded=True) as status:
                try:
                    resp = client.chat.completions.create(
                        model=st.session_state.assistant_model_id,
                        messages=api_messages,
                        stream=True
                    )
                    status.update(label="⚡ Risposta in arrivo...", state="running", expanded=False)
                except Exception as e:
                    err_str = str(e).lower()
                    if "429" in err_str or "rate limit" in err_str:
                        status.update(label="⏳ Limite raggiunto...", state="error")
                        wait_time = 20
                        placeholder = st.empty()
                        for i in range(wait_time, 0, -1):
                            placeholder.warning(f"⚠️ Limite al minuto raggiunto. Prossimo invio tra {i} secondi...")
                            time.sleep(1)
                        placeholder.empty()
                        st.error("Puoi riprovare ora!")
                        st.stop()
                    raise e
            
            with st.chat_message("assistant"):
                ai_content = st.write_stream(resp)
            
            st.session_state.assistant_chat_history.append({"role": "assistant", "content": ai_content})
            st.rerun()
        except Exception as e:
            st.error(f"Errore connessione OpenRouter: {str(e)}")
    st.stop()


def set_selected_tool(uid: str) -> None:
    st.session_state["selected_tool_uid"] = uid
    # Sincronizzazione Context per Assistente Full-Screen
    try:
        ctx_file = DATA_DIR / "active_context.json"
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(ctx_file, "w", encoding="utf-8") as f:
            json.dump({"selected_tool_uid": uid}, f)
    except:
        pass


def get_selected_tool(tools: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    uid = st.session_state.get("selected_tool_uid")
    if not uid:
        return None
    for t in tools:
        if t.get("uid") == uid:
            return t
    return None


def group_tools_by_region(tools: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for t in tools:
        reg = t.get("region") or "Generali"
        grouped.setdefault(reg, []).append(t)
    return grouped


# --- FUNZIONI DI SUPPORTO PER L'ASSISTENTE ---
def get_project_codebase_summary(root_dir):
    """Scansiona la codebase e restituisce un sommario con path e docstring."""
    summary = []
    root_path = Path(root_dir)
    exclude_dirs = {".venv", "venv", "__pycache__", ".git", ".idea", ".vscode", "node_modules", "data", "logs"}
    
    for path in root_path.rglob("*.py"):
        # Filtra directory escluse
        if any(part in exclude_dirs for part in path.parts):
            continue
            
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
            # Estrazione docstring (molto semplice)
            docstring = "Nessuna descrizione."
            match = re.search(r'^"""(.*?)"""', content, re.DOTALL)
            if match:
                docstring = match.group(1).strip()[:200].replace("\n", " ") + "..."
            
            rel_path = path.relative_to(root_path)
            summary.append(f"- FILE: {rel_path}\n  INFO: {docstring}")
        except Exception:
            continue
            
    return "\n".join(summary[:100])

@st.fragment
def render_ai_assistant(tools):
    if openai and "openrouter" in st.secrets:
        # Intestazione e Tasto Chiudi
        c_head, c_close = st.columns([0.9, 0.1])
        with c_head:
            st.markdown("### ✨ Assistente Intelligente CNA")
        with c_close:
            if st.button("❌", key="close_assistant_side", help="Torna al Tool"):
                st.session_state["show_assistant"] = False
                st.rerun()
        
        st.markdown("""
            <script>
            function triggerSidebar(action) {
                const doc = window.parent.document;
                const labels = action === 'collapse' ? ['Collapse sidebar', 'Chiudi sidebar'] : ['Expand sidebar', 'Apri sidebar'];
                let btn = null;
                for (let lab of labels) {
                    btn = doc.querySelector(`button[aria-label="${lab}"]`) || doc.querySelector(`button[title="${lab}"]`);
                    if (btn) break;
                }
                if (btn) { btn.click(); }
            }
            setTimeout(() => triggerSidebar('collapse'), 300);
            </script>
        """, unsafe_allow_html=True)

        if "assistant_chat_history" not in st.session_state:
            st.session_state.assistant_chat_history = []

        # Chat Bubbles Style - Scroll gestito dal Box esterno Full-Height
        # Rimuoviamo l'altezza fissa qui per lasciare che il box padre gestisca lo spazio
        chat_scroll_area = st.container(border=False)
        with chat_scroll_area:
            for msg in st.session_state.assistant_chat_history:
                with st.chat_message("user" if msg["role"] == "user" else "assistant"):
                    st.write(msg["content"])

        if prompt := st.chat_input("Chiedi all'assistente...", key="assistant_inline_input"):
            st.session_state.assistant_chat_history.append({"role": "user", "content": prompt})
            try:
                # Cache kb_text (knowledge base of tools)
                if "kb_text_cache" not in st.session_state:
                    kb = []
                    for t in tools:
                        if not t.get("import_error"):
                            kb.append(f"- TOOL: {t.get('name')}\n  DESC: {t.get('description')}")
                    st.session_state["kb_text_cache"] = "\n\n".join(kb)
                kb_text = st.session_state["kb_text_cache"]
                
                ctx_info = ""
                try:
                    ctx_file = DATA_DIR / "active_context.json"
                    if ctx_file.exists():
                        with open(ctx_file, "r", encoding="utf-8") as f:
                            uid = json.load(f).get("selected_tool_uid")
                            target = next((t for t in tools if t.get("uid") == uid), None)
                            if target and target.get("source_path"):
                                # Cache Minified Source Code
                                cache_key = f"src_min_cache_{_key_safe(uid)}"
                                if cache_key not in st.session_state:
                                    raw_src = Path(target["source_path"]).read_text(encoding="utf-8")
                                    st.session_state[cache_key] = _minify_code(raw_src)
                                
                                src_snippet = st.session_state[cache_key]
                                
                                # SMART CONTEXT: Priorità al tool attivo
                                kb_text = "" 

                                current_state = []
                                prefix_up = f"up_{_key_safe(uid)}_"
                                prefix_param = f"param_{_key_safe(uid)}_"
                                for k, v in st.session_state.items():
                                    if k.startswith(prefix_up):
                                        fname = v.name if hasattr(v, 'name') else (f"{len(v)} file" if isinstance(v, list) else str(v))
                                        current_state.append(f"- INPUT '{k[len(prefix_up):]}': {fname}")
                                    elif k.startswith(prefix_param):
                                        current_state.append(f"- PARAM '{k[len(prefix_param):]}': {v}")
                                
                                state_str = "\n".join(current_state) if current_state else "None"
                                ctx_info = (
                                    f"TOOL: {target['name']}\n"
                                    f"STATE: {state_str}\n"
                                    f"CODE:\n```python\n{src_snippet}\n```"
                                )
                except: pass

                # codebase_map = get_project_codebase_summary(BASE_DIR) 
                
                client = openai.OpenAI(
                    base_url=st.session_state.get("assistant_base_url", "https://openrouter.ai/api/v1"), 
                    api_key=st.secrets["openrouter"]["api_key"]
                )
                instructions = (
                    f"Sei l'Assistente CNA. Rispondi BREVE.\n"
                    f"CONTESTO:\n{kb_text}\n{ctx_info}"
                )
                
                messages = []
                history = st.session_state.assistant_chat_history[-8:]
                for i, m in enumerate(history):
                    content = m["content"]
                    if i == 0: content = f"[ISTRUZIONI: {instructions}]\n\n{content}"
                    messages.append({"role": "assistant" if m["role"] in ["model", "assistant"] else "user", "content": content})
                
                with chat_scroll_area:
                    with st.status("🔮 Elaborazione istruzioni...", expanded=True) as status:
                        try:
                            response_stream = client.chat.completions.create(
                                model=st.session_state.assistant_model_id, 
                                messages=messages,
                                stream=True
                            )
                            status.update(label="⚡ Risposta in arrivo...", state="running", expanded=False)
                        except Exception as e:
                            err_str = str(e).lower()
                            if "429" in err_str or "rate limit" in err_str:
                                status.update(label="⏳ Limite raggiunto...", state="error")
                                wait_time = 20
                                placeholder = st.empty()
                                for i in range(wait_time, 0, -1):
                                    placeholder.warning(f"⚠️ Limite al minuto raggiunto. Prossimo invio tra {i} secondi...")
                                    time.sleep(1)
                                placeholder.empty()
                                st.rerun()
                            raise e

                    with st.chat_message("assistant"):
                        full_response = st.write_stream(response_stream)
                
                st.session_state.assistant_chat_history.append({"role": "assistant", "content": full_response or "Error"})
                st.rerun()
            except Exception as e:
                st.session_state.assistant_chat_history.append({"role": "assistant", "content": f"❌ Errore: {str(e)}"})
                st.rerun()

def sidebar_regions(tools: List[Dict[str, Any]]) -> None:
    st.sidebar.markdown("## 🧰 Toolbox")
    
    # --- REGOLAZIONE INTENSITÀ BLU ---
    with st.sidebar.expander("🎨 Personalizza Tema", expanded=False):
        new_l = st.slider("Intensità Blu Sidebar", 10, 60, st.session_state["sidebar_lightness"])
        if new_l != st.session_state["sidebar_lightness"]:
            st.session_state["sidebar_lightness"] = new_l
            save_theme_config({"sidebar_lightness": new_l})
            st.rerun()

    # --- TOGGLE ASSISTENTE (Side Panel) ---
    if st.session_state["show_assistant"]:
        if st.sidebar.button("❌ Chiudi Assistente CNA", width="stretch", key="close_assistant_btn"):
            st.session_state["show_assistant"] = False
            st.rerun()
    else:
        if st.sidebar.button("✨ Apri Assistente CNA", width="stretch", key="toggle_assistant_btn"):
            st.session_state["show_assistant"] = True
            st.rerun()
            
    # --- CONFIGURAZIONE AI ---
    with st.sidebar.expander("⚙️ Configurazione AI", expanded=False):
        c_mod = st.text_input("Model ID", value=st.session_state.get("assistant_model_id", "arcee-ai/trinity-large-preview:free"), key="cfg_model_id")
        c_url = st.text_input("Base URL", value=st.session_state.get("assistant_base_url", "https://openrouter.ai/api/v1"), key="cfg_base_url")
        
        if st.button("💾 Salva Configurazione", width="stretch"):
            st.session_state.assistant_model_id = c_mod
            st.session_state.assistant_base_url = c_url
            save_ai_config({"model_id": c_mod, "base_url": c_url})
            st.toast("✅ Configurazione salvata con successo!")

    st.sidebar.markdown("---")

    # Root override (senza caption tecnica)
    root_val = st.sidebar.text_input("Cartella Radice Toolbox", value=str(BASE_DIR), key="root_override_input")
    if st.sidebar.button("Imposta root", help="Aggiorna TOOLBOX_HOME e ricarica la pagina.", width="stretch"):
        new_root = str(Path(root_val).expanduser().resolve())
        os.environ["TOOLBOX_HOME"] = new_root
        st.cache_resource.clear()
        st.cache_data.clear()
        st.session_state.pop("selected_tool_uid", None)
        st.rerun()

    q = st.sidebar.text_input("🔎 Cerca (tool/regione)", value="", placeholder="Es. Lazio, Prospetto, 2025…")

    if st.sidebar.button("🔄 Aggiorna Elenco e Cache", help="Ricarica la lista dei tool e svuota completamente la cache (utile se modifichi il codice).", width="stretch"):
        st.cache_resource.clear()
        st.cache_data.clear()
        # Pulizia completa dello stato dei widget per tutti i tool
        keys_to_clear = [k for k in st.session_state.keys() if k.startswith("up_") or k.startswith("param_") or k.startswith("_init_done_")]
        for k in keys_to_clear:
            del st.session_state[k]
        st.rerun()


    st.sidebar.divider()

    qn = q.strip().lower()

    def _match(t: Dict[str, Any]) -> bool:
        return (
            qn in str(t.get("name", "")).lower()
            or qn in str(t.get("id", "")).lower()
            or qn in str(t.get("uid", "")).lower()
            or qn in str(t.get("region", "")).lower()
        )

    view_tools = tools if not qn else [t for t in tools if _match(t)]
    grouped = group_tools_by_region(view_tools)

    if not grouped:
        st.sidebar.info("Nessun tool corrisponde alla ricerca.")
        return

    if get_selected_tool(tools) is None and tools:
        first_valid = next((t for t in tools if not t.get("import_error")), tools[0])
        set_selected_tool(first_valid.get("uid", ""))

    selected = get_selected_tool(tools)
    selected_region = selected.get("region") if selected else None

    # Ordinamento personalizzato
    custom_order = st.session_state.get("region_order", [])
    def _region_sort_key(r):
        r_str = str(r)
        if r_str in custom_order:
            return (0, custom_order.index(r_str))
        # Default: Generali per primo (se non in custom_order), poi alfabetico
        rank = 1 if r_str.lower() == "generali" else 2
        return (rank, r_str.lower())

    regions = sorted(grouped.keys(), key=_region_sort_key)

    for region_key in regions:
        # Espandi solo se c'è una ricerca attiva. Altrimenti chiuse di default.
        expanded = bool(qn)
        with st.sidebar.expander(f"📍 {region_key}", expanded=expanded):
            data_path = DATA_DIR / str(region_key)
            if data_path.exists() and data_path.is_dir():
                files = sorted([p.name for p in data_path.glob("*") if p.is_file()])
                if files:
                    st.caption("📁 File dati (locali)")
                    for n in files[:8]:
                        st.write(f"• {n}")
                    if len(files) > 8:
                        st.write(f"• … (+{len(files) - 8} altri)")
                    st.divider()

            for t in grouped[region_key]:
                disabled = bool(t.get("import_error")) or (t.get("runner") is None)
                label = str(t.get("name", t.get("id", "tool")))
                
                # Layout colonne: Tool + Matita + (opzionale) Email
                has_email = bool(t.get("email_reminder"))
                cols = st.columns([0.7, 0.15, 0.15]) if has_email else st.columns([0.8, 0.2])
                c1, c2 = cols[0], cols[1]
                with c1:
                    st.button(
                        label,
                        key=f"nav_{_key_safe(str(t.get('uid')))}",
                        disabled=disabled,
                        width="stretch",
                        on_click=set_selected_tool,
                        args=(str(t.get("uid")),),
                    )
                with c2:
                    if st.button("✏️", key=f"edit_{_key_safe(str(t.get('uid')))}", help="Apri posizione file"):
                        path = t.get("source_path")
                        if path:
                            subprocess.run(['explorer', '/select,', str(Path(path).absolute())])
                if has_email:
                    email_val = t.get("email_reminder")
                    email_tip = email_val if isinstance(email_val, str) else "File da mandare all'amministrazione"
                    with cols[2]:
                        st.button("📧", key=f"email_{_key_safe(str(t.get('uid')))}", help=email_tip)

    st.sidebar.divider()



# ------------------------------------------------------------
# Main UI helpers
# ------------------------------------------------------------
def save_upload_to_tmp(tmpdir: Path, f) -> Path:
    p = tmpdir / f.name
    buf = f.getbuffer()
    if isinstance(buf, io.BytesIO):
        data = buf.getvalue()
    else:
        try:
            data = bytes(buf)
        except Exception:
            data = buf if isinstance(buf, (bytes, bytearray, memoryview)) else b""
    p.write_bytes(data)
    return p


def render_params_list(params: List[Dict[str, Any]], tool_uid: str, tool_module: Any = None) -> Dict[str, Any]:
    """
    Renderizza la lista dei parametri.
    tool_module è opzionale: se passato, permette a tipi come 'dynamic_info' di chiamare funzioni del modulo.
    """
    values: Dict[str, Any] = {}
    if not params:
        return values

    current_section = None
    first_param = True

    for p in params:
        sec = p.get("section")
        if first_param and not sec:
            st.markdown("### Opzioni")
            current_section = "Opzioni"
        if sec and sec != current_section:
            st.markdown(f"### {sec}")
            current_section = sec
        elif not sec and current_section and current_section not in ("Opzioni", "Altre Opzioni"):
            st.markdown("### Altre Opzioni")
            current_section = "Altre Opzioni"
        first_param = False

        ptype = p.get("type")
        key = p.get("key")
        label = p.get("label", key)
        default = p.get("default")
        help_txt = p.get("help")
        widget_key = f"param_{_key_safe(tool_uid)}_{_key_safe(str(key))}"

        if ptype == "dynamic_info":
            func_name = p.get("function")
            if tool_module and func_name and hasattr(tool_module, func_name):
                try:
                    current_values = values.copy()
                    func = getattr(tool_module, func_name)
                    info_text = func(current_values)
                    if info_text:
                        st.info(info_text)
                except Exception as e:
                    st.warning(f"Errore calcolo anteprima ({func_name}): {e}")
            else:
                if not tool_module:
                    st.warning(f"Info dinamica '{label}' non disponibile (modulo non caricato)")
                else:
                    st.warning(f"Info dinamica '{label}' non disponibile (funzione '{func_name}' non trovata)")
            continue

        if ptype == "select":
            options = p.get("options", [])
            index = options.index(default) if default in options else 0
            values[key] = st.selectbox(label, options, index=index, help=help_txt, key=widget_key)
        elif ptype == "radio":
            options = p.get("options", [])
            index = options.index(default) if default in options else 0
            values[key] = st.radio(label, options, index=index, help=help_txt, key=widget_key)
        elif ptype == "multiselect":
            options = p.get("options", [])
            default_val = p.get("default", []) if isinstance(p.get("default"), list) else []
            values[key] = st.multiselect(label, options, default=default_val, help=help_txt, key=widget_key)
        elif ptype == "checkbox":
            values[key] = st.checkbox(label, value=bool(default) if default is not None else False, help=help_txt, key=widget_key)
        elif ptype == "number":
            min_v = p.get("min", 0)
            max_v = p.get("max", 10**9)
            step = p.get("step", 1)
            default_val = default if default is not None else min_v
            if widget_key not in st.session_state:
                st.session_state[widget_key] = default_val
            val_raw = st.session_state.get(widget_key, default_val)
            wants_float = any(
                isinstance(x, float) or (isinstance(x, str) and ("." in x or "," in x))
                for x in (min_v, max_v, step, val_raw)
            )
            def _as_float(x, fb):
                try:
                    return float(str(x).replace(",", ".")) if x is not None else fb
                except Exception:
                    return fb
            if wants_float:
                min_v = _as_float(min_v, 0.0)
                max_v = _as_float(max_v, 10**9)
                step = _as_float(step, 1.0)
                val = _as_float(val_raw, min_v)
            else:
                min_v = int(min_v)
                max_v = int(max_v)
                step = int(step)
                try:
                    val = int(val_raw)
                except Exception:
                    val = min_v
            st.session_state[widget_key] = val
            # Formato custom per alcuni parametri economici
            fmt = "%.6f" if wants_float else "%d"
            if wants_float:
                if key in {"pension_min", "aliquota_1", "aliquota_2", "aliquota_3"}:
                    fmt = "%.2f"
                elif key in {"coeff_maggiorazione"}:
                    fmt = "%.3f"

            values[key] = st.number_input(
                label,
                min_value=min_v,
                max_value=max_v,
                value=val,
                step=step,
                format=fmt,
                help=help_txt,
                key=widget_key,
            )
        elif ptype == "text":
            values[key] = st.text_input(label, value=str(default) if default is not None else "", help=help_txt, key=widget_key)
        elif ptype == "textarea":
            values[key] = st.text_area(label, value=str(default) if default is not None else "", help=help_txt, key=widget_key)
        elif ptype == "folder":
            import tkinter as tk
            from tkinter import filedialog
            
            # Callback per gestire il picker senza causare StreamlitAPIException
            def _pick_folder_cb(k):
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                path = filedialog.askdirectory()
                root.destroy()
                if path:
                    curr = st.session_state.get(k, "").strip()
                    if curr:
                        # Se è già presente, aggiunge una nuova riga (multi-folder)
                        if path not in curr.splitlines():
                            st.session_state[k] = f"{curr}\n{path}"
                    else:
                        st.session_state[k] = path

            c1, c2 = st.columns([0.85, 0.15])
            with c2:
                st.write(" ")
                # L'uso di on_click garantisce che lo stato venga aggiornato PRIMA del rendering dei widget
                st.button("📂", key=f"btn_{widget_key}", help="Aggiungi una cartella locale", on_click=_pick_folder_cb, args=(widget_key,))
            
            with c1:
                # Usiamo text_area per supportare percorsi multipli in modo leggibile
                values[key] = st.text_area(
                    label, 
                    value=st.session_state.get(widget_key, str(default) if default is not None else ""), 
                    help=help_txt, 
                    key=widget_key,
                    height=90
                )
        elif ptype == "file_path_info":
            # Parametro speciale che mostra un path con pulsanti per Aprirlo o Selezionarne uno nuovo
            import tkinter as tk
            from tkinter import filedialog
            
            # Valore attuale
            current_val = st.session_state.get(widget_key, str(default) if default is not None else "")
            
            def _pick_file_path(k, base_dir_hint, p_key, t_mod):
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                f_path = filedialog.askopenfilename(initialdir=base_dir_hint)
                root.destroy()
                if f_path:
                    # 1. Aggiorna stato sessione (immediato)
                    st.session_state[k] = f_path
                    
                    # 2. Persistenza su disco (per prossimo riavvio)
                    if t_mod and hasattr(t_mod, "__file__"):
                        try:
                            py_p = Path(t_mod.__file__)
                            if py_p.exists() and hasattr(t_mod, "TOOL"):
                                # Facciamo una copia profonda di TOOL per non sporcare il modulo in memoria
                                import copy
                                new_tool = copy.deepcopy(getattr(t_mod, "TOOL"))
                                
                                # Cerchiamo il parametro giusto negli 'inputs' o 'params'
                                found = False
                                for collection in ["inputs", "params"]:
                                    if collection in new_tool:
                                        for p in new_tool[collection]:
                                            if p.get("key") == p_key:
                                                p["default"] = f_path
                                                found = True
                                                break
                                    if found: break
                                
                                if found:
                                    # Usiamo l'helper esistente per scrivere su file
                                    update_tool_fields_in_file(py_p, {"params": new_tool.get("params", []), "inputs": new_tool.get("inputs", [])})
                        except Exception as e:
                            # Silenzioso in UI ma logghiamo se possibile
                            print(f"Errore salvataggio persistente: {e}")

            def _open_folder_of_file(f_path):
                if f_path:
                    f_p = Path(f_path)
                    if not f_p.is_absolute(): f_p = BASE_DIR / f_p
                    if f_p.exists():
                        subprocess.run(['explorer', '/select,', str(f_p.absolute())])
                    elif f_p.parent.exists():
                        subprocess.run(['explorer', str(f_p.parent.absolute())])

            # Layout compatto: Input largo + 2 bottoni piccoli
            c1, c2, c3 = st.columns([0.84, 0.08, 0.08], gap="small")
            
            with c1:
                values[key] = st.text_input(label, value=current_val, help=help_txt, key=widget_key)
            
            with c2:
                # Spacer per allineare i bottoni all'altezza dell'input (label inclusa)
                st.markdown("<div style='padding-top: 1.8rem;'></div>", unsafe_allow_html=True)
                st.button("📂", key=f"btn_pick_{widget_key}", help="Sfoglia file...", 
                          on_click=_pick_file_path, args=(widget_key, str(BASE_DIR), key, tool_module), width="stretch")
            
            with c3:
                st.markdown("<div style='padding-top: 1.8rem;'></div>", unsafe_allow_html=True)
                st.button("🔍", key=f"btn_open_{widget_key}", help="Apri in Explorer", 
                          on_click=_open_folder_of_file, args=(values[key],), width="stretch")

        else:
            st.error(f"Tipo parametro non gestito: {ptype}")

    return values


def _editor_panel(tool: Dict[str, Any]) -> None:
    uid = str(tool.get("uid"))
    src = Path(str(tool.get("source_path") or ""))
    with st.expander("⚙️ Impostazioni Tool", expanded=True):
        st.caption("Modifica direttamente il dict TOOL nel file .py. Se TOOL non è un dict literal parsabile, non verrà modificato.")
        name_val = st.text_input("Name (TOOL['name'])", value=str(tool.get("name", "")), key=f"ed_name_{_key_safe(uid)}")
        desc_val = st.text_area("Description (TOOL['description'])", value=str(tool.get("description", "")), height=110, key=f"ed_desc_{_key_safe(uid)}")
        inputs_json = json.dumps(tool.get("inputs", []) or [], ensure_ascii=False, indent=2)
        params_json = json.dumps(tool.get("params", []) or [], ensure_ascii=False, indent=2)
        st.markdown("**Inputs (JSON)**")
        inputs_text = st.text_area(
            "TOOL['inputs']",
            value=inputs_json,
            height=160,
            key=f"ed_inputs_{_key_safe(uid)}",
            help="Lista di dict. types: txt_multi | txt_single | xlsx_single",
        )
        st.markdown("**Params (JSON)**")
        params_text = st.text_area(
            "TOOL['params']",
            value=params_json,
            height=200,
            key=f"ed_params_{_key_safe(uid)}",
            help="Lista di dict. types: select | radio | checkbox | number | text | textarea",
        )
        b1, b2, _ = st.columns([0.20, 0.20, 0.60], vertical_alignment="center")
        if b1.button("Salva", key=f"ed_save_{_key_safe(uid)}", type="primary", width="stretch"):
            try:
                new_inputs = json.loads(inputs_text) if inputs_text.strip() else []
            except Exception as e:
                st.error(f"inputs JSON non valido: {e}")
                return
            try:
                new_params = json.loads(params_text) if params_text.strip() else []
            except Exception as e:
                st.error(f"params JSON non valido: {e}")
                return
            updates = {"name": str(name_val), "description": str(desc_val), "inputs": new_inputs, "params": new_params}
            ok, msg = update_tool_fields_in_file(src, updates)
            if ok:
                st.success(msg)
                st.session_state.pop("edit_mode_uid", None)
                load_tools_cached.clear()
                st.rerun()
            else:
                st.error(msg)
        if b2.button("Chiudi", key=f"ed_close_{_key_safe(uid)}", width="stretch"):
            st.session_state.pop("edit_mode_uid", None)
            st.rerun()
        st.write("")
        st.info("Qui puoi cambiare nome/descrizione e anche inputs/params. Se non vuoi toccare qualcosa, lascialo com'è.")

def render_tool(tool: Dict[str, Any]) -> None:
    tool_uid = str(tool.get("uid"))
    tool_title = str(tool.get("name", tool.get("id", "Tool")))

    st.markdown(
        """<style>
/* Header tool: colonne si restringono al contenuto → bottone ⚙️ incollato al titolo */
/* :not(:has(stVerticalBlockBordered)) esclude il blocco 70/30 esterno che contiene st.container(border=True) */
div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) {
    gap: 6px !important;
    align-items: center !important;
}
div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) > div[data-testid="stColumn"] {
    flex: 0 0 auto !important;
    width: fit-content !important;
    max-width: 92% !important;
}
div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) > div[data-testid="stColumn"]:last-child {
    margin-top: 18px !important;
}
div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) div[data-testid="stTooltipHoverTarget"] {
    justify-content: flex-start !important;
    width: auto !important;
}
</style>""",
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns([0.95, 0.05])
    with c1:
        # Titolo su riga singola (no wrap)
        st.markdown(f'<h1 style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 0;">{tool_title}</h1>', unsafe_allow_html=True)
    with c2:
        if st.button("⚙️", key=f"gear_{_key_safe(tool_uid)}", help="Impostazioni tool (modifica TOOL)", type="primary"):
            st.session_state["edit_mode_uid"] = tool_uid



    if st.session_state.get("edit_mode_uid") == tool_uid:
        _editor_panel(tool)
        st.divider()

    desc = (tool.get("description") or "").strip()
    if desc:
        with st.expander("📖 Consulta la Guida e Logica del Tool", expanded=False):
            st.markdown(desc, unsafe_allow_html=True)
    if tool.get("import_error") or tool.get("runner") is None:
        st.error("Questo tool non è eseguibile (errore import o runner mancante).")
        return

    inputs: List[Dict[str, Any]] = tool.get("inputs", []) or []
    uploads: Dict[str, Any] = {}

    # --- UI TOP RENDER (CUSTOM) ---
    tool_mod = tool.get("module_obj")
    if tool_mod and hasattr(tool_mod, "get_ui_top"):
        try:
            tool_mod.get_ui_top()
        except Exception as e:
            st.warning(f"Errore UI Top: {e}")
    # ------------------------------

    if tool.get("id") != "attivazione_profili":
        st.markdown("### Input file")
        if not inputs:
            st.caption("— Nessun input file richiesto —")

    for inp in inputs:
        itype = inp["type"]
        key = inp["key"]
        label = inp.get("label", key)
        required = bool(inp.get("required", False))
        widget_key = f"up_{_key_safe(tool_uid)}_{_key_safe(str(key))}"

        if key == "file_banca_dati":
            st.markdown("### Input da File Banca Dati (XLSX/SLK)")

        auto_upload = None
        auto_hint = ""
        if itype == "xlsx_single" and key != "file_banca_dati":
            # RIPRISTINO MIRATO: Precaricamento SOLO per Emilia Romagna (sindrinn_normalizer)
            # L'utente vuole che questo specifico tool parta con il file precaricato.
            if tool.get("id") == "sindrinn_normalizer":
                tool_dir = Path(tool.get("source_path", "")).resolve().parent
                candidates = [
                    tool_dir / "File" / "Prospetto_Emilia_Romagna.xlsx",
                    tool_dir / "Prospetto_Emilia_Romagna.xlsx",
                    BASE_DIR / "tools" / "Emilia-Romagna" / "File" / "Prospetto_Emilia_Romagna.xlsx",
                    BASE_DIR / "tools" / "Emilia-Romagna" / "Prospetto_Emilia_Romagna.xlsx",
                ]
                auto_path = next((p for p in candidates if p.exists() and p.is_file()), None)
                if auto_path:
                    class _DummyUpload(io.BytesIO):
                        def __init__(self, p: Path):
                            data = p.read_bytes()
                            super().__init__(data)
                            self.name = p.name
                        def getbuffer(self):
                            return super().getbuffer()
                    auto_upload = _DummyUpload(auto_path)
                    auto_hint = f"  ✅ **(precaricato: {auto_path.name})**"

        current_upload = st.session_state.get(widget_key, auto_upload)
        count_str = ""
        is_auto = auto_upload is not None and current_upload is auto_upload
        if itype.endswith("_multi"):
            count = len(current_upload) if isinstance(current_upload, list) else 0
            if count > 0:
                count_str = f"  ✅ **({count} file caricati)**"
        elif current_upload is not None and not is_auto:
            count_str = "  ✅ **(1 file caricato)**"
        elif auto_upload is not None:
            count_str = auto_hint

        final_label = f"{label}{count_str}"

        if itype == "txt_multi":
            uploads[key] = st.file_uploader(final_label, type=["txt"], accept_multiple_files=True, key=widget_key)
            if required and not uploads[key]:
                st.caption("⚠️ Campo richiesto")
        elif itype == "file_multi":
            uploads[key] = st.file_uploader(final_label, accept_multiple_files=True, key=widget_key)
            if required and not uploads[key]:
                st.caption("⚠️ Campo richiesto")
        elif itype == "txt_single":
            uploads[key] = st.file_uploader(final_label, type=["txt"], accept_multiple_files=False, key=widget_key)
            if required and uploads[key] is None:
                st.caption("⚠️ Campo richiesto")
        elif itype == "file_single":
            uploads[key] = st.file_uploader(final_label, accept_multiple_files=False, key=widget_key)
            if required and uploads[key] is None:
                st.caption("⚠️ Campo richiesto")
        elif itype == "xlsx_single":
            uploads[key] = st.file_uploader(
                final_label, type=["xlsx", "xls", "slk"], accept_multiple_files=False, key=widget_key
            )
            if uploads[key] is None and auto_upload is not None:
                uploads[key] = auto_upload
            if required and uploads[key] is None:
                st.caption("⚠️ Campo richiesto")
        elif itype == "warning":
            st.warning(label)
        elif itype == "info":
            st.info(label)
        elif itype == "error":
            st.error(label)
        elif itype == "success":
            st.success(label)
        elif itype == "markdown":
            st.markdown(label)
        else:
            st.error(f"Tipo input non gestito: {itype}")

        if inp.get("note"):
             st.warning(inp["note"])

    param_values: Dict[str, Any] = {}
    if tool.get("dynamic_params"):
        # Usiamo un contenitore stabile per evitare che i widget spariscano/si chiudano al variare dei parametri
        dyn_container = st.container()
        try:
            prefix = f"param_{_key_safe(tool_uid)}_"
            # Recupero parametri esistenti per permettere a get_dynamic_params di 'reagire'
            all_possible_params = (tool.get("params", []) or [])
            
            for k, v in list(st.session_state.items()):
                if k.startswith(prefix):
                    p_key_safe = k[len(prefix):]
                    param_found = False
                    for sp in all_possible_params:
                        if _key_safe(str(sp.get("key"))) == p_key_safe:
                            param_values[sp.get("key")] = v
                            param_found = True
                            break
                    if not param_found:
                        param_values[p_key_safe] = v
            
            dyn_params_def = tool["dynamic_params"](uploads, param_values)
            if dyn_params_def:
                with dyn_container:
                    st.markdown("### ⚡ Opzioni dinamiche")
                    tool_mod = tool.get("module_obj")
                    dyn_values = render_params_list(dyn_params_def, tool_uid, tool_module=tool_mod)
                param_values.update(dyn_values)
        except Exception as e:
            st.warning(f"⚠️ Impossibile caricare opzioni dinamiche: {e}")

    static_params = tool.get("params", []) or []
    tool_mod = tool.get("module_obj")
    static_values = render_params_list(static_params, tool_uid, tool_module=tool_mod)
    param_values.update(static_values)

    st.divider()

    results_key = f"results_{_key_safe(tool_uid)}"
    run_clicked = False
    if tool.get("id") != "attivazione_profili":
        run_clicked = st.button("▶️ Esegui", width="stretch", key=f"run_{_key_safe(tool_uid)}")
    
    if not run_clicked and results_key not in st.session_state:
        return

    for inp in inputs:
        if inp.get("type") == "xlsx_single" and inp.get("key") != "file_banca_dati":
            # RIPRISTINO FALLBACK MIRATO SOLO PER SINDINN NORMALIZER
            if tool.get("id") == "sindrinn_normalizer":
                key = inp["key"]
                if uploads.get(key) is None:
                    tool_dir = Path(tool.get("source_path", "")).resolve().parent
                    candidates = [
                        tool_dir / "File" / "Prospetto_Emilia_Romagna.xlsx",
                        tool_dir / "Prospetto_Emilia_Romagna.xlsx",
                        BASE_DIR / "tools" / "Emilia-Romagna" / "File" / "Prospetto_Emilia_Romagna.xlsx",
                        BASE_DIR / "tools" / "Emilia-Romagna" / "Prospetto_Emilia_Romagna.xlsx",
                    ]
                    auto_path = next((p for p in candidates if p.exists() and p.is_file()), None)
                    if auto_path:
                        class _DummyUpload:
                            def __init__(self, p: Path):
                                self.name = p.name
                                self._data = p.read_bytes()
                            def getbuffer(self):
                                return memoryview(self._data)
                        uploads[key] = _DummyUpload(auto_path)
                        st.caption(f"✅ File precaricato: {auto_path}")

    for inp in inputs:
        if inp.get("required"):
            v = uploads.get(inp["key"])
            if v is None or v == []:
                st.error(f"Manca input richiesto: {inp.get('label', inp['key'])}")
                return

    all_params_def = static_params + (dyn_params_def if tool.get("dynamic_params") and 'dyn_params_def' in locals() else [])
    for p in all_params_def:
        if p.get("required"):
            val = param_values.get(p.get("key"))
            if val is None or (isinstance(val, str) and not val.strip()):
                st.error(f"Manca opzione richiesta: {p.get('label', p.get('key'))}")
                return

    exact = tool.get("exact_txt")
    if exact is not None:
        multi_keys = [i["key"] for i in inputs if i["type"] == "txt_multi"]
        if multi_keys:
            mk = multi_keys[0]
            files = uploads.get(mk) or []
            if not isinstance(files, list) or len(files) != int(exact):
                st.error(f"Devi caricare esattamente {exact} file TXT (ne hai caricati {len(files) if isinstance(files, list) else 0}).")
                return

    if run_clicked:
        with st.spinner("Esecuzione in corso…"):
            with tempfile.TemporaryDirectory() as tmp:
                tmpdir = Path(tmp)
                out_dir = tmpdir / "out"
                out_dir.mkdir(parents=True, exist_ok=True)
                saved_inputs: Dict[str, Any] = {}
                for inp in inputs:
                    key = inp["key"]
                    itype = inp["type"]
                    v = uploads.get(key)
                    if itype in ("txt_multi", "file_multi"):
                        paths: List[Path] = []
                        for f in (v or []):
                            paths.append(save_upload_to_tmp(tmpdir, f))
                        saved_inputs[key] = paths
                    elif itype in ("txt_single", "xlsx_single", "file_single"):
                        saved_inputs[key] = None if v is None else save_upload_to_tmp(tmpdir, v)
                try:
                    out_files = tool["runner"](**saved_inputs, **param_values, out_dir=out_dir)
                    files_data = []
                    for pth in out_files:
                        pth = Path(pth)
                        files_data.append({"name": pth.name, "data": pth.read_bytes()})
                    st.session_state[results_key] = files_data
                except TypeError as te:
                    st.error("Firma della funzione run(...) non compatibile con i parametri/inputs dichiarati.")
                    st.exception(te)
                except Exception as e:
                    st.exception(e)

    if results_key in st.session_state:
        res = st.session_state[results_key]
        if not res:
            st.warning("Il tool non ha prodotto file.")
        else:
            st.success("Fatto! Risultati pronti.")

            # --- UI POST RESULTS (CUSTOM HOOK by Assistant) ---
            # Permette ai tool di mostrare dashboard o report persistenti PRIMA dei bottoni di download
            if tool_mod and hasattr(tool_mod, "get_ui_results"):
                try:
                    tool_mod.get_ui_results()
                except Exception as e:
                    st.warning(f"Errore UI Results: {e}")
            # --------------------------------------------------

            if len(res) > 1:
                dl_mode = st.radio(
                    "Scegli come scaricare:",
                    ["📦 Scarica tutto (.zip)", "📄 Scarica singoli file"],
                    horizontal=True,
                    key=f"dl_mode_{results_key}"
                )
                if dl_mode == "📦 Scarica tutto (.zip)":
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for item in res:
                            zf.writestr(item['name'], item['data'])
                    st.download_button(
                        label="📦 Scarica archivio ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="risultati_completi.zip",
                        mime="application/zip",
                        width="stretch",
                        type="primary"
                    )
                else:
                    for item in res:
                        st.download_button(
                            label=f"⬇️ Scarica {item['name']}",
                            data=item['data'],
                            file_name=item['name'],
                            width="stretch"
                        )
            else:
                for item in res:
                    st.download_button(
                        label=f"⬇️ Scarica {item['name']}",
                        data=item['data'],
                        file_name=item['name'],
                        width="stretch",
                        type="primary"
                    )

            validation_key = "emilia_validation_results"
            if validation_key in st.session_state:
                val_data = st.session_state[validation_key]
                st.markdown("---")
                st.subheader("📊 Risultati Validazione vs File Originale")
                if "html" in val_data:
                    st.markdown(val_data["html"], unsafe_allow_html=True)
                    if "results" in val_data:
                        for result in val_data.get("results", []):
                            if result["details"] and result["count"] > 0:
                                label_sheet = f"🔍 Dettagli {result['sheet']} ({result['count']} differenze)"
                                with st.expander(label_sheet, expanded=False):
                                    for i, diff in enumerate(result['details'][:100], 1):
                                        st.text(f"{i}. {diff}")
                                    if len(result['details']) > 100:
                                        st.caption(f"... e altre {len(result['details']) - 100} differenze")
                elif "error" in val_data:
                    st.error(f"Errore durante la validazione: {val_data['error']}")
                st.markdown("---")


# ------------------------------------------------------------
# Page Config & Initial Theme
# ------------------------------------------------------------
# Inizializzazione Stato Assistente
if "show_assistant" not in st.session_state:
    st.session_state["show_assistant"] = False


# Funzioni per la persistenza del tema
CONFIG_FILE = BASE_DIR / "theme_config.json"
AI_CONFIG_FILE = BASE_DIR / "ai_config.json"

def load_theme_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            pass
    return {"sidebar_lightness": 33, "region_order": []}

def save_theme_config(config):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)
    except:
        pass

def load_ai_config():
    if AI_CONFIG_FILE.exists():
        try:
            with open(AI_CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            pass
    return {
        "model_id": "arcee-ai/trinity-large-preview:free",
        "base_url": "https://openrouter.ai/api/v1"
    }

def save_ai_config(config):
    try:
        with open(AI_CONFIG_FILE, "w") as f:
            json.dump(config, f)
    except:
        pass

if "sidebar_lightness" not in st.session_state:
    config = load_theme_config()
    st.session_state["sidebar_lightness"] = config.get("sidebar_lightness", 33)
    st.session_state["region_order"] = config.get("region_order", [])

if "assistant_model_id" not in st.session_state or "assistant_base_url" not in st.session_state:
    ai_cfg = load_ai_config()
    st.session_state.assistant_model_id = ai_cfg.get("model_id", "arcee-ai/trinity-large-preview:free")
    st.session_state.assistant_base_url = ai_cfg.get("base_url", "https://openrouter.ai/api/v1")

# Iniezione CSS Globale CNA (Dinamica)
l_val = st.session_state["sidebar_lightness"]
st.markdown(f"""
    <style>
    /* TEMA CNA - BLU ISTITUZIONALE PROFESSIONALE */
    [data-testid="stAppViewContainer"], .main {{ background-color: #f0f4f8 !important; color: #003366 !important; }}
    [data-testid="stHeader"] {{ background-color: rgba(240, 244, 248, 0.8) !important; }}
    [data-testid="stSidebar"] {{ 
        background-color: hsl(210, 100%, {l_val}%) !important; 
        border-right: 1px solid hsl(210, 100%, {max(0, l_val-10)}%) !important; 
    }}
    /* Sidebar Text - Bianco su Blu CNA */
    [data-testid="stSidebar"] * {{ color: #ffffff !important; }}
    
    /* SIDEBAR TEXT & EXPANDERS - FIX TOTALE COLORE BIANCO E BORDI FINE */
    [data-testid="stSidebar"] [data-testid="stExpander"] summary,
    [data-testid="stSidebar"] [data-testid="stExpander"] summary *,
    [data-testid="stSidebar"] details summary, 
    [data-testid="stSidebar"] details summary * {{
        color: #ffffff !important;
        fill: #ffffff !important;
        text-decoration: none !important;
    }}

    /* Bordi Bianchi Fini per Expanders (Regioni) - Fix Totale */
    [data-testid="stSidebar"] [data-testid="stExpander"] {{
        border: 1px solid rgba(255, 255, 255, 0.4) !important;
        background-color: transparent !important;
        border-radius: 8px !important;
        margin-bottom: 10px !important;
        padding: 0 !important;
    }}

    /* Blocca lo sfondo trasparente su OGNI stato interno (Aperto, Chiuso, Focus) */
    [data-testid="stSidebar"] [data-testid="stExpander"] details,
    [data-testid="stSidebar"] [data-testid="stExpander"] details > div,
    [data-testid="stSidebar"] [data-testid="stExpander"] details[open],
    [data-testid="stSidebar"] [data-testid="stExpander"] details[open] > div,
    [data-testid="stSidebar"] [data-testid="stExpander"] summary {{
        border: none !important;
        background-color: transparent !important;
        background: transparent !important;
        box-shadow: none !important;
    }}

    /* Hover leggero sul titolo per feedback visivo */
    [data-testid="stSidebar"] [data-testid="stExpander"] summary:hover {{
        background-color: rgba(255, 255, 255, 0.1) !important;
    }}

    /* Bordi Bianchi Fini per i Pulsanti dei Tool */
    [data-testid="stSidebar"] div[data-testid="stButton"] button {{
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        background-color: rgba(255, 255, 255, 0.1) !important;
        color: #ffffff !important;
        border-radius: 4px !important;
        margin: 2px 0 !important;
        width: 100% !important;
        white-space: normal !important;
        word-break: break-word !important;
        height: auto !important;
        min-height: 38px !important;
        padding: 6px 10px !important;
        line-height: 1.35 !important;
    }}
    [data-testid="stSidebar"] div[data-testid="stButton"] button:hover {{
        background-color: rgba(255, 255, 255, 0.2) !important;
        border-color: #ffffff !important;
    }}

    /* Bordi Bianchi Fini per gli Input (Cerca, Root) - Sfondo Bianco e Testo Grigio */
    [data-testid="stSidebar"] div[data-testid="stTextInput"] input {{
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
        background-color: #ffffff !important; /* Sfondo Bianco */
        color: #555555 !important; /* Testo Grigio */
        border-radius: 4px !important;
        padding: 8px !important;
    }}
    
    /* Forza il colore del testo per la leggibilità */
    [data-testid="stSidebar"] input {{
        color: #555555 !important;
        -webkit-text-fill-color: #555555 !important;
    }}

    /* Stile Matita (Edit) nella Sidebar */
    [data-testid="stSidebar"] [data-testid="column"]:nth-child(2) button {{
        background-color: #ffffff !important;
        color: #0054a6 !important;
        border: 1px solid #ffffff !important;
        border-radius: 4px !important;
        padding: 0 !important;
        margin: 2px 0 !important;
        min-height: 38px !important;
        height: 38px !important;
        width: 100% !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 1.1rem !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.2) !important;
    }}
    [data-testid="stSidebar"] [data-testid="column"]:nth-child(2) button:hover {{
        background-color: #f0f7ff !important;
        border-color: #ffffff !important;
        transform: scale(1.05);
    }}

    /* Card e Contenitori - Bordi CNA Dinamici */
    div[data-testid="stVerticalBlockBordered"] {{ 
        background-color: #ffffff !important; 
        border: 1px solid #c0d1e2 !important; 
        border-left: 6px solid hsl(210, 100%, {l_val}%) !important;
        box-shadow: 0 4px 15px rgba(0,51,102,0.12) !important;
    }}
    
    /* Titles e Testi Dinamici */
    h1, h2, h3, h4 {{ color: hsl(210, 100%, {l_val}%) !important; font-weight: bold !important; }}
    .stMarkdown, .stCaption, p, span, label {{ color: hsl(210, 100%, {max(0, l_val-15)}%) !important; }}
    
    /* Bottoni reali (st.button) — NON usa il selettore generico "button"
       per evitare di colorare anche il ? del tooltip (che Streamlit rende
       come <button> ma NON è dentro [data-testid="stButton"]) */
    [data-testid="stButton"] button {{
        background-color: hsl(210, 100%, {l_val}%) !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15) !important;
        font-weight: 500 !important;
        height: 38px !important;
        min-height: 38px !important;
        padding: 0 12px !important;
        line-height: 38px !important;
        border-radius: 6px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 0.95rem !important;
    }}

    /* Testo interno ai bottoni reali */
    [data-testid="stButton"] button p,
    [data-testid="stButton"] button span {{
        background-color: transparent !important;
        background: transparent !important;
        color: #ffffff !important;
        margin: 0 !important;
        padding: 0 !important;
    }}

    [data-testid="stButton"] button:hover {{
        background-color: hsl(210, 100%, {max(0, l_val-10)}%) !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important;
    }}
    [data-testid="stButton"] button:hover * {{
        color: #ffffff !important;
    }}
    
    /* FIX SOVRAPPOSIZIONE E BORDI INPUT DINAMICI */
    div[data-baseweb="input"], 
    div[data-baseweb="select"] > div,
    div[data-testid="stTextInput"] > div,
    div[data-testid="stTextArea"] > div,
    div[data-testid="stNumberInput"] > div {{ 
        border: 1px solid hsl(210, 100%, {l_val}%) !important; 
        background-color: #ffffff !important;
        border-radius: 4px !important;
        overflow: hidden !important;
    }}

    div[data-testid="stTextInput"] input, 
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stNumberInput"] input,
    div[data-baseweb="select"] span {{ 
        border: none !important;
        box-shadow: none !important;
        background-color: transparent !important;
        color: #555555 !important; /* Testo Grigio scuro per contrasto su bianco */
        outline: none !important;
    }}

    /* FIX BOTTONI +/- (Number Input) Dinamici */
    div[data-testid="stNumberInput"] button {{
        background-color: hsl(210, 100%, {l_val}%) !important;
        color: #ffffff !important;
        border: none !important;
        border-left: 1px solid rgba(255, 255, 255, 0.3) !important;
        margin: 0 !important;
        height: 100% !important;
        border-radius: 0 !important;
    }}
    
    div[data-testid="stNumberInput"] button:hover {{
        background-color: hsl(210, 100%, {max(0, l_val-10)}%) !important;
    }}

    /* Bottoni azione (📂, 🔍, ✏️) nei blocchi orizzontali:
       larghezza 100% della colonna, emoji ben visibile,
       allineati verticalmente all'input (margin-top compensa la label sopra l'input) */
    [data-testid="stHorizontalBlock"] [data-testid="stButton"] button {{
        width: 100% !important;
        font-size: 1.2rem !important;
        padding: 0 8px !important;
    }}
    [data-testid="stHorizontalBlock"] [data-testid="stButton"] {{
        margin-top: auto !important;
    }}

    /* Bottone "Sfoglia file..." - stBaseButton-secondary ha specificità maggiore del selettore button globale */
    [data-testid="stBaseButton-secondary"],
    div[data-testid="stFileUploader"] button,
    div[data-testid="stFileUploaderDropzone"] button {{
        background-color: hsl(210, 100%, {l_val}%) !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15) !important;
        font-weight: 500 !important;
    }}
    [data-testid="stBaseButton-secondary"]:hover,
    div[data-testid="stFileUploader"] button:hover {{
        background-color: hsl(210, 100%, {max(0, l_val-10)}%) !important;
    }}
    [data-testid="stBaseButton-secondary"] p,
    [data-testid="stBaseButton-secondary"] span {{
        color: #ffffff !important;
    }}

    /* Drag & Drop (File Uploader) Dinamico */
    div[data-testid="stFileUploader"] {{
        background-color: #f8fafc !important;
        border: 2px dashed hsl(210, 100%, {l_val}%) !important;
        border-radius: 12px !important;
        padding: 15px !important;
        transition: all 0.3s ease;
    }}
    div[data-testid="stFileUploader"]:hover {{
        background-color: #f0f7ff !important;
        border-color: hsl(210, 100%, {max(0, l_val-10)}%) !important;
    }}
    div[data-testid="stFileUploader"] section {{ background-color: transparent !important; }}

    ::placeholder {{
        color: #aaaaaa !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #aaaaaa !important;
    }}
    ::-webkit-input-placeholder {{
        color: #aaaaaa !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #aaaaaa !important;
    }}
    /* Testi descrittivi (label) Dinamici */
    label {{ color: hsl(210, 100%, {max(0, l_val-10)}%) !important; font-weight: 600 !important; }}
    
    /* Footer/Caption */
    .stCaption {{ color: #4a6a8a !important; }}

    /* STILE SPECIALE PER ASSISTENTE OPENROUTER */
    div[data-testid="stSidebar"] [data-testid="stExpander"]:has(input[key="assistant_query"]) {{
        border: 2px solid #FFD700 !important;
        background-color: rgba(255, 215, 0, 0.05) !important;
    }}
    </style>
""", unsafe_allow_html=True)

tools = load_tools_cached()
sidebar_regions(tools)

if "show_assistant" not in st.session_state:
    st.session_state["show_assistant"] = False

if not tools:
    st.warning("Nessun tool trovato.")
else:
    selected = get_selected_tool(tools)
    if selected is None:
        selected = next((t for t in tools if not t.get("import_error")), tools[0])
        set_selected_tool(selected.get("uid", ""))

    # LAYOUT DINAMICO PANNELLO LATERALE (Integrato 70/30) con TWIN BOXES NATIVI
    if st.session_state["show_assistant"]:
        # CSS iniettato QUI (prima dei columns) con selector DOM reali verificati dall'inspector.
        # Target: stColumn > stVerticalBlock (il primo figlio diretto di ogni colonna del 70/30).
        # Il selector usa la catena completa con > per evitare collisioni con i VBlock annidati nel tool.
        st.markdown("""
            <style>
            /* 1. Rimuove il max-width sul contenitore principale → pannelli occupano tutta la larghezza */
            [data-testid="stMainBlockContainer"] {
                max-width: 100% !important;
                padding-left: 1rem !important;
                padding-right: 1rem !important;
            }
            /* 2. Allinea le colonne 70/30 in cima (evita che la colonna corta scivoli in basso) */
            [data-testid="stMainBlockContainer"]
            > [data-testid="stVerticalBlock"]
            > [data-testid="stHorizontalBlock"] {
                align-items: flex-start !important;
            }
            /* 3. Altezza fissa + scroll sui box delle colonne */
            [data-testid="stMainBlockContainer"]
            > [data-testid="stVerticalBlock"]
            > [data-testid="stHorizontalBlock"]
            > [data-testid="stColumn"]
            > [data-testid="stVerticalBlockBordered"] {
                height: calc(100vh - 80px) !important;
                max-height: calc(100vh - 80px) !important;
                overflow-y: auto !important;
                padding: 1rem !important;
                box-sizing: border-box !important;
            }
            </style>
        """, unsafe_allow_html=True)
        col_main, col_assist = st.columns([0.7, 0.3], gap="large")
        with col_main:
            # BOX 1: TOOL (Altezza gestita da CSS)
            with st.container(border=True):
                render_tool(selected)
        with col_assist:
            # BOX 2: ASSISTENTE (Altezza gestita da CSS 86vh)
            with st.container(border=True):
                render_ai_assistant(tools)
    else:
        # JS Trigger per Sidebar Expand (quando l'assistente è chiuso)
        st.markdown("""
            <script>
            setTimeout(() => {
                const doc = window.parent.document;
                const btn = doc.querySelector('button[aria-label="Expand sidebar"]') || doc.querySelector('button[title="Expand sidebar"]') || doc.querySelector('button[aria-label="Apri sidebar"]');
                if (btn) { btn.click(); }
            }, 300);
            </script>
        """, unsafe_allow_html=True)
        render_tool(selected)
