"""
core/styles.py — Design System CSS per Toolbox CNA
====================================================
Centralizza tutto lo stile dell'applicazione fuori da app.py.
La funzione inject_styles(l_val) inietta il CSS globale via st.markdown.

Vantaggi rispetto all'approccio inline precedente:
- Manutenibilità: tutto lo stile in un file
- Design tokens: CSS custom properties per coerenza
- Separazione degli interessi: logica e stile disaccoppiati
- Testabilità: generate_css() restituisce stringa testabile
"""
from __future__ import annotations

import streamlit as st


def generate_css(l_val: int) -> str:
    """
    Genera il CSS completo come stringa.

    Args:
        l_val: Luminosità sidebar (10–60). Determina il colore primario CNA dinamico.

    Returns:
        Stringa CSS completa da iniettare con st.markdown.
    """
    l_dark  = max(0, l_val - 10)
    l_text  = max(0, l_val - 15)
    l_xdark = max(0, l_val - 20)
    l_sub   = max(0, l_val - 6)

    return f"""<style>
/* ================================================================
   TOOLBOX CNA — DESIGN SYSTEM v2
   Generato da core/styles.py
   Luminosità primaria: {l_val}%
================================================================ */

/* ── 1. DESIGN TOKENS (CSS Custom Properties) ───────────────── */
:root {{
    /* Colori primari — derivati da l_val */
    --cna-primary:        hsl(210, 100%, {l_val}%);
    --cna-primary-dark:   hsl(210, 100%, {l_dark}%);
    --cna-primary-text:   hsl(210, 100%, {l_text}%);
    --cna-primary-xdark:  hsl(210, 100%, {l_xdark}%);

    /* Superfici */
    --cna-bg:             #f5f7fa;
    --cna-surface:        #ffffff;
    --cna-border:         #dde3ed;
    --cna-border-subtle:  #ecf0f6;

    /* Testo — gerarchia visiva */
    --cna-text-heading:   hsl(210, 55%, 18%);
    --cna-text-body:      hsl(215, 18%, 34%);
    --cna-text-muted:     hsl(215, 14%, 56%);
    --cna-text-white:     #ffffff;
    --cna-text-input:     #334155;

    /* Semantic */
    --cna-success:        #15803d;
    --cna-warning:        #b45309;
    --cna-error:          #b91c1c;

    /* Spacing scale */
    --sp-1: 4px;   --sp-2: 8px;   --sp-3: 12px;
    --sp-4: 16px;  --sp-5: 20px;  --sp-6: 24px;  --sp-8: 32px;

    /* Radius scale */
    --r-sm: 5px;  --r-md: 8px;  --r-lg: 12px;  --r-xl: 16px;

    /* Shadow scale */
    --sh-xs: 0 1px 2px rgba(0,40,90,.06);
    --sh-sm: 0 2px 6px rgba(0,40,90,.09);
    --sh-md: 0 4px 14px rgba(0,40,90,.12);
    --sh-lg: 0 8px 28px rgba(0,40,90,.16);

    /* Font */
    --font:      -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                 "Helvetica Neue", Arial, sans-serif;
    --font-mono: "Cascadia Code", "Fira Code", Consolas, monospace;

    /* Transizioni */
    --t-fast: 130ms cubic-bezier(.4, 0, .2, 1);
    --t-base: 200ms cubic-bezier(.4, 0, .2, 1);
}}

/* ── 2. LAYOUT BASE ─────────────────────────────────────────── */
[data-testid="stAppViewContainer"], .main {{
    background-color: var(--cna-bg) !important;
    color: var(--cna-text-body) !important;
    font-family: var(--font) !important;
}}

[data-testid="stHeader"] {{
    background-color: rgba(245, 247, 250, .93) !important;
    backdrop-filter: blur(10px) !important;
    border-bottom: 1px solid var(--cna-border-subtle) !important;
}}

/* Larghezza massima rimossa: layout full-width (desktop-first) */
[data-testid="stMainBlockContainer"] {{
    max-width: 100% !important;
    padding-left:  1.25rem !important;
    padding-right: 1.25rem !important;
    padding-top:   0.75rem !important;
}}

/* ── 3. SIDEBAR ─────────────────────────────────────────────── */
[data-testid="stSidebar"] {{
    background: linear-gradient(
        160deg,
        hsl(210, 100%, {l_val}%) 0%,
        hsl(210, 100%, {l_sub}%) 100%
    ) !important;
    border-right: 1px solid hsl(210, 100%, {l_dark}%) !important;
}}

/* Tutto il testo sidebar: bianco */
[data-testid="stSidebar"] * {{
    color: var(--cna-text-white) !important;
}}

/* Titolo sidebar */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2 {{
    font-size: 1rem !important;
    font-weight: 700 !important;
    letter-spacing: .06em !important;
    text-transform: uppercase !important;
    opacity: .85 !important;
}}

/* Expanders regioni */
[data-testid="stSidebar"] [data-testid="stExpander"] {{
    border: 1px solid rgba(255, 255, 255, .22) !important;
    background-color: rgba(255, 255, 255, .06) !important;
    border-radius: var(--r-md) !important;
    margin-bottom: var(--sp-2) !important;
    padding: 0 !important;
    transition: background-color var(--t-fast) !important;
}}

[data-testid="stSidebar"] [data-testid="stExpander"]:hover {{
    background-color: rgba(255, 255, 255, .11) !important;
}}

/* Reset stati interni expander sidebar */
[data-testid="stSidebar"] [data-testid="stExpander"] details,
[data-testid="stSidebar"] [data-testid="stExpander"] details > div,
[data-testid="stSidebar"] [data-testid="stExpander"] details[open],
[data-testid="stSidebar"] [data-testid="stExpander"] details[open] > div,
[data-testid="stSidebar"] [data-testid="stExpander"] summary {{
    border: none !important;
    background: transparent !important;
    box-shadow: none !important;
}}

[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] [data-testid="stExpander"] summary * {{
    color: var(--cna-text-white) !important;
    fill:  var(--cna-text-white) !important;
    text-decoration: none !important;
}}

[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover {{
    background-color: rgba(255, 255, 255, .08) !important;
}}

/* Pulsanti tool in sidebar */
[data-testid="stSidebar"] div[data-testid="stButton"] button {{
    border: 1px solid rgba(255, 255, 255, .18) !important;
    background-color: rgba(255, 255, 255, .08) !important;
    color: var(--cna-text-white) !important;
    border-radius: var(--r-sm) !important;
    margin: 2px 0 !important;
    width: 100% !important;
    white-space: normal !important;
    word-break: break-word !important;
    height: auto !important;
    min-height: 36px !important;
    padding: 5px 10px !important;
    line-height: 1.3 !important;
    font-size: .875rem !important;
    font-weight: 500 !important;
    text-align: left !important;
    transition: background-color var(--t-fast), border-color var(--t-fast) !important;
}}

[data-testid="stSidebar"] div[data-testid="stButton"] button:hover {{
    background-color: rgba(255, 255, 255, .18) !important;
    border-color:     rgba(255, 255, 255, .45) !important;
}}

[data-testid="stSidebar"] div[data-testid="stButton"] button:focus-visible {{
    outline: 2px solid rgba(255, 255, 255, .75) !important;
    outline-offset: 1px !important;
}}

/* Input testo sidebar (Cerca, Root) */
[data-testid="stSidebar"] div[data-testid="stTextInput"] input {{
    border: 1px solid rgba(255, 255, 255, .35) !important;
    background-color: rgba(255, 255, 255, .95) !important;
    color: var(--cna-text-input) !important;
    border-radius: var(--r-sm) !important;
    padding: var(--sp-2) var(--sp-3) !important;
    font-size: .875rem !important;
}}

[data-testid="stSidebar"] input {{
    color: var(--cna-text-input) !important;
    -webkit-text-fill-color: var(--cna-text-input) !important;
}}

[data-testid="stSidebar"] div[data-testid="stTextInput"] input:focus {{
    border-color: rgba(255, 255, 255, .75) !important;
    background-color: #ffffff !important;
    outline: none !important;
}}

/* Override wrapper input sidebar (evita bordi blu del main) */
[data-testid="stSidebar"] div[data-baseweb="input"],
[data-testid="stSidebar"] div[data-testid="stTextInput"] > div {{
    border: 1px solid rgba(255, 255, 255, .35) !important;
    background-color: rgba(255, 255, 255, .92) !important;
    border-radius: var(--r-sm) !important;
}}

/* Bottone matita ✏️ in sidebar */
[data-testid="stSidebar"] [data-testid="column"]:nth-child(2) button {{
    background-color: rgba(255, 255, 255, .88) !important;
    color: var(--cna-primary-xdark) !important;
    border: 1px solid rgba(255, 255, 255, .55) !important;
    border-radius: var(--r-sm) !important;
    padding: 0 !important;
    margin: 2px 0 !important;
    min-height: 36px !important;
    height:     36px !important;
    width: 100% !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    font-size: 1rem !important;
    box-shadow: var(--sh-xs) !important;
    transition: all var(--t-fast) !important;
}}

[data-testid="stSidebar"] [data-testid="column"]:nth-child(2) button:hover {{
    background-color: #ffffff !important;
    transform: scale(1.06) !important;
    box-shadow: var(--sh-sm) !important;
}}

/* ── 4. CONTENUTO PRINCIPALE — CARDS ────────────────────────── */
div[data-testid="stVerticalBlockBordered"] {{
    background-color: var(--cna-surface) !important;
    border: 1px solid var(--cna-border) !important;
    border-left: 4px solid var(--cna-primary) !important;
    border-radius: var(--r-lg) !important;
    box-shadow: var(--sh-sm) !important;
    transition: box-shadow var(--t-base) !important;
}}

div[data-testid="stVerticalBlockBordered"]:hover {{
    box-shadow: var(--sh-md) !important;
}}

/* ── 5. TIPOGRAFIA ───────────────────────────────────────────── */
h1 {{
    color: var(--cna-text-heading) !important;
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    letter-spacing: -.025em !important;
    line-height: 1.2 !important;
}}

h2 {{
    color: var(--cna-text-heading) !important;
    font-size: 1.35rem !important;
    font-weight: 650 !important;
    letter-spacing: -.01em !important;
    line-height: 1.3 !important;
}}

h3 {{
    color: var(--cna-primary-text) !important;
    font-size: .9rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: .06em !important;
    padding-bottom: var(--sp-2) !important;
    border-bottom: 1px solid var(--cna-border-subtle) !important;
    margin-top: var(--sp-5) !important;
    margin-bottom: var(--sp-3) !important;
}}

h4 {{
    color: var(--cna-primary-text) !important;
    font-size: .95rem !important;
    font-weight: 600 !important;
}}

.stMarkdown p, p {{
    color: var(--cna-text-body) !important;
    line-height: 1.65 !important;
}}

.stCaption {{
    color: var(--cna-text-muted) !important;
    font-size: .8rem !important;
}}

label {{
    color: var(--cna-primary-text) !important;
    font-weight: 600 !important;
    font-size: .875rem !important;
}}

/* ── 6. PULSANTI ────────────────────────────────────────────── */
[data-testid="stButton"] button {{
    background-color: var(--cna-primary) !important;
    color: var(--cna-text-white) !important;
    border: none !important;
    box-shadow: var(--sh-sm) !important;
    font-weight: 500 !important;
    font-family: var(--font) !important;
    height: 38px !important;
    min-height: 38px !important;
    padding: 0 16px !important;
    border-radius: var(--r-md) !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    font-size: .9rem !important;
    letter-spacing: .01em !important;
    transition: background-color var(--t-fast), box-shadow var(--t-fast),
                transform var(--t-fast) !important;
    cursor: pointer !important;
}}

[data-testid="stButton"] button p,
[data-testid="stButton"] button span {{
    color: var(--cna-text-white) !important;
    background: transparent !important;
    margin: 0 !important;
    padding: 0 !important;
}}

[data-testid="stButton"] button:hover {{
    background-color: var(--cna-primary-dark) !important;
    box-shadow: var(--sh-md) !important;
    transform: translateY(-1px) !important;
}}

[data-testid="stButton"] button:hover * {{
    color: var(--cna-text-white) !important;
}}

[data-testid="stButton"] button:active {{
    transform: translateY(0) !important;
    box-shadow: var(--sh-xs) !important;
}}

[data-testid="stButton"] button:focus-visible {{
    outline: 2px solid var(--cna-primary) !important;
    outline-offset: 2px !important;
}}

/* Bottoni icona in blocco orizzontale (📂 🔍 ✏️) */
[data-testid="stHorizontalBlock"] [data-testid="stButton"] button {{
    width: 100% !important;
    font-size: 1.15rem !important;
    padding: 0 var(--sp-2) !important;
}}

[data-testid="stHorizontalBlock"] [data-testid="stButton"] {{
    margin-top: auto !important;
}}

/* Bottone "Sfoglia file..." (secondary) */
[data-testid="stBaseButton-secondary"],
div[data-testid="stFileUploader"] button,
div[data-testid="stFileUploaderDropzone"] button {{
    background-color: var(--cna-primary) !important;
    color: var(--cna-text-white) !important;
    border: none !important;
    box-shadow: var(--sh-sm) !important;
    font-weight: 500 !important;
    border-radius: var(--r-md) !important;
    transition: background-color var(--t-fast) !important;
}}

[data-testid="stBaseButton-secondary"]:hover,
div[data-testid="stFileUploader"] button:hover {{
    background-color: var(--cna-primary-dark) !important;
}}

[data-testid="stBaseButton-secondary"] p,
[data-testid="stBaseButton-secondary"] span {{
    color: var(--cna-text-white) !important;
}}

/* ── 7. INPUT FORM ──────────────────────────────────────────── */
div[data-baseweb="input"],
div[data-baseweb="select"] > div,
div[data-testid="stTextInput"] > div,
div[data-testid="stTextArea"] > div,
div[data-testid="stNumberInput"] > div {{
    border: 1.5px solid var(--cna-border) !important;
    background-color: var(--cna-surface) !important;
    border-radius: var(--r-md) !important;
    overflow: hidden !important;
    transition: border-color var(--t-fast), box-shadow var(--t-fast) !important;
}}

div[data-baseweb="input"]:focus-within,
div[data-baseweb="select"] > div:focus-within,
div[data-testid="stTextInput"] > div:focus-within,
div[data-testid="stTextArea"] > div:focus-within,
div[data-testid="stNumberInput"] > div:focus-within {{
    border-color: var(--cna-primary) !important;
    box-shadow: 0 0 0 3px hsla(210, 100%, {l_val}%, .14) !important;
}}

div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stNumberInput"] input,
div[data-baseweb="select"] span {{
    border: none !important;
    box-shadow: none !important;
    background-color: transparent !important;
    color: var(--cna-text-input) !important;
    outline: none !important;
    font-family: var(--font) !important;
    font-size: .9rem !important;
}}

/* Bottoni +/- NumberInput */
div[data-testid="stNumberInput"] button {{
    background-color: var(--cna-primary) !important;
    color: var(--cna-text-white) !important;
    border: none !important;
    border-left: 1px solid hsl(210, 100%, {l_dark}%) !important;
    margin: 0 !important;
    height: 100% !important;
    border-radius: 0 !important;
    transition: background-color var(--t-fast) !important;
}}

div[data-testid="stNumberInput"] button:hover {{
    background-color: var(--cna-primary-dark) !important;
}}

/* Placeholder */
::placeholder {{
    color: #94a3b8 !important;
    opacity: 1 !important;
    -webkit-text-fill-color: #94a3b8 !important;
}}

::-webkit-input-placeholder {{
    color: #94a3b8 !important;
    opacity: 1 !important;
    -webkit-text-fill-color: #94a3b8 !important;
}}

/* ── 8. FILE UPLOADER ───────────────────────────────────────── */
div[data-testid="stFileUploader"] {{
    background-color: #f8fbff !important;
    border: 2px dashed var(--cna-primary) !important;
    border-radius: var(--r-xl) !important;
    padding: var(--sp-5) !important;
    transition: background-color var(--t-base), border-color var(--t-base) !important;
}}

div[data-testid="stFileUploader"]:hover {{
    background-color: #eef5ff !important;
    border-color: var(--cna-primary-dark) !important;
}}

div[data-testid="stFileUploader"] section {{
    background-color: transparent !important;
}}

/* ── 9. HEADER TOOL (titolo h1 + bottone ⚙️ inline) ─────────── */
div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) {{
    gap: 6px !important;
    align-items: center !important;
}}

div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) > div[data-testid="stColumn"] {{
    flex: 0 0 auto !important;
    width: fit-content !important;
    max-width: 92% !important;
}}

div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) > div[data-testid="stColumn"]:last-child {{
    margin-top: 18px !important;
}}

div[data-testid="stHorizontalBlock"]:has(h1):not(:has([data-testid="stVerticalBlockBordered"])) div[data-testid="stTooltipHoverTarget"] {{
    justify-content: flex-start !important;
    width: auto !important;
}}

/* ── 10. SPLIT LAYOUT assistente (70 / 30) ──────────────────── */
[data-testid="stMainBlockContainer"]
> [data-testid="stVerticalBlock"]
> [data-testid="stHorizontalBlock"] {{
    align-items: flex-start !important;
}}

[data-testid="stMainBlockContainer"]
> [data-testid="stVerticalBlock"]
> [data-testid="stHorizontalBlock"]
> [data-testid="stColumn"]
> [data-testid="stVerticalBlockBordered"] {{
    height: calc(100vh - 80px) !important;
    max-height: calc(100vh - 80px) !important;
    overflow-y: auto !important;
    padding: 1rem !important;
    box-sizing: border-box !important;
}}

/* ── 11. EXPANDERS (contenuto principale) ───────────────────── */
[data-testid="stExpander"] {{
    border: 1px solid var(--cna-border) !important;
    border-radius: var(--r-lg) !important;
    background-color: var(--cna-surface) !important;
    box-shadow: var(--sh-xs) !important;
    margin-bottom: var(--sp-3) !important;
    transition: box-shadow var(--t-base) !important;
}}

[data-testid="stExpander"]:hover {{
    box-shadow: var(--sh-sm) !important;
}}

[data-testid="stExpander"] summary {{
    font-weight: 600 !important;
    color: var(--cna-primary-text) !important;
}}

/* ── 12. ASSISTENTE AI — HIGHLIGHT GOLD ─────────────────────── */
div[data-testid="stSidebar"] [data-testid="stExpander"]:has(input[key="assistant_query"]) {{
    border: 2px solid #f0b429 !important;
    background-color: rgba(240, 180, 41, .06) !important;
}}

/* ── 13. DIVIDER ────────────────────────────────────────────── */
hr[data-testid="stDivider"] {{
    border-color: var(--cna-border) !important;
    opacity: 0.5 !important;
    margin: var(--sp-4) 0 !important;
}}

/* ── 14. TOAST ──────────────────────────────────────────────── */
div[data-testid="stToast"] {{
    border-radius: var(--r-lg) !important;
    box-shadow: var(--sh-lg) !important;
}}

/* ── 15. SKELETON LOADER ────────────────────────────────────── */
@keyframes cna-skeleton-pulse {{
    0%, 100% {{ opacity: 1; }}
    50%       {{ opacity: .4; }}
}}

.cna-skeleton {{
    background: linear-gradient(90deg, #e2e8f0 25%, #f1f5f9 50%, #e2e8f0 75%);
    background-size: 200% 100%;
    animation: cna-skeleton-pulse 1.6s ease-in-out infinite;
    border-radius: var(--r-md);
    min-height: 20px;
}}

/* ── 16. EMPTY STATE ────────────────────────────────────────── */
.cna-empty-state {{
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 4rem 2rem;
    text-align: center;
    opacity: .72;
}}

.cna-empty-state .cna-empty-icon {{
    font-size: 3.5rem;
    margin-bottom: 1rem;
}}

.cna-empty-state .cna-empty-title {{
    font-size: 1.15rem;
    font-weight: 600;
    color: var(--cna-text-heading);
    margin: 0 0 .4rem;
}}

.cna-empty-state .cna-empty-sub {{
    font-size: .9rem;
    color: var(--cna-text-muted);
    max-width: 320px;
    line-height: 1.55;
    margin: 0;
}}

/* ── 17. ERROR CARD ─────────────────────────────────────────── */
.cna-error-card {{
    background: #fff5f5;
    border: 1px solid #fca5a5;
    border-left: 4px solid var(--cna-error);
    border-radius: var(--r-lg);
    padding: var(--sp-4) var(--sp-5);
    margin: var(--sp-3) 0;
}}

.cna-error-card .cna-error-title {{
    color: #991b1b;
    font-weight: 700;
    font-size: .875rem;
    margin: 0 0 var(--sp-1);
}}

.cna-error-card .cna-error-detail {{
    color: #7f1d1d;
    font-size: .8rem;
    font-family: var(--font-mono);
    white-space: pre-wrap;
    margin: 0;
    max-height: 180px;
    overflow-y: auto;
}}

/* ── 18. BADGE REGIONE ──────────────────────────────────────── */
.cna-badge {{
    display: inline-flex;
    align-items: center;
    gap: 4px;
    background-color: var(--cna-border-subtle);
    color: var(--cna-text-muted);
    font-size: .72rem;
    font-weight: 600;
    letter-spacing: .04em;
    text-transform: uppercase;
    padding: 2px 8px;
    border-radius: var(--r-xl);
    vertical-align: middle;
}}

</style>"""


def inject_styles(l_val: int) -> None:
    """
    Inietta il CSS globale dell'applicazione.

    Args:
        l_val: Luminosità sidebar (10–60). Determina il colore primario CNA.
    """
    st.markdown(generate_css(l_val), unsafe_allow_html=True)
