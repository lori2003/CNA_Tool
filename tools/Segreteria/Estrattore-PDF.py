from __future__ import annotations

import base64
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# ==============================================================================
# TOOLBOX METADATA
# ==============================================================================
TOOL = {
    'id': 'estrattore_pdf_invoice',
    'name': 'Estrattore Fatture (Mobile)',
    'description': (
        "#### 📌 1. FINALITÀ DEL TOOL\n"
        "Permette di trasformare foto scattate da smartphone o file PDF in documenti digitali archiviabili, "
        "applicando un effetto 'scanner' professionale e gestendo l'upload remoto via tunnel sicuro.\n\n"
        "#### 🚀 2. COME UTILIZZARLO\n"
        "1. **Connessione:** Spunta *Avvia Sync* e clicca **Esegui** per creare il tunnel Cloudflare. "
        "Aggiorna la pagina per vedere il QR Code / URL da inviare al telefono.\n"
        "2. **Scatto:** Dal telefono apri l'URL e carica le foto delle fatture indicando Fornitore e Descrizione.\n"
        "3. **Archiviazione:** Carica file manualmente oppure spunta *Archivia tutti* e clicca **Esegui** "
        "per convertire e spostare tutti i file in coda nella cartella di destinazione.\n\n"
        "#### 🧠 3. LOGICA DI ELABORAZIONE\n"
        "* **Computer Vision (OpenCV):** Edge detection + correzione prospettiva + CLAHE.\n"
        "* **Secure Tunneling:** Tunnel Cloudflare temporaneo (cloudflared) — nessuna configurazione di rete.\n\n"
        "#### 📂 4. RISULTATO FINALE\n"
        "File PDF ottimizzati archiviati nelle cartelle fornitore su Disco F:."
    ),
    'inputs': [
        {'key': '_sec_tunnel',   'label': '### 🔄 Sincronizzazione Cloudflare Tunnel', 'type': 'markdown'},
        {'key': '_sec_upload',   'label': '### 📤 Caricamento Manuale dal PC',         'type': 'markdown'},
        {'key': 'files_manuali', 'label': 'Carica file (PDF, JPG, PNG, WEBP)',         'type': 'file_multi'},
        {'key': '_sec_dash',     'label': '### 📂 Dashboard Gestionale',               'type': 'markdown'},
    ],
    'params': [
        # ── Tunnel ─────────────────────────────────────────────────────────
        {'key': 'tunnel_status', 'label': 'Stato tunnel corrente',
         'type': 'dynamic_info', 'function': 'get_tunnel_status'},
        {'key': 'avvia_tunnel',  'label': '▶ Avvia Sync (Cloudflare)',
         'type': 'checkbox', 'default': False},
        {'key': 'ferma_tunnel',  'label': '⏹ Stop / Reset Tunnel',
         'type': 'checkbox', 'default': False},
        {'key': 'kill_zombie',   'label': '🔥 Kill Zombie (termina cloudflared orfani)',
         'type': 'checkbox', 'default': False},
        # ── Dashboard ───────────────────────────────────────────────────────
        {'key': 'dashboard_files', 'label': 'File in attesa di elaborazione',
         'type': 'dynamic_info', 'function': 'get_dashboard_info'},
        {'key': 'archivia_tutti', 'label': '📦 Archivia tutti i file dalla coda',
         'type': 'checkbox', 'default': False},
        # ── Configurazione ──────────────────────────────────────────────────
        {'key': 'output_path', 'label': '📁 Cartella Destinazione Archivio',
         'type': 'text', 'default': r'F:\Cna Pensionati\CNA PENSIONATI 2026\Fatture',
         'placeholder': r'Es: F:\CNA\Fatture'},
        # ── Elaborazione immagini ───────────────────────────────────────────
        {'key': 'auto_crop',        'label': '🔲 Auto-Crop (rilevamento bordi documento)',
         'type': 'checkbox', 'default': True},
        {'key': 'white_background', 'label': '⬜ Sfondo Bianco',
         'type': 'checkbox', 'default': True},
        {'key': 'brightness',       'label': '☀️ Luminosità (-50 / +50)',
         'type': 'number',   'default': 35},
        {'key': 'contrast',         'label': '🌗 Contrasto (0.5 – 2.0)',
         'type': 'number',   'default': 1.4},
        {'key': 'clahe_strength',   'label': '📊 CLAHE – contrasto locale (0 = disabilitato)',
         'type': 'number',   'default': 0.0},
    ],
}

# ==============================================================================
# CONFIG & CONSTANTS
# ==============================================================================
PORT = 8088
SCRIPT_PATH = Path(__file__).resolve()
TOOLBOX_ROOT = SCRIPT_PATH.parents[2]
LOG_FILE = TOOLBOX_ROOT / "estrattore_pdf.log"
SYNC_BRIDGE = TOOLBOX_ROOT / "sync_bridge"
SYNC_BRIDGE.mkdir(parents=True, exist_ok=True)
TUNNEL_URL_FILE = SYNC_BRIDGE / "tunnel_url.txt"

EXTENSION_DIR = SCRIPT_PATH.parent / "extension" / "x estrattore - pdf"
EXTENSION_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = EXTENSION_DIR / "config_manual.json"

# Regex per catturare URL da Cloudflare Tunnel (cloudflared).
# Formato: "https://random-words.trycloudflare.com"
URL_RE_CF = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", re.I)

VALID_EXTS = {".pdf", ".jpg", ".jpeg", ".png", ".webp"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp"}
CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0

# Stato tunnel persistente a livello di modulo (sopravvive tra le run)
_tunnel_proc: Optional[subprocess.Popen] = None

# ==============================================================================
# OPENCV LOCAL INSTALLATION (Lazy Load - NON modifica sys.path all'avvio!)
# ==============================================================================
OPENCV_LOCAL_DIR = EXTENSION_DIR / "OpenCV"
_opencv_path_added = False  # Flag per evitare di aggiungere più volte

def _get_cv2():
    """
    Importa cv2 in modo sicuro, cercando prima nell'installazione locale.
    IMPORTANTE: Modifica sys.path SOLO qui, non all'avvio del modulo.
    """
    global _opencv_path_added
    
    # 1. Prova import standard (installazione globale)
    try:
        import cv2
        return cv2
    except ImportError:
        pass
    except Exception:
        pass
    
    # 2. Se non trovato, prova dalla cartella locale
    if OPENCV_LOCAL_DIR.exists() and not _opencv_path_added:
        local_path = str(OPENCV_LOCAL_DIR)
        if local_path not in sys.path:
            sys.path.append(local_path)
            _opencv_path_added = True
        
        try:
            import cv2
            return cv2
        except ImportError:
            pass
        except Exception as e:
            if local_path in sys.path:
                sys.path.remove(local_path)
            log_event(f"OpenCV locale errore: {e}")
            return None
    
    # 3. OpenCV non disponibile - PROVA AUTO-INSTALLAZIONE (Limitata)
    # Evita ricorsione infinita usando st.session_state se disponibile
    install_tried = False
    try:
        from core.toolkit import ctx as st
        # Se abbiamo già provato in QUESTA sessione di esecuzione script, non riprovare
        if "opencv_install_attempted" in st.session_state:
            install_tried = True
    except: pass
        
    if not install_tried:
        log_event("⚠️ OpenCV non trovato. Tento installazione automatica...")
        try:
            from core.toolkit import ctx as st
            st.session_state["opencv_install_attempted"] = True
            st.toast("📦 Installazione OpenCV in corso...", icon="⏳")
        except: pass
            
        success, output = install_opencv_locally()
        
        if success:
            log_event("✅ OpenCV installato con successo! Riprovo import...")
            try:
                from core.toolkit import ctx as st
                st.toast("✅ OpenCV installato! Ricarica in corso...", icon="🎉")
                time.sleep(1) 
                st.rerun()
            except: pass
        else:
            log_event(f"❌ Installazione automatica fallita: {output}")

    return None

def install_opencv_locally():
    """Installa OpenCV nella cartella locale. Chiamato solo su richiesta utente."""
    import subprocess
    OPENCV_LOCAL_DIR.mkdir(parents=True, exist_ok=True)
    
    cmd = [
        sys.executable, "-m", "pip", "install",
        "--target", str(OPENCV_LOCAL_DIR),
        "--upgrade", "--no-cache-dir",
        "opencv-python-headless"  # Versione leggera senza GUI
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0, result.stdout + result.stderr


# ==============================================================================
# SERVER MODE CHECK (Fail Fast)
# ==============================================================================
if "--server" in sys.argv:
    try:
        import uvicorn
        from fastapi import Depends, FastAPI, File, HTTPException, UploadFile, Query, Form
        from fastapi.responses import HTMLResponse
        from fastapi.security import HTTPBasic, HTTPBasicCredentials
        import requests
        import qrcode
        from PIL import Image
    except ImportError as e:
        with open(LOG_FILE, "a") as f: f.write(f"[FATAL] Missing dependency: {e}\n")
        sys.exit(1)

from core.toolkit import ctx as st
try:
    import qrcode
    import requests
    from PIL import Image, ImageOps
except ImportError:
    pass

# ==============================================================================
# DATA CLASSES & UTILS
# ==============================================================================
@dataclass
class ManualFileState:
    source_filename: str
    file_hash: str
    current_path: str  # Path del file da archiviare (PDF)
    preview_path: str = ""  # Path per anteprima (immagine originale se disponibile)
    archived_path: str = ""  # Path del file archiviato (dopo "Archivia")
    name_prefix: str = ""
    description: str = ""
    imponibile: float = 0.0
    iva_percent: float = 22.0
    is_converted: bool = False
    processed: bool = False

    def __post_init__(self):
        # Se nome e descrizione sono vuoti, prova a estrarli dal nome file
        # Formato atteso: "NomeFornitore - Descrizione.ext" o "Nome - Desc_1.ext"
        if not self.name_prefix and not self.description:
            try:
                stem = Path(self.source_filename).stem
                if " - " in stem:
                    parts = stem.split(" - ", 1)
                    self.name_prefix = parts[0].strip()
                    desc_part = parts[1].strip()
                    # Rimuovi eventuale contatore _1, _2 finale se presente (opzionale, ma pulisce la UI)
                    if re.search(r"_\d+$", desc_part):
                        desc_part = re.sub(r"_\d+$", "", desc_part)
                    self.description = desc_part
            except: pass

def log_event(msg: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    try:
        with open(LOG_FILE, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{stamp}] {msg}\n")
    except: pass

def get_file_hash(path: Path) -> str:
    h = hashlib.md5()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(2*1024*1024), b""): h.update(chunk)
        return h.hexdigest()
    except: return ""

def load_config() -> Dict[str, Any]:
    default = {"output_path": r"F:\Cna Pensionati\CNA PENSIONATI 2026\FORNITORI"}
    if CONFIG_FILE.exists():
        try: return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except: pass
    return default

def save_config(cfg: Dict[str, Any]):
    try: CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except: pass

def apply_scanner_effect(img: Image.Image, force_rotate: int = 0) -> Image.Image:
    """Applica correzioni base all'immagine (EXIF, rotazione, nitidezza).
    
    Nota: La maggior parte dell'elaborazione è ora in apply_opencv_scanner.
    Questa funzione gestisce solo le correzioni residue.
    """
    from PIL import ImageFilter, ImageOps
    
    # 1. Correggi orientamento EXIF
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    
    # 2. Rotazione forzata se richiesta
    if force_rotate:
        img = img.rotate(-force_rotate, expand=True)
    
    # 3. Sharpening leggero per testo
    img = img.filter(ImageFilter.SHARPEN)
    
    return img


def apply_opencv_scanner(
    img: Image.Image, 
    rotation_override: int = 0,
    brightness: int = 0,        # -100 to +100
    contrast: float = 1.0,      # 0.5 to 2.0
    clahe_strength: float = 3.0, # 0 to 5.0 (0 = disabilitato)
    white_background: bool = True,
    auto_crop: bool = True,
    crop_margin: float = 0.02   # 0 to 0.1 (percentuale margine)
) -> Image.Image:
    """
    Scanner documenti completo con OpenCV.
    
    Args:
        img: Immagine PIL da processare
        rotation_override: Rotazione manuale (0, 90, 180, 270). 0 = auto-detect.
        brightness: Regolazione luminosità (-100 a +100)
        contrast: Moltiplicatore contrasto (0.5 a 2.0)
        clahe_strength: Forza CLAHE (0=off, 1-5)
        white_background: Se True, rende lo sfondo bianco
        auto_crop: Se True, tenta di rilevare e ritagliare il documento
        crop_margin: Margine interno dopo il crop (0-0.1)
    """
    cv2 = _get_cv2()
    if cv2 is None:
        log_event("OpenCV non disponibile.")
        return img
    
    try:
        import numpy as np
    except ImportError:
        log_event("NumPy non installato")
        return img
    
    try:
        # Converti PIL -> OpenCV (RGB -> BGR)
        img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        orig = img_cv.copy()
        processed = img_cv.copy()
        
        # ========================================
        # FASE 0: Rotazione manuale
        # ========================================
        if rotation_override:
            if rotation_override == 90:
                img_cv = cv2.rotate(img_cv, cv2.ROTATE_90_CLOCKWISE)
                orig = cv2.rotate(orig, cv2.ROTATE_90_CLOCKWISE)
            elif rotation_override == 180:
                img_cv = cv2.rotate(img_cv, cv2.ROTATE_180)
                orig = cv2.rotate(orig, cv2.ROTATE_180)
            elif rotation_override == 270:
                img_cv = cv2.rotate(img_cv, cv2.ROTATE_90_COUNTERCLOCKWISE)
                orig = cv2.rotate(orig, cv2.ROTATE_90_COUNTERCLOCKWISE)
            processed = img_cv.copy()
        
        # ========================================
        # FASE 1: Auto-Crop (rilevamento documento)
        # ========================================
        doc_found = False
        if auto_crop:
            # Pre-processing per edge detection
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            gray = cv2.GaussianBlur(gray, (5, 5), 0)
            
            # Edge detection con soglie adattive
            edged = cv2.Canny(gray, 50, 200)
            
            # Morfologia per chiudere i contorni
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
            edged = cv2.dilate(edged, kernel, iterations=2)
            edged = cv2.morphologyEx(edged, cv2.MORPH_CLOSE, kernel)
            
            # Trova contorni
            contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            if contours:
                # Ordina per area
                contours = sorted(contours, key=cv2.contourArea, reverse=True)
                
                for c in contours[:5]:
                    peri = cv2.arcLength(c, True)
                    approx = cv2.approxPolyDP(c, 0.02 * peri, True)
                    
                    area = cv2.contourArea(c)
                    img_area = img_cv.shape[0] * img_cv.shape[1]
                    
                    # Richiedi almeno 15% dell'immagine
                    if area > img_area * 0.15:
                        if len(approx) == 4:
                            # Contorno a 4 vertici trovato - prospettiva
                            pts = approx.reshape(4, 2)
                            rect = np.zeros((4, 2), dtype="float32")
                            
                            s = pts.sum(axis=1)
                            rect[0] = pts[np.argmin(s)]
                            rect[2] = pts[np.argmax(s)]
                            
                            diff = np.diff(pts, axis=1)
                            rect[1] = pts[np.argmin(diff)]
                            rect[3] = pts[np.argmax(diff)]
                            
                            (tl, tr, br, bl) = rect
                            maxWidth = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
                            maxHeight = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
                            
                            margin_w = int(maxWidth * crop_margin)
                            margin_h = int(maxHeight * crop_margin)
                            
                            dst = np.array([
                                [0, 0], 
                                [maxWidth - 1, 0],
                                [maxWidth - 1, maxHeight - 1], 
                                [0, maxHeight - 1]
                            ], dtype="float32")
                            
                            M = cv2.getPerspectiveTransform(rect, dst)
                            processed = cv2.warpPerspective(orig, M, (maxWidth, maxHeight))
                            
                            # Applica margine interno
                            if margin_h > 0 and margin_w > 0:
                                processed = processed[margin_h:-margin_h, margin_w:-margin_w]
                            
                            doc_found = True
                            log_event(f"OpenCV: Documento 4-vertici rilevato ({processed.shape[1]}x{processed.shape[0]})")
                            break
                        else:
                            # Fallback: usa bounding rect
                            x, y, w, h = cv2.boundingRect(c)
                            margin_w = int(w * crop_margin)
                            margin_h = int(h * crop_margin)
                            
                            x1 = max(0, x + margin_w)
                            y1 = max(0, y + margin_h)
                            x2 = min(orig.shape[1], x + w - margin_w)
                            y2 = min(orig.shape[0], y + h - margin_h)
                            
                            processed = orig[y1:y2, x1:x2].copy()
                            doc_found = True
                            log_event(f"OpenCV: Bounding rect usato ({processed.shape[1]}x{processed.shape[0]})")
                            break
        
        if not doc_found:
            log_event("OpenCV: Nessun ritaglio applicato")
            processed = orig.copy()
        
        # ========================================
        # FASE 2: Auto-rotazione (landscape -> portrait)
        # ========================================
        h_proc, w_proc = processed.shape[:2]
        if w_proc > h_proc * 1.2 and not rotation_override:
            processed = cv2.rotate(processed, cv2.ROTATE_90_CLOCKWISE)
            log_event("OpenCV: Auto-rotazione applicata")
        
        # ========================================
        # FASE 3: Luminosità e Contrasto
        # ========================================
        if brightness != 0 or contrast != 1.0:
            processed = cv2.convertScaleAbs(processed, alpha=contrast, beta=brightness)
        
        # ========================================
        # FASE 4: CLAHE (contrasto adattivo)
        # ========================================
        if clahe_strength > 0:
            lab = cv2.cvtColor(processed, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            clahe = cv2.createCLAHE(clipLimit=clahe_strength, tileGridSize=(8, 8))
            l = clahe.apply(l)
            lab = cv2.merge([l, a, b])
            processed = cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)
        
        # ========================================
        # FASE 5: Sfondo bianco
        # ========================================
        if white_background:
            gray_out = cv2.cvtColor(processed, cv2.COLOR_BGR2GRAY)
            _, bg_mask = cv2.threshold(gray_out, 200, 255, cv2.THRESH_BINARY)
            processed[bg_mask == 255] = [255, 255, 255]
        
        # Converti OpenCV -> PIL
        result = Image.fromarray(cv2.cvtColor(processed, cv2.COLOR_BGR2RGB))
        log_event("OpenCV: Elaborazione completata")
        return result
        
    except Exception as e:
        log_event(f"OpenCV errore: {e}")
        return img


def convert_to_pdf(img_path: Path) -> tuple[Path, Path]:
    """Converte immagine in PDF con effetto scansione. Restituisce (pdf_path, preview_path)."""
    pdf_path = img_path.with_suffix(".pdf")
    # Salva copia dell'immagine originale per anteprima
    preview_path = img_path.parent / f"_preview_{img_path.name}"
    try:
        # Copia immagine originale per anteprima
        shutil.copy2(img_path, preview_path)
        
        # Carica e applica effetto scanner
        img = Image.open(img_path)
        if img.mode != 'RGB': img = img.convert('RGB')
        
        # Step 1: OpenCV - Rilevamento bordi e correzione prospettiva (se disponibile)
        img = apply_opencv_scanner(img)
        
        # Step 2: Effetto scansione classico (luminosità, contrasto, nitidezza)
        img = apply_scanner_effect(img)
        
        # Salva come PDF
        img.save(pdf_path, "PDF", resolution=100.0)
        img_path.unlink()  # Elimina originale (abbiamo la copia)
        return pdf_path, preview_path
    except Exception as e:
        log_event(f"Conv err: {e}")
        return img_path, img_path  # Fallback: usa originale per entrambi

def kill_port_win(port: int) -> None:
    try:
        cmd = f'netstat -ano | findstr LISTENING | findstr ":{port}"'
        res = subprocess.run(cmd, shell=True, capture_output=True, text=True, errors="replace")
        if res.returncode == 0:
            for line in res.stdout.splitlines():
                parts = line.strip().split()
                if parts:
                    pid = parts[-1]
                    if pid.isdigit(): 
                        subprocess.run(f"taskkill /F /PID {pid}", shell=True, capture_output=True, creationflags=CREATE_NO_WINDOW)
    except: pass

def kill_cloudflared_processes() -> None:
    """Termina tutti i processi cloudflared orfani per evitare conflitti di tunnel."""
    try:
        subprocess.run("taskkill /F /IM cloudflared.exe", shell=True, capture_output=True, creationflags=CREATE_NO_WINDOW)
    except: pass

def find_cloudflared() -> str:
    """Restituisce il percorso dell'eseguibile cloudflared (PATH o bin/ locale)."""
    local = TOOLBOX_ROOT / "bin" / "cloudflared.exe"
    if local.exists():
        return str(local)
    found = shutil.which("cloudflared")
    if found:
        return found
    return "cloudflared"  # Fallback: assume sia in PATH


# ==============================================================================
# DYNAMIC INFO FUNCTIONS
# ==============================================================================

def get_tunnel_status(params: dict) -> str:
    """Mostra lo stato attuale del tunnel Cloudflare."""
    global _tunnel_proc
    is_running = _tunnel_proc is not None and _tunnel_proc.poll() is None

    if is_running:
        ctx.info("🚀 **Tunnel attivo** — server in esecuzione")
        if TUNNEL_URL_FILE.exists():
            try:
                url = TUNNEL_URL_FILE.read_text("utf-8").strip()
                if URL_RE_CF.match(url):
                    ctx.success(
                        f"✅ **Connessione stabile OK!**\n\n"
                        f"URL tunnel: `{url}`\n\n"
                        f"Scansiona il QR Code con il telefono o apri l'URL direttamente."
                    )
                    return url
            except Exception:
                pass
        ctx.warning("⏳ In attesa di URL da Cloudflare... (aggiorna tra qualche secondo)")
    else:
        ctx.info(
            "⏹ **Tunnel non attivo.**\n\n"
            "Spunta *▶ Avvia Sync* e clicca **Esegui** per avviarlo."
        )
    return ""


def get_dashboard_info(params: dict) -> str:
    """Mostra i file in attesa di elaborazione nella coda SYNC_BRIDGE."""
    files = sorted([
        f for f in SYNC_BRIDGE.glob("*")
        if f.suffix.lower() in VALID_EXTS and not f.name.startswith("_preview_")
    ])
    if not files:
        ctx.info("📭 Nessun file in coda.")
        return ""
    lines = [f"📁 **{len(files)} file in coda** (pronto per archivazione):"]
    for f in files:
        try:
            size_kb = f.stat().st_size // 1024
            lines.append(f"- `{f.name}` ({size_kb} KB)")
        except Exception:
            lines.append(f"- `{f.name}`")
    ctx.info("\n".join(lines))
    return str(len(files))


# ==============================================================================
# FASTAPI SERVER
# ==============================================================================
if "--server" in sys.argv:
    server_app = FastAPI()
    server_security = HTTPBasic(auto_error=False)

    def _auth(credentials: Optional[HTTPBasicCredentials] = Depends(server_security)):
        u, p = os.getenv("UP_USER", ""), os.getenv("UP_PASS", "")
        if u and (not credentials or credentials.username != u or credentials.password != p):
            raise HTTPException(401, headers={"WWW-Authenticate": "Basic"})

    @server_app.get("/health")
    def health_check():
        return {"status": "ok", "source": "Toolbox-CF"}

    # ==========================================================================
    # TESTING BLOCK - Mobile UI Enhanced
    # ==========================================================================
    def get_mobile_html(token: str, message: str = "") -> str:
        msg_html = f"<div class='toast'>{message}</div>" if message else ""
        return f"""
        <!DOCTYPE html>
        <html lang="it">
        <head>
            <meta charset="UTF-8">
            <title>SSH Upload</title>
            <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                :root {{ --p: #dc3545; --bg: #0f172a; --card: #1e293b; --success: #10b981; }}
                * {{ box-sizing: border-box; }}
                body {{ font-family: 'Inter', sans-serif; background: var(--bg); color: #f8fafc; margin: 0; min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 1rem; }}
                .card {{ background: var(--card); padding: 1.5rem; border-radius: 1.5rem; width: 100%; max-width: 400px; text-align: center; border: 2px solid var(--p); position: relative; }}
                h2 {{ margin: 0 0 1rem 0; font-weight: 700; color: white; font-size: 1.5rem; }}
                .file-label {{ display: flex; flex-direction: column; align-items: center; background: rgba(220,53,69,0.1); border: 3px dashed var(--p); padding: 2rem 1rem; border-radius: 1rem; cursor: pointer; transition: all 0.3s; }}
                .file-label:active {{ transform: scale(0.98); background: rgba(220,53,69,0.2); }}
                .icon {{ font-size: 3rem; margin-bottom: 0.5rem; }}
                input[type="file"] {{ display: none; }}
                .toast {{ position: fixed; top: 20px; left: 50%; transform: translateX(-50%); background: var(--success); padding: 0.75rem 1.5rem; border-radius: 0.75rem; font-weight: 600; z-index: 100; animation: fadeIn 0.3s; }}
                @keyframes fadeIn {{ from {{ opacity: 0; transform: translate(-50%, -20px); }} to {{ opacity: 1; transform: translate(-50%, 0); }} }}
                
                .upload-btn {{ display: none; width: 100%; padding: 1rem; background: var(--p); color: white; border: none; border-radius: 1rem; font-size: 1.1rem; font-weight: 700; cursor: pointer; margin-top: 1rem; transition: all 0.3s; }}
                .upload-btn.show {{ display: block; }}
                .add-btn {{ display: none; width: 100%; padding: 1rem; border-radius: 1rem; font-size: 1rem; font-weight: 700; cursor: pointer; margin-top: 0.5rem; transition: all 0.3s; background: rgba(255,255,255,0.1); color: white; border: 2px dashed rgba(255,255,255,0.3); }}
                .add-btn.show {{ display: block; }}
                
                /* File Rows */
                #file-list-container {{ max-height: 400px; overflow-y: auto; text-align: left; margin-bottom: 1rem; }}
                .file-row {{ display: flex; flex-direction: column; gap: 0.5rem; background: rgba(255,255,255,0.05); padding: 1rem; border-radius: 1rem; margin-bottom: 1rem; border: 1px solid rgba(255,255,255,0.1); }}
                .file-info {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.25rem; }}
                .file-name {{ font-weight: 600; font-size: 0.85rem; color: #cbd5e1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 200px; }}
                .file-remove {{ color: #dc3545; cursor: pointer; font-size: 1.2rem; padding: 0.2rem 0.5rem; font-weight: bold; }}
                .file-inputs {{ display: flex; gap: 0.5rem; }}
                .input-field {{ flex: 1; padding: 0.8rem; border-radius: 0.5rem; border: 1px solid #475569; background: #1e293b; color: white; outline: none; font-size: 0.9rem; width: 100%; }}
                .input-field:focus {{ border-color: var(--p); }}
                
                .hint {{ color: #94a3b8; font-size: 0.8rem; margin-top: 0.5rem; }}
            </style>
        </head>
        <body>
            <div class="card">
                {msg_html}
                <h2>📱 Caricamento Remoto</h2>
                
                <form id="up-form" method="post" enctype="multipart/form-data">
                    <div id="file-list-container"></div>
                
                    <label class="file-label" id="file-label">
                        <input type="file" id="file-input" multiple accept="image/*,application/pdf" capture="environment">
                        <div class="icon">📷</div>
                        <div>Tocca per aggiungere foto</div>
                    </label>
                    
                    <button type="button" class="add-btn" id="add-btn">➕ Aggiungi Altre</button>
                    <button type="button" class="upload-btn" id="upload-btn">🚀 CARICA TUTTO</button>
                </form>
            </div>
            
            <script>
                const fileInput = document.getElementById('file-input');
                const listContainer = document.getElementById('file-list-container');
                const uploadBtn = document.getElementById('upload-btn');
                const addBtn = document.getElementById('add-btn');
                const fileLabel = document.getElementById('file-label');
                
                let allFiles = []; // Array of objects: {{ file: File, id: uniqueId, name: '', desc: '' }}
                let uniqueIdCounter = 0;

                function render() {{
                    listContainer.innerHTML = '';
                    if (allFiles.length > 0) {{
                        fileLabel.style.display = 'none';
                        uploadBtn.classList.add('show');
                        addBtn.classList.add('show');
                        
                        allFiles.forEach((item) => {{
                            const div = document.createElement('div');
                            div.className = 'file-row';
                            div.innerHTML = `
                                <div class="file-info">
                                    <span class="file-name">📄 ${{item.file.name}}</span>
                                    <span class="file-remove" onclick="removeFile(${{item.id}})">✕</span>
                                </div>
                                <div class="file-inputs">
                                    <input type="text" class="input-field" placeholder="Fornitore" id="name-${{item.id}}" value="${{item.name}}" oninput="updateMeta(${{item.id}}, 'name', this.value)">
                                    <input type="text" class="input-field" placeholder="Descrizione" id="desc-${{item.id}}" value="${{item.desc}}" oninput="updateMeta(${{item.id}}, 'desc', this.value)">
                                </div>
                            `;
                            listContainer.appendChild(div);
                        }});
                        
                    }} else {{
                        fileLabel.style.display = 'flex';
                        uploadBtn.classList.remove('show');
                        addBtn.classList.remove('show');
                    }}
                }}

                window.removeFile = function(id) {{
                    allFiles = allFiles.filter(f => f.id !== id);
                    render();
                }};

                window.updateMeta = function(id, key, val) {{
                    const f = allFiles.find(item => item.id === id);
                    if (f) f[key] = val;
                }};

                fileInput.addEventListener('change', function() {{
                    for (let i = 0; i < this.files.length; i++) {{
                        allFiles.push({{ 
                            file: this.files[i], 
                            id: uniqueIdCounter++,
                            name: '',
                            desc: ''
                        }});
                    }}
                    this.value = '';
                    render();
                }});

                addBtn.addEventListener('click', () => fileInput.click());

                uploadBtn.addEventListener('click', function() {{
                    if (allFiles.length === 0) return;
                    
                    // Validazione
                    for (let item of allFiles) {{
                        if (!item.name || !item.desc) {{
                            alert(`⚠️ Compila Fornitore e Descrizione per: ${{item.file.name}}`);
                            return;
                        }}
                    }}
                    
                    uploadBtn.textContent = '⏳ Caricamento...';
                    uploadBtn.disabled = true;
                    addBtn.style.display = 'none';
                    
                    const formData = new FormData();
                    // Append grouped to ensure clean lists on backend
                    allFiles.forEach(item => formData.append('f', item.file));
                    allFiles.forEach(item => formData.append('names', item.name));
                    allFiles.forEach(item => formData.append('descs', item.desc));
                    
                    fetch('/up?t={token}', {{ method: 'POST', body: formData }})
                    .then(r => {{
                        if (!r.ok) throw new Error("Status " + r.status);
                        return r.text();
                    }})
                    .then(html => {{ document.open(); document.write(html); document.close(); }})
                    .catch(e => {{ 
                        alert('❌ Errore durante il caricamento: ' + e); 
                        uploadBtn.textContent = '🚀 CARICA TUTTO'; 
                        uploadBtn.disabled = false; 
                        addBtn.style.display = 'block'; 
                    }});
                }});
            </script>
        </body>
        </html>
        """
    # ==========================================================================
    # END TESTING BLOCK
    # ==========================================================================

    @server_app.get("/", response_class=HTMLResponse)
    def index(_=Depends(_auth)):
        return get_mobile_html(os.getenv("UP_TOKEN", ""))

    @server_app.post("/up")
    async def upload(
        t: str = Query(...), 
        f: List[UploadFile] = File(...), 
        names: List[str] = Form(...),  
        descs: List[str] = Form(...), 
        _=Depends(_auth)
    ):
        token = os.getenv("UP_TOKEN", "")
        if t != token: raise HTTPException(403)
        
        try:
            # Validazione lunghezze
            log_event(f"UP REQUEST: files={len(f)}, names={names}, descs={descs}")
            
            # Se per qualche motivo arrivano liste diverse, normalizziamo
            max_len =  max(len(f), len(names), len(descs))
            
            success_count = 0
            
            for i, file in enumerate(f):
                try:
                    name = names[i] if i < len(names) else "Sconosciuto"
                    desc = descs[i] if i < len(descs) else "File"
                    
                    log_event(f"Processing File {i}: {file.filename} -> {name} - {desc}")
                    
                    # 1. Determina nome base "Pulito"
                    if not name.strip() and not desc.strip():
                         base_name = Path(file.filename).stem
                    else:
                         base_name = f"{name.strip()} - {desc.strip()}"
                    
                    # Estensione sicura
                    orig_name = file.filename or "file"
                    ext = Path(orig_name).suffix
                    if not ext: ext = ".jpg"
                    
                    # 2. Genera nome file con gestione duplicati (counter incrementale)
                    final_name = f"{base_name}{ext}"
                    counter = 1
                    while (SYNC_BRIDGE / final_name).exists():
                        final_name = f"{base_name}_{counter}{ext}"
                        counter += 1
                    
                    data = await file.read()
                    log_event(f"Upload received: {final_name} ({len(data)} bytes)")
                    (SYNC_BRIDGE / final_name).write_bytes(data)
                    success_count += 1
                    
                except Exception as inner_e:
                    log_event(f"Single File Error: {inner_e}")
                    # Non fermare tutto se uno fallisce, ma segnalalo?
                    # Per semplicità, continuiamo
                    continue
                    
            return HTMLResponse(get_mobile_html(token, f"✅ {success_count} File Caricati con Successo!"))
            
        except Exception as e:
            log_event(f"CRITICAL UPLOAD ERROR: {e}")
            return HTMLResponse(get_mobile_html(token, f"❌ Errore Critico: {e}"))

    def run_server_process(port: int):
        log_event("--- START CLOUDFLARE TUNNEL SERVER ---")
        kill_port_win(port)
        time.sleep(0.5)

        # 1. AVVIA UVICORN IN BACKGROUND (così è pronto quando cloudflared si connette)
        def _run_uvicorn():
            log_event(f"Uvicorn Starting on 127.0.0.1:{port}...")
            try:
                uvicorn.run(server_app, host="127.0.0.1", port=port, log_level="warning")
            except Exception as e:
                log_event(f"Uvicorn Crash: {e}")

        uvicorn_thread = threading.Thread(target=_run_uvicorn, daemon=True)
        uvicorn_thread.start()
        time.sleep(1.0)  # Dai tempo a Uvicorn di partire

        # 2. CLOUDFLARE TUNNEL CON AUTO-RICONNESSIONE
        cf_bin = find_cloudflared()
        MAX_RETRIES = 999  # Praticamente infinito
        retry_count = 0

        while retry_count < MAX_RETRIES:
            try:
                TUNNEL_URL_FILE.unlink(missing_ok=True)
            except: pass

            log_event(f"Cloudflare Tunnel Starting (attempt {retry_count + 1}) using: {cf_bin}")

            cmd = [cf_bin, "tunnel", "--url", f"http://127.0.0.1:{port}", "--no-autoupdate"]

            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, creationflags=CREATE_NO_WINDOW)

            # Leggi l'output nel main thread (tiene vivo il processo)
            for line in proc.stdout:
                log_event(f"CF: {line.strip()}")
                m = URL_RE_CF.search(line)
                if m:
                    try: TUNNEL_URL_FILE.write_text(m.group(0), encoding="utf-8")
                    except: pass

            # cloudflared è terminato - aspetta e riprova
            exit_code = proc.wait()
            log_event(f"Cloudflare Tunnel Closed (exit code: {exit_code}). Reconnecting in 5s...")
            retry_count += 1
            time.sleep(5)

        log_event("Cloudflare Tunnel: Max retries reached.")

# ==============================================================================
# UI LOGIC
# ==============================================================================
def run(
    out_dir,
    output_path="",
    avvia_tunnel=False,
    ferma_tunnel=False,
    kill_zombie=False,
    archivia_tutti=False,
    files_manuali=None,
    auto_crop=True,
    white_background=True,
    brightness=35,
    contrast=1.4,
    clahe_strength=0.0,
    **kwargs,
):
    """Gestisce tunnel, upload manuale e archiviazione file."""
    global _tunnel_proc
    output_files = []

    # ── 1. Tunnel: stop / kill zombie ─────────────────────────────────────────
    if kill_zombie:
        kill_cloudflared_processes()
        kill_port_win(PORT)
        ctx.success("✅ Processi zombie cloudflared terminati.")

    if ferma_tunnel:
        if _tunnel_proc is not None:
            try:
                _tunnel_proc.terminate()
            except Exception:
                pass
            _tunnel_proc = None
        kill_cloudflared_processes()
        kill_port_win(PORT)
        try:
            TUNNEL_URL_FILE.unlink(missing_ok=True)
        except Exception:
            pass
        ctx.success("⏹ Tunnel fermato e risorse liberate.")

    # ── 2. Tunnel: avvio ──────────────────────────────────────────────────────
    if avvia_tunnel and not ferma_tunnel:
        is_running = _tunnel_proc is not None and _tunnel_proc.poll() is None
        if not is_running:
            token = base64.urlsafe_b64encode(os.urandom(12)).decode()
            env = os.environ.copy()
            env.update({"UP_TOKEN": token})
            args = [sys.executable, str(SCRIPT_PATH), "--server", "--port", str(PORT)]
            _tunnel_proc = subprocess.Popen(args, env=env, creationflags=CREATE_NO_WINDOW)
            log_event(f"Tunnel server avviato (PID {_tunnel_proc.pid})")
            ctx.success(
                f"▶ **Server avviato** (PID {_tunnel_proc.pid}).\n\n"
                "Aggiorna la pagina tra qualche secondo per vedere l'URL del tunnel."
            )
        else:
            ctx.info("ℹ️ Tunnel già in esecuzione.")

    # ── 3. Salva file caricati manualmente ────────────────────────────────────
    if files_manuali:
        if not isinstance(files_manuali, list):
            files_manuali = [files_manuali]
        saved = 0
        for fd in files_manuali:
            if not isinstance(fd, dict):
                continue
            filename = fd.get("filename", "file")
            content = fd.get("content", b"")
            ext = Path(filename).suffix.lower()
            if ext not in VALID_EXTS:
                ctx.warning(f"⚠️ Estensione non supportata: `{filename}`")
                continue
            bn = Path(filename).stem
            target = SYNC_BRIDGE / filename
            cnt = 1
            while target.exists():
                target = SYNC_BRIDGE / f"{bn}_{cnt}{ext}"
                cnt += 1
            target.write_bytes(content)
            saved += 1
            ctx.info(f"📥 Salvato in coda: `{target.name}`")
        if saved:
            ctx.success(f"✅ {saved} file salvati nella coda di elaborazione.")

    # ── 4. Archivia tutti i file dalla coda ───────────────────────────────────
    if archivia_tutti:
        dest_root = Path(output_path) if output_path else Path(
            r"F:\Cna Pensionati\CNA PENSIONATI 2026\Fatture"
        )
        dest_root.mkdir(parents=True, exist_ok=True)

        queue = sorted([
            f for f in SYNC_BRIDGE.glob("*")
            if f.suffix.lower() in VALID_EXTS and not f.name.startswith("_preview_")
        ])

        if not queue:
            ctx.warning("⚠️ Nessun file in coda da archiviare.")
        else:
            archived_count = 0
            for src in queue:
                try:
                    if src.suffix.lower() in IMAGE_EXTS:
                        # Converti immagine → PDF con effetto scanner
                        try:
                            img = Image.open(src)
                            if img.mode != "RGB":
                                img = img.convert("RGB")
                            img = apply_opencv_scanner(
                                img,
                                brightness=int(brightness),
                                contrast=float(contrast),
                                clahe_strength=float(clahe_strength),
                                white_background=bool(white_background),
                                auto_crop=bool(auto_crop),
                            )
                            img = apply_scanner_effect(img)
                            out_pdf = dest_root / (src.stem + ".pdf")
                            cnt = 1
                            while out_pdf.exists():
                                out_pdf = dest_root / f"{src.stem}_{cnt}.pdf"
                                cnt += 1
                            img.save(out_pdf, "PDF", resolution=100.0)
                            src.unlink(missing_ok=True)
                            ctx.success(f"✅ Archiviato: `{out_pdf.name}`")
                            output_files.append(out_pdf)
                        except Exception as e:
                            ctx.error(f"❌ Errore conversione `{src.name}`: {e}")
                    else:
                        # PDF nativo → copia diretta
                        out_pdf = dest_root / src.name
                        cnt = 1
                        while out_pdf.exists():
                            out_pdf = dest_root / f"{src.stem}_{cnt}{src.suffix}"
                            cnt += 1
                        shutil.copy2(src, out_pdf)
                        src.unlink(missing_ok=True)
                        ctx.success(f"✅ Archiviato: `{out_pdf.name}`")
                        output_files.append(out_pdf)
                    archived_count += 1
                except Exception as e:
                    ctx.error(f"❌ Errore archivazione `{src.name}`: {e}")

            if archived_count:
                ctx.success(f"📦 {archived_count} file archiviati in `{dest_root}`")

    # ── 5. File di riepilogo ──────────────────────────────────────────────────
    is_running = _tunnel_proc is not None and _tunnel_proc.poll() is None
    tunnel_url = ""
    if TUNNEL_URL_FILE.exists():
        try:
            tunnel_url = TUNNEL_URL_FILE.read_text("utf-8").strip()
        except Exception:
            pass
    remaining = len([
        f for f in SYNC_BRIDGE.glob("*")
        if f.suffix.lower() in VALID_EXTS and not f.name.startswith("_preview_")
    ])

    info_file = out_dir / "stato_estrattore.txt"
    info_file.write_text(
        "ESTRATTORE FATTURE — STATO CORRENTE\n"
        "=====================================\n"
        f"Tunnel: {'ATTIVO' if is_running else 'INATTIVO'}\n"
        f"URL:    {tunnel_url or 'N/A'}\n"
        f"File in coda: {remaining}\n"
        f"Cartella archivio: {output_path or 'non configurata'}\n",
        encoding="utf-8",
    )
    output_files.append(info_file)
    return output_files

def check_public_health(url: str) -> bool:
    try:
        resp = requests.get(f"{url.rstrip('/')}/health", timeout=5)
        return resp.status_code == 200 and "Toolbox" in resp.text
    except: return False

def get_ui_top():
    if "srv" not in st.session_state: st.session_state.srv = None
    if "tunnel_url" not in st.session_state: st.session_state.tunnel_url = None
    if "tunnel_verified" not in st.session_state: st.session_state.tunnel_verified = False

    with st.container():
        st.subheader("🔄 Sincronizzazione Cloudflare Tunnel")
        
        is_running = st.session_state.srv is not None and st.session_state.srv.poll() is None
        
        c1, c2 = st.columns([1, 1])

        # START
        if c1.button("▶ Avvia Sync", disabled=is_running, use_container_width=True, key="start"):
            token = base64.urlsafe_b64encode(os.urandom(12)).decode()
            env = os.environ.copy()
            env.update({"UP_TOKEN": token})
            
            args = [sys.executable, __file__, "--server", "--port", str(PORT)]
            st.session_state.srv = subprocess.Popen(args, env=env, creationflags=CREATE_NO_WINDOW)
            st.session_state.tunnel_url, st.session_state.tunnel_verified = None, False
            st.rerun()

        # STOP
        if c2.button("⏹ Stop / Reset", disabled=not is_running, use_container_width=True, key="stop"):
            if st.session_state.srv: st.session_state.srv.terminate()
            st.session_state.srv = None
            kill_cloudflared_processes()  # Pulisci tutti i tunnel cloudflared orfani
            kill_port_win(PORT)
            st.rerun()
        
        # KILL ZOMBIE (sempre attivo)
        c3 = st.columns([1])[0]
        if c3.button("🔥 Kill Zombie", use_container_width=True, key="kill_zombie", help="Termina tutti i processi cloudflared orfani"):
            kill_cloudflared_processes()
            kill_port_win(PORT)
            st.success("✅ Processi zombie terminati!")
            time.sleep(1)
            st.rerun()

        if is_running:
            st.info("🚀 Cloudflare Tunnel attivo. In attesa di assegnazione URL...")
            
            if not st.session_state.tunnel_verified:
                found = None
                if TUNNEL_URL_FILE.exists():
                    try: 
                        t = TUNNEL_URL_FILE.read_text("utf-8").strip()
                        if URL_RE_CF.match(t): found = t
                    except: pass
                
                if found:
                    # TRUST MODE: Se SSH ha scritto l'URL, è valido.
                    st.session_state.tunnel_url = found
                    st.session_state.tunnel_verified = True
                    st.rerun()
                else:
                    st.caption("Connessione a Cloudflare... (richiede qualche secondo)")
                    time.sleep(2)
                    st.rerun()
            else:
                # SEMPRE rileggi l'URL dal file per evitare cache stantia
                current_url = st.session_state.tunnel_url
                if TUNNEL_URL_FILE.exists():
                    try:
                        fresh_url = TUNNEL_URL_FILE.read_text("utf-8").strip()
                        if URL_RE_CF.match(fresh_url): current_url = fresh_url
                    except: pass
                
                st.success("✅ **Connessione Stabile OK!**")
                col_qr, col_txt = st.columns([1, 2])
                
                qr = qrcode.make(current_url)
                buf = io.BytesIO(); qr.save(buf, format="PNG")
                col_qr.image(buf.getvalue(), width=150)
                
                # URL nascosto per evitare blocco da Harmony Endpoint
                col_txt.markdown("### 📱 Scansiona il QR col telefono")
                col_txt.info("Funziona su WiFi e 4G. L'URL è codificato nel QR.")

    st.markdown("---")
    
    # === CARICAMENTO MANUALE ===
    st.subheader("📤 Caricamento Manuale")
    
    # Inizializza chiave dinamica per reset uploader
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    with st.expander("Carica file dal PC (Fatture/Immagini)", expanded=True):
        up_files = st.file_uploader(
            "Trascina qui i file o clicca per selezionare", 
            accept_multiple_files=True,
            type=["pdf", "jpg", "jpeg", "png", "webp"],
            key=f"uploader_{st.session_state.uploader_key}" # Chiave dinamica
        )
        
        if up_files:
            saved_c = 0
            for uf in up_files:
                # Logica salvataggio con gestione duplicati base
                bn = Path(uf.name).stem
                ex = Path(uf.name).suffix
                t_path = SYNC_BRIDGE / uf.name
                cnt = 1
                while t_path.exists():
                    t_path = SYNC_BRIDGE / f"{bn}_{cnt}{ex}"
                    cnt += 1
                
                t_path.write_bytes(uf.getvalue())
                saved_c += 1
            
            if saved_c > 0:
                # Incrementa chiave per resettare uploader al prossimo rerun
                st.session_state.uploader_key += 1
                st.success(f"✅ {saved_c} file caricati! Aggiorno...")
                time.sleep(1)
                st.rerun()

    st.markdown("---")
    # DASHBOARD
    st.subheader("📂 Dashboard Gestionale")
    
    # === DIAGNOSTICA LIBRERIE ===
    with st.expander("⚙️ Diagnostica Sistema", expanded=False):
        col1, col2, col3 = st.columns(3)
        
        # Check OpenCV usando il lazy loader
        cv2 = _get_cv2()
        if cv2 is not None:
            col1.success(f"✅ OpenCV {cv2.__version__}")
            opencv_ok = True
        else:
            col1.error("❌ OpenCV non installato")
            opencv_ok = False
        
        # Check NumPy
        try:
            import numpy as np
            col2.success(f"✅ NumPy {np.__version__}")
        except ImportError:
            col2.error("❌ NumPy non installato")
        
        # Check PIL
        try:
            from PIL import __version__ as pil_ver
            col3.success(f"✅ Pillow {pil_ver}")
        except:
            col3.warning("⚠️ Pillow (versione sconosciuta)")
        
        if opencv_ok:
            st.info("📸 Pipeline OpenCV attiva: rilevamento bordi + CLAHE + sfondo bianco")
        else:
            st.warning("📸 Pipeline base attiva: solo sharpening")
            st.caption(f"Cartella locale: `{OPENCV_LOCAL_DIR}`")
            
            col_inst1, col_inst2 = st.columns(2)
            if col_inst1.button("📦 Installa OpenCV", key="install_opencv", use_container_width=True):
                with st.spinner("Installazione in corso... (può richiedere 1-2 minuti)"):
                    success, output = install_opencv_locally()
                    if success:
                        st.success("✅ OpenCV installato! Ricarica la pagina.")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("❌ Installazione fallita:")
                        st.code(output)
            
            if col_inst2.button("🗑️ Forza Reinstallazione", key="force_reinstall_opencv", use_container_width=True):
                 with st.spinner("Rimozione e reinstallazione..."):
                    try:
                        import shutil
                        if OPENCV_LOCAL_DIR.exists():
                            shutil.rmtree(OPENCV_LOCAL_DIR)
                        success, output = install_opencv_locally()
                        if success:
                            st.success("✅ Reinstallazione completata! Ricarica la pagina.")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("❌ Reinstallazione fallita:")
                            st.code(output)
                    except Exception as e:
                        st.error(f"Errore rimozione: {e}")
    
    cfg = load_config()
    out_path = st.text_input("📁 Cartella Destinazione", value=cfg["output_path"])
    
    # Pulsanti Salva Percorso + Apri Cartella - SEMPRE VISIBILI
    col_save, col_open = st.columns([1, 1])
    if col_save.button("💾 Salva Percorso", key="save_path", use_container_width=True):
        cfg["output_path"] = out_path
        save_config(cfg)
        st.success("✅ Percorso salvato!")
    if col_open.button("📂 Apri Cartella", key="open_folder", use_container_width=True):
        folder = Path(out_path)
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(f'explorer "{folder}"', shell=True)
    
    
    
    cache = st.session_state.setdefault("cache", {})
    
    # Pulizia cache: rimuovi entry i cui file non esistono più
    stale_keys = [k for k, v in list(cache.items()) if not Path(v.current_path).exists()]
    for k in stale_keys:
        del cache[k]
    
    files = list(SYNC_BRIDGE.glob("*"))
    for f in files:
        if f.suffix.lower() not in VALID_EXTS: continue
        if f.name.startswith("_preview_"): continue  # Ignora file di anteprima
        h = get_file_hash(f)
        if h not in cache:
            target = f
            preview = str(f)
            is_conv = False
            if f.suffix.lower() in IMAGE_EXTS:
                target, preview_path = convert_to_pdf(f)
                h = get_file_hash(target)
                preview = str(preview_path)
                is_conv = True
            cache[h] = ManualFileState(f.name, h, str(target), preview_path=preview, is_converted=is_conv)

    if not cache:
        st.info("Nessuna fattura caricata.")
        if st.button("🔄 Aggiorna Lista", key="refresh_empty"):
            st.rerun()
        return

    # Pulsanti azione
    c_ref, c_clear = st.columns([1, 1])
    if c_ref.button("🔄 Aggiorna Lista", key="refresh_list", use_container_width=True):
        st.rerun()
    if c_clear.button("🧹 Svuota Dashboard", key="clear_all", use_container_width=True):
        for f in SYNC_BRIDGE.glob("*"): f.unlink()
        st.session_state.cache = {}
        st.rerun()
    

    # === FILE DA PROCESSARE ===
    pending = [(h, s) for h, s in cache.items() if not s.processed]
    completed = [(h, s) for h, s in cache.items() if s.processed]
    
    # CSS per bordo più spesso e rosso
    st.markdown("""
    <style>
    div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stVerticalBlockBorderWrapper"]) > div {
        border: 4px solid #dc3545 !important;
        border-radius: 10px !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    for h, state in pending:
        with st.container(border=True):
            # Titolo + Anteprima espandibile
            col_title, col_preview = st.columns([3, 1])
            col_title.write(f"📄 **{state.source_filename}**")
            
            # Anteprima - usa preview_path (immagine originale salvata)
            preview_file = Path(state.preview_path) if state.preview_path else Path(state.current_path)
            with st.expander("👁️ Anteprima & Modifica", expanded=False):
                if preview_file.exists() and preview_file.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}:
                    # Mostra immagine con correzione EXIF
                    try:
                        img = Image.open(preview_file)
                        img = ImageOps.exif_transpose(img)
                        
                        # Mini-editor controls
                        st.markdown("**🔧 Regolazioni Immagine**")
                        
                        # Inizializza stato per questo file
                        rot_key = f"rot_{h}"
                        bright_key = f"bright_{h}"
                        contrast_key = f"contrast_{h}"
                        clahe_key = f"clahe_{h}"
                        whitebg_key = f"whitebg_{h}"
                        autocrop_key = f"autocrop_{h}"
                        
                        if rot_key not in st.session_state:
                            st.session_state[rot_key] = 0
                        if bright_key not in st.session_state:
                            st.session_state[bright_key] = 35  # Default utente: 35
                        if contrast_key not in st.session_state:
                            st.session_state[contrast_key] = 1.4  # Default utente: 1.4
                        if clahe_key not in st.session_state:
                            st.session_state[clahe_key] = 0.0  # Default utente: 0.0 (disabilitato)
                        if whitebg_key not in st.session_state:
                            st.session_state[whitebg_key] = True
                        if autocrop_key not in st.session_state:
                            st.session_state[autocrop_key] = True
                        
                        # Rotazione
                        st.markdown("**Rotazione**")
                        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
                        if col_r1.button("↺ -90°", key=f"rotl_{h}", use_container_width=True):
                            st.session_state[rot_key] = (st.session_state[rot_key] - 90) % 360
                        if col_r2.button("↻ +90°", key=f"rotr_{h}", use_container_width=True):
                            st.session_state[rot_key] = (st.session_state[rot_key] + 90) % 360
                        if col_r3.button("⟳ 180°", key=f"rot180_{h}", use_container_width=True):
                            st.session_state[rot_key] = (st.session_state[rot_key] + 180) % 360
                        if col_r4.button("↩ Reset", key=f"rotreset_{h}", use_container_width=True):
                            st.session_state[rot_key] = 0
                        
                        current_rot = st.session_state[rot_key]
                        
                        # OpenCV Controls
                        st.markdown("**Parametri OpenCV**")
                        col_oct1, col_oct2 = st.columns(2)
                        st.session_state[autocrop_key] = col_oct1.checkbox(
                            "🔲 Auto-Crop", value=st.session_state[autocrop_key], key=f"cbac_{h}"
                        )
                        st.session_state[whitebg_key] = col_oct2.checkbox(
                            "⬜ Sfondo Bianco", value=st.session_state[whitebg_key], key=f"cbwb_{h}"
                        )
                        
                        st.session_state[bright_key] = st.slider(
                            "☀️ Luminosità", -50, 50, st.session_state[bright_key], 
                            key=f"sl_bright_{h}"
                        )
                        st.session_state[contrast_key] = st.slider(
                            "🌗 Contrasto", 0.5, 2.0, st.session_state[contrast_key], 
                            step=0.1, key=f"sl_contrast_{h}"
                        )
                        st.session_state[clahe_key] = st.slider(
                            "📊 CLAHE (contrasto locale)", 0.0, 5.0, st.session_state[clahe_key], 
                            step=0.5, key=f"sl_clahe_{h}"
                        )
                        
                        # Anteprima LIVE - applica OpenCV con tutti i parametri
                        try:
                            orig_img = Image.open(preview_file)
                            orig_img = ImageOps.exif_transpose(orig_img)
                            if orig_img.mode != 'RGB':
                                orig_img = orig_img.convert('RGB')
                            
                            # Applica OpenCV con tutti i parametri correnti
                            processed = apply_opencv_scanner(
                                orig_img,
                                rotation_override=current_rot,
                                brightness=st.session_state[bright_key],
                                contrast=st.session_state[contrast_key],
                                clahe_strength=st.session_state[clahe_key],
                                white_background=st.session_state[whitebg_key],
                                auto_crop=st.session_state[autocrop_key]
                            )
                            
                            st.image(processed, use_container_width=True)
                            st.caption("✨ Anteprima live con OpenCV")
                        except Exception as e:
                            # Fallback: mostra solo rotazione
                            if current_rot:
                                img = img.rotate(-current_rot, expand=True)
                            st.image(img, use_container_width=True)
                            st.caption(f"Rotazione: {current_rot}°")
                        
                    except Exception as e:
                        st.error(f"Errore caricamento: {e}")
                        st.image(str(preview_file), use_container_width=True)
                elif Path(state.current_path).suffix.lower() == ".pdf" and not state.is_converted:
                    # PDF nativo (non convertito) - mostra download
                    try:
                        with open(state.current_path, "rb") as f:
                            pdf_bytes = f.read()
                        st.download_button(
                            "📥 Scarica PDF", 
                            data=pdf_bytes, 
                            file_name=Path(state.current_path).name,
                            mime="application/pdf",
                            key=f"dl_{h}"
                        )
                        st.caption("PDF nativo - clicca per scaricare.")
                    except Exception as e:
                        st.error(f"Errore: {e}")
                else:
                    st.warning("Anteprima non disponibile.")
            
            # Riga 1: Nome e Descrizione (pre-popolati dal nome file)
            c_name, c_desc = st.columns([1, 1])
            state.name_prefix = c_name.text_input("Nome Fornitore", value=state.name_prefix, key=f"n_{h}")
            state.description = c_desc.text_input("Descrizione", value=state.description, key=f"d_{h}")
            
            # Riga 2: Calcolo IVA
            c_imp, c_iva, c_tot = st.columns([1, 0.5, 1])
            state.imponibile = c_imp.number_input("Imponibile (€)", min_value=0.0, step=0.01, key=f"imp_{h}", format="%.2f")
            state.iva_percent = c_iva.number_input("IVA %", min_value=0.0, max_value=100.0, value=22.0, step=1.0, key=f"iva_{h}")
            
            # Calcolo automatico totale
            iva_amount = state.imponibile * (state.iva_percent / 100)
            totale = state.imponibile + iva_amount
            c_tot.metric("Totale (€)", f"{totale:.2f}", delta=f"IVA: {iva_amount:.2f}")
            
            # Pulsante Archivia
            if st.button(f"🚀 Archivia", key=f"exec_{h}", use_container_width=True):
                # Costruisci nome file: "NomeFornitore - Descrizione.pdf" (SENZA PREZZO)
                nome_file = f"{state.name_prefix} - {state.description}"
                # if totale > 0: nome_file += f" - €{totale:.2f}" # RIMOSSO SU RICHIESTA UTENTE
                dest = Path(out_path) / f"{nome_file}.pdf"
                dest.parent.mkdir(parents=True, exist_ok=True)
                
                # Recupera impostazioni OpenCV da session_state
                rot_key = f"rot_{h}"
                bright_key = f"bright_{h}"
                contrast_key = f"contrast_{h}"
                clahe_key = f"clahe_{h}"
                whitebg_key = f"whitebg_{h}"
                autocrop_key = f"autocrop_{h}"
                
                rotation = st.session_state.get(rot_key, 0)
                brightness = st.session_state.get(bright_key, 0)
                contrast = st.session_state.get(contrast_key, 1.0)
                clahe_strength = st.session_state.get(clahe_key, 3.0)
                white_bg = st.session_state.get(whitebg_key, True)
                auto_crop = st.session_state.get(autocrop_key, True)
                
                # Se ci sono impostazioni personalizzate, ri-elabora l'immagine
                preview_file = Path(state.preview_path) if state.preview_path else None
                
                if preview_file and preview_file.exists() and preview_file.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}:
                    try:
                        orig_img = Image.open(preview_file)
                        orig_img = ImageOps.exif_transpose(orig_img)
                        if orig_img.mode != 'RGB':
                            orig_img = orig_img.convert('RGB')
                        
                        # Applica scanner con tutte le impostazioni
                        processed = apply_opencv_scanner(
                            orig_img, 
                            rotation_override=rotation,
                            brightness=brightness,
                            contrast=contrast,
                            clahe_strength=clahe_strength,
                            white_background=white_bg,
                            auto_crop=auto_crop
                        )
                        processed = apply_scanner_effect(processed)
                        
                        # Salva PDF
                        processed.save(dest, "PDF", resolution=100.0)
                        log_event(f"Archiviato con OpenCV: {nome_file}.pdf (rot={rotation}, bright={brightness}, contrast={contrast})")
                    except Exception as e:
                        log_event(f"Errore OpenCV, uso file originale: {e}")
                        shutil.copy2(state.current_path, dest)
                else:
                    # PDF nativo o file non immagine
                    shutil.copy2(state.current_path, dest)
                
                state.processed = True
                state.archived_path = str(dest)
                st.rerun()
    
    # === FILE COMPLETATI ===
    if completed:
        st.markdown("---")
        st.subheader("✅ Archiviati")
        for idx, (h, state) in enumerate(completed):
            with st.container(border=True):
                col_info, col_dl = st.columns([3, 1])
                col_info.write(f"✅ **{state.source_filename}** → `{state.name_prefix} - {state.description}`")
                
                # Pulsante download PDF archiviato
                archived = getattr(state, 'archived_path', state.current_path)
                if Path(archived).exists():
                    try:
                        with open(archived, "rb") as f:
                            pdf_bytes = f.read()
                        col_dl.download_button(
                            "📥 Scarica",
                            data=pdf_bytes,
                            file_name=Path(archived).name,
                            mime="application/pdf",
                            key=f"dlc_{idx}_{h[:8]}"
                        )
                    except: pass

if __name__ == "__main__":
    if "--server" in sys.argv:
        p = PORT
        if "--port" in sys.argv: p = int(sys.argv[sys.argv.index("--port")+1])
        if "run_server_process" in globals(): globals()["run_server_process"](p)
    else:
        if "st" in globals():
            try: st.set_page_config(page_title="SSH Sync", layout="wide")
            except: pass
            get_ui_top()
