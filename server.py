"""
server.py — FastAPI entry point per Toolbox CNA
================================================
Avvia con:  python server.py
oppure:     uvicorn server:app --reload --port 8501
"""
from __future__ import annotations

import contextlib
import os
import sys
import webbrowser
from pathlib import Path

# Fix encoding su Windows (terminale cp1252 non supporta emoji)
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    with contextlib.suppress(Exception):
        sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    with contextlib.suppress(Exception):
        sys.stderr.reconfigure(encoding="utf-8")

# ── Path setup ────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "core"))

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from api.tools_routes import router as tools_router, init_tools

# ── Lifespan (sostituisce on_event deprecato) ─────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    root = Path(os.getenv("TOOLBOX_HOME", str(BASE_DIR))).resolve()
    init_tools(root)
    yield

# ── App ───────────────────────────────────────────────────────
app = FastAPI(title="Toolbox CNA", version="3.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── API routes ────────────────────────────────────────────────
app.include_router(tools_router, prefix="/api")

# ── Static files ──────────────────────────────────────────────
WEB_DIR = BASE_DIR / "web"
app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")

# ── SPA fallback ──────────────────────────────────────────────
@app.get("/")
@app.get("/{full_path:path}")
async def spa(full_path: str = ""):
    # Evita che la SPA intercetti api o statici se non sono stati già gestiti
    if full_path.startswith("api/") or full_path.startswith("static/"):
        from fastapi import HTTPException
        raise HTTPException(status_code=404, detail="Not Found")
    
    index = WEB_DIR / "index.html"
    if index.exists():
        return FileResponse(str(index))
    return {"error": "web/index.html non trovato"}


# ── Dev entrypoint ────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8501))
    print(f"\nToolbox CNA  ->  http://localhost:{port}\n")
    # Apri browser automaticamente
    import threading
    threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)
