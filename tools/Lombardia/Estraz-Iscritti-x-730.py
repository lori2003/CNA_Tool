from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
from collections import Counter
import pandas as pd
import re
import json



# --- Helpers Universali ---

MESI_MAP = {
    '01': 'Gennaio', '1': 'Gennaio',
    '02': 'Febbraio', '2': 'Febbraio',
    '03': 'Marzo', '3': 'Marzo',
    '04': 'Aprile', '4': 'Aprile',
    '05': 'Maggio', '5': 'Maggio',
    '06': 'Giugno', '6': 'Giugno',
    '07': 'Luglio', '7': 'Luglio',
    '08': 'Agosto', '8': 'Agosto',
    '09': 'Settembre', '9': 'Settembre',
    '10': 'Ottobre',
    '11': 'Novembre',
    '12': 'Dicembre'
}
MESI_LIST = list(sorted(set(MESI_MAP.values()), key=lambda x: [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", 
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
].index(x)))

TOOL = {
    "id": "Estraz-Iscritti-x-730",
    "name": "Estrazione Revoche (Multi-Anno Dinamico)",
    "description": (
        "#### 📌 1. FINALITÀ DEL TOOL\n"
        "Analizza i flussi mensili INPS della Lombardia per identificare ed estrarre le 'Revoche Reali', "
        "distinguendole dalle variazioni tecniche o baseline di partenza.\n\n"
        "#### 🚀 2. COME UTILIZZARLO\n"
        "1. **Sorgente:** Carica file TXT manualmente o seleziona una cartella locale PC per scansione massiva.\n"
        "2. **Filtri:** Imposta i prefissi Sede INPS (es. 24 per Milano) e il Mese di Partenza (Baseline).\n"
        "3. **Analisi:** Il tool confronta i file cronologicamente per rilevare la scomparsa di Codici Fiscali attivi.\n\n"
        "#### 🧠 3. LOGICA DI ELABORAZIONE (SPECIFICHE)\n"
        "* **Multi-Anno Dinamico:** Rileva l'Anno Fiscale leggendo il contenuto del file (record tipo 1) superando eventuali errori nel nome del file.\n"
        "* **Motore di Confronto (Pool):** Implementa una memoria di stato che traccia la presenza dei CF tra i diversi mesi; la revoca scatta solo se un CF scompare in un mese e riappare come 'Codice Funzione 1' (Revoca).\n"
        "* **Decodifica Date:** Estrae la data di Decorrenza Storica dal record INPS formattandola in formato italiano standard.\n\n"
        "#### 📂 4. RISULTATO FINALE\n"
        "Report Excel multi-scheda con Riepilogo Generale, Dettaglio per Anno e Log di Audit completo."
    ),
    "inputs": [
        {
            "key": "files_txt",
            "label": "Opzione 1: Trascina qui i file TXT",
            "type": "txt_multi",
            "required": False,
            "note": "Usa questa sezione se preferisci caricare i file manualmente."
        }
    ],
    "params": [
        {
            "key": "folder_source",
            "label": "Opzione 2: Oppure seleziona una Cartella Locale",
            "type": "folder",
            "default": "",
            "help": "Seleziona una o più cartelle sul tuo PC che contengono i file .txt. Clicca più volte l'icona 📂 per aggiungerne diverse.",
            "required": False,
            "section": "Sorgente Dati"
        },
        {
            "key": "sedi_inps",
            "label": "Filtro Sede INPS (Inizia con)",
            "type": "text",
            "default": "24, 87", 
            "help": "Scrivi i prefissi da includere separati da virgola. (Si salva in automatico quando esegui)",
            "required": True,
            "section": "Filtri"
        },
        {
            "key": "start_token",
            "label": "Mese di Partenza (Baseline per Confronto)",
            "type": "select",
            "options": MESI_LIST,
            "default": "Gennaio", 
            "help": "Seleziona il mese base da cui partire per calcolare le revoche. (Si salva in automatico)",
            "required": True,
            "section": "Filtri"
        },
        {
            "key": "filename_excel",
            "label": "Nome file Excel Output",
            "type": "text",
            "default": "Revoche_Estratte.xlsx",
            "required": True,
            "section": "Output"
        }
    ]
}

def _key_safe(s: str) -> str:
    """Replica la logica di app.py per generare le chiavi univoche."""
    return re.sub(r"[^0-9A-Za-z_]+", "_", str(s))[:180]

def parse_month_from_name(name: str) -> Tuple[int, str]:
    name_clean = name.lower().replace(".txt", "")
    name_clean = re.sub(r"v\d+", "", name_clean)

    for m_num, m_name in MESI_MAP.items():
        if m_name.lower() in name_clean:
            return int(m_num), m_name
    
    match_year_month = re.search(r"(?:19|20)\d{2}[-_\.]?(\d{1,2})", name_clean)
    if match_year_month:
        try:
            m = int(match_year_month.group(1))
            if 1 <= m <= 12:
                return m, MESI_MAP.get(f"{m:02d}", "Sconosciuto")
        except:
            pass

    nums = re.findall(r"(?<!\d)\d{1,2}(?!\d)", name_clean)
    candidates = [int(n) for n in nums if 1 <= int(n) <= 12]
    if candidates:
        m = candidates[-1]
        return m, MESI_MAP.get(f"{m:02d}", "Sconosciuto")
        
    return 0, "Sconosciuto"

def detect_year_from_content(file_obj_or_path: Any, max_lines=200) -> int:
    years = []
    lines_read = 0
    try:
        if isinstance(file_obj_or_path, Path):
            with open(file_obj_or_path, "r", encoding="latin-1", errors="replace") as f:
                for line in f:
                    if lines_read >= max_lines: break
                    if len(line) >= 306:
                        data_dec = line[298:306].strip()
                        if len(data_dec) == 8 and data_dec.isdigit():
                            years.append(data_dec[-4:])
                    lines_read += 1
        elif hasattr(file_obj_or_path, "read"):
            content = file_obj_or_path.read().decode("latin-1", errors="replace")
            for line in content.splitlines()[:max_lines]:
                if len(line) >= 306:
                    data_dec = line[298:306].strip()
                    if len(data_dec) == 8 and data_dec.isdigit():
                        years.append(data_dec[-4:])
            file_obj_or_path.seek(0)
    except Exception:
        pass
    if not years: return 0
    c = Counter(years)
    most_common = c.most_common(1)
    if most_common: return int(most_common[0][0])
    return 0

def organize_files(files_list: List[Any], is_preview: bool = False) -> Dict[int, List[Dict]]:
    grouped = {}
    for f in files_list:
        fname = f.name
        year = detect_year_from_content(f)
        if year == 0: year = 9999
        m_num, m_name = parse_month_from_name(fname)
        grouped.setdefault(year, []).append({
            "obj": f, "filename": fname, "month_num": m_num, "base_month_name": m_name, "display_name": m_name
        })
        
    final_groups = {}
    for y, items in grouped.items():
        items.sort(key=lambda x: x["month_num"])
        counts = {}
        for item in items:
            base = item["base_month_name"]
            if base == "Sconosciuto":
                item["display_name"] = f"Mese_Ignoto_({item['filename']})"
                continue
            if base not in counts:
                counts[base] = 1
                item["display_name"] = base
            else:
                counts[base] += 1
                item["display_name"] = f"{base}_{counts[base]}"
        final_groups[y] = items
    return final_groups

# --- Config Management ---

CONFIG_PATH = Path(__file__).parent / "config.json"

def load_config() -> Dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        import json
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(data: Dict[str, Any]) -> None:
    import json
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f, indent=4)

# --- Dynamic Params (Preview & Startup Loader) ---

def get_dynamic_params(uploads: Dict[str, Any], current_params: Dict[str, Any]) -> List[Dict[str, Any]]:
    from core.toolkit import ctx as st
    
    # 1. STARTUP LOADER: Inietta i valori salvati nel Session State
    try:
        tool_id = TOOL["id"]
        init_key = f"_init_done_auto_{tool_id}"
        
        if not st.session_state.get(init_key, False):
            config = load_config()
            selected_uid = st.session_state.get("selected_tool_uid", tool_id)
            
            for p_key in ["sedi_inps", "start_token"]:
                val_saved = config.get(p_key)
                if not val_saved: continue
                
                # SPECIAL HANDLING: Mese di Partenza
                if p_key == "start_token":
                    val_str = str(val_saved).strip()
                    if val_str in MESI_MAP:
                        val_saved = MESI_MAP[val_str]
                    elif val_str.capitalize() in MESI_LIST:
                        val_saved = val_str.capitalize()
                    else:
                        val_saved = "Gennaio"

                safe_uid = _key_safe(selected_uid)
                safe_key = _key_safe(p_key)
                full_key = f"param_{safe_uid}_{safe_key}"
                st.session_state[full_key] = val_saved
                current_params[p_key] = val_saved 
            
            st.session_state[init_key] = True
            
    except Exception:
        pass

    # 2. ANTEPRIMA (Riquadro Verde)
    # Usiamo un dizionario per mappare il nome file al suo oggetto/percorso, 
    # evitando duplicati basati sul percorso assoluto o oggetto.
    unique_files_map = {} # path_str -> file_obj
    
    # A) Upload manuali
    manual_uploads = uploads.get("files_txt", [])
    for f in manual_uploads:
        # Per i file caricati da Streamlit, usiamo il nome come chiave univoca 
        # (Streamlit garantisce l'unicità nell'upload corrente)
        unique_files_map[f"upload://{f.name}"] = f

    # B) Integrazione cartelle locali
    folder_path_raw = current_params.get("folder_source", "").strip()
    if folder_path_raw:
        paths = [p.strip() for p in folder_path_raw.splitlines() if p.strip()]
        for p_str in paths:
            fdir = Path(p_str)
            if fdir.exists() and fdir.is_dir():
                folder_files = list(fdir.glob("*.txt")) + list(fdir.glob("*.TXT"))
                for ff in folder_files:
                    unique_files_map[str(ff.absolute())] = ff

    files = list(unique_files_map.values())
    if not files:
        return []
        
    files_sig = []
    for f in files:
        sz = f.size if hasattr(f, "size") else (f.stat().st_size if hasattr(f, "stat") else 0)
        fname = f.name if hasattr(f, "name") else (f.name if isinstance(f, Path) else "unknown")
        files_sig.append(f"{fname}_{sz}")
    input_hash = hash(tuple(sorted(files_sig)))

    groups = organize_files(files, is_preview=True)
    
    config = load_config()
    live_token = current_params.get("start_token")
    if live_token is None:
        live_token = config.get("start_token", "01")
    
    start_tok = str(live_token).strip().lower()

    report_lines = []
    sorted_years = sorted(groups.keys())
    for y in sorted_years:
        y_label = str(y) if y != 9999 else "ANNO NON RILEVATO"
        report_lines.append(f"📅 ANNO: {y_label}")
        
        items = groups[y]
        baseline_file = None
        if items:
            if start_tok:
                for item in items:
                    if (start_tok in item["filename"].lower()) or \
                       (start_tok in item["display_name"].lower()) or \
                       (start_tok in item["base_month_name"].lower()):
                        baseline_file = item
                        break
            if not baseline_file:
                baseline_file = items[0]

        for item in items:
            marker = ""
            if baseline_file and item is baseline_file:
                marker = "  ⬅️ FILE DI PARTENZA"
            report_lines.append(f"   • {item['filename']}  ->  {item['display_name']}{marker}")
        report_lines.append("")
        
    report = "\n".join(report_lines)
    
    # Calcolo altezza dinamica (min 100, max 600)
    line_count = len(report_lines)
    dyn_height = min(max(100, line_count * 22), 600)

    return [
        {
            "key": f"preview_report_{input_hash}",
            "section": "Anteprima Raggruppamento",
            "label": f"Report Analisi ({len(files)} file rilevati):",
            "type": "textarea",
            "default": report,
            "height": dyn_height,
            "required": False
        }
    ]

# --- Runner ---

def parse_record_line(line: str) -> Optional[Dict[str, str]]:
    if len(line) < 306: return None
    try:
        return {
            "CodiceFunzione": line[6:7],
            "CodiceFiscale": line[282:298].strip(),
            "DataDecorrenza": line[298:306].strip(),
            "SiglaProvincia": line[270:272].strip(),
            "SedeInps": line[2:6].strip(),
            "SedeGestioneInps": line[344:348].strip() if len(line) >= 348 else ""
        }
    except:
        return None

def read_text_any(file_obj_or_path: Any) -> str:
    """Legge il contenuto come testo supportando sia Path che oggetti di upload (Streamlit)."""
    try:
        if hasattr(file_obj_or_path, "read"):
            # Caso Streamlit UploadedFile
            content = file_obj_or_path.read().decode("latin-1", errors="replace")
            file_obj_or_path.seek(0)
            return content
        elif isinstance(file_obj_or_path, Path):
            # Caso Path locale
            return file_obj_or_path.read_text(encoding="latin-1", errors="replace")
    except Exception:
        pass
    return ""

def run(out_dir: Path, **kwargs) -> List[Path]:
    # --- AUTO-SAVE ON RUN (Persistence) ---
    sedi_val = kwargs.get("sedi_inps", "24, 87")
    start_token_val = kwargs.get("start_token", "Gennaio")
    filename_excel = kwargs.get("filename_excel", "Revoche_Estratte.xlsx")
    folder_source = kwargs.get("folder_source", "").strip()
    
    # Rilevamento univoco dei file
    unique_files_map = {}
    manual_files = kwargs.get("files_txt", [])
    for f in manual_files:
        unique_files_map[f"upload://{f.name}"] = f

    if folder_source:
        paths = [p.strip() for p in folder_source.splitlines() if p.strip()]
        for p_str in paths:
            fdir = Path(p_str)
            if fdir.exists() and fdir.is_dir():
                for ff in (list(fdir.glob("*.txt")) + list(fdir.glob("*.TXT"))):
                    unique_files_map[str(ff.absolute())] = ff
    
    files = list(unique_files_map.values())

    try:
        config = load_config()
        config["sedi_inps"] = sedi_val
        config["start_token"] = start_token_val
        config["folder_source"] = folder_source
        save_config(config)
    except Exception as e:
        print(f"Salvataggio config fallito: {e}")
    # --------------------------------------

    if not files: return []
        
    sedi_prefixes = [s.strip() for s in str(sedi_val).split(",") if s.strip()]
    sedi_nums = set()
    for s in sedi_prefixes:
        try: sedi_nums.add(int(s))
        except: pass
    
    if not filename_excel.endswith(".xlsx"): filename_excel += ".xlsx"

    # 1. Appiattimento e Ordinamento Globale
    groups = organize_files(files)
    all_items = []
    for y in groups:
        for item in groups[y]:
            item["year_label"] = y
            all_items.append(item)
    
    all_items.sort(key=lambda x: (x["year_label"], x["month_num"], x["filename"]))
    
    # 2. Ricerca punto di inizio
    start_idx = 0
    start_token = str(start_token_val).strip().lower()
    if start_token:
        for i, itm in enumerate(all_items):
            if (start_token in itm["filename"].lower()) or \
               (itm["display_name"] and start_token in itm["display_name"].lower()) or \
               (itm["base_month_name"] and start_token in itm["base_month_name"].lower()):
                start_idx = i
                break
    
    analysis_items = all_items[start_idx:]
    if not analysis_items: return []

    # 3. Loop di Analisi Unificato con Monitor di Validazione
    active_pool = {}           # CF -> {month, file, year, sede, prov, data}
    revocations_by_year = {}   # Year -> list of Revocations
    debug_discards = []        # CF, month, file, sede, motivo
    source_frames = {}         # MonthKey -> DataFrame (Per verifica manuale)

    stats = {
        "File Analizzati": len(analysis_items),
        "Revoche Potenziali (Codice 1)": 0,
        "Revoche in Baseline (Salto Tecnico)": 0,
        "Revoche Confermate (Identificate)": 0,
        "Scartate (Senza storico anno)": 0,
        "Nuovi Iscritti Rilevati": 0
    }

    last_year = None

    for idx, item in enumerate(analysis_items):
        y_group = item["year_label"]
        
        # --- ISOLAMENTO ANNUALE ---
        if last_year is not None and y_group != last_year:
            active_pool = {} 
        last_year = y_group
        
        current_month_name = item["display_name"]
        current_filename = item["filename"]
        
        try:
            content = read_text_any(item["obj"])
            lines = content.splitlines()
            
            active_lines = []
            revoke_lines = []
            all_source_rows = [] # Per dump verifica
            
            for line in lines:
                if len(line) < 306: continue
                s_inps_raw = line[2:6].strip()
                s_gest_raw = line[344:348].strip() if len(line) >= 348 else ""
                code = line[6:7]
                cf_raw = line[282:298].strip()
                dt_raw = line[298:306].strip()
                prov_raw = line[270:272].strip()

                # Prepariamo dati per il foglio di verifica manuale
                all_source_rows.append({
                    "Sede": s_inps_raw, "Sede_Gest": s_gest_raw, "Cod": code, 
                    "CF": cf_raw, "Data": dt_raw, "Prov": prov_raw
                })
                
                is_our_sede = any(s_inps_raw.startswith(p) for p in sedi_prefixes)
                if code == "1":
                    revoke_lines.append(line)
                elif is_our_sede:
                    active_lines.append(line)
            
            # Salviamo il frame nel dizionario globale per l'Excel finale
            source_frames[f"{current_month_name}_{y_group}"] = pd.DataFrame(all_source_rows)
            
            # --- CASO A: PRIMA GLI ATTIVI ---
            # Questo permette di beccare chi è attivo e revocato nello stesso file
            for line in active_lines:
                cf = line[282:298].strip()
                dt_raw = line[298:306].strip()
                is_valid_dt = len(dt_raw) == 8 and dt_raw.isdigit()
                
                if cf not in active_pool:
                    stats["Nuovi Iscritti Rilevati"] += 1
                    active_pool[cf] = {
                        "month": current_month_name, "file": current_filename,
                        "year": y_group, "sede": line[2:6].strip(),
                        "sede_gest": line[344:348].strip() if len(line) >= 348 else "",
                        "prov": line[270:272].strip(), "data": dt_raw if is_valid_dt else ""
                    }
                else:
                    if not active_pool[cf]["data"] and is_valid_dt:
                        active_pool[cf]["data"] = dt_raw

            # --- POI LE REVOCHE ---
            for idx_rev, line in enumerate(revoke_lines):
                stats["Revoche Potenziali (Codice 1)"] += 1
                if idx == 0: # File Baseline
                    stats["Revoche in Baseline (Salto Tecnico)"] += 1

                cf = line[282:298].strip()
                s_inps_rev = line[2:6].strip()
                s_gest_rev = line[344:348].strip() if len(line) >= 348 else ""
                
                if cf in active_pool:
                    orig = active_pool[cf]
                    
                    # LOGICA DI RECUPERO DATA
                    data_final = str(orig.get("data", "")).strip()
                    if not (len(data_final) == 8 and data_final.isdigit()):
                        alt_data = line[298:306].strip()
                        if len(alt_data) == 8 and alt_data.isdigit():
                            data_final = alt_data

                    # Marcatore Caso A (Stesso file)
                    marker = ""
                    if orig["file"] == current_filename:
                        marker = "[Transizione Rapida]"

                    rev = {
                        "Note": marker,
                        "Codice Fiscale": cf,
                        "Data Decorrenza": data_final, 
                        "Mese di Comparsa Iniziale": f"{orig['month']} ({orig['year']})",
                        "Mese Rilevamento Revoca": f"{current_month_name} ({y_group})",
                        "Sede INPS": s_inps_rev,
                        "Sede Gestione INPS": s_gest_rev,
                        "Provincia": line[270:272].strip(),
                        "File Revoca": current_filename,
                        "File Prima Apparizione": orig["file"]
                    }
                    revocations_by_year.setdefault(y_group, []).append(rev)
                    stats["Revoche Confermate (Identificate)"] += 1
                    del active_pool[cf]
                else:
                    stats["Scartate (Senza storico anno)"] += 1
                    if len(debug_discards) < 200: # Limite debug
                        debug_discards.append({
                            "Mese": current_month_name,
                            "File": current_filename,
                            "Codice Fiscale": cf,
                            "Sede": s_inps_rev,
                            "Motivo": "Revoca senza storico nel pool sedi"
                        })

            # B) CARICAMENTO/AGGIORNAMENTO POOL (Per i MESI FUTURI)
            for line in active_lines:
                cf = line[282:298].strip()
                dt_raw = line[298:306].strip()
                is_valid_dt = len(dt_raw) == 8 and dt_raw.isdigit()
                
                if cf not in active_pool:
                    stats["Nuovi Iscritti Rilevati"] += 1
                    active_pool[cf] = {
                        "month": current_month_name, "file": current_filename,
                        "year": y_group, "sede": line[2:6].strip(),
                        "sede_gest": line[344:348].strip() if len(line) >= 348 else "",
                        "prov": line[270:272].strip(), "data": dt_raw if is_valid_dt else ""
                    }
                else:
                    # Se il CF è già nel pool ma non aveva una data valida, proviamo a integrarla ora
                    if not active_pool[cf]["data"] and is_valid_dt:
                        active_pool[cf]["data"] = dt_raw
                    
        except Exception as e:
            print(f"Errore parsing {current_filename}: {e}")

    # 4. Generazione Excel (Risultati e Verifica Sorgenti)
    final_path = out_dir / filename_excel
    source_check_path = out_dir / "Verifica_Sorgenti_Dettaglio.xlsx"
    
    summary_data = []
    output_frames = {}
    
    for y in sorted(revocations_by_year.keys()):
        df_y = pd.DataFrame(revocations_by_year[y])
        cols = ["Note", "Codice Fiscale", "Data Decorrenza", "Mese di Comparsa Iniziale", "Mese Rilevamento Revoca", "Sede INPS", "Sede Gestione INPS", "Provincia", "File Revoca", "File Prima Apparizione"]
        df_y = df_y[cols]
        output_frames[y] = df_y
        summary_data.append({
            "Anno Rilevamento": y if y != 9999 else "Ignoto", 
            "Nuove Revoche Individuate": len(df_y)
        })
    
    if not summary_data:
        summary_data.append({"Info": "Nessuna revoca rilevata", "Conteggio": 0})
    
    summary_df = pd.DataFrame(summary_data)
    
    # --- CREAZIONE MONITOR PROFESSIONAL ---
    pot = stats["Revoche Potenziali (Codice 1)"]
    base = stats["Revoche in Baseline (Salto Tecnico)"]
    conf = stats["Revoche Confermate (Identificate)"]
    
    pot_target = pot - base
    perc = (conf / pot_target * 100) if pot_target > 0 else 0
    status = "ECCELLENTE" if perc > 85 else ("BUONO" if perc > 60 else "ATTENZIONE - Baseline scarna")
    
    monitor_rows = [
        {"Indicatore": "STATO QUALITA' ANALISI", "Valore": status},
        {"Indicatore": "Percentuale Identificazione (Extra-Baseline)", "Valore": f"{perc:.2f}%"},
        {"Indicatore": "----------------------------", "Valore": "---"},
        {"Indicatore": "File Analizzati Totali", "Valore": stats["File Analizzati"]},
        {"Indicatore": "Revoche Potenziali", "Valore": pot},
        {"Indicatore": "Revoche in Baseline (Scarto Tecnico)", "Valore": base},
        {"Indicatore": "Revoche Confermate (Nell'Excel)", "Valore": conf},
        {"Indicatore": "Revoche Scartate (Senza storico)", "Valore": stats["Scartate (Senza storico anno)"]},
        {"Indicatore": "Nuovi Iscritti Censiti", "Valore": stats["Nuovi Iscritti Rilevati"]},
        {"Indicatore": "Isolati a fine periodo (Pool)", "Valore": len(active_pool)}
    ]
    stats_df = pd.DataFrame(monitor_rows)

    # File 1: REVOCHE
    with pd.ExcelWriter(final_path, engine='openpyxl') as writer:
        stats_df.to_excel(writer, sheet_name="Monitor_Validazione", index=False)
        summary_df.to_excel(writer, sheet_name="Riepilogo", index=False)
        if debug_discards:
            df_debug = pd.DataFrame(debug_discards)
            df_debug.to_excel(writer, sheet_name="DEBUG_Scartati", index=False)
        
        for y in sorted(output_frames.keys()):
            df = output_frames[y].copy()
            sname = str(y) if y != 9999 else "Anno_Ignoto"
            if "Data Decorrenza" in df.columns:
                df["Data Decorrenza"] = df["Data Decorrenza"].apply(lambda d: f"{str(d)[0:2]}/{str(d)[2:4]}/{str(d)[4:8]}" if len(str(d)) == 8 and str(d).isdigit() else d)
            df.to_excel(writer, sheet_name=sname, index=False)
            
            # --- AUTO-ADATTAMENTO COLONNE (Solo per fogli anno) ---
            if sname.isdigit() and len(sname) == 4:
                ws = writer.sheets[sname]
                for col in ws.columns:
                    max_len = 0
                    column_letter = col[0].column_letter 
                    for cell in col:
                        try:
                            if cell.value:
                                length = len(str(cell.value))
                                if length > max_len: max_len = length
                        except:
                            pass
                    ws.column_dimensions[column_letter].width = max_len + 3

    # File 2: DETTAGLIO SORGENTI (Verifica)
    with pd.ExcelWriter(source_check_path, engine='openpyxl') as writer:
        for sname, df in source_frames.items():
            ws_name = sname.replace(" ", "_").replace("/", "-")[:31]
            df.to_excel(writer, sheet_name=ws_name, index=False)

    return [final_path, source_check_path]

