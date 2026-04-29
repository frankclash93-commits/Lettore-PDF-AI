import os
import sys
import subprocess
import traceback
import time
import json
import re
import threading
import sqlite3
import webbrowser
import ast
import logging
from json import JSONDecodeError
from datetime import datetime, timedelta

# Forza l'encoding di stdout su utf-8 quando possibile (Python 3.7+)
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    os.environ.setdefault("PYTHONUTF8", "1")

# Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

def safe_print(msg: str):
    try:
        print(msg)
    except UnicodeEncodeError:
        filtered = msg.encode('utf-8', errors='replace').decode('ascii', errors='ignore')
        print(filtered)

# Optional auto-install (commentato per sicurezza in ambienti di produzione)
def install_dependencies():
    dependencies = ['customtkinter', 'pypdf', 'requests', 'psutil', 'ollama', 'win10toast']
    for lib in dependencies:
        try:
            __import__(lib)
        except ImportError:
            safe_print(f"📦 Libreria '{lib}' mancante. Installazione in corso...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
                safe_print(f"✅ '{lib}' installata con successo!")
            except Exception as e:
                safe_print(f"❌ Errore installazione '{lib}': {e}")

# install_dependencies()  # lasciare commentato se non vuoi auto-install

# Import UI libs
try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except Exception:
    CTK_AVAILABLE = False

import tkinter as tk
from tkinter import filedialog, messagebox
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import ollama
    OLLAMA_AVAILABLE = True
except Exception:
    OLLAMA_AVAILABLE = False

import requests
import psutil

# win10toast
try:
    from win10toast import ToastNotifier
    TOAST_AVAILABLE = True
except Exception:
    ToastNotifier = None
    TOAST_AVAILABLE = False

import shutil  # per which()

# ---------------------------
# Configurazione estetica
# ---------------------------
ACCENT = "#03caf6"
PANEL = "#182437"
BG = "#010610"

# ---------------------------
# Utilità: rilevamento Ollama e profilo HW
# ---------------------------
def detect_ollama():
    paths = [
        shutil.which("ollama"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Ollama\ollama.exe"),
        r"C:\Program Files\Ollama\ollama.exe"
    ]
    for p in paths:
        if p and os.path.exists(p):
            try:
                out = subprocess.check_output([p, "--version"], stderr=subprocess.DEVNULL)
                return {"installed": True, "path": p, "version": out.decode().strip()}
            except Exception:
                return {"installed": True, "path": p, "version": "unknown"}
    return {"installed": False}

def hardware_profile():
    ram = round(psutil.virtual_memory().total / (1024**3))
    cpu = psutil.cpu_count(logical=True)
    gpu = ""
    try:
        if os.name == 'nt':
            gpu = subprocess.check_output("wmic path win32_VideoController get name", shell=True).decode(errors='ignore')
    except Exception:
        gpu = ""
    return {"ram": ram, "cpu": cpu, "gpu": gpu}

def suggerisci_modello_smart(hw):
    ram = hw.get("ram", 8) or 8
    gpu = (hw.get("gpu") or "").upper()
    if "NVIDIA" in gpu or "AMD" in gpu:
        if ram >= 16:
            return "llama3.1:8b", "🔥 Prestazioni alte"
        elif ram >= 10:
            return "llama3", "⚡ Bilanciato"
        else:
            return "phi3:mini", "💡 Leggero GPU"
    else:
        if ram <= 8:
            return "tinyllama", "💻 Base"
        else:
            return "phi3:mini", "🧠 CPU smart"

def is_executable_available(cmd_name: str) -> bool:
    return shutil.which(cmd_name) is not None

def avvia_e_forza_gpu():
    try:
        safe_print("Resettando Ollama per attivare la GPU (se presente)...")
        if os.name == 'nt' and is_executable_available("ollama"):
            try:
                subprocess.run("taskkill /F /IM ollama.exe /T", shell=True,
                            stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
            except Exception:
                pass
        check_gpu = ""
        try:
            if os.name == 'nt':
                check_gpu = subprocess.check_output("wmic path win32_VideoController get name", shell=True).decode(errors='ignore')
            else:
                try:
                    check_gpu = subprocess.check_output("lspci", shell=True).decode(errors='ignore')
                except Exception:
                    check_gpu = ""
        except Exception:
            check_gpu = ""
        if "AMD" in check_gpu.upper() or "RADEON" in check_gpu.upper():
            safe_print("GPU AMD rilevata. Configurazione Vulkan...")
            os.environ["OLLAMA_VULKAN"] = "1"
            os.environ["HSA_OVERRIDE_GFX_VERSION"] = "11.0.0"
        os.environ.setdefault("OLLAMA_MODELS", os.path.expanduser("~/.ollama/models"))
        if is_executable_available("ollama"):
            safe_print("Lancio del server Ollama con accelerazione hardware (se disponibile)...")
            popen_kwargs = {"env": os.environ}
            if os.name == 'nt':
                try:
                    popen_kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE
                except Exception:
                    pass
            try:
                subprocess.Popen(["ollama", "serve"], **popen_kwargs)
            except Exception as e:
                safe_print(f"Errore avvio Ollama: {e}")
        else:
            safe_print("Ollama non trovato nel PATH. Saltando avvio server Ollama.")
        safe_print("Attesa inizializzazione (5 secondi)...")
        time.sleep(5)
    except Exception as e:
        safe_print(f"Errore durante l'avvio forzato: {e}")
        traceback.print_exc()

try:
    t_gpu = threading.Thread(target=avvia_e_forza_gpu, daemon=True)
    t_gpu.start()
except Exception as e:
    safe_print(f"Errore avvio GPU thread: {e}")

# ---------------------------
# SmartBridge (web fallback)
# ---------------------------
class SmartBridge:
    def __init__(self):
        self.session = requests.Session()

    def ask(self, prompt):
        url = "https://text.pollinations.ai/"
        payload = {
            "messages": [{"role": "user", "content": prompt}],
            "model": "openai",
            "jsonMode": False,
            "seed": 42
        }
        try:
            resp = self.session.post(url, json=payload, timeout=45)
            if resp.status_code == 200:
                testo = resp.text
                try:
                    data = json.loads(testo)
                    return data.get('content', testo)
                except Exception:
                    testo_pulito = re.sub(r'\{.*?"content":\s*"(.*?)"\}', r'\1', testo, flags=re.DOTALL)
                    return testo_pulito.strip()
            return f"⚠️ Errore Bridge (Codice {resp.status_code})"
        except Exception as e:
            return f"❌ Connessione fallita: {str(e)}"

# ---------------------------
# FatturaAutomation (kept)
# ---------------------------
class FatturaAutomation:
    def __init__(self, app):
        self.app = app
        self.watched = set()
        self.running = False
        self.lock = threading.Lock()

    def auto_scan_folder(self, path):
        try:
            files = [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith('.pdf')]
            for f in files:
                if f not in self.watched:
                    with self.lock:
                        self.watched.add(f)
                    try:
                        reader = PdfReader(f)
                        pages = [p.extract_text() or "" for p in reader.pages]
                        self.app.pdf_pages = pages
                        self.app.pdf_text = " ".join(pages)
                        threading.Thread(target=self.app._extract_logic, daemon=True).start()
                    except Exception as e:
                        safe_print(f"Errore lettura PDF watcher: {e}")
        except Exception as e:
            safe_print(f"Errore scan folder: {e}")

    def auto_registra(self):
        pass

    def auto_scadenze(self):
        pass

# ---------------------------
# Helper: estrai primo oggetto JSON bilanciato
# ---------------------------
def _extract_first_json_object_from_text(text: str) -> str | None:
    if not text:
        return None
    start = text.find('{')
    if start == -1:
        return None
    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(text)):
        ch = text[i]
        if ch == '"' and not escape:
            in_string = not in_string
        if ch == '\\' and not escape:
            escape = True
            continue
        else:
            escape = False
        if in_string:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return text[start:i+1]
    return None

# ---------------------------
# Parsing utilities (pulizia risposta AI, estrazione importi)
# ---------------------------
def _clean_raw_ai_response(raw: str) -> str:
    if not raw:
        return raw
    m = re.search(r"```json\s*(\{.*?\})\s*```", raw, flags=re.DOTALL | re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"```\s*(\{.*?\})\s*```", raw, flags=re.DOTALL)
    if m:
        return m.group(1).strip()
    m = re.search(r"Message\s*\([^)]*content\s*=\s*(['\"]{1,3})(.*?)\1", raw, flags=re.DOTALL)
    if m:
        inner = m.group(2)
        m2 = re.search(r"```json\s*(\{.*?\})\s*```", inner, flags=re.DOTALL | re.IGNORECASE)
        if m2:
            return m2.group(1).strip()
        if inner.strip().startswith("{"):
            return inner.strip()
    m = re.search(r"content\s*=\s*(['\"]{1,3})(.*?)\1", raw, flags=re.DOTALL)
    if m:
        inner = m.group(2)
        m2 = re.search(r"```json\s*(\{.*?\})\s*```", inner, flags=re.DOTALL | re.IGNORECASE)
        if m2:
            return m2.group(1).strip()
        if inner.strip().startswith("{"):
            return inner.strip()
    json_like = _extract_first_json_object_from_text(raw)
    if json_like:
        return json_like.strip()
    return raw.strip()

def _parse_amounts_from_text(text: str) -> dict:
    res = {"imponibile": None, "iva": None, "totale": None}
    if not text:
        return res
    t = text.replace('\xa0', ' ').replace('\r', ' ')
    m_tot = re.search(r"(?:Totale(?:\s+fattura)?|Totale da pagare|Importo totale|Lordo|TOTALE)\s*[:\-]?\s*€?\s*([0-9\.\,]+)", t, flags=re.IGNORECASE)
    if not m_tot:
        m_tot = re.search(r"€\s*([0-9\.\,]+)", t)
    if m_tot:
        try:
            tot_s = m_tot.group(1).replace('.', '').replace(',', '.')
            res["totale"] = float(tot_s)
        except:
            res["totale"] = None
    m_iva = re.search(r"IVA\s*(?:\(?[0-9]{1,2}%\)?)?\s*[:\-]?\s*€?\s*([0-9\.\,]+)", t, flags=re.IGNORECASE)
    if m_iva:
        try:
            iva_s = m_iva.group(1).replace('.', '').replace(',', '.')
            res["iva"] = float(iva_s)
        except:
            res["iva"] = None
    else:
        m_iva2 = re.search(r"([0-9]{1,2})\s*%\s*[^\d]{0,6}€\s*([0-9\.\,]+)", t)
        if m_iva2:
            try:
                res["iva"] = float(m_iva2.group(2).replace('.', '').replace(',', '.'))
            except:
                res["iva"] = None
    m_imp = re.search(r"(?:Imponibile|Netto)\s*[:\-]?\s*€?\s*([0-9\.\,]+)", t, flags=re.IGNORECASE)
    if m_imp:
        try:
            imp_s = m_imp.group(1).replace('.', '').replace(',', '.')
            res["imponibile"] = float(imp_s)
        except:
            res["imponibile"] = None
    if res["imponibile"] is None and res["totale"] is not None and res["iva"] is not None:
        try:
            res["imponibile"] = round(res["totale"] - res["iva"], 2)
        except:
            res["imponibile"] = None
    if res["iva"] is None and res["totale"] is not None and res["imponibile"] is not None:
        try:
            res["iva"] = round(res["totale"] - res["imponibile"], 2)
        except:
            res["iva"] = None
    return res

def _extract_invoice_items(text: str, max_items: int = 8) -> list:
    items = []
    if not text:
        return items
    lines = [l.strip() for l in text.replace('\xa0', ' ').splitlines() if l.strip()]
    for ln in lines:
        m = re.search(r"(.{3,200}?)\s+€\s*([0-9\.\,]+)(?!\S)", ln)
        if not m:
            m = re.search(r"(.{3,200}?)\s+([0-9\.\,]+)\s*€(?!\S)", ln)
        if m:
            descr = m.group(1).strip()
            amt_s = m.group(2).replace('.', '').replace(',', '.')
            try:
                amt = float(amt_s)
            except:
                amt = None
            items.append({"descrizione": descr, "importo": amt})
        if len(items) >= max_items:
            break
    if not items:
        for ln in lines[:max_items]:
            m = re.search(r"([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2}))", ln)
            if m:
                try:
                    amt = float(m.group(1).replace('.', '').replace(',', '.'))
                except:
                    amt = None
                items.append({"descrizione": ln[:80], "importo": amt})
    return items

def _clean_piva(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"\b([0-9]{11})\b", s)
    return m.group(1) if m else ""

def _format_invoice_summary(parsed: dict, items: list, piva_doc: str, piva_azienda: str) -> str:
    def fmt_euro(v):
        try:
            return f"€ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "€ 0,00"
    data_em = parsed.get("data_emissione", "")
    data_readable = data_em
    try:
        dt = datetime.strptime(data_em, "%d/%m/%Y")
        mesi = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno","Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
        data_readable = f"{dt.day} {mesi[dt.month-1]} {dt.year}"
    except Exception:
        pass
    lines = []
    numero = parsed.get("numero", parsed.get("num", parsed.get("n", "N. sconosciuto")))
    lines.append(f"Sulla fattura n. {numero} sono presenti i seguenti dati:\n")
    lines.append(f"* Data: {data_readable}")
    fornitore = parsed.get("fornitore") or parsed.get("emittente") or parsed.get("mittente") or parsed.get("entita", "")
    f_indirizzo = parsed.get("fornitore_indirizzo") or parsed.get("indirizzo_fornitore") or parsed.get("indirizzo", "")
    lines.append("* Dati del Fornitore:")
    lines.append(f" + Nome: {fornitore or 'Sconosciuto'}")
    if f_indirizzo:
        lines.append(f" + Indirizzo: {f_indirizzo}")
    if piva_doc:
        lines.append(f" + P.IVA/C.F.: {piva_doc}")
    else:
        lines.append(f" + P.IVA/C.F.: Non rilevata")
    cliente = parsed.get("cliente") or parsed.get("destinatario") or parsed.get("cliente_entita") or ""
    c_indirizzo = parsed.get("cliente_indirizzo") or parsed.get("indirizzo_cliente") or ""
    if cliente or c_indirizzo:
        lines.append("* Dati del Cliente:")
        lines.append(f" + Nome: {cliente or 'Sconosciuto'}")
        if c_indirizzo:
            lines.append(f" + Indirizzo: {c_indirizzo}")
    lines.append("* Dettaglio Prodotti / Servizi")
    lines.append(" + Descrizione Quantità Prezzo Unitario Totale")
    if items:
        for it in items:
            descr = it.get("descrizione", "").strip()
            amt = it.get("importo")
            qty = it.get("quantita", "")
            unit = it.get("prezzo_unitario", "")
            if amt is not None:
                amt_str = fmt_euro(amt)
                if qty and unit:
                    lines.append(f"   - {descr} {qty} {fmt_euro(unit)} {amt_str}")
                else:
                    lines.append(f"   - {descr} {amt_str}")
            else:
                lines.append(f"   - {descr}")
    else:
        lines.append(" + (Nessuna riga dettagliata rilevata)")
    lines.append("* Riepilogo Economico:")
    imponibile = parsed.get("importo_netto") or parsed.get("imponibile") or ""
    iva = parsed.get("iva") or ""
    totale = parsed.get("totale") or ""
    if imponibile:
        lines.append(f" + Imponibile: {fmt_euro(imponibile)}")
    if iva:
        aliquota = parsed.get("aliquota") or parsed.get("iva_percentuale") or ""
        if aliquota:
            lines.append(f" + IVA ({aliquota}%): {fmt_euro(iva)}")
        else:
            lines.append(f" + IVA: {fmt_euro(iva)}")
    if totale:
        lines.append(f" + TOTALE FATTURA: {fmt_euro(totale)}")
    lines.append("* Note e Pagamento:")
    metodo = parsed.get("metodo_pagamento") or parsed.get("pagamento") or parsed.get("modalita_pagamento") or ""
    iban = parsed.get("iban") or parsed.get("IBAN") or ""
    scadenza = ""
    try:
        gg = int(parsed.get("giorni_scadenza", 0))
        if gg:
            scadenza = f"{gg} giorni data fattura."
    except:
        scadenza = parsed.get("scadenza", "")
    if metodo:
        lines.append(f" + Metodo di pagamento: {metodo}")
    if iban:
        lines.append(f" + IBAN: {iban}")
    if scadenza:
        lines.append(f" + Scadenza: {scadenza}")
    return "\n".join(lines)

# ---------------------------
# APP PRINCIPALE
# ---------------------------
if CTK_AVAILABLE:
    BaseAppClass = ctk.CTk
else:
    BaseAppClass = tk.Tk

class ApexLedgerApp(BaseAppClass):
    def __init__(self):
        if CTK_AVAILABLE:
            super().__init__()
            self.title("APEX LEDGER | Hybrid Intelligence SRL Edition 2026")
            self.geometry("1280x800")
        else:
            super().__init__()
            self.title("APEX LEDGER (Tk fallback)")
            self.geometry("1000x700")

        self.bridge = SmartBridge()
        self.pdf_text = ""
        self.pdf_pages = []
        self.mode = "LOCALE"

        try:
            self.hw = hardware_profile()
            self.ollama_info = detect_ollama()
            self.modello_locale, self.diagnosi_testo = suggerisci_modello_smart(self.hw)
        except Exception:
            self.modello_locale, self.diagnosi_testo = suggerisci_modello_smart({"ram":8,"cpu":2,"gpu":""})
            self.hw = {"ram": None, "cpu": None, "gpu": None}
            self.ollama_info = {"installed": False}

        self.history = []

        self.init_db()
        self.setup_azienda_if_missing()

        self.setup_ui()
        self.auto_start_ollama()

        # inizializza automazioni
        self.automation = FatturaAutomation(self)

        try:
            if hasattr(self, 'chat_area'):
                if self.ollama_info.get("installed"):
                    self.chat_area.insert("end", f"\n🧠 Ollama rilevato: {self.ollama_info.get('version')}\n")
                else:
                    self.chat_area.insert("end", "\n⚠️ Ollama non trovato (opzionale)\n")
                self.chat_area.insert("end", f"💡 Modello consigliato: {self.modello_locale} ({self.diagnosi_testo})\n")
        except Exception:
            pass

    def init_db(self):
        conn = sqlite3.connect('apex_ledger_v3.db')
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS contabilita 
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                      tipo TEXT, 
                      fornitore_cliente TEXT, 
                      importo_netto REAL, 
                      iva REAL,
                      totale_fattura REAL,
                      data_emissione TEXT, 
                      data_scadenza TEXT,
                      stato_pagamento TEXT, 
                      note TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS azienda (
            id INTEGER PRIMARY KEY,
            nome TEXT,
            piva TEXT
        )''')
        conn.commit()
        conn.close()

    def get_azienda(self):
        conn = sqlite3.connect('apex_ledger_v3.db')
        c = conn.cursor()
        c.execute("SELECT nome, piva FROM azienda LIMIT 1")
        row = c.fetchone()
        conn.close()
        return row

    def setup_azienda(self):
        try:
            from tkinter import simpledialog
            nome = simpledialog.askstring("Setup Azienda", "Inserisci denominazione azienda:")
            piva = simpledialog.askstring("Setup Azienda", "Inserisci Partita IVA:")
            if not nome or not piva:
                return
            conn = sqlite3.connect('apex_ledger_v3.db')
            c = conn.cursor()
            c.execute("INSERT INTO azienda (nome, piva) VALUES (?, ?)", (nome, piva))
            conn.commit()
            conn.close()
            self.mia_azienda = nome
            self.mia_piva = piva
        except Exception as e:
            safe_print(f"Errore setup azienda: {e}")

    def setup_azienda_if_missing(self):
        azienda = self.get_azienda()
        if not azienda:
            try:
                if CTK_AVAILABLE or True:
                    self.setup_azienda()
            except Exception:
                pass
        else:
            self.mia_azienda, self.mia_piva = azienda

    def auto_start_ollama(self):
        if is_executable_available("ollama"):
            try:
                subprocess.Popen(["ollama", "serve"], 
                                 creationflags=(0x08000000 if os.name == 'nt' else 0),
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
        else:
            safe_print("Ollama non trovato: server non avviato automaticamente.")

    def setup_ui(self):
        if CTK_AVAILABLE:
            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self.sidebar = ctk.CTkFrame(self, width=300, fg_color=BG, corner_radius=0)
            self.sidebar.grid(row=0, column=0, sticky="nsew")

            ctk.CTkLabel(self.sidebar, text="APEX LEDGER", font=("Segoe UI", 20, "bold"), text_color=ACCENT).pack(pady=(20, 8))
            ctk.CTkLabel(self.sidebar, text="Hybrid System SRL v4.0", font=("Segoe UI", 10), text_color="gray").pack(pady=(0, 12))

            self.maint_frame = ctk.CTkFrame(self.sidebar, fg_color=PANEL, corner_radius=12)
            self.maint_frame.pack(fill="x", padx=15, pady=10)
            ctk.CTkLabel(self.maint_frame, text="DIAGNOSI HARDWARE", font=("Segoe UI", 11, "bold"), text_color=ACCENT).pack(pady=(10, 5))
            ctk.CTkLabel(self.maint_frame, text=self.diagnosi_testo, font=("Segoe UI", 10), wraplength=250).pack(pady=5)

            self.btn_download = ctk.CTkButton(self.maint_frame, text=f"SCARICA {self.modello_locale.upper()}", 
                                            fg_color="#059669", hover_color="#047857", height=35,
                                            command=lambda: safe_print("Use installa_modello_cmd if needed"))
            self.btn_download.pack(pady=10, padx=15)

            ctk.CTkLabel(self.sidebar, text="MODALITÀ AI", font=("Segoe UI", 12, "bold")).pack(pady=(10, 5))
            self.mode_var = tk.StringVar(value="LOCALE")
            self.seg_button = ctk.CTkSegmentedButton(self.sidebar, values=["LOCALE", "AI SMART"], 
                                                    command=self.set_mode, variable=self.mode_var)
            self.seg_button.pack(pady=5, padx=20, fill="x")

            ctk.CTkButton(self.sidebar, text="📂 CARICA FATTURA (PDF / TXT)", command=self.load_pdf).pack(pady=(20, 8), padx=20, fill="x")
            self.btn_analyze = ctk.CTkButton(self.sidebar, text="⚡ ESTRAI E REGISTRA", fg_color="#047857", state="disabled", command=self.analyze_invoice)
            self.btn_analyze.pack(pady=8, padx=20, fill="x")
            self.btn_bilancio = ctk.CTkButton(self.sidebar, text="📊 BILANCIO MENSILE", fg_color="#b45309", command=self.mostra_bilancio)
            self.btn_bilancio.pack(pady=8, padx=20, fill="x")
        else:
            left_frame = tk.Frame(self, width=300, bg=BG)
            left_frame.pack(side="left", fill="y")
            tk.Label(left_frame, text="APEX LEDGER", fg=ACCENT, bg=BG, font=("Segoe UI", 18, "bold")).pack(pady=(20,8))
            tk.Button(left_frame, text="CARICA FATTURA (PDF / TXT)", command=self.load_pdf).pack(pady=8, padx=10, fill="x")
            self.btn_analyze = tk.Button(left_frame, text="ESTRAI E REGISTRA", state="disabled", command=self.analyze_invoice)
            self.btn_analyze.pack(pady=8, padx=10, fill="x")
            tk.Button(left_frame, text="BILANCIO MENSILE", command=self.mostra_bilancio).pack(pady=8, padx=10, fill="x")

        # Tabview ERP minimale (se customtkinter disponibile)
        if CTK_AVAILABLE:
            self.tabview = ctk.CTkTabview(self, width=800)
            self.tabview.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
            self.tab_chat = self.tabview.add("AI")
            self.tab_erp = self.tabview.add("ERP")

            # Chat area nella tab AI
            self.chat_area = ctk.CTkTextbox(self.tabview.tab("AI"), font=("Segoe UI", 12), corner_radius=8, border_width=1, border_color="#70a7ff")
            self.chat_area.pack(fill="both", expand=True, padx=10, pady=10)
            self.chat_area.insert("end", f"SISTEMA: Pronto. Modello locale suggerito: {self.modello_locale}.\n")

            # ERP: treeview per fatture
            import tkinter.ttk as ttk
            self.tree = ttk.Treeview(self.tabview.tab("ERP"), columns=("id","tipo","entita","totale","data_emissione","stato"), show="headings")
            for col in ("id","tipo","entita","totale","data_emissione","stato"):
                self.tree.heading(col, text=col)
            self.tree.pack(fill="both", expand=True, padx=10, pady=10)
            # popola lista ERP
            self.refresh_erp_list()
        else:
            self.chat_area = tk.Text(self, font=("Segoe UI", 11))
            self.chat_area.pack(fill="both", expand=True, padx=10, pady=10)
            self.chat_area.insert("end", f"SISTEMA: Pronto. Modello locale suggerito: {self.modello_locale}.\n")

        if CTK_AVAILABLE:
            self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
            self.input_frame.grid(row=0, column=1, padx=20, pady=25, sticky="s")
            self.entry = ctk.CTkEntry(self.input_frame, placeholder_text="Chiedi all'AI, fai analisi sui dati o carica un PDF...", width=650, height=40)
            self.entry.pack(side="left", padx=10)
            self.entry.bind("<Return>", lambda e: self.send_message())
            ctk.CTkButton(self.input_frame, text="INVIA", width=100, height=40, command=self.send_message).pack(side="right")
        else:
            bottom_frame = tk.Frame(self)
            bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10)
            self.entry = tk.Entry(bottom_frame, width=80)
            self.entry.pack(side="left", padx=6)
            self.entry.bind("<Return>", lambda e: self.send_message())
            tk.Button(bottom_frame, text="INVIA", command=self.send_message).pack(side="right", padx=6)

    def set_mode(self, m):
        self.mode = m
        self.chat_area.insert("end", f"\n[SISTEMA] Modalità impostata su: {m}\n")

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("Documenti", "*.pdf;*.txt"), ("PDF", "*.pdf"), ("Testo", "*.txt")])
        if path:
            try:
                if path.lower().endswith('.pdf'):
                    if PdfReader is None:
                        messagebox.showerror("Errore", "pypdf non disponibile: installa la libreria pypdf per leggere PDF.")
                        return
                    reader = PdfReader(path)
                    self.pdf_pages = [p.extract_text() or "" for p in reader.pages]
                    self.pdf_text = " ".join(self.pdf_pages)
                    pages_count = len(self.pdf_pages)
                else:
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        txt = f.read()
                    self.pdf_pages = [txt]
                    self.pdf_text = txt
                    pages_count = 1
                self.chat_area.insert("end", f"\n[FILE] Caricato {os.path.basename(path)} ({pages_count} pag.)\n")
                # Debug: PDF senza testo leggibile (scansione)
                if not self.pdf_text or not self.pdf_text.strip():
                    self.chat_area.insert("end", "⚠️ Il PDF non contiene testo leggibile (forse è una scansione). Considera l'uso di OCR (pytesseract).\n")
                try:
                    self.btn_analyze.configure(state="normal")
                except Exception:
                    try:
                        self.btn_analyze.config(state="normal")
                    except Exception:
                        pass

                if messagebox.askyesno("File Caricato", "Vuoi analizzare automaticamente questo documento per capire se è una fattura?"):
                    self.analyze_invoice()
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile leggere il file: {e}")

    def send_message(self):
        msg = self.entry.get()
        if not msg:
            return
        try:
            self.entry.delete(0, "end")
        except Exception:
            pass
        self.chat_area.insert("end", f"\n👤 TU: {msg}\n")
        threading.Thread(target=self._ai_logic, args=(msg,), daemon=True).start()

    def _ai_logic(self, query):
        parole_chiave = [p.lower() for p in query.split() if len(p) > 3]
        pdf_context = ""
        if hasattr(self, 'pdf_pages') and self.pdf_pages:
            pagine_trovate = [pag for pag in self.pdf_pages if any(kw in pag.lower() for kw in parole_chiave)]
            pdf_context = "\nDOC: " + "\n".join(pagine_trovate[:2]) if pagine_trovate else "\nDOC: " + "\n".join(self.pdf_pages[:1])

        memoria_testo = "\n".join([f"{m['role'].upper()}: {m['content']}" for m in self.history[-4:]]) if self.history else ""

        prompt = (
            "Sei APEX LEDGER, un assistente contabile AI.\n"
            f"CONTESTO DOCUMENTO (se presente): {pdf_context}\n\n"
            "REGOLE:\n"
            "1. Se la domanda riguarda il documento, usa i dati del contesto.\n"
            "2. Se la domanda è generale, usa la tua conoscenza.\n"
            "3. Rispondi in italiano in modo professionale.\n\n"
            f"STORICO CHAT:\n{memoria_testo}\n\n"
            f"DOMANDA ATTUALE: {query}"
        )

        risposta_finale = ""
        if self.mode == "AI SMART":
            self.chat_area.insert("end", "✨ SMART AI: ")
            try:
                risposta_finale = self.bridge.ask(prompt[:8000])
                self.chat_area.insert("end", f"{risposta_finale}\n")
            except Exception as e:
                self.chat_area.insert("end", f"Errore SmartBridge: {e}\n")
        else:
            self.chat_area.insert("end", f"LOCALE ({self.modello_locale}): ")
            if OLLAMA_AVAILABLE and is_executable_available("ollama"):
                try:
                    stream = ollama.chat(model=self.modello_locale, messages=[{'role': 'user', 'content': prompt}], stream=True)
                    for chunk in stream:
                        txt = chunk['message']['content']
                        risposta_finale += txt
                    self.chat_area.insert("end", "Risposta ricevuta e processata.\n")
                except Exception as e:
                    self.chat_area.insert("end", f"\nErrore locale: {e}\n")
            else:
                try:
                    fallback = self.bridge.ask(prompt[:8000])
                    risposta_finale = fallback
                    self.chat_area.insert("end", f"{fallback}\n")
                except Exception as e:
                    self.chat_area.insert("end", f"Errore fallback AI: {e}\n")

        self.history.append({"role": "user", "content": query})
        self.history.append({"role": "assistant", "content": risposta_finale})

    def analyze_invoice(self):
        self.chat_area.insert("end", "\n🔍 Analisi contabile automatica...\n")
        threading.Thread(target=self._extract_logic, daemon=True).start()

    def _extract_logic(self):
        prompt = (
            "Analizza questo documento. È una fattura? "
            "Rispondi SOLO in formato JSON rigoroso. Estrai:\n"
            "1. 'tipo_doc': 'fattura' o 'altro'\n"
            "2. 'tipo_movimento': 'ENTRATA' o 'USCITA'\n"
            "3. 'fornitore': {nome, indirizzo, piva_cf}\n"
            "4. 'cliente': {nome, indirizzo, piva_cf}\n"
            "5. 'righe': lista di {descrizione, quantita, prezzo_unitario, totale_riga}\n"
            "6. 'totale', 'importo_netto', 'iva'\n"
            "7. 'data_emissione' DD/MM/YYYY\n"
            "8. 'giorni_scadenza'\n"
        )

        raw = ""
        try:
            if OLLAMA_AVAILABLE and is_executable_available("ollama"):
                azienda_info = ""
                try:
                    azi = getattr(self, 'mia_azienda', None)
                    piva = getattr(self, 'mia_piva', None)
                    if azi or piva:
                        azienda_info = f"\nDATI AZIENDA: Nome: {azi or ''} PIVA: {piva or ''}\n"
                except Exception:
                    azienda_info = ""
                try:
                    res = ollama.chat(model=self.modello_locale, messages=[{'role': 'user', 'content': prompt + azienda_info + "\n\n" + self.pdf_text[:4000]}])
                    raw = res.get('message', {}).get('content', '') if isinstance(res, dict) else str(res)
                except Exception:
                    logging.exception("Errore chiamata Ollama")
                    raw = ""
        except Exception:
            raw = ""

        if not raw:
            m_date = re.search(r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})', self.pdf_text)
            date_str = m_date.group(1).replace('-', '/') if m_date else datetime.today().strftime("%d/%m/%Y")
            m_tot = re.search(r'€\s*([\d\.,]+)', self.pdf_text)
            if not m_tot:
                m_tot = re.search(r'([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2}))\s*€?', self.pdf_text)
            tot = 0.0
            if m_tot:
                tot_s = m_tot.group(1).replace('.', '').replace(',', '.')
                try:
                    tot = float(tot_s)
                except:
                    tot = 0.0

            ent_for = "Sconosciuto"
            ent_cli = ""
            piva_match = re.search(r'\b([0-9]{11})\b', self.pdf_text)
            if piva_match:
                piva_val = piva_match.group(1).strip()
            else:
                piva_val = ""
            m_ent = re.search(r'(Fornitore|Mittente|Da|Intestazione)\s*[:\-]?\s*([A-Z0-9a-z \.\'\-\,]{3,120})', self.pdf_text, flags=re.IGNORECASE)
            if m_ent:
                ent_for = m_ent.group(2).strip()

            data_obj = {
                "tipo_doc": "fattura" if tot > 0 else "altro",
                "tipo_movimento": "USCITA" if tot > 0 else "ENTRATA",
                "fornitore": {"nome": ent_for, "indirizzo": "", "piva_cf": piva_val},
                "cliente": {"nome": ent_cli, "indirizzo": "", "piva_cf": ""},
                "righe": [],
                "totale": round(tot, 2),
                "importo_netto": None,
                "iva": None,
                "data_emissione": date_str,
                "giorni_scadenza": 30
            }
            raw = json.dumps(data_obj)

        cleaned = _clean_raw_ai_response(raw)
        logging.debug("Cleaned AI response (hidden from UI): %s", cleaned[:2000])

        json_like = _extract_first_json_object_from_text(cleaned)
        if not json_like:
            m = re.search(r'\{.*?\}', cleaned, re.DOTALL)
            if m:
                json_like = m.group()

        parsed = None
        if json_like:
            try:
                parsed = json.loads(json_like)
            except JSONDecodeError:
                parsed = None
            if parsed is None:
                try:
                    parsed = ast.literal_eval(json_like)
                    if isinstance(parsed, dict):
                        parsed = json.loads(json.dumps(parsed))
                except Exception:
                    parsed = None
            if parsed is None:
                try:
                    s = json_like
                    s = re.sub(r"(?<=\{|,)\s*'([^']+)'\s*:", r'"\1":', s)
                    s = re.sub(r":\s*'([^']*)'", r': "\1"', s)
                    s = re.sub(r",\s*}", "}", s)
                    s = re.sub(r",\s*]", "]", s)
                    parsed = json.loads(s)
                except Exception:
                    parsed = None

        # >>> PATCH: fallback intelligente sempre attivo se parsed è None
        if parsed is None:
            logging.warning("Parsing JSON fallito → attivo fallback intelligente")

            # 🔁 FALLBACK REALE (SEMPRE ATTIVO)
            amounts = _parse_amounts_from_text(self.pdf_text)

            totale = amounts.get("totale") or 0.0

            m_date = re.search(r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})', self.pdf_text)
            data_str = m_date.group(1).replace('-', '/') if m_date else datetime.today().strftime("%d/%m/%Y")

            ent = "Sconosciuto"
            m_ent = re.search(r'(Fornitore|Mittente|Da)\s*[:\-]?\s*(.+)', self.pdf_text, re.IGNORECASE)
            if m_ent:
                ent = m_ent.group(2).strip()

            # 🧠 LOGICA INTELLIGENTE AZIENDA
            tipo_movimento = "USCITA"
            if hasattr(self, 'mia_piva') and self.mia_piva:
                if self.mia_piva in self.pdf_text:
                    tipo_movimento = "ENTRATA"

            parsed = {
                "tipo_doc": "fattura" if totale > 0 else "altro",
                "tipo_movimento": tipo_movimento,
                "fornitore": {"nome": ent},
                "cliente": {},
                "totale": totale,
                "importo_netto": amounts.get("imponibile"),
                "iva": amounts.get("iva"),
                "data_emissione": data_str,
                "giorni_scadenza": 30
            }

            self.chat_area.insert("end", "⚠️ AI non precisa → uso fallback intelligente\n")

        # Arricchisci con imponibile/iva se possibile
        amounts = _parse_amounts_from_text(self.pdf_text)
        totale = parsed.get("totale", None) or amounts.get("totale")
        imponibile = parsed.get("importo_netto", None) or amounts.get("imponibile")
        iva = parsed.get("iva", None) or amounts.get("iva")

        if imponibile is None and iva is None and totale is not None:
            if re.search(r"22\s*%|IVA\s*22", self.pdf_text, flags=re.IGNORECASE):
                try:
                    imponibile = round(totale / 1.22, 2)
                    iva = round(totale - imponibile, 2)
                except:
                    imponibile = None
                    iva = None

        if totale is not None:
            parsed["totale"] = round(float(totale), 2)
        if imponibile is not None:
            parsed["importo_netto"] = round(float(imponibile), 2)
        if iva is not None:
            parsed["iva"] = round(float(iva), 2)

        # Verifica P.IVA mittente vs azienda (se disponibile)
        mittente_match = False
        try:
            if hasattr(self, 'mia_piva') and self.mia_piva:
                if self.mia_piva in self.pdf_text:
                    mittente_match = True
        except Exception:
            mittente_match = False

        # Costruisci messaggio di conferma dettagliato (formattato internamente)
        msg_lines = []
        msg_lines.append(f"Trovata Fattura: {parsed.get('tipo_movimento','')}")
        fornitore = parsed.get("fornitore") or {}
        cliente = parsed.get("cliente") or {}
        if isinstance(fornitore, dict):
            msg_lines.append("Dati del Fornitore:")
            msg_lines.append(f"  Nome: {fornitore.get('nome','')}")
            if fornitore.get('indirizzo'):
                msg_lines.append(f"  Indirizzo: {fornitore.get('indirizzo')}")
            if fornitore.get('piva_cf'):
                msg_lines.append(f"  P.IVA/C.F.: {fornitore.get('piva_cf')}")
        else:
            msg_lines.append(f"Entità: {parsed.get('entita', parsed.get('fornitore',''))}")

        if isinstance(cliente, dict) and cliente.get('nome'):
            msg_lines.append("Dati del Cliente:")
            msg_lines.append(f"  Nome: {cliente.get('nome','')}")
            if cliente.get('indirizzo'):
                msg_lines.append(f"  Indirizzo: {cliente.get('indirizzo')}")
            if cliente.get('piva_cf'):
                msg_lines.append(f"  P.IVA/C.F.: {cliente.get('piva_cf')}")

        righe = parsed.get("righe", [])
        if righe:
            msg_lines.append("Dettaglio Prodotti / Servizi: (estratto)")
            for r in righe[:10]:
                desc = r.get("descrizione", "")[:80]
                q = r.get("quantita", "")
                pu = r.get("prezzo_unitario", "")
                tr = r.get("totale_riga", "")
                msg_lines.append(f"  - {desc} {q} x {pu} = {tr}")

        if "importo_netto" in parsed and parsed.get("importo_netto") is not None:
            msg_lines.append(f"Imponibile (netto): €{float(parsed['importo_netto']):.2f}")
        if "iva" in parsed and parsed.get("iva") is not None:
            msg_lines.append(f"IVA: €{float(parsed['iva']):.2f}")
        if "totale" in parsed:
            msg_lines.append(f"Totale (lordo): €{float(parsed['totale']):.2f}")
        msg_lines.append(f"Emessa il: {parsed.get('data_emissione','')}")
        if hasattr(self, 'mia_piva') and self.mia_piva:
            if mittente_match:
                msg_lines.append(f"ATTENZIONE: La Partita IVA dell'azienda ({self.mia_piva}) è presente nel documento: sembra EMESSA DALLA NOSTRA AZIENDA.")
            else:
                msg_lines.append(f"La Partita IVA dell'azienda ({self.mia_piva}) non è stata trovata nel documento.")

        msg_lines.append("")  # riga vuota
        msg_lines.append("Vuoi registrarla e impostare un promemoria nel calendario?")

        msg = "\n".join(msg_lines)

        try:
            if messagebox.askyesno("Registrazione Fattura SRL", msg):
                db_data = {
                    "tipo_movimento": parsed.get("tipo_movimento", parsed.get("tipo", "")),
                    "fornitore_cliente": fornitore.get('nome') if isinstance(fornitore, dict) else parsed.get("entita", ""),
                    "totale": parsed.get("totale", 0.0),
                    "data_emissione": parsed.get("data_emissione", datetime.today().strftime("%d/%m/%Y")),
                    "giorni_scadenza": parsed.get("giorni_scadenza", 30),
                    "importo_netto": parsed.get("importo_netto", None),
                    "iva": parsed.get("iva", None),
                    "note": ""
                }
                data_em = datetime.strptime(db_data["data_emissione"], "%d/%m/%Y")
                data_scad = data_em + timedelta(days=int(db_data["giorni_scadenza"]))
                self.salva_e_notifica(db_data, data_scad)
        except Exception as e:
            logging.exception("Errore conferma registrazione")

    def salva_e_notifica(self, data, data_scad):
        try:
            conn = sqlite3.connect('apex_ledger_v3.db')
            c = conn.cursor()
            c.execute("""INSERT INTO contabilita 
                        (tipo, fornitore_cliente, importo_netto, iva, totale_fattura, data_emissione, data_scadenza, stato_pagamento, note) 
                        VALUES (?,?,?,?,?,?,?,?,?)""",
                    (data.get('tipo_movimento',''), data.get('fornitore_cliente',''), 
                     data.get('importo_netto', None), data.get('iva', None),
                     float(data.get('totale',0.0)), data.get('data_emissione',''), data_scad.strftime("%d/%m/%Y"), "PENDENTE", data.get('note','')))
            conn.commit()
            conn.close()

            try:
                self.aggiorna_bilancio(data, float(data.get('totale', 0.0)))
            except Exception:
                pass

            self.chat_area.insert("end", f"✅ REGISTRATA: {data.get('fornitore_cliente','')} - €{float(data.get('totale',0.0)):.2f} (Scad. {data_scad.strftime('%d/%m/%Y')})\n")

            if TOAST_AVAILABLE:
                try:
                    toaster = ToastNotifier()
                    toaster.show_toast("APEX LEDGER", 
                                    f"Fattura {data.get('fornitore_cliente','')} registrata.\nPromemoria per il {data_scad.strftime('%d/%m/%Y')}",
                                    duration=5, threaded=True)
                except Exception:
                    safe_print("Notifica toast non disponibile.")

            try:
                webbrowser.open(f"outlookcal:calendar/action/compose&startdt={data_scad.strftime('%Y-%m-%dT09:00:00')}&subject=Scadenza%20Fattura%20{data.get('fornitore_cliente','').replace(' ', '%20')}")
            except Exception:
                pass

        except Exception as e:
            self.chat_area.insert("end", f"❌ Errore salvataggio: {e}\n")

    def aggiorna_bilancio(self, data, totale):
        try:
            if not hasattr(self, 'bilancio_cache'):
                self.bilancio_cache = {}
            mese = datetime.strptime(data.get('data_emissione'), "%d/%m/%Y").strftime("%m/%Y")
            key = (mese, data.get('tipo_movimento',''))
            self.bilancio_cache[key] = self.bilancio_cache.get(key, 0.0) + float(totale)
        except Exception:
            pass

    def mostra_bilancio(self):
        try:
            oggi = datetime.now()
            mese_corrente = oggi.strftime("%m/%Y")

            conn = sqlite3.connect('apex_ledger_v3.db')
            cursor = conn.cursor()

            query = "SELECT tipo, SUM(totale_fattura) FROM contabilita WHERE data_emissione LIKE ? GROUP BY tipo"
            cursor.execute(query, (f"%/{mese_corrente}",))
            risultati = cursor.fetchall()

            entrate = 0.0
            uscite = 0.0

            for tipo, somma in risultati:
                if tipo and tipo.upper() == 'ENTRATA':
                    entrate = somma or 0.0
                elif tipo and tipo.upper() == 'USCITA':
                    uscite = somma or 0.0

            utile = (entrate or 0.0) - (uscite or 0.0)

            report = (
                f"\n--- BILANCIO RAPIDO ({mese_corrente}) ---\n"
                f"ENTRATE: €{(entrate or 0.0):.2f}\n"
                f"USCITE: €{(uscite or 0.0):.2f}\n"
                f"-----------------------------\n"
            )

            if utile >= 0:
                report += f"UTILE PROVVISORIO: €{utile:.2f}\n\n"
            else:
                report += f"PERDITA PROVVISORIA: €{utile:.2f}\n\n"

            self.chat_area.insert("end", report)
            conn.close()

            try:
                self.refresh_erp_list()
            except Exception:
                pass

        except Exception as e:
            self.chat_area.insert("end", f"Errore calcolo bilancio: {e}\n")

    def refresh_erp_list(self):
        try:
            if not hasattr(self, 'tree'):
                return
            for i in self.tree.get_children():
                self.tree.delete(i)
            conn = sqlite3.connect('apex_ledger_v3.db')
            c = conn.cursor()
            c.execute("SELECT id, tipo, fornitore_cliente, totale_fattura, data_emissione, stato_pagamento FROM contabilita ORDER BY id DESC")
            rows = c.fetchall()
            conn.close()
            for r in rows:
                self.tree.insert("", "end", values=(r[0], r[1], r[2], f"€{(r[3] or 0.0):.2f}", r[4], r[5]))
        except Exception as e:
            safe_print(f"Errore refresh ERP list: {e}")

# ---------------------------
# Avvio applicazione
# ---------------------------
def main():
    try:
        app = ApexLedgerApp()
        try:
            app.mainloop()
        except Exception:
            tk.mainloop()
    except Exception as e:
        safe_print(f"Errore avvio applicazione: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
