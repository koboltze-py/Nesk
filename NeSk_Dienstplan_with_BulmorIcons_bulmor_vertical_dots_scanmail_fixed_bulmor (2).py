import pandas as pd
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk, simpledialog
GEN_ROOT_DIR = r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\04_Tagesdienstpläne"

try:
    import os as _os_sa
    _PARENT_SA = _os_sa.path.dirname(GEN_ROOT_DIR.rstrip("\\/"))
    _CAND_SA = _os_sa.path.join(_PARENT_SA, "05_Sonderaufgaben")
    SA_ROOT_DIR = _CAND_SA if _os_sa.path.isdir(_CAND_SA) else GEN_ROOT_DIR
except Exception:
    SA_ROOT_DIR = GEN_ROOT_DIR


SCAN_ROOT_DIR = r"C:\KyoScan"
import webbrowser
import urllib.parse
import os
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except Exception:
    PYODBC_AVAILABLE = False
try:
    from tkcalendar import Calendar
    TKCAL_AVAILABLE = True
except Exception:
    TKCAL_AVAILABLE = False
    class Calendar:
        def __init__(self, *a, **k):
            import tkinter as tk
            self._f = tk.Frame(a[0] if a else None)
            tk.Label(self._f, text='(Kalender nicht verfügbar)').pack()
        def pack(self, *a, **k): self._f.pack(*a, **k)
        def grid(self, *a, **k): self._f.grid(*a, **k)
        def selection_get(self):
            import datetime
            return datetime.date.today()
import openpyxl
import re
import json
# ===== Template-Style Export für "Sonderaufgaben" =====
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ===== Bulmor Status Persistence =====
def _bulmor_state_path():
    import os
    base = os.path.join(os.path.expanduser("~"), "AppData", "Local", "NeSk")
    try:
        os.makedirs(base, exist_ok=True)
    except Exception:
        base = os.getcwd()
    return os.path.join(base, "bulmor_status.json")

def load_bulmor_status():
    try:
        with open(_bulmor_state_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        # ensure keys
        for i in range(1,6):
            data.setdefault(f"Bulmor {i}", "fahrbereit")
        return data
    except Exception:
        return {f"Bulmor {i}": "fahrbereit" for i in range(1,6)}

def save_bulmor_status(state: dict):
    try:
        with open(_bulmor_state_path(), "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        try:
            messagebox.showwarning("Bulmor", f"Status konnte nicht gespeichert werden: {e}")
        except Exception:
            pass

def _export_sonderaufgaben_like_uploaded(assignments_rows, save_path, template_path=None, sheet_name="Sonderaufgaben"):
    """
    assignments_rows: Liste[Dict] mit Keys: 'Aufgabe', 'Tag', 'Nacht' (optional 'Bemerkung')
    save_path: Zielpfad .xlsx
    template_path: Pfad zu "Sonderaufgaben.xlsx" (wenn None -> im Skriptverzeichnis gesucht, sonst Fallback-Layout)
    """
    import os, re
    from datetime import datetime

    # 1) Vorlage bestimmen
    if template_path is None:
        # Erst im Skriptverzeichnis versuchen
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            script_dir = os.getcwd()
        cand = os.path.join(script_dir, "Sonderaufgaben.xlsx")
        if os.path.isfile(cand):
            template_path = cand

    # 2) Vorlage laden oder Fallback bauen
    if template_path and os.path.isfile(template_path):
        wb = load_workbook(template_path)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        # Minimaler Fallback (ähnliche Optik)
        ws["A1"] = "DRK KV Köln e.V."
        ws["C1"] = "Sonderaufgaben"
        ws["E1"] = "Sanitätsstation CGN"
        ws["A2"] = datetime.today()
        ws["C2"] = "Tagdienst"
        ws["E2"] = "Nachtdienst"
        ws["C1"].font = Font(name="Arial", size=26, bold=True)
        ws["C2"].font = Font(name="Arial", size=24, bold=True)
        ws["E2"].font = Font(name="Arial", size=24, bold=True)
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["G"].width = 25

    # 3) Aufgabenzeilen aus Vorlage erfassen + Bemerkung finden
    template_rows = {}
    remark_row = None
    for r in range(3, 500):
        val = ws.cell(row=r, column=1).value
        if val is None:
            continue
        txt = str(val).strip()
        if re.sub(r"\s+", "", txt).lower() == "bemerkung":
            remark_row = r
            continue
        if txt and txt.lower() not in ["tagdienst", "nachtdienst"]:
            template_rows[txt] = r

    def _clone_style(src_cell, dst_cell):
        try:
            if src_cell.has_style:
                dst_cell.font = src_cell.font.copy()
                dst_cell.fill = src_cell.fill.copy()
                dst_cell.alignment = src_cell.alignment.copy()
                dst_cell.border = src_cell.border.copy()
                dst_cell.number_format = src_cell.number_format
        except Exception:
            pass

    ref_row = min(template_rows.values()) if template_rows else None

    def _write_task_row(row_idx, aufgabe, tag_txt, nacht_txt):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2)
        ws.merge_cells(start_row=row_idx, start_column=3, end_row=row_idx, end_column=4)
        ws.merge_cells(start_row=row_idx, start_column=5, end_row=row_idx, end_column=6)

        cA = ws.cell(row=row_idx, column=1); cA.value = aufgabe
        cC = ws.cell(row=row_idx, column=3); cC.value = (tag_txt or "")
        cE = ws.cell(row=row_idx, column=5); cE.value = (nacht_txt or "")

        if ref_row:
            for src_col, dst_col in [(1,1),(3,3),(5,5)]:
                _clone_style(ws.cell(ref_row, src_col), ws.cell(row_idx, dst_col))
        else:
            base_font = Font(name="Arial", size=12)
            left = Alignment(horizontal="left", vertical="center")
            mid  = Alignment(horizontal="center", vertical="center")
            border = Border(left=Side(style="thin"), right=Side(style="thin"),
                            top=Side(style="thin"), bottom=Side(style="thin"))
            for col in [1,2,3,4,5,6]:
                c = ws.cell(row=row_idx, column=col)
                c.font = base_font; c.border = border
                c.alignment = mid if col in (3,4,5,6) else left

    last_task_row = max(template_rows.values()) if template_rows else 2
    attach_before = remark_row if remark_row else (last_task_row + 1)
    norm_tpl = {re.sub(r"\s+", " ", k).strip().lower(): v for k, v in template_rows.items()}

    # Zeilen schreiben
    for row in assignments_rows:
        aufgabe = str(row.get("Aufgabe", "")).strip()
        if not aufgabe:
            continue
        tag_txt = str(row.get("Tag", "")).strip() if row.get("Tag", None) is not None else ""
        nacht_txt = str(row.get("Nacht", "")).strip() if row.get("Nacht", None) is not None else ""

        key = re.sub(r"\s+", " ", aufgabe).lower()
        if key in norm_tpl:
            r = norm_tpl[key]
            ws.cell(row=r, column=3).value = tag_txt
            ws.cell(row=r, column=5).value = nacht_txt
        else:
            insert_row = attach_before
            ws.insert_rows(insert_row, amount=1)
            _write_task_row(insert_row, aufgabe, tag_txt, nacht_txt)
            attach_before += 1

    # Bemerkung (erste nicht-leere)
    if remark_row:
        for row in assignments_rows:
            b = row.get("Bemerkung", None)
            if b and str(b).strip():
                # bei Merges: Top-Left Zelle in Bereich C:D ist C
                ws.cell(row=remark_row, column=3).value = str(b).strip()
                break

    os.makedirs(os.path.dirname(save_path), exist_ok=True) if os.path.dirname(save_path) else None
    wb.save(save_path)
# ===== Ende Template-Style Export =====

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# High-DPI Fix (Windows)
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except (ImportError, AttributeError):
    pass

# Outlook/pywin32
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False  # Warnung zeigen wir nach Tk-Start

# Word/pywin32 (für .doc-Vorlage BTW-Check)
try:
    import win32com.client  # benötigt installiertes MS Word
    WORD_AVAILABLE = True
except Exception:
    WORD_AVAILABLE = False

import sys
def resource_path(*paths):
    """Pfad relativ zum App-Ordner (auch PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(sys.argv[0] if sys.argv and sys.argv[0] else __file__)))
    return os.path.join(base, *paths)

BTW_TEMPLATE_BASENAME = "FO_CGN_9_BTW-Check.doc"
BTW_TEMPLATE_PATH = resource_path(BTW_TEMPLATE_BASENAME)


KRANK_EXPORT_XLSM_BASENAME = "Krankmeldung_Digital September.xlsm"
KRANK_EXPORT_XLSM_PATH = resource_path(KRANK_EXPORT_XLSM_BASENAME)
# -------------------------------------------------------------------
# Konfiguration
# -------------------------------------------------------------------
CONFIG = {
    "image_path": r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\python\drk.jpg",
    "dispo_roles": ["DT", "DT3", "DN"],  # Dispo-Rollen (Word: Zeiten abrunden)
    "excluded_names": ["Peters"],  # für Word (Nachname)
    "excluded_services": ['k', 'krank', 'lg', 'lehrgang', 'ea', 'einarbeitung'],
    "email_recipients": {
        "to": "hildegard.eichler@koeln-bonn-airport.de; erste-hilfe-station-flughafen@drk-koeln.de",
        "cc": "leitung.fb2@drk-koeln.de; verwaltung.fb2@drk-koeln.de; flughafen2@drk-koeln.de; loahrs@gmx.de"
    }
}
CONFIG.setdefault('word_display_date_start_offset_days', -1)
CONFIG.setdefault('word_display_date_end_offset_days', 0)


# -------------------------------------------------------------------
# Hilfsfunktionen
# -------------------------------------------------------------------
def format_time(t):
    import numpy as _np
    from datetime import time as _dtime

    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""

    if isinstance(t, (pd.Timestamp, datetime)):
        return t.strftime("%H:%M")

    if isinstance(t, _dtime):
        return f"{t.hour:02d}:{t.minute:02d}"

    if isinstance(t, (_np.datetime64,)):
        try:
            ts = pd.to_datetime(t)
            return ts.strftime("%H:%M")
        except Exception:
            pass

    if isinstance(t, str):
        s = t.strip()
        m = re.match(r'^(\d{1,2}):(\d{2})(?::\d{2})?$', s)
        if m:
            hh = int(m.group(1)); mm = int(m.group(2))
            return f"{hh:02d}:{mm:02d}"
        try:
            ts = pd.to_datetime(s, errors='raise')
            return ts.strftime("%H:%M")
        except Exception:
            return s

    return str(t)

def extract_lastname(name):
    if isinstance(name, str) and "," in name:
        return name.split(",")[0].strip()

def extract_first2(name):
    """Erste zwei Buchstaben des Vornamens aus 'Nachname, Vorname' oder 'Vorname Nachname'."""
    if not isinstance(name, str):
        return ""
    s = name.strip()
    if "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
        if len(parts) > 1 and parts[1]:
            return parts[1].split()[0][:2]
    parts = s.split()
    if parts:
        return parts[0][:2]
    return ""
    return name.strip() if isinstance(name, str) else name

def group_by_time(entries):
    grouped = {}
    for time, name in entries:
        grouped.setdefault(time, []).append(name)
    return grouped

def add_heading(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

def add_time_block(doc, time, names):
    p = doc.add_paragraph()
    run = p.add_run(time)
    run.bold = True
    run.font.size = Pt(12)
    p.add_run("\t")
    for i, name in enumerate(names):
        if i > 0 and i % 6 == 0:
            p.add_run("\n" + "\t" * 3)
        elif i > 0:
            p.add_run(" / ")
        p.add_run(name)

def generate_time_choices(step_minutes=15):
    times = [""]
    for h in range(24):
        for m in range(0, 60, step_minutes):
            times.append(f"{h:02d}:{m:02d}")
    return times

def _sort_key_minutes(tstr):
    m = re.search(r'(\d{1,2}):(\d{2})', str(tstr))
    if m:
        hh = int(m.group(1)); mm = int(m.group(2))
        return hh * 60 + mm
    return 10**9

def floor_to_full_hour(timestr):
    m = re.match(r'^\s*(\d{1,2}):(\d{2})', str(timestr))
    if not m:
        return format_time(timestr)
    hh = int(m.group(1))
    return f"{hh:02d}:00"

# -------------------------------------------------------------------
# Word-Generator aus DataFrame (Tagesdienstplan)
# -------------------------------------------------------------------
def build_word_from_df(df, plan_date, save_path, paxzahl=None, header_date_start=None, header_date_end=None, selected_names=None):

    # Falls eine manuell bestätigte Auswahl übergeben wurde, nur diese Namen exportieren
    if selected_names is not None:
        name_col = 'Namen' if 'Namen' in df.columns else (df.columns[0] if len(df.columns) else None)
        if name_col:
            def _nm_ok(val):
                try:
                    n = extract_lastname(val) if name_col=='Namen' else str(val)
                    return n in set(selected_names)
                except Exception:
                    return False
            df = df[df[name_col].apply(_nm_ok)]
    doc = Document()

    # Header
    section = doc.sections[0]
    header = section.header
    htable = header.add_table(1, 2, width=Inches(6))
    htable.autofit = False
    hcell1 = htable.cell(0, 0)
    hcell1.width = Inches(1.5)
    try:
        if os.path.isfile(CONFIG["image_path"]):
            hcell1.paragraphs[0].add_run().add_picture(CONFIG["image_path"], width=Inches(1.25))
    except Exception:
        pass
    hcell2 = htable.cell(0, 1)
    hcell2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hcell2.paragraphs[0].text = "Deutsches Rotes Kreuz Kreisverband Köln e.V."
    hcell2.paragraphs[0].add_run("\n- Unfallhilfsstelle und Betreuungsstelle für Behinderte am Flughafen Köln/Bonn –")
    date_para = doc.add_paragraph()
    run_label = date_para.add_run("Datum:")
    run_label.bold = True
    run_label.font.size = Pt(14)
    display_start = header_date_start or (plan_date + timedelta(days=CONFIG.get('word_display_date_start_offset_days', -1)))
    display_end   = header_date_end   or (plan_date + timedelta(days=CONFIG.get('word_display_date_end_offset_days', 0)))
    run_date = date_para.add_run(f"		{display_start.strftime('%d.%m.%Y')} – {display_end.strftime('%d.%m.%Y')}")
    run_date.font.size = Pt(12)
    doc.add_paragraph()

    # Einträge gruppieren – aus df
    dispo_raw = []
    betreuer_raw = []

    for _, r in df.iterrows():
        if 'deactivated' in df.columns and bool(r.get('deactivated', False)):
            continue
        raw_name = r["Namen"]
        last = extract_lastname(raw_name)
        f2 = extract_first2(raw_name)
        name = last
        if name in CONFIG["excluded_names"]:
            continue

        dienst_raw = str(r["Dienst"]).strip() if pd.notnull(r["Dienst"]) else ""
        dienst_lower = dienst_raw.lower()

        if any(ex in dienst_lower for ex in CONFIG["excluded_services"]):
            continue

        beginn = format_time(r["Beginn"])
        ende   = format_time(r["Ende"])

        if any(dienst_raw.upper().startswith(role) for role in CONFIG["dispo_roles"]):
            b_out = floor_to_full_hour(beginn)
            e_out = floor_to_full_hour(ende)
            time_window = f"{b_out} bis {e_out}"
            dispo_raw.append((time_window, last, f2))
        else:
            time_window = f"{beginn} bis {ende}"
            betreuer_raw.append((time_window, last, f2))

    from collections import Counter
    all_last = [t[1] for t in dispo_raw] + [t[1] for t in betreuer_raw]
    counts = Counter(all_last)
    dispo_entries = []
    for tw, last, f2 in dispo_raw:
        disp = f"{last} {f2}" if counts.get(last, 0) > 1 and f2 else last
        dispo_entries.append((tw, disp))
    betreuer_entries = []
    for tw, last, f2 in betreuer_raw:
        disp = f"{last} {f2}" if counts.get(last, 0) > 1 and f2 else last
        betreuer_entries.append((tw, disp))
    dispo_grouped = group_by_time(dispo_entries)
    betreuer_grouped = group_by_time(betreuer_entries)

    add_heading(doc, "Disposition")
    for time in sorted(dispo_grouped.keys(), key=_sort_key_minutes):
        add_time_block(doc, time, dispo_grouped[time])
    doc.add_paragraph()

    add_heading(doc, "Behindertenbetreuer")
    for time in sorted(betreuer_grouped.keys(), key=_sort_key_minutes):
        add_time_block(doc, time, betreuer_grouped[time])

    if paxzahl is not None:
        doc.add_paragraph()
        pax_para = doc.add_paragraph(f"-- {paxzahl} --")
        pax_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    footer = section.footer
    table = footer.add_table(rows=1, cols=3, width=Inches(6))
    table.autofit = True
    cells = table.rows[0].cells
    cells[0].text = "Telefon: +49 220340 – 2323"
    cells[1].text = "email: flughafen@drk-koeln.de"
    cells[2].text = "Stationsleitung: Lars Peters"
    for cell in cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    doc.save(save_path)

# -------------------------------------------------------------------
# DienstplanTab – Excel einlesen & an Viewer weiterreichen
# -------------------------------------------------------------------

class VordruckeTab(ttk.Frame):
    def __init__(self, master, path, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.path = path

        ttk.Label(self, text="Vordrucke", font=("Segoe UI", 12, "bold")).pack(pady=5)

        self.tree = ttk.Treeview(self, columns=("Datei",), show="headings", height=15)
        self.tree.heading("Datei", text="Dateiname")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="Öffnen", command=self.open_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Drucken", command=self.print_file).pack(side=tk.LEFT, padx=5)

        self.load_files()

    def load_files(self):
        import os
        self.tree.delete(*self.tree.get_children())
        if os.path.exists(self.path):
            for file in os.listdir(self.path):
                full_path = os.path.join(self.path, file)
                if os.path.isfile(full_path):
                    self.tree.insert("", "end", values=(file,), iid=full_path)

    def open_file(self):
        import os, tkinter as tk
        sel = self.tree.selection()
        if sel:
            try:
                os.startfile(sel[0])
            except Exception as e:
                tk.messagebox.showerror("Fehler", f"Datei konnte nicht geöffnet werden:\n{e}")

    def print_file(self):
        import os, tkinter as tk
        sel = self.tree.selection()
        if sel:
            try:
                os.startfile(sel[0], "print")
            except Exception as e:
                tk.messagebox.showerror("Fehler", f"Datei konnte nicht gedruckt werden:\n{e}")

class DienstplanTab(tk.Frame):
    def __init__(self, master=None, bg_color=None, fg_color=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.configure(bg=self.bg_color)
        self.create_widgets()
        
    def create_widgets(self):
        # safety inits for new features
        self.kvs_time_map = getattr(self, 'kvs_time_map', {})
        self.deactivated_set = getattr(self, 'deactivated_set', set())
        frame = tk.Frame(self, padx=10, pady=10, bg=self.bg_color)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Excel-Quelldatei:", bg=self.bg_color, fg=self.fg_color).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_entry = tk.Entry(frame, width=60)
        self.file_entry.grid(row=0, column=1, pady=5)
        tk.Button(frame, text="Durchsuchen", command=self.select_file).grid(row=0, column=2, padx=5)

        tk.Button(frame, text="Neu laden", command=self.reload_selected_file).grid(row=0, column=3, padx=5)
        # Auto-Reload (silent) + Countdown
        try:
            self.reload_interval = 60  # Sekunden
            self._reload_seconds_left = self.reload_interval
            self._reload_label = tk.Label(frame, text=f"Auto-Reload in: {self._reload_seconds_left}s",
                                          bg=self.bg_color, fg=self.fg_color)
            self._reload_label.grid(row=0, column=4, padx=10, sticky=tk.W)
            self._reload_timer_id = None
            self._schedule_auto_reload_tick()
        except Exception:
            pass


        hint = tk.Label(frame, text="Die geladene Datei wird im Tab „Tagesdienstplan“ angezeigt.",
                        bg=self.bg_color, fg=self.fg_color)
        hint.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10,0))
        self._init_fs_browser(frame)
        self._init_fs_browser(frame)
        
    def _init_fs_browser(self, parent):
        """Erzeuge einen begrenzten Ordnerbaum unter GEN_ROOT_DIR samt Dateiliste."""
        import os
        wrap = tk.LabelFrame(parent, text="Explorer (Dienstpläne)", bg=self.bg_color, fg=self.fg_color)
        wrap.grid(row=2, column=0, columnspan=5, sticky="nsew", pady=(10,0))
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_columnconfigure(1, weight=2)
        wrap.grid_rowconfigure(0, weight=1)

        left = tk.Frame(wrap, bg=self.bg_color)
        left.grid(row=0, column=0, sticky="nsew", padx=(6,3), pady=6)
        tv_dirs = ttk.Treeview(left, columns=("abspath",), show="tree")
        vs_l = ttk.Scrollbar(left, orient="vertical", command=tv_dirs.yview)
        tv_dirs.configure(yscrollcommand=vs_l.set)
        tv_dirs.pack(side="left", fill="both", expand=True)
        vs_l.pack(side="left", fill="y")
        self._tv_dirs = tv_dirs

        right = tk.Frame(wrap, bg=self.bg_color)
        right.grid(row=0, column=1, sticky="nsew", padx=(3,6), pady=6)
        tv_files = ttk.Treeview(right, columns=("name","size","mtime","abspath"), show="headings", selectmode="extended")
        tv_files.heading("name", text="Datei"); tv_files.column("name", width=300, anchor="w")
        tv_files.heading("size", text="Größe"); tv_files.column("size", width=80, anchor="e")
        tv_files.heading("mtime", text="Geändert"); tv_files.column("mtime", width=140, anchor="w")
        tv_files["displaycolumns"] = ("name","size","mtime")
        vs_r = ttk.Scrollbar(right, orient="vertical", command=tv_files.yview)
        tv_files.configure(yscrollcommand=vs_r.set)
        tv_files.pack(side="left", fill="both", expand=True)
        vs_r.pack(side="left", fill="y")
        self._tv_files = tv_files

        sel_frame = tk.Frame(wrap, bg=self.bg_color)
        sel_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,6))
        tk.Button(sel_frame, text="Zur Auswahl hinzufügen", command=self._add_selected_files).pack(side="left")
        tk.Button(sel_frame, text="Aus Auswahl entfernen", command=self._remove_selected_files).pack(side="left", padx=6)
        tk.Button(sel_frame, text="Auswahl laden", command=self._load_selected_files).pack(side="left", padx=6)

        self._selected_files = []
        self._sel_list = tk.Listbox(sel_frame, height=3, selectmode="extended")
        self._sel_list.pack(side="left", fill="x", expand=True, padx=(10,0))

        root = GEN_ROOT_DIR
        self._root_dir = root
        root_text = os.path.basename(root.rstrip("\\/")) or root
        root_node = tv_dirs.insert("", "end", text=root_text, values=(root,), open=True)
        tv_dirs.insert(root_node, "end", text="...", values=(os.path.join(root, "__dummy__"),))

        def on_open(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            self._populate_dir(node)

        def on_select_dir(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            abspath = tv_dirs.set(node, "abspath")
            self._list_files(abspath)

        tv_dirs.bind("<<TreeviewOpen>>", on_open)
        tv_dirs.bind("<<TreeviewSelect>>", on_select_dir)

        def on_file_double(evt):
            iid = tv_files.focus()
            if not iid: return
            abspath = tv_files.set(iid, "abspath")
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, abspath)
            try:
                self.master.master.handle_file_load(abspath, switch_tab=False, pop_success=False)
            except Exception:
                pass
        tv_files.bind("<Double-1>", on_file_double)

        self._populate_dir(root_node)
        self._list_files(root)

    def _is_within_root(self, path):
        import os
        root = os.path.abspath(self._root_dir)
        try:
            ap = os.path.abspath(path)
        except Exception:
            return False
        return ap.startswith(root)

    def _populate_dir(self, node):
        import os, datetime
        tv = self._tv_dirs
        path = tv.set(node, "abspath")
        if not self._is_within_root(path):
            return
        kids = tv.get_children(node)
        for k in kids:
            if tv.item(k, "text") == "...":
                tv.delete(k)
        try:
            entries = sorted([d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))])
        except Exception:
            entries = []
        for d in entries:
            abspath = os.path.join(path, d)
            if not self._is_within_root(abspath):
                continue
            exists = any(tv.set(c, "abspath") == abspath for c in tv.get_children(node))
            if exists:
                continue
            new = tv.insert(node, "end", text=d, values=(abspath,))
            try:
                if any(os.path.isdir(os.path.join(abspath, x)) for x in os.listdir(abspath)):
                    tv.insert(new, "end", text="...", values=(os.path.join(abspath, "__dummy__"),))
            except Exception:
                pass

    def _list_files(self, dir_path):
        import os, datetime
        tvf = self._tv_files
        for c in tvf.get_children(): tvf.delete(c)
        if not self._is_within_root(dir_path):
            return
        try:
            items = os.listdir(dir_path)
        except Exception:
            items = []
        exts = (".xlsx", ".xlsm", ".xls")
        files = []
        for nm in items:
            abspath = os.path.join(dir_path, nm)
            if os.path.isfile(abspath) and nm.lower().endswith(exts):
                try:
                    st = os.stat(abspath)
                    size = st.st_size
                    mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M")
                except Exception:
                    size = 0; mtime = ""
                files.append((nm, size, mtime, abspath))
        for nm, size, mtime, abspath in sorted(files):
            tvf.insert("", "end", values=(nm, f"{size//1024} KB", mtime, abspath))

    def _add_selected_files(self):
        tvf = self._tv_files
        sel = tvf.selection()
        for iid in sel:
            abspath = tvf.set(iid, "abspath")
            if abspath and abspath not in self._selected_files:
                self._selected_files.append(abspath)
                self._sel_list.insert("end", abspath)

    def _remove_selected_files(self):
        idxs = list(self._sel_list.curselection())[::-1]
        for i in idxs:
            path = self._sel_list.get(i)
            self._sel_list.delete(i)
            try:
                self._selected_files.remove(path)
            except Exception:
                pass

    def _load_selected_files(self):
        if not self._selected_files:
            try:
                messagebox.showinfo("Auswahl", "Bitte Dateien rechts wählen und hinzufügen.")
            except Exception:
                pass
            return
        first = self._selected_files[0]
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, first)
        try:
            self.master.master.handle_file_load(first, switch_tab=False, pop_success=False)
        except Exception:
            pass


    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.master.master.handle_file_load(file_path, switch_tab=False, pop_success=False)


    def reload_selected_file(self):
        """Lädt die im Feld ausgewählte Datei erneut."""
        try:
            path = (self.file_entry.get() or '').strip()
        except Exception:
            path = ''
        if not path:
            try:
                messagebox.showwarning('Neu laden', 'Bitte zuerst eine Datei auswählen (Durchsuchen).')
            except Exception:
                pass
            return
        if not os.path.exists(path):
            try:
                messagebox.showerror('Neu laden', f'Datei nicht gefunden:\n{path}')
            except Exception:
                pass
            return
        ok = self.master.master.handle_file_load(path, switch_tab=False, pop_success=False)
        if ok:
            try:
                self._reload_seconds_left = self.reload_interval
            except Exception:
                pass
# -------------------------------------------------------------------
# Access-Datenbank
# -------------------------------------------------------------------
    def _schedule_auto_reload_tick(self):
        """Plan den nächsten Tick in 1s."""
        try:
            if getattr(self, '_reload_timer_id', None) is not None:
                self.after_cancel(self._reload_timer_id)
        except Exception:
            pass
        try:
            self._reload_timer_id = self.after(1000, self._auto_reload_tick)
        except Exception:
            pass

    def _auto_reload_tick(self):
        """Zählt runter und lädt die ausgewählte Datei still neu, wenn 0 erreicht."""
        try:
            sec = int(getattr(self, '_reload_seconds_left', 60)) - 1
        except Exception:
            sec = 59
        if sec <= 0:
            sec = int(getattr(self, 'reload_interval', 60))
            try:
                path = (self.file_entry.get() or '').strip()
            except Exception:
                path = ''
            if path and os.path.exists(path):
                try:
                    # silent reload
                    self.master.master.handle_file_load(path, switch_tab=False, pop_success=False)
                except Exception:
                    pass
        # Update Anzeige
        try:
            self._reload_seconds_left = sec
            if hasattr(self, '_reload_label'):
                self._reload_label.config(text=f"Auto-Reload in: {sec}s")
        except Exception:
            pass
        self._schedule_auto_reload_tick()



class AccessDB:
    def __init__(self, db_file):
        self.conn = None
        self.cursor = None
        self.db_file = db_file
        self.connect()

    def _ensure_schema(self):
        if not self.conn:
            return
        try:
            try:
                self.cursor.execute("SELECT KvS FROM Tabelle1 WHERE 1=0")
            except pyodbc.Error:
                self.cursor.execute("ALTER TABLE Tabelle1 ADD COLUMN KvS YESNO")
                self.conn.commit()
            try:
                self.cursor.execute("SELECT Gegangen_um FROM Tabelle1 WHERE 1=0")
            except pyodbc.Error:
                self.cursor.execute("ALTER TABLE Tabelle1 ADD COLUMN Gegangen_um DATETIME")
                self.conn.commit()
        except pyodbc.Error:
            pass

    # ---- Bulmor Werkstatt Table helpers ----
    def _ensure_bulmor_table(self):
        if not self.conn:
            return
        try:
            # Try to create table if not exists (Access SQL limited): attempt select
            try:
                self.cursor.execute("SELECT ID FROM BulmorWerkstatt WHERE 1=0")
            except pyodbc.Error:
                try:
                    create_sql = (
                        "CREATE TABLE BulmorWerkstatt ("
                        "ID AUTOINCREMENT PRIMARY KEY, "
                        "Bulmor TEXT(50), "
                        "Status TEXT(50), "
                        "WerkstattTermin DATETIME, "
                        "Freitext MEMO, "
                        "Erledigt YESNO"
                        ")"
                    )
                    self.cursor.execute(create_sql)
                    self.conn.commit()
                except Exception:
                    # If creation fails (permissions) ignore
                    pass
        except Exception:
            pass

    def save_bulmor_entry(self, bulmor_name, status, werkstatt_dt=None, freitext=None, erledigt=False):
        """Insert or update an entry for given bulmor_name and date; keep latest per Bulmor."""
        if not self.conn:
            return False
        try:
            self._ensure_bulmor_table()
            # Try to find existing row for same Bulmor
            self.cursor.execute("SELECT ID FROM BulmorWerkstatt WHERE Bulmor = ?", (bulmor_name,))
            row = self.cursor.fetchone()
            if row:
                # update
                self.cursor.execute(
                    "UPDATE BulmorWerkstatt SET Status = ?, WerkstattTermin = ?, Freitext = ?, Erledigt = ? WHERE ID = ?",
                    (status, werkstatt_dt, freitext, -1 if erledigt else 0, row[0])
                )
            else:
                self.cursor.execute(
                    "INSERT INTO BulmorWerkstatt (Bulmor, Status, WerkstattTermin, Freitext, Erledigt) VALUES (?, ?, ?, ?, ?)",
                    (bulmor_name, status, werkstatt_dt, freitext, -1 if erledigt else 0)
                )
            self.conn.commit()
            return True
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False

    def fetch_bulmor_entries(self):
        if not self.conn:
            return []
        try:
            self._ensure_bulmor_table()
            self.cursor.execute("SELECT ID, Bulmor, Status, WerkstattTermin, Freitext, Erledigt FROM BulmorWerkstatt")
            rows = self.cursor.fetchall()
            return rows
        except Exception:
            return []

    def mark_bulmor_done(self, entry_id, done=True):
        if not self.conn:
            return False
        try:
            self.cursor.execute("UPDATE BulmorWerkstatt SET Erledigt = ? WHERE ID = ?", (-1 if done else 0, entry_id))
            self.conn.commit()
            return True
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False

    def connect(self):
        try:
            if not os.path.exists(self.db_file):
                messagebox.showerror("Datenbankfehler", f"Datenbankfile nicht gefunden: {self.db_file}")
                return False
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' f'DBQ={self.db_file};')
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
            self._ensure_schema()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Verbindungsfehler zur Access-Datenbank: {e}\n\nStellen Sie sicher, dass der passende 'Microsoft Access Driver' installiert ist.")
            self.conn = None
            return False

    def close(self):
        if self.conn:
            self.conn.close()

    def get_or_create_mitarbeiter_id(self, name):
        if not name or not self.conn:
            return None
        try:
            self.cursor.execute("SELECT Mitarbeiter_ID FROM Tabelle2 WHERE Name = ?", (name,))
            result = self.cursor.fetchone()
            if result:
                return result[0]
            else:
                self.cursor.execute("INSERT INTO Tabelle2 (Name) VALUES (?)", (name,))
                self.conn.commit()
                return self.cursor.execute("SELECT @@IDENTITY").fetchone()[0]
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Zugriff auf Tabelle2: {e}")
            return None

    def save_krankmeldung(self, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=False, t_gegangen=None):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            meldender_id = self.get_or_create_mitarbeiter_id(s_meldender)
            angenommen_von_id = self.get_or_create_mitarbeiter_id(s_angenommen)
            if meldender_id is None or angenommen_von_id is None:
                return False
            insert_query = """
                INSERT INTO Tabelle1
                    ([Datum], [Meldender_ID], [Krank_von], [Krank_bis], [Anruf_um], [Angenommen_von_ID], [Bemerkung], [KvS], [Gegangen_um])
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            self.cursor.execute(insert_query, (
                d_datum,
                meldender_id,
                d_von,
                d_bis,
                t_anruf.time(),
                angenommen_von_id,
                s_bem,
                -1 if kvs else 0,
                t_gegangen.time() if t_gegangen else None
            ))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Speichern der Krankmeldung: {e}")
            self.conn.rollback()
            return False
            
    def update_krankmeldung(self, entry_id, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=False, t_gegangen=None):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            meldender_id = self.get_or_create_mitarbeiter_id(s_meldender)
            angenommen_von_id = self.get_or_create_mitarbeiter_id(s_angenommen)
            if meldender_id is None or angenommen_von_id is None:
                return False
            update_query = """
                UPDATE Tabelle1
                SET [Datum] = ?, [Meldender_ID] = ?, [Krank_von] = ?, [Krank_bis] = ?, [Anruf_um] = ?,
                    [Angenommen_von_ID] = ?, [Bemerkung] = ?, [KvS] = ?, [Gegangen_um] = ?
                WHERE Eintrag_ID = ?
            """
            self.cursor.execute(update_query, (
                d_datum,
                meldender_id,
                d_von,
                d_bis,
                t_anruf.time(),
                angenommen_von_id,
                s_bem,
                -1 if kvs else 0,
                t_gegangen.time() if t_gegangen else None,
                entry_id
            ))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Aktualisieren der Krankmeldung: {e}")
            self.conn.rollback()
            return False

    def delete_krankmeldung(self, entry_id):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            delete_query = "DELETE FROM Tabelle1 WHERE Eintrag_ID = ?"
            self.cursor.execute(delete_query, (entry_id,))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Löschen der Krankmeldung: {e}")
            self.conn.rollback()
            return False

    def fetch_all_krankmeldungen_with_id(self, start_date=None, end_date=None):
        if not self.conn:
            return []
        try:
            query = """
                SELECT
                    T1.Eintrag_ID,
                    T1.Datum,
                    T2_Meldender.Name AS Meldender,
                    T1.Krank_von,
                    T1.Krank_bis,
                    T1.Anruf_um,
                    T2_Angenommen.Name AS Angenommen_von,
                    T1.Bemerkung,
                    T1.KvS,
                    T1.Gegangen_um
                FROM
                    (Tabelle1 AS T1
                    INNER JOIN Tabelle2 AS T2_Meldender ON T1.Meldender_ID = T2_Meldender.Mitarbeiter_ID)
                    INNER JOIN Tabelle2 AS T2_Angenommen ON T1.Angenommen_von_ID = T2_Angenommen.Mitarbeiter_ID
            """
            params = []
            if start_date and end_date:
                query += " WHERE T1.Datum BETWEEN ? AND ?"
                end_of_day = datetime(end_date.year, end_date.month, end_date.day, 23, 59, 59)
                params = [start_date, end_of_day]
            query += " ORDER BY T1.Datum DESC;"
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der Daten: {e}")
            return []

    def fetch_sick_employees_in_range(self, start_date, end_date):
        if not self.conn:
            return []
        try:
            query = """
                SELECT DISTINCT T2.Name, T1.Krank_von, T1.Krank_bis
                FROM Tabelle1 AS T1
                INNER JOIN Tabelle2 AS T2 ON T1.Meldender_ID = T2.Mitarbeiter_ID
                WHERE T1.Krank_bis >= ? AND T1.Krank_von <= ?;
            """
            self.cursor.execute(query, (start_date, end_date))
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der aktuell Kranken: {e}")
            return []

    def fetch_kvs_for_date(self, date_obj):
        if not self.conn:
            return []
        try:
            plan_start = datetime(date_obj.year, date_obj.month, date_obj.day, 0, 0, 0)
            plan_end   = datetime(date_obj.year, date_obj.month, date_obj.day, 23, 59, 59)
            query = """
                SELECT T2.Name, T1.Gegangen_um, T1.Krank_von, T1.Krank_bis
                FROM Tabelle1 AS T1
                INNER JOIN Tabelle2 AS T2 ON T1.Meldender_ID = T2.Mitarbeiter_ID
                WHERE T1.KvS <> 0
                  AND T1.Krank_von <= ?
                  AND T1.Krank_bis  >= ?;
            """
            self.cursor.execute(query, (plan_end, plan_start))
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der KvS-Daten: {e}")
            return []
    
    def save_dienstplan_from_df(self, df, show_message=True, replace_for_date=True):
        if not self.conn:
            if show_message:
                messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        
        create_table_sql = """
            CREATE TABLE Dienstplan (
                Dienstplan_ID AUTOINCREMENT PRIMARY KEY,
                Datum DATE,
                Namen TEXT(255),
                Dienst TEXT(255),
                Beginn DATETIME,
                Ende DATETIME
            )
        """
        insert_sql = "INSERT INTO Dienstplan (Datum, Namen, Dienst, Beginn, Ende) VALUES (?, ?, ?, ?, ?)"

        plan_date_value = None
        try:
            if 'Datum' in df.columns and len(df) > 0:
                raw = df.iloc[0]['Datum']
                if isinstance(raw, (datetime, pd.Timestamp)):
                    plan_date_value = raw.date()
                else:
                    plan_date_value = pd.to_datetime(str(raw), dayfirst=True, errors='coerce')
                    if isinstance(plan_date_value, pd.Timestamp):
                        plan_date_value = plan_date_value.date()
                    elif isinstance(plan_date_value, datetime):
                        plan_date_value = plan_date_value.date()
                    else:
                        plan_date_value = None
        except Exception:
            plan_date_value = None

        try:
            try:
                self.cursor.execute(create_table_sql)
                self.conn.commit()
            except pyodbc.ProgrammingError:
                pass

            if replace_for_date and plan_date_value is not None:
                try:
                    self.cursor.execute("DELETE FROM Dienstplan WHERE Datum = ?", (plan_date_value,))
                    self.conn.commit()
                except pyodbc.Error:
                    pass

            for _, row in df.iterrows():
                if 'deactivated' in df.columns and bool(row.get('deactivated', False)):
                    continue
                if isinstance(row['Datum'], (datetime, pd.Timestamp)):
                    datum = row['Datum'].date()
                else:
                    try:
                        datum = pd.to_datetime(str(row['Datum']), dayfirst=True, errors='coerce')
                        datum = datum.date() if isinstance(datum, pd.Timestamp) else None
                    except Exception:
                        datum = None

                def _to_dt(v):
                    if isinstance(v, (datetime, pd.Timestamp)):
                        return v
                    try:
                        t = pd.to_datetime(str(v), errors='coerce')
                        return t.to_pydatetime() if isinstance(t, pd.Timestamp) else None
                    except Exception:
                        return None

                beginn = _to_dt(row['Beginn'])
                ende   = _to_dt(row['Ende'])

                self.cursor.execute(insert_sql, (datum, row['Namen'], row['Dienst'], beginn, ende))

            self.conn.commit()
            if show_message:
                messagebox.showinfo("Erfolg", "Dienstplandaten erfolgreich in die Datenbank gespeichert!")
            return True

        except pyodbc.Error as e:
            self.conn.rollback()
            if show_message:
                messagebox.showerror("Datenbankfehler", f"Fehler beim Speichern der Dienstplandaten: {e}")
            return False

# -------------------------------------------------------------------
# Krankmeldung-Tab (mit KvS)
# -------------------------------------------------------------------
class KrankmeldungTab(ttk.Frame):
    def __init__(self, master=None, db=None, style_vars=None, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.editing_id = None
        self.filter_manage_start_date = None
        self.filter_manage_end_date = None
        
        if style_vars:
            self.pastel_bg = style_vars.get('bg')
            self.pastel_fg = style_vars.get('fg')
            self.pastel_button = style_vars.get('button')
        else:
            self.pastel_bg = '#E0F7FA'
            self.pastel_fg = '#37474F'
            self.pastel_button = '#B2DFDB'
        
        self.configure(style='TFrame')
        self.style = ttk.Style(self)
        self.style.configure('TFrame', background=self.pastel_bg)
        self.style.configure('TLabel', background=self.pastel_bg, foreground=self.pastel_fg, font=('Segoe UI', 10))
        self.style.configure('TEntry', fieldbackground='white', foreground=self.pastel_fg, font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), foreground=self.pastel_fg, background=self.pastel_button, relief='flat')
        self.style.map('TButton', background=[('active', '#80CBC4'), ('pressed', '#4DB6AC')])
        self.style.configure('Treeview', background='white', foreground=self.pastel_fg, fieldbackground='white', rowheight=25, font=('Segoe UI', 10))
        # FIX: kein falsches Anführungszeichen!
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background='#B2DFDB', foreground=self.pastel_fg)

        self.main_container = ttk.Frame(self, padding="15", style='TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.form_frame = ttk.Frame(self.main_container, style='TFrame')
        self.form_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        self.tables_frame = ttk.Frame(self.main_container, style='TFrame')
        self.tables_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.create_form_widgets()
        self.create_tables_view()
        self.initialize_form()
        if self.db and self.db.conn:
            self.display_data_in_tables()

    def create_form_widgets(self):
        input_frame = ttk.Frame(self.form_frame, style='TFrame')
        input_frame.pack(padx=10, pady=10, fill=tk.Y, expand=False)
        
        ttk.Label(input_frame, text="Datum der Meldung:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_datum = ttk.Entry(input_frame)
        self.txt_datum.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Label(input_frame, text="Meldender:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_meldender = ttk.Entry(input_frame)
        self.txt_meldender.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Krank von:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.btn_krank_von = ttk.Button(input_frame, text="", command=lambda: self.show_calendar("von"))
        self.btn_krank_von.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Krank bis:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.btn_krank_bis = ttk.Button(input_frame, text="", command=lambda: self.show_calendar("bis"))
        self.btn_krank_bis.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Anruf um:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_anruf_um = ttk.Entry(input_frame)
        self.txt_anruf_um.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Angenommen von:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_angenommen_von = ttk.Entry(input_frame)
        self.txt_angenommen_von.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Bemerkung:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_bemerkung = ttk.Entry(input_frame)
        self.txt_bemerkung.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Label(input_frame, text="Kommentar:").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_kommentar = tk.Text(input_frame, height=3, width=30)
        self.txt_kommentar.grid(row=7, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        self.kvs_var = tk.BooleanVar(value=False)
        self.chk_kvs = ttk.Checkbutton(input_frame, text="KvS (krank von Station)", variable=self.kvs_var, command=self._toggle_kvs_controls)
        self.chk_kvs.grid(row=9, column=0, sticky=tk.W, padx=5, pady=(10, 0))

        ttk.Label(input_frame, text="Gegangen um:").grid(row=9, column=1, sticky=tk.W, padx=5, pady=(10, 0))
        self.cbo_gegangen_um = ttk.Combobox(input_frame, values=generate_time_choices(15), state="disabled", width=8)
        self.cbo_gegangen_um.grid(row=9, column=1, sticky=(tk.E), padx=5, pady=(10, 0))
        self.cbo_gegangen_um.set("")

        self.save_button = ttk.Button(input_frame, text="Speichern", command=self.cmd_speichern_click)
        self.save_button.grid(row=10, column=0, sticky=tk.W, padx=5, pady=10)

        self.cancel_button = ttk.Button(input_frame, text="Abbrechen", command=self.clear_form)
        self.cancel_button.grid(row=10, column=1, sticky=tk.W, padx=5, pady=10)

    

        self.export_xls_button = ttk.Button(input_frame, text="Export Excel (Meldedatum)", command=self.export_krankmeldungen_to_excel)
        self.export_xls_button.grid(row=10, column=2, sticky=tk.W, padx=5, pady=10)

    def export_krankmeldungen_to_excel(self):
        """Exportiert alle Krankmeldungen in die .xlsm-Monatsdatei (Blatt 'Krankmeldungen').
        Duplikate werden anhand (Datum, Meldender, Anruf um) verhindert."""
        if not getattr(self, "db", None) or not getattr(self.db, "conn", None):
            return
        target_path = KRANK_EXPORT_XLSM_PATH
        if not os.path.exists(target_path):
            return
        try:
            from openpyxl import load_workbook
            from datetime import datetime
            wb = load_workbook(target_path, keep_vba=True)
            ws = wb["Krankmeldungen"] if "Krankmeldungen" in wb.sheetnames else wb.active

            existing = set()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or all(v is None for v in row[:6]):
                    continue
                d, person, _, _, t, _ = (row + (None,)*6)[:6]
                try:
                    if isinstance(d, datetime): d = d.date()
                    elif hasattr(d, "date"): d = d.date()
                except Exception: pass
                try:
                    if isinstance(t, datetime): t = t.time()
                except Exception: pass
                existing.add((d, (person or "").strip(), t))

            rows = self.db.fetch_all_krankmeldungen()
            if not rows: 
                return

            def to_date(x):
                try:
                    if isinstance(x, datetime): return x.date()
                    if hasattr(x, "date"): return x.date()
                except Exception: pass
                return x

            def to_time(x):
                try:
                    if isinstance(x, datetime): return x.time()
                except Exception: pass
                return x

            for r in rows:
                try:
                    datum = r[1]; meldender = r[2]; krank_von = r[3]; krank_bis = r[4]; anruf_um = r[5]; angenommen_von = r[6]; bemerkung = r[7]
                except Exception:
                    datum = getattr(r, "Datum", None)
                    meldender = getattr(r, "Meldender", "")
                    krank_von = getattr(r, "Krank_von", None)
                    krank_bis = getattr(r, "Krank_bis", None)
                    anruf_um = getattr(r, "Anruf_um", None)
                    angenommen_von = getattr(r, "Angenommen_von", "")
                    bemerkung = getattr(r, "Bemerkung", "")

                key = (to_date(datum), (meldender or "").strip(), to_time(anruf_um))
                if key in existing:
                    continue

                ws.append([
                    to_date(datum),
                    (meldender or ""),
                    to_date(krank_von),
                    to_date(krank_bis),
                    to_time(anruf_um),
                    (angenommen_von or ""),
                    (bemerkung or ""),
                    None
                ])
                existing.add(key)

            wb.save(target_path)
        except Exception:
            return
    def _toggle_kvs_controls(self):
        self.cbo_gegangen_um.configure(state=("normal" if self.kvs_var.get() else "disabled"))
        if not self.kvs_var.get():
            self.cbo_gegangen_um.set("")

    def create_tables_view(self):
        sick_list_frame = ttk.Frame(self.tables_frame, style='TFrame')
        sick_list_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=10)

        ttk.Label(sick_list_frame, text="Aktuell kranke Mitarbeiter:", font=('Segoe UI', 11, 'bold')).pack(pady=(0, 5))
        self.sick_employees_tree = ttk.Treeview(sick_list_frame, columns=("Name", "Von", "Bis"), show="headings", selectmode="none")
        self.sick_employees_tree.heading("Name", text="Name")
        self.sick_employees_tree.heading("Von", text="Krank von")
        self.sick_employees_tree.heading("Bis", text="Krank bis")
        self.sick_employees_tree.column("Name", width=250, anchor="w")
        self.sick_employees_tree.column("Von", width=150, anchor="center")
        self.sick_employees_tree.column("Bis", width=150, anchor="center")
        self.sick_employees_tree.pack(fill=tk.BOTH, expand=True)
        
        manage_frame = ttk.Frame(self.tables_frame, style='TFrame')
        manage_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(manage_frame, text="Datensätze verwalten:", font=('Segoe UI', 11, 'bold')).pack(pady=(0, 5))
        
        top_frame = ttk.Frame(manage_frame, style='TFrame')
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        filter_frame = ttk.Frame(top_frame, style='TFrame')
        filter_frame.pack(side=tk.LEFT, fill=tk.X)
        
        ttk.Label(filter_frame, text="Datensätze filtern:").pack(side=tk.LEFT, padx=(0, 5))
        self.btn_manage_filter_start = ttk.Button(filter_frame, text="Startdatum", command=lambda: self.show_calendar("manage_start"))
        self.btn_manage_filter_start.pack(side=tk.LEFT, padx=2)

        self.btn_manage_filter_end = ttk.Button(filter_frame, text="Enddatum", command=lambda: self.show_calendar("manage_end"))
        self.btn_manage_filter_end.pack(side=tk.LEFT, padx=2)

        ttk.Button(filter_frame, text="Filtern", command=self.filter_manage_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="Zurücksetzen", command=self.reset_manage_filter).pack(side=tk.LEFT, padx=5)

        ttk.Button(top_frame, text="Daten exportieren", command=self.export_data).pack(side=tk.RIGHT, padx=5)

        self.manage_tree = ttk.Treeview(
            manage_frame,
            columns=("ID", "Datum", "Meldender", "Von", "Bis", "Anruf", "Angenommen", "Bemerkung", "KvS", "Gegangen um"),
            show="headings"
        )
        headers = [
            ("ID", 40, "center"),
            ("Datum", 120, "center"),
            ("Meldender", 150, "w"),
            ("Von", 120, "center"),
            ("Bis", 120, "center"),
            ("Anruf", 80, "center"),
            ("Angenommen", 150, "w"),
            ("Bemerkung", 250, "w"),
            ("KvS", 60, "center"),
            ("Gegangen um", 100, "center")
        ]
        for name, width, anchor in headers:
            self.manage_tree.heading(name, text=name)
            self.manage_tree.column(name, width=width, anchor=anchor)
        self.manage_tree.pack(fill=tk.BOTH, expand=True)

        action_frame = ttk.Frame(manage_frame, style='TFrame')
        action_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(action_frame, text="Bearbeiten", command=self.edit_selected_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Löschen", command=self.delete_selected_entry).pack(side=tk.LEFT, padx=5)

    def initialize_form(self):
        self.txt_datum.delete(0, tk.END)
        self.txt_datum.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.txt_anruf_um.delete(0, tk.END)
        self.txt_anruf_um.insert(0, datetime.now().strftime("%H:%M"))
        self.btn_krank_von.config(text=datetime.now().strftime("%d.%m.%Y"))
        self.btn_krank_bis.config(text=datetime.now().strftime("%d.%m.%Y"))
        self.txt_meldender.delete(0, tk.END)
        self.txt_angenommen_von.delete(0, tk.END)
        self.txt_bemerkung.delete(0, tk.END)
        try:
            self.txt_kommentar.delete('1.0','end')
        except Exception:
            pass
        self.kvs_var.set(False)
        self.cbo_gegangen_um.set("")
        self._toggle_kvs_controls()
        self.editing_id = None
        self.save_button.config(text="Speichern")

    def show_calendar(self, button_type):
        top = tk.Toplevel(self.master)
        top.title("Datum auswählen")
        top.geometry("400x400")
        top.configure(bg=self.pastel_bg)

        cal = Calendar(top, selectmode="day", date_pattern="dd.mm.yyyy",
                       background=self.pastel_button, foreground=self.pastel_fg,
                       headersbackground='#80CBC4', selectbackground='#4DB6AC')
        cal.pack(padx=10, pady=10)

        def get_date():
            selected_date = cal.get_date()
            if button_type == "von":
                self.btn_krank_von.config(text=selected_date)
            elif button_type == "bis":
                self.btn_krank_bis.config(text=selected_date)
            elif button_type == "manage_start":
                self.btn_manage_filter_start.config(text=selected_date)
            elif button_type == "manage_end":
                self.btn_manage_filter_end.config(text=selected_date)
            top.destroy()
        
        ttk.Button(top, text="Auswählen", command=get_date).pack(pady=10)

    def _parse_time_hhmm(self, s):
        s = (s or "").strip()
        if not s:
            return None
        m = re.match(r'^(\d{1,2}):(\d{2})$', s)
        if not m:
            raise ValueError("Zeitformat muss HH:MM sein.")
        hh, mm = int(m.group(1)), int(m.group(2))
        if not (0 <= hh <= 23 and 0 <= mm <= 59):
            raise ValueError("Ungültige Uhrzeit.")
        now = datetime.now()
        return datetime(now.year, now.month, now.day, hh, mm)

    def cmd_speichern_click(self):
        if not self.db or not self.db.conn:
            messagebox.showerror("Fehler", "Bitte verbinden Sie sich zuerst mit einer Datenbank.")
            return

        d_datum_str = self.txt_datum.get()
        s_meldender = self.txt_meldender.get().strip()
        d_von_str = self.btn_krank_von.cget("text")
        d_bis_str = self.btn_krank_bis.cget("text")
        t_anruf_str = self.txt_anruf_um.get()
        s_angenommen = self.txt_angenommen_von.get().strip()
        s_bem = self.txt_bemerkung.get().strip()
        # Kommentar (optional) an Bemerkung anhängen
        kommentar_val = ""
        try:
            kommentar_val = self.txt_kommentar.get("1.0", "end").strip()
        except Exception:
            kommentar_val = ""
        if kommentar_val:
            s_bem = (s_bem + (" | Kommentar: " + kommentar_val)) if s_bem else ("Kommentar: " + kommentar_val)
        kvs_flag = bool(self.kvs_var.get())
        t_gegangen_str = self.cbo_gegangen_um.get().strip()
        
        if not all([d_datum_str, s_meldender, d_von_str, d_bis_str, t_anruf_str, s_angenommen]):
            messagebox.showerror("Eingabefehler", "Alle Felder außer 'Bemerkung' und 'Gegangen um' müssen ausgefüllt sein.")
            return

        try:
            d_datum = datetime.strptime(d_datum_str, "%d.%m.%Y")
            d_von = datetime.strptime(d_von_str, "%d.%m.%Y")
            d_bis = datetime.strptime(d_bis_str, "%d.%m.%Y")
            t_anruf = datetime.strptime(t_anruf_str, "%H:%M")
            t_gegangen = self._parse_time_hhmm(t_gegangen_str) if t_gegangen_str else None
        except ValueError as e:
            messagebox.showerror("Formatfehler", f"Ungültiges Datums- oder Zeitformat: {e}\nErwartete Formate: Datum (dd.mm.yyyy), Zeit (HH:MM)")
            return
            
        if self.editing_id is None:
            success = self.db.save_krankmeldung(d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=kvs_flag, t_gegangen=t_gegangen)
        try:
            self.export_krankmeldungen_to_excel()
        except Exception:
            pass
            if success:
                messagebox.showinfo("Erfolg", "Krankmeldung erfolgreich gespeichert!")
        else:
            success = self.db.update_krankmeldung(self.editing_id, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=kvs_flag, t_gegangen=t_gegangen)
            if success:
                messagebox.showinfo("Erfolg", "Krankmeldung erfolgreich aktualisiert!")

        if success:
            self.clear_form()
            self.display_data_in_tables()
            try:
                app = self.master.master
                if hasattr(app, "excel_viewer_tab") and app.excel_viewer_tab and app.excel_viewer_tab.df is not None:
                    app.excel_viewer_tab.recompute_sickness_from_db()
                if hasattr(app, "statistics_tab") and app.statistics_tab:
                    app.statistics_tab.update_statistics(app.excel_viewer_tab.df if app.excel_viewer_tab.df is not None else None)
            except Exception:
                pass

    def clear_form(self):
        self.initialize_form()

    def edit_selected_entry(self):
        selected_item = self.manage_tree.focus()
        if not selected_item:
            messagebox.showwarning("Keine Auswahl", "Bitte wählen Sie einen Eintrag zum Bearbeiten aus.")
            return
        values = self.manage_tree.item(selected_item, 'values')
        self.editing_id = values[0]

        self.txt_datum.delete(0, tk.END); self.txt_datum.insert(0, values[1])
        self.txt_meldender.delete(0, tk.END); self.txt_meldender.insert(0, values[2])
        self.btn_krank_von.config(text=values[3])
        self.btn_krank_bis.config(text=values[4])
        self.txt_anruf_um.delete(0, tk.END); self.txt_anruf_um.insert(0, values[5])
        self.txt_angenommen_von.delete(0, tk.END); self.txt_angenommen_von.insert(0, values[6])
        self.txt_bemerkung.delete(0, tk.END); self.txt_bemerkung.insert(0, values[7])

        self.kvs_var.set(str(values[8]).strip().lower() in ("-1", "1", "true", "ja"))
        self._toggle_kvs_controls()
        self.cbo_gegangen_um.set(values[9] if values[9] else "")

        self.save_button.config(text="Aktualisieren")

    def delete_selected_entry(self):
        selected_item = self.manage_tree.focus()
        if not selected_item:
            messagebox.showwarning("Keine Auswahl", "Bitte wählen Sie einen Eintrag zum Löschen aus.")
            return
        entry_id = self.manage_tree.item(selected_item, 'values')[0]
        if messagebox.askyesno("Löschen bestätigen", f"Möchten Sie Eintrag {entry_id} wirklich löschen?"):
            if self.db and self.db.delete_krankmeldung(entry_id):
                messagebox.showinfo("Erfolg", "Eintrag erfolgreich gelöscht.")
                self.display_data_in_tables()
                try:
                    app = self.master.master
                    if hasattr(app, "excel_viewer_tab") and app.excel_viewer_tab and app.excel_viewer_tab.df is not None:
                        app.excel_viewer_tab.recompute_sickness_from_db()
                    if hasattr(app, "statistics_tab") and app.statistics_tab:
                        app.statistics_tab.update_statistics(app.excel_viewer_tab.df if app.excel_viewer_tab.df is not None else None)
                except Exception:
                    pass

    def display_data_in_tables(self):
        if not self.db or not self.db.conn:
            self.manage_tree.delete(*self.manage_tree.get_children())
            self.sick_employees_tree.delete(*self.sick_employees_tree.get_children())
            self.manage_tree.insert('', 'end', values=("Keine Datenbankverbindung.", "", "", "", "", "", "", "", "", ""))
            self.sick_employees_tree.insert('', 'end', values=("Keine Krankmeldungen im Zeitraum", "", ""))
            return

        all_data = self.db.fetch_all_krankmeldungen_with_id()
        self.manage_tree.delete(*self.manage_tree.get_children())
        self.sick_employees_tree.delete(*self.sick_employees_tree.get_children())

        for row in all_data:
            formatted_row = list(row)
            formatted_row[1] = row[1].strftime("%d.%m.%Y")
            formatted_row[3] = row[3].strftime("%d.%m.%Y")
            formatted_row[4] = row[4].strftime("%d.%m.%Y")
            formatted_row[5] = row[5].strftime("%H:%M")
            formatted_row[8] = "Ja" if (row[8] in (-1, 1, True)) else "Nein"
            formatted_row[9] = row[9].strftime("%H:%M") if row[9] else ""
            self.manage_tree.insert('', 'end', values=formatted_row)
        
        self.update_current_sickness_view()

    def update_current_sickness_view(self, start_date=None, end_date=None):
        if not self.sick_employees_tree:
            return

        for item in self.sick_employees_tree.get_children():
            self.sick_employees_tree.delete(item)

        if start_date is None or end_date is None:
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=3)

        if not self.db or not self.db.conn:
            self.sick_employees_tree.insert("", "end", values=("Keine Datenbankverbindung.", "", ""))
            return

        sick_employees = self.db.fetch_sick_employees_in_range(start_date, end_date)

        if not sick_employees:
            self.sick_employees_tree.insert("", "end", values=("Keine Krankmeldungen im Zeitraum", "", ""))
            return

        for name, krank_von, krank_bis in sick_employees:
            von_str = krank_von.strftime("%d.%m.%Y") if isinstance(krank_von, (datetime, pd.Timestamp)) else str(krank_von)
            bis_str = krank_bis.strftime("%d.%m.%Y") if isinstance(krank_bis, (datetime, pd.Timestamp)) else str(krank_bis)
            self.sick_employees_tree.insert("", "end", values=(name, von_str, bis_str))

    def filter_manage_data(self):
        if not self.db or not self.db.conn:
            messagebox.showwarning("Fehler", "Bitte verbinden Sie sich zuerst mit einer Datenbank.")
            return
        start_date_str = self.btn_manage_filter_start.cget("text")
        end_date_str = self.btn_manage_filter_end.cget("text")
        if start_date_str == "Startdatum" or end_date_str == "Enddatum":
            messagebox.showwarning("Fehler", "Bitte wählen Sie sowohl ein Start- als auch ein Enddatum aus.")
            return
        self.filter_manage_start_date = self.parse_date_de(start_date_str)
        self.filter_manage_end_date = self.parse_date_de(end_date_str)
        if not self.filter_manage_start_date or not self.filter_manage_end_date:
            messagebox.showwarning("Fehler", "Ungültiges Datumsformat.")
            return
        if self.filter_manage_end_date < self.filter_manage_start_date:
            messagebox.showwarning("Fehler", "Das Enddatum muss nach dem Startdatum liegen.")
            return
        self.update_table_view(self.manage_tree, self.db.fetch_all_krankmeldungen_with_id(self.filter_manage_start_date, self.filter_manage_end_date))
        
    def reset_manage_filter(self):
        self.filter_manage_start_date = None
        self.filter_manage_end_date = None
        self.btn_manage_filter_start.config(text="Startdatum")
        self.btn_manage_filter_end.config(text="Enddatum")
        self.display_data_in_tables()
    
    def update_table_view(self, treeview, data):
        for item in treeview.get_children():
            treeview.delete(item)
        if not data:
            treeview.insert("", "end", values=("Keine Daten im Zeitraum gefunden.", "", "", "", "", "", "", "", "", ""))
            return
        for row in data:
            formatted_row = list(row)
            formatted_row[1] = row[1].strftime("%d.%m.%Y")
            formatted_row[3] = row[3].strftime("%d.%m.%Y")
            formatted_row[4] = row[4].strftime("%d.%m.%Y")
            formatted_row[5] = row[5].strftime("%H:%M")
            formatted_row[8] = "Ja" if (row[8] in (-1, 1, True)) else "Nein"
            formatted_row[9] = row[9].strftime("%H:%M") if row[9] else ""
            treeview.insert('', 'end', values=formatted_row)

    def export_data(self):
        if not self.db or not self.db.conn:
            messagebox.showwarning("Exportfehler", "Keine Datenbankverbindung vorhanden.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-Dateien", "*.xlsx")], title="Daten exportieren")
        if not file_path:
            return
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Krankmeldungen"
            columns = ("ID", "Datum", "Meldender", "Krank von", "Krank bis", "Anruf um", "Angenommen von", "Bemerkung", "KvS", "Gegangen um")
            sheet.append(columns)
            for row_id in self.manage_tree.get_children():
                sheet.append(self.manage_tree.item(row_id)['values'])
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for column in sheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_length:
                        max_length = len(val)
                sheet.column_dimensions[get_column_letter(column[0].column)].width = max_length + 2
            workbook.save(file_path)
            messagebox.showinfo("Export erfolgreich", f"Daten wurden erfolgreich nach '{file_path}' exportiert.")
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Ein Fehler ist beim Exportieren der Daten aufgetreten: {e}")

    def parse_date_de(self, date_str):
        try:
            return datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            return None

# -------------------------------------------------------------------
# Excel-Viewer-Tab (Tagesdienstplan)
# -------------------------------------------------------------------

class ExcelViewerTab(tk.Frame):
    def __init__(self, master=None, db=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.db = db
        self.tree = None
        self.df = None
        self.file_path = None
        self.original_file_path = None
        self.header_row_index = None
        self.edit_entry = None
        self.edit_item = None
        self.edit_column = None
        self.kvs_time_map = {}  # Nachname -> "KvS HH:MM (von DD.MM.YYYY)"
        self.create_widgets()

    def update_treeview_from_df(self):
        
        # Tree leeren
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.df is None or self.df.empty:
            return

        # Spalte 'deactivated' sicherstellen
        if 'deactivated' not in self.df.columns:
            self.df['deactivated'] = False

        # Einträge aufbauen
        for index, row in self.df.iterrows():
            # Stelle sicher, dass erwartete Spalten existieren
            for col in ['Datum','Namen','Dienst','Beginn','Ende']:
                if col not in self.df.columns:
                    return

            row_values = [row.get('Datum',''),
                          row.get('Namen',''),
                          row.get('Dienst',''),
                          row.get('Beginn',''),
                          row.get('Ende','')]

            # Zeiten schön formatiert (ohne Sekunden)
            row_values[3] = format_time(row_values[3])
            row_values[4] = format_time(row_values[4])

            # KvS-Text anhängen
            try:
                lastname = extract_lastname(row.get('Namen',''))
            except Exception:
                lastname = ''
            kvs_text = getattr(self, 'kvs_time_map', {}).get(lastname, "")
            row_values.append(kvs_text)  # -> Spalte 'KvS'

            # Tags bestimmen
            tags = []
            if 'is_sick' in self.df.columns and bool(row.get('is_sick', False)):
                tags.append('sick_employee')
            if bool(row.get('deactivated', False)) or (index in getattr(self, 'deactivated_set', set())):
                tags.append('deactivated')

            # Zeile einfügen
            try:
                self.tree.insert("", "end", values=row_values, iid=index, tags=tuple(tags))
            except Exception:
                # Fallback ohne iid
                self.tree.insert("", "end", values=row_values, tags=tuple(tags))



    def create_widgets(self):
        frame = tk.Frame(self)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        h_scrollbar = ttk.Scrollbar(frame, orient="horizontal")
        v_scrollbar = ttk.Scrollbar(frame, orient="vertical")
        
        self.tree = ttk.Treeview(frame, xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set, selectmode='browse')
        # Configure columns & headings for the day plan
        try:
            self.tree['columns'] = ('Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS')
            self.tree['show'] = 'headings'
            for col, width in [('Datum', 110), ('Namen', 220), ('Dienst', 80), ('Beginn', 80), ('Ende', 80), ('KvS', 80)]:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=width, anchor='center')
        except Exception:
            pass
        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.tag_configure('sick_employee', foreground='red')

        # ---- Deaktivieren: Kontextmenü und Tag ----
        self.deactivated_set = set()
        self.tree.tag_configure('deactivated', foreground='#9E9E9E')
        self.ctx_menu = tk.Menu(self.tree, tearoff=0)
        self.ctx_menu.add_command(label="Mitarbeiter deaktivieren", command=self._ctx_deactivate)
        self.ctx_menu.add_command(label="Mitarbeiter wieder aktivieren", command=self._ctx_activate)
        # Rechtsklick-Binding
        self.tree.bind("<Button-3>", self._on_right_click)

        h_scrollbar.config(command=self.tree.xview)
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.pack(side='bottom', fill='x')
        v_scrollbar.pack(side='right', fill='y')
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
        container_bg = '#E0F7FA'
        btn_width = 25
        btn_font = ("Segoe UI", 10, "bold")
        btn_bg = "#B2DFDB"
        btn_fg = "#37474F"

        button_frame = tk.Frame(self, bg=container_bg)
        button_frame.pack(pady=10, fill="x")

        word_frame = tk.Frame(button_frame, bg=container_bg)
        word_frame.pack(side=tk.TOP, pady=5)

        self.word_button = tk.Button(
            word_frame, text="📄 Word-Dokument erstellen",
            command=self.create_word_from_current_plan,  # <- Methode ist unten vorhanden
            state=tk.DISABLED,
            width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge"
        )
        self.word_button.pack(side=tk.LEFT, padx=5)

        email_frame = tk.Frame(button_frame, bg=container_bg)
        email_frame.pack(side=tk.TOP, pady=5)

        self.email_button = tk.Button(
            email_frame, text="✉️ E-Mail-Entwurf (mit Word)",
            command=self.email_draft_from_current_plan,
            state=tk.DISABLED,
            width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge"
        )
        self.email_button.pack(side=tk.LEFT, padx=5)

        tools_frame = tk.Frame(button_frame, bg=container_bg)
        tools_frame.pack(side=tk.TOP, pady=5)

        self.mark_sick_button = tk.Button(tools_frame, text="Krankmeldungen markieren",
                                          command=self.mark_sick_employees, state=tk.DISABLED,
                                          width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.mark_sick_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(tools_frame, text="Daten in Datenbank speichern",
                                     command=self.save_to_db_click, state=tk.DISABLED,
                                     width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.export_button = tk.Button(tools_frame, text="Exportieren",
                                       command=self.export_to_excel, state=tk.DISABLED,
                                       width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.export_button.pack(side=tk.LEFT, padx=5)

    def load_data(self, file_path):
        self.file_path = file_path
        self.original_file_path = file_path
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)

            df_raw = pd.read_excel(file_path, header=None, engine="openpyxl")
            required_columns = ["Name", "Dienst", "Beginn", "Ende"]
            self.header_row_index = None

            for i in range(15):
                if set(required_columns).issubset(set(df_raw.iloc[i].dropna().tolist())):
                    self.header_row_index = i
                    break
            
            if self.header_row_index is None:
                messagebox.showerror("Fehler", "Die benötigten Spaltenüberschriften (Name, Dienst, Beginn, Ende) konnten in den ersten 15 Zeilen nicht gefunden werden.")
                self.df = None
                for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                    b.config(state=tk.DISABLED)
                return False

            df = pd.read_excel(file_path, header=self.header_row_index, engine="openpyxl")
            df = df[['Name', 'Dienst', 'Beginn', 'Ende']]
            df.dropna(subset=['Name', 'Dienst', 'Beginn', 'Ende'], inplace=True)
            df.rename(columns={'Name': 'Namen'}, inplace=True)

            df['OrigDienst'] = df['Dienst']
            df['OrigEnde'] = df['Ende']
            df['is_sick'] = False

            df['is_sick'] = df['is_sick'] | df['Dienst'].astype(str).str.strip().str.lower().str.startswith(('k', 'krank'))
            df.loc[df['is_sick'], 'Dienst'] = 'K'

            filename = os.path.basename(file_path)
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
            file_date_str = date_match.group(1) if date_match else "Unbekanntes Datum"
            df.insert(0, 'Datum', file_date_str)
            
            self.df = df
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            
            columns = ['Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS']
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=110 if col in ('Beginn', 'Ende', 'KvS') else 150 if col == 'Namen' else 100, anchor="w")
                
            self.update_treeview_from_df()

            if self.db and self.db.conn:
                self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)

            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.NORMAL)

            if self.db and self.db.conn:
                self.recompute_sickness_from_db()
            return True
        except KeyError as e:
            self.df = None
            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.DISABLED)
            messagebox.showerror("Fehler beim Laden der Datei", f"Die Datei enthält nicht die erforderlichen Spalten. Es fehlt: {e}")
            return False
        except Exception as e:
            self.df = None
            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.DISABLED)
            messagebox.showerror("Fehler beim Laden der Datei", f"Ein Fehler ist aufgetreten: {e}\n\nStellen Sie sicher, dass die Datei ein gültiges Excel-Format hat.")
            return False

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        self.edit_item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        self.edit_column = int(column_id[1:]) - 1
        if self.edit_column in [1, 2]:
            x, y, width, height = self.tree.bbox(self.edit_item, column_id)
            current_value = self.tree.item(self.edit_item, "values")[self.edit_column]
            self.edit_entry = ttk.Entry(self.tree)
            self.edit_entry.place(x=x, y=y, width=width, height=height)
            self.edit_entry.insert(0, current_value)
            self.edit_entry.focus_set()
            self.edit_entry.bind("<Return>", self.on_edit_finish)
            self.edit_entry.bind("<FocusOut>", self.on_edit_finish)
        else:
            messagebox.showwarning("Bearbeitung nicht möglich", "Diese Spalte kann nicht direkt im Viewer bearbeitet werden.")

    def on_edit_finish(self, event):
        if not self.edit_entry:
            return
        new_value = self.edit_entry.get()
        row_index = int(self.edit_item)
        column_name = self.tree['columns'][self.edit_column]
        if column_name in self.df.columns:
            self.df.loc[row_index, column_name] = new_value
            current_values = list(self.tree.item(self.edit_item, "values"))
            current_values[self.edit_column] = new_value
            self.tree.item(self.edit_item, values=current_values)
        self.edit_entry.destroy(); self.edit_entry = None; self.edit_item = None; self.edit_column = None

        try:
            if self.db and self.db.conn and self.df is not None:
                self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)
        except Exception:
            pass

    def save_to_db_click(self):
        if self.db and self.db.conn and self.df is not None:
            self.db.save_dienstplan_from_df(self.df, show_message=True, replace_for_date=True)
        else:
            messagebox.showwarning("Fehler", "Keine gültigen Daten zum Speichern oder keine aktive Datenbankverbindung.")
    
    def recompute_sickness_from_db(self):
        if not self.db or not self.db.conn or self.df is None:
            return
        try:
            filename = os.path.basename(self.file_path) if self.file_path else ""
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
            plan_date = datetime.strptime(date_match.group(1), "%d.%m.%Y") if date_match else datetime.today()
            end_date = plan_date.replace(hour=23, minute=59, second=59, microsecond=999999)

            sick_employees = self.db.fetch_sick_employees_in_range(plan_date, end_date)
            sick_lastnames = {extract_lastname(name) for name, _, _ in sick_employees}

            kvs_rows = self.db.fetch_kvs_for_date(plan_date)
            kvs_map = {}
            for name, t_gegangen, d_von, d_bis in kvs_rows:
                if d_von is None or d_bis is None:
                    continue
                key = extract_lastname(name)
                hhmm = ""
                if t_gegangen:
                    try:
                        if isinstance(t_gegangen, datetime):
                            hhmm = f"{t_gegangen.hour:02d}:{t_gegangen.minute:02d}"
                        else:
                            hh, mm = map(int, str(t_gegangen).split(":")[:2])
                            hhmm = f"{hh:02d}:{mm:02d}"
                    except Exception:
                        hhmm = ""
                try:
                    von_dt = pd.to_datetime(d_von)
                    von_str = von_dt.strftime("%d.%m.%Y")
                except Exception:
                    continue
                kvs_map[key] = f"KvS {hhmm} (von {von_str})" if hhmm else f"KvS (von {von_str})"

            self.kvs_time_map = kvs_map

            self.df['Dienst'] = self.df['OrigDienst']
            self.df['Ende']   = self.df['OrigEnde']

            excel_sick = self.df['Dienst'].astype(str).str.strip().str.lower().str.startswith(('k', 'krank'))
            self.df['is_sick'] = excel_sick.copy()
            self.df.loc[self.df['is_sick'], 'Dienst'] = 'K'

            for index, row in self.df.iterrows():
                lastname = extract_lastname(row['Namen'])
                if lastname in sick_lastnames:
                    self.df.at[index, 'is_sick'] = True
                    self.df.at[index, 'Dienst'] = 'K'

            self.update_treeview_from_df()

            try:
                if self.db and self.db.conn and self.df is not None:
                    self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)
            except Exception:
                pass

            try:
                app = self.master.master
                if hasattr(app, "statistics_tab") and app.statistics_tab:
                    app.statistics_tab.update_statistics(self.df)
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Fehler beim Aktualisieren", f"Krank-Status/KvS konnte nicht neu berechnet werden: {e}")

    def mark_sick_employees(self):
        if not self.db or not self.db.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung vorhanden.")
            return
        if self.df is None:
            messagebox.showwarning("Fehler", "Bitte laden Sie zuerst eine Excel-Datei.")
            return
        if 'OrigDienst' not in self.df.columns:
            self.df['OrigDienst'] = self.df['Dienst']
        if 'OrigEnde' not in self.df.columns:
            self.df['OrigEnde'] = self.df['Ende']
        self.recompute_sickness_from_db()
        messagebox.showinfo("Erfolg", "Dienstplan wurde mit den aktuellen Krankmeldungen abgeglichen (KvS-Kommentar im Zeitraum).")

    
       
    def _tree_row_iid_at(self, event):
        row_iid = self.tree.identify_row(event.y)
        return row_iid if row_iid != "" else None

    def _on_right_click(self, event):
        iid = self._tree_row_iid_at(event)
        if iid is None:
            return
        self.tree.selection_set(iid)
        self._last_rightclick_iid = iid
        idx = int(iid)
        if idx in getattr(self, 'deactivated_set', set()):
            self.ctx_menu.entryconfig(0, state="disabled")
            self.ctx_menu.entryconfig(1, state="normal")
        else:
            self.ctx_menu.entryconfig(0, state="normal")
            self.ctx_menu.entryconfig(1, state="disabled")
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _ctx_deactivate(self):
        iid = getattr(self, "_last_rightclick_iid", None)
        if iid is None:
            return
        idx = int(iid)
        if not hasattr(self, 'deactivated_set'):
            self.deactivated_set = set()
        self.deactivated_set.add(idx)
        if self.df is not None:
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            self.df.at[idx, 'deactivated'] = True
        self.update_treeview_from_df()
        try:
            app = self.master.master
            if hasattr(app, "statistics_tab") and app.statistics_tab:
                app.statistics_tab.update_statistics(self.df)
        except Exception:
            pass

    def _ctx_activate(self):
        iid = getattr(self, "_last_rightclick_iid", None)
        if iid is None:
            return
        idx = int(iid)
        if not hasattr(self, 'deactivated_set'):
            self.deactivated_set = set()
        if idx in self.deactivated_set:
            self.deactivated_set.discard(idx)
        if self.df is not None:
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            self.df.at[idx, 'deactivated'] = False
        self.update_treeview_from_df()
        try:
            app = self.master.master
            if hasattr(app, "statistics_tab") and app.statistics_tab:
                app.statistics_tab.update_statistics(self.df)
        except Exception:
            pass
    
    def export_to_excel(self):
        if self.df is None:
            messagebox.showwarning("Fehler", "Bitte laden Sie zuerst eine Excel-Datei.")
            return
        save_path = filedialog.asksaveasfilename(title="Dienstplan speichern unter", defaultextension=".xlsx", filetypes=[("Excel-Dateien", "*.xlsx")])
        if not save_path:
            return
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Dienstplan"
            headers = ['Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS']
            sheet.append(headers)
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for idx, row in self.df.iterrows():
                beg = format_time(row['Beginn'])
                end = format_time(row['Ende'])
                lastname = extract_lastname(row['Namen'])
                kvs_text = self.kvs_time_map.get(lastname, "")
                sheet.append([row['Datum'], row['Namen'], row['Dienst'], beg, end, kvs_text])
                if 'is_sick' in self.df.columns and self.df.loc[idx, 'is_sick']:
                    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    for cell in sheet[sheet.max_row]:
                        cell.fill = red_fill
            for column in sheet.columns:
                max_length = 0
                column_cells = [cell for cell in column]
                for cell in column_cells:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_length:
                        max_length = len(val)
                sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = max_length + 2
            workbook.save(save_path)
            messagebox.showinfo("Export erfolgreich", f"Daten wurden erfolgreich nach '{save_path}' exportiert.")
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Ein Fehler ist beim Exportieren aufgetreten: {e}")

    # -------- Word & E-Mail aus aktuellem Plan --------
    def create_word_from_current_plan(self, save_path=None, return_path=False):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Hinweis", "Kein Tagesdienstplan geladen.")
            return None if return_path else None

        if not save_path:
            save_path = filedialog.asksaveasfilename(
                title="Word-Dokument speichern",
                defaultextension=".docx",
                filetypes=[("Word-Dokumente", "*.docx")]
            )
            if not save_path:
                return None if return_path else None

        filename = os.path.basename(self.file_path) if self.file_path else ""
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        if date_match:
            try:
                plan_date = datetime.strptime(date_match.group(1), "%d.%m.%Y")
            except Exception:
                plan_date = datetime.today()
        else:
            plan_date = datetime.today()

        # --- Datumsbereich mit Kalender wählen ---
        try:
            from tkcalendar import Calendar
            _has_cal = True
        except Exception:
            _has_cal = False

        header_start = plan_date + timedelta(days=CONFIG.get('word_display_date_start_offset_days', -1))
        header_end   = plan_date + timedelta(days=CONFIG.get('word_display_date_end_offset_days', 0))

        if _has_cal:
            class _DateRangeDialog(tk.Toplevel):
                def __init__(self, parent, start_date, end_date):
                    super().__init__(parent)
                    self.title('Datumsbereich wählen')
                    self.transient(parent)
                    self.grab_set()
                    self.resizable(False, False)
                    frm = ttk.Frame(self, padding=10); frm.pack(fill='both', expand=True)
                    ttk.Label(frm, text='Start- und Enddatum wählen').grid(row=0, column=0, columnspan=2, pady=(0,8))
                    self.cal_start = Calendar(frm, selectmode='day', date_pattern='dd.mm.yyyy')
                    self.cal_end   = Calendar(frm, selectmode='day', date_pattern='dd.mm.yyyy')
                    self.cal_start.grid(row=1, column=0, padx=(0,8)); self.cal_end.grid(row=1, column=1, padx=(8,0))
                    try:
                        self.cal_start.selection_set(start_date)
                        self.cal_end.selection_set(end_date)
                    except Exception:
                        pass
                    btns = ttk.Frame(frm); btns.grid(row=2, column=0, columnspan=2, pady=(10,0))
                    ttk.Button(btns, text='OK', command=self._ok).pack(side='left', padx=(0,8))
                    ttk.Button(btns, text='Abbrechen', command=self._cancel).pack(side='left')
                    self.result = None
                def _ok(self):
                    s = self.cal_start.selection_get(); e = self.cal_end.selection_get()
                    if e < s:
                        messagebox.showerror('Ungültiger Bereich', 'Enddatum darf nicht vor dem Startdatum liegen.')
                        return
                    self.result = (s, e); self.destroy()
                def _cancel(self):
                    self.result = None; self.destroy()

            dlg = _DateRangeDialog(self, header_start, header_end)
            try:
                self.wait_window(dlg)
            except Exception:
                pass
            if not getattr(dlg, 'result', None):
                return None if return_path else None
            header_start, header_end = dlg.result
            # zu datetime umwandeln, falls date
            if hasattr(header_start, 'year') and not hasattr(header_start, 'hour'):
                header_start = datetime(header_start.year, header_start.month, header_start.day)
            if hasattr(header_end, 'year') and not hasattr(header_end, 'hour'):
                header_end = datetime(header_end.year, header_end.month, header_end.day)
        else:
            # Fallback: Texteingabe
            fmt = '%d.%m.%Y'
            s_str = simpledialog.askstring('Startdatum', 'Startdatum (TT.MM.JJJJ):', initialvalue=(header_start.strftime(fmt)), parent=self)
            if s_str is None:
                return None if return_path else None
            e_str = simpledialog.askstring('Enddatum', 'Enddatum (TT.MM.JJJJ):', initialvalue=(header_end.strftime(fmt)), parent=self)
            if e_str is None:
                return None if return_path else None
            try:
                header_start = datetime.strptime(s_str.strip(), fmt)
                header_end   = datetime.strptime(e_str.strip(), fmt)
            except Exception:
                messagebox.showerror('Ungültiges Datum', 'Bitte Daten im Format TT.MM.JJJJ eingeben.')
                return None if return_path else None
            if header_end < header_start:
                messagebox.showerror('Ungültiger Bereich', 'Enddatum darf nicht vor dem Startdatum liegen.')
                return None if return_path else None

        
        # --- Vorschau: Plan vs Export (mit Bearbeitung) ---
        def _detect_name_col(_df):
            for c in ('Namen','Name','Mitarbeiter','Mitarbeiter_Name','Nachname'):
                if c in _df.columns:
                    return c
            return _df.columns[0] if len(_df.columns) else 'Name'
        name_col_detected = _detect_name_col(self.df)

        def _row_to_pair(r):
            name = ''
            try:
                name = extract_lastname(r['Namen']) if 'Namen' in self.df.columns else str(r.get(name_col_detected,''))
            except Exception:
                name = str(r.get(name_col_detected,''))
            dienst = str(r.get('Dienst','')) if 'Dienst' in self.df.columns else ''
            return name.strip(), dienst.strip()

        # Ausgangslisten bilden
        plan_pairs = []
        for _, r in self.df.iterrows():
            nm, ds = _row_to_pair(r)
            if nm:
                plan_pairs.append((nm, ds))
        # Exportliste nach Export-Regeln initialisieren
        # F2-Mapping für Plan: merke pro (Nachname, Dienst) die Reihenfolge der Initialen
        plan_pair2f2 = {}
        for _, r in self.df.iterrows():
            nm, ds = _row_to_pair(r)
            if not nm:
                continue
            f2 = extract_first2(r.get('Namen', '')) if 'Namen' in self.df.columns else ''
            plan_pair2f2.setdefault((nm, ds), []).append(f2)

        export_pairs = []
        for _, r in self.df.iterrows():
            if 'deactivated' in self.df.columns and bool(r.get('deactivated', False)):
                continue
            nm, ds = _row_to_pair(r)
            if not nm:
                continue
            dienst_raw = ds.lower()
            if any(ex in dienst_raw for ex in CONFIG.get('excluded_services', [])):
                continue
            if nm in CONFIG.get('excluded_names', []):
                continue
            export_pairs.append((nm, ds))

        # Dialog
        # F2-Mapping für Export (gleiche Filter wie Exportliste)
        export_pair2f2 = {}
        for _, r in self.df.iterrows():
            nm, ds = _row_to_pair(r)
            if not nm:
                continue
            dienst_raw = ds.lower() if isinstance(ds, str) else ''
            if any(ex in dienst_raw for ex in CONFIG.get('excluded_services', [])):
                continue
            if nm in CONFIG.get('excluded_names', []):
                continue
            f2 = extract_first2(r.get('Namen', '')) if 'Namen' in self.df.columns else ''
            export_pair2f2.setdefault((nm, ds), []).append(f2)

        prev = tk.Toplevel(self)
        prev.title('Vorschau: Plan vs Word-Export')
        prev.transient(self)
        prev.grab_set()
        prev.geometry('900x520')
        frm = ttk.Frame(prev, padding=8)
        frm.pack(fill='both', expand=True)

        # 3-Spalten-Layout
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_columnconfigure(1, weight=1)
        frm.grid_columnconfigure(2, weight=1)
        frm.grid_rowconfigure(1, weight=1)

        ttk.Label(frm, text='Tagesdienstplan').grid(row=0, column=0, sticky='w')
        ttk.Label(frm, text='Abweichungen').grid(row=0, column=1, sticky='w')
        ttk.Label(frm, text='Word-Export').grid(row=0, column=2, sticky='w')

        def _mk_tv(parent):
            tv = ttk.Treeview(parent, columns=('Name','Dienst'), show='headings', height=14)
            tv.heading('Name', text='Name'); tv.column('Name', width=200, anchor='w')
            tv.heading('Dienst', text='Dienst'); tv.column('Dienst', width=140, anchor='w')
            vs = ttk.Scrollbar(parent, orient='vertical', command=tv.yview)
            tv.configure(yscrollcommand=vs.set)
            tv.grid(row=1, column=0, sticky='nsew')
            vs.grid(row=1, column=1, sticky='ns')
            return tv

        def _fill_tv_disambig(tv, pairs, pair2f2):
            from collections import Counter, defaultdict
            for it in tv.get_children(): tv.delete(it)
            counts = Counter([nm for nm,_ in pairs])
            used = defaultdict(int)
            for i, (nm, ds) in enumerate(pairs):
                f2_list = pair2f2.get((nm, ds), []) if pair2f2 else []
                idx = used[(nm, ds)]
                f2 = f2_list[idx] if idx < len(f2_list) else (f2_list[0] if f2_list else '')
                used[(nm, ds)] += 1
                disp = f"{nm} {f2}" if counts.get(nm, 0) > 1 and f2 else nm
                iid = f"{nm}|||{ds}|||{f2}|||{i}"
                tv.insert('', 'end', iid=iid, values=(disp, ds))

        left_frame = ttk.Frame(frm)
        left_frame.grid(row=1, column=0, sticky='nsew', padx=(0,6))
        mid_frame  = ttk.Frame(frm)
        mid_frame.grid(row=1, column=1, sticky='nsew', padx=6)
        right_frame= ttk.Frame(frm)
        right_frame.grid(row=1, column=2, sticky='nsew', padx=(6,0))

        tv_left  = _mk_tv(left_frame)
        tv_right = _mk_tv(right_frame)

        def _fill_tv(tv, pairs):
            for it in tv.get_children():
                tv.delete(it)
            for nm, ds in pairs:
                tv.insert('', 'end', values=(nm, ds))

        _fill_tv_disambig(tv_left, plan_pairs, plan_pair2f2)
        _fill_tv_disambig(tv_right, export_pairs, export_pair2f2)

        def _compute_sets():
            left_names  = set(nm for nm,_ in plan_pairs)
            right_names = set(nm for nm,_ in export_pairs)
            missing = sorted(left_names - right_names)
            extra   = sorted(right_names - left_names)
            return missing, extra

        ttk.Label(mid_frame, text='Im Plan, nicht im Export').grid(row=0, column=0, sticky='w')
        lb_missing = tk.Listbox(mid_frame, height=10, exportselection=False)
        lb_missing.grid(row=1, column=0, sticky='nsew')
        ttk.Label(mid_frame, text='Im Export, nicht im Plan').grid(row=2, column=0, sticky='w', pady=(8,0))
        lb_extra   = tk.Listbox(mid_frame, height=10, exportselection=False)
        lb_extra.grid(row=3, column=0, sticky='nsew')

        mid_frame.grid_rowconfigure(1, weight=1)
        mid_frame.grid_rowconfigure(3, weight=1)
        mid_frame.grid_columnconfigure(0, weight=1)

        def _refresh_diff():
            missing, extra = _compute_sets()
            lb_missing.delete(0, 'end'); lb_extra.delete(0, 'end')
            for n in missing: lb_missing.insert('end', n)
            for n in extra:   lb_extra.insert('end', n)

        _refresh_diff()

        btns_frame = ttk.Frame(frm)
        btns_frame.grid(row=2, column=2, sticky='e', pady=(8,0))

        def _selected_from(tv):
            vals = []
            for iid in tv.selection():
                v = tv.item(iid,'values')
                if v: vals.append((v[0], v[1]))
            return vals

        def _add_to_export():
            sel = _selected_from(tv_left)
            for nm, ds in sel:
                if (nm, ds) not in export_pairs:
                    export_pairs.append((nm, ds))
            _fill_tv_disambig(tv_right, export_pairs, export_pair2f2); _refresh_diff()

        def _remove_from_export():
            sel = _selected_from(tv_right)
            if not sel: return
            for pair in sel:
                try:
                    export_pairs.remove(pair)
                except ValueError:
                    export_pairs[:] = [p for p in export_pairs if p[0] != pair[0]]
            _fill_tv_disambig(tv_right, export_pairs, export_pair2f2); _refresh_diff()

        ttk.Button(btns_frame, text='<< Aus Export entfernen', command=_remove_from_export).grid(row=0, column=0, padx=6)
        ttk.Button(btns_frame, text='In Export übernehmen >>', command=_add_to_export).grid(row=0, column=1, padx=6)

        footer = ttk.Frame(frm)
        footer.grid(row=3, column=0, columnspan=3, sticky='ew', pady=(8,0))
        footer.grid_columnconfigure(0, weight=1)
        _confirmed = {'ok': False, 'names': []}
        def _ok_and_close():
            _confirmed['ok'] = True
            _confirmed['names'] = sorted(set(n for n,_ in export_pairs))
            prev.destroy()
        ttk.Button(footer, text='Abbrechen', command=lambda: prev.destroy()).pack(side='right')
        ttk.Button(footer, text='Fortfahren', command=_ok_and_close).pack(side='right', padx=(0,8))

        try:
            self.wait_window(prev)
        except Exception:
            pass
        if not _confirmed['ok']:
            return None if return_path else None
        selected_names = _confirmed['names']
        # --- Ende Vorschau ---

        paxzahl = simpledialog.askinteger(
            "PAX-Zahl eingeben",
            "Bitte geben Sie die PAX-Zahl für den Dienstplan ein:",
            minvalue=0, parent=self
        )
        if paxzahl is None:
            return None if return_path else None

        try:
            build_word_from_df(self.df, plan_date, save_path, paxzahl=paxzahl, header_date_start=header_start, header_date_end=header_end, selected_names=selected_names)
            if not return_path:
                messagebox.showinfo("Fertig", f"Dokument gespeichert unter:\n{save_path}")
                return None
            else:
                return save_path
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Erstellen des Word-Dokuments: {e}")
            return None

    def email_draft_from_current_plan(self):
        save_path = self.create_word_from_current_plan(return_path=True)
        if not save_path:
            return

        filename = os.path.basename(self.file_path) if self.file_path else ""
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        if date_match:
            try:
                d0 = datetime.strptime(date_match.group(1), "%d.%m.%Y")
                d1 = d0 + timedelta(days=1)
            except Exception:
                d0 = datetime.today(); d1 = d0 + timedelta(days=1)
        else:
            d0 = datetime.today(); d1 = d0 + timedelta(days=1)
        date_str = f"{d0.strftime('%d.%m.%Y')} – {d1.strftime('%d.%m.%Y')}"

        subject = f"Dienstplan für {date_str}"
        body = f"Hallo,\n\nanbei der Dienstplan für {date_str} als Anhang.\n\nMit freundlichen Grüßen,\nDRK Flughafen"

        try:
            if OUTLOOK_AVAILABLE:
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = CONFIG["email_recipients"]["to"]
                mail.CC = CONFIG["email_recipients"]["cc"]
                mail.Subject = subject
                mail.Body = body
                mail.Attachments.Add(os.path.abspath(save_path))
                mail.Display(True)
            else:
                body_mailto = body + f"\n\nPfad zur Datei (manuell anhängen): {save_path}"
                mailto_url = (f"mailto:?to={urllib.parse.quote(CONFIG['email_recipients']['to'])}"
                              f"&cc={urllib.parse.quote(CONFIG['email_recipients']['cc'])}"
                              f"&subject={urllib.parse.quote(subject)}"
                              f"&body={urllib.parse.quote(body_mailto)}")
                webbrowser.open(mailto_url)
        except Exception as e:
            messagebox.showerror("Fehler", f"E-Mail-Entwurf konnte nicht erstellt werden: {e}")

# -------------------------------------------------------------------
# Neuer Tab: Sonderaufgaben (neue Excel-Datei)
# -------------------------------------------------------------------

class SonderaufgabenTab(tk.Frame):
    AUFGABEN = [
        "Sauberkeit Station",
        "BTW Check + Sauberkeit",
        "E - mobby Check",
        "Bulmor 1 - 7312",
        "Bulmor 2 - 7892",
        "Bulmor 3 - 8092",
        "Bulmor 4 - 8794",
        "Bulmor 5 - 9982",
    ]

    NAME_HEADERS = {"name","namen","nachname","lastname","surname"}
    DIENST_HEADERS = {"dienst","schicht","shift"}

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        # sanfte Hintergrundfarbe wie im restlichen GUI
        try:
            self.configure(bg="#E0F7FA")
        except Exception:
            pass
        self.entries_tag = {}
        self.entries_nacht = {}
        # Zweite Combobox-Maps nur für Bulmor-Aufgaben
        self.entries_tag_2 = {}
        self.entries_nacht_2 = {}

        # Gesamtliste (Kompatibilität / manuelle Auswahl)
        self.namen_liste = []
        # Schicht-spezifische Listen
        self.namen_liste_tag = []
        self.namen_liste_nacht = []
        self.namen_liste_bulmor_tag = []
        self.namen_liste_bulmor_nacht = []

        self._build_ui()

    # --------- Color helpers (yellow detection) ---------
    def _hex_to_rgb(self, hexstr):
        s = str(hexstr or "").strip().lstrip("#")
        if len(s) == 8: s = s[2:]
        if len(s) != 6: return None
        try: return (int(s[0:2],16), int(s[2:4],16), int(s[4:6],16))
        except Exception: return None

    def _apply_tint(self, rgb, tint):
        if not rgb: return None
        r,g,b = rgb
        if tint in (None,0): return (r,g,b)
        if tint > 0:
            r = int(round(r + (255 - r) * tint))
            g = int(round(g + (255 - g) * tint))
            b = int(round(b + (255 - b) * tint))
        else:
            r = int(round(r * (1 + tint)))
            g = int(round(g * (1 + tint)))
            b = int(round(b * (1 + tint)))
        return (max(0,min(255,r)), max(0,min(255,g)), max(0,min(255,b)))

    def _color_obj_to_rgb(self, wb, color_obj):
        try:
            if not color_obj: return None
            tint = getattr(color_obj, "tint", None)
            if getattr(color_obj, "rgb", None):
                rgb = self._hex_to_rgb(color_obj.rgb)
                return self._apply_tint(rgb, tint) if rgb else None
            if getattr(color_obj, "theme", None) is not None:
                return None
            return None
        except Exception:
            return None

    def _get_effective_rgb(self, wb, fill):
        try:
            if not fill: return None
            pattern = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
            if str(pattern).lower() != "solid": return None
            for c in (getattr(fill,'fgColor',None), getattr(fill,'start_color',None), getattr(fill,'bgColor',None)):
                rgb = self._color_obj_to_rgb(wb, c)
                if rgb: return rgb
            return None
        except Exception:
            return None

    def _is_yellowish(self, rgb):
        if not rgb: return False
        r,g,b = rgb
        if r >= 180 and g >= 180 and b <= 150: return True
        try:
            import colorsys
            h,s,v = colorsys.rgb_to_hsv(r/255.0, g/255.0, b/255.0)
            deg = h*360
            return (25 <= deg <= 70) and s >= 0.25 and v >= 0.6
        except Exception:
            return False

    # --------------------------- UI ---------------------------
    
    def _build_ui(self):
        try:
            title = tk.Label(self, text="Sonderaufgaben – Eingabe (Tag/Nacht)",
                             font=("Segoe UI", 14, "bold"), bg="#E0F7FA", fg="#37474F")
            title.pack(pady=(10,5))
        except Exception:
            pass

        table = tk.Frame(self)
        try: table.configure(bg="#E0F7FA")
        except Exception: pass
        table.pack(padx=10, pady=10, fill="x")

        header_font = ("Segoe UI", 10, "bold")
        tk.Label(table, text="Aufgabe", font=header_font, anchor="w", width=35)\
            .grid(row=0, column=0, padx=5, pady=4, sticky="w")
        tk.Label(table, text="Tagschicht", font=header_font)\
            .grid(row=0, column=1, padx=5, pady=4)
        tk.Label(table, text="Nachtschicht", font=header_font)\
            .grid(row=0, column=2, padx=5, pady=4)

        for i, aufgabe in enumerate(self.AUFGABEN, start=1):
            tk.Label(table, text=aufgabe, anchor="w")\
                .grid(row=i, column=0, padx=5, pady=3, sticky="w")

            is_bul = self._is_bulmor(aufgabe)

            # Tag
            tag_cell = tk.Frame(table, bg=table.cget("bg"))
            tag_cell.grid(row=i, column=1, padx=5, pady=3, sticky="w")
            if is_bul:
                cb_t1 = ttk.Combobox(tag_cell, width=18, values=[""], state="normal")
                cb_t2 = ttk.Combobox(tag_cell, width=18, values=[""], state="normal")
                cb_t1.pack(side=tk.LEFT, padx=(0,4)); cb_t2.pack(side=tk.LEFT, padx=0)
                self.entries_tag[aufgabe] = cb_t1          # Backward-compatible primary
                self.entries_tag_2[aufgabe] = cb_t2        # Secondary
            else:
                cb_t = ttk.Combobox(tag_cell, width=22, values=[""], state="normal")
                cb_t.pack(side=tk.LEFT)
                self.entries_tag[aufgabe] = cb_t
                self.entries_tag_2[aufgabe] = None

            # Nacht
            nacht_cell = tk.Frame(table, bg=table.cget("bg"))
            nacht_cell.grid(row=i, column=2, padx=5, pady=3, sticky="w")
            if is_bul:
                cb_n1 = ttk.Combobox(nacht_cell, width=18, values=[""], state="normal")
                cb_n2 = ttk.Combobox(nacht_cell, width=18, values=[""], state="normal")
                cb_n1.pack(side=tk.LEFT, padx=(0,4)); cb_n2.pack(side=tk.LEFT, padx=0)
                self.entries_nacht[aufgabe] = cb_n1        # Backward-compatible primary
                self.entries_nacht_2[aufgabe] = cb_n2      # Secondary
            else:
                cb_n = ttk.Combobox(nacht_cell, width=22, values=[""], state="normal")
                cb_n.pack(side=tk.LEFT)
                self.entries_nacht[aufgabe] = cb_n
                self.entries_nacht_2[aufgabe] = None
        btns = tk.Frame(self)
        try: btns.configure(bg="#E0F7FA")
        except Exception: pass
        btns.pack(padx=10, pady=(0,10), fill="x")
        ttk.Button(btns, text="Excel (Namen) laden", command=self.load_names_from_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Dropdowns aktualisieren", command=self.update_name_choices_by_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Export nach Excel", command=self.export_sonderaufgaben_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Drucken", command=self.print_sonderaufgaben_excel).pack(side=tk.LEFT, padx=5)
        # Hinweistext
        hint = tk.Label(self, text="Nur gelb markierte Mitarbeiter werden als Bulmorfahrer identifiziert",
                        font=("Segoe UI", 9, "italic"),
                        bg="#E0F7FA", fg="#555555", anchor="w", justify="left")
        hint.pack(padx=10, pady=(5, 10), anchor="w")
        self._init_sa_browser(self)


    def _extract_lastname_robust(self, full_name: str) -> str:
        s = str(full_name or "").strip()
        if not s: return ""
        if "," in s: s = s.split(",",1)[0].strip()
        parts = [p for p in s.split() if p]
        if not parts: return s
        if len(parts) == 1: return parts[0]
        particles = {"von","van","de","der","den","di","da","le","du"}
        if len(parts) >= 2 and parts[-2].lower() in particles:
            return parts[-2] + " " + parts[-1]
        return parts[-1]

    def _is_bulmor(self, aufgabe: str) -> bool:
        return "bulmor" in (aufgabe or "").lower()

    
    def update_name_choices_by_task(self):
        tag_vals_all   = [""] + (self.namen_liste_tag if self.namen_liste_tag else [])
        nacht_vals_all = [""] + (self.namen_liste_nacht if self.namen_liste_nacht else [])

        tag_vals_bulmor   = [""] + (self.namen_liste_bulmor_tag if self.namen_liste_bulmor_tag else self.namen_liste_tag or [])
        nacht_vals_bulmor = [""] + (self.namen_liste_bulmor_nacht if self.namen_liste_bulmor_nacht else self.namen_liste_nacht or [])

        for aufgabe in self.AUFGABEN:
            is_bul = self._is_bulmor(aufgabe)
            vals_tag = tag_vals_bulmor if is_bul else tag_vals_all
            vals_nacht = nacht_vals_bulmor if is_bul else nacht_vals_all

            # Tag primary
            cb_t1 = self.entries_tag.get(aufgabe)
            if cb_t1 is not None:
                try: cb_t1.configure(values=vals_tag)
                except Exception: pass
            # Tag secondary (Bulmor)
            cb_t2 = self.entries_tag_2.get(aufgabe)
            if cb_t2 is not None:
                try: cb_t2.configure(values=vals_tag)
                except Exception: pass

            # Nacht primary
            cb_n1 = self.entries_nacht.get(aufgabe)
            if cb_n1 is not None:
                try: cb_n1.configure(values=vals_nacht)
                except Exception: pass
            # Nacht secondary (Bulmor)
            cb_n2 = self.entries_nacht_2.get(aufgabe)
            if cb_n2 is not None:
                try: cb_n2.configure(values=vals_nacht)
                except Exception: pass
    def _find_headers_same_row(self, ws):
        # exakte Header in der gleichen Zeile
        for r in range(1, min(ws.max_row, 60)+1):
            labels = {}
            for c in range(1, ws.max_column+1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str):
                    labels[c] = v.strip().lower()
            if not labels: continue
            name_cols = [c for c,low in labels.items() if low in self.NAME_HEADERS]
            dienst_cols = [c for c,low in labels.items() if low in self.DIENST_HEADERS]
            if name_cols and dienst_cols:
                return r, name_cols[0], dienst_cols[0]
        return None, None, None

    
    def get_task_names(self, aufgabe: str):
        """Liefert (tag_namen, nacht_namen) jeweils als Liste mit 1 oder 2 Einträgen (leere Strings werden gefiltert)."""
        res_tag, res_nacht = [], []
        cb_t1 = self.entries_tag.get(aufgabe); cb_t2 = self.entries_tag_2.get(aufgabe)
        cb_n1 = self.entries_nacht.get(aufgabe); cb_n2 = self.entries_nacht_2.get(aufgabe)
        for cb in (cb_t1, cb_t2):
            try:
                val = cb.get().strip() if cb else ""
                if val: res_tag.append(val)
            except Exception:
                pass
        for cb in (cb_n1, cb_n2):
            try:
                val = cb.get().strip() if cb else ""
                if val: res_nacht.append(val)
            except Exception:
                pass
        return res_tag, res_nacht

    
    
    def export_sonderaufgaben_to_excel(self):
            """
            Exportiert den Tab 'Sonderaufgaben' als Excel exakt im Layout der Vorlage 'Sonderaufgaben.xlsx'.
            Spalten/Zeilen lt. Template: A2=Datum, A3..=Aufgabe, C3..=Tagdienst, E3..=Nachtdienst.
            Hinweis: GUI-Hinweise werden NICHT exportiert.
            """
            from tkinter import filedialog, messagebox
            from openpyxl import load_workbook, Workbook
            import os, datetime
    
            # 1) Zielpfad abfragen
            out_path = filedialog.asksaveasfilename(
                title="Sonderaufgaben als Excel speichern",
                defaultextension=".xlsx",
                filetypes=[("Excel-Datei", "*.xlsx")]
            )
            if not out_path:
                return
    
            # 2) Vorlage suchen: bevorzugt im Skriptverzeichnis
            script_dir = os.path.dirname(os.path.abspath(__file__))
            candidates = [
                os.path.join(script_dir, "Sonderaufgaben.xlsx"),
                os.path.join(os.getcwd(), "Sonderaufgaben.xlsx"),
                os.path.expanduser("~/Sonderaufgaben.xlsx"),
            ]
            template_path = None
            for cand in candidates:
                if os.path.exists(cand):
                    template_path = cand
                    break
            # Wenn nicht gefunden: Nutzer fragen
            if not template_path:
                template_path = filedialog.askopenfilename(
                    title="Vorlage 'Sonderaufgaben.xlsx' auswählen",
                    filetypes=[("Excel-Dateien", "*.xlsx")]
                )
                if not template_path:
                    try:
                        messagebox.showwarning("Abbruch", "Keine Vorlage gewählt – Export ohne Layout wird abgebrochen.")
                    except Exception:
                        pass
                    return
    
            try:
                # 3) Vorlage öffnen (Formatierung bleibt erhalten), Werte überschreiben, unter out_path speichern
                wb = load_workbook(template_path)
                ws = wb[wb.sheetnames[0]]  # 'Sonderaufgaben'
    
                # 4) Datum in A2 setzen
                today = datetime.date.today()
                ws.cell(row=2, column=1, value=today)
    
                # 5) Aufgaben/Namen ab Zeile 3 schreiben
                start_row = 3
                for idx, aufgabe in enumerate(self.AUFGABEN, start=0):
                    r = start_row + idx
                    ws.cell(row=r, column=1, value=aufgabe)
    
                    # Tagdienst (Bulmor kann 2 Felder haben)
                    vals_tag = []
                    cb_t1 = self.entries_tag.get(aufgabe)
                    cb_t2 = getattr(self, "entries_tag_2", {}).get(aufgabe) if hasattr(self, "entries_tag_2") else None
                    for cb in (cb_t1, cb_t2):
                        if cb:
                            try:
                                v = cb.get().strip()
                                if v: vals_tag.append(v)
                            except Exception:
                                pass
                    ws.cell(row=r, column=3, value=(" / ".join(vals_tag) if vals_tag else None))
    
                    # Nachtdienst
                    vals_nacht = []
                    cb_n1 = self.entries_nacht.get(aufgabe)
                    cb_n2 = getattr(self, "entries_nacht_2", {}).get(aufgabe) if hasattr(self, "entries_nacht_2") else None
                    for cb in (cb_n1, cb_n2):
                        if cb:
                            try:
                                v = cb.get().strip()
                                if v: vals_nacht.append(v)
                            except Exception:
                                pass
                    ws.cell(row=r, column=5, value=(" / ".join(vals_nacht) if vals_nacht else None))
    
                # 6) Speichern als neue Datei (Vorlage bleibt unverändert)
                wb.save(out_path)
            except Exception as e:
                try:
                    from tkinter import messagebox as _mb
                    _mb.showerror("Export-Fehler", f"{e}\n\nVorlage: {template_path}")
                except Exception:
                    print("Export-Fehler:", e)
            return

    def print_sonderaufgaben_excel(self):
        """
        Erstellt die Sonderaufgabenliste mit der Vorlage 'Sonderaufgaben.xlsx' als TEMP-Datei
        und schickt sie – wenn möglich – direkt an den Drucker.
        Falls kein direkter Druck möglich ist, wird die Datei geöffnet bzw. der Pfad angezeigt.
        """
        from tkinter import filedialog, messagebox
        import os, sys, tempfile, datetime
        try:
            from openpyxl import load_workbook, Workbook
        except Exception as _e:
            try:
                messagebox.showerror("Druck-Fehler", f"openpyxl fehlt: {_e}\\n\\nInstalliere mit:\\n    pip install openpyxl")
            except Exception:
                pass
            return

        # Vorlage suchen (wie beim Export)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(script_dir, "Sonderaufgaben.xlsx"),
            os.path.join(os.getcwd(), "Sonderaufgaben.xlsx"),
            os.path.expanduser("~/Sonderaufgaben.xlsx"),
        ]
        template_path = None
        for cand in candidates:
            if os.path.exists(cand):
                template_path = cand
                break
        if not template_path:
            template_path = filedialog.askopenfilename(
                title="Vorlage 'Sonderaufgaben.xlsx' auswählen",
                filetypes=[("Excel-Dateien", "*.xlsx")]
            )
            if not template_path:
                try:
                    messagebox.showwarning("Abbruch", "Keine Vorlage gewählt – Druck abgebrochen.")
                except Exception:
                    pass
                return
        # Workbook aus Vorlage befüllen (genau wie beim Export)
        try:
            wb = load_workbook(template_path)
            ws = wb[wb.sheetnames[0]]

            # Datum A2
            today = datetime.date.today()
            ws.cell(row=2, column=1, value=today)

            # Aufgaben & Namen ab Zeile 3
            start_row = 3
            for idx, aufgabe in enumerate(self.AUFGABEN, start=0):
                r = start_row + idx
                ws.cell(row=r, column=1, value=aufgabe)

                # Tag (inkl. zweites Feld bei Bulmor)
                vals_tag = []
                cb_t1 = self.entries_tag.get(aufgabe)
                cb_t2 = getattr(self, "entries_tag_2", {}).get(aufgabe) if hasattr(self, "entries_tag_2") else None
                for cb in (cb_t1, cb_t2):
                    if cb:
                        try:
                            v = cb.get().strip()
                            if v: vals_tag.append(v)
                        except Exception:
                            pass
                ws.cell(row=r, column=3, value=(" / ".join(vals_tag) if vals_tag else None))

                # Nacht
                vals_n = []
                cb_n1 = self.entries_nacht.get(aufgabe)
                cb_n2 = getattr(self, "entries_nacht_2", {}).get(aufgabe) if hasattr(self, "entries_nacht_2") else None
                for cb in (cb_n1, cb_n2):
                    if cb:
                        try:
                            v = cb.get().strip()
                            if v: vals_n.append(v)
                        except Exception:
                            pass
                ws.cell(row=r, column=5, value=(" / ".join(vals_n) if vals_n else None))

            # TEMP-Datei speichern
            base = f"Sonderaufgaben_{today.strftime('%Y-%m-%d')}"
            tmpdir = tempfile.gettempdir()
            tmp_path = os.path.join(tmpdir, base + ".xlsx")
            # Sicherer Name (falls existiert)
            c = 1
            while os.path.exists(tmp_path):
                tmp_path = os.path.join(tmpdir, f"{base}_{c}.xlsx")
                c += 1
            wb.save(tmp_path)

            # Drucken (Windows bevorzugt)
            printed = False
            if sys.platform.startswith("win"):
                # Versuch 1: Excel COM
                try:
                    import win32com.client as win32
                    excel = win32.Dispatch("Excel.Application")
                    excel.Visible = False
                    wb2 = excel.Workbooks.Open(tmp_path)
                    wb2.PrintOut()
                    wb2.Close(SaveChanges=False)
                    excel.Quit()
                    printed = True
                except Exception:
                    # Versuch 2: Shell-Print via Dateiverknüpfung
                    try:
                        os.startfile(tmp_path, "print")
                        printed = True
                    except Exception:
                        printed = False

            if printed:
                return
            else:
                # Kein direkter Druck verfügbar → Datei öffnen / Hinweis zeigen
                try:
                    if sys.platform == "darwin":
                        os.system(f'open "{tmp_path}"')
                    elif sys.platform.startswith("linux"):
                        os.system(f'xdg-open "{tmp_path}"')
                    else:
                        os.startfile(tmp_path)
                except Exception:
                    pass
        except Exception as e:
            try:
                messagebox.showerror("Druck-Fehler", f"{e}")
            except Exception:
                print("Druck-Fehler:", e)

    def load_names_from_excel(self):
            # openpyxl optional und klar melden, wenn fehlend
            try:
                from openpyxl import load_workbook
                _OPENPY_AVAILABLE = True
            except Exception as _e:
                _OPENPY_AVAILABLE = False
                try:
                    from tkinter import messagebox as _mb
                    _mb.showerror("Fehlender Baustein", f"openpyxl ist nicht installiert:\n{_e}\n\nInstalliere mit:\n    pip install openpyxl")
                except Exception:
                    pass
                return
    
            path = filedialog.askopenfilename(
                title="Excel mit Namens- und Dienstspalte wählen",
                filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
            )
            if not path: return
    
            try:
                wb = load_workbook(path, data_only=True)
                # erstes brauchbares Sheet
                ws = None
                for sn in wb.sheetnames:
                    if wb[sn].max_row > 1 and wb[sn].max_column > 1:
                        ws = wb[sn]; break
                if ws is None:
                    try:
                        messagebox.showwarning("Hinweis", "Keine verwertbaren Datenblätter in der Datei gefunden.")
                    except Exception:
                        pass
                    return
    
                header_row, name_col, dienst_col = self._find_headers_same_row(ws)
                if not (header_row and name_col and dienst_col):
                    try:
                        messagebox.showwarning("Hinweis", "Ich habe keine passenden Spalten für Name und Dienst (gleiche Zeile) gefunden.")
                    except Exception:
                        pass
                    return
    
                def shift_from_dienst(val: str):
                    s = str(val or "").strip().lower()
                    if s.startswith(("t","t10","t8")): return "tag"
                    if s.startswith(("n","n10","nf")): return "nacht"
                    if s[:1] == "t": return "tag"
                    if s[:1] == "n": return "nacht"
                    return None
    
                tag_set, nacht_set = set(), set()
                bul_tag, bul_nacht = set(), set()
                for rr in range(header_row+1, ws.max_row+1):
                    nval = ws.cell(row=rr, column=name_col).value
                    if not nval: continue
                    dval = ws.cell(row=rr, column=dienst_col).value
                    schicht = shift_from_dienst(dval)
                    last = self._extract_lastname_robust(str(nval))
    
                    if schicht == "tag": tag_set.add(last)
                    elif schicht == "nacht": nacht_set.add(last)
    
                    fill = ws.cell(row=rr, column=dienst_col).fill
                    rgb = self._get_effective_rgb(wb, fill)
                    if self._is_yellowish(rgb):
                        if schicht == "tag": bul_tag.add(last)
                        elif schicht == "nacht": bul_nacht.add(last)
    
                self.namen_liste_tag = sorted(tag_set, key=lambda s: s.lower())
                self.namen_liste_nacht = sorted(nacht_set, key=lambda s: s.lower())
                self.namen_liste_bulmor_tag = sorted(bul_tag, key=lambda s: s.lower()) if bul_tag else self.namen_liste_tag[:]
                self.namen_liste_bulmor_nacht = sorted(bul_nacht, key=lambda s: s.lower()) if bul_nacht else self.namen_liste_nacht[:]
                self.namen_liste = sorted(set(self.namen_liste_tag + self.namen_liste_nacht), key=lambda s: s.lower())
    
                self.update_name_choices_by_task()
                try:
                        vals = [""] + (self.namen_liste_tag if getattr(self, 'namen_liste_tag', None) else [])
                except Exception:
                    pass
            except Exception as e:
                import traceback, os
                tb = traceback.format_exc()
                log_path = os.path.join(os.path.dirname(__file__), "sonderaufgaben_error.log")
                try:
                    with open(log_path, "w", encoding="utf-8") as f:
                        f.write(tb)
                except Exception:
                    pass
                try:
                    messagebox.showerror("Fehler beim Laden", f"{e}\n\nDetails in:\n{log_path}")
                except Exception:
                    print(tb)


    # ===== Explorer (Sonderaufgaben) unter SA_ROOT_DIR =====
    def _init_sa_browser(self, parent):
        import os
        wrap = tk.LabelFrame(parent, text="Explorer (Sonderaufgaben)", bg="#E0F7FA", fg="#37474F")
        wrap.pack(padx=10, pady=(0,10), fill="both", expand=True)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_columnconfigure(1, weight=2)
        wrap.grid_rowconfigure(0, weight=1)

        left = tk.Frame(wrap, bg="#E0F7FA")
        left.grid(row=0, column=0, sticky="nsew", padx=(6,3), pady=6)
        tv_dirs = ttk.Treeview(left, columns=("abspath",), show="tree")
        vs_l = ttk.Scrollbar(left, orient="vertical", command=tv_dirs.yview)
        tv_dirs.configure(yscrollcommand=vs_l.set)
        tv_dirs.pack(side="left", fill="both", expand=True)
        vs_l.pack(side="left", fill="y")
        self._sa_tv_dirs = tv_dirs

        right = tk.Frame(wrap, bg="#E0F7FA")
        right.grid(row=0, column=1, sticky="nsew", padx=(3,6), pady=6)
        tv_files = ttk.Treeview(right, columns=("name","size","mtime","abspath"), show="headings", selectmode="extended")
        tv_files.heading("name", text="Datei"); tv_files.column("name", width=300, anchor="w")
        tv_files.heading("size", text="Größe"); tv_files.column("size", width=80, anchor="e")
        tv_files.heading("mtime", text="Geändert"); tv_files.column("mtime", width=140, anchor="w")
        tv_files["displaycolumns"] = ("name","size","mtime")
        vs_r = ttk.Scrollbar(right, orient="vertical", command=tv_files.yview)
        tv_files.configure(yscrollcommand=vs_r.set)
        tv_files.pack(side="left", fill="both", expand=True)
        vs_r.pack(side="left", fill="y")
        self._sa_tv_files = tv_files

        sel_frame = tk.Frame(wrap, bg="#E0F7FA")
        sel_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,6))
        tk.Button(sel_frame, text="Zur Auswahl hinzufügen", command=self._sa_add_selected).pack(side="left")
        tk.Button(sel_frame, text="Aus Auswahl entfernen", command=self._sa_remove_selected).pack(side="left", padx=6)
        tk.Button(sel_frame, text="Auswahl laden (Namen)", command=self._sa_load_selected).pack(side="left", padx=6)

        self._sa_selected_files = []
        self._sa_sel_list = tk.Listbox(sel_frame, height=3, selectmode="extended")
        self._sa_sel_list.pack(side="left", fill="x", expand=True, padx=(10,0))

        root = SA_ROOT_DIR
        self._sa_root_dir = root
        root_text = os.path.basename(root.rstrip("\\/")) or root
        root_node = tv_dirs.insert("", "end", text=root_text, values=(root,), open=True)
        tv_dirs.insert(root_node, "end", text="...", values=(os.path.join(root, "__dummy__"),))

        def on_open(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            self._populate_sa_dir(node)

        def on_select_dir(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            abspath = tv_dirs.set(node, "abspath")
            self._list_sa_files(abspath)

        tv_dirs.bind("<<TreeviewOpen>>", on_open)
        tv_dirs.bind("<<TreeviewSelect>>", on_select_dir)

        def on_file_double(evt):
            iid = tv_files.focus()
            if not iid: return
            abspath = tv_files.set(iid, "abspath")
            self._load_names_from_excel_path(abspath)
        tv_files.bind("<Double-1>", on_file_double)

        self._populate_sa_dir(root_node)
        self._list_sa_files(root)

    def _is_within_sa_root(self, path):
        import os
        root = os.path.abspath(self._sa_root_dir)
        try:
            ap = os.path.abspath(path)
        except Exception:
            return False
        return ap.startswith(root)

    def _populate_sa_dir(self, node):
        import os
        tv = self._sa_tv_dirs
        path = tv.set(node, "abspath")
        if not self._is_within_sa_root(path):
            return
        for k in tv.get_children(node):
            if tv.item(k, "text") == "...":
                tv.delete(k)
        try:
            entries = sorted([d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))])
        except Exception:
            entries = []
        for d in entries:
            abspath = os.path.join(path, d)
            if not self._is_within_sa_root(abspath):
                continue
            exists = any(tv.set(c, "abspath") == abspath for c in tv.get_children(node))
            if exists:
                continue
            new = tv.insert(node, "end", text=d, values=(abspath,))
            try:
                if any(os.path.isdir(os.path.join(abspath, x)) for x in os.listdir(abspath)):
                    tv.insert(new, "end", text="...", values=(os.path.join(abspath, "__dummy__"),))
            except Exception:
                pass

    def _list_sa_files(self, dir_path):
        import os, datetime
        tvf = self._sa_tv_files
        for c in tvf.get_children(): tvf.delete(c)
        if not self._is_within_sa_root(dir_path):
            return
        try:
            items = os.listdir(dir_path)
        except Exception:
            items = []
        exts = (".xlsx", ".xlsm", ".xls")
        files = []
        for nm in items:
            abspath = os.path.join(dir_path, nm)
            if os.path.isfile(abspath) and nm.lower().endswith(exts):
                try:
                    st = os.stat(abspath)
                    size = st.st_size
                    mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M")
                except Exception:
                    size = 0; mtime = ""
                files.append((nm, size, mtime, abspath))
        for nm, size, mtime, abspath in sorted(files):
            tvf.insert("", "end", values=(nm, f"{size//1024} KB", mtime, abspath))

    def _sa_add_selected(self):
        tvf = self._sa_tv_files
        sel = tvf.selection()
        for iid in sel:
            abspath = tvf.set(iid, "abspath")
            if abspath and abspath not in self._sa_selected_files:
                self._sa_selected_files.append(abspath)
                self._sa_sel_list.insert("end", abspath)

    def _sa_remove_selected(self):
        idxs = list(self._sa_sel_list.curselection())[::-1]
        for i in idxs:
            path = self._sa_sel_list.get(i)
            self._sa_sel_list.delete(i)
            try:
                self._sa_selected_files.remove(path)
            except Exception:
                pass

    def _sa_load_selected(self):
        if not self._sa_selected_files:
            try:
                messagebox.showinfo("Auswahl", "Bitte Dateien rechts wählen und hinzufügen.")
            except Exception:
                pass
            return
        self._load_names_from_excel_path(self._sa_selected_files[0])

    def _load_names_from_excel_path(self, path):
        # Trick: Filedialog kurz überschreiben, damit load_names_from_excel(path) nutzt
        try:
            from tkinter import filedialog as _fd
            orig = _fd.askopenfilename
            _fd.askopenfilename = lambda **kw: path
            try:
                self.load_names_from_excel()
            finally:
                _fd.askopenfilename = orig
        except Exception as e:
            try:
                messagebox.showerror("Ladefehler", f"{e}")
            except Exception:
                print("Ladefehler:", e)
class StatisticsTab(tk.Frame):
    def __init__(self, master=None, style_vars=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        if style_vars:
            self.pastel_bg = style_vars.get('bg')
            self.pastel_fg = style_vars.get('fg')
        else:
            self.pastel_bg = '#E0F7FA'
            self.pastel_fg = '#37474F'
        self.configure(bg=self.pastel_bg)
        self.create_widgets()
        
    def create_widgets(self):
        container_frame = tk.Frame(self, bg=self.pastel_bg, padx=20, pady=20)
        container_frame.pack(expand=True, fill=tk.BOTH)
        
        title_label = tk.Label(container_frame, text="Tagesübersicht", font=("Segoe UI", 16, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)
        title_label.pack(pady=(0, 20))
        
        self.stats_frame = tk.Frame(container_frame, bg=self.pastel_bg)
        self.stats_frame.pack(fill=tk.X)
        
        self.short_list_frame = tk.Frame(container_frame, bg=self.pastel_bg)
        self.short_list_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        # Bulmor-Status beim Aufbau rendern
        try:
            self.refresh_bulmor()
        except Exception:
            pass
        
        self.update_statistics(None)

    def _parse_time(self, val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        if isinstance(val, (pd.Timestamp, datetime)):
            return datetime(2000, 1, 1, val.hour, val.minute, val.second)
        if isinstance(val, str):
            s = val.strip()
            m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', s)
            if m:
                hh = int(m.group(1)); mm = int(m.group(2)); ss = int(m.group(3) or 0)
                if 0 <= hh <= 23 and 0 <= mm <= 59 and 0 <= ss <= 59:
                    return datetime(2000, 1, 1, hh, mm, ss)
            try:
                ts = pd.to_datetime(s, errors='raise')
                return datetime(2000, 1, 1, ts.hour, ts.minute, ts.second)
            except Exception:
                return None
        return None

    def _shift_by_begin(self, begin_dt):
        if not begin_dt:
            return "unbekannt"
        return "tag" if 6 <= begin_dt.hour <= 17 else "nacht"

    def _duration_hours(self, begin_dt, end_dt):
        if not begin_dt or not end_dt:
            return None
        if end_dt < begin_dt:
            end_dt = end_dt + timedelta(days=1)
        return round((end_dt - begin_dt).total_seconds() / 3600.0, 2)

    def _role_code(self, s):
        s = (s or "").strip().upper()
        m = re.match(r'^[A-Z]+[0-9]*', s)
        return m.group(0) if m else s

    def update_statistics(self, df):
        for w in self.stats_frame.winfo_children(): w.destroy()
        for w in self.short_list_frame.winfo_children(): w.destroy()

        if df is None or df.empty:
            tk.Label(self.stats_frame, text="Bitte laden Sie eine Excel-Datei, um die Statistik anzuzeigen.",
                     font=("Segoe UI", 12), bg=self.pastel_bg, fg=self.pastel_fg).pack(pady=20)
            return

        work = df.copy()
        work['Dienst_clean'] = work['Dienst'].astype(str).str.strip().str.lower()
        work['Role'] = work['Dienst'].apply(self._role_code)

        # Deaktivierte Mitarbeitende ausschließen
        if 'deactivated' in work.columns:
            work = work[~work['deactivated'].astype(bool)]

        work = work[~work['Dienst_clean'].str.startswith(('dt', 'dn'))]
        name_norm = work['Namen'].astype(str).str.strip().str.lower()
        work = work[~name_norm.isin({'lars peters', 'peters, lars'})]

        is_sick_series = work['Dienst_clean'].str.startswith(('k', 'krank'))
        if 'is_sick' in work.columns:
            is_sick_series = is_sick_series | work['is_sick'].astype(bool)

        work['_begin_dt'] = work['Beginn'].apply(self._parse_time)
        work['_end_dt']   = work['Ende'].apply(self._parse_time)
        work['_shift']    = work['_begin_dt'].apply(self._shift_by_begin)
        work['_hours']    = work.apply(lambda r: self._duration_hours(r['_begin_dt'], r['_end_dt']), axis=1)

        kvs_heute = 0
        try:
            plan_date = None
            if 'Datum' in df.columns and len(df) > 0:
                first_date = str(df.iloc[0]['Datum'])
                try:
                    plan_date = datetime.strptime(first_date, "%d.%m.%Y")
                except Exception:
                    plan_date = pd.to_datetime(first_date, errors='coerce')
                    if pd.isna(plan_date): plan_date = None
                    elif isinstance(plan_date, pd.Timestamp): plan_date = plan_date.to_pydatetime()
            if plan_date is None: plan_date = datetime.today()

            app = self.master.master if hasattr(self, 'master') and hasattr(self.master, 'master') else None
            db = getattr(app, 'db', None) if app else None
            if db and db.conn:
                kvs_rows = db.fetch_kvs_for_date(plan_date)
                kvs_names = set()
                for name, _, d_von, d_bis in kvs_rows:
                    if not name: continue
                    ln = extract_lastname(str(name)).strip().lower()
                    if ln in ('lars peters', 'peters, lars'): continue
                    kvs_names.add(ln)
                kvs_heute = len(kvs_names)
        except Exception:
            kvs_heute = 0

        wanted_roles = ['T', 'T10', 'T8', 'N', 'N10', 'NF']
        counts_roles = {role: int((work['Role'] == role).sum()) for role in wanted_roles}
        # Nachnamen je Rolle für Anzeige hinter den Summen
        try:
            role_lastnames = {r: ', '.join(sorted({extract_lastname(n) for n in work.loc[work['Role']==r, 'Namen'].dropna()})) for r in wanted_roles}
        except Exception:
            role_lastnames = {r: '' for r in wanted_roles}

        t_roles = ['t', 't10', 't8']
        n_roles = ['n', 'n10', 'nf']
        count_tag = work[work['Dienst_clean'].str.startswith(tuple(t_roles))]['Dienst'].count()
        count_nacht = work[work['Dienst_clean'].str.startswith(tuple(n_roles))]['Dienst'].count()
        count_krank = int(is_sick_series.sum())
        sick_rows = work[is_sick_series]
        count_krank_tag = int((sick_rows['_shift'] == 'tag').sum())
        count_krank_nacht = int((sick_rows['_shift'] == 'nacht').sum())
        count_total = int(work.shape[0])

        row = 0
        tk.Label(self.stats_frame, text="Rollenübersicht:", font=("Segoe UI", 12, "bold"),
                 bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=(0,5))
        row += 1
        for r in ['T', 'T10', 'T8', 'N', 'N10', 'NF']:
            tk.Label(self.stats_frame, text=f"{r}:", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=3)
            tk.Label(self.stats_frame, text=str(counts_roles[r]), font=("Segoe UI", 12, "bold"),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=1, sticky='e', padx=10, pady=3)
            # Nachnamen rechts daneben
            try:
                names_text = role_lastnames.get(r, '')
            except Exception:
                names_text = ''
            tk.Label(self.stats_frame, text=names_text, font=("Segoe UI", 11),
                     bg=self.pastel_bg, fg=self.pastel_fg, anchor='w').grid(row=row, column=2, sticky='w', padx=10, pady=3)
            row += 1

        row += 1
        for label, value in [
            ("Tagdienste gesamt (T, T10, T8)", count_tag),
            ("Nachtdienste gesamt (N, N10, NF)", count_nacht),
            ("Krankmeldungen gesamt", count_krank),
            ("davon krank in Tagschicht", count_krank_tag),
            ("davon krank in Nachtschicht", count_krank_nacht),
            ("KvS heute", kvs_heute),
            ("Gesamtmitarbeiter im Dienst", count_total),
        ]:
            tk.Label(self.stats_frame, text=f"{label}:", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=3)
            tk.Label(self.stats_frame, text=str(value), font=("Segoe UI", 12, "bold"),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=1, sticky='e', padx=10, pady=3)
            row += 1

        

        # Liste der Kranken (Übersicht)
        try:
            sick_names = [str(n) for n in sick_rows['Namen'].dropna().tolist()]
            sick_text = ", ".join(sick_names) if sick_names else "—"
        except Exception:
            sick_text = "—"
        tk.Label(self.stats_frame, text="Krank (Namen):", font=("Segoe UI", 12, "bold"),
                 bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=(8,3))
        tk.Label(self.stats_frame, text=sick_text, font=("Segoe UI", 11),
                 bg=self.pastel_bg, fg=self.pastel_fg, anchor='w', justify="left", wraplength=900).grid(row=row, column=1, columnspan=2, sticky='w', padx=10, pady=(8,3))
        row += 1

        # Bulmor-Status in der Übersicht
        try:
            st = load_bulmor_status()
            lines = [f"Bulmor {i}: {st.get(f'Bulmor {i}', 'fahrbereit')}" for i in range(1, 6)]
            bulmor_text = "\n".join(lines)
        except Exception:
            bulmor_text = "—"
        tk.Label(self.stats_frame, text="Bulmor-Status:", font=("Segoe UI", 12, "bold"),
                 bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=(8,3))
        tk.Label(self.stats_frame, text=bulmor_text, font=("Segoe UI", 11),
                 bg=self.pastel_bg, fg=self.pastel_fg, anchor='w', justify="left", wraplength=900).grid(row=row, column=1, columnspan=2, sticky='w', padx=10, pady=(8,3))
        row += 1

        threshold = 10.0
        short_df = work[(work['_hours'].notna()) & (work['_hours'] < threshold) & (~is_sick_series)].copy()
        short_df = short_df[['Namen', 'Beginn', 'Ende', '_hours']].sort_values('_hours')

        tk.Label(self.short_list_frame, text=f"Mitarbeitende mit weniger als {threshold:g} Stunden:",
                 font=("Segoe UI", 12, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)\
          .grid(row=0, column=0, columnspan=2, sticky='w', padx=10, pady=(0,8))

        if short_df.empty:
            tk.Label(self.short_list_frame, text="— keine —", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=1, column=0, sticky='w', padx=20, pady=2)
        else:
            r = 1
            for _, rowv in short_df.iterrows():
                beg_str = format_time(rowv['Beginn'])
                end_str = format_time(rowv['Ende'])
                tk.Label(self.short_list_frame, text=f"{rowv['Namen']}:", font=("Segoe UI", 12),
                         bg=self.pastel_bg, fg=self.pastel_fg).grid(row=r, column=0, sticky='w', padx=20, pady=2)
                tk.Label(self.short_list_frame, text=f"{beg_str} – {end_str} ({rowv['_hours']:.2f} h)",
                         font=("Segoe UI", 12, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)\
                  .grid(row=r, column=1, sticky='w', padx=10, pady=2)
                r += 1

# -------------------------------------------------------------------
# Haupt-App
# -------------------------------------------------------------------



def refresh_bulmor(self):
    # (Re)render Bulmor-Status section
    try:
        if hasattr(self, "_bulmor_frame") and self._bulmor_frame.winfo_exists():
            self._bulmor_frame.destroy()
    except Exception:
        pass
    self._bulmor_frame = tk.LabelFrame(self, text="Bulmor-Status", bg=self.pastel_bg, fg=self.pastel_fg)
    self._bulmor_frame.pack(fill="x", padx=20, pady=(10,20))
    st = load_bulmor_status()
    # simple row showing Bulmor 1..5
    row = tk.Frame(self._bulmor_frame, bg=self.pastel_bg)
    row.pack(fill="x", padx=10, pady=6)
    for i in range(1,6):
        name = f"Bulmor {i}"
        val = st.get(name, "fahrbereit")
        tk.Label(row, text=f"{name}: {val}", bg=self.pastel_bg, fg=self.pastel_fg).pack(side="left", padx=10)


class BulmorTab(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(bg="#E0F7FA")
        self.vars = {}
        self._build_ui()

    def _build_ui(self):
        tk.Label(self, text="Bulmor Status", font=("Segoe UI", 14, "bold"),
                 bg="#E0F7FA", fg="#37474F").pack(anchor="w", padx=16, pady=(16,8))

        frm = tk.Frame(self, bg="#E0F7FA"); frm.pack(fill="x", padx=16, pady=(0,16))
        state = load_bulmor_status()

        # Widgets pro Bulmor sammeln
        self._widget_map = {}

        # ggf. letzte DB-Werte zum Vorbefüllen holen
        last = {}
        try:
            db = getattr(self.master.master, 'db', None)
            if db and getattr(db, 'conn', None):
                for rid, bul, st, wt, ft, erledigt in db.fetch_bulmor_entries():
                    last[bul] = {"Status": st, "WerkstattTermin": wt, "Freitext": ft}
        except Exception:
            pass

        for i in range(1,6):
            row = tk.Frame(frm, bg="#E0F7FA"); row.pack(fill="x", pady=6)
            name = f"Bulmor {i}"

            tk.Label(row, text=name + ":", width=12, anchor="w", bg="#E0F7FA", fg="#37474F").pack(side="left")
            v = tk.StringVar(value=state.get(name, "fahrbereit"))
            self.vars[name] = v
            cb = ttk.Combobox(row, textvariable=v, values=["defekt","fahrbereit","werkstatt"], width=12, state="readonly")
            cb.pack(side="left", padx=(0,8))
            cb.bind("<<ComboboxSelected>>", self._on_change)

            # Termin
            term_frame = tk.Frame(row, bg="#E0F7FA"); term_frame.pack(side="left")
            tk.Label(term_frame, text="Termin:", bg="#E0F7FA", fg="#37474F").pack(side="left")
            date_label = tk.Label(term_frame, text="—", width=12, anchor="w", bg="#E0F7FA", fg="#37474F")
            date_label.pack(side="left", padx=(4,4))
            ttk.Button(term_frame, text="…", width=3, command=lambda n=name: self._pick_wt_date(n)).pack(side="left", padx=(0,8))

            # Freitext
            tk.Label(row, text="Notiz:", bg="#E0F7FA", fg="#37474F").pack(side="left")
            text_entry = ttk.Entry(row, width=36)
            text_entry.pack(side="left", padx=(4,0))

            # Vorbefüllen
            preset = last.get(name, {})
            try:
                if isinstance(preset.get("Status"), str):
                    v.set(preset["Status"])
                if preset.get("WerkstattTermin") is not None:
                    date_label.config(text=str(preset["WerkstattTermin"])[:10])
                if isinstance(preset.get("Freitext"), str):
                    text_entry.insert(0, preset["Freitext"])
            except Exception:
                pass

            self._widget_map[name] = {
                "date_label": date_label,
                "text_entry": text_entry,
                "status_cb": cb,
            }

        tk.Label(self, text="Einstellung wird gespeichert und in der Tagesübersicht angezeigt.",
                 bg="#E0F7FA", fg="#555").pack(anchor="w", padx=16, pady=(6,16))
        ttk.Button(self, text="Speichern (Bulmor)", command=self.save_all_bulmor).pack(padx=16, pady=(0,10), anchor="w")

    def _pick_wt_date(self, bulmor_name):
        """Open date selector and set label"""
        try:
            from tkcalendar import DateEntry
            top = tk.Toplevel(self)
            top.title(f"Werkstatttermin für {bulmor_name}")
            cal = DateEntry(top, date_pattern="yyyy-mm-dd")
            cal.pack(padx=10, pady=10)
            def ok():
                sel = cal.get_date()
                lbl = self._widget_map[bulmor_name]["date_label"]
                try:
                    lbl.config(text=sel.strftime("%Y-%m-%d"))
                except Exception:
                    lbl.config(text=str(sel))
                top.destroy()
            ttk.Button(top, text="OK", command=ok).pack(pady=6)
            self.wait_window(top)
        except Exception:
            # fallback: askstring
            from tkinter import simpledialog
            s = simpledialog.askstring("Werkstatttermin", f"Werkstatttermin für {bulmor_name} (YYYY-MM-DD):", parent=self)
            if s:
                lbl = self._widget_map[bulmor_name]["date_label"]
                lbl.config(text=s)

    def _on_change(self, event=None):
        st = {k: v.get() for k, v in self.vars.items()}
        save_bulmor_status(st)


    def save_all_bulmor(self):
        """Save current status + werkstatt info to JSON and Access DB if available"""
        st = {k: v.get() for k, v in self.vars.items()}
        # save JSON as before
        save_bulmor_status(st)
        # also persist extra fields
        try:
            db = getattr(self.master.master, 'db', None)
            if db and getattr(db, 'conn', None):
                for i in range(1,6):
                    name = f"Bulmor {i}"
                    widgets = self._widget_map.get(name, {})
                    status = self.vars[name].get()
                    date_lbl = widgets.get("date_label")
                    date_txt = date_lbl.cget("text") if date_lbl else None
                    freetxt = widgets.get("text_entry").get() if widgets.get("text_entry") else None
                    # save into DB
                    db.save_bulmor_entry(name, status, werkstatt_dt=(date_txt or None), freitext=freetxt, erledigt=False)
        except Exception:
            pass

class UnifiedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NeSk")
        self.geometry("1400x900")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.pastel_bg = '#E0F7FA'
        self.pastel_fg = '#37474F'
        self.pastel_button = '#B2DFDB'
        self.configure(bg=self.pastel_bg)

        if not OUTLOOK_AVAILABLE:
            self.after(0, lambda: messagebox.showwarning(
                "Warnung",
                "Die 'pywin32'-Bibliothek wurde nicht gefunden.\n"
                "Die E-Mail-Funktion kann das Dokument nicht automatisch anhängen.\n"
                "Installiere sie mit: pip install pywin32"
            ))
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        style_vars = {'bg': self.pastel_bg, 'fg': self.pastel_fg, 'button': self.pastel_button}
        
        self.db = None

        self.statistics_tab = StatisticsTab(self.notebook, style_vars=style_vars)
        self.notebook.add(self.statistics_tab, text="Tagesübersicht")
        # --- Vordrucke Tab ---
        VORDRUCKE_PATH = r"C:\Vordrucke"
        self.vordrucke_tab = VordruckeTab(self.notebook, VORDRUCKE_PATH)
        self.notebook.add(self.vordrucke_tab, text="Vordrucke")

        # --- Automatisches Aktualisieren beim Tabwechsel ---
        def on_tab_changed(event):
            tab = event.widget.nametowidget(event.widget.select())
            if isinstance(tab, VordruckeTab):
                tab.load_files()
            try:
                # Wenn zur Tagesübersicht gewechselt wird, Bulmor-Status aktualisieren
                if isinstance(tab, StatisticsTab):
                    tab.refresh_bulmor()
            except Exception:
                pass

        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)
        
        self.krankmeldung_tab = KrankmeldungTab(self.notebook, db=self.db, style_vars=style_vars)
        self.notebook.add(self.krankmeldung_tab, text="Krankmeldung")
        
        self.dienstplan_tab = DienstplanTab(self.notebook, bg_color=self.pastel_bg, fg_color=self.pastel_fg)
        self.notebook.add(self.dienstplan_tab, text="Dienstplangenerator")

        # NEU: Sonderaufgaben-Tab
        self.sonderaufgaben_tab = SonderaufgabenTab(self.notebook)
        self.notebook.add(self.sonderaufgaben_tab, text="Sonderaufgaben")


        self.scanmail_tab = ScanMailTab(self.notebook, bg=self.pastel_bg)
        self.notebook.add(self.scanmail_tab, text="Scan‑Mail")

        # Bulmor Tab
        self.bulmor_tab = BulmorTab(self.notebook)
        self.notebook.add(self.bulmor_tab, text="Bulmor")
        self.excel_viewer_tab = ExcelViewerTab(self.notebook, db=self.db)
        self.notebook.add(self.excel_viewer_tab, text="Tagesdienstplan")
        
        self.connect_button = ttk.Button(self, text="Datenbank auswählen", command=self.select_and_connect_to_db, style='TButton')
        self.connect_button.pack(pady=10)
        
        # Automatische Verbindung beim Start mit festem Pfad
        self.connect_to_db(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\python\Database101.accdb")

    def handle_file_load(self, file_path, switch_tab=True, pop_success=True):
        success = self.excel_viewer_tab.load_data(file_path)
        if success and self.excel_viewer_tab.df is not None:
            self.statistics_tab.update_statistics(self.excel_viewer_tab.df)
            if switch_tab:
                self.notebook.select(self.excel_viewer_tab)  # old
        if success and self.excel_viewer_tab.df is not None:
            self.statistics_tab.update_statistics(self.excel_viewer_tab.df)
            if switch_tab:
                self.notebook.select(self.excel_viewer_tab)


# --- Explorer für Dienstpläne (unter GEN_ROOT_DIR) ---

    def connect_to_db(self, db_file):
        self.db = AccessDB(db_file)
        if self.db.conn:
            messagebox.showinfo("Verbindung erfolgreich", f"Verbindung zur Datenbank '{os.path.basename(db_file)}' erfolgreich hergestellt.")
            self.krankmeldung_tab.db = self.db
            self.excel_viewer_tab.db = self.db
            self.krankmeldung_tab.display_data_in_tables()
            if self.excel_viewer_tab.df is not None:
                self.excel_viewer_tab.recompute_sickness_from_db()
        else:
            self.db = None
            self.krankmeldung_tab.db = None
            self.excel_viewer_tab.db = None
            self.krankmeldung_tab.display_data_in_tables()

    def select_and_connect_to_db(self):
        db_file = filedialog.askopenfilename(title="Access-Datenbank auswählen", filetypes=[("Access-Datenbank", "*.accdb *.mdb")])
        if db_file:
            self.connect_to_db(db_file)


class ScanMailTab(tk.Frame):

    def _ask_for_date(self):
        """Fragt ein Datum über tkcalendar ab und gibt es als datetime.date zurück."""
        top = tk.Toplevel(self)
        top.title("Datum auswählen")

        cal = Calendar(top, selectmode="day", date_pattern="dd.mm.yyyy")
        cal.pack(padx=10, pady=10)

        selected_date = {"value": None}

        def on_ok():
            selected_date["value"] = cal.selection_get()
            top.destroy()

        ttk.Button(top, text="OK", command=on_ok).pack(pady=10)

        self.wait_window(top)
        return selected_date["value"]

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.bg_color = "#F5F7FA"
        self.fg_color = "#222"
        self.configure(bg=self.bg_color)
        self._build_ui()

    def _build_ui(self):
        title = tk.Label(self, text="Scan-Mail: PDF auswählen, Datum vergeben, als Entwurf öffnen", 
                         font=("Segoe UI", 12, "bold"), bg=self.bg_color, fg=self.fg_color)
        title.pack(padx=10, pady=(10, 6), anchor="w")

        # Date selector
        date_row = tk.Frame(self, bg=self.bg_color)
        date_row.pack(padx=10, pady=(0,8), anchor="w")
        tk.Label(date_row, text="Datum:", bg=self.bg_color, fg=self.fg_color).pack(side="left")
        try:
            from tkcalendar import DateEntry  # type: ignore
            self.date_picker = DateEntry(date_row, date_pattern="yyyy-mm-dd", width=12)
        except Exception:
            import datetime as _dt
            self.date_picker = tk.Entry(date_row, width=12)
            self.date_picker.insert(0, _dt.date.today().strftime("%Y-%m-%d"))
        # pack either widget
        try:
            self.date_picker.pack(side="left", padx=(6,0))
        except Exception:
            pass

        hint = tk.Label(self, text=f"Quelle: {SCAN_ROOT_DIR} (nur PDF). Ziel: entsprechender Monats-Unterordner unter {SCAN_ROOT_DIR}", 
                        bg=self.bg_color, fg="#555")
        hint.pack(padx=10, pady=(0,10), anchor="w")

        # Explorer (eingeschränkt auf SCAN_ROOT_DIR)
        self._init_scan_browser(self)

        # Action buttons
        btns = tk.Frame(self, bg=self.bg_color)
        btns.pack(padx=10, pady=(6,12), anchor="w")
        tk.Button(btns, text="Umbenennen + Speichern + Entwurf erstellen", command=self._sm_prepare_and_draft)\
            .pack(side="left")
        tk.Button(btns, text="Nur Entwurf mit letzter Datei", command=self._sm_draft_last)\
            .pack(side="left", padx=8)

    # --------- Explorer (Scan) ---------
    def _init_scan_browser(self, parent):
        import os
        wrap = tk.LabelFrame(parent, text="Explorer (KyoScan)", bg=self.bg_color, fg=self.fg_color)
        wrap.pack(padx=10, pady=(0,10), fill="both", expand=True)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_columnconfigure(1, weight=2)
        wrap.grid_rowconfigure(0, weight=1)

        left = tk.Frame(wrap, bg=self.bg_color)
        left.grid(row=0, column=0, sticky="nsew", padx=(6,3), pady=6)
        tv_dirs = ttk.Treeview(left, columns=("abspath",), show="tree")
        vs_l = ttk.Scrollbar(left, orient="vertical", command=tv_dirs.yview)
        tv_dirs.configure(yscrollcommand=vs_l.set)
        tv_dirs.pack(side="left", fill="both", expand=True)
        vs_l.pack(side="left", fill="y")
        self._sm_tv_dirs = tv_dirs

        right = tk.Frame(wrap, bg=self.bg_color)
        right.grid(row=0, column=1, sticky="nsew", padx=(3,6), pady=6)
        tv_files = ttk.Treeview(right, columns=("name","size","mtime","abspath"), show="headings", selectmode="extended")
        tv_files.heading("name", text="Datei"); tv_files.column("name", width=320, anchor="w")
        tv_files.heading("size", text="Größe"); tv_files.column("size", width=80, anchor="e")
        tv_files.heading("mtime", text="Geändert"); tv_files.column("mtime", width=150, anchor="w")
        tv_files["displaycolumns"] = ("name","size","mtime")
        vs_r = ttk.Scrollbar(right, orient="vertical", command=tv_files.yview)
        tv_files.configure(yscrollcommand=vs_r.set)
        tv_files.pack(side="left", fill="both", expand=True)
        vs_r.pack(side="left", fill="y")
        self._sm_tv_files = tv_files

        sel_frame = tk.Frame(wrap, bg=self.bg_color)
        sel_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,6))
        tk.Button(sel_frame, text="Zur Auswahl hinzufügen", command=self._sm_add_selected).pack(side="left")
        tk.Button(sel_frame, text="Aus Auswahl entfernen", command=self._sm_remove_selected).pack(side="left", padx=6)

        self._sm_selected = []
        self._sm_list = tk.Listbox(sel_frame, height=3, selectmode="extended")
        self._sm_list.pack(side="left", fill="x", expand=True, padx=(10,0))

        root = SCAN_ROOT_DIR
        self._sm_root_dir = root
        root_text = os.path.basename(root.rstrip("\\/")) or root
        root_node = tv_dirs.insert("", "end", text=root_text, values=(root,), open=True)
        tv_dirs.insert(root_node, "end", text="...", values=(os.path.join(root, "__dummy__"),))

        def on_open(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            self._populate_scan_dir(node)

        def on_select_dir(evt):
            sel = tv_dirs.selection()
            if not sel: return
            node = sel[0]
            abspath = tv_dirs.set(node, "abspath")
            self._list_scan_files(abspath)

        tv_dirs.bind("<<TreeviewOpen>>", on_open)
        tv_dirs.bind("<<TreeviewSelect>>", on_select_dir)

        def on_file_double(evt):
            iid = tv_files.focus()
            if not iid: return
            abspath = tv_files.set(iid, "abspath")
            if abspath and abspath.lower().endswith(".pdf"):
                self._sm_selected = [abspath]
                self._sm_list.delete(0, "end")
                self._sm_list.insert("end", abspath)
        tv_files.bind("<Double-1>", on_file_double)

        self._populate_scan_dir(root_node)
        self._list_scan_files(root)

    def _is_within_scan_root(self, path):
        import os
        root = os.path.abspath(self._sm_root_dir)
        try:
            ap = os.path.abspath(path)
        except Exception:
            return False
        return ap.startswith(root)

    def _populate_scan_dir(self, node):
        import os
        tv = self._sm_tv_dirs
        path = tv.set(node, "abspath")
        if not self._is_within_scan_root(path):
            return
        for k in tv.get_children(node):
            if tv.item(k, "text") == "...":
                tv.delete(k)
        try:
            entries = sorted([d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))])
        except Exception:
            entries = []
        for d in entries:
            abspath = os.path.join(path, d)
            if not self._is_within_scan_root(abspath):
                continue
            exists = any(tv.set(c, "abspath") == abspath for c in tv.get_children(node))
            if exists:
                continue
            new = tv.insert(node, "end", text=d, values=(abspath,))
            try:
                if any(os.path.isdir(os.path.join(abspath, x)) for x in os.listdir(abspath)):
                    tv.insert(new, "end", text="...", values=(os.path.join(abspath, "__dummy__"),))
            except Exception:
                pass

    def _list_scan_files(self, dir_path):
        import os, datetime
        tvf = self._sm_tv_files
        for c in tvf.get_children(): tvf.delete(c)
        if not self._is_within_scan_root(dir_path):
            return
        try:
            items = os.listdir(dir_path)
        except Exception:
            items = []
        files = []
        for nm in items:
            abspath = os.path.join(dir_path, nm)
            if os.path.isfile(abspath) and nm.lower().endswith(".pdf"):
                try:
                    st = os.stat(abspath)
                    size = st.st_size
                    mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M")
                except Exception:
                    size = 0; mtime = ""
                files.append((nm, size, mtime, abspath))
        for nm, size, mtime, abspath in sorted(files):
            tvf.insert("", "end", values=(nm, f"{size//1024} KB", mtime, abspath))

    def _sm_add_selected(self):
        tvf = self._sm_tv_files
        sel = tvf.selection()
        for iid in sel:
            abspath = tvf.set(iid, "abspath")
            if abspath and abspath.lower().endswith(".pdf") and abspath not in self._sm_selected:
                self._sm_selected.append(abspath)
                self._sm_list.insert("end", abspath)

    def _sm_remove_selected(self):
        idxs = list(self._sm_list.curselection())[::-1]
        for i in idxs:
            p = self._sm_list.get(i)
            self._sm_list.delete(i)
            try:
                self._sm_selected.remove(p)
            except Exception:
                pass

    # --------- Core actions ---------
    def _sm_prepare_and_draft(self):
        import os, shutil, datetime
        # Get selected file
        src = None
        if self._sm_selected:
            src = self._sm_selected[0]
        elif self._sm_tv_files.focus():
            src = self._sm_tv_files.set(self._sm_tv_files.focus(), "abspath")
        if not src or not src.lower().endswith(".pdf"):
            try:
                messagebox.showinfo("Auswahl", "Bitte eine PDF-Datei im rechten Bereich auswählen.")
            except Exception:
                print("Bitte eine PDF-Datei auswählen.")
            return

        # Get date
        try:
            val = self.date_picker.get_date()
        except Exception:
            try:
                s = self.date_picker.get()
            except Exception:
                s = ""
            try:
                val = datetime.datetime.strptime(s, "%Y-%m-%d").date()
            except Exception:
                val = datetime.date.today()
        d = val

        # Determine destination subfolder (month)
        dest_dir = self._sm_pick_month_subfolder(SCAN_ROOT_DIR, d)
        os.makedirs(dest_dir, exist_ok=True)

        # New filename
        newname = f"{d:%Y_%m_%d}.pdf"
        dest_path = os.path.join(dest_dir, newname)
        # Ensure uniqueness
        k = 2
        while os.path.exists(dest_path):
            dest_path = os.path.join(dest_dir, f"{d:%Y_%m_%d}_{k}.pdf")
            k += 1

        # Copy file
        try:
            shutil.copy2(src, dest_path)
            self._sm_last_saved = dest_path
        except Exception as e:
            try:
                messagebox.showerror("Speicherfehler", f"{e}")
            except Exception:
                print("Speicherfehler:", e)
            return

        # Draft email in Outlook
        self._sm_create_outlook_draft(dest_path, d)

    def _sm_draft_last(self):
        import datetime
        d = datetime.date.today()
        try:
            self._sm_create_outlook_draft(self._sm_last_saved, d)
        except Exception:
            try:
                messagebox.showinfo("Hinweis", "Es wurde noch keine Datei gespeichert.")
            except Exception:
                pass

    def _sm_pick_month_subfolder(self, root, date_obj):
        import os, calendar
        y, m = date_obj.year, date_obj.month
        candidates = []
        month_num = f"{y:04d}_{m:02d}"
        month_num2 = f"{y:04d}-{m:02d}"
        month_name = calendar.month_name[m]
        for d in os.listdir(root):
            p = os.path.join(root, d)
            if not os.path.isdir(p): 
                continue
            dn = d.lower()
            if month_num in d or month_num2 in d or month_name.lower() in dn:
                candidates.append(p)
        if candidates:
            # pick the one with longest match (more specific)
            return sorted(candidates, key=lambda s: len(s))[-1]
        # else default to YYYY_MM
        return os.path.join(root, month_num)

    def _sm_create_outlook_draft(self, attach_path, d):
        if not attach_path:
            return
        try:
            import win32com.client as win32  # requires pywin32
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.To = "leitung.fb2@drk-koeln.de"
            mail.Subject = f"Scan {d:%Y-%m-%d}"
            mail.Body = "Guten Tag,\n\nanbei der Scan.\n\nBeste Grüße"
            mail.Attachments.Add(Source=attach_path)
            mail.Display()  # show as draft
        except Exception as e:
            # Fallback Hinweis
            try:
                messagebox.showinfo("Entwurf", f"Outlook-Entwurf konnte nicht automatisch erstellt werden.\nDatei gespeichert unter:\n{attach_path}\n\nFehler: {e}")
            except Exception:
                print("Entwurf-Hinweis:", e, attach_path)



if __name__ == "__main__":
    try:
        app = UnifiedApp()
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        from tkinter import messagebox
        messagebox.showerror("Fehler beim Start", f"{e}")
    
def _export_sonderaufgaben_like_uploaded(assignments_rows, save_path, template_path=None, sheet_name="Sonderaufgaben"):
    """
    assignments_rows: Liste[Dict] mit Keys: 'Aufgabe', 'Tag', 'Nacht' (optional 'Bemerkung')
    save_path: Zielpfad .xlsx
    template_path: Pfad zu "Sonderaufgaben.xlsx" (wenn None -> im Skriptverzeichnis gesucht, sonst Fallback-Layout)
    """
    import os, re
    from datetime import datetime

    # 1) Vorlage bestimmen
    if template_path is None:
        # Erst im Skriptverzeichnis versuchen
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            script_dir = os.getcwd()
        cand = os.path.join(script_dir, "Sonderaufgaben.xlsx")
        if os.path.isfile(cand):
            template_path = cand

    # 2) Vorlage laden oder Fallback bauen
    if template_path and os.path.isfile(template_path):
        wb = load_workbook(template_path)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        # Minimaler Fallback (ähnliche Optik)
        ws["A1"] = "DRK KV Köln e.V."
        ws["C1"] = "Sonderaufgaben"
        ws["E1"] = "Sanitätsstation CGN"
        ws["A2"] = datetime.today()
        ws["C2"] = "Tagdienst"
        ws["E2"] = "Nachtdienst"
        ws["C1"].font = Font(name="Arial", size=26, bold=True)
        ws["C2"].font = Font(name="Arial", size=24, bold=True)
        ws["E2"].font = Font(name="Arial", size=24, bold=True)
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["G"].width = 25

    # 3) Aufgabenzeilen aus Vorlage erfassen + Bemerkung finden
    template_rows = {}
    remark_row = None
    for r in range(3, 500):
        val = ws.cell(row=r, column=1).value
        if val is None:
            continue
        txt = str(val).strip()
        if re.sub(r"\s+", "", txt).lower() == "bemerkung":
            remark_row = r
            continue
        if txt and txt.lower() not in ["tagdienst", "nachtdienst"]:
            template_rows[txt] = r

    def _clone_style(src_cell, dst_cell):
        try:
            if src_cell.has_style:
                dst_cell.font = src_cell.font.copy()
                dst_cell.fill = src_cell.fill.copy()
                dst_cell.alignment = src_cell.alignment.copy()
                dst_cell.border = src_cell.border.copy()
                dst_cell.number_format = src_cell.number_format
        except Exception:
            pass

    ref_row = min(template_rows.values()) if template_rows else None

    def _write_task_row(row_idx, aufgabe, tag_txt, nacht_txt):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2)
        ws.merge_cells(start_row=row_idx, start_column=3, end_row=row_idx, end_column=4)
        ws.merge_cells(start_row=row_idx, start_column=5, end_row=row_idx, end_column=6)

        cA = ws.cell(row=row_idx, column=1); cA.value = aufgabe
        cC = ws.cell(row=row_idx, column=3); cC.value = (tag_txt or "")
        cE = ws.cell(row=row_idx, column=5); cE.value = (nacht_txt or "")

        if ref_row:
            for src_col, dst_col in [(1,1),(3,3),(5,5)]:
                _clone_style(ws.cell(ref_row, src_col), ws.cell(row_idx, dst_col))
        else:
            base_font = Font(name="Arial", size=12)
            left = Alignment(horizontal="left", vertical="center")
            mid  = Alignment(horizontal="center", vertical="center")
            border = Border(left=Side(style="thin"), right=Side(style="thin"),
                            top=Side(style="thin"), bottom=Side(style="thin"))
            for col in [1,2,3,4,5,6]:
                c = ws.cell(row=row_idx, column=col)
                c.font = base_font; c.border = border
                c.alignment = mid if col in (3,4,5,6) else left

    last_task_row = max(template_rows.values()) if template_rows else 2
    attach_before = remark_row if remark_row else (last_task_row + 1)
    norm_tpl = {re.sub(r"\s+", " ", k).strip().lower(): v for k, v in template_rows.items()}

    # Zeilen schreiben
    for row in assignments_rows:
        aufgabe = str(row.get("Aufgabe", "")).strip()
        if not aufgabe:
            continue
        tag_txt = str(row.get("Tag", "")).strip() if row.get("Tag", None) is not None else ""
        nacht_txt = str(row.get("Nacht", "")).strip() if row.get("Nacht", None) is not None else ""

        key = re.sub(r"\s+", " ", aufgabe).lower()
        if key in norm_tpl:
            r = norm_tpl[key]
            ws.cell(row=r, column=3).value = tag_txt
            ws.cell(row=r, column=5).value = nacht_txt
        else:
            insert_row = attach_before
            ws.insert_rows(insert_row, amount=1)
            _write_task_row(insert_row, aufgabe, tag_txt, nacht_txt)
            attach_before += 1

    # Bemerkung (erste nicht-leere)
    if remark_row:
        for row in assignments_rows:
            b = row.get("Bemerkung", None)
            if b and str(b).strip():
                # bei Merges: Top-Left Zelle in Bereich C:D ist C
                ws.cell(row=remark_row, column=3).value = str(b).strip()
                break

    os.makedirs(os.path.dirname(save_path), exist_ok=True) if os.path.dirname(save_path) else None
    wb.save(save_path)
# ===== Ende Template-Style Export =====

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# High-DPI Fix (Windows)
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except (ImportError, AttributeError):
    pass

# Outlook/pywin32
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False  # Warnung zeigen wir nach Tk-Start

# -------------------------------------------------------------------
# Konfiguration
# -------------------------------------------------------------------
CONFIG = {
    "image_path": r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\python\drk.jpg",
    "dispo_roles": ["DT", "DT3", "DN"],  # Dispo-Rollen (Word: Zeiten abrunden)
    "excluded_names": ["Peters"],  # für Word (Nachname)
    "excluded_services": ['k', 'krank', 'lg', 'lehrgang', 'ea', 'einarbeitung'],
    "email_recipients": {
        "to": "hildegard.eichler@koeln-bonn-airport.de; erste-hilfe-station-flughafen@drk-koeln.de",
        "cc": "leitung.fb2@drk-koeln.de; verwaltung.fb2@drk-koeln.de; flughafen2@drk-koeln.de; loahrs@gmx.de"
    }
}

# -------------------------------------------------------------------
# Hilfsfunktionen
# -------------------------------------------------------------------
def format_time(t):
    import numpy as _np
    from datetime import time as _dtime

    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""

    if isinstance(t, (pd.Timestamp, datetime)):
        return t.strftime("%H:%M")

    if isinstance(t, _dtime):
        return f"{t.hour:02d}:{t.minute:02d}"

    if isinstance(t, (_np.datetime64,)):
        try:
            ts = pd.to_datetime(t)
            return ts.strftime("%H:%M")
        except Exception:
            pass

    if isinstance(t, str):
        s = t.strip()
        m = re.match(r'^(\d{1,2}):(\d{2})(?::\d{2})?$', s)
        if m:
            hh = int(m.group(1)); mm = int(m.group(2))
            return f"{hh:02d}:{mm:02d}"
        try:
            ts = pd.to_datetime(s, errors='raise')
            return ts.strftime("%H:%M")
        except Exception:
            return s

    return str(t)

def extract_lastname(name):
    if isinstance(name, str) and "," in name:
        return name.split(",")[0].strip()
    return name.strip() if isinstance(name, str) else name

def group_by_time(entries):
    grouped = {}
    for time, name in entries:
        grouped.setdefault(time, []).append(name)
    return grouped

def add_heading(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

def add_time_block(doc, time, names):
    p = doc.add_paragraph()
    run = p.add_run(time)
    run.bold = True
    run.font.size = Pt(12)
    p.add_run("\t")
    for i, name in enumerate(names):
        if i > 0 and i % 6 == 0:
            p.add_run("\n" + "\t" * 3)
        elif i > 0:
            p.add_run(" / ")
        p.add_run(name)

def generate_time_choices(step_minutes=15):
    times = [""]
    for h in range(24):
        for m in range(0, 60, step_minutes):
            times.append(f"{h:02d}:{m:02d}")
    return times

def _sort_key_minutes(tstr):
    m = re.search(r'(\d{1,2}):(\d{2})', str(tstr))
    if m:
        hh = int(m.group(1)); mm = int(m.group(2))
        return hh * 60 + mm
    return 10**9

def floor_to_full_hour(timestr):
    m = re.match(r'^\s*(\d{1,2}):(\d{2})', str(timestr))
    if not m:
        return format_time(timestr)
    hh = int(m.group(1))
    return f"{hh:02d}:00"

# -------------------------------------------------------------------
# Word-Generator aus DataFrame (Tagesdienstplan)
# -------------------------------------------------------------------
def build_word_from_df(df, plan_date, save_path, paxzahl=None):
    doc = Document()

    # Header
    section = doc.sections[0]
    header = section.header
    htable = header.add_table(1, 2, width=Inches(6))
    htable.autofit = False
    hcell1 = htable.cell(0, 0)
    hcell1.width = Inches(1.5)
    try:
        if os.path.isfile(CONFIG["image_path"]):
            hcell1.paragraphs[0].add_run().add_picture(CONFIG["image_path"], width=Inches(1.25))
    except Exception:
        pass
    hcell2 = htable.cell(0, 1)
    hcell2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    hcell2.paragraphs[0].text = "Deutsches Rotes Kreuz Kreisverband Köln e.V."
    hcell2.paragraphs[0].add_run("\n- Unfallhilfsstelle und Betreuungsstelle für Behinderte am Flughafen Köln/Bonn –")
    date_para = doc.add_paragraph()
    run_label = date_para.add_run("Datum:")
    run_label.bold = True
    run_label.font.size = Pt(14)
    run_date = date_para.add_run(f"\t\t{header_start.strftime('%d.%m.%Y')} – {header_end.strftime('%d.%m.%Y')}")
    run_date.font.size = Pt(12)
    doc.add_paragraph()

    # Einträge gruppieren – aus df
    dispo_entries = []
    betreuer_entries = []

    for _, r in df.iterrows():
        if 'deactivated' in df.columns and bool(r.get('deactivated', False)):
            continue
        name = extract_lastname(r["Namen"])
        if name in CONFIG["excluded_names"]:
            continue

        dienst_raw = str(r["Dienst"]).strip() if pd.notnull(r["Dienst"]) else ""
        dienst_lower = dienst_raw.lower()

        if any(ex in dienst_lower for ex in CONFIG["excluded_services"]):
            continue

        beginn = format_time(r["Beginn"])
        ende   = format_time(r["Ende"])

        if any(dienst_raw.upper().startswith(role) for role in CONFIG["dispo_roles"]):
            b_out = floor_to_full_hour(beginn)
            e_out = floor_to_full_hour(ende)
            time_window = f"{b_out} bis {e_out}"
            dispo_entries.append((time_window, name))
        else:
            time_window = f"{beginn} bis {ende}"
            betreuer_entries.append((time_window, name))

    dispo_grouped = group_by_time(dispo_entries)
    betreuer_grouped = group_by_time(betreuer_entries)

    add_heading(doc, "Disposition")
    for time in sorted(dispo_grouped.keys(), key=_sort_key_minutes):
        add_time_block(doc, time, dispo_grouped[time])
    doc.add_paragraph()

    add_heading(doc, "Behindertenbetreuer")
    for time in sorted(betreuer_grouped.keys(), key=_sort_key_minutes):
        add_time_block(doc, time, betreuer_grouped[time])

    if paxzahl is not None:
        doc.add_paragraph()
        pax_para = doc.add_paragraph(f"-- {paxzahl} --")
        pax_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    footer = section.footer
    table = footer.add_table(rows=1, cols=3, width=Inches(6))
    table.autofit = True
    cells = table.rows[0].cells
    cells[0].text = "Telefon: +49 220340 – 2323"
    cells[1].text = "email: flughafen@drk-koeln.de"
    cells[2].text = "Stationsleitung: Lars Peters"
    for cell in cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    doc.save(save_path)

# -------------------------------------------------------------------
# DienstplanTab – Excel einlesen & an Viewer weiterreichen
# -------------------------------------------------------------------

class VordruckeTab(ttk.Frame):
    def __init__(self, master, path, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.path = path

        ttk.Label(self, text="Vordrucke", font=("Segoe UI", 12, "bold")).pack(pady=5)

        self.tree = ttk.Treeview(self, columns=("Datei",), show="headings", height=15)
        self.tree.heading("Datei", text="Dateiname")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="Öffnen", command=self.open_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Drucken", command=self.print_file).pack(side=tk.LEFT, padx=5)

        self.load_files()

    def load_files(self):
        import os
        self.tree.delete(*self.tree.get_children())
        if os.path.exists(self.path):
            for file in os.listdir(self.path):
                full_path = os.path.join(self.path, file)
                if os.path.isfile(full_path):
                    self.tree.insert("", "end", values=(file,), iid=full_path)

    def open_file(self):
        import os, tkinter as tk
        sel = self.tree.selection()
        if sel:
            try:
                os.startfile(sel[0])
            except Exception as e:
                tk.messagebox.showerror("Fehler", f"Datei konnte nicht geöffnet werden:\n{e}")

    def print_file(self):
        import os, tkinter as tk
        sel = self.tree.selection()
        if sel:
            try:
                os.startfile(sel[0], "print")
            except Exception as e:
                tk.messagebox.showerror("Fehler", f"Datei konnte nicht gedruckt werden:\n{e}")

class DienstplanTab(tk.Frame):
    def __init__(self, master=None, bg_color=None, fg_color=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.configure(bg=self.bg_color)
        self.create_widgets()
        
    def create_widgets(self):
        # safety inits for new features
        self.kvs_time_map = getattr(self, 'kvs_time_map', {})
        self.deactivated_set = getattr(self, 'deactivated_set', set())
        frame = tk.Frame(self, padx=10, pady=10, bg=self.bg_color)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Excel-Quelldatei:", bg=self.bg_color, fg=self.fg_color).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_entry = tk.Entry(frame, width=60)
        self.file_entry.grid(row=0, column=1, pady=5)
        tk.Button(frame, text="Durchsuchen", command=self.select_file).grid(row=0, column=2, padx=5)

        tk.Button(frame, text="Neu laden", command=self.reload_selected_file).grid(row=0, column=3, padx=5)
        # Auto-Reload (silent) + Countdown
        try:
            self.reload_interval = 60  # Sekunden
            self._reload_seconds_left = self.reload_interval
            self._reload_label = tk.Label(frame, text=f"Auto-Reload in: {self._reload_seconds_left}s",
                                          bg=self.bg_color, fg=self.fg_color)
            self._reload_label.grid(row=0, column=4, padx=10, sticky=tk.W)
            self._reload_timer_id = None
            self._schedule_auto_reload_tick()
        except Exception:
            pass


        hint = tk.Label(frame, text="Die geladene Datei wird im Tab „Tagesdienstplan“ angezeigt.",
                        bg=self.bg_color, fg=self.fg_color)
        hint.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10,0))
        self._init_fs_browser(frame)
        self._init_fs_browser(frame)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.master.master.handle_file_load(file_path, switch_tab=False, pop_success=False)


    def reload_selected_file(self):
        """Lädt die im Feld ausgewählte Datei erneut."""
        try:
            path = (self.file_entry.get() or '').strip()
        except Exception:
            path = ''
        if not path:
            try:
                messagebox.showwarning('Neu laden', 'Bitte zuerst eine Datei auswählen (Durchsuchen).')
            except Exception:
                pass
            return
        if not os.path.exists(path):
            try:
                messagebox.showerror('Neu laden', f'Datei nicht gefunden:\n{path}')
            except Exception:
                pass
            return
        ok = self.master.master.handle_file_load(path, switch_tab=False, pop_success=False)
        if ok:
            try:
                self._reload_seconds_left = self.reload_interval
            except Exception:
                pass
# -------------------------------------------------------------------
# Access-Datenbank
# -------------------------------------------------------------------
    def _schedule_auto_reload_tick(self):
        """Plan den nächsten Tick in 1s."""
        try:
            if getattr(self, '_reload_timer_id', None) is not None:
                self.after_cancel(self._reload_timer_id)
        except Exception:
            pass
        try:
            self._reload_timer_id = self.after(1000, self._auto_reload_tick)
        except Exception:
            pass

    def _auto_reload_tick(self):
        """Zählt runter und lädt die ausgewählte Datei still neu, wenn 0 erreicht."""
        try:
            sec = int(getattr(self, '_reload_seconds_left', 60)) - 1
        except Exception:
            sec = 59
        if sec <= 0:
            sec = int(getattr(self, 'reload_interval', 60))
            try:
                path = (self.file_entry.get() or '').strip()
            except Exception:
                path = ''
            if path and os.path.exists(path):
                try:
                    # silent reload
                    self.master.master.handle_file_load(path, switch_tab=False, pop_success=False)
                except Exception:
                    pass
        # Update Anzeige
        try:
            self._reload_seconds_left = sec
            if hasattr(self, '_reload_label'):
                self._reload_label.config(text=f"Auto-Reload in: {sec}s")
        except Exception:
            pass
        self._schedule_auto_reload_tick()

class AccessDB:
    def __init__(self, db_file):
        self.conn = None
        self.cursor = None
        self.db_file = db_file
        self.connect()

    def _ensure_schema(self):
        if not self.conn:
            return
        try:
            try:
                self.cursor.execute("SELECT KvS FROM Tabelle1 WHERE 1=0")
            except pyodbc.Error:
                self.cursor.execute("ALTER TABLE Tabelle1 ADD COLUMN KvS YESNO")
                self.conn.commit()
            try:
                self.cursor.execute("SELECT Gegangen_um FROM Tabelle1 WHERE 1=0")
            except pyodbc.Error:
                self.cursor.execute("ALTER TABLE Tabelle1 ADD COLUMN Gegangen_um DATETIME")
                self.conn.commit()
        except pyodbc.Error:
            pass

    # ---- Bulmor Werkstatt Table helpers ----
    def _ensure_bulmor_table(self):
        if not self.conn:
            return
        try:
            # Try to create table if not exists (Access SQL limited): attempt select
            try:
                self.cursor.execute("SELECT ID FROM BulmorWerkstatt WHERE 1=0")
            except pyodbc.Error:
                try:
                    create_sql = (
                        "CREATE TABLE BulmorWerkstatt ("
                        "ID AUTOINCREMENT PRIMARY KEY, "
                        "Bulmor TEXT(50), "
                        "Status TEXT(50), "
                        "WerkstattTermin DATETIME, "
                        "Freitext MEMO, "
                        "Erledigt YESNO"
                        ")"
                    )
                    self.cursor.execute(create_sql)
                    self.conn.commit()
                except Exception:
                    # If creation fails (permissions) ignore
                    pass
        except Exception:
            pass

    def save_bulmor_entry(self, bulmor_name, status, werkstatt_dt=None, freitext=None, erledigt=False):
        """Insert or update an entry for given bulmor_name and date; keep latest per Bulmor."""
        if not self.conn:
            return False
        try:
            self._ensure_bulmor_table()
            # Try to find existing row for same Bulmor
            self.cursor.execute("SELECT ID FROM BulmorWerkstatt WHERE Bulmor = ?", (bulmor_name,))
            row = self.cursor.fetchone()
            if row:
                # update
                self.cursor.execute(
                    "UPDATE BulmorWerkstatt SET Status = ?, WerkstattTermin = ?, Freitext = ?, Erledigt = ? WHERE ID = ?",
                    (status, werkstatt_dt, freitext, -1 if erledigt else 0, row[0])
                )
            else:
                self.cursor.execute(
                    "INSERT INTO BulmorWerkstatt (Bulmor, Status, WerkstattTermin, Freitext, Erledigt) VALUES (?, ?, ?, ?, ?)",
                    (bulmor_name, status, werkstatt_dt, freitext, -1 if erledigt else 0)
                )
            self.conn.commit()
            return True
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False

    def fetch_bulmor_entries(self):
        if not self.conn:
            return []
        try:
            self._ensure_bulmor_table()
            self.cursor.execute("SELECT ID, Bulmor, Status, WerkstattTermin, Freitext, Erledigt FROM BulmorWerkstatt")
            rows = self.cursor.fetchall()
            return rows
        except Exception:
            return []

    def mark_bulmor_done(self, entry_id, done=True):
        if not self.conn:
            return False
        try:
            self.cursor.execute("UPDATE BulmorWerkstatt SET Erledigt = ? WHERE ID = ?", (-1 if done else 0, entry_id))
            self.conn.commit()
            return True
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False

    def connect(self):
        try:
            if not os.path.exists(self.db_file):
                messagebox.showerror("Datenbankfehler", f"Datenbankfile nicht gefunden: {self.db_file}")
                return False
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' f'DBQ={self.db_file};')
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
            self._ensure_schema()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Verbindungsfehler zur Access-Datenbank: {e}\n\nStellen Sie sicher, dass der passende 'Microsoft Access Driver' installiert ist.")
            self.conn = None
            return False

    def close(self):
        if self.conn:
            self.conn.close()

    def get_or_create_mitarbeiter_id(self, name):
        if not name or not self.conn:
            return None
        try:
            self.cursor.execute("SELECT Mitarbeiter_ID FROM Tabelle2 WHERE Name = ?", (name,))
            result = self.cursor.fetchone()
            if result:
                return result[0]
            else:
                self.cursor.execute("INSERT INTO Tabelle2 (Name) VALUES (?)", (name,))
                self.conn.commit()
                return self.cursor.execute("SELECT @@IDENTITY").fetchone()[0]
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Zugriff auf Tabelle2: {e}")
            return None

    def save_krankmeldung(self, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=False, t_gegangen=None):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            meldender_id = self.get_or_create_mitarbeiter_id(s_meldender)
            angenommen_von_id = self.get_or_create_mitarbeiter_id(s_angenommen)
            if meldender_id is None or angenommen_von_id is None:
                return False
            insert_query = """
                INSERT INTO Tabelle1
                    ([Datum], [Meldender_ID], [Krank_von], [Krank_bis], [Anruf_um], [Angenommen_von_ID], [Bemerkung], [KvS], [Gegangen_um])
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            self.cursor.execute(insert_query, (
                d_datum,
                meldender_id,
                d_von,
                d_bis,
                t_anruf.time(),
                angenommen_von_id,
                s_bem,
                -1 if kvs else 0,
                t_gegangen.time() if t_gegangen else None
            ))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Speichern der Krankmeldung: {e}")
            self.conn.rollback()
            return False
            
    def update_krankmeldung(self, entry_id, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=False, t_gegangen=None):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            meldender_id = self.get_or_create_mitarbeiter_id(s_meldender)
            angenommen_von_id = self.get_or_create_mitarbeiter_id(s_angenommen)
            if meldender_id is None or angenommen_von_id is None:
                return False
            update_query = """
                UPDATE Tabelle1
                SET [Datum] = ?, [Meldender_ID] = ?, [Krank_von] = ?, [Krank_bis] = ?, [Anruf_um] = ?,
                    [Angenommen_von_ID] = ?, [Bemerkung] = ?, [KvS] = ?, [Gegangen_um] = ?
                WHERE Eintrag_ID = ?
            """
            self.cursor.execute(update_query, (
                d_datum,
                meldender_id,
                d_von,
                d_bis,
                t_anruf.time(),
                angenommen_von_id,
                s_bem,
                -1 if kvs else 0,
                t_gegangen.time() if t_gegangen else None,
                entry_id
            ))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Aktualisieren der Krankmeldung: {e}")
            self.conn.rollback()
            return False

    def delete_krankmeldung(self, entry_id):
        if not self.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        try:
            delete_query = "DELETE FROM Tabelle1 WHERE Eintrag_ID = ?"
            self.cursor.execute(delete_query, (entry_id,))
            self.conn.commit()
            return True
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Löschen der Krankmeldung: {e}")
            self.conn.rollback()
            return False

    def fetch_all_krankmeldungen_with_id(self, start_date=None, end_date=None):
        if not self.conn:
            return []
        try:
            query = """
                SELECT
                    T1.Eintrag_ID,
                    T1.Datum,
                    T2_Meldender.Name AS Meldender,
                    T1.Krank_von,
                    T1.Krank_bis,
                    T1.Anruf_um,
                    T2_Angenommen.Name AS Angenommen_von,
                    T1.Bemerkung,
                    T1.KvS,
                    T1.Gegangen_um
                FROM
                    (Tabelle1 AS T1
                    INNER JOIN Tabelle2 AS T2_Meldender ON T1.Meldender_ID = T2_Meldender.Mitarbeiter_ID)
                    INNER JOIN Tabelle2 AS T2_Angenommen ON T1.Angenommen_von_ID = T2_Angenommen.Mitarbeiter_ID
            """
            params = []
            if start_date and end_date:
                query += " WHERE T1.Datum BETWEEN ? AND ?"
                end_of_day = datetime(end_date.year, end_date.month, end_date.day, 23, 59, 59)
                params = [start_date, end_of_day]
            query += " ORDER BY T1.Datum DESC;"
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der Daten: {e}")
            return []

    def fetch_sick_employees_in_range(self, start_date, end_date):
        if not self.conn:
            return []
        try:
            query = """
                SELECT DISTINCT T2.Name, T1.Krank_von, T1.Krank_bis
                FROM Tabelle1 AS T1
                INNER JOIN Tabelle2 AS T2 ON T1.Meldender_ID = T2.Mitarbeiter_ID
                WHERE T1.Krank_bis >= ? AND T1.Krank_von <= ?;
            """
            self.cursor.execute(query, (start_date, end_date))
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der aktuell Kranken: {e}")
            return []

    def fetch_kvs_for_date(self, date_obj):
        if not self.conn:
            return []
        try:
            plan_start = datetime(date_obj.year, date_obj.month, date_obj.day, 0, 0, 0)
            plan_end   = datetime(date_obj.year, date_obj.month, date_obj.day, 23, 59, 59)
            query = """
                SELECT T2.Name, T1.Gegangen_um, T1.Krank_von, T1.Krank_bis
                FROM Tabelle1 AS T1
                INNER JOIN Tabelle2 AS T2 ON T1.Meldender_ID = T2.Mitarbeiter_ID
                WHERE T1.KvS <> 0
                  AND T1.Krank_von <= ?
                  AND T1.Krank_bis  >= ?;
            """
            self.cursor.execute(query, (plan_end, plan_start))
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("Datenbankfehler", f"Fehler beim Abrufen der KvS-Daten: {e}")
            return []
    
    def save_dienstplan_from_df(self, df, show_message=True, replace_for_date=True):
        if not self.conn:
            if show_message:
                messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung.")
            return False
        
        create_table_sql = """
            CREATE TABLE Dienstplan (
                Dienstplan_ID AUTOINCREMENT PRIMARY KEY,
                Datum DATE,
                Namen TEXT(255),
                Dienst TEXT(255),
                Beginn DATETIME,
                Ende DATETIME
            )
        """
        insert_sql = "INSERT INTO Dienstplan (Datum, Namen, Dienst, Beginn, Ende) VALUES (?, ?, ?, ?, ?)"

        plan_date_value = None
        try:
            if 'Datum' in df.columns and len(df) > 0:
                raw = df.iloc[0]['Datum']
                if isinstance(raw, (datetime, pd.Timestamp)):
                    plan_date_value = raw.date()
                else:
                    plan_date_value = pd.to_datetime(str(raw), dayfirst=True, errors='coerce')
                    if isinstance(plan_date_value, pd.Timestamp):
                        plan_date_value = plan_date_value.date()
                    elif isinstance(plan_date_value, datetime):
                        plan_date_value = plan_date_value.date()
                    else:
                        plan_date_value = None
        except Exception:
            plan_date_value = None

        try:
            try:
                self.cursor.execute(create_table_sql)
                self.conn.commit()
            except pyodbc.ProgrammingError:
                pass

            if replace_for_date and plan_date_value is not None:
                try:
                    self.cursor.execute("DELETE FROM Dienstplan WHERE Datum = ?", (plan_date_value,))
                    self.conn.commit()
                except pyodbc.Error:
                    pass

            for _, row in df.iterrows():
                if 'deactivated' in df.columns and bool(row.get('deactivated', False)):
                    continue
                if isinstance(row['Datum'], (datetime, pd.Timestamp)):
                    datum = row['Datum'].date()
                else:
                    try:
                        datum = pd.to_datetime(str(row['Datum']), dayfirst=True, errors='coerce')
                        datum = datum.date() if isinstance(datum, pd.Timestamp) else None
                    except Exception:
                        datum = None

                def _to_dt(v):
                    if isinstance(v, (datetime, pd.Timestamp)):
                        return v
                    try:
                        t = pd.to_datetime(str(v), errors='coerce')
                        return t.to_pydatetime() if isinstance(t, pd.Timestamp) else None
                    except Exception:
                        return None

                beginn = _to_dt(row['Beginn'])
                ende   = _to_dt(row['Ende'])

                self.cursor.execute(insert_sql, (datum, row['Namen'], row['Dienst'], beginn, ende))

            self.conn.commit()
            if show_message:
                messagebox.showinfo("Erfolg", "Dienstplandaten erfolgreich in die Datenbank gespeichert!")
            return True

        except pyodbc.Error as e:
            self.conn.rollback()
            if show_message:
                messagebox.showerror("Datenbankfehler", f"Fehler beim Speichern der Dienstplandaten: {e}")
            return False

# -------------------------------------------------------------------
# Krankmeldung-Tab (mit KvS)
# -------------------------------------------------------------------
class KrankmeldungTab(ttk.Frame):
    def __init__(self, master=None, db=None, style_vars=None, **kwargs):
        super().__init__(master, **kwargs)
        self.db = db
        self.editing_id = None
        self.filter_manage_start_date = None
        self.filter_manage_end_date = None
        
        if style_vars:
            self.pastel_bg = style_vars.get('bg')
            self.pastel_fg = style_vars.get('fg')
            self.pastel_button = style_vars.get('button')
        else:
            self.pastel_bg = '#E0F7FA'
            self.pastel_fg = '#37474F'
            self.pastel_button = '#B2DFDB'
        
        self.configure(style='TFrame')
        self.style = ttk.Style(self)
        self.style.configure('TFrame', background=self.pastel_bg)
        self.style.configure('TLabel', background=self.pastel_bg, foreground=self.pastel_fg, font=('Segoe UI', 10))
        self.style.configure('TEntry', fieldbackground='white', foreground=self.pastel_fg, font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10, 'bold'), foreground=self.pastel_fg, background=self.pastel_button, relief='flat')
        self.style.map('TButton', background=[('active', '#80CBC4'), ('pressed', '#4DB6AC')])
        self.style.configure('Treeview', background='white', foreground=self.pastel_fg, fieldbackground='white', rowheight=25, font=('Segoe UI', 10))
        # FIX: kein falsches Anführungszeichen!
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background='#B2DFDB', foreground=self.pastel_fg)

        self.main_container = ttk.Frame(self, padding="15", style='TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.form_frame = ttk.Frame(self.main_container, style='TFrame')
        self.form_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        self.tables_frame = ttk.Frame(self.main_container, style='TFrame')
        self.tables_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.create_form_widgets()
        self.create_tables_view()
        self.initialize_form()
        if self.db and self.db.conn:
            self.display_data_in_tables()

    def create_form_widgets(self):
        input_frame = ttk.Frame(self.form_frame, style='TFrame')
        input_frame.pack(padx=10, pady=10, fill=tk.Y, expand=False)
        
        ttk.Label(input_frame, text="Datum der Meldung:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_datum = ttk.Entry(input_frame)
        self.txt_datum.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Label(input_frame, text="Meldender:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_meldender = ttk.Entry(input_frame)
        self.txt_meldender.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Krank von:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.btn_krank_von = ttk.Button(input_frame, text="", command=lambda: self.show_calendar("von"))
        self.btn_krank_von.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Krank bis:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.btn_krank_bis = ttk.Button(input_frame, text="", command=lambda: self.show_calendar("bis"))
        self.btn_krank_bis.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Anruf um:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_anruf_um = ttk.Entry(input_frame)
        self.txt_anruf_um.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Angenommen von:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_angenommen_von = ttk.Entry(input_frame)
        self.txt_angenommen_von.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Label(input_frame, text="Bemerkung:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_bemerkung = ttk.Entry(input_frame)
        self.txt_bemerkung.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Label(input_frame, text="Kommentar:").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
        self.txt_kommentar = tk.Text(input_frame, height=3, width=30)
        self.txt_kommentar.grid(row=7, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        self.kvs_var = tk.BooleanVar(value=False)
        self.chk_kvs = ttk.Checkbutton(input_frame, text="KvS (krank von Station)", variable=self.kvs_var, command=self._toggle_kvs_controls)
        self.chk_kvs.grid(row=9, column=0, sticky=tk.W, padx=5, pady=(10, 0))

        ttk.Label(input_frame, text="Gegangen um:").grid(row=9, column=1, sticky=tk.W, padx=5, pady=(10, 0))
        self.cbo_gegangen_um = ttk.Combobox(input_frame, values=generate_time_choices(15), state="disabled", width=8)
        self.cbo_gegangen_um.grid(row=9, column=1, sticky=(tk.E), padx=5, pady=(10, 0))
        self.cbo_gegangen_um.set("")

        self.save_button = ttk.Button(input_frame, text="Speichern", command=self.cmd_speichern_click)
        self.save_button.grid(row=10, column=0, sticky=tk.W, padx=5, pady=10)

        self.cancel_button = ttk.Button(input_frame, text="Abbrechen", command=self.clear_form)
        self.cancel_button.grid(row=10, column=1, sticky=tk.W, padx=5, pady=10)

    
        self.export_xls_button = ttk.Button(input_frame, text="Export Excel (Meldedatum)", command=self.export_krankmeldungen_to_excel)
        self.export_xls_button.grid(row=10, column=2, sticky=tk.W, padx=5, pady=10)

    def export_krankmeldungen_to_excel(self):
        """Exportiert alle Krankmeldungen in die .xlsm-Monatsdatei (Blatt 'Krankmeldungen').
        Duplikate werden anhand (Datum, Meldender, Anruf um) verhindert."""
        if not getattr(self, "db", None) or not getattr(self.db, "conn", None):
            return
        target_path = KRANK_EXPORT_XLSM_PATH
        if not os.path.exists(target_path):
            return
        try:
            from openpyxl import load_workbook
            from datetime import datetime
            wb = load_workbook(target_path, keep_vba=True)
            ws = wb["Krankmeldungen"] if "Krankmeldungen" in wb.sheetnames else wb.active

            existing = set()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or all(v is None for v in row[:6]):
                    continue
                d, person, _, _, t, _ = (row + (None,)*6)[:6]
                try:
                    if isinstance(d, datetime): d = d.date()
                    elif hasattr(d, "date"): d = d.date()
                except Exception: pass
                try:
                    if isinstance(t, datetime): t = t.time()
                except Exception: pass
                existing.add((d, (person or "").strip(), t))

            rows = self.db.fetch_all_krankmeldungen()
            if not rows: 
                return

            def to_date(x):
                try:
                    if isinstance(x, datetime): return x.date()
                    if hasattr(x, "date"): return x.date()
                except Exception: pass
                return x

            def to_time(x):
                try:
                    if isinstance(x, datetime): return x.time()
                except Exception: pass
                return x

            for r in rows:
                try:
                    datum = r[1]; meldender = r[2]; krank_von = r[3]; krank_bis = r[4]; anruf_um = r[5]; angenommen_von = r[6]; bemerkung = r[7]
                except Exception:
                    datum = getattr(r, "Datum", None)
                    meldender = getattr(r, "Meldender", "")
                    krank_von = getattr(r, "Krank_von", None)
                    krank_bis = getattr(r, "Krank_bis", None)
                    anruf_um = getattr(r, "Anruf_um", None)
                    angenommen_von = getattr(r, "Angenommen_von", "")
                    bemerkung = getattr(r, "Bemerkung", "")

                key = (to_date(datum), (meldender or "").strip(), to_time(anruf_um))
                if key in existing:
                    continue

                ws.append([
                    to_date(datum),
                    (meldender or ""),
                    to_date(krank_von),
                    to_date(krank_bis),
                    to_time(anruf_um),
                    (angenommen_von or ""),
                    (bemerkung or ""),
                    None
                ])
                existing.add(key)

            wb.save(target_path)
        except Exception:
            return
    def _toggle_kvs_controls(self):
        self.cbo_gegangen_um.configure(state=("normal" if self.kvs_var.get() else "disabled"))
        if not self.kvs_var.get():
            self.cbo_gegangen_um.set("")

    def create_tables_view(self):
        sick_list_frame = ttk.Frame(self.tables_frame, style='TFrame')
        sick_list_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=10)

        ttk.Label(sick_list_frame, text="Aktuell kranke Mitarbeiter:", font=('Segoe UI', 11, 'bold')).pack(pady=(0, 5))
        self.sick_employees_tree = ttk.Treeview(sick_list_frame, columns=("Name", "Von", "Bis"), show="headings", selectmode="none")
        self.sick_employees_tree.heading("Name", text="Name")
        self.sick_employees_tree.heading("Von", text="Krank von")
        self.sick_employees_tree.heading("Bis", text="Krank bis")
        self.sick_employees_tree.column("Name", width=250, anchor="w")
        self.sick_employees_tree.column("Von", width=150, anchor="center")
        self.sick_employees_tree.column("Bis", width=150, anchor="center")
        self.sick_employees_tree.pack(fill=tk.BOTH, expand=True)
        
        manage_frame = ttk.Frame(self.tables_frame, style='TFrame')
        manage_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(manage_frame, text="Datensätze verwalten:", font=('Segoe UI', 11, 'bold')).pack(pady=(0, 5))
        
        top_frame = ttk.Frame(manage_frame, style='TFrame')
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        filter_frame = ttk.Frame(top_frame, style='TFrame')
        filter_frame.pack(side=tk.LEFT, fill=tk.X)
        
        ttk.Label(filter_frame, text="Datensätze filtern:").pack(side=tk.LEFT, padx=(0, 5))
        self.btn_manage_filter_start = ttk.Button(filter_frame, text="Startdatum", command=lambda: self.show_calendar("manage_start"))
        self.btn_manage_filter_start.pack(side=tk.LEFT, padx=2)

        self.btn_manage_filter_end = ttk.Button(filter_frame, text="Enddatum", command=lambda: self.show_calendar("manage_end"))
        self.btn_manage_filter_end.pack(side=tk.LEFT, padx=2)

        ttk.Button(filter_frame, text="Filtern", command=self.filter_manage_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(filter_frame, text="Zurücksetzen", command=self.reset_manage_filter).pack(side=tk.LEFT, padx=5)

        ttk.Button(top_frame, text="Daten exportieren", command=self.export_data).pack(side=tk.RIGHT, padx=5)

        self.manage_tree = ttk.Treeview(
            manage_frame,
            columns=("ID", "Datum", "Meldender", "Von", "Bis", "Anruf", "Angenommen", "Bemerkung", "KvS", "Gegangen um"),
            show="headings"
        )
        headers = [
            ("ID", 40, "center"),
            ("Datum", 120, "center"),
            ("Meldender", 150, "w"),
            ("Von", 120, "center"),
            ("Bis", 120, "center"),
            ("Anruf", 80, "center"),
            ("Angenommen", 150, "w"),
            ("Bemerkung", 250, "w"),
            ("KvS", 60, "center"),
            ("Gegangen um", 100, "center")
        ]
        for name, width, anchor in headers:
            self.manage_tree.heading(name, text=name)
            self.manage_tree.column(name, width=width, anchor=anchor)
        self.manage_tree.pack(fill=tk.BOTH, expand=True)

        action_frame = ttk.Frame(manage_frame, style='TFrame')
        action_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(action_frame, text="Bearbeiten", command=self.edit_selected_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Löschen", command=self.delete_selected_entry).pack(side=tk.LEFT, padx=5)

    def initialize_form(self):
        self.txt_datum.delete(0, tk.END)
        self.txt_datum.insert(0, datetime.now().strftime("%d.%m.%Y"))
        self.txt_anruf_um.delete(0, tk.END)
        self.txt_anruf_um.insert(0, datetime.now().strftime("%H:%M"))
        self.btn_krank_von.config(text=datetime.now().strftime("%d.%m.%Y"))
        self.btn_krank_bis.config(text=datetime.now().strftime("%d.%m.%Y"))
        self.txt_meldender.delete(0, tk.END)
        self.txt_angenommen_von.delete(0, tk.END)
        self.txt_bemerkung.delete(0, tk.END)
        try:
            self.txt_kommentar.delete('1.0','end')
        except Exception:
            pass
        self.kvs_var.set(False)
        self.cbo_gegangen_um.set("")
        self._toggle_kvs_controls()
        self.editing_id = None
        self.save_button.config(text="Speichern")

    def show_calendar(self, button_type):
        top = tk.Toplevel(self.master)
        top.title("Datum auswählen")
        top.geometry("400x400")
        top.configure(bg=self.pastel_bg)

        cal = Calendar(top, selectmode="day", date_pattern="dd.mm.yyyy",
                       background=self.pastel_button, foreground=self.pastel_fg,
                       headersbackground='#80CBC4', selectbackground='#4DB6AC')
        cal.pack(padx=10, pady=10)

        def get_date():
            selected_date = cal.get_date()
            if button_type == "von":
                self.btn_krank_von.config(text=selected_date)
            elif button_type == "bis":
                self.btn_krank_bis.config(text=selected_date)
            elif button_type == "manage_start":
                self.btn_manage_filter_start.config(text=selected_date)
            elif button_type == "manage_end":
                self.btn_manage_filter_end.config(text=selected_date)
            top.destroy()
        
        ttk.Button(top, text="Auswählen", command=get_date).pack(pady=10)

    def _parse_time_hhmm(self, s):
        s = (s or "").strip()
        if not s:
            return None
        m = re.match(r'^(\d{1,2}):(\d{2})$', s)
        if not m:
            raise ValueError("Zeitformat muss HH:MM sein.")
        hh, mm = int(m.group(1)), int(m.group(2))
        if not (0 <= hh <= 23 and 0 <= mm <= 59):
            raise ValueError("Ungültige Uhrzeit.")
        now = datetime.now()
        return datetime(now.year, now.month, now.day, hh, mm)

    def cmd_speichern_click(self):
        if not self.db or not self.db.conn:
            messagebox.showerror("Fehler", "Bitte verbinden Sie sich zuerst mit einer Datenbank.")
            return

        d_datum_str = self.txt_datum.get()
        s_meldender = self.txt_meldender.get().strip()
        d_von_str = self.btn_krank_von.cget("text")
        d_bis_str = self.btn_krank_bis.cget("text")
        t_anruf_str = self.txt_anruf_um.get()
        s_angenommen = self.txt_angenommen_von.get().strip()
        s_bem = self.txt_bemerkung.get().strip()
        # Kommentar (optional) an Bemerkung anhängen
        kommentar_val = ""
        try:
            kommentar_val = self.txt_kommentar.get("1.0", "end").strip()
        except Exception:
            kommentar_val = ""
        if kommentar_val:
            s_bem = (s_bem + (" | Kommentar: " + kommentar_val)) if s_bem else ("Kommentar: " + kommentar_val)
        kvs_flag = bool(self.kvs_var.get())
        t_gegangen_str = self.cbo_gegangen_um.get().strip()
        
        if not all([d_datum_str, s_meldender, d_von_str, d_bis_str, t_anruf_str, s_angenommen]):
            messagebox.showerror("Eingabefehler", "Alle Felder außer 'Bemerkung' und 'Gegangen um' müssen ausgefüllt sein.")
            return

        try:
            d_datum = datetime.strptime(d_datum_str, "%d.%m.%Y")
            d_von = datetime.strptime(d_von_str, "%d.%m.%Y")
            d_bis = datetime.strptime(d_bis_str, "%d.%m.%Y")
            t_anruf = datetime.strptime(t_anruf_str, "%H:%M")
            t_gegangen = self._parse_time_hhmm(t_gegangen_str) if t_gegangen_str else None
        except ValueError as e:
            messagebox.showerror("Formatfehler", f"Ungültiges Datums- oder Zeitformat: {e}\nErwartete Formate: Datum (dd.mm.yyyy), Zeit (HH:MM)")
            return
            
        if self.editing_id is None:
            success = self.db.save_krankmeldung(d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=kvs_flag, t_gegangen=t_gegangen)
        try:
            self.export_krankmeldungen_to_excel()
        except Exception:
            pass
            if success:
                messagebox.showinfo("Erfolg", "Krankmeldung erfolgreich gespeichert!")
        else:
            success = self.db.update_krankmeldung(self.editing_id, d_datum, s_meldender, d_von, d_bis, t_anruf, s_angenommen, s_bem, kvs=kvs_flag, t_gegangen=t_gegangen)
            if success:
                messagebox.showinfo("Erfolg", "Krankmeldung erfolgreich aktualisiert!")

        if success:
            self.clear_form()
            self.display_data_in_tables()
            try:
                app = self.master.master
                if hasattr(app, "excel_viewer_tab") and app.excel_viewer_tab and app.excel_viewer_tab.df is not None:
                    app.excel_viewer_tab.recompute_sickness_from_db()
                if hasattr(app, "statistics_tab") and app.statistics_tab:
                    app.statistics_tab.update_statistics(app.excel_viewer_tab.df if app.excel_viewer_tab.df is not None else None)
            except Exception:
                pass

    def clear_form(self):
        self.initialize_form()

    def edit_selected_entry(self):
        selected_item = self.manage_tree.focus()
        if not selected_item:
            messagebox.showwarning("Keine Auswahl", "Bitte wählen Sie einen Eintrag zum Bearbeiten aus.")
            return
        values = self.manage_tree.item(selected_item, 'values')
        self.editing_id = values[0]

        self.txt_datum.delete(0, tk.END); self.txt_datum.insert(0, values[1])
        self.txt_meldender.delete(0, tk.END); self.txt_meldender.insert(0, values[2])
        self.btn_krank_von.config(text=values[3])
        self.btn_krank_bis.config(text=values[4])
        self.txt_anruf_um.delete(0, tk.END); self.txt_anruf_um.insert(0, values[5])
        self.txt_angenommen_von.delete(0, tk.END); self.txt_angenommen_von.insert(0, values[6])
        self.txt_bemerkung.delete(0, tk.END); self.txt_bemerkung.insert(0, values[7])

        self.kvs_var.set(str(values[8]).strip().lower() in ("-1", "1", "true", "ja"))
        self._toggle_kvs_controls()
        self.cbo_gegangen_um.set(values[9] if values[9] else "")

        self.save_button.config(text="Aktualisieren")

    def delete_selected_entry(self):
        selected_item = self.manage_tree.focus()
        if not selected_item:
            messagebox.showwarning("Keine Auswahl", "Bitte wählen Sie einen Eintrag zum Löschen aus.")
            return
        entry_id = self.manage_tree.item(selected_item, 'values')[0]
        if messagebox.askyesno("Löschen bestätigen", f"Möchten Sie Eintrag {entry_id} wirklich löschen?"):
            if self.db and self.db.delete_krankmeldung(entry_id):
                messagebox.showinfo("Erfolg", "Eintrag erfolgreich gelöscht.")
                self.display_data_in_tables()
                try:
                    app = self.master.master
                    if hasattr(app, "excel_viewer_tab") and app.excel_viewer_tab and app.excel_viewer_tab.df is not None:
                        app.excel_viewer_tab.recompute_sickness_from_db()
                    if hasattr(app, "statistics_tab") and app.statistics_tab:
                        app.statistics_tab.update_statistics(app.excel_viewer_tab.df if app.excel_viewer_tab.df is not None else None)
                except Exception:
                    pass

    def display_data_in_tables(self):
        if not self.db or not self.db.conn:
            self.manage_tree.delete(*self.manage_tree.get_children())
            self.sick_employees_tree.delete(*self.sick_employees_tree.get_children())
            self.manage_tree.insert('', 'end', values=("Keine Datenbankverbindung.", "", "", "", "", "", "", "", "", ""))
            self.sick_employees_tree.insert('', 'end', values=("Keine Krankmeldungen im Zeitraum", "", ""))
            return

        all_data = self.db.fetch_all_krankmeldungen_with_id()
        self.manage_tree.delete(*self.manage_tree.get_children())
        self.sick_employees_tree.delete(*self.sick_employees_tree.get_children())

        for row in all_data:
            formatted_row = list(row)
            formatted_row[1] = row[1].strftime("%d.%m.%Y")
            formatted_row[3] = row[3].strftime("%d.%m.%Y")
            formatted_row[4] = row[4].strftime("%d.%m.%Y")
            formatted_row[5] = row[5].strftime("%H:%M")
            formatted_row[8] = "Ja" if (row[8] in (-1, 1, True)) else "Nein"
            formatted_row[9] = row[9].strftime("%H:%M") if row[9] else ""
            self.manage_tree.insert('', 'end', values=formatted_row)
        
        self.update_current_sickness_view()

    def update_current_sickness_view(self, start_date=None, end_date=None):
        if not self.sick_employees_tree:
            return

        for item in self.sick_employees_tree.get_children():
            self.sick_employees_tree.delete(item)

        if start_date is None or end_date is None:
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=3)

        if not self.db or not self.db.conn:
            self.sick_employees_tree.insert("", "end", values=("Keine Datenbankverbindung.", "", ""))
            return

        sick_employees = self.db.fetch_sick_employees_in_range(start_date, end_date)

        if not sick_employees:
            self.sick_employees_tree.insert("", "end", values=("Keine Krankmeldungen im Zeitraum", "", ""))
            return

        for name, krank_von, krank_bis in sick_employees:
            von_str = krank_von.strftime("%d.%m.%Y") if isinstance(krank_von, (datetime, pd.Timestamp)) else str(krank_von)
            bis_str = krank_bis.strftime("%d.%m.%Y") if isinstance(krank_bis, (datetime, pd.Timestamp)) else str(krank_bis)
            self.sick_employees_tree.insert("", "end", values=(name, von_str, bis_str))

    def filter_manage_data(self):
        if not self.db or not self.db.conn:
            messagebox.showwarning("Fehler", "Bitte verbinden Sie sich zuerst mit einer Datenbank.")
            return
        start_date_str = self.btn_manage_filter_start.cget("text")
        end_date_str = self.btn_manage_filter_end.cget("text")
        if start_date_str == "Startdatum" or end_date_str == "Enddatum":
            messagebox.showwarning("Fehler", "Bitte wählen Sie sowohl ein Start- als auch ein Enddatum aus.")
            return
        self.filter_manage_start_date = self.parse_date_de(start_date_str)
        self.filter_manage_end_date = self.parse_date_de(end_date_str)
        if not self.filter_manage_start_date or not self.filter_manage_end_date:
            messagebox.showwarning("Fehler", "Ungültiges Datumsformat.")
            return
        if self.filter_manage_end_date < self.filter_manage_start_date:
            messagebox.showwarning("Fehler", "Das Enddatum muss nach dem Startdatum liegen.")
            return
        self.update_table_view(self.manage_tree, self.db.fetch_all_krankmeldungen_with_id(self.filter_manage_start_date, self.filter_manage_end_date))
        
    def reset_manage_filter(self):
        self.filter_manage_start_date = None
        self.filter_manage_end_date = None
        self.btn_manage_filter_start.config(text="Startdatum")
        self.btn_manage_filter_end.config(text="Enddatum")
        self.display_data_in_tables()
    
    def update_table_view(self, treeview, data):
        for item in treeview.get_children():
            treeview.delete(item)
        if not data:
            treeview.insert("", "end", values=("Keine Daten im Zeitraum gefunden.", "", "", "", "", "", "", "", "", ""))
            return
        for row in data:
            formatted_row = list(row)
            formatted_row[1] = row[1].strftime("%d.%m.%Y")
            formatted_row[3] = row[3].strftime("%d.%m.%Y")
            formatted_row[4] = row[4].strftime("%d.%m.%Y")
            formatted_row[5] = row[5].strftime("%H:%M")
            formatted_row[8] = "Ja" if (row[8] in (-1, 1, True)) else "Nein"
            formatted_row[9] = row[9].strftime("%H:%M") if row[9] else ""
            treeview.insert('', 'end', values=formatted_row)

    def export_data(self):
        if not self.db or not self.db.conn:
            messagebox.showwarning("Exportfehler", "Keine Datenbankverbindung vorhanden.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-Dateien", "*.xlsx")], title="Daten exportieren")
        if not file_path:
            return
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Krankmeldungen"
            columns = ("ID", "Datum", "Meldender", "Krank von", "Krank bis", "Anruf um", "Angenommen von", "Bemerkung", "KvS", "Gegangen um")
            sheet.append(columns)
            for row_id in self.manage_tree.get_children():
                sheet.append(self.manage_tree.item(row_id)['values'])
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for column in sheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_length:
                        max_length = len(val)
                sheet.column_dimensions[get_column_letter(column[0].column)].width = max_length + 2
            workbook.save(file_path)
            messagebox.showinfo("Export erfolgreich", f"Daten wurden erfolgreich nach '{file_path}' exportiert.")
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Ein Fehler ist beim Exportieren der Daten aufgetreten: {e}")

    def parse_date_de(self, date_str):
        try:
            return datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            return None

# -------------------------------------------------------------------
# Excel-Viewer-Tab (Tagesdienstplan)
# -------------------------------------------------------------------
class ExcelViewerTab(tk.Frame):
    def __init__(self, master=None, db=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.db = db
        self.tree = None
        self.df = None
        self.file_path = None
        self.original_file_path = None
        self.header_row_index = None
        self.edit_entry = None
        self.edit_item = None
        self.edit_column = None
        self.kvs_time_map = {}  # Nachname -> "KvS HH:MM (von DD.MM.YYYY)"
        self.create_widgets()

    def update_treeview_from_df(self):
        
        # Tree leeren
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.df is None or self.df.empty:
            return

        # Spalte 'deactivated' sicherstellen
        if 'deactivated' not in self.df.columns:
            self.df['deactivated'] = False

        # Einträge aufbauen
        for index, row in self.df.iterrows():
            # Stelle sicher, dass erwartete Spalten existieren
            for col in ['Datum','Namen','Dienst','Beginn','Ende']:
                if col not in self.df.columns:
                    return

            row_values = [row.get('Datum',''),
                          row.get('Namen',''),
                          row.get('Dienst',''),
                          row.get('Beginn',''),
                          row.get('Ende','')]

            # Zeiten schön formatiert (ohne Sekunden)
            row_values[3] = format_time(row_values[3])
            row_values[4] = format_time(row_values[4])

            # KvS-Text anhängen
            try:
                lastname = extract_lastname(row.get('Namen',''))
            except Exception:
                lastname = ''
            kvs_text = getattr(self, 'kvs_time_map', {}).get(lastname, "")
            row_values.append(kvs_text)  # -> Spalte 'KvS'

            # Tags bestimmen
            tags = []
            if 'is_sick' in self.df.columns and bool(row.get('is_sick', False)):
                tags.append('sick_employee')
            if bool(row.get('deactivated', False)) or (index in getattr(self, 'deactivated_set', set())):
                tags.append('deactivated')

            # Zeile einfügen
            try:
                self.tree.insert("", "end", values=row_values, iid=index, tags=tuple(tags))
            except Exception:
                # Fallback ohne iid
                self.tree.insert("", "end", values=row_values, tags=tuple(tags))



    def create_widgets(self):
        frame = tk.Frame(self)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        h_scrollbar = ttk.Scrollbar(frame, orient="horizontal")
        v_scrollbar = ttk.Scrollbar(frame, orient="vertical")
        
        self.tree = ttk.Treeview(frame, xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set, selectmode='browse')
        # Configure columns & headings for the day plan
        try:
            self.tree['columns'] = ('Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS')
            self.tree['show'] = 'headings'
            for col, width in [('Datum', 110), ('Namen', 220), ('Dienst', 80), ('Beginn', 80), ('Ende', 80), ('KvS', 80)]:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=width, anchor='center')
        except Exception:
            pass
        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.tag_configure('sick_employee', foreground='red')

        # ---- Deaktivieren: Kontextmenü und Tag ----
        self.deactivated_set = set()
        self.tree.tag_configure('deactivated', foreground='#9E9E9E')
        self.ctx_menu = tk.Menu(self.tree, tearoff=0)
        self.ctx_menu.add_command(label="Mitarbeiter deaktivieren", command=self._ctx_deactivate)
        self.ctx_menu.add_command(label="Mitarbeiter wieder aktivieren", command=self._ctx_activate)
        # Rechtsklick-Binding
        self.tree.bind("<Button-3>", self._on_right_click)

        h_scrollbar.config(command=self.tree.xview)
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.pack(side='bottom', fill='x')
        v_scrollbar.pack(side='right', fill='y')
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
        container_bg = '#E0F7FA'
        btn_width = 25
        btn_font = ("Segoe UI", 10, "bold")
        btn_bg = "#B2DFDB"
        btn_fg = "#37474F"

        button_frame = tk.Frame(self, bg=container_bg)
        button_frame.pack(pady=10, fill="x")

        word_frame = tk.Frame(button_frame, bg=container_bg)
        word_frame.pack(side=tk.TOP, pady=5)

        self.word_button = tk.Button(
            word_frame, text="📄 Word-Dokument erstellen",
            command=self.create_word_from_current_plan,  # <- Methode ist unten vorhanden
            state=tk.DISABLED,
            width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge"
        )
        self.word_button.pack(side=tk.LEFT, padx=5)

        email_frame = tk.Frame(button_frame, bg=container_bg)
        email_frame.pack(side=tk.TOP, pady=5)

        self.email_button = tk.Button(
            email_frame, text="✉️ E-Mail-Entwurf (mit Word)",
            command=self.email_draft_from_current_plan,
            state=tk.DISABLED,
            width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge"
        )
        self.email_button.pack(side=tk.LEFT, padx=5)

        tools_frame = tk.Frame(button_frame, bg=container_bg)
        tools_frame.pack(side=tk.TOP, pady=5)

        self.mark_sick_button = tk.Button(tools_frame, text="Krankmeldungen markieren",
                                          command=self.mark_sick_employees, state=tk.DISABLED,
                                          width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.mark_sick_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(tools_frame, text="Daten in Datenbank speichern",
                                     command=self.save_to_db_click, state=tk.DISABLED,
                                     width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.export_button = tk.Button(tools_frame, text="Exportieren",
                                       command=self.export_to_excel, state=tk.DISABLED,
                                       width=btn_width, font=btn_font, bg=btn_bg, fg=btn_fg, relief="ridge")
        self.export_button.pack(side=tk.LEFT, padx=5)

    def load_data(self, file_path):
        self.file_path = file_path
        self.original_file_path = file_path
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)

            df_raw = pd.read_excel(file_path, header=None, engine="openpyxl")
            required_columns = ["Name", "Dienst", "Beginn", "Ende"]
            self.header_row_index = None

            for i in range(15):
                if set(required_columns).issubset(set(df_raw.iloc[i].dropna().tolist())):
                    self.header_row_index = i
                    break
            
            if self.header_row_index is None:
                messagebox.showerror("Fehler", "Die benötigten Spaltenüberschriften (Name, Dienst, Beginn, Ende) konnten in den ersten 15 Zeilen nicht gefunden werden.")
                self.df = None
                for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                    b.config(state=tk.DISABLED)
                return False

            df = pd.read_excel(file_path, header=self.header_row_index, engine="openpyxl")
            df = df[['Name', 'Dienst', 'Beginn', 'Ende']]
            df.dropna(subset=['Name', 'Dienst', 'Beginn', 'Ende'], inplace=True)
            df.rename(columns={'Name': 'Namen'}, inplace=True)

            df['OrigDienst'] = df['Dienst']
            df['OrigEnde'] = df['Ende']
            df['is_sick'] = False

            df['is_sick'] = df['is_sick'] | df['Dienst'].astype(str).str.strip().str.lower().str.startswith(('k', 'krank'))
            df.loc[df['is_sick'], 'Dienst'] = 'K'

            filename = os.path.basename(file_path)
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
            file_date_str = date_match.group(1) if date_match else "Unbekanntes Datum"
            df.insert(0, 'Datum', file_date_str)
            
            self.df = df
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            
            columns = ['Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS']
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=110 if col in ('Beginn', 'Ende', 'KvS') else 150 if col == 'Namen' else 100, anchor="w")
                
            self.update_treeview_from_df()

            if self.db and self.db.conn:
                self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)

            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.NORMAL)

            if self.db and self.db.conn:
                self.recompute_sickness_from_db()
            return True
        except KeyError as e:
            self.df = None
            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.DISABLED)
            messagebox.showerror("Fehler beim Laden der Datei", f"Die Datei enthält nicht die erforderlichen Spalten. Es fehlt: {e}")
            return False
        except Exception as e:
            self.df = None
            for b in (self.save_button, self.mark_sick_button, self.export_button, self.word_button, self.email_button):
                b.config(state=tk.DISABLED)
            messagebox.showerror("Fehler beim Laden der Datei", f"Ein Fehler ist aufgetreten: {e}\n\nStellen Sie sicher, dass die Datei ein gültiges Excel-Format hat.")
            return False

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        self.edit_item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        self.edit_column = int(column_id[1:]) - 1
        if self.edit_column in [1, 2]:
            x, y, width, height = self.tree.bbox(self.edit_item, column_id)
            current_value = self.tree.item(self.edit_item, "values")[self.edit_column]
            self.edit_entry = ttk.Entry(self.tree)
            self.edit_entry.place(x=x, y=y, width=width, height=height)
            self.edit_entry.insert(0, current_value)
            self.edit_entry.focus_set()
            self.edit_entry.bind("<Return>", self.on_edit_finish)
            self.edit_entry.bind("<FocusOut>", self.on_edit_finish)
        else:
            messagebox.showwarning("Bearbeitung nicht möglich", "Diese Spalte kann nicht direkt im Viewer bearbeitet werden.")

    def on_edit_finish(self, event):
        if not self.edit_entry:
            return
        new_value = self.edit_entry.get()
        row_index = int(self.edit_item)
        column_name = self.tree['columns'][self.edit_column]
        if column_name in self.df.columns:
            self.df.loc[row_index, column_name] = new_value
            current_values = list(self.tree.item(self.edit_item, "values"))
            current_values[self.edit_column] = new_value
            self.tree.item(self.edit_item, values=current_values)
        self.edit_entry.destroy(); self.edit_entry = None; self.edit_item = None; self.edit_column = None

        try:
            if self.db and self.db.conn and self.df is not None:
                self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)
        except Exception:
            pass

    def save_to_db_click(self):
        if self.db and self.db.conn and self.df is not None:
            self.db.save_dienstplan_from_df(self.df, show_message=True, replace_for_date=True)
        else:
            messagebox.showwarning("Fehler", "Keine gültigen Daten zum Speichern oder keine aktive Datenbankverbindung.")
    
    def recompute_sickness_from_db(self):
        if not self.db or not self.db.conn or self.df is None:
            return
        try:
            filename = os.path.basename(self.file_path) if self.file_path else ""
            date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
            plan_date = datetime.strptime(date_match.group(1), "%d.%m.%Y") if date_match else datetime.today()
            end_date = plan_date.replace(hour=23, minute=59, second=59, microsecond=999999)

            sick_employees = self.db.fetch_sick_employees_in_range(plan_date, end_date)
            sick_lastnames = {extract_lastname(name) for name, _, _ in sick_employees}

            kvs_rows = self.db.fetch_kvs_for_date(plan_date)
            kvs_map = {}
            for name, t_gegangen, d_von, d_bis in kvs_rows:
                if d_von is None or d_bis is None:
                    continue
                key = extract_lastname(name)
                hhmm = ""
                if t_gegangen:
                    try:
                        if isinstance(t_gegangen, datetime):
                            hhmm = f"{t_gegangen.hour:02d}:{t_gegangen.minute:02d}"
                        else:
                            hh, mm = map(int, str(t_gegangen).split(":")[:2])
                            hhmm = f"{hh:02d}:{mm:02d}"
                    except Exception:
                        hhmm = ""
                try:
                    von_dt = pd.to_datetime(d_von)
                    von_str = von_dt.strftime("%d.%m.%Y")
                except Exception:
                    continue
                kvs_map[key] = f"KvS {hhmm} (von {von_str})" if hhmm else f"KvS (von {von_str})"

            self.kvs_time_map = kvs_map

            self.df['Dienst'] = self.df['OrigDienst']
            self.df['Ende']   = self.df['OrigEnde']

            excel_sick = self.df['Dienst'].astype(str).str.strip().str.lower().str.startswith(('k', 'krank'))
            self.df['is_sick'] = excel_sick.copy()
            self.df.loc[self.df['is_sick'], 'Dienst'] = 'K'

            for index, row in self.df.iterrows():
                lastname = extract_lastname(row['Namen'])
                if lastname in sick_lastnames:
                    self.df.at[index, 'is_sick'] = True
                    self.df.at[index, 'Dienst'] = 'K'

            self.update_treeview_from_df()

            try:
                if self.db and self.db.conn and self.df is not None:
                    self.db.save_dienstplan_from_df(self.df, show_message=False, replace_for_date=True)
            except Exception:
                pass

            try:
                app = self.master.master
                if hasattr(app, "statistics_tab") and app.statistics_tab:
                    app.statistics_tab.update_statistics(self.df)
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Fehler beim Aktualisieren", f"Krank-Status/KvS konnte nicht neu berechnet werden: {e}")

    def mark_sick_employees(self):
        if not self.db or not self.db.conn:
            messagebox.showerror("Fehler", "Keine aktive Datenbankverbindung vorhanden.")
            return
        if self.df is None:
            messagebox.showwarning("Fehler", "Bitte laden Sie zuerst eine Excel-Datei.")
            return
        if 'OrigDienst' not in self.df.columns:
            self.df['OrigDienst'] = self.df['Dienst']
        if 'OrigEnde' not in self.df.columns:
            self.df['OrigEnde'] = self.df['Ende']
        self.recompute_sickness_from_db()
        messagebox.showinfo("Erfolg", "Dienstplan wurde mit den aktuellen Krankmeldungen abgeglichen (KvS-Kommentar im Zeitraum).")

    
       
    def _tree_row_iid_at(self, event):
        row_iid = self.tree.identify_row(event.y)
        return row_iid if row_iid != "" else None

    def _on_right_click(self, event):
        iid = self._tree_row_iid_at(event)
        if iid is None:
            return
        self.tree.selection_set(iid)
        self._last_rightclick_iid = iid
        idx = int(iid)
        if idx in getattr(self, 'deactivated_set', set()):
            self.ctx_menu.entryconfig(0, state="disabled")
            self.ctx_menu.entryconfig(1, state="normal")
        else:
            self.ctx_menu.entryconfig(0, state="normal")
            self.ctx_menu.entryconfig(1, state="disabled")
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _ctx_deactivate(self):
        iid = getattr(self, "_last_rightclick_iid", None)
        if iid is None:
            return
        idx = int(iid)
        if not hasattr(self, 'deactivated_set'):
            self.deactivated_set = set()
        self.deactivated_set.add(idx)
        if self.df is not None:
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            self.df.at[idx, 'deactivated'] = True
        self.update_treeview_from_df()
        try:
            app = self.master.master
            if hasattr(app, "statistics_tab") and app.statistics_tab:
                app.statistics_tab.update_statistics(self.df)
        except Exception:
            pass

    def _ctx_activate(self):
        iid = getattr(self, "_last_rightclick_iid", None)
        if iid is None:
            return
        idx = int(iid)
        if not hasattr(self, 'deactivated_set'):
            self.deactivated_set = set()
        if idx in self.deactivated_set:
            self.deactivated_set.discard(idx)
        if self.df is not None:
            if 'deactivated' not in self.df.columns:
                self.df['deactivated'] = False
            self.df.at[idx, 'deactivated'] = False
        self.update_treeview_from_df()
        try:
            app = self.master.master
            if hasattr(app, "statistics_tab") and app.statistics_tab:
                app.statistics_tab.update_statistics(self.df)
        except Exception:
            pass
    
    def export_to_excel(self):
        if self.df is None:
            messagebox.showwarning("Fehler", "Bitte laden Sie zuerst eine Excel-Datei.")
            return
        save_path = filedialog.asksaveasfilename(title="Dienstplan speichern unter", defaultextension=".xlsx", filetypes=[("Excel-Dateien", "*.xlsx")])
        if not save_path:
            return
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Dienstplan"
            headers = ['Datum', 'Namen', 'Dienst', 'Beginn', 'Ende', 'KvS']
            sheet.append(headers)
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for idx, row in self.df.iterrows():
                beg = format_time(row['Beginn'])
                end = format_time(row['Ende'])
                lastname = extract_lastname(row['Namen'])
                kvs_text = self.kvs_time_map.get(lastname, "")
                sheet.append([row['Datum'], row['Namen'], row['Dienst'], beg, end, kvs_text])
                if 'is_sick' in self.df.columns and self.df.loc[idx, 'is_sick']:
                    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    for cell in sheet[sheet.max_row]:
                        cell.fill = red_fill
            for column in sheet.columns:
                max_length = 0
                column_cells = [cell for cell in column]
                for cell in column_cells:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_length:
                        max_length = len(val)
                sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = max_length + 2
            workbook.save(save_path)
            messagebox.showinfo("Export erfolgreich", f"Daten wurden erfolgreich nach '{save_path}' exportiert.")
        except Exception as e:
            messagebox.showerror("Exportfehler", f"Ein Fehler ist beim Exportieren aufgetreten: {e}")

    # -------- Word & E-Mail aus aktuellem Plan --------
    def create_word_from_current_plan(self, save_path=None, return_path=False):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Hinweis", "Kein Tagesdienstplan geladen.")
            return None if return_path else None

        if not save_path:
            save_path = filedialog.asksaveasfilename(
                title="Word-Dokument speichern",
                defaultextension=".docx",
                filetypes=[("Word-Dokumente", "*.docx")]
            )
            if not save_path:
                return None if return_path else None

        filename = os.path.basename(self.file_path) if self.file_path else ""
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        if date_match:
            try:
                plan_date = datetime.strptime(date_match.group(1), "%d.%m.%Y")
            except Exception:
                plan_date = datetime.today()
        else:
            plan_date = datetime.today()

        paxzahl = simpledialog.askinteger(
            "PAX-Zahl eingeben",
            "Bitte geben Sie die PAX-Zahl für den Dienstplan ein:",
            minvalue=0, parent=self
        )
        if paxzahl is None:
            return None if return_path else None

        try:
            build_word_from_df(self.df, plan_date, save_path, paxzahl=paxzahl)
            if not return_path:
                messagebox.showinfo("Fertig", f"Dokument gespeichert unter:\n{save_path}")
                return None
            else:
                return save_path
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Erstellen des Word-Dokuments: {e}")
            return None

    def email_draft_from_current_plan(self):
        save_path = self.create_word_from_current_plan(return_path=True)
        if not save_path:
            return

        filename = os.path.basename(self.file_path) if self.file_path else ""
        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', filename)
        if date_match:
            try:
                d0 = datetime.strptime(date_match.group(1), "%d.%m.%Y")
                d1 = d0 + timedelta(days=1)
            except Exception:
                d0 = datetime.today(); d1 = d0 + timedelta(days=1)
        else:
            d0 = datetime.today(); d1 = d0 + timedelta(days=1)
        date_str = f"{d0.strftime('%d.%m.%Y')} – {d1.strftime('%d.%m.%Y')}"

        subject = f"Dienstplan für {date_str}"
        body = f"Hallo,\n\nanbei der Dienstplan für {date_str} als Anhang.\n\nMit freundlichen Grüßen,\nDRK Flughafen"

        try:
            if OUTLOOK_AVAILABLE:
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = CONFIG["email_recipients"]["to"]
                mail.CC = CONFIG["email_recipients"]["cc"]
                mail.Subject = subject
                mail.Body = body
                mail.Attachments.Add(os.path.abspath(save_path))
                mail.Display(True)
            else:
                body_mailto = body + f"\n\nPfad zur Datei (manuell anhängen): {save_path}"
                mailto_url = (f"mailto:?to={urllib.parse.quote(CONFIG['email_recipients']['to'])}"
                              f"&cc={urllib.parse.quote(CONFIG['email_recipients']['cc'])}"
                              f"&subject={urllib.parse.quote(subject)}"
                              f"&body={urllib.parse.quote(body_mailto)}")
                webbrowser.open(mailto_url)
        except Exception as e:
            messagebox.showerror("Fehler", f"E-Mail-Entwurf konnte nicht erstellt werden: {e}")

# -------------------------------------------------------------------
# Neuer Tab: Sonderaufgaben (neue Excel-Datei)
# -------------------------------------------------------------------

class SonderaufgabenTab(tk.Frame):
    AUFGABEN = [
        "Sauberkeit Station",
        "BTW Check + Sauberkeit",
        "E - mobby Check",
        "Bulmor 1 - 7312",
        "Bulmor 2 - 7892",
        "Bulmor 3 - 8092",
        "Bulmor 4 - 8794",
        "Bulmor 5 - 9982",
    ]

    NAME_HEADERS = {"name","namen","nachname","lastname","surname"}
    DIENST_HEADERS = {"dienst","schicht","shift"}

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        # sanfte Hintergrundfarbe wie im restlichen GUI
        try:
            self.configure(bg="#E0F7FA")
        except Exception:
            pass
        self.entries_tag = {}
        self.entries_nacht = {}
        # Zweite Combobox-Maps nur für Bulmor-Aufgaben
        self.entries_tag_2 = {}
        self.entries_nacht_2 = {}

        # Gesamtliste (Kompatibilität / manuelle Auswahl)
        self.namen_liste = []
        # Schicht-spezifische Listen
        self.namen_liste_tag = []
        self.namen_liste_nacht = []
        self.namen_liste_bulmor_tag = []
        self.namen_liste_bulmor_nacht = []

        self._build_ui()

    # --------- Color helpers (yellow detection) ---------
    def _hex_to_rgb(self, hexstr):
        s = str(hexstr or "").strip().lstrip("#")
        if len(s) == 8: s = s[2:]
        if len(s) != 6: return None
        try: return (int(s[0:2],16), int(s[2:4],16), int(s[4:6],16))
        except Exception: return None

    def _apply_tint(self, rgb, tint):
        if not rgb: return None
        r,g,b = rgb
        if tint in (None,0): return (r,g,b)
        if tint > 0:
            r = int(round(r + (255 - r) * tint))
            g = int(round(g + (255 - g) * tint))
            b = int(round(b + (255 - b) * tint))
        else:
            r = int(round(r * (1 + tint)))
            g = int(round(g * (1 + tint)))
            b = int(round(b * (1 + tint)))
        return (max(0,min(255,r)), max(0,min(255,g)), max(0,min(255,b)))

    def _color_obj_to_rgb(self, wb, color_obj):
        try:
            if not color_obj: return None
            tint = getattr(color_obj, "tint", None)
            if getattr(color_obj, "rgb", None):
                rgb = self._hex_to_rgb(color_obj.rgb)
                return self._apply_tint(rgb, tint) if rgb else None
            if getattr(color_obj, "theme", None) is not None:
                return None
            return None
        except Exception:
            return None

    def _get_effective_rgb(self, wb, fill):
        try:
            if not fill: return None
            pattern = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
            if str(pattern).lower() != "solid": return None
            for c in (getattr(fill,'fgColor',None), getattr(fill,'start_color',None), getattr(fill,'bgColor',None)):
                rgb = self._color_obj_to_rgb(wb, c)
                if rgb: return rgb
            return None
        except Exception:
            return None

    def _is_yellowish(self, rgb):
        if not rgb: return False
        r,g,b = rgb
        if r >= 180 and g >= 180 and b <= 150: return True
        try:
            import colorsys
            h,s,v = colorsys.rgb_to_hsv(r/255.0, g/255.0, b/255.0)
            deg = h*360
            return (25 <= deg <= 70) and s >= 0.25 and v >= 0.6
        except Exception:
            return False

    # --------------------------- UI ---------------------------
    
    def _build_ui(self):
        try:
            title = tk.Label(self, text="Sonderaufgaben – Eingabe (Tag/Nacht)",
                             font=("Segoe UI", 14, "bold"), bg="#E0F7FA", fg="#37474F")
            title.pack(pady=(10,5))
        except Exception:
            pass

        table = tk.Frame(self)
        try: table.configure(bg="#E0F7FA")
        except Exception: pass
        table.pack(padx=10, pady=10, fill="x")

        header_font = ("Segoe UI", 10, "bold")
        tk.Label(table, text="Aufgabe", font=header_font, anchor="w", width=35)\
            .grid(row=0, column=0, padx=5, pady=4, sticky="w")
        tk.Label(table, text="Tagschicht", font=header_font)\
            .grid(row=0, column=1, padx=5, pady=4)
        tk.Label(table, text="Nachtschicht", font=header_font)\
            .grid(row=0, column=2, padx=5, pady=4)

        for i, aufgabe in enumerate(self.AUFGABEN, start=1):
            tk.Label(table, text=aufgabe, anchor="w")\
                .grid(row=i, column=0, padx=5, pady=3, sticky="w")

            is_bul = self._is_bulmor(aufgabe)

            # Tag
            tag_cell = tk.Frame(table, bg=table.cget("bg"))
            tag_cell.grid(row=i, column=1, padx=5, pady=3, sticky="w")
            if is_bul:
                cb_t1 = ttk.Combobox(tag_cell, width=18, values=[""], state="normal")
                cb_t2 = ttk.Combobox(tag_cell, width=18, values=[""], state="normal")
                cb_t1.pack(side=tk.LEFT, padx=(0,4)); cb_t2.pack(side=tk.LEFT, padx=0)
                self.entries_tag[aufgabe] = cb_t1          # Backward-compatible primary
                self.entries_tag_2[aufgabe] = cb_t2        # Secondary
            else:
                cb_t = ttk.Combobox(tag_cell, width=22, values=[""], state="normal")
                cb_t.pack(side=tk.LEFT)
                self.entries_tag[aufgabe] = cb_t
                self.entries_tag_2[aufgabe] = None

            # Nacht
            nacht_cell = tk.Frame(table, bg=table.cget("bg"))
            nacht_cell.grid(row=i, column=2, padx=5, pady=3, sticky="w")
            if is_bul:
                cb_n1 = ttk.Combobox(nacht_cell, width=18, values=[""], state="normal")
                cb_n2 = ttk.Combobox(nacht_cell, width=18, values=[""], state="normal")
                cb_n1.pack(side=tk.LEFT, padx=(0,4)); cb_n2.pack(side=tk.LEFT, padx=0)
                self.entries_nacht[aufgabe] = cb_n1        # Backward-compatible primary
                self.entries_nacht_2[aufgabe] = cb_n2      # Secondary
            else:
                cb_n = ttk.Combobox(nacht_cell, width=22, values=[""], state="normal")
                cb_n.pack(side=tk.LEFT)
                self.entries_nacht[aufgabe] = cb_n
                self.entries_nacht_2[aufgabe] = None
        btns = tk.Frame(self)
        try: btns.configure(bg="#E0F7FA")
        except Exception: pass
        btns.pack(padx=10, pady=(0,10), fill="x")
        ttk.Button(btns, text="Excel (Namen) laden", command=self.load_names_from_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Dropdowns aktualisieren", command=self.update_name_choices_by_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Export nach Excel", command=self.export_sonderaufgaben_to_excel).pack(side=tk.LEFT, padx=5)

        # Hinweistext
        hint = tk.Label(self, text="Nur gelb markierte Mitarbeiter werden als Bulmorfahrer identifiziert",
                        font=("Segoe UI", 9, "italic"),
                        bg="#E0F7FA", fg="#555555", anchor="w", justify="left")
        hint.pack(padx=10, pady=(5, 10), anchor="w")
        self._init_sa_browser(self)


    def _extract_lastname_robust(self, full_name: str) -> str:
        s = str(full_name or "").strip()
        if not s: return ""
        if "," in s: s = s.split(",",1)[0].strip()
        parts = [p for p in s.split() if p]
        if not parts: return s
        if len(parts) == 1: return parts[0]
        particles = {"von","van","de","der","den","di","da","le","du"}
        if len(parts) >= 2 and parts[-2].lower() in particles:
            return parts[-2] + " " + parts[-1]
        return parts[-1]

    def _is_bulmor(self, aufgabe: str) -> bool:
        return "bulmor" in (aufgabe or "").lower()

    
    def update_name_choices_by_task(self):
        tag_vals_all   = [""] + (self.namen_liste_tag if self.namen_liste_tag else [])
        nacht_vals_all = [""] + (self.namen_liste_nacht if self.namen_liste_nacht else [])

        tag_vals_bulmor   = [""] + (self.namen_liste_bulmor_tag if self.namen_liste_bulmor_tag else self.namen_liste_tag or [])
        nacht_vals_bulmor = [""] + (self.namen_liste_bulmor_nacht if self.namen_liste_bulmor_nacht else self.namen_liste_nacht or [])

        for aufgabe in self.AUFGABEN:
            is_bul = self._is_bulmor(aufgabe)
            vals_tag = tag_vals_bulmor if is_bul else tag_vals_all
            vals_nacht = nacht_vals_bulmor if is_bul else nacht_vals_all

            # Tag primary
            cb_t1 = self.entries_tag.get(aufgabe)
            if cb_t1 is not None:
                try: cb_t1.configure(values=vals_tag)
                except Exception: pass
            # Tag secondary (Bulmor)
            cb_t2 = self.entries_tag_2.get(aufgabe)
            if cb_t2 is not None:
                try: cb_t2.configure(values=vals_tag)
                except Exception: pass

            # Nacht primary
            cb_n1 = self.entries_nacht.get(aufgabe)
            if cb_n1 is not None:
                try: cb_n1.configure(values=vals_nacht)
                except Exception: pass
            # Nacht secondary (Bulmor)
            cb_n2 = self.entries_nacht_2.get(aufgabe)
            if cb_n2 is not None:
                try: cb_n2.configure(values=vals_nacht)
                except Exception: pass
    def _find_headers_same_row(self, ws):
        # exakte Header in der gleichen Zeile
        for r in range(1, min(ws.max_row, 60)+1):
            labels = {}
            for c in range(1, ws.max_column+1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str):
                    labels[c] = v.strip().lower()
            if not labels: continue
            name_cols = [c for c,low in labels.items() if low in self.NAME_HEADERS]
            dienst_cols = [c for c,low in labels.items() if low in self.DIENST_HEADERS]
            if name_cols and dienst_cols:
                return r, name_cols[0], dienst_cols[0]
        return None, None, None

    
    def get_task_names(self, aufgabe: str):
        """Liefert (tag_namen, nacht_namen) jeweils als Liste mit 1 oder 2 Einträgen (leere Strings werden gefiltert)."""
        res_tag, res_nacht = [], []
        cb_t1 = self.entries_tag.get(aufgabe); cb_t2 = self.entries_tag_2.get(aufgabe)
        cb_n1 = self.entries_nacht.get(aufgabe); cb_n2 = self.entries_nacht_2.get(aufgabe)
        for cb in (cb_t1, cb_t2):
            try:
                val = cb.get().strip() if cb else ""
                if val: res_tag.append(val)
            except Exception:
                pass
        for cb in (cb_n1, cb_n2):
            try:
                val = cb.get().strip() if cb else ""
                if val: res_nacht.append(val)
            except Exception:
                pass
        return res_tag, res_nacht

    
    def export_sonderaufgaben_to_excel(self):
        """
        Exportiert den Tab 'Sonderaufgaben' als Excel im Layout wie in 'Sonderaufgaben.xlsx':
        - Spalte A: Aufgabe
        - Spalte C: Tagdienst-Namen
        - Spalte E: Nachtdienst-Namen
        - Kopfzeile mit Datum in A2, 'Tagdienst' in C2, 'Nachtdienst' in E2
        Hinweis: Der GUI-Hinweistext wird NICHT exportiert.
        """
        # Dateiauswahl: Zieldatei
        from tkinter import filedialog, messagebox
        out_path = filedialog.asksaveasfilename(
            title="Sonderaufgaben als Excel speichern",
            defaultextension=".xlsx",
            filetypes=[("Excel-Datei", "*.xlsx")]
        )
        if not out_path:
            return
        try:
            # Versuche, die bereitgestellte Vorlage zu nutzen, um Layout/Format zu wahren
            from openpyxl import load_workbook, Workbook
            import datetime
            template_path = None
            # Prüfe mögliche Template-Dateien im gleichen Ordner
            for cand in ("Sonderaufgaben.xlsx", "./Sonderaufgaben.xlsx"):
                p = Path(cand)
                if p.exists():
                    template_path = p
                    break

            if template_path:
                wb = load_workbook(template_path)
                ws = wb[wb.sheetnames[0]]
            else:
                # Fallback: Neues Workbook mit minimaler Struktur
                wb = Workbook()
                ws = wb.active
                ws.title = "Sonderaufgaben"
                # Kopf
                ws.cell(row=1, column=3, value="Sonderaufgaben")
                ws.cell(row=2, column=3, value="Tagdienst")
                ws.cell(row=2, column=5, value="Nachtdienst")

            # Datum setzen (A2)
            today = datetime.date.today()
            ws.cell(row=2, column=1, value=today)

            # Mapping: ab Zeile 3 nach Vorlage
            start_row = 3
            for idx, aufgabe in enumerate(self.AUFGABEN, start=0):
                r = start_row + idx
                ws.cell(row=r, column=1, value=aufgabe)

                # Werte aus GUI holen
                # Für Bulmor können 2 Felder pro Schicht existieren (rückwärtskompatibles Mapping)
                try:
                    # Tag
                    vals_tag = []
                    cb_t1 = self.entries_tag.get(aufgabe)
                    cb_t2 = getattr(self, "entries_tag_2", {}).get(aufgabe) if hasattr(self, "entries_tag_2") else None
                    for cb in (cb_t1, cb_t2):
                        if cb:
                            try:
                                v = cb.get().strip()
                                if v: vals_tag.append(v)
                            except Exception:
                                pass
                    tag_str = " / ".join(vals_tag) if vals_tag else None
                    ws.cell(row=r, column=3, value=tag_str)

                    # Nacht
                    vals_nacht = []
                    cb_n1 = self.entries_nacht.get(aufgabe)
                    cb_n2 = getattr(self, "entries_nacht_2", {}).get(aufgabe) if hasattr(self, "entries_nacht_2") else None
                    for cb in (cb_n1, cb_n2):
                        if cb:
                            try:
                                v = cb.get().strip()
                                if v: vals_nacht.append(v)
                            except Exception:
                                pass
                    nacht_str = " / ".join(vals_nacht) if vals_nacht else None
                    ws.cell(row=r, column=5, value=nacht_str)
                except Exception:
                    # Fallback: setze leere Felder
                    ws.cell(row=r, column=3, value=None)
                    ws.cell(row=r, column=5, value=None)

            # Speichern
            wb.save(out_path)
        except Exception as e:
            try:
                messagebox.showerror("Export-Fehler", str(e))
            except Exception:
                print("Export-Fehler:", e)

    def load_names_from_excel(self):
        # openpyxl optional und klar melden, wenn fehlend
        try:
            from openpyxl import load_workbook
            _OPENPY_AVAILABLE = True
        except Exception as _e:
            _OPENPY_AVAILABLE = False
            try:
                from tkinter import messagebox as _mb
                _mb.showerror("Fehlender Baustein", f"openpyxl ist nicht installiert:\n{_e}\n\nInstalliere mit:\n    pip install openpyxl")
            except Exception:
                pass
            return

        path = filedialog.askopenfilename(
            title="Excel mit Namens- und Dienstspalte wählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls")]
        )
        if not path: return

        try:
            wb = load_workbook(path, data_only=True)
            # erstes brauchbares Sheet
            ws = None
            for sn in wb.sheetnames:
                if wb[sn].max_row > 1 and wb[sn].max_column > 1:
                    ws = wb[sn]; break
            if ws is None:
                try:
                    messagebox.showwarning("Hinweis", "Keine verwertbaren Datenblätter in der Datei gefunden.")
                except Exception:
                    pass
                return

            header_row, name_col, dienst_col = self._find_headers_same_row(ws)
            if not (header_row and name_col and dienst_col):
                try:
                    messagebox.showwarning("Hinweis", "Ich habe keine passenden Spalten für Name und Dienst (gleiche Zeile) gefunden.")
                except Exception:
                    pass
                return

            def shift_from_dienst(val: str):
                s = str(val or "").strip().lower()
                if s.startswith(("t","t10","t8")): return "tag"
                if s.startswith(("n","n10","nf")): return "nacht"
                if s[:1] == "t": return "tag"
                if s[:1] == "n": return "nacht"
                return None

            tag_set, nacht_set = set(), set()
            bul_tag, bul_nacht = set(), set()
            for rr in range(header_row+1, ws.max_row+1):
                nval = ws.cell(row=rr, column=name_col).value
                if not nval: continue
                dval = ws.cell(row=rr, column=dienst_col).value
                schicht = shift_from_dienst(dval)
                last = self._extract_lastname_robust(str(nval))

                if schicht == "tag": tag_set.add(last)
                elif schicht == "nacht": nacht_set.add(last)

                fill = ws.cell(row=rr, column=dienst_col).fill
                rgb = self._get_effective_rgb(wb, fill)
                if self._is_yellowish(rgb):
                    if schicht == "tag": bul_tag.add(last)
                    elif schicht == "nacht": bul_nacht.add(last)

            self.namen_liste_tag = sorted(tag_set, key=lambda s: s.lower())
            self.namen_liste_nacht = sorted(nacht_set, key=lambda s: s.lower())
            self.namen_liste_bulmor_tag = sorted(bul_tag, key=lambda s: s.lower()) if bul_tag else self.namen_liste_tag[:]
            self.namen_liste_bulmor_nacht = sorted(bul_nacht, key=lambda s: s.lower()) if bul_nacht else self.namen_liste_nacht[:]
            self.namen_liste = sorted(set(self.namen_liste_tag + self.namen_liste_nacht), key=lambda s: s.lower())

            self.update_name_choices_by_task()
        except Exception as e:
            import traceback, os
            tb = traceback.format_exc()
            log_path = os.path.join(os.path.dirname(__file__), "sonderaufgaben_error.log")
            try:
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(tb)
            except Exception:
                pass
            try:
                messagebox.showerror("Fehler beim Laden", f"{e}\n\nDetails in:\n{log_path}")
            except Exception:
                print(tb)

class StatisticsTab(tk.Frame):
    def __init__(self, master=None, style_vars=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        if style_vars:
            self.pastel_bg = style_vars.get('bg')
            self.pastel_fg = style_vars.get('fg')
        else:
            self.pastel_bg = '#E0F7FA'
            self.pastel_fg = '#37474F'
        self.configure(bg=self.pastel_bg)
        self.create_widgets()
        
    def create_widgets(self):
        container_frame = tk.Frame(self, bg=self.pastel_bg, padx=20, pady=20)
        container_frame.pack(expand=True, fill=tk.BOTH)
        
        title_label = tk.Label(container_frame, text="Tagesübersicht", font=("Segoe UI", 16, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)
        title_label.pack(pady=(0, 20))
        
        self.stats_frame = tk.Frame(container_frame, bg=self.pastel_bg)
        self.stats_frame.pack(fill=tk.X)
        
        self.short_list_frame = tk.Frame(container_frame, bg=self.pastel_bg)
        self.short_list_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        # Bulmor-Status beim Aufbau rendern
        try:
            self.refresh_bulmor()
        except Exception:
            pass
        
        self.update_statistics(None)

    def _parse_time(self, val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        if isinstance(val, (pd.Timestamp, datetime)):
            return datetime(2000, 1, 1, val.hour, val.minute, val.second)
        if isinstance(val, str):
            s = val.strip()
            m = re.match(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$', s)
            if m:
                hh = int(m.group(1)); mm = int(m.group(2)); ss = int(m.group(3) or 0)
                if 0 <= hh <= 23 and 0 <= mm <= 59 and 0 <= ss <= 59:
                    return datetime(2000, 1, 1, hh, mm, ss)
            try:
                ts = pd.to_datetime(s, errors='raise')
                return datetime(2000, 1, 1, ts.hour, ts.minute, ts.second)
            except Exception:
                return None
        return None

    def _shift_by_begin(self, begin_dt):
        if not begin_dt:
            return "unbekannt"
        return "tag" if 6 <= begin_dt.hour <= 17 else "nacht"

    def _duration_hours(self, begin_dt, end_dt):
        if not begin_dt or not end_dt:
            return None
        if end_dt < begin_dt:
            end_dt = end_dt + timedelta(days=1)
        return round((end_dt - begin_dt).total_seconds() / 3600.0, 2)

    def _role_code(self, s):
        s = (s or "").strip().upper()
        m = re.match(r'^[A-Z]+[0-9]*', s)
        return m.group(0) if m else s

    def update_statistics(self, df):
        for w in self.stats_frame.winfo_children(): w.destroy()
        for w in self.short_list_frame.winfo_children(): w.destroy()

        if df is None or df.empty:
            tk.Label(self.stats_frame, text="Bitte laden Sie eine Excel-Datei, um die Statistik anzuzeigen.",
                     font=("Segoe UI", 12), bg=self.pastel_bg, fg=self.pastel_fg).pack(pady=20)
            return

        work = df.copy()
        work['Dienst_clean'] = work['Dienst'].astype(str).str.strip().str.lower()
        work['Role'] = work['Dienst'].apply(self._role_code)

        # Deaktivierte Mitarbeitende ausschließen
        if 'deactivated' in work.columns:
            work = work[~work['deactivated'].astype(bool)]

        work = work[~work['Dienst_clean'].str.startswith(('dt', 'dn'))]
        name_norm = work['Namen'].astype(str).str.strip().str.lower()
        work = work[~name_norm.isin({'lars peters', 'peters, lars'})]

        is_sick_series = work['Dienst_clean'].str.startswith(('k', 'krank'))
        if 'is_sick' in work.columns:
            is_sick_series = is_sick_series | work['is_sick'].astype(bool)

        work['_begin_dt'] = work['Beginn'].apply(self._parse_time)
        work['_end_dt']   = work['Ende'].apply(self._parse_time)
        work['_shift']    = work['_begin_dt'].apply(self._shift_by_begin)
        work['_hours']    = work.apply(lambda r: self._duration_hours(r['_begin_dt'], r['_end_dt']), axis=1)

        kvs_heute = 0
        try:
            plan_date = None
            if 'Datum' in df.columns and len(df) > 0:
                first_date = str(df.iloc[0]['Datum'])
                try:
                    plan_date = datetime.strptime(first_date, "%d.%m.%Y")
                except Exception:
                    plan_date = pd.to_datetime(first_date, errors='coerce')
                    if pd.isna(plan_date): plan_date = None
                    elif isinstance(plan_date, pd.Timestamp): plan_date = plan_date.to_pydatetime()
            if plan_date is None: plan_date = datetime.today()

            app = self.master.master if hasattr(self, 'master') and hasattr(self.master, 'master') else None
            db = getattr(app, 'db', None) if app else None
            if db and db.conn:
                kvs_rows = db.fetch_kvs_for_date(plan_date)
                kvs_names = set()
                for name, _, d_von, d_bis in kvs_rows:
                    if not name: continue
                    ln = extract_lastname(str(name)).strip().lower()
                    if ln in ('lars peters', 'peters, lars'): continue
                    kvs_names.add(ln)
                kvs_heute = len(kvs_names)
        except Exception:
            kvs_heute = 0

        wanted_roles = ['T', 'T10', 'T8', 'N', 'N10', 'NF']
        counts_roles = {role: int((work['Role'] == role).sum()) for role in wanted_roles}
        # Nachnamen je Rolle für Anzeige hinter den Summen
        try:
            role_lastnames = {r: ', '.join(sorted({extract_lastname(n) for n in work.loc[work['Role']==r, 'Namen'].dropna()})) for r in wanted_roles}
        except Exception:
            role_lastnames = {r: '' for r in wanted_roles}

        t_roles = ['t', 't10', 't8']
        n_roles = ['n', 'n10', 'nf']
        count_tag = work[work['Dienst_clean'].str.startswith(tuple(t_roles))]['Dienst'].count()
        count_nacht = work[work['Dienst_clean'].str.startswith(tuple(n_roles))]['Dienst'].count()
        count_krank = int(is_sick_series.sum())
        sick_rows = work[is_sick_series]
        count_krank_tag = int((sick_rows['_shift'] == 'tag').sum())
        count_krank_nacht = int((sick_rows['_shift'] == 'nacht').sum())
        count_total = int(work.shape[0])

        row = 0
        tk.Label(self.stats_frame, text="Rollenübersicht:", font=("Segoe UI", 12, "bold"),
                 bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, columnspan=2, sticky='w', padx=10, pady=(0,5))
        row += 1
        for r in ['T', 'T10', 'T8', 'N', 'N10', 'NF']:
            tk.Label(self.stats_frame, text=f"{r}:", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=3)
            tk.Label(self.stats_frame, text=str(counts_roles[r]), font=("Segoe UI", 12, "bold"),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=1, sticky='e', padx=10, pady=3)
            # Nachnamen rechts daneben
            try:
                names_text = role_lastnames.get(r, '')
            except Exception:
                names_text = ''
            tk.Label(self.stats_frame, text=names_text, font=("Segoe UI", 11),
                     bg=self.pastel_bg, fg=self.pastel_fg, anchor='w').grid(row=row, column=2, sticky='w', padx=10, pady=3)
            row += 1

        row += 1
        for label, value in [
            ("Tagdienste gesamt (T, T10, T8)", count_tag),
            ("Nachtdienste gesamt (N, N10, NF)", count_nacht),
            ("Krankmeldungen gesamt", count_krank),
            ("davon krank in Tagschicht", count_krank_tag),
            ("davon krank in Nachtschicht", count_krank_nacht),
            ("KvS heute", kvs_heute),
            ("Gesamtmitarbeiter im Dienst", count_total),
        ]:
            tk.Label(self.stats_frame, text=f"{label}:", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=0, sticky='w', padx=10, pady=3)
            tk.Label(self.stats_frame, text=str(value), font=("Segoe UI", 12, "bold"),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=row, column=1, sticky='e', padx=10, pady=3)
            row += 1

        threshold = 10.0
        short_df = work[(work['_hours'].notna()) & (work['_hours'] < threshold) & (~is_sick_series)].copy()
        short_df = short_df[['Namen', 'Beginn', 'Ende', '_hours']].sort_values('_hours')

        tk.Label(self.short_list_frame, text=f"Mitarbeitende mit weniger als {threshold:g} Stunden:",
                 font=("Segoe UI", 12, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)\
          .grid(row=0, column=0, columnspan=2, sticky='w', padx=10, pady=(0,8))

        if short_df.empty:
            tk.Label(self.short_list_frame, text="— keine —", font=("Segoe UI", 12),
                     bg=self.pastel_bg, fg=self.pastel_fg).grid(row=1, column=0, sticky='w', padx=20, pady=2)
        else:
            r = 1
            for _, rowv in short_df.iterrows():
                beg_str = format_time(rowv['Beginn'])
                end_str = format_time(rowv['Ende'])
                tk.Label(self.short_list_frame, text=f"{rowv['Namen']}:", font=("Segoe UI", 12),
                         bg=self.pastel_bg, fg=self.pastel_fg).grid(row=r, column=0, sticky='w', padx=20, pady=2)
                tk.Label(self.short_list_frame, text=f"{beg_str} – {end_str} ({rowv['_hours']:.2f} h)",
                         font=("Segoe UI", 12, "bold"), bg=self.pastel_bg, fg=self.pastel_fg)\
                  .grid(row=r, column=1, sticky='w', padx=10, pady=2)
                r += 1

# -------------------------------------------------------------------
# Haupt-App
# -------------------------------------------------------------------
class UnifiedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NeSk")
        self.geometry("1400x900")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.pastel_bg = '#E0F7FA'
        self.pastel_fg = '#37474F'
        self.pastel_button = '#B2DFDB'
        self.configure(bg=self.pastel_bg)

        if not OUTLOOK_AVAILABLE:
            self.after(0, lambda: messagebox.showwarning(
                "Warnung",
                "Die 'pywin32'-Bibliothek wurde nicht gefunden.\n"
                "Die E-Mail-Funktion kann das Dokument nicht automatisch anhängen.\n"
                "Installiere sie mit: pip install pywin32"
            ))
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        style_vars = {'bg': self.pastel_bg, 'fg': self.pastel_fg, 'button': self.pastel_button}
        
        self.db = None

        self.statistics_tab = StatisticsTab(self.notebook, style_vars=style_vars)
        self.notebook.add(self.statistics_tab, text="Tagesübersicht")
        # --- Vordrucke Tab ---
        VORDRUCKE_PATH = r"C:\Vordrucke"
        self.vordrucke_tab = VordruckeTab(self.notebook, VORDRUCKE_PATH)
        self.notebook.add(self.vordrucke_tab, text="Vordrucke")

        # --- Automatisches Aktualisieren beim Tabwechsel ---
        def on_tab_changed(event):
            tab = event.widget.nametowidget(event.widget.select())
            if isinstance(tab, VordruckeTab):
                tab.load_files()
            try:
                # Wenn zur Tagesübersicht gewechselt wird, Bulmor-Status aktualisieren
                if isinstance(tab, StatisticsTab):
                    tab.refresh_bulmor()
            except Exception:
                pass

        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)
        
        self.krankmeldung_tab = KrankmeldungTab(self.notebook, db=self.db, style_vars=style_vars)
        self.notebook.add(self.krankmeldung_tab, text="Krankmeldung")
        
        self.dienstplan_tab = DienstplanTab(self.notebook, bg_color=self.pastel_bg, fg_color=self.pastel_fg)
        self.notebook.add(self.dienstplan_tab, text="Dienstplangenerator")

        # NEU: Sonderaufgaben-Tab
        self.sonderaufgaben_tab = SonderaufgabenTab(self.notebook)
        self.notebook.add(self.sonderaufgaben_tab, text="Sonderaufgaben")
        
        self.excel_viewer_tab = ExcelViewerTab(self.notebook, db=self.db)
        self.notebook.add(self.excel_viewer_tab, text="Tagesdienstplan")
        
        self.connect_button = ttk.Button(self, text="Datenbank auswählen", command=self.select_and_connect_to_db, style='TButton')
        self.connect_button.pack(pady=10)
        
        # Automatische Verbindung beim Start mit festem Pfad
        self.connect_to_db(r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26\python\Database101.accdb")

    def handle_file_load(self, file_path):
        success = self.excel_viewer_tab.load_data(file_path)
        if success and self.excel_viewer_tab.df is not None:
            self.statistics_tab.update_statistics(self.excel_viewer_tab.df)
            self.notebook.select(self.excel_viewer_tab)

    def connect_to_db(self, db_file):
        self.db = AccessDB(db_file)
        if self.db.conn:
            messagebox.showinfo("Verbindung erfolgreich", f"Verbindung zur Datenbank '{os.path.basename(db_file)}' erfolgreich hergestellt.")
            self.krankmeldung_tab.db = self.db
            self.excel_viewer_tab.db = self.db
            self.krankmeldung_tab.display_data_in_tables()
            if self.excel_viewer_tab.df is not None:
                self.excel_viewer_tab.recompute_sickness_from_db()
        else:
            self.db = None
            self.krankmeldung_tab.db = None
            self.excel_viewer_tab.db = None
            self.krankmeldung_tab.display_data_in_tables()

    def select_and_connect_to_db(self):
        db_file = filedialog.askopenfilename(title="Access-Datenbank auswählen", filetypes=[("Access-Datenbank", "*.accdb *.mdb")])
        if db_file:
            self.connect_to_db(db_file)

if __name__ == "__main__":
    try:
        app = UnifiedApp()
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        from tkinter import messagebox
        messagebox.showerror("Fehler beim Start", f"{e}")




# === Bulmor Icon Patch (auto-injected) ===
# Adds a compact Bulmor icon bar on the right side of the Tagesübersicht (StatisticsTab)
# and ensures it refreshes after each update_statistics call.
try:
    import tkinter as _tk  # noqa: F401
except Exception:
    _tk = None

def _bulmor_make_icon(_container, _status, _bg, _fg):
    """Create a tiny 'vehicle' icon Canvas indicating Bulmor status."""
    import tkinter as tk
    c = tk.Canvas(_container, width=32, height=16, highlightthickness=0, bg=_bg)
    color = "#2ecc71" if str(_status).strip().lower() == "fahrbereit" else "#e74c3c"
    # body
    c.create_rectangle(2, 6, 26, 13, fill=color, outline=color)
    # cabin
    c.create_rectangle(20, 3, 28, 10, fill=color, outline=color)
    # wheels
    c.create_oval(5, 11, 9, 15, fill="#333333", outline="#333333")
    c.create_oval(18, 11, 22, 15, fill="#333333", outline="#333333")
    # short handle
    c.create_line(28, 7, 30, 7, fill="#333333")
    return c



def _new_refresh_bulmor(self):
    """
    Create a header bar under the notebook tab 'Tagesübersicht':
    - Left: Title "Tagesübersicht"
    - Right: Bulmor icons (1–5, green=fahrbereit, red=defekt)

    Additionally, keep the textual "Bulmor-Status:" line in the stats section.
    """
    try:
        import tkinter as tk
    except Exception:
        return

    # resolve parent containers
    stats_parent = getattr(self, 'stats_frame', None)
    if stats_parent is None:
        return
    container = getattr(stats_parent, 'master', self)

    bg = getattr(self, "pastel_bg", getattr(self, "bg_color", "#f7f7fb"))
    fg = getattr(self, "pastel_fg", getattr(self, "fg_color", "#222222"))

    # Hide/remove old single title label if present to avoid duplicates
    try:
        # find a direct child of container that is a Label with the Tagesübersicht text
        for w in list(container.children.values()):
            if str(w.winfo_class()) == "Label":
                try:
                    if w.cget("text").strip().lower() == "tagesübersicht":
                        try:
                            w.pack_forget()
                            w.destroy()
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass

    # (Re)build combined title + icons header
    try:
        if not hasattr(self, "_title_header") or not self._title_header.winfo_exists():
            self._title_header = tk.Frame(container, bg=bg)
            # Insert before stats_frame so it sits directly under the tab
            try:
                self._title_header.pack(fill="x", pady=(0, 20), before=self.stats_frame)
            except Exception:
                self._title_header.pack(fill="x", pady=(0, 20))
        else:
            # Ensure position before stats_frame
            try:
                self._title_header.pack_forget()
                self._title_header.pack(fill="x", pady=(0, 20), before=self.stats_frame)
            except Exception:
                pass

        # clear previous content
        for ch in self._title_header.winfo_children():
            ch.destroy()

        left = tk.Frame(self._title_header, bg=bg)
        right = tk.Frame(self._title_header, bg=bg)
        left.pack(side="left", anchor="w")
        right.pack(side="right", anchor="e")

        tk.Label(
            left,
            text="Tagesübersicht",
            font=("Segoe UI", 16, "bold"),
            bg=bg, fg=fg
        ).pack()

        # status map (defaults to 'fahrbereit')
        st = {}
        try:
            st = load_bulmor_status()
            if not isinstance(st, dict):
                st = {}
        except Exception:
            st = {}

        def _icon(container, status):
            c = tk.Canvas(container, width=32, height=16, highlightthickness=0, bg=bg)
            color = "#2ecc71" if str(status).strip().lower() == "fahrbereit" else "#e74c3c"
            c.create_rectangle(2, 6, 26, 13, fill=color, outline=color)
            c.create_rectangle(20, 3, 28, 10, fill=color, outline=color)
            c.create_oval(5, 11, 9, 15, fill="#333333", outline="#333333")
            c.create_oval(18, 11, 22, 15, fill="#333333", outline="#333333")
            c.create_line(28, 7, 30, 7, fill="#333333")
            return c

        # label on the right
        tk.Label(right, text="Bulmor:", font=("Segoe UI", 10, "bold"), bg=bg, fg=fg).pack(side="left", padx=(0,8))

        # icons 1..5
        for i in range(1, 6):
            name = f"Bulmor {i}"
            val = st.get(name, "fahrbereit")
            cell = tk.Frame(right, bg=bg)
            cell.pack(side="left", padx=6)
            _icon(cell, val).pack()
            tk.Label(cell, text=str(i), font=("Segoe UI", 8, "bold"), bg=bg, fg=fg).pack()
    except Exception:
        return

# Patch update_statistics to call refresh_bulmor afterwards
try:
    _orig_update_statistics = StatisticsTab.update_statistics
    def _patched_update_statistics(self, *args, **kwargs):
        res = _orig_update_statistics(self, *args, **kwargs)
        try:
            # render / refresh Bulmor icons after stats update
            self.refresh_bulmor()
        except Exception:
            pass
        return res
    StatisticsTab.update_statistics = _patched_update_statistics
except Exception:
    pass
# === End Bulmor Icon Patch ===


# === Bulmor Init Hook Patch ===
try:
    _orig_stats_init = StatisticsTab.__init__
    def _patched_stats_init(self, *a, **kw):
        _orig_stats_init(self, *a, **kw)
        try:
            # initial render on tab creation
            if hasattr(self, "refresh_bulmor"):
                self.refresh_bulmor()
        except Exception:
            pass
    StatisticsTab.__init__ = _patched_stats_init
except Exception:
    pass
# === End Bulmor Init Hook Patch ===


# === Bulmor Right-Anchor Refinement ===
def _refined_refresh_bulmor(self):
    """
    Place Bulmor bar to the RIGHT of the main stats content:
    pack it into the parent container of self.stats_frame (container_frame).
    This avoids being destroyed when stats_frame is rebuilt.
    """
    try:
        import tkinter as tk
    except Exception:
        return

    bg = getattr(self, "pastel_bg", "#f7f7fb")
    fg = getattr(self, "pastel_fg", "#222222")

    # resolve container: parent of stats_frame
    try:
        parent = self.stats_frame.master if hasattr(self, "stats_frame") else self
    except Exception:
        parent = getattr(self, "master", self)

    # remove any previous bar
    try:
        if hasattr(self, "_bulmor_frame") and getattr(self._bulmor_frame, "winfo_exists", lambda: False)():
            self._bulmor_frame.destroy()
    except Exception:
        pass

    self._bulmor_frame = tk.LabelFrame(parent, text="Bulmor", bg=bg, fg=fg)
    # Pack right so it visibly sits at the right edge of the Tagesübersicht
    try:
        self._bulmor_frame.pack(side="right", anchor="ne", padx=10, pady=(0, 10))
    except Exception:
        try:
            self._bulmor_frame.place(relx=0.98, rely=0.02, anchor="ne")
        except Exception:
            pass

    # get status
    st = {}
    try:
        st = load_bulmor_status()
        if not isinstance(st, dict):
            st = {}
    except Exception:
        st = {}

    row = tk.Frame(self._bulmor_frame, bg=bg)
    row.pack(fill="x", padx=10, pady=6)

    def _mk(_container, _status):
        return _bulmor_make_icon(_container, _status, bg, fg)

    for i in range(1, 6):
        name = f"Bulmor {i}"
        val = st.get(name, "fahrbereit")
        cell = tk.Frame(row, bg=bg)
        cell.pack(side="left", padx=6)
        _mk(cell, val).pack()
        tk.Label(cell, text=str(i), font=("Segoe UI", 9, "bold"), bg=bg, fg=fg).pack()

    # --- Zusatzliste unter den Icons: Status / Termin / Notiz ---
    try:
        info_frame = tk.Frame(self._bulmor_frame, bg=bg)
        info_frame.pack(fill="x", padx=10, pady=(6,6))

        # DB-Objekt finden
        db = None
        try:
            db = getattr(getattr(self, 'master', None), 'db', None) or getattr(getattr(getattr(self, 'master', None), 'master', None), 'db', None)
        except Exception:
            db = None

        rows = []
        if db and getattr(db, 'conn', None):
            try:
                rows = db.fetch_bulmor_entries()  # (ID, Bulmor, Status, WerkstattTermin, Freitext, Erledigt)
            except Exception:
                rows = []

        latest = {}
        for rid, bul, stt, wt, ft, erledigt in rows:
            latest[bul] = {"Status": stt, "Termin": wt, "Freitext": ft}

        for i in range(1, 6):
            name = f"Bulmor {i}"
            d = latest.get(name, {})
            status = d.get("Status", "fahrbereit")
            term = d.get("Termin")
            term_txt = "—" if term in (None, "") else str(term)[:10]
            note = (d.get("Freitext") or "").strip()

            line = f"{name}: {status}   |   Termin: {term_txt}"
            if note:
                line += f"   |   Notiz: {note}"

            tk.Label(info_frame, text=line, bg=bg, fg=fg, anchor="w", justify="left", wraplength=320).pack(fill="x", pady=1)
    except Exception:
        pass


# Rebind method on the class
try:
    StatisticsTab.refresh_bulmor = _refined_refresh_bulmor
except Exception:
    pass
# Ensure update and init hooks still call the new one
try:
    def _ensure_call_after_update(self, *a, **kw):
        ret = _orig_update_statistics(self, *a, **kw)
        try:
            self.refresh_bulmor()
        except Exception:
            pass
        return ret
    StatisticsTab.update_statistics = _ensure_call_after_update
except Exception:
    pass
# === End Bulmor Right-Anchor Refinement ===
