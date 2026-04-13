"""
gui.py – Grafični vmesnik za Nepremičnine.net scraper
======================================================
Tkinter GUI za izbiro regij, vrst nepremičnin in nastavitev.
Za zagon:  python gui.py
"""

import sys
import os
import csv as _csv
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import math
import statistics

# Poti
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRAPER    = os.path.join(SCRIPT_DIR, "scraper.py")
ANALYZE    = os.path.join(SCRIPT_DIR, "analyze.py")
MODELI     = os.path.join(SCRIPT_DIR, "modeli.py")
CENIK      = os.path.join(SCRIPT_DIR, "cenik.py")

# Uvozi konstante iz scraper.py
sys.path.insert(0, SCRIPT_DIR)
try:
    from scraper import REGIJE, VRSTE, AKCIJE
except ImportError:
    REGIJE = {"Gorenjska": "gorenjska"}
    VRSTE  = {"Hiša": "hisa"}
    AKCIJE = ["prodaja", "najem"]

# ── Barve / stil ──────────────────────────────────────────────────────────────
BG        = "#1e1e2e"
BG2       = "#2a2a3e"
BG3       = "#313145"
ACCENT    = "#7c6af7"
ACCENT2   = "#5a4fcf"
GREEN     = "#4caf50"
RED       = "#f44336"
ORANGE    = "#ff9800"
FG        = "#e8e8f0"
FG2       = "#a0a0b8"
FONT      = ("Segoe UI", 10)
FONT_B    = ("Segoe UI", 10, "bold")
FONT_H    = ("Segoe UI", 13, "bold")
MONO      = ("Consolas", 9)


# ── Pomožni drsnični vsebnik ──────────────────────────────────────────────────
class _ScrollFrame(tk.Frame):
    """
    Splošni navpično drsnični vsebnik.
    Vsebino dodajaj v  .body  atribut (tk.Frame).
    Mousewheel deluje, ko je miška znotraj tega widgeta.
    """

    def __init__(self, parent, bg=BG, **kw):
        super().__init__(parent, bg=bg, **kw)

        c  = tk.Canvas(self, bg=bg, highlightthickness=0, borderwidth=0)
        sb = tk.Scrollbar(self, orient="vertical", command=c.yview,
                          bg=BG3, troughcolor=BG2)
        self.body  = tk.Frame(c, bg=bg)
        self._c    = c
        self._cmd  = lambda e: c.yview_scroll(int(-1 * (e.delta / 120)), "units")

        c.configure(yscrollcommand=sb.set)
        self._win = c.create_window((0, 0), window=self.body, anchor="nw")

        sb.pack(side="right", fill="y")
        c.pack(side="left",  fill="both", expand=True)

        # Posodobi scroll-region ko se vsebina spremeni
        self.body.bind("<Configure>",
                       lambda e: c.configure(scrollregion=c.bbox("all")))
        # Razširi notranjo širino ko se canvas spremeni
        c.bind("<Configure>",
               lambda e: c.itemconfig(self._win, width=e.width))

        # Mousewheel: aktivno, ko je kazalec znotraj tega widgeta
        self.bind("<Enter>", lambda e: c.bind_all("<MouseWheel>", self._cmd))
        self.bind("<Leave>", lambda e: c.unbind_all("<MouseWheel>"))

    def refresh_wheel(self):
        """Pokliči po dinamičnem dodajanju vsebine, da mousewheel deluje povsod."""
        def _rec(w):
            w.bind("<MouseWheel>", self._cmd)
            for ch in w.winfo_children():
                _rec(ch)
        _rec(self.body)


# ── CheckList ─────────────────────────────────────────────────────────────────
class CheckList(tk.Frame):
    """Razpisni seznam z check-boxom 'Vse' na vrhu."""

    def __init__(self, parent, title, items: dict[str, str],
                 list_h: int = 150, **kw):
        super().__init__(parent, bg=BG2, **kw)
        self._items = items  # {ime: slug}
        self._vars: dict[str, tk.BooleanVar] = {}

        # Naslov
        tk.Label(self, text=title, bg=BG2, fg=FG, font=FONT_B,
                 anchor="w").pack(fill="x", padx=8, pady=(8, 2))

        # "Vse" checkbox
        self._all_var = tk.BooleanVar(value=True)
        cb_all = tk.Checkbutton(self, text="✔ Vse", variable=self._all_var,
                                bg=BG2, fg=ACCENT, selectcolor=BG3,
                                activebackground=BG2, activeforeground=ACCENT,
                                font=FONT_B, anchor="w",
                                command=self._toggle_all)
        cb_all.pack(fill="x", padx=8, pady=2)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=8, pady=4)

        # Scrollable container
        canvas = tk.Canvas(self, bg=BG2, highlightthickness=0, height=list_h)
        sb = tk.Scrollbar(self, orient="vertical", command=canvas.yview,
                          bg=BG3, troughcolor=BG2)
        self._inner = tk.Frame(canvas, bg=BG2)
        self._inner.bind("<Configure>",
                         lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # Popravek: razširi notranje okno na celotno širino canvasa
        _win_id = canvas.create_window((0, 0), window=self._inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_win_id, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0))
        sb.pack(side="right", fill="y")

        # Mousewheel: ko je miška nad canvasom, pomikamo ta seznam
        _chk_scroll = lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _chk_scroll))
        canvas.bind("<Leave>", lambda e: self._restore_parent_scroll(canvas))
        self._inner.bind("<MouseWheel>", _chk_scroll)

        # Checkbox za vsako postavko
        for name in items:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(self._inner, text=name, variable=var,
                                bg=BG2, fg=FG, selectcolor=BG3,
                                activebackground=BG2, activeforeground=FG,
                                font=FONT, anchor="w",
                                command=self._on_item_change)
            cb.pack(fill="x", pady=1)
            cb.bind("<MouseWheel>", _chk_scroll)  # neposredno pomikanje
            self._vars[name] = var

    def _toggle_all(self):
        state = self._all_var.get()
        for var in self._vars.values():
            var.set(state)

    def _on_item_change(self):
        all_set = all(v.get() for v in self._vars.values())
        self._all_var.set(all_set)

    def _restore_parent_scroll(self, canvas):
        """Ko miška zapusti canvas, reaktivira scroll starševskega _ScrollFrame."""
        w = self.master
        while w is not None:
            if isinstance(w, _ScrollFrame):
                canvas.bind_all("<MouseWheel>", w._cmd)
                return
            w = getattr(w, "master", None)
        canvas.unbind_all("<MouseWheel>")

    def selected_slugs(self) -> list[str]:
        """Vrne seznam URL slugov izbranih postavk."""
        return [self._items[name] for name, var in self._vars.items() if var.get()]

    def selected_names(self) -> list[str]:
        return [name for name, var in self._vars.items() if var.get()]


# ── Dialog za nastavitve analize ──────────────────────────────────────────────
class AnalysisDialog(tk.Toplevel):
    """Modalni dialog za izbiro vrst, primerjave in tipov grafov."""

    # (key, kategorija, opis)
    _GRAPH_DEFS = [
        ("hist_cen",    "Histogrami – porazdelitev",                "Histogram cen  (€)"),
        ("hist_cm2",    "Histogrami – porazdelitev",                "Histogram cena/m²  (€/m²)"),
        ("scatter_m2",  "Razsevni diagrami – primerjaj z €",        "Površina (m²)  ↔  Cena (€)"),
        ("scatter_leto","Razsevni diagrami – primerjaj z €",        "Leto gradnje  ↔  Cena (€)"),
        ("bar_lok",     "Stolpičarji – primerjaj z €",              "Mediana cen po lokacijah (top 12)"),
    ]

    def __init__(self, parent: "ScraperGUI"):
        super().__init__(parent)
        self.title("📊  Nastavitve analize")
        self.configure(bg=BG)
        self.geometry("580x700")
        self.minsize(480, 520)
        self.resizable(True, True)
        self.grab_set()

        self._parent = parent
        self._csv_var = tk.StringVar(value=parent._csv_var.get())
        self._vrste_vars: dict[str, tk.BooleanVar] = {}
        self._all_vrste_var = tk.BooleanVar(value=True)
        self._graph_vars = {k: tk.BooleanVar(value=True)
                            for k, _, _ in self._GRAPH_DEFS}
        self._docx_var = tk.BooleanVar(value=True)
        default_docx = os.path.join(SCRIPT_DIR,
                                    f"analiza_{datetime.now().strftime('%Y%m%d_%H%M')}.docx")
        self._docx_path_var = tk.StringVar(value=default_docx)

        self._build_ui()

        # Provide a default CSV and load types
        csv = self._csv_var.get().strip() or self._find_default_csv()
        if csv:
            self._csv_var.set(csv)
            self._load_vrste(csv)

    # ── helpers ───────────────────────────────────────────────────────────────
    def _find_default_csv(self) -> str:
        # Najprej preveri akumulirano bazo iz scraping_runs/
        baza = os.path.join(SCRIPT_DIR, "scraping_runs", "baza.csv")
        if os.path.isfile(baza):
            return baza
        for name in ("nepremicnine_export_prodaja.csv", "nepremicnine_export_najem.csv",
                     "nepremicnine_export.csv"):
            p = os.path.join(SCRIPT_DIR, name)
            if os.path.isfile(p):
                return p
        return ""

    # ── UI gradnja ────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header (zgoraj, fiksno) ────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT2, padx=12, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊  Nastavitve analize",
                 bg=ACCENT2, fg="white", font=FONT_H).pack(side="left")

        # ── Gumbi (spodaj, fiksno – pakirani PRED telesom!) ───────────────
        btn_frame = tk.Frame(self, bg=BG, pady=8)
        btn_frame.pack(side="bottom", fill="x", padx=12)
        tk.Button(btn_frame, text="✖  Prekliči", bg=BG3, fg=FG2, font=FONT,
                  relief="flat", padx=12, pady=6, cursor="hand2",
                  command=self.destroy).pack(side="right", padx=(4, 0))
        tk.Button(btn_frame, text="📊  Zageni analizo", bg=ORANGE, fg="white",
                  font=FONT_B, relief="flat", padx=16, pady=6, cursor="hand2",
                  command=self._start_analysis).pack(side="right", padx=4)

        # ── Drsnično telo ─────────────────────────────────────────────────
        self._sf = _ScrollFrame(self, bg=BG)
        self._sf.pack(fill="both", expand=True, padx=12, pady=8)
        body = self._sf.body

        # ── CSV datoteka ───────────────────────────────────────────────────
        csv_lf = tk.LabelFrame(body, text=" CSV datoteka ",
                               bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        csv_lf.pack(fill="x", pady=(0, 8))
        csv_row = tk.Frame(csv_lf, bg=BG)
        csv_row.pack(fill="x", padx=6, pady=4)
        tk.Entry(csv_row, textvariable=self._csv_var, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT).pack(
            side="left", fill="x", expand=True)
        tk.Button(csv_row, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_csv).pack(side="left", padx=2)
        tk.Button(csv_row, text="↺", bg=ACCENT2, fg="white", relief="flat",
                  cursor="hand2", font=FONT_B,
                  command=lambda: self._load_vrste(self._csv_var.get())).pack(
            side="left", padx=(2, 0))

        # ── Filter po vrsti objekta ────────────────────────────────────────
        vrste_lf = tk.LabelFrame(body, text=" Filtriraj po vrsti objekta (iz CSV) ",
                                 bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        vrste_lf.pack(fill="x", pady=(0, 8))
        self._vrste_container = tk.Frame(vrste_lf, bg=BG)
        self._vrste_container.pack(fill="x", padx=6, pady=4)
        tk.Label(self._vrste_container,
                 text="Izberi CSV in klikni ↺ za prikaz vrst …",
                 bg=BG, fg=FG2, font=("Segoe UI", 9, "italic")).pack(anchor="w")

        # ── Grafi ─────────────────────────────────────────────────────────
        grafi_lf = tk.LabelFrame(body, text=" Grafi za generiranje ",
                                 bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        grafi_lf.pack(fill="x", pady=(0, 8))

        current_cat = None
        for key, cat, label in self._GRAPH_DEFS:
            if cat != current_cat:
                current_cat = cat
                sep_row = tk.Frame(grafi_lf, bg=BG)
                sep_row.pack(fill="x", padx=6, pady=(8, 0))
                tk.Label(sep_row, text=f"  {cat}",
                         bg=BG, fg=FG2,
                         font=("Segoe UI", 9, "bold")).pack(side="left")
                ttk.Separator(sep_row, orient="horizontal").pack(
                    side="left", fill="x", expand=True, padx=(6, 0))
            tk.Checkbutton(grafi_lf, text=f"    {label}",
                           variable=self._graph_vars[key],
                           bg=BG, fg=FG, selectcolor=BG3,
                           activebackground=BG, font=FONT,
                           anchor="w").pack(fill="x", padx=16, pady=1)

        # ── Izvoz DOCX ────────────────────────────────────────────────────
        docx_lf = tk.LabelFrame(body, text=" Izvoz v Word dokument (.docx) ",
                                bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        docx_lf.pack(fill="x", pady=(0, 8))
        docx_top = tk.Frame(docx_lf, bg=BG)
        docx_top.pack(fill="x", padx=6, pady=(4, 2))
        tk.Checkbutton(docx_top, text="Ustvari Word dokument z vsemi rezultati in grafi",
                       variable=self._docx_var, bg=BG, fg=FG,
                       selectcolor=BG3, activebackground=BG, font=FONT_B,
                       anchor="w").pack(side="left")
        docx_path_row = tk.Frame(docx_lf, bg=BG)
        docx_path_row.pack(fill="x", padx=6, pady=(0, 4))
        tk.Label(docx_path_row, text="Shrani kot:", bg=BG, fg=FG2,
                 font=FONT, width=9, anchor="w").pack(side="left")
        tk.Entry(docx_path_row, textvariable=self._docx_path_var, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT).pack(
            side="left", fill="x", expand=True)
        tk.Button(docx_path_row, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_docx).pack(side="left", padx=2)


    def _browse_docx(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word dokument", "*.docx"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR, title="Shrani Word dokument kot…")
        if path:
            self._docx_path_var.set(path)

    # ── CSV brskanje & nalaganje vrst ─────────────────────────────────────────
    def _browse_csv(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV datoteke", "*.csv"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR, title="Izberi CSV datoteko…")
        if path:
            self._csv_var.set(path)
            self._load_vrste(path)

    def _load_vrste(self, path: str):
        """Preberi unikatne VrstaObjekta vrednosti iz CSV in obnovi checkboxe."""
        if not path or not os.path.isfile(path):
            return
        try:
            vrste_set: set[str] = set()
            with open(path, newline="", encoding="utf-8-sig") as f:
                for row in _csv.DictReader(f, delimiter=";"):
                    v = row.get("VrstaObjekta", "").strip()
                    if v:
                        vrste_set.add(v)

            for w in self._vrste_container.winfo_children():
                w.destroy()
            self._vrste_vars.clear()

            if not vrste_set:
                tk.Label(self._vrste_container,
                         text="Ni najdenih vrst objekta v CSV.",
                         bg=BG, fg=FG2, font=FONT).pack(anchor="w")
                return

            self._all_vrste_var.set(True)

            def toggle_all():
                s = self._all_vrste_var.get()
                for v in self._vrste_vars.values():
                    v.set(s)

            def on_item_change():
                self._all_vrste_var.set(
                    all(v.get() for v in self._vrste_vars.values()))

            tk.Checkbutton(self._vrste_container, text="✔ Vse vrste",
                           variable=self._all_vrste_var,
                           bg=BG, fg=ACCENT, selectcolor=BG3,
                           activebackground=BG, font=FONT_B, anchor="w",
                           command=toggle_all).pack(fill="x")
            ttk.Separator(self._vrste_container,
                          orient="horizontal").pack(fill="x", pady=3)

            grid = tk.Frame(self._vrste_container, bg=BG)
            grid.pack(fill="x")
            cols = 3
            for i, vrsta in enumerate(sorted(vrste_set)):
                var = tk.BooleanVar(value=True)
                tk.Checkbutton(grid, text=vrsta, variable=var,
                               bg=BG, fg=FG, selectcolor=BG3,
                               activebackground=BG, font=FONT, anchor="w",
                               command=on_item_change).grid(
                    row=i // cols, column=i % cols, sticky="w", padx=8, pady=1)
                self._vrste_vars[vrsta] = var

        except Exception as exc:
            for w in self._vrste_container.winfo_children():
                w.destroy()
            tk.Label(self._vrste_container, text=f"Napaka pri branju CSV: {exc}",
                     bg=BG, fg=RED, font=FONT).pack(anchor="w")
        finally:
            # Obnovi mousewheel binding za dinamično dodano vsebino
            if hasattr(self, "_sf"):
                self._sf.refresh_wheel()

    # ── Zagon analize ─────────────────────────────────────────────────────────
    def _start_analysis(self):
        if self._parent._running:
            messagebox.showwarning("V teku",
                                   "Počakaj, da se trenutni proces konča.")
            return

        selected_grafi = [k for k, _ in
                          ((k, v) for k, v in self._graph_vars.items() if v.get())]
        if not selected_grafi:
            messagebox.showwarning("Izbor grafov",
                                   "Izberi vsaj en graf za generiranje!")
            return

        cmd = [sys.executable, "-X", "utf8", ANALYZE]
        csv_path = self._csv_var.get().strip()
        if csv_path:
            cmd += ["--csv", csv_path]
        cmd += ["--grafi", ",".join(selected_grafi)]

        if self._vrste_vars:
            sel_vrste = [v for v, var in self._vrste_vars.items() if var.get()]
            if sel_vrste and len(sel_vrste) < len(self._vrste_vars):
                cmd += ["--vrste", ",".join(sel_vrste)]

        if self._docx_var.get():
            cmd.append("--docx")
            docx_p = self._docx_path_var.get().strip()
            if docx_p:
                cmd += ["--docx-izhod", docx_p]

        self.destroy()
        self._parent._log_write(
            f"\n{'=' * 55}\n"
            f"  📊 Analiza: {', '.join(selected_grafi)}\n"
            f"{'=' * 55}\n\n", "head")
        self._parent._set_running(True)
        threading.Thread(target=self._parent._run_process,
                         args=(cmd,), daemon=True).start()


# ── ML Dialog ────────────────────────────────────────────────────────────────
class MLDialog(tk.Toplevel):
    """Dialog za zagon regresijskih ML modelov."""

    def __init__(self, parent: "ScraperGUI"):
        super().__init__(parent)
        self.title("🤖  Regresijski modeli")
        self.configure(bg=BG)
        self.geometry("520x560")
        self.minsize(440, 460)
        self.resizable(True, True)
        self.grab_set()

        self._parent = parent
        self._csv_var   = tk.StringVar(value=parent._csv_var.get())
        self._split_var = tk.DoubleVar(value=80.0)   # 80 % učna množica
        self._seed_var  = tk.IntVar(value=42)
        self._docx_var  = tk.BooleanVar(value=True)
        default_docx = os.path.join(SCRIPT_DIR,
                                    f"ml_porocilo_{datetime.now().strftime('%Y%m%d_%H%M')}.docx")
        self._docx_path_var = tk.StringVar(value=default_docx)

        self._build_ui()
        csv = self._csv_var.get().strip() or self._find_default_csv()
        if csv:
            self._csv_var.set(csv)

    # ── helpers ───────────────────────────────────────────────────────────────
    def _find_default_csv(self) -> str:
        # Najprej preveri akumulirano bazo iz scraping_runs/
        baza = os.path.join(SCRIPT_DIR, "scraping_runs", "baza.csv")
        if os.path.isfile(baza):
            return baza
        for name in ("nepremicnine_export_prodaja.csv", "nepremicnine_export_najem.csv",
                     "nepremicnine_export.csv"):
            p = os.path.join(SCRIPT_DIR, name)
            if os.path.isfile(p):
                return p
        return ""

    def _build_ui(self):
        # ── Header (zgoraj, fiksno) ────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT2, padx=12, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🤖  Regresijski modeli – napovedovanje cen",
                 bg=ACCENT2, fg="white", font=FONT_H).pack(side="left")

        # ── Gumbi (spodaj, fiksno – PRED telesom!) ────────────────────────
        bf = tk.Frame(self, bg=BG, pady=8)
        bf.pack(side="bottom", fill="x", padx=12)
        tk.Button(bf, text="✖  Prekliči", bg=BG3, fg=FG2, font=FONT,
                  relief="flat", padx=12, pady=6, cursor="hand2",
                  command=self.destroy).pack(side="right", padx=(4, 0))
        tk.Button(bf, text="🤖  Zageni modele", bg="#9c27b0", fg="white",
                  font=FONT_B, relief="flat", padx=16, pady=6, cursor="hand2",
                  command=self._start).pack(side="right", padx=4)

        # ── Drsnično telo ─────────────────────────────────────────────────
        sf = _ScrollFrame(self, bg=BG)
        sf.pack(fill="both", expand=True, padx=12, pady=8)
        body = sf.body

        # CSV
        csv_lf = tk.LabelFrame(body, text=" CSV datoteka ", bg=BG, fg=FG2,
                               font=FONT, bd=1, relief="groove")
        csv_lf.pack(fill="x", pady=(0, 8))
        cr = tk.Frame(csv_lf, bg=BG)
        cr.pack(fill="x", padx=6, pady=4)
        tk.Entry(cr, textvariable=self._csv_var, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT).pack(
            side="left", fill="x", expand=True)
        tk.Button(cr, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_csv).pack(side="left", padx=2)

        # Nastavitve
        set_lf = tk.LabelFrame(body, text=" Nastavitve ", bg=BG, fg=FG2,
                               font=FONT, bd=1, relief="groove")
        set_lf.pack(fill="x", pady=(0, 8))

        def srow(lbl, w):
            r = tk.Frame(set_lf, bg=BG)
            r.pack(fill="x", padx=6, pady=3)
            tk.Label(r, text=lbl, bg=BG, fg=FG2, font=FONT,
                     width=20, anchor="w").pack(side="left")
            w(r).pack(side="left")

        srow("Učna množica (%):",
             lambda r: tk.Spinbox(r, from_=50, to=95, increment=5,
                                  textvariable=self._split_var, width=6,
                                  bg=BG3, fg=FG, insertbackground=FG,
                                  buttonbackground=BG3, relief="flat",
                                  highlightthickness=0, font=FONT))
        srow("Naključno seme:",
             lambda r: tk.Spinbox(r, from_=0, to=9999,
                                  textvariable=self._seed_var, width=8,
                                  bg=BG3, fg=FG, insertbackground=FG,
                                  buttonbackground=BG3, relief="flat",
                                  highlightthickness=0, font=FONT))

        # Modeli info
        info_lf = tk.LabelFrame(body, text=" Modeli (vsi se zaženejo) ",
                                bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        info_lf.pack(fill="x", pady=(0, 8))
        for nm, hp in [
            ("Linearna regresija (OLS)", "brez HP"),
            ("Ridge regresija", "α ∈ {0.01, 0.1, 1, 10, 100, 500, 1000}"),
            ("Odločitveno drevo (CART)", "max_depth ∈ {2, 3, 4, 5, 6, 8}"),
            ("Naključni gozd", "n_trees ∈ {10,20}, depth ∈ {3,5,7}"),
        ]:
            row_f = tk.Frame(info_lf, bg=BG)
            row_f.pack(fill="x", padx=8, pady=1)
            tk.Label(row_f, text=f"✦ {nm}", bg=BG, fg=FG,
                     font=FONT, width=32, anchor="w").pack(side="left")
            tk.Label(row_f, text=hp, bg=BG, fg=FG2,
                     font=("Segoe UI", 8, "italic")).pack(side="left")

        # DOCX
        docx_lf = tk.LabelFrame(body, text=" DOCX poročilo ",
                                bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        docx_lf.pack(fill="x", pady=(0, 8))
        dtop = tk.Frame(docx_lf, bg=BG)
        dtop.pack(fill="x", padx=6, pady=(4, 2))
        tk.Checkbutton(dtop, text="Ustvari Word poročilo (modeli, grafi, primerjava)",
                       variable=self._docx_var, bg=BG, fg=FG,
                       selectcolor=BG3, activebackground=BG,
                       font=FONT_B, anchor="w").pack(side="left")
        drow = tk.Frame(docx_lf, bg=BG)
        drow.pack(fill="x", padx=6, pady=(0, 4))
        tk.Label(drow, text="Shrani kot:", bg=BG, fg=FG2,
                 font=FONT, width=9, anchor="w").pack(side="left")
        tk.Entry(drow, textvariable=self._docx_path_var, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT).pack(
            side="left", fill="x", expand=True)
        tk.Button(drow, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_docx).pack(side="left", padx=2)


    def _browse_csv(self):
        p = filedialog.askopenfilename(
            filetypes=[("CSV datoteke", "*.csv"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR, title="Izberi CSV datoteko…")
        if p: self._csv_var.set(p)

    def _browse_docx(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word dokument", "*.docx"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR, title="Shrani poročilo kot…")
        if p: self._docx_path_var.set(p)

    def _start(self):
        if self._parent._running:
            messagebox.showwarning("V teku", "Počakaj, da se trenutni proces konča.")
            return
        csv_p = self._csv_var.get().strip()
        cmd = [sys.executable, "-X", "utf8", MODELI,
               "--split", str(self._split_var.get() / 100),
               "--seed",  str(self._seed_var.get())]
        if csv_p:
            cmd += ["--csv", csv_p]
        if self._docx_var.get():
            cmd.append("--docx")
            dp = self._docx_path_var.get().strip()
            if dp:
                cmd += ["--docx-izhod", dp]
        self.destroy()
        self._parent._log_write(
            f"\n{'='*55}\n  🤖 ML Modeli: linearna, ridge, drevo, naključni gozd\n{'='*55}\n\n",
            "head")
        self._parent._set_running(True)
        threading.Thread(target=self._parent._run_process,
                         args=(cmd,), daemon=True).start()


# ── Dialog za napoved cene ───────────────────────────────────────────────────

def _model_status(csv_path: str, n_trees: int = 60, depth: int = 7) -> str:
    """Vrne besedilo o stanju predpomnjenega modela (npr. 'Shranjen pred 2 min')."""
    if not csv_path or not os.path.isfile(csv_path):
        return "model: CSV ni izbran"
    base       = os.path.splitext(os.path.abspath(csv_path))[0]
    cache_path = f"{base}_cenik_rf{n_trees}_d{depth}.pkl"
    if not os.path.isfile(cache_path):
        return "⚠  Model ni shranjen – prva napoved bo počasna"
    age_s = int(os.path.getmtime(cache_path))
    csv_m = int(os.path.getmtime(csv_path))
    if csv_m > age_s:
        return "⚠  CSV novejši od modela – model bo ponovljen"
    diff  = int(__import__("time").time()) - age_s
    if diff < 120:
        return f"⚡  Model shranjen pred {diff} s"
    elif diff < 3600:
        return f"⚡  Model shranjen pred {diff//60} min"
    elif diff < 86400:
        return f"⚡  Model shranjen pred {diff//3600} h"
    else:
        return f"⚡  Model shranjen pred {diff//86400} dni"


class CenikDialog(tk.Toplevel):
    """Interaktivni dialog za napovedovanje cene nepremičnine."""

    _ENERGIJA = ["A+", "A", "B", "C", "D", "E", "F", "G"]

    def __init__(self, parent: "ScraperGUI"):
        super().__init__(parent)
        self.title("🏠  Napovednik cen nepremičnin")
        self.configure(bg=BG)
        self.geometry("640x760")
        self.minsize(520, 620)
        self.resizable(True, True)

        self._parent = parent
        self._csv_var = tk.StringVar(value=parent._csv_var.get())
        self._process: subprocess.Popen | None = None
        self._retrain_running = False

        # Vhodna polja
        self._vrsta_var    = tk.StringVar()
        self._kraj_var     = tk.StringVar()
        self._povrsina_var = tk.DoubleVar(value=150.0)
        self._zem_var      = tk.DoubleVar(value=500.0)
        self._leto_var     = tk.IntVar(value=1990)
        self._sobe_var     = tk.DoubleVar(value=4.0)
        self._enr_var      = tk.StringVar(value="D")

        self._vrste: list[str] = []
        self._kraji: list[str] = []

        self._build_ui()

        csv = self._csv_var.get().strip() or self._find_default_csv()
        if csv:
            self._csv_var.set(csv)
            self._reload_csv(csv)
        self._update_model_status()

    # ── helpers ───────────────────────────────────────────────────────────────
    def _find_default_csv(self) -> str:
        # Najprej preveri akumulirano bazo iz scraping_runs/
        baza = os.path.join(SCRIPT_DIR, "scraping_runs", "baza.csv")
        if os.path.isfile(baza):
            return baza
        for name in ("nepremicnine_export_prodaja.csv", "nepremicnine_export_najem.csv",
                     "nepremicnine_export.csv"):
            p = os.path.join(SCRIPT_DIR, name)
            if os.path.isfile(p):
                return p
        return ""

    def _reload_csv(self, path: str):
        """Naloži edinstvene vrednosti za spustne menije iz CSV."""
        if not path or not os.path.isfile(path):
            return
        try:
            vrste_set: set[str] = set()
            kraji_set: set[str] = set()
            m2_vals, zem_vals, leto_vals, sob_vals = [], [], [], []
            with open(path, newline="", encoding="utf-8-sig") as fh:
                for r in _csv.DictReader(fh, delimiter=";"):
                    v = r.get("VrstaObjekta", "").strip()
                    if v: vrste_set.add(v)
                    k = (r.get("Obcina") or r.get("Lokacija", "")).strip()
                    if k: kraji_set.add(k)
                    for src, lst in [("VelikostM2", m2_vals),
                                     ("ZemljisteM2", zem_vals),
                                     ("LetoGradnje", leto_vals),
                                     ("StSob", sob_vals)]:
                        try:
                            val = float(str(r.get(src, "")).replace(",", "."))
                            if math.isfinite(val): lst.append(val)
                        except Exception:
                            pass
            self._vrste = sorted(vrste_set)
            self._kraji = sorted(kraji_set)
            self._vrsta_cb["values"] = self._vrste
            self._kraj_cb["values"]  = self._kraji
            if self._vrste: self._vrsta_var.set(self._vrste[0])
            if self._kraji: self._kraj_var.set(self._kraji[0])
            # Nastavi mediane kot privzete vrednosti
            def med(lst): return statistics.median(lst) if lst else 0.0
            self._povrsina_var.set(round(med(m2_vals), 0) if m2_vals else 150)
            self._zem_var.set(round(med(zem_vals), 0) if zem_vals else 500)
            if leto_vals:
                self._leto_var.set(int(med([v for v in leto_vals if v > 1900])))
            if sob_vals:
                self._sobe_var.set(round(med(sob_vals), 1))
        except Exception as exc:
            self._set_result(f"Napaka pri branju CSV:\n{exc}", error=True)

    # ── UI gradnja ────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header (zgoraj, fiksno) ────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT2, padx=12, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏠  Napovednik cen nepremičnin",
                 bg=ACCENT2, fg="white", font=FONT_H).pack(side="left")

        # ── Status bar (spodaj, fiksno – PRED telesom!) ────────────────────
        self._status_lbl = tk.Label(self, text="",
                                    bg=BG3, fg=FG2, font=("Segoe UI", 9),
                                    anchor="w")
        self._status_lbl.pack(side="bottom", fill="x", padx=12, pady=(0, 6))

        # ── Drsnično telo ─────────────────────────────────────────────────
        sf = _ScrollFrame(self, bg=BG)
        sf.pack(fill="both", expand=True, padx=12, pady=8)
        body = sf.body

        # ── CSV ───────────────────────────────────────────────────────────
        csv_lf = tk.LabelFrame(body, text=" CSV datoteka (učna baza) ",
                               bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        csv_lf.pack(fill="x", pady=(0, 8))
        cr = tk.Frame(csv_lf, bg=BG)
        cr.pack(fill="x", padx=6, pady=(4, 2))
        tk.Entry(cr, textvariable=self._csv_var, bg=BG3, fg=FG,
                 insertbackground=FG, relief="flat", font=FONT).pack(
            side="left", fill="x", expand=True)
        tk.Button(cr, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_csv).pack(side="left", padx=2)
        tk.Button(cr, text="↺", bg=ACCENT2, fg="white", relief="flat",
                  cursor="hand2", font=FONT_B,
                  command=lambda: self._reload_csv(self._csv_var.get())).pack(
            side="left", padx=(2, 0))

        # Vrstica: status modela + gumb Prenovi
        model_row = tk.Frame(csv_lf, bg=BG)
        model_row.pack(fill="x", padx=6, pady=(0, 4))
        self._model_status_lbl = tk.Label(
            model_row, text="", bg=BG, fg=FG2,
            font=("Segoe UI", 8, "italic"), anchor="w")
        self._model_status_lbl.pack(side="left", fill="x", expand=True)
        self._retrain_btn = tk.Button(
            model_row, text="🔄  Prenovi model",
            bg="#37474f", fg=FG, font=("Segoe UI", 8), relief="flat",
            padx=8, pady=2, cursor="hand2",
            command=self._prenovi_model)
        self._retrain_btn.pack(side="right")

        # ── Vhodni podatki ────────────────────────────────────────────────
        inp_lf = tk.LabelFrame(body, text=" Podatki o nepremičnini ",
                               bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        inp_lf.pack(fill="x", pady=(0, 8))
        inp_lf.columnconfigure(1, weight=1)
        inp_lf.columnconfigure(3, weight=1)

        def lbl(r, c, txt):
            tk.Label(inp_lf, text=txt, bg=BG, fg=FG2, font=FONT,
                     anchor="e", width=16).grid(row=r, column=c,
                     sticky="e", padx=(8, 4), pady=4)

        def entry(r, c, var, w=10, from_=0, to=9999, inc=1):
            sp = tk.Spinbox(inp_lf, from_=from_, to=to, increment=inc,
                            textvariable=var, width=w,
                            bg=BG3, fg=FG, insertbackground=FG,

                            buttonbackground=BG3, relief="flat",
                            highlightthickness=0, font=FONT)
            sp.grid(row=r, column=c, sticky="w", padx=(0, 8), pady=4)
            return sp

        def combo(r, c, var, values, w=18):
            cb = ttk.Combobox(inp_lf, textvariable=var, values=values,
                              width=w, font=FONT, state="readonly")
            cb.grid(row=r, column=c, sticky="w", padx=(0, 8), pady=4)
            return cb

        lbl(0, 0, "Vrsta objekta:")
        self._vrsta_cb = combo(0, 1, self._vrsta_var, self._vrste)
        lbl(0, 2, "Občina / kraj:")
        self._kraj_cb  = combo(0, 3, self._kraj_var,  self._kraji)

        lbl(1, 0, "Površina (m²):")
        entry(1, 1, self._povrsina_var, from_=10, to=5000, inc=10)
        lbl(1, 2, "Zemljišče (m²):")
        entry(1, 3, self._zem_var, from_=0, to=50000, inc=50)

        lbl(2, 0, "Leto gradnje:")
        entry(2, 1, self._leto_var, from_=1900, to=2026, inc=1)
        lbl(2, 2, "Število sob:")
        entry(2, 3, self._sobe_var, from_=0, to=20, inc=0.5)

        lbl(3, 0, "Energetski razred:")
        combo(3, 1, self._enr_var, self._ENERGIJA, w=6)
        tk.Label(inp_lf, text="(A+ = najboljši, G = najslabši)",
                 bg=BG, fg=FG2, font=("Segoe UI", 8, "italic"),
                 anchor="w").grid(row=3, column=2, columnspan=2, sticky="w",
                                  padx=(0, 8), pady=4)

        # ── Gumb Napovej ──────────────────────────────────────────────────
        tk.Button(body, text="🔍  Napovej ceno",
                  bg=GREEN, fg="white", font=("Segoe UI", 12, "bold"),
                  relief="flat", pady=10, cursor="hand2",
                  command=self._napovej).pack(fill="x", pady=(0, 8))

        # ── Rezultat ──────────────────────────────────────────────────────
        res_lf = tk.LabelFrame(body, text=" Rezultat ",
                               bg=BG, fg=FG2, font=FONT, bd=1, relief="groove")
        res_lf.pack(fill="x", pady=(0, 4))

        price_frame = tk.Frame(res_lf, bg=BG2, pady=10)
        price_frame.pack(fill="x", padx=8, pady=(8, 4))

        self._price_lbl = tk.Label(price_frame,
                                   text="—  vstavi podatke in klikni Napovej  —",
                                   bg=BG2, fg=FG2,
                                   font=("Segoe UI", 18, "bold"))
        self._price_lbl.pack()

        self._ci_lbl = tk.Label(price_frame, text="",
                                bg=BG2, fg=FG2, font=("Segoe UI", 10))
        self._ci_lbl.pack()

        self._detail_txt = tk.Text(res_lf, bg="#12121e", fg="#d4d4f0",
                                   font=MONO, relief="flat", wrap="word",
                                   height=7, state="disabled")
        sb_txt = tk.Scrollbar(res_lf, orient="vertical",
                              command=self._detail_txt.yview,
                              bg=BG3, troughcolor=BG2)
        self._detail_txt.configure(yscrollcommand=sb_txt.set)
        sb_txt.pack(side="right", fill="y", padx=(0, 4), pady=(4, 8))
        self._detail_txt.pack(fill="both", expand=True, padx=(8, 0), pady=(4, 8))
        self._detail_txt.tag_configure("grn", foreground="#4caf50")
        self._detail_txt.tag_configure("org", foreground="#ff9800")
        self._detail_txt.tag_configure("red", foreground="#f44336")


    # ── CSV brskanje ──────────────────────────────────────────────────────────
    def _browse_csv(self):
        p = filedialog.askopenfilename(
            filetypes=[("CSV datoteke", "*.csv"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR, title="Izberi CSV datoteko…")
        if p:
            self._csv_var.set(p)
            self._reload_csv(p)
            self._update_model_status()

    # ── Status modela ─────────────────────────────────────────────────────────
    def _update_model_status(self):
        """Posodobi napis o stanju predpomnjenega modela."""
        try:
            status = _model_status(self._csv_var.get().strip())
            color  = GREEN if status.startswith("⚡") else ORANGE
            self._model_status_lbl.config(text=f"  {status}", fg=color)
        except Exception:
            pass

    # ── Prenovi model ─────────────────────────────────────────────────────────
    def _prenovi_model(self):
        """Zažene cenik.py --samo-trening v ozadju in pokaže napredek."""
        if self._retrain_running:
            return
        csv_p = self._csv_var.get().strip()
        if not csv_p or not os.path.isfile(csv_p):
            messagebox.showwarning("CSV", "Najprej izberi veljavno CSV datoteko.")
            return

        self._retrain_running = True
        self._retrain_btn.config(state="disabled", text="⏳  Treniram …")
        self._model_status_lbl.config(text="  ⏳  Treniram model, prosim počakaj …",
                                      fg=ORANGE)
        self.update_idletasks()

        cmd = [sys.executable, "-X", "utf8", CENIK,
               "--csv", csv_p, "--samo-trening"]
        threading.Thread(target=self._run_retrain, args=(cmd,), daemon=True).start()

    def _run_retrain(self, cmd):
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", cwd=SCRIPT_DIR)
            out, _ = proc.communicate(timeout=300)
            ok     = "TRENING_OK=1" in out
        except Exception as e:
            ok  = False
            out = str(e)

        def _done():
            self._retrain_running = False
            self._retrain_btn.config(state="normal", text="🔄  Prenovi model")
            if ok:
                self._update_model_status()
                self._set_status("✓  Model uspešno shranjen.", GREEN)
            else:
                self._model_status_lbl.config(
                    text="  ✗  Napaka pri treningu.", fg=RED)
                self._set_status("✗  Napaka pri treningu modela.", RED)

        self.after(0, _done)

    # ── Prikaži rezultat ──────────────────────────────────────────────────────
    def _set_result(self, detail: str, price: str = "", ci: str = "",
                    price_color: str = FG, error: bool = False):
        self._price_lbl.config(text=price or "—",
                               fg=RED if error else price_color)
        self._ci_lbl.config(text=ci)
        self._detail_txt.configure(state="normal")
        self._detail_txt.delete("1.0", "end")
        self._detail_txt.insert("end", detail)
        self._detail_txt.configure(state="disabled")

    def _set_status(self, msg: str, color: str = FG2):
        self._status_lbl.config(text=f"  {msg}", fg=color)

    # ── Napoved ───────────────────────────────────────────────────────────────
    def _napovej(self):
        csv_p = self._csv_var.get().strip()
        if not csv_p or not os.path.isfile(csv_p):
            self._set_result("CSV datoteka ni veljavna.", error=True)
            return

        cmd = [sys.executable, "-X", "utf8", CENIK,
               "--csv",       csv_p,
               "--vrsta",     self._vrsta_var.get(),
               "--kraj",      self._kraj_var.get(),
               "--povrsina",  str(self._povrsina_var.get()),
               "--zemljisce", str(self._zem_var.get()),
               "--leto",      str(self._leto_var.get()),
               "--sobe",      str(self._sobe_var.get()),
               "--energija",  self._enr_var.get()]

        self._price_lbl.config(text="⏳  Računam…", fg=FG2)
        self._ci_lbl.config(text="")
        self._set_status("Treniram model in računam napoved…", ORANGE)
        self.update_idletasks()

        threading.Thread(target=self._run_pred, args=(cmd,), daemon=True).start()

    def _run_pred(self, cmd):
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", cwd=SCRIPT_DIR)
            out, _ = proc.communicate(timeout=120)
        except Exception as exc:
            self.after(0, self._set_result, f"Napaka pri izvajanju:\n{exc}",
                       "", "", FG, True)
            self.after(0, self._set_status, "Napaka.", RED)
            return

        # Razčleni KEY=VALUE vrstice
        kv: dict[str, str] = {}
        for line in out.splitlines():
            if "=" in line and not line.startswith("✓") and not line.startswith(" "):
                k, _, v = line.partition("=")
                kv[k.strip()] = v.strip()

        if "NAPAKA" in kv:
            self.after(0, self._set_result, kv["NAPAKA"], error=True)
            self.after(0, self._set_status, "Napaka.", RED)
            return
        if "NAPOVEDANA" not in kv:
            self.after(0, self._set_result,
                       f"Nepričakovan izhod modela:\n{out[:600]}", error=True)
            self.after(0, self._set_status, "Napaka.", RED)
            return

        cena      = float(kv["NAPOVEDANA"])
        ci_min    = float(kv.get("CI_MIN",  cena*0.85))
        ci_max    = float(kv.get("CI_MAX",  cena*1.15))
        sim_n     = kv.get("PODOBNI_N",  "?")
        sim_min   = float(kv.get("PODOBNI_MIN",  0))
        sim_max   = float(kv.get("PODOBNI_MAX",  0))
        sim_med   = float(kv.get("PODOBNI_MED",  cena))
        sim_povp  = float(kv.get("PODOBNI_POVP", cena))
        sim_filt  = kv.get("PODOBNI_FILTER", "")
        n_vzorcev = kv.get("VZORCI", "?")
        napaka_pct= float(kv.get("NAPAKA_PCT", 0))

        price_str = f"{cena:,.0f} €".replace(",", ".")
        ci_str    = f"90 % interval:  {ci_min:,.0f} – {ci_max:,.0f} €".replace(",", ".")

        # Barva: zelena = blizu medianu podobnih, oranžna = oddaljena
        col = GREEN if napaka_pct < 20 else (ORANGE if napaka_pct < 40 else RED)

        detail = (
            f"Učna baza:        {n_vzorcev} oglasov\n"
            f"Podobni oglasi:   {sim_n}  ({sim_filt})\n"
            f"  Razpon:         {sim_min:,.0f} – {sim_max:,.0f} €\n"
            f"  Povprečje:      {sim_povp:,.0f} €\n"
            f"  Mediana:        {sim_med:,.0f} €\n"
        ).replace(",", ".")

        self.after(0, self._set_result, detail, price_str, ci_str, col)
        self.after(0, self._set_status,
                   f"✓  Napoved izračunana  ({kv.get('VRSTA','')} | "
                   f"{kv.get('KRAJ','')} | {kv.get('POVRSINA','')} m²)", GREEN)


class ScraperGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nepremičnine.net Scraper")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("1050x720")
        self.minsize(820, 580)

        self._process: subprocess.Popen | None = None
        self._running = False

        self._build_ui()
        self._update_estimate()

    # ── Gradnja UI ────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Glava ─────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=ACCENT2, padx=16, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="🏠  Nepremičnine.net  Scraper",
                 bg=ACCENT2, fg="white", font=("Segoe UI", 14, "bold")).pack(side="left")
        self._status_lbl = tk.Label(header, text="● Pripravljen",
                                    bg=ACCENT2, fg="#c8f7c5",
                                    font=("Segoe UI", 10))
        self._status_lbl.pack(side="right")

        # ── Glavno področje (levo: nastavitve, desno: log) ────────────────────
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=12, pady=8)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # Leva plošča – nastavitve
        left = tk.Frame(main, bg=BG, width=380)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        left.pack_propagate(False)

        # Desna plošča – log
        right = tk.Frame(main, bg=BG2)
        right.grid(row=0, column=1, sticky="nsew")

        self._build_left(left)
        self._build_right(right)

        # ── Gumbi na dnu ──────────────────────────────────────────────────────
        bottom = tk.Frame(self, bg=BG, pady=8)
        bottom.pack(fill="x", padx=12)

        self._btn_start = tk.Button(bottom, text="▶  Začni scraping",
                                    bg=GREEN, fg="white", font=FONT_B,
                                    relief="flat", padx=20, pady=8,
                                    activebackground="#388e3c", cursor="hand2",
                                    command=self._start)
        self._btn_start.pack(side="left", padx=(0, 8))

        self._btn_stop = tk.Button(bottom, text="■  Ustavi",
                                   bg=RED, fg="white", font=FONT_B,
                                   relief="flat", padx=14, pady=8,
                                   activebackground="#c62828", cursor="hand2",
                                   state="disabled", command=self._stop)
        self._btn_stop.pack(side="left", padx=(0, 8))

        self._btn_analyze = tk.Button(bottom, text="📊  Analiza",
                                      bg=ORANGE, fg="white", font=FONT_B,
                                      relief="flat", padx=14, pady=8,
                                      activebackground="#e65100", cursor="hand2",
                                      command=self._run_analyze)
        self._btn_analyze.pack(side="left", padx=(0, 8))

        tk.Button(bottom, text="🤖  ML Modeli",
                  bg="#9c27b0", fg="white", font=FONT_B,
                  relief="flat", padx=14, pady=8,
                  activebackground="#6a1b9a", cursor="hand2",
                  command=self._run_ml).pack(side="left", padx=(0, 8))

        tk.Button(bottom, text="🏠  Cenik",
                  bg="#00897b", fg="white", font=FONT_B,
                  relief="flat", padx=14, pady=8,
                  activebackground="#00695c", cursor="hand2",
                  command=self._run_cenik).pack(side="left", padx=(0, 8))

        tk.Button(bottom, text="🗑  Počisti log",
                  bg=BG3, fg=FG2, font=FONT, relief="flat",
                  padx=10, pady=8, cursor="hand2",
                  command=self._clear_log).pack(side="left")

        self._estimate_lbl = tk.Label(bottom, text="",
                                      bg=BG, fg=FG2, font=("Segoe UI", 9))
        self._estimate_lbl.pack(side="right", padx=8)

    def _build_left(self, parent):
        # Celoten levi panel zavijemo v drsnični frame
        sf   = _ScrollFrame(parent, bg=BG)
        sf.pack(fill="both", expand=True)
        body = sf.body

        # ── Akcija ────────────────────────────────────────────────────────
        row0 = tk.Frame(body, bg=BG)
        row0.pack(fill="x", pady=(0, 6))
        tk.Label(row0, text="Akcija:", bg=BG, fg=FG2, font=FONT,
                 width=10, anchor="w").pack(side="left")
        self._akcija_var = tk.StringVar(value="prodaja")
        for a in AKCIJE:
            tk.Radiobutton(row0, text=a.capitalize(), variable=self._akcija_var,
                           value=a, bg=BG, fg=FG, selectcolor=BG3,
                           activebackground=BG, font=FONT).pack(side="left", padx=4)

        # ── Regije checklist ──────────────────────────────────────────────
        self._reg_list = CheckList(body, "Regije", REGIJE, list_h=160)
        self._reg_list.pack(fill="x", pady=(0, 6))
        for var in self._reg_list._vars.values():
            var.trace_add("write", lambda *_: self._update_estimate())
        self._reg_list._all_var.trace_add("write", lambda *_: self._update_estimate())

        # ── Vrste checklist ───────────────────────────────────────────────
        self._vrs_list = CheckList(body, "Vrsta nepremičnine", VRSTE, list_h=130)
        self._vrs_list.pack(fill="x", pady=(0, 6))
        for var in self._vrs_list._vars.values():
            var.trace_add("write", lambda *_: self._update_estimate())
        self._vrs_list._all_var.trace_add("write", lambda *_: self._update_estimate())

        # ── Nastavitve ────────────────────────────────────────────────────
        settings = tk.LabelFrame(body, text=" Nastavitve ", bg=BG, fg=FG2,
                                 font=FONT, bd=1, relief="groove")
        settings.pack(fill="x", pady=(0, 4))

        def row(f, lbl, widget_fn, **kw):
            r = tk.Frame(f, bg=BG)
            r.pack(fill="x", pady=2, padx=6)
            tk.Label(r, text=lbl, bg=BG, fg=FG2, font=FONT,
                     width=14, anchor="w").pack(side="left")
            widget_fn(r, **kw).pack(side="left", fill="x", expand=True)

        self._strani_var = tk.IntVar(value=0)
        row(settings, "Maks strani:",
            lambda p, **kw: tk.Spinbox(p, from_=0, to=999,
                textvariable=self._strani_var, width=6,
                bg=BG3, fg=FG, insertbackground=FG,
                buttonbackground=BG3, relief="flat",
                highlightthickness=0, font=FONT,
                command=self._update_estimate))
        tk.Label(settings, text="  (0 = vse strani samodejno)",
                 bg=BG, fg=FG2, font=("Segoe UI", 8)).pack(anchor="w", padx=20, pady=(0, 2))

        self._delay_var = tk.DoubleVar(value=1.5)
        row(settings, "Zamik (s):",
            lambda p, **kw: tk.Spinbox(p, from_=0.5, to=10.0, increment=0.5,
                textvariable=self._delay_var, width=6,
                bg=BG3, fg=FG, insertbackground=FG,
                buttonbackground=BG3, relief="flat",
                highlightthickness=0, font=FONT))

        # CSV datoteka
        csv_row = tk.Frame(settings, bg=BG)
        csv_row.pack(fill="x", pady=2, padx=6)
        tk.Label(csv_row, text="CSV izhod:", bg=BG, fg=FG2, font=FONT,
                 width=14, anchor="w").pack(side="left")
        self._csv_var = tk.StringVar(value="")
        tk.Entry(csv_row, textvariable=self._csv_var, width=22,
                 bg=BG3, fg=FG, insertbackground=FG,
                 relief="flat", font=FONT).pack(side="left", fill="x", expand=True)
        tk.Button(csv_row, text="…", bg=BG3, fg=FG, relief="flat",
                  cursor="hand2", command=self._browse_csv).pack(side="left", padx=2)

        # Checkboxes
        flags = tk.Frame(settings, bg=BG)
        flags.pack(fill="x", pady=4, padx=6)
        self._headless_var = tk.BooleanVar(value=False)
        tk.Checkbutton(flags, text="Brez okna (headless)",
                       variable=self._headless_var, bg=BG, fg=FG,
                       selectcolor=BG3, activebackground=BG, font=FONT).pack(side="left")
        self._nodb_var = tk.BooleanVar(value=True)
        tk.Checkbutton(flags, text="Samo CSV",
                       variable=self._nodb_var, bg=BG, fg=FG,
                       selectcolor=BG3, activebackground=BG, font=FONT).pack(side="left", padx=12)

    def _build_right(self, parent):
        tk.Label(parent, text="  Izhod / log", bg=BG2, fg=FG2,
                 font=FONT_B, anchor="w").pack(fill="x", padx=6, pady=(6, 2))

        self._log = tk.Text(parent, bg="#12121e", fg="#d4d4f0",
                            insertbackground=FG, font=MONO,
                            relief="flat", wrap="none", state="disabled")
        sb_y = tk.Scrollbar(parent, orient="vertical", command=self._log.yview,
                            bg=BG3, troughcolor=BG2)
        sb_x = tk.Scrollbar(parent, orient="horizontal", command=self._log.xview,
                            bg=BG3, troughcolor=BG2)
        self._log.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self._log.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # Tag barve
        self._log.tag_configure("ok",   foreground="#4caf50")
        self._log.tag_configure("warn", foreground="#ff9800")
        self._log.tag_configure("err",  foreground="#f44336")
        self._log.tag_configure("info", foreground="#64b5f6")
        self._log.tag_configure("head", foreground="#ce93d8", font=MONO + ("bold",))

    # ── Pomožne metode ────────────────────────────────────────────────────────
    def _browse_csv(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV datoteke", "*.csv"), ("Vse datoteke", "*.*")],
            initialdir=SCRIPT_DIR,
            title="Shrani CSV kot…")
        if path:
            self._csv_var.set(path)

    def _update_estimate(self, *_):
        try:
            n_reg  = len(self._reg_list.selected_slugs())
            n_vrs  = len(self._vrs_list.selected_slugs())
            n_str  = self._strani_var.get()
            combos = n_reg * n_vrs
            if n_str == 0:
                # Neznano število strani – prikažemo samo kombinacije
                self._estimate_lbl.config(
                    text=f"{combos} kombinacij × vse strani  (čas neznano)")
            else:
                pages = combos * n_str
                secs  = pages * 35
                h, m  = divmod(secs // 60, 60)
                est   = f"{h}h {m}min" if h else f"~{m} min"
                self._estimate_lbl.config(
                    text=f"{combos} kombinacij × maks {n_str} strani ≈ {est}")
        except Exception:
            pass

    def _log_write(self, text, tag=""):
        self._log.configure(state="normal")
        if tag:
            self._log.insert("end", text, tag)
        else:
            self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _set_running(self, running: bool):
        self._running = running
        state_start = "disabled" if running else "normal"
        state_stop  = "normal"   if running else "disabled"
        self._btn_start.config(state=state_start)
        self._btn_stop.config(state=state_stop)
        if running:
            self._status_lbl.config(text="● V teku…", fg="#fff176")
        else:
            self._status_lbl.config(text="● Pripravljen", fg="#c8f7c5")

    # ── Scraping ──────────────────────────────────────────────────────────────
    def _start(self):
        reg_slugs = self._reg_list.selected_slugs()
        vrs_slugs = self._vrs_list.selected_slugs()

        if not reg_slugs:
            messagebox.showwarning("Izbor", "Izberi vsaj eno regijo!")
            return
        if not vrs_slugs:
            messagebox.showwarning("Izbor", "Izberi vsaj eno vrsto nepremičnine!")
            return

        n_comb = len(reg_slugs) * len(vrs_slugs)
        strani_txt = "vse strani" if self._strani_var.get() == 0 else f"{self._strani_var.get()} strani"
        if n_comb > 10 or (n_comb > 3 and self._strani_var.get() == 0):
            ok = messagebox.askyesno(
                "Veliko kombinacij",
                f"Izbranih je {n_comb} kombinacij × {strani_txt}.\n"
                "To lahko traja dlje časa. Nadaljuješ?")
            if not ok:
                return

        # Sestavi argumente
        cmd = [sys.executable, "-X", "utf8", SCRAPER,
               "--regije", ",".join(reg_slugs),
               "--vrste",  ",".join(vrs_slugs),
               "--akcija", self._akcija_var.get(),
               "--strani", str(self._strani_var.get()),
               "--delay",  str(self._delay_var.get())]

        if self._nodb_var.get():
            cmd.append("--csv")
        if self._headless_var.get():
            cmd.append("--headless")
        if self._csv_var.get().strip():
            cmd += ["--izhod", self._csv_var.get().strip()]

        self._clear_log()
        ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        self._log_write(f"{'='*60}\n", "head")
        self._log_write(f"  Zagon: {ts}\n", "head")
        self._log_write(f"  Regije: {', '.join(self._reg_list.selected_names())}\n", "info")
        self._log_write(f"  Vrste:  {', '.join(self._vrs_list.selected_names())}\n", "info")
        self._log_write(f"  Ukaz: {' '.join(cmd)}\n", "info")
        self._log_write(f"{'='*60}\n\n", "head")

        self._set_running(True)
        threading.Thread(target=self._run_process, args=(cmd,), daemon=True).start()

    def _run_process(self, cmd):
        try:
            self._process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=SCRIPT_DIR,
            )
            for line in self._process.stdout:
                tag = ""
                ll = line.lower()
                if "✓" in line or "ok" in ll:
                    tag = "ok"
                elif "⚠" in line or "warn" in ll:
                    tag = "warn"
                elif "✗" in line or "error" in ll or "napaka" in ll:
                    tag = "err"
                elif line.startswith("  ") and "€" in line:
                    tag = "ok"
                self.after(0, self._log_write, line, tag)

            self._process.wait()
            rc = self._process.returncode
            self.after(0, self._on_done, rc)
        except Exception as e:
            self.after(0, self._log_write, f"\n✗ Napaka: {e}\n", "err")
            self.after(0, self._set_running, False)

    def _on_done(self, rc: int):
        self._set_running(False)
        if rc == 0:
            self._log_write("\n✓  Scraping končan.\n", "ok")
            self._status_lbl.config(text="● Končano ✓", fg="#4caf50")
        else:
            self._log_write(f"\n✗  Proces končan z napako (koda {rc}).\n", "err")
            self._status_lbl.config(text="● Napaka", fg="#f44336")

    def _stop(self):
        if self._process and self._process.poll() is None:
            self._process.terminate()
            self._log_write("\n⚠  Scraping prekinjen.\n", "warn")
        self._set_running(False)

    # ── Analiza ───────────────────────────────────────────────────────────────
    def _run_analyze(self):
        AnalysisDialog(self)

    def _run_ml(self):
        MLDialog(self)

    def _run_cenik(self):
        CenikDialog(self)


# ── Zagon ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    app = ScraperGUI()
    app.mainloop()

