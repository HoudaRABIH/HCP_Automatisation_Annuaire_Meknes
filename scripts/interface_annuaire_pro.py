from __future__ import annotations

from pathlib import Path
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# =============================================================================
# CONFIG
# =============================================================================

APP_TITLE = "HCP Annuaire Studio"
ROOT = Path(".")
SCRIPTS_DIR = ROOT / "scripts"
BACKEND = SCRIPTS_DIR / "generer_annuaire_complet.py"
DATA_DIR = ROOT / "donnees_brutes"
OUTPUT_DIR = ROOT / "sorties"
BASE_DOC = OUTPUT_DIR / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"
FINAL_DOC = OUTPUT_DIR / "Annuaire_HCP_2025_FINAL.docx"
REPORT_FILE = OUTPUT_DIR / "rapport_generation_complete.txt"

# HCP-inspired palette
COLORS = {
    "bg": "#F5F7FA",
    "surface": "#FFFFFF",
    "surface_alt": "#F8FAFC",
    "border": "#E3E8EF",
    "text": "#17202A",
    "muted": "#6B7280",
    "primary": "#8B1538",
    "primary_dark": "#68102A",
    "primary_soft": "#F7EAF0",
    "success": "#15803D",
    "success_soft": "#EAF7EF",
    "warning": "#B7791F",
    "warning_soft": "#FFF7E6",
    "danger": "#B42318",
    "danger_soft": "#FDECEC",
    "info": "#2563EB",
    "info_soft": "#EEF4FF",
    "sidebar": "#151B26",
    "sidebar_alt": "#1D2634",
    "sidebar_text": "#F8FAFC",
    "sidebar_muted": "#AEB8C7",
}

STEPS = [
    "Justice",
    "Culture / Loisirs",
    "Habitat / Foncier",
    "Population",
    "Industrie",
    "Énergie / Eau",
    "Validation finale",
]


# =============================================================================
# HELPERS
# =============================================================================

def open_path(path: Path):
    if not path.exists():
        raise FileNotFoundError(str(path))

    p = str(path.resolve())

    if os.name == "nt":
        os.startfile(p)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", p])
    else:
        subprocess.Popen(["xdg-open", p])


def count_excel_files() -> int:
    if not DATA_DIR.exists():
        return 0
    return len([
        p for p in DATA_DIR.glob("*.xlsx")
        if not p.name.startswith("~$")
    ])


def human_size(path: Path) -> str:
    if not path.exists():
        return "—"
    size = path.stat().st_size
    if size < 1024:
        return f"{size} o"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} Ko"
    return f"{size / (1024 * 1024):.2f} Mo"


# =============================================================================
# MAIN APP
# =============================================================================

class AnnuaireStudio(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1260x800")
        self.minsize(1080, 700)
        self.configure(bg=COLORS["bg"])

        self.proc: subprocess.Popen | None = None
        self.worker_thread: threading.Thread | None = None
        self.events = queue.Queue()

        self.run_started_at: float | None = None
        self.current_step = 0
        self.last_status = "ready"

        self._configure_styles()
        self._build_layout()
        self._refresh_project_state()

        self.after(100, self._poll_events)
        self.after(1000, self._tick_elapsed)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # -------------------------------------------------------------------------
    # STYLE
    # -------------------------------------------------------------------------

    def _configure_styles(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "App.TFrame",
            background=COLORS["bg"],
        )
        style.configure(
            "Surface.TFrame",
            background=COLORS["surface"],
        )
        style.configure(
            "Sidebar.TFrame",
            background=COLORS["sidebar"],
        )

        style.configure(
            "Title.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=("Segoe UI", 22, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10),
        )

        style.configure(
            "CardTitle.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "CardValue.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "CardHint.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 8),
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(20, 11),
            foreground="white",
            background=COLORS["primary"],
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COLORS["primary_dark"]),
                ("disabled", "#B8A0AA"),
            ],
            foreground=[("disabled", "#F4EEF1")],
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(13, 8),
            foreground=COLORS["text"],
            background=COLORS["surface"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLORS["surface_alt"])],
        )

        style.configure(
            "Danger.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(13, 8),
            foreground=COLORS["danger"],
            background=COLORS["danger_soft"],
            borderwidth=0,
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#F9D7D4")],
        )

        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#E9EDF2",
            background=COLORS["primary"],
            bordercolor="#E9EDF2",
            lightcolor=COLORS["primary"],
            darkcolor=COLORS["primary"],
            thickness=10,
        )

    # -------------------------------------------------------------------------
    # LAYOUT
    # -------------------------------------------------------------------------

    def _build_layout(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self._build_sidebar()

        self.main = tk.Frame(self, bg=COLORS["bg"])
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(3, weight=1)

        self._build_header()
        self._build_stats()
        self._build_run_panel()
        self._build_work_area()

    def _build_sidebar(self):
        sidebar = tk.Frame(
            self,
            bg=COLORS["sidebar"],
            width=248,
        )
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        brand = tk.Frame(sidebar, bg=COLORS["sidebar"])
        brand.pack(fill="x", padx=22, pady=(28, 20))

        logo = tk.Canvas(
            brand,
            width=46,
            height=46,
            bg=COLORS["sidebar"],
            highlightthickness=0,
        )
        logo.pack(side="left")
        logo.create_oval(3, 3, 43, 43, fill=COLORS["primary"], outline="")
        logo.create_text(
            23, 23,
            text="HCP",
            fill="white",
            font=("Segoe UI", 10, "bold")
        )

        brand_text = tk.Frame(brand, bg=COLORS["sidebar"])
        brand_text.pack(side="left", padx=(12, 0))

        tk.Label(
            brand_text,
            text="Annuaire Studio",
            bg=COLORS["sidebar"],
            fg=COLORS["sidebar_text"],
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")

        tk.Label(
            brand_text,
            text="Préfecture de Meknès",
            bg=COLORS["sidebar"],
            fg=COLORS["sidebar_muted"],
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        tk.Frame(
            sidebar,
            bg="#2B3545",
            height=1,
        ).pack(fill="x", padx=22, pady=(0, 18))

        self.nav_buttons = {}

        nav_items = [
            ("dashboard", "▦", "Tableau de bord"),
            ("logs", "≡", "Journal"),
            ("output", "▣", "Résultats"),
        ]

        for key, icon, label in nav_items:
            btn = tk.Button(
                sidebar,
                text=f"  {icon}    {label}",
                anchor="w",
                relief="flat",
                bd=0,
                bg=COLORS["sidebar"],
                fg=COLORS["sidebar_text"],
                activebackground=COLORS["sidebar_alt"],
                activeforeground="white",
                font=("Segoe UI", 10),
                padx=18,
                pady=12,
                cursor="hand2",
                command=lambda k=key: self._nav_action(k),
            )
            btn.pack(fill="x", padx=10, pady=2)
            self.nav_buttons[key] = btn

        self.nav_buttons["dashboard"].configure(
            bg=COLORS["primary"],
            font=("Segoe UI", 10, "bold"),
        )

        tk.Frame(
            sidebar,
            bg="#2B3545",
            height=1,
        ).pack(side="bottom", fill="x", padx=22, pady=(10, 0))

        bottom = tk.Frame(sidebar, bg=COLORS["sidebar"])
        bottom.pack(side="bottom", fill="x", padx=22, pady=22)

        tk.Label(
            bottom,
            text="Backend validé",
            bg=COLORS["sidebar"],
            fg="#83E0A3",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")

        tk.Label(
            bottom,
            text="Pipeline statistique HCP",
            bg=COLORS["sidebar"],
            fg=COLORS["sidebar_muted"],
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(2, 0))

    def _build_header(self):
        header = tk.Frame(self.main, bg=COLORS["bg"])
        header.grid(row=0, column=0, sticky="ew", padx=30, pady=(26, 16))
        header.columnconfigure(0, weight=1)

        left = tk.Frame(header, bg=COLORS["bg"])
        left.grid(row=0, column=0, sticky="w")

        ttk.Label(
            left,
            text="Générateur Annuaire HCP",
            style="Title.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            left,
            text="Générez, contrôlez et exportez l'annuaire statistique final depuis un seul écran.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        right = tk.Frame(header, bg=COLORS["bg"])
        right.grid(row=0, column=1, sticky="e")

        self.status_badge = tk.Label(
            right,
            text="●  PRÊT",
            bg=COLORS["success_soft"],
            fg=COLORS["success"],
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=7,
        )
        self.status_badge.pack(side="right")

    def _build_stats(self):
        stats = tk.Frame(self.main, bg=COLORS["bg"])
        stats.grid(row=1, column=0, sticky="ew", padx=30, pady=(0, 16))

        for i in range(4):
            stats.columnconfigure(i, weight=1, uniform="stat")

        self.stat_excel = self._create_stat_card(
            stats, 0,
            "FICHIERS EXCEL",
            "—",
            "Données brutes détectées"
        )
        self.stat_base = self._create_stat_card(
            stats, 1,
            "MODÈLE WORD",
            "—",
            "Base validée"
        )
        self.stat_output = self._create_stat_card(
            stats, 2,
            "DERNIER WORD",
            "—",
            "Document final"
        )
        self.stat_time = self._create_stat_card(
            stats, 3,
            "DURÉE",
            "00:00",
            "Dernière / actuelle"
        )

    def _create_stat_card(self, parent, col, title, value, hint):
        outer = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        outer.grid(
            row=0,
            column=col,
            sticky="nsew",
            padx=(0 if col == 0 else 6, 0 if col == 3 else 6),
        )

        inner = tk.Frame(outer, bg=COLORS["surface"])
        inner.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(
            inner,
            text=title,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w")

        value_label = tk.Label(
            inner,
            text=value,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 17, "bold"),
        )
        value_label.pack(anchor="w", pady=(5, 2))

        tk.Label(
            inner,
            text=hint,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        return value_label

    def _build_run_panel(self):
        panel = tk.Frame(
            self.main,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        panel.grid(row=2, column=0, sticky="ew", padx=30, pady=(0, 16))
        panel.columnconfigure(0, weight=1)

        top = tk.Frame(panel, bg=COLORS["surface"])
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 10))
        top.columnconfigure(0, weight=1)

        title_box = tk.Frame(top, bg=COLORS["surface"])
        title_box.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_box,
            text="Pipeline de génération",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w")

        self.run_hint = tk.Label(
            title_box,
            text="7 étapes validées • arrêt automatique en cas d'erreur",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
        )
        self.run_hint.pack(anchor="w", pady=(3, 0))

        actions = tk.Frame(top, bg=COLORS["surface"])
        actions.grid(row=0, column=1, sticky="e")

        self.verify_btn = ttk.Button(
            actions,
            text="Vérifier le projet",
            style="Secondary.TButton",
            command=self._manual_preflight,
        )
        self.verify_btn.pack(side="left", padx=(0, 8))

        self.run_btn = ttk.Button(
            actions,
            text="▶  Générer l'annuaire",
            style="Primary.TButton",
            command=self._start_generation,
        )
        self.run_btn.pack(side="left")

        progress_wrap = tk.Frame(panel, bg=COLORS["surface"])
        progress_wrap.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 18))
        progress_wrap.columnconfigure(0, weight=1)

        progress_head = tk.Frame(progress_wrap, bg=COLORS["surface"])
        progress_head.grid(row=0, column=0, sticky="ew")
        progress_head.columnconfigure(0, weight=1)

        self.step_label = tk.Label(
            progress_head,
            text="Prêt à générer",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 9, "bold"),
        )
        self.step_label.grid(row=0, column=0, sticky="w")

        self.percent_label = tk.Label(
            progress_head,
            text="0%",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9, "bold"),
        )
        self.percent_label.grid(row=0, column=1, sticky="e")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            progress_wrap,
            variable=self.progress_var,
            maximum=100,
            style="Modern.Horizontal.TProgressbar",
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(8, 12))

        steps_frame = tk.Frame(progress_wrap, bg=COLORS["surface"])
        steps_frame.grid(row=2, column=0, sticky="ew")

        self.step_widgets = []

        for i, step in enumerate(STEPS, start=1):
            cell = tk.Frame(steps_frame, bg=COLORS["surface"])
            cell.pack(side="left", fill="x", expand=True)

            dot = tk.Label(
                cell,
                text=str(i),
                bg="#E9EDF2",
                fg=COLORS["muted"],
                width=2,
                height=1,
                font=("Segoe UI", 8, "bold"),
            )
            dot.pack()

            label = tk.Label(
                cell,
                text=step,
                bg=COLORS["surface"],
                fg=COLORS["muted"],
                font=("Segoe UI", 7),
                wraplength=92,
                justify="center",
            )
            label.pack(pady=(4, 0))

            self.step_widgets.append((dot, label))

    def _build_work_area(self):
        work = tk.Frame(self.main, bg=COLORS["bg"])
        work.grid(row=3, column=0, sticky="nsew", padx=30, pady=(0, 26))
        work.columnconfigure(0, weight=3)
        work.columnconfigure(1, weight=2)
        work.rowconfigure(0, weight=1)

        self._build_log_panel(work)
        self._build_output_panel(work)

    def _build_log_panel(self, parent):
        panel = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        panel.rowconfigure(1, weight=1)
        panel.columnconfigure(0, weight=1)

        header = tk.Frame(panel, bg=COLORS["surface"])
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        header.columnconfigure(0, weight=1)

        tk.Label(
            header,
            text="Journal d'exécution",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self.cancel_btn = ttk.Button(
            header,
            text="Arrêter",
            style="Danger.TButton",
            state="disabled",
            command=self._cancel_generation,
        )
        self.cancel_btn.grid(row=0, column=1, sticky="e")

        log_wrap = tk.Frame(panel, bg=COLORS["surface"])
        log_wrap.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        log_wrap.rowconfigure(0, weight=1)
        log_wrap.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_wrap,
            wrap="word",
            relief="flat",
            bd=0,
            bg="#0F172A",
            fg="#DCE7F3",
            insertbackground="white",
            selectbackground="#334155",
            font=("Cascadia Mono", 9),
            padx=12,
            pady=10,
            state="disabled",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(
            log_wrap,
            orient="vertical",
            command=self.log_text.yview,
        )
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        self.log_text.tag_configure("success", foreground="#86EFAC")
        self.log_text.tag_configure("danger", foreground="#FCA5A5")
        self.log_text.tag_configure("warning", foreground="#FDE68A")
        self.log_text.tag_configure("step", foreground="#93C5FD")
        self.log_text.tag_configure("muted", foreground="#94A3B8")

    def _build_output_panel(self, parent):
        panel = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        panel.columnconfigure(0, weight=1)

        tk.Label(
            panel,
            text="Résultat",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=18, pady=(16, 4))

        tk.Label(
            panel,
            text="Document final généré",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
        ).pack(anchor="w", padx=18)

        self.output_card = tk.Frame(
            panel,
            bg=COLORS["surface_alt"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        self.output_card.pack(fill="x", padx=18, pady=16)

        output_inner = tk.Frame(
            self.output_card,
            bg=COLORS["surface_alt"],
        )
        output_inner.pack(fill="x", padx=14, pady=14)

        tk.Label(
            output_inner,
            text="WORD",
            bg=COLORS["primary_soft"],
            fg=COLORS["primary"],
            font=("Segoe UI", 8, "bold"),
            padx=8,
            pady=4,
        ).pack(anchor="w")

        self.output_name = tk.Label(
            output_inner,
            text=FINAL_DOC.name,
            bg=COLORS["surface_alt"],
            fg=COLORS["text"],
            font=("Segoe UI", 9, "bold"),
            anchor="w",
            justify="left",
            wraplength=320,
        )
        self.output_name.pack(anchor="w", pady=(10, 4))

        self.output_meta = tk.Label(
            output_inner,
            text="Pas encore généré",
            bg=COLORS["surface_alt"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
        )
        self.output_meta.pack(anchor="w")

        buttons = tk.Frame(panel, bg=COLORS["surface"])
        buttons.pack(fill="x", padx=18)

        self.open_word_btn = ttk.Button(
            buttons,
            text="Ouvrir le Word",
            style="Secondary.TButton",
            command=lambda: self._safe_open(FINAL_DOC),
        )
        self.open_word_btn.pack(fill="x", pady=(0, 8))

        self.open_report_btn = ttk.Button(
            buttons,
            text="Ouvrir le rapport",
            style="Secondary.TButton",
            command=lambda: self._safe_open(REPORT_FILE),
        )
        self.open_report_btn.pack(fill="x", pady=(0, 8))

        ttk.Button(
            buttons,
            text="Ouvrir le dossier sorties",
            style="Secondary.TButton",
            command=lambda: self._safe_open(OUTPUT_DIR),
        ).pack(fill="x")

        tk.Frame(
            panel,
            bg=COLORS["border"],
            height=1,
        ).pack(fill="x", padx=18, pady=18)

        tk.Label(
            panel,
            text="État du projet",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", padx=18)

        self.health_frame = tk.Frame(panel, bg=COLORS["surface"])
        self.health_frame.pack(fill="x", padx=18, pady=(10, 16))

        self.health_labels = {}
        for key, label in [
            ("excel", "Données Excel"),
            ("base", "Modèle Word"),
            ("backend", "Backend"),
        ]:
            row = tk.Frame(self.health_frame, bg=COLORS["surface"])
            row.pack(fill="x", pady=4)

            tk.Label(
                row,
                text=label,
                bg=COLORS["surface"],
                fg=COLORS["muted"],
                font=("Segoe UI", 8),
            ).pack(side="left")

            status = tk.Label(
                row,
                text="—",
                bg=COLORS["surface"],
                fg=COLORS["muted"],
                font=("Segoe UI", 8, "bold"),
            )
            status.pack(side="right")
            self.health_labels[key] = status

    # -------------------------------------------------------------------------
    # NAV
    # -------------------------------------------------------------------------

    def _nav_action(self, key):
        if key == "dashboard":
            return
        if key == "logs":
            self.log_text.focus_set()
            self.log_text.see("end")
        elif key == "output":
            self._safe_open(OUTPUT_DIR)

    # -------------------------------------------------------------------------
    # PROJECT STATE
    # -------------------------------------------------------------------------

    def _refresh_project_state(self):
        excel_count = count_excel_files()
        self.stat_excel.configure(text=str(excel_count))

        self.stat_base.configure(
            text="OK" if BASE_DOC.exists() else "MANQUANT",
            fg=COLORS["success"] if BASE_DOC.exists() else COLORS["danger"],
        )

        if FINAL_DOC.exists():
            self.stat_output.configure(
                text=human_size(FINAL_DOC),
                fg=COLORS["success"],
            )
            modified = datetime.fromtimestamp(
                FINAL_DOC.stat().st_mtime
            ).strftime("%d/%m/%Y %H:%M")
            self.output_meta.configure(
                text=f"{human_size(FINAL_DOC)} • {modified}"
            )
            self.open_word_btn.configure(state="normal")
        else:
            self.stat_output.configure(text="—")
            self.output_meta.configure(text="Pas encore généré")
            self.open_word_btn.configure(state="disabled")

        self.health_labels["excel"].configure(
            text=f"✓ {excel_count} fichiers" if excel_count >= 20 else f"⚠ {excel_count} fichiers",
            fg=COLORS["success"] if excel_count >= 20 else COLORS["warning"],
        )

        self.health_labels["base"].configure(
            text="✓ Disponible" if BASE_DOC.exists() else "✕ Introuvable",
            fg=COLORS["success"] if BASE_DOC.exists() else COLORS["danger"],
        )

        self.health_labels["backend"].configure(
            text="✓ Disponible" if BACKEND.exists() else "✕ Introuvable",
            fg=COLORS["success"] if BACKEND.exists() else COLORS["danger"],
        )

        if REPORT_FILE.exists():
            self.open_report_btn.configure(state="normal")
        else:
            self.open_report_btn.configure(state="disabled")

    def _preflight(self):
        errors = []
        warnings = []

        excel_count = count_excel_files()

        if not BACKEND.exists():
            errors.append(f"Backend introuvable : {BACKEND}")

        if not BASE_DOC.exists():
            errors.append(f"Modèle Word introuvable : {BASE_DOC}")

        if not DATA_DIR.exists():
            errors.append(f"Dossier Excel introuvable : {DATA_DIR}")
        elif excel_count < 20:
            warnings.append(
                f"Seulement {excel_count} fichiers Excel détectés."
            )

        return errors, warnings

    def _manual_preflight(self):
        errors, warnings = self._preflight()
        self._refresh_project_state()

        if errors:
            messagebox.showerror(
                "Vérification du projet",
                "Le projet n'est pas prêt :\n\n• "
                + "\n• ".join(errors)
            )
            return

        msg = "Projet prêt pour la génération."
        if warnings:
            msg += "\n\nAvertissements :\n• " + "\n• ".join(warnings)

        messagebox.showinfo("Vérification du projet", msg)

    # -------------------------------------------------------------------------
    # GENERATION
    # -------------------------------------------------------------------------

    def _start_generation(self):
        if self.proc is not None:
            return

        errors, warnings = self._preflight()

        if errors:
            messagebox.showerror(
                "Impossible de générer",
                "Corrigez d'abord :\n\n• " + "\n• ".join(errors)
            )
            return

        if warnings:
            if not messagebox.askyesno(
                "Avertissement",
                "\n".join(warnings) + "\n\nContinuer quand même ?"
            ):
                return

        self._clear_log()
        self._append_log(
            "HCP Annuaire Studio\n"
            "Démarrage du pipeline de génération...\n\n",
            "step"
        )

        self.current_step = 0
        self.run_started_at = time.time()
        self.last_status = "running"

        self._set_status("running", "GÉNÉRATION EN COURS")
        self.run_btn.configure(state="disabled")
        self.verify_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        self.progress_var.set(2)
        self.percent_label.configure(text="2%")
        self.step_label.configure(text="Pré-contrôles...")

        for dot, label in self.step_widgets:
            dot.configure(
                text="○",
                bg="#E9EDF2",
                fg=COLORS["muted"],
            )
            label.configure(fg=COLORS["muted"])

        self.worker_thread = threading.Thread(
            target=self._generation_worker,
            daemon=True,
        )
        self.worker_thread.start()

    def _generation_worker(self):
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"

            creationflags = 0
            if os.name == "nt":
                creationflags = getattr(
                    subprocess,
                    "CREATE_NO_WINDOW",
                    0
                )

            self.proc = subprocess.Popen(
                [sys.executable, "-u", str(BACKEND)],
                cwd=str(ROOT.resolve()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                creationflags=creationflags,
            )

            assert self.proc.stdout is not None

            for line in self.proc.stdout:
                self.events.put(("line", line))

            code = self.proc.wait()
            self.events.put(("finished", code))

        except Exception as exc:
            self.events.put(("exception", str(exc)))
        finally:
            self.proc = None

    def _cancel_generation(self):
        if self.proc is None:
            return

        if not messagebox.askyesno(
            "Arrêter la génération",
            "La génération actuelle sera interrompue.\n\nContinuer ?"
        ):
            return

        try:
            self.proc.terminate()
            self._append_log(
                "\nArrêt demandé par l'utilisateur...\n",
                "warning"
            )
            self.step_label.configure(text="Arrêt en cours...")
        except Exception as exc:
            messagebox.showerror("Erreur", str(exc))

    def _handle_backend_line(self, line: str):
        clean = line.rstrip("\n")

        tag = None
        upper = clean.upper()

        if "ECHEC" in upper or "TRACEBACK" in upper or "ERROR" in upper:
            tag = "danger"
        elif "OK" in upper or "VALIDÉ" in upper or "VALIDEE" in upper:
            tag = "success"
        elif re.search(r"\[\d+/\d+\]", clean):
            tag = "step"
        elif "WARNING" in upper or "AVERT" in upper:
            tag = "warning"

        self._append_log(line, tag)

        # Parse validated backend stage: [x/7]
        m = re.search(r"\[(\d+)/7\]\s*(.+)", clean)
        if m:
            step = int(m.group(1))
            name = m.group(2).strip()

            self.current_step = step
            pct = max(5, min(92, round((step - 0.15) / 7 * 100)))

            self.progress_var.set(pct)
            self.percent_label.configure(text=f"{pct}%")
            self.step_label.configure(
                text=f"Étape {step}/7 • {name}"
            )
            self._set_step_state(step)

        if "VALIDATION BACKEND" in upper:
            self.progress_var.set(95)
            self.percent_label.configure(text="95%")
            self.step_label.configure(text="Validation finale...")

        if "GENERATION COMPLETE : OK" in upper:
            self.progress_var.set(100)
            self.percent_label.configure(text="100%")
            self.step_label.configure(text="Génération terminée")

    def _set_step_state(self, active_step: int):
        for i, (dot, label) in enumerate(self.step_widgets, start=1):
            if i < active_step:
                dot.configure(
                    text="✓",
                    bg=COLORS["success_soft"],
                    fg=COLORS["success"],
                )
                label.configure(fg=COLORS["success"])
            elif i == active_step:
                dot.configure(
                    text=str(i),
                    bg=COLORS["primary"],
                    fg="white",
                )
                label.configure(
                    fg=COLORS["primary"],
                    font=("Segoe UI", 7, "bold"),
                )
            else:
                dot.configure(
                    text="○",
                    bg="#E9EDF2",
                    fg=COLORS["muted"],
                )
                label.configure(
                    fg=COLORS["muted"],
                    font=("Segoe UI", 7),
                )

    def _finish_success(self):
        self.last_status = "success"
        self._set_status("success", "TERMINÉ AVEC SUCCÈS")
        self.progress_var.set(100)
        self.percent_label.configure(text="100%")
        self.step_label.configure(text="Annuaire généré avec succès")

        for dot, label in self.step_widgets:
            dot.configure(
                text="✓",
                bg=COLORS["success_soft"],
                fg=COLORS["success"],
            )
            label.configure(fg=COLORS["success"])

        self.run_btn.configure(state="normal")
        self.verify_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")

        self._refresh_project_state()

        self._append_log(
            "\n✓ Pipeline terminé avec succès.\n",
            "success"
        )

        messagebox.showinfo(
            "Génération terminée",
            "L'annuaire HCP a été généré avec succès.\n\n"
            f"{FINAL_DOC}"
        )

    def _finish_failure(self, message="Une étape a échoué."):
        self.last_status = "failure"
        self._set_status("failure", "ÉCHEC")
        self.step_label.configure(text="Génération interrompue")

        self.run_btn.configure(state="normal")
        self.verify_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")

        self._append_log(
            f"\n✕ {message}\n",
            "danger"
        )

        messagebox.showerror(
            "Génération échouée",
            message
            + "\n\nConsultez le journal pour identifier l'étape concernée."
        )

    # -------------------------------------------------------------------------
    # EVENT LOOP
    # -------------------------------------------------------------------------

    def _poll_events(self):
        try:
            while True:
                event, payload = self.events.get_nowait()

                if event == "line":
                    self._handle_backend_line(payload)

                elif event == "finished":
                    code = int(payload)

                    if code == 0 and FINAL_DOC.exists():
                        self._finish_success()
                    else:
                        self._finish_failure(
                            f"Le backend s'est arrêté avec le code {code}."
                        )

                elif event == "exception":
                    self._finish_failure(str(payload))

        except queue.Empty:
            pass

        self.after(100, self._poll_events)

    def _tick_elapsed(self):
        if self.run_started_at and self.last_status == "running":
            elapsed = int(time.time() - self.run_started_at)
            mins, secs = divmod(elapsed, 60)
            self.stat_time.configure(
                text=f"{mins:02d}:{secs:02d}"
            )
        self.after(1000, self._tick_elapsed)

    # -------------------------------------------------------------------------
    # UI HELPERS
    # -------------------------------------------------------------------------

    def _set_status(self, status: str, label: str):
        if status == "running":
            self.status_badge.configure(
                text=f"●  {label}",
                bg=COLORS["info_soft"],
                fg=COLORS["info"],
            )
        elif status == "success":
            self.status_badge.configure(
                text=f"●  {label}",
                bg=COLORS["success_soft"],
                fg=COLORS["success"],
            )
        elif status == "failure":
            self.status_badge.configure(
                text=f"●  {label}",
                bg=COLORS["danger_soft"],
                fg=COLORS["danger"],
            )
        else:
            self.status_badge.configure(
                text=f"●  {label}",
                bg=COLORS["success_soft"],
                fg=COLORS["success"],
            )

    def _append_log(self, text, tag=None):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text, tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _safe_open(self, path: Path):
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(
                "Ouverture impossible",
                str(exc)
            )

    def _on_close(self):
        if self.proc is not None:
            if not messagebox.askyesno(
                "Quitter",
                "Une génération est en cours.\n"
                "Voulez-vous vraiment fermer l'application ?"
            ):
                return
            try:
                self.proc.terminate()
            except Exception:
                pass

        self.destroy()


if __name__ == "__main__":
    app = AnnuaireStudio()
    app.mainloop()
