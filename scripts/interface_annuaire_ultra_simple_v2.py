from __future__ import annotations

from pathlib import Path
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

ROOT = Path(".")
BACKEND = ROOT / "scripts" / "generer_annuaire_tous_secteurs_v3.py"
BACKEND_FINAL = ROOT / "sorties" / "Annuaire_HCP_2025_FINAL.docx"
REPORT = ROOT / "sorties" / "rapport_generation_tous_secteurs.txt"

STEPS = [
    "Transport", "Prix", "Environnement", "Tourisme",
    "Agriculture", "Santé", "Enseignement", "Ens. supérieur",
    "Emploi", "Justice", "Culture", "Habitat / Foncier",
    "Population", "Industrie", "Énergie / Eau", "Validation finale"
]


def excel_count(folder: Path | None) -> int:
    if not folder or not folder.exists() or not folder.is_dir():
        return 0
    return len([p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")])


def open_path(path: Path):
    if not path.exists():
        raise FileNotFoundError(str(path))
    if os.name == "nt":
        os.startfile(str(path.resolve()))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path.resolve())])
    else:
        subprocess.Popen(["xdg-open", str(path.resolve())])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Générateur Annuaire HCP — Simple SAFE")
        self.geometry("1180x820")
        self.minsize(980, 720)

        self.word_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str((ROOT / "sorties" / "Annuaire_HCP_2025_FINAL_UTILISATEUR.docx").resolve()))
        self.status_var = tk.StringVar(value="Prêt à configurer")
        self.step_var = tk.StringVar(value="En attente")
        self.percent_var = tk.StringVar(value="0%")

        self.proc = None
        self.q: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_ui()
        self._refresh_state()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        self.configure(bg="#f5f6f8")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)

        title = tk.Frame(self, bg="#f5f6f8")
        title.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 8))
        title.columnconfigure(0, weight=1)

        tk.Label(title, text="Générateur complet de l'Annuaire HCP", font=("Segoe UI", 22, "bold"), bg="#f5f6f8").grid(row=0, column=0, sticky="w")
        tk.Label(title, text="Choisir simplement : le modèle Word, le dossier Excel, le fichier de sortie, puis lancer le pipeline.", font=("Segoe UI", 10), fg="#555", bg="#f5f6f8").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.status_badge = tk.Label(title, textvariable=self.status_var, font=("Segoe UI", 10, "bold"), bg="#e9f6ee", fg="#137333", padx=12, pady=6)
        self.status_badge.grid(row=0, column=1, rowspan=2, sticky="e")

        cfg = ttk.LabelFrame(self, text="Configuration")
        cfg.grid(row=1, column=0, sticky="ew", padx=24, pady=8)
        cfg.columnconfigure(1, weight=1)

        # Word
        ttk.Label(cfg, text="Modèle Word initial :").grid(row=0, column=0, sticky="w", padx=12, pady=10)
        self.word_entry = ttk.Entry(cfg, textvariable=self.word_var)
        self.word_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=10)
        ttk.Button(cfg, text="Choisir...", command=self._choose_word).grid(row=0, column=2, padx=12, pady=10)

        # Excel folder
        ttk.Label(cfg, text="Dossier des fichiers Excel :").grid(row=1, column=0, sticky="w", padx=12, pady=10)
        self.excel_entry = ttk.Entry(cfg, textvariable=self.excel_var)
        self.excel_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=10)
        ttk.Button(cfg, text="Choisir...", command=self._choose_excel_folder).grid(row=1, column=2, padx=12, pady=10)

        # Output
        ttk.Label(cfg, text="Fichier de sortie final :").grid(row=2, column=0, sticky="w", padx=12, pady=10)
        self.output_entry = ttk.Entry(cfg, textvariable=self.output_var)
        self.output_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=10)
        ttk.Button(cfg, text="Choisir...", command=self._choose_output).grid(row=2, column=2, padx=12, pady=10)

        info = tk.Frame(self, bg="#f5f6f8")
        info.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 8))
        for i in range(3):
            info.columnconfigure(i, weight=1)

        self.word_info = self._small_card(info, 0, "Word")
        self.excel_info = self._small_card(info, 1, "Excel")
        self.output_info = self._small_card(info, 2, "Sortie")

        pipe = ttk.LabelFrame(self, text="Pipeline")
        pipe.grid(row=3, column=0, sticky="ew", padx=24, pady=8)
        pipe.columnconfigure(0, weight=1)

        top = tk.Frame(pipe)
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        top.columnconfigure(0, weight=1)

        tk.Label(top, textvariable=self.step_var, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(top, textvariable=self.percent_var, font=("Segoe UI", 10, "bold"), fg="#9e163e").grid(row=0, column=1, sticky="e")

        self.progress = ttk.Progressbar(pipe, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew", padx=12, pady=(4, 10))

        steps_frame = tk.Frame(pipe)
        steps_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        for c in range(4):
            steps_frame.columnconfigure(c, weight=1)
        self.step_labels = []
        for i, name in enumerate(STEPS, start=1):
            r = (i - 1) // 4
            c = (i - 1) % 4
            lbl = tk.Label(steps_frame, text=f"{i}. {name}", anchor="w", justify="left", fg="#666")
            lbl.grid(row=r, column=c, sticky="ew", padx=8, pady=4)
            self.step_labels.append(lbl)

        actions = tk.Frame(self, bg="#f5f6f8")
        actions.grid(row=4, column=0, sticky="nsew", padx=24, pady=(8, 24))
        actions.columnconfigure(0, weight=3)
        actions.columnconfigure(1, weight=1)
        actions.rowconfigure(0, weight=1)

        journal_box = ttk.LabelFrame(actions, text="Journal d'exécution")
        journal_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        journal_box.rowconfigure(0, weight=1)
        journal_box.columnconfigure(0, weight=1)

        self.log = tk.Text(journal_box, font=("Consolas", 10), bg="#0e1525", fg="#dbe7ff", insertbackground="white")
        self.log.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.log.configure(state="disabled")

        right = ttk.LabelFrame(actions, text="Actions")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self.run_btn = ttk.Button(right, text="Lancer le pipeline", command=self._start, state="disabled")
        self.run_btn.pack(fill="x", padx=12, pady=(12, 8))

        self.stop_btn = ttk.Button(right, text="Arrêter", command=self._stop, state="disabled")
        self.stop_btn.pack(fill="x", padx=12, pady=8)

        ttk.Separator(right).pack(fill="x", padx=12, pady=10)

        ttk.Button(right, text="Ouvrir le Word final", command=lambda: self._open_selected_output()).pack(fill="x", padx=12, pady=6)
        ttk.Button(right, text="Ouvrir le rapport", command=lambda: self._safe_open(REPORT)).pack(fill="x", padx=12, pady=6)
        ttk.Button(right, text="Ouvrir le dossier sorties", command=lambda: self._safe_open(ROOT / "sorties")).pack(fill="x", padx=12, pady=6)

        tip = tk.Label(right, text="Le pipeline met à jour tous les secteurs automatiquement après le choix des 3 entrées.", wraplength=220, justify="left", fg="#555")
        tip.pack(anchor="w", padx=12, pady=(16, 8))

    def _small_card(self, parent, col, title):
        frame = ttk.LabelFrame(parent, text=title)
        frame.grid(row=0, column=col, sticky="ew", padx=6)
        value = tk.Label(frame, text="—", font=("Segoe UI", 13, "bold"))
        value.pack(anchor="w", padx=12, pady=(10, 2))
        hint = tk.Label(frame, text="", fg="#666")
        hint.pack(anchor="w", padx=12, pady=(0, 10))
        return value, hint

    def _choose_word(self):
        path = filedialog.askopenfilename(title="Choisir le modèle Word", filetypes=[("Word", "*.docx")])
        if path:
            self.word_var.set(path)
            self._refresh_state()

    def _choose_excel_folder(self):
        path = filedialog.askdirectory(title="Choisir le dossier des Excel")
        if path:
            self.excel_var.set(path)
            self._refresh_state()

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Choisir le fichier de sortie",
            defaultextension=".docx",
            filetypes=[("Word", "*.docx")],
            initialfile="Annuaire_HCP_2025_FINAL.docx"
        )
        if path:
            self.output_var.set(path)
            self._refresh_state()

    def _refresh_state(self):
        word = Path(self.word_var.get()) if self.word_var.get() else None
        excel_dir = Path(self.excel_var.get()) if self.excel_var.get() else None
        output = Path(self.output_var.get()) if self.output_var.get() else None

        word_ok = bool(word and word.exists() and word.suffix.lower() == ".docx")
        excel_n = excel_count(excel_dir)
        excel_ok = bool(excel_dir and excel_dir.exists() and excel_n == 23)
        output_ok = bool(output and output.suffix.lower() == ".docx")

        self.word_info[0].configure(text="OK" if word_ok else "Non choisi", fg="#14804A" if word_ok else "#b42318")
        self.word_info[1].configure(text=word.name if word_ok else "Choisir un fichier .docx")

        self.excel_info[0].configure(text=f"{excel_n}/23", fg="#14804A" if excel_ok else "#b42318")
        self.excel_info[1].configure(text=str(excel_dir) if excel_dir else "Choisir un dossier")

        self.output_info[0].configure(text="OK" if output_ok else "Non choisi", fg="#14804A" if output_ok else "#b42318")
        self.output_info[1].configure(text=output.name if output_ok else "Choisir un fichier .docx")

        ready = word_ok and excel_ok and output_ok and BACKEND.exists()
        self.run_btn.configure(state="normal" if ready else "disabled")

        if ready:
            self.status_var.set("PRÊT")
            self.status_badge.configure(bg="#e9f6ee", fg="#137333")
        else:
            self.status_var.set("CONFIGURATION")
            self.status_badge.configure(bg="#eef2f6", fg="#42526b")

    def _set_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _start(self):
        word = Path(self.word_var.get())
        excel_dir = Path(self.excel_var.get())
        output = Path(self.output_var.get())

        if not BACKEND.exists():
            messagebox.showerror("Erreur", f"Backend introuvable :\n{BACKEND}")
            return
        if not word.exists():
            messagebox.showerror("Erreur", "Le modèle Word n'existe pas.")
            return
        n = excel_count(excel_dir)
        if n != 23:
            messagebox.showerror("Erreur", f"Le dossier Excel doit contenir exactement 23 fichiers .xlsx.\nDétectés : {n}")
            return
        if output.suffix.lower() != ".docx":
            messagebox.showerror("Erreur", "Le fichier de sortie doit être un .docx")
            return

        output.parent.mkdir(parents=True, exist_ok=True)

        self._clear_log()
        self._reset_pipeline()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("EN COURS")
        self.status_badge.configure(bg="#fff4e5", fg="#9a6700")
        self.step_var.set("Lancement du pipeline...")
        self.percent_var.set("0%")
        self.progress["value"] = 0

        threading.Thread(target=self._worker, args=(word, excel_dir, output), daemon=True).start()

    def _worker(self, word: Path, excel_dir: Path, output: Path):
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0

            self.proc = subprocess.Popen(
                [sys.executable, "-u", str(BACKEND), "--word", str(word), "--excel-dir", str(excel_dir)],
                cwd=str(ROOT.resolve()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                creationflags=flags,
            )

            assert self.proc.stdout is not None
            for line in self.proc.stdout:
                self.q.put(("line", line))

            code = self.proc.wait()
            if code == 0:
                if not BACKEND_FINAL.exists():
                    self.q.put(("error", "Le fichier final standard n'a pas été trouvé après génération."))
                    return
                shutil.copy2(BACKEND_FINAL, output)
                self.q.put(("copied", str(output)))
            self.q.put(("done", code))
        except Exception as e:
            self.q.put(("error", str(e)))
        finally:
            self.proc = None

    def _reset_pipeline(self):
        for i, lbl in enumerate(self.step_labels, start=1):
            lbl.configure(fg="#666", font=("Segoe UI", 9))

    def _mark_step(self, current: int):
        for i, lbl in enumerate(self.step_labels, start=1):
            if i < current:
                lbl.configure(fg="#14804A", font=("Segoe UI", 9, "bold"))
            elif i == current:
                lbl.configure(fg="#9e163e", font=("Segoe UI", 9, "bold"))
            else:
                lbl.configure(fg="#666", font=("Segoe UI", 9))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()

                if kind == "line":
                    line = str(payload)
                    self._set_log(line)
                    m = re.search(r"\[(\d+)/16\]\s*(.+)", line)
                    if m:
                        idx = int(m.group(1))
                        name = m.group(2).strip()
                        pct = round((idx - 1) / 16 * 100)
                        self.step_var.set(f"Étape {idx}/16 — {name}")
                        self.percent_var.set(f"{pct}%")
                        self.progress["value"] = pct
                        self._mark_step(idx)

                elif kind == "copied":
                    self._set_log(f"\nCopie finale utilisateur créée : {payload}\n")

                elif kind == "done":
                    code = int(payload)
                    self.stop_btn.configure(state="disabled")
                    self._refresh_state()
                    if code == 0:
                        self.progress["value"] = 100
                        self.percent_var.set("100%")
                        self.step_var.set("Pipeline terminé")
                        self.status_var.set("TERMINÉ")
                        self.status_badge.configure(bg="#e9f6ee", fg="#137333")
                        for lbl in self.step_labels:
                            lbl.configure(fg="#14804A", font=("Segoe UI", 9, "bold"))
                        messagebox.showinfo("Succès", f"L'annuaire a été généré avec succès.\n\nFichier final :\n{self.output_var.get()}")
                    else:
                        self.status_var.set("ÉCHEC")
                        self.status_badge.configure(bg="#fdecec", fg="#b42318")
                        self.run_btn.configure(state="normal")
                        messagebox.showerror("Échec", "Le pipeline a échoué. Consulte le journal d'exécution.")

                elif kind == "error":
                    self.stop_btn.configure(state="disabled")
                    self.run_btn.configure(state="normal")
                    self.status_var.set("ERREUR")
                    self.status_badge.configure(bg="#fdecec", fg="#b42318")
                    messagebox.showerror("Erreur", str(payload))

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _stop(self):
        if self.proc and messagebox.askyesno("Arrêter", "Voulez-vous arrêter le pipeline en cours ?"):
            self.proc.terminate()

    def _safe_open(self, path: Path):
        try:
            open_path(path)
        except Exception as e:
            messagebox.showerror("Ouverture impossible", str(e))

    def _open_selected_output(self):
        path = Path(self.output_var.get()) if self.output_var.get() else None
        if not path:
            messagebox.showerror("Erreur", "Aucun fichier de sortie n'est défini.")
            return
        self._safe_open(path)


if __name__ == "__main__":
    App().mainloop()
