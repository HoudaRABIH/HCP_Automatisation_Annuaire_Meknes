from pathlib import Path
import subprocess
import sys
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_TITLE = "Générateur Annuaire HCP"
ROOT = Path(".")
SCRIPTS = ROOT / "scripts"
BACKEND = SCRIPTS / "generer_annuaire_complet.py"
FINAL = ROOT / "sorties" / "Annuaire_HCP_2025_FINAL.docx"
LOG = ROOT / "sorties" / "rapport_generation_complete.txt"


class AnnuaireGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("940x650")
        self.minsize(820, 560)

        self.proc = None
        self.q = queue.Queue()

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)

        title = ttk.Label(
            main,
            text="Générateur Annuaire HCP",
            font=("Segoe UI", 18, "bold")
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            main,
            text="Génération automatisée du document statistique final",
            font=("Segoe UI", 10)
        )
        subtitle.pack(anchor="w", pady=(2, 16))

        info = ttk.LabelFrame(main, text="Projet", padding=12)
        info.pack(fill="x")

        ttk.Label(info, text="Dossier projet :").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )

        self.project_var = tk.StringVar(value=str(ROOT.resolve()))
        project_entry = ttk.Entry(
            info,
            textvariable=self.project_var,
            state="readonly"
        )
        project_entry.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Button(
            info,
            text="Ouvrir le dossier",
            command=self._open_project
        ).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(info, text="Backend :").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4
        )

        self.backend_var = tk.StringVar(value=str(BACKEND))
        ttk.Entry(
            info,
            textvariable=self.backend_var,
            state="readonly"
        ).grid(row=1, column=1, columnspan=2, sticky="ew", pady=4)

        info.columnconfigure(1, weight=1)

        controls = ttk.Frame(main)
        controls.pack(fill="x", pady=(16, 10))

        self.run_btn = ttk.Button(
            controls,
            text="Générer l'annuaire",
            command=self._start_generation
        )
        self.run_btn.pack(side="left")

        self.cancel_btn = ttk.Button(
            controls,
            text="Arrêter",
            command=self._cancel_generation,
            state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=(8, 0))

        self.open_final_btn = ttk.Button(
            controls,
            text="Ouvrir le Word final",
            command=self._open_final,
            state="disabled"
        )
        self.open_final_btn.pack(side="right")

        self.progress = ttk.Progressbar(
            main,
            mode="indeterminate"
        )
        self.progress.pack(fill="x", pady=(0, 10))

        status_frame = ttk.Frame(main)
        status_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(status_frame, text="Statut :").pack(side="left")
        self.status_var = tk.StringVar(value="Prêt")
        ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=(6, 0))

        log_frame = ttk.LabelFrame(main, text="Journal d'exécution", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            font=("Consolas", 9),
            state="disabled"
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(main)
        footer.pack(fill="x", pady=(10, 0))

        ttk.Button(
            footer,
            text="Ouvrir le rapport",
            command=self._open_log
        ).pack(side="left")

        ttk.Button(
            footer,
            text="Quitter",
            command=self.destroy
        ).pack(side="right")

    def _append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_path(self, path: Path):
        if not path.exists():
            messagebox.showwarning(
                "Fichier introuvable",
                f"Le chemin n'existe pas :\n{path}"
            )
            return

        try:
            if os.name == "nt":
                os.startfile(str(path.resolve()))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path.resolve())])
            else:
                subprocess.Popen(["xdg-open", str(path.resolve())])
        except Exception as exc:
            messagebox.showerror("Erreur", str(exc))

    def _open_project(self):
        self._open_path(ROOT)

    def _open_final(self):
        self._open_path(FINAL)

    def _open_log(self):
        self._open_path(LOG)

    def _preflight(self):
        missing = []

        if not BACKEND.exists():
            missing.append(str(BACKEND))

        base = ROOT / "sorties" / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"
        if not base.exists():
            missing.append(str(base))

        data_dir = ROOT / "donnees_brutes"
        if not data_dir.exists():
            missing.append(str(data_dir))

        if missing:
            raise FileNotFoundError(
                "Fichiers/dossiers requis introuvables :\n\n"
                + "\n".join(missing)
            )

    def _start_generation(self):
        if self.proc is not None:
            return

        try:
            self._preflight()
        except Exception as exc:
            messagebox.showerror("Pré-contrôle", str(exc))
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.status_var.set("Génération en cours...")
        self.run_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.open_final_btn.configure(state="disabled")
        self.progress.start(12)

        t = threading.Thread(
            target=self._worker,
            daemon=True
        )
        t.start()

    def _worker(self):
        try:
            env = os.environ.copy()
            # Force UTF-8 so accents render correctly in the GUI log.
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
                [sys.executable, str(BACKEND)],
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
                self.q.put(("log", line))

            code = self.proc.wait()

            if code == 0 and FINAL.exists():
                self.q.put(("done", 0))
            else:
                self.q.put(("done", code))

        except Exception as exc:
            self.q.put(("error", str(exc)))
        finally:
            self.proc = None

    def _cancel_generation(self):
        if self.proc is None:
            return

        if not messagebox.askyesno(
            "Arrêter",
            "Voulez-vous arrêter la génération en cours ?"
        ):
            return

        try:
            self.proc.terminate()
            self.status_var.set("Arrêt demandé...")
        except Exception as exc:
            messagebox.showerror("Erreur", str(exc))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()

                if kind == "log":
                    self._append_log(payload)

                elif kind == "done":
                    self.progress.stop()
                    self.run_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")

                    if payload == 0:
                        self.status_var.set("Terminé avec succès")
                        self.open_final_btn.configure(state="normal")
                        messagebox.showinfo(
                            "Génération terminée",
                            "Annuaire généré avec succès.\n\n"
                            f"{FINAL}"
                        )
                    else:
                        self.status_var.set("Échec")
                        messagebox.showerror(
                            "Génération échouée",
                            "Une étape a échoué.\n"
                            "Consultez le journal affiché et le rapport."
                        )

                elif kind == "error":
                    self.progress.stop()
                    self.run_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")
                    self.status_var.set("Erreur")
                    messagebox.showerror("Erreur", payload)

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)


if __name__ == "__main__":
    app = AnnuaireGUI()
    app.mainloop()
