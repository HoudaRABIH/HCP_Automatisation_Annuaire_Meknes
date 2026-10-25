
from pathlib import Path
import os, queue, re, subprocess, sys, threading, tkinter as tk
from tkinter import ttk, filedialog, messagebox

ROOT = Path(".")
BACKEND = ROOT / "scripts" / "generer_annuaire_tous_secteurs_v3.py"
FINAL = ROOT / "sorties" / "Annuaire_HCP_2025_FINAL.docx"

def count_excel(folder):
    p = Path(folder) if folder else None
    if not p or not p.exists(): return 0
    return len([f for f in p.glob("*.xlsx") if not f.name.startswith("~$")])

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Générateur Annuaire HCP")
        self.geometry("860x520")
        self.configure(bg="#f5f6f8")
        self.word = tk.StringVar()
        self.excel = tk.StringVar()
        self.status = tk.StringVar(value="En attente")
        self.q = queue.Queue()
        self.proc = None
        self._ui()
        self.after(100, self._poll)

    def _ui(self):
        tk.Label(self, text="Générateur Annuaire HCP", bg="#f5f6f8",
                 fg="#1f2937", font=("Segoe UI",22,"bold")).pack(anchor="w", padx=28, pady=(28,4))
        tk.Label(self, text="Choisissez le Word initial et le dossier Excel, puis lancez la génération.",
                 bg="#f5f6f8", fg="#6b7280", font=("Segoe UI",10)).pack(anchor="w", padx=28, pady=(0,20))

        card = tk.Frame(self, bg="white", highlightbackground="#d9dee7", highlightthickness=1)
        card.pack(fill="x", padx=28)
        card.columnconfigure(1, weight=1)

        tk.Label(card, text="1. Word initial", bg="white", font=("Segoe UI",11,"bold")).grid(
            row=0,column=0,columnspan=3,sticky="w",padx=20,pady=(20,8))
        ttk.Entry(card, textvariable=self.word).grid(row=1,column=0,columnspan=2,sticky="ew",padx=(20,10))
        ttk.Button(card,text="Choisir le Word",command=self._choose_word).grid(row=1,column=2,padx=(0,20))
        self.word_state = tk.Label(card,text="Word : NON SÉLECTIONNÉ",bg="white",fg="#b42318",
                                   font=("Segoe UI",9,"bold"))
        self.word_state.grid(row=2,column=0,columnspan=3,sticky="w",padx=20,pady=(7,18))

        ttk.Separator(card).grid(row=3,column=0,columnspan=3,sticky="ew",padx=20)

        tk.Label(card, text="2. Dossier des fichiers Excel", bg="white", font=("Segoe UI",11,"bold")).grid(
            row=4,column=0,columnspan=3,sticky="w",padx=20,pady=(18,8))
        ttk.Entry(card, textvariable=self.excel).grid(row=5,column=0,columnspan=2,sticky="ew",padx=(20,10))
        ttk.Button(card,text="Choisir le dossier",command=self._choose_excel).grid(row=5,column=2,padx=(0,20))
        self.excel_state = tk.Label(card,text="Excel : NON SÉLECTIONNÉ",bg="white",fg="#b42318",
                                    font=("Segoe UI",9,"bold"))
        self.excel_state.grid(row=6,column=0,columnspan=3,sticky="w",padx=20,pady=(7,20))

        self.run = ttk.Button(self,text="GÉNÉRER L'ANNUAIRE",command=self._start,state="disabled")
        self.run.pack(pady=(20,12))

        barbox = tk.Frame(self,bg="#f5f6f8")
        barbox.pack(fill="x",padx=28)
        tk.Label(barbox,textvariable=self.status,bg="#f5f6f8",font=("Segoe UI",9,"bold")).pack(anchor="w")
        self.bar = ttk.Progressbar(barbox,maximum=100)
        self.bar.pack(fill="x",pady=(6,8))

        self.open_btn = ttk.Button(self,text="Ouvrir le Word final",command=self._open_final,state="disabled")
        self.open_btn.pack()

    def _choose_word(self):
        p = filedialog.askopenfilename(title="Choisir le Word initial",filetypes=[("Word","*.docx")])
        if p:
            self.word.set(p)
            self._refresh()

    def _choose_excel(self):
        p = filedialog.askdirectory(title="Choisir le dossier des 23 fichiers Excel")
        if p:
            self.excel.set(p)
            self._refresh()

    def _refresh(self):
        wp = Path(self.word.get()) if self.word.get() else None
        word_ok = bool(wp and wp.exists() and wp.suffix.lower()==".docx")
        n = count_excel(self.excel.get())
        excel_ok = n == 23

        self.word_state.config(text="Word : PRÊT" if word_ok else "Word : NON SÉLECTIONNÉ",
                               fg="#16803f" if word_ok else "#b42318")
        self.excel_state.config(text=f"Excel : PRÊT • {n}/23" if excel_ok else f"Excel : {n}/23",
                                fg="#16803f" if excel_ok else "#b42318")

        ready = word_ok and excel_ok and BACKEND.exists() and self.proc is None
        self.run.config(state="normal" if ready else "disabled")
        self.open_btn.config(state="normal" if FINAL.exists() else "disabled")

    def _start(self):
        self.run.config(state="disabled")
        self.status.set("Démarrage...")
        self.bar["value"] = 1
        threading.Thread(target=self._worker,daemon=True).start()

    def _worker(self):
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"]="utf-8"
            env["PYTHONUTF8"]="1"
            flags = getattr(subprocess,"CREATE_NO_WINDOW",0) if os.name=="nt" else 0
            self.proc = subprocess.Popen(
                [sys.executable,"-u",str(BACKEND),"--word",self.word.get(),"--excel-dir",self.excel.get()],
                cwd=str(ROOT.resolve()),stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                text=True,encoding="utf-8",errors="replace",env=env,creationflags=flags)
            for line in self.proc.stdout:
                self.q.put(("line",line))
            self.q.put(("done",self.proc.wait()))
        except Exception as e:
            self.q.put(("error",str(e)))
        finally:
            self.proc = None

    def _poll(self):
        try:
            while True:
                kind,payload = self.q.get_nowait()
                if kind=="line":
                    m = re.search(r"\[(\d+)/16\]\s*(.+)",payload)
                    if m:
                        step = int(m.group(1)); name = m.group(2).strip()
                        self.bar["value"] = round(step/16*100)
                        self.status.set(f"Étape {step}/16 — {name}")
                elif kind=="done":
                    if payload==0 and FINAL.exists():
                        self.bar["value"]=100
                        self.status.set("Terminé avec succès")
                        self.open_btn.config(state="normal")
                        messagebox.showinfo("Succès","L'annuaire a été généré avec succès.")
                    else:
                        self.status.set("Échec")
                        messagebox.showerror("Échec","La génération a échoué.")
                    self._refresh()
                elif kind=="error":
                    self.status.set("Erreur")
                    messagebox.showerror("Erreur",payload)
                    self._refresh()
        except queue.Empty:
            pass
        self.after(100,self._poll)

    def _open_final(self):
        if not FINAL.exists():
            return
        if os.name=="nt": os.startfile(str(FINAL.resolve()))
        else: subprocess.Popen(["xdg-open",str(FINAL.resolve())])

if __name__=="__main__":
    App().mainloop()
