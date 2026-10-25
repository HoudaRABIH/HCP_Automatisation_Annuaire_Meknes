from __future__ import annotations
from pathlib import Path
import os, queue, re, subprocess, sys, threading, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

APP_TITLE = "HCP Annuaire Studio"
ROOT = Path(".")
BACKEND = ROOT / "scripts" / "generer_annuaire_tous_secteurs.py"
DATA_DIR = ROOT / "donnees_brutes"
FINAL = ROOT / "sorties" / "Annuaire_HCP_2025_FINAL.docx"
REPORT = ROOT / "sorties" / "rapport_generation_tous_secteurs.txt"

C = {
    "bg":"#F5F7FA","surface":"#FFFFFF","border":"#E3E8EF",
    "text":"#17202A","muted":"#6B7280","primary":"#8B1538",
    "primary_dark":"#68102A","success":"#15803D","success_soft":"#EAF7EF",
    "danger":"#B42318","danger_soft":"#FDECEC","info":"#2563EB",
    "info_soft":"#EEF4FF","sidebar":"#151B26","sidebar_muted":"#AEB8C7"
}

STEPS = [
    "Transport","Prix","Environnement","Tourisme",
    "Agriculture / Élevage / Forêt","Santé","Enseignement",
    "Enseignement supérieur","Emploi / Salaires","Justice",
    "Culture / Loisirs","Habitat / Construction / Foncier",
    "Population / Division administrative","Industrie / Artisanat",
    "Énergie / Eau","Validation finale"
]

def count_excels():
    return len([p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$")]) if DATA_DIR.exists() else 0

def open_path(p):
    if not p.exists(): raise FileNotFoundError(str(p))
    if os.name == "nt": os.startfile(str(p.resolve()))
    elif sys.platform == "darwin": subprocess.Popen(["open", str(p.resolve())])
    else: subprocess.Popen(["xdg-open", str(p.resolve())])

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x880")
        self.minsize(1160,760)
        self.configure(bg=C["bg"])
        self.word = tk.StringVar()
        self.proc = None
        self.q = queue.Queue()
        self._style()
        self._build()
        self._refresh()
        self.after(100,self._poll)

    def _style(self):
        s=ttk.Style(self)
        try:s.theme_use("clam")
        except:pass
        s.configure("Primary.TButton",font=("Segoe UI",10,"bold"),padding=(18,10),
                    foreground="white",background=C["primary"],borderwidth=0)
        s.map("Primary.TButton",background=[("active",C["primary_dark"])])
        s.configure("Secondary.TButton",font=("Segoe UI",9),padding=(12,8))
        s.configure("P.Horizontal.TProgressbar",troughcolor="#E9EDF2",
                    background=C["primary"],thickness=10)

    def _build(self):
        self.columnconfigure(1,weight=1); self.rowconfigure(0,weight=1)
        side=tk.Frame(self,bg=C["sidebar"],width=245); side.grid(row=0,column=0,sticky="nsew"); side.grid_propagate(False)
        tk.Label(side,text="HCP",bg=C["primary"],fg="white",font=("Segoe UI",11,"bold"),
                 padx=11,pady=11).pack(anchor="w",padx=24,pady=(30,12))
        tk.Label(side,text="Annuaire Studio",bg=C["sidebar"],fg="white",
                 font=("Segoe UI",15,"bold")).pack(anchor="w",padx=24)
        tk.Label(side,text="Génération complète",bg=C["sidebar"],fg=C["sidebar_muted"],
                 font=("Segoe UI",9)).pack(anchor="w",padx=24,pady=(3,28))
        tk.Label(side,text="▦   Tous les secteurs",bg=C["primary"],fg="white",
                 font=("Segoe UI",10,"bold"),anchor="w",padx=20,pady=12).pack(fill="x",padx=10)
        tk.Label(side,text="1 Word initial\n23 Excel automatiques\n16 étapes contrôlées",
                 bg=C["sidebar"],fg=C["sidebar_muted"],justify="left",
                 font=("Segoe UI",9)).pack(anchor="w",padx=24,pady=24)

        main=tk.Frame(self,bg=C["bg"]); main.grid(row=0,column=1,sticky="nsew")
        main.columnconfigure(0,weight=1); main.rowconfigure(4,weight=1)

        h=tk.Frame(main,bg=C["bg"]); h.grid(row=0,column=0,sticky="ew",padx=30,pady=(24,12)); h.columnconfigure(0,weight=1)
        tk.Label(h,text="Générateur complet de l'Annuaire HCP",bg=C["bg"],fg=C["text"],
                 font=("Segoe UI",22,"bold")).grid(row=0,column=0,sticky="w")
        tk.Label(h,text="Le Word sélectionné est actualisé secteur par secteur avec les fichiers Excel du projet.",
                 bg=C["bg"],fg=C["muted"],font=("Segoe UI",10)).grid(row=1,column=0,sticky="w",pady=(4,0))
        self.badge=tk.Label(h,text="●  PRÊT",bg=C["success_soft"],fg=C["success"],
                            font=("Segoe UI",9,"bold"),padx=12,pady=7)
        self.badge.grid(row=0,column=1,rowspan=2,sticky="e")

        card=tk.Frame(main,bg="white",highlightbackground=C["border"],highlightthickness=1)
        card.grid(row=1,column=0,sticky="ew",padx=30,pady=(0,12)); card.columnconfigure(0,weight=1)
        tk.Label(card,text="1. Choisir le Word initial",bg="white",fg=C["text"],
                 font=("Segoe UI",11,"bold")).grid(row=0,column=0,columnspan=2,sticky="w",padx=18,pady=(14,3))
        tk.Label(card,text="Le programme modifiera ensuite automatiquement TOUS les secteurs.",
                 bg="white",fg=C["muted"],font=("Segoe UI",8)).grid(row=1,column=0,columnspan=2,sticky="w",padx=18,pady=(0,10))
        ttk.Entry(card,textvariable=self.word,font=("Segoe UI",9)).grid(row=2,column=0,sticky="ew",padx=(18,8),pady=(0,14))
        ttk.Button(card,text="Parcourir...",style="Secondary.TButton",command=self._choose).grid(row=2,column=1,padx=(0,18),pady=(0,14))

        stats=tk.Frame(main,bg=C["bg"]); stats.grid(row=2,column=0,sticky="ew",padx=30,pady=(0,12))
        stats.columnconfigure(0,weight=1);stats.columnconfigure(1,weight=1);stats.columnconfigure(2,weight=1)
        self.excel=self._stat(stats,0,"EXCEL DÉTECTÉS","—","Automatique")
        self.wstat=self._stat(stats,1,"WORD INITIAL","NON","Sélection utilisateur")
        self.fstat=self._stat(stats,2,"DERNIER FINAL","—","Document généré")

        pipe=tk.Frame(main,bg="white",highlightbackground=C["border"],highlightthickness=1)
        pipe.grid(row=3,column=0,sticky="ew",padx=30,pady=(0,12));pipe.columnconfigure(0,weight=1)
        top=tk.Frame(pipe,bg="white");top.grid(row=0,column=0,sticky="ew",padx=18,pady=(14,8));top.columnconfigure(0,weight=1)
        tk.Label(top,text="Pipeline — tous les secteurs",bg="white",fg=C["text"],
                 font=("Segoe UI",11,"bold")).grid(row=0,column=0,sticky="w")
        self.run=ttk.Button(top,text="▶  Générer tout l'annuaire",style="Primary.TButton",command=self._start)
        self.run.grid(row=0,column=1,rowspan=2,sticky="e")
        self.step_label=tk.Label(top,text="En attente",bg="white",fg=C["muted"],font=("Segoe UI",8))
        self.step_label.grid(row=1,column=0,sticky="w",pady=(3,0))
        self.pct=tk.Label(pipe,text="0%",bg="white",fg=C["muted"],font=("Segoe UI",9,"bold"))
        self.pct.grid(row=1,column=0,sticky="e",padx=18)
        self.pvar=tk.DoubleVar(value=0)
        ttk.Progressbar(pipe,variable=self.pvar,maximum=100,style="P.Horizontal.TProgressbar").grid(row=2,column=0,sticky="ew",padx=18,pady=(2,10))

        grid=tk.Frame(pipe,bg="white");grid.grid(row=3,column=0,sticky="ew",padx=18,pady=(0,14))
        for c in range(8):grid.columnconfigure(c,weight=1,uniform="st")
        self.sw=[]
        for i,name in enumerate(STEPS,1):
            r=(i-1)//8;c=(i-1)%8
            cell=tk.Frame(grid,bg="white");cell.grid(row=r,column=c,sticky="nsew",padx=3,pady=5)
            dot=tk.Label(cell,text=str(i),bg="#E9EDF2",fg=C["muted"],width=3,font=("Segoe UI",8,"bold"))
            dot.pack()
            lab=tk.Label(cell,text=name,bg="white",fg=C["muted"],font=("Segoe UI",7),wraplength=120,justify="center")
            lab.pack(pady=(3,0))
            self.sw.append((dot,lab))

        bottom=tk.Frame(main,bg=C["bg"]);bottom.grid(row=4,column=0,sticky="nsew",padx=30,pady=(0,24))
        bottom.columnconfigure(0,weight=3);bottom.columnconfigure(1,weight=1);bottom.rowconfigure(0,weight=1)
        lp=tk.Frame(bottom,bg="white",highlightbackground=C["border"],highlightthickness=1)
        lp.grid(row=0,column=0,sticky="nsew",padx=(0,7));lp.rowconfigure(1,weight=1);lp.columnconfigure(0,weight=1)
        lh=tk.Frame(lp,bg="white");lh.grid(row=0,column=0,sticky="ew",padx=14,pady=10);lh.columnconfigure(0,weight=1)
        tk.Label(lh,text="Journal d'exécution",bg="white",fg=C["text"],font=("Segoe UI",10,"bold")).grid(row=0,column=0,sticky="w")
        self.stop=ttk.Button(lh,text="Arrêter",state="disabled",command=self._stop);self.stop.grid(row=0,column=1)
        self.log=tk.Text(lp,bg="#0F172A",fg="#DCE7F3",relief="flat",bd=0,font=("Cascadia Mono",9),
                         padx=10,pady=10,state="disabled")
        self.log.grid(row=1,column=0,sticky="nsew",padx=14,pady=(0,14))

        rp=tk.Frame(bottom,bg="white",highlightbackground=C["border"],highlightthickness=1)
        rp.grid(row=0,column=1,sticky="nsew",padx=(7,0))
        tk.Label(rp,text="Résultat",bg="white",fg=C["text"],font=("Segoe UI",10,"bold")).pack(anchor="w",padx=16,pady=(15,8))
        ttk.Button(rp,text="Ouvrir le Word final",command=lambda:self._open(FINAL)).pack(fill="x",padx=16,pady=5)
        ttk.Button(rp,text="Ouvrir le rapport",command=lambda:self._open(REPORT)).pack(fill="x",padx=16,pady=5)
        ttk.Button(rp,text="Ouvrir sorties",command=lambda:self._open(ROOT/"sorties")).pack(fill="x",padx=16,pady=5)

    def _stat(self,p,c,t,v,h):
        x=tk.Frame(p,bg="white",highlightbackground=C["border"],highlightthickness=1)
        x.grid(row=0,column=c,sticky="nsew",padx=(0 if c==0 else 6,0 if c==2 else 6))
        tk.Label(x,text=t,bg="white",fg=C["muted"],font=("Segoe UI",8,"bold")).pack(anchor="w",padx=16,pady=(12,0))
        val=tk.Label(x,text=v,bg="white",fg=C["text"],font=("Segoe UI",16,"bold"));val.pack(anchor="w",padx=16,pady=(4,1))
        tk.Label(x,text=h,bg="white",fg=C["muted"],font=("Segoe UI",8)).pack(anchor="w",padx=16,pady=(0,12))
        return val

    def _choose(self):
        p=filedialog.askopenfilename(title="Choisir le Word initial",filetypes=[("Word","*.docx")])
        if p:self.word.set(p);self._refresh()

    def _refresh(self):
        n=count_excels();self.excel.configure(text=str(n),fg=C["success"] if n>=20 else C["danger"])
        p=Path(self.word.get()) if self.word.get() else None
        self.wstat.configure(text="OK" if p and p.exists() else "NON",fg=C["success"] if p and p.exists() else C["danger"])
        self.fstat.configure(text=f"{FINAL.stat().st_size/(1024*1024):.2f} Mo" if FINAL.exists() else "—",
                             fg=C["success"] if FINAL.exists() else C["text"])

    def _start(self):
        p=Path(self.word.get()) if self.word.get() else None
        if not p or not p.exists():
            messagebox.showwarning("Word requis","Choisissez le Word initial.")
            return
        if count_excels()<20:
            messagebox.showerror("Excel","Le dossier donnees_brutes est incomplet.")
            return
        if not BACKEND.exists():
            messagebox.showerror("Backend",f"Introuvable : {BACKEND}")
            return
        self._clear();self.run.configure(state="disabled");self.stop.configure(state="normal")
        self.badge.configure(text="●  EN COURS",bg=C["info_soft"],fg=C["info"])
        self.pvar.set(1);self.pct.configure(text="1%")
        threading.Thread(target=self._worker,args=(p,),daemon=True).start()

    def _worker(self,p):
        try:
            env=os.environ.copy();env["PYTHONIOENCODING"]="utf-8";env["PYTHONUTF8"]="1"
            flags=getattr(subprocess,"CREATE_NO_WINDOW",0) if os.name=="nt" else 0
            self.proc=subprocess.Popen([sys.executable,"-u",str(BACKEND),"--word",str(p)],
                cwd=str(ROOT.resolve()),stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                text=True,encoding="utf-8",errors="replace",bufsize=1,env=env,creationflags=flags)
            for line in self.proc.stdout:self.q.put(("line",line))
            self.q.put(("done",self.proc.wait()))
        except Exception as e:self.q.put(("error",str(e)))
        finally:self.proc=None

    def _line(self,line):
        self._append(line)
        m=re.search(r"\[(\d+)/16\]\s*(.+)",line)
        if m:
            k=int(m.group(1));name=m.group(2).strip();pct=round(k/16*100)
            self.pvar.set(pct);self.pct.configure(text=f"{pct}%");self.step_label.configure(text=f"Étape {k}/16 • {name}")
            for i,(d,l) in enumerate(self.sw,1):
                if i<k:d.configure(text="✓",bg=C["success_soft"],fg=C["success"]);l.configure(fg=C["success"])
                elif i==k:d.configure(text=str(i),bg=C["primary"],fg="white");l.configure(fg=C["primary"])
                else:d.configure(text="○",bg="#E9EDF2",fg=C["muted"]);l.configure(fg=C["muted"])

    def _poll(self):
        try:
            while True:
                k,p=self.q.get_nowait()
                if k=="line":self._line(p)
                elif k=="done":
                    self.run.configure(state="normal");self.stop.configure(state="disabled")
                    if p==0 and FINAL.exists():
                        self.pvar.set(100);self.pct.configure(text="100%");self.badge.configure(text="●  SUCCÈS",bg=C["success_soft"],fg=C["success"])
                        for d,l in self.sw:d.configure(text="✓",bg=C["success_soft"],fg=C["success"]);l.configure(fg=C["success"])
                        self._refresh();messagebox.showinfo("Terminé","Tous les secteurs ont été traités.")
                    else:
                        self.badge.configure(text="●  ÉCHEC",bg=C["danger_soft"],fg=C["danger"])
                        messagebox.showerror("Échec","Une étape a échoué. Consultez le journal.")
                elif k=="error":
                    self.run.configure(state="normal");self.stop.configure(state="disabled")
                    messagebox.showerror("Erreur",p)
        except queue.Empty:pass
        self.after(100,self._poll)

    def _append(self,t):
        self.log.configure(state="normal");self.log.insert("end",t);self.log.see("end");self.log.configure(state="disabled")
    def _clear(self):
        self.log.configure(state="normal");self.log.delete("1.0","end");self.log.configure(state="disabled")
    def _stop(self):
        if self.proc and messagebox.askyesno("Arrêter","Arrêter la génération ?"):self.proc.terminate()
    def _open(self,p):
        try:open_path(p)
        except Exception as e:messagebox.showerror("Ouverture",str(e))

if __name__=="__main__":App().mainloop()
