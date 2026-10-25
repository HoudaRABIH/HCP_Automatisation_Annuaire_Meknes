from pathlib import Path
import shutil
import sys

ROOT = Path(".")
TARGET = ROOT / "modeles" / "modeles.docx"
BACKUPS = ROOT / "sorties" / "backups"

print("=" * 90)
print("RECUPERATION DU MODELE WORD")
print("=" * 90)

if TARGET.exists():
    print(f"Le modèle existe déjà : {TARGET}")
    sys.exit(0)

candidates = []
if BACKUPS.exists():
    candidates += list(BACKUPS.glob("modele_original_*.docx"))
    candidates += list(BACKUPS.glob("*modele*.docx"))

unique = {}
for p in candidates:
    try:
        unique[p.resolve()] = p
    except Exception:
        unique[p] = p

candidates = sorted(unique.values(), key=lambda p: p.stat().st_mtime, reverse=True)

if candidates:
    source = candidates[0]
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, TARGET)
    print(f"Sauvegarde trouvée : {source}")
    print(f"Modèle restauré : {TARGET}")
    print("RECUPERATION : OK")
    sys.exit(0)

print("Aucune sauvegarde modele_original_*.docx trouvée dans sorties\\backups.")
print("Le script n'utilisera PAS un Word intermédiaire/final à la place du modèle original.")
sys.exit(2)
