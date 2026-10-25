from pathlib import Path
import pandas as pd

# Dossier contenant les fichiers Excel bruts
dossier = Path("donnees_brutes")

# Chercher tous les fichiers .xlsx
fichiers_excel = sorted(dossier.glob("*.xlsx"))

print("=" * 70)
print("INVENTAIRE DES FICHIERS EXCEL HCP")
print("=" * 70)

print(f"\nNombre de fichiers Excel trouvés : {len(fichiers_excel)}\n")

for numero, fichier in enumerate(fichiers_excel, start=1):

    print(f"{numero}. {fichier.name}")

    try:
        excel = pd.ExcelFile(fichier)

        print(f"   Nombre de feuilles : {len(excel.sheet_names)}")

        for feuille in excel.sheet_names:
            print(f"      - {feuille}")

    except Exception as erreur:
        print(f"   ERREUR : {erreur}")

    print()

print("=" * 70)
print("FIN DE L'INVENTAIRE")
print("=" * 70)