from pathlib import Path
from docx import Document


DOSSIER_SORTIE = Path("sorties")

FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_OK.docx"
)


def remplacer_texte_preserver_style(cellule, nouveau_texte):
    nouveau_texte = str(nouveau_texte)

    if not cellule.paragraphs:
        cellule.text = nouveau_texte
        return

    p = cellule.paragraphs[0]

    if p.runs:
        p.runs[0].text = nouveau_texte
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.add_run(nouveau_texte)

    for p2 in cellule.paragraphs[1:]:
        for run in p2.runs:
            run.text = ""


def remplacer_dans_paragraphe_runs(paragraphe, ancien, nouveau):
    """
    Remplacement ciblé dans les runs existants pour préserver le formatage.
    """
    if ancien not in (paragraphe.text or ""):
        return 0

    n = 0
    for run in paragraphe.runs:
        if ancien in run.text:
            run.text = run.text.replace(ancien, nouveau)
            n += 1
    return n


if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Document introuvable : {FICHIER_ENTREE}"
    )

document = Document(FICHIER_ENTREE)


# ============================================================
# 9-10 — SUPERFICIES REBOISEES
# Source actuelle :
# A6 = 2023-2024
# D6 = 2024-2023
# B9 = provisoires (1)
# C9 = existantes (2)
# ============================================================

table_9_10 = document.tables[295 - 1]

if len(table_9_10.rows) != 6 or len(table_9_10.columns) != 4:
    raise RuntimeError("Structure Word 9-10 inattendue.")

# La cellule de gauche est fusionnée verticalement dans le modèle.
remplacer_texte_preserver_style(
    table_9_10.cell(0, 0),
    "2023-2024 En hectare"
)

# Retirer l'astérisque hérité de l'ancien modèle.
remplacer_texte_preserver_style(
    table_9_10.cell(0, 3),
    "2024-2023"
)

# Numéros de notes conformes à l'Excel courant.
remplacer_texte_preserver_style(
    table_9_10.cell(2, 1),
    "Provisoires(1)"
)

remplacer_texte_preserver_style(
    table_9_10.cell(2, 2),
    "Existantes(2)"
)


# ============================================================
# 9-11 — ESSENCES FORESTIERES
# Excel courant :
# Année 2024(1)
# ============================================================

table_9_11 = document.tables[297 - 1]

if len(table_9_11.rows) != 5 or len(table_9_11.columns) != 8:
    raise RuntimeError("Structure Word 9-11 inattendue.")

remplacer_texte_preserver_style(
    table_9_11.cell(0, 0),
    "Année 2024(1)"
)

remplacer_texte_preserver_style(
    table_9_11.cell(0, 7),
    "(1) 2024 سنة"
)


# ============================================================
# Corriger UNIQUEMENT le numéro de note associé au 9-11,
# sans toucher au texte de la source ni aux autres sources.
# On limite la recherche aux deux tableaux immédiatement après 9-11.
# ============================================================

for idx in range(297, min(len(document.tables), 300)):
    table = document.tables[idx]
    cellules_vues = set()

    for row in table.rows:
        for cell in row.cells:
            identifiant = id(cell._tc)
            if identifiant in cellules_vues:
                continue
            cellules_vues.add(identifiant)

            texte = cell.text or ""

            # Numéro de note isolé
            if texte.strip() == "(3)":
                remplacer_texte_preserver_style(cell, "(1)")
                continue

            # Phrase de note FR/AR : on change seulement le numéro.
            if "Inventaire National Forestier" in texte:
                for p in cell.paragraphs:
                    remplacer_dans_paragraphe_runs(p, "(3)", "(1)")

            if "الجرد الوطني" in texte:
                for p in cell.paragraphs:
                    remplacer_dans_paragraphe_runs(p, "(3)", "(1)")


document.save(FICHIER_SORTIE)

print("=" * 82)
print("CORRECTIONS AGRICULTURE APPLIQUEES")
print("=" * 82)
print("9-10 :")
print("  - 2024-2023 sans astérisque")
print("  - Provisoires(1)")
print("  - Existantes(2)")
print()
print("9-11 :")
print("  - Année 2024(1)")
print("  - note (1) mise en cohérence")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print()
print("Aucune autre source du Word n'a été modifiée.")
print("=" * 82)
