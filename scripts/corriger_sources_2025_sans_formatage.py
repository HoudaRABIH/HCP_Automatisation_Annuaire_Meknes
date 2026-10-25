from pathlib import Path
from docx import Document


DOSSIER_SORTIE = Path("sorties")

# IMPORTANT :
# On repart du document AVANT la correction des sources, car ce fichier
# conserve la mise en page originale correcte.
FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement_Sources2025_OK.docx"
)


MARQUEUR_FR = "Annuaire Statistique du Maroc"
MARQUEUR_AR = "النشرة الإحصائية السنوية للمغرب"


def remplacer_annee_dans_paragraphe(paragraphe):
    """
    Ne remplace QUE 2023 -> 2025 dans un paragraphe qui contient
    explicitement la source Annuaire Statistique du Maroc.

    On ne reconstruit jamais le paragraphe et on ne déplace aucun texte :
    chaque run garde donc sa police, taille, couleur, italique, RTL, etc.
    """
    texte = paragraphe.text or ""

    if (
        MARQUEUR_FR not in texte
        and MARQUEUR_AR not in texte
    ):
        return 0

    if "2023" not in texte:
        return 0

    nb = 0

    for run in paragraphe.runs:
        if "2023" in run.text:
            run.text = run.text.replace(
                "2023",
                "2025"
            )
            nb += 1

    return nb


def traiter_tableau(tableau):
    nb = 0
    cellules_vues = set()

    for ligne in tableau.rows:
        for cellule in ligne.cells:
            identifiant = id(cellule._tc)

            if identifiant in cellules_vues:
                continue

            cellules_vues.add(identifiant)

            for paragraphe in cellule.paragraphs:
                nb += remplacer_annee_dans_paragraphe(
                    paragraphe
                )

            for sous_tableau in cellule.tables:
                nb += traiter_tableau(
                    sous_tableau
                )

    return nb


if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Document d'entrée introuvable : {FICHIER_ENTREE}"
    )


document = Document(
    FICHIER_ENTREE
)

nb_remplacements = 0


# Paragraphes normaux
for paragraphe in document.paragraphs:
    nb_remplacements += remplacer_annee_dans_paragraphe(
        paragraphe
    )


# Tableaux
for tableau in document.tables:
    nb_remplacements += traiter_tableau(
        tableau
    )


# En-têtes / pieds de page si nécessaire
for section in document.sections:
    for paragraphe in section.header.paragraphs:
        nb_remplacements += remplacer_annee_dans_paragraphe(
            paragraphe
        )

    for tableau in section.header.tables:
        nb_remplacements += traiter_tableau(
            tableau
        )

    for paragraphe in section.footer.paragraphs:
        nb_remplacements += remplacer_annee_dans_paragraphe(
            paragraphe
        )

    for tableau in section.footer.tables:
        nb_remplacements += traiter_tableau(
            tableau
        )


document.save(
    FICHIER_SORTIE
)


print("=" * 82)
print("SOURCES ANNUAIRE 2025 - CORRECTION SANS TOUCHER AU FORMATAGE")
print("=" * 82)
print(f"Document d'entrée : {FICHIER_ENTREE}")
print(f"Document créé     : {FICHIER_SORTIE}")
print()
print(f"Runs modifiés : {nb_remplacements}")
print()
print("Seule l'année 2023 -> 2025 a été remplacée dans les sources")
print("'Annuaire Statistique du Maroc'.")
print()
print("Aucune police, taille, couleur, position, source RGPH/HCP/ministère")
print("ou autre texte n'a été modifié.")
print("=" * 82)
