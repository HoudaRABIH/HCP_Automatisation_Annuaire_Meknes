from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import re
import warnings


warnings.filterwarnings("ignore")


DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_SORTIE = Path("sorties")

FICHIER_SORTIE = DOSSIER_SORTIE / "catalogue_blocs.xlsx"

DOSSIER_SORTIE.mkdir(exist_ok=True)


# ------------------------------------------------------------
# OUTILS
# ------------------------------------------------------------

def est_feuille_ignoree(nom):
    nom = nom.strip().lower()

    return (
        nom.startswith("pg")
        or "sommaire" in nom
        or "graph" in nom
        or "grraph" in nom
        or "graphique" in nom
    )


def nettoyer_texte(valeur):
    if valeur is None:
        return ""

    return str(valeur).strip()


def contient_francais(texte):
    return bool(
        re.search(
            r"[A-Za-zÀ-ÿ]",
            texte
        )
    )


# Exemple reconnu :
# 2 - 8 ...
# 2-8 ...
# 11 -17 ...
# 16 – 4 ...
MOTIF_TITRE = re.compile(
    r"^\s*(\d+)\s*[-–—]\s*(\d+)"
)


MOTIF_ANNEE = re.compile(
    r"\b(?:19|20)\d{2}\b"
)


def detecter_numero_titre(texte):
    resultat = MOTIF_TITRE.search(texte)

    if not resultat:
        return None

    return (
        int(resultat.group(1)),
        int(resultat.group(2))
    )


def choisir_titre(candidats):
    """
    Plusieurs cellules de la même ligne peuvent contenir
    le même titre en français et en arabe.

    On privilégie le titre français.
    """

    candidats_fr = [
        texte
        for texte in candidats
        if contient_francais(texte)
    ]

    if candidats_fr:
        return max(
            candidats_fr,
            key=len
        )

    return max(
        candidats,
        key=len
    )


def construire_titre_complet(
    titre_principal,
    ligne_titre,
    ligne_suivante_titre,
    cellules_par_ligne
):
    """
    Ajoute éventuellement les lignes suivantes du titre.

    Exemple :

    ligne 3 :
    2 - 8 Evolution de la population infantile

    ligne 4 :
    selon le milieu de résidence (en milliers)

    devient :

    2 - 8 Evolution de la population infantile
    selon le milieu de résidence (en milliers)
    """

    parties = [titre_principal]

    limite = min(
        ligne_titre + 3,
        ligne_suivante_titre - 1
    )

    for numero_ligne in range(
        ligne_titre + 1,
        limite + 1
    ):

        cellules = cellules_par_ligne.get(
            numero_ligne,
            []
        )

        textes = []

        for cellule in cellules:

            texte = nettoyer_texte(
                cellule.value
            )

            if not texte:
                continue

            if not contient_francais(texte):
                continue

            textes.append(texte)

        if not textes:
            continue

        meilleur = max(
            textes,
            key=len
        )

        meilleur_min = meilleur.lower()

        # Ne pas incorporer les headers ou sources
        if meilleur_min.startswith("source"):
            break

        if meilleur_min.startswith("année"):
            break

        # Une ligne composée essentiellement d'une année
        if (
            MOTIF_ANNEE.search(meilleur)
            and len(meilleur) < 35
        ):
            break

        # Une vraie continuation de titre est généralement
        # suffisamment descriptive.
        if len(meilleur) >= 8:
            parties.append(meilleur)

    # Eviter les répétitions
    resultat = []

    for partie in parties:

        if partie not in resultat:
            resultat.append(partie)

    return " ".join(resultat)


def extraire_annees(cellules):
    annees = set()

    for cellule in cellules:

        texte = nettoyer_texte(
            cellule.value
        )

        for annee in MOTIF_ANNEE.findall(
            texte
        ):
            # findall avec ce motif retourne seulement
            # le groupe 19/20 ; on utilise finditer ensuite.
            pass

        for resultat in MOTIF_ANNEE.finditer(
            texte
        ):
            annees.add(
                resultat.group(0)
            )

    return ", ".join(
        sorted(annees)
    )


def contient_meknes(cellules):
    for cellule in cellules:

        texte = nettoyer_texte(
            cellule.value
        )

        if re.search(
            r"\bmekn[eè]s\b",
            texte,
            re.IGNORECASE
        ):
            return True

        if "مكناس" in texte:
            return True

    return False


def contient_fes_meknes(cellules):
    variantes = [
        "fès - meknès",
        "fès-meknès",
        "fes - meknes",
        "fes-meknes",
        "فاس ــ مكناس",
        "فاس - مكناس"
    ]

    for cellule in cellules:

        texte = nettoyer_texte(
            cellule.value
        ).lower()

        for variante in variantes:

            if variante.lower() in texte:
                return True

    return False


# ------------------------------------------------------------
# PREPARATION DU RAPPORT
# ------------------------------------------------------------

rapport = Workbook()

catalogue = rapport.active
catalogue.title = "Blocs"

catalogue.append([
    "ID",
    "Fichier",
    "Feuille",
    "Chapitre source",
    "N° tableau source",
    "Numéro complet",
    "Titre",
    "Ligne début",
    "Ligne fin",
    "Dernière colonne",
    "Meknès détecté",
    "Fès-Meknès détecté",
    "Années détectées",
    "Nombre cellules avec données"
])


resume = rapport.create_sheet(
    "Resume"
)

resume.append([
    "Indicateur",
    "Valeur"
])


anomalies = rapport.create_sheet(
    "Anomalies"
)

anomalies.append([
    "Fichier",
    "Feuille",
    "Observation"
])


# ------------------------------------------------------------
# ANALYSE
# ------------------------------------------------------------

fichiers_excel = sorted(
    fichier
    for fichier in DOSSIER_SOURCE.glob("*.xlsx")
    if not fichier.name.startswith("~$")
)


print("=" * 74)
print("DETECTION DES VRAIS TABLEAUX DANS LES FEUILLES EXCEL HCP")
print("=" * 74)


nombre_blocs = 0
nombre_feuilles_avec_blocs = 0
nombre_blocs_meknes = 0


for numero_fichier, fichier in enumerate(
    fichiers_excel,
    start=1
):

    print(
        f"[{numero_fichier}/{len(fichiers_excel)}] "
        f"{fichier.name}"
    )

    try:

        classeur = load_workbook(
            fichier,
            data_only=False,
            read_only=False
        )

        for feuille in classeur.worksheets:

            if est_feuille_ignoree(
                feuille.title
            ):
                continue

            # ----------------------------------------
            # CELLULES REELLEMENT UTILISEES
            # ----------------------------------------

            cellules_utilisees = [
                cellule
                for cellule in feuille._cells.values()
                if cellule.value not in (None, "")
            ]

            if not cellules_utilisees:
                continue

            cellules_par_ligne = {}

            for cellule in cellules_utilisees:

                cellules_par_ligne.setdefault(
                    cellule.row,
                    []
                ).append(cellule)

            derniere_ligne_reelle = max(
                cellule.row
                for cellule in cellules_utilisees
            )

            # ----------------------------------------
            # RECHERCHE DES LIGNES DE TITRES
            # ----------------------------------------

            titres_par_ligne = {}

            for cellule in cellules_utilisees:

                texte = nettoyer_texte(
                    cellule.value
                )

                numero = detecter_numero_titre(
                    texte
                )

                if numero is None:
                    continue

                titres_par_ligne.setdefault(
                    cellule.row,
                    []
                ).append(texte)

            lignes_titres = sorted(
                titres_par_ligne.keys()
            )

            if not lignes_titres:

                anomalies.append([
                    fichier.name,
                    feuille.title,
                    "Aucun titre numéroté détecté"
                ])

                continue

            nombre_feuilles_avec_blocs += 1

            # ----------------------------------------
            # CREATION DES BLOCS
            # ----------------------------------------

            for index, ligne_debut in enumerate(
                lignes_titres
            ):

                candidats = titres_par_ligne[
                    ligne_debut
                ]

                titre_principal = choisir_titre(
                    candidats
                )

                numero = detecter_numero_titre(
                    titre_principal
                )

                if numero is None:
                    continue

                chapitre_source = numero[0]
                numero_tableau_source = numero[1]

                numero_complet = (
                    f"{chapitre_source}-"
                    f"{numero_tableau_source}"
                )

                if index + 1 < len(
                    lignes_titres
                ):

                    prochaine_ligne = (
                        lignes_titres[
                            index + 1
                        ]
                    )

                    ligne_fin = (
                        prochaine_ligne - 1
                    )

                else:

                    prochaine_ligne = (
                        derniere_ligne_reelle + 1
                    )

                    ligne_fin = (
                        derniere_ligne_reelle
                    )

                titre_complet = (
                    construire_titre_complet(
                        titre_principal,
                        ligne_debut,
                        prochaine_ligne,
                        cellules_par_ligne
                    )
                )

                cellules_bloc = [
                    cellule
                    for cellule
                    in cellules_utilisees
                    if (
                        ligne_debut
                        <= cellule.row
                        <= ligne_fin
                    )
                ]

                if not cellules_bloc:
                    continue

                derniere_colonne = max(
                    cellule.column
                    for cellule
                    in cellules_bloc
                )

                meknes = contient_meknes(
                    cellules_bloc
                )

                fes_meknes = (
                    contient_fes_meknes(
                        cellules_bloc
                    )
                )

                annees = extraire_annees(
                    cellules_bloc
                )

                nombre_blocs += 1

                if meknes:
                    nombre_blocs_meknes += 1

                catalogue.append([
                    nombre_blocs,
                    fichier.name,
                    feuille.title,
                    chapitre_source,
                    numero_tableau_source,
                    numero_complet,
                    titre_complet,
                    ligne_debut,
                    ligne_fin,
                    get_column_letter(
                        derniere_colonne
                    ),
                    (
                        "OUI"
                        if meknes
                        else "NON"
                    ),
                    (
                        "OUI"
                        if fes_meknes
                        else "NON"
                    ),
                    annees,
                    len(cellules_bloc)
                ])

        classeur.close()

    except Exception as erreur:

        anomalies.append([
            fichier.name,
            "",
            str(erreur)
        ])

        print(
            f"   ERREUR : {erreur}"
        )


# ------------------------------------------------------------
# RESUME
# ------------------------------------------------------------

resume.append([
    "Fichiers analysés",
    len(fichiers_excel)
])

resume.append([
    "Feuilles contenant au moins un tableau",
    nombre_feuilles_avec_blocs
])

resume.append([
    "Vrais tableaux détectés",
    nombre_blocs
])

resume.append([
    "Tableaux contenant Meknès",
    nombre_blocs_meknes
])


# ------------------------------------------------------------
# MISE EN FORME
# ------------------------------------------------------------

largeurs = {
    "A": 8,
    "B": 45,
    "C": 28,
    "D": 18,
    "E": 20,
    "F": 18,
    "G": 85,
    "H": 15,
    "I": 15,
    "J": 18,
    "K": 18,
    "L": 22,
    "M": 30,
    "N": 25
}

for colonne, largeur in largeurs.items():

    catalogue.column_dimensions[
        colonne
    ].width = largeur


catalogue.freeze_panes = "A2"

catalogue.auto_filter.ref = (
    catalogue.dimensions
)

resume.column_dimensions["A"].width = 45
resume.column_dimensions["B"].width = 20

anomalies.column_dimensions["A"].width = 45
anomalies.column_dimensions["B"].width = 30
anomalies.column_dimensions["C"].width = 60


# ------------------------------------------------------------
# SAUVEGARDE
# ------------------------------------------------------------

rapport.save(
    FICHIER_SORTIE
)


print()
print(
    f"Feuilles contenant des tableaux : "
    f"{nombre_feuilles_avec_blocs}"
)

print(
    f"Vrais tableaux détectés : "
    f"{nombre_blocs}"
)

print(
    f"Tableaux contenant Meknès : "
    f"{nombre_blocs_meknes}"
)

print()
print(
    f"Fichier créé : "
    f"{FICHIER_SORTIE}"
)

print("=" * 74)
print("DETECTION TERMINEE")
print("=" * 74)