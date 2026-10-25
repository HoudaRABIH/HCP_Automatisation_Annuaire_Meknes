from pathlib import Path
from openpyxl import load_workbook
from docx import Document
import re
import unicodedata


# ============================================================
# CHEMINS
# ============================================================

DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_MODELES = Path("modeles")
DOSSIER_SORTIE = Path("sorties")

FICHIER_MAPPING = DOSSIER_SORTIE / "mapping_final_recettes.xlsx"
FICHIER_TRANSPORT = DOSSIER_SOURCE / "8. Transport_AS 2025.xlsx"
FICHIER_SORTIE = DOSSIER_SORTIE / "Annuaire_HCP_Transport.docx"


# ============================================================
# OUTILS GENERAUX
# ============================================================

def normaliser(texte):
    texte = str(texte or "").strip().lower()
    texte = texte.replace("’", "'")

    texte = "".join(
        caractere
        for caractere in unicodedata.normalize("NFD", texte)
        if unicodedata.category(caractere) != "Mn"
    )

    texte = re.sub(r"\s+", " ", texte)

    return texte


def feuille_par_nom_approximatif(classeur, nom):
    cible = normaliser(nom)

    for feuille in classeur.worksheets:
        if normaliser(feuille.title) == cible:
            return feuille

    raise KeyError(
        f"Feuille introuvable : {nom}"
    )


def formater_nombre(valeur):
    """
    Affichage HCP :
    - nombres entiers avec séparateurs par espaces
    - nombres décimaux Excel arrondis à l'entier le plus proche
      lorsque le tableau Word attend des effectifs
    - conserve une éventuelle étoile finale
    """
    if valeur is None:
        return ""

    if isinstance(valeur, bool):
        return str(valeur)

    etoile = ""

    if isinstance(valeur, str):
        texte = valeur.strip()

        if texte.endswith("*"):
            etoile = "*"
            texte = texte[:-1].strip()

        texte_compact = (
            texte
            .replace(" ", "")
            .replace("\u00A0", "")
            .replace(",", ".")
        )

        try:
            nombre = float(texte_compact)
        except ValueError:
            return valeur.strip()

    else:
        nombre = float(valeur)

    entier = int(round(nombre))

    return (
        f"{entier:,}".replace(",", " ")
        + etoile
    )


def remplacer_texte_preserver_style(cellule, nouveau_texte):
    """
    Remplace le contenu textuel en gardant la cellule, ses bordures,
    sa couleur, son alignement et le style du premier run.
    """
    if not cellule.paragraphs:
        cellule.text = nouveau_texte
        return

    paragraphe = cellule.paragraphs[0]

    if paragraphe.runs:
        paragraphe.runs[0].text = nouveau_texte

        for run in paragraphe.runs[1:]:
            run.text = ""

    else:
        paragraphe.add_run(
            nouveau_texte
        )

    for autre in cellule.paragraphs[1:]:
        for run in autre.runs:
            run.text = ""


def trouver_ligne_exacte(
    feuille,
    libelle,
    colonne=1,
    debut=1,
    fin=None
):
    cible = normaliser(libelle)

    if fin is None:
        fin = feuille.max_row

    for ligne in range(debut, fin + 1):
        valeur = feuille.cell(
            ligne,
            colonne
        ).value

        if normaliser(valeur) == cible:
            return ligne

    raise ValueError(
        f"'{libelle}' introuvable dans la feuille {feuille.title}"
    )


def trouver_annees(feuille, colonnes, debut=1, fin=15):
    annees = []

    for colonne in colonnes:
        annee = None

        for ligne in range(debut, fin + 1):
            valeur = feuille.cell(
                ligne,
                colonne
            ).value

            resultat = re.search(
                r"\b(20\d{2})\b",
                str(valeur or "")
            )

            if resultat:
                annee = resultat.group(1)
                break

        if annee is None:
            raise RuntimeError(
                f"Année introuvable dans colonne {colonne} "
                f"de la feuille {feuille.title}"
            )

        annees.append(annee)

    return annees


def charger_index_tables_word():
    wb = load_workbook(
        FICHIER_MAPPING,
        data_only=True
    )

    ws = wb["Mapping_final"]

    entetes = {
        cellule.value: cellule.column
        for cellule in ws[1]
    }

    index = {}

    for ligne in range(
        2,
        ws.max_row + 1
    ):
        code = str(
            ws.cell(
                ligne,
                entetes["N° Word"]
            ).value or ""
        ).strip()

        index[code] = {
            "table_donnees": int(
                ws.cell(
                    ligne,
                    entetes["Table Word données"]
                ).value
            )
        }

    wb.close()

    return index


def trouver_total_national_8_8(feuille):
    candidats = []

    for ligne in range(3, 57):
        if normaliser(
            feuille.cell(ligne, 1).value
        ) != "total":
            continue

        valeurs = [
            feuille.cell(
                ligne,
                colonne
            ).value
            for colonne in range(2, 6)
        ]

        if all(
            valeur not in (None, "")
            for valeur in valeurs
        ):
            candidats.append(ligne)

    if not candidats:
        raise RuntimeError(
            "Total national introuvable dans 8-8."
        )

    return candidats[-1]


def trouver_total_national_8_9(feuille):
    """
    Le tableau 8-9 contient le total national en A114,
    avec 2024 en C et 2023 en E.
    """
    for ligne in range(
        3,
        min(feuille.max_row, 130) + 1
    ):
        if normaliser(
            feuille.cell(ligne, 1).value
        ) != "total":
            continue

        valeur_2024 = feuille.cell(
            ligne,
            3
        ).value

        valeur_2023 = feuille.cell(
            ligne,
            5
        ).value

        if (
            valeur_2024 not in (None, "")
            and valeur_2023 not in (None, "")
        ):
            return ligne

    raise RuntimeError(
        "Total national introuvable dans 8-9."
    )


# ============================================================
# CHARGEMENT SOURCES
# ============================================================

if not FICHIER_TRANSPORT.exists():
    raise FileNotFoundError(
        f"Source introuvable : {FICHIER_TRANSPORT}"
    )

classeur = load_workbook(
    FICHIER_TRANSPORT,
    data_only=True,
    read_only=False
)

feuille_87 = feuille_par_nom_approximatif(
    classeur,
    "7(1)7(2)"
)

feuille_88 = feuille_par_nom_approximatif(
    classeur,
    "8"
)

feuille_89 = feuille_par_nom_approximatif(
    classeur,
    "9(1)9(2)"
)


# ============================================================
# EXTRACTION 14-1
# ============================================================

annee_14_1 = trouver_annees(
    feuille_88,
    [1],
    debut=1,
    fin=15
)[0]

ligne_nation_14_1 = (
    trouver_total_national_8_8(
        feuille_88
    )
)

ligne_region_14_1 = trouver_ligne_exacte(
    feuille_87,
    "Fès - Meknès",
    colonne=1,
    debut=3,
    fin=122
)

ligne_meknes_14_1 = trouver_ligne_exacte(
    feuille_87,
    "Meknès",
    colonne=1,
    debut=3,
    fin=122
)

nation_14_1 = [
    feuille_88.cell(
        ligne_nation_14_1,
        colonne
    ).value
    for colonne in range(2, 6)
]

region_14_1 = [
    feuille_87.cell(
        ligne_region_14_1,
        colonne
    ).value
    for colonne in range(2, 6)
]

meknes_14_1 = [
    feuille_87.cell(
        ligne_meknes_14_1,
        colonne
    ).value
    for colonne in range(2, 6)
]


# ============================================================
# EXTRACTION 14-2
# ============================================================

annees_14_2 = trouver_annees(
    feuille_89,
    [3, 5],
    debut=1,
    fin=12
)

annee_recente = annees_14_2[0]
annee_precedente = annees_14_2[1]

ligne_nation_14_2 = (
    trouver_total_national_8_9(
        feuille_89
    )
)

ligne_region_14_2 = trouver_ligne_exacte(
    feuille_89,
    "Fès - Meknès",
    colonne=1,
    debut=3,
    fin=128
)

ligne_meknes_14_2 = trouver_ligne_exacte(
    feuille_89,
    "Meknès",
    colonne=1,
    debut=3,
    fin=128
)

nation_14_2 = [
    feuille_89.cell(
        ligne_nation_14_2,
        3
    ).value,
    feuille_89.cell(
        ligne_nation_14_2,
        5
    ).value
]

region_14_2 = [
    feuille_89.cell(
        ligne_region_14_2,
        3
    ).value,
    feuille_89.cell(
        ligne_region_14_2,
        5
    ).value
]

meknes_14_2 = [
    feuille_89.cell(
        ligne_meknes_14_2,
        3
    ).value,
    feuille_89.cell(
        ligne_meknes_14_2,
        5
    ).value
]


classeur.close()


# ============================================================
# OUVERTURE MODELE WORD
# ============================================================

fichiers_word = [
    fichier
    for fichier in DOSSIER_MODELES.glob(
        "*.docx"
    )
    if not fichier.name.startswith(
        "~$"
    )
]

if not fichiers_word:
    raise FileNotFoundError(
        "Aucun .docx trouvé dans modeles."
    )

modele = fichiers_word[0]

document = Document(
    modele
)

index_tables = charger_index_tables_word()


# ============================================================
# MISE A JOUR 14-1
# ============================================================

table_14_1 = document.tables[
    index_tables["14-1"][
        "table_donnees"
    ] - 1
]

if (
    len(table_14_1.rows) != 8
    or len(table_14_1.columns) != 6
):
    raise RuntimeError(
        "Structure Word 14-1 inattendue."
    )

remplacer_texte_preserver_style(
    table_14_1.cell(0, 0),
    f"Année {annee_14_1}*"
)

remplacer_texte_preserver_style(
    table_14_1.cell(0, 5),
    f"سنة {annee_14_1}*"
)

for ligne_word, valeurs in zip(
    [5, 6, 7],
    [
        nation_14_1,
        region_14_1,
        meknes_14_1
    ]
):
    for colonne_word, valeur in enumerate(
        valeurs,
        start=1
    ):
        remplacer_texte_preserver_style(
            table_14_1.cell(
                ligne_word,
                colonne_word
            ),
            formater_nombre(
                valeur
            )
        )


# ============================================================
# MISE A JOUR 14-2
# ============================================================

table_14_2 = document.tables[
    index_tables["14-2"][
        "table_donnees"
    ] - 1
]

if (
    len(table_14_2.rows) != 6
    or len(table_14_2.columns) != 4
):
    raise RuntimeError(
        "Structure Word 14-2 inattendue."
    )

# Années
remplacer_texte_preserver_style(
    table_14_2.cell(0, 1),
    annee_recente
)

remplacer_texte_preserver_style(
    table_14_2.cell(0, 2),
    annee_precedente
)

# Données
for ligne_word, valeurs in zip(
    [3, 4, 5],
    [
        nation_14_2,
        region_14_2,
        meknes_14_2
    ]
):
    remplacer_texte_preserver_style(
        table_14_2.cell(
            ligne_word,
            1
        ),
        formater_nombre(
            valeurs[0]
        )
    )

    remplacer_texte_preserver_style(
        table_14_2.cell(
            ligne_word,
            2
        ),
        formater_nombre(
            valeurs[1]
        )
    )


# ============================================================
# SOURCES / NOTES
# ============================================================

# On ne modifie PAS encore les paragraphes de source du modèle ici.
# Ils seront traités de manière ciblée après validation visuelle des
# deux tableaux Transport, afin de ne pas modifier par erreur les
# références d'autres chapitres du document.


# ============================================================
# SAUVEGARDE
# ============================================================

DOSSIER_SORTIE.mkdir(
    exist_ok=True
)

document.save(
    FICHIER_SORTIE
)


# ============================================================
# RESULTAT
# ============================================================

print("=" * 78)
print("CHAPITRE TRANSPORT GENERE")
print("=" * 78)

print()
print("14-1 Immatriculation des véhicules")
print(
    f"Année : {annee_14_1}"
)
print(
    "Nation : "
    + " | ".join(
        formater_nombre(v)
        for v in nation_14_1
    )
)
print(
    "Région : "
    + " | ".join(
        formater_nombre(v)
        for v in region_14_1
    )
)
print(
    "Meknès : "
    + " | ".join(
        formater_nombre(v)
        for v in meknes_14_1
    )
)

print()
print("14-2 Permis de conduire délivrés")
print(
    f"Années : {annee_recente} / {annee_precedente}"
)
print(
    "Nation : "
    + " | ".join(
        formater_nombre(v)
        for v in nation_14_2
    )
)
print(
    "Région : "
    + " | ".join(
        formater_nombre(v)
        for v in region_14_2
    )
)
print(
    "Meknès : "
    + " | ".join(
        formater_nombre(v)
        for v in meknes_14_2
    )
)

print()
print(
    f"Document créé : {FICHIER_SORTIE}"
)

print()
print(
    "Le fichier modeles.docx original n'a pas été modifié."
)

print("=" * 78)
