from pathlib import Path
from openpyxl import load_workbook
from zipfile import ZipFile
from lxml import etree


# ============================================================
# TABLEAUX SANTE DU WORD
# ============================================================

CODES_WORD = [
    "2-1", "2-2", "2-3", "2-4", "2-5", "2-6",
    "2-7", "2-8", "2-9", "2-10", "2-11", "2-12",
    "2-13", "2-14", "2-15", "2-16", "2-17", "2-18",
    "2-20", "2-21", "2-22", "2-23", "2-24"
]


# ============================================================
# CHEMINS
# ============================================================

DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_MODELES = Path("modeles")
DOSSIER_SORTIE = Path("sorties")

FICHIER_MAPPING = DOSSIER_SORTIE / "mapping_final_recettes.xlsx"
FICHIER_CATALOGUE = DOSSIER_SORTIE / "catalogue_tables_reelles.xlsx"
FICHIER_SORTIE = DOSSIER_SORTIE / "diagnostic_sante.txt"

# On prend automatiquement le Word cumulé le plus récent.
CANDIDATS_WORD = [
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_VALIDE.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_FINAL.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture.docx",
]


# ============================================================
# OUTILS
# ============================================================

def txt(v):
    if v is None:
        return ""
    return str(v).replace("\n", " ").strip()


def charger_mapping():
    print("1/5 Lecture du mapping...", flush=True)

    wb = load_workbook(
        FICHIER_MAPPING,
        data_only=True,
        read_only=True
    )

    ws = wb["Mapping_final"]

    headers = {
        c.value: c.column
        for c in ws[1]
    }

    resultat = {}

    for r in range(2, ws.max_row + 1):
        code = txt(
            ws.cell(
                r,
                headers["N° Word"]
            ).value
        )

        if code not in CODES_WORD:
            continue

        resultat[code] = {
            "table_word": int(
                ws.cell(
                    r,
                    headers["Table Word données"]
                ).value
            ),
            "theme": txt(
                ws.cell(
                    r,
                    headers["Thème Word"]
                ).value
            ),
            "titre": txt(
                ws.cell(
                    r,
                    headers["Titre Word"]
                ).value
            ),
            "type": txt(
                ws.cell(
                    r,
                    headers["Type de recette"]
                ).value
            ),
            "sources": txt(
                ws.cell(
                    r,
                    headers["Sources Excel"]
                ).value
            ),
        }

    wb.close()

    return resultat


def charger_catalogue():
    print("2/5 Lecture du catalogue...", flush=True)

    wb = load_workbook(
        FICHIER_CATALOGUE,
        data_only=True,
        read_only=True
    )

    ws = wb["Tables_logiques"]

    headers = {
        c.value: c.column
        for c in ws[1]
    }

    resultat = {}

    for r in range(2, ws.max_row + 1):
        numero = txt(
            ws.cell(
                r,
                headers["Numéro complet"]
            ).value
        )

        resultat[numero] = {
            "fichier": txt(
                ws.cell(
                    r,
                    headers["Fichier"]
                ).value
            ),
            "titre": txt(
                ws.cell(
                    r,
                    headers["Titre principal"]
                ).value
            ),
            "plages": txt(
                ws.cell(
                    r,
                    headers["Plage(s)"]
                ).value
            ),
            "feuilles": txt(
                ws.cell(
                    r,
                    headers["Feuille(s)"]
                ).value
            ),
        }

    wb.close()

    return resultat


def choisir_word():
    for fichier in CANDIDATS_WORD:
        if fichier.exists():
            return fichier

    fichiers = [
        f
        for f in DOSSIER_MODELES.glob("*.docx")
        if not f.name.startswith("~$")
    ]

    if not fichiers:
        raise FileNotFoundError(
            "Aucun document Word trouvé."
        )

    return fichiers[0]


def extraire_tables_word_xml(docx_path, indices):
    """
    Lecture directe du XML Word :
    beaucoup plus rapide que python-docx pour ce gros document.
    """

    print("3/5 Lecture directe du Word...", flush=True)

    ns = {
        "w": (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )
    }

    with ZipFile(docx_path, "r") as z:
        xml = z.read(
            "word/document.xml"
        )

    root = etree.fromstring(
        xml
    )

    tables = root.xpath(
        ".//w:tbl",
        namespaces=ns
    )

    resultat = {}

    for no_objet in indices:
        idx = no_objet - 1

        if idx < 0 or idx >= len(tables):
            resultat[no_objet] = [
                "ERREUR: table Word introuvable"
            ]
            continue

        tbl = tables[idx]

        rows = tbl.xpath(
            "./w:tr",
            namespaces=ns
        )

        lignes = []

        for ir, row in enumerate(
            rows,
            start=1
        ):
            cells = row.xpath(
                "./w:tc",
                namespaces=ns
            )

            parts = []

            for ic, cell in enumerate(
                cells,
                start=1
            ):
                textes = cell.xpath(
                    ".//w:t/text()",
                    namespaces=ns
                )

                contenu = " ".join(
                    t.strip()
                    for t in textes
                    if t.strip()
                )

                parts.append(
                    f"C{ic}={contenu!r}"
                )

            lignes.append(
                f"R{ir}: "
                + " | ".join(parts)
            )

        resultat[no_objet] = lignes

    return resultat


CACHE_EXCEL = {}


def ouvrir_excel(nom_fichier):
    if nom_fichier not in CACHE_EXCEL:
        print(
            f"   Ouverture Excel : {nom_fichier}",
            flush=True
        )

        CACHE_EXCEL[nom_fichier] = load_workbook(
            DOSSIER_SOURCE / nom_fichier,
            data_only=True,
            read_only=True
        )

    return CACHE_EXCEL[
        nom_fichier
    ]


def extraire_plages(info):
    wb = ouvrir_excel(
        info["fichier"]
    )

    sortie = []

    plages = [
        p.strip()
        for p in info["plages"].split("|")
        if p.strip()
    ]

    for plage in plages:
        if "!" not in plage:
            continue

        nom_feuille, bornes = plage.rsplit(
            "!",
            1
        )

        if ":" not in bornes:
            continue

        debut_txt, fin_txt = bornes.split(
            ":",
            1
        )

        try:
            debut = int(debut_txt)
            fin = int(fin_txt)
        except ValueError:
            continue

        if nom_feuille not in wb.sheetnames:
            sortie.append(
                f"ERREUR: feuille {nom_feuille!r} absente"
            )
            continue

        ws = wb[nom_feuille]

        sortie.append(
            f"--- {nom_feuille} / lignes {debut}:{fin} ---"
        )

        # Suffisant pour les tableaux Santé.
        max_col = min(
            ws.max_column,
            30
        )

        for row in ws.iter_rows(
            min_row=debut,
            max_row=fin,
            min_col=1,
            max_col=max_col
        ):
            valeurs = []

            for cell in row:
                if cell.value in (
                    None,
                    ""
                ):
                    continue

                valeurs.append(
                    f"{cell.coordinate}="
                    f"{txt(cell.value)!r}"
                )

            if valeurs:
                sortie.append(
                    " | ".join(valeurs)
                )

    return sortie


# ============================================================
# EXECUTION
# ============================================================

mapping = charger_mapping()
catalogue = charger_catalogue()

word_path = choisir_word()

indices = [
    mapping[code]["table_word"]
    for code in CODES_WORD
    if code in mapping
]

word_tables = extraire_tables_word_xml(
    word_path,
    indices
)


print("4/5 Lecture des sources Santé...", flush=True)

rapport = [
    "=" * 110,
    "DIAGNOSTIC DU CHAPITRE SANTE",
    "=" * 110,
    f"Document Word analysé : {word_path}",
]


for position, code in enumerate(
    CODES_WORD,
    start=1
):
    print(
        f"   [{position}/{len(CODES_WORD)}] {code}",
        flush=True
    )

    rapport.extend([
        "",
        "#" * 110,
        f"TABLEAU WORD {code}",
        "#" * 110,
    ])

    if code not in mapping:
        rapport.append(
            "ERREUR: absent du mapping"
        )
        continue

    m = mapping[code]

    rapport.extend([
        f"Thème : {m['theme']}",
        f"Titre Word : {m['titre']}",
        f"Type : {m['type']}",
        f"Sources : {m['sources']}",
        f"N° objet Word : {m['table_word']}",
        "",
        "STRUCTURE WORD",
        "-" * 110,
    ])

    rapport.extend(
        word_tables.get(
            m["table_word"],
            [
                "ERREUR: structure Word absente"
            ]
        )
    )

    sources = [
        s.strip()
        for s in m["sources"].split("+")
        if s.strip()
    ]

    if not sources:
        rapport.append("")
        rapport.append(
            "AUCUNE SOURCE EXCEL : tableau à conserver dans le modèle."
        )

    for source in sources:
        rapport.extend([
            "",
            "STRUCTURE EXCEL",
            "-" * 110,
        ])

        if source not in catalogue:
            rapport.append(
                f"ERREUR: source {source} absente du catalogue"
            )
            continue

        info = catalogue[source]

        rapport.extend([
            f"Source {source}",
            f"Fichier : {info['fichier']}",
            f"Titre : {info['titre']}",
            f"Feuille(s) : {info['feuilles']}",
            f"Plage(s) : {info['plages']}",
        ])

        rapport.extend(
            extraire_plages(
                info
            )
        )


for wb in CACHE_EXCEL.values():
    wb.close()


print("5/5 Ecriture du diagnostic...", flush=True)

FICHIER_SORTIE.write_text(
    "\n".join(rapport),
    encoding="utf-8"
)


print()
print("=" * 82)
print("DIAGNOSTIC SANTE TERMINE")
print("=" * 82)
print(
    f"Tableaux analysés : {len(CODES_WORD)}"
)
print(
    f"Fichier créé : {FICHIER_SORTIE}"
)
print("=" * 82)
