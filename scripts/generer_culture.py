from pathlib import Path
from openpyxl import load_workbook
from docx import Document
import re
import unicodedata


# ============================================================
# CHEMINS
# ============================================================

DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_SORTIE = Path("sorties")

FICHIER_CULTURE = (
    DOSSIER_SOURCE
    / "16. Activités culturelles_Loisirs et Entraide_AS 2025.xlsx"
)

CANDIDATS_ENTREE = [
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_COLORS_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_FIXED.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx"
)


# ============================================================
# OUTILS
# ============================================================

def norm(v):
    s = str(v or "").strip().lower()
    s = (
        s.replace("’", "'")
         .replace("–", "-")
         .replace("—", "-")
         .replace("‐", "-")
    )
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def fmt(v):
    if v is None or v == "":
        return "-"

    if isinstance(v, str):
        t = v.strip()
        if t in ("-", "–", "…", "..."):
            return t
        try:
            n = float(
                t.replace(" ", "").replace(",", ".")
            )
        except ValueError:
            return t
    else:
        try:
            n = float(v)
        except (TypeError, ValueError):
            return str(v)

    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n)):,}".replace(",", " ")

    # Pour les recettes cinéma, conserver 3 décimales si nécessaire.
    if abs(n * 1000 - round(n * 1000)) < 1e-9:
        out = f"{n:.3f}".rstrip("0").rstrip(".")
    else:
        out = str(n)

    return out.replace(".", ",")


def set_cell(cell, value):
    value = str(value)

    if not cell.paragraphs:
        cell.text = value
        return

    p = cell.paragraphs[0]

    if p.runs:
        p.runs[0].text = value
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.add_run(value)

    for p2 in cell.paragraphs[1:]:
        for run in p2.runs:
            run.text = ""


def set_year(cell, year_text):
    old = cell.text or ""

    # Remplacer année simple ou année avec *
    new = re.sub(
        r"\b20\d{2}\*?\b",
        str(year_text),
        old,
        count=1
    )

    if new != old:
        set_cell(cell, new)


# ============================================================
# LOCALISATION WORD PAR TITRE — PAS PAR N° D'OBJET
# ============================================================

def find_word_data_table(document, code):
    """
    Recherche le titre 6-x, puis prend le premier vrai tableau de données
    qui suit. Cela reste correct même si le nombre total de tableaux Word
    a changé (413 au lieu de 414).
    """
    pattern = re.compile(
        rf"^\s*6\s*-\s*{re.escape(code)}\b",
        re.I
    )

    title_idx = None

    for i, table in enumerate(document.tables):
        if not table.rows:
            continue

        first = table.cell(0, 0).text or ""

        if pattern.search(first):
            title_idx = i
            break

    if title_idx is None:
        raise RuntimeError(
            f"Titre Word 6-{code} introuvable."
        )

    for j in range(
        title_idx + 1,
        min(len(document.tables), title_idx + 5)
    ):
        t = document.tables[j]

        # Les titres ont 1 ligne ; les vrais tableaux ont plusieurs lignes.
        if len(t.rows) >= 3:
            return t

    raise RuntimeError(
        f"Tableau de données Word 6-{code} introuvable après son titre."
    )


# ============================================================
# LOCALISATION DYNAMIQUE DES TABLES EXCEL 16-x
# ============================================================

TITLE_RE = re.compile(r"^\s*16\s*-\s*(\d+)\b", re.I)


def sheets_for_source(wb, source_num):
    result = []

    for ws in wb.worksheets:
        found = False

        for r in range(1, ws.max_row + 1):
            for c in range(1, min(ws.max_column, 10) + 1):
                value = ws.cell(r, c).value

                if not isinstance(value, str):
                    continue

                m = TITLE_RE.match(value)

                if m and int(m.group(1)) == source_num:
                    found = True
                    break

            if found:
                break

        if found:
            result.append(ws)

    return result


def row_numeric_count(ws, r):
    count = 0

    for c in range(2, min(ws.max_column, 15) + 1):
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)):
            count += 1
        elif isinstance(v, str):
            t = v.strip()
            if t in ("-", "…", "..."):
                count += 1
            else:
                try:
                    float(t.replace(" ", "").replace(",", "."))
                    count += 1
                except ValueError:
                    pass

    return count


def find_best_row(worksheets, labels):
    labels_n = {norm(x) for x in labels}
    candidates = []

    for ws in worksheets:
        for r in range(1, ws.max_row + 1):
            if norm(ws.cell(r, 1).value) in labels_n:
                candidates.append(
                    (row_numeric_count(ws, r), ws, r)
                )

    if not candidates:
        raise RuntimeError(
            f"Ligne Excel {labels!r} introuvable."
        )

    candidates.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return candidates[0][1], candidates[0][2]


def source_year(worksheets, default=None):
    years = []

    for ws in worksheets:
        for r in range(1, min(ws.max_row, 90) + 1):
            for c in range(1, min(ws.max_column, 10) + 1):
                v = ws.cell(r, c).value

                if v is None:
                    continue

                m = re.search(
                    r"\b(20\d{2})\s*(\*)?",
                    str(v)
                )

                if m:
                    years.append(
                        m.group(1) + ("*" if m.group(2) else "")
                    )

    if years:
        # La première année du tableau est celle à afficher.
        return years[0]

    return default


def get_geo_rows(worksheets, need_region=True):
    ws_nat, r_nat = find_best_row(
        worksheets,
        ["Ensemble", "Total", "Nation"]
    )

    result = {
        "nation": (ws_nat, r_nat)
    }

    if need_region:
        ws_reg, r_reg = find_best_row(
            worksheets,
            [
                "Fès - Meknès",
                "Fès-Meknès",
                "Fes - Meknes",
                "Fes-Meknes"
            ]
        )

        result["region"] = (
            ws_reg,
            r_reg
        )

    ws_mek, r_mek = find_best_row(
        worksheets,
        [
            "Meknès",
            "Meknes",
            "Meknès ( Ouisselane )",
            "Meknès (Ouislane )",
            "Meknes ( Ouisselane )",
        ]
    )

    result["meknes"] = (
        ws_mek,
        r_mek
    )

    return result


def val(ws, r, c):
    return fmt(
        ws.cell(r, c).value
    )


# ============================================================
# ECRITURE GENERIQUE GEOGRAPHIQUE
# ============================================================

def fill_geo_table(
    table,
    geo_rows,
    source_cols,
    word_rows,
    word_cols=None
):
    if word_cols is None:
        word_cols = list(
            range(1, 1 + len(source_cols))
        )

    for geo_name, wr in word_rows.items():
        ws, er = geo_rows[geo_name]

        for wc, ec in zip(
            word_cols,
            source_cols
        ):
            set_cell(
                table.cell(wr, wc),
                val(ws, er, ec)
            )


# ============================================================
# EXECUTION
# ============================================================

input_docx = None

for candidate in CANDIDATS_ENTREE:
    if candidate.exists():
        input_docx = candidate
        break

if input_docx is None:
    raise FileNotFoundError(
        "Aucun fichier Word d'entrée Culture trouvé dans sorties."
    )

if not FICHIER_CULTURE.exists():
    raise FileNotFoundError(
        f"Excel Culture introuvable : {FICHIER_CULTURE}"
    )


print("=" * 92)
print("GENERATION DU CHAPITRE ACTIVITES CULTURELLES ET LOISIRS")
print("=" * 92)
print(f"Word d'entrée : {input_docx}")
print()

doc = Document(input_docx)

wb = load_workbook(
    FICHIER_CULTURE,
    data_only=True,
    read_only=False
)


# ============================================================
# 6-1 — FOYERS FEMININS
# ============================================================

print("[1/11] 6-1 Foyers féminins", flush=True)

t = find_word_data_table(doc, "1")
sheets = sheets_for_source(wb, 1)

if not sheets:
    raise RuntimeError("Source Excel 16-1 introuvable.")

year = source_year(sheets, "2024")
set_year(t.cell(0, 0), year)
set_year(t.cell(0, len(t.columns) - 1), year)

geo = get_geo_rows(sheets)

# 7 indicateurs du modèle ; prendre B:H de la source.
fill_geo_table(
    t,
    geo,
    [2, 3, 4, 5, 6, 7, 8],
    {
        "nation": 5,
        "region": 6,
        "meknes": 7,
    }
)


# ============================================================
# 6-2 — JARDINS D'ENFANTS
# ============================================================

print("[2/11] 6-2 Jardins d'enfants", flush=True)

t = find_word_data_table(doc, "2")
sheets = sheets_for_source(wb, 2)

if not sheets:
    raise RuntimeError("Source Excel 16-2 introuvable.")

year = source_year(sheets, "2024*")
set_year(t.cell(0, 0), year)
set_year(t.cell(0, len(t.columns) - 1), year)

geo = get_geo_rows(sheets)

# La source 2025 fournit :
# B = activités de rayonnement
# C = activités normales
# D = nombre d'établissements
#
# Les colonnes "jardinières Aide/Fixe" n'existent plus dans la source.
# On les met à "-" pour ne PAS conserver des valeurs 2021 sous une année 2024.
for name, wr in {
    "nation": 5,
    "region": 6,
    "meknes": 7,
}.items():
    ws, er = geo[name]

    set_cell(t.cell(wr, 1), val(ws, er, 2))
    set_cell(t.cell(wr, 2), val(ws, er, 3))
    set_cell(t.cell(wr, 3), "-")
    set_cell(t.cell(wr, 4), "-")
    set_cell(t.cell(wr, 5), val(ws, er, 4))


# ============================================================
# 6-3 — MAISONS DES JEUNES
# ============================================================

print("[3/11] 6-3 Maisons des jeunes", flush=True)

t = find_word_data_table(doc, "3")
sheets = sheets_for_source(wb, 3)
year = source_year(sheets, "2022")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, len(t.columns) - 1), year)

geo = get_geo_rows(sheets)

fill_geo_table(
    t,
    geo,
    [2, 3, 4, 5],
    {
        "nation": 4,
        "region": 5,
        "meknes": 6,
    }
)


# ============================================================
# 6-4 — CENTRES D'ACCUEIL
# ============================================================

print("[4/11] 6-4 Centres d'accueil", flush=True)

t = find_word_data_table(doc, "4")
sheets = sheets_for_source(wb, 4)
year = source_year(sheets, "2018")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 3), year)

geo = get_geo_rows(
    sheets,
    need_region=False
)

# Nation = Ensemble, Meknès = Meknès (Ouisselane)
for name, wr in {
    "nation": 2,
    "meknes": 3,
}.items():
    ws, er = geo[name]
    set_cell(t.cell(wr, 1), val(ws, er, 2))
    set_cell(t.cell(wr, 2), val(ws, er, 4))


# ============================================================
# 6-5 — COLONIES DE VACANCES
# ============================================================

print("[5/11] 6-5 Colonies de vacances", flush=True)

t = find_word_data_table(doc, "5")
sheets = sheets_for_source(wb, 5)
year = source_year(sheets, "2019")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 2), year)

geo = get_geo_rows(sheets)

for name, wr in {
    "nation": 1,
    "region": 2,
    "meknes": 3,
}.items():
    ws, er = geo[name]
    set_cell(
        t.cell(wr, 1),
        val(ws, er, 3)
    )


# ============================================================
# 6-6 — INSTALLATIONS SPORTIVES
# ============================================================

print("[6/11] 6-6 Installations sportives", flush=True)

t = find_word_data_table(doc, "6")
sheets = sheets_for_source(wb, 8)
year = source_year(sheets, "2024")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 7), year)

geo = get_geo_rows(sheets)

fill_geo_table(
    t,
    geo,
    [2, 3, 4, 5, 6, 7],
    {
        "nation": 3,
        "region": 4,
        "meknes": 5,
    }
)


# ============================================================
# 6-7 — ACTIVITE CINEMATOGRAPHIQUE
# ============================================================

print("[7/11] 6-7 Activité cinématographique", flush=True)

t = find_word_data_table(doc, "7")

# Le mapping initial 16-11 est un tableau NATIONAL multi-années,
# alors que le Word exige Nation/Région/Meknès.
# On cherche d'abord une vraie table géographique 16-10.
sheets_geo_cinema = sheets_for_source(wb, 10)

can_update_cinema = False

if sheets_geo_cinema:
    try:
        geo_cinema = get_geo_rows(
            sheets_geo_cinema
        )
        can_update_cinema = True
    except Exception:
        can_update_cinema = False

if can_update_cinema:
    year = source_year(
        sheets_geo_cinema,
        "2024"
    )

    set_year(t.cell(0, 0), year)
    set_year(t.cell(0, 5), year)

    # Détection des colonnes par mots d'en-tête.
    def detect_col(sheets, words):
        words = [norm(w) for w in words]

        for ws in sheets:
            for c in range(2, min(ws.max_column, 12) + 1):
                header = []

                for r in range(1, min(ws.max_row, 25) + 1):
                    v = ws.cell(r, c).value
                    if v not in (None, ""):
                        header.append(str(v))

                h = norm(" ".join(header))

                if all(w in h for w in words):
                    return c

        raise RuntimeError(
            f"Colonne cinéma {words} introuvable."
        )

    c_rec = detect_col(
        sheets_geo_cinema,
        ["recette"]
    )
    c_ent = detect_col(
        sheets_geo_cinema,
        ["entree"]
    )
    c_places = detect_col(
        sheets_geo_cinema,
        ["place"]
    )
    c_salles = detect_col(
        sheets_geo_cinema,
        ["salle"]
    )

    for name, wr in {
        "nation": 5,
        "region": 6,
        "meknes": 7,
    }.items():
        ws, er = geo_cinema[name]

        for wc, ec in zip(
            [1, 2, 3, 4],
            [c_rec, c_ent, c_places, c_salles]
        ):
            set_cell(
                t.cell(wr, wc),
                val(ws, er, ec)
            )

else:
    print(
        "   -> 6-7 conservé : aucune source géographique Nation/Région/Meknès détectée."
    )


# ============================================================
# 6-8 — PATRIMOINE CULTUREL
# ============================================================

print("[8/11] 6-8 Patrimoine culturel", flush=True)

t = find_word_data_table(doc, "8")
sheets = sheets_for_source(wb, 13)
year = source_year(sheets, "2024")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 3), year)

geo = get_geo_rows(sheets)

for name, wr in {
    "nation": 1,
    "region": 2,
    "meknes": 3,
}.items():
    ws, er = geo[name]

    entries = ws.cell(er, 2).value

    def num_or_zero(v):
        if v in (None, "", "-", "…", "..."):
            return 0
        try:
            return float(
                str(v).replace(" ", "").replace(",", ".")
            )
        except ValueError:
            return 0

    sites = (
        num_or_zero(ws.cell(er, 3).value)
        + num_or_zero(ws.cell(er, 4).value)
    )

    set_cell(
        t.cell(wr, 1),
        fmt(entries)
    )
    set_cell(
        t.cell(wr, 2),
        fmt(sites)
    )


# ============================================================
# 6-9 — BIBLIOTHEQUES
# ============================================================

print("[9/11] 6-9 Bibliothèques", flush=True)

t = find_word_data_table(doc, "9")
sheets = sheets_for_source(wb, 15)
year = source_year(sheets, "2024*")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 5), year)

geo = get_geo_rows(sheets)

fill_geo_table(
    t,
    geo,
    [2, 3, 4, 5],
    {
        "nation": 1,
        "region": 2,
        "meknes": 3,
    }
)


# ============================================================
# 6-10 — THEATRES / SALLES DE SPECTACLE
# ============================================================

print("[10/11] 6-10 Théâtres et salles de spectacle", flush=True)

t = find_word_data_table(doc, "10")
sheets = sheets_for_source(wb, 16)
year = source_year(sheets, "2024*")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 5), year)

geo = get_geo_rows(sheets)

fill_geo_table(
    t,
    geo,
    [2, 3, 4, 5],
    {
        "nation": 1,
        "region": 2,
        "meknes": 3,
    }
)


# ============================================================
# 6-11 — CONSERVATOIRES
# ============================================================

print("[11/11] 6-11 Conservatoires de musique", flush=True)

t = find_word_data_table(doc, "11")
sheets = sheets_for_source(wb, 17)
year = source_year(sheets, "2024")

set_year(t.cell(0, 0), year)
set_year(t.cell(0, 4), year)

geo = get_geo_rows(sheets)

fill_geo_table(
    t,
    geo,
    [2, 3, 4],
    {
        "nation": 2,
        "region": 3,
        "meknes": 4,
    }
)


# ============================================================
# SAUVEGARDE
# ============================================================

wb.close()

print()
print("Sauvegarde du Word...", flush=True)

doc.save(
    FICHIER_SORTIE
)

print()
print("=" * 92)
print("CHAPITRE CULTURE / LOISIRS GENERE")
print("=" * 92)
print("6-1  : actualisé")
print("6-2  : actualisé ; colonnes jardinières indisponibles -> '-'")
print("6-3  : actualisé")
print("6-4  : données 2018 conservées/validées depuis la source")
print("6-5  : données 2019 conservées/validées depuis la source")
print("6-6  : actualisé 2024")
print(
    "6-7  : actualisé seulement si une source cinéma géographique "
    "Nation/Région/Meknès est détectée ; sinon conservé"
)
print("6-8  : actualisé 2024 ; sites = avec billetterie + sans billetterie")
print("6-9  : actualisé 2024*")
print("6-10 : actualisé 2024*")
print("6-11 : actualisé 2024")
print()
print("Les sources Word n'ont pas été modifiées.")
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 92)
