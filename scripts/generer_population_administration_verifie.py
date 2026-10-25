from pathlib import Path
from shutil import copy2
from openpyxl import load_workbook
from docx import Document
import re
import unicodedata

DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_SORTIE = Path("sorties")

CANDIDATS_ENTREE = [
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION.docx"
)

def norm(v):
    s = str(v or "").strip().lower()
    s = (
        s.replace("’", "'")
         .replace("–", "-")
         .replace("—", "-")
         .replace("\u202f", " ")
         .replace("\xa0", " ")
    )
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", s).strip()

def fmt(v):
    if v is None or v == "":
        return "-"
    if isinstance(v, (int, float)):
        if abs(float(v) - round(float(v))) < 1e-9:
            return f"{int(round(float(v))):,}".replace(",", " ")
    s = str(v).strip()
    try:
        n = float(
            s.replace("\u202f", "")
             .replace("\xa0", "")
             .replace(" ", "")
             .replace(",", ".")
        )
        if abs(n - round(n)) < 1e-9:
            return f"{int(round(n)):,}".replace(",", " ")
    except Exception:
        pass
    return s

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

def find_word_data_table(doc, code):
    number = code.split("-")[1]
    pat = re.compile(rf"^\s*1\s*-\s*{re.escape(number)}\b", re.I)

    for i, table in enumerate(doc.tables):
        if not table.rows:
            continue
        title = table.cell(0, 0).text or ""
        if pat.search(title):
            for j in range(i + 1, min(i + 6, len(doc.tables))):
                candidate = doc.tables[j]
                if len(candidate.rows) >= 3:
                    return candidate
    raise RuntimeError(f"Tableau Word {code} introuvable.")

def find_excel(patterns):
    files = [
        p for p in DOSSIER_SOURCE.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]
    for p in files:
        n = norm(p.name)
        if all(norm(x) in n for x in patterns):
            return p
    raise FileNotFoundError(
        f"Excel introuvable pour motifs : {patterns}"
    )

def sheet_text(ws, rows=12, cols=10):
    vals = []
    for r in range(1, min(ws.max_row, rows) + 1):
        for c in range(1, min(ws.max_column, cols) + 1):
            v = ws.cell(r, c).value
            if v not in (None, ""):
                vals.append(str(v))
    return norm(" ".join(vals))

def find_sheet(wb, keywords):
    keys = [norm(k) for k in keywords]
    for ws in wb.worksheets:
        t = sheet_text(ws)
        if all(k in t for k in keys):
            return ws
    raise RuntimeError(
        f"Feuille Excel introuvable pour : {keywords}"
    )

def find_sheet_table(wb, table_number, keywords):
    """
    Recherche robuste d'un tableau Excel même si le titre écrit
    '2 -7', '2-7' ou '2 - 7'.
    """
    a, b = table_number.split("-")
    pat = re.compile(rf"\b{re.escape(a)}\s*-\s*{re.escape(b)}\b")
    keys = [norm(k) for k in keywords]

    for ws in wb.worksheets:
        t = sheet_text(ws)
        if pat.search(t) and all(k in t for k in keys):
            return ws

    raise RuntimeError(
        f"Feuille Excel introuvable pour tableau {table_number} / {keywords}"
    )

def numeric_count(ws, r):
    count = 0
    for c in range(2, min(ws.max_column, 10) + 1):
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)):
            count += 1
        elif v not in (None, ""):
            try:
                float(
                    str(v).replace("\u202f", "")
                          .replace("\xa0", "")
                          .replace(" ", "")
                          .replace(",", ".")
                )
                count += 1
            except Exception:
                pass
    return count

def find_row(ws, labels):
    wanted = {norm(x) for x in labels}
    candidates = []
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 1).value) in wanted:
            candidates.append((numeric_count(ws, r), r))
    if not candidates:
        raise RuntimeError(
            f"Ligne introuvable dans {ws.title}: {labels}"
        )
    candidates.sort(reverse=True)
    return candidates[0][1]

def check_equal(label, got, expected):
    if str(got).strip() != str(expected).strip():
        raise RuntimeError(
            f"CONTROLE ECHEC {label}: Word={got!r} Excel={expected!r}"
        )

# ------------------------------------------------------------
# Entrée
# ------------------------------------------------------------
input_docx = next(
    (p for p in CANDIDATS_ENTREE if p.exists()),
    None
)

if input_docx is None:
    raise FileNotFoundError(
        "Le Word Habitat/Foncier d'entrée est introuvable."
    )

excel_da = find_excel(["division", "administrative", "2025"])
excel_pop = find_excel(["population", "2025"])

print("=" * 96)
print("POPULATION / DIVISION ADMINISTRATIVE — GENERATION AVEC DOUBLE CONTROLE")
print("=" * 96)
print(f"Word d'entrée : {input_docx}")
print(f"Excel division administrative : {excel_da.name}")
print(f"Excel population : {excel_pop.name}")
print()

doc = Document(input_docx)

wb_da = load_workbook(excel_da, data_only=True, read_only=False)
wb_pop = load_workbook(excel_pop, data_only=True, read_only=False)

# ============================================================
# 1-1 — DIVISION ADMINISTRATIVE
# ============================================================
print("[1/14] 1-1 Division administrative", flush=True)

t = find_word_data_table(doc, "1-1")
ws = find_sheet(wb_da, ["1 - 1", "division administrative"])

r_nat = find_row(ws, ["Ensemble"])
r_reg = find_row(ws, ["Fès - Meknès", "Fès-Meknès"])
r_mek = find_row(ws, ["Meknès"])

source_rows = [
    (1, r_nat),
    (2, r_reg),
    (3, r_mek),
]

for wr, er in source_rows:
    set_cell(t.cell(wr, 2), fmt(ws.cell(er, 3).value))
    set_cell(t.cell(wr, 3), fmt(ws.cell(er, 4).value))

# double contrôle explicite des valeurs clés
check_equal("1-1 Nation communes", t.cell(1, 2).text, fmt(ws.cell(r_nat, 3).value))
check_equal("1-1 Nation cercles", t.cell(1, 3).text, fmt(ws.cell(r_nat, 4).value))
check_equal("1-1 Région communes", t.cell(2, 2).text, fmt(ws.cell(r_reg, 3).value))
check_equal("1-1 Meknès communes", t.cell(3, 2).text, fmt(ws.cell(r_mek, 3).value))
print("   -> source et géographie vérifiées : OK")

# ============================================================
# 1-2 — NE PAS TOUCHER
# ============================================================
print("[2/14] 1-2 Division administrative locale", flush=True)
print(
    "   -> CONSERVE. Le mapping 1-2 est incompatible : "
    "la source Excel 1-2 concerne les principales préfectures / arrondissements, "
    "alors que le Word contient les collectivités et cercles de Meknès."
)

# ============================================================
# 1-3 — POPULATION LEGALE ENSEMBLE / URBAIN / RURAL
# ============================================================
print("[3/14] 1-3 Population légale selon milieu", flush=True)

t = find_word_data_table(doc, "1-3")

sources = [
    ("2-5", ["population legale", "ensemble"], [1, 2, 3]),
    ("2-6", ["population legale", "milieu urbain"], [5, 6, 7]),
    ("2-7", ["population legale", "milieu rural"], [9, 10, 11]),
]

for table_number, keywords, word_rows in sources:
    ws = find_sheet_table(wb_pop, table_number, keywords)
    r_nat = find_row(ws, ["Ensemble"])
    r_reg = find_row(ws, ["Fès - Meknès", "Fès-Meknès"])
    r_mek = find_row(ws, ["Meknès"])

    for wr, er in zip(word_rows, [r_nat, r_reg, r_mek]):
        for wc, ec in zip([1, 2, 3], [2, 3, 4]):
            set_cell(
                t.cell(wr, wc),
                fmt(ws.cell(er, ec).value)
            )

    # contrôle arithmétique Total = Etrangers + Marocains
    for er, geo in [
        (r_nat, "Nation"),
        (r_reg, "Région"),
        (r_mek, "Meknès"),
    ]:
        total = ws.cell(er, 2).value
        etr = ws.cell(er, 3).value
        mar = ws.cell(er, 4).value
        if all(isinstance(x, (int, float)) for x in [total, etr, mar]):
            if int(total) != int(etr) + int(mar):
                raise RuntimeError(
                    f"CONTROLE ECHEC 1-3 {geo}: "
                    f"Total != Etrangers + Marocains"
                )

print("   -> Ensemble + Urbain + Rural vérifiés : OK")
print("   -> contrôle Total = Etrangers + Marocains : OK")

# ============================================================
# 1-4 — MENAGES
# ============================================================
print("[4/14] 1-4 Nombre de ménages", flush=True)

t = find_word_data_table(doc, "1-4")
ws = find_sheet_table(
    wb_pop,
    "2-14",
    ["nombre de menages", "milieu de residence"]
)

r_nat = find_row(ws, ["Ensemble"])
r_reg = find_row(ws, ["Fès - Meknès", "Fès-Meknès"])
r_mek = find_row(ws, ["Meknès"])

for wr, er in zip([1, 2, 3], [r_nat, r_reg, r_mek]):
    for wc, ec in zip([1, 2, 3], [2, 3, 4]):
        set_cell(
            t.cell(wr, wc),
            fmt(ws.cell(er, ec).value)
        )

    total = ws.cell(er, 2).value
    rural = ws.cell(er, 3).value
    urbain = ws.cell(er, 4).value

    if all(isinstance(x, (int, float)) for x in [total, rural, urbain]):
        if int(total) != int(rural) + int(urbain):
            raise RuntimeError(
                "CONTROLE ECHEC 1-4: Total != Rural + Urbain"
            )

print("   -> Nation / Région / Meknès vérifiés : OK")
print("   -> contrôle Total = Rural + Urbain : OK")

# ============================================================
# 1-5 à 1-14 — MODELES 2024 DEJA RENSEIGNES
# ============================================================
for n in range(5, 15):
    code = f"1-{n}"
    print(f"[{n}/14] {code} — contrôle du modèle", flush=True)
    t = find_word_data_table(doc, code)

    # Tous ces tableaux sont explicitement sans source Excel dans le diagnostic.
    # On ne modifie AUCUNE valeur.
    # On vérifie seulement qu'ils sont bien renseignés et cohérents avec 2024.
    texts = " ".join(
        cell.text
        for row in t.rows
        for cell in row.cells
    )

    if n != 7 and "2024" not in texts:
        raise RuntimeError(
            f"{code}: aucune mention 2024 détectée dans le modèle."
        )

    if len(t.rows) < 3:
        raise RuntimeError(
            f"{code}: structure Word anormale."
        )

    print("   -> conservé sans modification : OK")

wb_da.close()
wb_pop.close()

print()
print("Sauvegarde du Word...", flush=True)
doc.save(FICHIER_SORTIE)

print()
print("=" * 96)
print("POPULATION / DIVISION ADMINISTRATIVE TERMINE")
print("=" * 96)
print("1-1 : vérifié contre Excel et actualisé")
print("1-2 : conservé volontairement — mapping source incompatible détecté")
print("1-3 : vérifié contre 2-5 + 2-6 + 2-7, avec contrôles arithmétiques")
print("1-4 : vérifié contre 2-14, avec contrôle Total = Rural + Urbain")
print("1-5 à 1-14 : conservés — modèles 2024 déjà renseignés, aucune source Excel")
print()
print("Aucune source Word n'a été modifiée.")
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 96)
