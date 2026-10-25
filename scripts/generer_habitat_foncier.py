from pathlib import Path
from openpyxl import load_workbook
from docx import Document
import re
import unicodedata

DOSSIER_SOURCE = Path("donnees_brutes")
DOSSIER_SORTIE = Path("sorties")

FICHIER_EXCEL = DOSSIER_SOURCE / "7. Foncier_AS 2025.xlsx"

CANDIDATS_ENTREE = [
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_COLORS_OK.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER.docx"
)

def norm(v):
    s = str(v or "").strip().lower()
    s = (
        s.replace("’", "'")
         .replace("–", "-")
         .replace("—", "-")
         .replace("‐", "-")
         .replace("\u202f", " ")
         .replace("\xa0", " ")
    )
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def fmt(v, decimals=2):
    if v is None or v == "":
        return "-"
    if isinstance(v, str):
        t = v.strip()
        if t in ("-", "–", "…", "..."):
            return t
        try:
            n = float(
                t.replace("\u202f", "")
                 .replace("\xa0", "")
                 .replace(" ", "")
                 .replace(",", ".")
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

    out = f"{n:,.{decimals}f}"
    out = out.replace(",", "X").replace(".", ",").replace("X", " ")
    if "," in out:
        out = out.rstrip("0").rstrip(",")
    return out

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

def replace_2023_source_with_2025(table):
    for row in table.rows:
        for cell in row.cells:
            text = cell.text or ""
            if (
                "Annuaire Statistique du Maroc 2023" in text
                or "النشرة الإحصائية السنوية للمغرب 2023" in text
            ):
                new = (
                    text.replace(
                        "Annuaire Statistique du Maroc 2023",
                        "Annuaire Statistique du Maroc 2025"
                    ).replace(
                        "النشرة الإحصائية السنوية للمغرب 2023",
                        "النشرة الإحصائية السنوية للمغرب 2025"
                    )
                )
                set_cell(cell, new)

def find_word_data_table(doc, code):
    number = code.split("-")[1]
    pat = re.compile(rf"^\s*13\s*-\s*{re.escape(number)}\b", re.I)

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

def sheet_text(ws, max_rows=18):
    vals = []
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(r, c).value
            if v not in (None, ""):
                vals.append(str(v))
    return norm(" ".join(vals))

def sheets_matching(wb, keywords, mode=None):
    keywords = [norm(k) for k in keywords]
    result = []
    for ws in wb.worksheets:
        text = sheet_text(ws)
        if not all(k in text for k in keywords):
            continue
        if mode == "urbain_rural":
            if "urbain + rural" not in text and "urbain+rural" not in text:
                continue
        elif mode == "milieu_rural":
            if "milieu rural" not in text:
                continue
        result.append(ws)
    return result

def numeric_count(ws, r):
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
                    float(
                        t.replace("\u202f", "")
                         .replace("\xa0", "")
                         .replace(" ", "")
                         .replace(",", ".")
                    )
                    count += 1
                except ValueError:
                    pass
    return count

def find_row(sheets, labels):
    labels = {norm(x) for x in labels}
    candidates = []
    for ws in sheets:
        for r in range(1, ws.max_row + 1):
            if norm(ws.cell(r, 1).value) in labels:
                candidates.append((numeric_count(ws, r), ws, r))
    if not candidates:
        return None, None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2]

def geo_rows(sheets):
    return {
        "nation": find_row(sheets, ["Total", "Ensemble", "Nation"]),
        "region": find_row(
            sheets,
            ["Fès - Meknès", "Fès-Meknès", "Fes - Meknes", "Fes-Meknes"]
        ),
        "menzeh": find_row(sheets, ["Meknès Menzeh", "Meknes Menzeh"]),
        "ismailia": find_row(
            sheets,
            ["Meknès Ismailia", "Meknès Ismaila", "Meknes Ismailia", "Meknes Ismaila"]
        ),
    }

def all_geo_found(rows):
    return all(ws is not None and r is not None for ws, r in rows.values())

def value(ws, row, col, decimals=2):
    return fmt(ws.cell(row, col).value, decimals=decimals)

def fill_number_surface_section(table, word_start_row, sheets):
    rows = geo_rows(sheets)
    if not all_geo_found(rows):
        return False

    order = [
        ("nation", "nombre"),
        ("nation", "surface"),
        ("region", "nombre"),
        ("region", "surface"),
        ("menzeh", "nombre"),
        ("menzeh", "surface"),
        ("ismailia", "nombre"),
        ("ismailia", "surface"),
    ]

    for offset, (geo, metric) in enumerate(order):
        ws, er = rows[geo]
        wr = word_start_row + offset
        excel_cols = [5, 6, 7] if metric == "nombre" else [2, 3, 4]

        for wc, ec in zip([2, 3, 4], excel_cols):
            set_cell(
                table.cell(wr, wc),
                value(ws, er, ec, decimals=2)
            )
    return True

def set_years_three(table):
    for row in [0, 1]:
        for col, year in zip([2, 3, 4], ["2024", "2023", "2022"]):
            set_cell(table.cell(row, col), year)

input_docx = None
for candidate in CANDIDATS_ENTREE:
    if candidate.exists():
        input_docx = candidate
        break

if input_docx is None:
    raise FileNotFoundError(
        "Fichier Word Culture introuvable dans sorties."
    )

if not FICHIER_EXCEL.exists():
    raise FileNotFoundError(
        f"Excel Foncier introuvable : {FICHIER_EXCEL}"
    )

print("=" * 96)
print("GENERATION HABITAT / CONSTRUCTION / FONCIER")
print("=" * 96)
print(f"Word d'entrée : {input_docx}")
print()

doc = Document(input_docx)
wb = load_workbook(FICHIER_EXCEL, data_only=True, read_only=False)

print("[1/13] 13-1  conservé — déjà RGPH 2024")
print("[2/13] 13-4  conservé — déjà RGPH 2024")
print("[3/13] 13-6  conservé — déjà RGPH 2024")
print("[4/13] 13-7  conservé — déjà RGPH 2024")
print("[5/13] 13-8  conservé — déjà RGPH 2024")
print("[6/13] 13-9  conservé — aucune nouvelle source correspondante")
print("[7/13] 13-10 conservé — aucune nouvelle source correspondante")

print("[8/13] 13-11 Réquisitions", flush=True)
t = find_word_data_table(doc, "13-11")
urb = sheets_matching(wb, ["requisitions", "deposees"], mode="urbain_rural")
rur = sheets_matching(wb, ["requisitions", "deposees"], mode="milieu_rural")

ok_u = False
ok_r = False
if urb:
    set_years_three(t)
    ok_u = fill_number_surface_section(t, 2, urb)
if rur:
    ok_r = fill_number_surface_section(t, 11, rur)
print(
    f"   Urbain+rural : {'OK' if ok_u else 'CONSERVE'} | "
    f"Milieu rural : {'OK' if ok_r else 'CONSERVE'}"
)

print("[9/13] 13-12 Immatriculations", flush=True)
t = find_word_data_table(doc, "13-12")
urb = sheets_matching(wb, ["immatric"], mode="urbain_rural")
rur = sheets_matching(wb, ["immatric"], mode="milieu_rural")

updated = False
if urb:
    set_years_three(t)
    updated = fill_number_surface_section(t, 2, urb) or updated
if rur:
    updated = fill_number_surface_section(t, 11, rur) or updated

if updated:
    print("   -> Actualisé avec les sources 'immatriculations' détectées.")
else:
    print(
        "   -> Aucune source Excel sémantiquement validée "
        "pour les immatriculations : tableau conservé."
    )

print("[10/13] 13-13 Morcellement", flush=True)
t = find_word_data_table(doc, "13-13")
urb = sheets_matching(wb, ["morcellement"], mode="urbain_rural")
rur = sheets_matching(wb, ["morcellement"], mode="milieu_rural")

ok_u = False
ok_r = False
if urb:
    set_years_three(t)
    ok_u = fill_number_surface_section(t, 2, urb)
if rur:
    ok_r = fill_number_surface_section(t, 11, rur)

print(
    f"   Urbain+rural : {'OK' if ok_u else 'CONSERVE'} | "
    f"Milieu rural : {'OK' if ok_r else 'CONSERVE'}"
)

print("[11/13] 13-14 Hypothèques", flush=True)
t = find_word_data_table(doc, "13-14")
sheets = sheets_matching(wb, ["hypotheques"], mode="urbain_rural")
if not sheets:
    sheets = sheets_matching(wb, ["hypotheques"])
rows = geo_rows(sheets) if sheets else {}

if sheets and all_geo_found(rows):
    for row in [0, 1]:
        set_cell(t.cell(row, 2), "2024")
        set_cell(t.cell(row, 3), "2023")

    order = [
        ("nation", "nombre"),
        ("nation", "valeur"),
        ("region", "nombre"),
        ("region", "valeur"),
        ("menzeh", "nombre"),
        ("menzeh", "valeur"),
        ("ismailia", "nombre"),
        ("ismailia", "valeur"),
    ]

    for offset, (geo, metric) in enumerate(order):
        ws, er = rows[geo]
        wr = 2 + offset
        cols = [5, 6] if metric == "nombre" else [2, 3]
        for wc, ec in zip([2, 3], cols):
            set_cell(
                t.cell(wr, wc),
                value(ws, er, ec, decimals=2)
            )
    print("   -> 2024 / 2023 actualisés.")
else:
    print("   -> Source non validée : conservé.")

print("[12/13] 13-15 Recettes", flush=True)
t = find_word_data_table(doc, "13-15")
sheets = sheets_matching(wb, ["recettes", "realisees"])
rows = geo_rows(sheets) if sheets else {}

if sheets and all_geo_found(rows):
    for col, year in zip([1, 2, 3], ["2024", "2023", "2022"]):
        set_cell(t.cell(0, col), year)

    for geo, wr in {
        "nation": 1,
        "region": 2,
        "menzeh": 3,
        "ismailia": 4,
    }.items():
        ws, er = rows[geo]
        for wc, ec in zip([1, 2, 3], [2, 3, 4]):
            set_cell(
                t.cell(wr, wc),
                value(ws, er, ec, decimals=0)
            )

    replace_2023_source_with_2025(t)
    print("   -> 2024 / 2023 / 2022 actualisés.")
else:
    print("   -> Source non validée : conservé.")

print("[13/13] 13-16 Plans cadastraux", flush=True)
t = find_word_data_table(doc, "13-16")

urb = sheets_matching(wb, ["plans", "cadastraux"], mode="urbain_rural")
rur = sheets_matching(wb, ["plans", "cadastraux"], mode="milieu_rural")

def fill_plans_section(table, start_row, sheets):
    rows = geo_rows(sheets)
    if not all_geo_found(rows):
        return False

    year_cols = None
    for ws in sheets:
        for r in range(1, min(ws.max_row, 15) + 1):
            found = {}
            for c in range(1, min(ws.max_column, 12) + 1):
                v = str(ws.cell(r, c).value or "").strip()
                if v in ("2024", "2023", "2022"):
                    found[v] = c
            if all(y in found for y in ("2024", "2023", "2022")):
                year_cols = [found["2024"], found["2023"], found["2022"]]
                break
        if year_cols:
            break

    if not year_cols:
        return False

    for geo, wr in {
        "nation": start_row,
        "region": start_row + 1,
        "menzeh": start_row + 2,
        "ismailia": start_row + 3,
    }.items():
        ws, er = rows[geo]
        for wc, ec in zip([1, 2, 3], year_cols):
            set_cell(
                table.cell(wr, wc),
                value(ws, er, ec, decimals=0)
            )
    return True

for col, year in zip([1, 2, 3], ["2024", "2023", "2022"]):
    set_cell(t.cell(0, col), year)

ok_u = fill_plans_section(t, 1, urb) if urb else False
ok_r = fill_plans_section(t, 6, rur) if rur else False

print(
    f"   Urbain+rural : {'OK' if ok_u else 'CONSERVE'} | "
    f"Milieu rural : {'OK' if ok_r else 'CONSERVE'}"
)

wb.close()

print()
print("Sauvegarde du Word...", flush=True)
doc.save(FICHIER_SORTIE)

print()
print("=" * 96)
print("HABITAT / CONSTRUCTION / FONCIER TERMINE")
print("=" * 96)
print("13-1, 13-4, 13-6, 13-7, 13-8 : conservés (déjà 2024)")
print("13-9, 13-10 : conservés (pas de nouvelle source correspondante)")
print("13-11 à 13-16 : mise à jour seulement après validation sémantique")
print("                 du titre Excel et du milieu Urbain/Rural.")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 96)
