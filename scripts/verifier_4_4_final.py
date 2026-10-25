from pathlib import Path
from zipfile import ZipFile
from lxml import etree
import re
import unicodedata

DOCX = Path("sorties") / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx"
OUT = Path("sorties") / "verification_4_4_finale.txt"

TARGETS = {
    "Cité Universitaire Meknès - Total": "2668",
    "Cité Universitaire Meknès - Féminin": "1908",
    "Cité Universitaire Meknès - Etrangers": "61",
    "Nation - Total": "60915",
    "Nation - Féminin": "38491",
    "Nation - Etrangers": "1487",
}

def norm(s):
    s = str(s or "").replace("\u00a0", " ").replace("\u202f", " ")
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", s).strip().lower()

def digits(s):
    return re.sub(r"\D", "", str(s or ""))

if not DOCX.exists():
    raise FileNotFoundError(DOCX)

paragraphs = []

with ZipFile(DOCX, "r") as zf:
    for name in zf.namelist():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue

        try:
            root = etree.fromstring(zf.read(name))
        except Exception:
            continue

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        for p in root.xpath(".//w:p", namespaces=ns):
            parts = p.xpath(".//w:t/text()", namespaces=ns)
            if not parts:
                continue

            text = "".join(parts)
            paragraphs.append((name, text))

report = []
report.append("=" * 110)
report.append("VERIFICATION FINALE 4-4 — CITE UNIVERSITAIRE")
report.append("=" * 110)
report.append(f"Word : {DOCX}")
report.append("")

results = {}

for label, target in TARGETS.items():
    matches = []

    for idx, (name, text) in enumerate(paragraphs):
        d = digits(text)
        if target in d:
            matches.append((idx, name, text))

    results[label] = matches

    report.append(f"{label} [{target}] : {'PRESENT' if matches else 'ABSENT'}")

    for idx, name, text in matches[:10]:
        report.append(f"  {name} | P{idx+1} | {text!r}")

report.append("")
report.append("-" * 110)
report.append("CONTEXTE AUTOUR DE 'cité universitaire' / 'meknès'")
report.append("-" * 110)

for idx, (name, text) in enumerate(paragraphs):
    n = norm(text)

    if "cite universitaire" in n or "meknes" in n:
        start = max(0, idx - 4)
        end = min(len(paragraphs), idx + 8)

        report.append("")
        report.append(f"CONTEXTE {name} autour P{idx+1}")

        for j in range(start, end):
            report.append(f"  P{j+1}: {paragraphs[j][1]!r}")

missing = [label for label, matches in results.items() if not matches]

report.append("")
report.append("=" * 110)

if not missing:
    report.append("VERDICT : LES 6 VALEURS ATTENDUES SONT PRESENTES DANS LE XML WORD.")
else:
    report.append("VERDICT : AU MOINS UNE VALEUR ATTENDUE N'A PAS ETE RETROUVEE.")
    for label in missing:
        report.append(f"  MANQUANTE : {label}")

report.append("=" * 110)

OUT.write_text("\n".join(report), encoding="utf-8")

print("=" * 100)
print("VERIFICATION 4-4 TERMINEE")
print("=" * 100)

for label, matches in results.items():
    print(f"{label}: {'PRESENT' if matches else 'ABSENT'}")

print()
print(f"Rapport créé : {OUT}")
print("=" * 100)
