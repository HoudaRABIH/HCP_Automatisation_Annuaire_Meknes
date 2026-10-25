from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import sys
import re

SORTIES = Path("sorties")
REFERENCES = Path("references")

INPUT = SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante.docx"
REFERENCE = REFERENCES / "REFERENCE_EDUCATION_2025.docx"
OUTPUT = SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}
W = "{" + W_NS + "}"
R = "{" + R_NS + "}"


def text_of(el):
    return "".join(el.xpath(".//w:t/text()", namespaces=NS))


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def body_and_bounds(root):
    body = root.find(f".//{W}body")
    if body is None:
        raise RuntimeError("w:body introuvable.")

    children = list(body)

    # Il existe une occurrence 3-1 dans la table des matières.
    # L'occurrence réelle du chapitre est la DERNIERE occurrence avant 6-1.
    starts = []
    for i, el in enumerate(children):
        t = norm(text_of(el))
        if "3-1" in t and "taux de scolarisation" in t:
            starts.append(i)

    if not starts:
        raise RuntimeError("Début réel du bloc Enseignement (3-1) introuvable.")

    start = starts[-1]

    end = None
    for i in range(start + 1, len(children)):
        t = norm(text_of(children[i]))
        if "6-1" in t and (
            "biblioth" in t
            or "culture" in t
            or "loisirs" in t
        ):
            end = i
            break

    # Fallback: premier vrai 6-1 après l'enseignement.
    if end is None:
        for i in range(start + 1, len(children)):
            t = norm(text_of(children[i]))
            if re.search(r"\b6\s*-\s*1\b", t):
                end = i
                break

    if end is None or end <= start:
        raise RuntimeError(
            f"Fin du bloc Enseignement introuvable : start={start}, end={end}"
        )

    return body, start, end


def collect_relationship_ids(elements):
    ids = set()
    for el in elements:
        for node in el.iter():
            for attr, val in node.attrib.items():
                if attr.startswith(R) and attr.endswith("id"):
                    ids.add(val)
    return ids


def relationship_ids(zf):
    name = "word/_rels/document.xml.rels"
    root = etree.fromstring(zf.read(name))
    return {el.get("Id") for el in root if el.get("Id")}


if not INPUT.exists():
    raise FileNotFoundError(f"Word Santé introuvable : {INPUT}")

if not REFERENCE.exists():
    raise FileNotFoundError(
        f"Référence Enseignement validée introuvable : {REFERENCE}"
    )

print("=" * 96)
print("ENSEIGNEMENT + ENSEIGNEMENT SUPERIEUR — APPLICATION DU BLOC VALIDE")
print("=" * 96)
print(f"Entrée : {INPUT}")
print(f"Référence : {REFERENCE}")
print()

with ZipFile(REFERENCE, "r") as zr:
    ref_doc = etree.fromstring(zr.read("word/document.xml"))
    ref_body, rs, re_ = body_and_bounds(ref_doc)

    block = [
        deepcopy(el)
        for el in list(ref_body)[rs:re_]
    ]

    if not block:
        raise RuntimeError("Bloc Enseignement de référence vide.")

    donor_rids = collect_relationship_ids(block)

with ZipFile(INPUT, "r") as zi:
    target_doc = etree.fromstring(zi.read("word/document.xml"))
    target_body, ts, te = body_and_bounds(target_doc)

    target_rids = relationship_ids(zi)
    missing = sorted(donor_rids - target_rids)

    if missing:
        raise RuntimeError(
            "Le bloc Enseignement utilise des relations absentes du Word initial : "
            + ", ".join(missing)
        )

    children = list(target_body)

    for el in children[ts:te]:
        target_body.remove(el)

    for offset, el in enumerate(block):
        target_body.insert(ts + offset, deepcopy(el))

    new_xml = etree.tostring(
        target_doc,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )

    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zo.writestr(item, data)

# Sanity validation
with ZipFile(OUTPUT, "r") as zf:
    final_root = etree.fromstring(zf.read("word/document.xml"))
    text = norm(text_of(final_root))

checks = [
    "3-1",
    "3-21",
    "4 - 1",
    "4 - 4",
    "4 - 5",
]

missing_checks = [x for x in checks if x not in text]
if missing_checks:
    raise RuntimeError(
        "Validation Enseignement échouée : "
        + ", ".join(missing_checks)
    )

print("Enseignement général : OK")
print("Enseignement supérieur : OK")
print("Bloc complet 3-x / 4-x / 5-x restauré depuis la référence validée.")
print(f"Document créé : {OUTPUT}")
print("=" * 96)
