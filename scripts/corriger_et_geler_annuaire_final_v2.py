from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import re
import unicodedata
import shutil
import tempfile

SORTIES = Path("sorties")

DONOR = SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"
TARGET = SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx"
OUTPUT = SORTIES / "Annuaire_HCP_2025_FINAL.docx"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

EXPECTED_44 = {
    "Meknès Total": "2668",
    "Meknès Féminin": "1908",
    "Meknès Etrangers": "61",
    "Nation Total": "60915",
    "Nation Féminin": "38491",
    "Nation Etrangers": "1487",
}

def norm(s):
    s = str(s or "").replace("\u00a0", " ").replace("\u202f", " ")
    s = s.replace("’", "'").replace("–", "-").replace("—", "-")
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", s).strip().lower()

def child_text(el):
    return "".join(el.xpath(".//w:t/text()", namespaces=NS))

def paragraph_numeric_values(elements):
    vals = []
    for el in elements:
        for p in el.xpath(".//w:p", namespaces=NS):
            text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
            d = re.sub(r"\D", "", text)
            if d:
                vals.append(d)
    return vals

def find_44_bounds(root):
    body = root.find(f".//{W}body")
    if body is None:
        raise RuntimeError("w:body introuvable.")

    children = list(body)
    start = None
    end = None

    for i, el in enumerate(children):
        t = norm(child_text(el))

        if start is None:
            if (
                re.search(r"\b4\s*-\s*4\b", t)
                and "effectif des etudiants residant" in t
                and "cite" in t
            ):
                start = i
                continue

        if start is not None:
            if (
                re.search(r"\b4\s*-\s*5\b", t)
                and "effectif des laureats" in t
            ):
                end = i
                break

    if start is None or end is None or end <= start:
        raise RuntimeError(
            f"Bloc 4-4 impossible à délimiter : start={start}, end={end}"
        )

    return body, start, end

def collect_rids(elements):
    ids = set()
    for el in elements:
        for node in el.iter():
            for attr, val in node.attrib.items():
                if attr.startswith(R) and attr.endswith("id"):
                    ids.add(val)
    return ids

def rel_ids_from_zip(zf):
    name = "word/_rels/document.xml.rels"
    root = etree.fromstring(zf.read(name))
    return {el.get("Id") for el in root if el.get("Id")}

def _replace_across_text_nodes(nodes, old, new):
    """
    Remplace une chaîne même si Word l'a découpée sur plusieurs w:t.
    La distribution des caractères est conservée autant que possible
    pour ne pas casser le formatage des runs.
    """
    if not nodes:
        return 0

    original_parts = [node.text or "" for node in nodes]
    combined = "".join(original_parts)

    if old not in combined:
        return 0

    replaced = combined.replace(old, new)
    lengths = [len(x) for x in original_parts]

    pos = 0
    for i, node in enumerate(nodes):
        if i < len(nodes) - 1:
            take = lengths[i]
            node.text = replaced[pos:pos + take]
            pos += take
        else:
            node.text = replaced[pos:]

    return combined.count(old)


def patch_annuaire_text(root):
    """
    Corrige les textes visibles au niveau PARAGRAPHE, afin de gérer les
    chaînes coupées entre plusieurs runs XML.
    """
    fr = 0
    ar = 0
    typo = 0

    for p in root.xpath(".//w:p", namespaces=NS):
        nodes = p.xpath(".//w:t", namespaces=NS)

        if not nodes:
            continue

        # 2023 -> 2025 uniquement pour les sources Annuaire.
        fr += _replace_across_text_nodes(
            nodes,
            "Annuaire Statistique du Maroc 2023",
            "Annuaire Statistique du Maroc 2025"
        )

        ar += _replace_across_text_nodes(
            nodes,
            "النشرة الإحصائية السنوية للمغرب 2023",
            "النشرة الإحصائية السنوية للمغرب 2025"
        )

        # Certaines occurrences arabes n'ont pas d'espace avant l'année.
        ar += _replace_across_text_nodes(
            nodes,
            "النشرة الإحصائية السنوية للمغرب2023",
            "النشرة الإحصائية السنوية للمغرب2025"
        )

        # Supprime l'artefact même s'il est dans un run séparé.
        typo += _replace_across_text_nodes(
            nodes,
            "cccccccc",
            ""
        )

    return fr, ar, typo

def validate_44(root, label):
    body, start, end = find_44_bounds(root)
    elems = list(body)[start:end]
    vals = paragraph_numeric_values(elems)

    missing = {
        name: v for name, v in EXPECTED_44.items()
        if v not in vals
    }

    print(f"{label} — valeurs numériques détectées dans 4-4 :")
    for name, v in EXPECTED_44.items():
        print(f"   {name}: {'OK' if v in vals else 'ABSENT'}")

    return missing

if not DONOR.exists():
    raise FileNotFoundError(
        f"Donneur Education corrigé introuvable : {DONOR}"
    )

if not TARGET.exists():
    raise FileNotFoundError(
        f"Word final intermédiaire introuvable : {TARGET}"
    )

print("=" * 104)
print("CORRECTION FINALE V2 — RESTAURATION 4-4 + SOURCES ANNUAIRE")
print("=" * 104)
print(f"Donneur 4-4 : {DONOR}")
print(f"Cible : {TARGET}")
print()

with ZipFile(DONOR, "r") as zd:
    donor_doc = etree.fromstring(zd.read("word/document.xml"))
    donor_body, ds, de = find_44_bounds(donor_doc)
    donor_block = [
        deepcopy(el)
        for el in list(donor_body)[ds:de]
    ]

    missing_donor = validate_44(donor_doc, "DONNEUR")

    if missing_donor:
        raise RuntimeError(
            "Le fichier donneur ne contient pas toutes les valeurs 4-4 attendues : "
            + ", ".join(missing_donor.keys())
        )

    donor_rids = collect_rids(donor_block)

with ZipFile(TARGET, "r") as zt:
    target_doc = etree.fromstring(zt.read("word/document.xml"))
    target_body, ts, te = find_44_bounds(target_doc)

    print()
    missing_target = validate_44(target_doc, "CIBLE AVANT CORRECTION")

    # Sécurité : si le bloc donneur utilise des relations, vérifier
    # qu'elles existent déjà dans la cible (documents de même lignée).
    target_rel_ids = rel_ids_from_zip(zt)
    missing_rids = sorted(donor_rids - target_rel_ids)

    if missing_rids:
        raise RuntimeError(
            "Le bloc 4-4 donneur référence des relations absentes de la cible : "
            + ", ".join(missing_rids)
        )

    # Supprimer uniquement le bloc 4-4 cible.
    children = list(target_body)
    for el in children[ts:te]:
        target_body.remove(el)

    # Insérer le bloc 4-4 validé du fichier Education corrigé.
    for offset, el in enumerate(donor_block):
        target_body.insert(ts + offset, deepcopy(el))

    # Corriger les sources Annuaire dans document.xml.
    fr_doc, ar_doc, typo_doc = patch_annuaire_text(target_doc)

    new_document_xml = etree.tostring(
        target_doc,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )

    # Réécrire le DOCX sans altérer les autres fichiers.
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as zout:
        total_fr = fr_doc
        total_ar = ar_doc
        total_typo = typo_doc

        for item in zt.infolist():
            data = zt.read(item.filename)

            if item.filename == "word/document.xml":
                data = new_document_xml

            elif item.filename.startswith("word/") and item.filename.endswith(".xml"):
                try:
                    root = etree.fromstring(data)
                    fr, ar, typo = patch_annuaire_text(root)
                    total_fr += fr
                    total_ar += ar
                    total_typo += typo

                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone="yes"
                    )
                except Exception:
                    pass

            zout.writestr(item, data)

print()
print("Bloc 4-4 restauré depuis le Word Education corrigé.")
print(f"Sources FR Annuaire 2023 corrigées : {total_fr}")
print(f"Sources AR Annuaire 2023 corrigées : {total_ar}")
print(f"Artefacts 'cccccccc' corrigés : {total_typo}")

# ============================================================================
# VALIDATION FINALE DU FICHIER GENERE
# ============================================================================

with ZipFile(OUTPUT, "r") as zf:
    final_doc = etree.fromstring(zf.read("word/document.xml"))

    print()
    missing_final = validate_44(final_doc, "FINAL")

    old_fr = 0
    old_ar = 0
    typo_left = 0

    for name in zf.namelist():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue

        try:
            root = etree.fromstring(zf.read(name))
        except Exception:
            continue

        for p in root.xpath(".//w:p", namespaces=NS):
            text = "".join(p.xpath(".//w:t/text()", namespaces=NS))

            if "Annuaire Statistique du Maroc 2023" in text:
                old_fr += 1

            if (
                "النشرة الإحصائية السنوية للمغرب 2023" in text
                or "النشرة الإحصائية السنوية للمغرب2023" in text
            ):
                old_ar += 1

            if "cccccccc" in text:
                typo_left += 1

    if missing_final:
        raise RuntimeError(
            "ECHEC FINAL 4-4 : "
            + ", ".join(missing_final.keys())
        )

    if old_fr or old_ar or typo_left:
        raise RuntimeError(
            f"ECHEC FINAL SOURCES : FR2023={old_fr}, AR2023={old_ar}, "
            f"cccccccc={typo_left}"
        )

print()
print("=" * 104)
print("VALIDATION AUTOMATIQUE FINALE : OK")
print("=" * 104)
print("4-4 : 6/6 valeurs attendues présentes")
print("Annuaire 2023 restant : 0")
print("Artefact cccccccc restant : 0")
print()
print(f"Fichier final créé : {OUTPUT}")
print("=" * 104)
