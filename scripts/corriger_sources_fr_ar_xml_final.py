from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree


DOSSIER_SORTIE = Path("sorties")

CANDIDATS = [
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_SOURCES_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_VALIDE.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_SOURCES_FINAL.docx"
)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def texte_paragraphe(p):
    return "".join(
        p.xpath(".//w:t/text()", namespaces=NS)
    ).strip()


def get_or_add(parent, tag):
    el = parent.find(f"w:{tag}", namespaces=NS)
    if el is None:
        el = etree.Element(f"{{{NS['w']}}}{tag}")
        parent.append(el)
    return el


def aligner_paragraphe(p, direction):
    """
    direction = 'right' ou 'left'
    Travaille directement en OOXML : fonctionne aussi dans les zones de texte.
    """
    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is None:
        pPr = etree.Element(f"{{{NS['w']}}}pPr")
        p.insert(0, pPr)

    jc = get_or_add(pPr, "jc")
    jc.set(f"{{{NS['w']}}}val", direction)

    bidi = pPr.find("w:bidi", namespaces=NS)

    if direction == "right":
        if bidi is None:
            bidi = etree.Element(f"{{{NS['w']}}}bidi")
            pPr.append(bidi)
        bidi.set(f"{{{NS['w']}}}val", "1")

        # RTL également au niveau des runs du paragraphe arabe.
        for r in p.xpath(".//w:r", namespaces=NS):
            rPr = r.find("w:rPr", namespaces=NS)
            if rPr is None:
                rPr = etree.Element(f"{{{NS['w']}}}rPr")
                r.insert(0, rPr)

            rtl = rPr.find("w:rtl", namespaces=NS)
            if rtl is None:
                rtl = etree.Element(f"{{{NS['w']}}}rtl")
                rPr.append(rtl)
            rtl.set(f"{{{NS['w']}}}val", "1")

    else:
        if bidi is not None:
            pPr.remove(bidi)


def est_source_arabe(t):
    return (
        "المصدر" in t
        or "النشرة الإحصائية السنوية للمغرب" in t
    )


def est_source_francaise(t):
    tl = t.lower()
    return (
        tl.startswith("source")
        or "annuaire statistique du maroc" in tl
    )


fichier_entree = None

for candidat in CANDIDATS:
    if candidat.exists():
        fichier_entree = candidat
        break

if fichier_entree is None:
    raise FileNotFoundError(
        "Aucun document Santé récent trouvé dans le dossier sorties."
    )


print("=" * 90)
print("CORRECTION OOXML GLOBALE DES SOURCES FR / AR")
print("=" * 90)
print(f"Document d'entrée : {fichier_entree}")
print()

nb_fr = 0
nb_ar = 0

with ZipFile(fichier_entree, "r") as zin:
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)

    # .//w:p inclut aussi les paragraphes dans les cellules, zones de texte,
    # formes et objets Word présents dans document.xml.
    for p in root.xpath(".//w:p", namespaces=NS):
        t = texte_paragraphe(p)

        if not t:
            continue

        if est_source_arabe(t):
            aligner_paragraphe(p, "right")
            nb_ar += 1

        elif est_source_francaise(t):
            aligner_paragraphe(p, "left")
            nb_fr += 1

    nouveau_xml = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )

    with ZipFile(FICHIER_SORTIE, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "word/document.xml":
                data = nouveau_xml

            zout.writestr(item, data)


print(f"Sources françaises corrigées : {nb_fr}")
print(f"Sources arabes corrigées      : {nb_ar}")
print()
print("Aucune donnée numérique n'a été modifiée.")
print("Aucun texte de source n'a été remplacé ou recréé.")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 90)
