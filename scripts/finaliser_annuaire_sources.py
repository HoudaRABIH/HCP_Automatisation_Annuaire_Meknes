from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree


DOSSIER_SORTIE = Path("sorties")

# On prend automatiquement le fichier Santé le plus récent.
CANDIDATS = [
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_SOURCES_FINAL.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_SOURCES_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_VALIDE.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante_OK.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN.docx"
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag):
    return f"{{{W}}}{tag}"


def texte_paragraphe(p):
    return "".join(
        t.text or ""
        for t in p.xpath(".//w:t", namespaces=NS)
    ).strip()


def est_arabe(texte):
    return any("\u0600" <= c <= "\u06ff" for c in texte)


def est_source_annuaire_fr(texte):
    t = " ".join((texte or "").split()).lower()
    return (
        "annuaire statistique du maroc" in t
        or t.startswith("source : annuaire statistique")
        or t.startswith("source: annuaire statistique")
    )


def est_source_annuaire_ar(texte):
    t = " ".join((texte or "").split())
    return (
        "المصدر" in t
        and (
            "النشرة الإحصائية السنوية للمغرب" in t
            or "الإحصائية السنوية للمغرب" in t
        )
    )


def get_or_add_ppr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        pPr = etree.Element(qn("pPr"))
        p.insert(0, pPr)

    return pPr


def get_or_add_rpr(r):
    rPr = r.find("w:rPr", namespaces=NS)

    if rPr is None:
        rPr = etree.Element(qn("rPr"))
        r.insert(0, rPr)

    return rPr


def set_alignment_and_direction(p, arabe):
    """
    Corrige la direction VISUELLE sans modifier le texte :
    FR = gauche + LTR
    AR = droite + RTL
    """
    pPr = get_or_add_ppr(p)

    jc = pPr.find("w:jc", namespaces=NS)
    if jc is None:
        jc = etree.Element(qn("jc"))
        pPr.append(jc)

    jc.set(
        qn("val"),
        "right" if arabe else "left"
    )

    bidi = pPr.find("w:bidi", namespaces=NS)

    if arabe:
        if bidi is None:
            bidi = etree.Element(qn("bidi"))
            pPr.append(bidi)

        bidi.set(qn("val"), "1")

    elif bidi is not None:
        pPr.remove(bidi)

    # Direction de chaque run.
    for r in p.xpath(".//w:r", namespaces=NS):
        rPr = get_or_add_rpr(r)

        rtl = rPr.find("w:rtl", namespaces=NS)

        if arabe:
            if rtl is None:
                rtl = etree.Element(qn("rtl"))
                rPr.append(rtl)

            rtl.set(qn("val"), "1")

        elif rtl is not None:
            rPr.remove(rtl)


def trouver_ancetre(element, nom_tag):
    parent = element.getparent()

    while parent is not None:
        if parent.tag == qn(nom_tag):
            return parent
        parent = parent.getparent()

    return None


def cellules_directes_de_ligne(tr):
    return tr.findall("w:tc", namespaces=NS)


def deplacer_paragraphe_dans_cellule(p, tc_cible):
    parent = p.getparent()

    if parent is tc_cible:
        return False

    parent.remove(p)

    # Ajouter avant les propriétés de cellule finales éventuelles n'est
    # pas nécessaire : les paragraphes sont des enfants normaux de w:tc.
    tc_cible.append(p)

    return True


fichier_entree = None

for candidat in CANDIDATS:
    if candidat.exists():
        fichier_entree = candidat
        break

if fichier_entree is None:
    raise FileNotFoundError(
        "Aucun fichier Santé récent trouvé dans le dossier sorties."
    )


print("=" * 94)
print("NETTOYAGE FINAL DES SOURCES ANNUAIRE FR / AR")
print("=" * 94)
print(f"Document d'entrée : {fichier_entree}")
print()


with ZipFile(fichier_entree, "r") as zin:
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)

    paragraphes = root.xpath(".//w:p", namespaces=NS)

    sources_fr = []
    sources_ar = []

    for p in paragraphes:
        texte = texte_paragraphe(p)

        if est_source_annuaire_fr(texte):
            sources_fr.append(p)

        elif est_source_annuaire_ar(texte):
            sources_ar.append(p)

    deplacements_fr = 0
    deplacements_ar = 0

    # --------------------------------------------------------
    # FRANÇAIS : première cellule de la ligne + LTR/gauche
    # --------------------------------------------------------
    for p in sources_fr:
        set_alignment_and_direction(
            p,
            arabe=False
        )

        tc = trouver_ancetre(
            p,
            "tc"
        )

        tr = trouver_ancetre(
            p,
            "tr"
        )

        if tc is None or tr is None:
            continue

        cellules = cellules_directes_de_ligne(
            tr
        )

        if len(cellules) >= 2:
            tc_gauche = cellules[0]

            if deplacer_paragraphe_dans_cellule(
                p,
                tc_gauche
            ):
                deplacements_fr += 1

    # --------------------------------------------------------
    # ARABE : dernière cellule de la ligne + RTL/droite
    # --------------------------------------------------------
    for p in sources_ar:
        set_alignment_and_direction(
            p,
            arabe=True
        )

        tc = trouver_ancetre(
            p,
            "tc"
        )

        tr = trouver_ancetre(
            p,
            "tr"
        )

        if tc is None or tr is None:
            continue

        cellules = cellules_directes_de_ligne(
            tr
        )

        if len(cellules) >= 2:
            tc_droite = cellules[-1]

            if deplacer_paragraphe_dans_cellule(
                p,
                tc_droite
            ):
                deplacements_ar += 1

    # Une deuxième passe de direction après les déplacements.
    for p in sources_fr:
        set_alignment_and_direction(
            p,
            arabe=False
        )

    for p in sources_ar:
        set_alignment_and_direction(
            p,
            arabe=True
        )

    nouveau_xml = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )

    with ZipFile(
        FICHIER_SORTIE,
        "w",
        ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(
                item.filename
            )

            if item.filename == "word/document.xml":
                data = nouveau_xml

            zout.writestr(
                item,
                data
            )


print(f"Sources Annuaire FR trouvées : {len(sources_fr)}")
print(f"Sources Annuaire AR trouvées : {len(sources_ar)}")
print(f"Sources FR déplacées à gauche : {deplacements_fr}")
print(f"Sources AR déplacées à droite : {deplacements_ar}")
print()
print("Le texte des sources n'a PAS été réécrit.")
print("Les données numériques n'ont PAS été modifiées.")
print("Les autres sources (RGPH, ministères, HCP, etc.) n'ont PAS été touchées.")
print()
print(f"Document final créé : {FICHIER_SORTIE}")
print("=" * 94)
