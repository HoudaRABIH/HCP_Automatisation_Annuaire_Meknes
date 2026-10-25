from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import re


DOSSIER_SORTIE = Path("sorties")

# IMPORTANT:
# On repart du fichier JUSTICE propre, AVANT les scripts qui ont créé des doublons.
FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_FIXED.docx"
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag):
    return f"{{{W}}}{tag}"


def texte_paragraphe(p):
    return "".join(
        p.xpath(".//w:t/text()", namespaces=NS)
    ).strip()


def normaliser(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def est_source_fr(t):
    low = normaliser(t).lower()
    compact = re.sub(r"\s+", "", low)

    return (
        "2025" in low
        and (
            "annuairestatistiquedumaroc" in compact
            or (
                "source" in low
                and "annuaire" in low
                and "maroc" in low
            )
        )
    )


def est_source_ar(t):
    s = normaliser(t)

    return (
        "2025" in s
        and (
            "النشرة الإحصائية السنوية للمغرب" in s
            or (
                "المصدر" in s
                and "الإحصائية" in s
                and "المغرب" in s
            )
        )
    )


def get_sectpr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        return None

    sect = pPr.find("w:sectPr", namespaces=NS)

    return deepcopy(sect) if sect is not None else None


def remove_sectpr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        return

    sect = pPr.find("w:sectPr", namespaces=NS)

    if sect is not None:
        pPr.remove(sect)


def add_sectpr(p, sect):
    if sect is None:
        return

    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        pPr = etree.Element(qn("pPr"))
        p.insert(0, pPr)

    old = pPr.find("w:sectPr", namespaces=NS)

    if old is not None:
        pPr.remove(old)

    pPr.append(deepcopy(sect))


def trouver_modeles_8_10(body):
    enfants = list(body)

    index_8_10 = None

    for i, el in enumerate(enfants):
        t = normaliser(
            " ".join(
                el.xpath(".//w:t/text()", namespaces=NS)
            )
        )

        if t.startswith(
            "8 - 10 Activités des sections des juridictions de proximité"
        ):
            index_8_10 = i
            break

    if index_8_10 is None:
        raise RuntimeError(
            "Impossible de localiser 8-10 pour prendre son style de source."
        )

    modele_fr = None
    modele_ar = None

    for el in enfants[index_8_10 + 1:]:
        t = normaliser(
            " ".join(
                el.xpath(".//w:t/text()", namespaces=NS)
            )
        )

        if t.startswith("8 - 11"):
            break

        if el.tag != qn("p"):
            continue

        if est_source_fr(t) and not est_source_ar(t):
            modele_fr = deepcopy(el)

        if est_source_ar(t) and not est_source_fr(t):
            modele_ar = deepcopy(el)

    if modele_fr is None or modele_ar is None:
        raise RuntimeError(
            "Sources FR/AR propres de 8-10 introuvables."
        )

    remove_sectpr(modele_fr)
    remove_sectpr(modele_ar)

    return modele_fr, modele_ar


def construire_source_mixte(orig, modele_fr, modele_ar):
    """
    Cas où FR + AR sont actuellement dans UN SEUL paragraphe.
    On garde UN SEUL paragraphe pour ne pas modifier la pagination.
    """

    p = deepcopy(orig)

    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        pPr = etree.Element(qn("pPr"))
        p.insert(0, pPr)

    # Paragraphe LTR
    bidi = pPr.find("w:bidi", namespaces=NS)

    if bidi is not None:
        pPr.remove(bidi)

    jc = pPr.find("w:jc", namespaces=NS)

    if jc is None:
        jc = etree.SubElement(pPr, qn("jc"))

    jc.set(qn("val"), "left")

    # Même zone horizontale que la source 8-10.
    ind = pPr.find("w:ind", namespaces=NS)

    if ind is None:
        ind = etree.SubElement(pPr, qn("ind"))

    for attr in list(ind.attrib):
        ind.attrib.pop(attr)

    ind.set(qn("left"), "497")
    ind.set(qn("right"), "497")

    # Tabulation à droite pour la source arabe.
    tabs = pPr.find("w:tabs", namespaces=NS)

    if tabs is not None:
        pPr.remove(tabs)

    tabs = etree.Element(qn("tabs"))

    tab = etree.SubElement(
        tabs,
        qn("tab")
    )
    tab.set(qn("val"), "right")
    tab.set(qn("pos"), "9000")

    pPr.insert(1, tabs)

    if pPr.find("w:keepLines", namespaces=NS) is None:
        etree.SubElement(
            pPr,
            qn("keepLines")
        )

    # Garder uniquement pPr.
    for child in list(p):
        if child is not pPr:
            p.remove(child)

    # Style FR depuis 8-10.
    fr_rpr = modele_fr.find(
        "w:r/w:rPr",
        namespaces=NS
    )

    r_fr = etree.SubElement(
        p,
        qn("r")
    )

    if fr_rpr is not None:
        r_fr.append(
            deepcopy(fr_rpr)
        )

    t_fr = etree.SubElement(
        r_fr,
        qn("t")
    )
    t_fr.text = (
        "Source : Annuaire Statistique du Maroc 2025"
    )

    # Tabulation.
    r_tab = etree.SubElement(
        p,
        qn("r")
    )
    etree.SubElement(
        r_tab,
        qn("tab")
    )

    # Style AR depuis 8-10.
    ar_rpr = modele_ar.find(
        "w:r[2]/w:rPr",
        namespaces=NS
    )

    if ar_rpr is None:
        ar_rpr = modele_ar.find(
            "w:r/w:rPr",
            namespaces=NS
        )

    r_ar = etree.SubElement(
        p,
        qn("r")
    )

    if ar_rpr is not None:
        rp = deepcopy(ar_rpr)

        rtl = rp.find(
            "w:rtl",
            namespaces=NS
        )

        if rtl is None:
            rtl = etree.SubElement(
                rp,
                qn("rtl")
            )

        rtl.set(
            qn("val"),
            "1"
        )

        r_ar.append(
            rp
        )

    t_ar = etree.SubElement(
        r_ar,
        qn("t")
    )
    t_ar.text = (
        "المصدر: النشرة الإحصائية السنوية للمغرب 2025"
    )

    return p


if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Document introuvable : {FICHIER_ENTREE}"
    )


print("=" * 94)
print("CORRECTION DEFINITIVE DES SOURCES ANNUAIRE 2025")
print("=" * 94)
print(f"Entrée : {FICHIER_ENTREE}")
print()
print("Le fichier précédent avec les sources dupliquées n'est PAS utilisé.")
print()


with ZipFile(FICHIER_ENTREE, "r") as zin:

    root = etree.fromstring(
        zin.read(
            "word/document.xml"
        )
    )

    body = root.find(
        "w:body",
        namespaces=NS
    )

    if body is None:
        raise RuntimeError(
            "w:body introuvable."
        )

    modele_fr, modele_ar = trouver_modeles_8_10(
        body
    )

    compte_mixte = 0
    compte_fr = 0
    compte_ar = 0

    # ========================================================
    # 1) NORMALISER CHAQUE SOURCE SANS AJOUTER DE DOUBLONS
    # ========================================================

    for p in list(
        body.findall(
            "w:p",
            namespaces=NS
        )
    ):
        t = normaliser(
            texte_paragraphe(p)
        )

        fr = est_source_fr(t)
        ar = est_source_ar(t)

        if not fr and not ar:
            continue

        sect = get_sectpr(p)

        if fr and ar:
            # 1 paragraphe avant -> 1 paragraphe après.
            nouveau = construire_source_mixte(
                p,
                modele_fr,
                modele_ar
            )

            add_sectpr(
                nouveau,
                sect
            )

            body.replace(
                p,
                nouveau
            )

            compte_mixte += 1

        elif fr:
            nouveau = deepcopy(
                modele_fr
            )

            add_sectpr(
                nouveau,
                sect
            )

            body.replace(
                p,
                nouveau
            )

            compte_fr += 1

        elif ar:
            nouveau = deepcopy(
                modele_ar
            )

            add_sectpr(
                nouveau,
                sect
            )

            body.replace(
                p,
                nouveau
            )

            compte_ar += 1

    # ========================================================
    # 2) SI FR ET AR SONT SEPARES PAR DES NOTES,
    #    REMONTER AR JUSTE APRES FR
    # ========================================================

    enfants = list(body)

    indexes_tables = [
        i
        for i, el in enumerate(enfants)
        if el.tag == qn("tbl")
    ]

    deplacements = 0

    for k, debut in enumerate(
        indexes_tables
    ):
        fin = (
            indexes_tables[k + 1]
            if k + 1 < len(indexes_tables)
            else len(enfants)
        )

        region = enfants[
            debut + 1:fin
        ]

        sources_fr = [
            el
            for el in region
            if (
                el.tag == qn("p")
                and est_source_fr(
                    texte_paragraphe(el)
                )
                and not est_source_ar(
                    texte_paragraphe(el)
                )
            )
        ]

        sources_ar = [
            el
            for el in region
            if (
                el.tag == qn("p")
                and est_source_ar(
                    texte_paragraphe(el)
                )
                and not est_source_fr(
                    texte_paragraphe(el)
                )
            )
        ]

        if (
            len(sources_fr) == 1
            and len(sources_ar) == 1
        ):
            p_fr = sources_fr[0]
            p_ar = sources_ar[0]

            index_fr = body.index(
                p_fr
            )
            index_ar = body.index(
                p_ar
            )

            if index_ar != index_fr + 1:
                body.remove(
                    p_ar
                )

                index_fr = body.index(
                    p_fr
                )

                body.insert(
                    index_fr + 1,
                    p_ar
                )

                deplacements += 1

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


print("=" * 94)
print("TERMINE")
print("=" * 94)
print(f"Sources mixtes corrigées : {compte_mixte}")
print(f"Sources FR normalisées    : {compte_fr}")
print(f"Sources AR normalisées    : {compte_ar}")
print(f"Sources AR remontées à côté du FR : {deplacements}")
print()
print("AUCUNE source n'a été dupliquée.")
print("AUCUNE donnée numérique n'a été modifiée.")
print("Les sources RGPH / Ministères / HCP ne sont pas touchées.")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 94)
