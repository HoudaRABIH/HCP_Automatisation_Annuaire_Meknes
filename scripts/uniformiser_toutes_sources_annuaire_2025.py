from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import re


DOSSIER_SORTIE = Path("sorties")

FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_UNIFORMES.docx"
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag):
    return f"{{{W}}}{tag}"


def texte_paragraphe(p):
    return "".join(
        p.xpath(".//w:t/text()", namespaces=NS)
    ).strip()


def normaliser_espaces(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def est_source_fr_annuaire_2025(t):
    x = normaliser_espaces(t).lower()

    return (
        "2025" in x
        and (
            "annuaire statistique du maroc" in x
            or "maroc du statistique annuaire" in x
        )
    )


def est_source_ar_annuaire_2025(t):
    x = normaliser_espaces(t)

    return (
        "2025" in x
        and (
            "النشرة الإحصائية السنوية للمغرب" in x
            or (
                "المصدر" in x
                and "الإحصائية" in x
                and "المغرب" in x
            )
        )
    )


def contient_fr_et_ar(t):
    return (
        est_source_fr_annuaire_2025(t)
        and est_source_ar_annuaire_2025(t)
    )


def retirer_sectpr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        return None

    sect = pPr.find("w:sectPr", namespaces=NS)

    if sect is None:
        return None

    copie = deepcopy(sect)
    pPr.remove(sect)

    return copie


def recuperer_sectpr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        return None

    sect = pPr.find("w:sectPr", namespaces=NS)

    return deepcopy(sect) if sect is not None else None


def ajouter_sectpr(p, sect):
    if sect is None:
        return

    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        pPr = etree.Element(qn("pPr"))
        p.insert(0, pPr)

    ancien = pPr.find("w:sectPr", namespaces=NS)

    if ancien is not None:
        pPr.remove(ancien)

    pPr.append(deepcopy(sect))


def trouver_modele_8_10(root):
    """
    Cherche les deux paragraphes de source Annuaire 2025 situés
    après le tableau 8-10. On utilise leur format exact comme modèle.
    """

    body = root.find("w:body", namespaces=NS)

    if body is None:
        raise RuntimeError("w:body introuvable.")

    enfants = list(body)

    index_titre_8_10 = None

    for i, el in enumerate(enfants):
        texte = " ".join(
            el.xpath(".//w:t/text()", namespaces=NS)
        )

        if (
            "8 - 10" in texte
            and "juridictions de proximité" in texte.lower()
        ):
            index_titre_8_10 = i
            break

    if index_titre_8_10 is None:
        raise RuntimeError(
            "Impossible de localiser le tableau 8-10 servant de modèle."
        )

    modele_fr = None
    modele_ar = None

    # Chercher après 8-10 jusqu'à l'apparition de 8-11.
    for el in enfants[index_titre_8_10 + 1:]:
        texte_global = normaliser_espaces(
            " ".join(
                el.xpath(".//w:t/text()", namespaces=NS)
            )
        )

        if "8 - 11" in texte_global:
            break

        for p in el.xpath(".//w:p", namespaces=NS):
            t = texte_paragraphe(p)

            if modele_fr is None and est_source_fr_annuaire_2025(t):
                modele_fr = deepcopy(p)

            if modele_ar is None and est_source_ar_annuaire_2025(t):
                modele_ar = deepcopy(p)

        # Si l'élément lui-même est un paragraphe.
        if el.tag == qn("p"):
            t = texte_paragraphe(el)

            if modele_fr is None and est_source_fr_annuaire_2025(t):
                modele_fr = deepcopy(el)

            if modele_ar is None and est_source_ar_annuaire_2025(t):
                modele_ar = deepcopy(el)

        if modele_fr is not None and modele_ar is not None:
            break

    if modele_fr is None or modele_ar is None:
        raise RuntimeError(
            "Impossible de récupérer les deux sources FR/AR propres de 8-10."
        )

    # Le modèle ne doit transporter aucun saut de section.
    retirer_sectpr(modele_fr)
    retirer_sectpr(modele_ar)

    return modele_fr, modele_ar


if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Document introuvable : {FICHIER_ENTREE}"
    )


print("=" * 96)
print("UNIFORMISATION GLOBALE DES SOURCES 'ANNUAIRE STATISTIQUE DU MAROC 2025'")
print("=" * 96)
print(f"Document d'entrée : {FICHIER_ENTREE}")
print()
print("Modèle visuel utilisé : source du tableau 8-10")
print("Cible : TOUS les chapitres du document")
print()


with ZipFile(FICHIER_ENTREE, "r") as zin:
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)

    modele_fr, modele_ar = trouver_modele_8_10(root)

    # Snapshot avant modifications.
    paragraphes = list(
        root.xpath(".//w:p", namespaces=NS)
    )

    nb_fr = 0
    nb_ar = 0
    nb_mixtes = 0

    for p in paragraphes:
        parent = p.getparent()

        # Peut déjà avoir été supprimé lors du traitement d'un paragraphe mixte.
        if parent is None:
            continue

        t = texte_paragraphe(p)

        cible_fr = est_source_fr_annuaire_2025(t)
        cible_ar = est_source_ar_annuaire_2025(t)

        if not cible_fr and not cible_ar:
            continue

        # Conserver un éventuel saut de section de ce paragraphe.
        sect = recuperer_sectpr(p)

        position = parent.index(p)

        if cible_fr and cible_ar:
            # Cas où FR + AR se sont retrouvés fusionnés dans un même paragraphe.
            nouveau_fr = deepcopy(modele_fr)
            nouveau_ar = deepcopy(modele_ar)

            # Le saut de section reste sur le dernier paragraphe.
            ajouter_sectpr(nouveau_ar, sect)

            parent.remove(p)
            parent.insert(position, nouveau_fr)
            parent.insert(position + 1, nouveau_ar)

            nb_fr += 1
            nb_ar += 1
            nb_mixtes += 1

        elif cible_fr:
            nouveau = deepcopy(modele_fr)
            ajouter_sectpr(nouveau, sect)

            parent.remove(p)
            parent.insert(position, nouveau)

            nb_fr += 1

        elif cible_ar:
            nouveau = deepcopy(modele_ar)
            ajouter_sectpr(nouveau, sect)

            parent.remove(p)
            parent.insert(position, nouveau)

            nb_ar += 1

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
            data = zin.read(item.filename)

            if item.filename == "word/document.xml":
                data = nouveau_xml

            zout.writestr(item, data)


print("=" * 96)
print("SOURCES UNIFORMISEES")
print("=" * 96)
print(f"Sources françaises corrigées : {nb_fr}")
print(f"Sources arabes corrigées      : {nb_ar}")
print(f"Paragraphes FR+AR fusionnés réparés : {nb_mixtes}")
print()
print("Toutes les sources ciblées utilisent maintenant EXACTEMENT")
print("le même format que la source correcte du tableau 8-10.")
print()
print("NON MODIFIE :")
print("  - données numériques")
print("  - titres")
print("  - notes")
print("  - sources RGPH / HCP / Ministères / autres organismes")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 96)
