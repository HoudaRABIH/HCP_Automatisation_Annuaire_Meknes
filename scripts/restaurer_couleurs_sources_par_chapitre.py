from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import re


DOSSIER_SORTIE = Path("sorties")

# Référence pour récupérer la COULEUR originale de chaque chapitre
FICHIER_REFERENCE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx"
)

# Fichier où la POSITION des sources est déjà corrigée
FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_FIXED.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_COLORS_OK.docx"
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


def est_source_annuaire_2025(p):
    t = normaliser(
        texte_paragraphe(p)
    )
    low = t.lower()
    compact = re.sub(r"\s+", "", low)

    return (
        "2025" in t
        and (
            "annuairestatistiquedumaroc" in compact
            or "marocdustatistiqueannuaire" in compact
            or "النشرة الإحصائية السنوية للمغرب" in t
            or (
                "المصدر" in t
                and "الإحصائية" in t
                and "المغرب" in t
            )
        )
    )


def couleur_du_paragraphe(p):
    """
    Récupère la couleur explicite utilisée dans le paragraphe source
    du fichier propre d'origine.
    """

    # 1) Couleur des runs visibles
    couleurs = []

    for r in p.xpath(".//w:r", namespaces=NS):
        txt = "".join(
            r.xpath(".//w:t/text()", namespaces=NS)
        )

        if not txt.strip():
            continue

        color = r.find(
            "w:rPr/w:color",
            namespaces=NS
        )

        if color is not None:
            val = color.get(qn("val"))

            if val and val.lower() not in ("auto", "000000"):
                couleurs.append(val)

    if couleurs:
        # La première couleur visible est normalement celle du chapitre.
        return couleurs[0]

    # 2) Couleur au niveau du paragraphe
    color = p.find(
        "w:pPr/w:rPr/w:color",
        namespaces=NS
    )

    if color is not None:
        val = color.get(qn("val"))

        if val and val.lower() != "auto":
            return val

    # 3) Si noir est explicitement utilisé, le conserver aussi.
    for r in p.xpath(".//w:r", namespaces=NS):
        color = r.find(
            "w:rPr/w:color",
            namespaces=NS
        )

        if color is not None:
            val = color.get(qn("val"))

            if val and val.lower() != "auto":
                return val

    return None


def appliquer_couleur(p, couleur):
    if not couleur:
        return 0

    nb = 0

    for r in p.xpath(".//w:r", namespaces=NS):
        txt = "".join(
            r.xpath(".//w:t/text()", namespaces=NS)
        )

        # Ne pas colorer les runs purement techniques (tabulations seules).
        if not txt.strip():
            continue

        rPr = r.find(
            "w:rPr",
            namespaces=NS
        )

        if rPr is None:
            rPr = etree.Element(
                qn("rPr")
            )
            r.insert(
                0,
                rPr
            )

        color = rPr.find(
            "w:color",
            namespaces=NS
        )

        if color is None:
            color = etree.SubElement(
                rPr,
                qn("color")
            )

        color.set(
            qn("val"),
            couleur
        )

        # Supprimer d'éventuels attributs thème qui peuvent écraser la couleur.
        for attr in (
            qn("themeColor"),
            qn("themeTint"),
            qn("themeShade"),
        ):
            if attr in color.attrib:
                del color.attrib[attr]

        nb += 1

    return nb


if not FICHIER_REFERENCE.exists():
    raise FileNotFoundError(
        f"Fichier de référence introuvable : {FICHIER_REFERENCE}"
    )

if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Fichier SOURCES_FIXED introuvable : {FICHIER_ENTREE}"
    )


print("=" * 96)
print("RESTAURATION DES COULEURS DE SOURCE PAR CHAPITRE")
print("=" * 96)
print(f"Référence couleurs : {FICHIER_REFERENCE}")
print(f"Entrée mise en page : {FICHIER_ENTREE}")
print()


with ZipFile(
    FICHIER_REFERENCE,
    "r"
) as zref:

    root_ref = etree.fromstring(
        zref.read(
            "word/document.xml"
        )
    )

    sources_ref = [
        p
        for p in root_ref.xpath(
            ".//w:p",
            namespaces=NS
        )
        if est_source_annuaire_2025(p)
    ]


with ZipFile(
    FICHIER_ENTREE,
    "r"
) as zin:

    root = etree.fromstring(
        zin.read(
            "word/document.xml"
        )
    )

    sources = [
        p
        for p in root.xpath(
            ".//w:p",
            namespaces=NS
        )
        if est_source_annuaire_2025(p)
    ]

    print(
        f"Sources trouvées dans la référence : {len(sources_ref)}"
    )
    print(
        f"Sources trouvées dans le fichier corrigé : {len(sources)}"
    )

    if len(sources_ref) != len(sources):
        raise RuntimeError(
            "Le nombre de sources n'est pas identique entre les deux fichiers. "
            "Aucune modification n'a été faite pour éviter un mauvais appariement."
        )

    nb_paragraphes = 0
    nb_runs = 0
    couleurs_vues = []

    for pref, p in zip(
        sources_ref,
        sources
    ):
        couleur = couleur_du_paragraphe(
            pref
        )

        if couleur:
            couleurs_vues.append(
                couleur
            )

            nb_runs += appliquer_couleur(
                p,
                couleur
            )

            nb_paragraphes += 1

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


print()
print("=" * 96)
print("COULEURS RESTAUREES")
print("=" * 96)
print(f"Paragraphes source recolorés : {nb_paragraphes}")
print(f"Runs recolorés               : {nb_runs}")
print(
    "Couleurs distinctes restaurées : "
    + ", ".join(
        sorted(set(couleurs_vues))
    )
)
print()
print("La POSITION des sources reste celle du fichier SOURCES_FIXED.")
print("La COULEUR revient à celle propre à chaque chapitre.")
print()
print("Aucune donnée numérique, année, note ou autre source n'a été modifiée.")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 96)
