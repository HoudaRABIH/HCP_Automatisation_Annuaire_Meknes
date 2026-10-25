from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree


DOSSIER_SORTIE = Path("sorties")

FICHIER_ENTREE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx"
)

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_SOURCES_OK.docx"
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(tag):
    return f"{{{W}}}{tag}"


def texte(p):
    return "".join(
        p.xpath(".//w:t/text()", namespaces=NS)
    ).strip()


def est_source(p):
    t = texte(p).lower()
    return (
        "annuaire statistique du maroc" in t
        or "maroc du statistique annuaire" in t
        or "النشرة الإحصائية السنوية للمغرب" in texte(p)
    )


def trouver_top_level_tables(body):
    return [
        el for el in body
        if el.tag == qn("tbl")
    ]


def child_index(body, element):
    for i, child in enumerate(body):
        if child is element:
            return i
    raise RuntimeError("Élément introuvable dans le body.")


def source_paragraphs_between(body, start_idx, end_idx):
    result = []

    for i in range(start_idx + 1, end_idx):
        el = body[i]

        if el.tag != qn("p"):
            continue

        if est_source(el):
            result.append(el)

    return result


def get_sectpr(source_paras):
    for p in reversed(source_paras):
        pPr = p.find("w:pPr", namespaces=NS)

        if pPr is None:
            continue

        sect = pPr.find("w:sectPr", namespaces=NS)

        if sect is not None:
            return deepcopy(sect)

    return None


def remove_sectpr(p):
    pPr = p.find("w:pPr", namespaces=NS)

    if pPr is None:
        return

    sect = pPr.find("w:sectPr", namespaces=NS)

    if sect is not None:
        pPr.remove(sect)


def attach_sectpr(p, sect):
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


if not FICHIER_ENTREE.exists():
    raise FileNotFoundError(
        f"Document introuvable : {FICHIER_ENTREE}"
    )


print("=" * 92)
print("UNIFORMISATION DES SOURCES JUSTICE SUR LE MODELE 8-10")
print("=" * 92)
print(f"Document d'entrée : {FICHIER_ENTREE}")
print()


with ZipFile(FICHIER_ENTREE, "r") as zin:
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)

    body = root.find("w:body", namespaces=NS)

    if body is None:
        raise RuntimeError("w:body introuvable.")

    tables = trouver_top_level_tables(body)

    if len(tables) < 267:
        raise RuntimeError(
            f"Seulement {len(tables)} tableaux top-level trouvés."
        )

    # --------------------------------------------------------
    # MODELE VISUEL = SOURCE DE 8-10
    #
    # 8-10 data table = objet Word 265
    # 8-11 title      = objet Word 266
    # Entre les deux, le document possède déjà la source correcte :
    #   FR à gauche
    #   AR à droite
    # --------------------------------------------------------

    table_8_10 = tables[265 - 1]
    table_8_11_title = tables[266 - 1]

    idx_8_10 = child_index(body, table_8_10)
    idx_8_11_title = child_index(body, table_8_11_title)

    template_paras = source_paragraphs_between(
        body,
        idx_8_10,
        idx_8_11_title
    )

    template_fr = None
    template_ar = None

    for p in template_paras:
        t = texte(p)

        if (
            "Annuaire Statistique du Maroc" in t
            and "النشرة الإحصائية" not in t
        ):
            template_fr = deepcopy(p)

        if (
            "النشرة الإحصائية السنوية للمغرب" in t
            and "Annuaire Statistique du Maroc" not in t
            and "Maroc du Statistique Annuaire" not in t
        ):
            template_ar = deepcopy(p)

    if template_fr is None or template_ar is None:
        raise RuntimeError(
            "Impossible de récupérer les deux paragraphes source propres de 8-10."
        )

    # Les templates ne doivent jamais transporter un saut de section.
    remove_sectpr(template_fr)
    remove_sectpr(template_ar)

    # --------------------------------------------------------
    # TABLEAUX DE DONNEES JUSTICE
    # --------------------------------------------------------
    # 8-1  241
    # 8-2  244
    # 8-3  246
    # 8-4  249
    # 8-5  251
    # 8-6  254
    # 8-7  257
    # 8-8  259
    # 8-9  262
    # 8-10 265 -> déjà le modèle correct
    # 8-11 267
    #
    justice_data_objects = [
        241, 244, 246, 249, 251,
        254, 257, 259, 262, 267
    ]

    corrections = 0

    for no_objet in justice_data_objects:
        # Recalculer la liste à chaque passe car le body est modifié.
        tables_now = trouver_top_level_tables(body)

        data_tbl = tables_now[no_objet - 1]

        # Le prochain tableau top-level marque la fin de la zone source.
        next_tbl = tables_now[no_objet]

        start_idx = child_index(body, data_tbl)
        end_idx = child_index(body, next_tbl)

        old_sources = source_paragraphs_between(
            body,
            start_idx,
            end_idx
        )

        if not old_sources:
            print(
                f"Objet {no_objet}: aucune source Annuaire trouvée — ignoré."
            )
            continue

        # Conserver un éventuel saut de section de l'ancien paragraphe.
        sectpr = get_sectpr(old_sources)

        # Emplacement d'insertion = premier ancien paragraphe source.
        first_source_idx = child_index(
            body,
            old_sources[0]
        )

        # Supprimer uniquement les anciens paragraphes de SOURCE.
        for p in old_sources:
            parent = p.getparent()
            parent.remove(p)

        # Cloner exactement le visuel 8-10.
        new_fr = deepcopy(template_fr)
        new_ar = deepcopy(template_ar)

        # Le saut de section doit rester au même endroit logique,
        # sur le dernier paragraphe de source.
        attach_sectpr(
            new_ar,
            sectpr
        )

        body.insert(
            first_source_idx,
            new_fr
        )
        body.insert(
            first_source_idx + 1,
            new_ar
        )

        corrections += 1

        print(
            f"Objet {no_objet}: source uniformisée comme 8-10."
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
            data = zin.read(item.filename)

            if item.filename == "word/document.xml":
                data = nouveau_xml

            zout.writestr(
                item,
                data
            )


print()
print("=" * 92)
print("SOURCES JUSTICE CORRIGEES")
print("=" * 92)
print(f"Blocs de source corrigés : {corrections}")
print()
print("Tous les blocs corrigés utilisent maintenant EXACTEMENT")
print("la même structure visuelle que la source du tableau 8-10 :")
print("  - Français à gauche")
print("  - Arabe à droite")
print("  - même police / taille / espacement")
print()
print("Aucune donnée numérique n'a été modifiée.")
print("Le tableau 8-10 n'a pas été touché : il sert de modèle.")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 92)
