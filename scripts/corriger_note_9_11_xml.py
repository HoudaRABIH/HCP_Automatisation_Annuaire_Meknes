from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import tempfile
import shutil


DOSSIER_SORTIE = Path("sorties")

CANDIDATS = [
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_FINAL.docx",
    DOSSIER_SORTIE / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_OK.docx",
]

FICHIER_SORTIE = (
    DOSSIER_SORTIE
    / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_VALIDE.docx"
)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def texte_paragraphe(paragraphe):
    return "".join(
        paragraphe.xpath(".//w:t/text()", namespaces=NS)
    )


def remplacer_dans_paragraphe(paragraphe, ancien, nouveau):
    """
    Remplace ancien -> nouveau directement dans les noeuds w:t.
    Cela conserve le formatage Word, car on ne recrée ni paragraphe ni run.
    """
    noeuds = paragraphe.xpath(".//w:t", namespaces=NS)

    texte_total = "".join(
        n.text or ""
        for n in noeuds
    )

    if ancien not in texte_total:
        return 0

    # Cas simple : l'expression est dans un seul noeud.
    for n in noeuds:
        if n.text and ancien in n.text:
            n.text = n.text.replace(ancien, nouveau)
            return 1

    # Cas Word découpé sur plusieurs runs/noeuds.
    # On reconstruit uniquement le texte visible de ce paragraphe,
    # en conservant le premier noeud et en vidant les autres.
    nouveau_total = texte_total.replace(ancien, nouveau)

    if noeuds:
        noeuds[0].text = nouveau_total

        for n in noeuds[1:]:
            n.text = ""

        return 1

    return 0


fichier_entree = None

for candidat in CANDIDATS:
    if candidat.exists():
        fichier_entree = candidat
        break

if fichier_entree is None:
    raise FileNotFoundError(
        "Aucun fichier Agriculture FINAL/OK trouvé."
    )


print("=" * 82)
print("CORRECTION OOXML DE LA NOTE 9-11")
print("=" * 82)
print(f"Document : {fichier_entree}")
print()


with ZipFile(fichier_entree, "r") as zin:
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)

    paragraphes = root.xpath(".//w:p", namespaces=NS)

    indices_note = []

    for i, p in enumerate(paragraphes):
        texte = texte_paragraphe(p)

        if (
            "Inventaire National Forestier" in texte
            or "الجرد الوطني" in texte
        ):
            indices_note.append(i)

    print(
        "Paragraphe(s) de note trouvé(s) : "
        + (
            ", ".join(str(i) for i in indices_note)
            if indices_note
            else "AUCUN"
        )
    )

    if not indices_note:
        raise RuntimeError(
            "La note 9-11 n'a pas été trouvée dans le XML Word."
        )

    nb = 0

    # Corriger la phrase de note elle-même.
    for i in indices_note:
        nb += remplacer_dans_paragraphe(
            paragraphes[i],
            "(3)",
            "(1)"
        )

        # Corriger également un numéro isolé "(3)" placé juste avant/après
        # la phrase, ce qui correspond à la mise en page visible dans Word.
        for j in range(
            max(0, i - 6),
            min(len(paragraphes), i + 7)
        ):
            texte_voisin = texte_paragraphe(
                paragraphes[j]
            ).strip()

            if texte_voisin == "(3)":
                nb += remplacer_dans_paragraphe(
                    paragraphes[j],
                    "(3)",
                    "(1)"
                )

    nouveau_xml = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone="yes"
    )

    # Réécrire le DOCX sans toucher aux autres fichiers internes.
    with ZipFile(FICHIER_SORTIE, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "word/document.xml":
                data = nouveau_xml

            zout.writestr(item, data)


print()
print(f"Remplacement(s) effectué(s) : {nb}")
print()
print("Résultat attendu :")
print("  Année 2024(1)")
print("  (1) Statistiques de l'Inventaire National Forestier.")
print("  équivalent arabe également en (1)")
print()
print(f"Document créé : {FICHIER_SORTIE}")
print("=" * 82)
