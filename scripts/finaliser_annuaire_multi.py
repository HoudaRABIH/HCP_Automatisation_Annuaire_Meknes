from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from lxml import etree
import os
import re
import unicodedata

ROOT = Path('.')
SORTIES = ROOT / 'sorties'
YEAR = os.environ.get('HCP_EDITION_YEAR', '').strip()
if not re.fullmatch(r'20\d{2}', YEAR):
    raise RuntimeError(f'Edition HCP invalide : {YEAR!r}')

DONOR = SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx'
TARGET = SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx'
OUTPUT = SORTIES / f'Annuaire_HCP_{YEAR}_FINAL.docx'

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def norm(s):
    s = str(s or '').replace('\u00a0', ' ').replace('\u202f', ' ')
    s = s.replace('’', "'").replace('–', '-').replace('—', '-')
    s = ''.join(
        c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn'
    )
    return re.sub(r'\s+', ' ', s).strip().lower()


def child_text(el):
    return ''.join(el.xpath('.//w:t/text()', namespaces=NS))


def find_44_bounds(root):
    body = root.find(f'.//{W}body')
    if body is None:
        raise RuntimeError('w:body introuvable.')

    children = list(body)
    start = None
    end = None

    for i, el in enumerate(children):
        t = norm(child_text(el))
        if start is None:
            if (
                re.search(r'\b4\s*-\s*4\b', t)
                and 'effectif des etudiants residant' in t
                and 'cite' in t
            ):
                start = i
                continue
        if start is not None:
            if (
                re.search(r'\b4\s*-\s*5\b', t)
                and 'effectif des laureats' in t
            ):
                end = i
                break

    if start is None or end is None or end <= start:
        raise RuntimeError(
            f'Bloc 4-4 impossible à délimiter : start={start}, end={end}'
        )
    return body, start, end


def collect_rids(elements):
    ids = set()
    for el in elements:
        for node in el.iter():
            for attr, val in node.attrib.items():
                if attr.startswith(R) and attr.endswith('id'):
                    ids.add(val)
    return ids


def rel_ids_from_zip(zf):
    name = 'word/_rels/document.xml.rels'
    root = etree.fromstring(zf.read(name))
    return {el.get('Id') for el in root if el.get('Id')}


def rewrite_nodes(nodes, new_text):
    lengths = [len(n.text or '') for n in nodes]
    pos = 0
    for i, node in enumerate(nodes):
        if i < len(nodes) - 1:
            take = lengths[i]
            node.text = new_text[pos:pos + take]
            pos += take
        else:
            node.text = new_text[pos:]


def patch_visible_text(root):
    source_changes = 0
    typo_changes = 0
    year_re = re.compile(r"20\d{2}")
    ar_marker = norm("النشرة الإحصائية السنوية للمغرب")

    for p in root.xpath(".//w:p", namespaces=NS):
        # Direct text nodes of this paragraph only (excluding nested text-box
        # paragraphs). This is required for 4-4 where the source line lives
        # directly in a container paragraph after nested shape paragraphs.
        all_nodes = p.xpath(".//w:t", namespaces=NS)
        direct_nodes = []
        for node in all_nodes:
            ancestors = node.xpath("ancestor::w:p[1]", namespaces=NS)
            if ancestors and ancestors[0] is p:
                direct_nodes.append(node)

        if direct_nodes:
            direct_text = "".join(n.text or "" for n in direct_nodes)
            direct_new = direct_text.replace("cccccccc", "")
            if direct_new != direct_text:
                typo_changes += direct_text.count("cccccccc")

            dnt = norm(direct_new)
            if (
                "annuaire statistique du maroc" in dnt
                or ar_marker in dnt
            ):
                newer = year_re.sub(YEAR, direct_new)
                if newer != direct_new:
                    source_changes += 1
                    direct_new = newer

            if direct_new != direct_text:
                rewrite_nodes(direct_nodes, direct_new)

        # Leaf paragraphs are safe to rewrite as a whole.
        if p.xpath(".//w:p", namespaces=NS):
            continue
        nodes = p.xpath(".//w:t", namespaces=NS)
        if not nodes:
            continue
        text = "".join(n.text or "" for n in nodes)
        new = text.replace("cccccccc", "")
        if new != text:
            typo_changes += text.count("cccccccc")

        nt = norm(new)
        if (
            "annuaire statistique du maroc" in nt
            or ar_marker in nt
        ):
            newer = year_re.sub(YEAR, new)
            if newer != new:
                source_changes += 1
                new = newer

        if new != text:
            rewrite_nodes(nodes, new)

    return source_changes, typo_changes

def donor_sanity(block):
    text = norm(' '.join(child_text(el) for el in block))
    # 4-4 is a floating-text-box block; its data may not include geography
    # labels. We require multiple numeric values and the expected heading.
    digits = re.findall(r'\d+', text)
    if len(digits) < 6:
        raise RuntimeError('Bloc donneur 4-4 incomplet : moins de 6 valeurs numériques.')


def main():
    if not DONOR.exists():
        raise FileNotFoundError(f'Donneur Education introuvable : {DONOR}')
    if not TARGET.exists():
        raise FileNotFoundError(f'Word intermédiaire introuvable : {TARGET}')

    print('=' * 104)
    print(f'FINALISATION MULTI-ANNEE — EDITION {YEAR}')
    print('=' * 104)

    with ZipFile(DONOR, 'r') as zd:
        donor_doc = etree.fromstring(zd.read('word/document.xml'))
        donor_body, ds, de = find_44_bounds(donor_doc)
        donor_block = [deepcopy(el) for el in list(donor_body)[ds:de]]
        if not donor_block:
            raise RuntimeError('Bloc donneur 4-4 vide.')
        donor_sanity(donor_block)
        donor_rids = collect_rids(donor_block)

    with ZipFile(TARGET, 'r') as zt:
        target_doc = etree.fromstring(zt.read('word/document.xml'))
        target_body, ts, te = find_44_bounds(target_doc)

        missing_rids = sorted(donor_rids - rel_ids_from_zip(zt))
        if missing_rids:
            raise RuntimeError(
                'Relations Word manquantes pour restaurer 4-4 : '
                + ', '.join(missing_rids)
            )

        children = list(target_body)
        for el in children[ts:te]:
            target_body.remove(el)
        for offset, el in enumerate(donor_block):
            target_body.insert(ts + offset, deepcopy(el))

        source_total = 0
        typo_total = 0
        s, t = patch_visible_text(target_doc)
        source_total += s
        typo_total += t

        new_document_xml = etree.tostring(
            target_doc,
            xml_declaration=True,
            encoding='UTF-8',
            standalone='yes',
        )

        with ZipFile(OUTPUT, 'w', ZIP_DEFLATED) as zout:
            for item in zt.infolist():
                data = zt.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = new_document_xml
                elif item.filename.startswith('word/') and item.filename.endswith('.xml'):
                    try:
                        root = etree.fromstring(data)
                        s, t = patch_visible_text(root)
                        source_total += s
                        typo_total += t
                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding='UTF-8',
                            standalone='yes',
                        )
                    except Exception:
                        pass
                zout.writestr(item, data)

    # Final structural checks.
    with ZipFile(OUTPUT, 'r') as zf:
        final_doc = etree.fromstring(zf.read('word/document.xml'))
        _, fs, fe = find_44_bounds(final_doc)
        body = final_doc.find(f'.//{W}body')
        block = list(body)[fs:fe]
        donor_sanity(block)

        old_source_years = 0
        typo_left = 0
        for name in zf.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'):
                continue
            try:
                root = etree.fromstring(zf.read(name))
            except Exception:
                continue
            for p in root.xpath('.//w:p', namespaces=NS):
                text = ''.join(p.xpath('.//w:t/text()', namespaces=NS))
                nt = norm(text)
                if 'cccccccc' in text:
                    typo_left += 1
                if (
                    'annuaire statistique du maroc' in nt
                    or norm('النشرة الإحصائية السنوية للمغرب') in nt
                ):
                    years = set(re.findall(r'20\d{2}', text))
                    if years and years != {YEAR}:
                        old_source_years += 1

        if typo_left:
            raise RuntimeError(f'Artefact cccccccc restant : {typo_left}')
        if old_source_years:
            raise RuntimeError(
                f'Sources Annuaire avec une année différente de {YEAR} : {old_source_years}'
            )

    print('Bloc 4-4 restauré depuis le donneur Education mis à jour.')
    print(f'Sources Annuaire synchronisées : {source_total}')
    print(f'Artefacts supprimés : {typo_total}')
    print(f'Fichier final : {OUTPUT}')
    print('=' * 104)


if __name__ == '__main__':
    main()
