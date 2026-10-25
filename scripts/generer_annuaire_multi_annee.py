from pathlib import Path
import argparse
import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from datetime import datetime

ROOT = Path('.')
SCRIPTS = ROOT / 'scripts'
SORTIES = ROOT / 'sorties'
CANONICAL_DATA = ROOT / 'donnees_brutes'
MODELES = ROOT / 'modeles'
REFERENCES = ROOT / 'references'

MODEL_CANONICAL = MODELES / 'modeles.docx'
REFERENCE_EDUCATION = REFERENCES / 'REFERENCE_EDUCATION_2025.docx'  # layout donor only
EXISTING_EDUCATION = SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx'
LOG = SORTIES / 'rapport_generation_tous_secteurs.txt'

REQUIRED_SCRIPTS = [
    'generer_transport.py',
    'generer_prix.py',
    'generer_environnement.py',
    'generer_tourisme.py',
    'generer_agriculture.py',
    'generer_sante_corrige.py',
    'appliquer_education_reference.py',
    'mettre_a_jour_education_multi.py',
    'generer_justice.py',
    'generer_culture.py',
    'generer_habitat_foncier.py',
    'generer_population_administration_verifie_v2.py',
    'generer_industrie_verifie.py',
    'generer_energie_eau_verifie_v2.py',
    'finaliser_annuaire_multi.py',
    'synchroniser_synthese_encadres_multi.py',
]

# Legacy file names expected by the already validated 2025 sector generators.
# The selected workbook for year X is copied temporarily under these names;
# only the FILE NAME is adapted, never its contents.
LEGACY_EXCEL_NAMES = {
    1:  '1.Division Administrative_AS 2025.xlsx',
    2:  '2. population_AS 2025.xlsx',
    3:  '3. Agriculture_AS 20252.xlsx',  # historical typo required by generer_agriculture.py
    4:  '4. MINES_AS 2025.xlsx',
    5:  '5.Energie et Eau_AS 2025.xlsx',
    6:  '6. Industrie et artisanat_AS 2025.xlsx',
    7:  '7. Foncier_AS 2025.xlsx',
    8:  '8. Transport_AS 2025.xlsx',
    9:  '9. Post-télécomm_AS 2025.xlsx',
    10: '10. Tourisme_AS 2025.xlsx',
    11: '11. Enseignement_AS 2025.xlsx',
    12: '12. Santé_AS 2025.xlsx',
    13: '13. Prévoyance sociale_AS 2025.xlsx',
    14: '14-Emploi et salaire_AS 2025.xlsx',
    15: "15.Justice et Sauvegarde de l'Enfance _AS2025.xlsx",
    16: '16. Activités culturelles_Loisirs et Entraide_AS 2025.xlsx',
    17: '17.Indice des prix_AS  2025.xlsx',
    18: '18. Commerce extérieur_AS 2025.xlsx',
    19: '19. Monnaie et crédits_AS 2025.xlsx',
    20: '20. Bourse_AS_ 2025.xlsx',
    21: '21. Finances publiques_AS 2025.xlsx',
    22: '22. compte_AS 2025.xlsx',
    23: '23. Environnement_AS 2025.xlsx',
}

OUTPUTS = {
    'transport': SORTIES / 'Annuaire_HCP_Transport.docx',
    'prix': SORTIES / 'Annuaire_HCP_Transport_Prix.docx',
    'env': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement.docx',
    'env_alias': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement_Sources2025_OK.docx',
    'tourisme': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement_Tourisme.docx',
    'agri': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture.docx',
    'agri_alias': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_VALIDE.docx',
    'sante': SORTIES / 'Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante.docx',
    'education': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx',
    'justice': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx',
    'culture': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx',
    'habitat': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER.docx',
    'population': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION.docx',
    'industrie': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE.docx',
    'energie': SORTIES / 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx',
}


def final_path(year):
    return SORTIES / f'Annuaire_HCP_{year}_FINAL.docx'


def banner(text):
    print('\n' + '=' * 108)
    print(text)
    print('=' * 108)


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('--word', required=True, help='Word initial/template')
    p.add_argument('--excel-dir', required=True, help='Dossier contenant les 23 Excel')
    return p.parse_args()


def write_log(lines):
    SORTIES.mkdir(parents=True, exist_ok=True)
    LOG.write_text('\n'.join(lines), encoding='utf-8')


def fail(msg, log):
    log.append('ECHEC : ' + msg)
    write_log(log)
    raise RuntimeError(msg)


def source_files(folder):
    return sorted(
        [p for p in folder.glob('*.xlsx') if not p.name.startswith('~$')],
        key=lambda p: _source_sort_key(p.name),
    )


def prefix_number(name):
    m = re.match(r'\s*(\d+)', str(name))
    return int(m.group(1)) if m else None


def edition_year_in_name(name):
    years = set(re.findall(r'20\d{2}', str(name)))
    if not years:
        return None
    if len(years) != 1:
        raise RuntimeError(f'Plusieurs années détectées dans le nom : {name}')
    return next(iter(years))


def validate_and_detect_sources(selected_dir, log):
    if not selected_dir.exists() or not selected_dir.is_dir():
        fail(f'Dossier Excel introuvable : {selected_dir}', log)

    files = source_files(selected_dir)
    if len(files) != 23:
        fail(f'23 fichiers Excel attendus, {len(files)} détectés.', log)

    by_prefix = {}
    file_years = {}
    for p in files:
        n = prefix_number(p.name)
        if n is None or n not in LEGACY_EXCEL_NAMES:
            fail(f'Numéro de source 1..23 introuvable dans : {p.name}', log)
        if n in by_prefix:
            fail(f'Deux fichiers portent le numéro {n}.', log)
        by_prefix[n] = p
        try:
            y = edition_year_in_name(p.name)
        except RuntimeError as exc:
            fail(str(exc), log)
        if y is None:
            fail(
                f"Année d'édition non détectable dans le nom : {p.name}. "
                "Conservez l'année dans le nom des 23 fichiers.",
                log,
            )
        file_years[p.name] = y

    missing = sorted(set(range(1, 24)) - set(by_prefix))
    if missing:
        fail('Sources numérotées manquantes : ' + ', '.join(map(str, missing)), log)

    years = sorted(set(file_years.values()))
    if len(years) != 1:
        details = ', '.join(f'{name}={year}' for name, year in file_years.items())
        fail('Les 23 fichiers ne correspondent pas à la même édition : ' + details, log)

    return files, by_prefix, years[0]


def ensure_education_reference(log):
    REFERENCES.mkdir(parents=True, exist_ok=True)
    if REFERENCE_EDUCATION.exists():
        print(f'Référence de mise en page Enseignement : {REFERENCE_EDUCATION}')
        return

    if EXISTING_EDUCATION.exists():
        shutil.copy2(EXISTING_EDUCATION, REFERENCE_EDUCATION)
        log.append(f'Référence Enseignement créée : {REFERENCE_EDUCATION}')
        return

    fail(
        'Référence de mise en page Enseignement absente : '
        'references\\REFERENCE_EDUCATION_2025.docx',
        log,
    )


def stage_model(selected, log):
    if not selected.exists():
        fail(f'Word initial introuvable : {selected}', log)
    if selected.suffix.lower() != '.docx':
        fail('Le Word initial doit être un fichier .docx.', log)

    MODELES.mkdir(parents=True, exist_ok=True)
    if selected.resolve() == MODEL_CANONICAL.resolve():
        log.append(f'Word utilisateur utilisé directement : {selected}')
        return {'staged': False, 'backup': None, 'created': False}

    backup = None
    created = not MODEL_CANONICAL.exists()
    if MODEL_CANONICAL.exists():
        bdir = SORTIES / 'backups'
        bdir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup = bdir / f'modele_original_{stamp}.docx'
        shutil.copy2(MODEL_CANONICAL, backup)

    shutil.copy2(selected, MODEL_CANONICAL)
    log.append(f'Word utilisateur : {selected}')
    if backup:
        log.append(f'Sauvegarde modèle précédent : {backup}')
    return {'staged': True, 'backup': backup, 'created': created}


def restore_model(ctx):
    if not ctx or not ctx.get('staged'):
        return
    try:
        backup = ctx.get('backup')
        if backup and Path(backup).exists():
            shutil.copy2(backup, MODEL_CANONICAL)
        elif ctx.get('created') and MODEL_CANONICAL.exists():
            MODEL_CANONICAL.unlink()
    except Exception as exc:
        print(f'AVERTISSEMENT restauration modèle : {exc}')


def stage_excel_sources(selected_files, by_prefix, log):
    """
    Snapshot user sources, save current canonical data, then expose exactly
    23 compatibility aliases expected by the validated legacy generators.
    """
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    snapshot = SORTIES / '_staging' / f'sources_selectionnees_{stamp}'
    backup = SORTIES / 'backups' / f'donnees_brutes_{stamp}'
    snapshot.mkdir(parents=True, exist_ok=True)
    backup.mkdir(parents=True, exist_ok=True)

    # Snapshot first: this also protects the case selected_dir == donnees_brutes.
    snap_by_prefix = {}
    for n, src in by_prefix.items():
        dst = snapshot / f'{n:02d}.xlsx'
        shutil.copy2(src, dst)
        snap_by_prefix[n] = dst

    CANONICAL_DATA.mkdir(parents=True, exist_ok=True)
    for p in source_files(CANONICAL_DATA):
        shutil.copy2(p, backup / p.name)
        p.unlink()

    for n in range(1, 24):
        shutil.copy2(snap_by_prefix[n], CANONICAL_DATA / LEGACY_EXCEL_NAMES[n])

    log.append(f'Snapshot sources utilisateur : {snapshot}')
    log.append(f'Sauvegarde donnees_brutes : {backup}')
    return {'snapshot': snapshot, 'backup': backup}


def restore_excel_sources(ctx):
    if not ctx:
        return
    try:
        for p in source_files(CANONICAL_DATA):
            p.unlink()
        backup = Path(ctx['backup'])
        for p in source_files(backup):
            shutil.copy2(p, CANONICAL_DATA / p.name)
        shutil.rmtree(ctx['snapshot'], ignore_errors=True)
        print('Sources Excel originales restaurées.')
    except Exception as exc:
        print(f'AVERTISSEMENT restauration Excel : {exc}')


def _sans_accents(text):
    text = unicodedata.normalize('NFKD', str(text))
    return ''.join(ch for ch in text if not unicodedata.combining(ch))


def _source_sort_key(name):
    m = re.match(r'\s*(\d+)', str(name))
    return (int(m.group(1)) if m else 9999, str(name).lower())


def _source_label(name):
    value = Path(str(name)).stem
    value = re.sub(r'^\s*\d+\s*[.\-_]*\s*', '', value)
    value = re.sub(r'\s*_?\s*AS\s*_?\s*20\d{2}.*$', '', value, flags=re.I)
    value = re.sub(r'\s+', ' ', value).strip(' _-.')
    return value or Path(str(name)).stem


def _classer_source(name):
    n = _sans_accents(Path(str(name)).stem).lower()
    if 'transport' in n:
        return 'transport'
    if 'prix' in n:
        return 'prix'
    if 'environnement' in n:
        return 'environnement'
    if 'tourisme' in n:
        return 'tourisme'
    if any(k in n for k in ('agriculture', 'elevage', 'foret')):
        return 'agriculture'
    if 'sant' in n:
        return 'sante'
    if any(k in n for k in ('enseignement', 'education', 'formation professionnelle')):
        return 'education'
    if any(k in n for k in ('emploi', 'salaire', 'travail')):
        return 'emploi'
    if 'justice' in n:
        return 'justice'
    if any(k in n for k in ('culture', 'loisirs', 'entraide')):
        return 'culture'
    if any(k in n for k in ('habitat', 'construction', 'foncier', 'cadastre')):
        return 'habitat'
    if any(k in n for k in ('population', 'division administrative', 'demograph')):
        return 'population'
    if any(k in n for k in ('industrie', 'artisanat')):
        return 'industrie'
    if any(k in n for k in ('energie', 'eau')):
        return 'energie'
    return 'autre'


class Progression24:
    def __init__(self, original_files, log):
        self.sources = [p.name for p in sorted(original_files, key=lambda p: _source_sort_key(p.name))]
        self.log = log
        self.done = set()
        self.count = 0

    def _emit(self, name, note=None):
        if name in self.done:
            return
        self.done.add(name)
        self.count += 1
        label = _source_label(name)
        if note:
            label += ' — ' + note
        print(f'@@PROGRESS24@@ {self.count}/24 | {label}', flush=True)
        self.log.append(f'[PROGRESSION {self.count}/24] {name}' + (f' — {note}' if note else ''))

    def phase_terminee(self, phase, note=None):
        for name in self.sources:
            if _classer_source(name) == phase:
                self._emit(name, note)

    def terminer_sources_restantes(self):
        for name in self.sources:
            if name not in self.done:
                self._emit(name, 'source contrôlée / contenu conservé')
        if self.count != 23:
            raise RuntimeError(f'Progression incohérente : {self.count}/23 sources.')

    def validation_finale(self):
        if self.count != 23:
            self.terminer_sources_restantes()
        self.count = 24
        print('@@PROGRESS24@@ 24/24 | Validation finale', flush=True)
        self.log.append('[PROGRESSION 24/24] Validation finale')


def preflight(log):
    banner('PRE-CONTROLE')
    excels = source_files(CANONICAL_DATA)
    if len(excels) != 23:
        fail(f'23 alias Excel attendus, {len(excels)} détectés.', log)

    missing = [name for name in REQUIRED_SCRIPTS if not (SCRIPTS / name).exists()]
    if missing:
        fail('Scripts manquants : ' + ', '.join(missing), log)

    mapping = SORTIES / 'mapping_final_recettes.xlsx'
    if not mapping.exists():
        fail(f'Mapping requis introuvable : {mapping}', log)
    print('Pré-contrôle : OK')


def cleanup_outputs(final, log):
    banner('NETTOYAGE DES INTERMEDIAIRES')
    for p in list(OUTPUTS.values()) + [final]:
        if p.exists():
            p.unlink()
            log.append(f'Supprimé avant run : {p}')


def run_script(step_no, display_name, script_name, expected_output, log, year):
    banner(f'[{step_no}/16] {display_name}')
    script = SCRIPTS / script_name
    env = os.environ.copy()
    env['HCP_EDITION_YEAR'] = year
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    start = time.time()

    creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0) if os.name == 'nt' else 0
    proc = subprocess.run(
        [sys.executable, '-u', str(script)],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding='utf-8',
        errors='replace',
        env=env,
        creationflags=creationflags,
    )

    print(proc.stdout, end='' if proc.stdout.endswith('\n') else '\n')
    log.extend([
        '', '-' * 108,
        f'[{step_no}/16] {display_name}',
        f'Script : {script_name}',
        f'Code : {proc.returncode}',
        f'Durée : {time.time() - start:.1f}s',
        proc.stdout.rstrip(),
    ])

    if proc.returncode != 0:
        fail(f'{display_name} a échoué.', log)
    if expected_output is not None and not expected_output.exists():
        fail(f"{display_name} n'a pas créé {expected_output.name}.", log)


def copy_alias(src, dst, label, log):
    if not src.exists():
        fail(f'Impossible de préparer {label}: {src} absent.', log)
    shutil.copy2(src, dst)
    log.append(f'Alias {label}: {src} -> {dst}')


def no_op_step(step_no, name, reason, log):
    banner(f'[{step_no}/16] {name}')
    print(reason)
    log.extend(['', '-' * 108, f'[{step_no}/16] {name}', reason])


def validate_final(path, log):
    if not path.exists():
        fail('Word final absent.', log)
    if path.stat().st_size < 100_000:
        fail('Word final anormalement petit.', log)
    with path.open('rb') as f:
        if f.read(2) != b'PK':
            fail("Le résultat final n'est pas un DOCX valide.", log)


def main():
    args = parse_args()
    selected_word = Path(args.word)
    selected_excel_dir = Path(args.excel_dir)
    SORTIES.mkdir(parents=True, exist_ok=True)

    log = [
        '=' * 108,
        'RAPPORT — GENERATION MULTI-ANNEE TOUS SECTEURS',
        '=' * 108,
        f'Début : {datetime.now().isoformat(timespec="seconds")}',
        f'Word initial : {selected_word}',
        f'Dossier Excel : {selected_excel_dir}',
    ]

    model_ctx = None
    excel_ctx = None
    banner('ANNUAIRE HCP — GENERATION MULTI-ANNEE')

    try:
        original_files, by_prefix, year = validate_and_detect_sources(selected_excel_dir, log)
        final = final_path(year)
        os.environ['HCP_EDITION_YEAR'] = year
        print(f'@@YEAR@@ {year}', flush=True)
        print(f'Edition détectée : {year}')
        log.append(f'Edition détectée : {year}')

        ensure_education_reference(log)
        model_ctx = stage_model(selected_word, log)
        excel_ctx = stage_excel_sources(original_files, by_prefix, log)

        preflight(log)
        cleanup_outputs(final, log)
        progression = Progression24(original_files, log)

        run_script(1, 'Transport', 'generer_transport.py', OUTPUTS['transport'], log, year)
        progression.phase_terminee('transport')
        run_script(2, 'Prix', 'generer_prix.py', OUTPUTS['prix'], log, year)
        progression.phase_terminee('prix')
        run_script(3, 'Environnement', 'generer_environnement.py', OUTPUTS['env'], log, year)
        progression.phase_terminee('environnement')
        copy_alias(OUTPUTS['env'], OUTPUTS['env_alias'], 'entrée Tourisme', log)

        run_script(4, 'Tourisme', 'generer_tourisme.py', OUTPUTS['tourisme'], log, year)
        progression.phase_terminee('tourisme')
        run_script(5, 'Agriculture / Élevage / Forêt', 'generer_agriculture.py', OUTPUTS['agri'], log, year)
        progression.phase_terminee('agriculture')
        copy_alias(OUTPUTS['agri'], OUTPUTS['agri_alias'], 'entrée Santé', log)

        run_script(6, 'Santé', 'generer_sante_corrige.py', OUTPUTS['sante'], log, year)
        progression.phase_terminee('sante')

        # Education reference provides layout only; actual annual values are then
        # injected from the selected annual Education workbook.
        run_script(7, 'Enseignement — structure', 'appliquer_education_reference.py', OUTPUTS['education'], log, year)
        run_script(7, 'Enseignement — données de l\'édition', 'mettre_a_jour_education_multi.py', OUTPUTS['education'], log, year)
        no_op_step(8, 'Enseignement supérieur', 'Inclus et actualisé dans la phase Enseignement multi-année.', log)
        progression.phase_terminee('education')

        no_op_step(
            9,
            'Emploi / Salaires',
            "Le contenu est conservé lorsque le mapping annuel n'est pas validé ; aucune valeur n'est inventée.",
            log,
        )
        progression.phase_terminee('emploi', 'source contrôlée / contenu conservé')

        run_script(10, 'Justice', 'generer_justice.py', OUTPUTS['justice'], log, year)
        progression.phase_terminee('justice')
        run_script(11, 'Culture / Loisirs', 'generer_culture.py', OUTPUTS['culture'], log, year)
        progression.phase_terminee('culture')
        run_script(12, 'Habitat / Construction / Foncier', 'generer_habitat_foncier.py', OUTPUTS['habitat'], log, year)
        progression.phase_terminee('habitat')
        run_script(13, 'Population / Division administrative', 'generer_population_administration_verifie_v2.py', OUTPUTS['population'], log, year)
        progression.phase_terminee('population')
        run_script(14, 'Industrie / Artisanat', 'generer_industrie_verifie.py', OUTPUTS['industrie'], log, year)
        progression.phase_terminee('industrie')
        run_script(15, 'Énergie / Eau', 'generer_energie_eau_verifie_v2.py', OUTPUTS['energie'], log, year)
        progression.phase_terminee('energie')
        progression.terminer_sources_restantes()

        run_script(16, 'Finalisation multi-année', 'finaliser_annuaire_multi.py', final, log, year)
        run_script(16, 'Synthèse / encadrés multi-année', 'synchroniser_synthese_encadres_multi.py', final, log, year)

        validate_final(final, log)
        progression.validation_finale()

        log.extend([
            '', '=' * 108,
            'GENERATION MULTI-ANNEE : OK',
            f'Edition : {year}',
            f'Final : {final}',
            f'Fin : {datetime.now().isoformat(timespec="seconds")}',
        ])
        write_log(log)
        banner('GENERATION MULTI-ANNEE : OK')
        print(f'Word final : {final}')
        print(f'Rapport : {LOG}')

    except Exception as exc:
        log.append(f'ERREUR FINALE : {exc}')
        write_log(log)
        banner('GENERATION MULTI-ANNEE : ECHEC')
        print(str(exc))
        print(f'Rapport : {LOG}')
        sys.exit(1)

    finally:
        restore_excel_sources(excel_ctx)
        restore_model(model_ctx)


if __name__ == '__main__':
    main()
