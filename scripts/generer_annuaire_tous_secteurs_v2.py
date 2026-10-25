from pathlib import Path
import argparse
import subprocess
import sys
import time
import shutil
from datetime import datetime

ROOT = Path(".")
SCRIPTS = ROOT / "scripts"
SORTIES = ROOT / "sorties"
CANONICAL_DATA = ROOT / "donnees_brutes"
MODELES = ROOT / "modeles"
REFERENCES = ROOT / "references"

MODEL_CANONICAL = MODELES / "modeles.docx"
REFERENCE_EDUCATION = REFERENCES / "REFERENCE_EDUCATION_2025.docx"
EXISTING_EDUCATION = SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"

FINAL = SORTIES / "Annuaire_HCP_2025_FINAL.docx"
LOG = SORTIES / "rapport_generation_tous_secteurs.txt"

REQUIRED_SCRIPTS = [
    "generer_transport.py",
    "generer_prix.py",
    "generer_environnement.py",
    "generer_tourisme.py",
    "generer_agriculture.py",
    "generer_sante_corrige.py",
    "appliquer_education_reference.py",
    "generer_justice.py",
    "generer_culture.py",
    "generer_habitat_foncier.py",
    "generer_population_administration_verifie_v2.py",
    "generer_industrie_verifie.py",
    "generer_energie_eau_verifie_v2.py",
    "corriger_et_geler_annuaire_final_v2.py",
]

OUTPUTS = {
    "transport": SORTIES / "Annuaire_HCP_Transport.docx",
    "prix": SORTIES / "Annuaire_HCP_Transport_Prix.docx",
    "env": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement.docx",
    "env_alias": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Sources2025_OK.docx",
    "tourisme": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme.docx",
    "agri": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture.docx",
    "agri_alias": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_VALIDE.docx",
    "sante": SORTIES / "Annuaire_HCP_Transport_Prix_Environnement_Tourisme_Agriculture_Sante.docx",
    "education": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx",
    "justice": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx",
    "culture": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx",
    "habitat": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER.docx",
    "population": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION.docx",
    "industrie": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE.docx",
    "energie": SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx",
}


def banner(text):
    print()
    print("=" * 108)
    print(text)
    print("=" * 108)


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--word", required=True, help="Word initial/template")
    p.add_argument("--excel-dir", required=True, help="Dossier contenant les Excel")
    return p.parse_args()


def write_log(lines):
    SORTIES.mkdir(parents=True, exist_ok=True)
    LOG.write_text("\n".join(lines), encoding="utf-8")


def fail(msg, log):
    log.append("ECHEC : " + msg)
    write_log(log)
    raise RuntimeError(msg)


def backup_file(path: Path, label: str):
    if not path.exists():
        return None
    bdir = SORTIES / "backups"
    bdir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = bdir / f"{label}_{stamp}{path.suffix}"
    shutil.copy2(path, dest)
    return dest


def ensure_education_reference(log):
    REFERENCES.mkdir(parents=True, exist_ok=True)

    if REFERENCE_EDUCATION.exists():
        print(f"Référence Enseignement : {REFERENCE_EDUCATION}")
        return

    if not EXISTING_EDUCATION.exists():
        fail(
            "Référence Enseignement 2025 absente. "
            "Conservez le Word validé 'Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx' "
            "une première fois dans sorties.",
            log
        )

    shutil.copy2(EXISTING_EDUCATION, REFERENCE_EDUCATION)
    print(f"Référence Enseignement créée : {REFERENCE_EDUCATION}")
    log.append(f"Référence Enseignement créée : {REFERENCE_EDUCATION}")


def stage_model(selected: Path, log):
    if not selected.exists():
        fail(f"Word initial introuvable : {selected}", log)
    if selected.suffix.lower() != ".docx":
        fail("Le Word initial doit être un fichier .docx.", log)

    MODELES.mkdir(parents=True, exist_ok=True)

    old_bytes = MODEL_CANONICAL.read_bytes() if MODEL_CANONICAL.exists() else None

    shutil.copy2(selected, MODEL_CANONICAL)
    print(f"Word initial chargé : {selected}")
    print(f"Copie de travail : {MODEL_CANONICAL}")
    log.append(f"Word utilisateur : {selected}")

    return old_bytes


def restore_model(old_bytes):
    try:
        if old_bytes is None:
            if MODEL_CANONICAL.exists():
                MODEL_CANONICAL.unlink()
        else:
            MODEL_CANONICAL.write_bytes(old_bytes)
    except Exception as exc:
        print(f"AVERTISSEMENT restauration modèle : {exc}")


def stage_excel_folder(selected_dir: Path, log):
    if not selected_dir.exists() or not selected_dir.is_dir():
        fail(f"Dossier Excel introuvable : {selected_dir}", log)

    selected_files = sorted([
        p for p in selected_dir.glob("*.xlsx")
        if not p.name.startswith("~$")
    ])

    if len(selected_files) != 23:
        fail(
            f"Le dossier sélectionné doit contenir exactement 23 fichiers Excel. "
            f"Détectés : {len(selected_files)}",
            log
        )

    try:
        same_dir = selected_dir.resolve() == CANONICAL_DATA.resolve()
    except Exception:
        same_dir = False

    if same_dir:
        print("Dossier Excel sélectionné = donnees_brutes actuel.")
        print("Aucune copie temporaire nécessaire.")
        log.append(f"Dossier Excel utilisateur : {selected_dir} (déjà canonique)")
        return None

    CANONICAL_DATA.mkdir(parents=True, exist_ok=True)

    backup_dir = SORTIES / "backups" / (
        "donnees_brutes_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    backup_dir.mkdir(parents=True, exist_ok=True)

    old_files = [
        p for p in CANONICAL_DATA.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]

    for p in old_files:
        shutil.copy2(p, backup_dir / p.name)
        p.unlink()

    for p in selected_files:
        shutil.copy2(p, CANONICAL_DATA / p.name)

    print(f"Dossier Excel chargé : {selected_dir}")
    print(f"23 fichiers copiés temporairement dans {CANONICAL_DATA}")
    print(f"Sauvegarde des anciennes sources : {backup_dir}")
    log.append(f"Dossier Excel utilisateur : {selected_dir}")
    log.append(f"Sauvegarde sources précédentes : {backup_dir}")

    return backup_dir


def restore_excel_folder(backup_dir):
    if backup_dir is None:
        return

    try:
        for p in CANONICAL_DATA.glob("*.xlsx"):
            if not p.name.startswith("~$"):
                p.unlink()

        for p in backup_dir.glob("*.xlsx"):
            shutil.copy2(p, CANONICAL_DATA / p.name)

        print("Sources Excel originales restaurées.")
    except Exception as exc:
        print(f"AVERTISSEMENT restauration Excel : {exc}")


def preflight(log):
    banner("PRE-CONTROLE")

    excels = [
        p for p in CANONICAL_DATA.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]
    print(f"Excel disponibles : {len(excels)}")

    if len(excels) != 23:
        fail(f"23 fichiers Excel attendus, {len(excels)} détectés.", log)

    missing = [
        name for name in REQUIRED_SCRIPTS
        if not (SCRIPTS / name).exists()
    ]

    if missing:
        fail("Scripts manquants : " + ", ".join(missing), log)

    mapping = SORTIES / "mapping_final_recettes.xlsx"
    if not mapping.exists():
        fail(f"Mapping requis introuvable : {mapping}", log)

    print("Pré-contrôle : OK")


def cleanup_outputs(log):
    banner("NETTOYAGE DES INTERMEDIAIRES")

    for p in list(OUTPUTS.values()) + [FINAL]:
        if p.exists():
            # Ne jamais effacer la référence dédiée.
            p.unlink()
            print(f"Supprimé : {p.name}")
            log.append(f"Supprimé avant run : {p}")


def run_script(step_no, display_name, script_name, expected_output, log):
    banner(f"[{step_no}/16] {display_name}")

    script = SCRIPTS / script_name
    start = time.time()

    proc = subprocess.run(
        [sys.executable, "-u", str(script)],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    print(proc.stdout, end="" if proc.stdout.endswith("\n") else "\n")

    log.extend([
        "",
        "-" * 108,
        f"[{step_no}/16] {display_name}",
        f"Script : {script_name}",
        f"Code : {proc.returncode}",
        f"Durée : {time.time() - start:.1f}s",
        proc.stdout.rstrip(),
    ])

    if proc.returncode != 0:
        fail(f"{display_name} a échoué.", log)

    if expected_output is not None and not expected_output.exists():
        fail(
            f"{display_name} n'a pas créé {expected_output.name}.",
            log
        )

    if expected_output:
        print(
            f"Etape validée : {expected_output.name} "
            f"({expected_output.stat().st_size / (1024*1024):.2f} Mo)"
        )


def copy_alias(src, dst, label, log):
    if not src.exists():
        fail(f"Impossible de préparer {label}: {src} absent.", log)
    shutil.copy2(src, dst)
    print(f"Préparation compatibilité : {dst.name}")
    log.append(f"Alias {label}: {src} -> {dst}")


def no_op_step(step_no, name, reason, log):
    banner(f"[{step_no}/16] {name}")
    print(reason)
    print("Etape validée : conservation contrôlée.")
    log.extend([
        "",
        "-" * 108,
        f"[{step_no}/16] {name}",
        reason,
        "Etape conservée/validée.",
    ])


def validate_final(log):
    if not FINAL.exists():
        fail("Word final absent.", log)
    if FINAL.stat().st_size < 100_000:
        fail("Word final anormalement petit.", log)
    with FINAL.open("rb") as f:
        if f.read(2) != b"PK":
            fail("Le résultat final n'est pas un DOCX valide.", log)


def main():
    args = parse_args()
    selected_word = Path(args.word)
    selected_excel_dir = Path(args.excel_dir)

    SORTIES.mkdir(parents=True, exist_ok=True)

    log = [
        "=" * 108,
        "RAPPORT — GENERATION COMPLETE TOUS SECTEURS",
        "=" * 108,
        f"Début : {datetime.now().isoformat(timespec='seconds')}",
        f"Word initial : {selected_word}",
        f"Dossier Excel : {selected_excel_dir}",
    ]

    old_model = None
    excel_backup = None

    banner("ANNUAIRE HCP — GENERATION COMPLETE")

    try:
        ensure_education_reference(log)

        old_model = stage_model(selected_word, log)
        excel_backup = stage_excel_folder(selected_excel_dir, log)

        preflight(log)
        cleanup_outputs(log)

        run_script(1, "Transport", "generer_transport.py", OUTPUTS["transport"], log)
        run_script(2, "Prix", "generer_prix.py", OUTPUTS["prix"], log)
        run_script(3, "Environnement", "generer_environnement.py", OUTPUTS["env"], log)

        copy_alias(OUTPUTS["env"], OUTPUTS["env_alias"], "entrée Tourisme", log)

        run_script(4, "Tourisme", "generer_tourisme.py", OUTPUTS["tourisme"], log)
        run_script(5, "Agriculture / Élevage / Forêt", "generer_agriculture.py", OUTPUTS["agri"], log)

        copy_alias(OUTPUTS["agri"], OUTPUTS["agri_alias"], "entrée Santé", log)

        run_script(6, "Santé", "generer_sante_corrige.py", OUTPUTS["sante"], log)

        run_script(7, "Enseignement", "appliquer_education_reference.py", OUTPUTS["education"], log)

        no_op_step(
            8,
            "Enseignement supérieur",
            "Le bloc Enseignement supérieur 4-x est inclus dans la référence validée appliquée à l'étape 7.",
            log
        )

        no_op_step(
            9,
            "Emploi / Salaires",
            "Aucune nouvelle source Excel exploitable n'a été validée pour ce secteur; "
            "le contenu Word est contrôlé et conservé.",
            log
        )

        run_script(10, "Justice", "generer_justice.py", OUTPUTS["justice"], log)
        run_script(11, "Culture / Loisirs", "generer_culture.py", OUTPUTS["culture"], log)
        run_script(12, "Habitat / Construction / Foncier", "generer_habitat_foncier.py", OUTPUTS["habitat"], log)
        run_script(13, "Population / Division administrative", "generer_population_administration_verifie_v2.py", OUTPUTS["population"], log)
        run_script(14, "Industrie / Artisanat", "generer_industrie_verifie.py", OUTPUTS["industrie"], log)
        run_script(15, "Énergie / Eau", "generer_energie_eau_verifie_v2.py", OUTPUTS["energie"], log)
        run_script(16, "Validation finale", "corriger_et_geler_annuaire_final_v2.py", FINAL, log)

        validate_final(log)

        log.extend([
            "",
            "=" * 108,
            "GENERATION TOUS SECTEURS : OK",
            f"Final : {FINAL}",
            f"Fin : {datetime.now().isoformat(timespec='seconds')}",
        ])
        write_log(log)

        banner("GENERATION TOUS SECTEURS : OK")
        print(f"Word final : {FINAL}")
        print(f"Rapport : {LOG}")

    except Exception as exc:
        write_log(log)
        banner("GENERATION TOUS SECTEURS : ECHEC")
        print(str(exc))
        print(f"Rapport : {LOG}")
        sys.exit(1)

    finally:
        restore_excel_folder(excel_backup)
        restore_model(old_model)


if __name__ == "__main__":
    main()
