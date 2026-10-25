from pathlib import Path
import argparse
import subprocess
import sys
import time
import shutil
import re
import unicodedata
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
    "corriger_indicateurs_synthese_v2.py",
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
    """
    Prépare le modèle sans JAMAIS supprimer ni écraser le fichier choisi
    par l'utilisateur.
    """
    if not selected.exists():
        fail(f"Word initial introuvable : {selected}", log)
    if selected.suffix.lower() != ".docx":
        fail("Le Word initial doit être un fichier .docx.", log)

    MODELES.mkdir(parents=True, exist_ok=True)

    selected_resolved = selected.resolve()
    canonical_resolved = MODEL_CANONICAL.resolve()

    # Si l'utilisateur choisit directement modeles\\modeles.docx,
    # on l'utilise tel quel. Aucune copie, aucune suppression, aucune restauration.
    if selected_resolved == canonical_resolved:
        print(f"Word initial utilisé directement : {selected}")
        log.append(f"Word utilisateur utilisé directement : {selected}")
        return {"staged": False, "backup": None, "created": False}

    backup = None
    created = not MODEL_CANONICAL.exists()

    if MODEL_CANONICAL.exists():
        bdir = SORTIES / "backups"
        bdir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = bdir / f"modele_original_{stamp}.docx"
        shutil.copy2(MODEL_CANONICAL, backup)
        print(f"Sauvegarde modèle de travail existant : {backup}")

    shutil.copy2(selected, MODEL_CANONICAL)
    print(f"Word utilisateur : {selected}")
    print(f"Copie de travail créée : {MODEL_CANONICAL}")
    log.append(f"Word utilisateur : {selected}")
    if backup:
        log.append(f"Sauvegarde modèle précédent : {backup}")

    return {"staged": True, "backup": backup, "created": created}


def restore_model(ctx):
    """
    Ne touche qu'à la copie de travail modeles\\modeles.docx quand elle a
    réellement été créée/remplacée par stage_model().
    Le fichier utilisateur sélectionné n'est JAMAIS supprimé.
    """
    if not ctx or not ctx.get("staged"):
        return

    backup = ctx.get("backup")
    created = ctx.get("created", False)

    try:
        if backup and Path(backup).exists():
            shutil.copy2(backup, MODEL_CANONICAL)
            print("Modèle de travail précédent restauré.")
        elif created and MODEL_CANONICAL.exists():
            MODEL_CANONICAL.unlink()
            print("Copie temporaire du modèle retirée.")
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



def _sans_accents(text):
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _source_sort_key(path):
    """
    Trie d'abord par numéro au début du nom (1., 2., ... 23.),
    puis alphabétiquement.
    """
    m = re.match(r"\s*(\d+)", path.name)
    return (int(m.group(1)) if m else 9999, path.name.lower())


def _source_label(path):
    """
    Affichage lisible à partir du vrai nom du fichier sélectionné.
    On garde le nom du secteur tel qu'il apparaît dans le fichier,
    en retirant seulement les suffixes techniques _AS 2025...
    """
    name = path.stem
    name = re.sub(r"^\s*\d+\s*[.\-_]*\s*", "", name)
    name = re.sub(r"\s*_?\s*AS\s*20\d{2}.*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip(" _-.")
    return name or path.stem


def _classer_source_excel(path):
    """
    Associe chaque vrai fichier Excel à la phase interne qui le couvre.
    Cette fonction ne modifie PAS le traitement des données.
    Elle sert uniquement au suivi 23 sources + validation.
    """
    n = _sans_accents(path.stem).lower()

    if "transport" in n:
        return "transport"
    if "prix" in n or "indice des prix" in n:
        return "prix"
    if "environnement" in n or "meteor" in n:
        return "environnement"
    if "tourisme" in n:
        return "tourisme"
    if any(k in n for k in ("agriculture", "elevage", "foret", "forest")):
        return "agriculture"
    if "sante" in n:
        return "sante"
    if any(k in n for k in ("enseignement", "education", "formation professionnelle")):
        return "education"
    if any(k in n for k in ("emploi", "salaire", "travail")):
        return "emploi"
    if "justice" in n:
        return "justice"
    if any(k in n for k in ("culture", "loisirs", "entraide")):
        return "culture"
    if any(k in n for k in ("habitat", "construction", "foncier", "cadastre")):
        return "habitat"
    if any(k in n for k in ("population", "division administrative", "demograph")):
        return "population"
    if any(k in n for k in ("industrie", "artisanat", "commerce")):
        return "industrie"
    if any(k in n for k in ("energie", "eau")):
        return "energie"

    # Source présente parmi les 23 mais sans phase spécifique dans
    # le backend actuel. Elle sera affichée comme source contrôlée
    # avant la validation finale, sans prétendre qu'un tableau a été
    # modifié à partir d'elle.
    return "autre"


class Progression24:
    """
    Flux d'affichage indépendant des 16 phases techniques historiques.

    L'interface voit exactement :
        23 sources Excel réelles
        + 1 validation finale
        = 24 étapes
    """
    def __init__(self, log):
        self.sources = sorted(
            [
                p for p in CANONICAL_DATA.glob("*.xlsx")
                if not p.name.startswith("~$")
            ],
            key=_source_sort_key,
        )

        if len(self.sources) != 23:
            fail(
                f"Progression 24 impossible : 23 Excel attendus, "
                f"{len(self.sources)} détectés.",
                log
            )

        self.log = log
        self.done = set()
        self.count = 0

    def _emit_source(self, path, note=None):
        resolved = str(path.resolve())
        if resolved in self.done:
            return

        self.done.add(resolved)
        self.count += 1

        label = _source_label(path)
        if note:
            label = f"{label} — {note}"

        # Marqueur machine lu par l'interface V4.
        print(
            f"@@PROGRESS24@@ {self.count}/24 | {label}",
            flush=True
        )
        self.log.append(
            f"[PROGRESSION {self.count}/24] {path.name}"
            + (f" — {note}" if note else "")
        )

    def phase_terminee(self, phase):
        for path in self.sources:
            if _classer_source_excel(path) == phase:
                self._emit_source(path)

    def terminer_sources_restantes(self):
        """
        Les fichiers non associés à une phase dédiée sont présents et
        contrôlés, mais le backend ne doit pas inventer une mise à jour.
        """
        for path in self.sources:
            if str(path.resolve()) not in self.done:
                self._emit_source(path, "source contrôlée / contenu conservé")

        if self.count != 23:
            fail(
                f"Progression incohérente : {self.count}/23 sources comptabilisées.",
                self.log
            )

    def validation_finale(self):
        if self.count != 23:
            self.terminer_sources_restantes()

        self.count = 24
        print(
            "@@PROGRESS24@@ 24/24 | Validation finale",
            flush=True
        )
        self.log.append("[PROGRESSION 24/24] Validation finale")


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

    model_ctx = None
    excel_backup = None

    banner("ANNUAIRE HCP — GENERATION COMPLETE — FINAL V5 / PROGRESSION 24")

    try:
        ensure_education_reference(log)

        model_ctx = stage_model(selected_word, log)
        excel_backup = stage_excel_folder(selected_excel_dir, log)

        preflight(log)
        cleanup_outputs(log)

        progression = Progression24(log)

        run_script(1, "Transport", "generer_transport.py", OUTPUTS["transport"], log)
        progression.phase_terminee("transport")
        run_script(2, "Prix", "generer_prix.py", OUTPUTS["prix"], log)
        progression.phase_terminee("prix")
        run_script(3, "Environnement", "generer_environnement.py", OUTPUTS["env"], log)
        progression.phase_terminee("environnement")

        copy_alias(OUTPUTS["env"], OUTPUTS["env_alias"], "entrée Tourisme", log)

        run_script(4, "Tourisme", "generer_tourisme.py", OUTPUTS["tourisme"], log)
        progression.phase_terminee("tourisme")
        run_script(5, "Agriculture / Élevage / Forêt", "generer_agriculture.py", OUTPUTS["agri"], log)
        progression.phase_terminee("agriculture")

        copy_alias(OUTPUTS["agri"], OUTPUTS["agri_alias"], "entrée Santé", log)

        run_script(6, "Santé", "generer_sante_corrige.py", OUTPUTS["sante"], log)
        progression.phase_terminee("sante")

        run_script(7, "Enseignement", "appliquer_education_reference.py", OUTPUTS["education"], log)

        no_op_step(
            8,
            "Enseignement supérieur",
            "Le bloc Enseignement supérieur 4-x est inclus dans la référence validée appliquée à l'étape 7.",
            log
        )
        progression.phase_terminee("education")

        no_op_step(
            9,
            "Emploi / Salaires",
            "Aucune nouvelle source Excel exploitable n'a été validée pour ce secteur; "
            "le contenu Word est contrôlé et conservé.",
            log
        )
        progression.phase_terminee("emploi")

        run_script(10, "Justice", "generer_justice.py", OUTPUTS["justice"], log)
        progression.phase_terminee("justice")
        run_script(11, "Culture / Loisirs", "generer_culture.py", OUTPUTS["culture"], log)
        progression.phase_terminee("culture")
        run_script(12, "Habitat / Construction / Foncier", "generer_habitat_foncier.py", OUTPUTS["habitat"], log)
        progression.phase_terminee("habitat")
        run_script(13, "Population / Division administrative", "generer_population_administration_verifie_v2.py", OUTPUTS["population"], log)
        progression.phase_terminee("population")
        run_script(14, "Industrie / Artisanat", "generer_industrie_verifie.py", OUTPUTS["industrie"], log)
        progression.phase_terminee("industrie")
        run_script(15, "Énergie / Eau", "generer_energie_eau_verifie_v2.py", OUTPUTS["energie"], log)
        progression.phase_terminee("energie")
        progression.terminer_sources_restantes()
        run_script(
            16,
            "Validation finale",
            "corriger_et_geler_annuaire_final_v2.py",
            FINAL,
            log
        )

        run_script(
            16,
            "Synchronisation synthèse / encadrés",
            "corriger_indicateurs_synthese_v2.py",
            FINAL,
            log
        )

        validate_final(log)
        progression.validation_finale()

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
        restore_model(model_ctx)


if __name__ == "__main__":
    main()
