from pathlib import Path
import subprocess
import sys
import time
import shutil
from datetime import datetime

ROOT = Path(".")
SCRIPTS = ROOT / "scripts"
SORTIES = ROOT / "sorties"
DONNEES = ROOT / "donnees_brutes"

BASE_VALIDEE = SORTIES / "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED.docx"
FINAL = SORTIES / "Annuaire_HCP_2025_FINAL.docx"
LOG = SORTIES / "rapport_generation_complete.txt"

# Backend V1:
# les chapitres déjà validés jusqu'à Enseignement sont contenus dans BASE_VALIDEE.
# On rejoue ensuite automatiquement tous les générateurs validés restants.
ETAPES = [
    {
        "nom": "Justice",
        "script": "generer_justice.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE.docx",
    },
    {
        "nom": "Culture / Loisirs",
        "script": "generer_culture.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE.docx",
    },
    {
        "nom": "Habitat / Foncier",
        "script": "generer_habitat_foncier.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER.docx",
    },
    {
        "nom": "Population / Division administrative",
        "script": "generer_population_administration_verifie_v2.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION.docx",
    },
    {
        "nom": "Industrie",
        "script": "generer_industrie_verifie.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE.docx",
    },
    {
        "nom": "Energie / Eau + revalidation Prix énergie/eau",
        "script": "generer_energie_eau_verifie_v2.py",
        "sortie": "Annuaire_HCP_FINAL_CLEAN_EDUCATION_CORRECTED_JUSTICE_CULTURE_HABITAT_FONCIER_POPULATION_INDUSTRIE_ENERGIE_EAU.docx",
    },
    {
        "nom": "Gel final + restauration contrôlée 4-4",
        "script": "corriger_et_geler_annuaire_final_v2.py",
        "sortie": "Annuaire_HCP_2025_FINAL.docx",
    },
]

# Fichiers intermédiaires exacts à supprimer avant un run.
# Cela empêche qu'un script choisisse par erreur un ancien résultat.
INTERMEDIAIRES = [
    SORTIES / e["sortie"]
    for e in ETAPES
    if e["sortie"] != FINAL.name
]


def banner(text):
    print()
    print("=" * 104)
    print(text)
    print("=" * 104)


def fail(message, log_lines):
    log_lines.append(f"ECHEC : {message}")
    LOG.write_text("\n".join(log_lines), encoding="utf-8")
    raise RuntimeError(message)


def preflight(log_lines):
    banner("PRE-CONTROLE")

    if not SCRIPTS.exists():
        fail(f"Dossier scripts introuvable : {SCRIPTS}", log_lines)

    if not DONNEES.exists():
        fail(f"Dossier donnees_brutes introuvable : {DONNEES}", log_lines)

    if not BASE_VALIDEE.exists():
        fail(
            "Base validée introuvable : "
            f"{BASE_VALIDEE}\n"
            "Cette version du backend part du Word validé après Enseignement.",
            log_lines
        )

    missing_scripts = [
        e["script"] for e in ETAPES
        if not (SCRIPTS / e["script"]).exists()
    ]

    if missing_scripts:
        fail(
            "Scripts manquants : " + ", ".join(missing_scripts),
            log_lines
        )

    excels = [
        p for p in DONNEES.glob("*.xlsx")
        if not p.name.startswith("~$")
    ]

    print(f"Base Word : {BASE_VALIDEE}")
    print(f"Excel détectés : {len(excels)}")
    print(f"Etapes backend : {len(ETAPES)}")

    log_lines.append(f"Base Word : {BASE_VALIDEE}")
    log_lines.append(f"Excel détectés : {len(excels)}")

    if len(excels) < 20:
        fail(
            f"Seulement {len(excels)} fichiers Excel détectés dans donnees_brutes. "
            "Le dossier source semble incomplet.",
            log_lines
        )

    print("Pré-contrôle : OK")


def nettoyer_intermediaires(log_lines):
    banner("NETTOYAGE DES ANCIENS INTERMEDIAIRES")

    removed = 0

    for p in INTERMEDIAIRES:
        if p.exists():
            p.unlink()
            print(f"Supprimé : {p.name}")
            log_lines.append(f"Ancien intermédiaire supprimé : {p}")
            removed += 1

    # Le fichier final précédent est conservé comme sauvegarde au lieu d'être écrasé sans trace.
    if FINAL.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = SORTIES / f"Annuaire_HCP_2025_FINAL_BACKUP_{stamp}.docx"
        shutil.copy2(FINAL, backup)
        print(f"Sauvegarde de l'ancien final : {backup.name}")
        log_lines.append(f"Sauvegarde ancien final : {backup}")

    if removed == 0:
        print("Aucun ancien intermédiaire à supprimer.")


def executer_etape(index, etape, log_lines):
    nom = etape["nom"]
    script = SCRIPTS / etape["script"]
    sortie = SORTIES / etape["sortie"]

    banner(f"[{index}/{len(ETAPES)}] {nom}")
    print(f"Script : {script}")

    before = sortie.stat().st_mtime if sortie.exists() else None
    start = time.time()

    proc = subprocess.run(
        [sys.executable, str(script)],
        cwd=ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8",
        errors="replace",
    )

    # Afficher exactement la sortie du sous-script.
    print(proc.stdout, end="" if proc.stdout.endswith("\n") else "\n")

    duration = time.time() - start

    log_lines.append("")
    log_lines.append("-" * 104)
    log_lines.append(f"ETAPE {index}/{len(ETAPES)} : {nom}")
    log_lines.append(f"Script : {script}")
    log_lines.append(f"Durée : {duration:.1f} s")
    log_lines.append(f"Code retour : {proc.returncode}")
    log_lines.append(proc.stdout.rstrip())

    if proc.returncode != 0:
        fail(
            f"L'étape '{nom}' a échoué. "
            f"Voir {LOG}.",
            log_lines
        )

    if not sortie.exists():
        fail(
            f"L'étape '{nom}' s'est terminée sans créer {sortie}.",
            log_lines
        )

    after = sortie.stat().st_mtime

    if before is not None and after <= before:
        fail(
            f"L'étape '{nom}' n'a pas régénéré son fichier de sortie.",
            log_lines
        )

    size_mb = sortie.stat().st_size / (1024 * 1024)

    print()
    print(f"Etape validée : {sortie.name} ({size_mb:.2f} Mo)")
    log_lines.append(f"Sortie validée : {sortie} ({size_mb:.2f} Mo)")


def validation_finale(log_lines):
    banner("VALIDATION BACKEND")

    if not FINAL.exists():
        fail("Le fichier final n'existe pas.", log_lines)

    if FINAL.stat().st_size < 100_000:
        fail(
            f"Le fichier final semble anormalement petit : {FINAL.stat().st_size} octets.",
            log_lines
        )

    # Contrôles binaires simples : un DOCX est un ZIP commençant par PK.
    with FINAL.open("rb") as f:
        magic = f.read(2)

    if magic != b"PK":
        fail("Le fichier final n'est pas un DOCX/ZIP valide.", log_lines)

    size_mb = FINAL.stat().st_size / (1024 * 1024)

    print(f"Final : {FINAL}")
    print(f"Taille : {size_mb:.2f} Mo")
    print("Validation backend : OK")

    log_lines.append("")
    log_lines.append("=" * 104)
    log_lines.append("VALIDATION BACKEND : OK")
    log_lines.append(f"Final : {FINAL}")
    log_lines.append(f"Taille : {size_mb:.2f} Mo")


def main():
    SORTIES.mkdir(parents=True, exist_ok=True)

    log_lines = [
        "=" * 104,
        "RAPPORT GENERATION COMPLETE — BACKEND V1",
        "=" * 104,
        f"Début : {datetime.now().isoformat(timespec='seconds')}",
        f"Python : {sys.executable}",
        "",
        "NOTE : cette version part de la base validée après Enseignement.",
        "Le repositionnement global des lignes Source est volontairement EXCLU.",
    ]

    banner("GENERATION ANNUAIRE HCP — BACKEND V1")
    print("Source positioning : ignoré, conformément à la demande.")
    print("Mode : un seul lancement, arrêt immédiat au premier échec.")

    try:
        preflight(log_lines)
        nettoyer_intermediaires(log_lines)

        for i, etape in enumerate(ETAPES, start=1):
            executer_etape(i, etape, log_lines)

        validation_finale(log_lines)

        log_lines.append(f"Fin : {datetime.now().isoformat(timespec='seconds')}")
        LOG.write_text("\n".join(log_lines), encoding="utf-8")

        banner("GENERATION COMPLETE : OK")
        print(f"Word final : {FINAL}")
        print(f"Rapport : {LOG}")
        print()
        print("Prochaine étape : interface graphique (GUI).")

    except Exception as exc:
        print()
        print("=" * 104)
        print("GENERATION COMPLETE : ECHEC")
        print("=" * 104)
        print(str(exc))
        print(f"Rapport : {LOG}")
        sys.exit(1)


if __name__ == "__main__":
    main()
