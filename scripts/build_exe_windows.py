from pathlib import Path
import os
import py_compile
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"

ENTRY = SCRIPTS / "interface_annuaire_exe.py"
BACKEND = SCRIPTS / "generer_annuaire_tous_secteurs_exe.py"

WORKERS = [
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

ASSETS = [
    (
        ROOT / "sorties" / "mapping_final_recettes.xlsx",
        "assets",
    ),
    (
        ROOT / "references" / "REFERENCE_EDUCATION_2025.docx",
        "assets",
    ),
]

EXE_NAME = "Generateur_Annuaire_Statistique_Meknes_2025_HCP"


def add_data(src, dest):
    return f"{src}{os.pathsep}{dest}"


def fail(message):
    print("\nERREUR :", message)
    raise SystemExit(1)


def main():
    if os.name != "nt":
        fail("Le build EXE doit être exécuté sous Windows.")

    files = [ENTRY, BACKEND]
    files += [SCRIPTS / name for name in WORKERS]
    files += [src for src, _ in ASSETS]

    missing = [p for p in files if not p.exists()]

    if missing:
        print("Fichiers manquants :")
        for p in missing:
            print(" -", p)
        raise SystemExit(2)

    print("Vérification des scripts...")
    for p in [ENTRY, BACKEND] + [SCRIPTS / name for name in WORKERS]:
        py_compile.compile(str(p), doraise=True)
    print("OK")

    try:
        import PyInstaller  # noqa: F401
    except Exception:
        fail(
            "PyInstaller absent. Lance d'abord : "
            "python -m pip install pyinstaller"
        )

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name",
        EXE_NAME,
    ]

    cmd += ["--add-data", add_data(BACKEND, "scripts")]

    for name in WORKERS:
        cmd += [
            "--add-data",
            add_data(SCRIPTS / name, "scripts"),
        ]

    for source, destination in ASSETS:
        cmd += [
            "--add-data",
            add_data(source, destination),
        ]

    # Workers are executed dynamically with runpy, so collect their
    # dependencies explicitly.
    for package in (
        "docx",
        "lxml",
        "openpyxl",
        "pandas",
        "numpy",
        "matplotlib",
        "jinja2",
    ):
        cmd += ["--collect-all", package]

    cmd.append(str(ENTRY))

    print("\nConstruction de l'EXE...")
    print("Cela peut prendre plusieurs minutes.\n")

    result = subprocess.run(cmd, cwd=ROOT)

    if result.returncode != 0:
        fail("PyInstaller a échoué.")

    exe = ROOT / "dist" / f"{EXE_NAME}.exe"

    if not exe.exists():
        fail(f"EXE introuvable après build : {exe}")

    print("\n" + "=" * 90)
    print("EXE CREE AVEC SUCCES")
    print("=" * 90)
    print(exe)
    print()
    print("Le PC cible n'a pas besoin de Python.")
    print(
        r"Les fichiers de travail seront créés dans "
        r"%USERPROFILE%\Documents\Annuaire_HCP_2025"
    )


if __name__ == "__main__":
    main()
