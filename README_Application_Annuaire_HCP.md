# Application d'Actualisation de l'Annuaire Statistique de Meknès

## 1\. Présentation

Cette application a été développée dans le cadre de l'automatisation de l'actualisation de l'Annuaire Statistique de Meknès.

Son objectif est d'automatiser l'intégration des données statistiques provenant de fichiers Excel dans un modèle Word de l'annuaire, tout en conservant la structure générale du document.

L'application peut être utilisée de deux manières :

1. **Depuis le code source Python**, pour la maintenance, le développement ou la reprise technique du projet.
2. **Depuis l'exécutable Windows (.exe)**, pour une utilisation simple par un utilisateur final sans avoir à lancer manuellement les scripts Python.



## 2\. Fonctionnement général

Le fonctionnement de l'application est le suivant :

1. Sélection d'un fichier Word initial compatible avec le modèle de l'annuaire.
2. Sélection d'un dossier contenant les 23 fichiers Excel sources.
3. Vérification automatique du fichier Word et du nombre de fichiers Excel.
4. Exécution successive des traitements sectoriels.
5. Mise à jour des tableaux de l'annuaire.
6. Validation finale.
7. Génération du document Word final.

Le document final généré est nommé :

Annuaire\_HCP\_2025\_FINAL.docx



## 3\. Configuration recommandée

### Système d'exploitation

* Windows 10 ou Windows 11
* Architecture 64 bits recommandée

### Version de Python

* **Version recommandée : Python 3.14.x 64 bits**
* **Version minimale : Python 3.11**

Le code utilise des fonctionnalités de Python nécessitant Python 3.10 ou une version plus récente.

Lors de l'installation de Python sous Windows, il est important de cocher l'option :

Add Python to PATH



## 4\. Bibliothèques et dépendances

Les principales bibliothèques utilisées dans l'environnement de déploiement sont :

* **OpenPyXL** : lecture et traitement des fichiers Excel.
* **Pandas** : traitement et manipulation de données tabulaires.
* **python-docx** : lecture et modification des documents Word.
* **lxml** : manipulation avancée de la structure XML des documents Word.
* **NumPy** : dépendance de calcul utilisée dans l'environnement de traitement.
* **Matplotlib** : bibliothèque prévue dans l'environnement packagé.
* **Jinja2** : bibliothèque prévue dans l'environnement packagé.
* **PyInstaller** : création de l'exécutable Windows.
* **Tkinter** : interface graphique de l'application.

### Important concernant Tkinter

Tkinter ne doit normalement pas être installé avec pip.

Il est inclus avec l'installation officielle de Python sous Windows lorsque le composant **Tcl/Tk and IDLE** est installé.

Pour vérifier son fonctionnement :

**cmd :**
python -m tkinter


Si une petite fenêtre graphique s'ouvre, Tkinter est correctement installé.

## 5\. Arborescence attendue du projet

Pour exécuter l'application depuis le code source, il est recommandé de conserver l'organisation suivante :


Annuaire\_HCP\_2025/
│
├── README.md
├── requirements.txt
│
├── scripts/
│   ├── interface\_annuaire\_exe.py
│   ├── generer\_annuaire\_tous\_secteurs\_v5.py
│   ├── generer\_annuaire\_tous\_secteurs\_exe.py
│   ├── appliquer\_education\_reference.py
│   ├── corriger\_et\_geler\_annuaire\_final\_v2.py
│   ├── corriger\_indicateurs\_synthese\_v2.py
│   ├── generer\_agriculture.py
│   ├── generer\_culture.py
│   ├── generer\_energie\_eau\_verifie\_v2.py
│   ├── generer\_environnement.py
│   ├── generer\_habitat\_foncier.py
│   ├── generer\_industrie\_verifie.py
│   ├── generer\_justice.py
│   ├── generer\_population\_administration\_verifie\_v2.py
│   ├── generer\_prix.py
│   ├── generer\_sante\_corrige.py
│   ├── generer\_tourisme.py
│   └── generer\_transport.py
│
├── references/
│   └── REFERENCE\_EDUCATION\_2025.docx
│
├── sorties/
│   └── mapping\_final\_recettes.xlsx
│
├── modeles/
└── donnees\_brutes/


### Fichiers à ne pas renommer

Les fichiers suivants sont appelés par le programme avec un nom précis :


references\\REFERENCE\_EDUCATION\_2025.docx
sorties\\mapping\_final\_recettes.xlsx


Il est donc recommandé de conserver exactement ces noms et ces emplacements.



## 6\. Installation de Python

Télécharger et installer Python 3.11 en version 64 bits.

Pendant l'installation :

1. Cocher **Add Python to PATH**.
2. Conserver le composant **Tcl/Tk and IDLE**.
3. Terminer l'installation.

Ensuite, ouvrir l'invite de commandes Windows (cmd) et vérifier :

**cmd:**
python --version


Vérifier également `pip` :

**cmd:**

pip --version


## 7\. Installation de l'environnement Python

Ouvrir une invite de commandes dans le dossier principal du projet.

Exemple :

**cmd:**

cd /d C:\\Annuaire\_HCP\_2025


### 7.1 Création d'un environnement virtuel

Il est recommandé d'utiliser un environnement virtuel afin d'isoler les bibliothèques du projet :

**cmd:**

python -m venv .venv


### 7.2 Activation de l'environnement virtuel

Sous Windows :

**cmd:**

.venv\\Scripts\\activate


Une fois l'environnement activé, le terminal doit normalement afficher (.venv) au début de la ligne.

### 7.3 Mise à jour de pip

**cmd:**

python -m pip install --upgrade pip


### 7.4 Installation des dépendances

Toutes les bibliothèques Python nécessaires sont regroupées dans le fichier `requirements.txt`.

Exécuter :

**cmd:**

python -m pip install -r requirements.txt




## 8\. Vérification de l'installation

Après l'installation des dépendances, vérifier Tkinter :

**cmd:**

python -m tkinter
```

Une petite fenêtre doit apparaître.

Il est également possible de vérifier les principales bibliothèques avec :

**cmd:**

python -c "import openpyxl, pandas, docx, lxml; print('Bibliothèques principales : OK')"


Le résultat attendu est :

Bibliothèques principales : OK


## 9\. Exécution depuis le code source

### Important

La commande doit être exécutée depuis le **dossier principal du projet**, c'est-à-dire le dossier contenant `scripts`, `references` et `sorties`.

Exemple :

**cmd:**

cd /d C:\\Annuaire\_HCP\_2025


Activer l'environnement virtuel si nécessaire :

**cmd:**

.venv\\Scripts\\activate


Lancer ensuite l'application :

**cmd:**

python scripts\\interface\_annuaire\_exe.py


L'interface graphique du générateur s'ouvre.



## 10\. Fichiers d'entrée nécessaires

L'application demande deux éléments à l'utilisateur.

### 10.1 Fichier Word initial

Cliquer sur :

**Choisir le Word**

Le fichier doit être au format `.docx`.

### 10.2 Dossier des fichiers Excel

Cliquer sur :

**Choisir le dossier**

Le dossier sélectionné doit contenir **exactement 23 fichiers `.xlsx`** correspondant aux sources statistiques attendues.



## 11\. Génération de l'annuaire

Lorsque le fichier Word et les 23 fichiers Excel ont été validés, le bouton :

**GÉNÉRER L'ANNUAIRE**

devient disponible.

Cliquer sur ce bouton pour commencer le traitement.

L'interface affiche l'avancement de la génération sous la forme :

Étape 1/24
Étape 2/24
...
Étape 24/24


Les 24 étapes correspondent au suivi des 23 sources Excel et à la validation finale.

Pendant l'exécution, il est recommandé de :

* ne pas déplacer les fichiers sources ;
* ne pas renommer les fichiers utilisés par le programme ;
* attendre la fin du traitement avant de fermer l'application.

À la fin du traitement, l'interface affiche :

`Terminé avec succès`



## 12\. Emplacement du document final

### Exécution depuis le code source

Lorsque l'application est lancée depuis le code source, le document final est enregistré dans :

sorties\\Annuaire\_HCP\_2025\_FINAL.docx
```

Le chemin est relatif au dossier principal du projet.

Exemple :

C:\\Annuaire\_HCP\_2025\\sorties\\Annuaire\_HCP\_2025\_FINAL.docx
```

Un rapport technique de génération est également produit dans le dossier `sorties` :
sorties\\rapport\_generation\_tous\_secteurs.txt
```

Ce rapport peut être utilisé pour vérifier le déroulement du traitement en cas de problème.



## 13\. Utilisation de la version Windows (.exe)

Pour une utilisation quotidienne, la version exécutable est la méthode la plus simple.

Nom de l'application :


Generateur\_Annuaire\_Statistique\_Meknes\_2025\_HCP.exe



L'utilisateur peut lancer l'application par double-clic.

Dans cette version, Python n'a normalement pas besoin d'être installé sur le poste utilisateur, car l'environnement nécessaire est intégré à l'exécutable.

Au premier lancement, l'application prépare automatiquement son dossier de travail dans :

C:\\Users\\<NomUtilisateur>\\Documents\\Annuaire\_HCP\_2025
Le document final est alors généré dans :

C:\\Users\\<NomUtilisateur>\\Documents\\Annuaire\_HCP\_2025\\sorties\\Annuaire\_HCP\_2025\_FINAL.docx
Les ressources intégrées nécessaires au programme sont également copiées dans ce dossier de travail lorsqu'elles n'existent pas encore.

## 14\. Création de l'exécutable Windows

La création de l'exécutable est réservée à la maintenance ou au redéploiement de l'application.

PyInstaller doit être installé :

**cmd:**

python -m pip install pyinstaller
```

## 15\. Remarques concernant les éditions futures

La version actuelle de l'application a été développée et validée avec la structure de l'Annuaire Statistique 2025.

Le projet peut servir de base aux éditions futures.

Cependant, avant d'utiliser l'application avec une nouvelle édition, il faut vérifier :

* que la structure du modèle Word n'a pas changé ;
* que les fichiers Excel conservent une organisation compatible ;
* que les tableaux et colonnes utilisés pour le mapping restent identiques ;
* que les règles sectorielles restent applicables.

Si les structures changent, le mapping ou certains scripts devront être adaptés avant la génération.

L'application ne doit donc pas être considérée comme automatiquement compatible avec toute nouvelle année sans contrôle préalable.

## 16\. Résumé des commandes principales

### Installation

**cmds:**

cd /d C:\\Annuaire\_HCP\_2025
python -m venv .venv
.venv\\Scripts\\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Vérification

**cmds:**

python --version
python -m tkinter
python -c "import openpyxl, pandas, docx, lxml; print('Bibliothèques principales : OK')"
```

### Lancement

**cmds:**

python scripts\\interface\_annuaire\_exe.py
```

## 17\. Auteur et contexte

**Projet :** Automatisation de l'actualisation de l'Annuaire Statistique de Meknès  
**Développement :** RABIH Houda  
**Organisme d'accueil :** Haut-Commissariat au Plan – Direction Provinciale de Meknès  
**Encadrant professionnel :** M. HOUNNI Mustapha  
**Année de référence de la version validée :** 2025

\---

## 18\. Note de maintenance

Pour préserver le bon fonctionnement du projet :

* conserver l'arborescence des dossiers ;
* éviter de renommer les scripts appelés par le pipeline ;
* conserver les deux ressources de référence aux emplacements prévus ;
* effectuer une sauvegarde avant toute modification ;
* tester toute nouvelle version sur une copie des données avant son déploiement.

