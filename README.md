# DDT PDF Excel

Outil de Diagne Digital & Tech : recherche un numéro de référence (ex. `COLIS-2026-00481`)
dans tous les PDF du dossier `PDF_Colis` (sous-dossiers compris) et écrit les nombres
trouvés dans `resultat.xlsx`.

## Ce que fait l'outil

- Trouve la référence même si la ponctuation ou la casse diffèrent
  (`COLIS-2026-00481` = `colis 2026 00481` = `COLIS/2026/00481`),
  sans confondre avec une référence plus longue (`COLIS-2026-004810`).
- Extrait les nombres des pages où la référence apparaît (`850 000`, `25,5`, `15/10/2026`…).
- Feuille **Résultats** : fichier (lien cliquable), page(s), nombres extraits.
- Feuille **Non analysés** : PDF sans texte (scans, OCR nécessaire) ou illisibles.

## Installation (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Pour générer les PDF de test : `pip install -r requirements-dev.txt`.

## Lancement

```powershell
python main.py COLIS-2026-00481
```

Sans argument, le numéro est demandé au clavier. Ferme `resultat.xlsx` dans Excel avant de relancer.

Options :

```powershell
python main.py COLIS-2026-00481 --dossier "D:\Dossier client\PDF" --sortie "$env:TEMP\resultat.xlsx"
```

- `--dossier` : dossier des PDF (par défaut `PDF_Colis` à côté de `main.py`) ;
- `--sortie` : fichier Excel produit (par défaut `resultat.xlsx` à côté de `main.py`).

Code de sortie : `0` = succès (même si aucun PDF ne contient le numéro), `1` = erreur
(dossier introuvable, fichier Excel ouvert, option inconnue…).

## OCR (Tesseract)

Les pages scannées (images sans texte) sont lues par OCR avec Tesseract, via PyMuPDF.
Une page passe à l'OCR si elle a moins de 20 caractères, ou moins de 200 caractères avec
des images sur au moins 70 % de sa surface (scan d'appli téléphone avec un filigrane du type
« Scanned with CamScanner »). Les autres pages sont lues normalement.

- **Langues** : `fra.traineddata` et `eng.traineddata` doivent être **dans le même dossier
  tessdata**. Un dossier qui n'a que l'une des deux est ignoré. L'installeur UB Mannheim
  n'installe que `eng` par défaut : télécharger `fra.traineddata` depuis
  [tesseract-ocr/tessdata](https://github.com/tesseract-ocr/tessdata) et le copier à côté.
- **Ordre de recherche du dossier tessdata** (le premier qui contient les deux langues gagne) :
  1. la variable d'environnement `TESSDATA_PREFIX` (le dossier tessdata ou son dossier parent) ;
  2. `tesseract\tessdata` à côté de `main.py` (prévu pour la version empaquetée en .exe) ;
  3. les installations standard : `C:\Program Files\Tesseract-OCR\tessdata`,
     `C:\Program Files (x86)\Tesseract-OCR\tessdata`,
     `%LOCALAPPDATA%\Programs\Tesseract-OCR\tessdata`.
- **Sans Tesseract**, l'outil fonctionne : les PDF texte sont analysés normalement, un
  avertissement s'affiche et les pages scannées sont listées dans la feuille « Non analysés »
  avec la raison « OCR non disponible (Tesseract introuvable) ».

La colonne « Lecture » de la feuille « Résultats » indique comment les pages trouvées ont été
lues : « texte », « OCR » ou « texte + OCR ». Pour simuler une machine sans Tesseract :
`python tests\lancer_sans_tesseract.py COLIS-2026-00481`.

## Index

Pour ne pas relire (et repasser à l'OCR) tous les PDF à chaque recherche, le texte de chaque
page est gardé dans un index SQLite.

- **Où** : `%LOCALAPPDATA%\DDT_Recherche_PDF\index_<empreinte>.sqlite`, un fichier par
  dossier de PDF (l'empreinte est calculée à partir du chemin du dossier). L'index n'est
  jamais écrit dans le dossier des PDF, qui peut être sur OneDrive, sur un réseau ou en
  lecture seule. Deux recherches lancées en même temps ne se bloquent pas.
- **Mise à jour automatique, à chaque recherche** :
  - les PDF nouveaux, ou dont la taille ou la date de modification a changé, sont relus ;
  - les PDF supprimés sont retirés de l'index ;
  - les autres sont lus depuis l'index ;
  - une page scannée qui n'a pas pu passer à l'OCR (Tesseract absent, erreur d'OCR) est
    relue au lancement suivant ; un PDF illisible (mot de passe, fichier abîmé) n'est relu
    que s'il change.
- **Reconstruction complète** : automatique si les réglages de lecture changent (langues ou
  résolution de l'OCR, seuils de détection des scans, version de l'index) ou si le fichier
  d'index est abîmé. Pour la forcer : `python main.py COLIS-2026-00481 --reindexer`.
  On peut aussi simplement supprimer le fichier `.sqlite`.
- **Sans index** : `python main.py COLIS-2026-00481 --sans-index` lit directement tous les PDF.

Le terminal indique « Index : X PDF lus, Y depuis le cache, Z retirés » et affiche la
progression toutes les 10 PDF pendant une indexation longue.

**Conseil pour l'installation chez un client** : la première indexation passe tous les scans
à l'OCR (environ 0,4 s par page scannée). Lance-la une fois à l'installation, sans attendre
une première recherche :

```powershell
python main.py --indexer-seulement --dossier "D:\Dossier client\PDF"
```

## Construire la version livrable (.exe)

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

Produit `dist\Recherche_PDF\` (moteur PyInstaller en mode onedir, langues OCR, licences,
sources, `LISEZMOI.txt` pour le client, scripts de tâche planifiée). Le poste du client n'a
besoin ni de Python ni de Tesseract. Copier ensuite `Recherche_PDF.xlsm` à la racine de
`dist\Recherche_PDF\` : la macro utilise `moteur\Recherche_PDF.exe` s'il est présent, sinon
le `.venv` du projet.

## Créer le classeur Recherche_PDF.xlsm

Le bouton « Rechercher » lance `main.py` depuis Excel. Le classeur doit être enregistré
**à la racine du projet**, à côté de `main.py` et du dossier `.venv`.

1. Dans Excel, crée un nouveau classeur et enregistre-le sous `Recherche_PDF.xlsm`
   (type : *Classeur Excel (prenant en charge les macros)*).
2. Crée trois feuilles nommées exactement : `Recherche`, `Résultats`, `Non analysés`.
3. Mets en page la feuille `Recherche` :

   | Cellule | Contenu |
   |---|---|
   | A2 | Numéro à rechercher |
   | B2 | *(saisie, ex. COLIS-2026-00481)* |
   | A3 | Dossier des PDF (facultatif) |
   | B3 | *(vide = `PDF_Colis` à côté du classeur)* |

   Mets A2:A3 en gras, élargis les colonnes A (≈ 30) et B (≈ 60), et encadre B2:B3.
4. Sélectionne B2, clic droit > Format de cellule > **Texte**, **AVANT** de saisir un numéro.
   Sinon Excel supprime les zéros en tête : `00481` devient `481` (et passer la cellule en
   Texte après coup ne les rend pas : il faut ressaisir le numéro). La macro refuse un B2
   qui contient un nombre.
5. Si l'onglet **Développeur** n'est pas visible : Fichier > Options > Personnaliser le ruban >
   coche *Développeur*.
6. Importe la macro : `Alt+F11` > Fichier > Importer un fichier… > `Rechercher.bas`.
   Le module s'appelle `modRecherche`. Ferme l'éditeur VBA.
7. Ajoute le bouton : Développeur > Insérer > Bouton (contrôle de formulaire), dessine-le
   (par exemple en D2:E3), choisis la macro `Rechercher` puis OK.
   Renomme le bouton « Rechercher » (clic droit > Modifier le texte).
8. Enregistre le classeur.

Utilisation : saisis le numéro en B2, éventuellement un dossier en B3, puis clique sur
« Rechercher ». Les feuilles `Résultats` et `Non analysés` sont remplacées à chaque recherche.
En cas d'erreur, le message affiche le contenu de `journal.txt` (créé à côté du classeur).

## Tests

```powershell
python -m unittest discover tests
python tests\creer_pdf_test.py
python main.py COLIS-2026-00481
```
