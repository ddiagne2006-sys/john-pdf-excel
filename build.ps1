<#
Construit la version livrable dans dist\Recherche_PDF :
  moteur\ (exe PyInstaller + tesseract\tessdata), LISEZMOI.txt, LICENCES\,
  source\ (sources, scripts d'indexation, guide technicien), PDF_Colis\ (vide). Recherche_PDF.xlsm est à ajouter à la main.

Lancement (depuis le dossier du projet) :
  powershell -ExecutionPolicy Bypass -File build.ps1 [-Tessdata C:\outils\tessdata]
#>
param(
    # Dossier contenant fra.traineddata et eng.traineddata
    [string]$Tessdata = $(if ($env:TESSDATA_PREFIX) { $env:TESSDATA_PREFIX } else { 'C:\outils\tessdata' })
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# Fichiers officiels de github.com/tesseract-ocr : eng = tessdata_fast, fra = tessdata
$langues = [ordered]@{
    'eng.traineddata' = '7D4322BD2A7749724879683FC3912CB542F19906C83BCC1A52132556427170B2'
    'fra.traineddata' = 'EAC01C1D72540D6090FACB7B2F42DD0A2EE8FC57C5BE1B20548AE668E2761913'
}
$livraison = 'dist\Recherche_PDF'
$py = '.venv-build\Scripts\python.exe'

function Copier-Licence([string]$Motif, [string]$Destination) {
    $fichier = Get-ChildItem $Motif -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $fichier) { throw "Licence introuvable : $Motif" }
    Copy-Item $fichier.FullName (Join-Path "$livraison\LICENCES" $Destination)
}

Write-Host '1/6 Nettoyage de build\ et dist\'
foreach ($dossier in 'build', 'dist') {
    if (Test-Path $dossier) { Remove-Item -Recurse -Force $dossier }
}

Write-Host '2/6 Environnement de construction (.venv-build)'
if (-not (Test-Path $py)) {
    python -m venv .venv-build
    if ($LASTEXITCODE) { throw 'Création de .venv-build impossible' }
}
& $py -m pip install --disable-pip-version-check -q -r requirements.txt -r requirements-build.txt
if ($LASTEXITCODE) { throw 'Installation des dépendances impossible' }

Write-Host "3/6 Vérification des langues OCR dans $Tessdata"
foreach ($nom in $langues.Keys) {
    $fichier = Join-Path $Tessdata $nom
    if (-not (Test-Path $fichier)) { throw "Fichier de langue introuvable : $fichier" }
    if ((Get-FileHash $fichier -Algorithm SHA256).Hash -ne $langues[$nom]) {
        throw "Empreinte inattendue pour $fichier (ce n'est pas le fichier officiel attendu)"
    }
}

Write-Host '4/6 PyInstaller'
& $py -m PyInstaller --noconfirm --clean --log-level WARN `
    --distpath build\pyinstaller --workpath build\travail Recherche_PDF.spec
if ($LASTEXITCODE) { throw 'PyInstaller a échoué' }

Write-Host "5/6 Assemblage de $livraison"
New-Item -ItemType Directory -Force $livraison | Out-Null
Move-Item build\pyinstaller\moteur "$livraison\moteur"
foreach ($dossier in 'moteur\tesseract\tessdata', 'LICENCES', 'source', 'PDF_Colis') {
    New-Item -ItemType Directory -Force (Join-Path $livraison $dossier) | Out-Null
}
foreach ($nom in $langues.Keys) {
    Copy-Item (Join-Path $Tessdata $nom) "$livraison\moteur\tesseract\tessdata\"
}
Copy-Item livraison\LISEZMOI.txt $livraison     # document client
Copy-Item licences\* "$livraison\LICENCES\"      # index des licences, Tesseract, tessdata, Leptonica
$sitePackages = '.venv-build\Lib\site-packages'
Copier-Licence "$sitePackages\pymupdf-*.dist-info\COPYING" 'PyMuPDF-MuPDF-AGPL-3.0.txt'
Copier-Licence "$sitePackages\openpyxl-*.dist-info\LICENCE.rst" 'openpyxl-LICENSE.txt'
Copier-Licence "$sitePackages\et_xmlfile-*.dist-info\LICENCE.rst" 'et_xmlfile-LICENSE.txt'
Copier-Licence "$sitePackages\pyinstaller-*.dist-info\licenses\COPYING.txt" 'PyInstaller-COPYING.txt'
$prefixePython = & $py -c "import sys; print(sys.base_prefix)"
Copier-Licence "$prefixePython\LICENSE.txt" 'Python-LICENSE.txt'
# Sources (licence AGPL de PyMuPDF) : programme et instructions de construction
Copy-Item main.py, requirements.txt, requirements-build.txt, Recherche_PDF.spec, build.ps1, Rechercher.bas `
    "$livraison\source\"
# Pour le technicien : scripts de tâche planifiée et guide (pas à la racine, le client n'en a pas besoin)
Copy-Item livraison\installer_indexation.ps1, livraison\desinstaller_indexation.ps1, INSTALLATION_TECHNICIEN.txt `
    "$livraison\source\"

Write-Host '6/6 Taille'
function Taille-Mo([string]$Chemin) {
    [math]::Round((Get-ChildItem $Chemin -Recurse -File | Measure-Object Length -Sum).Sum / 1MB, 1)
}
Write-Host ("  moteur\_internal   : {0} Mo" -f (Taille-Mo "$livraison\moteur\_internal"))
Write-Host ("  moteur\tesseract   : {0} Mo" -f (Taille-Mo "$livraison\moteur\tesseract"))
Write-Host ("  Total livré        : {0} Mo ({1} fichiers)" -f (Taille-Mo $livraison),
    (Get-ChildItem $livraison -Recurse -File).Count)
Write-Host "Terminé : $((Resolve-Path $livraison).Path)"
Write-Host 'Reste à ajouter : Recherche_PDF.xlsm (classeur avec la macro, voir README.md).'
