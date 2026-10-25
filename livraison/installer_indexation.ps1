<#
Crée une tâche planifiée Windows qui met à jour l'index des PDF chaque jour à 7 h et à
l'ouverture de session, sans fenêtre, pour l'utilisateur courant (sans droits administrateur).

Le script est livré dans Recherche_PDF\source\ ; le moteur est dans Recherche_PDF\moteur\.
Utilisation :
  powershell -ExecutionPolicy Bypass -File C:\DDT\Recherche_PDF\source\installer_indexation.ps1 -Dossier "D:\Archives\Colis 2026"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Dossier,
    [string]$NomTache = 'DDT Recherche PDF - Indexation'
)
$ErrorActionPreference = 'Stop'

# Version sans fenêtre du moteur (même programme que Recherche_PDF.exe)
$moteur = Join-Path (Split-Path $PSScriptRoot -Parent) 'moteur\Recherche_PDF_silencieux.exe'
if (-not (Test-Path $moteur)) { throw "Programme introuvable : $moteur" }
if (-not (Test-Path $Dossier -PathType Container)) { throw "Dossier des PDF introuvable : $Dossier" }
$Dossier = (Resolve-Path $Dossier).Path
# Un \ juste avant le guillemet fermant ("D:\") serait lu comme \" par le programme
if ($Dossier.EndsWith('\')) { $Dossier += '.' }

$utilisateur = "$env:USERDOMAIN\$env:USERNAME"
$action = New-ScheduledTaskAction -Execute $moteur `
    -Argument "--indexer-seulement --dossier `"$Dossier`"" `
    -WorkingDirectory (Split-Path $moteur)

$chaqueMatin = New-ScheduledTaskTrigger -Daily -At '07:00'
$aLaConnexion = New-ScheduledTaskTrigger -AtLogOn -User $utilisateur
$aLaConnexion.Delay = 'PT2M'   # laisse la session s'ouvrir avant de lire les PDF

$reglages = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4)
# Interactive + Limited : tourne dans la session de l'utilisateur, sans mot de passe ni élévation
$principal = New-ScheduledTaskPrincipal -UserId $utilisateur -LogonType Interactive -RunLevel Limited
$description = "Met à jour l'index de Recherche_PDF pour le dossier $Dossier (Diagne Digital & Tech)."

try {
    Register-ScheduledTask -TaskName $NomTache -Action $action -Trigger $chaqueMatin, $aLaConnexion `
        -Settings $reglages -Principal $principal -Description $description -Force | Out-Null
    Write-Host "Tâche « $NomTache » installée : chaque jour à 7 h et à l'ouverture de session."
}
catch [Microsoft.Management.Infrastructure.CimException] {
    # Certains postes refusent le déclencheur « ouverture de session » sans droits administrateur
    Write-Warning "Déclencheur « ouverture de session » refusé ($($_.Exception.Message.Trim()))."
    Register-ScheduledTask -TaskName $NomTache -Action $action -Trigger $chaqueMatin `
        -Settings $reglages -Principal $principal -Description $description -Force | Out-Null
    Write-Host "Tâche « $NomTache » installée : chaque jour à 7 h (ou dès que l'ordinateur est allumé)."
}
Write-Host "Dossier indexé : $Dossier"
Write-Host "Pour la lancer tout de suite : Start-ScheduledTask -TaskName '$NomTache'"
