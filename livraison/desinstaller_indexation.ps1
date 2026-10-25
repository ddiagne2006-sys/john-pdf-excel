<#
Retire la tâche planifiée créée par installer_indexation.ps1.
Avec -SupprimerIndex, supprime aussi les index déjà construits (%LOCALAPPDATA%\DDT_Recherche_PDF).

Utilisation :
  powershell -ExecutionPolicy Bypass -File C:\DDT\Recherche_PDF\source\desinstaller_indexation.ps1 [-SupprimerIndex]
#>
param(
    [string]$NomTache = 'DDT Recherche PDF - Indexation',
    [switch]$SupprimerIndex
)
$ErrorActionPreference = 'Stop'

if (Get-ScheduledTask -TaskName $NomTache -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $NomTache -Confirm:$false
    Write-Host "Tâche « $NomTache » supprimée."
}
else {
    Write-Host "Aucune tâche « $NomTache » : rien à supprimer."
}

if ($SupprimerIndex) {
    $index = Join-Path $env:LOCALAPPDATA 'DDT_Recherche_PDF'
    if (Test-Path $index) {
        Remove-Item -Recurse -Force $index
        Write-Host "Index supprimés : $index (ils seront reconstruits à la prochaine recherche)."
    }
}
