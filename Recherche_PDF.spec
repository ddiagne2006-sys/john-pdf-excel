# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller : construction du moteur (à lancer via build.ps1, pas directement).
#
# Mode onedir (moteur\Recherche_PDF.exe + moteur\_internal\) et non onefile : un onefile se
# décompresse dans %TEMP% à chaque lancement (lent) et il est plus souvent bloqué par les antivirus.
#
# Deux exécutables partagent le même dossier _internal :
#   Recherche_PDF.exe             console : macro Excel, ligne de commande, double-clic
#   Recherche_PDF_silencieux.exe  sans fenêtre : tâche planifiée d'indexation

a = Analysis(
    ["main.py"],
    excludes=["tkinter"],
)
pyz = PYZ(a.pure)

# -X utf8, comme « python -X utf8 » dans la macro : la sortie redirigée vers journal.txt
# est écrite en UTF-8, et la macro la relit en UTF-8.
options = [("X utf8", None, "OPTION")]

# upx=False : la compression UPX fait souvent réagir les antivirus
moteur = EXE(pyz, a.scripts, options, exclude_binaries=True, name="Recherche_PDF",
             console=True, upx=False)
silencieux = EXE(pyz, a.scripts, options, exclude_binaries=True, name="Recherche_PDF_silencieux",
                 console=False, upx=False)

COLLECT(moteur, silencieux, a.binaries, a.datas, name="moteur", upx=False)
