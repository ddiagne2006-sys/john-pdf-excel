"""
Lance main.py comme si Tesseract n'était pas installé (mode dégradé), sans rien renommer
sur la machine : TESSDATA_PREFIX pointe vers un dossier vide, et le dossier « à côté de
main.py » et les installations standard sont masqués.

Lancement :  python tests\\lancer_sans_tesseract.py COLIS-2026-00481 [--sortie FICHIER.xlsx]
"""

from pathlib import Path
import os
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import main  # noqa: E402

if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as vide:
        os.environ["TESSDATA_PREFIX"] = vide
        main.BASE_DIR = Path(vide)       # masque tesseract\tessdata à côté de main.py
        main.CHEMINS_TESSDATA = []       # masque C:\Program Files\Tesseract-OCR…
        main.DOSSIER_INDEX = Path(vide)  # index vide : sinon les textes OCR déjà en cache seraient servis
        code = main.main()
    sys.exit(code)
