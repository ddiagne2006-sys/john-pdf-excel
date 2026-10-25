"""
Tests unitaires de l'extraction des nombres et de la recherche de la référence.

Lancement :  python -m unittest discover tests
"""

from pathlib import Path
from unittest import mock
import os
import sys
import tempfile
import unittest

# main.py est à la racine du projet, un niveau au-dessus de tests/
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pymupdf  # noqa: E402

import main  # noqa: E402
from main import (  # noqa: E402
    couverture_images, extraire_nombres, motif_reference, page_sans_texte, trouver_tessdata,
)


class TestExtraireNombres(unittest.TestCase):
    def test_espace_insecable(self):
        self.assertEqual(extraire_nombres("Montant : 850 000 GNF"), ["850 000"])

    def test_espace_fine_insecable(self):
        self.assertEqual(extraire_nombres("Montant : 850 000 GNF"), ["850 000"])

    def test_decimale_virgule(self):
        self.assertEqual(extraire_nombres("Poids : 25,5 kg"), ["25,5"])


class TestUnitesCollees(unittest.TestCase):
    """Un nombre suivi d'une unité connue collée : seul le nombre est gardé."""

    def test_monnaie_collee(self):
        self.assertEqual(extraire_nombres("Montant:850000GNF"), ["850000"])

    def test_decimale_collee(self):
        self.assertEqual(extraire_nombres("Poids:25,5kg"), ["25,5"])

    def test_milliers_et_decimale_colles(self):
        self.assertEqual(extraire_nombres("Total 1 250 000,50FG"), ["1 250 000,50"])

    def test_milliers_insecables_colles(self):
        self.assertEqual(extraire_nombres("850 000GNF"), ["850 000"])

    def test_insensible_a_la_casse(self):
        self.assertEqual(extraire_nombres("25,5KG 12gnf 3Pcs"), ["25,5", "12", "3"])

    def test_toutes_les_unites(self):
        for unite in main.UNITES:
            with self.subTest(unite=unite):
                self.assertEqual(extraire_nombres(f"7{unite}"), ["7"])

    def test_references_inchangees(self):
        self.assertEqual(extraire_nombres("F-118 COLIS-2026-00481 15/10/2026"),
                         ["F-118", "COLIS-2026-00481", "15/10/2026"])

    def test_code_alphanumerique_non_coupe(self):
        self.assertEqual(extraire_nombres("A12B"), ["A12B"])

    def test_unite_inconnue_non_coupee(self):
        self.assertEqual(extraire_nombres("12km 5gr"), ["12km", "5gr"])

    def test_unite_separee_inchangee(self):
        self.assertEqual(extraire_nombres("850 000 GNF, 25,5 kg"), ["850 000", "25,5"])


class TestMotifReference(unittest.TestCase):
    def setUp(self):
        self.motif = motif_reference("COLIS-2026-00481")

    def test_trouve_forme_differente(self):
        self.assertIsNotNone(self.motif.search("livraison du colis 2026 00481"))

    def test_ignore_reference_plus_longue(self):
        self.assertIsNone(self.motif.search("Ref : COLIS-2026-004810"))

    def test_ignore_autre_reference(self):
        self.assertIsNone(self.motif.search("Ref : COLIS-2026-00482"))


class TestPageSansTexte(unittest.TestCase):
    """Seuil : moins de 20 caractères non blancs = page scannée."""

    def test_19_caracteres(self):
        self.assertTrue(page_sans_texte("  12345 67890\n123456789 \n"))

    def test_20_caracteres(self):
        self.assertFalse(page_sans_texte("  12345 67890\n1234567890 \n"))

    def test_page_vide(self):
        self.assertTrue(page_sans_texte(""))

    def test_blancs_seulement(self):
        self.assertTrue(page_sans_texte(" \n\t  " * 50))


class TestScanAvecFiligrane(unittest.TestCase):
    """Moins de 200 caractères ET images sur au moins 70 % de la page => OCR."""

    LARGEUR, HAUTEUR = 595, 842   # A4 en points

    def creer_page(self, images=(), lignes=()):
        """Page en mémoire avec des images (rectangles) et des lignes de texte."""
        document = pymupdf.open()
        self.addCleanup(document.close)
        page = document.new_page(width=self.LARGEUR, height=self.HAUTEUR)
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
        pixmap.clear_with(128)
        for rect in images:
            page.insert_image(pymupdf.Rect(rect), stream=pixmap.tobytes("png"))
        for i, ligne in enumerate(lignes):
            page.insert_text((50, 60 + 14 * i), ligne, fontsize=10)
        return page

    def test_filigrane_sur_image_pleine_page(self):
        page = self.creer_page([(0, 0, self.LARGEUR, self.HAUTEUR)], ["Scanned with CamScanner"])
        self.assertTrue(page_sans_texte(page.get_text(), page))

    def test_petit_logo_reste_texte(self):
        lignes = ["Ref : COLIS-2026-00481", "Montant : 850 000 GNF", "Poids : 25,5 kg"]
        page = self.creer_page([(400, 20, 520, 80)], lignes)
        self.assertLess(couverture_images(page), 0.70)
        self.assertFalse(page_sans_texte(page.get_text(), page))

    def test_image_pleine_page_avec_beaucoup_de_texte_reste_texte(self):
        lignes = [f"Ligne {i} : texte du document numérique" for i in range(10)]   # > 200 caractères
        page = self.creer_page([(0, 0, self.LARGEUR, self.HAUTEUR)], lignes)
        self.assertFalse(page_sans_texte(page.get_text(), page))

    def test_sans_page_seule_la_regle_des_20_caracteres(self):
        self.assertFalse(page_sans_texte("Scanned with CamScanner"))

    def test_couverture_union_sans_double_comptage(self):
        # Deux images de 50 % de la page qui se chevauchent sur 25 % : union = 75 %, pas 100 %
        page = self.creer_page([(0, 0, self.LARGEUR, self.HAUTEUR * 0.50),
                                (0, self.HAUTEUR * 0.25, self.LARGEUR, self.HAUTEUR * 0.75)])
        self.assertAlmostEqual(couverture_images(page), 0.75, places=2)

    def test_couverture_bornee_a_la_page(self):
        # Image qui déborde largement de la page : couverture plafonnée à 100 %
        page = self.creer_page([(-100, -100, self.LARGEUR + 100, self.HAUTEUR + 100)])
        self.assertAlmostEqual(couverture_images(page), 1.0, places=2)


class TestTrouverTessdata(unittest.TestCase):
    """trouver_tessdata() passe à l'étape suivante si un dossier est absent ou incomplet."""

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.racine = Path(temp.name)

        # Fausse installation standard qui contient fra et eng
        self.standard = self.racine / "Tesseract-OCR" / "tessdata"
        self.standard.mkdir(parents=True)
        for langue in ("fra", "eng"):
            (self.standard / f"{langue}.traineddata").touch()

        # Isole le test de la machine : pas de tesseract\tessdata « à côté de main.py »,
        # et une seule installation standard, la fausse
        for nom, valeur in (("BASE_DIR", self.racine / "projet"), ("CHEMINS_TESSDATA", [self.standard])):
            patch = mock.patch.object(main, nom, valeur)
            patch.start()
            self.addCleanup(patch.stop)

    def test_prefix_inexistant_passe_a_l_etape_suivante(self):
        with mock.patch.dict(os.environ, {"TESSDATA_PREFIX": str(self.racine / "inexistant")}):
            self.assertEqual(trouver_tessdata(), self.standard)

    def test_prefix_sans_francais_passe_a_l_etape_suivante(self):
        incomplet = self.racine / "incomplet"
        incomplet.mkdir()
        (incomplet / "eng.traineddata").touch()
        with mock.patch.dict(os.environ, {"TESSDATA_PREFIX": str(incomplet)}):
            self.assertEqual(trouver_tessdata(), self.standard)

    def test_prefix_valide_prioritaire(self):
        prefixe = self.racine / "prefixe"
        prefixe.mkdir()
        for langue in ("fra", "eng"):
            (prefixe / f"{langue}.traineddata").touch()
        with mock.patch.dict(os.environ, {"TESSDATA_PREFIX": str(prefixe)}):
            self.assertEqual(trouver_tessdata(), prefixe)

    def test_rien_trouve_renvoie_none(self):
        with mock.patch.object(main, "CHEMINS_TESSDATA", []), \
             mock.patch.dict(os.environ, {"TESSDATA_PREFIX": str(self.racine / "inexistant")}):
            self.assertIsNone(trouver_tessdata())


if __name__ == "__main__":
    unittest.main()
