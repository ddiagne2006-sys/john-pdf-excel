"""
Génère les PDF de test dans PDF_Colis pour vérifier main.py.

Prérequis :  pip install -r requirements-dev.txt   (reportlab, pillow)
Lancement :  python tests\\creer_pdf_test.py
"""

from pathlib import Path
import io
import random

import pymupdf
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent.parent
PDF_DIR = BASE_DIR / "PDF_Colis"

# Helvetica (police standard) n'a pas de glyphe pour U+202F : il faut une police TrueType intégrée.
POLICE_UNICODE = Path("C:/Windows/Fonts/arial.ttf")
ESPACES_INSECABLES = {"\u00a0": "U+00A0 espace insécable", "\u202f": "U+202F espace fine insécable"}

CONDITIONS_GENERALES = [
    "CONDITIONS GÉNÉRALES",
    "Article 3 : réclamation sous 48 heures après livraison.",
    "Article 7 : indemnité plafonnée à 150 000 GNF par colis.",
    "Article 12 : stockage gratuit 5 jours, puis 2 500 GNF par jour.",
]


def ecrire_page(pdf, lignes, police="Helvetica"):
    """Écrit une page de texte, une ligne sous l'autre."""
    pdf.setFont(police, 12)
    y = A4[1] - 72
    for ligne in lignes:
        pdf.drawString(72, y, ligne)
        y -= 20
    pdf.showPage()


def creer_pdf(chemin, pages):
    chemin.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(chemin), pagesize=A4)
    for lignes in pages:
        ecrire_page(pdf, lignes)
    pdf.save()
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def bordereau(reference):
    return [
        "BORDEREAU D'EXPÉDITION",
        f"Ref : {reference}",
        "Date : 15/10/2026",
        "Tél : +224622123456",
        "Poids : 25,5 kg",
        "Montant : 850 000 GNF",
        "Nombre de colis : 4",
    ]


def creer_bordereau_insecable(chemin):
    """850 000 écrit avec U+00A0 puis avec U+202F, et contrôle de ce que le PDF contient vraiment."""
    if not POLICE_UNICODE.exists():
        raise FileNotFoundError(f"Police introuvable : {POLICE_UNICODE}")
    pdfmetrics.registerFont(TTFont("ArialUnicode", str(POLICE_UNICODE)))
    pdf = canvas.Canvas(str(chemin), pagesize=A4)
    ecrire_page(pdf, [
        "BORDEREAU D'EXPÉDITION",
        "Ref : COLIS-2026-00481",
        "Montant : 850 000 GNF",
        "Montant TTC : 850 000 GNF",
    ], police="ArialUnicode")
    pdf.save()
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")

    with pymupdf.open(chemin) as lu:
        texte = lu[0].get_text()
    for caractere, nom in ESPACES_INSECABLES.items():
        etat = "présente" if caractere in texte else "ABSENTE (relue comme espace normale)"
        print(f"  {nom} dans le texte extrait : {etat}")


def creer_scan_sans_texte(chemin):
    """Une page avec seulement des formes : aucun texte extractible, comme un scan."""
    pdf = canvas.Canvas(str(chemin), pagesize=A4)
    pdf.setFillGray(0.85)
    pdf.rect(72, 400, 450, 300, fill=1, stroke=0)
    pdf.setStrokeGray(0.3)
    for y in range(420, 700, 20):
        pdf.line(90, y, 500, y)
    pdf.showPage()
    pdf.save()
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def pdf_texte_en_memoire(lignes):
    tampon = io.BytesIO()
    pdf = canvas.Canvas(tampon, pagesize=A4)
    ecrire_page(pdf, lignes)
    pdf.save()
    return tampon.getvalue()


def image_scan(lignes, dpi, abime=False):
    """
    Rend une page texte en image, comme un scanner.
    abime=True : niveaux de gris, rotation de 1,5°, points parasites (bruit toujours identique).
    """
    with pymupdf.open(stream=pdf_texte_en_memoire(lignes), filetype="pdf") as source:
        pix = source[0].get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY if abime else pymupdf.csRGB)
    image = Image.frombytes("L" if abime else "RGB", (pix.width, pix.height), pix.samples)
    if abime:
        image = image.rotate(1.5, resample=Image.Resampling.BICUBIC, fillcolor=255)
        hasard = random.Random(481)
        pixels = image.load()
        for _ in range(image.width * image.height // 200):   # 0,5 % des pixels
            x, y = hasard.randrange(image.width), hasard.randrange(image.height)
            pixels[x, y] = hasard.choice((0, 90, 170))
    return image


def ajouter_page_scannee(document, image):
    """Ajoute une page A4 qui ne contient que l'image : aucune couche texte."""
    tampon = io.BytesIO()
    image.save(tampon, format="PNG")
    page = document.new_page(width=A4[0], height=A4[1])
    page.insert_image(page.rect, stream=tampon.getvalue())


def creer_scan(chemin, lignes, dpi=300, abime=False):
    with pymupdf.open() as document:
        ajouter_page_scannee(document, image_scan(lignes, dpi, abime))
        document.save(str(chemin), garbage=3, deflate=True)
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def creer_mixte(chemin):
    """Page 1 en texte sans la référence, page 2 scannée avec la référence."""
    with pymupdf.open(stream=pdf_texte_en_memoire(CONDITIONS_GENERALES), filetype="pdf") as document:
        ajouter_page_scannee(document, image_scan(bordereau("COLIS-2026-00481"), 300))
        document.save(str(chemin), garbage=3, deflate=True)
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def creer_scan_filigrane(chemin):
    """Image du bordereau sur toute la page + couche texte en bas, comme les applis de scan sur téléphone."""
    with pymupdf.open() as document:
        ajouter_page_scannee(document, image_scan(bordereau("COLIS-2026-00481"), 300))
        document[0].insert_text((A4[0] / 2 - 55, A4[1] - 20), "Scanned with CamScanner",
                                fontsize=9, color=(0.5, 0.5, 0.5))
        document.save(str(chemin), garbage=3, deflate=True)
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def creer_bordereau_logo(chemin):
    """PDF texte normal avec un petit logo (bien moins de 70 % de la page) : doit rester lu en texte."""
    tampon = io.BytesIO()
    Image.new("RGB", (240, 120), (30, 90, 160)).save(tampon, format="PNG")
    with pymupdf.open(stream=pdf_texte_en_memoire(bordereau("COLIS-2026-00481")), filetype="pdf") as document:
        document[0].insert_image(pymupdf.Rect(A4[0] - 192, 20, A4[0] - 72, 80), stream=tampon.getvalue())
        document.save(str(chemin), garbage=3, deflate=True)
    print(f"Créé : {chemin.relative_to(BASE_DIR)}")


def main():
    # Doit être trouvé (page 1 seulement : la page 2 ne contient pas la référence)
    creer_pdf(PDF_DIR / "bordereau_00481.pdf", [bordereau("COLIS-2026-00481"), CONDITIONS_GENERALES])
    creer_bordereau_insecable(PDF_DIR / "bordereau_insecable.pdf")

    # Ne doivent PAS être trouvés
    creer_pdf(PDF_DIR / "bordereau_00482.pdf", [bordereau("COLIS-2026-00482")])
    creer_pdf(PDF_DIR / "piege_004810.pdf", [bordereau("COLIS-2026-004810")])

    # Doit être trouvé : sous-dossier, extension en majuscules, référence écrite autrement
    creer_pdf(PDF_DIR / "2026" / "Facture_F-118.PDF", [[
        "FACTURE F-118",
        "Objet : livraison du colis 2026 00481",
        "Total : 1 250 000,50 GNF",
    ]])

    # Doit apparaître dans la feuille "Non analysés"
    creer_scan_sans_texte(PDF_DIR / "scan_sans_texte.pdf")

    # Scans (images sans couche texte) : lus par OCR
    creer_scan(PDF_DIR / "scan_00481.pdf", bordereau("COLIS-2026-00481"), dpi=300)
    creer_scan(PDF_DIR / "scan_mauvais_00481.pdf", bordereau("COLIS-2026-00481"), dpi=150, abime=True)
    creer_mixte(PDF_DIR / "mixte_00481.pdf")
    creer_scan(PDF_DIR / "scan_00482.pdf", bordereau("COLIS-2026-00482"), dpi=300)   # ne doit PAS être trouvé

    # Scan d'appli téléphone avec filigrane texte : lecture « OCR »
    creer_scan_filigrane(PDF_DIR / "scan_filigrane_00481.pdf")
    # PDF texte avec un petit logo : lecture « texte »
    creer_bordereau_logo(PDF_DIR / "bordereau_logo_00481.pdf")


if __name__ == "__main__":
    main()
