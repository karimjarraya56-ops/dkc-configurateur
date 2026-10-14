"""Export de la liste de matériel en Excel et en PDF (avec photos et codes)."""
import io
from datetime import date
from pathlib import Path

from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

ROUGE = "E30613"
COLONNES = ["photo", "code", "designation", "finition", "dimensions", "quantite", "unite", "page"]

# Polices Unicode disponibles selon le système (accents, «», ×, –)
POLICES = [
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
]


def _vignette(chemin, hauteur_px):
    """Image PNG redimensionnée (fond blanc) en mémoire, ou None."""
    if not chemin:
        return None
    im = Image.open(chemin).convert("RGB")
    ratio = hauteur_px / im.height
    im = im.resize((max(1, int(im.width * ratio)), hauteur_px), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf, im.width, im.height


def excel(lignes, t, titre, logo):
    """lignes : dicts (photo, code, designation, finition, dimensions, quantite, unite, page)."""
    wb = Workbook()
    ws = wb.active
    ws.title = titre[:31]
    ws.sheet_view.showGridLines = False
    largeurs = [12, 12, 52, 30, 26, 10, 8, 10]
    for i, l in enumerate(largeurs):
        ws.column_dimensions[chr(65 + i)].width = l

    ws.row_dimensions[1].height = 40
    if logo and Path(logo).exists():
        v = _vignette(logo, 40)
        img = XLImage(v[0])
        ws.add_image(img, "A1")
    ws["C1"] = titre
    ws["C1"].font = Font(size=16, bold=True, color="1C1F24")
    ws["C1"].alignment = Alignment(vertical="center")
    ws["C2"] = f"{t('date')} : {date.today().strftime('%d/%m/%Y')}"
    ws["C2"].font = Font(size=10, color="6B7280")

    entete = 4
    fin_ligne = Side(style="thin", color="E6E8EC")
    for i, col in enumerate(COLONNES):
        c = ws.cell(row=entete, column=i + 1, value=t(f"col_{col}"))
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=ROUGE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[entete].height = 22

    for n, lg in enumerate(lignes):
        r = entete + 1 + n
        ws.row_dimensions[r].height = 48
        valeurs = ["", lg["code"], lg["designation"], lg["finition"], lg["dimensions"], lg["quantite"],
                   lg["unite"], lg["page"]]
        for i, v in enumerate(valeurs):
            c = ws.cell(row=r, column=i + 1, value=v)
            c.alignment = Alignment(vertical="center", wrap_text=True,
                                    horizontal="center" if i in (1, 5, 6, 7) else "left")
            c.border = Border(bottom=fin_ligne)
        ws.cell(row=r, column=2).font = Font(bold=True, color=ROUGE)
        ws.cell(row=r, column=6).font = Font(bold=True)
        v = _vignette(lg.get("photo"), 56)
        if v:
            img = XLImage(v[0])
            img.width, img.height = min(v[1], 80), 56 * min(v[1], 80) / v[1]
            ws.add_image(img, f"A{r}")

    total = entete + len(lignes) + 2
    ws.cell(row=total, column=3, value=f"{t('total_articles')} : {len(lignes)}").font = Font(bold=True)
    ws.cell(row=total + 1, column=3, value=t("mention_doc")).font = Font(size=9, italic=True, color="6B7280")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class _PDF(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font(self.police, size=8)
        self.set_text_color(107, 114, 128)
        self.cell(0, 6, self.mention, align="L")
        self.cell(0, 6, f"{self.page_no()}/{{nb}}", align="R")


def pdf(lignes, t, titre, logo):
    doc = _PDF(orientation="P", unit="mm", format="A4")
    doc.mention = t("mention_doc")
    police = None
    for normal, gras in POLICES:
        if Path(normal).exists() and Path(gras).exists():
            doc.add_font("Txt", "", normal)
            doc.add_font("Txt", "B", gras)
            police = "Txt"
            break
    doc.police = police or "Helvetica"
    if police is None:  # police de secours sans Unicode complet
        for lg in lignes:
            for k in ("designation", "finition", "dimensions"):
                lg[k] = lg[k].replace("–", "-").replace("«", '"').replace("»", '"')
    doc.alias_nb_pages()
    doc.set_auto_page_break(auto=True, margin=16)
    doc.add_page()

    if logo and Path(logo).exists():
        doc.image(str(logo), x=12, y=10, h=10)
    doc.set_xy(12, 24)
    doc.set_font(doc.police, "B", 16)
    doc.set_text_color(28, 31, 36)
    doc.cell(0, 8, titre, new_x="LMARGIN", new_y="NEXT")
    doc.set_font(doc.police, size=9)
    doc.set_text_color(107, 114, 128)
    doc.cell(0, 5, f"{t('date')} : {date.today().strftime('%d/%m/%Y')}   ·   {t('total_articles')} : {len(lignes)}",
             new_x="LMARGIN", new_y="NEXT")
    doc.ln(4)

    # Colonnes : photo, code, désignation (+ finition, dimensions), quantité, page
    larg = {"photo": 26, "code": 26, "designation": 98, "quantite": 22, "page": 14}
    hauteur = 22

    def entete():
        doc.set_fill_color(227, 6, 19)
        doc.set_text_color(255, 255, 255)
        doc.set_font(doc.police, "B", 9)
        for col in larg:
            doc.cell(larg[col], 8, t(f"col_{col}"), fill=True, align="C")
        doc.ln()

    entete()
    for lg in lignes:
        if doc.get_y() + hauteur > doc.h - 18:
            doc.add_page()
            entete()
        y = doc.get_y()
        x = doc.l_margin
        v = _vignette(lg.get("photo"), 120)
        if v:
            w = min(larg["photo"] - 4, (hauteur - 4) * v[1] / v[2])
            h = w * v[2] / v[1]
            doc.image(v[0], x=x + (larg["photo"] - w) / 2, y=y + (hauteur - h) / 2, w=w, h=h)
        x += larg["photo"]
        doc.set_xy(x, y)
        doc.set_font(doc.police, "B", 10)
        doc.set_text_color(227, 6, 19)
        doc.cell(larg["code"], hauteur, lg["code"], align="C")
        x += larg["code"]
        doc.set_xy(x + 1, y + 3)
        doc.set_text_color(28, 31, 36)
        doc.set_font(doc.police, "B", 9)
        doc.multi_cell(larg["designation"] - 2, 4.5, lg["designation"], new_x="LEFT", new_y="NEXT", max_line_height=4.5)
        doc.set_font(doc.police, size=8)
        doc.set_text_color(107, 114, 128)
        doc.multi_cell(larg["designation"] - 2, 4, f"{lg['finition']}  ·  {lg['dimensions']}", new_x="LEFT", new_y="NEXT")
        x += larg["designation"]
        doc.set_xy(x, y)
        doc.set_font(doc.police, "B", 11)
        doc.set_text_color(28, 31, 36)
        doc.cell(larg["quantite"], hauteur, f"{lg['quantite']} {lg['unite']}", align="C")
        x += larg["quantite"]
        doc.set_font(doc.police, size=9)
        doc.set_text_color(107, 114, 128)
        doc.cell(larg["page"], hauteur, str(lg["page"]), align="C")
        doc.set_draw_color(230, 232, 236)
        doc.line(doc.l_margin, y + hauteur, doc.l_margin + sum(larg.values()), y + hauteur)
        doc.set_xy(doc.l_margin, y + hauteur)
    return bytes(doc.output())
