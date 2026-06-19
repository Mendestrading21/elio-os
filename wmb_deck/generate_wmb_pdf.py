# -*- coding: utf-8 -*-
"""
Générateur du deck premium WallStreet Market Brief (WMB) — version PDF.

Rendu vectoriel autonome via reportlab (aucune dépendance LibreOffice).
Format 16:9 (13,333 x 7,5 in / 960 x 540 pt), ~40 pages.

Esthétique : rapport de banque privée — sobre, blanc dominant, grille
rigoureuse, accents parcimonieux, data-viz lisibles dessinées à la main.

Polices embarquées : un serif élégant (Liberation Serif, proxy de Playfair
Display, absent de l'environnement) pour les titres ; une sans-serif propre
(Liberation Sans, métriquement identique à Arial) pour le corps.
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth, registerFontFamily
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
import math

# ---------------------------------------------------------------------------
#  Polices
# ---------------------------------------------------------------------------
LIB = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("Serif",      LIB + "LiberationSerif-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Serif-Bold", LIB + "LiberationSerif-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Serif-It",   LIB + "LiberationSerif-Italic.ttf"))
pdfmetrics.registerFont(TTFont("Serif-BdIt", LIB + "LiberationSerif-BoldItalic.ttf"))
pdfmetrics.registerFont(TTFont("Sans",       LIB + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Sans-Bold",  LIB + "LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Sans-It",    LIB + "LiberationSans-Italic.ttf"))
pdfmetrics.registerFont(TTFont("Sans-BdIt",  LIB + "LiberationSans-BoldItalic.ttf"))
registerFontFamily("Serif", normal="Serif", bold="Serif-Bold", italic="Serif-It", boldItalic="Serif-BdIt")
registerFontFamily("Sans", normal="Sans", bold="Sans-Bold", italic="Sans-It", boldItalic="Sans-BdIt")

SERIF, SERIF_B, SERIF_I = "Serif", "Serif-Bold", "Serif-It"
SANS, SANS_B, SANS_I = "Sans", "Sans-Bold", "Sans-It"

# ---------------------------------------------------------------------------
#  THEME
# ---------------------------------------------------------------------------
NAVY   = HexColor("#0F2742")
STEEL  = HexColor("#1E3A5F")
SLATE  = HexColor("#5B7A9D")
GOLD   = HexColor("#C9A34E")
RED    = HexColor("#B23A3A")
GREEN  = HexColor("#2E7D5B")
BG     = HexColor("#F5F6F8")
RULE   = HexColor("#E2E5EA")
INK    = HexColor("#1A1A1A")
OFFW   = HexColor("#F7F7F4")
WHITE  = HexColor("#FFFFFF")
NAVY_DK = HexColor("#1B3A5C")   # numéro de section sur fond navy
MUTE_LT = HexColor("#9DB2C9")   # texte clair sur navy
MUTE_LT2 = HexColor("#C7D2DE")

IN = 72.0
W_IN, H_IN = 13.333, 7.5
PW, PH = W_IN * IN, H_IN * IN
M = 0.62                      # marge (in)
CW = W_IN - 2 * M             # largeur de contenu (in)

# ---------------------------------------------------------------------------
#  Primitives (coordonnées en pouces, origine en HAUT à gauche)
# ---------------------------------------------------------------------------

def _y(top_in, h_in=0.0):
    """Convertit un top (origine haut) vers la coordonnée reportlab (bas)."""
    return PH - (top_in + h_in) * IN


def rect(c, x, y, w, h, fill=None, stroke=None, lw=0.75, radius=None):
    c.saveState()
    if stroke is not None:
        c.setStrokeColor(stroke); c.setLineWidth(lw)
    do_fill = 1 if fill is not None else 0
    do_stroke = 1 if stroke is not None else 0
    if fill is not None:
        c.setFillColor(fill)
    if radius:
        c.roundRect(x * IN, _y(y, h), w * IN, h * IN, radius * IN, stroke=do_stroke, fill=do_fill)
    else:
        c.rect(x * IN, _y(y, h), w * IN, h * IN, stroke=do_stroke, fill=do_fill)
    c.restoreState()


def hline(c, x, y, w, color, lw=0.75):
    c.saveState()
    c.setStrokeColor(color); c.setLineWidth(lw)
    yy = _y(y)
    c.line(x * IN, yy, (x + w) * IN, yy)
    c.restoreState()


def vline(c, x, y, h, color, lw=0.75):
    c.saveState()
    c.setStrokeColor(color); c.setLineWidth(lw)
    c.line(x * IN, _y(y), x * IN, _y(y + h))
    c.restoreState()


def line_text(c, x, y_top, h, s, font, size, color, align="l", vmid=True):
    """Texte sur une ligne. y_top/h définissent la boîte ; centré verticalement."""
    c.setFont(font, size); c.setFillColor(color)
    if vmid:
        baseline = _y(y_top + h / 2) - size * 0.34
    else:
        baseline = _y(y_top) - size * 0.92
    xpt = x * IN
    if align == "l":
        c.drawString(xpt, baseline, s)
    elif align == "r":
        c.drawRightString((x + 0) * IN, baseline, s)
    elif align == "c":
        c.drawCentredString(xpt, baseline, s)


def line_text_box(c, x, y_top, w, h, s, font, size, color, align="l", vmid=True):
    """Comme line_text mais align right/center relatif à la boîte [x, x+w]."""
    c.setFont(font, size); c.setFillColor(color)
    baseline = _y(y_top + h / 2) - size * 0.34 if vmid else _y(y_top) - size * 0.92
    if align == "l":
        c.drawString(x * IN, baseline, s)
    elif align == "r":
        c.drawRightString((x + w) * IN, baseline, s)
    elif align == "c":
        c.drawCentredString((x + w / 2) * IN, baseline, s)


def para(c, x, y_top, w, h, html, font=SANS, size=12, color=INK, leading=None,
         align=TA_LEFT, space_after=4, anchor="top"):
    """Paragraphe(s) avec retour à la ligne (platypus). `html` : str ou liste de str."""
    if leading is None:
        leading = size * 1.2
    style = ParagraphStyle("s", fontName=font, fontSize=size, leading=leading,
                           textColor=color, alignment=align, spaceAfter=space_after)
    if isinstance(html, str):
        html = [html]
    flow = [Paragraph(t, style) for t in html]
    # hauteur totale réelle
    total = 0
    sizes = []
    for f in flow:
        fw, fh = f.wrap(w * IN, h * IN)
        sizes.append(fh)
        total += fh
    if anchor == "middle":
        y0 = y_top + (h - total / IN) / 2
    else:
        y0 = y_top
    cur = y0
    for f, fh in zip(flow, sizes):
        f.drawOn(c, x * IN, _y(cur, fh / IN))
        cur += fh / IN
    return total / IN


def soft_shadow_rect(c, x, y, w, h, fill, radius=0.045):
    """Carte claire avec ombre très légère (décalage discret)."""
    c.saveState()
    c.setFillColor(HexColor("#0F2742"))
    c.setFillAlpha(0.07)
    c.roundRect((x + 0.035) * IN, _y(y + 0.05, h) - 0.0, w * IN, h * IN, radius * IN, stroke=0, fill=1)
    c.setFillAlpha(1)
    c.restoreState()
    rect(c, x, y, w, h, fill=fill, stroke=RULE, lw=0.75, radius=radius)


# ---------------------------------------------------------------------------
#  Helpers de contenu (format FR, couleurs)
# ---------------------------------------------------------------------------

def fmt_fr(value, decimals=2, signed=False, suffix=""):
    neg = value < 0
    s = f"{abs(value):,.{decimals}f}".replace(",", " ").replace(".", ",")
    if signed:
        s = ("-" if neg else "+") + s
    elif neg:
        s = "-" + s
    return s + suffix


def color_for_change(v):
    return GREEN if v >= 0 else RED


PAGE = {"n": 0}


def page_bg(c, color):
    rect(c, -0.1, -0.1, W_IN + 0.2, H_IN + 0.2, fill=color)


def footer(c, dark=False):
    PAGE["n"] += 1
    tc = SLATE if not dark else MUTE_LT
    rc = RULE if not dark else STEEL
    y = H_IN - 0.46
    hline(c, M, y, CW, rc, 0.5)
    line_text_box(c, M, y + 0.04, 4, 0.3, "WallStreet Market Brief", SANS, 8.5, tc, "l", vmid=False)
    line_text_box(c, W_IN - M - 2, y + 0.04, 2, 0.3, f"{PAGE['n']:02d}", SANS, 8.5, tc, "r", vmid=False)


def title_band(c, title, section="", subtitle=None):
    x, y = M, 0.5
    rect(c, x, y + 0.06, 0.14, 0.42, fill=GOLD)
    line_text(c, x + 0.30, y - 0.02, 0.62, title, SERIF_B, 31, NAVY, "l")
    if section:
        line_text_box(c, W_IN - M - 3.6, y + 0.10, 3.6, 0.35, section.upper(), SANS_B, 10, SLATE, "r")
    by = y + 0.62
    hline(c, x, by, CW, NAVY, 1.1)
    if subtitle:
        para(c, x, by + 0.10, CW, 0.4, _it(subtitle), SANS_I, 13.5, STEEL)
        return by + 0.62
    return by + 0.22


def _it(s):
    return f'<i>{s}</i>'


def _b(s):
    return f'<b>{s}</b>'


# ---------------------------------------------------------------------------
#  Composants
# ---------------------------------------------------------------------------

def kpi_block(c, x, y, w, number, label, context="", color=GOLD, num_size=50):
    # nombre (peut contenir un \n)
    parts = number.split("\n")
    line_h = num_size / IN * 1.02
    cy = y
    for p in parts:
        line_text_box(c, x, cy, w, line_h, p, SERIF_B, num_size, color, "l")
        cy += line_h
    yl = (y + 0.95) if len(parts) == 1 else (cy + 0.05)
    line_text_box(c, x, yl, w, 0.30, label.upper(), SANS_B, 11, NAVY, "l", vmid=False)
    if context:
        para(c, x, yl + 0.34, w, 0.9, context, SANS, 10.5, SLATE, leading=12.5)


def card(c, x, y, w, h, kicker=None, title=None, body=None, accent=None,
         fill=WHITE, title_color=NAVY):
    soft_shadow_rect(c, x, y, w, h, fill)
    pad = 0.22
    cy = y + pad
    if accent is not None:
        rect(c, x + pad, cy + 0.02, 0.30, 0.30, fill=accent)
        tx = x + pad + 0.44
        tw = w - 2 * pad - 0.44
    else:
        tx, tw = x + pad, w - 2 * pad
    if kicker:
        line_text_box(c, tx, cy, tw, 0.26, kicker.upper(), SANS_B, 9, SLATE, "l", vmid=False)
        cy += 0.26
    if title:
        line_text_box(c, tx, cy, tw, 0.36, title, SERIF_B, 15, title_color, "l", vmid=False)
        cy += 0.46
    if body:
        if isinstance(body, list):
            body = " ".join([b for b in body if b])
        para(c, x + pad, cy, w - 2 * pad, h - (cy - y) - pad, body, SANS, 10.5, INK, leading=14.5)


# ---------------------------------------------------------------------------
#  Data-viz dessinées à la main
# ---------------------------------------------------------------------------

def bar_chart(c, x, y, w, h, categories, values, signed=False, highlight_last=False):
    """Histogramme. x,y,w,h = zone de tracé (in)."""
    n = len(values)
    vmax = max(values); vmin = min(values)
    if signed:
        lo = min(0, vmin); hi = max(0, vmax)
    else:
        lo = 0; hi = vmax
    span = (hi - lo) or 1
    pad_top = 0.45
    plot_top = y + pad_top
    plot_h = h - pad_top - 0.5         # 0.5 pour labels de catégorie
    plot_bottom = plot_top + plot_h
    zero_y = plot_top + plot_h * (hi - 0) / span if signed else plot_bottom
    # ligne de base
    hline(c, x, zero_y, w, RULE, 0.8)
    slot = w / n
    bw = slot * 0.46
    for i, v in enumerate(values):
        cx = x + slot * (i + 0.5)
        bx = cx - bw / 2
        bh = plot_h * (abs(v - 0)) / span
        if signed:
            col = color_for_change(v)
            by = zero_y - bh if v >= 0 else zero_y
        else:
            col = NAVY if not (highlight_last and i == n - 1) else GOLD
            by = plot_bottom - bh
        rect(c, bx, by, bw, bh, fill=col)
        # étiquette de valeur
        if signed:
            lbl = fmt_fr(v, 1, signed=True)
            ly = (by - 0.24) if v >= 0 else (by + bh + 0.02)
        else:
            lbl = fmt_fr(v, 0)
            ly = by - 0.24
        line_text_box(c, cx - slot / 2, ly, slot, 0.22, lbl, SANS_B, 10, NAVY, "c", vmid=False)
        # catégorie
        for k, cl in enumerate(categories[i].split("\n")):
            line_text_box(c, cx - slot / 2, plot_bottom + 0.10 + k * 0.16, slot, 0.16,
                          cl, SANS, 9.5, SLATE, "c", vmid=False)


def line_chart(c, x, y, w, h, categories, series, colors, legend=True):
    """Graphique en ligne. series = dict {nom: [valeurs]}."""
    allv = [v for s in series.values() for v in s]
    hi = max(allv); lo = min(allv)
    span = (hi - lo) or 1
    pad_top = 0.3
    leg_h = 0.4 if legend else 0.0
    plot_top = y + pad_top
    plot_h = h - pad_top - 0.45 - leg_h
    plot_bottom = plot_top + plot_h
    n = len(categories)
    # grille horizontale légère (4 lignes)
    for g in range(5):
        gy = plot_top + plot_h * g / 4
        hline(c, x, gy, w, RULE, 0.4)
    def px(i): return x + w * i / (n - 1)
    def py(v): return plot_top + plot_h * (hi - v) / span
    names = list(series.keys())
    for si, name in enumerate(names):
        col = colors[si % len(colors)]
        vals = series[name]
        c.saveState(); c.setStrokeColor(col); c.setLineWidth(2.2)
        pts = [(px(i) * IN, _y(py(v))) for i, v in enumerate(vals)]
        for i in range(len(pts) - 1):
            c.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1])
        c.restoreState()
        for (ptx, pty) in pts:
            c.setFillColor(col); c.circle(ptx, pty, 3.1, stroke=0, fill=1)
            c.setFillColor(WHITE); c.circle(ptx, pty, 1.2, stroke=0, fill=1)
    # catégories
    for i, cat in enumerate(categories):
        line_text_box(c, px(i) - 0.6, plot_bottom + 0.10, 1.2, 0.18, cat, SANS, 9.5, SLATE, "c", vmid=False)
    # légende
    if legend:
        lx = x
        ly = plot_bottom + 0.42
        for si, name in enumerate(names):
            col = colors[si % len(colors)]
            rect(c, lx, ly, 0.22, 0.10, fill=col)
            tw = stringWidth(name, SANS, 10) / IN
            line_text_box(c, lx + 0.30, ly - 0.04, tw + 0.2, 0.18, name, SANS, 10, SLATE, "l", vmid=False)
            lx += 0.30 + tw + 0.5


def donut(c, cx, cy, r_out, r_in, pct, accent):
    """Jauge donut : segment 'accent' = pct, reste gris ; trou au centre."""
    cxp, cyp = cx * IN, _y(cy)
    rop = r_out * IN
    # reste (cercle plein gris)
    c.setFillColor(RULE); c.circle(cxp, cyp, rop, stroke=0, fill=1)
    # segment valeur (wedge depuis le haut, sens horaire)
    extent = -360.0 * pct / 100.0
    c.setFillColor(accent)
    c.wedge(cxp - rop, cyp - rop, cxp + rop, cyp + rop, 90, extent, stroke=0, fill=1)
    # trou
    c.setFillColor(WHITE); c.circle(cxp, cyp, r_in * IN, stroke=0, fill=1)


# ---------------------------------------------------------------------------
#  SLIDES
# ---------------------------------------------------------------------------

def s_cover(c, title, subtitle, kicker, date_str):
    page_bg(c, NAVY)
    rect(c, M, 1.5, 0.6, 0.10, fill=GOLD)
    line_text_box(c, M, 1.78, CW, 0.35, kicker.upper(), SANS_B, 13, MUTE_LT, "l", vmid=False)
    line_text_box(c, M, 2.45, CW, 0.9, title, SERIF_B, 54, OFFW, "l")
    para(c, M, 4.35, CW - 2.0, 1.0, _it(subtitle), SANS_I, 18, MUTE_LT2, leading=24)
    hline(c, M, H_IN - 1.15, CW, STEEL, 1.0)
    line_text_box(c, M, H_IN - 1.0, 8, 0.35, "wallstreetmarketbrief.ch", SANS_B, 12, GOLD, "l", vmid=False)
    line_text_box(c, W_IN - M - 4, H_IN - 1.0, 4, 0.35, date_str, SANS, 12, MUTE_LT, "r", vmid=False)
    c.showPage()


def s_contents(c, entries):
    page_bg(c, WHITE)
    top = title_band(c, "Sommaire", "Plan du document")
    n = len(entries); half = (n + 1) // 2
    col_w = (CW - 0.6) / 2
    for i, (num, label, desc) in enumerate(entries):
        col = 0 if i < half else 1
        row = i if i < half else i - half
        x = M + col * (col_w + 0.6)
        y = top + 0.30 + row * 0.95
        line_text_box(c, x, y, 0.8, 0.6, num, SERIF_B, 26, GOLD, "l")
        line_text_box(c, x + 0.85, y, col_w - 0.85, 0.38, label, SERIF_B, 15, NAVY, "l", vmid=False)
        para(c, x + 0.85, y + 0.40, col_w - 0.85, 0.4, desc, SANS, 10, SLATE)
        hline(c, x + 0.85, y + 0.80, col_w - 0.85, RULE, 0.5)
    footer(c)
    c.showPage()


def s_divider(c, num, title, subtitle):
    page_bg(c, NAVY)
    line_text_box(c, M - 0.05, 2.2, 3, 1.8, num, SERIF_B, 120, NAVY_DK, "l")
    rect(c, M, 4.35, 0.6, 0.09, fill=GOLD)
    line_text_box(c, M, 4.6, CW, 0.7, title, SERIF_B, 40, OFFW, "l", vmid=False)
    para(c, M, 5.55, CW - 1.5, 0.9, _it(subtitle), SANS_I, 15, MUTE_LT, leading=22)
    c.showPage()


def s_kpi_row(c, title, section, kpis, subtitle=None, footnote=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    n = len(kpis); gap = 0.5
    cw = (CW - gap * (n - 1)) / n
    y = top + 0.75
    for i, k in enumerate(kpis):
        x = M + i * (cw + gap)
        kpi_block(c, x, y, cw, k["number"], k["label"], k.get("context", ""),
                  color=k.get("color", GOLD), num_size=k.get("size", 50))
    for i in range(1, n):
        x = M + i * (cw + gap) - gap / 2
        vline(c, x, y + 0.1, 1.7, RULE, 0.75)
    if footnote:
        para(c, M, H_IN - 0.95, CW, 0.4, footnote, SANS, 10, SLATE, leading=12.5)
    footer(c)
    c.showPage()


def s_editorial(c, title, section, paragraphs, pull_quote=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section)
    col_w = CW * 0.60
    para(c, M, top + 0.20, col_w, H_IN - top - 0.9, paragraphs, SANS, 12, INK,
         leading=17.5, space_after=10)
    if pull_quote:
        qx = M + col_w + 0.5
        qw = CW - col_w - 0.5
        rect(c, qx, top + 0.30, 0.10, 1.9, fill=GOLD)
        para(c, qx + 0.28, top + 0.25, qw - 0.28, 2.6, _it(pull_quote), SERIF_I, 19, NAVY, leading=25)
    footer(c)
    c.showPage()


def s_cards(c, title, section, cards, subtitle=None, cols=3):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    gap = 0.4
    cw = (CW - gap * (cols - 1)) / cols
    ch = H_IN - top - 1.05
    for i, cd in enumerate(cards):
        x = M + i * (cw + gap)
        card(c, x, top + 0.15, cw, ch, kicker=cd.get("kicker"), title=cd.get("title"),
             body=cd.get("body"), accent=cd.get("accent"))
    footer(c)
    c.showPage()


def s_scoreboard(c, title, section, indices):
    page_bg(c, WHITE)
    top = title_band(c, title, section,
                     "Clôture de la veille — données illustratives, vérifiées sur deux sources")
    cols = 4
    rows = (len(indices) + cols - 1) // cols
    gap = 0.28
    cw = (CW - gap * (cols - 1)) / cols
    avail = H_IN - top - 0.95
    chh = (avail - gap * (rows - 1)) / rows
    for i, (name, value, change) in enumerate(indices):
        r, cc = divmod(i, cols)
        x = M + cc * (cw + gap)
        y = top + 0.15 + r * (chh + gap)
        soft_shadow_rect(c, x, y, cw, chh, WHITE, radius=0.06)
        col = color_for_change(change)
        rect(c, x, y + 0.14, 0.07, chh - 0.28, fill=col)
        line_text_box(c, x + 0.22, y + 0.13, cw - 0.4, 0.26, name, SANS_B, 10.5, SLATE, "l", vmid=False)
        line_text_box(c, x + 0.22, y + 0.40, cw - 0.4, 0.40, value, SERIF_B, 21, NAVY, "l", vmid=False)
        line_text_box(c, x + 0.22, y + chh - 0.40, cw - 0.4, 0.28,
                      fmt_fr(change, 2, signed=True, suffix=" %"), SANS_B, 12, col, "l", vmid=False)
    footer(c)
    c.showPage()


def s_breaking(c, tag, headline, lede):
    page_bg(c, NAVY)
    rect(c, M, 1.7, 2.4, 0.55, fill=RED)
    line_text_box(c, M, 1.7, 2.4, 0.55, tag.upper(), SANS_B, 17, WHITE, "c")
    para(c, M, 2.55, CW - 1.0, 2.0, headline, SERIF_B, 36, OFFW, leading=42)
    rect(c, M, 4.95, 0.10, 1.0, fill=RED)
    para(c, M + 0.28, 4.95, CW - 1.0, 1.1, _it(lede), SANS_I, 15, MUTE_LT2, leading=21)
    footer(c, dark=True)
    c.showPage()


def s_process(c, title, section, steps, subtitle=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    n = len(steps); gap = 0.28
    cw = (CW - gap * (n - 1)) / n
    y = top + 0.4
    ch = H_IN - top - 1.3
    for i, (head, body) in enumerate(steps):
        x = M + i * (cw + gap)
        rect(c, x, y, cw, ch, fill=BG, stroke=RULE, lw=0.75, radius=0.05)
        line_text_box(c, x + 0.18, y + 0.18, cw - 0.36, 0.5, f"{i+1:02d}", SERIF_B, 26, GOLD, "l", vmid=False)
        para(c, x + 0.18, y + 0.78, cw - 0.36, 0.7, head, SERIF_B, 13.5, NAVY, leading=16)
        para(c, x + 0.18, y + 1.46, cw - 0.36, ch - 1.6, body, SANS, 10.5, INK, leading=14)
        if i < n - 1:
            line_text_box(c, x + cw + 0.01, y + ch / 2 - 0.2, gap - 0.02, 0.4, "›", SANS_B, 18, SLATE, "c")
    footer(c)
    c.showPage()


def s_table(c, title, section, headers, rows, subtitle=None, col_w=None,
            right_cols=None, change_cols=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    right_cols = right_cols or []
    change_cols = change_cols or []
    ncols = len(headers)
    if col_w is None:
        col_w = [1.0 / ncols] * ncols
    tot = sum(col_w)
    widths = [CW * cwd / tot for cwd in col_w]
    xs = [M]
    for wdt in widths[:-1]:
        xs.append(xs[-1] + wdt)
    nrows = len(rows)
    head_h = 0.5
    avail = H_IN - top - 1.0
    row_h = min(0.56, (avail - head_h) / nrows)
    y = top + 0.18
    # en-tête
    rect(c, M, y, CW, head_h, fill=NAVY)
    for j, htext in enumerate(headers):
        al = "r" if j in right_cols else "l"
        xx = xs[j] + (widths[j] - 0.16 if al == "r" else 0.12)
        line_text_box(c, xs[j] + 0.12, y, widths[j] - 0.24, head_h, htext, SANS_B, 11, OFFW, al)
    # lignes
    cy = y + head_h
    for i, row in enumerate(rows):
        if i % 2 == 1:
            rect(c, M, cy, CW, row_h, fill=BG)
        for j, val in enumerate(row):
            al = "r" if j in right_cols else "l"
            color = NAVY if j == 0 else INK
            font = SANS_B if j == 0 else SANS
            if j in change_cols:
                try:
                    num = float(str(val).replace("+", "").replace("%", "").replace(" ", "").replace(",", "."))
                    color = color_for_change(num); font = SANS_B
                except ValueError:
                    pass
            # wrap si texte long
            txt = str(val)
            if stringWidth(txt, font, 10.5) / IN > widths[j] - 0.24 and al == "l":
                para(c, xs[j] + 0.12, cy + 0.07, widths[j] - 0.24, row_h, txt, font, 10, color, leading=12)
            else:
                line_text_box(c, xs[j] + 0.12, cy, widths[j] - 0.24, row_h, txt, font, 10.5, color, al)
        hline(c, M, cy, CW, RULE, 0.4)
        cy += row_h
    hline(c, M, cy, CW, RULE, 0.4)
    footer(c)
    c.showPage()


def s_bar(c, title, section, categories, values, subtitle=None, signed=False,
          highlight_last=False, commentary=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    cw = CW if not commentary else CW * 0.64
    bar_chart(c, M, top + 0.2, cw, H_IN - top - 1.0, categories, values,
              signed=signed, highlight_last=highlight_last)
    if commentary:
        cx = M + cw + 0.5
        ckw = CW - cw - 0.5
        rect(c, cx, top + 0.35, 0.10, 0.9, fill=GOLD)
        for k, (h, b) in enumerate(commentary):
            yy = top + 0.30 + k * 1.5
            line_text_box(c, cx + 0.26, yy, ckw - 0.26, 0.4, h, SERIF_B, 15, NAVY, "l", vmid=False)
            para(c, cx + 0.26, yy + 0.45, ckw - 0.26, 1.0, b, SANS, 10.5, INK, leading=14)
    footer(c)
    c.showPage()


def s_line(c, title, section, categories, series, subtitle=None, colors=None, commentary=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    cw = CW if not commentary else CW * 0.64
    line_chart(c, M, top + 0.2, cw, H_IN - top - 1.0, categories, series,
               colors or [NAVY, GOLD, SLATE])
    if commentary:
        cx = M + cw + 0.5
        ckw = CW - cw - 0.5
        rect(c, cx, top + 0.35, 0.10, 0.9, fill=GOLD)
        for k, (h, b) in enumerate(commentary):
            yy = top + 0.30 + k * 1.5
            line_text_box(c, cx + 0.26, yy, ckw - 0.26, 0.4, h, SERIF_B, 15, NAVY, "l", vmid=False)
            para(c, cx + 0.26, yy + 0.45, ckw - 0.26, 1.0, b, SANS, 10.5, INK, leading=14)
    footer(c)
    c.showPage()


def s_gauges(c, title, section, gauges, subtitle=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    n = len(gauges); gap = 0.5
    cw = (CW - gap * (n - 1)) / n
    gsz = min(cw - 0.4, 2.5)
    y = top + 0.55
    for i, g in enumerate(gauges):
        x = M + i * (cw + gap)
        cx = x + cw / 2
        cy = y + gsz / 2
        accent = g.get("color", GOLD)
        donut(c, cx, cy, gsz / 2, gsz / 2 * 0.66, g["value"], accent)
        line_text_box(c, x, cy - 0.30, cw, 0.6, fmt_fr(g["value"], 0, suffix=" %"), SERIF_B, 30, NAVY, "c")
        line_text_box(c, x, y + gsz + 0.18, cw, 0.3, g["label"].upper(), SANS_B, 11, NAVY, "c", vmid=False)
        if g.get("context"):
            para(c, x, y + gsz + 0.52, cw, 0.8, g["context"], SANS, 9.5, SLATE, leading=12, align=TA_CENTER)
    footer(c)
    c.showPage()


def s_two_col(c, title, section, left, right, subtitle=None):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    gap = 0.5
    cw = (CW - gap) / 2
    ch = H_IN - top - 1.05
    for col, data in [(0, left), (1, right)]:
        accent = data["accent"]
        x = M + col * (cw + gap)
        soft_shadow_rect(c, x, top + 0.15, cw, ch, WHITE, radius=0.04)
        rect(c, x, top + 0.15, cw, 0.62, fill=accent, radius=0.04)
        rect(c, x, top + 0.40, cw, 0.37, fill=accent)
        line_text_box(c, x + 0.25, top + 0.15, cw - 0.5, 0.62, data["heading"].upper(), SANS_B, 13, WHITE, "l")
        yy = top + 0.95
        for item in data["items"]:
            line_text_box(c, x + 0.25, yy, 0.3, 0.3, "—", SANS_B, 12, accent, "l", vmid=False)
            hh = para(c, x + 0.55, yy - 0.02, cw - 0.8, 0.9, item, SANS, 10.5, INK, leading=13.5)
            yy += max(0.40, hh + 0.16)
    footer(c)
    c.showPage()


def s_palette(c, title, section, swatches, subtitle=None, typo=True):
    page_bg(c, WHITE)
    top = title_band(c, title, section, subtitle)
    n = len(swatches); gap = 0.3
    cw = (CW - gap * (n - 1)) / n
    y = top + 0.35
    chh = 1.7
    light = {"F5F6F8", "E2E5EA", "F7F7F4"}
    for i, (name, hexv, role) in enumerate(swatches):
        x = M + i * (cw + gap)
        rect(c, x, y, cw, chh, fill=HexColor("#" + hexv), stroke=RULE, lw=0.5, radius=0.05)
        tcol = NAVY if hexv in light else WHITE
        line_text_box(c, x + 0.18, y + chh - 0.42, cw - 0.36, 0.3, "#" + hexv, SANS_B, 11, tcol, "l", vmid=False)
        line_text_box(c, x, y + chh + 0.14, cw, 0.3, name, SANS_B, 11, NAVY, "c", vmid=False)
        para(c, x, y + chh + 0.46, cw, 0.6, role, SANS, 9.5, SLATE, align=TA_CENTER, leading=12)
    if typo:
        ty = y + chh + 1.45
        hline(c, M, ty, CW, RULE, 0.5)
        line_text_box(c, M, ty + 0.18, CW / 2, 0.5, "Playfair Display", SERIF_B, 26, NAVY, "l", vmid=False)
        para(c, M, ty + 0.72, CW / 2 - 0.3, 0.4, "Titres — empattements, ton presse premium", SANS, 10.5, SLATE)
        line_text_box(c, M + CW / 2, ty + 0.20, CW / 2, 0.5, "Arial / Sans-serif", SANS_B, 24, NAVY, "l", vmid=False)
        para(c, M + CW / 2, ty + 0.72, CW / 2, 0.4, "Corps, labels, chiffres, axes — lisibilité", SANS, 10.5, SLATE)
    footer(c)
    c.showPage()


def s_quote(c, quote, attribution):
    page_bg(c, BG)
    line_text_box(c, M + 0.1, 1.7, 1.5, 1.4, "“", SERIF_B, 120, GOLD, "l")
    para(c, M + 0.4, 2.95, CW - 1.2, 2.4, _it(quote), SERIF_I, 30, NAVY, leading=39)
    hline(c, M + 0.4, H_IN - 1.5, 1.2, GOLD, 1.5)
    line_text_box(c, M + 0.4, H_IN - 1.35, CW, 0.3, attribution, SANS_B, 13, STEEL, "l", vmid=False)
    footer(c)
    c.showPage()


def s_disclaimer(c, title, section, paragraphs):
    page_bg(c, WHITE)
    top = title_band(c, title, section)
    rect(c, M, top + 0.2, CW, H_IN - top - 1.0, fill=BG, stroke=RULE, lw=0.75, radius=0.02)
    para(c, M + 0.4, top + 0.5, CW - 0.8, H_IN - top - 1.5, paragraphs, SANS, 10.5, SLATE,
         leading=15.5, space_after=9)
    footer(c)
    c.showPage()


def s_closing(c, title, lines, contact):
    page_bg(c, NAVY)
    rect(c, M, 2.7, 0.6, 0.10, fill=GOLD)
    line_text_box(c, M, 2.95, CW, 0.9, title, SERIF_B, 46, OFFW, "l", vmid=False)
    para(c, M, 4.3, CW - 2.0, 1.0, lines, SANS, 16, MUTE_LT2, leading=24)
    hline(c, M, H_IN - 1.15, CW, STEEL, 1.0)
    line_text_box(c, M, H_IN - 1.0, 8, 0.35, contact, SANS_B, 13, GOLD, "l", vmid=False)
    line_text_box(c, W_IN - M - 6, H_IN - 1.0, 6, 0.35,
                  "Informatif et éducatif — aucun conseil en investissement", SANS_I, 10, MUTE_LT, "r", vmid=False)
    c.showPage()


# ---------------------------------------------------------------------------
#  BUILD
# ---------------------------------------------------------------------------

def build(path="/home/user/elio-os/wmb_deck/WMB_Presentation.pdf"):
    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle("WallStreet Market Brief — Business plan & présentation")
    c.setAuthor("WallStreet Market Brief")

    # 1 Couverture
    s_cover(c, "WallStreet Market Brief",
            "La newsletter financière premium de la Suisse francophone — vision, marché, modèle et exécution.",
            "Business plan & dossier de présentation", "Juin 2026 · Confidentiel")

    # 2 Sommaire
    s_contents(c, [
        ("1", "La vision", "Résumé, problème, solution, valeurs"),
        ("2", "L'offre produit", "Briefings, modules, thèmes, social"),
        ("3", "Marché & opportunité", "TAM/SAM/SOM, tendances, concurrence"),
        ("4", "Modèle économique", "Abonnement, économie unitaire, funnel"),
        ("5", "Produit, technologie & site", "Plateforme, pipeline IA, CMS, identité"),
        ("6", "Finances & risques", "Projections, scénarios, KPIs, feuille de route"),
    ])

    # SECTION 1
    s_divider(c, "1", "La vision",
              "Transformer le bruit permanent des marchés en un brief clair, court et structuré.")

    s_kpi_row(c, "Résumé exécutif", "La vision",
        [
            {"number": "49", "label": "CHF / mois", "context": "Abonnement premium récurrent, assumé pour un public à forte valeur.", "color": GOLD},
            {"number": "2 500\n3 500", "label": "Abonnés visés à 3 ans", "context": "Un noyau d'abonnés payants fidèles, Romandie puis francophonie.", "color": NAVY, "size": 40},
            {"number": "1,5–2 M", "label": "CHF de revenu annualisé", "context": "Run-rate visé à 3 ans, structure légère et rentable.", "color": NAVY, "size": 38},
        ],
        subtitle="Une niche claire et peu servie, une rigueur éditoriale absolue, une production industrialisée par l'IA.",
        footnote="WMB transforme le bruit des marchés en un brief lisible en quelques minutes, pour des francophones qui veulent comprendre sans y passer leurs journées. Positionnement strictement informatif et éducatif — jamais de conseil.")

    s_editorial(c, "Vision & mission", "La vision",
        [
            "Faire de WMB la référence francophone du brief de marché : aussi rigoureux qu'une salle de marché, aussi pédagogique qu'un bon professeur, et assez concis pour tenir dans un café du matin.",
            "À terme, WMB veut être le premier réflexe du francophone qui veut comprendre les marchés — l'équivalent suisse et francophone des grandes newsletters financières anglo-saxonnes.",
            "Sa mission : démocratiser une information financière de qualité, claire et honnête, pour un public aujourd'hui largement exclu. Ni jargon réservé aux professionnels, ni contenus racoleurs promettant des gains. WMB rend le lecteur autonome dans sa lecture des marchés.",
        ],
        pull_quote="Aussi rigoureux qu'une salle de marché, aussi pédagogique qu'un bon professeur.")

    s_cards(c, "Cinq valeurs fondatrices", "La vision",
        [
            {"kicker": "Clarté", "title": "Sans bruit", "body": "Sans blabla, sans jargon non expliqué. L'essentiel, lisible en quelques minutes.", "accent": NAVY},
            {"kicker": "Rigueur", "title": "Chaque chiffre sourcé", "body": "Daté et vérifié sur au moins deux sources. La fiabilité est un actif de marque.", "accent": GOLD},
            {"kicker": "Neutralité & indépendance", "title": "Aucun biais", "body": "Aucune opinion directionnelle, aucun biais commercial caché. La valeur vient de l'abonné, jamais d'un émetteur de produits financiers.", "accent": STEEL},
        ],
        subtitle="Clarté, rigueur, neutralité, pédagogie, indépendance.")

    s_editorial(c, "Le problème", "La vision",
        [
            "L'investisseur francophone moyen est mal servi. L'information de qualité est en anglais : Bloomberg, CNBC, newsletters US supposent de lire l'anglais financier couramment — une barrière réelle pour une large part du public romand, français et belge.",
            "L'information en français est soit institutionnelle et aride (rapports bancaires illisibles), soit racoleuse (contenus « devenez riche » sans rigueur ni cadre).",
            "Le bruit est permanent — réseaux sociaux, alertes, chaînes d'information continue : trop de signal faible, pas assez de synthèse fiable. Et le temps manque : la cible (cadres, professionnels, indépendants) n'a pas une heure par jour à consacrer à la veille de marché.",
        ],
        pull_quote="Trop de signal faible, pas assez de synthèse fiable.")

    s_cards(c, "La solution : l'essentiel, vérifié, structuré", "La vision",
        [
            {"kicker": "Chaque matin", "title": "Briefing Quotidien", "body": "Flash de quelques minutes sur la clôture de la veille : indices, devises, matières premières, crypto, fait marquant et agenda du jour.", "accent": NAVY},
            {"kicker": "Chaque dimanche", "title": "Brief hebdomadaire", "body": "Un dossier complet lisible en 15 à 20 minutes : USA, Monde, Suisse, Actions US, en cadre SI / ALORS / SINON.", "accent": GOLD},
            {"kicker": "Chaque mois", "title": "Thème du Mois", "body": "Une analyse approfondie d'un secteur ou d'une tendance : drivers, acteurs, risques, scénarios, glossaire.", "accent": STEEL},
        ],
        subtitle="Le tout en français, pédagogique, neutre — pour des décisions éclairées, sans jamais dire quoi faire.")

    # SECTION 2
    s_divider(c, "2", "L'offre produit",
              "Des formats complémentaires : une partie gratuite pour acquérir, une partie premium pour monétiser.")

    s_process(c, "Une offre en entonnoir", "L'offre produit",
        [
            ("Contenu social", "Instagram / Facebook adossés aux briefings, identité visuelle propre. Notoriété et haut de funnel."),
            ("Briefing Quotidien", "Le produit d'appel gratuit : visible, régulier, il installe l'habitude et fait connaître la marque."),
            ("Brief hebdomadaire", "Le cœur de l'offre payante : modules USA, Monde, Suisse, Actions US, chaque dimanche à 09h00."),
            ("Thème du Mois", "La montée en valeur : lecture longue qui renforce la valeur perçue de l'abonnement."),
        ],
        subtitle="Du gratuit qui acquiert au premium qui retient.")

    s_cards(c, "Le brief hebdomadaire : quatre modules", "L'offre produit",
        [
            {"kicker": "Module USA", "title": "La référence US", "body": "Indices, VIX, taux, crédit, Fed, secteurs, Mag 7, earnings, scénarios et risques.", "accent": NAVY},
            {"kicker": "Module Monde", "title": "International", "body": "Europe, Asie, émergents, géopolitique — la lecture globale des marchés.", "accent": STEEL},
            {"kicker": "Module Suisse", "title": "Spécificité helvétique", "body": "SMI, BNS, franc suisse, valeurs helvétiques, immobilier.", "accent": GOLD},
            {"kicker": "Actions US", "title": "Focus titres", "body": "Une sélection de valeurs en cadre conditionnel SI / ALORS / SINON, sans recommandation.", "accent": RED},
        ],
        subtitle="Un dossier complet, modulaire, lisible en 15 à 20 minutes.", cols=4)

    s_scoreboard(c, "Scoreboard des marchés", "L'offre produit — signature",
        [
            ("S&P 500", "5 633,09", 1.08), ("Nasdaq", "18 461,33", 1.24),
            ("Dow Jones", "41 209,75", 0.62), ("VIX", "16,84", -4.30),
            ("CAC 40", "7 554,29", -0.38), ("DAX", "18 720,10", 0.21),
            ("SMI", "12 048,55", 0.44), ("EUR/USD", "1,0876", 0.15),
            ("USD/CHF", "0,8521", -0.22), ("Or (oz)", "2 412,30", 0.73),
            ("WTI", "78,42", -1.15), ("Bitcoin", "63 180", 2.41),
        ])

    s_breaking(c, "Breaking",
               "La Fed maintient ses taux et ouvre la porte à une baisse en septembre",
               "Réservé aux mouvements majeurs : décisions Fed / BCE / BNS, variations d'indices supérieures à 3 %, surprises de résultats des Mag 7. Format éditorial à observer, sans interprétation directionnelle.")

    # SECTION 3
    s_divider(c, "3", "Marché & opportunité",
              "Un marché francophone premium peu servi, là où l'anglo-saxon est saturé.")

    s_kpi_row(c, "Dimensionnement du marché", "Marché & opportunité",
        [
            {"number": "Dizaines\nde millions", "label": "TAM — francophones", "context": "Suisse, France et Belgique sensibles à l'investissement.", "color": SLATE, "size": 30},
            {"number": "100–200 k", "label": "SAM — Romandie", "context": "5 à 10 % de la population romande adulte intéressée par l'info de marché.", "color": NAVY, "size": 36},
            {"number": "2 500\n3 500", "label": "SOM — à 3 ans", "context": "Une fraction de pour-cent du marché adressable. La contrainte est l'exécution.", "color": GOLD, "size": 40},
        ],
        subtitle="Approche TAM / SAM / SOM — dimensionnement stratégique, pas une promesse de résultat.",
        footnote="La Suisse compte ≈ 9,24 millions d'habitants début 2026, dont la Romandie représente près de 24 % (plus de 2 millions de personnes). Estimations combinant données publiques (OFS) et hypothèses de pénétration.")

    s_cards(c, "Quatre tendances porteuses", "Marché & opportunité",
        [
            {"kicker": "Retail", "title": "Investissement individuel", "body": "L'essor des courtiers en ligne élargit le public qui veut comprendre les marchés, y compris en Suisse.", "accent": NAVY},
            {"kicker": "Abonnement", "title": "Économie des newsletters", "body": "Forte croissance des newsletters payantes ; le public paie désormais directement l'éditeur pour du contenu de qualité.", "accent": GOLD},
            {"kicker": "Langue & pouvoir d'achat", "title": "La faille francophone", "body": "La quasi-totalité du contenu premium est anglophone. 49 CHF/mois est absorbable par la cible romande.", "accent": STEEL},
        ])

    s_table(c, "Paysage concurrentiel", "Marché & opportunité",
        ["Catégorie", "Exemples", "Limite pour la cible WMB"],
        [
            ["Médias financiers anglophones", "Bloomberg, CNBC, Morning Brew", "En anglais, peu centrés sur la Suisse"],
            ["Presse économique suisse/française", "Titres économiques, pages bourse", "Aride, lent, peu pédagogique"],
            ["Contenus finance réseaux sociaux", "Influenceurs, chaînes YouTube", "Souvent racoleurs, non sourcés, conseil déguisé"],
            ["Recherche bancaire", "Notes des banques", "Réservée aux clients, jargon, conflits d'intérêts"],
        ],
        subtitle="Quatre catégories, quatre limites pour le francophone qui veut comprendre vite et bien.",
        col_w=[0.32, 0.28, 0.40])

    s_two_col(c, "Le positionnement unique de WMB", "Marché & opportunité",
        left={"heading": "Ce que WMB est", "accent": GREEN, "items": [
            "Français et pensé pour la Suisse romande, puis France et Belgique.",
            "Premium : qualité, format et prix assumés pour un public à forte valeur.",
            "Strictement éducatif : contexte, faits, scénarios — jamais de conseil.",
            "Plus pédagogique que la presse, plus rigoureux que les influenceurs.",
        ]},
        right={"heading": "Ce que WMB n'est pas", "accent": RED, "items": [
            "Ni une newsletter anglophone de plus, déconnectée de la Suisse.",
            "Ni un rapport bancaire aride réservé aux initiés.",
            "Ni un compte « finance » racoleur promettant des gains.",
            "Ni un service de conseil ou de signaux de trading réglementé.",
        ]},
        subtitle="Un espace vide à l'intersection de quatre attributs rarement réunis : français + suisse + premium + éducatif.")

    s_cards(c, "Avantages défendables", "Marché & opportunité",
        [
            {"kicker": "Marque", "title": "Confiance & rigueur", "body": "Une réputation de fiabilité se construit lentement et se copie difficilement.", "accent": GOLD},
            {"kicker": "Habitude", "title": "Le rendez-vous matinal", "body": "Le briefing quotidien crée un réflexe récurrent et une rétention forte.", "accent": NAVY},
            {"kicker": "Pipeline", "title": "Production assistée par IA", "body": "Un savoir-faire propriétaire (schémas, prompts, contrôle qualité) difficile à reproduire en solo.", "accent": STEEL},
        ],
        subtitle="Quatre moats : marque, habitude quotidienne, pipeline propriétaire et identité éditoriale.")

    # SECTION 4
    s_divider(c, "4", "Modèle économique",
              "Un abonnement récurrent premium, alimenté par un funnel gratuit et la publicité Meta.")

    s_kpi_row(c, "Le modèle d'abonnement", "Modèle économique",
        [
            {"number": "49", "label": "CHF / mois", "context": "Prix de référence mensuel, volontairement premium.", "color": GOLD},
            {"number": "490", "label": "CHF / an", "context": "Option annuelle (deux mois offerts) — meilleure rétention.", "color": NAVY},
            {"number": "5–15 $", "label": "Marché dominant", "context": "WMB assume un volume plus faible mais une valeur par abonné élevée.", "color": SLATE, "size": 42},
        ],
        subtitle="Source principale : l'abonnement. Sources secondaires à activer : B2B, sponsoring encadré, contenus premium.",
        footnote="Le positionnement à 49 CHF impose une exigence de qualité et une différenciation forte entre le gratuit et le payant. Toute source de revenu est subordonnée à l'indépendance éditoriale.")

    s_kpi_row(c, "Économie unitaire (illustrative)", "Modèle économique",
        [
            {"number": "750–880", "label": "CHF · LTV brute", "context": "À 49 CHF/mois et 15–18 mois de durée de vie moyenne.", "color": GREEN, "size": 40},
            {"number": "40–100", "label": "CHF · CAC", "context": "Via le funnel gratuit + publicité Meta.", "color": NAVY, "size": 40},
            {"number": "> 5 : 1", "label": "Ratio LTV / CAC", "context": "Cible saine, compatible avec une croissance financée par la marge.", "color": GOLD, "size": 44},
        ],
        subtitle="Une économie unitaire saine est la condition d'une croissance autofinancée.",
        footnote="Chiffres illustratifs fondés sur des hypothèses explicites — ni données réalisées, ni garantie de résultat.")

    s_process(c, "Le funnel d'acquisition", "Modèle économique",
        [
            ("Notoriété", "Contenu Instagram/Facebook + publicité Meta ciblée sur la Romandie."),
            ("Capture", "Inscription gratuite au Briefing Quotidien, l'aimant principal."),
            ("Activation", "Le lecteur prend l'habitude du rendez-vous matinal."),
            ("Conversion", "Passage au payant, déclenché par la valeur du brief hebdo et des thèmes."),
            ("Rétention", "Qualité constante, offre annuelle, montée en valeur via les thèmes."),
        ],
        subtitle="De l'inconnu à l'abonné fidèle, en cinq étapes mesurables.")

    s_bar(c, "Repères de conversion gratuit → payant", "Modèle économique",
        ["Plancher\nmarché", "Médiane\nmarché", "Cible\nWMB", "Meilleures\npublications"],
        [2.0, 3.0, 3.5, 8.0],
        subtitle="La conversion se matérialise souvent après plusieurs mois d'exposition gratuite.",
        highlight_last=False,
        commentary=[
            ("2 à 5 %", "La fourchette typique du marché des newsletters payantes, médiane autour de 3 %."),
            ("Réchauffer dans la durée", "Un briefing quotidien régulier entretient l'audience jusqu'à la conversion."),
        ])

    s_cards(c, "Les canaux d'acquisition", "Modèle économique",
        [
            {"kicker": "Payant", "title": "Meta", "body": "Facebook/Instagram : canal principal, mesurable et scalable, ciblé sur la Romandie.", "accent": GOLD},
            {"kicker": "Organique", "title": "Social & SEO", "body": "Contenu visuel WMB régulier ; glossaire et thématiques comme aimant de recherche.", "accent": NAVY},
            {"kicker": "Relationnel", "title": "Parrainage & partenariats", "body": "Bouche-à-oreille entre abonnés, médias romands, communautés finance, événements locaux.", "accent": STEEL},
        ])

    # SECTION 5
    s_divider(c, "5", "Produit, technologie & site",
              "Une plateforme éditoriale pilotée par schémas, plus qu'une simple newsletter.")

    s_table(c, "Architecture technique", "Produit & technologie",
        ["Couche", "Technologie", "Rôle"],
        [
            ["Base de données", "TiDB + Drizzle ORM", "Données structurées des contenus et abonnés"],
            ["Stockage de fichiers", "S3", "Assets et médias"],
            ["Paiements / abonnements", "Stripe", "Gestion de l'abonnement à 49 CHF/mois"],
            ["E-mailing", "Brevo (SMTP)", "Envoi des briefings et modules"],
            ["Données de marché", "Finnhub", "Cours, indices, devises (API)"],
            ["Authentification", "OAuth + JWT (cookies HttpOnly)", "Sessions sécurisées"],
            ["Front", "PWA, rendu côté client", "Application installable, expérience native"],
        ],
        subtitle="Une pile moderne, sécurisée et adaptée à une petite équipe.",
        col_w=[0.26, 0.34, 0.40])

    s_process(c, "Le pipeline éditorial assisté par IA", "Produit & technologie",
        [
            ("Détection", "Dates et session de marché concernées, détectées automatiquement."),
            ("Recherche", "Multi-sources, avec vérification croisée (au moins deux sources par chiffre)."),
            ("Assemblage", "Dans un schéma JSON ou Markdown contraint, un format par type de contenu."),
            ("Contrôle qualité", "Cohérence des chiffres, sources datées, conventions de format."),
        ],
        subtitle="L'avantage opérationnel de WMB : une cadence quotidienne soutenable à très petite équipe.")

    s_cards(c, "Le site : trois produits en un", "Le site WMB",
        [
            {"kicker": "Média récurrent", "title": "Les briefings", "body": "Briefing quotidien gratuit et brief hebdomadaire premium (US + Monde + Suisse), 20 minutes de lecture.", "accent": NAVY},
            {"kicker": "Référence", "title": "La bibliothèque", "body": "Glossaire de 1 007+ termes, 54 indicateurs techniques, 15+ thématiques d'investissement.", "accent": GOLD},
            {"kicker": "Pédagogie", "title": "La formation", "body": "Un produit pédagogique à l'investissement et des outils complémentaires.", "accent": STEEL},
        ],
        subtitle="Un fonds documentaire durable qui justifie le premium et nourrit le référencement naturel.")

    s_editorial(c, "Le CMS piloté par schémas", "Le site WMB",
        [
            "C'est le cœur opérationnel du site, et son originalité. Le contenu n'est pas mis en page à la main : il est produit dans des formats stricts que le site transforme automatiquement en interface.",
            "Les modules (USA, Monde, Suisse, Actions US, Briefing) sont produits en JSON, selon un schéma précis par module. Le site lit le JSON et génère cartes, tableaux et sections.",
            "Le Thème du Mois est produit en Markdown. Le site détecte automatiquement des motifs et déclenche des widgets premium : cartes numérotées pour les drivers, cartes colorées par niveau pour les risques, cartes SI / ALORS / SINON pour les scénarios.",
            "Séparer le contenu du rendu supprime la mise en page manuelle, garantit une cohérence visuelle parfaite et permet une cadence quotidienne soutenable. C'est l'atout structurel majeur.",
        ],
        pull_quote="Contenu structuré, rendu automatisé : l'atout structurel majeur.")

    s_palette(c, "L'identité visuelle", "Le site WMB",
        [
            ("Ardoise foncée", "344453", "Thème du site, chrome PWA"),
            ("Navy éditorial", "1E3A5F", "Mots surlignés, barres, tags"),
            ("Rouge BREAKING", "B23A3A", "Cartouche et mode alerte"),
            ("Or", "C9A34E", "Mots de titre, accents"),
        ],
        subtitle="Une charte systématisée : modes standard et BREAKING, symboles par sujet, règles strictes.")

    s_cards(c, "Les forces de la plateforme", "Analyse du site",
        [
            {"kicker": "Structure", "title": "CMS par schémas", "body": "Vitesse et cohérence : la mise en page manuelle disparaît, la cadence devient soutenable.", "accent": GOLD},
            {"kicker": "Valeur", "title": "Une vraie plateforme", "body": "Glossaire, indicateurs, thématiques, formation : justifie le premium, crée du SEO, augmente la rétention.", "accent": NAVY},
            {"kicker": "Marque", "title": "Identité & discipline", "body": "Charte systématisée, double sourcing, datation, zéro donnée inventée : la fiabilité est un actif.", "accent": STEEL},
        ])

    s_two_col(c, "Faiblesses & chantiers prioritaires", "Analyse du site",
        left={"heading": "Risques opérationnels", "accent": RED, "items": [
            "Copier-coller manuel : un JSON malformé peut casser le rendu.",
            "Rendu côté client : risque d'indexation SEO du contenu de référence.",
            "Source de données unique (Finnhub) : exposition aux écarts de prix.",
            "Dépendance au fondateur et au pipeline ; observabilité limitée.",
        ]},
        right={"heading": "Recommandations", "accent": GREEN, "items": [
            "Validateur de schéma à la publication (champs, types, zéro astérisque).",
            "Sécuriser le SEO (rendu serveur ou pré-rendu des pages de référence).",
            "Double source pour les prix critiques, source de référence explicite.",
            "Analytics relié aux KPIs ; dunning Stripe et mise en avant de l'annuel.",
        ]},
        subtitle="Les chantiers ne touchent pas à la vision, solide, mais à la robustesse opérationnelle.")

    # SECTION 6
    s_divider(c, "6", "Finances & risques",
              "Une trajectoire modélisée, des scénarios explicites, une discipline du risque.")

    s_editorial(c, "Conformité & cadre éditorial", "Finances & risques",
        [
            "La règle fondatrice de WMB est aussi sa protection juridique : informatif et éducatif uniquement.",
            "Jamais de conseil en investissement. Jamais de recommandation d'achat ou de vente. Jamais d'objectif de prix ni de signal de trading. Jamais de promesse de performance.",
            "WMB livre du contexte, des faits, des chiffres sourcés et des scénarios conditionnels (SI / ALORS / SINON). Un disclaimer figure dans chaque édition, chaque e-mail et chaque thème.",
            "Compte tenu du cadre suisse, une validation juridique formelle des mentions et du périmètre éditorial est recommandée avant toute montée en échelle.",
        ],
        pull_quote="Informatif et éducatif uniquement — jamais de conseil.")

    s_cards(c, "Équipe & organisation", "Finances & risques",
        [
            {"kicker": "Aujourd'hui", "title": "Fondateur & pipeline", "body": "Un fondateur-rédacteur en chef ; le pipeline IA démultiplie une structure très légère.", "accent": NAVY},
            {"kicker": "Demain", "title": "Recrutements clés", "body": "Acquisition/growth, développement, éditorial, design/social — à mesure que le revenu se consolide.", "accent": GOLD},
            {"kicker": "Philosophie", "title": "Équipe resserrée", "body": "Forte automatisation : la marge dégagée finance la croissance, pas une structure lourde.", "accent": STEEL},
        ])

    s_table(c, "Trajectoire à 3 ans — scénario de base", "Finances & risques",
        ["Indicateur", "Fin année 1", "Fin année 2", "Fin année 3"],
        [
            ["Audience gratuite (liste)", "≈ 8 000", "≈ 25 000", "≈ 50 000"],
            ["Abonnés payants", "≈ 300", "≈ 1 200", "≈ 3 000"],
            ["Revenu mensuel récurrent (MRR)", "≈ 14 700 CHF", "≈ 58 800 CHF", "≈ 147 000 CHF"],
            ["Revenu annualisé (run-rate ARR)", "≈ 176 000 CHF", "≈ 706 000 CHF", "≈ 1 760 000 CHF"],
            ["Revenu encaissé (ordre de grandeur)", "≈ 80 000 CHF", "≈ 400 000 CHF", "≈ 1 100 000 CHF"],
        ],
        subtitle="Projections illustratives fondées sur des hypothèses explicites — non une garantie de résultat.",
        col_w=[0.40, 0.20, 0.20, 0.20], right_cols=[1, 2, 3])

    s_bar(c, "Revenu mensuel récurrent (MRR)", "Finances & risques",
        ["Année 1", "Année 2", "Année 3"], [14700, 58800, 147000],
        subtitle="Le MRR de fin d'année — moteur du revenu annualisé.",
        highlight_last=True,
        commentary=[
            ("× 10 en trois ans", "De ≈ 14 700 CHF à ≈ 147 000 CHF de MRR, porté par la montée en charge des abonnés."),
            ("Run-rate ≈ 1,76 M", "Le revenu annualisé de fin d'année 3, base d'un modèle rentable et durable."),
        ])

    s_line(c, "Audience gratuite & abonnés payants", "Finances & risques",
        ["Lancement", "Année 1", "Année 2", "Année 3"],
        {"Audience gratuite": [1000, 8000, 25000, 50000], "Abonnés payants": [40, 300, 1200, 3000]},
        subtitle="Le gratuit alimente le payant : la conversion se construit dans la durée.",
        colors=[SLATE, GOLD],
        commentary=[
            ("Le gratuit d'abord", "Une liste qui croît de 1 000 à 50 000 inscrits, réchauffée par le briefing quotidien."),
            ("≈ 3 000 abonnés", "Le SOM visé : une fraction de pour-cent du marché adressable."),
        ])

    s_cards(c, "Trois scénarios à 3 ans", "Finances & risques",
        [
            {"kicker": "Prudent", "title": "≈ 1 200–1 500 abonnés", "body": "Conversion ≈ 2 %, croissance plus lente. Le modèle reste viable et léger.", "accent": SLATE},
            {"kicker": "Base", "title": "≈ 3 000 abonnés", "body": "Conversion 3–4 %, rétention soutenue par l'annuel. Run-rate ARR ≈ 1,76 M CHF.", "accent": GOLD},
            {"kicker": "Ambitieux", "title": "5 000+ abonnés", "body": "Conversion ≈ 5 %, extension France/Belgique réussie et offre B2B. Revenus diversifiés.", "accent": GREEN},
        ],
        subtitle="La taille du marché n'est pas la contrainte ; l'exécution l'est.")

    s_gauges(c, "Indicateurs clés à piloter", "Finances & risques",
        [
            {"value": 4, "label": "Conversion cible", "context": "Gratuit → payant, dans la fourchette 3–4 % du scénario de base.", "color": GOLD},
            {"value": 45, "label": "Ouverture e-mail", "context": "Engagement des briefings — indicateur d'habitude et de rétention.", "color": NAVY},
            {"value": 60, "label": "Part d'abonnés annuels", "context": "L'annuel stabilise le revenu et réduit le churn.", "color": STEEL},
        ],
        subtitle="Croissance de la liste, conversion, MRR/ARR, churn, CAC, LTV/CAC, engagement, NPS.")

    s_table(c, "Risques & mitigation", "Finances & risques",
        ["Risque", "Description", "Mitigation"],
        [
            ["Dépendance à un canal", "Forte dépendance à Meta (coûts, algorithme)", "Diversifier : organique, parrainage, SEO"],
            ["Churn élevé", "Le prix premium augmente l'exigence de valeur", "Offre annuelle, qualité, montée en valeur"],
            ["Différenciation faible", "Si le payant n'apporte pas assez", "Réserver une vraie valeur exclusive"],
            ["Risque réglementaire", "Requalification en conseil financier", "Discipline éditoriale + validation juridique"],
            ["Dépendance au fondateur", "Concentration des compétences clés", "Documenter, recruter, automatiser"],
            ["Fiabilité des données", "Une erreur chiffrée nuit à la confiance", "Double sourcing, contrôle qualité, datation"],
        ],
        subtitle="Six risques majeurs, chacun assorti d'une mitigation concrète.",
        col_w=[0.24, 0.40, 0.36])

    s_process(c, "Feuille de route", "Finances & risques",
        [
            ("0–6 mois", "Consolider la cadence, finaliser le dashboard, lancer les premières campagnes Meta, structurer le funnel."),
            ("6–18 mois", "Atteindre le premier millier d'abonnés, affiner l'économie unitaire, tester l'offre B2B."),
            ("18–36 mois", "Étendre à la France et la Belgique, diversifier les revenus, installer WMB comme la référence."),
        ],
        subtitle="Court, moyen et long terme : de la consolidation à la référence francophone.")

    s_quote(c,
            "La taille du marché n'est pas la contrainte. La contrainte — et l'opportunité — est l'exécution : la régularité, la qualité, et la patience de construire la confiance.",
            "WallStreet Market Brief — Conclusion du business plan")

    s_disclaimer(c, "Disclaimer", "Mentions légales",
        [
            "WallStreet Market Brief (WMB) produit un contenu strictement informatif et éducatif. Les éditions, briefings, modules et thèmes ne constituent en aucun cas un conseil en investissement, une recommandation d'achat ou de vente, un objectif de prix, un signal de trading ou une promesse de performance.",
            "Les chiffres de marché présentés dans ce document sont illustratifs et destinés à montrer le format éditorial. Les projections financières sont des hypothèses de travail fondées sur des paramètres explicites ; elles ne constituent ni des données réalisées, ni une garantie de résultat.",
            "Chaque décision d'investissement relève de la seule responsabilité du lecteur, le cas échéant après consultation d'un conseiller dûment autorisé. Compte tenu du cadre réglementaire suisse relatif aux services financiers, le périmètre éditorial et les mentions de WMB font l'objet d'une validation juridique avant toute montée en échelle.",
            "Document de travail confidentiel destiné à la direction, aux partenaires et aux investisseurs potentiels. Toute diffusion externe requiert l'accord préalable de WMB.",
        ])

    s_closing(c, "Merci",
        "WallStreet Market Brief — la référence francophone du brief de marché. Clean, rigoureux, premium, strictement éducatif.",
        "wallstreetmarketbrief.ch")

    c.save()
    print("PDF enregistré :", path, "| pages :", PAGE["n"] + 6)  # +sections/couv approximatif
    return path


if __name__ == "__main__":
    build()
