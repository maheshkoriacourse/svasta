#!/usr/bin/env python3
"""SVASTA investor deck builder v2 — 14 slides, full clean build."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

FOREST   = RGBColor(0x1B, 0x43, 0x32)
FOREST_D = RGBColor(0x0E, 0x2A, 0x1D)
FOREST_L = RGBColor(0x2D, 0x6A, 0x4F)
GOLD     = RGBColor(0xE9, 0xB4, 0x4C)
TERRA    = RGBColor(0xC1, 0x50, 0x2E)
CREAM    = RGBColor(0xFA, 0xF6, 0xEF)
SAND     = RGBColor(0xF1, 0xE8, 0xD7)
LEAF     = RGBColor(0x74, 0xA5, 0x7F)
INK      = RGBColor(0x22, 0x30, 0x1F)
MUTED    = RGBColor(0x5C, 0x6B, 0x58)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
SOFT     = RGBColor(0xD7, 0xE4, 0xDC)

HEAD, BODY = "Georgia", "Calibri"
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
N = [0]

def slide():
    N[0] += 1
    return prs.slides.add_slide(BLANK), N[0]

def rect(s, x, y, w, h, fill):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp

def txt(s, x, y, w, h, text, size=18, color=INK, bold=False, font=BODY,
        align=PP_ALIGN.LEFT, italic=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        r = p.add_run(); r.text = line
        f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic
        f.name = font; f.color.rgb = color
    return tb

def bullets(s, x, y, w, h, items, size=14, gap=12):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for lead, rest in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        r = p.add_run(); r.text = lead + "  "
        r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = FOREST; r.font.name = BODY
        r2 = p.add_run(); r2.text = rest
        r2.font.size = Pt(size); r2.font.color.rgb = INK; r2.font.name = BODY
    return tb

def card(s, x, y, w, h, big, label, dark=False):
    rect(s, x, y, w, h, FOREST_L if dark else WHITE)
    txt(s, x, y + h*0.22, w, 0.8, big, size=32, color=GOLD if dark else FOREST, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
    txt(s, x + 0.15, y + h*0.55, w - 0.3, 0.8, label, size=11.5, color=SOFT if dark else MUTED, align=PP_ALIGN.CENTER)

def eyebrow(s, text, x=0.9, y=0.55, dark=False):
    txt(s, x, y, 9, 0.35, text.upper(), size=12, color=GOLD if dark else TERRA, bold=True)

def pnum(n, dark=False):
    txt(prs.slides[-1], SW-0.9, SH-0.5, 0.5, 0.4, f"{n:02d}", size=11, color=GOLD if dark else MUTED, align=PP_ALIGN.RIGHT)

# ---------- 1 COVER ----------
s, n = slide(); rect(s, 0, 0, SW, SH, FOREST); rect(s, 0, 0, 0.18, SH, GOLD); rect(s, 8.6, 0, 4.74, SH, FOREST_L)
txt(s, 8.6, 1.9, 4.7, 1.2, "❁", size=92, color=GOLD, align=PP_ALIGN.CENTER)
txt(s, 8.6, 3.35, 4.7, 0.8, "SVASTA", size=46, color=GOLD, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
txt(s, 8.6, 4.2, 4.7, 0.5, "AYURVEDA · VERIFIED", size=13, color=CREAM, align=PP_ALIGN.CENTER)
txt(s, 0.9, 1.45, 7.3, 0.4, "INVESTOR & PARTNER BRIEFING — 2026", size=14, color=GOLD, bold=True)
txt(s, 0.9, 2.05, 7.4, 2.3, "Ancient Wisdom.\nBottled for\nthe World.", size=52, color=CREAM, bold=True, font=HEAD)
txt(s, 0.9, 4.7, 7.0, 0.9, "The global ayurvedic wellness beverage company —\n5,000 farmers · 12 countries · zero compromise.", size=16, color=SOFT)
txt(s, 0.9, 6.45, 7, 0.4, "CONFIDENTIAL · SVASTA WELLNESS PVT. LTD.", size=10, color=RGBColor(0x9F,0xB8,0xA8))
pnum(n, dark=True)

# ---------- 2 OPPORTUNITY ----------
s, n = slide(); rect(s, 0, 0, SW, SH, CREAM)
eyebrow(s, "The Opportunity")
txt(s, 0.9, 0.95, 11.3, 1.1, "The world is switching its drink.\nAyurveda is the answer it hasn't tasted yet.", size=30, color=FOREST, bold=True, font=HEAD)
bullets(s, 0.9, 2.75, 5.7, 3.6, [
    ("$1.9T", "global wellness economy — beverages are its fastest-growing aisle"),
    ("$54B", "herbal & botanical drinks market by 2030 (9.1% CAGR)"),
    ("68%", "of global consumers read ingredient labels before buying"),
    ("Zero", "home-grown global ayurvedic beverage brand — the captain's seat is open"),
])
rect(s, 7.0, 2.75, 5.4, 3.7, FOREST)
txt(s, 7.4, 3.05, 4.6, 0.4, "WHY NOW", size=13, color=GOLD, bold=True)
txt(s, 7.4, 3.6, 4.7, 2.7, "• Post-pandemic immunity is a permanent purchase driver\n• Sugar taxes & clean-label rules reshape soft drinks\n• 180M+ yoga/ayurveda practitioner mindshare worldwide\n• India's wellness exports grew 4x in five years", size=14, color=CREAM)
pnum(n)

# ---------- 3 BRAND STORY ----------
s, n = slide(); rect(s, 0, 0, SW, SH, SAND)
eyebrow(s, "Brand Story")
txt(s, 0.9, 0.95, 11.4, 1.5, "From a grandmother's kitchen in Kerala\nto shelves on four continents.", size=32, color=FOREST, bold=True, font=HEAD)
rect(s, 0.9, 2.95, 5.6, 3.55, WHITE)
txt(s, 1.25, 3.25, 4.9, 0.4, "THE BELIEF", size=12, color=TERRA, bold=True)
txt(s, 1.25, 3.7, 4.9, 2.6, "The world's oldest science of wellbeing deserves the world's newest standards of quality.\n\nSVASTA — from 'svastha': established in one's own best self. We don't bottle a drink; we bottle a daily ritual of returning to balance.", size=13.5, color=MUTED)
rect(s, 6.85, 2.95, 5.55, 3.55, FOREST)
txt(s, 7.2, 3.25, 4.9, 0.4, "THE FOUNDER'S LINE", size=12, color=GOLD, bold=True)
txt(s, 7.2, 3.75, 4.9, 2.0, '"My mother healed our whole street with a copper pot and seven herbs. SVASTA is that copper pot — reforged in steel, standardised in labs, shipped to the world."', size=15, color=CREAM, italic=True, font=HEAD)
txt(s, 7.2, 5.85, 4.8, 0.4, "— Mahesh Koria, Founder & CEO", size=13, color=GOLD, bold=True)
pnum(n)

# ---------- 4 PRODUCTS ----------
s, n = slide(); rect(s, 0, 0, SW, SH, CREAM)
eyebrow(s, "The Portfolio")
txt(s, 0.9, 0.9, 11, 0.7, "Four Rituals. Every Body. Every Day.", size=30, color=FOREST, bold=True, font=HEAD)
cards = [
    ("🌟", "SVASTA GOLDEN", "IMMUNITY + VITALITY", "Turmeric · Ashwagandha · Tulsi", "22 kcal · zero refined sugar", "$3.49"),
    ("🌿", "SVASTA COOL", "FOCUS + CALM", "Amla · Brahmi · Mint", "18 kcal · natural vitamin C", "$3.49"),
    ("⚡", "SVASTA BOOST", "ENERGY, NO CRASH", "Shatavari · Gokshura · Ginger", "25 kcal · electrolytes", "$3.99"),
    ("🌙", "SVASTA NIGHT", "DEEP REST", "Jatamansi · Chamomile · Nutmeg", "15 kcal · melatonin-free", "$3.99"),
]
for i, (em, name, bene, ings, call, price) in enumerate(cards):
    x = 0.9 + i * 3.0
    rect(s, x, 1.95, 2.8, 4.15, WHITE)
    rect(s, x, 1.95, 2.8, 0.14, GOLD if i in (0, 2) else TERRA)
    txt(s, x, 2.2, 2.8, 0.8, em, size=40, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 3.1, 2.4, 0.45, name, size=15, color=FOREST, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 3.55, 2.4, 0.35, bene, size=10, color=TERRA, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 4.0, 2.4, 0.6, ings, size=12, color=INK, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 4.7, 2.4, 0.4, call, size=10.5, color=MUTED, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 5.3, 2.4, 0.5, price + " / 250ml", size=14, color=FOREST, bold=True, align=PP_ALIGN.CENTER)
txt(s, 0.9, 6.45, 11.5, 0.5, "Classical ayurvedic pairings, standardised to clinical-grade actives. Zero artificial anything — in every market, no exceptions.", size=12.5, color=MUTED, italic=True)
pnum(n)

# ---------- 5 VIRAL ENGINE ----------
s, n = slide(); rect(s, 0, 0, SW, SH, FOREST)
rect(s, 0, 0, SW, 0.14, GOLD)
eyebrow(s, "Growth Engine", dark=True)
txt(s, 0.9, 0.95, 11, 0.7, "The Viral Engine: What's Your Dosha?", size=30, color=CREAM, bold=True, font=HEAD)
txt(s, 0.9, 1.9, 5.7, 2.3, "A 60-second quiz maps your ayurvedic dosha and assigns your daily ritual. Every result card is built to be shared — customers become the marketing department.", size=15, color=SOFT)
card(s, 0.9, 4.4, 2.6, 1.7, "2.1M", "quiz completions", dark=True)
card(s, 3.75, 4.4, 2.6, 1.7, "38%", "share rate", dark=True)
card(s, 6.9, 4.4, 2.6, 1.7, "27%", "buy-after-quiz", dark=True)
card(s, 9.75, 4.4, 2.6, 1.7, "$0", "seed-phase adspend", dark=True)
rect(s, 6.9, 1.9, 5.5, 2.15, FOREST_L)
txt(s, 7.25, 2.2, 4.8, 0.4, "CAMPAIGN STACK", size=12, color=GOLD, bold=True)
txt(s, 7.25, 2.65, 4.85, 1.5, "#MyDoshaMyRitual — always-on quiz engine\n#SwapTheSoda — 30-day UGC challenge (9 markets)\n#HarGharSwasth — India home-wellness movement", size=13.5, color=CREAM)
pnum(n, dark=True)

# ---------- 6 SCIENCE ----------
s, n = slide(); rect(s, 0, 0, SW, SH, CREAM)
eyebrow(s, "Ayurveda, Verified")
txt(s, 0.9, 0.95, 11, 0.7, "Tradition You Can Test.", size=30, color=FOREST, bold=True, font=HEAD)
txt(s, 0.9, 1.75, 10.5, 0.5, "We honour the classical texts — and we test like a pharma lab. That is the SVASTA standard.", size=15, color=MUTED)
card(s, 0.9, 2.9, 3.75, 2.6, "40+", "standardised actives — every herb lot fingerprinted before it enters a bottle")
card(s, 4.85, 2.9, 3.75, 2.6, "7", "university partnerships — ongoing clinical collaborations, India & EU")
card(s, 8.8, 2.9, 3.75, 2.6, "0", "artificial colours, flavours, preservatives or refined sugar — ever")
rect(s, 0.9, 5.95, 11.65, 0.9, SAND)
txt(s, 1.25, 6.15, 11, 0.6, "Certifications: FSSAI · USDA Organic · EU Organic · AYUSH · HACCP + GMP · Halal · Kosher", size=14, color=FOREST, bold=True, align=PP_ALIGN.CENTER)
pnum(n)

# ---------- 7 MARKETS ----------
s, n = slide(); rect(s, 0, 0, SW, SH, SAND)
eyebrow(s, "Global Footprint")
txt(s, 0.9, 0.95, 11, 0.7, "Born in India. Poured Everywhere.", size=30, color=FOREST, bold=True, font=HEAD)
card(s, 0.9, 2.2, 2.75, 1.9, "12", "countries live")
card(s, 3.85, 2.2, 2.75, 1.9, "1,400+", "retail doors")
card(s, 6.8, 2.2, 2.75, 1.9, "31", "online markets")
card(s, 9.75, 2.2, 2.7, 1.9, "94%", "D2C repeat rate")
txt(s, 0.9, 4.6, 4.0, 0.4, "PHASE-2 WAVE (2027-28)", size=12, color=TERRA, bold=True)
txt(s, 0.9, 5.05, 5.6, 1.8, "🇮🇳 India (HQ) · 🇦🇪 UAE · 🇸🇦 KSA · 🇬🇧 UK · 🇺🇸 USA\n🇸🇬 Singapore · 🇩🇪 Germany · 🇫🇷 France · 🇨🇦 Canada\n🇦🇺 Australia · 🇯🇵 Japan · 🇳🇱 Netherlands", size=15, color=INK)
rect(s, 6.8, 4.6, 5.6, 2.15, FOREST)
txt(s, 7.15, 4.9, 4.9, 0.4, "NEXT-GEN ENTRY", size=12, color=GOLD, bold=True)
txt(s, 7.15, 5.4, 4.9, 1.2, "Q1: Korea & Brazil via e-comm-only\nQ3: Nordic retail via partner distribution", size=14, color=CREAM)
pnum(n)

# ---------- 8 DISTRIBUTION ----------
s, n = slide(); rect(s, 0, 0, SW, SH, CREAM)
eyebrow(s, "Distribution Engine")
txt(s, 0.9, 0.95, 11, 0.7, "Every Shelf. Every Screen. Every City.", size=30, color=FOREST, bold=True, font=HEAD)
cols = [
    ("🛒", "Modern Trade", "42% of volume", "1,400+ premium grocery, convenience & pharmacy doors; planogram leadership in 6 chains."),
    ("📱", "D2C & Marketplaces", "35% of volume", "svasta.com subscription ritual boxes + Amazon, Flipkart, Noon, iHerb. 94% repeat rate."),
    ("🏨", "HoReCa & Airlines", "23% of volume", "5-star spas, gyms, cafés; onboard menus with 3 international carriers."),
]
for i, (ic, name, share, body) in enumerate(cols):
    x = 0.9 + i * 4.0
    rect(s, x, 2.1, 3.7, 3.9, WHITE)
    txt(s, x, 2.4, 3.75, 0.7, ic, size=40, align=PP_ALIGN.CENTER)
    txt(s, x + 0.25, 3.25, 3.5, 0.45, name, size=17, color=FOREST, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
    txt(s, x + 0.25, 3.62, 3.5, 0.4, share, size=15, color=TERRA, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.25, 4.15, 3.5, 1.7, body, size=12.5, color=MUTED, align=PP_ALIGN.CENTER)
txt(s, 0.9, 6.35, 11.5, 0.5, "Cold-chain certified logistics · 48-hour metro delivery · single global brand book, local campaign playbook per market.", size=12.5, color=MUTED, italic=True)
pnum(n)

# ---------- 9 SUSTAINABILITY ----------
s, n = slide(); rect(s, 0, 0, SW, SH, FOREST)
eyebrow(s, "Sustainability", dark=True)
txt(s, 0.9, 0.95, 11.5, 0.8, "Good for You. Good for the Grower.\nGood for the Ground.", size=28, color=CREAM, bold=True, font=HEAD)
sus = [
    ("5,000+", "Farmer Partners", "Direct-sourcing clusters; 3x-market organic premiums; zero-debt goal by 2028."),
    ("100%", "Circular Packaging", "Glass-first today; 90% rPET by 2027; zero packaging-to-landfill by 2030."),
    ("3x", "Water-Positive 2030", "Rainwater harvest + watershed restoration in every sourcing district."),
    ("Net-0", "Carbon 2030", "Solar-powered plants now; full value-chain neutrality, independently audited."),
]
for i, (big, name, body) in enumerate(sus):
    x = 0.9 + i * 2.95
    rect(s, x, 2.9, 2.75, 3.4, FOREST_L)
    txt(s, x, 3.25, 2.75, 0.7, big, size=30, color=GOLD, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 4.05, 2.35, 0.4, name, size=13, color=CREAM, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.25, 4.55, 2.25, 1.7, body, size=10.5, color=SOFT, align=PP_ALIGN.CENTER)
txt(s, 0.9, 6.7, 11.5, 0.4, "All 2030 commitments audited annually and published openly — wellness that costs the earth is not wellness.", size=12, color=SOFT, italic=True)
pnum(n, dark=True)

# ---------- 10 FINANCIALS ----------
s, n = slide(); rect(s, 0, 0, SW, SH, CREAM)
eyebrow(s, "The Business")
txt(s, 0.9, 0.95, 11, 0.7, "Unit Economics That Compound.", size=30, color=FOREST, bold=True, font=HEAD)
table_data = [
    ["Metric", "Year 1", "Year 3", "Year 5"],
    ["Net revenue", "$4.2M", "$46M", "$140M"],
    ["Gross margin", "52%", "58%", "61%"],
    ["EBITDA margin", "-12%", "14%", "22%"],
    ["D2C share", "38%", "41%", "45%"],
    ["LTV / CAC (D2C)", "3.1x", "4.4x", "5.2x"],
]
rows, cols_n = len(table_data), len(table_data[0])
tbl_shape = s.shapes.add_table(rows, cols_n, Inches(0.9), Inches(2.1), Inches(11.6), Inches(3.4))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(3.6)
for ci in range(1, 4):
    tbl.columns[cols_n - cols_n + 1].width = Inches(2.66)
for r in range(rows):
    for c in range(cols_n):
        cell = tbl.cell(r, c)
        cell.text = table_data[r][c]
        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        run = para.runs[0]
        run.font.size = Pt(14)
        run.font.name = BODY
        if r == 0:
            run.font.bold = True; run.font.color.rgb = CREAM
            cell.fill.solid(); cell.fill.fore_color.rgb = FOREST
        elif r % 2 == 1:
            run.font.color.rgb = INK; cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
        else:
            run.font.color.rgb = INK; cell.fill.solid(); cell.fill.fore_color.rgb = SAND
txt(s, 0.9, 5.85, 11.5, 0.9, "Model: subscription-first D2C for data & margin, modern trade for scale, HoReCa for margin + brand theatre. Unit economics positive per market at ~18 months; group EBITDA positive in Year 3.", size=12.5, color=MUTED)
pnum(n)

# ---------- 11 LEADERSHIP ----------
s, n = slide(); rect(s, 0, 0, SW, SH, SAND)
eyebrow(s, "Leadership")
txt(s, 0.9, 0.95, 11, 0.7, "Built by Operators. Guided by Healers.", size=30, color=FOREST, bold=True, font=HEAD)
lead = [
    ("MK", "Mahesh Koria", "FOUNDER & CEO", "JBIMS MBA; 19+ years scaling BFSI operations & data businesses. From Dharavi — raised by a maid who believed in plants and prayer. SVASTA is his love letter to both.", True),
    ("AR", "Dr. Anaya Raghavan", "CHIEF SCIENCE OFFICER", "PhD Pharmacognosy (IIT Bombay); 15 years in phytochemical standardisation; leads the clinical validation program.", False),
    ("RS", "Rohan Sharma", "CHIEF COMMERCIAL OFFICER", "Ex-global beverages VP; built 40-market distribution networks across GCC, EU and APAC.", False),
]
for i, (ini, name, role, bio, founder) in enumerate(lead):
    x = 0.9 + i * 4.0
    rect(s, x, 2.1, 3.75, 4.1, WHITE)
    if founder:
        rect(s, x, 2.1, 3.75, 0.12, GOLD)
    av = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1.475), Inches(2.45), Inches(0.8), Inches(0.8))
    av.fill.solid(); av.fill.fore_color.rgb = GOLD; av.line.fill.background(); av.shadow.inherit = False
    tf = av.text_frame; para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = ini; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = FOREST_D; r.font.name = HEAD
    txt(s, x + 0.2, 3.45, 3.35, 0.45, name, size=16, color=FOREST, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
    txt(s, x + 0.2, 3.95, 3.35, 0.35, role, size=9.5, color=TERRA, bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.35, 4.4, 3.05, 1.7, bio, size=11, color=MUTED, align=PP_ALIGN.CENTER)
pnum(n)

# ---------- 12 PARTNERS ----------
s, n = slide(); rect(s, 0, 0, SW, SH, SAND)
eyebrow(s, "Partnerships")
txt(s, 0.9, 0.95, 11, 0.7, "We Grow With the Best.", size=30, color=FOREST, bold=True, font=HEAD)
chips = [
    "5,000-Farm Sourcing Collective", "Taj Wellness Spas", "Emirates First-Class", "NABL Labs",
    "AIIMS Research Program", "Nature's Basket", "Amazon Global", "Whole Foods EU",
]
for i, ch in enumerate(chips):
    x = 0.9 + (i % 3) * 4.0
    y = 2.2 + (i // 3) * 1.15
    rect(s, x, y, 3.75, 0.85, WHITE)
    txt(s, x + 0.2, y + 0.18, 3.4, 0.5, ch, size=13.5, color=FOREST, bold=True, align=PP_ALIGN.CENTER)
txt(s, 0.9, 6.3, 11.5, 0.5, "Partnership thesis: credibility partners for science, reach partners for distribution, culture partners for story.", size=12.5, color=MUTED, italic=True)
pnum(n)

# ---------- 13 ROADMAP ----------
s, n = slide(); rect(s, 0, 0, SW, SH, FOREST)
rect(s, 0, 0, SW, 0.14, GOLD)
eyebrow(s, "Roadmap", dark=True)
txt(s, 0.9, 0.95, 11, 0.7, "From 12 Countries to a Global Category.", size=30, color=CREAM, bold=True, font=HEAD)
phases = [
    ("2026-27", "PROVE", "Deepen India + GCC; subscription engine to 100k rituals/month; 2 clinical studies published."),
    ("2027-28", "SCALE", "EU + APAC retail push; 1,400 → 4,000 doors; rPET line; ₹700M ARR run-rate."),
    ("2028-30", "LEAD", "Category captain: 20+ countries, Net-0 carbon, ayurvedic beverages = a named global category."),
]
for i, (yr, tag, body) in enumerate(phases):
    x = 0.9 + i * 4.0
    rect(s, x, 2.3, 3.75, 3.6, FOREST_L)
    txt(s, x + 0.25, 2.6, 3.2, 0.5, yr, size=20, color=GOLD, bold=True, font=HEAD)
    txt(s, x + 0.25, 3.15, 3.2, 0.4, tag, size=12, color=CREAM, bold=True)
    txt(s, x + 0.25, 3.7, 3.25, 1.8, body, size=12, color=SOFT)
rect(s, 0.9, 6.05, 11.65, 0.85, GOLD)
txt(s, 1.2, 6.25, 11, 0.5, "THE END GAME: SVASTA becomes to Ayurveda what Coca-Cola is to cola — the default global name, with dignity built in.", size=14, color=FOREST_D, bold=True, align=PP_ALIGN.CENTER)
pnum(n, dark=True)

# ---------- 14 CLOSING ----------
s, n = slide(); rect(s, 0, 0, SW, SH, FOREST)
txt(s, 1.2, 2.3, 11, 0.5, "❁", size=54, color=GOLD, align=PP_ALIGN.CENTER)
txt(s, 1.2, 3.2, 11, 0.9, "Your Body Knows. Now You Do Too.", size=36, color=CREAM, bold=True, font=HEAD, align=PP_ALIGN.CENTER)
txt(s, 1.2, 4.3, 11, 0.6, "Ancient Wisdom. Bottled for the World.", size=18, color=GOLD, align=PP_ALIGN.CENTER, italic=True)
txt(s, 1.2, 5.2, 11, 0.5, "hello@svasta.com  ·  svasta.com", size=15, color=SOFT, align=PP_ALIGN.CENTER)
txt(s, 1.2, 6.3, 11, 0.4, "© 2026 SVASTA Wellness Pvt. Ltd. · Confidential", size=10, color=RGBColor(0x9F,0xB8,0xA8), align=PP_ALIGN.CENTER)
pnum(n, dark=True)

prs.save("/home/maheshkoria/hermes-builds/svasta/SVASTA-deck.pptx")
print(f"DECK SAVED: {len(prs.slides.slides if hasattr(prs.slides,'slides') else prs.slides._sldIdLst)} slides -> SVASTA-deck.pptx")