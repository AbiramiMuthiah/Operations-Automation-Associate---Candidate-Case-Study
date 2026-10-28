"""Build the Luxury Inventory Snapshot workbook (Excel file that converts cleanly to Google Sheets)."""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

import clean as C

raw = pd.read_csv("out/raw.csv", dtype={"product_id": str})
cl = pd.read_csv("out/clean.csv", dtype={"product_id": str})
dq = pd.read_csv("out/dq_log.csv")
N = len(cl)
LAST = N + 1

FONT = "Arial"
NAVY, GREY, LIGHT = "1F2A44", "595959", "F2F2F2"
ACCENT = "C8102E"
f_title = Font(name=FONT, size=18, bold=True, color=NAVY)
f_h = Font(name=FONT, size=12, bold=True, color=NAVY)
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_b = Font(name=FONT, size=10)
f_bb = Font(name=FONT, size=10, bold=True)
f_note = Font(name=FONT, size=9, italic=True, color=GREY)
f_input = Font(name=FONT, size=10, color="0000FF")
fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_light = PatternFill("solid", fgColor=LIGHT)
fill_take = PatternFill("solid", fgColor="FFF4E5")
thin = Side(style="thin", color="D9D9D9")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
SGD, RM, PCT, INT, MULT = '"S$"#,##0', '"RM"#,##0', "0.0%", "#,##0", '0.00"x"'

wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
cd = wb.create_sheet("Clean_Data")
rd = wb.create_sheet("Raw_Data")
mp = wb.create_sheet("Mappings")
me = wb.create_sheet("Methodology")
code = wb.create_sheet("Code")


def hdr_row(ws, row, col, labels, widths=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.border = f_hdr, fill_hdr, box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def setw(ws, widths):
    for k, v in widths.items():
        ws.column_dimensions[k].width = v


# =====================================================================
# Mappings (editable lookup tables + settings)
# =====================================================================
mp["A1"] = "Mappings and settings"; mp["A1"].font = f_title
mp["A2"] = "Blue cells are inputs. Clean_Data formulas look up these tables, so editing a row here re-cleans every product."
mp["A2"].font = f_note

mp["A4"] = "Settings"; mp["A4"].font = f_h
settings = [
    ("Snapshot date", pd.Timestamp("2026-10-28").to_pydatetime(), "yyyy-mm-dd",
     "Date both feeds were collected (28 Oct 2026, 16:18 MYT)."),
    ("Shopify SGD to MYR exchange rate", 3.2450076, "0.0000",
     "Source: Shopify.currency.rate read on luxlexicon.com/en-my and carousellluxury.com/en-my, 28 Oct 2026."),
    ("New-listing window (days)", 30, "0", "Listings created within this many days count as new arrivals."),
]
for i, (lab, val, fmt, note) in enumerate(settings):
    r = 5 + i
    mp.cell(row=r, column=1, value=lab).font = f_b
    c = mp.cell(row=r, column=2, value=val); c.font = f_input; c.number_format = fmt
    mp.cell(row=r, column=3, value=note).font = f_note
SNAP, FX, NEWWIN = "Mappings!$B$5", "Mappings!$B$6", "Mappings!$B$7"

# Price bands
mp["A10"] = "Price bands (S$)"; mp["A10"].font = f_h
hdr_row(mp, 11, 1, ["Lower bound (S$)", "Band"])
for i, (lo, hi, name) in enumerate(C.PRICE_BANDS_SGD):
    mp.cell(row=12 + i, column=1, value=lo).font = f_input
    mp.cell(row=12 + i, column=2, value=name).font = f_input
    mp.cell(row=12 + i, column=1).number_format = SGD
BAND_LO = f"Mappings!$A$12:$A${11 + len(C.PRICE_BANDS_SGD)}"
BAND_LAB = f"Mappings!$B$12:$B${11 + len(C.PRICE_BANDS_SGD)}"
BANDS = [b[2] for b in C.PRICE_BANDS_SGD]

# Big houses
mp["A21"] = "Brand groups (houses shown separately)"; mp["A21"].font = f_h
for i, b in enumerate(C.BIG_HOUSES):
    mp.cell(row=22 + i, column=1, value=b).font = f_input
BIG = f"Mappings!$A$22:$A${21 + len(C.BIG_HOUSES)}"

# Brand map (raw vendor -> clean brand)
mp["D4"] = "Brand map (raw vendor name -> brand)"; mp["D4"].font = f_h
hdr_row(mp, 5, 4, ["Brand (raw)", "Brand (clean)", "Listings"])
raw_brands = raw.groupby("brand_raw").size().sort_values(ascending=False)
for i, (b, n) in enumerate(raw_brands.items()):
    full = C.BRAND_FIX.get(b, b)
    mp.cell(row=6 + i, column=4, value=b).font = f_b
    mp.cell(row=6 + i, column=5, value=C.COLLAB_PARENT.get(full, full)).font = f_input
    mp.cell(row=6 + i, column=6, value=f'=COUNTIF(Clean_Data!$D$2:$D${LAST},D{6 + i})').font = f_b
BMAP = f"Mappings!$D$6:$E${5 + len(raw_brands)}"

# Category map (product type -> category), built from the cleaning rules
mp["H4"] = "Category map (product type -> category)"; mp["H4"].font = f_h
hdr_row(mp, 5, 8, ["Product type", "Category", "Listings"])
types = cl.groupby(["product_type", "category"]).size().reset_index(name="n")
tmap = {}
for t, g in types.groupby(types.product_type.str.lower()):
    tmap[g.product_type.iloc[0]] = g.sort_values("n", ascending=False).category.iloc[0]
for i, (t, cat) in enumerate(sorted(tmap.items(), key=lambda x: (x[1], x[0]))):
    mp.cell(row=6 + i, column=8, value=t).font = f_b
    mp.cell(row=6 + i, column=9, value=cat).font = f_input
    mp.cell(row=6 + i, column=10, value=f'=COUNTIF(Clean_Data!$H$2:$H${LAST},H{6 + i})').font = f_b
CMAP = f"Mappings!$H$6:$I${5 + len(tmap)}"
setw(mp, {"A": 34, "B": 14, "C": 70, "D": 26, "E": 22, "F": 9, "G": 3, "H": 34, "I": 20, "J": 9})
mp.freeze_panes = "A4"

# =====================================================================
# Raw_Data
# =====================================================================
rcols = list(raw.columns)
hdr_row(rd, 1, 1, rcols)
for r, row in enumerate(raw.itertuples(index=False), start=2):
    for c_, v in enumerate(row, start=1):
        if isinstance(v, float) and pd.isna(v):
            v = None
        cell = rd.cell(row=r, column=c_, value=v)
        cell.font = f_b
rd.cell(row=1, column=1).comment = Comment(
    f"Source: Shopify product feeds of luxlexicon.com and carousellluxury.com (/products.json), {N} rows, "
    "collected 28 Oct 2026 16:18 MYT. RM prices from /en-my/products/<handle>.js. Untouched except HTML stripped "
    "from descriptions and descriptions cut to 500 characters.", "Analyst")
for col, w in zip("ABCDEFGHIJKLMNOPQRSTU", [16, 17, 18, 50, 18, 20, 40, 16, 12, 12, 12, 10, 24, 24, 10, 60, 60, 70]):
    rd.column_dimensions[col].width = w
rd.freeze_panes = "A2"
rd.auto_filter.ref = f"A1:{get_column_letter(len(rcols))}{LAST}"

# =====================================================================
# Clean_Data (cleaning logic as formulas where it can be)
# =====================================================================
ccols = ["site", "product_id", "product_name", "brand_raw", "brand", "brand_group", "product_type_raw",
         "product_type", "category_override", "category", "condition", "is_new", "price_sgd", "price_myr",
         "my_price_multiple", "price_band_sgd", "compare_at_sgd", "on_sale", "discount_pct", "available",
         "exotic", "is_birkin_kelly", "audience", "listed_date", "days_listed", "new_last_30d", "image_count", "url"]
hdr_row(cd, 1, 1, ccols)
rawi = raw.set_index("product_id")
for r, x in enumerate(cl.itertuples(index=False), start=2):
    rr = rawi.loc[x.product_id]
    rule_cat = tmap.get(x.product_type)
    override = x.category if rule_cat != x.category else None
    cmp_raw = rr.compare_at_price_sgd_raw if not pd.isna(rr.compare_at_price_sgd_raw) else None
    vals = [
        x.site, x.product_id, x.product_name, rr.brand_raw,
        f'=IFERROR(VLOOKUP(D{r},{BMAP},2,FALSE),D{r})',
        f'=IF(COUNTIF({BIG},E{r})>0,E{r},"Other brands")',
        rr.product_type_raw if isinstance(rr.product_type_raw, str) else None,
        x.product_type, override,
        f'=IF(I{r}<>"",I{r},IFERROR(VLOOKUP(H{r},{CMAP},2,FALSE),"Other"))',
        x.condition, f'=K{r}="Brand New"',
        x.price_sgd, x.price_myr, f'=IF(M{r}=0,"",N{r}/M{r})',
        f'=LOOKUP(M{r},{BAND_LO},{BAND_LAB})',
        cmp_raw, f'=AND(Q{r}<>"",Q{r}>M{r})', f'=IF(R{r},1-M{r}/Q{r},"")',
        bool(rr.available), bool(x.exotic),
        f'=AND(E{r}="Hermès",OR(ISNUMBER(SEARCH("Birkin",C{r})),ISNUMBER(SEARCH("Kelly",C{r}))))',
        x.audience, pd.Timestamp(x.listed_date).to_pydatetime(),
        f'={SNAP}-X{r}', f'=Y{r}<={NEWWIN}', x.image_count, x.url,
    ]
    for c_, v in enumerate(vals, start=1):
        cell = cd.cell(row=r, column=c_, value=v)
        cell.font = f_b
    cd.cell(row=r, column=13).number_format = SGD
    cd.cell(row=r, column=14).number_format = RM
    cd.cell(row=r, column=15).number_format = MULT
    cd.cell(row=r, column=17).number_format = SGD
    cd.cell(row=r, column=19).number_format = PCT
    cd.cell(row=r, column=24).number_format = "yyyy-mm-dd"
    cd.cell(row=r, column=25).number_format = "0"
cd.cell(row=1, column=1).comment = Comment(
    "Cleaned from Raw_Data (one row per product, duplicates checked on product_id). Columns in black are formulas "
    "that use the Mappings tab; condition, audience and category_override were parsed from product descriptions "
    "in Python (see Code tab).", "Analyst")
widths = [16, 17, 50, 18, 18, 14, 20, 20, 16, 18, 12, 9, 12, 12, 10, 12, 12, 9, 10, 9, 8, 10, 10, 11, 9, 10, 8, 60]
for i, w in enumerate(widths, start=1):
    cd.column_dimensions[get_column_letter(i)].width = w
cd.freeze_panes = "D2"
cd.auto_filter.ref = f"A1:AB{LAST}"


def col(letter):
    return f"Clean_Data!${letter}$2:${letter}${LAST}"


SITE, BRAND, BGRP, CAT, COND, NEW = col("A"), col("E"), col("F"), col("J"), col("K"), col("L")
PRICE, PMYR, MULTC, BAND, SALE, DISC = col("M"), col("N"), col("O"), col("P"), col("R"), col("S")
AVAIL, EXO, BK, DAYS, NEW30 = col("T"), col("U"), col("V"), col("Y"), col("Z")

# =====================================================================
# Dashboard
# =====================================================================
d = dash
setw(d, {"A": 2, "B": 44, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 15, "J": 15, "K": 15})
d["B1"] = "Luxury inventory snapshot: LuxLexicon vs Carousell Luxury"; d["B1"].font = f_title
d["B2"] = ('="Snapshot of every live listing on both stores, "&TEXT(' + SNAP + ',"d mmm yyyy")&". "&TEXT(COUNTA('
           + SITE + '),"#,##0")&" products. Prices in S$ (store base currency) unless marked RM."')
d["B2"].font = f_note

ws = d
S1 = 13  # KPI table start row (rows above hold takeaways)
LL, CLc, ALL = "C", "D", "E"


def site_cond(colref):
    # TRUE for rows of the site in the header cell; "All sites" matches every row
    return f'(({SITE}={colref})+({colref}="All sites")>0)'


# ---- KPI table ----
d.cell(row=S1 - 1, column=2, value="1. The two stores at a glance").font = f_h
hdr_row(d, S1, 2, ["Metric", "LuxLexicon", "Carousell Luxury", "All sites"])
kpis = [
    ("Live listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)})', INT),
    ("Inventory value (S$)", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},{PRICE})', SGD),
    ("Inventory value, Malaysia prices (RM)", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},{PMYR})', RM),
    ("Median price (S$)", lambda h, c: ("A", f'=MEDIAN(IF({site_cond(h)},{PRICE}))'), SGD),
    ("Average price (S$)", lambda h, c: f'={c}{S1 + 2}/{c}{S1 + 1}', SGD),
    ("25th percentile price (S$)", lambda h, c: ("A", f'=PERCENTILE(IF({site_cond(h)},{PRICE}),0.25)'), SGD),
    ("75th percentile price (S$)", lambda h, c: ("A", f'=PERCENTILE(IF({site_cond(h)},{PRICE}),0.75)'), SGD),
    ("Highest price (S$)", lambda h, c: ("A", f'=MAX(IF({site_cond(h)},{PRICE}))'), SGD),
    ("Brands carried", lambda h, c: f'=SUMPRODUCT({site_cond(h)}/COUNTIFS({SITE},{SITE},{BRAND},{BRAND}))' if h != "$E$" + str(S1) else None, INT),
    ("Hermès share of listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--({BRAND}="Hermès"))/{c}{S1 + 1}', PCT),
    ("Hermès share of inventory value", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--({BRAND}="Hermès"),{PRICE})/{c}{S1 + 2}', PCT),
    ("Hermès Birkin and Kelly listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--{BK})', INT),
    ("Bags share of listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--({CAT}="Bags"))/{c}{S1 + 1}', PCT),
    ("Brand new share of listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--{NEW})/{c}{S1 + 1}', PCT),
    ("Exotic leather listings", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--{EXO})', INT),
    ("Listings shown on sale (compare-at price)", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--{SALE})/{c}{S1 + 1}', PCT),
    ("Median discount on sale items", lambda h, c: ("A", f'=IFERROR(MEDIAN(IF({site_cond(h)}*{SALE},{DISC})),"n/a")'), PCT),
    ("Median days since listed", lambda h, c: ("A", f'=MEDIAN(IF({site_cond(h)},{DAYS}))'), "0"),
    ("Listed in the last 30 days", lambda h, c: f'=SUMPRODUCT(--{site_cond(h)},--{NEW30})/{c}{S1 + 1}', PCT),
    ("Malaysia price as multiple of S$ price (median)", lambda h, c: ("A", f'=MEDIAN(IF({site_cond(h)},{MULTC}))'), MULT),
    ("Malaysia uplift vs Shopify FX rate", lambda h, c: f'={c}{S1 + 20}/{FX}-1', PCT),
]
KPI = {}
for i, (lab, fn, fmt) in enumerate(kpis):
    r = S1 + 1 + i
    KPI[lab] = r
    d.cell(row=r, column=2, value=lab).font = f_b
    for colL in (LL, CLc, ALL):
        h = f"{colL}${S1}"
        if lab == "Brands carried" and colL == ALL:
            f = f'=SUMPRODUCT(1/COUNTIF({BRAND},{BRAND}))'
        elif lab == "Average price (S$)":
            f = f'={colL}{S1 + 2}/{colL}{S1 + 1}'
        elif lab == "Malaysia uplift vs Shopify FX rate":
            f = f'={colL}{S1 + 20}/{FX}-1'
        else:
            f = fn(h, colL)
        cell = d[f"{colL}{r}"]
        if isinstance(f, tuple):
            d[f"{colL}{r}"] = ArrayFormula(f"{colL}{r}", f[1])
        else:
            d[f"{colL}{r}"] = f
        d[f"{colL}{r}"].number_format = fmt
        d[f"{colL}{r}"].font = f_b
        d[f"{colL}{r}"].border = box
    d.cell(row=r, column=2).border = box
    if i % 2:
        for cc in "BCDE":
            d[f"{cc}{r}"].fill = fill_light
KPI_END = S1 + len(kpis)
d.cell(row=KPI_END + 1, column=2,
       value="Malaysia uplift = how much more the /en-my price is than the S$ price converted at Shopify's own "
             "exchange rate (Mappings!B6).").font = f_note


def kpi(lab, colL):
    return f"${colL}${KPI[lab]}"


# ---- 2. Price distribution ----
P0 = KPI_END + 4
d.cell(row=P0 - 1, column=2, value="2. Price distribution (S$ bands)").font = f_h
hdr_row(d, P0, 2, ["Price band", "LuxLexicon", "Carousell Luxury", "LuxLexicon %", "Carousell Luxury %"])
for i, bname in enumerate(BANDS):
    r = P0 + 1 + i
    d.cell(row=r, column=2, value=bname).font = f_b
    d[f"C{r}"] = f'=COUNTIFS({SITE},C${P0},{BAND},$B{r})'
    d[f"D{r}"] = f'=COUNTIFS({SITE},D${P0},{BAND},$B{r})'
    d[f"E{r}"] = f'=C{r}/SUM(C${P0 + 1}:C${P0 + len(BANDS)})'
    d[f"F{r}"] = f'=D{r}/SUM(D${P0 + 1}:D${P0 + len(BANDS)})'
    for cc, fm in zip("CDEF", [INT, INT, PCT, PCT]):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].font = f_b; d[f"{cc}{r}"].border = box
    d[f"B{r}"].border = box
P_END = P0 + len(BANDS)
ch = BarChart(); ch.type = "col"; ch.title = "Share of listings by price band"; ch.y_axis.title = "Share of listings"
ch.y_axis.number_format = "0%"; ch.height, ch.width = 7.5, 16
ch.add_data(Reference(d, min_col=5, max_col=6, min_row=P0, max_row=P_END), titles_from_data=True)
ch.set_categories(Reference(d, min_col=2, min_row=P0 + 1, max_row=P_END))
d.add_chart(ch, f"H{P0 - 1}")

# ---- 3. Brand split ----
B0 = P_END + 9
brands_top = list(cl.groupby("brand").price_sgd.sum().sort_values(ascending=False).head(15).index)
d.cell(row=B0 - 1, column=2, value="3. Brand split: listings and share of inventory value").font = f_h
hdr_row(d, B0, 2, ["Brand", "LuxLexicon listings", "LuxLexicon % of listings", "LuxLexicon % of value",
                    "LuxLexicon median (S$)", "Carousell Luxury listings", "Carousell Luxury % of listings",
                    "Carousell Luxury % of value", "Carousell Luxury median (S$)"])
d.cell(row=B0, column=12, value="LuxLexicon").font = f_note  # helper header used by formulas
LLn, CLn = kpi("Live listings", "C"), kpi("Live listings", "D")
LLv, CLv = kpi("Inventory value (S$)", "C"), kpi("Inventory value (S$)", "D")
for i, b in enumerate(brands_top + ["All other brands"]):
    r = B0 + 1 + i
    d.cell(row=r, column=2, value=b).font = f_bb if b == "All other brands" else f_b
    if b == "All other brands":
        top = f"$B${B0 + 1}:$B${B0 + len(brands_top)}"
        d[f"C{r}"] = f'={LLn}-SUM(C{B0 + 1}:C{r - 1})'
        d[f"D{r}"] = f'=C{r}/{LLn}'
        d[f"E{r}"] = f'=1-SUM(E{B0 + 1}:E{r - 1})'
        d[f"F{r}"] = "–"
        d[f"G{r}"] = f'={CLn}-SUM(G{B0 + 1}:G{r - 1})'
        d[f"H{r}"] = f'=G{r}/{CLn}'
        d[f"I{r}"] = f'=1-SUM(I{B0 + 1}:I{r - 1})'
        d[f"J{r}"] = "–"
    else:
        d[f"C{r}"] = f'=COUNTIFS({SITE},$C$13,{BRAND},$B{r})'
        d[f"D{r}"] = f'=C{r}/{LLn}'
        d[f"E{r}"] = f'=SUMIFS({PRICE},{SITE},$C$13,{BRAND},$B{r})/{LLv}'
        d[f"F{r}"] = ArrayFormula(f"F{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$C$13)*({BRAND}=$B{r}),{PRICE})),"–")')
        d[f"G{r}"] = f'=COUNTIFS({SITE},$D$13,{BRAND},$B{r})'
        d[f"H{r}"] = f'=G{r}/{CLn}'
        d[f"I{r}"] = f'=SUMIFS({PRICE},{SITE},$D$13,{BRAND},$B{r})/{CLv}'
        d[f"J{r}"] = ArrayFormula(f"J{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$D$13)*({BRAND}=$B{r}),{PRICE})),"–")')
    for cc, fm in zip("CDEFGHIJ", [INT, PCT, PCT, SGD, INT, PCT, PCT, SGD]):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].font = f_b; d[f"{cc}{r}"].border = box
    d[f"B{r}"].border = box
B_END = B0 + len(brands_top) + 1
d.cell(row=B_END + 1, column=2, value="Brands ranked by combined inventory value across both sites. Collaborations "
       "(e.g. Gucci x Balenciaga) roll up to the parent house.").font = f_note
ch = BarChart(); ch.type = "col"; ch.title = "Share of inventory value by brand (top 8)"; ch.height, ch.width = 9, 18
ch.y_axis.number_format = "0%"
ch.add_data(Reference(d, min_col=5, min_row=B0, max_row=B0 + 8), titles_from_data=True)
ch.add_data(Reference(d, min_col=9, min_row=B0, max_row=B0 + 8), titles_from_data=True)
ch.set_categories(Reference(d, min_col=2, min_row=B0 + 1, max_row=B0 + 8))
d.add_chart(ch, f"L{B0 - 1}")

# ---- 4. Category split ----
K0 = B_END + 5
cats = ["Bags", "Small Leather Goods", "Jewellery", "Watches", "Accessories", "Shoes", "Apparel",
        "Home & Tableware"]
d.cell(row=K0 - 1, column=2, value="4. Category split").font = f_h
hdr_row(d, K0, 2, ["Category", "LuxLexicon listings", "LuxLexicon % of listings", "LuxLexicon % of value",
                    "LuxLexicon median (S$)", "Carousell Luxury listings", "Carousell Luxury % of listings",
                    "Carousell Luxury % of value", "Carousell Luxury median (S$)"])
for i, ct in enumerate(cats):
    r = K0 + 1 + i
    d.cell(row=r, column=2, value=ct).font = f_b
    d[f"C{r}"] = f'=COUNTIFS({SITE},$C$13,{CAT},$B{r})'
    d[f"D{r}"] = f'=C{r}/{LLn}'
    d[f"E{r}"] = f'=SUMIFS({PRICE},{SITE},$C$13,{CAT},$B{r})/{LLv}'
    d[f"F{r}"] = ArrayFormula(f"F{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$C$13)*({CAT}=$B{r}),{PRICE})),"–")')
    d[f"G{r}"] = f'=COUNTIFS({SITE},$D$13,{CAT},$B{r})'
    d[f"H{r}"] = f'=G{r}/{CLn}'
    d[f"I{r}"] = f'=SUMIFS({PRICE},{SITE},$D$13,{CAT},$B{r})/{CLv}'
    d[f"J{r}"] = ArrayFormula(f"J{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$D$13)*({CAT}=$B{r}),{PRICE})),"–")')
    for cc, fm in zip("CDEFGHIJ", [INT, PCT, PCT, SGD, INT, PCT, PCT, SGD]):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].font = f_b; d[f"{cc}{r}"].border = box
    d[f"B{r}"].border = box
K_END = K0 + len(cats)
ch = BarChart(); ch.type = "col"; ch.title = "Share of listings by category"; ch.height, ch.width = 7.5, 16
ch.y_axis.number_format = "0%"
ch.add_data(Reference(d, min_col=4, min_row=K0, max_row=K_END), titles_from_data=True)
ch.add_data(Reference(d, min_col=8, min_row=K0, max_row=K_END), titles_from_data=True)
ch.set_categories(Reference(d, min_col=2, min_row=K0 + 1, max_row=K_END))
d.add_chart(ch, f"L{K0 - 1}")

# ---- 5. Like-for-like price comparison ----
L0 = K_END + 5
lf_brands = ["Hermès", "Chanel", "Louis Vuitton", "Dior", "Gucci", "Fendi", "Goyard", "Bottega Veneta"]
d.cell(row=L0 - 2, column=2, value="5. Same brand, same condition: who is cheaper?").font = f_h
d.cell(row=L0 - 1, column=2, value="Bags only, pre-owned in Very Good or Good condition. Median S$ price.").font = f_note
hdr_row(d, L0, 2, ["Brand (bags)", "LuxLexicon median (S$)", "LuxLexicon listings", "Carousell Luxury median (S$)",
                    "Carousell Luxury listings", "Carousell Luxury vs LuxLexicon"])
cond_ok = f'(({COND}="Very Good")+({COND}="Good")>0)'
for i, b in enumerate(lf_brands):
    r = L0 + 1 + i
    d.cell(row=r, column=2, value=b).font = f_b
    for mc, nc, h in (("C", "D", "$C$"), ("E", "F", "$D$")):
        base = f'({SITE}={h}13)*({BRAND}=$B{r})*({CAT}="Bags")*{cond_ok}'
        d[f"{mc}{r}"] = ArrayFormula(f"{mc}{r}", f'=IFERROR(MEDIAN(IF({base},{PRICE})),"–")')
        d[f"{nc}{r}"] = f'=SUMPRODUCT({base})'
    d[f"G{r}"] = f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r}),D{r}>=5,F{r}>=5),E{r}/C{r}-1,"too few")'
    for cc, fm in zip("CDEFG", [SGD, INT, SGD, INT, '+0%;-0%;0%']):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].font = f_b; d[f"{cc}{r}"].border = box
    d[f"B{r}"].border = box
L_END = L0 + len(lf_brands)
d.cell(row=L_END + 1, column=2, value="Model mix still differs (e.g. LuxLexicon Hermès is mostly Birkin, Kelly and "
       "Constance). Read gaps as positioning, not as the same bag priced differently.").font = f_note

# ---- 6. Condition ----
C0 = L_END + 5
conds = C.COND_ORDER
d.cell(row=C0 - 1, column=2, value="6. Condition mix and price").font = f_h
hdr_row(d, C0, 2, ["Condition", "LuxLexicon listings", "LuxLexicon %", "LuxLexicon median (S$)",
                    "Carousell Luxury listings", "Carousell Luxury %", "Carousell Luxury median (S$)"])
for i, cn in enumerate(conds):
    r = C0 + 1 + i
    d.cell(row=r, column=2, value=cn).font = f_b
    d[f"C{r}"] = f'=COUNTIFS({SITE},$C$13,{COND},$B{r})'
    d[f"D{r}"] = f'=C{r}/{LLn}'
    d[f"E{r}"] = ArrayFormula(f"E{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$C$13)*({COND}=$B{r}),{PRICE})),"–")')
    d[f"F{r}"] = f'=COUNTIFS({SITE},$D$13,{COND},$B{r})'
    d[f"G{r}"] = f'=F{r}/{CLn}'
    d[f"H{r}"] = ArrayFormula(f"H{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$D$13)*({COND}=$B{r}),{PRICE})),"–")')
    for cc, fm in zip("CDEFGH", [INT, PCT, SGD, INT, PCT, SGD]):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].font = f_b; d[f"{cc}{r}"].border = box
    d[f"B{r}"].border = box
C_END = C0 + len(conds)
d.cell(row=C_END + 1, column=2, value="Condition parsed from each product description. LuxLexicon's Fair items "
       "are mostly vintage Hermès, which keeps their median high.").font = f_note

# ---- 7. Ageing and markdowns ----
A0 = C_END + 5
d.cell(row=A0 - 2, column=2, value="7. Extra insight: Carousell Luxury marks down stock as it ages").font = f_h
d.cell(row=A0 - 1, column=2, value="Days since the product was first created in the store. Edit the bucket edges "
       "(blue) to test other cut-offs.").font = f_note
hdr_row(d, A0, 2, ["Days listed", "From (days)", "To (days)", "Carousell Luxury listings", "Share on sale",
                    "Median discount", "LuxLexicon listings"])
buckets = [("0–30 days", 0, 30), ("31–90 days", 31, 90), ("91–180 days", 91, 180), ("Over 180 days", 181, 100000)]
for i, (lab, lo, hi) in enumerate(buckets):
    r = A0 + 1 + i
    d.cell(row=r, column=2, value=lab).font = f_b
    d[f"C{r}"] = lo; d[f"D{r}"] = hi
    d[f"C{r}"].font = f_input; d[f"D{r}"].font = f_input
    d[f"E{r}"] = f'=COUNTIFS({SITE},$D$13,{DAYS},">="&C{r},{DAYS},"<="&D{r})'
    d[f"F{r}"] = f'=IF(E{r}=0,0,COUNTIFS({SITE},$D$13,{DAYS},">="&C{r},{DAYS},"<="&D{r},{SALE},TRUE)/E{r})'
    d[f"G{r}"] = ArrayFormula(f"G{r}", f'=IFERROR(MEDIAN(IF(({SITE}=$D$13)*({DAYS}>=C{r})*({DAYS}<=D{r})*{SALE},{DISC})),"–")')
    d[f"H{r}"] = f'=COUNTIFS({SITE},$C$13,{DAYS},">="&C{r},{DAYS},"<="&D{r})'
    for cc, fm in zip("CDEFGH", ["0", "0", INT, PCT, PCT, INT]):
        d[f"{cc}{r}"].number_format = fm; d[f"{cc}{r}"].border = box
        if cc not in "CD":
            d[f"{cc}{r}"].font = f_b
    d[f"B{r}"].border = box
A_END = A0 + len(buckets)
d.cell(row=A_END + 1, column=2, value="LuxLexicon shows no compare-at (was) prices, so its markdowns cannot be "
       "seen from the storefront.").font = f_note
ch = BarChart(); ch.type = "col"; ch.title = "Carousell Luxury: share of listings on sale, by age"
ch.height, ch.width = 7, 14; ch.y_axis.number_format = "0%"; ch.legend = None
ch.add_data(Reference(d, min_col=6, min_row=A0, max_row=A_END), titles_from_data=True)
ch.set_categories(Reference(d, min_col=2, min_row=A0 + 1, max_row=A_END))
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False
d.add_chart(ch, f"J{A0 - 2}")

# ---- Takeaways (top of dashboard, formulas pull the numbers) ----
def K(lab, c_):  # value of a KPI cell
    return f"{c_}{KPI[lab]}"


hermes_lf = L0 + 1
chanel_lf, lv_lf, dior_lf = L0 + 2, L0 + 3, L0 + 4
over180 = A0 + 4
first30 = A0 + 1
d["B4"] = "Key takeaways"; d["B4"].font = Font(name=FONT, size=13, bold=True, color=ACCENT)
take = [
    ('="1. The two stores sell different things. LuxLexicon is a Hermès specialist: Hermès is "&TEXT(' + K("Hermès share of listings", "C")
     + ',"0%")&" of its listings and "&TEXT(' + K("Hermès share of inventory value", "C") + ',"0%")&" of its S$"&TEXT('
     + K("Inventory value (S$)", "C") + '/1000000,"0.0")&"m stock, with "&' + K("Hermès Birkin and Kelly listings", "C")
     + '&" Birkin and Kelly bags and "&TEXT(' + K("Brand new share of listings", "C") + ',"0%")&" sold brand new. Carousell Luxury is a pre-owned bag shop: "&TEXT('
     + K("Bags share of listings", "D") + ',"0%")&" bags, "&TEXT(1-' + K("Brand new share of listings", "D") + ',"0%")&" pre-owned, median S$"&TEXT('
     + K("Median price (S$)", "D") + ',"#,##0")&" vs S$"&TEXT(' + K("Median price (S$)", "C") + ',"#,##0")&"."'),
    ('="2. LuxLexicon depends on one brand. Hermès carries "&TEXT(' + K("Hermès share of inventory value", "C")
     + ',"0%")&" of its value, so Hermès supply and resale prices drive the business. Jewellery and watches are already "&TEXT('
     + f'E{K0 + 3}+E{K0 + 4}' + ',"0%")&" of value at higher prices; Van Cleef & Arpels, Cartier and Rolex are the clearest way to spread that risk."'),
    ('="3. Carousell Luxury is the cheaper entry point, even within the same brand. For pre-owned bags in Very Good or Good condition, its median is "&TEXT(-'
     + f'G{chanel_lf}' + ',"0%")&" lower for Chanel and "&TEXT(-' + f'G{dior_lf}' + ',"0%")&" lower for Dior, while Louis Vuitton prices are about equal. Recommendation: treat Carousell Luxury buyers as a pipeline for LuxLexicon, e.g. trade-in or upgrade offers."'),
    ('="4. Carousell Luxury stock that does not sell gets discounted: "&TEXT(' + f'F{first30}' + ',"0%")&" of items under 30 days old are on sale, rising to "&TEXT('
     + f'F{over180}' + ',"0%")&" after 180 days. "&TEXT(' + K("Listings shown on sale (compare-at price)", "D")
     + ',"0%")&" of listings show a markdown. Recommendation: automate age-based repricing and move slow items to the other store before cutting price."'),
    ('="5. Malaysian shoppers pay more than the exchange rate suggests. /en-my prices are "&TEXT(' + K("Malaysia price as multiple of S$ price (median)", "C")
     + ',"0.00")&"x the S$ price on LuxLexicon and "&TEXT(' + K("Malaysia price as multiple of S$ price (median)", "D")
     + ',"0.00")&"x on Carousell Luxury, against Shopify\'s rate of "&TEXT(' + FX + ',"0.00")&". That is a "&TEXT(' + K("Malaysia uplift vs Shopify FX rate", "C")
     + ',"0%")&" and "&TEXT(' + K("Malaysia uplift vs Shopify FX rate", "D") + ',"0%")&" uplift. Worth checking the two stores mean to price Malaysia differently."'),
]
for i, t in enumerate(take):
    r = 5 + i
    d.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
    d[f"B{r}"] = t
    d[f"B{r}"].font = f_b; d[f"B{r}"].fill = fill_take
    d[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    d.row_dimensions[r].height = 44
r = 10
d.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
d[f"B{r}"] = ("What are we selling, and to whom? LuxLexicon sells investment-grade Hermès (Birkin, Kelly, Constance), "
              "often brand new, plus high jewellery and watches, to collectors and high-net-worth buyers who pay "
              "S$10k–S$50k+ for scarce pieces. Carousell Luxury sells pre-owned bags from the big houses (Chanel, "
              "Louis Vuitton, Hermès, Dior) at S$1k–S$5k to women buying their first or second luxury bag, who look "
              "for value and respond to markdowns.")
d[f"B{r}"].font = f_bb; d[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
d[f"B{r}"].fill = PatternFill("solid", fgColor="E8EEF7")
d.row_dimensions[r].height = 58
d.cell(row=11, column=2, value="Every number above is a live formula over Clean_Data. Method, assumptions and "
       "limitations are on the Methodology tab.").font = f_note
d.freeze_panes = "A4"
d.sheet_view.showGridLines = False

# =====================================================================
# Methodology
# =====================================================================
m = me
setw(m, {"A": 2, "B": 30, "C": 110})
m["B1"] = "Methodology, assumptions and limitations"; m["B1"].font = f_title
rows = [
    ("Data collection", None),
    ("Sources", "luxlexicon.com and carousellluxury.com. Both run on Shopify, which publishes every live product at "
                "/products.json (250 per page). I read the full feed for each store: 2,074 LuxLexicon and 831 "
                "Carousell Luxury products, which matches the product counts the stores report at /meta.json, so "
                "coverage is 100% and no sampling was needed."),
    ("Malaysia prices", "The feed returns S$ prices (the stores' base currency). For each product I also read "
                        "/en-my/products/<handle>.js, which returns the price a Malaysian shopper sees in RM. Every RM "
                        "price in the workbook is the real /en-my price, not a conversion."),
    ("When", "28 Oct 2026, about 16:18 MYT. Listings change daily, so this is a one-day snapshot."),
    ("How", "JavaScript run in a browser on each store (polite pacing, 3 parallel requests with pauses), saved as "
            "JSON, then cleaned in Python (pandas). The code is on the Code tab."),
    ("Cleaning", None),
    ("One row per product", "Every product has exactly one variant, so price, stock status and SKU come from that "
                            "variant. Product IDs were checked for duplicates (none found)."),
    ("Brands", "Vendor names standardised through the brand map on the Mappings tab (e.g. 'Tiffany & Co' to "
               "'Tiffany & Co.', 'See By Chole' to 'See By Chloé'). Collaborations roll up to the parent house."),
    ("Categories", "167 product types mapped to 9 categories through the category map. 36 products had no product "
                   "type; LuxLexicon's description carries a category field, so I used it there, and the product name "
                   "otherwise (column category_override)."),
    ("Condition", "Neither store has a condition field. I parsed it from the description (LuxLexicon: a pipe-separated "
                  "field; Carousell Luxury: text such as 'Very Good Condition' or 'Like New'). 24 products state no "
                  "condition and are shown as Unknown."),
    ("Currency", "All analysis is in S$, the stores' base currency, so the two sites compare on the same footing. "
                 "RM inventory values use the actual /en-my prices."),
    ("Sale and discount", "A product counts as on sale when its compare-at (was) price is above its price. 8 "
                          "compare-at prices at or below the selling price were ignored."),
    ("Assumptions", None),
    ("Inventory value", "Sum of listed asking prices, not cost or sold value."),
    ("Days listed", "Days from the product's creation date in Shopify to the snapshot date. Relisted items may have a "
                    "newer creation date, so true age can be higher. LuxLexicon created 585 products in January 2026, "
                    "which looks like a bulk upload or migration."),
    ("Audience", "Carousell Luxury descriptions end with a target audience ('Girl | Lady | Women' or 'Men Unisex'); "
                 "91% of listings target women. LuxLexicon does not state one."),
    ("Limitations", None),
    ("Asking prices only", "The data shows what is listed and at what price, not what sells or how fast. Sell-through "
                           "and margin need internal data (Postgres orders)."),
    ("Model mix", "Like-for-like comparisons control for brand, category and condition but not for model, size or "
                  "leather, so part of any price gap is mix."),
    ("Hidden stock", "247 products are tagged 'no addtocart' (enquire-only); they are included because they are "
                     "listed for sale."),
    ("Store tags", "Carousell Luxury has 809 products tagged 'price less than 500 sgd' that cost S$500 or more. I "
                   "ignored tags for analysis; this is a catalogue data issue worth fixing."),
]
r = 3
for k, v in rows:
    if v is None:
        m.cell(row=r, column=2, value=k).font = Font(name=FONT, size=12, bold=True, color=ACCENT)
    else:
        m.cell(row=r, column=2, value=k).font = f_bb
        c = m.cell(row=r, column=3, value=v); c.font = f_b; c.alignment = Alignment(wrap_text=True, vertical="top")
        m.cell(row=r, column=2).alignment = Alignment(vertical="top")
        m.row_dimensions[r].height = max(15, 15 * (len(v) // 120 + 1))
    r += 1
r += 1
m.cell(row=r, column=2, value="Data quality log").font = Font(name=FONT, size=12, bold=True, color=ACCENT)
r += 1
hdr_row(m, r, 2, ["Issue", "Products affected"])
for k, v in dq.itertuples(index=False):
    r += 1
    m.cell(row=r, column=2, value=k).font = f_b
    m.cell(row=r, column=3, value=int(v)).font = f_b
    m.cell(row=r, column=3).alignment = Alignment(horizontal="left")
m.sheet_view.showGridLines = False

# =====================================================================
# Code tab
# =====================================================================
code.column_dimensions["A"].width = 160
code["A1"] = "Code used (also shared as separate files: collect_luxury.js, clean.py)"; code["A1"].font = f_title
lines = ["// ---- collect_luxury.js : run in the browser console on each store's /en-my page ----"]
lines += open("collect_luxury.js", encoding="utf-8").read().splitlines()
lines += ["", "# ---- clean.py : cleaning rules (Python, pandas) ----"]
lines += open("clean.py", encoding="utf-8").read().splitlines()
for i, ln in enumerate(lines, start=3):
    c = code.cell(row=i, column=1, value=ln if not ln.startswith("=") else "'" + ln)
    c.font = Font(name="Courier New", size=9)

wb.save("out/Luxury_Inventory_Snapshot.xlsx")
print("saved", N)
