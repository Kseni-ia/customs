#!/usr/bin/env python3
"""HS code summary — Sonavox shipment CZ26011282 / FCR TECGO26080029 / container CIMU1716969.

PRAVIDLO: HS kódy se nevytvářejí vlastním odvozením. V souhrnu smí být pouze kód,
který je doslova uveden v některém z podkladů (e-mail „SONAVOX HS KODY“ nebo předchozí
celní prohlášení Sonavox). Položky bez takového kódu zůstávají PRÁZDNÉ a vyžadují se
od klienta — viz list „Chybějící HS kódy“.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

S1 = "Suzhou Sonavox International Trading Co., Ltd"
S2 = "Suzhou Sonavox Electronics Co., Ltd"
S3 = "Suzhou Yanlong Electronic Product Co., Ltd"
INV = {S1: "CZ20260725", S2: "CZ20260725-1", S3: "CZ20260725-2"}

MISSING = "CHYBÍ – VYŽÁDAT OD KLIENTA"

# Zdroje kódů (doslovné citace z podkladů)
SRC_MAIL = "e-mail SONAVOX HS KODY"
SRC_CP = "předchozí celní prohlášení Sonavox"
SRC_BOTH = "e-mail SONAVOX HS KODY + předchozí celní prohlášení"

# sender, item, part, EN desc, qty, net, gross, EUR, HS (None = chybí), CZ desc, zdroj kódu
LINES = [
    # --- Sender 1 — CZ20260725 ---
    (S1, 1, "010-KD1622S-3.8A-CZ", "Voice coil", 65000, 86.5, 216.33, 9841.00, "8504500000",
     "Kmitací cívka – měděný drát (jednoduché vinutí) navinutý na papírovém válci (formeru)",
     SRC_MAIL + ": „Voice Coil / Kmitací cívka – HS 85045000 00“ (ZISZ 30-0292-2016)"),
    (S1, 2, "040-62E4-CR-CZ", "Front plate", 35000, 2345.0, 2385.00, 3542.00, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru",
     SRC_CP + ": „Front plate – deska“, CN 8518 90 00 DK 00 (POZOR: e-mail uvádí ...90 – rozpor)"),
    (S1, 3, "040-80I6-CRB-CZ", "Front plate", 32400, 5103.0, 5183.00, 7464.96, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru",
     SRC_CP + ": „Front plate – deska“, CN 8518 90 00 DK 00 (POZOR: e-mail uvádí ...90 – rozpor)"),
    (S1, 4, "040-54D3-CR-1-CZ", "Front plate", 25000, 1012.5, 1032.91, 1377.50, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru",
     SRC_CP + ": „Front plate – deska“, CN 8518 90 00 DK 00 (POZOR: e-mail uvádí ...90 – rozpor)"),
    (S1, 5, "021-100F00-9-CZ", "Spider", 5600, 10.5, 15.00, 449.68, "8518900000",
     "Středící membrána (spider) – naimpregnovaná textilie, součást autoreproduktoru",
     SRC_BOTH + ": „Spider – 8518 9000 00“"),
    (S1, 6, "021-80E15-3-CZ", "Spider", 72000, 19.4, 25.00, 7041.60, "8518900000",
     "Středící membrána (spider) – naimpregnovaná textilie, součást autoreproduktoru",
     SRC_BOTH + ": „Spider – 8518 9000 00“"),
    (S1, 19, "140-35D14-44XP1-CZ", "Cone", 8000, 16.0, 20.00, 2926.40, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru",
     SRC_BOTH + ": „Cone – HS 85189000 00“"),
    (S1, 10, "181-35113-1SN-CZ", "Terminal", 75000, 37.5, 40.00, 1612.50, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru",
     SRC_CP + ": „terminal – části reproduktorů“, CN 8518 90 00 DK 00"),
    (S1, 11, "181-35113-2SN-CZ", "Terminal", 125000, 62.5, 70.00, 2687.50, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru",
     SRC_CP + ": „terminal – části reproduktorů“, CN 8518 90 00 DK 00"),
    (S1, 12, "181-40133-1SN-CZ", "Terminal", 200000, 100.0, 110.00, 2540.00, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru",
     SRC_CP + ": „terminal – části reproduktorů“, CN 8518 90 00 DK 00"),
    (S1, 13, "181-40133-2SN-CZ", "Terminal", 200000, 100.0, 110.00, 2540.00, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru",
     SRC_CP + ": „terminal – části reproduktorů“, CN 8518 90 00 DK 00"),
    (S1, 7, "142-630H10-15H-CZ", "Gasket", 960, 7.2, 10.00, 100.80, "3926909790",
     "Plastové těsnění (gasket) – ostatní výrobky z plastů",
     SRC_CP + ": „Gassket – plastové těsnění“, CN 3926 90 97 DK 90"),
    (S1, 8, "142-40X25-25D-CZ", "Gasket", 2000, 15.0, 20.00, 75.80, "3926909790",
     "Plastové těsnění (gasket) – ostatní výrobky z plastů",
     SRC_CP + ": „Gassket – plastové těsnění“, CN 3926 90 97 DK 90"),
    (S1, 17, "JK-65425-01", "Gauge", 5000, 57.78, 70.00, 899.50, "3926909790",
     "Plastový kalibr (gauge) – montážní přípravek, ostatní výrobky z plastů",
     SRC_BOTH + ": „Gauge – 3926 9097 90“"),
    (S1, 18, "JK-35113-01-CZ", "Gauge", 2000, 23.11, 28.00, 302.00, "3926909790",
     "Plastový kalibr (gauge) – montážní přípravek, ostatní výrobky z plastů",
     SRC_BOTH + ": „Gauge – 3926 9097 90“"),
    (S1, 14, "029-2513-AL-CZ", "Nut", 180000, 299.7, 310.50, 12600.00, "7318159519",
     "Matice / vložka (nut, insert) – závitové zboží ze železa nebo oceli",
     SRC_MAIL + ": „Insert (nut) – 7318159519“ (POZOR: materiál dílu neověřen – viz Poznámky)"),
    (S1, 16, "008-XH-2Y-DK-6-CZ", "Harness", 51200, 235.52, 250.00, 8120.32, "7413000090",
     "Kabelový svazek – splétaná lanka z mědi, elektricky neizolovaná",
     SRC_CP + ": „harness – kabelový svazek“, CN 7413 00 00 DK 90"),
    # --- Sender 1 — BEZ KÓDU ---
    (S1, 9, "J2-10X002-0001-CZ", "Waterproof membrane", 210000, 30.24, 40.00, 8820.00, "8518900000",
     "Vodotěsná membrána – část autoreproduktoru",
     "e-mail Alice Knápková (Sonavox Technology CZ), 12.08.2026" + ": „Waterproof membrane – 8518900000“"),
    (S1, 15, "029A-000004-CZ", "Spring nut", 30000, 12.6, 15.00, 540.00, "731816",
     "Pružná matice (spring nut) – matice ze železa nebo oceli",
     "e-mail Alice Knápková (Sonavox Technology CZ), 12.08.2026" + ": „029A-000004-CZ – HS 731816“ (POZOR: jen 6 míst, "
     "nutno doplnit na 10 – 7318 16 má šest podpoložek)"),

    # --- Sender 2 — CZ20260725-1 ---
    (S2, 1, "212-8075-CZ", "Ceramics circle", 19152, 11299.68, 11688.80, 19918.08, "8518900000",
     "Magnetický obvod (ceramics circle) – feritové jádro (nezmagnetované), součást autoreproduktoru",
     SRC_MAIL + ": „CERAMICS CIRCLE / magnetický obvod – 8518900000“"),
    (S2, 2, "IM-50196-AUDICZ", "Car Speaker", 2016, 1149.12, 1219.04, 5140.80, "8518210000",
     "Reproduktor (autoreproduktor), jediný v jedné skříni",
     SRC_BOTH + ": „SPEAKER / reproduktor – 8518210000“, CN 8518 21 00 DK 00"),

    # --- Sender 3 — CZ20260725-2 ---
    (S3, 1, "140-80I27-42X-CZ", "Cone", 33600, 856.8, 1134.00, 16836.96, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru",
     SRC_BOTH + ": „Cone – HS 85189000 00“"),
    (S3, 2, "140-80E30-58X-CZ", "Cone", 2160, 52.2, 84.12, 1280.88, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru",
     SRC_BOTH + ": „Cone – HS 85189000 00“"),
    (S3, 3, "140-40E17-72XC-CZ", "Cone", 8000, 48.0, 56.30, 2712.80, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru",
     SRC_BOTH + ": „Cone – HS 85189000 00“"),
]

HS_META = {
    "8504500000": ("Ne", "Ne", "Zařazení podloženo závaznou informací ZISZ 30-0292-2016."),
    "8518900000": ("Ne", "Ne", "Části a součásti reproduktorů, DK 00 dle předchozích celních prohlášení. "
                               "Zahrnuje i Waterproof membrane dle sdělení klienta z 12.08.2026."),
    "8518210000": ("Ne", "Ne", "Kompletní autoreproduktor, jediný v jedné skříni."),
    "3926909790": ("Ne", "Ne", "Ostatní výrobky z plastů, DK 90 dle předchozího celního prohlášení."),
    "7413000090": ("Ne", "Ne", "Splétaná lanka z mědi, neizolovaná, DK 90 dle předchozího celního prohlášení."),
    "731816": ("ANO – Y137", "Ne", "Matice ze železa/oceli. Kód 7318 16 NENÍ v seznamu nař. (EU) 2022/191, "
                                    "antidumping se tedy neuplatní. POZOR: kód je jen 6místný, nutno doplnit na 10."),
    "7318159519": ("ANO – Y137", "ANO – 90,2 %", "Antidumping potvrzen výpočtem v Helios: 3,7 % clo + 86,5 % AD = 90,2 %, "
                                  "základ 304 416 Kč → 274 584 Kč. CBAM pokryt prohlášením Y137. "
                                  "ROZPOR: spring nut má dle klienta 7318 16 (matice, bez AD) – viz Poznámky."),
    None: ("K OVĚŘENÍ", "K OVĚŘENÍ", "Bez HS kódu nelze určit. Vyžádat od klienta – viz list „Chybějící HS kódy“."),
}

INCOTERM = "FOB (bez uvedeného přístavu/místa) – viz Poznámky"

HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BOLD = Font(name="Arial", size=10, bold=True)
BASE = Font(name="Arial", size=10)
RED = Font(name="Arial", size=10, bold=True, color="C00000")
TITLE = Font(name="Arial", size=13, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FILL_SENDER = PatternFill("solid", fgColor="DDEBF7")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")
FILL_RED = PatternFill("solid", fgColor="FCE4E4")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.Workbook()

# ------------------------------------------------------------------ Line items
det = wb.active
det.title = "Line items"
cols = ["Sender", "Invoice No.", "Item", "Part Number", "Description of Goods (EN)", "HS kód",
        "Popis zboží (CZ)", "Quantity (pcs)", "Net weight (kg)", "Gross weight (kg)",
        "Customs value (EUR)", "Zdroj HS kódu"]
det.append(cols)
for i in range(1, len(cols) + 1):
    c = det.cell(row=1, column=i)
    c.font, c.fill = HDR, FILL_HDR
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for sender, item, part, en, qty, net, gross, val, hs, cz, src in LINES:
    det.append([sender, INV[sender], item, part, en, hs or MISSING, cz, qty, net, gross, val, src])

last = det.max_row
for r in range(2, last + 1):
    for c in range(1, len(cols) + 1):
        cell = det.cell(row=r, column=c)
        cell.font, cell.border = BASE, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    det.cell(row=r, column=6).number_format = "@"
    det.cell(row=r, column=8).number_format = "#,##0"
    for c in (9, 10, 11):
        det.cell(row=r, column=c).number_format = "#,##0.00"
    if det.cell(row=r, column=6).value == MISSING:
        for c in range(1, len(cols) + 1):
            det.cell(row=r, column=c).fill = FILL_RED
        det.cell(row=r, column=6).font = RED

tr = last + 1
det.cell(row=tr, column=7, value="CELKEM / TOTAL").font = BOLD
for c, L in ((8, "H"), (9, "I"), (10, "J"), (11, "K")):
    cell = det.cell(row=tr, column=c, value=f"=SUM({L}2:{L}{last})")
    cell.font, cell.fill = BOLD, FILL_TOT
    cell.number_format = "#,##0" if c == 8 else "#,##0.00"
for c in range(1, len(cols) + 1):
    det.cell(row=tr, column=c).fill = FILL_TOT
    det.cell(row=tr, column=c).border = BORDER

for col, w in zip("ABCDEFGHIJKL", [36, 14, 7, 22, 22, 24, 46, 13, 13, 14, 15, 62]):
    det.column_dimensions[col].width = w
det.freeze_panes = "A2"
det.auto_filter.ref = f"A1:L{last}"

# ------------------------------------------------------------------ HS Summary
ws = wb.create_sheet("HS Summary", 0)
ws["A1"] = "HS CODE SUMMARY – Sonavox Technology CZ s.r.o."
ws["A1"].font = TITLE
meta = [
    ("Reference MW:", "CZ26011282 (Josef Šrejber – RAIL) | režim: volný oběh | celnice: A1 Ostrava | příjezd cca 34. týden"),
    ("Odesílatelé / Senders:", "3 – zásilka je členěna PRIMÁRNĚ dle odesílatele"),
    ("Příjemce / Consignee:", "Sonavox Technology CZ s.r.o., Lhotka nad Bečvou 93, 756 41 Lešná, Czech Republic"),
    ("FCR / HBL No.:", "TECGO26080029 (T.H.I. Group Ltd.), ref. ZIHWB260805FM022"),
    ("Nákladní list SMGS:", "38055706"),
    ("Kontejner / plomba:", "CIMU1716969 / 40HQ / plomba ZH176441"),
    ("Faktury:", "CZ20260725, CZ20260725-1, CZ20260725-2, vše ze dne 28.07.2026"),
    ("Trasa:", "Zhengzhou (Putian), Henan → Malaszewicze (PL) → Czech Republic – ŽELEZNIČNÍ přeprava"),
    ("Země původu:", "Čína (Suzhou, 35029)"),
    ("Měna:", "EUR"),
    ("INCOTERM:", INCOTERM),
    ("Preferenční věta:", "NENALEZENA v žádném dokladu – viz list Poznámky"),
    ("Celkem kolí / objem:", "51 palet / 51,40 CBM / brutto 24 133 kg (souhlasí s FCR i SMGS)"),
    ("Doplněné HS kódy:", "2 položky doplnil klient 12.08.2026 – viz list „Doplněné HS kódy“ "
                          "(u Spring nut jen 6 míst, nutno doplnit)"),
    ("CBAM:", "Prohlášení de minimis Y137 (do 50 t/rok) ze dne 12.08.2026, podepsala Ing. Michaela Ryan"),
    ("ANTIDUMPING:", "NEVYŘEŠEN – prohlášení Y137 se týká CBAM, nikoli antidumpingu. Viz Poznámky."),
]
r = 3
for k, v in meta:
    ws.cell(row=r, column=1, value=k).font = BOLD
    ws.cell(row=r, column=2, value=v).font = BASE
    if k.startswith(("Preferenční", "INCOTERM", "CHYBĚJÍCÍ")):
        for c in (1, 2):
            ws.cell(row=r, column=c).fill = FILL_WARN
        ws.cell(row=r, column=2).font = RED
    r += 1

r += 1
scols = ["Odesílatel / Sender", "Faktura", "HS kód", "Popis zboží (CZ)", "Množství (ks)",
         "Čistá hmotnost (kg)", "Hrubá hmotnost (kg)", "Celní hodnota (EUR)",
         "Incoterm", "CBAM", "Antidumping", "Zdroj HS kódu / poznámka"]
hdr = r
for i, c in enumerate(scols, start=1):
    cell = ws.cell(row=hdr, column=i, value=c)
    cell.font, cell.fill = HDR, FILL_HDR
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws.row_dimensions[hdr].height = 32

row = hdr + 1
blocks = []
for sender in (S1, S2, S3):
    keys = []
    for ln in LINES:
        if ln[0] != sender:
            continue
        # položky bez kódu se NEslučují – každá zůstává samostatně
        k = ln[8] if ln[8] else ("__MISSING__", ln[2])
        if k not in keys:
            keys.append(k)
    start = row
    for k in keys:
        if isinstance(k, tuple):
            members = [ln for ln in LINES if ln[0] == sender and ln[8] is None and ln[2] == k[1]]
            hs_disp, hs_key = MISSING, None
        else:
            members = [ln for ln in LINES if ln[0] == sender and ln[8] == k]
            hs_disp, hs_key = k, k
        czs = []
        for m in members:
            if m[9] not in czs:
                czs.append(m[9])
        srcs = []
        for m in members:
            if m[10] not in srcs:
                srcs.append(m[10])
        cbam, ad, note = HS_META[hs_key]
        ws.cell(row=row, column=1, value=sender)
        ws.cell(row=row, column=2, value=INV[sender])
        ws.cell(row=row, column=3, value=hs_disp).number_format = "@"
        ws.cell(row=row, column=4, value="; ".join(czs))
        for ci, sc in ((5, "H"), (6, "I"), (7, "J"), (8, "K")):
            if hs_key:
                f = (f"=SUMIFS('Line items'!{sc}$2:{sc}${last},'Line items'!$A$2:$A${last},$A{row},"
                     f"'Line items'!$F$2:$F${last},$C{row})")
            else:  # missing-code rows keyed by part number instead
                f = (f"=SUMIFS('Line items'!{sc}$2:{sc}${last},'Line items'!$A$2:$A${last},$A{row},"
                     f"'Line items'!$D$2:$D${last},\"{k[1]}\")")
            ws.cell(row=row, column=ci, value=f)
        ws.cell(row=row, column=9, value=INCOTERM)
        ws.cell(row=row, column=10, value=cbam)
        ws.cell(row=row, column=11, value=ad)
        ws.cell(row=row, column=12, value=note if hs_key else srcs[0])
        for c in range(1, 13):
            cell = ws.cell(row=row, column=c)
            cell.font, cell.border = BASE, BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=5).number_format = "#,##0"
        for c in (6, 7, 8):
            ws.cell(row=row, column=c).number_format = "#,##0.00"
        if hs_key is None or cbam == "ANO" or ad in ("ANO", "K OVĚŘENÍ"):
            for c in range(1, 13):
                ws.cell(row=row, column=c).fill = FILL_RED
            ws.cell(row=row, column=3).font = RED
            for c in (10, 11):
                ws.cell(row=row, column=c).font = RED
        row += 1
    end = row - 1
    ws.cell(row=row, column=4, value=f"Mezisoučet – {sender}").font = BOLD
    for c, L in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
        cell = ws.cell(row=row, column=c, value=f"=SUM({L}{start}:{L}{end})")
        cell.font = BOLD
        cell.number_format = "#,##0" if c == 5 else "#,##0.00"
    for c in range(1, 13):
        ws.cell(row=row, column=c).fill = FILL_SENDER
        ws.cell(row=row, column=c).border = BORDER
    blocks.append(row)
    row += 1

gt = row
ws.cell(row=gt, column=4, value="CELKEM ZÁSILKA / TOTAL SHIPMENT").font = BOLD
for c, L in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
    cell = ws.cell(row=gt, column=c, value="=" + "+".join(f"{L}{b}" for b in blocks))
    cell.font = BOLD
    cell.number_format = "#,##0" if c == 5 else "#,##0.00"
for c in range(1, 13):
    ws.cell(row=gt, column=c).fill = FILL_TOT
    ws.cell(row=gt, column=c).border = BORDER

cc = gt + 2
ws.cell(row=cc, column=1, value="Kontrola / cross-check").font = BOLD
ws.cell(row=cc, column=2, value="Ze souhrnu").font = BOLD
ws.cell(row=cc, column=3, value="Z dokladů").font = BOLD
ws.cell(row=cc, column=4, value="Rozdíl").font = BOLD
checks = [
    ("Množství celkem vs. součet faktur", f"=E{gt}", 1324160 + 21168 + 43760),
    ("Čistá hmotnost celkem (kg)", f"=F{gt}", 9574.05 + 12448.80 + 957.00),
    ("Hrubá hmotnost celkem (kg) vs. FCR/SMGS 24 133 kg", f"=G{gt}", 24133.00),
    ("Celní hodnota celkem (EUR) vs. součet faktur", f"=H{gt}", 73481.56 + 25058.88 + 20830.64),
]
cr = cc + 1
for label, ref, doc in checks:
    ws.cell(row=cr, column=1, value=label).font = BASE
    ws.cell(row=cr, column=2, value=ref).font = BASE
    ws.cell(row=cr, column=3, value=doc).font = BASE
    ws.cell(row=cr, column=4, value=f"=B{cr}-C{cr}").font = BASE
    for c in (2, 3, 4):
        ws.cell(row=cr, column=c).number_format = "#,##0.00"
    cr += 1

for col, w in zip("ABCDEFGHIJKL", [38, 14, 24, 52, 13, 15, 15, 16, 32, 11, 13, 62]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = f"A{hdr + 1}"

# --------------------------------------------------------- Chybějící HS kódy
mi = wb.create_sheet("Doplněné HS kódy", 1)
mi["A1"] = "HS KÓDY DODATEČNĚ DOPLNĚNÉ KLIENTEM"
mi["A1"].font = TITLE
mi["A2"] = ("Tyto dvě položky neměly HS kód v žádném původním podkladu. Kódy dodala Alice Knápková "
            "(Sonavox Technology CZ) e-mailem dne 12.08.2026 – viz sloupec Zdroj.")
mi["A2"].font = BASE
mi["A2"].alignment = Alignment(wrap_text=True, vertical="top")
mi.merge_cells("A2:J2")
mi.row_dimensions[2].height = 30

RESOLVED = [
    (S1, 9, "J2-10X002-0001-CZ", "Waterproof membrane", 210000, 30.24, 40.00, 8820.00,
     "8518900000", "OK – úplný 10místný kód",
     "e-mail Alice Knápková, 12.08.2026: „Waterproof membrane  8518900000“"),
    (S1, 15, "029A-000004-CZ", "Spring nut", 30000, 12.6, 15.00, 540.00,
     "731816", "NEÚPLNÝ – jen 6 míst, nutno doplnit na 10",
     "e-mail Alice Knápková, 12.08.2026: „029A-000004-CZ  HS 731816“"),
]

mcols = ["Odesílatel", "Faktura", "Pol. č.", "Part Number", "Popis dle faktury (EN)",
         "Množství (ks)", "Čistá hm. (kg)", "Hrubá hm. (kg)", "Celní hodnota (EUR)",
         "HS kód od klienta", "Stav kódu", "Zdroj"]
r = 4
for i, c in enumerate(mcols, start=1):
    cell = mi.cell(row=r, column=i, value=c)
    cell.font, cell.fill = HDR, FILL_HDR
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
mi.row_dimensions[r].height = 32
r += 1
for sender, item, part, en, qty, net, gross, val, hs, stav, zdroj in RESOLVED:
    mi.append([sender, INV[sender], item, part, en, qty, net, gross, val, hs, stav, zdroj])
    for c in range(1, 13):
        cell = mi.cell(row=r, column=c)
        cell.font, cell.border = BASE, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    mi.cell(row=r, column=6).number_format = "#,##0"
    for c in (7, 8, 9):
        mi.cell(row=r, column=c).number_format = "#,##0.00"
    mi.cell(row=r, column=10).number_format = "@"
    if stav.startswith("NEÚPLNÝ"):
        for c in range(1, 13):
            mi.cell(row=r, column=c).fill = FILL_RED
        mi.cell(row=r, column=11).font = RED
    r += 1
for col, w in zip("ABCDEFGHIJKL", [34, 13, 8, 22, 24, 12, 13, 13, 15, 16, 34, 52]):
    mi.column_dimensions[col].width = w

# ------------------------------------------------------------------ Poznámky
nt = wb.create_sheet("Poznámky & flagy")
nt["A1"] = "POZNÁMKY, UPOZORNĚNÍ A NESROVNALOSTI"
nt["A1"].font = TITLE
notes = [
    ("KRITICKÉ", "Prohlášení Y137 řeší CBAM, NIKOLI antidumping",
     "Na dotaz k antidumpingu u položky 7318 15 95 19 zaslal klient dne 12.08.2026 prohlášení podepsané "
     "Ing. Michaelou Ryan: „Sonavox Technology CZ s.r.o. uplatňuje výjimku de minimis – za kalendářní rok "
     "naše dovozy nepřesáhnou kumulativně 50 tun čisté hmotnosti (kód Y137)“. Kód Y137 je však kód CBAM "
     "(uhlíkové vyrovnávací opatření), NE antidumping. Antidumpingové clo dle nař. (EU) 2022/191 žádnou "
     "obdobnou výjimku de minimis nemá. Clo 274 584 Kč tedy tímto prohlášením NEODPADÁ. "
     "NUTNO KLIENTOVI VYSVĚTLIT A ZNOVU VYŽÁDAT NÁZEV A ADRESU ČÍNSKÉHO VÝROBCE."),
    ("KRITICKÉ", "Rozpor 7318 16 vs. 7318 15 95 – možná úspora 274 584 Kč",
     "Klient dne 12.08.2026 zařadil Spring nut (029A-000004-CZ) pod HS 7318 16, tedy pod MATICE. "
     "Položka 14 „Nut“ (029-2513-AL-CZ) je přitom dle tabulky SONAVOX pod 7318 15 95 19, tedy pod "
     "ŠROUBY A SVORNÍKY. Pokud je i díl 029-2513-AL-CZ maticí, patří rovněž pod 7318 16 – a ten v seznamu "
     "nař. (EU) 2022/191 NENÍ, takže by antidumping zcela odpadl. Rozdíl činí 274 584 Kč. "
     "NUTNO S KLIENTEM VYJASNIT, PROČ JSOU DVA PODOBNÉ DÍLY V RŮZNÝCH POLOŽKÁCH."),
    ("K OVĚŘENÍ", "Spring nut – kód je jen 6místný",
     "Klient uvedl „HS 731816“, tedy pouze 6 míst. Do celního prohlášení je nutný 10místný kód; "
     "podpoložka 7318 16 se dále dělí podle typu matice a vnitřního průměru. Nutno vyžádat doplnění."),
    ("VYŘEŠENO", "Waterproof membrane – kód doplněn klientem",
     "Klient dne 12.08.2026 zařadil J2-10X002-0001-CZ (210 000 ks, 8 820,00 EUR) pod 8518 90 00 00, "
     "tedy jako část reproduktoru. Kód je úplný a odpovídá DK 00 používanému u ostatních částí. "
     "Položka je proto sloučena do řádku 8518 90 00 00."),
    ("KRITICKÉ", "Preferenční věta / preferenční původ (COO)",
     "V žádném z dodaných dokladů (3× faktura, 3× packing list, FCR TECGO26080029, nákladní list SMGS 38055706) "
     "NENÍ uvedena preferenční věta o původu zboží. Původ zboží = ČÍNA. EU nemá s Čínou dohodu o preferenčním "
     "obchodu, preferenční sazební zacházení tedy NELZE uplatnit a použije se plná celní sazba třetí země."),
    ("KRITICKÉ", "Incoterm – chybí místo a termín neodpovídá druhu přepravy",
     "Všechny tři faktury i packing listy uvádějí pouze „Terms of Payment: FOB“ – BEZ přístavu / jmenovaného "
     "místa. Navíc FOB je dle Incoterms 2020 určen jen pro námořní a vnitrozemskou vodní přepravu, ale zásilka "
     "jde ŽELEZNICÍ (Zhengzhou/Putian → Malaszewicze → CZ); správný termín je FCA s uvedeným místem. "
     "FCR je „freight collect“, takže bez určení dodacího místa nelze spolehlivě určit, zda se do celní "
     "hodnoty připočítává dopravné do hranice EU."),
    ("KRITICKÉ", "Antidumping u matic – vyčísleno 274 584 Kč",
     "Položka 14 (Nut, 029-2513-AL-CZ, 180 000 ks, 12 600,00 EUR) pod kódem 7318 15 95 19. Výpočet v Helios: "
     "druh A00, základ 304 416 Kč, sazba 90,2 % (= 3,7 % clo + 86,5 % antidumping), částka 274 584 Kč. "
     "Sazba 86,5 % je ZBYTKOVÁ pro neznámého/nespolupracujícího výrobce; s názvem a adresou čínského výrobce "
     "lze získat doplňkový kód TARIC se sazbou 22,1–48,8 % (clo celkem 25,8–52,5 %, tj. 78 539–159 818 Kč) "
     "nebo 39,6 % pro spolupracující výrobce (131 812 Kč). Možná úspora až cca 196 000 Kč. "
     "Tři možné scénáře: (a) HLINÍK → 7616 10 00, clo cca 6 %, BEZ antidumpingu; (b) OCELOVÁ MATICE → 7318 16, "
     "clo cca 3,7 %, BEZ antidumpingu; (c) OCELOVÝ ŠROUB / VLOŽKA → ex 7318 15 95, clo 3,7 % + ANTIDUMPING "
     "22,1–86,5 %. Nař. (EU) 2022/191 se vztahuje na vruty, samořezné šrouby, ostatní šrouby a svorníky s hlavou "
     "a podložky (7318 12 90, 7318 14 91, 7318 14 99, 7318 15 58, 7318 15 68, 7318 15 82, 7318 15 88, "
     "ex 7318 15 95, ex 7318 21 00, ex 7318 22 00) – kód 7318 16 pro MATICE v seznamu NENÍ. "
     "Bez jména čínského výrobce se uplatní zbytková sazba 86,5 %. NUTNO ZJISTIT MATERIÁL, DRUH DÍLU A VÝROBCE."),
    ("VYŘEŠENO", "CBAM – prohlášení de minimis Y137 doloženo",
     "Položky 14 a 15 (7318 15 95 19 a 7318 16, celkem 312,30 kg čisté hmotnosti) spadají do přílohy I "
     "nařízení CBAM. Klient doložil dne 12.08.2026 podepsané prohlášení, že dovozy Sonavox Technology CZ s.r.o. "
     "nepřesáhnou za kalendářní rok kumulativně 50 tun čisté hmotnosti – kód Y137 (do 50 t ročně). "
     "V celním prohlášení se tedy u těchto položek uvede Y137. "
     "POZOR: platnost je vázána na kumulativní roční objem – při dalších dovozech je nutno sledovat součet. "
     "Ostatní položky zásilky (kap. 39, 85) do CBAM nespadají."),
    ("K OVĚŘENÍ", "CBAM registr příjemců nebyl k dispozici",
     "Referenční soubor reference/CBAM_receivers_CZ.xlsx (list List1) není v repozitáři k dispozici, "
     "nebylo proto možné ověřit zápis příjemce Sonavox Technology CZ s.r.o. (IČO 19670338, DIČ CZ19670338) "
     "v registru ani porovnat deklarovaný kód Y137 se záznamem. Doporučeno ověřit a případně příjemce "
     "do registru doplnit."),
    ("ROZPOR", "Front plate – e-mail a celní prohlášení se liší",
     "E-mail SONAVOX uvádí u „FRONT PLATE / ocelová deska“ kód 8518900090. Předchozí celní prohlášení Sonavox "
     "však deklaruje „Front plate – deska“ pod CN 8518 90 00 s DK 00. V souhrnu je použita varianta "
     "z celního prohlášení (8518 9000 00), protože jde o skutečně přijaté prohlášení. Přípony 40 / 90 / 99, "
     "které e-mail u 8518 90 00 používá, se v žádném z prohlášení nevyskytují – vždy je tam DK 00. "
     "POKUD MÁ PLATIT VARIANTA Z E-MAILU, JE NUTNO ŘÁDEK OPRAVIT."),
    ("ROZPOR", "Terminal – e-mail a celní prohlášení se liší",
     "E-mail SONAVOX nemá řádek „Terminal“; nejbližší je „Pin, lug / kontakt – 74093900“. Předchozí celní "
     "prohlášení však deklaruje přímo „terminal – části reproduktorů“ pod CN 8518 90 00 DK 00. V souhrnu je "
     "použit kód z celního prohlášení, protože odpovídá názvu položky doslova."),
    ("NESROVNALOST", "FCR uvádí jen jednoho odesílatele, zásilka je od tří prodávajících",
     "FCR TECGO26080029 uvádí jako shipper pouze SUZHOU SONAVOX ELECTRONICS CO., LTD, zboží však pochází "
     "od tří různých prodávajících dle tří samostatných faktur. Souhrn je proto členěn dle odesílatele."),
    ("NESROVNALOST", "Popis zboží ve FCR neuvádí všechny položky",
     "FCR vyjmenovává: Speaker component, Speaker, Voice coil, Front plate, Gasket, Waterproof membrane, "
     "Terminal, Nut, Spring nut, Harness, Gauge, Cone. CHYBÍ SPIDER (77 600 ks) a CERAMICS CIRCLE (19 152 ks) – "
     "zřejmě zahrnuty pod obecné „Speaker component“."),
    ("NESROVNALOST", "Nákladní list SMGS deklaruje celý kontejner pod jedním kódem",
     "SMGS 38055706 uvádí HS 851890 / 85189000 „Speaker component“ (+ HS 99020000) pro celý kontejner, "
     "ačkoli zboží spadá do kapitol 39, 73 a 85. Přepravní zjednodušení bez vlivu na celní zařazení, "
     "může však vyvolat dotaz celního úřadu."),
    ("NESROVNALOST", "Rozdílná čísla bankovních účtů (CZ20260725)",
     "Packing list uvádí účet EUR 500158216108, faktura EUR 468960663282. Bez vlivu na celní řízení, "
     "ale před úhradou ověřit u dodavatele."),
    ("NESROVNALOST", "ETD/ETA neodpovídá datu vypravení",
     "Packing list CZ20260725-2 uvádí ETD 29.07.2026 / ETA 29.08.2026; FCR byl vystaven 04.08.2026 "
     "a vlak má č. V.WB2026/08/05. U zbylých dvou sad dokladů jsou ETD/ETA prázdné."),
    ("OK", "Faktury a packing listy spolu souhlasí",
     "Porovnáno po položkách: 19 / 2 / 3 řádků v obou dokladech každé sady, shodná čísla dílů, množství "
     "i popisy. Žádná položka není jen na faktuře nebo jen na packing listu."),
    ("OK", "Kontrolní součty souhlasí",
     "Hrubá hmotnost 9 950,74 + 12 907,84 + 1 274,42 = 24 133,00 kg = FCR i SMGS. Kolí 12 + 23 + 16 = 51 palet. "
     "Objem 15,24 + 16,64 + 19,52 = 51,40 CBM = FCR. Množství i celní hodnota souhlasí s fakturami."),
    ("OK", "Sleva / rabat",
     "V žádné faktuře není uvedena obchodní sleva ani rabat; součet položek = Sub Total. "
     "Poměrné rozpuštění slevy do celní hodnoty se neuplatňuje."),
    ("INFO", "Kódy nebyly nikdy doplňovány vlastním odvozením",
     "Každý HS kód v tomto souhrnu je doslovně převzat z e-mailu „SONAVOX HS KODY“ nebo z předchozího celního "
     "prohlášení Sonavox; zdroj je uveden u každého řádku ve sloupci „Zdroj HS kódu“ a na listu Line items. "
     "Kde podklad chybí, zůstává kód prázdný. Platnost kódů nebyla ověřena přímo v databázi TARIC – "
     "doporučeno před podáním prohlášení zkontrolovat."),
]
r = 3
for i, h in enumerate(["Priorita", "Téma", "Popis"], start=1):
    cell = nt.cell(row=r, column=i, value=h)
    cell.font, cell.fill = HDR, FILL_HDR
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = BORDER
r += 1
for prio, topic, text in notes:
    nt.cell(row=r, column=1, value=prio)
    nt.cell(row=r, column=2, value=topic)
    nt.cell(row=r, column=3, value=text)
    for c in range(1, 4):
        cell = nt.cell(row=r, column=c)
        cell.font, cell.border = BASE, BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    if prio == "KRITICKÉ":
        for c in range(1, 4):
            nt.cell(row=r, column=c).fill = FILL_RED
        nt.cell(row=r, column=1).font = RED
        nt.cell(row=r, column=2).font = BOLD
    elif prio in ("ROZPOR", "NESROVNALOST"):
        for c in range(1, 4):
            nt.cell(row=r, column=c).fill = FILL_WARN
        nt.cell(row=r, column=2).font = BOLD
    nt.row_dimensions[r].height = 62
    r += 1
nt.column_dimensions["A"].width = 15
nt.column_dimensions["B"].width = 46
nt.column_dimensions["C"].width = 118
nt.freeze_panes = "A4"

# ------------------------------------------------------------------ Zdroje
src = wb.create_sheet("Zdroje")
src["A1"] = "ZDROJOVÉ DOKLADY"
src["A1"].font = TITLE
rows = [
    ("Doklad", "Obsah / použití"),
    ("TRCZ20260725.xlsx (PL, INV)", "Faktura + packing list CZ20260725 – Suzhou Sonavox International Trading (19 položek)"),
    ("ELCZ202607251_...GCS.xlsx (2 soubory)", "Faktura + packing list CZ20260725-1 – Suzhou Sonavox Electronics (2 položky)"),
    ("YLCZ202607252.xlsx (PL, INV)", "Faktura + packing list CZ20260725-2 – Suzhou Yanlong Electronic Product (3 položky)"),
    ("HBL_TECGO26080029.pdf", "Forwarder's Cargo Receipt T.H.I. Group – trasa, kontejner, brutto, kolí, objem"),
    ("CIMU1716969.pdf", "Nákladní list SMGS 38055706 – železniční přeprava, plomba, hmotnosti, trasa"),
    ("Re_SONAVOX_HS_KODY.eml", "Referenční tabulka HS kódů s českými popisy (Ing. Michaela Ryan, Sonavox Technology CZ)"),
    ("Předchozí celní prohlášení Sonavox (2×)", "Položky dřívějších prohlášení na obdobné zásilky – zdroj kódů "
     "8518 90 00 DK 00, 8518 21 00 DK 00, 7413 00 00 DK 90, 3926 90 97 DK 90"),
    ("E-mail Josef Šrejber, OR: CZ26011282", "Zadání k vyclení – reference, celnice A1 Ostrava, režim volný oběh, příjezd 34. týden"),
    ("", ""),
    ("PRAVIDLO", "HS kódy se v tomto souhrnu NEVYTVÁŘEJÍ vlastním odvozením. Použit je pouze kód doslovně "
                 "uvedený v některém z výše uvedených podkladů; zdroj je zaznamenán u každého řádku. "
                 "Položky bez podkladu zůstávají bez kódu a vyžadují se od klienta – list „Chybějící HS kódy“."),
]
r = 3
for a, b in rows:
    src.cell(row=r, column=1, value=a).font = BOLD if a in ("Doklad", "PRAVIDLO") else BASE
    src.cell(row=r, column=2, value=b).font = BOLD if a == "Doklad" else BASE
    src.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    if a == "PRAVIDLO":
        for c in (1, 2):
            src.cell(row=r, column=c).fill = FILL_RED
        src.cell(row=r, column=1).font = RED
        src.row_dimensions[r].height = 46
    r += 1
src.column_dimensions["A"].width = 42
src.column_dimensions["B"].width = 112

out = "/home/user/customs/output/HS_Code_Summary_Sonavox_Technology_CZ_TECGO26080029.xlsx"
wb.save(out)
print("saved", out)
