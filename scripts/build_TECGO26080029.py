#!/usr/bin/env python3
"""Build HS code summary for Sonavox shipment TECGO26080029 / container CIMU1716969."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

S1 = "Suzhou Sonavox International Trading Co., Ltd"
S2 = "Suzhou Sonavox Electronics Co., Ltd"
S3 = "Suzhou Yanlong Electronic Product Co., Ltd"

INV = {S1: "CZ20260725", S2: "CZ20260725-1", S3: "CZ20260725-2"}

# sender, invoice item no, part number, EN description, qty, net, gross, EUR, HS code, CZ description
LINES = [
    # --- Sender 1: Suzhou Sonavox International Trading Co., Ltd (CZ20260725) ---
    (S1, 1, "010-KD1622S-3.8A-CZ", "Voice coil", 65000, 86.5, 216.33, 9841.00, "8504500000",
     "Kmitací cívka – měděný drát (jednoduché vinutí) navinutý na papírovém válci (formeru)"),
    (S1, 2, "040-62E4-CR-CZ", "Front plate", 35000, 2345.0, 2385.00, 3542.00, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru"),
    (S1, 3, "040-80I6-CRB-CZ", "Front plate", 32400, 5103.0, 5183.00, 7464.96, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru"),
    (S1, 4, "040-54D3-CR-1-CZ", "Front plate", 25000, 1012.5, 1032.91, 1377.50, "8518900000",
     "Ocelová deska (front plate) – součást autoreproduktoru"),
    (S1, 5, "021-100F00-9-CZ", "Spider", 5600, 10.5, 15.00, 449.68, "8518900000",
     "Středící membrána (spider) – naimpregnovaná textilie, součást autoreproduktoru"),
    (S1, 6, "021-80E15-3-CZ", "Spider", 72000, 19.4, 25.00, 7041.60, "8518900000",
     "Středící membrána (spider) – naimpregnovaná textilie, součást autoreproduktoru"),
    (S1, 19, "140-35D14-44XP1-CZ", "Cone", 8000, 16.0, 20.00, 2926.40, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru"),
    (S1, 7, "142-630H10-15H-CZ", "Gasket", 960, 7.2, 10.00, 100.80, "3926909790",
     "Plastové těsnění (gasket) – ostatní výrobky z plastů"),
    (S1, 8, "142-40X25-25D-CZ", "Gasket", 2000, 15.0, 20.00, 75.80, "3926909790",
     "Plastové těsnění (gasket) – ostatní výrobky z plastů"),
    (S1, 17, "JK-65425-01", "Gauge", 5000, 57.78, 70.00, 899.50, "3926909790",
     "Plastový kalibr (gauge) – montážní přípravek, ostatní výrobky z plastů"),
    (S1, 18, "JK-35113-01-CZ", "Gauge", 2000, 23.11, 28.00, 302.00, "3926909790",
     "Plastový kalibr (gauge) – montážní přípravek, ostatní výrobky z plastů"),
    (S1, 9, "J2-10X002-0001-CZ", "Waterproof membrane", 210000, 30.24, 40.00, 8820.00, "3921909090",
     "Plastová membrána (vodotěsná membrána) – ochranná membrána reproduktoru"),
    (S1, 10, "181-35113-1SN-CZ", "Terminal", 75000, 37.5, 40.00, 1612.50, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru"),
    (S1, 11, "181-35113-2SN-CZ", "Terminal", 125000, 62.5, 70.00, 2687.50, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru"),
    (S1, 12, "181-40133-1SN-CZ", "Terminal", 200000, 100.0, 110.00, 2540.00, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru"),
    (S1, 13, "181-40133-2SN-CZ", "Terminal", 200000, 100.0, 110.00, 2540.00, "8518900000",
     "Kontakt / vývod (terminal) – část autoreproduktoru"),
    (S1, 14, "029-2513-AL-CZ", "Nut", 180000, 299.7, 310.50, 12600.00, "7318159519",
     "Matice / vložka (nut, insert) – závitové zboží ze železa nebo oceli"),
    (S1, 15, "029A-000004-CZ", "Spring nut", 30000, 12.6, 15.00, 540.00, "7318159519",
     "Matice / vložka (nut, insert) – závitové zboží ze železa nebo oceli"),
    (S1, 16, "008-XH-2Y-DK-6-CZ", "Harness", 51200, 235.52, 250.00, 8120.32, "7413000090",
     "Kabelový svazek – splétané lanko, slitina měď (78 %) a cín, neizolované"),

    # --- Sender 2: Suzhou Sonavox Electronics Co., Ltd (CZ20260725-1) ---
    (S2, 1, "212-8075-CZ", "Ceramics circle", 19152, 11299.68, 11688.80, 19918.08, "8518900000",
     "Magnetický obvod (ceramics circle) – feritové jádro (nezmagnetované), součást autoreproduktoru"),
    (S2, 2, "IM-50196-AUDICZ", "Car Speaker", 2016, 1149.12, 1219.04, 5140.80, "8518210000",
     "Reproduktor (autoreproduktor), jediný v jedné skříni"),

    # --- Sender 3: Suzhou Yanlong Electronic Product Co., Ltd (CZ20260725-2) ---
    (S3, 1, "140-80I27-42X-CZ", "Cone", 33600, 856.8, 1134.00, 16836.96, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru"),
    (S3, 2, "140-80E30-58X-CZ", "Cone", 2160, 52.2, 84.12, 1280.88, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru"),
    (S3, 3, "140-40E17-72XC-CZ", "Cone", 8000, 48.0, 56.30, 2712.80, "8518900000",
     "Membrána s vlnkou (cone) – membrána elektrodynamického reproduktoru"),
]

# HS code -> (CBAM, antidumping, classification note)
HS_META = {
    "8504500000": ("Ne", "Ne", "Zařazení dle ZISZ 30-0292-2016 (kmitací cívka = ostatní tlumivky)."),
    "8518900000": ("Ne", "Ne", "Části a součásti reproduktorů. Přípona DK 00 a zařazení terminalu i front plate "
                                "pod tento kód dle předchozího celního prohlášení Sonavox (viz list Zdroje)."),
    "8518210000": ("Ne", "Ne", "Kompletní autoreproduktor, jediný v jedné skříni."),
    "3926909790": ("Ne", "Ne", "Ostatní výrobky z plastů. Gasket + Gauge na jednom řádku, DK 90 – potvrzeno "
                                "předchozím celním prohlášením Sonavox."),
    "3921909090": ("Ne", "Ne", "Ostatní desky, listy, fólie a pásy z plastů."),
    "7318159519": ("ANO", "ANO", "Spojovací materiál ze železa/oceli, původ ČÍNA – viz list Poznámky."),
    "7413000090": ("Ne", "Ne", "Splétaná lanka z mědi, neizolovaná."),
}

INCOTERM = "FOB (bez uvedeného přístavu/místa) – viz Poznámky"

HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BOLD = Font(name="Arial", size=10, bold=True)
BASE = Font(name="Arial", size=10)
TITLE = Font(name="Arial", size=13, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FILL_SENDER = PatternFill("solid", fgColor="DDEBF7")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")
FILL_RED = PatternFill("solid", fgColor="FCE4E4")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.Workbook()

# ---------------------------------------------------------------- Line items
det = wb.active
det.title = "Line items"
det_cols = ["Sender", "Invoice No.", "Item", "Part Number", "Description of Goods (EN)",
            "HS code", "Popis zboží (CZ)", "Quantity (pcs)", "Net weight (kg)",
            "Gross weight (kg)", "Customs value (EUR)"]
det.append(det_cols)
for i, c in enumerate(det.iter_cols(min_row=1, max_row=1), start=1):
    cell = det.cell(row=1, column=i)
    cell.font = HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for ln in LINES:
    sender, item, part, en, qty, net, gross, val, hs, cz = ln
    det.append([sender, INV[sender], item, part, en, hs, cz, qty, net, gross, val])

det_last = det.max_row
for r in range(2, det_last + 1):
    for c in range(1, len(det_cols) + 1):
        cell = det.cell(row=r, column=c)
        cell.font = BASE
        cell.border = BORDER
    det.cell(row=r, column=6).alignment = Alignment(horizontal="left")
    det.cell(row=r, column=6).number_format = "@"
    det.cell(row=r, column=8).number_format = "#,##0"
    for c in (9, 10):
        det.cell(row=r, column=c).number_format = "#,##0.00"
    det.cell(row=r, column=11).number_format = "#,##0.00"
    det.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical="top")

tr = det_last + 1
det.cell(row=tr, column=7, value="CELKEM / TOTAL").font = BOLD
for c, col in ((8, "H"), (9, "I"), (10, "J"), (11, "K")):
    cell = det.cell(row=tr, column=c, value=f"=SUM({col}2:{col}{det_last})")
    cell.font = BOLD
    cell.fill = FILL_TOT
    cell.number_format = "#,##0" if c == 8 else "#,##0.00"
for c in range(1, len(det_cols) + 1):
    det.cell(row=tr, column=c).fill = FILL_TOT
    det.cell(row=tr, column=c).border = BORDER

for col, w in zip("ABCDEFGHIJK", [38, 14, 7, 22, 22, 13, 52, 13, 13, 14, 16]):
    det.column_dimensions[col].width = w
det.freeze_panes = "A2"
det.auto_filter.ref = f"A1:K{det_last}"

# ---------------------------------------------------------------- HS Summary
ws = wb.create_sheet("HS Summary", 0)
ws["A1"] = "HS CODE SUMMARY – Sonavox Technology CZ s.r.o."
ws["A1"].font = TITLE
meta = [
    ("Odesílatelé / Senders:", "3 (viz členění níže) – zásilka je členěna PRIMÁRNĚ dle odesílatele"),
    ("Příjemce / Consignee:", "Sonavox Technology CZ s.r.o., Lhotka nad Bečvou 93, 756 41 Lešná, Czech Republic"),
    ("FCR / HBL No.:", "TECGO26080029 (T.H.I. Group Ltd.), ref. ZIHWB260805FM022"),
    ("Nákladní list SMGS:", "38055706"),
    ("Kontejner / plomba:", "CIMU1716969 / 40HQ / plomba ZH176441"),
    ("Faktury:", "CZ20260725, CZ20260725-1, CZ20260725-2, vše ze dne 28.07.2026"),
    ("Trasa:", "Zhengzhou (Putian), Henan → Malaszewicze (PL) → Czech Republic – ŽELEZNIČNÍ přeprava"),
    ("Země původu:", "Čína (Suzhou, 35029)"),
    ("Destinace:", "Czech Republic / EU"),
    ("Měna:", "EUR"),
    ("INCOTERM:", INCOTERM),
    ("Preferenční věta:", "NENALEZENA v žádném dokladu – viz list Poznámky (kritické)"),
    ("Celkem kolí / objem:", "51 palet / 51,40 CBM / brutto 24 133 kg (souhlasí s FCR i SMGS)"),
]
r = 3
for k, v in meta:
    ws.cell(row=r, column=1, value=k).font = BOLD
    ws.cell(row=r, column=2, value=v).font = BASE
    if "Preferenční" in k or "INCOTERM" in k:
        ws.cell(row=r, column=1).fill = FILL_WARN
        ws.cell(row=r, column=2).fill = FILL_WARN
        ws.cell(row=r, column=2).font = Font(name="Arial", size=10, bold=True, color="C00000")
    r += 1

r += 1
cols = ["Odesílatel / Sender", "Faktura", "HS kód", "Popis zboží (CZ)", "Množství (ks)",
        "Čistá hmotnost (kg)", "Hrubá hmotnost (kg)", "Celní hodnota (EUR)",
        "Incoterm", "CBAM", "Antidumping", "Poznámka"]
hdr_row = r
for i, c in enumerate(cols, start=1):
    cell = ws.cell(row=hdr_row, column=i, value=c)
    cell.font = HDR
    cell.fill = FILL_HDR
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws.row_dimensions[hdr_row].height = 32

# group: sender order S1, S2, S3; within sender keep first-appearance HS order
row = hdr_row + 1
sender_blocks = []
for sender in (S1, S2, S3):
    hs_order = []
    for ln in LINES:
        if ln[0] == sender and ln[8] not in hs_order:
            hs_order.append(ln[8])
    start = row
    for hs in hs_order:
        members = [ln for ln in LINES if ln[0] == sender and ln[8] == hs]
        czs = []
        for m in members:
            if m[9] not in czs:
                czs.append(m[9])
        cz = "; ".join(czs)
        cbam, ad, note = HS_META[hs]
        ws.cell(row=row, column=1, value=sender)
        ws.cell(row=row, column=2, value=INV[sender])
        hc = ws.cell(row=row, column=3, value=hs)
        hc.number_format = "@"
        ws.cell(row=row, column=4, value=cz)
        for col_i, src_col in ((5, "H"), (6, "I"), (7, "J"), (8, "K")):
            ws.cell(row=row, column=col_i,
                    value=f"=SUMIFS('Line items'!{src_col}$2:{src_col}${det_last},"
                          f"'Line items'!$A$2:$A${det_last},$A{row},"
                          f"'Line items'!$F$2:$F${det_last},$C{row})")
        ws.cell(row=row, column=9, value=INCOTERM)
        ws.cell(row=row, column=10, value=cbam)
        ws.cell(row=row, column=11, value=ad)
        ws.cell(row=row, column=12, value=note)
        for c in range(1, 13):
            cell = ws.cell(row=row, column=c)
            cell.font = BASE
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=5).number_format = "#,##0"
        for c in (6, 7, 8):
            ws.cell(row=row, column=c).number_format = "#,##0.00"
        if cbam == "ANO" or ad == "ANO":
            for c in range(1, 13):
                ws.cell(row=row, column=c).fill = FILL_RED
            for c in (10, 11):
                ws.cell(row=row, column=c).font = Font(name="Arial", size=10, bold=True, color="C00000")
        row += 1
    end = row - 1
    # subtotal per sender
    ws.cell(row=row, column=4, value=f"Mezisoučet – {sender}").font = BOLD
    for c, L in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
        cell = ws.cell(row=row, column=c, value=f"=SUM({L}{start}:{L}{end})")
        cell.font = BOLD
        cell.number_format = "#,##0" if c == 5 else "#,##0.00"
    for c in range(1, 13):
        ws.cell(row=row, column=c).fill = FILL_SENDER
        ws.cell(row=row, column=c).border = BORDER
    sender_blocks.append((start, end, row))
    row += 1

# grand total
gt = row
ws.cell(row=gt, column=4, value="CELKEM ZÁSILKA / TOTAL SHIPMENT").font = BOLD
for c, L in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
    parts = "+".join(f"{L}{sub}" for _, _, sub in sender_blocks)
    cell = ws.cell(row=gt, column=c, value=f"={parts}")
    cell.font = BOLD
    cell.number_format = "#,##0" if c == 5 else "#,##0.00"
for c in range(1, 13):
    ws.cell(row=gt, column=c).fill = FILL_TOT
    ws.cell(row=gt, column=c).border = BORDER

# cross-check row
cc = gt + 2
ws.cell(row=cc, column=1, value="Kontrola / cross-check").font = BOLD
checks = [
    ("Množství celkem vs. součet faktur", f"=E{gt}", 1324160 + 21168 + 43760),
    ("Čistá hmotnost celkem (kg)", f"=F{gt}", 9574.05 + 12448.80 + 957.00),
    ("Hrubá hmotnost celkem (kg) vs. FCR/SMGS 24 133 kg", f"=G{gt}", 24133.00),
    ("Celní hodnota celkem (EUR) vs. součet faktur", f"=H{gt}", 73481.56 + 25058.88 + 20830.64),
]
cr = cc + 1
ws.cell(row=cc, column=2, value="Ze souhrnu").font = BOLD
ws.cell(row=cc, column=3, value="Z dokladů").font = BOLD
ws.cell(row=cc, column=4, value="Rozdíl").font = BOLD
for label, ref, doc in checks:
    ws.cell(row=cr, column=1, value=label).font = BASE
    ws.cell(row=cr, column=2, value=ref).font = BASE
    ws.cell(row=cr, column=3, value=doc).font = BASE
    ws.cell(row=cr, column=4, value=f"=B{cr}-C{cr}").font = BASE
    for c in (2, 3, 4):
        ws.cell(row=cr, column=c).number_format = "#,##0.00"
    cr += 1

for col, w in zip("ABCDEFGHIJKL", [40, 15, 13, 58, 13, 15, 15, 16, 34, 9, 12, 55]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = f"A{hdr_row + 1}"

# ---------------------------------------------------------------- Notes
nt = wb.create_sheet("Poznámky & flagy")
nt["A1"] = "POZNÁMKY, UPOZORNĚNÍ A NESROVNALOSTI"
nt["A1"].font = TITLE
notes = [
    ("KRITICKÉ", "Preferenční věta / preferenční původ (COO)",
     "V žádném z dodaných dokladů (3× faktura, 3× packing list, FCR TECGO26080029, nákladní list SMGS 38055706) "
     "NENÍ uvedena preferenční věta o původu zboží. Původ zboží = ČÍNA. EU nemá s Čínou dohodu o preferenčním "
     "obchodu, preferenční sazební zacházení tedy NELZE uplatnit a při proclení se použije plná celní sazba "
     "třetí země (erga omnes). Doporučení: nevyžadovat od dodavatele preferenční prohlášení, ale potvrdit "
     "hodnotu a sazby v TARIC před podáním celního prohlášení."),
    ("KRITICKÉ", "Incoterm – chybí místo/přístav a termín neodpovídá druhu přepravy",
     "Všechny tři faktury i packing listy uvádějí pouze „Terms of Payment: FOB“ – BEZ uvedení přístavu / "
     "jmenovaného místa. Dvojí problém: (1) dle standardu musí být vždy uveden termín i město (např. FCA Suzhou); "
     "(2) FOB je dle Incoterms 2020 určen výhradně pro námořní a vnitrozemskou vodní přepravu, ale tato zásilka "
     "jde ŽELEZNICÍ (Zhengzhou/Putian → Malaszewicze → CZ). Pro železniční přepravu je správný termín FCA "
     "s uvedeným místem. Nutno vyžádat od odesílatele opravu / doplnění, jinak nelze spolehlivě určit, zda "
     "je do celní hodnoty nutno připočíst dopravné do hranice EU."),
    ("KRITICKÉ", "Antidumping – spojovací materiál ze železa/oceli (HS 7318), původ Čína",
     "Položky 14 (Nut, 029-2513-AL-CZ) a 15 (Spring nut, 029A-000004-CZ) od odesílatele "
     "Suzhou Sonavox International Trading, zařazené pod 7318 15 95, spadají do rozsahu platných antidumpingových "
     "opatření EU na spojovací prostředky ze železa nebo oceli pocházející z ČLR (nař. (EU) 2022/191 ve znění "
     "pozdějších předpisů, včetně opatření proti obcházení). Celní hodnota těchto položek: 13 140,00 EUR. "
     "Nutno ověřit doplňkový kód TARIC podle konkrétního čínského výrobce – sazba se dle výrobce významně liší."),
    ("KRITICKÉ", "CBAM – spojovací materiál ze železa/oceli (HS 7318)",
     "Tytéž položky pod kódem 7318 15 95 spadají do přílohy I nařízení CBAM (železo a ocel). Celkem 210 000 ks, "
     "čistá hmotnost 312,30 kg. Nutno zahrnout do čtvrtletního CBAM hlášení (vykazuje se čistá hmotnost "
     "a zabudované emise). Ostatní položky zásilky (kap. 35, 39, 74, 85) do CBAM nespadají."),
    ("K OVĚŘENÍ", "Nesoulad: „Nut 029-2513-AL-CZ“ – ocel nebo hliník?",
     "Označení dílu obsahuje „-AL-“, což naznačuje HLINÍK, avšak podklad SONAVOX (e-mail „SONAVOX HS KODY“) "
     "přiřazuje „Insert (nut)“ kód 7318 15 95 19 (železo/ocel). Hliníková matice by patřila pod 7616 10 00. "
     "Zařazení má přímý dopad: 7318 = antidumping ANO + CBAM ANO; 7616 = antidumping NE + CBAM ANO. "
     "V souhrnu je konzervativně použito 7318 15 95 19. NUTNO POTVRDIT MATERIÁL U DODAVATELE."),
    ("VYŘEŠENO", "Přípona u 8518 90 00 = DK 00 (opraveno dle předchozích celních prohlášení)",
     "Referenční tabulka SONAVOX uvádí u částí reproduktorů souběžně přípony 00, 40, 90 i 99. Ve dvou "
     "předchozích celních prohlášeních Sonavox je však pro VŠECHNY části reproduktorů použito jednotně "
     "8518 90 00 s doplňkovým kódem DK 00 – včetně Front plate („Front plate – deska“), T-yoke, Cone, Spider, "
     "Dust cap a Bracket. Front plate byl proto v tomto souhrnu OPRAVEN z 8518 9000 90 na 8518 9000 00. "
     "Přípony 40 / 90 / 99 z tabulky SONAVOX se v praxi nepoužívají."),
    ("K OVĚŘENÍ", "Waterproof membrane – kód přiřazen odvozením",
     "Položka 9 (J2-10X002-0001-CZ, Waterproof membrane, 210 000 ks, 8 820,00 EUR) není v referenční tabulce "
     "SONAVOX uvedena pod tímto názvem. Byl použit kód 3921 90 90 90 podle nejbližší položky "
     "„Protective membrane plast / Plastová membrána“. Doporučeno potvrdit u dodavatele."),
    ("VYŘEŠENO", "Terminal – OPRAVENO z 7409 39 00 na 8518 90 00 00",
     "Položky 10–13 (Terminal, celkem 600 000 ks, 9 380,00 EUR) byly původně zařazeny pod 7409 39 00 podle "
     "položky „Pin, lug / kontakt“ z tabulky SONAVOX. Předchozí celní prohlášení Sonavox však deklaruje "
     "„terminal – části reproduktorů“ pod 8518 90 00 DK 00. Zařazení bylo opraveno na 8518 90 00 00, což je "
     "i věcně správnější (hotový díl určený výhradně pro reproduktor = část zboží čísla 8518 dle pozn. 2 "
     "ke třídě XVI; 7409 je naopak polotovar – desky a pásy z mědi)."),
    ("NESROVNALOST", "FCR uvádí jen jednoho odesílatele, zásilka je však od tří prodávajících",
     "FCR TECGO26080029 uvádí jako shipper pouze SUZHOU SONAVOX ELECTRONICS CO., LTD. Zboží v kontejneru však "
     "pochází od TŘÍ různých prodávajících (Suzhou Sonavox International Trading, Suzhou Sonavox Electronics, "
     "Suzhou Yanlong Electronic Product) dle tří samostatných faktur. Celní prohlášení musí být členěno "
     "dle jednotlivých odesílatelů – proto je souhrn zpracován primárně dle odesílatele."),
    ("NESROVNALOST", "Popis zboží ve FCR neuvádí všechny položky",
     "FCR vyjmenovává: Speaker component, Speaker, Voice coil, Front plate, Gasket, Waterproof membrane, "
     "Terminal, Nut, Spring nut, Harness, Gauge, Cone. CHYBÍ zde SPIDER (středící membrána, 77 600 ks) "
     "a CERAMICS CIRCLE (magnetický obvod, 19 152 ks) – zřejmě jsou zahrnuty pod obecné „Speaker component“, "
     "ale doporučeno upozornit dopravce."),
    ("NESROVNALOST", "Nákladní list SMGS deklaruje celý kontejner pod jedním kódem 8518 90",
     "SMGS 38055706 uvádí HS 851890 / 85189000 „Speaker component“ (+ HS 99020000 sběrné zboží) pro celý "
     "kontejner. Skutečné zboží však spadá do kapitol 35, 39, 73, 74 a 85. Jde o zjednodušení přepravního "
     "dokladu, které nemá vliv na celní zařazení, ale může vyvolat dotaz celního úřadu – vhodné mít "
     "připravené rozúčtování dle faktur."),
    ("NESROVNALOST", "Rozdílná čísla bankovních účtů mezi packing listem a fakturou (CZ20260725)",
     "U dokladu CZ20260725 (Suzhou Sonavox International Trading) uvádí packing list účet EUR 500158216108, "
     "zatímco faktura uvádí EUR 468960663282. Nemá vliv na celní řízení, ale před úhradou doporučeno "
     "ověřit správné číslo účtu přímo u dodavatele (riziko podvodné změny platebních údajů)."),
    ("NESROVNALOST", "ETD/ETA neodpovídá datu vypravení",
     "Packing list CZ20260725-2 (Yanlong) uvádí ETD 29.07.2026 / ETA 29.08.2026; FCR byl však vystaven "
     "04.08.2026 a vlak má č. V.WB2026/08/05 (05.08.2026). U zbylých dvou sad dokladů jsou ETD/ETA prázdné. "
     "Doporučeno sjednotit."),
    ("VYŘEŠENO", "Potvrzeno předchozími celními prohlášeními Sonavox",
     "Z dvou předchozích celních prohlášení na obdobné zásilky Sonavox byly potvrzeny tyto kódy: "
     "Cone, Spider, Dust cap, T-yoke, Front plate, Bracket, Speaker component i Terminal = 8518 90 00 DK 00; "
     "Car speaker = 8518 21 00 DK 00; Harness / kabelový svazek = 7413 00 00 DK 90; "
     "Gasket (plastové těsnění) + Gauge (plastový kalibr) společně na jednom řádku = 3926 90 97 DK 90. "
     "Potvrzen je rovněž postup slučování více komponent pod jeden kód na jeden řádek."),
    ("OK", "Kontrolní součty souhlasí",
     "Hrubá hmotnost 9 950,74 + 12 907,84 + 1 274,42 = 24 133,00 kg = FCR (24 133 KGS) i SMGS (24 133 kg). "
     "Kolí 12 + 23 + 16 = 51 palet = FCR (51 PLTS) i SMGS (51). Objem 15,24 + 16,64 + 19,52 = 51,40 CBM = FCR "
     "(51,4 CBM). Množství i celní hodnota souhlasí s fakturami."),
    ("OK", "Sleva / rabat",
     "V žádné z faktur není uvedena obchodní sleva ani rabat; součet položek = Sub Total. Poměrné rozpuštění "
     "slevy do celní hodnoty se tedy neuplatňuje."),
    ("INFO", "Dopravné / celní hodnota",
     "FCR uvádí FREIGHT COLLECT (dopravné hradí příjemce) a fakturační podmínka je FOB. Do celní hodnoty "
     "je proto nutno připočíst náklady dopravy a pojištění do místa vstupu do EU (hranice EU – Malaszewicze/PL). "
     "Podklady k dopravnému nebyly součástí dodaných dokladů."),
]
r = 3
nt.cell(row=r, column=1, value="Priorita").font = HDR
nt.cell(row=r, column=2, value="Téma").font = HDR
nt.cell(row=r, column=3, value="Popis").font = HDR
for c in range(1, 4):
    nt.cell(row=r, column=c).fill = FILL_HDR
    nt.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    nt.cell(row=r, column=c).border = BORDER
r += 1
for prio, topic, text in notes:
    nt.cell(row=r, column=1, value=prio)
    nt.cell(row=r, column=2, value=topic)
    nt.cell(row=r, column=3, value=text)
    for c in range(1, 4):
        cell = nt.cell(row=r, column=c)
        cell.font = BASE
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    if prio == "KRITICKÉ":
        for c in range(1, 4):
            nt.cell(row=r, column=c).fill = FILL_RED
        nt.cell(row=r, column=1).font = Font(name="Arial", size=10, bold=True, color="C00000")
        nt.cell(row=r, column=2).font = BOLD
    elif prio in ("K OVĚŘENÍ", "NESROVNALOST"):
        for c in range(1, 4):
            nt.cell(row=r, column=c).fill = FILL_WARN
        nt.cell(row=r, column=2).font = BOLD
    nt.row_dimensions[r].height = 58
    r += 1

nt.column_dimensions["A"].width = 15
nt.column_dimensions["B"].width = 46
nt.column_dimensions["C"].width = 118
nt.freeze_panes = "A4"

# ---------------------------------------------------------------- Sources
src = wb.create_sheet("Zdroje")
src["A1"] = "ZDROJOVÉ DOKLADY A PŘIŘAZENÍ HS KÓDŮ"
src["A1"].font = TITLE
rows = [
    ("Doklad", "Obsah / použití"),
    ("TRCZ20260725.xlsx (listy PL, INV)", "Faktura + packing list CZ20260725 – Suzhou Sonavox International Trading Co., Ltd (19 položek)"),
    ("ELCZ202607251_...GCS.xlsx (2 soubory)", "Faktura + packing list CZ20260725-1 – Suzhou Sonavox Electronics Co., Ltd (2 položky)"),
    ("YLCZ202607252.xlsx (listy PL, INV)", "Faktura + packing list CZ20260725-2 – Suzhou Yanlong Electronic Product Co., Ltd (3 položky)"),
    ("HBL_TECGO26080029.pdf", "Forwarder's Cargo Receipt T.H.I. Group Ltd. – trasa, kontejner, brutto, kolí, objem"),
    ("CIMU1716969.pdf", "Nákladní list SMGS č. 38055706 – železniční přeprava, plomba, hmotnosti, trasa"),
    ("Re_SONAVOX_HS_KODY.eml", "Referenční tabulka HS kódů s českými popisy (Ing. Michaela Ryan, Sonavox Technology CZ) – zdroj HS kódů a českých popisů"),
    ("Předchozí celní prohlášení Sonavox (2 ks)", "Screenshoty položek dřívějších celních prohlášení na obdobné "
     "zásilky – zdroj potvrzení kódů 8518 90 00 DK 00, 8518 21 00 DK 00, 7413 00 00 DK 90 a 3926 90 97 DK 90"),
    ("", ""),
    ("POZOR", "Faktury ANI packing listy neobsahují žádné HS kódy. Kódy v tomto souhrnu pocházejí "
              "z referenční tabulky „SONAVOX HS KODY“ a z předchozích celních prohlášení Sonavox. "
              "Kde se oba zdroje liší, má přednost celní prohlášení (Front plate, Terminal). "
              "Zbývající nepotvrzené položky jsou na listu Poznámky & flagy označeny jako „K OVĚŘENÍ“: "
              "Nut / Spring nut (materiál) a Waterproof membrane. Voice coil je podložen závaznou "
              "informací o sazebním zařazení ZISZ 30-0292-2016."),
]
r = 3
for a, b in rows:
    src.cell(row=r, column=1, value=a).font = BOLD if a in ("Doklad", "POZOR") else BASE
    src.cell(row=r, column=2, value=b).font = BOLD if a == "Doklad" else BASE
    src.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    if a == "POZOR":
        src.cell(row=r, column=1).fill = FILL_RED
        src.cell(row=r, column=2).fill = FILL_RED
        src.cell(row=r, column=1).font = Font(name="Arial", size=10, bold=True, color="C00000")
        src.row_dimensions[r].height = 46
    r += 1
src.column_dimensions["A"].width = 42
src.column_dimensions["B"].width = 110

out = "/home/user/customs/output/HS_Code_Summary_Sonavox_Technology_CZ_TECGO26080029.xlsx"
wb.save(out)
print("saved", out)
