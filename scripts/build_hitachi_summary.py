#!/usr/bin/env python3
"""Build per-item (thyristor / diode) summary for the EX HITACHI 1797 shipment
(Hitachi Energy CZ -> Shindengen, Japan)."""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

# One record per invoiced item. net_amount = total invoice price (EUR).
# gross weights taken from the paired delivery note.
rows = [
    # type, material, description, hs, qty, net_kg, gross_kg, unit_price, net_amount, invoice, dn, incoterms, currency
    ("Diode",     "5SDD 24F2800",  "5SDD 24F2800 Rectifier diodes",       "85411000",  40, 19.6,  21.6,  168.00,  6720.00,  "810678828", "0923633858", "FCA Prague", "EUR"),
    ("Diode",     "5SDA 14F5007",  "5SDA 14F5007 Avalanche diodes",       "85411000",   5,  2.35,  2.6,  209.00,  1045.00,  "810678829", "0923633901", "FCA Prague", "EUR"),
    ("Thyristor", "5STP 3328L0011","YST 45 P28 E / Thyristor",            "85413000", 100, 140.0, 140.0, 484.00, 48400.00,  "810678830", "0923633955", "FCA Prague", "EUR"),
    ("Thyristor", "5STP 12F4200",  "YST 14-29 P42D / Thyristor",          "85413000",   5,  2.4,   2.75, 181.00,   905.00,  "810678833", "0923633974", "FCA Prague", "EUR"),
    ("Diode",     "5SDD 3050L0003","YSD 45-02 P50 / Diode",               "85411000",  47, 58.75, 58.75, 549.00, 25803.00,  "810678834", "0923634051", "FCA Prague", "EUR"),
    ("Thyristor", "5STP2552L0025", "PCT 5200 V 2500 A SF Thyristor",      "85413000",  60, 87.0,  None,  442.00, 26520.00,  "810678838", "0923634362", "FCA Prague", "EUR"),
    ("Diode",     "5SDD7102B0001", "5SDD 7102B0001 Welding diodes",       "85411000", 400, 56.0,  60.0,   71.00, 28400.00,  "810678853", "0923637756", "FCA Prague", "EUR"),
    ("Diode",     "5SDD0135Z0401", "5SDD 0135Z0401 Welding diodes",       "85411000", 200, 28.0,  34.0,  108.00, 21600.00,  "810678854", "0923637801", "FCA Prague", "EUR"),
]

headers = ["Type", "Material No.", "Description", "HS code", "Quantity",
           "Net weight (kg)", "Gross weight (kg)", "Unit price (EUR)",
           "Total invoice price (EUR)", "Invoice No.", "Delivery Note No.",
           "Incoterms", "Currency"]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Summary"

# Styles
hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(bold=True, color="FFFFFF", size=11)
title_font = Font(bold=True, size=14)
sub_font = Font(italic=True, size=10, color="555555")
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
diode_fill = PatternFill("solid", fgColor="EAF3FB")
thyr_fill = PatternFill("solid", fgColor="FBF3EA")
total_fill = PatternFill("solid", fgColor="D9E1F2")
total_font = Font(bold=True)

# Title block
ws["A1"] = "EX HITACHI 1797 — Item summary (thyristors & diodes)"
ws["A1"].font = title_font
ws["A2"] = "Exporter: Hitachi Energy Czech Republic s.r.o.  |  Consignee: Shindengen Electric Mfg. Co., Ltd. (Japan)  |  Invoice date: 14.07.2026"
ws["A2"].font = sub_font
ws["A3"] = "8 invoices / 8 delivery notes.  Export declaration, 0% VAT.  Origin of goods: CZ."
ws["A3"].font = sub_font

hdr_row = 5
for c, h in enumerate(headers, start=1):
    cell = ws.cell(row=hdr_row, column=c, value=h)
    cell.fill = hdr_fill
    cell.font = hdr_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

group_fill = PatternFill("solid", fgColor="2E75B6")
group_font = Font(bold=True, color="FFFFFF", size=11)
sub_fill = PatternFill("solid", fgColor="D6E4F0")


def write_item(r, rec):
    (typ, mat, desc, hs, qty, net, gross, unit, total, inv, dn, inco, cur) = rec
    vals = [typ, mat, desc, hs, qty, net, gross, unit, total, inv, dn, inco, cur]
    fill = thyr_fill if typ == "Thyristor" else diode_fill
    for c, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = border
        cell.fill = fill
        if c == 5:
            cell.number_format = "#,##0"
        if c in (6, 7, 8, 9):
            cell.number_format = "#,##0.00"
        if c in (10, 11):
            cell.number_format = "@"
        if c in (4, 5, 6, 7, 8, 9, 12, 13):
            cell.alignment = Alignment(horizontal="center")
    if gross is None:
        ws.cell(row=r, column=7, value="n/a").alignment = Alignment(horizontal="center")


def write_group(r, label, recs):
    # group banner
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    gc = ws.cell(row=r, column=1, value=label)
    gc.fill = group_fill
    gc.font = group_font
    gc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for c in range(1, 14):
        ws.cell(row=r, column=c).border = border
    r += 1
    for rec in recs:
        write_item(r, rec)
        r += 1
    # subtotal
    q = sum(x[4] for x in recs)
    nt = sum(x[5] for x in recs)
    gt = sum((x[6] or 0) for x in recs)
    vt = sum(x[8] for x in recs)
    ws.cell(row=r, column=3, value=f"Subtotal — {label} ({len(recs)} items)")
    ws.cell(row=r, column=5, value=q).number_format = "#,##0"
    ws.cell(row=r, column=6, value=round(nt, 2)).number_format = "#,##0.00"
    ws.cell(row=r, column=7, value=round(gt, 2)).number_format = "#,##0.00"
    ws.cell(row=r, column=9, value=round(vt, 2)).number_format = "#,##0.00"
    for c in range(1, 14):
        cell = ws.cell(row=r, column=c)
        cell.fill = sub_fill
        cell.font = total_font
        cell.border = border
        if c in (4, 5, 6, 7, 8, 9, 12, 13):
            cell.alignment = Alignment(horizontal="center")
    return r + 1


diodes = [x for x in rows if x[0] == "Diode"]
thyristors = [x for x in rows if x[0] == "Thyristor"]

r = hdr_row + 1
r = write_group(r, "DIODES  —  HS 8541 10 00", diodes)
r = write_group(r, "THYRISTORS  —  HS 8541 30 00", thyristors)

# Grand total row
tot_row = r
ws.cell(row=tot_row, column=1, value="GRAND TOTAL")
ws.cell(row=tot_row, column=3, value="8 items (5 diode types, 3 thyristor types)")
qty_total = sum(x[4] for x in rows)
net_total = sum(x[5] for x in rows)
gross_total = sum((x[6] or 0) for x in rows)
val_total = sum(x[8] for x in rows)
ws.cell(row=tot_row, column=5, value=qty_total).number_format = "#,##0"
ws.cell(row=tot_row, column=6, value=round(net_total, 2)).number_format = "#,##0.00"
ws.cell(row=tot_row, column=7, value=round(gross_total, 2)).number_format = "#,##0.00"
ws.cell(row=tot_row, column=9, value=round(val_total, 2)).number_format = "#,##0.00"
ws.cell(row=tot_row, column=12, value="FCA Prague")
ws.cell(row=tot_row, column=13, value="EUR")
for c in range(1, 14):
    cell = ws.cell(row=tot_row, column=c)
    cell.fill = total_fill
    cell.font = total_font
    cell.border = border
    if c in (4, 5, 6, 7, 8, 9, 12, 13):
        cell.alignment = Alignment(horizontal="center")

# Notes
n = tot_row + 2
notes = [
    "Consistency checks:",
    "  • Incoterms — CONSISTENT across all 16 documents: FCA Prague (minor text variants 'FCA Prague, CZ' / 'FCA  Prague', same term).",
    "  • Currency  — CONSISTENT across all invoices: EUR (each invoice also shown in CZK at 1 EUR = 24.25 CZK, informational only).",
    "  • Gross weight for DN 0923634362 (5STP2552L0025) is printed as 0 on the delivery note — likely a data error; net weight 87 kg. Please verify.",
    "  • No 'preferenční věta' (preferential-origin declaration) found in any document. Goods origin CZ; export to Japan at 0% VAT.",
    "  • Delivery-note numbers are given in the 092… form shown on the delivery notes; invoices list the same number without the leading zero.",
]
for i, txt in enumerate(notes):
    cell = ws.cell(row=n + i, column=1, value=txt)
    cell.font = Font(bold=(i == 0))

# Column widths
widths = [11, 15, 32, 11, 10, 15, 17, 15, 22, 13, 17, 13, 10]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = "A6"

os.makedirs("/home/user/customs/output", exist_ok=True)
out = "/home/user/customs/output/HS_Code_Summary_Hitachi_Shindengen_EX1797.xlsx"
wb.save(out)
print("saved", out)
print("totals: qty", qty_total, "net", round(net_total,2), "gross", round(gross_total,2), "value EUR", round(val_total,2))
