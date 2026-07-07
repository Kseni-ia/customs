# CLAUDE.md — Standing instructions for customs / shipment work

These are my standing preferences. Read this file at the start of every session and
follow it whenever we work on a shipment. If anything here conflicts with what I ask
in the moment, the in-the-moment request wins — but otherwise treat this as the default.

---

## What this project is

I prepare customs documentation for shipments — primarily **HS code summaries** produced
as Excel files. Each shipment has a client and a shipment number
(e.g. "Czech Healthcare, shipment 196652").

## Output files

- Save all deliverables to the `output/` folder.
- Name Excel files: `HS_Code_Summary_<Client>_<ShipmentNumber>.xlsx`
  - Example: `output/HS_Code_Summary_Czech_Healthcare_196652.xlsx`
- Use underscores instead of spaces in file names.

## HS code summary — default rules

- **Merge line items to unique HS codes.** Don't list the same HS code on multiple rows —
  consolidate (e.g. 9019 raw rows → the set of distinct HS codes they map to).
- **Do NOT include an Incoterm column** in the summary.
- Include **gross weights** AND **net weights** per HS code.
- Include the **total quantity** for each HS code (sum the quantities of the line items
  merged into that HS code).
- Apply the **proportional trade discount** to customs values when a discount applies to
  the shipment (distribute it across line items in proportion to value).
- **Use the Czech goods descriptions from the invoice** (the Czech column, e.g. "Rychlovarná
  konvice", "Odpadkový koš") as the primary description column — that's what goes on the CZ
  customs declaration. Keep the English name only as a secondary column.
- Flag **CBAM** (Carbon Border Adjustment Mechanism) applicable items.
- Add an **antidumping note** for any items subject to antidumping duties.

## Document checks — COO / preferential origin (VERY IMPORTANT)

- **Always check the shipment documents (invoices, etc.) for a `preferenční věta`** — a
  preferential-origin declaration. This is about **COO (country of origin)** and it directly
  affects the duty owed at clearance.
- The statement looks like this (example, Korea):
  > "The exporter of the products covered by this document (customs authorization
  > No. 020-12-200186) declares that, except where otherwise clearly indicated, these
  > products are of Republic of Korea preferential origin."
- **Why it matters:** without it, the standard third-country duty applies (e.g. **Jinzi Korea
  → 2.7 % clo**). With a valid preferential statement, the goods qualify for **preferential
  tariff treatment** (reduced / zero duty) instead.
- **If a `preferenční věta` is present in the docs, flag it to me explicitly** as a prominent
  note — it is critical for correct clearance of the shipment. Don't bury it; call it out.

## Defaults & assumptions

<!-- Fill these in so I stop asking each time. Delete any that don't apply. -->
- Default destination country: <!-- e.g. Czech Republic / EU -->
- Default currency: <!-- e.g. EUR -->
- Preferred columns (in order): <!-- e.g. HS code | Description | Gross weight | Customs value | CBAM | Antidumping -->

## How I like to work

- When a new shipment starts, ask me only for what you can't infer (client, shipment
  number, source data file) — don't re-ask about the rules above.
- Show me the resulting file / a summary of what changed after each update.
- Commit each meaningful change with a clear message.

## Notes / things to remember

<!-- Add per-client quirks, recurring HS codes, past decisions, etc. here over time. -->
