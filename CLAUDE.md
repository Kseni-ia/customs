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

### ⚠️ Required columns — NEVER omit these (one row per unique HS code)

Every HS code summary **MUST** include the following columns for **each** HS code. Do not
drop any of them, even if the source data makes them tedious to compute:

1. **HS code**
2. **Description** — see the description-language rule below (Czech prioritised).
3. **Net weight** per HS code (sum the net weights of the merged line items).
4. **Gross weight** per HS code (sum the gross weights of the merged line items).
5. **Quantity** — the total quantity for the HS code (sum the quantities of the merged line items).
6. **Price / customs value** for the HS code.

If any of these values is genuinely missing from the source data, keep the column and tell
me it's missing — don't silently drop the whole column.

### Description language — Czech is prioritised

- **If a Czech description of the item is available in the source docs, use it.** The Czech
  description takes priority over any other language for the Description column.
- Only fall back to another language when no Czech description exists.

### Classification — sender first, then HS code

- **Classify the shipment by sender first, then by HS code.** The **sender takes priority**
  over the HS code.
- **Merge line items to unique HS codes _within each sender_.** Don't list the same HS code
  on multiple rows for the same sender — consolidate (e.g. 9019 raw rows → the set of
  distinct HS codes they map to).
- **If the same HS code appears under different senders, keep it on separate rows — one per
  sender.** Do NOT merge identical HS codes across senders; the shipment must be classified
  per sender in that case.

### Other rules

- **Always include the Incoterm** in the summary — both the **Incoterm term** (e.g. FCA,
  EXW, DAP, CIP) **and the city / named place** that goes with it (e.g. `FCA Prague`,
  `EXW Seoul`). Never leave out the city.
- Apply the **proportional trade discount** to customs values when a discount applies to
  the shipment (distribute it across line items in proportion to value).
- Flag **CBAM** (Carbon Border Adjustment Mechanism) applicable items. See the
  **CBAM lookup** section below for what to report.
- Add an **antidumping note** for any items subject to antidumping duties.

## CBAM lookup — by CZ receiver (VERY IMPORTANT)

- The CBAM registry lives in **`reference/CBAM_receivers_CZ.xlsx`** (sheet `List1`).
  This is the source of truth for CBAM status. Whenever a shipment contains a
  **CBAM-applicable HS code**, look the receiver up in this file and report the result.
- **Look up by the CZ receiver** (the importing company in the Czech Republic) — match on
  the **`Firma`** column (or its **`EORI`**). CBAM status is tied to the receiver, not the
  sender or the HS code.
- **When the receiver is found, show me their CBAM status** from the row — specifically:
  - the **`Y KÓD`** (`Y128`, `Y238`, `Y137`, `Y237`) — this is the declaration code to use,
  - the **authorization number** (`povolení`, e.g. `CBAM-CZ-2025-…`) if present,
  - the **application** reference (`podaná žádost`, e.g. `APPL-CZ-…`) if present,
  - the **EORI** and issue date (`datum vydání`) when available.

  Quick guide to the Y codes:
  - **`Y128`** — authorization granted (`povolení`); full CBAM authorised declarant.
  - **`Y238`** — application submitted (`podaná žádost`), not yet authorised.
  - **`Y137`** — below the de-minimis threshold (`do 50 t ročně`, under 50 t/year).
- **If the receiver company is NOT in the file, tell me explicitly** — don't guess or assume
  a status. Flag it prominently so I can check the company and add it to the registry.
- The file also lists **non-EU companies that use our shared CBAM account**
  (`CBAM-CZ-2025-QGM67089385721`) in the right-hand columns. If a shipment involves one of
  these, note that it goes under the shared account.

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
