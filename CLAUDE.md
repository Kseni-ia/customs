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

### MOREK CZ s.r.o. — Guangdong Grandview Technology (heat-shrink tubing)

- **Sender quirk:** Grandview leaves the **HS CODE column on their invoices completely
  empty** — every line, both invoices. Expect to classify from scratch each time, and
  it's worth asking them to start filling it in.
- **HS codes come from the forwarder, not the documents.** Maurice Ward (Lukáš
  Jirkovský, `customs.praha@mauriceward.com`) sends them by e-mail with the ETA and
  JSD routing — ask him for them if they haven't arrived. His e-mail is also the only
  **Czech-language** description of the goods, so use his wording verbatim.
- **Codes used for shipment CZ26008741 (Aug 2026)**, per his e-mail:
  - `8547 20 00` — smršťovací trubice silnostěnná a středněstěnná (heavy + medium wall)
  - `3917 40 00` — spojovací set (MSTS kits)
  - `8547 90 00` — smršťovací bužírka tenkostěnná (thinwall sleeving)
- **Two standing reservations about those codes** — raised, not yet resolved:
  - Thinwall is *wholly plastic*, so `8547 90 00` (residual, aimed at metal conduit
    lined with insulating material) sits oddly beside heavy/medium wall in `8547 20 00`.
  - **Note 8 to Chapter 39** limits heading 3917 to tubes *conveying gases or liquids*,
    and `3917 40` covers fittings for such pipes — so it's doubtful for cable-jointing
    sets made purely of heat-shrink sleeve. GIR 3(b) would point at `8547 20 00`.
    The kits are ~36 % of shipment value, so **the duty difference is material** —
    check both rates in TARIC. **A BTI is worth getting** given the recurring volume.
- **Their B/L declares a single blanket `8547 90 00`** for the whole container, which
  matches none of the three-code split — don't rely on it.
- **China = no preferential origin is possible.** The EU has no FTA with China, so a
  `preferenční věta` can never help here — full third-country duty always applies.
  Still check for it, but don't expect one.
- **Packing lists give gross weight per pallet only**, never per item — apportion pro
  rata to net weight within each container so the total ties to the B/L.
- **Their invoice PDFs have overlapping text runs** in the left "Description of Goods"
  column, so it extracts as scrambled characters. The **"Printing" column is clean** —
  use that (product code + size + length) plus the Customer PN instead.
- **EXW + freight collect** — the invoice total is *not* the customs value. Freight and
  insurance to the EU border must be added; Maurice Ward sends the amount with the MRN,
  then apportion it across the HS codes by value.
- **Origin is stated on the invoice face only** (`Country of Origin: CHINA`) — there is
  no separate certificate of origin in the pack, and none is needed for a
  non-preferential declaration absent a trade-defence measure.
