# Domain Glossary · CDF Helper

This glossary names the concepts of the CDF Helper domain. It is a glossary and nothing else —
no implementation details, no specs, no decisions. Terms are captured here the moment they crystallize.

## Terms

- **Generation job** — a single request to produce one customs-declaration 报关清单 for a vessel, port, and date.
  A generation job has a status (running / done / failed) and a live progress log. It may run in the background
  (for example, while the AI enrichment estimates missing weights and prices) and be polled for progress.

- **Spare part (备件)** — a line item in the source file: name, quantity, unit, and optionally a spec/model,
  weight (KG), and unit price (RMB). The unit of parsing and of the generated 报关清单.

- **Per-unit weight** — the `weight` value carried on a Spare part: kilograms per single item (kg/piece). This is a source-side quantity. When the generator fills the 重量(KG) column of the 报关清单 sheet, it writes the **per-row total** weight = `per-unit weight × quantity`; the sheet's total row sums that column. A part with no known weight contributes a blank cell (not zero).

- **Per-row weight** — the value placed in the generated sheet's 重量(KG) column for a line item: `per-unit weight × quantity`. The template's total row aggregates these (e.g. `=SUM(E$3:...E_n)`), so the column holds per-row totals, not per-unit values.