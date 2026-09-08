# Domain Glossary · CDF Helper

This glossary names the concepts of the CDF Helper domain. It is a glossary and nothing else —
no implementation details, no specs, no decisions. Terms are captured here the moment they crystallize.

## Terms

- **Generation job** — a single request to produce one customs-declaration 报关清单 for a vessel, port, and date.
  A generation job has a status (running / done / failed) and a live progress log. It may run in the background
  (for example, while the AI enrichment estimates missing weights and prices) and be polled for progress.

- **Spare part (备件)** — a line item in the source file: name, quantity, unit, and optionally a spec/model,
  weight (KG), and unit price (RMB). The unit of parsing and of the generated 报关清单.