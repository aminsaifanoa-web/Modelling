# DCF Valuation Model — Project Meridian

A five-year unlevered free-cash-flow DCF for a hypothetical industrial target ("Meridian Industrials, Inc.", USD millions). Built as live, auditable Excel formulas — no pasted values anywhere in the output cells.

## Files

- `DCF_Model.xlsx` — live workbook
- `DCF_Model.pdf` — PDF export of all tabs

## Tab structure

| Tab | Contents |
|---|---|
| **Cover** | Contents, conventions, units, disclaimer |
| **Assumptions** | Valuation date, CAPM WACC, operating drivers, terminal-value inputs, capital structure & shares |
| **OperatingModel** | Revenue → EBITDA → EBIT → NOPAT → unlevered FCF, 2027E–2031E, with margins and totals |
| **DCFValuation** | Discount periods (mid-year adjusted), discount factors, PV of FCF, both terminal values, EV → equity → per-share |
| **Sensitivity** | Implied price across WACC × terminal-g and WACC × exit-multiple grids |

## Key assumptions (all editable inputs)

- Risk-free **4.0%**, ERP **5.0%**, beta **1.0** → **WACC 9.00%**, **mid-year convention** on
- LTM revenue **$1,000**, growth **8% → 4%** fading, EBITDA margin **30%**, D&A 7%, CapEx 8% of revenue, ΔNWC 12% of the revenue *change*, tax **25%**
- Terminal value (two methods, both shown): Gordon **g = 2.5%**, or exit **9.0× EV/EBITDA** (primary)
- Net debt **$700** (debt $800 / cash $100) · diluted shares **50mm** · current price **$45.00**

## Verified outputs (Excel-calculated and cross-checked)

| Output | Exit multiple (primary) | Gordon growth |
|---|---|---|
| PV of forecast FCF | $762.2 | $762.2 |
| Terminal value | $3,611.6 | $3,330.3 |
| Enterprise value | **$3,212.8** | $3,021.9 |
| Equity value | $2,512.8 | $2,221.9 |
| Implied price | **$50.26 (+11.7%)** | $46.44 (+3.2%) |

The Gordon TV implies a terminal EV/EBITDA of **8.30×** — a sanity cross-check between the two methods. Forecast uFCF: 165.9 / 178.7 / 190.7 / 201.7 / 211.2.

## Design notes

- **Mid-year convention**: discount periods are 0.5 … 4.5, driven by a 1/0 toggle in Assumptions (`C12`) — set it to 0 for year-end discounting.
- **Sensitivity grids without DataTables**: each grid cell is a full `SUMPRODUCT` recomputation of the DCF at that row/column combination (WACC × g, WACC × exit multiple), so the grids stay live even if someone opens the file with calculation-what-if tables blocked. Both grids are exact, not linearised approximations, and are verified against the independent mirror at their centres and corners.
- **Check cell**: `DCFValuation!H27` shows `OK` when `PV of FCF + PV of TV = Enterprise value`.

## Verification

Built by `scripts/build_dcf.ps1` (Excel COM). The script computes an independent PowerShell mirror of every key output, runs `CalculateFullRebuild`, and asserts Excel's cells equal the mirror before saving: 17 assertions (PVs, TVs, EVs, equity, prices, upside, implied multiple, plus grid centres and corners) — all pass; the workbook check cell reads `OK`.
