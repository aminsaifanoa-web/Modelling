# LBO Model — Project Lightyear

A standard private-equity buyout model for a hypothetical target ("WidgetCo Industrials, Inc.", USD millions): sources & uses, full five-year debt schedule with a **100% cash sweep**, and sponsor returns. All live Excel formulas, no circular references.

## Files

- `LBO_Model.xlsx` — live workbook
- `LBO_Model.pdf` — PDF export of all tabs

## Tab structure

| Tab | Contents |
|---|---|
| **Cover** | Contents, conventions, disclaimer |
| **Assumptions** | Transaction terms, financing structure (TLB + senior notes), operating drivers, exit |
| **SourcesUses** | Uses (EV + fees), sources (debt + sponsor-equity plug), funding %, entry leverage |
| **DebtSchedule** | Levered FCF, TLB mandatory amortisation + cash sweep, notes (bullet), cash build, credit stats |
| **Returns** | Exit value → equity → MOIC / IRR, with an IRR identity cross-check |
| **Sensitivity** | IRR and MOIC grids across entry × exit EV/EBITDA multiples (exact recomputation) |

## Transaction (verified outputs)

| Item | Value |
|---|---|
| LTM EBITDA (2026A) | **$500.0** (revenue $1,250, 40% margin) |
| Entry multiple | **8.0×** → EV **$4,000.0** |
| Transaction fees (3% of EV) | $120.0 |
| **Total uses** | **$4,120.0** |

| Sources | Amount | % |
|---|---|---|
| Term Loan B (4.4×, **9%**, 5% mandatory amort.) | $2,200.0 | 53.4% |
| Senior notes (1.2×, **11%**, bullet) | $600.0 | 14.6% |
| **Sponsor equity (plug)** | **$1,320.0** | 32.0% |

## Debt schedule & credit (verified outputs)

Levered FCF after interest, cash taxes, CapEx and ΔNWC grows **$150.1 → $277.6**; the sweep applies 100% of post-amortisation FCF to the TLB:

| Year | 1 | 2 | 3 | 4 | 5 |
|---|---|---|---|---|---|
| TLB ending balance | 2,049.9 | 1,872.2 | 1,664.3 | 1,423.1 | **1,145.5** |
| Net leverage (×) | 5.05 | 4.48 | 3.91 | 3.33 | **2.74** |

Total net debt falls from **$2,800.0 → $1,745.5** — **$1,054.5 (37.7%)** paid down over the hold.

## Returns (verified outputs)

- Revenue grows 5% p.a.; EBITDA compounds to **$638.1** by year 5
- Exit at **8.0×** → exit EV **$5,105.1**, less net debt $1,745.5 → **exit equity $3,359.6**
- **MOIC 2.55× · IRR 20.5%** on $1,320.0 invested

## Design notes

- **No circularity**: interest is charged on *beginning-of-period* balances — slightly conservative, keeps every cell auditable.
- **Check cell**: `Returns!C18` shows `OK` when `(1 + IRR)^hold = MOIC`.
- **Exact grids**: because the debt-paydown path is independent of entry/exit multiples, each IRR/MOIC grid cell re-derives returns closed-form at that combination — no DataTables needed, and grids survive stripped-down calculation settings.

## Verification

Built by `scripts/build_lbo.ps1` (Excel COM). The script computes an independent PowerShell mirror (full FCF and debt-path simulation), runs `CalculateFullRebuild`, and asserts workbook cells equal the mirror before saving: **27 assertions** (S&U, per-year FCF / TLB / leverage, exit values, MOIC, IRR, paydown, and grid centres + corners) — all pass; the internal MOIC↔IRR check cell reads `OK`.
