# Modelling — study-quality finance models

Two self-contained Excel models, built and **verified by Excel itself**: every headline output was recalculated by Excel and cross-checked cell-by-cell against an independent PowerShell "mirror" implementation before the files were saved. Nothing is hard-coded — the workbooks are live formulas end to end.

| Project | Model | Headline outputs (verified) |
|---|---|---|
| [`DCF-Model/`](DCF-Model/) | **Project Meridian** — unlevered DCF valuation | EV **$3,212.8mm** (exit multiple) / $3,021.9mm (Gordon) · price **$50.26 (+11.7%)** / $46.44 (+3.2%) |
| [`LBO-Model/`](LBO-Model/) | **Project Lightyear** — leveraged buyout | **MOIC 2.55x · IRR 20.5%** on $1,320mm sponsor equity · 5.6x entry → 2.7x exit leverage |

## What's in each folder

- `<Model>.xlsx` — the live workbook (all formulas auditable, colour-coded inputs)
- `<Model>.pdf` — a PDF printout of every tab
- Short README with assumptions, outputs and structure

## Conventions (both workbooks)

- **Blue bold font** = input you can change
- **Black** = formula · **Green** = link pulling from another tab
- Every tab ends in a visible **"OK" check cell** (EV↔price / MOIC↔IRR identities)
- Units: USD millions unless noted. Fictitious illustrative data — **not investment advice**

## Rebuilding / re-verifying

`scripts/build_dcf.ps1` and `scripts/build_lbo.ps1` regenerate the workbooks from scratch via Excel COM and refuse to save unless every key cell matches the independent mirror (see each script header). Requires Windows + Microsoft Excel.

## Disclaimer

Training/portfolio models with invented data. Nothing here is a recommendation, valuation advice, or a model of any real company.
