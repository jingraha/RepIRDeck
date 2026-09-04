# Replit Series E investor materials

This repository generates an illustrative 26-slide Replit Series E investor presentation and its supporting source and assumptions workbook.

The materials use public information plus clearly labeled analyst estimates. They do not contain Replit confidential information and are not investment advice.

## Build

```powershell
python -m pip install -r requirements.txt
python -m src.build
```

Outputs:

- `output/replit_series_e_investor_deck_v3.pptx`
- `output/replit_series_e_sourcebook_v3.xlsx`

## Update assumptions or evidence

- Edit operating, transaction, market, and scenario inputs in `data/assumptions.yaml`.
- Add or update citations in `data/public_sources.yaml`.
- Update slide headlines and source mappings in `data/slide_content.yaml`.
- Run `python -m src.build` to regenerate both Office files from the same model.

## Evidence labels

| Label | Meaning |
|---|---|
| Public | Direct or attributable public disclosure |
| Derived | Arithmetic based on public facts |
| Est. | Analyst estimate based on stated assumptions |
| MT | Forward-looking management target |

Every headline metric is listed in the sourcebook's `Slide Reconciliation` tab. Public URLs and claim-level evidence are available in `Source Register` and `Public Facts`.

The first five workbook tabs are:

1. `Assumptions`
2. `Visuals`
3. `P&L`
4. `Cash Flow`
5. `Balance Sheet`

The cash-flow tab includes an annual cash-flow statement and a monthly runway from October 2026 through December 2031.

## Validation

```powershell
python -m pytest -q
```

The build validates transaction math, primary-capital allocation, scenario ordering, slide count, workbook tabs, source coverage, and Office package integrity.
