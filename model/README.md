# United Spirits Limited (USL) — Integrated Financial Model & Initiation of Coverage

Fully-linked, auditable 3-statement equity-research model for **United Spirits Ltd (NSE: UNITDSPR)**, FY2022–FY2030, on a **consolidated** basis. Built bottom-up from the company's filings.

## Deliverables
| File | Description |
|---|---|
| `USL_Equity_Research_Report.md` | Full initiation report (Steps 1–14): business understanding, historical analysis, all schedules, integrated statements, DCF + relative + historical valuation, scenarios, recommendation. All tables auto-generated from the model. |
| `USL_Financial_Model.xlsx` | Auditable workbook — 14 tabs: Cover, Assumptions, Income Statement, Balance Sheet, Cash Flow, Revenue Buildup, Working Capital, PPE Schedule, Lease Schedule, Debt & Cash, DCF, Comps, Historical Valuation, Scenarios. |
| `usl_engine.py` | Reusable model engine: historical database + `run(assumptions)` → fully-linked forecast (IS/BS/CF). |
| `usl_build.py` | Scenario runner + valuation (2-stage DCF, FCFF, WACC, sensitivity, trading comps, historical valuation). Writes `build_state.pkl`. |
| `usl_excel.py` | Builds the Excel workbook from `build_state.pkl`. |
| `usl_report.py` | Builds the markdown report from `build_state.pkl`. |

## How to run
```bash
pip install numpy pandas openpyxl
python3 usl_build.py     # runs engine + valuation -> build_state.pkl (prints checks)
python3 usl_excel.py     # -> USL_Financial_Model.xlsx
python3 usl_report.py    # -> USL_Equity_Research_Report.md
```

## Key design choices
- **Topline = Net Sales Value (NSV)** = revenue net of excise duty (excise is a pass-through state levy ~55% of gross revenue). Gross revenue & excise shown as a memo.
- **FY2026 is the last actual** (results filed 14-May-2026); forecast window is **FY27E–FY30E**.
- **RCB / "Sports"** was classified by the company as a **discontinued operation / held-for-sale** in FY26, so the core forecast is the **Beverage-Alcohol** continuing business; RCB monetisation is treated as separate optionality (held conservatively at ₹2,500 cr in valuation).
- **Balance sheet balances to ₹0 every year by construction**: equity rolls forward (PAT − dividends); all operating items are driver-based; **treasury (cash + deposits + current investments) is the balancing item**; the cash-flow statement is derived from BS/IS movements.
- Source data: USL Annual Report FY2023 (FY22 & FY23, originally ₹ mn → ÷10), Integrated Annual Report FY2024-25 (FY24 & FY25, ₹ cr), FY2026 consolidated results (`fy 2026.xlsx`).

## Headline
- **Rating: NEUTRAL / HOLD.** CMP ₹1,259 (10-Jun-26) · 12-m base target **₹1,160 (−8%)** · Bull ₹1,500 / Bear ₹770.
- High-quality compounder (net cash, ~26% RoCE, premiumisation) but richly valued; accumulate < ~₹1,050, trim > ~₹1,450.

*Educational model for analytical purposes only — not investment advice. Assumptions are the analyst's estimates.*
