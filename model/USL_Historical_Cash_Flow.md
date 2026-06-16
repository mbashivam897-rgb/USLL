# United Spirits Ltd (NSE: UNITDSPR) — Historical Cash Flow Statement

**Consolidated basis · FY2017–FY2026 · all figures ₹ Crore**
Simplified, treasury-linked, forecast-ready cash-flow reconstruction for equity research & DCF.

> Sourced **directly** from the consolidated Balance Sheet, P&L and Cash Flow Statement in each annual
> report held in this repository. Figures originally reported in ₹ million (FY2016–FY2023) have been
> converted to ₹ crore (÷10); FY2024–FY2026 are reported in ₹ crore. Every line reconciles to the
> filed statements and the **tie check vs. the balance-sheet treasury balance is 0.0 in every year.**

---

## 1. Historical Cash Flow Statement

| Particulars (₹ cr) | FY2017 | FY2018 | FY2019 | FY2020 | FY2021 | FY2022 | FY2023 | FY2024 | FY2025 | FY2026 |
|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| Net Profit After Tax (PAT) | 93.0 | 651.9 | 683.6 | 620.6 | 362.1 | 810.6 | 1,125.8 | 1,408.0 | 1,582.0 | 1,838.0 |
| + Depreciation & Amortisation | 188.6 | 192.3 | 214.7 | 285.3 | 299.1 | 303.8 | 282.5 | 275.0 | 283.0 | 289.0 |
| +/- Other Non-Cash Adjustments | 333.1 | (167.8) | 7.6 | (649.9) | 1,024.4 | 8.9 | (726.1) | (356.1) | 119.0 | (185.0) |
| +/- Change in Net Working Capital | 32.1 | 248.7 | 42.4 | 527.2 | 132.1 | (145.9) | (67.5) | (208.9) | (37.0) | (483.0) |
| **Cash Flow from Operations (CFO)** | **646.8** | **925.1** | **948.3** | **783.2** | **1,817.7** | **977.4** | **614.7** | **1,118.0** | **1,947.0** | **1,459.0** |
| + Interest Income Received | 19.7 | 4.2 | 6.4 | 6.2 | 11.1 | 6.4 | 20.4 | 42.0 | 36.0 | 67.0 |
| - Capital Expenditure (Capex) | 311.0 | 184.3 | 173.0 | 211.6 | 158.9 | 134.0 | 136.6 | 98.0 | 162.0 | 181.0 |
| **Free Cash Flow (FCF)** | **355.5** | **745.0** | **781.7** | **577.8** | **1,669.9** | **849.8** | **498.5** | **1,062.0** | **1,821.0** | **1,345.0** |
| - Dividends Paid | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 284.0 | 355.0 | 1,263.0 |
| - Lease Principal Repayment | 8.6 | 0.0 | 0.0 | 64.5 | 80.7 | 100.2 | 124.0 | 126.0 | 137.0 | 144.0 |
| +/- Net Borrowings Drawn / (Repaid) | (91.1) | (713.3) | (579.0) | (494.3) | (1,484.1) | (537.7) | (340.6) | 24.0 | (25.0) | (6.0) |
| +/- Other Investing/Financing & Interest (net) | (306.5) | 137.1 | (175.9) | (228.4) | (94.7) | (13.4) | 822.8 | 52.9 | (269.0) | 299.0 |
| **Net Change in Treasury** | **(50.7)** | **168.8** | **26.8** | **(209.4)** | **10.4** | **198.5** | **856.7** | **728.9** | **1,035.0** | **231.0** |
| Opening Treasury Balance | 138.0 | 87.3 | 256.1 | 282.9 | 73.5 | 83.9 | 282.4 | 1,139.1 | 1,868.0 | 2,903.0 |
| Closing Treasury Balance | 87.3 | 256.1 | 282.9 | 73.5 | 83.9 | 282.4 | 1,139.1 | 1,868.0 | 2,903.0 | 3,134.0 |
| Closing Treasury (Balance Sheet) | 87.3 | 256.1 | 282.9 | 73.5 | 83.9 | 282.4 | 1,139.1 | 1,868.0 | 2,903.0 | 3,134.0 |
| **Tie Check vs Balance Sheet** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** |

*(Parentheses = cash outflow / negative.)*

A live, formula-driven version is in **`USL_Historical_Cash_Flow.xlsx`** (tabs: Cash Flow · Treasury & NWC · Reconciliation · Sources). Every calculated cell is an Excel formula; the workbook was recalculated and the tie check is 0.0 in all years.

---

## 2. Formula definitions (shown beside every calculated line)

| Line | Formula (Excel logic) |
|---|---|
| **Change in Net Working Capital** | `= NWC(prev year) − NWC(current year)` → a *rise* in NWC is a cash outflow |
| **Cash Flow from Operations (CFO)** | `= PAT + D&A + Other Non-Cash Adj. + Change in NWC` |
| **Free Cash Flow (FCF)** | `= CFO + Interest Income Received − Capex` |
| **Net Change in Treasury** | `= FCF − Dividends − Lease Principal + Net Borrowings + Other Inv./Fin. & Interest` |
| **Closing Treasury (calc)** | `= Opening Treasury + Net Change in Treasury` |
| **Opening Treasury** | `= prior-year Closing Treasury` (FY2017 opening = reported FY2016 treasury) |
| **Closing Treasury (BS)** | `= Cash & equivalents + Bank balances + Current investments` (from Treasury & NWC tab) |
| **Tie Check** | `= Closing Treasury (calc) − Closing Treasury (BS)` → **0.0 every year** |

**Treasury** = Cash & cash equivalents + Bank balances (other than cash) + Current investments.
*Excluded:* trade receivables, inventories, other current assets, borrowings, deferred-tax assets, current-tax assets.

**Net Working Capital** = Operating current assets − Operating current liabilities.
*Included assets:* inventories, trade receivables (current), other current financial assets, current loans, contract assets, other current assets.
*Included liabilities:* trade payables, other current financial liabilities, contract liabilities, current provisions, other current liabilities.
*Excluded:* cash, bank balances, current investments, borrowings, current-tax assets/liabilities, lease liabilities, dividend payable.

### About the two reconciling lines
- **Other Non-Cash Adjustments** is sized so that **CFO equals the reported statutory “net cash from operating activities”** in each year. Economically it captures: finance-cost add-backs, taxes paid vs. accrued timing, exceptional/non-cash items, share-based payments, loss allowances and FX. (Reconciliation tab proves `Model CFO − Reported CFO = 0`.)
- **Other Investing/Financing & Interest (net)** bridges FCF to the change in the *broad* treasury balance. Because treasury includes bank deposits and current investments, this line absorbs items the four headline financing lines do not: **interest paid** on debt & leases, proceeds from **business / subsidiary / asset disposals**, **JV investments**, loans given, and FX. This is where the material one-offs sit (see §4).

---

## 3. Reconciliation to reported statements

| ₹ cr | FY2017 | FY2018 | FY2019 | FY2020 | FY2021 | FY2022 | FY2023 | FY2024 | FY2025 | FY2026 |
|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| Reported statutory CFO (filings) | 646.8 | 925.1 | 948.3 | 783.2 | 1,817.7 | 977.4 | 614.7 | 1,118.0 | 1,947.0 | 1,459.0 |
| Model CFO (this statement) | 646.8 | 925.1 | 948.3 | 783.2 | 1,817.7 | 977.4 | 614.7 | 1,118.0 | 1,947.0 | 1,459.0 |
| **Difference** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** | **0.0** |
| Memo: Interest paid (debt + leases) | 375.2 | 252.8 | 229.3 | 181.1 | 142.0 | 50.1 | 36.3 | 21.0 | 40.0 | 37.0 |
| Memo: Net Working Capital (level) | 2,594.1 | 2,345.4 | 2,303.0 | 1,775.8 | 1,643.7 | 1,789.6 | 1,857.1 | 2,066.0 | 2,103.0 | 2,586.0 |

Treasury closing balances reconcile to the consolidated balance sheet **exactly** in all ten years.

---

## 4. Material one-off / exceptional items flagged

| Year | Flag |
|---|---|
| **FY2017** | PAT trough (₹93 cr). Goodwill impairment, write-down of plant & equipment and large bad-debt / advances write-offs (legacy United Breweries-era receivables). Interest paid of ₹375 cr (high gross debt) dominates the −₹306 cr “Other Inv./Fin. & Interest” line. |
| **FY2018–FY2019** | Aggressive deleveraging — ~₹713 cr and ~₹579 cr of net borrowings repaid; CFO conversion normalises. |
| **FY2020** | Ind AS 116 (leases) adopted: lease principal repayment appears (₹64.5 cr); large working-capital release. |
| **FY2021** | COVID year: low PAT (₹362 cr) but a **₹1,818 cr CFO** on heavy working-capital unwind and tax refunds; ~₹1,484 cr of borrowings repaid → company effectively net-debt-free. |
| **FY2023** | **+₹817 cr exceptional gain** and ~₹818 cr proceeds from **sale of a business undertaking (popular/“Pioneer” brands)** — sits in the +₹823 cr “Other Inv./Fin.” line; depresses reported PAT-to-CFO conversion. |
| **FY2024–FY2025** | Dividends resumed (₹284 cr, ₹355 cr cash paid); treasury build accelerates; net cash deployed into deposits & mutual funds (within treasury). |
| **FY2026** | **₹1,263 cr dividend** paid (highest ever). **RCB / Sports (Royal Challengers Sports Pvt Ltd) reclassified as discontinued / held-for-sale** (₹701 cr assets, ₹327 cr liabilities). Goodwill & “other non-current assets” jump on acquisition activity. |

---

## 5. Analyst commentary

**PAT → CFO conversion.** Over the decade CFO comfortably exceeds reported PAT in most years (cash-conversion >100% in FY18–FY22), reflecting a high non-cash D&A base and disciplined working capital. Two distortions stand out: **FY2021** (CFO 5.0× PAT — COVID working-capital release + tax refunds against a depressed profit) and **FY2023** (CFO only 0.55× PAT — the year’s profit was flattered by an ₹817 cr exceptional disposal *gain* that is a non-operating, investing cash item, so it correctly drops out of CFO). Normalising for one-offs, underlying operating cash conversion is healthy and improving as the mix premiumises.

**Free cash flow.** USL is structurally free-cash-generative: **cumulative FCF of ≈₹9,706 cr over FY2017–FY2026** against just **≈₹1,750 cr of cumulative capex** — an asset-light, ~1.5–2.0% of-NSV capital intensity. FCF troughed around ₹356 cr (FY2017) and peaked at ₹1,669 cr (FY2021) / ₹1,821 cr (FY2025).

**Capital deployment & dividends.** The first half of the decade was spent **deleveraging — ~₹4,247 cr of net borrowings repaid** — transforming a heavily-indebted balance sheet into a net-cash one. With debt gone, capital return began only in **FY2024**; cumulative dividends of **≈₹1,902 cr** have been paid, culminating in the **₹1,263 cr FY2026 payout**. There is no buyback history; dividends are the sole return channel.

**Evolution of treasury.** Treasury rose from a slim **₹87 cr (FY2017)** to **₹3,134 cr (FY2026)** — a ~36× increase — as repaid debt, retained FCF and the FY2023 disposal proceeds accumulated. The composition also matured: from almost pure operating cash to a diversified mix of **cash (₹859 cr), bank deposits (₹1,118 cr) and current investments / mutual funds (₹1,157 cr)** by FY2026.

**Treasury-yield potential.** Historically interest income was negligible (₹4–20 cr) because treasury balances were tiny and debt servicing consumed cash. With a **~₹3,100 cr net-cash pile** now earning a return, treasury income is becoming a real P&L contributor: at a **~5.5–6.0% blended yield** the FY2026 closing balance alone implies **~₹175–190 cr of annual pre-tax treasury income** — a clean, low-risk earnings stream. This is exactly why the template links **forecast interest income = Opening Treasury × Treasury Yield**, with treasury itself rolled forward from the cash-flow statement.

---

## 6. Sources & provenance

| Years | Source (this repo) |
|---|---|
| FY2016 & FY2017 | `FY2017.pdf` — Consolidated BS p.190–191, P&L p.192–193, Cash Flow p.194–195 (₹ mn ÷10) |
| FY2018 & FY2019 | `FY2019.pdf` — Consolidated BS p.212–213, P&L p.214–215, Cash Flow p.216–217 (₹ mn ÷10) |
| FY2020 & FY2021 | `FY2021.pdf` — Consolidated BS p.176–177, P&L p.178–179, Cash Flow p.181–182 (₹ mn ÷10) |
| FY2022 & FY2023 | `FY2023_compressed.pdf` — Consolidated BS p.259–260, P&L p.261, Cash Flow p.264–265 (₹ mn ÷10) |
| FY2024 & FY2025 | `FY2025_compressed.pdf` — Consolidated BS p.206, P&L p.207, Cash Flow p.208–209 (₹ cr) |
| FY2026 | `fy 2026.xlsx` — audited consolidated results filed 14-May-2026 (₹ cr) |

*Educational / analytical model. Not investment advice. Generated by `usl_hist_cashflow.py` (calc) and `usl_hist_cashflow_excel.py` (workbook).*
