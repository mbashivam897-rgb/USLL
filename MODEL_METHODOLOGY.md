# United Spirits Ltd (USL) — Debt Schedule & Cash Schedule

**File:** `USL_Debt_Cash_Schedule.xlsx` · **Basis:** Consolidated · **Units:** ₹ crore (1 crore = 10 million)
**History:** FY2022–FY2025 (Actuals) · **Forecast:** FY2026–FY2030 (Estimates)
**Sources:** USL Integrated Annual Report FY2022‑23 (reported in ₹ million → divided by 10) and FY2024‑25 (reported in ₹ crore).

---

## 1. The central insight

United Spirits is **structurally net‑cash / debt‑free**. The FY2022‑23 Annual Report states plainly: *"As the Company is debt‑free, exposure to interest rate risk is negligible."* The only material interest‑bearing liability on the balance sheet is **lease liabilities recognised under Ind AS 116**; contractual borrowings have been run down to ~nil.

Therefore the model treats interest in two halves:

| | Drives | Where |
|---|---|---|
| **Interest EXPENSE** | Lease liabilities + (negligible) borrowings | **Debt Schedule** tab |
| **Interest INCOME** | A large and growing treasury (cash + bank deposits + debt mutual funds) | **Cash Schedule** tab |

For USL, **net interest is an income, not an expense**, and is widening every year.

---

## 2. Workbook structure

| Tab | Purpose |
|---|---|
| **Cover** | Orientation and conventions |
| **Assumptions** | Every forecast driver (blue font = hard input) with a written justification in column L |
| **Historicals** | Reconstructed consolidated P&L and balance‑sheet items, FY22A–FY25A |
| **Debt Schedule** | Lease & borrowings roll‑forward → computes finance cost (interest expense) |
| **Cash Schedule** | Free‑cash‑flow build → treasury roll‑forward → computes interest income |
| **Output Summary** | Total debt, net debt, net interest, leverage & coverage ratios, equity roll |

**Colour key:** blue figures are hard‑coded inputs; black figures are live formulas. Historical actual columns (FY22A–FY25A) are hard‑coded from the annual reports; forecast columns (FY26E–FY30E) are 100 % formula‑driven from the Assumptions tab.

---

## 3. How interest is calculated

**Lease liability roll (no circularity):**
```
Closing lease = Opening + New lease additions − Principal repayment − Terminations
Lease interest = Lease rate × AVERAGE(Opening, Closing)
```
The closing balance does not depend on the interest charge, so the calculation is clean.
*Historical lease interest is set to the actual reported figure so total finance cost ties exactly to the P&L (variance = 0 in every actual year). The "rate × average" method is shown as a memo line and is used for all forecast years.*

**Borrowings roll:** Opening + Drawdowns − Repayments = Closing; interest = rate × average. Forecast borrowings = 0 (debt‑free policy), so no forecast interest expense.

**Total finance cost = Lease interest + Borrowing interest + Other interest** (the "interest – others" line in the finance‑cost note: deferred sales‑tax, MSME, unwinding of discounts, bank charges).

**Interest income (no circularity):** `Treasury yield × Opening treasury`. Using the opening balance avoids the income↔cash circular reference, so the workbook recalculates in any spreadsheet application.

---

## 4. Historical anchors (₹ crore)

| | FY22A | FY23A | FY24A | FY25A |
|---|---:|---:|---:|---:|
| Net sales (ex‑excise) | 9,712 | 10,612 | 11,321 | 12,069 |
| EBITDA (derived) | 1,608 | 1,416 | 2,000 | 2,236 |
| Total lease liabilities | 264 | 182 | 240 | 480 |
| Total borrowings | 342 | 1 | 25 | 0 |
| Finance cost (reported) | 88 | 104 | 76 | 89 |
| Treasury (cash+deposits+MF) | 282 | 1,139 | 1,868 | 2,903 |
| **Net debt / (net cash)** | **323** | **(956)** | **(1,603)** | **(2,423)** |

---

## 5. Forecast assumptions & justification

| Driver | FY26E→FY30E | Rationale |
|---|---|---|
| Net‑sales growth | 9.5% → 10% → 10% → 9% → 9% | USL reaffirmed double‑digit topline guidance; premiumization (Prestige & Above ≈ 90% of mix). India premium‑spirits value CAGR ≈ 12%, IMFL value ≈ 13%. Kept slightly conservative for state‑excise / regulatory risk. |
| EBITDA margin (% net sales) | 19.0% → 21.0% | Mix premiumization + operating leverage; H1 FY26 reported EBITDA up ~31%. Builds on FY25 ≈ 18.5%. |
| Effective tax rate | 25.5% | Normalised Indian corporate rate. |
| Lease borrowing rate | 8.5% | Ind AS 116 incremental borrowing rate; consistent with implied FY24–25 lease interest. |
| New lease additions | ₹190–210 cr/yr | Normalised recurring warehouse/plant leases (FY25's ₹394 cr was a one‑off plant & equipment spike). |
| Borrowings | nil | Debt‑free policy; large undrawn working‑capital facilities (~₹1,300 cr). |
| Treasury yield | 6.75% | Indian short‑term debt / liquid‑fund yields. |
| Dividend payout | 30% of PAT | Up from ~22% in FY25, supported by strong cash generation. |
| Capex | ₹175 → ₹250 cr | Asset‑light model; capacity, premiumization and supply‑restructuring. |
| Working‑capital drag | 5% of incremental net sales | Inventory/receivables investment as the business grows. |

---

## 6. Headline outputs (₹ crore)

| | FY26E | FY27E | FY28E | FY29E | FY30E |
|---|---:|---:|---:|---:|---:|
| EBITDA | 2,511 | 2,835 | 3,198 | 3,573 | 3,990 |
| Total debt (all leases) | 530 | 560 | 580 | 600 | 620 |
| Treasury | 3,995 | 5,253 | 6,708 | 8,379 | 10,315 |
| Net cash | (3,466) | (4,693) | (6,128) | (7,779) | (9,695) |
| Interest expense | 90 | 91 | 92 | 92 | 92 |
| Interest income | 196 | 270 | 355 | 453 | 566 |
| **Net interest (income)** | **(106)** | **(178)** | **(263)** | **(361)** | **(474)** |
| Net debt / EBITDA | (1.4x) | (1.7x) | (1.9x) | (2.2x) | (2.4x) |
| Interest coverage (EBITDA/int exp) | 28x | 31x | 35x | 39x | 43x |

---

## 7. How to flex the model
Change any blue input on the **Assumptions** tab and the Debt Schedule, Cash Schedule and Output Summary recalculate automatically. Key sensitivities: net‑sales growth and EBITDA margin (drive cash and therefore interest income), the lease borrowing rate (drives interest expense) and the treasury yield (drives interest income).

> Industry figures referenced above were sourced from publicly available market research and company disclosures and have been paraphrased/summarised. *Content was rephrased for compliance with licensing restrictions.*
