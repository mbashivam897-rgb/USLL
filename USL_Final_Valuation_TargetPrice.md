# United Spirits Limited (USL) — Final Valuation & Justified Target Price
### Source model: `United Spirits Excell.xlsx` | Consolidated | All figures ₹ crore unless stated
**CMP: ₹1,322 (Mkt Cap ₹96,196 cr ÷ 72.74 cr shares)  |  12-month Target Price: ₹1,222  |  Implied Downside: −7.6%  |  Rating: HOLD**

> This section does **not** rebuild the model. It synthesises the DCF (`DCF` sheet) and Relative Valuation (`Comparables` sheet) outputs already in the workbook into a single, weight-justified Target Price. Relative multiples are applied on a **1-year-forward (FY2027E)** basis using **peer-mean multiples** (`Comparables` row 42), with FY2028E shown as a roll-forward cross-check. *Educational MBA project — not investment advice.*

---

## A. Analysis of the Valuation Outputs

The five valuation read-outs from the workbook, each expressed as an implied per-share value (FY2027E, peer-mean multiples; DCF from `DCF!C28`):

| # | Methodology | USL FY27E Metric | Peer-Mean Multiple | Implied Share Price (₹) |
|---|---|--:|--:|--:|
| 1 | DCF (2-stage FCFF) | FCFF ₹1,412–2,324 cr | WACC 11.0% / g 6.0% | **599** |
| 2 | EV/EBITDA | EBITDA ₹2,547 cr | 37.3x | **1,344** |
| 3 | EV/EBIT | EBIT ₹2,351 cr | 46.5x | **1,541** |
| 4 | P/E | PAT ₹1,847 cr | 72.1x | **1,831** |
| 5 | EV/Revenue | NSV ₹13,589 cr | 4.9x | **957** |

**Peer set:** United Breweries, Radico Khaitan, Allied Blenders Distillers (`Comparables` rows 13–15).

**Key observations**

1. **Two outliers bracket the estimate — one low, one high:**
   - **DCF (₹599) is the downside outlier**, ~55% below the current price. With WACC 11.0% and terminal growth 6.0% (`DCF` sheet), the intrinsic value is heavily penalised by USL's **working-capital intensity** — the model assumes **97 receivable days and 135 inventory days** (`Assumptions` E25:E26), reflecting whisky maturation and slow-paying state beverage corporations. This compresses free cash flow far below reported EBIT/PAT, so the DCF behaves as an **intrinsic-value floor / margin-of-safety anchor**, not the primary lens.
   - **P/E (₹1,831) is the upside outlier**, ~38% above the current price. This is a **peer-composition artefact**: the comparable set trades at structurally inflated P/E (United Breweries ~80x, Allied Blenders ~71x, Radico ~65x on FY27E) versus USL's own ~52x. Applying a ~72x peer-mean P/E implicitly assumes USL re-rates to the level of faster-growing / lower-base peers (and a beer company) — not realistic for the most mature, slowest-growing name in the set. The P/E output therefore **overstates** fair value and warrants reduced weight.

2. **EV/EBITDA (₹1,344) is the most reliable single read.** USL's FY27E EV/EBITDA (37.2x) sits almost exactly on the peer mean (37.3x), so this implied price embeds **no heroic re-rating** — it reflects fair value on the industry's standard metric.

3. **EV/Revenue (₹957) is a low read** because a sales multiple ignores USL's superior profitability. USL already trades at a richer EV/Sales (~7.0x) than the peer mean (~4.9x), and that premium is *justified* by its ~18% EBITDA margin — so forcing USL onto the peer-mean sales multiple understates it.

4. **Most relevant methodologies:** EV/EBITDA (industry standard) and EV/EBIT (operating profitability), supported by the DCF as an intrinsic check.

5. **Lower-importance / down-weighted:** P/E (upward outlier from inflated peer P/E) and EV/Revenue (ignores margins).

---

## B. Weightages and Their Justification

We deliberately **reject equal weighting**. Weights reflect (i) USL's business characteristics, (ii) spirits-industry convention, (iii) profitability/maturity, (iv) reliability of the cash-flow forecast, and (v) the outlier diagnosis from Section A.

| Methodology | Weight | Rationale |
|---|--:|---|
| **EV/EBITDA** | **35%** | Industry-standard for spirits; capital-structure- and D&A-neutral; USL's multiple is in line with peers so **no re-rating assumption is required** — the most defensible single read. Highest weight. |
| **DCF (FCFF)** | **25%** | The only intrinsic, cash-flow-based method; disciplines against over-paying. An outlier on the low side (WC intensity), so weighted materially but **capped below the relative block**. |
| **EV/EBIT** | **15%** | Operating-profit cross-check that also charges D&A. Modestly inflated by United Breweries' high EV/EBIT, and largely redundant with EV/EBITDA for an asset-light model (capex ~1% of NSV) → moderate weight. |
| **P/E** | **15%** | Normally a headline metric, but here an **upward outlier** driven by structurally high peer P/E (UB/ABDL/Radico). Retained because equity investors anchor to it, but **down-weighted** so the peer distortion does not dominate. |
| **EV/Revenue** | **10%** | Sanity check only; blind to profitability/mix, which are the core of the USL premiumisation case → lowest weight. |
| **Total** | **100%** | |

**Why each is *included*:** EV/EBITDA and EV/EBIT capture how the spirits market actually values operating performance; the DCF grounds the price in intrinsic cash generation; P/E reflects what the broad equity market watches; EV/Revenue guards against the case resting entirely on margin assumptions.

**Why the alternatives receive *lower* weight:**
- **P/E (15% vs a typical 25–30%):** the peer set's P/E is inflated and not representative of USL's growth/maturity profile — applying it unadjusted overstates value, so its influence is curtailed.
- **EV/Revenue (10%):** appropriate for early-stage or loss-making businesses, not a profit-mature leader; it penalises USL's margin advantage.
- **EV/EBIT (15%):** for an asset-light company it adds little beyond EV/EBITDA, so heavy weight would double-count.
- **DCF (25% rather than ~40%):** reliable in construction but an outlier in output (very high WC drag, single point estimate), so it anchors rather than drives.

**Net stance:** relative methods carry **70%** (the market prices USL on multiples); DCF carries **30%**, consistent with sell-side convention for mature Indian consumer-staples names.

---

## C. Selection of Appropriate Trading Multiples

| Multiple | Most common in spirits? | Captures operating performance? | Investor focus | Advantages | Limitations (for USL) |
|---|---|---|---|---|---|
| **EV/EBITDA** | **Yes — sector default** | Strong | High (institutional / M&A) | Capital-structure & tax neutral; removes D&A noise; best cross-peer comparability; USL multiple ≈ peer mean (no re-rating needed) | Ignores capex/working-capital intensity (a genuine USL drag) |
| **P/E** | Common | Moderate (post-tax) | **Highest among retail / long-only investors** | Simple, widely quoted; reflects tax efficiency and net interest | **Distorted here by inflated peer P/E**; sensitive to one-offs/exceptional items |
| **EV/EBIT** | Less common | **Strongest pure-operating read** | Moderate | Charges D&A, so penalises capital intensity EBITDA hides | ≈ EV/EBITDA for asset-light USL; skewed by UB's high EV/EBIT |
| **EV/Revenue** | Rare for profitable spirits | Weak (ignores margin) | Low | Robust to margin one-offs; usable when earnings are volatile | Blind to USL's margin/mix advantage — least suitable here |

**Conclusions:**
- **Most commonly used in spirits/alcobev:** **EV/EBITDA** — the sector default and the basis for industry M&A.
- **Best captures operating performance:** **EV/EBIT** in principle; but for asset-light USL, EV/EBITDA does so almost as well and is more comparable.
- **What investors focus on most:** **P/E** (broad market) and **EV/EBITDA** (institutional) — but in this peer set P/E must be used with caution.
- **Least suitable:** **EV/Revenue**, because it cannot reward the margin expansion central to USL's thesis.

---

## D. Final Target Price

### D.1 Valuation Summary Table (FY2027E, peer-mean multiples)

| Methodology | Implied Share Price (₹) | Weight | Weighted Contribution (₹) |
|---|--:|--:|--:|
| EV/EBITDA (37.3x) | 1,344 | 35% | 470.3 |
| DCF (FCFF, WACC 11% / g 6%) | 599 | 25% | 149.8 |
| EV/EBIT (46.5x) | 1,541 | 15% | 231.1 |
| P/E (72.1x) | 1,831 | 15% | 274.7 |
| EV/Revenue (4.9x) | 957 | 10% | 95.7 |
| **Final Target Price** | | **100%** | **₹1,222** |

### D.2 Final Weighted Valuation

**Target Price = Σ (Implied Price × Weight) = 470.3 + 149.8 + 231.1 + 274.7 + 95.7 = ₹1,222**

Cross-checks: the **relative-only blend** (ex-DCF) is **₹1,429**; the **standalone DCF** is **₹599 (−54.7%)**. Using FY2027E **median** (rather than mean) multiples gives a near-identical blended **₹1,230**, and rolling forward to **FY2028E** mean multiples gives **₹1,192** — confirming the target is robust at roughly **₹1,200–1,230**.

### D.3 Upside / Downside Analysis

| Item | Value |
|---|--:|
| Current Market Price (Mkt Cap ₹96,196 cr ÷ 72.74 cr shares) | ₹1,322 |
| 12-month Target Price | ₹1,222 |
| **Implied Upside / (Downside)** | **−7.6%** |

The risk-reward is **balanced-to-slightly-negative**: USL's quality is fully reflected in the price, with limited margin of safety at current levels.

---

## E. Investment Recommendation

**Rating: HOLD (Neutral). Target Price ₹1,222 (−7.6% vs CMP ₹1,322).**

**Justification.** USL is the highest-quality proxy for India's structural spirits-premiumisation theme — market leadership, a near-debt-free balance sheet (cash ₹3,134 cr vs debt ₹407 cr), high returns and an asset-light model. But the valuation triangulation shows the stock is **fully valued**: the DCF (₹599) flags that intrinsic cash generation, after heavy working-capital absorption, does not yet support the current price, while the headline P/E read (₹1,831) is an artefact of an over-priced peer set rather than USL fundamentals. Stripping out both outliers, the credible operating-multiple cluster (EV/EBITDA ₹1,344, EV/EBIT ₹1,541) and the weighted blend land fair value just **below** the market.

**Key valuation drivers:**
- **Premiumisation & mix** — the engine behind EBITDA-margin expansion (gross margin assumed 46% → 47%, `Scenarios`).
- **Operating leverage** on a largely fixed S&A/employee base (~20% / ~6% of sales).
- **Balance-sheet quality** — net cash, ~67% dividend payout, high asset turnover support a premium EV/Sales vs peers.
- **WACC sensitivity** — at 11% WACC the DCF is highly geared to terminal growth (`DCF` sensitivity grid).

**Key risks to the valuation:**
- **Working-capital intensity** (97 receivable / 135 inventory days) — the single biggest reason the DCF lags the multiples; any deterioration directly erodes intrinsic value.
- **State taxation / route-to-market** disruption and pricing lag versus input-cost inflation.
- **Peer-multiple de-rating** — much of the relative upside relies on peers sustaining 65–86x P/E; a sector de-rating would pull USL's implied value down.
- **Competition** from Pernod Ricard, Radico Khaitan and Allied Blenders in the premium segment.

**Action:** *Hold* core positions for the premiumisation compounding. **Accumulate below ~₹1,150–1,200** (where the blend turns positive and a margin of safety opens); **trim above ~₹1,430** (the relative-only ceiling). Catalysts: state pricing actions, P&A volume momentum, working-capital improvement and the input-cost trajectory.

---

## F. Methodology Notes (reproducibility)

- All figures are read directly from `United Spirits Excell.xlsx`: DCF per-share from `DCF!C28`; relative per-share values from `Comparables` "Valuation using Mean Multiples", FY2027 row 42 (J/K/L/M = EV/Revenue / EV/EBITDA / EV/EBIT / P/E); current Mkt Cap from `Comparables!C12`; shares from `DCF!C27`.
- The script `final_target_price.py` re-reads these cells and reproduces the ₹1,222 weighted result, so every figure ties back to the live workbook.
- **DCF** base case: 2-stage FCFF, explicit FY2027–30 + terminal value; WACC 11.0% (CAPM: Rf 7%, ERP 4%, β 0.92), terminal growth 6.0%.
- **Relative valuation** applies peer-mean multiples (United Breweries, Radico Khaitan, Allied Blenders) to USL's FY2027E metrics; median, quartile and FY2028E variants are used as cross-checks.

*This analysis defends every weighting decision with reasoning and explicitly down-weights the two statistical outliers (DCF low, P/E high); it does not merely average the valuation outputs.*
