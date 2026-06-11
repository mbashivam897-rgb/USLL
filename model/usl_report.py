"""Generate the full equity-research report (markdown) with model-linked tables."""
import pickle, numpy as np
st=pickle.load(open("build_state.pkl","rb"))
M=st["M"];R=st["R"];CF=st["CF"];FC=st["FC"];A=st["A"];comps=st["comps"];sens=st["sens"]
base_dcf=st["base_dcf"];base_rel=st["base_rel"];hist_val=st["hist_val"]
S_base=st["S_base"];S_bull=st["S_bull"];S_bear=st["S_bear"];CMP=st["CMP"]
HIST=["FY22","FY23","FY24","FY25","FY26"];FCST=["FY27E","FY28E","FY29E","FY30E"];YEARS=HIST+FCST
SH=72.78

def n(x,d=0):
    if x is None or (isinstance(x,float) and x!=x): return ""
    return f"{x:,.{d}f}"
def p(x,d=1):
    if x is None or (isinstance(x,float) and x!=x): return ""
    return f"{x*100:.{d}f}%"
def xx(x,d=1):
    if x is None or (isinstance(x,float) and x!=x): return ""
    return f"{x:.{d}f}x"

def tbl(headers, rows):
    out="| "+" | ".join(headers)+" |\n"
    out+="|"+"|".join(["---"]+["--:"]*(len(headers)-1))+"|\n"
    for r in rows: out+="| "+" | ".join(str(c) for c in r)+" |\n"
    return out+"\n"

def yrow(label, d, fmt=n, yrs=YEARS, d2=0):
    return [label]+[fmt(d.get(y),d2) if isinstance(d,dict) else "" for y in yrs]

L=[]
def w(s=""): L.append(s)

# ============ HEADER ============
w("# United Spirits Limited (USL) — Equity Research Initiation")
w("### NSE: UNITDSPR | BSE: 532432 | Sector: Beverages (Spirits) | Consolidated | FY2022–FY2030")
w("")
w("> **Rating: NEUTRAL / HOLD**  |  **CMP (10-Jun-2026): ₹1,259**  |  **12-m Target: ₹1,160 (−8%)**  |  Bull ₹1,500 / Bear ₹770")
w(">")
w("> Currency ₹ crore unless stated. Sources: USL Annual Reports FY2023 & FY2024-25 (consolidated), FY2026 consolidated results (filed 14-May-2026); market data ~10-Jun-2026. **Educational model — not investment advice.**")
w("")
w("**One-line view:** USL is the highest-quality, market-leading play on India's structural spirits-premiumisation theme, with a fortress balance sheet (net cash, ~26% RoCE) and a multi-year margin-expansion runway. But after a powerful re-rating, the stock already capitalises much of that optimism — our blended fair value of **₹1,160** implies the risk-reward is broadly balanced-to-slightly-negative. We would turn constructive below ~₹1,050.")
w("")
w("---")

# ============ STEP 1: BUSINESS UNDERSTANDING ============
w("## Step 1 — Business Understanding")
w("")
w("**Business model.** United Spirits Limited (USL) is India's largest beverage-alcohol (\"alcobev\") company and the local subsidiary of **Diageo plc** (~56% ownership). It manufactures, sources, markets and sells Indian-Made Foreign Liquor (IMFL) — whisky, brandy, rum, vodka and gin — across the full price ladder. The economic engine is **premiumisation**: shifting the mix from low-margin \"Popular\" brands to high-margin \"Prestige & Above\" (P&A) brands, and trading consumers up *within* brands and categories.")
w("")
w("**Revenue segments.** Post-FY2026 the group reports two segments: **(1) Beverage alcohol** — the core franchise (~96% of net sales), and **(2) Sports** — the Royal Challengers Bengaluru (RCB) IPL/WPL franchise held via Royal Challengers Sports Pvt Ltd. In FY2026 the Sports business was **classified as a discontinued operation / held-for-sale**, so reported continuing operations are effectively the pure beverage-alcohol business. Within beverages, the company manages two tiers:")
w("- **Prestige & Above (P&A):** the growth/margin driver — ~**90% of net sales** and ~84% of volume in FY26 (vs 70% of NSV in FY20). Sub-segments: Lower/Mid/Upper Prestige, Premium and Luxury.")
w("- **Popular:** mass-market, regional; deliberately de-emphasised. In FY2023 USL franchised/slump-sold 32 brands (incl. 11 popular brands) to sharpen focus on P&A.")
w("")
w("**Key brands.** Lower-to-upper Prestige: *McDowell's No.1* (>25 mn cases p.a.), *Royal Challenge, Signature, Antiquity, Royal Challenge American Pride*. Premium & Luxury (incl. imported Diageo portfolio): *Black & White, Black Dog, Johnnie Walker, Smirnoff, Singleton, Tanqueray, Don Julio*. USL has ~**63 brands**, of which **8 sell over 1 million cases** annually.")
w("")
w("**Distribution network.** Pan-India reach via **>70,000 retail outlets**. India's alcobev distribution is state-controlled: in many states USL sells to **government corporations** (e.g., TASMAC, APSBCL, Kerala/Karnataka beverage corps) which then control wholesale/retail; elsewhere it uses private wholesalers or auction/retail models. A **single state customer accounts for ~30% of revenue** — a structural concentration risk, but typical for the industry.")
w("")
w("**Manufacturing model.** Asset-light and distributed: ~**40 facilities** combining owned distilleries/bottling units with **tie-up/contract manufacturing units (TMUs)** and franchisee bottlers. State-level taxation and inter-state movement restrictions force largely *in-state* production, hence the dispersed footprint. The ongoing multi-year **\"supply agility programme\"** is optimising this footprint (recognised as exceptional items).")
w("")
w("**Industry structure.** India is the world's largest whisky market by volume and among the fastest-growing spirits markets. The IMFL market is an oligopoly: **USL (Diageo)** and **Pernod Ricard India** are the two majors, followed by **Allied Blenders & Distillers (ABDL), Radico Khaitan, Tilaknagar Industries** and **Globus Spirits**, plus regional players. USL is the volume leader; Pernod leads premium whisky value via Royal Stag/Blenders Pride/Imperial Blue (the last sold to Tilaknagar in 2025).")
w("")
w("**Key competitors & market-share trend.** USL remains the **#1 spirits company by volume**, but faces share pressure in mid-premium whisky from Pernod, ABDL and a fast-premiumising Radico. USL's strategy is value share over volume share — accepting some volume cession in Popular while compounding P&A value. Volume fell from ~79 mn cases (FY22) to ~64 mn (FY25) largely due to the deliberate Popular exit; net sales nonetheless rose every year on premiumisation.")
w("")
w("**Regulatory environment.** Alcohol is a **State subject** in India (outside GST). Each state independently controls licensing, pricing approvals, distribution route-to-market (RTM), label registration, and **excise duty** — creating 30+ distinct regulatory regimes. Risks include prohibition (e.g., Gujarat, Bihar), sudden RTM changes (e.g., the Delhi policy reversals), and periodic 'dry days'/advertising bans (alcohol advertising is prohibited, so brands rely on surrogate advertising).")
w("")
w("**Alcohol taxation structure.** Excise duty is the dominant levy and is **collected from consumers and remitted to states** — it is a pass-through that inflates gross revenue but is economically neutral to USL. In FY26, excise (₹15,349 cr) was **~55% of gross revenue**; the company and analysts therefore track **Net Sales Value (NSV = revenue net of excise)** as the true topline. States also levy additional duties, label/brand-registration fees, and litre/case-based special fees. Because state-set Maximum Retail Prices often lag input inflation, **pricing power is constrained** — USL must negotiate price increases state-by-state.")
w("")
w("**State-wise pricing & taxation impact.** Realisation and margins vary widely by state depending on the tax model (ad-valorem vs specific duty) and RTM. Favourable pricing actions in several states delivered ~₹243 cr of pricing benefit in FY25; conversely, adverse RTM shifts (Delhi, Andhra Pradesh, etc.) can dent volumes and mix in any given year. This makes the topline a **portfolio of 30+ state outcomes**, smoothing company-level volatility but capping the speed of price recovery.")
w("")
w("**Industry growth drivers.** (i) Low per-capita consumption + a large, young, legal-drinking-age population; (ii) **premiumisation** — consumers \"drinking better, not more\"; (iii) rising urbanisation and disposable incomes; (iv) growing female participation and on-premise (pubs/bars) culture; (v) conversion from country liquor to branded IMFL; (vi) white-spirits and ready-to-drink expansion.")
w("")
w("**Key risks.** (i) Adverse state taxation/excise hikes not passed through; (ii) RTM disruptions or prohibition; (iii) input-cost inflation (Extra Neutral Alcohol/grain, glass, packaging); (iv) intensifying premium competition; (v) customer concentration (state corporations & receivables); (vi) advertising restrictions; (vii) ESG/health-policy headwinds.")
w("")

# ============ STEP 2: HISTORICAL ANALYSIS ============
w("## Step 2 — Historical Analysis (FY2022–FY2026)")
w("")
w("### 2.1 Revenue analysis")
w(tbl(["₹ crore / metric"]+HIST,[
 yrow("Volume (mn cases)",M["volume_mn_cases"],n,HIST,1),
 yrow("NSV per case (₹, realisation)",M["nsv_per_case"],n,HIST),
 yrow("Net Sales Value (NSV)",M["net_revenue"],n,HIST),
 ["NSV growth %","","+9.3%","+6.7%","+6.6%","+3.3%"],
 yrow("P&A salience (% of NSV)",M["pa_nsv_salience"],n,HIST),
 yrow("P&A salience (% of volume)",M["pa_vol_salience"],n,HIST),
 yrow("Gross revenue (incl. excise)",M["gross_revenue"],n,HIST),
 yrow("Excise duty",M["excise_duty"],n,HIST),
]))
w("**Read-through.** USL's volume *declined* over FY22–FY25 (79→64 mn cases) yet NSV *grew every year* — the clearest evidence of the premiumisation engine. The FY23 portfolio reshape (slump-sale of 32 brands + franchising of 11 Popular brands) cut ~8% of volume but lifted P&A's NSV salience from 84%→87% and realisation from ₹1,459→₹1,605/case. By FY26, P&A is ~90% of NSV and realisation is ~₹1,872/case. Topline growth is now **realisation-led** (premium mix + pricing) rather than volume-led.")
w("")
w("### 2.2 Profitability analysis (core beverage)")
w(tbl(["Margin (%)"]+HIST,[
 yrow("Gross margin",R["gross_margin"],p,HIST),
 yrow("EBITDA margin",R["bev_ebitda_margin"],p,HIST),
 yrow("EBIT margin",R["ebit_margin"],p,HIST),
 yrow("PAT margin (reported)",R["pat_margin"],p,HIST),
]))
w("EBITDA margin expanded from **16.0% (FY22) → 18.3% (FY26)** — ~230 bps — despite a 2022-23 input-cost spike (glass, ENA) that compressed gross margin in FY23 (41.5%). Premiumisation + the supply-agility cost programme + operating leverage drove the recovery. PAT margins are flattered in FY23 (one-off ₹176 cr exceptional gain on the brand sale) and FY26 (₹129 cr discontinued-ops gain).")
w("")
w("### 2.3 Cost structure (% of NSV)")
w(tbl(["% of NSV"]+HIST,[
 yrow("COGS (raw material + packaging + traded goods)",R["cogs_pct"],p,HIST),
 yrow("Employee cost",R["emp_pct"],p,HIST),
 yrow("Advertising & promotion (A&P)",R["adp_pct"],p,HIST),
 yrow("Other operating expense (incl. distribution)",R["othexp_pct"],p,HIST),
]))
w("**Cost drivers.** *COGS* (~53-56% of NSV) is dominated by ENA/grain spirit, scotch concentrate, and glass/packaging; it spiked in FY23 on commodity inflation, then eased. *A&P* has risen structurally (7.2%→10.4% of NSV) as USL invests behind premium/luxury brand-building — a deliberate, margin-accretive investment because P&A carries far higher gross margins. *Employee cost* has been held flat-to-down as a share of NSV (~5%) via productivity (~₹388 cr enterprise-productivity savings in FY25). *Other operating expense* (freight, warehousing, royalties, overheads) trends with scale and the leased manufacturing footprint.")
w("")

# ============ STEP 3: WORKING CAPITAL ============
w("## Step 3 — Working Capital Analysis")
w(tbl(["Metric"]+HIST,[
 yrow("Trade receivables (current, ₹ cr)",M["receivables_cur"],n,HIST),
 yrow("Inventory (₹ cr)",M["inventory"],n,HIST),
 yrow("Trade payables (₹ cr)",M["trade_payables"],n,HIST),
 yrow("Net trade WC (₹ cr)",R["nwc"],n,HIST),
 yrow("Receivable days (on gross rev)",R["recv_days"],n,HIST,1),
 yrow("Inventory days (on COGS)",R["inv_days"],n,HIST,1),
 yrow("Payable days (on COGS)",R["pay_days"],n,HIST,1),
 yrow("Cash conversion cycle (days)",R["ccc"],n,HIST,1),
]))
w("**Formulae:** Receivable days = Trade receivables ÷ **gross revenue** × 365 (receivables embed excise billed to state corporations, so gross is the correct denominator); Inventory days = Inventory ÷ COGS × 365; Payable days = Trade payables ÷ COGS × 365; CCC = Inventory + Receivable − Payable days.")
w("")
w("**Efficiency read.** Inventory is structurally high (130–145 days) because premium whisky requires **maturation** (scotch/grain spirit ageing). Receivables (~45–48 days on gross) reflect dependence on **state beverage corporations**, whose payment cycles USL cannot fully control — a key reason working capital ties up cash. Payables (~125–130 days) partly offset this. The CCC drifted up into FY26 as inventory built for the premium ramp; management's disclosed **working-capital-to-NSV ratio held at ~14–17%**. This is a working-capital-intensive model — the principal drag on free-cash conversion.")
w("")

# ============ STEP 4: FIXED ASSET ============
w("## Step 4 — Fixed Asset (PPE) Analysis")
w("USL is **asset-light** — much production runs through tie-up/contract units, so owned PPE is modest and capex historically tiny (₹100–180 cr p.a.). The gross-block roll-forward (consolidated) from the notes:")
w("")
w(tbl(["PPE gross block (₹ cr)","FY24","FY25"],[
 ["Opening gross block","2,012","1,893"],
 ["  + Additions","126","126"],
 ["  − Disposals","(47)","(41)"],
 ["  − Transfer to investment property","(198)","(39)"],
 ["Closing gross block","1,893","1,939"],
 ["Accumulated depreciation (closing)","1,049","1,089"],
 ["**Net PPE**","**844**","**850**"],
 ["PPE depreciation charge","121","113"],
]))
w(tbl(["Metric"]+HIST,[
 yrow("Net PPE incl. ROU (₹ cr)",M["ppe"],n,HIST),
 yrow("Capex (₹ cr)",{y:M["capex"][y] for y in HIST},n,HIST),
 yrow("Capex / NSV %",{y:R["capex_sales"][y] for y in HIST},p,HIST),
 yrow("Depreciation / NSV %",{y:R["dep_rate"][y] for y in HIST},p,HIST),
 yrow("Asset turnover (NSV / Net PPE)",{y:R["asset_turn"][y] for y in HIST},n,HIST,1),
]))
w("**Useful lives** (per notes): Buildings 5–60y; Plant & equipment 5–15y; Furniture 10y; Office equipment 3–5y; Vehicles 5y; computer software 5y; the RCB **franchise rights amortised over 50 years/IPL seasons**. **Asset turnover is very high (~14–15x NSV/Net PPE)** — the hallmark of the asset-light model. **Capex forecast:** we model ₹270–340 cr p.a. (~2% of NSV) — a modest step-up over history for premium/luxury maturation capacity, white-spirits and selective in-house bottling, plus maintenance; the bulk of capacity remains leased/tolled (captured in the lease schedule below).")
w("")

# ============ STEP 5: LEASE SCHEDULE ============
w("## Step 5 — Lease Accounting Schedule (Ind AS 116)")
w("Leases are material for USL because tie-up manufacturing and warehousing carry **in-substance fixed payments** capitalised as right-of-use (ROU) assets / lease liabilities. Historical roll-forward (notes) + forecast:")
w("")
w(tbl(["ROU asset (₹ cr)","FY24","FY25"],[
 ["Opening","173","227"],
 ["  + Additions / adjustments","184","394"],
 ["  − Depreciation","(129)","(147)"],
 ["  − Terminations / transfers","(1)","(17)"],
 ["Closing ROU asset","227","457"],
]))
w(tbl(["Lease liability (₹ cr)"]+FCST,[
 yrow("Opening",{y:M["lease_liab"][YEARS[YEARS.index(y)-1]] for y in FCST},n,FCST),
 yrow("  + New leases",{y:M["rou_add"][y] for y in FCST},n,FCST),
 yrow("  + Interest accretion",{y:M["lease_int"][y] for y in FCST},n,FCST),
 yrow("  − Lease payments",{y:-M["lease_pay"][y] for y in FCST},n,FCST),
 yrow("Closing lease liability",{y:M["lease_liab"][y] for y in FCST},n,FCST),
 yrow("ROU asset — closing",{y:M["rou_fc"][y] for y in FCST},n,FCST),
]))
w("**Assumptions:** new leases ~1.8–2.0% of NSV (capacity growth), ROU depreciation ~32% (short ~3-yr TMU leases), incremental borrowing rate ~8.5%. FY26 total lease cash outflow was ~₹144 cr (principal) + interest. Lease liabilities stay small relative to EBITDA (<0.3x), so leverage remains negligible even on a post-IFRS-16 basis.")
w("")

# ============ STEP 6: DEBT ============
w("## Step 6 — Debt Schedule")
w(tbl(["₹ crore"]+HIST,[
 yrow("Gross borrowings",M["borrowings"],n,HIST),
 yrow("Lease liabilities",M["lease_liab"],n,HIST),
 yrow("Total debt (incl. leases)",{y:M["borrowings"][y]+M["lease_liab"][y] for y in HIST},n,HIST),
 yrow("Total cash + treasury",{y:M["cash"][y]+M["bank_deposits"][y]+M["cur_investments"][y] for y in HIST},n,HIST),
 yrow("Net debt / (net cash)",{y:R["net_debt"][y] for y in HIST},n,HIST),
]))
w("**Why USL runs near-zero leverage.** USL deleveraged from **~₹3,400 cr of gross debt (FY22)** to virtually nil by FY25-26, repaying working-capital borrowings out of operating cash. Reasons: (i) strong, self-funding cash generation; (ii) **Diageo parentage** removes any need for external gearing and imposes conservative treasury discipline; (iii) high RoCE means reinvestment needs are low and surplus is returned as dividends. The effective interest cost is now dominated by **lease interest (~8.5%)**; conventional debt interest is immaterial. The company is firmly **net cash (~₹2,700 cr)**.")
w("")

# ============ STEP 7: CASH & INVESTMENTS ============
w("## Step 7 — Cash & Investments (Treasury) Schedule")
w(tbl(["₹ crore"]+YEARS,[
 yrow("Cash & equivalents",M["cash"]),
 yrow("Bank deposits",M["bank_deposits"]),
 yrow("Current investments (debt MF/bonds)",M["cur_investments"]),
 yrow("Total treasury",{y:M["cash"][y]+M["bank_deposits"][y]+M["cur_investments"][y] for y in YEARS}),
 yrow("Other income (treasury, ₹ cr)",M["other_income"]),
]))
w("**Treasury income yield = interest/MF income ÷ average treasury balance.** Realised yields run ~5.5–6% on a mix of bank term deposits and debt mutual funds. As the cash pile compounds (we model net retention after a ~60% payout), treasury income becomes a growing, low-risk contributor to other income (₹260–375 cr over the forecast). Note FY26 other income (₹478 cr) was elevated by one-off asset-disposal gains; we normalise this in the forecast.")
w("")

# ============ STEP 8: FORECAST DRIVERS ============
w("## Step 8 — Forecast Drivers (FY2027E–FY2030E)")
w("> FY2026 is the **last actual** (results filed 14-May-2026). The forecast window is **FY27E–FY30E**. We forecast the **continuing beverage-alcohol business**; RCB/Sports is treated as discontinued/held-for-sale (monetisation = separate optionality).")
w("")
w(tbl(["Driver"]+FCST,[
 yrow("Volume growth %",A["vol_growth"],p,FCST),
 yrow("Realisation (NSV/case) growth %",A["real_growth"],p,FCST),
 ["NSV growth % (≈ vol × price)"]+[p((1+A["vol_growth"][y])*(1+A["real_growth"][y])-1) for y in FCST],
 yrow("Volume (mn cases)",{y:M["volume_mn_cases"][y] for y in FCST},n,FCST,1),
 yrow("NSV per case (₹)",{y:M["nsv_per_case"][y] for y in FCST},n,FCST),
 yrow("Net Sales Value (₹ cr)",{y:M["net_revenue"][y] for y in FCST},n,FCST),
]))
w("**Assumptions & justification.**")
w("- **Volume +4.0–4.5%:** India IMFL/P&A volumes growing mid-single-digit; USL has lapped the Popular-exit drag, so reported volume now grows with the premium portfolio. Conservative vs the high-single-digit P&A category growth, reflecting continued Popular attrition.")
w("- **Realisation +4.8–5.3%:** continued **premiumisation** (P&A salience drifting toward ~92% of NSV) plus state-level price increases. Below the FY22–26 realisation CAGR (~6.4%) to be prudent on pricing approvals.")
w("- **Net result: NSV ~+10% p.a. in FY27–28, fading to ~+9%** — consistent with management's stated medium-term ambition of *sustained double-digit topline growth*.")
w("- **Urban consumption & market share:** premium concentrated in urban/metro and on-premise; we assume USL holds value share in P&A, ceding some Popular volume.")
w("")

# ============ STEP 9: MARGIN FORECAST ============
w("## Step 9 — Margin Forecast")
w(tbl(["Margin %"]+FCST,[
 yrow("Gross margin",{y:R["gross_margin"][y] for y in FCST},p,FCST),
 yrow("EBITDA margin (core)",{y:R["bev_ebitda_margin"][y] for y in FCST},p,FCST),
 yrow("EBIT margin",{y:R["ebit_margin"][y] for y in FCST},p,FCST),
 yrow("A&P (% of NSV)",{y:R["adp_pct"][y] for y in FCST},p,FCST),
]))
w("EBITDA margin expands ~**50–60 bps per year (18.3%→20.5%)**, driven by: (i) **premiumisation** lifting gross margin (47.0%→48.2% of NSV); (ii) **operating leverage** on employee/overhead as NSV compounds; (iii) **supply-agility productivity**. This is partly *re-invested* in **A&P (~10.2–10.3% of NSV)** to sustain premium brand-building. We stop short of the >22% margins implied by bulls, since pricing remains state-gated and competition is intensifying. Management's earlier *\"mid-high-teens EBITDA margin\"* goal has already been achieved; we model continued, measured expansion toward the low-20s.")
w("")

# ============ STEP 10: INTEGRATED STATEMENTS ============
w("## Step 10 — Integrated Financial Statements")
w("### 10.1 Income Statement (₹ cr; NSV basis, continuing operations)")
ebit={y:(M["bev_ebitda"][y]-M["dep_amort"][y]) for y in YEARS}
w(tbl(["₹ crore"]+YEARS,[
 yrow("Net Sales Value (NSV)",M["net_revenue"]),
 yrow("COGS",M["cogs"]),
 yrow("Gross Profit",M["bev_gp"]),
 yrow("Employee cost",M["employee"]),
 yrow("A&P",M["ad_promo"]),
 yrow("Other operating expense",M["other_expense"]),
 yrow("EBITDA (core)",M["bev_ebitda"]),
 yrow("Depreciation & amortisation",M["dep_amort"]),
 yrow("EBIT",ebit),
 yrow("Other income",M["other_income"]),
 yrow("Finance costs",M["finance_cost"]),
 yrow("Exceptional / JV",{y:M["exceptional"][y]+M["jv_share"][y] for y in YEARS}),
 yrow("PBT",M["pbt"]),
 yrow("Tax",M["tax"]),
 yrow("PAT (continuing)",M["pat_continuing"]),
 yrow("PAT (reported)",M["pat_reported"]),
 yrow("EPS (₹)",M["eps_reported"],n,YEARS,1),
]))
w("### 10.2 Balance Sheet (₹ cr)")
w(tbl(["₹ crore"]+YEARS,[
 yrow("Net PPE (incl. ROU)",M["ppe"]),
 yrow("Goodwill + Intangibles + Inv. property",{y:M["goodwill"][y]+M["intangibles"][y]+M["inv_property"][y]+M["cwip"][y] for y in YEARS}),
 yrow("Inventories",M["inventory"]),
 yrow("Trade receivables (cur+non-cur)",{y:M["receivables_cur"][y]+M["receivables_nc"][y] for y in YEARS}),
 yrow("Cash + deposits + investments",{y:M["cash"][y]+M["bank_deposits"][y]+M["cur_investments"][y] for y in YEARS}),
 yrow("Other assets (incl. tax under protest)",{y:M["other_assets"][y]+M["other_fin_assets"][y]+M["dta_net"][y]+M["cur_tax_assets"][y]+M["jv_invest"][y]+M["assets_hfs"][y] for y in YEARS}),
 yrow("TOTAL ASSETS",M["total_assets"]),
 yrow("Total equity",M["equity"]),
 yrow("Borrowings + lease liabilities",{y:M["borrowings"][y]+M["lease_liab"][y] for y in YEARS}),
 yrow("Trade payables",M["trade_payables"]),
 yrow("Other liabilities & provisions",{y:M["dtl"][y]+M["provisions"][y]+M["cur_tax_liab"][y]+M["other_fin_liab"][y]+M["other_cur_liab"][y]+M["liab_hfs"][y] for y in YEARS}),
 yrow("TOTAL EQUITY & LIABILITIES",M["total_liab_eq"]),
 yrow("Balance check (≈0)",{y:M["total_assets"][y]-M["total_liab_eq"][y] for y in YEARS},n,YEARS,1),
]))
w("### 10.3 Cash Flow Statement (₹ cr) — ties to the change in treasury")
w(tbl(["₹ crore"]+FCST,[
 yrow("Operating cash flow (CFO)",{y:CF[y]["cfo"] for y in FCST},n,FCST),
 yrow("Investing cash flow (CFI, incl. RCB disposal)",{y:CF[y]["cfi"] for y in FCST},n,FCST),
 yrow("Financing cash flow (CFF)",{y:CF[y]["cff"] for y in FCST},n,FCST),
 yrow("Net change in cash & treasury",{y:CF[y]["net"] for y in FCST},n,FCST),
 yrow("Opening treasury",{y:CF[y]["t_open"] for y in FCST},n,FCST),
 yrow("Closing treasury (= opening + net)",{y:CF[y]["t_close"] for y in FCST},n,FCST),
]))
w("**All three statements are fully linked and the balance sheet balances to ₹0 every year.** Equity rolls forward by PAT less dividends (~60–62% payout); treasury is the balancing item and is deployed across cash, deposits and debt mutual funds.")
w("")

# ============ STEP 11: VALUATION ============
w("## Step 11 — Valuation")
w("### 11.1 DCF (FCFF)")
w(tbl(["₹ crore"]+FCST,[
 yrow("EBIT",{y:FC[y]["ebit"] for y in FCST},n,FCST),
 yrow("NOPAT (EBIT × (1−25.5%))",{y:FC[y]["ebit"]-FC[y]["tax"] for y in FCST},n,FCST),
 yrow("+ D&A",{y:FC[y]["da"] for y in FCST},n,FCST),
 yrow("− Capex",{y:-FC[y]["capex"] for y in FCST},n,FCST),
 yrow("− New leases",{y:-FC[y]["rou_add"] for y in FCST},n,FCST),
 yrow("− Δ Working capital",{y:-FC[y]["dwc"] for y in FCST},n,FCST),
 yrow("FCFF",{y:FC[y]["fcff"] for y in FCST},n,FCST),
]))
w(tbl(["DCF bridge","Value"],[
 ["WACC","10.25%"],["Terminal growth (g)","6.25%"],
 ["PV of explicit + fade FCFF (₹ cr)",n(base_dcf["pv_explicit"])],
 ["PV of terminal value (₹ cr)",n(base_dcf["pv_tv"])],
 ["**Enterprise Value (₹ cr)**",f"**{n(base_dcf['ev'])}**"],
 ["+ Net cash (₹ cr)",n(base_dcf["net_cash"])],
 ["+ RCB/held-for-sale optionality (₹ cr)",n(base_dcf["rcb"])],
 ["**Equity Value (₹ cr)**",f"**{n(base_dcf['equity_val'])}**"],
 ["**DCF value / share (₹)**",f"**{n(base_dcf['tp'])}**"],
]))
w("**WACC build:** risk-free ~6.5% (India 10-yr G-sec), equity risk premium ~6.0%, beta ~0.65 (defensive staple, net cash) → cost of equity ~10.4%; ~100% equity-funded → WACC **~10.25%**. We use a **2-stage DCF** (explicit FY27–30 + a premiumisation *fade* FY31–36 with NSV growth easing 8.5%→6.0% and EBITDA margin rising toward ~22.5%) before a Gordon terminal at **g = 6.25%** (India nominal long-run consumption growth).")
w("")
w("**Honest finding:** even on these supportive inputs, the DCF yields ~**₹740** — well below the ₹1,259 market price. A pure cash-flow valuation cannot justify USL's multiple because (a) high working-capital intensity depresses FCFF conversion and (b) the market is implicitly applying a **sub-10% discount rate and >7% perpetual growth** — i.e., capitalising a very long premiumisation runway. The DCF is therefore a *sanity floor*, not the primary lens.")
w("")
w("**Sensitivity — DCF target price (₹): WACC (rows) × terminal growth (cols)**")
sens_rows=[[idx]+[n(sens.loc[idx,c]) for c in sens.columns] for idx in sens.index]
w(tbl(["WACC \\ g"]+list(sens.columns),sens_rows))
w("### 11.2 Trading comparables (≈10-Jun-2026)")
crows=[]
for _,row in comps.iterrows():
    crows.append([row["Company"],n(row["MktCap"]),n(row["EV"]),n(row["NetRev"]),n(row["EBITDA"]),
                  xx(row["EV/EBITDA"]),xx(row["EV/Sales"]),xx(row["P/E"]),p(row["EBITDA_margin"])])
w(tbl(["Company","MktCap","EV","NetRev FY26","EBITDA","EV/EBITDA","EV/Sales","P/E","EBITDA mgn"],crows))
w("**Read:** USL trades at **~39x EV/EBITDA and ~50–53x P/E** — a *discount* to the fastest-growing premium peer **Radico Khaitan** (~46x / ~78x; +25% revenue growth FY26) but a *premium* to mass-market peers **ABDL** (~21x) and **Globus** (~9x; a lower-margin bulk-alcohol/IMIL play). The premium over mass peers is justified by USL's leadership, mix and ~26% RoCE; the discount to Radico reflects Radico's faster growth. Comps therefore frame USL as **fully but not egregiously valued**. *(Tilaknagar FY26 PAT normalised for Imperial Blue one-offs; ABDL figures approximate.)*")
w("")
w("### 11.3 Historical valuation")
w(tbl(["Metric"]+HIST,[
 yrow("Market cap (₹ cr)",{y:hist_val[y]["mcap"] for y in HIST}),
 yrow("P/E (x)",{y:hist_val[y]["pe"] for y in HIST},n,HIST,1),
 yrow("EV/EBITDA (x)",{y:hist_val[y]["ev_ebitda"] for y in HIST},n,HIST,1),
]))
w("USL's 5-year average P/E is **~60x**; the current ~50x sits **below** that average (the stock has de-rated ~21% over the past year), and EV/EBITDA has compressed from ~48x (FY25) to ~39x (FY26). On its *own* history, USL looks **modestly cheap** — the bull's strongest argument and the reason we rate HOLD rather than Reduce.")
w("")

# ============ STEP 12: INVESTMENT THESIS ============
w("## Step 12 — Investment Thesis")
w("**Key positives**")
w("- **Premiumisation compounding:** P&A is ~90% of NSV and still mixing up; realisation rising ~5%+ p.a. structurally.")
w("- **Market leadership + Diageo portfolio:** unrivalled brand stable across the price ladder; access to global luxury brands (Johnnie Walker, Don Julio, Tanqueray).")
w("- **Margin expansion runway:** EBITDA margin 18.3%→~20.5% via mix, operating leverage and supply-agility savings.")
w("- **Fortress balance sheet & cash generation:** net cash ~₹2,700 cr, ~26% RoCE, ~60% dividend payout, asset-light (~14x asset turnover).")
w("- **Optionality:** RCB monetisation (held-for-sale) could unlock a sizeable, un-modelled cash inflow.")
w("")
w("**Key risks**")
w("- **Taxation/regulation:** state excise hikes not passed through; prohibition; RTM disruptions (Delhi/AP precedents).")
w("- **Pricing lag:** state-set MRPs lag input inflation, capping margin recovery speed.")
w("- **Competition:** Pernod, Radico, ABDL and Tilaknagar (post-Imperial Blue) intensifying premium competition.")
w("- **Input inflation:** ENA/grain, glass, scotch-concentrate and packaging volatility.")
w("- **Concentration:** ~30% of revenue from one state customer; receivables tied to state corporations.")
w("- **Valuation:** rich multiple leaves limited margin of safety; de-rating risk if growth disappoints.")
w("")

# ============ STEP 13: SENSITIVITY / SCENARIOS ============
w("## Step 13 — Scenario Analysis (Bull / Base / Bear)")
w(tbl(["Metric","Bear","Base","Bull"],[
 ["FY30E Net Sales (₹ cr)",n(S_bear["nsv30"]),n(S_base["nsv30"]),n(S_bull["nsv30"])],
 ["FY30E EBITDA (₹ cr)",n(S_bear["ebitda30"]),n(S_base["ebitda30"]),n(S_bull["ebitda30"])],
 ["FY30E PAT (₹ cr)",n(S_bear["pat30"]),n(S_base["pat30"]),n(S_bull["pat30"])],
 ["FY30E EPS (₹)",n(S_bear["eps30"],1),n(S_base["eps30"],1),n(S_bull["eps30"],1)],
 ["NSV CAGR FY26-30",p(S_bear["nsv_cagr"]),p(S_base["nsv_cagr"]),p(S_bull["nsv_cagr"])],
 ["EPS CAGR FY26-30",p(S_bear["eps_cagr"]),p(S_base["eps_cagr"]),p(S_bull["eps_cagr"])],
 ["**Target price (₹)**",f"**{n(S_bear['tp'])}**",f"**{n(S_base['tp'])}**",f"**{n(S_bull['tp'])}**"],
 ["Upside/(downside) vs ₹1,259",p(S_bear["upside"]),p(S_base["upside"]),p(S_bull["upside"])],
]))
w("- **Bull (₹1,500, +19%):** volume +6%, realisation +6%, EBITDA margin →~22%, RCB monetised at a premium; WACC 9.75%/g 6.75%.")
w("- **Base (₹1,160, −8%):** volume +4–4.5%, realisation +5%, margin →20.5%; WACC 10.25%/g 6.25%.")
w("- **Bear (₹770, −39%):** volume +2.5%, realisation +4%, margin stalls ~18.8%, tax 26%, de-rating; WACC 11.5%/g 5%.")
w("")

# ============ STEP 14: OUTPUT / RECOMMENDATION ============
w("## Step 14 — Investment Recommendation")
w(f"**Rating: NEUTRAL / HOLD.  12-month target price: ₹1,160 (−8% vs CMP ₹1,259).**")
w("")
w("USL is a best-in-class, structurally advantaged compounder — the cleanest large-cap proxy for Indian spirits premiumisation, with a fortress balance sheet and improving returns. However, the investment case and the valuation case diverge: the **business** deserves a premium, but the **stock** at ~50x earnings / ~39x EV/EBITDA already discounts years of double-digit growth and margin expansion. Our blended fair value (60% relative multiples + 40% 2-stage DCF) of **₹1,160** sits just below the market, and the scenario range (₹770–₹1,500) is roughly symmetric around the current price with a slightly negative skew.")
w("")
w("**Action:** *Hold* core positions for the compounding and optionality (RCB monetisation, faster-than-modelled premiumisation). **Accumulate below ~₹1,050–1,100** (where the margin of safety opens up and the DCF/relative blend turns positive); **trim above ~₹1,450**. Key catalysts to watch: RCB sale crystallisation, state pricing actions, P&A volume momentum, and input-cost trajectory.")
w("")
w("---")
w("### Deliverables checklist (all in this report + `USL_Financial_Model.xlsx`)")
w("1. Forecast assumptions ✓ 2. Historical tables ✓ 3. Revenue build-up ✓ 4. Working-capital schedule ✓ 5. PPE schedule ✓ 6. Lease schedule ✓ 7. Debt schedule ✓ 8. Cash/treasury schedule ✓ 9. Income statement ✓ 10. Balance sheet ✓ 11. Cash-flow statement ✓ 12. DCF ✓ 13. Relative valuation ✓ 14. Sensitivity tables ✓ 15. Recommendation ✓")
w("")
w("*Model build: consolidated basis; FY26 = last actual; topline = Net Sales Value (ex-excise). All schedules are linked and the balance sheet balances every year. Assumptions are the analyst's estimates and should be stress-tested against quarterly disclosures.*")

open("USL_Equity_Research_Report.md","w").write("\n".join(L))
print("Wrote USL_Equity_Research_Report.md  (chars:", len("\n".join(L)), ")")
