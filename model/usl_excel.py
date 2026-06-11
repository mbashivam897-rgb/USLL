"""
Build the auditable Excel workbook USL_Financial_Model.xlsx from build_state.pkl.
"""
import pickle, numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter

st = pickle.load(open("build_state.pkl","rb"))
M=st["M"]; R=st["R"]; CF=st["CF"]; FC=st["FC"]; A=st["A"]
base_dcf=st["base_dcf"]; base_rel=st["base_rel"]; sens=st["sens"]; comps=st["comps"]
hist_val=st["hist_val"]; S_base=st["S_base"]; S_bull=st["S_bull"]; S_bear=st["S_bear"]; CMP=st["CMP"]
M_bull=st["M_bull"]; M_bear=st["M_bear"]

HIST=["FY22","FY23","FY24","FY25","FY26"]; FCST=["FY27E","FY28E","FY29E","FY30E"]; YEARS=HIST+FCST
SHARES=72.78

wb=Workbook()
# ---- styles ----
NAVY="1F3864"; BLUE="2E5496"; LBLUE="D9E1F2"; GREY="F2F2F2"; GREEN="E2EFDA"; GOLD="FFF2CC"
hdr=Font(bold=True,color="FFFFFF",size=11); title=Font(bold=True,color="FFFFFF",size=14)
boldf=Font(bold=True,size=10); ital=Font(italic=True,size=9,color="555555")
fillN=PatternFill("solid",fgColor=NAVY); fillB=PatternFill("solid",fgColor=BLUE)
fillL=PatternFill("solid",fgColor=LBLUE); fillG=PatternFill("solid",fgColor=GREY)
fillGr=PatternFill("solid",fgColor=GREEN); fillGo=PatternFill("solid",fgColor=GOLD)
thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
ctr=Alignment(horizontal="center"); rgt=Alignment(horizontal="right"); lft=Alignment(horizontal="left",wrap_text=True)

def sheet(name):
    ws=wb.create_sheet(name); ws.sheet_view.showGridLines=False; return ws

def put(ws,r,c,v,font=None,fill=None,fmt=None,align=None,bd=False):
    cell=ws.cell(row=r,column=c,value=v)
    if font:cell.font=font
    if fill:cell.fill=fill
    if fmt:cell.number_format=fmt
    if align:cell.alignment=align
    if bd:cell.border=border
    return cell

NUM="#,##0"; NUM1="#,##0.0"; PC="0.0%"; PCX="0.0%"; X="0.0\"x\""; RS="\"₹\"#,##0"

def table(ws, start_row, rows, years=YEARS, num_fmt=NUM, hdr_label="₹ crore", title_text=None,
          pct_rows=None, x_rows=None, one_dp_rows=None, bold_rows=None, shade_fcst=True):
    pct_rows=pct_rows or set(); x_rows=x_rows or set(); one_dp_rows=one_dp_rows or set(); bold_rows=bold_rows or set()
    r=start_row
    if title_text:
        put(ws,r,1,title_text,title,fillB); 
        for c in range(2,len(years)+2): put(ws,r,c,None,fill=fillB)
        r+=1
    # header
    put(ws,r,1,hdr_label,hdr,fillN,align=lft,bd=True)
    for j,y in enumerate(years):
        f=fillN
        put(ws,r,2+j,y,hdr,f,align=ctr,bd=True)
    r+=1
    for label,data in rows:
        is_bold = label in bold_rows
        put(ws,r,1,label,boldf if is_bold else None,fillL if is_bold else None,align=lft,bd=True)
        for j,y in enumerate(years):
            v=data.get(y,None) if isinstance(data,dict) else None
            fmt=PC if label in pct_rows else (X if label in x_rows else (NUM1 if label in one_dp_rows else num_fmt))
            fcell_fill = fillG if (shade_fcst and y in FCST) else (fillL if is_bold else None)
            c=put(ws,r,2+j,(round(v,4) if isinstance(v,(int,float)) and not (v!=v) else v),
                  boldf if is_bold else None, fcell_fill, fmt, rgt, True)
        r+=1
    return r

# column widths helper
def widths(ws,first=34,rest=12):
    ws.column_dimensions["A"].width=first
    for c in range(2,14): ws.column_dimensions[get_column_letter(c)].width=rest

# =====================================================================================
# 1. COVER
# =====================================================================================
ws=wb.active; ws.title="Cover"; ws.sheet_view.showGridLines=False; widths(ws,40,14)
put(ws,1,1,"UNITED SPIRITS LIMITED (NSE: UNITDSPR | BSE: 532432)",title,fillN)
for c in range(2,7): put(ws,1,c,None,fill=fillN)
put(ws,2,1,"Integrated Financial Model & Initiation of Coverage  |  FY2022–FY2030  |  Consolidated",Font(bold=True,color="FFFFFF"),fillB)
for c in range(2,7): put(ws,2,c,None,fill=fillB)
rows=[
 ("Rating","NEUTRAL / HOLD"),
 ("Current Market Price (10-Jun-2026)","₹ 1,259"),
 ("12-month Target Price (Base)","₹ 1,160"),
 ("Implied upside/(downside)","(8)%"),
 ("Bull / Bear targets","₹ 1,500  /  ₹ 770"),
 ("Market capitalisation","₹ 91,639 cr"),
 ("Shares outstanding","72.78 cr  (FV ₹2)"),
 ("Parent","Diageo plc (~56%)"),
 ("FY26 Net Sales Value (NSV)","₹ 12,467 cr"),
 ("FY26 Core EBITDA / margin","₹ 2,286 cr  /  18.3%"),
 ("FY26 PAT (reported)","₹ 1,838 cr"),
 ("FY26 EPS / RoCE","₹ 25.9  /  ~26%"),
 ("Valuation methodology","60% relative multiples + 40% 2-stage DCF"),
 ("Reporting basis","Consolidated; topline = Net Sales Value (ex-excise)"),
 ("Note","FY26 = last actual; RCB/'Sports' = discontinued/held-for-sale; core forecast = Beverage Alcohol"),
]
r=4
for k,v in rows:
    put(ws,r,1,k,boldf,fillL,align=lft,bd=True); put(ws,r,2,v,None,None,align=lft,bd=True)
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6); r+=1
put(ws,r+1,1,"Currency: INR crore unless stated. Source: USL Annual Reports FY23 & FY25, FY26 results (consolidated). Market data ~10-Jun-2026.",ital)
put(ws,r+2,1,"Disclaimer: Educational model for analytical purposes; not investment advice. Assumptions are the analyst's estimates.",ital)
ws.merge_cells(start_row=r+1,start_column=1,end_row=r+1,end_column=6)
ws.merge_cells(start_row=r+2,start_column=1,end_row=r+2,end_column=6)

# =====================================================================================
# 2. ASSUMPTIONS
# =====================================================================================
ws=sheet("Assumptions"); widths(ws,38,12)
arows=[
 ("Volume growth (mn cases) %",{**{y:None for y in HIST},**A["vol_growth"]}),
 ("Realisation (NSV/case) growth %",{**{y:None for y in HIST},**A["real_growth"]}),
 ("Core (Beverage) EBITDA margin %",{**{y:None for y in HIST},**A["bev_ebitda_margin"]}),
 ("Gross margin (% of NSV)",{**{y:None for y in HIST},**A["gross_margin_nsv"]}),
 ("A&P (% of NSV)",{**{y:None for y in HIST},**A["ad_promo_pct_nsv"]}),
 ("Employee cost (% of NSV)",{**{y:None for y in HIST},**A["employee_pct_nsv"]}),
 ("Excise duty (% of gross revenue)",{**{y:None for y in HIST},**A["excise_pct_gross"]}),
 ("Effective tax rate %",{**{y:None for y in HIST},**A["tax_rate"]}),
 ("Dividend payout %",{**{y:None for y in HIST},**A["payout"]}),
 ("Treasury yield %",{**{y:None for y in HIST},**A["treasury_yield"]}),
 ("Capex (% of NSV)",{**{y:None for y in HIST},**A["capex_pct_nsv"]}),
 ("New lease additions (% of NSV)",{**{y:None for y in HIST},**A["lease_add_pct_nsv"]}),
 ("Receivable days (on gross rev)",{**{y:None for y in HIST},**A["recv_days_gross"]}),
 ("Inventory days (on COGS)",{**{y:None for y in HIST},**A["inv_days_cogs"]}),
 ("Payable days (on COGS)",{**{y:None for y in HIST},**A["pay_days_cogs"]}),
]
pcts={"Volume growth (mn cases) %","Realisation (NSV/case) growth %","Core (Beverage) EBITDA margin %",
 "Gross margin (% of NSV)","A&P (% of NSV)","Employee cost (% of NSV)","Excise duty (% of gross revenue)",
 "Effective tax rate %","Dividend payout %","Treasury yield %","Capex (% of NSV)","New lease additions (% of NSV)"}
onedp={"Receivable days (on gross rev)","Inventory days (on COGS)","Payable days (on COGS)"}
table(ws,1,arows,title_text="FORECAST ASSUMPTIONS — BASE CASE (FY27E–FY30E)",pct_rows=pcts,one_dp_rows=onedp,
      bold_rows={"Core (Beverage) EBITDA margin %"})

# =====================================================================================
# 3. INCOME STATEMENT
# =====================================================================================
ws=sheet("Income Statement"); widths(ws,38,12)
ebit={y:(M["bev_ebitda"][y]-M["dep_amort"][y]) for y in YEARS}
isrows=[
 ("Volume (mn cases)",M["volume_mn_cases"]),
 ("NSV per case (₹)",M["nsv_per_case"]),
 ("Gross revenue (incl. excise)",M["gross_revenue"]),
 ("Less: Excise duty",M["excise_duty"]),
 ("Net Sales Value (NSV)",M["net_revenue"]),
 ("Cost of goods sold",M["cogs"]),
 ("Gross Profit",M["bev_gp"]),
 ("Employee cost",M["employee"]),
 ("Advertising & promotion",M["ad_promo"]),
 ("Other operating expense",M["other_expense"]),
 ("EBITDA (core, ex-other income)",M["bev_ebitda"]),
 ("Depreciation & amortisation",M["dep_amort"]),
 ("EBIT",ebit),
 ("Other income (treasury)",M["other_income"]),
 ("Finance costs",M["finance_cost"]),
 ("Share of JV profit/(loss)",M["jv_share"]),
 ("Exceptional items",M["exceptional"]),
 ("Profit before tax",M["pbt"]),
 ("Tax expense",M["tax"]),
 ("PAT (continuing operations)",M["pat_continuing"]),
 ("Discontinued ops (RCB)",M["disc_ops_pat"]),
 ("PAT (reported)",M["pat_reported"]),
 ("EPS (₹)",M["eps_reported"]),
]
table(ws,1,isrows,title_text="CONSOLIDATED INCOME STATEMENT (NSV basis; continuing operations)",
      one_dp_rows={"Volume (mn cases)","EPS (₹)"},num_fmt=NUM,
      bold_rows={"Net Sales Value (NSV)","Gross Profit","EBITDA (core, ex-other income)","EBIT","Profit before tax","PAT (reported)"})
r=len(isrows)+4
# margins block
marg=[
 ("NSV growth %",{**{"FY22":None},**R["nsv_growth"]}),
 ("Gross margin %",R["gross_margin"]),
 ("EBITDA margin %",R["bev_ebitda_margin"]),
 ("EBIT margin %",R["ebit_margin"]),
 ("PAT margin %",R["pat_margin"]),
]
table(ws,r,marg,title_text="MARGIN & GROWTH ANALYSIS",pct_rows={x[0] for x in marg})

# =====================================================================================
# 4. BALANCE SHEET
# =====================================================================================
ws=sheet("Balance Sheet"); widths(ws,40,12)
treas_cash={y:M["cash"][y] for y in YEARS}
bsrows=[
 ("ASSETS",{}),
 ("Net property, plant & equipment (incl. ROU)",M["ppe"]),
 ("Capital work-in-progress",M["cwip"]),
 ("Goodwill",M["goodwill"]),
 ("Other intangible assets",M["intangibles"]),
 ("Investment property",M["inv_property"]),
 ("Investment in JVs",M["jv_invest"]),
 ("Inventories",M["inventory"]),
 ("Trade receivables (current)",M["receivables_cur"]),
 ("Trade receivables (non-current)",M["receivables_nc"]),
 ("Cash & cash equivalents",M["cash"]),
 ("Bank deposits",M["bank_deposits"]),
 ("Current investments (treasury)",M["cur_investments"]),
 ("Other financial assets",M["other_fin_assets"]),
 ("Deferred tax assets (net)",M["dta_net"]),
 ("Current tax assets",M["cur_tax_assets"]),
 ("Other assets (incl. tax under protest)",M["other_assets"]),
 ("Assets held for sale (RCB)",M["assets_hfs"]),
 ("TOTAL ASSETS",M["total_assets"]),
 ("EQUITY & LIABILITIES",{}),
 ("Total equity",M["equity"]),
 ("Borrowings",M["borrowings"]),
 ("Lease liabilities",M["lease_liab"]),
 ("Trade payables",M["trade_payables"]),
 ("Deferred tax liabilities",M["dtl"]),
 ("Provisions",M["provisions"]),
 ("Current tax liabilities",M["cur_tax_liab"]),
 ("Other financial liabilities",M["other_fin_liab"]),
 ("Other current liabilities",M["other_cur_liab"]),
 ("Liabilities of disposal group (RCB)",M["liab_hfs"]),
 ("TOTAL EQUITY & LIABILITIES",M["total_liab_eq"]),
]
table(ws,1,bsrows,title_text="CONSOLIDATED BALANCE SHEET",
      bold_rows={"ASSETS","EQUITY & LIABILITIES","TOTAL ASSETS","TOTAL EQUITY & LIABILITIES","Total equity"})
r=len(bsrows)+4
chk={y:(M["total_assets"][y]-M["total_liab_eq"][y]) for y in YEARS}
table(ws,r,[("Balance check (Assets − E&L)",chk)],title_text="INTEGRITY CHECK (should be 0)",num_fmt=NUM1)

# =====================================================================================
# 5. CASH FLOW (forecast)
# =====================================================================================
ws=sheet("Cash Flow"); widths(ws,38,12)
cfo={y:CF[y]["cfo"] for y in FCST}; cfi={y:CF[y]["cfi"] for y in FCST}
cff={y:CF[y]["cff"] for y in FCST}; net={y:CF[y]["net"] for y in FCST}
capex={y:-CF[y]["capex"] for y in FCST}
copen={y:CF[y]["cash_open"] for y in FCST}; cclose={y:CF[y]["cash_close"] for y in FCST}
# include historical reported CF for context
cforows=[
 ("Net cash from operating activities",{**{y:M['cfo'][y] for y in HIST},**cfo}),
 ("  of which: Capex (PPE+intangibles)",{**{y:-M['capex'][y] for y in HIST},**capex}),
 ("Net cash from investing activities",{**{y:M['cfi'][y] for y in HIST},**cfi}),
 ("Net cash from financing activities",{**{y:M['cff'][y] for y in HIST},**cff}),
 ("  of which: Dividends paid",{**{y:-M['dividend_paid'][y] for y in HIST},**{y:-M['dividend_paid'][y] for y in FCST}}),
 ("Net change in cash & equivalents",{**{y:None for y in HIST},**net}),
 ("Closing cash & equivalents",{**{y:M['cash'][y] for y in HIST},**{y:M['cash'][y] for y in FCST}}),
 ("Memo: Treasury (cash+deposits+inv)",{y:M['cash'][y]+M['bank_deposits'][y]+M['cur_investments'][y] for y in YEARS}),
]
table(ws,1,cforows,title_text="CASH FLOW STATEMENT (historical reported + forecast derived)",
      bold_rows={"Net cash from operating activities","Net change in cash & equivalents","Closing cash & equivalents"})

# =====================================================================================
# 6. REVENUE BUILD-UP
# =====================================================================================
ws=sheet("Revenue Buildup"); widths(ws,38,12)
revrows=[
 ("Volume (mn cases)",M["volume_mn_cases"]),
 ("  Volume growth %",{**{"FY22":None},**{y:(M["volume_mn_cases"][y]/M["volume_mn_cases"][YEARS[YEARS.index(y)-1]]-1) for y in YEARS[1:]}}),
 ("NSV per case (₹) — realisation",M["nsv_per_case"]),
 ("  Realisation growth %",{**{"FY22":None},**{y:(M["nsv_per_case"][y]/M["nsv_per_case"][YEARS[YEARS.index(y)-1]]-1) for y in YEARS[1:]}}),
 ("Net Sales Value (₹ cr) = Vol × Realisation",M["net_revenue"]),
 ("  NSV growth %",{**{"FY22":None},**R["nsv_growth"]}),
 ("P&A salience — % of NSV",M["pa_nsv_salience"]),
 ("P&A salience — % of volume",M["pa_vol_salience"]),
 ("Excise duty",M["excise_duty"]),
 ("Gross revenue (incl. excise)",M["gross_revenue"]),
]
table(ws,1,revrows,title_text="REVENUE BUILD-UP (Volume × Realisation; premiumisation)",
      one_dp_rows={"Volume (mn cases)"},
      pct_rows={"  Volume growth %","  Realisation growth %","  NSV growth %"},
      bold_rows={"Net Sales Value (₹ cr) = Vol × Realisation"})
# note: P&A salience expressed as number not %
r2=len(revrows)+4
put(ws,r2,1,"P&A = Prestige & Above. Salience shown as % (e.g. 90 = 90%). NSV growth ≈ (1+vol)(1+realisation)−1.",ital)

# =====================================================================================
# 7. WORKING CAPITAL
# =====================================================================================
ws=sheet("Working Capital"); widths(ws,38,12)
wcrows=[
 ("Trade receivables (current)",M["receivables_cur"]),
 ("Inventories",M["inventory"]),
 ("Trade payables",M["trade_payables"]),
 ("Net trade working capital",R["nwc"]),
 ("Receivable days (on gross rev)",R["recv_days"]),
 ("Receivable days (on NSV, memo)",R["recv_days_net"]),
 ("Inventory days (on COGS)",R["inv_days"]),
 ("Payable days (on COGS)",R["pay_days"]),
 ("Cash conversion cycle (days)",R["ccc"]),
]
table(ws,1,wcrows,title_text="WORKING CAPITAL SCHEDULE",
      one_dp_rows={"Receivable days (on gross rev)","Receivable days (on NSV, memo)","Inventory days (on COGS)","Payable days (on COGS)","Cash conversion cycle (days)"},
      bold_rows={"Net trade working capital","Cash conversion cycle (days)"})
r=len(wcrows)+4
put(ws,r,1,"Receivable days computed on GROSS revenue (receivables embed excise billed to state corporations). Inventory/Payable days on COGS. CCC = Inv + Recv − Pay.",ital)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10)

# =====================================================================================
# 8. PPE SCHEDULE
# =====================================================================================
ws=sheet("PPE Schedule"); widths(ws,38,12)
# historical net PPE roll (reported) + forecast owned PPE + ROU
ppe_open={"FY24":M["ppe"]["FY23"],"FY25":M["ppe"]["FY24"],"FY26":M["ppe"]["FY25"]+M["rou"]["FY25"]}
ppe_rows=[
 ("Owned PPE — opening (net)",{**{y:None for y in HIST},**{y:M["ppe_owned"][YEARS[YEARS.index(y)-1]] for y in FCST}}),
 ("  + Capex",{**{y:M["capex"][y] for y in HIST},**M["capex_fc"]}),
 ("  − Depreciation (PPE)",{**{y:None for y in HIST},**{y:-M["ppe_dep"][y] for y in FCST}}),
 ("Owned PPE — closing (net)",{**{y:None for y in HIST},**{y:M["ppe_owned"][y] for y in FCST}}),
 ("ROU asset — opening",{**{y:M["rou"][YEARS[YEARS.index(y)-1]] for y in ["FY24","FY25"]},**{y:M["rou_fc"][YEARS[YEARS.index(y)-1]] for y in FCST}}),
 ("  + New leases (ROU additions)",{"FY25":394,"FY24":184,**M["rou_add"]}),
 ("  − ROU depreciation",{"FY25":-147,"FY24":-129,**{y:-M["rou_dep"][y] for y in FCST}}),
 ("ROU asset — closing",{**{y:M["rou"][y] for y in HIST},**{y:M["rou_fc"][y] for y in FCST}}),
 ("Reported Net PPE (incl. ROU)",M["ppe"]),
 ("Capital work-in-progress",M["cwip"]),
 ("Capex / NSV %",R["capex_sales"]),
 ("Depreciation / NSV %",R["dep_rate"]),
 ("Asset turnover (NSV / Net PPE)",R["asset_turn"]),
]
table(ws,1,ppe_rows,title_text="FIXED-ASSET (PPE) & ROU SCHEDULE",
      pct_rows={"Capex / NSV %","Depreciation / NSV %"},x_rows={"Asset turnover (NSV / Net PPE)"},
      bold_rows={"Owned PPE — closing (net)","ROU asset — closing","Reported Net PPE (incl. ROU)"})

# =====================================================================================
# 9. LEASE SCHEDULE
# =====================================================================================
ws=sheet("Lease Schedule"); widths(ws,38,12)
lease_rows=[
 ("Lease liability — opening",{**{y:None for y in HIST},**{y:M["lease_liab"][YEARS[YEARS.index(y)-1]] for y in FCST}}),
 ("  + New leases",{**{y:None for y in HIST},**M["rou_add"]}),
 ("  + Interest accretion",{"FY24":21,"FY25":40,**M["lease_int"]}),
 ("  − Lease payments",{"FY24":-126,"FY25":-137,**{y:-M["lease_pay"][y] for y in FCST}}),
 ("Lease liability — closing",M["lease_liab"]),
 ("ROU asset — closing",{**{y:M["rou"][y] for y in HIST},**{y:M["rou_fc"][y] for y in FCST}}),
 ("Lease interest (P&L finance cost)",{"FY24":21,"FY25":40,**M["lease_int"]}),
]
table(ws,1,lease_rows,title_text="LEASE SCHEDULE (Ind AS 116)",
      bold_rows={"Lease liability — closing"})

# =====================================================================================
# 10. DEBT & CASH SCHEDULE
# =====================================================================================
ws=sheet("Debt & Cash"); widths(ws,38,12)
debt_rows=[
 ("Borrowings (gross debt)",M["borrowings"]),
 ("Lease liabilities",M["lease_liab"]),
 ("Total debt (incl. leases)",{y:M["borrowings"][y]+M["lease_liab"][y] for y in YEARS}),
 ("Cash & equivalents",M["cash"]),
 ("Bank deposits",M["bank_deposits"]),
 ("Current investments",M["cur_investments"]),
 ("Total cash & treasury",{y:M["cash"][y]+M["bank_deposits"][y]+M["cur_investments"][y] for y in YEARS}),
 ("Net debt / (net cash)",R["net_debt"]),
 ("Interest income (other income)",M["other_income"]),
 ("Treasury yield % (income/avg cash)",{**{y:None for y in HIST},**A["treasury_yield"]}),
]
table(ws,1,debt_rows,title_text="DEBT, CASH & TREASURY SCHEDULE",
      pct_rows={"Treasury yield % (income/avg cash)"},
      bold_rows={"Total debt (incl. leases)","Total cash & treasury","Net debt / (net cash)"})
r=len(debt_rows)+4
put(ws,r,1,"USL is net cash: deleveraged from ~₹3,400 cr gross debt (FY22) to near-zero. Diageo parentage + strong FCF obviate the need for leverage.",ital)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10)

# =====================================================================================
# 11. DCF
# =====================================================================================
ws=sheet("DCF"); widths(ws,38,14)
fc=st["FC"]
dcf_rows=[
 ("EBIT",{y:fc[y]["ebit"] for y in FCST}),
 ("Less: Tax on EBIT (25.5%)",{y:-fc[y]["tax"] for y in FCST}),
 ("NOPAT",{y:fc[y]["ebit"]-fc[y]["tax"] for y in FCST}),
 ("Add: Depreciation & amortisation",{y:fc[y]["da"] for y in FCST}),
 ("Less: Capex",{y:-fc[y]["capex"] for y in FCST}),
 ("Less: New lease additions",{y:-fc[y]["rou_add"] for y in FCST}),
 ("Less: Change in working capital",{y:-fc[y]["dwc"] for y in FCST}),
 ("Free Cash Flow to Firm (FCFF)",{y:fc[y]["fcff"] for y in FCST}),
]
table(ws,1,dcf_rows,years=FCST,title_text="DCF — EXPLICIT FCFF (FY27E–FY30E)",
      bold_rows={"NOPAT","Free Cash Flow to Firm (FCFF)"})
r=len(dcf_rows)+4
out=[
 ("WACC","10.25%"),("Terminal growth (g)","6.25%"),
 ("PV of explicit + fade FCFF (₹ cr)",f"{base_dcf['pv_explicit']:,.0f}"),
 ("PV of terminal value (₹ cr)",f"{base_dcf['pv_tv']:,.0f}"),
 ("Enterprise Value (₹ cr)",f"{base_dcf['ev']:,.0f}"),
 ("Add: Net cash (₹ cr)",f"{base_dcf['net_cash']:,.0f}"),
 ("Add: RCB / held-for-sale optionality (₹ cr)",f"{base_dcf['rcb']:,.0f}"),
 ("Equity Value (₹ cr)",f"{base_dcf['equity_val']:,.0f}"),
 ("DCF value per share (₹)",f"{base_dcf['tp']:,.0f}"),
 ("Relative-valuation value per share (₹)",f"{base_rel['blend']:,.0f}"),
 ("Blended 12-m Target (60% rel / 40% DCF) (₹)","1,160"),
]
put(ws,r,1,"DCF OUTPUT & BRIDGE",title,fillB); 
for c in range(2,4): put(ws,r,c,None,fill=fillB)
r+=1
for k,v in out:
    put(ws,r,1,k,boldf,fillL,align=lft,bd=True); put(ws,r,2,v,None,None,align=rgt,bd=True); r+=1
put(ws,r+1,1,"2-stage DCF: explicit FY27–30 + premiumisation fade FY31–36 + Gordon terminal. RCB held conservatively at ₹2,500 cr (media valuations higher).",ital)
ws.merge_cells(start_row=r+1,start_column=1,end_row=r+1,end_column=6)

# Sensitivity
r=r+3
put(ws,r,1,"SENSITIVITY — Target price (₹): WACC (rows) × terminal growth (cols)",title,fillB)
for c in range(2,8): put(ws,r,c,None,fill=fillB)
r+=1
put(ws,r,1,"WACC \\ g",hdr,fillN,align=ctr,bd=True)
for j,col in enumerate(sens.columns): put(ws,r,2+j,col,hdr,fillN,align=ctr,bd=True)
r+=1
for i,idx in enumerate(sens.index):
    put(ws,r,1,idx,boldf,fillL,align=ctr,bd=True)
    for j,col in enumerate(sens.columns):
        v=sens.loc[idx,col]
        fill=fillGr if abs(v-1160)<120 else None
        put(ws,r,2+j,v,None,fill,RS,rgt,True)
    r+=1

# =====================================================================================
# 12. RELATIVE VALUATION (COMPS)
# =====================================================================================
ws=sheet("Comps"); widths(ws,26,13)
put(ws,1,1,"TRADING COMPARABLES (market data ~10-Jun-2026; ₹ cr)",title,fillB)
for c in range(2,11): put(ws,1,c,None,fill=fillB)
cols=["Company","MktCap","EV","NetRev FY26","EBITDA FY26","PAT FY26","EV/EBITDA","EV/Sales","P/E","EBITDA margin"]
for j,c in enumerate(cols): put(ws,2,1+j,c,hdr,fillN,align=ctr,bd=True)
rr=3
for _,row in comps.iterrows():
    vals=[row["Company"],row["MktCap"],row["EV"],row["NetRev"],row["EBITDA"],row["PAT"],
          row["EV/EBITDA"],row["EV/Sales"],row["P/E"],row["EBITDA_margin"]]
    for j,v in enumerate(vals):
        fmt=None
        if j in (1,2,3,4,5): fmt=NUM
        if j in (6,7,8): fmt=X
        if j==9: fmt=PC
        fill=fillGo if row["Company"].startswith("United") else None
        al=lft if j==0 else rgt
        put(ws,rr,1+j,(round(v,2) if isinstance(v,(int,float)) else v),boldf if j==0 else None,fill,fmt,al,True)
    rr+=1
# peer median
med=comps.iloc[1:][["EV/EBITDA","EV/Sales","P/E","EBITDA_margin"]].median()
put(ws,rr,1,"Peer median (ex-USL)",boldf,fillL,align=lft,bd=True)
for j in range(1,6): put(ws,rr,1+j,None,fill=fillL,bd=True)
put(ws,rr,7,round(med["EV/EBITDA"],2),boldf,fillL,X,rgt,True)
put(ws,rr,8,round(med["EV/Sales"],2),boldf,fillL,X,rgt,True)
put(ws,rr,9,round(med["P/E"],2),boldf,fillL,X,rgt,True)
put(ws,rr,10,round(med["EBITDA_margin"],3),boldf,fillL,PC,rgt,True)
rr+=2
put(ws,rr,1,"Tilaknagar FY26 PAT normalised (reported ₹21 cr distorted by Imperial Blue acquisition one-offs). ABDL mkt cap/EV approximate. USL EV reflects net-cash balance sheet.",ital)
ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=10)

# =====================================================================================
# 13. HISTORICAL VALUATION
# =====================================================================================
ws=sheet("Historical Valuation"); widths(ws,30,13)
hv_rows=[
 ("Market capitalisation (₹ cr)",{y:hist_val[y]["mcap"] for y in HIST}),
 ("PAT (reported, ₹ cr)",{y:M["pat_reported"][y] for y in HIST}),
 ("P/E (x)",{y:hist_val[y]["pe"] for y in HIST}),
 ("Enterprise value (₹ cr)",{y:hist_val[y]["ev"] for y in HIST}),
 ("Core EBITDA (₹ cr)",{y:M["bev_ebitda"][y] for y in HIST}),
 ("EV/EBITDA (x)",{y:hist_val[y]["ev_ebitda"] for y in HIST}),
]
table(ws,1,hv_rows,years=HIST,title_text="HISTORICAL VALUATION (FY22–FY26)",
      x_rows={"P/E (x)","EV/EBITDA (x)"},
      bold_rows={"P/E (x)","EV/EBITDA (x)"})
r=len(hv_rows)+4
put(ws,r,1,"5-yr average P/E ~60x; current ~50x (below average). EV/EBITDA de-rated from ~48x (FY25) to ~39x (FY26). Market cap peaked at ~₹1.02 lakh cr (FY25).",ital)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)

# =====================================================================================
# 14. SCENARIOS
# =====================================================================================
ws=sheet("Scenarios"); widths(ws,30,15)
put(ws,1,1,"SCENARIO ANALYSIS (FY30E outcomes & 12-m target price)",title,fillB)
for c in range(2,6): put(ws,1,c,None,fill=fillB)
hdrs=["Metric","Bear","Base","Bull"]
for j,h in enumerate(hdrs): put(ws,2,1+j,h,hdr,fillN,align=ctr,bd=True)
srows=[
 ("FY30E Net Sales (₹ cr)",[S_bear["nsv30"],S_base["nsv30"],S_bull["nsv30"]],NUM),
 ("FY30E EBITDA (₹ cr)",[S_bear["ebitda30"],S_base["ebitda30"],S_bull["ebitda30"]],NUM),
 ("FY30E PAT (₹ cr)",[S_bear["pat30"],S_base["pat30"],S_bull["pat30"]],NUM),
 ("FY30E EPS (₹)",[S_bear["eps30"],S_base["eps30"],S_bull["eps30"]],NUM1),
 ("NSV CAGR FY26-30",[S_bear["nsv_cagr"],S_base["nsv_cagr"],S_bull["nsv_cagr"]],PC),
 ("EPS CAGR FY26-30",[S_bear["eps_cagr"],S_base["eps_cagr"],S_bull["eps_cagr"]],PC),
 ("DCF value (₹)",[S_bear["tp_dcf"],S_base["tp_dcf"],S_bull["tp_dcf"]],RS),
 ("Relative value (₹)",[S_bear["tp_rel"],S_base["tp_rel"],S_bull["tp_rel"]],RS),
 ("Target price (₹)",[S_bear["tp"],S_base["tp"],S_bull["tp"]],RS),
 ("Upside / (downside) vs ₹1,259",[S_bear["upside"],S_base["upside"],S_bull["upside"]],PC),
]
rr=3
for label,vals,fmt in srows:
    bold=label.startswith("Target")
    put(ws,rr,1,label,boldf if bold else None,fillL if bold else None,align=lft,bd=True)
    for j,v in enumerate(vals):
        fill=fillGo if bold else None
        put(ws,rr,2+j,round(v,1) if fmt==NUM1 else (round(v,4) if fmt==PC else round(v)),boldf if bold else None,fill,fmt,rgt,True)
    rr+=1
put(ws,rr+1,1,"Bull: 6%/6% vol/price, EBITDA margin →22%, WACC 9.75%/g 6.75%. Bear: 2.5%/4%, margin ~18.8%, WACC 11.5%/g 5%. Target = 60% relative + 40% DCF.",ital)
ws.merge_cells(start_row=rr+1,start_column=1,end_row=rr+1,end_column=5)

# reorder so Cover first
wb._sheets.sort(key=lambda s: 0 if s.title=="Cover" else 1)
wb.save("USL_Financial_Model.xlsx")
print("Saved USL_Financial_Model.xlsx with sheets:", [s.title for s in wb._sheets])
