"""
USL valuation + output builder.
Runs base/bull/bear scenarios on the engine, builds DCF (2-stage), trading comps,
historical valuation, sensitivity tables, then writes an Excel workbook and a
markdown tables file used by the research report.
"""
import numpy as np, pandas as pd, copy, json
import usl_engine as E
from usl_engine import YEARS, HIST, FCST, SHARES, prev

# =====================================================================================
# RUN SCENARIOS
# =====================================================================================
def bull():
    A = E.base_assumptions()
    A["vol_growth"]={"FY27E":0.060,"FY28E":0.060,"FY29E":0.055,"FY30E":0.055}
    A["real_growth"]={"FY27E":0.062,"FY28E":0.060,"FY29E":0.058,"FY30E":0.055}
    A["bev_ebitda_margin"]={"FY27E":0.193,"FY28E":0.202,"FY29E":0.212,"FY30E":0.220}
    A["gross_margin_nsv"]={"FY27E":0.475,"FY28E":0.482,"FY29E":0.489,"FY30E":0.495}
    A["payout"]={"FY27E":0.55,"FY28E":0.55,"FY29E":0.58,"FY30E":0.58}
    return A
def bear():
    A = E.base_assumptions()
    A["vol_growth"]={"FY27E":0.025,"FY28E":0.025,"FY29E":0.020,"FY30E":0.020}
    A["real_growth"]={"FY27E":0.040,"FY28E":0.040,"FY29E":0.038,"FY30E":0.035}
    A["bev_ebitda_margin"]={"FY27E":0.182,"FY28E":0.184,"FY29E":0.186,"FY30E":0.188}
    A["gross_margin_nsv"]={"FY27E":0.462,"FY28E":0.463,"FY29E":0.464,"FY30E":0.465}
    A["tax_rate"]={"FY27E":0.260,"FY28E":0.260,"FY29E":0.260,"FY30E":0.260}
    return A

A_base = E.base_assumptions()
M = E.run(A_base)
M_bull = E.run(bull())
M_bear = E.run(bear())

# =====================================================================================
# RATIOS (base)
# =====================================================================================
def R_build(M):
    R={}
    R["nsv_growth"]={y:(M["net_revenue"][y]/M["net_revenue"][prev(y)]-1) for y in YEARS[1:]}
    R["bev_ebitda_margin"]={y:M["bev_ebitda"][y]/M["net_revenue"][y] for y in YEARS}
    R["gross_margin"]={y:M["bev_gp"][y]/M["net_revenue"][y] for y in YEARS}
    R["ebit"]={y:(M["bev_ebitda"][y]-M["dep_amort"][y]) for y in YEARS}
    R["ebit_margin"]={y:R["ebit"][y]/M["net_revenue"][y] for y in YEARS}
    R["pat_margin"]={y:M["pat_reported"][y]/M["net_revenue"][y] for y in YEARS}
    R["adp_pct"]={y:M["ad_promo"][y]/M["net_revenue"][y] for y in YEARS}
    R["emp_pct"]={y:M["employee"][y]/M["net_revenue"][y] for y in YEARS}
    R["cogs_pct"]={y:M["cogs"][y]/M["net_revenue"][y] for y in YEARS}
    R["othexp_pct"]={y:M["other_expense"][y]/M["net_revenue"][y] for y in YEARS}
    R["recv_days"]={y:M["receivables_cur"][y]/M["gross_revenue"][y]*365 for y in YEARS}
    R["recv_days_net"]={y:M["receivables_cur"][y]/M["net_revenue"][y]*365 for y in YEARS}
    R["inv_days"]={y:M["inventory"][y]/M["cogs"][y]*365 for y in YEARS}
    R["pay_days"]={y:M["trade_payables"][y]/M["cogs"][y]*365 for y in YEARS}
    R["ccc"]={y:R["inv_days"][y]+R["recv_days"][y]-R["pay_days"][y] for y in YEARS}
    R["nwc"]={y:(M["inventory"][y]+M["receivables_cur"][y]-M["trade_payables"][y]) for y in YEARS}
    R["net_debt"]={y:(M["borrowings"][y]+M["lease_liab"][y]-M["cash"][y]-M["bank_deposits"][y]-M["cur_investments"][y]) for y in YEARS}
    R["roce"]={y:R["ebit"][y]*(1-0.255)/((M["equity"][y]+M["borrowings"][y]+M["lease_liab"][y])) for y in YEARS}
    R["roe"]={y:(M["pat_reported"][y]/(0.5*(M["equity"][y]+M["equity"][prev(y)])) if y!="FY22" else np.nan) for y in YEARS}
    R["asset_turn"]={y:M["net_revenue"][y]/M["ppe"][y] for y in YEARS}
    R["capex"]={y:(M["capex"][y] if y in HIST else M["capex_fc"][y]) for y in YEARS}
    R["capex_sales"]={y:R["capex"][y]/M["net_revenue"][y] for y in YEARS}
    R["dep_rate"]={y:M["dep_amort"][y]/M["net_revenue"][y] for y in YEARS}
    return R
R = R_build(M)

# =====================================================================================
# CASH FLOW (forecast, derived) & FCFF
# =====================================================================================
def cashflow(M):
    """Cash-flow statement that ties EXACTLY to the change in treasury (the BS plug).
       Net = CFO+CFI+CFF = treasury_t - treasury_{t-1} by construction."""
    CF={}
    for y in FCST:
        p=prev(y)
        pat=M["pat_reported"][y]; da=M["dep_amort"][y]; lint=M["lease_int"][y]
        # working-capital cash effect: +Δ(op liabilities) − Δ(op assets)
        d_assets = ((M["inventory"][y]-M["inventory"][p])
                    +(M["receivables_cur"][y]-M["receivables_cur"][p])
                    +(M["receivables_nc"][y]-M["receivables_nc"][p])
                    +(M["other_assets"][y]-M["other_assets"][p]))
        d_liab   = ((M["trade_payables"][y]-M["trade_payables"][p])
                    +(M["provisions"][y]-M["provisions"][p])
                    +(M["cur_tax_liab"][y]-M["cur_tax_liab"][p])
                    +(M["other_cur_liab"][y]-M["other_cur_liab"][p]))
        dwc = d_liab - d_assets
        cfo = pat + da + dwc + lint                      # lease interest added back (financing)
        capex = M["capex_fc"][y]
        # RCB / held-for-sale disposal (FY27): assets derecognised net of disposal-group liabilities
        d_hfs = (M["assets_hfs"][p]-M["assets_hfs"][y]) - (M["liab_hfs"][p]-M["liab_hfs"][y])
        cfi = -capex + d_hfs                             # ROU additions are non-cash
        d_borrow = M["borrowings"][y]-M["borrowings"][p]
        cff = -M["dividend_paid"][y] - M["lease_pay"][y] + d_borrow
        net = cfo+cfi+cff
        topen = M["treasury"][p]; tclose = topen+net
        CF[y]=dict(cfo=cfo,da=da,dwc=dwc,lint=lint,capex=capex,cfi=cfi,cff=cff,net=net,d_hfs=d_hfs,
                   div=M["dividend_paid"][y],lease_pay=M["lease_pay"][y],d_borrow=d_borrow,
                   t_open=topen,t_close=tclose,t_bs=M["treasury"][y],
                   tie=tclose-M["treasury"][y])
    return CF
CF = cashflow(M)
print("\n=== Cash-flow tie check (CF closing treasury − BS treasury, should be ~0) ===")
for y in FCST: print(f"  {y}: {CF[y]['tie']:+.3f}")

# FCFF for each forecast year (unlevered)
def fcff_series(M):
    out={}
    for y in FCST:
        p=prev(y)
        ebit=M["bev_ebitda"][y]-M["dep_amort"][y]
        tax=ebit*0.255
        nwc_y=M["inventory"][y]+M["receivables_cur"][y]+M["receivables_nc"][y]-M["trade_payables"][y]-M["other_cur_liab"][y]
        nwc_p=M["inventory"][p]+M["receivables_cur"][p]+M["receivables_nc"][p]-M["trade_payables"][p]-M["other_cur_liab"][p]
        dwc=nwc_y-nwc_p
        capex=M["capex_fc"][y]; rou_add=M["rou_add"][y]
        fcff=ebit-tax+M["dep_amort"][y]-capex-rou_add-dwc
        out[y]=dict(ebit=ebit,tax=tax,da=M["dep_amort"][y],capex=capex,rou_add=rou_add,dwc=dwc,fcff=fcff)
    return out
FC = fcff_series(M)

# =====================================================================================
# DCF — 2-stage: explicit FY27-FY30 + premiumisation fade FY31-FY36 + terminal
# =====================================================================================
def dcf(wacc=0.105, g_term=0.060, base_M=M, verbose=False):
    # stage-1 explicit FCFF (FY27-FY30)
    fc = fcff_series(base_M)
    flows=[fc[y]["fcff"] for y in FCST]
    # stage-2 fade FY31-FY36: NSV growth fades 9%->6%, EBITDA margin -> 22.5%, FCFF conversion improves
    nsv = base_M["net_revenue"]["FY30E"]
    ebitda_m = base_M["bev_ebitda"]["FY30E"]/nsv
    da_pct = base_M["dep_amort"]["FY30E"]/nsv
    # fade path
    growths=[0.085,0.080,0.075,0.070,0.065,0.060]
    margins=[ebitda_m+ (0.225-ebitda_m)*k/6 for k in range(1,7)]
    fade=[]
    for i,gr in enumerate(growths):
        nsv=nsv*(1+gr)
        ebitda=nsv*margins[i]
        da=nsv*da_pct
        ebit=ebitda-da
        nopat=ebit*(1-0.255)
        capex=nsv*0.018; rou=nsv*0.016
        dwc=nsv*gr*0.16   # incremental NWC ~16% of incremental sales
        fcff=nopat+da-capex-rou-dwc
        fade.append(fcff)
    allflows=flows+fade
    n=len(allflows)
    pv=sum(allflows[i]/(1+wacc)**(i+1) for i in range(n))
    tv=allflows[-1]*(1+g_term)/(wacc-g_term)
    pv_tv=tv/(1+wacc)**n
    ev=pv+pv_tv
    net_cash=-(base_M["borrowings"]["FY26"]+base_M["lease_liab"]["FY26"]
               -base_M["cash"]["FY26"]-base_M["bank_deposits"]["FY26"]-base_M["cur_investments"]["FY26"])
    rcb_value=2500.0   # optional value of RCB (held-for-sale) net of tax, conservative
    equity_val=ev+net_cash+rcb_value
    tp=equity_val/SHARES
    if verbose:
        print(f"  WACC {wacc:.1%} g {g_term:.1%}: EV {ev:,.0f} +netcash {net_cash:,.0f} +RCB {rcb_value:,.0f} = EqV {equity_val:,.0f}  TP Rs {tp:,.0f}")
    return dict(ev=ev,pv_explicit=pv,pv_tv=pv_tv,net_cash=net_cash,rcb=rcb_value,equity_val=equity_val,tp=tp,
                flows=allflows,tv=tv)

base_dcf = dcf(0.1025,0.0625,M,verbose=True)

# Sensitivity table (WACC x terminal g)
waccs=[0.095,0.100,0.105,0.110,0.115]
gs=[0.050,0.055,0.060,0.065,0.070]
sens=pd.DataFrame(index=[f"{w:.1%}" for w in waccs], columns=[f"{x:.1%}" for x in gs])
for w in waccs:
    for x in gs:
        sens.loc[f"{w:.1%}",f"{x:.1%}"]=round(dcf(w,x,M)["tp"])

# =====================================================================================
# RELATIVE VALUATION TARGET (target multiple on FY28E)
# =====================================================================================
def rel_val(M, tgt_ev_ebitda, tgt_pe, fy="FY28E"):
    ebitda=M["bev_ebitda"][fy]; eps=M["eps_reported"][fy]; pat=M["pat_reported"][fy]
    net_cash=-(M["borrowings"]["FY26"]+M["lease_liab"]["FY26"]-M["cash"]["FY26"]-M["bank_deposits"]["FY26"]-M["cur_investments"]["FY26"])
    tp_ev=(ebitda*tgt_ev_ebitda+net_cash+2500)/SHARES
    tp_pe=eps*tgt_pe
    return dict(tp_ev=tp_ev,tp_pe=tp_pe,blend=0.5*(tp_ev+tp_pe),ebitda=ebitda,eps=eps)
base_rel = rel_val(M, 34.0, 50.0)

# =====================================================================================
# SCENARIO SUMMARY (Revenue/EBITDA/PAT/EPS FY30 + target price + CAGR)
# =====================================================================================
CMP=1259.0
def scen_summary(Ms, name, wacc, g, ev_mult, pe_mult, w_rel=0.60):
    d=dcf(wacc,g,Ms)
    rel=rel_val(Ms, ev_mult, pe_mult)
    tp=w_rel*rel["blend"]+(1-w_rel)*d["tp"]
    nsv_cagr=(Ms["net_revenue"]["FY30E"]/Ms["net_revenue"]["FY26"])**0.25-1
    eps_cagr=(Ms["eps_reported"]["FY30E"]/M["eps_reported"]["FY26"])**0.25-1
    return dict(name=name,
                nsv30=Ms["net_revenue"]["FY30E"],ebitda30=Ms["bev_ebitda"]["FY30E"],
                pat30=Ms["pat_reported"]["FY30E"],eps30=Ms["eps_reported"]["FY30E"],
                nsv_cagr=nsv_cagr,eps_cagr=eps_cagr,
                tp_dcf=d["tp"],tp_rel=rel["blend"],tp=tp,upside=tp/CMP-1)
S_base=scen_summary(M,"Base",0.1025,0.0625,34.0,50.0)
S_bull=scen_summary(M_bull,"Bull",0.0975,0.0675,40.0,58.0)
S_bear=scen_summary(M_bear,"Bear",0.1150,0.0500,26.0,38.0)

# =====================================================================================
# HISTORICAL VALUATION
# =====================================================================================
mcap_hist={"FY21":40427,"FY22":64607,"FY23":55010,"FY24":82500,"FY25":101924,"FY26":91639}
pat_hist={"FY21":310,"FY22":810.6,"FY23":1125.8,"FY24":1408,"FY25":1582,"FY26":1838}
ebitda_hist={"FY21":988,"FY22":1510,"FY23":1419,"FY24":1708,"FY25":2058,"FY26":2286}
netcash_hist={"FY22":-(341.7+263.7-54.5-5.8-222.1),"FY23":-(1.1+182.2-115.1-768.2-255.8),
              "FY24":-(25+240-1052-217-599),"FY25":-(0+480-1328-702-873),"FY26":-(6+407-859-1118-1157)}
hist_val={}
for y in ["FY22","FY23","FY24","FY25","FY26"]:
    pe=mcap_hist[y]/pat_hist[y]
    ev=mcap_hist[y]-netcash_hist[y]
    eve=ev/ebitda_hist[y]
    hist_val[y]=dict(mcap=mcap_hist[y],pe=pe,ev=ev,ev_ebitda=eve)

# =====================================================================================
# TRADING COMPS (approx, market data ~10-Jun-2026; INR cr) — sourced; estimates labelled
# =====================================================================================
comps=pd.DataFrame([
 # name, price, mcap, ev, net_rev_FY26, ebitda_FY26, pat_FY26
 ["United Spirits (USL)",1259, 91639, 88918, 12467, 2286, 1716],
 ["Radico Khaitan",      None, 46854, 47264, 6050, 1033, 604],
 ["Allied Blenders (ABDL)",None,11500, 12100, 3949, 568, 266],
 ["Tilaknagar Inds",     None, 10500, 14400, 2346, 445, 145],   # PAT normalised (FY26 reported 21 distorted by Imperial Blue one-offs)
 ["Globus Spirits",      None, 2839, 3198, 3625, 360, 91],
], columns=["Company","Price","MktCap","EV","NetRev","EBITDA","PAT"])
comps["EV/EBITDA"]=comps["EV"]/comps["EBITDA"]
comps["EV/Sales"]=comps["EV"]/comps["NetRev"]
comps["P/E"]=comps["MktCap"]/comps["PAT"]
comps["EBITDA_margin"]=comps["EBITDA"]/comps["NetRev"]

# =====================================================================================
# PRINT KEY OUTPUTS
# =====================================================================================
pd.options.display.float_format=lambda x:f"{x:,.1f}"
print("\n=== DCF base ==="); print({k:round(v,0) for k,v in base_dcf.items() if k in ("ev","pv_explicit","pv_tv","net_cash","equity_val","tp")})
print("\n=== Relative val (FY28E, 34x EV/EBITDA, 50x PE) ==="); print({k:round(v,0) for k,v in base_rel.items()})
print("\n=== Scenarios ===")
for s in (S_bear,S_base,S_bull):
    print(f"  {s['name']:5s} NSV30 {s['nsv30']:,.0f}  EBITDA30 {s['ebitda30']:,.0f}  PAT30 {s['pat30']:,.0f}  EPS30 {s['eps30']:,.1f}  TP {s['tp']:,.0f}  upside {s['upside']:+.0%}")
print("\n=== Sensitivity (TP, WACC rows x g cols) ==="); print(sens.to_string())
print("\n=== Comps ===")
print(comps.to_string(index=False))
print("\n=== Historical valuation ===")
for y,v in hist_val.items(): print(f"  {y}: MCap {v['mcap']:,}  P/E {v['pe']:.1f}x  EV/EBITDA {v['ev_ebitda']:.1f}x")

# persist for excel/report
import pickle
pickle.dump(dict(M=M,M_bull=M_bull,M_bear=M_bear,R=R,CF=CF,FC=FC,A=A_base,
                 base_dcf=base_dcf,base_rel=base_rel,sens=sens,comps=comps,hist_val=hist_val,
                 S_base=S_base,S_bull=S_bull,S_bear=S_bear,CMP=CMP),
            open("build_state.pkl","wb"))
print("\nSaved build_state.pkl")
