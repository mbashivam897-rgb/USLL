"""
USL integrated model ENGINE — reusable build function for scenario analysis.
All figures INR crore unless noted. Consolidated basis. Topline = Net Sales Value (NSV).
FY26 = last actual; FY27E-FY30E forecast. Core = Beverage Alcohol (continuing ops);
RCB/Sports treated as discontinued/held-for-sale (excluded from core forecast).
"""
import numpy as np, pandas as pd, copy

HIST = ["FY22","FY23","FY24","FY25","FY26"]
FCST = ["FY27E","FY28E","FY29E","FY30E"]
YEARS = HIST + FCST
SHARES = 72.78  # crore shares

# ---------------- Historical reported consolidated data ----------------
H = {
 "gross_revenue":{"FY22":31061.8,"FY23":27815.4,"FY24":26018,"FY25":27276,"FY26":27816},
 "excise_duty":{"FY22":21349.4,"FY23":17203.8,"FY24":14697,"FY25":15207,"FY26":15349},
 "net_revenue":{"FY22":9712.4,"FY23":10611.6,"FY24":11321,"FY25":12069,"FY26":12467},
 "other_income":{"FY22":35.5,"FY23":73.1,"FY24":225,"FY25":336,"FY26":478},
 "cogs":{"FY22":5291.5,"FY23":6064.1,"FY24":6048,"FY25":6397,"FY26":6674},
 "employee":{"FY22":653.1,"FY23":610.0,"FY24":547,"FY25":609,"FY26":635},
 "ad_promo":{"FY22":694.9,"FY23":921.8,"FY24":1048,"FY25":1130,"FY26":1295},
 "other_expense":{"FY22":1464.8,"FY23":1598.8,"FY24":1677,"FY25":1690,"FY26":1577},
 "dep_amort":{"FY22":303.8,"FY23":282.5,"FY24":275,"FY25":283,"FY26":289},
 "finance_cost":{"FY22":88.0,"FY23":103.9,"FY24":76,"FY25":89,"FY26":158},
 "jv_share":{"FY22":0.0,"FY23":-1.4,"FY24":-1,"FY25":-7,"FY26":-7},
 "exceptional":{"FY22":-165.2,"FY23":176.4,"FY24":-17,"FY25":-65,"FY26":-91},
 "pbt":{"FY22":1086.6,"FY23":1278.6,"FY24":1857,"FY25":2135,"FY26":2226},
 "tax":{"FY22":276.0,"FY23":152.8,"FY24":449,"FY25":553,"FY26":510},
 "pat_continuing":{"FY22":810.6,"FY23":1125.8,"FY24":1408,"FY25":1582,"FY26":1716},
 "disc_ops_pat":{"FY22":0.0,"FY23":0.0,"FY24":0,"FY25":0,"FY26":129},
 "pat_reported":{"FY22":810.6,"FY23":1125.8,"FY24":1408,"FY25":1582,"FY26":1838},
 "eps_reported":{"FY22":11.68,"FY23":16.01,"FY24":19.83,"FY25":22.28,"FY26":25.89},
 "bev_nsv":{"FY22":9424,"FY23":10374,"FY24":10692,"FY25":11573,"FY26":12467},
 "bev_ebitda":{"FY22":1510,"FY23":1419,"FY24":1708,"FY25":2058,"FY26":2286},
 "bev_gp":{"FY22":4140,"FY23":4302,"FY24":4644,"FY25":5176,"FY26":5793},
 "volume_mn_cases":{"FY22":79.2,"FY23":72.5,"FY24":61.5,"FY25":64.0,"FY26":66.6},
 "nsv_per_case":{"FY22":1459,"FY23":1605,"FY24":1740,"FY25":1810,"FY26":1872},
 "pa_nsv_salience":{"FY22":84,"FY23":87,"FY24":87,"FY25":89,"FY26":90},
 "pa_vol_salience":{"FY22":54,"FY23":66,"FY24":82,"FY25":83,"FY26":84},
 "ppe":{"FY22":1215.2,"FY23":978.3,"FY24":844,"FY25":850,"FY26":1215},
 "rou":{"FY22":260.6,"FY23":172.6,"FY24":227,"FY25":457,"FY26":0},
 "cwip":{"FY22":87.7,"FY23":66.8,"FY24":37,"FY25":72,"FY26":77},
 "goodwill":{"FY22":21.0,"FY23":1.3,"FY24":1,"FY25":1,"FY26":55},
 "intangibles":{"FY22":358.0,"FY23":372.7,"FY24":349,"FY25":329,"FY26":76},
 "inv_property":{"FY22":0.0,"FY23":25.3,"FY24":139,"FY25":75,"FY26":58},
 "jv_invest":{"FY22":0.0,"FY23":10.1,"FY24":46,"FY25":50,"FY26":14},
 "inventory":{"FY22":2156.7,"FY23":2230.0,"FY24":2063,"FY25":2305,"FY26":2668},
 "receivables_cur":{"FY22":2373.6,"FY23":2434.0,"FY24":3056,"FY25":3410,"FY26":3609},
 "receivables_nc":{"FY22":148.5,"FY23":146.2,"FY24":365,"FY25":340,"FY26":351},
 "cash":{"FY22":54.5,"FY23":115.1,"FY24":1052,"FY25":1328,"FY26":859},
 "bank_deposits":{"FY22":5.8,"FY23":768.2,"FY24":217,"FY25":702,"FY26":1118},
 "cur_investments":{"FY22":222.1,"FY23":255.8,"FY24":599,"FY25":873,"FY26":1157},
 "other_fin_assets":{"FY22":225.7,"FY23":279.6,"FY24":148,"FY25":640,"FY26":368},
 "dta_net":{"FY22":147.8,"FY23":157.3,"FY24":177,"FY25":155,"FY26":284},
 "cur_tax_assets":{"FY22":1261.4,"FY23":1335.6,"FY24":1358,"FY25":1115,"FY26":0},
 "other_assets":{"FY22":501.9,"FY23":519.9,"FY24":561,"FY25":524,"FY26":1858},
 "assets_hfs":{"FY22":0.0,"FY23":0.0,"FY24":0,"FY25":0,"FY26":701},
 "total_assets":{"FY22":8911.3,"FY23":9761.5,"FY24":11249,"FY25":13248,"FY26":14469},
 "equity":{"FY22":4874.8,"FY23":5999.5,"FY24":7121,"FY25":8104,"FY26":8957},
 "borrowings":{"FY22":341.7,"FY23":1.1,"FY24":25,"FY25":0,"FY26":6},
 "lease_liab":{"FY22":263.7,"FY23":182.2,"FY24":240,"FY25":480,"FY26":407},
 "trade_payables":{"FY22":1582.1,"FY23":1782.5,"FY24":1954,"FY25":2239,"FY26":2383},
 "dtl":{"FY22":48.3,"FY23":45.1,"FY24":73,"FY25":72,"FY26":18},
 "provisions":{"FY22":502.9,"FY23":385.4,"FY24":383,"FY25":391,"FY26":541},
 "cur_tax_liab":{"FY22":276.5,"FY23":282.9,"FY24":332,"FY25":380,"FY26":571},
 "other_fin_liab":{"FY22":182.3,"FY23":284.6,"FY24":276,"FY25":477,"FY26":273},
 "other_cur_liab":{"FY22":839.0,"FY23":798.2,"FY24":845,"FY25":1105,"FY26":986},
 "liab_hfs":{"FY22":0.0,"FY23":0.0,"FY24":0,"FY25":0,"FY26":327},
 "total_liab_eq":{"FY22":8911.3,"FY23":9761.5,"FY24":11249,"FY25":13248,"FY26":14469},
 "cfo":{"FY22":977.4,"FY23":614.7,"FY24":1118,"FY25":1947,"FY26":1459},
 "capex":{"FY22":134.0,"FY23":136.6,"FY24":98,"FY25":162,"FY26":181},
 "cfi":{"FY22":-312.7,"FY23":-53.2,"FY24":226,"FY25":-1114,"FY26":-428},
 "cff":{"FY22":-688.0,"FY23":-500.9,"FY24":-407,"FY25":-557,"FY26":-1450},
 "dividend_paid":{"FY22":0.0,"FY23":284.0,"FY24":284,"FY25":355,"FY26":1263},
}

def base_assumptions():
    return {
     "vol_growth":{"FY27E":0.045,"FY28E":0.045,"FY29E":0.040,"FY30E":0.040},
     "real_growth":{"FY27E":0.053,"FY28E":0.053,"FY29E":0.053,"FY30E":0.048},
     "bev_ebitda_margin":{"FY27E":0.188,"FY28E":0.194,"FY29E":0.200,"FY30E":0.205},
     "gross_margin_nsv":{"FY27E":0.470,"FY28E":0.474,"FY29E":0.478,"FY30E":0.482},
     "ad_promo_pct_nsv":{"FY27E":0.103,"FY28E":0.103,"FY29E":0.102,"FY30E":0.102},
     "employee_pct_nsv":{"FY27E":0.050,"FY28E":0.049,"FY29E":0.048,"FY30E":0.047},
     "excise_pct_gross":{"FY27E":0.552,"FY28E":0.552,"FY29E":0.551,"FY30E":0.551},
     "tax_rate":{"FY27E":0.255,"FY28E":0.255,"FY29E":0.255,"FY30E":0.255},
     "payout":{"FY27E":0.60,"FY28E":0.60,"FY29E":0.62,"FY30E":0.62},
     "treasury_yield":{"FY27E":0.058,"FY28E":0.059,"FY29E":0.060,"FY30E":0.060},
     "capex_pct_nsv":{"FY27E":0.020,"FY28E":0.020,"FY29E":0.019,"FY30E":0.019},
     "ppe_dep_rate":{"FY27E":0.115,"FY28E":0.115,"FY29E":0.115,"FY30E":0.115},
     "lease_add_pct_nsv":{"FY27E":0.020,"FY28E":0.018,"FY29E":0.018,"FY30E":0.018},
     "rou_dep_rate":{"FY27E":0.32,"FY28E":0.32,"FY29E":0.32,"FY30E":0.32},
     "lease_int_rate":{"FY27E":0.085,"FY28E":0.085,"FY29E":0.085,"FY30E":0.085},
     "lease_pay_pct":{"FY27E":0.34,"FY28E":0.32,"FY29E":0.32,"FY30E":0.32},
     "recv_days_gross":{"FY27E":47.0,"FY28E":46.5,"FY29E":46.0,"FY30E":45.5},
     "inv_days_cogs":{"FY27E":143.0,"FY28E":141.0,"FY29E":139.0,"FY30E":137.0},
     "pay_days_cogs":{"FY27E":130.0,"FY28E":130.0,"FY29E":130.0,"FY30E":130.0},
     "recv_nc_pct_nsv":{"FY27E":0.027,"FY28E":0.027,"FY29E":0.026,"FY30E":0.026},
     "other_cl_pct_nsv":{"FY27E":0.079,"FY28E":0.078,"FY29E":0.078,"FY30E":0.077},
     "provisions_pct_nsv":{"FY27E":0.043,"FY28E":0.043,"FY29E":0.042,"FY30E":0.042},
     "finance_cost_other":{"FY27E":52,"FY28E":55,"FY29E":58,"FY30E":60},
    }

def prev(y): return YEARS[YEARS.index(y)-1]

def run(A):
    M = {k: dict(v) for k,v in H.items()}
    # revenue
    for y in FCST:
        p=prev(y)
        M["volume_mn_cases"][y]=M["volume_mn_cases"][p]*(1+A["vol_growth"][y])
        M["nsv_per_case"][y]=M["nsv_per_case"][p]*(1+A["real_growth"][y])
        M["bev_nsv"][y]=M["volume_mn_cases"][y]*1e6*M["nsv_per_case"][y]/1e7
        M["net_revenue"][y]=M["bev_nsv"][y]
        M["excise_duty"][y]=M["net_revenue"][y]/(1-A["excise_pct_gross"][y])*A["excise_pct_gross"][y]
        M["gross_revenue"][y]=M["net_revenue"][y]+M["excise_duty"][y]
        M["pa_nsv_salience"][y]=min(95,M["pa_nsv_salience"][p]+0.8)
        M["pa_vol_salience"][y]=min(92,M["pa_vol_salience"][p]+1.2)
    # P&L operating
    for y in FCST:
        nsv=M["net_revenue"][y]
        M["bev_gp"][y]=nsv*A["gross_margin_nsv"][y]
        M["cogs"][y]=nsv-M["bev_gp"][y]
        M["ad_promo"][y]=nsv*A["ad_promo_pct_nsv"][y]
        M["employee"][y]=nsv*A["employee_pct_nsv"][y]
        M["bev_ebitda"][y]=nsv*A["bev_ebitda_margin"][y]
        M["other_expense"][y]=M["bev_gp"][y]-M["ad_promo"][y]-M["employee"][y]-M["bev_ebitda"][y]
    # PPE schedule
    M["ppe_owned"]={"FY26":900.0}; M["rou_fc"]={"FY26":315.0}; M["lease_liab_fc"]={"FY26":M["lease_liab"]["FY26"]}
    M["capex_fc"]={};M["ppe_dep"]={};M["rou_add"]={};M["rou_dep"]={};M["lease_int"]={};M["lease_pay"]={}
    for y in FCST:
        p=prev(y)
        capex=M["net_revenue"][y]*A["capex_pct_nsv"][y]; M["capex_fc"][y]=capex
        dep_ppe=(M["ppe_owned"][p]+capex*0.5)*A["ppe_dep_rate"][y]; M["ppe_dep"][y]=dep_ppe
        M["ppe_owned"][y]=M["ppe_owned"][p]+capex-dep_ppe
        add=M["net_revenue"][y]*A["lease_add_pct_nsv"][y]; M["rou_add"][y]=add
        rou_dep=(M["rou_fc"][p]+add*0.5)*A["rou_dep_rate"][y]; M["rou_dep"][y]=rou_dep
        M["rou_fc"][y]=M["rou_fc"][p]+add-rou_dep
        opening=M["lease_liab_fc"][p]
        interest=opening*A["lease_int_rate"][y]; payment=opening*A["lease_pay_pct"][y]
        M["lease_int"][y]=interest; M["lease_pay"][y]=payment
        M["lease_liab_fc"][y]=opening+add+interest-payment
        M["lease_liab"][y]=M["lease_liab_fc"][y]
        M["ppe"][y]=M["ppe_owned"][y]+M["rou_fc"][y]; M["rou"][y]=0
        M["dep_amort"][y]=M["ppe_dep"][y]+M["rou_dep"][y]+15
    # BS operating items
    for y in FCST:
        p=prev(y); nsv=M["net_revenue"][y]; gr=M["gross_revenue"][y]; cogs=M["cogs"][y]
        M["receivables_cur"][y]=A["recv_days_gross"][y]/365*gr
        M["receivables_nc"][y]=nsv*A["recv_nc_pct_nsv"][y]
        M["inventory"][y]=A["inv_days_cogs"][y]/365*cogs
        M["trade_payables"][y]=A["pay_days_cogs"][y]/365*cogs
        M["provisions"][y]=nsv*A["provisions_pct_nsv"][y]
        M["other_cur_liab"][y]=nsv*A["other_cl_pct_nsv"][y]
        M["cur_tax_liab"][y]=M["cur_tax_liab"][p]*1.05
        M["other_fin_liab"][y]=M["other_fin_liab"][p]
        for k in ["cwip","goodwill","intangibles","inv_property","jv_invest","dta_net","dtl","other_fin_assets"]:
            M[k][y]=M[k][p]
        M["cur_tax_assets"][y]=0
        M["other_assets"][y]=1191+667.0*(nsv/M["net_revenue"]["FY26"])
        M["borrowings"][y]=6; M["assets_hfs"][y]=0; M["liab_hfs"][y]=0
    # treasury & below-EBITDA (iterate for circularity)
    M["treasury"]={}; M["ebit"]={}; M["total_liab"]={}
    base_t=M["cash"]["FY26"]+M["bank_deposits"]["FY26"]+M["cur_investments"]["FY26"]
    M["treasury"]["FY26"]=base_t
    for y in FCST: M["treasury"][y]=base_t
    for _ in range(60):
        for y in FCST:
            p=prev(y)
            avg=0.5*(M["treasury"][p]+M["treasury"][y])
            M["other_income"][y]=avg*A["treasury_yield"][y]+60
            M["finance_cost"][y]=M["lease_int"][y]+A["finance_cost_other"][y]
            M["jv_share"][y]=0; M["exceptional"][y]=0
            ebit=M["bev_ebitda"][y]-M["dep_amort"][y]; M["ebit"][y]=ebit
            pbt=ebit+M["other_income"][y]-M["finance_cost"][y]
            M["pbt"][y]=pbt; M["tax"][y]=pbt*A["tax_rate"][y]
            M["pat_continuing"][y]=pbt-M["tax"][y]; M["disc_ops_pat"][y]=0
            M["pat_reported"][y]=M["pat_continuing"][y]
            M["eps_reported"][y]=M["pat_reported"][y]/SHARES
            M["dividend_paid"][y]=M["pat_reported"][y]*A["payout"][y]
            M["equity"][y]=M["equity"][p]+M["pat_reported"][y]-M["dividend_paid"][y]
            nontreas=(M["ppe"][y]+M["cwip"][y]+M["goodwill"][y]+M["intangibles"][y]+M["inv_property"][y]
                      +M["jv_invest"][y]+M["inventory"][y]+M["receivables_cur"][y]+M["receivables_nc"][y]
                      +M["other_fin_assets"][y]+M["dta_net"][y]+M["cur_tax_assets"][y]+M["other_assets"][y]+M["assets_hfs"][y])
            tl=(M["borrowings"][y]+M["lease_liab"][y]+M["trade_payables"][y]+M["dtl"][y]+M["provisions"][y]
                +M["cur_tax_liab"][y]+M["other_fin_liab"][y]+M["other_cur_liab"][y]+M["liab_hfs"][y])
            M["total_liab"][y]=tl
            M["treasury"][y]=M["equity"][y]+tl-nontreas
            M["total_assets"][y]=nontreas+M["treasury"][y]
            M["total_liab_eq"][y]=M["equity"][y]+tl
    for y in FCST:
        t=M["treasury"][y]
        M["cash"][y]=t*0.27; M["bank_deposits"][y]=t*0.36; M["cur_investments"][y]=t*0.37
    return M
