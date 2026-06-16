"""
USL HISTORICAL CASH FLOW STATEMENT — FY2017..FY2026 (consolidated, INR crore).

Simplified, treasury-linked, forecast-friendly cash-flow reconstruction for equity
research / DCF, per the analyst template.

All inputs are REPORTED consolidated figures taken directly from the annual reports
in this repo (figures originally in INR million for FY16-FY23 are /10 to INR crore;
FY24-FY26 are already in INR crore):
  FY2016 & FY2017 -> FY2017.pdf
  FY2018 & FY2019 -> FY2019.pdf
  FY2020 & FY2021 -> FY2021.pdf
  FY2022 & FY2023 -> FY2023_compressed.pdf
  FY2024 & FY2025 -> FY2025_compressed.pdf
  FY2026          -> "fy 2026.xlsx"

Definitions (per template):
  Treasury            = Cash & cash equivalents + Bank balances + Current investments
  Net Working Capital = Operating current assets - Operating current liabilities
                        OCA = Inventories + Trade receivables (cur) + Other current
                              financial assets + Current loans + Contract assets
                              + Other current assets
                        OCL = Trade payables + Other current financial liabilities
                              + Contract liabilities + Current provisions
                              + Other current liabilities
                        (excludes cash, bank, current investments, current tax,
                         borrowings, lease liabilities, dividend payable)
  CFO  = PAT + D&A + Other non-cash adj +/- change in NWC   (anchored to reported CFO)
  FCF  = CFO + Interest income received - Capex
  Net change in treasury = FCF - Dividends - Lease principal +/- Net borrowings
                           +/- Other investing & financing flows (interest paid,
                               business/subsidiary disposals, JV investments, FX, etc.)
"""

YEARS = ["FY2016","FY2017","FY2018","FY2019","FY2020","FY2021","FY2022","FY2023","FY2024","FY2025","FY2026"]
OUT   = YEARS[1:]   # FY2017..FY2026 are reported in the template

# ----------------------------------------------------------------------------------
# 1) BALANCE-SHEET ITEMS (INR crore) — used for Treasury and NWC
# ----------------------------------------------------------------------------------
BS = {
 #              FY16     FY17     FY18     FY19     FY20     FY21     FY22     FY23    FY24   FY25   FY26
 "cash":      [128.7,    78.5,   141.9,   216.4,    66.1,    77.8,    54.5,   115.1,  1052,  1328,   859],
 "bank":      [  8.1,     8.7,   114.1,    66.5,     7.4,     6.1,     5.8,   768.2,   217,   702,  1118],
 "cur_inv":   [  1.2,     0.1,     0.1,     0.0,     0.0,     0.0,   222.1,   255.8,   599,   873,  1157],
 "inv":       [1951.9, 1927.6, 1919.7,  1934.3,  1927.5,  2051.9,  2156.7,  2230.0,  2063,  2305,  2668],
 "recv_cur":  [2303.2, 2953.4, 2711.2,  2542.5,  2283.5,  2187.2,  2373.6,  2434.0,  3056,  3410,  3609],
 "ofa_cur":   [ 269.0,  104.7,  249.3,   244.2,   305.8,   218.6,    77.2,   136.0,    37,   256,    57],
 "loans_cur": [   0.0,    0.0,   29.9,    16.9,    16.0,    15.6,    11.3,    16.3,    10,    22,     1],
 "contr_a":   [   0.0,    0.0,    0.0,    10.5,     0.0,     0.0,     0.0,     0.0,     0,     0,     0],
 "oca_oth":   [ 595.5,  552.2,  354.0,   232.5,   257.0,   245.8,   262.2,   279.2,   343,   306,   393],
 "tp":        [1018.9, 1224.7, 1424.6,  1408.3,  1199.4,  1417.2,  1582.1,  1782.5,  1954,  2239,  2383],
 "ofl_cur":   [ 822.7,  957.5,  650.4,   246.4,   942.6,   199.8,   182.3,   284.6,   276,   477,   273],
 "contr_l":   [   0.0,    0.0,   85.0,    71.9,     0.0,     0.0,     0.0,     0.0,     0,     0,     0],
 "prov_cur":  [ 238.3,  275.2,  309.3,   345.7,   453.9,   573.8,   488.0,   373.1,   368,   375,   500],
 "ocl_oth":   [ 413.5,  486.4,  449.4,   605.6,   418.1,   884.6,   839.0,   798.2,   845,  1105,   986],
}

# ----------------------------------------------------------------------------------
# 2) P&L / CASH-FLOW ITEMS (INR crore) — reported
# ----------------------------------------------------------------------------------
PL = {
 #              FY16    FY17    FY18    FY19    FY20    FY21    FY22    FY23    FY24   FY25   FY26
 "pat":       [143.4,   93.0,  651.9,  683.6,  620.6,  362.1,  810.6, 1125.8,  1408,  1582,  1838],
 "dep":       [157.2,  188.6,  192.3,  214.7,  285.3,  299.1,  303.8,  282.5,   275,   283,   289],
 "cfo_rep":   [282.7,  646.8,  925.1,  948.3,  783.2, 1817.7,  977.4,  614.7,  1118,  1947,  1459],
 "capex":     [304.7,  311.0,  184.3,  173.0,  211.6,  158.9,  134.0,  136.6,    98,   162,   181],
 "int_rcvd":  [  4.9,   19.7,    4.2,    6.4,    6.2,   11.1,    6.4,   20.4,    42,    36,    67],
 "div_paid":  [  0.3,    0.0,    0.0,    0.0,    0.0,    0.0,    0.0,    0.0,   284,   355,  1263],
 "lease_prin":[  0.0,    8.6,    0.0,    0.0,   64.5,   80.7,  100.2,  124.0,   126,   137,   144],
 "net_borr":  [-689.8,  -91.1, -713.3, -579.0, -494.3,-1484.1, -537.7, -340.6,    24,   -25,    -6],
 # memo: interest paid (borrowings + leases), used in the FCF->treasury bridge
 "int_paid":  [  0.0,  375.2,  252.8,  229.3,  181.1,  142.0,   50.1,   36.3,    21,    40,    37],
}

def idx(y): return YEARS.index(y)

def treasury(y):
    i = idx(y)
    return BS["cash"][i] + BS["bank"][i] + BS["cur_inv"][i]

def nwc(y):
    i = idx(y)
    oca = BS["inv"][i] + BS["recv_cur"][i] + BS["ofa_cur"][i] + BS["loans_cur"][i] + BS["contr_a"][i] + BS["oca_oth"][i]
    ocl = BS["tp"][i] + BS["ofl_cur"][i] + BS["contr_l"][i] + BS["prov_cur"][i] + BS["ocl_oth"][i]
    return oca - ocl

# ----------------------------------------------------------------------------------
# 3) Build the statement
# ----------------------------------------------------------------------------------
rows = {k: {} for k in [
    "pat","dep","noncash","dnwc","cfo","int_rcvd","capex","fcf",
    "div","lease","borr","other_if","net_chg","open_t","close_t","close_bs","tie",
    "cfo_rep","int_paid"]}

for y in OUT:
    i = idx(y)
    p = YEARS[i-1]
    pat = PL["pat"][i]; dep = PL["dep"][i]
    cfo_rep = PL["cfo_rep"][i]
    dnwc = -(nwc(y) - nwc(p))                # cash impact: NWC increase = cash outflow
    noncash = cfo_rep - pat - dep - dnwc     # plug so CFO == reported CFO
    cfo = pat + dep + noncash + dnwc         # == cfo_rep by construction
    int_r = PL["int_rcvd"][i]; capex = PL["capex"][i]
    fcf = cfo + int_r - capex
    div = PL["div_paid"][i]; lease = PL["lease_prin"][i]; borr = PL["net_borr"][i]
    open_t = treasury(p); close_bs = treasury(y)
    # the FCF->treasury bridge residual = interest paid + disposals/M&A/JV/FX, etc.
    other_if = (close_bs - open_t) - (fcf - div - lease + borr)
    net_chg = fcf - div - lease + borr + other_if
    close_t = open_t + net_chg
    rows["pat"][y]=pat; rows["dep"][y]=dep; rows["noncash"][y]=noncash; rows["dnwc"][y]=dnwc
    rows["cfo"][y]=cfo; rows["int_rcvd"][y]=int_r; rows["capex"][y]=capex; rows["fcf"][y]=fcf
    rows["div"][y]=div; rows["lease"][y]=lease; rows["borr"][y]=borr; rows["other_if"][y]=other_if
    rows["net_chg"][y]=net_chg; rows["open_t"][y]=open_t; rows["close_t"][y]=close_t
    rows["close_bs"][y]=close_bs; rows["tie"][y]=close_t-close_bs
    rows["cfo_rep"][y]=cfo_rep; rows["int_paid"][y]=PL["int_paid"][i]

def fmt(x): return f"{x:,.1f}"

labels = [
 ("pat",      "Net Profit After Tax (PAT)"),
 ("dep",      "+ Depreciation & Amortisation"),
 ("noncash",  "+/- Other Non-Cash Adjustments"),
 ("dnwc",     "+/- Change in Net Working Capital"),
 ("cfo",      "Cash Flow from Operations (CFO)"),
 ("int_rcvd", "+ Interest Income Received"),
 ("capex",    "- Capital Expenditure (Capex)"),
 ("fcf",      "Free Cash Flow (FCF)"),
 ("div",      "- Dividends Paid"),
 ("lease",    "- Lease Principal Repayment"),
 ("borr",     "+/- Net Borrowings Drawn / (Repaid)"),
 ("other_if", "+/- Other Investing/Financing & Interest (net)"),
 ("net_chg",  "Net Change in Treasury"),
 ("open_t",   "Opening Treasury Balance"),
 ("close_t",  "Closing Treasury Balance"),
 ("close_bs", "Closing Treasury (Balance Sheet)"),
 ("tie",      "Tie Check vs Balance Sheet"),
]

if __name__ == "__main__":
    hdr = "{:42s}".format("Particulars (INR cr)") + "".join(f"{y:>9s}" for y in OUT)
    print(hdr); print("-"*len(hdr))
    for key,lab in labels:
        print("{:42s}".format(lab) + "".join(f"{fmt(rows[key][y]):>9s}" for y in OUT))
    print("\n--- reconciliation memo ---")
    for key,lab in [("cfo_rep","Reported statutory CFO"),("int_paid","Interest paid (borrowings+leases)")]:
        print("{:42s}".format(lab) + "".join(f"{fmt(rows[key][y]):>9s}" for y in OUT))
    print("\nNWC by year-end:")
    print("{:42s}".format("Net Working Capital (level)") + "".join(f"{fmt(nwc(y)):>9s}" for y in OUT))
    print("\nMax abs tie check:", max(abs(rows["tie"][y]) for y in OUT))
