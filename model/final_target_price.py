"""Final blended Target Price for USL: DCF + 4 trading multiples, weighted.
All implied prices computed on FY28E metrics (1-yr-forward) with the FY26 net-cash bridge."""
import usl_engine as E
from usl_engine import SHARES
import usl_build as B  # reuse the exact DCF used in the report

A = E.base_assumptions(); M = E.run(A)
def ebit(y): return M["bev_ebitda"][y] - M["dep_amort"][y]
net_cash = -(M["borrowings"]["FY26"] + M["lease_liab"]["FY26"]
             - M["cash"]["FY26"] - M["bank_deposits"]["FY26"] - M["cur_investments"]["FY26"])
rcb = 2500.0
CMP = 1259.0
fy = "FY28E"
nsv, ebitda, ebit28, eps = M["net_revenue"][fy], M["bev_ebitda"][fy], ebit(fy), M["eps_reported"][fy]

def ev_price(metric, mult): return (metric*mult + net_cash + rcb) / SHARES

# DCF target (same as report base case)
dcf_tp = B.dcf(0.1025, 0.0625, M)["tp"]

# Target multiples (forward, FY28E) - justified vs peers + USL history
methods = [
    ("DCF (2-stage FCFF)",      dcf_tp,                       0.30),
    ("EV/EBITDA",               ev_price(ebitda, 34.0),       0.28),
    ("P/E",                     eps*50.0,                     0.22),
    ("EV/EBIT",                 ev_price(ebit28, 38.0),       0.12),
    ("EV/Revenue",              ev_price(nsv, 6.0),           0.08),
]

print(f"{'Methodology':<24}{'Implied Px':>12}{'Weight':>9}{'Contribution':>14}")
tp_final = 0.0
for name, px, w in methods:
    contrib = px*w
    tp_final += contrib
    print(f"{name:<24}{px:>12,.0f}{w:>8.0%}{contrib:>14,.1f}")
print("-"*59)
print(f"{'FINAL TARGET PRICE':<24}{'':>12}{sum(w for _,_,w in methods):>8.0%}{tp_final:>14,.0f}")
print()
print(f"CMP (10-Jun-2026)   : Rs {CMP:,.0f}")
print(f"Target Price        : Rs {tp_final:,.0f}")
print(f"Upside/(Downside)   : {tp_final/CMP-1:+.1%}")
print()
# relative-only blend (ex-DCF) for reference
rel = [(p,w) for n,p,w in methods if not n.startswith("DCF")]
rw = sum(w for _,w in rel)
rel_only = sum(p*w for p,w in rel)/rw
print(f"Relative-only blend : Rs {rel_only:,.0f}  ({rel_only/CMP-1:+.1%})")
print(f"DCF value           : Rs {dcf_tp:,.0f}  ({dcf_tp/CMP-1:+.1%})")
