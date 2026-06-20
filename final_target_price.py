"""Final blended Target Price for USL.
Reads the implied per-share values DIRECTLY from 'United Spirits Excell.xlsx'
(DCF sheet C28; Comparables 'Valuation using Mean Multiples', FY2027 row 42).
Applies justified, non-equal weights and computes the weighted target price."""
import openpyxl

wb = openpyxl.load_workbook("/projects/sandbox/USLL/United Spirits Excell.xlsx", data_only=True)
dcf = wb["DCF"]; cmp_ws = wb["Comparables"]

# --- Current market price (from Comparables C12 Mkt Cap / shares) ---
mktcap = cmp_ws["C12"].value          # 96,196 cr
shares = dcf["C27"].value             # 72.74 cr
CMP = mktcap / shares

# --- DCF intrinsic value (DCF!C28) ---
dcf_px = dcf["C28"].value             # 599.08

# --- Relative valuation: peer-MEAN multiples, FY2027E (1-yr forward), row 42 ---
# columns: J=EV/Revenue, K=EV/EBITDA, L=EV/EBIT, M=P/E
evrev = cmp_ws["J42"].value
evebitda = cmp_ws["K42"].value
evebit = cmp_ws["L42"].value
pe = cmp_ws["M42"].value

methods = [
    ("EV/EBITDA",            evebitda, 0.35),
    ("DCF (FCFF, WACC 11%)", dcf_px,   0.25),
    ("EV/EBIT",              evebit,   0.15),
    ("P/E",                  pe,       0.15),
    ("EV/Revenue",           evrev,    0.10),
]

print(f"CMP = MktCap {mktcap:,.0f} / shares {shares} = Rs {CMP:,.0f}\n")
print(f"{'Methodology':<24}{'Implied Px':>12}{'Weight':>9}{'Contribution':>14}")
tp = 0.0
for n, px, w in methods:
    tp += px*w
    print(f"{n:<24}{px:>12,.0f}{w:>8.0%}{px*w:>14,.1f}")
print("-"*59)
print(f"{'FINAL TARGET PRICE':<24}{'':>12}{sum(w for *_,w in methods):>8.0%}{tp:>14,.0f}")
print(f"\nTarget Price      : Rs {tp:,.0f}")
print(f"Upside/(Downside) : {tp/CMP-1:+.1%}")

# reference: relative-only blend (ex-DCF)
rel = [(p,w) for n,p,w in methods if not n.startswith('DCF')]
rw = sum(w for _,w in rel)
print(f"\nRelative-only blend: Rs {sum(p*w for p,w in rel)/rw:,.0f}")
print(f"DCF intrinsic      : Rs {dcf_px:,.0f}  ({dcf_px/CMP-1:+.1%})")
