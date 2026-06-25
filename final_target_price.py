"""USL — REAL valuation. Primary: forward EV/EBITDA (the only multiple on which USL
and peers are genuinely comparable). DCF = intrinsic anchor. P/E = USL-normalised cross-check
(peer P/E rejected: United Breweries is beer; Radico/ABDL are higher-growth, so peer P/E is not
a valid comparable). Reads inputs from United Spirits Excell.xlsx."""
import openpyxl
wb = openpyxl.load_workbook("/projects/sandbox/USLL/United Spirits Excell.xlsx", data_only=True)
C = wb["Comparables"]; D = wb["DCF"]; B = wb["Beta Calculation"]

shares   = D["C27"].value                         # 72.74 cr
cash     = D["C24"].value                          # 3,134
debt     = -D["C25"].value                         # 407
net_cash = cash - debt                             # 2,727
CMP      = B["F5"].value                            # 1,246.7  latest weekly close 31-May-2026
stale_cap= C["C12"].value                           # 96,196 (=> 1,322/sh) -> stale

# FY2027E (1-yr forward) USL operating metrics
ebitda27 = C["K12"].value      # 2,547.32
eps27    = C["M12"].value / shares    # PAT 1,846.99 / shares
# FY2028E for roll-forward
ebitda28 = C["P12"].value      # 2,824.49
eps28    = C["R12"].value / shares

# Peer FY27 EV/EBITDA stats (from model)
peer_mean, peer_med, peer_lq, peer_uq = C["K27"].value, C["K28"].value, C["K29"].value, C["K30"].value

def ev_ebitda_price(ebitda, mult):
    return (ebitda*mult + net_cash) / shares

# --- Target multiples (justified) ---
TGT_EVEBITDA = 38.0     # ~ peer mean 37.3 / median 39.2; slight discount to median for slower vol growth
TGT_PE       = 52.0     # USL's own ~3-yr avg & peer lower-quartile; peer-mean 72x rejected (not comparable)
dcf_px       = D["C28"].value   # 599

p_evebitda = ev_ebitda_price(ebitda27, TGT_EVEBITDA)
p_pe       = TGT_PE * eps27
p_dcf      = dcf_px

methods = [
    ("EV/EBITDA (38.0x FY27E)",          p_evebitda, 0.60),
    ("P/E (52.0x FY27E, USL-normalised)", p_pe,       0.25),
    ("DCF (FCFF, WACC 11%/g 6%)",         p_dcf,      0.15),
]
print(f"CMP (latest close 31-May-26): Rs {CMP:,.1f}   [stale comps cap implies Rs {stale_cap/shares:,.0f}]")
print(f"Net cash: {net_cash:,.0f} cr | shares {shares} cr | FY27E EBITDA {ebitda27:,.0f} | FY27E EPS {eps27:,.2f}\n")

# USL implied trading multiples at the LATEST price
ev_now = CMP*shares - net_cash
print(f"At CMP {CMP:,.0f}: USL trades FY27E EV/EBITDA {ev_now/ebitda27:.1f}x | P/E {CMP/eps27:.1f}x")
print(f"Peer FY27 EV/EBITDA: mean {peer_mean:.1f}x  median {peer_med:.1f}x  LQ {peer_lq:.1f}x  UQ {peer_uq:.1f}x\n")

print(f"{'Methodology':<38}{'Implied':>10}{'Wt':>6}{'Contrib':>10}")
tp=0.0
for n,p,w in methods:
    tp+=p*w; print(f"{n:<38}{p:>10,.0f}{w:>6.0%}{p*w:>10,.1f}")
print("-"*64)
print(f"{'BASE-CASE TARGET PRICE':<38}{'':>10}{'100%':>6}{tp:>10,.0f}")
print(f"\nTarget Rs {tp:,.0f} vs CMP Rs {CMP:,.0f}  =>  {tp/CMP-1:+.1%}\n")

# Scenario analysis (bear / base / bull) on EV/EBITDA + DCF bookend
bear = ev_ebitda_price(ebitda27, peer_lq)          # lower-quartile multiple
bull = ev_ebitda_price(ebitda28, peer_uq)          # upper-quartile on FY28E (roll-forward + re-rating)
print("Scenario bookends:")
print(f"  Bear  (DCF floor)                : Rs {dcf_px:,.0f}  ({dcf_px/CMP-1:+.0%})")
print(f"  Bear  (LQ {peer_lq:.0f}x EV/EBITDA FY27): Rs {bear:,.0f}  ({bear/CMP-1:+.0%})")
print(f"  Base                              : Rs {tp:,.0f}  ({tp/CMP-1:+.0%})")
print(f"  Bull  (UQ {peer_uq:.0f}x EV/EBITDA FY28): Rs {bull:,.0f}  ({bull/CMP-1:+.0%})")
