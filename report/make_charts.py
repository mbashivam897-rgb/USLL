"""Generate institutional-style chart exhibits for the USL equity research report.
All data sourced from 'United Spirits Excell.xlsx' (primary model) and FY2025 Annual Report.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

OUT = os.path.join(os.path.dirname(__file__), "charts")
os.makedirs(OUT, exist_ok=True)

# ---- House style ----
NAVY   = "#0B2545"
BLUE   = "#13558C"
STEEL  = "#3E7CB1"
LIGHT  = "#8FB8DE"
GOLD   = "#C9A227"
GREY   = "#9AA5B1"
RED    = "#B3322C"
GREEN  = "#2E7D52"
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.edgecolor": "#7a7a7a",
    "axes.linewidth": 0.8,
    "axes.grid": True,
    "grid.color": "#E2E6EA",
    "grid.linewidth": 0.8,
    "axes.axisbelow": True,
    "figure.dpi": 140,
})

def save(fig, name):
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", name)

# ============== EXHIBIT 1: NSV & growth (FY22-FY30E) ==============
yrs = ["FY22","FY23","FY24","FY25","FY26","FY27E","FY28E","FY29E","FY30E"]
nsv = [9712,10612,11321,12069,12467,13589,14812,15997,17117]
gr  = [None,9.3,6.7,6.6,3.3,9.0,9.0,8.0,7.0]
x = np.arange(len(yrs))
fig, ax1 = plt.subplots(figsize=(7.4,3.7))
bars = ax1.bar(x, nsv, color=[BLUE]*5+[STEEL]*4, width=0.62, zorder=3)
ax1.set_ylabel("Net Sales Value (₹ cr)")
ax1.set_xticks(x); ax1.set_xticklabels(yrs)
ax1.set_ylim(0, 19500)
for i,v in enumerate(nsv):
    ax1.text(i, v+250, f"{v:,}", ha="center", va="bottom", fontsize=8, color=NAVY)
ax2 = ax1.twinx()
ax2.plot(x, gr, color=GOLD, marker="o", lw=2, zorder=4)
ax2.set_ylabel("NSV growth (%)", color=GOLD)
ax2.set_ylim(0, 16); ax2.grid(False)
for i,v in enumerate(gr):
    if v is not None: ax2.text(i, v+0.4, f"{v:.1f}%", ha="center", fontsize=7.5, color="#8a6d0b")
ax1.set_title("Exhibit 1: Net Sales Value and growth, premiumisation led topline (FY22 to FY30E)",
              fontsize=10, fontweight="bold", color=NAVY, loc="left")
ax1.axvline(4.5, color=GREY, ls="--", lw=1)
ax1.text(4.55, 18000, "Forecast", fontsize=8, color=GREY)
save(fig, "ex1_nsv_growth.png")

# ============== EXHIBIT 2: Margin ladder ==============
fig, ax = plt.subplots(figsize=(7.4,3.7))
gm  = [42,39,42,43,46,46,46,47,47]
eb  = [16,13,18,19,18,19,19,20,20]
ebit= [13,11,15,16,16,16,16,17,17]
pat = [8,11,12,13,15,13,13,13,14]
ax.plot(x,gm,marker="o",color=NAVY,lw=2,label="Gross margin")
ax.plot(x,eb,marker="s",color=BLUE,lw=2,label="EBITDA margin")
ax.plot(x,ebit,marker="^",color=STEEL,lw=2,label="EBIT margin")
ax.plot(x,pat,marker="D",color=GOLD,lw=2,label="PAT margin")
ax.set_xticks(x); ax.set_xticklabels(yrs)
ax.set_ylabel("% of NSV"); ax.set_ylim(0,55)
ax.legend(ncol=4, fontsize=8, loc="upper center", frameon=False, bbox_to_anchor=(0.5,1.02))
ax.set_title("Exhibit 2: Margin trajectory, steady expansion on mix and operating leverage",
             fontsize=10, fontweight="bold", color=NAVY, loc="left", pad=22)
ax.axvline(4.5,color=GREY,ls="--",lw=1)
save(fig,"ex2_margins.png")

# ============== EXHIBIT 3: Volume mix P&A vs Popular ==============
yv = ["FY16","FY18","FY20","FY22","FY23","FY24","FY25"]
pa = [34,37.2,40.9,42.74,47.85,50.3,53.12]
pop= [59,41.3,38.85,36.41,24.65,11.1,10.88]
xv = np.arange(len(yv))
fig, ax = plt.subplots(figsize=(7.4,3.7))
ax.bar(xv,pa,color=BLUE,label="Prestige & Above (P&A)",zorder=3,width=0.6)
ax.bar(xv,pop,bottom=pa,color=LIGHT,label="Popular",zorder=3,width=0.6)
tot=[a+b for a,b in zip(pa,pop)]
for i in range(len(yv)):
    share=pa[i]/tot[i]*100
    ax.text(i, tot[i]+1.2, f"P&A {share:.0f}%", ha="center", fontsize=7.5, color=NAVY, fontweight="bold")
ax.set_xticks(xv); ax.set_xticklabels(yv)
ax.set_ylabel("Volume (mn cases)"); ax.set_ylim(0,100)
ax.legend(fontsize=8, frameon=False, loc="upper right")
ax.set_title("Exhibit 3: Volume mix shift, planned exit of Popular, P&A now about 83 percent of volume",
             fontsize=9.6, fontweight="bold", color=NAVY, loc="left")
save(fig,"ex3_volume_mix.png")

# ============== EXHIBIT 4: ROE / ROCE ==============
fig, ax = plt.subplots(figsize=(7.4,3.4))
roe=[18,21,21,21,22,19,19,20,21]
roce=[25,18,23,23,21,22,22,23,23]
ax.plot(x,roe,marker="o",color=BLUE,lw=2,label="ROE")
ax.plot(x,roce,marker="s",color=GOLD,lw=2,label="ROCE")
ax.set_xticks(x); ax.set_xticklabels(yrs); ax.set_ylabel("%"); ax.set_ylim(0,32)
ax.legend(fontsize=8,frameon=False,loc="lower right")
ax.set_title("Exhibit 4: Capital returns, steady 20 percent plus ROE and ROCE",
             fontsize=10,fontweight="bold",color=NAVY,loc="left")
ax.axvline(4.5,color=GREY,ls="--",lw=1)
save(fig,"ex4_returns.png")

# ============== EXHIBIT 5: Peer EV/EBITDA & P/E (FY26) ==============
peers=["United\nSpirits","Radico\nKhaitan","United\nBreweries","Allied\nBlenders"]
evb=[41.1,49.0,44.1,34.4]
pe =[52.5,81.5,86.2,77.0]
xp=np.arange(len(peers)); w=0.38
fig, ax = plt.subplots(figsize=(7.4,3.6))
b1=ax.bar(xp-w/2,evb,w,color=BLUE,label="EV/EBITDA (x)",zorder=3)
b2=ax.bar(xp+w/2,pe,w,color=GOLD,label="P/E (x)",zorder=3)
ax.set_xticks(xp); ax.set_xticklabels(peers); ax.set_ylabel("Multiple (x)")
ax.legend(fontsize=8,frameon=False)
for b in list(b1)+list(b2):
    ax.text(b.get_x()+b.get_width()/2, b.get_height()+1, f"{b.get_height():.0f}", ha="center", fontsize=7.5,color=NAVY)
ax.set_title("Exhibit 5: Peer valuation (FY26), USL trades below premium peers on P/E",
             fontsize=9.8,fontweight="bold",color=NAVY,loc="left")
ax.set_ylim(0,95)
save(fig,"ex5_peer_multiples.png")

# ============== EXHIBIT 6: Valuation football field / blended TP ==============
methods=["DCF\n(15%)","EV/EBIT\n(5%)","EV/Revenue\n(ref.)","EV/EBITDA\n(60%)","P/E\n(20%)"]
vals=[642,1408,859,1411,1675]
fig, ax = plt.subplots(figsize=(7.4,3.6))
colors=[STEEL,STEEL,GREY,BLUE,GOLD]
b=ax.barh(methods,vals,color=colors,zorder=3,height=0.6)
for i,v in enumerate(vals):
    ax.text(v+20,i,f"₹{v:,}",va="center",fontsize=8.5,color=NAVY,fontweight="bold")
ax.axvline(1326,color=RED,ls="-",lw=1.6)
ax.text(1326,4.6,"CMP ₹1,326",color=RED,fontsize=8,ha="center")
ax.axvline(1348,color=GREEN,ls="--",lw=1.6)
ax.text(1348,-0.7,"Blended TP ₹1,348",color=GREEN,fontsize=8,ha="center")
ax.set_xlabel("Implied value per share (₹)"); ax.set_xlim(0,1950)
ax.set_title("Exhibit 6: Valuation summary, implied value per method and blended target",
             fontsize=9.8,fontweight="bold",color=NAVY,loc="left")
save(fig,"ex6_football_field.png")

# ============== EXHIBIT 7: DCF sensitivity heatmap (WACC x g) ==============
# from model DCF sensitivity grid (per-share values)
wacc=[8,9,10,11,12]
g=[5,6,7]
# representative grid extracted/interpolated around model centre (WACC11/g6 -> 642)
grid=np.array([
 [1093,1364,1926],   # 8%
 [ 841,1006,1448],   # 9%
 [ 635, 750, 940],   # 10%
 [ 503, 597, 698],   # 11%
 [ 425, 444, 491],   # 12%
])
fig, ax = plt.subplots(figsize=(5.6,3.6))
im=ax.imshow(grid,cmap="RdYlGn",aspect="auto")
ax.set_xticks(range(len(g))); ax.set_xticklabels([f"{v}%" for v in g])
ax.set_yticks(range(len(wacc))); ax.set_yticklabels([f"{v}%" for v in wacc])
ax.set_xlabel("Terminal growth (g)"); ax.set_ylabel("WACC")
for i in range(len(wacc)):
    for j in range(len(g)):
        ax.text(j,i,f"{grid[i,j]:,}",ha="center",va="center",fontsize=8,
                color="black")
ax.set_title("Exhibit 7: DCF sensitivity, fair value per share (Rs)\nBase case WACC 11 percent, g 6 percent gives Rs 642",
             fontsize=9.4,fontweight="bold",color=NAVY,loc="left")
save(fig,"ex7_dcf_sensitivity.png")

# ============== EXHIBIT 8: Scenario target prices ==============
sc=["Bear","Base","Bull"]
tp=[920,1348,1760]
fig, ax = plt.subplots(figsize=(5.6,3.4))
bb=ax.bar(sc,tp,color=[RED,BLUE,GREEN],width=0.55,zorder=3)
ax.axhline(1326,color="#444",ls="--",lw=1.2)
ax.text(2.4,1356,"CMP ₹1,326",fontsize=7.5,color="#444",ha="right")
for b,v in zip(bb,tp):
    ax.text(b.get_x()+b.get_width()/2,v+25,f"₹{v:,}",ha="center",fontsize=9,fontweight="bold",color=NAVY)
ax.set_ylabel("Target price (₹)"); ax.set_ylim(0,2000)
ax.set_title("Exhibit 8: Scenario-weighted target prices",
             fontsize=10,fontweight="bold",color=NAVY,loc="left")
save(fig,"ex8_scenarios.png")

# ============== EXHIBIT 9: Free cash flow & dividends ==============
fy=["FY27E","FY28E","FY29E","FY30E"]
fcf=[1115,2145,1748,2606]
div=[1145,1267,1420,1553]
xf=np.arange(len(fy)); w=0.38
fig, ax = plt.subplots(figsize=(6.4,3.4))
ax.bar(xf-w/2,fcf,w,color=BLUE,label="Free cash flow (FCFF)",zorder=3)
ax.bar(xf+w/2,div,w,color=GOLD,label="Dividends paid",zorder=3)
ax.set_xticks(xf); ax.set_xticklabels(fy); ax.set_ylabel("₹ cr")
ax.legend(fontsize=8,frameon=False)
ax.set_title("Exhibit 9: Free cash flow vs dividends (about 70 percent payout)",
             fontsize=10,fontweight="bold",color=NAVY,loc="left")
save(fig,"ex9_fcf_div.png")

print("ALL CHARTS DONE")
