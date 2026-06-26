"""Illustrative diagrams + quarterly charts for the USL Stock Note style report.
Saves PNGs into report/charts/. No tildes or em-dashes in any label text.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Polygon, Circle
import numpy as np
import os

OUT = os.path.join(os.path.dirname(__file__), "charts")
os.makedirs(OUT, exist_ok=True)

NAVY="#0B2545"; BLUE="#13558C"; STEEL="#3E7CB1"; LIGHT="#8FB8DE"; PALE="#DCE7F2"
GOLD="#C9A227"; GOLDD="#8A6D0B"; GREY="#9AA5B1"; RED="#B3322C"; GREEN="#2E7D52"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":10,
    "axes.edgecolor":"#7a7a7a","axes.linewidth":0.8,"axes.grid":True,
    "grid.color":"#E6EAEE","axes.axisbelow":True,"figure.dpi":150})

def save(fig,name):
    fig.savefig(os.path.join(OUT,name),bbox_inches="tight",facecolor="white")
    plt.close(fig); print("saved",name)

def rbox(ax,x,y,w,h,text,fc,tc="white",fs=9,bold=True,ec=None):
    box=FancyBboxPatch((x,y),w,h,boxstyle="round,pad=0.02,rounding_size=0.06",
        fc=fc,ec=ec or fc,lw=1.2,zorder=2)
    ax.add_patch(box)
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",color=tc,fontsize=fs,
        fontweight="bold" if bold else "normal",zorder=3,wrap=True)

# ============ DIAGRAM A: Industry Value Chain (flow) ============
fig,ax=plt.subplots(figsize=(8.2,2.7)); ax.set_xlim(0,10); ax.set_ylim(0,3); ax.axis("off")
ax.text(5,2.82,"Exhibit: Indian Spirits Value Chain and Where the Margin Sits",
    ha="center",fontsize=11,fontweight="bold",color=NAVY)
stages=[("Inputs\n(ENA / grain\nspirit, scotch\nconcentrate,\nglass)",STEEL),
        ("Distillation\n& Blending\n(own + tie-up\nunits)",BLUE),
        ("Bottling &\nPackaging\n(40+ plants,\nin-state)",BLUE),
        ("State Route\nto Market\n(excise, govt\ncorporations)",GOLD),
        ("Retail\n(70,000+\noutlets, bars,\nhotels)",STEEL),
        ("Consumer\n(premium,\nexperiential\ndemand)",NAVY)]
w=1.42; gap=0.16; x=0.15
for i,(t,c) in enumerate(stages):
    rbox(ax,x,0.85,w,1.25,t,c,fs=7.6)
    if i<len(stages)-1:
        ax.add_patch(FancyArrowPatch((x+w,1.48),(x+w+gap,1.48),
            arrowstyle="-|>",mutation_scale=14,color=GREY,lw=2))
    x+=w+gap
ax.text(0.15+ (w+gap)*2 +w/2,0.5,"State excise duty (pass-through to states) is the single largest cost in the chain",
    ha="center",fontsize=7.5,color=RED,style="italic")
ax.text(0.15+(w+gap)*4+w/2,0.5,"USL value capture is highest in premium blending and brand equity",
    ha="center",fontsize=7.5,color=GREEN,style="italic")
save(fig,"dgm_value_chain.png")

# ============ DIAGRAM B: Premiumisation Pyramid ============
fig,ax=plt.subplots(figsize=(7.6,4.4)); ax.set_xlim(0,10); ax.set_ylim(0,6.4); ax.axis("off")
ax.text(5,6.15,"Exhibit: The Premiumisation Pyramid (price ladder and growth)",
    ha="center",fontsize=11,fontweight="bold",color=NAVY)
tiers=[  # (label, price band, 5yr CAGR, color, ybottom, half-width-bottom, half-width-top)
 ("LUXURY","Above Rs 2,000 / case","approx 30-35% CAGR",GOLD,4.7,1.0,0.1),
 ("PREMIUM","Rs 800 to 2,000","approx 15% CAGR",BLUE,3.4,2.1,1.0),
 ("PRESTIGE","Rs 400 to 800","approx 5% CAGR",STEEL,2.1,3.2,2.1),
 ("POPULAR","Below Rs 400","approx flat to negative",GREY,0.8,4.3,3.2)]
cx=5
for label,band,cagr,color,yb,hwb,hwt in tiers:
    pts=[(cx-hwb,yb),(cx+hwb,yb),(cx+hwt,yb+1.2),(cx-hwt,yb+1.2)]
    ax.add_patch(Polygon(pts,closed=True,fc=color,ec="white",lw=2,zorder=2))
    ax.text(cx,yb+0.6,label,ha="center",va="center",color="white",fontweight="bold",fontsize=11,zorder=3)
    ax.text(cx+hwb+0.25,yb+0.6,f"{band}   |   {cagr}",ha="left",va="center",fontsize=8.2,color=NAVY,zorder=3)
ax.annotate("Prestige & Above (P&A)\nnow about 89-90% of USL net sales",
    xy=(cx-1.6,4.4),xytext=(0.2,5.4),fontsize=8.5,color=GREEN,fontweight="bold",
    arrowprops=dict(arrowstyle="-|>",color=GREEN,lw=1.6))
ax.text(5,0.35,"Value migrates up the pyramid: small volume at the top, large share of industry profit",
    ha="center",fontsize=8,color=GREY,style="italic")
save(fig,"dgm_premium_pyramid.png")

# ============ DIAGRAM C: Porter's Five Forces ============
fig,ax=plt.subplots(figsize=(8.0,5.2)); ax.set_xlim(0,10); ax.set_ylim(0,9); ax.axis("off")
ax.text(5,8.7,"Exhibit: Porter's Five Forces, Indian Spirits Industry",
    ha="center",fontsize=11,fontweight="bold",color=NAVY)
rbox(ax,3.5,3.7,3.0,1.5,"COMPETITIVE\nRIVALRY\n(High and rising)",NAVY,fs=9.5)
rbox(ax,3.6,6.7,2.8,1.4,"New Entrants\n(Low threat)",GREEN,fs=8.5)
rbox(ax,3.6,0.5,2.8,1.4,"Substitutes\n(Moderate: beer,\nwine, no/low alc.)",GOLD,fs=8.2)
rbox(ax,0.2,3.8,2.8,1.3,"Supplier Power\n(Moderate: ENA,\nglass, scotch)",STEEL,fs=8.2)
rbox(ax,7.0,3.8,2.8,1.3,"Buyer Power\n(High: state\ncorporations)",RED,fs=8.5)
for (x0,y0),(x1,y1) in [((5,6.7),(5,5.2)),((5,3.7),(5,1.9)),((3.0,4.45),(3.5,4.45)),((7.0,4.45),(6.5,4.45))]:
    ax.add_patch(FancyArrowPatch((x0,y0),(x1,y1),arrowstyle="-|>",mutation_scale=13,color=GREY,lw=1.8))
ax.text(5,0.15,"Net: high entry barriers and premiumisation are positives; state control of pricing and rising rivalry are the constraints",
    ha="center",fontsize=7.6,color=GREY,style="italic")
save(fig,"dgm_porter.png")

# ============ DIAGRAM D: Competitive positioning map ============
fig,ax=plt.subplots(figsize=(7.6,5.0))
# x = premium mix / price positioning (low->high), y = FY26 EBITDA margin
players={
 "United Spirits":(8.8,18.3,2279,NAVY),
 "Radico Khaitan":(7.6,17.1,1038,BLUE),
 "Allied Blenders":(5.2,13.8,542,STEEL),
 "Tilaknagar Inds.":(4.3,16.0,300,GOLD),
 "United Breweries":(4.8,8.9,824,GREY)}
for name,(x,y,size,c) in players.items():
    ax.scatter(x,y,s=size*0.55,color=c,alpha=0.75,edgecolor="white",lw=1.5,zorder=3)
    ax.text(x,y-1.7,name,ha="center",fontsize=8.3,fontweight="bold",color=NAVY,zorder=4)
ax.set_xlim(3,10); ax.set_ylim(5,23)
ax.set_xlabel("Premium portfolio positioning  (lower  to  higher)",fontsize=9)
ax.set_ylabel("FY26 EBITDA margin (%)",fontsize=9)
ax.set_title("Exhibit: Competitive Positioning Map (bubble size = EBITDA, Rs cr)",
    fontsize=10.5,fontweight="bold",color=NAVY,loc="left")
ax.axhline(14,color=GREY,ls="--",lw=0.9); ax.axvline(6.5,color=GREY,ls="--",lw=0.9)
ax.text(9.6,21.5,"Premium\n+ high margin",ha="right",fontsize=7.5,color=GREEN,style="italic")
save(fig,"dgm_positioning.png")

# ============ DIAGRAM E: Industry size bar + consumption pie ============
fig,(a1,a2)=plt.subplots(1,2,figsize=(8.4,3.4))
yrs=["2021","2024","2030P"]; size=[52.4,58,77]
b=a1.bar(yrs,size,color=[LIGHT,STEEL,BLUE],width=0.6,zorder=3)
for bi,v in zip(b,size): a1.text(bi.get_x()+bi.get_width()/2,v+1,f"${v}bn",ha="center",fontsize=9,fontweight="bold",color=NAVY)
a1.set_ylim(0,90); a1.set_ylabel("Market size (US$ bn)")
a1.set_title("India alcobev market size",fontsize=10,fontweight="bold",color=NAVY,loc="left")
seg=["IMFL spirits","IMIL (country)","Beer","Wine / Others"]; share=[36.0,32.9,30.4,0.7]
cols=[BLUE,STEEL,GOLD,GREY]
a2.pie(share,labels=seg,autopct="%1.0f%%",colors=cols,startangle=90,
    textprops={"fontsize":8},wedgeprops={"edgecolor":"white","linewidth":1.2})
a2.set_title("Alcohol consumption mix",fontsize=10,fontweight="bold",color=NAVY)
fig.suptitle("Exhibit: A large, growing market dominated by Indian spirits",
    fontsize=11,fontweight="bold",color=NAVY,y=1.02)
save(fig,"dgm_industry_size.png")

# ============ QUARTERLY CHARTS ============
q=["Q1\nFY26","Q2\nFY26","Q3\nFY26","Q4\nFY26"]
nsv=[2549,3170,3683,3046]; ebitda=[415,672,618,591]; margin=[16.3,21.2,16.8,19.4]; pat=[417,464,418,539]
x=np.arange(4)

# Q-1: NSV bars + EBITDA margin line
fig,ax1=plt.subplots(figsize=(6.6,3.6))
bb=ax1.bar(x,nsv,color=BLUE,width=0.58,zorder=3)
for bi,v in zip(bb,nsv): ax1.text(bi.get_x()+bi.get_width()/2,v+50,f"{v:,}",ha="center",fontsize=8.5,color=NAVY)
ax1.set_ylabel("Net Sales Value (Rs cr)"); ax1.set_xticks(x); ax1.set_xticklabels(q); ax1.set_ylim(0,4300)
ax2=ax1.twinx(); ax2.plot(x,margin,color=GOLD,marker="o",lw=2.2,zorder=4)
for i,v in enumerate(margin): ax2.text(i,v+0.5,f"{v}%",ha="center",fontsize=8,color=GOLDD,fontweight="bold")
ax2.set_ylabel("EBITDA margin (%)",color=GOLDD); ax2.set_ylim(0,26); ax2.grid(False)
ax1.set_title("Exhibit: FY26 quarterly NSV and EBITDA margin",fontsize=10,fontweight="bold",color=NAVY,loc="left")
save(fig,"q_nsv_margin.png")

# Q-2: EBITDA and PAT bars
fig,ax=plt.subplots(figsize=(6.6,3.5)); w=0.38
b1=ax.bar(x-w/2,ebitda,w,color=BLUE,label="EBITDA (Rs cr)",zorder=3)
b2=ax.bar(x+w/2,pat,w,color=GOLD,label="PAT, consolidated (Rs cr)",zorder=3)
for bs in (b1,b2):
    for bi in bs: ax.text(bi.get_x()+bi.get_width()/2,bi.get_height()+8,f"{int(bi.get_height())}",ha="center",fontsize=8,color=NAVY)
ax.set_xticks(x); ax.set_xticklabels(q); ax.set_ylabel("Rs cr"); ax.set_ylim(0,780)
ax.legend(fontsize=8,frameon=False)
ax.set_title("Exhibit: FY26 quarterly EBITDA and PAT",fontsize=10,fontweight="bold",color=NAVY,loc="left")
save(fig,"q_ebitda_pat.png")

# Q-3: YoY growth lines (NSV, PAT)
nsv_yoy=[8.4,11.5,7.6,3.4]; pat_yoy=[-14,36,25,28]
fig,ax=plt.subplots(figsize=(6.6,3.4))
ax.plot(x,nsv_yoy,marker="o",color=BLUE,lw=2.2,label="NSV growth YoY")
ax.plot(x,pat_yoy,marker="s",color=GREEN,lw=2.2,label="PAT growth YoY")
ax.axhline(0,color=GREY,lw=1)
for i,v in enumerate(nsv_yoy): ax.text(i,v+2.5,f"{v}%",ha="center",fontsize=8,color=BLUE)
for i,v in enumerate(pat_yoy): ax.text(i,v+ (3 if v>=0 else -5),f"{v}%",ha="center",fontsize=8,color=GREEN)
ax.set_xticks(x); ax.set_xticklabels(q); ax.set_ylabel("YoY growth (%)"); ax.set_ylim(-22,46)
ax.legend(fontsize=8,frameon=False,loc="upper right")
ax.set_title("Exhibit: FY26 quarterly growth (NSV vs PAT)",fontsize=10,fontweight="bold",color=NAVY,loc="left")
save(fig,"q_growth.png")

print("ALL DIAGRAMS DONE")
