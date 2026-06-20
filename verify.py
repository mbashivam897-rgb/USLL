import openpyxl, warnings
warnings.filterwarnings("ignore")

OUT = "/projects/sandbox/USLL/United Spirits Excell - 8 Year Forecast.xlsx"
VER = "/projects/sandbox/USLL/_verify.xlsx"

wb = openpyxl.load_workbook(OUT, data_only=False)
dcf = wb["DCF"]
for r in range(18, 28):
    for c in range(11, 21):  # K..T : strip unsupported what-if data table
        dcf.cell(row=r, column=c).value = None
wb.save(VER)

import formulas
xl = formulas.ExcelModel().loads(VER).finish()
sol = xl.calculate()

def g(sheet, cell):
    suffix = ("]" + sheet + "'!" + cell).upper()
    for k, v in sol.items():
        if k.upper().endswith(suffix):
            try:
                return v.value[0, 0]
            except Exception:
                return v.value
    return float("nan")

cols8 = ["M","N","O","P","Q","R","S","T"]
dcols  = ["C","D","E","F","G","H","I","J"]
yrs = list(range(2027, 2035))

print("=== Balance Sheet check (Assets - L&E), should be ~0 ===")
for col, yr in zip(cols8, yrs):
    t = g("BS", col+"51"); a = g("BS", col+"48")
    print(f"  {yr}: test={t:.4f}   Assets={a:,.0f}")

print("\n=== Key forecast outputs ===")
for col, yr in zip(cols8, yrs):
    s = g("PL", col+"6"); e = g("PL", col+"15"); p = g("PL", col+"23"); c = g("BS", col+"45")
    print(f"  {yr}: Sales={s:,.0f}  EBIT={e:,.0f}  PAT={p:,.0f}  Cash={c:,.0f}")

print("\n=== DCF (8-year explicit) ===")
for col, yr in zip(dcols, yrs):
    f = g("DCF", col+"12"); pv = g("DCF", col+"15")
    print(f"  {yr}: FCF={f:,.0f}  PV={pv:,.0f}")
for lbl, cell in [("WACC (N9)","N9"),("Terminal FCF","C19"),("Terminal Value","C20"),
                  ("PV of TV","C21"),("Sum PV FCF","C22"),("Enterprise Value","C23"),
                  ("Equity Value","C26"),("PER SHARE","C28")]:
    v = g("DCF", cell)
    print(f"  {lbl:<16}: {v:,.4f}" if cell=="N9" else f"  {lbl:<16}: {v:,.2f}")

print("\n=== ROU / PPE / Intangible net asset sanity over horizon ===")
for col, yr in zip(cols8, yrs):
    r = g("BS", col+"34"); pp = g("BS", col+"32"); it = g("BS", col+"35")
    print(f"  {yr}: ROU={r:,.0f}  PPE={pp:,.0f}  Intang={it:,.0f}")
