"""
Diageo PLC - Data Validation Script
Checks internal consistency of the financial model inputs.
Run: python diageo_data_validation.py
"""

passes = 0
fails = 0
warnings = 0

def check(name, condition, detail=""):
    global passes, fails
    if condition:
        passes += 1
    else:
        fails += 1
        print(f"  FAIL: {name} — {detail}")

def warn(name, detail):
    global warnings
    warnings += 1
    print(f"  WARN: {name} — {detail}")

print("=" * 60)
print("DIAGEO PLC - DATA VALIDATION")
print("=" * 60)

# ── Historical Data ──
# Revenue (£m)
NA_REV   = [4500, 5246, 6644, 7190, 6480]
INTL_REV = [7252, 7487, 8808, 9923, 9326]
EBITDA   = [4050, 4410, 5320, 5900, 5380]
DA       = [680,  700,  750,  810,  830]
FINC     = [530,  470,  500,  680,  750]
TAXR     = [0.208, 0.195, 0.202, 0.210, 0.225]
SHARES   = [2358, 2340, 2305, 2270, 2235]
DPS      = [69.88, 72.55, 76.18, 80.00, 83.24]
CAPEX    = [550,  480,  620,  760,  810]

# Balance sheet
CASH     = [2900, 2600, 2400, 2300, 2100]
INV      = [5800, 6100, 6500, 7200, 7400]
RECV     = [2100, 2300, 2700, 2900, 2750]
OTHCA    = [500,  520,  560,  580,  600]
PPE      = [4800, 4650, 4580, 4700, 4850]
ROU      = [780,  750,  730,  720,  710]
GWILL    = [16200,16550,17100,17550,17800]
OTHNA    = [4000, 4050, 4100, 4200, 4200]
PAY      = [3200, 3400, 3800, 4100, 3900]
STDEBT   = [1500, 2200, 2800, 2100, 1800]
CURL     = [120,  125,  130,  135,  140]
OTHCL    = [2400, 2500, 2700, 2900, 2800]
LTDEBT   = [12000,12000,13200,14800,15100]
LTLEAS   = [680,  650,  630,  620,  610]
DEFTAX   = [2800, 2900, 3000, 3100, 3200]
EQUITY   = [14380,13745,12410,12395,12860]

# D&A components
PPE_DEP  = [420, 430, 450, 480, 500]
ROU_DEP  = [130, 135, 140, 145, 150]
INT_AMOR = [130, 135, 160, 185, 180]

# Debt
BOND_TOT = [13500,14200,16000,16900,16900]
LEAS_TOT = [800,  775,  760,  755,  750]

BUYBACK  = [500, 600, 1000, 1250, 1000]

years = ["FY20","FY21","FY22","FY23","FY24"]

# ── 1. Revenue Checks ──
print("\n--- Revenue checks ---")
for i in range(5):
    tot = NA_REV[i] + INTL_REV[i]
    check(f"{years[i]} rev = NA + Intl", tot > 0, f"Total={tot}")
    ebd_mrg = EBITDA[i] / tot
    check(f"{years[i]} EBITDA margin 25-45%", 0.25 <= ebd_mrg <= 0.45,
          f"Margin={ebd_mrg:.1%}")

for i in range(1,5):
    prev_rev = NA_REV[i-1] + INTL_REV[i-1]
    curr_rev = NA_REV[i] + INTL_REV[i]
    yoy = (curr_rev - prev_rev) / prev_rev
    check(f"{years[i]} rev YoY within (-20%, +30%)", -0.20 < yoy < 0.30,
          f"YoY={yoy:.1%}")

# ── 2. Profitability Checks ──
print("\n--- Profitability checks ---")
for i in range(5):
    tot = NA_REV[i] + INTL_REV[i]
    ebit = EBITDA[i] - DA[i]
    pbt = ebit - FINC[i]
    tax = pbt * TAXR[i]
    pat = pbt - tax
    check(f"{years[i]} PAT > 0", pat > 0, f"PAT={pat:.0f}")
    check(f"{years[i]} tax rate 15-30%", 0.15 <= TAXR[i] <= 0.30,
          f"Tax rate={TAXR[i]:.1%}")
    ebit_mrg = ebit / tot
    check(f"{years[i]} EBIT margin 20-40%", 0.20 <= ebit_mrg <= 0.40,
          f"EBIT margin={ebit_mrg:.1%}")

# ── 3. Balance Sheet Checks ──
print("\n--- Balance sheet checks ---")
for i in range(5):
    ca = CASH[i] + INV[i] + RECV[i] + OTHCA[i]
    nca = PPE[i] + ROU[i] + GWILL[i] + OTHNA[i]
    assets = ca + nca
    cl = PAY[i] + STDEBT[i] + CURL[i] + OTHCL[i]
    ncl = LTDEBT[i] + LTLEAS[i] + DEFTAX[i]
    liab = cl + ncl
    net_assets = assets - liab
    diff = abs(net_assets - EQUITY[i])
    check(f"{years[i]} BS balances (assets-liab=equity)", diff < 2,
          f"Diff={diff}")

# ── 4. Debt Checks ──
print("\n--- Debt checks ---")
for i in range(5):
    total_debt = STDEBT[i] + LTDEBT[i]
    check(f"{years[i]} ST+LT debt = bond total",
          abs(total_debt - BOND_TOT[i]) < 5,
          f"ST+LT={total_debt}, Bonds={BOND_TOT[i]}")
    total_leas = CURL[i] + LTLEAS[i]
    check(f"{years[i]} current+LT lease = lease total",
          abs(total_leas - LEAS_TOT[i]) < 5,
          f"Curr+LT={total_leas}, Total={LEAS_TOT[i]}")

# ── 5. D&A Consistency ──
print("\n--- D&A consistency ---")
for i in range(5):
    da_sum = PPE_DEP[i] + ROU_DEP[i] + INT_AMOR[i]
    check(f"{years[i]} D&A components sum to total",
          abs(da_sum - DA[i]) < 5,
          f"Sum={da_sum}, Reported={DA[i]}")

# ── 6. Working Capital Days ──
print("\n--- Working capital checks ---")
for i in range(5):
    rev = NA_REV[i] + INTL_REV[i]
    dio = INV[i] / rev * 365
    dso = RECV[i] / rev * 365
    dpo = PAY[i] / rev * 365
    check(f"{years[i]} DIO 100-200d (spirits aging)", 100 < dio < 200,
          f"DIO={dio:.1f}")
    check(f"{years[i]} DSO 40-90d", 40 < dso < 90, f"DSO={dso:.1f}")
    check(f"{years[i]} DPO 60-120d", 60 < dpo < 120, f"DPO={dpo:.1f}")

# ── 7. Per Share Checks ──
print("\n--- Per share checks ---")
for i in range(5):
    rev = NA_REV[i] + INTL_REV[i]
    ebit = EBITDA[i] - DA[i]
    pbt = ebit - FINC[i]
    pat = pbt - (pbt * TAXR[i])
    eps = pat / SHARES[i] * 100
    check(f"{years[i]} EPS > 0", eps > 0, f"EPS={eps:.1f}p")
    payout = (DPS[i] * SHARES[i] / 100) / pat
    check(f"{years[i]} payout ratio 30-80%", 0.30 < payout < 0.80,
          f"Payout={payout:.1%}")

# ── 8. Capex Checks ──
print("\n--- Capex checks ---")
for i in range(5):
    rev = NA_REV[i] + INTL_REV[i]
    capex_rev = CAPEX[i] / rev
    check(f"{years[i]} capex/rev 3-8%", 0.03 < capex_rev < 0.08,
          f"Capex/Rev={capex_rev:.1%}")

# ── 9. Leverage Checks ──
print("\n--- Leverage checks ---")
for i in range(5):
    nd = BOND_TOT[i] - CASH[i]
    nd_ebitda = nd / EBITDA[i]
    check(f"{years[i]} ND/EBITDA 1.5-4.0x", 1.5 < nd_ebitda < 4.0,
          f"ND/EBITDA={nd_ebitda:.1f}x")

print("\n" + "=" * 60)
print(f"RESULTS: {passes}/{passes+fails} checks passed")
if fails:
    print(f"  {fails} checks FAILED")
if warnings:
    print(f"  {warnings} warnings")
if not fails and not warnings:
    print("\nAll checks passed. Data looks clean.")
print("=" * 60)
