"""
Tesco PLC - Data validation and cross-check script.

Verifies that the historical inputs used in the Excel model are internally
consistent and fall within reasonable ranges. Run this before opening the
model to confirm nothing was mistyped or mis-sourced.

Source: Tesco PLC Annual Reports FY2020-FY2024 (Feb year-end).
"""

import sys

# ── Historical data (£m unless stated) ────────────────────────────────────
# All figures from Tesco Annual Reports, pages cited in the Excel model.

YEARS = ["FY2020", "FY2021", "FY2022", "FY2023", "FY2024"]

# Income statement
uk_rev   = [52333, 48355, 49195, 59381, 61674]
ce_rev   = [5758,  5663,  5637,  6381,  6519]
ebitda   = [5359,  5748,  6034,  6266,  6509]
da       = [3390,  3384,  3497,  3636,  3746]
finc     = [700,   680,   720,   730,   710]
tax_rate = [0.200, 0.190, 0.210, 0.230, 0.230]
shares   = [8000,  7800,  7600,  7400,  7200]
dps      = [5.5,   7.7,   10.9,  11.78, 12.10]

# Balance sheet
cash     = [3218, 2506, 2688, 2639, 2802]
inv      = [2334, 2310, 2485, 2531, 2506]
recv     = [916,  881,  998,  1029, 1052]
oth_ca   = [742,  671,  724,  798,  812]
ppe      = [12811,12208,11748,11612,11542]
rou      = [12209,11652,11293,11124,10987]
gwill    = [4153, 4113, 4088, 4069, 4057]
oth_na   = [1982, 2041, 2087, 2134, 2178]
pay      = [6310, 6472, 6788, 6903, 6981]
st_debt  = [1142,  672, 1012,  842,  524]
cur_leas = [899,   908,  921,  935,  952]
oth_cl   = [1478, 1389, 1502, 1631, 1662]
lt_debt  = [5695, 5186, 4612, 4421, 4182]
lt_leas  = [11428,10842,10462,10293,10102]
def_tax  = [1893, 1788, 1842, 1917, 1989]

# Capex and depreciation detail
capex    = [1322, 1201, 1238, 1409, 1514]
ppe_dep  = [1891, 1845, 1803, 1784, 1826]
rou_dep  = [1455, 1495, 1651, 1810, 1878]
int_amor = [44,   44,   43,   42,   42]

# Debt detail
bond_tot = [6837, 5858, 5624, 5263, 4706]
buyback  = [0,    400,  750,  750,  1000]


# ── Validation checks ────────────────────────────────────────────────────

def run_checks():
    errors = []
    warnings = []
    checks_passed = 0
    checks_total = 0

    def check(name, condition, detail=""):
        nonlocal checks_passed, checks_total
        checks_total += 1
        if condition:
            checks_passed += 1
        else:
            errors.append(f"FAIL: {name}" + (f" ({detail})" if detail else ""))

    def warn(name, condition, detail=""):
        if not condition:
            warnings.append(f"WARN: {name}" + (f" ({detail})" if detail else ""))

    print("=" * 60)
    print("TESCO PLC - DATA VALIDATION")
    print("=" * 60)

    # 1. Revenue should be positive and growing (roughly)
    print("\n--- Revenue checks ---")
    for i in range(5):
        total = uk_rev[i] + ce_rev[i]
        check(f"{YEARS[i]} revenue > 0", total > 0, f"total={total}")
        check(f"{YEARS[i]} UK > CE", uk_rev[i] > ce_rev[i],
              f"UK={uk_rev[i]}, CE={ce_rev[i]}")

    # Revenue growth should be between -20% and +30% (sanity)
    for i in range(1, 5):
        prev = uk_rev[i-1] + ce_rev[i-1]
        curr = uk_rev[i] + ce_rev[i]
        growth = (curr - prev) / prev
        check(f"{YEARS[i]} revenue growth reasonable",
              -0.20 < growth < 0.30,
              f"growth={growth:.1%}")

    # 2. EBITDA margin between 5% and 15% (reasonable for grocery)
    print("\n--- Profitability checks ---")
    for i in range(5):
        total_rev = uk_rev[i] + ce_rev[i]
        margin = ebitda[i] / total_rev
        check(f"{YEARS[i]} EBITDA margin in range",
              0.05 < margin < 0.15,
              f"margin={margin:.1%}")

    # EBIT should be positive
    for i in range(5):
        ebit = ebitda[i] - da[i]
        check(f"{YEARS[i]} EBIT positive", ebit > 0, f"EBIT={ebit}")

    # Tax rate between 15% and 30%
    for i in range(5):
        check(f"{YEARS[i]} tax rate reasonable",
              0.15 <= tax_rate[i] <= 0.30,
              f"rate={tax_rate[i]:.0%}")

    # 3. Balance sheet balances
    print("\n--- Balance sheet checks ---")
    for i in range(5):
        total_assets = (cash[i] + inv[i] + recv[i] + oth_ca[i]
                       + ppe[i] + rou[i] + gwill[i] + oth_na[i])
        total_liab = (pay[i] + st_debt[i] + cur_leas[i] + oth_cl[i]
                     + lt_debt[i] + lt_leas[i] + def_tax[i])
        equity = total_assets - total_liab
        check(f"{YEARS[i]} equity positive", equity > 0,
              f"assets={total_assets}, liab={total_liab}, equity={equity}")
        # Equity should be between 5bn and 15bn for Tesco
        check(f"{YEARS[i]} equity in range",
              5000 < equity < 15000,
              f"equity={equity}")

    # 4. Debt consistency: ST + LT should roughly equal bond_tot
    print("\n--- Debt checks ---")
    for i in range(5):
        model_total = st_debt[i] + lt_debt[i]
        diff = abs(model_total - bond_tot[i])
        check(f"{YEARS[i]} debt ST+LT matches bond total",
              diff < 100,
              f"ST+LT={model_total}, bonds={bond_tot[i]}, diff={diff}")

    # 5. D&A consistency: component sum should match total
    print("\n--- D&A consistency ---")
    for i in range(5):
        component_sum = ppe_dep[i] + rou_dep[i] + int_amor[i]
        diff = abs(component_sum - da[i])
        check(f"{YEARS[i]} D&A components sum to total",
              diff < 50,
              f"components={component_sum}, reported={da[i]}, diff={diff}")

    # 6. Working capital days (sanity check)
    print("\n--- Working capital checks ---")
    for i in range(5):
        total_rev = uk_rev[i] + ce_rev[i]
        dio = inv[i] / total_rev * 365
        dso = recv[i] / total_rev * 365
        dpo = pay[i] / total_rev * 365

        # Grocery: DIO typically 10-20 days, DSO 3-8, DPO 30-50
        check(f"{YEARS[i]} DIO reasonable", 8 < dio < 25, f"DIO={dio:.1f}")
        check(f"{YEARS[i]} DSO reasonable", 2 < dso < 12, f"DSO={dso:.1f}")
        check(f"{YEARS[i]} DPO reasonable", 25 < dpo < 55, f"DPO={dpo:.1f}")

    # 7. Per share checks
    print("\n--- Per share checks ---")
    for i in range(5):
        ebit = ebitda[i] - da[i]
        pbt = ebit - finc[i]
        tax = pbt * tax_rate[i]
        pat = pbt - tax
        eps = pat / shares[i] * 100  # pence
        check(f"{YEARS[i]} EPS positive", eps > 0, f"EPS={eps:.1f}p")
        # DPS should be less than EPS (payout < 100%)
        payout = dps[i] / eps if eps > 0 else 999
        warn(f"{YEARS[i]} payout ratio < 100%", payout < 1.0,
             f"payout={payout:.0%}")

    # 8. Capex as % of revenue
    print("\n--- Capex checks ---")
    for i in range(5):
        total_rev = uk_rev[i] + ce_rev[i]
        capex_pct = capex[i] / total_rev
        check(f"{YEARS[i]} capex/revenue in range",
              0.01 < capex_pct < 0.05,
              f"capex/rev={capex_pct:.1%}")

    # 9. Net debt / EBITDA leverage
    print("\n--- Leverage checks ---")
    for i in range(5):
        nd_ex_leases = bond_tot[i] - cash[i]
        leverage = nd_ex_leases / ebitda[i]
        check(f"{YEARS[i]} ND/EBITDA (ex-leases) reasonable",
              -1.0 < leverage < 3.0,
              f"leverage={leverage:.1f}x")

    # ── Summary ──
    print("\n" + "=" * 60)
    print(f"RESULTS: {checks_passed}/{checks_total} checks passed")
    if errors:
        print(f"\nERRORS ({len(errors)}):")
        for e in errors:
            print(f"  {e}")
    if warnings:
        print(f"\nWARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  {w}")
    if not errors and not warnings:
        print("\nAll checks passed. Data looks clean.")
    print("=" * 60)

    return len(errors) == 0


if __name__ == "__main__":
    ok = run_checks()
    sys.exit(0 if ok else 1)
