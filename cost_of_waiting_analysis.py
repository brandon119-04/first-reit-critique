#!/usr/bin/env python3
"""
Cost of Delayed Reinvestment Analysis: First REIT (SGX: AW9U)
Implements:
1. Model A: DPU Recovery Deadline & Net Present Value of Delay NPV(T)
2. Model B: P/NAV Recovery, Cost of Equity & Required DPU Sensitivity Matrix
3. Model C: Financing Deadline & Equity Dilution Constraint Model
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

OUTPUT_DIR = "/Users/brandonwong/.gemini/antigravity/scratch/first-reit-critique/data_exports"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# Base Parameters (First REIT 9M 2026 / Post-Tranche 1 Filings)
# -------------------------------------------------------------
UNITS = 2076.62  # Millions of units
NAVPU_PRE = 0.2094  # S$0.2094 (20.94 cents, 9M 2026 actual)
MARKET_PRICE = 0.170  # S$0.170 current trading price
CURRENT_P_NAV = MARKET_PRICE / NAVPU_PRE  # ~0.812x
GEARING_CURRENT = 0.156  # 15.6% post-T1 debt retirement (or 0.275 9M26)
TOTAL_ASSETS = 605.0  # S$ Millions
TOTAL_DEBT = 94.4  # S$ Millions (Post S$362.7M repayment)
CASH_UNALLOCATED = 101.5  # S$ Millions

# Income surrender & interest savings
LOST_NPI_ANNUAL = 35.26  # S$M/yr lost from 11 divested assets
SAVED_INTEREST_ANNUAL = 18.80  # S$M/yr saved on S$362.7M debt @ 5.18%
CASH_YIELD_RATE = 0.028  # 2.80% on MAS T-bills
CASH_INCOME_ANNUAL = CASH_UNALLOCATED * CASH_YIELD_RATE  # S$2.842M/yr
NET_INCOME_DRAG_ANNUAL = LOST_NPI_ANNUAL - SAVED_INTEREST_ANNUAL - CASH_INCOME_ANNUAL  # S$13.618M/yr
DRAG_PER_UNIT_ANNUAL = NET_INCOME_DRAG_ANNUAL / UNITS  # S$0.00656/unit/yr (0.656 cents/yr)

# Cost of equity and growth
KE = 0.095  # 9.5% baseline cost of equity
G = 0.010   # 1.0% perpetual growth rate
DISCOUNT_RATE = KE  # 9.5%

# =============================================================
# 1. MODEL A: DPU RECOVERY DEADLINE & DELAY ECONOMICS NPV(T)
# =============================================================
# We evaluate reinvestment after T years of delay (T = 0, 1, 2, 3, 4, 5).
# Reinvestment Scenarios:
# Scenario 1 (Dilutive - German Senior Living): S$250M @ 4.5% yield, 50% debt @ 4.5%, RETT 6.5%
# Scenario 2 (Neutral - Developed Market Core): S$250M @ 5.5% yield, 50% debt @ 4.0%, fees 3.0%
# Scenario 3 (Accretive - High Yield Healthcare): S$250M @ 6.5% yield, 50% debt @ 4.0%, fees 3.0%

delays = [0, 1, 2, 3, 4, 5]
records_model_a = []

for T in delays:
    # Cumulative PV of income lost during waiting period [1..T]
    if T == 0:
        pv_lost_income = 0.0
    else:
        pv_lost_income = sum(NET_INCOME_DRAG_ANNUAL / ((1 + DISCOUNT_RATE) ** t) for t in range(1, T + 1))
    
    # Benefit B_T at year T for 3 asset acquisition scenarios:
    # Asset size S$250M, 50% debt funded (S$125M debt, S$125M equity from cash)
    scenarios = {
        "German_Senior_Living_4.5pct": {"yield": 0.045, "debt_cost": 0.045, "frictions": 0.075},
        "Developed_Core_5.5pct":       {"yield": 0.055, "debt_cost": 0.040, "frictions": 0.035},
        "Accretive_Target_6.5pct":     {"yield": 0.065, "debt_cost": 0.040, "frictions": 0.035}
    }
    
    row = {
        "Delay_Years_T": T,
        "Cumulative_PV_Lost_Income_SGD_M": round(pv_lost_income, 3),
        "Cumulative_Lost_DPU_Cents": round((pv_lost_income / UNITS) * 100, 3)
    }
    
    for s_name, s_params in scenarios.items():
        acq_npi = 250.0 * s_params["yield"]
        acq_int = 125.0 * s_params["debt_cost"]
        inc_dist = acq_npi - acq_int  # S$M incremental cash flow to equity
        # Capitalized value of perpetual cash flow
        cap_val = inc_dist / (KE - G)
        # Cost of equity deployed + non-recoverable acquisition fees/taxes
        net_equity_cost = 125.0 + (250.0 * s_params["frictions"])
        # B_T: Net economic value created at acquisition year T
        B_T = cap_val - net_equity_cost
        
        # Discount B_T back to t=0
        pv_benefit = B_T / ((1 + DISCOUNT_RATE) ** T)
        npv_T = -pv_lost_income + pv_benefit
        npv_per_unit_cents = (npv_T / UNITS) * 100
        
        row[f"Benefit_{s_name}_SGD_M"] = round(B_T, 2)
        row[f"NPV_{s_name}_SGD_M"] = round(npv_T, 2)
        row[f"NPV_Per_Unit_{s_name}_Cents"] = round(npv_per_unit_cents, 3)
    
    records_model_a.append(row)

df_model_a = pd.DataFrame(records_model_a)
df_model_a.to_csv(os.path.join(OUTPUT_DIR, "cost_of_delayed_reinvestment_model.csv"), index=False)
print("Saved Model A to cost_of_delayed_reinvestment_model.csv")


# =============================================================
# 2. MODEL B: P/NAV RECOVERY & REQUIRED DPU SENSITIVITY MATRIX
# =============================================================
# Formula: DPU_required = (P/NAV)_target * NAVPU * (k_e - g)
nav_benchmarks = [0.2094, 0.2100, 0.2200]
target_p_nav_list = [0.60, 0.70, 0.80, 0.85, 0.90, 0.95, 1.00, 1.05, 1.10]
ke_list = [0.080, 0.085, 0.090, 0.095, 0.100, 0.105, 0.110]
g_default = 0.010

records_model_b = []
for p_nav in target_p_nav_list:
    for ke_val in ke_list:
        spread = ke_val - g_default
        req_dpu = p_nav * NAVPU_PRE * spread
        implied_price = req_dpu / spread
        records_model_b.append({
            "Target_P_NAV": p_nav,
            "Cost_of_Equity_ke": ke_val,
            "Growth_Rate_g": g_default,
            "Equity_Spread_ke_minus_g": round(spread, 4),
            "NAV_Per_Unit_SGD": NAVPU_PRE,
            "Implied_Unit_Price_SGD": round(implied_price, 4),
            "Required_Forward_DPU_SGD": round(req_dpu, 5),
            "Required_Forward_DPU_Cents": round(req_dpu * 100, 3)
        })

df_model_b = pd.DataFrame(records_model_b)
df_model_b.to_csv(os.path.join(OUTPUT_DIR, "pnav_dpu_sensitivity_matrix.csv"), index=False)
print("Saved Model B to pnav_dpu_sensitivity_matrix.csv")


# =============================================================
# 3. MODEL C: FINANCING DEADLINE & EQUITY DILUTION
# =============================================================
# NAVPU_post = (NAV_pre + N_new * P_issue - C) / (N_old + N_new)
# We test equity raising amounts from S$20M to S$150M at 3 issue pricing scenarios:
# 1. At current market price: S$0.170
# 2. At 5% placement discount: S$0.1615
# 3. At 10% placement discount: S$0.1530
issue_sizes_m = [20.0, 40.0, 60.0, 80.0, 100.0, 125.0, 150.0]
pricing_scenarios = {
    "At_Market_0.170": 0.170,
    "Discount_5pct_0.1615": 0.1615,
    "Discount_10pct_0.1530": 0.1530
}
issue_cost_pct = 0.025  # 2.5% underwriting and advisory fees
nav_pre_total = UNITS * NAVPU_PRE  # S$434.84M

records_model_c = []
for size_m in issue_sizes_m:
    for p_name, issue_p in pricing_scenarios.items():
        n_new = size_m / issue_p  # Millions of new units
        costs = size_m * issue_cost_pct
        nav_post_total = nav_pre_total + size_m - costs
        units_post = UNITS + n_new
        navpu_post = nav_post_total / units_post
        nav_dilution_cents = (NAVPU_PRE - navpu_post) * 100
        nav_dilution_pct = ((navpu_post - NAVPU_PRE) / NAVPU_PRE) * 100
        
        records_model_c.append({
            "Equity_Raise_SGD_M": size_m,
            "Issue_Pricing_Scenario": p_name,
            "Issue_Price_SGD": issue_p,
            "New_Units_Issued_M": round(n_new, 2),
            "Pro_Forma_Units_M": round(units_post, 2),
            "Issue_Costs_SGD_M": round(costs, 3),
            "NAVPU_Pre_Cents": round(NAVPU_PRE * 100, 2),
            "NAVPU_Post_Cents": round(navpu_post * 100, 3),
            "NAV_Dilution_Cents": round(nav_dilution_cents, 3),
            "NAV_Dilution_Percent": round(nav_dilution_pct, 2),
            "Breaches_2pct_Dilution_Ceiling": abs(nav_dilution_pct) > 2.0,
            "Breaches_5pct_Dilution_Ceiling": abs(nav_dilution_pct) > 5.0
        })

df_model_c = pd.DataFrame(records_model_c)
df_model_c.to_csv(os.path.join(OUTPUT_DIR, "equity_dilution_financing_capacity.csv"), index=False)
print("Saved Model C to equity_dilution_financing_capacity.csv")


# =============================================================
# 4. INSTITUTIONAL PLOT: THE COST OF WAITING TRADE-OFF
# =============================================================
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6), dpi=300)

# Subplot 1: Cumulative Income Drag vs Delay T
years = df_model_a["Delay_Years_T"]
lost_inc = df_model_a["Cumulative_PV_Lost_Income_SGD_M"]
lost_dpu = df_model_a["Cumulative_Lost_DPU_Cents"]

ax1.plot(years, lost_inc, 'o-', color='#e11d48', linewidth=2.5, markersize=7, label="PV Cumulative Lost Income (S$M)")
ax1.set_title("Model A: Cumulative Income Drag from Waiting\n(Net of Interest Savings & Cash Yield)", fontsize=12, fontweight='bold', pad=12)
ax1.set_xlabel("Reinvestment Delay T (Years)", fontsize=10, fontweight='semibold')
ax1.set_ylabel("PV of Income Drag (S$ Millions)", fontsize=10, fontweight='semibold', color='#e11d48')
ax1.tick_params(axis='y', labelcolor='#e11d48')
ax1.grid(True, linestyle='--', alpha=0.5)

ax1_twin = ax1.twinx()
ax1_twin.plot(years, lost_dpu, 's--', color='#9333ea', linewidth=2.0, markersize=6, label="Cumulative DPU Loss (¢/unit)")
ax1_twin.set_ylabel("Cumulative DPU Drag (Cents/Unit)", fontsize=10, fontweight='semibold', color='#9333ea')
ax1_twin.tick_params(axis='y', labelcolor='#9333ea')
ax1_twin.grid(False)

# Subplot 2: Net Strategy NPV(T) across 3 Reinvestment Quality Curves
npv_german = df_model_a["NPV_Per_Unit_German_Senior_Living_4.5pct_Cents"]
npv_core = df_model_a["NPV_Per_Unit_Developed_Core_5.5pct_Cents"]
npv_accretive = df_model_a["NPV_Per_Unit_Accretive_Target_6.5pct_Cents"]

ax2.plot(years, npv_accretive, 'd-', color='#10b981', linewidth=2.5, markersize=7, label="Accretive Acquisition (6.5% Yield)")
ax2.plot(years, npv_core, 'o-', color='#3b82f6', linewidth=2.5, markersize=7, label="Neutral Core Acquisition (5.5% Yield)")
ax2.plot(years, npv_german, 'x-', color='#ef4444', linewidth=2.5, markersize=7, label="German Senior Living Option (4.5% Yield)")

# Benchmark Hold Alternative (Scenario A NPV = 13.61 cents) and Capital Return (18.65 cents)
ax2.axhline(0, color='gray', linestyle=':', linewidth=1.5, label="Zero Net NPV Line")
ax2.set_title("The Reinvestment Delay Trade-Off:\nNet Strategy NPV(T) vs. Waiting Time", fontsize=12, fontweight='bold', pad=12)
ax2.set_xlabel("Reinvestment Delay T (Years)", fontsize=10, fontweight='semibold')
ax2.set_ylabel("Net Strategy NPV per Unit (Singapore Cents)", fontsize=10, fontweight='semibold')
ax2.legend(loc='lower left', frameon=True, facecolor='white', framealpha=0.9, fontsize=9)
ax2.grid(True, linestyle='--', alpha=0.5)

plt.suptitle("First REIT (SGX: AW9U) — How Long Can First REIT Afford to Wait?\nThe Cost of Delayed Reinvestment & The Waiting Trade-Off", fontsize=14, fontweight='bold', y=1.02)
plt.tight_layout()

chart_path = os.path.join(OUTPUT_DIR, "cost_of_waiting_tradeoff_plot.png")
chart_artifact_path = "/Users/brandonwong/.gemini/antigravity/brain/4b69b333-e895-416e-b499-9f6ffa7fb747/cost_of_waiting_tradeoff_plot.png"
plt.savefig(chart_path)
plt.savefig(chart_artifact_path)
plt.close()
print(f"Saved institutional plot to {chart_path} and {chart_artifact_path}")
