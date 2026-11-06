#!/usr/bin/env python3
"""
Year-End Price-to-NAV (P/NAV) Analysis Engine for First REIT (SGX: AW9U)
========================================================================
Uses yfinance to extract year-end market closing prices, matches them with
audited reported Net Asset Value (NAV per unit) from SGX annual filings,
computes historical P/NAV multiples, premiums/discounts, compares against
benchmark Parkway Life REIT (SGX: C2PU), exports CSVs, and plots trends.
"""

import os
import sys
import json
import csv
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf

# Official Audited Reported NAV per unit (attributable to unitholders, in SGD)
# Sourced from First REIT and Parkway Life REIT Annual Reports and Circulars
OFFICIAL_NAV_DATA = {
    "AW9U.SI": {
        2017: 1.0540,
        2018: 1.0180,
        2019: 0.9990,
        2020: 0.3665,  # Post-Dec 2020 rights issue (pre-rights was 0.4994)
        2021: 0.3665,
        2022: 0.3070,
        2023: 0.2720,
        2024: 0.2470,
        2025: 0.2343,
        2026: 0.2094   # 3Q26 reported (pro forma post-sale: 0.2136)
    },
    "C2PU.SI": {
        2017: 1.7600,
        2018: 1.8300,
        2019: 1.9100,
        2020: 1.9700,
        2021: 2.3500,
        2022: 2.3400,
        2023: 2.3400,
        2024: 2.4300,
        2025: 2.5700,
        2026: 2.5700   # 1H26 reported
    }
}

def get_year_end_price_to_nav(
    ticker="AW9U.SI",
    benchmark_ticker="C2PU.SI",
    start_year=2018,
    end_year=2026,
    output_dir=None,
    save_plot=True
):
    """
    Fetches year-end prices using yfinance and computes P/NAV multiples and discounts.
    """
    if output_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "data_exports")
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"Fetching historical year-end prices from Yahoo Finance for {ticker} and {benchmark_ticker}...")
    
    # Download unadjusted historical prices
    tickers = [ticker, benchmark_ticker]
    data = yf.download(tickers, start=f"{start_year}-01-01", end=f"{end_year+1}-01-01", auto_adjust=False, progress=False)
    
    close_prices = data["Close"]
    adj_close_prices = data["Adj Close"] if "Adj Close" in data else close_prices
    
    results = []
    
    for yr in range(start_year, end_year + 1):
        # Filter for year
        yr_data = close_prices[close_prices.index.year == yr]
        if yr_data.empty:
            continue
            
        last_date = yr_data.index[-1]
        date_str = last_date.strftime("%Y-%m-%d")
        
        # Target REIT
        price_reit = float(yr_data[ticker].dropna().iloc[-1])
        adj_price_reit = float(adj_close_prices[adj_close_prices.index.year == yr][ticker].dropna().iloc[-1])
        nav_reit = OFFICIAL_NAV_DATA.get(ticker, {}).get(yr, np.nan)
        
        p_to_nav_reit = price_reit / nav_reit if (nav_reit and not np.isnan(nav_reit) and nav_reit > 0) else np.nan
        prem_disc_reit = ((price_reit - nav_reit) / nav_reit * 100) if p_to_nav_reit else np.nan
        
        # Benchmark Peer
        price_bench = float(yr_data[benchmark_ticker].dropna().iloc[-1])
        nav_bench = OFFICIAL_NAV_DATA.get(benchmark_ticker, {}).get(yr, np.nan)
        p_to_nav_bench = price_bench / nav_bench if (nav_bench and not np.isnan(nav_bench) and nav_bench > 0) else np.nan
        prem_disc_bench = ((price_bench - nav_bench) / nav_bench * 100) if p_to_nav_bench else np.nan
        
        results.append({
            "Year": yr,
            "Date": date_str,
            f"Price_{ticker}": price_reit,
            f"AdjPrice_{ticker}": adj_price_reit,
            f"NAV_{ticker}": nav_reit,
            f"P_NAV_{ticker}": p_to_nav_reit,
            f"Prem_Disc_{ticker}_Pct": prem_disc_reit,
            f"Price_{benchmark_ticker}": price_bench,
            f"NAV_{benchmark_ticker}": nav_bench,
            f"P_NAV_{benchmark_ticker}": p_to_nav_bench,
            f"Prem_Disc_{benchmark_ticker}_Pct": prem_disc_bench
        })
        
    df = pd.DataFrame(results)
    
    # Save CSV
    csv_path = os.path.join(output_dir, "first_reit_year_end_p_to_nav.csv")
    df.to_csv(csv_path, index=False)
    print(f"Saved Year-End P/NAV CSV to: {csv_path}")
    
    # Plot Trend
    plot_path = os.path.join(output_dir, "year_end_p_to_nav_trend.png")
    if save_plot:
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), dpi=200, sharex=True)
        plt.subplots_adjust(hspace=0.25)
        
        years = df["Year"]
        
        # Panel 1: Price vs NAV per unit for First REIT
        ax1.plot(years, df[f"Price_{ticker}"], marker="o", color="#2563eb", linewidth=2.5, label=f"Year-End Market Price ({ticker})")
        ax1.plot(years, df[f"NAV_{ticker}"], marker="s", color="#dc2626", linewidth=2.5, linestyle="--", label=f"Reported NAV per Unit ({ticker})")
        
        # Annotate values
        for i, row in df.iterrows():
            ax1.annotate(f"S${row[f'Price_{ticker}']:.3f}", (row["Year"], row[f"Price_{ticker}"]),
                         textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8, fontweight="bold", color="#1e40af")
            ax1.annotate(f"NAV {row[f'NAV_{ticker}']:.3f}", (row["Year"], row[f"NAV_{ticker}"]),
                         textcoords="offset points", xytext=(0, -14), ha="center", fontsize=8, color="#991b1b")
                         
        ax1.set_title(f"First REIT ({ticker}): Historical Year-End Market Price vs. Net Asset Value (NAV)", fontsize=13, fontweight="bold")
        ax1.set_ylabel("Singapore Dollars (S$)", fontsize=10)
        ax1.grid(True, linestyle="--", alpha=0.5)
        ax1.legend(loc="upper right", fontsize=10)
        
        # Panel 2: P/NAV Multiple Comparison: First REIT vs Parkway Life REIT
        ax2.plot(years, df[f"P_NAV_{ticker}"], marker="o", color="#2563eb", linewidth=2.8, label=f"First REIT P/NAV ({ticker})")
        ax2.plot(years, df[f"P_NAV_{benchmark_ticker}"], marker="^", color="#059669", linewidth=2.8, label=f"Parkway Life REIT P/NAV ({benchmark_ticker})")
        
        # Parity Line (1.0x NAV)
        ax2.axhline(y=1.0, color="#64748b", linestyle=":", linewidth=1.5, label="NAV Parity (1.0x Multiple)")
        ax2.fill_between(years, 0, 1.0, color="#f43f5e", alpha=0.08, label="Discount to NAV Zone")
        ax2.fill_between(years, 1.0, 2.5, color="#10b981", alpha=0.08, label="Premium to NAV Zone")
        
        for i, row in df.iterrows():
            # First REIT P/NAV
            lbl_reit = f"{row[f'P_NAV_{ticker}']:.2f}x\n({row[f'Prem_Disc_{ticker}_Pct']:+.1f}%)"
            ax2.annotate(lbl_reit, (row["Year"], row[f"P_NAV_{ticker}"]),
                         textcoords="offset points", xytext=(0, -22), ha="center", fontsize=8, fontweight="bold", color="#1e40af")
            # Parkway P/NAV
            lbl_bench = f"{row[f'P_NAV_{benchmark_ticker}']:.2f}x"
            ax2.annotate(lbl_bench, (row["Year"], row[f"P_NAV_{benchmark_ticker}"]),
                         textcoords="offset points", xytext=(0, 10), ha="center", fontsize=8, fontweight="bold", color="#065f46")
                         
        ax2.set_title("Historical Year-End Price-to-NAV (P/NAV) Multiple: First REIT vs. Parkway Life REIT", fontsize=13, fontweight="bold")
        ax2.set_xlabel("Year (as of Last Trading Day)", fontsize=10)
        ax2.set_ylabel("P/NAV Ratio (Multiple)", fontsize=10)
        ax2.set_ylim(0.3, 2.4)
        ax2.grid(True, linestyle="--", alpha=0.5)
        ax2.legend(loc="upper left", fontsize=9)
        
        fig.tight_layout()
        fig.savefig(plot_path, dpi=200, bbox_inches="tight")
        print(f"Saved Year-End P/NAV plot to: {plot_path}")
        plt.close(fig)
        
    return df, csv_path, plot_path

if __name__ == "__main__":
    df, csv_p, plot_p = get_year_end_price_to_nav()
    print("\n=== YEAR-END PRICE TO NAV SUMMARY TABLE ===")
    print(df[["Year", "Date", "Price_AW9U.SI", "NAV_AW9U.SI", "P_NAV_AW9U.SI", "Prem_Disc_AW9U.SI_Pct", "P_NAV_C2PU.SI"]].to_string(index=False))
