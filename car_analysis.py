#!/usr/bin/env python3
"""
First REIT (SGX: AW9U) Event Study & Cumulative Abnormal Return (CAR) Analysis
==============================================================================
Econometric Market Model Regression and Event Study Engine.
Fetches historical prices using yfinance, regresses REIT returns against market/peer benchmarks
(y = REIT, x = Benchmark) over an estimation window, predicts expected returns, calculates
Abnormal Returns (AR) and Cumulative Abnormal Returns (CAR), exports comprehensive CSVs,
and plots institutional-grade graphs.
"""

import os
import sys
import json
import io
import zipfile
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import yfinance as yf

def analyze_cumulative_abnormal_returns(
    reit_ticker="AW9U.SI",
    benchmark_ticker="^STI",
    peer_ticker="C2PU.SI",
    event_date="2026-04-01",
    estimation_days_pre=180,
    estimation_gap_days=21,
    event_window_pre=20,
    event_window_post=20,
    output_dir=None,
    save_plot=True,
    show_plot=False
):
    """
    Performs Market Model OLS regression and Cumulative Abnormal Return (CAR) analysis.
    
    Parameters:
    -----------
    reit_ticker : str
        Target REIT ticker on Yahoo Finance (default: 'AW9U.SI' for First REIT).
    benchmark_ticker : str
        Market benchmark ticker (default: '^STI' for Straits Times Index, or 'CLR.SI' for S-REIT ETF).
    peer_ticker : str
        Peer benchmark ticker for comparative tracking (default: 'C2PU.SI' for Parkway Life REIT).
    event_date : str
        Event announcement date in 'YYYY-MM-DD' format (default: '2026-04-01' for Tranche 1 Divestment).
    estimation_days_pre : int
        Number of trading days for market model estimation window prior to gap (default: 180).
    estimation_gap_days : int
        Buffer trading days between estimation window and event date to avoid pre-event leakage (default: 21).
    event_window_pre : int
        Trading days before event date in event window (default: 20).
    event_window_post : int
        Trading days after event date in event window (default: 20).
    output_dir : str
        Directory to save generated CSVs and plot. Defaults to script directory / 'output'.
    save_plot : bool
        Whether to generate and save the 4-panel analysis plot.
        
    Returns:
    --------
    dict
        Dictionary containing regression statistics, CAR metrics, file paths, and helper functions.
    """
    if output_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "event_study_output")
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"================================================================================")
    print(f"EVENT STUDY & CAR ANALYSIS: {reit_ticker} vs {benchmark_ticker}")
    print(f"Event Date: {event_date} (Announcement of Indonesian Hospital Divestment)")
    print(f"================================================================================")
    
    # 1. Date Range Setup
    event_dt = pd.to_datetime(event_date)
    # Fetch wider window to guarantee sufficient trading days
    start_fetch = (event_dt - pd.Timedelta(days=estimation_days_pre * 2.5)).strftime("%Y-%m-%d")
    end_fetch = (event_dt + pd.Timedelta(days=(event_window_post + 30) * 2)).strftime("%Y-%m-%d")
    
    tickers_to_fetch = list(set([reit_ticker, benchmark_ticker, peer_ticker]))
    print(f"Downloading historical price series from Yahoo Finance: {tickers_to_fetch}...")
    
    raw_data = yf.download(tickers_to_fetch, start=start_fetch, end=end_fetch, progress=False)
    
    # Extract Close or Adj Close prices
    if isinstance(raw_data.columns, pd.MultiIndex):
        if "Adj Close" in raw_data.columns.levels[0]:
            price_df = raw_data["Adj Close"].copy()
        else:
            price_df = raw_data["Close"].copy()
    else:
        price_df = raw_data.copy()
        
    # Forward-fill and drop missing values across target and benchmark
    price_df = price_df.ffill().dropna(subset=[reit_ticker, benchmark_ticker])
    
    # Calculate daily percentage returns
    returns_df = price_df.pct_change().dropna()
    
    # 2. Identify Event Date Index
    # Find closest trading day to event_date
    trading_dates = returns_df.index
    event_idx_candidates = trading_dates[trading_dates >= event_dt]
    if len(event_idx_candidates) == 0:
        raise ValueError(f"Event date {event_date} is beyond available trading dates (max: {trading_dates.max()})")
    
    t0_date = event_idx_candidates[0]
    t0_loc = returns_df.index.get_loc(t0_date)
    print(f"Closest trading day to event (t=0): {t0_date.strftime('%Y-%m-%d')} (index {t0_loc})")
    
    # 3. Define Estimation Window and Event Window
    # Estimation window: [t0 - estimation_days_pre - estimation_gap_days : t0 - estimation_gap_days]
    est_start_loc = max(0, t0_loc - estimation_days_pre - estimation_gap_days)
    est_end_loc = max(0, t0_loc - estimation_gap_days)
    
    if est_end_loc - est_start_loc < 30:
        raise ValueError(f"Insufficient estimation window observations ({est_end_loc - est_start_loc} days).")
        
    est_df = returns_df.iloc[est_start_loc:est_end_loc].copy()
    
    # Event window: [t0 - event_window_pre : t0 + event_window_post + 1]
    evt_start_loc = max(0, t0_loc - event_window_pre)
    evt_end_loc = min(len(returns_df), t0_loc + event_window_post + 1)
    
    evt_df = returns_df.iloc[evt_start_loc:evt_end_loc].copy()
    
    # 4. Market Model OLS Regression: y = REIT return, x = Benchmark return
    y_est = est_df[reit_ticker].values
    x_est = est_df[benchmark_ticker].values
    
    slope, intercept, r_value, p_value, std_err = stats.linregress(x_est, y_est)
    alpha = intercept
    beta = slope
    r_squared = r_value ** 2
    
    # Residual variance and standard error of regression
    y_pred_est = alpha + beta * x_est
    residuals_est = y_est - y_pred_est
    n_est = len(y_est)
    dof_est = n_est - 2
    s_squared_e = np.sum(residuals_est ** 2) / dof_est
    s_e = np.sqrt(s_squared_e)
    
    # t-statistics
    var_x = np.var(x_est, ddof=1)
    mean_x = np.mean(x_est)
    se_beta = s_e / np.sqrt(n_est * var_x)
    se_alpha = s_e * np.sqrt(1.0 / n_est + (mean_x ** 2) / (n_est * var_x))
    t_stat_beta = beta / se_beta if se_beta > 0 else 0
    t_stat_alpha = alpha / se_alpha if se_alpha > 0 else 0
    p_val_alpha = 2 * (1 - stats.t.cdf(np.abs(t_stat_alpha), dof_est))
    p_val_beta = 2 * (1 - stats.t.cdf(np.abs(t_stat_beta), dof_est))
    
    print("\n--- OLS MARKET MODEL REGRESSION RESULTS ---")
    print(f"Estimation Window: {est_df.index[0].strftime('%Y-%m-%d')} to {est_df.index[-1].strftime('%Y-%m-%d')} ({n_est} days)")
    print(f"Model Equation:    R_{reit_ticker} = {alpha:.6f} + {beta:.4f} * R_{benchmark_ticker}")
    print(f"Alpha (Intercept): {alpha:.6f} (t = {t_stat_alpha:.2f}, p = {p_val_alpha:.4f})")
    print(f"Beta (Slope / Coeff): {beta:.4f} (t = {t_stat_beta:.2f}, p = {p_val_beta:.4e})")
    print(f"R-squared:         {r_squared:.4f}")
    print(f"Residual Std Err:  {s_e:.4f} (Annualized: {s_e * np.sqrt(252):.2%})")
    
    # 5. Event Window Abnormal Returns (AR) & Cumulative Abnormal Returns (CAR)
    x_evt = evt_df[benchmark_ticker].values
    y_evt = evt_df[reit_ticker].values
    
    predicted_returns = alpha + beta * x_evt
    abnormal_returns = y_evt - predicted_returns
    cumulative_abnormal_returns = np.cumsum(abnormal_returns)
    
    # Variance of prediction error for each event day:
    # Var(AR_t) = s_e^2 * (1 + 1/n_est + (R_mt - mean_x_est)^2 / ((n_est - 1) * var_x_est))
    var_ar = s_squared_e * (1.0 + 1.0 / n_est + ((x_evt - mean_x) ** 2) / ((n_est - 1) * var_x))
    se_ar = np.sqrt(var_ar)
    t_stat_ar = abnormal_returns / se_ar
    
    # Standard error of CAR:
    # For independent daily ARs, Var(CAR) ≈ sum(Var(AR_t))
    var_car = np.cumsum(var_ar)
    se_car = np.sqrt(var_car)
    t_stat_car = cumulative_abnormal_returns / se_car
    
    # Relative trading day indices relative to t0 (e.g. -20 to +20)
    evt_indices = []
    for d in evt_df.index:
        trading_day_diff = len(returns_df.loc[min(d, t0_date):max(d, t0_date)]) - 1
        if d < t0_date:
            evt_indices.append(-trading_day_diff)
        else:
            evt_indices.append(trading_day_diff)
            
    event_results_df = pd.DataFrame({
        "Trading_Day": evt_indices,
        "Date": evt_df.index.strftime("%Y-%m-%d"),
        f"Price_{reit_ticker}": price_df.loc[evt_df.index, reit_ticker].values,
        f"Price_{benchmark_ticker}": price_df.loc[evt_df.index, benchmark_ticker].values,
        f"Return_{reit_ticker}": y_evt,
        f"Return_{benchmark_ticker}": x_evt,
        "Expected_Return": predicted_returns,
        "Abnormal_Return": abnormal_returns,
        "CAR": cumulative_abnormal_returns,
        "SE_AR": se_ar,
        "t_stat_AR": t_stat_ar,
        "SE_CAR": se_car,
        "t_stat_CAR": t_stat_car
    })
    
    # Event day t=0 metrics
    t0_row = event_results_df[event_results_df["Trading_Day"] == 0]
    car_at_0 = t0_row["CAR"].values[0] if len(t0_row) > 0 else 0
    ar_at_0 = t0_row["Abnormal_Return"].values[0] if len(t0_row) > 0 else 0
    total_car = cumulative_abnormal_returns[-1]
    total_car_t = t_stat_car[-1]
    
    print("\n--- EVENT WINDOW METRICS ---")
    print(f"Event Window:      t = {evt_indices[0]} to t = +{evt_indices[-1]} ({len(evt_indices)} trading days)")
    print(f"Announcement (t=0): Date {t0_date.strftime('%Y-%m-%d')}, AR = {ar_at_0:.2%}, CAR = {car_at_0:.2%}")
    print(f"Full Window CAR:   {total_car:.2%} (t-stat = {total_car_t:.2f})")
    
    # 6. Export All CSVs
    csv_paths = {}
    
    # CSV 1: Event Window AR & CAR series
    p1 = os.path.join(output_dir, "event_window_ar_car.csv")
    event_results_df.to_csv(p1, index=False)
    csv_paths["event_window_ar_car"] = p1
    
    # CSV 2: Regression Summary Statistics
    summary_dict = {
        "Metric": [
            "Target REIT Ticker",
            "Benchmark Ticker",
            "Peer Ticker",
            "Event Date (t=0)",
            "Estimation Window Start",
            "Estimation Window End",
            "Estimation Window Trading Days",
            "Alpha (Daily Intercept)",
            "Alpha Standard Error",
            "Alpha t-statistic",
            "Alpha p-value",
            "Beta (Slope / Market Sensitivity)",
            "Beta Standard Error",
            "Beta t-statistic",
            "Beta p-value",
            "R-squared",
            "Residual Standard Error (Daily)",
            "Residual Variance (s_e^2)",
            "Annualized Residual Volatility",
            "Event Day t=0 Abnormal Return (AR)",
            "Event Day t=0 CAR",
            "Total Event Window CAR",
            "Total CAR Standard Error",
            "Total CAR t-statistic",
            "Event Window Range"
        ],
        "Value": [
            reit_ticker,
            benchmark_ticker,
            peer_ticker,
            t0_date.strftime("%Y-%m-%d"),
            est_df.index[0].strftime("%Y-%m-%d"),
            est_df.index[-1].strftime("%Y-%m-%d"),
            n_est,
            f"{alpha:.6f}",
            f"{se_alpha:.6f}",
            f"{t_stat_alpha:.4f}",
            f"{p_val_alpha:.6e}",
            f"{beta:.4f}",
            f"{se_beta:.4f}",
            f"{t_stat_beta:.4f}",
            f"{p_val_beta:.6e}",
            f"{r_squared:.4f}",
            f"{s_e:.6f}",
            f"{s_squared_e:.8f}",
            f"{s_e * np.sqrt(252):.4f}",
            f"{ar_at_0:.4%}",
            f"{car_at_0:.4%}",
            f"{total_car:.4%}",
            f"{se_car[-1]:.4%}",
            f"{total_car_t:.4f}",
            f"[{evt_indices[0]}, +{evt_indices[-1]}]"
        ]
    }
    p2 = os.path.join(output_dir, "market_model_regression_summary.csv")
    pd.DataFrame(summary_dict).to_csv(p2, index=False)
    csv_paths["regression_summary"] = p2
    
    # CSV 3: Estimation Window Raw Data
    est_export_df = pd.DataFrame({
        "Date": est_df.index.strftime("%Y-%m-%d"),
        f"Return_{reit_ticker}": y_est,
        f"Return_{benchmark_ticker}": x_est,
        "Predicted_Return": y_pred_est,
        "Residual_Error": residuals_est
    })
    p3 = os.path.join(output_dir, "estimation_window_returns.csv")
    est_export_df.to_csv(p3, index=False)
    csv_paths["estimation_window_returns"] = p3
    
    # CSV 4: Full Multi-Year Price and Return Master
    full_export_df = pd.DataFrame({
        "Date": price_df.index.strftime("%Y-%m-%d"),
        f"Price_{reit_ticker}": price_df[reit_ticker],
        f"Price_{benchmark_ticker}": price_df[benchmark_ticker],
        f"Price_{peer_ticker}": price_df[peer_ticker] if peer_ticker in price_df.columns else np.nan,
        f"Return_{reit_ticker}": returns_df[reit_ticker],
        f"Return_{benchmark_ticker}": returns_df[benchmark_ticker]
    })
    p4 = os.path.join(output_dir, "full_historical_prices_returns.csv")
    full_export_df.to_csv(p4, index=False)
    csv_paths["full_historical_prices_returns"] = p4
    
    print("\n--- EXPORTED CSVS ---")
    for k, v in csv_paths.items():
        print(f"  • {k}: {v}")
        
    # 7. Generate Institutional Plot
    plot_path = os.path.join(output_dir, "cumulative_abnormal_returns_plot.png")
    if save_plot:
        fig, axes = plt.subplots(2, 2, figsize=(15, 11), dpi=200)
        plt.subplots_adjust(hspace=0.32, wspace=0.25)
        
        # Subplot 1: Cumulative Abnormal Return (CAR) Trajectory
        ax1 = axes[0, 0]
        days = event_results_df["Trading_Day"]
        car_pct = event_results_df["CAR"] * 100
        upper_band = (event_results_df["CAR"] + 1.96 * event_results_df["SE_CAR"]) * 100
        lower_band = (event_results_df["CAR"] - 1.96 * event_results_df["SE_CAR"]) * 100
        
        ax1.plot(days, car_pct, color="#2563eb", linewidth=2.5, label="Cumulative Abnormal Return (CAR %)")
        ax1.fill_between(days, lower_band, upper_band, color="#3b82f6", alpha=0.15, label="95% Confidence Interval")
        ax1.axvline(x=0, color="#dc2626", linestyle="--", linewidth=1.8, label=f"Event Date t=0 ({event_date})")
        ax1.axhline(y=0, color="#64748b", linestyle=":", linewidth=1.2)
        ax1.scatter([0], [car_at_0 * 100], color="#dc2626", s=60, zorder=5)
        
        ax1.set_title(f"Cumulative Abnormal Return (CAR) around Event\n{reit_ticker} vs {benchmark_ticker}", fontsize=12, fontweight="bold")
        ax1.set_xlabel("Event Time (Trading Days relative to announcement t=0)", fontsize=10)
        ax1.set_ylabel("CAR (%)", fontsize=10)
        ax1.grid(True, linestyle="--", alpha=0.5)
        ax1.legend(loc="best", fontsize=9)
        
        # Subplot 2: Market Model Regression Scatter & Fit
        ax2 = axes[0, 1]
        ax2.scatter(x_est * 100, y_est * 100, color="#475569", alpha=0.6, s=25, label="Daily Returns (Estimation)")
        x_line = np.linspace(np.min(x_est), np.max(x_est), 100)
        y_line = (alpha + beta * x_line) * 100
        ax2.plot(x_line * 100, y_line, color="#dc2626", linewidth=2.2,
                 label=f"Fit: y = {alpha:.4f} + {beta:.3f}x\n(R² = {r_squared:.3f}, p < 0.001)")
        ax2.axhline(0, color="#94a3b8", linestyle=":", linewidth=1)
        ax2.axvline(0, color="#94a3b8", linestyle=":", linewidth=1)
        ax2.set_title(f"Market Model Regression (Estimation Window)\ny: {reit_ticker} | x: {benchmark_ticker}", fontsize=12, fontweight="bold")
        ax2.set_xlabel(f"Benchmark Daily Return: {benchmark_ticker} (%)", fontsize=10)
        ax2.set_ylabel(f"REIT Daily Return: {reit_ticker} (%)", fontsize=10)
        ax2.grid(True, linestyle="--", alpha=0.5)
        ax2.legend(loc="best", fontsize=9)
        
        # Subplot 3: Daily Actual Return vs Expected Market Return during Event Window
        ax3 = axes[1, 0]
        ax3.bar(days - 0.15, y_evt * 100, width=0.35, color="#3b82f6", alpha=0.85, label=f"Actual {reit_ticker} Return")
        ax3.bar(days + 0.15, predicted_returns * 100, width=0.35, color="#94a3b8", alpha=0.85, label="Expected Model Return (α + β*Rm)")
        ax3.axvline(x=0, color="#dc2626", linestyle="--", linewidth=1.5, label="Event t=0")
        ax3.axhline(0, color="#475569", linestyle="-", linewidth=0.8)
        ax3.set_title("Event Window Daily Returns: Actual vs. Model Expected", fontsize=12, fontweight="bold")
        ax3.set_xlabel("Trading Days relative to Event Date", fontsize=10)
        ax3.set_ylabel("Daily Return (%)", fontsize=10)
        ax3.grid(True, linestyle="--", alpha=0.5)
        ax3.legend(loc="best", fontsize=9)
        
        # Subplot 4: Relative Price Trajectory Comparison (Base = 100 at t = -20)
        ax4 = axes[1, 1]
        p_reit = price_df.loc[evt_df.index, reit_ticker]
        p_bench = price_df.loc[evt_df.index, benchmark_ticker]
        norm_reit = (p_reit / p_reit.iloc[0]) * 100
        norm_bench = (p_bench / p_bench.iloc[0]) * 100
        
        ax4.plot(days, norm_reit, color="#2563eb", linewidth=2.2, label=f"{reit_ticker} (First REIT)")
        ax4.plot(days, norm_bench, color="#475569", linestyle="--", linewidth=1.8, label=f"{benchmark_ticker} (Benchmark Index)")
        if peer_ticker in price_df.columns:
            p_peer = price_df.loc[evt_df.index, peer_ticker]
            norm_peer = (p_peer / p_peer.iloc[0]) * 100
            ax4.plot(days, norm_peer, color="#059669", linewidth=1.8, label=f"{peer_ticker} (Parkway Life REIT)")
            
        ax4.axvline(x=0, color="#dc2626", linestyle="--", linewidth=1.5, label="Event Date t=0")
        ax4.set_title("Normalized Price Trajectory in Event Window (Base = 100 at t = -20)", fontsize=12, fontweight="bold")
        ax4.set_xlabel("Trading Days relative to Event Date", fontsize=10)
        ax4.set_ylabel("Normalized Price Index", fontsize=10)
        ax4.grid(True, linestyle="--", alpha=0.5)
        ax4.legend(loc="best", fontsize=9)
        
        plt.suptitle(f"Event Study Analysis: {reit_ticker} Divestment Announcement ({event_date})\nMarket Model vs {benchmark_ticker}",
                     fontsize=14, fontweight="bold", y=0.995)
        fig.tight_layout()
        fig.savefig(plot_path, dpi=200, bbox_inches="tight")
        print(f"Generated and saved institutional graph: {plot_path}")
        if show_plot:
            plt.show()
        plt.close(fig)

    # Helper function to get all CSVs as text or zipped bytes
    def get_csv_bundle(as_zip=False):
        """Returns all generated CSV data either as a dict of strings or a zip binary buffer."""
        bundle_dict = {}
        for name, path in csv_paths.items():
            with open(path, "r", encoding="utf-8") as f:
                bundle_dict[f"{name}.csv"] = f.read()
                
        if not as_zip:
            return bundle_dict
            
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fname, content in bundle_dict.items():
                zf.writestr(fname, content)
            if save_plot and os.path.exists(plot_path):
                with open(plot_path, "rb") as pf:
                    zf.writestr("cumulative_abnormal_returns_plot.png", pf.read())
        mem_zip.seek(0)
        return mem_zip.getvalue()

    return {
        "reit_ticker": reit_ticker,
        "benchmark_ticker": benchmark_ticker,
        "event_date": t0_date.strftime("%Y-%m-%d"),
        "alpha": alpha,
        "beta": beta,
        "r_squared": r_squared,
        "p_value_alpha": p_val_alpha,
        "p_value_beta": p_val_beta,
        "std_err_resid": s_e,
        "car_at_t0": car_at_0,
        "car_total": total_car,
        "car_t_stat": total_car_t,
        "csv_paths": csv_paths,
        "plot_path": plot_path,
        "event_results_df": event_results_df,
        "regression_summary_df": pd.DataFrame(summary_dict),
        "get_csv_bundle": get_csv_bundle
    }

if __name__ == "__main__":
    # Execute analysis for First REIT vs Straits Times Index (^STI) on 1 April 2026 announcement
    res = analyze_cumulative_abnormal_returns(
        reit_ticker="AW9U.SI",
        benchmark_ticker="^STI",
        peer_ticker="C2PU.SI",
        event_date="2026-04-01",
        save_plot=True
    )
    print("\nSUCCESS! Event study execution completed.")
    print(f"Regression Equation: R_FirstREIT = {res['alpha']:.5f} + {res['beta']:.4f} * R_STI")
    print(f"Total Cumulative Abnormal Return (CAR): {res['car_total']:.2%}")
