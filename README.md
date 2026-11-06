# First REIT (SGX: AW9U) · Institutional Strategic Critique & Econometric Capital Reallocation Model

> **Central Investment Question**: *Is First REIT's Divestment Strategy Creating Value for Unitholders, or Simply Replacing One Set of Risks with Another?*

An interactive institutional research web platform, econometric event study, and capital reallocation model evaluating the complete Indonesian hospital portfolio divestment by First Real Estate Investment Trust (**SGX: AW9U**).

---

## 📌 Executive Summary

First REIT's divestment strategy—commenced under its December 2021 "2.0 Growth Strategy" and culminating in the **S$471.5M Tranche 1** Indonesian hospital exit completed in October 2026 (with an additional **S$294.8M** under a Tranche 2 put option expiring 31 March 2027)—presents a classic corporate finance dilemma: **A defensibly necessary balance-sheet rescue that nevertheless fails as an accretive capital reallocation strategy.**

### Key Findings:
1. **The Necessity (Chapter 1)**: Financially mandatory. Without liquidating the Indonesian portfolio, First REIT faced imminent covenant breaches on its S$262.7M syndicated debt facility in May 2026 and an un-refinanceable S$100M CGIF-guaranteed bond on 19 October 2026.
2. **Disposal Economics (Chapter 2)**: The headline 2.1% valuation premium (S$9.7M) is an accounting illusion. Evaluated against original purchase costs, the 14 Indonesian properties crystallized a permanent capital loss of **-S$137.96M (-15.3%)**, and post-IPO acquisitions lost **-S$210.89M (-27.8%)**.
3. **Market Reaction (Chapter 3)**: Econometric market model regression ($\alpha = 0.0003$, $\beta = 0.3812$) against the Straits Times Index (`^STI`) proves the market reacted negatively to the divestment strategy announcement (**-3.23% CAR**) and post-T1 execution (**-4.39% CAR**), signaling deep skepticism over unreplaced earnings.
4. **The Reallocation Dilemma (Chapter 4)**: Insurmountable yield cliff. First REIT disposed of assets generating an 10.88% gross yield (~7.5% net cash yield) to recycle capital into developed markets (Japan, Singapore, Australia) where 4.0%–5.5% healthcare yields offer zero or negative spreads over local debt costs.
5. **Sell vs. Hold Counterfactual (Chapter 5)**: While selling rescued credit solvency (slashing gearing from 45.7% to 15.6%), it eroded unitholder equity value by -58.1% (NAV per unit down from 50.1¢ peak to 21.0¢).
6. **Transaction Evaluation (Chapter 6)**: Tranche 1 generated S$464.2M net cash, paid down S$362.7M debt, but left a net annual distributable cash deficit of **-S$16.46M/yr**, dropping core DPU to 1.39¢ and forward consensus to 1.01¢.
7. **Cost of Delayed Reinvestment (Chapter 7)**:
   * **Model A (DPU Recovery Deadline)**: Holding idle cash destroys **S$1.17M to S$2.55M/yr** in lost net income, creating a strict **4.5 to 7.0-year economic deadline** before cash drag outweighs any redeployment benefit.
   * **Model B (Gordon Growth P/NAV)**: To recover to parity (1.00× P/NAV at 10% $k_e$), First REIT requires an unachievable **2.08¢ DPU (+106% accretion)**.
   * **Model C (Financing Wall)**: Trading at **0.86× P/NAV**, First REIT is locked out of equity-funded acquisitions; raising S$50M equity dilutes NAV by **-4.3%**, and raising S$150M dilutes NAV by **-10.8%**.
8. **Final Verdict (Chapter 8)**: The divestment was necessary for survival, but value-destructive for equity. The Manager should execute Tranche 2, exhaust debt-funded acquisitions without equity dilution, and evaluate a total portfolio wind-up if trading discounts persist.

---

## 🚀 Interactive Web Features

The web platform features six interactive financial modeling tools:

* **Property Portfolio & Capital Loss Matrix**: Interactive audit of all 14 Indonesian hospital assets, comparing purchase costs, peak appraisals, and agreed exit prices with SGX filing links.
* **Multi-Year Trajectory & Cliff Analysis**: Dynamic bar chart engine visualizing the collapse in recurring DPU (2.61¢ to 1.01¢), book value erosion (50.1¢ to 21.0¢), and debt deleveraging (45.7% to 15.6%).
* **Yield Reinvestment Simulator**: Real-time interactive calculation engine modeling post-divestment equity redeployment, target asset yields (3.5%–7.5%), debt borrowing costs, and resulting DPU accretion/dilution.
* **Cumulative Abnormal Return (CAR) Visualizer**: Interactive econometric event study chart tracking First REIT (`AW9U.SI`) vs. the Straits Times Index (`^STI`) across all transaction event windows.
* **Cost of Waiting Dynamic Simulator**: Sliders to stress-test target P/NAV, cost of equity ($k_e$), long-term growth rate ($g$), and required DPU to break through the equity financing wall.
* **Historical Year-End P/NAV Tracker**: 9-year annual comparison (2018–2026) benchmarked against premier healthcare peer Parkway Life REIT (`C2PU.SI`).
* **Verified Sources & Primary Evidence Library**: 18 direct regulatory filings, circulars, and institutional broker reports (Phillip Securities, Beansprout Research, Maybank IBG).
* **One-Click Data Exporters**: 13 individual CSV datasets, complete JSON bundle, and a consolidated master ZIP archive.

---

## 🛠️ Project Structure

```
.
├── index.html                           # Hallmark-certified interactive web application
├── car_analysis.py                      # Python yfinance market model & CAR event study script
├── cost_of_waiting_analysis.py          # Python financial runway & cost of waiting simulation
├── p_to_nav_analysis.py                 # Python year-end P/NAV extraction & peer benchmarking
├── data_bundle.json                     # Complete consolidated JSON dataset
├── package.json                         # Node server metadata & local runner script
├── cumulative_abnormal_returns_plot.png # High-res econometric CAR event study plot
├── cost_of_waiting_tradeoff_plot.png    # High-res reinvestment delay trade-off graphic
├── year_end_p_to_nav_trend.png          # High-res P/NAV valuation multiple trend plot
├── data_exports/                        # 13 verified CSV files & master ZIP archive
└── event_study_output/                  # Econometric regression output tables & logs
```

---

## 💻 Quick Start & Local Execution

### Option 1: Direct Browser
Open `index.html` directly in any modern web browser.

### Option 2: Local HTTP Server
```bash
# Clone the repository
git clone https://github.com/brandon119-04/first-reit-critique.git
cd first-reit-critique

# Serve locally using npx or Python
npx serve .
# or:
python3 -m http.server 3000
```
Open `http://localhost:3000` in your browser.

---

## 📊 Running the Econometric Models

### Requirements
```bash
pip install yfinance pandas numpy scipy matplotlib
```

### 1. Cumulative Abnormal Return (CAR) Event Study
```bash
python3 car_analysis.py
```
Regresses First REIT (`AW9U.SI`) returns on Straits Times Index (`^STI`) using a 250-trading-day estimation window to calculate Alpha, Beta, Abnormal Returns ($AR$), and Cumulative Abnormal Returns ($CAR$).

### 2. Year-End Price-to-NAV Analysis
```bash
python3 p_to_nav_analysis.py
```
Pulls fiscal year-end closing prices for First REIT and Parkway Life REIT from Yahoo Finance and calculates historical P/NAV discounts.

### 3. Cost of Waiting & Reinvestment Runway
```bash
python3 cost_of_waiting_analysis.py
```
Runs the 3-model institutional framework (Model A: DPU recovery deadline; Model B: Gordon Growth required DPU; Model C: Non-dilutive financing capacity).

---

## 📜 Citations & Sources

Every metric cited in this research is traceable to primary disclosures:
1. **SGX Circular**: Proposed Divestment of Indonesian Hospital Portfolio (Tranche 1 & Tranche 2 Put Option)
2. **First REIT Annual Reports**: Audited Financial Disclosures (FY2018 through FY2025)
3. **Phillip Securities Research**: First REIT Equity Research Reports (October 2026 & AGM Updates)
4. **Beansprout Research**: Healthcare REIT Sector Review & Yield Analysis
5. **SGX Circular**: Master Lease Agreement (MLA) Restructuring & Renounceable Rights Issue (December 2020 – February 2021)
6. **MAS Regulatory Framework**: Code on Collective Investment Schemes (Property Funds Appendix 6)

---

## ⚖️ License & Attribution
Antigravity Research & Financial Engineering. Built with institutional standards and Hallmark design principles.
