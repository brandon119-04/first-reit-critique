import os
import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

out_dir = "/Users/brandonwong/.gemini/antigravity/scratch/first-reit-2024-2025-extraction"
os.makedirs(out_dir, exist_ok=True)

# ---------------------------------------------------------
# 1. FINANCIAL HIGHLIGHTS & KEY METRICS (FY2021 - FY2025)
# ---------------------------------------------------------
highlights_rows = [
    ["Category", "Metric", "Unit", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "YoY Change (FY25 vs FY24)", "AR2025 Reference"],
    ["Income Statement Summary", "Rental and Other Income", "S$'million", 102.3, 111.3, 108.6, 102.2, 100.5, "-1.7%", "p. 2, 4, 125"],
    ["Income Statement Summary", "Rental and Other Income (Ex-FRS 116 Straight-Lining)", "S$'million", "-", "-", "-", 92.0, 93.1, "+1.2%", "p. 25, 125, Note 4"],
    ["Income Statement Summary", "Net Property and Other Income (NPI)", "S$'million", 100.2, 108.6, 105.3, 98.5, 97.3, "-1.1%", "p. 2, 4, 125"],
    ["Income Statement Summary", "Cash Net Property Income (Ex-FRS 116 Straight-Lining)", "S$'million", "-", "-", "-", 88.2, 89.9, "+1.9%", "Derived (Note 16, 17)"],
    ["Income Statement Summary", "Distributable Amount", "S$'million", 42.1, 52.4, 51.4, 49.3, 45.8, "-7.0%", "p. 2, 4, 126"],
    ["Per Unit Metrics", "Distribution Per Unit (DPU) - Full Year", "Singapore cents", 2.61, 2.64, 2.48, 2.36, 2.17, "-8.1%", "p. 2, 4, 126"],
    ["Per Unit Metrics", "DPU - 1Q", "Singapore cents", 0.65, 0.66, 0.62, 0.60, 0.58, "-3.3%", "p. 44, 173"],
    ["Per Unit Metrics", "DPU - 2Q", "Singapore cents", 0.65, 0.66, 0.62, 0.60, 0.55, "-8.3%", "p. 44, 173"],
    ["Per Unit Metrics", "DPU - 3Q", "Singapore cents", 0.65, 0.66, 0.62, 0.58, 0.52, "-10.3%", "p. 44, 173"],
    ["Per Unit Metrics", "DPU - 4Q", "Singapore cents", 0.66, 0.66, 0.62, 0.58, 0.52, "-10.3%", "p. 44, 173"],
    ["Per Unit Metrics", "Basic Earnings Per Unit (EPU)", "Singapore cents", 4.00, 1.56, 2.98, 1.68, 1.29, "-23.2%", "p. 4, 125, 171"],
    ["Per Unit Metrics", "Diluted Earnings Per Unit (EPU)", "Singapore cents", 4.00, 1.56, 2.98, 1.66, 1.29, "-22.3%", "p. 125, 171"],
    ["Balance Sheet Summary", "Assets Under Management (AUM / Inv Properties)", "S$'million", 962.4, 1145.3, 1139.5, 1118.0, 1022.6, "-8.5%", "p. 2, 4, 124"],
    ["Balance Sheet Summary", "Total Assets", "S$'million", 1019.0, 1207.5, 1187.3, 1160.5, 1089.0, "-6.2%", "p. 124"],
    ["Balance Sheet Summary", "Total Gross Borrowings (Face Value)", "S$'million", 369.9, 463.7, 458.9, 459.6, 458.0, "-0.3%", "p. 161-162"],
    ["Balance Sheet Summary", "Total Carrying Borrowings", "S$'million", 366.8, 459.7, 452.7, 453.5, 454.8, "+0.3%", "p. 124, 161"],
    ["Balance Sheet Summary", "Total Liabilities", "S$'million", 425.1, 543.5, 527.2, 528.2, 528.5, "+0.0%", "p. 124"],
    ["Balance Sheet Summary", "Net Assets Attributable to Unitholders", "S$'million", 591.9, 632.0, 626.9, 599.0, 527.2, "-12.0%", "p. 124"],
    ["Balance Sheet Summary", "Perpetual Securities", "S$'million", 33.3, 33.3, 33.3, 33.3, 33.3, "0.0% (Redeemed 8 Jan 2026)", "p. 124, 166"],
    ["Balance Sheet Summary", "Total Net Assets", "S$'million", 625.2, 665.3, 660.2, 632.3, 560.5, "-11.4%", "p. 124"],
    ["Per Unit & Capital Structure", "Units in Issue (Year-End)", "'000", 1613029, 2058581, 2076925, 2094447, 2110969, "+0.8%", "p. 4, 124, 167"],
    ["Per Unit & Capital Structure", "Weighted Average Units in Issue (Basic)", "'000", 1499382, 1981551, 2067938, 2086175, 2111058, "+1.2%", "p. 4, 171"],
    ["Per Unit & Capital Structure", "Net Asset Value (NAV) Per Unit", "Singapore cents", 36.69, 30.70, 30.18, 28.60, 24.97, "-12.7%", "p. 2, 5, 124"],
    ["Portfolio Cap Rates & Yields", "Portfolio Implied Gross Cap Rate (Ex-FRS 116 Cash Rent / AUM)", "%", "-", "-", "-", "8.23%", "9.10%", "+87 bps", "Derived (pp. 26-41, 133)"],
    ["Portfolio Cap Rates & Yields", "Portfolio Implied Cash NPI Cap Rate (Ex-FRS 116 Cash NPI / AUM)", "%", "-", "-", "-", "7.89%", "8.79%", "+90 bps", "Derived (Note 16, 17, 28)"],
    ["Portfolio Cap Rates & Yields", "Portfolio Accounting NPI Cap Rate (Reported NPI incl. FRS 116 / AUM)", "%", "10.41%", "9.48%", "9.24%", "8.81%", "9.52%", "+71 bps", "Derived (p. 4, 124-125)"],
    ["Portfolio Cap Rates & Yields", "Singapore Implied Gross / Cash NPI Cap Rate (Ex-FRS 116)", "%", "-", "-", "-", "14.47% / 12.69%", "15.71% / 14.10%", "+124 bps / +141 bps", "p. 33-34, 134, 175"],
    ["Portfolio Cap Rates & Yields", "Japan Implied Gross / Cash NPI Cap Rate (Ex-FRS 116)", "%", "-", "-", "-", "5.41% / 4.46%", "5.78% / 4.90%", "+37 bps / +44 bps", "p. 26-32, 139-141, 175"],
    ["Portfolio Cap Rates & Yields", "Indonesia (14 Active) Implied Gross / Cash NPI Cap Rate (Ex-FRS 116)", "%", "-", "-", "-", "8.95% / 8.84%", "9.66% / 9.56%", "+71 bps / +72 bps", "p. 35-41, 135-138, 175"],
    ["Portfolio Cap Rates & Yields", "Valuer Terminal Cap Rate - Indonesia Restructured HGB Hospitals", "%", "-", "-", "-", "9.25% - 10.66%", "9.25% - 11.18%", "Upper bound +52 bps", "Note 4(ii), p. 159-160"],
    ["Portfolio Cap Rates & Yields", "Valuer Terminal Cap Rate - Indonesia Non-Restructured HGB", "%", "-", "-", "-", "9.50%", "9.00% - 9.86%", "-50 to +36 bps", "Note 4(ii), p. 159-160"],
    ["Portfolio Cap Rates & Yields", "Valuer Terminal Cap Rate - Japan Nursing Homes (14 Freehold)", "%", "-", "-", "-", "4.30% - 5.00%", "4.30% - 5.00%", "Unchanged", "Note 4(ii), p. 160"],
    ["Portfolio Cap Rates & Yields", "Valuer DCF Discount Rate - Singapore / Japan / Indonesia", "%", "-", "-", "-", "8.75% / 4.00-4.70% / 9.00-13.00%", "8.75% / 4.00-4.70% / 9.30-13.20%", "ID +20 to +65 bps", "Note 4(ii), p. 159-160"],
    ["Key Financial Ratios", "Interest Coverage Ratio (ICR)", "times", 5.2, 5.0, 4.1, 3.6, 3.7, "+0.1x", "p. 2, 5"],
    ["Key Financial Ratios", "Aggregate Leverage (Gearing Ratio)", "%", "33.6%", "38.5%", "38.7%", "39.6%", "42.1%", "+250 bps", "p. 2, 5"],
    ["Key Financial Ratios", "All-in Cost of Debt per Annum", "%", "4.2%", "3.7%", "5.0%", "5.0%", "4.5%", "-50 bps", "p. 2, 5"],
    ["Key Financial Ratios", "Debt on Fixed Rates or Hedged", "%", "56.1%", "59.6%", "87.2%", "56.9%", "46.1%", "-10.8% pts", "p. 5"],
    ["Key Financial Ratios", "Weighted Average Term to Debt Maturity", "years", 1.2, 3.4, 3.6, 2.5, 1.5, "-1.0 yr", "p. 20"],
    ["Key Financial Ratios", "Percentage of Debt Under Social Finance", "%", "-", "96.7%", "76.8%", "77.3%", "78.7%", "+1.4% pts", "p. 2, 5"],
    ["Key Financial Ratios", "Derivative Financial Instruments as % of NAV", "%", "-", "0.94%", "0.49%", "0.01%", "0.00%", "-0.01% pts", "p. 5"],
    ["Unit Price & Market Metrics", "Opening Unit Price", "S$", "-", "-", "-", 0.265, 0.255, "-3.8%", "p. 44"],
    ["Unit Price & Market Metrics", "Closing Unit Price (31 Dec)", "S$", 0.305, 0.260, 0.265, 0.255, 0.275, "+7.8%", "p. 44"],
    ["Unit Price & Market Metrics", "Highest Unit Price", "S$", "-", "-", "-", 0.280, 0.290, "+3.6%", "p. 44"],
    ["Unit Price & Market Metrics", "Lowest Unit Price", "S$", "-", "-", "-", 0.240, 0.245, "+2.1%", "p. 44"],
    ["Unit Price & Market Metrics", "Price-to-NAV Ratio (Year-End)", "times", 0.83, 0.85, 0.88, 0.89, 1.10, "+0.21x", "Derived / p. 44"],
    ["Unit Price & Market Metrics", "Market Capitalisation (31 Dec)", "S$'million", 492.0, 535.2, 550.4, 534.1, 580.5, "+8.7%", "p. 44"],
    ["Unit Price & Market Metrics", "Distribution Yield (at Closing Price)", "%", "8.6%", "10.2%", "9.4%", "9.3%", "7.9%", "-140 bps", "p. 44"],
    ["Unit Price & Market Metrics", "Total Return (1-Year Price + Distribution)", "%", "-", "-", "-", "5.8%", "17.3%", "+11.5% pts", "p. 44"],
    ["Unit Price & Market Metrics", "Total Trading Volume", "million units", "-", "-", "-", 296.7, 399.8, "+34.7%", "p. 44"],
    ["Unit Price & Market Metrics", "Average Daily Traded Volume", "million units", "-", "-", "-", 1.177, 1.586, "+34.7%", "p. 44"],
    ["Geography Breakdown - Income", "Indonesia Rental & Other Income", "S$'million (%)", "-", "-", "-", "84.6 (82.8%)", "82.9 (82.5%)", "-1.9%", "p. 4, 175"],
    ["Geography Breakdown - Income", "Japan Rental & Other Income", "S$'million (%)", "-", "-", "-", "13.4 (13.1%)", "13.4 (13.3%)", "-0.3%", "p. 4, 175"],
    ["Geography Breakdown - Income", "Singapore Rental & Other Income", "S$'million (%)", "-", "-", "-", "4.2 (4.1%)", "4.2 (4.2%)", "-0.3%", "p. 4, 175"],
    ["Geography Breakdown - Income Ex-FRS116", "Indonesia Rental Income (Ex-FRS 116 Straight-Lining)", "IDR billion / S$'m", "621 / 76", "645 / 75", "694 / 74", "727 / 74", "764 / 75", "+5.1% (IDR) / +1.4% (SGD)", "EGM Pres p. 11 / AR p. 169"],
    ["Geography Breakdown - AUM", "Indonesia AUM (Carrying Value)", "S$'million (%)", "-", "-", "-", "839.9 (75.1%)", "762.0 (74.5%)", "-9.3%", "p. 4, 133"],
    ["Geography Breakdown - AUM", "Japan AUM (Carrying Value)", "S$'million (%)", "-", "-", "-", "247.7 (22.2%)", "232.0 (22.7%)", "-6.4%", "p. 4, 133"],
    ["Geography Breakdown - AUM", "Singapore AUM (Carrying Value)", "S$'million (%)", "-", "-", "-", "30.4 (2.7%)", "28.6 (2.8%)", "-5.8%", "p. 4, 133"],
    ["Portfolio Operational Summary", "Number of Properties (Year-End)", "count", 19, 32, 32, 32, 31, "-1 (IAHCC divested 4 Dec 2025)", "p. 2, 24"],
    ["Portfolio Operational Summary", "Total Gross Floor Area (GFA)", "sqm", 353569, 444558, 446664, 450085, 432159, "-4.0% (excl. 17,926 sqm IAHCC)", "p. 2, 134-142"],
    ["Portfolio Operational Summary", "Maximum Beds / Rooms", "beds / rooms", "-", "-", "-", 6498, 6305, "-193 rooms (IAHCC)", "p. 2, 24"],
    ["Portfolio Operational Summary", "Committed Occupancy", "%", "100.0%", "100.0%", "100.0%", "100.0%", "100.0%", "Unchanged", "p. 2"],
    ["Portfolio Operational Summary", "Weighted Average Lease Expiry (WALE by GFA)", "years", 11.5, 12.5, 11.5, 10.6, 10.0, "-0.6 yrs", "p. 2, 25"],
]

# ---------------------------------------------------------
# 2. STATEMENTS OF FINANCIAL POSITION (FY2024 vs FY2025 + 1H2026)
# ---------------------------------------------------------
bs_rows = [
    ["Section", "Line Item", "Note (AR25)", "Group FY2024 (S$'000)", "Group FY2025 (S$'000)", "Group YoY Change (%)", "Trust FY2024 (S$'000)", "Trust FY2025 (S$'000)", "Trust YoY Change (%)", "Group 1H2026 (30 Jun 2026, S$'000)"],
    ["Non-Current Assets", "Plant and equipment", "", 64, 49, "-23.4%", 0, 0, "N/A", 41],
    ["Non-Current Assets", "Investment properties", "Note 4", 1118008, 1022619, "-8.5%", 30400, 28650, "-5.8%", 529307],
    ["Non-Current Assets", "Investments in subsidiaries", "Note 5", 0, 0, "N/A", 746614, 714723, "-4.3%", 0],
    ["Non-Current Assets", "Loan to a subsidiary", "Note 6", 0, 0, "N/A", 12639, 1998, "-84.2%", 0],
    ["Non-Current Assets", "Total Non-Current Assets", "", 1118072, 1022668, "-8.5%", 789653, 745371, "-5.6%", 529348],
    ["Current Assets", "Loan to subsidiaries", "Note 6", 0, 0, "N/A", 4190, 8406, "+100.6%", 0],
    ["Current Assets", "Trade and other receivables", "Note 7", 6724, 9372, "+39.4%", 17850, 17223, "-3.5%", 674],
    ["Current Assets", "Other assets", "Note 8", 771, 732, "-5.1%", 406, 394, "-3.0%", 2258],
    ["Current Assets", "Derivative financial instruments", "Note 15", 260, 0, "-100.0%", 260, 0, "-100.0%", 7],
    ["Current Assets", "Cash and cash equivalents", "Note 9", 34674, 56197, "+62.1%", 11796, 36662, "+210.8%", 34747],
    ["Current Assets", "Assets held for sale (11 ID properties in 1H26)", "Note 7 (1H26)", 0, 0, "N/A", 0, 0, "N/A", 460074],
    ["Current Assets", "Total Current Assets", "", 42429, 66301, "+56.3%", 34502, 62685, "+81.7%", 497760],
    ["Total Assets", "TOTAL ASSETS", "", 1160501, 1088969, "-6.2%", 824155, 808056, "-2.0%", 1027108],
    ["Non-Current Liabilities", "Deferred tax liabilities", "Note 10", 49369, 47041, "-4.7%", 0, 0, "N/A", 37059],
    ["Non-Current Liabilities", "Borrowings (non-current)", "Note 11", 452602, 194168, "-57.1%", 336394, 98757, "-70.6%", 91704],
    ["Non-Current Liabilities", "Loan from a subsidiary", "Note 6", 0, 0, "N/A", 0, 3984, "N.M.", 0],
    ["Non-Current Liabilities", "Other liabilities (security deposits, non-current)", "Note 12", 8069, 7773, "-3.7%", 2121, 2163, "+2.0%", 7614],
    ["Non-Current Liabilities", "Total Non-Current Liabilities", "", 510040, 248982, "-51.2%", 338515, 104904, "-69.0%", 136377],
    ["Current Liabilities", "Borrowings (current)", "Note 11", 880, 260681, "+29522.8%", 0, 246123, "N.M.", 375150],
    ["Current Liabilities", "Other liabilities (security deposits & rentals in advance)", "Note 12", 3246, 3036, "-6.5%", 104, 106, "+1.9%", 1149],
    ["Current Liabilities", "Loans from subsidiaries", "Note 6", 0, 0, "N/A", 32861, 30437, "-7.4%", 0],
    ["Current Liabilities", "Derivative financial instruments", "Note 15", 192, 0, "-100.0%", 192, 0, "-100.0%", 459],
    ["Current Liabilities", "Trade and other payables", "Note 13", 12641, 14482, "+14.6%", 47966, 75020, "+56.4%", 14813],
    ["Current Liabilities", "Income tax payable", "", 1230, 1293, "+5.1%", 7, 5, "-28.6%", 1088],
    ["Current Liabilities", "Liabilities directly associated with assets held for sale", "Note 7 (1H26)", 0, 0, "N/A", 0, 0, "N/A", 10631],
    ["Current Liabilities", "Total Current Liabilities", "", 18189, 279492, "+1436.6%", 81130, 351691, "+333.5%", 403290],
    ["Total Liabilities", "TOTAL LIABILITIES", "", 528229, 528474, "+0.05%", 419645, 456595, "+8.8%", 539667],
    ["Net Assets & Equity", "Net assets attributable to unitholders", "Note 15", 598990, 527213, "-12.0%", 371228, 318179, "-14.3%", 487441],
    ["Net Assets & Equity", "Perpetual securities holders' funds", "Note 14", 33282, 33282, "0.0%", 33282, 33282, "0.0%", 0],
    ["Net Assets & Equity", "TOTAL NET ASSETS", "", 632272, 560495, "-11.4%", 404510, 351461, "-13.1%", 487441],
    ["Per Unit Metrics", "Units in issue ('000)", "Note 15", 2094447, 2110969, "+0.8%", 2094447, 2110969, "+0.8%", 2119234],
    ["Per Unit Metrics", "Net asset value (NAV) per unit attributable to unitholders (cents)", "Note 15", 28.60, 24.97, "-12.7%", 17.72, 15.07, "-15.0%", 23.00],
]

# ---------------------------------------------------------
# 3. STATEMENTS OF TOTAL RETURN & DISTRIBUTION (FY2024 vs FY2025 + 1H2026)
# ---------------------------------------------------------
pl_rows = [
    ["Statement / Section", "Line Item", "Note (AR25)", "Group FY2024 (S$'000)", "Group FY2025 (S$'000)", "Group YoY Change (%)", "Trust FY2024 (S$'000)", "Trust FY2025 (S$'000)", "Group 1H2025 (S$'000)", "Group 1H2026 (S$'000)"],
    ["Statement of Total Return", "Rental and other income", "Note 16", 102211, 100531, "-1.6%", 30115, 30083, 50467, 46277],
    ["Income Breakdown (Note 16)", "  - Rental income from investment properties (incl. variable & FRS 116)", "Note 16", 102211, 100531, "-1.6%", 4241, 4226, "-", "-"],
    ["Income Breakdown (Note 16)", "  - Dividend income from subsidiaries (Trust only)", "Note 16", 0, 0, "N/A", 25874, 25857, "-", "-"],
    ["Income Breakdown (Note 16)", "  - Memo: Variable rent component included in rental income", "Note 16", 2778, 3926, "+41.3%", 0, 0, "-", "-"],
    ["Income Breakdown (Note 16)", "  - Memo: FRS 116 rental straight-lining adjustment", "Note 4 / Note A", 10227, 7446, "-27.2%", -306, -268, 3827, 2473],
    ["Income Breakdown (Note 16)", "  - Memo: Rental & other income ex-FRS 116 straight-lining", "Derived", 91984, 93085, "+1.2%", 30421, 30351, 46640, 43804],
    ["Statement of Total Return", "Property operating expenses", "Note 17", -3758, -3191, "-15.1%", -226, -254, -1536, -1294],
    ["OpEx Breakdown (Note 17)", "  - Property tax expenses", "Note 17", -1395, -1428, "+2.4%", 0, 0, -724, -657],
    ["OpEx Breakdown (Note 17)", "  - Valuation expenses", "Note 17", -251, -311, "+23.9%", -16, -24, -152, -105],
    ["OpEx Breakdown (Note 17)", "  - Professional fees", "Note 17", -1273, -930, "-26.9%", -30, -32, -454, -416],
    ["OpEx Breakdown (Note 17)", "  - Repairs and maintenance", "Note 17", -482, -236, "-51.0%", 0, 0, -76, -58],
    ["OpEx Breakdown (Note 17)", "  - Others (Insurance, etc.)", "Note 17", -357, -286, "-19.9%", -180, -198, -130, -58],
    ["Statement of Total Return", "NET PROPERTY AND OTHER INCOME (NPI)", "", 98453, 97340, "-1.1%", 29889, 29829, 48931, 44983],
    ["Statement of Total Return", "Finance income", "Note 18", 362, 269, "-25.7%", 1241, 687, 153, 99],
    ["Statement of Total Return", "Manager's management fees", "Note 19", -8690, -8502, "-2.2%", -8647, -8458, -4243, -3884],
    ["Mgmt Fee Breakdown (Note 19)", "  - Base fees", "Note 19", -3729, -3575, "-4.1%", -3686, -3531, -1782, -1629],
    ["Mgmt Fee Breakdown (Note 19)", "  - Performance fees", "Note 19", -4961, -4927, "-0.7%", -4961, -4927, -2461, -2255],
    ["Statement of Total Return", "Asset management fees (Japan portfolio)", "", -1006, -1008, "+0.2%", 0, 0, -506, -461],
    ["Statement of Total Return", "Trustee fees", "", -355, -341, "-3.9%", -355, -341, -171, -156],
    ["Statement of Total Return", "Finance costs", "Note 18", -22773, -20900, "-8.2%", -20662, -18852, -10871, -9474],
    ["Finance Cost Breakdown (Note 18)", "  - Interest expense on borrowings", "Note 18", -18242, -16195, "-11.2%", -16364, -14346, -8350, -6857],
    ["Finance Cost Breakdown (Note 18)", "  - Amortisation of transaction costs (borrowings, CGIF guarantee, LC)", "Note 18", -4531, -4705, "+3.8%", -4298, -4506, -2521, -2617],
    ["Statement of Total Return", "Other expenses", "Note 20", -1553, -4647, "+199.2%", -1516, -4871, -1071, -1104],
    ["Other Expenses Breakdown (Note 20)", "  - Project expenses (Strategic review / divestment advisory)", "Note 20", -239, -3431, "+1335.6%", -239, -3431, -272, -383],
    ["Other Expenses Breakdown (Note 20)", "  - Handling, processing, professional & other trust expenses", "Note 20", -1314, -1216, "-7.5%", -1277, -1440, -799, -721],
    ["Statement of Total Return", "NET INCOME / (LOSS)", "", 64438, 62211, "-3.5%", -50, -2006, 32222, 30003],
    ["Non-Operating / Fair Value Items", "Net fair value (losses)/gains on investment properties", "Note 4", -9578, -3528, "-63.2%", -1300, -1645, 2632, -2473],
    ["Non-Operating / Fair Value Items", "Losses on disposal of subsidiary (PT KSS / IAHCC)", "Note 30", 0, -7535, "N.M.", 0, 0, 0, 0],
    ["Non-Operating / Fair Value Items", "Reversal of/(Impairment losses) on investments in subsidiaries", "Note 5", 0, 0, "N/A", 7012, -5601, 0, 0],
    ["Non-Operating / Fair Value Items", "Net fair value gains/(losses) of derivative financial instruments", "", 440, 473, "+7.5%", 440, 473, 267, -452],
    ["Non-Operating / Fair Value Items", "Net foreign exchange (losses)/gains", "", -4724, -8308, "+75.9%", -181, 30, -3824, -3836],
    ["Statement of Total Return", "TOTAL RETURN / (LOSS) BEFORE INCOME TAX", "", 50576, 43313, "-14.4%", 5921, -8749, 31297, 23242],
    ["Statement of Total Return", "Income tax expenses", "Note 22", -13822, -14331, "+3.7%", 0, 0, -7364, -7535],
    ["Tax Breakdown (Note 22)", "  - Current tax expense", "Note 22", -8756, -8780, "+0.3%", 0, 0, "-", "-"],
    ["Tax Breakdown (Note 22)", "  - Withholding tax", "Note 22", -3660, -4601, "+25.7%", 0, 0, "-", "-"],
    ["Tax Breakdown (Note 22)", "  - Deferred tax expense (Note 10)", "Note 22", -1406, -950, "-32.4%", 0, 0, "-", "-"],
    ["Statement of Total Return", "TOTAL RETURN / (LOSS) AFTER INCOME TAX", "", 36754, 28982, "-21.1%", 5921, -8749, 23933, 15707],
    ["Other Comprehensive Income", "Foreign currency translation differences (foreign operations)", "", -3169, -56459, "+1681.6%", 0, 0, -37381, -35263],
    ["Other Comprehensive Income", "TOTAL COMPREHENSIVE RETURN / (LOSS) FOR THE YEAR", "", 33585, -27477, "N.M.", 5921, -8749, -13448, -19556],
    ["Attribution of Total Return", "Total return attributable to: Unitholders of the Trust", "", 35093, 27326, "-22.1%", 4260, -10405, 23111, 15675],
    ["Attribution of Total Return", "Total return attributable to: Perpetual securities holders", "", 1661, 1656, "-0.3%", 1661, 1656, 822, 32],
    ["Statement of Distribution", "Amount available for distribution at beginning of year", "", 12906, 12113, "-6.1%", 12906, 12113, 12113, 11028],
    ["Statement of Distribution", "Total return/(loss) for the year after income tax", "", 36754, 28982, "-21.1%", 5921, -8749, 23933, 15707],
    ["Statement of Distribution", "Adjustments for tax purposes (Note A)", "Note A", 12562, 16869, "+34.3%", 43395, 54600, -65, 5030],
    ["Statement of Distribution", "ANNUAL DISTRIBUTABLE AMOUNT (Current Year Generation)", "", 49316, 45851, "-7.0%", 49316, 45851, 23868, 20737],
    ["Statement of Distribution", "Total amount available for distribution to unitholders", "", 62222, 57964, "-6.8%", 62222, 57964, 35981, 31765],
    ["Statement of Distribution", "Distributions paid to unitholders during the year", "Note 24", -50109, -46936, "-6.3%", -50109, -46936, -24370, -21596],
    ["Statement of Distribution", "Amount available for distribution at end of year", "", 12113, 11028, "-9.0%", 12113, 11028, 11611, 10169],
    ["Statement of Distribution", "Distribution per unit (DPU in cents)", "", 2.36, 2.17, "-8.1%", 2.36, 2.17, 1.13, 0.98],
    ["Note A - Tax Adjustments", "Manager's management fees settled in units", "Note A", 4291, 4180, "-2.6%", 4291, 4180, 2099, 1926],
    ["Note A - Tax Adjustments", "Changes in fair values on investment properties, net of deferred tax", "Note A", 10984, 4478, "-59.2%", 1300, 1645, -1410, 3373],
    ["Note A - Tax Adjustments", "Adjustment on rental straight-lining", "Note A", -10227, -7446, "-27.2%", 306, 268, -3827, -2473],
    ["Note A - Tax Adjustments", "Losses on disposal of a subsidiary", "Note A", 0, 7535, "N.M.", 0, 0, 0, 0],
    ["Note A - Tax Adjustments", "(Reversal of)/impairment losses on investments in subsidiaries", "Note A", 0, 0, "N/A", -7012, 5601, 0, 0],
    ["Note A - Tax Adjustments", "Net fair value (gains)/losses of derivative financial instruments", "Note A", -440, -473, "+7.5%", -440, -473, -267, 452],
    ["Note A - Tax Adjustments", "Amount reserved for distribution to perpetual securities holders", "Note A", -1661, -1656, "-0.3%", -1661, -1656, -822, -32],
    ["Note A - Tax Adjustments", "Capital repayment from subsidiaries (Redemption of pref shares & loans)", "Note A", 0, 0, "N/A", 36065, 33405, 0, 0],
    ["Note A - Tax Adjustments", "Net settlement gains on currency hedging", "Note A", 646, 541, "-16.3%", 646, 541, 270, 0],
    ["Note A - Tax Adjustments", "Foreign exchange adjustment losses", "Note A", 3140, 4071, "+29.6%", 181, -30, 2095, 2090],
    ["Note A - Tax Adjustments", "Other non-tax deductible items and adjustments", "Note A", 5829, 5639, "-3.3%", 9719, 11119, 1797, -306],
    ["Note A - Tax Adjustments", "Total Note A Adjustments for Tax Purposes", "Note A", 12562, 16869, "+34.3%", 43395, 54600, -65, 5030],
]

# ---------------------------------------------------------
# 4. CASH FLOWS & UNITHOLDERS' FUNDS (FY2024 vs FY2025 + 1H2026)
# ---------------------------------------------------------
cf_rows = [
    ["Statement / Section", "Line Item", "Note (AR25)", "Group FY2024 (S$'000)", "Group FY2025 (S$'000)", "Group 1H2025 (S$'000)", "Group 1H2026 (S$'000)"],
    ["Operating Activities", "Total return for the year/period after income tax", "", 36754, 28982, 23933, 15707],
    ["Operating Activities", "Adj: Finance income", "Note 18", -362, -269, -153, -99],
    ["Operating Activities", "Adj: Interest expense", "Note 18", 18242, 16195, 8350, 6857],
    ["Operating Activities", "Adj: Amortisation of transaction costs", "Note 18", 4531, 4705, 2521, 2617],
    ["Operating Activities", "Adj: Depreciation expense", "Note 20", 20, 15, 8, 7],
    ["Operating Activities", "Adj: Unrealised foreign exchange losses", "", 3140, 4071, 2095, 2090],
    ["Operating Activities", "Adj: Losses/(gains) on disposal of quoted shares", "Note 20", 1, -2, -2, 0],
    ["Operating Activities", "Adj: Adjustment on rental straight-lining", "Note 4", -10227, -7446, -3827, -2473],
    ["Operating Activities", "Adj: Net fair value losses/(gains) on investment properties", "Note 4", 9578, 3528, -2632, 2473],
    ["Operating Activities", "Adj: Losses on disposal of subsidiary (PT KSS)", "Note 30", 0, 7535, 0, 0],
    ["Operating Activities", "Adj: Net fair value (gains)/losses of derivative financial instruments", "", -440, -473, -267, 452],
    ["Operating Activities", "Adj: Manager's management fees settled in units", "", 4291, 4180, 2099, 1926],
    ["Operating Activities", "Adj: Manager's divestment fees settled in units", "", 0, 0, 0, 130],
    ["Operating Activities", "Adj: Income tax expenses", "Note 22", 13822, 14331, 7364, 7535],
    ["Operating Activities", "Adj: Other non-cash items", "", -3, 3, 0, 0],
    ["Operating Activities", "Operating cash flows before working capital changes", "", 79347, 75355, 39489, 37222],
    ["Operating Activities", "Change in trade and other receivables", "", -1309, -4656, -1157, -1105],
    ["Operating Activities", "Change in other assets", "", 469, -41, -601, -1616],
    ["Operating Activities", "Change in trade and other payables", "", -156, 1791, -963, 2872],
    ["Operating Activities", "Change in other liabilities", "", -186, 5447, 251, -917],
    ["Operating Activities", "Cash generated from operating activities", "", 78165, 77896, 37019, 36456],
    ["Operating Activities", "Income taxes paid", "", -12392, -13221, -5944, -5960],
    ["Operating Activities", "NET CASH FLOWS FROM OPERATING ACTIVITIES", "", 65773, 64675, 31075, 30496],
    ["Investing Activities", "Capital expenditure on investment properties", "Note 4", -5859, -3746, -1206, -179],
    ["Investing Activities", "Net cash inflow on disposal of subsidiary (PT KSS / IAHCC)", "Note 30", 0, 22440, 0, 0],
    ["Investing Activities", "Interest received", "", 501, 298, 240, 183],
    ["Investing Activities", "Investments in quoted shares", "", -163, -155, -109, -117],
    ["Investing Activities", "Proceeds from disposals of quoted shares", "", 82, 154, 31, 117],
    ["Investing Activities", "NET CASH FLOWS (USED IN)/FROM INVESTING ACTIVITIES", "", -5439, 18991, -1044, 4],
    ["Financing Activities", "Distributions paid to unitholders", "Note 24", -50109, -46936, -24370, -21596],
    ["Financing Activities", "Distributions paid to perpetual securities holders", "", -1661, -1656, -835, -835],
    ["Financing Activities", "Proceeds from borrowings", "Note 11", 7000, 6000, 6000, 16000],
    ["Financing Activities", "Repayment of borrowings", "Note 11", -885, -873, -448, -406],
    ["Financing Activities", "Payment of transaction costs related to borrowings", "Note 11", -1669, -1831, -1087, -2235],
    ["Financing Activities", "Interest paid", "Note 11", -18240, -15825, -8491, -6724],
    ["Financing Activities", "Proceeds from settlement of derivative financial instruments", "", 646, 541, 270, 0],
    ["Financing Activities", "Perpetual securities redemption (8 Jan 2026)", "", 0, 0, 0, -33250],
    ["Financing Activities", "NET CASH FLOWS USED IN FINANCING ACTIVITIES", "", -64918, -60580, -28961, -49046],
    ["Cash Summary", "Net (decrease)/increase in cash and cash equivalents", "", -4584, 23086, 1070, -18546],
    ["Cash Summary", "Cash and cash equivalents at beginning of year/period", "", 40319, 34674, 34674, 56197],
    ["Cash Summary", "Effect of exchange rate fluctuations on cash held", "", -1061, -1563, -201, -814],
    ["Cash Summary", "CASH AND CASH EQUIVALENTS AT END OF YEAR/PERIOD", "Note 9", 34674, 56197, 35543, 36837],
]

# ---------------------------------------------------------
# 5. SEGMENTS, VALUATION INPUTS, BORROWINGS, LEASES, TENANTS & ARREARS
# ---------------------------------------------------------
seg_debt_rows = [
    ["Table / Category", "Item / Segment / Facility", "Sub-Item / Currency / Terms", "FY2024 (S$'000 or %)", "FY2025 (S$'000 or %)", "YoY Change / Remarks", "Reference"],
    ["Operating Segments (Note 28)", "Singapore", "Rental and other income", 4242, 4231, "-0.3% (+2% before FRS 116 straight-lining)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Rental and other income", 84577, 82946, "-1.9% (+4.5% local base rent escalation offset by IDR depreciation & IAHCC disposal)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Rental and other income", 13392, 13354, "-0.3% (Stable in JPY; slight JPY/SGD FX translation effect)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Rental and other income", 102211, 100531, "-1.6%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Singapore", "Property operating expenses", -543, -461, "-15.1% (Cash NPI Margin: 87.66% in FY24 vs 89.76% in FY25)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Property operating expenses", -876, -743, "-15.2% (Cash NPI Margin: 98.82% in FY24 vs 99.02% in FY25)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Property operating expenses", -2339, -1987, "-15.0% (Cash NPI Margin: 82.54% in FY24 vs 85.17% in FY25)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Property operating expenses", -3758, -3191, "-15.1% (Cash NPI Margin: 95.91% in FY24 vs 96.57% in FY25)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Singapore", "Net property and other income (NPI)", 3699, 3770, "+1.9%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Net property and other income (NPI)", 83701, 82203, "-1.8%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Net property and other income (NPI)", 11053, 11367, "+2.8%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Net property and other income (NPI)", 98453, 97340, "-1.1%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Singapore", "Net fair value losses on investment properties", -1300, -1645, "+26.5%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Net fair value (losses)/gains on investment properties", -5878, 481, "N.M.", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Net fair value losses on investment properties", -2400, -2364, "-1.5%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Net fair value losses on investment properties", -9578, -3528, "-63.2%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Income tax expense", -11873, -12471, "+5.0%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Income tax expense", -1949, -1860, "-4.6%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Income tax expense", -13822, -14331, "+3.7%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Singapore", "Segment Assets", 43393, 66699, "+53.7% (Higher cash at Trust from IAHCC sale proceeds)", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Indonesia", "Segment Assets", 853893, 774236, "-9.3%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Japan", "Segment Assets", 263215, 248034, "-5.8%", "Note 28, p. 175"],
    ["Operating Segments (Note 28)", "Total Group", "Total Segment Assets", 1160501, 1088969, "-6.2%", "Note 28, p. 175"],
    ["Segment Cap Rates & Yields", "Singapore (3 Nursing Homes)", "Implied Gross Cap Rate (Ex-FRS 116) | Cash NPI Cap Rate | Accounting NPI Cap Rate", "14.47% | 12.69% | 12.17%", "15.71% | 14.10% | 13.16%", "+124 bps Gross | +141 bps Cash NPI | +99 bps Acct NPI", "Derived (pp. 33-34, 134, 175)"],
    ["Segment Cap Rates & Yields", "Japan (14 Nursing Homes)", "Implied Gross Cap Rate | Cash NPI Cap Rate | Accounting NPI Cap Rate", "5.41% | 4.46% | 4.46%", "5.78% | 4.90% | 4.90%", "+37 bps Gross | +44 bps Cash & Acct NPI", "Derived (pp. 26-32, 139-141, 175)"],
    ["Segment Cap Rates & Yields", "Indonesia (14 Active Properties)", "Implied Gross Cap Rate (Ex-FRS 116) | Cash NPI Cap Rate", "8.95% | 8.84%", "9.66% | 9.56%", "+71 bps Gross | +72 bps Cash NPI (2026 Exit Cap Rate: 9.60% Gross / 9.51% NPI)", "Derived (pp. 35-41, 135-138, 175)"],
    ["Segment Cap Rates & Yields", "Indonesia (Total incl. IAHCC)", "Implied Gross Cap Rate (Ex-FRS 116) | Cash NPI Cap Rate | Accounting NPI Cap Rate", "8.85% | 8.75% | 9.97%", "9.91% | 9.81% | 10.79%", "+106 bps Gross | +106 bps Cash NPI | +82 bps Acct NPI", "Derived (pp. 133, 175)"],
    ["Segment Cap Rates & Yields", "Total Portfolio", "Implied Gross Cap Rate (Ex-FRS 116) | Cash NPI Cap Rate | Accounting NPI Cap Rate", "8.23% | 7.89% | 8.81%", "9.10% | 8.79% | 9.52%", "+87 bps Gross | +90 bps Cash NPI | +71 bps Acct NPI", "Derived (pp. 4, 125, 133, 175)"],
    ["Valuation Inputs (Note 4(ii))", "Singapore (3 Nursing Homes)", "Valuer DCF Discount Rate | Terminal Cap Rate", "8.75% | N/A (Note 1)", "8.75% | N/A (Note 1)", "Note 1: Direct discounting to end of leasehold term (land reverts to Govt; 6.3-11.6 yrs remaining)", "Note 4(ii), pp. 159-160"],
    ["Valuation Inputs (Note 4(ii))", "Indonesia - Restructured Properties (13 Hospitals, Note 2)", "Valuer DCF Discount Rate | Terminal Cap Rate (HGB vs BOT)", "Disc: 12.34%-13.00% | Term Cap: 9.25%-10.66% (BOT: N/A)", "Disc: 12.99%-13.20% | Term Cap: 9.25%-11.18% (BOT: N/A)", "Note 2 & 4: 10 HGB hospitals use 9.25%-11.18% terminal cap rate; 3 BOT hospitals (Sriwijaya, SHKP, SHBB) use N/A (fixed BOT term)", "Note 4(ii), pp. 159-160"],
    ["Valuation Inputs (Note 4(ii))", "Indonesia - Non-Restructured Properties (SHLC, HAMD, LPK, LPB, Note 3)", "Valuer DCF Discount Rate | Terminal Cap Rate (HGB vs BOT)", "Disc: 9.00%-12.34% | Term Cap: 9.50% (BOT: N/A)", "Disc: 9.30%-12.99% | Term Cap: 9.00%-9.86% (BOT: N/A)", "Note 3 & 4: HGB assets (SHLC, HAMD) use 9.00%-9.86% terminal cap rate; BOT malls (LPK, LPB) use N/A (fixed BOT term)", "Note 4(ii), pp. 159-160"],
    ["Valuation Inputs (Note 4(ii))", "Japan (14 Freehold Nursing Homes)", "Valuer DCF Discount Rate | Terminal Cap Rate", "Disc: 4.00%-4.70% | Term Cap: 4.30%-5.00%", "Disc: 4.00%-4.70% | Term Cap: 4.30%-5.00%", "Unchanged YoY across all 14 freehold nursing homes (Colliers KK, CBRE KK, C&W KK)", "Note 4(ii), p. 160"],
    ["Major Customers (Note 28)", "Top 1 Customer (PT Siloam International Hospitals Tbk)", "Indonesia Revenue (S$'000 & % of Group)", "42,930 (42.0%)", "45,484 (45.2%)", "+5.9% YoY", "Note 28, p. 176"],
    ["Major Customers (Note 28)", "Top 2 Customer (PT Lippo Karawaci Tbk)", "Indonesia Revenue (S$'000 & % of Group)", "35,842 (35.1%)", "31,964 (31.8%)", "-10.8% YoY", "Note 28, p. 176"],
    ["Related Party Transactions (Note 27)", "Property rental income from related parties", "Group Total (S$'000)", 54702, 31964, "-41.6% (Change in related-party classification of Siloam after Sight Investment acquisition in Sep 2024)", "Note 27, p. 174"],
    ["Borrowings Schedule (Note 11)", "CGIF Guaranteed Bonds (Due Apr 2027)", "SGD | 3.25% fixed | Face Value / Carrying Amount", "100,000 / 97,762", "100,000 / 98,757", "Guaranteed by ADB's CGIF", "Note 11, p. 161-162"],
    ["Borrowings Schedule (Note 11)", "Social Term Loan A & RCF (OCBC & CIMB)", "SGD | 3.39%-5.56% (2024: 5.27%-5.97%) | Face / Carrying", "240,713 / 238,632", "246,713 / 246,123", "Due May 2026; extended to May 2027 on 2 Feb / 26 Mar 2026", "Note 11, p. 161-162"],
    ["Borrowings Schedule (Note 11)", "Social Term Loan B (Shinsei Trust Bank)", "JPY 1.66B | 1.31% | Face Value / Carrying Amount", "14,608 / 14,486", "13,778 / 13,728", "Due 27 Sep 2026 (Ayase & Komaki nursing homes)", "Note 11, p. 161-162"],
    ["Borrowings Schedule (Note 11)", "Secured TMK Bonds B (JPY 2B) & Term Loan C (JPY 9.75B)", "JPY 11.75B | 1.50% | Kiraboshi Bank | Face / Carrying", "104,280 / 102,602", "97,525 / 96,241", "Due June 2030 (12 Japan nursing homes)", "Note 11, p. 161-162"],
    ["Borrowings Schedule (Note 11)", "Total Borrowings", "Total Face Value / Total Carrying Amount (S$'000)", "459,601 / 453,482", "458,016 / 454,849", "Unamortised txn costs: S$6,119k (2024) vs S$3,167k (2025)", "Note 11, p. 161-162"],
    ["Lease Receivables Maturity (Note 26)", "Within 1 year", "Undiscounted Lease Payments to be Received (S$'000)", 91161, 86782, "-4.8% (Trust: 4,495 in FY24 vs 4,585 in FY25)", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "1 to 2 years", "Undiscounted Lease Payments to be Received (S$'000)", 87647, 80801, "-7.8% (Trust: 4,585 in FY24 vs 1,576 in FY25)", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "2 to 3 years", "Undiscounted Lease Payments to be Received (S$'000)", 86598, 78416, "-9.4% (Trust: 1,576 in FY24 vs 0 in FY25)", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "3 to 4 years", "Undiscounted Lease Payments to be Received (S$'000)", 84572, 81153, "-4.0%", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "4 to 5 years", "Undiscounted Lease Payments to be Received (S$'000)", 87558, 83887, "-4.2%", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "After 5 years", "Undiscounted Lease Payments to be Received (S$'000)", 667592, 533285, "-20.1%", "Note 26, p. 173"],
    ["Lease Receivables Maturity (Note 26)", "Total Undiscounted Lease Receivables", "Group Total (S$'000)", 1105128, 944324, "-14.6% (Trust Total: 10,656 in FY24 vs 6,161 in FY25)", "Note 26, p. 173"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "1. PT Siloam International Hospitals Tbk & subsidiaries", "% of Rental Income (before FRS 116) / Implied S$'m", "41.1% (~S$37.8m)", "44.3% (~S$41.2m)", "+3.2% pts share", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "2. PT Lippo Karawaci Tbk & PT Andromeda Sakti", "% of Rental Income (before FRS 116) / Implied S$'m", "33.5% (~S$30.8m)", "30.7% (~S$28.6m)", "-2.8% pts share", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "3. Hikari Heights Varus Co., Ltd. (7 Japan Homes)", "% of Rental Income (before FRS 116) / Implied S$'m", "9.4% (~S$8.7m)", "9.3% (S$8.7m)", "-0.1% pts share", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "4. PT Metropolis Propertindo Utama & PT Bumi Sarana Sejahtera", "% of Rental Income (before FRS 116) / Implied S$'m", "6.0% (~S$5.5m)", "5.7% (~S$5.3m)", "-0.3% pts share", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "5. BEC Co., Ltd. (2 Japan Homes)", "% of Rental Income (before FRS 116) / Implied S$'m", "2.1% (~S$1.9m)", "2.1% (S$1.9m)", "Unchanged", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "6. The Lentor Residence Pte. Ltd. (1 SG Home)", "% of Rental Income (before FRS 116) / Implied S$'m", "2.0% (~S$1.86m)", "2.1% (S$1.9m)", "+0.1% pts share (+2% rent escalation)", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "7. Orchard Care Co., Ltd. (3 Japan Homes)", "% of Rental Income (before FRS 116) / Implied S$'m", "1.6% (~S$1.5m)", "1.5% (S$1.5m)", "-0.1% pts share", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "8. Precious Homes Pte. Ltd. (Bukit Merah, SG)", "% of Rental Income (before FRS 116) / Implied S$'m", "1.4% (~S$1.27m)", "1.4% (S$1.3m)", "Unchanged (+2% rent escalation)", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "9. Precious Homes Bukit Panjang Pte. Ltd. (Bukit Panjang, SG)", "% of Rental Income (before FRS 116) / Implied S$'m", "1.4% (~S$1.27m)", "1.4% (S$1.3m)", "Unchanged (+2% rent escalation)", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "10. Benesse Style Care Co., Ltd. (Bon Sejour Komaki, JP)", "% of Rental Income (before FRS 116) / Implied S$'m", "0.9% (~S$0.8m)", "0.9% (S$0.8m)", "Unchanged", "p. 25"],
    ["Tenant Mix (% of Rent ex-FRS 116)", "11. Social Welfare Research Institute Co., Ltd. (Ayase, JP)", "% of Rental Income (before FRS 116) / Implied S$'m", "0.6% (~S$0.5m)", "0.6% (S$0.5m)", "Unchanged", "p. 25"],
    ["Trade Sector Mix", "Healthcare (Hospitals & Nursing Homes)", "% of Rental Income (before FRS 116)", "89.2%", "89.4%", "+0.2% pts", "p. 25"],
    ["Trade Sector Mix", "Hospitality (Hotel Aryaduta Manado & IAHCC)", "% of Rental Income (before FRS 116)", "5.7%", "5.6%", "-0.1% pts", "p. 25"],
    ["Trade Sector Mix", "Retail (Lippo Plaza Kupang & Lippo Plaza Baubau)", "% of Rental Income (before FRS 116)", "5.1%", "5.0%", "-0.1% pts", "p. 25"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "Trade Receivables - Current (not past due)", "Group (S$'000)", 2159, 2243, "+3.9%", "Note 7 / 29, p. 180"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "Trade Receivables - Past due 31 to 60 days", "Group (S$'000)", 0, 236, "N.M.", "Note 7 / 29, p. 180"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "Trade Receivables - Past due 61 to 90 days", "Group (S$'000)", 595, 804, "+35.1%", "Note 7 / 29, p. 180"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "Trade Receivables - Past due over 90 days", "Group (S$'000)", 3518, 5547, "+57.7%", "Note 7 / 29, p. 180"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "Total Trade Receivables", "Group (S$'000)", 6272, 8830, "+40.8% (3rd parties: 8,490 in FY25 vs 5,834 in FY24; Related: 340 vs 438)", "Note 7, p. 158"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "PT MPU Rental Arrears - Lippo Plaza Kupang (Mall)", "IDR / S$'000 Equivalent", "Included in past due", "IDR 43,365,038,000 (~S$3,339k)", "Rose to IDR 52.56B (~S$3,837k) by 30 Jun 2026", "Note 7, p. 158"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "PT MPU Rental Arrears - 3 Hospitals (Sriwijaya, Purwakarta, Kupang)", "IDR / S$'000 Equivalent", "Included in past due", "IDR 45,888,728,000 (~S$3,534k)", "Rose to IDR 52.63B (~S$3,842k) by 30 Jun 2026", "Note 7, p. 158"],
    ["Trade Receivables & PT MPU Arrears (Note 7)", "PT MPU Total Rental Arrears at Year-End", "IDR / S$'000 Equivalent", "S$4.1m+ past due", "IDR 89,253,766,000 (~S$6,873k)", "Security deposit held: IDR 50.53B (~S$3,891k); IDR 20B (~S$1,540k) paid early 2026", "Note 7, p. 158"],
    ["Financial Ratios (Note 25)", "Expenses to average net assets ratio - excl. performance fees", "%", "1.08%", "1.70%", "+62 bps (driven by S$3.43m strategic review project expenses)", "Note 25, p. 172"],
    ["Financial Ratios (Note 25)", "Expenses to average net assets ratio - incl. performance fees", "%", "1.89%", "2.57%", "+68 bps", "Note 25, p. 172"],
    ["Financial Ratios (Note 25)", "Total operating expenses (CIS Code definition, incl. taxes & fees)", "S$'000 (% of NAV)", "29,184 (4.87%)", "32,020 (6.07%)", "+9.7% (+120 bps of NAV)", "Note 25, p. 172"],
    ["Capital Management (Note 29)", "Net debt to adjusted capital ratio", "%", "66.24%", "71.12%", "+4.88% pts (Trust: 87.31% in FY24 vs 91.58% in FY25)", "Note 29, p. 189"],
]

# ---------------------------------------------------------
# 6. COMPLETE ASSET-LEVEL PORTFOLIO DATABASE (ALL 32 PROPERTIES)
# ---------------------------------------------------------
asset_headers = [
    "No", "Country", "Property_Name", "City_Region", "Address", "Property_Type", "Centre_of_Excellence",
    "Acquisition_Date", "Purchase_Price_Orig_Currency", "Purchase_Price_SGD_M",
    "Land_Area_sqm", "GFA_sqm", "Land_Title", "Remaining_Land_Tenure",
    "Max_Beds_or_Rooms", "Operational_Beds_FY25", "Outpatient_Total_FY25_k", "Outpatient_BPJS_FY25_k", "Inpatient_Total_FY25_k", "Inpatient_BPJS_FY25_k",
    "Lease_Term_Structure", "Lease_Expiry_Date", "Remaining_Lease_Years_End2025",
    "Tenant_Master_Lessee", "Appointed_Valuer_2025",
    "FY2024_Carrying_Value_SGD_k", "FY2024_Carrying_Value_SGD_M", "FY2024_Pct_of_NAV",
    "FY2025_Appraised_Value_Local", "FY2025_Carrying_Value_SGD_k", "FY2025_Carrying_Value_SGD_M", "FY2025_Pct_of_NAV", "Valuation_YoY_Change_Pct",
    "FY2024_Implied_Rental_Income_ExFRS116_SGD_M", "FY2025_Reported_Rental_Income_ExFRS116_SGD_M",
    "FY2024_Implied_Gross_Cap_Rate_Pct", "FY2025_Implied_Gross_Cap_Rate_Pct", "Gross_Cap_Rate_YoY_Change_bps",
    "FY2024_Implied_Cash_NPI_Cap_Rate_Pct", "FY2025_Implied_Cash_NPI_Cap_Rate_Pct", "FY2025_Gross_Yield_on_Purchase_Cost_Pct",
    "Valuer_DCF_Discount_Rate_FY24", "Valuer_DCF_Discount_Rate_FY25",
    "Valuer_Terminal_Cap_Rate_FY24", "Valuer_Terminal_Cap_Rate_FY25",
    "Implied_2026_Exit_Cap_Rate_Pct",
    "Jun2026_Carrying_or_HFS_Status", "Agreed_2026_Exit_or_Put_Consideration_Local", "Agreed_2026_Exit_or_Put_Consideration_SGD_M", "AR2025_Pages"
]

raw_assets_data = [
    # SINGAPORE (3 Properties)
    [1, "Singapore", "Precious Homes @ Bukit Merah", "Bukit Merah, Singapore", "6 Lengkok Bahru, Singapore 159051", "Nursing Home", "Nursing & Senior Care",
     "11 Apr 2007", "SGD 11.8M", 11.8, 1984, 3593, "30-yr leasehold from 22 Apr 2002", "6.3 years",
     "248 beds", 248, "N/A", "N/A", "N/A", "N/A",
     "10+10 years (2% p.a. fixed escalation)", "10 Apr 2027", 1.3,
     "Precious Homes Pte. Ltd.", "Newmark",
     7500, 7.50, "1.25%", "SGD 6.8M", 6800, 6.80, "1.29%", "-9.3%",
     1.27, 1.30, "19.1%", "Retained (30 Jun 2026 Value: S$6,857k)", "N/A - Retained Core Asset", "N/A", "pp. 33, 134, 159"],
    [2, "Singapore", "Precious Homes @ Bukit Panjang", "Bukit Panjang, Singapore", "21 Senja Road, Singapore 677736", "Nursing Home", "Nursing & Senior Care",
     "11 Apr 2007", "SGD 11.5M", 11.5, 2000, 3563, "30-yr leasehold from 14 May 2003", "7.4 years",
     "248 beds", 248, "N/A", "N/A", "N/A", "N/A",
     "10+10 years (2% p.a. fixed escalation)", "10 Apr 2027", 1.3,
     "Precious Homes Bukit Panjang Pte. Ltd.", "Newmark",
     8100, 8.10, "1.36%", "SGD 7.6M", 7550, 7.55, "1.43%", "-6.8%",
     1.27, 1.30, "17.2%", "Retained (30 Jun 2026 Value: S$7,576k)", "N/A - Retained Core Asset", "N/A", "pp. 33, 134, 159"],
    [3, "Singapore", "The Lentor Residence", "Lentor, Singapore", "51 Lentor Avenue, Singapore 786876", "Nursing Home", "Nursing & Senior Care",
     "8 Jun 2007", "SGD 12.8M", 12.8, 2486, 4005, "99-yr leasehold from 20 Aug 1938", "11.6 years",
     "208 beds", 208, "N/A", "N/A", "N/A", "N/A",
     "10+10+10 years (2% p.a. fixed escalation)", "7 Jun 2027 (+10yr opt)", 11.4,
     "The Lentor Residence Pte. Ltd.", "Newmark",
     14800, 14.80, "2.47%", "SGD 14.3M", 14300, 14.30, "2.71%", "-3.4%",
     1.86, 1.90, "13.3%", "Retained (30 Jun 2026 Value: S$14,321k)", "N/A - Retained Core Asset", "N/A", "pp. 34, 134, 159"],

    # JAPAN (14 Properties)
    [4, "Japan", "Hikari Heights Varus Ishiyama", "Sapporo, Hokkaido", "Ishiyama 1-jo 3-3-33, Minami-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 839M", 7.0, 4413, 8747, "Freehold", "Freehold",
     "117 rooms", 117, "N/A", "N/A", "N/A", "N/A",
     "30+5 years (Capped CPI escalation every 3-5 yrs)", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     7506, 7.51, "1.25%", "JPY 840M", 6972, 6.97, "1.32%", "-7.1%",
     0.60, 0.60, "8.6%", "Retained (30 Jun 2026 Value: S$6,781k)", "N/A - Retained Core Asset", "N/A", "pp. 26, 139, 160"],
    [5, "Japan", "Hikari Heights Varus Tsukisamu-Koen", "Sapporo, Hokkaido", "Misono 9-jo 8-5-1, Toyohira-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 641M", 5.3, 2249, 4362, "Freehold", "Freehold",
     "58 rooms", 58, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     6063, 6.06, "1.01%", "JPY 681M", 5652, 5.65, "1.07%", "-6.8%",
     0.40, 0.40, "7.1%", "Retained (30 Jun 2026 Value: S$5,448k)", "N/A - Retained Core Asset", "N/A", "pp. 26, 139, 160"],
    [6, "Japan", "Hikari Heights Varus Fujino", "Sapporo, Hokkaido", "Fujino 3-jo 11-10-11, Minami-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 1,574M", 13.1, 7230, 9782, "Freehold", "Freehold",
     "139 rooms", 139, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     14784, 14.78, "2.47%", "JPY 1,640M", 13612, 13.61, "2.58%", "-7.9%",
     0.90, 0.90, "6.6%", "Retained (30 Jun 2026 Value: S$13,120k)", "N/A - Retained Core Asset", "N/A", "pp. 27, 139, 160"],
    [7, "Japan", "Hikari Heights Varus Kotoni", "Sapporo, Hokkaido", "Nijuyonken 4-jo 1-3-1, Nishi-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 6,209M", 51.7, 11033, 20756, "Freehold", "Freehold",
     "281 rooms", 281, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     58784, 58.78, "9.81%", "JPY 6,640M", 55112, 55.11, "10.45%", "-6.2%",
     2.90, 2.90, "5.3%", "Retained (30 Jun 2026 Value: S$53,120k)", "N/A - Retained Core Asset", "N/A", "pp. 27, 139, 160"],
    [8, "Japan", "Hikari Heights Varus Makomanai-Koen", "Sapporo, Hokkaido", "Makomanai Midoricho 1-1-1, Minami-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 4,475M", 37.3, 6653, 13301, "Freehold", "Freehold",
     "161 rooms", 161, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     42504, 42.50, "7.10%", "JPY 4,800M", 39840, 39.84, "7.56%", "-6.3%",
     2.10, 2.10, "5.3%", "Retained (30 Jun 2026 Value: S$38,400k)", "N/A - Retained Core Asset", "N/A", "pp. 28, 140, 160"],
    [9, "Japan", "Varus Cuore Yamanote", "Sapporo, Hokkaido", "Yamanote 6-jo 2-1-1, Nishi-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 1,007M", 8.4, 1668, 2808, "Freehold", "Freehold",
     "59 rooms", 59, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     9944, 9.94, "1.66%", "JPY 1,120M", 9296, 9.30, "1.76%", "-6.5%",
     0.50, 0.50, "5.4%", "Retained (30 Jun 2026 Value: S$8,960k)", "N/A - Retained Core Asset", "N/A", "pp. 28, 139, 160"],
    [10, "Japan", "Varus Cuore Sapporo-Kita & Annex", "Sapporo, Hokkaido", "Tonden 8-jo 9-3-7, Kita-ku, Sapporo", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 2,847M", 23.7, 5269, 7637, "Freehold", "Freehold",
     "216 rooms", 216, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Hikari Heights Varus Co., Ltd.", "Colliers KK",
     26312, 26.31, "4.39%", "JPY 2,970M", 24651, 24.65, "4.68%", "-6.3%",
     1.30, 1.30, "5.3%", "Retained (30 Jun 2026 Value: S$23,760k)", "N/A - Retained Core Asset", "N/A", "pp. 29, 140, 160"],
    [11, "Japan", "Elysion Gakuenmae", "Nara-shi, Nara", "Nakatomigaoka 1-1994-6, Nara-shi, Nara", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 1,610M", 13.4, 1898, 3790, "Freehold", "Freehold",
     "92 rooms", 92, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "BEC Co., Ltd.", "CBRE KK",
     14168, 14.17, "2.37%", "JPY 1,600M", 13280, 13.28, "2.52%", "-6.3%",
     0.70, 0.70, "5.3%", "Retained (30 Jun 2026 Value: S$12,800k)", "N/A - Retained Core Asset", "N/A", "pp. 29, 140, 160"],
    [12, "Japan", "Elysion Mamigaoka & Annex", "Katsuragi-gun, Nara", "Umami-minami 4-1-19, Koryocho, Kitakatsuragi-gun, Nara", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 2,370M", 19.8, 6997, 10259, "Freehold", "Freehold",
     "160 rooms", 160, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "BEC Co., Ltd.", "CBRE KK",
     19360, 19.36, "3.23%", "JPY 2,180M", 18094, 18.09, "3.43%", "-6.5%",
     1.20, 1.20, "6.6%", "Retained (30 Jun 2026 Value: S$17,448k)", "N/A - Retained Core Asset", "N/A", "pp. 30, 140, 160"],
    [13, "Japan", "Orchard Amanohashidate", "Miyazu, Kyoto", "Aza Mannen Koaza Akaiwa 1060-1, Miyazu, Kyoto", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 933M", 7.8, 2694, 2927, "Freehold", "Freehold",
     "60 rooms", 60, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Orchard Care Co., Ltd.", "C&W KK",
     7814, 7.81, "1.30%", "JPY 886M", 7354, 7.35, "1.39%", "-5.9%",
     0.50, 0.50, "6.8%", "Retained (30 Jun 2026 Value: S$7,093k)", "N/A - Retained Core Asset", "N/A", "pp. 30, 140, 160"],
    [14, "Japan", "Orchard Kaichi North", "Matsumoto, Nagano", "Kaichi 2-3-50, Matsumoto, Nagano", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 1,303M", 10.9, 2833, 5058, "Freehold", "Freehold",
     "79 rooms", 79, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Orchard Care Co., Ltd.", "C&W KK",
     11968, 11.97, "2.00%", "JPY 1,350M", 11205, 11.21, "2.13%", "-6.4%",
     0.70, 0.70, "6.2%", "Retained (30 Jun 2026 Value: S$10,800k)", "N/A - Retained Core Asset", "N/A", "pp. 31, 141, 160"],
    [15, "Japan", "Orchard Kaichi West", "Matsumoto, Nagano", "Kaichi 2-3-50, Matsumoto, Nagano", "Nursing Home", "Eldercare",
     "1 Mar 2022", "JPY 405M", 3.4, 797, 1561, "Freehold", "Freehold",
     "29 rooms", 29, "N/A", "N/A", "N/A", "N/A",
     "30+5 years", "24 Apr 2043", 22.3,
     "Orchard Care Co., Ltd.", "C&W KK",
     4127, 4.13, "0.69%", "JPY 469M", 3893, 3.89, "0.74%", "-5.7%",
     0.30, 0.30, "7.7%", "Retained (30 Jun 2026 Value: S$3,752k)", "N/A - Retained Core Asset", "N/A", "pp. 31, 141, 160"],
    [16, "Japan", "Loyal Residence Ayase", "Ayase, Kanagawa", "4695-1, 7-chome, Fukaya, Ayase City, Kanagawa", "Nursing Home", "Eldercare",
     "29 Sep 2022", "JPY 1,130M", 9.4, 2803, 3387, "Freehold", "Freehold",
     "80 rooms", 80, "N/A", "N/A", "N/A", "N/A",
     "30 years", "30 Sep 2043", 17.7,
     "Social Welfare Research Institute Co., Ltd.", "C&W KK",
     10208, 10.21, "1.70%", "JPY 1,160M", 9628, 9.63, "1.83%", "-5.7%",
     0.50, 0.50, "5.2%", "Retained (30 Jun 2026 Value: S$9,280k)", "N/A - Retained Core Asset", "N/A", "pp. 32, 141, 160"],
    [17, "Japan", "Medical Rehabilitation Home Bon Sejour Komaki", "Komaki, Aichi", "3-chome-1 Shiroyama, Komaki, Aichi", "Nursing Home", "Medical Rehabilitation & Eldercare",
     "27 Sep 2022", "JPY 1,450M", 12.1, 8230, 8858, "Freehold", "Freehold",
     "124 rooms", 124, "N/A", "N/A", "N/A", "N/A",
     "20+5 years", "21 May 2027 (+5yr opt)", 6.4,
     "Benesse Style Care Co., Ltd.", "C&W KK",
     14168, 14.17, "2.37%", "JPY 1,610M", 13363, 13.36, "2.54%", "-5.7%",
     0.80, 0.80, "6.0%", "Retained (30 Jun 2026 Value: S$12,880k)", "N/A - Retained Core Asset", "N/A", "pp. 32, 141, 160"],

    # INDONESIA (14 Active Properties + 1 Divested in Dec 2025)
    [18, "Indonesia", "Siloam Hospitals Yogyakarta", "Yogyakarta, Central Java", "Jalan Laksda Adi Sucipto No. 32-34, Yogyakarta", "Hospital", "Neuroscience & Cardiology",
     "22 Dec 2017", "SGD 27.0M", 27.0, 13715, 12474, "HGB", "Renewable HGB",
     "249 beds", 70, 132.00, 80.69, 5.78, 3.42,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Taruna Perkasa Megah & PT Lippo Karawaci Tbk", "Knight Frank",
     19641, 19.64, "3.28%", "IDR 230.6B", 17759, 17.76, "3.37%", "-9.6%",
     1.72, 1.70, "9.6%", "Retained / Put Option (30 Jun 2026 Value: S$16,837k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 17.90, "pp. 35, 74, 138, 159"],
    [19, "Indonesia", "Siloam Hospitals Baubau (SHBB) & Lippo Plaza Baubau (LPB)", "Baubau, Southeast Sulawesi", "Jalan Sultan Hasanuddin No. 50, 52, 54 & 58, Baubau", "Hospital & Retail Mall", "Emergency & Trauma",
     "10 Oct 2017", "SGD 28.5M", 28.5, 21874, 21934, "BOT scheme to 15 Jun 2044", "18.5 years",
     "140 beds (plus Mall)", 100, 57.39, 48.90, 8.24, 7.74,
     "SHBB: 15+15 yrs; LPB: 15+15 yrs", "SHBB: 31 Dec 2035; LPB: 9 Oct 2032", 25.0,
     "SHBB: PT Bina Bahtera Sejati & LPKR; LPB: PT Andromeda Sakti", "Newmark",
     23858, 23.86, "3.98%", "SHBB: IDR 134.0B (S$10.3M) + LPB: S$12.1M", 22416, 22.42, "4.25%", "-6.0%",
     2.53, 2.50, "11.2%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "SHBB: IDR 137.0B (S$10.41M) + LPB: S$12.0M", 22.41, "pp. 35, 74, 138, 159"],
    [20, "Indonesia", "Siloam Hospitals Labuan Bajo", "Labuan Bajo, East Nusa Tenggara", "Jalan Gabriel Gampur, Labuan Bajo", "Hospital", "Emergency Medicine, Internal Medicine & Neuroscience",
     "30 Dec 2016", "SGD 20.0M", 20.0, 2837, 7604, "HGB", "Renewable HGB",
     "124 beds", 90, 55.04, 48.06, 6.51, 6.16,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Lintas Buana Jaya & PT Lippo Karawaci Tbk", "Newmark",
     10542, 10.54, "1.76%", "IDR 125.8B", 9686, 9.69, "1.84%", "-8.1%",
     0.91, 0.90, "9.3%", "Retained / Put Option (30 Jun 2026 Value: S$9,183k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 9.80, "pp. 36, 74, 138, 159"],
    [21, "Indonesia", "Siloam Hospitals Kupang (SHKP) & Lippo Plaza Kupang (LPK)", "Kupang, East Nusa Tenggara", "Jalan Veteran No. 4, Arena Pameran Fatululi, Kupang", "Hospital & Retail Mall", "Emergency & Trauma, O&G & Paediatrics",
     "14 Dec 2015", "SGD 75.0M", 75.0, 66060, 55368, "BOT scheme to 21 May 2041", "16.4 years",
     "416 beds (plus Mall)", 180, 156.21, 136.72, 12.30, 11.27,
     "SHKP: 15+15 yrs; LPK: 15+15 yrs", "SHKP: 31 Dec 2035; LPK: 13 Dec 2030", 25.0,
     "SHKP: PT Krisolis Jaya Mandiri & PT MPU; LPK: PT Bumi Sarana Sejahtera", "Newmark",
     49843, 49.84, "8.32%", "SHKP: IDR 223.4B (S$17.2M) + LPK: S$29.7M", 46904, 46.90, "8.90%", "-5.9%",
     5.56, 5.50, "11.7%", "Held for Sale / Prepaid Lease (Approved 23 Jun 2026 EGM)", "SHKP: IDR 228.5B (S$17.37M) + LPK: S$29.1M", 46.47, "pp. 36, 74, 137-138, 159"],
    [22, "Indonesia", "Siloam Sriwijaya", "Palembang, South Sumatra", "Jalan POM IX, Komplek Palembang Square, Palembang", "Hospital (Strata)", "Emergency & Trauma, Gastroenterology",
     "29 Dec 2014", "SGD 39.2M", 39.2, 23311, 15336, "Strata Title on BOT scheme to 24 Jan 2041", "16.1 years",
     "357 beds", 161, 280.19, 56.85, 13.77, 2.84,
     "15+15 years (Achieved 8% GOR performance rent in FY25)", "31 Dec 2035", 25.0,
     "PT Rumah Sakit Siloam Hospitals Sumsel & PT MPU", "Newmark",
     22015, 22.02, "3.68%", "IDR 323.1B", 24878, 24.88, "4.72%", "+13.0%",
     2.95, 3.20, "12.9%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "IDR 330,432M", 25.11, "pp. 37, 74, 137, 159"],
    [23, "Indonesia", "Siloam Hospitals Purwakarta", "Purwakarta, West Java", "Jalan Raya Bungursari No. 1, Purwakarta", "Hospital", "Emergency & Trauma",
     "28 May 2014", "SGD 31.0M", 31.0, 7990, 9595, "HGB", "Renewable HGB",
     "255 beds", 287, 198.39, 117.55, 25.18, 20.69,
     "15+15 years (Achieved 8% GOR performance rent in FY25)", "31 Dec 2035", 25.0,
     "PT Berlian Cahaya Indah & PT MPU", "Newmark",
     22601, 22.60, "3.77%", "IDR 282.7B", 21766, 21.77, "4.13%", "-3.7%",
     1.95, 2.00, "9.2%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "IDR 289,092M", 21.97, "pp. 37, 74, 137, 159"],
    [24, "Indonesia", "Siloam Hospitals Bali", "Kuta, Bali", "Jalan Sunset Road No. 818, Kuta, Badung, Bali", "Hospital", "Cardiology, Emergency & Trauma, Orthopaedics",
     "13 May 2013", "SGD 97.3M", 97.3, 9025, 20958, "HGB", "Renewable HGB",
     "281 beds", 100, 155.08, 52.95, 8.41, 2.67,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Newmark",
     64245, 64.25, "10.72%", "IDR 763.7B", 58803, 58.80, "11.15%", "-8.5%",
     5.36, 5.30, "9.0%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "IDR 781,010M", 59.36, "pp. 38, 74, 136, 159"],
    [25, "Indonesia", "Siloam Hospitals TB Simatupang", "South Jakarta", "Jalan Letjend. TB Simatupang, Jalan R.A. Kartini No. 8, Cilandak", "Hospital", "Cardiology, Emergency & Trauma, Neuroscience, Oncology",
     "22 May 2013", "SGD 93.1M", 93.1, 2489, 18605, "HGB", "Renewable HGB",
     "269 beds", 99, 115.68, 0.00, 6.19, 0.00,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Newmark",
     42802, 42.80, "7.15%", "IDR 508.2B", 39134, 39.13, "7.42%", "-8.6%",
     3.54, 3.50, "8.9%", "Retained / Put Option (30 Jun 2026 Value: S$37,101k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 39.50, "pp. 38, 74, 137, 159"],
    [26, "Indonesia", "Siloam Hospitals Manado (SHMD) & Hotel Aryaduta Manado (HAMD)", "Manado, North Sulawesi", "Jalan Sam Ratulangi No. 22 & Jalan Piere Tendean No. 1, Manado", "Hospital & Hotel", "Emergency & Trauma",
     "30 Nov 2012", "SGD 83.6M", 83.6, 5518, 36051, "HGB", "Renewable HGB",
     "238 beds & 199 rooms (437 total)", 160, 92.56, 45.80, 9.93, 6.76,
     "SHMD: 15+15 yrs (IDR); HAMD: 15+15 yrs (SGD)", "SHMD: 31 Dec 2035; HAMD: 29 Nov 2027", 25.0,
     "SHMD: Siloam & LPKR; HAMD: PT Lippo Karawaci Tbk", "Newmark",
     79365, 79.37, "13.25%", "SHMD: IDR 477.6B (S$36.8M) + HAMD: S$42.6M", 79372, 79.37, "15.05%", "+0.01%",
     6.95, 6.90, "8.7%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "SHMD: IDR 488.4B (S$37.12M) + HAMD: S$41.3M", 78.42, "pp. 39, 74, 136, 159"],
    [27, "Indonesia", "Siloam Hospitals Makassar", "Makassar, South Sulawesi", "Jalan Metro Tanjung Bunga Kav 3-5, Makassar", "Hospital", "Cardiology, Emergency & Trauma, Endocrinology",
     "30 Nov 2012", "SGD 59.3M", 59.3, 3963, 15686, "HGB", "Renewable HGB",
     "362 beds", 222, 242.47, 109.24, 14.92, 6.82,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "15 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Newmark",
     66563, 66.56, "11.11%", "IDR 791.4B", 60939, 60.94, "11.56%", "-8.4%",
     5.56, 5.50, "9.0%", "Retained / Put Option (30 Jun 2026 Value: S$57,773k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 61.50, "pp. 39, 74, 136, 159"],
    [28, "Indonesia", "Mochtar Riady Comprehensive Cancer Centre (MRCCC)", "Central Jakarta", "Jalan Garnisun Dalam No. 2-3, Semanggi, Central Jakarta", "Hospital", "Emergency & Trauma, Gastroenterology, Oncology",
     "30 Dec 2010", "SGD 170.5M", 170.5, 4145, 37933, "HGB", "Renewable HGB",
     "334 beds", 213, 178.32, 29.45, 12.77, 3.90,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Newmark",
     127803, 127.80, "21.34%", "IDR 1,519.2B", 116976, 116.98, "22.19%", "-8.5%",
     10.61, 10.50, "9.0%", "Retained / Put Option (30 Jun 2026 Value: S$110,900k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 118.10, "pp. 40, 74, 135, 159"],
    [29, "Indonesia", "Siloam Hospitals Lippo Cikarang (SHLC)", "Bekasi, West Java", "Jalan Mohammad Husni Thamrin Kav. 105, Lippo Cikarang, Bekasi", "Hospital", "Emergency & Trauma, Internal Medicine, Urology",
     "31 Dec 2010", "SGD 35.0M", 35.0, 9900, 13256, "HGB", "Renewable HGB",
     "164 beds", 120, 146.27, 35.80, 7.71, 1.19,
     "15+0.5 yrs (SGD lease renewed 13 Mar 2026 to 31 Dec 2026 + 6m opt; +17% rent in FY25 via variable rent)", "31 Dec 2026 (+6m opt)", 0.5,
     "PT East Jakarta Medika (Siloam subsidiary)", "Knight Frank",
     50065, 50.07, "8.36%", "SGD 47.6M (IDR 618.3B)", 47589, 47.59, "9.03%", "-4.9%",
     4.44, 5.20, "10.9%", "Retained / Put Option (30 Jun 2026 Value: S$45,117k)", "Put Option (Part of IDR 3,879.4B / S$294.8M)", 48.00, "pp. 40, 74, 135, 159"],
    [30, "Indonesia", "Siloam Hospitals Lippo Village", "Tangerang, Banten", "Jalan Siloam No. 6, Lippo Karawaci 1600, Tangerang, Banten", "Hospital", "Cardiology, Emergency & Trauma, Neuroscience, Orthopaedics",
     "11 Dec 2006", "SGD 94.3M", 94.3, 17442, 32696, "HGB", "Renewable HGB",
     "308 beds", 274, 263.43, 0.28, 17.97, 0.74,
     "15+15 years (4.5% base escalation or 8% GOR in IDR)", "31 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Knight Frank",
     162487, 162.49, "27.13%", "IDR 1,924.8B", 148209, 148.21, "28.11%", "-8.8%",
     13.85, 13.70, "9.2%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "IDR 1,968,489M", 149.61, "pp. 41, 74, 135, 159"],
    [31, "Indonesia", "Siloam Hospitals Kebon Jeruk", "West Jakarta", "Jalan Raya Perjuangan Kav. 8, Kebon Jeruk, West Jakarta", "Hospital", "Cardiology, Emergency & Trauma, Orthopaedics, Urology",
     "11 Dec 2006", "SGD 50.6M", 50.6, 11420, 20268, "HGB", "Renewable HGB",
     "250 beds", 185, 186.20, 0.00, 12.62, 0.00,
     "15+15 years (Achieved 8% GOR performance rent in FY25)", "15 Dec 2035", 25.0,
     "PT Siloam International Hospitals Tbk & PT Lippo Karawaci Tbk", "Knight Frank",
     70345, 70.35, "11.74%", "IDR 877.7B", 67586, 67.59, "12.82%", "-3.9%",
     6.75, 7.20, "10.7%", "Held for Sale (Approved at 23 Jun 2026 EGM)", "IDR 897,664M", 68.22, "pp. 41, 74, 135, 159"],
    [32, "Indonesia", "Imperial Aryaduta Hotel & Country Club (IAHCC) [Divested 4 Dec 2025]", "Tangerang, Banten", "Jalan Boulevard Jend. Sudirman, Lippo Village 1300, Tangerang", "Hotel & Country Club", "Hospitality & Country Club",
     "11 Dec 2006", "SGD 21.2M", 21.2, 54410, 17926, "HGB", "Divested 4 Dec 2025",
     "193 rooms", "Divested", "N/A", "N/A", "N/A", "N/A",
     "Rolling 1-yr IDR leases prior to divestment on 4 Dec 2025", "Divested 4 Dec 2025", 0.0,
     "PT Lippo Karawaci Tbk (prior to disposal of PT Karya Sentra Sejahtera)", "N/A (Divested)",
     27723, 27.72, "4.63%", "Divested (Carrying value S$25,743k at disposal)", 0, 0.00, "0.00%", "-100.0% (Divested)",
     1.67, 1.90, "N/A", "Divested 4 Dec 2025 (Gross S$25.91M / Net Cash S$22.44M)", "IDR 332.2B Gross", 25.91, "pp. 19, 138, 159, 190 (Note 30)"],

    # PORTFOLIO TOTALS / SUBTOTALS
    ["-", "Subtotal", "Singapore Portfolio (3 Nursing Homes)", "Singapore", "-", "Nursing Homes", "-",
     "-", "SGD 36.1M", 36.1, 6470, 11161, "-", "-",
     "704 beds", 704, "N/A", "N/A", "N/A", "N/A",
     "-", "-", "-", "-", "Newmark",
     30400, 30.40, "5.08%", "SGD 28.65M", 28650, 28.65, "5.43%", "-5.8%",
     4.40, 4.50, "15.7%", "30 Jun 2026 Value: S$28,754k (Segment Rent w/ FRS 116: FY24 S$4,242k, FY25 S$4,231k)", "-", "-", "pp. 33-34, 133-134, 159"],
    ["-", "Subtotal", "Japan Portfolio (14 Nursing Homes)", "Japan", "-", "Nursing Homes", "-",
     "-", "JPY 26,793M", 223.3, 64767, 103234, "Freehold", "Freehold",
     "1,655 rooms", 1655, "N/A", "N/A", "N/A", "N/A",
     "-", "-", "-", "-", "Colliers / CBRE / C&W",
     247710, 247.71, "41.35%", "JPY 27,946M", 231952, 231.95, "44.00%", "-6.4%",
     13.40, 13.40, "5.8%", "30 Jun 2026 Value: S$223,642k (Segment Rent: FY24 S$13,392k, FY25 S$13,354k)", "-", "-", "pp. 26-32, 133, 139-141, 160"],
    ["-", "Subtotal", "Indonesia Active Portfolio (14 Properties)", "Indonesia", "-", "Hospitals, Malls & Hotel", "-",
     "-", "SGD 874.4M", 874.4, 221700, 317764, "-", "-",
     "3,747 beds & 199 rooms (3,946 total)", 2261, 2259.22, 762.28, 162.29, 74.18,
     "-", "-", "-", "-", "Newmark / Knight Frank",
     812175, 812.18, "135.59%", "IDR 8,185.2B + Non-Restructured SGD", 762017, 762.02, "144.54%", "-6.2%",
     72.68, 73.60, "9.7%", "30 Jun 2026: S$276,911k Continuing + S$449,906k Held for Sale", "Stage 1: S$471.5M + Stage 2 Put: S$294.8M", 766.30, "pp. 35-41, 74, 133, 135-138, 159"],
    ["-", "Subtotal", "Indonesia Total Portfolio (Incl. IAHCC in FY24/FY25)", "Indonesia", "-", "All Indonesia Properties", "-",
     "-", "SGD 895.6M", 895.6, 276110, 335690, "-", "-",
     "4,139 beds/rooms (3,946 active)", 2261, 2259.22, 762.28, 162.29, 74.18,
     "-", "-", "-", "-", "-",
     839898, 839.90, "140.22%", "-", 762017, 762.02, "144.54%", "-9.3%",
     74.35, 75.50, "9.9%", "Segment Rent w/ FRS 116: FY24 S$84,577k, FY25 S$82,946k (Ex-FRS 116: FY24 S$74.35M, FY25 S$75.50M)", "-", "-", "pp. 4, 133, 159, 175"],
    ["-", "GRAND TOTAL", "Total First REIT Portfolio (31 Active Properties at End-2025; 32 at End-2024)", "All Countries", "-", "Full Portfolio", "-",
     "-", "SGD 1,155.0M", 1155.0, 347347, 432159, "-", "-",
     "6,305 active (6,498 incl. IAHCC)", 4620, 2259.22, 762.28, 162.29, 74.18,
     "WALE 10.0 yrs", "-", 10.0, "11 Tenants", "-",
     1118008, 1118.01, "186.65%", "-", 1022619, 1022.62, "193.97%", "-8.5%",
     91.98, 93.09, "9.1%", "Total Segment Rent w/ FRS 116: FY24 S$102,211k, FY25 S$100,531k", "-", "-", "pp. 2, 4, 24-42, 133-143, 159-160"],
]

# Segment Cash NPI margins (Ex-FRS 116) from Note 28:
# Singapore: FY24 = (4400 - 543) / 4400 = 87.659%; FY25 = (4500 - 461) / 4500 = 89.756%
# Japan: FY24 = (13392 - 2339) / 13392 = 82.534%; FY25 = (13354 - 1987) / 13354 = 85.121%
# Indonesia: FY24 = (74350 - 876) / 74350 = 98.822%; FY25 = (75500 - 743) / 75500 = 99.016%
# Total Portfolio: FY24 = (91984 - 3758) / 91984 = 95.915%; FY25 = (93085 - 3191) / 93085 = 96.572%
npi_margins_fy24 = {"Singapore": 0.87659, "Japan": 0.82534, "Indonesia": 0.98822}
npi_margins_fy25 = {"Singapore": 0.89756, "Japan": 0.85121, "Indonesia": 0.99016}

assets_data = []
for r in raw_assets_data:
    no = r[0]
    country = r[1]
    prop_name = r[2]
    purchase_sgd_m = r[9]
    val24_k = r[25]
    val25_k = r[29]
    rent24_m = r[33]
    rent25_m = r[34]
    exit_sgd_m = r[38]

    # Gross Cap Rates (Ex-FRS 116 Rent / Carrying Value)
    gross_cap_24 = (rent24_m * 1000.0 / val24_k) * 100.0 if val24_k and val24_k > 0 else None
    if no == 32:
        # IAHCC divested 4 Dec 2025 at carrying value S$25,743k
        gross_cap_25 = None
        cap_yoy_bps = "Divested 4 Dec 2025 (7.38% on S$25.74M carrying value)"
    else:
        gross_cap_25 = (rent25_m * 1000.0 / val25_k) * 100.0 if val25_k and val25_k > 0 else None
        cap_yoy_bps = f"{round((gross_cap_25 - gross_cap_24) * 100):+d} bps" if (gross_cap_24 is not None and gross_cap_25 is not None) else "N/A"

    # Cash NPI Cap Rates
    if country in npi_margins_fy24:
        m24 = npi_margins_fy24[country]
        m25 = npi_margins_fy25[country]
    elif "Singapore" in prop_name:
        m24, m25 = npi_margins_fy24["Singapore"], npi_margins_fy25["Singapore"]
    elif "Japan" in prop_name:
        m24, m25 = npi_margins_fy24["Japan"], npi_margins_fy25["Japan"]
    elif "Indonesia" in prop_name:
        m24, m25 = npi_margins_fy24["Indonesia"], npi_margins_fy25["Indonesia"]
    else:
        m24, m25 = 0.95915, 0.96572

    npi_cap_24 = gross_cap_24 * m24 if gross_cap_24 is not None else None
    npi_cap_25 = gross_cap_25 * m25 if gross_cap_25 is not None else None

    # Gross Yield on Purchase Cost
    if isinstance(purchase_sgd_m, (int, float)) and purchase_sgd_m > 0:
        yield_on_cost_25 = f"{(rent25_m / purchase_sgd_m) * 100.0:.2f}%"
    else:
        yield_on_cost_25 = "N/A"

    # Valuer DCF Discount Rate & Terminal Cap Rate (Note 4(ii), AR2025 pp. 159-160)
    if country == "Singapore" or "Singapore" in prop_name:
        disc_24 = "8.75%"
        disc_25 = "8.75%"
        term_24 = "N/A (Leasehold reverts to Govt)"
        term_25 = "N/A (Leasehold reverts to Govt)"
    elif country == "Japan" or "Japan" in prop_name:
        disc_24 = "4.00% - 4.70%"
        disc_25 = "4.00% - 4.70%"
        term_24 = "4.30% - 5.00%"
        term_25 = "4.30% - 5.00%"
    elif no in [18, 20, 23, 24, 25, 27, 28, 30, 31]:
        # Restructured HGB Indonesia Hospitals
        disc_24 = "12.34% - 13.00%"
        disc_25 = "12.99% - 13.20%"
        term_24 = "9.25% - 10.66%"
        term_25 = "9.25% - 11.18%"
    elif no == 22:
        # Siloam Sriwijaya (Restructured BOT Hospital)
        disc_24 = "12.34% - 13.00%"
        disc_25 = "12.99% - 13.20%"
        term_24 = "N/A (Fixed BOT term to Jan 2041)"
        term_25 = "N/A (Fixed BOT term to Jan 2041)"
    elif no in [19, 21]:
        # Integrated Restructured Hospital + Non-Restructured Mall on BOT title (Baubau & Kupang)
        disc_24 = "Hosp: 12.34%-13.00%; Mall: 9.00%-12.34%"
        disc_25 = "Hosp: 12.99%-13.20%; Mall: 9.30%-12.99%"
        term_24 = "N/A (Fixed BOT term)"
        term_25 = "N/A (Fixed BOT term)"
    elif no == 26:
        # Integrated SHMD (Restructured) + HAMD (Non-Restructured) on HGB title
        disc_24 = "SHMD: 12.34%-13.00%; HAMD: 9.00%-12.34%"
        disc_25 = "SHMD: 12.99%-13.20%; HAMD: 9.30%-12.99%"
        term_24 = "SHMD: 9.25%-10.66%; HAMD: 9.50%"
        term_25 = "SHMD: 9.25%-11.18%; HAMD: 9.00%-9.86%"
    elif no == 29:
        # SHLC (Non-Restructured HGB Hospital)
        disc_24 = "9.00% - 12.34%"
        disc_25 = "9.30% - 12.99%"
        term_24 = "9.50%"
        term_25 = "9.00% - 9.86%"
    elif no == 32:
        # IAHCC (Divested 4 Dec 2025)
        disc_24 = "Direct Comparison (S$510 psm)"
        disc_25 = "N/A (Divested 4 Dec 2025)"
        term_24 = "N/A (Direct Comparison)"
        term_25 = "N/A (Divested 4 Dec 2025)"
    elif "Indonesia" in prop_name:
        disc_24 = "9.00% - 13.00%"
        disc_25 = "9.30% - 13.20%"
        term_24 = "9.25% - 10.66% (HGB) / N/A (BOT)"
        term_25 = "9.00% - 11.18% (HGB) / N/A (BOT)"
    else:
        disc_24 = "4.00% - 13.00%"
        disc_25 = "4.00% - 13.20%"
        term_24 = "4.30% - 10.66%"
        term_25 = "4.30% - 11.18%"

    # Implied 2026 Exit Cap Rate
    if isinstance(exit_sgd_m, (int, float)) and exit_sgd_m > 0:
        exit_cap_str = f"{(rent25_m / exit_sgd_m) * 100.0:.2f}%"
    else:
        exit_cap_str = "N/A (Retained)"

    gross_cap_24_str = f"{gross_cap_24:.2f}%" if gross_cap_24 is not None else "N/A"
    gross_cap_25_str = f"{gross_cap_25:.2f}%" if gross_cap_25 is not None else "N/A (Divested)"
    npi_cap_24_str = f"{npi_cap_24:.2f}%" if npi_cap_24 is not None else "N/A"
    npi_cap_25_str = f"{npi_cap_25:.2f}%" if npi_cap_25 is not None else "N/A (Divested)"

    new_row = (
        r[:35]
        + [
            gross_cap_24_str,
            gross_cap_25_str,
            cap_yoy_bps,
            npi_cap_24_str,
            npi_cap_25_str,
            yield_on_cost_25,
            disc_24,
            disc_25,
            term_24,
            term_25,
            exit_cap_str,
        ]
        + r[36:]
    )
    assets_data.append(new_row)

# ---------------------------------------------------------
# 7. ESG & OPERATIONAL METRICS (AR pp. 73-74, 83)
# ---------------------------------------------------------
esg_rows = [
    ["Category", "Metric / Country", "Unit", "FY2023", "FY2024", "FY2025", "YoY Change (FY25 vs FY24)", "AR2025 Reference"],
    ["Energy Consumption", "Singapore Nursing Homes", "MWh", 1806.10, 1833.42, 1830.65, "-0.2%", "p. 83"],
    ["Energy Consumption", "Indonesia Properties", "MWh", 60349.10, 62229.97, 56037.00, "-10.0%", "p. 83"],
    ["Energy Consumption", "Japan Nursing Homes", "MWh", 10443.13, 10847.29, 10726.32, "-1.1%", "p. 83"],
    ["Energy Consumption", "Total Portfolio Energy Consumption", "MWh", 72598.33, 74910.68, 68593.97, "-8.4%", "p. 83"],
    ["Energy Intensity", "Singapore Energy Intensity", "MWh / sqm", 0.162, 0.164, 0.164, "0.0%", "p. 83"],
    ["Energy Intensity", "Indonesia Energy Intensity", "MWh / sqm", 0.160, 0.165, 0.176, "+6.7%", "p. 83"],
    ["Energy Intensity", "Japan Energy Intensity", "MWh / sqm", 0.101, 0.105, 0.104, "-1.0%", "p. 83"],
    ["Energy Intensity", "Total Portfolio Energy Intensity", "MWh / sqm", 0.148, 0.153, 0.159, "+3.9%", "p. 83"],
    ["GHG Emissions (Scope 3 Downstream Leased Assets)", "Singapore Emissions", "tCO2e", 753, 755, 732, "-3.0%", "p. 83"],
    ["GHG Emissions (Scope 3 Downstream Leased Assets)", "Indonesia Emissions", "tCO2e", 48175, 49514, 44377, "-10.4%", "p. 83"],
    ["GHG Emissions (Scope 3 Downstream Leased Assets)", "Japan Emissions", "tCO2e", 3868, 3981, 3918, "-1.6%", "p. 83"],
    ["GHG Emissions (Scope 3 Downstream Leased Assets)", "Total Portfolio Scope 3 Emissions", "tCO2e", 52796, 54251, 49027, "-9.6%", "p. 83"],
    ["GHG Emissions Intensity", "Total Portfolio Emissions Intensity", "tCO2e / sqm", 0.108, 0.111, 0.113, "+1.8%", "p. 83"],
    ["Water Consumption", "Singapore Nursing Homes", "megalitres (ML)", 64.96, 65.81, 65.58, "-0.3%", "p. 83"],
    ["Water Consumption", "Indonesia Properties", "megalitres (ML)", 695.01 if False else 726.31, 695.01, 608.05, "-12.5%", "p. 83"],
    ["Water Consumption", "Japan Nursing Homes", "megalitres (ML)", 109.76, 113.09, 108.04, "-4.5%", "p. 83"],
    ["Water Consumption", "Total Portfolio Water Consumption", "megalitres (ML)", 901.03, 873.92, 781.66, "-10.6%", "p. 83"],
    ["Water Intensity", "Total Portfolio Water Intensity", "m3 / sqm", 1.835, 1.780, 1.809, "+1.6%", "p. 83"],
    ["Hospital Social Access (Indonesia 14 Hospitals)", "Total Operational Beds", "beds", "-", "-", 2261, "Max Capacity: 3,747 hospital beds", "p. 73-74"],
    ["Hospital Social Access (Indonesia 14 Hospitals)", "Total Outpatient Attendances", "thousand visits", "-", "-", 2259.22, "BPJS (Universal Healthcare): 762.28k (33.7%)", "p. 73-74"],
    ["Hospital Social Access (Indonesia 14 Hospitals)", "Total Inpatient Admissions", "thousand admissions", "-", "-", 162.29, "BPJS (Universal Healthcare): 74.18k (45.7%)", "p. 73-74"],
]

# Write individual CSV files
datasets = [
    ("01_Financial_Highlights_5Yr.csv", "Financial_Highlights", highlights_rows),
    ("02_Balance_Sheet_2024_2025.csv", "Statements_of_Financial_Pos", bs_rows),
    ("03_Income_and_Distribution_2024_2025.csv", "Total_Return_and_Distribution", pl_rows),
    ("04_Cash_Flows_2024_2025.csv", "Cash_Flows", cf_rows),
    ("05_Segments_Debt_Leases_Notes_2024_2025.csv", "Segments_Debt_Leases_Arrears", seg_debt_rows),
    ("06_Asset_Level_Portfolio_2024_2025.csv", "Asset_Level_Portfolio_2024_25", [asset_headers] + assets_data),
    ("07_ESG_and_Hospital_Operations.csv", "ESG_and_Hospital_Metrics", esg_rows),
]

for filename, _, rows in datasets:
    csv_path = os.path.join(out_dir, filename)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

# Create Styled Excel Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
sub_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
bold_font = Font(name="Calibri", size=10, bold=True)
reg_font = Font(name="Calibri", size=10)
thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

for _, sheet_name, rows in datasets:
    ws = wb.create_sheet(title=sheet_name)
    for r_idx, row in enumerate(rows, start=1):
        ws.append(row)
        is_header = (r_idx == 1)
        is_total = any(isinstance(x, str) and any(k in x.upper() for k in ["TOTAL", "SUBTOTAL", "NET PROPERTY", "NET INCOME", "ANNUAL DISTRIBUTABLE"]) for x in row[:3])
        for c_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            if is_header:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif is_total:
                cell.fill = sub_fill
                cell.font = bold_font
            else:
                cell.font = reg_font
            if not is_header and isinstance(cell.value, (int, float)):
                if abs(cell.value) >= 100 and isinstance(cell.value, int):
                    cell.number_format = "#,##0"
                elif isinstance(cell.value, float):
                    cell.number_format = "#,##0.00"

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

xlsx_path = os.path.join(out_dir, "First_REIT_FY2024_FY2025_Complete_Extraction.xlsx")
wb.save(xlsx_path)
print("Successfully generated Excel and CSV datasets with full Cap Rate columns in:", out_dir)
