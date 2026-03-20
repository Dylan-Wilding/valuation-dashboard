# Valuation Dashboard

A fundamentals-driven equity screening tool that automates EPS/P-E scenario analysis and generates interactive Excel dashboards. Built to quickly stress-test valuation assumptions across large ticker universes (S&P 500, Euronext, or custom lists).

## What It Does

Given a list of tickers, the script:

1. **Fetches earnings data** (GAAP and Street/Adjusted EPS) via `yfinance`
2. **Reconstructs rolling TTM EPS** from quarterly financials with a 45-day reporting delay to avoid look-ahead bias
3. **Builds historical P/E ranges** using 3-year daily price/earnings data (5th–95th percentile)
4. **Pulls analyst consensus estimates** (current FY and next FY) with dispersion and surprise metrics
5. **Scrapes insider transactions** from OpenInsider and scores conviction using a three-pillar model
6. **Outputs a formatted Excel workbook** with per-ticker dashboards and a sortable comparison table

## Output Structure

Each run produces an `.xlsx` file with:

| Sheet | Contents |
|---|---|
| Per-ticker tabs | 7×7 scenario grids for Implied Price, Upside/Downside %, PEG Ratio, and Holden Score |
| Per-ticker tabs | Summary statistics panel (market data, earnings basis, analyst forecasts, resilience metrics, insider activity) |
| `Comparison` | Side-by-side table across all tickers: P/E, Forward P/E, PEG, Holden Score, Resilience Ratio, Conviction Score |
| `Inputs` | Raw data layer backing all formulas (enables dropdown-driven scenario switching in Excel) |

The dashboards use Excel formulas (not static values), so you can toggle between GAAP vs. Street EPS, Analyst Consensus vs. Historical CAGR growth, and Current vs. Next FY—directly in the spreadsheet.

## Insider Conviction Scoring

Insider buying data is scored on a 0–10 scale across three pillars:

| Pillar | Max Score | Logic |
|---|---|---|
| **Materiality** | 2 | Net buying as % of market cap, or absolute dollar thresholds |
| **Breadth** | 4 | Number of unique buyers (capped at 4) |
| **Depth** | 4 | Average stake increase %, with a single-buyer cap to prevent false signals |

Thresholds are calibrated against existing academic literature on insider trading predictability and are defined as named constants at the top of the script for easy adjustment.

## Usage

```bash
# Default: runs against the ticker list defined in the script
python main.py

# Custom tickers and output file
python main.py --tickers NVDA AAPL MSFT AMZN --output my_screen.xlsx
```

## Dependencies

```
pandas
numpy
xlsxwriter
yfinance
requests
beautifulsoup4
```

Install with:
```bash
pip install pandas numpy xlsxwriter yfinance requests beautifulsoup4
```

## Limitations

- Insider data is US-only (OpenInsider coverage). European/Asian tickers will return zero activity.
- P/E calculations require positive earnings. Stocks with negative EPS are handled via fallbacks but produce less meaningful outputs.
- Data quality depends on `yfinance` and may lag or be incomplete for some tickers.

## License

© 2026 Dylan H Wilding. All rights reserved.

This project is provided for informational and research purposes only and does not constitute investment advice.
