# Date: 2026-03-14
# Author: Dylan H WILDING
# LLMs used : Gemini 3.1 Pro, Claude Sonnet 4.6
# Objective: EPS/PE estimates dashboard, scenario analysis, insider buying

import pandas as pd
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import yfinance as yf
import datetime
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import argparse

# --- CONFIGURATION ---

TICKERS = ["NVDA", "AMZN", "GOOG", "META", "TEP.PA", "BNP.PA", "CRM", "MU", "TEAM", "APP", "ADBE", "WDAY", "INTU", "TXN", "NOW", "FIS", "ASML", "SNDK", "SOFI", "MRVL"]

BASE_DIR = Path(__file__).resolve().parent
FILENAME = BASE_DIR / "Dashboard.xlsx"

USER_AGENT = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

# --- MODEL TUNING CONSTANTS ---
# Insider Conviction Scoring Thresholds, based on the existing academic literature.
INSIDER_DOLLAR_LARGE = 1_000_000
INSIDER_PCT_LARGE = 0.0002
INSIDER_DOLLAR_MODERATE = 250_000
INSIDER_PCT_MODERATE = 0.00005

INSIDER_STAKE_PCT_FOR_SCORE_4 = 100
INSIDER_STAKE_PCT_FOR_SCORE_3 = 50
INSIDER_STAKE_PCT_FOR_SCORE_2 = 20
INSIDER_STAKE_PCT_FOR_SCORE_1 = 5

MIN_HISTORICAL_PE_LOW = 5.0
MIN_HISTORICAL_PE_HIGH = 10.0

ROLLING_PE_PERIOD = "2y"
GAAP_REPORTING_DELAY_DAYS = 45
MAX_VALID_PE = 300
PE_LOW_QUANTILE = 0.05
PE_HIGH_QUANTILE = 0.95
PE_LOW_HIST_FALLBACK_MULT = 0.9
CAGR_YEARS = 3
OPENINSIDER_DAYS = 180

CURRENCY_MAP = {
    'USD': '$', 'EUR': '€', 'GBP': '£',
    'JPY': '¥', 'CNY': '¥', 'INR': '₹',
    'CAD': 'C$', 'AUD': 'A$'
}

REGION_MAP = {
    'United States': 'North America', 'Canada': 'North America',
    'France': 'Europe', 'Germany': 'Europe', 'United Kingdom': 'Europe',
    'Spain': 'Europe', 'Italy': 'Europe', 'Netherlands': 'Europe',
    'Japan': 'Asia-Pac', 'China': 'Asia-Pac', 'India': 'Asia-Pac',
    'Taiwan': 'Asia-Pac', 'Australia': 'Asia-Pac'
}


# --- REGIME CLASSIFICATION ---
def classify_regime(eps_trend, pe_trend, fwd_growth):
    """Classify stock into ERG+ regime based on directional signals."""
    THRESHOLD = 0.03
    e_up = eps_trend > THRESHOLD
    p_up = pe_trend > THRESHOLD
    p_down = pe_trend < -THRESHOLD
    f_up = fwd_growth > THRESHOLD
    f_down = fwd_growth < -THRESHOLD

    if e_up:
        if p_down:
            if f_up:     return "Golden Gap", "\U0001f7e2 Strong Opportunity"
            elif f_down: return "Value Trap (Peak)", "\U0001f534 Avoid"
            else:        return "Value Trap Risk", "\U0001f7e1 Investigate"
        elif p_up:
            if f_up:     return "Growth Expansion", "\U0001f7e2 Momentum"
            elif f_down: return "Late-Cycle Excess", "\U0001f7e1 Trim"
            else:        return "Late-Cycle Excess", "\U0001f7e1 Trim"
        else:
            if f_up:     return "Confirmed Growth", "\U0001f7e2 Momentum"
            elif f_down: return "Decelerating", "\U0001f7e1 Investigate"
            else:        return "Growth Stalling", "\U0001f7e1 Investigate"
    else:
        if p_down:
            if f_up:     return "Turnaround", "\U0001f7e1 Speculative"
            elif f_down: return "Decline", "\U0001f534 Avoid"
            else:        return "Decline", "\U0001f534 Avoid"
        elif p_up:
            if f_up:     return "Recovery Expected", "\U0001f7e1 Early Entry"
            elif f_down: return "Overvalued", "\U0001f534 Avoid"
            else:        return "Stagnation", "\U0001f7e1 Investigate"
        else:
            if f_up:     return "Turnaround", "\U0001f7e1 Speculative"
            elif f_down: return "Decline", "\U0001f534 Avoid"
            else:        return "Stagnation", "\u26aa Neutral"


# --- DATA GATHERING ---
def get_insider_data(ticker):
    """
    Scrapes OpenInsider for the last 6 months.
    Returns dictionary with Net Buying, Unique Buyers count, and Average Stake Increase %.
    """
    clean_sym = ticker.split('.')[0]
    print(f"  [{clean_sym}] Fetching OpenInsider data (L6M)...")

    url = f"http://openinsider.com/screener?s={clean_sym}&o=&pl=&ph=&ll=&lh=&fd={OPENINSIDER_DAYS}&fdr=&td=&tdr=&fdlyl=&fdlyh=&daysago=&xp=1&xs=1&vl=&vh=&ocl=&och=&sic1=-1&sicl=100&sich=9999&grp=0&nfl=&nfh=&nil=&nih=&nol=&noh=&v2l=&v2h=&oc2l=&oc2h=&sortcol=0&cnt=100&page=1"

    default_res = {'net_buying': 0.0, 'unique_buyers': 0, 'avg_stake_inc': 0.0}

    try:
        response = requests.get(url, headers=USER_AGENT, timeout=10)
        if response.status_code != 200:
            print(f"  [{clean_sym}] OpenInsider request failed with status code: {response.status_code}")
            return default_res

        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find('table', {'class': 'tinytable'})

        if not table:
            print(f"  [{clean_sym}] No insider table found. Assuming 0 activity or foreign ticker.")
            return default_res

        net_buying = 0.0
        unique_buyers_set = set()
        stake_increases = []

        rows = table.find('tbody').find_all('tr')
        valid_trades = 0

        for row in rows:
            cols = row.find_all('td')
            if len(cols) < 12: continue

            insider_name = cols[4].text.strip()
            trade_type = cols[6].text.strip()
            delta_own_txt = cols[10].text.strip().replace('%', '').replace('+', '').replace(',', '')
            value_txt = cols[11].text.strip().replace('$', '').replace(',', '').replace('+', '').replace('-', '')

            try:
                val = abs(float(value_txt))
                valid_trades += 1
            except ValueError:
                val = 0.0

            if 'Purchase' in trade_type:
                net_buying += val
                unique_buyers_set.add(insider_name)

                # Parse Stake Increase (capped at 100% to prevent extreme skewing from "New" positions)
                if delta_own_txt.lower() == 'new' or '>999' in delta_own_txt:
                    stake_inc = 100.0
                else:
                    try:
                        stake_inc = float(delta_own_txt)
                        stake_inc = min(stake_inc, 100.0)  # Cap at 100%
                    except ValueError:
                        stake_inc = 0.0

                if stake_inc > 0:
                    stake_increases.append(stake_inc)

            elif 'Sale' in trade_type:
                net_buying -= val

        unique_buyers_count = len(unique_buyers_set)
        avg_stake_inc = sum(stake_increases) / len(stake_increases) if stake_increases else 0.0

        print(
            f"  [{clean_sym}] Parsed {valid_trades} trades. Net Buy: ${net_buying:,.0f} | Unique Buyers: {unique_buyers_count} | Avg Stake Inc: {avg_stake_inc:.1f}%")
        return {'net_buying': net_buying, 'unique_buyers': unique_buyers_count, 'avg_stake_inc': avg_stake_inc}

    except Exception as e:
        print(f"  [{clean_sym}] OpenInsider scraping error: {e}")
        return default_res


def get_data(ticker):
    """Fetches Basic (GAAP) and reconstructs Street (Adjusted) EPS from earnings history, plus Next FY estimates."""
    clean_ticker = ticker.split()[0]
    print(f"\n[{clean_ticker}] Starting data fetch process...")

    try:
        stock = yf.Ticker(clean_ticker)
        info = stock.info
        print(f"  [{clean_ticker}] Successfully retrieved yfinance info dictionary.")
    except Exception as e:
        print(f"  [{clean_ticker}] CRITICAL ERROR: Failed to instantiate yfinance Ticker or retrieve info. ({e})")
        return None

    company_name = info.get('longName', info.get('shortName', clean_ticker))
    market_cap = info.get('marketCap', 0.0)
    currency_code = info.get('currency', 'USD').upper()
    curr_price = info.get('currentPrice', info.get('regularMarketPrice', 0.0))
    currency_symbol = CURRENCY_MAP.get(currency_code, '$')
    sector = info.get('sector', 'Unknown')
    country = info.get('country', 'Unknown')
    region = REGION_MAP.get(country, country)

    if curr_price == 0: return None

    eps_basic_ttm = info.get('trailingEps')
    if eps_basic_ttm is None: eps_basic_ttm = info.get('dilutedEpsTrailingTwelveMonths', 0.0)

    eps_basic_prior = eps_basic_ttm
    try:
        fin = stock.financials
        if 'Basic EPS' in fin.index and len(fin.columns) >= 2:
            val = fin.loc['Basic EPS'].iloc[1]
            if not np.isnan(val): eps_basic_prior = val
    except:
        pass

    eps_street_ttm = eps_basic_ttm
    eps_street_prior = eps_basic_prior
    surprise_display_val = 0.0

    try:
        dates = stock.earnings_dates
        if dates is not None and not dates.empty:
            dates.index = dates.index.tz_localize(None)
            actual_col = next((col for col in dates.columns if 'Actual' in str(col) or 'Reported' in str(col)), None)

            if actual_col:
                now = pd.Timestamp.now()
                past = dates[dates.index < now].dropna(subset=[actual_col]).sort_index(ascending=False)

                if len(past) >= 4:
                    last_4 = past.head(4)
                    if last_4[actual_col].sum() != 0: eps_street_ttm = last_4[actual_col].sum()
                    if 'Surprise(%)' in dates.columns:
                        surp = last_4['Surprise(%)'].dropna()
                        surprise_display_val = surp.abs().mean() if (surp > 0).any() and (
                                    surp < 0).any() else surp.mean()

                if len(past) >= 8:
                    prior_4 = past.iloc[4:8]
                    if prior_4[actual_col].sum() != 0: eps_street_prior = prior_4[actual_col].sum()
    except:
        pass

    analyst_count = info.get('numberOfAnalystOpinions', 0)
    est_year_str = "Current FY"
    est_year_str_nxt = "Next FY"

    try:
        nxt_fy_ts = info.get('nextFiscalYearEnd')
        if nxt_fy_ts:
            est_year_dt = pd.to_datetime(nxt_fy_ts, unit='s')
            est_year_str = f"FY {est_year_dt.year} ({est_year_dt.strftime('%b')})"
            nxt_year_dt = est_year_dt + pd.DateOffset(years=1)
            est_year_str_nxt = f"FY {nxt_year_dt.year} ({nxt_year_dt.strftime('%b')})"
    except:
        pass

    eps_mid, eps_high, eps_low = 0, 0, 0
    eps_mid_nxt, eps_high_nxt, eps_low_nxt = 0, 0, 0
    base_eps = eps_street_ttm if eps_street_ttm > 0 else eps_basic_ttm

    # --- EPS Trend (90d): % change from 90-day-ago consensus to current ---
    eps_90d_change = 0.0
    try:
        eps_trend_df = stock.eps_trend
        if eps_trend_df is not None and not eps_trend_df.empty:
            # yfinance eps_trend: index = fiscal period ('0y', '1y'), columns = time offsets ('current', '90daysAgo', ...)
            period = '0y' if '0y' in eps_trend_df.index else ('1y' if '1y' in eps_trend_df.index else None)
            if period and 'current' in eps_trend_df.columns and '90daysAgo' in eps_trend_df.columns:
                current_est = eps_trend_df.loc[period, 'current']
                ago_90d_est = eps_trend_df.loc[period, '90daysAgo']
                if pd.notna(current_est) and pd.notna(ago_90d_est) and float(ago_90d_est) != 0:
                    eps_90d_change = (float(current_est) - float(ago_90d_est)) / abs(float(ago_90d_est))
                    print(f"  [{clean_ticker}] EPS Trend 90d: {ago_90d_est:.2f} -> {current_est:.2f} = {eps_90d_change:+.1%}")
            else:
                print(f"  [{clean_ticker}] eps_trend shape={eps_trend_df.shape} index={list(eps_trend_df.index)} cols={list(eps_trend_df.columns)}")
    except Exception as e:
        print(f"  [{clean_ticker}] EPS trend 90d fetch failed: {e}")

    try:
        est = stock.earnings_estimate
        if est is not None and '0y' in est.index:
            eps_mid = est.loc['0y', 'avg']
            eps_high = est.loc['0y', 'high']
            eps_low = est.loc['0y', 'low']
            if abs(eps_high - eps_low) < 0.001:
                eps_low, eps_high = eps_mid * 0.75, eps_mid * 1.25
        else:
            raise ValueError

        if est is not None and '1y' in est.index:
            eps_mid_nxt = est.loc['1y', 'avg']
            eps_high_nxt = est.loc['1y', 'high']
            eps_low_nxt = est.loc['1y', 'low']
            if abs(eps_high_nxt - eps_low_nxt) < 0.001:
                eps_low_nxt, eps_high_nxt = eps_mid_nxt * 0.75, eps_mid_nxt * 1.25
        else:
            eps_mid_nxt = eps_mid * 1.10
            eps_high_nxt, eps_low_nxt = eps_mid_nxt * 1.25, eps_mid_nxt * 0.90

    except:
        eps_mid = base_eps * 1.10
        eps_high, eps_low = eps_mid * 1.25, eps_mid * 0.90
        eps_mid_nxt = eps_mid * 1.10
        eps_high_nxt, eps_low_nxt = eps_mid_nxt * 1.25, eps_mid_nxt * 0.90

    pe_current = curr_price / eps_basic_ttm if (eps_basic_ttm and eps_basic_ttm != 0) else 0
    pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_current * 0.7)
    pe_high_hist = max(MIN_HISTORICAL_PE_HIGH, pe_current * 1.3)

    # --- Historical 1Y Trend variables (computed inside rolling PE block) ---
    perf_1y = 0.0
    eps_trend_1y = 0.0
    pe_trend_1y = 0.0

    # Calculate ROLLING Historical P/E Ranges safely via yfinance
    try:
        hist = stock.history(period=ROLLING_PE_PERIOD)
        if not hist.empty:
            hist.index = hist.index.tz_localize(None)
            closes = hist['Close']

            # 1) Rolling GAAP
            ttm_gaap_series = pd.Series(index=closes.index, dtype=float)
            q_fin = stock.quarterly_financials
            if q_fin is not None and 'Basic EPS' in q_fin.index:
                basic_eps = q_fin.loc['Basic EPS'].dropna().sort_index(ascending=True)
                rolling_4q_gaap = basic_eps.rolling(4).sum().dropna()
                if not rolling_4q_gaap.empty:
                    # Delay by GAAP_REPORTING_DELAY_DAYS to avoid look-ahead bias
                    gaap_avail_dates = rolling_4q_gaap.index + pd.Timedelta(days=GAAP_REPORTING_DELAY_DAYS)
                    gaap_ttm_df = pd.DataFrame({'eps': rolling_4q_gaap.values}, index=gaap_avail_dates).sort_index()
                    gaap_ttm_df = gaap_ttm_df[~gaap_ttm_df.index.duplicated(keep='last')]
                    ttm_gaap_series = gaap_ttm_df['eps'].reindex(closes.index, method='ffill')

            # 2) Rolling Street
            ttm_street_series = pd.Series(index=closes.index, dtype=float)
            dates = stock.earnings_dates
            if dates is not None and not dates.empty:
                dates = dates.copy()
                dates.index = dates.index.tz_localize(None)
                actual_col = next((col for col in dates.columns if 'Actual' in str(col) or 'Reported' in str(col)), None)
                if actual_col:
                    actuals = dates[actual_col].dropna().sort_index(ascending=True)
                    rolling_4q_street = actuals.rolling(4).sum().dropna()
                    if not rolling_4q_street.empty:
                        street_ttm_df = pd.DataFrame({'eps': rolling_4q_street.values}, index=rolling_4q_street.index).sort_index()
                        street_ttm_df = street_ttm_df[~street_ttm_df.index.duplicated(keep='last')]
                        ttm_street_series = street_ttm_df['eps'].reindex(closes.index, method='ffill')

            daily_street_pe = closes / ttm_street_series
            daily_gaap_pe = closes / ttm_gaap_series

            daily_pe = daily_street_pe if not daily_street_pe.dropna().empty else daily_gaap_pe
            daily_pe = daily_pe[(daily_pe > 0) & (daily_pe < MAX_VALID_PE)].dropna()

            if not daily_pe.empty:
                pe_low_hist = daily_pe.quantile(PE_LOW_QUANTILE)
                pe_high_hist = daily_pe.quantile(PE_HIGH_QUANTILE)

            # --- Compute Historical 1Y Trends from the rolling series ---
            one_year_ago = closes.index[-1] - pd.DateOffset(years=1)

            # 1Y Price Performance
            closes_1y = closes[closes.index >= one_year_ago]
            if len(closes_1y) >= 2:
                perf_1y = (closes_1y.iloc[-1] / closes_1y.iloc[0]) - 1

            # 1Y EPS Trend (prefer street, fall back to GAAP)
            eps_series = ttm_street_series if not ttm_street_series.dropna().empty else ttm_gaap_series
            eps_1y = eps_series[eps_series.index >= one_year_ago].dropna()
            if len(eps_1y) >= 2 and eps_1y.iloc[0] != 0:
                eps_trend_1y = (eps_1y.iloc[-1] / eps_1y.iloc[0]) - 1
            elif eps_street_prior and eps_street_prior != 0:
                eps_trend_1y = (eps_street_ttm / eps_street_prior) - 1

            # 1Y P/E Trend (from daily PE series, or derive via identity)
            pe_1y = daily_pe[daily_pe.index >= one_year_ago]
            if len(pe_1y) >= 2 and pe_1y.iloc[0] != 0:
                pe_trend_1y = (pe_1y.iloc[-1] / pe_1y.iloc[0]) - 1
            elif (1 + eps_trend_1y) != 0:
                pe_trend_1y = ((1 + perf_1y) / (1 + eps_trend_1y)) - 1

    except Exception as e:
        print(f"  [{clean_ticker}] Rolling PE calculation fallback failed: {e}")

    # Only allow fallback when PE is positive and genuinely below historical range
    if 0 < pe_current < pe_low_hist:
        pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_current * PE_LOW_HIST_FALLBACK_MULT)

    # Final safety floor: PE bounds must never fall below configured minimums
    pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_low_hist)
    pe_high_hist = max(MIN_HISTORICAL_PE_HIGH, pe_high_hist)

    cagr_3y = 0.0
    try:
        fin_annual = stock.financials
        if 'Basic EPS' in fin_annual.index:
            eps_years = fin_annual.loc['Basic EPS'].sort_index(ascending=True)
            if len(eps_years) >= (CAGR_YEARS + 1) and eps_years.iloc[-(CAGR_YEARS + 1)] > 0 and eps_years.iloc[-1] > 0:
                cagr_3y = ((eps_years.iloc[-1] / eps_years.iloc[-(CAGR_YEARS + 1)]) ** (1 / CAGR_YEARS)) - 1
    except:
        pass

    # --- Forward-Confirmed ERG+ ---
    base_eps_fwd = eps_street_ttm if eps_street_ttm > 0 else eps_basic_ttm
    if base_eps_fwd and base_eps_fwd > 0 and eps_mid and eps_mid > 0:
        implied_fwd_growth = (eps_mid - base_eps_fwd) / base_eps_fwd
    else:
        implied_fwd_growth = 0.0

    erg_raw = (eps_trend_1y - perf_1y) if eps_trend_1y > 0 else 0.0
    if eps_trend_1y > 0.01:
        fcr_raw = implied_fwd_growth / eps_trend_1y
    else:
        fcr_raw = 0.0
    fcr = max(-1.0, min(2.0, fcr_raw))
    erg_plus = erg_raw * max(fcr, 0)

    regime, regime_signal = classify_regime(eps_trend_1y, pe_trend_1y, implied_fwd_growth)

    insider_data = get_insider_data(clean_ticker)

    # --- Quality & Risk Metrics ---
    profit_margin = info.get('profitMargins', 0.0)
    if profit_margin is None: profit_margin = 0.0
    
    fcf = info.get('freeCashflow', 0.0)
    if fcf is None: fcf = 0.0
    fcf_yield = (fcf / market_cap) if market_cap and market_cap > 0 else 0.0
    
    raw_de = info.get('debtToEquity', None)
    if raw_de is not None:
        de_ratio = raw_de / 100.0  # yfinance returns percentage e.g., 85.5 for 0.85x
    else:
        de_ratio = -1.0 # Use -1 to denote missing
        
    # --- Dividend Yield & Ex-Date (Replacing ROE) ---
    div_yield = 0.0
    div_ex_date = "N/A"
    
    try:
        divs = stock.dividends
        info_div_rate = info.get('dividendRate', 0.0)
        info_ex_date = info.get('exDividendDate')
        
        # 1. Selection: Next vs Last Ex-Date
        cal = stock.calendar
        next_ex = cal.get('Ex-Dividend Date') if cal else None
        
        today = datetime.date.today()
        is_next_declared = False
        
        if next_ex and isinstance(next_ex, datetime.date) and next_ex >= today:
            div_ex_date = next_ex.strftime('%Y-%m-%d')
            is_next_declared = True
        elif info_ex_date:
            try:
                dt_ex = datetime.date.fromtimestamp(info_ex_date)
                if dt_ex >= today:
                    div_ex_date = dt_ex.strftime('%Y-%m-%d')
                    is_next_declared = True
                else:
                    div_ex_date = dt_ex.strftime('%Y-%m-%d')
            except:
                pass
        
        if div_ex_date == "N/A" and not divs.empty:
            div_ex_date = divs.index[-1].strftime('%Y-%m-%d')

        # 2. Annualization Logic (Strict Trailing/Forward Summing)
        if not divs.empty:
            # Detect Frequency (count in last 13 months)
            divs_naive = divs.copy()
            divs_naive.index = divs_naive.index.tz_localize(None)
            thirteen_mo_ago = pd.Timestamp.now() - pd.DateOffset(months=13)
            recent_divs = divs_naive[divs_naive.index >= thirteen_mo_ago]
            raw_freq = len(recent_divs)
            
            # Map to discrete categories
            if raw_freq >= 4: freq = 4
            elif raw_freq >= 2: freq = 2
            else: freq = 1
            
            # Identify Next amount if declared (via info['lastDividendValue'])
            # falling back to latest historical if no next is known
            next_amt = info.get('lastDividendValue', divs.iloc[-1])
            
            if is_next_declared:
                # Next known + (Freq - 1) most recent historical payments
                # This "swaps" out the oldest for the newest declared without multipliers
                annual_sum = next_amt + divs.tail(freq - 1).sum() if freq > 1 else next_amt
            else:
                # Grounded Reality: Sum the last F actually paid dividends
                annual_sum = divs.tail(freq).sum()
            
            div_yield = annual_sum / curr_price if curr_price > 0 else 0.0
            
            div_yield = annual_sum / curr_price if curr_price > 0 else 0.0
        elif info_div_rate:
            # Fallback if no history but Rate exists
            div_yield = info_div_rate / curr_price if curr_price > 0 else 0.0

    except Exception as e:
        print(f"  [{clean_ticker}] Dividend calculation failed: {e}")

    # --- Analyst Price Target (12M Consensus) ---
    analyst_target_mean = 0.0
    analyst_target_median = 0.0
    analyst_target_low = 0.0
    analyst_target_high = 0.0
    analyst_target_upside = 0.0
    multiple_expansion_signal = 0.0
    analyst_target_count = 0

    try:
        apt = stock.analyst_price_targets
        if apt and isinstance(apt, dict):
            analyst_target_mean = apt.get('mean', 0.0) or 0.0
            analyst_target_median = apt.get('median', 0.0) or 0.0
            analyst_target_low = apt.get('low', 0.0) or 0.0
            analyst_target_high = apt.get('high', 0.0) or 0.0
            # Use numberOfAnalystOpinions from info as the target count
            analyst_target_count = info.get('numberOfAnalystOpinions', 0) or 0
            if analyst_target_mean > 0 and curr_price > 0:
                analyst_target_upside = (analyst_target_mean / curr_price) - 1
            # Multiple Expansion Signal via multiplicative decomposition:
            # (1 + price_return) = (1 + eps_growth) × (1 + multiple_change)
            # So: multiple_change = (1 + analyst_upside) / (1 + eps_growth) - 1
            # This correctly handles high-growth stocks where additive would distort.
            if (1 + implied_fwd_growth) != 0:
                multiple_expansion_signal = (1 + analyst_target_upside) / (1 + implied_fwd_growth) - 1
            else:
                multiple_expansion_signal = analyst_target_upside
            print(f"  [{clean_ticker}] Analyst 12M Target: {currency_symbol}{analyst_target_mean:.2f} "
                  f"(Upside: {analyst_target_upside:+.1%}, Mult.Exp: {multiple_expansion_signal:+.1%}, "
                  f"Coverage: {analyst_target_count} analysts)")
        else:
            print(f"  [{clean_ticker}] No analyst price target data available.")
    except Exception as e:
        print(f"  [{clean_ticker}] Analyst price target fetch failed: {e}")

    return {
        'ticker': clean_ticker, 'company_name': company_name, 'market_cap': market_cap,
        'currency': currency_symbol, 'price': curr_price,
        'sector': sector, 'region': region,
        'est_year_str': est_year_str, 'est_year_str_nxt': est_year_str_nxt,
        'eps_basic_ttm': eps_basic_ttm, 'eps_basic_prior': eps_basic_prior,
        'eps_street_ttm': eps_street_ttm, 'eps_street_prior': eps_street_prior,
        'eps_low': eps_low, 'eps_mid': eps_mid, 'eps_high': eps_high,
        'eps_low_nxt': eps_low_nxt, 'eps_mid_nxt': eps_mid_nxt, 'eps_high_nxt': eps_high_nxt,
        'pe_current': pe_current, 'pe_low_hist': pe_low_hist, 'pe_high_hist': pe_high_hist,
        'cagr_3y': cagr_3y, 'surprise_avg': surprise_display_val, 'analysts': analyst_count,
        'insider_net': insider_data['net_buying'], 'insider_buy_count': insider_data['unique_buyers'],
        'insider_avg_stake_inc': insider_data['avg_stake_inc'],
        'perf_1y': perf_1y, 'eps_trend_1y': eps_trend_1y, 'pe_trend_1y': pe_trend_1y,
        'implied_fwd_growth': implied_fwd_growth, 'fcr': fcr, 'erg_plus': erg_plus,
        'regime': regime, 'regime_signal': regime_signal,
        'profit_margin': profit_margin, 'fcf_yield': fcf_yield, 'de_ratio': de_ratio,
        'div_yield': div_yield, 'div_ex_date': div_ex_date,
        'analyst_target_mean': analyst_target_mean, 'analyst_target_median': analyst_target_median,
        'analyst_target_low': analyst_target_low, 'analyst_target_high': analyst_target_high,
        'analyst_target_upside': analyst_target_upside,
        'multiple_expansion_signal': multiple_expansion_signal,
        'analyst_target_count': analyst_target_count,
        'eps_90d_change': eps_90d_change
    }

# (See copyright notice at top of file)

# --- FORMATTING ---
def get_formats(workbook, symbol):
    base = {'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter'}
    num_fmt = f'"{symbol}"#,##0.00'

    def add(props): return workbook.add_format({**base, **props})

    return {
        'title': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'border': 1, 'font_size': 12}),
        'pe_label': add({'bold': True, 'bg_color': '#92D050', 'border': 1}),
        'pe_val_base': add({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'num_format': '0.0x'}),
        'pe_val_mid': add({'bold': True, 'bg_color': '#A6A6A6', 'border': 1, 'num_format': '0.0x'}),
        'eps_label': add({'bold': True, 'bg_color': '#92D050', 'border': 1, 'rotation': 90}),
        'eps_val': add({'bold': True, 'bg_color': 'white', 'border': 1, 'num_format': num_fmt}),
        'eps_mid': add({'bold': True, 'bg_color': '#FFCCFF', 'border': 1, 'num_format': num_fmt}),
        'eps_low': add({'bold': True, 'bg_color': '#CCFFFF', 'border': 1, 'num_format': num_fmt}),
        'outer_zone': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0'}),
        'mid_zone': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0'}),
        'center_zone': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0'}),
        'outer_zone_pct': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'mid_zone_pct': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'center_zone_pct': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'holden_outer': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0.0%'}),
        'holden_mid': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0.0%'}),
        'holden_center': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0.0%'}),
        'stat_head': add({'bold': True, 'align': 'left', 'bottom': 1, 'bg_color': '#E7E6E6'}),
        'stat_subhead': add({'bold': True, 'align': 'left', 'italic': True, 'font_color': '#595959', 'bottom': 1}),
        'stat_label': add({'align': 'left'}),
        'stat_val': add({'bold': True, 'num_format': num_fmt, 'align': 'right'}),
        'stat_val_txt': add({'bold': True, 'align': 'right'}),
        'stat_val_mcap': add({'bold': True, 'num_format': '"$"#,##0', 'align': 'right'}),
        'stat_val_int': add({'bold': True, 'num_format': '#,##0', 'align': 'right'}),
        'stat_val_score_int': add(
            {'bold': True, 'num_format': '0', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1}),
        'stat_val_pe': add({'bold': True, 'num_format': '0.00x', 'align': 'right'}),
        'stat_val_score': add({'bold': True, 'num_format': '0.0%', 'align': 'right'}),
        'stat_val_fcr': add({'bold': True, 'num_format': '0.00', 'align': 'right'}),
        'stat_val_real': add(
            {'bold': True, 'num_format': '0.0%', 'align': 'right', 'bg_color': '#E2EFDA', 'border': 1}),
        'stat_val_blue': add(
            {'bold': True, 'num_format': num_fmt, 'align': 'right', 'bg_color': '#CCFFFF', 'border': 1}),
        'stat_val_purple': add(
            {'bold': True, 'num_format': num_fmt, 'align': 'right', 'bg_color': '#FFCCFF', 'border': 1}),
        'stat_val_pe_grey': add(
            {'bold': True, 'num_format': '0.00x', 'align': 'right', 'bg_color': '#A6A6A6', 'border': 1}),
        'est_growth_base': add({'italic': True, 'font_color': '#006100', 'num_format': '+0.0%', 'align': 'left'}),
        'est_growth_neg': add({'italic': True, 'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'left'}),
        'diag_ok': add({'bold': True, 'font_color': '#006100', 'align': 'right'}),
        'diag_warn': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'right', 'border': 1}),
        'diag_mid': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align': 'right', 'border': 1}),

        # Insider Specific Formatting
        'insider_pos': add(
            {'bold': True, 'font_color': '#006100', 'bg_color': '#C6EFCE', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),
        'insider_neg': add(
            {'bold': True, 'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),
        'insider_neutral': add(
            {'bold': True, 'font_color': '#595959', 'bg_color': '#F2F2F2', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),

        'peg_deep_green': add({'bg_color': '#006100', 'font_color': 'white', 'border': 1, 'num_format': '0.00'}),
        'peg_green': add({'bg_color': '#C6EFCE', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_yellow': add({'bg_color': '#FFEB9C', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_orange': add({'bg_color': '#FFCC99', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_red': add({'bg_color': '#FFC7CE', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_black': add({'bg_color': '#000000', 'font_color': 'white', 'border': 1, 'num_format': '0.00'}),
        'peg_nm': add({'bg_color': '#F2F2F2', 'font_color': 'gray', 'border': 1}),

        # --- Quality & Risk Tiers ---
        'tier_dkgreen': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_green': add({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_yellow': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_red': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_black': add({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        
        'tier_de_dkgreen': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_green': add({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_yellow': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_red': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_black': add({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'num_format': '0.00x', 'align': 'right', 'border': 1}),

        'resilience_good': add(
            {'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.00', 'align': 'right',
             'border': 1}),
        'resilience_bad': add(
            {'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.00', 'align': 'right',
             'border': 1}),
        'cushion_val': add({'font_color': '#006100', 'num_format': '0.0%', 'align': 'right'}),
        'risk_val': add({'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'right'}),
        'legend_bold': add({'bold': True, 'font_size': 9, 'align': 'left'}),
        'legend_norm': add({'font_size': 9, 'align': 'left', 'italic': True, 'font_color': '#595959'}),

        'input_header': add({'bold': True, 'bg_color': '#808080', 'font_color': 'white', 'border': 1}),
        'dash_label': add({'bold': True, 'bg_color': '#E7E6E6', 'align': 'left', 'border': 1}),
        'dash_input': add({'bg_color': '#FFF2CC', 'border': 1, 'align': 'left'}),

        'comp_head': add({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1}),
        'comp_ticker': add({'bold': True, 'align': 'left'}),
        'comp_txt': add({'align': 'left'}),
        'comp_mcap': add({'num_format': '"$"#,##0'}),
        'comp_num': add({'num_format': '#,##0.00'}),
        'comp_pct': add({'num_format': '0.0%'}),
        'comp_pe': add({'num_format': '0.0x'}),
        'comp_dollar': add({'num_format': '"$"#,##0'}),
        'comp_int': add({'num_format': '#,##0'})
    }


# --- ANALYTICS DASHBOARD ---
ANALYTICS_MIN_STOCKS = 50

def build_analytics_dashboard(workbook, comp_data, col_headers):
    """Build an Analytics Dashboard sheet with charts, tables, and rankings.
    Only called when universe size >= ANALYTICS_MIN_STOCKS.
    """
    df = pd.DataFrame(comp_data, columns=col_headers)
    for col in ['Market Cap', 'Price', 'Target Price', 'Implied Upside',
                'Current P/E (Adj)', 'Forward P/E', 'PEG Ratio', 'Holden Score',
                'Safety Cushion', 'Resilience Ratio', 'Insider Net L6M ($)',
                'Unique Buyers', 'Avg Stake Inc (%)', 'Conviction Score (0-10)',
                '1Y Perf', '1Y ERG Score', 'Impl. FWD Growth', 'FCR', 'ERG+',
                'Net Profit Margin', 'FCF Yield', 'Debt/Equity', 'Div. Yield']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    df['PEG Ratio'] = df['PEG Ratio'].clip(upper=10)
    df.loc[df['Debt/Equity'] < 0, 'Debt/Equity'] = np.nan

    ws = workbook.add_worksheet('Analytics')
    ws.hide_gridlines(2)
    ws.set_tab_color('#2F5496')
    n = len(df)

    # --- Dashboard formats ---
    B = {'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter'}
    def F(p): return workbook.add_format({**B, **p})

    f_title = F({'bold': True, 'font_size': 14, 'font_color': 'white',
                 'bg_color': '#1F3864', 'align': 'left'})
    f_kl = F({'font_size': 8, 'font_color': '#8B8B8B', 'bg_color': '#F2F2F2', 'top': 1})
    f_ki = F({'bold': True, 'font_size': 18, 'font_color': '#1F3864',
              'bg_color': '#F2F2F2', 'bottom': 2, 'bottom_color': '#4472C4'})
    f_kp = F({'bold': True, 'font_size': 18, 'font_color': '#1F3864',
              'num_format': '0.0%', 'bg_color': '#F2F2F2',
              'bottom': 2, 'bottom_color': '#4472C4'})
    f_kd = F({'bold': True, 'font_size': 18, 'font_color': '#1F3864',
              'num_format': '0.0', 'bg_color': '#F2F2F2',
              'bottom': 2, 'bottom_color': '#4472C4'})
    f_sec = F({'bold': True, 'font_size': 12, 'font_color': 'white',
               'bg_color': '#2F5496', 'align': 'left', 'border': 1})
    f_th = F({'bold': True, 'bg_color': '#D6DCE4', 'border': 1, 'font_size': 9})
    f_td = F({'border': 1, 'font_size': 9, 'bg_color': '#F2F2F2', 'align': 'left'})
    f_ti = F({'border': 1, 'font_size': 9, 'bg_color': '#F2F2F2', 'num_format': '0'})
    f_tp = F({'border': 1, 'font_size': 9, 'bg_color': '#F2F2F2', 'num_format': '0.0%'})
    f_tn = F({'border': 1, 'font_size': 9, 'bg_color': '#F2F2F2', 'num_format': '0.00'})
    f_tx = F({'border': 1, 'font_size': 9, 'bg_color': '#F2F2F2', 'num_format': '0.0x'})
    f_rh = F({'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
              'border': 1, 'font_size': 9})
    f_r = F({'border': 1, 'font_size': 9, 'align': 'left'})
    f_rp = F({'border': 1, 'font_size': 9, 'num_format': '0.0%'})
    f_rn = F({'border': 1, 'font_size': 9, 'num_format': '0.00'})
    f_rx = F({'border': 1, 'font_size': 9, 'num_format': '0.0x'})
    f_ri = F({'border': 1, 'font_size': 9, 'num_format': '0'})
    f_rd = F({'border': 1, 'font_size': 9, 'num_format': '"$"#,##0'})

    # === TITLE ===
    ws.merge_range(0, 0, 0, 18,
        f'  HOLDEN ANALYTICS DASHBOARD  -  {n} Stocks  -  '
        f'{datetime.date.today().strftime("%Y-%m-%d")}', f_title)

    # === KPI CARDS ===
    for i, (lbl, val, vf) in enumerate([
        ('Universe', n, f_ki),
        ('Med. Upside', df['Implied Upside'].median(), f_kp),
        ('Med. Fwd P/E', df['Forward P/E'].median(), f_kd),
        ('Med. PEG', df['PEG Ratio'].median(), f_kd),
        ('% Golden Gap', (df['Regime'] == 'Golden Gap').mean(), f_kp),
        ('% Avoid', df['Signal'].str.contains('Avoid', na=False).mean(), f_kp),
        ('Med. Conviction', df['Conviction Score (0-10)'].median(), f_kd),
    ]):
        c = i * 2
        ws.merge_range(2, c, 2, c + 1, lbl, f_kl)
        ws.merge_range(3, c, 3, c + 1, val, vf)

    # === Helper: write summary table ===
    def write_table(start_row, title, title_span, headers, data, col_fmts):
        ws.merge_range(start_row, 0, start_row, title_span, f'  {title}', f_sec)
        hr = start_row + 1
        for j, h in enumerate(headers):
            ws.write(hr, j, h, f_th)
        for i, row_vals in enumerate(data):
            for j, val in enumerate(row_vals):
                ws.write(hr + 1 + i, j, val, col_fmts[j])
        return hr + len(data)

    # === Helper: write ranking table ===
    def write_ranking(start_row, col_start, title, headers, rows, fmts):
        span = len(headers) - 1
        ws.merge_range(start_row, col_start, start_row, col_start + span, title, f_rh)
        for j, h in enumerate(headers):
            ws.write(start_row + 1, col_start + j, h, f_th)
        for i, vals in enumerate(rows):
            for j, v in enumerate(vals):
                ws.write(start_row + 2 + i, col_start + j, v, fmts[j])

    # === SECTOR ANALYSIS ===
    s = 5
    sec = df.groupby('Sector').agg(
        Count=('Ticker', 'count'), FPE=('Forward P/E', 'median'),
        PEG=('PEG Ratio', 'median'), Up=('Implied Upside', 'median'),
        Hld=('Holden Score', 'median'), Conv=('Conviction Score (0-10)', 'median'),
        Mgn=('Net Profit Margin', 'median'), FCF=('FCF Yield', 'median'),
    ).sort_values('Count', ascending=False).reset_index()
    sec_data = [[r.Sector, r.Count, r.FPE, r.PEG, r.Up, r.Hld, r.Conv, r.Mgn, r.FCF]
                for _, r in sec.iterrows()]
    se = write_table(s, 'SECTOR ANALYSIS', 8,
        ['Sector', 'Count', 'Med. Fwd P/E', 'Med. PEG', 'Med. Upside',
         'Med. Holden', 'Med. Conv.', 'Med. Margin', 'Med. FCF'],
        sec_data, [f_td, f_ti, f_tx, f_tn, f_tp, f_tp, f_ti, f_tp, f_tp])
    shr = s + 1

    c1 = workbook.add_chart({'type': 'column'})
    c1.add_series({'name': 'Fwd P/E', 'categories': ['Analytics', shr+1, 0, se, 0],
                   'values': ['Analytics', shr+1, 2, se, 2], 'fill': {'color': '#4472C4'}})
    c1.add_series({'name': 'PEG', 'categories': ['Analytics', shr+1, 0, se, 0],
                   'values': ['Analytics', shr+1, 3, se, 3], 'fill': {'color': '#ED7D31'},
                   'y2_axis': True})
    c1.set_title({'name': 'Valuation by Sector', 'name_font': {'size': 10, 'bold': True}})
    c1.set_x_axis({'num_font': {'rotation': -45, 'size': 7}})
    c1.set_y_axis({'name': 'Fwd P/E', 'num_format': '0.0x', 'name_font': {'size': 8}})
    c1.set_y2_axis({'name': 'PEG', 'num_format': '0.0', 'name_font': {'size': 8}})
    c1.set_size({'width': 560, 'height': 340})
    c1.set_legend({'position': 'bottom', 'font': {'size': 8}})
    c1.set_style(10)
    ws.insert_chart(s, 10, c1)

    c2 = workbook.add_chart({'type': 'column'})
    c2.add_series({'name': 'Upside', 'categories': ['Analytics', shr+1, 0, se, 0],
                   'values': ['Analytics', shr+1, 4, se, 4], 'fill': {'color': '#70AD47'}})
    c2.add_series({'name': 'Holden', 'categories': ['Analytics', shr+1, 0, se, 0],
                   'values': ['Analytics', shr+1, 5, se, 5], 'fill': {'color': '#FFC000'}})
    c2.set_title({'name': 'Opportunity by Sector', 'name_font': {'size': 10, 'bold': True}})
    c2.set_x_axis({'num_font': {'rotation': -45, 'size': 7}})
    c2.set_y_axis({'name': '%', 'num_format': '0%', 'name_font': {'size': 8}})
    c2.set_size({'width': 560, 'height': 340})
    c2.set_legend({'position': 'bottom', 'font': {'size': 8}})
    c2.set_style(10)
    ws.insert_chart(s, 18, c2)

    # === REGION ANALYSIS ===
    rg = se + 2
    df['_gg'] = (df['Regime'] == 'Golden Gap').astype(int)
    reg = df.groupby('Region').agg(
        Count=('Ticker', 'count'), FPE=('Forward P/E', 'median'),
        PEG=('PEG Ratio', 'median'), Up=('Implied Upside', 'median'),
        GG=('_gg', 'mean'), Conv=('Conviction Score (0-10)', 'median'),
        Mgn=('Net Profit Margin', 'median'),
    ).sort_values('Count', ascending=False).reset_index()
    reg_data = [[r.Region, r.Count, r.FPE, r.PEG, r.Up, r.GG, r.Conv, r.Mgn]
                for _, r in reg.iterrows()]
    re_end = write_table(rg, 'REGION ANALYSIS', 7,
        ['Region', 'Count', 'Med. Fwd P/E', 'Med. PEG', 'Med. Upside',
         '% Golden Gap', 'Med. Conv.', 'Med. Margin'],
        reg_data, [f_td, f_ti, f_tx, f_tn, f_tp, f_tp, f_ti, f_tp])
    rhr = rg + 1

    c3 = workbook.add_chart({'type': 'column'})
    c3.add_series({'name': 'Fwd P/E', 'categories': ['Analytics', rhr+1, 0, re_end, 0],
                   'values': ['Analytics', rhr+1, 2, re_end, 2], 'fill': {'color': '#4472C4'}})
    c3.add_series({'name': 'Upside', 'categories': ['Analytics', rhr+1, 0, re_end, 0],
                   'values': ['Analytics', rhr+1, 4, re_end, 4], 'fill': {'color': '#70AD47'},
                   'y2_axis': True})
    c3.set_title({'name': 'Region Overview', 'name_font': {'size': 10, 'bold': True}})
    c3.set_y_axis({'name': 'Fwd P/E', 'num_format': '0.0x', 'name_font': {'size': 8}})
    c3.set_y2_axis({'name': 'Upside', 'num_format': '0%', 'name_font': {'size': 8}})
    c3.set_size({'width': 560, 'height': 300})
    c3.set_legend({'position': 'bottom', 'font': {'size': 8}})
    c3.set_style(10)
    ws.insert_chart(rg, 10, c3)

    # === REGIME DISTRIBUTION ===
    rm = re_end + 2
    rgm = df.groupby('Regime').agg(
        Count=('Ticker', 'count'), Up=('Implied Upside', 'median'),
        ERG=('ERG+', 'median'),
    ).reset_index()
    rgm['Pct'] = rgm['Count'] / n
    rgm = rgm.sort_values('Count', ascending=False).reset_index(drop=True)
    rgm_data = [[r.Regime, r.Count, r.Pct, r.Up, r.ERG] for _, r in rgm.iterrows()]
    rme = write_table(rm, 'REGIME DISTRIBUTION', 4,
        ['Regime', 'Count', '% of Univ.', 'Med. Upside', 'Med. ERG+'],
        rgm_data, [f_td, f_ti, f_tp, f_tp, f_tp])
    rmr = rm + 1

    pal = ['#70AD47', '#4472C4', '#FFC000', '#ED7D31', '#FF4444', '#9B59B6',
           '#1ABC9C', '#E74C3C', '#3498DB', '#F39C12', '#2ECC71', '#E67E22']
    c4 = workbook.add_chart({'type': 'doughnut'})
    c4.add_series({
        'name': 'Regime', 'categories': ['Analytics', rmr+1, 0, rme, 0],
        'values': ['Analytics', rmr+1, 1, rme, 1],
        'data_labels': {'percentage': True, 'category': True, 'separator': '\n',
                        'font': {'size': 7}},
        'points': [{'fill': {'color': pal[i % len(pal)]}} for i in range(len(rgm))],
    })
    c4.set_title({'name': 'Regime Mix', 'name_font': {'size': 10, 'bold': True}})
    c4.set_size({'width': 520, 'height': 380})
    c4.set_legend({'position': 'left', 'font': {'size': 7}})
    ws.insert_chart(rm, 10, c4)

    # === SCATTER PLOTS ===
    sc = max(rme + 2, rm + 22)
    ws.merge_range(sc, 0, sc, 25, '  VALUE DISCOVERY - Scatter Analysis', f_sec)
    clr = n  # Comparison sheet last data row (1-indexed)
    # Column indices in the Comparison sheet (Region appended at end, no shift)
    # PEG=9, Upside=6, Conviction=16, FwdPE=8, ERG+=24
    for cfg in [
        {'name': 'PEG vs Upside', 'xc': 9, 'yc': 6, 'xt': 'PEG Ratio',
         'yt': 'Implied Upside', 'xf': '0.0', 'yf': '0%',
         'mk': 'circle', 'cl': '#4472C4', 'bc': '#2F5496',
         'xmin': 0, 'xmax': 5, 'col_off': 0},
        {'name': 'Conviction vs Upside', 'xc': 16, 'yc': 6,
         'xt': 'Conviction (0-10)', 'yt': 'Implied Upside', 'xf': '0', 'yf': '0%',
         'mk': 'diamond', 'cl': '#70AD47', 'bc': '#375623',
         'xmin': None, 'xmax': None, 'col_off': 9},
        {'name': 'Fwd P/E vs ERG+', 'xc': 8, 'yc': 24, 'xt': 'Fwd P/E',
         'yt': 'ERG+', 'xf': '0.0x', 'yf': '0%',
         'mk': 'triangle', 'cl': '#ED7D31', 'bc': '#C55A11',
         'xmin': 0, 'xmax': 60, 'col_off': 18},
    ]:
        ch = workbook.add_chart({'type': 'scatter'})
        ch.add_series({
            'name': cfg['name'],
            'categories': ['Comparison', 1, cfg['xc'], clr, cfg['xc']],
            'values': ['Comparison', 1, cfg['yc'], clr, cfg['yc']],
            'marker': {'type': cfg['mk'], 'size': 4,
                       'fill': {'color': cfg['cl']}, 'border': {'color': cfg['bc']}},
        })
        ch.set_title({'name': cfg['name'], 'name_font': {'size': 10}})
        xa = {'name': cfg['xt'], 'num_format': cfg['xf']}
        if cfg['xmin'] is not None: xa['min'] = cfg['xmin']
        if cfg['xmax'] is not None: xa['max'] = cfg['xmax']
        ch.set_x_axis(xa)
        ch.set_y_axis({'name': cfg['yt'], 'num_format': cfg['yf']})
        ch.set_size({'width': 460, 'height': 340})
        ch.set_legend({'none': True})
        ch.set_style(11)
        ws.insert_chart(sc + 1, cfg['col_off'], ch)

    # === DISTRIBUTION ANALYSIS ===
    ds = sc + 22
    ws.merge_range(ds, 0, ds, 25, '  DISTRIBUTION ANALYSIS', f_sec)
    dr = ds + 1
    for offset, col_name, bins, fill_color in [
        (0, 'Implied Upside',
         [(-999, -0.20, '<-20%'), (-0.20, 0, '-20% to 0%'), (0, 0.20, '0% to 20%'),
          (0.20, 0.50, '20% to 50%'), (0.50, 999, '>50%')], '#4472C4'),
        (10, 'Forward P/E',
         [(-999, 10, '<10x'), (10, 15, '10-15x'), (15, 20, '15-20x'),
          (20, 30, '20-30x'), (30, 50, '30-50x'), (50, 999, '>50x')], '#ED7D31'),
        (20, 'PEG Ratio',
         [(0, 0.75, '<0.75'), (0.75, 1.0, '0.75-1.0'), (1.0, 1.5, '1.0-1.5'),
          (1.5, 2.0, '1.5-2.0'), (2.0, 3.0, '2.0-3.0'), (3.0, 999, '>3.0')], '#70AD47'),
    ]:
        ws.write(dr, offset, col_name + ' Range', f_th)
        ws.write(dr, offset + 1, 'Count', f_th)
        for bi, (lo, hi, lbl) in enumerate(bins):
            cnt = int(((df[col_name] > lo) & (df[col_name] <= hi)).sum())
            ws.write(dr + 1 + bi, offset, lbl, f_td)
            ws.write(dr + 1 + bi, offset + 1, cnt, f_ti)
        de = dr + len(bins)
        ch = workbook.add_chart({'type': 'column'})
        ch.add_series({
            'name': f'{col_name} Distribution',
            'categories': ['Analytics', dr + 1, offset, de, offset],
            'values': ['Analytics', dr + 1, offset + 1, de, offset + 1],
            'fill': {'color': fill_color}, 'gap': 50,
        })
        ch.set_title({'name': f'{col_name} Distribution', 'name_font': {'size': 10}})
        ch.set_y_axis({'name': '# Stocks'})
        ch.set_size({'width': 420, 'height': 300})
        ch.set_legend({'none': True})
        ch.set_style(10)
        ws.insert_chart(ds + 1, offset + 3, ch)

    # === TOP / BOTTOM RANKINGS ===
    tk = ds + 22
    ws.merge_range(tk, 0, tk, 25, '  TOP / BOTTOM RANKINGS', f_sec)

    # Top 10 Holden Score
    t10 = df.nlargest(10, 'Holden Score')
    t10_rows = [[i+1, r['Ticker'], r['Sector'], r['Holden Score'], r['Implied Upside'],
                 r['PEG Ratio'], r['Forward P/E'], r['Conviction Score (0-10)']]
                for i, (_, r) in enumerate(t10.iterrows())]
    write_ranking(tk + 1, 0, 'Top 10 - Highest Holden Score',
        ['#', 'Ticker', 'Sector', 'Holden', 'Upside', 'PEG', 'Fwd P/E', 'Conv.'],
        t10_rows, [f_ri, f_r, f_r, f_rp, f_rp, f_rn, f_rx, f_ri])

    # Top 10 Golden Gap
    gg = df[df['Regime'] == 'Golden Gap'].nlargest(10, 'ERG+')
    gg_rows = [[i+1, r['Ticker'], r['Sector'], r['ERG+'], r['Implied Upside'],
                r['Forward P/E'], r['FCR'], r['Signal']]
               for i, (_, r) in enumerate(gg.iterrows())]
    write_ranking(tk + 1, 10, 'Top 10 - Golden Gap (by ERG+)',
        ['#', 'Ticker', 'Sector', 'ERG+', 'Upside', 'Fwd P/E', 'FCR', 'Signal'],
        gg_rows, [f_ri, f_r, f_r, f_rp, f_rp, f_rx, f_rn, f_r])

    # Bottom 10 Upside
    b10 = df.nsmallest(10, 'Implied Upside')
    b10_rows = [[i+1, r['Ticker'], r['Sector'], r['Implied Upside'],
                 r['Forward P/E'], r['PEG Ratio'], r['Regime'], r['Signal']]
                for i, (_, r) in enumerate(b10.iterrows())]
    write_ranking(tk + 14, 0, 'Bottom 10 - Most Overvalued',
        ['#', 'Ticker', 'Sector', 'Upside', 'Fwd P/E', 'PEG', 'Regime', 'Signal'],
        b10_rows, [f_ri, f_r, f_r, f_rp, f_rx, f_rn, f_r, f_r])

    # Top 10 Conviction
    cv10 = df.nlargest(10, 'Conviction Score (0-10)')
    cv_rows = [[i+1, r['Ticker'], r['Sector'], r['Conviction Score (0-10)'],
                r['Insider Net L6M ($)'], r['Unique Buyers'], r['Implied Upside'],
                r['Signal']]
               for i, (_, r) in enumerate(cv10.iterrows())]
    write_ranking(tk + 14, 10, 'Top 10 - Highest Insider Conviction',
        ['#', 'Ticker', 'Sector', 'Conv.', 'Net Buy ($)', 'Buyers', 'Upside', 'Signal'],
        cv_rows, [f_ri, f_r, f_r, f_ri, f_rd, f_ri, f_rp, f_r])

    # === REGIME x SECTOR MATRIX ===
    mx = tk + 28
    ws.merge_range(mx, 0, mx, 17, '  REGIME x SECTOR MATRIX (Stock Counts)', f_sec)
    cross = pd.crosstab(df['Sector'], df['Regime'])
    cross = cross.reindex(columns=sorted(cross.columns))
    mhr = mx + 1
    ws.write(mhr, 0, 'Sector \\ Regime', f_th)
    for j, rg_name in enumerate(cross.columns):
        ws.write(mhr, j + 1, rg_name, f_th)
    ws.write(mhr, len(cross.columns) + 1, 'Total', f_th)
    for i, (sector_name, crow) in enumerate(cross.iterrows()):
        r = mhr + 1 + i
        ws.write(r, 0, sector_name, f_td)
        for j, rg_name in enumerate(cross.columns):
            ws.write(r, j + 1, int(crow[rg_name]), f_ti)
        ws.write(r, len(cross.columns) + 1, int(crow.sum()), f_ti)
    if len(cross) > 0:
        rng = '{}:{}'.format(
            xl_rowcol_to_cell(mhr + 1, 1),
            xl_rowcol_to_cell(mhr + len(cross), len(cross.columns)))
        ws.conditional_format(rng, {
            'type': '3_color_scale',
            'min_color': '#FFFFFF', 'mid_color': '#BDD7EE', 'max_color': '#2F5496'})

    # === COLUMN WIDTHS ===
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 8, 13)
    ws.set_column(9, 9, 3)
    ws.set_column(10, 10, 22)
    ws.set_column(11, 17, 13)
    ws.set_column(18, 19, 3)
    ws.set_column(20, 21, 12)
    ws.set_column(22, 30, 10)

    print(f"  Analytics Dashboard: {n} stocks, {len(sec)} sectors, {len(reg)} regions.")


# --- MAIN LOGIC ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Holden Valuation Model")
    parser.add_argument("--tickers", nargs="+", help="List of tickers to process (e.g. ABC XYZ)")
    parser.add_argument("--output", type=str, help="Output Excel filename (e.g. custom.xlsx)")
    args = parser.parse_args()

    if args.tickers:
        TICKERS = args.tickers
    if args.output:
        FILENAME = BASE_DIR / args.output

    # Deduplicate TICKERS while preserving order
    TICKERS = list(dict.fromkeys(TICKERS))

    ALL_DATA = []

    print("\n==============================")
    print("Initiating Batch Data Sequence")
    print("==============================\n")

    for t in TICKERS:
        d = get_data(t)
        if d:
            ALL_DATA.append(d)
        else:
            print(f"[{t}] FAILED. Moving to next ticker.\n")

    if not ALL_DATA:
        print("\n[!] No data successfully fetched. Exiting without generating spreadsheet.")
        import sys
        sys.exit(0)

    print("\nGenerating Dashboards...")
    writer = pd.ExcelWriter(
        FILENAME,
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    # 1. GENERATE DASHBOARDS
    data_start_row = 10
    used_sheet_names = set()

    for i, d in enumerate(ALL_DATA):
        ticker = d['ticker']
        print(f"  Formatting dashboard for {ticker}...")

        # --- Unique and Valid Sheet Naming ---
        # Excel forbidden: \ / ? * : [ ]
        clean_name = ticker.replace('\\','').replace('/','').replace('?','').replace('*','').replace(':','').replace('[','').replace(']','')
        base_name = clean_name[:31] # Excel limit
        
        sheet_name = base_name
        counter = 1
        while sheet_name.lower() in used_sheet_names:
            suffix = f"_{counter}"
            sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
            counter += 1
        
        used_sheet_names.add(sheet_name.lower())
        
        sheet = workbook.add_worksheet(sheet_name)
        sheet.hide_gridlines(2)
        fmt = get_formats(workbook, d['currency'])

        inp_row = data_start_row + 1 + i


        def get_ref(col_idx):
            return f"Inputs!{xl_rowcol_to_cell(inp_row, col_idx, row_abs=True, col_abs=True)}"


        # Mapping References
        ref_ticker = get_ref(0)
        ref_curr = get_ref(1)
        ref_price = get_ref(2)
        ref_b_ttm = get_ref(3)
        ref_b_prior = get_ref(4)
        ref_s_ttm = get_ref(5)
        ref_s_prior = get_ref(6)
        ref_elo, ref_emid, ref_ehigh = get_ref(7), get_ref(8), get_ref(9)
        ref_pec, ref_pel, ref_peh = get_ref(10), get_ref(11), get_ref(12)
        ref_cagr, ref_sector, ref_region = get_ref(13), get_ref(14), get_ref(15)
        ref_surprise = get_ref(16)
        ref_analysts = get_ref(17)
        ref_est_year = get_ref(18)
        ref_insider = get_ref(19)
        ref_name = get_ref(20)
        ref_mcap = get_ref(21)
        ref_elo_nxt = get_ref(22)
        ref_emid_nxt = get_ref(23)
        ref_ehigh_nxt = get_ref(24)
        ref_est_year_nxt = get_ref(25)
        ref_insider_count = get_ref(26)
        ref_avg_stake = get_ref(27)
        ref_perf_1y = get_ref(28)
        ref_eps_trend_1y = get_ref(29)
        ref_pe_trend_1y = get_ref(30)
        ref_implied_fwd = get_ref(31)
        ref_fcr = get_ref(32)
        ref_erg_plus = get_ref(33)
        ref_regime = get_ref(34)
        ref_regime_signal = get_ref(35)
        
        # Quality & Risk mapping (Columns 36, 37, 38, 39, 40)
        ref_profit_margin = get_ref(36)
        ref_fcf_yield = get_ref(37)
        ref_de_ratio = get_ref(38)
        ref_div_yield = get_ref(39)
        ref_div_ex_date = get_ref(40)

        # Analyst Price Target mapping (Columns 41-48)
        ref_analyst_target_mean = get_ref(41)
        ref_analyst_target_median = get_ref(42)
        ref_analyst_target_low = get_ref(43)
        ref_analyst_target_high = get_ref(44)
        ref_analyst_target_upside = get_ref(45)
        ref_mult_expansion = get_ref(46)
        ref_analyst_target_count = get_ref(47)
        ref_eps_90d_change = get_ref(48)

        dash_row, dash_col = 2, 15
        lists = {'Growth': 'Inputs!$A$2:$A$3', 'EPS': 'Inputs!$B$2:$B$3', 'PE': 'Inputs!$C$2:$C$4',
                 'Type': 'Inputs!$D$2:$D$3', 'FY': 'Inputs!$E$2:$E$3'}

        sheet.write(dash_row, dash_col, "EPS Type", fmt['dash_label'])
        sheet.write(dash_row + 1, dash_col, "Growth Basis", fmt['dash_label'])
        sheet.write(dash_row + 2, dash_col, "EPS Basis", fmt['dash_label'])
        sheet.write(dash_row + 3, dash_col, "P/E Mode", fmt['dash_label'])
        sheet.write(dash_row + 4, dash_col, "Target FY Period", fmt['dash_label'])

        sheet.data_validation(dash_row, dash_col + 1, dash_row, dash_col + 1,
                              {'validate': 'list', 'source': lists['Type']})
        sheet.write(dash_row, dash_col + 1, 'Street (Adjusted)', fmt['dash_input'])
        sheet.data_validation(dash_row + 1, dash_col + 1, dash_row + 1, dash_col + 1,
                              {'validate': 'list', 'source': lists['Growth']})
        sheet.write(dash_row + 1, dash_col + 1, 'Analyst Consensus', fmt['dash_input'])
        sheet.data_validation(dash_row + 2, dash_col + 1, dash_row + 2, dash_col + 1,
                              {'validate': 'list', 'source': lists['EPS']})
        sheet.write(dash_row + 2, dash_col + 1, 'TTM (Current)', fmt['dash_input'])
        sheet.data_validation(dash_row + 3, dash_col + 1, dash_row + 3, dash_col + 1,
                              {'validate': 'list', 'source': lists['PE']})
        sheet.write(dash_row + 3, dash_col + 1, 'Flexible (Hist)', fmt['dash_input'])
        sheet.data_validation(dash_row + 4, dash_col + 1, dash_row + 4, dash_col + 1,
                              {'validate': 'list', 'source': lists['FY']})
        sheet.write(dash_row + 4, dash_col + 1, 'Current', fmt['dash_input'])

        cell_type = xl_rowcol_to_cell(dash_row, dash_col + 1)
        cell_growth = xl_rowcol_to_cell(dash_row + 1, dash_col + 1)
        cell_eps = xl_rowcol_to_cell(dash_row + 2, dash_col + 1)
        cell_pe = xl_rowcol_to_cell(dash_row + 3, dash_col + 1)
        cell_fy = xl_rowcol_to_cell(dash_row + 4, dash_col + 1)
        addr_type_input = xl_rowcol_to_cell(dash_row, dash_col + 1, row_abs=True, col_abs=True)

        dyn_elo = f'IF({cell_fy}="Next", {ref_elo_nxt}, {ref_elo})'
        dyn_emid = f'IF({cell_fy}="Next", {ref_emid_nxt}, {ref_emid})'
        dyn_ehigh = f'IF({cell_fy}="Next", {ref_ehigh_nxt}, {ref_ehigh})'

        f_act_ttm = f'IF({cell_type}="Basic (GAAP)", {ref_b_ttm}, {ref_s_ttm})'
        f_calc_base = f'IF({cell_eps}="TTM (Current)", {f_act_ttm}, {dyn_elo})'
        f_calc_growth = f'IF({cell_growth}="Analyst Consensus", ({dyn_emid}-{f_act_ttm})/{f_act_ttm}, {ref_cagr})'
        f_pe_rt = f'{ref_price}/{f_act_ttm}'
        f_target_mid = f'({f_act_ttm}*(1+{f_calc_growth}))'
        f_pe_rt_bounded = f'MAX({MIN_HISTORICAL_PE_LOW}, {f_pe_rt})'
        f_calc_pemid = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 30, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 20, {f_pe_rt_bounded}))'
        f_pe_down_step = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 5, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 0, ({f_calc_pemid}-MIN({f_calc_pemid}-3, {ref_pel}))/3))'
        f_pe_up_step = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 5, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 0, (MAX({f_calc_pemid}+3, {ref_peh})-{f_calc_pemid})/3))'
        f_calc_espstep = f'({f_target_mid}-{f_calc_base})/3'
        f_step_upper = f'({dyn_ehigh}-{f_target_mid})/3'
        pe_steps_offsets = [-3, -2, -1, 0, 1, 2, 3]


        def draw_grid_formulas(start_row, table_type):
            start_col = 1
            titles = {'PRICE': f'Implied Stock Price: {d["ticker"]}', 'UPSIDE': 'Implied Upside / Downside %',
                      'PEG': 'Implied PEG Ratio', 'HOLDEN': 'The Holden Score (Upside Efficiency)'}
            sheet.merge_range(start_row, start_col, start_row, start_col + 8, titles[table_type], fmt['title'])
            pe_row = start_row + 2
            sheet.merge_range(start_row + 1, start_col + 2, start_row + 1, start_col + 8, "P/E Multiple",
                              fmt['pe_label'])

            for j in range(7):
                offset = pe_steps_offsets[j]
                f_pe_cell_fmt = fmt['pe_val_mid'] if j == 3 else fmt['pe_val_base']
                if offset < 0:
                    form = f'={f_calc_pemid} + ({f_pe_down_step}*{offset})'
                elif offset > 0:
                    form = f'={f_calc_pemid} + ({f_pe_up_step}*{offset})'
                else:
                    form = f'={f_calc_pemid}'
                sheet.write_formula(pe_row, start_col + 2 + j, form, f_pe_cell_fmt)

            sheet.merge_range(start_row + 3, start_col, start_row + 9, start_col, "EPS", fmt['eps_label'])
            eps_col, eps_start_row = start_col + 1, start_row + 3

            for i in range(7):
                r = eps_start_row + i
                f_eps_cell_fmt = fmt['eps_mid'] if i == 3 else (fmt['eps_low'] if i == 0 else fmt['eps_val'])
                if i <= 3:
                    form = f'={f_calc_base} + ({f_calc_espstep}*{i})'
                else:
                    form = f'={f_target_mid} + ({f_step_upper}*{i - 3})'
                sheet.write_formula(r, eps_col, form, f_eps_cell_fmt)

            for i in range(7):
                r = eps_start_row + i
                eps_cell = xl_rowcol_to_cell(r, eps_col, col_abs=True)
                for j in range(7):
                    c = start_col + 2 + j
                    pe_cell = xl_rowcol_to_cell(pe_row, c, row_abs=True)
                    zone = 'center_zone' if abs(i - 3) == 0 and abs(j - 3) == 0 else (
                        'mid_zone' if abs(i - 3) <= 1 and abs(j - 3) <= 1 else 'outer_zone')
                    if table_type in ['UPSIDE', 'HOLDEN']: zone += '_pct'

                    if table_type == 'PRICE':
                        sheet.write_formula(r, c, f'={eps_cell}*{pe_cell}', fmt[zone])
                    elif table_type == 'UPSIDE':
                        sheet.write_formula(r, c, f'=({eps_cell}*{pe_cell} - {ref_price}) / {ref_price}', fmt[zone])
                    elif table_type == 'PEG':
                        growth_denom = f'(({eps_cell}-{f_act_ttm})/{f_act_ttm}*100)'
                        peg_f = f'=IF({growth_denom} < 0.1, "NM", {pe_cell}/{growth_denom})'
                        sheet.write_formula(r, c, peg_f, fmt['peg_nm'])
                    elif table_type == 'HOLDEN':
                        growth_denom = f'(({eps_cell}-{f_act_ttm})/{f_act_ttm}*100)'
                        upside_part = f'(({eps_cell}*{pe_cell} - {ref_price}) / {ref_price})'
                        peg_part = f'({pe_cell}/{growth_denom})'
                        form = f'=IF(OR({growth_denom}<0.1, {upside_part}<0), "NM", {upside_part}/{peg_part})'
                        sheet.write_formula(r, c, form, fmt[f'holden_{zone.split("_")[0]}'])

            if table_type == 'PEG':
                rng = f"{xl_rowcol_to_cell(start_row + 3, start_col + 2)}:{xl_rowcol_to_cell(start_row + 9, start_col + 8)}"
                for crit, val, f in [('<', 0.75, 'peg_deep_green'), ('between', 0.75, 'peg_green'),
                                     ('between', 1.0, 'peg_yellow'),
                                     ('between', 1.5, 'peg_orange'), ('between', 2.0, 'peg_red'),
                                     ('>', 3.0, 'peg_black')]:
                    props = {'type': 'cell', 'criteria': crit, 'format': fmt[f]}
                    if crit == 'between':
                        props.update({'minimum': val, 'maximum':
                            {'peg_green': 1.0, 'peg_yellow': 1.5, 'peg_orange': 2.0, 'peg_red': 3.0}[f]})
                    else:
                        props['value'] = val
                    sheet.conditional_format(rng, props)


        draw_grid_formulas(1, 'PRICE')
        draw_grid_formulas(14, 'UPSIDE')
        draw_grid_formulas(27, 'PEG')
        draw_grid_formulas(40, 'HOLDEN')

        # --- SUMMARY STATISTICS ---
        stats_col, r = 11, 1
        sheet.write(r, stats_col, "Summary Statistics", fmt['stat_head'])
        r += 1
        sheet.write(r, stats_col, "Market Data", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Name", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_name}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Sector", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_sector}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Market Cap", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_mcap}', fmt['stat_val_mcap'])
        r += 1
        sheet.write(r, stats_col, "Region", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_region}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Current Price", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_price}', fmt['stat_val'])
        addr_price = xl_rowcol_to_cell(r, stats_col + 1)

        r += 2
        sheet.write(r, stats_col, "Earnings Basis", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "GAAP EPS (Diluted)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_b_ttm}', fmt['stat_val_blue'])
        addr_basic = xl_rowcol_to_cell(r, stats_col + 1)
        r += 1
        sheet.write(r, stats_col, "Street EPS (Adj)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_s_ttm}', fmt['stat_val_blue'])
        addr_street = xl_rowcol_to_cell(r, stats_col + 1)

        r += 2
        r_forecast_start = r

        # Current FY Block
        sheet.write(r, stats_col, "Analyst Forecasts", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Target FY Period", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_est_year}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "EPS Low (Target FY)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_elo}', fmt['stat_val_blue'])
        r += 1
        sheet.write(r, stats_col, "EPS Consensus", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_emid}', fmt['stat_val_purple'])
        sheet.write_formula(r, stats_col + 2, f'=({ref_emid}-{f_act_ttm})/{f_act_ttm}', fmt['est_growth_base'])
        r += 1
        sheet.write(r, stats_col, "EPS High (Target FY)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_ehigh}', fmt['stat_val'])
        r += 1
        sheet.write(r, stats_col, "Growth Diagnosis", fmt['stat_label'])
        f_diag = f'=IF({addr_type_input}="Basic (GAAP)", IF({ref_b_ttm}<{ref_b_prior}, "⚠️ Cyclical Rebound", "✔ Organic"), IF({ref_s_ttm}<{ref_s_prior}, "⚠️ Cyclical Rebound", "✔ Organic"))'
        sheet.write_formula(r, stats_col + 1, f_diag, fmt['diag_ok'])

        # Next FY Block
        nxt_col = stats_col + 4
        r_nxt = r_forecast_start
        sheet.write(r_nxt, nxt_col, "Analyst Forecasts (Next FY)", fmt['stat_subhead'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "Target FY Period", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_est_year_nxt}', fmt['stat_val_txt'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS Low (Target FY)", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_elo_nxt}', fmt['stat_val_blue'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS Consensus", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_emid_nxt}', fmt['stat_val_purple'])
        sheet.write_formula(r_nxt, nxt_col + 2, f'=({ref_emid_nxt}-{f_act_ttm})/{f_act_ttm}', fmt['est_growth_base'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS High (Target FY)", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_ehigh_nxt}', fmt['stat_val'])

        r += 1
        sheet.write(r, stats_col, "Credibility (Surprise)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_surprise}/100', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "Estimate Dispersion", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=({dyn_ehigh}-{dyn_elo})/{dyn_emid}', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "EPS Trend (90d)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_eps_90d_change}', fmt['stat_val_score'])
        eps90d_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(eps90d_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(eps90d_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Analyst Coverage (#)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analysts}', fmt['stat_val_txt'])

        # --- VALUATION LOGIC --- (moved above Quality & Risk)
        r += 2
        sheet.write(r, stats_col, "Valuation Logic", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Implied P/E (Active)", fmt['stat_label'])
        addr_active_pe = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1,
                            f'={addr_price} / IF({addr_type_input}="Basic (GAAP)", {addr_basic}, {addr_street})',
                            fmt['stat_val_pe_grey'])
        r += 1
        sheet.write(r, stats_col, "Forward P/E (Est.)", fmt['stat_label'])
        addr_fwd_pe = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'={addr_price} / {dyn_emid}', fmt['stat_val_pe'])

        # --- QUALITY & RISK PROFILE ---
        r += 2
        sheet.write(r, stats_col, "Quality & Risk Profile", fmt['stat_subhead'])
        
        # Net Profit Margin
        r += 1
        sheet.write(r, stats_col, "Net Profit Margin", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_profit_margin}', fmt['stat_val_score'])
        cell_pm = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': '>=', 'value': 0.30, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.20, 'maximum': 0.2999, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.10, 'maximum': 0.1999, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.05, 'maximum': 0.0999, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': '<', 'value': 0.05, 'format': fmt['tier_black']})

        # FCF Yield
        r += 1
        sheet.write(r, stats_col, "Free Cash Flow Yield", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_fcf_yield}', fmt['stat_val_score'])
        cell_fcf = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': '>=', 'value': 0.08, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.05, 'maximum': 0.0799, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.025, 'maximum': 0.0499, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.0, 'maximum': 0.0249, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': '<', 'value': 0.0, 'format': fmt['tier_black']})
        
        # Debt-To-Equity
        r += 1
        sheet.write(r, stats_col, "Debt-to-Equity", fmt['stat_label'])
        f_de = f'=IF({ref_de_ratio}<0, "N/A", {ref_de_ratio})'
        sheet.write_formula(r, stats_col + 1, f_de, fmt['stat_val_pe'])
        cell_de = xl_rowcol_to_cell(r, stats_col + 1)
        # Using format mapping
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 0.0, 'maximum': 0.499, 'format': fmt['tier_de_dkgreen']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 0.5, 'maximum': 0.999, 'format': fmt['tier_de_green']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 1.0, 'maximum': 1.999, 'format': fmt['tier_de_yellow']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 2.0, 'maximum': 3.999, 'format': fmt['tier_de_red']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': '>=', 'value': 4.0, 'format': fmt['tier_de_black']})
        
        # Dividend Yield
        r += 1
        sheet.write(r, stats_col, "Dividend Yield", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_div_yield}', fmt['stat_val_score'])
        cell_div = xl_rowcol_to_cell(r, stats_col + 1)
        # Use existing ROE styling logic for yield
        sheet.conditional_format(cell_div, {'type': 'cell', 'criteria': '>=', 'value': 0.05, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_div, {'type': 'cell', 'criteria': 'between', 'minimum': 0.03, 'maximum': 0.0499, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_div, {'type': 'cell', 'criteria': 'between', 'minimum': 0.015, 'maximum': 0.0299, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_div, {'type': 'cell', 'criteria': 'between', 'minimum': 0.005, 'maximum': 0.0149, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_div, {'type': 'cell', 'criteria': '<', 'value': 0.005, 'format': fmt['tier_black']})
        
        # Dividend Ex Date
        r += 1
        sheet.write(r, stats_col, "Dividend Ex Date", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_div_ex_date}', fmt['stat_val_txt'])

        r += 2
        sheet.write(r, stats_col, "Holden Score (Base)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, '=G47', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "Realizable Upside %", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, '=IF(G34="NM", "NM", G21/MAX(1, G34))', fmt['stat_val_real'])

        # --- HISTORICAL TRENDS (1Y) ---
        r += 2
        sheet.write(r, stats_col, "Historical Trends (1Y)", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "1Y Price Performance", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_perf_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y EPS Trend", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_eps_trend_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y P/E Trend", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_pe_trend_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y Earnings-Return Gap (ERG)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=IF({ref_eps_trend_1y}>0, {ref_eps_trend_1y} - {ref_perf_1y}, "N/A (EPS < 0)")', fmt['stat_val_score'])
        ufa_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(ufa_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(ufa_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})

        # --- FORWARD-CONFIRMED ERG ---
        r += 2
        sheet.write(r, stats_col, "Forward-Confirmed ERG", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Implied FWD Growth", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_implied_fwd}', fmt['stat_val_score'])
        fwd_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(fwd_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(fwd_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Growth Confirm. (FCR)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_fcr}', fmt['stat_val_fcr'])
        fcr_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(fcr_cell, {'type': 'cell', 'criteria': '<', 'value': 0.3, 'format': fmt['diag_warn']})
        sheet.conditional_format(fcr_cell, {'type': 'cell', 'criteria': '>=', 'value': 0.7, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "ERG+ (Confirmed)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_erg_plus}', fmt['stat_val_score'])
        erg_plus_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(erg_plus_cell, {'type': 'cell', 'criteria': '<=', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(erg_plus_cell, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Regime", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_regime}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Implied Signal", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_regime_signal}', fmt['stat_val_txt'])

        # --- ANALYST PRICE TARGET (12M CONSENSUS) ---
        r += 2
        sheet.write(r, stats_col, "Analyst Price Target (12M)", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Mean Target (12M)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_mean}', fmt['stat_val'])
        r += 1
        sheet.write(r, stats_col, "Median Target (12M)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_median}', fmt['stat_val'])
        r += 1
        sheet.write(r, stats_col, "Target Low (12M)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_low}', fmt['stat_val_blue'])
        r += 1
        sheet.write(r, stats_col, "Target High (12M)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_high}', fmt['stat_val_blue'])
        r += 1
        sheet.write(r, stats_col, "Price Target Coverage (#)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_count}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Analyst Implied Upside", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analyst_target_upside}', fmt['stat_val_score'])
        analyst_up_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(analyst_up_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(analyst_up_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Multiple Expansion Signal", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_mult_expansion}', fmt['stat_val_score'])
        mexp_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(mexp_cell, {'type': 'cell', 'criteria': '>', 'value': 0.02, 'format': fmt['diag_ok']})
        sheet.conditional_format(mexp_cell, {'type': 'cell', 'criteria': '<', 'value': -0.02, 'format': fmt['diag_warn']})
        sheet.conditional_format(mexp_cell, {'type': 'cell', 'criteria': 'between', 'minimum': -0.02, 'maximum': 0.02, 'format': fmt['diag_mid']})
        r += 1
        sheet.write(r, stats_col, "Implied 12M P/E", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1,
                            f'=IF({dyn_emid}=0, "N/A", {ref_analyst_target_mean}/{dyn_emid})',
                            fmt['stat_val_pe'])

        # --- INSIDER ACTIVITY & CONVICTION SCORING ENGINE ---
        r += 2
        sheet.write(r, stats_col, "Insider Activity (L6M)", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Net Buying ($)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_insider}', fmt['insider_neutral'])
        insider_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(insider_cell,
                                 {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt['insider_pos']})
        sheet.conditional_format(insider_cell,
                                 {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['insider_neg']})
        r += 1
        sheet.write(r, stats_col, "Unique Buyers (#)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_insider_count}', fmt['stat_val_int'])
        r += 1
        sheet.write(r, stats_col, "Avg Stake Inc. (%)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_avg_stake}/100', fmt['stat_val_score'])

        # Conviction Scoring Formulas (Lone Wolf Fix)
        r += 1
        sheet.write(r, stats_col, "Conviction Score (0-10)", fmt['stat_label'])

        # Pillar 1: Materiality (Max 2)
        p1 = f'IF({ref_insider}<=0, 0, IF(OR({ref_insider}/MAX(1,{ref_mcap})>{INSIDER_PCT_LARGE}, {ref_insider}>={INSIDER_DOLLAR_LARGE}), 2, IF(OR({ref_insider}/MAX(1,{ref_mcap})>{INSIDER_PCT_MODERATE}, {ref_insider}>={INSIDER_DOLLAR_MODERATE}), 1, 0)))'
        # Pillar 2: Breadth (Max 4)
        p2 = f'IF({ref_insider_count}>=4, 4, {ref_insider_count})'
        # Pillar 3: Depth (Max 4) - Capped at 1 if Unique Buyers == 1
        p3 = f'IF({ref_insider_count}=0, 0, IF({ref_insider_count}=1, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_1}, 1, 0), IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_4}, 4, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_3}, 3, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_2}, 2, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_1}, 1, 0))))))'

        score_formula = f'=({p1}) + ({p2}) + ({p3})'
        sheet.write_formula(r, stats_col + 1, score_formula, fmt['stat_val_score_int'])

        # --- HOLDEN RESILIENCE --- (moved to bottom)
        r += 2
        sheet.write(r, stats_col, "Holden Resilience", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Allowed Safety Cushion", fmt['stat_label'])
        addr_cushion = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'=({addr_active_pe} - {addr_fwd_pe}) / {addr_active_pe}',
                            fmt['cushion_val'])
        r += 1
        sheet.write(r, stats_col, "Hist. Downside Risk", fmt['stat_label'])
        addr_risk = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'=MAX(0, ({addr_active_pe} - {ref_pel}) / {addr_active_pe})',
                            fmt['risk_val'])
        r += 1
        sheet.write(r, stats_col, "Resilience Ratio", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=IF({addr_risk}<=0, 99, {addr_cushion} / {addr_risk})',
                            fmt['resilience_good'])

        foot_row = r + 2
        sheet.write(foot_row, stats_col, "Holden Resilience Interpretation:", fmt['legend_bold'])
        sheet.write(foot_row + 1, stats_col, "> 1.0: Growth cushion covers historical downside (Safe)",
                    fmt['legend_norm'])

        # --- Ownership Watermark ---
        watermark_fmt = workbook.add_format({
            'font_name': 'Arial', 'font_size': 8, 'font_color': '#A6A6A6',
            'italic': True, 'align': 'left'
        })
        sheet.write(foot_row + 3, stats_col,
                    f"Holden Valuation Model \u00a9 {datetime.date.today().year} Dylan H Wilding. All rights reserved.",
                    watermark_fmt)
        sheet.set_header('', {'margin': 0})
        sheet.set_footer(
            '&L&8&I Holden Valuation Model \u00a9 Dylan H Wilding — Proprietary & Confidential'
        )

        sheet.set_column(0, 0, 2)
        sheet.set_column(1, 1, 6)
        sheet.set_column(2, 2, 12)
        sheet.set_column(3, 9, 10)
        sheet.set_column(11, 11, 22)
        sheet.set_column(12, 12, 16)
        sheet.set_column(13, 13, 12)
        sheet.set_column(14, 14, 2)
        sheet.set_column(15, 15, 22)
        sheet.set_column(16, 16, 18)

    # 2. CREATE COMPARISON SHEET
    print("  Generating Comparison Sheet...")
    ws_comp = workbook.add_worksheet("Comparison")
    fmt_comp_h = fmt['comp_head']

    cols = [
        "Ticker", "Company Name", "Sector", "Market Cap",
        "Price", "Target Price", "Implied Upside",
        "Current P/E (Adj)", "Forward P/E", "PEG Ratio",
        "Holden Score", "Safety Cushion", "Resilience Ratio",
        "Insider Net L6M ($)", "Unique Buyers", "Avg Stake Inc (%)",
        "Conviction Score (0-10)", "Growth Diagnosis",
        "1Y Perf", "1Y EPS \u0394", "1Y P/E \u0394", "1Y ERG Score",
        "Impl. FWD Growth", "FCR", "ERG+", "Regime", "Signal",
        "Net Profit Margin", "FCF Yield", "Debt/Equity", "Div. Yield", "Ex Date", "Region",
        "Analyst Mean Target", "Analyst Implied Upside", "Mult. Expansion Signal", "Implied 12M P/E"
    ]
    ws_comp.write_row(0, 0, cols, fmt_comp_h)
    comp_data = []

    for d in ALL_DATA:
        tick, name, sector, mcap, price = d['ticker'], d['company_name'], d['sector'], d['market_cap'], d['price']
        street_ttm, street_prior = d['eps_street_ttm'], d['eps_street_prior']
        basic_ttm, basic_prior = d['eps_basic_ttm'], d['eps_basic_prior']

        eps_ttm = street_ttm if street_ttm != 0 else basic_ttm
        if eps_ttm == 0: eps_ttm = 0.01
        eps_fwd = d['eps_mid'] if d['eps_mid'] != 0 else 0.01

        pe_low_hist = d['pe_low_hist']
        pe_curr = price / eps_ttm
        target_price = pe_curr * eps_fwd
        upside = (target_price - price) / price if price != 0 else 0
        pe_fwd = price / eps_fwd
        growth = (eps_fwd - eps_ttm) / eps_ttm
        peg = pe_curr / (growth * 100) if growth > 0.001 else 999.0
        holden_score = upside / peg if (peg > 0 and upside > 0 and peg != 999.0) else 0
        cushion = (pe_curr - pe_fwd) / pe_curr if pe_curr != 0 else 0
        risk = (pe_curr - pe_low_hist) / pe_curr if pe_curr != 0 else 0
        resilience = cushion / risk if risk > 0 else 99.0

        insider = d['insider_net']
        insider_count = d['insider_buy_count']
        avg_stake = d['insider_avg_stake_inc']

        # Calculate Conviction Score (Lone Wolf Fix applied here)
        p1 = 0
        if insider > 0:
            m = max(1, mcap)
            if (insider / m > INSIDER_PCT_LARGE) or (insider >= INSIDER_DOLLAR_LARGE):
                p1 = 2
            elif (insider / m > INSIDER_PCT_MODERATE) or (insider >= INSIDER_DOLLAR_MODERATE):
                p1 = 1
        p2 = min(4, insider_count)
        p3 = 0
        if insider_count == 1:
            if avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_1: p3 = 1
        elif insider_count > 1:
            if avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_4:
                p3 = 4
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_3:
                p3 = 3
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_2:
                p3 = 2
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_1:
                p3 = 1

        conviction_score = p1 + p2 + p3
        growth_diag = "⚠️ Cyclical Rebound" if (
            street_ttm < street_prior if street_ttm != 0 else basic_ttm < basic_prior) else "✔ Organic"

        eps_trend_1y = d['eps_trend_1y']
        pe_trend_1y = d['pe_trend_1y']
        perf_1y = d['perf_1y']
        erg_score = (eps_trend_1y - perf_1y) if eps_trend_1y > 0 else "N/A"
        implied_fwd_g = d['implied_fwd_growth']
        fcr_val = d['fcr']
        erg_plus_val = d['erg_plus']
        regime_val = d['regime']
        signal_val = d['regime_signal']

        # Implied 12M P/E = Analyst Mean Target / Forward EPS
        implied_12m_pe = d['analyst_target_mean'] / eps_fwd if (d['analyst_target_mean'] > 0 and eps_fwd > 0.01) else 0.0

        comp_data.append([tick, name, sector, mcap, price, target_price, upside, pe_curr, pe_fwd, peg,
                          holden_score, cushion, resilience, insider, insider_count, avg_stake / 100, conviction_score,
                          growth_diag, perf_1y, eps_trend_1y, pe_trend_1y, erg_score,
                          implied_fwd_g, fcr_val, erg_plus_val, regime_val, signal_val,
                          d['profit_margin'], d['fcf_yield'], d['de_ratio'], d['div_yield'], d['div_ex_date'], d['region'],
                          d['analyst_target_mean'], d['analyst_target_upside'], d['multiple_expansion_signal'],
                          implied_12m_pe])

    for i, row in enumerate(comp_data):
        ws_comp.write(i + 1, 0, row[0], fmt['comp_ticker'])
        ws_comp.write(i + 1, 1, row[1], fmt['comp_txt'])
        ws_comp.write(i + 1, 2, row[2], fmt['comp_txt'])
        ws_comp.write(i + 1, 3, row[3], fmt['comp_mcap'])
        ws_comp.write(i + 1, 4, row[4], fmt['comp_num'])
        ws_comp.write(i + 1, 5, row[5], fmt['comp_num'])
        ws_comp.write(i + 1, 6, row[6], fmt['comp_pct'])
        ws_comp.write(i + 1, 7, row[7], fmt['comp_pe'])
        ws_comp.write(i + 1, 8, row[8], fmt['comp_pe'])
        ws_comp.write(i + 1, 9, row[9], fmt['comp_num'])
        ws_comp.write(i + 1, 10, row[10], fmt['comp_pct'])
        ws_comp.write(i + 1, 11, row[11], fmt['comp_pct'])
        ws_comp.write(i + 1, 12, row[12], fmt['comp_num'])
        ws_comp.write(i + 1, 13, row[13], fmt['comp_dollar'])
        ws_comp.write(i + 1, 14, row[14], fmt['comp_int'])
        ws_comp.write(i + 1, 15, row[15], fmt['stat_val_score'])
        ws_comp.write(i + 1, 16, row[16], fmt['comp_int'])
        ws_comp.write(i + 1, 17, row[17], fmt['comp_txt'])
        ws_comp.write(i + 1, 18, row[18], fmt['comp_pct'])
        ws_comp.write(i + 1, 19, row[19], fmt['comp_pct'])
        ws_comp.write(i + 1, 20, row[20], fmt['comp_pct'])
        if isinstance(row[21], str):
            ws_comp.write(i + 1, 21, row[21], fmt['comp_txt'])
        else:
            ws_comp.write(i + 1, 21, row[21], fmt['comp_pct'])
        ws_comp.write(i + 1, 22, row[22], fmt['comp_pct'])
        ws_comp.write(i + 1, 23, row[23], fmt['comp_num'])
        ws_comp.write(i + 1, 24, row[24], fmt['comp_pct'])
        ws_comp.write(i + 1, 25, row[25], fmt['comp_txt'])
        ws_comp.write(i + 1, 26, row[26], fmt['comp_txt'])
        ws_comp.write(i + 1, 27, row[27], fmt['comp_pct'])
        ws_comp.write(i + 1, 28, row[28], fmt['comp_pct'])
        val_de = row[29]
        if val_de < 0:
            ws_comp.write(i + 1, 29, "N/A", fmt['comp_txt'])
        else:
            ws_comp.write(i + 1, 29, val_de, fmt['comp_num'])
        ws_comp.write(i + 1, 30, row[30], fmt['comp_pct'])
        ws_comp.write(i + 1, 31, row[31], fmt['comp_txt'])
        ws_comp.write(i + 1, 32, row[32], fmt['comp_txt'])  # Region
        ws_comp.write(i + 1, 33, row[33], fmt['comp_num'])  # Analyst Mean Target
        ws_comp.write(i + 1, 34, row[34], fmt['comp_pct'])  # Analyst Implied Upside
        ws_comp.write(i + 1, 35, row[35], fmt['comp_pct'])  # Mult. Expansion Signal
        ws_comp.write(i + 1, 36, row[36], fmt['comp_pe'])   # Implied 12M P/E

    if comp_data:
        ws_comp.add_table(0, 0, len(comp_data), len(cols) - 1, {
            'columns': [{'header': c} for c in cols],
            'style': 'TableStyleMedium2',
            'name': 'ValuationComparison'
        })

    ws_comp.set_column(0, 0, 10)
    ws_comp.set_column(1, 1, 25)
    ws_comp.set_column(2, 2, 22)
    ws_comp.set_column(3, 3, 16)
    ws_comp.set_column(4, 16, 14)
    ws_comp.set_column(17, 17, 20)
    ws_comp.set_column(18, 24, 14)
    ws_comp.set_column(25, 26, 22)
    ws_comp.set_column(27, 28, 14)
    ws_comp.set_column(29, 31, 14)
    ws_comp.set_column(32, 32, 18)  # Region
    ws_comp.set_column(33, 36, 16)  # Analyst Target columns

    # --- Ownership Watermark (Comparison Sheet) ---
    watermark_fmt_comp = workbook.add_format({
        'font_name': 'Arial', 'font_size': 8, 'font_color': '#A6A6A6',
        'italic': True, 'align': 'left'
    })
    ws_comp.write(len(comp_data) + 2, 0,
                  f"Holden Valuation Model \u00a9 {datetime.date.today().year} Dylan H Wilding. All rights reserved.",
                  watermark_fmt_comp)
    ws_comp.set_footer(
        '&L&8&I Holden Valuation Model \u00a9 Dylan H Wilding — Proprietary & Confidential'
    )

    # 3. GENERATE ANALYTICS DASHBOARD
    if len(comp_data) >= ANALYTICS_MIN_STOCKS:
        print("  Generating Analytics Dashboard...")
        build_analytics_dashboard(workbook, comp_data, cols)
    else:
        print(f"  Skipping Analytics Dashboard (need {ANALYTICS_MIN_STOCKS}+ stocks, have {len(comp_data)}).")

    # 4. CREATE INPUTS SHEET
    print("  Generating Inputs Sheet...")
    ws_inputs = workbook.add_worksheet("Inputs")
    try:
        ws_inputs.set_tab_color('#808080')
    except:
        pass

    lists = {'Growth': ['Analyst Consensus', 'Historical CAGR'], 'EPS': ['TTM (Current)', 'Analyst Low Est'],
             'PE': ['Flexible (Hist)', 'Static (15-45x)', 'Custom (20x)'],
             'Type': ['Basic (GAAP)', 'Street (Adjusted)'],
             'FY': ['Current', 'Next']}

    ws_inputs.write(0, 0, "List: Growth", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 0, lists['Growth'])
    ws_inputs.write(0, 1, "List: EPS", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 1, lists['EPS'])
    ws_inputs.write(0, 2, "List: PE", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 2, lists['PE'])
    ws_inputs.write(0, 3, "List: Type", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 3, lists['Type'])
    ws_inputs.write(0, 4, "List: FY", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 4, lists['FY'])

    headers = ['Ticker', 'Currency', 'Price', 'EPS Basic TTM', 'EPS Basic Prior', 'EPS Street TTM', 'EPS Street Prior',
               'EPS Low', 'EPS Mid', 'EPS High', 'PE Current', 'PE Low Hist', 'PE High Hist', 'CAGR 3Y', 'Sector',
               'Region',
               'Surprise Avg 4Q', 'Analyst Count', 'Est Target FY', 'Insider Net L6M', 'Company Name', 'Market Cap',
               'EPS Low Nxt', 'EPS Mid Nxt', 'EPS High Nxt', 'Est Target FY Nxt', 'Unique Buyers Count',
               'Avg Stake Increase %', 'Perf 1Y', 'EPS Trend 1Y', 'PE Trend 1Y',
               'Implied FWD Growth', 'FCR', 'ERG+', 'Regime', 'Regime Signal',
               'Profit Margin', 'FCF Yield', 'Debt to Equity', 'Div Yield', 'Ex Date',
               'Analyst Target Mean', 'Analyst Target Median', 'Analyst Target Low', 'Analyst Target High',
               'Analyst Target Upside', 'Multiple Expansion Signal', 'Analyst Target Count',
               'EPS 90d Change']

    fmt_head = workbook.add_format({'bold': True, 'bottom': 1})
    ws_inputs.write_row(data_start_row, 0, headers, fmt_head)

    for i, d in enumerate(ALL_DATA):
        r = data_start_row + 1 + i
        row_data = [d['ticker'], d['currency'], d['price'], d['eps_basic_ttm'], d.get('eps_basic_prior', 0),
                    d.get('eps_street_ttm', 0), d.get('eps_street_prior', 0), d['eps_low'], d['eps_mid'], d['eps_high'],
                    d['pe_current'], d['pe_low_hist'], d['pe_high_hist'], d['cagr_3y'], d['sector'], d['region'],
                    d.get('surprise_avg', 0), d.get('analysts', 0), d['est_year_str'], d['insider_net'],
                    d['company_name'], d['market_cap'], d['eps_low_nxt'], d['eps_mid_nxt'], d['eps_high_nxt'],
                    d['est_year_str_nxt'], d['insider_buy_count'], d['insider_avg_stake_inc'],
                    d['perf_1y'], d['eps_trend_1y'], d['pe_trend_1y'],
                    d['implied_fwd_growth'], d['fcr'], d['erg_plus'], d['regime'], d['regime_signal'],
                    d['profit_margin'], d['fcf_yield'], d['de_ratio'], d['div_yield'], d['div_ex_date'],
                    d['analyst_target_mean'], d['analyst_target_median'],
                    d['analyst_target_low'], d['analyst_target_high'],
                    d['analyst_target_upside'], d['multiple_expansion_signal'],
                    d['analyst_target_count'], d['eps_90d_change']]
        ws_inputs.write_row(r, 0, row_data)

    writer.close()
    print(f"\nSaved successfully to: {FILENAME}")

# =============================================================================
# Holden Valuation Model
# Copyright (c) 2026 Dylan H Wilding. All rights reserved.
#
# PROPRIETARY AND CONFIDENTIAL — This source code, its algorithms, scoring
# methodologies, and all associated outputs are the exclusive intellectual
# property of Dylan H Wilding. Unauthorized copying, distribution,
# modification, reverse-engineering, or commercial use of this file, in
# whole or in part, is strictly prohibited without prior written consent.
#
# Any use of this model by third parties (including but not limited to
# investment clubs, partnerships, or firms) constitutes a limited,
# non-exclusive, revocable license granted at the sole discretion of the
# author. Such use does not transfer ownership or intellectual property
# rights of any kind.
# =============================================================================
