"""
Pulls a company's market data, historical financial statements, analyst
estimates, and the risk-free rate from Yahoo Finance (via the yfinance
library) and packages them into a single CompanyData object that the
rest of the app (dcf_engine, excel_export, app.py) consumes.

This module is the only place that talks to Yahoo Finance. If Yahoo
changes a field name, fix it here.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone

import pandas as pd
import yfinance as yf


class TickerNotFoundError(Exception):
    """Raised when a ticker has no usable financial statements (e.g. it's
    an ETF/fund, an index, or an invalid symbol)."""


@dataclass
class CompanyData:
    ticker: str
    name: str
    sector: str
    industry: str
    currency: str
    as_of: str  # human-readable retrieval timestamp, for citations

    # --- Live market data ---
    price: float
    shares_out_mm: float
    market_cap_mm: float
    beta: float | None
    total_debt_mm: float
    total_cash_mm: float
    analyst_target_mean: float | None
    analyst_count: int | None

    # --- Macro ---
    risk_free_rate: float | None  # 10-Yr UST yield, as a decimal (e.g. 0.048)

    # --- Analyst consensus revenue growth (decimals), may be None ---
    consensus_growth_y1: float | None
    consensus_growth_y2: float | None

    # --- Historical annual financials, oldest -> newest, $ millions ---
    fiscal_years: list[str]  # e.g. ["FY2022A", "FY2023A", "FY2024A", "FY2025A"]
    revenue: list[float]
    ebit: list[float]
    ebitda: list[float]
    da: list[float]  # depreciation & amortization
    capex: list[float]  # positive number = cash spent
    pretax_income: list[float]
    tax_provision: list[float]
    net_income: list[float]
    ocf: list[float]
    fcf: list[float]
    bs_total_debt: list[float]
    bs_total_equity: list[float]
    bs_working_capital: list[float]
    bs_cash: list[float]
    bs_shares: list[float]


def _row_mm(df: pd.DataFrame | None, name: str, cols) -> list[float | None]:
    """Extract a row from a yfinance statement, in $ millions, aligned to `cols`."""
    if df is None or name not in df.index:
        return [None] * len(cols)
    s = df.loc[name]
    out = []
    for c in cols:
        v = s.get(c) if c in s.index else None
        out.append(round(float(v) / 1e6, 3) if v is not None and pd.notna(v) else None)
    return out


def _first_present(df: pd.DataFrame | None, names: list[str], cols) -> list[float | None]:
    """Try several possible row labels (Yahoo renames these occasionally) and
    return the first one that exists with data."""
    if df is None:
        return [None] * len(cols)
    for name in names:
        if name in df.index:
            row = _row_mm(df, name, cols)
            if any(v is not None for v in row):
                return row
    return [None] * len(cols)


def fetch_risk_free_rate() -> float | None:
    """Live 10-Year US Treasury yield from the ^TNX index (quoted x10)."""
    try:
        tnx = yf.Ticker("^TNX")
        hist = tnx.history(period="5d")
        if not hist.empty:
            return round(float(hist["Close"].dropna().iloc[-1]) / 100.0, 5)
    except Exception:
        pass
    return None


def fetch_company_data(ticker: str) -> CompanyData:
    ticker = ticker.strip().upper()
    t = yf.Ticker(ticker)
    info = t.info or {}

    fin = t.financials  # annual income statement
    bs = t.balance_sheet
    cf = t.cashflow

    if fin is None or fin.empty or "Total Revenue" not in fin.index:
        raise TickerNotFoundError(
            f"'{ticker}' doesn't have standard company financial statements on "
            f"Yahoo Finance (this happens for ETFs, funds, and indices). "
            f"A DCF model requires an operating company — try its stock ticker instead."
        )

    # Keep only columns (fiscal years) where revenue was actually reported,
    # oldest first.
    rev_row = fin.loc["Total Revenue"]
    valid_cols = [c for c in fin.columns if pd.notna(rev_row.get(c))]
    valid_cols = sorted(valid_cols)  # ascending date order
    if len(valid_cols) == 0:
        raise TickerNotFoundError(f"No annual revenue history found for '{ticker}'.")

    fiscal_years = [f"FY{pd.Timestamp(c).year}A" for c in valid_cols]

    revenue = _row_mm(fin, "Total Revenue", valid_cols)
    ebit = _first_present(fin, ["EBIT"], valid_cols)
    ebitda = _first_present(fin, ["EBITDA"], valid_cols)
    da = _first_present(fin, ["Reconciled Depreciation",
                              "Depreciation Amortization Depletion Income Statement",
                              "Depreciation And Amortization In Income Statement"], valid_cols)
    pretax = _first_present(fin, ["Pretax Income"], valid_cols)
    tax_prov = _first_present(fin, ["Tax Provision"], valid_cols)
    net_income = _first_present(fin, ["Net Income"], valid_cols)

    capex_raw = _first_present(cf, ["Capital Expenditure"], valid_cols)
    capex = [(-v if v is not None else None) for v in capex_raw]  # store as positive spend
    ocf = _first_present(cf, ["Operating Cash Flow"], valid_cols)
    fcf = _first_present(cf, ["Free Cash Flow"], valid_cols)
    da_cf = _first_present(cf, ["Depreciation And Amortization", "Depreciation Amortization Depletion"], valid_cols)
    for i in range(len(da)):
        if da[i] is None:
            da[i] = da_cf[i]

    bs_total_debt = _first_present(bs, ["Total Debt"], valid_cols)
    bs_total_equity = _first_present(bs, ["Stockholders Equity", "Common Stock Equity"], valid_cols)
    bs_working_capital = _first_present(bs, ["Working Capital"], valid_cols)
    bs_cash = _first_present(bs, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"], valid_cols)
    bs_shares = _first_present(bs, ["Ordinary Shares Number", "Share Issued"], valid_cols)

    # If EBIT wasn't reported directly, derive it from EBITDA - D&A, or Pretax + Interest.
    for i in range(len(ebit)):
        if ebit[i] is None and ebitda[i] is not None and da[i] is not None:
            ebit[i] = round(ebitda[i] - da[i], 3)
    for i in range(len(ebitda)):
        if ebitda[i] is None and ebit[i] is not None and da[i] is not None:
            ebitda[i] = round(ebit[i] + da[i], 3)

    price = info.get("currentPrice") or info.get("regularMarketPrice")
    shares_out = info.get("sharesOutstanding")
    market_cap = info.get("marketCap")

    # Analyst consensus revenue growth for the current and next fiscal year.
    g_y1 = g_y2 = None
    try:
        rev_est = t.revenue_estimate
        if rev_est is not None and not rev_est.empty:
            if "0y" in rev_est.index and pd.notna(rev_est.loc["0y", "growth"]):
                g_y1 = round(float(rev_est.loc["0y", "growth"]), 4)
            if "+1y" in rev_est.index and pd.notna(rev_est.loc["+1y", "growth"]):
                g_y2 = round(float(rev_est.loc["+1y", "growth"]), 4)
    except Exception:
        pass

    return CompanyData(
        ticker=ticker,
        name=info.get("longName") or info.get("shortName") or ticker,
        sector=info.get("sector") or "n/a",
        industry=info.get("industry") or "n/a",
        currency=info.get("currency") or "USD",
        as_of=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        price=float(price) if price else float("nan"),
        shares_out_mm=round(shares_out / 1e6, 2) if shares_out else float("nan"),
        market_cap_mm=round(market_cap / 1e6, 2) if market_cap else float("nan"),
        beta=info.get("beta"),
        total_debt_mm=round(info.get("totalDebt", 0) / 1e6, 2) if info.get("totalDebt") else 0.0,
        total_cash_mm=round(info.get("totalCash", 0) / 1e6, 2) if info.get("totalCash") else 0.0,
        analyst_target_mean=info.get("targetMeanPrice"),
        analyst_count=info.get("numberOfAnalystOpinions"),
        risk_free_rate=fetch_risk_free_rate(),
        consensus_growth_y1=g_y1,
        consensus_growth_y2=g_y2,
        fiscal_years=fiscal_years,
        revenue=revenue,
        ebit=ebit,
        ebitda=ebitda,
        da=da,
        capex=capex,
        pretax_income=pretax,
        tax_provision=tax_prov,
        net_income=net_income,
        ocf=ocf,
        fcf=fcf,
        bs_total_debt=bs_total_debt,
        bs_total_equity=bs_total_equity,
        bs_working_capital=bs_working_capital,
        bs_cash=bs_cash,
        bs_shares=bs_shares,
    )
