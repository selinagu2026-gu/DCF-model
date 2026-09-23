"""
DCF Valuation Model — a Streamlit web app.

Enter any stock ticker, pull its live market data and historical financial
statements from Yahoo Finance, edit the DCF assumptions (WACC, growth,
margins, terminal growth), and see the implied share price update live —
plus a one-click download of the same model as a formula-driven Excel
workbook.

Run with:   streamlit run app.py
"""

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from dcf.data_fetcher import CompanyData, TickerNotFoundError, fetch_company_data
from dcf.dcf_engine import (
    Assumptions,
    compute_valuation,
    compute_wacc,
    default_assumptions,
    project_fcf,
    sensitivity_grid,
)
from dcf.excel_export import build_workbook

st.set_page_config(page_title="DCF Valuation Model", layout="wide")


@st.cache_data(ttl=900, show_spinner=False)
def _cached_fetch(ticker: str) -> CompanyData:
    return fetch_company_data(ticker)


def _fmt_money(x, decimals=0):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "n/a"
    return f"${x:,.{decimals}f}"


def _fmt_pct(x, decimals=1):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "n/a"
    return f"{x * 100:,.{decimals}f}%"


def _seed_assumption_state(data: CompanyData):
    d = default_assumptions(data)
    st.session_state["mkt_price"] = round(d.price, 2)
    st.session_state["mkt_shares"] = round(d.shares_out_mm, 2)
    st.session_state["mkt_debt"] = round(d.total_debt_mm, 2)
    st.session_state["mkt_cash"] = round(d.total_cash_mm, 2)
    st.session_state["rf_rate"] = round(d.risk_free_rate * 100, 3)
    st.session_state["erp"] = round(d.equity_risk_premium * 100, 3)
    st.session_state["beta"] = round(d.beta, 3)
    st.session_state["kd_pre"] = round(d.pretax_cost_of_debt * 100, 3)
    st.session_state["tax_rate"] = round(d.tax_rate * 100, 3)
    for i, g in enumerate(d.revenue_growth):
        st.session_state[f"g{i + 1}"] = round(g * 100, 3)
    st.session_state["margin_start"] = round(d.ebit_margin_start * 100, 3)
    st.session_state["margin_target"] = round(d.ebit_margin_target * 100, 3)
    st.session_state["da_pct"] = round(d.da_pct * 100, 3)
    st.session_state["capex_pct"] = round(d.capex_pct * 100, 3)
    st.session_state["nwc_pct"] = round(d.nwc_pct * 100, 3)
    st.session_state["term_g"] = round(d.terminal_growth * 100, 3)


# ---------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------
with st.sidebar:
    st.header("Company")
    ticker_input = st.text_input("Ticker", value=st.session_state.get("ticker_input", "DIS")).strip().upper()
    fetch_clicked = st.button("Fetch / Refresh Data", type="primary", use_container_width=True)

    if fetch_clicked:
        try:
            with st.spinner(f"Fetching {ticker_input} from Yahoo Finance..."):
                data = _cached_fetch(ticker_input)
            st.session_state["company_data"] = data
            st.session_state["ticker_input"] = ticker_input
            _seed_assumption_state(data)
        except TickerNotFoundError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Couldn't fetch data for '{ticker_input}': {e}")

    if "company_data" in st.session_state:
        if st.button("Reset assumptions to defaults", use_container_width=True):
            _seed_assumption_state(st.session_state["company_data"])

        st.divider()
        st.subheader("Market Data")
        st.session_state["mkt_price"] = st.number_input("Current Share Price ($)", value=float(st.session_state["mkt_price"]), min_value=0.01, step=0.01)
        st.session_state["mkt_shares"] = st.number_input("Shares Outstanding (mm)", value=float(st.session_state["mkt_shares"]), min_value=0.01, step=1.0)
        st.session_state["mkt_debt"] = st.number_input("Total Debt ($mm)", value=float(st.session_state["mkt_debt"]), min_value=0.0, step=10.0)
        st.session_state["mkt_cash"] = st.number_input("Total Cash & Equivalents ($mm)", value=float(st.session_state["mkt_cash"]), min_value=0.0, step=10.0)

        st.divider()
        st.subheader("WACC Assumptions")
        st.session_state["rf_rate"] = st.number_input("Risk-Free Rate (%)", value=float(st.session_state["rf_rate"]), step=0.05, format="%.2f")
        st.session_state["erp"] = st.number_input("Equity Risk Premium (%)", value=float(st.session_state["erp"]), step=0.10, format="%.2f")
        st.session_state["beta"] = st.number_input("Beta", value=float(st.session_state["beta"]), step=0.01, format="%.2f")
        st.session_state["kd_pre"] = st.number_input("Pre-Tax Cost of Debt (%)", value=float(st.session_state["kd_pre"]), step=0.05, format="%.2f")
        st.session_state["tax_rate"] = st.number_input("Tax Rate (%)", value=float(st.session_state["tax_rate"]), step=0.5, format="%.2f")

        st.divider()
        st.subheader("Growth & Margin Assumptions")
        for i in range(5):
            st.session_state[f"g{i + 1}"] = st.number_input(f"Year {i + 1} Revenue Growth (%)", value=float(st.session_state[f"g{i + 1}"]), step=0.25, format="%.2f")
        st.session_state["margin_start"] = st.number_input("Starting EBIT Margin (%)", value=float(st.session_state["margin_start"]), step=0.25, format="%.2f")
        st.session_state["margin_target"] = st.number_input("Target EBIT Margin, Year 5 (%)", value=float(st.session_state["margin_target"]), step=0.25, format="%.2f")
        st.session_state["da_pct"] = st.number_input("D&A (% of Revenue)", value=float(st.session_state["da_pct"]), step=0.1, format="%.2f")
        st.session_state["capex_pct"] = st.number_input("Capex (% of Revenue)", value=float(st.session_state["capex_pct"]), step=0.1, format="%.2f")
        st.session_state["nwc_pct"] = st.number_input("Incr. in NWC (% of Revenue)", value=float(st.session_state["nwc_pct"]), step=0.1, format="%.2f")
        st.session_state["term_g"] = st.number_input("Terminal Growth Rate (%)", value=float(st.session_state["term_g"]), step=0.1, format="%.2f")

# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------
st.title("DCF Valuation Model")
st.caption("Live data from Yahoo Finance. Educational purposes only — not investment advice.")

if "company_data" not in st.session_state:
    st.info("Enter a ticker in the sidebar and click **Fetch / Refresh Data** to build a DCF model.")
    st.stop()

data: CompanyData = st.session_state["company_data"]

current = Assumptions(
    risk_free_rate=st.session_state["rf_rate"] / 100,
    equity_risk_premium=st.session_state["erp"] / 100,
    beta=st.session_state["beta"],
    pretax_cost_of_debt=st.session_state["kd_pre"] / 100,
    tax_rate=st.session_state["tax_rate"] / 100,
    revenue_growth=[st.session_state[f"g{i + 1}"] / 100 for i in range(5)],
    ebit_margin_start=st.session_state["margin_start"] / 100,
    ebit_margin_target=st.session_state["margin_target"] / 100,
    da_pct=st.session_state["da_pct"] / 100,
    capex_pct=st.session_state["capex_pct"] / 100,
    nwc_pct=st.session_state["nwc_pct"] / 100,
    terminal_growth=st.session_state["term_g"] / 100,
    price=st.session_state["mkt_price"],
    shares_out_mm=st.session_state["mkt_shares"],
    market_cap_mm=st.session_state["mkt_price"] * st.session_state["mkt_shares"],
    total_debt_mm=st.session_state["mkt_debt"],
    total_cash_mm=st.session_state["mkt_cash"],
    base_revenue_mm=data.revenue[-1],
)

wacc_info = compute_wacc(current)
wacc = wacc_info["wacc"]

st.subheader(f"{data.name} ({data.ticker})")
st.caption(f"{data.sector} – {data.industry}  |  Data retrieved {data.as_of}")

if wacc <= current.terminal_growth:
    st.error("WACC must be greater than the terminal growth rate. Adjust the assumptions in the sidebar.")
    st.stop()

projection = project_fcf(current)
valuation = compute_valuation(current, projection, wacc)
w_values, g_values, grid = sensitivity_grid(current, projection, wacc)

# --- Headline metrics ---
c1, c2, c3, c4 = st.columns(4)
c1.metric("Current Price", _fmt_money(current.price, 2))
c2.metric("Implied Value / Share", _fmt_money(valuation["implied_price"], 2),
          delta=_fmt_pct(valuation["upside"]))
c3.metric("WACC", _fmt_pct(wacc, 2))
c4.metric("Analyst Avg. Target", _fmt_money(data.analyst_target_mean, 2) if data.analyst_target_mean else "n/a",
          help=f"{data.analyst_count or 0} analysts (Yahoo Finance)")

if valuation["tv_pct_ev"] > 0.80:
    st.warning(f"Terminal value is {valuation['tv_pct_ev']*100:.0f}% of enterprise value — "
               f"this valuation is highly sensitive to the terminal growth / WACC assumptions.")

tab_hist, tab_wacc, tab_dcf, tab_sens, tab_export = st.tabs(
    ["Historicals", "WACC Build", "DCF Projection", "Sensitivity", "Export"]
)

# --- Historicals ---
with tab_hist:
    hist_df = pd.DataFrame({
        "Revenue ($mm)": data.revenue,
        "EBIT ($mm)": data.ebit,
        "EBIT Margin": [e / r if e is not None and r else None for e, r in zip(data.ebit, data.revenue)],
        "D&A ($mm)": data.da,
        "Capex ($mm)": data.capex,
        "Free Cash Flow ($mm)": data.fcf,
        "Net Income ($mm)": data.net_income,
        "Total Debt ($mm)": data.bs_total_debt,
        "Working Capital ($mm)": data.bs_working_capital,
    }, index=data.fiscal_years).T
    st.dataframe(hist_df.style.format(lambda v: "n/a" if v is None or pd.isna(v) else (f"{v:,.1%}" if abs(v) < 5 else f"{v:,.0f}")), use_container_width=True)

    fig = go.Figure()
    fig.add_bar(x=data.fiscal_years, y=data.revenue, name="Revenue")
    fig.add_bar(x=data.fiscal_years, y=data.ebit, name="EBIT")
    fig.update_layout(barmode="group", height=350, title="Historical Revenue & EBIT ($mm)")
    st.plotly_chart(fig, use_container_width=True)

# --- WACC ---
with tab_wacc:
    wacc_rows = {
        "Risk-Free Rate": _fmt_pct(current.risk_free_rate, 2),
        "Beta": f"{current.beta:.2f}",
        "Equity Risk Premium": _fmt_pct(current.equity_risk_premium, 2),
        "Cost of Equity (CAPM)": _fmt_pct(wacc_info["cost_of_equity"], 2),
        "Pre-Tax Cost of Debt": _fmt_pct(current.pretax_cost_of_debt, 2),
        "Tax Rate": _fmt_pct(current.tax_rate, 2),
        "After-Tax Cost of Debt": _fmt_pct(wacc_info["after_tax_cost_of_debt"], 2),
        "Weight of Equity": _fmt_pct(wacc_info["weight_equity"], 1),
        "Weight of Debt": _fmt_pct(wacc_info["weight_debt"], 1),
        "WACC": _fmt_pct(wacc, 2),
    }
    st.table(pd.Series(wacc_rows, name="Value"))

# --- DCF Projection ---
with tab_dcf:
    proj_years = [f"FY{int(data.fiscal_years[-1][2:6]) + i}E" for i in range(1, 6)]
    proj_df = pd.DataFrame({
        "Revenue Growth": [row["growth"] for row in projection],
        "Total Revenue": [row["revenue"] for row in projection],
        "EBIT Margin": [row["ebit_margin"] for row in projection],
        "EBIT": [row["ebit"] for row in projection],
        "NOPAT": [row["nopat"] for row in projection],
        "D&A": [row["da"] for row in projection],
        "Capex": [row["capex"] for row in projection],
        "Incr. NWC": [row["nwc"] for row in projection],
        "Unlevered FCF": [row["ufcf"] for row in projection],
        "PV of FCF": valuation["pv_list"],
    }, index=proj_years).T
    st.dataframe(proj_df.style.format(lambda v: f"{v:,.1%}" if abs(v) < 3 else f"{v:,.0f}"), use_container_width=True)

    fig = go.Figure()
    fig.add_bar(x=proj_years, y=[row["ufcf"] for row in projection], name="Unlevered FCF")
    fig.add_bar(x=proj_years, y=valuation["pv_list"], name="PV of FCF")
    fig.update_layout(barmode="group", height=350, title="Projected Unlevered Free Cash Flow ($mm)")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("**Valuation Summary**")
    v1, v2, v3 = st.columns(3)
    v1.metric("Sum PV of Explicit FCF", _fmt_money(valuation["sum_pv"]))
    v1.metric("PV of Terminal Value", _fmt_money(valuation["pv_terminal_value"]))
    v2.metric("Enterprise Value", _fmt_money(valuation["enterprise_value"]))
    v2.metric("Net Debt", _fmt_money(valuation["net_debt"]))
    v3.metric("Equity Value", _fmt_money(valuation["equity_value"]))
    v3.metric("Implied Value / Share", _fmt_money(valuation["implied_price"], 2))

# --- Sensitivity ---
with tab_sens:
    st.caption("Implied share price across WACC (rows) and terminal growth rate (columns). "
               "Explicit-period FCF held constant at the base case above.")
    z = grid
    fig = go.Figure(data=go.Heatmap(
        z=z,
        x=[f"{g*100:.2f}%" for g in g_values],
        y=[f"{w*100:.2f}%" for w in w_values],
        text=[[("n/m" if pd.isna(v) else f"${v:,.2f}") for v in row] for row in z],
        texttemplate="%{text}",
        colorscale="RdYlGn",
        colorbar_title="$/share",
    ))
    fig.update_layout(height=400, xaxis_title="Terminal Growth Rate", yaxis_title="WACC")
    st.plotly_chart(fig, use_container_width=True)

# --- Export ---
with tab_export:
    st.markdown("Download the full model — Cover, Assumptions, Historicals, DCF Model, and Sensitivity tabs — "
                "as a formula-driven Excel workbook. Editing any blue cell in Excel will recalculate the whole model.")
    xlsx_buf = build_workbook(data, current)
    st.download_button(
        "Download Excel Model (.xlsx)",
        data=xlsx_buf,
        file_name=f"{data.ticker}_DCF_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.divider()
st.caption("Data: Yahoo Finance via the yfinance library. This tool is for educational purposes only and does not "
           "constitute investment advice. Review and adjust every assumption before relying on the output.")
