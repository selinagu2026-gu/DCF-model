"""
Pure calculation logic for the DCF model: WACC, 5-year unlevered free cash
flow projection, terminal value, implied share price, and a WACC x terminal
growth sensitivity grid.

Nothing in this file touches the network, Streamlit, or Excel — it only
takes numbers in and returns numbers out, so the same logic can be reused
by both the web app and the Excel exporter (keeping them from drifting
out of sync with each other).
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
import statistics

from .data_fetcher import CompanyData

N_YEARS = 5


@dataclass
class Assumptions:
    risk_free_rate: float
    equity_risk_premium: float
    beta: float
    pretax_cost_of_debt: float
    tax_rate: float
    revenue_growth: list[float]  # length N_YEARS, Y1..Y5
    ebit_margin_start: float
    ebit_margin_target: float
    da_pct: float
    capex_pct: float
    nwc_pct: float
    terminal_growth: float
    # Market data, carried alongside so the whole model can be recomputed
    # from one Assumptions object.
    price: float
    shares_out_mm: float
    market_cap_mm: float
    total_debt_mm: float
    total_cash_mm: float
    base_revenue_mm: float


def _effective_tax_rates(data: CompanyData) -> list[float]:
    rates = []
    for pretax, prov in zip(data.pretax_income, data.tax_provision):
        if pretax and prov is not None and pretax > 0:
            rates.append(prov / pretax)
    return rates


def default_assumptions(data: CompanyData) -> Assumptions:
    """Derive a full, editable set of DCF assumptions straight from the
    fetched historical financials — every field here is a starting point
    the user can override in the UI."""

    base_revenue = data.revenue[-1]
    base_ebit = data.ebit[-1]
    margin_start = (base_ebit / base_revenue) if base_ebit is not None and base_revenue else 0.10

    tax_rates = _effective_tax_rates(data)
    if tax_rates:
        tax_rate = min(max(statistics.median(tax_rates), 0.10), 0.35)
    else:
        tax_rate = 0.25

    # Capex and D&A are often lumpy year to year (e.g. a one-off capacity
    # buildout), so average the ratio over the last up-to-3 actual years
    # rather than taking a single year, which would over- or under-state
    # the forward run rate.
    n_avg = min(3, len(data.revenue))

    def _avg_pct(values: list[float | None]) -> float | None:
        ratios = [v / r for v, r in zip(values[-n_avg:], data.revenue[-n_avg:]) if v is not None and r]
        return statistics.mean(ratios) if ratios else None

    da_pct = _avg_pct(data.da)
    if da_pct is None:
        da_pct = 0.05
    capex_pct = _avg_pct(data.capex)
    if capex_pct is None:
        capex_pct = 0.05

    risk_free = data.risk_free_rate if data.risk_free_rate is not None else 0.045
    beta = data.beta if data.beta is not None else 1.0

    terminal_growth = 0.025

    g1 = data.consensus_growth_y1
    if g1 is None:
        # Fall back to the most recent historical YoY growth rate.
        if len(data.revenue) >= 2 and data.revenue[-2]:
            g1 = data.revenue[-1] / data.revenue[-2] - 1
        else:
            g1 = 0.05
    g2 = data.consensus_growth_y2 if data.consensus_growth_y2 is not None else g1 * 0.8

    # Fade years 3-5 linearly from Y2's growth down to the terminal rate.
    g3 = g2 + (terminal_growth - g2) * (1 / 3)
    g4 = g2 + (terminal_growth - g2) * (2 / 3)
    g5 = terminal_growth

    return Assumptions(
        risk_free_rate=round(risk_free, 4),
        equity_risk_premium=0.05,
        beta=round(beta, 3),
        pretax_cost_of_debt=round(risk_free + 0.015, 4),
        tax_rate=round(tax_rate, 4),
        revenue_growth=[round(g1, 4), round(g2, 4), round(g3, 4), round(g4, 4), round(g5, 4)],
        ebit_margin_start=round(margin_start, 4),
        ebit_margin_target=round(margin_start, 4),
        da_pct=round(da_pct, 4),
        capex_pct=round(capex_pct, 4),
        nwc_pct=0.0,
        terminal_growth=terminal_growth,
        price=data.price,
        shares_out_mm=data.shares_out_mm,
        market_cap_mm=data.market_cap_mm,
        total_debt_mm=data.total_debt_mm,
        total_cash_mm=data.total_cash_mm,
        base_revenue_mm=base_revenue,
    )


def compute_wacc(a: Assumptions) -> dict:
    cost_of_equity = a.risk_free_rate + a.beta * a.equity_risk_premium
    after_tax_kd = a.pretax_cost_of_debt * (1 - a.tax_rate)
    total_cap = a.market_cap_mm + a.total_debt_mm
    weight_equity = a.market_cap_mm / total_cap if total_cap else 1.0
    weight_debt = a.total_debt_mm / total_cap if total_cap else 0.0
    wacc = weight_equity * cost_of_equity + weight_debt * after_tax_kd
    return dict(
        cost_of_equity=cost_of_equity,
        after_tax_cost_of_debt=after_tax_kd,
        weight_equity=weight_equity,
        weight_debt=weight_debt,
        wacc=wacc,
    )


def project_fcf(a: Assumptions) -> list[dict]:
    rows = []
    prev_revenue = a.base_revenue_mm
    for t in range(1, N_YEARS + 1):
        growth = a.revenue_growth[t - 1]
        revenue = prev_revenue * (1 + growth)
        margin = a.ebit_margin_start + (a.ebit_margin_target - a.ebit_margin_start) * (t / N_YEARS)
        ebit = revenue * margin
        taxes = ebit * a.tax_rate
        nopat = ebit - taxes
        da = revenue * a.da_pct
        capex = revenue * a.capex_pct
        nwc = revenue * a.nwc_pct
        ufcf = nopat + da - capex - nwc
        rows.append(dict(
            year=t, growth=growth, revenue=revenue, ebit_margin=margin, ebit=ebit,
            taxes=taxes, nopat=nopat, da=da, capex=capex, nwc=nwc, ufcf=ufcf,
        ))
        prev_revenue = revenue
    return rows


def _discount(ufcf_list: list[float], wacc: float, terminal_growth: float) -> dict:
    pv_list = [cf / (1 + wacc) ** (i + 1) for i, cf in enumerate(ufcf_list)]
    sum_pv = sum(pv_list)
    n = len(ufcf_list)
    tv = ufcf_list[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_tv = tv / (1 + wacc) ** n
    ev = sum_pv + pv_tv
    return dict(pv_list=pv_list, sum_pv=sum_pv, terminal_value=tv, pv_terminal_value=pv_tv, enterprise_value=ev)


def compute_valuation(a: Assumptions, projection: list[dict], wacc: float) -> dict:
    ufcf_list = [row["ufcf"] for row in projection]
    disc = _discount(ufcf_list, wacc, a.terminal_growth)
    net_debt = a.total_debt_mm - a.total_cash_mm
    equity_value = disc["enterprise_value"] - net_debt
    implied_price = equity_value / a.shares_out_mm if a.shares_out_mm else float("nan")
    upside = (implied_price / a.price - 1) if a.price else float("nan")
    tv_pct_ev = disc["pv_terminal_value"] / disc["enterprise_value"] if disc["enterprise_value"] else float("nan")
    return dict(
        **disc,
        net_debt=net_debt,
        equity_value=equity_value,
        implied_price=implied_price,
        upside=upside,
        tv_pct_ev=tv_pct_ev,
    )


def sensitivity_grid(a: Assumptions, projection: list[dict], base_wacc: float,
                      w_offsets=(-0.02, -0.01, 0.0, 0.01, 0.02),
                      g_offsets=(-0.01, -0.005, 0.0, 0.005, 0.01)) -> tuple[list[float], list[float], list[list[float]]]:
    """Returns (wacc_values, growth_values, price_grid) where price_grid[i][j]
    is the implied share price at wacc_values[i], growth_values[j]."""
    ufcf_list = [row["ufcf"] for row in projection]
    net_debt = a.total_debt_mm - a.total_cash_mm
    wacc_values = [round(base_wacc + off, 5) for off in w_offsets]
    growth_values = [round(a.terminal_growth + off, 5) for off in g_offsets]

    grid = []
    for w in wacc_values:
        row = []
        for g in growth_values:
            if w <= g:
                row.append(float("nan"))
                continue
            disc = _discount(ufcf_list, w, g)
            equity_value = disc["enterprise_value"] - net_debt
            price = equity_value / a.shares_out_mm if a.shares_out_mm else float("nan")
            row.append(price)
        grid.append(row)
    return wacc_values, growth_values, grid
