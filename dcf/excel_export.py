"""
Builds the same styled, formula-driven Excel DCF workbook (Cover,
Assumptions, Historicals, DCF Model, Sensitivity) that was originally
hand-built for Disney, but generically for whatever ticker + Assumptions
the web app currently has loaded.

Every number that came from Yahoo Finance is written as a hardcoded
(blue) input cell with a source comment; every derived number is a
live Excel formula, so the workbook keeps working if you edit the
blue cells after downloading it.
"""

from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from .data_fetcher import CompanyData
from .dcf_engine import Assumptions, N_YEARS

FONT_NAME = "Arial"

BLUE = Font(name=FONT_NAME, color="0000FF", size=10)
BLACK = Font(name=FONT_NAME, color="000000", size=10)
BLACK_B = Font(name=FONT_NAME, color="000000", size=10, bold=True)
GREEN = Font(name=FONT_NAME, color="008000", size=10)
TITLE_F = Font(name=FONT_NAME, color="FFFFFF", size=14, bold=True)
SUBTITLE_F = Font(name=FONT_NAME, color="404040", size=9, italic=True)
HDR_F = Font(name=FONT_NAME, color="FFFFFF", size=10, bold=True)
SECTION_F = Font(name=FONT_NAME, color="FFFFFF", size=11, bold=True)
NOTE_F = Font(name=FONT_NAME, size=8, italic=True, color="808080")

TITLE_FILL = PatternFill("solid", fgColor="0B3D91")
SECTION_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FILL = PatternFill("solid", fgColor="2E5395")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
KEY_FILL = PatternFill("solid", fgColor="D9E1F2")

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
TOPONLY = Border(top=Side(style="thin", color="000000"))

CUR0 = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT1 = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM0 = '#,##0;(#,##0);"-"'


def _comment(ws, cell, text):
    c = Comment(text, "DCF Model")
    c.width = 340
    c.height = 150
    ws[cell].comment = c


def _section_row(ws, row, text, last_col_letter):
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = SECTION_F
    c.fill = SECTION_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def build_workbook(data: CompanyData, a: Assumptions) -> BytesIO:
    n_hist = len(data.fiscal_years)
    hist_cols = [get_column_letter(2 + i) for i in range(n_hist)]  # B, C, D, ...
    last_hist_col = hist_cols[-1]
    last_year_num = int(data.fiscal_years[-1][2:6])
    proj_years = [f"FY{last_year_num + i}E" for i in range(1, N_YEARS + 1)]
    proj_cols = ["B", "C", "D", "E", "F"]

    wb = Workbook()

    # ---------------------------------------------------------------
    # COVER
    # ---------------------------------------------------------------
    cov = wb.active
    cov.title = "Cover"
    cov.sheet_view.showGridLines = False
    for i, w in enumerate([4, 95, 4], start=1):
        cov.column_dimensions[get_column_letter(i)].width = w

    cov["B2"] = f"{data.name} ({data.ticker})"
    cov["B2"].font = Font(name=FONT_NAME, size=20, bold=True, color="0B3D91")
    cov["B3"] = "Discounted Cash Flow (DCF) Valuation Model"
    cov["B3"].font = Font(name=FONT_NAME, size=14, color="404040")
    cov["B5"] = f"Generated: {a.__dict__.get('as_of', data.as_of)}"
    cov["B5"].font = Font(name=FONT_NAME, size=10, italic=True)

    notes = [
        ("How this workbook is organized", True),
        ("Assumptions – market data, CAPM/WACC build, and forward operating assumptions. Edit the blue cells to run your own scenarios.", False),
        (f"Historicals – {data.fiscal_years[0]}–{data.fiscal_years[-1]} actual financial statements.", False),
        ("DCF Model – 5-year unlevered free cash flow projection, terminal value, and implied share price.", False),
        ("Sensitivity – implied share price grid across WACC and terminal growth rate.", False),
        ("", False),
        ("Color legend", True),
        ("Blue text = hardcoded input / assumption (safe to edit)", False),
        ("Black text = formula (do not overwrite)", False),
        ("Green text = link to another sheet", False),
        ("Yellow fill = key output cell", False),
        ("", False),
        ("Data sources", True),
        (f"Market data, beta, historical financial statements: Yahoo Finance (via the yfinance Python library), retrieved {data.as_of}.", False),
        ("Risk-free rate: 10-Year U.S. Treasury yield (Yahoo Finance ^TNX index).", False),
        (f"Analyst consensus revenue growth (where available): Yahoo Finance revenue estimates, {data.analyst_count or 'n/a'} analysts, avg price target "
         f"{'$'+format(data.analyst_target_mean, ',.2f') if data.analyst_target_mean else 'n/a'}.", False),
        ("", False),
        ("Important", True),
        ("Generated by the DCF Modle web app. Educational purposes only — not investment advice. Forward assumptions (growth, margins, WACC, terminal growth) are editable estimates; review before relying on the output.", False),
    ]
    r = 7
    for text, is_header in notes:
        cell = cov[f"B{r}"]
        cell.value = text
        if is_header:
            cell.font = Font(name=FONT_NAME, size=12, bold=True, color="1F4E79")
        else:
            cell.font = Font(name=FONT_NAME, size=10, color="000000")
            cell.alignment = Alignment(wrap_text=True)
        r += 1
    cov.row_dimensions[1].height = 6

    # ---------------------------------------------------------------
    # ASSUMPTIONS
    # ---------------------------------------------------------------
    asm = wb.create_sheet("Assumptions")
    asm.sheet_view.showGridLines = False
    for i, w in enumerate([42, 16, 55], start=1):
        asm.column_dimensions[get_column_letter(i)].width = w

    asm.merge_cells("A1:C1")
    asm["A1"] = f"{data.name} ({data.ticker}) — Assumptions & WACC"
    asm["A1"].font = TITLE_F
    asm["A1"].fill = TITLE_FILL
    asm["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    asm.row_dimensions[1].height = 22
    asm.merge_cells("A2:C2")
    asm["A2"] = "All $ figures in millions except per-share data. Blue = input, Black = formula, Green = link to another sheet."
    asm["A2"].font = SUBTITLE_F

    def kv(row, lbl, value, font=BLUE, fmt=None, note=None, key=False):
        asm[f"A{row}"] = lbl
        asm[f"A{row}"].font = BLACK
        c = asm[f"B{row}"]
        c.value = value
        c.font = font
        if fmt:
            c.number_format = fmt
        if key:
            c.fill = KEY_FILL
            c.border = BORDER
        if note:
            asm[f"C{row}"] = note
            asm[f"C{row}"].font = NOTE_F

    row = 4
    _section_row(asm, row, "MARKET DATA", "C"); row += 1
    A_PRICE = row; kv(row, "Current Share Price", a.price, BLUE, CUR2, f"Source: Yahoo Finance, retrieved {data.as_of}"); row += 1
    A_SHARES = row; kv(row, "Shares Outstanding (mm)", a.shares_out_mm, BLUE, NUM0, "Source: Yahoo Finance"); row += 1
    A_MKTCAP = row; kv(row, "Market Capitalization", f"=B{A_PRICE}*B{A_SHARES}", BLACK, CUR0); row += 1
    A_DEBT = row; kv(row, "Total Debt", a.total_debt_mm, BLUE, CUR0, "Source: Yahoo Finance (most recent reported quarter)"); row += 1
    A_CASH = row; kv(row, "Total Cash & Equivalents", a.total_cash_mm, BLUE, CUR0, "Source: Yahoo Finance"); row += 1
    A_NETDEBT = row; kv(row, "Net Debt", f"=B{A_DEBT}-B{A_CASH}", BLACK, CUR0); row += 1
    A_BETA = row; kv(row, "Beta", a.beta, BLUE, '0.00', "Source: Yahoo Finance"); row += 1

    row += 1
    _section_row(asm, row, "WACC BUILD", "C"); row += 1
    A_RF = row; kv(row, "Risk-Free Rate (10-Yr U.S. Treasury)", a.risk_free_rate, BLUE, PCT2, "Source: Yahoo Finance ^TNX"); row += 1
    A_ERP = row; kv(row, "Equity Risk Premium", a.equity_risk_premium, BLUE, PCT2, "Assumption: long-run U.S. equity risk premium. Edit as needed."); row += 1
    A_COE = row; kv(row, "Cost of Equity (CAPM)", f"=B{A_RF}+B{A_BETA}*B{A_ERP}", BLACK, PCT2); row += 1
    A_KD_PRE = row; kv(row, "Pre-Tax Cost of Debt", a.pretax_cost_of_debt, BLUE, PCT2, "Assumption: risk-free rate + estimated credit spread. Edit as needed."); row += 1
    A_TAX = row; kv(row, "Normalized Tax Rate", a.tax_rate, BLUE, PCT1, "Assumption: median historical effective tax rate (see Historicals), clipped to 10%-35%."); row += 1
    A_KD_POST = row; kv(row, "After-Tax Cost of Debt", f"=B{A_KD_PRE}*(1-B{A_TAX})", BLACK, PCT2); row += 1
    A_WE = row; kv(row, "Weight of Equity  E/(D+E)", f"=B{A_MKTCAP}/(B{A_MKTCAP}+B{A_DEBT})", BLACK, PCT1); row += 1
    A_WD = row; kv(row, "Weight of Debt  D/(D+E)", f"=B{A_DEBT}/(B{A_MKTCAP}+B{A_DEBT})", BLACK, PCT1); row += 1
    A_WACC = row; kv(row, "WACC", f"=B{A_WE}*B{A_COE}+B{A_WD}*B{A_KD_POST}", BLACK_B, PCT2, key=True); row += 1

    row += 1
    _section_row(asm, row, "DCF PROJECTION ASSUMPTIONS", "C"); row += 1
    A_BASEREV = row; kv(row, f"Base Revenue, {data.fiscal_years[-1]}", f"=Historicals!{last_hist_col}5", GREEN, CUR0); row += 1
    growth_rows = []
    for i, g in enumerate(a.revenue_growth):
        gr = row
        kv(row, f"Revenue Growth — {proj_years[i]}", g, BLUE, PCT1,
           "Analyst consensus" if i < 2 and (data.consensus_growth_y1 is not None) else "Assumption: faded toward terminal growth. Edit as needed.")
        growth_rows.append(gr)
        row += 1
    A_MARGIN0 = row; kv(row, f"{data.fiscal_years[-1]} EBIT Margin (start)", a.ebit_margin_start, BLUE, PCT1, "Based on most recent actual fiscal year."); row += 1
    A_MARGIN5 = row; kv(row, f"Target EBIT Margin ({proj_years[-1]})", a.ebit_margin_target, BLUE, PCT1, "Assumption: default holds margin flat at the latest actual level. Edit if you expect margin expansion/contraction."); row += 1
    A_DA = row; kv(row, "D&A (% of Revenue)", a.da_pct, BLUE, PCT1, "Based on most recent actual fiscal year D&A / Revenue."); row += 1
    A_CAPEX = row; kv(row, "Capital Expenditures (% of Revenue)", a.capex_pct, BLUE, PCT1, "Based on most recent actual fiscal year Capex / Revenue."); row += 1
    A_NWC = row; kv(row, "Increase in Net Working Capital (% of Revenue)", a.nwc_pct, BLUE, PCT1, "Assumption: net neutral by default. Edit if you have a view."); row += 1
    A_TERMG = row; kv(row, "Terminal Growth Rate", a.terminal_growth, BLUE, PCT1, "Assumption: approximates long-run nominal GDP growth. Must be less than WACC."); row += 1

    asm.freeze_panes = "A4"

    # ---------------------------------------------------------------
    # HISTORICALS
    # ---------------------------------------------------------------
    hist = wb.create_sheet("Historicals")
    hist.sheet_view.showGridLines = False
    for i, w in enumerate([34] + [14] * n_hist, start=1):
        hist.column_dimensions[get_column_letter(i)].width = w

    last_col_letter = get_column_letter(1 + n_hist)
    hist.merge_cells(f"A1:{last_col_letter}1")
    hist["A1"] = f"{data.name} ({data.ticker}) — Historical Financials ($ in millions)"
    hist["A1"].font = TITLE_F
    hist["A1"].fill = TITLE_FILL
    hist["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    hist.row_dimensions[1].height = 22
    hist.merge_cells(f"A2:{last_col_letter}2")
    hist["A2"] = f"Source: Yahoo Finance annual financial statements (via yfinance), retrieved {data.as_of}."
    hist["A2"].font = SUBTITLE_F

    hdr_row = 4
    hist[f"A{hdr_row}"] = "Fiscal Year"
    hist[f"A{hdr_row}"].font = HDR_F
    hist[f"A{hdr_row}"].fill = HDR_FILL
    for i, y in enumerate(data.fiscal_years):
        col = hist_cols[i]
        c = hist[f"{col}{hdr_row}"]
        c.value = y
        c.font = HDR_F
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")

    def data_row(row, lbl, vals, fmt=CUR0, font=BLUE, bold=False):
        hist[f"A{row}"] = lbl
        hist[f"A{row}"].font = BLACK_B if bold else BLACK
        for i, v in enumerate(vals):
            col = hist_cols[i]
            c = hist[f"{col}{row}"]
            c.value = v if v is not None else "n/a"
            c.font = font
            if v is not None:
                c.number_format = fmt

    r = hdr_row + 1
    H_REV = r; data_row(r, "Total Revenue", data.revenue); r += 1
    H_REVG = r
    hist[f"A{r}"] = "Revenue Growth %"; hist[f"A{r}"].font = BLACK
    hist[f"{hist_cols[0]}{r}"] = "n/a"; hist[f"{hist_cols[0]}{r}"].font = BLACK
    for i in range(1, n_hist):
        col, prev = hist_cols[i], hist_cols[i - 1]
        c = hist[f"{col}{r}"]
        c.value = f"={col}{H_REV}/{prev}{H_REV}-1"
        c.font = BLACK
        c.number_format = PCT1
    r += 1
    H_EBIT = r; data_row(r, "EBIT", data.ebit); r += 1
    H_EBITM = r
    hist[f"A{r}"] = "EBIT Margin %"; hist[f"A{r}"].font = BLACK
    for i in range(n_hist):
        col = hist_cols[i]
        c = hist[f"{col}{r}"]
        c.value = f"={col}{H_EBIT}/{col}{H_REV}"
        c.font = BLACK
        c.number_format = PCT1
    r += 1
    H_DA = r; data_row(r, "D&A", data.da); r += 1
    H_EBITDA = r
    hist[f"A{r}"] = "EBITDA (EBIT + D&A)"; hist[f"A{r}"].font = BLACK
    for i in range(n_hist):
        col = hist_cols[i]
        c = hist[f"{col}{r}"]
        c.value = f"={col}{H_EBIT}+{col}{H_DA}"
        c.font = BLACK
        c.number_format = CUR0
    r += 1
    H_CAPEX = r; data_row(r, "Capital Expenditures", data.capex); r += 1
    H_PRETAX = r; data_row(r, "Pretax Income", data.pretax_income); r += 1
    H_TAXPROV = r; data_row(r, "Tax Provision", data.tax_provision); r += 1
    H_TAXRATE = r
    hist[f"A{r}"] = "Effective Tax Rate"; hist[f"A{r}"].font = BLACK
    for i in range(n_hist):
        col = hist_cols[i]
        c = hist[f"{col}{r}"]
        c.value = f"={col}{H_TAXPROV}/{col}{H_PRETAX}"
        c.font = BLACK
        c.number_format = PCT1
    r += 1
    H_NI = r; data_row(r, "Net Income", data.net_income); r += 1
    H_OCF = r; data_row(r, "Operating Cash Flow", data.ocf); r += 1
    H_FCF = r
    hist[f"A{r}"] = "Free Cash Flow (OCF – Capex)"; hist[f"A{r}"].font = BLACK
    for i in range(n_hist):
        col = hist_cols[i]
        c = hist[f"{col}{r}"]
        c.value = f"={col}{H_OCF}-{col}{H_CAPEX}"
        c.font = BLACK
        c.number_format = CUR0
    r += 1
    H_TOTDEBT = r; data_row(r, "Total Debt (FY-end)", data.bs_total_debt); r += 1
    H_TOTEQ = r; data_row(r, "Total Equity (FY-end)", data.bs_total_equity); r += 1
    H_WC = r; data_row(r, "Working Capital (FY-end)", data.bs_working_capital); r += 1

    hist.freeze_panes = "B5"

    # ---------------------------------------------------------------
    # DCF MODEL
    # ---------------------------------------------------------------
    dcf = wb.create_sheet("DCF Model")
    dcf.sheet_view.showGridLines = False
    for i, w in enumerate([36, 14, 14, 14, 14, 14], start=1):
        dcf.column_dimensions[get_column_letter(i)].width = w

    dcf.merge_cells("A1:F1")
    dcf["A1"] = f"{data.name} ({data.ticker}) — Unlevered Free Cash Flow & DCF ($ in millions)"
    dcf["A1"].font = TITLE_F
    dcf["A1"].fill = TITLE_FILL
    dcf["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    dcf.row_dimensions[1].height = 22
    dcf.merge_cells("A2:F2")
    dcf["A2"] = f"5-year explicit forecast ({proj_years[0]}–{proj_years[-1]}) + Gordon Growth terminal value. All projection drivers link to the Assumptions tab."
    dcf["A2"].font = SUBTITLE_F

    hdr_row = 4
    dcf[f"A{hdr_row}"] = "Fiscal Year"
    dcf[f"A{hdr_row}"].font = HDR_F
    dcf[f"A{hdr_row}"].fill = HDR_FILL
    for i, y in enumerate(proj_years):
        col = proj_cols[i]
        c = dcf[f"{col}{hdr_row}"]
        c.value = y
        c.font = HDR_F
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")

    r = hdr_row + 1
    D_PERIOD = r
    dcf[f"A{r}"] = "Discount Period (t)"; dcf[f"A{r}"].font = BLACK
    for i, col in enumerate(proj_cols):
        c = dcf[f"{col}{r}"]; c.value = i + 1; c.font = BLACK; c.number_format = "0"
        c.alignment = Alignment(horizontal="center")
    r += 1

    D_GROWTH = r
    dcf[f"A{r}"] = "Revenue Growth %"; dcf[f"A{r}"].font = BLACK
    for i, col in enumerate(proj_cols):
        c = dcf[f"{col}{r}"]; c.value = f"=Assumptions!$B${growth_rows[i]}"; c.font = GREEN; c.number_format = PCT1
    r += 1

    D_REV = r
    dcf[f"A{r}"] = "Total Revenue"; dcf[f"A{r}"].font = BLACK_B
    for i, col in enumerate(proj_cols):
        prev = f"Assumptions!$B${A_BASEREV}" if i == 0 else f"{proj_cols[i-1]}{r}"
        c = dcf[f"{col}{r}"]; c.value = f"={prev}*(1+{col}{D_GROWTH})"; c.font = BLACK_B; c.number_format = CUR0
    r += 1

    D_MARGIN = r
    dcf[f"A{r}"] = "EBIT Margin % (linear ramp to target)"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]
        c.value = f"=Assumptions!$B${A_MARGIN0}+(Assumptions!$B${A_MARGIN5}-Assumptions!$B${A_MARGIN0})*({col}{D_PERIOD}/{N_YEARS})"
        c.font = GREEN; c.number_format = PCT1
    r += 1

    D_EBIT = r
    dcf[f"A{r}"] = "EBIT"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_REV}*{col}{D_MARGIN}"; c.font = BLACK; c.number_format = CUR0
    r += 1

    D_TAXRATE = r
    dcf[f"A{r}"] = "Tax Rate"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"=Assumptions!$B${A_TAX}"; c.font = GREEN; c.number_format = PCT1
    r += 1

    D_TAXES = r
    dcf[f"A{r}"] = "Less: Taxes on EBIT"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_EBIT}*{col}{D_TAXRATE}"; c.font = BLACK; c.number_format = CUR0
    r += 1

    D_NOPAT = r
    dcf[f"A{r}"] = "NOPAT (EBIAT)"; dcf[f"A{r}"].font = BLACK_B
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_EBIT}-{col}{D_TAXES}"; c.font = BLACK_B; c.number_format = CUR0
    r += 1

    D_DA = r
    dcf[f"A{r}"] = "Add: D&A"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_REV}*Assumptions!$B${A_DA}"; c.font = BLACK; c.number_format = CUR0
    r += 1

    D_CAPEX = r
    dcf[f"A{r}"] = "Less: Capital Expenditures"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_REV}*Assumptions!$B${A_CAPEX}"; c.font = BLACK; c.number_format = CUR0
    r += 1

    D_NWC = r
    dcf[f"A{r}"] = "Less: Increase in Net Working Capital"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_REV}*Assumptions!$B${A_NWC}"; c.font = BLACK; c.number_format = CUR0
    r += 1

    D_UFCF = r
    dcf[f"A{r}"] = "Unlevered Free Cash Flow"; dcf[f"A{r}"].font = BLACK_B
    for col in proj_cols:
        c = dcf[f"{col}{r}"]
        c.value = f"={col}{D_NOPAT}+{col}{D_DA}-{col}{D_CAPEX}-{col}{D_NWC}"
        c.font = BLACK_B; c.number_format = CUR0; c.border = TOPONLY
    r += 1

    D_DF = r
    dcf[f"A{r}"] = "Discount Factor @ WACC"; dcf[f"A{r}"].font = BLACK
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"=1/(1+Assumptions!$B${A_WACC})^{col}{D_PERIOD}"; c.font = GREEN; c.number_format = '0.000'
    r += 1

    D_PVFCF = r
    dcf[f"A{r}"] = "PV of Unlevered FCF"; dcf[f"A{r}"].font = BLACK_B
    for col in proj_cols:
        c = dcf[f"{col}{r}"]; c.value = f"={col}{D_UFCF}*{col}{D_DF}"; c.font = BLACK_B; c.number_format = CUR0
    r += 1

    r += 1
    _section_row(dcf, r, "VALUATION SUMMARY", "F"); r += 1

    def sv(row, lbl, value, font=BLACK, fmt=CUR0, key=False, note=None):
        dcf[f"A{row}"] = lbl; dcf[f"A{row}"].font = BLACK
        c = dcf[f"B{row}"]; c.value = value; c.font = font; c.number_format = fmt
        if key:
            c.fill = YELLOW_FILL
            c.border = BORDER
            c.font = Font(name=FONT_NAME, size=11, bold=True)
        if note:
            dcf[f"C{row}"] = note; dcf[f"C{row}"].font = NOTE_F

    D_SUMPV = r; sv(r, f"Sum of PV of Explicit FCF ({proj_years[0]}–{proj_years[-1]})", f"=SUM(B{D_PVFCF}:F{D_PVFCF})"); r += 1
    D_TERMG_L = r; sv(r, "Terminal Growth Rate", f"=Assumptions!$B${A_TERMG}", GREEN, PCT1); r += 1
    D_TV = r; sv(r, f"Terminal Value (Gordon Growth, end of {proj_years[-1]})", f"=F{D_UFCF}*(1+B{D_TERMG_L})/(Assumptions!$B${A_WACC}-B{D_TERMG_L})"); r += 1
    D_PVTV = r; sv(r, "PV of Terminal Value", f"=B{D_TV}*F{D_DF}"); r += 1
    D_EV = r; sv(r, "Enterprise Value", f"=B{D_SUMPV}+B{D_PVTV}", BLACK_B); dcf[f"B{r}"].border = TOPONLY; r += 1
    D_ND = r; sv(r, "Less: Net Debt", f"=-Assumptions!$B${A_NETDEBT}", GREEN); r += 1
    D_EQV = r; sv(r, "Equity Value", f"=B{D_EV}+B{D_ND}", BLACK_B); dcf[f"B{r}"].border = TOPONLY; r += 1
    D_SHARES = r; sv(r, "Shares Outstanding (mm)", f"=Assumptions!$B${A_SHARES}", GREEN, NUM0); r += 1
    D_IMPPRICE = r; sv(r, "Implied Value per Share", f"=B{D_EQV}/B{D_SHARES}", BLACK, CUR2, key=True); r += 1
    D_CURPRICE = r; sv(r, "Current Share Price", f"=Assumptions!$B${A_PRICE}", GREEN, CUR2); r += 1
    D_UPSIDE = r; sv(r, "Implied Upside / (Downside)", f"=B{D_IMPPRICE}/B{D_CURPRICE}-1", BLACK_B, PCT1); r += 1
    D_TVPCT = r; sv(r, "Terminal Value as % of Enterprise Value", f"=B{D_PVTV}/B{D_EV}", BLACK, PCT1,
                     note="Sanity check: if this exceeds ~75-80%, the valuation is highly sensitive to terminal assumptions."); r += 1

    dcf.freeze_panes = "B5"

    # ---------------------------------------------------------------
    # SENSITIVITY
    # ---------------------------------------------------------------
    sens = wb.create_sheet("Sensitivity")
    sens.sheet_view.showGridLines = False
    for i, w in enumerate([24, 14, 14, 14, 14, 14], start=1):
        sens.column_dimensions[get_column_letter(i)].width = w

    sens.merge_cells("A1:F1")
    sens["A1"] = f"{data.name} ({data.ticker}) — Sensitivity: Implied Share Price ($)"
    sens["A1"].font = TITLE_F
    sens["A1"].fill = TITLE_FILL
    sens["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    sens.row_dimensions[1].height = 22
    sens.merge_cells("A2:F2")
    sens["A2"] = ("Implied share price across WACC (rows) and terminal growth rate (columns). "
                  "Explicit-period unlevered FCF held constant at the DCF Model base case; only the discount rate and terminal growth vary.")
    sens["A2"].font = SUBTITLE_F
    sens.row_dimensions[2].height = 26
    sens["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    sens["B3"] = "WACC \\ Terminal g →"
    sens["B3"].font = HDR_F; sens["B3"].fill = HDR_FILL
    sens["B3"].alignment = Alignment(horizontal="center", wrap_text=True)

    g_cols = ["C", "D", "E", "F", "G"]
    g_offsets = [-0.01, -0.005, 0, 0.005, 0.01]
    for col, off in zip(g_cols, g_offsets):
        c = sens[f"{col}3"]
        c.value = f"=Assumptions!$B${A_TERMG}{'+' if off >= 0 else ''}{off}" if off != 0 else f"=Assumptions!$B${A_TERMG}"
        c.font = HDR_F; c.fill = HDR_FILL; c.number_format = PCT1
        c.alignment = Alignment(horizontal="center")

    w_rows = [4, 5, 6, 7, 8]
    w_offsets = [-0.01, -0.005, 0, 0.005, 0.01]
    for wrow, off in zip(w_rows, w_offsets):
        c = sens[f"B{wrow}"]
        c.value = f"=Assumptions!$B${A_WACC}{'+' if off >= 0 else ''}{off}" if off != 0 else f"=Assumptions!$B${A_WACC}"
        c.font = HDR_F; c.fill = HDR_FILL; c.number_format = PCT2
        c.alignment = Alignment(horizontal="center")

    for wrow in w_rows:
        for col in g_cols:
            c = sens[f"{col}{wrow}"]
            c.value = (
                f"=IFERROR((SUMPRODUCT('DCF Model'!$B${D_UFCF}:$F${D_UFCF},"
                f"1/(1+$B{wrow})^('DCF Model'!$B${D_PERIOD}:$F${D_PERIOD}))"
                f"+('DCF Model'!$F${D_UFCF}*(1+{col}$3)/($B{wrow}-{col}$3))"
                f"/(1+$B{wrow})^'DCF Model'!$F${D_PERIOD}"
                f"-Assumptions!$B${A_NETDEBT})/Assumptions!$B${A_SHARES},\"n/m\")"
            )
            c.font = BLACK; c.number_format = CUR2
            c.alignment = Alignment(horizontal="center")
            c.border = BORDER
            if wrow == 6 and col == "E":
                c.fill = YELLOW_FILL
                c.font = Font(name=FONT_NAME, size=10, bold=True)

    sens.freeze_panes = "C4"

    r2 = 10
    sens[f"A{r2}"] = "Current Share Price"; sens[f"A{r2}"].font = BLACK
    sens[f"B{r2}"] = f"=Assumptions!$B${A_PRICE}"; sens[f"B{r2}"].font = GREEN; sens[f"B{r2}"].number_format = CUR2
    r2 += 1
    if data.analyst_target_mean:
        sens[f"A{r2}"] = "Analyst Avg. Price Target (for reference)"; sens[f"A{r2}"].font = BLACK
        sens[f"B{r2}"] = data.analyst_target_mean; sens[f"B{r2}"].font = BLUE; sens[f"B{r2}"].number_format = CUR2
        sens[f"C{r2}"] = f"Source: Yahoo Finance, {data.analyst_count or 'n/a'} analysts, retrieved {data.as_of}"
        sens[f"C{r2}"].font = NOTE_F

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
