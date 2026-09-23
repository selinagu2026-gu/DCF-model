# DCF Modle

A small Streamlit web app that builds a discounted cash flow (DCF)
valuation model for any public company, using live data pulled from
Yahoo Finance.

## What it does

1. You type a stock ticker (e.g. `DIS`, `AAPL`, `MSFT`).
2. It pulls, via the `yfinance` library:
   - Current price, shares outstanding, market cap, beta, total debt/cash
   - The last 4-5 years of annual revenue, EBIT, D&A, capex, taxes, and
     free cash flow
   - The live 10-Year US Treasury yield (risk-free rate)
   - Analyst consensus revenue growth estimates, where available
3. It derives a full set of starting DCF assumptions from that data
   (WACC build, revenue growth fade, margins, D&A%, capex%, tax rate),
   which you can then edit in the sidebar — the whole model recalculates
   live as you move any input.
4. It shows historicals, the WACC build, the 5-year unlevered free cash
   flow projection, the implied share price, and a WACC × terminal
   growth sensitivity grid.
5. You can download the same model as a formatted, formula-driven Excel
   workbook (5 tabs: Cover, Assumptions, Historicals, DCF Model,
   Sensitivity) that keeps recalculating if you edit it further in Excel.

## Project layout

```
app.py                 Streamlit UI — run this
dcf/
  data_fetcher.py       Talks to Yahoo Finance (yfinance), returns a CompanyData object
  dcf_engine.py          Pure DCF math: WACC, FCF projection, terminal value, sensitivity
  excel_export.py        Builds the downloadable .xlsx workbook from the same data/assumptions
requirements.txt
```

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

This opens the app in your browser (usually http://localhost:8501).

## Notes & limitations

- Data comes from Yahoo Finance via `yfinance`, an unofficial wrapper
  around Yahoo's public endpoints. Fields occasionally go missing or
  get renamed by Yahoo — if a ticker's historicals look incomplete,
  that's the usual cause.
- Works best for standard operating companies. ETFs, mutual funds,
  indices, and some financial companies (banks, insurers) don't report
  the line items a standard DCF needs and won't fetch cleanly.
- All forward-looking assumptions (growth beyond the analyst-covered
  years, margin targets, terminal growth, equity risk premium, cost of
  debt) are reasonable *defaults*, not predictions — review and adjust
  them before relying on the output.
- Educational purposes only. Not investment advice.
