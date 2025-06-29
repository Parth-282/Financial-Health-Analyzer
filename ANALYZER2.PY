import yfinance as yf
import pandas as pd
import numpy as np
import datetime
import xlsxwriter

# --- Safe data access helpers ---
def safe_get(df, col, date):
    return df.at[date, col] if col in df.columns and date in df.index else np.nan

def get_first_valid_value(df, date, columns):
    for col in columns:
        val = safe_get(df, col, date)
        if not pd.isna(val):
            return val
    return np.nan

def find_valid_year(balance_sheet):
    for date in balance_sheet.index:
        equity = get_first_valid_value(balance_sheet, date, ['Total Stockholder Equity', 'Stockholders Equity', 'Common Stock Equity'])
        liabilities = get_first_valid_value(balance_sheet, date, ['Total Liab', 'Total Liabilities Net Minority Interest'])
        current_assets = get_first_valid_value(balance_sheet, date, ['Total Current Assets', 'Current Assets'])
        current_liabilities = get_first_valid_value(balance_sheet, date, ['Total Current Liabilities', 'Current Liabilities'])
        if not any(pd.isna(x) for x in [equity, liabilities, current_assets, current_liabilities]):
            return date
    return balance_sheet.index[0]

# --- Analyze one company ---
def analyze_company(ticker):
    try:
        stock = yf.Ticker(ticker)
        bs = stock.balance_sheet.T
        is_ = stock.financials.T
        info = stock.info
        date = find_valid_year(bs)

        revenue = get_first_valid_value(is_, date, ['Total Revenue'])
        net_income = get_first_valid_value(is_, date, ['Net Income'])
        ebit = get_first_valid_value(is_, date, ['EBIT', 'Operating Income'])
        interest_expense = abs(get_first_valid_value(is_, date, ['Interest Expense']))

        total_assets = get_first_valid_value(bs, date, ['Total Assets'])
        equity = get_first_valid_value(bs, date, ['Total Stockholder Equity', 'Stockholders Equity', 'Common Stock Equity'])
        liabilities = get_first_valid_value(bs, date, ['Total Liab', 'Total Liabilities Net Minority Interest'])
        current_assets = get_first_valid_value(bs, date, ['Total Current Assets', 'Current Assets'])
        current_liabilities = get_first_valid_value(bs, date, ['Total Current Liabilities', 'Current Liabilities'])

        pe_ratio = info.get('trailingPE', np.nan)
        eps = info.get('trailingEps', np.nan)
        current_price = info.get('currentPrice', np.nan)
        fair_value = eps * 25 if not pd.isna(eps) else np.nan  # Assuming sector PE = 25
        recommendation = (
            "Undervalued" if not pd.isna(current_price) and fair_value > current_price * 1.1 else
            "Overvalued" if not pd.isna(current_price) and fair_value < current_price * 0.9 else
            "Fairly Valued"
        )

        ratios = {
            'Net Profit Margin (%)': (net_income / revenue * 100) if revenue else np.nan,
            'ROA (%)': (net_income / total_assets * 100) if total_assets else np.nan,
            'ROE (%)': (net_income / equity * 100) if equity else np.nan,
            'Current Ratio': (current_assets / current_liabilities) if current_liabilities else np.nan,
            'Debt to Equity': (liabilities / equity) if equity else np.nan,
            'Interest Coverage Ratio': (ebit / interest_expense) if interest_expense else np.nan,
            'P/E Ratio': pe_ratio,
            'P/B Ratio': info.get('priceToBook', np.nan),
            'EPS (TTM)': eps,
            'Current Price': current_price,
            'Fair Value Estimate': fair_value,
            'Valuation Recommendation': recommendation
        }

        return pd.Series(ratios, name=ticker)
    except Exception as e:
        print(f"Error processing {ticker}: {e}")
        return pd.Series(name=ticker)

# --- Tickers to analyze ---
tickers = ['TCS.NS', 'INFY.NS', 'WIPRO.NS', 'HCLTECH.NS']

# --- Run analysis ---
df = pd.DataFrame([analyze_company(tkr) for tkr in tickers]).round(2).T

# Split and clean numeric part
df_numeric = df.drop('Valuation Recommendation')
df_numeric = df_numeric.apply(pd.to_numeric, errors='coerce')
df_cleaned = pd.concat([df_numeric, df.loc[['Valuation Recommendation']]])

# Rankings
ranking_metrics = ['ROE (%)', 'Net Profit Margin (%)', 'ROA (%)']
ranking_df = df_numeric.loc[ranking_metrics].rank(axis=1, ascending=False).mean().sort_values()
ranking_df = pd.DataFrame({'Average Rank (lower is better)': ranking_df})

# Summary Insights
insights = [
    f"\U0001F3C6 Best ROE: {df_numeric.loc['ROE (%)'].idxmax()} ({df_numeric.loc['ROE (%)'].max()}%)",
    f"\U0001F3C6 Best Net Margin: {df_numeric.loc['Net Profit Margin (%)'].idxmax()} ({df_numeric.loc['Net Profit Margin (%)'].max()}%)",
    f"⚠️ Highest D/E: {df_numeric.loc['Debt to Equity'].idxmax()} ({df_numeric.loc['Debt to Equity'].max()})"
]

# --- Terminal Output ---
print("\n\U0001F4CA SUMMARY INSIGHTS")
for insight in insights:
    print("•", insight)

print("\n\U0001F3C5 COMPANY RANKINGS (Lower rank = better performance on ROE, Margin, ROA):")
print(ranking_df)

print("\n\U0001F4C8 VALUATION RECOMMENDATIONS:")
recommendation_table = df_cleaned.loc[['Current Price', 'Fair Value Estimate', 'Valuation Recommendation']]
print(recommendation_table.T)

# --- Save Excel output ---
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
filename = f"Financial_Health_Analyzer_{timestamp}.xlsx"
workbook = xlsxwriter.Workbook(filename)

header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
number_format = workbook.add_format({'num_format': '0.00', 'border': 1})
currency_format = workbook.add_format({'num_format': '₹#,##0', 'border': 1})
default_format = workbook.add_format({'border': 1})
highlight_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
highlight_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

# Sheet 1
ws_ratios = workbook.add_worksheet('Financial Ratios')
ws_ratios.freeze_panes(1, 1)
ws_ratios.write(0, 0, "Metric", header_format)
for col, company in enumerate(df_cleaned.columns, start=1):
    ws_ratios.write(0, col, company, header_format)

for row, metric in enumerate(df_cleaned.index, start=1):
    ws_ratios.write(row, 0, metric, header_format)
    for col, company in enumerate(df_cleaned.columns, start=1):
        val = df_cleaned.at[metric, company]
        fmt = (
            percent_format if '(%' in metric else
            currency_format if 'Price' in metric or 'Value' in metric else
            default_format
        )
        if metric == 'Valuation Recommendation':
            fmt = highlight_green if val == 'Undervalued' else highlight_red if val == 'Overvalued' else default_format
        ws_ratios.write(row, col, val, fmt)
ws_ratios.set_column(0, len(df_cleaned.columns), 18)

# Sheet 2: Summary
ws_summary = workbook.add_worksheet('Summary Insights')
ws_summary.write(0, 0, "Summary Insights", header_format)
for i, insight in enumerate(insights, start=2):
    ws_summary.write(i, 0, insight)

# Sheet 3: Rankings
ws_rank = workbook.add_worksheet('Company Rankings')
ws_rank.write(0, 0, "Company", header_format)
ws_rank.write(0, 1, "Average Rank (lower is better)", header_format)
for i, (company, score) in enumerate(ranking_df['Average Rank (lower is better)'].items(), start=1):
    ws_rank.write(i, 0, company)
    ws_rank.write(i, 1, score, number_format)
ws_rank.set_column(0, 1, 30)

# Sheet 4: Charts
ws_chart = workbook.add_worksheet('Charts')
chart_metrics = ['ROE (%)', 'Net Profit Margin (%)', 'Debt to Equity']
row_offset = 0
for metric in chart_metrics:
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': metric,
        'categories': ['Financial Ratios', 0, 1, 0, len(df.columns)],
        'values':     ['Financial Ratios', df.index.get_loc(metric) + 1, 1, df.index.get_loc(metric) + 1, len(df.columns)],
    })
    chart.set_title({'name': metric})
    chart.set_x_axis({'name': 'Company'})
    chart.set_y_axis({'name': metric})
    ws_chart.insert_chart(row_offset, 0, chart, {'x_offset': 25, 'y_offset': 10})
    row_offset += 15

workbook.close()
print(f"\n✅ Excel file saved as: {filename}")
