import yfinance as yf
import pandas as pd
import numpy as np
import datetime
import xlsxwriter
import matplotlib.pyplot as plt
import os


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
        fair_value = eps * 25 if not pd.isna(eps) else np.nan

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

# --- Clean numeric data ---
df_numeric = df.drop('Valuation Recommendation')
df_numeric = df_numeric.apply(pd.to_numeric, errors='coerce')
df_cleaned = pd.concat([df_numeric, df.loc[['Valuation Recommendation']]])

# --- Rankings ---
ranking_metrics = ['ROE (%)', 'Net Profit Margin (%)', 'ROA (%)']
ranking_df = df_numeric.loc[ranking_metrics].rank(axis=1, ascending=False).mean().sort_values()
ranking_df = pd.DataFrame({'Average Rank (lower is better)': ranking_df})

# --- Summary Insights ---
insights = [
    f"🏆 Best ROE: {df_numeric.loc['ROE (%)'].idxmax()} ({df_numeric.loc['ROE (%)'].max()}%)",
    f"🏆 Best Net Margin: {df_numeric.loc['Net Profit Margin (%)'].idxmax()} ({df_numeric.loc['Net Profit Margin (%)'].max()}%)",
    f"⚠️ Highest D/E: {df_numeric.loc['Debt to Equity'].idxmax()} ({df_numeric.loc['Debt to Equity'].max()})"
]

# --- Terminal Output ---
print("\n SUMMARY INSIGHTS")
for insight in insights:
    print("•", insight)

print("\n COMPANY RANKINGS:")
print(ranking_df)

print("\n VALUATION RECOMMENDATIONS:")
recommendation_table = df_cleaned.loc[['Current Price', 'Fair Value Estimate', 'Valuation Recommendation']]
print(recommendation_table.T)


# --- Save charts as PNG ---
charts_folder = 'Charts_Output'
os.makedirs(charts_folder, exist_ok=True)

chart_metrics = [
    'ROE (%)', 'Net Profit Margin (%)', 'Debt to Equity',
    'Current Ratio', 'Interest Coverage Ratio', 'P/E Ratio'
]

# --- Single-metric bar charts ---
for metric in chart_metrics:
    plt.figure(figsize=(8, 6))
    data = df_numeric.loc[metric]
    bars = plt.bar(data.index, data.values, color='skyblue', edgecolor='black')

    plt.title(f'{metric} Comparison')
    plt.ylabel(metric)
    plt.xlabel('Company')
    plt.xticks(rotation=45)

    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + (yval * 0.02 if yval != 0 else 0.02),
                 round(yval, 2), ha='center', fontsize=9)

    plt.tight_layout()
    filename = os.path.join(charts_folder, f'{metric.replace(" ", "_").replace("/", "-")}.png')
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"✅ Chart saved as {filename}")

# --- Fair Value vs Current Price chart ---
plt.figure(figsize=(8, 6))
data = df_numeric.loc[['Current Price', 'Fair Value Estimate']]
x = np.arange(len(data.columns))
width = 0.35

fig, ax = plt.subplots(figsize=(8, 6))
bars1 = ax.bar(x - width/2, data.loc['Current Price'], width, label='Current Price', color='lightcoral')
bars2 = ax.bar(x + width/2, data.loc['Fair Value Estimate'], width, label='Fair Value', color='mediumseagreen')

ax.set_ylabel('Price (₹)')
ax.set_title('Fair Value vs Current Price')
ax.set_xticks(x)
ax.set_xticklabels(data.columns, rotation=45)
ax.legend()

for bars in [bars1, bars2]:
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height:.2f}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 5),
                    textcoords="offset points",
                    ha='center', va='bottom')

plt.tight_layout()
fv_filename = os.path.join(charts_folder, 'Fair_Value_vs_Current_Price.png')
plt.savefig(fv_filename, dpi=300, bbox_inches='tight')
plt.close()
print(f"✅ Chart saved as {fv_filename}")


# --- Save Excel output ---
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
filename = f"Financial_Health_Analyzer_{timestamp}.xlsx"
workbook = xlsxwriter.Workbook(filename)

# Formatting
header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
number_format = workbook.add_format({'num_format': '0.00', 'border': 1})
currency_format = workbook.add_format({'num_format': '₹#,##0', 'border': 1})
default_format = workbook.add_format({'border': 1})
highlight_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
highlight_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

# Sheet 1: Financial Ratios
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

# Fair Value vs Current Price Chart
chart = workbook.add_chart({'type': 'column'})
chart.add_series({
    'name': 'Current Price',
    'categories': ['Financial Ratios', 0, 1, 0, len(df.columns)],
    'values':     ['Financial Ratios', df.index.get_loc('Current Price') + 1, 1, df.index.get_loc('Current Price') + 1, len(df.columns)],
})
chart.add_series({
    'name': 'Fair Value Estimate',
    'categories': ['Financial Ratios', 0, 1, 0, len(df.columns)],
    'values':     ['Financial Ratios', df.index.get_loc('Fair Value Estimate') + 1, 1, df.index.get_loc('Fair Value Estimate') + 1, len(df.columns)],
})
chart.set_title({'name': 'Fair Value vs Current Price'})
chart.set_x_axis({'name': 'Company'})
chart.set_y_axis({'name': 'Price (₹)'})
chart.set_style(10)
ws_chart.insert_chart(row_offset, 0, chart, {'x_offset': 25, 'y_offset': 10})
row_offset += 15

workbook.close()
print(f"\n✅ Excel file saved as: {filename}")
