import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter, MonthLocator
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter
import uuid

# Streamlit page configuration
st.set_page_config(layout="wide", page_title="Black-Scholes Volatility Calculator")
st.image("Qapita1.png", width=150)
st.title("Black-Scholes Volatility Calculator")
st.markdown("---")

# Sidebar inputs
st.sidebar.header("Input Parameters")
num_tickers = st.sidebar.number_input("Number of Companies", min_value=1, max_value=10, value=2)
tickers = [st.sidebar.text_input(f"Enter Ticker {i+1} (e.g., INFY.NS)", "").strip().upper() for i in range(num_tickers)]
tickers = [t for t in tickers if t]  # Remove empty tickers
start_date = st.sidebar.date_input("Start Date", datetime.today() - timedelta(days=365))
end_date = st.sidebar.date_input("End Date", datetime.today())

# Validate inputs
if start_date >= end_date:
    st.sidebar.error("End date must be after start date.")
    st.stop()

# Display model explanation
st.markdown("### 📘 Understanding Volatility in the Black-Scholes Model")
st.markdown(
    """
The **Black-Scholes Model** is used to price options, with **volatility (σ)** as a key input, representing stock price fluctuations. 
Volatility is calculated as:
"""
)
st.latex(r"\text{Annualized Volatility} = \sigma = \text{Std Dev of Daily Returns} \times \sqrt{252}")
st.markdown(
    """
- **Std Dev of Daily Returns**: Daily price movement percentage.
- **252**: Typical trading days in a year.
- **Data Source**: Adjusted close prices (if available) or close prices, accounting for splits and dividends.
"""
)

@st.cache_data
def fetch_stock_data(ticker, start, end):
    """Fetch stock data from yfinance with error handling."""
    try:
        data = yf.download(ticker, start=start, end=end, progress=False)
        if data.empty:
            return None
        if 'Adj Close' in data.columns and not data['Adj Close'].dropna().empty:
            data = data[['Adj Close']].rename(columns={'Adj Close': 'Price'})
        elif 'Close' in data.columns and not data['Close'].dropna().empty:
            data = data[['Close']].rename(columns={'Close': 'Price'})
        else:
            return None
        return data
    except Exception as e:
        st.warning(f"Error fetching data for {ticker}: {str(e)}")
        return None

def process_stock_data(data, ticker):
    """Process stock data for volatility calculation and visualization."""
    data['Price'] = data['Price'].round(2)
    data['Daily % Change'] = data['Price'].pct_change() * 100
    data.dropna(inplace=True)
    data.reset_index(inplace=True)
    data['Date'] = pd.to_datetime(data['Date'])
    
    std_dev = data['Daily % Change'].std()
    annual_vol = std_dev * np.sqrt(252)
    days = len(data)
    
    return data, {
        "Ticker": ticker,
        "Std Dev (%)": std_dev,
        "No. of Days": days,
        "Period": f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}",
        "Annualized Volatility (%)": annual_vol
    }

def plot_daily_change(data, ticker):
    """Generate and display daily % change plot."""
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(data['Date'], data['Daily % Change'], label=f'{ticker} % Change', color='royalblue')
    ax.set_xlabel("Date", fontsize=12)
    ax.set_ylabel("Daily % Change (%)", fontsize=12)
    ax.set_title(f"{ticker} Daily % Change", fontsize=14)
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.xaxis.set_major_locator(MonthLocator())
    ax.xaxis.set_major_formatter(DateFormatter("%b '%y"))
    plt.xticks(rotation=45)
    plt.legend()
    st.pyplot(fig)

def generate_excel(data_dict, summary_dict):
    """Generate Excel file with summary and ticker sheets."""
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # Formats
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_size': 12})
    cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 11})
    pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'border': 1, 'align': 'center', 'font_size': 11})
    date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center', 'font_size': 11})
    price_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center', 'font_size': 11})

    # Summary sheet
    sheet = workbook.add_worksheet("Summary")
    headers = ["Ticker", "Std Dev (%)", "No. of Days", "Period", "Annualized Volatility (%)"]
    for col, header in enumerate(headers):
        sheet.write(0, col, header, header_fmt)
    
    for row, item in enumerate(summary_dict.values(), start=1):
        sheet.write(row, 0, item["Ticker"], cell_fmt)
        sheet.write_number(row, 1, item["Std Dev (%)"], pct_fmt)
        sheet.write_number(row, 2, item["No. of Days"], cell_fmt)
        sheet.write(row, 3, item["Period"], cell_fmt)
        sheet.write_number(row, 4, item["Annualized Volatility (%)"], pct_fmt)
    
    sheet.set_column(0, 0, 15)
    sheet.set_column(1, 1, 12)
    sheet.set_column(2, 2, 12)
    sheet.set_column(3, 3, 25)
    sheet.set_column(4, 4, 20)

    # Ticker sheets
    for ticker, df in data_dict.items():
        ws = workbook.add_worksheet(ticker[:31])  # Excel sheet name limit
        headers = ["Date", "Price", "Daily % Change"]
        for col, header in enumerate(headers):
            ws.write(0, col, header, header_fmt)
        
        for row, row_data in enumerate(df.itertuples(name=None), start=1):
            ws.write_datetime(row, 0, pd.to_datetime(row_data[1]), date_fmt)  # Date is first column (index 1)
            ws.write_number(row, 1, row_data[2], price_fmt)  # Price is second column (index 2)
            ws.write_number(row, 2, row_data[3], pct_fmt)  # Daily % Change is third column (index 3)
        
        ws.set_column(0, 0, 15)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 15)

    workbook.close()
    output.seek(0)
    return output

# Main logic
if st.sidebar.button("Calculate Volatility") and tickers:
    with st.spinner("Fetching and processing data..."):
        stock_data_dict = {}
        volatility_summary = {}
        
        for ticker in tickers:
            data = fetch_stock_data(ticker, start_date, end_date)
            if data is None:
                st.warning(f"No valid price data found for {ticker}. Skipping.")
                continue
            
            processed_data, summary = process_stock_data(data, ticker)
            stock_data_dict[ticker] = processed_data
            volatility_summary[ticker] = summary
            
            # Display data
            st.subheader(f"Data for {ticker}", anchor=None)
            display_df = processed_data.copy()
            display_df['Date'] = display_df['Date'].dt.strftime('%d/%m/%Y')
            display_df['Daily % Change'] = display_df['Daily % Change'].map('{:.2f}%'.format)
            display_df['Price'] = display_df['Price'].map('{:.2f}'.format)
            st.dataframe(display_df[['Date', 'Price', 'Daily % Change']], use_container_width=True, height=300)
            
            # Plot
            st.write("**Daily % Change Chart**")
            plot_daily_change(processed_data, ticker)
        
        if volatility_summary:
            st.markdown("### 📊 Volatility Summary")
            summary_df = pd.DataFrame(list(volatility_summary.values()))
            summary_df["Std Dev (%)"] = summary_df["Std Dev (%)"].map("{:.2f}%".format)
            summary_df["Annualized Volatility (%)"] = summary_df["Annualized Volatility (%)"].map("{:.2f}%".format)
            st.table(summary_df[["Ticker", "Std Dev (%)", "No. of Days", "Period", "Annualized Volatility (%)"]])
            
            # Excel download
            excel_file = generate_excel(stock_data_dict, volatility_summary)
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_file,
                file_name="volatility_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=str(uuid.uuid4())
            )
        
        st.markdown("---")
        st.markdown("**Note**: Adjusted closing prices are used when available to account for splits, dividends, and corporate actions. If unavailable, closing prices are used.")
else:
    st.info("Enter valid tickers and click 'Calculate Volatility' to see results.")