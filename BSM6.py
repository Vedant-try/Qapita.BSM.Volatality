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
st.set_page_config(layout="wide", page_title="ESOP Valuation Calculator")
st.image("Qapita1.png", width=150)  # Assuming logo is available
st.title("ESOP Valuation Calculator")
st.markdown("---")

# Model explanation (from BSM6.py)
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

# Sidebar inputs
st.sidebar.header("Input Parameters")
company_name = st.sidebar.text_input("Company Name", "Example Inc.")
ticker = st.sidebar.text_input("Enter Ticker (e.g., INFY.NS)", "").strip().upper()
grant_date = st.sidebar.date_input("Grant Date", datetime.today() - timedelta(days=365))
num_vests = st.sidebar.number_input("Number of Vests", min_value=1, max_value=10, value=3)
period_unit = st.sidebar.selectbox("Period Denomination", ["Years", "Quarters", "Months", "Days"], index=0)

# Conversion factors for period denominations
period_factors = {"Years": 365, "Quarters": 365/4, "Months": 365/12, "Days": 1}
period_divisor = period_factors[period_unit]

# Maximum date supported by pandas (2262-04-11)
MAX_DATE = datetime(2262, 4, 11).date()

# Period conversion reference data (defined in global scope)
period_ref_data = [
    {"Period": "1 Year", "Days": 365},
    {"Period": "2 Years", "Days": 730},
    {"Period": "3 Years", "Days": 1095}
]

# Validate inputs
if not ticker:
    st.sidebar.error("Please enter a valid ticker.")
    st.stop()
if grant_date > datetime.today().date():
    st.sidebar.error("Grant date cannot be in the future.")
    st.stop()

# Vesting table
st.markdown("### Vesting Schedule")
col1, col2 = st.columns([3, 3])
with col1:
    st.markdown(f"Enter vesting details (Periods in {period_unit}):")
    vesting_data = []
    colors = ["#DDEBF7", "#E6F3E6", "#FFE6E6", "#E6E6FA", "#FFFACD"]  # Colors for distinguishing vests
    for i in range(num_vests):
        st.markdown(f"#### Vesting {i+1}", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(f"<div style='background-color:{colors[i % len(colors)]};padding:10px;border-radius:5px;'>", unsafe_allow_html=True)
            vesting_row = {}
            vesting_row["Vesting"] = f"Vesting {i+1}"
            
            # Input choice: Vesting Date or Vesting Period
            input_method = st.radio(f"Input method for Vesting {i+1}", ["Vesting Date", f"Vesting Period ({period_unit})"], key=f"input_method_{i}")
            
            if input_method == "Vesting Date":
                vesting_date = st.date_input(f"Vesting Date {i+1} (DD/MM/YYYY)", 
                                            min_value=grant_date,
                                            max_value=MAX_DATE,
                                            key=f"vest_date_{i}")
                vesting_row["Vesting Date"] = vesting_date.strftime('%d/%m/%Y')
                vesting_period_days = (vesting_date - grant_date).days
                vesting_row["Vesting Period (Days)"] = vesting_period_days
                vesting_row["Vesting Period"] = vesting_period_days / period_divisor
            else:
                # Calculate max vesting period to keep vesting_date <= MAX_DATE
                max_days = (MAX_DATE - grant_date).days
                max_period = max_days / period_divisor
                vesting_period = st.number_input(f"Vesting Period for Vesting {i+1} ({period_unit})", 
                                                min_value=0.0, 
                                                max_value=max_period,
                                                value=min(1.0, max_period),
                                                step=0.1, 
                                                key=f"vest_period_{i}")
                vesting_row["Vesting Period"] = vesting_period
                vesting_row["Vesting Period (Days)"] = vesting_period * period_divisor
                vesting_date = grant_date + timedelta(days=int(vesting_row["Vesting Period (Days)"]))
                if vesting_date > MAX_DATE:
                    st.error(f"Vesting {i+1}: Vesting period results in a date beyond 11/04/2262, which is not supported. Please reduce the vesting period to {max_period:.2f} {period_unit} or less.")
                    st.stop()
                vesting_row["Vesting Date"] = vesting_date.strftime('%d/%m/%Y')
            
            vesting_row["Weight (%)"] = st.number_input(f"Weight (%) for Vesting {i+1}", 
                                                      min_value=0.0, max_value=100.0, 
                                                      value=100.0/num_vests, step=0.1, key=f"weight_{i}")
            
            # Input choice: Exercise End Date or Exercise Validity
            exercise_input_method = st.radio(f"Exercise input for Vesting {i+1}", ["Exercise End Date", f"Exercise Validity ({period_unit})"], 
                                            key=f"exercise_input_method_{i}")
            
            if exercise_input_method == "Exercise End Date":
                exercise_end_date = st.date_input(f"Exercise End Date {i+1} (DD/MM/YYYY)", 
                                                min_value=vesting_date,
                                                max_value=MAX_DATE,
                                                value=vesting_date,  # Set default to vesting_date
                                                key=f"exercise_end_date_{i}")
                vesting_row["Exercise End Date"] = exercise_end_date.strftime('%d/%m/%Y')
                try:
                    vesting_date_dt = pd.to_datetime(vesting_row["Vesting Date"], format='%d/%m/%Y').date()
                    exercise_validity_days = (exercise_end_date - vesting_date_dt).days
                    vesting_row["Exercise Validity (Days)"] = exercise_validity_days
                    vesting_row["Exercise Validity"] = exercise_validity_days / period_divisor
                except Exception as e:
                    st.error(f"Vesting {i+1}: Invalid date conversion for vesting date {vesting_row['Vesting Date']}. Error: {str(e)}")
                    st.stop()
            else:
                # Calculate max exercise validity to keep exercise_end_date <= MAX_DATE
                max_exercise_days = (MAX_DATE - vesting_date).days
                max_exercise_period = max_exercise_days / period_divisor
                exercise_validity = st.number_input(f"Exercise Validity for Vesting {i+1} ({period_unit})", 
                                                  min_value=0.0, 
                                                  max_value=max_exercise_period,
                                                  value=min(1.0, max_exercise_period),
                                                  step=0.1, 
                                                  key=f"exercise_validity_{i}")
                vesting_row["Exercise Validity"] = exercise_validity
                vesting_row["Exercise Validity (Days)"] = exercise_validity * period_divisor
                try:
                    vesting_date_dt = pd.to_datetime(vesting_row["Vesting Date"], format='%d/%m/%Y').date()
                    exercise_end_date = vesting_date_dt + timedelta(days=int(vesting_row["Exercise Validity (Days)"]))
                    if exercise_end_date > MAX_DATE:
                        st.error(f"Vesting {i+1}: Exercise validity results in a date beyond 11/04/2262, which is not supported. Please reduce the exercise validity to {max_exercise_period:.2f} {period_unit} or less.")
                        st.stop()
                    vesting_row["Exercise End Date"] = exercise_end_date.strftime('%d/%m/%Y')
                except Exception as e:
                    st.error(f"Vesting {i+1}: Invalid date conversion for vesting date {vesting_row['Vesting Date']}. Error: {str(e)}")
                    st.stop()
            
            vesting_row["Min Life (Days)"] = vesting_row["Vesting Period (Days)"]
            vesting_row["Min Life"] = vesting_row["Vesting Period"]
            vesting_row["Max Life (Days)"] = vesting_row["Vesting Period (Days)"] + vesting_row["Exercise Validity (Days)"]
            vesting_row["Max Life"] = vesting_row["Vesting Period"] + vesting_row["Exercise Validity"]
            vesting_row["Avg Life (Days)"] = (vesting_row["Min Life (Days)"] + vesting_row["Max Life (Days)"]) / 2
            vesting_row["Avg Life"] = vesting_row["Min Life"] + vesting_row["Exercise Validity"] / 2
            volatility_start_date = grant_date - timedelta(days=int(vesting_row["Avg Life (Days)"]))
            vesting_row["Volatility Start Date"] = volatility_start_date.strftime('%d/%m/%Y')
            vesting_row["Volatility End Date"] = grant_date.strftime('%d/%m/%Y')
            vesting_row["Annualized Volatility (%)"] = 0.0  # Initialize volatility
            vesting_data.append(vesting_row)
            st.markdown("</div>", unsafe_allow_html=True)

# Convert vesting data to DataFrame
vesting_df = pd.DataFrame(vesting_data)

# Display vesting table
with col2:
    st.markdown(f"**Vesting Schedule (Periods in {period_unit})**")
    display_df = vesting_df[["Vesting", "Vesting Date", "Vesting Period", "Weight (%)", 
                            "Exercise Validity", "Exercise End Date", 
                            "Min Life", "Max Life", "Avg Life", 
                            "Volatility Start Date", "Volatility End Date"]]
    display_df = display_df.copy()
    display_df["Vesting Period"] = display_df["Vesting Period"].map(f"{{:.2f}} {period_unit}".format)
    display_df["Exercise Validity"] = display_df["Exercise Validity"].map(f"{{:.2f}} {period_unit}".format)
    display_df["Min Life"] = display_df["Min Life"].map(f"{{:.2f}} {period_unit}".format)
    display_df["Max Life"] = display_df["Max Life"].map(f"{{:.2f}} {period_unit}".format)
    display_df["Avg Life"] = display_df["Avg Life"].map(f"{{:.2f}} {period_unit}".format)
    display_df["Weight (%)"] = display_df["Weight (%)"].map("{:.2f}%".format)
    st.dataframe(display_df, use_container_width=True, height=300)

# Volatility calculation functions (from BSM6.py)
@st.cache_data
def fetch_stock_data(ticker, start, end):
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

def process_stock_data(data, ticker, start_date, end_date):
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
    return fig

def generate_excel(company_name, ticker, vesting_df, stock_data_dict, volatility_summary, period_ref_data):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # Formats
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center', 'font_size': 12})
    cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 11})
    pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'border': 1, 'align': 'center', 'font_size': 11})
    date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center', 'font_size': 11})
    period_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center', 'font_size': 11})

    # Summary sheet
    sheet = workbook.add_worksheet("Summary")
    sheet.write(0, 0, "ESOP Valuation Report", workbook.add_format({'bold': True, 'font_size': 14}))
    sheet.write(2, 0, "Company Name", header_fmt)
    sheet.write(2, 1, company_name, cell_fmt)
    sheet.write(3, 0, "Ticker", header_fmt)
    sheet.write(3, 1, ticker, cell_fmt)
    sheet.write(4, 0, "Grant Date", header_fmt)
    sheet.write(4, 1, grant_date.strftime('%d/%m/%Y'), date_fmt)
    sheet.write(5, 0, "Period Unit", header_fmt)
    sheet.write(5, 1, period_unit, cell_fmt)
    
    # Vesting schedule
    sheet.write(7, 0, "Vesting Schedule", header_fmt)
    headers = ["Vesting", "Vesting Date", f"Vesting Period ({period_unit})", "Weight (%)", 
               f"Exercise Validity ({period_unit})", "Exercise End Date", 
               f"Min Life ({period_unit})", f"Max Life ({period_unit})", f"Avg Life ({period_unit})", 
               "Volatility Start Date", "Volatility End Date", "Annualized Volatility (%)"]
    for col, header in enumerate(headers):
        sheet.write(9, col, header, header_fmt)
    
    for row, item in enumerate(vesting_df.to_dict('records'), start=10):
        sheet.write(row, 0, item["Vesting"], cell_fmt)
        sheet.write(row, 1, item["Vesting Date"], date_fmt)
        sheet.write_number(row, 2, item["Vesting Period"], period_fmt)
        sheet.write_number(row, 3, item["Weight (%)"], pct_fmt)
        sheet.write_number(row, 4, item["Exercise Validity"], period_fmt)
        sheet.write(row, 5, item["Exercise End Date"], date_fmt)
        sheet.write_number(row, 6, item["Min Life"], period_fmt)
        sheet.write_number(row, 7, item["Max Life"], period_fmt)
        sheet.write_number(row, 8, item["Avg Life"], period_fmt)
        sheet.write(row, 9, item["Volatility Start Date"], date_fmt)
        sheet.write(row, 10, item["Volatility End Date"], date_fmt)
        sheet.write_number(row, 11, item["Annualized Volatility (%)"], pct_fmt)
    
    sheet.set_column(0, 0, 15)
    sheet.set_column(1, 1, 15)
    sheet.set_column(2, 8, 12)
    sheet.set_column(9, 10, 15)
    sheet.set_column(11, 11, 20)

    # Volatility summary sheet
    vol_sheet = workbook.add_worksheet("Volatility Summary")
    vol_headers = ["Vesting", "Ticker", "Std Dev (%)", "No. of Days", "Period", "Annualized Volatility (%)"]
    for col, header in enumerate(vol_headers):
        vol_sheet.write(0, col, header, header_fmt)
    
    for row, (vesting, summary) in enumerate(volatility_summary.items(), start=1):
        vol_sheet.write(row, 0, vesting, cell_fmt)
        vol_sheet.write(row, 1, summary["Ticker"], cell_fmt)
        vol_sheet.write_number(row, 2, summary["Std Dev (%)"], pct_fmt)
        vol_sheet.write_number(row, 3, summary["No. of Days"], cell_fmt)
        vol_sheet.write(row, 4, summary["Period"], cell_fmt)
        vol_sheet.write_number(row, 5, summary["Annualized Volatility (%)"], pct_fmt)
    
    vol_sheet.set_column(0, 0, 15)
    vol_sheet.set_column(1, 1, 15)
    vol_sheet.set_column(2, 2, 12)
    vol_sheet.set_column(3, 3, 12)
    vol_sheet.set_column(4, 4, 25)
    vol_sheet.set_column(5, 5, 20)

    # Stock data sheets
    for vesting, data in stock_data_dict.items():
        ws = workbook.add_worksheet(f"Data_{vesting[:28]}")  # Excel sheet name limit
        headers = ["Date", "Price", "Daily % Change"]
        for col, header in enumerate(headers):
            ws.write(0, col, header, header_fmt)
        
        for row, row_data in enumerate(data.itertuples(name=None), start=1):
            ws.write(row, 0, row_data[1].strftime('%d/%m/%Y'), date_fmt)
            ws.write_number(row, 1, row_data[2], period_fmt)
            ws.write_number(row, 2, row_data[3], pct_fmt)
        
        ws.set_column(0, 0, 15)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 15)

    # Period conversion reference sheet (at the end)
    period_sheet = workbook.add_worksheet("Period Conversion")
    period_sheet.write(0, 0, "Period Conversion Reference", header_fmt)
    ref_headers = ["Period", "Days"]
    for col, header in enumerate(ref_headers):
        period_sheet.write(1, col, header, header_fmt)
    for row, ref in enumerate(period_ref_data, start=2):
        period_sheet.write(row, 0, ref["Period"], cell_fmt)
        period_sheet.write_number(row, 1, ref["Days"], cell_fmt)
    period_sheet.set_column(0, 0, 15)
    period_sheet.set_column(1, 1, 15)

    workbook.close()
    output.seek(0)
    return output

# Main logic
if st.sidebar.button("Calculate ESOP Valuation"):
    with st.spinner("Fetching and processing data..."):
        stock_data_dict = {}
        volatility_summary = {}
        
        for index, row in vesting_df.iterrows():
            vesting = row["Vesting"]
            start_date = pd.to_datetime(row["Volatility Start Date"], format='%d/%m/%Y')
            end_date = pd.to_datetime(row["Volatility End Date"], format='%d/%m/%Y')
            data = fetch_stock_data(ticker, start_date, end_date)
            if data is None:
                st.warning(f"No valid price data for {vesting} ({ticker}). Skipping.")
                volatility_summary[vesting] = {
                    "Ticker": ticker,
                    "Std Dev (%)": 0,
                    "No. of Days": 0,
                    "Period": f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}",
                    "Annualized Volatility (%)": 0
                }
                vesting_df.loc[vesting_df["Vesting"] == vesting, "Annualized Volatility (%)"] = 0
                continue
            
            processed_data, summary = process_stock_data(data, ticker, start_date, end_date)
            stock_data_dict[vesting] = processed_data
            volatility_summary[vesting] = summary
            vesting_df.loc[vesting_df["Vesting"] == vesting, "Annualized Volatility (%)"] = summary["Annualized Volatility (%)"]
            
            # Display data (as in BSM6.py)
            st.subheader(f"Data for {vesting} ({ticker})", anchor=None)
            display_df = processed_data.copy()
            display_df['Date'] = display_df['Date'].dt.strftime('%d/%m/%Y')
            display_df['Daily % Change'] = display_df['Daily % Change'].map('{:.2f}%'.format)
            display_df['Price'] = display_df['Price'].map('{:.2f}'.format)
            st.dataframe(display_df[['Date', 'Price', 'Daily % Change']], use_container_width=True, height=300)
            
            # Plot
            st.write(f"**Daily % Change Chart for {vesting}**")
            fig = plot_daily_change(processed_data, ticker)
            st.pyplot(fig)
        
        if volatility_summary:
            st.markdown("### 📊 Volatility Summary")
            summary_df = pd.DataFrame(list(volatility_summary.values()))
            summary_df["Vesting"] = volatility_summary.keys()
            summary_df["Std Dev (%)"] = summary_df["Std Dev (%)"].map("{:.2f}%".format)
            summary_df["Annualized Volatility (%)"] = summary_df["Annualized Volatility (%)"].map("{:.2f}%".format)
            summary_df = summary_df[["Vesting", "Ticker", "Std Dev (%)", "No. of Days", "Period", "Annualized Volatility (%)"]]
            st.table(summary_df)
            
            # Display updated vesting table with volatility
            st.markdown(f"**Updated Vesting Schedule with Volatility (Periods in {period_unit})**")
            display_df = vesting_df[["Vesting", "Vesting Date", "Vesting Period", "Weight (%)", 
                                    "Exercise Validity", "Exercise End Date", 
                                    "Min Life", "Max Life", "Avg Life", 
                                    "Volatility Start Date", "Volatility End Date", "Annualized Volatility (%)"]]
            display_df = display_df.copy()
            display_df["Vesting Period"] = display_df["Vesting Period"].map(f"{{:.2f}} {period_unit}".format)
            display_df["Exercise Validity"] = display_df["Exercise Validity"].map(f"{{:.2f}} {period_unit}".format)
            display_df["Min Life"] = display_df["Min Life"].map(f"{{:.2f}} {period_unit}".format)
            display_df["Max Life"] = display_df["Max Life"].map(f"{{:.2f}} {period_unit}".format)
            display_df["Avg Life"] = display_df["Avg Life"].map(f"{{:.2f}} {period_unit}".format)
            display_df["Weight (%)"] = display_df["Weight (%)"].map("{:.2f}%".format)
            display_df["Annualized Volatility (%)"] = display_df["Annualized Volatility (%)"].map("{:.2f}%".format)
            st.dataframe(display_df, use_container_width=True, height=300)
            
            # Excel download
            excel_file = generate_excel(company_name, ticker, vesting_df, stock_data_dict, volatility_summary, period_ref_data)
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_file,
                file_name=f"{company_name}_ESOP_Valuation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=str(uuid.uuid4())
            )
        
        st.markdown("---")
        st.markdown("**Note**: Adjusted closing prices are used when available to account for splits, dividends, and corporate actions. If unavailable, closing prices are used.")
        
        # Period conversion reference table (at the end)
        st.markdown("### Period Conversion Reference (Years to Days)")
        period_ref_df = pd.DataFrame(period_ref_data)
        st.table(period_ref_df)
else:
    st.info("Enter valid inputs and click 'Calculate ESOP Valuation' to see results.")
