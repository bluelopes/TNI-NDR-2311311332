# Importing Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import matplotlib
import streamlit as st
from PIL import Image
import os #for opening png -> # os.system("start DELTA_Graph.png") 
import matplotlib.dates as mdates

# Importing Functions
from rsi import cal_rsi 
from ema import cal_ema

# ----- Start of dealing with months ----- #
# Setting up font
matplotlib.rcParams['font.family'] = 'DejaVu Sans' 

# Setting up pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 2000)
pd.set_option('display.max_colwidth', None)

# Reading Excel file
df = pd.read_excel("Delta_data.xlsx", sheet_name="Delta_dat", skiprows=1)

# Setting Column names
df.columns = [
    "Date", "Open", "Highest", "Lowest", "Average", "Close",
    "Change", "Change(%)", "Volume(000 shares)", "Value(Million Baht)",
    "SET Index", "SET Change(%)"
]

# Converting Months
thai_months = {
    "ม.ค.": "01", "ก.พ.": "02", "มี.ค.": "03", "เม.ย.": "04",
    "พ.ค.": "05", "มิ.ย.": "06", "ก.ค.": "07", "ส.ค.": "08",
    "ก.ย.": "09", "ต.ค.": "10", "พ.ย.": "11", "ธ.ค.": "12"
}

# Converting thai years into global years
def convert_thai_date(thai_date_str):
    for th, num in thai_months.items():
        if th in thai_date_str:
            day, month_th, year_th = thai_date_str.replace(",", "").split()
            month = thai_months[month_th]
            year = int(year_th) - 543
            return f"{year}-{month}-{int(day):02d}"
    return None


df = df[~df["Date"].isna()]
df = df[~df["Date"].astype(str).str.contains("Date")]


df["Date"] = df["Date"].apply(convert_thai_date)
df = df.dropna(subset=["Date"])
df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
df = df.dropna(subset=["Date"])

# ----- End of dealing with months ----- #


# ค่า RSI ไม่มีเดือน 5 กับ 11 เพราะข้อมูลจาก settrade มีไม่ครบ 14 วันที่เป็น standard นํามาหาค่า rsi
df['RSI'] = cal_rsi(df)
df['EMA'] = cal_ema(df)
# ----- Start working with Graph ----- #
# Building and Setting Graph
df_sorted = df.sort_values("Date")
X = df_sorted["Date"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
y = df_sorted["Close"].values

#Trend Line && Closing Value Graph
model = LinearRegression()
model.fit(X, y)
trend = model.predict(X)

plt.figure(figsize=(12, 6))
plt.plot(df_sorted["Date"], y, label = "Actual Closing Price")
plt.plot(df_sorted["Date"], trend, label = "Trend (Linear Regression)", linestyle = "--", color = "red")
plt.title("Closing Price Trend")
plt.xlabel("Date")
plt.ylabel("Closing Price (Baht)")
plt.legend()
plt.grid(True)
plt.tight_layout()

plt.savefig("DELTA_Graph.png")
image = Image.open("DELTA_Graph.png")


#RSI Graph
fig, ax = plt.subplots(figsize=(12, 6))

ax.plot(df["Date"], df["RSI"], label='RSI (14)', color='purple')
ax.axhline(70, color ='red', linestyle = '--', label='Overbought (70)')
ax.axhline(30, color = 'green', linestyle = '--', label='Oversold (30)')

ax.set_title("Relative Strength Index (RSI)")
ax.set_xlabel("Date")
ax.set_ylabel("RSI Value")
ax.legend()
ax.grid(True)

# Date axis
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)

plt.tight_layout()
plt.savefig("RSI.png")
image1 = Image.open("RSI.png")


# EMA Graph
plt.figure(figsize=(12,6))
plt.plot(df['Date'], df['Close'], label = 'Actual Closing Price', color='blue')
plt.plot(df['Date'], df['EMA'], label = 'Exponential Moving Average  (20 days)', color='orange')
plt.title('Closing Price with Exponential Moving Average (20 days)')
plt.xlabel('Date')
plt.ylabel('Price')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("EMA.png")
image2 = Image.open("EMA.png")

# ----- End working with Graph ----- #

# ----- Start Building Webpage ----- #

st.set_page_config(layout="wide")


# เเบ่ง column เพื่อเเยกส่วน display => col1 ผั่งซ้าย col2 ตรงกลาง col3 ฝั่งขวา
col1, col2, col3 = st.columns([1, 4, 1])

with col2: # centered = col2

    with st.container(border = False):
        st.markdown(
        """
        <div style="background-color: #b6ae63; padding: 20px; border-top-left-radius: 20px; border-top-right-radius: 20px; margin-bottom: 10px">
            <h1 style="text-align:center;">Stock DELTA (หุ้น DELTA)</h1>
            <h3>Delta Electronics India Manufacturing Private Limited</h1>
            <h3>บริษัทเดลต้า อีเลคโทรนิคส์ (ประเทศไทย) จำกัด (มหาชน)</h2>
            <h4>Data from the last 6 months as of May 19 2025</h2>
        </div>
        """,
        unsafe_allow_html = True
    )

    with st.container():
        show_price_chart = st.checkbox("Show Closing Price Graph", value=True, key="price_chart")
        if show_price_chart:
            st.image(image, caption="", use_container_width=True)

    with st.container():
        show_rsi_chart = st.checkbox("Show Relative Strength Index Graph", value=True, key="rsi_chart")
        if show_rsi_chart:
            st.image(image1, caption="", use_container_width=True)

    with st.container():
        show_ema_chart = st.checkbox("Show Exponential Moving Average Graph", value=True, key="ema_chart")
        if show_ema_chart:
            st.image(image2, caption="", use_container_width=True)

# month & year column
df["month"] = df["Date"].dt.month
df["year"] = df["Date"].dt.year

# month choices
months = df["month"].unique()
months.sort()

# looping through month in option (ย้อนหลัง 6 เดือน)
month_options = ["All"] + [f"{month:02d}" for month in months]

# select box for selecting month
selected_month = st.selectbox("Choose month", month_options)

# filtering data by choice selected
if selected_month != "All":
    # if choice != ทั้งหมด then choice = selected int month (converted into number already)
    filtered_df = df[df["month"] == int(selected_month)]
else:
    # copy all data (choice = ทั้งหมด) || สร้างสําเนา dataframe โดยไม่ให้เกิดการเปลี่ยนเเปลงตาม choice 
    filtered_df = df.copy()

# filtering to show only date (without showing time)
filtered_df["Date"] = filtered_df["Date"].dt.date

# filtering index starting from 1 to n
filtered_df.index = range(1, len(filtered_df) + 1)

# displaying filtered data
st.dataframe(filtered_df)