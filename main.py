# Importing Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import matplotlib
import streamlit as st
from PIL import Image
import os #for opening png -> # os.system("start DELTA_Graph.png") 
from rsi import cal_rsi # 
import matplotlib.dates as mdates

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
    "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
    "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
    "SET Index", "SET เปลี่ยนแปลง(%)"
]

# Converting Months
thai_months = {
    "ม.ค.": "01", "ก.พ.": "02", "มี.ค.": "03", "เม.ย.": "04",
    "พ.ค.": "05", "มิ.ย.": "06", "ก.ค.": "07", "ส.ค.": "08",
    "ก.ย.": "09", "ต.ค.": "10", "พ.ย.": "11", "ธ.ค.": "12"
}

# Converting thai dates into global dates
def convert_thai_date(thai_date_str):
    for th, num in thai_months.items():
        if th in thai_date_str:
            day, month_th, year_th = thai_date_str.replace(",", "").split()
            month = thai_months[month_th]
            year = int(year_th) - 543
            return f"{year}-{month}-{int(day):02d}"
    return None


df = df[~df["วันที่"].isna()]
df = df[~df["วันที่"].astype(str).str.contains("วันที่")]


df["วันที่"] = df["วันที่"].apply(convert_thai_date)
df = df.dropna(subset=["วันที่"])
df["วันที่"] = pd.to_datetime(df["วันที่"], errors='coerce')
df = df.dropna(subset=["วันที่"])

# ----- End of dealing with months ----- #
# print(df.head(10))

# ค่า RSI ไม่มีเดือน 5 กับ 11 เพราะข้อมูลจาก settrade มีไม่ครบ 14 วันที่เป็น standard นํามาหาค่า rsi
df['RSI'] = cal_rsi(df)

# ----- Start working with Graph ----- #
# Building and Setting Graph
df_sorted = df.sort_values("วันที่")
X = df_sorted["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
y = df_sorted["ราคาปิด"].values

#Trend Line && Closing Value Graph
model = LinearRegression()
model.fit(X, y)
trend = model.predict(X)

plt.figure(figsize=(12, 6))
plt.plot(df_sorted["วันที่"], y, label="Actual Closing Price")
plt.plot(df_sorted["วันที่"], trend, label="Trend (Linear Regression)", linestyle="--", color="red")
plt.title("DELTA Closing Price Trend")
plt.xlabel("Date")
plt.ylabel("Closing Price (Baht)")
plt.legend()
plt.grid(True)
plt.tight_layout()

plt.savefig("DELTA_Graph.png")
plt.close() 
image = Image.open("DELTA_Graph.png")
# image.show()

#RSI Graph
fig, ax = plt.subplots(figsize=(12, 4))

ax.plot(df["วันที่"], df["RSI"], label='RSI (14)', color='purple')
ax.axhline(70, color='red', linestyle='--', label='Overbought (70)')
ax.axhline(30, color='green', linestyle='--', label='Oversold (30)')

ax.set_title("RSI (Relative Strength Index)")
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
# image1.show()

# plt.show()
# ----- End working with Graph ----- #

# ----- Start Building Webpage ----- #

st.set_page_config(layout="wide")



# เเบ่ง column เพื่อเเยกส่วน display => col1 ผั่งซ้าย col2 ตรงกลาง col3 ฝั่งขวา
col1, col2, col3 = st.columns([1, 4, 1])

with col2: # centered = col2

    with st.container(border=False):
        st.markdown(
        """
        <div style="background-color: #b6ae63; padding: 20px; border-top-left-radius: 20px; border-top-right-radius: 20px; margin-bottom: 10px">
            <h1>DELTA</h1>
            <p style="font-size: 28px;">บริษัทเดลต้า อีเลคโทรนิคส์ (ประเทศไทย) จำกัด (มหาชน)</p>
            <p style="font-size: 18px;">ข้อมูลย้อนหลัง 6 เดือนของวันที่ 19 พ.ค. 2568</p>
        </div>
        """,
        unsafe_allow_html=True
    )

        price_chart_slot = st.empty()
        rsi_chart_slot = st.empty()
        #Checkbox 
        show_price_chart = st.checkbox("แสดงกราฟราคาปิด", value=True)
        show_rsi_chart = st.checkbox("แสดงกราฟค่า RSI", value=True)


        if show_price_chart:
            price_chart_slot.image(image, caption="", use_container_width=True)
        else:
            price_chart_slot.empty()

        if show_rsi_chart:
            rsi_chart_slot.image(image1, caption="", use_container_width=True)
        else:
            rsi_chart_slot.empty()



# month & year column
df["เดือน"] = df["วันที่"].dt.month
df["ปี"] = df["วันที่"].dt.year

# month choices
months = df["เดือน"].unique()
months.sort()

# looping through month in option (ย้อนหลัง 6 เดือน)
month_options = ["ทั้งหมด"] + [f"{month:02d}" for month in months]

# select box for selecting month
selected_month = st.selectbox("เลือกเดือน", month_options)

# filtering data by choice selected
if selected_month != "ทั้งหมด":
    # if choice != ทั้งหมด then choice = selected int month (converted into number already)
    filtered_df = df[df["เดือน"] == int(selected_month)]
else:
    # copy all data (choice = ทั้งหมด) || สร้างสําเนา dataframe โดยไม่ให้เกิดการเปลี่ยนเเปลงตาม choice 
    filtered_df = df.copy()

# filtering to show only date (without showing time)
filtered_df["วันที่"] = filtered_df["วันที่"].dt.date

# filtering index starting from 1 to n
filtered_df.index = range(1, len(filtered_df) + 1)

# displaying filtered data
st.dataframe(filtered_df)