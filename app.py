from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Dashboard Training",
    page_icon="📊",
    layout="wide",
)

DATA_FILE = Path("data.xlsx")

st.title("📊 Dashboard Training")
st.caption("CS HAD | Internal Training Material | 2026")

if not DATA_FILE.exists():
    st.error("Không tìm thấy file data.xlsx. Hãy đặt file cùng thư mục với app.py.")
    st.stop()

df = pd.read_excel(DATA_FILE)

required_columns = {"Date", "Customer", "Bookings", "Processing_Time_Min"}
missing_columns = required_columns.difference(df.columns)
if missing_columns:
    st.error("File Excel thiếu cột: " + ", ".join(sorted(missing_columns)))
    st.stop()

df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
df["Bookings"] = pd.to_numeric(df["Bookings"], errors="coerce").fillna(0)
df["Processing_Time_Min"] = pd.to_numeric(
    df["Processing_Time_Min"], errors="coerce"
).fillna(0)

st.sidebar.header("Bộ lọc")
customers = sorted(df["Customer"].dropna().astype(str).unique().tolist())
selected_customers = st.sidebar.multiselect(
    "Customer",
    options=customers,
    default=customers,
)

filtered = df[df["Customer"].astype(str).isin(selected_customers)].copy()

col1, col2, col3 = st.columns(3)
col1.metric("Total Bookings", f"{int(filtered['Bookings'].sum()):,}")
col2.metric("Customers", filtered["Customer"].nunique())
col3.metric(
    "Avg. Processing Time",
    f"{filtered['Processing_Time_Min'].mean():.1f} min",
)

st.subheader("Booking Volume by Customer")
booking_by_customer = (
    filtered.groupby("Customer", as_index=False)["Bookings"].sum()
    .sort_values("Bookings", ascending=False)
)
fig = px.bar(
    booking_by_customer,
    x="Customer",
    y="Bookings",
    text_auto=True,
)
fig.update_layout(xaxis_title="", yaxis_title="Bookings")
st.plotly_chart(fig, use_container_width=True)

st.subheader("Source Data")
st.dataframe(filtered, hide_index=True, use_container_width=True)
