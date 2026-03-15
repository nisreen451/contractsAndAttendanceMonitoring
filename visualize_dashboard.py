import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Contracts & Attendance Dashboard", layout="wide")
st.title("Contracts & Attendance Dashboard")

# Load your processed Excel
file_path = "OutputFiles/contractsWithAttendanceMonitoring_1_2026.xlsx"
final_df = pd.read_excel(file_path, sheet_name="Contracts Data")

# Pie chart: Attendance Error Flags
st.subheader("Attendance Errors Distribution")
fig1 = px.pie(
    final_df, 
    names="Attendance Error Flag", 
    title="Contracts with Attendance Errors"
)
st.plotly_chart(fig1)

# Bar chart: Contracts with 7 Continuous No Attendance
st.subheader("Contracts with 7 Continuous No Attendance Days")
no_attendance_df = final_df[final_df["No Attendance 7 Continuous Days"]]
fig2 = px.bar(
    no_attendance_df, 
    x="Contract No.", 
    y="No Attendance 7 Continuous Days", 
    title="7 Continuous No Attendance Records"
)
st.plotly_chart(fig2)

# KPI Metrics Table
st.subheader("KPIs Summary")
kpi_df = pd.read_excel(file_path, sheet_name="KPIs_Comparison")
st.dataframe(kpi_df)
