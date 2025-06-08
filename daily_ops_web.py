
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

def categorize_services(row):
    if row['V'] == "√":
        return "ON CALL - NEEDED ENGINEER SUPPORT", "2_TECH_SUPPORT"
    elif "CANCELED" in str(row['X']).upper():
        return "CANCELED", "3_CANCELED"
    elif row['R'] == "√":
        services = ["Transit"]
        if row['S'] == "√":
            services.append("Headset")
        if row['T'] == "√":
            services.append("Daily Check")
        if row['U'] == "√":
            services.append("Weekly Check")
        return ", ".join(services), "1_TRANSIT"
    else:
        return "Per Landing", "4_ON_CALL"

def format_excel(file):
    df = pd.read_excel(file, sheet_name=0, skiprows=4)

    df['Services'], df['SortKey'] = zip(*df.apply(categorize_services, axis=1))

    df['WO#'] = df['Q']
    df['Station'] = "KKIA"
    df['Customer'] = df['F'].astype(str).str[:2]
    df['Flight No.'] = df['F']
    df['Registration Code'] = df['E']
    df['Aircraft'] = df['D']
    df['Date'] = pd.to_datetime(df['B']).dt.strftime("%m/%d/%Y")

    def format_datetime(cell):
        try:
            return pd.to_datetime(cell).strftime("%m/%d/%Y %H:%M:%S")
        except:
            return ""

    df['STA.'] = df['G'].apply(format_datetime)
    df['ATA.'] = df['H'].apply(format_datetime)
    df['STD.'] = df['J'].apply(format_datetime)
    df['ATD.'] = df['K'].apply(format_datetime)

    df['Is Canceled'] = df['Services'].str.contains("CANCELED", case=False)
    df['Employees'] = df['O'].fillna('') + ', ' + df['P'].fillna('')
    df['Employees'] = df['Employees'].str.strip(', ')
    df['Remarks'] = ""
    df['Comments'] = ""

    df = df.sort_values(by='SortKey')

    final_df = df[['WO#', 'Station', 'Customer', 'Flight No.', 'Registration Code', 'Aircraft',
                   'Date', 'STA.', 'ATA.', 'STD.', 'ATD.', 'Is Canceled', 'Services', 'Employees',
                   'Remarks', 'Comments']]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Template')
    output.seek(0)
    return output

# Streamlit UI
st.title("🛫 Daily Ops Formatter (Web Tool)")
uploaded_file = st.file_uploader("📄 Upload the 'Daily Ops Report' Excel file", type=["xlsx"])

if uploaded_file:
    result = format_excel(uploaded_file)
    st.success("✅ File processed successfully!")
    st.download_button(
        label="📥 Download Formatted Excel File",
        data=result,
        file_name="Formatted_Daily_Ops.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
