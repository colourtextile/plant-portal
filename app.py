import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io

# --- CONFIGURATION ---
EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"
st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# (Style code waisa hi rakhein jaisa aapne bheja tha)
st.markdown("""
<style>
    .global-header { text-align: center; font-weight: 800; font-size: 44px; background: linear-gradient(45deg, #FF5733, #FFC300); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .sidebar-brand-box { background: linear-gradient(135deg, #1F4E79, #2c3e50); padding: 18px; border-radius: 8px; color: white; text-align: center; }
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATES ---
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True}
    }
if "party_options" not in st.session_state:
    st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]
if "item_options" not in st.session_state:
    st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

# --- DASHBOARD LOGIC ---
if st.session_state["logged_in"]:
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    
    # 📅 Filter Range
    filter_type = st.radio("📅 Select Dashboard Filter Range:", ["☀️ Day-Wise Filter", "📆 Month-Wise Filter"], horizontal=True)
    
    # Logic to load and filter data
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        df['Date'] = df['Date'].astype(str).str.strip()
        df['parsed_date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
        df['Month_Year'] = df['parsed_date'].dt.strftime('%B %Y')

        if filter_type == "☀️ Day-Wise Filter":
            sel_date = st.date_input("Choose Date", datetime.now()).strftime("%d-%m-%Y")
            filtered_df = df[df["Date"] == sel_date]
            label = f"SELECTED DATE: {sel_date}"
        else:
            months = sorted(df['Month_Year'].dropna().unique())
            sel_month = st.selectbox("Choose Month", months)
            filtered_df = df[df["Month_Year"] == sel_month]
            label = f"SELECTED MONTH: {sel_month.upper()}"

        if not filtered_df.empty:
            # --- 1. ITEM SUMMARY (0 Hide Logic) ---
            item_summary = filtered_df.groupby('Item Type')['Total Pcs'].sum().reset_index()
            item_summary = item_summary[item_summary['Total Pcs'] > 0]
            
            # --- 2. PARTY SUMMARY (New) ---
            party_summary = filtered_df.groupby('Party Name')['Total Pcs'].sum().reset_index()
            party_summary = party_summary[party_summary['Total Pcs'] > 0]

            col1, col2 = st.columns(2)

            with col1:
                st.markdown(f"### 🏭 ITEM PIECES SUMMARY ({label})")
                st.table(item_summary.rename(columns={'Item Type': 'Item', 'Total Pcs': 'Pcs'}))
                st.success(f"TOTAL PRODUCED: {item_summary['Total Pcs'].sum()} Pcs")

            with col2:
                st.markdown(f"### 📦 PARTY WISE SUMMARY ({label})")
                st.table(party_summary.rename(columns={'Party Name': 'Party', 'Total Pcs': 'Pcs'}))
                st.info(f"TOTAL DISPATCH: {party_summary['Total Pcs'].sum()} Pcs")
        else:
            st.warning("No data found for the selected filter.")
