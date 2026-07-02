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

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"
st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# --- CSS STYLING (Same as your provided file) ---
st.markdown("""
<style>
    .reportview-container { background: #f8f9fa; }
    .global-header { text-align: center; margin: 0px 0px 20px 0px; font-weight: 800; font-size: 44px; letter-spacing: 2px; background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff); background-size: 300% 300%; -webkit-background-clip: text; -webkit-text-fill-color: transparent; animation: gradientShift 6s ease infinite; border-bottom: 2px solid #eaeaea; padding-bottom: 10px; }
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #eef2f5; }
</style>
""", unsafe_allow_html=True)

# --- (Yahan wahi Session States aur Helpers honge jo aapki file mein hain) ---
# ... [Session States aur Functions yahan wahi rahengi] ...

# --- MAIN LOGIC ---
if not st.session_state["logged_in"]:
    # ... [Login code] ...
    pass
else:
    # ... [Sidebar and other menus] ...
    
    # --- YEH HAI AAPKA UPDATED DASHBOARD LOGIC ---
    if nav_choice == "📊 Dashboard":
        st.markdown("<h2 style='color: #1F4E79; font-weight: bold;'>📊 Live Analytics Dashboard</h2>", unsafe_allow_html=True)
        filter_type = st.radio("Filter Range:", ["☀️ Day-Wise Filter", "📆 Month-Wise Filter"], horizontal=True)
        
        # Data Filtering
        if excel_loaded and not df.empty:
            df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
            df['Party Name'] = df['Party Name'].astype(str).str.strip()
            df['Date'] = df['Date'].astype(str).str.strip()
            df['parsed_date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
            df['Month_Year'] = df['parsed_date'].dt.strftime('%B %Y')
            
            if filter_type == "☀️ Day-Wise Filter":
                sel_date = st.date_input("Date", datetime.now()).strftime("%d-%m-%Y")
                filtered_df = df[df['Date'] == sel_date]
            else:
                months = sorted(df['Month_Year'].dropna().unique())
                sel_month = st.selectbox("Month", months)
                filtered_df = df[df['Month_Year'] == sel_month]
            
            # Aggragation with 0-Hide Logic
            if not filtered_df.empty:
                item_groups = filtered_df.groupby('Item Type')['Total Pcs'].sum()
                item_groups = item_groups[item_groups > 0]
                
                party_groups = filtered_df.groupby('Party Name')['Total Pcs'].sum()
                party_groups = party_groups[party_groups > 0]
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                    st.write("### 🏭 ITEM SUMMARY")
                    st.table(item_groups.reset_index())
                    st.markdown('</div>', unsafe_allow_html=True)
                with c2:
                    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                    st.write("### 📦 PARTY SUMMARY")
                    st.table(party_groups.reset_index())
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("No data found for this selection.")
