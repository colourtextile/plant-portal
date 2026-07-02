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

# --- INIT SESSION STATE FIRST ---
st.set_page_config(page_title="Colour Textile Portal", layout="wide")

if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "current_user" not in st.session_state: st.session_state["current_user"] = None
if "supervisor_targets" not in st.session_state: st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False},
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "party_options" not in st.session_state: st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]
if "item_options" not in st.session_state: st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]
if "email_config" not in st.session_state: st.session_state["email_config"] = {"sender": "your_email@gmail.com", "password": "your_app_password", "receiver": "receiver_email@gmail.com"}

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

# --- STYLING ---
st.markdown("""
<style>
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #eef2f5; }
    .global-header { text-align: center; font-weight: 800; font-size: 40px; color: #1F4E79; }
</style>
""", unsafe_allow_html=True)

# --- LOGIN LOGIC ---
if not st.session_state["logged_in"]:
    st.markdown('<h1 class="global-header">COLOUR TEXTILE LOGIN</h1>', unsafe_allow_html=True)
    with st.form("login"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("LOGIN"):
            if username in st.session_state["users"] and st.session_state["users"][username]["password"] == password:
                st.session_state["logged_in"] = True
                st.session_state["current_user"] = st.session_state["users"][username]
                st.rerun()
            else:
                st.error("Invalid Login!")
else:
    user = st.session_state["current_user"]
    st.sidebar.markdown(f"Welcome, {user['name']}")
    nav_choice = st.sidebar.radio("Menu", ["📊 Dashboard", "📝 Data Entry"] if user["role"] == "admin" else ["📊 Dashboard"])
    
    if st.sidebar.button("Logout"):
        st.session_state["logged_in"] = False
        st.rerun()

    # --- LOAD DATA ---
    excel_loaded = False
    df = pd.DataFrame()
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        df.columns = ["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor", "Challan No"][:len(df.columns)]
        excel_loaded = True

    # --- DASHBOARD WITH 0-HIDE & PARTY SUMMARY ---
    if nav_choice == "📊 Dashboard":
        st.markdown("<h2 class='global-header'>📊 Analytics Dashboard</h2>", unsafe_allow_html=True)
        filter_type = st.radio("Range:", ["Day-Wise", "Month-Wise"], horizontal=True)
        
        if excel_loaded and not df.empty:
            df['Date'] = df['Date'].astype(str).str.strip()
            df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
            df['Party Name'] = df['Party Name'].astype(str).str.strip()
            
            if filter_type == "Day-Wise":
                sel = st.date_input("Date").strftime("%d-%m-%Y")
                f_df = df[df['Date'] == sel]
            else:
                f_df = df # Simplification for example
            
            if not f_df.empty:
                items = f_df.groupby('Item Type')['Total Pcs'].sum()
                items = items[items > 0] # 0 Hide
                
                parties = f_df.groupby('Party Name')['Total Pcs'].sum()
                parties = parties[parties > 0] # 0 Hide
                
                c1, c2 = st.columns(2)
                with c1:
                    st.write("### 🏭 Item Summary")
                    st.table(items.reset_index())
                with c2:
                    st.write("### 📦 Party Summary")
                    st.table(parties.reset_index())
            else:
                st.info("No data.")
