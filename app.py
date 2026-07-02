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

# --- 💅 PREMIUM GLOBAL STYLING ---
st.markdown("""
<style>
    .global-header { text-align: center; margin: 0px 0px 20px 0px; font-weight: 800; font-size: 44px; letter-spacing: 2px; background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff); background-size: 300% 300%; -webkit-background-clip: text; -webkit-text-fill-color: transparent; animation: gradientShift 6s ease infinite; }
    @keyframes gradientShift { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .brand-title { text-align: center; font-weight: 800; font-size: 42px; background: linear-gradient(45deg, #FF5733, #FFC300); -webkit-background-clip: text; -webkit-text-fill-color: transparent; padding-top: 40px; }
    .sidebar-brand-box { background: linear-gradient(135deg, #1F4E79, #2c3e50); padding: 18px; border-radius: 8px; margin-bottom: 25px; border-left: 5px solid #FF5733; text-align: center; }
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #eef2f5; }
</style>
""", unsafe_allow_html=True)

# --- INITIALIZATION ---
if "supervisor_targets" not in st.session_state: st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "party_options" not in st.session_state: st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics"]
if "item_options" not in st.session_state: st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA"]
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "email_config" not in st.session_state: st.session_state["email_config"] = {"sender": "your_email@gmail.com", "password": "your_app_password", "receiver": "receiver_email@gmail.com"}

# --- HELPER FUNCTION ---
def send_daywise_backup_email():
    cfg = st.session_state["email_config"]
    current_today = datetime.now().strftime("%d-%m-%Y")
    if not os.path.exists(EXCEL_FILE): return False
    try:
        main_df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        main_df['Date'] = main_df['Date'].astype(str).str.strip()
        day_wise_df = main_df[main_df['Date'] == current_today]
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer: day_wise_df.to_excel(writer, index=False)
        excel_buffer.seek(0)
        # Email logic... (simplified for space)
        return True
    except: return False

# --- LOGIN ---
if not st.session_state["logged_in"]:
    st.markdown('<h1 class="brand-title">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("LOGIN"):
            if username in st.session_state["users"] and st.session_state["users"][username]["password"] == password:
                st.session_state["logged_in"] = True
                st.session_state["current_user"] = st.session_state["users"][username]
                st.rerun()
else:
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    
    # --- SIDEBAR ---
    st.sidebar.markdown(f"### Welcome {user['name']}")
    nav_options = ["📊 Dashboard"]
    if user["role"] == "admin":
        nav_options.extend(["📝 Data Entry", "🎯 Target Settings", "⚙️ Factory Configurations Desk", "📧 Email Setup"])
    nav_choice = st.sidebar.radio("Navigation", nav_options)

    # --- CONTENT LOGIC ---
    if nav_choice == "📊 Dashboard":
        st.subheader("📊 Live Analytics")
        filter_type = st.radio("Filter Type:", ["☀️ Day-Wise", "📆 Month-Wise"], horizontal=True)
        # ... (Dashboard logic wahi rahegi) ...
        st.write("Dashboard Data Here...")

    elif nav_choice == "⚙️ Factory Configurations Desk":
        st.subheader("⚙️ Factory Configurations Desk")
        t1, t2, t3 = st.tabs(["🏢 Manage Parties", "📦 Manage Items", "👥 Supervisors"])
        with t1:
            st.write("Party Management...")
        with t2:
            st.write("Item Management...")
        with t3:
            st.write("Supervisor Management...")

    # Logout
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
