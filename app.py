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
    .global-header { text-align: center; font-weight: 800; font-size: 44px; background: linear-gradient(45deg, #FF5733, #30c381); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .brand-title { text-align: center; font-weight: 800; font-size: 42px; background: linear-gradient(45deg, #FF5733, #30c381); -webkit-background-clip: text; -webkit-text-fill-color: transparent; padding-top: 40px; }
    .sidebar-brand-box { background: linear-gradient(135deg, #1F4E79, #2c3e50); padding: 18px; border-radius: 8px; margin-bottom: 25px; border-left: 5px solid #FF5733; text-align: center; }
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #eef2f5; }
</style>
""", unsafe_allow_html=True)

# --- SESSION STATES INIT ---
if "supervisor_targets" not in st.session_state: st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "party_options" not in st.session_state: st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics"]
if "item_options" not in st.session_state: st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA"]
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "email_config" not in st.session_state: st.session_state["email_config"] = {"sender": "", "password": "", "receiver": ""}

# --- HELPER FUNCTIONS ---
def send_daywise_backup_email():
    # ... (Email logic wahi rahegi) ...
    pass

# --- MAIN APP ---
if not st.session_state["logged_in"]:
    st.markdown('<h1 class="brand-title">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    with st.form("login"):
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
    st.sidebar.markdown(f"""
    <div class="sidebar-brand-box">
        <div style="color: #FFC300; font-size: 20px; font-weight: 800;">WELCOME</div>
        <div style="color: #FFFFFF; font-size: 22px; font-weight: 900;">COLOUR TEXTILE</div>
    </div>
    """, unsafe_allow_html=True)
    
    nav_options = ["📊 Dashboard"]
    if user["role"] == "admin":
        nav_options.extend(["📝 Data Entry", "🎯 Target Settings", "⚙️ Factory Configs", "📧 Setup Email Auto-Backup"])
        
    nav_choice = st.sidebar.radio("🧭 Navigation Menu", nav_options)
    
    # --- LOGIC BASED ON NAVIGATION ---
    
    # 1. DASHBOARD
    if nav_choice == "📊 Dashboard":
        st.subheader("📊 Live Analytics Dashboard")
        filter_type = st.radio("Select Filter:", ["☀️ Day-Wise", "📆 Month-Wise"], horizontal=True)
        # ... (Dashboard logic yahan rahegi) ...
        st.info("Dashboard content here...")

    # 2. FACTORY CONFIGS (NEW SIDEBAR LOCATION)
    elif nav_choice == "⚙️ Factory Configs":
        st.subheader("⚙️ Factory Configurations Desk")
        t1, t2, t3 = st.tabs(["🏢 Manage Parties", "📦 Manage Items", "👥 Supervisors"])
        with t1:
            st.write("Party Management Logic Here")
        with t2:
            st.write("Item Management Logic Here")
        with t3:
            st.write("Supervisor Management Logic Here")

    # 3. OTHER PAGES
    elif nav_choice == "📝 Data Entry":
        st.subheader("📝 Quick Data Entry")
        # ...

    elif nav_choice == "📧 Setup Email Auto-Backup":
        st.subheader("📧 Email Settings")
        # ...

    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
