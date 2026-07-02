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

# --- PREMIUM STYLING ---
st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #0f172a !important; padding: 20px 10px; }
    .premium-sidebar-header {
        background: linear-gradient(135deg, #1e293b, #334155);
        padding: 20px; border-radius: 12px; margin-bottom: 30px;
        border-left: 5px solid #38bdf8; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        text-align: center;
    }
    .user-chip {
        display: inline-block; background: #0ea5e9; color: white;
        padding: 4px 12px; border-radius: 20px; font-size: 11px;
        font-weight: 700; margin-top: 10px; text-transform: uppercase;
    }
    .global-header {
        text-align: center; font-weight: 800; font-size: 44px;
        background: linear-gradient(45deg, #FF5733, #30c381);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
</style>
""", unsafe_allow_html=True)

# --- INITIALIZATION ---
if "supervisor_targets" not in st.session_state: st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "party_options" not in st.session_state: st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]
if "item_options" not in st.session_state: st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False

# --- LOGIC ---
if not st.session_state["logged_in"]:
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    with st.columns([1,1,1])[1]:
        with st.form("login"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("LOGIN"):
                if username in st.session_state["users"] and st.session_state["users"][username]["password"] == password:
                    st.session_state["logged_in"] = True
                    st.session_state["current_user"] = st.session_state["users"][username]
                    st.rerun()
                else: st.error("Invalid Login")
else:
    user = st.session_state["current_user"]
    st.sidebar.markdown(f"""
    <div class="premium-sidebar-header">
        <div style="color: #ffffff; font-size: 20px; font-weight: 900;">COLOUR TEXTILE</div>
        <div class="user-chip">{user['role'].upper()}</div>
        <div style="color: #cbd5e1; font-size: 13px; margin-top: 15px;">👤 {user['name']}</div>
    </div>
    """, unsafe_allow_html=True)
    
    nav_options = ["📊 Dashboard"]
    if user["role"] == "admin": nav_options.extend(["📝 Data Entry", "🎯 Target Settings"])
    choice = st.sidebar.radio("Navigation", nav_options)
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
    
    st.header(choice)
    st.write("Main Application Logic Here.")
