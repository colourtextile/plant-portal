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

# --- 💅 STYLING ---
st.markdown("""
<style>
    .global-header { text-align: center; font-weight: 800; font-size: 40px; color: #1F4E79; }
    .sidebar-brand-box { background: #1F4E79; padding: 20px; border-radius: 10px; color: white; text-align: center; }
</style>
""", unsafe_allow_html=True)

# --- INITIALIZATION ---
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "supervisor_targets" not in st.session_state: st.session_state["supervisor_targets"] = {"Ramesh": 500}
if "party_options" not in st.session_state: st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics"]
if "item_options" not in st.session_state: st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA"]
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False

# --- LOGIN ---
if not st.session_state["logged_in"]:
    st.markdown('<h1 style="text-align:center;">COLOUR TEXTILE SYSTEM</h1>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("LOGIN"):
            if username in st.session_state["users"] and st.session_state["users"][username]["password"] == password:
                st.session_state["logged_in"] = True
                st.session_state["current_user"] = st.session_state["users"][username]
                st.rerun()
            else: st.error("Invalid Login!")
else:
    user = st.session_state["current_user"]
    st.sidebar.markdown(f'<div class="sidebar-brand-box"><h3>{user["name"]}</h3><p>Role: {user["role"].upper()}</p></div>', unsafe_allow_html=True)
    nav_choice = st.sidebar.radio("Navigation", ["📊 Dashboard", "📝 Data Entry", "⚙️ Configurations"])
    
    df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()

    # --- DASHBOARD LOGIC (Filter Updated) ---
    if nav_choice == "📊 Dashboard":
        st.subheader("📊 Live Analytics Dashboard")
        filter_type = st.radio("Select Filter:", ["Day-Wise", "Month-Wise"], horizontal=True)
        
        if not df.empty:
            df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
            if filter_type == "Day-Wise":
                sel_date = st.date_input("Select Date")
                view_df = df[df['Date'].dt.date == sel_date]
            else:
                sel_month = st.selectbox("Select Month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
                view_df = df[df['Date'].dt.month_name() == sel_month]
            st.dataframe(view_df)

    # --- CONFIGURATIONS LOGIC (Supervisor Edit Updated) ---
    elif nav_choice == "⚙️ Configurations" and user["role"] == "admin":
        st.subheader("⚙️ System Configurations")
        t1, t2 = st.tabs(["Parties", "Supervisors"])
        with t2:
            col_u1, col_u2 = st.columns(2)
            with col_u2:
                st.markdown("**✏️ Edit Info & Permissions**")
                all_sups = [u for u in st.session_state["users"] if st.session_state["users"][u]["role"] == "supervisor"]
                selected_sup = st.selectbox("Select Supervisor", all_sups)
                curr = st.session_state["users"][selected_sup]
                
                edit_name = st.text_input("Name", value=curr["name"])
                edit_pass = st.text_input("Password", value=curr["password"])
                edit_entry = st.checkbox("Allow Entry", value=curr.get("p_entry", True))
                edit_view = st.checkbox("Allow View", value=curr.get("p_view", True))
                edit_edit = st.checkbox("Allow Edit", value=curr.get("p_edit", False))
                
                if st.button("Update Account"):
                    st.session_state["users"][selected_sup].update({"name": edit_name, "password": edit_pass, "p_entry": edit_entry, "p_view": edit_view, "p_edit": edit_edit})
                    st.success("Updated!")
                    st.rerun()

    if st.sidebar.button("Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
