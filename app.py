import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

# CONFIGURATION
EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"
st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# --- INITIALIZATION ---
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False

# --- LOGIN LOGIC ---
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
    st.sidebar.markdown(f"## 👤 {user['name']}")
    st.sidebar.write(f"**Role:** {user['role'].upper()}")
    nav_choice = st.sidebar.radio("Navigation", ["📊 Dashboard", "📝 Data Entry", "⚙️ Configurations"])
    
    # LOAD DATA
    df = pd.DataFrame()
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
        except Exception as e:
            st.error(f"Error loading file: {e}")

    # --- DASHBOARD LOGIC (FIXED) ---
    if nav_choice == "📊 Dashboard":
        st.subheader("📊 Live Analytics Dashboard")
        if not df.empty:
            # Fix KeyError: Identify date column dynamically
            date_col = next((col for col in df.columns if 'date' in col.lower()), None)
            
            if date_col:
                df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
                filter_type = st.radio("Select Filter:", ["Day-Wise", "Month-Wise"], horizontal=True)
                
                if filter_type == "Day-Wise":
                    sel_date = st.date_input("Select Date")
                    view_df = df[df[date_col].dt.date == sel_date]
                else:
                    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
                    sel_month = st.selectbox("Select Month", months)
                    view_df = df[df[date_col].dt.month_name() == sel_month]
                st.dataframe(view_df)
            else:
                st.error(f"Date column not found! Found columns: {list(df.columns)}")
        else:
            st.info("No data available.")

    # --- CONFIGURATIONS LOGIC ---
    elif nav_choice == "⚙️ Configurations" and user["role"] == "admin":
        st.subheader("⚙️ System Configurations")
        with st.expander("Manage Supervisor Permissions", expanded=True):
            col_u1, col_u2 = st.columns(2)
            with col_u2:
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
                    st.success("Updated Successfully!")
                    st.rerun()

    if st.sidebar.button("Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
