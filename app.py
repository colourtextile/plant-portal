import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

st.set_page_config(page_title="Plant Production Portal", layout="wide")

# Default Master Admin and initial supervisors
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin"},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor"},
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor"}
    }

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = None

# LOGIN SCREEN
if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #1F4E79;'>🔐 Plant Portal - Secure Login</h2>", unsafe_content_type=True)
    with st.form("login_form"):
        username = st.text_input("Username / ID").strip()
        password = st.text_input("Password", type="password").strip()
        if st.form_submit_button("LOGIN"):
            if username in st.session_state["users"] and st.session_state["users"][username]["password"] == password:
                st.session_state["logged_in"] = True
                st.session_state["current_user"] = st.session_state["users"][username]
                st.rerun()
            else:
                st.error("❌ Galat ID ya Password!")

# MAIN APPLICATION
else:
    user = st.session_state["current_user"]
    st.sidebar.markdown(f"### 👤 Active: **{user['name']}**")
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()

    # Admin Panel is visible only to Admin role
    options = ["📝 Supervisor Data Entry", "📊 Executive Dashboard"]
    if user["role"] == "admin":
        options.append("👑 Master: Manage Users")
        
    menu = st.sidebar.radio("Navigation Menu", options)

    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili! Kripya check karein.")
    else:
        if menu == "📝 Supervisor Data Entry":
            st.subheader("Supervisor Data Entry Form")
            with st.form("entry_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    date_input = st.date_input("Date", datetime.now())
                    design_no = st.text_input("Design Number", placeholder="e.g., DS-1001")
                    party_name = st.selectbox("Select Party Name", ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"])
                    item_type = st.selectbox("Select Item Type", ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"])
                with col2:
                    total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1)
                    st.text_input("Supervisor Name", value=user["name"], disabled=True)
                
                if st.form_submit_button("SAVE"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Galti: Fresh Pcs + Seconds Pcs ka jod Total Pcs ke barabar hona chahiye!")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"]])
                        wb.save(EXCEL_FILE)
                        st.success(f"🎉 Data Save Ho Gaya! (Saved by: {user['name']})")

        elif menu == "📊 Executive Dashboard":
            st.subheader("Live Production Reports")
            df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
            if df.empty:
                st.warning("Excel me koi data nahi hai.")
            else:
                st.dataframe(df, hide_index=True, use_container_width=True)

        elif menu == "👑 Master: Manage Users":
            st.subheader("Add New Supervisor ID/Password")
            with st.form("add_user_form"):
                new_id = st.text_input("New Username / ID")
                new_pass = st.text_input("Set Password")
                new_name = st.text_input("Supervisor Real Name")
                if st.form_submit_button("Create User"):
                    if new_id and new_pass and new_name:
                        st.session_state["users"][new_id] = {"password": new_pass, "name": new_name, "role": "supervisor"}
                        st.success(f"🎉 New User '{new_name}' Successfully Created!")
                    else:
                        st.error("Saari details bharna zaroori hai.")