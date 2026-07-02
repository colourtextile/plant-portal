import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

# --- CONFIGURATION ---
EXCEL_FILE = "erp_data.xlsx"
st.set_page_config(page_title="My ERP System", layout="wide")

# --- DATA INITIALIZATION ---
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["Date", "User", "Category", "Amount", "Description"])
    df.to_excel(EXCEL_FILE, index=False)

# --- AUTHENTICATION ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

def login():
    st.sidebar.title("Login")
    user = st.sidebar.text_input("Username")
    pwd = st.sidebar.text_input("Password", type="password")
    if st.sidebar.button("Login"):
        if user == "admin" and pwd == "1234": # Yahan apna secure login set karein
            st.session_state.logged_in = True
            st.session_state.user = user
            st.rerun()
        else:
            st.sidebar.error("Invalid Credentials")

# --- MAIN APP ---
def main_app():
    st.title("📊 ERP Management Dashboard")
    st.write(f"Welcome, {st.session_state.user}!")
    
    # Entry Form
    with st.form("entry_form"):
        col1, col2 = st.columns(2)
        date = col1.date_input("Date", datetime.now())
        category = col2.selectbox("Category", ["Sales", "Purchase", "Expense", "Other"])
        amount = st.number_input("Amount", min_value=0.0)
        desc = st.text_input("Description")
        
        if st.form_submit_button("Save Entry"):
            new_data = pd.DataFrame([[date, st.session_state.user, category, amount, desc]], 
                                    columns=["Date", "User", "Category", "Amount", "Description"])
            
            # Save to Excel
            try:
                with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                    df = pd.read_excel(EXCEL_FILE)
                    updated_df = pd.concat([df, new_data], ignore_index=True)
                    updated_df.to_excel(writer, index=False)
                st.toast("✅ Data Saved Successfully!")
            except Exception as e:
                st.error(f"Error saving data: {e}")

    # Display Data
    st.subheader("Recent Entries")
    df = pd.read_excel(EXCEL_FILE)
    st.dataframe(df, use_container_width=True)

    # Download Button
    st.download_button("📥 Download Report", data=open(EXCEL_FILE, "rb").read(), 
                       file_name="erp_report.xlsx", mime="application/vnd.ms-excel")

# --- RUN APP ---
if st.session_state.logged_in:
    main_app()
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()
else:
    login()
