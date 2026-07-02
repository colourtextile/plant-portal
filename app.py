import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

st.set_page_config(page_title="Plant Production Portal", layout="wide")

# 1. INITIALIZING DATA STRUCTURES
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin"},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor"},
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor"}
    }

if "party_options" not in st.session_state:
    st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]

if "item_options" not in st.session_state:
    st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = None

# 2. LOGIN SCREEN
if not st.session_state["logged_in"]:
    st.markdown("<h2 style='text-align: center; color: #1F4E79;'>🔐 Plant Portal - Secure Login</h2>", unsafe_allow_html=True)
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

# 3. MAIN APPLICATION (AFTER LOGIN)
else:
    user = st.session_state["current_user"]
    
    st.sidebar.markdown(f"### 👤 Welcome, **{user['name']}**")
    st.sidebar.markdown(f"📋 Role: `{user['role'].upper()}`")
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()

    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili! Kripya check karein.")
    else:
        # --- ROLE 1: SUPERVISOR ---
        if user["role"] == "supervisor":
            st.title("📝 Supervisor Data Entry Portal")
            with st.form("entry_form_sup", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    date_input = st.date_input("Date", datetime.now())
                    design_no = st.text_input("Design Number", placeholder="e.g., DS-1001")
                    party_name = st.selectbox("Select Party Name", st.session_state["party_options"])
                    item_type = st.selectbox("Select Item Type", st.session_state["item_options"])
                with col2:
                    total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1)
                    st.text_input("Supervisor Name", value=user["name"], disabled=True)
                
                if st.form_submit_button("SAVE ENTRY"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Galti: Total Pcs ka jod sahi nahi hai!")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"]])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved!")

        # --- ROLE 2: ADMIN ALL-IN-ONE DESKBOARD ---
        elif user["role"] == "admin":
            st.title("👑 Admin Master Deskboard")
            
            # Form Hidden inside Expander
            with st.expander("📝 Quick Data Entry Form", expanded=False):
                with st.form("entry_form_admin", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    with col1:
                        date_input = st.date_input("Date", datetime.now(), key="adm_date")
                        design_no = st.text_input("Design Number", placeholder="e.g., DS-1001", key="adm_des")
                        party_name = st.selectbox("Select Party Name", st.session_state["party_options"], key="adm_party")
                        item_type = st.selectbox("Select Item Type", st.session_state["item_options"], key="adm_item")
                    with col2:
                        total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1, key="adm_tot")
                        fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1, key="adm_fr")
                        seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1, key="adm_sec")
                    if st.form_submit_button("SAVE ADMIN ENTRY"):
                        if fresh_pcs + seconds_pcs != total_pcs:
                            st.error("❌ Total Pcs galat hain!")
                        else:
                            wb = openpyxl.load_workbook(EXCEL_FILE)
                            ws = wb["Supervisor Entry"]
                            ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"]])
                            wb.save(EXCEL_FILE)
                            st.success("🎉 Entry Saved!")
                            st.rerun()

            st.markdown("---")
            
            # Load Excel Data
            df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
            if not df.empty and df.shape[1] > 6:
                df.columns = ["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor"]
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                item_groups = df.groupby('Item Type')['Total Pcs'].sum().to_dict()
            else:
                item_groups = {}

            # Process items list safely
            items_list = [it.upper().strip() for it in st.session_state["item_options"]]
            
            # Create Perfect Native Layout Table (No HTML Breaks!)
            summary_rows = []
            current_date = datetime.now().strftime("%d-%m-%y")
            
            summary_rows.append({"PRODUCTION": "DATE", "VALUE / PCS": current_date})
            total_sum = 0
            for it in items_list:
                val = item_groups.get(it, 0)
                summary_rows.append({"PRODUCTION": it, "VALUE / PCS": str(val)})
                total_sum += val
            summary_rows.append({"PRODUCTION": "TOTAL PCS", "VALUE / PCS": str(total_sum)})
            
            summary_table_df = pd.DataFrame(summary_rows)

            # --- SCREEN COLUMNS FOR GRAPH AND TABLE ---
            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.subheader("📊 Live Summary Table")
                # Custom CSS styling for beautiful Clean grid lines & colors
                st.dataframe(
                    summary_table_df, 
                    hide_index=True, 
                    use_container_width=True
                )
                
            with col_right:
                st.subheader("🍩 Item Wise Percentage Graph")
                # Create Round Percentage Chart data
                chart_labels = []
                chart_values = []
                for it in items_list:
                    val = item_groups.get(it, 0)
                    if val > 0:  # Only show items that have data
                        chart_labels.append(it)
                        chart_values.append(val)
                
                if chart_values:
                    chart_df = pd.DataFrame({"Items": chart_labels, "Pieces": chart_values})
                    st.vega_lite_chart(chart_df, {
                        'mark': {'type': 'arc', 'innerRadius': 50, 'tooltip': True},
                        'encoding': {
                            'theta': {'field': 'Pieces', 'type': 'quantitative'},
                            'color': {'field': 'Items', 'type': 'nominal', 'scale': {'scheme': 'category20'}},
                        },
                        'view': {'stroke': None}
                    }, use_container_width=True)
                else:
                    st.info("Graph dekhne ke liye pehle kam-se-kam ek data entry kijiye!")

            # --- PART D: RAW DATA & CONTROLS ---
            st.markdown("---")
            st.subheader("📋 Raw Excel Logs")
            st.dataframe(df, hide_index=True, use_container_width=True, height=250)

            st.markdown("---")
            st.subheader("⚙️ System Configuration")
            t1, t2 = st.tabs(["👥 Supervisors Accounts", "🏢 Dropdowns Management"])
            with t1:
                col_u1, col_u2 = st.columns(2)
                with col_u1:
                    with st.form("add_user_form", clear_on_submit=True):
                        st.markdown("**Add New Supervisor**")
                        new_id = st.text_input("Username / ID").strip()
                        new_pass = st.text_input("Password").strip()
                        new_name = st.text_input("Name").strip()
                        if st.form_submit_button("Create"):
                            if new_id and new_pass and new_name:
                                st.session_state["users"][new_id] = {"password": new_pass, "name": new_name, "role": "supervisor"}
                                st.success("User Created!")
                with col_u2:
                    st.write("Active Accounts:", list(st.session_state["users"].keys()))
            with t2:
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.write("Current Parties:", st.session_state["party_options"])
                    new_party = st.text_input("Add Party").strip()
                    if st.button("➕ Party Add"):
                        if new_party and new_party not in st.session_state["party_options"]:
                            st.session_state["party_options"].append(new_party)
                            st.rerun()
                with col_d2:
                    st.write("Current Items:", st.session_state["item_options"])
                    new_item = st.text_input("Add Item").strip()
                    if st.button("➕ Item Add"):
                        if new_item and new_item not in st.session_state["item_options"]:
                            st.session_state["item_options"].append(new_item.upper())
                            st.rerun()
