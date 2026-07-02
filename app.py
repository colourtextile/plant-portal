import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

st.set_page_config(page_title="Plant Production Portal", layout="wide")

# 1. INITIALIZING DATA STRUCTURES IN SESSION STATE
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
    
    # Sidebar for logout and info
    st.sidebar.markdown(f"### 👤 Welcome, **{user['name']}**")
    st.sidebar.markdown(f"📋 Role: `{user['role'].upper()}`")
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()

    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili! Kripya check karein.")
    else:
        # -------------------------------------------------------------
        # ROLE 1: IF SUPERVISOR LOGS IN (Only show entry form)
        # -------------------------------------------------------------
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
                        st.error("❌ Galti: Fresh + Seconds ka total Total Pieces ke barabar hona chahiye!")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"]])
                        wb.save(EXCEL_FILE)
                        st.success(f"🎉 Entry Saved Successfully by {user['name']}!")

        # -------------------------------------------------------------
        # ROLE 2: IF ADMIN LOGS IN (SHOW FULL DESKBOARD ON ONE PAGE)
        # -------------------------------------------------------------
        elif user["role"] == "admin":
            st.title("👑 Admin Master Deskboard")
            
            # --- PART A: ADMIN DATA ENTRY FORM ---
            with st.expander("📝 Quick Data Entry Form (Click to Open/Close)", expanded=False):
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
                        st.text_input("Logged In As", value=user["name"], disabled=True)
                    
                    if st.form_submit_button("SAVE ADMIN ENTRY"):
                        if fresh_pcs + seconds_pcs != total_pcs:
                            st.error("❌ Galti: Fresh + Seconds ka total Total Pieces ke barabar hona chahiye!")
                        else:
                            wb = openpyxl.load_workbook(EXCEL_FILE)
                            ws = wb["Supervisor Entry"]
                            ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"]])
                            wb.save(EXCEL_FILE)
                            st.success("🎉 Data saved directly to Excel sheet!")
                            st.rerun()

            # --- PART B: IMAGE 5E37A1 STYLE PRODUCTION REPORT ---
            st.markdown("---")
            
            df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
            
            col_dash1, col_dash2 = st.columns([1, 2])
            
            with col_dash1:
                st.subheader("📊 Live Production Summary")
                
                # Excel style summary table code
                summary_data = []
                total_all_pcs = 0
                
                if not df.empty and df.shape[1] > 6:
                    # Clean and standardize names
                    df.columns = ["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor"]
                    df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                    
                    # Group by Item Type and sum Total Pcs
                    item_groups = df.groupby('Item Type')['Total Pcs'].sum().to_dict()
                    
                    # Build list based on unique items
                    for item in st.session_state["item_options"]:
                        item_upper = item.upper().strip()
                        pcs = item_groups.get(item_upper, 0)
                        summary_data.append({"PRODUCTION": item_upper, "VALUE / PCS": f"{pcs:,}"})
                        total_all_pcs += pcs
                else:
                    for item in st.session_state["item_options"]:
                        summary_data.append({"PRODUCTION": item.upper(), "VALUE / PCS": "0"})
                
                # Add Total row at bottom
                summary_data.append({"PRODUCTION": "TOTAL PCS", "VALUE / PCS": f"{total_all_pcs:,}"})
                summary_df = pd.DataFrame(summary_data)
                
                # Render beautiful custom style table (Yellow header, light-blue date like image)
                current_date_str = datetime.now().strftime("%d-%m-%y")
                st.markdown(
                    f"""
                    <table style="width:100%; border-collapse: collapse; border: 2px solid black; text-align: left; font-family: sans-serif;">
                        <tr style="background-color: #FFFF00; color: black; font-weight: bold; text-align: center; border-bottom: 2px solid black;">
                            <td colspan="2" style="padding: 10px; font-size: 18px; letter-spacing: 1px;">PRODUCTION</td>
                        </tr>
                        <tr style="background-color: #8EA9DB; color: black; font-weight: bold; border-bottom: 1px solid black;">
                            <td style="padding: 8px;">DATE</td>
                            <td style="padding: 8px; text-align: right;">{current_date_str}</td>
                        </tr>
                    """, unsafe_allow_html=True
                )
                
                for index, row in summary_df.iterrows():
                    is_total = row['PRODUCTION'] == "TOTAL PCS"
                    bg_color = "#FFFFFF"
                    font_weight = "bold" if is_total else "normal"
                    border_top = "2px solid black" if is_total else "1px solid #D3D3D3"
                    font_size = "16px" if is_total else "14px"
                    
                    st.markdown(
                        f"""
                        <tr style="background-color: {bg_color}; font-weight: {font_weight}; font-size: {font_size}; border-top: {border_top}; border-bottom: 1px solid black;">
                            <td style="padding: 8px; border-right: 1px solid black;">{row['PRODUCTION']}</td>
                            <td style="padding: 8px; text-align: right;">{row['VALUE / PCS']}</td>
                        </tr>
                        """, unsafe_allow_html=True
                    )
                st.markdown("</table><br>", unsafe_allow_html=True)
                
            with col_dash2:
                st.subheader("📋 Raw Excel Entry Logs")
                if df.empty:
                    st.warning("Excel me abhi koi data nahi hai.")
                else:
                    st.dataframe(df, hide_index=True, use_container_width=True, height=350)

            # --- PART C: ADMINISTRATIVE CONTROLS ---
            st.markdown("---")
            st.subheader("⚙️ System Control Center (Manage Settings)")
            tab1, tab2 = st.tabs(["👥 Manage Supervisors", "🏢 Manage Dropdowns (Party & Items)"])
            
            with tab1:
                col_u1, col_u2 = st.columns(2)
                with col_u1:
                    with st.form("add_user_form", clear_on_submit=True):
                        st.markdown("**Add New Supervisor Account**")
                        new_id = st.text_input("New Username / ID").strip()
                        new_pass = st.text_input("Set Password").strip()
                        new_name = st.text_input("Supervisor Real Name").strip()
                        if st.form_submit_button("Create User"):
                            if new_id and new_pass and new_name:
                                st.session_state["users"][new_id] = {"password": new_pass, "name": new_name, "role": "supervisor"}
                                st.success(f"🎉 New User '{new_name}' Created!")
                            else:
                                st.error("❌ Saari details bharein.")
                with col_u2:
                    st.markdown("**Active Users Accounts List:**")
                    for u_id, u_info in st.session_state["users"].items():
                        st.write(f"- `{u_id}` : {u_info['name']} ({u_info['role']})")

            with tab2:
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.markdown("**Party Names**")
                    st.write(st.session_state["party_options"])
                    new_party = st.text_input("Add New Party").strip()
                    if st.button("➕ Add Party"):
                        if new_party and new_party not in st.session_state["party_options"]:
                            st.session_state["party_options"].append(new_party)
                            st.success("Party Added!")
                            st.rerun()
                with col_d2:
                    st.markdown("**Item Types**")
                    st.write(st.session_state["item_options"])
                    new_item = st.text_input("Add New Item Type").strip()
                    if st.button("➕ Add Item"):
                        if new_item and new_item not in st.session_state["item_options"]:
                            st.session_state["item_options"].append(new_item.upper())
                            st.success("Item Added!")
                            st.rerun()
