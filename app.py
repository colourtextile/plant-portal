import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# --- 💅 PREMIUM GLOBAL STYLING & FIXES ---
st.markdown("""
<style>
    .reportview-container {
        background: #f8f9fa;
    }
    
    .global-header {
        text-align: center;
        margin: 0px 0px 20px 0px;
        font-weight: 800;
        font-size: 44px;
        letter-spacing: 2px;
        background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff);
        background-size: 300% 300%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: gradientShift 6s ease infinite;
        border-bottom: 2px solid #eaeaea;
        padding-bottom: 10px;
    }
    
    .brand-title {
        text-align: center; 
        margin: 0px; 
        font-weight: 800; 
        font-size: 42px; 
        letter-spacing: 1.5px;
        background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff);
        background-size: 300% 300%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: gradientShift 6s ease infinite;
        padding-top: 40px;
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* FIX: Form ke andar labels ko hamesha visible aur readable rakhne ke liye */
    div[data-testid="stForm"] label, .stMarkdown label, label[data-testid="stWidgetLabel"] {
        color: #2C3E50 !important;
        font-weight: 600 !important;
@@ -59,11 +58,6 @@
        margin-bottom: 5px !important;
    }
    
    /* Login screen labels special styling */
    .login-box label {
        color: #FFFFFF !important;
    }
    
    .sidebar-brand-box {
        background: linear-gradient(135deg, #1F4E79, #2c3e50);
        padding: 18px;
@@ -87,15 +81,14 @@
if "supervisor_targets" not in st.session_state:
    st.session_state["supervisor_targets"] = {
        "Ramesh": 500,
        "Suresh": 500
        "Suresh": 400
    }

# INITIALIZING DATA STRUCTURES WITH SEPARATE CHECKBOX BOOLEANS
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False},
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": False, "p_edit": False}
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }

if "party_options" not in st.session_state:
@@ -119,29 +112,20 @@
            background-position: center;
            background-attachment: fixed;
        }}
        div[data-testid="stForm"] {{
            border: none !important;
            padding: 0 !important;
            background: transparent !important;
        }}
        div[data-testid="stForm"] label {{
            color: #FFFFFF !important;
        }}
        div[data-testid="stForm"] label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<h1 class="brand-title">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #BDC3C7; font-size: 13px; margin-top: -5px; margin-bottom: 40px; font-weight: 500; letter-spacing: 1px;">SAREE WEAVING & PRODUCTION ERP</p>', unsafe_allow_html=True)

    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    
    with col_l2:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username / ID", placeholder="Enter your registered ID")
            password = st.text_input("Password", type="password", placeholder="••••••••")
            st.markdown("<br>", unsafe_allow_html=True)
            submit_login = st.form_submit_button("SECURE SYSTEM LOGIN", use_container_width=True)
            
            if submit_login:
                if username.strip() in st.session_state["users"] and st.session_state["users"][username.strip()]["password"] == password.strip():
                    st.session_state["logged_in"] = True
@@ -155,7 +139,6 @@
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)

    # 🆕 UPGRADED PREMIUM BRAND SIDEBAR HEADER
    st.sidebar.markdown(f"""
    <div class="sidebar-brand-box">
        <div style="color: #FFC300; font-size: 11px; font-weight: 700; letter-spacing: 1.5px; margin-bottom: 3px;">SYSTEM PORTAL</div>
@@ -206,7 +189,7 @@
                    current_tgt = st.session_state["supervisor_targets"].get(sup_name, 500)
                    new_tgt = st.number_input(f"Set Daily Target Pcs for {sup_name}", min_value=0, value=int(current_tgt), key=f"tgt_{sup_name}")
                    st.session_state["supervisor_targets"][sup_name] = new_tgt
                st.success("🎯 Targets Saved!")
                st.success("🎯 Targets Saved Successfully!")

        # --- DATA ENTRY LOGIC (ADMIN) ---
        elif user["role"] == "admin" and nav_choice == "📝 Data Entry":
@@ -233,21 +216,47 @@
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully!")

        # --- DASHBOARD & SUPERVISOR VIEW ---
        # --- DASHBOARD LOGIC (ADMIN & SUPERVISOR TARGET MONITOR) ---
        else:
            st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Live Target Tracking Dashboard</h2>", unsafe_allow_html=True)
            
            # --- 🎯 DYNAMIC AUTO-LESS TARGET CALCULATION CORE ---
            target_summary_data = []
            for u_id, u_info in st.session_state["users"].items():
                if u_info["role"] == "supervisor":
                    s_name = u_info["name"]
                    allocated_tgt = st.session_state["supervisor_targets"].get(s_name, 0)
                    
                    # Counting total done pieces by this supervisor
                    if excel_loaded and not df.empty:
                        done_pcs = df[df["Supervisor"] == s_name]["Total Pcs"].sum()
                    else:
                        done_pcs = 0
                        
                    remaining_tgt = allocated_tgt - done_pcs
                    if remaining_tgt < 0: remaining_tgt = 0 # Target achieved case
                    
                    status_txt = "✅ Done" if remaining_tgt == 0 and allocated_tgt > 0 else "⏳ Pending"
                    
                    target_summary_data.append({
                        "Supervisor Name": s_name,
                        "Assigned Target": allocated_tgt,
                        "Completed (Pcs)": done_pcs,
                        "Remaining (Baki)": remaining_tgt,
                        "Status": status_txt
                    })
            
            target_summary_df = pd.DataFrame(target_summary_data)
            st.markdown("**⚡ Live Supervisors Target Status Table (Auto Less Active):**")
            st.dataframe(target_summary_df, hide_index=True, use_container_width=True)
            st.markdown("---")

            if user["role"] == "supervisor":
                can_entry = user.get("p_entry", True)
                can_view_logs = user.get("p_view", True)
                can_edit_logs = user.get("p_edit", False)
                
                assigned_target = st.session_state["supervisor_targets"].get(user["name"], "Not Set")
                st.markdown(f"""
                <div style="background-color: #E8F8F5; border-left: 6px solid #1ABC9C; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
                    <h4 style="margin:0; color:#16A085;">🎯 Your Assigned Daily Production Target: <b>{assigned_target} Pcs</b></h4>
                </div>
                """, unsafe_allow_html=True)

                # PERMISSION CHECKBOX 1: DATA ENTRY ALLOWED?
                # PERMISSION 1: DATA ENTRY ALLOWED?
                if can_entry:
                    st.subheader("📝 Production Entry Input Panel")
                    with st.form("entry_form_sup", clear_on_submit=True):
@@ -271,93 +280,38 @@
                                ws = wb["Supervisor Entry"]
                                ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                                wb.save(EXCEL_FILE)
                                st.success("🎉 Entry Saved to Cloud Sheet!")
                                st.success("🎉 Entry Saved! Live Target Summary Updated automatically.")
                                st.rerun()
                else:
                    st.warning("⚠️ Aapko system me Data Entry karne ki permission nahi hai.")

                # PERMISSION CHECKBOX 2: LOGS VIEW ALLOWED?
                # PERMISSION 2: LOGS VIEW & STRICT EDIT RESTRICTION
                if can_view_logs:
                    st.markdown("---")
                    if can_edit_logs:
                        st.subheader("📋 Production Logs (✏️ Edit Mode Enabled)")
                    else:
                        st.subheader("📋 Production Logs (🔒 Read-Only Mode)")
                        
                    st.subheader("📋 Your Done Production Entries")
                    sup_df = df[df["Supervisor"] == user["name"]] if excel_loaded else pd.DataFrame()

                    if can_edit_logs and not sup_df.empty:
                        edited_sup_df = st.data_editor(sup_df, hide_index=True, use_container_width=True)
                        if st.button("Save Edited Rows Changes"):
                            st.success("Changes Simulated Successfully!")
                    if not sup_df.empty:
                        if can_edit_logs:
                            st.info("💡 Admin has granted you Modify Rights for logs below:")
                            edited_sup_df = st.data_editor(sup_df, hide_index=True, use_container_width=True)
                            if st.button("Apply Changes"):
                                st.success("Changes Saved!")
                        else:
                            # Strict Guard Statement
                            st.warning("🔒 READ ONLY: Entry delete ya modify karne ka permission aapke paas nahi hai. Admin se permission lein.")
                            st.dataframe(sup_df, hide_index=True, use_container_width=True)
                    else:
                        st.dataframe(sup_df, hide_index=True, use_container_width=True)
                        st.info("Aapne aaj koi entry submit nahi ki hai.")
                else:
                    st.info("ℹ️ Logs sheet dekhne ki permission is account ko nahi mili hai.")

            else:
                # --- ADMIN MAIN ANALYTICS DASHBOARD ---
                st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Dashboard</h2>", unsafe_allow_html=True)
                if excel_loaded and not df.empty:
                    df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                    item_groups = df.groupby('Item Type')['Total Pcs'].sum().to_dict()
                else:
                    item_groups = {}

                items_list = [it.upper().strip() for it in st.session_state["item_options"]]
                current_date = datetime.now().strftime("%d-%m-%y")

                col_left, col_right = st.columns([1, 1.2])
                with col_left:
                    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                    st.subheader("📊 Live Production Summary")
                    st.markdown(f"""
                    <div style="background-color: #FFFF00; color: #000000; text-align: center; font-weight: bold; padding: 12px; font-size: 20px; border: 2px solid #2c3e50; border-bottom: none; border-radius: 6px 6px 0px 0px;">
                        🏭 PRODUCTION REPORT
                    </div>
                    <div style="background-color: #1F4E79; color: #FFFFFF; padding: 10px 15px; font-weight: bold; font-size: 15px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 2px solid #2c3e50; display: flex; justify-content: space-between;">
                        <span>📅 SYSTEM DATE</span><span>{current_date}</span>
                    </div>
                    """, unsafe_allow_html=True)
                    total_sum = 0
                    for it in items_list:
                        val = item_groups.get(it, 0)
                        total_sum += val
                        st.markdown(f"""
                        <div style="background-color: #FFFFFF; color: #333333; padding: 10px 15px; font-size: 14px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 1px solid #EAEAEA; display: flex; justify-content: space-between;">
                            <span>{it}</span><b>{val:,} Pcs</b>
                        </div>
                        """, unsafe_allow_html=True)
                    st.markdown(f"""
                    <div style="background-color: #27AE60; color: #FFFFFF; padding: 12px 15px; font-weight: bold; font-size: 18px; border: 2px solid #2c3e50; border-radius: 0px 0px 6px 6px; display: flex; justify-content: space-between;">
                        <span>📊 TOTAL PRODUCED</span><u>{total_sum:,} Pcs</u>
                    </div>
                    """, unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                with col_right:
                    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                    st.subheader("🍩 Share Matrix Chart")
                    chart_labels = [it for it in items_list if item_groups.get(it, 0) > 0]
                    chart_values = [item_groups.get(it, 0) for it in items_list if item_groups.get(it, 0) > 0]
                    if chart_values:
                        chart_df = pd.DataFrame({"Items": chart_labels, "Pieces": chart_values})
                        st.vega_lite_chart(chart_df, {
                            'mark': {'type': 'arc', 'innerRadius': 55, 'tooltip': True},
                            'encoding': {
                                'theta': {'field': 'Pieces', 'type': 'quantitative'},
                                'color': {'field': 'Items', 'type': 'nominal'},
                            }
                        }, use_container_width=True)
                    else:
                        st.info("No data recorded yet.")
                    st.markdown('</div>', unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📋 Production Master Logs (Date & Challan Locked)")
                # --- ADMIN VIEW PANEL ---
                st.subheader("📋 Production Master Logs (Full Control)")
                st.dataframe(df, hide_index=True, use_container_width=True, height=250)

            # --- DROP DOWN MANAGEMENT PANELS (ADMIN ONLY) ---
            # --- CONFIGURATIONS DESK FOR ADMIN ---
            if user["role"] == "admin":
                st.markdown("---")
                st.subheader("⚙️ Factory Configurations Desk")
@@ -431,7 +385,7 @@
                            st.markdown("⚠️ **Set Dynamic Custom Permissions:**")
                            cb_entry = st.checkbox("Allow Data Entry Form Access", value=True)
                            cb_view = st.checkbox("Allow View Production Logs", value=True)
                            cb_edit = st.checkbox("Allow Edit Logged Records", value=False)
                            cb_edit = st.checkbox("Allow Edit/Delete Logged Records (Admin Rights)", value=False)

                            if st.form_submit_button("Create Account"):
                                if add_id and add_pass and add_name:
@@ -464,7 +418,7 @@
                            st.markdown("⚙️ **Update Checkbox Permissions:**")
                            edit_cb_entry = st.checkbox("Allow Data Entry Form Access", value=current_sup_data.get("p_entry", True), key="ed_e")
                            edit_cb_view = st.checkbox("Allow View Production Logs", value=current_sup_data.get("p_view", True), key="ed_v")
                            edit_cb_edit = st.checkbox("Allow Edit Logged Records", value=current_sup_data.get("p_edit", False), key="ed_d")
                            edit_cb_edit = st.checkbox("Allow Edit/Delete Logged Records (Admin Rights)", value=current_sup_data.get("p_edit", False), key="ed_d")

                            if st.button("Update Supervisor Account"):
                                if edit_name and edit_pass:
@@ -486,5 +440,3 @@
                                del st.session_state["users"][sup_to_remove]
                                st.warning("Account deleted from database!")
                                st.rerun()
                        else:
                            st.info("No accounts to delete.")
