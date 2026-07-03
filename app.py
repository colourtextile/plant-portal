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
import json

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"
CONFIG_FILE = "system_config.json"

st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# --- 💅 PREMIUM GLOBAL STYLING & FIXES ---
st.markdown("""
<style>
    .user-chip {
        display: inline-block;
        background: #0ea5e9;
        color: white;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* ORIGINAL DASHBOARD STYLING */
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

    div[data-testid="stForm"] label, .stMarkdown label, label[data-testid="stWidgetLabel"] {
        color: #2C3E50 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        letter-spacing: 0.5px;
        margin-bottom: 5px !important;
    }
    
    .dashboard-card {
        background: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #eef2f5;
    }
</style>
""", unsafe_allow_html=True)

# --- 💾 PERMANENT DATABASE LOGIC FOR CONFIG & USERS ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    else:
        # Default Configs if file not found
        default_config = {
            "supervisor_targets": {"Ramesh": 500, "Suresh": 400},
            "party_options": ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"],
            "item_options": ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"],
            "email_config": {"sender": "your_email@gmail.com", "password": "your_app_password", "receiver": "receiver_email@gmail.com"},
            "users": {
                "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
                "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False},
                "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
            }
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(default_config, f, indent=4)
        return default_config

def save_config():
    config_data = {
        "supervisor_targets": st.session_state["supervisor_targets"],
        "party_options": st.session_state["party_options"],
        "item_options": st.session_state["item_options"],
        "email_config": st.session_state["email_config"],
        "users": st.session_state["users"]
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config_data, f, indent=4)

# Load into Session State
if "config_loaded" not in st.session_state:
    app_config = load_config()
    st.session_state["supervisor_targets"] = app_config["supervisor_targets"]
    st.session_state["party_options"] = app_config["party_options"]
    st.session_state["item_options"] = app_config["item_options"]
    st.session_state["email_config"] = app_config["email_config"]
    st.session_state["users"] = app_config["users"]
    st.session_state["config_loaded"] = True

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = None

# --- 📧 HELPER FUNCTION TO SEND DAY-WISE DATA EMAIL ---
def send_daywise_backup_email():
    cfg = st.session_state["email_config"]
    current_today = datetime.now().strftime("%d-%m-%Y")
    
    if not os.path.exists(EXCEL_FILE):
        return False
        
    try:
        main_df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        main_df['Date'] = main_df['Date'].astype(str).str.strip()
        day_wise_df = main_df[main_df['Date'] == current_today]
        
        if day_wise_df.empty:
            day_wise_df = pd.DataFrame(columns=main_df.columns)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            day_wise_df.to_excel(writer, index=False, sheet_name="Today Backup")
        excel_buffer.seek(0)

        msg = MIMEMultipart()
        msg['From'] = cfg["sender"]
        msg['To'] = cfg["receiver"]
        msg['Subject'] = f"📍 Colour Textile Day-Wise Backup Report - {current_today}"
        
        part = MIMEBase('application', "octet-stream")
        part.set_payload(excel_buffer.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Production_Backup_{current_today}.xlsx"')
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(cfg["sender"], cfg["password"])
        server.sendmail(cfg["sender"], cfg["receiver"], msg.as_string())
        server.close()
        return True
    except Exception as e:
        return False

# --- 🔒 MAIN APPLICATION LOGIN SCREEN ---
if not st.session_state["logged_in"]:
    st.markdown(f"""
    <style>
        .stApp {{
            background-image: linear-gradient(rgba(0, 0, 0, 0.6), rgba(0, 0, 0, 0.6)), url("https://images.unsplash.com/photo-1617627143750-d86bc21e42bb?q=80&w=1920&auto=format&fit=crop");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        div[data-testid="stForm"] label {{ color: #FFFFFF !important; }}
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
                    st.session_state["current_user"] = st.session_state["users"][username.strip()]
                    st.rerun()
                else:
                    st.error("❌ Invalid ID or Password")

# --- 👑 MAIN PORTAL SCREEN (AFTER LOGIN) ---
else:
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    
    # --- 📐 SIDEBAR BRAND FORMAT & USER DETAILS ---
    st.sidebar.markdown(f"""
    <div class="premium-sidebar-header">
        <div style="color: #38bdf8; font-size: 12px; font-weight: 800; letter-spacing: 2px; margin-bottom: 5px;">WELCOME</div>
        <div style="color: #ffffff; font-size: 20px; font-weight: 900; letter-spacing: 1px;">COLOUR TEXTILE</div>
        <div class="user-chip">{user['role'].upper()}</div>
        <div style="color: #cbd5e1; font-size: 13px; margin-top: 15px; font-weight: 500;">👤 User: {user['name']}</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Navigation items dynamically based on permissions
    nav_options = []
    if user.get("p_view", True) or user["role"] == "admin": nav_options.append("📊 Dashboard")
    if user.get("p_entry", True) or user["role"] == "admin": nav_options.append("📝 Data Entry")
    if user.get("p_edit", False) or user["role"] == "admin": nav_options.append("📋 Edit & Delete Records")
    
    if user["role"] == "admin":
        nav_options.extend(["🎯 Target Settings", "📧 Setup Email Auto-Backup", "⚙️ Factory Config"])
        
    nav_choice = st.sidebar.radio("🧭 Navigation Menu", nav_options)
        
    excel_loaded = False
    df = pd.DataFrame()
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
            if not df.empty:
                if df.shape[1] == 8:
                    df.columns = ["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor"]
                    df["Challan No"] = "-" 
                elif df.shape[1] == 9:
                    df.columns = ["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor", "Challan No"]
                ordered_cols = ["Date", "Challan No", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor"]
                df = df[ordered_cols]
                excel_loaded = True
        except:
            pass

    # --- ⏰ 5 TIMES AUTO EMAIL BACKGROUND TRIGGER ---
    now_time = datetime.now().strftime("%H:%M")
    target_slots = ["10:00", "13:00", "16:00", "19:00", "22:00"]
    if now_time in target_slots:
        if "last_triggered_slot" not in st.session_state or st.session_state["last_triggered_slot"] != now_time:
            send_daywise_backup_email()
            st.session_state["last_triggered_slot"] = now_time

    if st.sidebar.button("🚪 Logout System"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()

    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili! System ke folders check karein.")
    else:
        # --- 📧 EMAIL AUTO-BACKUP CONFIG PANEL ---
        if user["role"] == "admin" and nav_choice == "📧 Setup Email Auto-Backup":
            st.subheader("📧 Day-Wise Auto Email Schedule Configurations (5 Times Daily)")
            st.info("System automatic niche diye gaye samay par sirf usi din ka filtered data mail karega:\n\n ⏰ **10:00 AM | 01:00 PM | 04:00 PM | 07:00 PM | 10:00 PM**")
            
            with st.form("mail_config_form"):
                sender = st.text_input("Sender Gmail Address", st.session_state["email_config"]["sender"])
                password = st.text_input("Sender Google App Password (16-Digit)", st.session_state["email_config"]["password"], type="password")
                receiver = st.text_input("Receiver Backup Email Address", st.session_state["email_config"]["receiver"])
                
                if st.form_submit_button("Save Email Settings"):
                    st.session_state["email_config"] = {"sender": sender, "password": password, "receiver": receiver}
                    save_config()  # Save permanently
                    st.success("✅ Email Configuration Settings Saved Locally!")
            
            if st.button("🚀 Test Send Instant Day-wise Data Now"):
                with st.spinner("Sending Day-wise Excel file attachment..."):
                    if send_daywise_backup_email():
                        st.success("🎉 Email Sent Successfully with today's day-wise data filter!")
                    else:
                        st.error("❌ Email transmission failed. Please check credentials or App password.")

        # --- TARGET SETTINGS LOGIC ---
        elif user["role"] == "admin" and nav_choice == "🎯 Target Settings":
            st.subheader("🎯 Supervisor Target Configurations")
            all_users = list(st.session_state["users"].keys())
            sups_list = [st.session_state["users"][u]["name"] for u in all_users if st.session_state["users"][u]["role"] == "supervisor"]
            if sups_list:
                for sup_name in sups_list:
                    current_tgt = st.session_state["supervisor_targets"].get(sup_name, 500)
                    new_tgt = st.number_input(f"Set Daily Target Pcs for {sup_name}", min_value=0, value=int(current_tgt), key=f"tgt_{sup_name}")
                    st.session_state["supervisor_targets"][sup_name] = new_tgt
                
                if st.button("💾 Save All Targets"):
                    save_config() # Save permanently
                    st.success("🎯 Targets Saved Successfully!")

        # --- DATA ENTRY LOGIC ---
        elif nav_choice == "📝 Data Entry":
            st.subheader(f"📝 Quick Data Entry Panel ({user['name']})")
            with st.form("entry_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    date_input = st.date_input("Select Date", datetime.now())
                    challan_no = st.text_input("Enter Challan Number", placeholder="e.g. CH-9921")
                    design_no = st.text_input("Enter Design Number", placeholder="e.g. D-401")
                    party_name = st.selectbox("Select Party Name Account", st.session_state["party_options"])
                with col2:
                    item_type = st.selectbox("Select Fabric / Item Type", st.session_state["item_options"])
                    total_pcs = st.number_input("Total Pieces Count", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Quality Pieces", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Damage Pieces", min_value=0, step=1)
                
                if st.form_submit_button("SAVE ENTRY"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Total mismatch! (Fresh + Seconds) must be equal to Total Pcs.")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully!")

        # --- EDIT & DELETE (ROW-WISE FORM SYSTEM) ---
        elif nav_choice == "📋 Edit & Delete Records":
            st.subheader("📋 Edit & Delete Logged Records (Row-Wise)")
            st.info("💡 **Kaise use karein:** Niche drop-down list se jis entry (row) ko aapko edit ya delete karna hai use select karein. Uski puri details form mein aa jayengi. Change karke Update ya Delete dabayein.")
            
            if excel_loaded and not df.empty:
                if user["role"] != "admin":
                    display_df = df[df["Supervisor"] == user["name"]].copy()
                else:
                    display_df = df.copy()
                    
                if not display_df.empty:
                    options = []
                    for idx, row in display_df.iterrows():
                        options.append(f"Row ID: {idx} | Date: {row['Date']} | Design: {row['Design No']} | Party: {row['Party Name']} | Pcs: {row['Total Pcs']}")
                    
                    selected_option = st.selectbox("🔍 Select Entry to Modify:", options)
                    selected_idx = int(selected_option.split(" | ")[0].replace("Row ID: ", ""))
                    row_data = df.loc[selected_idx]
                    
                    st.markdown(f"### ✏️ Update Entry Details (Row ID: {selected_idx})")
                    
                    with st.form("edit_row_form"):
                        col1, col2 = st.columns(2)
                        with col1:
                            e_date = st.text_input("Date (DD-MM-YYYY)", value=str(row_data["Date"]))
                            e_challan = st.text_input("Challan No", value=str(row_data["Challan No"]))
                            e_design = st.text_input("Design No", value=str(row_data["Design No"]))
                            
                            p_val = str(row_data["Party Name"])
                            p_index = st.session_state["party_options"].index(p_val) if p_val in st.session_state["party_options"] else 0
                            e_party = st.selectbox("Party Name", st.session_state["party_options"], index=p_index)
                            
                        with col2:
                            i_val = str(row_data["Item Type"])
                            i_index = st.session_state["item_options"].index(i_val) if i_val in st.session_state["item_options"] else 0
                            e_item = st.selectbox("Item Type", st.session_state["item_options"], index=i_index)
                            
                            e_total = st.number_input("Total Pcs", value=int(row_data["Total Pcs"]), step=1)
                            e_fresh = st.number_input("Fresh Pcs", value=int(row_data["Fresh Pcs"]), step=1)
                            e_seconds = st.number_input("Seconds Pcs", value=int(row_data["Seconds Pcs"]), step=1)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            update_btn = st.form_submit_button("💾 UPDATE THIS RECORD", use_container_width=True)
                        with btn_col2:
                            delete_btn = st.form_submit_button("❌ DELETE THIS RECORD", use_container_width=True)
                            
                    if update_btn:
                        if e_fresh + e_seconds != e_total:
                            st.error("❌ Total mismatch! (Fresh + Seconds) must be equal to Total Pcs.")
                        else:
                            df.at[selected_idx, "Date"] = e_date
                            df.at[selected_idx, "Challan No"] = e_challan
                            df.at[selected_idx, "Design No"] = e_design
                            df.at[selected_idx, "Party Name"] = e_party
                            df.at[selected_idx, "Item Type"] = e_item
                            df.at[selected_idx, "Total Pcs"] = e_total
                            df.at[selected_idx, "Fresh Pcs"] = e_fresh
                            df.at[selected_idx, "Seconds Pcs"] = e_seconds
                            
                            df.to_excel(EXCEL_FILE, sheet_name="Supervisor Entry", index=False)
                            st.success("✅ Record successfully updated in Database!")
                            st.rerun()
                            
                    if delete_btn:
                        df = df.drop(index=selected_idx)
                        df.to_excel(EXCEL_FILE, sheet_name="Supervisor Entry", index=False)
                        st.warning("🗑️ Record successfully deleted from Database!")
                        st.rerun()
                        
                else:
                    st.warning("No records available to edit for your account.")
            else:
                st.warning("Excel file is empty or missing. Please make entries first.")

        # --- CONFIGURATIONS DESK FOR ADMIN ---
        elif user["role"] == "admin" and nav_choice == "⚙️ Factory Config":
            st.subheader("⚙️ Factory Configurations Desk")
            t1, t2, t3 = st.tabs(["🏢 Manage Parties", "📦 Manage Items", "👥 Supervisors Accounts"])
            
            with t1:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**➕ Add New Party**")
                    new_party = st.text_input("Enter Party Name to Add", placeholder="e.g. Balaji Corp").strip()
                    if st.button("Save New Party"):
                        if new_party and new_party not in st.session_state["party_options"]:
                            st.session_state["party_options"].append(new_party)
                            save_config() # Save permanently
                            st.success(f"Added: {new_party}")
                            st.rerun()
                with col2:
                    st.markdown("**✏️ Edit Party Name**")
                    party_to_edit = st.selectbox("Select Target Party to Edit", st.session_state["party_options"], key="edt_p")
                    edited_party_name = st.text_input("Enter New Modified Name", value=party_to_edit)
                    if st.button("Update Party Name"):
                        if edited_party_name and party_to_edit:
                            idx = st.session_state["party_options"].index(party_to_edit)
                            st.session_state["party_options"][idx] = edited_party_name
                            save_config() # Save permanently
                            st.success("Updated!")
                            st.rerun()
                with col3:
                    st.markdown("**❌ Remove Party**")
                    party_to_remove = st.selectbox("Select Party to Delete", st.session_state["party_options"], key="rem_p")
                    if st.button("Delete Party From System", type="primary"):
                        if party_to_remove in st.session_state["party_options"]:
                            st.session_state["party_options"].remove(party_to_remove)
                            save_config() # Save permanently
                            st.rerun()

            with t2:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**➕ Add New Item Type**")
                    new_item = st.text_input("Enter Item Name to Add", placeholder="e.g. SILK SAREE").strip()
                    if st.button("Save New Item"):
                        if new_item and new_item.upper() not in st.session_state["item_options"]:
                            st.session_state["item_options"].append(new_item.upper())
                            save_config() # Save permanently
                            st.success("Added!")
                            st.rerun()
                with col2:
                    st.markdown("**✏️ Edit Item Name**")
                    item_to_edit = st.selectbox("Select Item to Edit from List", st.session_state["item_options"], key="edt_i")
                    edited_item_name = st.text_input("Enter New Name for Selected Item", value=item_to_edit)
                    if st.button("Update Item Name"):
                        if edited_item_name and item_to_edit:
                            idx = st.session_state["item_options"].index(item_to_edit)
                            st.session_state["item_options"][idx] = edited_item_name.upper()
                            save_config() # Save permanently
                            st.success("Updated!")
                            st.rerun()
                with col3:
                    st.markdown("**❌ Remove Item Type**")
                    item_to_remove = st.selectbox("Select Item to Delete from System", st.session_state["item_options"], key="rem_i")
                    if st.button("Delete Item Category", type="primary"):
                        if item_to_remove in st.session_state["item_options"]:
                            st.session_state["item_options"].remove(item_to_remove)
                            save_config() # Save permanently
                            st.rerun()

            with t3:
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    st.markdown("**➕ Add New Supervisor Account**")
                    with st.form("add_user_form", clear_on_submit=True):
                        add_id = st.text_input("Set Login ID / Username", placeholder="e.g. ramesh01").strip()
                        add_pass = st.text_input("Set Account Password", type="password", placeholder="••••••••").strip()
                        add_name = st.text_input("Supervisor Full Real Name", placeholder="e.g. Ramesh Kumar").strip()
                        
                        st.markdown("⚠️ **Set Dynamic Custom Permissions:**")
                        cb_entry = st.checkbox("Allow Data Entry Form Access", value=True)
                        cb_view = st.checkbox("Allow View Production Logs", value=True)
                        cb_edit = st.checkbox("Allow Edit/Delete Logged Records", value=False)
                        
                        if st.form_submit_button("Create Account"):
                            if add_id and add_pass and add_name:
                                if add_id not in st.session_state["users"]:
                                    st.session_state["users"][add_id] = {
                                        "password": add_pass, 
                                        "name": add_name, 
                                        "role": "supervisor",
                                        "p_entry": cb_entry,
                                        "p_view": cb_view,
                                        "p_edit": cb_edit
                                    }
                                    st.session_state["supervisor_targets"][add_name] = 500 # Default target
                                    save_config() # Save permanently
                                    st.success(f"Supervisor '{add_name}' Created & Saved Permanently!")
                                    st.rerun()
                            
                with col_u2:
                    st.markdown("**✏️ Edit Info & Checkbox Permissions**")
                    all_users = list(st.session_state["users"].keys())
                    sups_only = [u for u in all_users if st.session_state["users"][u]["role"] == "supervisor"]
                    
                    if sups_only:
                        selected_sup = st.selectbox("Select Supervisor ID to Modify", sups_only, key="sel_sup_edt")
                        current_sup_data = st.session_state["users"][selected_sup]
                        
                        edit_name = st.text_input("Edit Full Name Display", value=current_sup_data["name"])
                        edit_pass = st.text_input("Edit Security Password", value=current_sup_data["password"])
                        
                        st.markdown("⚙️ **Update Checkbox Permissions:**")
                        edit_cb_entry = st.checkbox("Allow Data Entry Form Access", value=current_sup_data.get("p_entry", True), key="ed_e")
                        edit_cb_view = st.checkbox("Allow View Production Logs", value=current_sup_data.get("p_view", True), key="ed_v")
                        edit_cb_edit = st.checkbox("Allow Edit/Delete Logged Records", value=current_sup_data.get("p_edit", False), key="ed_d")
                        
                        if st.button("Update Supervisor Account"):
                            if edit_name and edit_pass:
                                old_name = current_sup_data["name"]
                                st.session_state["users"][selected_sup]["name"] = edit_name
                                st.session_state["users"][selected_sup]["password"] = edit_pass
                                st.session_state["users"][selected_sup]["p_entry"] = edit_cb_entry
                                st.session_state["users"][selected_sup]["p_view"] = edit_cb_view
                                st.session_state["users"][selected_sup]["p_edit"] = edit_cb_edit
                                
                                # Update targets dict if name changed
                                if old_name != edit_name and old_name in st.session_state["supervisor_targets"]:
                                    st.session_state["supervisor_targets"][edit_name] = st.session_state["supervisor_targets"].pop(old_name)
                                
                                save_config() # Save permanently
                                st.success("Account & Checkbox Permissions updated Permanently!")
                                st.rerun()
                        
                with col_u3:
                    st.markdown("**❌ Remove Supervisor**")
                    if sups_only:
                        sup_to_remove = st.selectbox("Select Supervisor ID to Delete", sups_only, key="sel_sup_rem")
                        if st.button("Delete Supervisor Account", type="primary"):
                            sup_name_to_del = st.session_state["users"][sup_to_remove]["name"]
                            del st.session_state["users"][sup_to_remove]
                            if sup_name_to_del in st.session_state["supervisor_targets"]:
                                del st.session_state["supervisor_targets"][sup_name_to_del]
                            save_config() # Save permanently
                            st.warning("Account deleted from database permanently!")
                            st.rerun()

        # --- DASHBOARD LOGIC ---
        elif nav_choice == "📊 Dashboard":
            st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Live Analytics Dashboard</h2>", unsafe_allow_html=True)
            
            filter_type = st.radio("📅 Select Dashboard Filter Range:", ["☀️ Day-Wise Filter", "📆 Month-Wise Filter"], horizontal=True)
            
            filtered_df_by_range = pd.DataFrame()
            display_range_label = ""
            
            if excel_loaded and not df.empty:
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                df['Date'] = df['Date'].astype(str).str.strip()
                df['parsed_date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
                df['Month_Year'] = df['parsed_date'].dt.strftime('%B %Y')
            
            if filter_type == "☀️ Day-Wise Filter":
                selected_filter_date = st.date_input("Choose Specific Date to View Summary", datetime.now())
                formatted_filter_date = selected_filter_date.strftime("%d-%m-%Y")
                display_range_label = f"SELECTED DATE: {formatted_filter_date}"
                
                if excel_loaded and not df.empty:
                    filtered_df_by_range = df[df["Date"] == formatted_filter_date]
            
            else:
                current_month_year = datetime.now().strftime('%B %Y')
                available_months = []
                if excel_loaded and not df.empty:
                    available_months = sorted(list(df['Month_Year'].dropna().unique()))
                
                if current_month_year not in available_months:
                    available_months.append(current_month_year)
                    
                selected_month = st.selectbox("Choose Month and Year to View Summary", available_months, 
                                             index=available_months.index(current_month_year) if current_month_year in available_months else 0)
                display_range_label = f"SELECTED MONTH: {selected_month.upper()}"
                
                if excel_loaded and not df.empty:
                    filtered_df_by_range = df[df["Month_Year"] == selected_month]

            if not filtered_df_by_range.empty:
                item_groups = filtered_df_by_range.groupby('Item Type')['Total Pcs'].sum().to_dict()
                party_groups = filtered_df_by_range.groupby('Party Name')['Total Pcs'].sum().to_dict()
            else:
                item_groups = {}
                party_groups = {}

            items_list = [it.upper().strip() for it in st.session_state["item_options"]]
            party_list = st.session_state["party_options"]

            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.markdown(f"""
                <div style="background-color: #FFFF00; color: #000000; text-align: center; font-weight: bold; padding: 12px; font-size: 20px; border: 2px solid #2c3e50; border-bottom: none; border-radius: 6px 6px 0px 0px;">
                    🏭 ITEM PIECES SUMMARY
                </div>
                <div style="background-color: #1F4E79; color: #FFFFFF; padding: 10px 15px; font-weight: bold; font-size: 15px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 2px solid #2c3e50; text-align: center;">
                    {display_range_label}
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

                st.markdown('<div class="dashboard-card" style="margin-top: 20px;">', unsafe_allow_html=True)
                st.markdown(f"""
                <div style="background-color: #FFA500; color: #000000; text-align: center; font-weight: bold; padding: 12px; font-size: 20px; border: 2px solid #2c3e50; border-bottom: none; border-radius: 6px 6px 0px 0px;">
                    🏢 PARTY PIECES SUMMARY
                </div>
                <div style="background-color: #1F4E79; color: #FFFFFF; padding: 10px 15px; font-weight: bold; font-size: 15px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 2px solid #2c3e50; text-align: center;">
                    {display_range_label}
                </div>
                """, unsafe_allow_html=True)
                
                party_total_sum = 0
                for pt in party_list:
                    val = party_groups.get(pt, 0)
                    if val > 0:
                        party_total_sum += val
                        st.markdown(f"""
                        <div style="background-color: #FFFFFF; color: #333333; padding: 10px 15px; font-size: 14px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 1px solid #EAEAEA; display: flex; justify-content: space-between;">
                            <span>{pt}</span><b>{val:,} Pcs</b>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown(f"""
                <div style="background-color: #27AE60; color: #FFFFFF; padding: 12px 15px; font-weight: bold; font-size: 18px; border: 2px solid #2c3e50; border-radius: 0px 0px 6px 6px; display: flex; justify-content: space-between;">
                    <span>📊 TOTAL PRODUCED</span><u>{party_total_sum:,} Pcs</u>
                </div>
                """, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
            with col_right:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.markdown(f"### 🍩 Share Matrix Graph ({display_range_label.replace('SELECTED ', '')})")
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
                    st.info(f"Is range ({display_range_label.replace('SELECTED ', '')}) me graph ke liye koi data entry nahi mili.")
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("---")

            st.markdown("### 🎯 Supervisor Live Target Tracker (Auto-Less Status)")
            target_summary_data = []
            
            current_today_str = datetime.now().strftime("%d-%m-%Y")
            if excel_loaded and not df.empty:
                today_df = df[df["Date"] == current_today_str]
            else:
                today_df = pd.DataFrame()

            for u_id, u_info in st.session_state["users"].items():
                if u_info["role"] == "supervisor":
                    s_name = u_info["name"]
                    allocated_tgt = st.session_state["supervisor_targets"].get(s_name, 0)
                    
                    if not today_df.empty:
                        done_pcs = today_df[today_df["Supervisor"] == s_name]["Total Pcs"].sum()
                    else:
                        done_pcs = 0
                        
                    pending = allocated_tgt - done_pcs
                    target_summary_data.append({
                        "Supervisor Name": s_name, 
                        "Daily Target Allocated": allocated_tgt, 
                        "Completed Today": done_pcs, 
                        "Pending Target": pending if pending > 0 else 0,
                        "Status": "✅ Target Achieved" if pending <= 0 else "⏳ Work in Progress"
                    })
                    
            if target_summary_data:
                st.dataframe(pd.DataFrame(target_summary_data), use_container_width=True)
            else:
                st.info("No active supervisor targets found in the system.")
