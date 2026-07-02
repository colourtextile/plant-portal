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

    div[data-testid="stForm"] label, .stMarkdown label, label[data-testid="stWidgetLabel"] {
        color: #2C3E50 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        letter-spacing: 0.5px;
        margin-bottom: 5px !important;
    }
    
    .sidebar-brand-box {
        background: linear-gradient(135deg, #1F4E79, #2c3e50);
        padding: 18px;
        border-radius: 8px;
        margin-bottom: 25px;
        border-left: 5px solid #FF5733;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
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

# --- SESSION STATES INIT ---
if "supervisor_targets" not in st.session_state:
    st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}

if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False},
        "suresh02": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }

if "party_options" not in st.session_state:
    st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]

if "item_options" not in st.session_state:
    st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = None

# Email Configuration states
if "email_config" not in st.session_state:
    st.session_state["email_config"] = {
        "sender": "your_email@gmail.com",
        "password": "your_app_password",
        "receiver": "receiver_email@gmail.com"
    }

# --- 📧 HELPER FUNCTION TO SEND DAY-WISE DATA EMAIL ---
def send_daywise_backup_email():
    cfg = st.session_state["email_config"]
    current_today = datetime.now().strftime("%d-%m-%Y")
    
    if not os.path.exists(EXCEL_FILE):
        return False
        
    try:
        # Load main file and filter only today's records
        main_df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        main_df['Date'] = main_df['Date'].astype(str).str.strip()
        day_wise_df = main_df[main_df['Date'] == current_today]
        
        if day_wise_df.empty:
            # Agar us din ka koi data nahi hai to khali format file bhej sakte hain ya skip kar sakte hain
            day_wise_df = pd.DataFrame(columns=main_df.columns)

        # Create temporary in-memory excel for today's data only
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            day_wise_df.to_excel(writer, index=False, sheet_name="Today Backup")
        excel_buffer.seek(0)

        # Email Setup
        msg = MIMEMultipart()
        msg['From'] = cfg["sender"]
        msg['To'] = cfg["receiver"]
        msg['Subject'] = f"📍 Colour Textile Day-Wise Backup Report - {current_today}"
        
        # Attach File
        part = MIMEBase('application', "octet-stream")
        part.set_payload(excel_buffer.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Production_Backup_{current_today}.xlsx"')
        msg.attach(part)
        
        # SMTP Server logic
        server = smtplib.SMTP('smtp.gmail.com', 557)
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
                    st.session_state["current_user"] = st.session_state["users"][username.strip()]
                    st.rerun()
                else:
                    st.error("❌ Invalid ID or Password")

# --- 👑 MAIN PORTAL SCREEN (AFTER LOGIN) ---
else:
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    
    # --- 📐 UPDATED SIDEBAR BRAND FORMAT & USER DETAILS ---
    st.sidebar.markdown(f"""
    <div class="sidebar-brand-box">
        <div style="color: #FFC300; font-size: 20px; font-weight: 800; letter-spacing: 2px; margin-bottom: -5px;">WELCOME</div>
        <div style="color: #FFFFFF; font-size: 22px; font-weight: 900; letter-spacing: 1px;">COLOUR TEXTILE</div>
        <div style="border-top: 1px solid rgba(255,255,255,0.2); margin: 12px 0px;"></div>
        <div style="color: #E0E6ED; font-size: 14px; font-weight: 500; text-align: left; padding-left: 5px;">👤 User: <span style="color: #30c381; font-weight: 700;">{user['name']}</span></div>
        <div style="color: #B4C6E7; font-size: 12px; margin-top: 2px; text-align: left; padding-left: 5px;">📋 Access Level: <b>{user['role'].upper()}</b></div>
    </div>
    """, unsafe_allow_html=True)
    
    # Navigation items including Email Setup
    nav_options = ["📊 Dashboard"]
    if user["role"] == "admin":
        nav_options.extend(["📝 Data Entry", "🎯 Target Settings", "📧 Setup Email Auto-Backup"])
        
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
    # Trigger slots check automatically on dashboard interactions
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
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili!")
    else:
        # --- 📧 EMAIL AUTO-BACKUP CONFIG PANEL ---
        if user["role"] == "admin" and nav_choice == "📧 Setup Email Auto-Backup":
            st.subheader("📧 Day-Wise Auto Email Schedule Configurations (5 Times Daily)")
            st.info("System automatic niche diye gaye samay par sirf usi din ka filtered data mail karega:\n\n ⏰ **10:00 AM | 01:00 PM | 04:00 PM | 07:00 PM | 10:00 PM**")
            
            with st.form("mail_config_form"):
                sender = st.text_input("Sender Gmail Address", st.session_state["email_config"]["sender"])
                password = st.text_input("Sender Google App Password (16-Digit)", st.session_state["email_config"]["password"], type="password")
                receiver = st.text_input("Receiver Backup Email Address", st.session_state["email_config"]["receiver"])
                
                if st.form_submit_button("Save Email Engine Settings"):
                    st.session_state["email_config"] = {"sender": sender, "password": password, "receiver": receiver}
                    st.success("✅ Email Configuration Settings Saved Locally!")
            
            if st.button("🚀 Test Send Instant Day-wise Data Now"):
                with st.spinner("Sending Day-wise Excel file attachment..."):
                    if send_daywise_backup_email():
                        st.success("🎉 Email Sent Successfully with today's day-wise data filter!")
                    else:
                        st.error("❌ Email transmission failed. Please check credentials or Google App security keys.")

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
                st.success("🎯 Targets Saved Successfully!")

        # --- DATA ENTRY LOGIC (ADMIN) ---
        elif user["role"] == "admin" and nav_choice == "📝 Data Entry":
            st.subheader("📝 Quick Data Entry Panel")
            with st.form("entry_form_admin", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    date_input = st.date_input("Select Date", datetime.now())
                    challan_no = st.text_input("Enter Challan Number Locked", placeholder="e.g. CH-9921")
                    design_no = st.text_input("Enter Design Number", placeholder="e.g. D-401")
                    party_name = st.selectbox("Select Party Name Account", st.session_state["party_options"])
                with col2:
                    item_type = st.selectbox("Select Fabric / Item Type", st.session_state["item_options"])
                    total_pcs = st.number_input("Total Pieces Count", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Quality Pieces", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Damage Pieces", min_value=0, step=1)
                if st.form_submit_button("SAVE ADMIN ENTRY"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Total mismatch! (Fresh + Seconds) must be equal to Total Pcs.")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully!")

        # --- DASHBOARD LOGIC ---
        else:
            st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Live Analytics Dashboard</h2>", unsafe_allow_html=True)
            
            # --- 📅 NO 1: DATE SELECTION & PRODUCTION BREAKDOWN ---
            st.markdown("### 📅 Filter Dashboard by Date")
            selected_filter_date = st.date_input("Choose Date to View Summary & Graph", datetime.now())
            formatted_filter_date = selected_filter_date.strftime("%d-%m-%Y")

            if excel_loaded and not df.empty:
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                df['Date'] = df['Date'].astype(str).str.strip()
                filtered_df_by_date = df[df["Date"] == formatted_filter_date]
                item_groups = filtered_df_by_date.groupby('Item Type')['Total Pcs'].sum().to_dict()
            else:
                filtered_df_by_date = pd.DataFrame()
                item_groups = {}

            items_list = [it.upper().strip() for it in st.session_state["item_options"]]

            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.markdown(f"""
                <div style="background-color: #FFFF00; color: #000000; text-align: center; font-weight: bold; padding: 12px; font-size: 20px; border: 2px solid #2c3e50; border-bottom: none; border-radius: 6px 6px 0px 0px;">
                    🏭 ITEM PIECES SUMMARY
                </div>
                <div style="background-color: #1F4E79; color: #FFFFFF; padding: 10px 15px; font-weight: bold; font-size: 15px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 2px solid #2c3e50; display: flex; justify-content: space-between;">
                    <span>📅 SELECTED DATE</span><span>{formatted_filter_date}</span>
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
                st.markdown(f"### 🍩 Share Matrix Graph ({formatted_filter_date})")
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
                    st.info(f"Is date ({formatted_filter_date}) me graph ke liye koi data entry nahi mili.")
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("---")

            # --- 🎯 NO 2: TARGET STATUS TABLE ---
            st.markdown("### 🎯 Supervisor Live Target Tracker (Auto-Less Status)")
            target_summary_data = []
            for u_id, u_info in st.session_state["users"].items():
                if u_info["role"] == "supervisor":
                    s_name = u_info["name"]
                    allocated_tgt = st.session_state["supervisor_targets"].get(s_name, 0)
                    
                    if excel_loaded and not df.empty:
                        done_pcs = df[df["Supervisor"] == s_name]["Total Pcs"].sum()
                    else:
                        done_pcs = 0
                        
                    remaining_tgt = allocated_tgt - done_pcs
                    if remaining_tgt < 0: remaining_tgt = 0
                    
                    status_txt = "✅ Done" if remaining_tgt == 0 and allocated_tgt > 0 else "⏳ Pending"
                    
                    target_summary_data.append({
                        "Supervisor Name": s_name,
                        "Assigned Target": allocated_tgt,
                        "Completed (Pcs)": done_pcs,
                        "Remaining (Baki)": remaining_tgt,
                        "Status": status_txt
                    })
            
            target_summary_df = pd.DataFrame(target_summary_data)
            st.dataframe(target_summary_df, hide_index=True, use_container_width=True)
            st.markdown("---")

            # --- 👥 SUPERVISOR PERSONAL VIEW / ADMIN VIEW PANELS ---
            if user["role"] == "supervisor":
                can_entry = user.get("p_entry", True)
                can_view_logs = user.get("p_view", True)
                can_edit_logs = user.get("p_edit", False)

                if can_entry:
                    st.subheader("📝 Production Entry Input Panel")
                    with st.form("entry_form_sup", clear_on_submit=True):
                        col1, col2 = st.columns(2)
                        with col1:
                            date_input = st.date_input("Select Production Date", datetime.now())
                            challan_no = st.text_input("Enter Challan Number", placeholder="e.g. CH-205")
                            design_no = st.text_input("Enter Design Number/Code", placeholder="e.g. DS-902")
                            party_name = st.selectbox("Choose Party Name", st.session_state["party_options"])
                        with col2:
                            item_type = st.selectbox("Choose Item Type", st.session_state["item_options"])
                            total_pcs = st.number_input("Enter Total Pieces", min_value=0, step=1)
                            fresh_pcs = st.number_input("Enter Fresh Pieces", min_value=0, step=1)
                            seconds_pcs = st.number_input("Enter Seconds Pieces", min_value=0, step=1)
                        
                        if st.form_submit_button("SAVE PRODUCTION ENTRY"):
                            if fresh_pcs + seconds_pcs != total_pcs:
                                st.error("❌ Calculation mismatch! Fresh + Seconds must equal Total Pieces.")
                            else:
                                wb = openpyxl.load_workbook(EXCEL_FILE)
                                ws = wb["Supervisor Entry"]
                                ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                                wb.save(EXCEL_FILE)
                                st.success("🎉 Entry Saved!")
                                st.rerun()

                if can_view_logs:
                    st.markdown("---")
                    st.subheader("📋 Your Done Production Entries")
                    sup_df = df[df["Supervisor"] == user["name"]] if excel_loaded else pd.DataFrame()
                    
                    if not sup_df.empty:
                        if can_edit_logs:
                            st.data_editor(sup_df, hide_index=True, use_container_width=True)
                        else:
                            st.dataframe(sup_df, hide_index=True, use_container_width=True)
            else:
                st.subheader("📋 Production Master Logs (Full Control)")
                st.dataframe(df, hide_index=True, use_container_width=True, height=250)

            # --- CONFIGURATIONS DESK FOR ADMIN ---
            if user["role"] == "admin":
                st.markdown("---")
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
                                st.success("Updated!")
                                st.rerun()
                    with col3:
                        st.markdown("**❌ Remove Party**")
                        party_to_remove = st.selectbox("Select Party to Delete", st.session_state["party_options"], key="rem_p")
                        if st.button("Delete Party From System", type="primary"):
                            if party_to_remove in st.session_state["party_options"]:
                                st.session_state["party_options"].remove(party_to_remove)
                                st.rerun()

                with t2:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown("**➕ Add New Item Type**")
                        new_item = st.text_input("Enter Item Name to Add", placeholder="e.g. SILK SAREE").strip()
                        if st.button("Save New Item"):
                            if new_item and new_item.upper() not in st.session_state["item_options"]:
                                st.session_state["item_options"].append(new_item.upper())
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
                                st.success("Updated!")
                                st.rerun()
                    with col3:
                        st.markdown("**❌ Remove Item Type**")
                        item_to_remove = st.selectbox("Select Item to Delete from System", st.session_state["item_options"], key="rem_i")
                        if st.button("Delete Item Category", type="primary"):
                            if item_to_remove in st.session_state["item_options"]:
                                st.session_state["item_options"].remove(item_to_remove)
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
                            cb_edit = st.checkbox("Allow Edit/Delete Logged Records (Admin Rights)", value=False)
                            
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
                                        st.success(f"Supervisor '{add_name}' Created!")
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
                            edit_cb_edit = st.checkbox("Allow Edit/Delete Logged Records (Admin Rights)", value=current_sup_data.get("p_edit", False), key="ed_d")
                            
                            if st.button("Update Supervisor Account"):
                                if edit_name and edit_pass:
                                    st.session_state["users"][selected_sup]["name"] = edit_name
                                    st.session_state["users"][selected_sup]["password"] = edit_pass
                                    st.session_state["users"][selected_sup]["p_entry"] = edit_cb_entry
                                    st.session_state["users"][selected_sup]["p_view"] = edit_cb_view
                                    st.session_state["users"][selected_sup]["p_edit"] = edit_cb_edit
                                    st.success("Account & Checkbox Permissions updated!")
                                    st.rerun()
                            
                    with col_u3:
                        st.markdown("**❌ Remove Supervisor**")
                        if sups_only:
                            sup_to_remove = st.selectbox("Select Supervisor ID to Delete", sups_only, key="sel_sup_rem")
                            if st.button("Delete Supervisor Account", type="primary"):
                                del st.session_state["users"][sup_to_remove]
                                st.warning("Account deleted from database!")
                                st.rerun()
