import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from apscheduler.schedulers.background import BackgroundScheduler
import atexit

EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"

# --- 📧 EMAIL CONFIGURATION ---
EMAIL_SENDER = "aapka_gmail@gmail.com"
EMAIL_PASSWORD = "xxxx xxxx xxxx xxxx"       
EMAIL_RECEIVER = "aapka_mail@gmail.com"      

st.set_page_config(page_title="Colour Textile Portal", layout="wide")

# --- 💅 CLEAN & PREMIUM LOGIN OVERRIDES (NO MORE CARDS / BOXES) ---
st.markdown("""
<style>
    .reportview-container {
        background: #f8f9fa;
    }
    
    /* Dynamic Multi-color Gradient Textile Title */
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

    /* Input text styling for high contrast readability against dark background */
    div[data-testid="stForm"] label {
        color: #FFFFFF !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        letter-spacing: 0.5px;
    }
    
    /* Giving the form containers a clean frameless alignment */
    div[data-testid="stForm"] {
        border: none !important;
        padding: 0 !important;
        background: transparent !important;
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

# --- 🛡️ AUTOMATIC EMAIL BACKUP FUNCTION ---
def send_excel_backup_email():
    if not os.path.exists(EXCEL_FILE):
        return
    try:
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"🔄 AUTOMATIC BACKUP: Plant Production Data ({current_time})"
        
        with open(EXCEL_FILE, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= Backup_{datetime.now().strftime('%d_%m_%Y_%H_%M')}.xlsx",
            )
            msg.attach(part)
            
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
    except:
        pass

# --- 🕒 SCHEDULER ---
if "scheduler_started" not in st.session_state:
    scheduler = BackgroundScheduler(timezone="Asia/Kolkata")
    backup_hours = [10, 13, 16, 19, 22] 
    for hr in backup_hours:
        scheduler.add_job(send_excel_backup_email, 'cron', hour=hr, minute=0)
    scheduler.start()
    st.session_state["scheduler_started"] = True
    atexit.register(lambda: scheduler.shutdown())

# INITIALIZING DATA STRUCTURES
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

# --- 🔒 MAIN APPLICATION LOGIN SCREEN ---
if not st.session_state["logged_in"]:
    # Restoring the perfect premium room/chair background wallpaper from image_5f303d.jpg
    st.markdown(f"""
    <style>
        .stApp {{
            background-image: linear-gradient(rgba(0, 0, 0, 0.55), rgba(0, 0, 0, 0.55)), url("https://images.unsplash.com/photo-1617627143750-d86bc21e42bb?q=80&w=1920&auto=format&fit=crop");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
    </style>
    """, unsafe_allow_html=True)
    
    # Large frameless dynamic title header
    st.markdown('<h1 class="brand-title">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #BDC3C7; font-size: 13px; margin-top: -5px; margin-bottom: 40px; font-weight: 500; letter-spacing: 1px;">SAREE WEAVING & PRODUCTION ERP</p>', unsafe_allow_html=True)
    
    # Beautiful large layout container for login inputs without any restrictive white card boxes
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    
    with col_l2:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username / ID ID", placeholder="Enter your registered ID")
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
    
    st.sidebar.markdown(f"### 👤 Welcome, **{user['name']}**")
    st.sidebar.markdown(f"📋 Role: `{user['role'].upper()}`")
    
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

    # --- ADVANCED REPORTS ---
    if user["role"] == "admin" and excel_loaded and not df.empty:
        st.sidebar.markdown("---")
        with st.sidebar.expander("📊 Advanced Reports Center", expanded=False):
            all_parties = ["All Parties"] + sorted(list(df["Party Name"].dropna().unique()))
            all_designs = ["All Designs"] + sorted(list(df["Design No"].dropna().astype(str).unique()))
            all_items = ["All Items"] + sorted(list(df["Item Type"].dropna().unique()))
            all_supervisors = ["All Supervisors"] + sorted(list(df["Supervisor"].dropna().unique()))
            all_challans = ["All Challans"] + sorted(list(df["Challan No"].dropna().astype(str).unique()))
            
            filter_party = st.selectbox("1. Filter by Party", all_parties)
            filter_design = st.selectbox("2. Filter by Design No", all_designs)
            filter_item = st.selectbox("3. Filter by Item Type", all_items)
            filter_supervisor = st.selectbox("4. Filter by Supervisor", all_supervisors)
            filter_challan = st.selectbox("5. Filter by Challan No", all_challans)
            
            filtered_df = df.copy()
            if filter_party != "All Parties":
                filtered_df = filtered_df[filtered_df["Party Name"] == filter_party]
            if filter_design != "All Designs":
                filtered_df = filtered_df[filtered_df["Design No"].astype(str) == filter_design]
            if filter_item != "All Items":
                filtered_df = filtered_df[filtered_df["Item Type"] == filter_item]
            if filter_supervisor != "All Supervisors":
                filtered_df = filtered_df[filtered_df["Supervisor"] == filter_supervisor]
            if filter_challan != "All Challans":
                filtered_df = filtered_df[filtered_df["Challan No"].astype(str) == filter_challan]
                
            tot_filtered_pcs = filtered_df["Total Pcs"].sum() if not filtered_df.empty else 0
            st.markdown(f"""
            <div style="background-color: #f1f3f4; padding: 8px; border-radius: 4px; margin-bottom: 10px;">
                <p style="margin:0; font-size:12px; color:#5f6368;">Found Rows: <b>{len(filtered_df)}</b></p>
                <p style="margin:0; font-size:12px; color:#5f6368;">Total Selected Pcs: <b>{tot_filtered_pcs:,}</b></p>
            </div>
            """, unsafe_allow_html=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='Filtered Report', index=False)
            
            st.download_button(
                label="📥 Download Filtered Excel",
                data=buffer.getvalue(),
                file_name=f"Advanced_Report_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    if user["role"] == "admin":
        st.sidebar.markdown("---")
        if st.sidebar.button("📧 Send Backup to Mail Now"):
            send_excel_backup_email()
            st.sidebar.success("✅ Backup mail sent!")

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()

    # --- MAIN WORKSPACE SYSTEM CONTENT ---
    if not os.path.exists(EXCEL_FILE):
        st.error(f"Excel file '{EXCEL_FILE}' nahi mili!")
    else:
        if user["role"] == "supervisor":
            st.title("📝 Colour Textile Production Entry Panel")
            with st.form("entry_form_sup", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    date_input = st.date_input("Date", datetime.now())
                    challan_no = st.text_input("Challan Number", placeholder="e.g., CH-501")
                    design_no = st.text_input("Design Number", placeholder="e.g., DS-1001")
                    party_name = st.selectbox("Select Party Name", st.session_state["party_options"])
                with col2:
                    item_type = st.selectbox("Select Item Type", st.session_state["item_options"])
                    total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1)
                
                if st.form_submit_button("SAVE PRODUCTION ENTRY"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Galti: Total Pcs ka jod sahi nahi hai!")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Data Entry Saved Successfully!")
                        st.rerun()

        elif user["role"] == "admin":
            st.markdown("<h1 style='color: #1F4E79; font-weight: bold;'>🏭 Colour Textile Master Control Room</h1>", unsafe_allow_html=True)
            
            with st.expander("📝 Quick Data Entry Form (Admin View)", expanded=False):
                with st.form("entry_form_admin", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    with col1:
                        date_input = st.date_input("Date", datetime.now(), key="adm_date")
                        challan_no = st.text_input("Challan Number", placeholder="e.g., CH-501", key="adm_chal")
                        design_no = st.text_input("Design Number", placeholder="e.g., DS-1001", key="adm_des")
                        party_name = st.selectbox("Select Party Name", st.session_state["party_options"], key="adm_party")
                    with col2:
                        item_type = st.selectbox("Select Item Type", st.session_state["item_options"], key="adm_item")
                        total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1, key="adm_tot")
                        fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1, key="adm_fr")
                        seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1, key="adm_sec")
                    if st.form_submit_button("SAVE ADMIN ENTRY"):
                        if fresh_pcs + seconds_pcs != total_pcs:
                            st.error("❌ Total Pcs calculation mismatch!")
                        else:
                            wb = openpyxl.load_workbook(EXCEL_FILE)
                            ws = wb["Supervisor Entry"]
                            ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                            wb.save(EXCEL_FILE)
                            st.success("🎉 Entry Saved!")
                            st.rerun()

            st.markdown("---")
            
            if excel_loaded and not df.empty:
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                item_groups = df.groupby('Item Type')['Total Pcs'].sum().to_dict()
            else:
                item_groups = {}

            items_list = [it.upper().strip() for it in st.session_state["item_options"]]
            current_date = datetime.now().strftime("%d-%m-%y")

            # --- METRICS BOARDS ---
            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.subheader("📊 Live Production Summary")
                
                st.markdown(f"""
                <div style="background-color: #FFFF00; color: #000000; text-align: center; font-weight: bold; padding: 12px; font-size: 20px; border: 2px solid #2c3e50; border-bottom: none; border-radius: 6px 6px 0px 0px; font-family: sans-serif;">
                    🏭 PRODUCTION REPORT
                </div>
                <div style="background-color: #1F4E79; color: #FFFFFF; padding: 10px 15px; font-weight: bold; font-size: 15px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 2px solid #2c3e50; display: flex; justify-content: space-between;">
                    <span>📅 SYSTEM DATE</span>
                    <span>{current_date}</span>
                </div>
                """, unsafe_allow_html=True)
                
                total_sum = 0
                for it in items_list:
                    val = item_groups.get(it, 0)
                    total_sum += val
                    st.markdown(f"""
                    <div style="background-color: #FFFFFF; color: #333333; padding: 10px 15px; font-size: 14px; border-left: 2px solid #2c3e50; border-right: 2px solid #2c3e50; border-bottom: 1px solid #EAEAEA; display: flex; justify-content: space-between;">
                        <span style="font-weight: 500;">{it}</span>
                        <span style="font-weight: bold; color: #2c3e50;">{val:,} Pcs</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown(f"""
                <div style="background-color: #27AE60; color: #FFFFFF; padding: 12px 15px; font-weight: bold; font-size: 18px; border: 2px solid #2c3e50; border-radius: 0px 0px 6px 6px; display: flex; justify-content: space-between; box-shadow: 0px 4px 10px rgba(39, 174, 96, 0.25);">
                    <span>📊 TOTAL PRODUCED</span>
                    <span style="text-decoration: underline; font-size: 19px; letter-spacing: 0.5px;">{total_sum:,} Pcs</span>
                </div>
                """, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
            with col_right:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.subheader("🍩 Share Share Matrix Chart")
                chart_labels = []
                chart_values = []
                for it in items_list:
                    val = item_groups.get(it, 0)
                    if val > 0:
                        chart_labels.append(it)
                        chart_values.append(val)
                
                if chart_values:
                    chart_df = pd.DataFrame({"Items": chart_labels, "Pieces": chart_values})
                    st.vega_lite_chart(chart_df, {
                        'mark': {'type': 'arc', 'innerRadius': 55, 'tooltip': True},
                        'encoding': {
                            'theta': {'field': 'Pieces', 'type': 'quantitative'},
                            'color': {'field': 'Items', 'type': 'nominal', 'scale': {'scheme': 'tablet10'}},
                        },
                        'view': {'stroke': None}
                    }, use_container_width=True)
                else:
                    st.info("No data recorded yet.")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- VIEW MASTER DATAFRAME LOGS ---
            st.markdown("---")
            st.subheader("📋 Production Master Logs (Date & Challan Locked)")
            st.dataframe(df, hide_index=True, use_container_width=True, height=250)

            # --- TABS FOR DROPDOWNS MANAGEMENT ---
            st.markdown("---")
            st.subheader("⚙️ Factory Configurations Desk")
            t1, t2, t3 = st.tabs(["🏢 Manage Parties", "📦 Manage Items", "👥 Supervisors Accounts"])
            
            with t1:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**➕ Add New Party**")
                    new_party = st.text_input("Party Name to Add").strip()
                    if st.button("Save New Party"):
                        if new_party and new_party not in st.session_state["party_options"]:
                            st.session_state["party_options"].append(new_party)
                            st.success(f"Added: {new_party}")
                            st.rerun()
                with col2:
                    st.markdown("**✏️ Edit Party Name**")
                    party_to_edit = st.selectbox("Select Party to Edit", st.session_state["party_options"], key="edt_p")
                    edited_party_name = st.text_input("Enter New Name for Party", value=party_to_edit)
                    if st.button("Update Party Name"):
                        if edited_party_name and party_to_edit:
                            idx = st.session_state["party_options"].index(party_to_edit)
                            st.session_state["party_options"][idx] = edited_party_name
                            st.success("Party name updated successfully!")
                            st.rerun()
                with col3:
                    st.markdown("**❌ Remove Party**")
                    party_to_remove = st.selectbox("Select Party to Remove", st.session_state["party_options"], key="rem_p")
                    if st.button("Delete Party", type="primary"):
                        if party_to_remove in st.session_state["party_options"]:
                            st.session_state["party_options"].remove(party_to_remove)
                            st.warning(f"Removed: {party_to_remove}")
                            st.rerun()

            with t2:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**➕ Add New Item Type**")
                    new_item = st.text_input("Item Name to Add").strip()
                    if st.button("Save New Item"):
                        if new_item and new_item.upper() not in st.session_state["item_options"]:
                            st.session_state["item_options"].append(new_item.upper())
                            st.success(f"Added: {new_item.upper()}")
                            st.rerun()
                with col2:
                    st.markdown("**✏️ Edit Item Name**")
                    item_to_edit = st.selectbox("Select Item to Edit", st.session_state["item_options"], key="edt_i")
                    edited_item_name = st.text_input("Enter New Name for Item", value=item_to_edit)
                    if st.button("Update Item Name"):
                        if edited_item_name and item_to_edit:
                            idx = st.session_state["item_options"].index(item_to_edit)
                            st.session_state["item_options"][idx] = edited_item_name.upper()
                            st.success("Item name updated successfully!")
                            st.rerun()
                with col3:
                    st.markdown("**❌ Remove Item Type**")
                    item_to_remove = st.selectbox("Select Item to Remove", st.session_state["item_options"], key="rem_i")
                    if st.button("Delete Item", type="primary"):
                        if item_to_remove in st.session_state["item_options"]:
                            st.session_state["item_options"].remove(item_to_remove)
                            st.warning(f"Removed: {item_to_remove}")
                            st.rerun()

            with t3:
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    st.markdown("**➕ Add New Supervisor**")
                    with st.form("add_user_form", clear_on_submit=True):
                        add_id = st.text_input("New Username / ID").strip()
                        add_pass = st.text_input("Set Password").strip()
                        add_name = st.text_input("Supervisor Real Name").strip()
                        if st.form_submit_button("Create Account"):
                            if add_id and add_pass and add_name:
                                if add_id not in st.session_state["users"]:
                                    st.session_state["users"][add_id] = {"password": add_pass, "name": add_name, "role": "supervisor"}
                                    st.success(f"Supervisor '{add_name}' Created!")
                                    st.rerun()
                                else:
                                    st.error("❌ ID exists!")
                            else:
                                st.error("❌ All fields required.")
                                
                with col_u2:
                    st.markdown("**✏️ Edit Supervisor Info**")
                    all_users = list(st.session_state["users"].keys())
                    sups_only = [u for u in all_users if st.session_state["users"][u]["role"] == "supervisor"]
                    
                    if sups_only:
                        selected_sup = st.selectbox("Select ID to Edit", sups_only, key="sel_sup_edt")
                        current_sup_data = st.session_state["users"][selected_sup]
                        edit_name = st.text_input("Change Full Name", value=current_sup_data["name"])
                        edit_pass = st.text_input("Change Password", value=current_sup_data["password"])
                        
                        if st.button("Update Supervisor Account"):
                            if edit_name and edit_pass:
                                st.session_state["users"][selected_sup]["name"] = edit_name
                                st.session_state["users"][selected_sup]["password"] = edit_pass
                                st.success("Account updated!")
                                st.rerun()
                    else:
                        st.info("No supervisors setup.")
                        
                with col_u3:
                    st.markdown("**❌ Remove Supervisor**")
                    if sups_only:
                        sup_to_remove = st.selectbox("Select ID to Delete", sups_only, key="sel_sup_rem")
                        if st.button("Delete Supervisor Account", type="primary"):
                            del st.session_state["users"][sup_to_remove]
                            st.warning("Account deleted!")
                            st.rerun()
                    else:
                        st.info("No accounts to delete.")
