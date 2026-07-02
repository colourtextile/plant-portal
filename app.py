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

# --- 📧 EMAIL CONFIGURATION (Apna Details Yahan Bharein) ---
EMAIL_SENDER = "aapka_gmail@gmail.com"        # Bhejne wale ka email
EMAIL_PASSWORD = "xxxx xxxx xxxx xxxx"       # Gmail ka 16-digit App Password 
EMAIL_RECEIVER = "aapka_mail@gmail.com"      # Jiss par backup mangwana hai

st.set_page_config(page_title="Plant Production Portal", layout="wide")

# --- 🛡️ AUTOMATIC EMAIL BACKUP FUNCTION ---
def send_excel_backup_email():
    if not os.path.exists(EXCEL_FILE):
        return
    
    try:
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Create Message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECEIVER
        msg['Subject'] = f"🔄 AUTOMATIC BACKUP: Plant Production Data ({current_time})"
        
        # Attach File
        with open(EXCEL_FILE, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= Backup_{datetime.now().strftime('%d_%m_%Y_%H_%M')}.xlsx",
            )
            msg.attach(part)
            
        # Connect and Send via Gmail SMTP
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"✅ Backup successfully sent at {current_time}")
    except Exception as e:
        print(f"❌ Backup failed: {str(e)}")

# --- 🕒 SCHEDULER: DIN ME 5 BAAR AUTOMATIC RUN ---
# Yeh ensure karega ki scheduler sirf ek baar initialize ho jab server start ho
if "scheduler_started" not in st.session_state:
    scheduler = BackgroundScheduler(timezone="Asia/Kolkata")
    
    # Din me 5 baar ka time set kiya hai (Aap apne hisab se hours badal sakte hain)
    # 1. Subah 10:00 baje | 2. Dopahar 1:00 baje | 3. Shaam 4:00 baje | 4. Raat 7:00 baje | 5. Raat 10:00 baje
    backup_hours = [10, 13, 16, 19, 22] 
    
    for hr in backup_hours:
        scheduler.add_job(send_excel_backup_email, 'cron', hour=hr, minute=0)
        
    scheduler.start()
    st.session_state["scheduler_started"] = True
    atexit.register(lambda: scheduler.shutdown())

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
    
    # --- SIDEBAR AREA ---
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
                excel_loaded = True
        except Exception as e:
            pass

    # REPORT GENERATOR IN SIDEBAR
    if user["role"] == "admin" and excel_loaded and not df.empty:
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📊 Download Reports")
        
        all_parties = ["All Parties"] + sorted(list(df["Party Name"].dropna().unique()))
        all_designs = ["All Designs"] + sorted(list(df["Design No"].dropna().astype(str).unique()))
        
        filter_party = st.sidebar.selectbox("Filter by Party", all_parties)
        filter_design = st.sidebar.selectbox("Filter by Design", all_designs)
        
        filtered_df = df.copy()
        if filter_party != "All Parties":
            filtered_df = filtered_df[filtered_df["Party Name"] == filter_party]
        if filter_design != "All Designs":
            filtered_df = filtered_df[filtered_df["Design No"] == filter_design]
            
        st.sidebar.markdown(f"Found Rows: `{len(filtered_df)}`")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, sheet_name='Filtered Report', index=False)
        
        st.sidebar.download_button(
            label="📥 Download Excel Report",
            data=buffer.getvalue(),
            file_name=f"Report_{filter_party}_{filter_design}_{datetime.now().strftime('%d-%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Manual Backup Button for Admin
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

    # --- MAIN PORTAL SCREEN ---
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
                    challan_no = st.text_input("Challan Number", placeholder="e.g., CH-501")
                    design_no = st.text_input("Design Number", placeholder="e.g., DS-1001")
                    party_name = st.selectbox("Select Party Name", st.session_state["party_options"])
                with col2:
                    item_type = st.selectbox("Select Item Type", st.session_state["item_options"])
                    total_pcs = st.number_input("Total Pieces Produced", min_value=0, step=1)
                    fresh_pcs = st.number_input("Fresh Pieces (Grade A)", min_value=0, step=1)
                    seconds_pcs = st.number_input("Seconds Pieces (Grade B)", min_value=0, step=1)
                
                if st.form_submit_button("SAVE ENTRY"):
                    if fresh_pcs + seconds_pcs != total_pcs:
                        st.error("❌ Galti: Total Pcs ka jod sahi nahi hai!")
                    else:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully with Challan!")
                        st.rerun()

        # --- ROLE 2: ADMIN ALL-IN-ONE DESKBOARD ---
        elif user["role"] == "admin":
            st.title("👑 Admin Master Deskboard")
            
            with st.expander("📝 Quick Data Entry Form", expanded=False):
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
                            st.error("❌ Total Pcs galat hain!")
                        else:
                            wb = openpyxl.load_workbook(EXCEL_FILE)
                            ws = wb["Supervisor Entry"]
                            ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                            wb.save(EXCEL_FILE)
                            st.success("🎉 Entry Saved Successfully!")
                            st.rerun()

            st.markdown("---")
            
            if excel_loaded and not df.empty:
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                item_groups = df.groupby('Item Type')['Total Pcs'].sum().to_dict()
            else:
                item_groups = {}

            items_list = [it.upper().strip() for it in st.session_state["item_options"]]
            current_date = datetime.now().strftime("%d-%m-%y")

            # --- SCREEN COLUMNS FOR GRAPH AND BEAUTIFUL TABLE ---
            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.subheader("📊 Live Summary Table")
                
                html_table = f"""
                <div style="border: 2px solid #000000; border-radius: 4px; overflow: hidden; font-family: Arial, sans-serif;">
                    <div style="background-color: #FFFF00; color: #000000; text-align: center; font-weight: bold; padding: 10px; font-size: 18px; border-bottom: 2px solid #000000;">
                        PRODUCTION
                    </div>
                    <div style="background-color: #4F81BD; color: #FFFFFF; display: flex; justify-content: space-between; font-weight: bold; padding: 8px 15px; font-size: 15px; border-bottom: 1px solid #000000;">
                        <span>DATE</span>
                        <span>{current_date}</span>
                    </div>
                """
                total_sum = 0
                for it in items_list:
                    val = item_groups.get(it, 0)
                    total_sum += val
                    html_table += f"""
                    <div style="background-color: #FFFFFF; color: #000000; display: flex; justify-content: space-between; padding: 8px 15px; font-size: 14px; border-bottom: 1px solid #E0E0E0;">
                        <span>{it}</span>
                        <span style="font-weight: bold;">{val:,}</span>
                    </div>
                    """
                
                html_table += f"""
                    <div style="background-color: #F2F2F2; color: #000000; display: flex; justify-content: space-between; font-weight: bold; padding: 10px 15px; font-size: 16px; border-top: 2px solid #000000;">
                        <span>TOTAL PCS</span>
                        <span style="color: #000000; text-decoration: underline;">{total_sum:,}</span>
                    </div>
                </div>
                """
                st.markdown(html_table, unsafe_allow_html=True)
                
            with col_right:
                st.subheader("🍩 Item Wise Percentage Graph")
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
                        'mark': {'type': 'arc', 'innerRadius': 50, 'tooltip': True},
                        'encoding': {
                            'theta': {'field': 'Pieces', 'type': 'quantitative'},
                            'color': {'field': 'Items', 'type': 'nominal', 'scale': {'scheme': 'category20'}},
                        },
                        'view': {'stroke': None}
                    }, use_container_width=True)
                else:
                    st.info("Graph dekhne ke liye data entry kijiye!")

            # --- PART D: RAW DATA & ADVANCED CONTROLS ---
            st.markdown("---")
            st.subheader("📋 Raw Excel Logs")
            st.dataframe(df, hide_index=True, use_container_width=True, height=220)

            st.markdown("---")
            st.subheader("⚙️ System Control Center (Edit / Remove Panel)")
            t1, t2, t3 = st.tabs(["🏢 Manage Parties", "📦 Manage Items", "👥 Supervisors Accounts"])
            
            # TAB 1: PARTY MANAGEMENT
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

            # TAB 2: ITEM MANAGEMENT
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

            # TAB 3: SUPERVISOR ACCOUNTS
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
                                    st.error("❌ Yeh ID pehle se bani hui hai!")
                            else:
                                st.error("❌ Saari details bharna zaroori hai.")
                                
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
                                st.success(f"ID `{selected_sup}` updated successfully!")
                                st.rerun()
                    else:
                        st.info("Koi supervisor account nahi mila.")
                        
                with col_u3:
                    st.markdown("**❌ Remove Supervisor**")
                    if sups_only:
                        sup_to_remove = st.selectbox("Select ID to Delete", sups_only, key="sel_sup_rem")
                        if st.button("Delete Supervisor Account", type="primary"):
                            del st.session_state["users"][sup_to_remove]
                            st.warning(f"ID `{sup_to_remove}` deleted hamesha ke liye!")
                            st.rerun()
                    else:
                        st.info("No accounts to delete.")
                    
                    st.markdown("---")
                    st.markdown("**👥 Current Active List:**")
                    for u_id, u_info in st.session_state["users"].items():
                        st.write(f"- `{u_id}` : {u_info['name']}")
