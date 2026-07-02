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

# ==========================================
# 1. PAGE CONFIGURATION & CSS
# ==========================================
EXCEL_FILE = "Final_Plant_System_With_All_Dropdowns.xlsx"
st.set_page_config(page_title="Colour Textile Portal", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #0f172a !important; padding: 20px 10px; }
    .premium-sidebar-header { background: linear-gradient(135deg, #1e293b, #334155); padding: 20px; border-radius: 12px; margin-bottom: 25px; border-left: 5px solid #38bdf8; text-align: center; }
    .user-chip { display: inline-block; background: #0ea5e9; color: white; padding: 4px 14px; border-radius: 20px; font-size: 11px; font-weight: 700; margin-top: 10px; text-transform: uppercase; }
    .dashboard-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #eef2f5; }
    .stButton>button { border-radius: 6px; font-weight: bold; }
    .metric-box { background-color: #f1f5f9; padding: 15px; border-radius: 8px; text-align: center; }
    .metric-title { font-size: 14px; color: #64748b; }
    .metric-value { font-size: 24px; font-weight: bold; color: #0f172a; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SESSION STATE INITIALIZATION
# ==========================================
if "supervisor_targets" not in st.session_state: 
    st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}
if "users" not in st.session_state:
    st.session_state["users"] = {
        "admin": {"password": "plant123", "name": "Admin Master", "role": "admin", "p_entry": True, "p_view": True, "p_edit": True},
        "ramesh01": {"password": "ramesh@123", "name": "Ramesh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False},
        "suresh01": {"password": "suresh@123", "name": "Suresh", "role": "supervisor", "p_entry": True, "p_view": True, "p_edit": False}
    }
if "party_options" not in st.session_state: 
    st.session_state["party_options"] = ["Krishna Textiles", "Balaji Fabrics", "Radhe Shyam Corp"]
if "item_options" not in st.session_state: 
    st.session_state["item_options"] = ["SAREE", "SUIT", "DUPATTA", "ONLY TOP"]
if "logged_in" not in st.session_state: 
    st.session_state["logged_in"] = False
if "email_config" not in st.session_state:
    st.session_state["email_config"] = {"sender": "", "password": "", "receiver": ""}

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def load_excel_data():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
            # Safety check for Date column to prevent KeyError
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                df['Month_Year'] = df['Date'].dt.strftime('%B %Y')
            return df, True
        except Exception as e:
            st.sidebar.error(f"Error loading Excel: {e}")
            return pd.DataFrame(), False
    else:
        # Create blank template if file doesn't exist
        df = pd.DataFrame(columns=["Date", "Design No", "Party Name", "Item Type", "Total Pcs", "Fresh Pcs", "Seconds Pcs", "Supervisor", "Challan No"])
        df.to_excel(EXCEL_FILE, sheet_name="Supervisor Entry", index=False)
        return df, True

def send_email_backup(sender, password, receiver, file_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = f"Production Backup - {datetime.now().strftime('%d %b %Y')}"
        
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file_path)}")
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        return True, "Email sent successfully!"
    except Exception as e:
        return False, str(e)

# ==========================================
# 4. LOGIN SYSTEM
# ==========================================
if not st.session_state["logged_in"]:
    st.markdown('<br><br><h1 style="text-align:center; color: #0f172a;">🏢 COLOUR TEXTILE PORTAL</h1><p style="text-align:center; color: #64748b;">Secure Production Management System</p><br>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            st.markdown("### 🔐 Login")
            user_in = st.text_input("Username")
            pass_in = st.text_input("Password", type="password")
            submit_login = st.form_submit_button("LOGIN TO PORTAL", use_container_width=True)
            
            if submit_login:
                if user_in in st.session_state["users"] and st.session_state["users"][user_in]["password"] == pass_in:
                    st.session_state["logged_in"] = True
                    st.session_state["current_user"] = st.session_state["users"][user_in]
                    st.rerun()
                else:
                    st.error("❌ Invalid Username or Password")
    st.stop()

# ==========================================
# 5. SIDEBAR NAVIGATION
# ==========================================
user = st.session_state["current_user"]
df, excel_loaded = load_excel_data()

st.sidebar.markdown(f"""
<div class="premium-sidebar-header">
    <div style="color: #ffffff; font-size: 20px; font-weight: 900;">COLOUR TEXTILE</div>
    <div class="user-chip">{user['role'].upper()}</div>
    <div style="color: #cbd5e1; font-size: 13px; margin-top: 15px;">👤 {user['name']}</div>
</div>
""", unsafe_allow_html=True)

nav_options = ["📊 Dashboard", "📝 Production Entry", "📋 Edit & Delete Records"]
if user["role"] == "admin":
    nav_options.extend(["🎯 Target Settings", "📧 Email Auto-Backup", "⚙️ Factory Config"])

choice = st.sidebar.radio("🧭 Navigation Menu", nav_options)
st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout", use_container_width=True):
    st.session_state["logged_in"] = False
    st.rerun()

# ==========================================
# 6. MAIN PAGES LOGIC
# ==========================================

# --- PAGE: DASHBOARD ---
if choice == "📊 Dashboard":
    st.header("📊 Live Analytics Dashboard")
    
    filter_type = st.radio("Select View:", ["Day-Wise", "Month-Wise", "All Data"], horizontal=True)
    data = df.copy()
    
    if not data.empty and 'Date' in data.columns:
        if filter_type == "Day-Wise":
            sel_date = st.date_input("Choose Date", datetime.now())
            data = data[data['Date'].dt.date == sel_date]
        elif filter_type == "Month-Wise":
            if 'Month_Year' in data.columns:
                months = data['Month_Year'].dropna().unique().tolist()
                if months:
                    sel_month = st.selectbox("Choose Month", months)
                    data = data[data['Month_Year'] == sel_month]
                else:
                    data = pd.DataFrame()
    
    if data.empty:
        st.info("No production data available for the selected filter.")
    else:
        # Top Metrics
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f'<div class="metric-box"><div class="metric-title">Total Pieces</div><div class="metric-value">{data["Total Pcs"].sum()}</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-box"><div class="metric-title">Fresh Pieces</div><div class="metric-value">{data["Fresh Pcs"].sum()}</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-box"><div class="metric-title">Seconds Pieces</div><div class="metric-value">{data["Seconds Pcs"].sum()}</div></div>', unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Summaries
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="dashboard-card"><h3>🏭 ITEM PIECES SUMMARY</h3>', unsafe_allow_html=True)
            if 'Item Type' in data.columns: 
                st.dataframe(data.groupby('Item Type')['Total Pcs'].sum().reset_index(), use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="dashboard-card"><h3>🏢 PARTY PIECES SUMMARY</h3>', unsafe_allow_html=True)
            if 'Party Name' in data.columns: 
                st.dataframe(data.groupby('Party Name')['Total Pcs'].sum().reset_index(), use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

# --- PAGE: DATA ENTRY ---
elif choice == "📝 Production Entry":
    st.header("📝 Production Entry Input Panel")
    
    if user["role"] == "admin" or user.get("p_entry", False):
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
            
            if st.form_submit_button("SAVE PRODUCTION ENTRY", type="primary"):
                if fresh_pcs + seconds_pcs != total_pcs:
                    st.error("❌ Calculation mismatch! Fresh + Seconds must equal Total Pieces.")
                elif not design_no or not party_name:
                    st.error("❌ Please fill in all mandatory fields.")
                else:
                    try:
                        wb = openpyxl.load_workbook(EXCEL_FILE)
                        ws = wb["Supervisor Entry"]
                        ws.append([
                            date_input.strftime("%d-%m-%Y"), 
                            design_no, 
                            party_name, 
                            item_type, 
                            total_pcs, 
                            fresh_pcs, 
                            seconds_pcs, 
                            user["name"], 
                            challan_no
                        ])
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully!")
                    except Exception as e:
                        st.error(f"Error saving to Excel: {e}")
    else:
        st.warning("You do not have permission to make new entries.")

# --- PAGE: EDIT & DELETE RECORDS ---
elif choice == "📋 Edit & Delete Records":
    st.header("📋 Manage Production Records")
    
    if user["role"] == "admin" or user.get("p_view", False):
        st.info("💡 **How to use:** \n- **Edit:** Click inside any cell to change the data.\n- **Delete:** Click the checkbox on the left of a row and press the 'Delete' key on your keyboard.\n- **Save:** Click the Save button below to apply changes.")
        
        if not df.empty:
            # Drop the Month_Year helper column before editing
            if 'Month_Year' in df.columns:
                edit_view_df = df.drop(columns=['Month_Year'])
            else:
                edit_view_df = df.copy()

            # Filter for supervisor if not admin
            if user["role"] != "admin":
                user_data = edit_view_df[edit_view_df["Supervisor"] == user["name"]]
                if not user_data.empty:
                    if user.get("p_edit", False):
                        edited_df = st.data_editor(user_data, num_rows="dynamic", use_container_width=True)
                        if st.button("💾 Save My Changes", type="primary"):
                            # Admin + other supervisors data
                            rest_df = edit_view_df[edit_view_df["Supervisor"] != user["name"]]
                            # Combine
                            final_df = pd.concat([rest_df, edited_df], ignore_index=True)
                            final_df.to_excel(EXCEL_FILE, sheet_name="Supervisor Entry", index=False)
                            st.success("✅ Your records updated successfully!")
                            st.rerun()
                    else:
                        st.dataframe(user_data, hide_index=True, use_container_width=True)
                        st.warning("You have View-Only access. Contact Admin for Edit permissions.")
                else:
                    st.info("You haven't made any entries yet.")
            
            # Admin View (Can edit everything)
            else:
                edited_df = st.data_editor(edit_view_df, num_rows="dynamic", use_container_width=True)
                if st.button("💾 Save All Changes (Edit/Delete)", type="primary"):
                    edited_df.to_excel(EXCEL_FILE, sheet_name="Supervisor Entry", index=False)
                    st.success("✅ Database updated successfully!")
                    st.rerun()
        else:
            st.warning("No data found in the database.")
    else:
        st.error("Access Denied.")

# --- PAGE: TARGET SETTINGS (ADMIN ONLY) ---
elif choice == "🎯 Target Settings":
    st.header("🎯 Supervisor Targets")
    st.write("Set daily production targets for your supervisors.")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Current Targets")
        for sup, target in st.session_state["supervisor_targets"].items():
            st.write(f"**{sup}**: {target} Pcs")
            
    with c2:
        st.subheader("Update Target")
        with st.form("target_form"):
            sup_name = st.selectbox("Select Supervisor", list(st.session_state["supervisor_targets"].keys()))
            new_target = st.number_input("New Target (Pcs)", min_value=0, step=50)
            if st.form_submit_button("Update Target"):
                st.session_state["supervisor_targets"][sup_name] = new_target
                st.success(f"Target for {sup_name} updated to {new_target}!")
                st.rerun()

# --- PAGE: EMAIL BACKUP (ADMIN ONLY) ---
elif choice == "📧 Email Auto-Backup":
    st.header("📧 System Email Configurations")
    st.info("⚠️ Ensure you are using a 16-Digit Google App Password, NOT your regular Gmail password.")
    
    with st.form("email_settings"):
        sender = st.text_input("Sender Gmail Address", value=st.session_state["email_config"]["sender"])
        password = st.text_input("Sender Google App Password (16-Digit)", type="password", value=st.session_state["email_config"]["password"])
        receiver = st.text_input("Receiver Backup Email Address", value=st.session_state["email_config"]["receiver"])
        
        if st.form_submit_button("💾 Save Email Settings"):
            st.session_state["email_config"]["sender"] = sender
            st.session_state["email_config"]["password"] = password.replace(" ", "") # Removing accidental spaces
            st.session_state["email_config"]["receiver"] = receiver
            st.success("Email configurations saved locally in session!")
            
    st.markdown("---")
    st.subheader("Test Backup")
    if st.button("🚀 Test Send Instant Backup Now"):
        conf = st.session_state["email_config"]
        if not conf["sender"] or not conf["password"] or not conf["receiver"]:
            st.error("Please save email settings first!")
        elif not os.path.exists(EXCEL_FILE):
            st.error("Excel file does not exist yet to backup.")
        else:
            with st.spinner("Sending Email..."):
                success, msg = send_email_backup(conf["sender"], conf["password"], conf["receiver"], EXCEL_FILE)
                if success:
                    st.success(msg)
                else:
                    st.error(f"Email transmission failed. Error: {msg}")

# --- PAGE: FACTORY CONFIG (ADMIN ONLY) ---
elif choice == "⚙️ Factory Config":
    st.header("⚙️ Factory Configurations Desk")
    
    t1, t2, t3 = st.tabs(["Manage Parties", "Manage Items", "Manage Supervisor Access"])
    
    with t1:
        st.subheader("Current Parties")
        st.write(", ".join(st.session_state["party_options"]))
        with st.form("add_party"):
            new_p = st.text_input("New Party Name")
            if st.form_submit_button("Add Party") and new_p:
                if new_p not in st.session_state["party_options"]:
                    st.session_state["party_options"].append(new_p)
                    st.success(f"{new_p} added!")
                    st.rerun()
                
    with t2:
        st.subheader("Current Items")
        st.write(", ".join(st.session_state["item_options"]))
        with st.form("add_item"):
            new_i = st.text_input("New Item Name")
            if st.form_submit_button("Add Item") and new_i:
                new_i = new_i.upper()
                if new_i not in st.session_state["item_options"]:
                    st.session_state["item_options"].append(new_i)
                    st.success(f"{new_i} added!")
                    st.rerun()
                    
    with t3:
        st.subheader("Supervisor Edit Permissions")
        st.write("Grant or revoke Edit/Delete access for supervisors.")
        for uname, udata in st.session_state["users"].items():
            if udata["role"] == "supervisor":
                curr_status = udata.get("p_edit", False)
                new_status = st.checkbox(f"Allow {udata['name']} to Edit/Delete their records", value=curr_status, key=f"chk_{uname}")
                if new_status != curr_status:
                    st.session_state["users"][uname]["p_edit"] = new_status
                    st.success(f"Permissions updated for {udata['name']}. Please refresh.")

# ==========================================
# END OF FILE
# ==========================================
