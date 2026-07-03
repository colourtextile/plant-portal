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

@@ -65,6 +70,7 @@
        margin-bottom: 25px;
        border-left: 5px solid #FF5733;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
    }
    
    .dashboard-card {
@@ -78,11 +84,9 @@
</style>
""", unsafe_allow_html=True)

# --- SESSION STATES INIT ---
if "supervisor_targets" not in st.session_state:
    st.session_state["supervisor_targets"] = {
        "Ramesh": 500,
        "Suresh": 400
    }
    st.session_state["supervisor_targets"] = {"Ramesh": 500, "Suresh": 400}

if "users" not in st.session_state:
    st.session_state["users"] = {
@@ -102,6 +106,62 @@
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
@@ -139,20 +199,23 @@
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)

    # --- 📐 UPDATED SIDEBAR BRAND FORMAT & USER DETAILS ---
    st.sidebar.markdown(f"""
    <div class="sidebar-brand-box">
        <div style="color: #FFC300; font-size: 11px; font-weight: 700; letter-spacing: 1.5px; margin-bottom: 3px;">SYSTEM PORTAL</div>
        <div style="color: #FFFFFF; font-size: 16px; font-weight: 800; letter-spacing: 0.5px;">WELCOME TO COLOUR TEXTILE</div>
        <div style="border-top: 1px solid rgba(255,255,255,0.15); margin: 8px 0px;"></div>
        <div style="color: #E0E6ED; font-size: 14px; font-weight: 500;">👤 User: <span style="color: #30c381; font-weight: 700;">{user['name']}</span></div>
        <div style="color: #B4C6E7; font-size: 12px; margin-top: 2px;">📋 Access Level: <b>{user['role'].upper()}</b></div>
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
        nav_choice = st.sidebar.radio("🧭 Navigation Menu", ["📊 Dashboard", "📝 Data Entry", "🎯 Target Settings"])
    else:
        nav_choice = st.sidebar.radio("🧭 Navigation Menu", ["📊 Dashboard"])
        nav_options.extend(["📝 Data Entry", "🎯 Target Settings", "📧 Setup Email Auto-Backup"])
        
    nav_choice = st.sidebar.radio("🧭 Navigation Menu", nav_options)

    excel_loaded = False
    df = pd.DataFrame()
@@ -171,6 +234,15 @@
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
@@ -179,8 +251,29 @@
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
        if user["role"] == "admin" and nav_choice == "🎯 Target Settings":
        elif user["role"] == "admin" and nav_choice == "🎯 Target Settings":
            st.subheader("🎯 Supervisor Target Configurations")
            all_users = list(st.session_state["users"].keys())
            sups_list = [st.session_state["users"][u]["name"] for u in all_users if st.session_state["users"][u]["role"] == "supervisor"]
@@ -216,25 +309,92 @@
                        wb.save(EXCEL_FILE)
                        st.success("🎉 Entry Saved Successfully!")

        # --- DASHBOARD LOGIC (ADMIN & SUPERVISOR TARGET MONITOR) ---
        # --- DASHBOARD LOGIC ---
        else:
            st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Live Target Tracking Dashboard</h2>", unsafe_allow_html=True)
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

            # --- 🎯 DYNAMIC AUTO-LESS TARGET CALCULATION CORE ---
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

                    # Counting total done pieces by this supervisor
                    if excel_loaded and not df.empty:
                        done_pcs = df[df["Supervisor"] == s_name]["Total Pcs"].sum()
                    else:
                        done_pcs = 0

                    remaining_tgt = allocated_tgt - done_pcs
                    if remaining_tgt < 0: remaining_tgt = 0 # Target achieved case
                    if remaining_tgt < 0: remaining_tgt = 0

                    status_txt = "✅ Done" if remaining_tgt == 0 and allocated_tgt > 0 else "⏳ Pending"

@@ -247,16 +407,15 @@
                    })

            target_summary_df = pd.DataFrame(target_summary_data)
            st.markdown("**⚡ Live Supervisors Target Status Table (Auto Less Active):**")
            st.dataframe(target_summary_df, hide_index=True, use_container_width=True)
            st.markdown("---")

            # --- 👥 SUPERVISOR PERSONAL VIEW / ADMIN VIEW PANELS ---
            if user["role"] == "supervisor":
                can_entry = user.get("p_entry", True)
                can_view_logs = user.get("p_view", True)
                can_edit_logs = user.get("p_edit", False)

                # PERMISSION 1: DATA ENTRY ALLOWED?
                if can_entry:
                    st.subheader("📝 Production Entry Input Panel")
                    with st.form("entry_form_sup", clear_on_submit=True):
@@ -280,34 +439,20 @@
                                ws = wb["Supervisor Entry"]
                                ws.append([date_input.strftime("%d-%m-%Y"), design_no, party_name, item_type, total_pcs, fresh_pcs, seconds_pcs, user["name"], challan_no])
                                wb.save(EXCEL_FILE)
                                st.success("🎉 Entry Saved! Live Target Summary Updated automatically.")
                                st.success("🎉 Entry Saved!")
                                st.rerun()
                else:
                    st.warning("⚠️ Aapko system me Data Entry karne ki permission nahi hai.")

                # PERMISSION 2: LOGS VIEW & STRICT EDIT RESTRICTION
                if can_view_logs:
                    st.markdown("---")
                    st.subheader("📋 Your Done Production Entries")
                    sup_df = df[df["Supervisor"] == user["name"]] if excel_loaded else pd.DataFrame()

                    if not sup_df.empty:
                        if can_edit_logs:
                            st.info("💡 Admin has granted you Modify Rights for logs below:")
                            edited_sup_df = st.data_editor(sup_df, hide_index=True, use_container_width=True)
                            if st.button("Apply Changes"):
                                st.success("Changes Saved!")
                            st.data_editor(sup_df, hide_index=True, use_container_width=True)
                        else:
                            # Strict Guard Statement
                            st.warning("🔒 READ ONLY: Entry delete ya modify karne ka permission aapke paas nahi hai. Admin se permission lein.")
                            st.dataframe(sup_df, hide_index=True, use_container_width=True)
                    else:
                        st.info("Aapne aaj koi entry submit nahi ki hai.")
                else:
                    st.info("ℹ️ Logs sheet dekhne ki permission is account ko nahi mili hai.")
                    
            else:
                # --- ADMIN VIEW PANEL ---
                st.subheader("📋 Production Master Logs (Full Control)")
                st.dataframe(df, hide_index=True, use_container_width=True, height=250)

@@ -400,8 +545,6 @@
                                        }
                                        st.success(f"Supervisor '{add_name}' Created!")
                                        st.rerun()
                                    else:
                                        st.error("❌ ID already exists!")

                    with col_u2:
                        st.markdown("**✏️ Edit Info & Checkbox Permissions**")
@@ -429,14 +572,12 @@
                                    st.session_state["users"][selected_sup]["p_edit"] = edit_cb_edit
                                    st.success("Account & Checkbox Permissions updated!")
                                    st.rerun()
                        else:
                            st.info("No supervisors setup.")

                    with col_u3:
                        st.markdown("**❌ Remove Supervisor**")
                        if sups_only:
                            sup_to_remove = st.selectbox("Select Supervisor ID to Delete", sups_only, key="sel_sup_rem")
                            if st.button("Delete Supervisor Account", type="primary"):
                                del st.session_state["users"][sup_to_remove]
                                st.warning("Account deleted from database!")
                                st.rerun()
