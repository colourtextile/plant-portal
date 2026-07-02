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
    /* Premium Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #0f172a !important;
        padding: 20px 10px;
    }
    
    .premium-sidebar-header {
        background: linear-gradient(135deg, #1e293b, #334155);
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 30px;
        border-left: 5px solid #38bdf8;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        text-align: center;
    }
    
    .user-chip {
        display: inline-block;
        background: #0ea5e9;
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .reportview-container { background: #f8f9fa; }
    
    .global-header {
        text-align: center; margin: 0px 0px 20px 0px; font-weight: 800; font-size: 44px; letter-spacing: 2px;
        background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff);
        background-size: 300% 300%; -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        animation: gradientShift 6s ease infinite; border-bottom: 2px solid #eaeaea; padding-bottom: 10px;
    }
    
    .brand-title {
        text-align: center; margin: 0px; font-weight: 800; font-size: 42px; letter-spacing: 1.5px;
        background: linear-gradient(45deg, #FF5733, #FFC300, #30c381, #247ba0, #a066ff);
        background-size: 300% 300%; -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        animation: gradientShift 6s ease infinite; padding-top: 40px;
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; }
    }

    div[data-testid="stForm"] label, .stMarkdown label, label[data-testid="stWidgetLabel"] {
        color: #2C3E50 !important; font-weight: 600 !important; font-size: 14px !important; margin-bottom: 5px !important;
    }
    
    .dashboard-card {
        background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px; border: 1px solid #eef2f5;
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

if "email_config" not in st.session_state:
    st.session_state["email_config"] = {
        "sender": "your_email@gmail.com",
        "password": "your_app_password",
        "receiver": "receiver_email@gmail.com"
    }

# --- 📧 HELPER FUNCTION ---
def send_daywise_backup_email():
    cfg = st.session_state["email_config"]
    current_today = datetime.now().strftime("%d-%m-%Y")
    if not os.path.exists(EXCEL_FILE): return False
    try:
        main_df = pd.read_excel(EXCEL_FILE, sheet_name="Supervisor Entry")
        main_df['Date'] = main_df['Date'].astype(str).str.strip()
        day_wise_df = main_df[main_df['Date'] == current_today]
        if day_wise_df.empty: day_wise_df = pd.DataFrame(columns=main_df.columns)
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
        server.starttls()
        server.login(cfg["sender"], cfg["password"])
        server.sendmail(cfg["sender"], cfg["receiver"], msg.as_string())
        server.close()
        return True
    except Exception: return False

# --- 🔒 LOGIN SCREEN ---
if not st.session_state["logged_in"]:
    st.markdown("""<style>.stApp {background-image: linear-gradient(rgba(0, 0, 0, 0.6), rgba(0, 0, 0, 0.6)), url("https://images.unsplash.com/photo-1617627143750-d86bc21e42bb?q=80&w=1920&auto=format&fit=crop"); background-size: cover; background-position: center; background-attachment: fixed;} div[data-testid="stForm"] label { color: #FFFFFF !important; }</style>""", unsafe_allow_html=True)
    st.markdown('<h1 class="brand-title">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username / ID", placeholder="Enter your registered ID")
            password = st.text_input("Password", type="password", placeholder="••••••••")
            if st.form_submit_button("SECURE SYSTEM LOGIN", use_container_width=True):
                if username.strip() in st.session_state["users"] and st.session_state["users"][username.strip()]["password"] == password.strip():
                    st.session_state["logged_in"] = True
                    st.session_state["current_user"] = st.session_state["users"][username.strip()]
                    st.rerun()
                else: st.error("❌ Invalid ID or Password")

# --- 👑 MAIN PORTAL SCREEN ---
else:
    user = st.session_state["current_user"]
    st.markdown('<h1 class="global-header">COLOUR TEXTILE</h1>', unsafe_allow_html=True)
    
    # Premium Sidebar
    st.sidebar.markdown(f"""
    <div class="premium-sidebar-header">
        <div style="color: #38bdf8; font-size: 12px; letter-spacing: 2px;">WELCOME</div>
        <div style="color: #ffffff; font-size: 20px; font-weight: 900;">COLOUR TEXTILE</div>
        <div class="user-chip">{user['role'].upper()}</div>
        <div style="color: #cbd5e1; font-size: 13px; margin-top: 15px;">👤 {user['name']}</div>
    </div>
    """, unsafe_allow_html=True)
    
    nav_options = ["📊 Dashboard"]
    if user["role"] == "admin": nav_options.extend(["📝 Data Entry", "🎯 Target Settings", "📧 Setup Email Auto-Backup"])
    nav_choice = st.sidebar.radio("🧭 Navigation Menu", nav_options)
    
    if st.sidebar.button("🚪 Logout System"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = None
        st.rerun()
        
    # --- (Baki pura original code yahan rahega...) ---
    st.write("---")
    st.info("Portal Loaded Successfully. Navigate using the sidebar.")
```[cite: 1]

Ab batayein, aur kya changes karne hain?
