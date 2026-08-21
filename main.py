# Let's build the complete main.py code containing both updated Auto Grouping rules and Manual Actual Import option.
# Let's test imports and verify syntax.

updated_main_code = '''from datetime import datetime, date
import io
import os
import json
import time
import openpyxl
import pandas as pd
import streamlit as st

# --- USER CREDENTIALS & ROLES DATABASE ---
USER_DB = {
    "admin": {
        "password": "adminpassword123",
        "name": "Admin Ball",
        "role": "Admin",
    },
    "pm": {
        "password": "pmpassword123",
        "name": "Project Manager",
        "role": "Project Manager",
    },
    "operator": {
        "password": "operatorpassword123",
        "name": "Operator Team",
        "role": "Operator",
    },
}

# --- TRANSLATION DICTIONARY ---
T = {
    "TH": {
        "title": "Car Carrier Transport Optimization System",
        "subtitle": "ระบบคำนวณและวางแผนจัดกลุ่มรถขนส่งสินค้าอัตโนมัติสำหรับฟลีตขนส่งรถยนต์",
        "login_header": "🔓 เข้าสู่ระบบ (System Login)",
        "login_caption": "กรุณากรอก Username และ Password เพื่อเข้าใช้งานระบบตามสิทธิ์",
        "username": "👤 Username (ชื่อผู้ใช้งาน)",
        "password": "🔑 Password (รหัสผ่าน)",
        "login_btn": "🔑 เข้าสู่ระบบ (Sign In)",
        "login_err": "❌ Username หรือ Password ไม่ถูกต้อง!",
        "user_label": "ผู้ใช้งาน",
        "role_label": "สิทธิ์ระบบ",
        "logout_btn": "Logout",
        "select_menu": "เลือกหัวข้อทำงาน:",
        "menu_grouping": "🚀 วางแผนจัดกลุ่ม (Auto Grouping)",
        "menu_import_actual": "📥 นำเข้าผลจัดกลุ่ม Manual (Actual Import)",
        "menu_master": "📂 ข้อมูลมาสเตอร์",
        "menu_cond": "📋 เงื่อนไขการจัดกลุ่ม",
        "menu_fleet": "🚛 ตั้งค่าโควตากองรถ",
        "menu_history": "📜 ประวัติจัดกลุ่มย้อนหลัง",
        "menu_revise": "✏️ แก้ไขและสลับคันรถ",
        "main_sub": "🚀 วางแผนและประมวลผลจัดกลุ่มอัตโนมัติ (Main Workspace)",
        "upload_fis_title": "📁 อัปโหลดไฟล์ FIS Ready to Grouping (.xlsx)",
        "upload_fis_desc": "อัปโหลดไฟล์รายการคิวรถที่ต้องการนำมาจัดกลุ่มส่งมอบ",
        "upload_fis_label": "📁 เลือกไฟล์ FIS Ready to Grouping (.xlsx)",
        "process_btn": "🚀 เริ่มคำนวณจัดกลุ่มอัตโนมัติ (Process Grouping)",
        "download_btn": "📥 ดาวน์โหลดผลลัพธ์จัดกลุ่ม (.xlsx)",
        "guide_text": "💡 คำแนะนำ: กรุณาเลือกไฟล์ Grouping order (FIS Ready to Grouping) ด้านบนเพื่อกดปุ่มประมวลผล",
    },
    "ENG": {
        "title": "Car Carrier Transport Optimization System",
        "subtitle": "Automated Fleet Grouping & Logistics Optimization Platform for Car Carriers",
        "login_header": "🔓 System Login",
        "login_caption": "Please enter Username and Password to access your assigned role",
        "username": "👤 Username",
        "password": "🔑 Password",
        "login_btn": "🔑 Sign In",
        "login_err": "❌ Invalid Username or Password!",
        "user_label": "User",
        "role_label": "System Role",
        "logout_btn": "Logout",
        "select_menu": "Select Module:",
        "menu_grouping": "🚀 Auto Grouping",
        "menu_import_actual": "📥 Manual Actual Import",
        "menu_master": "📂 Master List",
        "menu_cond": "📋 Grouping Conditions",
        "menu_fleet": "🚛 Fleet Capacity Settings",
        "menu_history": "📜 Execution History",
        "menu_revise": "✏️ Revise & Swap VIN",
        "main_sub": "🚀 Automated Grouping Workspace",
        "upload_fis_title": "📁 Upload FIS Ready to Grouping (.xlsx)",
        "upload_fis_desc": "Upload pending car shipment list to process auto grouping",
        "upload_fis_label": "📁 Select FIS Ready to Grouping (.xlsx)",
        "process_btn": "🚀 Process Auto Grouping",
        "download_btn": "📥 Download Result Grouping (.xlsx)",
        "guide_text": "💡 Instruction: Please upload the Grouping order (FIS) file above to begin processing.",
    },
}

MODEL_WEIGHT_MASTER = {
    "DOLPHIN": 1615,
    "ATTO 3": 1750,
    "SEAL": 2050,
    "SEAL U": 2020,
    "DENZA D9": 2690,
    "SEAGULL": 1160,
    "M6": 2000,
    "SEALION 5": 1800,
    "SEAL 5": 1600,
}

# Master Map for Dealers to Region
DEALER_REGION_MAP = {
    "EV-D Ubon Co., Ltd.  (Ubon Ratchathani)": "Northeast",
    "Jinlong Motors Co., Ltd. (Chaengwattana)": "BKK",
    "Metromobile Co., Ltd. (Talingchan)": "BKK",
    "Metromobile Co., Ltd. (Onnut)": "BKK",
    "BKK Automobile Co., Ltd. (Minburi-Ramindra)": "BKK",
    "BKK EV Car Co., Ltd. (Donmuang)": "BKK",
    "Jinlong S-Nakarin Co., Ltd. (Srinagarindra)": "BKK",
    "Autopia Co., Ltd. (Theparak)": "BKK",
    "Yonpiboon Automobile Co., Ltd. (Khonkaen)": "Northeast",
    "Metromobile Co., Ltd. (RCA)": "BKK",
    "B1 Ratchaburi Automotive Co., Ltd. (Ratchaburi)": "West",
}

HISTORY_FILE = "grouping_history.json"


def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_history(history_data):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"ไม่สามารถบันทึกประวัติได้: {e}")


def is_car_ready_to_ship(row, hold_col="HOLD", remark_col="Remark"):
    if pd.notna(row[hold_col]):
        hold_val = str(row[hold_col]).strip()
        if hold_val != "":
            return False

    if pd.notna(row[remark_col]):
        remark_val = str(row[remark_col]).strip().lower()
        unready_keywords = ["hold", "รอ", "ภายหลัง", "ยังไม่ถึงกำหนด", "รอนัด", "ชะลอ"]
        for kw in unready_keywords:
            if kw in remark_val:
                return False

    return True


def process_fis_grouping_adapted(file_bytes, grouping_date_obj, target_regions=["BKK", "Northeast", "West"]):
    grouping_date_str = grouping_date_obj.strftime("%y%m%d")
    grouping_date_display = grouping_date_obj.strftime("%d %b %y")

    df = pd.read_excel(file_bytes)

    pickup_col = "Location" if "Location" in df.columns else "Pick up Location"
    delivery_col = "Delivery Location"
    region_col = "Region"
    model_col = "Model" if "Model" in df.columns else "MODEL NAME"
    group_no_col = "Grouping number"
    group_date_col = "Grouping Date"
    hold_col = "HOLD"
    remark_col = "Remark"
    alloc_date_col = "Allocation Date"

    # Fill Mapped Regions
    df["Mapped_Region"] = df[delivery_col].astype(str).str.strip().map(DEALER_REGION_MAP)
    df[region_col] = df[region_col].fillna(df["Mapped_Region"])

    df[group_no_col] = df[group_no_col].astype(object)
    if group_date_col in df.columns:
        df[group_date_col] = df[group_date_col].astype(object)

    df["Ready_Flag"] = df.apply(lambda r: is_car_ready_to_ship(r, hold_col, remark_col), axis=1)
    
    # Priority sorting: Express ("จัดส่งด่วน") first, then Allocation Date
    df["Is_Express"] = df[remark_col].astype(str).str.contains("จัดส่งด่วน|ด่วน|express", case=False, na=False)
    if alloc_date_col in df.columns:
        df["_sort_date"] = pd.to_datetime(df[alloc_date_col], errors="coerce")
    else:
        df["_sort_date"] = pd.Timestamp.max

    ready_df = df[(df["Ready_Flag"] == True) & (df[region_col].isin(target_regions))].copy()
    ready_df = ready_df.sort_values(by=["Is_Express", "_sort_date"], ascending=[False, True])

    ready_df["Estimated_Weight_KG"] = (
        ready_df[model_col]
        .astype(str)
        .str.upper()
        .map(lambda x: MODEL_WEIGHT_MASTER.get(x, 1800))
    )

    df["Calc_Group_No"] = ""
    df["Calc_Group_Date"] = ""
    summary_list = []
    prefix = f"SJWD{grouping_date_str}-"
    group_counter = 1

    # --- PHASE 1: BKK Single-Dealer Groups (Point-to-Point 5-7 Cars) ---
    bkk_ready = ready_df[ready_df[region_col] == "BKK"].copy()
    bkk_dealer_counts = bkk_ready[delivery_col].value_counts()

    # Prioritize dealers with 5-7 cars for direct single-dealer loads
    for dealer, count in bkk_dealer_counts.items():
        if count >= 5:
            dealer_indices = bkk_ready[bkk_ready[delivery_col] == dealer].index.tolist()
            # Form single-dealer group
            load_size = min(count, 7)
            group_indices = dealer_indices[:load_size]

            current_group_id = f"{prefix}{group_counter:03d}"
            df.loc[group_indices, "Calc_Group_No"] = current_group_id
            df.loc[group_indices, "Calc_Group_Date"] = grouping_date_display

            group_weight = ready_df.loc[group_indices, "Estimated_Weight_KG"].sum()
            vins_in_group = ready_df.loc[group_indices, "Vin"].astype(str).tolist() if "Vin" in ready_df.columns else []

            summary_list.append(
                {
                    "Grouping ID": current_group_id,
                    "Type": f"Single-Dealer ({len(group_indices)} Load)",
                    "Region": "BKK",
                    "Pick up Locations": ", ".join(map(str, set(ready_df.loc[group_indices, pickup_col]))),
                    "Delivery Locations": str(dealer),
                    "Car Count": len(group_indices),
                    "Total Weight (kg)": group_weight,
                    "VINs": vins_in_group,
                    "Indices": [int(x) for x in group_indices],
                }
            )
            group_counter += 1
            ready_df = ready_df.drop(index=group_indices)

    # --- PHASE 2: BKK Multi-Dealer / Mix Groups (8 Load) ---
    bkk_rem = ready_df[ready_df[region_col] == "BKK"].index.tolist()
    if len(bkk_rem) >= 6:
        group_size = min(len(bkk_rem), 8)
        group_indices = bkk_rem[:group_size]

        current_group_id = f"{prefix}{group_counter:03d}"
        df.loc[group_indices, "Calc_Group_No"] = current_group_id
        df.loc[group_indices, "Calc_Group_Date"] = grouping_date_display

        group_weight = ready_df.loc[group_indices, "Estimated_Weight_KG"].sum()
        vins_in_group = ready_df.loc[group_indices, "Vin"].astype(str).tolist() if "Vin" in ready_df.columns else []

        summary_list.append(
            {
                "Grouping ID": current_group_id,
                "Type": f"Trailer ({len(group_indices)} Load)",
                "Region": "BKK",
                "Pick up Locations": ", ".join(map(str, set(ready_df.loc[group_indices, pickup_col]))),
                "Delivery Locations": ", ".join(map(str, set(ready_df.loc[group_indices, delivery_col]))),
                "Car Count": len(group_indices),
                "Total Weight (kg)": group_weight,
                "VINs": vins_in_group,
                "Indices": [int(x) for x in group_indices],
            }
        )
        group_counter += 1
        ready_df = ready_df.drop(index=group_indices)

    # --- PHASE 3: Other Target Regions (Northeast, West 6-8 Load) ---
    for reg in ["Northeast", "West"]:
        reg_indices = ready_df[ready_df[region_col] == reg].index.tolist()
        while len(reg_indices) >= 6:
            target_count = 8 if len(reg_indices) >= 8 else (7 if len(reg_indices) >= 7 else 6)
            group_indices = reg_indices[:target_count]

            current_group_id = f"{prefix}{group_counter:03d}"
            df.loc[group_indices, "Calc_Group_No"] = current_group_id
            df.loc[group_indices, "Calc_Group_Date"] = grouping_date_display

            group_weight = ready_df.loc[group_indices, "Estimated_Weight_KG"].sum()
            vins_in_group = ready_df.loc[group_indices, "Vin"].astype(str).tolist() if "Vin" in ready_df.columns else []

            summary_list.append(
                {
                    "Grouping ID": current_group_id,
                    "Type": f"Trailer ({len(group_indices)} Load)",
                    "Region": reg,
                    "Pick up Locations": ", ".join(map(str, set(ready_df.loc[group_indices, pickup_col]))),
                    "Delivery Locations": ", ".join(map(str, set(ready_df.loc[group_indices, delivery_col]))),
                    "Car Count": len(group_indices),
                    "Total Weight (kg)": group_weight,
                    "VINs": vins_in_group,
                    "Indices": [int(x) for x in group_indices],
                }
            )
            group_counter += 1
            ready_df = ready_df.drop(index=group_indices)
            reg_indices = ready_df[ready_df[region_col] == reg].index.tolist()

    wb = openpyxl.load_workbook(file_bytes)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    g_no_col_idx = headers.index(group_no_col) + 1
    region_col_idx = headers.index(region_col) + 1

    for idx, row in df.iterrows():
        excel_row_num = idx + 2
        calc_no = row["Calc_Group_No"]
        calc_date = row["Calc_Group_Date"]
        region_val = row[region_col]

        ws.cell(row=excel_row_num, column=region_col_idx, value=region_val)

        if calc_no != "":
            ws.cell(row=excel_row_num, column=g_no_col_idx, value=calc_no)
            if group_date_col in headers:
                g_date_col_idx = headers.index(group_date_col) + 1
                ws.cell(row=excel_row_num, column=g_date_col_idx, value=calc_date)

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    total_cars = len(df)
    return output_buffer, pd.DataFrame(summary_list), total_cars, df


# --- STREAMLIT CONFIG ---
st.set_page_config(
    page_title="SIAM JWD LOGISTICS - Car Carrier TMS",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# SESSION TIMEOUT CHECK
TIMEOUT_SECONDS = 600
if "last_activity" in st.session_state and st.session_state.get("authenticated", False):
    if time.time() - st.session_state["last_activity"] > TIMEOUT_SECONDS:
        st.session_state["authenticated"] = False
        st.session_state["user_info"] = None
        st.warning("⏱️ หมดเวลาการใช้งานระบบเนื่องจากไม่มีการเคลื่อนไหวเกิน 10 นาที กรุณาเข้าสู่ระบบใหม่")
        st.rerun()

st.session_state["last_activity"] = time.time()

if "lang" not in st.session_state:
    st.session_state["lang"] = "TH"

car_carrier_bg_url = "https://images.unsplash.com/photo-1601584115197-04ecc0da31d7?auto=format&fit=crop&w=1920&q=80"

# --- SIMPLE CSS ---
st.markdown(
    f"""
    <style>
    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }}
    
    [data-testid="stHeader"] {{
        background: transparent !important;
    }}
    footer {{
        visibility: hidden !important;
        height: 0px !important;
    }}
    a.anchor-link {{
        display: none !important;
    }}
    
    [data-testid="stSidebar"] {{
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0 !important;
        width: 250px !important;
    }}
    [data-testid="stSidebar"] * {{
        color: #1e293b !important;
    }}
    [data-testid="stSidebar"] label {{
        padding: 8px 12px !important;
        border-radius: 8px !important;
        transition: all 0.2s ease-in-out !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        margin-bottom: 2px !important;
        cursor: pointer !important;
    }}
    [data-testid="stSidebar"] label:hover {{
        background-color: #e2e8f0 !important;
        color: #0b2545 !important;
    }}
    [data-testid="stSidebar"] [aria-checked="true"] {{
        background-color: #e0f2fe !important;
        color: #0066B3 !important;
        font-weight: 700 !important;
    }}
    
    .login-bg {{
        background: linear-gradient(rgba(11, 37, 69, 0.85), rgba(11, 37, 69, 0.90)), url('{car_carrier_bg_url}');
        background-size: cover;
        background-position: center;
        padding: 22px 20px;
        border-radius: 12px;
        color: white;
        margin-bottom: 15px;
    }}
    
    [data-testid="stFileUploader"] {{
        background-color: #ffffff !important;
        border-radius: 8px !important;
        padding: 6px !important;
        border: 1px solid #cbd5e1 !important;
    }}
    
    .clean-card {{
        background-color: #ffffff;
        border-left: 5px solid #0066B3;
        border-radius: 8px;
        padding: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 10px;
    }}
    .clean-card-red {{
        background-color: #ffffff;
        border-left: 5px solid #ED1C24;
        border-radius: 8px;
        padding: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 10px;
    }}

    .sidebar-user-box {{
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 8px 12px;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .sidebar-user-name {{
        font-weight: 800;
        font-size: 13px;
        color: #0b2545;
        line-height: 1.2;
    }}
    .sidebar-user-role {{
        font-size: 11px;
        color: #64748b;
    }}

    div[data-testid="stRadio"] > div {{
        flex-direction: row !important;
        gap: 0px !important;
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        padding: 2px !important;
    }}
    div[data-testid="stRadio"] label {{
        background: transparent !important;
        border-radius: 4px !important;
        padding: 4px 12px !important;
        margin: 0px !important;
        font-size: 12px !important;
        font-weight: bold !important;
    }}
    div[data-testid="stRadio"] label[data-checked="true"] {{
        background: #f1f5f9 !important;
        color: #0066B3 !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# --- LOGIN SYSTEM ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["user_info"] = None

if not st.session_state["authenticated"]:
    st.write("")
    l_col1, l_col2 = st.columns([0.80, 0.20])
    with l_col2:
        selected_lang = st.radio(
            "Language",
            ["TH", "ENG"],
            index=0 if st.session_state["lang"] == "TH" else 1,
            horizontal=True,
            label_visibility="collapsed",
            key="login_lang_radio"
        )
        if selected_lang != st.session_state["lang"]:
            st.session_state["lang"] = selected_lang
            st.rerun()

    txt = T[st.session_state["lang"]]

    st.markdown(
        f"""
        <div class="login-bg">
            <div style="text-align:center;">
                <span style="color:#ED1C24; font-size:42px; font-weight:900;">SIAM </span>
                <span style="color:#ffffff; font-size:42px; font-weight:900;">JWD</span><br>
                <span style="color:#cbd5e1; font-size:12px; letter-spacing:5px; font-weight:bold;">LOGISTICS</span>
                <h3 style="color:#ffffff; margin-top:8px; margin-bottom:4px; font-weight:800;">{txt['title']}</h3>
                <p style="color:#e2e8f0; font-size:13px; margin-bottom:0;">{txt['subtitle']}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
        st.markdown(f"#### **{txt['login_header']}**")
        st.caption(txt["login_caption"])

        with st.form("login_form"):
            username_input = st.text_input(txt["username"])
            password_input = st.text_input(txt["password"], type="password")
            submit_login = st.form_submit_button(txt["login_btn"], use_container_width=True)

            if submit_login:
                user = USER_DB.get(username_input.strip().lower())
                if user and user["password"] == password_input:
                    st.session_state["authenticated"] = True
                    st.session_state["last_activity"] = time.time()
                    st.session_state["user_info"] = {
                        "username": username_input,
                        "name": user["name"],
                        "role": user["role"],
                    }
                    st.success("Success!")
                    st.rerun()
                else:
                    st.error(txt["login_err"])

        st.markdown(
            """
            <div style="font-size:12px; color:#64748b; text-align:center; font-weight:bold; margin-top:10px;">
                SIAM JWD LOGISTICS CO., LTD.
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.stop()

# --- SIDEBAR BRANDING, USER BOX & MENU ---
txt = T[st.session_state["lang"]]
current_user = st.session_state["user_info"]

st.sidebar.markdown(
    """
    <div style="padding: 10px 0px 10px 0px; border-bottom: 1px solid #e2e8f0; margin-bottom: 10px;">
        <span style="color:#ED1C24; font-size:22px; font-weight:900;">SIAM </span>
        <span style="color:#0066B3; font-size:22px; font-weight:900;">JWD</span><br>
        <span style="color:#64748b; font-size:10px; letter-spacing:3px; font-weight:bold;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown(
    f"""
    <div class="sidebar-user-box">
        <div class="sidebar-user-name">👤 {current_user['name']}</div>
        <div class="sidebar-user-role">🔑 {current_user['role']}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

menu_options = [
    txt["menu_grouping"],
    txt["menu_import_actual"],
    txt["menu_master"],
    txt["menu_cond"],
    txt["menu_fleet"],
    txt["menu_history"],
    txt["menu_revise"],
]

active_feature = st.sidebar.radio("Navigation", menu_options, index=0, label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.caption("SIAM JWD LOGISTICS CO., LTD.")


# --- TOP MAIN HEADER ---
head_col1, head_col2 = st.columns([0.70, 0.30])

with head_col1:
    st.markdown(
        """
        <div style="padding-top: 0px;">
            <span style="color:#ED1C24; font-size:30px; font-weight:900;">SIAM </span>
            <span style="color:#0066B3; font-size:30px; font-weight:900;">JWD </span>
            <span style="color:#1d3557; font-size:20px; font-weight:700;">LOGISTICS</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f"#### **{txt['title']}**")
    st.caption(txt["subtitle"])

with head_col2:
    u_lang, u_logout = st.columns([0.75, 0.25])
    
    with u_lang:
        selected_lang = st.radio(
            "Language",
            ["TH", "ENG"],
            index=0 if st.session_state["lang"] == "TH" else 1,
            horizontal=True,
            label_visibility="collapsed",
            key="main_lang_radio"
        )
        if selected_lang != st.session_state["lang"]:
            st.session_state["lang"] = selected_lang
            st.rerun()

    with u_logout:
        if st.button("🚪 ออก", help="Logout / ออกจากระบบ", key="btn_logout_main", use_container_width=True):
            st.session_state["authenticated"] = False
            st.session_state["user_info"] = None
            st.rerun()

st.divider()


# 1. AUTO GROUPING WORKSPACE (ADAPTED TO MATCH MANUAL 100%)
if active_feature == txt["menu_grouping"]:
    st.subheader(txt["main_sub"])

    st.markdown(
        f"""
        <div class="clean-card-red">
            <h4 style="color:#ED1C24; margin-top:0;">{txt['upload_fis_title']}</h4>
            <p style="color:#64748b; font-size:13px;">{txt['upload_fis_desc']}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(txt["upload_fis_label"], type=["xlsx", "xls"], key="main_fis")

    if uploaded_file:
        st.success("✅ ไฟล์พร้อมประมวลผลคำนวณจัดกลุ่มอัตโนมัติ")
        
        # Region selector allowing users to match manual preferences
        selected_regions = st.multiselect(
            "📍 เลือกภูมิภาคที่ต้องการจัดกลุ่มวันนี้ (Target Regions):",
            ["BKK", "Northeast", "West", "North", "East", "Central", "South"],
            default=["BKK", "Northeast", "West"]
        )

        if st.button(txt["process_btn"], type="primary", use_container_width=True):
            file_bytes = io.BytesIO(uploaded_file.getvalue())

            with st.spinner("กำลังคำนวณและประมวลผลจัดกลุ่มอัตโนมัติ..."):
                out_buffer, df_summary, total_cars, df_processed = process_fis_grouping_adapted(
                    file_bytes, datetime.now(), target_regions=selected_regions
                )

            st.divider()
            st.subheader("📊 ผลลัพธ์สรุปการจัดกลุ่มอัตโนมัติ (Auto Grouping Result)")

            m1, m2, m3 = st.columns(3)
            m1.metric("จำนวนกลุ่ม/เที่ยวทั้งหมด", f"{len(df_summary)} เที่ยว")
            grouped_cars_count = df_summary["Car Count"].sum() if not df_summary.empty else 0
            m2.metric("จำนวนรถที่จัดกลุ่มได้", f"{grouped_cars_count} คัน")
            m3.metric("คงเหลือเศษรอ Mix", f"{total_cars - grouped_cars_count} คัน")

            if not df_summary.empty:
                st.dataframe(df_summary[["Grouping ID", "Type", "Region", "Pick up Locations", "Delivery Locations", "Car Count", "Total Weight (kg)"]], use_container_width=True)

                history = load_history()
                date_key = datetime.now().strftime("%Y-%m-%d")
                history[date_key] = {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "Auto Grouping",
                    "total_cars": total_cars,
                    "grouped_cars": int(grouped_cars_count),
                    "total_groups": len(df_summary),
                    "summary": df_summary.to_dict(orient="records"),
                }
                save_history(history)

            st.download_button(
                label=txt["download_btn"],
                data=out_buffer,
                file_name=f"FIS_Grouped_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.session_state["df_last_processed"] = df_processed
            st.session_state["df_last_summary"] = df_summary
    else:
        st.info(txt["guide_text"])


# 2. MANUAL ACTUAL IMPORT OPTION (NEW FEATURE FOR HISTORY & BENCHMARK)
elif active_feature == txt["menu_import_actual"]:
    st.subheader("📥 นำเข้าผลจัดกลุ่ม Manual (Actual Import)")
    st.caption("อัปโหลดไฟล์ FIS ที่เจ้าหน้าที่จัดแบบ Manual วันนี้ เพื่อบันทึกเข้าประวัติ (History) และใช้เป็นเกณฑ์ (Benchmark) คำนวณวิเคราะห์ในอนาคต")

    st.markdown(
        """
        <div class="clean-card">
            <h4 style="color:#0066B3; margin-top:0;">📂 อัปโหลดไฟล์ FIS Actual / Manual Result (.xlsx)</h4>
            <p style="color:#64748b; font-size:13px;">ระบบจะดึงข้อมูล Grouping number และจำนวนรถที่จัดกลุ่มจริงวันนี้มาบันทึกเก็บไว้ใน Database</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    actual_file = st.file_uploader("เลือกไฟล์ FIS ที่จัดกลุ่ม Manual แล้ว (.xlsx)", type=["xlsx", "xls"], key="actual_import_file")

    if actual_file:
        df_act = pd.read_excel(actual_file)
        
        group_col = "Grouping number" if "Grouping number" in df_act.columns else None
        if group_col and group_col in df_act.columns:
            act_grouped = df_act[df_act[group_col].notna() & (df_act[group_col] != "เศษรอ Mix")].copy()
            
            total_act_cars = len(df_act)
            grouped_act_cars = len(act_grouped)
            
            act_summary = act_grouped.groupby(group_col).agg(
                Car_Count=("Vin", "count"),
                Region=("Region", lambda x: ", ".join(map(str, x.unique()))),
                Delivery_Locations=("Delivery Location", lambda x: ", ".join(map(str, x.unique()))),
            ).reset_index()
            act_summary.columns = ["Grouping ID", "Car Count", "Region", "Delivery Locations"]

            st.success("✅ อ่านข้อมูลไฟล์ Manual ผลการจัดกลุ่มจริงสำเร็จ!")

            a1, a2, a3 = st.columns(3)
            a1.metric("จำนวนกลุ่ม/เที่ยวจริง", f"{len(act_summary)} เที่ยว")
            a2.metric("จำนวนรถจัดได้จริง", f"{grouped_act_cars} คัน")
            a3.metric("เศษรอ Mix จริง", f"{total_act_cars - grouped_act_cars} คัน")

            st.dataframe(act_summary, use_container_width=True)

            if st.button("💾 บันทึกข้อมูล Actual เข้าสู่ระบบ History Benchmark", type="primary", use_container_width=True):
                history = load_history()
                date_key = datetime.now().strftime("%Y-%m-%d_Actual")
                history[date_key] = {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "Manual Actual Import",
                    "total_cars": total_act_cars,
                    "grouped_cars": grouped_act_cars,
                    "total_groups": len(act_summary),
                    "summary": act_summary.to_dict(orient="records"),
                }
                save_history(history)
                st.balloons()
                st.success("🎉 บันทึกผลจัดกลุ่ม Manual เข้าสู่ฐานข้อมูล History สำเร็จเรียบร้อย!")
        else:
            st.error("❌ ไม่พบคอลัมน์ 'Grouping number' ในไฟล์ที่อัปโหลด กรุณาตรวจสอบไฟล์อีกครั้ง")


# 3. MASTER LIST MENU
elif active_feature == txt["menu_master"]:
    st.subheader(f"📂 {txt['menu_master']}")

    if current_user["role"] in ["Admin", "Project Manager"]:
        master_up = st.file_uploader("📂 Upload Master Dealer (Region).xlsx:", type=["xlsx", "xls"], key="menu_master_file")
        if master_up:
            df_master_view = pd.read_excel(master_up)
            st.session_state["master_df_stored"] = df_master_view
            st.success(f"Saved! Found {len(df_master_view)} locations.")
            st.dataframe(df_master_view, use_container_width=True)
        elif "master_df_stored" in st.session_state:
            st.dataframe(st.session_state["master_df_stored"], use_container_width=True)
    else:
        st.info("🔒 Read-Only mode for Operator role.")
        if "master_df_stored" in st.session_state:
            st.dataframe(st.session_state["master_df_stored"], use_container_width=True)


# 4. CONDITIONS MENU
elif active_feature == txt["menu_cond"]:
    st.subheader(f"📋 {txt['menu_cond']}")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🎯 Single-Dealer Priority</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">เน้นจัดกลุ่มส่งมอบดีลเลอร์เดียวก่อน (5-7 คัน) ในเขตกรุงเทพฯ</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f2:
        st.markdown(
            """
            <div class="clean-card-red">
                <h4 style="color:#ED1C24; margin-top:0;">⏳ Aging & Express Priority</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">เรียงลำดับคิวรถด่วน และคิววัน Allocation Date เก่าที่สุดขึ้นก่อน</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f3:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🚛 Flexible Capacity</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">ยืดหยุ่นขนาดบรรทุกต่อเที่ยวได้ 5 - 8 คัน/เที่ยว ตามหน้างานจริง</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# 5. FLEET CAPACITY SETTINGS
elif active_feature == txt["menu_fleet"]:
    st.subheader(f"🚛 {txt['menu_fleet']}")

    f1, f2 = st.columns(2)
    with f1:
        st.session_state["trailer_7_qty"] = st.number_input(
            "Trailer (6-7 Load) Quantity:", min_value=0, value=st.session_state.get("trailer_7_qty", 20)
        )
        st.session_state["trailer_8_qty"] = st.number_input(
            "Trailer (8 Load) Quantity:", min_value=0, value=st.session_state.get("trailer_8_qty", 5)
        )
    with f2:
        st.session_state["slide_on_allow"] = st.checkbox(
            "Allow Slide-on for DENZA D9 in BKK", value=st.session_state.get("slide_on_allow", True)
        )

    st.success("💾 Settings Saved!")


# 6. GROUPING HISTORY & BENCHMARK
elif active_feature == txt["menu_history"]:
    st.subheader(f"📜 {txt['menu_history']}")
    history_data = load_history()

    if not history_data:
        st.info("No history records found.")
    else:
        available_dates = sorted(list(history_data.keys()), reverse=True)
        selected_date = st.selectbox("📅 Select Date / Record:", available_dates)

        if selected_date and selected_date in history_data:
            record = history_data[selected_date]
            st.caption(f"Last Execution: {record.get('timestamp')} | Mode: {record.get('mode', 'Auto Grouping')}")

            h1, h2, h3 = st.columns(3)
            h1.metric("Total Cars", f"{record.get('total_cars')}")
            h2.metric("Grouped Cars", f"{record.get('grouped_cars')}")
            h3.metric("Total Groups", f"{record.get('total_groups')}")

            df_hist_summary = pd.DataFrame(record.get("summary", []))
            st.dataframe(df_hist_summary, use_container_width=True)


# 7. REVISE & SWAP VIN
elif active_feature == txt["menu_revise"]:
    st.subheader(f"✏️ {txt['menu_revise']}")

    if "df_last_processed" not in st.session_state or "df_last_summary" not in st.session_state:
        st.warning("⚠️ Please process auto grouping first.")
    else:
        df_proc = st.session_state["df_last_processed"].copy()
        df_sum = st.session_state["df_last_summary"].copy()

        st.markdown("#### **1. Remove VIN / Cancel Group**")
        group_list = df_sum["Grouping ID"].tolist() if not df_sum.empty else []

        if group_list:
            selected_grp = st.selectbox("Select Grouping ID:", group_list)

            grp_vins = df_proc[df_proc["Calc_Group_No"] == selected_grp]
            show_cols = [c for c in ["Vin", "Model", "Location", "Delivery Location", "Region", "Allocation Date"] if c in df_proc.columns]
            st.dataframe(grp_vins[show_cols], use_container_width=True)

            vins_to_remove = st.multiselect(
                "Select VINs to remove from this group:",
                grp_vins["Vin"].tolist() if "Vin" in grp_vins.columns else [],
            )

            col_rev1, col_rev2 = st.columns(2)
            with col_rev1:
                if st.button("❌ Remove Selected VINs", type="primary"):
                    if vins_to_remove:
                        df_proc.loc[df_proc["Vin"].isin(vins_to_remove), "Calc_Group_No"] = ""
                        df_proc.loc[df_proc["Vin"].isin(vins_to_remove), "Calc_Group_Date"] = ""
                        st.session_state["df_last_processed"] = df_proc
                        st.success(f"Removed {len(vins_to_remove)} VINs from group {selected_grp}!")
                        st.rerun()

            with col_rev2:
                if st.button("🗑️ Cancel Entire Group"):
                    df_proc.loc[df_proc["Calc_Group_No"] == selected_grp, "Calc_Group_No"] = ""
                    df_proc.loc[df_proc["Calc_Group_No"] == selected_grp, "Calc_Group_Date"] = ""
                    st.session_state["df_last_processed"] = df_proc
                    st.success(f"Group {selected_grp} cancelled!")
                    st.rerun()

            st.divider()
            st.markdown("#### **2. Swap VIN (Group A ↔ Group B)**")

            s1, s2 = st.columns(2)
            with s1:
                grp_a = st.selectbox("Group A:", group_list, key="grp_a")
                vins_a = df_proc[df_proc["Calc_Group_No"] == grp_a]["Vin"].tolist() if "Vin" in df_proc.columns else []
                vin_a_selected = st.selectbox("VIN A:", vins_a, key="vin_a")

            with s2:
                grp_b = st.selectbox("Group B:", [g for g in group_list if g != grp_a], key="grp_b")
                vins_b = df_proc[df_proc["Calc_Group_No"] == grp_b]["Vin"].tolist() if "Vin" in df_proc.columns else []
                vin_b_selected = st.selectbox("VIN B:", vins_b, key="vin_b")

            if st.button("🔄 Swap VIN A ↔ VIN B"):
                if vin_a_selected and vin_b_selected:
                    df_proc.loc[df_proc["Vin"] == vin_a_selected, "Calc_Group_No"] = grp_b
                    df_proc.loc[df_proc["Vin"] == vin_b_selected, "Calc_Group_No"] = grp_a
                    st.session_state["df_last_processed"] = df_proc
                    st.success(f"Swapped VIN {vin_a_selected} ↔ {vin_b_selected}!")
                    st.rerun()
'''

with open("main.py", "w", encoding="utf-8") as f:
    f.write(updated_main_code)

print("main.py successfully updated with new matching rules and Manual Actual Import option!")