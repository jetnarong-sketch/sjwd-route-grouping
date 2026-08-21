# Generate updated main.py where Top Header contains:
# Left side: SIAM JWD LOGISTICS Header
# Right side: Language Switcher + User Info Card (Name, Role, Logout button) matching image 2 layout perfectly.

updated_main = '''from datetime import datetime, date
import io
import os
import json
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

# --- TRANSLATION DICTIONARY (TH / ENG) ---
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
        "menu_master": "📂 Master list",
        "menu_cond": "📋 Conditions (เงื่อนไขการจัดกลุ่ม)",
        "menu_fleet": "🚛 Fleet Capacity Settings",
        "menu_history": "📜 ประวัติการจัดกลุ่มย้อนหลัง",
        "menu_revise": "✏️ Revise & Swap VIN",
        "main_sub": "🚀 วางแผนและประมวลผลจัดกลุ่มอัตโนมัติ (Main Workspace)",
        "upload_fis_title": "📁 Upload FIS Ready to Grouping (.xlsx)",
        "upload_fis_desc": "อัปโหลดไฟล์รายการคิวรถที่ต้องการนำมาจัดกลุ่มส่งมอบ",
        "upload_fis_label": "📁 เลือกไฟล์ FIS Ready to Grouping (.xlsx)",
        "process_btn": "🚀 เริ่มคำนวณจัดกลุ่มอัตโนมัติ (Process Grouping)",
        "download_btn": "📥 Download Result grouping (.xlsx)",
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
        "menu_grouping": "🚀 Auto Grouping Workspace",
        "menu_master": "📂 Master List Management",
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


def get_max_delivery_locations(region_str):
    reg = str(region_str).strip().upper()
    if reg == "BKK":
        return 4
    elif reg == "EAST":
        return 6
    else:
        return 6


def process_fis_grouping_with_capacity(file_bytes, master_region_df, grouping_date_obj, fleet_capacity):
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

    prefix = f"SJWD{grouping_date_str}-"
    group_counter = 1

    master_map = dict(
        zip(
            master_region_df["Delivery Location"].astype(str).str.strip(),
            master_region_df["Region"].astype(str).str.strip(),
        )
    )

    df_delivery_clean = df[delivery_col].astype(str).str.strip()
    missing_locations = [loc for loc in df_delivery_clean.unique() if loc not in master_map and pd.notna(loc) and loc != "nan"]

    if missing_locations:
        return None, None, None, missing_locations, df

    df["Mapped_Region"] = df_delivery_clean.map(master_map)
    df[region_col] = df[region_col].fillna(df["Mapped_Region"])

    df[group_no_col] = df[group_no_col].astype(object)
    if group_date_col in df.columns:
        df[group_date_col] = df[group_date_col].astype(object)

    df["Ready_Flag"] = df.apply(lambda r: is_car_ready_to_ship(r, hold_col, remark_col), axis=1)
    ready_df = df[df["Ready_Flag"] == True].copy()

    if alloc_date_col in ready_df.columns:
        temp_alloc_date = pd.to_datetime(ready_df[alloc_date_col], errors="coerce")
        ready_df = ready_df.assign(_temp_sort_date=temp_alloc_date).sort_values(by="_temp_sort_date", ascending=True)

    ready_df["Estimated_Weight_KG"] = (
        ready_df[model_col]
        .astype(str)
        .str.upper()
        .map(lambda x: MODEL_WEIGHT_MASTER.get(x, 1800))
    )

    df["Calc_Group_No"] = ""
    df["Calc_Group_Date"] = ""
    summary_list = []

    if fleet_capacity.get("slide_on", True):
        slide_on_mask = (
            ready_df[model_col]
            .astype(str)
            .str.upper()
            .str.contains("DENZA D9|D9", regex=True)
        ) & (ready_df[region_col].astype(str).str.upper().str.strip() == "BKK")
        slide_on_indices = ready_df[slide_on_mask].index.tolist()

        for idx in slide_on_indices:
            current_group_id = f"{prefix}{group_counter:03d}"
            df.loc[idx, "Calc_Group_No"] = current_group_id
            df.loc[idx, "Calc_Group_Date"] = grouping_date_display

            group_weight = ready_df.loc[idx, "Estimated_Weight_KG"]
            pick_loc = ready_df.loc[idx, pickup_col]
            del_loc = ready_df.loc[idx, delivery_col]

            summary_list.append(
                {
                    "Grouping ID": current_group_id,
                    "Type": "Slide-on",
                    "Region": "BKK (Slide on)",
                    "Pick up Locations": str(pick_loc),
                    "Delivery Locations": str(del_loc),
                    "Car Count": 1,
                    "Total Weight (kg)": group_weight,
                    "VINs": [str(ready_df.loc[idx, "Vin"])] if "Vin" in ready_df.columns else [],
                    "Indices": [int(idx)],
                }
            )
            group_counter += 1
        ready_df_trailer = ready_df.drop(index=slide_on_indices)
    else:
        ready_df_trailer = ready_df.copy()

    trailer_7_quota = fleet_capacity.get("trailer_7", 999)
    trailer_8_quota = fleet_capacity.get("trailer_8", 999)

    for region_name, region_batch in ready_df_trailer.groupby(region_col, dropna=False):
        pending_indices = region_batch.index.tolist()
        max_deliv = get_max_delivery_locations(region_name)

        while len(pending_indices) >= 6:
            if trailer_7_quota <= 0 and trailer_8_quota <= 0:
                break

            group_indices = []
            pickups_in_group = set()
            deliveries_in_group = set()

            if trailer_7_quota > 0:
                target_count = 7 if len(pending_indices) >= 7 else 6
                use_quota_type = 7
            else:
                target_count = 8 if len(pending_indices) >= 8 else 6
                use_quota_type = 8

            for idx in pending_indices:
                curr_pickup = region_batch.loc[idx, pickup_col]
                curr_delivery = region_batch.loc[idx, delivery_col]

                temp_pickups = pickups_in_group | {curr_pickup}
                temp_deliveries = deliveries_in_group | {curr_delivery}

                if len(temp_pickups) <= 4 and len(temp_deliveries) <= max_deliv:
                    group_indices.append(idx)
                    pickups_in_group = temp_pickups
                    deliveries_in_group = temp_deliveries

                if len(group_indices) == target_count:
                    break

            if len(group_indices) >= 6:
                current_group_id = f"{prefix}{group_counter:03d}"
                df.loc[group_indices, "Calc_Group_No"] = current_group_id
                df.loc[group_indices, "Calc_Group_Date"] = grouping_date_display

                group_weight = ready_df.loc[group_indices, "Estimated_Weight_KG"].sum()
                vins_in_group = ready_df.loc[group_indices, "Vin"].astype(str).tolist() if "Vin" in ready_df.columns else []

                summary_list.append(
                    {
                        "Grouping ID": current_group_id,
                        "Type": f"Trailer ({len(group_indices)} Load)",
                        "Region": region_name,
                        "Pick up Locations": ", ".join(map(str, pickups_in_group)),
                        "Delivery Locations": ", ".join(map(str, deliveries_in_group)),
                        "Car Count": len(group_indices),
                        "Total Weight (kg)": group_weight,
                        "VINs": vins_in_group,
                        "Indices": [int(x) for x in group_indices],
                    }
                )

                if use_quota_type == 7:
                    trailer_7_quota -= 1
                else:
                    trailer_8_quota -= 1

                group_counter += 1
                for g_idx in group_indices:
                    pending_indices.remove(g_idx)
            else:
                break

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
    return output_buffer, pd.DataFrame(summary_list), total_cars, [], df


# --- STREAMLIT CONFIG & COMPACT RESPONSIVE THEME ---
st.set_page_config(
    page_title="SIAM JWD LOGISTICS - Car Carrier TMS",
    page_icon="🚚",
    layout="wide",
)

if "lang" not in st.session_state:
    st.session_state["lang"] = "TH"

car_carrier_bg_url = "https://images.unsplash.com/photo-1601584115197-04ecc0da31d7?auto=format&fit=crop&w=1920&q=80"

st.markdown(
    f"""
    <style>
    .block-container {{
        padding-top: 0.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }}
    
    .css-1544g2n, .e1ewe6wb4, [data-testid="stHeader"], footer {{
        visibility: hidden !important;
        height: 0px !important;
    }}
    a.anchor-link {{
        display: none !important;
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
    
    [data-testid="stSidebar"] {{
        background-color: #0b2545 !important;
    }}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {{
        color: #ffffff !important;
    }}
    
    [data-testid="stFileUploader"] {{
        background-color: #f8f9fa !important;
        border-radius: 8px !important;
        padding: 6px !important;
        border: 1px solid #cbd5e1 !important;
    }}
    [data-testid="stFileUploader"] * {{
        color: #1a202c !important;
    }}
    [data-testid="stFileUploader"] button {{
        background-color: #0066B3 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        border: none !important;
    }}
    
    div.stButton > button:first-child {{
        background-color: #ED1C24 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 8px 20px !important;
        box-shadow: 0px 4px 10px rgba(237, 28, 36, 0.3) !important;
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
    
    /* สไตล์การ์ดผู้ใช้งานมุมขวาบนตรงตามตัวอย่างระบบ SIAM JWD */
    .user-profile-box {{
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 6px 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    }}
    
    .user-name-text {{
        font-weight: 800;
        font-size: 13px;
        color: #0b2545;
        line-height: 1.1;
    }}
    
    .user-role-text {{
        font-size: 11px;
        color: #64748b;
    }}
    
    .lang-btn-box [data-testid="stMarkdownContainer"] p {{
        font-size: 12px;
        font-weight: bold;
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
    # TOP RIGHT LANGUAGE SELECTOR FOR LOGIN PAGE
    top_col1, top_col2 = st.columns([0.80, 0.20])
    with top_col2:
        selected_lang_choice = st.radio(
            "Lang",
            ["TH 🇹🇭", "ENG 🇬🇧"],
            horizontal=True,
            key="top_right_login_lang",
            label_visibility="collapsed",
        )
        clean_choice = "TH" if "TH" in selected_lang_choice else "ENG"
        if clean_choice != st.session_state["lang"]:
            st.session_state["lang"] = clean_choice
            st.rerun()

    txt = T[st.session_state["lang"]]

    # Banner โลโก้และหัวข้อ
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

# --- SIDEBAR BRANDING & MENU ---
st.sidebar.markdown(
    """
    <div style="text-align:center; padding: 10px 0px 15px 0px;">
        <span style="color:#ED1C24; font-size:26px; font-weight:900;">SIAM </span>
        <span style="color:#ffffff; font-size:26px; font-weight:900;">JWD</span><br>
        <span style="color:#8da9c4; font-size:11px; letter-spacing:4px; font-weight:bold;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)

txt = T[st.session_state["lang"]]
current_user = st.session_state["user_info"]

menu_options = [
    txt["menu_grouping"],
    txt["menu_master"],
    txt["menu_cond"],
    txt["menu_fleet"],
    txt["menu_history"],
    txt["menu_revise"],
]

active_feature = st.sidebar.radio(txt["select_menu"], menu_options, index=0)

st.sidebar.divider()
st.sidebar.caption("SIAM JWD LOGISTICS CO., LTD.")


# --- TOP MAIN HEADER: LOGO LEFT & USER + LANG TOP RIGHT ---
head_col1, head_col2 = st.columns([0.60, 0.40])

with head_col1:
    st.markdown(
        """
        <div style="padding-top: 5px;">
            <span style="color:#ED1C24; font-size:36px; font-weight:900;">SIAM </span>
            <span style="color:#0066B3; font-size:36px; font-weight:900;">JWD </span>
            <span style="color:#1d3557; font-size:24px; font-weight:700;">LOGISTICS</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(f"#### **{txt['title']}**")
    st.caption(txt["subtitle"])

with head_col2:
    # Top Right Controls: Language Selector + User Profile Card & Logout Button
    u_col1, u_col2, u_col3 = st.columns([0.42, 0.43, 0.15])
    
    with u_col1:
        top_lang = st.radio(
            "LangToggle",
            ["TH 🇹🇭", "ENG 🇬🇧"],
            horizontal=True,
            key="top_main_lang_toggle",
            label_visibility="collapsed",
        )
        clean_top_lang = "TH" if "TH" in top_lang else "ENG"
        if clean_top_lang != st.session_state["lang"]:
            st.session_state["lang"] = clean_top_lang
            st.rerun()

    with u_col2:
        role_label = current_user["role"]
        st.markdown(
            f"""
            <div class="user-profile-box">
                <div>
                    <div class="user-name-text">{current_user['name']}</div>
                    <div class="user-role-text">{role_label}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with u_col3:
        if st.button("🚪", help="Logout / ออกจากระบบ", use_container_width=True):
            st.session_state["authenticated"] = False
            st.session_state["user_info"] = None
            st.rerun()

st.divider()


# 1. AUTO GROUPING WORKSPACE
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

    st.write("")
    master_df_to_use = st.session_state.get("master_df_stored", None)

    if uploaded_file:
        if master_df_to_use is None:
            st.warning("⚠️ Master list missing! Please upload Dealer (Region).xlsx:")
            master_region_file_local = st.file_uploader("📂 Dealer (Region).xlsx:", type=["xlsx", "xls"], key="temp_master_up")
            if master_region_file_local:
                master_df_to_use = pd.read_excel(master_region_file_local)
                st.session_state["master_df_stored"] = master_df_to_use

        if master_df_to_use is not None:
            st.success("✅ File is ready for grouping process.")
            if st.button(txt["process_btn"], type="primary", use_container_width=True):
                file_bytes = io.BytesIO(uploaded_file.getvalue())

                capacity_settings = {
                    "trailer_7": st.session_state.get("trailer_7_qty", 20),
                    "trailer_8": st.session_state.get("trailer_8_qty", 5),
                    "slide_on": st.session_state.get("slide_on_allow", True),
                }

                with st.spinner("Processing grouping..."):
                    out_buffer, df_summary, total_cars, missing_locs, df_processed = process_fis_grouping_with_capacity(
                        file_bytes, master_df_to_use, datetime.now(), capacity_settings
                    )

                if missing_locs:
                    st.error("❌ Missing delivery locations found in Master list!")
                    for m_loc in missing_locs:
                        st.write(f"- 📍 **{m_loc}**")
                else:
                    st.divider()
                    st.subheader("📊 Grouping Result Summary")

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total Groups", f"{len(df_summary)}")
                    grouped_cars_count = df_summary["Car Count"].sum() if not df_summary.empty else 0
                    m2.metric("Grouped Cars", f"{grouped_cars_count}")
                    m3.metric("Pending / Unassigned Cars", f"{total_cars - grouped_cars_count}")

                    if not df_summary.empty:
                        st.dataframe(df_summary, use_container_width=True)

                        history = load_history()
                        date_key = datetime.now().strftime("%Y-%m-%d")
                        history[date_key] = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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


# 2. MASTER LIST MENU
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


# 3. CONDITIONS MENU
elif active_feature == txt["menu_cond"]:
    st.subheader(f"📋 {txt['menu_cond']}")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🎯 Auto Matching</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">Delivery Location to Region auto mapping from Master list.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f2:
        st.markdown(
            """
            <div class="clean-card-red">
                <h4 style="color:#ED1C24; margin-top:0;">⏳ Aging Priority</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">Sort by Allocation Date from oldest to newest.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f3:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🚛 Route & Slide-on</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">6-8 Cars per trailer load (DENZA D9 in BKK uses Slide-on).</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# 4. FLEET CAPACITY SETTINGS
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


# 5. GROUPING HISTORY
elif active_feature == txt["menu_history"]:
    st.subheader(f"📜 {txt['menu_history']}")
    history_data = load_history()

    if not history_data:
        st.info("No history records found.")
    else:
        available_dates = sorted(list(history_data.keys()), reverse=True)
        selected_date = st.selectbox("📅 Select Date:", available_dates)

        if selected_date and selected_date in history_data:
            record = history_data[selected_date]
            st.caption(f"Last Execution: {record.get('timestamp')}")

            h1, h2, h3 = st.columns(3)
            h1.metric("Total Cars", f"{record.get('total_cars')}")
            h2.metric("Grouped Cars", f"{record.get('grouped_cars')}")
            h3.metric("Total Groups", f"{record.get('total_groups')}")

            df_hist_summary = pd.DataFrame(record.get("summary", []))
            st.dataframe(df_hist_summary, use_container_width=True)


# 6. REVISE & SWAP VIN
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
    f.write(updated_main)

print("Successfully generated main.py matching image 2 UI layout!")