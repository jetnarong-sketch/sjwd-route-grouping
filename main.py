from datetime import datetime, date
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
        "title": "ระบบจัดการการขนส่งยานพาหนะ",
        "subtitle": "จัดการคำสั่งซื้อ วางแผนการขนส่ง และติดตามการส่งมอบอย่างมีประสิทธิภาพ",
        "login_header": "เข้าสู่ระบบ",
        "login_caption": "กรอกข้อมูลเพื่อเข้าใช้งานระบบ",
        "username": "อีเมล / Username",
        "password": "รหัสผ่าน",
        "login_btn": "เข้าสู่ระบบ",
        "login_err": " Username หรือ Password ไม่ถูกต้อง!",
        "user_label": "ผู้ใช้งาน",
        "role_label": "สิทธิ์ระบบ",
        "logout_btn": "Logout",
        "select_menu": "เลือกหัวข้อทำงาน:",
        "menu_grouping": "แดชบอร์ดผู้ดูแลระบบ",
        "menu_master": "จัดการคำสั่งซื้อ",
        "menu_cond": "วางแผนการเดินทาง",
        "menu_fleet": "ตั้งค่าโควตากองรถ",
        "menu_history": "ประวัติจัดกลุ่มย้อนหลัง",
        "menu_revise": "แก้ไขและสลับคันรถ",
        "main_sub": "แดชบอร์ดผู้ดูแลระบบ",
        "upload_fis_title": "อัปโหลดไฟล์ FIS Ready to Grouping (.xlsx)",
        "upload_fis_desc": "อัปโหลดไฟล์รายการคิวรถที่ต้องการนำมาจัดกลุ่มส่งมอบ",
        "upload_fis_label": "เลือกไฟล์ FIS Ready to Grouping (.xlsx)",
        "process_btn": "เริ่มคำนวณจัดกลุ่มอัตโนมัติ (Process Grouping)",
        "download_btn": "ดาวน์โหลดผลลัพธ์จัดกลุ่ม (.xlsx)",
        "guide_text": " คำแนะนำ: กรุณาเลือกไฟล์ Grouping order (FIS) ด้านบนเพื่อกดปุ่มประมวลผล",
    },
    "ENG": {
        "title": "Car Carrier Transport Management System",
        "subtitle": "Manage shipment orders, plan logistics, and track deliveries efficiently",
        "login_header": "System Login",
        "login_caption": "Please enter your credentials to log in",
        "username": "Email / Username",
        "password": "Password",
        "login_btn": "Sign In",
        "login_err": " Invalid Username or Password!",
        "user_label": "User",
        "role_label": "System Role",
        "logout_btn": "Logout",
        "select_menu": "Select Module:",
        "menu_grouping": "Admin Dashboard",
        "menu_master": "Order Management",
        "menu_cond": "Trip Planning",
        "menu_fleet": "Fleet Settings",
        "menu_history": "Grouping History",
        "menu_revise": "Revise & Swap VIN",
        "main_sub": "System Administrator Dashboard",
        "upload_fis_title": "Upload FIS Ready to Grouping (.xlsx)",
        "upload_fis_desc": "Upload pending car shipment list to process auto grouping",
        "upload_fis_label": "Select FIS Ready to Grouping (.xlsx)",
        "process_btn": "Process Auto Grouping",
        "download_btn": "Download Result Grouping (.xlsx)",
        "guide_text": " Instruction: Please upload the FIS file above to begin processing.",
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


# --- STREAMLIT CONFIG ---
st.set_page_config(
    page_title="SIAM JWD LOGISTICS - TMS",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# SESSION TIMEOUT CHECK (10 MINS INACTIVITY AUTO LOGOUT)
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

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }
    
    [data-testid="stHeader"] {
        background: transparent !important;
    }
    footer {
        visibility: hidden !important;
        height: 0px !important;
    }
    a.anchor-link {
        display: none !important;
    }
    
    /* SIDEBAR THEME EXACTLY LIKE IMAGE 2 */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e5e7eb !important;
        width: 260px !important;
    }
    
    [data-testid="stSidebar"] * {
        color: #1f2937 !important;
    }
    
    [data-testid="stSidebar"] label {
        padding: 10px 14px !important;
        border-radius: 8px !important;
        transition: all 0.2s ease-in-out !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        margin-bottom: 3px !important;
        cursor: pointer !important;
    }
    
    [data-testid="stSidebar"] label:hover {
        background-color: #f3f4f6 !important;
        color: #1d4ed8 !important;
    }
    
    [data-testid="stSidebar"] [aria-checked="true"] {
        background-color: #eff6ff !important;
        color: #1d4ed8 !important;
        font-weight: 700 !important;
        border-left: 4px solid #1d4ed8 !important;
    }
    
    /* LOGIN PAGE SPLIT SCREEN (IMAGE 1 STYLE) */
    .login-left-banner {
        background: linear-gradient(135deg, #0b2545 0%, #1d3557 100%);
        height: 100vh;
        padding: 60px 40px;
        color: white;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }

    /* สวิตช์สลับภาษาทรงแคปซูลมินิมอล (文A TH | EN) */
    .lang-switcher-btn div.stButton > button {
        background-color: #ffffff !important;
        color: #6b7280 !important;
        border: 1px solid #e5e7eb !important;
        border-radius: 6px !important;
        height: 32px !important;
        padding: 0px 10px !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    .lang-switcher-btn-active div.stButton > button {
        background-color: #ffffff !important;
        color: #1d4ed8 !important;
        border: 1px solid #1d4ed8 !important;
        border-radius: 6px !important;
        height: 32px !important;
        padding: 0px 10px !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }

    /* ปุ่ม Logout สัญลักษณ์ [-> ตรงตามภาพที่ 2 */
    .logout-btn-exact div.stButton > button {
        background-color: transparent !important;
        color: #374151 !important;
        border: none !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        height: 34px !important;
        width: 34px !important;
        padding: 0px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    .logout-btn-exact div.stButton > button:hover {
        color: #dc2626 !important;
    }

    /* ปุ่มฟอร์มเข้าสู่ระบบสีน้ำเงินเข้มสไตล์ภาพที่ 1 */
    .login-submit-btn div.stButton > button {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        height: 44px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(29, 78, 216, 0.25) !important;
    }

    .brand-pill-btn {
        background-color: #eff6ff;
        border: 1px solid #bfdbfe;
        color: #1d4ed8;
        border-radius: 6px;
        padding: 4px 10px;
        font-size: 12px;
        font-weight: 700;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        height: 32px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- LOGIN SYSTEM ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["user_info"] = None

if not st.session_state["authenticated"]:
    # TOP RIGHT LANGUAGE SWITCHER
    st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
    l_top1, l_top2 = st.columns([0.84, 0.16])
    with l_top2:
        is_th = st.session_state["lang"] == "TH"
        st.markdown("<div style='display:flex; align-items:center; gap:6px; background:#fff; border:1px solid #e5e7eb; padding:2px 8px; border-radius:8px;'>", unsafe_allow_html=True)
        c_ic, c_t, c_sp, c_e = st.columns([0.2, 0.35, 0.1, 0.35])
        with c_ic:
            st.markdown("<span style='font-size:13px; font-weight:bold;'>文<sub>A</sub></span>", unsafe_allow_html=True)
        with c_t:
            st.markdown(f'<div class="{"lang-switcher-btn-active" if is_th else "lang-switcher-btn"}">', unsafe_allow_html=True)
            if st.button("TH", key="lg_th_btn"):
                st.session_state["lang"] = "TH"
                st.session_state["last_activity"] = time.time()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with c_sp:
            st.markdown("<span style='color:#d1d5db;'>|</span>", unsafe_allow_html=True)
        with c_e:
            st.markdown(f'<div class="{"lang-switcher-btn-active" if not is_th else "lang-switcher-btn"}">', unsafe_allow_html=True)
            if st.button("EN", key="lg_en_btn"):
                st.session_state["lang"] = "ENG"
                st.session_state["last_activity"] = time.time()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    txt = T[st.session_state["lang"]]

    # LOGIN PAGE SPLIT SCREEN DESIGN (EXACT MATCH IMAGE 1)
    col_banner, col_form = st.columns([1.1, 1])

    with col_banner:
        st.markdown(
            f"""
            <div style="background: linear-gradient(180deg, #0b2545 0%, #133a68 100%); padding: 50px 40px; border-radius: 16px; color: white; min-height: 480px; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <div style="background: white; width: 140px; padding: 8px; border-radius: 8px; margin-bottom: 30px;">
                        <span style="color:#ED1C24; font-size:20px; font-weight:900;">SIAM </span>
                        <span style="color:#0066B3; font-size:20px; font-weight:900;">JWD</span><br>
                        <span style="color:#64748b; font-size:9px; letter-spacing:2px; font-weight:bold;">LOGISTICS</span>
                    </div>
                    <h2 style="font-size: 28px; font-weight: 800; color: #ffffff; margin-bottom: 12px;">{txt['title']}</h2>
                    <p style="color: #cbd5e1; font-size: 14px; line-height: 1.6;">{txt['subtitle']}</p>
                </div>
                <div style="font-size: 12px; color: #94a3b8; font-weight: 500;">
                    © 2026 YARD - Transportation Management System (TMS). สงวนลิขสิทธิ์
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_form:
        st.markdown("<div style='padding: 20px 30px;'>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='font-size:28px; font-weight:800; color:#111827; margin-bottom:4px;'>{txt['login_header']}</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='color:#6b7280; font-size:14px; margin-bottom:24px;'>{txt['login_caption']}</p>", unsafe_allow_html=True)

        with st.form("login_form"):
            username_input = st.text_input(txt["username"], placeholder="jetnarong@siamjwd.com")
            password_input = st.text_input(txt["password"], type="password", placeholder="••••••••••••")
            
            st.markdown("<div style='text-align:right; margin-bottom:15px;'><a href='#' style='color:#2563eb; font-size:13px; text-decoration:none; font-weight:600;'>ลืมรหัสผ่าน?</a></div>", unsafe_allow_html=True)
            
            st.markdown('<div class="login-submit-btn">', unsafe_allow_html=True)
            submit_login = st.form_submit_button(f"→]  {txt['login_btn']}", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

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
        st.markdown("</div>", unsafe_allow_html=True)

    st.stop()


# --- SIDEBAR BRANDING & MENU (IMAGE 2 STYLE) ---
st.sidebar.markdown(
    """
    <div style="padding: 12px 10px; display: flex; align-items: center; gap: 8px; border-bottom: 1px solid #f3f4f6; margin-bottom: 12px;">
        <span style="font-size: 20px;">☰</span>
        <div>
            <span style="color:#ED1C24; font-size:18px; font-weight:900;">SIAM </span>
            <span style="color:#0066B3; font-size:18px; font-weight:900;">JWD</span><br>
            <span style="color:#64748b; font-size:8px; letter-spacing:2px; font-weight:bold;">LOGISTICS</span>
        </div>
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

active_feature = st.sidebar.radio("Navigation", menu_options, index=0, label_visibility="collapsed")

st.sidebar.markdown("---")
st.sidebar.caption("© 2026 SIAM JWD LOGISTICS CO., LTD.")


# --- TOP MAIN HEADER: LOGO LEFT & USER + LANG TOP RIGHT (EXACT MATCH IMAGE 2) ---
head_col1, head_col2 = st.columns([0.4, 0.6])

with head_col1:
    st.markdown(
        f"""
        <div style="padding-top: 5px;">
            <h2 style="font-size: 22px; font-weight: 800; color: #111827; margin: 0;">{txt['main_sub']}</h2>
            <p style="font-size: 13px; color: #6b7280; margin: 0;">กำลังแสดงข้อมูลของ REVER AUTOMOTIVE</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

with head_col2:
    # องค์ประกอบมุมขวาบนเรียงขนานเป๊ะ บาลานซ์ตามภาพที่ 2
    is_th = st.session_state["lang"] == "TH"
    
    st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
    u_lang, u_brand, u_user, u_logout = st.columns([0.28, 0.28, 0.32, 0.12])

    # 1. ปุ่มสลับภาษา 文A TH | EN
    with u_lang:
        st.markdown("<div style='display:flex; align-items:center; gap:4px; background:#fff; border:1px solid #e5e7eb; padding:2px 6px; border-radius:6px; height:32px;'>", unsafe_allow_html=True)
        c_ic, c_t, c_sp, c_e = st.columns([0.2, 0.35, 0.05, 0.35])
        with c_ic:
            st.markdown("<span style='font-size:12px; font-weight:bold;'>文<sub>A</sub></span>", unsafe_allow_html=True)
        with c_t:
            st.markdown(f'<div class="{"lang-switcher-btn-active" if is_th else "lang-switcher-btn"}">', unsafe_allow_html=True)
            if st.button("TH", key="hdr_th_btn"):
                st.session_state["lang"] = "TH"
                st.session_state["last_activity"] = time.time()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with c_sp:
            st.markdown("<span style='color:#d1d5db; font-size:12px;'>|</span>", unsafe_allow_html=True)
        with c_e:
            st.markdown(f'<div class="{"lang-switcher-btn-active" if not is_th else "lang-switcher-btn"}">', unsafe_allow_html=True)
            if st.button("EN", key="hdr_en_btn"):
                st.session_state["lang"] = "ENG"
                st.session_state["last_activity"] = time.time()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # 2. ปุ่มเลือกแบรนด์ RE REVER ∨
    with u_brand:
        st.markdown(
            """
            <div class="brand-pill-btn">
                <span style="background:#2563eb; color:white; border-radius:4px; padding:1px 4px; font-size:10px;">RE</span>
                <span>REVER</span>
                <span style="font-size:10px; color:#6b7280;">∨</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 3. ชื่อผู้ใช้งาน Admin Ball / Admin (ไร้กรอบครอบทับสไตล์ภาพที่ 2)
    with u_user:
        st.markdown(
            f"""
            <div style="text-align: right; padding-right: 8px;">
                <div style="font-weight: 700; font-size: 13px; color: #111827; line-height: 1.1;">{current_user['name']}</div>
                <div style="font-size: 11px; color: #6b7280; line-height: 1.1;">{current_user['role']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 4. ปุ่ม Logout ไอคอน [->
    with u_logout:
        st.markdown('<div class="logout-btn-exact">', unsafe_allow_html=True)
        if st.button("[→", help="Logout / ออกจากระบบ", key="btn_logout_main"):
            st.session_state["authenticated"] = False
            st.session_state["user_info"] = None
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

st.divider()


# 1. AUTO GROUPING WORKSPACE
if active_feature == txt["menu_grouping"]:
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
            st.success(" File is ready for grouping process.")
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
                    st.error(" Missing delivery locations found in Master list!")
                    for m_loc in missing_locs:
                        st.write(f"-  **{m_loc}**")
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
        st.info(" Read-Only mode for Operator role.")
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
                <h4 style="color:#0066B3; margin-top:0;"> Auto Matching</h4>
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
                <h4 style="color:#0066B3; margin-top:0;"> Route & Slide-on</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">6-8 Cars per trailer load (DENZA D9 in BKK uses Slide-on).</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# 4. FLEET CAPACITY SETTINGS
elif active_feature == txt["menu_fleet"]:
    st.subheader(f" {txt['menu_fleet']}")

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
                if st.button(" Remove Selected VINs", type="primary"):
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