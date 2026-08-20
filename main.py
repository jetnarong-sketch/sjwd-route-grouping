from datetime import datetime, date
import io
import os
import json
import openpyxl
import pandas as pd
import streamlit as st

# Master Data น้ำหนักรถ (กก.)
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
        unready_keywords = [
            "hold",
            "รอ",
            "ภายหลัง",
            "ยังไม่ถึงกำหนด",
            "รอนัด",
            "ชะลอ",
        ]
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


def process_fis_grouping_with_capacity(
    file_bytes, master_region_df, grouping_date_obj, fleet_capacity
):
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
    missing_locations = [
        loc
        for loc in df_delivery_clean.unique()
        if loc not in master_map and pd.notna(loc) and loc != "nan"
    ]

    if missing_locations:
        return None, None, None, missing_locations, df

    df["Mapped_Region"] = df_delivery_clean.map(master_map)
    df[region_col] = df[region_col].fillna(df["Mapped_Region"])

    df[group_no_col] = df[group_no_col].astype(object)
    if group_date_col in df.columns:
        df[group_date_col] = df[group_date_col].astype(object)

    df["Ready_Flag"] = df.apply(
        lambda r: is_car_ready_to_ship(r, hold_col, remark_col), axis=1
    )
    ready_df = df[df["Ready_Flag"] == True].copy()

    if alloc_date_col in ready_df.columns:
        temp_alloc_date = pd.to_datetime(
            ready_df[alloc_date_col], errors="coerce"
        )
        ready_df = ready_df.assign(
            _temp_sort_date=temp_alloc_date
        ).sort_values(by="_temp_sort_date", ascending=True)

    ready_df["Estimated_Weight_KG"] = (
        ready_df[model_col]
        .astype(str)
        .str.upper()
        .map(lambda x: MODEL_WEIGHT_MASTER.get(x, 1800))
    )

    df["Calc_Group_No"] = ""
    df["Calc_Group_Date"] = ""
    summary_list = []

    # 1. Slide on D9 BKK
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

    # 2. Capacity Tracking
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


# --- STREAMLIT CONFIG & CORPORATE THEME ---
st.set_page_config(
    page_title="SIAM JWD LOGISTICS - Auto Grouping System",
    page_icon="🚛",
    layout="wide",
)

st.markdown(
    """
    <style>
    /* แถบข้าง Sidebar สีน้ำเงินเข้มองค์กร */
    [data-testid="stSidebar"] {
        background-color: #0b2545 !important;
    }
    /* ข้อความใน Sidebar เป็นสีขาว */
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3, 
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] label {
        color: #ffffff !important;
    }
    
    /* กล่อง File Uploader อ่านง่าย */
    [data-testid="stFileUploader"] {
        background-color: #f8f9fa !important;
        border-radius: 10px !important;
        padding: 10px !important;
        border: 1px solid #cbd5e1 !important;
    }
    [data-testid="stFileUploader"] * {
        color: #1a202c !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #0066B3 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        border: none !important;
    }
    
    /* ปุ่มประมวลผล สีแดงองค์กร */
    div.stButton > button:first-child {
        background-color: #ED1C24 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 24px !important;
        box-shadow: 0px 4px 10px rgba(237, 28, 36, 0.3) !important;
    }
    
    /* การ์ดฟีเจอร์หน้าหลัก */
    .clean-card {
        background-color: #ffffff;
        border-left: 5px solid #0066B3;
        border-radius: 8px;
        padding: 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 15px;
    }
    .clean-card-red {
        background-color: #ffffff;
        border-left: 5px solid #ED1C24;
        border-radius: 8px;
        padding: 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 15px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- SIDEBAR BRANDING & FUNCTIONAL MENU ---
st.sidebar.markdown(
    """
    <div style="text-align:center; padding: 10px 0px 20px 0px;">
        <span style="color:#ED1C24; font-size:26px; font-weight:900; font-family:sans-serif;">SIAM </span>
        <span style="color:#ffffff; font-size:26px; font-weight:900; font-family:sans-serif;">JWD</span><br>
        <span style="color:#8da9c4; font-size:12px; letter-spacing:4px; font-weight:bold;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.title("⚙️ Control Panel")
st.sidebar.caption("กดเลือกหัวข้อที่ Sidebar เพื่อแสดงผลลัพธ์บน Main Panel ขวามือ")

active_feature = st.sidebar.radio(
    "เลือกหัวข้อทำงาน:",
    [
        "🚀 วางแผนจัดกลุ่ม (Auto Grouping)",
        "📂 Master list (Dealer Region)",
        "📋 Conditions (เงื่อนไขการจัดกลุ่ม)",
        "🚛 Fleet Capacity Settings",
        "📜 ประวัติการจัดกลุ่มย้อนหลัง",
        "✏️ Revise & Swap VIN",
    ],
    index=0,
)

st.sidebar.divider()
st.sidebar.caption("SIAM JWD LOGISTICS CO., LTD.")


# --- MAIN PANEL HEADER ---
st.markdown(
    """
    <div style="padding-bottom: 10px;">
        <span style="color:#ED1C24; font-size:46px; font-weight:900; font-family:sans-serif;">SIAM </span>
        <span style="color:#0066B3; font-size:46px; font-weight:900; font-family:sans-serif;">JWD </span>
        <span style="color:#1d3557; font-size:32px; font-weight:700; font-family:sans-serif;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("### **Auto Fleet Grouping & Logistics Optimization System**")
st.caption("ระบบคำนวณและวางแผนจัดกลุ่มรถขนส่งสินค้าอัตโนมัติ (Automated Car Carrier Optimization)")
st.divider()


# --- DISPLAY CONTENT ON MAIN PANEL RIGHT SIDE BASED ON SIDEBAR CLICK ---

# 1. AUTO GROUPING WORKSPACE ON MAIN PANEL
if active_feature == "🚀 วางแผนจัดกลุ่ม (Auto Grouping)":
    st.subheader("🚀 วางแผนและประมวลผลจัดกลุ่มอัตโนมัติ (Main Workspace)")
    st.caption("อัปโหลดไฟล์ข้อมูล และกดปุ่มประมวลผลจัดกลุ่มบนหน้าหลักขวามือได้ทันที")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">1. Master list</h4>
                <p style="color:#64748b; font-size:13px;">อัปโหลดไฟล์ <b>Dealer (Region).xlsx</b> (หรือใช้ Master ล่าสุดในระบบ)</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        master_region_file = st.file_uploader(
            "📂 เลือกไฟล์ Dealer (Region).xlsx", type=["xlsx", "xls"], key="main_master"
        )

    with col_up2:
        st.markdown(
            """
            <div class="clean-card-red">
                <h4 style="color:#ED1C24; margin-top:0;">2. Grouping order</h4>
                <p style="color:#64748b; font-size:13px;">อัปโหลดไฟล์ <b>FIS Ready to Grouping (.xlsx)</b> คิวรถที่ต้องการจัดกลุ่ม</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader(
            "📁 เลือกไฟล์ FIS Ready to Grouping (.xlsx)", type=["xlsx", "xls"], key="main_fis"
        )

    st.write("")

    if uploaded_file and master_region_file:
        st.success("✅ อัปโหลดไฟล์ครบถ้วนแล้ว พร้อมสำหรับการประมวลผลจัดกลุ่ม")
        if st.button("🚀 เริ่มคำนวณจัดกลุ่มอัตโนมัติ (Process Grouping)", type="primary", use_container_width=True):
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            master_df = pd.read_excel(master_region_file)

            capacity_settings = {
                "trailer_7": st.session_state.get("trailer_7_qty", 20),
                "trailer_8": st.session_state.get("trailer_8_qty", 5),
                "slide_on": st.session_state.get("slide_on_allow", True),
            }

            with st.spinner("กำลังตรวจสอบ Master Region และประมวลผลคิวขนส่ง..."):
                out_buffer, df_summary, total_cars, missing_locs, df_processed = (
                    process_fis_grouping_with_capacity(
                        file_bytes, master_df, datetime.now(), capacity_settings
                    )
                )

            if missing_locs:
                st.error("❌ ไม่สามารถประมวลผลได้ เนื่องจากพบ Delivery Location ที่ไม่มีในไฟล์ Master!")
                st.warning("กรุณาเพิ่มข้อมูล Delivery Location ดังต่อไปนี้ลงในไฟล์ Master list (Dealer (Region).xlsx) ก่อนประมวลผลใหม่:")
                for m_loc in missing_locs:
                    st.write(f"- 📍 **{m_loc}**")
            else:
                st.divider()
                st.subheader("📊 ผลลัพธ์การจัดกลุ่มจัดส่ง (SIAM JWD LOGISTICS)")

                m1, m2, m3 = st.columns(3)
                m1.metric("จำนวนกลุ่มที่สร้างได้", f"{len(df_summary)} กลุ่ม")
                grouped_cars_count = df_summary["Car Count"].sum() if not df_summary.empty else 0
                m2.metric("จำนวนรถที่จัดกลุ่มสำเร็จ", f"{grouped_cars_count} คัน")
                m3.metric("รถที่ไม่เข้าเงื่อนไข/รอจัดกลุ่มใหม่", f"{total_cars - grouped_cars_count} คัน")

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
                    st.success("💾 บันทึกประวัติการจัดกลุ่มลงระบบเรียบร้อยแล้ว!")

                st.download_button(
                    label="📥 Download Result grouping (.xlsx)",
                    data=out_buffer,
                    file_name=f"FIS_Grouped_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.session_state["df_last_processed"] = df_processed
                st.session_state["df_last_summary"] = df_summary
    else:
        st.info("💡 **คำแนะนำ:** กรุณาเลือกไฟล์ Master list และ Grouping order ด้านบนเพื่อกดปุ่มประมวลผลจัดกลุ่ม")


# 2. MASTER LIST SEPARATE MENU ON MAIN PANEL
elif active_feature == "📂 Master list (Dealer Region)":
    st.subheader("📂 การจัดการข้อมูล Master list (Dealer & Region Mapping)")
    st.caption("ตรวจสอบและอัปโหลดข้อมูลแมปสถานที่จัดส่งสินค้า (Delivery Location) คู่กับ Region")

    master_up = st.file_uploader("📂 อัปโหลดไฟล์ Master Dealer (Region).xlsx ใหม่:", type=["xlsx", "xls"], key="menu_master_file")
    
    if master_up:
        df_master_view = pd.read_excel(master_up)
        st.success(f"อัปโหลดไฟล์ Master สำเร็จ! พบรายการสถานที่ส่งทั้งหมด {len(df_master_view)} รายการ")
        st.dataframe(df_master_view, use_container_width=True)
    else:
        st.info("📌 คุณสามารถอัปโหลดไฟล์ Master Dealer (Region).xlsx เพื่อตรวจสอบจับคู่ Delivery Location และ Region ที่นี่")


# 3. CONDITIONS SEPARATE MENU ON MAIN PANEL
elif active_feature == "📋 Conditions (เงื่อนไขการจัดกลุ่ม)":
    st.subheader("📋 เงื่อนไขการจัดกลุ่มจัดส่งอัตโนมัติ (Grouping Conditions)")
    st.caption("ระบบจะนำเงื่อนไขทั้ง 3 ส่วนนี้ไปประมวลผลในการจัดกลุ่มรถทุกครั้ง")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🎯 Auto Matching</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">จับคู่ Delivery Location กับ Region และเติมช่องตกหล่นจาก Master อัตโนมัติ</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f2:
        st.markdown(
            """
            <div class="clean-card-red">
                <h4 style="color:#ED1C24; margin-top:0;">⏳ Aging Priority</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">เรียงลำดับคิวรถตาม Allocation Date จากอดีตไปหาปัจจุบัน ป้องกันสินค้าค้างส่ง</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_f3:
        st.markdown(
            """
            <div class="clean-card">
                <h4 style="color:#0066B3; margin-top:0;">🚛 Route & Slide-on</h4>
                <p style="color:#4a5568; font-size:14px; margin:0;">คุมจำนวนรถเฉลี่ย 6-7 คันต่อเทรลเลอร์ (DENZA D9 ใน BKK จัดส่งแบบ Slide on ได้ทันที)</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# 4. FLEET CAPACITY SETTINGS ON MAIN PANEL
elif active_feature == "🚛 Fleet Capacity Settings":
    st.subheader("🚛 ตั้งค่า Fleet Capacity (กำหนดโควตากองรถ)")
    st.caption("กำหนดจำกัดจำนวนเทรลเลอร์และเงื่อนไขประเภทรถขนส่งสำหรับนำไปคำนวณในระบบ")

    f1, f2 = st.columns(2)
    with f1:
        st.session_state["trailer_7_qty"] = st.number_input(
            "จำนวน Trailer (6-7 Load) ที่มีพร้อมวิ่ง:", min_value=0, value=st.session_state.get("trailer_7_qty", 20)
        )
        st.session_state["trailer_8_qty"] = st.number_input(
            "จำนวน Trailer (8 Load) ที่มีพร้อมวิ่ง:", min_value=0, value=st.session_state.get("trailer_8_qty", 5)
        )
    with f2:
        st.session_state["slide_on_allow"] = st.checkbox(
            "อนุญาตให้ส่งแบบ Slide-on สำหรับ DENZA D9 ใน BKK ได้ทันที", value=st.session_state.get("slide_on_allow", True)
        )

    st.success("💾 บันทึกการตั้งค่าโควตากองรถเรียบร้อยแล้ว!")


# 5. GROUPING HISTORY ON MAIN PANEL
elif active_feature == "📜 ประวัติการจัดกลุ่มย้อนหลัง":
    st.subheader("📜 เรียกดูประวัติการจัดกลุ่มย้อนหลัง")
    history_data = load_history()

    if not history_data:
        st.info("ยังไม่มีประวัติการจัดกลุ่มในระบบ")
    else:
        available_dates = sorted(list(history_data.keys()), reverse=True)
        selected_date = st.selectbox("📅 เลือกวันที่ต้องการดูประวัติ:", available_dates)

        if selected_date and selected_date in history_data:
            record = history_data[selected_date]
            st.caption(f"เวลาประมวลผลล่าสุด: {record.get('timestamp')}")

            h1, h2, h3 = st.columns(3)
            h1.metric("จำนวนรถทั้งหมด", f"{record.get('total_cars')} คัน")
            h2.metric("จัดกลุ่มสำเร็จ", f"{record.get('grouped_cars')} คัน")
            h3.metric("จำนวนกลุ่มสร้างได้", f"{record.get('total_groups')} กลุ่ม")

            df_hist_summary = pd.DataFrame(record.get("summary", []))
            st.dataframe(df_hist_summary, use_container_width=True)


# 6. REVISE & SWAP VIN ON MAIN PANEL
elif active_feature == "✏️ Revise & Swap VIN":
    st.subheader("✏️ แก้ไข / ยกเลิก / สลับคันรถใน Grouping (Revise & Swap)")

    if "df_last_processed" not in st.session_state or "df_last_summary" not in st.session_state:
        st.warning("⚠️ กรุณาประมวลผลจัดกลุ่มที่หัวข้อ '🚀 วางแผนจัดกลุ่ม' ก่อนทำการแก้ไข")
    else:
        df_proc = st.session_state["df_last_processed"].copy()
        df_sum = st.session_state["df_last_summary"].copy()

        st.markdown("#### **1. ยกเลิกกลุ่ม หรือ ปลดบางคันออกจากกลุ่ม (Remove VIN)**")
        group_list = df_sum["Grouping ID"].tolist() if not df_sum.empty else []

        if group_list:
            selected_grp = st.selectbox("เลือก Grouping ID ที่ต้องการ Revise:", group_list)

            grp_vins = df_proc[df_proc["Calc_Group_No"] == selected_grp]
            st.write(f"รายการรถในกลุ่ม **{selected_grp}** ({len(grp_vins)} คัน):")

            show_cols = [c for c in ["Vin", "Model", "Location", "Delivery Location", "Region", "Allocation Date"] if c in df_proc.columns]
            st.dataframe(grp_vins[show_cols], use_container_width=True)

            vins_to_remove = st.multiselect(
                "เลือก VIN ที่ต้องการปลดออกจากกลุ่มนี้ (เพื่อนำกลับไปคิวรอ):",
                grp_vins["Vin"].tolist() if "Vin" in grp_vins.columns else []
            )

            col_rev1, col_rev2 = st.columns(2)
            with col_rev1:
                if st.button("❌ ปลด VIN ที่เลือกออกจากกลุ่ม", type="primary"):
                    if vins_to_remove:
                        df_proc.loc[df_proc["Vin"].isin(vins_to_remove), "Calc_Group_No"] = ""
                        df_proc.loc[df_proc["Vin"].isin(vins_to_remove), "Calc_Group_Date"] = ""
                        st.session_state["df_last_processed"] = df_proc
                        st.success(f"ปลด {len(vins_to_remove)} VIN ออกจากกลุ่ม {selected_grp} เรียบร้อยแล้ว!")
                        st.rerun()

            with col_rev2:
                if st.button("🗑️ ยกเลิกกลุ่มนี้ทั้งหมด (Cancel Whole Group)"):
                    df_proc.loc[df_proc["Calc_Group_No"] == selected_grp, "Calc_Group_No"] = ""
                    df_proc.loc[df_proc["Calc_Group_No"] == selected_grp, "Calc_Group_Date"] = ""
                    st.session_state["df_last_processed"] = df_proc
                    st.success(f"ยกเลิกกลุ่ม {selected_grp} ทั้งหมดเรียบร้อยแล้ว!")
                    st.rerun()

            st.divider()
            st.markdown("#### **2. สลับคันรถระหว่างกลุ่ม (Swope VIN)**")

            s1, s2 = st.columns(2)
            with s1:
                grp_a = st.selectbox("เลือก กลุ่ม A:", group_list, key="grp_a")
                vins_a = df_proc[df_proc["Calc_Group_No"] == grp_a]["Vin"].tolist() if "Vin" in df_proc.columns else []
                vin_a_selected = st.selectbox("เลือก VIN จากกลุ่ม A:", vins_a, key="vin_a")

            with s2:
                grp_b = st.selectbox("เลือก กลุ่ม B:", [g for g in group_list if g != grp_a], key="grp_b")
                vins_b = df_proc[df_proc["Calc_Group_No"] == grp_b]["Vin"].tolist() if "Vin" in df_proc.columns else []
                vin_b_selected = st.selectbox("เลือก VIN จากกลุ่ม B:", vins_b, key="vin_b")

            if st.button("🔄 สลับคันรถ (Swap VIN A ↔ VIN B)"):
                if vin_a_selected and vin_b_selected:
                    df_proc.loc[df_proc["Vin"] == vin_a_selected, "Calc_Group_No"] = grp_b
                    df_proc.loc[df_proc["Vin"] == vin_b_selected, "Calc_Group_No"] = grp_a
                    st.session_state["df_last_processed"] = df_proc
                    st.success(f"สลับ VIN {vin_a_selected} ↔ {vin_b_selected} สำเร็จ!")
                    st.rerun()