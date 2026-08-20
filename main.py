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
    page_title="SIAM JWD LOGISTICS - TMS Dashboard",
    page_icon="🚛",
    layout="wide",
)

st.markdown(
    """
    <style>
    /* Clean Light Sidebar */
    [data-testid="stSidebar"] {
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0 !important;
    }
    [data-testid="stSidebar"] * {
        color: #1e293b !important;
    }
    
    /* Active Menu Radio Style */
    div[data-testid="stSidebar"] div[role="radiogroup"] > label {
        padding: 10px 14px !important;
        border-radius: 8px !important;
        margin-bottom: 4px !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
        background-color: #e2e8f0 !important;
    }
    
    /* Primary Red Buttons */
    div.stButton > button:first-child {
        background-color: #2563eb !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 8px 18px !important;
    }
    
    /* Action Primary Upload Button */
    button[kind="primary"] {
        background-color: #0066B3 !important;
    }
    
    .clean-table-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 24px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-top: 15px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# SIDEBAR BRANDING & CLEAN MENU
st.sidebar.markdown(
    """
    <div style="padding: 10px 0px 15px 0px; border-bottom: 1px solid #e2e8f0; margin-bottom: 15px;">
        <span style="color:#ED1C24; font-size:24px; font-weight:900; font-family:sans-serif;">SIAM </span>
        <span style="color:#0066B3; font-size:24px; font-weight:900; font-family:sans-serif;">JWD</span><br>
        <span style="color:#64748b; font-size:11px; letter-spacing:3px; font-weight:bold;">LOGISTICS</span>
    </div>
    """,
    unsafe_allow_html=True,
)

menu_choice = st.sidebar.radio(
    "เมนูหลัก (Navigation)",
    [
        "🗺️ วางแผนการเดินทาง",
        "📦 จัดการคำสั่งซื้อ",
        "🚛 Fleet Capacity",
        "📜 ประวัติการจัดกลุ่ม",
        "✏️ แก้ไข / สลับ VIN",
    ],
    index=0,
)

st.sidebar.divider()
st.sidebar.caption("👤 Admin Ball (Administrator)")
st.sidebar.caption("SIAM JWD LOGISTICS CO., LTD.")

# TOP BAR HEADER
col_title, col_actions = st.columns([2, 1])
with col_title:
    st.markdown(f"## **{menu_choice.split(' ')[1]}**")
    st.caption("จัดกลุ่มรถเพื่อจัดส่งและมอบหมายให้คนขับ (Automated Optimization System)")

with col_actions:
    st.write("")
    if st.button("⬆️ อัปโหลดการจัดกลุ่ม (Upload FIS)", type="primary", use_container_width=True):
        st.session_state["show_upload_dialog"] = True

# --- POPUP UPLOAD DIALOG / EXPANDABLE ---
if st.session_state.get("show_upload_dialog", False):
    with st.expander("📂 **กล่องนำเข้าไฟล์สำหรับการจัดกลุ่ม (Upload Files)**", expanded=True):
        u1, u2 = st.columns(2)
        with u1:
            st.markdown("##### **1. Master list**")
            master_region_file = st.file_uploader(
                "อัปโหลดไฟล์ Dealer (Region).xlsx", type=["xlsx", "xls"], key="dlg_master"
            )
        with u2:
            st.markdown("##### **2. Grouping order**")
            uploaded_file = st.file_uploader(
                "อัปโหลดไฟล์ FIS Ready to Grouping (.xlsx)", type=["xlsx", "xls"], key="dlg_fis"
            )

        col_run, col_close = st.columns([3, 1])
        with col_run:
            if uploaded_file and master_region_file:
                if st.button("🚀 ประมวลผลจัดกลุ่มทันที", use_container_width=True):
                    file_bytes = io.BytesIO(uploaded_file.getvalue())
                    master_df = pd.read_excel(master_region_file)

                    capacity_settings = {
                        "trailer_7": st.session_state.get("trailer_7_qty", 20),
                        "trailer_8": st.session_state.get("trailer_8_qty", 5),
                        "slide_on": st.session_state.get("slide_on_allow", True),
                    }

                    with st.spinner("กำลังคำนวณและจัดกลุ่มรถอัตโนมัติ..."):
                        out_buffer, df_summary, total_cars, missing_locs, df_processed = (
                            process_fis_grouping_with_capacity(
                                file_bytes, master_df, datetime.now(), capacity_settings
                            )
                        )

                    if missing_locs:
                        st.error("❌ ไม่พบ Delivery Location ใน Master!")
                        for m_loc in missing_locs:
                            st.write(f"- {m_loc}")
                    else:
                        st.session_state["df_last_processed"] = df_processed
                        st.session_state["df_last_summary"] = df_summary
                        st.session_state["out_buffer"] = out_buffer
                        st.session_state["total_cars"] = total_cars

                        # Save History
                        history = load_history()
                        date_key = datetime.now().strftime("%Y-%m-%d")
                        history[date_key] = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "total_cars": total_cars,
                            "grouped_cars": int(df_summary["Car Count"].sum() if not df_summary.empty else 0),
                            "total_groups": len(df_summary),
                            "summary": df_summary.to_dict(orient="records"),
                        }
                        save_history(history)
                        st.session_state["show_upload_dialog"] = False
                        st.rerun()

        with col_close:
            if st.button("ปิดหน้าต่าง", use_container_width=True):
                st.session_state["show_upload_dialog"] = False
                st.rerun()

st.divider()

# --- MODULE 1: MAIN TRIP PLANNING WORKSPACE ---
if menu_choice == "🗺️ วางแผนการเดินทาง":
    # FILTER BAR
    f1, f2, f3, f4 = st.columns([1, 2, 2, 1])
    with f1:
        st.selectbox("ประเภททริป", ["ทั้งหมด", "Trailer", "Slide-on"])
    with f2:
        st.text_input("🔍 ค้นหาด้วยเลขทริป / Dealer / VIN...", placeholder="พิมพ์คำค้นหา...")
    with f3:
        st.date_input("วันที่วางแผน", datetime.now())
    with f4:
        st.selectbox("สถานะ", ["ทุกสถานะ", "จัดกลุ่มแล้ว", "รอจัดกลุ่ม"])

    st.write("")

    # RESULTS DATA DISPLAY
    if "df_last_summary" in st.session_state:
        df_summary = st.session_state["df_last_summary"]
        total_cars = st.session_state.get("total_cars", 0)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("จำนวนทริปทั้งหมด", f"{len(df_summary)} ทริป")
        grouped_cars_count = df_summary["Car Count"].sum() if not df_summary.empty else 0
        m2.metric("รถที่จัดทริปสำเร็จ", f"{grouped_cars_count} คัน")
        m3.metric("รถค้างส่ง/รอทริป", f"{total_cars - grouped_cars_count} คัน")
        
        with m4:
            if "out_buffer" in st.session_state:
                st.download_button(
                    label="📥 ส่งออกการจัดกลุ่ม (.xlsx)",
                    data=st.session_state["out_buffer"],
                    file_name=f"FIS_Grouped_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        st.markdown("### **รายการทริปเดินทางที่คำนวณสำเร็จ**")
        st.dataframe(df_summary, use_container_width=True)

    else:
        st.markdown(
            """
            <div style="text-align: center; padding: 60px 20px; background-color: #ffffff; border-radius: 12px; border: 1px dashed #cbd5e1;">
                <p style="font-size: 40px; margin-bottom: 10px;">📋</p>
                <h4 style="color: #475569; margin: 0;">ไม่พบทริปการเดินทาง</h4>
                <p style="color: #94a3b8; font-size: 14px; margin-top: 5px;">กรุณากดปุ่ม <b>'⬆️ อัปโหลดการจัดกลุ่ม'</b> ที่มุมขวาบนเพื่อนำเข้าไฟล์และสร้างทริปอัตโนมัติ</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

# --- MODULE 2: ORDER MANAGEMENT ---
elif menu_choice == "📦 จัดการคำสั่งซื้อ":
    st.subheader("📦 รายการคำสั่งซื้อและสถานะคิวรถ (Order Queue)")
    if "df_last_processed" in st.session_state:
        df_proc = st.session_state["df_last_processed"]
        show_cols = [c for c in ["Vin", "Model", "Location", "Delivery Location", "Region", "Allocation Date", "Calc_Group_No", "Ready_Flag"] if c in df_proc.columns]
        st.dataframe(df_proc[show_cols], use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูลคำสั่งซื้อในระบบ")

# --- MODULE 3: FLEET CAPACITY ---
elif menu_choice == "🚛 Fleet Capacity":
    st.subheader("🚛 ตั้งค่ากองรถและโควตาเทรลเลอร์ (Fleet Capacity)")
    fc1, fc2 = st.columns(2)
    with fc1:
        st.session_state["trailer_7_qty"] = st.number_input(
            "จำนวน Trailer (6-7 Load) พร้อมวิ่ง:", min_value=0, value=st.session_state.get("trailer_7_qty", 20)
        )
        st.session_state["trailer_8_qty"] = st.number_input(
            "จำนวน Trailer (8 Load) พร้อมวิ่ง:", min_value=0, value=st.session_state.get("trailer_8_qty", 5)
        )
    with fc2:
        st.session_state["slide_on_allow"] = st.checkbox(
            "อนุญาตให้ส่งแบบ Slide-on สำหรับ DENZA D9 ใน BKK ได้ทันที", value=st.session_state.get("slide_on_allow", True)
        )
    st.success("💾 บันทึกการตั้งค่า Fleet Capacity เรียบร้อยแล้ว")

# --- MODULE 4: HISTORY ---
elif menu_choice == "📜 ประวัติการจัดกลุ่ม":
    st.subheader("📜 เรียกดูประวัติการจัดกลุ่มย้อนหลัง")
    history_data = load_history()
    if not history_data:
        st.info("ยังไม่มีประวัติในระบบ")
    else:
        selected_date = st.selectbox("📅 เลือกวันที่:", sorted(list(history_data.keys()), reverse=True))
        if selected_date in history_data:
            rec = history_data[selected_date]
            st.dataframe(pd.DataFrame(rec.get("summary", [])), use_container_width=True)

# --- MODULE 5: REVISE ---
elif menu_choice == "✏️ แก้ไข / สลับ VIN":
    st.subheader("✏️ แก้ไข / ยกเลิก / สลับคันรถใน Grouping")
    if "df_last_processed" in st.session_state and "df_last_summary" in st.session_state:
        df_proc = st.session_state["df_last_processed"].copy()
        df_sum = st.session_state["df_last_summary"].copy()
        group_list = df_sum["Grouping ID"].tolist() if not df_sum.empty else []

        selected_grp = st.selectbox("เลือก Grouping ID ที่ต้องการ Revise:", group_list)
        grp_vins = df_proc[df_proc["Calc_Group_No"] == selected_grp]
        st.dataframe(grp_vins, use_container_width=True)
    else:
        st.warning("⚠️ กรุณาประมวลผลจัดกลุ่มก่อนทำการแก้ไข")