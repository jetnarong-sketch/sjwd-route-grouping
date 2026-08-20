from datetime import datetime, date
import io
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

# โค้ด SVG โลโก้ SIAM JWD LOGISTICS ชัดเจน ไม่ต้องโหลดไฟล์ แสดงผล 100%
LOGO_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 520 120" width="100%" height="auto">
  <g>
    <!-- Gear Left Red/Blue Icon -->
    <path d="M 68 15 C 64 15 52 25 42 32 L 25 45 C 22 47 18 53 18 58 C 18 63 22 69 25 71 L 42 84 C 52 91 64 101 68 101 C 72 101 84 91 94 84 L 111 71 C 114 69 118 63 118 58 C 118 53 114 47 111 45 L 94 32 C 84 25 72 15 68 15 Z" fill="none"/>
    <!-- Red Top Half Gear -->
    <path d="M 65 10 C 50 10 32 20 22 30 L 10 40 C 5 45 2 52 2 58 L 75 58 C 75 42 62 30 45 30 L 35 30 C 50 18 80 18 95 30 L 85 30 C 68 30 55 42 55 58 L 132 58 C 132 52 129 45 124 40 L 112 30 C 102 20 84 10 65 10 Z" fill="#ED1C24"/>
    <!-- Blue Bottom Half Gear -->
    <path d="M 65 106 C 80 106 98 96 108 86 L 120 76 C 125 71 128 64 128 58 L 55 58 C 55 74 68 86 85 86 L 95 86 C 80 98 50 98 35 86 L 45 86 C 62 86 75 74 75 58 L 2 58 C 2 64 5 71 10 76 L 22 86 C 32 96 50 106 65 106 Z" fill="#0066B3"/>
    <!-- White S Cutout inside Gear -->
    <path d="M 85 38 L 45 38 C 35 38 28 44 28 52 C 28 58 35 62 48 62 L 82 62 C 95 62 102 66 102 72 C 102 80 95 86 85 86 L 45 86" fill="none" stroke="#FFFFFF" stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>
    
    <!-- Text: SIAM JWD -->
    <text x="145" y="62" font-family="'Arial Black', 'Arial', sans-serif" font-weight="900" font-style="italic" font-size="44" fill="#000000" letter-spacing="-1">SIAM JWD</text>
    
    <!-- Text: LOGISTICS -->
    <text x="185" y="94" font-family="'Arial', sans-serif" font-weight="700" font-size="22" fill="#000000" letter-spacing="9">LOGISTICS</text>
  </g>
</svg>
"""


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


def convert_string_to_dd_mmm_yy(val_str):
    if not val_str or not isinstance(val_str, str):
        return val_str

    clean_str = val_str.replace(" 00:00:00", "").strip()

    try:
        dt = datetime.strptime(clean_str, "%Y-%m-%d")
        return dt.strftime("%d %b %y")
    except ValueError:
        pass

    try:
        dt = datetime.strptime(clean_str, "%d/%m/%Y")
        return dt.strftime("%d %b %y")
    except ValueError:
        pass

    try:
        dt = pd.to_datetime(clean_str, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d %b %y")
    except Exception:
        pass

    return val_str


def process_fis_grouping_preserve_format(
    file_bytes, master_region_df, grouping_date_obj
):
    grouping_date_str = grouping_date_obj.strftime("%y%m%d")
    grouping_date_display = grouping_date_obj.strftime("%d %b %y")

    df = pd.read_excel(file_bytes)

    pickup_col = "Pick up Location"
    delivery_col = "Delivery Location"
    region_col = "Region"
    model_col = "Model"
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
        return None, None, None, missing_locations

    df["Mapped_Region"] = df_delivery_clean.map(master_map)
    df[region_col] = df[region_col].fillna(df["Mapped_Region"])

    df[group_no_col] = df[group_no_col].astype(object)
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

    for region_name, region_batch in ready_df.groupby(region_col, dropna=False):
        pending_indices = region_batch.index.tolist()

        while len(pending_indices) >= 6:
            group_indices = []
            pickups_in_group = set()
            deliveries_in_group = set()

            for idx in pending_indices:
                curr_pickup = region_batch.loc[idx, pickup_col]
                curr_delivery = region_batch.loc[idx, delivery_col]

                temp_pickups = pickups_in_group | {curr_pickup}
                temp_deliveries = deliveries_in_group | {curr_delivery}

                if len(temp_pickups) <= 3 and len(temp_deliveries) <= 3:
                    group_indices.append(idx)
                    pickups_in_group = temp_pickups
                    deliveries_in_group = temp_deliveries

                if len(group_indices) == 8:
                    break

            if len(group_indices) >= 6:
                current_group_id = f"{prefix}{group_counter:03d}"
                df.loc[group_indices, "Calc_Group_No"] = current_group_id
                df.loc[group_indices, "Calc_Group_Date"] = (
                    grouping_date_display
                )

                group_weight = ready_df.loc[
                    group_indices, "Estimated_Weight_KG"
                ].sum()

                summary_list.append(
                    {
                        "Grouping ID": current_group_id,
                        "Region": region_name,
                        "Pick up Locations": ", ".join(
                            map(str, pickups_in_group)
                        ),
                        "Delivery Locations": ", ".join(
                            map(str, deliveries_in_group)
                        ),
                        "Car Count": len(group_indices),
                        "Total Weight (kg)": group_weight,
                    }
                )

                group_counter += 1
                for g_idx in group_indices:
                    pending_indices.remove(g_idx)
            else:
                break

    wb = openpyxl.load_workbook(file_bytes)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    g_no_col_idx = headers.index(group_no_col) + 1
    g_date_col_idx = headers.index(group_date_col) + 1
    region_col_idx = headers.index(region_col) + 1

    date_header_names = ["Gate In", "Allocation Date", "Grouping Date", "วันที่รับ"]
    target_date_cols = [
        headers.index(h) + 1 for h in date_header_names if h in headers
    ]

    for idx, row in df.iterrows():
        excel_row_num = idx + 2
        calc_no = row["Calc_Group_No"]
        calc_date = row["Calc_Group_Date"]
        region_val = row[region_col]

        ws.cell(row=excel_row_num, column=region_col_idx, value=region_val)

        if calc_no != "":
            ws.cell(row=excel_row_num, column=g_no_col_idx, value=calc_no)
            ws.cell(row=excel_row_num, column=g_date_col_idx, value=calc_date)

    for r in range(2, ws.max_row + 1):
        for c in target_date_cols:
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                if isinstance(cell.value, (datetime, date)):
                    cell.value = cell.value.strftime("%d %b %y")
                else:
                    cell.value = convert_string_to_dd_mmm_yy(str(cell.value))
                cell.number_format = "@"

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    total_cars = len(df)
    return output_buffer, pd.DataFrame(summary_list), total_cars, []


# --- Streamlit Layout Configuration ---
st.set_page_config(
    page_title="SIAM JWD LOGISTICS - Auto Grouping System",
    page_icon="🚛",
    layout="wide",
)

# --- SIDEBAR: CONTROL PANEL ---
st.sidebar.markdown(
    f'<div style="max-width:280px; margin-bottom:10px;">{LOGO_SVG}</div>',
    unsafe_allow_html=True,
)

st.sidebar.title("⚙️ Control Panel")
st.sidebar.caption("ศูนย์จัดการไฟล์และตั้งค่าการประมวลผล")

st.sidebar.subheader("1. Master list")
master_region_file = st.sidebar.file_uploader(
    "📂 Upload Dealer (Region).xlsx",
    type=["xlsx", "xls"],
    help="ไฟล์ Master แมปสถานที่ส่งกับ Region",
)

st.sidebar.subheader("2. Grouping order")
uploaded_file = st.sidebar.file_uploader(
    "📁 Upload FIS Ready to Grouping (.xlsx)",
    type=["xlsx", "xls"],
    help="ไฟล์รายการรถที่ต้องการนำมาจัดกลุ่ม",
)

date_input = datetime.now()

run_btn = False
if uploaded_file and master_region_file:
    st.sidebar.write("")
    run_btn = st.sidebar.button(
        "🚀 ประมวลผลจัดกลุ่มอัตโนมัติ", type="primary", use_container_width=True
    )

st.sidebar.divider()
st.sidebar.caption("SIAM JWD LOGISTICS CO., LTD.")


# --- MAIN PANEL ---
st.markdown(
    f'<div style="max-width:440px; margin-bottom:15px;">{LOGO_SVG}</div>',
    unsafe_allow_html=True,
)

st.markdown("### **Auto Fleet Grouping & Logistics Optimization System**")
st.caption(
    "ระบบคำนวณและวางแผนจัดกลุ่มรถขนส่งสินค้าอัตโนมัติ (Automated Car Carrier Optimization)"
)

if not run_btn:
    st.divider()
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        st.markdown("### 🎯 Auto Matching")
        st.write(
            "จับคู่ Delivery Location กับ Region และเติมช่องตกหล่นจาก Master อัตโนมัติ"
        )
    with col_f2:
        st.markdown("### ⏳ Aging Priority")
        st.write(
            "เรียงลำดับคิวรถตาม Allocation Date จากอดีตไปหาปัจจุบัน ป้องกันสินค้าค้างส่ง"
        )
    with col_f3:
        st.markdown("### 🚛 Route Control")
        st.write(
            "คุมจำนวนรถ 6-8 คันต่อเทรลเลอร์ และจุดรับ-ส่งไม่เกินอย่างละ 3 จุดต่อเที่ยววิ่ง"
        )

    st.info(
        "👈 **เริ่มต้นใช้งาน:** กรุณาอัปโหลดไฟล์ Master list และ Grouping order ที่แถบซ้ายมือ (Control Panel)"
    )

else:
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    master_df = pd.read_excel(master_region_file)

    with st.spinner("กำลังตรวจสอบ Master Region และประมวลผลคิวขนส่ง..."):
        out_buffer, df_summary, total_cars, missing_locs = (
            process_fis_grouping_preserve_format(
                file_bytes, master_df, date_input
            )
        )

    if missing_locs:
        st.error(
            "❌ ไม่สามารถประมวลผลได้ เนื่องจากพบ Delivery Location ที่ไม่มีในไฟล์ Master!"
        )
        st.warning(
            "กรุณาเพิ่มข้อมูล Delivery Location ดังต่อไปนี้ลงในไฟล์ Master list (Dealer (Region).xlsx) ก่อนประมวลผลใหม่:"
        )
        for m_loc in missing_locs:
            st.write(f"- 📍 **{m_loc}**")
    else:
        st.divider()
        st.subheader("📊 สรุปผลการจัดกลุ่มจัดส่ง (SIAM JWD LOGISTICS)")

        m1, m2, m3 = st.columns(3)
        m1.metric("จำนวนกลุ่มที่สร้างได้", f"{len(df_summary)} กลุ่ม")
        grouped_cars_count = (
            df_summary["Car Count"].sum() if not df_summary.empty else 0
        )
        m2.metric("จำนวนรถที่จัดกลุ่มสำเร็จ", f"{grouped_cars_count} คัน")
        m3.metric(
            "รถที่ไม่เข้าเงื่อนไข/รอจัดกลุ่มใหม่",
            f"{total_cars - grouped_cars_count} คัน",
        )

        if not df_summary.empty:
            st.dataframe(df_summary, use_container_width=True)
        else:
            st.warning(
                "ไม่พบคันรถที่ตรงตามเงื่อนไขครบ 6-8 คัน หรือรถส่วนใหญ่อยู่ในสถานะ HOLD/เลื่อนส่ง"
            )

        st.download_button(
            label="📥 Download Result grouping",
            data=out_buffer,
            file_name=f"FIS_Grouped_{date_input.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )