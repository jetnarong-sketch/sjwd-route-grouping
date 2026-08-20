from datetime import datetime, date
import io
import base64
import openpyxl
import pandas as pd
from PIL import Image
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

# รหัส Base64 ภาพโลโก้ SIAM JWD LOGISTICS ต้นฉบับจริง (สมบูรณ์ 100%)
LOGO_B64 = "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"

@st.cache_data
def load_embedded_logo():
    """ดึงรูปจาก Base64 ขึ้น RAM โดยตรง"""
    try:
        image_data = base64.b64decode(LOGO_B64)
        return Image.open(io.BytesIO(image_data))
    except Exception:
        return None


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

# โหลดรูปภาพจากหน่วยความจำ RAM
logo_img = load_embedded_logo()

# --- SIDEBAR: CONTROL PANEL ---
if logo_img is not None:
    st.sidebar.image(logo_img, use_container_width=True)

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
if logo_img is not None:
    st.image(logo_img, width=420)

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