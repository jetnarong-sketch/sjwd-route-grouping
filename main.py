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


def is_car_ready_to_ship(row, hold_col="HOLD", remark_col="Remark"):
    """ข้อ 9: ตรวจสอบความพร้อมในการจัดส่งผ่านช่อง HOLD และ Remark"""
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
    """แปลงข้อความวันที่ YYYY-MM-DD หรือ DD/MM/YYYY ให้กลายเป็น '31 Jul 26'"""
    if not val_str or not isinstance(val_str, str):
        return val_str

    clean_str = val_str.replace(" 00:00:00", "").strip()

    # ลองแปลงรูปแบบ YYYY-MM-DD
    try:
        dt = datetime.strptime(clean_str, "%Y-%m-%d")
        return dt.strftime("%d %b %y")
    except ValueError:
        pass

    # ลองแปลงรูปแบบ DD/MM/YYYY
    try:
        dt = datetime.strptime(clean_str, "%d/%m/%Y")
        return dt.strftime("%d %b %y")
    except ValueError:
        pass

    # ใช้ pandas to_datetime สำรองกรณีติดรูปแบบอื่น
    try:
        dt = pd.to_datetime(clean_str, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d %b %y")
    except Exception:
        pass

    return val_str


def process_fis_grouping_preserve_format(file_bytes, grouping_date_obj):
    grouping_date_str = grouping_date_obj.strftime("%y%m%d")  # รูปแบบ YYMMDD
    grouping_date_display = grouping_date_obj.strftime("%d %b %y")

    # อ่าน DataFrame มาประมวลผล Logic
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

    df[group_no_col] = df[group_no_col].astype(object)
    df[group_date_col] = df[group_date_col].astype(object)

    # 1. กรองเฉพาะรถที่พร้อมส่ง (ข้อ 9)
    df["Ready_Flag"] = df.apply(
        lambda r: is_car_ready_to_ship(r, hold_col, remark_col), axis=1
    )
    ready_df = df[df["Ready_Flag"] == True].copy()

    # 2. ข้อ 8: เรียงลำดับคิวตาม Aging Allocation Date
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

        while len(pending_indices) >= 6:  # ข้อ 3: ต้องมีอย่างน้อย 6 คัน
            group_indices = []
            pickups_in_group = set()
            deliveries_in_group = set()

            for idx in pending_indices:
                curr_pickup = region_batch.loc[idx, pickup_col]
                curr_delivery = region_batch.loc[idx, delivery_col]

                temp_pickups = pickups_in_group | {curr_pickup}
                temp_deliveries = deliveries_in_group | {curr_delivery}

                # ข้อ 5, 6, 7: Pick up <= 3 จุด และ Delivery <= 3 จุด
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

    # โหลดไฟล์เข้า openpyxl เพื่อเขียนผลลัพธ์และแปลง Format วันที่
    wb = openpyxl.load_workbook(file_bytes)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    g_no_col_idx = headers.index(group_no_col) + 1
    g_date_col_idx = headers.index(group_date_col) + 1

    date_header_names = ["Gate In", "Allocation Date", "Grouping Date", "วันที่รับ"]
    target_date_cols = [
        headers.index(h) + 1 for h in date_header_names if h in headers
    ]

    # 1. เขียนข้อมูล Grouping Number & Date
    for idx, row in df.iterrows():
        excel_row_num = idx + 2
        calc_no = row["Calc_Group_No"]
        calc_date = row["Calc_Group_Date"]

        if calc_no != "":
            ws.cell(row=excel_row_num, column=g_no_col_idx, value=calc_no)
            ws.cell(row=excel_row_num, column=g_date_col_idx, value=calc_date)

    # 2. บังคับเปลี่ยนวันที่ในคอลัมน์วันที่ทั้งหมดให้เป็น '31 Jul 26'
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
    return output_buffer, pd.DataFrame(summary_list), total_cars


# --- Streamlit Web App Interface ---
st.set_page_config(
    page_title="SJWD - FIS Auto Grouping System",
    page_icon="🚛",
    layout="wide",
)

st.title("🚛 ระบบจัดกลุ่มรถขนส่งอัตโนมัติ (FIS Delivery Optimization)")
st.caption(
    "ประมวลผลจัดกลุ่มรถขนส่งอัตโนมัติ อ้างอิงเงื่อนไข Aging, สถานะความพร้อม และ Multi-stop Constraint"
)

uploaded_file = st.file_uploader(
    "อัปโหลดไฟล์ FIS Ready to Grouping for delivery (.xlsx)",
    type=["xlsx", "xls"],
)

if uploaded_file:
    file_bytes = io.BytesIO(uploaded_file.getvalue())
    st.success("อัปโหลดไฟล์สำเร็จ!")

    col1, col2 = st.columns(2)
    with col1:
        date_input = st.date_input("เลือกวันที่จัดกลุ่ม", datetime.now())

    with col2:
        st.write("")
        st.write("")
        run_btn = st.button("🚀 ประมวลผลจัดกลุ่มอัตโนมัติ", type="primary")

    if run_btn:
        with st.spinner("กำลังประมวลผลและเปลี่ยน Format วันที่เป็น '31 Jul 26'..."):
            out_buffer, df_summary, total_cars = (
                process_fis_grouping_preserve_format(file_bytes, date_input)
            )

        st.divider()
        st.subheader("📊 สรุปผลการจัดกลุ่มจัดส่ง")

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
        