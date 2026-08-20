from datetime import datetime
import io
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
    # 9.1 ตรวจสอบช่อง HOLD (ต้องเป็นค่าว่าง/Blank เท่านั้น)
    if pd.notna(row[hold_col]):
        hold_val = str(row[hold_col]).strip()
        if hold_val != "":
            return False

    # 9.2 ตรวจสอบช่อง Remark (ต้องไม่มีคำสั่ง Hold หรือเลื่อนการส่ง)
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


def process_fis_grouping(df, grouping_date_obj):
    grouping_date_str = grouping_date_obj.strftime("%y%m%d")  # รูปแบบ YYMMDD
    grouping_date_display = grouping_date_obj.strftime(
        "%d/%m/%Y"
    )  # สำหรับลงในช่อง Grouping Date

    pickup_col = "Pick up Location"
    delivery_col = "Delivery Location"
    region_col = "Region"
    model_col = "Model"
    group_no_col = "Grouping number"
    group_date_col = "Grouping Date"
    hold_col = "HOLD"
    remark_col = "Remark"
    alloc_date_col = "Allocation Date"

    # เปลี่ยนคำนำหน้าเป็น SJWD
    prefix = f"SJWD{grouping_date_str}-"
    group_counter = 1

    # 1. กรองเฉพาะรถที่พร้อมส่ง (ข้อ 9)
    df["Ready_Flag"] = df.apply(
        lambda r: is_car_ready_to_ship(r, hold_col, remark_col), axis=1
    )
    ready_df = df[df["Ready_Flag"] == True].copy()

    # 2. ข้อ 8: เรียงลำดับคิวตาม Aging Allocation Date (เก่าสุดขึ้นก่อน)
    if alloc_date_col in ready_df.columns:
        ready_df[alloc_date_col] = pd.to_datetime(
            ready_df[alloc_date_col], errors="coerce"
        )
        ready_df = ready_df.sort_values(by=alloc_date_col, ascending=True)

    # คำนวณน้ำหนักประเมิน
    ready_df["Estimated_Weight_KG"] = (
        ready_df[model_col]
        .astype(str)
        .str.upper()
        .map(lambda x: MODEL_WEIGHT_MASTER.get(x, 1800))
    )

    df["Calc_Group_No"] = ""
    df["Calc_Group_Date"] = ""
    summary_list = []

    # แบ่งกลุ่มตาม Region เพื่อตีกรอบพื้นที่จัดส่ง
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

                # ข้อ 3: สูงสุด 8 คันต่อกลุ่ม
                if len(group_indices) == 8:
                    break

            # ข้อ 3: ตรวจสอบว่ารวมได้ระหว่าง 6 ถึง 8 คันหรือไม่
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

    # ข้อ 4: อัปเดตคอลัมน์ (คันที่จัดไม่ได้ ให้คงค่าเดิมไว้)
    mask = df["Calc_Group_No"] != ""
    df.loc[mask, group_no_col] = df.loc[mask, "Calc_Group_No"]
    df.loc[mask, group_date_col] = df.loc[mask, "Calc_Group_Date"]

    df.drop(
        columns=["Ready_Flag", "Calc_Group_No", "Calc_Group_Date"],
        inplace=True,
        errors="ignore",
    )

    return df, pd.DataFrame(summary_list)


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
    df_raw = pd.read_excel(uploaded_file)
    st.success(f"อัปโหลดไฟล์สำเร็จ! จำนวนรถทั้งหมดในไฟล์: {len(df_raw)} คัน")

    col1, col2 = st.columns(2)
    with col1:
        date_input = st.date_input("เลือกวันที่จัดกลุ่ม", datetime.now())

    with col2:
        st.write("")
        st.write("")
        run_btn = st.button("🚀 ประมวลผลจัดกลุ่มอัตโนมัติ", type="primary")

    if run_btn:
        with st.spinner("กำลังตรวจสอบเงื่อนไขและจัดกลุ่มคิวขนส่ง..."):
            df_result, df_summary = process_fis_grouping(df_raw, date_input)

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
            f"{len(df_raw) - grouped_cars_count} คัน",
        )

        if not df_summary.empty:
            st.dataframe(df_summary, use_container_width=True)
        else:
            st.warning(
                "ไม่พบคันรถที่ตรงตามเงื่อนไขครบ 6-8 คัน หรือรถส่วนใหญ่อยู่ในสถานะ HOLD/เลื่อนส่ง"
            )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์",
            data=buffer,
            file_name=f"FIS_Grouped_{date_input.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        