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
}


def find_column_name(df, keywords):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for kw in keywords:
            if kw.lower() in col_clean:
                return col
    return None


def process_grouping(df, grouping_date_str):
    pickup_col = (
        find_column_name(df, ["location pick up", "pickup", "location"])
        or df.columns[12]
    )
    delivery_col = (
        find_column_name(df, ["location delivery", "delivery location", "delivery"])
        or df.columns[13]
    )
    region_col = find_column_name(df, ["region"]) or df.columns[14]
    model_col = find_column_name(df, ["model"]) or df.columns[8]
    group_col = (
        find_column_name(
            df, ["groupping  number", "grouping number", "grouping"]
        )
        or "Groupping  Number"
    )

    prefix = f"SJWD{grouping_date_str}-"
    group_counter = 1

    df["Estimated_Weight_KG"] = (
        df[model_col]
        .astype(str)
        .str.upper()
        .map(lambda x: MODEL_WEIGHT_MASTER.get(x, 1800))
    )
    df["Calculated_Grouping_Number"] = ""

    grouped_batches = df.groupby(
        [pickup_col, delivery_col, region_col], dropna=False
    )

    summary_list = []

    for (pickup, delivery, region), batch in grouped_batches:
        indices = batch.index.tolist()
        i = 0
        total_cars = len(indices)

        while i < total_cars:
            current_group_id = f"{prefix}{group_counter:03d}"

            if (total_cars - i) == 1:
                batch_size = 1
            elif (total_cars - i) <= 7:
                batch_size = min(7, total_cars - i)
            else:
                batch_size = 8

            selected_indices = indices[i : i + batch_size]
            group_weight = df.loc[
                selected_indices, "Estimated_Weight_KG"
            ].sum()

            df.loc[selected_indices, "Calculated_Grouping_Number"] = (
                current_group_id
            )

            summary_list.append(
                {
                    "Grouping ID": current_group_id,
                    "Pickup": pickup,
                    "Delivery Location": delivery,
                    "Region": region,
                    "Car Count": len(selected_indices),
                    "Total Weight (kg)": group_weight,
                }
            )

            group_counter += 1
            i += batch_size

    df[group_col] = df["Calculated_Grouping_Number"]
    df.drop(
        columns=["Estimated_Weight_KG", "Calculated_Grouping_Number"],
        inplace=True,
    )

    return df, pd.DataFrame(summary_list)


# --- ส่วนอินเทอร์เฟซหน้าเว็บ (Streamlit App) ---
st.set_page_config(
    page_title="SJWD Route & Vehicle Grouping System", layout="wide"
)

st.title("🚛 ระบบจัดกลุ่มรถและวางแผนเส้นทางขนส่งอัตโนมัติ")
st.write(
    "อัปโหลดไฟล์รายงาน Excel เพื่อประมวลผลจัดกลุ่มรหัสขนส่ง (Grouping Number) อัตโนมัติ"
)

uploaded_file = st.file_uploader(
    "เลือกหรือลากไฟล์ Excel (.xlsx) มาวางที่นี่", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    st.success(f"อัปโหลดไฟล์สำเร็จ! จำนวนรายการทั้งหมด: {len(df_raw)} คัน")

    col1, col2 = st.columns(2)
    with col1:
        date_input = st.date_input("เลือกวันที่จัดกลุ่ม", datetime.now())
        grouping_date_str = date_input.strftime("%Y%m%d")

    with col2:
        st.write("")
        st.write("")
        run_btn = st.button("🚀 ประมวลผลจัดกลุ่ม", type="primary")

    if run_btn:
        with st.spinner("กำลังประมวลผลจัดกลุ่ม..."):
            df_result, df_summary = process_grouping(
                df_raw, grouping_date_str
            )

        st.divider()
        st.subheader("📊 สรุปผลการจัดกลุ่มขนส่ง")

        # Metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("จำนวนกลุ่มทั้งหมด (Groups)", len(df_summary))
        m2.metric("จำนวนรถทั้งหมด", len(df_result))
        m3.metric(
            "น้ำหนักเฉลี่ยต่อกลุ่ม (กก.)",
            f"{df_summary['Total Weight (kg)'].mean():,.2f}",
        )

        st.dataframe(df_summary, use_container_width=True)

        # เตรียมไฟล์ Download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel ผลลัพธ์",
            data=buffer,
            file_name=f"BYD_Delivery_Grouped_{grouping_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )