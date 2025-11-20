import streamlit as st
import pandas as pd
from io import BytesIO

st.title("MKU Exam Extractor")
st.write("Upload your exam timetable Excel file and select your units.")

# File uploader
uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx"])

def filter_my_units(file, selected_units):
    df = pd.read_excel(file, header=5)
    df.columns = df.columns.str.strip().str.upper()

    selected_units = [u.upper() for u in selected_units]

    filtered = df[df["COURSE_CODE"].isin(selected_units)]

    clean = filtered[["EXAMS DATE", "SESSION TIME", "COURSE_CODE", "COURSE_TITLE"]].copy()

    is_date = clean["EXAMS DATE"].astype(str).str.match(r'^\d{4}-\d{2}-\d{2}$')
    clean.loc[is_date, "EXAMS DATE"] = pd.to_datetime(
        clean.loc[is_date, "EXAMS DATE"], errors='coerce'
    ).dt.date

    clean = clean.sort_values(by=["EXAMS DATE", "SESSION TIME"]).reset_index(drop=True)

    return clean

if uploaded_file is not None:
    df_preview = pd.read_excel(uploaded_file, header=5)
    df_preview.columns = df_preview.columns.str.strip().str.upper()
    available_units = df_preview["COURSE_CODE"].dropna().unique().tolist()
    available_units.sort()

    # Multi select for units
    selected_units = st.multiselect(
        "Select your units", options=available_units
    )

    if selected_units:
        result = filter_my_units(uploaded_file, selected_units)
        st.subheader("Filtered Exam Timetable")
        st.dataframe(result)

        buffer = BytesIO()
        result.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            "Download as Excel",
            data=buffer,
            file_name="filtered_exam_timetable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Please select at least one unit.")
