import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.title("MKU Exam Extractor")
st.subheader("Only works for Undergraduate (Bachelor's and Diploma) exam timetables. File uploaded must be in Excel format.")
st.write("Upload your exam timetable Excel file and select your units.")

uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx"])

def clean_date(date_val):
    """Standardises date strings by removing day names"""
    if pd.isna(date_val):
        return None
    
    d_str = str(date_val).strip()
    
    if re.match(r'^\d{4}-\d{2}-\d{2}.*', d_str):
        return pd.to_datetime(d_str).date()
        
    d_str = re.sub(r'^(MON|TUE|WED|THU|THUR|FRI|SAT|SUN)\s+', '', d_str, flags=re.IGNORECASE)
    
    d_str = re.sub(r'(?<=\d)(st|nd|rd|th)', '', d_str, flags=re.IGNORECASE)
    
    # - cleaned string parse
    try:
        return pd.to_datetime(d_str).date()
    except:
        return None

def load_data(file):
    """loads the dataframe by detecting headers and standardising columns."""
    
    preview = pd.read_excel(file, header=None, nrows=10)
    header_row = 0
    
    for i, row in preview.iterrows():
        row_str = row.astype(str).str.upper().values
        
        if any("CODE" in x for x in row_str) and (any("TIME" in x for x in row_str) or any("DATE" in x for x in row_str)):
            header_row = i
            break
            
    df = pd.read_excel(file, header=header_row)
    
    df.columns = df.columns.str.strip().str.upper()
    
    col_map = {
        'UNIT CODE': 'COURSE_CODE',
        'CODE': 'COURSE_CODE',
        'UNIT CODE & NAME': 'COURSE_CODE', 
        'COURSE_CODE': 'COURSE_CODE',
        'UNIT NAME': 'COURSE_TITLE',
        'COURSE_TITLE': 'COURSE_TITLE',
        'DAY & DATE': 'EXAMS DATE',
        'DAY/DATE': 'EXAMS DATE',
        'EXAMS DATE': 'EXAMS DATE',
        'TIME': 'SESSION TIME',
        'SESSION TIME': 'SESSION TIME'
    }
    
    df.rename(columns=col_map, inplace=True)
    
    # TODO: Handle Merged Cells for dates & time that are merged vertically
    cols_to_fill = ['EXAMS DATE', 'SESSION TIME']
    for col in cols_to_fill:
        if col in df.columns:
            df[col] = df[col].ffill()

    if 'COURSE_CODE' in df.columns:
        code_pattern = r'([A-Z]{2,4}\s?[-]?\s?\d{3,4}[A-Z]?)'
        
        if 'COURSE_TITLE' not in df.columns:
             df['COURSE_TITLE'] = df['COURSE_CODE'].astype(str).str.replace(code_pattern, '', regex=True).str.strip(' -')
             
        df['COURSE_CODE'] = df['COURSE_CODE'].astype(str).str.extract(code_pattern, expand=False)
        
    return df

if uploaded_file is not None:
    df = load_data(uploaded_file)
    if 'COURSE_CODE' not in df.columns:
        st.error("Could not detect 'Unit Code' column. Please check the file format.")
    else:
        df['COURSE_CODE'] = df['COURSE_CODE'].str.upper().str.strip()
        available_units = df["COURSE_CODE"].dropna().unique().tolist()
        available_units.sort()

        selected_units = st.multiselect("Select your units", options=available_units)

        if selected_units:
            filtered = df[df["COURSE_CODE"].isin(selected_units)].copy()
            if 'EXAMS DATE' in filtered.columns:
                filtered['EXAMS DATE'] = filtered['EXAMS DATE'].apply(clean_date)
                
            if 'SESSION TIME' in filtered.columns:
                 filtered = filtered.sort_values(by=["EXAMS DATE", "SESSION TIME"])
            
            cols_to_show = [c for c in ["EXAMS DATE", "SESSION TIME", "COURSE_CODE", "COURSE_TITLE", "VENUE"] if c in filtered.columns]
            st.subheader("Filtered Exam Timetable")
            st.dataframe(filtered[cols_to_show])

            # Download button
            buffer = BytesIO()
            filtered[cols_to_show].to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            st.download_button(
                "Download as Excel",
                data=buffer,
                file_name="filtered_exam_timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Please select at least one unit.")