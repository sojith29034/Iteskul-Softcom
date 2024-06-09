import streamlit as st
import pandas as pd
from io import BytesIO

# Helper function to calculate attendance percentage
def calculate_attendance(df):
    total_days = len(df.columns) - 1  # Exclude the 'Student Name' column
    df['Attendance %'] = df.iloc[:, 1:].sum(axis=1) / total_days * 100
    return df

# Helper function to find students absent for 3 or more consecutive days
def find_consecutive_absentees(df):
    absentees = []
    for index, row in df.iterrows():
        attendance_series = row[1:-1]
        max_consecutive_absent = (attendance_series == 0).astype(int).groupby(attendance_series.ne(0).cumsum()).sum().max()
        if max_consecutive_absent >= 3:
            absentees.append(row['Student Name'])
    return absentees

# Streamlit app
st.title("Student Attendance Report")

uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    class_reports = {}
    for uploaded_file in uploaded_files:
        class_name = uploaded_file.name.split('.')[0]  # Assuming file name is the class name
        st.write(f"Processing {class_name}...")

        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Calculate attendance percentage
        df = calculate_attendance(df)

        # Find students with attendance below 75%
        low_attendance = df[df['Attendance %'] < 75]['Student Name'].tolist()

        # Find students with 3 or more consecutive absences
        consecutive_absentees = find_consecutive_absentees(df)

        # Store the results in a dictionary
        class_reports[class_name] = {
            "Attendance Data": df,
            "Low Attendance": low_attendance,
            "Consecutive Absentees": consecutive_absentees
        }

    # Display the results
    for class_name, report in class_reports.items():
        st.subheader(f"Class: {class_name}")
        st.write("Attendance Data")
        st.dataframe(report["Attendance Data"])

        st.write("Students with Attendance Below 75%")
        st.write(report["Low Attendance"])

        st.write("Students Absent for 3 or More Consecutive Days")
        st.write(report["Consecutive Absentees"])
