import streamlit as st
import pandas as pd
import re

# Helper function to calculate attendance percentage
def calculate_attendance(df):
    df['Total Sessions'] = df.iloc[:, 1:].apply(lambda row: row.count(), axis=1)
    df['Present Sessions'] = df.iloc[:, 1:].apply(lambda row: row.value_counts().get('P', 0), axis=1)
    df['Attendance %'] = (df['Present Sessions'] / df['Total Sessions']) * 100
    
    df.insert(1, 'Total Sessions', df.pop('Total Sessions'))
    df.insert(2, 'Present Sessions', df.pop('Present Sessions'))
    df.insert(3, 'Attendance %', df.pop('Attendance %'))
    
    return df

# Helper function to find students absent for 3 or more consecutive days
def find_consecutive_absentees(df):
    absentees = []
    for index, row in df.iterrows():
        attendance_series = row[1:-3]  # Exclude the 'Student Name', 'Total Sessions', 'Present Sessions', and 'Attendance %' columns
        max_consecutive_absent = (attendance_series == 'A').astype(int).groupby(attendance_series.ne('A').cumsum()).sum().max()
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


        # Read the "Teachers" sheet to get Teacher's Name, skipping initial metadata rows
        # teacher = pd.read_excel(uploaded_file, sheet_name='Teachers Note', skiprows=1)
        # teacher = teacher.iloc[:, :1]
        # teacherName = teacher["ITESKUL Teacher's Note"].str.extract(r'Trainer\s*(\d+)', flags=re.IGNORECASE)
        # st.write(teacherName)
        # # st.dataframe(teacher)
        # text = ""
        # for index, row in teacher.iterrows():
        #     for column_name, cell_value in row.items():
        #         text += str(cell_value) + " "  # Convert cell value to string and add to text
        #     text += "\n"  # Add newline after each row

        # # Display the regular text
        # print(text)
        
        
        # Read the "Attendance" sheet, skipping initial metadata rows
        df = pd.read_excel(uploaded_file, sheet_name='Attendance', skiprows=2)

        # Drop the first column if it is not relevant
        if df.columns[0].lower().startswith('unnamed'):
            df.drop(df.columns[0], axis=1, inplace=True)

        # Rename the first column to 'Student Name'
        df.rename(columns={df.columns[0]: 'Student Name'}, inplace=True)
        
        # Drop rows where 'Student Name' is empty or NaN
        df = df.dropna(subset=['Student Name'])

        # Drop empty columns
        df = df.dropna(axis=1, how='all')

        # Keep only columns containing "P" and "A"
        df = df.loc[:, ['Student Name'] + [col for col in df.columns[1:] if df[col].isin(['P', 'A', 'p', 'a']).any()]]

        # Calculate attendance percentage
        df = calculate_attendance(df)
        df.index = df.index + 1

        # Find students with 3 or more consecutive absences
        consecutive_absentees = find_consecutive_absentees(df)
        
        # Find students with attendance below 75%
        low_attendance = df[df['Attendance %'] < 75]['Student Name'].tolist()
        # Highlight rows where Attendance % < 75
        df_styled = df.style.apply(lambda x: ['background: #FFA07A' if x.name in low_attendance else '' for _ in x], axis=1)
        
        

        # Store the results in a dictionary
        class_reports[class_name] = {
            "Attendance Data": df_styled,
            "Low Attendance": low_attendance,
            "Consecutive Absentees": consecutive_absentees
        }

    # Display the results
    st.markdown("<hr>", unsafe_allow_html=True)
    
    for class_name, report in class_reports.items():
        st.subheader(f"Class: {class_name}")
        st.write("Attendance Data")
        st.dataframe(report["Attendance Data"])

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        if report["Low Attendance"]:
            st.write("Students with Attendance Below 75%")
            st.markdown("<ol>", unsafe_allow_html=True)
            for student in report["Low Attendance"]:
                st.markdown(f"<li>{student}</li>", unsafe_allow_html=True)
            st.markdown("</ol>", unsafe_allow_html=True)
        else:
            st.write("No students with attendance below 75%.")
            
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        if report["Consecutive Absentees"]:
            st.write("Students Absent for 3 or More Consecutive Days")
            st.markdown("<ul>", unsafe_allow_html=True)
            for student in report["Consecutive Absentees"]:
                st.markdown(f"<li>{student}</li>", unsafe_allow_html=True)
            st.markdown("</ul>", unsafe_allow_html=True)
        else:
            st.write("No students were absent for 3 or more consecutive days.")
        
        st.markdown("<br><br><hr>", unsafe_allow_html=True)
