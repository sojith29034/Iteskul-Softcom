import streamlit as st
import pandas as pd
from datetime import datetime

# Helper function to calculate attendance percentage
def calculate_attendance(df):
    df['Total Sessions'] = df.iloc[:, 1:].apply(lambda row: row.count(), axis=1)
    df['Present Sessions'] = df.iloc[:, 1:].apply(lambda row: row.value_counts().get('P', 0), axis=1)
    df['Attendance %'] = (df['Present Sessions'] / df['Total Sessions']) * 100
    
    df.insert(1, 'Total Sessions', df.pop('Total Sessions'))
    df.insert(2, 'Present Sessions', df.pop('Present Sessions'))
    df.insert(3, 'Attendance %', df.pop('Attendance %'))
    
    return df

# Helper function to find students absent for 3 consecutive days
def find_consecutive_absentees(df):
    absentees = []
    for index, row in df.iterrows():
        attendance_series = row[-6:-3]  # Get the last three attendance columns before the summary columns
        attendance_series = attendance_series.str.upper()  # Convert to uppercase to handle both 'A' and 'a'
        
        # Check if all three values in the attendance_series are 'A'
        if attendance_series.tolist() == ['A', 'A', 'A']:
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
        
        # Change column names to DD/MM/YYYY format if they are timestamps
        for i in range(3, len(df.columns)):
            df.columns.values[i] = datetime.strptime(str(df.columns.values[i]), "%Y-%m-%d %H:%M:%S").strftime("%d-%m-%Y")
            
        # Get Last Date of attendance being updated
        last_date = df.columns[-1]

        # Keep only columns containing "P" and "A"
        df = df.loc[:, ['Student Name'] + [col for col in df.columns[1:] if df[col].isin(['P', 'A', 'p', 'a']).any()]]

        # Calculate attendance percentage
        df = calculate_attendance(df)

        # Find students with 3 or more consecutive absences
        consecutive_absentees = find_consecutive_absentees(df)
        
        # Find students with attendance below 75%
        low_attendance = df[df['Attendance %'] < 75]['Student Name'].tolist()

        # Store all the data for this class in the same sheet
        # class_data = pd.DataFrame({
        #     "Low Attendance": [', '.join(low_attendance)],
        #     "Consecutive Absentees": [', '.join(consecutive_absentees)]
        # })

        df_styled = df.style.apply(lambda x: ['background: #FFA07A' if x['Attendance %'] < 75 else '' for _ in x], axis=1)
        
        # Store the results in a dictionary
        class_reports[class_name] = {
            "Last date": last_date,
            "Attendance Data": df_styled,
            "Low Attendance": low_attendance,
            "Consecutive Absentees": consecutive_absentees
        }

    # Display the results
    st.markdown("<hr>", unsafe_allow_html=True)
    
    for class_name, report in class_reports.items():
        st.subheader(f"Class: {class_name}")
        st.write(f"Last updated on: {report['Last date']}")
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
            st.write("Students Absent for 3 Consecutive Days")
            st.markdown("<ul>", unsafe_allow_html=True)
            for student in report["Consecutive Absentees"]:
                st.markdown(f"<li>{student}</li>", unsafe_allow_html=True)
            st.markdown("</ul>", unsafe_allow_html=True)
        else:
            st.write("No students were absent for 3 consecutive days.")
        
        st.markdown("<br><br><hr>", unsafe_allow_html=True)
        
        
    # Download link for all data
    # if st.button('Download All Data as Excel'):
    #     output = pd.ExcelWriter('student_attendance_report.xlsx')
    #     for class_name, report in class_reports.items():
    #         report["Attendance Data"].to_excel(output, sheet_name=f"{class_name}_Attendance", index=False)
    #         class_data = pd.DataFrame({
    #             "Attendance < 75%": [', '.join(report["Low Attendance"])],
    #             # "3 Consecutive Absentees": [', '.join(report["Consecutive Absentees"])]
    #             "3 Consecutive Absentees": [[student] for student in report["Consecutive Absentees"]]
    #         })
    #         class_data.to_excel(output, sheet_name=f"{class_name}_Summary", index=False)
    #     output.close()
    #     with open('student_attendance_report.xlsx', 'rb') as f:
    #         file_contents = f.read()
    #     st.download_button(label='Download', data=file_contents, file_name='student_attendance_report.xlsx', mime='application/octet-stream')


    # # Create a list to store DataFrames
    # dataframes = []

    # for class_name, report in class_reports.items():
    #     # Extracting the DataFrame from the styler object
    #     report_df = report["Attendance Data"].data

    #     for _, row in report_df.iterrows():
    #         dataframes.append(pd.DataFrame({
    #             "Class Name": [class_name],
    #             "Student Name": [row["Student Name"]],
    #             "Attendance %": [row["Attendance %"]],
    #             "Low Attendance": [row["Student Name"] in report["Low Attendance"]],
    #             "Consecutive Absentees": [row["Student Name"] in report["Consecutive Absentees"]]
    #         }))

    # # Concatenate all DataFrames into one
    # all_data = pd.concat(dataframes, ignore_index=True)

    # # Display the combined DataFrame
    # st.write("Combined Data:")
    # st.dataframe(all_data)

    # # Download the combined DataFrame as Excel
    # if st.button('Download Combined Data as Excel'):
    #     output = pd.ExcelWriter('combined_student_attendance_report.xlsx')
    #     all_data.to_excel(output, index=False)
    #     output.close()
    #     with open('combined_student_attendance_report.xlsx', 'rb') as f:
    #         file_contents = f.read()
    #     st.download_button(label='Download', data=file_contents, file_name='combined_student_attendance_report.xlsx', mime='application/octet-stream')



    # Create ExcelWriter object
    output = pd.ExcelWriter('student_attendance_report.xlsx')

    for class_name, report in class_reports.items():
        # Save each DataFrame to a separate sheet
        report["Attendance Data"].to_excel(output, sheet_name=f"{class_name}_Attendance", index=False)

    # Save the Excel file
    output.close()

    # Download the Excel file
    with open('student_attendance_report.xlsx', 'rb') as f:
        file_contents = f.read()
    st.download_button(label='Download', data=file_contents, file_name='student_attendance_report.xlsx', mime='application/octet-stream')