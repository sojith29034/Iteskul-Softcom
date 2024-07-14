import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import requests


def run_main_app():      
    st.set_page_config(page_title="Student Attendance Report", layout="wide")
    
    st.markdown("""
        <style>
            .reportview-container {margin-top: -2em;}
            .st-emotion-cache-1jicfl2 {padding: 2rem 3rem 10rem;}
            h1#student-attendance-report {text-align: center;}
            header #MainMenu {visibility: hidden; display: none;}
            .stActionButton {visibility: hidden; display: none;}
            .stDeployButton {display:none;}
            footer {visibility: hidden;}
            stDecoration {display:none;}
            .stTabs button {margin-right: 50px;}
            .viewerBadge_container__r5tak {display: none;}
            p.credits {user-select: none; filter: opacity(0);}
        </style>
    """, unsafe_allow_html=True)
    
    
    # Function to fetch sample files from GitHub (adjust the URL and filenames as needed)
    def fetch_sample_files():
        files = [
            'https://github.com/sojith29034/Iteskul-Softcom/blob/main_branch/StudentData/German%20A1-WD-08.00pm-RR-23052024.xlsx',
            'https://github.com/sojith29034/Iteskul-Softcom/blob/main_branch/StudentData/Japanese%20N3-TR-2.00pm%20-SS-240923.xlsx',
            'https://github.com/sojith29034/Iteskul-Softcom/blob/main_branch/StudentData/Japanese%20N5-WN-2.00pm%20-WT-240923.xlsx'
        ]
        uploaded_files = []
    
        for file in files:
            url = file
            st.write(url)
            response = requests.get(url)
            if response.status_code == 200:
                uploaded_files.append((BytesIO(response.content)))
            else:
                st.warning(f"Failed to fetch {file} from GitHub.")
    
        return uploaded_files

        
    
    # Helper function to calculate attendance percentage
    def calculate_attendance(df):
        # Calculate total sessions and present sessions
        df['Total Sessions'] = df.iloc[:, 1:].apply(lambda row: row.count(), axis=1)
        df['Present Sessions'] = df.iloc[:, 1:].apply(lambda row: row.value_counts().get('P', 0) + row.value_counts().get('p', 0), axis=1)
        df['Attendance %'] = (df['Present Sessions'] / df['Total Sessions']) * 100
    
        # Reorder columns to have Student Name, Total Sessions, Present Sessions, and Attendance % at the start
        columns_order = ['Student Name', 'Total Sessions', 'Present Sessions', 'Attendance %'] + [col for col in df.columns if col not in ['Student Name', 'Total Sessions', 'Present Sessions', 'Attendance %']]
        df = df[columns_order]
    
        return df
    
    # Helper function to find students who left
    def students_left(df):
        left = []
        for index, row in df.iterrows():
            attendance = row[:-3]  # Get the last three attendance columns before the summary columns
            attendance = attendance.str.upper()  # Convert to uppercase 
            
            # Check if any student has "Left"
            if attendance.str.contains("LEFT").any():
                left.append(row['Student Name'])
        
        return left
    
    
    def find_consecutive_absentees(df):
        absentees = []
        for index, row in df.iterrows():
            attendance_series = row[-6:-3]  # Get the last three attendance columns before the summary columns
            attendance_series = attendance_series.str.upper()  # Convert to uppercase to handle both 'A' and 'a'
            
            # Check if all three values in the attendance_series are 'A'
            if attendance_series.tolist() == ['A', 'A', 'A']:
                absentees.append(row['Student Name'])
        
        return absentees
    
    # Helper function to find students absent for a total of 5 or more days
    def find_absentees(df):
        five_absentees = []
        ten_absentees = []
        
        for index, row in df.iterrows():
            attendance_series = row[3:]  # Get all attendance columns after the summary columns
            attendance_series = attendance_series.str.upper()  # Convert to uppercase to handle both 'A' and 'a'
            
            # Count total absences
            total_absences = attendance_series.tolist().count('A')
            
            if row['Total Sessions'] >= 25:
                if total_absences >= 10:
                    ten_absentees.append(row['Student Name'])
                    five_absentees.append(row['Student Name'])
            if row['Total Sessions'] >= 10:
                if total_absences >= 5:
                    if row['Student Name'] not in five_absentees:
                        five_absentees.append(row['Student Name']) 
        
        # Return absentees or default statement
        if ten_absentees:
            return five_absentees, ten_absentees
        elif five_absentees:
            return five_absentees, []
        else:
            return [], []
        
    # Highlight function
    def highlight_rows(row, condition_list):
        if row['Student Name'] in condition_list:
            return ['background-color: yellow'] * len(row)
        return [''] * len(row)
    
    
  
    # Helper function to pad lists to the same length
    def pad_lists(lists):
        max_len = max(len(lst) for lst in lists)
        return [lst + [""] * (max_len - len(lst)) for lst in lists]
    
    # Function to generate excel sheets
    def generate_excel(reports):
        output = BytesIO()
        
        # Create a Pandas Excel writer using openpyxl as the engine
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            trainer_report_data = []
    
            for class_name, report in reports.items():
                # Add class attendance data to a sheet
                df = report["Attendance Data"]
                df.to_excel(writer, sheet_name=f"{class_name[:20]} Data", index=False, startrow=1)
    
                # Add class summary to a separate sheet
                summary_data = {
                    "Attendance < 75%": report["Low Attendance"],
                    "3 Consecutive Absents": report["Consecutive Absentees"],
                    "5 Absents (at least 10 sessions)": report["Five Absent"],
                    "10 Absents (at least 25 sessions)": report["Ten Absent"],
                    "Discontinued": report["Discontinued"]
                }
    
                # Ensure the lists are padded to the same length
                padded_data = pad_lists(list(summary_data.values()))
    
                # Create a DataFrame for the summary data
                summary_df = pd.DataFrame(
                    padded_data, 
                    index=summary_data.keys()
                ).transpose()
                
                summary_df.to_excel(writer, sheet_name=f"{class_name[:20]} Summary", index=False, startrow=1)
    
                # Collect trainer report data
                trainer_report_data.append({
                    "Class": class_name,
                    "Trainer": report["Trainer"],
                    "Last Updated Date": report["Last date"]
                })
    
            # Add the Trainer's Report sheet
            trainer_report_df = pd.DataFrame(trainer_report_data)
            trainer_report_df.to_excel(writer, sheet_name="Trainer's Report", index=False)
    
        # Get the workbook object from the writer
        workbook = writer.book
        
        for class_name in reports.keys():
            for sheet_name_suffix in [" Data", " Summary"]:
                sheet_name = f"{class_name[:20]}{sheet_name_suffix}"
                if sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]
                    
                    # Insert the class name in the first row and merge cells
                    sheet.insert_rows(1)
                    sheet["A1"] = class_name
                    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=sheet.max_column)
                    
                    # Align the text to the left
                    sheet["A1"].alignment = Alignment(horizontal="left")
    
        # Save the modified workbook back to BytesIO
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        return output
    
    def find_trainer_notes(excel_file):
        # Load the workbook
        wb = load_workbook(excel_file, read_only=True)
        
        # Check sheet names and find "Teachers Note"
        sheet_names = wb.sheetnames
        if "Teachers Note" not in sheet_names:
            return None
        
        # Load the sheet
        ws = wb["Teachers Note"]
        
        # Search for "Trainer: " in the sheet
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str) and "Trainer" in cell:
                    return cell
        
        return None
    
    ###############################################################################################################
    ###############################################################################################################
    ###############################################################################################################
            
    # Streamlit app
    st.markdown("<p class='credits'>Made by <a href='https://github.com/sojith29034'>Sojith Sunny</a></p>", unsafe_allow_html=True)

    if st.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()
        
    st.title("Student Attendance Report")
    
    if st.checkbox("Use Sample Files"):
        uploaded_files = fetch_sample_files()
    else:
        uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)
    
    if uploaded_files:
        class_reports = {}
        for uploaded_file in uploaded_files:
            class_name = uploaded_file.name.split('.xlsx')[0]  # get class name from the file name
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
            
            # Drop rows where student has left
            students_left_list = students_left(df)
            df = df[~df['Student Name'].isin(students_left_list)]
            
            # Clean up student names (removing additional details if present)
            for i in range(len(students_left_list)):
                students_left_list[i] = students_left_list[i].split('(')[0].strip()
    
            # Drop empty columns
            df = df.dropna(axis=1, how='all')
            
            # Reset the index
            df.reset_index(drop=True, inplace=True)
            
            # Start index of dataframe from 1
            df.index = df.index + 1
            
            # Change column names to DD/MM/YYYY format if they are timestamps
            for i in range(3, len(df.columns)):
                df.columns.values[i] = datetime.strptime(str(df.columns.values[i]), "%Y-%m-%d %H:%M:%S").strftime("%d-%m-%Y")
                
            # Get Last Date of attendance being updated
            last_date = df.columns[-1]
    
            # Find Trainer Name
            trainer_name = find_trainer_notes(uploaded_file)
            if trainer_name:
                trainer_name = trainer_name.split(":")[1].strip()
            else:
                trainer_name = "Not Found"
            
            # Keep only columns containing "P" and "A"
            df = df.loc[:, ['Student Name'] + [col for col in df.columns[1:] if df[col].isin(['P', 'A', 'p', 'a']).any()]]
    
            # Calculate attendance percentage
            df = calculate_attendance(df)
    
            # Find students with 3 or more consecutive absences
            consecutive_absentees = find_consecutive_absentees(df)
            
            # Find students absent for 5 or more days
            five_absent, ten_absent = find_absentees(df)
            
            # Find students with attendance below 75%
            low_attendance = df[df['Attendance %'] < 75]['Student Name'].tolist()
    
            # Store the results in a dictionary
            class_reports[class_name] = {
                "Last date": last_date,
                "Trainer": trainer_name,
                "Attendance Data": df,
                "Low Attendance": low_attendance,
                "Five Absent": five_absent,
                "Ten Absent": ten_absent,
                "Consecutive Absentees": consecutive_absentees,
                "Discontinued": students_left_list
            }
            
        excel_data = generate_excel(class_reports)
        st.download_button(
            label="Download Attendance Report",
            data=excel_data,
            file_name=f"attendance_report_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
        # Group classes by language
        language_classes = {}
        for class_name, report in class_reports.items():
            lang = class_name.split(' ')[0]  # get the first part of the class name as the language
            if lang not in language_classes:
                language_classes[lang] = {}
            language_classes[lang][class_name] = report
            
        # Display the results in tabs
        st.markdown("<hr>", unsafe_allow_html=True)
        
        tab_titles = list(language_classes.keys())
        tabs = st.tabs(tab_titles)
        
        # Display each language in its own tab
        for tab, lang in zip(tabs, language_classes.keys()):
            with tab:
                st.subheader(f"{lang} Classes")
                    
                # Get the list of class names for the current language
                class_names = list(language_classes[lang].keys())
                
                # Select the first class by default
                default_class = class_names[0] if class_names else None
                
                # Create a dropdown menu to select classes for the current language
                selected_classes = st.multiselect("Select Classes", class_names, default=default_class)
                
                # Iterate over selected classes and display their details
                for class_name in selected_classes:
                    report = language_classes[lang][class_name]
                    
                    # Separate language tabs with a horizontal rule
                    st.markdown("<hr>", unsafe_allow_html=True)
                    
                    st.subheader(f"Class: {class_name}")
                    st.write(f"Trainer: {report['Trainer']}")
                    st.write(f"Last updated on: {report['Last date']}")
                    st.write("Attendance Data")
                    st.dataframe(report["Attendance Data"])
    
                    sub_tab_titles = ["Attendance < 75%", "3 Consecutive Absents", "5 Absents (atleast 10 sessions)", "10 Absents (atleast 25 sessions)","Discontinued"]
                    sub_tabs = st.tabs(sub_tab_titles)
    
                    sub_tab_content = {
                        "Attendance < 75%": report["Low Attendance"],
                        "3 Consecutive Absents": report["Consecutive Absentees"],
                        "5 Absents (atleast 10 sessions)": report["Five Absent"],
                        "10 Absents (atleast 25 sessions)": report["Ten Absent"],
                        "Discontinued": report["Discontinued"]
                    }
    
                    for sub_tab, sub_tab_title in zip(sub_tabs, sub_tab_titles):
                        with sub_tab:
                            st.write(sub_tab_title)
                            details = sub_tab_content.get(sub_tab_title)
                            if details:
                                st.markdown("<ul>", unsafe_allow_html=True)
                                for student in details:
                                    st.markdown(f"<li>{student}</li>", unsafe_allow_html=True)
                                st.markdown("</ul>", unsafe_allow_html=True)
                            else:
                                st.write(f"No student has {sub_tab_title.lower()}.")
    
                # Separate language tabs with a horizontal rule
                st.markdown("<hr>", unsafe_allow_html=True)
                
        excel_data = generate_excel(class_reports)
        st.download_button(
            label="Download Cumulative Attendance Report",
            data=excel_data,
            file_name=f"attendance_report_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    run_main_app()
