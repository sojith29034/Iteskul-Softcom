# pip install os numpy pandas

import os
import pandas as pd
from tkinter import Tk, filedialog
import numpy as np

def merge_excel_files():
    # Initialize Tkinter root
    root = Tk()
    root.withdraw()  # Hide the root window

    # Open directory dialog to select the folder containing Excel and CSV files
    folder_path = filedialog.askdirectory(
        title="Select Folder Containing Excel and CSV Files",
    )
    
    # If no folder is selected, exit the function
    if not folder_path:
        print("No folder selected.")
        return

    # List to store dataframes
    df_list = []
    summary_data = []

    # Get all Excel and CSV files in the selected folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xlsm') or f.endswith('.CSV')]

    # Print the list of files found
    print(f"Files found: {excel_files}")

    # Read and append dataframes
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        print(f"Reading file: {file_path}")
        
        if file.endswith('.xlsx'):
            df = pd.read_excel(file_path)  # Read with dtype=str to ensure NA values are preserved as strings
        elif file.endswith('.xlsm'):
            df = pd.read_excel(file_path, sheet_name='Data')
        elif file.endswith('.CSV'):
            df = pd.read_csv(file_path) 
            
        # Replace Nan with "NA"
        df = df.replace(np.nan, "NA")
        
        df_list.append(df)
        
        # Append to summary
        summary_data.append({"Excel File": file, "Number of Entries": len(df)})
    
    if not df_list:
        print("No Excel or CSV files found or no data read from files.")
        return
    
    # Concatenate all dataframes
    merged_df = pd.concat(df_list, ignore_index=True)
    
    # Create summary dataframe
    summary_df = pd.DataFrame(summary_data)
    
    # Calculate total entries
    total_entries = summary_df["Number of Entries"].sum()
    total_row = pd.DataFrame([{"Excel File": "Total", "Number of Entries": total_entries}])
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)
    
    # Define the output file path with the folder name
    folder_name = os.path.basename(folder_path)
    output_file = os.path.join(folder_path, f'{folder_name}_master.xlsx')
    
    # Write merged dataframe and summary to an Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='Merged Data', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"Master file created: {output_file}")

# Merge excel files and create summary
merge_excel_files()
