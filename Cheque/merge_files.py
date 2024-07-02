# pip install os numpy pandas

import os
import zipfile
import pandas as pd
from glob import glob
import tkinter as tk
from tkinter import filedialog
import numpy as np

# Function to select the directory containing files
def select_directory():
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title='Select Directory Containing Files')
    return directory

# Function to detect and extract ZIP files, then return a list of extracted CSV files
def extract_zip_files(zip_dir, extract_dir):
    # Create the extract directory if it doesn't exist
    os.makedirs(extract_dir, exist_ok=True)

    # Extract all zip files
    for zip_file in glob(os.path.join(zip_dir, '*.zip')):
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

    # Find all CSV files in the extraction directory
    csv_files = glob(os.path.join(extract_dir, '**', '*.csv'), recursive=True)
    return csv_files

# Main process
def main():
    # Select directory using file dialog
    directory = select_directory()

    if not directory:
        print("No folder selected.")
        return

    # Initialize lists to store dataframes and summary data
    df_list = []
    summary_data = []

    # Check for ZIP files in the directory
    zip_files = glob(os.path.join(directory, '*.zip'))

    if zip_files:
        # Extract ZIP files and get CSV files
        csv_files = extract_zip_files(directory, directory)
        # Read and append CSV files to the dataframe list
        for file in csv_files:
            df = pd.read_csv(file)
            df = df.replace("/", "-")
            df_list.append(df)
            summary_data.append({"File": os.path.basename(file), "Number of Entries": len(df)})
    else:
        # Get all Excel and CSV files in the selected folder
        excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx') or f.endswith('.xlsm') or f.endswith('.CSV')]

        if not excel_files:
            print("No Excel or CSV files found.")
            return

        # Read and append Excel and CSV files to the dataframe list
        for file in excel_files:
            file_path = os.path.join(directory, file)
            if file.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file.endswith('.xlsm'):
                df = pd.read_excel(file_path, sheet_name='Data')
            elif file.endswith('.CSV'):
                df = pd.read_csv(file_path)
            df = df.replace(np.nan, "NA")
            df_list.append(df)
            summary_data.append({"File": file, "Number of Entries": len(df)})

    if not df_list:
        print("No data read from files.")
        return

    # Concatenate all dataframes
    merged_df = pd.concat(df_list, ignore_index=True)
    # Replace '/' with '-' in all columns
    merged_df.replace('/', '-', inplace=True, regex=True)

    # Create summary dataframe
    summary_df = pd.DataFrame(summary_data)

    # Calculate total entries
    total_entries = summary_df["Number of Entries"].sum()
    total_row = pd.DataFrame([{"File": "Total", "Number of Entries": total_entries}])
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)

    # Define the output file path with the folder name
    folder_name = os.path.basename(directory)
    output_file = os.path.join(directory, f'{folder_name}_master.xlsx')

    # Write merged dataframe and summary to an Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='Merged Data', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    print(f'Master file created: {output_file}')


if __name__ == "__main__":
    main()
