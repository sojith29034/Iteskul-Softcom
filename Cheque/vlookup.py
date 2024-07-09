import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook

def upload_bank_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        bank_file_entry.delete(0, tk.END)
        bank_file_entry.insert(0, filepath)

def upload_data_entry_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        data_entry_file_entry.delete(0, tk.END)
        data_entry_file_entry.insert(0, filepath)

def upload_vlookup_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        vlookup_file_entry.delete(0, tk.END)
        vlookup_file_entry.insert(0, filepath)

def compare_files():
    bank_file = bank_file_entry.get()
    data_entry_file = data_entry_file_entry.get()
    vlookup_file = vlookup_file_entry.get()

    if not bank_file or not data_entry_file:
        messagebox.showerror("Error", "Please upload all files")
        return

    try:
        bank_df = pd.read_excel(bank_file)
        data_entry_df = pd.read_excel(data_entry_file)
        
        # Save the original data types
        original_bank_dtypes = bank_df.dtypes
        original_data_entry_dtypes = data_entry_df.dtypes
        
        bank_df = pd.read_excel(bank_file, dtype=str)
        data_entry_df = pd.read_excel(data_entry_file, dtype=str)
        
        # Clean column names in the bank file
        bank_df.columns = bank_df.columns.str.strip()

        # Clean column names in the data entry file
        data_entry_df.columns = data_entry_df.columns.str.strip()
        
        # Ensure necessary columns are present
        required_bank_columns = ['INST NO', 'INST AMOUNT']
        required_data_entry_columns = ['Cheque Number', 'Amount', 'Account Number', 'Name']
        
        for col in required_bank_columns:
            if col not in bank_df.columns:
                raise KeyError(f"Column '{col}' not found in the bank master file.")
        
        for col in required_data_entry_columns:
            if col not in data_entry_df.columns:
                raise KeyError(f"Column '{col}' not found in the data entry master file.")

        # Concatenate columns
        bank_df['Concat'] = bank_df['INST NO'] + " " + bank_df['INST AMOUNT']
        data_entry_df['Concat'] = data_entry_df['Cheque Number'] + " " + data_entry_df['Amount']

        # Merge dataframes on the concatenated fields
        merged_df = pd.merge(bank_df, data_entry_df[['Concat', 'Account Number', 'Name']], on='Concat', how='left')

        # Convert columns back to original data types
        for col in original_bank_dtypes.index:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].astype(original_bank_dtypes[col])
        
        for col in original_data_entry_dtypes.index:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].astype(original_data_entry_dtypes[col])

        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, merged_df.to_string())
        
        if vlookup_file:
            # Load the existing VLOOKUP file
            vlookup_df = pd.read_excel(vlookup_file)
            # vlookup_df = vlookup_df.astype(str)  # Ensure same dtype for comparison
            
            # Filter out rows that are already in the VLOOKUP file
            new_entries = merged_df[~merged_df['Concat'].isin(vlookup_df['Concat'])]
            
            # Append the new data to the VLOOKUP dataframe
            updated_vlookup_df = pd.concat([vlookup_df, new_entries], ignore_index=True).drop_duplicates()
        else:
            updated_vlookup_df = merged_df

        # To save as excel file
        def save_to_excel():
            save_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_filepath:
                try:
                    # Create a Pandas Excel writer using XlsxWriter as the engine
                    writer = pd.ExcelWriter(save_filepath, engine='openpyxl')
                    updated_vlookup_df.to_excel(writer, index=False, sheet_name='Sheet1')

                    # Access the workbook and worksheet objects
                    workbook = writer.book
                    worksheet = workbook['Sheet1']

                    # Convert all cells in the worksheet to 'General' format
                    for row in worksheet.iter_rows(min_row=2):  # Start from the second row (header is row 1)
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.isdigit() and len(cell.value)<10:
                                # Convert the string to an integer
                                cell.value = int(cell.value)  # Use float() if it's a floating point number

                        # Set number format to General to ensure it displays as a number
                        cell.number_format = 'General'

                    # Save the workbook
                    # writer.save()
                    writer.close()

                    messagebox.showinfo("Success", f"File saved successfully to {save_filepath}")
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving file: {e}")
            else:
                messagebox.showwarning("Warning", "Please enter a valid file name.")


        save_button.config(state=tk.NORMAL)
        save_button.config(command=save_to_excel)

    except KeyError as ke:
        messagebox.showerror("Error", str(ke))
    except Exception as e:
        messagebox.showerror("Error", f"Error processing files: {e}")

# Create the main window
root = tk.Tk()
root.title("VLOOKUP Update Dashboard")

# Bank master file upload
bank_master_label = tk.Label(root, text="Bank Master File:")
bank_master_label.grid(row=0, column=0, padx=10, pady=10)
bank_file_entry = tk.Entry(root, width=50)
bank_file_entry.grid(row=0, column=1, padx=10, pady=10)
bank_master_button = tk.Button(root, text="Browse", command=upload_bank_file)
bank_master_button.grid(row=0, column=2, padx=10, pady=10)

# Data entry master file upload
data_entry_master_label = tk.Label(root, text="Data Entry Master File:")
data_entry_master_label.grid(row=1, column=0, padx=10, pady=10)
data_entry_file_entry = tk.Entry(root, width=50)
data_entry_file_entry.grid(row=1, column=1, padx=10, pady=10)
data_entry_master_button = tk.Button(root, text="Browse", command=upload_data_entry_file)
data_entry_master_button.grid(row=1, column=2, padx=10, pady=10)

# VLOOKUP file upload
vlookup_file_label = tk.Label(root, text="VLOOKUP File:")
vlookup_file_label.grid(row=2, column=0, padx=10, pady=10)
vlookup_file_entry = tk.Entry(root, width=50)
vlookup_file_entry.grid(row=2, column=1, padx=10, pady=10)
vlookup_file_button = tk.Button(root, text="Browse", command=upload_vlookup_file)
vlookup_file_button.grid(row=2, column=2, padx=10, pady=10)

# Compare button
compare_button = tk.Button(root, text="Compare", command=compare_files)
compare_button.grid(row=3, column=1, pady=20)

# Save button (initially disabled)
save_button = tk.Button(root, text="Save to Excel", state=tk.DISABLED)
save_button.grid(row=3, column=2, pady=20)

# Result display
result_text = tk.Text(root, wrap=tk.NONE, width=100, height=20)
result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

# Start the GUI event loop
root.mainloop()
