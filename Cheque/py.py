import pytesseract
from PIL import Image
import os
import re
import openpyxl

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_cheque_data(image_path):
    # Open the image file
    img = Image.open(image_path)
    
    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(img)
    
    # print(text)
    
    # Extract cheque number, amount, account number, and name using regex or string operations
    cheque_number = extract_cheque_number(text)
    amount = extract_amount(text)
    account_number = extract_account_number(text)
    name = extract_name(text)
    
    return cheque_number, amount, account_number, name

def extract_cheque_number(text):
    # Example regex to match a cheque number pattern
    match = re.search(r'Cheque Number:\s*(\d+)', text, re.IGNORECASE)
    return match.group(1) if match else 'Not Found'

def extract_amount(text):
    # Example regex to match an amount pattern
    match = re.search(r'RUPEES \s*\$?([\d,]+\.\d{2})', text, re.IGNORECASE)
    return match.group(1) if match else 'Not Found'

def extract_account_number(text):
    # Example regex to match an account number pattern
    match = re.search(r'A/c No. \s*(\d+)', text, re.IGNORECASE)
    return match.group(1) if match else 'Not Found'

def extract_name(text):
    # Example regex to match a name pattern
    match = re.search(r'On Demand Pay \s*([A-Za-z\s]+)', text, re.IGNORECASE)
    return match.group(1) if match else 'Not Found'

def main():
    # Folder containing the TIFF files
    folder_path = 'C:/Users/sojit/Downloads/tiff'
    
    # Create a new Excel workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Cheque Data'
    
    # Write headers to the Excel sheet
    headers = ['File Name', 'Cheque Number', 'Amount', 'Account Number', 'Name']
    sheet.append(headers)
    
    # Process each TIFF file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('f.tif') or file_name.lower().endswith('f.tiff'):
            file_path = os.path.join(folder_path, file_name)
            cheque_number, amount, account_number, name = extract_cheque_data(file_path)
            sheet.append([file_name, cheque_number, amount, account_number, name])
    
    # Save the Excel workbook
    workbook.save('Cheque_Data.xlsx')

if __name__ == '__main__':
    main()
