import re

text = "Here is the Cheque Number: 123456."

match = re.search(r'Cheque Number:\s*(\d+)', text, re.IGNORECASE)
if match:
    cheque_number = match.group(1)
    print(f"Cheque Number: {cheque_number}")
else:
    print("Cheque Number not found.")
