# pip install pillow openpyxl


from tkinter import Tk, Label, Entry, Button
from tkinter import filedialog
from PIL import Image, ImageTk
import os
import openpyxl

# Function to save data to Excel file
def save_to_excel(filename, cheque_number, amount, account_number, name):
    wb = openpyxl.load_workbook('cheque_data.xlsx')
    sheet = wb.active
    sheet.append([filename, cheque_number, amount, account_number, name])
    wb.save('cheque_data.xlsx')

class ChequeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cheque Image Viewer and Data Entry")

        self.current_index = 0
        self.images = self.load_images()

        self.filename_label = Label(root, text="Filename:")
        self.filename_label.grid(row=0, column=0, sticky="w")

        self.filename_var = Label(root, text="")
        self.filename_var.grid(row=0, column=1, sticky="w")

        self.image_label = Label(root)
        self.image_label.grid(row=1, column=0, columnspan=2)

        self.cheque_number_label = Label(root, text="Cheque Number:")
        self.cheque_number_label.grid(row=2, column=0, sticky="w")

        self.cheque_number_entry = Entry(root)
        self.cheque_number_entry.grid(row=2, column=1)

        self.amount_label = Label(root, text="Amount:")
        self.amount_label.grid(row=3, column=0, sticky="w")

        self.amount_entry = Entry(root)
        self.amount_entry.grid(row=3, column=1)

        self.account_number_label = Label(root, text="Account Number:")
        self.account_number_label.grid(row=4, column=0, sticky="w")

        self.account_number_entry = Entry(root)
        self.account_number_entry.grid(row=4, column=1)

        self.name_label = Label(root, text="Name:")
        self.name_label.grid(row=5, column=0, sticky="w")

        self.name_entry = Entry(root)
        self.name_entry.grid(row=5, column=1)

        self.previous_button = Button(root, text="Previous", command=self.show_previous)
        self.previous_button.grid(row=6, column=0)

        self.next_button = Button(root, text="Next", command=self.show_next)
        self.next_button.grid(row=6, column=1)

        self.update_display()

    def load_images(self):
        folder_path = filedialog.askdirectory(title="Select Folder with Images")
        if not folder_path:
            exit()

        image_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.tif') and file.lower().endswith('f.tif')]

        images = []
        for file in image_files:
            image_path = os.path.join(folder_path, file)
            image = Image.open(image_path)
            images.append((image, file))

        return images

    def update_display(self):
        if self.current_index >= len(self.images):
            self.current_index = 0
        elif self.current_index < 0:
            self.current_index = len(self.images) - 1

        image, filename = self.images[self.current_index]
        self.filename_var.config(text=filename)
        self.image_tk = ImageTk.PhotoImage(image)
        self.image_label.config(image=self.image_tk)

    def show_previous(self):
        self.current_index -= 1
        self.update_display()

    def show_next(self):
        filename = self.filename_var.cget("text")
        cheque_number = self.cheque_number_entry.get()
        amount = self.amount_entry.get()
        account_number = self.account_number_entry.get()
        name = self.name_entry.get()

        save_to_excel(filename, cheque_number, amount, account_number, name)

        self.current_index += 1
        self.update_display()

if __name__ == "__main__":
    root = Tk()
    app = ChequeApp(root)
    root.mainloop()
