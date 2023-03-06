import tkinter as tk
import xlwings as xw
import openpyxl
import shutil
from tkinter import messagebox
import psutil

################################################################################
#
class CountingApp:
    def __init__(self, master):
        self.master = master
        self.start_number = tk.IntVar(value='')
        self.current_number = tk.IntVar(value=None)
        self.input_string_1 = tk.StringVar(value='')
        self.input_string_2 = tk.StringVar(value='')
        self.input_string_3 = tk.StringVar(value='')
        self.input_int = tk.StringVar(value='')
        self.page_one()

    def page_one(self):
        self.start_label = tk.Label(self.master, text="Enter starting number:")
        self.start_label.pack()
        self.start_entry = tk.Entry(self.master, textvariable=self.start_number)
        self.start_entry.pack()
        self.continue_button = tk.Button(self.master, text="Continue", command=self.page_two)
        self.continue_button.pack()

    def page_two(self):
        self.start_label.destroy()
        self.start_entry.destroy()
        self.continue_button.destroy()

        if self.start_number.get() is not None:
            self.current_number.set(self.start_number.get())
        else:
            self.current_number.set('')

        self.show_current = tk.Label(self.master, text="Receipt No.")
        self.show_current.pack()
        self.current_value = tk.Label(self.master, textvariable=self.current_number)
        self.current_value.pack()


        self.string_label_1 = tk.Label(self.master, text="ชื่อ")
        self.string_label_1.pack()
        self.string_entry_1 = tk.Entry(self.master, textvariable=self.input_string_1)
        self.string_entry_1.pack()

        self.string_label_2 = tk.Label(self.master, text="วันที่ [วว/ดด/20xx]")
        self.string_label_2.pack()
        self.string_entry_2 = tk.Entry(self.master, textvariable=self.input_string_2)
        self.string_entry_2.pack()

        self.string_label_3 = tk.Label(self.master, text="รายการ")
        self.string_label_3.pack()
        self.string_entry_3 = tk.Entry(self.master, textvariable=self.input_string_3)
        self.string_entry_3.pack()

        self.int_label = tk.Label(self.master, text="ราคาต่อหน่วย")
        self.int_label.pack()
        self.int_entry = tk.Entry(self.master, textvariable=self.input_int)
        self.int_entry.pack()

        

        self.count_button = tk.Button(self.master, text="Next", command=self.count)
        self.count_button.pack(side="left")

        self.quit_button = tk.Button(self.master, text="Quit", command=self.master.quit)
        self.quit_button.pack(side="right")

    def count(self):
        current = self.current_number.get()
        self.current_number.set(current)

        input_string_1 = self.input_string_1.get()
        input_string_2 = self.input_string_2.get()
        input_string_3 = self.input_string_3.get()
        input_int = self.input_int.get()

       

        if not input_string_1 or not input_string_2 or not input_string_3 or not input_int:
            messagebox.showerror("Error", "กรุณากรอกให้ครบ")
            return
        try:
            input_int = int(input_int)
        except ValueError:
            messagebox.showerror("Error", "ใส่ตัวเลข")
            return
        
        input_int = int(self.input_int.get())

        self.input_string_1.set(value='')
        self.input_string_2.set(value='')
        self.input_string_3.set(value='')
        self.input_int.set(value='')


        #############################
        #Exel
        # Open the template file
        template_path = 'C:/Users/Windows10/Desktop/python/template.xlsx'
        wb_template = openpyxl.load_workbook(template_path)

        # Get the sheet you want to populate with data
        sheet = wb_template.active
        file_name = f"{current}.xlsx"
        shutil.copy2(template_path, file_name)
        wb = xw.Book(file_name)  # this will open a new workbook
        sheet = wb.sheets['ใบเสร็จรับเงิน']

        #รับค่า
        name = input_string_1
        date = input_string_2
        descrition = input_string_3
        unitprice = input_int
        amount = unitprice

        #แก้ไข
        sheet['E9'].value=name
        sheet['L12'].value=date
        sheet['L10'].value=current
        sheet['C18'].value=descrition
        sheet['J18'].value=unitprice
        sheet['L18'].value=amount
        sheet['L28'].value=amount

        #check if want to continue
    

        #fix Read-Only
        import os
        # check if the file is read-only
        if os.access(file_name, os.W_OK):
            print("")
        else:
            # change the read-only attribute
            os.chmod(file_name, 0o600)
            print("")

        wb.save(f'C:/Users/Windows10/Desktop/python/{file_name}')
        

        # close Excel workbook and terminate Excel process
        wb.close()
        for process in psutil.process_iter():
            if process.name() == "EXCEL.EXE":
                process.kill()


        wb_template.close()
        current += 1
        self.current_number.set(current)
################################################################################



# create tkinter window
window = tk.Tk()
window.title("Receipt Program")
window.geometry("300x300")

# create counting app object
app = CountingApp(window)

# start tkinter main loop
window.mainloop()