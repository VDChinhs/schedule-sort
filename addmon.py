from tkinter import *
from tkinter import ttk,messagebox
import tkinter as tk
import openpyxl
import Guilec
import os

# GUI đế thêm môn học 
class Addmonhoc:
    def __init__(self,root,nguoigoi,classgoi):
        self.classgoi = classgoi
        self.nguoigoi = nguoigoi
        self.root = root
        self.root.geometry('400x120')
        self.root.title('Thêm môn học')

        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Môn học")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)

        self.name_entry = ttk.Entry(self.widgets_frame,font=("Helvetica", 20))
        self.name_entry.insert(0,"Tên")
        self.name_entry.bind("<FocusIn>", lambda e: self.name_entry.delete('0', 'end'))
        self.name_entry.grid(row=0,column=0,sticky='ew')

        button = ttk.Button(self.widgets_frame, text="Thêm", command=self.insert_row)
        button.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        self.name_entry.bind("<Return>", self.perform_insert)
    
    def perform_insert(self, event=None):
        self.insert_row()

    def insert_row(self):
        col = []
        name = self.name_entry.get()
        path = os.getcwd() + "\\alpha.xlsx"
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        data = list(sheet.values)

        if len(data) == 0:
            row_values = ["Môn"]
            sheet.append(row_values)

        for i in data[1:]:
            col.append(i[0])
        for j in col:
            if j.lower() == name.lower():
                messagebox.showerror(title="Lỗi", message="Môn học đã tồn tại")
                return
            
        row_values = [name]
        sheet.append(row_values)
        workbook.save(path)

        if self.classgoi.getname() == "Guigiangvien":
            Guilec.Guigiangvien(self.nguoigoi).close(self.nguoigoi)
            Guilec.Guigiangvien(tk.Tk())   

        workbook.close()
        self.root.destroy()

# if __name__ == "__main__":
#     root = tk.Tk()
#     Addmonhoc(root,root)
#     root.mainloop()