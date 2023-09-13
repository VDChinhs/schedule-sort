from tkinter import *
from tkinter import ttk,messagebox
import tkinter as tk
import openpyxl
import addlec
import addmon
import readalpha as ra
import os


# GUI hiển thị chí số alpha giáo viên ứng với mỗi môn học
class Guigiangvien:

    def __init__(self,window):

        self.window = window
        self.path = os.getcwd() + "\\alpha.xlsx"

        self.window.title('Lịch giảng dạy')

        self.frame = ttk.Frame(self.window)
        self.frame.pack()

        # Khung thu 1

        self.style = ttk.Style()
        self.style.configure("My.TLabelframe.Label", font=("Helvetica", 12))
        self.style.configure("My.TButton", font=("Helvetica", 12))

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Thông tin giảng viên", style="My.TLabelframe")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)

        self.status_combobox = ttk.Combobox(self.widgets_frame, values=self.rowfirst(), font=("Helvetica", 12))
        self.status_combobox.current(0)
        self.status_combobox.grid(row=0, column=0, padx=5, pady=5,  sticky="ew")

        self.status_combobox1 = ttk.Combobox(self.widgets_frame, values=self.colfirst(), font=("Helvetica", 12))
        self.status_combobox1.current(0)
        self.status_combobox1.grid(row=1, column=0, padx=5, pady=5,  sticky="ew")

        self.name_entry = ttk.Entry(self.widgets_frame, font=("Helvetica", 12))
        self.name_entry.insert(0,"Alpha")
        self.name_entry.bind("<FocusIn>", lambda e: self.name_entry.delete('0', 'end'))
        self.name_entry.grid(row=2,column=0,sticky='ew')

        self.button1 = ttk.Button(self.widgets_frame, text="Sửa", command=self.sua, style="My.TButton")
        self.button1.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

        self.separator = ttk.Separator(self.widgets_frame)
        self.separator.grid(row=4, column=0, padx=(20, 10), pady=10, sticky="ew")

        self.button1 = ttk.Button(self.widgets_frame, text="Kiểm tra", command=self.kiemtra, style="My.TButton")
        self.button1.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")


        # Khung thu 2
        self.treeFrame = ttk.Frame(self.frame)
        self.treeFrame.grid(row=0, column=1, pady=10)
        self.treeScroll = ttk.Scrollbar(self.treeFrame)
        self.treeScroll.pack(side="right", fill="y")

        self.cols = self.rowfirst()
        self.treeview = ttk.Treeview(self.treeFrame, show="headings",yscrollcommand=self.treeScroll.set, columns=self.cols, height=20)
        for i in self.cols:
            if i == "Môn":
                self.treeview.column(i, width=250)
            else:
                self.treeview.column(i, width=100)

        self.treeview.pack()
        self.treeScroll.config(command=self.treeview.yview) 


        def select(event):
            indexrow = 0
            for i in self.treeview.selection():
                indexrow = self.treeview.index(i)
            self.status_combobox1.current(indexrow)

        def on_treeview_cell_select(event):
            region = self.treeview.identify_region(event.x, event.y)
            selected_item = self.treeview.selection()
            if region == "cell":
                selected_column = self.treeview.identify_column(event.x)  
                selected_column_id = int(selected_column[1:]) - 1
                selected_value = self.treeview.item(selected_item[0], "values")[int(selected_column[1:]) - 1]
                self.status_combobox.current(selected_column_id)
                if (selected_value.isnumeric()):
                    self.name_entry.delete('0','end')
                    self.name_entry.insert(0,selected_value)
                else:
                    self.name_entry.delete('0','end')  

        self.treeview.bind('<<TreeviewSelect>>',select)
        self.treeview.bind("<ButtonRelease-1>", on_treeview_cell_select)

        self.load_data()

        self.menubar = Menu(window)
        self.window.config(menu = self.menubar)

        # self.fileMenu = Menu(self.menubar, tearoff = 0)
        # self.menubar.add_cascade(label = "File", menu = self.fileMenu)
        # self.fileMenu.add_cascade(label = "Open",command= self.open)
        # self.fileMenu.add_cascade(label = "Edit",command= self.edit)
        # self.fileMenu.add_cascade(label = "Close",command= self.close)

        self.lecturerMenu = Menu(self.menubar, tearoff = 0)
        self.menubar.add_cascade(label = "Giảng viên", menu = self.lecturerMenu)
        self.lecturerMenu.add_cascade(label = "Thông tin", command= self.thongtin)
        self.lecturerMenu.add_cascade(label = "Thêm giảng viên", command= self.addgiangvien)
        self.lecturerMenu.add_cascade(label = "Thêm môn học", command= self.addmonhoc)

    def open():
        print("open file")

    def edit():
        print("edit")

    def close(self, window):
        self.window = window
        window.destroy()

    def thongtin(self):
        print(len(self.rowfirst()))

    def addgiangvien(self):
        root = tk.Tk()
        addlec.Addlec(root,self.window,Guigiangvien)

    def addmonhoc(self):
        root = tk.Tk()
        addmon.Addmonhoc(root,self.window,Guigiangvien)

    def colfirst(self):
        col = []
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook.active
        data = list(sheet.values)
        for i in data[1:]:
            col.append(i[0])
        return col

    def rowfirst(self):
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook.active
        row = list(sheet.values)
        return row[0]

    def load_data(self):
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook.active

        list_values = list(sheet.values)
        for col_name in list_values[0]:
            self.treeview.heading(col_name, text=col_name)

        for value in list_values[1:]:
            value_list = list(value)
            for i in range(len(value_list)):
                if value_list[i] == None:
                    value_list[i] = '-'
            self.treeview.insert('', tk.END, values=value_list)

    def sua(self):  
        alpha = self.name_entry.get()
        giangvien = self.status_combobox.get()
        monhoc = self.status_combobox1.get()
        if alpha.isnumeric() or alpha == '':
            workbook = openpyxl.load_workbook(self.path)
            sheet = workbook.active
            sheet.cell(column = self.rowfirst().index(giangvien) + 1, row = self.colfirst().index(monhoc) + 2,value = alpha)
            workbook.save(self.path)
            self.treeview.delete(*self.treeview.get_children())
            self.load_data()
        else:
            messagebox.showerror(title = "Lỗi",message = "Vui lòng nhập số")
    
    def kiemtra(self):
        listalpha = ra.listalpha()
        for tenmon in listalpha:
            tong = 0
            listpoint = list(listalpha[tenmon].values())
            for i in listpoint:
                tong = tong + int(i)
            if tong != 12:
                messagebox.showwarning(title = "Chú ý",message = tenmon + ": tổng hệ số alpha chưa = 12")
                break

    def getname():
        return Guigiangvien.__name__

# if __name__ == "__main__":
#     window = Tk()
#     Guigiangvien(window)
#     window.mainloop()