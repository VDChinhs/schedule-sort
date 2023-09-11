from tkinter import *
from tkinter import ttk,messagebox,filedialog
from openpyxl.styles import NamedStyle
import tkinter as tk
import openpyxl
import addlec
import addmon
import manager as mn
import Guilec
import os


#GUI hiện thị lịch
class Guische:
    def __init__(self,window):
        self.cansave = False
        self.path = os.getcwd() + "\\alpha.xlsx"
        self.checktrung = FALSE
        self.list_copy = []

        self.window = window
        self.window.title('Lịch giảng dạy')
    
        self.frame = ttk.Frame(self.window)
        self.frame.pack()

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Xử lý")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)

        self.name_entry = ttk.Entry(self.widgets_frame, font=("Helvetica", 12))
        self.name_entry.insert(0,"Alpha")
        # self.name_entry.bind("<FocusIn>", lambda e: self.name_entry.delete('0', 'end'))
        self.name_entry.grid(row=2,column=0,sticky='ew')

        self.button1 = ttk.Button(self.widgets_frame, text="Sửa", command=self.sua)
        self.button1.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

        self.button1 = ttk.Button(self.widgets_frame, text="Xếp giảng viên", command=self.xepgiangvien)
        self.button1.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

        self.button1 = ttk.Button(self.widgets_frame, text="Kiểm tra lớp trùng", command=self.checkloptrung)
        self.button1.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")

        self.button1 = ttk.Button(self.widgets_frame, text="Xuất File Excel", command=self.save)
        self.button1.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")

        self.treeFrame = ttk.Frame(self.frame)
        self.treeFrame.grid(row=0, column=1, pady=10)
        self.treeScroll = ttk.Scrollbar(self.treeFrame)
        self.treeScroll.pack(side="right", fill="y")

        self.cols = ["STT","Mã học phần","Tên môn học", "Mã Lớp học phần", "Thứ","Tiết","Tuần","Giảng viên","Ghép","Trùng"]
        self.treeview = ttk.Treeview(self.treeFrame, show="headings",yscrollcommand=self.treeScroll.set, columns=self.cols, height=35)
        for i in self.cols:
            if i == "Tên môn học":
                self.treeview.column(i, width=270)
            elif i == "Tuần":
                self.treeview.column(i, width=140)
            else:
                self.treeview.column(i, width=100)

        self.treeview.pack()
        self.treeScroll.config(command=self.treeview.yview) 

        def on_treeview_cell_select(event):
            region = self.treeview.identify_region(event.x, event.y)
            selected_item = self.treeview.selection()
            if region == "cell":
                selected_column = self.treeview.identify_column(event.x)  
                selected_value = self.treeview.item(selected_item[0], "values")[int(selected_column[1:]) - 1]
                self.name_entry.delete('0','end')
                self.name_entry.insert(0,selected_value)

        self.treeview.bind("<ButtonRelease-1>", on_treeview_cell_select)

        # self.load_data()

        self.menubar = Menu(window)
        self.window.config(menu = self.menubar)

        self.fileMenu = Menu(self.menubar, tearoff = 0)
        self.menubar.add_cascade(label = "File", menu = self.fileMenu)
        self.fileMenu.add_cascade(label = "Open",command= self.open)
        self.fileMenu.add_cascade(label = "Edit",command= self.edit)
        self.fileMenu.add_cascade(label = "Close",command= self.close)

        self.lecturerMenu = Menu(self.menubar, tearoff = 0)
        self.menubar.add_cascade(label = "Giảng viên", menu = self.lecturerMenu)
        self.lecturerMenu.add_cascade(label = "Thông tin", command= self.thongtin)
        self.lecturerMenu.add_cascade(label = "Thêm giảng viên", command= self.addgiangvien)
        self.lecturerMenu.add_cascade(label = "Thêm môn học", command= self.addmonhoc)
    

    def close(self):
        self.window.quit()

    def open(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls",)])
        if self.file_path:
            print("Thư mục được chọn:", self.file_path)
            self.treeview.delete(*self.treeview.get_children())

            self.list_ = mn.readfile(self.file_path)
            for i in self.list_:
                self.list_copy.append(i)

            mn.lopghep(self.list_)
            self.load_data(self.list_)
            self.cansave = True
        else:
            print("Không có thư mục nào được chọn.")

    def edit(self):
        self.treeview.delete(*self.treeview.get_children())

    def thongtin(self):
        if len(self.colfirst()) == 0 and len(self.rowfirst()) == 0:
            messagebox.showwarning(title="Chú ý",message="Không có thông tin giảng viên vui lòng thêm giảng viên và môn học")
        elif len(self.colfirst()) == 0 :
            messagebox.showwarning(title="Chú ý",message="Không có thông tin môn học vui lòng thêm môn học")
        elif len(self.rowfirst()) == 0 :
            messagebox.showwarning(title="Chú ý",message="Không có thông tin giảng viên vui lòng thêm giảng viên")
        else:
            Guilec.Guigiangvien(tk.Tk())

    def addgiangvien(self):
        root = tk.Tk()
        addlec.Addlec(root,self.window,Guische)

    def addmonhoc(self):
        root = tk.Tk()
        addmon.Addmonhoc(root,self.window,Guische)

    def sua(self):
        pass

    def creatfile():
        file_path = os.getcwd() + "\\alpha.xlsx"

        if not os.path.exists(file_path):
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Sheet1"
            workbook.save(file_path)
            workbook.close()

    def xepgiangvien(self):
        try:
            self.treeview.delete(*self.treeview.get_children())
            list_xep = mn.readfile(self.file_path)
            self.checktrung = True
            mn.lopghep(list_xep)
            mn.xepgiangvien(list_xep)
            self.load_data(list_xep)
            self.list_ = list_xep
        except:
            messagebox.showwarning(title="Chú ý",message="Không có thông tin lịch dạy vui lòng thêm lịch dạy")
            return
        messagebox.showinfo(title="Thông báo",message="Xếp giảng viên xong")

    def checkloptrung(self):
        if self.checktrung:
            self.treeview.delete(*self.treeview.get_children())
            mn.checktrung(self.list_)
            self.load_data(self.list_)
        else:
           messagebox.showwarning(title="Chú ý",message="Không có thông tin giảng viên vui lòng thêm lịch dạy") 
           return
        messagebox.showinfo(title="Thông báo",message="Kiểm tra trùng lịch xong")
        
    def save(self):
        if self.cansave:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if file_path:
                # Tạo một workbook và một worksheet
                workbook = openpyxl.Workbook()
                worksheet = workbook.active
                worksheet.title = "Sheet1"
                date_style = NamedStyle(name='date_style')
                date_style.number_format = 'DD-MM-YYYY'
                
                head = ["STT","Mã học phần","Tên môn học","Mã lớp học","Lớp ghép","Thứ","Tiết","Phòng học","Số TC","Bắt đầu","Kết thúc","1234567890123456789012","Giảng viên"]
                worksheet.append(head)
                list_data = mn.list_save(self.list_)
                for i in list_data:
                    worksheet.append(i)
                # Đổi style ngày
                column_J = worksheet['J']
                for cell in column_J:
                    cell.style = date_style
                
                column_K = worksheet['K']
                for cell in column_K:
                    cell.style = date_style
                # Lưu workbook
                workbook.save(file_path)
                workbook.close()
        else:
            messagebox.showwarning(title="Chú ý",message="Không có thông tin lịch dạy vui lòng thêm lịch dạy")
    
    def load_data(self,data):
        # data dạng đối tượng rồi mới chuyển thành dạng list
        list_data = mn.list_print(data)

        for col_name in self.cols:
            self.treeview.heading(col_name, text=col_name)

        for value in list_data:
            value_list = list(value)
            for i in range(len(value_list)):
                if value_list[i] == None:
                    value_list[i] = '-'
            self.treeview.insert('', tk.END, values=value_list)

    def rowfirst(self):
        row = []
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook.active
        data = list(sheet.values)
        if len(data) == 0:
            return row
        else:
            row = data[0]
            return row[1:]
    
    def colfirst(self):
        col = []
        workbook = openpyxl.load_workbook(self.path)
        sheet = workbook.active
        data = list(sheet.values)
        for i in data[1:]:
            col.append(i[0])
        return col
    
    def getname():
        return Guische.__name__

if __name__ == "__main__":
    Guische.creatfile()
    window = Tk()
    Guische(window)
    window.mainloop()