import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from GoogleCalendar.CalendarFunc import CalendarFunc
from GoogleCalendar.commonVar import commonVar
import GoogleCalendar.commonVar as com
from concurrent.futures import ThreadPoolExecutor

class GuiCreatCal:
    
    def __init__(self, root: Tk):

        self.created_list = []
        self.root = root
        self.lecNameList = commonVar.get_lecNameList(com.common)
        self.checkbox_vars = []

        self.root.title("Tạo lịch cho cán bộ giảng dạy")
        self.root.geometry(f'400x600')

        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        style = ttk.Style(self.root)
        style.configure("TCheckbutton", font=("Helvetica", 13))
        style.configure("Custom.TButton", font=("Helvetica", 13))

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Cán bộ giảng dạy                    ",style="My.TLabelframe")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)

        self.select_all_var = tk.IntVar()
        self.select_all_checkbox = ttk.Checkbutton(self.widgets_frame, text="Select All", variable=self.select_all_var, command=self.toggle_select_all)

        self.select_all_checkbox.pack(anchor=tk.W)
        
        for lec in self.lecNameList:
            var = tk.IntVar()
            self.checkbox = ttk.Checkbutton(self.widgets_frame, text=lec, variable= var)
            self.checkbox.pack(anchor=tk.W)
            self.checkbox_vars.append(var)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.widgets_frame, variable=self.progress_var, maximum=100,length=200)
        self.progress_bar.pack()
            
        self.show_button = ttk.Button(self.widgets_frame, text="Tạo lịch", command=self.start_processing, style="Custom.TButton")
        self.show_button.pack(pady=10)     

        self.executor = ThreadPoolExecutor(max_workers=4) 

    def toggle_select_all(self):
            select_all_state = self.select_all_var.get()
            for var in self.checkbox_vars:
                var.set(select_all_state)

    def create(self):
        selected_lec = [lec for lec, var in zip(self.lecNameList, self.checkbox_vars) if var.get() == 1]
        index = 0
        if len(selected_lec) > 0:
            IDDict = CalendarFunc.get_calendar_id()
            if type(IDDict):
                lec = set(selected_lec)
                uncommon = [item for item in lec if item not in IDDict]
                newID = {}
                for lec in uncommon:
                    ID = CalendarFunc.newCalendar(lec)

                    index = index + 1
                    process = (index / len(uncommon)) * 100
                    self.progress_var.set(process)  
                    self.root.update_idletasks()

                messagebox.showinfo("Thông báo", f"Tạo lịch thành công cho: {uncommon}",parent = self.root)
                self.root.destroy()
            else:
                CalendarFunc.CreateIDFile()
                self.create()
        else:
            messagebox.showwarning(title="Chú ý",message="Vui lòng chọn cán bộ giảng dạy",parent = self.root)

    def start_processing(self):
        self.executor.submit(self.create)