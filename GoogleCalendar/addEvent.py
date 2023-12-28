import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox, filedialog, Menu, Tk, Toplevel, Label, simpledialog
from GoogleCalendar.CalendarFunc import CalendarFunc
from GoogleCalendar.commonVar import commonVar
import GoogleCalendar.commonVar as com
from GoogleCalendar.lecturer import lecturer

class GuiAddEvent:
    
    def __init__(self, root):
        
        self.root = root
        self.IDDICT = CalendarFunc.get_calendar_id()
        self.lecNameList = self.IDDICT.keys()
        self.checkbox_vars = []

        self.root.title("Thêm sự kiện vào lịch giảng viên")
        self.root.geometry(f'400x600')

        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        style = ttk.Style(self.root)
        style.configure("My.TLabelframe.Label", font=("Helvetica", 13))
        style.configure("TCheckbutton", font=("Helvetica", 13))
        style.configure("Custom.TButton", font=("Helvetica", 13))

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Cán bộ giảng dạy",style="My.TLabelframe")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)

        self.select_all_var = tk.IntVar()
        self.select_all_checkbox = ttk.Checkbutton(self.widgets_frame, text="Select All", variable=self.select_all_var, command=self.toggle_select_all)
        self.select_all_checkbox.pack(anchor=tk.W)
        
        for lec in self.lecNameList:
            var = tk.IntVar()
            self.checkbox = ttk.Checkbutton(self.widgets_frame, text=lec, variable= var)
            self.checkbox.pack(anchor=tk.W)
            self.checkbox_vars.append(var)
            
        self.show_button = ttk.Button(self.widgets_frame, text="Thêm sự kiện", command=self.addEvent, style="Custom.TButton")
        self.show_button.pack(pady=10)     
            
    def toggle_select_all(self):
        select_all_state = self.select_all_var.get()
        for var in self.checkbox_vars:
            var.set(select_all_state) 
                   
    def addEvent(self):
        selected_lec = [lec for lec, var in zip(self.lecNameList, self.checkbox_vars) if var.get() == 1]
        if len(selected_lec) > 0:
            IDDict = self.IDDICT
            lecList = commonVar.get_lecList(com.common)
            firstMon = commonVar.get_firstMon(com.common)
            addList = []
            for lecName in selected_lec:
                for lec in lecList:
                    if lec.Name == lecName:
                        addList.append(lec)
            
            for i in range (len(addList)):
                totalElement = len(addList[i].Schedule)
                for j in range(totalElement):
                    tempData = lecturer.getEventData(i, j, addList)
                    
                    eData = list(tempData)
                    tmpStages = CalendarFunc.getStages(int(eData[2]), eData[3], firstMon)
                    stages = [element for row in tmpStages for element in row] 
                    
                    for k in range(0, len(stages), 2):
                        date_obj0 = stages[k]
                        date_obj1 = stages[k+1]
                        s_date_str = date_obj0.strftime("%Y-%m-%d")
                        e_date_str = date_obj1.strftime("%Y%m%d")
                        eData[5] = s_date_str
                        eData[6] = e_date_str
                        ev = CalendarFunc.eventCreate(*eData)   
                                
            self.root.destroy()
        else:
            messagebox.showwarning(title="Chú ý",message="Vui lòng chọn cán bộ giảng dạy",parent = self.root)