import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from GoogleCalendar.CalendarFunc import CalendarFunc
from concurrent.futures import ThreadPoolExecutor

class GuiRemoveCalendar:
    
    def __init__(self, root: Tk):

        self.id_list = {}
        self.lecNameList = []
        self.root = root
        calendar_list = CalendarFunc.get_calendar_list()

        self.root.title("Xóa lịch của cán bộ giảng dạy")
        self.root.geometry(f'400x600')

        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        style = ttk.Style(self.root)
        style.configure("My.TLabelframe.Label", font=("Helvetica", 13))
        style.configure("TCheckbutton", font=("Helvetica", 13))
        style.configure("Custom.TButton", font=("Helvetica", 13))

        for calendar_list_entry in calendar_list['items']:
            self.lecNameList.append(calendar_list_entry['summary'])
            self.id_list.update({calendar_list_entry['summary']:calendar_list_entry['id']})

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Cán bộ giảng dạy                    ",style="My.TLabelframe")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)
            
        self.checkbox_vars = []
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
            
        self.show_button = ttk.Button(self.widgets_frame, text="Xóa lịch", command=self.start_processing, style="Custom.TButton")
        self.show_button.pack(pady=10)   

        self.executor = ThreadPoolExecutor(max_workers=4) 

    def toggle_select_all(self):
        select_all_state = self.select_all_var.get()
        for var in self.checkbox_vars:
            var.set(select_all_state)
    
    def getID(dict, keys):
        return [dict[key] for key in keys if key in dict]
    
    def getKey(dict, target_value):
        for key, value in dict.items():
            if value == target_value:
                return key
        return None
    
    def remove(self):
        selected_lec = [lec for lec, var in zip(self.lecNameList, self.checkbox_vars) if var.get() == 1]
        index = 0
        if len(selected_lec) > 0:
            IDDict = self.id_list
            confirm = messagebox.askokcancel("Xác nhận", f"Xác nhận xóa lịch của: {selected_lec}",parent = self.root)
            if confirm:
                if IDDict:
                    IDList = GuiRemoveCalendar.getID(IDDict, selected_lec)
                    print(IDList)
                    removedList = []
                    for ID in IDList:
                        removedID = CalendarFunc.calendarRemove(ID)
                        removedCale = GuiRemoveCalendar.getKey(IDDict, removedID)
                        removedList.append(removedCale)

                        index = index + 1
                        process = (index / len(IDList)) * 100
                        self.progress_var.set(process)  
                        self.root.update_idletasks()

                    messagebox.showinfo("Xóa thành công", f"Danh sách xóa: {removedList}",parent = self.root)
                    self.root.destroy()      
                else:
                    return
            else:
                return
        else:
            messagebox.showwarning(title="Chú ý",message="Vui lòng chọn cán bộ giảng dạy",parent = self.root)

    def start_processing(self):
        self.executor.submit(self.remove)