import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox, filedialog, Menu, Tk, Toplevel, Label, simpledialog
from GoogleCalendar.CalendarFunc import CalendarFunc

class GuiShareCal:
    
    def __init__(self, root):
        
        self.root = root

        self.root.title ("Lấy link chia sẻ lịch")
        self.root.geometry("400x200")
        
        self.frame = ttk.Frame(self.root)
        self.frame.pack()

        style = ttk.Style(self.root)
        style.configure("My.TLabelframe.Label", font=("Helvetica", 13))
        style.configure("Custom.TButton", font=("Helvetica", 13))

        self.IDDICT = CalendarFunc.get_calendar_id()
        self.lecNameList = list(self.IDDICT)
        
        ttk.Label(self.frame, text = "Chọn lịch :",font = ("Helvetica", 13)).grid(column = 0,row = 15, padx = 10, pady = 25) 
        n = tk.StringVar() 
        self.cbb = ttk.Combobox(self.frame, values= self.lecNameList, width = 27, textvariable = n, font=("Helvetica", 13))
        self.cbb.grid(column = 1, row = 15) 
        
        self.show_button = ttk.Button(self.frame, text="Lấy lịch", command=self.immigrate, style="Custom.TButton")
        self.show_button.grid(column= 1, row= 17)   

        
    def immigrate(self):
        selected = self.cbb.get()
        calID = self.IDDICT[selected]
        email_input = simpledialog.askstring('Nhập email', 'email')
        if email_input:
            result = CalendarFunc.invite(calID, email_input)
            if result == 'successful':
                messagebox.showinfo('Thành công', f'Đã gửi link lịch tới email: {email_input}',parent = self.root)
                self.window.destroy()
            else:
                messagebox.showerror('Lỗi', result)