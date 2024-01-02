import os
import tkinter as tk
from tkinter import ttk,messagebox,filedialog,Menu,Tk
import manager as mn
from GoogleCalendar.lecturer import lecturer
from GoogleCalendar.commonVar import commonVar
import GoogleCalendar.commonVar as com
from GoogleCalendar.CalendarFunc import CalendarFunc
from GoogleCalendar.CreatCal import GuiCreatCal
from GoogleCalendar.RemoteCal import GuiRemoveCalendar
from GoogleCalendar.addEvent import GuiAddEvent
from GoogleCalendar.ShareLink import GuiShareCal

class GoogleCalendar:
    def __init__(self, window: Tk, calendar_list):

        self.fullSche = None
        self.lecNameList = None
        self.lecList = None
        self.list_ = calendar_list
        
        self.window = window
        window.title("Google Calendar")
        self.window.geometry("400x300+150+100")

        self.style = ttk.Style(self.window)

        # self.window.tk.call("source", "theme-light.tcl")
        # self.style.theme_use("theme-light")

        self.style.configure("My.TLabelframe.Label", font=("Helvetica", 13))
        self.style.configure("Custom.TButton", font=("Helvetica", 13))
        
        self.frame = ttk.Frame(self.window)
        self.frame.pack()

        self.widgets_frame = ttk.LabelFrame(self.frame, text= "Tác vụ",style="My.TLabelframe")
        self.widgets_frame.grid(row=0,column=0, padx=20, pady=10)
            
        # self.separator = ttk.Separator(self.widgets_frame)
        # self.separator.grid(row=1, column=0, padx=(20, 10), pady=10, sticky="ew")

        self.button1 = ttk.Button(self.widgets_frame, text="Tạo lịch", command=self.calCreate, style="Custom.TButton")
        self.button1.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
        
        self.button2 = ttk.Button(self.widgets_frame, text="Xóa lịch", command=self.calRemove, style="Custom.TButton")
        self.button2.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

        self.button3 = ttk.Button(self.widgets_frame, text="Thêm sự kiện vào lịch", command=self.addEvent, style="Custom.TButton")
        self.button3.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")
        
        self.button4 = ttk.Button(self.widgets_frame, text="Lấy link chia sẻ lịch", command=self.getShareableLink, style="Custom.TButton")
        self.button4.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")
        
        self.menubar = Menu(self.window)
        self.window.config(menu = self.menubar)

        self.fileMenu = Menu(self.menubar, tearoff = 0,font = ("Helvetica",13))
        self.menubar.add_cascade(label = "File", menu = self.fileMenu)
        self.fileMenu.add_cascade(label = "Cập nhật lịch đã sắp xếp",command=self.updateSche)
        self.fileMenu.add_cascade(label = "Open File Excel",command= self.open)

        self.menuacc = Menu(self.menubar, tearoff = 0,font = ("Helvetica",13))
        self.menubar.add_cascade(label = "Tài khoản", menu = self.menuacc)
        self.menuacc.add_cascade(label = "Đăng nhập",command= self.login)
        self.menuacc.add_cascade(label = "Đăng xuất",command= self.logout)

        if CalendarFunc.checklogin():
            self.solich = len(CalendarFunc.get_calendar_id())
        else:
            self.solich = 0

        firstMon = mn.get_first_monday(self.list_)
        lecNameList = lecturer.getLecNameList(self.list_)
        lecList = lecturer.lecCre(lecNameList)
        lecturer.scheAdd(lecList, self.list_)
        commonVar.Set_fullSche(com.common, self.list_)
        commonVar.Set_lecNameList(com.common, lecNameList)
        commonVar.Set_lecList(com.common, lecList)
        commonVar.Set_firstMon(com.common, firstMon)

    def updateSche(self):
        firstMon = mn.get_first_monday(self.list_)
        lecNameList = lecturer.getLecNameList(self.list_)
        lecList = lecturer.lecCre(lecNameList)
        lecturer.scheAdd(lecList, self.list_)
        commonVar.Set_fullSche(com.common, self.list_)
        commonVar.Set_lecNameList(com.common, lecNameList)
        commonVar.Set_lecList(com.common, lecList)
        commonVar.Set_firstMon(com.common, firstMon)

    def open(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls",)])
        if self.file_path:
            print("Thư mục được chọn:", self.file_path)

            try:
                if os.path.splitext(self.file_path)[1] == ".xlsx":
                    self.list_ = mn.readfile(self.file_path, readlec = True) 
                    
                if os.path.splitext(self.file_path)[1] == ".xls":
                    self.list_ = mn.readfilexls(self.file_path, readlec = True)
            except:
                messagebox.showwarning(title="Chú ý",message="File Excel không đúng định dạng",parent = self.window)
                return

            firstMon = mn.startday(self.list_)
            lecNameList = lecturer.getLecNameList(self.list_)
            lecList = lecturer.lecCre(lecNameList)
            lecturer.scheAdd(lecList, self.list_)
            commonVar.Set_fullSche(com.common, self.list_)
            commonVar.Set_lecNameList(com.common, lecNameList)
            commonVar.Set_lecList(com.common, lecList)
            commonVar.Set_firstMon(com.common, firstMon)
            self.cansave = True

        else:
            print("Không có thư mục nào được chọn.")

    def login(self):
        creds = CalendarFunc.credsCheck()
        if creds:
            self.solich = len(CalendarFunc.get_calendar_id())
            messagebox.showinfo("Thông báo", "Đã đăng nhập!",parent = self.window)

    def logout(self):
        if not CalendarFunc.checklogin():
            messagebox.showinfo("Thông báo","Chưa đăng nhập",parent = self.window)
            return
        
        result = messagebox.askokcancel("Đăng xuất", "Bạn có muốn đăng xuất?",parent = self.window)
        if result:
            credRemove = CalendarFunc.credRemove()
            if credRemove:
                messagebox.showerror("Lỗi", "Lỗi",parent = self.window)
            else:
                messagebox.showinfo("Thông báo","Đăng xuất thành công!",parent = self.window)

    def calCreate(self):
        if not CalendarFunc.checklogin():
            messagebox.showinfo("Thông báo","Vui lòng đăng nhập",parent = self.window)
            return
        if (self.list_):
            root = tk.Toplevel(self.window)
            GuiCreatCal(root)
        else:
            messagebox.showerror(title="Lỗi",message="Không có thông tin lịch dạy vui lòng thêm lịch dạy",parent = self.window)

    def calRemove(self):
        if not CalendarFunc.checklogin():
            messagebox.showinfo("Thông báo","Vui lòng đăng nhập",parent = self.window)
            return
        if self.solich > 0: 
            root = tk.Toplevel(self.window)
            GuiRemoveCalendar(root)
        else:
            messagebox.showerror(title="Lỗi",message="Không có thông tin cán bộ giảng dạy vui lòng thêm tạo lịch trước",parent = self.window)

    def addEvent(self):
        if not CalendarFunc.checklogin():
            messagebox.showinfo("Thông báo","Vui lòng đăng nhập",parent = self.window)
            return
        if self.solich > 0:
            root = tk.Toplevel(self.window)
            GuiAddEvent(root)
        else:
            messagebox.showerror(title="Lỗi",message="Không có thông tin cán bộ giảng dạy vui lòng thêm tạo lịch trước",parent = self.window)

    def getShareableLink(self):
        if not CalendarFunc.checklogin():
            messagebox.showinfo("Thông báo","Vui lòng đăng nhập",parent = self.window)
            return
        if self.solich > 0:
            root = tk.Toplevel(self.window)
            GuiShareCal(root)
        else:
            messagebox.showerror(title="Lỗi",message="Không có thông tin cán bộ giảng dạy vui lòng thêm tạo lịch trước",parent = self.window)

# if __name__ == "__main__":
#     window = Tk()
#     GoogleCalendar(window, [])
#     window.mainloop()
