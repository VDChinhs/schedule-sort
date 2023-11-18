import tkinter as tk
from tkinter import Tk,ttk
import manager as mn

class Guianalysis:
    def __init__(self,window,list_data):

        self.list_data = list_data

        self.window = window
        self.window.title('Phân tích lịch trùng')
        self.window.geometry('+200+150')
        self.window.iconbitmap("schedule.ico")

        self.notebook = ttk.Notebook(self.window)

        self.style = ttk.Style(self.window)

        self.window.tk.call("source", "theme-light.tcl")
        self.style.theme_use("theme-light")

        cols = ["Môn","Môn Trùng","Thứ","TIết","Tuần"]
        list_print = mn.list_printtrung(self.list_data)

        for i in list_print:
            self.create_tab(i[0], cols, i[1:])

        self.notebook.pack(expand=1, fill='both')
    
    def create_table(self, tab, columns,data):
        tree = ttk.Treeview(tab, columns=columns, show="headings", height=20)
        self.style.configure("Treeview.Heading", font=("Helvetica", 12))
        tree.tag_configure("my_font", font=("Helvetica", 12))
        for col in columns:
            tree.heading(col, text=col)
            for i in columns:
                if i == "Môn":
                    tree.column(i, width=330)
                elif i == "Môn Trùng":
                    tree.column(i, width=330)
                elif i == "Thứ":
                    tree.column(i, width=50)
                elif i == "Tiết":
                    tree.column(i, width=75)
                elif i == "Tuần":
                    tree.column(i, width=210)
                else:
                    tree.column(i, width=100)

        for row in data:
            tree.insert("", tk.END, values=row, tags=("my_font"))

        tree.pack(expand=1, fill='both')

    def create_tab(self, tab_name, columns,data):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=tab_name)
        self.create_table(tab, columns,data)