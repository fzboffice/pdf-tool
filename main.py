from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter, PdfFileReader
from icon import img
import base64
import os
import math
from merge import Merge
from split import Split


class Application():
    def __init__(self):
        self.root = Tk()
        tmp = open('tmp.ico', 'wb+')
        tmp.write(base64.b64decode(img))
        tmp.close()
        self.root.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        self.root.iconbitmap()
        self.root.title("pdf tool v1.0.0")
        self.root.resizable(0, 0)
        main_width, main_height = 800, 570
        scr_width, scr_height = self.root.winfo_screenwidth(
        ), self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d' % (main_width, main_height,
                                            (scr_width-main_width)/2, (scr_height-main_height)/2))
        # 界面分为两个frame
        # 1.功能选择框架
        self.selFrm = Frame(self.root)
        self.selFrm.pack(side='top', fill='x')
        # 合并按钮
        self.v = IntVar()
        self.v.set(1)
        self.selBtn_merge = Radiobutton(
            self.selFrm, text='合并pdf', width=15, height=5, command=self.mergeBtnEvent, variable=self.v, value=1, indicatoron=0)
        self.selBtn_merge.pack(side='left', expand='yes', pady=30)
        # 拆分按钮
        self.selBtn_split = Radiobutton(
            self.selFrm, text='拆分pdf', width=15, height=5, command=self.splitBtnEvent, variable=self.v, value=2, indicatoron=0)
        self.selBtn_split.pack(side='left', expand='yes', pady=30)
        # 插入按钮
        self.selBtn_add = Radiobutton(
            self.selFrm, text='插入pdf', width=15, height=5, command=self.insertBtnEvent, variable=self.v, value=3, indicatoron=0)
        self.selBtn_add.pack(side='left', expand='yes', pady=30)
        # 删除按钮
        self.selBtn_del = Radiobutton(
            self.selFrm, text='删除pdf', width=15, height=5, command=self.delBtnEvent, variable=self.v, value=4, indicatoron=0)
        self.selBtn_del.pack(side='left', expand='yes', pady=30)

        # 2.功能框架
        self.workFrm = Frame(self.root)
        self.workFrm.pack()

        # 3.进度框架
        self.prgFrm = Frame(self.root)
        self.prgFrm.pack(side='bottom')

        self.canvas = Canvas(self.prgFrm, width=800, height=22, bg="white")
        self.canvas.pack()
        self.prgsRectagl = self.canvas.create_rectangle(
            0, 0, 0, 0, width=0, fill="green")

        self.selBtn_merge.invoke()

    # 设置进度条
    def progress(self, x):
        self.canvas.coords(self.prgsRectagl, (0, 0, 800 * x/100, 22))
        self.root.update()

    # 允许所有的按钮
    def enable(self, w, state_text='normal'):
        def cstate(widget):
            if widget.winfo_children:
                for w in widget.winfo_children():
                    if isinstance(w, Button):
                        w.config(state=state_text)
                    cstate(w)
        cstate(w)
    # 禁用所有的按钮

    def disable(self, w):
        self.enable(w, 'disabled')

    # 合并功能选择按钮事件
    def mergeBtnEvent(self):
        for widget in self.workFrm.winfo_children():
            widget.destroy()
        Merge(self)

    # 分割功能选择按钮事件
    def splitBtnEvent(self):
        for widget in self.workFrm.winfo_children():
            widget.destroy()
        Split(self)

    # 插入功能选择按钮事件
    def insertBtnEvent(self):
        for widget in self.workFrm.winfo_children():
            widget.destroy()

    # 删除功能选择按钮事件
    def delBtnEvent(self):
        for widget in self.workFrm.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    main = Application()
    main.root.mainloop()
