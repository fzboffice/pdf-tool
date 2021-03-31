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
from insert import Insert
from delete import Delete

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
        self.main_width, self.main_height = 800, 570
        scr_width, scr_height = self.root.winfo_screenwidth(
        ), self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d' % (self.main_width, self.main_height,
                                            (scr_width-self.main_width)/2, (scr_height-self.main_height)/2))
        # 界面分为两个frame
        # 1.功能选择框架
        self.selFrm = Frame(self.root)
        self.selFrm.pack(side='top', fill='x')
        # 功能选择按钮
        self.selV = IntVar()
        self.selV.set(1)
        # 合并按钮
        selBtn_merge = Radiobutton(self.selFrm, text='合并pdf', width=15, height=5, command=self.selBtnEvent,
                                   variable=self.selV, value=1, indicatoron=0)
        selBtn_merge.pack(side='left', expand='yes', pady=30)
        # 拆分按钮
        Radiobutton(self.selFrm, text='拆分pdf', width=15, height=5, command=self.selBtnEvent,
                    variable=self.selV, value=2, indicatoron=0).pack(side='left', expand='yes', pady=30)
        # 插入按钮
        Radiobutton(self.selFrm, text='插入pdf', width=15, height=5, command=self.selBtnEvent,
                    variable=self.selV, value=3, indicatoron=0).pack(side='left', expand='yes', pady=30)
        # 删除按钮
        Radiobutton(self.selFrm, text='删除pdf', width=15, height=5, command=self.selBtnEvent,
                    variable=self.selV, value=4, indicatoron=0).pack(side='left', expand='yes', pady=30)

        # 2.功能框架
        self.workFrm = Frame(self.root)
        self.workFrm.pack()

        # 3.进度框架
        self.prgFrm = Frame(self.root)
        self.prgFrm.pack(side='bottom')

        self.canvas = Canvas(self.prgFrm, width=self.main_width,
                             height=22, bg="white")
        self.canvas.pack()
        self.prgsRectagl = self.canvas.create_rectangle(
            0, 0, 0, 0, width=0, fill="green")

        # 初始功能
        selBtn_merge.invoke()
    # 设置进度条

    def progress(self, x):
        self.canvas.coords(
            self.prgsRectagl, (0, 0, self.main_width * x/100, 22))
        self.root.update()

    # 允许交互
    def enable(self, widget, state_text='normal'):
        def cstate(widget):
            if widget.winfo_children:
                for w in widget.winfo_children():
                    if isinstance(w, Button) or isinstance(w, Radiobutton) or isinstance(w, Entry):
                        w.config(state=state_text)
                    cstate(w)
        cstate(widget)

    # 禁用交互
    def disable(self, widget):
        self.enable(widget, 'disabled')

    # 功能选择按钮事件
    def selBtnEvent(self):
        for widget in self.workFrm.winfo_children():
            widget.destroy()
        if self.selV.get() == 1:
            Merge(self)
        if self.selV.get() == 2:
            Split(self)
        if self.selV.get() == 3:
            Insert(self)
        if self.selV.get() == 4:
            Delete(self)


if __name__ == "__main__":
    main = Application()
    main.root.mainloop()
