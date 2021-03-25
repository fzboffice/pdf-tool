from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileMerger
from icon import img
import base64
import os

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
        main_width, main_height = 800, 600
        scr_width, scr_height = self.root.winfo_screenwidth(
        ), self.root.winfo_screenheight()
        self.root.geometry('%dx%d+%d+%d' % (main_width, main_height,
                                            (scr_width-main_width)/2, (scr_height-main_height)/2))
        # 界面分为两个frame
        # 1.功能选择框架
        selFrm = Frame(self.root)
        selFrm.pack(side='top', fill='x')
        # 合并按钮
        self.selbtn_merge = Button(
            selFrm, text='合并pdf', width=15, height=5, command=self.selbtn_merge_event)
        self.selbtn_merge.pack(side='left', expand='yes', pady=30)
        # 拆分按钮
        self.selbtn_split = Button(
            selFrm, text='拆分pdf', width=15, height=5, command=self.selbtn_merge_event)
        self.selbtn_split.pack(side='left', expand='yes', pady=30)
        # 插入按钮
        self.selbtn_add = Button(
            selFrm, text='插入pdf', width=15, height=5, command=self.selbtn_merge_event)
        self.selbtn_add.pack(side='left', expand='yes', pady=30)
        # 删除按钮
        self.selbtn_del = Button(
            selFrm, text='删除pdf', width=15, height=5, command=self.selbtn_merge_event)
        self.selbtn_del.pack(side='left', expand='yes', pady=30)

        # 2.功能框架
        self.workfrm = Frame(self.root)
        self.workfrm.pack(side='top')

        self.selbtn_merge_event()

    def selbtn_merge_event(self):
        #print('merge press')
        self.selbtn_merge.config(relief='sunken')
        self.selbtn_split.config(relief='raised')
        self.selbtn_add.config(relief='raised')
        self.selbtn_del.config(relief='raised')
        for widget in self.workfrm.winfo_children():
            widget.destroy()
        # list
        self.fileList = Listbox(
            self.workfrm, width=82, height=20, selectmode='extended')
        self.fileList.pack(side='left')
        self.fileList.bind('<Button-1>', self.show)
        self.fileList.bind('<B1-Motion>', self.showInfo)
        # scrollbar
        sc = Scrollbar(self.workfrm, command=self.fileList.yview)
        sc.pack(side=LEFT, fill=Y)
        self.fileList.config(yscrollcommand=sc.set)

        # 按钮
        self.fileBtnAdd = Button(
            self.workfrm, text='添加文件', width=15, height=5, command=self.fileBtnAdd_event)
        self.fileBtnAdd.pack(side='top', expand='yes')
        self.fileBtnMov = Button(
            self.workfrm, text='移除文件', width=15, height=5, command=self.fileBtnMov_event)
        self.fileBtnMov.pack(side='top', expand='yes')
        self.fileBtnExe = Button(
            self.workfrm, text='执行合并', width=15, height=5, command=self.fileBtnExe_event)
        self.fileBtnExe.pack(side='top', expand='yes')

    def fileBtnAdd_event(self):
        filenames = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')])
        if len(filenames) == 0:
            return
        for item in filenames:
            self.fileList.insert(END, item)

    def fileBtnMov_event(self):
        indices = self.fileList.curselection()
        if len(indices) == 0:
            return
        for index in reversed(indices):
            self.fileList.delete(index)

    def fileBtnExe_event(self):
        if self.fileList.size() <= 1:
            return
        file_merger = PdfFileMerger()
        for pdf in self.fileList.get(0, END):
            file_merger.append(pdf)
        filename = filedialog.asksaveasfilename(
            initialfile='合并文件', defaultextension=".pdf", filetypes=[('PDF file', '*.pdf')])
        if len(filename) == 0:
            return
        file_merger.write(filename)
        messagebox.showinfo('提示', '合并已完成')

    def show(self, event):
        # nearest可以传回最接近y坐标在Listbox的索引
        # 传回目前选项的索引
        self.fileList.index = self.fileList.nearest(event.y)

    def showInfo(self, event):
        # 获取目前选项的新索引
        newIndex = self.fileList.nearest(event.y)
        # 判断，如果向上拖拽
        if newIndex < self.fileList.index:
            # 获取新位置的内容
            x = self.fileList.get(newIndex)
            # 删除新内容
            self.fileList.delete(newIndex)
            # 将新内容插入，相当于插入我们移动后的位置
            self.fileList.insert(newIndex + 1, x)
            # 把需要移动的索引值变成我们所希望的索引，达到了移动的目的
            self.fileList.index = newIndex
        elif newIndex > self.fileList.index:
            # 获取新位置的内容
            x = self.fileList.get(newIndex)
            # 删除新内容
            self.fileList.delete(newIndex)
            # 将新内容插入，相当于插入我们移动后的位置
            self.fileList.insert(newIndex - 1, x)
            # 把需要移动的索引值变成我们所希望的索引，达到了移动的目的
            self.fileList.index = newIndex


main = Application()
main.root.mainloop()
