from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter, PdfFileReader
import os


class Merge():
    def __init__(self, ui):
        self.ui = ui
        # list
        self.ui.fileList = Listbox(
            self.ui.workFrm, width=82, height=20, selectmode='extended')
        self.ui.fileList.pack(side='left')
        self.ui.fileList.bind('<Button-1>', self.show)
        self.ui.fileList.bind('<B1-Motion>', self.showInfo)
        # scrollbar
        sc = Scrollbar(self.ui.workFrm, command=self.ui.fileList.yview)
        sc.pack(side=LEFT, fill=Y)
        self.ui.fileList.config(yscrollcommand=sc.set)

        # 按钮
        self.ui.fileBtnAdd = Button(
            self.ui.workFrm, text='添加文件', width=12, height=2, command=self.addBtnEvent)
        self.ui.fileBtnAdd.pack(side='top', expand='yes')
        self.ui.fileBtnMov = Button(
            self.ui.workFrm, text='移除文件', width=12, height=2, command=self.delBtnEvent)
        self.ui.fileBtnMov.pack(side='top', expand='yes')
        self.ui.fileBtnExe = Button(
            self.ui.workFrm, text='执行合并', width=12, height=2, command=self.exeBtnEvent)
        self.ui.fileBtnExe.pack(side='top', expand='yes')

    def show(self, event):
        # nearest可以传回最接近y坐标在Listbox的索引
        # 传回目前选项的索引
        self.ui.fileList.index = self.ui.fileList.nearest(event.y)

    def showInfo(self, event):
        # 获取目前选项的新索引
        newIndex = self.ui.fileList.nearest(event.y)
        # 判断，如果向上拖拽
        if newIndex < self.ui.fileList.index:
            # 获取新位置的内容
            x = self.ui.fileList.get(newIndex)
            # 删除新内容
            self.ui.fileList.delete(newIndex)
            # 将新内容插入，相当于插入我们移动后的位置
            self.ui.fileList.insert(newIndex + 1, x)
            # 把需要移动的索引值变成我们所希望的索引，达到了移动的目的
            self.ui.fileList.index = newIndex
        elif newIndex > self.ui.fileList.index:
            # 获取新位置的内容
            x = self.ui.fileList.get(newIndex)
            # 删除新内容
            self.ui.fileList.delete(newIndex)
            # 将新内容插入，相当于插入我们移动后的位置
            self.ui.fileList.insert(newIndex - 1, x)
            # 把需要移动的索引值变成我们所希望的索引，达到了移动的目的
            self.ui.fileList.index = newIndex

    def addBtnEvent(self):
        filenames = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')])
        if len(filenames) == 0:
            return
        for item in filenames:
            self.ui.fileList.insert(END, item)

    def delBtnEvent(self):
        indices = self.ui.fileList.curselection()
        if len(indices) == 0:
            return
        for index in reversed(indices):
            self.ui.fileList.delete(index)

    def exeBtnEvent(self):
        if self.ui.fileList.size() <= 1:
            return
        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        try:
            merger = PdfFileMerger()
            i = 0
            for pdf in self.ui.fileList.get(0, END):
                merger.append(pdf)
                i += 1
                self.ui.progress(100*i/self.ui.fileList.size())
        except:
            messagebox.showerror('错误', '合并失败 pdf文件可能无效')
            self.ui.progress(0)
            self.ui.enable(self.ui.self.uirm)
            self.ui.enable(self.ui.workFrm)
            return
        filename = filedialog.asksaveasfilename(
            initialfile=os.path.splitext(os.path.basename(self.ui.fileList.get(0)))[0]+'_merge', defaultextension=".pdf", filetypes=[('PDF file', '*.pdf')])
        if len(filename) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        merger.write(filename)
        merger.close()
        messagebox.showinfo('提示', '合并已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)
