from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileWriter, PdfFileReader
import os

def isDigit(x):
    try:
        x = int(x)
        return isinstance(x, int)
    except ValueError:
        return False

class Delete():
    def __init__(self,ui):
        self.ui=ui
        Frame1 = Frame(self.ui.workFrm)
        Frame1.pack(side='top')
        self.ui.sel_fileLabel = Label(Frame1, width=60)
        self.ui.sel_fileLabel.pack(side='top')
        self.ui.fileBtnSel = Button(
            Frame1, text='选择pdf', command=self.fileSelBtnEvent)
        self.ui.fileBtnSel.pack(side='top')
        
        self.ui.Label3 = Label(Frame1, text='请按照格式填写 页数或者页数-页数(逗号隔开)')
        self.ui.Label3.pack(side='top')

        self.entry = Entry(Frame1, justify='center')
        self.entry.pack()
        Label(Frame1, text='举例 "1,2-4,5,7-10"').pack(side='top')

        self.ui.exeBtn = Button(
            Frame1, text='执行删除', command=self.exeBtnEvent)
        self.ui.exeBtn.pack(side='top')

    def fileSelBtnEvent(self):
        filename = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')], multiple=False)
        if len(filename) == 0:
            return
        self.ui.sel_fileLabel.config(text=filename[0])

    def exeBtnEvent(self):
        sfile_path = self.ui.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '删除失败 请正确选择文件')
            return
        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '删除失败 pdf文件不正确')
            return
        fstr = self.entry.get().split(',')
        for index in fstr:
            if not isDigit(index):
                fstr2 = index.split('-')
                if len(fstr2) != 2:
                    messagebox.showerror('错误', '删除失败 格式错误')
                    return
                for index2 in fstr2:
                    if not isDigit(index2):
                        messagebox.showerror('错误', '删除失败 格式错误')
                        return
                if int(fstr2[0]) >= int(fstr2[1]) or int(fstr2[0]) < 1 or int(fstr2[1]) > pdf_num:
                    messagebox.showerror(
                        '错误', '删除失败 输入不合法 有可能左边比右边大 有可能输入了0 还有可能超过了原文件的页数')
                    return
            else:
                if int(index) > pdf_num or int(index) == 0:
                    messagebox.showerror(
                        '错误', '删除失败 输入不合法 有可能输入了0 有可能超过了原文件的页数')
                    return
        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        
        del_range = []
        for index in fstr:
            if isDigit(index):
                del_range.append(int(index))
            else :
                fstr2 = index.split('-')
                for i in fstr2:
                    if del_range.count(int(i)) == 0:
                        del_range.append(int(i)) 
        del_range.sort()
        pdf_output = PdfFileWriter()
        for i in range(pdf_num):
            if del_range.count(i+1) == 0:
                pdf_output.addPage(pdf_input.getPage(i))
            self.ui.progress(100*(i+1)/pdf_num)
        with open(dfile_path+'\\' + os.path.splitext(
                    os.path.basename(sfile_path))[0]+'_delete_' + self.entry.get() + '.pdf', 'wb') as f:
                pdf_output.write(f)
        messagebox.showinfo('提示', '删除已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)