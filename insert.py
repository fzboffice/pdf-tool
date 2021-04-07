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


class Insert():
    def __init__(self,ui):
        self.ui=ui
        Frame1 = Frame(self.ui.workFrm)
        Frame1.pack(side='top')
        self.ui.sel_fileLabel = Label(Frame1, width=60)
        self.ui.sel_fileLabel.pack(side='top')
        self.ui.fileBtnSel = Button(
            Frame1, text='选择被插入的pdf', command=self.fileSelBtnEvent)
        self.ui.fileBtnSel.pack(side='top')

        self.ui.sel_fileLabel2 = Label(Frame1, width=60)
        self.ui.sel_fileLabel2.pack(side='top')
        self.ui.fileBtnSel2 = Button(
            Frame1, text='选择要插入的pdf', command=self.fileSelBtn2Event)
        self.ui.fileBtnSel2.pack(side='top')
        
        self.ui.Label3 = Label(Frame1, text='请输入插入到第几页')
        self.ui.Label3.pack(side='top')

        self.entry = Entry(Frame1, justify='center')
        self.entry.pack()

        self.ui.exeBtn = Button(
            Frame1, text='执行插入', command=self.exeBtnEvent)
        self.ui.exeBtn.pack(side='top',pady = 5)

    def fileSelBtnEvent(self):
        filename = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')], multiple=False)
        if len(filename) == 0:
            return
        self.ui.sel_fileLabel.config(text=filename[0])

    def fileSelBtn2Event(self):
        filename = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')], multiple=False)
        if len(filename) == 0:
            return
        self.ui.sel_fileLabel2.config(text=filename[0])

    def exeBtnEvent(self):
        sfile_path = self.ui.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '插入失败 请正确选择要分割的文件')
            return
        page_num =  pdf_input.getNumPages()

        sfile_path2 = self.ui.sel_fileLabel2.cget('text')
        try:
            pdf_input2 = PdfFileReader(sfile_path2)
        except:
            messagebox.showerror('错误', '插入失败 请正确选择要分割的文件')
            return
        page_num2 =  pdf_input2.getNumPages()

        if not isDigit(self.entry.get()):
            messagebox.showerror('错误', '插入失败 请正确输入插入页数')
            return
        page_to_insert = int(self.entry.get())
        if page_to_insert < 1 or page_to_insert>page_num+1 :
            messagebox.showerror('错误', '插入失败 请正确输入插入页数')
            return

        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        pdf_output = PdfFileWriter()
        with open(dfile_path+'\\' + os.path.splitext(
                    os.path.basename(sfile_path))[0]+'_insert' + '.pdf', 'wb') as f:
            for i in range(page_num):
                pdf_output.addPage(pdf_input.getPage(i))
            for i in reversed(range(page_num2)):
                pdf_output.insertPage(pdf_input2.getPage(i),page_to_insert-1)
                self.ui.progress(100*(i+1)/page_num2)
            pdf_output.write(f)
        messagebox.showinfo('提示', '插入已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)
