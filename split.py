from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import math

def isDigit(x):
    try:
        x = int(x)
        return isinstance(x, int)
    except ValueError:
        return False

class Split():
    def __init__(self, ui):
        self.ui = ui
        Frame1 = Frame(self.ui.workFrm)
        Frame1.pack(side='top')
        self.ui.Frame2 = Frame(self.ui.workFrm)
        self.ui.Frame2.pack(side='top')
        # 文件路径label
        self.ui.sel_fileLabel = Label(Frame1, width=60)
        self.ui.sel_fileLabel.pack(side='top')

        self.ui.fileBtnSel = Button(
            Frame1, text='选择文件', command=self.fileSelBtnEvent)
        self.ui.fileBtnSel.pack(side='top')
        # 单选框
        self.ui.v = IntVar()
        self.ui.v.set(0)
        r1 = Radiobutton(Frame1, text="指定每个文档页数", variable=self.ui.v,
                         value=1, command=self.mode1)
        r1.pack(side='left', expand='yes', pady=20)
        Radiobutton(Frame1, text="指定生成文档个数", variable=self.ui.v,
                    value=2, command=self.mode2).pack(side='left', expand='yes', pady=20)
        Radiobutton(Frame1, text="自定义生成文档", variable=self.ui.v,
                    value=3, command=self.mode3).pack(side='left', expand='yes', pady=20)
        r1.select()
        r1.invoke()

    # 模式1
    def mode1(self):
        for widget in self.ui.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.ui.Frame2, text='每').pack(side='top')
        self.ui.split_mode1_entry = Entry(self.ui.Frame2, justify='center')
        self.ui.split_mode1_entry.pack(side='top')
        Label(self.ui.Frame2, text='页成为一个文档').pack(side='top')
        Button(self.ui.Frame2, text='执行分割',
               command=self.mode1ExeBtnEvent).pack(side='top')
    
    # 模式2
    def mode2(self):
        for widget in self.ui.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.ui.Frame2, text='分割成').pack(side='top')
        self.split_mode2_entry = Entry(self.ui.Frame2, justify='center')
        self.split_mode2_entry.pack(side='top')
        Label(self.ui.Frame2, text='个文档').pack(side='top')
        Button(self.ui.Frame2, text='执行分割',
               command=self.mode2ExeBtnEvent).pack(side='top')
    
    # 模式3
    def mode3(self):
        for widget in self.ui.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.ui.Frame2, text='请按照格式填写 页数或者页数-页数(逗号隔开)').pack(side='top')
        self.split_mode3_entry = Entry(self.ui.Frame2, justify='center')
        self.split_mode3_entry.pack(side='top')
        Label(self.ui.Frame2, text='举例 "1,2-4,5,7-10"').pack(side='top')
        Button(self.ui.Frame2, text='执行分割',
               command=self.mode3ExeBtnEvent).pack(side='top')
    
    # 选择文件按钮
    def fileSelBtnEvent(self):
        filename = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')], multiple=False)
        if len(filename) == 0:
            return
        self.ui.sel_fileLabel.config(text=filename[0])

    # 模式1执行按钮
    def mode1ExeBtnEvent(self):
        sfile_path = self.ui.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '分割失败 请正确选择要分割的文件')
            return

        if isDigit(self.ui.split_mode1_entry.get()) != True:
            messagebox.showerror('错误', '分割失败 请输入数字')
            return
        step = int(self.ui.split_mode1_entry.get())
        if step <= 0:
            messagebox.showerror('错误', '分割失败 输入大于0的数字')
            return

        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '分割失败 pdf文件不正确')
            return
        if pdf_num <= step:
            messagebox.showerror('错误', '分割失败 原文件页数不足以分割')
            return
        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        i = 0
        while i < pdf_num:
            pdf_output = PdfFileWriter()
            j = 0
            while j < step:
                pdf_output.addPage(pdf_input.getPage(i))
                i += 1
                j += 1
                if(i == pdf_num):
                    break
            out_stream = open(dfile_path+'\\' + os.path.splitext(
                os.path.basename(sfile_path))[0]+'_splite_' + str(math.ceil(i/step)) + '.pdf', 'wb')
            pdf_output.write(out_stream)
            self.ui.progress(100*i/pdf_num)

        messagebox.showinfo('提示', '拆分已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)

    # 模式2执行按钮
    def mode2ExeBtnEvent(self):
        sfile_path = self.ui.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '分割失败 请正确选择要分割的文件')
            return
        if isDigit(self.split_mode2_entry.get()) != True:
            messagebox.showerror('错误', '分割失败 请输入数字')
            return
        file_num = int(self.split_mode2_entry.get())
        if file_num <= 1:
            messagebox.showerror('错误', '分割失败 输入大于1的数字')
            return
        
        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '分割失败 pdf文件不正确')
            return
        if pdf_num < file_num:
            messagebox.showerror('错误', '分割失败 原文件页数不足以分割')
            return
        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        Quot_left = pdf_num // file_num
        Quot_right = math.ceil(pdf_num / file_num)
        mod_left = pdf_num % Quot_left
        mod_right = pdf_num % Quot_right
        quot = 0
        if mod_left <= Quot_right - mod_right:
            quot = Quot_left
        else:
            quot = Quot_right
        i = 0
        j = 0
        while i < file_num:
            pdf_output = PdfFileWriter()
            k = 0
            if i == file_num-1:
                left_page_num = pdf_num-j
                while k < left_page_num:
                    pdf_output.addPage(pdf_input.getPage(j))
                    j += 1
                    k += 1
            else:
                while k < quot:
                    pdf_output.addPage(pdf_input.getPage(j))
                    j += 1
                    k += 1
            i += 1
            out_stream = open(dfile_path+'\\' + os.path.splitext(
                os.path.basename(sfile_path))[0]+'_splite_' + str(i) + '.pdf', 'wb')
            pdf_output.write(out_stream)
            self.ui.progress(100*i/file_num)
        messagebox.showinfo('提示', '拆分已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)

    # 模式3执行按钮
    def mode3ExeBtnEvent(self):
        sfile_path = self.ui.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '分割失败 请正确选择要分割的文件')
            return
        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '分割失败 pdf文件不正确')
            return
        fstr = self.split_mode3_entry.get().split(',')
        for index in fstr:
            if not isDigit(index):
                fstr2 = index.split('-')
                if len(fstr2) != 2:
                    messagebox.showerror('错误', '分割失败 格式错误')
                    return
                for index2 in fstr2:
                    if not isDigit(index2):
                        messagebox.showerror('错误', '分割失败 格式错误')
                        return
                if int(fstr2[0]) >= int(fstr2[1]) or int(fstr2[0]) < 1 or int(fstr2[1]) > pdf_num:
                    messagebox.showerror(
                        '错误', '分割失败 输入不合法 有可能左边比右边大 有可能输入了0 还有可能超过了原文件的页数')
                    return
            else:
                if int(index) > pdf_num or int(index) == 0:
                    messagebox.showerror(
                        '错误', '分割失败 输入不合法 有可能输入了0 有可能超过了原文件的页数')
                    return

        self.ui.disable(self.ui.selFrm)
        self.ui.disable(self.ui.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.ui.progress(0)
            self.ui.enable(self.ui.selFrm)
            self.ui.enable(self.ui.workFrm)
            return
        j = 0
        for index in fstr:
            if isDigit(index):
                pdf_output = PdfFileWriter()
                pdf_output.addPage(pdf_input.getPage(int(index)-1))
                out_stream = open(dfile_path+'\\' + os.path.splitext(
                    os.path.basename(sfile_path))[0]+'_splite_' + index + '.pdf', 'wb')
                pdf_output.write(out_stream)
            else:
                fstr2 = index.split('-')
                i = 0
                pdf_output = PdfFileWriter()
                while i < int(fstr2[1])-int(fstr2[0])+1:
                    pdf_output.addPage(pdf_input.getPage(int(fstr2[0])+i-1))
                    i += 1
                out_stream = open(dfile_path+'\\' + os.path.splitext(
                    os.path.basename(sfile_path))[0]+'_splite_' + index + '.pdf', 'wb')
                pdf_output.write(out_stream)
            j += 1
            self.ui.progress(100*j/len(fstr))
        messagebox.showinfo('提示', '拆分已完成')
        self.ui.progress(0)
        self.ui.enable(self.ui.selFrm)
        self.ui.enable(self.ui.workFrm)
