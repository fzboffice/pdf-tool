from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfFileWriter, PdfFileReader
from icon import img
import base64
import os
import math


def isDigit(x):
    try:
        x = int(x)
        return isinstance(x, int)
    except ValueError:
        return False


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
        self.selbtn_merge = Button(
            self.selFrm, text='合并pdf', width=15, height=5, command=self.selbtn_merge_event)
        self.selbtn_merge.pack(side='left', expand='yes', pady=30)
        # 拆分按钮
        self.selbtn_split = Button(
            self.selFrm, text='拆分pdf', width=15, height=5, command=self.selbtn_spilt_event)
        self.selbtn_split.pack(side='left', expand='yes', pady=30)
        # 插入按钮
        self.selbtn_add = Button(
            self.selFrm, text='插入pdf', width=15, height=5, command=self.selbtn_add_event)
        self.selbtn_add.pack(side='left', expand='yes', pady=30)
        # 删除按钮
        self.selbtn_del = Button(
            self.selFrm, text='删除pdf', width=15, height=5, command=self.selbtn_del_event)
        self.selbtn_del.pack(side='left', expand='yes', pady=30)

        # 2.功能框架
        self.workFrm = Frame(self.root)
        self.workFrm.pack()

        # 3.进度框架
        self.prgfrm = Frame(self.root)
        self.prgfrm.pack(side='bottom')

        self.canvas = Canvas(self.prgfrm, width=800, height=22, bg="white")
        self.canvas.pack()
        self.prgsRectagl = self.canvas.create_rectangle(
            0, 0, 0, 0, width=0, fill="green")

        self.selbtn_merge_event()
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

    def selbtn_merge_event(self):
        # print('merge press')
        self.selbtn_merge.config(relief='sunken')
        self.selbtn_split.config(relief='raised')
        self.selbtn_add.config(relief='raised')
        self.selbtn_del.config(relief='raised')
        for widget in self.workFrm.winfo_children():
            widget.destroy()
        # list
        self.fileList = Listbox(
            self.workFrm, width=82, height=20, selectmode='extended')
        self.fileList.pack(side='left')
        self.fileList.bind('<Button-1>', self.show)
        self.fileList.bind('<B1-Motion>', self.showInfo)

        # scrollbar
        sc = Scrollbar(self.workFrm, command=self.fileList.yview)
        sc.pack(side=LEFT, fill=Y)
        self.fileList.config(yscrollcommand=sc.set)

        # 按钮
        self.fileBtnAdd = Button(
            self.workFrm, text='添加文件', width=12, height=2, command=self.merge_fileBtnAdd_event)
        self.fileBtnAdd.pack(side='top', expand='yes')
        self.fileBtnMov = Button(
            self.workFrm, text='移除文件', width=12, height=2, command=self.merge_fileBtnMov_event)
        self.fileBtnMov.pack(side='top', expand='yes')
        self.fileBtnExe = Button(
            self.workFrm, text='执行合并', width=12, height=2, command=self.merge_fileBtnExe_event)
        self.fileBtnExe.pack(side='top', expand='yes')

    def merge_fileBtnAdd_event(self):
        filenames = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')])
        if len(filenames) == 0:
            return
        for item in filenames:
            self.fileList.insert(END, item)

    def merge_fileBtnMov_event(self):
        indices = self.fileList.curselection()
        if len(indices) == 0:
            return
        for index in reversed(indices):
            self.fileList.delete(index)

    def merge_fileBtnExe_event(self):
        if self.fileList.size() <= 1:
            return
        # self.disable(self.workFrm)
        self.disable(self.selFrm)
        self.disable(self.workFrm)
        try:
            file_merger = PdfFileMerger()
            i = 0
            for pdf in self.fileList.get(0, END):
                file_merger.append(pdf)
                i += 1
                self.progress(100*i/self.fileList.size())
        except:
            messagebox.showerror('错误', '合并失败 pdf文件可能无效')
            self.progress(0)
            self.enable(self.selFrm)
            self.enable(self.workFrm)
            return
        filename = filedialog.asksaveasfilename(
            initialfile=os.path.splitext(os.path.basename(self.fileList.get(0)))[0]+'_merge', defaultextension=".pdf", filetypes=[('PDF file', '*.pdf')])
        if len(filename) == 0:
            self.progress(0)
            self.enable(self.selFrm)
            self.enable(self.workFrm)
            return
        file_merger.write(filename)
        file_merger.close()
        messagebox.showinfo('提示', '合并已完成')
        self.progress(0)
        self.enable(self.selFrm)
        self.enable(self.workFrm)

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
    # 分割功能选择按钮事件

    def selbtn_spilt_event(self):
        self.selbtn_merge.config(relief='raised')
        self.selbtn_split.config(relief='sunken')
        self.selbtn_add.config(relief='raised')
        self.selbtn_del.config(relief='raised')
        for widget in self.workFrm.winfo_children():
            widget.destroy()

        Frame1 = Frame(self.workFrm)
        Frame1.pack(side='top')
        self.Frame2 = Frame(self.workFrm)
        self.Frame2.pack(side='top')
        # 文件路径label
        self.sel_fileLabel = Label(Frame1, width=60)
        self.sel_fileLabel.pack(side='top')

        self.fileBtnSel = Button(
            Frame1, text='选择文件', command=self.spilt_fileBtnSel_event)
        self.fileBtnSel.pack(side='top')
        # 单选框
        self.v = IntVar()
        self.v.set(0)
        r1 = Radiobutton(Frame1, text="指定每个文档页数", variable=self.v,
                         value=1, command=self.spilt_mode_sel_mode1)
        r1.pack(side='left', expand='yes', pady=20)
        Radiobutton(Frame1, text="指定生成文档个数", variable=self.v,
                    value=2, command=self.spilt_mode_sel_mode2).pack(side='left', expand='yes', pady=20)
        Radiobutton(Frame1, text="自定义生成文档", variable=self.v,
                    value=3, command=self.spilt_mode_sel_mode3).pack(side='left', expand='yes', pady=20)
        r1.select()
        r1.invoke()

    def spilt_mode_sel_mode1(self):
        for widget in self.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.Frame2, text='每').pack(side='top')
        self.split_mode1_entry = Entry(self.Frame2, justify='center')
        self.split_mode1_entry.pack(side='top')
        Label(self.Frame2, text='页成为一个文档').pack(side='top')
        Button(self.Frame2, text='执行分割',
               command=self.spilt_mode1_exeBtn_event).pack(side='top')
    # 模式2

    def spilt_mode_sel_mode2(self):
        for widget in self.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.Frame2, text='分割成').pack(side='top')
        self.split_mode2_entry = Entry(self.Frame2, justify='center')
        self.split_mode2_entry.pack(side='top')
        Label(self.Frame2, text='个文档').pack(side='top')
        Button(self.Frame2, text='执行分割',
               command=self.spilt_mode2_exeBtn_event).pack(side='top')

    def spilt_mode_sel_mode3(self):
        for widget in self.Frame2.winfo_children():
            widget.destroy()
        # label
        Label(self.Frame2, text='请按照格式填写 页数或者页数-页数(逗号隔开)').pack(side='top')
        self.split_mode3_entry = Entry(self.Frame2, justify='center')
        self.split_mode3_entry.pack(side='top')
        Label(self.Frame2, text='举例 "1,2-4,5,7-10"').pack(side='top')
        Button(self.Frame2, text='执行分割',
               command=self.spilt_mode3_exeBtn_event).pack(side='top')
    # 模式1选择文件按钮

    def spilt_fileBtnSel_event(self):
        filename = filedialog.askopenfilenames(
            filetypes=[('PDF file', '*.pdf')], multiple=False)
        if len(filename) == 0:
            return
        self.sel_fileLabel.config(text=filename[0])

    # 执行模式1执行按钮
    def spilt_mode1_exeBtn_event(self):
        if isDigit(self.split_mode1_entry.get()) != True:
            messagebox.showerror('错误', '分割失败 请输入数字')
            return
        step = int(self.split_mode1_entry.get())
        if step <= 0:
            messagebox.showerror('错误', '分割失败 输入大于0的数字')
            return

        sfile_path = self.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '分割失败 请正确选择要分割的文件')
            return

        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '分割失败 pdf文件不正确')
            return
        if pdf_num <= step:
            messagebox.showerror('错误', '分割失败 原文件页数不足以分割')
            return
        self.disable(self.selFrm)
        self.disable(self.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.progress(0)
            self.enable(self.selFrm)
            self.enable(self.workFrm)
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
            self.progress(100*i/pdf_num)

        messagebox.showinfo('提示', '拆分已完成')
        self.progress(0)
        self.enable(self.selFrm)
        self.enable(self.workFrm)

    # 执行模式2执行按钮
    def spilt_mode2_exeBtn_event(self):
        if isDigit(self.split_mode2_entry.get()) != True:
            messagebox.showerror('错误', '分割失败 请输入数字')
            return
        file_num = int(self.split_mode2_entry.get())
        if file_num <= 1:
            messagebox.showerror('错误', '分割失败 输入大于1的数字')
            return
        sfile_path = self.sel_fileLabel.cget('text')
        try:
            pdf_input = PdfFileReader(sfile_path)
        except:
            messagebox.showerror('错误', '分割失败 请正确选择要分割的文件')
            return
        pdf_num = pdf_input.getNumPages()
        if pdf_num <= 0:
            messagebox.showerror('错误', '分割失败 pdf文件不正确')
            return
        if pdf_num < file_num:
            messagebox.showerror('错误', '分割失败 原文件页数不足以分割')
            return
        self.disable(self.selFrm)
        self.disable(self.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.progress(0)
            self.enable(self.selFrm)
            self.enable(self.workFrm)
            return

        Quot_left = pdf_num // file_num
        Quot_right = math.ceil(pdf_num / file_num)
        mod_left = pdf_num % Quot_left
        mod_right = pdf_num % Quot_right

        # 11 3  335  443
        # 10 3  334  442
        # 17 3  557  665
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
            self.progress(100*i/file_num)

        messagebox.showinfo('提示', '拆分已完成')
        self.progress(0)
        self.enable(self.selFrm)
        self.enable(self.workFrm)

    def spilt_mode3_exeBtn_event(self):
        sfile_path = self.sel_fileLabel.cget('text')
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
                    messagebox.showerror('错误', '分割失败 输入不合法 有可能左边比右边大 有可能输入了0 还有可能超过了原文件的页数')
                    return
            else:
                if int(index) > pdf_num or int(index) == 0:
                    messagebox.showerror('错误', '分割失败 输入不合法 有可能输入了0 有可能超过了原文件的页数')
                    return

        self.disable(self.selFrm)
        self.disable(self.workFrm)
        dfile_path = filedialog.askdirectory()
        if len(dfile_path) == 0:
            self.progress(0)
            self.enable(self.selFrm)
            self.enable(self.workFrm)
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
                    i+=1
                out_stream = open(dfile_path+'\\' + os.path.splitext(
                os.path.basename(sfile_path))[0]+'_splite_' + index + '.pdf', 'wb')
                pdf_output.write(out_stream)
            j+=1
            self.progress(100*j/len(fstr))
        messagebox.showinfo('提示', '拆分已完成')
        self.progress(0)
        self.enable(self.selFrm)
        self.enable(self.workFrm)
    # 插入功能选择按钮事件

    def selbtn_add_event(self):
        self.selbtn_merge.config(relief='raised')
        self.selbtn_split.config(relief='raised')
        self.selbtn_add.config(relief='sunken')
        self.selbtn_del.config(relief='raised')
        for widget in self.workFrm.winfo_children():
            widget.destroy()

    # 删除功能选择按钮事件
    def selbtn_del_event(self):
        self.selbtn_merge.config(relief='raised')
        self.selbtn_split.config(relief='raised')
        self.selbtn_add.config(relief='raised')
        self.selbtn_del.config(relief='sunken')
        for widget in self.workFrm.winfo_children():
            widget.destroy()


main = Application()
main.root.mainloop()
