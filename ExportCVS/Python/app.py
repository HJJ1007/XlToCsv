from tkinter import*
from tkinter import filedialog, messagebox
from tkinterdnd2 import *
import tkinter.ttk as ttk
import os
import sys
try:
    os.chdir(sys._MEIPASS)
    print(sys._MEIPASS)
except:
    os.chdir(os.getcwd())
# 툴팁

import FileFormatUtil


class CreateToolTip(object):
    """
    create a tooltip for a given widget
    """

    def __init__(self, widget, text='widget info'):
        self.waittime = 500  # miliseconds
        self.wraplength = 180  # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        # self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(self.tw, text=self.text, justify='left',
                      background="#ffffff", relief='solid', borderwidth=1,
                      wraplength=self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


def SelectFileBtn():
    dir_file = filedialog.askopenfilename(parent=root, initialdir=os.getcwd(
    ), title='Please select a directory', filetypes=[("Excel files", "*.xlsx; *.xlsm")])
    if dir_file == '':
        return

    t_Log.config(state='normal')
    t_Log.delete('1.0', END)
    t_Log.config(state='disabled')
    SetProgress(0)
    label.config(text='', foreground='black')

    os.chdir(os.path.dirname(dir_file))
    SetSheetList(FF.SetFileName(dir_file))
    SetEntryTxt(inputFileEntry, dir_file)
    SetEntryTxt(outputDirEntry, FF.SetSavePath(dir_file, True))


def SelectOutputDirBtn():
    dir = filedialog.askdirectory(
        parent=root, initialdir=outputDirEntry.get(), title='Please select a directory')
    if dir == '':
        return
    SetEntryTxt(outputDirEntry, FF.SetSavePath(dir))


def convertingBtn(kind):

    if outputDirEntry.get() == '':
        messagebox.showerror('error', '저장 경로를 지정하지 않았습니다.')
        return
    if len(FF.convertList) == 0:
        msgBox = messagebox.askquestion(
            'error', '변환할 파일을 선택하지 않았습니다.\n시트 전체를 변환하시겠습니까?')
        if msgBox == 'yes':
            FF.convertList = FF.sNames
        else:
            return
    openBtn.config(state=DISABLED)
    conversionToCsvBtn.config(state=DISABLED)
    # conversionToJsonBtn.config(state=DISABLED)
    ret = FF.Convert(kind)
    openBtn.config(state=NORMAL)
    conversionToCsvBtn.config(state=NORMAL)
    # conversionToJsonBtn.config(state=NORMAL)
    if not ret:
        messagebox.showerror(
            "error", "csv 변환 중에 에러가 발생했습니다.\ncsv파일이 열려 있다면 닫아주세요.")


def OpenFolder():
    pathName = FF.GetFileName()

    if not pathName:
        messagebox.showerror("error", "변환할 파일을 선택하세요.")
        InsertLog('Error : not selected file', 'red')
        return

    if not FF.completeConvert:
        messagebox.showerror("error", "아직 변환하지 않았습니다. 변환 후 다시 시도하세요.")
        InsertLog('Error : not converted yet', 'red')
        return

    InsertLog(
        'Opened the folder containing the successfully converted files.', 'blue')
    path = os.path.realpath(FF.savePath)
    os.startfile(path)


def SetProgress(currentValue):
    p_var.set(currentValue)
    progressbar.update()
    label.config(text='Converting : {:.2f} %'.format(
        currentValue), foreground='black')
    if currentValue >= 100:
        label.config(text='Complete!!', foreground='blue')


def InsertLog(txt, color='black'):
    t_Log.config(state='normal')
    t_Log.insert('end', txt + '\n')
    t_Log.tag_add(color + '_fg', 'end-2c linestart', 'end-1c')
    t_Log.config(state='disabled')
    t_Log.see('end')


def SetSheetList(sNames):
    sheetListbox.delete(0, END)
    if not sNames:
        return
    for name in sNames:
        sheetListbox.insert(END, name)


def OnSelect(event):
    widget = event.widget
    try:
        value = []
        for i in widget.curselection():
            value.append(widget.get(i))
    except:
        return
    FF.convertList = value


def SetEntryTxt(entry, txt):
    entry.config(state='normal')
    entry.delete(0, 'end')
    entry.insert(0, txt)
    entry.config(state='readonly')


def SetIsFirstRowConvert():
    # FF.delFirstrow = delFirstrow.get()
    # print(FF.delFirstrow)
    FF.delFirstrow = False


def SetSelectFileName(event):
    exel_tuple = ('.xlsx', 'xlsm')
    if event.data.endswith(exel_tuple):
        dir_file = event.data
        t_Log.config(state='normal')
        t_Log.delete('1.0', END)
        t_Log.config(state='disabled')
        SetProgress(0)
        label.config(text='', foreground='black')
        os.chdir(os.path.dirname(dir_file))
        SetSheetList(FF.SetFileName(dir_file))
        SetEntryTxt(inputFileEntry, dir_file)
        SetEntryTxt(outputDirEntry, FF.SetSavePath(dir_file, True))


root = Tk()
root.title("xlsxToCsv App")
root.geometry("750x520+660+300")
root.resizable(False, False)
# 변수 선언
dir_file = ''

# 변환할 파일 경로 1
inputdir_frame = Frame(root, relief='raised')

inputFileTxt = Label(inputdir_frame)
inputFileTxt.config(text='파일 경로')
inputFileTxt.pack(side='left', padx=20)

inputFileEntry = Entry(inputdir_frame)
inputFileEntry.pack(side='left', fill='x', expand=YES)
inputFileEntry.config(state='readonly')
inputFileEntry.drop_target_register(DND_FILES)
inputFileEntry.dnd_bind('<<Drop>>', SetSelectFileName)

selectBtn = Button(inputdir_frame, text='파일 선택', overrelief="solid",
                   command=SelectFileBtn, repeatdelay=1000, repeatinterval=100)
selectBtn.pack(side='left', padx=20)


FF = FileFormatUtil.FileFormat(dir_file)

# 시트 리스트4
middleFrame = Frame(root)

rowFrame = Frame(middleFrame)
sheetFrame = Frame(rowFrame)

questionImg = PhotoImage(file="./image/question_mark.png")

sheetlabel = Button(rowFrame, text='SHEET 목록  ', image=questionImg,
                    compound=RIGHT, highlightthickness=0, bd=0, relief="sunken")
questionImg_ttp = CreateToolTip(
    sheetlabel, 'sheet 목록을 클릭하지 않은 기본 값은 sheet 전체입니다. \nsheet 목록에서 csv로 변환할 sheet를 클릭합니다.')

s_scrollbar = Scrollbar(sheetFrame)
s_scrollbar.pack(side="right", fill="y")

sheetListbox = Listbox(
    sheetFrame, yscrollcommand=s_scrollbar.set, selectmode='multiple', activestyle='none')
sheetListbox.pack(side="left", fill='both', expand=YES)
sheetListbox.bind("<<ListboxSelect>>", OnSelect)
s_scrollbar["command"] = sheetListbox.yview

sheetlabel.pack(side='top', fill='x')
sheetFrame.pack(side='top', fill='both', expand=YES, padx=5)
rowFrame.pack(side='left', fill='both', expand=YES, padx=5)
# 로그
Logframe = Frame(middleFrame)

l_scrollbar = Scrollbar(Logframe)
l_scrollbar.pack(side="right", fill="y")

l_scrollbar2 = Scrollbar(Logframe, orient='horizontal')
l_scrollbar2.pack(side="bottom", fill="x")

t_Log = Text(Logframe, yscrollcommand=l_scrollbar.set,
             xscrollcommand=l_scrollbar2.set, wrap=NONE)
t_Log.tag_configure("red_fg", foreground="red")
t_Log.tag_configure("blue_fg", foreground="blue")
t_Log.tag_configure("yellow_fg", foreground="yellow")
t_Log.tag_configure("black_fg", foreground="black")
t_Log.tag_configure("orange_fg", foreground="orange")

t_Log.pack(side="left", fill='both', expand=YES)

l_scrollbar["command"] = t_Log.yview
l_scrollbar2["command"] = t_Log.xview

Logframe.pack(side='left', padx=5)

label = Label(root)

# 게이지 바 5
p_var = DoubleVar()
progressbar = ttk.Progressbar(
    root, orient="horizontal", mode="determinate", variable=p_var, maximum=100)

# 출력 파일 경로 2
outputdir_frame = Frame(root, relief='raised')

outputFileTxt = Label(outputdir_frame)
outputFileTxt.config(text='저장 경로')
outputFileTxt.pack(side='left', padx=20)

outputDirEntry = Entry(outputdir_frame)
outputDirEntry.pack(side='left', fill='x', expand=YES)
outputDirEntry.config(state='readonly')

selectBtn2 = Button(outputdir_frame, text='경로 선택', overrelief="solid",
                    command=SelectOutputDirBtn, repeatdelay=1000, repeatinterval=100)
selectBtn2.pack(side='left', padx=20)


# 변환, 폴더 열기 버튼 3
btnFrame = Frame(root, relief="solid")

delFirstrow = BooleanVar()
firstRowConvert_checkBtn = Checkbutton(
    btnFrame, text='데이터 타입 제거', variable=delFirstrow, command=SetIsFirstRowConvert, state='disabled')


btnFrame2 = Frame(btnFrame)
conversionToCsvBtn = Button(btnFrame2, text='csv 변환', overrelief="solid",
                            command=lambda: convertingBtn('csv'), repeatdelay=1000, repeatinterval=100, width=10, pady=5)
# conversionToJsonBtn = Button(btnFrame2, text='json 변환', overrelief="solid",
#                        command=lambda : convertingBtn('json'), repeatdelay=1000, repeatinterval=100, width=10, pady=5)
openBtn = Button(btnFrame2, text='폴더 열기', overrelief="solid",
                 command=OpenFolder, repeatdelay=1000, repeatinterval=100, width=10, pady=5)
ClearLogBtn = Button(btnFrame2, text='로그 지우기', overrelief="solid",
                     command=lambda: (t_Log.config(state='normal'), t_Log.delete('1.0', END), t_Log.config(state='disabled')), repeatdelay=1000, repeatinterval=100, width=10, pady=5)

firstRowConvert_checkBtn.   pack(side='left')
conversionToCsvBtn.         pack(side='left', padx=15)
# conversionToJsonBtn.        pack(side='left')
openBtn.                    pack(side='left', padx=15)
ClearLogBtn.                pack(side='left')
btnFrame2.                  pack(side='left', padx=90)

inputdir_frame.pack(side='top', fill='x', pady=5)
outputdir_frame.pack(side='top', fill='x')
btnFrame.pack(side='top', pady=10, fill='x')
middleFrame.pack(side='top', fill='both')
label.pack(side='top', pady=5)
progressbar.pack(side='top', fill='x', padx=10)

FF.SetLoadingbarEventBind(SetProgress)
FF.PrintLogEventBind(InsertLog)
os.chdir('C:/')

root.mainloop()
