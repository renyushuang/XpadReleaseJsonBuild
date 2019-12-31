import os

from tkinter import *
from tkinter import filedialog

filename = None
path = None
selectValue: IntVar = None
listBox: Listbox = None


def insertListBoxMessage(item):
    global listBox
    listBox.insert("end", item)
    listBox.see(END)


def selectPath():
    global filename
    global path
    filename = filedialog.askopenfilename(filetypes=[("excel格式", "xlsx")])
    path.set(filename)


def startCreateJson():
    global selectValue
    value = selectValue.get()
    if value == 1:
        os.system("python3 XpadJsonBuild_1.0.py %s" % filename)
    if value == 2:
        os.system("python3 XpadJsonBuild_2.0.py %s" % filename)
    if value == 3:
        os.system("python3 XpadJsonBuild_data_pip.py %s" % filename)
    pass


def creatMainUi():
    global path
    global selectValue
    global listBox

    root = Tk()
    root.title("XpadJsonBuild")
    root.geometry("1000x618")
    root.resizable(False, False)

    path = StringVar()
    selectValue = IntVar()
    selectValue.set(1)

    topFrame = Frame(root)
    topFrame.pack(side=TOP)

    Label(topFrame, text="目标路径:").pack(side=LEFT, padx=5, pady=10)
    Entry(topFrame, textvariable=path).pack(side=LEFT, padx=5, pady=10)
    Button(topFrame, text="路径选择", command=selectPath).pack(side=LEFT, padx=5, pady=10)
    Button(topFrame, text="开始生成", command=startCreateJson).pack(side=LEFT, padx=5, pady=10)

    middleFrame = Frame(root)
    middleFrame.pack(side=TOP)

    rb1 = Radiobutton(middleFrame, text='1.0.py', variable=selectValue, value=1)
    rb1.pack(side=LEFT)

    rb2 = Radiobutton(middleFrame, text='2.0.py', variable=selectValue, value=2)
    rb2.pack(side=LEFT)

    rb3 = Radiobutton(middleFrame, text='data_pip.py', variable=selectValue, value=3)
    rb3.pack(side=LEFT)

    bottomFrame = Frame(root)
    scrollbar = Scrollbar(bottomFrame)
    scrollbar.pack(side=RIGHT, fill=Y)
    listBox = Listbox(root, yscrollcommand=scrollbar.set)
    scrollbar.config(command=listBox.yview)
    listBox.pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
    bottomFrame.pack(side=TOP, fill=BOTH, expand=YES)

    root.mainloop()

if __name__ == '__main__':
    creatMainUi()
