import os

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PIL import ImageTk

import XpadJsonBuild_data_pip
import XpadJsonBuild_1
import XpadJsonBuild_2

# 生成应用
# sudo pyinstaller XpadJsonBuild_GUI.py -p XpadJsonBuild_1.py -p XpadJsonBuild_2.py -p XpadJsonBuild_data_pip.py --hidden-import XpadJsonBuild_1 --hidden-import XpadJsonBuild_2 --hidden-import XpadJsonBuild_data_pip

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
        XpadJsonBuild_1.creatMainUi()
    if value == 2:
        XpadJsonBuild_2.creatMainUi()
    if value == 3:
        XpadJsonBuild_data_pip.creatMainUi()


def creatMainUi():
    global path
    global selectValue
    global listBox

    root = Tk()
    root.title("XpadJsonBuild")
    root.geometry("1000x618")
    root.resizable(False, False)
    root.config()

    path = StringVar()
    selectValue = IntVar()
    selectValue.set(1)

    middleFrame = Frame(root)
    middleFrame.pack(side=TOP)

    rb1 = Radiobutton(middleFrame, text='1.0.py', variable=selectValue, value=1)
    rb1.pack(side=LEFT)

    rb2 = Radiobutton(middleFrame, text='2.0.py', variable=selectValue, value=2)
    rb2.pack(side=LEFT)

    rb3 = Radiobutton(middleFrame, text='data_pip.py', variable=selectValue, value=3)
    rb3.pack(side=LEFT)

    ttk.Button(middleFrame, text="选择", command=startCreateJson).pack(side=LEFT, padx=5, pady=10)

    bottomFrame = Frame(root)
    bottomFrame.pack(side=TOP)

    title: Text = Text(bottomFrame, width=50, height=5)
    title.pack(side=TOP, padx=5, pady=10)
    title.insert(INSERT, "简介")
    content = '\n1.0 : Xpad1.0 广告json 生成工具 \n2.0 : Xpad2.0 广告json 生成工具 \ndata_pip : 产品数据通道 广告配置'
    title.insert(INSERT, content)
    title.config(state=DISABLED, font=("Arial", 20))

    root.mainloop()


if __name__ == '__main__':
    creatMainUi()
