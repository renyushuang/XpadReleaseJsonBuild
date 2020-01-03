# coding=utf-8
import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

import json
import logging
import openpyxl
import os

AD_TYPE_OPEN = "开屏"
AD_TYPE_NATIVE_PLATE = "原生模版"
AD_TYPE_PLATE_RENDER = "模版渲染"
AD_TYPE_INTERSTITIAL = "插屏"
AD_TYPE_FULL_SCREEN_VIDEO = "全屏视频"
AD_TYPE_NATIVE = "原生"
AD_TYPE_NATIVE_FEED = "自渲染"
AD_TYPE_BANNER = "banner"

# extra-type
AD_TYPE_CSJ_NAME = "穿山甲"
AD_TYPE_CSJ = "csj"
# 穿山甲开屏
AD_TYPE_CSJ_OPEN: int = 11
# 穿山甲个性化模版banner
AD_TYPE_CSJ_PERSONAL_PLATE_BANNER: int = 12
# 穿山甲自渲染信息流
AD_TYPE_CSJ_FEED: int = 13
# 穿山甲个性化模版信息流
AD_TYPE_CSJ_PERSONAL_PLATE_FEED: int = 14
# 穿山甲全屏视频
AD_TYPE_CSJ_FULL_SCREEN_VIDEO: int = 15
# 穿山甲个性化模版插屏
AD_TYPE_CSJ_PERSONAL_PLATE_INTERSTITIAL: int = 16

# extra-type
AD_TYPE_YLH_NAME = "广点通"
AD_TYPE_YLH = "ylh"
# 优量汇开屏
AD_TYPE_YLH_OPEN: int = 21
# 优量汇banner2.0
AD_TYPE_YLH_BANNER: int = 22
# 优量汇插屏2.0
AD_TYPE_YLH_INTERSTITIAL: int = 23
# 优量汇原生模版
AD_TYPE_YLH_PERSONAL_PLATE_FEED: int = 24
# 优量汇原生自渲染
AD_TYPE_YLH_FEED: int = 25

# extra-type
AD_TYPE_KSH_NAME = "快手"
AD_TYPE_KSH = "ksh"
# 快手原生自渲染
AD_TYPE_KSH_FEED: int = 31
# 快手模版信息流
AD_TYPE_KSH_PERSONAL_PLATE_FEED: int = 32
# 快手全屏视频
AD_TYPE_KSH_FULL_SCREEN_VIDEO: int = 33

excelPath = ""

mAdResultMap = {}

filename = None
path = None
listBox: Listbox = None
native_ad_refr_it: IntVar = None
native_ad_refr_t_u: IntVar = None


def insertListBoxMessage(item):
    if listBox is not None:
        listBox.insert("end", item)
        listBox.see(END)


def findTitleInColum(name, adSheet):
    maxColumn = int(adSheet.max_column)
    for columnIndex in range(0, maxColumn):
        titleValue = adSheet.cell(row=1, column=columnIndex + 1).value
        if titleValue == name:
            return columnIndex + 1
    insertListBoxMessage("title 不存在 -" + name)
    logging.error("title 不存在 -" + name)
    exit()


def getTitleColumValue(name, adSheet):
    index = findTitleInColum(name, adSheet)
    value = adSheet.cell(row=2, column=index).value
    if value is not None:
        return value

    insertListBoxMessage("当前 " + name + "值为 None")
    logging.warning("当前 " + name + "值为 None")

    return ""


def getCloumeValueColumValue(row, name, adSheet):
    index = findTitleInColum(name, adSheet)
    value = adSheet.cell(row=row, column=index).value

    return value


def hasAdTypeString(adDetailsValuelist: list, typeName):
    length = len(adDetailsValuelist)

    for i in range(0, length):
        if adDetailsValuelist[i] == typeName:
            return True

    return False


def getAdExtraTypeValue(platformValue, adDetailsValue, adSourceIdValue):
    adDetailsValuelist: list = adDetailsValue.split("-")
    if len(adDetailsValuelist) < 2:
        insertListBoxMessage("备注 命名错误" + str(adSourceIdValue))
        logging.error("备注 命名错误" + str(adSourceIdValue))

    if platformValue == AD_TYPE_CSJ:
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_OPEN):
            return AD_TYPE_CSJ_OPEN
        elif hasAdTypeString(adDetailsValuelist, AD_TYPE_INTERSTITIAL):
            return AD_TYPE_CSJ_PERSONAL_PLATE_INTERSTITIAL
        elif hasAdTypeString(adDetailsValuelist, AD_TYPE_FULL_SCREEN_VIDEO):
            return AD_TYPE_CSJ_FULL_SCREEN_VIDEO
        elif hasAdTypeString(adDetailsValuelist, AD_TYPE_PLATE_RENDER):
            return AD_TYPE_CSJ_PERSONAL_PLATE_FEED
        elif hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE_FEED):
            return AD_TYPE_CSJ_FEED
        elif hasAdTypeString(adDetailsValuelist, AD_TYPE_BANNER):
            return AD_TYPE_CSJ_PERSONAL_PLATE_BANNER
        else:
            insertListBoxMessage("不支持这种类型 " + adDetailsValue + "穿山甲的广告id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 " + adDetailsValue + "穿山甲的广告id为 = " + str(adSourceIdValue))

    elif platformValue == AD_TYPE_YLH:
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_OPEN):
            return AD_TYPE_YLH_OPEN
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE_PLATE):
            return AD_TYPE_YLH_PERSONAL_PLATE_FEED
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_INTERSTITIAL):
            return AD_TYPE_YLH_INTERSTITIAL
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
            return AD_TYPE_YLH_FEED
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_BANNER):
            return AD_TYPE_YLH_BANNER
        else:
            insertListBoxMessage("不支持这种类型 " + adDetailsValue + "广点通id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 " + adDetailsValue + "广点通id为 = " + str(adSourceIdValue))

        pass
    elif platformValue == AD_TYPE_KSH:
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
            return AD_TYPE_KSH_FEED
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_PLATE_RENDER):
            return AD_TYPE_KSH_PERSONAL_PLATE_FEED
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_FULL_SCREEN_VIDEO) or hasAdTypeString(adDetailsValuelist,
                                                                                             AD_TYPE_INTERSTITIAL):
            return AD_TYPE_KSH_FULL_SCREEN_VIDEO
        if hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
            return AD_TYPE_KSH_PERSONAL_PLATE_FEED
        else:
            insertListBoxMessage("不支持这种类型 " + adDetailsValue + "快手id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 " + adDetailsValue + "快手id为 = " + str(adSourceIdValue))

        pass
    else:
        insertListBoxMessage("没有这个类型 广告id是 = " + str(adSourceIdValue))
        logging.error("没有这个类型 广告id是 = " + str(adSourceIdValue))
        exit()
    insertListBoxMessage("不支持这种类型 " + adDetailsValue + "id为 = " + str(adSourceIdValue))
    logging.error("不支持这种类型 " + adDetailsValue + "id为 = " + str(adSourceIdValue))
    return None


def addChannelIds(slot: list, adSheet):
    maxColum = adSheet.max_row
    global native_ad_refr_it
    global native_ad_refr_t_u
    global native_ad_new_usr_st

    for sidRowIndex in range(1, maxColum):
        currentIndex = sidRowIndex + 1
        sidValue = getCloumeValueColumValue(currentIndex, "sid", adSheet)

        if sidValue is not None:
            sidDataPipItem = {}
            mAdResultMap[str(sidValue)] = sidDataPipItem
            adUnitNameValue: str = getCloumeValueColumValue(currentIndex, "广告位名称", adSheet)
            if adUnitNameValue == AD_TYPE_OPEN:
                sidDataPipItem["ad_sw_o"] = True
                sidDataPipItem["ad_sw_n"] = True
                sidDataPipItem["ad_pro_h"] = 0
                sidDataPipItem["ad_new_usr_st"] = native_ad_new_usr_st.get()
                sidDataPipItem["ad_refr_sw"] = False
                sidDataPipItem["ad_refr_it"] = 10000
                sidDataPipItem["ad_refr_t_u"] = 10
                insertListBoxMessage("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            elif adUnitNameValue.find(AD_TYPE_INTERSTITIAL) > 0:
                sidDataPipItem["ad_sw_o"] = True
                sidDataPipItem["ad_sw_n"] = True
                sidDataPipItem["ad_pro_h"] = 0
                sidDataPipItem["ad_new_usr_st"] = native_ad_new_usr_st.get()
                sidDataPipItem["ad_refr_sw"] = False
                sidDataPipItem["ad_refr_it"] = 10000
                sidDataPipItem["ad_refr_t_u"] = 10
                insertListBoxMessage("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            elif adUnitNameValue.find(AD_TYPE_NATIVE) > 0:
                sidDataPipItem["ad_sw_o"] = True
                sidDataPipItem["ad_sw_n"] = True
                sidDataPipItem["ad_pro_h"] = 0
                sidDataPipItem["ad_new_usr_st"] = native_ad_new_usr_st.get()
                sidDataPipItem["ad_refr_sw"] = True
                sidDataPipItem["ad_refr_it"] = native_ad_refr_it.get()
                sidDataPipItem["ad_refr_t_u"] = native_ad_refr_t_u.get()
                insertListBoxMessage("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            else:
                insertListBoxMessage("这个类型不支持" + adUnitNameValue)
                logging.warning("这个类型不支持" + adUnitNameValue)

        # platformValue = getCloumeValueColumValue(currentIndex, "Platform", adSheet)
        # adSourceIdValue = getCloumeValueColumValue(currentIndex, "广告ID", adSheet)
        # adWatingTimeValue = getCloumeValueColumValue(currentIndex, "广告超时时间", adSheet)
        # adExpriedTImeValue = getCloumeValueColumValue(currentIndex, "广告过期时间", adSheet)
        # adDetailsValue = getCloumeValueColumValue(currentIndex, "备注", adSheet)
        # adExtraType = getAdExtraTypeValue(platformValue, adDetailsValue, adSourceIdValue)
        # adPriorityValue = getCloumeValueColumValue(currentIndex, "广告优先级", adSheet)

    pass


def main(excelPath):
    # 开始读取
    wb = openpyxl.load_workbook(excelPath)
    sheetNames = wb.sheetnames
    adSheet = wb[str(sheetNames[0])]
    maxColumn = int(adSheet.max_column)
    insertListBoxMessage("最大列 = " + str(maxColumn))
    insertListBoxMessage("最大行 = " + str(adSheet.max_row))
    print("最大列 = " + str(maxColumn))
    print("最大行 = " + str(adSheet.max_row))

    appLicenseId = getTitleColumValue("app license id", adSheet)
    csjAppId = getTitleColumValue("穿山甲应用ID", adSheet)
    ylhAppId = getTitleColumValue("广点通应用ID", adSheet)
    kshAppId = getTitleColumValue("快手应用ID", adSheet)
    appId = getTitleColumValue("appId", adSheet)

    slot = []

    addChannelIds(slot, adSheet)


def startAndBuild(excelPath):
    logging.basicConfig(level=logging.INFO)
    insertListBoxMessage("XPAD data pip json脚本生成工具")
    logging.info("XPAD data pip json脚本生成工具")
    argLen = len(sys.argv)
    # if argLen < 2:
    #     logging.error("请输入想要解析的文件")
    #     exit()
    # excelPath = sys.argv[1]

    if not os.path.exists(excelPath):
        insertListBoxMessage("需要解析的Excel文件 不存在")

        logging.error("需要解析的Excel文件 不存在")
        exit()

    fileSub = os.path.splitext(os.path.basename(excelPath))[1]
    if fileSub != ".xlsx":
        insertListBoxMessage("请输入正确的Excel文件->" + fileSub)
        logging.error("请输入正确的Excel文件->" + fileSub)
        exit()
    insertListBoxMessage("请确保需要解析的广告数据表在第一个...")
    logging.warning("请确保需要解析的广告数据表在第一个...")
    insertListBoxMessage("开始生成 ..." + str(datetime.datetime.now()))
    logging.info("开始生成 ...")

    main(excelPath)

    jsonResult = json.dumps(mAdResultMap, indent=4, ensure_ascii=False)
    fileName = os.path.splitext(os.path.basename(excelPath))[0]
    fileDir = os.path.dirname(excelPath)
    resultAdFileJson = os.path.join(fileDir, fileName + "_xpad_ad_release_data_pip.json")
    logging.info(resultAdFileJson)

    if os.path.exists(resultAdFileJson):
        resultAdFile = open(resultAdFileJson, 'w')
        resultAdFile.write(jsonResult)
        resultAdFile.close()
    else:
        resultAdFile = open(resultAdFileJson, 'a')
        resultAdFile.write(jsonResult)
        resultAdFile.close()
    insertListBoxMessage("生成成功")
    logging.info("生成成功")


def selectPath():
    global filename
    global path
    filename = filedialog.askopenfilename(filetypes=[("excel格式", "xlsx")])
    insertListBoxMessage("选择路径 :" + filename)
    path.set(filename)


def startCreateJson():
    global listBox
    listBox.delete(0, END)
    global filename
    startAndBuild(filename)
    pass


def creatMainUi():
    global path
    global listBox
    global native_ad_refr_it
    global native_ad_refr_t_u
    global native_ad_new_usr_st

    root = Tk()
    root.title("XPAD data pip json脚本生成工具")
    root.geometry("1000x618")
    root.resizable(False, False)

    path = StringVar()
    native_ad_new_usr_st = IntVar()
    native_ad_new_usr_st.set(253370736000000)

    native_ad_refr_it = IntVar()
    native_ad_refr_it.set(10000)
    native_ad_refr_t_u = IntVar()
    native_ad_refr_t_u.set(10)

    topFrame = Frame(root)
    topFrame.pack(side=TOP)

    Label(topFrame, text="目标路径:").pack(side=LEFT, padx=5, pady=10)
    Entry(topFrame, textvariable=path).pack(side=LEFT, padx=5, pady=10)
    ttk.Button(topFrame, text="路径选择", command=selectPath).pack(side=LEFT, padx=5, pady=10)
    ttk.Button(topFrame, text="开始生成", command=startCreateJson).pack(side=LEFT, padx=5, pady=10)

    middleFrame = Frame(root)
    middleFrame.pack(side=TOP)

    Label(middleFrame, text="新用户时间戳").pack(side=TOP, padx=5, pady=10)
    newUserAdFrame = Frame(middleFrame)
    newUserAdFrame.pack(side=TOP)
    Label(newUserAdFrame, text="ad_new_usr_st").pack(side=LEFT, padx=5, pady=10)
    Entry(newUserAdFrame, textvariable=native_ad_new_usr_st).pack(side=LEFT, padx=5, pady=10)

    Label(middleFrame, text="原生广告配置").pack(side=TOP, padx=5, pady=10)
    nativeAdFrame = Frame(middleFrame)
    nativeAdFrame.pack(side=TOP)

    Label(nativeAdFrame, text="ad_refr_it").pack(side=LEFT, padx=5, pady=10)
    Entry(nativeAdFrame, textvariable=native_ad_refr_it).pack(side=LEFT, padx=5, pady=10)
    Label(nativeAdFrame, text="ad_refr_t_u").pack(side=LEFT, padx=5, pady=10)
    Entry(nativeAdFrame, textvariable=native_ad_refr_t_u).pack(side=LEFT, padx=5, pady=10)

    bottomFrame = Frame(root)
    scrollbar = Scrollbar(bottomFrame)
    scrollbar.pack(side=RIGHT, fill=Y)
    listBox = Listbox(root, yscrollcommand=scrollbar.set)
    scrollbar.config(command=listBox.yview)
    listBox.pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
    bottomFrame.pack(side=TOP, fill=BOTH, expand=YES)

    insertListBoxMessage("欢迎来到XPAD XPAD data pip json脚本生成工具")
    insertListBoxMessage("请选择路径 :")

    root.mainloop()


if __name__ == '__main__':
    argLen = len(sys.argv)

    if argLen < 2:
        creatMainUi()
    else:
        excelPath = sys.argv[1]
        startAndBuild(excelPath)
