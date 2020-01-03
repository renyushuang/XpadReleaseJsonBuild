# coding=utf-8
import openpyxl, os, json, sys, logging
from tkinter import *
from tkinter import filedialog
import datetime

AD_TYPE_OPEN = "开屏"
AD_TYPE_NATIVE_PLATE = "原生模版"
AD_TYPE_PLATE_RENDER = "模版渲染"
AD_TYPE_PLATE = "模版"
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


def insertListBoxMessage(item):
    global listBox
    listBox.insert("end", item)
    listBox.see(END)


def findTitleInColum(name, adSheet):
    maxColumn = int(adSheet.max_column)
    for columnIndex in range(0, maxColumn):
        titleValue = adSheet.cell(row=1, column=columnIndex + 1).value
        if titleValue == name:
            return columnIndex + 1

    insertListBoxMessage("title 不存在 -- " + name)
    logging.error("title 不存在 -- " + name)
    insertListBoxMessage("解析将会停止")


def getTitleColumValue(name, adSheet):
    index = findTitleInColum(name, adSheet)
    value = adSheet.cell(row=2, column=index).value
    if value is not None:
        return value
    insertListBoxMessage("当前 " + name + "值为 None")
    logging.error("当前 " + name + "值为 None")

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
        insertListBoxMessage("备注 命名错误 -- " + str(adSourceIdValue))
        logging.error("备注 命名错误 -- " + str(adSourceIdValue))

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
            insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + " -- 穿山甲的广告id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 -- " + adDetailsValue + " -- 穿山甲的广告id为 = " + str(adSourceIdValue))

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
            insertListBoxMessage("不支持这种类型 --  " + adDetailsValue + "-- 广点通id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 --  " + adDetailsValue + "-- 广点通id为 = " + str(adSourceIdValue))

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
            insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + "-- 快手id为 = " + str(adSourceIdValue))
            logging.error("不支持这种类型 -- " + adDetailsValue + "-- 快手id为 = " + str(adSourceIdValue))

        pass
    else:
        insertListBoxMessage("没有这个类型 广告id是 = " + str(adSourceIdValue))
        logging.error("没有这个类型 广告id是 = " + str(adSourceIdValue))
        insertListBoxMessage("解析将会停止")
    insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + "-- id为 = " + str(adSourceIdValue))
    logging.error("不支持这种类型 -- " + adDetailsValue + "-- id为 = " + str(adSourceIdValue))
    return None


def get_merged_cells_value(adSheet, row_index, col_index):
    merged = adSheet.merged_cells
    for (min_col, min_row, max_col, max_row) in merged:
        if (row_index >= min_row[1] and row_index <= max_row[1]):
            if (col_index >= min_col[1] and col_index <= max_col[1]):
                cell_value = adSheet.cell(min_row[1], min_col[1])
                # print('该单元格[%d,%d]属于合并单元格，值为[%s]' % (row_index, col_index, cell_value.value))
                return cell_value.value
    return None


def addChannelIds(slot, adSheet):
    maxColum = adSheet.max_row

    adSidItem = {}
    extra = {}
    for sidRowIndex in range(1, maxColum):
        currentIndex = sidRowIndex + 1
        sidValue = getCloumeValueColumValue(currentIndex, "sid", adSheet)

        if sidValue is not None:
            adSidItem = {}
            extra = {}
            slot[str(sidValue)] = adSidItem
            adSidItem["extra"] = extra
            adSidItem["sid"] = sidValue
            adSidItem["size"] = ""
            adSidItem["status"] = 1
            adWatingTimeValue = getCloumeValueColumValue(currentIndex, "广告超时时间", adSheet)
            if adWatingTimeValue is None:
                adWatingTimeValue = 4000
            elif adWatingTimeValue < 1000:
                insertListBoxMessage("广告超时时间 是毫秒的 =" + str(sidValue) + "--对应 = " + adWatingTimeValue + " 是错误的")
                logging.error("广告超时时间 是毫秒的 =" + str(sidValue) + "--对应 = " + adWatingTimeValue + " 是错误的")

            adSidItem["wt"] = adWatingTimeValue

            adPriorityValue = getCloumeValueColumValue(currentIndex, "广告优先级", adSheet)
            if adPriorityValue is None:
                adPriorityValue = ""
            adSidItem["st"] = adPriorityValue
            adUnitNameValue: str = getCloumeValueColumValue(currentIndex, "广告位名称", adSheet)

            if adUnitNameValue == AD_TYPE_OPEN or adUnitNameValue.find(AD_TYPE_NATIVE) > 0:
                adSidItem["type"] = 4
                extra["ylh_pixel_w"] = 1280
                extra["ylh_pixel_h"] = 720

            elif adUnitNameValue.find(AD_TYPE_INTERSTITIAL) > 0:
                adSidItem["type"] = 2
                extra["image_ratio"] = "2:3"
        else:
            # 跳过不属于当前单元格的输出
            merged_value = get_merged_cells_value(adSheet, currentIndex, 4)
            if merged_value != adSidItem["sid"]:
                continue

        platformValue = getCloumeValueColumValue(currentIndex, "Platform", adSheet)
        if platformValue is None:
            insertListBoxMessage(
                "Platform is None -- sid = " + str(adSidItem["sid"]) + "----- 行 = " + str(currentIndex))
            logging.error("Platform is None -- sid = " + str(adSidItem["sid"]) + "----- 行 = " + str(currentIndex))
            continue

        adSourceIdValue = getCloumeValueColumValue(currentIndex, "广告ID", adSheet)
        if adSourceIdValue is None:
            insertListBoxMessage("广告ID is None -- sid = " + str(adSidItem["sid"]) + "----- 行 = " + str(currentIndex))
            logging.error("广告ID is None -- sid = " + str(adSidItem["sid"]) + "----- 行 = " + str(currentIndex))
            continue

        adDetailsValue: str = getCloumeValueColumValue(currentIndex, "备注", adSheet)
        if adDetailsValue is None:
            insertListBoxMessage("备注内容为None  跳过 = " + str(adSourceIdValue))
            logging.warning("备注内容为None  跳过 = " + str(adSourceIdValue))
            continue

        adExtraType = getAdExtraTypeValue(platformValue, adDetailsValue, adSourceIdValue)
        if adSidItem.get("type") == 4:
            adSidItem[str(platformValue)] = adSourceIdValue
            if adDetailsValue.find(AD_TYPE_PLATE) > 0:
                if platformValue == AD_TYPE_YLH:
                    extra["subtype"] = 1
                else:
                    key = "subtype_" + platformValue
                    extra[key] = 1



            elif adDetailsValue.find(AD_TYPE_NATIVE) > 0:
                if platformValue == AD_TYPE_YLH:
                    extra["subtype"] = 2
                else:
                    key = "subtype_" + platformValue
                    extra[key] = 2
            elif adDetailsValue.find(AD_TYPE_OPEN) > 0:
                extra["subtype"] = 1
                key = "subtype_" + platformValue
                extra[key] = 1


        elif adSidItem.get("type") == 2:
            if adDetailsValue.find(AD_TYPE_INTERSTITIAL) > 0:
                adSidItem[str(platformValue)] = adSourceIdValue
            elif adDetailsValue.find(AD_TYPE_FULL_SCREEN_VIDEO) > 0:
                key = str(platformValue) + "_ext"
                adSidItem[key] = adSourceIdValue


def main(excelPath):
    # 开始读取
    wb = openpyxl.load_workbook(excelPath)
    sheetNames = wb.sheetnames
    adSheet = wb[str(sheetNames[0])]
    maxColumn = int(adSheet.max_column)

    insertListBoxMessage("最大列 = " + str(maxColumn))
    insertListBoxMessage("最大列 = " + str(maxColumn))
    logging.info("最大列 = " + str(maxColumn))
    logging.info("最大行 = " + str(adSheet.max_row))

    appLicenseId = getTitleColumValue("app license id", adSheet)
    csjAppId = getTitleColumValue("穿山甲应用ID", adSheet)
    ylhAppId = getTitleColumValue("广点通应用ID", adSheet)
    kshAppId = getTitleColumValue("快手应用ID", adSheet)
    appId = getTitleColumValue("appId", adSheet)

    mAdResultMap["ls"] = appLicenseId
    mAdResultMap["csj"] = csjAppId
    mAdResultMap["ylh"] = ylhAppId
    mAdResultMap["ksh"] = kshAppId
    mAdResultMap["appid"] = appId

    slot = {}
    mAdResultMap["slot"] = slot

    addChannelIds(slot, adSheet)

    mAdResultMap["status"] = 1


def startAndBuild(excelPath):
    logging.basicConfig(level=logging.INFO)
    insertListBoxMessage("XPAD 1.0 json脚本生成工具")
    logging.info("XPAD 1.0 json脚本生成工具")
    # argLen = len(sys.argv)
    # if argLen < 2:
    #     logging.error("请输入想要解析的文件")
    #     exit()
    # excelPath = sys.argv[1]

    if not os.path.exists(excelPath):
        insertListBoxMessage("需要解析的Excel文件 不存在")
        logging.error("需要解析的Excel文件 不存在")
        insertListBoxMessage("解析将会停止")

    fileSub = os.path.splitext(os.path.basename(excelPath))[1]
    if fileSub != ".xlsx":
        insertListBoxMessage("请输入正确的Excel文件->" + fileSub)
        logging.error("请输入正确的Excel文件->" + fileSub)
        insertListBoxMessage("解析将会停止")

    insertListBoxMessage("请确保需要解析的广告数据表在第一个...")
    insertListBoxMessage("开始生成 ..." + str(datetime.datetime.now()))

    logging.warning("请确保需要解析的广告数据表在第一个...")
    logging.info("开始生成 ...")
    main(excelPath)

    jsonResult = json.dumps(mAdResultMap, indent=4, ensure_ascii=False)
    fileName = os.path.splitext(os.path.basename(excelPath))[0]
    fileDir = os.path.dirname(excelPath)
    resultAdFileJson = os.path.join(fileDir, fileName + "_xpad_ad_release_1.0.json")
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

    root = Tk()
    root.title("XPAD 1.0 json脚本生成工具")
    root.geometry("1000x618")
    root.resizable(False, False)

    path = StringVar()

    topFrame = Frame(root)
    topFrame.pack(side=TOP)

    Label(topFrame, text="目标路径:").pack(side=LEFT, padx=5, pady=10)
    Entry(topFrame, textvariable=path).pack(side=LEFT, padx=5, pady=10)
    Button(topFrame, text="路径选择", command=selectPath).pack(side=LEFT, padx=5, pady=10)
    Button(topFrame, text="开始生成", command=startCreateJson).pack(side=LEFT, padx=5, pady=10)

    bottomFrame = Frame(root)
    scrollbar = Scrollbar(bottomFrame)
    scrollbar.pack(side=RIGHT, fill=Y)
    listBox = Listbox(root, yscrollcommand=scrollbar.set)
    scrollbar.config(command=listBox.yview)
    listBox.pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
    bottomFrame.pack(side=TOP, fill=BOTH, expand=YES)
    insertListBoxMessage("欢迎来到XPAD 1.0 json脚本生成工具")
    insertListBoxMessage("请选择路径 :")
    root.mainloop()


if __name__ == '__main__':
    # startAndBuild()
    creatMainUi()
