# coding=utf-8
import openpyxl, os, json, sys, logging

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


def findTitleInColum(name, adSheet):
    maxColumn = int(adSheet.max_column)
    for columnIndex in range(0, maxColumn):
        titleValue = adSheet.cell(row=1, column=columnIndex + 1).value
        if titleValue == name:
            return columnIndex + 1

    logging.error("title 不存在 -" + name)
    exit()


def getTitleColumValue(name, adSheet):
    index = findTitleInColum(name, adSheet)
    value = adSheet.cell(row=2, column=index).value
    if value is not None:
        return value

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
            logging.error("不支持这种类型 " + adDetailsValue + "快手id为 = " + str(adSourceIdValue))

        pass
    else:
        logging.error("没有这个类型 广告id是 = " + str(adSourceIdValue))
        exit()

    logging.error("不支持这种类型 " + adDetailsValue + "id为 = " + str(adSourceIdValue))
    return None


def addChannelIds(slot: list, adSheet):
    maxColum = adSheet.max_row

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
                sidDataPipItem["ad_new_usr_st"] = 253370736000000
                sidDataPipItem["ad_refr_sw"] = False
                sidDataPipItem["ad_refr_it"] = 10000
                sidDataPipItem["ad_refr_t_u"] = 10
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            elif adUnitNameValue.find(AD_TYPE_INTERSTITIAL) > 0:
                sidDataPipItem["ad_sw_o"] = True
                sidDataPipItem["ad_sw_n"] = True
                sidDataPipItem["ad_pro_h"] = 0
                sidDataPipItem["ad_new_usr_st"] = 253370736000000
                sidDataPipItem["ad_refr_sw"] = False
                sidDataPipItem["ad_refr_it"] = 10000
                sidDataPipItem["ad_refr_t_u"] = 10
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            elif adUnitNameValue.find(AD_TYPE_NATIVE) > 0:
                sidDataPipItem["ad_sw_o"] = True
                sidDataPipItem["ad_sw_n"] = True
                sidDataPipItem["ad_pro_h"] = 0
                sidDataPipItem["ad_new_usr_st"] = 253370736000000
                sidDataPipItem["ad_refr_sw"] = True
                sidDataPipItem["ad_refr_it"] = 10000
                sidDataPipItem["ad_refr_t_u"] = 10
                logging.info("" + adUnitNameValue + "----" + str(sidDataPipItem["ad_refr_sw"]))
            else:
                logging.warning("这个类型不支持" + adUnitNameValue)

        # platformValue = getCloumeValueColumValue(currentIndex, "Platform", adSheet)
        # adSourceIdValue = getCloumeValueColumValue(currentIndex, "广告ID", adSheet)
        # adWatingTimeValue = getCloumeValueColumValue(currentIndex, "广告超时时间", adSheet)
        # adExpriedTImeValue = getCloumeValueColumValue(currentIndex, "广告过期时间", adSheet)
        # adDetailsValue = getCloumeValueColumValue(currentIndex, "备注", adSheet)
        # adExtraType = getAdExtraTypeValue(platformValue, adDetailsValue, adSourceIdValue)
        # adPriorityValue = getCloumeValueColumValue(currentIndex, "广告优先级", adSheet)

    pass


def main():
    # 开始读取
    wb = openpyxl.load_workbook(excelPath)
    sheetNames = wb.sheetnames
    adSheet = wb[str(sheetNames[0])]
    maxColumn = int(adSheet.max_column)

    print("最大列 = " + str(maxColumn))
    print("最大行 = " + str(adSheet.max_row))

    appLicenseId = getTitleColumValue("app license id", adSheet)
    csjAppId = getTitleColumValue("穿山甲应用ID", adSheet)
    ylhAppId = getTitleColumValue("广点通应用ID", adSheet)
    kshAppId = getTitleColumValue("快手应用ID", adSheet)
    appId = getTitleColumValue("appId", adSheet)

    slot = []

    addChannelIds(slot, adSheet)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    logging.info("XPAD 2.0 json脚本生成工具")
    argLen = len(sys.argv)
    if argLen < 2:
        logging.error("请输入想要解析的文件")
        exit()
    excelPath = sys.argv[1]

    if not os.path.exists(excelPath):
        logging.error("需要解析的Excel文件 不存在")
        exit()

    fileSub = os.path.splitext(os.path.basename(excelPath))[1]
    if fileSub != ".xlsx":
        logging.error("请输入正确的Excel文件->" + fileSub)
        exit()
    logging.warning("请确保需要解析的广告数据表在第一个...")
    logging.info("开始生成 ...")

    main()

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
    logging.info("生成成功")
