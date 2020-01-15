import logging
from tkinter import *

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

csj_id_test = "839119524"
ylh_id_test = "6050194279806441"
ksh_id_test = "5011000106"


class BaseXpadJsonBuild:
    filename = None
    path = None
    listBox: Listbox = None

    def insertListBoxMessage(self, item):
        if self.listBox is not None:
            self.listBox.insert("end", item)
            self.listBox.see(END)

    def findTitleInColum(self, name, adSheet):
        maxColumn = int(adSheet.max_column)
        for columnIndex in range(0, maxColumn):
            titleValue = adSheet.cell(row=1, column=columnIndex + 1).value
            if titleValue == name:
                return columnIndex + 1

        self.insertListBoxMessage("title 不存在 -- " + name)
        logging.error("title 不存在 -- " + name)
        self.insertListBoxMessage("解析将会停止")

    def getTitleColumValue(self, name, adSheet):
        index = self.findTitleInColum(name, adSheet)
        value = adSheet.cell(row=2, column=index).value
        if value is not None:
            return value
        self.insertListBoxMessage("当前 " + name + "值为 None")
        logging.error("当前 " + name + "值为 None")

        return ""

    def getCloumeValueColumValue(self, row, name, adSheet):
        index = self.findTitleInColum(name, adSheet)
        value = adSheet.cell(row=row, column=index).value
        return value

    def hasAdTypeString(self, adDetailsValuelist: list, typeName):
        length = len(adDetailsValuelist)

        for i in range(0, length):
            if adDetailsValuelist[i] == typeName:
                return True

        return False

    def getAdExtraTypeValue(self, platformValue, adDetailsValue, adSourceIdValue):
        adDetailsValuelist: list = adDetailsValue.split("-")
        if len(adDetailsValuelist) < 2:
            self.insertListBoxMessage("备注 命名错误 -- " + str(adSourceIdValue))
            logging.error("备注 命名错误 -- " + str(adSourceIdValue))

        if platformValue == AD_TYPE_CSJ:
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_OPEN):
                return AD_TYPE_CSJ_OPEN
            elif self.hasAdTypeString(adDetailsValuelist, AD_TYPE_INTERSTITIAL):
                return AD_TYPE_CSJ_PERSONAL_PLATE_INTERSTITIAL
            elif self.hasAdTypeString(adDetailsValuelist, AD_TYPE_FULL_SCREEN_VIDEO):
                return AD_TYPE_CSJ_FULL_SCREEN_VIDEO
            elif self.hasAdTypeString(adDetailsValuelist, AD_TYPE_PLATE_RENDER):
                return AD_TYPE_CSJ_PERSONAL_PLATE_FEED
            elif self.hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE_FEED):
                return AD_TYPE_CSJ_FEED
            elif self.hasAdTypeString(adDetailsValuelist, AD_TYPE_BANNER):
                return AD_TYPE_CSJ_PERSONAL_PLATE_BANNER
            else:
                self.insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + " -- 穿山甲的广告id为 = " + str(adSourceIdValue))
                logging.error("不支持这种类型 -- " + adDetailsValue + " -- 穿山甲的广告id为 = " + str(adSourceIdValue))

        elif platformValue == AD_TYPE_YLH:
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_OPEN):
                return AD_TYPE_YLH_OPEN
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE_PLATE):
                return AD_TYPE_YLH_PERSONAL_PLATE_FEED
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_INTERSTITIAL):
                return AD_TYPE_YLH_INTERSTITIAL
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
                return AD_TYPE_YLH_FEED
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_BANNER):
                return AD_TYPE_YLH_BANNER
            else:
                self.insertListBoxMessage("不支持这种类型 --  " + adDetailsValue + "-- 广点通id为 = " + str(adSourceIdValue))
                logging.error("不支持这种类型 --  " + adDetailsValue + "-- 广点通id为 = " + str(adSourceIdValue))

            pass
        elif platformValue == AD_TYPE_KSH:
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
                return AD_TYPE_KSH_FEED
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_PLATE_RENDER):
                return AD_TYPE_KSH_PERSONAL_PLATE_FEED
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_FULL_SCREEN_VIDEO) or self.hasAdTypeString(
                    adDetailsValuelist,
                    AD_TYPE_INTERSTITIAL):
                return AD_TYPE_KSH_FULL_SCREEN_VIDEO
            if self.hasAdTypeString(adDetailsValuelist, AD_TYPE_NATIVE):
                return AD_TYPE_KSH_PERSONAL_PLATE_FEED
            else:
                self.insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + "-- 快手id为 = " + str(adSourceIdValue))
                logging.error("不支持这种类型 -- " + adDetailsValue + "-- 快手id为 = " + str(adSourceIdValue))

            pass
        else:
            self.insertListBoxMessage("没有这个类型 广告id是 = " + str(adSourceIdValue))
            logging.error("没有这个类型 广告id是 = " + str(adSourceIdValue))
            self.insertListBoxMessage("解析将会停止")

        self.insertListBoxMessage("不支持这种类型 -- " + adDetailsValue + "-- id为 = " + str(adSourceIdValue))
        logging.error("不支持这种类型 -- " + adDetailsValue + "-- id为 = " + str(adSourceIdValue))
        return None

    def get_merged_cells_value(self, adSheet, row_index, col_index):
        merged = adSheet.merged_cells
        for (min_col, min_row, max_col, max_row) in merged:
            if (row_index >= min_row[1] and row_index <= max_row[1]):
                if (col_index >= min_col[1] and col_index <= max_col[1]):
                    cell_value = adSheet.cell(min_row[1], min_col[1])
                    # print('该单元格[%d,%d]属于合并单元格，值为[%s]' % (row_index, col_index, cell_value.value))
                    return cell_value.value
        return None

    def checkPlatformValueData(self, platformValue, adSourceIdValue):
        if platformValue == AD_TYPE_CSJ:
            if len(str(adSourceIdValue)) != len(csj_id_test):
                self.insertListBoxMessage(
                    "这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))
                logging.warning("这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))

        elif platformValue == AD_TYPE_YLH:
            if len(str(adSourceIdValue)) != len(ylh_id_test):
                self.insertListBoxMessage(
                    "这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))
                logging.warning("这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))
        elif platformValue == AD_TYPE_KSH:
            if len(str(adSourceIdValue)) != len(ksh_id_test):
                self.insertListBoxMessage(
                    "这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))
                logging.warning("这个id对应的类型不太对- 类型 = " + str(platformValue) + " - 广告id = " + str(adSourceIdValue))
        pass
