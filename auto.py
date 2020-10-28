from pywinauto.application import Application
import time
from pywinauto import keyboard, mouse
import pyperclip
import requests
import zipfile
from xml.dom.minidom import parseString
from datetime import datetime
import xml.etree.ElementTree as et
from xml.etree.ElementTree import fromstring, ElementTree
import xmltodict
import pandas as pd

USERNAME1 = "chencheng"
PASSWORD1 = "cc_137813"
USERNAME2 = "chengtao"
PASSWORD2 = "888888"

D_COORD = {
    "一级菜单-商务管理": (208, 49),
    "商务管理-分析报表": (265, 212),
    "分析报表-同比发货趋势分析": (300, 212),
    "趋势分析-单位选择": (69, 162),
    "趋势分析-产品选择": (240, 162),
    "导出": (320, 72),
    "复制导出链接": (1070, 445),
    "导出成功-确定": (1125, 585),
}


def login():

    # 对于Windows中自带应用程序，直接执行，对于外部应用应输入完整路径
    app = Application(backend="uia").start(r"C:\Program Files (x86)\CRM系统管理\pmgr.exe")
    time.sleep(2)

    # 登录
    keyboard.send_keys(USERNAME1)
    keyboard.send_keys("{VK_TAB}")
    keyboard.send_keys(PASSWORD1)
    keyboard.send_keys("{VK_RETURN}")
    time.sleep(10)
    keyboard.send_keys(USERNAME2)
    keyboard.send_keys("{VK_TAB}")
    keyboard.send_keys(PASSWORD2)
    keyboard.send_keys("{VK_RETURN}")
    time.sleep(3)

    mouse.click(coords=D_COORD["一级菜单-商务管理"])
    mouse.click(coords=D_COORD["商务管理-分析报表"])
    mouse.click(coords=D_COORD["分析报表-同比发货趋势分析"])


def get_data_url():
    # 打开发货分析并获得导出链接
    time.sleep(10)
    mouse.click(coords=D_COORD["导出"])
    time.sleep(2)
    mouse.click(coords=D_COORD["复制导出链接"])
    time.sleep(1)
    mouse.click(coords=D_COORD["导出成功-确定"])
    time.sleep(1)
    excel_url = pyperclip.paste()  # 远程导出excel链接

    # 通过链接处理导出内容
    req = requests.get(excel_url)
    url_content = req.content
    excel_file = open("downloaded.xlsx", "wb")
    excel_file.write(url_content)
    excel_file.close()


def repair_excel():  # CRM系统下载的xlsx有错误，不能使用任何类似xlrd的python excel包处理，要先修复
    # Excel xlsx文件实际上是个zip文件，用zipFile包读取里面实际存放数据的sheet1.xml文件
    with zipfile.ZipFile("downloaded.xlsx", "r") as zbad:
        sheet_xml = zbad.read("sheet1.xml")

    # 以下部分解析sheet1.xml
    dom = parseString(sheet_xml)
    ele_sales = dom.getElementsByTagName("c")  # 读取所有单元格Node元素
    dict_sales = {}
    for i in range(ele_sales.length):
        cells_cr = ele_sales[i].attributes.getNamedItem("r").nodeValue  # 单元格Node的属性"r"是次单元格的行业字符串，类似A1（表示第一行，第一列）的格式
        if ele_sales[i].hasChildNodes():
            sales = ele_sales[i].childNodes[0].childNodes[0].nodeValue  # 单元格Node的孙Node包含这个单元格的值
        else:
            sales = 0  # 如果单元格Node无子Node说明单元格为空
        dict_sales[cells_cr] = sales  # 将所有单元格暂存入字典

    dict_sales_reindex = {}
    for k, v in dict_sales.items():
        if "1" not in k and "8" not in k and "A" not in k and "N" not in k:  # 不要标题行、标题列、汇总行、汇总列，剩下的为纯数据
            dict_sales_reindex[cell_to_timestamp(k)] = v  # 将原字典的key从行列字符串转换为日期

    df = pd.DataFrame.from_dict(dict_sales_reindex, orient="index")  # 转换为pandas dataframe
    df = df.sort_index()  # dataframe日期排序
    print(df)


def cell_to_timestamp(cell_cr):  # 根据单元格行列字符串转换到对应的日期
    STR_C = "BCDEFGHIJKLMN"
    STR_R = "234567"[::-1]
    c = cell_cr[0]
    r = cell_cr[1]
    month = STR_C.index(c) + 1
    year = STR_R.index(r) + 2015
    datestamp = datetime(year=year, month=month, day=1)
    return datestamp


# # 屏幕截图
# img = pag.screenshot()
# open_cv_image = np.array(img)
# open_cv_image = open_cv_image[:, :, ::-1].copy() # Convert RGB to BGR
# cv2.imshow('image',open_cv_image)

# mouse.click(coords=(69, 162))
# mouse.click(coords=(240, 162))


if __name__ == "__main__":
    # cell_to_timestamp('B2')
    # login()
    # get_data_url()
    repair_excel()
