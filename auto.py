from pywinauto.application import Application
import time
from pywinauto import keyboard, mouse
import pyperclip
import requests
import zipfile
from xml.dom.minidom import parseString
from datetime import datetime
import numpy as np
import pandas as pd
from linear import liner_forecast
import pyautogui

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
    "查询": (755, 135),
    "导出": (320, 72),
    "复制导出链接": (1070, 445),
    "导出成功-确定": (1125, 585),
    "单位选择下拉": (69, 164),
    "产品选择下拉": (240, 164),
    "金额": (46, 180),
    "数量": (46, 194),
    "折算标准数量": (681, 168),
    "产品选择确定": (227, 354),
    "产品选择全清": (124, 354),
    "泰嘉": (94, 186),
    "信立坦": (94, 226),
    "泰加宁": (94, 206),
    "信达怡": (94, 246),
    "泰仪": (94, 326),
    "判断色块_加载": (182, 914),
    "判断色块_导出": (934, 317),
}


def check_loaded(pixel, color):
    time.sleep(1)
    while True:
        im = pyautogui.screenshot()
        check = im.getpixel(pixel)
        time.sleep(1)
        print(check)
        if check == color:
            break
    time.sleep(1)


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

    # 打开发货分析界面
    mouse.click(coords=D_COORD["一级菜单-商务管理"])
    time.sleep(1)
    mouse.click(coords=D_COORD["商务管理-分析报表"])
    time.sleep(1)
    mouse.click(coords=D_COORD["分析报表-同比发货趋势分析"])
    # time.sleep(10)
    check_loaded(D_COORD["判断色块_加载"], (255, 0, 0))


def get_data_url(product, metric):

    # 根据品牌和单位获取数据下载链接
    if product != "所有产品总计":
        mouse.click(coords=D_COORD["产品选择下拉"])
        time.sleep(1)
        mouse.click(coords=D_COORD["产品选择全清"])
        time.sleep(1)
        mouse.click(coords=D_COORD[product])
        time.sleep(1)
        mouse.click(coords=D_COORD["产品选择确定"])
        time.sleep(1)

    mouse.click(coords=D_COORD["单位选择下拉"])
    time.sleep(1)
    mouse.click(coords=D_COORD[metric])  # 选择单位（金额/数量）
    time.sleep(1)
    if metric == "数量":
        mouse.click(coords=D_COORD["折算标准数量"])  # 如果单位是数量选择折算标准数量
        time.sleep(1)

    mouse.click(coords=D_COORD["查询"])
    check_loaded(D_COORD["判断色块_加载"], (255, 0, 0))
    mouse.click(coords=D_COORD["导出"])
    check_loaded(D_COORD["判断色块_导出"], (0, 0, 132))
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


def repair_excel(product):  # CRM系统下载的xlsx有错误，不能使用任何类似xlrd的python excel包处理，要先修复
    # Excel xlsx文件实际上是个zip文件，用zipFile包读取里面实际存放数据的sheet1.xml文件
    with zipfile.ZipFile("downloaded.xlsx", "r") as zbad:
        sheet_xml = zbad.read("sheet1.xml")

    # 以下部分解析sheet1.xml
    dom = parseString(sheet_xml)
    ele_sales = dom.getElementsByTagName("c")  # 读取所有单元格Node元素（tag为c）
    dict_sales = {}
    for i in range(ele_sales.length):
        cells_cr = ele_sales[i].attributes.getNamedItem("r").nodeValue  # 单元格Node的属性"r"是次单元格的行业字符串，类似A1（表示第一行，第一列）的格式
        if ele_sales[i].hasChildNodes():
            sales = ele_sales[i].childNodes[0].childNodes[0].nodeValue  # 单元格Node的孙Node包含这个单元格的值
        else:
            sales = np.nan  # 如果单元格Node无子Node说明单元格为空
        dict_sales[cells_cr] = sales  # 将所有单元格暂存入字典

    list_value = []
    list_index = []
    for k, v in dict_sales.items():
        if "1" not in k and "8" not in k and "A" not in k and "N" not in k:  # 不要标题行、标题列、汇总行、汇总列，剩下的为纯数据
            list_value.append(v)
            list_index.append(cell_to_timestamp(k))  # 将原字典的key从行列字符串转换为日期

    df = pd.Series(list_value, index=list_index)  # 转换为pandas series
    df.name = product
    df = df.sort_index()  # dataframe日期排序

    return df


def cell_to_timestamp(cell_cr):  # 根据单元格行列字符串转换到对应的日期
    STR_C = "BCDEFGHIJKLMN"
    STR_R = "234567"[::-1]
    c = cell_cr[0]
    r = cell_cr[1]
    month = STR_C.index(c) + 1
    year = STR_R.index(r) + 2015
    datestamp = datetime(year=year, month=month, day=1)
    return datestamp


if __name__ == "__main__":
    year = 2020
    month = 10
    # product_list = ["泰嘉", "泰加宁"]
    # metric_list = ["金额", "数量"]
    product_list = ["信立坦"]
    metric_list = ["金额"]

    login()

    now = datetime.now().strftime("%Y%m%d%H%M%S")
    writer = pd.ExcelWriter("data/" + now + ".xlsx", engine="xlsxwriter")
    for i, product in enumerate(product_list):
        for metric in metric_list:
            get_data_url(product, metric)
            df = repair_excel(product)
            if i == 0:
                show_index = True
                start_col = i
            else:
                show_index = False
                start_col = i + 1
            df.to_excel(writer, sheet_name=metric, startrow=0, startcol=start_col, index=show_index)
            liner_forecast(
                df,
                year,
                month,
                product,
                metric,
                timestamp=datetime.strptime(now, "%Y%m%d%H%M%S").strftime("%Y-%m-%d %H:%M:%S"),
            )
    writer.save()
