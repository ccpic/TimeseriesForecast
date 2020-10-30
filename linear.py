import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import calendar
from chart_func import *


from matplotlib.font_manager import FontProperties

myfont = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf", size=14)
mpl.rcParams["axes.unicode_minus"] = False


def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


# product_list = ["泰嘉", "泰加宁", "信立坦", "信达怡", "泰仪"]
product_list = ["信立坦"]
metric_list = ["金额", "数量"]
year = 2020
month = 10

# for product in product_list:
#     for metric in metric_list:
#         path = product + "_" + metric + ".xlsx"
#         df = pd.read_excel(io=path, index_col=0)
#         print(df)
#         print(product, metric)
#         df.drop(np.nan, axis=0, inplace=True)
#         df.drop("合计", axis=1, inplace=True)
#         list_index = []
#         list_value = []
#         for i in range(0, df.shape[0]):
#             for j in range(0, df.shape[1]):
#                 d = datetime.date(int(df.index[i]), int(df.columns[j][:-1]), 1)
#                 list_index.append(d)
#                 list_value.append(df.iloc[i, j])
#
#         df = pd.Series(list_value, index=list_index)
#         df = df.sort_index()
#         if month < 12:
#             df = df.iloc[: month - 12]
#         ytd = df.iloc[month * (-1) :].sum()
#         ytd_ya = df.iloc[month * (-1) - 12 : -12].sum()
#         df_mat = df.rolling(window=12).sum()
#         df_combined = df_mat
#         print(ytd, ytd_ya)
#
#         for rolling in [12, 6, 3]:
#
#             df_cut = df_mat[rolling * (-1) :]
#             print(df_cut)
#             coef = np.polyfit(x=list(range(1, rolling + 1)), y=df_cut.values, deg=1)
#             p = np.poly1d(coef)
#
#             list_index = []
#             list_value = []
#             for i in range(rolling + 1 - month, 12 * 3 - month + 1 - (12 - rolling)):
#                 list_index.append(add_months(datetime.date(year, month, 1), rolling * (-1) + i))
#                 list_value.append(p(i))
#             df2 = pd.Series(list_value, index=list_index)
#             df_combined = pd.concat([df_combined, df2], axis=1)
#         df_combined.columns = ["Actual-MAT", "Trend-12M", "Trend-6M", "Trend-3M"]
#         df_combined = df_combined.iloc[
#             24:,
#         ]
#         df_combined[df_combined < 0] = np.nan
#         print(df_combined)
#
#         actual_mat_value = df_mat[-1]
#         forecast_value = df_combined.iloc[47, -3:]
#         forecast_value = pd.concat([pd.Series(ytd), pd.Series(actual_mat_value), forecast_value])
#         formatted_forecast_value = ["{:,.0f}".format(member) for member in forecast_value]
#         ytd_gr = ytd / ytd_ya - 1
#         actual_mat_gr = df_mat[-1] / df_mat[-12] - 1
#         forecast_gr = df_combined.iloc[47, -3:] / df_combined.iloc[35, 0] - 1
#         forecast_gr = pd.concat([pd.Series(ytd_gr), pd.Series(actual_mat_gr), forecast_gr])
#         formatted_forecast_gr = ["{:+.1%}".format(member) for member in forecast_gr]
#         table_data = [formatted_forecast_value, formatted_forecast_gr]
#
#         # print(table_data)
#
#         plot_line(
#             df_combined,
#             "plots/" + product + "_" + metric + ".png",
#             table_data=table_data,
#             title=product + "（" + metric + "）" + "发货表现及趋势预测",
#         )


def liner_forecast(df, latest_year, latest_month, product, metric):
    df = df.astype(float)
    if latest_month < 12:
        df = df.iloc[: latest_month - 12]
    ytd = df.iloc[latest_month * (-1):].sum()
    ytd_ya = df.iloc[latest_month * (-1) - 12: -12].sum()
    df_mat = df.rolling(window=12).sum()
    df_combined = df_mat
    # print(ytd, ytd_ya)

    for rolling in [12, 6, 3]:

        df_cut = df_mat[rolling * (-1):]
        coef = np.polyfit(x=list(range(1, rolling + 1)), y=df_cut.values, deg=1)
        p = np.poly1d(coef)

        list_index = []
        list_value = []
        for i in range(rolling + 1 - latest_month, 12 * 3 - latest_month + 1 - (12 - rolling)):
            list_index.append(add_months(datetime.date(latest_year, latest_month, 1), rolling * (-1) + i))
            list_value.append(p(i))
        df2 = pd.Series(list_value, index=list_index)
        df_combined = pd.concat([df_combined, df2], axis=1)
    df_combined.columns = ["Actual-MAT", "Trend-12M", "Trend-6M", "Trend-3M"]
    df_combined = df_combined.iloc[
                  24:,
                  ]
    df_combined[df_combined < 0] = np.nan
    print(df_combined)

    actual_mat_value = df_mat[-1]
    forecast_value = df_combined.iloc[47, -3:]
    forecast_value = pd.concat([pd.Series(ytd), pd.Series(actual_mat_value), forecast_value])
    print(forecast_value)
    formatted_forecast_value = ["{:,.0f}".format(float(member)) for member in forecast_value]
    ytd_gr = ytd / ytd_ya - 1
    actual_mat_gr = df_mat[-1] / df_mat[-12] - 1
    forecast_gr = df_combined.iloc[47, -3:] / df_combined.iloc[35, 0] - 1
    forecast_gr = pd.concat([pd.Series(ytd_gr), pd.Series(actual_mat_gr), forecast_gr])
    formatted_forecast_gr = ["{:+.1%}".format(float(member)) for member in forecast_gr]
    table_data = [formatted_forecast_value, formatted_forecast_gr]

    # print(table_data)

    plot_line(
        df_combined,
        "plots/" + product + "_" + metric + ".png",
        table_data=table_data,
        title=product + "（" + metric + "）" + "发货表现及趋势预测",
    )


if __name__ == "__main__":
    for product in product_list:
        for metric in metric_list:
            path = product + "_" + metric + ".xlsx"
            df = pd.read_excel(io=path, index_col=0)
            print(df)
            print(product, metric)
            df.drop(np.nan, axis=0, inplace=True)
            df.drop("合计", axis=1, inplace=True)
            list_index = []
            list_value = []
            for i in range(0, df.shape[0]):
                for j in range(0, df.shape[1]):
                    d = datetime.date(int(df.index[i]), int(df.columns[j][:-1]), 1)
                    list_index.append(d)
                    list_value.append(df.iloc[i, j])

            df = pd.Series(list_value, index=list_index)
            df = df.sort_index()
            liner_forecast(df, 2020, 9, product, metric)