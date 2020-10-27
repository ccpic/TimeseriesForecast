import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import calendar
from chart_func import *
import statsmodels.api as sm


from matplotlib.font_manager import FontProperties
myfont=FontProperties(fname=r'C:\Windows\Fonts\simhei.ttf',size=14)
mpl.rcParams['axes.unicode_minus'] = False

def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year,month)[1])
    return datetime.date(year, month, day)

brand_list = ['泰嘉', '泰加宁', '信立坦', '信达怡', '泰仪']
metric_list = ['金额', '数量']
year = 2020
month = 9

path = '信立坦_数量.xlsx'
df = pd.read_excel(io=path, index_col=0)

df.drop(np.nan, axis=0, inplace=True)
df.drop('合计', axis=1, inplace=True)
list_index = []
list_value = []
for i in range(0, df.shape[0]):
    for j in range(0, df.shape[1]):
        d = datetime.date(int(df.index[i]), int(df.columns[j][:-1]), 1)
        list_index.append(d)
        list_value.append(df.iloc[i,j])

df = pd.Series(list_value, index = list_index)
df = df.sort_index()
df.dropna(inplace=True)
# df= df[35:-1] #最后一个月有发货记录但不是整月，去除
df.index = pd.to_datetime(df.index)
print(df)
# df.plot(figsize=(19, 4))
# plt.show()

decomposition = sm.tsa.seasonal_decompose(df, model='additive')
fig = decomposition.plot()
plt.show()

p = d = q = range(0, 2)
pdq = list(itertools.product(p, d, q))
seasonal_pdq = [(x[0], x[1], x[2], 12) for x in list(itertools.product(p, d, q))]

# for param in pdq:
#     for param_seasonal in seasonal_pdq:
#         try:
#             mod = sm.tsa.statespace.SARIMAX(df,order=param,seasonal_order=param_seasonal,enforce_stationarity=False,enforce_invertibility=False)
#             results = mod.fit()
#             print('ARIMA{}x{}12 - AIC:{}'.format(param,param_seasonal,results.aic))
#         except:
#             continue

mod = sm.tsa.statespace.SARIMAX(df,
                                order=(0, 0, 0),
                                seasonal_order=(1, 0, 1, 3),
                                enforce_stationarity=False,
                                enforce_invertibility=False)
results = mod.fit()
print(results.summary().tables[1])

pred_uc = results.get_forecast(steps=12+(12-month))
pred_ci = pred_uc.conf_int()
ax = df.plot(label='observed', figsize=(14, 9))
pred_uc.predicted_mean.plot(ax=ax, label='Forecast')
ax.fill_between(pred_ci.index,
                pred_ci.iloc[:, 0],
                pred_ci.iloc[:, 1], color='k', alpha=.25)
ax.set_xlabel('Date')
ax.set_ylabel('Sales')
plt.legend()
plt.show()

y_forecasted = pred_uc.predicted_mean
print(y_forecasted.head(12).sum())