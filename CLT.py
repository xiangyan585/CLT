import xlrd
import xlwt
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.pyplot as plt2
from matplotlib.font_manager import FontProperties
import random
from xlutils.copy import copy

# 设置字体，避免乱码
font_set = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=12)
plt.title('中心极限定理', fontproperties=font_set)
plt2.title('中心极限定理2', fontproperties=font_set)

# 设置精度，只留一个小数位
np.set_printoptions(precision=0)

# 读取EXCEL 表格 
data = xlrd.open_workbook('20nm RP_0.5mM_5M_150mv.xlsx')
table = data.sheet_by_name('150mv')
tRows = table.nrows
tCols = table.ncols

TotalNumber  = 50000
ChooseNumber  = 300 
#从原始数据中选取 m 个，生成n组数据
raw_data_duration =table.col_values(5)
raw_data_amplitude =table.col_values(6)

samples_mean_duration = []
samples_mean_amplitude = []

for i in range(0,TotalNumber):
    samples_duration = []
    samples_amplitude = []
    samples_duration = random.sample(raw_data_duration, ChooseNumber)  #从list中随机获取5个元素，作为一个片断返回 
    samples_amplitude = random.sample(raw_data_amplitude, ChooseNumber)  #从list中随机获取5个元素，作为一个片断返回 
    samples_mean_duration.append(sum(samples_duration)/ChooseNumber)
    samples_mean_amplitude.append(sum(samples_amplitude)/ChooseNumber)

#写入数据
new_workbook = copy(data)
new_worksheet = new_workbook.get_sheet(0)
for j in range(0,TotalNumber):
    new_worksheet.write(j,9, samples_mean_duration[j])
    new_worksheet.write(j,10, samples_mean_amplitude[j])
new_workbook.save('CLT_150mv.xls')
# 转换下格式
samples_mean_np = np.array(samples_mean_duration)
samples_mean_np2 = np.array(samples_mean_amplitude)


# 画图
print(samples_mean_np)
plt.hist(samples_mean_np, bins=100)
plt.show()

# 画图
print(samples_mean_np2)
plt2.hist(samples_mean_np2, bins=100)

plt2.show()
