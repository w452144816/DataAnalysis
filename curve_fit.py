# import os
import numpy as np
# # from scipy import log
# from scipy.optimize import curve_fit
# import matplotlib.pyplot as plt
import math
from sklearn.metrics import r2_score
#
# # 字体
# plt.rcParams['font.sans-serif'] = ['SimHei']
#
#
# # 拟合函数
# def func(x, a, b):
#     #    y = a * log(x) + b
#     y = x / (a * x + b)
#     return y
#
#
# # 拟合的坐标点
# # x0 = [24.9704, 24.01, 23.0496, 21.1288, 18.2476, 13.4456, 7.6832, 0.9604]
# # y0 = [3.92, 7.84, 11.76, 15.68, 19.6, 23.52, 27.44, 31.36]
# x0 = [2, 4, 8, 10, 24, 28, 32, 48]
# y0 = [6.66,8.35,10.81,11.55,13.63,13.68,13.69,13.67]
# # 拟合，可选择不同的method
# result = curve_fit(func, x0, y0, method='trf')
# a, b = result[0]
#
# # 绘制拟合曲线用
# x1 = np.arange(2, 48, 0.1)
# # y1 = a * log(x1) + b
# y1 = x1 / (a * x1 + b)
#
# x0 = np.array(x0)
# y0 = np.array(y0)
# # 计算r2
# y2 = x0 / (a * x0 + b)
# # y2 = a * log(x0) + b
# r2 = r2_score(y0, y2)
#
# # plt.figure(figsize=(7.5, 5))
# # 坐标字体大小
# plt.tick_params(labelsize=11)
# # 原数据散点
# plt.scatter(x0, y0, s=30, marker='o')
#
# # 横纵坐标起止
# plt.xlim((0, 50))
# plt.ylim((0, round(max(y0)) + 2))
#
# # 拟合曲线
# plt.plot(x1, y1, "blue")
# plt.title("标题", fontsize=13)
# plt.xlabel('X（h）', fontsize=12)
# plt.ylabel('Y（%）', fontsize=12)
#
# # 指定点，y=9时求x
# p = round(9 * b / (1 - 9 * a), 2)
# # p = b/(math.log(9/a))
# p = round(p, 2)
# # 显示坐标点
# plt.scatter(p, 9, s=20, marker='x')
# # 显示坐标点横线、竖线
# plt.vlines(p, 0, 9, colors="c", linestyles="dashed")
# plt.hlines(9, 0, p, colors="c", linestyles="dashed")
# # 显示坐标点坐标值
# plt.text(p, 9, (float('%.2f' % p), 9), ha='left', va='top', fontsize=11)
# # 显示公式
# m = round(max(y0) / 10, 1)
# print(m)
# plt.text(48, m, 'y= x/(' + str(round(a, 2)) + '*x+' + str(round(b, 2)) + ')', ha='right', fontsize=12)
# plt.text(48, m, r'$R^2=$' + str(round(r2, 3)), ha='right', va='top', fontsize=12)
#
# # True 显示网格
# # linestyle 设置线显示的类型(一共四种)
# # color 设置网格的颜色
# # linewidth 设置网格的宽度
# plt.grid(True, linestyle="--", color="g", linewidth="0.5")
# plt.show()
from sklearn import preprocessing

def __sst(y_no_fitting):
    """
    计算SST(total sum of squares) 总平方和
    :param y_no_predicted: List[int] or array[int] 待拟合的y
    :return: 总平方和SST
    """
    y_mean = sum(y_no_fitting) / len(y_no_fitting)
    s_list =[(y - y_mean)**2 for y in y_no_fitting]
    sst = sum(s_list)
    return sst


def __ssr(y_fitting, y_no_fitting):
    """
    计算SSR(regression sum of squares) 回归平方和
    :param y_fitting: List[int] or array[int]  拟合好的y值
    :param y_no_fitting: List[int] or array[int] 待拟合y值
    :return: 回归平方和SSR
    """
    y_mean = sum(y_no_fitting) / len(y_no_fitting)
    s_list =[(y - y_mean)**2 for y in y_fitting]
    ssr = sum(s_list)
    return ssr


def __sse(y_fitting, y_no_fitting):
    """
    计算SSE(error sum of squares) 残差平方和
    :param y_fitting: List[int] or array[int] 拟合好的y值
    :param y_no_fitting: List[int] or array[int] 待拟合y值
    :return: 残差平方和SSE
    """
    s_list = [(y_fitting[i] - y_no_fitting[i])**2 for i in range(len(y_fitting))]
    sse = sum(s_list)
    return sse

def goodness_of_fit(y_fitting, y_no_fitting):
    """
    计算拟合优度R^2
    :param y_fitting: List[int] or array[int] 拟合好的y值
    :param y_no_fitting: List[int] or array[int] 待拟合y值
    :return: 拟合优度R^2
    """
    SSR = __ssr(y_fitting, y_no_fitting)
    SST = __sst(y_no_fitting)
    rr = SSR /SST
    return rr

import numpy as np
import matplotlib.pyplot as plt
from scipy import optimize
from scipy.optimize import curve_fit

def data_scale(data):
    digit_count = len(str(data[0]))
    return [x / (np.power(10, digit_count - 1)) for x in data], digit_count - 1

def un_data_scale(data, scale):
    return [x * np.power(10, scale) for x in data]


# 忽略除以0的报错
np.seterr(divide='ignore', invalid='ignore')

# 拟合数据集
# x = [4, 8, 10, 12, 25, 32, 43, 58, 63, 69, 79]
# y = [20, 33, 50, 56, 42, 31, 33, 46, 65, 75, 78]
# x = [3.92, 7.84, 11.76, 15.68, 19.6, 23.52, 27.44, 31.36]
# y = [24.9704, 24.01, 23.0496, 21.1288, 18.2476, 13.4456, 7.6832, 0.9604]

x = [7728, 8855, 9928, 11054, 12128, 13255, 14328, 15455]
y = [3187, 3145, 3074, 2962, 2792, 2567, 2335, 2019]
# x = [7.728, 8.855, 9.928, 11.054, 12.128, 13.255, 14.328, 15.455]
# y = [3.187, 3.145, 3.074, 2.962, 2.792, 2.567, 2.335, 2.019]

normalized_xdata, x_scale = data_scale(x)
normalized_ydata, y_scale = data_scale(y)

x_arr = np.array(normalized_xdata)
y_arr = np.array(normalized_ydata)

# 一阶函数方程(直线)
def func_1(x, a, b):
    return a*x + b


# 二阶曲线方程
def func_2(x, a, b, c):
    return a * np.power(x, 2) + b * x + c


# 三阶曲线方程
def func_3(x, a, b, c, d):
    return a * np.power(x, 3) + b * np.power(x, 2) + c * x + d


# 四阶曲线方程
def func_4(x, a, b, c, d, e):
    return a * np.power(x, 4) + b * np.power(x, 3) + c * np.power(x, 2) + d * x + e

# 五阶曲线方程
def func_5(x, a, b, c, d, e, f):
    return a * np.power(x, 5) + b * np.power(x, 4) + c * np.power(x, 3) + d * np.power(x, 2) * x + e + f


# 拟合参数都放在popt里，popt是个数组，参数顺序即你自定义函数中传入的参数的顺序
popt1, pcov1 = curve_fit(func_1, x_arr, y_arr)
a1 = popt1[0]
b1 = popt1[1]
popt2, pcov2 = curve_fit(func_2, x_arr, y_arr)
a2 = popt2[0]
b2 = popt2[1]
c2 = popt2[2]
popt3, pcov3 = curve_fit(func_3, x_arr, y_arr)
a3 = popt3[0]
b3 = popt3[1]
c3 = popt3[2]
d3 = popt3[3]
popt4, pcov4 = curve_fit(func_4, x_arr, y_arr,method='trf')
a4 = popt4[0]
b4 = popt4[1]
c4 = popt4[2]
d4 = popt4[3]
e4 = popt4[4]
popt5, pcov5 = curve_fit(func_5, x_arr, y_arr,method='trf')
a5 = popt5[0]
b5 = popt5[1]
c5 = popt5[2]
d5 = popt5[3]
e5 = popt5[4]
f5 = popt5[5]

yvals1 = func_1(x_arr, a1, b1)
print("一阶拟合数据为: ", yvals1)
yvals2 = func_2(x_arr, a2, b2, c2)
print("二阶拟合数据为: ", yvals2)
yvals3 = func_3(x_arr, a3, b3, c3, d3)
print("三阶拟合数据为: ", yvals3)
yvals4 = func_4(x_arr, a4, b4, c4, d4, e4)
print("四阶拟合数据为: ", yvals4)
yvals5 = func_5(x_arr, a5, b5, c5, d5, e5, f5)
print("五阶拟合数据为: ", yvals5)

rr1 = goodness_of_fit(yvals1, normalized_ydata)
print("一阶曲线拟合优度为%.5f" % rr1)
rr2 = goodness_of_fit(yvals2, normalized_ydata)
print("二阶曲线拟合优度为%.5f" % rr2)
rr3 = goodness_of_fit(yvals3, normalized_ydata)
print("三阶曲线拟合优度为%.5f" % rr3)
rr4 = goodness_of_fit(yvals4, normalized_ydata)
print("四阶曲线拟合优度为%.5f" % rr4)
rr5 = goodness_of_fit(yvals5, normalized_ydata)
print("五阶曲线拟合优度为%.5f" % rr5)


p = round((25 - b1) / a1, 2)
#p = b/(math.log(9/a))
# p = round(p, 2)
# 显示坐标点


# 绘制拟合曲线用
x1 = np.arange(5000, 16000, 10)
x1_d = np.array(data_scale(x1)[0])

# y1 = a4 * np.power(x1, 4)
y1_d = a4 * np.power(x1_d, 4) + b4 * np.power(x1_d, 3) + c4 * np.power(x1_d, 2) + d4 * x1_d + e4
y1 = un_data_scale(y1_d, y_scale)

# x0 = np.array(x)
# y0 = np.array(y)
# 计算r2
# y2 = a4 * np.power(x0, 4) + b4 * np.power(x0, 3) + c4 * np.power(x0, 2) + d4 * x0 + e4
# y2 = a * log(x0) + b
# r2 = r2_score(y0, y2)

# figure3 = plt.figure(figsize=(8,6))

plt.ylim((min(y)-min(y)*0.1, max(y)+max(y)*0.1))
plt.xlim((min(x)-min(x)*0.1, max(x)+max(x)*0.1))
# plt.plot(x_arr, yvals1, color="#72CD28", label='一阶拟合曲线')
# plt.plot(x_arr, yvals2, color="#EBBD43", label='二阶拟合曲线')
# plt.plot(x_arr, yvals3, color="#50BFFB", label='三阶拟合曲线')
# plt.plot(x_arr, yvals4, color="gold", label='四阶拟合曲线')
# plt.plot(x_arr, yvals5, color="red", label='五阶拟合曲线')
plt.plot(x1, y1, "blue")
plt.scatter(x_arr, y_arr, color='black', marker="X", label='原始数据')
# plt.scatter(p , 25,s=20,marker='x')
plt.xlabel('x')
plt.ylabel('y')
plt.legend(loc=4)    # 指定legend的位置右下角
plt.title('curve_fit 1~5阶拟合曲线')
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
plt.show()