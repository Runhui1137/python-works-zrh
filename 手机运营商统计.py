'''
    作者: 张润晖
    2021年4月29日 13:44:54 完成
    开源地址: 
    本程序基于pyechart v1开发, 0.x版本的pyecharts无法运行该程序

'''

import csv
from pyecharts.charts import Pie
from pyecharts import options as opts
# 从百度搜到的运营商号段字符串
yd_string = '''134、135、136、137、138、139、147、148、150、151、152、157、158、159、178、182、183、184、187、188、198、1440、1703、1705、1706'''
lt_string = '''130、131、132、155、156、185、186、145、146、166、167、175、176、1704、1707、1708、1709、1710、1711、1712、1713、1714、1715、1716、1717、1718、1719'''
dx_string = '''133、153、177、180、181、189、191、199、1349、1410、1700、1701、1702、1740'''

# 将字符串进行拆分, 将号段存入数组
set_yd = set()
set_lt = set()
set_dx = set()

# 定义一个forEach函数, 方便遍历数据使用
def forEach(iterator, function):
    for e in iterator:
        function(e)
# 提取百度得到的字符串中的每个号段, 并添加到三大运营商对应的集合set中
forEach(yd_string.split(sep="、"), lambda e : set_yd.add(e))
forEach(lt_string.split(sep="、"), lambda e : set_lt.add(e))
forEach(dx_string.split(sep="、"), lambda e : set_dx.add(e))

#读取csv文件, 并将号码读取到tel_nums列表中
tel_nums = []
with open(file='./软件18学生详细名单.csv', encoding="UTF-8") as file :
    items = csv.reader(file)
    # 将每一行中的手机号提取到tel_nums列表中
    for e in enumerate(items):
        if e[0] >= 1 :
            tel_nums.append(e[1][10])

# 遍历列表, 查看号码的前三位和前四位是否与号段集合存在"存在关系", 并使用cnt_xxx变量来记录每个运营商的使用人数
cnt_yd = cnt_lt = cnt_dx = cnt_other = 0
for tel_num in tel_nums:
    three = "{0}".format(tel_num[0:3])
    four = "{0}".format(tel_num[0:4])
    if three in set_yd or four in set_yd:
        cnt_yd += 1
    elif three in set_lt or four in set_lt:
        cnt_lt += 1
    elif three in set_dx or four in set_dx:
        cnt_dx += 1
    else:
        cnt_other += 1;
print("移动:{0} \n联通: {1}\n电信: {2} \n其它/虚拟号段:{3} ".format(cnt_yd, cnt_lt, cnt_dx, cnt_other))

# 绘图
print("绘图ing...")
pie = Pie().add(
    series_name="软件工程2018级手机运营商统计",
    data_pair=[('移动', cnt_yd), ('联通', cnt_lt),('电信', cnt_dx),('其它/虚拟号段', cnt_other)],
    label_opts=opts.LabelOpts(is_show=True, formatter="{b}:{c} - {d}%")
).set_global_opts(
    title_opts=opts.TitleOpts(
        title='软件工程2018级手机运营商统计',
        pos_left="center",
    ),
    legend_opts=opts.LegendOpts(is_show=True, pos_bottom=0)
).render("out.html")

print("绘图完成!")
