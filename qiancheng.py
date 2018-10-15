# -*- coding:utf-8 -*-
import urllib.request
import re
from xlrd import open_workbook
from xlutils.copy import copy


# 获取原码
def get_content(page):
    url = 'https://jobs.51job.com/gaojiruanjian/p' + str(page)
    a = urllib.request.urlopen(url)  # 打开网址
    html = a.read().decode('gbk')  # 读取源代码并转为unicode
    return html


def get(html):

    reg = re.compile(
        r'class="e ">.*?<p class="info">.*?<span class="title">.*?<a title="(.*?)" target="_blank".*?<a title="(.*?)".*?<span class="location name">(.*?)</span>.*?<span class="location">(.*?)</span>.*?<span class="time">(.*?)</span>\r\n\t\t\t\t\t\t\t\t</p>\r\n\t\t\t\t\t\t\t\t<p class="order">\r\n\t\t\t\t\t\t\t\t\t学历要求：(.*?)<span>.*?</span>工作经验：(.*?)<span>.*?</span>公司性质：(.*?)<span>.*?</span>公司规模：(.*?)\t\t\t\t\t\t\t\t</p>\r\n\t\t\t\t\t\t\t\t<p class="text" title="(.*?)">',
        re.S)  # 匹配换行符
    items = re.findall(reg, html)
    # print(items)
    return items


def excel_write(items, index):
    # 爬取到的内容写入excel表格
    rb = open_workbook('高级软件工程师.xls')  # 打开表格
    wb = copy(rb)
    ws = wb.get_sheet(0)
    for n in range(1, 11):
        for i in range(1, 11):
            nn = n + index * 10
            ii = i - 1
            ws.write(nn, ii, items[n-1][ii])
    wb.save('计算机软件.xls')


for n in range(2953):
    print("正在保存第", str(n), "页的数据.....")
    excel_write(get(get_content(page=n)), n-1)
