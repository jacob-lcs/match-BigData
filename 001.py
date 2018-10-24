# -*- coding:utf-8 -*-

import re
from urllib import request
import requests
import urllib.request
import time
import random
from openpyxl import workbook
import os
import xlwt

firms = []
salary = ''

os.chdir('C:\\Users\\Administrator\\Desktop')


def getReview(url):
    proxy = {'http': '61.135.217.7:80'}
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36'
    headers = {'User-Agent': user_agent}
    print("url:", url)
    a = urllib.request.Request(url=url, headers=headers)  # 打开网址
    content = urllib.request.urlopen(a).read()
    # content = a.read().decode('gbk', 'ignore')  # 读取源代码并转为unicode

    # req = requests.get(url)
    # res = request.urlopen(req)

    content = content.decode('utf-8')
    # print("获取到的网址为：", content)
    pattern = re.compile('.*?<a ka="com.*?-review" href="(.*?)" class="weird" target="_blank">.*?</a>.*?<a ka="com.*?-salary".*?class="weird".*?>工资(.*?)</a>', re.S)
    hrefs = re.findall(pattern, content)
    print("获取到的网址为：", hrefs)
    for href in hrefs:
        full_url = 'https://www.kanzhun.com' + href[0]
        print("正在获取", full_url, "的数据...")
        salary = href[1]
        getData(full_url, salary)
        time.sleep(5 + random.random())


def getData(url, salary):
    firm_review = []

    # proxy = {'http': 'http://114.230.41.78:808'}
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36'
    headers = {'User-Agent': user_agent}

    req = requests.get(url=url, headers=headers)

    # 爬公司
    content = req.content.decode('utf-8')
    # print("公司网页：", content)
    pattern1 = re.compile('<strong class="f_24">.*?<a href=".*?" ka="head-title">(.*?)</a>.*?</strong>', re.S)
    firm_review = re.findall(pattern1, content)
    print("公司为：", firm_review)

    # 爬点评总数
    pattern2 = re.compile('<span.*?f_12 grey_99 ml5.*?>(.*?)</span>', re.S)
    p = re.findall(pattern2, content)
    for i in p:
        firm_review.append(i)
        print("点评总数为：", i)

    # 爬各词汇
    pattern3 = re.compile('<a.*?review-label.*?>(.*?)</a>', re.S)
    pre = re.findall(pattern3, content)
    print("爬取到的词汇为：", pre)
    for i in range(len(pre)):
        pre[i] = re.sub(r'\s+', '', pre[i])  # 去掉/n,/t

    if (pre):
        search(u'福利好', pre, firm_review)
        search(u'待遇好', pre, firm_review)
        search(u'待遇不错', pre, firm_review)
        search(u'福利不错', pre, firm_review)
        search(u'发展空间大', pre, firm_review)
        search(u'机会多', pre, firm_review)
        search(u'工资低', pre, firm_review)
        search(u'加班', pre, firm_review)
        search(u'压力大', pre, firm_review)
        search(u'流动性大', pre, firm_review)
        search(u'年轻', pre, firm_review)
        search(u'满意', pre, firm_review)
        search(u'不满意', pre, firm_review)

    if (firm_review):
        print(firm_review)
        if (len(firm_review) == 15):
            firms.append(
                [salary, firm_review[0], firm_review[1], firm_review[2], firm_review[3], firm_review[4], firm_review[5],
                 firm_review[6], firm_review[7], firm_review[8], firm_review[9], firm_review[10], firm_review[11],
                 firm_review[12], firm_review[13], firm_review[14]])
            print("保存成功！")


def search(words, array, final):
    origin = len(final)
    for i in range(len(array)):
        if (array[i].find(words) != -1):
            x = re.findall(r"\d+", array[i])
            for i in x:
                final.append(i)
            break

    if ((len(final) - origin) == 0):  # 如果没有这种评论，数值设为0
        final.append('0')


def excel_write(items):
    # 爬取到的内容写入excel表格
    j = 1
    for item in items:
        for i in range(len(item)):
            ws.write(j, int(i % 16), item[i])
        j = j + 1


if __name__ == '__main__':

    newTable = '看准网.xls'  # 表格名称
    wb = xlwt.Workbook(encoding='utf-8')  # 创建excel文件，声明编码
    ws = wb.add_sheet('看准网')  # 创建表格
    headData = ['工资', '公司', '点评数', '福利好', '待遇好', '待遇不错', '福利不错', '发展空间大', '机会多', '工资低', '加班', '压力大', '流动性大', '年轻', '满意',
                '不满意']  # 表头部信息

    for colnum in range(0, 16):
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))  # 行，列

    print("写入表头成功！")
    all = [
        'https://www.kanzhun.com/plc52p'
        # 'https://www.kanzhun.com/plc4p',
        # 'https://www.kanzhun.com/plc2p',
        # 'https://www.kanzhun.com/plc3p',
        # 'https://www.kanzhun.com/plc121p',
        # 'https://www.kanzhun.com/plc116p',
        # 'https://www.kanzhun.com/plc72p',
        # 'https://www.kanzhun.com/plc69p',
        # 'https://www.kanzhun.com/plc100p',
        # 'https://www.kanzhun.com/plc70p',
        # 'https://www.kanzhun.com/plc105p',
        # 'https://www.kanzhun.com/plc120p',

    ]
    for u in all:
        for i in range(1, 11):
            url = u + str(i) + '.html'
            getReview(url)
            time.sleep(5 + random.random())

    excel_write(firms)

    print("============SUCCEED================")
    wb.save(newTable)