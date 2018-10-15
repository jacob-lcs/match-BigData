# -*- coding:utf-8 -*-
import urllib.request
import re
import xlwt  # 用来创建excel文档并写入数据

url_pre = 'https://jobs.51job.com/yinxiaoshejishi/p'
yeshu = 5
table_name = "音效设计师.xls"
sheet_name = "音效设计师"


# 获取原码
def get_content(page):
    url = url_pre + str(page)
    a = urllib.request.urlopen(url)  # 打开网址
    html = a.read().decode('gbk', 'ignore')  # 读取源代码并转为unicode
    return html


def get(html):
    reg = re.compile(
        r'class="e ">.*?<p class="info">.*?<span class="title">.*?<a title="(.*?)" target="_blank".*?<a title="(.*?)".*?<span class="location name">(.*?)</span>.*?<span class="location">(.*?)</span>.*?<span class="time">(.*?)</span>\r\n\t\t\t\t\t\t\t\t</p>\r\n\t\t\t\t\t\t\t\t<p class="order">\r\n\t\t\t\t\t\t\t\t\t学历要求：(.*?)<span>.*?</span>工作经验：(.*?)<span>.*?</span>公司性质：(.*?)<span>.*?</span>公司规模：(.*?)\t\t\t\t\t\t\t\t</p>\r\n\t\t\t\t\t\t\t\t<p class="text" title="(.*?)">',
        re.S)  # 匹配换行符
    items = re.findall(reg, html)
    return items


def excel_write(items, index):
    # 爬取到的内容写入excel表格
    for item in items:  # 职位信息
        for i in range(0, 10):
            # print item[i]
            ws.write(index, i, item[i])  # 行，列，数据
        index += 1


newTable = table_name  # 表格名称
wb = xlwt.Workbook(encoding='utf-8')  # 创建excel文件，声明编码
ws = wb.add_sheet(sheet_name)  # 创建表格
headData = ['招聘职位', '公司', '地址', '薪资', '日期', '学历要求', '工作经验', '公司性质', '公司规模', '职位说明']  # 表头部信息
for colnum in range(0, 10):
    ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))  # 行，列

sum = yeshu + 1
for each in range(sum):
    print("正在保存第" + str(each) + "页的数据...", "已完成", int(each/sum * 100), "%")
    index = (each - 1) * 10 + 1
    excel_write(get(get_content(each)), index)

print("============SUCCEED================")
wb.save(newTable)
