import xlrd
import re
import xlwt


def read_excel(file_excel, infos, n):  # 读excel并将需要的数据分类放在数组里
    info_file = xlrd.open_workbook(file_excel)
    info_sheet = info_file.sheets()[n]  # 获取第一个sheet
    row_count = info_sheet.nrows  # 获取数据的行
    for row in range(1, row_count):
        infos.append(
            {
                'position': info_sheet.cell(row, 0).value,  # 获取第一列的值
                'company': info_sheet.cell(row, 1).value,
                'addition': info_sheet.cell(row, 2).value,
                'salary': info_sheet.cell(row, 3).value,
                'Education': info_sheet.cell(row, 5).value,
                'work-experience': info_sheet.cell(row, 6).value,
                'Company-Type': info_sheet.cell(row, 7).value,
                'Company-Size': info_sheet.cell(row, 8).value,
                'Job-description': info_sheet.cell(row, 9).value
            }
        )
    return infos


def excel_write(items):
    # 爬取到的内容写入excel表格
    len = 1
    for item in items:  # 职位信息
        ws.write(len, 0, item['addition'])  # 行，列，数据
        ws.write(len, 1, item['少于50人'])
        ws.write(len, 2, item['50-150人'])
        ws.write(len, 3, item['150-500人'])
        ws.write(len, 4, item['500-1000人'])
        ws.write(len, 5, item['1000-5000人'])
        ws.write(len, 6, item['5000-10000人'])
        ws.write(len, 7, item['10000人以上'])
        len += 1


if __name__ == "__main__":
    infos = []  # 新建空字典
    for n in range(6):  # 存储sheet
        print("正在存储第", str(n + 1), "个sheet...")
        read_excel('./data/前程无忧（总）.xlsx', infos, n)
    print("保存成功！！")
    count = len(infos)  # 总计数
    print("数据总数为：", str(count))

    cs2 = []  # 新建空字典
    total = 0
    for key in infos:
        addition = key['addition']
        confirm = 0
        for lcs in cs2:
            if lcs['addition'] == addition:  # 遍历到的地址在字典中
                confirm = 1
                break
        if confirm == 0:
            c = 0
            if '-' in addition:
                matchObj = re.match(r'(.*)-.*', addition, re.M | re.I)  # 正则表达式匹配城市名
                if matchObj:
                    addition = matchObj.group(1)
                c = 0
            for lcs in cs2:
                if lcs['addition'] == addition:  # 遍历到的地址在字典中
                    c = 1
                    break
            if c == 0:
                addition2 = {'addition': addition, '少于50人': 0, '50-150人': 0, '150-500人': 0, '500-1000人': 0, '1000-5000人': 0, '5000-10000人': 0, '10000人以上': 0}
                cs2.append(addition2)
        for ad in cs2:
            if ad["addition"] == addition:
                if key['Company-Size'] == '少于50人':
                    ad['少于50人'] += 1
                elif key['Company-Size'] == '50-150人':
                    ad['50-150人'] += 1
                elif key['Company-Size'] == '150-500人':
                    ad['150-500人'] += 1
                elif key['Company-Size'] == '500-1000人':
                    ad['500-1000人'] += 1
                elif key['Company-Size'] == '1000-5000人':
                    ad['1000-5000人'] += 1
                elif key['Company-Size'] == '5000-10000人':
                    ad['5000-10000人'] += 1
                elif key['Company-Size'] == '10000人以上':
                    ad['10000人以上'] += 1
                elif key['Company-Size'] == '':
                    ad['少于50人'] += 1
    print(cs2)
    newTable = "各城市公司规模统计.xls"  # 表格名称
    wb = xlwt.Workbook(encoding='utf-8')  # 创建excel文件，声明编码
    ws = wb.add_sheet('各城市公司规模统计')  # 创建表格
    headData = ['地点', '少于50人', '50-150人', '150-500人', '500-1000人', '1000-5000人', '5000-10000人', '10000人以上']  # 表头部信息
    for colnum in range(0, 8):
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))  # 行，列
    excel_write(cs2)
    wb.save(newTable)
