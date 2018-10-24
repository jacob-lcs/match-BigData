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
        ws.write(len, 1, item['本科'])
        ws.write(len, 2, item['博士'])
        ws.write(len, 3, item['初中及以下'])
        ws.write(len, 4, item['大专'])
        ws.write(len, 5, item['高中'])
        ws.write(len, 6, item['硕士'])
        ws.write(len, 7, item['中技'])
        ws.write(len, 8, item['中专'])
        ws.write(len, 9, item['无要求'])
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
                addition2 = {'addition': addition, '本科': 0, '博士': 0, '初中及以下': 0, '大专': 0, '高中': 0, '硕士': 0, '中专': 0,
                             '中技': 0, '无要求': 0}
                cs2.append(addition2)
        for ad in cs2:
            if ad["addition"] == addition:
                if key['Education'] == '本科':
                    ad['本科'] += 1
                elif key['Education'] == '博士':
                    ad['博士'] += 1
                elif key['Education'] == '初中及以下':
                    ad['初中及以下'] += 1
                elif key['Education'] == '大专':
                    ad['大专'] += 1
                elif key['Education'] == '高中':
                    ad['高中'] += 1
                elif key['Education'] == '硕士':
                    ad['硕士'] += 1
                elif key['Education'] == '中技':
                    ad['中技'] += 1
                elif key['Education'] == '中专':
                    ad['中专'] += 1
                elif key['Education'] == '':
                    ad['无要求'] += 1
    print(cs2)
    newTable = "各城市学历要求.xls"  # 表格名称
    wb = xlwt.Workbook(encoding='utf-8')  # 创建excel文件，声明编码
    ws = wb.add_sheet('学历要求')  # 创建表格
    headData = ['地点', '本科', '博士', '初中及以下', '大专', '高中', '硕士', '中技', '中专', '无要求']  # 表头部信息
    for colnum in range(0, 10):
        ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))  # 行，列
    excel_write(cs2)
    wb.save(newTable)
