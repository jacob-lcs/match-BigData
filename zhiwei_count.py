import xlrd
import re


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


if __name__ == "__main__":
    infos = []  # 新建空字典
    for n in range(6):  # 存储sheet
        print("正在存储第", str(n+1), "个sheet...")
        read_excel('./data/前程无忧（总）.xlsx', infos, n)
    print("保存成功！！")
    count = len(infos)  # 总计数
    print("数据总数为：", str(count))
    # ----------------每个城市职位数量---------------------------#
    cs = {}  # 新建空字典
    for key in infos:
        addition = key['addition']
        confirm = 0
        for soft_key in cs:
            if addition in cs:  # 遍历到的地址在字典中
                confirm = 1
            else:  # 遍历到的地址不在字典中
                confirm = 0
        if confirm == 0:
            if '-' in addition:
                matchObj = re.match(r'(.*)-.*', addition, re.M | re.I)  # 正则表达式匹配城市名
                if matchObj.group(1) in cs:
                    cs[matchObj.group(1)] += 1
                else:
                    cs[matchObj.group(1)] = 1
            else:
                cs[addition] = 1
        else:
            cs[addition] += 1
    print(cs)
    total = 0
    for key in cs:  # 所有岗位总计数
        total += cs[key]

    print("职位总数为：", total)
