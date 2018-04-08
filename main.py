# -*- coding:utf-8 -*-

import xlrd


# 打开Excel文件
def open_excel(file_path):
    try:
        data = xlrd.open_workbook(file_path)
        return data
    except Exception as e:
        print("文件打开失败!"+str(e))

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     by_index：通过索引顺序获取，默认第一个索引
def excel_table_byindex(file, by_index = 0):
    data = open_excel(file)
    table = data.sheet_by_index(by_index)  # 根据索引来获取excel中的sheet
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    list =[]  # 装读取结果的序列
    for rownum in range(0, nrows):  # 遍历每一行的内容
         row = table.row_values(rownum)  # 根据行号获取行
         if row:
             app = [] # 某一行
             for i in range(0, ncols):  # 一列列地读取行的内容
                app.append(row[i])
             list.append(app)  # 装载数据
    return list


def main():
    excel_1 = excel_table_byindex("rawData/吉力沈辽路站OS管控流水20180402.xlsx", 0)
    excel_2 = excel_table_byindex("rawData/吉力沈辽路站正星管控流水20180402.XLS", 0)
    print(len(excel_1))

    print(excel_2)



if __name__ == "__main__":
    main()