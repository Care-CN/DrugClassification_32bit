import sqlite3
import xlrd
# 正则表达式
import re

import xlsxwriter


def DB(drug_name, drug_manufactor):
    conn = sqlite3.connect('recourse/database/drug_classification.db')
    cur = conn.cursor()
    right_table = ''
    # 定义查询出的匹配到的表名
    match_table = []
    drug_name_select_base_drug = cur.execute("select * from `base_drug` where 药品名称='" + drug_name + "'")
    if len(drug_name_select_base_drug.fetchall()) != 0:
        match_table.append('base_drug')
    drug_name_select_47_base_drug = cur.execute("select * from `4+7_base_drug` where 药品名称='" + drug_name + "'")
    if len(drug_name_select_47_base_drug.fetchall()) != 0:
        match_table.append('4+7_base_drug')
    drug_name_select_47_non_base_drug = cur.execute("select * from `4+7_non_base_drug` where 药品名称='" + drug_name + "'")
    if len(drug_name_select_47_non_base_drug.fetchall()) != 0:
        match_table.append('4+7_non_base_drug')
    if len(match_table) == 1:
        right_table = match_table[0]
    if len(match_table) > 1:
        for match in match_table:
            manufactor_list = cur.execute(
                "select `生产厂家` from `" + match + "` where 药品名称='" + drug_name + "'").fetchall()
            print('表：',match,'            厂家：',manufactor_list[0][0])
            if manufactor_list[0][0] != '' and manufactor_list[0][0] is not None:
                if manufactor_list[0][0][0:2] == drug_manufactor[0:2]:
                    right_table = match
        print(match_table)
    cur.close()
    conn.close()
    return right_table
    # print('---end---')


def testxlrd():
    files_path_list = [
        # r'D:\PyCharm\PJ\DrugClassification_32bit\基药非基药分类\20221231192151.xlsx',
        # r'D:\PyCharm\PJ\DrugClassification_32bit\基药非基药分类\20221231192212.xlsx',
        # r'D:\PyCharm\PJ\DrugClassification_32bit\基药非基药分类\20221231195537.xlsx',
        r"C:\Users\DELL\Desktop\基药非基药分类\库存管理-包含0库存.xlsx"
    ]
    files_path_list2 = [r"F:\新建文件夹\20230115191743.xlsx",
                        r"F:\新建文件夹\20230115191806.xlsx",
                        r"F:\新建文件夹\20230115191824.xlsx"]

    for file_path in files_path_list:
        # 读取传入的xlsx文件
        workbook = xlrd.open_workbook(filename=file_path)
        # 获取第二个sheet表格
        table = workbook.sheets()[1]

        # 行列下标从0开始，定义，药名从第三行第一列开始，所以row=2、col=0
        row = 2
        col = 1
        # 获取药品名单元格内的值
        while row < table.nrows:  # table.nrows是表格的有效行数
            if table.cell_value(rowx=row, colx=col) != '':
                drug_name = table.cell_value(rowx=row, colx=col)
                # 药名规范化：去掉".",厂家只对比前两个字
                drug_name = drug_name.replace('.', '')
                # 去掉" "
                drug_name = drug_name.replace(' ', '')
                # 药名最后一个字是数字的去掉数字
                while drug_name[-1] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                    drug_name = drug_name.replace(drug_name[-1], '')
                # 正则表达式 去掉括号及括号中内容
                drug_name = re.sub('\\(.*?\\)', '', drug_name)
                drug_name = re.sub('（.*?）', '', drug_name)
                drug_name = re.sub('\\(.*?）', '', drug_name)
                drug_name = re.sub('（.*?\\)', '', drug_name)
                # 药品厂家
                drug_manufactor = table.cell_value(rowx=row, colx=4)
                drug_type = DB(drug_name, drug_manufactor)
                if drug_type == '':
                    print('药品编号',table.cell_value(rowx=row, colx=0))
                    print('药品名：', drug_name)
                    print('药品厂家：', drug_manufactor)
                    print('药品类型：', drug_type)
            print('--------------------------------')
            row += 1
        print(file_path)
    print('------END------')


if __name__ == '__main__':
    # testDB()
    testxlrd()
