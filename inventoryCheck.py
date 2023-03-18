import sqlite3  # python内置数据库
import xlrd  # 读xlsx
import re  # 正则表达式


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
            # print(drug_name,'-------',drug_manufactor)
            # print('表：', match, '            厂家：', manufactor_list[0][0])
            if manufactor_list[0][0] != '' and manufactor_list[0][0] is not None:
                if manufactor_list[0][0][0:2] == drug_manufactor[0:2]:
                    right_table = match
    #     print(match_table)
    # if right_table == '':
    #     print('未找到正确药品类型！')
    # else:
    #     print('药品类型查找成功！')
    cur.close()
    conn.close()
    return right_table, match_table


def check(drug_stock_path):
    try:
        # 读取传入的xlsx文件
        workbook = xlrd.open_workbook(filename=drug_stock_path)
    except BaseException:
        return False
    # 获取第1个sheet表格
    table = workbook.sheets()[0]
    # 判别读取的xlsx格式是否正确
    if table.cell_value(rowx=0, colx=0) != '药房库存管理':
        return False
    # 行列下标从0开始，定义，药名从第三行第二列开始，所以row=2、col=1
    row = 2
    col = 1
    # 无法识别的药品计数
    count = 0
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
            drug_manufactor = table.cell_value(rowx=row, colx=6)
            drug_type = DB(drug_name, drug_manufactor)
            if drug_type[0] == '':
                # print('药品编号', table.cell_value(rowx=row, colx=0))
                print('药品行号：', row + 1)
                print('药品名：', drug_name)
                print('药品厂家：', drug_manufactor)
                print('药品类型：', drug_type[0])
                print('药品名所在数据库表：', drug_type[1])
                print('--------------------------------')
                count += 1
        row += 1
    print('共', count, '种无法识别的药品')
    print(drug_stock_path)
    return True
