import PySimpleGUI as sg  # GUI
import datetime  # 时间
import xlsxwriter  # 写xlsx
import xlrd  # 读xlsx
import re  # 正则表达式
import sqlite3  # python内置数据库
import inventoryCheck

# 额外功能：对数据库进行增删改查，重置数据库，备份数据库（未完成）
from PySimpleGUI import SYSTEM_TRAY_MESSAGE_ICON_WARNING


def contrastDB(drug_name, drug_manufactor):
    # 连接db文件
    conn = sqlite3.connect('recourse/database/drug_classification.db')
    cur = conn.cursor()
    # 定义默认的药品类型
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
    # 药名只符合一种类型
    if len(match_table) == 1:
        right_table = match_table[0]
    # 药名符合多种类型
    if len(match_table) > 1:
        # 对比类型中的厂家前两个字
        for match in match_table:
            manufactor_list = cur.execute(
                "select `生产厂家` from `" + match + "` where 药品名称='" + drug_name + "'").fetchall()
            if manufactor_list[0][0] != '' and manufactor_list[0][0] is not None:
                if manufactor_list[0][0][0:2] == drug_manufactor[0:2]:
                    right_table = match
    # 关闭db连接
    cur.close()
    conn.close()
    return right_table


def classification(files_path_list):
    # 定义返回的结果列表
    result = []
    # 获取被分类文件所在文件夹地址
    xlsx_path = files_path_list[0][0:-len(files_path_list[0].split('/')[-1])]
    # 定义分类表（基药、4+7基药、4+7非基药、无法识别）文件的保存地址
    save_file_path = xlsx_path + '药品分类.xlsx'
    # 创建工作表
    sava_file_xlsx = xlsxwriter.Workbook(save_file_path)
    # 为工作表添加sheet
    base_drug_worksheet = sava_file_xlsx.add_worksheet("基药")
    base_drug_47_worksheet = sava_file_xlsx.add_worksheet("4+7基药")
    base_drug_47_non_worksheet = sava_file_xlsx.add_worksheet("4+7非基药")
    unrecognized_worksheet = sava_file_xlsx.add_worksheet("无法识别")
    # 定义每个工作表的有效行域
    save_file_effective_area = {'base_drug_worksheet_row': 2,
                                'base_drug_47_worksheet_row': 2,
                                'base_drug_47_non_worksheet_row': 2,
                                'unrecognized_worksheet_row': 2}
    # 定义格式
    # 标题格式
    title_format = sava_file_xlsx.add_format({'font_size': 16,  # 字体大小
                                              'bold': True,  # 是否粗体
                                              'align': 'center',  # 水平居中对齐
                                              'valign': 'vcenter'  # 垂直居中对齐
                                              })
    # 表头格式
    herder_format = sava_file_xlsx.add_format({'bold': True,  # 是否粗体
                                               'align': 'left',  # 水平左对齐
                                               'valign': 'vcenter',  # 垂直居中对齐
                                               'border': 1  # 边框，0:无边框；1:外边框；
                                               })
    # 左对齐格式
    left_format = sava_file_xlsx.add_format({'align': 'left',  # 水平左对齐
                                             'valign': 'vcenter'  # 垂直居中对齐
                                             })
    # 右对齐格式
    right_format = sava_file_xlsx.add_format({'align': 'right',  # 水平右对齐
                                              'valign': 'vcenter'  # 垂直居中对齐
                                              })
    # 边框格式
    frame_format = sava_file_xlsx.add_format({'border': 1  # 边框，0:无边框；1:外边框；
                                              })
    # 合并单元格并写入表标题
    base_drug_worksheet.merge_range('A1:E1', '城关村基药药房发药统计明细', title_format)
    base_drug_47_worksheet.merge_range('A1:E1', '城关村4+7基药药房发药统计明细', title_format)
    base_drug_47_non_worksheet.merge_range('A1:E1', '城关村4+7非基药药房发药统计明细', title_format)
    unrecognized_worksheet.merge_range('A1:E1', '无法识别', title_format)
    # 定义表头
    xlsx_header = ['药品名称', '药品规格', '单位', '发药数', '发药金额']
    # 写入表头
    xlsx_header_col = 0
    for item in xlsx_header:
        base_drug_worksheet.write(1, xlsx_header_col, item, herder_format)
        base_drug_47_worksheet.write(1, xlsx_header_col, item, herder_format)
        base_drug_47_non_worksheet.write(1, xlsx_header_col, item, herder_format)
        unrecognized_worksheet.write(1, xlsx_header_col, item, herder_format)
        xlsx_header_col += 1
    # 读取传入的xlsx文件
    for file_path in files_path_list:
        # 读取传入的xlsx文件
        try:
            workbook = xlrd.open_workbook(filename=file_path)
        except BaseException:
            time = datetime.datetime.now()
            print('!!!!!!!!!!!!!!!!!!!!!!!!!\n'
                  '时间：  ', time)
            print(file_path + '文件错误，或者该文件非xlsx文件')
            print('!!!!!!!!!!!!!!!!!!!!!!!!!')
            continue
        # 获取第一个sheet表格
        table = workbook.sheets()[0]
        # 判别读取的xlsx格式是否正确
        if table.cell_value(
                rowx=0, colx=0) != '药房发药统计明细' or table.cell_value(
            rowx=1, colx=0) != '药品名称' or table.cell_value(
            rowx=1, colx=1) != '药品规格' or table.cell_value(
            rowx=1, colx=2) != '单位' or table.cell_value(
            rowx=1, colx=3) != '发药数' or table.cell_value(
            rowx=1, colx=4) != '发药金额':
            result.append(file_path + '非药房发药统计明细表格格式！转换失败！')
            continue
        # 行列下标从0开始，定义，药名从第三行第一列开始，所以row=2、col=0
        row = 2
        col = 0
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
                # 与数据库进行对比，判别药品类型
                drug_type = contrastDB(drug_name, drug_manufactor)
                # 定义默认的sheet类型
                sheet_type = unrecognized_worksheet
                # 定义默认的sheet的行
                row_name = 'unrecognized_worksheet_row'
                # drug_type改变sheet_type和row_name
                if drug_type == '':
                    sheet_type = unrecognized_worksheet
                    row_name = 'unrecognized_worksheet_row'
                if drug_type == 'base_drug':
                    sheet_type = base_drug_worksheet
                    row_name = 'base_drug_worksheet_row'
                if drug_type == '4+7_base_drug':
                    sheet_type = base_drug_47_worksheet
                    row_name = 'base_drug_47_worksheet_row'
                if drug_type == '4+7_non_base_drug':
                    sheet_type = base_drug_47_non_worksheet
                    row_name = 'base_drug_47_non_worksheet_row'
                # 分类好后写入新的xlsx文件
                for record_col in range(0, 5):
                    # 读取要分类的一行记录
                    row_value = table.cell_value(rowx=row, colx=record_col)
                    # “发药金额”列需要转为数字最后才能进行求和操作
                    if record_col == 4:
                        row_value = float(row_value)
                    # 把记录写入生成的文件
                    sheet_type.write(save_file_effective_area[row_name], record_col, row_value, frame_format)
                # 生成文件的对应sheet的有效行+1
                save_file_effective_area[row_name] += 1
            # 行+1，准备读取下一行记录
            row += 1
        # 加入转换成功提示
        result.append(file_path + '文件转换成功！')
    # 定义Excel求和函数语句
    # 为防止求和函数出错，需要判别后赋予正确的求和
    if save_file_effective_area['unrecognized_worksheet_row'] != 2:
        unrecognized_worksheet_sum_string = '=SUM(E3:E' + str(
            save_file_effective_area['unrecognized_worksheet_row']) + ')'
    else:
        unrecognized_worksheet_sum_string = '0'
    if save_file_effective_area['base_drug_worksheet_row'] != 2:
        base_drug_worksheet_sum_string = '=SUM(E3:E' + str(save_file_effective_area['base_drug_worksheet_row']) + ')'
    else:
        base_drug_worksheet_sum_string = '0'
    if save_file_effective_area['base_drug_47_worksheet_row'] != 2:
        base_drug_47_worksheet_sum_string = '=SUM(E3:E' + str(
            save_file_effective_area['base_drug_47_worksheet_row']) + ')'
    else:
        base_drug_47_worksheet_sum_string = '0'
    if save_file_effective_area['base_drug_47_non_worksheet_row'] != 2:
        base_drug_47_non_worksheet_sum_string = '=SUM(E3:E' + str(
            save_file_effective_area['base_drug_47_non_worksheet_row']) + ') '
    else:
        base_drug_47_non_worksheet_sum_string = '0'
    # 将各个sheet的求和写入
    unrecognized_worksheet.write(save_file_effective_area['unrecognized_worksheet_row'], 4,
                                 unrecognized_worksheet_sum_string)
    base_drug_worksheet.write(save_file_effective_area['base_drug_worksheet_row'], 4,
                              base_drug_worksheet_sum_string)
    base_drug_47_worksheet.write(save_file_effective_area['base_drug_47_worksheet_row'], 4,
                                 base_drug_47_worksheet_sum_string)
    base_drug_47_non_worksheet.write(save_file_effective_area['base_drug_47_non_worksheet_row'], 4,
                                     base_drug_47_non_worksheet_sum_string)
    # 调整xlsx格式
    base_drug_worksheet.set_row(0, 20.4)  # 第0行的行高
    base_drug_worksheet.set_row(1, 20)  # 第1行的行高
    base_drug_worksheet.set_column(0, 0, 27, left_format)  # 第0列到第0列的列宽，左对齐
    base_drug_worksheet.set_column(1, 3, 12, left_format)  # 第1列到第3列的列宽，左对齐
    base_drug_worksheet.set_column(4, 4, 12, right_format)  # 第4列到第4列的列宽，右对齐
    base_drug_47_worksheet.set_row(0, 20.4)  # 第0行的行高
    base_drug_47_worksheet.set_row(1, 20)  # 第1行的行高
    base_drug_47_worksheet.set_column(0, 0, 27, left_format)  # 第0列到第0列的列宽，左对齐
    base_drug_47_worksheet.set_column(1, 3, 12, left_format)  # 第1列到第3列的列宽，左对齐
    base_drug_47_worksheet.set_column(4, 4, 12, right_format)  # 第4列到第4列的列宽，右对齐
    base_drug_47_non_worksheet.set_row(0, 20.4)  # 第0行的行高
    base_drug_47_non_worksheet.set_row(1, 20)  # 第1行的行高
    base_drug_47_non_worksheet.set_column(0, 0, 27, left_format)  # 第0列到第0列的列宽，左对齐
    base_drug_47_non_worksheet.set_column(1, 3, 12, left_format)  # 第1列到第3列的列宽，左对齐
    base_drug_47_non_worksheet.set_column(4, 4, 12, right_format)  # 第4列到第4列的列宽，右对齐
    unrecognized_worksheet.set_row(0, 20.4)  # 第0行的行高
    unrecognized_worksheet.set_row(1, 20)  # 第1行的行高
    unrecognized_worksheet.set_column(0, 0, 27, left_format)  # 第0列到第0列的列宽，左对齐
    unrecognized_worksheet.set_column(1, 3, 12, left_format)  # 第1列到第3列的列宽，左对齐
    unrecognized_worksheet.set_column(4, 4, 12, right_format)  # 第4列到第4列的列宽，右对齐
    # 保存生成的分类xlsx文件
    sava_file_xlsx.close()
    # 加入生成文件的地址
    result.append('生成的文件地址：' + save_file_path)
    # 返回生成结果列表
    return result


# GUI
def main():
    sg.theme('DarkPurple1')
    layout = [
        [sg.Text('药房发药统计——药品分类器', font=('得意黑', 14)),
         sg.Button(key='setup',
                   image_filename=r"recourse/icon/数据库设置.png",
                   button_color=sg.theme_background_color(),
                   pad=((590, 0), (0, 0)),
                   border_width=0)],
        [sg.Text('>请先选择文件<', key='filenames', size=(100, 4), font=('得意黑', 12), text_color='white')],
        [sg.Output(size=(100, 12), font=('得意黑', 12), key='output')],
        [sg.FilesBrowse('选择文件', key='files', target='files', enable_events=True,
                        file_types=(("ALL xlsx Files", "*.xlsx"),)),
         sg.Button('开始分类'),
         sg.Button('退出'),
         sg.Text('———————————————————   By 张一丁   Version : 1.2        2023年01月28日',
                 font=('得意黑', 11)),
         sg.Button('使用须知')]]
    window = sg.Window('药房发药统计——药品分类器',
                       layout,
                       font=('得意黑', 17),
                       default_element_size=(50, 1),
                       # no_titlebar=True,  # 去除顶部状态栏
                       grab_anywhere=True  # 允许随意拖动窗口
                       )
    while True:
        event, values = window.read()
        if event == 'files':
            files_path = values['files']
            files_path = files_path.replace(';', '\n')
            window['filenames'].update(files_path)
        if event == '开始分类':
            window.FindElement('output').Update('')  # 清空输出框
            time = datetime.datetime.now()
            # 以”;“为分隔符将选择的多文件路径分割字符串为列表
            files_path_list = values['files'].split(';')
            if files_path_list[0] != '':
                # 转换结果
                conversion_results = classification(files_path_list)
                print('*************************')
                print('时间：  ', time)
                for result in conversion_results:
                    print(result)
                print('*************************')

            else:
                print('!!!!!!!!!!!!!!!!!!!!!!!!!\n'
                      '时间：  ', time, '\n'
                                     '未选取文件，请先选择正确的文件\n'
                                     '!!!!!!!!!!!!!!!!!!!!!!!!!')
        if event == 'setup':
            # window.hide()  # 隐藏窗口
            # window.UnHide()  # 取消隐藏窗口
            op = sg.popup_yes_no('确认要进行数据库数据核对吗？', font=('得意黑', 14), no_titlebar=True)
            if op == 'Yes':
                # 窗口显示文本框和浏览按钮, 以便选择文件
                drug_stock_path = sg.popup_get_file("请选择药房库存文件", multiple_files=False,
                                                    font=('得意黑', 14), file_types=(("ALL xlsx Files", "*.xlsx"),), )
                if drug_stock_path != '' and drug_stock_path is not None:
                    window.FindElement('output').Update('')  # 清空输出框
                    if inventoryCheck.check(drug_stock_path):
                        sg.popup_notify("数据库核对信息已下发！", icon=r'recourse/icon/正确.png', location=(600, 500))
                    else:
                        sg.popup_notify("数据库核对失败！", icon=r'recourse/icon/错误.png', location=(600, 500))
        if event == '使用须知':
            window.FindElement('output').Update('')  # 清空输出框
            print('#########################')
            print('--使用须知：')
            print('第一步：点击“选择文件”\n',
                  '第二步：选中需要分类的药房发药统计表\n',
                  '第三步：点击右下角“打开”\n',
                  '第四步：点击“开始分类”\n',
                  '系统提示分类是否成功。')
            print('#########################')
        if event in (None, '退出'):
            break
    window.close()


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    main()
