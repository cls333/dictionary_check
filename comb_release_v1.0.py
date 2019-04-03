# -*- coding:utf-8 -*-
import xlrd
import xlwt
import openpyxl
import sqlite3
import re
import sys
import os

input_file_name = 'NDS1560.xlsx'


# input_file_name = '数据字典_V0.1.xlsx'


def get_sheet_data_by_name(file_name, sheet_name):
    book = xlrd.open_workbook(file_name)
    if sheet_name in book.sheet_names():
        return book.sheet_by_name(sheet_name)
    else:
        return None


def create_table_sql(type):
    table_info = '''
    create table table_info(
      sys_name varchar (32),
      sys_code varchar (32),
      sys_module varchar (32),
      schema varchar (32),
      tab_prefix varchar (32),
      tab_enname varchar (64),
      tab_zhname varchar (1024),
      tab_desc   varchar (1024),
      tab_type   varchar (32),
      exists_pkey varchar (1),
      importance_degree varchar (8),
      primary_key varchar (1024),
      public_key  varchar (1024),
      last_date   integer
    )
    '''

    column_info = '''
    create table column_info(
      sys_name varchar (32),
      sys_code varchar (32),
      sys_module varchar (32),
      schema varchar (32),
      tab_prefix varchar (32),
      tab_enname varchar (64),
      tab_zhname varchar (1024),
      col_no   integer,
      col_enname varchar (32),
      col_zhname varchar (32),
      col_type varchar (8),
      is_primary varchar (1024),
      allowed_null  varchar (1024),
      is_code   varchar (32),
      cite_code varchar (32)
    )
    '''
    code_info = '''
    create table code_info(
      sys_name varchar (32),
      sys_code varchar (32),
      sys_module varchar (32),
      schema varchar (32),
      tab_prefix varchar (32),
      tab_enname varchar (64),
      tab_zhname varchar (1024),
      col_enname varchar (32),
      col_zhname varchar (32),
      col_type varchar (8),
      value varchar (32),
      desc varchar (1024)
    )
    '''
    if type == 'table':
        return table_info
    elif type == 'column':
        return column_info
    elif type == 'code':
        return code_info


def r01_tab(c):
    """表级信息存在空值"""
    __rule_name__ = "表级信息存在空值"
    _rule_name = "表级信息存在空值"
    c.execute('select * '
              'from table_info '
              'where sys_name == \'\' '
              'or sys_code == \'\' '
              'or sys_module == \'\' '
              'or schema == \'\' '
              'or tab_prefix == \'\' '
              'or tab_zhname == \'\' '
              'or exists_pkey == \'\' '
              'or (exists_pkey == \'Y\' '
              'and primary_key == \'\')')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r02_col(c):
    """字段级信息存在空值"""
    _rule_name = "字段级信息存在空值"
    c.execute('select * '
              'from column_info '
              'where tab_enname == \'\' '
              'or col_no == \'\' '
              'or col_enname == \'\' '
              'or col_zhname == \'\' '
              'or col_type == \'\' '
              'or is_primary == \'\' '
              'or allowed_null == \'\' '
              'or is_code == \'\' '
              'or cite_code == \'\'')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r03_cod(c):
    """代码级信息存在空值"""
    _rule_name = "代码级信息存在空值"
    c.execute('select * '
              'from code_info '
              'where tab_enname == \'\' '
              'or col_enname == \'\' '
              'or value == \'\' '
              'or desc == \'\'')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r04_tab(c):
    """表级信息表中文名存在非法字符"""
    _rule_name = "表级信息表中文名存在非法字符"
    c.execute('select * from table_info')
    for res in c.fetchall():
        star = res[6]
        rule = re.compile("[^a-zA-Z0-9\u4e00-\u9fa5]")
        if star != rule.sub('', star) or ' ' in star:
            yield sys._getframe().f_code.co_name, _rule_name, res


def r05_col(c):
    """字段级信息字段中文名存在非法字符"""
    _rule_name = "字段级信息字段中文名存在非法字符"
    c.execute('select * from column_info')
    for res in c.fetchall():
        star = res[9]
        rule = re.compile("[^a-zA-Z0-9\u4e00-\u9fa5]")
        if star != rule.sub('', star) or ' ' in star:
            yield sys._getframe().f_code.co_name, _rule_name, res


def r06_cod(c):
    """代码级信息代码值存在非法字符"""
    _rule_name = "代码级信息代码值存在非法字符"
    c.execute('select * from code_info')
    for res in c.fetchall():
        # star1 = res[11]
        star2 = res[10]
        rule = re.compile("[^a-zA-Z0-9\u4e00-\u9fa5]")
        # if star1 != rule.sub('', star1) or star2 != rule.sub('', star2):
        if star2 != rule.sub('', star2):
            yield sys._getframe().f_code.co_name, _rule_name, res


def r07_tab(c):
    """表级信息存在重复表"""
    _rule_name = "表级信息存在重复表"
    c.execute('select * '
              'from table_info '
              'where tab_enname in '
              '(select tab_enname '
              'from table_info '
              'group by tab_enname '
              'having count(*) >= 2) '
              'order by tab_enname')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r08_col(c):
    """字段级信息存在重复字段"""
    _rule_name = "字段级信息存在重复字段"
    c.execute('select * '
              'from column_info '
              'where (tab_enname||col_enname) '
              'in (select tab_enname||col_enname '
              'from column_info '
              'group by tab_enname,col_enname '
              'having count(*) >= 2 )'
              'order by tab_enname, col_no')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r09_cod(c):
    """代码级信息存在重复代码值"""
    _rule_name = "代码级信息存在重复代码值"
    c.execute('select * '
              'from code_info '
              'where (tab_enname||col_enname||value) '
              'in (select tab_enname||col_enname||value '
              'from code_info '
              'group by tab_enname,col_enname,value '
              'having count(*) > 1) '
              'order by tab_enname, col_enname, value')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r10_tab(c):
    """表级信息中表的主键字段在字段级信息中不存在"""
    _rule_name = "表级信息中表的主键字段在字段级信息中不存在"
    c.execute('select tab_enname,primary_key||\';\'||public_key, * '
              'from table_info '
              'where tab_enname != \'\'')
    temp = []
    for res in c.fetchall():
        # 循环之前清空列表
        del temp[:]
        # 索引为空时直接跳过
        if res[1] != ';':
            for j in re.findall(r'[(](.*?)[)]', res[1]):
                temp.extend(j.split(','))
            index_keys = tuple(set(temp))
            tab_enname = res[0]
            for k in index_keys:
                c.execute('select count(*) from column_info where tab_enname == \'{0}\' '
                          'and col_enname == \'{1}\''.format(tab_enname, k))
                keys_count = c.fetchall()[0][0]

                if keys_count == 0:
                    yield sys._getframe().f_code.co_name, _rule_name, res[2:]


def r11_col(c):
    """字段级信息中字段长度为1未标记为代码字段"""
    _rule_name = '字段级信息中字段长度为1未标记为代码字段'
    c.execute('select col_type, * from column_info '
              'where is_code == \'N\' '
              'and tab_enname != \'\' '
              'and col_enname != \'\'')
    for res in c.fetchall():
        # if res[0].upper()[-7:] in ('CHER(1)', 'CTER(1)'):
        if res[0].upper()[-3:] == '(1)':
            yield sys._getframe().f_code.co_name, _rule_name, res[1:]


def r12_col(c):
    """代码级信息中码值字段只有一个码值"""
    _rule_name = '代码级信息中码值字段只有一个码值'
    c.execute('select tab_enname, col_enname, col_type, * '
              'from column_info '
              'where is_code == \'Y\' '
              'and tab_enname != \'\' '
              'and col_enname != \'\'')

    for res in c.fetchall():
        c.execute('select count(*) '
                  'from code_info '
                  'where tab_enname == \'{0}\' '
                  'and col_enname == \'{1}\''.format(res[0], res[1]))
        for j in c.fetchall():
            if j[0] == 1:
                yield sys._getframe().f_code.co_name, _rule_name, res[1:]


def r13_col(c):
    """字段级信息中非代码字段存在码值"""
    _rule_name = '字段级信息中非代码字段存在码值'
    c.execute('select tab_enname, col_enname, col_type, * '
              'from column_info '
              'where is_code == \'N\' '
              'and tab_enname != \'\' '
              'and col_enname != \'\'')
    for res in c.fetchall():
        c.execute('select count(*) '
                  'from code_info '
                  'where tab_enname == \'{0}\' '
                  'and col_enname == \'{1}\''.format(res[0], res[1]))
        for j in c.fetchall():
            if j[0] > 0:
                yield sys._getframe().f_code.co_name, _rule_name, res[3:]


def r14_tab(c):
    """表级信息中存在的表没有字段信息"""
    _rule_name = "表级信息中存在的表没有字段信息"
    c.execute('select tab_enname, * '
              'from table_info '
              'where tab_enname != \'\'')
    for res in c.fetchall():
        c.execute('select count(*) '
                  'from column_info '
                  'where tab_enname == \'{0}\''.format(res[0]))
        col_count = c.fetchall()[0][0]
        if col_count == 0:
            yield sys._getframe().f_code.co_name, _rule_name, res[1:]


def r15_col(c):
    """字段级信息中存在的表没有表信息"""
    _rule_name = "字段级信息中存在的表没有表信息"
    c.execute('select tab_enname, * '
              'from column_info '
              'where tab_enname != \'\' '
              'and col_enname != \'\'')
    for res in c.fetchall():
        c.execute('select count(*) '
                  'from table_info '
                  'where tab_enname == \'{0}\''.format(res[0]))
        col_count = c.fetchall()[0][0]
        if col_count == 0:
            yield sys._getframe().f_code.co_name, _rule_name, res[1:]


def r16_cod(c):
    """代码级信息中代码值和注释完全一致"""
    _rule_name = "代码级信息中代码值和注释完全一致"
    c.execute('select * '
              'from code_info '
              'where value == desc '
              'and value != \'\'')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r17_tab(c):
    """表级信息中表英文名和中文名一致"""
    _rule_name = "表级信息中表英文名和中文名一致"
    c.execute('select * '
              'from table_info '
              'where tab_enname == tab_zhname '
              'and tab_enname != \'\'')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


def r18_col(c):
    """字段级信息中字段英文名和中文名一致"""
    _rule_name = "字段级信息中字段英文名和中文名一致"
    c.execute('select * '
              'from column_info '
              'where col_enname == col_zhname '
              'and col_enname != \'\' '
              'and tab_enname != \'\'')
    for res in c.fetchall():
        yield sys._getframe().f_code.co_name, _rule_name, res


if __name__ == '__main__':
    # 创建数据库
    con = sqlite3.connect(':memory:')
    # con = sqlite3.connect('comb.db')
    cur = con.cursor()
    # 创建表级信息数据表
    cur.executescript(create_table_sql('table'))
    # 创建字段级信息数据表
    cur.executescript(create_table_sql('column'))
    # 创建代码级信息数据表
    cur.executescript(create_table_sql('code'))

    # 加载表级信息到内存
    sheet = get_sheet_data_by_name(input_file_name, '表级信息')
    if sheet:
        # 跳过表头取表数据
        for r in range(2, sheet.nrows):
            d = sheet.row_values(r)
            cur.execute('insert into table_info values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                        (d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9],
                         d[10], d[11], d[12], d[13]))
    else:
        raise Exception('表级信息表不存在')

    # 加载字段级信息数据到内存
    sheet = get_sheet_data_by_name(input_file_name, '字段级信息')
    if sheet:
        # 跳过表头取表数据
        for r in range(2, sheet.nrows):
            d = sheet.row_values(r)
            cur.execute('insert into column_info values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                        (d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10],
                         d[11], d[12], d[13], d[14]))
    else:
        raise Exception('字段级信息表不存在')

    # 加载代码级信息数据到内存
    sheet = get_sheet_data_by_name(input_file_name, '代码级信息')
    if sheet:
        # 跳过表头取表数据
        for r in range(2, sheet.nrows):
            d = sheet.row_values(r)
            cur.execute('insert into code_info values(?,?,?,?,?,?,?,?,?,?,?,?)',
                        (d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9],
                         d[10], d[11]))
    else:
        raise Exception('代码级信息表不存在')

    con.commit()

    # 选择要执行的检查函数
    func_list = [r01_tab, r02_col, r03_cod, r04_tab, r05_col, r06_cod,
                 r07_tab, r08_col, r09_cod, r10_tab, r11_col, r12_col,
                 r13_col, r14_tab, r15_col, r16_cod, r17_tab, r18_col]
    # func_list = [r01_tab]
    # 执行函数并输出结果
    for func in func_list:
        basename, extension = os.path.splitext(input_file_name)
        func_name = getattr(func, '__name__')
        func_doc = getattr(func, '__doc__')
        output_name = basename + '_' + func_name + '_' + func_doc + '({0}).xls'
        # 创建excel数据表
        # wb = openpyxl.Workbook()
        # if func_name[-3:] == "tab":
        #     sheet = wb.active
        #     sheet.title = '表级信息'
        #     # 添加表头
        #     sheet.append(('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
        #                   '表描述', '表类型', '是否存在主键', '重要程度', '唯一索引', '非唯一索引', '表最新数据日期'))
        #     sheet.append(('', '', '', '', '', '', '', '', '', '', '', '', '', ''))
        # elif func_name[-3:] == "col":
        #     sheet = wb.active
        #     sheet.title = '字段级信息'
        #     sheet.append(('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名',
        #                   '表中文名', '字段序号', '字段英文名', '字段中文名', '字段类型', '是否主键',
        #                   '是否允许为空', '是否代码字段', '是否引用代码表', '字段注释'))
        #     sheet.append(())
        # elif func_name[-3:] == "cod":
        #     sheet = wb.active
        #     sheet.title = '代码级信息'
        #     sheet.append(('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
        #                   '字段英文名', '字段中文名', '字段类型', '代码取值', '代码注释'))
        #     sheet.append(())
        #
        # # 在表格尾部添加检核结果
        # for i in func(cur):
        #     print(list(i))
        #     sheet.append(list(i[2]))
        # wb.save(output_name)

        wb = xlwt.Workbook()
        # 创建表级信息sheet
        if func_name[-3:] == "tab":
            header = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
                      '表描述', '表类型', '是否存在主键', '重要程度', '唯一索引', '非唯一索引', '表最新数据日期')
            sheet = wb.add_sheet('表级信息')
        # 创建字段级信息sheet
        if func_name[-3:] == "col":
            header = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名',
                      '表中文名', '字段序号', '字段英文名', '字段中文名', '字段类型', '是否主键', '是否允许为空',
                      '是否代码字段', '是否引用代码表', '字段注释')
            sheet = wb.add_sheet('字段级信息')
        # 创建代码级信息sheet
        if func_name[-3:] == "cod":
            header = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
                      '字段英文名', '字段中文名', '字段类型', '代码取值', '代码注释')
            sheet = wb.add_sheet('表级信息')

        # 开始写入表头
        # 初始化cell定位
        x = 0
        y = 0
        # Create the pattern
        pattern = xlwt.Pattern()
        # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan,
        # 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal,
        # 22 = Light Gray, 23 = Dark Gray, the list goes on...
        pattern.pattern_fore_colour = 22
        # Create the pattern
        style = xlwt.XFStyle()
        # Add pattern to style
        style.pattern = pattern
        for i in header:
            # 写入字段值
            sheet.write(x, y, i, style)
            # 列递增
            y = y + 1
        # 行递增
        x = x + 1
        # 写入数据
        for i in func(cur):
            # print('  row_value: ', i[2])
            # 重置列定位
            y = 0
            for j in i[2]:
                # 写入字段值
                # print(' cell_value: ', j)
                sheet.write(x, y, j)
                # 列递增
                y = y + 1
            # 行递增
            x = x + 1
        # 保存文件，有检核结果时保存，否则不保存
        if x - 1 > 0:
            output_file_name = output_name.format(x - 1)
            print('save_to_file: ', output_file_name)
            wb.save(output_file_name)
    # 关闭游标
    con.close()
