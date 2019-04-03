# -*- coding:utf-8 -*-
import xlrd
import xlsxwriter
import sqlite3
import re
import sys
import os

input_file = 'NDS1560.xlsx'


# input_file = '数据字典_V0.1.xlsx'


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
    """字段级信息标记为码值的字段只有一个码值"""
    _rule_name = '字段级信息标记为码值的字段只有一个码值'
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


def r19_col(c):
    """字段级信息字段中文名相同但是否代码字段标识不同"""
    _rule_name = "字段级信息字段中文名相同但是否代码字段标识不同"
    c.execute('select distinct a.* from column_info a, column_info b '
              'where a.col_zhname == b.col_zhname '
              'and a.is_code != b.is_code '
              'and a.col_zhname != \'\' '
              'and a.is_code != \'\' '
              'and b.is_code != \'\' '
              'order by a.col_zhname, a.tab_enname, a.col_enname')
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
    sheet = get_sheet_data_by_name(input_file, '表级信息')
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
    sheet = get_sheet_data_by_name(input_file, '字段级信息')
    if sheet:
        # 跳过表头取表数据
        for r in range(2, sheet.nrows):
            d = sheet.row_values(r)
            cur.execute('insert into column_info values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                        (d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10],
                         d[11], d[12], str(d[13]), d[14]))
    else:
        raise Exception('字段级信息表不存在')

    # 加载代码级信息数据到内存
    sheet = get_sheet_data_by_name(input_file, '代码级信息')
    if sheet:
        # 跳过表头取表数据
        for i in range(2, sheet.nrows):  # 指定从1开始，到最后一列，跳过表头
            nos = []
            for j in range(sheet.ncols):
                # 判断python读取的返回类型  0 --empty,1 --string, 2 --number(都是浮点), 3 --date, 4 --boolean, 5 --error
                ctype = sheet.cell(i, j).ctype
                # 获取单元格的值
                no = sheet.cell(i, j).value
                if ctype == 2:
                    # 将浮点转换成整数再转换成字符串
                    no = str(int(no))
                nos.append(no)
            # print(nos)
            # 写入sqlite
            cur.execute('insert into code_info values(?,?,?,?,?,?,?,?,?,?,?,?)',
                        (nos[0], nos[1], nos[2], nos[3], nos[4], nos[5], nos[6], nos[7], nos[8], nos[9],
                         nos[10], nos[11]))
        # for r in range(2, sheet.nrows):
        #     d = sheet.row_values(r)
        #     print(d)
        #     cur.execute('insert into code_info values(?,?,?,?,?,?,?,?,?,?,?,?)',
        #                 (d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9],
        #                  d[10], d[11]))
    else:
        raise Exception('代码级信息表不存在')

    con.commit()

    # 选择要执行的检查函数
    func_list = [r01_tab, r02_col, r03_cod, r04_tab, r05_col, r06_cod,
                 r07_tab, r08_col, r09_cod, r10_tab, r11_col, r12_col,
                 r13_col, r14_tab, r15_col, r16_cod, r17_tab, r18_col,
                 r19_col]
    # func_list = [r19_col]
    # 执行函数并输出结果
    for func in func_list:
        basename, extension = os.path.splitext(input_file)
        func_name = getattr(func, '__name__')
        func_doc = getattr(func, '__doc__')
        output_name = basename + '_' + func_name + '_' + func_doc + '({0})' + extension
        # 创建临时文件并写入文件头
        temp_file = '_t' + extension
        wb = xlsxwriter.Workbook(temp_file)
        if func_name[-3:] == "tab":
            sheet = wb.add_worksheet('表级信息')
            # 添加表头
            headings = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
                        '表描述', '表类型', '是否存在主键', '重要程度', '唯一索引', '非唯一索引', '表最新数据日期')
        elif func_name[-3:] == "col":
            sheet = wb.add_worksheet('字段级信息')
            headings = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名',
                        '表中文名', '字段序号', '字段英文名', '字段中文名', '字段类型', '是否主键',
                        '是否允许为空', '是否代码字段', '是否引用代码表', '字段注释')
        elif func_name[-3:] == "cod":
            sheet = wb.add_worksheet('代码级信息')
            headings = ('系统名称', '系统代码', '系统模块', '模式名（Schema）', '表分类', '表英文名', '表中文名',
                        '字段英文名', '字段中文名', '字段类型', '代码取值', '代码注释')

        sheet.write_row('A{0}'.format(1), headings)

        # 从第2行写入检核数据
        x = 2
        for i in func(cur):
            # print('  row_value: ', x, i[2])
            # 从第2行开始写入数据
            sheet.write_row('A{0}'.format(x), i[2])
            # 行递增
            x = x + 1
        # 关闭文件
        wb.close()
        # 将临时文件重命名为正式文件
        if x - 2 > 0:
            output_file = output_name.format(x - 2)
            print('output_file: ', output_file)
            os.rename(temp_file, output_file)
        else:
            os.remove(temp_file)
    # 关闭游标
    con.close()
