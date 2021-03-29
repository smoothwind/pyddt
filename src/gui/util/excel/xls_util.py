# -*- coding: UTF-8 -*-
import os.path as oph
import win32com.client as win32
import xlrd
import xlutils.copy as xls_copy
import xlwt

from src.gui.util.config import LOG

__all__ = ['write_content', 'ms_convert']


def parser_content(reader, content_name):
    """
    根据ExeclReader对象文件解析目录
    :param reader: xlrd.book.Book 对象
    :param content_name: str 目录sheet的名称
    :return: contents
    """
    contents = []
    # openpyxl.Workbook.get_sheet_by_name()

    """ 
        for sheet in reader.book.sheets():
        if sheet.name != '目录':
            # print(sheet.cell_value(1,1),sheet.cell_value(1,3))
            hyperlink = "=HYPERLINK(\"#%s!A1\",\"%s\")" % (sheet.name, sheet.name)
            contents.append((sheet.cell_value(1, 1), sheet.cell_value(1, 3), hyperlink))
    """
    for sheet in reader.sheet_names():
        if sheet != content_name:
            # print(sheet.cell_value(1,1),sheet.cell_value(1,3))
            hyperlink = "HYPERLINK(\"#%s!A1\",\"%s\")" % (sheet, sheet)
            ws = reader.sheet_by_name(sheet)
            contents.append((ws.cell_value(1, 1), ws.cell_value(1, 3), xlwt.Formula(hyperlink)))

    LOG.debug("目录解析完毕")
    return contents


def write_content(file_path, content_name=None):
    """
    自动写入目录
    :param file_path: 文件名称
    :param content_name: 指定目录页名称
    :return: None
    """
    if not oph.exists(file_path):
        LOG.error("文件不存在 %s" % file_path)
        print("文件不存在 %s" % file_path)
        return

    if oph.splitext(file_path)[-1][1:] not in ['xls', 'xlsx']:
        LOG.error("不支持此文件格式 %s" % file_path)
        print("不支持此文件格式 %s" % file_path)
        return

    if content_name is not None:
        if isinstance(content_name, str):
            content_name = content_name
        else:
            content_name = "CONTENTS"
    else:
        content_name = "CONTENTS"

    _xls_reader = xlrd.open_workbook(file_path)
    _contents = parser_content(_xls_reader, content_name)
    _wb = xls_copy.copy(_xls_reader)
    # 到此处用不到了，释放
    _xls_reader.release_resources()

    _table = None
    try:
        _table = _wb.sheet_by_name(content_name)
    except:
        _table = _wb.add_sheet(content_name, cell_overwrite_ok=True)
    finally:
        if _table is None:
            LOG.error("无法打开或创建Sheet: %s" % content_name)
            return

    _table.write(0, 0, "表名")
    _table.write(0, 1, "备注")
    _table.write(0, 2, "链接")
    for _i, _content in enumerate(_contents):
        # 写入目录
        _table.write(_i + 1, 0, _content[0])
        _table.write(_i + 1, 1, _content[1])
        _table.write(_i + 1, 2, _content[2])
        _sheet_name = _content[0]
        if _content[1] != "":
            _sheet_name = _content[1]

        ### 更新返回
        tgt_sheet = _wb.get_sheet(_sheet_name)
        # _address = '=HYPERLINK("#%s!A%d","%s")' % (content_name, int(_i + 1), content_name)
        # tgt_sheet.write(1, 4, _address)
        # _address = 'HYPERLINK("#%s!A%d","%s")' % (content_name, int(_i + 1), content_name)
        tgt_sheet.write(1, 4, xlwt.Formula('HYPERLINK("#{}!A1"; "{}")'.format(content_name, content_name)))

        LOG.debug("写入目录：%s-%s-%s" % _content)
    """
    table.write(1, 0, '内容1')  # 括号内分别为行数、列数、内容
    table.write(1, 1, '内容2')
    table.write(1, 2, '内容3')
    """
    print(type(_wb))
    # _wb.save(file_path) # 不兼容xlsx
    if oph.splitext(file_path)[-1][1:] == ".xls":
        _wb.save(file_path)
    else:
        # todo: 待支持xls2xlsx
        file_path.replace(".xlsx", ".xls")
        _wb.save(file_path)
    return


if __name__ == '__main__':
    write_content("export/outp2.xlsx")
    '''
    # reader = pd.ExcelFile("export/outp2.xlsx")
    xlsx_reader = xlrd.open_workbook("export/outp2.xlsx")
    contents = parser_content(xlsx_reader)
    print(contents)
    wb = xls_copy.copy(xlsx_reader)
    try:
        table = wb.sheet_by_name('目录')
    except:
        table = wb.add_sheet('目录')
    finally:
        table.write(0, 0, "表名")
        table.write(0, 1, "备注")
        table.write(0, 2, "链接")
        for i, content in enumerate(contents):
            table.write(i+1, 0, content[0])
            table.write(i+1, 1, content[1])
            table.write(i+1, 2, content[2])
            sheet_name = content[0]
            if content[1] != "":
                sheet_name = content[1]
            tgt_sheet = wb.get_sheet(sheet_name)
            addr = '=HYPERLINK(CELL("address", 目录!A%d),"目录")' % int(i+1)
            tgt_sheet.write(1, 4, addr)
            LOG.debug("写入目录：%s-%s-%s" % content)
        """
        table.write(1, 0, '内容1')  # 括号内分别为行数、列数、内容
        table.write(1, 1, '内容2')
        table.write(1, 2, '内容3')
        """

    # wb.save("export/outp2.xls")
    '''

    """ 待处理 pandas 不支持增量更新 execl
    writer = pd.ExcelWriter("export/outp2.xlsx")
    pd.DataFrame(contents).to_excel(writer, sheet_name="目录")
    writer.save();
    writer.close()
    """

"""
    for i, item in enumerate(contents):
        #writer.write_cells([], sheet_name="目录", startrow=i, startcol=0)
        #writer.write_cells(item[1], sheet_name="目录", startrow=i, startcol=1)
        #writer.write_cells(item[2], sheet_name="目录", startrow=i, startcol=2)

        writer.write_cells(contents, sheet_name="目录")
        writer.write_cells("=HYPERLINK(CELL(\"address\", 目录!A1),\"目录\")", sheet_name=item[0], startrow=1, startcol=6)

    writer.close()
"""


def ms_convert(src_file: str, target_file: str) -> bool:
    """
    execl文件类型转换
    :param src_file: 源文件地址
    :param target_file: 目的文件地址
    :return:
    """

    if not oph.exists(src_file):
        LOG.error("文件路径不存在：%s" % src_file)
        return False

    src_file = oph.abspath(src_file)
    target_file = oph.abspath(target_file)
    print("%s --> %s" % (src_file, target_file))

    excel = win32.gencache.EnsureDispatch('Excel.Application')

    wb = excel.Workbooks.Open(src_file)
    wb.SaveAs(target_file)
    wb.Close()
    excel.Application.Quit()
    return True
