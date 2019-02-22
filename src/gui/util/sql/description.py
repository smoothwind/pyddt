# -*- coding: UTF-8 -*-
import cx_Oracle as cx
from pandas import DataFrame
from src.gui.util.config import LOG
from .oracle import TAB_COMMENT_STATS, TABLE_DES_STATS, GET_ALL_USER_OBJ

__all__ = ['get_table_docs', 'write_to_execl', 'split_table_name']

"""
"""


def get_table_docs(cursor, table_owner, table_name):
    table_name = table_name.upper().strip()
    if isinstance(cursor, cx.Cursor):
        _tab_comments = TAB_COMMENT_STATS % (repr(table_owner), repr(table_name))
        _tab_descs = TABLE_DES_STATS % (repr(table_owner), repr(table_name))
        _comment = cursor.execute(_tab_comments).fetchall()
        _desc = cursor.execute(_tab_descs).fetchall()
        comment = DataFrame(_comment)
        desc = DataFrame(_desc)
        return {'col_desc': comment, 'tab_comment': desc}
    else:
        LOG.error('get_table_docs: \'cursor\' is not a cx_Oracle.Cursor type. ')
        LOG.error('get_table_docs: additional info: %s ' % type(cursor))
    return None


def write_to_execl(writer, dd):
    col_desc = DataFrame(dd.get('tab_comment'))
    tab_comments = DataFrame(dd.get('col_desc'))
    sheet_name: str = tab_comments.loc[0, 1]

    if not str(tab_comments.loc[0, 3]).strip().isspace() and tab_comments.loc[0, 3] is not None:
        sheet_name = str(tab_comments.loc[0, 3])
    print(sheet_name)
    print(type(sheet_name))
    try:
        tab_comments.to_excel(writer, sheet_name=sheet_name, index=False, header=['用户', '表名', '类型', '备注'], startrow=0)
        col_desc.to_excel(writer, sheet_name=sheet_name, index=False,
                          header=['序号', '列英文名', '数据类型', '是否可空', '默认值', '列中文名'], startrow=5, inf_rep="")
    except:
        LOG.error("%s write failed!" % sheet_name)
    finally:
        writer.save()


def split_table_name(table_name: str):
    """
    分割表名
    :param table_name: 表名 “owner.table”
    :return: 返回 元组类型的表名（owner， table）
    """
    return table_name.split('.', maxsplit=2)


def get_all_table(cursor: cx.Cursor, owner: str):
    """
    获取某个用户下的所有表
    :param cursor: 游标 类型 cx_Oracle.Cursor
    :param owner:  用户 schema
    :return:
    """
    STMTS = GET_ALL_USER_OBJ % repr(owner.upper().strip())
    res = cursor.execute(STMTS).fetchall()
    tab_list = []
    for i, tables in enumerate(res):
        tab_list.append(tables[1])
    return tab_list
