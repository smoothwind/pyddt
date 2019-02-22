# -*- coding: UTF-8 -*-
import os
import tkinter as tk
from configparser import ConfigParser
from tkinter import ttk
from tkinter import messagebox

from src.gui.work_panel import ObjectViewer
import cx_Oracle as Cx

ini_file_name = 'options.ini'


def parse_conn_list(cf, section, key):
    values: list = []
    if isinstance(cf, ConfigParser) and str(section).__len__() > 0 and str(key).__len__() > 0 and\
            cf is not None:
        try:
            value_str = cf.get(section, key)
            values = list(value_str.replace("[", "").replace("]", "").replace('"', "").replace("'", "").strip()
                          .split(","))
            print(values)
        except BaseException as err:
            print(err)
    return values


def save_conn_list(cf, path, section, key, value_list):
    if isinstance(cf, ConfigParser) and str(section).__len__() > 0 and str(key).__len__() > 0 and cf is not None and\
            len(value_list) > 0:
        try:
            cf.set(section, key, repr(value_list))
            cf.write(open(path, 'w',encoding="utf-8"))
        except BaseException as err:
            print(err)
    else:
        print("链接配置保存错误！")


if __name__ == '__main__':

    """
    ############################################# 1。变量 ##############################################################
    """
    Db_Cursor = None
    default_conn = ""
    conn_list: list = []
    cf = ConfigParser()
    cf.read(filenames=ini_file_name, encoding="utf-8")
    conn_list = parse_conn_list(cf=cf, section="oracle", key="list")

    """
    ################################################ 2.事件处理函数 ########################################################
    """


    def combobox_conn_list_event_handler(event):
        """
        组合下拉框选中
        :param event: 事件
        :return: null
        """
        # tk.Event.mro()
        global default_conn
        default_conn = combobox_conn_list.get()
        if not conn_list.__contains__(default_conn):
            conn_list.append(default_conn)

        if event.type == tk.EventType.FocusOut:
            print("失去焦点")
        else:
            messagebox.showinfo("提示", "已选择连接:" + default_conn)


    """
    ########################################## 窗体及组件定义 ###########################################################
    """
    window = tk.Tk()
    window_width = 400
    window_height = 150
    window.title("Oracle数据字典生成工具")
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    align_fmt = "%dx%d+%d+%d" % (window_width, window_height, (screen_width - window_width) / 2,
                                 (screen_height - window_height) / 2)
    window.geometry(align_fmt)
    """
    各类控件配置
    """
    lbl_db_conn: tk.Label = ttk.Label(window, text="数据库链接")
    # ttk.Button(window, text="测试", command=button_test_connection).grid(column=1, row=2)
    combobox_conn_list_text: tk.StringVar = tk.StringVar()
    combobox_conn_list_text.set('请输入或选择连接')
    combobox_conn_list: ttk.Combobox = ttk.Combobox(values=conn_list, textvariable=combobox_conn_list_text)
    btn_test_conn: ttk.Button = ttk.Button(window, text="测试", command=lambda: button_test_connection(default_conn))
    btn_get_conn = ttk.Button(window, text="下一步", command=lambda: button_get_connection(default_conn))
    """
    控件布局
    """
    lbl_db_conn.grid(row=3, stick="w", pady=10)
    combobox_conn_list.grid(row=3, column=1, stick="e")
    btn_test_conn.grid(row=4, stick="w", pady=10)
    btn_get_conn.grid(row=4, column=1, stick="e")

    """
    事件绑定
    """
    combobox_conn_list.bind('<<ComboboxSelected>>', combobox_conn_list_event_handler)
    combobox_conn_list.bind('<FocusOut>', combobox_conn_list_event_handler, add="+")


    def button_test_connection(connection_string):
        _conn = None
        global btn_test_conn
        btn_test_conn.configure(state="disable")
        print(combobox_conn_list_text.get())

        if isinstance(connection_string, str):
            try:
                _conn = Cx.connect(connection_string)
                if _conn is not None:
                    messagebox.showinfo("连接测试", "连接测试成功！")
                    save_conn_list(cf=cf, path=ini_file_name, section="oracle", key="list", value_list=conn_list)
                else:
                    messagebox.showerror("连接测试", "连接测试失败! \n[数据库链接: %s]" % connection_string)
            except (Cx.DatabaseError, Exception) as err:
                error_tips = ""
                if len(connection_string) == 0:
                    error_tips = "[数据库链接不能为空!]"
                else:
                    error_tips = "[数据库链接: %s]" % connection_string.strip()
                messagebox.showerror("连接测试", " 连接数据库发生异常!\n错误信息:%s \n%s" % (err, error_tips))
            finally:
                if isinstance(_conn, Cx.Connection):
                    _conn.close()
            btn_test_conn.configure(state="active")
        else:
            messagebox.showinfo("连接测试", "连接测试发生异常！")
        btn_test_conn.configure(state="active")
        return False


    def button_get_connection(connection_string):
        global btn_get_conn
        btn_get_conn.configure(state='disable')
        if isinstance(connection_string, str):
            if len(connection_string.strip()) == 0:
                connection_string = combobox_conn_list_text.get()
            try:
                global Db_Cursor
                Db_Cursor = Cx.connect(connection_string)
                messagebox.showinfo("提示", "数据库已连接！")
                btn_get_conn.configure(state='active')
                obj = ObjectViewer(Db_Cursor)
                save_conn_list(cf=cf, path=ini_file_name, section="oracle", key="list", value_list=conn_list)
                return True
            except (Cx.DatabaseError, Exception) as err:
                messagebox.showerror("提示", "连接数据库发生异常, 错误信息:%s " % err)
                btn_get_conn.configure(state='active')

        btn_get_conn.configure(state='active')
        return False


    window.size()
    window.mainloop()
