# -*- coding: UTF-8 -*-
import tkinter as tk
from configparser import ConfigParser
from tkinter import messagebox
from tkinter import ttk

import cx_Oracle as Cx

from src.gui.work_panel import ObjectViewer


class EntryWin:
    connection_string: str
    conn_list: list = []
    default_conn = ""
    btn_get_conn: ttk.Button
    btn_test_conn: ttk.Button
    ini_file_name = ""
    cf: ConfigParser
    db_cursor = None
    window: tk.Tk
    lbl_db_conn: tk.Label
    combobox_conn_list_text: tk.StringVar
    combobox_conn_list: ttk.Combobox

    def __init__(self, ini_file_name):
        self.connection_string = ""
        if ini_file_name is None:
            self.ini_file_name = ""
        self.cf = ConfigParser()
        self.cf.read(filenames=self.ini_file_name, encoding="utf-8")
        self.conn_list = EntryWin.parse_conn_list(cf=self.cf, section="oracle", key="list")
        self.window = tk.Tk()
        self.window.title("Oracle数据字典生成工具")
        self.window.geometry("280x100")

        self.lbl_db_conn = ttk.Label(self.window, text="数据库链接")
        self.combobox_conn_list_text = tk.StringVar()
        self.combobox_conn_list_text.set('请输入或选择连接')
        self.combobox_conn_list = ttk.Combobox(values=self.conn_list, textvariable=self.combobox_conn_list_text)
        self.btn_test_conn = ttk.Button(self.window, text="测试",
                                        command=self.button_test_connection)
        self.btn_get_conn = ttk.Button(self.window, text="下一步",
                                       command=self.button_get_connection)

        """
            控件布局
        """
        self.lbl_db_conn.grid(row=3, stick="w", pady=10)
        self.combobox_conn_list.grid(row=3, column=1, stick="e")
        self.btn_test_conn.grid(row=4, stick="w", pady=10)
        self.btn_get_conn.grid(row=4, column=1, stick="e")

        """
        事件绑定
        """
        self.combobox_conn_list.bind('<<ComboboxSelected>>', self.combobox_conn_list_event_handler)
        self.combobox_conn_list.bind('<FocusOut>', self.combobox_conn_list_event_handler, add="+")
        self.window.size()
        self.window.mainloop()

    @staticmethod
    def parse_conn_list(cf: ConfigParser, section, key) -> list:
        values: list = []
        if isinstance(cf, ConfigParser) and str(section).__len__() > 0 and str(key).__len__() > 0 and \
                cf is not None:
            try:
                value_str = cf.get(section, key)
                values = list(value_str.replace("[", "").replace("]", "").replace('"', "").replace("'", "").strip()
                              .split(","))
                print(values)
            except BaseException as err:
                print(err)
        return values

    @staticmethod
    def save_conn_list(cf, path, section, key, value_list):
        if isinstance(cf, ConfigParser) and str(section).__len__() > 0 and str(key).__len__() > 0 \
                and cf is not None and len(value_list) > 0:
            try:
                cf.set(section, key, repr(value_list))
                cf.write(open(path, 'w', encoding="utf-8"))
            except BaseException as err:
                print(err)
        else:
            print("链接配置保存错误！")

    def button_test_connection(self):
        self.btn_test_conn.configure(state="disable")
        _conn: Cx.Connection
        self.default_conn = self.combobox_conn_list_text.get()

        if isinstance(self.connection_string, str):
            try:
                _conn = Cx.connect(self.connection_string)
                if _conn is not None:
                    messagebox.showinfo("连接测试", "连接测试成功！")
                    EntryWin.save_conn_list(cf=self.cf, path=self.ini_file_name, section="oracle", key="list",
                                            value_list=self.conn_list)
                else:
                    messagebox.showerror("连接测试", "连接测试失败! \n[数据库链接: %s]" % self.connection_string)
            except (Cx.DatabaseError, Exception) as err:
                error_tips = ""
                if len(self.connection_string) == 0:
                    error_tips = "[数据库链接不能为空!]"
                else:
                    error_tips = "[数据库链接: %s]" % self.connection_string.strip()
                messagebox.showerror("连接测试", " 连接数据库发生异常!\n错误信息:%s \n%s" % (err, error_tips))
            finally:
                if isinstance(_conn, Cx.Connection):
                    _conn.close()
            self.btn_test_conn.configure(state="active")
        else:
            messagebox.showinfo("连接测试", "连接测试发生异常！")
        self.btn_test_conn.configure(state="active")
        return False

    def button_get_connection(self):
        self.btn_get_conn.configure(state='disable')
        if isinstance(self.connection_string, str):
            if len(self.connection_string.strip()) == 0:
                self.connection_string = self.combobox_conn_list_text.get()
            try:
                self.db_cursor = Cx.connect(self.connection_string)
                messagebox.showinfo("提示", "数据库已连接！")
                self.btn_get_conn.configure(state='active')
                obj = ObjectViewer(self.db_cursor)
                EntryWin.save_conn_list(cf=self.cf, path=self.ini_file_name, section="oracle", key="list",
                                        value_list=self.conn_list)
                return True
            except (Cx.DatabaseError, Exception) as err:
                messagebox.showerror("提示", "连接数据库发生异常, 错误信息:%s " % err)
                self.btn_get_conn.configure(state='active')

        self.btn_get_conn.configure(state='active')
        return False

    def combobox_conn_list_event_handler(self, event):
        """
        组合下拉框选中
        :param event: 事件
        :return: null
        """
        # tk.Event.mro()
        self.default_conn = self.combobox_conn_list.get()
        if not self.conn_list.__contains__(self.default_conn):
            self.conn_list.append(self.default_conn)
        if event.type == tk.EventType.FocusOut:
            print("失去焦点")
        else:
            messagebox.showinfo("提示", "已选择连接:" + self.default_conn)
