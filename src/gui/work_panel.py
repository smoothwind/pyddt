#!/usr/bin/python
# -*- coding: UTF-8 -*-
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import cx_Oracle as Cx
from docx import Document
from pandas import ExcelWriter

from util.config import LOG
from util.excel.xls_util import write_content
from util.sql.description import write_to_word, write_to_execl, get_table_docs, split_table_name
from util.sql.oracle import GET_ALL_TABLE
from util.sysinfo import SysInfo


class ObjectViewer:
    table_list: dict = {}
    selected_list = set()
    tab_tree: ttk.Treeview
    table_desc_list = []
    window: tk.Tk
    frame_top: tk.Frame
    frame_mid: tk.Frame
    frame_mid_left: tk.Frame
    frame_mid_left_top: tk.Frame
    frame_mid_left_bottom: tk.Frame
    frame_mid_right: tk.Frame
    frame_mid_right_right: tk.Frame
    frame_bottom: tk.Frame
    btn_add: tk.Button
    btn_del: tk.Button
    btn_clean: tk.Button
    btn_add_all: tk.Button
    btn_download: tk.Button
    box_prepared_export: tk.Listbox

    def __init__(self, connection):
        self.window = tk.Tk()
        self.window.title("Oracle数据字典生成工具")

        if connection is None:
            messagebox.showerror("error", "数据库链接不能为None！")
        if isinstance(connection, Cx.Connection) is not True:
            messagebox.showerror("error", "不是有效的数据库链接！")

        self.__screen__ = SysInfo.get_screen_size()
        if self.__screen__ is not None:
            self.__main_width__ = self.__screen__[0] / 2
            self.__main_height__ = self.__screen__[1] / 2

        self.__conn__: Cx.Connection = connection
        self.__cursor__: Cx.Cursor = connection.cursor()

        self.frame_top = tk.Frame(self.window, height=0.1 * self.__main_height__, width=self.__main_width__)
        self.frame_mid = tk.Frame(self.window, height=0.8 * self.__main_height__, width=self.__main_width__)

        self.frame_mid_left = tk.Frame(self.frame_mid, height=0.8 * self.__main_height__,
                                       width=0.35 * self.__main_width__)
        self.frame_mid_left_top = tk.Frame(self.frame_mid_left, width=0.28 * self.__main_width__)
        self.frame_mid_left_bottom = tk.Frame(self.frame_mid_left, width=0.28 * self.__main_width__)

        self.frame_mid_right = tk.Frame(self.frame_mid, height=0.8 * self.__main_height__,
                                        width=0.6 * self.__main_width__)

        self.frame_mid_right_left = tk.Frame(self.frame_mid_right, height=0.8 * self.__main_height__,
                                             width=0.1 * self.__main_width__)
        self.frame_mid_right_right = tk.Frame(self.frame_mid_right, height=0.8 * self.__main_height__,
                                              width=0.5 * self.__main_width__)

        self.frame_bottom = tk.Frame(self.window, height=0.1 * self.__main_height__, width=self.__main_width__)
        self.frame_top.pack(side="top")
        self.frame_mid_left_top.pack(side='top')
        self.frame_mid_left_bottom.pack(side='bottom')
        self.frame_mid_left.pack(side='left')

        self.frame_mid_right_left.pack(side="left")
        self.frame_mid_right_right.pack(side="right")
        self.frame_mid_right.pack(side="right")
        self.frame_mid.pack()

        self.frame_bottom.pack(side="bottom")

        self.tab_tree = ttk.Treeview(self.frame_mid_left_top, selectmode='extended', padding=[10, 0, 5, 10])
        self.__get_all_tables__()
        self.set_tree()
        self.tab_tree.heading("#0", text="数据库对象浏览器", anchor="w")
        self.tab_tree.pack(side='left', fill='x')
        self.vsb = tk.Scrollbar(self.frame_mid_left_top, orient="vertical", command=self.tab_tree.yview)
        self.vsb.pack(side='right', fill='y')
        self.tab_tree.configure(yscrollcommand=self.vsb.set)

        self.hsb = ttk.Scrollbar(self.frame_mid_left_bottom, orient="horizontal", command=self.tab_tree.xview)
        self.hsb.pack(side='bottom', fill='x')
        self.tab_tree.configure(xscrollcommand=self.hsb.set)

        self.btn_add = tk.Button(self.frame_mid_right_left, text="->", width=10, command=self.btn_add_click)
        self.btn_del = tk.Button(self.frame_mid_right_left, text="<-", width=10, command=self.btn_del_click)
        self.btn_add_all = tk.Button(self.frame_mid_right_left, text="->>", width=10, command=self.btn_add_all_click)
        self.btn_clean = tk.Button(self.frame_mid_right_left, text="<<-", width=10, command=self.btn_clean_click)
        self.btn_download = tk.Button(self.frame_bottom, text="生成文档", width=10, command=self.btn_download_click)
        self.box_prepared_export = tk.Listbox(self.frame_mid_right_right, selectmode="multiple")

        self.b_vsb = tk.Scrollbar(self.frame_mid_right_right, orient="vertical", command=self.box_prepared_export.yview)
        self.b_hsb = tk.Scrollbar(self.frame_mid_right_right, orient="horizontal",
                                  command=self.box_prepared_export.xview)
        self.box_prepared_export.configure(yscrollcommand=self.b_vsb.set, xscrollcommand=self.b_hsb.set)
        self.b_vsb.pack(side="right", fill='y')
        self.b_hsb.pack(side="bottom", fill='x')

        self.btn_add.grid(row=3)
        self.btn_del.grid(row=5)
        self.btn_add_all.grid(row=7)
        self.btn_clean.grid(row=9)
        self.btn_download.grid()
        self.box_prepared_export.pack()
        self.window.geometry("%dx%d" % (self.__main_width__, self.__main_height__))
        self.window.size()
        self.window.mainloop()

    def __get_all_tables__(self):
        """
        :return: 查询所有非SYS用户的表
        """
        if len(self.table_list) == 0:
            res = self.__cursor__.execute(GET_ALL_TABLE).fetchall()
            for i, tab in enumerate(res):
                if self.table_list.__contains__(tab[0]):
                    self.table_list.get(tab[0]).append("%s.%s" % (tab[0], tab[1]))
                else:
                    self.table_list.update({tab[0]: ["%s.%s" % (tab[0], tab[1])]})
            # print(self.table_list)
        return True

    def get_all_tables(self):
        if len(self.table_list) == 0:
            self.__get_all_tables__()
        return self.table_list

    def get_all_users(self):
        if len(self.table_list.keys()) == 0:
            self.__get_all_tables__()
        return self.table_list.keys()

    def set_tree(self):
        owners = list(self.table_list.keys())
        owners.sort()
        for i, owner in enumerate(owners):
            branch = self.tab_tree.insert("", i, text=owner, values=owner)
            tabs = list(self.table_list.get(owner))
            tabs.sort()
            for j, tab in enumerate(tabs):
                self.tab_tree.insert(branch, j, text=tab, values=tab)

    def btn_add_click(self):
        add_items = []
        for i, it in enumerate(self.tab_tree.selection()):
            valid_table_name: str = self.tab_tree.item(it, "values")[0]
            if valid_table_name.__contains__("."):
                add_items.append(valid_table_name)
        self.selected_list.update(add_items)
        self.insert_into_box()
        print("已从列表中添加了%s" % add_items)

    def btn_del_click(self):
        """
        从待导出列表中，将选中的项目删除
        :return:
        """
        del_set = self.box_prepared_export.curselection()
        if len(list(del_set)) == 0:
            LOG.info("待删除列表：未选中任何项！")
            return
        else:
            for i, it in enumerate(del_set):
                del_item = self.box_prepared_export.get(it)
                self.selected_list.discard(del_item)
                LOG.info("已从列表中删除：%s" % del_item)

        self.insert_into_box()
        print("已从列表中删除")

    def btn_add_all_click(self):
        print("添加所有表到选中列表%s" % self.table_list.values())
        for i, val in enumerate(list(self.table_list.keys())):
            for j, it in enumerate(self.table_list.get(val)):
                self.selected_list.add(it)
            # self.selected_list.add(self.table_list.get(val))

        self.insert_into_box()

    def btn_clean_click(self):
        print("清空已选!")
        self.selected_list.clear()
        self.insert_into_box()

    def insert_into_box(self):
        items = list(self.selected_list)
        items.sort()
        counts = self.box_prepared_export.size()
        if counts != 0:
            self.box_prepared_export.delete(first=0, last=counts)
        if len(items) != 0:
            for i, it in enumerate(items):
                self.box_prepared_export.insert(i, it)

    def btn_download_click(self):
        # 禁用
        self.btn_download.configure(state="disabled")
        # 开线程执行下载操作
        try:
            t_download = threading.Thread(target=self.process_doc_download)
            t_download.setDaemon(True)
            t_download.start()
        except threading.ThreadError as err:
            messagebox.showerror("错误", "线程开启失败！\n%s" % err)
            self.btn_download.configure(state="active")
        else:
            self.btn_download.configure(state="active")

    def process_doc_download(self, file_output_path=None):
        """
        文档生成
        todo： 待处理事项： 1.word文件style定制;  word 文件目录生成。
        :param file_output_path: 文件保存路径
        :return: None
        """
        if file_output_path is None:
            import os
            if not os.path.exists("export"):
                os.mkdir("export")
            now = datetime.now()
            file_output_path = "export/dict_%s.xls" % now.strftime("%Y-%m-%d_%H-%M-%S")

        if len(list(self.selected_list)) == 0:
            messagebox.showinfo("提示:", "未选中任何表!")
            return

        writer = ExcelWriter(file_output_path)
        document = Document()  # todo: Just for testing, please delete it

        for i, tab in enumerate(list(self.selected_list)):
            owner = ""
            tab_name = ""
            try:
                owner, tab_name = split_table_name(tab)
            except ValueError:
                # 无法解析出用户名
                continue

            dd = get_table_docs(self.__cursor__, owner, tab_name)
            write_to_execl(writer, dd)
            write_to_word(document=document, dd=dd)  # todo: Just for testing, please delete it
        write_content(file_output_path)
        file_output_path_word = file_output_path.split(".")[0] + ".docx"
        document.save(file_output_path_word)  # todo:  Just for testing, please delete it

        self.btn_download.configure(state="active")
        messagebox.showinfo("提示", "文档《%s》\n  《%s》\n已生成完毕！" % (file_output_path, file_output_path_word))


if __name__ == '__main__':
    obj = ObjectViewer(None)
