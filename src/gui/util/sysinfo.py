# -*- coding: UTF-8 -*-
import platform


class SysInfo:

    @staticmethod
    def get_screen_size():
        __sys__ = platform.system()
        if __sys__ == 'Windows':
            from win32api import GetSystemMetrics
            return GetSystemMetrics(0), GetSystemMetrics(1)
