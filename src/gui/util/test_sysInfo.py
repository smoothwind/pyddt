from unittest import TestCase

from src import SysInfo


class TestSysInfo(TestCase):
    def test_get_screen_size(self):
        screen_size = SysInfo.get_screen_size()
        if screen_size is None:
            self.fail("失败")
        elif not isinstance(screen_size, tuple):
            self.fail("返回错误")
        else:
            print(screen_size)
            pass
