import ctypes
from ctypes import wintypes
import time
import logging
from typing import Optional, List, Dict

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Windows API常量定义
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
BM_CLICK = 0x00F5
WM_COMMAND = 0x0111
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_MOUSEMOVE = 0x0200
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_CONTROL = 0x11
VK_SHIFT = 0x10
VK_DOWN = 0x28
SW_SHOW = 5
SW_RESTORE = 9

# 定义Windows API函数
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

class RECT(ctypes.Structure):
    _fields_ = [
        ('left', wintypes.LONG),
        ('top', wintypes.LONG),
        ('right', wintypes.LONG),
        ('bottom', wintypes.LONG)
    ]

class POINT(ctypes.Structure):
    _fields_ = [('x', wintypes.LONG), ('y', wintypes.LONG)]

class FeiShuAutomation:
    def __init__(self):
        self.window_titles = []
        self.hwnd_cache = {}

    def enum_windows_proc(self, hwnd, param):
        """枚举窗口回调函数"""
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            
            # 检查是否是飞书主窗口
            if "飞书" in buff.value and user32.IsWindowVisible(hwnd):
                self.window_titles.append((hwnd, buff.value))
        
        return True

    def find_feishu_window(self) -> Optional[int]:
        """查找飞书主窗口"""
        self.window_titles = []
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(self.enum_windows_proc)
        user32.EnumWindows(enum_proc, 0)
        
        # 优先查找标题完全匹配的窗口
        for hwnd, title in self.window_titles:
            if "飞书" in title and "Feishu" in title.lower():
                return hwnd
        
        # 如果没找到完全匹配的，返回包含"飞书"的窗口
        for hwnd, title in self.window_titles:
            if "飞书" in title and "主窗口" in title:
                return hwnd
                
        # 最后返回任何包含"飞书"的可见窗口
        for hwnd, title in self.window_titles:
            if "飞书" in title:
                return hwnd
                
        return None

if __name__ == "__main__":
    automation = FeiShuAutomation()
    hwnd = automation.find_feishu_window()
    print(f"Found Feishu window: {hwnd}")