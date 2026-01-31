import ctypes
from ctypes import wintypes
import time
import re
import logging
from typing import Optional, Tuple, List, Dict

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
                logger.info(f"找到飞书主窗口: {title}")
                return hwnd
        
        # 如果没找到完全匹配的，返回包含"飞书"的窗口
        for hwnd, title in self.window_titles:
            if "飞书" in title and "主窗口" in title:
                logger.info(f"找到飞书主窗口: {title}")
                return hwnd
                
        # 最后返回任何包含"飞书"的可见窗口
        for hwnd, title in self.window_titles:
            if "飞书" in title:
                logger.info(f"找到飞书窗口: {title}")
                return hwnd
                
        logger.warning("未找到飞书窗口")
        return None

    def find_child_window_by_class(self, parent_hwnd: int, class_name: str) -> Optional[int]:
        """通过类名查找子窗口"""
        child = user32.FindWindowExW(parent_hwnd, 0, class_name, None)
        return child if child != 0 else None

    def find_child_window_by_title(self, parent_hwnd: int, title: str) -> Optional[int]:
        """通过标题查找子窗口"""
        child = user32.FindWindowExW(parent_hwnd, 0, None, title)
        return child if child != 0 else None

    def get_all_child_windows(self, parent_hwnd: int) -> List[Dict]:
        """获取所有子窗口信息"""
        children = []
        
        def enum_child_proc(hwnd, param):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                
                class_name = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, class_name, 256)
                
                rect = RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                
                children.append({
                    'hwnd': hwnd,
                    'title': buff.value,
                    'class': class_name.value,
                    'rect': rect
                })
            return True
        
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(enum_child_proc)
        user32.EnumChildWindows(parent_hwnd, enum_proc, 0)
        
        return children

    def send_text_to_window(self, hwnd: int, text: str) -> bool:
        """向窗口发送文本"""
        try:
            # 确保窗口可编辑
            user32.EnableWindow(hwnd, True)
            user32.SetFocus(hwnd)
            
            # 清空现有文本
            user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.c_wchar_p(""))
            time.sleep(0.1)
            
            # 发送新文本
            result = user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.c_wchar_p(text))
            time.sleep(0.2)  # 给予时间让文本显示
            
            return result != 0
        except Exception as e:
            logger.error(f"发送文本到窗口失败: {e}")
            return False

    def simulate_key_combination(self, hwnd: int, *keys: int):
        """模拟按键组合"""
        # 按下所有键
        for key in keys:
            user32.PostMessageW(hwnd, WM_KEYDOWN, key, 0)
            time.sleep(0.02)
        
        time.sleep(0.05)  # 延迟确保组合键生效
        
        # 释放所有键（按相反顺序）
        for key in reversed(keys):
            user32.PostMessageW(hwnd, WM_KEYUP, key, 0)
            time.sleep(0.02)

    def click_window(self, hwnd: int):
        """点击窗口"""
        # 获取窗口中心点
        rect = RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        
        center_x = (rect.left + rect.right) // 2
        center_y = (rect.top + rect.bottom) // 2
        
        # 转换为相对于父窗口的坐标
        screen_point = POINT(center_x, center_y)
        user32.ScreenToClient(hwnd, ctypes.byref(screen_point))
        
        # 发送鼠标事件
        lparam = screen_point.x | (screen_point.y << 16)
        user32.PostMessageW(hwnd, WM_LBUTTONDOWN, 1, lparam)
        time.sleep(0.05)
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0, lparam)

    def activate_feishu_window(self, hwnd: int) -> bool:
        """激活飞书窗口"""
        try:
            # 恢复窗口（如果最小化）
            user32.ShowWindow(hwnd, SW_RESTORE)
            # 将窗口带到前台
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.error(f"激活窗口失败: {e}")
            return False

    def find_search_box(self, feishu_hwnd: int) -> Optional[int]:
        """查找搜索框"""
        children = self.get_all_child_windows(feishu_hwnd)
        
        # 常见的飞书搜索框类名
        potential_classes = [
            'Chrome_WidgetWin_0', 'Edit', 'RICHEDIT50W', 
            'RichEdit20W', 'Chrome_RenderWidgetHostHWND',
            'TEdit', 'ATL:00E72ED4'
        ]
        
        for child in children:
            # 检查类名
            if child['class'] in potential_classes:
                # 检查窗口尺寸和位置
                width = child['rect'].right - child['rect'].left
                height = child['rect'].bottom - child['rect'].top
                
                # 搜索框通常在顶部区域，有一定宽度但高度适中
                window_rect = RECT()
                user32.GetWindowRect(feishu_hwnd, ctypes.byref(window_rect))
                
                # 检查是否在顶部区域
                relative_top = child['rect'].top - window_rect.top
                if relative_top < 100 and width > 100 and height < 50:
                    logger.info(f"找到潜在搜索框: {child['title']} (Class: {child['class']})")
                    return child['hwnd']
        
        # 如果没有找到明显的搜索框，尝试根据文本提示查找
        for child in children:
            if child['class'] in potential_classes:
                # 检查是否可能包含搜索功能
                if any(keyword in child['title'].lower() for keyword in ['搜索', 'search', '输入']):
                    logger.info(f"找到搜索框: {child['title']}")
                    return child['hwnd']
        
        # 尝试通过遍历所有编辑控件来找到最可能的搜索框
        for child in children:
            if child['class'] in ['Edit', 'RICHEDIT50W', 'RichEdit20W']:
                # 检查位置和大小
                width = child['rect'].right - child['rect'].left
                height = child['rect'].bottom - child['rect'].top
                relative_top = child['rect'].top - window_rect.top
                
                if relative_top < 80 and width > 150 and height < 40:
                    logger.info(f"找到搜索框候选: {child['title']}")
                    return child['hwnd']
        
        logger.warning("未找到搜索框")
        return None

    def find_message_input(self, feishu_hwnd: int) -> Optional[int]:
        """查找消息输入框"""
        children = self.get_all_child_windows(feishu_hwnd)
        
        # 消息输入框的常见类名
        potential_classes = ['RICHEDIT50W', 'RichEdit20W', 'Edit']
        
        # 获取主窗口的位置信息
        window_rect = RECT()
        user32.GetWindowRect(feishu_hwnd, ctypes.byref(window_rect))
        window_height = window_rect.bottom - window_rect.top
        
        for child in children:
            if child['class'] in potential_classes:
                # 消息输入框通常位于窗口底部
                relative_bottom = window_rect.bottom - child['rect'].bottom
                width = child['rect'].right - child['rect'].left
                height = child['rect'].bottom - child['rect'].top
                
                # 检查是否在底部且有合适尺寸
                if relative_bottom < 150 and width > 200 and height > 30:
                    logger.info(f"找到消息输入框: {child['title']}")
                    return child['hwnd']
        
        # 如果没有找到明显的输入框，尝试其他策略
        for child in children:
            if child['class'] in potential_classes:
                # 检查是否有"输入"相关的文本
                if any(keyword in child['title'] for keyword in ['输入', 'message', 'text']):
                    logger.info(f"找到消息输入框: {child['title']}")
                    return child['hwnd']
        
        logger.warning("未找到消息输入框")
        return None

    def select_contact_from_search(self, search_box_hwnd: int, contact_name: str) -> bool:
        """从搜索结果中选择联系人"""
        # 搜索完成后，通常需要选择联系人
        # 这里我们假设搜索后第一个结果是我们要找的
        time.sleep(1.0)  # 等待搜索完成
        
        # 按向下箭头键选择第一个结果
        user32.PostMessageW(search_box_hwnd, WM_KEYDOWN, VK_DOWN, 0)
        time.sleep(0.1)
        user32.PostMessageW(search_box_hwnd, WM_KEYUP, VK_DOWN, 0)
        time.sleep(0.1)
        
        # 按回车键确认选择
        user32.PostMessageW(search_box_hwnd, WM_KEYDOWN, VK_RETURN, 0)
        time.sleep(0.05)
        user32.PostMessageW(search_box_hwnd, WM_KEYUP, VK_RETURN, 0)
        
        time.sleep(1.0)  # 等待切换到聊天界面
        return True

    def send_message(self, contact_name: str, message: str) -> bool:
        """发送消息给指定联系人"""
        try:
            # 1. 查找飞书窗口
            feishu_hwnd = self.find_feishu_window()
            if not feishu_hwnd:
                logger.error("未找到飞书窗口")
                return False
            
            logger.info(f"找到飞书窗口: {feishu_hwnd}")
            
            # 2. 激活窗口
            if not self.activate_feishu_window(feishu_hwnd):
                logger.error("无法激活飞书窗口")
                return False
            
            # 3. 查找搜索框
            search_box = self.find_search_box(feishu_hwnd)
            if not search_box:
                logger.error("未找到搜索框")
                return False
            
            # 4. 在搜索框中输入联系人姓名
            if not self.send_text_to_window(search_box, contact_name):
                logger.error("无法在搜索框输入联系人姓名")
                return False
            
            time.sleep(0.5)
            
            # 5. 按回车键开始搜索
            user32.PostMessageW(search_box, WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.05)
            user32.PostMessageW(search_box, WM_KEYUP, VK_RETURN, 0)
            
            # 6. 等待搜索结果并选择联系人
            time.sleep(1.5)
            self.select_contact_from_search(search_box, contact_name)
            
            # 7. 查找消息输入框（现在应该在聊天界面）
            message_input = self.find_message_input(feishu_hwnd)
            if not message_input:
                logger.error("未找到消息输入框")
                return False
            
            # 8. 输入消息
            if not self.send_text_to_window(message_input, message):
                logger.error("无法在消息框输入消息")
                return False
            
            # 9. 发送消息（使用Ctrl+Enter组合键）
            self.simulate_key_combination(message_input, VK_CONTROL, VK_RETURN)
            
            logger.info(f"成功向 '{contact_name}' 发送消息: '{message}'")
            return True
            
        except Exception as e:
            logger.error(f"发送消息时发生错误: {e}")
            return False

def main():
    """主函数示例"""
    automation = FeiShuAutomation()
    
    print("飞书自动化消息发送工具")
    print("=" * 40)
    
    contact_name = input("请输入联系人姓名: ").strip()
    if not contact_name:
        print("联系人姓名不能为空!")
        return
    
    message = input("请输入要发送的消息: ").strip()
    if not message:
        print("消息内容不能为空!")
        return
    
    print(f"\n正在向 '{contact_name}' 发送消息...")
    print(f"消息内容: {message}")
    
    success = automation.send_message(contact_name, message)
    
    if success:
        print("\n✓ 消息发送成功!")
    else:
        print("\n✗ 消息发送失败!")
        print("请确保:")
        print("- 飞书已启动并登录")
        print("- 联系人姓名正确")
        print("- 飞书窗口没有被其他窗口遮挡")

if __name__ == "__main__":
    main()