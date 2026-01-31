import ctypes
from ctypes import wintypes
import time
import re

# Windows API常量定义
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
BM_CLICK = 0x00F5
WM_COMMAND = 0x0111
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_CONTROL = 0x11
VK_SHIFT = 0x10

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

def enum_windows_proc(hwnd, param):
    """枚举窗口回调函数"""
    # 获取窗口标题
    length = user32.GetWindowTextLengthW(hwnd)
    if length > 0:
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        
        # 检查是否是飞书主窗口
        if "飞书" in buff.value and user32.IsWindowVisible(hwnd):
            window_titles.append((hwnd, buff.value))
    
    return True

EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

def find_feishu_window():
    """查找飞书主窗口"""
    global window_titles
    window_titles = []
    enum_proc = EnumWindowsProc(enum_windows_proc)
    user32.EnumWindows(enum_proc, 0)
    
    for hwnd, title in window_titles:
        if "飞书" in title and "Feishu" in title:
            return hwnd
    
    # 如果没找到精确匹配，返回包含"飞书"的窗口
    for hwnd, title in window_titles:
        if "飞书" in title:
            return hwnd
            
    return None

def find_child_window_by_class(parent_hwnd, class_name):
    """通过类名查找子窗口"""
    child = user32.FindWindowExW(parent_hwnd, 0, class_name, None)
    return child

def find_child_window_by_title(parent_hwnd, title):
    """通过标题查找子窗口"""
    child = user32.FindWindowExW(parent_hwnd, 0, None, title)
    return child

def get_all_child_windows(parent_hwnd):
    """获取所有子窗口"""
    children = []
    
    def enum_child_proc(hwnd, param):
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            
            class_name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, class_name, 256)
            
            children.append({
                'hwnd': hwnd,
                'title': buff.value,
                'class': class_name.value
            })
        return True
    
    enum_proc = EnumWindowsProc(enum_child_proc)
    user32.EnumChildWindows(parent_hwnd, enum_proc, 0)
    
    return children

def send_text_to_window(hwnd, text):
    """向窗口发送文本"""
    # 使用SendMessage发送文本
    text_buffer = ctypes.c_wchar_p(text)
    result = user32.SendMessageW(hwnd, WM_SETTEXT, 0, text_buffer)
    return result

def get_window_text(hwnd):
    """获取窗口文本"""
    length = user32.GetWindowTextLengthW(hwnd)
    if length > 0:
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        return buff.value
    return ""

def find_contact_and_send_message(contact_name, message):
    """查找联系人并发送消息"""
    # 查找飞书主窗口
    feishu_hwnd = find_feishu_window()
    if not feishu_hwnd:
        print("未找到飞书窗口")
        return False
    
    print(f"找到飞书窗口: {get_window_text(feishu_hwnd)} (HWND: {feishu_hwnd})")
    
    # 获取所有子窗口以查找搜索框和聊天区域
    children = get_all_child_windows(feishu_hwnd)
    
    search_box = None
    contact_list = None
    message_input = None
    
    # 根据常见的控件类型查找相关组件
    for child in children:
        # 查找可能的搜索框（通常是编辑框）
        if child['class'] == 'Chrome_WidgetWin_0':
            # 尝试找到搜索输入框
            edit_hwnd = find_child_window_by_class(child['hwnd'], 'Chrome_RenderWidgetHostHWND')
            if edit_hwnd:
                search_box = edit_hwnd
        
        # 查找联系人列表
        if 'list' in child['class'].lower() or 'tree' in child['class'].lower():
            contact_list = child['hwnd']
        
        # 查找消息输入框
        if 'edit' in child['class'].lower() or 'richedit' in child['class'].lower():
            message_input = child['hwnd']
    
    # 如果没有找到搜索框，尝试更通用的方法
    if not search_box:
        # 遍历所有编辑控件
        for child in children:
            if 'Edit' in child['class'] or child['class'] == 'RichEdit20W' or child['class'] == 'RICHEDIT50W':
                # 这可能是搜索框或消息输入框
                if len(child['title']) == 0:  # 搜索框通常初始无标题
                    search_box = child['hwnd']
                    break
    
    # 如果没有找到消息输入框，继续寻找
    if not message_input:
        for child in children:
            if 'Edit' in child['class'] or child['class'] == 'RichEdit20W' or child['class'] == 'RICHEDIT50W':
                if '输入' in child['title'] or 'message' in child['title'].lower() or len(get_window_text(child['hwnd'])) > 0:
                    message_input = child['hwnd']
                    break
    
    print(f"搜索框 HWND: {search_box}")
    print(f"消息输入框 HWND: {message_input}")
    
    if not search_box:
        print("未找到搜索框")
        return False
    
    # 聚焦到搜索框
    user32.SetForegroundWindow(search_box)
    time.sleep(0.5)
    
    # 输入联系人姓名进行搜索
    send_text_to_window(search_box, contact_name)
    time.sleep(1)
    
    # 模拟按回车键选择第一个结果
    user32.PostMessageW(search_box, WM_KEYDOWN, VK_RETURN, 0)
    time.sleep(0.1)
    user32.PostMessageW(search_box, WM_KEYUP, VK_RETURN, 0)
    time.sleep(1)
    
    # 找到消息输入框（现在应该切换到了联系人的聊天界面）
    # 重新获取子窗口，因为界面可能已更新
    updated_children = get_all_child_windows(feishu_hwnd)
    
    for child in updated_children:
        if 'Edit' in child['class'] or child['class'] == 'RichEdit20W' or child['class'] == 'RICHEDIT50W':
            if '输入' in child['title'] or 'message' in child['title'].lower() or len(get_window_text(child['hwnd'])) == 0:
                message_input = child['hwnd']
                break
    
    if not message_input:
        print("未找到消息输入框")
        return False
    
    print(f"找到消息输入框: {message_input}")
    
    # 聚焦到消息输入框
    user32.SetFocus(message_input)
    time.sleep(0.2)
    
    # 发送消息
    send_text_to_window(message_input, message)
    time.sleep(0.5)
    
    # 模拟按Ctrl+Enter发送消息（或单独的发送按钮点击）
    user32.PostMessageW(message_input, WM_KEYDOWN, VK_CONTROL, 0)
    time.sleep(0.1)
    user32.PostMessageW(message_input, WM_KEYDOWN, VK_RETURN, 0)
    time.sleep(0.1)
    user32.PostMessageW(message_input, WM_KEYUP, VK_RETURN, 0)
    time.sleep(0.1)
    user32.PostMessageW(message_input, WM_KEYUP, VK_CONTROL, 0)
    
    print(f"已向 {contact_name} 发送消息: {message}")
    return True

def feishu_send_message(contact_name, message):
    """
    主函数：向指定联系人发送消息
    :param contact_name: 联系人姓名
    :param message: 要发送的消息
    """
    try:
        success = find_contact_and_send_message(contact_name, message)
        return success
    except Exception as e:
        print(f"发送消息时出错: {str(e)}")
        return False

# 更精确的实现版本
def feishu_send_message_advanced(contact_name, message):
    """
    改进版：更精确地控制飞书窗口
    """
    # 查找飞书主窗口
    feishu_hwnd = find_feishu_window()
    if not feishu_hwnd:
        print("未找到飞书窗口")
        return False
    
    print(f"飞书主窗口: {get_window_text(feishu_hwnd)}")
    
    # 确保窗口处于前台
    user32.ShowWindow(feishu_hwnd, 9)  # SW_RESTORE
    user32.SetForegroundWindow(feishu_hwnd)
    time.sleep(0.5)
    
    # 使用更精确的方法定位控件
    # 飞书的界面结构通常为：
    # 主窗口 -> 搜索栏/联系人面板 -> 聊天区域 -> 消息输入框
    
    # 尝试找到侧边栏中的搜索框（通常是第一个可见的编辑控件）
    children = get_all_child_windows(feishu_hwnd)
    
    # 查找搜索输入框（通常在顶部导航区域）
    search_input = None
    for child in children:
        if child['class'] in ['Chrome_SearchBox', 'Edit', 'ATL:00E72ED4', 'TEdit']:
            if "搜索" in child['title'] or "search" in child['title'].lower():
                search_input = child['hwnd']
                break
    
    # 如果没有找到特定的搜索框，使用更通用的方法
    if not search_input:
        # 寻找可能作为搜索框的编辑控件
        for child in children:
            if child['class'] in ['Edit', 'RICHEDIT50W', 'RichEdit20W', 'Chrome_RenderWidgetHostHWND']:
                # 获取当前窗口文本，如果为空可能是输入框
                if len(get_window_text(child['hwnd'])) == 0:
                    # 检查窗口大小，搜索框通常较宽
                    rect = RECT()
                    user32.GetWindowRect(child['hwnd'], ctypes.byref(rect))
                    width = rect.right - rect.left
                    height = rect.bottom - rect.top
                    if width > 100 and height < 50:  # 合理的搜索框尺寸
                        search_input = child['hwnd']
                        break
    
    if not search_input:
        print("未找到搜索输入框")
        return False
    
    # 清空搜索框并输入联系人姓名
    send_text_to_window(search_input, "")
    time.sleep(0.2)
    send_text_to_window(search_input, contact_name)
    time.sleep(0.5)
    
    # 按回车键进行搜索
    user32.PostMessageW(search_input, WM_KEYDOWN, VK_RETURN, 0)
    time.sleep(0.1)
    user32.PostMessageW(search_input, WM_KEYUP, VK_RETURN, 0)
    time.sleep(1.5)  # 等待搜索结果加载
    
    # 此时应该已经进入与联系人的聊天界面
    # 重新获取子窗口以找到消息输入区域
    updated_children = get_all_child_windows(feishu_hwnd)
    
    # 查找消息输入框（通常是底部的大编辑区域）
    message_input = None
    for child in updated_children:
        if child['class'] in ['RICHEDIT50W', 'RichEdit20W', 'Edit']:
            # 检查窗口文本和位置
            rect = RECT()
            user32.GetWindowRect(child['hwnd'], ctypes.byref(rect))
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            
            # 消息输入框通常在底部且有一定宽度
            window_rect = RECT()
            user32.GetWindowRect(feishu_hwnd, ctypes.byref(window_rect))
            window_bottom = window_rect.bottom - window_rect.top
            
            # 检查是否在窗口底部附近且有合理尺寸
            if height > 20 and width > 200 and (window_bottom - rect.bottom + window_rect.top) < 150:
                message_input = child['hwnd']
                break
    
    if not message_input:
        print("未找到消息输入框")
        return False
    
    # 输入消息
    send_text_to_window(message_input, message)
    time.sleep(0.3)
    
    # 发送消息 - 尝试多种方式
    # 方法1: Ctrl+Enter
    user32.PostMessageW(message_input, WM_KEYDOWN, VK_CONTROL, 0)
    time.sleep(0.05)
    user32.PostMessageW(message_input, WM_KEYDOWN, VK_RETURN, 0)
    time.sleep(0.05)
    user32.PostMessageW(message_input, WM_KEYUP, VK_RETURN, 0)
    time.sleep(0.05)
    user32.PostMessageW(message_input, WM_KEYUP, VK_CONTROL, 0)
    
    print(f"成功向 '{contact_name}' 发送消息: '{message}'")
    return True

if __name__ == "__main__":
    # 示例用法
    contact = input("请输入联系人姓名: ")
    msg = input("请输入要发送的消息: ")
    
    success = feishu_send_message_advanced(contact, msg)
    if success:
        print("消息发送成功！")
    else:
        print("消息发送失败！")