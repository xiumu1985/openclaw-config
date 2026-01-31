import time
import subprocess

def manual_send():
    print("手动执行飞书消息发送流程...")
    
    # 步骤1: 激活飞书主窗口
    print("1. 查找并激活飞书窗口...")
    try:
        import ctypes
        from ctypes import wintypes
        
        user32 = ctypes.windll.user32
        
        def enum_windows_proc(hwnd, param):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buff, length + 1)
                
                # 查找飞书主窗口
                if "飞书" in buff.value and user32.IsWindowVisible(hwnd):
                    # 恢复并激活窗口
                    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                    user32.SetForegroundWindow(hwnd)
                    print(f"   已激活窗口: {buff.value}")
                    time.sleep(1)
                    return False  # 找到后停止枚举
            return True
        
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(
            enum_windows_proc
        )
        user32.EnumWindows(enum_proc, 0)
        
    except Exception as e:
        print(f"   窗口激活时出错: {e}")
    
    time.sleep(2)
    
    # 步骤2: 使用剪贴板和键盘模拟发送消息
    print("2. 准备发送消息...")
    try:
        import pyperclip
        import ctypes
        import ctypes.wintypes
        
        # 设置剪贴板内容为联系人名称
        pyperclip.copy("朱宝")
        time.sleep(0.5)
        
        # 模拟快捷键和输入
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        
        # 模拟 Ctrl+K (快速查找)
        user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL
        time.sleep(0.05)
        user32.keybd_event(0x4B, 0, 0, 0)  # K
        time.sleep(0.05)
        user32.keybd_event(0x4B, 0, 2, 0)  # 释放K
        time.sleep(0.05)
        user32.keybd_event(0x11, 0, 2, 0)  # 释放VK_CONTROL
        time.sleep(0.5)
        
        # 模拟 Ctrl+V (粘贴联系人名称)
        user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL
        time.sleep(0.05)
        user32.keybd_event(0x56, 0, 0, 0)  # V
        time.sleep(0.05)
        user32.keybd_event(0x56, 0, 2, 0)  # 释放V
        time.sleep(0.05)
        user32.keybd_event(0x11, 0, 2, 0)  # 释放VK_CONTROL
        time.sleep(1)
        
        # 模拟回车选择联系人
        user32.keybd_event(0x0D, 0, 0, 0)  # VK_RETURN
        time.sleep(0.05)
        user32.keybd_event(0x0D, 0, 2, 0)  # 释放VK_RETURN
        time.sleep(2)
        
        # 设置剪贴板内容为消息内容
        pyperclip.copy("床前明月光，疑是地上霜。举头望明月，低头思故乡。")
        time.sleep(0.5)
        
        # 模拟 Ctrl+V (粘贴消息内容)
        user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL
        time.sleep(0.05)
        user32.keybd_event(0x56, 0, 0, 0)  # V
        time.sleep(0.05)
        user32.keybd_event(0x56, 0, 2, 0)  # 释放V
        time.sleep(0.05)
        user32.keybd_event(0x11, 0, 2, 0)  # 释放VK_CONTROL
        time.sleep(0.5)
        
        # 模拟回车发送消息
        user32.keybd_event(0x0D, 0, 0, 0)  # VK_RETURN
        time.sleep(0.05)
        user32.keybd_event(0x0D, 0, 2, 0)  # 释放VK_RETURN
        
        print("   消息发送指令已执行")
        
    except ImportError:
        print("   需要安装pyperclip库: pip install pyperclip")
    except Exception as e:
        print(f"   发送消息时出错: {e}")
    
    print("操作完成，请检查飞书应用。")

if __name__ == "__main__":
    manual_send()