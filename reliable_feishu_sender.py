import subprocess
import time
import pyautogui
import pygetwindow as gw
import os

def send_message_to_zhubao():
    """
    更可靠的飞书自动化脚本
    """
    try:
        print("正在寻找飞书窗口...")
        
        # 尝试通过进程激活飞书
        subprocess.run(["tasklist", "/FI", "IMAGENAME eq Feishu.exe"], check=True)
        
        # 等待窗口加载
        time.sleep(3)
        
        # 寻找飞书窗口
        feishu_windows = []
        for window in gw.getAllWindows():
            if '飞书' in window.title or 'Lark' in window.title or 'Feishu' in window.title:
                feishu_windows.append(window)
        
        if feishu_windows:
            # 激活第一个找到的飞书窗口
            win = feishu_windows[0]
            win.activate()
            time.sleep(1)
            print(f"已激活飞书窗口: {win.title}")
        else:
            print("未找到飞书窗口，尝试启动飞书...")
            # 尝试启动飞书
            subprocess.Popen([r"D:\Feishu\app\Feishu.exe"])
            time.sleep(5)  # 等待应用启动
        
        # 按下 Ctrl+K 打开搜索
        print("按下快捷键打开搜索...")
        pyautogui.hotkey('ctrl', 'k')
        time.sleep(1)
        
        # 输入联系人姓名 "朱宝"
        print("输入联系人姓名...")
        pyautogui.write('朱宝', interval=0.2)
        time.sleep(1)
        
        # 按下回车选择联系人
        pyautogui.press('enter')
        time.sleep(2)
        
        # 输入李白的《静夜思》
        poem = "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
        print("输入李白的《静夜思》...")
        pyautogui.write(poem, interval=0.1)
        time.sleep(1)
        
        # 按下回车发送消息
        pyautogui.press('enter')
        
        print("消息已发送给朱宝：李白的《静夜思》")
        
    except ImportError as e:
        print(f"缺少必要的Python库: {e}")
        print("请运行: pip install pyautogui pygetwindow")
    except Exception as e:
        print(f"执行过程中出现错误: {e}")

if __name__ == "__main__":
    send_message_to_zhubao()