#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
使用专业飞书自动化脚本发送李白《静夜思》给朱宝
"""

from advanced_feishu_automation import FeiShuAutomation

def main():
    print("正在启动飞书自动化...")
    
    # 创建自动化实例
    automation = FeiShuAutomation()
    
    # 准备发送的信息
    contact_name = "朱宝"
    message_content = "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
    
    print(f"正在向 {contact_name} 发送消息...")
    print(f"消息内容: {message_content}")
    
    # 发送消息
    success = automation.send_message(contact_name, message_content)
    
    if success:
        print("✓ 消息发送成功！")
        print("李白的《静夜思》已成功发送给朱宝")
    else:
        print("✗ 消息发送失败！")
        print("可能的原因：")
        print("- 飞书窗口未找到")
        print("- 飞书未登录或未启动")
        print("- 联系人名称不正确")
        print("- 飞书界面结构发生变化")

if __name__ == "__main__":
    main()