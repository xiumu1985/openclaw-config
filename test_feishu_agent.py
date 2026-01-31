"""
飞书自动化代理测试脚本
用于验证各项功能是否正常工作
"""

import time
from advanced_feishu_automation import AdvancedFeishuAutomationAgent


def test_basic_functionality():
    """测试基本功能"""
    print("开始测试飞书自动化代理...")
    
    # 创建代理实例
    agent = AdvancedFeishuAutomationAgent()
    
    print("✓ 代理实例创建成功")
    
    # 测试配置加载
    assert agent.config is not None, "配置应成功加载"
    print("✓ 配置加载正常")
    
    # 测试坐标计算
    test_coord = type('', (), {'x_ratio': 0.5, 'y_ratio': 0.5})()  # Mock坐标对象
    x, y = agent.calculate_absolute_coords(test_coord)
    assert x > 0 and y > 0, "坐标计算应返回有效值"
    print("✓ 坐标计算正常")
    
    print("基本功能测试完成")


def demo_usage():
    """演示使用方法"""
    print("\n" + "="*50)
    print("飞书自动化代理演示")
    print("="*50)
    
    agent = AdvancedFeishuAutomationAgent()
    
    print("\n接下来将演示:")
    print("1. 查找或启动飞书应用")
    print("2. 搜索指定联系人")
    print("3. 向联系人发送中文消息")
    
    # 获取用户输入
    contact_name = input("\n请输入要发送消息的联系人姓名 (直接回车使用默认值): ").strip()
    if not contact_name:
        contact_name = "测试联系人"  # 默认值
    
    message_content = input("请输入要发送的中文消息 (直接回车使用默认值): ").strip()
    if not message_content:
        message_content = "这是一条测试消息，用于验证飞书自动化代理的功能。"
    
    print(f"\n准备向 '{contact_name}' 发送消息...")
    print(f"消息内容: {message_content}")
    
    # 执行发送操作
    success = agent.send_to_contact(contact_name, message_content)
    
    if success:
        print("\n✓ 消息发送成功!")
    else:
        print("\n✗ 消息发送失败!")
        print("可能的原因:")
        print("- 飞书未正确安装")
        print("- 联系人名称不正确")
        print("- 飞书界面布局与预期不符")
        print("- 权限不足")


def show_configuration():
    """显示当前配置信息"""
    print("\n" + "-"*30)
    print("当前配置信息")
    print("-"*30)
    
    agent = AdvancedFeishuAutomationAgent()
    config = agent.config
    
    print(f"应用名称: {config['app_settings']['app_name']}")
    print(f"可执行文件路径数量: {len(config['app_settings']['executable_paths'])}")
    print(f"中文输入支持: {config['features']['enable_chinese_input']}")
    print(f"用户确认要求: {config['features']['require_user_confirmation']}")
    print(f"动作间暂停时间: {config['timing_settings']['pause_after_action']}秒")
    
    print("\nGUI坐标设置:")
    coords = config['gui_coordinates']
    for element, pos in coords.items():
        print(f"  {element}: ({pos['x_ratio']:.2f}, {pos['y_ratio']:.2f})")


if __name__ == "__main__":
    print("飞书GUI自动化代理 - 测试工具")
    print("请选择要执行的操作:")
    print("1. 运行基本功能测试")
    print("2. 显示当前配置")
    print("3. 运行完整演示")
    
    choice = input("请输入选项 (1/2/3): ").strip()
    
    if choice == "1":
        test_basic_functionality()
    elif choice == "2":
        show_configuration()
    elif choice == "3":
        demo_usage()
    else:
        print("无效选项，运行完整演示...")
        demo_usage()