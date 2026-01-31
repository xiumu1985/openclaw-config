# 飞书GUI自动化代理使用指南

## 概述

飞书GUI自动化代理是一个专门设计用来精确控制飞书应用的工具，能够自动发送中文内容给指定联系人。

## 安装依赖

在使用此工具前，请确保安装以下Python库：

```bash
pip install pyautogui pygetwindow pywin32
```

## 配置文件说明

### feishu_config.json

该文件包含了自动化代理的所有配置参数：

- `app_settings`: 应用程序相关设置
  - `app_name`: 应用名称（飞书/Lark）
  - `executable_paths`: 飞书可执行文件的可能路径
  - `protocol_uri`: 飞书协议URI

- `gui_coordinates`: GUI元素坐标设置（以屏幕比例表示）
  - `search_box`: 搜索框位置
  - `contact_item`: 联系人项目位置
  - `message_input`: 消息输入框位置

- `timing_settings`: 时间延迟设置
  - `pause_after_action`: 每个动作后的暂停时间
  - `app_launch_wait_time`: 应用启动等待时间
  - `search_wait_time`: 搜索等待时间
  - `contact_select_wait_time`: 联系人选择等待时间

- `features`: 功能开关
  - `enable_chinese_input`: 是否启用中文输入支持
  - `use_image_recognition`: 是否使用图像识别（暂未实现）
  - `require_user_confirmation`: 是否需要用户确认

## 使用方法

### 1. 基础使用

```python
from advanced_feishu_automation import AdvancedFeishuAutomationAgent

# 创建代理实例
agent = AdvancedFeishuAutomationAgent()

# 向联系人发送消息
success = agent.send_to_contact("联系人姓名", "要发送的中文消息")
```

### 2. 自定义配置

可以通过传递自定义配置文件路径来使用不同的配置：

```python
agent = AdvancedFeishuAutomationAgent("custom_config.json")
```

### 3. 调整坐标

如果默认坐标不准确，可以修改 `feishu_config.json` 中的 `gui_coordinates` 部分。
坐标值是相对于屏幕尺寸的比例值（0.0-1.0之间）。

例如：
- `x_ratio: 0.5` 表示水平方向的中心点
- `y_ratio: 0.1` 表示垂直方向距离顶部10%的位置

## 注意事项

1. **权限要求**: 此工具需要控制鼠标和键盘的权限
2. **屏幕分辨率**: 坐标是基于屏幕比例计算的，适用于不同分辨率
3. **中文支持**: 工具已优化中文输入，但可能需要手动切换输入法
4. **安全确认**: 建议保持 `require_user_confirmation` 开启以避免误发消息
5. **飞书版本**: 不同版本的飞书界面可能略有差异，需要相应调整坐标

## 故障排除

### 常见问题

1. **找不到飞书窗口**
   - 确认飞书已正确安装
   - 检查 `executable_paths` 中的路径是否正确
   - 确认飞书进程正在运行

2. **点击位置不准确**
   - 调整 `gui_coordinates` 中的坐标比例
   - 检查飞书界面布局是否与预期一致

3. **中文输入问题**
   - 确保系统输入法支持中文
   - 可尝试手动切换到中文输入法

4. **消息发送失败**
   - 检查联系人名称是否准确
   - 确认网络连接正常
   - 验证飞书账户权限

## 扩展功能

未来可以考虑添加的功能：
- 图像识别精确定位
- 消息发送状态确认
- 批量发送联系人列表
- 消息模板系统
- 日志记录和报告