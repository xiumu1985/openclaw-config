# 飞书自动化消息发送工具

这是一个使用Windows API直接操作飞书窗口的自动化脚本，能够可靠地向指定联系人发送中文消息。

## 功能特点

- 使用Windows API直接与飞书窗口交互
- 支持发送中文内容
- 不依赖键盘模拟，更加稳定可靠
- 自动查找联系人并发送消息

## 文件说明

- `feishu_automation.py`: 基础版本的自动化脚本
- `advanced_feishu_automation.py`: 改进版本，包含更好的错误处理和日志记录
- `README.md`: 本说明文件

## 使用方法

### 1. 准备工作

确保您的系统满足以下要求：
- Windows操作系统
- Python 3.6+
- 已安装飞书客户端并已登录

### 2. 运行脚本

运行改进版脚本（推荐）：

```bash
python advanced_feishu_automation.py
```

按照提示输入联系人姓名和消息内容。

### 3. 编程方式使用

您也可以在自己的代码中导入并使用：

```python
from advanced_feishu_automation import FeiShuAutomation

automation = FeiShuAutomation()
success = automation.send_message("联系人姓名", "要发送的消息")
```

## 注意事项

1. **权限要求**: 可能需要管理员权限才能正常操作其他应用程序窗口
2. **飞书窗口状态**: 飞书必须处于打开和登录状态
3. **窗口遮挡**: 确保飞书窗口没有被其他窗口完全遮挡
4. **联系人准确性**: 输入的联系人姓名必须完全匹配飞书中的名称

## 错误排查

如果遇到问题，请检查：

1. 飞书是否已经启动并登录
2. 是否以足够权限运行脚本
3. 联系人姓名拼写是否正确
4. 飞书窗口是否被其他程序遮挡

## 技术原理

该脚本使用Windows API函数：
- `FindWindowExW`: 查找窗口句柄
- `SendMessageW`: 向窗口发送消息
- `PostMessageW`: 发送窗口消息
- `SetForegroundWindow`: 激活窗口

通过直接操作窗口句柄，绕过了键盘模拟的不稳定性，确保中文内容能够正确发送。