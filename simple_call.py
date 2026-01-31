# 简化的调用脚本
import sys
import os

# 将当前目录添加到Python路径
sys.path.append(os.getcwd())

# 导入自动化类
from advanced_feishu_automation import FeiShuAutomation

# 创建实例并调用
automation = FeiShuAutomation()
result = automation.send_message('朱宝', '床前明月光，疑是地上霜。举头望明月，低头思故乡。')
print(f"发送结果: {result}")