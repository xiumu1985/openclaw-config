' 简单的VBScript自动化脚本
WScript.Sleep 2000

' 激活飞书窗口 (通过Alt+Tab循环直到找到)
CreateObject("WScript.Shell").SendKeys "^+{ESC}"  ' 打开任务视图
WScript.Sleep 1000

' 尝试通过快捷键激活飞书
CreateObject("WScript.Shell").SendKeys "%{TAB}"  ' Alt+Tab
WScript.Sleep 500
CreateObject("WScript.Shell").SendKeys "%{TAB}"  ' 继续Alt+Tab
WScript.Sleep 500

' 打开搜索 (Ctrl+K)
CreateObject("WScript.Shell").SendKeys "^k"
WScript.Sleep 1000

' 输入联系人 "朱宝"
CreateObject("WScript.Shell").SendKeys "朱宝"
WScript.Sleep 1000

' 按Enter选择联系人
CreateObject("WScript.Shell").SendKeys "~"  ' Enter key
WScript.Sleep 2000

' 输入李白的《静夜思》
CreateObject("WScript.Shell").SendKeys "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
WScript.Sleep 500

' 按Enter发送消息
CreateObject("WScript.Shell").SendKeys "~"  ' Enter key

WScript.Echo "消息已尝试发送给朱宝"