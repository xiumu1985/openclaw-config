#SingleInstance Force
#NoEnv
SetBatchLines, -1

; 激活飞书窗口并发送李白的《静夜思》给朱宝

; 尝试激活飞书窗口
WinGet, FeishuList, List, ahk_exe Feishu.exe

if (FeishuList > 0) {
    ; 获取最后一个（最新的）飞书窗口
    WinActivate, % "ahk_id " FeishuList%FeishuList%
    WinWaitActive, % "ahk_id " FeishuList%FeishuList%, , 3
    if (!ErrorLevel) {
        ; 窗口已激活
        Sleep, 1000
        
        ; 按下 Ctrl+K 打开搜索
        Send, ^k
        Sleep, 1000
        
        ; 输入联系人姓名 "朱宝"
        Clipboard := "朱宝"
        Send, ^v
        Sleep, 1000
        
        ; 按回车选择联系人
        Send, {Enter}
        Sleep, 2000  ; 等待对话窗口加载
        
        ; 输入李白的《静夜思》
        Clipboard := "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
        Send, ^v
        Sleep, 500
        
        ; 按回车发送消息
        Send, {Enter}
        
        MsgBox, 李白的《静夜思》已成功发送给朱宝
    } else {
        MsgBox, 无法激活飞书窗口
    }
} else {
    MsgBox, 未找到飞书进程
}