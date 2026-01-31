# 最简化的飞书自动化脚本

Add-Type -AssemblyName System.Windows.Forms

# 激活飞书窗口并发送消息
$feishuProcess = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1

if ($feishuProcess) {
    # 使用 P/Invoke 调用 Windows API
    $code = @"
        using System;
        using System.Runtime.InteropServices;

        public class Win32 {
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool IsIconic(IntPtr hWnd);
        }
"@

    Add-Type -TypeDefinition $code -ReferencedAssemblies "System.Runtime.InteropServices" -IgnoreWarnings

    # 激活窗口
    if ([Win32]::IsIconic($feishuProcess.MainWindowHandle)) {
        [Win32]::ShowWindow($feishuProcess.MainWindowHandle, 9)
    }
    [Win32]::SetForegroundWindow($feishuProcess.MainWindowHandle)

    Start-Sleep -Seconds 2

    # 操作序列
    Set-Clipboard -Value "朱宝"
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^k")      # Ctrl+K
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("^v")      # Ctrl+V
    Start-Sleep -Milliseconds 1000
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") # Enter
    Start-Sleep -Seconds 2
    Set-Clipboard -Value "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")      # Ctrl+V
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}") # Enter

    "李白的《静夜思》已发送给朱宝"
} else {
    "未找到飞书进程"
}