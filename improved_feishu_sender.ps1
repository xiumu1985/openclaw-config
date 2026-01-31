# 改进的飞书自动化脚本
# 使用更稳定的窗口激活和消息发送机制

Add-Type -AssemblyName System.Windows.Forms

try {
    # 设置剪贴板内容
    Set-Clipboard -Value "朱宝"
    Write-Output "已设置联系人到剪贴板"
    
    Start-Sleep -Milliseconds 200
    
    # 使用Win32 API激活飞书窗口
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;

        public class FeishuActivator
        {
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
            
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
            
            [DllImport("user32.dll")]
            public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder strText, int maxCount);
            
            [DllImport("user32.dll")]
            public static extern int GetWindowTextLength(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern bool IsWindowVisible(IntPtr hWnd);
            
            [DllImport("user32.dll")]
            public static extern IntPtr GetShellWindow();
            
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        }
"@

    # 激活飞书窗口
    $feishuProcesses = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue
    if ($feishuProcesses) {
        $feishuProcess = $feishuProcesses | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
        
        if ($feishuProcess) {
            Write-Output "发现飞书进程，正在激活窗口..."
            
            # 尝试激活窗口
            [FeishuActivator]::ShowWindow($feishuProcess.MainWindowHandle, 9)  # SW_RESTORE
            Start-Sleep -Milliseconds 1000
            [FeishuActivator]::SetForegroundWindow($feishuProcess.MainWindowHandle)
            
            Write-Output "窗口已激活，等待2秒..."
            Start-Sleep -Seconds 2
            
            # 执行键盘操作序列
            # 按下 Alt+Shift+K 打开快速搜索
            [System.Windows.Forms.SendKeys]::SendWait("%+k")  # Alt+Shift+K
            Write-Output "已按下搜索快捷键"
            
            Start-Sleep -Milliseconds 1000
            
            # 粘贴联系人姓名
            [System.Windows.Forms.SendKeys]::SendWait("^v")  # Ctrl+V
            Write-Output "已粘贴联系人姓名"
            
            Start-Sleep -Milliseconds 1000
            
            # 按回车选择联系人
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Write-Output "已选择联系人"
            
            Start-Sleep -Seconds 2  # 等待对话窗口加载
            
            # 设置诗歌内容到剪贴板
            Set-Clipboard -Value "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
            Start-Sleep -Milliseconds 200
            
            # 粘贴诗歌内容
            [System.Windows.Forms.SendKeys]::SendWait("^v")  # Ctrl+V
            Write-Output "已粘贴诗歌内容"
            
            Start-Sleep -Milliseconds 500
            
            # 按回车发送消息
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Write-Output "已发送消息"
            
            Write-Output "李白的《静夜思》已成功发送给朱宝"
        } else {
            Write-Output "未找到具有主窗口句柄的飞书进程"
        }
    } else {
        Write-Output "未找到飞书进程"
    }
} catch {
    Write-Output "发生错误: $($_.Exception.Message)"
}