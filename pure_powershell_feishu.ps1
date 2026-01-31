# 纯PowerShell飞书自动化脚本

Add-Type -AssemblyName System.Windows.Forms

try {
    # 确保飞书正在运行
    $feishuProcess = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    
    if ($null -eq $feishuProcess) {
        Write-Output "未找到飞书进程，尝试启动..."
        Start-Process -FilePath "D:\Feishu\app\Feishu.exe"
        Start-Sleep -Seconds 5  # 等待启动
        $feishuProcess = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    }
    
    if ($null -ne $feishuProcess) {
        Write-Output "找到飞书进程，正在激活窗口..."
        
        # 使用 Windows API 通过 P/Invoke
        $signature = @'
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
'@
        
        $user32 = Add-Type -MemberDefinition $signature -Name "User32" -PassThru
        
        # 恢复窗口并置顶
        if ($user32::IsIconic($feishuProcess.MainWindowHandle)) {
            $user32::ShowWindow($feishuProcess.MainWindowHandle, 9)  # SW_RESTORE
        }
        
        Start-Sleep -Milliseconds 500
        $user32::SetForegroundWindow($feishuProcess.MainWindowHandle)
        
        Write-Output "窗口已激活，准备发送消息..."
        Start-Sleep -Seconds 2
        
        # 设置剪贴板内容为联系人姓名
        Set-Clipboard -Value "朱宝"
        Start-Sleep -Milliseconds 200
        
        # 按下 Ctrl+K 打开搜索
        [System.Windows.Forms.SendKeys]::SendWait("^k")
        Write-Output "已按下搜索快捷键"
        Start-Sleep -Milliseconds 1000
        
        # 粘贴联系人姓名
        [System.Windows.Forms.SendKeys]::SendWait("^v")
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
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Write-Output "已粘贴诗歌内容"
        Start-Sleep -Milliseconds 500
        
        # 按回车发送消息
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Write-Output "已发送消息"
        
        Write-Output "李白的《静夜思》已成功发送给朱宝"
    } else {
        Write-Output "无法找到或启动飞书进程"
    }
} catch {
    Write-Output "发生错误: $($_.Exception.Message)"
    Write-Output "错误详情: $($_.Exception)"
}