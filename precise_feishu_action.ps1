# 精确的飞书操作脚本

# 添加类型定义用于Windows API调用
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool SetActiveWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    public const int WM_KEYDOWN = 0x0100;
    public const int WM_KEYUP = 0x0101;
    public const int VK_CONTROL = 0x11;
    public const int VK_V = 0x56;
    public const int VK_K = 0x4B;
    public const int VK_RETURN = 0x0D;
    public const int SW_RESTORE = 9;
}

public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
"@

# 回调函数用于枚举窗口
$callback = [Win32+EnumWindowsProc] {
    param($hWnd, $lParam)
    
    $length = [Win32]::GetWindowTextLength($hWnd)
    if ($length -gt 0) {
        $sb = New-Object System.Text.StringBuilder($length + 1)
        [Win32]::GetWindowText($hWnd, $sb, $sb.Capacity)
        $title = $sb.ToString()
        
        # 检查是否是飞书窗口
        if ($title -match "飞书" -and [Win32]::IsWindowVisible($hWnd)) {
            Write-Host "发现飞书窗口: $title"
            
            # 激活窗口
            [Win32]::ShowWindow($hWnd, 9)
            [Win32]::SetForegroundWindow($hWnd)
            Start-Sleep -Milliseconds 1000
            
            # 现在执行操作序列
            # 1. 设置剪贴板为联系人名称
            Set-Clipboard -Value "朱宝"
            Start-Sleep -Milliseconds 200
            
            # 2. 模拟 Ctrl+K
            [Win32]::keybd_event([byte]17, 0, 0, [UIntPtr]::Zero)  # Ctrl down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]75, 0, 0, [UIntPtr]::Zero)  # K down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]75, 0, 2, [UIntPtr]::Zero)  # K up
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]17, 0, 2, [UIntPtr]::Zero) # Ctrl up
            Start-Sleep -Milliseconds 500
            
            # 3. 模拟 Ctrl+V (粘贴联系人名称)
            [Win32]::keybd_event([byte]17, 0, 0, [UIntPtr]::Zero)  # Ctrl down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]86, 0, 0, [UIntPtr]::Zero)  # V down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]86, 0, 2, [UIntPtr]::Zero)  # V up
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]17, 0, 2, [UIntPtr]::Zero) # Ctrl up
            Start-Sleep -Milliseconds 1000
            
            # 4. 按回车选择联系人
            [Win32]::keybd_event([byte]13, 0, 0, [UIntPtr]::Zero)  # Enter down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]13, 0, 2, [UIntPtr]::Zero) # Enter up
            Start-Sleep -Seconds 2
            
            # 5. 设置剪贴板为消息内容
            Set-Clipboard -Value "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
            Start-Sleep -Milliseconds 200
            
            # 6. 模拟 Ctrl+V (粘贴消息内容)
            [Win32]::keybd_event([byte]17, 0, 0, [UIntPtr]::Zero)  # Ctrl down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]86, 0, 0, [UIntPtr]::Zero)  # V down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]86, 0, 2, [UIntPtr]::Zero)  # V up
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]17, 0, 2, [UIntPtr]::Zero) # Ctrl up
            Start-Sleep -Milliseconds 500
            
            # 7. 按回车发送消息
            [Win32]::keybd_event([byte]13, 0, 0, [UIntPtr]::Zero)  # Enter down
            Start-Sleep -Milliseconds 50
            [Win32]::keybd_event([byte]13, 0, 2, [UIntPtr]::Zero) # Enter up
            
            Write-Host "操作完成：已尝试向朱宝发送《静夜思》"
            return $false  # 停止枚举
        }
    }
    return $true  # 继续枚举
}

# 开始枚举窗口
[Win32]::EnumWindows($callback, [IntPtr]::Zero)

Write-Host "脚本执行完成"