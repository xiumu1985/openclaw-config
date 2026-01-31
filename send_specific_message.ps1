# Simple PowerShell Script to Send Specific Message via Feishu using Keyboard Shortcuts

# Load required assembly
Add-Type -AssemblyName System.Windows.Forms

try {
    # Get the Feishu process
    $feishuProcess = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    
    if ($null -eq $feishuProcess) {
        Write-Output "No Feishu process with visible window found."
        exit 1
    }
    
    Write-Output "Found Feishu process with ID $($feishuProcess.Id)"
    
    # Use Win32 API to bring window to foreground
    Add-Type @"
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
    
    # Restore window if minimized
    if ([Win32]::IsIconic($feishuProcess.MainWindowHandle)) {
        [Win32]::ShowWindow($feishuProcess.MainWindowHandle, 9)  # SW_RESTORE
    }
    
    # Bring window to foreground
    Start-Sleep -Milliseconds 500
    [Win32]::SetforegroundWindow($feishuProcess.MainWindowHandle)
    
    Start-Sleep -Milliseconds 1000
    
    # Wait for window to be active
    [System.Windows.Forms.SendKeys]::SendWait("%+k")  # Alt+Shift+K (quick search in Feishu)
    Start-Sleep -Milliseconds 500
    
    # Type the contact name
    [System.Windows.Forms.SendKeys]::SendWait("朱宝")
    Start-Sleep -Milliseconds 1000
    
    # Press Enter to select the contact from search results
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 1500  # Wait for chat to open
    
    # Type the specific message "你在干嘛呢"
    [System.Windows.Forms.SendKeys]::SendWait("你在干嘛呢")
    Start-Sleep -Milliseconds 500
    
    # Press Enter to send the message
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    Write-Output "Successfully attempted to send message: '你在干嘛呢' to '朱宝'"
    Write-Output "Please verify manually if the message was sent successfully."
    
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
    Write-Output "Full error details: $($_.Exception)"
}