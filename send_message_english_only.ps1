# Script to send Chinese message via Feishu using clipboard method
# All comments in English to avoid parsing issues

Add-Type -AssemblyName System.Windows.Forms

try {
    $feishuProcess = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    
    if ($null -eq $feishuProcess) {
        Write-Output "No Feishu process with visible window found."
        exit 1
    }
    
    Write-Output "Found Feishu process with ID $($feishuProcess.Id)"
    
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
    
    if ([Win32]::IsIconic($feishuProcess.MainWindowHandle)) {
        [Win32]::ShowWindow($feishuProcess.MainWindowHandle, 9)
    }
    
    Start-Sleep -Milliseconds 500
    [Win32]::SetForegroundWindow($feishuProcess.MainWindowHandle)
    
    Start-Sleep -Milliseconds 1000
    
    [System.Windows.Forms.SendKeys]::SendWait("%+k")
    Start-Sleep -Milliseconds 500
    
    Set-Clipboard -Value "Zhu Bao"  # Contact name in clipboard
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 1000
    
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 1500
    
    Set-Clipboard -Value "Ni hao ma"  # Message in clipboard
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 500
    
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    Write-Output "Successfully attempted to send message via clipboard method."
    Write-Output "Please verify manually if the message was sent successfully."
    
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
}