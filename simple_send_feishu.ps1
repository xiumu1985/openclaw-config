# Simple PowerShell Script to Send Message via Feishu using Keyboard Shortcuts

param(
    [string]$Contact = "朱宝",
    [string]$Message = "你在干嘛"
)

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
    [Win32]::SetForegroundWindow($feishuProcess.MainWindowHandle)
    
    Start-Sleep -Milliseconds 1000
    
    # Wait for window to be active
    [System.Windows.Forms.SendKeys]::SendWait("%+k")  # Alt+Shift+K (quick search in Feishu)
    Start-Sleep -Milliseconds 500
    
    # Type the contact name
    [System.Windows.Forms.SendKeys]::SendWait($Contact)
    Start-Sleep -Milliseconds 1000
    
    # Press Enter to select the contact from search results
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 1500  # Wait for chat to open
    
    # Type the message
    [System.Windows.Forms.SendKeys]::SendWait($Message)
    Start-Sleep -Milliseconds 500
    
    # Press Enter to send the message
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    Write-Output "Successfully attempted to send message: '$($Message)' to '$($Contact)'"
    Write-Output "Please verify manually if the message was sent successfully."
    
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
    Write-Output "Full error details: $($_.Exception)"
}