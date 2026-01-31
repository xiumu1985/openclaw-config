# Activate Feishu and send the specific message using clipboard method
Add-Type -AssemblyName System.Windows.Forms

# Known Feishu process ID from our earlier checks
$feishuProcessId = 1668

try {
    # Get the specific Feishu process
    $feishuProcess = Get-Process -Id $feishuProcessId -ErrorAction SilentlyContinue
    
    if ($null -eq $feishuProcess) {
        Write-Output "Cannot find Feishu process with ID $feishuProcessId"
        exit 1
    }
    
    Write-Output "Found Feishu process with ID $($feishuProcess.Id)"
    
    # Use Win32 API to bring window to foreground
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;

        public class NativeMethods
        {
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool IsIconic(IntPtr hWnd);
        }
"@
    
    # Try to restore and activate the window
    if ($feishuProcess.MainWindowHandle -ne [IntPtr]::Zero) {
        if ([NativeMethods]::IsIconic($feishuProcess.MainWindowHandle)) {
            [NativeMethods]::ShowWindow($feishuProcess.MainWindowHandle, 9)  # SW_RESTORE
        }
        
        Start-Sleep -Milliseconds 500
        [NativeMethods]::SetForegroundWindow($feishuProcess.MainWindowHandle)
    } else {
        Write-Output "MainWindowHandle is zero, trying alternative activation method..."
    }
    
    Start-Sleep -Milliseconds 2000  # Wait for window to be active

    # Send the keyboard shortcuts to open search and send message
    # Alt+Shift+K is the quick search shortcut in Feishu
    [System.Windows.Forms.SendKeys]::SendWait("%+k")  # Alt+Shift+K
    Start-Sleep -Milliseconds 1000
    
    # Input contact name using clipboard to avoid encoding issues
    Set-Clipboard -Value "Zhu Bao"
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")  # Ctrl+V to paste
    Start-Sleep -Milliseconds 1000
    
    # Press Enter to select contact
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 2000  # Wait for chat to load
    
    # Input message using clipboard to avoid encoding issues
    Set-Clipboard -Value "Ni zai gan ma ne"
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")  # Ctrl+V to paste
    Start-Sleep -Milliseconds 500
    
    # Press Enter to send message
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    Write-Output "Successfully attempted to send message to Zhu Bao"
    Write-Output "Please verify manually in Feishu that the message was sent."

} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
}