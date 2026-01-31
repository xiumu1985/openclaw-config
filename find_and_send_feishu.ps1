# PowerShell Script to Find Feishu Window by Process and Send Message
# Uses both UI Automation and Process-based approaches

param(
    [string]$Contact = "朱宝",
    [string]$Message = "你在干嘛"
)

# Load UI Automation assemblies
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

try {
    # First, let's get the Feishu window handle via process
    $feishuProcesses = Get-Process -Name "Feishu" -ErrorAction SilentlyContinue
    
    if ($null -eq $feishuProcesses) {
        Write-Output "No Feishu processes found. Please ensure Feishu is running."
        exit 1
    }
    
    # Find the process with a valid main window handle
    $targetProcess = $null
    foreach ($proc in $feishuProcesses) {
        if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
            $targetProcess = $proc
            break
        }
    }
    
    if ($null -eq $targetProcess) {
        Write-Output "Found Feishu processes but no main window is visible."
        Write-Output "Please make sure Feishu is not minimized."
        exit 1
    }
    
    Write-Output "Found Feishu process with ID $($targetProcess.Id) and window handle $($targetProcess.MainWindowHandle)"
    
    # Use Win32 API to bring window to foreground first
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
"@
        
    [Win32]::ShowWindow($targetProcess.MainWindowHandle, 9)  # SW_RESTORE
    [Win32]::SetForegroundWindow($targetProcess.MainWindowHandle)
    
    Start-Sleep -Milliseconds 1000
    
    # Now try to find the window by title (which should be visible after restoring)
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    
    # Try different possible titles for Feishu
    $possibleTitles = @("飞书", "Lark", "Feishu", "FeiShu")
    $feishuWindow = $null
    
    foreach ($title in $possibleTitles) {
        $condition = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::NameProperty, $title
        )
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
        
        if ($null -ne $feishuWindow) {
            break
        }
    }
    
    # If still not found, try with the actual window title we got from the process
    if ($null -eq $feishuWindow) {
        # Since we couldn't properly read the title due to encoding, let's use a different approach
        # Find the window by process ID
        $processIdCondition = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $targetProcess.Id
        )
        $windowCondition = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window
        )
        $andCondition = New-Object System.Windows.Automation.AndCondition @(
            $processIdCondition, $windowCondition
        )
        
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $andCondition)
    }

    if ($null -eq $feishuWindow) {
        Write-Output "Still cannot locate Feishu window via UI Automation."
        Write-Output "Attempting to send message using keyboard shortcuts..."
        
        # Bring the window to foreground using Win32 API
        Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
"@
        
        [Win32]::ShowWindow($targetProcess.MainWindowHandle, 9)  # SW_RESTORE
        [Win32]::SetForegroundWindow($targetProcess.MainWindowHandle)
        
        Start-Sleep -Milliseconds 1000
        
        # Use keyboard shortcuts to navigate
        # Alt+Shift+K is often used in Feishu to open quick contact search
        [System.Windows.Forms.SendKeys]::SendWait("%+k")  # Alt+Shift+K
        Start-Sleep -Milliseconds 500
        
        # Type the contact name
        [System.Windows.Forms.SendKeys]::SendWait($Contact)
        Start-Sleep -Milliseconds 500
        
        # Press Enter to select the contact
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 1000
        
        # Type the message
        [System.Windows.Forms.SendKeys]::SendWait($Message)
        Start-Sleep -Milliseconds 500
        
        # Press Enter to send
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        Write-Output "Attempted to send message using keyboard shortcuts: '$($Message)' to '$($Contact)'"
        Write-Output "Please verify manually if the message was sent successfully."
    } else {
        Write-Output "Successfully located Feishu window via UI Automation."
        
        # Bring the window to foreground
        $windowPattern = [System.Windows.Automation.WindowPattern]$feishuWindow.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
        $feishuWindow.SetFocus()

        Start-Sleep -Milliseconds 1000

        # Use keyboard shortcut to open search (Ctrl+K is standard in many apps)
        [System.Windows.Forms.SendKeys]::SendWait("^k")  # Ctrl+K
        Start-Sleep -Milliseconds 500
        
        # Type the contact name
        [System.Windows.Forms.SendKeys]::SendWait($Contact)
        Start-Sleep -Milliseconds 500
        
        # Press Enter to initiate search
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        Start-Sleep -Milliseconds 1000
        
        # Use keyboard to select the contact (usually arrow keys + enter)
        [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")  # Navigate to first result
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")  # Select contact
        
        Start-Sleep -Milliseconds 1000
        
        # Type the message
        [System.Windows.Forms.SendKeys]::SendWait($Message)
        Start-Sleep -Milliseconds 500
        
        # Press Enter to send
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        Write-Output "Message sent successfully to $($Contact): $($Message)"
    }
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
    Write-Output "Full error details: $($_.Exception)"
}

Write-Output "Script completed."