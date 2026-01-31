# PowerShell Script to Send Message via Feishu
# Requires UI Automation libraries

param(
    [string]$Contact = "朱宝",
    [string]$Message = "你在干嘛"
)

# Load UI Automation assemblies
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

try {
    # Find the Feishu window
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, 
        New-Object System.Windows.Automation.AndCondition @(
            New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::NameProperty, "飞书"
            )
        ))

    if ($null -eq $feishuWindow) {
        # Try alternative window title
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, 
            [System.Windows.Automation.Condition]::New([System.Windows.Automation.PropertyCondition]::New(
                [System.Windows.Automation.AutomationElement]::NameProperty, "Lark")))
    }

    if ($null -eq $feishuWindow) {
        # Try for English version
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, 
            [System.Windows.Automation.Condition]::New([System.Windows.Automation.PropertyCondition]::New(
                [System.Windows.Automation.AutomationElement]::NameProperty, "Feishu")))
    }

    if ($null -eq $feishuWindow) {
        Write-Output "Cannot find Feishu/Lark window. Please make sure it's open and visible."
        exit 1
    }

    Write-Output "Found Feishu window. Attempting to send message..."

    # Bring the window to foreground
    $windowPattern = [System.Windows.Automation.WindowPattern]$feishuWindow.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
    $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
    $feishuWindow.SetFocus()

    Start-Sleep -Milliseconds 1000

    # Find the contact search box (may be in different locations based on UI)
    $searchBox = $null
    
    # Method 1: Look for search box with placeholder text
    $searchBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.AndCondition]::New(@(
            [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit),
            [System.Windows.Automation.OrCondition]::New(@(
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "搜索"),
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "Search"),
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "搜索或输入用户名"),
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "Search or enter username")
            ))
        )))

    if ($null -eq $searchBox) {
        # Method 2: Look for search icon that might trigger a textbox
        $searchIcon = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.AndCondition]::New(@(
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button),
                [System.Windows.Automation.OrCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "搜索"),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "Search")
                ))
            )))

        if ($null -ne $searchIcon) {
            # Click the search icon to reveal the search box
            $invokePattern = [System.Windows.Automation.InvokePattern]$searchIcon.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            $invokePattern.Invoke()
            
            Start-Sleep -Milliseconds 500
            
            # Now look for the revealed search box
            $searchBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.AndCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
                )))
        }
    }

    if ($null -eq $searchBox) {
        Write-Output "Could not locate search box. Trying general approach..."
        
        # Method 3: Use keyboard shortcut Ctrl+K which usually opens search in many apps
        [System.Windows.Forms.SendKeys]::SendWait("^k")  # Ctrl+K
        Start-Sleep -Milliseconds 500
        
        # Now try to find the search box again
        $searchBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.AndCondition]::New(@(
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
            )))
    }

    if ($null -ne $searchBox) {
        # Focus the search box
        $searchBox.SetFocus()
        Start-Sleep -Milliseconds 500
        
        # Clear any existing text and type the contact name
        $textPattern = [System.Windows.Automation.TextPattern]$searchBox.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
        $textPattern.DocumentRange.SetText($Contact)
        
        Start-Sleep -Milliseconds 500
        
        # Press Enter to initiate search
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        
        Start-Sleep -Milliseconds 1000
        
        # Look for the contact in search results
        $contactItem = $null
        
        # Try finding contact by name in various UI elements
        $contactItem = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.AndCondition]::New(@(
                [System.Windows.Automation.OrCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Custom),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataGrid)
                )),
                [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, $Contact)
            )))
        
        if ($null -eq $contactItem) {
            # Alternative: Look for text containing the contact name
            $contactItem = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.AndCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, $Contact)
                )))
        }

        if ($null -ne $contactItem) {
            # Try to click on the contact
            try {
                $invokePattern = [System.Windows.Automation.InvokePattern]$contactItem.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                $invokePattern.Invoke()
            } catch {
                # If invoke fails, try selection pattern
                try {
                    $selectionItemPattern = [System.Windows.Automation.SelectionItemPattern]$contactItem.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
                    $selectionItemPattern.Select()
                } catch {
                    # Last resort: try clicking with TransformPattern
                    try {
                        $transformPattern = [System.Windows.Automation.TransformPattern]$contactItem.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern)
                        $transformPattern.Move(10, 10) # Small movement to bring to focus
                    } catch {
                        Write-Output "Could not interact with contact item directly"
                    }
                }
            }
            
            Start-Sleep -Milliseconds 1000
            
            # Find the message input box
            $messageBox = $null
            
            # Common selectors for message input areas
            $messageBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.AndCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit),
                    [System.Windows.Automation.OrCondition]::New(@(
                        [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ClassNameProperty, "TextEditor"),
                        [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "输入消息"),
                        [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "Type a message"),
                        [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::NameProperty, "输入")
                    ))
                )))
            
            if ($null -eq $messageBox) {
                # Alternative: Look for any edit control in the lower part of the window
                $bounds = $feishuWindow.Current.BoundingRectangle
                $lowerHalfCondition = [System.Windows.Automation.AndCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::IsEnabledProperty, $true)
                ))
                
                $allEditControls = $feishuWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $lowerHalfCondition)
                
                foreach ($control in $allEditControls) {
                    $ctrlBounds = $control.Current.BoundingRectangle
                    # Check if the control is in the lower half of the window (where message input typically is)
                    if ($ctrlBounds.Top -gt ($bounds.Top + ($bounds.Height / 2))) {
                        $messageBox = $control
                        break
                    }
                }
            }

            if ($null -ne $messageBox) {
                # Focus and send the message
                $messageBox.SetFocus()
                Start-Sleep -Milliseconds 500
                
                # Clear any existing text and type the message
                $textPattern = [System.Windows.Automation.TextPattern]$messageBox.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
                $textPattern.DocumentRange.SetText($Message)
                
                Start-Sleep -Milliseconds 500
                
                # Press Enter to send
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                
                Write-Output "Message sent successfully to $($Contact): $($Message)"
            } else {
                Write-Output "Could not find message input box. The UI might be different than expected."
                Write-Output "Available edit controls in window:"
                $allEditControls = $feishuWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, 
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit))
                foreach ($control in $allEditControls) {
                    Write-Output "  - Name: '$($control.Current.Name)', ClassName: '$($control.Current.ClassName)'"
                }
            }
        } else {
            Write-Output "Contact $Contact not found in search results."
            Write-Output "Available contacts in search results:"
            $contacts = $feishuWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.AndCondition]::New(@(
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem),
                    [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::IsEnabledProperty, $true)
                )))
            foreach ($item in $contacts) {
                if ($item.Current.Name -and $item.Current.Name.Length -gt 0) {
                    Write-Output "  - $($item.Current.Name)"
                }
            }
        }
    } else {
        Write-Output "Could not locate search box in the UI. The Feishu interface might be different than expected."
        Write-Output "Available windows on desktop:"
        $allWindows = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, 
            [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window))
        foreach ($window in $allWindows) {
            if ($window.Current.Name -and $window.Current.Name.Length -gt 0) {
                Write-Output "  - $($window.Current.Name) (Class: $($window.Current.ClassName))"
            }
        }
    }
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
    Write-Output "Full error details: $($_.Exception)"
}

Write-Output "Script completed."