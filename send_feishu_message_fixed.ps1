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
    $conditionName = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::NameProperty, "飞书"
    )
    $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $conditionName)

    if ($null -eq $feishuWindow) {
        # Try alternative window title
        $conditionName2 = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::NameProperty, "Lark"
        )
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $conditionName2)
    }

    if ($null -eq $feishuWindow) {
        # Try for English version
        $conditionName3 = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::NameProperty, "Feishu"
        )
        $feishuWindow = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $conditionName3)
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
    
    # Create conditions for search box
    $conditionCtrlType = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit
    )
    
    $conditionNameSearch1 = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::NameProperty, "搜索"
    )
    $conditionNameSearch2 = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::NameProperty, "Search"
    )
    $conditionNameSearch3 = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::NameProperty, "搜索或输入用户名"
    )
    $conditionNameSearch4 = New-Object System.Windows.Automation.PropertyCondition @(
        [System.Windows.Automation.AutomationElement]::NameProperty, "Search or enter username"
    )
    
    $orConditionNames = New-Object System.Windows.Automation.OrCondition @(
        $conditionNameSearch1, $conditionNameSearch2, $conditionNameSearch3, $conditionNameSearch4
    )
    
    $andCondition = New-Object System.Windows.Automation.AndCondition @(
        $conditionCtrlType, $orConditionNames
    )
    
    $searchBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $andCondition)

    if ($null -eq $searchBox) {
        Write-Output "Could not locate search box by name. Trying general approach..."
        
        # Use keyboard shortcut Ctrl+K which usually opens search in many apps
        [System.Windows.Forms.SendKeys]::SendWait("^k")  # Ctrl+K
        Start-Sleep -Milliseconds 500
        
        # Now try to find any edit control
        $anyEditCondition = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit
        )
        $searchBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $anyEditCondition)
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
        
        # Create conditions for finding contact
        $conditionNameContact = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::NameProperty, $Contact
        )
        
        $conditionCtrlType1 = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem
        )
        $conditionCtrlType2 = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Custom
        )
        $conditionCtrlType3 = New-Object System.Windows.Automation.PropertyCondition @(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataGrid
        )
        
        $orCtrlTypes = New-Object System.Windows.Automation.OrCondition @(
            $conditionCtrlType1, $conditionCtrlType2, $conditionCtrlType3
        )
        
        $andConditionContact = New-Object System.Windows.Automation.AndCondition @(
            $orCtrlTypes, $conditionNameContact
        )
        
        $contactItem = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $andConditionContact)
        
        if ($null -eq $contactItem) {
            # Alternative: Look for text containing the contact name
            $conditionCtrlTypeText = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text
            )
            $andConditionText = New-Object System.Windows.Automation.AndCondition @(
                $conditionCtrlTypeText, $conditionNameContact
            )
            $contactItem = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $andConditionText)
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
            
            # Create conditions for message box
            $conditionCtrlTypeEdit = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit
            )
            $conditionClassTextEditor = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::ClassNameProperty, "TextEditor"
            )
            $conditionNameInput1 = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::NameProperty, "输入消息"
            )
            $conditionNameInput2 = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::NameProperty, "Type a message"
            )
            $conditionNameInput3 = New-Object System.Windows.Automation.PropertyCondition @(
                [System.Windows.Automation.AutomationElement]::NameProperty, "输入"
            )
            
            $orNameConditions = New-Object System.Windows.Automation.OrCondition @(
                $conditionNameInput1, $conditionNameInput2, $conditionNameInput3
            )
            
            $andMsgCondition = New-Object System.Windows.Automation.AndCondition @(
                $conditionCtrlTypeEdit, $orNameConditions
            )
            
            $messageBox = $feishuWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $andMsgCondition)
            
            if ($null -eq $messageBox) {
                # Alternative: Look for any enabled edit control in the lower part of the window
                $allEditControls = $feishuWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $conditionCtrlTypeEdit)
                
                # Get window bounds to determine lower half
                $bounds = $feishuWindow.Current.BoundingRectangle
                $lowerYThreshold = $bounds.Top + ($bounds.Height / 2)
                
                foreach ($control in $allEditControls) {
                    if ($control.Current.IsEnabled) {
                        $ctrlBounds = $control.Current.BoundingRectangle
                        # Check if the control is in the lower half of the window (where message input typically is)
                        if ($ctrlBounds.Top -gt $lowerYThreshold) {
                            $messageBox = $control
                            break
                        }
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
            }
        } else {
            Write-Output "Contact $($Contact) not found in search results."
        }
    } else {
        Write-Output "Could not locate search box in the UI. The Feishu interface might be different than expected."
    }
} catch {
    Write-Output "An error occurred: $($_.Exception.Message)"
}

Write-Output "Script completed."