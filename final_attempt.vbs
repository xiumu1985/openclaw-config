' VBScript to automate Feishu
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Wait for a moment to ensure Feishu is loaded
WScript.Sleep 3000

' Activate Feishu window using Alt+Tab or by searching
WshShell.Run "mshta vbscript:Execute(""CreateObject(""""Shell.Application"""").MinimizeAll:close"")", 0
WScript.Sleep 1000

' Press Win+S to open search
WshShell.SendKeys "^%{ESC}"
WScript.Sleep 1000

' Search for Feishu
WshShell.SendKeys "Feishu"
WScript.Sleep 1000

' Press Enter to open Feishu
WshShell.SendKeys "{ENTER}"
WScript.Sleep 3000

' Now perform the sequence to send the message
' 1. Press Ctrl+K to search for contact
WshShell.SendKeys "^k"
WScript.Sleep 1000

' 2. Type the contact name "朱宝"
WshShell.SendKeys "朱宝"
WScript.Sleep 1000

' 3. Press Enter to select the contact
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000

' 4. Type the poem
WshShell.SendKeys "床前明月光，疑是地上霜。举头望明月，低头思故乡。"
WScript.Sleep 500

' 5. Press Enter to send the message
WshShell.SendKeys "{ENTER}"

WScript.Echo "Message sent to Zhu Bao"