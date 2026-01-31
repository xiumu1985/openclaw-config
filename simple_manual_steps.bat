@echo off
echo 步骤1: 确保飞书应用已启动...
timeout /t 2 /nobreak >nul

echo 步骤2: 激活飞书窗口...
powershell -Command "Get-Process Feishu -ErrorAction SilentlyContinue | ForEach-Object { $_.MainWindowHandle; Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.Interaction]::AppActivate($_.MainWindowHandle) }"
timeout /t 2 /nobreak >nul

echo 步骤3: 设置剪贴板内容为联系人"朱宝"...
powershell -Command "Set-Clipboard -Value '朱宝'"
timeout /t 1 /nobreak >nul

echo 步骤4: 模拟键盘操作 - Ctrl+K 搜索...
powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; [System.Windows.Forms.SendKeys]::SendWait('^k'); Start-Sleep -Milliseconds 500"
timeout /t 1 /nobreak >nul

echo 步骤5: 模拟键盘操作 - Ctrl+V 粘贴联系人名称...
powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; [System.Windows.Forms.SendKeys]::SendWait('^v'); Start-Sleep -Milliseconds 500"
timeout /t 1 /nobreak >nul

echo 步骤6: 模拟键盘操作 - Enter 选择联系人...
powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; [System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); Start-Sleep -Seconds 2"
timeout /t 3 /nobreak >nul

echo 步骤7: 设置剪贴板内容为诗句...
powershell -Command "Set-Clipboard -Value '床前明月光，疑是地上霜。举头望明月，低头思故乡。'"
timeout /t 1 /nobreak >nul

echo 步骤8: 模拟键盘操作 - Ctrl+V 粘贴诗句...
powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; [System.Windows.Forms.SendKeys]::SendWait('^v'); Start-Sleep -Milliseconds 500"
timeout /t 1 /nobreak >nul

echo 步骤9: 模拟键盘操作 - Enter 发送消息...
powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')"

echo 操作完成！已尝试向朱宝发送《静夜思》
pause