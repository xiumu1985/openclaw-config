@echo off
chcp 65001 >nul
echo 正在激活飞书并发送消息...

powershell -ExecutionPolicy Bypass -Command ^
"[System.Threading.Thread]::CurrentThread.CurrentCulture = 'zh-CN'; ^
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'zh-CN'; ^
$feishuProc = Get-Process -Name 'Feishu' -Id 1668 -ErrorAction SilentlyContinue; ^
if ($feishuProc) { ^
  Add-Type -AssemblyName System.Windows.Forms; ^
  $signature = @' ^
[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); ^
[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr hWnd); ^
[DllImport(\"user32.dll\")] public static extern bool IsIconic(IntPtr hWnd); ^
'@; ^
  $user32 = Add-Type -MemberDefinition $signature -Name 'User32' -PassThru; ^
  if ($user32::IsIconic($feishuProc.MainWindowHandle)) { ^
    $user32::ShowWindow($feishuProc.MainWindowHandle, 9); ^
  } ^
  $user32::SetForegroundWindow($feishuProc.MainWindowHandle); ^
  Start-Sleep -Seconds 2; ^
  Set-Clipboard -Value '朱宝'; ^
  Start-Sleep -Milliseconds 200; ^
  [System.Windows.Forms.SendKeys]::SendWait('^k'); ^
  Start-Sleep -Milliseconds 500; ^
  [System.Windows.Forms.SendKeys]::SendWait('^v'); ^
  Start-Sleep -Milliseconds 1000; ^
  [System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); ^
  Start-Sleep -Seconds 2; ^
  Set-Clipboard -Value '床前明月光，疑是地上霜。举头望明月，低头思故乡。'; ^
  Start-Sleep -Milliseconds 200; ^
  [System.Windows.Forms.SendKeys]::SendWait('^v'); ^
  Start-Sleep -Milliseconds 500; ^
  [System.Windows.Forms.SendKeys]::SendWait('{ENTER}'); ^
  Write-Output '消息已发送'; ^
} else { ^
  Write-Output '未找到飞书进程'; ^
}"