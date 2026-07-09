<#
.Synopsis
Creates a Desktop shortcut for launching Exchange Recipient Admin Center.
.Description
Run this once from wherever this project's files actually live. It resolves
its own folder and the current user's Desktop dynamically, so the same
script works unmodified for any user/location - no hardcoded paths.
#>

$targetFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath = Join-Path $targetFolder "Launch-ExchangeRecipientAdmin.bat"
$iconPath = Join-Path $targetFolder "images\favicon.ico"
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Exchange Recipient Admin Center.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $batPath
$shortcut.WorkingDirectory = $targetFolder
if (Test-Path $iconPath) {
    $shortcut.IconLocation = $iconPath
}
else {
    $shortcut.IconLocation = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe,0"
}
$shortcut.Description = "Exchange Recipient Admin Center"
$shortcut.Save()

Write-Host "Shortcut created at: $shortcutPath"
