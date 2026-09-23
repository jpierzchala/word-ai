#Requires -Version 7.0
$ErrorActionPreference = 'Stop'
$tokenPath = Join-Path $env:LOCALAPPDATA 'WordAiSecure\secrets\pairing-token'
# Run manually immediately before connecting the Word add-in. Does not print it.
Set-Clipboard -Value ([IO.File]::ReadAllText($tokenPath).Trim())
Write-Output 'Token copied. Paste it into Word AI Secure Live, then clear the clipboard.'
