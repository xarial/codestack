SET vaultName=%1
SET filePath=%2
SET action=%3

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0get-url.ps1" %vaultName% %filePath% %action%