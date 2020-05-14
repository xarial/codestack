SET vaultName=%1
SET userName=%2
SET password=%3
SET folderName=%4
SET targetUserName=%5
SET permissions=%6

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0set-permissions.ps1" %vaultName% %userName% %password% %folderName% %targetUserName% %permissions%