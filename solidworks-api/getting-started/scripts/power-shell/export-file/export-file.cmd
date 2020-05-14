SET inputFilePath=%1
SET outFilePath=%2

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0export-file.ps1" %inputFilePath% %outFilePath%