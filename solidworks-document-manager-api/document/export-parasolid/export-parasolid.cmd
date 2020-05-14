SET inputFilePath=%1
SET outDirPath=%2

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0export-parasolid.ps1" %inputFilePath% %outDirPath%