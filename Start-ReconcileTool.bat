@echo off
if exist "%~dp0Start-ReconcileTool.exe" (
  start "" "%~dp0Start-ReconcileTool.exe"
) else (
  start "" wscript.exe "%~dp0Start-ReconcileTool.vbs"
)
exit /b
