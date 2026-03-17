@echo off
setlocal EnableExtensions

set "EXE=%~dp0build\WorkLogLite.exe"
if exist "%EXE%" (
  start "" "%EXE%"
  exit /b 0
)

echo.
echo WorkLogLite.exe not found at:
echo   %EXE%
echo.
echo If you are a user: you need a built release (an exe) to run.
echo If you are a developer: run build.bat first to compile it.
echo.
pause
exit /b 1

