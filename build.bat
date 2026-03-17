@echo off
setlocal EnableExtensions

REM Build script for MSVC Developer Command Prompt
REM Output: build\WorkLogLite.exe

set NOPAUSE=0
if /I "%~1"=="--no-pause" set NOPAUSE=1

where cl >nul 2>nul
if not errorlevel 1 goto :have_cl

REM Try to bootstrap MSVC environment automatically via vswhere + vcvars64.bat.
set "VSINSTALL="
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" goto :check_cl

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSINSTALL=%%I"
if not defined VSINSTALL goto :check_cl

if exist "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" call "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" >nul 2>nul

:check_cl
where cl >nul 2>nul
if not errorlevel 1 goto :have_cl

echo.
echo ERROR: MSVC compiler cl.exe not found.
echo.
echo Fix:
echo 1. Install "Visual Studio Build Tools 2022" (or Visual Studio 2022).
echo 2. In installer, select workload: "Desktop development with C++".
echo    Ensure components include:
echo    - MSVC v143 x64/x86 build tools
echo    - Windows 10 or 11 SDK
echo 3. Then rerun this script (double click is OK), or run in:
echo    "x64 Native Tools Command Prompt for VS 2022"
echo.
echo Tip (winget, admin may be required):
echo   winget install --id Microsoft.VisualStudio.2022.BuildTools -e --source winget ^
echo     --accept-package-agreements --accept-source-agreements --silent ^
echo     --override "--wait --passive --add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"
echo.
if "%NOPAUSE%"=="0" pause
exit /b 1

:have_cl

if not exist build mkdir build

set SOURCES=src\main.cpp src\storage.cpp src\report.cpp src\categories.cpp src\tasks.cpp src\recurring.cpp src\demo.cpp src\win32_util.cpp

cl ^
  /utf-8 ^
  /nologo /std:c++17 /EHsc ^
  /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /D_CRT_SECURE_NO_WARNINGS ^
  /DWINVER=0x0601 /D_WIN32_WINNT=0x0601 /D_WIN32_IE=0x0600 ^
  /O2 /MT ^
  /W4 /wd4100 /wd4505 ^
  /Fe:build\WorkLogLite.new.exe ^
  %SOURCES% ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shlwapi.lib ole32.lib shell32.lib crypt32.lib

if errorlevel 1 (
  echo.
  echo Build failed.
  if "%NOPAUSE%"=="0" pause
  exit /b 1
)

echo.
if exist build\\WorkLogLite.exe (
  del /f /q build\\WorkLogLite.exe >nul 2>nul
)
move /y build\\WorkLogLite.new.exe build\\WorkLogLite.exe >nul 2>nul

if exist build\\WorkLogLite.exe (
  echo Build OK: build\\WorkLogLite.exe
) else (
  echo Build OK: build\\WorkLogLite.new.exe
  echo Note: build\\WorkLogLite.exe is probably running/locked. Close it and rename manually.
)
if "%NOPAUSE%"=="0" pause
