@echo off
setlocal EnableExtensions

set NOPAUSE=0
if /I "%~1"=="--no-pause" set NOPAUSE=1

set SOURCES=tests\smoke_tests.cpp src\storage.cpp src\report.cpp src\categories.cpp src\tasks.cpp src\demo.cpp src\win32_util.cpp src\data_ops.cpp

where cl >nul 2>nul
if not errorlevel 1 goto :build_msvc

set "VSINSTALL="
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" goto :check_msvc

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSINSTALL=%%I"
if not defined VSINSTALL goto :check_msvc

if exist "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" call "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" >nul 2>nul

:check_msvc
where cl >nul 2>nul
if not errorlevel 1 goto :build_msvc

set "GPP=%~dp0.toolchain\w64devkit\w64devkit\bin\g++.exe"
if exist "%GPP%" goto :build_mingw

echo.
echo ERROR: No C++ toolchain found ^(cl.exe or bundled MinGW^).
if "%NOPAUSE%"=="0" pause
exit /b 1

:build_msvc
if not exist build mkdir build

cl ^
  /utf-8 ^
  /nologo /std:c++17 /EHsc ^
  /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /D_CRT_SECURE_NO_WARNINGS ^
  /DWINVER=0x0601 /D_WIN32_WINNT=0x0601 /D_WIN32_IE=0x0600 ^
  /O2 /MT ^
  /I src ^
  /Fe:build\WorkLogLiteSmokeTests.exe ^
  %SOURCES% ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shlwapi.lib ole32.lib shell32.lib crypt32.lib bcrypt.lib

if errorlevel 1 goto :fail
goto :run

:build_mingw
if not exist build mkdir build
set "PATH=%~dp0.toolchain\w64devkit\w64devkit\bin;%PATH%"

echo.
echo Using bundled MinGW toolchain:
echo   %GPP%
echo.

"%GPP%" ^
  -std=c++17 -O2 ^
  -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -D_CRT_SECURE_NO_WARNINGS ^
  -DWINVER=0x0601 -D_WIN32_WINNT=0x0601 -D_WIN32_IE=0x0600 ^
  -municode ^
  -I src ^
  -o build\WorkLogLiteSmokeTests.exe ^
  %SOURCES% ^
  -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshlwapi -lole32 -lshell32 -lcrypt32 -lbcrypt -luuid

if errorlevel 1 goto :fail
goto :run

:run
echo.
echo Running smoke tests...
build\WorkLogLiteSmokeTests.exe
set RC=%ERRORLEVEL%
if not "%RC%"=="0" goto :fail_run

echo.
echo Smoke tests OK.
if "%NOPAUSE%"=="0" pause
exit /b 0

:fail_run
echo.
echo Smoke tests failed.
if "%NOPAUSE%"=="0" pause
exit /b %RC%

:fail
echo.
echo Test build failed.
if "%NOPAUSE%"=="0" pause
exit /b 1
