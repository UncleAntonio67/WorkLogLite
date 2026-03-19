@echo off
setlocal EnableExtensions

REM Build WorkLogLite into build\WorkLogLite.exe
REM Preferred: MSVC (cl.exe)
REM Fallback: bundled MinGW (w64devkit) if present at .toolchain\w64devkit\w64devkit\bin

set NOPAUSE=0
if /I "%~1"=="--no-pause" set NOPAUSE=1

set SOURCES=src\main.cpp src\storage.cpp src\report.cpp src\categories.cpp src\tasks.cpp src\recurring.cpp src\demo.cpp src\win32_util.cpp src\screenshot_tool.cpp

where cl >nul 2>nul
if not errorlevel 1 goto :build_msvc

REM Try to bootstrap MSVC environment automatically via vswhere + vcvars64.bat.
set "VSINSTALL="
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" goto :check_msvc

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSINSTALL=%%I"
if not defined VSINSTALL goto :check_msvc

if exist "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" call "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" >nul 2>nul

:check_msvc
where cl >nul 2>nul
if not errorlevel 1 goto :build_msvc

REM Fallback: bundled MinGW (w64devkit).
set "GPP=%~dp0.toolchain\w64devkit\w64devkit\bin\g++.exe"
if exist "%GPP%" goto :build_mingw

echo.
echo ERROR: No C++ toolchain found (cl.exe or bundled MinGW).
echo.
echo Fix option 1 (recommended):
echo   Install "Visual Studio Build Tools 2022" with workload "Desktop development with C++".
echo.
echo Fix option 2:
echo   Download / prepare .toolchain\w64devkit in this repo, then rerun build.bat.
echo.
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
  /W4 /wd4100 /wd4505 ^
  /Fe:build\WorkLogLite.new.exe ^
  %SOURCES% ^
  user32.lib gdi32.lib comctl32.lib comdlg32.lib shlwapi.lib ole32.lib shell32.lib crypt32.lib bcrypt.lib

if errorlevel 1 goto :build_fail
goto :build_done

:build_mingw
if not exist build mkdir build

set "PATH=%~dp0.toolchain\w64devkit\w64devkit\bin;%PATH%"

echo.
echo Using bundled MinGW toolchain:
echo   %GPP%
echo.

"%GPP%" ^
  -std=c++17 -O2 -s ^
  -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -D_CRT_SECURE_NO_WARNINGS ^
  -DWINVER=0x0601 -D_WIN32_WINNT=0x0601 -D_WIN32_IE=0x0600 ^
  -municode -mwindows ^
  -static -static-libgcc -static-libstdc++ ^
  -o build\WorkLogLite.new.exe ^
  %SOURCES% ^
  -luser32 -lgdi32 -lcomctl32 -lcomdlg32 -lshlwapi -lole32 -lshell32 -lcrypt32 -lbcrypt -luuid

if errorlevel 1 goto :build_fail
goto :build_done

:build_done
echo.
if exist build\WorkLogLite.exe (
  del /f /q build\WorkLogLite.exe >nul 2>nul
)
move /y build\WorkLogLite.new.exe build\WorkLogLite.exe >nul 2>nul

if exist build\WorkLogLite.exe (
  echo Build OK: build\WorkLogLite.exe
) else (
  echo Build OK: build\WorkLogLite.new.exe
  echo Note: build\WorkLogLite.exe is probably running/locked. Close it and rename manually.
)
if "%NOPAUSE%"=="0" pause
exit /b 0

:build_fail
echo.
echo Build failed.
if "%NOPAUSE%"=="0" pause
exit /b 1

