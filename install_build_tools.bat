@echo off
setlocal EnableExtensions

echo This will install Visual Studio Build Tools 2022 (C++ toolchain).
echo It may download several GB and can take a while.
echo.

where winget >nul 2>nul
if errorlevel 1 (
  echo ERROR: winget not found.
  echo Please install "App Installer" from Microsoft Store, then rerun.
  echo.
  pause
  exit /b 1
)

winget install --id Microsoft.VisualStudio.2022.BuildTools -e --source winget ^
  --accept-package-agreements --accept-source-agreements --silent ^
  --override "--wait --passive --add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"

echo.
echo Done. Now run:
echo   build.bat
echo.
pause

