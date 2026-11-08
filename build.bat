@echo off
setlocal
cd /d "%~dp0"

where dotnet >nul 2>nul
if errorlevel 1 (
    echo [ERROR] .NET SDK not found. Please install the .NET 10 SDK.
    exit /b 1
)

echo Building excel2img...
dotnet publish -c Release -o bin
if errorlevel 1 (
    echo [ERROR] Build failed.
    exit /b 1
)

if exist "bin\Release" rmdir /s /q "bin\Release"

echo.
echo Build succeeded: %~dp0bin\excel2img.exe
endlocal
exit /b 0
