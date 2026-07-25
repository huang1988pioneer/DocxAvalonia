@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0app-out\Release\WordViewHost.exe" (
  dotnet build "%~dp0DocxAvalonia.csproj" -c Release -v q
)
start "" "%~dp0app-out\Release\WordViewHost.exe" %*
