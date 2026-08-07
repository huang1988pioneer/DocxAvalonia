@echo off
setlocal
cd /d "%~dp0"
set EXE=%~dp0bin\Release\net8.0\DocxAvalonia.exe
if not exist "%EXE%" (
  dotnet build "%~dp0DocxAvalonia.csproj" -c Release -v q
)
start "" "%EXE%" %*
