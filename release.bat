@echo off
rem release.bat -- build the TRSToolGui download zip. Needs the .NET 10 SDK.
rem The zip, left in this folder, holds:
rem   With .NET 10 installed\TRSToolGui.exe  small; needs the .NET 10 Desktop Runtime
rem   No .NET needed\TRSToolGui.exe          self-contained (win-x64); runs on Windows 10+
rem   License.txt
rem   Source\                                 this project's source (the windows, and Core\ for everything else)
rem trs80-tool is NOT bundled: on first start the GUI asks where it is and saves the
rem answer in TRSToolGui.cfg next to TRSToolGui.exe.
setlocal
cd /d "%~dp0"
set ZIPNAME=TRSToolGui-v1.2.zip
set STAGE=%TEMP%\TRSToolGui-release
set WITHNET=%STAGE%\With .NET 10 installed
set NONET=%STAGE%\No .NET needed

if exist "%STAGE%" rd /s /q "%STAGE%"
mkdir "%WITHNET%" || exit /b 1
mkdir "%NONET%" || exit /b 1
mkdir "%STAGE%\Source\Help" || exit /b 1
mkdir "%STAGE%\Source\Core" || exit /b 1

dotnet clean -c Release

rem Framework-dependent single file
dotnet publish TRSToolGui.vbproj -c Release -r win-x64 -p:SelfContained=false -p:PublishSingleFile=true -p:DebugType=none -o "%STAGE%\pub-fd" || exit /b 1
copy /y "%STAGE%\pub-fd\TRSToolGui.exe" "%WITHNET%\" >nul || exit /b 1

rem Self-contained single file
dotnet publish TRSToolGui.vbproj -c Release -r win-x64 -p:SelfContained=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o "%STAGE%\pub-sc" || exit /b 1
copy /y "%STAGE%\pub-sc\TRSToolGui.exe" "%NONET%\" >nul || exit /b 1

rd /s /q "%STAGE%\pub-fd"
rd /s /q "%STAGE%\pub-sc"
rd /s /q "bin"
rd /s /q "obj"
if exist "Core\bin" rd /s /q "Core\bin"
if exist "Core\obj" rd /s /q "Core\obj"

copy /y Help\License.txt "%STAGE%\" >nul || exit /b 1
copy /y *.vb "%STAGE%\Source\" >nul || exit /b 1
copy /y TRSToolGui.vbproj "%STAGE%\Source\" >nul || exit /b 1
copy /y Core\*.vb "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y Core\TRSToolGui.Core.vbproj "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y release.bat "%STAGE%\Source\" >nul || exit /b 1
copy /y Help\*.txt "%STAGE%\Source\Help\" >nul || exit /b 1

if exist "%ZIPNAME%" del "%ZIPNAME%"
powershell -NoProfile -Command "Compress-Archive -Path '%STAGE%\*' -DestinationPath '%~dp0%ZIPNAME%'" || exit /b 1
rd /s /q "%STAGE%"

echo.
echo Done. %ZIPNAME% is in %~dp0
