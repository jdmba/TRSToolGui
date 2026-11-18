@echo off
rem release.bat -- build the TRSToolGui download zips. Needs the .NET 10 SDK.
rem Two zips are left in this folder: the Windows one and the Linux one (TRSToolGui-vX.Y-linux-x64.zip,
rem holding the self-contained Linux program TRSToolGui, License.txt and Manual.txt).
rem The Windows zip holds:
rem   With .NET 10 installed\TRSToolGui.exe  small; needs the .NET 10 Desktop Runtime
rem   No .NET needed\TRSToolGui.exe          self-contained (win-x64); runs on Windows 10+
rem   License.txt
rem   Source\                                 this project's source (the Windows windows, Eto\ for the Linux ones,
rem                                           and Core\ for everything else)
rem trs80-tool is NOT bundled: on first start the GUI asks where it is and saves the
rem answer in TRSToolGui.cfg next to TRSToolGui.exe.
setlocal
cd /d "%~dp0"
set ZIPNAME=TRSToolGui-v1.3.zip
set LINUXZIP=TRSToolGui-v1.3-linux-x64.zip
set STAGE=%TEMP%\TRSToolGui-release
set LSTAGE=%TEMP%\TRSToolGui-release-linux
set WITHNET=%STAGE%\With .NET 10 installed
set NONET=%STAGE%\No .NET needed

if exist "%STAGE%" rd /s /q "%STAGE%"
if exist "%LSTAGE%" rd /s /q "%LSTAGE%"
mkdir "%LSTAGE%\zip" || exit /b 1
mkdir "%WITHNET%" || exit /b 1
mkdir "%NONET%" || exit /b 1
mkdir "%STAGE%\Source\Help" || exit /b 1
mkdir "%STAGE%\Source\Core" || exit /b 1
mkdir "%STAGE%\Source\Eto" || exit /b 1

dotnet clean -c Release

rem Framework-dependent single file
dotnet publish TRSToolGui.vbproj -c Release -r win-x64 -p:SelfContained=false -p:PublishSingleFile=true -p:DebugType=none -o "%STAGE%\pub-fd" || exit /b 1
copy /y "%STAGE%\pub-fd\TRSToolGui.exe" "%WITHNET%\" >nul || exit /b 1

rem Self-contained single file
dotnet publish TRSToolGui.vbproj -c Release -r win-x64 -p:SelfContained=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o "%STAGE%\pub-sc" || exit /b 1
copy /y "%STAGE%\pub-sc\TRSToolGui.exe" "%NONET%\" >nul || exit /b 1

rem The Linux program: self-contained single file for linux-x64 (Eto.Forms and GtkSharp come from NuGet)
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release -r linux-x64 -p:SelfContained=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o "%LSTAGE%\pub" || exit /b 1
copy /y "%LSTAGE%\pub\TRSToolGui" "%LSTAGE%\zip\" >nul || exit /b 1
copy /y Help\License.txt "%LSTAGE%\zip\" >nul || exit /b 1
copy /y Help\Manual.txt "%LSTAGE%\zip\" >nul || exit /b 1

rd /s /q "%STAGE%\pub-fd"
rd /s /q "%STAGE%\pub-sc"
rd /s /q "%LSTAGE%\pub"
rd /s /q "bin"
rd /s /q "obj"
if exist "Core\bin" rd /s /q "Core\bin"
if exist "Core\obj" rd /s /q "Core\obj"
if exist "Eto\bin" rd /s /q "Eto\bin"
if exist "Eto\obj" rd /s /q "Eto\obj"

copy /y Help\License.txt "%STAGE%\" >nul || exit /b 1
copy /y *.vb "%STAGE%\Source\" >nul || exit /b 1
copy /y TRSToolGui.vbproj "%STAGE%\Source\" >nul || exit /b 1
copy /y Core\*.vb "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y Core\TRSToolGui.Core.vbproj "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y Eto\*.vb "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y Eto\TRSToolGui.Eto.vbproj "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y release.bat "%STAGE%\Source\" >nul || exit /b 1
copy /y Help\*.txt "%STAGE%\Source\Help\" >nul || exit /b 1

if exist "%ZIPNAME%" del "%ZIPNAME%"
powershell -NoProfile -Command "Compress-Archive -Path '%STAGE%\*' -DestinationPath '%~dp0%ZIPNAME%'" || exit /b 1
rd /s /q "%STAGE%"

if exist "%LINUXZIP%" del "%LINUXZIP%"
powershell -NoProfile -Command "Compress-Archive -Path '%LSTAGE%\zip\*' -DestinationPath '%~dp0%LINUXZIP%'" || exit /b 1
rd /s /q "%LSTAGE%"

echo.
echo Done. %ZIPNAME% and %LINUXZIP% are in %~dp0
