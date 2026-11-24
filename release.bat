rem release.bat -- build the TRSToolGui download zips. Needs the .NET 10 SDK. Everything builds on Windows.
rem Four zips are left in this folder:
rem   TRSToolGui-vX.Y.zip               Windows (described below)
rem   TRSToolGui-vX.Y-linux-x64.zip     the self-contained Linux program TRSToolGui, License.txt and Manual.txt
rem   TRSToolGui-vX.Y-macos-arm64.zip   TRSToolGui.app for a Mac with Apple Silicon, License.txt and Manual.txt
rem   TRSToolGui-vX.Y-macos-x64.zip     the same for a Mac with an Intel processor
rem The two Mac zips are written by unixzip.ps1 and not by Compress-Archive, because the program inside
rem TRSToolGui.app has to keep its Unix "may be run" permission, which a zip made the Windows way doesn't record.
rem The Windows zip holds:
rem   With .NET 10 installed\TRSToolGui.exe  small; needs the .NET 10 Desktop Runtime
rem   No .NET needed\TRSToolGui.exe          self-contained (win-x64); runs on Windows 10+
rem   License.txt
rem   Source\                                 this project's source (the Windows windows, Eto\ for the Linux and
rem                                           Mac ones, and Core\ for everything else)
rem trs80-tool is NOT bundled: on first start the GUI asks where it is and saves the
rem answer in TRSToolGui.cfg next to TRSToolGui.exe.
setlocal
cd /d "%~dp0"
set ZIPNAME=TRSToolGui-v1.8.zip
set LINUXZIP=TRSToolGui-v1.8-linux-x64.zip
set MACZIP=TRSToolGui-v1.8-macos-arm64.zip
set MACX64ZIP=TRSToolGui-v1.8-macos-x64.zip
set STAGE=%TEMP%\TRSToolGui-release
set LSTAGE=%TEMP%\TRSToolGui-release-linux
set MSTAGE=%TEMP%\TRSToolGui-release-mac
set WITHNET=%STAGE%\With .NET 10 installed
set NONET=%STAGE%\No .NET needed

if exist "%STAGE%" rd /s /q "%STAGE%"
if exist "%LSTAGE%" rd /s /q "%LSTAGE%"
if exist "%MSTAGE%" rd /s /q "%MSTAGE%"
mkdir "%LSTAGE%\zip" || exit /b 1
mkdir "%MSTAGE%\arm64" || exit /b 1
mkdir "%MSTAGE%\x64" || exit /b 1
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

rem The Mac program, once for each kind of Mac: the same project published for osx-arm64 and osx-x64.
rem The Eto.Platform.Mac64 package wraps the published file in TRSToolGui.app, in the publish folder.
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release -r osx-arm64 -p:SelfContained=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o "%MSTAGE%\pub-arm64" || exit /b 1
if not exist "%MSTAGE%\pub-arm64\TRSToolGui.app\Contents\MacOS\TRSToolGui" echo *** No TRSToolGui.app came out of the osx-arm64 publish. *** & exit /b 1
xcopy /e /i /q /y "%MSTAGE%\pub-arm64\TRSToolGui.app" "%MSTAGE%\arm64\TRSToolGui.app" >nul || exit /b 1
copy /y Help\License.txt "%MSTAGE%\arm64\" >nul || exit /b 1
copy /y Help\Manual.txt "%MSTAGE%\arm64\" >nul || exit /b 1

dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release -r osx-x64 -p:SelfContained=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o "%MSTAGE%\pub-x64" || exit /b 1
if not exist "%MSTAGE%\pub-x64\TRSToolGui.app\Contents\MacOS\TRSToolGui" echo *** No TRSToolGui.app came out of the osx-x64 publish. *** & exit /b 1
xcopy /e /i /q /y "%MSTAGE%\pub-x64\TRSToolGui.app" "%MSTAGE%\x64\TRSToolGui.app" >nul || exit /b 1
copy /y Help\License.txt "%MSTAGE%\x64\" >nul || exit /b 1
copy /y Help\Manual.txt "%MSTAGE%\x64\" >nul || exit /b 1

rd /s /q "%STAGE%\pub-fd"
rd /s /q "%STAGE%\pub-sc"
rd /s /q "%LSTAGE%\pub"
rd /s /q "%MSTAGE%\pub-arm64"
rd /s /q "%MSTAGE%\pub-x64"
rd /s /q "bin"
rd /s /q "obj"
if exist "Core\bin" rd /s /q "Core\bin"
if exist "Core\obj" rd /s /q "Core\obj"
if exist "Eto\bin" rd /s /q "Eto\bin"
if exist "Eto\obj" rd /s /q "Eto\obj"

copy /y Help\License.txt "%STAGE%\" >nul || exit /b 1
copy /y *.vb "%STAGE%\Source\" >nul || exit /b 1
copy /y TRSToolGui.vbproj "%STAGE%\Source\" >nul || exit /b 1
copy /y TRSToolGui.ico "%STAGE%\Source\" >nul || exit /b 1
copy /y Core\*.vb "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y Core\TRSToolGui.Core.vbproj "%STAGE%\Source\Core\" >nul || exit /b 1
copy /y Eto\*.vb "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y Eto\TRSToolGui.Eto.vbproj "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y Eto\Icon.icns "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y Eto\Icon.png "%STAGE%\Source\Eto\" >nul || exit /b 1
copy /y release.bat "%STAGE%\Source\" >nul || exit /b 1
copy /y unixzip.ps1 "%STAGE%\Source\" >nul || exit /b 1
copy /y Help\*.txt "%STAGE%\Source\Help\" >nul || exit /b 1

if exist "%ZIPNAME%" del "%ZIPNAME%"
powershell -NoProfile -Command "Compress-Archive -Path '%STAGE%\*' -DestinationPath '%~dp0%ZIPNAME%'" || exit /b 1
rd /s /q "%STAGE%"

if exist "%LINUXZIP%" del "%LINUXZIP%"
powershell -NoProfile -Command "Compress-Archive -Path '%LSTAGE%\zip\*' -DestinationPath '%~dp0%LINUXZIP%'" || exit /b 1
rd /s /q "%LSTAGE%"

if exist "%MACZIP%" del "%MACZIP%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0unixzip.ps1" -Source "%MSTAGE%\arm64" -Zip "%~dp0%MACZIP%" -Executable "TRSToolGui.app/Contents/MacOS/TRSToolGui" || exit /b 1
if exist "%MACX64ZIP%" del "%MACX64ZIP%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0unixzip.ps1" -Source "%MSTAGE%\x64" -Zip "%~dp0%MACX64ZIP%" -Executable "TRSToolGui.app/Contents/MacOS/TRSToolGui" || exit /b 1
rd /s /q "%MSTAGE%"

echo.
echo Done. %ZIPNAME%, %LINUXZIP%, %MACZIP% and %MACX64ZIP% are in %~dp0
