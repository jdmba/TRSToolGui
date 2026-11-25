# TRSToolGui

A front end for Lawrence Kesteloot's [trs80-tool](https://www.my-trs-80.com/tool/), in the spirit of Matthew Reed's TRSTools.

Open a TRS-80 disk image (DMK, JV1, JV3, DSK, HFE, IMD, SCP, HDV, gzipped or not) and work with it the way TRSTools did: list, view, extract, add, delete and rename files, change attributes, passwords, the disk name and the date, and create new floppy and hard disk images. trs80-tool does all the disk work; TRSToolGui asks it for everything as data (`disk --json`), so each trs80-tool update flows straight through. At startup the GUI asks trs80-tool what it can do (`help --json`) and greys out anything that version can't do. One trs80-tool is kept running in the background, so nothing waits on a program start.

It also exposes trs80-tool features TRSTools never had: a sector map, a granule map, signs of copy protection, imaging artifacts, check and repair, trim, format, hard disk partitions, boot in the emulator, disassembly and assembly. It adds two things of its own, both read-only and both reading the DMK file directly: an analysis of the image (the DOS, model and version from the disk's own bytes, copy-protection signs, damage versus imaging artifacts, and tail tracks), and a hexadecimal listing of chosen tracks (Disk > Track hex dump), sector by sector with the IDAM table decoded, for reading or for handing to something that will not take a `.dmk`.

**TRSToolGui 1.8.1 needs trs80-tool 5.1.0 or later.** With trs80-tool 5.2.0 it also changes and creates Montezuma Micro CP/M disks, lets you name the layout of a CP/M disk, opens gzipped images, and maps hard disk partitions. With trs80-tool 5.3.0 it copies files straight from one image to another (Disk > Import from another image), and a file that cannot be added no longer stops the rest of a multi-file add.

TRSToolGui runs on Windows, on Linux (x64) from version 1.3, and on the Mac from version 1.4. The Mac version is new: it was built on Windows without a Mac to try it on, so reports from Mac users are wanted.

From version 1.6 the program's own menus, buttons and dialogs can be French: Preferences has a Language list, and the choice takes effect the next time the program starts. The disk analysis reports, whatever trs80-tool itself prints, the manual and the changelog stay English. The French is a first translation and has not been through a reviewer who knows both the language and TRS-80 floppies, so a word that reads wrongly is worth [an issue](../../issues).

## Download

Get the zip for your system from [Releases](../../releases).

**Windows:** `TRSToolGui-v1.8.1.zip` holds:

| Folder | What |
|---|---|
| `With .NET 10 installed\TRSToolGui.exe` | Small. Needs the [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0). |
| `No .NET needed\TRSToolGui.exe` | Self-contained, runs on Windows 10 and later. |
| `Source\` | This repository's source. |

**Linux (x64):** `TRSToolGui-v1.8.1-linux-x64.zip` holds one self-contained file, `TRSToolGui`, with the license and the manual. Make it runnable once with `chmod +x TRSToolGui`, then run `./TRSToolGui`. It needs GTK 3, which desktop Linux has.

**Mac:** `TRSToolGui-v1.8.1-macos-arm64.zip` is for a Mac with Apple Silicon (M1 and later), `TRSToolGui-v1.8.1-macos-x64.zip` for a Mac with an Intel processor. Each holds `TRSToolGui.app`, the license and the manual. It needs macOS 14 or later. Unzip it and drag `TRSToolGui.app` to Applications before you start it. The app is not signed with an Apple developer certificate, so the first start is refused: open System Settings > Privacy & Security and click Open Anyway, or run `xattr -dr com.apple.quarantine /Applications/TRSToolGui.app` in Terminal. A trs80-tool downloaded with a browser needs `xattr -d com.apple.quarantine trs80-tool` once.

**trs80-tool is not included.** Download the one for your system from <https://www.my-trs-80.com/tool/>. On Windows, TRSToolGui asks on first start where it is and remembers the answer in `TRSToolGui.cfg` next to `TRSToolGui.exe`. On Linux it looks beside itself and on the PATH first, and keeps its settings in `~/.config/TRSToolGui/TRSToolGui.cfg`. On a Mac it looks in the folder `TRSToolGui.app` is in, on the PATH, and in `/opt/homebrew/bin` and `/usr/local/bin`, and keeps its settings in `~/Library/Application Support/TRSToolGui/TRSToolGui.cfg`.

See `Help\Manual.txt` for every menu and window (its last two sections say what differs on Linux and on the Mac), and `Help\Changelog.txt` for what changed.

## Reporting problems

Report anything that fails, gives a wrong result or looks wrong as an issue at <https://github.com/jdmba/TRSToolGui/issues>, one problem per issue. The form there asks for the two versions (Help > About shows them), your system, the steps, and the text of any error. Every report is read, and fixes are listed in `Help\Changelog.txt` with the number of the issue. If you can't tell whether the fault is in TRSToolGui or in trs80-tool, report it here.

## Building

Needs the .NET 10 SDK. Everything builds on Windows.

```
dotnet build TRSToolGui.vbproj -c Release
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release -r osx-arm64
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release -r osx-x64
```

The second line makes the Linux program; the last two make `TRSToolGui.app` for the two kinds of Mac. `release.bat` builds the two Windows executables, the Linux one, the two Mac apps, and the four download zips. The Mac zips are written by `unixzip.ps1`, which keeps the Unix permission that lets the program inside `TRSToolGui.app` run; a zip made the Windows way loses it.

The source is in three projects. `Core\TRSToolGui.Core.vbproj` is everything that is not a window: the trs80-tool adapter, the disk model, the DMK and assembler-source analysis, the settings, and the few things that differ between operating systems. It targets plain `net10.0`, references no window toolkit, and builds and runs on Windows, Linux and macOS. `TRSToolGui.vbproj` is the Windows Forms program. `Eto\TRSToolGui.Eto.vbproj` is the Linux and Mac program: the same windows written for [Eto.Forms](https://github.com/picoe/Eto), built with its GTK platform for Linux and its Mac platform for macOS. NuGet supplies both. Both programs reference Core.

## License

MIT; see [Help/License.txt](Help/License.txt). TRSToolGui is by Ira Goldklang and Claude AI.

trs80-tool is © Lawrence Kesteloot, MIT License. TRSTools is © 2009–2021 Matthew Reed; TRSToolGui is modelled on it, and the PDRIVE line of its DMK analysis is converted from his source (see the License).
