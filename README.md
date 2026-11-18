# TRSToolGui

A front end for Lawrence Kesteloot's [trs80-tool](https://www.my-trs-80.com/tool/), in the spirit of Matthew Reed's TRSTools.

Open a TRS-80 disk image (DMK, JV1, JV3, DSK, HFE, IMD, SCP, HDV, gzipped or not) and work with it the way TRSTools did: list, view, extract, add, delete and rename files, change attributes, passwords, the disk name and the date, and create new floppy and hard disk images. trs80-tool does all the disk work; TRSToolGui asks it for everything as data (`disk --json`), so each trs80-tool update flows straight through. At startup the GUI asks trs80-tool what it can do (`help --json`) and greys out anything that version can't do. One trs80-tool is kept running in the background, so nothing waits on a program start.

It also exposes trs80-tool features TRSTools never had: a sector map, a granule map, signs of copy protection, imaging artifacts, check and repair, trim, format, hard disk partitions, boot in the emulator, disassembly and assembly. It adds its own read-only analysis of DMK images: the DOS, model and version from the disk's own bytes, copy-protection signs, damage versus imaging artifacts, and tail tracks.

**TRSToolGui 1.3 needs trs80-tool 5.1.0 or later.** With trs80-tool 5.2.0 it also changes and creates Montezuma Micro CP/M disks, lets you name the layout of a CP/M disk, opens gzipped images, and maps hard disk partitions.

TRSToolGui runs on Windows and, from version 1.3, on Linux (x64). A Mac version is being looked into.

## Download

Get the zip for your system from [Releases](../../releases).

**Windows:** `TRSToolGui-v1.3.zip` holds:

| Folder | What |
|---|---|
| `With .NET 10 installed\TRSToolGui.exe` | Small. Needs the [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0). |
| `No .NET needed\TRSToolGui.exe` | Self-contained, runs on Windows 10 and later. |
| `Source\` | This repository's source. |

**Linux (x64):** `TRSToolGui-v1.3-linux-x64.zip` holds one self-contained file, `TRSToolGui`, with the license and the manual. Make it runnable once with `chmod +x TRSToolGui`, then run `./TRSToolGui`. It needs GTK 3, which desktop Linux has.

**trs80-tool is not included.** Download the one for your system from <https://www.my-trs-80.com/tool/>. On Windows, TRSToolGui asks on first start where it is and remembers the answer in `TRSToolGui.cfg` next to `TRSToolGui.exe`. On Linux it looks beside itself and on the PATH first, and keeps its settings in `~/.config/TRSToolGui/TRSToolGui.cfg`.

See `Help\Manual.txt` for every menu and window (its last section says what differs on Linux), and `Help\Changelog.txt` for what changed.

## Building

Needs the .NET 10 SDK. Both programs build on Windows.

```
dotnet build TRSToolGui.vbproj -c Release
dotnet publish Eto\TRSToolGui.Eto.vbproj -c Release
```

`release.bat` builds the two Windows executables, the Linux one, and both download zips.

The source is in three projects. `Core\TRSToolGui.Core.vbproj` is everything that is not a window: the trs80-tool adapter, the disk model, the DMK and assembler-source analysis, the settings, and the few things that differ between operating systems. It targets plain `net10.0`, references no window toolkit, and builds and runs on Windows, Linux and macOS. `TRSToolGui.vbproj` is the Windows Forms program. `Eto\TRSToolGui.Eto.vbproj` is the Linux program: the same windows written for [Eto.Forms](https://github.com/picoe/Eto) with its GTK platform, which NuGet supplies. Both reference Core.

## License

MIT; see [Help/License.txt](Help/License.txt). TRSToolGui is by Ira Goldklang and Claude AI.

trs80-tool is © Lawrence Kesteloot, MIT License. TRSTools is © 2009–2021 Matthew Reed; TRSToolGui is modelled on it, and the PDRIVE line of its DMK analysis is converted from his source (see the License).
