# TRSToolGui

A Windows front end for Lawrence Kesteloot's [trs80-tool](https://www.my-trs-80.com/tool/), in the spirit of Matthew Reed's TRSTools.

Open a TRS-80 disk image (DMK, JV1, JV3, DSK, HFE, IMD, SCP, HDV) and work with it the way TRSTools did: list, view, extract, add, delete and rename files, change attributes, passwords, the disk name and the date, and create new floppy and hard disk images. trs80-tool does all the disk work; TRSToolGui asks it for everything as data (`disk --json`), so each trs80-tool update flows straight through. At startup the GUI asks trs80-tool what it can do (`help --json`) and greys out anything that version can't do. One trs80-tool is kept running in the background, so nothing waits on a program start.

It also exposes trs80-tool features TRSTools never had: a sector map, a granule map, signs of copy protection, imaging artifacts, check and repair, trim, format, hard disk partitions, boot in the emulator, disassembly and assembly. It adds its own read-only analysis of DMK images: the DOS, model and version from the disk's own bytes, copy-protection signs, damage versus imaging artifacts, and tail tracks.

**TRSToolGui 1.1 needs trs80-tool 5.1.0 or later.**

## Download

Get the zip from [Releases](../../releases). It holds:

| Folder | What |
|---|---|
| `With .NET 10 installed\TRSToolGui.exe` | Small. Needs the [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0). |
| `No .NET needed\TRSToolGui.exe` | Self-contained, runs on Windows 10 and later. |
| `Source\` | This repository's source. |

**trs80-tool is not included.** Download it from <https://www.my-trs-80.com/tool/>. On first start TRSToolGui asks where it is and remembers the answer in `TRSToolGui.cfg` next to `TRSToolGui.exe`.

See `Help\Manual.txt` for every menu and window, and `Help\Changelog.txt` for what changed.

## Building

Needs the .NET 10 SDK on Windows.

```
dotnet build TRSToolGui.vbproj -c Release
```

`release.bat` builds both executables and the download zip.

## License

MIT; see [Help/License.txt](Help/License.txt). TRSToolGui is by Ira Goldklang and Claude AI.

trs80-tool is © Lawrence Kesteloot, MIT License. TRSTools is © 2009–2021 Matthew Reed; TRSToolGui is modelled on it, and the PDRIVE line of its DMK analysis is converted from his source (see the License).
