# unixzip.ps1 -- write a zip whose entries carry Unix permissions. Used by release.bat for the Mac downloads.
#
# A Mac or Linux program has to keep its "may be run" permission when it is unzipped. The zips that
# Windows makes (Compress-Archive, Explorer) record no Unix permissions, so a program taken out of one
# does not start until the user has run chmod on it. Inside a Mac application (TRSToolGui.app) that
# shows as "the application can't be opened", with nothing to say why.
#
# This script writes the zip itself: every folder as rwxr-xr-x, every file as rw-r--r--, the files named
# with -Executable as rwxr-xr-x, and every entry marked as made on Unix, which is what makes an unzip
# program believe the permissions.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File unixzip.ps1 -Source <folder> -Zip <file.zip> -Executable <path>[,<path>...]
# The zip holds what is inside <folder>, not the folder itself. Each -Executable path is given as it will
# be inside the zip, with forward slashes: TRSToolGui.app/Contents/MacOS/TRSToolGui
# It runs in Windows PowerShell 5.1 and in PowerShell 7.
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Zip,
    [string[]]$Executable = @()
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# The Unix mode of an entry goes in the high 16 bits of its "external attributes".
# 040755 folder (with the DOS folder bit), 100644 file, 100755 program; written as signed 32-bit numbers.
$DirAttributes = 0x41ED0010
$FileAttributes = -2119958528
$ProgramAttributes = -2115174400

function Write-Zip([string]$root, [string]$zipPath, [hashtable]$programs) {
    $stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::Create)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($item in (Get-ChildItem -LiteralPath $root -Recurse -Force | Sort-Object FullName)) {
                if (-not $item.FullName.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { throw "unixzip: $($item.FullName) is not under $root." }
                $name = $item.FullName.Substring($root.Length + 1).Replace('\', '/')
                if ($item.PSIsContainer) {
                    $entry = $archive.CreateEntry($name + '/')
                    $entry.ExternalAttributes = $DirAttributes
                    continue
                }
                $entry = $archive.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $entry.LastWriteTime = [System.DateTimeOffset]$item.LastWriteTime
                if ($programs.ContainsKey($name)) {
                    $entry.ExternalAttributes = $ProgramAttributes
                    $programs[$name] = $true
                }
                else {
                    $entry.ExternalAttributes = $FileAttributes
                }
                $from = [System.IO.File]::OpenRead($item.FullName)
                try {
                    $to = $entry.Open()
                    try { $from.CopyTo($to) } finally { $to.Dispose() }
                }
                finally { $from.Dispose() }
            }
        }
        finally { $archive.Dispose() }
    }
    finally { $stream.Dispose() }
}

# Windows writes "made on MS-DOS" into every entry of the zip's directory, and an unzip program then
# ignores the Unix mode. The byte that says so is the second byte of "version made by": 3 means Unix.
function Set-MadeOnUnix([string]$zipPath) {
    $bytes = [System.IO.File]::ReadAllBytes($zipPath)
    $end = -1
    for ($i = $bytes.Length - 22; $i -ge 0; $i--) {
        if ($bytes[$i] -eq 0x50 -and $bytes[$i + 1] -eq 0x4B -and $bytes[$i + 2] -eq 0x05 -and $bytes[$i + 3] -eq 0x06) { $end = $i; break }
    }
    if ($end -lt 0) { throw "unixzip: $zipPath has no end-of-directory record." }
    $count = [System.BitConverter]::ToUInt16($bytes, $end + 10)
    $pos = [int][System.BitConverter]::ToUInt32($bytes, $end + 16)
    for ($n = 0; $n -lt $count; $n++) {
        if ($bytes[$pos] -ne 0x50 -or $bytes[$pos + 1] -ne 0x4B -or $bytes[$pos + 2] -ne 0x01 -or $bytes[$pos + 3] -ne 0x02) {
            throw "unixzip: entry $n of the directory of $zipPath is not where it should be."
        }
        $bytes[$pos + 5] = 3
        $pos += 46 + [System.BitConverter]::ToUInt16($bytes, $pos + 28) + [System.BitConverter]::ToUInt16($bytes, $pos + 30) + [System.BitConverter]::ToUInt16($bytes, $pos + 32)
    }
    [System.IO.File]::WriteAllBytes($zipPath, $bytes)
    return $count
}

# ---- main
$root = (Get-Item -LiteralPath $Source).FullName.TrimEnd('\', '/')
$zipPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Zip)
$programs = @{}
foreach ($given in $Executable) {
    foreach ($one in ($given -split ',')) {
        $name = $one.Trim().Replace('\', '/').Trim('/')
        if ($name.Length -gt 0) { $programs[$name] = $false }
    }
}

Write-Zip $root $zipPath $programs
$entries = Set-MadeOnUnix $zipPath

foreach ($name in @($programs.Keys)) {
    if (-not $programs[$name]) {
        Remove-Item -LiteralPath $zipPath -Force
        throw "unixzip: $name is not in $root, so nothing in the zip would be runnable. No zip was written."
    }
}
Write-Host "unixzip: $zipPath, $entries entries, $($programs.Count) marked as programs"
