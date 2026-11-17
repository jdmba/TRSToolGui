' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HostFiles.vb — gzip on the host side. trs80-tool 5.2.0 reads and writes
'  gzipped files in every command, and goes by the file NAME: a name that ends
'  in ".gz" is compressed, whatever is inside. So "disk.dmk.gz" is a DMK image,
'  and a host file named "X.gz" that is added to a disk goes on decompressed.
'  The helpers here let the GUI follow the same rule where it reads a file
'  itself, and keep a ".gz" file as it is when the user wants it stored so.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression

Public Module HostFiles

    Public Const GzipSuffix As String = ".gz"

    ''' <summary>True when the file name ends in ".gz" (any case), which is how trs80-tool decides.</summary>
    Public Function IsGzipName(pathOrName As String) As Boolean
        Return Not String.IsNullOrEmpty(pathOrName) AndAlso pathOrName.EndsWith(GzipSuffix, StringComparison.OrdinalIgnoreCase) AndAlso
            Path.GetFileName(pathOrName).Length > GzipSuffix.Length
    End Function

    ''' <summary>The path or name without a trailing ".gz".</summary>
    Public Function WithoutGzipSuffix(pathOrName As String) As String
        If IsGzipName(pathOrName) Then Return pathOrName.Substring(0, pathOrName.Length - GzipSuffix.Length)
        Return pathOrName
    End Function

    ''' <summary>
    ''' The extension that says what kind of image a file is, lower case with its dot: ".dmk" for both
    ''' "GAME.DMK" and "game.dmk.gz".
    ''' </summary>
    Public Function ImageExtension(imagePath As String) As String
        Return Path.GetExtension(WithoutGzipSuffix(If(imagePath, ""))).ToLowerInvariant()
    End Function

    ''' <summary>The bytes of a file, decompressed when its name ends in ".gz".</summary>
    Public Function ReadAllBytesUncompressed(filePath As String) As Byte()
        If Not IsGzipName(filePath) Then Return File.ReadAllBytes(filePath)
        Using source As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
            Using gz As New GZipStream(source, CompressionMode.Decompress)
                Using target As New MemoryStream()
                    gz.CopyTo(target)
                    Return target.ToArray()
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Decompress a ".gz" file into the folder, under its name without ".gz", and return the new file's path.
    ''' Throws InvalidDataException when the contents aren't gzip data.
    ''' </summary>
    Public Function GunzipInto(gzPath As String, folder As String) As String
        Dim target As String = Path.Combine(folder, Path.GetFileName(WithoutGzipSuffix(gzPath)))
        File.WriteAllBytes(target, ReadAllBytesUncompressed(gzPath))
        Return target
    End Function

    ''' <summary>
    ''' Copy a ".gz" file into the folder under a name trs80-tool won't take for a compressed file (the dot
    ''' before "gz" becomes an underscore), and return the copy's path. Adding the copy stores the bytes as
    ''' they are; the caller gives the TRS-80 name itself.
    ''' </summary>
    Public Function CopyUnderPlainName(gzPath As String, folder As String) As String
        Dim name As String = Path.GetFileName(gzPath)
        Dim target As String = Path.Combine(folder, name.Substring(0, name.Length - GzipSuffix.Length) & "_gz")
        File.Copy(gzPath, target, True)
        Return target
    End Function

End Module
