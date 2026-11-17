' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HostPlatform.vb — everything that differs between Windows, Linux and macOS
'  and is not about a window toolkit: what trs80-tool is called and where to
'  look for it, where the settings file lives, how file paths compare, the
'  default line ending for extracted text, and the names of the fonts.
'
'  On Windows every answer here is what TRSToolGui 1.1 did, so the Windows
'  program behaves as before. The Linux and macOS answers are for the port.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO

Public Module HostPlatform

    Public ReadOnly Property IsWindows As Boolean
        Get
            Return OperatingSystem.IsWindows()
        End Get
    End Property

    Public ReadOnly Property IsMacOS As Boolean
        Get
            Return OperatingSystem.IsMacOS()
        End Get
    End Property

    Public ReadOnly Property IsLinux As Boolean
        Get
            Return Not IsWindows AndAlso Not IsMacOS
        End Get
    End Property

    ''' <summary>The file name trs80-tool is released under on this system.</summary>
    Public ReadOnly Property ToolFileName As String
        Get
            Return If(IsWindows, "trs80-tool.exe", "trs80-tool")
        End Get
    End Property

    ''' <summary>
    ''' How two host file paths compare. Windows and macOS file systems ignore case; Linux file systems don't.
    ''' TRS-80 file names are not host paths and always compare ignoring case.
    ''' </summary>
    Public ReadOnly Property PathComparison As StringComparison
        Get
            Return If(IsLinux, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    Public Function SamePath(a As String, b As String) As Boolean
        Return String.Equals(a, b, PathComparison)
    End Function

    ''' <summary>
    ''' The folder of TRSToolGui.cfg. On Windows it is the program's own folder, so the program and its settings
    ''' travel together. On macOS that folder is inside the signed application and on Linux it may not be
    ''' writable, so there the file goes in the user's own configuration folder.
    ''' </summary>
    Public ReadOnly Property SettingsDirectory As String
        Get
            If IsWindows Then Return AppContext.BaseDirectory
            Dim home As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            If IsMacOS Then Return Path.Combine(home, "Library", "Application Support", "TRSToolGui")
            Dim xdg As String = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME")
            If Not String.IsNullOrWhiteSpace(xdg) AndAlso Path.IsPathRooted(xdg) Then Return Path.Combine(xdg, "TRSToolGui")
            Return Path.Combine(home, ".config", "TRSToolGui")
        End Get
    End Property

    ''' <summary>The line ending extracted TRS-80 text gets when no preference has been saved: CR/LF on Windows, LF elsewhere.</summary>
    Public ReadOnly Property DefaultConversionFromTrs80 As TextConversionOut
        Get
            Return If(IsWindows, TextConversionOut.CrToCrLf, TextConversionOut.CrToLf)
        End Get
    End Property

    ''' <summary>
    ''' A fixed-width font that has the box-drawing characters of trs80-tool's text maps: Consolas on Windows,
    ''' Menlo on macOS, DejaVu Sans Mono on Linux.
    ''' </summary>
    Public ReadOnly Property MonospaceFontName As String
        Get
            If IsWindows Then Return "Consolas"
            If IsMacOS Then Return "Menlo"
            Return "DejaVu Sans Mono"
        End Get
    End Property

    ''' <summary>The system's font for window text.</summary>
    Public ReadOnly Property UiFontName As String
        Get
            If IsWindows Then Return "Segoe UI"
            If IsMacOS Then Return "Helvetica Neue"
            Return "DejaVu Sans"
        End Get
    End Property

    ''' <summary>A proportional serif font for the viewer's "proportional" choice.</summary>
    Public ReadOnly Property SerifFontName As String
        Get
            If IsWindows Then Return "Times New Roman"
            If IsMacOS Then Return "Times"
            Return "DejaVu Serif"
        End Get
    End Property

    ''' <summary>The full path of a program on the PATH, or Nothing. On Windows the PATHEXT extensions are tried too.</summary>
    Public Function FindOnPath(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then Return Nothing
        Dim pathVar As String = Environment.GetEnvironmentVariable("PATH")
        If String.IsNullOrEmpty(pathVar) Then Return Nothing
        Dim names As New List(Of String) From {fileName}
        If IsWindows AndAlso Path.GetExtension(fileName).Length = 0 Then
            Dim pathExt As String = Environment.GetEnvironmentVariable("PATHEXT")
            If String.IsNullOrEmpty(pathExt) Then pathExt = ".COM;.EXE;.BAT;.CMD"
            For Each ext In pathExt.Split(";"c, StringSplitOptions.RemoveEmptyEntries)
                names.Add(fileName & ext.ToLowerInvariant())
            Next
        End If
        For Each folder In pathVar.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            For Each name In names
                Try
                    Dim candidate As String = Path.Combine(folder.Trim().Trim(""""c), name)
                    If File.Exists(candidate) Then Return Path.GetFullPath(candidate)
                Catch
                    ' A malformed PATH entry: skip it.
                End Try
            Next
        Next
        Return Nothing
    End Function

End Module
