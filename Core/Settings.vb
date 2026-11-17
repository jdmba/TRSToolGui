' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Settings.vb — user preferences, stored as JSON in TRSToolGui.cfg in the
'  folder this program runs from (TRSTool kept them in the registry under
'  HKCU\Software\Matthew Reed\TRSTools; this program keeps its own file).
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.Json

Public Enum TextConversionOut
    None = 0
    CrToCrLf = 1
    CrToLf = 2
End Enum

Public Enum TextConversionIn
    None = 0
    CrLfToCr = 1
End Enum

Public Class Settings
    ''' <summary>Full path of trs80-tool (may be renamed). Empty = look for it (Trs80Tool.Locate), else ask at startup.</summary>
    Public Property ToolPath As String = ""
    Public Property ShowInvisible As Boolean = False
    Public Property ShowSystem As Boolean = False
    Public Property ExtractDirectory As String = ""
    Public Property ConversionFromTrs80 As TextConversionOut = HostPlatform.DefaultConversionFromTrs80
    Public Property ConversionToTrs80 As TextConversionIn = TextConversionIn.CrLfToCr
    ''' <summary>Let trs80-tool tokenize BASIC text (.bas) when adding. Off = store bytes exactly, as TRSTool did.</summary>
    Public Property TokenizeBasicOnAdd As Boolean = True
    ''' <summary>Record the host file's modification date on files added (TRSTool did).</summary>
    Public Property UseHostDateOnAdd As Boolean = True
    ''' <summary>
    ''' A host file named *.gz that is added to a disk: False = stored as it is (what TRSTools did), True = unzipped
    ''' first and the file inside stored (what trs80-tool 5.2.0 does by itself).
    ''' </summary>
    Public Property UnzipGzOnAdd As Boolean = False
    Public Property ExitOnEscape As Boolean = False
    Public Property ConfirmDelete As Boolean = True
    Public Property ViewerMonospaced As Boolean = True
    Public Property LowercaseExtractNames As Boolean = False
    Public Property Mru As New List(Of String)
    Public Property WindowWidth As Integer = 760
    Public Property WindowHeight As Integer = 480
    Public Property WindowX As Integer = -1
    Public Property WindowY As Integer = -1
    Public Property WindowMaximized As Boolean = False
    Public Property ColumnWidths As List(Of Integer) = Nothing
    ''' <summary>Last trs80-tool version the user was told about (so the "newer than tested" note shows once).</summary>
    Public Property LastNotedVersion As String = ""
    ''' <summary>If set, every trs80-tool call gets --log-file THIS (the tool appends its logs; for bug reports to Lawrence).</summary>
    Public Property EngineLogFile As String = ""
    ''' <summary>Keep one trs80-tool running in the background ("trs80-tool mcp") instead of starting it for every action.</summary>
    Public Property KeepEngineRunning As Boolean = True
    ''' <summary>
    ''' Linux and macOS: the terminal program for the disk shell and the emulator, with the words that go before
    ''' the script, for example "xterm -e". Blank = find one. Not used on Windows.
    ''' </summary>
    Public Property TerminalCommand As String = ""
    ''' <summary>Option words for every "disasm" call (viewer and Tools → Disassemble). Blank = DefaultDisasmOptions.</summary>
    Public Property DisasmOptions As String = DefaultDisasmOptions

    Public Const MaxMru As Integer = 9
    ''' <summary>
    ''' TRSTools-style listing: address and bytes, every byte decoded as an
    ''' instruction, no labels, upper case, hex as 0C8H, an ASCII comment per line.
    ''' </summary>
    Public Const DefaultDisasmOptions As String = "--data-comment --upper --full --no-labels --listing --no-known --hex-format h"
    ''' <summary>The plain-source default saved by an earlier build; Load replaces it with DefaultDisasmOptions.</summary>
    Private Const PlainDisasmOptions As String = "--upper --no-known --hex-format h"

    ''' <summary>DisasmOptions split into the words DisasmAsync takes (blank falls back to the default).</summary>
    Public Function DisasmOptionWords() As List(Of String)
        Dim line As String = If(String.IsNullOrWhiteSpace(DisasmOptions), DefaultDisasmOptions, DisasmOptions)
        Return line.Split(New Char() {" "c, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries).ToList()
    End Function

    ''' <summary>
    ''' TRSToolGui.cfg. On Windows it is beside TRSToolGui.exe, so the program and its settings travel together;
    ''' on Linux and macOS it is in the user's configuration folder (HostPlatform.SettingsDirectory).
    ''' </summary>
    Public Shared ReadOnly Property FilePath As String
        Get
            Return Path.Combine(HostPlatform.SettingsDirectory, "TRSToolGui.cfg")
        End Get
    End Property

    ''' <summary>Where earlier builds kept the settings; read if TRSToolGui.cfg does not exist yet.</summary>
    Private Shared ReadOnly Property LegacyFilePath As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TRSToolGui", "settings.json")
        End Get
    End Property

    Public Shared Function Load() As Settings
        Try
            Dim source As String = If(File.Exists(FilePath), FilePath, LegacyFilePath)
            If File.Exists(source) Then
                Dim s As Settings = JsonSerializer.Deserialize(Of Settings)(File.ReadAllText(source), New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True, .AllowTrailingCommas = True})
                If s IsNot Nothing Then
                    If s.Mru Is Nothing Then s.Mru = New List(Of String)
                    If s.DisasmOptions = PlainDisasmOptions Then s.DisasmOptions = DefaultDisasmOptions
                    Return s
                End If
            End If
        Catch
            ' Corrupt or unreadable settings: start fresh rather than fail.
        End Try
        Return New Settings()
    End Function

    ''' <summary>Write TRSToolGui.cfg. Returns False if it could not be written (e.g. a read-only program folder).</summary>
    Public Function Save() As Boolean
        Try
            ' The program's own folder exists; a per-user configuration folder may not yet.
            If Not HostPlatform.IsWindows Then Directory.CreateDirectory(HostPlatform.SettingsDirectory)
            File.WriteAllText(FilePath, JsonSerializer.Serialize(Me, New JsonSerializerOptions With {.WriteIndented = True}))
            Return True
        Catch
            ' Settings are a convenience; never let saving them break the program.
            Return False
        End Try
    End Function

    Public Sub AddMru(path As String)
        If String.IsNullOrWhiteSpace(path) Then Return
        Mru.RemoveAll(Function(x) String.Equals(x, path, HostPlatform.PathComparison))
        Mru.Insert(0, path)
        While Mru.Count > MaxMru
            Mru.RemoveAt(Mru.Count - 1)
        End While
    End Sub

    Public Sub RemoveMru(path As String)
        Mru.RemoveAll(Function(x) String.Equals(x, path, HostPlatform.PathComparison))
    End Sub
End Class
