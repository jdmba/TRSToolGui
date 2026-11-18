' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Program.vb — entry point of the Eto.Forms program (Linux).
'  Usage: TRSToolGui [image file]
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports Eto.Forms

Public Module Program

    Public Const AppName As String = AppInfo.AppName
    Public Const AppVersion As String = AppInfo.AppVersion
    Public Const ReleaseVersion As String = AppInfo.ReleaseVersion
    Public Const ReleaseDate As String = AppInfo.ReleaseDate
    Public Const AppAuthors As String = AppInfo.AppAuthors

    Public Settings As Settings
    Public Tool As Trs80Tool

    Public ReadOnly TempRoot As String = Path.Combine(Path.GetTempPath(), "TRSToolGui", Guid.NewGuid().ToString("N").Substring(0, 8))

    <STAThread>
    Public Sub Main(args As String())
        Dim app As New Application(Eto.Platform.Detect)

        Settings = Settings.Load()
        Dim toolPath As String = Trs80Tool.Locate(Settings.ToolPath)
        If toolPath Is Nothing Then toolPath = AskForTool()
        Tool = New Trs80Tool(If(toolPath, ""))
        Tool.LogFile = If(Settings.EngineLogFile, "")
        Tool.UseServer = Settings.KeepEngineRunning
        Tool.TerminalCommand = If(Settings.TerminalCommand, "")

        Dim startFile As String = Nothing
        For Each a In args
            If a.StartsWith("-", StringComparison.Ordinal) Then Continue For
            startFile = a
        Next

        Try
            app.Run(New MainForm(startFile))
        Finally
            Tool.Shutdown()
            Settings.Save()
            CleanTemp()
        End Try
    End Sub

    ''' <summary>Where to get trs80-tool, worded for the system the program runs on.</summary>
    Public ReadOnly Property ToolDownloadHint As String
        Get
            Dim which As String = If(HostPlatform.IsWindows, """trs80-tool for Windows 10+""", If(HostPlatform.IsMacOS, "the macOS download", "the Linux download"))
            Return "Download it from " & AppInfo.ToolUrl & " (" & which & ")"
        End Get
    End Property

    ''' <summary>The file filter for locating trs80-tool: it has no extension outside Windows.</summary>
    Public ReadOnly Property ToolFileFilter As String
        Get
            Return If(HostPlatform.IsWindows, "Programs (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|All files (*.*)|*.*", "All files|*")
        End Get
    End Property

    Private Function AskForTool() As String
        Ui.Info(Nothing, "TRSToolGui needs trs80-tool, which was not found." & vbLf & vbLf &
            ToolDownloadHint & "." & If(HostPlatform.IsWindows, "", " A copy beside TRSToolGui, or one on the PATH, is found without asking.") & vbLf & vbLf &
            "Click OK, then choose " & Trs80Tool.DefaultExeName & ". Its location is saved in:" & vbLf & Settings.FilePath, AppName)
        Dim chosen As String = Ui.PickFile(Nothing, "Locate trs80-tool", ToolFileFilter)
        If chosen Is Nothing OrElse Not File.Exists(chosen) Then Return Nothing
        Settings.ToolPath = chosen
        If Not Settings.Save() Then
            Ui.Warn(Nothing, "The trs80-tool location could not be saved to" & vbLf & Settings.FilePath & vbLf & vbLf &
                "The location will be used for this session only.", AppName)
        End If
        Return Settings.ToolPath
    End Function

    Public Function ReadHelpText(name As String) As String
        Using st As Stream = GetType(Program).Assembly.GetManifestResourceStream("TRSToolGui." & name)
            If st Is Nothing Then Return "(" & name & " is missing from this build.)"
            Using sr As New StreamReader(st, System.Text.Encoding.UTF8)
                Return sr.ReadToEnd()
            End Using
        End Using
    End Function

    Public Function NewTempDir() As String
        Dim d As String = Path.Combine(TempRoot, Guid.NewGuid().ToString("N").Substring(0, 8))
        Directory.CreateDirectory(d)
        Return d
    End Function

    Public Sub CleanTemp()
        Try
            If Directory.Exists(TempRoot) Then Directory.Delete(TempRoot, True)
            ' Also remove the parent if it is now empty (other instances may be using it).
            Dim parent As String = Path.GetDirectoryName(TempRoot)
            If Directory.Exists(parent) AndAlso Directory.GetFileSystemEntries(parent).Length = 0 Then Directory.Delete(parent)
        Catch
        End Try
    End Sub

End Module
