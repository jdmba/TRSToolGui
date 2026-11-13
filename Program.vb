' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Program.vb — entry point.
'  Usage: TRSToolGui.exe [image file]
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Windows.Forms

Public Module Program

    Public Const AppName As String = "TRSToolGui"
    Public Const AppVersion As String = "v261112"   ' build stamp, date-based, like Ira's other tools (v26MMDD)
    Public Const ReleaseVersion As String = "v1.1"   ' the release number shown in Help > About
    Public Const ReleaseDate As String = "November 12, 2026"   ' with a matching entry in Help\Changelog.txt
    Public Const AppAuthors As String = "Ira Goldklang and Claude AI"

    Public Settings As Settings
    Public Tool As Trs80Tool

    ''' <summary>Per-run temporary folder for drag-out/copy/viewer files; deleted at exit.</summary>
    Public ReadOnly TempRoot As String = Path.Combine(Path.GetTempPath(), "TRSToolGui", Guid.NewGuid().ToString("N").Substring(0, 8))

    <STAThread>
    Public Sub Main(args As String())
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.SetHighDpiMode(HighDpiMode.SystemAware)

        Settings = Settings.Load()
        Dim toolPath As String = Trs80Tool.Locate(Settings.ToolPath)
        If toolPath Is Nothing Then toolPath = AskForTool()
        Tool = New Trs80Tool(If(toolPath, ""))
        Tool.LogFile = If(Settings.EngineLogFile, "")
        Tool.UseServer = Settings.KeepEngineRunning

        Dim startFile As String = Nothing
        For Each a In args
            If a.StartsWith("-", StringComparison.Ordinal) OrElse a.StartsWith("/", StringComparison.Ordinal) Then Continue For
            startFile = a
        Next

        Try
            Application.Run(New MainForm(startFile))
        Finally
            Tool.Shutdown()
            Settings.Save()
            CleanTemp()
        End Try
    End Sub

    ''' <summary>
    ''' Neither the configured path nor the default trs80-tool location exists: make the user point at trs80-tool,
    ''' then save the choice to TRSToolGui.cfg. Returns Nothing if they cancel (the program then runs without
    ''' the engine and asks again at the next start).
    ''' </summary>
    Private Function AskForTool() As String
        MessageBox.Show("TRSToolGui needs trs80-tool, which was not found." & vbCrLf & vbCrLf &
            "If you don't have it, download ""trs80-tool for Windows 10+"" from https://www.my-trs-80.com/tool/." & vbCrLf & vbCrLf &
            "Click OK, then choose trs80-tool.exe. Its location is saved in:" & vbCrLf & Settings.FilePath,
            AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Using ofd As New OpenFileDialog()
            ofd.Title = "Locate trs80-tool"
            ofd.Filter = "Programs (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|All files (*.*)|*.*"
            ofd.FileName = Trs80Tool.DefaultExeName
            If ofd.ShowDialog() <> DialogResult.OK OrElse Not File.Exists(ofd.FileName) Then Return Nothing
            Settings.ToolPath = ofd.FileName
        End Using
        If Not Settings.Save() Then
            MessageBox.Show("The trs80-tool location could not be saved to" & vbCrLf & Settings.FilePath & vbCrLf & vbCrLf &
                "This program's folder may be read-only. The location will be used for this session only.",
                AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Return Settings.ToolPath
    End Function

    ''' <summary>Text of a file embedded from the Help folder (Manual.txt, Changelog.txt, License.txt).</summary>
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
