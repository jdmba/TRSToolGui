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

    ' The name and the version stamps live in the Core library (AppInfo.vb), shared with the Linux program.
    Public Const AppName As String = AppInfo.AppName
    Public Const AppVersion As String = AppInfo.AppVersion
    Public Const ReleaseVersion As String = AppInfo.ReleaseVersion
    Public Const ReleaseDate As String = AppInfo.ReleaseDate
    Public Const AppAuthors As String = AppInfo.AppAuthors

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
        ' The language goes in before anything reaches the screen, because text already built is not
        ' revisited. An empty setting, which is what a new installation has, leaves English in force.
        If Not String.IsNullOrEmpty(Settings.Language) Then UiText.Use(Settings.Language)
        Dim toolPath As String = Trs80Tool.Locate(Settings.ToolPath)
        If toolPath Is Nothing Then toolPath = AskForTool()
        Tool = New Trs80Tool(If(toolPath, ""))
        Tool.LogFile = If(Settings.EngineLogFile, "")
        Tool.UseServer = Settings.KeepEngineRunning
        Tool.TerminalCommand = If(Settings.TerminalCommand, "")

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
        MessageBox.Show(UiText.T("msg.toolmissing.head") & vbCrLf & vbCrLf &
            UiText.T("msg.toolmissing.download.windows") & vbCrLf & vbCrLf &
            UiText.T("msg.toolmissing.choose", "trs80-tool.exe") & vbCrLf & Settings.FilePath,
            AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("dlg.prefs.picktool")
            ofd.Filter = UiText.T("filter.programs")
            ofd.FileName = Trs80Tool.DefaultExeName
            If ofd.ShowDialog() <> DialogResult.OK OrElse Not File.Exists(ofd.FileName) Then Return Nothing
            Settings.ToolPath = ofd.FileName
        End Using
        If Not Settings.Save() Then
            MessageBox.Show(UiText.TCrLf("msg.toolsavefailed", Settings.FilePath, UiText.T("msg.toolsavefailed.readonly")),
                AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Return Settings.ToolPath
    End Function

    ''' <summary>Text of a file embedded from the Help folder (Manual.txt, Changelog.txt, License.txt).</summary>
    Public Function ReadHelpText(name As String) As String
        Using st As Stream = GetType(Program).Assembly.GetManifestResourceStream("TRSToolGui." & name)
            If st Is Nothing Then Return UiText.T("msg.helpmissing", name)
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
