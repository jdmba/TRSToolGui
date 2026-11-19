' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Program.vb — entry point of the Eto.Forms program (Linux and macOS).
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

    Private _shutDown As Boolean = False

    <STAThread>
    Public Sub Main(args As String())
        Dim app As New Application(Eto.Platform.Detect)
        If HostPlatform.IsMacOS Then
            app.Name = AppName
            ' On a Mac the program ends inside Quit: Run never comes back, so the Finally below is not reached
            ' there. The same work is done when the application says it is about to end.
            AddHandler app.Terminating, Sub(sender, e)
                                            If Not e.Cancel Then ShutDown()
                                        End Sub
        End If

        Settings = Settings.Load()
        Dim toolPath As String = Trs80Tool.Locate(Settings.ToolPath)
        ' On a Mac nothing can be asked before the application is running; there the main window asks, once
        ' it is on the screen (MainForm.CheckEngineAsync).
        If toolPath Is Nothing AndAlso Not HostPlatform.IsMacOS Then toolPath = AskForTool()
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
            ShutDown()
        End Try
    End Sub

    ''' <summary>The work at the end of a run, done once: stop the background trs80-tool, save the settings, remove the temp folder.</summary>
    Private Sub ShutDown()
        If _shutDown Then Return
        _shutDown = True
        Tool.Shutdown()
        Settings.Save()
        CleanTemp()
    End Sub

    ''' <summary>End the program. On a Mac closing the main window only hides it, so Quit is asked for by name.</summary>
    Public Sub ExitProgram(mainWindow As Form)
        If HostPlatform.IsMacOS Then
            Application.Instance.Quit()
        Else
            mainWindow.Close()
        End If
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

    ''' <summary>Where a trs80-tool is found without asking, said for the system the program runs on ("" on Windows).</summary>
    Public ReadOnly Property ToolFoundWhereHint As String
        Get
            If HostPlatform.IsWindows Then Return ""
            If Not HostPlatform.IsMacOS Then Return " A copy beside TRSToolGui, or one on the PATH, is found without asking."
            Dim hint As String = " A copy in the folder that TRSToolGui.app is in, or one in /usr/local/bin or /opt/homebrew/bin, is found without asking."
            If HostPlatform.IsMacTranslocated Then
                hint &= vbLf & vbLf & "macOS is running TRSToolGui from a temporary copy of its folder, because the app has not been moved since it was " &
                    "unzipped. A trs80-tool beside the app can't be seen from there. To have it found: quit, drag TRSToolGui.app to another folder " &
                    "(Applications, for one) in Finder, put trs80-tool beside it, and start it again."
            End If
            Return hint
        End Get
    End Property

    ''' <summary>
    ''' What a Mac user has to do once to a trs80-tool that came from a browser download, or "" on other
    ''' systems: the file needs its execute permission, and macOS blocks it until its download mark is gone.
    ''' </summary>
    Public Function MacToolHint(toolPath As String) As String
        If Not HostPlatform.IsMacOS Then Return ""
        Dim quoted As String = "'" & If(toolPath, "trs80-tool").Replace("'", "'\''") & "'"
        Return "On a Mac, a trs80-tool that was downloaded with a browser has to be made runnable and allowed once. " &
            "Open Terminal and enter these two lines, then start TRSToolGui again:" & vbLf & vbLf &
            "chmod +x " & quoted & vbLf &
            "xattr -d com.apple.quarantine " & quoted & vbLf & vbLf &
            "(If the second line answers ""No such xattr"", the file was already allowed.)"
    End Function

    Private Function AskForTool() As String
        Ui.Info(Nothing, "TRSToolGui needs trs80-tool, which was not found." & vbLf & vbLf &
            ToolDownloadHint & "." & ToolFoundWhereHint & vbLf & vbLf &
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
