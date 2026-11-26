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
#If TRSMAC Then
            ' The Finder's launch document reaches the application delegate, which AppKit calls itself,
            ' so the delegate goes in before Run; listening for the 'odoc' event as well covers an
            ' application that is already running. Both are in MacOpenDocuments.vb, which is in the Mac
            ' build and no other, and the note at the top of that file says why one is not enough.
            MacOpenDocuments.Install(app)
            MacOpenDocuments.Listen()
#End If
            ' On a Mac the program ends inside Quit: Run never comes back, so the Finally below is not reached
            ' there. The same work is done when the application says it is about to end.
            AddHandler app.Terminating, Sub(sender, e)
                                            If Not e.Cancel Then ShutDown()
                                        End Sub
        End If

        Settings = Settings.Load()
        ' The language goes in before anything reaches the screen, because text already built is not
        ' revisited. An empty setting, which is what a new installation has, leaves English in force.
        If Not String.IsNullOrEmpty(Settings.Language) Then UiText.Use(Settings.Language)

        ' A Mac gets a Window menu from Eto, holding Minimize, Zoom and Bring All To Front, and
        ' TRSToolGui puts the list of open images into a Window menu of its own. There must be ONE such
        ' menu, not two, and Eto finds its own by name (MenuItemCollection.GetSubmenu, comparing the text
        ' with the & taken out) - which fails the moment the program is in French and ours says Fenetre.
        ' LocalizeString is Eto's own event for exactly this: answering it for "&Window" with whatever
        ' the language table holds makes the two names agree, so Eto's three items go into our menu.
        ' It has to be in place before any menu bar is built, which is why it is here and not in a form.
        AddHandler app.LocalizeString,
            Sub(sender As Object, e As LocalizeEventArgs)
                If String.Equals(e.Text, "&Window", StringComparison.Ordinal) Then e.LocalizedText = UiText.T("menu.window")
            End Sub
        Dim toolPath As String = Trs80Tool.Locate(Settings.ToolPath)
        ' On a Mac nothing can be asked before the application is running; there the main window asks, once
        ' it is on the screen (MainForm.CheckEngineAsync).
        If toolPath Is Nothing AndAlso Not HostPlatform.IsMacOS Then toolPath = AskForTool()
        Tool = New Trs80Tool(If(toolPath, ""))
        Tool.LogFile = If(Settings.EngineLogFile, "")
        Tool.UseServer = Settings.KeepEngineRunning
        Tool.TerminalCommand = If(Settings.TerminalCommand, "")

        ' EVERY file named on the command line, not only the last of them: a file manager asked to open
        ' several images can start the program with all of them at once, and until v1.8.3 this loop kept
        ' overwriting its answer so that all but one were silently dropped. (A Mac never passes an opened
        ' document on the command line at all; there they arrive as an Apple event, see MacOpenDocuments.)
        Dim startFiles As New List(Of String)
        For Each a In args
            If a.StartsWith("-", StringComparison.Ordinal) Then Continue For
            startFiles.Add(a)
        Next

        Try
            Dim first As New MainForm(If(startFiles.Count > 0, startFiles(0), Nothing))
            ' The rest are asked for rather than opened here: a window cannot be cascaded from one that
            ' is not on the screen yet, so they wait for the first window to be shown. AskToOpen holds
            ' them until MainForm.OnShown has told it where to send them.
            For i As Integer = 1 To startFiles.Count - 1
                AskToOpen(startFiles(i))
            Next
            app.Run(first)
        Finally
            ShutDown()
        End Try
    End Sub

    ''' <summary>
    ''' Where a window says how to open a file the system has asked for. A Mac never passes an opened
    ''' document on the command line; it sends an Apple event that can arrive either side of the main
    ''' window appearing, so whichever is ready first wins: the window sets this when it is shown and
    ''' takes anything already waiting, and an event after that goes straight here.
    ''' </summary>
    Public Opener As Action(Of String)

    ''' <summary>
    ''' The files that arrived before any window was ready. A LIST since v1.8.3, and that is the whole of
    ''' one half of zebbchi's report: it used to be a single String, so when the Finder sent three images
    ''' in one 'odoc' event each AskToOpen overwrote the one before it and two of the three were lost
    ''' before any window had been built. Opening three images and getting one was the symptom.
    ''' </summary>
    Private ReadOnly _waitingToOpen As New List(Of String)

    ''' <summary>Open this file, now if a window is ready for it and otherwise as soon as one is.</summary>
    Public Sub AskToOpen(imagePath As String)
        If String.IsNullOrEmpty(imagePath) Then Return
        Dim ready As Action(Of String) = Opener
        If ready Is Nothing Then
            SyncLock _waitingToOpen
                _waitingToOpen.Add(imagePath)
            End SyncLock
            Return
        End If
        Try
            Application.Instance.AsyncInvoke(Sub() ready(imagePath))
        Catch
            SyncLock _waitingToOpen
                _waitingToOpen.Add(imagePath)
            End SyncLock
        End Try
    End Sub

    ''' <summary>The files that arrived before any window was ready. Taken only once; never Nothing.</summary>
    Public Function TakeWaitingToOpen() As List(Of String)
        SyncLock _waitingToOpen
            Dim waiting As New List(Of String)(_waitingToOpen)
            _waitingToOpen.Clear()
            Return waiting
        End SyncLock
    End Function

    ''' <summary>The work at the end of a run, done once: stop the background trs80-tool, save the settings, remove the temp folder.</summary>
    Private Sub ShutDown()
        If _shutDown Then Return
        _shutDown = True
        Tool.Shutdown()
        Settings.Save()
        CleanTemp()
    End Sub

    ''' <summary>
    ''' End the program. On a Mac that is Quit by name, because closing a window there leaves the
    ''' application running, which is what a Mac user expects. Elsewhere every main window is closed, not
    ''' just the one the command came from: Exit means the program, and more than one window can be open.
    ''' </summary>
    Public Sub ExitProgram(mainWindow As Form)
        If HostPlatform.IsMacOS Then
            Application.Instance.Quit()
        ElseIf WindowList.Count > 1 Then
            WindowList.CloseAll()
        Else
            mainWindow.Close()
        End If
    End Sub

    ''' <summary>Where to get trs80-tool, worded for the system the program runs on.</summary>
    Public ReadOnly Property ToolDownloadHint As String
        Get
            Dim which As String = If(HostPlatform.IsWindows, UiText.T("msg.tooldownload.windows"), If(HostPlatform.IsMacOS, UiText.T("msg.tooldownload.mac"), UiText.T("msg.tooldownload.linux")))
            Return UiText.T("msg.tooldownload", AppInfo.ToolUrl, which)
        End Get
    End Property

    ''' <summary>The file filter for locating trs80-tool: it has no extension outside Windows.</summary>
    Public ReadOnly Property ToolFileFilter As String
        Get
            Return If(HostPlatform.IsWindows, UiText.T("filter.programs"), UiText.T("filter.allfiles"))
        End Get
    End Property

    ''' <summary>Where a trs80-tool is found without asking, said for the system the program runs on ("" on Windows).</summary>
    Public ReadOnly Property ToolFoundWhereHint As String
        Get
            If HostPlatform.IsWindows Then Return ""
            If Not HostPlatform.IsMacOS Then Return UiText.T("msg.toolfound.unix")
            Dim hint As String = UiText.T("msg.toolfound.mac")
            If HostPlatform.IsMacTranslocated Then
                hint &= vbLf & vbLf & UiText.T("msg.toolfound.translocated")
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
        ' The two command lines are what the user types, so they stay here.
        Return UiText.T("msg.mactoolhint.intro") & vbLf & vbLf &
            "chmod +x " & quoted & vbLf &
            "xattr -d com.apple.quarantine " & quoted & vbLf & vbLf &
            UiText.T("msg.mactoolhint.note")
    End Function

    Private Function AskForTool() As String
        Ui.Info(Nothing, UiText.T("msg.toolmissing.head") & vbLf & vbLf &
            ToolDownloadHint & "." & ToolFoundWhereHint & vbLf & vbLf &
            UiText.T("msg.toolmissing.choose", Trs80Tool.DefaultExeName) & vbLf & Settings.FilePath, AppName)
        Dim chosen As String = Ui.PickFile(Nothing, UiText.T("dlg.prefs.picktool"), ToolFileFilter)
        If chosen Is Nothing OrElse Not File.Exists(chosen) Then Return Nothing
        Settings.ToolPath = chosen
        If Not Settings.Save() Then
            Ui.Warn(Nothing, UiText.T("msg.toolsavefailed", Settings.FilePath, UiText.T("msg.toolsavefailed.session")), AppName)
        End If
        Return Settings.ToolPath
    End Function

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
