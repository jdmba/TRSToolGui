' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  MacOpenDocuments.vb — macOS only, and compiled only into the Mac build.
'
'  A Mac never hands a program an opened document on its command line. The Finder
'  sends the application an Apple event — 'odoc' in the 'aevt' class — and a program
'  that is not listening for it never learns that a file was opened at all. The file
'  types declared in the bundle's Info.plist are what get the Finder as far as
'  starting the program; this is what carries the path the rest of the way.
'
'  WHAT v1.7 AND v1.8 BOTH GOT WRONG, and how zebbchi's log settled it. Both releases
'  tried to be the 'odoc' handler themselves. v1.7 registered before Application.Run;
'  v1.8 registered again from NSApplicationWillFinishLaunchingNotification, on the
'  belief that -[NSApplication finishLaunching] posts that notification *after* it
'  installs AppKit's own handlers, so ours would go over the top of them. The v1.8 log
'  from a Mac says otherwise:
'
'      waiting for NSApplicationWillFinishLaunchingNotification
'      listening for open-document events (before the application starts), class 6170706C id 6F646F63
'      listening for open-document events (as the application finishes launching), class 6170706C id 6F646F63
'
'  — and then nothing. Both registrations went in (6170706C/6F646F63 is 'appl'/'odoc',
'  which also settles that MonoMac's AEEventClass.Application and AEEventID.OpenDocuments
'  are the right constants), no event ever arrived, and the user still got AppKit's
'  "cannot open files in the ... format" dialog. So the order is the other way round:
'  finishLaunching posts WillFinishLaunching FIRST, installs its own handlers after it,
'  and dispatches the launch document before it posts DidFinishLaunching. There is
'  therefore NO moment at which registering an 'odoc' handler can win the launch
'  document — not before Run, not at Will, and not at Did, which comes too late.
'
'  WHAT DOES WIN. AppKit's own 'odoc' handler looks for somewhere to send the file, in
'  order: an NSDocumentController, then the application delegate's application:openURLs:,
'  application:openFiles: or application:openFile:. Only when it finds none of those does
'  it put up the dialog. Giving the delegate one of those methods is the documented route
'  and the one that cannot be raced, because AppKit calls it itself.
'
'  Eto.Forms makes that reachable without any runtime trickery: Eto.Mac.AppDelegate is
'  public, MonoMac's NSApplicationDelegate.OpenFiles and .OpenFile are Overridable, and
'  Eto.Mac.Forms.ApplicationHandler.AppDelegate is a settable property that Eto reads when
'  it starts. So MacAppDelegate below simply inherits Eto's own delegate and adds the two
'  methods, and Install() hands it to the handler before Application.Run. Nothing of Eto's
'  is replaced or worked around.
'
'  The 'odoc' registration is KEPT, for two reasons: an application that is already
'  running when the Finder sends the event, and as the fallback that says in the log
'  whether this analysis was right. The two can never both fire for one event — AppKit's
'  handler is what calls the delegate, so either AppKit has the event (delegate) or this
'  class has it (handler) — so a file is never opened twice.
'
'  Every event is written to TRSToolGui-open.log beside TRSToolGui.cfg, whether or not
'  the file then opens, because that log is the evidence a problem report needs. The
'  line each route writes names the route, so one run says which one did the work.
'
'  SEVERAL FILES AT ONCE, which is what v1.8.3 is about. zebbchi reported on
'  26 November 2026 that on a Mac there was no telling how many images were open:
'  he would select a few in the Finder, open them, see ONE window, and find the
'  program behaving as though more than one image were open. Both halves of that
'  were real, and neither was a Mac fault.
'
'  The Finder sends every file of one open in a SINGLE 'odoc' event, so OpenFiles
'  below is called once with the whole array and calls Program.AskToOpen for each
'  name. What happened to them after that was the fault, in two different ways:
'
'    * at LAUNCH, no window exists yet, so every path went into Program's waiting
'      slot - which was one String, each write overwriting the one before it. Three
'      images in, one image out, and the other two gone without a word. It is a
'      List now, and MainForm.OnShown sends every one of them through the
'      window-list dispatcher.
'
'    * while the program was ALREADY RUNNING, every path was handed to the one and
'      only main window, which opened them one after another into itself. The last
'      one won the file list, but each had already added itself to Open Recent and
'      each had written to the title bar and the status line on its way past, which
'      is exactly why the program looked as though several images were open while
'      only one window was there. They were all in the same window.
'
'  So the answer to his question is yes, and it is yes from v1.8.3: one window per
'  image, each cascaded from the one in front, each with its image named in its own
'  title bar, and a Window menu listing them with the frontmost ticked. Nothing in
'  THIS file had to change for that - it was already handing every name over one at
'  a time, which was right all along - and that is why only its header moved.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports MonoMac.AppKit
Imports MonoMac.Foundation
Imports MonoMac.ObjCRuntime

''' <summary>
''' Eto's own application delegate with the two methods AppKit looks for when the Finder opens a
''' document. Installed by MacOpenDocuments.Install before Application.Run.
''' </summary>
Friend NotInheritable Class MacAppDelegate
    Inherits Eto.Mac.AppDelegate

    ''' <summary>AppKit's preferred route: every file of one Finder open, in one call.</summary>
    Public Overrides Sub OpenFiles(sender As NSApplication, filenames As String())
        Dim count As Integer = If(filenames Is Nothing, 0, filenames.Length)
        MacOpenDocuments.Note("the application delegate was given " & count.ToString(CultureInfo.InvariantCulture) &
                              If(count = 1, " file", " files") & " to open")
        Try
            If filenames IsNot Nothing Then
                For i As Integer = 0 To filenames.Length - 1
                    MacOpenDocuments.Note("  " & (i + 1).ToString(CultureInfo.InvariantCulture) & ": " & If(filenames(i), "(nothing)"))
                    If Not String.IsNullOrEmpty(filenames(i)) Then Program.AskToOpen(filenames(i))
                Next
            End If
        Catch ex As Exception
            MacOpenDocuments.Note("failed while opening the delegate's files: " & ex.Message)
        End Try
        ' AppKit waits to be told how the open went before it lets the launch finish.
        Try
            If sender IsNot Nothing Then sender.ReplyToOpenOrPrint(NSApplicationDelegateReply.Success)
        Catch
        End Try
    End Sub

    ''' <summary>The older one-file route, for a system that asks for it instead.</summary>
    Public Overrides Function OpenFile(sender As NSApplication, filename As String) As Boolean
        MacOpenDocuments.Note("the application delegate was given one file to open: " & If(filename, "(nothing)"))
        If String.IsNullOrEmpty(filename) Then Return False
        Try
            Program.AskToOpen(filename)
        Catch ex As Exception
            MacOpenDocuments.Note("failed while opening the delegate's file: " & ex.Message)
            Return False
        End Try
        Return True
    End Function

End Class

''' <summary>Receives the Finder's open-document events on a Mac. Not built on any other system.</summary>
Friend NotInheritable Class MacOpenDocuments
    Inherits NSObject

    ''' <summary>keyDirectObject, '----': where an 'odoc' event keeps its list of files.</summary>
    Private Const KeyDirectObject As UInteger = &H2D2D2D2DUI

    ''' <summary>Posted by -[NSApplication finishLaunching] before AppKit installs its own handlers.</summary>
    Private Const WillFinishLaunching As String = "NSApplicationWillFinishLaunchingNotification"

    ''' <summary>Held for the life of the program: the event manager does not keep the handler alive.</summary>
    Private Shared _listener As MacOpenDocuments

    ''' <summary>True once the notification has been asked for, so it is asked for only once.</summary>
    Private Shared _watching As Boolean

    ''' <summary>
    ''' Give Eto the delegate that carries application:openFiles:. Call before Application.Run — this is
    ''' the route that actually receives the Finder's launch document; see the note at the top of this file.
    ''' </summary>
    Friend Shared Sub Install(app As Eto.Forms.Application)
        Try
            Dim handler As Eto.Mac.Forms.ApplicationHandler = TryCast(app.Handler, Eto.Mac.Forms.ApplicationHandler)
            If handler Is Nothing Then
                Note("the Eto application handler is not the Mac one, so no delegate was installed")
                Return
            End If
            handler.AppDelegate = New MacAppDelegate()
            Note("installed the application delegate that answers application:openFiles:")
        Catch ex As Exception
            Note("could not install the application delegate: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Also listen for 'odoc' directly. This cannot win the launch document — AppKit's own handler goes in
    ''' after NSApplicationWillFinishLaunchingNotification — but it serves an application that is already
    ''' running, and the log line it writes says whether the delegate above did the work instead.
    ''' </summary>
    Friend Shared Sub Listen()
        If _listener Is Nothing Then
            Try
                _listener = New MacOpenDocuments()
            Catch ex As Exception
                Note("could not make the open-document listener: " & ex.Message)
                Return
            End Try
        End If

        If Not _watching Then
            Try
                NSNotificationCenter.DefaultCenter.AddObserver(
                    _listener,
                    New Selector("appWillFinishLaunching:"),
                    New NSString(WillFinishLaunching),
                    Nothing)
                _watching = True
                Note("waiting for " & WillFinishLaunching)
            Catch ex As Exception
                Note("could not wait for " & WillFinishLaunching & ": " & ex.Message)
            End Try
        End If

        Register("before the application starts")
    End Sub

    ''' <summary>AppKit is about to install its own handlers. Ours goes in again, for a later event.</summary>
    <Export("appWillFinishLaunching:")>
    Public Sub AppWillFinishLaunching(notification As NSNotification)
        Register("as the application finishes launching")
    End Sub

    ''' <summary>Put this class in as the handler of 'aevt'/'odoc'. Doing it again simply replaces it.</summary>
    Private Shared Sub Register(occasion As String)
        If _listener Is Nothing Then Return
        Try
            NSAppleEventManager.SharedAppleEventManager.SetEventHandler(
                _listener,
                New Selector("handleOpenDocuments:withReplyEvent:"),
                AEEventClass.Application,
                AEEventID.OpenDocuments)
            Note("listening for open-document events (" & occasion & "), class " &
                 CUInt(AEEventClass.Application).ToString("X8", CultureInfo.InvariantCulture) & " id " &
                 CUInt(AEEventID.OpenDocuments).ToString("X8", CultureInfo.InvariantCulture))
        Catch ex As Exception
            Note("could not listen for open-document events (" & occasion & "): " & ex.Message)
        End Try
    End Sub

    <Export("handleOpenDocuments:withReplyEvent:")>
    Public Sub HandleOpenDocuments(openEvent As NSAppleEventDescriptor, reply As NSAppleEventDescriptor)
        Try
            If openEvent Is Nothing Then
                Note("an open-document event arrived with nothing in it")
                Return
            End If
            Dim files As NSAppleEventDescriptor = openEvent.ParamDescriptorForKeyword(KeyDirectObject)
            If files Is Nothing Then
                Note("an open-document event arrived with no file list")
                Return
            End If
            Dim count As Integer = CInt(files.NumberOfItems)
            Note("an open-document event arrived at this class's own handler, " &
                 count.ToString(CultureInfo.InvariantCulture) & If(count = 1, " item", " items"))
            For i As Integer = 1 To count
                Dim item As NSAppleEventDescriptor = files.DescriptorAtIndex(i)
                Dim raw As String = If(item Is Nothing, Nothing, item.StringValue)
                Dim path As String = ToPath(raw)
                Note("  " & i.ToString(CultureInfo.InvariantCulture) & ": " &
                     If(raw, "(nothing)") & "  ->  " & If(path, "(no path)"))
                If path IsNot Nothing Then Program.AskToOpen(path)
            Next
        Catch ex As Exception
            Note("failed while reading an open-document event: " & ex.Message)
        End Try
    End Sub

    ''' <summary>The file's path. An 'odoc' item is usually a file URL; anything else is taken as it is.</summary>
    Private Shared Function ToPath(value As String) As String
        If String.IsNullOrEmpty(value) Then Return Nothing
        If value.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
            Try
                Return New Uri(value).LocalPath
            Catch
                Return Nothing
            End Try
        End If
        Return value
    End Function

    ''' <summary>One line in TRSToolGui-open.log. Never throws, and never grows without limit.</summary>
    Friend Shared Sub Note(text As String)
        Try
            Dim logFile As String = LogPath()
            If logFile Is Nothing Then Return
            Try
                Dim already As New FileInfo(logFile)
                If already.Exists AndAlso already.Length > 65536 Then already.Delete()
            Catch
            End Try
            File.AppendAllText(logFile,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) &
                "  " & text & Environment.NewLine, Encoding.UTF8)
        Catch
        End Try
    End Sub

    Private Shared Function LogPath() As String
        Try
            Dim folder As String = Path.GetDirectoryName(Settings.FilePath)
            If String.IsNullOrEmpty(folder) Then Return Nothing
            Return Path.Combine(folder, "TRSToolGui-open.log")
        Catch
            Return Nothing
        End Try
    End Function

End Class
