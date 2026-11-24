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
'  WHEN the listening starts is the whole trick, and v1.7 got it wrong. AppKit puts in
'  its own 'odoc' handler inside -[NSApplication finishLaunching], which runs when
'  Application.Run is called. A handler registered before Run is therefore thrown away
'  a moment later by AppKit's own, and AppKit's handler, finding no NSDocument class
'  and no openFile: on the delegate, is what puts up
'
'      The document "x.dmk" could not be opened.
'      TRSToolGui cannot open files in the "TRS-80 floppy disk image" format.
'
'  — a message that comes out of the bundle's own CFBundleTypeName. The dialog is
'  proof that AppKit, not this class, took the event.
'
'  finishLaunching posts NSApplicationWillFinishLaunchingNotification immediately
'  after it installs those handlers, and the launch document is not dispatched until
'  afterwards. So that notification is the one moment when registering wins: late
'  enough that AppKit cannot overwrite it, early enough that nothing has been missed.
'  Registering before Run is kept as well, for the case where the application is
'  already up and the notification has long since been and gone.
'
'  The notification is named by its string rather than by NSApplication's constant so
'  that this file goes on needing nothing from MonoMac.AppKit, and it is asked for with
'  an NSObject and a Selector — the same pattern the event handler itself uses, and the
'  only one this file already proves builds.
'
'  Every event is written to TRSToolGui-open.log beside TRSToolGui.cfg, whether or
'  not the file then opens, because that log is the evidence a problem report needs:
'  no entry at all means the event never arrived and the Info.plist is the suspect;
'  an entry with no path means the descriptor is not shaped the way this code reads it.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports MonoMac.Foundation
Imports MonoMac.ObjCRuntime

''' <summary>Receives the Finder's open-document events on a Mac. Not built on any other system.</summary>
Friend NotInheritable Class MacOpenDocuments
    Inherits NSObject

    ''' <summary>keyDirectObject, '----': where an 'odoc' event keeps its list of files.</summary>
    Private Const KeyDirectObject As UInteger = &H2D2D2D2DUI

    ''' <summary>Posted by -[NSApplication finishLaunching] just after AppKit installs its own handlers.</summary>
    Private Const WillFinishLaunching As String = "NSApplicationWillFinishLaunchingNotification"

    ''' <summary>Held for the life of the program: the event manager does not keep the handler alive.</summary>
    Private Shared _listener As MacOpenDocuments

    ''' <summary>True once the notification has been asked for, so it is asked for only once.</summary>
    Private Shared _watching As Boolean

    ''' <summary>
    ''' Start listening. Call once, before Application.Run. The handler is put in twice: now, which covers an
    ''' application that is already running, and again when AppKit says it is about to finish launching, which
    ''' is the registration that survives — see the note at the top of this file.
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

    ''' <summary>AppKit has just put its own handlers in. This is the moment to put ours over the top.</summary>
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
            Note("an open-document event arrived, " & count.ToString(CultureInfo.InvariantCulture) &
                 If(count = 1, " item", " items"))
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
    Private Shared Sub Note(text As String)
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
