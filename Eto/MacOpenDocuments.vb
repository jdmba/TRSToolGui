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
'  Eto.Forms has no portable event for this, and Eto.Mac64 2.11 implements neither
'  application:openFile: nor application:openFiles:, so the event is taken straight
'  from NSAppleEventManager. Registration happens before Application.Run, because
'  the event is delivered as soon as the run loop starts — and it can arrive either
'  side of the main window appearing, which is why the path goes to Program.AskToOpen
'  rather than to a window this class knows about.
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

    ''' <summary>Held for the life of the program: the event manager does not keep the handler alive.</summary>
    Private Shared _listener As MacOpenDocuments

    ''' <summary>Start listening. Call once, before Application.Run.</summary>
    Friend Shared Sub Listen()
        If _listener IsNot Nothing Then Return
        Try
            Dim listener As New MacOpenDocuments()
            NSAppleEventManager.SharedAppleEventManager.SetEventHandler(
                listener,
                New Selector("handleOpenDocuments:withReplyEvent:"),
                AEEventClass.Application,
                AEEventID.OpenDocuments)
            _listener = listener
            Note("listening for open-document events")
        Catch ex As Exception
            Note("could not listen for open-document events: " & ex.Message)
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
