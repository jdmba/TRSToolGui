' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  ConsoleRun.vb — one run of trs80-tool in a console window of its own (the
'  interactive disk shell, the emulator). The caller keeps the disk locked
'  until HasExited is true, and can listen for Exited.
'
'  Windows: the system opens a console window for trs80-tool, and the run is
'  that process. This is what TRSToolGui 1.1 did.
'
'  Linux and macOS: a program started by a GUI has no terminal, so trs80-tool
'  is started through a short shell script in a terminal program (Terminal.app
'  through "open" on macOS; the first terminal emulator found on Linux, or the
'  one named in the settings). Terminal programs hand the work to a process the
'  GUI never sees, so the script itself reports the end: it writes trs80-tool's
'  exit status to a marker file, which is watched here.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks

Public Class ConsoleRun
    Implements IDisposable

    ''' <summary>Raised once, on a background thread, when the run has ended.</summary>
    Public Event Exited As EventHandler

    Private _process As Process = Nothing
    Private _folder As String = Nothing
    Private _marker As String = Nothing
    Private _ended As Integer = 0
    Private _cancel As CancellationTokenSource = Nothing

    Private Sub New()
    End Sub

    ''' <summary>The terminal programs tried on Linux, in order, each with the words that go before the script.</summary>
    Private Shared ReadOnly LinuxTerminals As String()() = {
        New String() {"x-terminal-emulator", "-e"},
        New String() {"gnome-terminal", "--"},
        New String() {"konsole", "-e"},
        New String() {"xfce4-terminal", "-x"},
        New String() {"mate-terminal", "-x"},
        New String() {"lxterminal", "-e"},
        New String() {"alacritty", "-e"},
        New String() {"kitty"},
        New String() {"xterm", "-e"}}

    Public ReadOnly Property HasExited As Boolean
        Get
            If Volatile.Read(_ended) <> 0 Then Return True
            If _process IsNot Nothing Then Return _process.HasExited
            Return _marker IsNot Nothing AndAlso File.Exists(_marker)
        End Get
    End Property

    ''' <summary>Windows: start the program with a console window of its own and follow that process.</summary>
    Friend Shared Function StartWithOwnWindow(psi As ProcessStartInfo) As ConsoleRun
        Dim run As New ConsoleRun()
        Dim p As Process = Process.Start(psi)
        If p Is Nothing Then Throw New ToolException("The system did not start trs80-tool.")
        run._process = p
        AddHandler p.Exited, Sub(sender, e) run.SignalEnded()
        p.EnableRaisingEvents = True
        Return run
    End Function

    ''' <summary>
    ''' Linux and macOS: run the program, with its arguments, in a terminal window.
    ''' </summary>
    ''' <param name="programPath">trs80-tool.</param>
    ''' <param name="arguments">Every argument, program options included.</param>
    ''' <param name="terminalCommand">The user's terminal command (program and the words before the script), or blank to find one.</param>
    Friend Shared Function StartInTerminal(programPath As String, arguments As IEnumerable(Of String), terminalCommand As String) As ConsoleRun
        Dim run As New ConsoleRun()
        run._folder = Path.Combine(Path.GetTempPath(), "TRSToolGui", "console-" & Guid.NewGuid().ToString("N").Substring(0, 8))
        Directory.CreateDirectory(run._folder)
        run._marker = Path.Combine(run._folder, "ended")
        ' ".command" is what macOS opens in Terminal; Linux doesn't mind the name.
        Dim script As String = Path.Combine(run._folder, "run.command")

        Dim sb As New StringBuilder()
        sb.Append("#!/bin/sh").Append(vbLf)
        sb.Append("marker=").Append(Quote(run._marker)).Append(vbLf)
        ' The marker is written however the script ends, a closed terminal window included.
        sb.Append("trap 'echo $? > ""$marker""' EXIT").Append(vbLf)
        sb.Append("trap 'exit 129' HUP INT TERM").Append(vbLf)
        sb.Append("cd ").Append(Quote(Path.GetDirectoryName(programPath))).Append(" || exit 1").Append(vbLf)
        sb.Append(Quote(programPath))
        For Each a In arguments
            sb.Append(" "c).Append(Quote(a))
        Next
        sb.Append(vbLf)
        File.WriteAllText(script, sb.ToString(), New UTF8Encoding(False))
        If Not OperatingSystem.IsWindows() Then
            File.SetUnixFileMode(script, UnixFileMode.UserRead Or UnixFileMode.UserWrite Or UnixFileMode.UserExecute)
        End If

        Dim words As List(Of String) = TerminalWords(terminalCommand)
        If words Is Nothing Then
            run.RemoveFolder()
            Throw New ToolException("No terminal program was found to run trs80-tool in. Name one in the settings file (TerminalCommand), for example ""xterm -e"".")
        End If
        Dim psi As New ProcessStartInfo(words(0)) With {.UseShellExecute = False, .WorkingDirectory = run._folder}
        For i As Integer = 1 To words.Count - 1
            psi.ArgumentList.Add(words(i))
        Next
        psi.ArgumentList.Add(script)
        Try
            Using starter As Process = Process.Start(psi)
            End Using
        Catch ex As Exception
            run.RemoveFolder()
            Throw New ToolException("The terminal program """ & words(0) & """ could not be started: " & ex.Message)
        End Try

        run._cancel = New CancellationTokenSource()
        Dim token As CancellationToken = run._cancel.Token
        Task.Run(Async Function()
                     Try
                         While Not File.Exists(run._marker)
                             Await Task.Delay(400, token)
                         End While
                         run.SignalEnded()
                     Catch ex As OperationCanceledException
                     End Try
                 End Function)
        Return run
    End Function

    ''' <summary>The terminal program and the words before the script: the user's, or the first one found.</summary>
    Private Shared Function TerminalWords(terminalCommand As String) As List(Of String)
        If Not String.IsNullOrWhiteSpace(terminalCommand) Then
            Return terminalCommand.Split(New Char() {" "c, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries).ToList()
        End If
        If HostPlatform.IsMacOS Then Return New List(Of String) From {"open"}
        For Each candidate In LinuxTerminals
            Dim found As String = HostPlatform.FindOnPath(candidate(0))
            If found IsNot Nothing Then
                Dim words As New List(Of String) From {found}
                For i As Integer = 1 To candidate.Length - 1
                    words.Add(candidate(i))
                Next
                Return words
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>A word in single quotes for /bin/sh.</summary>
    Private Shared Function Quote(word As String) As String
        Return "'" & If(word, "").Replace("'", "'\''") & "'"
    End Function

    Private Sub SignalEnded()
        If Interlocked.Exchange(_ended, 1) <> 0 Then Return
        RaiseEvent Exited(Me, EventArgs.Empty)
    End Sub

    Private Sub RemoveFolder()
        Try
            If _folder IsNot Nothing AndAlso Directory.Exists(_folder) Then Directory.Delete(_folder, True)
        Catch
        End Try
    End Sub

    ''' <summary>Stop following the run and remove its script. The program itself is left alone.</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            If _cancel IsNot Nothing Then _cancel.Cancel()
        Catch
        End Try
        Try
            If _process IsNot Nothing Then _process.Dispose()
        Catch
        End Try
        _process = Nothing
        RemoveFolder()
    End Sub
End Class
