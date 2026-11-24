' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Trs80Tool.vb — the engine adapter.
'
'  This is the ONLY file in TRSToolGui that knows trs80-tool's command line
'  and the shape of its output. Every other file asks this class for typed
'  results. When Lawrence Kesteloot changes trs80-tool, this is the file to
'  update.
'
'  Rules followed here (trs80-tool v5.1.x and v5.2.x):
'    * JSON everywhere it exists: dir --json, info --json --verbose,
'      sectors --json, help --json, and "disk IMAGE --json ..." for every
'      disk command. No text of a disk command is parsed; the text a command
'      would have printed comes back in the JSON ("messages") and is only
'      ever shown to the user.
'    * two transports, same JSON documents:
'        - one launch per call (RunAsync), or
'        - one long-running "trs80-tool mcp" server (EngineServer) that runs
'          dir / info / sectors / disk / disk_write in its own process, so a
'          call costs no program start. Any server trouble falls back to
'          launches for the rest of the session.
'    * program options (--color off --no-pager) come BEFORE the command word
'    * changes go through "disk IMAGE --json [--partition N] CMD ... : CMD ..."
'      (or the server's disk_write); nothing is saved if a command fails
'    * failures: the exit status says why (3 not found, 4 read-only, 5 write-
'      protected, 6 damaged, 7 disk full, 1 anything else); with --json the
'      error text and its code are inside the document
'    * hexdump, disasm, asm, convert and the text reports of sectors are plain
'      launches whose text is shown as printed; console programs (the disk
'      shell, the emulator) are started with StartInConsole and never parsed
'    * nothing here is about Windows: the tool's file name, where it is looked
'      for and how a console run is started come from HostPlatform and
'      ConsoleRun
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>Why a trs80-tool call failed, from its exit status (trs80-tool 5.1.0 and later).</summary>
Public Enum ToolFailure
    None = 0
    General = 1
    NotFound = 3
    ReadOnlyDisk = 4
    WriteProtected = 5
    Damaged = 6
    DiskFull = 7
End Enum

''' <summary>Result of one trs80-tool call.</summary>
Public Class ToolResult
    Public Property ExitCode As Integer
    Public Property StdOut As String = ""
    Public Property StdErr As String = ""
    Public Property StdOutBytes As Byte() = Nothing
    Public Property TimedOut As Boolean
    Public Property CommandLine As String = ""
    ''' <summary>The parsed document of a "disk --json" run, else Nothing.</summary>
    Public Property Disk As Trs80Tool.DiskRunJson = Nothing
    ''' <summary>The error text taken from a JSON document, when there is one.</summary>
    Public Property ErrorMessage As String = Nothing

    Public ReadOnly Property Succeeded As Boolean
        Get
            Return ExitCode = 0 AndAlso Not TimedOut
        End Get
    End Property

    ''' <summary>Why the call failed, from the exit status.</summary>
    Public ReadOnly Property Failure As ToolFailure
        Get
            If Succeeded Then Return ToolFailure.None
            Select Case ExitCode
                Case 3 : Return ToolFailure.NotFound
                Case 4 : Return ToolFailure.ReadOnlyDisk
                Case 5 : Return ToolFailure.WriteProtected
                Case 6 : Return ToolFailure.Damaged
                Case 7 : Return ToolFailure.DiskFull
            End Select
            Return ToolFailure.General
        End Get
    End Property

    ''' <summary>The message to show a user when the call failed.</summary>
    Public ReadOnly Property ErrorText As String
        Get
            If TimedOut Then Return "trs80-tool did not finish in time and was stopped." & vbCrLf & CommandLine
            If Not String.IsNullOrWhiteSpace(ErrorMessage) Then Return ErrorMessage.Trim()
            Dim s As String = StdOut.Trim()
            If s.Length = 0 Then s = WithoutStackTrace(StdErr).Trim()
            If s.Length = 0 Then s = "trs80-tool exited with code " & ExitCode.ToString(CultureInfo.InvariantCulture) & "."
            Return s
        End Get
    End Property

    ''' <summary>
    ''' An unexpected exception in trs80-tool prints a Node stack trace to stderr ("Error: ..." then "    at ..."
    ''' lines). Keep the message, drop the frames.
    ''' </summary>
    Private Shared Function WithoutStackTrace(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Dim kept As New List(Of String)
        For Each line In text.Replace(vbCr, "").Split(ControlChars.Lf)
            If line.TrimStart().StartsWith("at ", StringComparison.Ordinal) AndAlso line.StartsWith(" ", StringComparison.Ordinal) Then Exit For
            kept.Add(line)
        Next
        Return String.Join(vbCrLf, kept)
    End Function
End Class

''' <summary>Thrown by the typed wrappers when trs80-tool reports a failure.</summary>
Public Class ToolException
    Inherits Exception
    Public Property Result As ToolResult
    Public Sub New(result As ToolResult)
        MyBase.New(result.ErrorText)
        Me.Result = result
    End Sub
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
    ''' <summary>Why the call failed (General when there is no result to tell from).</summary>
    Public ReadOnly Property Failure As ToolFailure
        Get
            If Result Is Nothing Then Return ToolFailure.General
            Return Result.Failure
        End Get
    End Property
End Class

''' <summary>What the startup probe ("--version" and "help --json") found out about the installed trs80-tool.</summary>
Public Class EngineCapabilities
    Public Property ToolPath As String = ""
    Public Property VersionText As String = ""
    Public Property Version As Version = Nothing
    Public Property GitHash As String = ""
    Public Property TopLevelCommands As New List(Of String)
    Public Property DiskCommands As New List(Of String)
    ''' <summary>Option names ("--json") by top-level command name.</summary>
    Public Property TopLevelOptions As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    ''' <summary>Option names ("--text") by disk command name.</summary>
    Public Property DiskOptions As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    ''' <summary>Options for a whole disk run ("--json", "--partition").</summary>
    Public Property DiskGlobalOptions As New List(Of String)
    ''' <summary>The DOSes "disk create" and "disk format" take.</summary>
    Public Property FormatDoses As New List(Of Trs80Tool.HelpDosJson)
    ''' <summary>The DOSes whose disks "disk" can write to.</summary>
    Public Property WritableDoses As New List(Of Trs80Tool.HelpDosJson)
    ''' <summary>The system disks "disk create --system" can make.</summary>
    Public Property SystemDisks As New List(Of Trs80Tool.HelpSystemDiskJson)
    Public Property ReadFormats As New List(Of String)
    Public Property ConvertFormats As New List(Of String)
    Public Property DiskWriteFormats As New List(Of String)
    Public Property DiskCreateFormats As New List(Of String)
    Public Property Problems As New List(Of String)

    Public Function HasTopLevel(name As String) As Boolean
        Return TopLevelCommands.Contains(name, StringComparer.OrdinalIgnoreCase)
    End Function
    Public Function HasDisk(name As String) As Boolean
        Return DiskCommands.Contains(name, StringComparer.OrdinalIgnoreCase)
    End Function
    ''' <summary>True when the top-level command has the option ("sectors", "--protection").</summary>
    Public Function HasTopLevelOption(command As String, optionName As String) As Boolean
        Dim list As List(Of String) = Nothing
        Return TopLevelOptions.TryGetValue(command, list) AndAlso list.Contains(optionName, StringComparer.OrdinalIgnoreCase)
    End Function
    ''' <summary>True when the disk command has the option ("add", "--text").</summary>
    Public Function HasDiskOption(command As String, optionName As String) As Boolean
        Dim list As List(Of String) = Nothing
        Return DiskOptions.TryGetValue(command, list) AndAlso list.Contains(optionName, StringComparer.OrdinalIgnoreCase)
    End Function
    Public Function CanCreate(formatName As String) As Boolean
        Return DiskCreateFormats.Contains(formatName, StringComparer.OrdinalIgnoreCase)
    End Function
    ''' <summary>True when the probed trs80-tool is this version or later.</summary>
    Public Function AtLeast(minimum As Version) As Boolean
        Return Version IsNot Nothing AndAlso Version >= minimum
    End Function
    Public ReadOnly Property IsUsable As Boolean
        Get
            Return Version IsNot Nothing AndAlso Version >= Trs80Tool.MinimumVersion
        End Get
    End Property
End Class

Public Class Trs80Tool

    ' ---- version policy (update when a new trs80-tool is tested) ----
    Public Shared ReadOnly MinimumVersion As New Version(5, 1, 0)
    Public Shared ReadOnly MaxTestedVersion As New Version(5, 3, 0)
    ''' <summary>
    ''' The first trs80-tool that writes CP/M disks (and makes them with "create cpm"), reads and writes gzipped
    ''' files, and maps hard disk partitions. "help --json" doesn't list any of the three, so the GUI goes by this.
    ''' </summary>
    Public Shared ReadOnly Version520 As New Version(5, 2, 0)
    ''' <summary>
    ''' The first trs80-tool whose "disk" command has "import" (copy files from another floppy, hard disk,
    ''' CAS or WAV image onto the open disk) and "--keep-going" on "add". Both are in "help --json", so the
    ''' GUI asks the probe as well: see CanImport and CanKeepGoing.
    ''' </summary>
    Public Shared ReadOnly Version530 As New Version(5, 3, 0)

    ''' <summary>The name trs80-tool is released under on this system: trs80-tool.exe on Windows, trs80-tool elsewhere.</summary>
    Public Shared ReadOnly Property DefaultExeName As String
        Get
            Return HostPlatform.ToolFileName
        End Get
    End Property

    ''' <summary>True when the probed trs80-tool is 5.2.0 or later (False before the probe has run).</summary>
    Public ReadOnly Property Has520 As Boolean
        Get
            Return Capabilities IsNot Nothing AndAlso Capabilities.AtLeast(Version520)
        End Get
    End Property

    ''' <summary>True when the probed trs80-tool is 5.3.0 or later (False before the probe has run).</summary>
    Public ReadOnly Property Has530 As Boolean
        Get
            Return Capabilities IsNot Nothing AndAlso Capabilities.AtLeast(Version530)
        End Get
    End Property

    ''' <summary>
    ''' True when "disk import" copies files from another image. Both halves of the test matter. Up to
    ''' 5.2.0 "import" was an ALIAS of "add", so on a 5.2.0 tool "import OTHER.DMK" means "add OTHER.DMK"
    ''' and would store the source image as a file on the open disk. ToolCapabilities fills DiskCommands
    ''' from the command NAMES in "help --json" and never from their aliases, so HasDisk("import") is
    ''' already False there; the version test is what keeps that safe if the probe ever learns aliases.
    ''' </summary>
    Public ReadOnly Property CanImport As Boolean
        Get
            Return Has530 AndAlso Capabilities.HasDisk("import")
        End Get
    End Property

    ''' <summary>
    ''' True when "add" (and "import") take "--keep-going": skip a file that can't be added, with a warning,
    ''' and add the rest, instead of adding nothing. New in 5.3.0.
    ''' </summary>
    Public ReadOnly Property CanKeepGoing As Boolean
        Get
            Return Capabilities IsNot Nothing AndAlso Capabilities.HasDiskOption("add", "--keep-going")
        End Get
    End Property

    ' Top-level commands and disk sub-commands this GUI relies on (for the capability probe). "import" is
    ' deliberately absent: it arrived in 5.3.0 and the minimum this GUI needs is 5.1.0, so demanding it would
    ' report "disk command missing: import" on a perfectly good 5.1.0 or 5.2.0 tool. CanImport gates it instead.
    Public Shared ReadOnly ExpectedTopLevel As String() = {"dir", "info", "convert", "extract", "disk", "sectors", "hexdump", "disasm", "asm", "run", "mcp", "help"}
    Public Shared ReadOnly ExpectedDisk As String() = {"dir", "extract", "list", "add", "kill", "rename", "attrib", "label", "protect", "info", "partition", "free", "map", "check", "repair", "create", "format", "trim", "boot"}

    Private _toolPath As String = ""
    Private _logFile As String = ""
    Private _useServer As Boolean = True

    ''' <summary>Where trs80-tool is. Changing it stops the background server.</summary>
    Public Property ToolPath As String
        Get
            Return _toolPath
        End Get
        Set(value As String)
            If Not String.Equals(_toolPath, value, HostPlatform.PathComparison) Then StopServer()
            _toolPath = If(value, "")
        End Set
    End Property

    Public Property Capabilities As EngineCapabilities = Nothing
    Public Property TimeoutMs As Integer = 180000

    ''' <summary>When not empty, "--log-file PATH" is added to every call (program option, before the command word).</summary>
    Public Property LogFile As String
        Get
            Return _logFile
        End Get
        Set(value As String)
            If Not String.Equals(_logFile, If(value, ""), HostPlatform.PathComparison) Then StopServer()
            _logFile = If(value, "")
        End Set
    End Property

    ''' <summary>
    ''' Keep one trs80-tool running ("trs80-tool mcp") and send it the dir / info / sectors / disk calls, instead
    ''' of starting the program for each one. Turned off automatically for the session if the server misbehaves.
    ''' </summary>
    Public Property UseServer As Boolean
        Get
            Return _useServer
        End Get
        Set(value As Boolean)
            If Not value Then StopServer()
            _useServer = value
            If value Then _serverFailed = False
        End Set
    End Property

    ''' <summary>
    ''' Linux and macOS: the terminal program for the disk shell and the emulator, with the words that go before
    ''' the script (for example "xterm -e"). Blank = find one. Not used on Windows.
    ''' </summary>
    Public Property TerminalCommand As String = ""

    ' ---- CP/M layout chosen by the user, per image (trs80-tool 5.2.0: --cpm-format NAME) ----
    ' A CP/M disk doesn't say how it is laid out, so trs80-tool works it out, and can be told instead. The name
    ' given here goes on every call that reads or changes that image, for as long as the program runs.
    Private ReadOnly _cpmFormats As New Dictionary(Of String, String)(If(HostPlatform.IsLinux, StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase))

    ''' <summary>Read this image as the named CP/M layout from now on; Nothing or "" goes back to trs80-tool's own choice.</summary>
    Public Sub SetCpmFormat(imagePath As String, formatName As String)
        Dim key As String = Path.GetFullPath(imagePath)
        If String.IsNullOrWhiteSpace(formatName) Then
            _cpmFormats.Remove(key)
        Else
            _cpmFormats(key) = formatName.Trim()
        End If
    End Sub

    ''' <summary>The CP/M layout the user chose for this image, or "" when trs80-tool decides.</summary>
    Public Function CpmFormatOf(imagePath As String) As String
        Dim name As String = Nothing
        If Not String.IsNullOrEmpty(imagePath) AndAlso _cpmFormats.TryGetValue(Path.GetFullPath(imagePath), name) Then Return name
        Return ""
    End Function

    ''' <summary>"disk IMAGE help cpm-format": trs80-tool's own table of the CP/M layouts and their names, as it prints it.</summary>
    Public Async Function CpmFormatHelpAsync(imagePath As String) As Task(Of String)
        Dim r As ToolResult = Await RunAsync(New String() {"disk", imagePath, "help", "cpm-format"})
        Return If(r.Succeeded, r.StdOut, r.ErrorText)
    End Function

    ''' <summary>Log of every call (for the Engine Info dialog).</summary>
    Public ReadOnly Property CallLog As New List(Of String)

    Private Shared ReadOnly JsonOpts As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True, .AllowTrailingCommas = True}
    Private Shared ReadOnly Utf8NoBom As New UTF8Encoding(False)

    Public Sub New(toolPath As String)
        _toolPath = If(toolPath, "")
    End Sub

    ' ------------------------------------------------------------------ locating

    ''' <summary>Ira's own (renamed) copy of trs80-tool. Used when no location is configured.</summary>
    Public Const DefaultInstallPath As String = "C:\DOS\TRS\castool.exe"

    ''' <summary>
    ''' Find trs80-tool: the configured path (TRSToolGui.cfg), then a trs80-tool in this program's own folder
    ''' (on a Mac also the folder that holds TRSToolGui.app), then (Windows) DefaultInstallPath, then the PATH,
    ''' then (Mac) the usual folders for command-line tools. Returns Nothing if none exists; the program then
    ''' asks the user where it is.
    ''' </summary>
    Public Shared Function Locate(configuredPath As String) As String
        If Not String.IsNullOrWhiteSpace(configuredPath) AndAlso File.Exists(configuredPath) Then Return Path.GetFullPath(configuredPath)
        Dim beside As String = Path.Combine(AppContext.BaseDirectory, DefaultExeName)
        If File.Exists(beside) Then Return beside
        Dim outerFolder As String = HostPlatform.ProgramFolder
        If Not HostPlatform.SamePath(outerFolder, AppContext.BaseDirectory) Then
            beside = Path.Combine(outerFolder, DefaultExeName)
            If File.Exists(beside) Then Return beside
        End If
        If HostPlatform.IsWindows AndAlso File.Exists(DefaultInstallPath) Then Return DefaultInstallPath
        Dim onPath As String = HostPlatform.FindOnPath(DefaultExeName)
        If onPath IsNot Nothing Then Return onPath
        For Each folder In HostPlatform.ExtraToolFolders
            Dim candidate As String = Path.Combine(folder, DefaultExeName)
            If File.Exists(candidate) Then Return candidate
        Next
        Return Nothing
    End Function

    ' ------------------------------------------------------------------ running (one launch per call)

    Private Sub AddProgramOptions(psi As ProcessStartInfo, color As Boolean)
        ' Never let the tool page or colour: these are program-level options and go first.
        If Not color Then
            psi.ArgumentList.Add("--color")
            psi.ArgumentList.Add("off")
        End If
        psi.ArgumentList.Add("--no-pager")
        If Not String.IsNullOrWhiteSpace(LogFile) Then
            psi.ArgumentList.Add("--log-file")
            psi.ArgumentList.Add(LogFile)
        End If
    End Sub

    Private Sub ThrowIfNoTool()
        If String.IsNullOrWhiteSpace(ToolPath) OrElse Not File.Exists(ToolPath) Then
            Throw New ToolException("trs80-tool was not found. Set its location in Edit → Preferences.")
        End If
    End Sub

    ''' <summary>
    ''' Run trs80-tool with the given arguments (program options are added automatically).
    ''' </summary>
    ''' <param name="args">The command word and its arguments.</param>
    ''' <param name="stdinText">Text to feed on standard input (a disk script), or Nothing.</param>
    ''' <param name="binaryStdout">Capture stdout as raw bytes (for "list --raw").</param>
    Public Async Function RunAsync(args As IEnumerable(Of String), Optional stdinText As String = Nothing, Optional binaryStdout As Boolean = False, Optional timeoutMs As Integer = 0) As Task(Of ToolResult)
        ThrowIfNoTool()
        If timeoutMs <= 0 Then timeoutMs = Me.TimeoutMs
        Dim psi As New ProcessStartInfo(ToolPath) With {
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .RedirectStandardInput = stdinText IsNot Nothing,
            .CreateNoWindow = True,
            .StandardErrorEncoding = Encoding.UTF8
        }
        If Not binaryStdout Then psi.StandardOutputEncoding = Encoding.UTF8
        ' A script names host files, which may have any characters: send it as UTF-8, as Node reads it.
        If stdinText IsNot Nothing Then psi.StandardInputEncoding = Utf8NoBom
        AddProgramOptions(psi, False)
        For Each a In args
            psi.ArgumentList.Add(a)
        Next
        ' Belt and braces for chalk/pager detection inside Node.
        psi.Environment.Item("NO_COLOR") = "1"
        psi.Environment.Item("TRS80_TOOL_PAGER") = ""
        psi.Environment.Item("PAGER") = ""

        Dim result As New ToolResult With {.CommandLine = DescribeCommandLine(args)}
        LogCall(result.CommandLine)

        Using p As Process = Process.Start(psi)
            Dim errTask As Task(Of String) = p.StandardError.ReadToEndAsync()
            Dim outTextTask As Task(Of String) = Nothing
            Dim outBytesTask As Task = Nothing
            Dim ms As MemoryStream = Nothing
            If binaryStdout Then
                ms = New MemoryStream()
                outBytesTask = p.StandardOutput.BaseStream.CopyToAsync(ms)
            Else
                outTextTask = p.StandardOutput.ReadToEndAsync()
            End If
            If stdinText IsNot Nothing Then
                Try
                    Await p.StandardInput.WriteAsync(stdinText)
                    p.StandardInput.Close()
                Catch ex As IOException
                    ' The tool may exit before reading everything; the exit code tells the story.
                End Try
            End If
            Using cts As New CancellationTokenSource(timeoutMs)
                Try
                    Await p.WaitForExitAsync(cts.Token)
                Catch ex As OperationCanceledException
                    result.TimedOut = True
                    Try
                        p.Kill(True)
                    Catch
                    End Try
                End Try
            End Using
            If binaryStdout Then
                Await outBytesTask
                result.StdOutBytes = ms.ToArray()
                result.StdOut = ""
            Else
                result.StdOut = Await outTextTask
            End If
            result.StdErr = Await errTask
            If Not result.TimedOut Then result.ExitCode = p.ExitCode Else result.ExitCode = -1
        End Using
        Return result
    End Function

    Private Shared Function DescribeCommandLine(args As IEnumerable(Of String)) As String
        Dim sb As New StringBuilder("trs80-tool --color off --no-pager")
        For Each a In args
            sb.Append(" "c)
            If a.Length = 0 OrElse a.IndexOfAny(New Char() {" "c, """"c}) >= 0 Then
                sb.Append(""""c).Append(a.Replace("""", "\""")).Append(""""c)
            Else
                sb.Append(a)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Sub LogCall(line As String)
        SyncLock CallLog
            CallLog.Add(DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) & "  " & line)
            If CallLog.Count > 400 Then CallLog.RemoveRange(0, CallLog.Count - 400)
        End SyncLock
    End Sub

    Private Shared Sub ThrowIfFailed(r As ToolResult)
        If Not r.Succeeded Then Throw New ToolException(r)
    End Sub

    ' ------------------------------------------------------------------ running (the long-running server)

    Private _server As Process = Nothing
    Private _serverFailed As Boolean = False
    Private _serverNextId As Integer = 1
    Private ReadOnly _serverLock As New SemaphoreSlim(1, 1)

    ''' <summary>How long the server may take to start and answer "initialize".</summary>
    Private Const ServerStartTimeoutMs As Integer = 30000

    ''' <summary>True while the background server is running.</summary>
    Public ReadOnly Property ServerRunning As Boolean
        Get
            Dim p As Process = _server
            If p Is Nothing Then Return False
            Try
                Return Not p.HasExited
            Catch
                Return False
            End Try
        End Get
    End Property

    ''' <summary>A line for the Engine Info dialog saying how calls are being made.</summary>
    Public ReadOnly Property TransportText As String
        Get
            If Not _useServer Then Return "one program start per call (the background server is turned off in Preferences)"
            If _serverFailed Then Return "one program start per call (the background server failed earlier in this session)"
            If Capabilities IsNot Nothing AndAlso Not Capabilities.HasTopLevel("mcp") Then Return "one program start per call (this trs80-tool has no ""mcp"" command)"
            If ServerRunning Then Return "background server running (trs80-tool mcp): disk calls cost no program start"
            Return "background server, started at the first disk call"
        End Get
    End Property

    Private Function ServerWanted() As Boolean
        If Not _useServer OrElse _serverFailed Then Return False
        If Capabilities Is Nothing OrElse Not Capabilities.IsUsable Then Return False
        Return Capabilities.HasTopLevel("mcp")
    End Function

    ''' <summary>Stop the background server, if it is running. Safe to call at any time.</summary>
    Public Sub StopServer()
        Dim p As Process = _server
        _server = Nothing
        If p Is Nothing Then Return
        Try
            If Not p.HasExited Then
                Try
                    p.StandardInput.Close()      ' end of input makes the server leave by itself
                Catch
                End Try
                If Not p.WaitForExit(500) Then p.Kill(True)
            End If
        Catch
        End Try
        Try
            p.Dispose()
        Catch
        End Try
    End Sub

    ''' <summary>Call when the program ends.</summary>
    Public Sub Shutdown()
        StopServer()
    End Sub

    Private Sub ServerGiveUp(reason As String)
        _serverFailed = True
        LogCall("[server] turned off for this session: " & reason)
        StopServer()
    End Sub

    ''' <summary>One line from the server, or Nothing at end of stream or after the time allowed.</summary>
    Private Shared Async Function ReadLineWithTimeoutAsync(reader As StreamReader, timeoutMs As Integer) As Task(Of String)
        Dim readTask As Task(Of String) = reader.ReadLineAsync()
        Dim finished As Task = Await Task.WhenAny(readTask, Task.Delay(timeoutMs))
        If finished IsNot readTask Then Return Nothing
        Return Await readTask
    End Function

    ''' <summary>Read replies until the one with this id; Nothing if the server stops answering.</summary>
    Private Shared Async Function ReadReplyAsync(p As Process, id As Integer, timeoutMs As Integer) As Task(Of String)
        Dim watch As Stopwatch = Stopwatch.StartNew()
        Do
            Dim remaining As Integer = timeoutMs - CInt(Math.Min(watch.ElapsedMilliseconds, Integer.MaxValue \ 2))
            If remaining <= 0 Then Return Nothing
            Dim line As String = Await ReadLineWithTimeoutAsync(p.StandardOutput, remaining)
            If line Is Nothing Then Return Nothing
            If line.Length = 0 OrElse line(0) <> "{"c Then Continue Do
            Try
                Using doc As JsonDocument = JsonDocument.Parse(line)
                    Dim idEl As JsonElement
                    If doc.RootElement.TryGetProperty("id", idEl) AndAlso idEl.ValueKind = JsonValueKind.Number AndAlso idEl.GetInt32() = id Then Return line
                End Using
            Catch ex As JsonException
                ' Not a reply; keep reading.
            End Try
        Loop
    End Function

    ''' <summary>Start the server and shake hands. Must be called with _serverLock held. False = could not.</summary>
    Private Async Function EnsureServerAsync() As Task(Of Boolean)
        If ServerRunning Then Return True
        StopServer()
        Try
            Dim psi As New ProcessStartInfo(ToolPath) With {
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .RedirectStandardInput = True,
                .CreateNoWindow = True,
                .StandardOutputEncoding = Encoding.UTF8,
                .StandardErrorEncoding = Encoding.UTF8,
                .StandardInputEncoding = Utf8NoBom
            }
            AddProgramOptions(psi, False)
            psi.ArgumentList.Add("mcp")
            psi.Environment.Item("NO_COLOR") = "1"
            psi.Environment.Item("TRS80_TOOL_PAGER") = ""
            psi.Environment.Item("PAGER") = ""
            LogCall("[server] trs80-tool --color off --no-pager mcp")
            Dim p As Process = Process.Start(psi)
            If p Is Nothing Then Return False
            ' The server logs to stderr; drain it so it can never block on a full pipe.
            AddHandler p.ErrorDataReceived, Sub(sender, e)
                                            End Sub
            p.BeginErrorReadLine()
            _server = p

            Dim id As Integer = _serverNextId
            _serverNextId += 1
            Dim init As New Dictionary(Of String, Object) From {
                {"jsonrpc", "2.0"}, {"id", id}, {"method", "initialize"},
                {"params", New Dictionary(Of String, Object) From {
                    {"protocolVersion", "2024-11-05"},
                    {"capabilities", New Dictionary(Of String, Object)()},
                    {"clientInfo", New Dictionary(Of String, Object) From {{"name", "TRSToolGui"}, {"version", "1.1"}}}}}}
            Await p.StandardInput.WriteLineAsync(JsonSerializer.Serialize(init))
            Await p.StandardInput.FlushAsync()
            Dim reply As String = Await ReadReplyAsync(p, id, ServerStartTimeoutMs)
            If reply Is Nothing Then Return False
            Dim note As New Dictionary(Of String, Object) From {{"jsonrpc", "2.0"}, {"method", "notifications/initialized"}}
            Await p.StandardInput.WriteLineAsync(JsonSerializer.Serialize(note))
            Await p.StandardInput.FlushAsync()
            Return True
        Catch ex As Exception
            LogCall("[server] could not start: " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>What a server call came back with.</summary>
    Private Class ServerReply
        ''' <summary>The server answered (whatever the answer says).</summary>
        Public Answered As Boolean
        ''' <summary>The request had been sent when the server stopped answering, so it may have run.</summary>
        Public MayHaveRun As Boolean
        Public Text As String = ""
        Public IsError As Boolean
    End Class

    ''' <summary>
    ''' Call one tool of the server. Answered = False means the transport failed (the server is then turned off
    ''' for the session); the caller decides whether the call can be repeated as a launch.
    ''' </summary>
    Private Async Function ServerCallAsync(toolName As String, arguments As Dictionary(Of String, Object), describe As String, timeoutMs As Integer) As Task(Of ServerReply)
        Dim result As New ServerReply
        Await _serverLock.WaitAsync()
        Try
            If Not Await EnsureServerAsync() Then
                ServerGiveUp("it did not start")
                Return result
            End If
            Dim p As Process = _server
            Dim id As Integer = _serverNextId
            _serverNextId += 1
            Dim request As New Dictionary(Of String, Object) From {
                {"jsonrpc", "2.0"}, {"id", id}, {"method", "tools/call"},
                {"params", New Dictionary(Of String, Object) From {{"name", toolName}, {"arguments", arguments}}}}
            LogCall("[server] " & describe)
            Dim line As String = Nothing
            Try
                Await p.StandardInput.WriteLineAsync(JsonSerializer.Serialize(request))
                Await p.StandardInput.FlushAsync()
                result.MayHaveRun = True
                line = Await ReadReplyAsync(p, id, timeoutMs)
            Catch ex As Exception
                line = Nothing
            End Try
            If line Is Nothing Then
                ServerGiveUp("no answer to " & toolName)
                Return result
            End If
            Try
                Using doc As JsonDocument = JsonDocument.Parse(line)
                    Dim root As JsonElement = doc.RootElement
                    Dim res As JsonElement
                    If Not root.TryGetProperty("result", res) Then
                        ' A protocol error (unknown tool, say): this server can't do what we ask.
                        Dim errEl As JsonElement
                        Dim msg As String = "protocol error"
                        If root.TryGetProperty("error", errEl) Then
                            Dim m As JsonElement
                            If errEl.ValueKind = JsonValueKind.Object AndAlso errEl.TryGetProperty("message", m) Then msg = If(m.GetString(), msg)
                        End If
                        ServerGiveUp(msg)
                        result.MayHaveRun = False
                        Return result
                    End If
                    Dim content As JsonElement
                    Dim sb As New StringBuilder()
                    If res.TryGetProperty("content", content) AndAlso content.ValueKind = JsonValueKind.Array Then
                        For Each item As JsonElement In content.EnumerateArray()
                            Dim t As JsonElement
                            If item.TryGetProperty("text", t) AndAlso t.ValueKind = JsonValueKind.String Then sb.Append(t.GetString())
                        Next
                    End If
                    Dim isErr As JsonElement
                    result.IsError = res.TryGetProperty("isError", isErr) AndAlso isErr.ValueKind = JsonValueKind.True
                    result.Text = sb.ToString()
                    result.Answered = True
                End Using
            Catch ex As JsonException
                ServerGiveUp("an answer that isn't JSON")
            End Try
            Return result
        Finally
            _serverLock.Release()
        End Try
    End Function

    ''' <summary>
    ''' Get one of the tool's JSON documents, by the server when it is in use, else by a launch.
    ''' </summary>
    ''' <param name="serverTool">The server's tool name ("dir", "info", "sectors", "disk", "disk_write").</param>
    ''' <param name="serverArgs">That tool's arguments.</param>
    ''' <param name="cliArgs">The same call as command-line words.</param>
    ''' <param name="changes">True when the call may change a file: it is then never repeated after a doubtful failure.</param>
    ''' <param name="stdinText">A disk script for the launch, or Nothing.</param>
    Private Async Function JsonCallAsync(serverTool As String, serverArgs As Dictionary(Of String, Object), cliArgs As List(Of String), changes As Boolean, Optional stdinText As String = Nothing, Optional timeoutMs As Integer = 0) As Task(Of ToolResult)
        ThrowIfNoTool()
        If timeoutMs <= 0 Then timeoutMs = Me.TimeoutMs
        If ServerWanted() Then
            Dim describe As String = DescribeCommandLine(cliArgs)
            Dim reply As ServerReply = Await ServerCallAsync(serverTool, serverArgs, describe, timeoutMs)
            If reply.Answered Then
                Dim r As New ToolResult With {.CommandLine = describe, .StdOut = reply.Text, .ExitCode = If(reply.IsError, 1, 0)}
                ' A tool that threw answers with plain text ("Error: ...") instead of a document.
                If reply.IsError AndAlso Not reply.Text.TrimStart().StartsWith("{", StringComparison.Ordinal) Then
                    Dim msg As String = reply.Text.Trim()
                    If msg.StartsWith("Error: ", StringComparison.Ordinal) Then msg = msg.Substring(7)
                    r.ErrorMessage = msg
                End If
                Return r
            End If
            If changes AndAlso reply.MayHaveRun Then
                Return New ToolResult With {.CommandLine = describe, .ExitCode = 1,
                    .ErrorMessage = "trs80-tool stopped answering while it was working on the image, so it is not known whether the change was made. Check the file list before trying again."}
            End If
            ' Nothing was changed: fall through and do it as a launch.
        End If
        Return Await RunAsync(cliArgs, stdinText, False, timeoutMs)
    End Function

    ' ------------------------------------------------------------------ version & probe

    ''' <summary>Parse "5.1.0 (git abc, built ...)" → 5.1.0. Returns Nothing when unparseable.</summary>
    Public Shared Function ParseVersion(versionOutput As String) As Version
        If versionOutput Is Nothing Then Return Nothing
        Dim m As Match = Regex.Match(versionOutput.Trim(), "^v?(\d+)\.(\d+)(?:\.(\d+))?")
        If Not m.Success Then Return Nothing
        Dim major As Integer = Integer.Parse(m.Groups(1).Value, CultureInfo.InvariantCulture)
        Dim minor As Integer = Integer.Parse(m.Groups(2).Value, CultureInfo.InvariantCulture)
        Dim build As Integer = 0
        If m.Groups(3).Success Then build = Integer.Parse(m.Groups(3).Value, CultureInfo.InvariantCulture)
        Return New Version(major, minor, build)
    End Function

    ' ---- help --json (what this version of the tool can do)

    Public Class HelpOptionJson
        Public Property Name As String
        Public Property ShortName As String
        Public Property Description As String
    End Class

    Public Class HelpCommandJson
        Public Property Name As String
        Public Property Aliases As List(Of String)
        Public Property Description As String
        Public Property Options As List(Of HelpOptionJson)
    End Class

    Public Class HelpDiskJson
        Public Property Options As List(Of HelpOptionJson)
        Public Property Commands As List(Of HelpCommandJson)
    End Class

    Public Class HelpDosJson
        ''' <summary>The name to give on the command line, such as "ldos-5.3".</summary>
        Public Property Name As String
        Public Property Aliases As List(Of String)
        ''' <summary>The DOS's name for people, such as "LDOS 5.3".</summary>
        Public Property Label As String
    End Class

    Public Class HelpSystemDiskJson
        Public Property Name As String
        ''' <summary>The DOS by the name "disk create" takes.</summary>
        Public Property Dos As String
        Public Property Model As Integer
        Public Property Version As String
        Public Property Label As String
    End Class

    Public Class HelpImageFormatsJson
        Public Property Read As List(Of String)
        Public Property Convert As List(Of String)
        Public Property DiskWrite As List(Of String)
        Public Property DiskCreate As List(Of String)
    End Class

    Public Class HelpJson
        Public Property Version As String
        Public Property GitHash As String
        Public Property Options As List(Of HelpOptionJson)
        Public Property Commands As List(Of HelpCommandJson)
        Public Property Disk As HelpDiskJson
        Public Property WritableDoses As List(Of HelpDosJson)
        Public Property FormatDoses As List(Of HelpDosJson)
        Public Property SystemDisks As List(Of HelpSystemDiskJson)
        Public Property ImageFormats As HelpImageFormatsJson
    End Class

    Private Shared Function OptionNames(options As List(Of HelpOptionJson)) As List(Of String)
        Dim names As New List(Of String)
        If options Is Nothing Then Return names
        For Each o In options
            If Not String.IsNullOrEmpty(o.Name) Then names.Add(o.Name)
        Next
        Return names
    End Function

    ''' <summary>
    ''' Run --version and "help --json" and record what the installed tool can do.
    ''' Never throws for a tool that runs; a missing/unrunnable tool gives Capabilities with Version = Nothing.
    ''' </summary>
    Public Async Function ProbeAsync() As Task(Of EngineCapabilities)
        Dim caps As New EngineCapabilities With {.ToolPath = ToolPath}
        Try
            Dim v As ToolResult = Await RunAsync(New String() {"--version"})
            caps.VersionText = (v.StdOut & v.StdErr).Trim()
            caps.Version = ParseVersion(caps.VersionText)
            If caps.Version Is Nothing Then caps.Problems.Add("Could not parse the version from: " & caps.VersionText)
        Catch ex As Exception
            caps.Problems.Add("Could not run trs80-tool: " & ex.Message)
            Capabilities = caps
            Return caps
        End Try
        If caps.Version IsNot Nothing AndAlso caps.Version < MinimumVersion Then
            caps.Problems.Add("trs80-tool " & caps.Version.ToString() & " is older than the minimum " & MinimumVersion.ToString() & " this program needs (""disk --json"" and ""help --json""). Download the current trs80-tool from https://www.my-trs-80.com/tool/.")
            Capabilities = caps
            Return caps
        End If
        Try
            Dim h As ToolResult = Await RunAsync(New String() {"help", "--json"})
            Dim help As HelpJson = ParseJson(Of HelpJson)(h)
            If help Is Nothing OrElse help.Commands Is Nothing Then
                caps.Problems.Add("""help --json"" did not answer with JSON: " & h.ErrorText)
            Else
                caps.GitHash = If(help.GitHash, "")
                For Each c In help.Commands
                    If String.IsNullOrEmpty(c.Name) Then Continue For
                    caps.TopLevelCommands.Add(c.Name)
                    caps.TopLevelOptions(c.Name) = OptionNames(c.Options)
                Next
                If help.Disk IsNot Nothing Then
                    caps.DiskGlobalOptions = OptionNames(help.Disk.Options)
                    If help.Disk.Commands IsNot Nothing Then
                        For Each c In help.Disk.Commands
                            If String.IsNullOrEmpty(c.Name) Then Continue For
                            caps.DiskCommands.Add(c.Name)
                            caps.DiskOptions(c.Name) = OptionNames(c.Options)
                        Next
                    End If
                End If
                If help.FormatDoses IsNot Nothing Then caps.FormatDoses = help.FormatDoses
                If help.WritableDoses IsNot Nothing Then caps.WritableDoses = help.WritableDoses
                If help.SystemDisks IsNot Nothing Then caps.SystemDisks = help.SystemDisks
                If help.ImageFormats IsNot Nothing Then
                    If help.ImageFormats.Read IsNot Nothing Then caps.ReadFormats = help.ImageFormats.Read
                    If help.ImageFormats.Convert IsNot Nothing Then caps.ConvertFormats = help.ImageFormats.Convert
                    If help.ImageFormats.DiskWrite IsNot Nothing Then caps.DiskWriteFormats = help.ImageFormats.DiskWrite
                    If help.ImageFormats.DiskCreate IsNot Nothing Then caps.DiskCreateFormats = help.ImageFormats.DiskCreate
                End If
            End If
        Catch ex As Exception
            caps.Problems.Add("""help --json"" failed: " & ex.Message)
        End Try
        For Each c In ExpectedTopLevel
            If caps.TopLevelCommands.Count > 0 AndAlso Not caps.HasTopLevel(c) Then caps.Problems.Add("Top-level command missing: " & c)
        Next
        For Each c In ExpectedDisk
            If caps.DiskCommands.Count > 0 AndAlso Not caps.HasDisk(c) Then caps.Problems.Add("disk command missing: " & c)
        Next
        If caps.DiskGlobalOptions.Count > 0 AndAlso Not caps.DiskGlobalOptions.Contains("--json") Then caps.Problems.Add("""disk"" has no --json option; this program needs it.")
        If caps.Version IsNot Nothing AndAlso caps.Version > MaxTestedVersion Then
            caps.Problems.Add("trs80-tool " & caps.Version.ToString() & " is newer than the last version this program was tested with (" & MaxTestedVersion.ToString() & "). It will probably work; check Help → Engine capabilities if something misbehaves.")
        End If
        Capabilities = caps
        Return caps
    End Function

    ' ------------------------------------------------------------------ JSON DTOs (trs80-tool's --json output)

    Public Class InfilesJson(Of T)
        Public Property Infiles As List(Of T)
    End Class

    Public Class DamageJson
        Public Property MissingSectors As List(Of SectorPositionJson)
        Public Property CrcErrorSectors As List(Of SectorPositionJson)
        Public Property DirectoryErrors As List(Of String)
        Public Overrides Function ToString() As String
            Dim parts As New List(Of String)
            If MissingSectors IsNot Nothing AndAlso MissingSectors.Count > 0 Then parts.Add(MissingSectors.Count.ToString(CultureInfo.InvariantCulture) & " unreadable sector(s)")
            If CrcErrorSectors IsNot Nothing AndAlso CrcErrorSectors.Count > 0 Then parts.Add(CrcErrorSectors.Count.ToString(CultureInfo.InvariantCulture) & " CRC error(s)")
            If DirectoryErrors IsNot Nothing AndAlso DirectoryErrors.Count > 0 Then parts.AddRange(DirectoryErrors)
            Return String.Join(", ", parts)
        End Function
    End Class

    Public Class SectorPositionJson
        Public Property Cylinder As Integer
        Public Property Side As Integer
        Public Property Sector As Integer
    End Class

    ''' <summary>A file's TRSDOS-family details in "dir --json" (json.ts, v5.1.0).</summary>
    Public Class TrsdosFileJson
        Public Property OsFile As Boolean
        Public Property Hidden As Boolean
        Public Property Protection As String
        ''' <summary>Hash of the update password; Nothing where the entry has none (TRSDOS 1.3 system files).</summary>
        Public Property UpdatePasswordHash As Integer?
        ''' <summary>Hash of the access password; Nothing where those bytes hold a time stamp.</summary>
        Public Property AccessPasswordHash As Integer?
        ''' <summary>"" for none or "PASSWORD" when the hash is one the tool can name; Nothing for any other password.</summary>
        Public Property UpdatePassword As String
        Public Property AccessPassword As String
        ''' <summary>The disk keeps a time stamp where the access password would go (LDOS 5.3, LS-DOS 6.3).</summary>
        Public Property TimeStampsReplaceAccessPassword As Boolean
        ''' <summary>Directory entry byte 0 as stored.</summary>
        Public Property EntryByte0 As Integer?
        ''' <summary>Directory entry byte 1 as stored (LDOS / TRSDOS 6: bit 7 CREATEd, bit 6 MOD).</summary>
        Public Property EntryByte1 As Integer?
    End Class

    Public Class CpmFileJson
        Public Property OsFile As Boolean
        Public Property Hidden As Boolean
        Public Property User As Integer
        Public Property Attributes As List(Of String)
    End Class

    Public Class GenericContainerJson
        Public Property OsFile As Boolean
        Public Property Hidden As Boolean
    End Class

    Public Class FileJson
        Public Property Name As String
        Public Property Size As Long
        Public Property Sha256 As String
        Public Property ContentType As String
        Public Property Description As String
        Public Property [Error] As String
        Public Property [Date] As String
        Public Property Damage As DamageJson
        Public Property Contents As String
        Public Property OutputName As String
        Public Property Trsdos As TrsdosFileJson
        Public Property Cpm As CpmFileJson
        Public Property MicroDos As GenericContainerJson
        Public Property UcsdPascal As GenericContainerJson
        Public Property ModelIITrsdos As GenericContainerJson
        Public Property TrsdosII As GenericContainerJson
        Public Property XenixTar As GenericContainerJson
    End Class

    Public Class FilesInfileJson
        Public Property Infile As String
        Public Property [Error] As String
        Public Property ErrorCode As Integer?
        Public Property Warnings As List(Of String)
        Public Property Os As String
        Public Property Files As List(Of FileJson)
    End Class

    Public Class TrackGeometryInfoJson
        Public Property FirstSector As Integer
        Public Property LastSector As Integer
        Public Property Density As String
        Public Property SectorSize As Integer
    End Class

    Public Class DataTrackInfoJson
        Public Property SectorSpan As Integer
        Public Property Density As String
        Public Property SectorSize As Integer
    End Class

    Public Class GeometryInfoJson
        Public Property CylinderCount As Integer
        Public Property SideCount As Integer
        Public Property BootTrack As TrackGeometryInfoJson
        Public Property DataTrack As DataTrackInfoJson
        Public Property Homogeneous As Boolean
    End Class

    ''' <summary>NEWDOS/80 PDRIVE settings as "info --json --verbose" gives them (v5.1.0).</summary>
    Public Class PdriveJson
        ''' <summary>"table" (read from the disk's drive table) or "layout" (worked out from how the disk was decoded).</summary>
        Public Property Source As String
        ''' <summary>Which cylinder's sector 2 held the table (0, or 1 when cylinder 0's was blank); table only.</summary>
        Public Property TableCylinder As Integer?
        ''' <summary>The line as PDRIVE takes it: TI=A,TD=E,TC=40,SPT=18,TSR=3,GPL=2,DDSL=17,DDGA=2.</summary>
        Public Property Text As String
    End Class

    Public Class TrsdosInfoJson
        Public Property DiskName As String
        Public Property DiskDate As String
        Public Property AutoCommand As String
        Public Property GatPassword As Integer?
        ''' <summary>"" for no master password or "PASSWORD"; Nothing when it is some other password.</summary>
        Public Property GatPasswordText As String
        Public Property GatError As String
        Public Property AllocationSides As Integer
        Public Property DirLump As Integer
        Public Property LumpTerm As String
        Public Property SectorsPerGranule As Integer
        Public Property GranulesPerLump As Integer
        Public Property Pdrive As PdriveJson
    End Class

    Public Class CpmInfoJson
        Public Property Format As String
        Public Property AllocationSides As Integer
        Public Property CylinderCount As Integer
        Public Property ReservedTrackCount As Integer
        Public Property SectorSize As Integer
        Public Property Skew As Integer
        Public Property BlockSize As Integer
        Public Property BlockCount As Integer
        Public Property DirEntryCount As Integer
    End Class

    Public Class OsInfoJson
        Public Property Name As String
        Public Property FileCount As Integer
        Public Property Trsdos As TrsdosInfoJson
        Public Property Cpm As CpmInfoJson
    End Class

    Public Class FloppyInfoJson
        Public Property Format As String
        Public Property Os As OsInfoJson
        Public Property Geometry As GeometryInfoJson
    End Class

    Public Class PartitionInfoJson
        Public Property Number As Integer
        Public Property CylinderOffset As Integer
        Public Property HeadOffset As Integer
        Public Property CylinderCount As Integer
        Public Property HeadCount As Integer
        Public Property Os As OsInfoJson
    End Class

    Public Class HardDiskInfoJson
        Public Property CylinderCount As Integer
        Public Property HeadCount As Integer
        Public Property Partitions As List(Of PartitionInfoJson)
    End Class

    Public Class CassetteInfoJson
        Public Property FileCount As Integer
    End Class

    Public Class InfoJson
        Public Property Infile As String
        Public Property [Error] As String
        Public Property ErrorCode As Integer?
        Public Property Warnings As List(Of String)
        Public Property Size As Long?
        Public Property Sha256 As String
        Public Property ContentType As String
        Public Property Description As String
        Public Property Floppy As FloppyInfoJson
        Public Property HardDisk As HardDiskInfoJson
        Public Property Cassette As CassetteInfoJson
    End Class

    ' ---- "disk --json" (disk/DiskJson.ts, v5.1.0)

    ''' <summary>One command of a disk run.</summary>
    Public Class DiskCommandJson
        ''' <summary>The command as typed.</summary>
        Public Property Command As String
        ''' <summary>The command's name without aliases.</summary>
        Public Property Name As String
        ''' <summary>What the command would have printed, line by line. For people.</summary>
        Public Property Messages As List(Of String)
        Public Property Warnings As List(Of String)
        Public Property [Error] As String
        Public Property ErrorCode As Integer?
        ''' <summary>The command's own result; its shape depends on Name.</summary>
        Public Property Result As JsonElement

        ''' <summary>The result as a typed object, or Nothing when the command failed or has none.</summary>
        Public Function GetResult(Of T As Class)() As T
            If Result.ValueKind <> JsonValueKind.Object Then Return Nothing
            Try
                Return JsonSerializer.Deserialize(Of T)(Result.GetRawText(), JsonOpts)
            Catch ex As JsonException
                Return Nothing
            End Try
        End Function
    End Class

    ''' <summary>The whole of what "disk --json" prints.</summary>
    Public Class DiskRunJson
        Public Property Image As String
        Public Property Commands As List(Of DiskCommandJson)
        ''' <summary>The path the changes were saved to, or Nothing if nothing was saved.</summary>
        Public Property Saved As String
        Public Property WouldHaveChanged As Boolean?
        Public Property Warnings As List(Of String)
        Public Property [Error] As String
        Public Property ErrorCode As Integer?

        ''' <summary>The first command with this name, or Nothing.</summary>
        Public Function Find(name As String) As DiskCommandJson
            If Commands Is Nothing Then Return Nothing
            For Each c In Commands
                If String.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase) Then Return c
            Next
            Return Nothing
        End Function

        ''' <summary>The typed result of the first command with this name, or Nothing.</summary>
        Public Function ResultOf(Of T As Class)(name As String) As T
            Dim c As DiskCommandJson = Find(name)
            If c Is Nothing Then Return Nothing
            Return c.GetResult(Of T)()
        End Function

        ''' <summary>Every warning of the run: the decoders' and each command's.</summary>
        Public Function AllWarnings() As List(Of String)
            Dim all As New List(Of String)
            If Warnings IsNot Nothing Then all.AddRange(Warnings)
            If Commands IsNot Nothing Then
                For Each c In Commands
                    If c.Warnings IsNot Nothing Then all.AddRange(c.Warnings)
                Next
            End If
            Return all
        End Function

        ''' <summary>The run's error: a command's, else the run's own; Nothing when everything worked.</summary>
        Public Function FirstError(ByRef code As Integer) As String
            code = 0
            If Commands IsNot Nothing Then
                For Each c In Commands
                    If c.Error IsNot Nothing Then
                        code = If(c.ErrorCode, 1)
                        Return c.Error
                    End If
                Next
            End If
            If [Error] IsNot Nothing Then
                code = If(ErrorCode, 1)
                Return [Error]
            End If
            Return Nothing
        End Function

        ''' <summary>The text the run would have printed: every command's lines, then the save line.</summary>
        Public Function MessageText() As String
            Dim sb As New StringBuilder()
            If Commands IsNot Nothing Then
                For Each c In Commands
                    If c.Warnings IsNot Nothing Then
                        For Each w In c.Warnings
                            sb.Append(w).Append(ControlChars.Lf)
                        Next
                    End If
                    If c.Messages IsNot Nothing Then
                        For Each m In c.Messages
                            sb.Append(m).Append(ControlChars.Lf)
                        Next
                    End If
                Next
            End If
            If Saved IsNot Nothing Then
                sb.Append("Saved ").Append(Saved).Append("."c).Append(ControlChars.Lf)
            ElseIf WouldHaveChanged.HasValue AndAlso WouldHaveChanged.Value Then
                sb.Append("Not saved (preview): the image would have changed.").Append(ControlChars.Lf)
            End If
            Return sb.ToString()
        End Function
    End Class

    ''' <summary>
    ''' The free space on a disk ("free", and inside "info"). trs80-tool 5.2.0 counts a TRSDOS-family disk in granules
    ''' and a CP/M disk in blocks; Unit says which ("granule" or "block"; Nothing from 5.1.0, which only had granules).
    ''' </summary>
    Public Class FreeResultJson
        Public Property Unit As String
        ''' <summary>Granules new files could be given.</summary>
        Public Property FreeGranules As Integer
        Public Property GranuleBytes As Integer
        Public Property FreeBytes As Long
        ''' <summary>Granules the GAT calls free but the image doesn't fully hold; new files skip them.</summary>
        Public Property MissingGranules As Integer
        ''' <summary>Granules the GAT calls free past the disk's proven size.</summary>
        Public Property PastEndGranules As Integer
        ''' <summary>CP/M: blocks new files could be given (no directory entry claims them).</summary>
        Public Property FreeBlocks As Integer
        Public Property BlockBytes As Integer
        ''' <summary>CP/M: unclaimed blocks with a sector the image lacks or a CRC error; new files skip them.</summary>
        Public Property MissingBlocks As Integer

        Public ReadOnly Property IsBlocks As Boolean
            Get
                Return String.Equals(Unit, "block", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property
    End Class

    Public Class PartitionLocationJson
        Public Property CylinderOffset As Integer
        Public Property HeadOffset As Integer
        Public Property CylinderCount As Integer
        Public Property HeadCount As Integer
    End Class

    ''' <summary>One partition of a hard disk, as "disk info" and "disk partition" give it.</summary>
    Public Class DiskPartitionJson
        Public Property Number As Integer
        Public Property DiskName As String
        Public Property Location As PartitionLocationJson
        ''' <summary>The partition's DOS name and version.</summary>
        Public Property Os As String
        Public Property FileCount As Integer
        Public Property Writable As Boolean
        Public Property ReadOnlyReason As String
        Public Property Selected As Boolean
    End Class

    Public Class HardDiskGeometryJson
        Public Property CylinderCount As Integer
        Public Property HeadCount As Integer
        Public Property SectorsPerTrack As Integer
        Public Property SectorSize As Integer
    End Class

    ''' <summary>The result of "disk info".</summary>
    Public Class InfoResultJson
        Public Property Description As String
        Public Property Os As OsInfoJson
        Public Property Geometry As GeometryInfoJson
        Public Property HardDiskGeometry As HardDiskGeometryJson
        ''' <summary>Which DOS's rules a change follows ("ldos-5"); Nothing when the disk can't be written by rules.</summary>
        Public Property Rules As String
        Public Property Writable As Boolean
        Public Property ReadOnlyReason As String
        ''' <summary>The image's write-protect flag; Nothing when the format has none (JV1, IMD).</summary>
        Public Property WriteProtected As Boolean?
        Public Property Free As FreeResultJson
        Public Property Partitions As List(Of DiskPartitionJson)
    End Class

    Public Class GranuleJson
        Public Property Lump As Integer
        Public Property Granule As Integer
    End Class

    ''' <summary>One thing "disk check" found. Severity: "layout", "damage" or "note".</summary>
    Public Class FindingJson
        Public Property Kind As String
        Public Property Severity As String
        Public Property Message As String
        Public Property Files As List(Of String)
        Public Property Granules As List(Of GranuleJson)
        ''' <summary>Only in "repair": why it wasn't fixed.</summary>
        Public Property Advice As String
    End Class

    Public Class TrackRefJson
        Public Property Cylinder As Integer
        Public Property Side As Integer
    End Class

    ''' <summary>One sign of copy protection. Verdict: "protected", "dosProtected", "possible" or "info".</summary>
    Public Class ProtectionSignJson
        Public Property Kind As String
        Public Property Verdict As String
        ''' <summary>Scheme numbers of the trs-80.com archive this sign has been seen with ("S-101").</summary>
        Public Property Schemes As List(Of String)
        ''' <summary>The scheme the sign identifies by itself, when it does.</summary>
        Public Property Scheme As String
        Public Property Message As String
        Public Property Tracks As List(Of TrackRefJson)
        Public Property Sectors As List(Of SectorPositionJson)
    End Class

    ''' <summary>One kind of imaging artifact: "readBeganLate", "readOneRevolution" or "converterCrcMarker".</summary>
    Public Class ImagingArtifactJson
        Public Property Kind As String
        Public Property Message As String
        Public Property Tracks As List(Of TrackRefJson)
        Public Property Sectors As List(Of SectorPositionJson)
    End Class

    ''' <summary>The result of "disk check".</summary>
    Public Class CheckResultJson
        Public Property Findings As List(Of FindingJson)
        Public Property Writable As Boolean
        Public Property ReadOnlyReason As String
        ''' <summary>Floppies only: signs of copy protection (empty = none found, which proves nothing).</summary>
        Public Property Protection As List(Of ProtectionSignJson)
        Public Property ImagingArtifacts As List(Of ImagingArtifactJson)
    End Class

    ''' <summary>The result of "disk repair".</summary>
    Public Class RepairResultJson
        Public Property Changes As List(Of String)
        Public Property Remaining As List(Of FindingJson)
        Public Property FreeableUnclaimedGranules As Integer
        Public Property MarkableMissingGranules As Integer
        Public Property FixablePrograms As List(Of String)
        Public Property Writable As Boolean
        Public Property ReadOnlyReason As String
    End Class

    Public Class ListedFileJson
        Public Property Name As String
        ''' <summary>"basic", "text" or "hex".</summary>
        Public Property Kind As String
        ''' <summary>The file's bytes as stored on the disk, base64.</summary>
        Public Property Contents As String
    End Class

    ''' <summary>The result of "disk list".</summary>
    Public Class ListResultJson
        Public Property Files As List(Of ListedFileJson)
        Public Property Listing As List(Of String)
    End Class

    ''' <summary>One file "disk add" put on the disk.</summary>
    Public Class AddedFileJson
        ''' <summary>The host path it came from, as given.</summary>
        Public Property Host As String
        ''' <summary>Its name on the disk, written "HELLO.BAS".</summary>
        Public Property Name As String
        ''' <summary>Its size on the disk in bytes, after any conversion.</summary>
        Public Property Size As Long
        Public Property Replaced As Boolean
        ''' <summary>What the conversion did, as the tool words it. For people.</summary>
        Public Property Conversions As List(Of String)
    End Class

    ''' <summary>A host file that "disk add --keep-going" left out.</summary>
    Public Class SkippedHostFileJson
        Public Property Host As String
        ''' <summary>Why it wasn't added, as the tool words it. For people.</summary>
        Public Property Problem As String
    End Class

    ''' <summary>
    ''' The result of "disk add". "Skipped" is always empty without "--keep-going", and the field itself
    ''' only exists in trs80-tool 5.3.0 and later, so it comes back Nothing from an older tool.
    ''' </summary>
    Public Class AddResultJson
        Public Property Files As List(Of AddedFileJson)
        Public Property Skipped As List(Of SkippedHostFileJson)
    End Class

    ''' <summary>One file "disk import" copied onto the disk (trs80-tool 5.3.0).</summary>
    Public Class ImportedFileJson
        ''' <summary>Its name in the image it came from: "HELLO.BAS", or "HARDA/HELLO.BAS" from a partition.</summary>
        Public Property [From] As String
        ''' <summary>Its name on the disk, written "HELLO.BAS".</summary>
        Public Property Name As String
        Public Property Size As Long
        Public Property Replaced As Boolean
        ''' <summary>What the conversion did ("Converted HELLO.BAS from Model III to Model 4 BASIC tokens.").</summary>
        Public Property Conversions As List(Of String)
    End Class

    ''' <summary>A file "disk import --keep-going" left out.</summary>
    Public Class SkippedFileJson
        Public Property [From] As String
        Public Property Problem As String
    End Class

    ''' <summary>The result of "disk import". "Skipped" is always empty without "--keep-going".</summary>
    Public Class ImportResultJson
        ''' <summary>The image the files were copied from, as given.</summary>
        Public Property Source As String
        Public Property Files As List(Of ImportedFileJson)
        Public Property Skipped As List(Of SkippedFileJson)
    End Class

    ''' <summary>One granule of "disk map".</summary>
    Public Class MapCellJson
        Public Property Lump As Integer
        Public Property Granule As Integer
        Public Property Sectors As List(Of SectorPositionJson)
        ''' <summary>
        ''' "free", "notOnImage", "pastEnd", "file", "otherFile", "directory", "boot", "lockedOut", "shared",
        ''' "claimedButFree", "unclaimed" (a later tool may add kinds).
        ''' </summary>
        Public Property Kind As String
        ''' <summary>Index into the map's files; only when Kind is "file".</summary>
        Public Property File As Integer?
        ''' <summary>Every file whose directory entries claim the granule.</summary>
        Public Property Owners As List(Of String)
    End Class

    Public Class MapFileJson
        Public Property Name As String
        Public Property GranuleCount As Integer
        Public Property ColorIndex As Integer?
    End Class

    Public Class MapTotalsJson
        Public Property GranuleCount As Integer
        Public Property FreeGranuleCount As Integer
        Public Property NotOnImageGranuleCount As Integer
        Public Property PastEndGranuleCount As Integer
        Public Property LockedOutGranuleCount As Integer
        Public Property FreeByteCount As Long
        Public Property GranuleByteCount As Integer
    End Class

    Public Class MapLayoutJson
        Public Property SectorsPerTrack As Integer
        Public Property FirstSectorNumber As Integer
        Public Property AllocationSides As Integer
        Public Property ImageSideCount As Integer
        Public Property LastCylinder As Integer
        ' From trs80-tool 5.2.0: "sectors" (a floppy) or "granules" (a hard disk partition), the lumps, and the DOS's
        ' word for a lump ("cylinder" where lumps are cylinders, otherwise "lump"). Nothing / 0 from 5.1.0.
        Public Property Grid As String
        Public Property GranulesPerLump As Integer
        Public Property LumpCount As Integer
        Public Property LumpTerm As String
    End Class

    ''' <summary>The result of "disk map": where everything is on a TRSDOS-family floppy.</summary>
    Public Class MapResultJson
        Public Property Cells As List(Of MapCellJson)
        Public Property Files As List(Of MapFileJson)
        Public Property KindCounts As Dictionary(Of String, Integer)
        Public Property Totals As MapTotalsJson
        Public Property Layout As MapLayoutJson
    End Class

    ' ------------------------------------------------------------------ read-only queries

    Private Shared Function ParseJson(Of T As Class)(r As ToolResult) As T
        Dim text As String = r.StdOut
        If text Is Nothing Then Return Nothing
        Dim start As Integer = text.IndexOf("{"c)
        If start < 0 Then Return Nothing
        Try
            Return JsonSerializer.Deserialize(Of T)(text.Substring(start), JsonOpts)
        Catch ex As JsonException
            Return Nothing
        End Try
    End Function

    ''' <summary>dir --json [--system] IMAGE → the one infile entry.</summary>
    Public Async Function DirJsonAsync(imagePath As String, includeSystem As Boolean) As Task(Of FilesInfileJson)
        Dim full As String = Path.GetFullPath(imagePath)
        Dim cli As New List(Of String) From {"dir", "--json"}
        If includeSystem Then cli.Add("--system")
        Dim sargs As New Dictionary(Of String, Object) From {{"images", New String() {full}}, {"system", includeSystem}}
        Dim cpmFormat As String = CpmFormatOf(full)
        If cpmFormat.Length > 0 Then
            cli.Add("--cpm-format") : cli.Add(cpmFormat)
            sargs("cpmFormat") = cpmFormat
        End If
        cli.Add(full)
        Dim r As ToolResult = Await JsonCallAsync("dir", sargs, cli, False)
        Dim parsed As InfilesJson(Of FilesInfileJson) = ParseJson(Of InfilesJson(Of FilesInfileJson))(r)
        If parsed Is Nothing OrElse parsed.Infiles Is Nothing OrElse parsed.Infiles.Count = 0 Then Throw New ToolException(r)
        Dim entry As FilesInfileJson = parsed.Infiles(0)
        If entry.Files Is Nothing Then entry.Files = New List(Of FileJson)
        Return entry
    End Function

    ''' <summary>info --json --verbose IMAGE → the one infile entry.</summary>
    Public Async Function InfoJsonAsync(imagePath As String) As Task(Of InfoJson)
        Dim full As String = Path.GetFullPath(imagePath)
        Dim cli As New List(Of String) From {"info", "--json", "--verbose"}
        Dim sargs As New Dictionary(Of String, Object) From {{"images", New String() {full}}, {"verbose", True}}
        Dim cpmFormat As String = CpmFormatOf(full)
        If cpmFormat.Length > 0 Then
            cli.Add("--cpm-format") : cli.Add(cpmFormat)
            sargs("cpmFormat") = cpmFormat
        End If
        cli.Add(full)
        Dim r As ToolResult = Await JsonCallAsync("info", sargs, cli, False)
        Dim parsed As InfilesJson(Of InfoJson) = ParseJson(Of InfilesJson(Of InfoJson))(r)
        If parsed Is Nothing OrElse parsed.Infiles Is Nothing OrElse parsed.Infiles.Count = 0 Then Throw New ToolException(r)
        Return parsed.Infiles(0)
    End Function

    ' ------------------------------------------------------------------ disk runs ("disk IMAGE --json ...")

    ''' <summary>Quote one word for the disk shell's tokenizer: double quotes; \ and " escaped with a backslash.</summary>
    Public Shared Function QuoteScriptWord(word As String) As String
        If word Is Nothing Then word = ""
        If word.Length > 0 AndAlso Regex.IsMatch(word, "^[A-Za-z0-9_./\-]+$") Then Return word
        Return """" & word.Replace("\", "\\").Replace("""", "\""") & """"
    End Function

    ''' <summary>Longest command line sent as arguments; a longer run goes as a script on standard input.</summary>
    Private Const MaxCommandLineChars As Integer = 12000

    ''' <summary>
    ''' Run disk commands in one go and return the run as JSON. Never throws for a failure the tool reports:
    ''' check Succeeded (ErrorText has the tool's sentence, Failure why). StdOut holds the text the commands would
    ''' have printed; Disk the parsed document.
    ''' </summary>
    ''' <param name="imagePath">The image.</param>
    ''' <param name="partition">Partition number or name for a hard disk, else Nothing.</param>
    ''' <param name="commands">Each element is one command's words.</param>
    ''' <param name="changes">True when a command may change the image or write host files (extract).</param>
    ''' <param name="dryRun">Write nothing; say whether the image would have changed.</param>
    ''' <param name="force">The run's --force.</param>
    Public Async Function DiskRunAsync(imagePath As String, partition As String, commands As IEnumerable(Of IEnumerable(Of String)), changes As Boolean, Optional dryRun As Boolean = False, Optional force As Boolean = False) As Task(Of ToolResult)
        Dim full As String = Path.GetFullPath(imagePath)
        Dim globals As New List(Of String)
        If Not String.IsNullOrEmpty(partition) Then
            globals.Add("--partition")
            globals.Add(partition)
        End If
        Dim cpmFormat As String = CpmFormatOf(full)
        If cpmFormat.Length > 0 Then
            globals.Add("--cpm-format")
            globals.Add(cpmFormat)
        End If
        If dryRun Then globals.Add("--dry-run")
        If force Then globals.Add("--force")

        Dim words As New List(Of String)(globals)
        Dim script As New StringBuilder()
        Dim length As Integer = 0
        Dim commandCount As Integer = 0
        Dim first As Boolean = True
        For Each cmd In commands
            If Not first Then words.Add(":")
            first = False
            commandCount += 1
            Dim quoted As New List(Of String)
            For Each w In cmd
                words.Add(w)
                quoted.Add(QuoteScriptWord(w))
                length += w.Length + 3
            Next
            script.Append(String.Join(" ", quoted)).Append(ControlChars.Lf)
        Next

        Dim cli As New List(Of String) From {"disk", full, "--json"}
        Dim stdinText As String = Nothing
        If length > MaxCommandLineChars Then
            ' Too long for a command line: the same commands as a script (one per line; saved at the end).
            cli.AddRange(globals)
            stdinText = script.ToString()
        Else
            cli.AddRange(words)
        End If
        Dim sargs As New Dictionary(Of String, Object) From {{"image", full}, {"words", words.ToArray()}}
        Dim serverTool As String = If(changes AndAlso Not dryRun, "disk_write", "disk")
        ' Every changing command of a run costs the tool a second or two on a large disk: allow for a long run.
        Dim allowMs As Integer = TimeoutMs + If(changes, Math.Min(commandCount, 2000) * 5000, 0)
        Dim r As ToolResult = Await JsonCallAsync(serverTool, sargs, cli, changes AndAlso Not dryRun, stdinText, allowMs)
        If stdinText IsNot Nothing Then r.CommandLine &= "  < script of " & script.ToString().Split(ControlChars.Lf).Length.ToString(CultureInfo.InvariantCulture) & " lines"
        If r.TimedOut OrElse r.ErrorMessage IsNot Nothing Then Return r

        Dim run As DiskRunJson = ParseJson(Of DiskRunJson)(r)
        If run Is Nothing Then
            ' Not a document: a crash or a usage error. Keep whatever text there is as the error.
            If r.ExitCode = 0 Then r.ExitCode = 1
            Return r
        End If
        r.Disk = run
        Dim json As String = r.StdOut
        Dim code As Integer
        Dim err As String = run.FirstError(code)
        r.StdOut = run.MessageText()
        If err IsNot Nothing Then
            r.ErrorMessage = err
            r.ExitCode = If(code = 0, 1, code)
        Else
            r.ExitCode = 0
        End If
        Return r
    End Function

    Private Shared Function One(ParamArray words As String()) As List(Of IEnumerable(Of String))
        Return New List(Of IEnumerable(Of String)) From {words}
    End Function

    ''' <summary>Everything the main window needs to know about an open disk, from one "info : check" run ("info : free" on CP/M).</summary>
    Public Class DiskStatus
        Public Property Info As InfoResultJson
        ''' <summary>The "free" result on a CP/M disk, whose "info" has no free space; Nothing otherwise.</summary>
        Public Property Free As FreeResultJson
        ''' <summary>The lines "info" prints (description first, then indented details).</summary>
        Public Property InfoLines As New List(Of String)
        ''' <summary>Nothing when "check" doesn't apply to the disk (no TRSDOS-family file system).</summary>
        Public Property Check As CheckResultJson
        Public Property CheckLines As New List(Of String)
        Public Property Warnings As New List(Of String)
        ''' <summary>Why "info" itself failed, else Nothing.</summary>
        Public Property [Error] As String
    End Class

    ''' <summary>
    ''' disk IMAGE [--partition P] --json info : check → free space, write-protect flag, whether and why the disk
    ''' is read-only, the partitions of a hard disk, the check findings, signs of copy protection and imaging
    ''' artifacts. "check" is last, so a disk it doesn't apply to still gives the "info" result. Never throws.
    ''' On a CP/M disk (cpm = True) the run is "info : free" instead: trs80-tool 5.2.0 writes CP/M disks and counts
    ''' their free blocks with "free", but its "info" has no free space for them and "check" doesn't cover them.
    ''' </summary>
    Public Async Function DiskStatusAsync(imagePath As String, partition As String, Optional cpm As Boolean = False) As Task(Of DiskStatus)
        Dim st As New DiskStatus
        Dim r As ToolResult
        Try
            r = Await DiskRunAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {New String() {"info"}, New String() {If(cpm, "free", "check")}}, False)
        Catch ex As Exception
            st.Error = ex.Message
            Return st
        End Try
        If r.Disk Is Nothing Then
            st.Error = r.ErrorText
            Return st
        End If
        st.Warnings = r.Disk.AllWarnings()
        Dim infoCmd As DiskCommandJson = r.Disk.Find("info")
        If infoCmd Is Nothing OrElse infoCmd.Error IsNot Nothing Then
            st.Error = If(infoCmd IsNot Nothing, infoCmd.Error, r.ErrorText)
            Return st
        End If
        st.Info = infoCmd.GetResult(Of InfoResultJson)()
        If infoCmd.Messages IsNot Nothing Then st.InfoLines.AddRange(infoCmd.Messages)
        Dim checkCmd As DiskCommandJson = r.Disk.Find("check")
        If checkCmd IsNot Nothing AndAlso checkCmd.Error Is Nothing Then
            st.Check = checkCmd.GetResult(Of CheckResultJson)()
            If checkCmd.Messages IsNot Nothing Then st.CheckLines.AddRange(checkCmd.Messages)
        End If
        Dim freeCmd As DiskCommandJson = r.Disk.Find("free")
        If freeCmd IsNot Nothing AndAlso freeCmd.Error Is Nothing Then st.Free = freeCmd.GetResult(Of FreeResultJson)()
        Return st
    End Function

    ''' <summary>
    ''' The bytes of one file exactly as stored: "disk list NAME --raw" as a launch, or the base64 "contents" of
    ''' "disk --json list NAME" through the server.
    ''' </summary>
    Public Async Function ReadFileBytesAsync(imagePath As String, partition As String, trsName As String) As Task(Of Byte())
        If ServerWanted() Then
            Dim jr As ToolResult = Await DiskRunAsync(imagePath, partition, One("list", trsName), False)
            If jr.Disk IsNot Nothing Then
                If Not jr.Succeeded Then Throw New ToolException(jr)
                Dim lr As ListResultJson = jr.Disk.ResultOf(Of ListResultJson)("list")
                If lr IsNot Nothing AndAlso lr.Files IsNot Nothing AndAlso lr.Files.Count = 1 AndAlso lr.Files(0).Contents IsNot Nothing Then
                    Try
                        Return Convert.FromBase64String(lr.Files(0).Contents)
                    Catch ex As FormatException
                        ' Fall through to the launch.
                    End Try
                End If
            End If
        End If
        ThrowIfNoTool()
        Dim r As ToolResult = Await RunAsync(DiskWords(imagePath, partition, "list", trsName, "--raw"), Nothing, True)
        If Not r.Succeeded Then
            ' The error text went to stdout as bytes; surface it.
            r.StdOut = If(r.StdOutBytes Is Nothing, "", Encoding.UTF8.GetString(r.StdOutBytes))
            Throw New ToolException(r)
        End If
        Return If(r.StdOutBytes, Array.Empty(Of Byte)())
    End Function

    ''' <summary>disk IMAGE list NAME [--hex] [--basic 3|4] → the listing text (LF line ends).</summary>
    Public Async Function ListTextAsync(imagePath As String, partition As String, trsName As String, hex As Boolean, basicDialect As String) As Task(Of String)
        Dim words As New List(Of String) From {"list", trsName}
        If hex Then words.Add("--hex")
        If Not String.IsNullOrEmpty(basicDialect) Then
            words.Add("--basic")
            words.Add(basicDialect)
        End If
        Dim r As ToolResult = Await DiskRunAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words}, False)
        ThrowIfFailed(r)
        Return r.StdOut
    End Function

    ''' <summary>
    ''' disasm HOSTFILE [options] → text. Option words as the tool takes them: --listing, --org ADDR, --entry A,B,
    ''' --no-labels, --no-known, --no-binary, --upper, --full, --data-comment, --hex-format c|dollar|h.
    ''' </summary>
    Public Async Function DisasmAsync(hostFile As String, optionWords As IEnumerable(Of String)) As Task(Of String)
        Dim args As New List(Of String) From {"disasm"}
        If optionWords IsNot Nothing Then args.AddRange(optionWords)
        args.Add(hostFile)
        Dim r As ToolResult = Await RunAsync(args)
        ThrowIfFailed(r)
        Return r.StdOut
    End Function

    ''' <summary>asm SRC OUT [--model 1|3|4] [--listing FILE] [--baud N] → the tool's output (not thrown; check Succeeded).</summary>
    Public Async Function AsmAsync(sourceFile As String, outFile As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim args As New List(Of String) From {"asm"}
        If optionWords IsNot Nothing Then args.AddRange(optionWords)
        args.Add(sourceFile)
        args.Add(outFile)
        Return Await RunAsync(args)
    End Function

    ''' <summary>
    ''' convert IN OUT (any file conversion: programs, floppy image formats, hi-res screens to PNG). Throws on failure.
    ''' </summary>
    Public Async Function ConvertAsync(inFile As String, outFile As String, Optional force As Boolean = False) As Task(Of ToolResult)
        Dim args As New List(Of String) From {"convert"}
        If force Then args.Add("--force")
        args.Add(inFile)
        args.Add(outFile)
        Dim r As ToolResult = Await RunAsync(args)
        ThrowIfFailed(r)
        Return r
    End Function

    ''' <summary>
    ''' convert IN... OUT with option words (--baud N, --start ADDR|auto, --system, --cpm-records, --force).
    ''' Several inputs are allowed for CAS/WAV outputs. Not thrown; check Succeeded.
    ''' </summary>
    Public Async Function ConvertFilesAsync(inFiles As IEnumerable(Of String), outFile As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim args As New List(Of String) From {"convert"}
        If optionWords IsNot Nothing Then args.AddRange(optionWords)
        args.AddRange(inFiles)
        args.Add(outFile)
        Return Await RunAsync(args)
    End Function

    ''' <summary>hexdump HOSTFILE [--no-collapse] → text.</summary>
    Public Async Function HexdumpAsync(hostFile As String, Optional noCollapse As Boolean = False) As Task(Of String)
        Dim args As New List(Of String) From {"hexdump"}
        If noCollapse Then args.Add("--no-collapse")
        args.Add(hostFile)
        Dim r As ToolResult = Await RunAsync(args)
        ThrowIfFailed(r)
        Return r.StdOut
    End Function

    ''' <summary>
    ''' disk IMAGE [--partition P] map [PATTERN...] → the granule map (TRSDOS-family floppies only): the drawing as
    ''' text in StdOut and the cells in Disk ("map" result, MapResultJson). Not thrown.
    ''' </summary>
    Public Async Function DiskMapAsync(imagePath As String, partition As String, patterns As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"map"}
        If patterns IsNot Nothing Then words.AddRange(patterns)
        Return Await DiskRunAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words}, False)
    End Function

    ''' <summary>disk IMAGE [--partition P] check → the report (not thrown; a non-TRSDOS disk gives the tool's error text).</summary>
    Public Async Function DiskCheckAsync(imagePath As String, partition As String) As Task(Of ToolResult)
        Return Await DiskRunAsync(imagePath, partition, One("check"), False)
    End Function

    ''' <summary>
    ''' disk IMAGE [--partition P] [--dry-run] repair [--free-unclaimed] [--fix-program-sizes] [--mark-missing-used]
    ''' → the tool's report. Not thrown: the report is shown to the user either way.
    ''' </summary>
    Public Async Function DiskRepairAsync(imagePath As String, partition As String, freeUnclaimed As Boolean, fixProgramSizes As Boolean, markMissingUsed As Boolean, dryRun As Boolean) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"repair"}
        If freeUnclaimed Then words.Add("--free-unclaimed")
        If fixProgramSizes Then words.Add("--fix-program-sizes")
        If markMissingUsed Then words.Add("--mark-missing-used")
        Return Await DiskRunAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words}, True, dryRun)
    End Function

    ''' <summary>
    ''' disk IMAGE [--partition P] format [DOS] [create options] [--force]. --force is required when the disk has
    ''' files (the GUI confirms first). On a hard disk it formats the chosen partition. Not thrown.
    ''' </summary>
    Public Async Function DiskFormatAsync(imagePath As String, partition As String, dosName As String, optionWords As IEnumerable(Of String), force As Boolean) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"format"}
        If Not String.IsNullOrEmpty(dosName) Then words.Add(dosName)
        If optionWords IsNot Nothing Then words.AddRange(optionWords)
        If force Then words.Add("--force")
        Return Await DiskRunAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words}, True)
    End Function

    ''' <summary>disk IMAGE trim [--cylinders N] [--side 0] [--force]. Not thrown.</summary>
    Public Async Function DiskTrimAsync(imagePath As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"trim"}
        If optionWords IsNot Nothing Then words.AddRange(optionWords)
        Return Await DiskRunAsync(imagePath, Nothing, New List(Of IEnumerable(Of String)) From {words}, True)
    End Function

    ''' <summary>
    ''' disk IMAGE partition --add DOS [options]: make a partition in a hard disk's free space. The tool's text
    ''' (StdOut) includes the settings the DOS's hard disk driver must be given. Not thrown.
    ''' </summary>
    Public Async Function DiskPartitionAddAsync(imagePath As String, dosName As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"partition", "--add", dosName}
        If optionWords IsNot Nothing Then words.AddRange(optionWords)
        Return Await DiskRunAsync(imagePath, Nothing, New List(Of IEnumerable(Of String)) From {words}, True)
    End Function

    ''' <summary>disk IMAGE partition --delete N [--force]: erase a partition of a hard disk. Not thrown.</summary>
    Public Async Function DiskPartitionDeleteAsync(imagePath As String, partitionNumber As Integer, force As Boolean) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"partition", "--delete", partitionNumber.ToString(CultureInfo.InvariantCulture)}
        If force Then words.Add("--force")
        Return Await DiskRunAsync(imagePath, Nothing, New List(Of IEnumerable(Of String)) From {words}, True)
    End Function

    ''' <summary>
    ''' Start trs80-tool in its own console window (for the interactive disk shell and the terminal emulator),
    ''' without redirecting anything; the caller may wait for Exited. Program options are added as usual.
    ''' Throws ToolException when the tool isn't found.
    ''' </summary>
    Public Function StartInConsole(args As IEnumerable(Of String), Optional keepColor As Boolean = True) As ConsoleRun
        ThrowIfNoTool()
        Dim psi As New ProcessStartInfo(ToolPath) With {
            .UseShellExecute = True,
            .CreateNoWindow = False,
            .WindowStyle = ProcessWindowStyle.Normal,
            .WorkingDirectory = Path.GetDirectoryName(ToolPath)
        }
        AddProgramOptions(psi, keepColor)
        For Each a In args
            psi.ArgumentList.Add(a)
        Next
        LogCall("[console] " & DescribeCommandLine(args))
        ' Windows gives the program a console window of its own. Linux and macOS don't, so there it runs in a
        ' terminal program (ConsoleRun).
        If HostPlatform.IsWindows Then Return ConsoleRun.StartWithOwnWindow(psi)
        Return ConsoleRun.StartInTerminal(ToolPath, psi.ArgumentList.ToList(), TerminalCommand)
    End Function

    ' ------------------------------------------------------------------ raw sectors (sectors --json)

    ''' <summary>One CRC of a sector (ID field or data) as the image format recorded it, from sectors.ts.</summary>
    Public Class SectorCrcJson
        Public Property Stored As Integer
        Public Property Computed As Integer
        Public ReadOnly Property Valid As Boolean
            Get
                Return Stored = Computed
            End Get
        End Property
    End Class

    ''' <summary>One sector of "sectors --json" (schema in sectors.ts, v5.1.0). Fields after Files exist for present sectors only.</summary>
    Public Class SectorJson
        Public Property Cylinder As Integer
        Public Property Side As Integer
        Public Property Sector As Integer
        ''' <summary>No readable sector at this position.</summary>
        Public Property Missing As Boolean
        ''' <summary>Sector number outside the usual numbers of its track (copy protection, perhaps).</summary>
        Public Property OutsideUsualNumbers As Boolean
        ''' <summary>Files that use this sector, in directory order (more than one = cross-linked).</summary>
        Public Property Files As List(Of String)
        Public Property Density As String
        Public Property Size As Integer?
        Public Property Deleted As Boolean?
        Public Property DataAddressMark As Integer?
        Public Property CrcError As Boolean?
        Public Property IdCrc As SectorCrcJson
        Public Property DataCrc As SectorCrcJson
        ''' <summary>
        ''' Why the sector is missing or bad, when the imaging explains it: "readBeganLate", "readOneRevolution"
        ''' or "converterCrcMarker". Nothing otherwise.
        ''' </summary>
        Public Property ImagingArtifact As String
        ''' <summary>Base64 data; only with --contents.</summary>
        Public Property Contents As String

        Public ReadOnly Property IsBad As Boolean
            Get
                Return Missing OrElse (CrcError.HasValue AndAlso CrcError.Value)
            End Get
        End Property

        ''' <summary>The map letter the tool prints, with its precedence: ? missing, C CRC error, X deleted mark, S/D density.</summary>
        Public ReadOnly Property MapLetter As String
            Get
                If Missing Then Return "?"
                If CrcError.HasValue AndAlso CrcError.Value Then Return "C"
                If Deleted.HasValue AndAlso Deleted.Value Then Return "X"
                If String.Equals(Density, "single", StringComparison.OrdinalIgnoreCase) Then Return "S"
                Return "D"
            End Get
        End Property

        ''' <summary>The imaging artifact in words, or "" when there is none.</summary>
        Public ReadOnly Property ImagingArtifactText As String
            Get
                Select Case If(ImagingArtifact, "")
                    Case "" : Return ""
                    Case "readBeganLate" : Return "the read of the track began late, after this sector had gone by"
                    Case "readOneRevolution" : Return "the read stopped after one revolution, cutting this sector short"
                    Case "converterCrcMarker" : Return "the stored CRC is a converter's marker for a sector it couldn't read"
                End Select
                Return ImagingArtifact
            End Get
        End Property

        ''' <summary>The header line the tool prints for the sector with --contents / --bad.</summary>
        Public ReadOnly Property HeaderText As String
            Get
                Dim h As String = "Cylinder " & Cylinder.ToString(CultureInfo.InvariantCulture) & ", side " & Side.ToString(CultureInfo.InvariantCulture) & ", sector " & Sector.ToString(CultureInfo.InvariantCulture) & ": "
                If Missing Then Return h & "missing"
                h &= If(String.Equals(Density, "single", StringComparison.OrdinalIgnoreCase), "single", "double") & " density"
                If Deleted.HasValue AndAlso Deleted.Value Then h &= ", marked as deleted"
                If OutsideUsualNumbers Then h &= ", outside the usual sector numbers"
                If CrcError.HasValue AndAlso CrcError.Value Then
                    h &= ", CRC error"
                    Dim parts As New List(Of String)
                    If IdCrc IsNot Nothing AndAlso Not IdCrc.Valid Then parts.Add("ID " & IdCrc.Stored.ToString("X4", CultureInfo.InvariantCulture) & " != " & IdCrc.Computed.ToString("X4", CultureInfo.InvariantCulture))
                    If DataCrc IsNot Nothing AndAlso Not DataCrc.Valid Then parts.Add("data " & DataCrc.Stored.ToString("X4", CultureInfo.InvariantCulture) & " != " & DataCrc.Computed.ToString("X4", CultureInfo.InvariantCulture))
                    If parts.Count > 0 Then h &= " (" & String.Join(", ", parts) & ")"
                End If
                If Size.HasValue Then h &= ", " & Size.Value.ToString(CultureInfo.InvariantCulture) & " bytes"
                If DataAddressMark.HasValue Then h &= ", DAM 0x" & DataAddressMark.Value.ToString("X2", CultureInfo.InvariantCulture)
                Return h
            End Get
        End Property

        ''' <summary>The sector bytes, or Nothing when absent or not requested.</summary>
        Public Function GetBytes() As Byte()
            If Missing OrElse Contents Is Nothing Then Return Nothing
            Try
                Return Convert.FromBase64String(Contents)
            Catch
                Return Nothing
            End Try
        End Function
    End Class

    Public Class SectorsInfileJson
        Public Property Infile As String
        Public Property [Error] As String
        Public Property ErrorCode As Integer?
        Public Property Warnings As List(Of String)
        Public Property Description As String
        Public Property Os As String
        Public Property Sectors As List(Of SectorJson)
        ''' <summary>Signs of copy protection; only when asked for (--protection).</summary>
        Public Property Protection As List(Of ProtectionSignJson)
        ''' <summary>What the imaging did to the disk that looks like damage.</summary>
        Public Property ImagingArtifacts As List(Of ImagingArtifactJson)
    End Class

    ''' <summary>
    ''' sectors --json [--contents] [--bad] [--protection] IMAGE → the one infile entry (floppies only). With
    ''' badOnly, a clean floppy is omitted from the tool's JSON, so the result then has an empty Sectors list and
    ''' no error. Throws ToolException when the tool reports an error for the file (not a floppy, unreadable).
    ''' </summary>
    Public Async Function SectorsInfileAsync(imagePath As String, includeContents As Boolean, badOnly As Boolean, Optional protection As Boolean = False) As Task(Of SectorsInfileJson)
        Dim full As String = Path.GetFullPath(imagePath)
        Dim cli As New List(Of String) From {"sectors", "--json"}
        If includeContents Then cli.Add("--contents")
        If badOnly Then cli.Add("--bad")
        If protection Then cli.Add("--protection")
        cli.Add(full)
        Dim sargs As New Dictionary(Of String, Object) From {{"images", New String() {full}}, {"contents", includeContents}, {"bad", badOnly}, {"protection", protection}}
        Dim r As ToolResult = Await JsonCallAsync("sectors", sargs, cli, False)
        Dim parsed As InfilesJson(Of SectorsInfileJson) = ParseJson(Of InfilesJson(Of SectorsInfileJson))(r)
        If parsed Is Nothing OrElse parsed.Infiles Is Nothing Then Throw New ToolException(r)
        If parsed.Infiles.Count = 0 Then Return New SectorsInfileJson With {.Infile = imagePath, .Sectors = New List(Of SectorJson)}
        Dim entry As SectorsInfileJson = parsed.Infiles(0)
        If entry.Error IsNot Nothing AndAlso (entry.Sectors Is Nothing OrElse entry.Sectors.Count = 0) Then Throw New ToolException(entry.Error)
        If entry.Sectors Is Nothing Then entry.Sectors = New List(Of SectorJson)
        Return entry
    End Function

    ''' <summary>
    ''' sectors [--bad | --protection] IMAGE → the tool's own text report (map and legend; with --bad only the bad
    ''' sectors and their files; with --protection the signs of copy protection). Not thrown.
    ''' </summary>
    Public Async Function SectorsTextAsync(imagePath As String, badOnly As Boolean, Optional protection As Boolean = False) As Task(Of ToolResult)
        Dim args As New List(Of String) From {"sectors"}
        If protection Then
            args.Add("--protection")
        ElseIf badOnly Then
            args.Add("--bad")
        End If
        args.Add(imagePath)
        Return Await RunAsync(args)
    End Function

    ' ------------------------------------------------------------------ changes

    ''' <summary>
    ''' "disk IMAGE [--partition P] [--cpm-format NAME] cmd args...", with the CP/M layout the user chose for the
    ''' image: for launches without --json, and for the disk shell.
    ''' </summary>
    Public Function DiskWords(imagePath As String, partition As String, ParamArray words As String()) As List(Of String)
        Dim args As New List(Of String) From {"disk", imagePath}
        If Not String.IsNullOrEmpty(partition) Then
            args.Add("--partition")
            args.Add(partition)
        End If
        Dim cpmFormat As String = CpmFormatOf(imagePath)
        If cpmFormat.Length > 0 Then
            args.Add("--cpm-format")
            args.Add(cpmFormat)
        End If
        args.AddRange(words)
        Return args
    End Function

    ''' <summary>Build "disk IMAGE [--partition P] cmd args..." argument lists (for launches without --json).</summary>
    Public Shared Function DiskArgs(imagePath As String, partition As String, ParamArray words As String()) As List(Of String)
        Dim args As New List(Of String) From {"disk", imagePath}
        If Not String.IsNullOrEmpty(partition) Then
            args.Add("--partition")
            args.Add(partition)
        End If
        args.AddRange(words)
        Return args
    End Function

    ''' <summary>
    ''' Run disk commands that change the image (or write host files) in one all-or-nothing run. Throws
    ''' ToolException when the tool reports a failure; its Failure says why.
    ''' </summary>
    Public Async Function DiskChainAsync(imagePath As String, partition As String, commands As IEnumerable(Of IEnumerable(Of String)), Optional dryRun As Boolean = False, Optional force As Boolean = False) As Task(Of ToolResult)
        Dim r As ToolResult = Await DiskRunAsync(imagePath, partition, commands, True, dryRun, force)
        ThrowIfFailed(r)
        Return r
    End Function

    Public Async Function DiskKillAsync(imagePath As String, partition As String, trsNames As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"kill"}
        words.AddRange(trsNames)
        Return Await DiskChainAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words})
    End Function

    Public Async Function DiskRenameAsync(imagePath As String, partition As String, oldName As String, newName As String) As Task(Of ToolResult)
        Return Await DiskChainAsync(imagePath, partition, One("rename", oldName, newName))
    End Function

    ''' <summary>attrib NAME with the given option words (e.g. "--invisible", "--protect", "5", "--date", "1983-01-05").</summary>
    Public Async Function DiskAttribAsync(imagePath As String, partition As String, trsName As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"attrib", trsName}
        words.AddRange(optionWords)
        Return Await DiskChainAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words})
    End Function

    ''' <summary>label with option words (--name X --date YYYY-MM-DD --auto CMD / --no-auto --password PW).</summary>
    Public Async Function DiskLabelSetAsync(imagePath As String, partition As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"label"}
        words.AddRange(optionWords)
        Return Await DiskChainAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words})
    End Function

    Public Async Function DiskProtectSetAsync(imagePath As String, writeProtected As Boolean) As Task(Of ToolResult)
        Return Await DiskChainAsync(imagePath, Nothing, One("protect", If(writeProtected, "on", "off")))
    End Function

    ''' <summary>
    ''' create [DOS] [--cylinders N --sides N --density single|double --model N --name X --date YYYY-MM-DD --system
    ''' --minimal --heads N --partitions N]. The image file must not exist yet and must end in .dmk (a floppy) or
    ''' .hdv (a hard disk; DOS may then be Nothing for a blank one). StdOut has the tool's text, which for a hard
    ''' disk includes the settings its driver must be given.
    ''' </summary>
    Public Async Function DiskCreateAsync(imagePath As String, dosName As String, optionWords As IEnumerable(Of String)) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"create"}
        If Not String.IsNullOrEmpty(dosName) Then words.Add(dosName)
        words.AddRange(optionWords)
        Return Await DiskChainAsync(imagePath, Nothing, New List(Of IEnumerable(Of String)) From {words})
    End Function

    ''' <summary>disk IMAGE extract NAME... --out-dir DIR [--lowercase] [--mark-damaged] → "Wrote ..." lines.</summary>
    Public Async Function DiskExtractAsync(imagePath As String, partition As String, trsNames As IEnumerable(Of String), outDir As String, lowercase As Boolean, Optional markDamaged As Boolean = False) As Task(Of ToolResult)
        Dim words As New List(Of String) From {"extract"}
        words.AddRange(trsNames)
        words.Add("--out-dir")
        words.Add(outDir)
        If lowercase Then words.Add("--lowercase")
        If markDamaged Then words.Add("--mark-damaged")
        Return Await DiskChainAsync(imagePath, partition, New List(Of IEnumerable(Of String)) From {words})
    End Function

    ''' <summary>
    ''' Top-level extract IMAGE [FULLNAME...] --out-dir DIR --system [--lowercase] [--mark-damaged] [--cpm-records] [--ucsd-pascal-raw].
    ''' Used instead of "disk extract" when the CP/M-records or UCSD-raw options are wanted (those exist only here).
    ''' Names are the tool's full names (PARTDIR/NAME.EXT on a hard disk). Throws on failure.
    ''' </summary>
    Public Async Function ExtractTopLevelAsync(imagePath As String, fullNames As IEnumerable(Of String), outDir As String, lowercase As Boolean, markDamaged As Boolean, cpmRecords As Boolean, ucsdRaw As Boolean) As Task(Of ToolResult)
        Dim args As New List(Of String) From {"extract", imagePath}
        args.AddRange(fullNames)
        args.Add("--out-dir")
        args.Add(outDir)
        args.Add("--system")
        Dim cpmFormat As String = CpmFormatOf(imagePath)
        If cpmFormat.Length > 0 Then
            args.Add("--cpm-format") : args.Add(cpmFormat)
        End If
        If lowercase Then args.Add("--lowercase")
        If markDamaged Then args.Add("--mark-damaged")
        If cpmRecords Then args.Add("--cpm-records")
        If ucsdRaw Then args.Add("--ucsd-pascal-raw")
        Dim r As ToolResult = Await RunAsync(args)
        ThrowIfFailed(r)
        Return r
    End Function

    ''' <summary>disk IMAGE extract NAME --as HOSTPATH (exactly one file).</summary>
    Public Async Function DiskExtractAsAsync(imagePath As String, partition As String, trsName As String, hostPath As String) As Task(Of ToolResult)
        Return Await DiskChainAsync(imagePath, partition, One("extract", trsName, "--as", hostPath))
    End Function

    ''' <summary>One host file to add, as the "add" sub-command needs it.</summary>
    Public Class AddItem
        Public Property HostPath As String
        ''' <summary>TRS-80 name to force with --as, or Nothing to let the tool choose.</summary>
        Public Property AsName As String
        Public Property Overwrite As Boolean
        ''' <summary>True = copy bytes exactly (--binary).</summary>
        Public Property Binary As Boolean
        ''' <summary>True = treat as text whatever the extension: line ends become CR, nothing else (--text). Not with Binary.</summary>
        Public Property Text As Boolean
        ''' <summary>Date to record (YYYY-MM-DD), or Nothing.</summary>
        Public Property DateYmd As String
        Public Property Invisible As Boolean
        Public Property System As Boolean
        ''' <summary>
        ''' True = "--keep-going": a file that can't be added is skipped with a warning and the rest are
        ''' added, instead of the whole command adding nothing. trs80-tool 5.3.0 and later (CanKeepGoing).
        ''' </summary>
        Public Property KeepGoing As Boolean

        Public Function ToWords() As List(Of String)
            Dim w As New List(Of String) From {"add", HostPath}
            If Not String.IsNullOrEmpty(AsName) Then
                w.Add("--as")
                w.Add(AsName)
            End If
            If Overwrite Then w.Add("--overwrite")
            If Binary Then
                w.Add("--binary")
            ElseIf Text Then
                w.Add("--text")
            End If
            If Not String.IsNullOrEmpty(DateYmd) Then
                w.Add("--date")
                w.Add(DateYmd)
            End If
            If Invisible Then w.Add("--invisible")
            If System Then w.Add("--system")
            If KeepGoing Then w.Add("--keep-going")
            Return w
        End Function
    End Class

    ''' <summary>
    ''' Add files in one all-or-nothing run. Files that take the same options (and no forced name) share one
    ''' "add" command, because every changing command of a run costs the tool a second or more on a large disk.
    ''' </summary>
    Public Async Function DiskAddAsync(imagePath As String, partition As String, items As IEnumerable(Of AddItem)) As Task(Of ToolResult)
        Dim cmds As New List(Of IEnumerable(Of String))
        Dim group As List(Of String) = Nothing      ' the words of the "add" being built: add HOST... then options
        Dim groupKey As String = Nothing
        Dim groupHosts As Integer = 0
        For Each it In items
            Dim words As List(Of String) = it.ToWords()
            If Not String.IsNullOrEmpty(it.AsName) Then
                group = Nothing
                cmds.Add(words)
                Continue For
            End If
            ' words = add HOST option...; files with the same options can share the command.
            Dim key As String = String.Join(ChrW(1), words.GetRange(2, words.Count - 2))
            If group IsNot Nothing AndAlso key = groupKey AndAlso groupHosts < 200 Then
                group.Insert(1 + groupHosts, it.HostPath)
                groupHosts += 1
            Else
                group = words
                groupKey = key
                groupHosts = 1
                cmds.Add(group)
            End If
        Next
        Return Await DiskChainAsync(imagePath, partition, cmds)
    End Function

    ''' <summary>
    ''' One run of "disk import": the image to copy files from, which of its files, and how. A file keeps its
    ''' own date and its system and hidden flags, and BASIC is converted between Model III and Model 4 when
    ''' the two disks differ, none of which the GUI's own extract-and-add path can do.
    ''' </summary>
    Public Class ImportRequest
        ''' <summary>The floppy, hard disk, CAS or WAV image to copy from.</summary>
        Public Property SourceImage As String
        ''' <summary>
        ''' Names or wildcards in the source image. Empty = every file but the system files; a wildcard never
        ''' matches a system file, so a system file has to be named.
        ''' </summary>
        Public Property Names As New List(Of String)
        ''' <summary>TRS-80 (or CP/M) name to force with --as. The tool refuses it for more than one file.</summary>
        Public Property AsName As String
        Public Property Overwrite As Boolean
        ''' <summary>True = copy every file byte for byte, BASIC included (--binary).</summary>
        Public Property Binary As Boolean
        Public Property System As Boolean
        Public Property Invisible As Boolean
        ''' <summary>Date to record instead of each file's own (YYYY-MM-DD), or Nothing to keep theirs.</summary>
        Public Property DateYmd As String
        ''' <summary>True = skip a file that can't be copied and copy the rest (--keep-going).</summary>
        Public Property KeepGoing As Boolean

        Public Function ToWords() As List(Of String)
            Dim w As New List(Of String) From {"import", SourceImage}
            If Names IsNot Nothing Then w.AddRange(Names)
            If Not String.IsNullOrEmpty(AsName) Then
                w.Add("--as")
                w.Add(AsName)
            End If
            If Binary Then w.Add("--binary")
            If Overwrite Then w.Add("--overwrite")
            If System Then w.Add("--system")
            If Invisible Then w.Add("--invisible")
            If Not String.IsNullOrEmpty(DateYmd) Then
                w.Add("--date")
                w.Add(DateYmd)
            End If
            If KeepGoing Then w.Add("--keep-going")
            Return w
        End Function
    End Class

    ''' <summary>
    ''' disk IMAGE import SOURCE [NAME...] [options] - copy files from another image onto the open disk.
    ''' Only call it when CanImport is True: on trs80-tool 5.2.0 and earlier "import" is an alias of "add"
    ''' and would store the source image itself on the disk.
    ''' </summary>
    Public Async Function DiskImportAsync(imagePath As String, partition As String, request As ImportRequest,
                                          Optional dryRun As Boolean = False) As Task(Of ToolResult)
        Return Await DiskChainAsync(imagePath, partition, One(request.ToWords().ToArray()), dryRun)
    End Function

End Class
