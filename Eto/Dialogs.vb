' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Dialogs.vb — the small dialogs of the Eto.Forms program: Extract,
'  Preferences, Rename, Confirm File Replace, Confirm File Overwrite, the
'  emulator's model choice and Help > Engine capabilities.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports Eto.Drawing
Imports Eto.Forms

''' <summary>Edit → Extract: where to, which files, and trs80-tool's extraction switches.</summary>
Public Class ExtractForm
    Inherits OkCancelDialog

    Private ReadOnly _txtDir As New TextBox With {.Width = 420}
    Private ReadOnly _rbSelected As RadioButton
    Private ReadOnly _rbAll As RadioButton
    Private ReadOnly _chkLower As New CheckBox With {.Text = "Lower-case host file names"}
    Private ReadOnly _chkDamaged As New CheckBox With {.Text = "Mark damaged files: add [DAMAGED] to the host name (--mark-damaged)"}
    Private ReadOnly _chkCpm As New CheckBox With {.Text = "CP/M: keep whole 128-byte records instead of ending text at Ctrl-Z (--cpm-records)"}
    Private ReadOnly _chkUcsd As New CheckBox With {.Text = "UCSD Pascal: keep text files as on the disk (--ucsd-pascal-raw)"}

    Public Sub New(selectedCount As Integer)
        MyBase.New("Extract")
        _rbSelected = New RadioButton With {.Text = "Selected files"}
        _rbAll = New RadioButton(_rbSelected) With {.Text = "All files shown"}
        Dim startDir As String = Program.Settings.ExtractDirectory
        If String.IsNullOrEmpty(startDir) Then startDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ' On Linux that is empty when the user has no Documents folder: start from the home folder then.
        If String.IsNullOrEmpty(startDir) Then startDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        _txtDir.Text = startDir
        If selectedCount = 0 Then
            _rbSelected.Enabled = False
            _rbAll.Checked = True
        Else
            _rbSelected.Checked = True
        End If
        _chkLower.Checked = Program.Settings.LowercaseExtractNames
        Finish(Ui.Down(Nothing,
                       Ui.Lbl("Directory to extract to:"),
                       Ui.Across(_txtDir, Ui.Btn("Browse...", AddressOf Browse)),
                       Ui.Across(Ui.Group("Files to extract", Ui.Down(Nothing, _rbSelected, _rbAll))),
                       _chkLower, _chkDamaged, _chkCpm, _chkUcsd))
    End Sub

    Private Sub Browse()
        Dim chosen As String = Ui.PickFolder(Me, "Select the folder to extract to", _txtDir.Text)
        If chosen IsNot Nothing Then _txtDir.Text = chosen
    End Sub

    Protected Overrides Function Accept() As Boolean
        If _txtDir.Text.Trim().Length = 0 Then
            Ui.Warn(Me, "Choose a directory.", "Extract")
            Return False
        End If
        Return True
    End Function

    Public ReadOnly Property TargetDirectory As String
        Get
            Return _txtDir.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property ExtractAll As Boolean
        Get
            Return _rbAll.Checked
        End Get
    End Property

    Public ReadOnly Property Lowercase As Boolean
        Get
            Return Ui.IsOn(_chkLower)
        End Get
    End Property

    Public ReadOnly Property Options As ExtractOptions
        Get
            Return New ExtractOptions With {.Lowercase = Ui.IsOn(_chkLower), .MarkDamaged = Ui.IsOn(_chkDamaged), .CpmRecords = Ui.IsOn(_chkCpm), .UcsdRaw = Ui.IsOn(_chkUcsd)}
        End Get
    End Property
End Class

''' <summary>The extraction switches chosen in the Extract dialog.</summary>
Public Class ExtractOptions
    Public Property Lowercase As Boolean
    Public Property MarkDamaged As Boolean
    Public Property CpmRecords As Boolean
    Public Property UcsdRaw As Boolean

    ''' <summary>--cpm-records and --ucsd-pascal-raw exist only on trs80-tool's top-level extract command.</summary>
    Public ReadOnly Property NeedsTopLevelExtract As Boolean
        Get
            Return CpmRecords OrElse UcsdRaw
        End Get
    End Property
End Class

''' <summary>Edit → Preferences.</summary>
Public Class PreferencesForm
    Inherits OkCancelDialog

    Private ReadOnly _txtTool As New TextBox With {.Width = 420}
    Private ReadOnly _rbFromNone, _rbFromCrLf, _rbFromLf, _rbToNone, _rbToCr As RadioButton
    Private ReadOnly _chkTokenize As New CheckBox With {.Text = "Let trs80-tool tokenize BASIC text (.bas) and convert cassette programs (.cas/.3bn) when adding"}
    Private ReadOnly _chkHostDate As New CheckBox With {.Text = "Record the host file's date on files added (where the DOS stores dates)"}
    Private ReadOnly _chkUnzipGz As New CheckBox With {.Text = "Unzip files named *.gz when adding them, and store the file inside (off = store the .gz file as it is)"}
    Private ReadOnly _chkEsc As New CheckBox With {.Text = "Exit upon ESC"}
    Private ReadOnly _chkConfirmDelete As New CheckBox With {.Text = "Confirm before deleting files"}
    Private ReadOnly _chkServer As New CheckBox With {.Text = "Keep trs80-tool running in the background (faster: one program start instead of one per action)"}
    Private ReadOnly _txtLog As New TextBox With {.Width = 420}
    Private ReadOnly _txtDisasm As New TextBox With {.Width = 420}
    Private ReadOnly _txtTerminal As New TextBox With {.Width = 420}

    Public Sub New()
        MyBase.New("Preferences")
        Dim s As Settings = Program.Settings
        _txtTool.Text = s.ToolPath
        Dim status As String = If(String.IsNullOrEmpty(Program.Tool.ToolPath), "Currently: not found", "Currently using: " & Program.Tool.ToolPath & If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, "  (v" & Program.Tool.Capabilities.Version.ToString() & ")", ""))
        Dim g0 As GroupBox = Ui.Group("trs80-tool (the engine)", Ui.Down(Nothing,
            Ui.Lbl("Location of " & Trs80Tool.DefaultExeName & " (blank = look beside this program and on the PATH, then ask):"),
            Ui.Across(_txtTool, Ui.Btn("Browse...", AddressOf BrowseTool)),
            Ui.Lbl(status, True)))

        _rbFromNone = New RadioButton With {.Text = "No conversion"}
        _rbFromCrLf = New RadioButton(_rbFromNone) With {.Text = "Convert CR to CR/LF"}
        _rbFromLf = New RadioButton(_rbFromNone) With {.Text = "Convert CR to LF"}
        Select Case s.ConversionFromTrs80
            Case TextConversionOut.CrToCrLf : _rbFromCrLf.Checked = True
            Case TextConversionOut.CrToLf : _rbFromLf.Checked = True
            Case Else : _rbFromNone.Checked = True
        End Select
        Dim g1 As GroupBox = Ui.Group("Text conversion from TRS-80 (extract, copy)", Ui.Down(Nothing, _rbFromNone, _rbFromCrLf, _rbFromLf))

        _rbToNone = New RadioButton With {.Text = "No conversion"}
        _rbToCr = New RadioButton(_rbToNone) With {.Text = "Convert line ends to CR"}
        If s.ConversionToTrs80 = TextConversionIn.CrLfToCr Then _rbToCr.Checked = True Else _rbToNone.Checked = True
        Dim g2 As GroupBox = Ui.Group("Text conversion to TRS-80 (add, paste, drop)", Ui.Down(Nothing, _rbToNone, _rbToCr, Ui.Lbl("(applies to files that look like text)", True)))

        _chkTokenize.Checked = s.TokenizeBasicOnAdd
        _chkHostDate.Checked = s.UseHostDateOnAdd
        _chkUnzipGz.Checked = s.UnzipGzOnAdd
        _chkEsc.Checked = s.ExitOnEscape
        _chkConfirmDelete.Checked = s.ConfirmDelete
        _chkServer.Checked = s.KeepEngineRunning
        _txtLog.Text = s.EngineLogFile
        _txtDisasm.Text = s.DisasmOptions
        _txtTerminal.Text = s.TerminalCommand

        Dim pair As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        pair.Rows.Add(New TableRow(New TableCell(g1, True), New TableCell(g2, True)))

        Dim rows As New List(Of Control) From {
            g0, pair, _chkTokenize, _chkHostDate, _chkUnzipGz, _chkEsc, _chkConfirmDelete, _chkServer,
            Ui.Lbl("trs80-tool log file (--log-file on every call; blank = none):"),
            Ui.Across(_txtLog, Ui.Btn("Browse...", AddressOf BrowseLog)),
            Ui.Lbl("Disassembly options (words after ""disasm""; blank = the default):"),
            Ui.Across(_txtDisasm, Ui.Btn("Default", Sub() _txtDisasm.Text = Settings.DefaultDisasmOptions))}
        If Not HostPlatform.IsWindows Then
            rows.Add(Ui.Lbl("Terminal for the trs80-tool shell and the emulator (blank = find one; the command line is added after it):"))
            rows.Add(Ui.Across(_txtTerminal, Ui.Lbl("e.g.  xterm -e", True)))
        End If
        rows.Add(Ui.Lbl("Settings file: " & Settings.FilePath, True))
        Finish(Ui.Down(Nothing, rows.ToArray()))
    End Sub

    Private Sub BrowseTool()
        Dim chosen As String = Ui.PickFile(Me, "Locate trs80-tool", Program.ToolFileFilter, If(_txtTool.Text.Trim().Length > 0, _txtTool.Text.Trim(), Nothing))
        If chosen IsNot Nothing Then _txtTool.Text = chosen
    End Sub

    Private Sub BrowseLog()
        Dim chosen As String = Ui.PickSave(Me, "trs80-tool log file", "Log files (*.log;*.txt)|*.log;*.txt|All files (*.*)|*.*", Nothing, "log")
        If chosen IsNot Nothing Then _txtLog.Text = chosen
    End Sub

    Protected Overrides Function Accept() As Boolean
        Dim s As Settings = Program.Settings
        Dim p As String = _txtTool.Text.Trim()
        If p.Length > 0 AndAlso Not File.Exists(p) Then
            If Not Ui.Ask(Me, p & vbLf & vbLf & "does not exist. Keep it anyway?", "Preferences", MessageBoxType.Warning) Then Return False
        End If
        s.ToolPath = p
        s.ConversionFromTrs80 = If(_rbFromCrLf.Checked, TextConversionOut.CrToCrLf, If(_rbFromLf.Checked, TextConversionOut.CrToLf, TextConversionOut.None))
        s.ConversionToTrs80 = If(_rbToCr.Checked, TextConversionIn.CrLfToCr, TextConversionIn.None)
        s.TokenizeBasicOnAdd = Ui.IsOn(_chkTokenize)
        s.UseHostDateOnAdd = Ui.IsOn(_chkHostDate)
        s.UnzipGzOnAdd = Ui.IsOn(_chkUnzipGz)
        s.ExitOnEscape = Ui.IsOn(_chkEsc)
        s.ConfirmDelete = Ui.IsOn(_chkConfirmDelete)
        s.EngineLogFile = _txtLog.Text.Trim()
        s.KeepEngineRunning = Ui.IsOn(_chkServer)
        s.DisasmOptions = _txtDisasm.Text.Trim()
        s.TerminalCommand = _txtTerminal.Text.Trim()
        Program.Tool.LogFile = s.EngineLogFile
        Program.Tool.UseServer = s.KeepEngineRunning
        Program.Tool.TerminalCommand = s.TerminalCommand
        Return True
    End Function
End Class

''' <summary>Ask for a TRS-80 file name (rename, or a host file whose name the disk can't hold).</summary>
Public Class RenameForm
    Inherits OkCancelDialog

    Private ReadOnly _txt As New TextBox With {.Width = 320}
    Private ReadOnly _title As String

    Public Sub New(title As String, prompt As String, initial As String, message As String)
        MyBase.New(title)
        _title = title
        _txt.Text = If(initial, "").ToUpperInvariant()
        Dim rows As New List(Of Control)
        If Not String.IsNullOrEmpty(message) Then rows.Add(Ui.Lbl(message))
        rows.Add(Ui.Lbl(prompt))
        rows.Add(_txt)
        Finish(Ui.Down(Nothing, rows.ToArray()))
        AddHandler Shown, Sub()
                              _txt.Focus()
                              Dim slash As Integer = _txt.Text.IndexOfAny(New Char() {"/"c, "."c})
                              Try
                                  _txt.Selection = New Range(Of Integer)(0, If(slash > 0, slash, _txt.Text.Length) - 1)
                              Catch
                              End Try
                          End Sub
    End Sub

    Protected Overrides Function Accept() As Boolean
        If Not TextUtil.IsValidTrsName(ResultName) Then
            Ui.Warn(Me, """" & _txt.Text & """ is not a valid TRS-80 filename." & vbLf & vbLf & "Use 1–8 letters or digits starting with a letter, then an optional / or . and up to 3 letters or digits.", _title)
            Return False
        End If
        Return True
    End Function

    ''' <summary>The name typed, in upper case, as TRS-80 names are.</summary>
    Public ReadOnly Property ResultName As String
        Get
            Return _txt.Text.Trim().ToUpperInvariant()
        End Get
    End Property
End Class

''' <summary>A file being added is already on the disk: replace it?</summary>
Public Class ReplaceForm
    Inherits Dialog(Of Boolean)

    Private ReadOnly _chkAll As New CheckBox With {.Text = "Do this for all remaining files"}

    Public Sub New(existing As DiskFile, newName As String, newSize As Long, newDate As DateTime)
        Title = "Confirm File Replace"
        Resizable = False
        Padding = New Padding(Ui.Margin)
        Dim oldText As String = existing.SlashName & vbLf & TextUtil.FormatBytes(existing.Size) & " bytes" & vbLf & If(existing.DateValue.HasValue, existing.DateValue.Value.ToString("g", CultureInfo.CurrentCulture), If(existing.DateText.Length > 0, existing.DateText, "(no date)"))
        Dim newText As String = newName & vbLf & TextUtil.FormatBytes(newSize) & " bytes" & vbLf & newDate.ToString("g", CultureInfo.CurrentCulture)
        Content = Ui.Down(Nothing,
                          Ui.Lbl("This disk image already contains a file named '" & existing.SlashName & "'."),
                          Ui.Lbl("Would you like to replace the existing file"),
                          New Panel With {.Padding = New Padding(18, 0, 0, 0), .Content = Ui.Lbl(oldText)},
                          Ui.Lbl("with this one?"),
                          New Panel With {.Padding = New Padding(18, 0, 0, 0), .Content = Ui.Lbl(newText)},
                          _chkAll)
        Dim btnYes As New Button With {.Text = "Yes"}
        Dim btnNo As New Button With {.Text = "No"}
        AddHandler btnYes.Click, Sub() Close(True)
        AddHandler btnNo.Click, Sub() Close(False)
        PositiveButtons.Add(btnYes)
        NegativeButtons.Add(btnNo)
        DefaultButton = btnNo
        AbortButton = btnNo
    End Sub

    ''' <summary>True when the user chose Yes.</summary>
    Public Function ShowReplace(parent As Control) As Boolean
        Return ShowModal(parent)
    End Function

    Public ReadOnly Property ApplyToAll As Boolean
        Get
            Return Ui.IsOn(_chkAll)
        End Get
    End Property
End Class

''' <summary>A file being extracted exists in the target folder: overwrite it, skip it, or stop.</summary>
Public Class OverwriteForm
    Inherits Dialog(Of DialogResult)

    Private ReadOnly _chkAll As New CheckBox With {.Text = "Do this for all remaining files"}

    Public Sub New(fileName As String, folder As String, offerAll As Boolean)
        Title = "Confirm File Overwrite"
        Resizable = False
        Padding = New Padding(Ui.Margin)
        Result = DialogResult.Cancel
        Dim rows As New List(Of Control) From {Ui.Lbl(fileName & " exists in " & folder & "."), Ui.Lbl("Overwrite it?"), Ui.Lbl("(Cancel stops extracting.)", True)}
        If offerAll Then rows.Add(_chkAll)
        Content = Ui.Down(Nothing, rows.ToArray())
        Dim btnYes As New Button With {.Text = "Yes"}
        Dim btnNo As New Button With {.Text = "No"}
        Dim btnCancel As New Button With {.Text = "Cancel"}
        AddHandler btnYes.Click, Sub() Close(DialogResult.Yes)
        AddHandler btnNo.Click, Sub() Close(DialogResult.No)
        AddHandler btnCancel.Click, Sub() Close(DialogResult.Cancel)
        PositiveButtons.Add(btnYes)
        NegativeButtons.Add(btnNo)
        NegativeButtons.Add(btnCancel)
        DefaultButton = btnNo
        AbortButton = btnCancel
    End Sub

    Public Function ShowChoice(parent As Control) As DialogResult
        Return ShowModal(parent)
    End Function

    Public ReadOnly Property ApplyToAll As Boolean
        Get
            Return Ui.IsOn(_chkAll)
        End Get
    End Property
End Class

''' <summary>Tools → Run the emulator without a disk: which machine.</summary>
Public Class RunModelForm
    Inherits OkCancelDialog

    Private ReadOnly _cmb As New DropDown With {.Width = 280}

    Public Sub New()
        MyBase.New("Run the emulator")
        _cmb.Items.Add("Model I (Level II BASIC)")
        _cmb.Items.Add("Model III")
        _cmb.Items.Add("Model 4")
        _cmb.SelectedIndex = 1
        Finish(Ui.Down(Nothing, Ui.Lbl("Machine to emulate (no disk mounted):"), _cmb), "Run")
    End Sub

    Public ReadOnly Property Model As Integer
        Get
            Return If(_cmb.SelectedIndex = 0, 1, If(_cmb.SelectedIndex = 1, 3, 4))
        End Get
    End Property
End Class

''' <summary>Help → Engine capabilities: what the probe found, how the engine is called, and the recent calls.</summary>
Public Class EngineInfoForm
    Inherits Dialog

    Private ReadOnly _tool As Trs80Tool
    Private ReadOnly _txt As TextArea
    Public Property ReprobeRequested As Boolean = False

    Public Sub New(tool As Trs80Tool)
        _tool = tool
        Title = "trs80-tool engine capabilities"
        Resizable = True
        ClientSize = New Size(820, 560)
        Padding = New Padding(Ui.Margin)
        _txt = Ui.ReportBox(BuildReport(), 9.0F)
        Dim btnClose As Button = Ui.Btn("Close", Sub() Close())
        Dim buttons As TableLayout = Ui.Across(Ui.Btn("Re-check engine", Sub()
                                                                            ReprobeRequested = True
                                                                            Close()
                                                                        End Sub),
                                               Ui.Btn("Copy to clipboard", Sub() Ui.CopyTextToClipboard(_txt.Text)), Nothing, btnClose)
        Content = Ui.Down(_txt, _txt, buttons)
        AbortButton = btnClose
        DefaultButton = btnClose
    End Sub

    Private Function BuildReport() As String
        Dim sb As New StringBuilder()
        Dim caps As EngineCapabilities = _tool.Capabilities
        sb.AppendLine(Program.AppName & " " & Program.AppVersion)
        sb.AppendLine("Engine path:        " & If(String.IsNullOrEmpty(_tool.ToolPath), "(not found)", _tool.ToolPath))
        sb.AppendLine("Minimum version:    " & Trs80Tool.MinimumVersion.ToString())
        sb.AppendLine("Last tested with:   " & Trs80Tool.MaxTestedVersion.ToString())
        If caps Is Nothing Then
            sb.AppendLine("The engine has not been probed (not found or not yet run).")
            Return sb.ToString()
        End If
        sb.AppendLine("Reported version:   " & caps.VersionText)
        sb.AppendLine("Parsed version:     " & If(caps.Version Is Nothing, "(unparseable)", caps.Version.ToString()))
        sb.AppendLine()
        sb.AppendLine("How it is called:   " & _tool.TransportText)
        sb.AppendLine()
        sb.AppendLine("Top-level commands this program uses: " & String.Join(", ", Trs80Tool.ExpectedTopLevel))
        sb.AppendLine("  this trs80-tool has: " & If(caps.TopLevelCommands.Count = 0, "(not reported)", String.Join(", ", caps.TopLevelCommands)))
        sb.AppendLine("disk commands this program uses:      " & String.Join(", ", Trs80Tool.ExpectedDisk))
        sb.AppendLine("  this trs80-tool has: " & If(caps.DiskCommands.Count = 0, "(not reported)", String.Join(", ", caps.DiskCommands)))
        sb.AppendLine()
        sb.AppendLine("DOSes it can format a disk for:  " & DosList(caps.FormatDoses))
        sb.AppendLine("DOSes whose disks it can change: " & DosList(caps.WritableDoses))
        sb.AppendLine("Image formats: reads " & String.Join(", ", caps.ReadFormats) & "; converts to " & String.Join(", ", caps.ConvertFormats) & "; changes " & String.Join(", ", caps.DiskWriteFormats) & "; creates " & String.Join(", ", caps.DiskCreateFormats))
        sb.AppendLine()
        If caps.Problems.Count = 0 Then
            sb.AppendLine("No problems found.")
        Else
            sb.AppendLine("Problems / notes:")
            For Each p In caps.Problems
                sb.AppendLine("  - " & p)
            Next
        End If
        sb.AppendLine()
        sb.AppendLine("Change log for new versions: https://www.my-trs-80.com/tool/#change-log")
        sb.AppendLine("disk command reference:       https://www.my-trs-80.com/tool/disk.html")
        sb.AppendLine()
        sb.AppendLine("---- recent engine calls (newest last) ----")
        SyncLock _tool.CallLog
            Dim start As Integer = Math.Max(0, _tool.CallLog.Count - 60)
            For i As Integer = start To _tool.CallLog.Count - 1
                sb.AppendLine(_tool.CallLog(i))
            Next
        End SyncLock
        sb.AppendLine()
        sb.AppendLine("---- options of each command (from ""trs80-tool help --json"") ----")
        For Each cmdName In caps.TopLevelCommands
            Dim opts As List(Of String) = Nothing
            If caps.TopLevelOptions.TryGetValue(cmdName, opts) Then sb.AppendLine(cmdName & ": " & String.Join(" ", opts))
        Next
        sb.AppendLine("disk (whole run): " & String.Join(" ", caps.DiskGlobalOptions))
        For Each cmdName In caps.DiskCommands
            Dim opts As List(Of String) = Nothing
            If caps.DiskOptions.TryGetValue(cmdName, opts) Then sb.AppendLine("disk " & cmdName & ": " & String.Join(" ", opts))
        Next
        Return sb.ToString()
    End Function

    Private Shared Function DosList(doses As List(Of Trs80Tool.HelpDosJson)) As String
        If doses Is Nothing OrElse doses.Count = 0 Then Return "(not reported)"
        Dim names As New List(Of String)
        For Each d In doses
            names.Add(If(String.IsNullOrEmpty(d.Label), d.Name, d.Label))
        Next
        Return String.Join(", ", names)
    End Function
End Class
