' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Dialogs.vb — the small dialogs: Extract, Preferences, Rename, Confirm File
'  Replace, Engine capabilities. (One file; each class is self-contained.)
' ============================================================================
Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

' ---------------------------------------------------------------- Extract

Public Class ExtractForm
    Inherits Form

    Private ReadOnly _txtDir As New TextBox()
    Private ReadOnly _rbSelected As New RadioButton() With {.Text = "Selected files", .AutoSize = True}
    Private ReadOnly _rbAll As New RadioButton() With {.Text = "All files shown", .AutoSize = True}
    Private ReadOnly _chkLower As New CheckBox() With {.Text = "Lower-case host file names", .AutoSize = True}
    Private ReadOnly _chkDamaged As New CheckBox() With {.Text = "Mark damaged files: add [DAMAGED] to the host name (--mark-damaged)", .AutoSize = True}
    Private ReadOnly _chkCpm As New CheckBox() With {.Text = "CP/M: keep whole 128-byte records instead of ending text at Ctrl-Z (--cpm-records)", .AutoSize = True}
    Private ReadOnly _chkUcsd As New CheckBox() With {.Text = "UCSD Pascal: keep text files as on the disk (--ucsd-pascal-raw)", .AutoSize = True}

    Public Sub New(selectedCount As Integer)
        Text = "Extract"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(520, 280)

        Controls.Add(New Label() With {.Text = "Directory to extract to:", .AutoSize = True, .Location = New Point(12, 12)})
        _txtDir.Bounds = New Rectangle(12, 32, 450, 23)
        _txtDir.Text = If(String.IsNullOrEmpty(Program.Settings.ExtractDirectory), Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Program.Settings.ExtractDirectory)
        Dim btnBrowse As New Button() With {.Text = "...", .Bounds = New Rectangle(468, 31, 40, 25)}
        AddHandler btnBrowse.Click, Sub() Browse()
        Controls.Add(_txtDir) : Controls.Add(btnBrowse)

        Dim grp As New GroupBox() With {.Text = "Files to extract", .Bounds = New Rectangle(12, 66, 250, 76)}
        _rbSelected.Location = New Point(14, 22)
        _rbAll.Location = New Point(14, 46)
        grp.Controls.Add(_rbSelected) : grp.Controls.Add(_rbAll)
        Controls.Add(grp)
        If selectedCount = 0 Then
            _rbSelected.Enabled = False
            _rbAll.Checked = True
        Else
            _rbSelected.Checked = True
        End If
        _chkLower.Location = New Point(12, 150)
        _chkLower.Checked = Program.Settings.LowercaseExtractNames
        _chkDamaged.Location = New Point(12, 174)
        _chkCpm.Location = New Point(12, 198)
        _chkUcsd.Location = New Point(12, 222)
        Controls.AddRange(New Control() {_chkLower, _chkDamaged, _chkCpm, _chkUcsd})

        Dim btnOk As New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(338, 246, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(428, 246, 80, 26)}
        AddHandler btnOk.Click, Sub(s, e)
                                    If _txtDir.Text.Trim().Length = 0 Then
                                        MessageBox.Show(Me, "Choose a directory.", "Extract", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                        DialogResult = DialogResult.None
                                    End If
                                End Sub
        Controls.Add(btnOk) : Controls.Add(btnCancel)
        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    Private Sub Browse()
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select the folder to extract to"
            fbd.UseDescriptionForTitle = True
            If System.IO.Directory.Exists(_txtDir.Text) Then fbd.SelectedPath = _txtDir.Text
            If fbd.ShowDialog(Me) = DialogResult.OK Then _txtDir.Text = fbd.SelectedPath
        End Using
    End Sub

    Public ReadOnly Property Directory As String
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
            Return _chkLower.Checked
        End Get
    End Property

    Public ReadOnly Property Options As ExtractOptions
        Get
            Return New ExtractOptions With {.Lowercase = _chkLower.Checked, .MarkDamaged = _chkDamaged.Checked, .CpmRecords = _chkCpm.Checked, .UcsdRaw = _chkUcsd.Checked}
        End Get
    End Property
End Class

''' <summary>Extraction switches the user can set; CpmRecords/UcsdRaw force the top-level "extract" command.</summary>
Public Class ExtractOptions
    Public Property Lowercase As Boolean
    Public Property MarkDamaged As Boolean
    Public Property CpmRecords As Boolean
    Public Property UcsdRaw As Boolean
    Public ReadOnly Property NeedsTopLevelExtract As Boolean
        Get
            Return CpmRecords OrElse UcsdRaw
        End Get
    End Property
End Class

' ---------------------------------------------------------------- Preferences

Public Class PreferencesForm
    Inherits Form

    Private ReadOnly _txtTool As New TextBox()
    Private ReadOnly _lblToolStatus As New Label() With {.AutoSize = True, .ForeColor = SystemColors.GrayText}
    Private ReadOnly _rbFromNone As New RadioButton() With {.Text = "No conversion", .AutoSize = True}
    Private ReadOnly _rbFromCrLf As New RadioButton() With {.Text = "Convert CR to CR/LF", .AutoSize = True}
    Private ReadOnly _rbFromLf As New RadioButton() With {.Text = "Convert CR to LF", .AutoSize = True}
    Private ReadOnly _rbToNone As New RadioButton() With {.Text = "No conversion", .AutoSize = True}
    Private ReadOnly _rbToCr As New RadioButton() With {.Text = "Convert CR/LF to CR", .AutoSize = True}
    Private ReadOnly _chkTokenize As New CheckBox() With {.Text = "Let trs80-tool tokenize BASIC text (.bas) and convert cassette programs (.cas/.3bn) when adding", .AutoSize = True}
    Private ReadOnly _chkHostDate As New CheckBox() With {.Text = "Record the Windows file's date on files added (where the DOS stores dates)", .AutoSize = True}
    Private ReadOnly _chkEsc As New CheckBox() With {.Text = "Exit upon ESC", .AutoSize = True}
    Private ReadOnly _chkConfirmDelete As New CheckBox() With {.Text = "Confirm before deleting files", .AutoSize = True}
    Private ReadOnly _chkServer As New CheckBox() With {.Text = "Keep trs80-tool running in the background (faster: one program start instead of one per action)", .AutoSize = True}
    Private ReadOnly _txtLog As New TextBox() With {.Width = 330}
    Private ReadOnly _txtDisasm As New TextBox() With {.Width = 446}

    Public Sub New()
        Text = "Preferences"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 506)
        Dim s As Settings = Program.Settings

        Dim g0 As New GroupBox() With {.Text = "trs80-tool (the engine)", .Bounds = New Rectangle(12, 10, 536, 86)}
        g0.Controls.Add(New Label() With {.Text = "Location of trs80-tool.exe (blank = ask at startup if not found):", .AutoSize = True, .Location = New Point(12, 20)})
        _txtTool.Bounds = New Rectangle(12, 40, 400, 23)
        _txtTool.Text = s.ToolPath
        Dim btnBrowse As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(420, 39, 80, 25)}
        AddHandler btnBrowse.Click, Sub() BrowseTool()
        _lblToolStatus.Location = New Point(12, 66)
        _lblToolStatus.Text = If(String.IsNullOrEmpty(Program.Tool.ToolPath), "Currently: not found", "Currently using: " & Program.Tool.ToolPath & If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, "  (v" & Program.Tool.Capabilities.Version.ToString() & ")", ""))
        g0.Controls.AddRange(New Control() {_txtTool, btnBrowse, _lblToolStatus})

        Dim g1 As New GroupBox() With {.Text = "Text conversion from TRS-80 (extract, copy, drag out)", .Bounds = New Rectangle(12, 104, 262, 100)}
        _rbFromNone.Location = New Point(14, 22) : _rbFromCrLf.Location = New Point(14, 46) : _rbFromLf.Location = New Point(14, 70)
        g1.Controls.AddRange(New Control() {_rbFromNone, _rbFromCrLf, _rbFromLf})
        Select Case s.ConversionFromTrs80
            Case TextConversionOut.CrToCrLf : _rbFromCrLf.Checked = True
            Case TextConversionOut.CrToLf : _rbFromLf.Checked = True
            Case Else : _rbFromNone.Checked = True
        End Select

        Dim g2 As New GroupBox() With {.Text = "Text conversion to TRS-80 (add, paste, drop)", .Bounds = New Rectangle(286, 104, 262, 100)}
        _rbToNone.Location = New Point(14, 22) : _rbToCr.Location = New Point(14, 46)
        g2.Controls.AddRange(New Control() {_rbToNone, _rbToCr})
        If s.ConversionToTrs80 = TextConversionIn.CrLfToCr Then _rbToCr.Checked = True Else _rbToNone.Checked = True
        g2.Controls.Add(New Label() With {.Text = "(applies to files that look like text)", .AutoSize = True, .Location = New Point(14, 72), .ForeColor = SystemColors.GrayText})

        _chkTokenize.Location = New Point(12, 216) : _chkTokenize.Checked = s.TokenizeBasicOnAdd
        _chkHostDate.Location = New Point(12, 242) : _chkHostDate.Checked = s.UseHostDateOnAdd
        _chkEsc.Location = New Point(12, 268) : _chkEsc.Checked = s.ExitOnEscape
        _chkConfirmDelete.Location = New Point(12, 294) : _chkConfirmDelete.Checked = s.ConfirmDelete
        _chkServer.Location = New Point(12, 320) : _chkServer.Checked = s.KeepEngineRunning

        Dim lblLog As New Label() With {.Text = "trs80-tool log file (--log-file on every call; blank = none):", .AutoSize = True, .Location = New Point(12, 349)}
        _txtLog.Location = New Point(12, 369)
        _txtLog.Text = s.EngineLogFile
        Dim btnLog As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(350, 368, 80, 25)}
        AddHandler btnLog.Click, Sub()
                                     Using sfd As New SaveFileDialog()
                                         sfd.Title = "trs80-tool log file"
                                         sfd.Filter = "Log files (*.log;*.txt)|*.log;*.txt|All files (*.*)|*.*"
                                         sfd.OverwritePrompt = False
                                         If sfd.ShowDialog(Me) = DialogResult.OK Then _txtLog.Text = sfd.FileName
                                     End Using
                                 End Sub
        Dim lblDisasm As New Label() With {.Text = "Disassembly options (words after ""disasm""; blank = the default):", .AutoSize = True, .Location = New Point(12, 399)}
        _txtDisasm.Location = New Point(12, 419)
        _txtDisasm.Text = s.DisasmOptions
        Dim btnDisasmDefault As New Button() With {.Text = "Default", .Bounds = New Rectangle(466, 418, 80, 25)}
        AddHandler btnDisasmDefault.Click, Sub() _txtDisasm.Text = Settings.DefaultDisasmOptions
        Dim lblSettings As New Label() With {.Text = "Settings file: " & Settings.FilePath, .AutoSize = True, .Location = New Point(12, 448), .ForeColor = SystemColors.GrayText}

        Dim btnOk As New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(378, 468, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(468, 468, 80, 26)}
        AddHandler btnOk.Click, Sub() Apply()
        Controls.AddRange(New Control() {g0, g1, g2, _chkTokenize, _chkHostDate, _chkEsc, _chkConfirmDelete, _chkServer, lblLog, _txtLog, btnLog, lblDisasm, _txtDisasm, btnDisasmDefault, lblSettings, btnOk, btnCancel})
        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    Private Sub BrowseTool()
        Using ofd As New OpenFileDialog()
            ofd.Title = "Locate trs80-tool"
            ofd.Filter = "Programs (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|All files (*.*)|*.*"
            ofd.FileName = If(_txtTool.Text.Length > 0, _txtTool.Text, Trs80Tool.DefaultExeName)
            If ofd.ShowDialog(Me) = DialogResult.OK Then _txtTool.Text = ofd.FileName
        End Using
    End Sub

    Private Sub Apply()
        Dim s As Settings = Program.Settings
        Dim p As String = _txtTool.Text.Trim()
        If p.Length > 0 AndAlso Not File.Exists(p) Then
            If MessageBox.Show(Me, p & vbCrLf & vbCrLf & "does not exist. Keep it anyway?", "Preferences", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
                DialogResult = DialogResult.None
                Return
            End If
        End If
        s.ToolPath = p
        s.ConversionFromTrs80 = If(_rbFromCrLf.Checked, TextConversionOut.CrToCrLf, If(_rbFromLf.Checked, TextConversionOut.CrToLf, TextConversionOut.None))
        s.ConversionToTrs80 = If(_rbToCr.Checked, TextConversionIn.CrLfToCr, TextConversionIn.None)
        s.TokenizeBasicOnAdd = _chkTokenize.Checked
        s.UseHostDateOnAdd = _chkHostDate.Checked
        s.ExitOnEscape = _chkEsc.Checked
        s.ConfirmDelete = _chkConfirmDelete.Checked
        s.EngineLogFile = _txtLog.Text.Trim()
        s.KeepEngineRunning = _chkServer.Checked
        s.DisasmOptions = _txtDisasm.Text.Trim()
        Program.Tool.LogFile = s.EngineLogFile
        Program.Tool.UseServer = s.KeepEngineRunning
    End Sub
End Class

' ---------------------------------------------------------------- Rename

Public Class RenameForm
    Inherits Form

    Private ReadOnly _txt As New TextBox()

    ''' <param name="message">Optional explanation shown above the box (e.g. "x is not a valid TRS-80 filename.").</param>
    Public Sub New(title As String, prompt As String, initial As String, message As String)
        Text = title
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        Dim y As Integer = 12
        If Not String.IsNullOrEmpty(message) Then
            Dim lbl As New Label() With {.Text = message, .AutoSize = False, .Bounds = New Rectangle(12, y, 356, 36), .AutoEllipsis = True}
            Controls.Add(lbl)
            y += 42
        End If
        Controls.Add(New Label() With {.Text = prompt, .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        _txt.Bounds = New Rectangle(12, y, 356, 23)
        _txt.Text = If(initial, "")
        _txt.CharacterCasing = CharacterCasing.Upper
        Controls.Add(_txt)
        y += 34
        Dim btnOk As New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(198, y, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(288, y, 80, 26)}
        AddHandler btnOk.Click, Sub()
                                    If Not TextUtil.IsValidTrsName(_txt.Text) Then
                                        MessageBox.Show(Me, """" & _txt.Text & """ is not a valid TRS-80 filename." & vbCrLf & vbCrLf & "Use 1–8 letters or digits starting with a letter, then an optional / or . and up to 3 letters or digits.", title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                        DialogResult = DialogResult.None
                                    End If
                                End Sub
        Controls.Add(btnOk) : Controls.Add(btnCancel)
        AcceptButton = btnOk : CancelButton = btnCancel
        ClientSize = New Size(380, y + 38)
        AddHandler Shown, Sub()
                              _txt.Focus()
                              Dim slash As Integer = _txt.Text.IndexOfAny(New Char() {"/"c, "."c})
                              _txt.Select(0, If(slash > 0, slash, _txt.Text.Length))
                          End Sub
    End Sub

    Public ReadOnly Property ResultName As String
        Get
            Return _txt.Text.Trim()
        End Get
    End Property
End Class

' ---------------------------------------------------------------- Confirm File Replace

Public Class ReplaceForm
    Inherits Form

    Private ReadOnly _chkAll As New CheckBox() With {.Text = "Do this for all remaining files", .AutoSize = True}

    Public Sub New(existing As DiskFile, newName As String, newSize As Long, newDate As DateTime)
        Text = "Confirm File Replace"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(400, 230)
        Controls.Add(New Label() With {.Text = "This disk image already contains a file named '" & existing.SlashName & "'.", .AutoSize = True, .Location = New Point(12, 12)})
        Controls.Add(New Label() With {.Text = "Would you like to replace the existing file", .AutoSize = True, .Location = New Point(12, 40)})
        Controls.Add(New Label() With {.Text = existing.SlashName & vbCrLf & TextUtil.FormatBytes(existing.Size) & " bytes" & vbCrLf & If(existing.DateValue.HasValue, existing.DateValue.Value.ToString("g", CultureInfo.CurrentCulture), If(existing.DateText.Length > 0, existing.DateText, "(no date)")), .AutoSize = True, .Location = New Point(30, 60)})
        Controls.Add(New Label() With {.Text = "with this one?", .AutoSize = True, .Location = New Point(12, 112)})
        Controls.Add(New Label() With {.Text = newName & vbCrLf & TextUtil.FormatBytes(newSize) & " bytes" & vbCrLf & newDate.ToString("g", CultureInfo.CurrentCulture), .AutoSize = True, .Location = New Point(30, 132)})
        _chkAll.Location = New Point(12, 186)
        Controls.Add(_chkAll)
        Dim btnYes As New Button() With {.Text = "Yes", .DialogResult = DialogResult.Yes, .Bounds = New Rectangle(218, 196, 80, 26)}
        Dim btnNo As New Button() With {.Text = "No", .DialogResult = DialogResult.No, .Bounds = New Rectangle(308, 196, 80, 26)}
        Controls.Add(btnYes) : Controls.Add(btnNo)
        AcceptButton = btnNo : CancelButton = btnNo
    End Sub

    Public ReadOnly Property ApplyToAll As Boolean
        Get
            Return _chkAll.Checked
        End Get
    End Property
End Class

' ---------------------------------------------------------------- Engine capabilities

Public Class EngineInfoForm
    Inherits Form

    Private ReadOnly _tool As Trs80Tool
    Private ReadOnly _txt As New TextBox()
    ''' <summary>Set when the user asked to re-run the engine probe. Not a designer property (WFO1000).</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ReprobeRequested As Boolean = False

    Public Sub New(tool As Trs80Tool)
        _tool = tool
        Text = "trs80-tool engine capabilities"
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        Size = New Size(760, 560)
        MinimumSize = New Size(500, 300)
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font("Consolas", 9.0F)
            .Dock = DockStyle.Fill
        End With
        Dim panel As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        Dim btnReprobe As New Button() With {.Text = "Re-check engine", .Bounds = New Rectangle(8, 7, 130, 26)}
        AddHandler btnReprobe.Click, Sub()
                                         ReprobeRequested = True
                                         Close()
                                     End Sub
        Dim btnCopy As New Button() With {.Text = "Copy to clipboard", .Bounds = New Rectangle(146, 7, 130, 26)}
        AddHandler btnCopy.Click, Sub()
                                      Try : Clipboard.SetText(_txt.Text) : Catch : End Try
                                  End Sub
        Dim btnClose As New Button() With {.Text = "Close", .DialogResult = DialogResult.OK, .Anchor = AnchorStyles.Right Or AnchorStyles.Top, .Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)}
        panel.Controls.AddRange(New Control() {btnReprobe, btnCopy, btnClose})
        Controls.Add(_txt)
        Controls.Add(panel)
        CancelButton = btnClose
        _txt.Text = BuildReport()
        _txt.Select(0, 0)
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
        Return sb.ToString().Replace(vbCr, "").Replace(vbLf, vbCrLf)
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
