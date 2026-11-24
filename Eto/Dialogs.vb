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
    Private ReadOnly _chkLower As New CheckBox With {.Text = UiText.T("dlg.extract.lower")}
    Private ReadOnly _chkDamaged As New CheckBox With {.Text = UiText.T("dlg.extract.damaged")}
    Private ReadOnly _chkCpm As New CheckBox With {.Text = UiText.T("dlg.extract.cpm")}
    Private ReadOnly _chkUcsd As New CheckBox With {.Text = UiText.T("dlg.extract.ucsd")}

    Public Sub New(selectedCount As Integer)
        MyBase.New(UiText.T("dlg.extract.title"))
        _rbSelected = New RadioButton With {.Text = UiText.T("dlg.extract.selected")}
        _rbAll = New RadioButton(_rbSelected) With {.Text = UiText.T("dlg.extract.all")}
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
                       Ui.Lbl(UiText.T("dlg.extract.dir")),
                       Ui.Across(_txtDir, Ui.Btn(UiText.T("btn.browse"), AddressOf Browse)),
                       Ui.Across(Ui.Group(UiText.T("dlg.extract.files"), Ui.Down(Nothing, _rbSelected, _rbAll))),
                       _chkLower, _chkDamaged, _chkCpm, _chkUcsd))
    End Sub

    Private Sub Browse()
        Dim chosen As String = Ui.PickFolder(Me, UiText.T("dlg.extract.pickfolder"), _txtDir.Text)
        If chosen IsNot Nothing Then _txtDir.Text = chosen
    End Sub

    Protected Overrides Function Accept() As Boolean
        If _txtDir.Text.Trim().Length = 0 Then
            Ui.Warn(Me, UiText.T("dlg.extract.nodir"), UiText.T("dlg.extract.title"))
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

''' <summary>Edit → Preferences (on a Mac: the application menu).</summary>
Public Class PreferencesForm
    Inherits OkCancelDialog

    Private ReadOnly _txtTool As New TextBox With {.Width = 420}
    Private ReadOnly _rbFromNone, _rbFromCrLf, _rbFromLf, _rbToNone, _rbToCr As RadioButton
    Private ReadOnly _chkTokenize As New CheckBox With {.Text = UiText.T("dlg.prefs.tokenize")}
    Private ReadOnly _chkHostDate As New CheckBox With {.Text = UiText.T("dlg.prefs.hostdate.unix")}
    Private ReadOnly _chkUnzipGz As New CheckBox With {.Text = UiText.T("dlg.prefs.unzipgz")}
    Private ReadOnly _chkEsc As New CheckBox With {.Text = UiText.T("dlg.prefs.esc")}
    Private ReadOnly _chkConfirmDelete As New CheckBox With {.Text = UiText.T("dlg.prefs.confirmdelete")}
    Private ReadOnly _chkServer As New CheckBox With {.Text = UiText.T("dlg.prefs.server")}
    Private ReadOnly _cmbLanguage As New DropDown With {.Width = 200}
    Private ReadOnly _txtLog As New TextBox With {.Width = 420}
    Private ReadOnly _txtDisasm As New TextBox With {.Width = 420}
    Private ReadOnly _txtTerminal As New TextBox With {.Width = 420}

    Public Sub New()
        MyBase.New(UiText.T("dlg.prefs.title"))
        Dim s As Settings = Program.Settings
        _txtTool.Text = s.ToolPath
        Dim status As String = If(String.IsNullOrEmpty(Program.Tool.ToolPath), UiText.T("dlg.prefs.toolnotfound"), If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, UiText.T("dlg.prefs.toolversion", UiText.T("dlg.prefs.toolusing", Program.Tool.ToolPath), Program.Tool.Capabilities.Version.ToString()), UiText.T("dlg.prefs.toolusing", Program.Tool.ToolPath)))
        Dim g0 As GroupBox = Ui.Group(UiText.T("dlg.prefs.engine"), Ui.Down(Nothing,
            Ui.Lbl(UiText.T("dlg.prefs.toollocation.unix", Trs80Tool.DefaultExeName)),
            Ui.Across(_txtTool, Ui.Btn(UiText.T("btn.browse"), AddressOf BrowseTool)),
            Ui.Lbl(status, True)))

        _rbFromNone = New RadioButton With {.Text = UiText.T("dlg.prefs.noconv")}
        _rbFromCrLf = New RadioButton(_rbFromNone) With {.Text = UiText.T("dlg.prefs.crtocrlf")}
        _rbFromLf = New RadioButton(_rbFromNone) With {.Text = UiText.T("dlg.prefs.crtolf")}
        Select Case s.ConversionFromTrs80
            Case TextConversionOut.CrToCrLf : _rbFromCrLf.Checked = True
            Case TextConversionOut.CrToLf : _rbFromLf.Checked = True
            Case Else : _rbFromNone.Checked = True
        End Select
        Dim g1 As GroupBox = Ui.Group(UiText.T("dlg.prefs.fromtrs.unix"), Ui.Down(Nothing, _rbFromNone, _rbFromCrLf, _rbFromLf))

        _rbToNone = New RadioButton With {.Text = UiText.T("dlg.prefs.noconv")}
        _rbToCr = New RadioButton(_rbToNone) With {.Text = UiText.T("dlg.prefs.lineendstocr")}
        If s.ConversionToTrs80 = TextConversionIn.CrLfToCr Then _rbToCr.Checked = True Else _rbToNone.Checked = True
        Dim g2 As GroupBox = Ui.Group(UiText.T("dlg.prefs.totrs"), Ui.Down(Nothing, _rbToNone, _rbToCr, Ui.Lbl(UiText.T("dlg.prefs.textonly"), True)))

        _chkTokenize.Checked = s.TokenizeBasicOnAdd
        _chkHostDate.Checked = s.UseHostDateOnAdd
        _chkUnzipGz.Checked = s.UnzipGzOnAdd
        _chkEsc.Checked = s.ExitOnEscape
        _chkConfirmDelete.Checked = s.ConfirmDelete
        _chkServer.Checked = s.KeepEngineRunning
        _txtLog.Text = s.EngineLogFile
        _txtDisasm.Text = s.DisasmOptions
        _txtTerminal.Text = s.TerminalCommand
        ' The language of the program's own text. Only the languages whose table is in this build are
        ' offered, and a change takes effect at the next start: text already on screen is not revisited.
        For Each lang As UiText.Language In UiText.Languages
            _cmbLanguage.Items.Add(lang.NativeName)
        Next
        _cmbLanguage.SelectedIndex = LanguageIndex()

        Dim pair As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        pair.Rows.Add(New TableRow(New TableCell(g1, True), New TableCell(g2, True)))

        Dim rows As New List(Of Control) From {
            g0, pair, _chkTokenize, _chkHostDate, _chkUnzipGz, _chkEsc, _chkConfirmDelete, _chkServer,
            Ui.Lbl(UiText.T("dlg.prefs.language")),
            Ui.Across(_cmbLanguage, Ui.Lbl(UiText.T("dlg.prefs.langnote"), True)),
            Ui.Lbl(UiText.T("dlg.prefs.logfile")),
            Ui.Across(_txtLog, Ui.Btn(UiText.T("btn.browse"), AddressOf BrowseLog)),
            Ui.Lbl(UiText.T("dlg.prefs.disasm")),
            Ui.Across(_txtDisasm, Ui.Btn(UiText.T("btn.default"), Sub() _txtDisasm.Text = Settings.DefaultDisasmOptions))}
        If HostPlatform.IsMacOS Then
            rows.Add(Ui.Lbl(UiText.T("dlg.prefs.terminal.mac")))
            rows.Add(Ui.Across(_txtTerminal, Ui.Lbl(UiText.T("dlg.prefs.terminalhint.mac"), True)))
        ElseIf Not HostPlatform.IsWindows Then
            rows.Add(Ui.Lbl(UiText.T("dlg.prefs.terminal.linux")))
            rows.Add(Ui.Across(_txtTerminal, Ui.Lbl(UiText.T("dlg.prefs.terminalhint.linux"), True)))
        End If
        rows.Add(Ui.Lbl(UiText.T("dlg.prefs.settingsfile", Settings.FilePath), True))
        Finish(Ui.Down(Nothing, rows.ToArray()))
    End Sub

    ''' <summary>Where the language in force sits in the offered list, or 0 (English) if it is not in it.</summary>
    Private Function LanguageIndex() As Integer
        Dim want As String = If(Program.Settings.Language, "").Trim()
        If want.Length = 0 Then want = UiText.DefaultCode
        For i As Integer = 0 To UiText.Languages.Count - 1
            If String.Equals(UiText.Languages(i).Code, want, StringComparison.OrdinalIgnoreCase) Then Return i
        Next
        Return 0
    End Function

    Private Sub BrowseTool()
        Dim chosen As String = Ui.PickFile(Me, UiText.T("dlg.prefs.picktool"), Program.ToolFileFilter, If(_txtTool.Text.Trim().Length > 0, _txtTool.Text.Trim(), Nothing))
        If chosen IsNot Nothing Then _txtTool.Text = chosen
    End Sub

    Private Sub BrowseLog()
        Dim chosen As String = Ui.PickSave(Me, UiText.T("dlg.prefs.picklog"), UiText.T("filter.log"), Nothing, "log")
        If chosen IsNot Nothing Then _txtLog.Text = chosen
    End Sub

    Protected Overrides Function Accept() As Boolean
        Dim s As Settings = Program.Settings
        Dim p As String = _txtTool.Text.Trim()
        If p.Length > 0 AndAlso Not File.Exists(p) Then
            If Not Ui.Ask(Me, UiText.T("dlg.prefs.notexist", p), UiText.T("dlg.prefs.title"), MessageBoxType.Warning) Then Return False
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
        ' Saved, but not put in force here: the language is read at the next start.
        Dim li As Integer = _cmbLanguage.SelectedIndex
        If li >= 0 AndAlso li < UiText.Languages.Count Then s.Language = UiText.Languages(li).Code
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
            Ui.Warn(Me, UiText.T("dlg.rename.invalid", _txt.Text), _title)
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

    Private ReadOnly _chkAll As New CheckBox With {.Text = UiText.T("dlg.replace.all")}

    Public Sub New(existing As DiskFile, newName As String, newSize As Long, newDate As DateTime)
        Title = UiText.T("dlg.replace.title")
        Resizable = False
        Padding = New Padding(Ui.Margin)
        Dim oldText As String = existing.SlashName & vbLf & TextUtil.FormatBytes(existing.Size) & " bytes" & vbLf & If(existing.DateValue.HasValue, existing.DateValue.Value.ToString("g", CultureInfo.CurrentCulture), If(existing.DateText.Length > 0, existing.DateText, "(no date)"))
        Dim newText As String = newName & vbLf & TextUtil.FormatBytes(newSize) & " bytes" & vbLf & newDate.ToString("g", CultureInfo.CurrentCulture)
        Content = Ui.Down(Nothing,
                          Ui.Lbl(UiText.T("dlg.replace.contains", existing.SlashName)),
                          Ui.Lbl(UiText.T("dlg.replace.would")),
                          New Panel With {.Padding = New Padding(18, 0, 0, 0), .Content = Ui.Lbl(oldText)},
                          Ui.Lbl(UiText.T("dlg.replace.withthis")),
                          New Panel With {.Padding = New Padding(18, 0, 0, 0), .Content = Ui.Lbl(newText)},
                          _chkAll)
        Dim btnYes As New Button With {.Text = UiText.T("btn.yes")}
        Dim btnNo As New Button With {.Text = UiText.T("btn.no")}
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

    Private ReadOnly _chkAll As New CheckBox With {.Text = UiText.T("dlg.replace.all")}

    Public Sub New(fileName As String, folder As String, offerAll As Boolean)
        Title = UiText.T("dlg.overwrite.title")
        Resizable = False
        Padding = New Padding(Ui.Margin)
        Result = DialogResult.Cancel
        Dim rows As New List(Of Control) From {Ui.Lbl(UiText.T("dlg.overwrite.exists", fileName, folder)), Ui.Lbl(UiText.T("dlg.overwrite.question")), Ui.Lbl(UiText.T("dlg.overwrite.cancelnote"), True)}
        If offerAll Then rows.Add(_chkAll)
        Content = Ui.Down(Nothing, rows.ToArray())
        Dim btnYes As New Button With {.Text = UiText.T("btn.yes")}
        Dim btnNo As New Button With {.Text = UiText.T("btn.no")}
        Dim btnCancel As New Button With {.Text = UiText.T("btn.cancel")}
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
        MyBase.New(UiText.T("dlg.runmodel.title"))
        ' SelectedIndex decides the model, so these are text and nothing else.
        _cmb.Items.Add(UiText.T("word.model1.levelii"))
        _cmb.Items.Add(UiText.T("word.model3"))
        _cmb.Items.Add(UiText.T("word.model4"))
        _cmb.SelectedIndex = 1
        Finish(Ui.Down(Nothing, Ui.Lbl(UiText.T("dlg.runmodel.machine")), _cmb), UiText.T("btn.run"))
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
        Title = UiText.T("dlg.engine.title")
        Resizable = True
        ClientSize = New Size(820, 560)
        Padding = New Padding(Ui.Margin)
        _txt = Ui.ReportBox(BuildReport(), 9.0F)
        Dim btnClose As Button = Ui.Btn(UiText.T("btn.close"), Sub() Close())
        Dim buttons As TableLayout = Ui.Across(Ui.Btn(UiText.T("btn.recheckengine"), Sub()
                                                                            ReprobeRequested = True
                                                                            Close()
                                                                        End Sub),
                                               Ui.Btn(UiText.T("btn.copyclipboard"), Sub() Ui.CopyTextToClipboard(_txt.Text)), Nothing, btnClose)
        Content = Ui.Down(_txt, _txt, buttons)
        AbortButton = btnClose
        DefaultButton = btnClose
    End Sub

    Private Function BuildReport() As String
        Dim sb As New StringBuilder()
        Dim caps As EngineCapabilities = _tool.Capabilities
        sb.AppendLine(Program.AppName & " " & Program.AppVersion)
        ' The labels are padded to a shared width so the report's columns still line up
        ' when the words are longer in another language. In English PadRight gives back
        ' exactly the spacing this report always had.
        sb.AppendLine(UiText.T("dlg.engine.enginepath").PadRight(20) & If(String.IsNullOrEmpty(_tool.ToolPath), UiText.T("dlg.engine.notfound"), _tool.ToolPath))
        sb.AppendLine(UiText.T("dlg.engine.minversion").PadRight(20) & Trs80Tool.MinimumVersion.ToString())
        sb.AppendLine(UiText.T("dlg.engine.lasttested").PadRight(20) & Trs80Tool.MaxTestedVersion.ToString())
        If caps Is Nothing Then
            sb.AppendLine(UiText.T("dlg.engine.notprobed"))
            Return sb.ToString()
        End If
        sb.AppendLine(UiText.T("dlg.engine.reportedversion").PadRight(20) & caps.VersionText)
        sb.AppendLine(UiText.T("dlg.engine.parsedversion").PadRight(20) & If(caps.Version Is Nothing, UiText.T("dlg.engine.unparseable"), caps.Version.ToString()))
        sb.AppendLine()
        sb.AppendLine(UiText.T("dlg.engine.howcalled").PadRight(20) & _tool.TransportText)
        sb.AppendLine()
        sb.AppendLine(UiText.T("dlg.engine.toplevelused").PadRight(38) & String.Join(", ", Trs80Tool.ExpectedTopLevel))
        sb.AppendLine("  " & UiText.T("dlg.engine.toolhas", If(caps.TopLevelCommands.Count = 0, UiText.T("dlg.engine.notreported"), String.Join(", ", caps.TopLevelCommands))))
        sb.AppendLine(UiText.T("dlg.engine.diskused").PadRight(38) & String.Join(", ", Trs80Tool.ExpectedDisk))
        sb.AppendLine("  " & UiText.T("dlg.engine.toolhas", If(caps.DiskCommands.Count = 0, UiText.T("dlg.engine.notreported"), String.Join(", ", caps.DiskCommands))))
        sb.AppendLine()
        sb.AppendLine(UiText.T("dlg.engine.formatdoses").PadRight(33) & DosList(caps.FormatDoses))
        sb.AppendLine(UiText.T("dlg.engine.writabledoses").PadRight(33) & DosList(caps.WritableDoses))
        sb.AppendLine(UiText.T("dlg.engine.formats", String.Join(", ", caps.ReadFormats), String.Join(", ", caps.ConvertFormats), String.Join(", ", caps.DiskWriteFormats), String.Join(", ", caps.DiskCreateFormats)))
        sb.AppendLine()
        If caps.Problems.Count = 0 Then
            sb.AppendLine(UiText.T("dlg.engine.noproblems"))
        Else
            sb.AppendLine(UiText.T("dlg.engine.problems"))
            For Each p In caps.Problems
                sb.AppendLine("  - " & p)
            Next
        End If
        sb.AppendLine()
        ' The two widths differ by one, as they always have; PadRight keeps that exactly.
        sb.AppendLine(UiText.T("dlg.engine.changelog").PadRight(29) & "https://www.my-trs-80.com/tool/#change-log")
        sb.AppendLine(UiText.T("dlg.engine.diskref").PadRight(30) & "https://www.my-trs-80.com/tool/disk.html")
        sb.AppendLine()
        sb.AppendLine("---- " & UiText.T("dlg.engine.recentcalls") & " ----")
        SyncLock _tool.CallLog
            Dim start As Integer = Math.Max(0, _tool.CallLog.Count - 60)
            For i As Integer = start To _tool.CallLog.Count - 1
                sb.AppendLine(_tool.CallLog(i))
            Next
        End SyncLock
        sb.AppendLine()
        sb.AppendLine("---- " & UiText.T("dlg.engine.cmdoptions") & " ----")
        For Each cmdName In caps.TopLevelCommands
            Dim opts As List(Of String) = Nothing
            If caps.TopLevelOptions.TryGetValue(cmdName, opts) Then sb.AppendLine(cmdName & ": " & String.Join(" ", opts))
        Next
        sb.AppendLine(UiText.T("dlg.engine.diskglobal").PadRight(18) & String.Join(" ", caps.DiskGlobalOptions))
        For Each cmdName In caps.DiskCommands
            Dim opts As List(Of String) = Nothing
            If caps.DiskOptions.TryGetValue(cmdName, opts) Then sb.AppendLine("disk " & cmdName & ": " & String.Join(" ", opts))
        Next
        Return sb.ToString()
    End Function

    Private Shared Function DosList(doses As List(Of Trs80Tool.HelpDosJson)) As String
        If doses Is Nothing OrElse doses.Count = 0 Then Return UiText.T("dlg.engine.notreported")
        Dim names As New List(Of String)
        For Each d In doses
            names.Add(If(String.IsNullOrEmpty(d.Label), d.Name, d.Label))
        Next
        Return String.Join(", ", names)
    End Function
End Class
