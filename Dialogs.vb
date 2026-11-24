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
    Private ReadOnly _rbSelected As New RadioButton() With {.Text = UiText.T("dlg.extract.selected"), .AutoSize = True}
    Private ReadOnly _rbAll As New RadioButton() With {.Text = UiText.T("dlg.extract.all"), .AutoSize = True}
    Private ReadOnly _chkLower As New CheckBox() With {.Text = UiText.T("dlg.extract.lower"), .AutoSize = True}
    Private ReadOnly _chkDamaged As New CheckBox() With {.Text = UiText.T("dlg.extract.damaged"), .AutoSize = True}
    Private ReadOnly _chkCpm As New CheckBox() With {.Text = UiText.T("dlg.extract.cpm"), .AutoSize = True}
    Private ReadOnly _chkUcsd As New CheckBox() With {.Text = UiText.T("dlg.extract.ucsd"), .AutoSize = True}

    Public Sub New(selectedCount As Integer)
        Icon = Program.WindowIcon
        Text = UiText.T("dlg.extract.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(520, 280)

        Controls.Add(New Label() With {.Text = UiText.T("dlg.extract.dir"), .AutoSize = True, .Location = New Point(12, 12)})
        _txtDir.Bounds = New Rectangle(12, 32, 450, 23)
        _txtDir.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        _txtDir.Text = If(String.IsNullOrEmpty(Program.Settings.ExtractDirectory), Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Program.Settings.ExtractDirectory)
        Dim btnBrowse As New Button() With {.Text = "...", .Bounds = New Rectangle(468, 31, 40, 25), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnBrowse.Click, Sub() Browse()
        Controls.Add(_txtDir) : Controls.Add(btnBrowse)

        Dim grp As New GroupBox() With {.Text = UiText.T("dlg.extract.files"), .Bounds = New Rectangle(12, 66, 250, 76)}
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

        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(338, 246, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(428, 246, 80, 26)}
        AddHandler btnOk.Click, Sub(s, e)
                                    If _txtDir.Text.Trim().Length = 0 Then
                                        MessageBox.Show(Me, UiText.T("dlg.extract.nodir"), UiText.T("dlg.extract.title"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                        DialogResult = DialogResult.None
                                    End If
                                End Sub
        Controls.Add(btnOk) : Controls.Add(btnCancel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves. The four
        ' option lines are whole sentences, and they are what decides this window's width.
        UiFit.GrowToFit(grp)
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 12, 10, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    Private Sub Browse()
        Using fbd As New FolderBrowserDialog()
            fbd.Description = UiText.T("dlg.extract.pickfolder")
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
    Private ReadOnly _rbFromNone As New RadioButton() With {.Text = UiText.T("dlg.prefs.noconv"), .AutoSize = True}
    Private ReadOnly _rbFromCrLf As New RadioButton() With {.Text = UiText.T("dlg.prefs.crtocrlf"), .AutoSize = True}
    Private ReadOnly _rbFromLf As New RadioButton() With {.Text = UiText.T("dlg.prefs.crtolf"), .AutoSize = True}
    Private ReadOnly _rbToNone As New RadioButton() With {.Text = UiText.T("dlg.prefs.noconv"), .AutoSize = True}
    Private ReadOnly _rbToCr As New RadioButton() With {.Text = UiText.T("dlg.prefs.crlftocr"), .AutoSize = True}
    Private ReadOnly _chkTokenize As New CheckBox() With {.Text = UiText.T("dlg.prefs.tokenize"), .AutoSize = True}
    Private ReadOnly _chkHostDate As New CheckBox() With {.Text = UiText.T("dlg.prefs.hostdate"), .AutoSize = True}
    Private ReadOnly _chkUnzipGz As New CheckBox() With {.Text = UiText.T("dlg.prefs.unzipgz"), .AutoSize = True}
    Private ReadOnly _chkEsc As New CheckBox() With {.Text = UiText.T("dlg.prefs.esc"), .AutoSize = True}
    Private ReadOnly _chkConfirmDelete As New CheckBox() With {.Text = UiText.T("dlg.prefs.confirmdelete"), .AutoSize = True}
    Private ReadOnly _chkServer As New CheckBox() With {.Text = UiText.T("dlg.prefs.server"), .AutoSize = True}
    Private ReadOnly _cmbLanguage As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 200}
    Private ReadOnly _txtLog As New TextBox() With {.Width = 330}
    Private ReadOnly _txtDisasm As New TextBox() With {.Width = 446}

    Public Sub New()
        Icon = Program.WindowIcon
        Text = UiText.T("dlg.prefs.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 582)
        Dim s As Settings = Program.Settings

        Dim g0 As New GroupBox() With {.Text = UiText.T("dlg.prefs.engine"), .Bounds = New Rectangle(12, 10, 536, 86)}
        g0.Controls.Add(New Label() With {.Text = UiText.T("dlg.prefs.toollocation", Trs80Tool.DefaultExeName), .AutoSize = True, .Location = New Point(12, 20)})
        _txtTool.Bounds = New Rectangle(12, 40, 400, 23)
        _txtTool.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        _txtTool.Text = s.ToolPath
        Dim btnBrowse As New Button() With {.Text = UiText.T("btn.browse"), .Bounds = New Rectangle(420, 39, 80, 25)}
        AddHandler btnBrowse.Click, Sub() BrowseTool()
        _lblToolStatus.Location = New Point(12, 66)
        _lblToolStatus.Text = If(String.IsNullOrEmpty(Program.Tool.ToolPath), UiText.T("dlg.prefs.toolnotfound"), If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, UiText.T("dlg.prefs.toolversion", UiText.T("dlg.prefs.toolusing", Program.Tool.ToolPath), Program.Tool.Capabilities.Version.ToString()), UiText.T("dlg.prefs.toolusing", Program.Tool.ToolPath)))
        g0.Controls.AddRange(New Control() {_txtTool, btnBrowse, _lblToolStatus})

        Dim g1 As New GroupBox() With {.Text = UiText.T("dlg.prefs.fromtrs"), .Bounds = New Rectangle(12, 104, 262, 100)}
        _rbFromNone.Location = New Point(14, 22) : _rbFromCrLf.Location = New Point(14, 46) : _rbFromLf.Location = New Point(14, 70)
        g1.Controls.AddRange(New Control() {_rbFromNone, _rbFromCrLf, _rbFromLf})
        Select Case s.ConversionFromTrs80
            Case TextConversionOut.CrToCrLf : _rbFromCrLf.Checked = True
            Case TextConversionOut.CrToLf : _rbFromLf.Checked = True
            Case Else : _rbFromNone.Checked = True
        End Select

        Dim g2 As New GroupBox() With {.Text = UiText.T("dlg.prefs.totrs"), .Bounds = New Rectangle(286, 104, 262, 100)}
        _rbToNone.Location = New Point(14, 22) : _rbToCr.Location = New Point(14, 46)
        g2.Controls.AddRange(New Control() {_rbToNone, _rbToCr})
        If s.ConversionToTrs80 = TextConversionIn.CrLfToCr Then _rbToCr.Checked = True Else _rbToNone.Checked = True
        g2.Controls.Add(New Label() With {.Text = UiText.T("dlg.prefs.textonly"), .AutoSize = True, .Location = New Point(14, 72), .ForeColor = SystemColors.GrayText})

        _chkTokenize.Location = New Point(12, 216) : _chkTokenize.Checked = s.TokenizeBasicOnAdd
        _chkHostDate.Location = New Point(12, 242) : _chkHostDate.Checked = s.UseHostDateOnAdd
        _chkUnzipGz.Location = New Point(12, 268) : _chkUnzipGz.Checked = s.UnzipGzOnAdd
        _chkEsc.Location = New Point(12, 294) : _chkEsc.Checked = s.ExitOnEscape
        _chkConfirmDelete.Location = New Point(12, 320) : _chkConfirmDelete.Checked = s.ConfirmDelete
        _chkServer.Location = New Point(12, 346) : _chkServer.Checked = s.KeepEngineRunning

        ' The language of the program's own text. Only the languages whose table is in this build are
        ' offered, and a change takes effect at the next start: text already on screen is not revisited.
        Dim lblLang As New Label() With {.Text = UiText.T("dlg.prefs.language"), .AutoSize = True, .Location = New Point(12, 375)}
        _cmbLanguage.Location = New Point(12, 395)
        For Each lang As UiText.Language In UiText.Languages
            _cmbLanguage.Items.Add(lang.NativeName)
        Next
        _cmbLanguage.SelectedIndex = LanguageIndex()
        Dim lblLangNote As New Label() With {.Text = UiText.T("dlg.prefs.langnote"), .AutoSize = True, .Location = New Point(222, 398), .ForeColor = SystemColors.GrayText}
        Dim lblLog As New Label() With {.Text = UiText.T("dlg.prefs.logfile"), .AutoSize = True, .Location = New Point(12, 425)}
        _txtLog.Location = New Point(12, 445)
        _txtLog.Text = s.EngineLogFile
        Dim btnLog As New Button() With {.Text = UiText.T("btn.browse"), .Bounds = New Rectangle(350, 444, 80, 25)}
        AddHandler btnLog.Click, Sub()
                                     Using sfd As New SaveFileDialog()
                                         sfd.Title = UiText.T("dlg.prefs.picklog")
                                         sfd.Filter = UiText.T("filter.log")
                                         sfd.OverwritePrompt = False
                                         If sfd.ShowDialog(Me) = DialogResult.OK Then _txtLog.Text = sfd.FileName
                                     End Using
                                 End Sub
        Dim lblDisasm As New Label() With {.Text = UiText.T("dlg.prefs.disasm"), .AutoSize = True, .Location = New Point(12, 475)}
        _txtDisasm.Location = New Point(12, 495)
        _txtDisasm.Text = s.DisasmOptions
        Dim btnDisasmDefault As New Button() With {.Text = UiText.T("btn.default"), .Bounds = New Rectangle(466, 494, 80, 25)}
        AddHandler btnDisasmDefault.Click, Sub() _txtDisasm.Text = Settings.DefaultDisasmOptions
        Dim lblSettings As New Label() With {.Text = UiText.T("dlg.prefs.settingsfile", Settings.FilePath), .AutoSize = True, .Location = New Point(12, 524), .ForeColor = SystemColors.GrayText}

        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(378, 544, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(468, 544, 80, 26)}
        AddHandler btnOk.Click, Sub() Apply()
        Controls.AddRange(New Control() {g0, g1, g2, _chkTokenize, _chkHostDate, _chkUnzipGz, _chkEsc, _chkConfirmDelete, _chkServer, lblLang, _cmbLanguage, lblLangNote, lblLog, _txtLog, btnLog, lblDisasm, _txtDisasm, btnDisasmDefault, lblSettings, btnOk, btnCancel})

        ' Make room for the language in force (UiFit.vb); in English nothing here moves. The two
        ' conversion boxes sit side by side and their captions are long sentences, so each is
        ' widened for its own caption and the right-hand one is pushed clear of the left.
        UiFit.SpreadRow(12, _cmbLanguage, lblLangNote)
        UiFit.SpreadRow(10, _txtLog, btnLog)
        UiFit.SpreadRow(10, _txtDisasm, btnDisasmDefault)
        UiFit.GrowToFit(g0)
        UiFit.FitCaption(g1)
        UiFit.FitCaption(g2)
        UiFit.GrowToFit(g1)
        UiFit.GrowToFit(g2)
        UiFit.SpreadRow(12, g1, g2)
        UiFit.GrowForm(Me)
        ' The three boxes that run the width of the window follow it.
        g0.Width = Math.Max(g0.Width, ClientSize.Width - 24)
        _txtLog.Width = Math.Max(_txtLog.Width, btnLog.Left - _txtLog.Left - 10)
        _txtDisasm.Width = Math.Max(_txtDisasm.Width, btnDisasmDefault.Left - _txtDisasm.Left - 10)
        UiFit.GrowToFit(g0)
        UiFit.ButtonRow(ClientSize.Width - 12, 10, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel
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
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("dlg.prefs.picktool")
            ofd.Filter = UiText.T("filter.programs")
            ofd.FileName = If(_txtTool.Text.Length > 0, _txtTool.Text, Trs80Tool.DefaultExeName)
            If ofd.ShowDialog(Me) = DialogResult.OK Then _txtTool.Text = ofd.FileName
        End Using
    End Sub

    Private Sub Apply()
        Dim s As Settings = Program.Settings
        Dim p As String = _txtTool.Text.Trim()
        If p.Length > 0 AndAlso Not File.Exists(p) Then
            If MessageBox.Show(Me, UiText.TCrLf("dlg.prefs.notexist", p), UiText.T("dlg.prefs.title"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
                DialogResult = DialogResult.None
                Return
            End If
        End If
        s.ToolPath = p
        s.ConversionFromTrs80 = If(_rbFromCrLf.Checked, TextConversionOut.CrToCrLf, If(_rbFromLf.Checked, TextConversionOut.CrToLf, TextConversionOut.None))
        s.ConversionToTrs80 = If(_rbToCr.Checked, TextConversionIn.CrLfToCr, TextConversionIn.None)
        s.TokenizeBasicOnAdd = _chkTokenize.Checked
        s.UseHostDateOnAdd = _chkHostDate.Checked
        s.UnzipGzOnAdd = _chkUnzipGz.Checked
        s.ExitOnEscape = _chkEsc.Checked
        s.ConfirmDelete = _chkConfirmDelete.Checked
        s.EngineLogFile = _txtLog.Text.Trim()
        s.KeepEngineRunning = _chkServer.Checked
        s.DisasmOptions = _txtDisasm.Text.Trim()
        ' Saved, but not put in force here: the language is read at the next start.
        Dim li As Integer = _cmbLanguage.SelectedIndex
        If li >= 0 AndAlso li < UiText.Languages.Count Then s.Language = UiText.Languages(li).Code
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
        Icon = Program.WindowIcon
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
        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(198, y, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(288, y, 80, 26)}
        AddHandler btnOk.Click, Sub()
                                    If Not TextUtil.IsValidTrsName(_txt.Text) Then
                                        MessageBox.Show(Me, UiText.TCrLf("dlg.rename.invalid", _txt.Text), title, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                        DialogResult = DialogResult.None
                                    End If
                                End Sub
        Controls.Add(btnOk) : Controls.Add(btnCancel)
        AcceptButton = btnOk : CancelButton = btnCancel

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        ClientSize = New Size(380, y + 38)
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 12, 10, btnOk, btnCancel)
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

    Private ReadOnly _chkAll As New CheckBox() With {.Text = UiText.T("dlg.replace.all"), .AutoSize = True}

    Public Sub New(existing As DiskFile, newName As String, newSize As Long, newDate As DateTime)
        Icon = Program.WindowIcon
        Text = UiText.T("dlg.replace.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(400, 230)
        Controls.Add(New Label() With {.Text = UiText.T("dlg.replace.contains", existing.SlashName), .AutoSize = True, .Location = New Point(12, 12)})
        Controls.Add(New Label() With {.Text = UiText.T("dlg.replace.would"), .AutoSize = True, .Location = New Point(12, 40)})
        Controls.Add(New Label() With {.Text = existing.SlashName & vbCrLf & UiText.T("word.bytes", TextUtil.FormatBytes(existing.Size)) & vbCrLf & If(existing.DateValue.HasValue, existing.DateValue.Value.ToString("g", CultureInfo.CurrentCulture), If(existing.DateText.Length > 0, existing.DateText, UiText.T("word.nodateparen"))), .AutoSize = True, .Location = New Point(30, 60)})
        Controls.Add(New Label() With {.Text = UiText.T("dlg.replace.withthis"), .AutoSize = True, .Location = New Point(12, 112)})
        Controls.Add(New Label() With {.Text = newName & vbCrLf & UiText.T("word.bytes", TextUtil.FormatBytes(newSize)) & vbCrLf & newDate.ToString("g", CultureInfo.CurrentCulture), .AutoSize = True, .Location = New Point(30, 132)})
        _chkAll.Location = New Point(12, 186)
        Controls.Add(_chkAll)
        Dim btnYes As New Button() With {.Text = UiText.T("btn.yes"), .DialogResult = DialogResult.Yes, .Bounds = New Rectangle(218, 196, 80, 26)}
        Dim btnNo As New Button() With {.Text = UiText.T("btn.no"), .DialogResult = DialogResult.No, .Bounds = New Rectangle(308, 196, 80, 26)}
        Controls.Add(btnYes) : Controls.Add(btnNo)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 12, 10, btnYes, btnNo)

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
        Icon = Program.WindowIcon
        _tool = tool
        Text = UiText.T("dlg.engine.title")
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        Size = New Size(760, 560)
        MinimumSize = New Size(500, 300)
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
            .Dock = DockStyle.Fill
        End With
        Dim panel As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        Dim btnReprobe As New Button() With {.Text = UiText.T("btn.recheckengine"), .Bounds = New Rectangle(8, 7, 130, 26)}
        AddHandler btnReprobe.Click, Sub()
                                         ReprobeRequested = True
                                         Close()
                                     End Sub
        Dim btnCopy As New Button() With {.Text = UiText.T("btn.copyclipboard"), .Bounds = New Rectangle(146, 7, 130, 26)}
        AddHandler btnCopy.Click, Sub()
                                      Try : Clipboard.SetText(_txt.Text) : Catch : End Try
                                  End Sub
        Dim btnClose As New Button() With {.Text = UiText.T("btn.close"), .DialogResult = DialogResult.OK, .Anchor = AnchorStyles.Right Or AnchorStyles.Top, .Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)}
        panel.Controls.AddRange(New Control() {btnReprobe, btnCopy, btnClose})
        Controls.Add(_txt)
        Controls.Add(panel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.ButtonRowLeft(8, 8, btnReprobe, btnCopy)
        UiFit.FitButton(btnClose)
        btnClose.Left = panel.ClientSize.Width - btnClose.Width - 16
        CancelButton = btnClose
        _txt.Text = BuildReport()
        _txt.Select(0, 0)
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
        Return sb.ToString().Replace(vbCr, "").Replace(vbLf, vbCrLf)
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
