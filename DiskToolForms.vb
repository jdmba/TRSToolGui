' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskToolForms.vb — windows for the trs80-tool features that TRSTool never
'  had: a generic text-report window (granule map, sectors
'  text, hexdump, …), Check & repair, Trim, and Boot-in-emulator. Each one
'  only builds option words and shows what trs80-tool printed; MainForm runs
'  the commands and refreshes the image afterwards.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms

' ---------------------------------------------------------------- Text report

''' <summary>
''' Shows a block of trs80-tool text in a monospaced box (Consolas renders the map's box characters),
''' with Copy / Save / Close. Optionally a filter box whose text is handed to a refresher that returns new text
''' (used by the granule map's PATTERN argument).
''' </summary>
Public Class TextReportForm
    Inherits Form

    Private ReadOnly _txt As New TextBox()
    Private ReadOnly _filter As New TextBox()
    Private ReadOnly _btnRun As New Button()
    Private ReadOnly _refresher As Func(Of String, Task(Of String))
    Private ReadOnly _baseTitle As String

    Public Sub New(title As String, text As String)
        Me.New(title, text, Nothing, Nothing)
    End Sub

    ''' <param name="filterLabel">Label for the filter box, or Nothing for no filter row.</param>
    ''' <param name="refresher">Called with the filter text; returns the new report text.</param>
    Public Sub New(title As String, text As String, filterLabel As String, refresher As Func(Of String, Task(Of String)))
        Icon = Program.WindowIcon
        _baseTitle = title
        _refresher = refresher
        Me.Text = title
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(900, 620)
        MinimumSize = New Size(480, 300)
        ShowInTaskbar = True
        KeyPreview = True
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.5F)
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With
        Dim panel As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        Dim btnCopy As New Button() With {.Text = UiText.T("btn.copyclipboard"), .Bounds = New Rectangle(8, 7, 130, 26)}
        AddHandler btnCopy.Click, Sub()
                                      Try : Clipboard.SetText(If(_txt.Text.Length = 0, " ", _txt.Text)) : Catch : End Try
                                  End Sub
        Dim btnSave As New Button() With {.Text = UiText.T("btn.saveas"), .Bounds = New Rectangle(146, 7, 100, 26)}
        AddHandler btnSave.Click, Sub() SaveText()
        Dim btnClose As New Button() With {.Text = UiText.T("btn.close"), .DialogResult = DialogResult.OK, .Anchor = AnchorStyles.Top Or AnchorStyles.Right, .Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)}
        AddHandler btnClose.Click, Sub() Close()
        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.ButtonRowLeft(8, 8, btnCopy, btnSave)
        btnCopy.Top = 7 : btnSave.Top = 7
        UiFit.FitButton(btnClose)
        btnClose.Left = ClientSize.Width - btnClose.Width - 16
        panel.Controls.AddRange(New Control() {btnCopy, btnSave, btnClose})
        If filterLabel IsNot Nothing AndAlso refresher IsNot Nothing Then
            Dim top As New Panel() With {.Dock = DockStyle.Top, .Height = 34}
            Dim lbl As New Label() With {.Text = filterLabel, .AutoSize = True, .Location = New Point(8, 9)}
            _filter.Bounds = New Rectangle(lbl.PreferredWidth + 16, 6, 300, 23)
            _btnRun.Text = UiText.T("btn.run")
            _btnRun.Bounds = New Rectangle(_filter.Right + 8, 5, 70, 25)
            AddHandler _btnRun.Click, Async Sub() Await RefreshAsync()
            AddHandler _filter.KeyDown, Async Sub(s, e)
                                            If e.KeyCode = Keys.Enter Then
                                                e.Handled = True : e.SuppressKeyPress = True
                                                Await RefreshAsync()
                                            End If
                                        End Sub
            UiFit.AfterAll(lbl, 8, _filter, _btnRun)
            UiFit.FitButton(_btnRun)
            top.Controls.AddRange(New Control() {lbl, _filter, _btnRun})
            Controls.Add(_txt)
            Controls.Add(top)
        Else
            Controls.Add(_txt)
        End If
        Controls.Add(panel)
        CancelButton = btnClose
        SetText(text)
    End Sub

    Public Sub SetText(text As String)
        _txt.Text = If(text, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
        _txt.Select(0, 0)
    End Sub

    Private Async Function RefreshAsync() As Task
        If _refresher Is Nothing Then Return
        UseWaitCursor = True
        _btnRun.Enabled = False
        Try
            SetText(Await _refresher(_filter.Text.Trim()))
            Me.Text = _baseTitle & If(_filter.Text.Trim().Length > 0, " [" & _filter.Text.Trim() & "]", "")
        Catch ex As Exception
            SetText("(" & ex.Message & ")")
        Finally
            UseWaitCursor = False
            _btnRun.Enabled = True
        End Try
    End Function

    Private Sub SaveText()
        Using sfd As New SaveFileDialog()
            sfd.Title = UiText.T("caption.savereport")
            sfd.Filter = UiText.T("filter.textfile")
            sfd.FileName = Regex.Replace(_baseTitle, "[^A-Za-z0-9_.-]+", "_") & ".txt"
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                File.WriteAllText(sfd.FileName, _txt.Text, Encoding.UTF8)
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, UiText.T("caption.savereport"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
End Class

' ---------------------------------------------------------------- Check & repair

''' <summary>
''' "disk check" on open, then "disk repair [--free-unclaimed] [--fix-program-sizes]" with a dry-run preview.
''' Changed is True once a real repair ran and saved, so the caller refreshes the image.
''' </summary>
Public Class CheckRepairForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txt As New TextBox()
    Private ReadOnly _chkFree As New CheckBox() With {.Text = UiText.T("dlg.checkrepair.free"), .AutoSize = True}
    Private ReadOnly _chkSizes As New CheckBox() With {.Text = UiText.T("dlg.checkrepair.sizes"), .AutoSize = True}
    Private ReadOnly _chkMissing As New CheckBox() With {.Text = UiText.T("dlg.checkrepair.missing"), .AutoSize = True}
    Private ReadOnly _btnCheck As New Button()
    Private ReadOnly _btnPreview As New Button()
    Private ReadOnly _btnRepair As New Button()
    Private ReadOnly _btnClose As New Button()
    Private _changed As Boolean = False

    Public Sub New(disk As DiskImage)
        Icon = Program.WindowIcon
        _disk = disk
        Dim shown As String = UiText.T("dlg.checkrepair.title", Path.GetFileName(disk.Path))
        If disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing Then shown = UiText.T("title.partition", shown, disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Text = shown
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(860, 600)
        MinimumSize = New Size(560, 360)
        ShowInTaskbar = False
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.5F)
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With
        Dim opts As New Panel() With {.Dock = DockStyle.Bottom, .Height = 84}
        _chkFree.Location = New Point(10, 8)
        _chkSizes.Location = New Point(10, 32)
        _chkMissing.Location = New Point(10, 56)
        _chkMissing.Enabled = Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.HasDiskOption("repair", "--mark-missing-used")
        opts.Controls.AddRange(New Control() {_chkFree, _chkSizes, _chkMissing})
        Dim panel As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        _btnCheck.Text = UiText.T("btn.recheck") : _btnCheck.Bounds = New Rectangle(8, 7, 90, 26)
        _btnPreview.Text = UiText.T("btn.previewrepair") : _btnPreview.Bounds = New Rectangle(104, 7, 170, 26)
        _btnRepair.Text = UiText.T("btn.repair") : _btnRepair.Bounds = New Rectangle(280, 7, 90, 26)
        _btnClose.Text = UiText.T("btn.close") : _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right : _btnClose.Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)
        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        AddHandler _btnPreview.Click, Async Sub() Await RunRepairAsync(True)
        AddHandler _btnRepair.Click, Async Sub() Await RunRepairAsync(False)
        AddHandler _btnClose.Click, Sub() Close()
        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.ButtonRowLeft(8, 6, _btnCheck, _btnPreview, _btnRepair)
        _btnCheck.Top = 7 : _btnPreview.Top = 7 : _btnRepair.Top = 7
        UiFit.FitButton(_btnClose)
        _btnClose.Left = ClientSize.Width - _btnClose.Width - 16
        panel.Controls.AddRange(New Control() {_btnCheck, _btnPreview, _btnRepair, _btnClose})
        Controls.Add(_txt)
        Controls.Add(opts)
        Controls.Add(panel)
        CancelButton = _btnClose
        Dim canRepair As Boolean = disk.WriteProtected.HasValue AndAlso Not disk.WriteProtected.Value OrElse Not disk.WriteProtected.HasValue
        _btnPreview.Enabled = canRepair
        _btnRepair.Enabled = canRepair
        If Not canRepair Then _txt.Text = UiText.T("dlg.checkrepair.wp") & vbCrLf & vbCrLf
    End Sub

    Public ReadOnly Property Changed As Boolean
        Get
            Return _changed
        End Get
    End Property

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Await RunCheckAsync()
    End Sub

    Private Sub SetBusy(busy As Boolean)
        UseWaitCursor = busy
        _btnCheck.Enabled = Not busy
        _btnPreview.Enabled = Not busy
        _btnRepair.Enabled = Not busy
    End Sub

    Private Sub Append(header As String, body As String)
        Dim sb As New StringBuilder(_txt.Text)
        If sb.Length > 0 Then sb.AppendLine()
        sb.AppendLine("==== " & header & " ====")
        sb.AppendLine(If(body, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf).TrimEnd())
        _txt.Text = sb.ToString()
        _txt.Select(_txt.Text.Length, 0)
        _txt.ScrollToCaret()
    End Sub

    Private Async Function RunCheckAsync() As Task
        SetBusy(True)
        Try
            Dim r As ToolResult = Await Program.Tool.DiskCheckAsync(_disk.Path, _disk.PartitionArg)
            Append("disk check", If(r.Succeeded, r.StdOut, r.ErrorText))
        Catch ex As Exception
            Append("disk check", ex.Message)
        Finally
            SetBusy(False)
        End Try
    End Function

    Private Async Function RunRepairAsync(dryRun As Boolean) As Task
        If Not dryRun Then
            If MessageBox.Show(Me, UiText.TCrLf("dlg.checkrepair.confirm"), UiText.T("caption.repair"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
        End If
        SetBusy(True)
        Try
            Dim missing As Boolean = _chkMissing.Enabled AndAlso _chkMissing.Checked
            Dim r As ToolResult = Await Program.Tool.DiskRepairAsync(_disk.Path, _disk.PartitionArg, _chkFree.Checked, _chkSizes.Checked, missing, dryRun)
            Append(If(dryRun, "disk --dry-run repair", "disk repair") & If(_chkFree.Checked, " --free-unclaimed", "") & If(_chkSizes.Checked, " --fix-program-sizes", "") & If(missing, " --mark-missing-used", ""), If(r.Succeeded, r.StdOut, r.ErrorText))
            If Not dryRun AndAlso r.Succeeded AndAlso r.Disk IsNot Nothing AndAlso r.Disk.Saved IsNot Nothing Then _changed = True
        Catch ex As Exception
            Append("disk repair", ex.Message)
        Finally
            SetBusy(False)
        End Try
    End Function
End Class

' ---------------------------------------------------------------- Trim

''' <summary>Options for "disk trim": drop unused cylinders (default), keep cylinders 0..N-1, or drop side 1; --force.</summary>
Public Class TrimForm
    Inherits Form

    Private ReadOnly _rbDefault As New RadioButton() With {.Text = UiText.T("dlg.trim.default"), .AutoSize = True, .Checked = True}
    Private ReadOnly _rbCylinders As New RadioButton() With {.Text = UiText.T("dlg.trim.cylinders"), .AutoSize = True}
    Private ReadOnly _rbSide As New RadioButton() With {.Text = UiText.T("dlg.trim.side"), .AutoSize = True}
    Private ReadOnly _numCyl As New NumericUpDown() With {.Minimum = 1, .Maximum = 255, .Value = 35, .Width = 60}
    Private ReadOnly _chkForce As New CheckBox() With {.Text = UiText.T("dlg.trim.force") & " " & UiText.T("dlg.trim.forcenote"), .AutoSize = True}

    Public Sub New(disk As DiskImage)
        Icon = Program.WindowIcon
        Text = UiText.T("dlg.trim.title", Path.GetFileName(disk.Path))
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(640, 200)
        If disk.Geometry IsNot Nothing AndAlso disk.Geometry.CylinderCount > 0 Then _numCyl.Value = Math.Min(255, Math.Max(1, disk.Geometry.CylinderCount))
        Controls.Add(New Label() With {.Text = UiText.T("dlg.trim.intro"), .AutoSize = True, .Location = New Point(12, 12)})
        _rbDefault.Location = New Point(14, 40)
        _rbCylinders.Location = New Point(14, 66)
        _numCyl.Location = New Point(_rbCylinders.Left + _rbCylinders.PreferredSize.Width + 8, 64)
        _rbSide.Location = New Point(14, 92)
        _chkForce.Location = New Point(14, 124)
        Controls.AddRange(New Control() {_rbDefault, _rbCylinders, _numCyl, _rbSide, _chkForce})
        AddHandler _rbCylinders.CheckedChanged, Sub() _numCyl.Enabled = _rbCylinders.Checked
        _numCyl.Enabled = False
        Dim btnOk As New Button() With {.Text = UiText.T("btn.trim"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(460, 162, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(550, 162, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves. The
        ' force line is a whole sentence and is the widest thing on this window.
        UiFit.SpreadRow(8, _rbCylinders, _numCyl)
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)
        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _rbCylinders.Checked Then
            w.Add("--cylinders") : w.Add(CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture))
        ElseIf _rbSide.Checked Then
            w.Add("--side") : w.Add("0")
        End If
        If _chkForce.Checked Then w.Add("--force")
        Return w
    End Function
End Class

' ---------------------------------------------------------------- Boot in the emulator

''' <summary>
''' Options for "trs80-tool run IMAGE": model, writable, extra images to mount in the next drives, printer file.
''' The emulator runs in its own console window (the GUI just starts it).
''' </summary>
Public Class BootForm
    Inherits Form

    Private ReadOnly _rbM1 As New RadioButton() With {.Text = UiText.T("word.model1"), .AutoSize = True}
    Private ReadOnly _rbM3 As New RadioButton() With {.Text = UiText.T("word.model3"), .AutoSize = True}
    Private ReadOnly _rbM4 As New RadioButton() With {.Text = UiText.T("word.model4"), .AutoSize = True}
    Private ReadOnly _chkWritable As New CheckBox() With {.Text = UiText.T("dlg.boot.writable"), .AutoSize = True}
    Private ReadOnly _lstMount As New ListBox()
    Private ReadOnly _txtPrinter As New TextBox() With {.Width = 360}

    Public Sub New(disk As DiskImage)
        Icon = Program.WindowIcon
        Text = UiText.T("dlg.boot.title", Path.GetFileName(disk.Path))
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 360)
        Controls.Add(New Label() With {.Text = UiText.TCrLf("dlg.boot.intro.windows"), .AutoSize = False, .Bounds = New Rectangle(12, 10, 536, 36)})
        Dim grp As New GroupBox() With {.Text = UiText.T("dlg.boot.model"), .Bounds = New Rectangle(12, 50, 536, 50)}
        _rbM1.Location = New Point(14, 20) : _rbM3.Location = New Point(110, 20) : _rbM4.Location = New Point(210, 20)
        grp.Controls.AddRange(New Control() {_rbM1, _rbM3, _rbM4})
        Controls.Add(grp)
        Select Case GuessModel(disk.DosName)
            Case 1 : _rbM1.Checked = True
            Case 4 : _rbM4.Checked = True
            Case Else : _rbM3.Checked = True
        End Select
        _chkWritable.Location = New Point(14, 110)
        Controls.Add(_chkWritable)
        Controls.Add(New Label() With {.Text = UiText.T("dlg.boot.mount"), .AutoSize = True, .Location = New Point(12, 140)})
        _lstMount.Bounds = New Rectangle(12, 160, 430, 90)
        Controls.Add(_lstMount)
        Dim btnAdd As New Button() With {.Text = UiText.T("btn.addellipsis"), .Bounds = New Rectangle(452, 160, 96, 26)}
        Dim btnRemove As New Button() With {.Text = UiText.T("btn.remove"), .Bounds = New Rectangle(452, 192, 96, 26)}
        AddHandler btnAdd.Click, Sub() AddMount()
        AddHandler btnRemove.Click, Sub()
                                        If _lstMount.SelectedIndex >= 0 Then _lstMount.Items.RemoveAt(_lstMount.SelectedIndex)
                                    End Sub
        Controls.Add(btnAdd) : Controls.Add(btnRemove)
        Controls.Add(New Label() With {.Text = UiText.T("dlg.boot.printer"), .AutoSize = True, .Location = New Point(12, 262)})
        _txtPrinter.Location = New Point(12, 282)
        Controls.Add(_txtPrinter)
        Dim btnPrinter As New Button() With {.Text = UiText.T("btn.browse"), .Bounds = New Rectangle(380, 281, 80, 25)}
        AddHandler btnPrinter.Click, Sub()
                                         Using sfd As New SaveFileDialog()
                                             sfd.Title = UiText.T("dlg.boot.pickprinter")
                                             sfd.Filter = UiText.T("filter.textfile")
                                             sfd.OverwritePrompt = False
                                             If sfd.ShowDialog(Me) = DialogResult.OK Then _txtPrinter.Text = sfd.FileName
                                         End Using
                                     End Sub
        Controls.Add(btnPrinter)
        Dim btnOk As New Button() With {.Text = UiText.T("btn.boot"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(380, 322, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(468, 322, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.SpreadRow(12, _rbM1, _rbM3, _rbM4)
        UiFit.GrowToFit(grp)
        UiFit.GrowForm(Me)
        grp.Width = Math.Max(grp.Width, ClientSize.Width - 24)
        _lstMount.Width = Math.Max(_lstMount.Width, btnAdd.Left - _lstMount.Left - 10)
        UiFit.ButtonRow(ClientSize.Width - 12, 8, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    ''' <summary>1, 3 or 4 from the DOS name: TRSDOS 6 / LS-DOS 6 → 4; TRSDOS 2.x, VTOS, DoubleDOS, NEWDOS/80 v1 → I; else III.</summary>
    Public Shared Function GuessModel(dosName As String) As Integer
        Dim d As String = If(dosName, "")
        If Regex.IsMatch(d, "^(TRSDOS|LS-DOS) 6\.") OrElse d.IndexOf("Model 4", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 4
        If Regex.IsMatch(d, "^TRSDOS 2\.") OrElse Regex.IsMatch(d, "^(VTOS|DoubleDOS|NEWDOS 2|NEWDOS/80 v1)", RegexOptions.IgnoreCase) Then Return 1
        Return 3
    End Function

    Private Sub AddMount()
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("dlg.boot.pickmount")
            ofd.Filter = UiText.T("filter.mountimages")
            ofd.Multiselect = True
            If ofd.ShowDialog(Me) = DialogResult.OK Then
                For Each f In ofd.FileNames
                    If Not _lstMount.Items.Contains(f) Then _lstMount.Items.Add(f)
                Next
            End If
        End Using
    End Sub

    ''' <summary>The words after "run": IMAGE --model N [--writable] [--printer F] [--mount A B ...] (--mount last: it is variadic).</summary>
    Public Function RunWords(imagePath As String) As List(Of String)
        Dim w As New List(Of String) From {"run", imagePath, "--model", If(_rbM1.Checked, "1", If(_rbM4.Checked, "4", "3"))}
        If _chkWritable.Checked Then w.Add("--writable")
        If _txtPrinter.Text.Trim().Length > 0 Then
            w.Add("--printer") : w.Add(_txtPrinter.Text.Trim())
        End If
        If _lstMount.Items.Count > 0 Then
            w.Add("--mount")
            For Each o In _lstMount.Items
                w.Add(o.ToString())
            Next
        End If
        Return w
    End Function
End Class

' ---------------------------------------------------------------- Import from another image

''' <summary>
''' Disk → Import from another image: "disk IMG import SOURCE [PATTERN...]", new in trs80-tool 5.3.0.
''' The engine reads the other image and copies the files, so each one keeps its own date and its system
''' and hidden flags and BASIC is converted between Model III and Model 4 where the two disks differ —
''' none of which the extract-to-a-temp-file-and-add-it path behind Copy, Paste and drag can do.
''' Which files to copy is left to the engine's own DOS wildcard matching (5.3.0 made those rules the same
''' on the host and on a disk), so this window collects a source, patterns and options and nothing more.
''' Preview is the same command with --dry-run: it writes nothing and prints what would happen.
''' </summary>
Public Class ImportForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txt As New TextBox()
    Private ReadOnly _lblSource As New Label() With {.Text = UiText.T("dlg.import.source"), .AutoSize = True}
    Private ReadOnly _txtSource As New TextBox() With {.ReadOnly = True, .Text = UiText.T("dlg.import.nosource"), .BackColor = SystemColors.Control}
    Private ReadOnly _btnBrowse As New Button() With {.Text = UiText.T("btn.browse")}
    Private ReadOnly _rbAll As New RadioButton() With {.Text = UiText.T("dlg.import.whichall"), .AutoSize = True, .Checked = True}
    Private ReadOnly _rbNames As New RadioButton() With {.Text = UiText.T("dlg.import.whichnames"), .AutoSize = True}
    Private ReadOnly _txtNames As New TextBox() With {.Width = 240, .Enabled = False}
    Private ReadOnly _lblHint As New Label() With {.Text = UiText.T("dlg.import.nameshint"), .AutoSize = False, .ForeColor = SystemColors.GrayText}
    Private ReadOnly _chkOverwrite As New CheckBox() With {.Text = UiText.T("dlg.import.overwrite"), .AutoSize = True}
    Private ReadOnly _chkKeepGoing As New CheckBox() With {.Text = UiText.T("dlg.import.keepgoing"), .AutoSize = True, .Checked = True}
    Private ReadOnly _chkBinary As New CheckBox() With {.Text = UiText.T("dlg.import.binary"), .AutoSize = True}
    Private ReadOnly _chkSystem As New CheckBox() With {.Text = UiText.T("dlg.import.system"), .AutoSize = True}
    Private ReadOnly _chkInvisible As New CheckBox() With {.Text = UiText.T("dlg.import.invisible"), .AutoSize = True}
    Private ReadOnly _lblDate As New Label() With {.Text = UiText.T("dlg.import.usedate"), .AutoSize = True}
    Private ReadOnly _txtDate As New TextBox() With {.Width = 110}
    Private ReadOnly _btnPreview As New Button() With {.Text = UiText.T("btn.previewimport")}
    Private ReadOnly _btnImport As New Button() With {.Text = UiText.T("btn.import")}
    Private ReadOnly _btnClose As New Button() With {.Text = UiText.T("btn.close")}
    Private _source As String = Nothing
    Private _changed As Boolean = False
    Private _copied As Integer = 0

    Public Sub New(disk As DiskImage)
        Icon = Program.WindowIcon
        _disk = disk
        Dim shown As String = UiText.T("dlg.import.title", Path.GetFileName(disk.Path))
        If disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing Then shown = UiText.T("title.partition", shown, disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Text = shown
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(880, 640)
        MinimumSize = New Size(620, 460)
        ShowInTaskbar = False
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.5F)
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With

        ' trs80-tool 5.2.0 has no --keep-going; the box says so by being off and disabled.
        If Not Program.Tool.CanKeepGoing Then
            _chkKeepGoing.Checked = False
            _chkKeepGoing.Enabled = False
        End If

        Dim top As New Panel() With {.Dock = DockStyle.Top, .Height = 282}
        Dim lblIntro As New Label() With {.Text = UiText.T("dlg.import.intro"), .AutoSize = True, .Location = New Point(10, 10)}
        Dim lblKeeps As New Label() With {.Text = UiText.T("dlg.import.keeps"), .AutoSize = False, .ForeColor = SystemColors.GrayText,
                                          .Location = New Point(10, 32), .Size = New Size(840, 34)}
        _lblSource.Location = New Point(10, 76)
        _txtSource.Location = New Point(UiFit.After(_lblSource, _txtSource), 73)
        _txtSource.Width = 520
        _btnBrowse.Bounds = New Rectangle(_txtSource.Right + 8, 72, 90, 25)
        _rbAll.Location = New Point(12, 110)
        _rbNames.Location = New Point(12, 136)
        _txtNames.Location = New Point(UiFit.After(_rbNames, _txtNames), 134)
        _lblHint.Location = New Point(30, 160)
        _lblHint.Size = New Size(820, 32)
        _chkOverwrite.Location = New Point(12, 198)
        _chkKeepGoing.Location = New Point(12, 220)
        _chkBinary.Location = New Point(12, 242)
        _lblDate.Location = New Point(430, 198)
        _txtDate.Location = New Point(430, 218)
        _chkSystem.Location = New Point(430, 242)
        _chkInvisible.Location = New Point(430, 264)
        top.Controls.AddRange(New Control() {lblIntro, lblKeeps, _lblSource, _txtSource, _btnBrowse, _rbAll, _rbNames,
                                             _txtNames, _lblHint, _chkOverwrite, _chkKeepGoing, _chkBinary,
                                             _lblDate, _txtDate, _chkSystem, _chkInvisible})

        Dim bar As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        _btnPreview.Bounds = New Rectangle(8, 7, 110, 26)
        _btnImport.Bounds = New Rectangle(126, 7, 110, 26)
        _btnClose.Bounds = New Rectangle(bar.Width - 96, 7, 80, 26)
        _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        bar.Controls.AddRange(New Control() {_btnPreview, _btnImport, _btnClose})

        AddHandler _btnBrowse.Click, Sub() PickSource()
        AddHandler _rbNames.CheckedChanged, Sub() _txtNames.Enabled = _rbNames.Checked
        AddHandler _btnPreview.Click, Async Sub() Await RunAsync(True)
        AddHandler _btnImport.Click, Async Sub() Await RunAsync(False)
        AddHandler _btnClose.Click, Sub() Close()

        Controls.Add(_txt)
        Controls.Add(top)
        Controls.Add(bar)
        CancelButton = _btnClose

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.ButtonRowLeft(8, 8, _btnPreview, _btnImport)
        _btnPreview.Top = 7 : _btnImport.Top = 7
        UiFit.FitButton(_btnClose)
        _btnClose.Left = ClientSize.Width - _btnClose.Width - 16
        UpdateEnabled()
    End Sub

    ''' <summary>True once a run saved the image, so MainForm knows to refresh.</summary>
    Public ReadOnly Property Changed As Boolean
        Get
            Return _changed
        End Get
    End Property

    ''' <summary>How many files the last real run copied, for the status line.</summary>
    Public ReadOnly Property Copied As Integer
        Get
            Return _copied
        End Get
    End Property

    Private Sub UpdateEnabled()
        Dim ready As Boolean = Not String.IsNullOrEmpty(_source)
        _btnPreview.Enabled = ready
        _btnImport.Enabled = ready
    End Sub

    Private Sub PickSource()
        Using dlg As New OpenFileDialog()
            dlg.Title = UiText.T("dlg.import.source")
            dlg.Filter = UiText.T("filter.mountimages")
            dlg.CheckFileExists = True
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim chosen As String = Path.GetFullPath(dlg.FileName)
            ' The engine refuses to add an image to itself; it cannot know that this is the same disk by
            ' another spelling, so the obvious case is caught here before the command is built.
            If String.Equals(chosen, Path.GetFullPath(_disk.Path), HostPlatform.PathComparison) Then
                MessageBox.Show(Me, UiText.T("msg.import.samefile"), UiText.T("btn.import"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            _source = chosen
            _txtSource.Text = chosen
            UpdateEnabled()
        End Using
    End Sub

    Private Sub SetBusy(busy As Boolean)
        UseWaitCursor = busy
        _btnPreview.Enabled = Not busy AndAlso Not String.IsNullOrEmpty(_source)
        _btnImport.Enabled = _btnPreview.Enabled
        _btnBrowse.Enabled = Not busy
    End Sub

    Private Sub Append(header As String, body As String)
        Dim sb As New StringBuilder(_txt.Text)
        If sb.Length > 0 Then sb.AppendLine()
        sb.AppendLine("==== " & header & " ====")
        sb.AppendLine(If(body, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf).TrimEnd())
        _txt.Text = sb.ToString()
        _txt.Select(_txt.Text.Length, 0)
        _txt.ScrollToCaret()
    End Sub

    ''' <summary>The request, or Nothing with a message box when the window is not filled in properly.</summary>
    Private Function BuildRequest() As Trs80Tool.ImportRequest
        If String.IsNullOrEmpty(_source) Then
            MessageBox.Show(Me, UiText.T("msg.import.nosource"), UiText.T("btn.import"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return Nothing
        End If
        Dim req As New Trs80Tool.ImportRequest With {
            .SourceImage = _source,
            .Overwrite = _chkOverwrite.Checked,
            .Binary = _chkBinary.Checked,
            .System = _chkSystem.Checked,
            .Invisible = _chkInvisible.Checked,
            .KeepGoing = _chkKeepGoing.Enabled AndAlso _chkKeepGoing.Checked}
        If _rbNames.Checked Then
            For Each part In _txtNames.Text.Split(New Char() {" "c, ControlChars.Tab, ";"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                req.Names.Add(part)
            Next
            If req.Names.Count = 0 Then
                MessageBox.Show(Me, UiText.T("msg.import.nonames"), UiText.T("btn.import"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return Nothing
            End If
        End If
        Dim typed As String = _txtDate.Text.Trim()
        If typed.Length > 0 Then
            Dim parsed As Date
            If Not Date.TryParseExact(typed, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
                MessageBox.Show(Me, UiText.T("msg.import.baddate"), UiText.T("btn.import"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If
            req.DateYmd = typed
        End If
        Return req
    End Function

    Private Async Function RunAsync(dryRun As Boolean) As Task
        Dim req As Trs80Tool.ImportRequest = BuildRequest()
        If req Is Nothing Then Return
        SetBusy(True)
        Try
            Dim r As ToolResult = Await Program.Tool.DiskImportAsync(_disk.Path, _disk.PartitionArg, req, dryRun)
            Dim header As String = "disk " & If(dryRun, "--dry-run ", "") & String.Join(" ", req.ToWords())
            Append(header, If(r.Succeeded, r.StdOut, r.ErrorText))
            If Not dryRun AndAlso r.Succeeded Then
                If r.Disk IsNot Nothing AndAlso r.Disk.Saved IsNot Nothing Then _changed = True
                Dim res As Trs80Tool.ImportResultJson = If(r.Disk Is Nothing, Nothing, r.Disk.ResultOf(Of Trs80Tool.ImportResultJson)("import"))
                _copied = If(res Is Nothing OrElse res.Files Is Nothing, 0, res.Files.Count)
            End If
        Catch ex As Exception
            Append("disk import", ex.Message)
        Finally
            SetBusy(False)
        End Try
    End Function
End Class
