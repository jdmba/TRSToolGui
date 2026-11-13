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
            .Font = New Font("Consolas", 9.5F)
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With
        Dim panel As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        Dim btnCopy As New Button() With {.Text = "Copy to clipboard", .Bounds = New Rectangle(8, 7, 130, 26)}
        AddHandler btnCopy.Click, Sub()
                                      Try : Clipboard.SetText(If(_txt.Text.Length = 0, " ", _txt.Text)) : Catch : End Try
                                  End Sub
        Dim btnSave As New Button() With {.Text = "Save as...", .Bounds = New Rectangle(146, 7, 100, 26)}
        AddHandler btnSave.Click, Sub() SaveText()
        Dim btnClose As New Button() With {.Text = "Close", .DialogResult = DialogResult.OK, .Anchor = AnchorStyles.Top Or AnchorStyles.Right, .Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)}
        AddHandler btnClose.Click, Sub() Close()
        panel.Controls.AddRange(New Control() {btnCopy, btnSave, btnClose})
        If filterLabel IsNot Nothing AndAlso refresher IsNot Nothing Then
            Dim top As New Panel() With {.Dock = DockStyle.Top, .Height = 34}
            Dim lbl As New Label() With {.Text = filterLabel, .AutoSize = True, .Location = New Point(8, 9)}
            _filter.Bounds = New Rectangle(lbl.PreferredWidth + 16, 6, 300, 23)
            _btnRun.Text = "Run"
            _btnRun.Bounds = New Rectangle(_filter.Right + 8, 5, 70, 25)
            AddHandler _btnRun.Click, Async Sub() Await RefreshAsync()
            AddHandler _filter.KeyDown, Async Sub(s, e)
                                            If e.KeyCode = Keys.Enter Then
                                                e.Handled = True : e.SuppressKeyPress = True
                                                Await RefreshAsync()
                                            End If
                                        End Sub
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
            sfd.Title = "Save report"
            sfd.Filter = "Text file (*.txt)|*.txt|All files (*.*)|*.*"
            sfd.FileName = Regex.Replace(_baseTitle, "[^A-Za-z0-9_.-]+", "_") & ".txt"
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                File.WriteAllText(sfd.FileName, _txt.Text, Encoding.UTF8)
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, "Save report", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
    Private ReadOnly _chkFree As New CheckBox() With {.Text = "Also free granules the GAT marks used but no file claims (--free-unclaimed)", .AutoSize = True}
    Private ReadOnly _chkSizes As New CheckBox() With {.Text = "Also fix program sizes whose transfer record is past the directory size (--fix-program-sizes)", .AutoSize = True}
    Private ReadOnly _chkMissing As New CheckBox() With {.Text = "Also mark used the free granules the image doesn't fully hold, so the DOS won't use them either (--mark-missing-used)", .AutoSize = True}
    Private ReadOnly _btnCheck As New Button()
    Private ReadOnly _btnPreview As New Button()
    Private ReadOnly _btnRepair As New Button()
    Private ReadOnly _btnClose As New Button()
    Private _changed As Boolean = False

    Public Sub New(disk As DiskImage)
        _disk = disk
        Text = "Check & repair - " & Path.GetFileName(disk.Path) & If(disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing, " - partition " & disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture), "")
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(860, 600)
        MinimumSize = New Size(560, 360)
        ShowInTaskbar = False
        With _txt
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font("Consolas", 9.5F)
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
        _btnCheck.Text = "Re-check" : _btnCheck.Bounds = New Rectangle(8, 7, 90, 26)
        _btnPreview.Text = "Preview repair (dry run)" : _btnPreview.Bounds = New Rectangle(104, 7, 170, 26)
        _btnRepair.Text = "Repair" : _btnRepair.Bounds = New Rectangle(280, 7, 90, 26)
        _btnClose.Text = "Close" : _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right : _btnClose.Bounds = New Rectangle(panel.Width - 96, 7, 80, 26)
        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        AddHandler _btnPreview.Click, Async Sub() Await RunRepairAsync(True)
        AddHandler _btnRepair.Click, Async Sub() Await RunRepairAsync(False)
        AddHandler _btnClose.Click, Sub() Close()
        panel.Controls.AddRange(New Control() {_btnCheck, _btnPreview, _btnRepair, _btnClose})
        Controls.Add(_txt)
        Controls.Add(opts)
        Controls.Add(panel)
        CancelButton = _btnClose
        Dim canRepair As Boolean = disk.WriteProtected.HasValue AndAlso Not disk.WriteProtected.Value OrElse Not disk.WriteProtected.HasValue
        _btnPreview.Enabled = canRepair
        _btnRepair.Enabled = canRepair
        If Not canRepair Then _txt.Text = "The image's write-protect flag is set; clear it in Disk properties before repairing." & vbCrLf & vbCrLf
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
            If MessageBox.Show(Me, "Repair writes to the image (trs80-tool saves it only if every step succeeds)." & vbCrLf & vbCrLf & "Go ahead?", "Repair", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
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

    Private ReadOnly _rbDefault As New RadioButton() With {.Text = "Drop only the cylinders trs80-tool can prove are past the disk's end (default)", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbCylinders As New RadioButton() With {.Text = "Keep cylinders 0 to N-1, with N =", .AutoSize = True}
    Private ReadOnly _rbSide As New RadioButton() With {.Text = "Drop side 1, keeping only side 0 (--side 0)", .AutoSize = True}
    Private ReadOnly _numCyl As New NumericUpDown() With {.Minimum = 1, .Maximum = 255, .Value = 35, .Width = 60}
    Private ReadOnly _chkForce As New CheckBox() With {.Text = "Force: drop tracks whose data can't be proven unused, or trim a write-protected image (--force; never tracks with signs of copy protection)", .AutoSize = True}

    Public Sub New(disk As DiskImage)
        Text = "Trim - " & Path.GetFileName(disk.Path)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(640, 200)
        If disk.Geometry IsNot Nothing AndAlso disk.Geometry.CylinderCount > 0 Then _numCyl.Value = Math.Min(255, Math.Max(1, disk.Geometry.CylinderCount))
        Controls.Add(New Label() With {.Text = "Remove the part of the image the disk doesn't use. trs80-tool refuses to drop tracks that hold data unless forced.", .AutoSize = True, .Location = New Point(12, 12)})
        _rbDefault.Location = New Point(14, 40)
        _rbCylinders.Location = New Point(14, 66)
        _numCyl.Location = New Point(_rbCylinders.Left + _rbCylinders.PreferredSize.Width + 8, 64)
        _rbSide.Location = New Point(14, 92)
        _chkForce.Location = New Point(14, 124)
        Controls.AddRange(New Control() {_rbDefault, _rbCylinders, _numCyl, _rbSide, _chkForce})
        AddHandler _rbCylinders.CheckedChanged, Sub() _numCyl.Enabled = _rbCylinders.Checked
        _numCyl.Enabled = False
        Dim btnOk As New Button() With {.Text = "Trim", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(460, 162, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(550, 162, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)
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

    Private ReadOnly _rbM1 As New RadioButton() With {.Text = "Model I", .AutoSize = True}
    Private ReadOnly _rbM3 As New RadioButton() With {.Text = "Model III", .AutoSize = True}
    Private ReadOnly _rbM4 As New RadioButton() With {.Text = "Model 4", .AutoSize = True}
    Private ReadOnly _chkWritable As New CheckBox() With {.Text = "Let the emulated DOS write to the image (--writable; otherwise every floppy is write-protected)", .AutoSize = True}
    Private ReadOnly _lstMount As New ListBox()
    Private ReadOnly _txtPrinter As New TextBox() With {.Width = 360}

    Public Sub New(disk As DiskImage)
        Text = "Boot in the emulator - " & Path.GetFileName(disk.Path)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 360)
        Controls.Add(New Label() With {.Text = "Starts ""trs80-tool run"" in a new console window with this image in drive 0. Close that window (or Ctrl-C) to stop the emulator; then use View → Refresh here if the DOS changed the image.", .AutoSize = False, .Bounds = New Rectangle(12, 10, 536, 36)})
        Dim grp As New GroupBox() With {.Text = "TRS-80 model (--model)", .Bounds = New Rectangle(12, 50, 536, 50)}
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
        Controls.Add(New Label() With {.Text = "Also mount in the next drives (--mount):", .AutoSize = True, .Location = New Point(12, 140)})
        _lstMount.Bounds = New Rectangle(12, 160, 430, 90)
        Controls.Add(_lstMount)
        Dim btnAdd As New Button() With {.Text = "Add...", .Bounds = New Rectangle(452, 160, 96, 26)}
        Dim btnRemove As New Button() With {.Text = "Remove", .Bounds = New Rectangle(452, 192, 96, 26)}
        AddHandler btnAdd.Click, Sub() AddMount()
        AddHandler btnRemove.Click, Sub()
                                        If _lstMount.SelectedIndex >= 0 Then _lstMount.Items.RemoveAt(_lstMount.SelectedIndex)
                                    End Sub
        Controls.Add(btnAdd) : Controls.Add(btnRemove)
        Controls.Add(New Label() With {.Text = "Printer output file (--printer, optional):", .AutoSize = True, .Location = New Point(12, 262)})
        _txtPrinter.Location = New Point(12, 282)
        Controls.Add(_txtPrinter)
        Dim btnPrinter As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(380, 281, 80, 25)}
        AddHandler btnPrinter.Click, Sub()
                                         Using sfd As New SaveFileDialog()
                                             sfd.Title = "Printer output file"
                                             sfd.Filter = "Text file (*.txt)|*.txt|All files (*.*)|*.*"
                                             sfd.OverwritePrompt = False
                                             If sfd.ShowDialog(Me) = DialogResult.OK Then _txtPrinter.Text = sfd.FileName
                                         End Using
                                     End Sub
        Controls.Add(btnPrinter)
        Dim btnOk As New Button() With {.Text = "Boot", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(380, 322, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(468, 322, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)
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
            ofd.Title = "Image to mount"
            ofd.Filter = "Disk and cassette images (*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.cas;*.wav)|*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.cas;*.wav|All files (*.*)|*.*"
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
