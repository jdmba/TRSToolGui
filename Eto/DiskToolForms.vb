' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskToolForms.vb — windows of the Eto.Forms program for trs80-tool features
'  TRSTool never had: the text report window (Copy / Save / optional filter),
'  Check & repair, Trim, and Boot in the emulator.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

''' <summary>
''' Monospaced text with Copy, Save and Close. With a filter label and a refresher it also has a box whose
''' text is handed to the refresher (the granule map's PATTERN, for one).
''' </summary>
Public Class TextReportForm
    Inherits ToolWindow

    Private ReadOnly _txt As TextArea
    Private ReadOnly _filter As New TextBox With {.Width = 300}
    Private ReadOnly _btnRun As New Button With {.Text = "Run"}
    Private ReadOnly _refresher As Func(Of String, Task(Of String))
    Private ReadOnly _baseTitle As String

    Public Sub New(title As String, text As String)
        Me.New(title, text, Nothing, Nothing)
    End Sub

    Public Sub New(title As String, text As String, filterLabel As String, refresher As Func(Of String, Task(Of String)))
        _baseTitle = title
        _refresher = refresher
        Me.Title = title
        ClientSize = New Size(900, 620)
        MinimumSize = New Size(480, 300)
        Padding = New Padding(Ui.Margin)
        _txt = Ui.ReportBox(text)
        Dim buttons As TableLayout = Ui.Across(Ui.Btn("Copy to clipboard", Sub() Ui.CopyTextToClipboard(_txt.Text)), Ui.Btn("Save as...", AddressOf SaveText), Nothing, Ui.Btn("Close", Sub() Close()))
        If filterLabel IsNot Nothing AndAlso refresher IsNot Nothing Then
            AddHandler _btnRun.Click, Async Sub() Await RefreshAsync()
            AddHandler _filter.KeyDown, Async Sub(s, e)
                                            If e.Key = Keys.Enter Then
                                                e.Handled = True
                                                Await RefreshAsync()
                                            End If
                                        End Sub
            Content = Ui.Down(_txt, Ui.Across(New Label With {.Text = filterLabel, .VerticalAlignment = VerticalAlignment.Center}, _filter, _btnRun), _txt, buttons)
        Else
            Content = Ui.Down(_txt, _txt, buttons)
        End If
    End Sub

    Public Sub SetText(text As String)
        _txt.Text = Ui.Lf(text)
        Try
            _txt.CaretIndex = 0
        Catch
        End Try
    End Sub

    Private Async Function RefreshAsync() As Task
        If _refresher Is Nothing Then Return
        _btnRun.Enabled = False
        Try
            SetText(Await _refresher(_filter.Text.Trim()))
            Title = _baseTitle & If(_filter.Text.Trim().Length > 0, " [" & _filter.Text.Trim() & "]", "")
        Catch ex As Exception
            SetText("(" & ex.Message & ")")
        Finally
            _btnRun.Enabled = True
        End Try
    End Function

    Private Sub SaveText()
        Dim target As String = Ui.PickSave(Me, "Save report", "Text file (*.txt)|*.txt|All files (*.*)|*.*", Regex.Replace(_baseTitle, "[^A-Za-z0-9_.-]+", "_") & ".txt", "txt")
        If target Is Nothing Then Return
        Try
            File.WriteAllText(target, Ui.HostLines(_txt.Text), New UTF8Encoding(True))
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, "Save report")
        End Try
    End Sub
End Class

''' <summary>Disk → Check and repair: "disk check" on opening, then a dry run or the real "disk repair".</summary>
Public Class CheckRepairForm
    Inherits Dialog

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txt As TextArea
    Private ReadOnly _chkFree As New CheckBox With {.Text = "Also free granules the GAT marks used but no file claims (--free-unclaimed)"}
    Private ReadOnly _chkSizes As New CheckBox With {.Text = "Also fix program sizes whose transfer record is past the directory size (--fix-program-sizes)"}
    Private ReadOnly _chkMissing As New CheckBox With {.Text = "Also mark used the free granules the image doesn't fully hold, so the DOS won't use them either (--mark-missing-used)"}
    Private ReadOnly _btnCheck As New Button With {.Text = "Re-check"}
    Private ReadOnly _btnPreview As New Button With {.Text = "Preview repair (dry run)"}
    Private ReadOnly _btnRepair As New Button With {.Text = "Repair"}
    Private ReadOnly _btnClose As New Button With {.Text = "Close"}
    Private ReadOnly _canRepair As Boolean
    Private _changed As Boolean = False

    Public Sub New(disk As DiskImage)
        _disk = disk
        Title = "Check & repair - " & Path.GetFileName(disk.Path) & If(disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing, " - partition " & disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture), "")
        Resizable = True
        ClientSize = New Size(880, 600)
        MinimumSize = New Size(560, 360)
        Padding = New Padding(Ui.Margin)
        _txt = Ui.ReportBox("")
        _chkMissing.Enabled = Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.HasDiskOption("repair", "--mark-missing-used")
        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        AddHandler _btnPreview.Click, Async Sub() Await RunRepairAsync(True)
        AddHandler _btnRepair.Click, Async Sub() Await RunRepairAsync(False)
        AddHandler _btnClose.Click, Sub() Close()
        Content = Ui.Down(_txt, _txt, _chkFree, _chkSizes, _chkMissing, Ui.Across(_btnCheck, _btnPreview, _btnRepair, Nothing, _btnClose))
        AbortButton = _btnClose
        _canRepair = Not (disk.WriteProtected.HasValue AndAlso disk.WriteProtected.Value)
        _btnPreview.Enabled = _canRepair
        _btnRepair.Enabled = _canRepair
        If Not _canRepair Then _txt.Text = "The image's write-protect flag is set; clear it in Disk properties before repairing." & vbLf & vbLf
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
        _btnCheck.Enabled = Not busy
        _btnPreview.Enabled = Not busy AndAlso _canRepair
        _btnRepair.Enabled = Not busy AndAlso _canRepair
    End Sub

    Private Sub Append(header As String, body As String)
        Dim sb As New StringBuilder(_txt.Text)
        If sb.Length > 0 Then sb.Append(vbLf)
        sb.Append("==== " & header & " ====" & vbLf)
        sb.Append(Ui.Lf(body).TrimEnd() & vbLf)
        _txt.Text = sb.ToString()
        Try
            _txt.CaretIndex = _txt.Text.Length
        Catch
        End Try
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
            If Not Ui.Ask(Me, "Repair writes to the image (trs80-tool saves it only if every step succeeds)." & vbLf & vbLf & "Go ahead?", "Repair", MessageBoxType.Warning, True) Then Return
        End If
        SetBusy(True)
        Try
            Dim free As Boolean = Ui.IsOn(_chkFree)
            Dim sizes As Boolean = Ui.IsOn(_chkSizes)
            Dim missing As Boolean = _chkMissing.Enabled AndAlso Ui.IsOn(_chkMissing)
            Dim r As ToolResult = Await Program.Tool.DiskRepairAsync(_disk.Path, _disk.PartitionArg, free, sizes, missing, dryRun)
            Append(If(dryRun, "disk --dry-run repair", "disk repair") & If(free, " --free-unclaimed", "") & If(sizes, " --fix-program-sizes", "") & If(missing, " --mark-missing-used", ""), If(r.Succeeded, r.StdOut, r.ErrorText))
            If Not dryRun AndAlso r.Succeeded AndAlso r.Disk IsNot Nothing AndAlso r.Disk.Saved IsNot Nothing Then _changed = True
        Catch ex As Exception
            Append("disk repair", ex.Message)
        Finally
            SetBusy(False)
        End Try
    End Function
End Class

''' <summary>Disk → Trim: the option words for "disk IMG trim".</summary>
Public Class TrimForm
    Inherits OkCancelDialog

    Private ReadOnly _rbDefault, _rbCylinders, _rbSide As RadioButton
    Private ReadOnly _numCyl As New NumericStepper With {.MinValue = 1, .MaxValue = 255, .Value = 35, .DecimalPlaces = 0, .Width = 80}
    Private ReadOnly _chkForce As New CheckBox With {.Text = "Force: drop tracks whose data can't be proven unused, or trim a write-protected image" & vbLf & "(--force; never tracks with signs of copy protection)"}

    Public Sub New(disk As DiskImage)
        MyBase.New("Trim - " & Path.GetFileName(disk.Path))
        _rbDefault = New RadioButton With {.Text = "Drop only the cylinders trs80-tool can prove are past the disk's end (default)", .Checked = True}
        _rbCylinders = New RadioButton(_rbDefault) With {.Text = "Keep cylinders 0 to N-1, with N ="}
        _rbSide = New RadioButton(_rbDefault) With {.Text = "Drop side 1, keeping only side 0 (--side 0)"}
        If disk.Geometry IsNot Nothing AndAlso disk.Geometry.CylinderCount > 0 Then _numCyl.Value = Math.Min(255, Math.Max(1, disk.Geometry.CylinderCount))
        AddHandler _rbCylinders.CheckedChanged, Sub() _numCyl.Enabled = _rbCylinders.Checked
        _numCyl.Enabled = False
        Finish(Ui.Down(Nothing,
                       Ui.Lbl("Remove the part of the image the disk doesn't use. trs80-tool refuses to drop tracks that hold data unless forced."),
                       _rbDefault, Ui.Across(_rbCylinders, _numCyl), _rbSide, _chkForce), "Trim")
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _rbCylinders.Checked Then
            w.Add("--cylinders") : w.Add(CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture))
        ElseIf _rbSide.Checked Then
            w.Add("--side") : w.Add("0")
        End If
        If Ui.IsOn(_chkForce) Then w.Add("--force")
        Return w
    End Function
End Class

''' <summary>Disk → Boot in the emulator: the words for "trs80-tool run IMG ...".</summary>
Public Class BootForm
    Inherits OkCancelDialog

    Private ReadOnly _rbM1, _rbM3, _rbM4 As RadioButton
    Private ReadOnly _chkWritable As New CheckBox With {.Text = "Let the emulated DOS write to the image (--writable; otherwise every floppy is write-protected)"}
    Private ReadOnly _lstMount As New ListBox With {.Size = New Size(460, 100)}
    Private ReadOnly _txtPrinter As New TextBox With {.Width = 380}

    Public Sub New(disk As DiskImage)
        MyBase.New("Boot in the emulator - " & Path.GetFileName(disk.Path))
        _rbM1 = New RadioButton With {.Text = "Model I"}
        _rbM3 = New RadioButton(_rbM1) With {.Text = "Model III"}
        _rbM4 = New RadioButton(_rbM1) With {.Text = "Model 4"}
        Select Case GuessModel(disk.DosName)
            Case 1 : _rbM1.Checked = True
            Case 4 : _rbM4.Checked = True
            Case Else : _rbM3.Checked = True
        End Select
        Dim mountButtons As TableLayout = Ui.Down(Nothing, Ui.Btn("Add...", AddressOf AddMount), Ui.Btn("Remove", Sub()
                                                                                                                   If _lstMount.SelectedIndex >= 0 Then _lstMount.Items.RemoveAt(_lstMount.SelectedIndex)
                                                                                                               End Sub))
        Finish(Ui.Down(Nothing,
                       Ui.Wrapped("Starts ""trs80-tool run"" in a terminal window with this image in drive 0. Close that window (or Ctrl-C) to stop the emulator; the file list here is read again when it ends.", 560),
                       Ui.Group("TRS-80 model (--model)", Ui.Across(_rbM1, _rbM3, _rbM4)),
                       _chkWritable,
                       Ui.Lbl("Also mount in the next drives (--mount):"),
                       Ui.Across(_lstMount, mountButtons),
                       Ui.Lbl("Printer output file (--printer, optional):"),
                       Ui.Across(_txtPrinter, Ui.Btn("Browse...", Sub()
                                                                      Dim chosen As String = Ui.PickSave(Me, "Printer output file", "Text file (*.txt)|*.txt|All files (*.*)|*.*", Nothing, "txt")
                                                                      If chosen IsNot Nothing Then _txtPrinter.Text = chosen
                                                                  End Sub))), "Boot")
    End Sub

    Public Shared Function GuessModel(dosName As String) As Integer
        Dim d As String = If(dosName, "")
        If Regex.IsMatch(d, "^(TRSDOS|LS-DOS) 6\.") OrElse d.IndexOf("Model 4", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 4
        If Regex.IsMatch(d, "^TRSDOS 2\.") OrElse Regex.IsMatch(d, "^(VTOS|DoubleDOS|NEWDOS 2|NEWDOS/80 v1)", RegexOptions.IgnoreCase) Then Return 1
        Return 3
    End Function

    Private Sub AddMount()
        Using ofd As New OpenFileDialog()
            ofd.Title = "Image to mount"
            Ui.SetFilters(ofd, "Disk and cassette images (*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.cas;*.wav)|*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.cas;*.wav|All files (*.*)|*.*")
            ofd.MultiSelect = True
            If ofd.ShowDialog(Me) <> DialogResult.Ok Then Return
            For Each f In ofd.Filenames
                Dim name As String = f
                If Not _lstMount.Items.Any(Function(it) it.Text = name) Then _lstMount.Items.Add(name)
            Next
        End Using
    End Sub

    Public Function RunWords(imagePath As String) As List(Of String)
        Dim w As New List(Of String) From {"run", imagePath, "--model", If(_rbM1.Checked, "1", If(_rbM4.Checked, "4", "3"))}
        If Ui.IsOn(_chkWritable) Then w.Add("--writable")
        If _txtPrinter.Text.Trim().Length > 0 Then
            w.Add("--printer") : w.Add(_txtPrinter.Text.Trim())
        End If
        If _lstMount.Items.Count > 0 Then
            w.Add("--mount")
            For Each o In _lstMount.Items
                w.Add(o.Text)
            Next
        End If
        Return w
    End Function
End Class
