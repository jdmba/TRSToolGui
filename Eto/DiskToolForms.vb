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
    Private ReadOnly _btnRun As New Button With {.Text = UiText.T("btn.run")}
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
        Dim buttons As TableLayout = Ui.Across(Ui.Btn(UiText.T("btn.copyclipboard"), Sub() Ui.CopyTextToClipboard(_txt.Text)), Ui.Btn(UiText.T("btn.saveas"), AddressOf SaveText), Nothing, Ui.Btn(UiText.T("btn.close"), Sub() Close()))
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
        Dim target As String = Ui.PickSave(Me, UiText.T("caption.savereport"), UiText.T("filter.textfile"), Regex.Replace(_baseTitle, "[^A-Za-z0-9_.-]+", "_") & ".txt", "txt")
        If target Is Nothing Then Return
        Try
            File.WriteAllText(target, Ui.HostLines(_txt.Text), New UTF8Encoding(True))
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, UiText.T("caption.savereport"))
        End Try
    End Sub
End Class

''' <summary>Disk → Check and repair: "disk check" on opening, then a dry run or the real "disk repair".</summary>
Public Class CheckRepairForm
    Inherits Dialog

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txt As TextArea
    Private ReadOnly _chkFree As New CheckBox With {.Text = UiText.T("dlg.checkrepair.free")}
    Private ReadOnly _chkSizes As New CheckBox With {.Text = UiText.T("dlg.checkrepair.sizes")}
    Private ReadOnly _chkMissing As New CheckBox With {.Text = UiText.T("dlg.checkrepair.missing")}
    Private ReadOnly _btnCheck As New Button With {.Text = UiText.T("btn.recheck")}
    Private ReadOnly _btnPreview As New Button With {.Text = UiText.T("btn.previewrepair")}
    Private ReadOnly _btnRepair As New Button With {.Text = UiText.T("btn.repair")}
    Private ReadOnly _btnClose As New Button With {.Text = UiText.T("btn.close")}
    Private ReadOnly _canRepair As Boolean
    Private _changed As Boolean = False

    Public Sub New(disk As DiskImage)
        _disk = disk
        Dim shown As String = UiText.T("dlg.checkrepair.title", Path.GetFileName(disk.Path))
        If disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing Then shown = UiText.T("title.partition", shown, disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Title = shown
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
        If Not _canRepair Then _txt.Text = UiText.T("dlg.checkrepair.wp") & vbLf & vbLf
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
            If Not Ui.Ask(Me, UiText.T("dlg.checkrepair.confirm"), UiText.T("caption.repair"), MessageBoxType.Warning, True) Then Return
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
    Private ReadOnly _chkForce As New CheckBox With {.Text = UiText.T("dlg.trim.force") & vbLf & UiText.T("dlg.trim.forcenote")}

    Public Sub New(disk As DiskImage)
        MyBase.New(UiText.T("dlg.trim.title", Path.GetFileName(disk.Path)))
        _rbDefault = New RadioButton With {.Text = UiText.T("dlg.trim.default"), .Checked = True}
        _rbCylinders = New RadioButton(_rbDefault) With {.Text = UiText.T("dlg.trim.cylinders")}
        _rbSide = New RadioButton(_rbDefault) With {.Text = UiText.T("dlg.trim.side")}
        If disk.Geometry IsNot Nothing AndAlso disk.Geometry.CylinderCount > 0 Then _numCyl.Value = Math.Min(255, Math.Max(1, disk.Geometry.CylinderCount))
        AddHandler _rbCylinders.CheckedChanged, Sub() _numCyl.Enabled = _rbCylinders.Checked
        _numCyl.Enabled = False
        Finish(Ui.Down(Nothing,
                       Ui.Lbl(UiText.T("dlg.trim.intro")),
                       _rbDefault, Ui.Across(_rbCylinders, _numCyl), _rbSide, _chkForce), UiText.T("btn.trim"))
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
    Private ReadOnly _chkWritable As New CheckBox With {.Text = UiText.T("dlg.boot.writable")}
    Private ReadOnly _lstMount As New ListBox With {.Size = New Size(460, 100)}
    Private ReadOnly _txtPrinter As New TextBox With {.Width = 380}

    Public Sub New(disk As DiskImage)
        MyBase.New(UiText.T("dlg.boot.title", Path.GetFileName(disk.Path)))
        _rbM1 = New RadioButton With {.Text = UiText.T("word.model1")}
        _rbM3 = New RadioButton(_rbM1) With {.Text = UiText.T("word.model3")}
        _rbM4 = New RadioButton(_rbM1) With {.Text = UiText.T("word.model4")}
        Select Case GuessModel(disk.DosName)
            Case 1 : _rbM1.Checked = True
            Case 4 : _rbM4.Checked = True
            Case Else : _rbM3.Checked = True
        End Select
        Dim mountButtons As TableLayout = Ui.Down(Nothing, Ui.Btn(UiText.T("btn.addellipsis"), AddressOf AddMount), Ui.Btn(UiText.T("btn.remove"), Sub()
                                                                                                                   If _lstMount.SelectedIndex >= 0 Then _lstMount.Items.RemoveAt(_lstMount.SelectedIndex)
                                                                                                               End Sub))
        Finish(Ui.Down(Nothing,
                       Ui.Wrapped(UiText.T("dlg.boot.intro.unix"), 560),
                       Ui.Group(UiText.T("dlg.boot.model"), Ui.Across(_rbM1, _rbM3, _rbM4)),
                       _chkWritable,
                       Ui.Lbl(UiText.T("dlg.boot.mount")),
                       Ui.Across(_lstMount, mountButtons),
                       Ui.Lbl(UiText.T("dlg.boot.printer")),
                       Ui.Across(_txtPrinter, Ui.Btn(UiText.T("btn.browse"), Sub()
                                                                      Dim chosen As String = Ui.PickSave(Me, UiText.T("dlg.boot.pickprinter"), UiText.T("filter.textfile"), Nothing, "txt")
                                                                      If chosen IsNot Nothing Then _txtPrinter.Text = chosen
                                                                  End Sub))), UiText.T("btn.boot"))
    End Sub

    Public Shared Function GuessModel(dosName As String) As Integer
        Dim d As String = If(dosName, "")
        If Regex.IsMatch(d, "^(TRSDOS|LS-DOS) 6\.") OrElse d.IndexOf("Model 4", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 4
        If Regex.IsMatch(d, "^TRSDOS 2\.") OrElse Regex.IsMatch(d, "^(VTOS|DoubleDOS|NEWDOS 2|NEWDOS/80 v1)", RegexOptions.IgnoreCase) Then Return 1
        Return 3
    End Function

    Private Sub AddMount()
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("dlg.boot.pickmount")
            Ui.SetFilters(ofd, UiText.T("filter.mountimages"))
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

''' <summary>
''' Disk → Import from another image: "disk IMG import SOURCE [PATTERN...]", new in trs80-tool 5.3.0.
''' The same window as the Windows program's (DiskToolForms.vb), built on Eto. The engine copies the files,
''' so each keeps its own date and its system and hidden flags and BASIC is converted between Model III and
''' Model 4 where the two disks differ. Preview is the same command with --dry-run.
''' </summary>
Public Class ImportForm
    Inherits Dialog

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txt As TextArea
    Private ReadOnly _txtSource As New TextBox With {.ReadOnly = True, .Text = UiText.T("dlg.import.nosource"), .Width = 460}
    Private ReadOnly _btnBrowse As Button
    Private ReadOnly _rbAll, _rbNames As RadioButton
    Private ReadOnly _txtNames As New TextBox With {.Width = 240, .Enabled = False}
    Private ReadOnly _chkOverwrite As New CheckBox With {.Text = UiText.T("dlg.import.overwrite")}
    Private ReadOnly _chkKeepGoing As New CheckBox With {.Text = UiText.T("dlg.import.keepgoing"), .Checked = True}
    Private ReadOnly _chkBinary As New CheckBox With {.Text = UiText.T("dlg.import.binary")}
    Private ReadOnly _chkSystem As New CheckBox With {.Text = UiText.T("dlg.import.system")}
    Private ReadOnly _chkInvisible As New CheckBox With {.Text = UiText.T("dlg.import.invisible")}
    Private ReadOnly _txtDate As New TextBox With {.Width = 120}
    Private ReadOnly _btnPreview As New Button With {.Text = UiText.T("btn.previewimport")}
    Private ReadOnly _btnImport As New Button With {.Text = UiText.T("btn.import")}
    Private ReadOnly _btnClose As New Button With {.Text = UiText.T("btn.close")}
    Private _source As String = Nothing
    Private _changed As Boolean = False
    Private _copied As Integer = 0

    Public Sub New(disk As DiskImage)
        _disk = disk
        Dim shown As String = UiText.T("dlg.import.title", Path.GetFileName(disk.Path))
        If disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing Then shown = UiText.T("title.partition", shown, disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Title = shown
        Resizable = True
        ClientSize = New Size(900, 640)
        MinimumSize = New Size(620, 460)
        Padding = New Padding(Ui.Margin)
        _txt = Ui.ReportBox("")
        _btnBrowse = Ui.Btn(UiText.T("btn.browse"), AddressOf PickSource)
        _rbAll = New RadioButton With {.Text = UiText.T("dlg.import.whichall"), .Checked = True}
        _rbNames = New RadioButton(_rbAll) With {.Text = UiText.T("dlg.import.whichnames")}

        ' trs80-tool 5.2.0 has no --keep-going; the box says so by being off and disabled.
        If Not Program.Tool.CanKeepGoing Then
            _chkKeepGoing.Checked = False
            _chkKeepGoing.Enabled = False
        End If

        AddHandler _rbNames.CheckedChanged, Sub() _txtNames.Enabled = _rbNames.Checked
        AddHandler _btnPreview.Click, Async Sub() Await RunAsync(True)
        AddHandler _btnImport.Click, Async Sub() Await RunAsync(False)
        AddHandler _btnClose.Click, Sub() Close()

        Content = Ui.Down(_txt,
                          Ui.Lbl(UiText.T("dlg.import.intro")),
                          Ui.Lbl(UiText.T("dlg.import.keeps"), True),
                          Ui.Across(Ui.Lbl(UiText.T("dlg.import.source")), _txtSource, _btnBrowse),
                          _rbAll,
                          Ui.Across(_rbNames, _txtNames),
                          Ui.Lbl(UiText.T("dlg.import.nameshint"), True),
                          _chkOverwrite, _chkKeepGoing, _chkBinary, _chkSystem, _chkInvisible,
                          Ui.Across(Ui.Lbl(UiText.T("dlg.import.usedate")), _txtDate),
                          _txt,
                          Ui.Across(_btnPreview, _btnImport, Nothing, _btnClose))
        AbortButton = _btnClose
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
        Dim chosen As String = Ui.PickFile(Me, UiText.T("dlg.import.source"), UiText.T("filter.mountimages"))
        If String.IsNullOrEmpty(chosen) Then Return
        chosen = Path.GetFullPath(chosen)
        ' The engine refuses to add an image to itself; it cannot know that this is the same disk by another
        ' spelling, so the obvious case is caught here before the command is built.
        If String.Equals(chosen, Path.GetFullPath(_disk.Path), HostPlatform.PathComparison) Then
            Ui.Warn(Me, UiText.T("msg.import.samefile"), UiText.T("btn.import"))
            Return
        End If
        _source = chosen
        _txtSource.Text = chosen
        UpdateEnabled()
    End Sub

    Private Sub SetBusy(busy As Boolean)
        _btnPreview.Enabled = Not busy AndAlso Not String.IsNullOrEmpty(_source)
        _btnImport.Enabled = _btnPreview.Enabled
        _btnBrowse.Enabled = Not busy
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

    ''' <summary>The request, or Nothing with a message when the window is not filled in properly.</summary>
    Private Function BuildRequest() As Trs80Tool.ImportRequest
        If String.IsNullOrEmpty(_source) Then
            Ui.Info(Me, UiText.T("msg.import.nosource"), UiText.T("btn.import"))
            Return Nothing
        End If
        Dim req As New Trs80Tool.ImportRequest With {
            .SourceImage = _source,
            .Overwrite = Ui.IsOn(_chkOverwrite),
            .Binary = Ui.IsOn(_chkBinary),
            .System = Ui.IsOn(_chkSystem),
            .Invisible = Ui.IsOn(_chkInvisible),
            .KeepGoing = _chkKeepGoing.Enabled AndAlso Ui.IsOn(_chkKeepGoing)}
        If _rbNames.Checked Then
            For Each part In _txtNames.Text.Split(New Char() {" "c, ControlChars.Tab, ";"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                req.Names.Add(part)
            Next
            If req.Names.Count = 0 Then
                Ui.Info(Me, UiText.T("msg.import.nonames"), UiText.T("btn.import"))
                Return Nothing
            End If
        End If
        Dim typed As String = _txtDate.Text.Trim()
        If typed.Length > 0 Then
            Dim parsed As Date
            If Not Date.TryParseExact(typed, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
                Ui.Warn(Me, UiText.T("msg.import.baddate"), UiText.T("btn.import"))
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
