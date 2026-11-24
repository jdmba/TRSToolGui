' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HostToolForms.vb — dialogs of the Eto.Forms program for trs80-tool's
'  commands on host files: ConvertFileForm ("trs80-tool convert") and
'  AssembleForm ("trs80-tool asm"). Each shows the tool's own output and
'  remembers the file it produced, so the main window can offer to add it to
'  the open disk.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

''' <summary>Tools → Convert file(s): trs80-tool convert INFILES… OUTFILE with its switches.</summary>
Public Class ConvertFileForm
    Inherits Dialog

    Private ReadOnly _lstIn As New ListBox With {.Size = New Size(560, 100)}
    Private ReadOnly _txtOut As New TextBox With {.Width = 560}
    Private ReadOnly _cmbBaud As New DropDown With {.Width = 120}
    Private ReadOnly _txtStart As New TextBox With {.Width = 120}
    Private ReadOnly _chkSystem As New CheckBox With {.Text = UiText.T("dlg.convert.system")}
    Private ReadOnly _chkCpm As New CheckBox With {.Text = UiText.T("dlg.convert.cpm")}
    Private ReadOnly _chkForce As New CheckBox With {.Text = UiText.T("dlg.convert.force")}
    Private ReadOnly _txtResult As TextArea
    Private ReadOnly _btnRun As New Button With {.Text = UiText.T("btn.convert")}
    Private ReadOnly _btnClose As New Button With {.Text = UiText.T("btn.close")}
    Private _outputPath As String = Nothing

    Public Sub New(initialInputs As IEnumerable(Of String))
        Title = UiText.T("dlg.convert.title")
        Resizable = True
        ClientSize = New Size(720, 580)
        MinimumSize = New Size(600, 480)
        Padding = New Padding(Ui.Margin)
        If initialInputs IsNot Nothing Then
            For Each f In initialInputs
                _lstIn.Items.Add(f)
            Next
        End If
        ' Only item 0 is text: the others are the numbers trs80-tool is given (SelectedIndex > 0 below).
        For Each item In New String() {UiText.T("dlg.convert.bauddefault"), "250", "500", "1500"}
            _cmbBaud.Items.Add(item)
        Next
        _cmbBaud.SelectedIndex = 0
        _txtResult = Ui.ReportBox("", 9.0F)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        AddHandler _btnClose.Click, Sub() Close()
        Dim inButtons As TableLayout = Ui.Down(Nothing, Ui.Btn(UiText.T("btn.addellipsis"), AddressOf AddInputs), Ui.Btn(UiText.T("btn.remove"), Sub()
                                                                                                                  If _lstIn.SelectedIndex >= 0 Then _lstIn.Items.RemoveAt(_lstIn.SelectedIndex)
                                                                                                              End Sub))
        Dim inRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        inRow.Rows.Add(New TableRow(New TableCell(_lstIn, True), New TableCell(inButtons)))
        Dim outRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        outRow.Rows.Add(New TableRow(New TableCell(_txtOut, True), New TableCell(Ui.Btn(UiText.T("btn.browse"), AddressOf BrowseOutput))))
        Content = Ui.Down(_txtResult,
                          Ui.Lbl(UiText.T("dlg.convert.inputs")),
                          inRow,
                          Ui.Lbl(UiText.T("dlg.convert.output")),
                          outRow,
                          Ui.Fields(UiText.T("dlg.convert.baud"), Ui.Across(_cmbBaud),
                                    UiText.T("dlg.convert.start"), Ui.Across(_txtStart)),
                          _chkSystem, _chkCpm, _chkForce,
                          Ui.Lbl(UiText.T("msg.engineoutput")),
                          _txtResult,
                          Ui.Across(Nothing, _btnRun, _btnClose))
        AbortButton = _btnClose
        UpdateEnabled()
    End Sub

    ''' <summary>The file produced by the last successful conversion, or Nothing.</summary>
    Public ReadOnly Property OutputPath As String
        Get
            Return _outputPath
        End Get
    End Property

    Private Sub UpdateEnabled()
        Dim ext As String = Path.GetExtension(_txtOut.Text.Trim()).ToLowerInvariant()
        _cmbBaud.Enabled = (ext = ".cas" OrElse ext = ".wav")
        _txtStart.Enabled = (ext = ".cmd" OrElse ext = ".3bn")
        _chkForce.Enabled = (ext = ".dmk" OrElse ext = ".jv1" OrElse ext = ".jv3" OrElse ext = ".hfe" OrElse ext = ".imd")
    End Sub

    Private Sub AddInputs()
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("dlg.convert.pickinputs")
            Ui.SetFilters(ofd, UiText.T("filter.convertin"))
            ofd.MultiSelect = True
            If ofd.ShowDialog(Me) <> DialogResult.Ok Then Return
            For Each f In ofd.Filenames
                Dim name As String = f
                If Not _lstIn.Items.Any(Function(it) it.Text = name) Then _lstIn.Items.Add(name)
            Next
        End Using
        If _txtOut.Text.Trim().Length = 0 AndAlso _lstIn.Items.Count = 1 Then
            ' Suggest an output next to the input with the "other" extension.
            Dim first As String = _lstIn.Items(0).Text
            Dim suggest As String = Nothing
            Select Case Path.GetExtension(first).ToLowerInvariant()
                Case ".bas" : suggest = ".asc"
                Case ".asc" : suggest = ".bas"
                Case ".cmd" : suggest = ".3bn"
                Case ".3bn", ".cas" : suggest = ".cmd"
                Case ".wav" : suggest = ".cas"
                Case ".rom", ".bin", ".hex" : suggest = ".cmd"
                Case ".jv1", ".jv3", ".dsk", ".hfe", ".imd", ".scp" : suggest = ".dmk"
                Case ".dmk" : suggest = ".jv3"
            End Select
            If suggest IsNot Nothing Then _txtOut.Text = Path.ChangeExtension(first, suggest)
        End If
    End Sub

    Private Sub BrowseOutput()
        Dim chosen As String = Ui.PickSave(Me, UiText.T("dlg.convert.pickoutput"), UiText.T("filter.convertout"),
                                           If(_txtOut.Text.Trim().Length > 0, _txtOut.Text.Trim(), Nothing))
        If chosen IsNot Nothing Then _txtOut.Text = chosen
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _cmbBaud.Enabled AndAlso _cmbBaud.SelectedIndex > 0 Then
            w.Add("--baud") : w.Add(_cmbBaud.Items(_cmbBaud.SelectedIndex).Text)
        End If
        If _txtStart.Enabled AndAlso _txtStart.Text.Trim().Length > 0 Then
            w.Add("--start") : w.Add(_txtStart.Text.Trim())
        End If
        If Ui.IsOn(_chkSystem) Then w.Add("--system")
        If Ui.IsOn(_chkCpm) Then w.Add("--cpm-records")
        If _chkForce.Enabled AndAlso Ui.IsOn(_chkForce) Then w.Add("--force")
        Return w
    End Function

    Private Async Function RunAsync() As Task
        If _lstIn.Items.Count = 0 Then
            Ui.Warn(Me, UiText.T("dlg.convert.noinputs"), UiText.T("btn.convert"))
            Return
        End If
        Dim outFile As String = _txtOut.Text.Trim()
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            Ui.Warn(Me, UiText.T("dlg.convert.nooutput"), UiText.T("btn.convert"))
            Return
        End If
        Dim inputs As New List(Of String)
        For Each o In _lstIn.Items
            inputs.Add(o.Text)
        Next
        _btnRun.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.ConvertFilesAsync(inputs, outFile, OptionWords())
            _txtResult.Text = Ui.Lf(r.CommandLine & vbLf & vbLf & If(r.Succeeded, r.StdOut, r.ErrorText))
            If r.Succeeded AndAlso File.Exists(outFile) Then _outputPath = outFile
        Catch ex As Exception
            _txtResult.Text = ex.Message
        Finally
            _btnRun.Enabled = True
        End Try
    End Function
End Class

''' <summary>Tools → Assemble: trs80-tool asm SOURCE.asm OUTPUT [--model N] [--listing OUT.lst] [--baud N].</summary>
Public Class AssembleForm
    Inherits Dialog

    Private ReadOnly _txtSrc As New TextBox With {.Width = 560}
    Private ReadOnly _txtOut As New TextBox With {.Width = 560}
    Private ReadOnly _rbM1, _rbM3, _rbM4 As RadioButton
    Private ReadOnly _chkListing As New CheckBox With {.Text = UiText.T("dlg.asm.listing"), .Checked = True}
    Private ReadOnly _cmbBaud As New DropDown With {.Width = 130}
    Private ReadOnly _txtResult As TextArea
    Private ReadOnly _btnRun As New Button With {.Text = UiText.T("btn.assemble")}
    Private ReadOnly _btnClose As New Button With {.Text = UiText.T("btn.close")}
    Private _outputPath As String = Nothing

    Public Sub New()
        Title = UiText.T("dlg.asm.title")
        Resizable = True
        ClientSize = New Size(720, 520)
        MinimumSize = New Size(600, 420)
        Padding = New Padding(Ui.Margin)
        _rbM3 = New RadioButton With {.Text = UiText.T("dlg.asm.m3"), .Checked = True}
        _rbM1 = New RadioButton(_rbM3) With {.Text = UiText.T("dlg.asm.m1")}
        _rbM4 = New RadioButton(_rbM3) With {.Text = UiText.T("dlg.asm.m4")}
        ' Only item 0 is text; the rest are the numbers trs80-tool is given.
        For Each item In New String() {UiText.T("dlg.asm.bauddefault"), "250", "500", "1500"}
            _cmbBaud.Items.Add(item)
        Next
        _cmbBaud.SelectedIndex = 0
        _txtResult = Ui.ReportBox("", 9.0F)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        AddHandler _btnClose.Click, Sub() Close()
        Dim srcRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        srcRow.Rows.Add(New TableRow(New TableCell(_txtSrc, True), New TableCell(Ui.Btn(UiText.T("btn.browse"), AddressOf BrowseSource))))
        Dim outRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        outRow.Rows.Add(New TableRow(New TableCell(_txtOut, True), New TableCell(Ui.Btn(UiText.T("btn.browse"), AddressOf BrowseOutput))))
        Content = Ui.Down(_txtResult,
                          Ui.Lbl(UiText.T("dlg.asm.source")), srcRow,
                          Ui.Lbl(UiText.T("dlg.asm.output")), outRow,
                          Ui.Across(Ui.Group(UiText.T("dlg.asm.target"), Ui.Across(_rbM1, _rbM3, _rbM4)),
                                    New Label With {.Text = "    " & UiText.T("dlg.asm.baud"), .VerticalAlignment = VerticalAlignment.Center}, _cmbBaud),
                          _chkListing,
                          Ui.Lbl(UiText.T("msg.engineoutput")),
                          _txtResult,
                          Ui.Across(Nothing, _btnRun, _btnClose))
        AbortButton = _btnClose
        UpdateEnabled()
    End Sub

    Public ReadOnly Property OutputPath As String
        Get
            Return _outputPath
        End Get
    End Property

    Private Sub UpdateEnabled()
        Dim ext As String = Path.GetExtension(_txtOut.Text.Trim()).ToLowerInvariant()
        _cmbBaud.Enabled = (ext = ".cas" OrElse ext = ".wav")
    End Sub

    Private Sub BrowseSource()
        Dim chosen As String = Ui.PickFile(Me, UiText.T("dlg.asm.picksource"), UiText.T("filter.asmsource"))
        If chosen Is Nothing Then Return
        _txtSrc.Text = chosen
        If _txtOut.Text.Trim().Length = 0 Then _txtOut.Text = Path.ChangeExtension(chosen, ".cmd")
    End Sub

    Private Sub BrowseOutput()
        Dim chosen As String = Ui.PickSave(Me, UiText.T("dlg.asm.pickoutput"), UiText.T("filter.asmoutput"),
                                           If(_txtOut.Text.Trim().Length > 0, _txtOut.Text.Trim(), Nothing), "cmd")
        If chosen IsNot Nothing Then _txtOut.Text = chosen
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String) From {"--model", If(_rbM1.Checked, "1", If(_rbM4.Checked, "4", "3"))}
        If Ui.IsOn(_chkListing) AndAlso _txtOut.Text.Trim().Length > 0 Then
            w.Add("--listing") : w.Add(Path.ChangeExtension(_txtOut.Text.Trim(), ".lst"))
        End If
        If _cmbBaud.Enabled AndAlso _cmbBaud.SelectedIndex > 0 Then
            w.Add("--baud") : w.Add(_cmbBaud.Items(_cmbBaud.SelectedIndex).Text)
        End If
        Return w
    End Function

    Private Async Function RunAsync() As Task
        Dim src As String = _txtSrc.Text.Trim()
        Dim outFile As String = _txtOut.Text.Trim()
        If src.Length = 0 OrElse Not File.Exists(src) Then
            Ui.Warn(Me, UiText.T("dlg.asm.nosource"), UiText.T("btn.assemble"))
            Return
        End If
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            Ui.Warn(Me, UiText.T("dlg.asm.nooutput"), UiText.T("btn.assemble"))
            Return
        End If
        _btnRun.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.AsmAsync(src, outFile, OptionWords())
            _txtResult.Text = Ui.Lf(r.CommandLine & vbLf & vbLf & If(r.Succeeded, If(r.StdOut.Trim().Length = 0, UiText.T("dlg.asm.nomessages"), r.StdOut), r.ErrorText))
            If r.Succeeded AndAlso File.Exists(outFile) Then _outputPath = outFile
        Catch ex As Exception
            _txtResult.Text = ex.Message
        Finally
            _btnRun.Enabled = True
        End Try
    End Function
End Class
