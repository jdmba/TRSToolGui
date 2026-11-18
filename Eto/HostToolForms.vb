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
    Private ReadOnly _chkSystem As New CheckBox With {.Text = "Include system files when the input is a disk image (--system)"}
    Private ReadOnly _chkCpm As New CheckBox With {.Text = "CP/M: keep whole 128-byte records (--cpm-records)"}
    Private ReadOnly _chkForce As New CheckBox With {.Text = "Floppy images: write even if the new format can't hold everything (--force)"}
    Private ReadOnly _txtResult As TextArea
    Private ReadOnly _btnRun As New Button With {.Text = "Convert"}
    Private ReadOnly _btnClose As New Button With {.Text = "Close"}
    Private _outputPath As String = Nothing

    Public Sub New(initialInputs As IEnumerable(Of String))
        Title = "Convert file(s) (trs80-tool convert)"
        Resizable = True
        ClientSize = New Size(720, 580)
        MinimumSize = New Size(600, 480)
        Padding = New Padding(Ui.Margin)
        If initialInputs IsNot Nothing Then
            For Each f In initialInputs
                _lstIn.Items.Add(f)
            Next
        End If
        For Each item In New String() {"default", "250", "500", "1500"}
            _cmbBaud.Items.Add(item)
        Next
        _cmbBaud.SelectedIndex = 0
        _txtResult = Ui.ReportBox("", 9.0F)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        AddHandler _btnClose.Click, Sub() Close()
        Dim inButtons As TableLayout = Ui.Down(Nothing, Ui.Btn("Add...", AddressOf AddInputs), Ui.Btn("Remove", Sub()
                                                                                                                  If _lstIn.SelectedIndex >= 0 Then _lstIn.Items.RemoveAt(_lstIn.SelectedIndex)
                                                                                                              End Sub))
        Dim inRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        inRow.Rows.Add(New TableRow(New TableCell(_lstIn, True), New TableCell(inButtons)))
        Dim outRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        outRow.Rows.Add(New TableRow(New TableCell(_txtOut, True), New TableCell(Ui.Btn("Browse...", AddressOf BrowseOutput))))
        Content = Ui.Down(_txtResult,
                          Ui.Lbl("Input file(s) — several inputs are combined into one CAS/WAV output:"),
                          inRow,
                          Ui.Lbl("Output file (the extension picks the conversion: .bas .asc .cmd .3bn .cas .wav .dmk .jv1 .jv3 .hfe .imd):"),
                          outRow,
                          Ui.Fields("Cassette baud (--baud, CAS/WAV):", Ui.Across(_cmbBaud),
                                    "Start address (--start, hex or auto; ROM/BIN → CMD/3BN):", Ui.Across(_txtStart)),
                          _chkSystem, _chkCpm, _chkForce,
                          Ui.Lbl("trs80-tool output:"),
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
            ofd.Title = "Input file(s) to convert"
            Ui.SetFilters(ofd, "TRS-80 files (*.bas;*.asc;*.cmd;*.3bn;*.cas;*.wav;*.rom;*.bin;*.hex;*.asm;*.dmk;*.jv1;*.jv3;*.dsk;*.hfe;*.imd;*.scp)|*.bas;*.asc;*.cmd;*.3bn;*.cas;*.wav;*.rom;*.bin;*.hex;*.asm;*.dmk;*.jv1;*.jv3;*.dsk;*.hfe;*.imd;*.scp|All files (*.*)|*.*")
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
        Dim chosen As String = Ui.PickSave(Me, "Output file", "BASIC tokenized (*.bas)|*.bas|BASIC text (*.asc)|*.asc|CMD program (*.cmd)|*.cmd|Cassette system program (*.3bn)|*.3bn|Cassette image (*.cas)|*.cas|Cassette audio (*.wav)|*.wav|DMK floppy image (*.dmk)|*.dmk|JV3 floppy image (*.jv3)|*.jv3|JV1 floppy image (*.jv1)|*.jv1|HFE floppy image (*.hfe)|*.hfe|IMD floppy image (*.imd)|*.imd|All files (*.*)|*.*",
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
            Ui.Warn(Me, "Add at least one input file.", "Convert")
            Return
        End If
        Dim outFile As String = _txtOut.Text.Trim()
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            Ui.Warn(Me, "Give an output file with an extension; trs80-tool converts by the output's extension.", "Convert")
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
    Private ReadOnly _chkListing As New CheckBox With {.Text = "Write a listing file next to the output (--listing OUT.lst)", .Checked = True}
    Private ReadOnly _cmbBaud As New DropDown With {.Width = 130}
    Private ReadOnly _txtResult As TextArea
    Private ReadOnly _btnRun As New Button With {.Text = "Assemble"}
    Private ReadOnly _btnClose As New Button With {.Text = "Close"}
    Private _outputPath As String = Nothing

    Public Sub New()
        Title = "Assemble (trs80-tool asm)"
        Resizable = True
        ClientSize = New Size(720, 520)
        MinimumSize = New Size(600, 420)
        Padding = New Padding(Ui.Margin)
        _rbM3 = New RadioButton With {.Text = "Model III (default)", .Checked = True}
        _rbM1 = New RadioButton(_rbM3) With {.Text = "Model I"}
        _rbM4 = New RadioButton(_rbM3) With {.Text = "Model 4"}
        For Each item In New String() {"default (500)", "250", "500", "1500"}
            _cmbBaud.Items.Add(item)
        Next
        _cmbBaud.SelectedIndex = 0
        _txtResult = Ui.ReportBox("", 9.0F)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        AddHandler _btnClose.Click, Sub() Close()
        Dim srcRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        srcRow.Rows.Add(New TableRow(New TableCell(_txtSrc, True), New TableCell(Ui.Btn("Browse...", AddressOf BrowseSource))))
        Dim outRow As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        outRow.Rows.Add(New TableRow(New TableCell(_txtOut, True), New TableCell(Ui.Btn("Browse...", AddressOf BrowseOutput))))
        Content = Ui.Down(_txtResult,
                          Ui.Lbl("Z80 source file (.asm):"), srcRow,
                          Ui.Lbl("Output file (.cmd, .3bn, .cas or .wav):"), outRow,
                          Ui.Across(Ui.Group("Target (--model; defines __MODEL__ for the source)", Ui.Across(_rbM1, _rbM3, _rbM4)),
                                    New Label With {.Text = "    Cassette baud (--baud):", .VerticalAlignment = VerticalAlignment.Center}, _cmbBaud),
                          _chkListing,
                          Ui.Lbl("trs80-tool output:"),
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
        Dim chosen As String = Ui.PickFile(Me, "Z80 source", "Assembler source (*.asm;*.z80;*.src;*.txt)|*.asm;*.z80;*.src;*.txt|All files (*.*)|*.*")
        If chosen Is Nothing Then Return
        _txtSrc.Text = chosen
        If _txtOut.Text.Trim().Length = 0 Then _txtOut.Text = Path.ChangeExtension(chosen, ".cmd")
    End Sub

    Private Sub BrowseOutput()
        Dim chosen As String = Ui.PickSave(Me, "Assembled output", "CMD program (*.cmd)|*.cmd|Cassette system program (*.3bn)|*.3bn|Cassette image (*.cas)|*.cas|Cassette audio (*.wav)|*.wav",
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
            Ui.Warn(Me, "Choose an existing source file.", "Assemble")
            Return
        End If
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            Ui.Warn(Me, "Give an output file with an extension (.cmd, .3bn, .cas or .wav).", "Assemble")
            Return
        End If
        _btnRun.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.AsmAsync(src, outFile, OptionWords())
            _txtResult.Text = Ui.Lf(r.CommandLine & vbLf & vbLf & If(r.Succeeded, If(r.StdOut.Trim().Length = 0, "Assembled without messages.", r.StdOut), r.ErrorText))
            If r.Succeeded AndAlso File.Exists(outFile) Then _outputPath = outFile
        Catch ex As Exception
            _txtResult.Text = ex.Message
        Finally
            _btnRun.Enabled = True
        End Try
    End Function
End Class
