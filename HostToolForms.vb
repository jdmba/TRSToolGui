' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HostToolForms.vb — Tools menu windows that work on host files through
'  trs80-tool: Convert file(s) ("convert IN... OUT")
'  and Assemble ("asm SRC OUT"). Each builds the option words, runs the tool,
'  shows what it printed, and tells MainForm the output path so it can offer
'  to add the result to the open disk.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

' ---------------------------------------------------------------- Convert file(s)

''' <summary>
''' trs80-tool convert: BAS↔ASC, CMD↔3BN, programs → CAS/WAV (--baud), ROM/BIN → CMD/3BN (--start), EDTASM → ASC,
''' several inputs → one CAS/WAV, floppy image → DMK/JV1/JV3/HFE/IMD (--force). The output extension decides.
''' </summary>
Public Class ConvertFileForm
    Inherits Form

    Private ReadOnly _lstIn As New ListBox()
    Private ReadOnly _txtOut As New TextBox()
    Private ReadOnly _cmbBaud As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 120}
    Private ReadOnly _txtStart As New TextBox() With {.Width = 120}
    Private ReadOnly _chkSystem As New CheckBox() With {.Text = "Include system files when the input is a disk image (--system)", .AutoSize = True}
    Private ReadOnly _chkCpm As New CheckBox() With {.Text = "CP/M: keep whole 128-byte records (--cpm-records)", .AutoSize = True}
    Private ReadOnly _chkForce As New CheckBox() With {.Text = "Floppy images: write even if the new format can't hold everything (--force)", .AutoSize = True}
    Private ReadOnly _txtResult As New TextBox()
    Private ReadOnly _btnRun As New Button()
    Private ReadOnly _btnClose As New Button()
    Private _outputPath As String = Nothing

    ''' <param name="initialInputs">Files to start with (e.g. one staged from the open disk), or Nothing.</param>
    Public Sub New(initialInputs As IEnumerable(Of String))
        Text = "Convert file(s) (trs80-tool convert)"
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(720, 560)
        MinimumSize = New Size(600, 480)
        ShowInTaskbar = False
        Dim y As Integer = 10
        Controls.Add(New Label() With {.Text = "Input file(s) — several inputs are combined into one CAS/WAV output:", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        _lstIn.Bounds = New Rectangle(12, y, 560, 90)
        _lstIn.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        _lstIn.HorizontalScrollbar = True
        Controls.Add(_lstIn)
        Dim btnAdd As New Button() With {.Text = "Add...", .Bounds = New Rectangle(584, y, 100, 26), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        Dim btnRemove As New Button() With {.Text = "Remove", .Bounds = New Rectangle(584, y + 32, 100, 26), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnAdd.Click, Sub() AddInputs()
        AddHandler btnRemove.Click, Sub()
                                        If _lstIn.SelectedIndex >= 0 Then _lstIn.Items.RemoveAt(_lstIn.SelectedIndex)
                                    End Sub
        Controls.Add(btnAdd) : Controls.Add(btnRemove)
        If initialInputs IsNot Nothing Then
            For Each f In initialInputs
                _lstIn.Items.Add(f)
            Next
        End If
        y += 100
        Controls.Add(New Label() With {.Text = "Output file (the extension picks the conversion: .bas .asc .cmd .3bn .cas .wav .dmk .jv1 .jv3 .hfe .imd):", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        _txtOut.Bounds = New Rectangle(12, y, 560, 23)
        _txtOut.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Controls.Add(_txtOut)
        Dim btnOut As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(584, y - 1, 100, 26), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnOut.Click, Sub() BrowseOutput()
        Controls.Add(btnOut)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        y += 34
        Controls.Add(New Label() With {.Text = "Cassette baud (--baud, CAS/WAV):", .AutoSize = True, .Location = New Point(12, y + 3)})
        _cmbBaud.Items.AddRange(New Object() {"default", "250", "500", "1500"})
        _cmbBaud.SelectedIndex = 0
        _cmbBaud.Location = New Point(230, y)
        Controls.Add(_cmbBaud)
        Controls.Add(New Label() With {.Text = "Start address (--start, hex or auto; ROM/BIN → CMD/3BN):", .AutoSize = True, .Location = New Point(370, y + 3)})
        y += 28
        _txtStart.Location = New Point(370, y)
        Controls.Add(_txtStart)
        y += 32
        _chkSystem.Location = New Point(12, y) : y += 24
        _chkCpm.Location = New Point(12, y) : y += 24
        _chkForce.Location = New Point(12, y) : y += 30
        Controls.AddRange(New Control() {_chkSystem, _chkCpm, _chkForce})
        Controls.Add(New Label() With {.Text = "trs80-tool output:", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        With _txtResult
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
            .Bounds = New Rectangle(12, y, 672, 110)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .BackColor = SystemColors.Window
        End With
        Controls.Add(_txtResult)
        _btnRun.Text = "Convert"
        _btnRun.Bounds = New Rectangle(504, 486, 90, 26)
        _btnRun.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        _btnClose.Text = "Close"
        _btnClose.DialogResult = DialogResult.Cancel
        _btnClose.Bounds = New Rectangle(604, 486, 80, 26)
        _btnClose.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Controls.Add(_btnRun) : Controls.Add(_btnClose)
        CancelButton = _btnClose
        UpdateEnabled()
    End Sub

    ''' <summary>The last output written successfully, or Nothing.</summary>
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
            ofd.Filter = "TRS-80 files (*.bas;*.asc;*.cmd;*.3bn;*.cas;*.wav;*.rom;*.bin;*.hex;*.asm;*.dmk;*.jv1;*.jv3;*.dsk;*.hfe;*.imd;*.scp)|*.bas;*.asc;*.cmd;*.3bn;*.cas;*.wav;*.rom;*.bin;*.hex;*.asm;*.dmk;*.jv1;*.jv3;*.dsk;*.hfe;*.imd;*.scp|All files (*.*)|*.*"
            ofd.Multiselect = True
            If ofd.ShowDialog(Me) = DialogResult.OK Then
                For Each f In ofd.FileNames
                    If Not _lstIn.Items.Contains(f) Then _lstIn.Items.Add(f)
                Next
                If _txtOut.Text.Trim().Length = 0 AndAlso _lstIn.Items.Count = 1 Then
                    ' Suggest the natural partner: BAS→ASC, ASC→BAS, CMD→3BN, 3BN→CMD, CAS→CMD, WAV→CAS, image→DMK.
                    Dim first As String = _lstIn.Items(0).ToString()
                    Dim e As String = Path.GetExtension(first).ToLowerInvariant()
                    Dim suggest As String = Nothing
                    Select Case e
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
            End If
        End Using
    End Sub

    Private Sub BrowseOutput()
        Using sfd As New SaveFileDialog()
            sfd.Title = "Output file"
            sfd.Filter = "BASIC tokenized (*.bas)|*.bas|BASIC text (*.asc)|*.asc|CMD program (*.cmd)|*.cmd|Cassette system program (*.3bn)|*.3bn|Cassette image (*.cas)|*.cas|Cassette audio (*.wav)|*.wav|DMK floppy image (*.dmk)|*.dmk|JV3 floppy image (*.jv3)|*.jv3|JV1 floppy image (*.jv1)|*.jv1|HFE floppy image (*.hfe)|*.hfe|IMD floppy image (*.imd)|*.imd|All files (*.*)|*.*"
            sfd.OverwritePrompt = True
            If _txtOut.Text.Trim().Length > 0 Then
                Try
                    sfd.InitialDirectory = Path.GetDirectoryName(_txtOut.Text.Trim())
                    sfd.FileName = Path.GetFileName(_txtOut.Text.Trim())
                Catch
                End Try
            End If
            If sfd.ShowDialog(Me) = DialogResult.OK Then _txtOut.Text = sfd.FileName
        End Using
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _cmbBaud.Enabled AndAlso _cmbBaud.SelectedIndex > 0 Then
            w.Add("--baud") : w.Add(_cmbBaud.SelectedItem.ToString())
        End If
        If _txtStart.Enabled AndAlso _txtStart.Text.Trim().Length > 0 Then
            w.Add("--start") : w.Add(_txtStart.Text.Trim())
        End If
        If _chkSystem.Checked Then w.Add("--system")
        If _chkCpm.Checked Then w.Add("--cpm-records")
        If _chkForce.Enabled AndAlso _chkForce.Checked Then w.Add("--force")
        Return w
    End Function

    Private Async Function RunAsync() As Task
        If _lstIn.Items.Count = 0 Then
            MessageBox.Show(Me, "Add at least one input file.", "Convert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim outFile As String = _txtOut.Text.Trim()
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            MessageBox.Show(Me, "Give an output file with an extension; trs80-tool converts by the output's extension.", "Convert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim inputs As New List(Of String)
        For Each o In _lstIn.Items
            inputs.Add(o.ToString())
        Next
        _btnRun.Enabled = False
        UseWaitCursor = True
        Try
            Dim r As ToolResult = Await Program.Tool.ConvertFilesAsync(inputs, outFile, OptionWords())
            _txtResult.Text = (r.CommandLine & vbCrLf & vbCrLf & If(r.Succeeded, r.StdOut, r.ErrorText)).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            If r.Succeeded AndAlso File.Exists(outFile) Then _outputPath = outFile
        Catch ex As Exception
            _txtResult.Text = ex.Message
        Finally
            UseWaitCursor = False
            _btnRun.Enabled = True
        End Try
    End Function
End Class

' ---------------------------------------------------------------- Assemble

''' <summary>trs80-tool asm SRC OUT [--model 1|3|4] [--listing FILE] [--baud N]. OUT's extension picks CMD / 3BN / CAS / WAV.</summary>
Public Class AssembleForm
    Inherits Form

    Private ReadOnly _txtSrc As New TextBox()
    Private ReadOnly _txtOut As New TextBox()
    Private ReadOnly _rbM1 As New RadioButton() With {.Text = "Model I", .AutoSize = True}
    Private ReadOnly _rbM3 As New RadioButton() With {.Text = "Model III (default)", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbM4 As New RadioButton() With {.Text = "Model 4", .AutoSize = True}
    Private ReadOnly _chkListing As New CheckBox() With {.Text = "Write a listing file next to the output (--listing OUT.lst)", .AutoSize = True, .Checked = True}
    Private ReadOnly _cmbBaud As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 120}
    Private ReadOnly _txtResult As New TextBox()
    Private ReadOnly _btnRun As New Button()
    Private ReadOnly _btnClose As New Button()
    Private _outputPath As String = Nothing

    Public Sub New()
        Text = "Assemble (trs80-tool asm)"
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(720, 500)
        MinimumSize = New Size(600, 420)
        ShowInTaskbar = False
        Dim y As Integer = 10
        Controls.Add(New Label() With {.Text = "Z80 source file (.asm):", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        _txtSrc.Bounds = New Rectangle(12, y, 560, 23) : _txtSrc.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Controls.Add(_txtSrc)
        Dim btnSrc As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(584, y - 1, 100, 26), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnSrc.Click, Sub() BrowseSource()
        Controls.Add(btnSrc)
        y += 34
        Controls.Add(New Label() With {.Text = "Output file (.cmd, .3bn, .cas or .wav):", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        _txtOut.Bounds = New Rectangle(12, y, 560, 23) : _txtOut.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Controls.Add(_txtOut)
        Dim btnOut As New Button() With {.Text = "Browse...", .Bounds = New Rectangle(584, y - 1, 100, 26), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        AddHandler btnOut.Click, Sub() BrowseOutput()
        Controls.Add(btnOut)
        AddHandler _txtOut.TextChanged, Sub() UpdateEnabled()
        y += 34
        Dim grp As New GroupBox() With {.Text = "Target (--model; defines __MODEL__ for the source)", .Bounds = New Rectangle(12, y, 400, 50)}
        _rbM1.Location = New Point(14, 20) : _rbM3.Location = New Point(100, 20) : _rbM4.Location = New Point(250, 20)
        grp.Controls.AddRange(New Control() {_rbM1, _rbM3, _rbM4})
        Controls.Add(grp)
        Controls.Add(New Label() With {.Text = "Cassette baud (--baud):", .AutoSize = True, .Location = New Point(430, y + 20)})
        _cmbBaud.Items.AddRange(New Object() {"default (500)", "250", "500", "1500"})
        _cmbBaud.SelectedIndex = 0
        _cmbBaud.Location = New Point(570, y + 16)
        Controls.Add(_cmbBaud)
        y += 60
        _chkListing.Location = New Point(12, y)
        Controls.Add(_chkListing)
        y += 30
        Controls.Add(New Label() With {.Text = "trs80-tool output:", .AutoSize = True, .Location = New Point(12, y)})
        y += 20
        With _txtResult
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
            .Bounds = New Rectangle(12, y, 672, 150)
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .BackColor = SystemColors.Window
        End With
        Controls.Add(_txtResult)
        _btnRun.Text = "Assemble"
        _btnRun.Bounds = New Rectangle(494, 426, 100, 26)
        _btnRun.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        AddHandler _btnRun.Click, Async Sub() Await RunAsync()
        _btnClose.Text = "Close"
        _btnClose.DialogResult = DialogResult.Cancel
        _btnClose.Bounds = New Rectangle(604, 426, 80, 26)
        _btnClose.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Controls.Add(_btnRun) : Controls.Add(_btnClose)
        CancelButton = _btnClose
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
        Using ofd As New OpenFileDialog()
            ofd.Title = "Z80 source"
            ofd.Filter = "Assembler source (*.asm;*.z80;*.src;*.txt)|*.asm;*.z80;*.src;*.txt|All files (*.*)|*.*"
            If ofd.ShowDialog(Me) = DialogResult.OK Then
                _txtSrc.Text = ofd.FileName
                If _txtOut.Text.Trim().Length = 0 Then _txtOut.Text = Path.ChangeExtension(ofd.FileName, ".cmd")
            End If
        End Using
    End Sub

    Private Sub BrowseOutput()
        Using sfd As New SaveFileDialog()
            sfd.Title = "Assembled output"
            sfd.Filter = "CMD program (*.cmd)|*.cmd|Cassette system program (*.3bn)|*.3bn|Cassette image (*.cas)|*.cas|Cassette audio (*.wav)|*.wav"
            sfd.OverwritePrompt = True
            If _txtOut.Text.Trim().Length > 0 Then
                Try
                    sfd.InitialDirectory = Path.GetDirectoryName(_txtOut.Text.Trim())
                    sfd.FileName = Path.GetFileName(_txtOut.Text.Trim())
                Catch
                End Try
            End If
            If sfd.ShowDialog(Me) = DialogResult.OK Then _txtOut.Text = sfd.FileName
        End Using
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String) From {"--model", If(_rbM1.Checked, "1", If(_rbM4.Checked, "4", "3"))}
        If _chkListing.Checked AndAlso _txtOut.Text.Trim().Length > 0 Then
            w.Add("--listing") : w.Add(Path.ChangeExtension(_txtOut.Text.Trim(), ".lst"))
        End If
        If _cmbBaud.Enabled AndAlso _cmbBaud.SelectedIndex > 0 Then
            w.Add("--baud") : w.Add(_cmbBaud.SelectedItem.ToString())
        End If
        Return w
    End Function

    Private Async Function RunAsync() As Task
        Dim src As String = _txtSrc.Text.Trim()
        Dim outFile As String = _txtOut.Text.Trim()
        If src.Length = 0 OrElse Not File.Exists(src) Then
            MessageBox.Show(Me, "Choose an existing source file.", "Assemble", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If outFile.Length = 0 OrElse Path.GetExtension(outFile).Length = 0 Then
            MessageBox.Show(Me, "Give an output file with an extension (.cmd, .3bn, .cas or .wav).", "Assemble", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _btnRun.Enabled = False
        UseWaitCursor = True
        Try
            Dim r As ToolResult = Await Program.Tool.AsmAsync(src, outFile, OptionWords())
            _txtResult.Text = (r.CommandLine & vbCrLf & vbCrLf & If(r.Succeeded, If(r.StdOut.Trim().Length = 0, "Assembled without messages.", r.StdOut), r.ErrorText)).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            If r.Succeeded AndAlso File.Exists(outFile) Then _outputPath = outFile
        Catch ex As Exception
            _txtResult.Text = ex.Message
        Finally
            UseWaitCursor = False
            _btnRun.Enabled = True
        End Try
    End Function
End Class
