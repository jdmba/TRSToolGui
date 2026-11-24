' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  ViewerForm.vb — the File Viewer of the Eto.Forms program, modelled on
'  TRSTool's: a view chooser (Graphics / EDTASM / BASIC / Disassembly / Text /
'  Hexadecimal), a font choice for text, Options, Save as and Done. All
'  decoding is done by trs80-tool: BASIC and hex by "disk list", disassembly
'  by "disasm", EDTASM by "convert" to .asc, hi-res pictures by "convert" to
'  PNG. The two option dialogs are in this file too.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class ViewerForm
    Inherits ToolWindow

    Private Enum ViewerKind
        Graphics
        Edtasm
        Basic
        Disassembly
        Text
        Hex
    End Enum

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _file As DiskFile
    Private _bytes As Byte() = Nothing

    Private ReadOnly _text As New TextArea With {.ReadOnly = True, .Wrap = False}
    Private ReadOnly _pic As New ImageView()
    Private ReadOnly _picScroll As New Scrollable With {.BackgroundColor = Colors.Black}
    Private ReadOnly _view As New Panel()
    ''' <summary>The PNG trs80-tool made of a hi-res file ("convert NAME.HR NAME.png"); Save as writes these bytes.</summary>
    Private _pngBytes As Byte() = Nothing
    Private ReadOnly _combo As New DropDown With {.Width = 150}
    Private ReadOnly _rbMono As RadioButton
    Private ReadOnly _rbProp As RadioButton
    Private ReadOnly _btnOptions As New Button With {.Text = UiText.T("btn.options"), .Enabled = False}
    ''' <summary>Per-viewer disassembly words, starting from the Preferences default.</summary>
    Private _disasmWords As List(Of String) = Program.Settings.DisasmOptionWords()
    Private _basicDialect As String = ""          ' "" = let the DOS decide, "3" or "4" = --basic N
    Private ReadOnly _lblInfo As New Label With {.TextColor = Ui.GrayText, .VerticalAlignment = VerticalAlignment.Center}
    ''' <summary>The layout and line count of an assembler source file; Nothing for any other file.</summary>
    Private _asm As AsmSourceInfo

    Private ReadOnly _kinds As New List(Of ViewerKind)
    Private _current As ViewerKind = ViewerKind.Hex
    Private _loading As Boolean = False
    Private ReadOnly _cache As New Dictionary(Of ViewerKind, String)

    Private Shared ReadOnly MonoFont As Font = Ui.Mono(10.0F)
    Private Shared ReadOnly PropFont As Font = Fonts.Serif(11.0F)

    Public Sub New(disk As DiskImage, file As DiskFile)
        _disk = disk
        _file = file
        _rbMono = New RadioButton With {.Text = UiText.T("dlg.viewer.mono")}
        _rbProp = New RadioButton(_rbMono) With {.Text = UiText.T("dlg.viewer.prop")}
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Title = UiText.T("dlg.viewer.title", _file.SlashName)
        ClientSize = New Size(820, 560)
        MinimumSize = New Size(420, 260)
        Padding = New Padding(Ui.Margin)
        _text.Font = MonoFont
        _picScroll.Content = _pic
        _view.Content = _text
        AddHandler _rbMono.CheckedChanged, Sub() ApplyFont()
        AddHandler _rbProp.CheckedChanged, Sub() ApplyFont()
        AddHandler _combo.SelectedIndexChanged, Async Sub() Await ShowKindAsync()
        AddHandler _btnOptions.Click, Async Sub() Await EditOptionsAsync()
        Dim bottom As TableLayout = Ui.Across(Ui.Group(UiText.T("dlg.viewer.font"), Ui.Across(_rbMono, _rbProp)), _combo, _btnOptions, _lblInfo, Nothing, Ui.Btn(UiText.T("btn.saveas"), AddressOf SaveCurrent), Ui.Btn(UiText.T("btn.done"), Sub() Close()))
        Content = Ui.Down(_view, _view, bottom)
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        _rbMono.Checked = Program.Settings.ViewerMonospaced
        _rbProp.Checked = Not _rbMono.Checked
        _text.Text = UiText.T("dlg.viewer.working")
        Try
            _bytes = Await Program.Tool.ReadFileBytesAsync(_disk.Path, _disk.PartitionArg, _file.TrsName)
        Catch ex As Exception
            _bytes = Nothing
            Ui.Fail(Me, UiText.T("dlg.viewer.readfailed", _file.SlashName, ex.Message), UiText.T("caption.fileviewer"))
        End Try
        If _bytes Is Nothing Then
            Close()
            Return
        End If
        _lblInfo.Text = UiText.T("word.bytes", TextUtil.FormatBytes(_bytes.LongLength)) & If(_file.Description.Length > 0, "  •  " & _file.Description, "")
        DetectKinds()
        ' Assembler source: say which layout it is in and where it really ends (old slack past that is not shown).
        If _kinds(0) = ViewerKind.Edtasm OrElse (_kinds(0) = ViewerKind.Text AndAlso AsmSource.IsSourceExtension(_file.Extension)) Then
            _asm = AsmSource.Analyze(_bytes)
            _lblInfo.Text &= "  •  " & _asm.Summary()
        End If
        Dim basicKind As String = BasicKindLabel()
        If basicKind.Length > 0 Then Title = UiText.T("dlg.viewer.titlekind", UiText.T("dlg.viewer.title", _file.SlashName), basicKind)
        _loading = True
        _combo.Items.Clear()
        For Each k In _kinds
            _combo.Items.Add(KindName(k))
        Next
        _loading = False
        _combo.SelectedIndex = 0
    End Sub

    Protected Overrides Sub OnClosed(e As EventArgs)
        Program.Settings.ViewerMonospaced = _rbMono.Checked
        MyBase.OnClosed(e)
    End Sub

    Private Shared Function KindName(k As ViewerKind) As String
        Select Case k
            Case ViewerKind.Graphics : Return UiText.T("dlg.viewer.kind.graphics")
            Case ViewerKind.Edtasm : Return UiText.T("dlg.viewer.kind.edtasm")
            Case ViewerKind.Basic : Return UiText.T("dlg.viewer.kind.basic")
            Case ViewerKind.Disassembly : Return UiText.T("dlg.viewer.kind.disassembly")
            Case ViewerKind.Text : Return UiText.T("dlg.viewer.kind.text")
            Case Else : Return UiText.T("dlg.viewer.kind.hex")
        End Select
    End Function

    ''' <summary>
    ''' The same kinds as a tag for the file name the Save dialog suggests. These stay
    ''' English whatever language the windows are in: a translated name would put accents
    ''' and spaces into a file name on the host.
    ''' </summary>
    Private Shared Function KindFileTag(k As ViewerKind) As String
        Select Case k
            Case ViewerKind.Graphics : Return "Graphics"
            Case ViewerKind.Edtasm : Return "EDTASM"
            Case ViewerKind.Basic : Return "BASIC"
            Case ViewerKind.Disassembly : Return "Disassembly"
            Case ViewerKind.Text : Return "Text"
            Case Else : Return "Hexadecimal"
        End Select
    End Function

    ''' <summary>
    ''' TRSTool's detection order — EDTASM, tokenized BASIC, /CMD, text, else hex — with the engine's contentType
    ''' as a second opinion for /CMD and EDTASM and the only word on hi-res pictures (HR, CHR, SHR).
    ''' </summary>
    Private Sub DetectKinds()
        _kinds.Clear()
        Dim trsdos As Boolean = _disk.IsTrsdosFamily
        Dim ct As String = _file.ContentType
        If trsdos AndAlso ct = "hiResScreen" Then
            _kinds.Add(ViewerKind.Graphics) : _kinds.Add(ViewerKind.Hex)
        ElseIf trsdos AndAlso (ct = "edtasm" OrElse FileKinds.IsEdtasm(_bytes)) Then
            _kinds.Add(ViewerKind.Edtasm) : _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        ElseIf (trsdos OrElse _disk.DosName.Length = 0) AndAlso FileKinds.IsTokenizedBasic(_bytes) Then
            ' Also on a disk whose DOS trs80-tool doesn't recognise: a tokenized program doesn't depend on the DOS.
            _kinds.Add(ViewerKind.Basic) : _kinds.Add(ViewerKind.Hex)
        ElseIf trsdos AndAlso (ct = "cmdProgram" OrElse FileKinds.IsCmd(_bytes)) Then
            _kinds.Add(ViewerKind.Disassembly) : _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        ElseIf TextUtil.LooksLikeText(_file.Extension, _bytes) Then
            _kinds.Add(ViewerKind.Text) : _kinds.Add(ViewerKind.Hex)
        Else
            _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        End If
    End Sub

    ''' <summary>The kind of BASIC program for the title bar; empty when the file isn't BASIC.</summary>
    Private Function BasicKindLabel() As String
        If _kinds.Contains(ViewerKind.Basic) Then
            Return If(FileKinds.IsModel4Basic(_bytes), UiText.T("dlg.viewer.basic.m4"), UiText.T("dlg.viewer.basic.m13"))
        End If
        If _kinds.Count > 0 AndAlso _kinds(0) = ViewerKind.Text AndAlso FileKinds.IsAsciiBasic(_bytes) Then Return UiText.T("dlg.viewer.basic.ascii")
        Return ""
    End Function

    Private Async Function ShowKindAsync() As Task
        If _loading OrElse _combo.SelectedIndex < 0 OrElse _combo.SelectedIndex >= _kinds.Count Then Return
        _current = _kinds(_combo.SelectedIndex)
        _rbMono.Enabled = (_current = ViewerKind.Text)
        _rbProp.Enabled = (_current = ViewerKind.Text)
        _btnOptions.Enabled = (_current = ViewerKind.Disassembly OrElse _current = ViewerKind.Basic)

        If _current = ViewerKind.Graphics Then
            _view.Content = _picScroll
            If _pic.Image Is Nothing Then
                Try
                    _pic.Image = Await GraphicsAsync()
                Catch ex As Exception
                    _pic.Image = Nothing
                    Ui.Fail(Me, UiText.T("dlg.viewer.picfailed", ex.Message), UiText.T("caption.fileviewer"))
                End Try
            End If
            Return
        End If
        _view.Content = _text

        Dim content As String = Nothing
        If _cache.TryGetValue(_current, content) Then
            SetText(content)
            Return
        End If
        _text.Text = UiText.T("dlg.viewer.working")
        Try
            Select Case _current
                Case ViewerKind.Basic
                    content = Await Program.Tool.ListTextAsync(_disk.Path, _disk.PartitionArg, _file.TrsName, False, If(_basicDialect.Length = 0, Nothing, _basicDialect))
                Case ViewerKind.Hex
                    Try
                        content = Await Program.Tool.ListTextAsync(_disk.Path, _disk.PartitionArg, _file.TrsName, True, Nothing)
                    Catch
                        content = TextUtil.HexDump(_bytes)
                    End Try
                Case ViewerKind.Text
                    ' An M-ZAL TED file is binary records: show its lines. For other assembler source the
                    ' text stops where the file really ends, so slack left by an older, longer version isn't shown.
                    content = If(AsmSource.TedListing(_bytes, _asm),
                                 TextUtil.BytesToDisplayText(If(_asm Is Nothing, _bytes, AsmSource.ContentBytes(_bytes, _asm)), _disk.IsCpm))
                Case ViewerKind.Disassembly
                    content = Await DisassembleAsync()
                Case ViewerKind.Edtasm
                    content = Await EdtasmAsync()
            End Select
        Catch ex As Exception
            content = "(" & ex.Message & ")"
        End Try
        If content Is Nothing Then content = ""
        _cache(_current) = content
        SetText(content)
    End Function

    Private Sub SetText(content As String)
        ApplyFont()
        ' A NUL would end the text for GTK; show it as the control picture a hex-minded reader expects.
        _text.Text = Ui.Lf(content).Replace(ControlChars.NullChar, "␀"c)
        Try
            _text.CaretIndex = 0
        Catch
        End Try
    End Sub

    Private Sub ApplyFont()
        _text.Font = If(_current = ViewerKind.Text AndAlso _rbProp.Checked, PropFont, MonoFont)
    End Sub

    ''' <summary>Extract the file to a temp folder so host-file commands (disasm, convert) can read it.</summary>
    Private Async Function StageHostCopyAsync() As Task(Of String)
        Dim folder As String = Program.NewTempDir()
        Dim host As String = Path.Combine(folder, TextUtil.HostNameFromTrs(_file.TrsName, False))
        Await Program.Tool.DiskExtractAsAsync(_disk.Path, _disk.PartitionArg, _file.TrsName, host)
        Return host
    End Function

    ''' <summary>A hi-res screen file as trs80-tool draws it: extract it, "convert" it to PNG, load the PNG.</summary>
    Private Async Function GraphicsAsync() As Task(Of Image)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Dim png As String = Path.ChangeExtension(host, ".png")
            Await Program.Tool.ConvertAsync(host, png, False)
            _pngBytes = File.ReadAllBytes(png)
            Return New Bitmap(_pngBytes)
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    Private Async Function DisassembleAsync() As Task(Of String)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Return Await Program.Tool.DisasmAsync(host, _disasmWords)
        Catch ex As ToolException
            Return UiText.T("dlg.viewer.disasmfailed", ex.Message) & vbLf & vbLf & TextUtil.HexDump(_bytes)
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    Private Async Function EdtasmAsync() As Task(Of String)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Dim asc As String = Path.ChangeExtension(host, ".asc")
            Await Program.Tool.ConvertAsync(host, asc, False)
            Return TextUtil.ExpandTabs(Encoding.Latin1.GetString(File.ReadAllBytes(asc)))
        Catch ex As ToolException
            ' Fall back to TRSTool's own decoding if the engine balks, using the real end of the file.
            Return FileKinds.DecodeEdtasmLocally(If(_asm Is Nothing, _bytes, AsmSource.ContentBytes(_bytes, _asm)))
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    ''' <summary>Options…: disasm flags for the Disassembly view, the token dialect for the BASIC view.</summary>
    Private Async Function EditOptionsAsync() As Task
        If _current = ViewerKind.Disassembly Then
            Dim dlg As New DisasmOptionsForm(_disasmWords)
            If Not dlg.ShowOk(Me) Then Return
            _disasmWords = dlg.OptionWords()
            _cache.Remove(ViewerKind.Disassembly)
        ElseIf _current = ViewerKind.Basic Then
            Dim dlg As New BasicDialectForm(_basicDialect)
            If Not dlg.ShowOk(Me) Then Return
            _basicDialect = dlg.Dialect
            _cache.Remove(ViewerKind.Basic)
        Else
            Return
        End If
        Await ShowKindAsync()
    End Function

    Private Sub SaveCurrent()
        Dim target As String
        If _current = ViewerKind.Graphics Then
            target = Ui.PickSave(Me, UiText.T("caption.saveview"), UiText.T("filter.png"), _file.BaseName & ".png", "png")
        Else
            target = Ui.PickSave(Me, UiText.T("caption.saveview"), UiText.T("filter.textfile"), _file.BaseName & "_" & KindFileTag(_current) & ".txt", "txt")
        End If
        If target Is Nothing Then Return
        Try
            If _current = ViewerKind.Graphics Then
                If _pngBytes IsNot Nothing Then File.WriteAllBytes(target, _pngBytes)
            Else
                File.WriteAllText(target, Ui.HostLines(_text.Text), New UTF8Encoding(True))
            End If
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, UiText.T("caption.saveviewerr"))
        End Try
    End Sub
End Class

''' <summary>Every "trs80-tool disasm" flag as a control; OptionWords() returns the words for the command line.</summary>
Public Class DisasmOptionsForm
    Inherits OkCancelDialog

    Private ReadOnly _chkListing As New CheckBox With {.Text = UiText.T("dlg.disasm.listing")}
    Private ReadOnly _chkNoBinary As New CheckBox With {.Text = UiText.T("dlg.disasm.nobinary")}
    Private ReadOnly _chkUpper As New CheckBox With {.Text = UiText.T("dlg.disasm.upper")}
    Private ReadOnly _chkNoLabels As New CheckBox With {.Text = UiText.T("dlg.disasm.nolabels")}
    Private ReadOnly _chkNoKnown As New CheckBox With {.Text = UiText.T("dlg.disasm.noknown")}
    Private ReadOnly _chkFull As New CheckBox With {.Text = UiText.T("dlg.disasm.full")}
    Private ReadOnly _chkDataComment As New CheckBox With {.Text = UiText.T("dlg.disasm.datacomment")}
    Private ReadOnly _cmbHex As New DropDown With {.Width = 170}
    Private ReadOnly _txtOrg As New TextBox With {.Width = 100}
    Private ReadOnly _txtEntry As New TextBox With {.Width = 260}
    Private ReadOnly _chkSaveDefault As New CheckBox With {.Text = UiText.T("dlg.disasm.savedefault")}

    Public Sub New(current As IEnumerable(Of String))
        MyBase.New(UiText.T("dlg.disasm.title"))
        ' SelectedIndex decides the --hex-format word, so these four are text only.
        For Each item In New String() {UiText.T("dlg.disasm.hex.default"), UiText.T("dlg.disasm.hex.c"), UiText.T("dlg.disasm.hex.dollar"), UiText.T("dlg.disasm.hex.h")}
            _cmbHex.Items.Add(item)
        Next
        _cmbHex.SelectedIndex = 0
        Finish(Ui.Down(Nothing, _chkListing, _chkNoBinary, _chkUpper, _chkNoLabels, _chkNoKnown, _chkFull, _chkDataComment,
                       Ui.Fields(UiText.T("dlg.disasm.hexformat"), Ui.Across(_cmbHex),
                                 UiText.T("dlg.disasm.org"), Ui.Across(_txtOrg),
                                 UiText.T("dlg.disasm.entry"), Ui.Across(_txtEntry)),
                       _chkSaveDefault))
        LoadWords(current)
    End Sub

    Protected Overrides Function Accept() As Boolean
        If Ui.IsOn(_chkSaveDefault) Then
            Program.Settings.DisasmOptions = String.Join(" ", OptionWords())
            Program.Settings.Save()
        End If
        Return True
    End Function

    Private Sub LoadWords(words As IEnumerable(Of String))
        If words Is Nothing Then Return
        Dim list As List(Of String) = words.ToList()
        For i As Integer = 0 To list.Count - 1
            Select Case list(i)
                Case "--listing" : _chkListing.Checked = True
                Case "--no-binary" : _chkNoBinary.Checked = True
                Case "--upper" : _chkUpper.Checked = True
                Case "--no-labels" : _chkNoLabels.Checked = True
                Case "--no-known" : _chkNoKnown.Checked = True
                Case "--full" : _chkFull.Checked = True
                Case "--data-comment" : _chkDataComment.Checked = True
                Case "--hex-format"
                    If i + 1 < list.Count Then
                        Select Case list(i + 1)
                            Case "c" : _cmbHex.SelectedIndex = 1
                            Case "dollar" : _cmbHex.SelectedIndex = 2
                            Case "h" : _cmbHex.SelectedIndex = 3
                        End Select
                    End If
                Case "--org"
                    If i + 1 < list.Count Then _txtOrg.Text = list(i + 1)
                Case "--entry"
                    If i + 1 < list.Count Then _txtEntry.Text = list(i + 1)
            End Select
        Next
    End Sub

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If Ui.IsOn(_chkListing) Then w.Add("--listing")
        If Ui.IsOn(_chkNoBinary) Then w.Add("--no-binary")
        If Ui.IsOn(_chkUpper) Then w.Add("--upper")
        If Ui.IsOn(_chkNoLabels) Then w.Add("--no-labels")
        If Ui.IsOn(_chkNoKnown) Then w.Add("--no-known")
        If Ui.IsOn(_chkFull) Then w.Add("--full")
        If Ui.IsOn(_chkDataComment) Then w.Add("--data-comment")
        Select Case _cmbHex.SelectedIndex
            Case 1 : w.Add("--hex-format") : w.Add("c")
            Case 2 : w.Add("--hex-format") : w.Add("dollar")
            Case 3 : w.Add("--hex-format") : w.Add("h")
        End Select
        If _txtOrg.Text.Trim().Length > 0 Then
            w.Add("--org") : w.Add(_txtOrg.Text.Trim())
        End If
        If _txtEntry.Text.Trim().Length > 0 Then
            w.Add("--entry") : w.Add(_txtEntry.Text.Trim().Replace(" ", ""))
        End If
        Return w
    End Function
End Class

''' <summary>Which BASIC's tokens "disk list" should use (trs80-tool: --basic 3|4).</summary>
Public Class BasicDialectForm
    Inherits OkCancelDialog

    Private ReadOnly _rbDefault, _rb3, _rb4 As RadioButton

    Public Sub New(current As String)
        MyBase.New(UiText.T("dlg.basicdialect.title"))
        _rbDefault = New RadioButton With {.Text = UiText.T("dlg.basicdialect.default")}
        _rb3 = New RadioButton(_rbDefault) With {.Text = UiText.T("dlg.basicdialect.m13")}
        _rb4 = New RadioButton(_rbDefault) With {.Text = UiText.T("dlg.basicdialect.m4")}
        Select Case current
            Case "3" : _rb3.Checked = True
            Case "4" : _rb4.Checked = True
            Case Else : _rbDefault.Checked = True
        End Select
        Finish(Ui.Down(Nothing, _rbDefault, _rb3, _rb4))
    End Sub

    Public ReadOnly Property Dialect As String
        Get
            If _rb3.Checked Then Return "3"
            If _rb4.Checked Then Return "4"
            Return ""
        End Get
    End Property
End Class
