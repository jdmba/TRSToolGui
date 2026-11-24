' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  ViewerForm.vb — the "File viewer" (Enter / double-click / Preview), with
'  TRSTool's representation combo: Graphics, EDTASM, BASIC, Disassembly,
'  Text, Hexadecimal. Representations come from trs80-tool wherever it has
'  one (list, list --hex, disasm, convert; hi-res screens are its PNGs); Text
'  is rendered here from the file's bytes. "Options..." exposes trs80-tool's
'  disasm flags and the BASIC dialect (list --basic 3|4).
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class ViewerForm
    Inherits Form

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

    Private ReadOnly _text As New RichTextBox()
    Private ReadOnly _picPanel As New Panel()
    Private ReadOnly _pic As New PictureBox()
    ''' <summary>The PNG trs80-tool made of a hi-res screen, exactly as it wrote it.</summary>
    Private _pngBytes As Byte() = Nothing
    Private ReadOnly _combo As New ComboBox()
    Private ReadOnly _grpFont As New GroupBox()
    Private ReadOnly _rbMono As New RadioButton()
    Private ReadOnly _rbProp As New RadioButton()
    Private ReadOnly _btnDone As New Button()
    Private ReadOnly _btnSave As New Button()
    Private ReadOnly _btnOptions As New Button()
    ''' <summary>Option words for "disasm"; starts from Preferences → Disassembly options.</summary>
    Private _disasmWords As List(Of String) = Program.Settings.DisasmOptionWords()
    ''' <summary>"" = the disk's default, else "3" or "4" for list --basic.</summary>
    Private _basicDialect As String = ""
    Private ReadOnly _lblInfo As New Label()
    ''' <summary>FixASM's verdict when the file is assembler source (EDTASM, or text with a source extension); else Nothing.</summary>
    Private _asm As AsmSourceInfo
    Private ReadOnly _kinds As New List(Of ViewerKind)
    Private _current As ViewerKind = ViewerKind.Hex
    Private _loading As Boolean = False
    Private ReadOnly _cache As New Dictionary(Of ViewerKind, String)

    Private Shared ReadOnly MonoFont As New Font(HostPlatform.MonospaceFontName, 10.0F)
    Private Shared ReadOnly PropFont As New Font(HostPlatform.SerifFontName, 11.0F)

    Public Sub New(disk As DiskImage, file As DiskFile)
        _disk = disk
        _file = file
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Text = UiText.T("dlg.viewer.title", _file.SlashName)
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(760, 520)
        MinimumSize = New Size(420, 260)
        ShowInTaskbar = False
        KeyPreview = True

        With _text
            .ReadOnly = True
            .Multiline = True
            .WordWrap = False
            .ScrollBars = RichTextBoxScrollBars.Both
            .Font = MonoFont
            .BackColor = SystemColors.Window
            .DetectUrls = False
            .HideSelection = False
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        End With
        With _picPanel
            .AutoScroll = True
            .BackColor = Color.Black
            .BorderStyle = BorderStyle.Fixed3D
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            .Visible = False
        End With
        _pic.SizeMode = PictureBoxSizeMode.AutoSize
        _picPanel.Controls.Add(_pic)

        _grpFont.Text = UiText.T("dlg.viewer.font")
        _rbMono.Text = UiText.T("dlg.viewer.mono") : _rbMono.AutoSize = True
        _rbProp.Text = UiText.T("dlg.viewer.prop") : _rbProp.AutoSize = True
        _grpFont.Controls.Add(_rbMono)
        _grpFont.Controls.Add(_rbProp)
        _rbMono.Location = New Point(10, 18)
        _rbProp.Location = New Point(110, 18)
        _grpFont.Size = New Size(215, 44)
        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.SpreadRow(10, _rbMono, _rbProp)
        UiFit.GrowToFit(_grpFont, 8)
        AddHandler _rbMono.CheckedChanged, Sub() ApplyFont()
        AddHandler _rbProp.CheckedChanged, Sub() ApplyFont()

        _combo.DropDownStyle = ComboBoxStyle.DropDownList
        _combo.Width = 150
        AddHandler _combo.SelectedIndexChanged, Async Sub() Await ShowKindAsync()

        _btnDone.Text = UiText.T("btn.done")
        _btnDone.DialogResult = DialogResult.OK
        AddHandler _btnDone.Click, Sub() Close()
        _btnSave.Text = UiText.T("btn.saveas")
        AddHandler _btnSave.Click, Sub() SaveCurrent()
        _btnOptions.Text = UiText.T("btn.options")
        _btnOptions.Enabled = False
        AddHandler _btnOptions.Click, Async Sub() Await EditOptionsAsync()

        _lblInfo.AutoSize = True
        _lblInfo.ForeColor = SystemColors.GrayText

        Controls.AddRange(New Control() {_text, _picPanel, _grpFont, _combo, _btnOptions, _btnSave, _btnDone, _lblInfo})
        AcceptButton = _btnDone
        CancelButton = _btnDone
        AddHandler Me.Resize, Sub() Layout2()
        Layout2()
    End Sub

    Private Sub Layout2()
        Dim m As Integer = 10
        Dim bottomH As Integer = 50
        Dim w As Integer = ClientSize.Width
        Dim h As Integer = ClientSize.Height
        _text.Bounds = New Rectangle(m, m, w - 2 * m, h - bottomH - 2 * m)
        _picPanel.Bounds = _text.Bounds
        Dim y As Integer = h - bottomH + 2
        _grpFont.Location = New Point(m, y - 4)
        _combo.Location = New Point(_grpFont.Right + 12, y + 10)
        _btnOptions.Bounds = New Rectangle(_combo.Right + 8, y + 8, 80, 26)
        UiFit.FitButton(_btnOptions)
        _lblInfo.Location = New Point(_btnOptions.Right + 12, y + 14)
        _btnDone.Bounds = New Rectangle(w - m - 80, y + 8, 80, 26)
        _btnSave.Bounds = New Rectangle(_btnDone.Left - 90, y + 8, 84, 26)
        ' The gap of 6 is what the two fixed positions above come to in English, so this moves
        ' nothing there and only makes room when a longer language needs it (UiFit.vb).
        UiFit.ButtonRow(w - m, 6, _btnSave, _btnDone)
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        _rbMono.Checked = Program.Settings.ViewerMonospaced
        _rbProp.Checked = Not _rbMono.Checked
        UseWaitCursor = True
        Try
            _bytes = Await Program.Tool.ReadFileBytesAsync(_disk.Path, _disk.PartitionArg, _file.TrsName)
        Catch ex As Exception
            UseWaitCursor = False
            MessageBox.Show(Me, UiText.TCrLf("dlg.viewer.readfailed", _file.SlashName, ex.Message), UiText.T("caption.fileviewer"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close()
            Return
        End Try
        UseWaitCursor = False
        _lblInfo.Text = UiText.T("word.bytes", TextUtil.FormatBytes(_bytes.LongLength)) & If(_file.Description.Length > 0, "  •  " & _file.Description, "")
        DetectKinds()
        If _kinds(0) = ViewerKind.Edtasm OrElse (_kinds(0) = ViewerKind.Text AndAlso AsmSource.IsSourceExtension(_file.Extension)) Then
            _asm = AsmSource.Analyze(_bytes)
            _lblInfo.Text &= "  •  " & _asm.Summary()
        End If
        Dim basicKind As String = BasicKindLabel()
        If basicKind.Length > 0 Then Text = UiText.T("dlg.viewer.titlekind", UiText.T("dlg.viewer.title", _file.SlashName), basicKind)
        _loading = True
        _combo.Items.Clear()
        For Each k In _kinds
            _combo.Items.Add(KindName(k))
        Next
        _loading = False
        _combo.SelectedIndex = 0
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Program.Settings.ViewerMonospaced = _rbMono.Checked
        MyBase.OnFormClosing(e)
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

    ''' <summary>TRSTool's detection order, helped by trs80-tool's contentType.</summary>
    Private Sub DetectKinds()
        _kinds.Clear()
        Dim trsdos As Boolean = _disk.IsTrsdosFamily
        Dim ct As String = _file.ContentType
        If trsdos AndAlso ct = "hiResScreen" Then
            ' trs80-tool recognises hi-res screen dumps itself (HR, HRG, HRX, CHR, SHR) and turns them into PNGs.
            _kinds.Add(ViewerKind.Graphics) : _kinds.Add(ViewerKind.Hex)
        ElseIf trsdos AndAlso (ct = "edtasm" OrElse IsEdtasm(_bytes)) Then
            _kinds.Add(ViewerKind.Edtasm) : _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        ElseIf (trsdos OrElse _disk.DosName.Length = 0) AndAlso IsTokenizedBasic(_bytes) Then
            _kinds.Add(ViewerKind.Basic) : _kinds.Add(ViewerKind.Hex)
        ElseIf trsdos AndAlso (ct = "cmdProgram" OrElse IsCmd(_bytes)) Then
            _kinds.Add(ViewerKind.Disassembly) : _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        ElseIf TextUtil.LooksLikeText(_file.Extension, _bytes) Then
            _kinds.Add(ViewerKind.Text) : _kinds.Add(ViewerKind.Hex)
        Else
            _kinds.Add(ViewerKind.Hex) : _kinds.Add(ViewerKind.Text)
        End If
    End Sub

    ' TRSTool's byte tests (Filters\EDTASM.cpp, LevelIIBasic.cpp, CMDFile.cpp)
    Private Shared Function IsEdtasm(b As Byte()) As Boolean
        Try
            Dim pos As Integer = If(b(0) = &HD3, 7, 0)
            For i As Integer = 0 To 4
                Dim ch As Byte = b(pos + i)
                If ch < (&H30 Or &H80) OrElse ch > (&H39 Or &H80) Then Return False
            Next
            Dim c As Byte = b(pos + 5)
            Return c = &H20 OrElse c = &H9
        Catch ex As IndexOutOfRangeException
            Return False
        End Try
    End Function

    ''' <summary>
    ''' FF, then line records: next-line pointer (2 bytes), line number (2), tokens, 00; a 0 pointer ends it.
    ''' The pointers are load addresses (Model 4 BASIC loads well above Model III, so no address limit), but
    ''' consecutive ones differ by exactly the record lengths. Model 4 binary constants can hold 00 bytes, so
    ''' each 00 in the first line is tried as its end and the chain is followed from there (up to 8 lines).
    ''' trs80-tool then picks the token table itself (guessBasicDialect).
    ''' </summary>
    Private Shared Function IsTokenizedBasic(b As Byte()) As Boolean
        If b.Length < 7 OrElse b(0) <> &HFF Then Return False
        Dim ptr0 As Integer = b(1) Or (CInt(b(2)) << 8)
        If ptr0 = 0 Then Return False
        ' TRSTool's rule, and FFFF = the placeholder pointers trs80-tool writes when it tokenizes (BASIC relinks on load).
        If ptr0 <= &H8500 OrElse ptr0 = &HFFFF Then Return True
        Dim lastEnd As Integer = Math.Min(b.Length - 1, 1 + 4 + 300)
        For z As Integer = 5 To lastEnd
            If b(z) <> 0 Then Continue For
            Dim base As Integer = ptr0 - (z + 1)
            If BasicChainHolds(b, z + 1, base, b(3) Or (CInt(b(4)) << 8)) Then Return True
        Next
        Return False
    End Function

    Private Shared Function BasicChainHolds(b As Byte(), startPos As Integer, base As Integer, prevLine As Integer) As Boolean
        Dim p As Integer = startPos
        For n As Integer = 1 To 8
            If p = b.Length Then Return False                    ' no end marker before the bytes run out
            If p + 1 >= b.Length Then Return b(p) = 0           ' a single trailing byte
            Dim ptr As Integer = b(p) Or (CInt(b(p + 1)) << 8)
            If ptr = 0 Then Return True                         ' end of program
            If p + 4 >= b.Length Then Return False
            Dim lineNo As Integer = b(p + 2) Or (CInt(b(p + 3)) << 8)
            If lineNo <= prevLine Then Return False
            Dim nextP As Integer = ptr - base
            If nextP <= p + 4 OrElse nextP > b.Length OrElse b(nextP - 1) <> 0 Then Return False
            prevLine = lineNo
            p = nextP
        Next
        Return True
    End Function

    ''' <summary>"ASCII", "Model I/III Tokenized" or "Model 4 Tokenized" for a BASIC program, else "" (title bar only).</summary>
    Private Function BasicKindLabel() As String
        If _kinds.Contains(ViewerKind.Basic) Then
            Return If(IsModel4Basic(_bytes), UiText.T("dlg.viewer.basic.m4"), UiText.T("dlg.viewer.basic.m13"))
        End If
        If _kinds.Count > 0 AndAlso _kinds(0) = ViewerKind.Text AndAlso IsAsciiBasic(_bytes) Then Return UiText.T("dlg.viewer.basic.ascii")
        Return ""
    End Function

    ''' <summary>
    ''' Model 4 BASIC (MS BASIC-80 5.x) writes FF-prefixed function tokens and binary number constants
    ''' (bytes 0B-1F); Level II / Model III BASIC has neither (tokens stop at FB, numbers stay ASCII).
    ''' Quoted strings and the rest of a REM / ' line are skipped (REM is 93 or FB on Model III, 8F on Model 4).
    ''' </summary>
    Private Shared Function IsModel4Basic(b As Byte()) As Boolean
        Dim p As Integer = 1
        Dim m4 As Integer = 0
        While p + 4 < b.Length
            Dim ptr As Integer = b(p) Or (CInt(b(p + 1)) << 8)
            If ptr = 0 Then Exit While
            p += 4
            Dim inString As Boolean = False
            Dim inRem As Boolean = False
            While p < b.Length AndAlso b(p) <> 0
                Dim c As Byte = b(p)
                If inString Then
                    If c = &H22 Then inString = False
                ElseIf Not inRem Then
                    If c = &H22 Then
                        inString = True
                    ElseIf c = &H93 OrElse c = &HFB OrElse c = &H8F Then
                        inRem = True
                    ElseIf c >= &HFC OrElse (c >= &HB AndAlso c <= &H1F) Then
                        m4 += 1
                    End If
                End If
                p += 1
            End While
            ' A binary constant can contain 00; a later line's chain resynchronises enough for a vote.
            p += 1
        End While
        Return m4 >= 2
    End Function

    ''' <summary>Plain-text BASIC: the first few non-blank lines each start with a line number, ascending.</summary>
    Private Shared Function IsAsciiBasic(b As Byte()) As Boolean
        Dim lines As String() = Encoding.Latin1.GetString(b).Split({ChrW(13), ChrW(10)}, StringSplitOptions.RemoveEmptyEntries)
        Dim checkedLines As Integer = 0
        Dim prev As Integer = -1
        For Each raw In lines
            Dim s As String = raw.TrimStart(" "c)
            If s.Length = 0 OrElse s = ChrW(&H1A) Then Continue For
            Dim n As Integer = 0
            While n < s.Length AndAlso n < 6 AndAlso Char.IsDigit(s(n))
                n += 1
            End While
            If n = 0 OrElse n = 6 OrElse n = s.Length Then Return False
            If Not (s(n) = " "c OrElse s(n) = "'"c OrElse Char.IsLetter(s(n))) Then Return False   ' "10 PRINT", "10PRINT"; not "1. item"
            Dim lineNo As Integer = Integer.Parse(s.Substring(0, n), Globalization.CultureInfo.InvariantCulture)
            If lineNo <= prev OrElse lineNo > 65529 Then Return False
            prev = lineNo
            checkedLines += 1
            If checkedLines = 5 Then Exit For
        Next
        Return checkedLines > 0
    End Function

    Private Shared Function IsCmd(b As Byte()) As Boolean
        If b.Length <= 1 Then Return False
        Dim first As Integer = b(0)
        Return first = 1 OrElse first = 3 OrElse first = 5 OrElse first = 7 OrElse first = &H1F
    End Function

    Private Async Function ShowKindAsync() As Task
        If _loading OrElse _combo.SelectedIndex < 0 OrElse _combo.SelectedIndex >= _kinds.Count Then Return
        _current = _kinds(_combo.SelectedIndex)
        _grpFont.Enabled = (_current = ViewerKind.Text)
        _btnOptions.Enabled = (_current = ViewerKind.Disassembly OrElse _current = ViewerKind.Basic)
        If _current = ViewerKind.Graphics Then
            _text.Visible = False
            _picPanel.Visible = True
            If _pic.Image Is Nothing Then
                Try
                    _pic.Image = Await GraphicsAsync()
                Catch ex As Exception
                    _pic.Image = Nothing
                    MessageBox.Show(Me, UiText.T("dlg.viewer.picfailed", ex.Message), UiText.T("caption.fileviewer"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
            Return
        End If
        _picPanel.Visible = False
        _text.Visible = True
        Dim content As String = Nothing
        If _cache.TryGetValue(_current, content) Then
            SetText(content)
            Return
        End If
        UseWaitCursor = True
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
                    content = If(AsmSource.TedListing(_bytes, _asm),
                                 TextUtil.BytesToDisplayText(If(_asm Is Nothing, _bytes, AsmSource.ContentBytes(_bytes, _asm)), _disk.IsCpm))
                Case ViewerKind.Disassembly
                    content = Await DisassembleAsync()
                Case ViewerKind.Edtasm
                    content = Await EdtasmAsync()
            End Select
        Catch ex As Exception
            content = "(" & ex.Message & ")"
        Finally
            UseWaitCursor = False
        End Try
        If content Is Nothing Then content = ""
        _cache(_current) = content
        SetText(content)
    End Function

    Private Sub SetText(content As String)
        _text.SuspendLayout()
        _text.Text = content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
        ApplyFont()
        _text.SelectionStart = 0
        _text.SelectionLength = 0
        _text.ScrollToCaret()
        _text.ResumeLayout()
    End Sub

    Private Sub ApplyFont()
        _text.Font = If(_current = ViewerKind.Text AndAlso _rbProp.Checked, PropFont, MonoFont)
    End Sub

    ''' <summary>Write the file to a temp host file through the engine, for disasm/convert.</summary>
    Private Async Function StageHostCopyAsync() As Task(Of String)
        Dim folder As String = Program.NewTempDir()
        Dim host As String = Path.Combine(folder, TextUtil.HostNameFromTrs(_file.TrsName, False))
        Await Program.Tool.DiskExtractAsAsync(_disk.Path, _disk.PartitionArg, _file.TrsName, host)
        Return host
    End Function

    ''' <summary>
    ''' The hi-res screen as a picture: the file is staged on the host and "trs80-tool convert X.HR X.png" draws it
    ''' (640x480 for a Model 4 screen, 512x384 for a Model III one: each row twice, as the pixels are tall).
    ''' The PNG's bytes are kept for Save as.
    ''' </summary>
    Private Async Function GraphicsAsync() As Task(Of Image)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Dim png As String = Path.ChangeExtension(host, ".png")
            Await Program.Tool.ConvertAsync(host, png, False)
            _pngBytes = File.ReadAllBytes(png)
            Using ms As New MemoryStream(_pngBytes)
                Using loaded As Image = Image.FromStream(ms)
                    Return New Bitmap(loaded)     ' a copy that doesn't need the stream
                End Using
            End Using
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    Private Async Function DisassembleAsync() As Task(Of String)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Return Await Program.Tool.DisasmAsync(host, _disasmWords)
        Catch ex As ToolException
            ' TRSTool fell back to hex when the CMD couldn't be parsed.
            Return UiText.T("dlg.viewer.disasmfailed", ex.Message) & vbCrLf & vbCrLf & TextUtil.HexDump(_bytes)
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    Private Async Function EdtasmAsync() As Task(Of String)
        Dim host As String = Await StageHostCopyAsync()
        Try
            Dim asc As String = Path.ChangeExtension(host, ".asc")
            Await Program.Tool.ConvertAsync(host, asc, False)
            Dim textOut As String = Encoding.Latin1.GetString(File.ReadAllBytes(asc))
            Return TextUtil.ExpandTabs(textOut)
        Catch ex As ToolException
            Return DecodeEdtasmLocally(If(_asm Is Nothing, _bytes, AsmSource.ContentBytes(_bytes, _asm)))
        Finally
            Try : Directory.Delete(Path.GetDirectoryName(host), True) : Catch : End Try
        End Try
    End Function

    ''' <summary>TRSTool's EDTASM filter (fallback only): strip D3 header, clear bit 7 of the 5-digit line number, tab after it.</summary>
    Private Shared Function DecodeEdtasmLocally(b As Byte()) As String
        Dim sb As New StringBuilder(b.Length)
        Dim pos As Integer = If(b.Length > 0 AndAlso b(0) = &HD3, 7, 0)
        Dim linePos As Integer = 0
        For i As Integer = pos To b.Length - 1
            Dim ch As Byte = b(i)
            If ch = &HD Then
                sb.Append(vbCrLf) : linePos = -1
            ElseIf i = b.Length - 1 AndAlso ch = &H1A Then
                Continue For
            Else
                If linePos < 5 Then ch = CByte(ch And &H7F)
                If linePos = 5 Then ch = 9
                sb.Append(ChrW(ch))
            End If
            linePos += 1
        Next
        Return TextUtil.ExpandTabs(sb.ToString())
    End Function

    ''' <summary>Edit the disasm flags or the BASIC dialect, then re-render the current view.</summary>
    Private Async Function EditOptionsAsync() As Task
        If _current = ViewerKind.Disassembly Then
            Using dlg As New DisasmOptionsForm(_disasmWords)
                If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
                _disasmWords = dlg.OptionWords()
            End Using
            _cache.Remove(ViewerKind.Disassembly)
        ElseIf _current = ViewerKind.Basic Then
            Using dlg As New BasicDialectForm(_basicDialect)
                If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
                _basicDialect = dlg.Dialect
            End Using
            _cache.Remove(ViewerKind.Basic)
        Else
            Return
        End If
        Await ShowKindAsync()
    End Function

    Private Sub SaveCurrent()
        Using sfd As New SaveFileDialog()
            sfd.Title = UiText.T("caption.saveview")
            If _current = ViewerKind.Graphics Then
                sfd.Filter = UiText.T("filter.png")
                sfd.FileName = _file.BaseName & ".png"
            Else
                sfd.Filter = UiText.T("filter.textfile")
                sfd.FileName = _file.BaseName & "_" & KindFileTag(_current) & ".txt"
            End If
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                If _current = ViewerKind.Graphics Then
                    If _pngBytes IsNot Nothing Then File.WriteAllBytes(sfd.FileName, _pngBytes)
                Else
                    File.WriteAllText(sfd.FileName, _text.Text, Encoding.UTF8)
                End If
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, UiText.T("caption.saveviewerr"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

End Class


''' <summary>trs80-tool "disasm" flags. Produces the option words the viewer passes to DisasmAsync.</summary>
Public Class DisasmOptionsForm
    Inherits Form

    Private ReadOnly _chkListing As New CheckBox() With {.Text = UiText.T("dlg.disasm.listing"), .AutoSize = True}
    Private ReadOnly _chkNoBinary As New CheckBox() With {.Text = UiText.T("dlg.disasm.nobinary"), .AutoSize = True}
    Private ReadOnly _chkUpper As New CheckBox() With {.Text = UiText.T("dlg.disasm.upper"), .AutoSize = True}
    Private ReadOnly _chkNoLabels As New CheckBox() With {.Text = UiText.T("dlg.disasm.nolabels"), .AutoSize = True}
    Private ReadOnly _chkNoKnown As New CheckBox() With {.Text = UiText.T("dlg.disasm.noknown"), .AutoSize = True}
    Private ReadOnly _chkFull As New CheckBox() With {.Text = UiText.T("dlg.disasm.full"), .AutoSize = True}
    Private ReadOnly _chkDataComment As New CheckBox() With {.Text = UiText.T("dlg.disasm.datacomment"), .AutoSize = True}
    Private ReadOnly _cmbHex As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 160}
    Private ReadOnly _txtOrg As New TextBox() With {.Width = 90}
    Private ReadOnly _txtEntry As New TextBox() With {.Width = 240}
    Private ReadOnly _chkSaveDefault As New CheckBox() With {.Text = UiText.T("dlg.disasm.savedefault"), .AutoSize = True}

    Public Sub New(current As IEnumerable(Of String))
        Text = UiText.T("dlg.disasm.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(470, 360)
        Dim y As Integer = 12
        For Each c In New CheckBox() {_chkListing, _chkNoBinary, _chkUpper, _chkNoLabels, _chkNoKnown, _chkFull, _chkDataComment}
            c.Location = New Point(12, y)
            Controls.Add(c)
            y += 26
        Next
        y += 6
        Dim lblHex As New Label() With {.Text = UiText.T("dlg.disasm.hexformat"), .AutoSize = True, .Location = New Point(12, y + 3)}
        Controls.Add(lblHex)
        ' SelectedIndex decides the --hex-format word, so these four are text only.
        _cmbHex.Items.AddRange(New Object() {UiText.T("dlg.disasm.hex.default"), UiText.T("dlg.disasm.hex.c"), UiText.T("dlg.disasm.hex.dollar"), UiText.T("dlg.disasm.hex.h")})
        _cmbHex.SelectedIndex = 0
        _cmbHex.Location = New Point(260, y)
        Controls.Add(_cmbHex)
        y += 32
        Dim lblOrg As New Label() With {.Text = UiText.T("dlg.disasm.org"), .AutoSize = True, .Location = New Point(12, y + 3)}
        Controls.Add(lblOrg)
        _txtOrg.Location = New Point(330, y)
        Controls.Add(_txtOrg)
        y += 32
        Controls.Add(New Label() With {.Text = UiText.T("dlg.disasm.entry"), .AutoSize = True, .Location = New Point(12, y + 3)})
        y += 24
        _txtEntry.Location = New Point(12, y)
        Controls.Add(_txtEntry)
        y += 32
        _chkSaveDefault.Location = New Point(12, y)
        Controls.Add(_chkSaveDefault)
        y += 34
        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(290, y, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(380, y, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)
        AcceptButton = btnOk : CancelButton = btnCancel
        AddHandler btnOk.Click, Sub()
                                    If _chkSaveDefault.Checked Then
                                        Program.Settings.DisasmOptions = String.Join(" ", OptionWords())
                                        Program.Settings.Save()
                                    End If
                                End Sub
        ' Make room for the language in force (UiFit.vb); in English nothing here moves. The
        ' seven option lines are whole sentences and decide the width.
        UiFit.After(lblHex, _cmbHex, 10)
        UiFit.After(lblOrg, _txtOrg, 10)
        ClientSize = New Size(470, y + 38)
        UiFit.GrowForm(Me)
        _txtEntry.Width = Math.Max(_txtEntry.Width, ClientSize.Width - 24)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)
        LoadWords(current)
    End Sub

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
        If _chkListing.Checked Then w.Add("--listing")
        If _chkNoBinary.Checked Then w.Add("--no-binary")
        If _chkUpper.Checked Then w.Add("--upper")
        If _chkNoLabels.Checked Then w.Add("--no-labels")
        If _chkNoKnown.Checked Then w.Add("--no-known")
        If _chkFull.Checked Then w.Add("--full")
        If _chkDataComment.Checked Then w.Add("--data-comment")
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

''' <summary>Which BASIC token table "list" should use: the disk's default, Model III or Model 4 (list --basic 3|4).</summary>
Public Class BasicDialectForm
    Inherits Form

    Private ReadOnly _rbDefault As New RadioButton() With {.Text = UiText.T("dlg.basicdialect.default"), .AutoSize = True, .Location = New Point(14, 14)}
    Private ReadOnly _rb3 As New RadioButton() With {.Text = UiText.T("dlg.basicdialect.m13"), .AutoSize = True, .Location = New Point(14, 40)}
    Private ReadOnly _rb4 As New RadioButton() With {.Text = UiText.T("dlg.basicdialect.m4"), .AutoSize = True, .Location = New Point(14, 66)}

    Public Sub New(current As String)
        Text = UiText.T("dlg.basicdialect.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(320, 140)
        Controls.AddRange(New Control() {_rbDefault, _rb3, _rb4})
        Select Case current
            Case "3" : _rb3.Checked = True
            Case "4" : _rb4.Checked = True
            Case Else : _rbDefault.Checked = True
        End Select
        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(140, 102, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(230, 102, 80, 26)}
        Controls.Add(btnOk) : Controls.Add(btnCancel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel
    End Sub

    Public ReadOnly Property Dialect As String
        Get
            If _rb3.Checked Then Return "3"
            If _rb4.Checked Then Return "4"
            Return ""
        End Get
    End Property
End Class
