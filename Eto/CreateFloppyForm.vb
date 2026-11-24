' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  CreateFloppyForm.vb — "Create New Floppy Image" in the Eto.Forms program,
'  and the same dialog as "Format Disk Image". The choices are limited to
'  what trs80-tool's "disk create" and "disk format" can make.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text.RegularExpressions
Imports Eto.Drawing
Imports Eto.Forms

Public Class CreateFloppyForm
    Inherits OkCancelDialog

    Private ReadOnly _rbJv1, _rbJv3, _rbDmk As RadioButton
    Private ReadOnly _rbT23, _rbT13, _rbLdos, _rbT62, _rbLs63, _rbCpm As RadioButton
    ''' <summary>trs80-tool 5.2.0 formats and creates Montezuma Micro CP/M disks (create cpm, format cpm).</summary>
    Private ReadOnly _cpmAvailable As Boolean = Program.Tool IsNot Nothing AndAlso Program.Tool.Has520
    Private ReadOnly _chkSystem As New CheckBox With {.Text = UiText.T("dlg.newfloppy.system")}
    Private ReadOnly _chkMinimal As New CheckBox With {.Text = UiText.T("dlg.newfloppy.minimal"), .Enabled = False}
    Private ReadOnly _rbModel1, _rbModel3 As RadioButton
    Private ReadOnly _rbSd, _rbDd As RadioButton
    Private ReadOnly _rbSs, _rbDs As RadioButton
    Private ReadOnly _rb35, _rb40, _rb80, _rbOther As RadioButton
    Private ReadOnly _txtTracks As New TextBox With {.Width = 60, .Text = "40", .Enabled = False}
    Private ReadOnly _txtName As New TextBox With {.Width = 120, .MaxLength = 8, .Text = "TRS80DSK"}
    Private ReadOnly _txtDate As New TextBox With {.Width = 120}
    Private _updating As Boolean
    Private ReadOnly _formatMode As Boolean

    Public Sub New()
        Me.New(False, Nothing, 0, 0, Nothing)
    End Sub

    ''' <summary>
    ''' formatMode = True turns the dialog into "Format Disk Image" for the open DMK: the image type and the
    ''' System box are left out, and the DOS and geometry start from the open disk's.
    ''' </summary>
    Public Sub New(formatMode As Boolean, currentDos As String, cylinders As Integer, sides As Integer, density As String)
        MyBase.New(If(formatMode, UiText.T("dlg.newfloppy.formattitle"), UiText.T("dlg.newfloppy.title")))
        _formatMode = formatMode
        _rbDmk = New RadioButton With {.Text = "DMK", .Checked = True}
        _rbJv1 = New RadioButton(_rbDmk) With {.Text = "JV1"}
        _rbJv3 = New RadioButton(_rbDmk) With {.Text = "JV3"}
        _rbLdos = New RadioButton With {.Text = UiText.T("dos.ldos53.m13"), .Checked = True}
        _rbT23 = New RadioButton(_rbLdos) With {.Text = UiText.T("dos.trsdos23.m1")}
        _rbT13 = New RadioButton(_rbLdos) With {.Text = UiText.T("dos.trsdos13.m3")}
        _rbT62 = New RadioButton(_rbLdos) With {.Text = UiText.T("dos.trsdos62.m4")}
        _rbLs63 = New RadioButton(_rbLdos) With {.Text = UiText.T("dos.lsdos63.m4")}
        _rbCpm = New RadioButton(_rbLdos) With {.Text = UiText.T("dos.cpm.mm22.m4")}
        _rbModel3 = New RadioButton With {.Text = "III", .Checked = True}
        _rbModel1 = New RadioButton(_rbModel3) With {.Text = "I"}
        _rbDd = New RadioButton With {.Text = UiText.T("dlg.newfloppy.double"), .Checked = True}
        _rbSd = New RadioButton(_rbDd) With {.Text = UiText.T("dlg.newfloppy.single")}
        _rbSs = New RadioButton With {.Text = "1", .Checked = True}
        _rbDs = New RadioButton(_rbSs) With {.Text = "2"}
        _rb40 = New RadioButton With {.Text = "40", .Checked = True}
        _rb35 = New RadioButton(_rb40) With {.Text = "35"}
        _rb80 = New RadioButton(_rb40) With {.Text = "80"}
        _rbOther = New RadioButton(_rb40) With {.Text = UiText.T("dlg.newfloppy.other")}
        _txtDate.Text = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        BuildUi()
        If formatMode Then
            Dim d As String = If(currentDos, "")
            If Regex.IsMatch(d, "^TRSDOS 2\.") Then
                _rbT23.Checked = True
            ElseIf Regex.IsMatch(d, "^TRSDOS 1\.") Then
                _rbT13.Checked = True
            ElseIf Regex.IsMatch(d, "^TRSDOS 6\.") Then
                _rbT62.Checked = True
            ElseIf Regex.IsMatch(d, "^LS-DOS 6\.") Then
                _rbLs63.Checked = True
            ElseIf _cpmAvailable AndAlso d.IndexOf("CP/M", StringComparison.OrdinalIgnoreCase) >= 0 Then
                _rbCpm.Checked = True
            Else
                _rbLdos.Checked = True
            End If
            If sides = 2 Then _rbDs.Checked = True
            If String.Equals(density, "single", StringComparison.OrdinalIgnoreCase) Then
                _rbSd.Checked = True : _rbModel1.Checked = True
            ElseIf String.Equals(density, "double", StringComparison.OrdinalIgnoreCase) Then
                _rbDd.Checked = True : _rbModel3.Checked = True
            End If
            Select Case cylinders
                Case 0
                Case 35 : _rb35.Checked = True
                Case 40 : _rb40.Checked = True
                Case 80 : _rb80.Checked = True
                Case Else
                    _rbOther.Checked = True
                    _txtTracks.Text = cylinders.ToString(CultureInfo.InvariantCulture)
            End Select
        End If
        ApplyInterlocks()
    End Sub

    ''' <summary>jv1 / jv3 / dmk (lower case, no dot).</summary>
    Public ReadOnly Property ImageTypeExtension As String
        Get
            If _rbJv1.Checked Then Return "jv1"
            If _rbJv3.Checked Then Return "jv3"
            Return "dmk"
        End Get
    End Property

    Public ReadOnly Property ImageTypeFilter As String
        Get
            Select Case ImageTypeExtension
                Case "jv1" : Return UiText.T("filter.jv1")
                Case "jv3" : Return UiText.T("filter.jv3")
                Case Else : Return UiText.T("filter.dmk")
            End Select
        End Get
    End Property

    Public ReadOnly Property IsCpm As Boolean
        Get
            Return _rbCpm.Checked
        End Get
    End Property

    ''' <summary>trs80-tool DOS name for "disk create".</summary>
    Public ReadOnly Property DosName As String
        Get
            If _rbCpm.Checked Then Return "cpm"
            If _rbT23.Checked Then Return "trsdos-2.3"
            If _rbT13.Checked Then Return "trsdos-1.3"
            If _rbT62.Checked Then Return "trsdos-6.2"
            If _rbLs63.Checked Then Return "lsdos-6.3"
            Return "ldos-5.3"
        End Get
    End Property

    ''' <summary>The option words after "create DOS" (or "format DOS").</summary>
    Public Function CreateOptionWords() As List(Of String)
        Dim w As New List(Of String)
        Dim sys As Boolean = Ui.IsOn(_chkSystem)
        If _rbCpm.Checked Then
            ' A Montezuma disk: 40 cylinders (170 KB), or 80 with two sides (710 KB); one side of 80 is
            ' refused by trs80-tool with its own sentence. It has no label, so no name and no date.
            If sys Then
                w.Add("--system")
            Else
                w.Add("--cylinders") : w.Add(If(_rb80.Checked, "80", "40"))
                w.Add("--sides") : w.Add(If(_rbDs.Checked, "2", "1"))
            End If
            Return w
        End If
        If sys Then
            w.Add("--system")
            If Ui.IsOn(_chkMinimal) Then w.Add("--minimal")
            If _rbLdos.Checked Then
                w.Add("--model") : w.Add(If(_rbModel1.Checked, "1", "3"))
            End If
        Else
            w.Add("--cylinders") : w.Add(TrackCount().ToString(CultureInfo.InvariantCulture))
            w.Add("--sides") : w.Add(If(_rbDs.Checked, "2", "1"))
            w.Add("--density") : w.Add(If(_rbSd.Checked, "single", "double"))
            If _rbLdos.Checked Then
                w.Add("--model") : w.Add(If(_rbModel1.Checked, "1", "3"))
            End If
        End If
        Dim nm As String = _txtName.Text.Trim().ToUpperInvariant()
        If nm.Length > 0 Then
            w.Add("--name") : w.Add(nm)
        End If
        Dim ymd As String = TextUtil.ToYmd(_txtDate.Text)
        If ymd IsNot Nothing Then
            w.Add("--date") : w.Add(ymd)
        End If
        Return w
    End Function

    Private Function TrackCount() As Integer
        If _rb35.Checked Then Return 35
        If _rb40.Checked Then Return 40
        If _rb80.Checked Then Return 80
        Dim n As Integer
        If Integer.TryParse(_txtTracks.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, n) AndAlso n >= 1 AndAlso n <= 255 Then Return n
        Return 40
    End Function

    Private Sub BuildUi()
        Dim dosRows As New List(Of Control) From {_rbT23, _rbT13, Ui.Across(_rbLdos, Ui.Lbl("    " & UiText.T("dlg.newfloppy.formodel")), _rbModel1, _rbModel3), _rbT62, _rbLs63}
        If _cpmAvailable Then dosRows.Add(_rbCpm)
        Dim gDos As GroupBox = Ui.Group(UiText.T("dlg.newfloppy.dos"), Ui.Down(Nothing, dosRows.ToArray()))

        Dim firstLeft As Control
        If _formatMode Then
            firstLeft = Ui.Wrapped(UiText.T("dlg.newfloppy.formatnote"), 170)
        Else
            firstLeft = Ui.Group(UiText.T("dlg.newfloppy.imagetype"), Ui.Down(Nothing, _rbJv1, _rbJv3, _rbDmk))
        End If
        Dim top As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        top.Rows.Add(New TableRow(New TableCell(firstLeft), New TableCell(gDos, True)))

        Dim gDensity As GroupBox = Ui.Group(UiText.T("dlg.newfloppy.density"), Ui.Down(Nothing, _rbSd, _rbDd))
        Dim gSides As GroupBox = Ui.Group(UiText.T("dlg.newfloppy.sides"), Ui.Down(Nothing, _rbSs, _rbDs))
        Dim gTracks As GroupBox = Ui.Group(UiText.T("dlg.newfloppy.tracks"), Ui.Down(Nothing, Ui.Across(_rb35, _rb40, _rb80), Ui.Across(_rbOther, _txtTracks)))
        Dim middle As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        middle.Rows.Add(New TableRow(New TableCell(gDensity, True), New TableCell(gSides, True), New TableCell(gTracks, True)))

        Dim gVolume As GroupBox = Ui.Group(UiText.T("dlg.newfloppy.volume"), Ui.Fields(UiText.T("dlg.newfloppy.name"), Ui.Across(_txtName), UiText.T("dlg.newfloppy.date"), Ui.Across(_txtDate)))

        Dim rows As New List(Of Control) From {top}
        If Not _formatMode Then rows.Add(Ui.Across(_chkSystem, _chkMinimal))
        rows.Add(middle)
        rows.Add(gVolume)
        Finish(Ui.Down(Nothing, rows.ToArray()), If(_formatMode, UiText.T("btn.format"), Nothing))

        For Each rb In New RadioButton() {_rbJv1, _rbJv3, _rbDmk, _rbT23, _rbT13, _rbLdos, _rbT62, _rbLs63, _rbCpm, _rbModel1, _rbModel3, _rbSd, _rbDd, _rbSs, _rbDs, _rb35, _rb40, _rb80, _rbOther}
            AddHandler rb.CheckedChanged, Sub() ApplyInterlocks()
        Next
        AddHandler _chkSystem.CheckedChanged, Sub() ApplyInterlocks()
    End Sub

    ''' <summary>
    ''' TRSTool's interlocks, narrowed to what trs80-tool can create:
    '''  - JV1 → single density, 1 side (the JV1 format can't hold anything else); TRSDOS 1.3 unavailable.
    '''  - TRSDOS 2.3 → 35 cylinders, single density, 1 side (the only layout it supports).
    '''  - TRSDOS 1.3 → 40 cylinders, double density, 1 side.
    '''  - TRSDOS 6 / LS-DOS 6 → double density only; any cylinder count; 1 or 2 sides.
    '''  - LDOS 5.3 → Model I = single density, Model III = double density.
    '''  - System disk → geometry fixed by the bundled disk; the geometry controls are disabled.
    ''' </summary>
    Private Sub ApplyInterlocks()
        If _updating Then Return
        _updating = True
        Try
            Dim jv1 As Boolean = _rbJv1.Checked
            Dim t23 As Boolean = _rbT23.Checked
            Dim t13 As Boolean = _rbT13.Checked
            Dim ldos As Boolean = _rbLdos.Checked
            Dim m4 As Boolean = _rbT62.Checked OrElse _rbLs63.Checked
            Dim sys As Boolean = Ui.IsOn(_chkSystem)
            Dim cpm As Boolean = _rbCpm.Checked

            _rbT13.Enabled = Not jv1
            _rbCpm.Enabled = Not jv1
            If jv1 AndAlso t13 Then _rbLdos.Checked = True : t13 = False : ldos = True
            If jv1 AndAlso cpm Then _rbLdos.Checked = True : cpm = False : ldos = True

            If cpm Then
                ' Montezuma Micro: double density; 40 cylinders (one or two sides) or 80 with two sides.
                ' The bundled system disk is 40 cylinders, one side.
                _rbDd.Checked = True
                If Not _rb80.Checked Then _rb40.Checked = True
                If sys Then _rb40.Checked = True : _rbSs.Checked = True
            End If

            If jv1 Then _rbSd.Checked = True : _rbSs.Checked = True
            If t23 Then _rbSd.Checked = True : _rbSs.Checked = True : _rb35.Checked = True
            If t13 Then _rbDd.Checked = True : _rbSs.Checked = True : _rb40.Checked = True
            If m4 Then _rbDd.Checked = True
            If ldos Then
                If _rbModel1.Checked Then _rbSd.Checked = True Else _rbDd.Checked = True
            End If

            _rbSd.Enabled = Not sys AndAlso Not t13 AndAlso Not m4 AndAlso Not ldos AndAlso Not cpm
            _rbDd.Enabled = Not sys AndAlso Not jv1 AndAlso Not t23 AndAlso Not ldos
            _rbSs.Enabled = Not sys
            _rbDs.Enabled = Not sys AndAlso Not jv1 AndAlso Not t23 AndAlso Not t13
            _rb35.Enabled = Not sys AndAlso Not t13 AndAlso Not cpm
            _rb40.Enabled = Not sys AndAlso Not t23
            _rb80.Enabled = Not sys AndAlso Not t23 AndAlso Not t13
            _rbOther.Enabled = Not sys AndAlso Not t23 AndAlso Not t13 AndAlso Not cpm
            _txtTracks.Enabled = _rbOther.Enabled AndAlso _rbOther.Checked
            _rbModel1.Enabled = ldos : _rbModel3.Enabled = ldos
            ' "create cpm --system" copies the whole bundled disk; it has no --minimal.
            _chkMinimal.Enabled = sys AndAlso Not cpm
            If Not sys OrElse cpm Then _chkMinimal.Checked = False
            _txtName.Enabled = Not cpm
            _txtDate.Enabled = Not cpm
        Finally
            _updating = False
        End Try
    End Sub

    Public ReadOnly Property IsFormatMode As Boolean
        Get
            Return _formatMode
        End Get
    End Property

    Protected Overrides Function Accept() As Boolean
        Dim caption As String = UiText.T("caption.newfloppyerror")
        If _rbCpm.Checked Then Return True     ' nothing to check: no name, no date, and the track count is 40 or 80
        Dim nm As String = _txtName.Text.Trim()
        If nm.Length = 0 OrElse _txtDate.Text.Trim().Length = 0 Then
            Ui.Fail(Me, UiText.T("dlg.newfloppy.needname"), caption)
            Return False
        End If
        If Not Regex.IsMatch(nm, "^[A-Za-z][A-Za-z0-9]{0,7}$") Then
            Ui.Fail(Me, UiText.T("dlg.newfloppy.badname"), caption)
            Return False
        End If
        If TextUtil.ToYmd(_txtDate.Text) Is Nothing Then
            Ui.Fail(Me, UiText.T("dlg.newfloppy.baddate"), caption)
            Return False
        End If
        If _rbOther.Checked Then
            Dim n As Integer
            If Not Integer.TryParse(_txtTracks.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, n) OrElse n < 1 OrElse n > 255 Then
                Ui.Fail(Me, UiText.T("dlg.newfloppy.badtracks"), caption)
                Return False
            End If
        End If
        Return True
    End Function
End Class
