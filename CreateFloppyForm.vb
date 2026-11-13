' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  CreateFloppyForm.vb — "Create New Floppy Image". Produces the DOS name and
'  option words for "disk NEW.dmk create DOS ...". TRSTool's interlocks are
'  kept (TRSDOS 2.3 = single density/1 side/35 tracks; TRSDOS 1.3 = double
'  density/1 side/40 tracks; JV1 = single density/1 side).
'  Format mode: the same dialog drives "disk IMG format DOS ..."
'  for the open image — image type and system-disk choices are hidden and the
'  current disk's DOS/geometry are preselected.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class CreateFloppyForm
    Inherits Form

    Private ReadOnly _rbJv1 As New RadioButton() With {.Text = "JV1", .AutoSize = True}
    Private ReadOnly _rbJv3 As New RadioButton() With {.Text = "JV3", .AutoSize = True}
    Private ReadOnly _rbDmk As New RadioButton() With {.Text = "DMK", .AutoSize = True, .Checked = True}

    Private ReadOnly _rbT23 As New RadioButton() With {.Text = "TRSDOS 2.3 (Model I)", .AutoSize = True}
    Private ReadOnly _rbT13 As New RadioButton() With {.Text = "TRSDOS 1.3 (Model III)", .AutoSize = True}
    Private ReadOnly _rbLdos As New RadioButton() With {.Text = "LDOS 5.3 (Model I/III)", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbT62 As New RadioButton() With {.Text = "TRSDOS 6.2 (Model 4)", .AutoSize = True}
    Private ReadOnly _rbLs63 As New RadioButton() With {.Text = "LS-DOS 6.3 (Model 4)", .AutoSize = True}
    Private ReadOnly _chkSystem As New CheckBox() With {.Text = "System (bootable) disk from trs80-tool's bundled DOS", .AutoSize = True}
    Private ReadOnly _chkMinimal As New CheckBox() With {.Text = "only the DOS's own files (--minimal)", .AutoSize = True, .Enabled = False}

    Private ReadOnly _rbModel1 As New RadioButton() With {.Text = "I", .AutoSize = True}
    Private ReadOnly _rbModel3 As New RadioButton() With {.Text = "III", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbModel4 As New RadioButton() With {.Text = "4", .AutoSize = True}

    Private ReadOnly _rbSd As New RadioButton() With {.Text = "Single", .AutoSize = True}
    Private ReadOnly _rbDd As New RadioButton() With {.Text = "Double", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbSs As New RadioButton() With {.Text = "1", .AutoSize = True, .Checked = True}
    Private ReadOnly _rbDs As New RadioButton() With {.Text = "2", .AutoSize = True}
    Private ReadOnly _rb35 As New RadioButton() With {.Text = "35", .AutoSize = True}
    Private ReadOnly _rb40 As New RadioButton() With {.Text = "40", .AutoSize = True, .Checked = True}
    Private ReadOnly _rb80 As New RadioButton() With {.Text = "80", .AutoSize = True}
    Private ReadOnly _rbOther As New RadioButton() With {.Text = "Other:", .AutoSize = True}
    Private ReadOnly _txtTracks As New TextBox() With {.Width = 40, .Text = "40", .Enabled = False}
    Private ReadOnly _txtName As New TextBox() With {.Width = 110, .MaxLength = 8, .Text = "TRS80DSK"}
    Private ReadOnly _txtDate As New TextBox() With {.Width = 110}
    Private _updating As Boolean
    Private ReadOnly _formatMode As Boolean

    Public Sub New()
        Me.New(False, Nothing, 0, 0, Nothing)
    End Sub

    ''' <summary>
    ''' formatMode: dialog for "disk IMG format" of the open image. currentDos/cylinders/sides/density preselect the
    ''' controls from the open disk (0/Nothing = leave the defaults).
    ''' </summary>
    Public Sub New(formatMode As Boolean, currentDos As String, cylinders As Integer, sides As Integer, density As String)
        _formatMode = formatMode
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

    ''' <summary>"dmk", "jv1" or "jv3".</summary>
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
                Case "jv1" : Return "JV1 floppy image (*.jv1;*.dsk)|*.jv1;*.dsk"
                Case "jv3" : Return "JV3 floppy image (*.jv3;*.dsk)|*.jv3;*.dsk"
                Case Else : Return "DMK floppy image (*.dmk)|*.dmk"
            End Select
        End Get
    End Property

    ''' <summary>The DOS argument of "create": trsdos-2.3, trsdos-1.3, ldos-5.3, trsdos-6.2, lsdos-6.3.</summary>
    Public ReadOnly Property DosName As String
        Get
            If _rbT23.Checked Then Return "trsdos-2.3"
            If _rbT13.Checked Then Return "trsdos-1.3"
            If _rbT62.Checked Then Return "trsdos-6.2"
            If _rbLs63.Checked Then Return "lsdos-6.3"
            Return "ldos-5.3"
        End Get
    End Property

    Public Function CreateOptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _chkSystem.Checked Then
            w.Add("--system")
            If _chkMinimal.Checked Then w.Add("--minimal")
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
        Text = If(_formatMode, "Format Disk Image", "Create New Floppy Image")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(470, 420)

        Dim g1 As New GroupBox() With {.Text = "Disk image type", .Bounds = New Rectangle(10, 10, 140, 100)}
        _rbJv1.Location = New Point(14, 24) : _rbJv3.Location = New Point(14, 48) : _rbDmk.Location = New Point(14, 72)
        g1.Controls.AddRange(New Control() {_rbJv1, _rbJv3, _rbDmk})

        Dim g2 As New GroupBox() With {.Text = "File system (DOS)", .Bounds = New Rectangle(160, 10, 300, 150)}
        _rbT23.Location = New Point(14, 22) : _rbT13.Location = New Point(14, 44) : _rbLdos.Location = New Point(14, 66)
        _rbT62.Location = New Point(14, 88) : _rbLs63.Location = New Point(14, 110)
        Dim lblModel As New Label() With {.Text = "LDOS for Model:", .AutoSize = True, .Location = New Point(180, 68)}
        _rbModel1.Location = New Point(180, 88) : _rbModel3.Location = New Point(220, 88)
        g2.Controls.AddRange(New Control() {_rbT23, _rbT13, _rbLdos, _rbT62, _rbLs63, lblModel, _rbModel1, _rbModel3})

        _chkSystem.Location = New Point(14, 170)
        _chkMinimal.Location = New Point(300, 170)

        Dim g3 As New GroupBox() With {.Text = "Density", .Bounds = New Rectangle(10, 200, 140, 76)}
        _rbSd.Location = New Point(14, 22) : _rbDd.Location = New Point(14, 46)
        g3.Controls.AddRange(New Control() {_rbSd, _rbDd})

        Dim g4 As New GroupBox() With {.Text = "Sides", .Bounds = New Rectangle(160, 200, 100, 76)}
        _rbSs.Location = New Point(14, 22) : _rbDs.Location = New Point(14, 46)
        g4.Controls.AddRange(New Control() {_rbSs, _rbDs})

        Dim g5 As New GroupBox() With {.Text = "Number of tracks (cylinders)", .Bounds = New Rectangle(270, 200, 190, 100)}
        _rb35.Location = New Point(14, 22) : _rb40.Location = New Point(70, 22) : _rb80.Location = New Point(126, 22)
        _rbOther.Location = New Point(14, 52) : _txtTracks.Location = New Point(80, 50)
        g5.Controls.AddRange(New Control() {_rb35, _rb40, _rb80, _rbOther, _txtTracks})

        Dim g6 As New GroupBox() With {.Text = "Volume information", .Bounds = New Rectangle(10, 285, 250, 90)}
        g6.Controls.Add(New Label() With {.Text = "Name (8 chars):", .AutoSize = True, .Location = New Point(12, 26)})
        _txtName.Location = New Point(120, 22)
        g6.Controls.Add(New Label() With {.Text = "Date (YYYY-MM-DD):", .AutoSize = True, .Location = New Point(12, 56)})
        _txtDate.Location = New Point(120, 52)
        g6.Controls.AddRange(New Control() {_txtName, _txtDate})

        Dim btnOk As New Button() With {.Text = If(_formatMode, "Format", "OK"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(290, 385, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(380, 385, 80, 26)}
        AddHandler btnOk.Click, AddressOf OnOk

        Controls.AddRange(New Control() {g1, g2, _chkSystem, _chkMinimal, g3, g4, g5, g6, btnOk, btnCancel})
        If _formatMode Then
            ' format always produces a DMK and never a system disk; the open image keeps its name.
            g1.Visible = False
            _chkSystem.Visible = False
            _chkMinimal.Visible = False
            _chkSystem.Checked = False
            Dim note As New Label() With {.Text = "Erases every file on the open image and formats it as an empty data disk (the image stays a DMK).", .AutoSize = False, .Bounds = New Rectangle(10, 10, 140, 100)}
            Controls.Add(note)
        End If
        AcceptButton = btnOk
        CancelButton = btnCancel

        For Each rb In New RadioButton() {_rbJv1, _rbJv3, _rbDmk, _rbT23, _rbT13, _rbLdos, _rbT62, _rbLs63, _rbModel1, _rbModel3, _rbSd, _rbDd, _rbSs, _rbDs, _rb35, _rb40, _rb80, _rbOther}
            AddHandler rb.CheckedChanged, Sub() ApplyInterlocks()
        Next
        AddHandler _chkSystem.CheckedChanged, Sub() ApplyInterlocks()
    End Sub

    Private Sub ApplyInterlocks()
        If _updating Then Return
        _updating = True
        Try
            Dim jv1 As Boolean = _rbJv1.Checked
            Dim t23 As Boolean = _rbT23.Checked
            Dim t13 As Boolean = _rbT13.Checked
            Dim ldos As Boolean = _rbLdos.Checked
            Dim m4 As Boolean = _rbT62.Checked OrElse _rbLs63.Checked
            Dim sys As Boolean = _chkSystem.Checked

            ' JV1 holds single-density, single-sided disks only (TRSDOS 1.3 is double density).
            _rbT13.Enabled = Not jv1
            If jv1 AndAlso t13 Then _rbLdos.Checked = True : t13 = False : ldos = True
            If jv1 Then _rbSd.Checked = True : _rbSs.Checked = True
            If t23 Then _rbSd.Checked = True : _rbSs.Checked = True : _rb35.Checked = True
            If t13 Then _rbDd.Checked = True : _rbSs.Checked = True : _rb40.Checked = True
            If m4 Then _rbDd.Checked = True
            If ldos Then
                If _rbModel1.Checked Then _rbSd.Checked = True Else _rbDd.Checked = True
            End If

            _rbSd.Enabled = Not sys AndAlso Not t13 AndAlso Not m4 AndAlso Not ldos
            _rbDd.Enabled = Not sys AndAlso Not jv1 AndAlso Not t23 AndAlso Not ldos
            _rbSs.Enabled = Not sys
            _rbDs.Enabled = Not sys AndAlso Not jv1 AndAlso Not t23 AndAlso Not t13
            _rb35.Enabled = Not sys AndAlso Not t13
            _rb40.Enabled = Not sys AndAlso Not t23
            _rb80.Enabled = Not sys AndAlso Not t23 AndAlso Not t13
            _rbOther.Enabled = Not sys AndAlso Not t23 AndAlso Not t13
            _txtTracks.Enabled = _rbOther.Enabled AndAlso _rbOther.Checked
            _rbModel1.Enabled = ldos : _rbModel3.Enabled = ldos
            _rbDmk.Enabled = True
            _chkMinimal.Enabled = sys
            If Not sys Then _chkMinimal.Checked = False
        Finally
            _updating = False
        End Try
    End Sub

    ''' <summary>True when the dialog is driving "disk format" rather than "disk create".</summary>
    Public ReadOnly Property IsFormatMode As Boolean
        Get
            Return _formatMode
        End Get
    End Property

    Private Sub OnOk(sender As Object, e As EventArgs)
        Dim nm As String = _txtName.Text.Trim()
        If nm.Length = 0 OrElse _txtDate.Text.Trim().Length = 0 Then
            MessageBox.Show(Me, "Volume name and date are required!", "Error creating floppy image", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DialogResult = DialogResult.None
            Return
        End If
        If Not System.Text.RegularExpressions.Regex.IsMatch(nm, "^[A-Za-z][A-Za-z0-9]{0,7}$") Then
            MessageBox.Show(Me, "The volume name must be 1–8 letters and digits, starting with a letter.", "Error creating floppy image", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DialogResult = DialogResult.None
            Return
        End If
        If TextUtil.ToYmd(_txtDate.Text) Is Nothing Then
            MessageBox.Show(Me, "The date must be a valid date, preferably YYYY-MM-DD.", "Error creating floppy image", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DialogResult = DialogResult.None
            Return
        End If
        If _rbOther.Checked Then
            Dim n As Integer
            If Not Integer.TryParse(_txtTracks.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, n) OrElse n < 1 OrElse n > 255 Then
                MessageBox.Show(Me, "Track count is not valid!", "Error creating floppy image", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DialogResult = DialogResult.None
                Return
            End If
        End If
    End Sub
End Class
