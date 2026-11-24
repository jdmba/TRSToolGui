' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HardDiskForms.vb — dialogs for trs80-tool's hard disk commands:
'    CreateHardDiskForm  File → New hard disk image  ("disk NEW.hdv create ...")
'    PartitionAddForm    View → Partition → Add a partition ("disk IMG partition --add DOS ...")
'  Both only collect the option words; trs80-tool does the work and checks
'  the numbers, and its own error text is shown when it refuses.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms

''' <summary>Options for a new HDV hard disk image: shape, and optionally a DOS with equal partitions.</summary>
Public Class CreateHardDiskForm
    Inherits Form

    Private ReadOnly _numCyl As New NumericUpDown() With {.Minimum = 1, .Maximum = 1023, .Value = 840, .Width = 70}
    Private ReadOnly _numHeads As New NumericUpDown() With {.Minimum = 1, .Maximum = 8, .Value = 6, .Width = 50}
    Private ReadOnly _rbBlank As New RadioButton() With {.Text = UiText.T("dlg.newhd.blank"), .AutoSize = True}
    Private ReadOnly _rbLdos As New RadioButton() With {.Text = UiText.T("dlg.newhd.ldos"), .AutoSize = True}
    Private ReadOnly _rbLsdos As New RadioButton() With {.Text = UiText.T("dlg.newhd.lsdos"), .AutoSize = True, .Checked = True}
    Private ReadOnly _chkPartitions As New CheckBox() With {.Text = UiText.T("dlg.newhd.parts"), .AutoSize = True}
    Private ReadOnly _numPartitions As New NumericUpDown() With {.Minimum = 1, .Maximum = 26, .Value = 4, .Width = 50, .Enabled = False}
    Private ReadOnly _cmbModel As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 90}
    Private ReadOnly _txtName As New TextBox() With {.Width = 90, .MaxLength = 8, .CharacterCasing = CharacterCasing.Upper}
    Private ReadOnly _lblModel As New Label() With {.Text = UiText.T("dlg.newhd.model"), .AutoSize = True}
    Private ReadOnly _lblName As New Label() With {.Text = UiText.T("dlg.newhd.name"), .AutoSize = True}

    Public Sub New()
        Text = UiText.T("dlg.newhd.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 330)

        Dim g1 As New GroupBox() With {.Text = UiText.T("dlg.newhd.shape"), .Bounds = New Rectangle(10, 10, 540, 62)}
        Dim lblCyl As New Label() With {.Text = UiText.T("dlg.newhd.cylinders"), .AutoSize = True, .Location = New Point(14, 28)}
        g1.Controls.Add(lblCyl)
        _numCyl.Location = New Point(140, 25)
        g1.Controls.Add(_numCyl)
        Dim lblHeads As New Label() With {.Text = UiText.T("dlg.newhd.heads"), .AutoSize = True, .Location = New Point(250, 28)}
        g1.Controls.Add(lblHeads)
        _numHeads.Location = New Point(330, 25)
        g1.Controls.Add(_numHeads)

        Dim g2 As New GroupBox() With {.Text = UiText.T("dlg.newhd.contents"), .Bounds = New Rectangle(10, 80, 540, 196)}
        _rbLsdos.Location = New Point(14, 24)
        _rbLdos.Location = New Point(14, 48)
        _rbBlank.Location = New Point(14, 72)
        _chkPartitions.Location = New Point(34, 104)
        _numPartitions.Location = New Point(220, 102)
        Dim lblPartsHint As New Label() With {.Text = UiText.T("dlg.newhd.partshint"), .AutoSize = True, .Location = New Point(280, 105), .ForeColor = SystemColors.GrayText}
        g2.Controls.Add(lblPartsHint)
        _lblModel.Location = New Point(34, 134)
        _cmbModel.Location = New Point(220, 130)
        _cmbModel.Items.AddRange(New Object() {UiText.T("word.model3"), UiText.T("word.model1")})
        _cmbModel.SelectedIndex = 0
        _lblName.Location = New Point(34, 162)
        _txtName.Location = New Point(220, 158)
        Dim lblNameHint As New Label() With {.Text = UiText.T("dlg.newhd.namehint"), .AutoSize = True, .Location = New Point(320, 162), .ForeColor = SystemColors.GrayText}
        g2.Controls.Add(lblNameHint)
        g2.Controls.AddRange(New Control() {_rbLsdos, _rbLdos, _rbBlank, _chkPartitions, _numPartitions, _lblModel, _cmbModel, _lblName, _txtName})

        Dim btnOk As New Button() With {.Text = UiText.T("btn.create"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(380, 292, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(470, 292, 80, 26)}
        Controls.AddRange(New Control() {g1, g2, btnOk, btnCancel})

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.AfterAll(lblCyl, 8, _numCyl)
        UiFit.SpreadRow(16, _numCyl, lblHeads)
        UiFit.AfterAll(lblHeads, 8, _numHeads)
        UiFit.SpreadRow(12, _chkPartitions, _numPartitions, lblPartsHint)
        UiFit.AfterAll(_lblModel, 8, _cmbModel)
        UiFit.AfterAll(_lblName, 8, _txtName)
        UiFit.SpreadRow(12, _txtName, lblNameHint)
        UiFit.GrowToFit(g1)
        UiFit.GrowToFit(g2)
        g1.Width = Math.Max(g1.Width, g2.Width)
        g2.Width = Math.Max(g1.Width, g2.Width)
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel

        AddHandler _rbBlank.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _rbLdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _rbLsdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _chkPartitions.CheckedChanged, Sub() UpdateEnabled()
        UpdateEnabled()
    End Sub

    Private Sub UpdateEnabled()
        Dim withDos As Boolean = Not _rbBlank.Checked
        _chkPartitions.Enabled = withDos
        _numPartitions.Enabled = withDos AndAlso _chkPartitions.Checked
        _cmbModel.Enabled = _rbLdos.Checked
        _lblModel.Enabled = _rbLdos.Checked
        _txtName.Enabled = withDos
        _lblName.Enabled = withDos
    End Sub

    ''' <summary>The DOS for "create" ("ldos-5.3" or "lsdos-6.3"), or Nothing for a blank disk.</summary>
    Public ReadOnly Property DosName As String
        Get
            If _rbBlank.Checked Then Return Nothing
            Return If(_rbLdos.Checked, "ldos-5.3", "lsdos-6.3")
        End Get
    End Property

    ''' <summary>The option words for "create": --cylinders N --heads N [--partitions N] [--model 1|3] [--name X].</summary>
    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String) From {
            "--cylinders", CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture),
            "--heads", CInt(_numHeads.Value).ToString(CultureInfo.InvariantCulture)}
        If Not _rbBlank.Checked Then
            If _chkPartitions.Checked Then
                w.Add("--partitions") : w.Add(CInt(_numPartitions.Value).ToString(CultureInfo.InvariantCulture))
            End If
            If _rbLdos.Checked Then
                w.Add("--model") : w.Add(If(_cmbModel.SelectedIndex = 1, "1", "3"))
            End If
            Dim name As String = _txtName.Text.Trim()
            If name.Length > 0 Then
                w.Add("--name") : w.Add(name)
            End If
        End If
        Return w
    End Function
End Class

''' <summary>Options for a new partition in a hard disk's free space.</summary>
Public Class PartitionAddForm
    Inherits Form

    Private ReadOnly _rbLdos As New RadioButton() With {.Text = UiText.T("dlg.partadd.ldos"), .AutoSize = True}
    Private ReadOnly _rbLsdos As New RadioButton() With {.Text = UiText.T("dos.lsdos63.m4"), .AutoSize = True, .Checked = True}
    Private ReadOnly _cmbModel As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 90}
    Private ReadOnly _chkCyl As New CheckBox() With {.Text = UiText.T("dlg.partadd.cylinders"), .AutoSize = True}
    Private ReadOnly _numCyl As New NumericUpDown() With {.Minimum = 35, .Maximum = 203, .Value = 203, .Width = 60, .Enabled = False}
    Private ReadOnly _txtName As New TextBox() With {.Width = 90, .MaxLength = 8, .CharacterCasing = CharacterCasing.Upper}
    Private ReadOnly _lblModel As New Label() With {.Text = UiText.T("dlg.partadd.model"), .AutoSize = True}

    Public Sub New()
        Text = UiText.T("dlg.partadd.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(480, 232)

        Controls.Add(New Label() With {.Text = UiText.TCrLf("dlg.partadd.intro"), .AutoSize = True, .Location = New Point(12, 10)})
        _rbLsdos.Location = New Point(14, 52)
        _rbLdos.Location = New Point(14, 76)
        _lblModel.Location = New Point(120, 78)
        _cmbModel.Location = New Point(220, 74)
        _cmbModel.Items.AddRange(New Object() {UiText.T("word.model3"), UiText.T("word.model1")})
        _cmbModel.SelectedIndex = 0
        _chkCyl.Location = New Point(14, 112)
        _numCyl.Location = New Point(160, 110)
        Dim lblCylHint As New Label() With {.Text = UiText.T("dlg.partadd.cylhint"), .AutoSize = True, .Location = New Point(230, 113), .ForeColor = SystemColors.GrayText}
        Controls.Add(lblCylHint)
        Dim lblName As New Label() With {.Text = UiText.T("dlg.partadd.name"), .AutoSize = True, .Location = New Point(14, 148)}
        Controls.Add(lblName)
        _txtName.Location = New Point(160, 144)
        Dim lblNameHint As New Label() With {.Text = UiText.T("dlg.partadd.namehint"), .AutoSize = True, .Location = New Point(260, 148), .ForeColor = SystemColors.GrayText}
        Controls.Add(lblNameHint)
        Dim btnOk As New Button() With {.Text = UiText.T("btn.add"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(300, 194, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(390, 194, 80, 26)}
        Controls.AddRange(New Control() {_rbLsdos, _rbLdos, _lblModel, _cmbModel, _chkCyl, _numCyl, _txtName, btnOk, btnCancel})

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.SpreadRow(12, _rbLdos, _lblModel)
        UiFit.AfterAll(_lblModel, 8, _cmbModel)
        UiFit.SpreadRow(12, _chkCyl, _numCyl, lblCylHint)
        UiFit.AfterAll(lblName, 8, _txtName)
        UiFit.SpreadRow(12, _txtName, lblNameHint)
        UiFit.GrowForm(Me)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)

        AcceptButton = btnOk : CancelButton = btnCancel

        AddHandler _rbLdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _chkCyl.CheckedChanged, Sub() UpdateEnabled()
        UpdateEnabled()
    End Sub

    Private Sub UpdateEnabled()
        _cmbModel.Enabled = _rbLdos.Checked
        _lblModel.Enabled = _rbLdos.Checked
        _numCyl.Enabled = _chkCyl.Checked
    End Sub

    ''' <summary>The DOS for "partition --add": "ldos-5.3" or "lsdos-6.3".</summary>
    Public ReadOnly Property DosName As String
        Get
            Return If(_rbLdos.Checked, "ldos-5.3", "lsdos-6.3")
        End Get
    End Property

    ''' <summary>The option words after "partition --add DOS": [--cylinders N] [--name X] [--model 1|3].</summary>
    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _chkCyl.Checked Then
            w.Add("--cylinders") : w.Add(CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture))
        End If
        Dim name As String = _txtName.Text.Trim()
        If name.Length > 0 Then
            w.Add("--name") : w.Add(name)
        End If
        If _rbLdos.Checked Then
            w.Add("--model") : w.Add(If(_cmbModel.SelectedIndex = 1, "1", "3"))
        End If
        Return w
    End Function
End Class
