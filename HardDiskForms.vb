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
    Private ReadOnly _rbBlank As New RadioButton() With {.Text = "Blank: no partitions (add them later with View → Partition → Add a partition)", .AutoSize = True}
    Private ReadOnly _rbLdos As New RadioButton() With {.Text = "LDOS 5.3 partitions", .AutoSize = True}
    Private ReadOnly _rbLsdos As New RadioButton() With {.Text = "LS-DOS 6.3 partitions (Model 4)", .AutoSize = True, .Checked = True}
    Private ReadOnly _chkPartitions As New CheckBox() With {.Text = "Number of equal partitions:", .AutoSize = True}
    Private ReadOnly _numPartitions As New NumericUpDown() With {.Minimum = 1, .Maximum = 26, .Value = 4, .Width = 50, .Enabled = False}
    Private ReadOnly _cmbModel As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 90}
    Private ReadOnly _txtName As New TextBox() With {.Width = 90, .MaxLength = 8, .CharacterCasing = CharacterCasing.Upper}
    Private ReadOnly _lblModel As New Label() With {.Text = "LDOS boot sector for:", .AutoSize = True}
    Private ReadOnly _lblName As New Label() With {.Text = "Disk name (one partition only):", .AutoSize = True}

    Public Sub New()
        Text = "Create New Hard Disk Image"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(560, 330)

        Dim g1 As New GroupBox() With {.Text = "Shape (the defaults are what FreHD makes)", .Bounds = New Rectangle(10, 10, 540, 62)}
        g1.Controls.Add(New Label() With {.Text = "Cylinders (1-1023):", .AutoSize = True, .Location = New Point(14, 28)})
        _numCyl.Location = New Point(140, 25)
        g1.Controls.Add(_numCyl)
        g1.Controls.Add(New Label() With {.Text = "Heads (1-8):", .AutoSize = True, .Location = New Point(250, 28)})
        _numHeads.Location = New Point(330, 25)
        g1.Controls.Add(_numHeads)

        Dim g2 As New GroupBox() With {.Text = "Contents", .Bounds = New Rectangle(10, 80, 540, 196)}
        _rbLsdos.Location = New Point(14, 24)
        _rbLdos.Location = New Point(14, 48)
        _rbBlank.Location = New Point(14, 72)
        _chkPartitions.Location = New Point(34, 104)
        _numPartitions.Location = New Point(220, 102)
        g2.Controls.Add(New Label() With {.Text = "(otherwise as few as keep each at 203 cylinders or fewer)", .AutoSize = True, .Location = New Point(280, 105), .ForeColor = SystemColors.GrayText})
        _lblModel.Location = New Point(34, 134)
        _cmbModel.Location = New Point(220, 130)
        _cmbModel.Items.AddRange(New Object() {"Model III", "Model I"})
        _cmbModel.SelectedIndex = 0
        _lblName.Location = New Point(34, 162)
        _txtName.Location = New Point(220, 158)
        g2.Controls.Add(New Label() With {.Text = "(blank = HARDA, HARDB, ...)", .AutoSize = True, .Location = New Point(320, 162), .ForeColor = SystemColors.GrayText})
        g2.Controls.AddRange(New Control() {_rbLsdos, _rbLdos, _rbBlank, _chkPartitions, _numPartitions, _lblModel, _cmbModel, _lblName, _txtName})

        Dim btnOk As New Button() With {.Text = "Create...", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(380, 292, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(470, 292, 80, 26)}
        Controls.AddRange(New Control() {g1, g2, btnOk, btnCancel})
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

    Private ReadOnly _rbLdos As New RadioButton() With {.Text = "LDOS 5.3", .AutoSize = True}
    Private ReadOnly _rbLsdos As New RadioButton() With {.Text = "LS-DOS 6.3 (Model 4)", .AutoSize = True, .Checked = True}
    Private ReadOnly _cmbModel As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 90}
    Private ReadOnly _chkCyl As New CheckBox() With {.Text = "Cylinders (35-203):", .AutoSize = True}
    Private ReadOnly _numCyl As New NumericUpDown() With {.Minimum = 35, .Maximum = 203, .Value = 203, .Width = 60, .Enabled = False}
    Private ReadOnly _txtName As New TextBox() With {.Width = 90, .MaxLength = 8, .CharacterCasing = CharacterCasing.Upper}
    Private ReadOnly _lblModel As New Label() With {.Text = "Boot sector for:", .AutoSize = True}

    Public Sub New()
        Text = "Add a Partition"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(480, 232)

        Controls.Add(New Label() With {.Text = "trs80-tool formats an empty data partition in the image's free space and prints the" & vbCrLf & "settings the DOS's hard disk driver must be given to find it.", .AutoSize = True, .Location = New Point(12, 10)})
        _rbLsdos.Location = New Point(14, 52)
        _rbLdos.Location = New Point(14, 76)
        _lblModel.Location = New Point(120, 78)
        _cmbModel.Location = New Point(220, 74)
        _cmbModel.Items.AddRange(New Object() {"Model III", "Model I"})
        _cmbModel.SelectedIndex = 0
        _chkCyl.Location = New Point(14, 112)
        _numCyl.Location = New Point(160, 110)
        Controls.Add(New Label() With {.Text = "(otherwise the rest of the free block, at most 203)", .AutoSize = True, .Location = New Point(230, 113), .ForeColor = SystemColors.GrayText})
        Controls.Add(New Label() With {.Text = "Disk name:", .AutoSize = True, .Location = New Point(14, 148)})
        _txtName.Location = New Point(160, 144)
        Controls.Add(New Label() With {.Text = "(blank = HARD and the first free letter)", .AutoSize = True, .Location = New Point(260, 148), .ForeColor = SystemColors.GrayText})
        Dim btnOk As New Button() With {.Text = "Add", .DialogResult = DialogResult.OK, .Bounds = New Rectangle(300, 194, 80, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(390, 194, 80, 26)}
        Controls.AddRange(New Control() {_rbLsdos, _rbLdos, _lblModel, _cmbModel, _chkCyl, _numCyl, _txtName, btnOk, btnCancel})
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
