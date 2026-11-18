' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  HardDiskForms.vb — hard disk dialogs of the Eto.Forms program:
'  CreateHardDiskForm (File → New hard disk image) and PartitionAddForm
'  (View → Partition → Add a partition). Each hands back a DOS name and the
'  option words for trs80-tool.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports Eto.Drawing
Imports Eto.Forms

''' <summary>File → New hard disk image: the words for "disk NEW.hdv create [DOS] ...".</summary>
Public Class CreateHardDiskForm
    Inherits OkCancelDialog

    Private ReadOnly _numCyl As New NumericStepper With {.MinValue = 1, .MaxValue = 1023, .Value = 840, .DecimalPlaces = 0, .Width = 90}
    Private ReadOnly _numHeads As New NumericStepper With {.MinValue = 1, .MaxValue = 8, .Value = 6, .DecimalPlaces = 0, .Width = 70}
    Private ReadOnly _rbBlank, _rbLdos, _rbLsdos As RadioButton
    Private ReadOnly _chkPartitions As New CheckBox With {.Text = "Number of equal partitions:"}
    Private ReadOnly _numPartitions As New NumericStepper With {.MinValue = 1, .MaxValue = 26, .Value = 4, .DecimalPlaces = 0, .Width = 70, .Enabled = False}
    Private ReadOnly _cmbModel As New DropDown With {.Width = 110}
    Private ReadOnly _txtName As New TextBox With {.Width = 110, .MaxLength = 8}
    Private ReadOnly _lblModel As New Label With {.Text = "LDOS boot sector for:", .VerticalAlignment = VerticalAlignment.Center}
    Private ReadOnly _lblName As New Label With {.Text = "Disk name (one partition only):", .VerticalAlignment = VerticalAlignment.Center}

    Public Sub New()
        MyBase.New("Create New Hard Disk Image")
        _rbLsdos = New RadioButton With {.Text = "LS-DOS 6.3 partitions (Model 4)", .Checked = True}
        _rbLdos = New RadioButton(_rbLsdos) With {.Text = "LDOS 5.3 partitions"}
        _rbBlank = New RadioButton(_rbLsdos) With {.Text = "Blank: no partitions (add them later with View → Partition → Add a partition)"}
        _cmbModel.Items.Add("Model III")
        _cmbModel.Items.Add("Model I")
        _cmbModel.SelectedIndex = 0
        Dim g1 As GroupBox = Ui.Group("Shape (the defaults are what FreHD makes)",
            Ui.Across(New Label With {.Text = "Cylinders (1-1023):", .VerticalAlignment = VerticalAlignment.Center}, _numCyl, New Label With {.Text = "    Heads (1-8):", .VerticalAlignment = VerticalAlignment.Center}, _numHeads))
        Dim inner As TableLayout = Ui.Fields(
            _chkPartitions, Ui.Across(_numPartitions, Ui.Lbl("(otherwise as few as keep each at 203 cylinders or fewer)", True)),
            _lblModel, Ui.Across(_cmbModel),
            _lblName, Ui.Across(_txtName, Ui.Lbl("(blank = HARDA, HARDB, ...)", True)))
        Dim g2 As GroupBox = Ui.Group("Contents", Ui.Down(Nothing, _rbLsdos, _rbLdos, _rbBlank, New Panel With {.Padding = New Padding(20, 4, 0, 0), .Content = inner}))
        Finish(Ui.Down(Nothing, g1, g2), "Create...")
        AddHandler _rbBlank.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _rbLdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _rbLsdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _chkPartitions.CheckedChanged, Sub() UpdateEnabled()
        UpdateEnabled()
    End Sub

    Private Sub UpdateEnabled()
        Dim withDos As Boolean = Not _rbBlank.Checked
        _chkPartitions.Enabled = withDos
        _numPartitions.Enabled = withDos AndAlso Ui.IsOn(_chkPartitions)
        _cmbModel.Enabled = _rbLdos.Checked
        _lblModel.Enabled = _rbLdos.Checked
        _txtName.Enabled = withDos
        _lblName.Enabled = withDos
    End Sub

    ''' <summary>The DOS word for "create", or Nothing for a blank image.</summary>
    Public ReadOnly Property DosName As String
        Get
            If _rbBlank.Checked Then Return Nothing
            Return If(_rbLdos.Checked, "ldos-5.3", "lsdos-6.3")
        End Get
    End Property

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String) From {
            "--cylinders", CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture),
            "--heads", CInt(_numHeads.Value).ToString(CultureInfo.InvariantCulture)}
        If Not _rbBlank.Checked Then
            If Ui.IsOn(_chkPartitions) Then
                w.Add("--partitions") : w.Add(CInt(_numPartitions.Value).ToString(CultureInfo.InvariantCulture))
            End If
            If _rbLdos.Checked Then
                w.Add("--model") : w.Add(If(_cmbModel.SelectedIndex = 1, "1", "3"))
            End If
            Dim name As String = _txtName.Text.Trim().ToUpperInvariant()
            If name.Length > 0 Then
                w.Add("--name") : w.Add(name)
            End If
        End If
        Return w
    End Function
End Class

''' <summary>View → Partition → Add a partition: the words for "disk IMG partition --add DOS ...".</summary>
Public Class PartitionAddForm
    Inherits OkCancelDialog

    Private ReadOnly _rbLdos, _rbLsdos As RadioButton
    Private ReadOnly _cmbModel As New DropDown With {.Width = 110}
    Private ReadOnly _chkCyl As New CheckBox With {.Text = "Cylinders (35-203):"}
    Private ReadOnly _numCyl As New NumericStepper With {.MinValue = 35, .MaxValue = 203, .Value = 203, .DecimalPlaces = 0, .Width = 80, .Enabled = False}
    Private ReadOnly _txtName As New TextBox With {.Width = 110, .MaxLength = 8}
    Private ReadOnly _lblModel As New Label With {.Text = "    Boot sector for:", .VerticalAlignment = VerticalAlignment.Center}

    Public Sub New()
        MyBase.New("Add a Partition")
        _rbLsdos = New RadioButton With {.Text = "LS-DOS 6.3 (Model 4)", .Checked = True}
        _rbLdos = New RadioButton(_rbLsdos) With {.Text = "LDOS 5.3"}
        _cmbModel.Items.Add("Model III")
        _cmbModel.Items.Add("Model I")
        _cmbModel.SelectedIndex = 0
        Finish(Ui.Down(Nothing,
                       Ui.Lbl("trs80-tool formats an empty data partition in the image's free space and prints the" & vbLf & "settings the DOS's hard disk driver must be given to find it."),
                       _rbLsdos,
                       Ui.Across(_rbLdos, _lblModel, _cmbModel),
                       Ui.Fields(_chkCyl, Ui.Across(_numCyl, Ui.Lbl("(otherwise the rest of the free block, at most 203)", True)),
                                 "Disk name:", Ui.Across(_txtName, Ui.Lbl("(blank = HARD and the first free letter)", True)))), "Add")
        AddHandler _rbLdos.CheckedChanged, Sub() UpdateEnabled()
        AddHandler _chkCyl.CheckedChanged, Sub() UpdateEnabled()
        UpdateEnabled()
    End Sub

    Private Sub UpdateEnabled()
        _cmbModel.Enabled = _rbLdos.Checked
        _lblModel.Enabled = _rbLdos.Checked
        _numCyl.Enabled = Ui.IsOn(_chkCyl)
    End Sub

    Public ReadOnly Property DosName As String
        Get
            Return If(_rbLdos.Checked, "ldos-5.3", "lsdos-6.3")
        End Get
    End Property

    Public Function OptionWords() As List(Of String)
        Dim w As New List(Of String)
        If Ui.IsOn(_chkCyl) Then
            w.Add("--cylinders") : w.Add(CInt(_numCyl.Value).ToString(CultureInfo.InvariantCulture))
        End If
        Dim name As String = _txtName.Text.Trim().ToUpperInvariant()
        If name.Length > 0 Then
            w.Add("--name") : w.Add(name)
        End If
        If _rbLdos.Checked Then
            w.Add("--model") : w.Add(If(_cmbModel.SelectedIndex = 1, "1", "3"))
        End If
        Return w
    End Function
End Class
