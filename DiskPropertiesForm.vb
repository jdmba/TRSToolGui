' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskPropertiesForm.vb — "Virtual disk properties": volume name/date/auto
'  command (editable on writable TRSDOS-family disks), capacity, free space,
'  DOS, geometry, damage, file name, image format, PDRIVE, write-protect flag,
'  plus trs80-tool's own detail lines and a Check button.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Public Class DiskPropertiesForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txtName As New TextBox()
    Private ReadOnly _txtDate As New TextBox()
    Private ReadOnly _txtAuto As New TextBox()
    Private ReadOnly _chkNoAuto As New CheckBox()
    Private ReadOnly _chkWp As New CheckBox()
    Private ReadOnly _chkSetMaster As New CheckBox()
    Private ReadOnly _txtMaster As New TextBox()
    Private ReadOnly _txtDetails As New TextBox()
    Private ReadOnly _btnCheck As New Button()

    ''' <summary>Every label-and-value row, so the value column can clear the widest label (UiFit.vb).</summary>
    Private ReadOnly _rows As New List(Of KeyValuePair(Of Control, Control))()
    Private _wpOriginal As Boolean

    Public Sub New(disk As DiskImage)
        Icon = Program.WindowIcon
        _disk = disk
        BuildUi()
    End Sub

    ''' <summary>Option words for "disk label" for whatever the user changed; empty if nothing.</summary>
    Public Function LabelOptionWords() As List(Of String)
        Dim w As New List(Of String)
        If Not _disk.IsWritable OrElse Not _disk.IsTrsdosFamily Then Return w
        Dim newName As String = _txtName.Text.Trim().ToUpperInvariant()
        If _txtName.Enabled AndAlso newName.Length > 0 AndAlso Not String.Equals(newName, _disk.VolumeName, StringComparison.Ordinal) Then
            w.Add("--name") : w.Add(newName)
        End If
        Dim dateText As String = _txtDate.Text.Trim()
        If _txtDate.Enabled AndAlso dateText.Length > 0 AndAlso Not String.Equals(dateText, _disk.VolumeDate, StringComparison.Ordinal) Then
            Dim ymd As String = TextUtil.ToYmd(dateText)
            If ymd IsNot Nothing Then
                w.Add("--date") : w.Add(ymd)
            End If
        End If
        If _chkSetMaster.Enabled AndAlso _chkSetMaster.Checked Then
            w.Add("--password") : w.Add(_txtMaster.Text.Trim())
        End If
        If _chkNoAuto.Enabled AndAlso _chkNoAuto.Checked AndAlso _disk.AutoCommand.Length > 0 Then
            w.Add("--no-auto")
        ElseIf _txtAuto.Enabled AndAlso Not _chkNoAuto.Checked AndAlso Not String.Equals(_txtAuto.Text.Trim(), _disk.AutoCommand, StringComparison.Ordinal) Then
            If _txtAuto.Text.Trim().Length = 0 Then
                If _disk.AutoCommand.Length > 0 Then w.Add("--no-auto")
            Else
                w.Add("--auto") : w.Add(_txtAuto.Text.Trim())
            End If
        End If
        Return w
    End Function

    Public ReadOnly Property WriteProtectChanged As Boolean
        Get
            Return _chkWp.Enabled AndAlso _chkWp.Checked <> _wpOriginal
        End Get
    End Property

    Public ReadOnly Property WriteProtectNew As Boolean
        Get
            Return _chkWp.Checked
        End Get
    End Property

    Private Sub BuildUi()
        Text = UiText.T("dlg.diskprops.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(520, 600)

        Dim y As Integer = 10
        Dim lblDateHint As Label = Nothing
        Dim lblMasterHint As Label = Nothing
        Dim grp1 As New GroupBox() With {.Left = 10, .Top = y, .Width = 500}
        grp1.Text = If(_disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing, UiText.T("dlg.diskprops.partition", _disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture), _disk.Partitions.Count.ToString(CultureInfo.InvariantCulture)), UiText.T("dlg.diskprops.volume"))
        Dim gy As Integer = 22
        Dim editable As Boolean = _disk.IsWritable AndAlso _disk.IsTrsdosFamily
        If _disk.IsTrsdosFamily Then
            AddRow(grp1, gy, UiText.T("dlg.diskprops.name"), _txtName, 120)
            _txtName.MaxLength = 8
            _txtName.Text = _disk.VolumeName
            _txtName.Enabled = editable
            gy += 26
            AddRow(grp1, gy, UiText.T("dlg.diskprops.date"), _txtDate, 160)
            _txtDate.Text = _disk.VolumeDate
            _txtDate.Enabled = editable
            lblDateHint = AddLabel(grp1, 300, gy + 3, UiText.T("dlg.diskprops.datehint"))
            gy += 26
            AddRow(grp1, gy, UiText.T("dlg.diskprops.auto"), _txtAuto, 260)
            _txtAuto.Text = _disk.AutoCommand
            _txtAuto.Enabled = editable
            _chkNoAuto.Text = UiText.T("dlg.diskprops.autoremove") : _chkNoAuto.AutoSize = True
            _chkNoAuto.Location = New Point(400, gy + 2)
            _chkNoAuto.Enabled = editable AndAlso _disk.AutoCommand.Length > 0
            grp1.Controls.Add(_chkNoAuto)
            gy += 26
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.master"), If(_disk.MasterPassword.Length > 0, _disk.MasterPassword, UiText.T("dlg.diskprops.masterunknown"))) : gy += 22
            _chkSetMaster.Text = UiText.T("dlg.diskprops.setmaster") : _chkSetMaster.AutoSize = True
            _chkSetMaster.Location = New Point(12, gy + 2)
            _chkSetMaster.Enabled = editable
            _txtMaster.Bounds = New Rectangle(160, gy, 120, 23)
            _txtMaster.MaxLength = 8
            _txtMaster.Enabled = False
            AddHandler _chkSetMaster.CheckedChanged, Sub() _txtMaster.Enabled = _chkSetMaster.Checked
            grp1.Controls.Add(_chkSetMaster) : grp1.Controls.Add(_txtMaster)
            _rows.Add(New KeyValuePair(Of Control, Control)(_chkSetMaster, _txtMaster))
            lblMasterHint = AddLabel(grp1, 300, gy + 3, UiText.T("dlg.diskprops.masterhint"))
            gy += 28
        End If
        If _disk.ApproxCapacityBytes > 0 Then AddStatic(grp1, gy, UiText.T("dlg.diskprops.capacity"), UiText.T("dlg.diskprops.capacityvalue", TextUtil.FormatKb(_disk.ApproxCapacityBytes), TextUtil.FormatBytes(_disk.ApproxCapacityBytes))) : gy += 22
        AddStatic(grp1, gy, UiText.T("dlg.diskprops.free"), If(_disk.FreeText.Length > 0, _disk.FreeText, UiText.T("dlg.diskprops.freeunknown"))) : gy += 22
        AddStatic(grp1, gy, UiText.T("dlg.diskprops.dos"), If(_disk.DosName.Length > 0, _disk.DosName, UiText.T("dlg.diskprops.dosnone"))) : gy += 22
        If _disk.Geometry IsNot Nothing Then
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.cylinders"), _disk.Geometry.CylinderCount.ToString(CultureInfo.InvariantCulture)) : gy += 22
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.sides"), If(_disk.AllocationSides > 0 AndAlso _disk.AllocationSides < _disk.Geometry.SideCount, UiText.T("dlg.diskprops.sidesvalue", _disk.Geometry.SideCount.ToString(CultureInfo.InvariantCulture), _disk.AllocationSides.ToString(CultureInfo.InvariantCulture)), _disk.Geometry.SideCount.ToString(CultureInfo.InvariantCulture))) : gy += 22
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.density"), _disk.DensityText) : gy += 22
        ElseIf _disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing Then
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.cylinders"), _disk.ActivePartition.CylinderCount.ToString(CultureInfo.InvariantCulture)) : gy += 22
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.heads"), _disk.ActivePartition.HeadCount.ToString(CultureInfo.InvariantCulture)) : gy += 22
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.location"), _disk.ActivePartition.Location) : gy += 22
        End If
        AddStatic(grp1, gy, UiText.T("dlg.diskprops.damaged"), If(_disk.Damaged, UiText.T("dlg.diskprops.damagedyes"), UiText.T("dlg.diskprops.damagedno"))) : gy += 22
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette AndAlso _disk.CheckText.Length > 0 Then
            AddStatic(grp1, gy, UiText.T("dlg.diskprops.protection"), ProtectionSummary()) : gy += 22
            If _disk.ImagingArtifacts.Count > 0 Then
                AddStatic(grp1, gy, UiText.T("dlg.diskprops.artifacts"), UiText.T("dlg.diskprops.artifactsvalue", _disk.ImagingArtifacts.Count.ToString(CultureInfo.InvariantCulture))) : gy += 22
            End If
        End If
        AddStatic(grp1, gy, UiText.T("dlg.diskprops.writable"), If(_disk.IsWritable, UiText.T("dlg.diskprops.writableyes"), UiText.T("dlg.diskprops.writableno", If(_disk.ReadOnlyReason, "")))) : gy += 22
        grp1.Height = gy + 8
        Controls.Add(grp1)
        y = grp1.Bottom + 10

        Dim grp2 As New GroupBox() With {.Left = 10, .Top = y, .Width = 500, .Text = UiText.T("dlg.diskprops.image")}
        gy = 22
        AddStatic(grp2, gy, UiText.T("dlg.diskprops.filename"), _disk.Path) : gy += 22
        AddStatic(grp2, gy, UiText.T("dlg.diskprops.format"), If(_disk.Description.Length > 0, UiText.T("dlg.diskprops.formatvalue", _disk.ImageFormat, _disk.Description), _disk.ImageFormat)) : gy += 22
        If _disk.PdriveText.Length > 0 Then
            Dim txtPd As New TextBox() With {.ReadOnly = True, .Left = 130, .Top = gy, .Width = 355, .Text = _disk.PdriveText}
            Dim lblPd As Label = AddLabel(grp2, 12, gy + 3, UiText.T("dlg.diskprops.pdrive"))
            grp2.Controls.Add(txtPd)
            _rows.Add(New KeyValuePair(Of Control, Control)(lblPd, txtPd))
            gy += 26
            Dim lblSrc As New Label() With {.AutoSize = False, .Location = New Point(130, gy), .Size = New Size(360, 30), .Text = UiText.T("dlg.diskprops.pdrivesource", _disk.PdriveSource), .ForeColor = SystemColors.GrayText}
            grp2.Controls.Add(lblSrc)
            gy += 32
        End If
        ' GUI-side identification from the bytes (DosIdentifier.vb), beside trs80-tool's "File system" line above
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette Then
            Dim fromBytes As String = IdentifyFromBytes()
            If fromBytes.Length > 0 Then
                Dim lblFb As Label = AddLabel(grp2, 12, gy + 3, UiText.T("dlg.diskprops.frombytes"))
                Dim txtId As New TextBox() With {.ReadOnly = True, .Left = 130, .Top = gy, .Width = 355, .Text = fromBytes}
                grp2.Controls.Add(txtId)
                _rows.Add(New KeyValuePair(Of Control, Control)(lblFb, txtId))
                gy += 26
            End If
        End If
        _chkWp.Text = UiText.T("dlg.diskprops.wp") : _chkWp.AutoSize = True
        _chkWp.Location = New Point(14, gy)
        If _disk.WriteProtected.HasValue Then
            _wpOriginal = _disk.WriteProtected.Value
            _chkWp.Checked = _wpOriginal
            _chkWp.Enabled = True
        Else
            _chkWp.Enabled = False
            _chkWp.Text = UiText.T("dlg.diskprops.wpna", _disk.ImageFormat)
        End If
        grp2.Controls.Add(_chkWp)
        gy += 26
        AddLabel(grp2, 12, gy, UiText.T("dlg.diskprops.details"))
        gy += 18
        With _txtDetails
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Vertical
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
            .Location = New Point(12, gy)
            .Size = New Size(476, 110)
            .Text = String.Join(vbCrLf, _disk.DetailLines)
        End With
        grp2.Controls.Add(_txtDetails)
        gy += 118
        grp2.Height = gy + 8
        Controls.Add(grp2)
        y = grp2.Bottom + 10

        _btnCheck.Text = UiText.T("btn.check")
        _btnCheck.Bounds = New Rectangle(10, y, 90, 26)
        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        Dim btnDmk As New Button() With {.Text = UiText.T("btn.dmkanalysis"), .Bounds = New Rectangle(106, y, 110, 26), .Enabled = Not _disk.IsHardDisk AndAlso Not _disk.IsCassette}
        AddHandler btnDmk.Click, Async Sub()
                                     Dim f As New DiskInfoForm(_disk.Path)
                                     UiFit.ShowOver(f, Me)
                                     Await f.LoadAsync()
                                 End Sub
        Controls.Add(btnDmk)
        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(340, y, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(430, y, 80, 26)}
        Controls.AddRange(New Control() {_btnCheck, btnOk, btnCancel})

        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        ' The value column clears the widest label, the two grey hints follow their boxes, and
        ' the groups and the window grow to hold what is now wider.
        UiFit.AlignColumn(_rows, 10)
        UiFit.SpreadRow(12, _txtDate, lblDateHint)
        UiFit.SpreadRow(12, _txtMaster, lblMasterHint)
        UiFit.SpreadRow(12, _txtAuto, _chkNoAuto)
        UiFit.GrowToFit(grp1)
        UiFit.GrowToFit(grp2)
        Dim wide As Integer = Math.Max(grp1.Width, grp2.Width)
        grp1.Width = wide : grp2.Width = wide
        _txtDetails.Width = Math.Max(_txtDetails.Width, grp2.ClientSize.Width - 24)
        UiFit.GrowForm(Me)
        UiFit.ButtonRowLeft(10, 6, _btnCheck, btnDmk)

        AcceptButton = btnOk
        CancelButton = btnCancel
        ' The height is known only now; the width is whatever the layout pass needed, never less
        ' than the 520 this window has always been.
        ClientSize = New Size(Math.Max(520, ClientSize.Width), y + 36)
        UiFit.ButtonRow(ClientSize.Width - 10, 10, btnOk, btnCancel)
    End Sub

    ''' <summary>Adds a label and gives it back, so the layout pass can measure it.</summary>
    Private Function AddLabel(parent As Control, x As Integer, y As Integer, text As String) As Label
        Dim lbl As New Label() With {.AutoSize = True, .Location = New Point(x, y), .Text = text}
        parent.Controls.Add(lbl)
        Return lbl
    End Function

    Private Sub AddRow(parent As Control, y As Integer, label As String, box As TextBox, width As Integer)
        Dim lbl As Label = AddLabel(parent, 12, y + 3, label)
        box.Location = New Point(130, y)
        box.Width = width
        parent.Controls.Add(box)
        _rows.Add(New KeyValuePair(Of Control, Control)(lbl, box))
    End Sub

    Private Sub AddStatic(parent As Control, y As Integer, label As String, value As String)
        Dim lbl As Label = AddLabel(parent, 12, y, label)
        Dim val As New Label() With {.AutoSize = False, .Location = New Point(130, y), .Size = New Size(360, 18), .Text = value, .AutoEllipsis = True}
        parent.Controls.Add(val)
        _rows.Add(New KeyValuePair(Of Control, Control)(lbl, val))
    End Sub

    ''' <summary>DosIdentifier's one-line result for the Disk properties dialog; "" when the file isn't a DMK or can't be read.</summary>
    Private Function IdentifyFromBytes() As String
        Try
            Dim img As DmkImage = DmkImage.Load(_disk.Path)
            If img.Kind <> "DMK" Then Return UiText.T("dlg.diskprops.notdmk", img.Kind)
            Return DosIdentifier.Identify(img).Result
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>One line on the signs of copy protection trs80-tool found (the Check button and the sector map list them).</summary>
    Private Function ProtectionSummary() As String
        If _disk.ProtectionSigns.Count = 0 Then Return UiText.T("dlg.diskprops.prot.none")
        Dim names As New List(Of String)
        For Each sign In _disk.ProtectionSigns
            Dim name As String = If(String.IsNullOrEmpty(sign.Scheme), "", sign.Scheme)
            If name.Length = 0 AndAlso sign.Schemes IsNot Nothing AndAlso sign.Schemes.Count > 0 Then name = String.Join("/", sign.Schemes)
            If name.Length > 0 AndAlso Not names.Contains(name) Then names.Add(name)
        Next
        Dim n As Integer = _disk.ProtectionSigns.Count
        Dim nText As String = n.ToString(CultureInfo.InvariantCulture)
        Dim text As String = If(n = 1, UiText.T("dlg.diskprops.prot.one", nText), UiText.T("dlg.diskprops.prot.many", nText))
        If _disk.ProtectedSignCount = 0 Then text = UiText.T("dlg.diskprops.prot.notconclusive", text)
        If names.Count > 0 Then text = UiText.T("dlg.diskprops.prot.withschemes", text, String.Join("; ", names))
        Return text
    End Function

    Private Async Function RunCheckAsync() As Threading.Tasks.Task
        _btnCheck.Enabled = False
        UseWaitCursor = True
        Try
            Dim r As ToolResult = Await Program.Tool.DiskCheckAsync(_disk.Path, _disk.PartitionArg)
            Dim report As String = If(r.Succeeded, r.StdOut, UiText.T("dlg.diskprops.checkfailed", r.ErrorText))
            _txtDetails.Text = report.Replace(vbCr, "").Replace(vbLf, vbCrLf)
        Finally
            UseWaitCursor = False
            _btnCheck.Enabled = True
        End Try
    End Function
End Class
