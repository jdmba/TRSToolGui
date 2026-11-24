' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskPropertiesForm.vb — "Virtual disk properties" in the Eto.Forms program:
'  volume information (name, date, AUTO command, master password), what
'  trs80-tool reports about the disk, the write-protect flag, a Check button
'  (disk check) and the DMK analysis.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class DiskPropertiesForm
    Inherits OkCancelDialog

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _txtName As New TextBox With {.Width = 130, .MaxLength = 8}
    Private ReadOnly _txtDate As New TextBox With {.Width = 170}
    Private ReadOnly _txtAuto As New TextBox With {.Width = 280}
    Private ReadOnly _chkNoAuto As New CheckBox With {.Text = UiText.T("dlg.diskprops.autoremove")}
    Private ReadOnly _chkWp As New CheckBox With {.Text = UiText.T("dlg.diskprops.wp")}
    Private ReadOnly _chkSetMaster As New CheckBox With {.Text = UiText.T("dlg.diskprops.setmaster")}
    Private ReadOnly _txtMaster As New TextBox With {.Width = 120, .MaxLength = 8}
    Private _txtDetails As TextArea
    Private ReadOnly _btnCheck As New Button With {.Text = UiText.T("btn.check")}
    Private _wpOriginal As Boolean
    Private _labelEditable As Boolean
    Private _wpAvailable As Boolean

    Public Sub New(disk As DiskImage)
        MyBase.New(UiText.T("dlg.diskprops.title"))
        _disk = disk
        BuildUi()
    End Sub

    ''' <summary>Option words for "disk label" (only what changed). Empty = nothing to set.</summary>
    Public Function LabelOptionWords() As List(Of String)
        Dim w As New List(Of String)
        If Not _labelEditable Then Return w
        Dim newName As String = _txtName.Text.Trim().ToUpperInvariant()
        If newName.Length > 0 AndAlso Not String.Equals(newName, _disk.VolumeName, StringComparison.Ordinal) Then
            w.Add("--name") : w.Add(newName)
        End If
        Dim dateText As String = _txtDate.Text.Trim()
        If dateText.Length > 0 AndAlso Not String.Equals(dateText, _disk.VolumeDate, StringComparison.Ordinal) Then
            Dim ymd As String = TextUtil.ToYmd(dateText)
            If ymd IsNot Nothing Then
                w.Add("--date") : w.Add(ymd)
            End If
        End If
        If Ui.IsOn(_chkSetMaster) Then
            w.Add("--password") : w.Add(_txtMaster.Text.Trim())
        End If
        If _chkNoAuto.Enabled AndAlso Ui.IsOn(_chkNoAuto) AndAlso _disk.AutoCommand.Length > 0 Then
            w.Add("--no-auto")
        ElseIf Not Ui.IsOn(_chkNoAuto) AndAlso Not String.Equals(_txtAuto.Text.Trim(), _disk.AutoCommand, StringComparison.Ordinal) Then
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
            Return _wpAvailable AndAlso Ui.IsOn(_chkWp) <> _wpOriginal
        End Get
    End Property

    Public ReadOnly Property WriteProtectNew As Boolean
        Get
            Return Ui.IsOn(_chkWp)
        End Get
    End Property

    Private Sub BuildUi()
        Dim editable As Boolean = _disk.IsWritable AndAlso _disk.IsTrsdosFamily
        _labelEditable = editable

        ' ---- volume
        Dim v As New List(Of Object)
        If _disk.IsTrsdosFamily Then
            _txtName.Text = _disk.VolumeName
            _txtName.Enabled = editable
            v.Add(UiText.T("dlg.diskprops.name")) : v.Add(Ui.Across(_txtName))
            _txtDate.Text = _disk.VolumeDate
            _txtDate.Enabled = editable
            v.Add(UiText.T("dlg.diskprops.date")) : v.Add(Ui.Across(_txtDate, Ui.Lbl(UiText.T("dlg.diskprops.datehint"), True)))
            _txtAuto.Text = _disk.AutoCommand
            _txtAuto.Enabled = editable
            _chkNoAuto.Enabled = editable AndAlso _disk.AutoCommand.Length > 0
            v.Add(UiText.T("dlg.diskprops.auto")) : v.Add(Ui.Across(_txtAuto, _chkNoAuto))
            v.Add(UiText.T("dlg.diskprops.master")) : v.Add(Ui.Lbl(If(_disk.MasterPassword.Length > 0, _disk.MasterPassword, UiText.T("dlg.diskprops.masterunknown"))))
            _chkSetMaster.Enabled = editable
            _txtMaster.Enabled = False
            AddHandler _chkSetMaster.CheckedChanged, Sub() _txtMaster.Enabled = Ui.IsOn(_chkSetMaster)
            v.Add(_chkSetMaster) : v.Add(Ui.Across(_txtMaster, Ui.Lbl(UiText.T("dlg.diskprops.masterhint"), True)))
        End If
        If _disk.ApproxCapacityBytes > 0 Then v.Add(UiText.T("dlg.diskprops.capacity")) : v.Add(Ui.Lbl(UiText.T("dlg.diskprops.capacityvalue", TextUtil.FormatKb(_disk.ApproxCapacityBytes), TextUtil.FormatBytes(_disk.ApproxCapacityBytes))))
        v.Add(UiText.T("dlg.diskprops.free")) : v.Add(Ui.Lbl(If(_disk.FreeText.Length > 0, _disk.FreeText, UiText.T("dlg.diskprops.freeunknown"))))
        v.Add(UiText.T("dlg.diskprops.dos")) : v.Add(Ui.Lbl(If(_disk.DosName.Length > 0, _disk.DosName, UiText.T("dlg.diskprops.dosnone"))))
        If _disk.Geometry IsNot Nothing Then
            v.Add(UiText.T("dlg.diskprops.cylinders")) : v.Add(Ui.Lbl(_disk.Geometry.CylinderCount.ToString(CultureInfo.InvariantCulture)))
            v.Add(UiText.T("dlg.diskprops.sides")) : v.Add(Ui.Lbl(If(_disk.AllocationSides > 0 AndAlso _disk.AllocationSides < _disk.Geometry.SideCount, UiText.T("dlg.diskprops.sidesvalue", _disk.Geometry.SideCount.ToString(CultureInfo.InvariantCulture), _disk.AllocationSides.ToString(CultureInfo.InvariantCulture)), _disk.Geometry.SideCount.ToString(CultureInfo.InvariantCulture))))
            v.Add(UiText.T("dlg.diskprops.density")) : v.Add(Ui.Lbl(_disk.DensityText))
        ElseIf _disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing Then
            v.Add(UiText.T("dlg.diskprops.cylinders")) : v.Add(Ui.Lbl(_disk.ActivePartition.CylinderCount.ToString(CultureInfo.InvariantCulture)))
            v.Add(UiText.T("dlg.diskprops.heads")) : v.Add(Ui.Lbl(_disk.ActivePartition.HeadCount.ToString(CultureInfo.InvariantCulture)))
            v.Add(UiText.T("dlg.diskprops.location")) : v.Add(Ui.Lbl(_disk.ActivePartition.Location))
        End If
        v.Add(UiText.T("dlg.diskprops.damaged")) : v.Add(Ui.Lbl(If(_disk.Damaged, UiText.T("dlg.diskprops.damagedyes"), UiText.T("dlg.diskprops.damagedno"))))
        ' What "disk check" found beyond damage: signs of copy protection and imaging artifacts.
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette AndAlso _disk.CheckText.Length > 0 Then
            v.Add(UiText.T("dlg.diskprops.protection")) : v.Add(Ui.Lbl(ProtectionSummary()))
            If _disk.ImagingArtifacts.Count > 0 Then
                v.Add(UiText.T("dlg.diskprops.artifacts")) : v.Add(Ui.Lbl(UiText.T("dlg.diskprops.artifactsvalue", _disk.ImagingArtifacts.Count.ToString(CultureInfo.InvariantCulture))))
            End If
        End If
        v.Add(UiText.T("dlg.diskprops.writable")) : v.Add(Ui.Wrapped(If(_disk.IsWritable, UiText.T("dlg.diskprops.writableyes"), UiText.T("dlg.diskprops.writableno", If(_disk.ReadOnlyReason, ""))), 420))
        Dim grp1 As GroupBox = Ui.Group(If(_disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing, UiText.T("dlg.diskprops.partition", _disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture), _disk.Partitions.Count.ToString(CultureInfo.InvariantCulture)), UiText.T("dlg.diskprops.volume")), Ui.Fields(v.ToArray()))

        ' ---- image
        Dim d As New List(Of Object)
        d.Add(UiText.T("dlg.diskprops.filename")) : d.Add(New TextBox With {.ReadOnly = True, .Text = _disk.Path, .Width = 420})
        d.Add(UiText.T("dlg.diskprops.format")) : d.Add(Ui.Lbl(If(_disk.Description.Length > 0, UiText.T("dlg.diskprops.formatvalue", _disk.ImageFormat, _disk.Description), _disk.ImageFormat)))
        If _disk.PdriveText.Length > 0 Then
            ' NEWDOS/80: the PDRIVE line as trs80-tool gives it, to copy, and where its values come from.
            d.Add(UiText.T("dlg.diskprops.pdrive")) : d.Add(New TextBox With {.ReadOnly = True, .Text = _disk.PdriveText, .Width = 420})
            d.Add("") : d.Add(Ui.Wrapped(UiText.T("dlg.diskprops.pdrivesource", _disk.PdriveSource), 420))
        End If
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette Then
            ' The program's own identification from the boot sector, directory and GAT bytes.
            Dim fromBytes As String = IdentifyFromBytes()
            If fromBytes.Length > 0 Then d.Add(UiText.T("dlg.diskprops.frombytes")) : d.Add(New TextBox With {.ReadOnly = True, .Text = fromBytes, .Width = 420})
        End If
        If _disk.WriteProtected.HasValue Then
            _wpOriginal = _disk.WriteProtected.Value
            _chkWp.Checked = _wpOriginal
            _wpAvailable = True
        Else
            _chkWp.Enabled = False
            _chkWp.Text = UiText.T("dlg.diskprops.wpna", _disk.ImageFormat)
        End If
        _txtDetails = New TextArea With {.ReadOnly = True, .Wrap = False, .Font = Ui.Mono(9.0F), .Size = New Size(540, 130), .Text = String.Join(vbLf, _disk.DetailLines)}
        Dim grp2 As GroupBox = Ui.Group(UiText.T("dlg.diskprops.image"), Ui.Down(Nothing, Ui.Fields(d.ToArray()), _chkWp, Ui.Lbl(UiText.T("dlg.diskprops.details")), _txtDetails))

        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        Dim btnDmk As New Button With {.Text = UiText.T("btn.dmkanalysis"), .Enabled = Not _disk.IsHardDisk AndAlso Not _disk.IsCassette}
        AddHandler btnDmk.Click, Async Sub()
                                     Dim f As New DiskInfoForm(_disk.Path)
                                     f.ShowOwned(Me)
                                     Await f.LoadAsync()
                                 End Sub
        Finish(Ui.Down(Nothing, grp1, grp2, Ui.Across(_btnCheck, btnDmk)))
    End Sub

    ''' <summary>The DOS as the program's own reading of the DMK names it; a note when the file isn't a DMK.</summary>
    Private Function IdentifyFromBytes() As String
        Try
            Dim img As DmkImage = DmkImage.Load(_disk.Path)
            If img.Kind <> "DMK" Then Return UiText.T("dlg.diskprops.notdmk", img.Kind)
            Return DosIdentifier.Identify(img).Result
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>One line on what "disk check" found: how many signs, and the schemes it named.</summary>
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

    Private Async Function RunCheckAsync() As Task
        _btnCheck.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.DiskCheckAsync(_disk.Path, _disk.PartitionArg)
            _txtDetails.Text = Ui.Lf(If(r.Succeeded, r.StdOut, UiText.T("dlg.diskprops.checkfailed", r.ErrorText)))
        Catch ex As Exception
            _txtDetails.Text = ex.Message
        Finally
            _btnCheck.Enabled = True
        End Try
    End Function
End Class
