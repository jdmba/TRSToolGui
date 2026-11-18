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
    Private ReadOnly _chkNoAuto As New CheckBox With {.Text = "Remove"}
    Private ReadOnly _chkWp As New CheckBox With {.Text = "Write protected (the image's own flag)"}
    Private ReadOnly _chkSetMaster As New CheckBox With {.Text = "Set master password:"}
    Private ReadOnly _txtMaster As New TextBox With {.Width = 120, .MaxLength = 8}
    Private _txtDetails As TextArea
    Private ReadOnly _btnCheck As New Button With {.Text = "Check..."}
    Private _wpOriginal As Boolean
    Private _labelEditable As Boolean
    Private _wpAvailable As Boolean

    Public Sub New(disk As DiskImage)
        MyBase.New("Virtual disk properties")
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
            v.Add("Volume name:") : v.Add(Ui.Across(_txtName))
            _txtDate.Text = _disk.VolumeDate
            _txtDate.Enabled = editable
            v.Add("Volume date:") : v.Add(Ui.Across(_txtDate, Ui.Lbl("(new date as YYYY-MM-DD)", True)))
            _txtAuto.Text = _disk.AutoCommand
            _txtAuto.Enabled = editable
            _chkNoAuto.Enabled = editable AndAlso _disk.AutoCommand.Length > 0
            v.Add("AUTO command:") : v.Add(Ui.Across(_txtAuto, _chkNoAuto))
            v.Add("Master password:") : v.Add(Ui.Lbl(If(_disk.MasterPassword.Length > 0, _disk.MasterPassword, "(unknown)")))
            _chkSetMaster.Enabled = editable
            _txtMaster.Enabled = False
            AddHandler _chkSetMaster.CheckedChanged, Sub() _txtMaster.Enabled = Ui.IsOn(_chkSetMaster)
            v.Add(_chkSetMaster) : v.Add(Ui.Across(_txtMaster, Ui.Lbl("(blank = none; trs80-tool: label --password)", True)))
        End If
        If _disk.ApproxCapacityBytes > 0 Then v.Add("Capacity (approx.):") : v.Add(Ui.Lbl(TextUtil.FormatKb(_disk.ApproxCapacityBytes) & "  (" & TextUtil.FormatBytes(_disk.ApproxCapacityBytes) & " bytes, from the geometry)"))
        v.Add("Free space:") : v.Add(Ui.Lbl(If(_disk.FreeText.Length > 0, _disk.FreeText, "(not known for this DOS)")))
        v.Add("File system:") : v.Add(Ui.Lbl(If(_disk.DosName.Length > 0, _disk.DosName, "(none recognized)")))
        If _disk.Geometry IsNot Nothing Then
            v.Add("Cylinders:") : v.Add(Ui.Lbl(_disk.Geometry.CylinderCount.ToString(CultureInfo.InvariantCulture)))
            v.Add("Sides:") : v.Add(Ui.Lbl(_disk.Geometry.SideCount.ToString(CultureInfo.InvariantCulture) & If(_disk.AllocationSides > 0 AndAlso _disk.AllocationSides < _disk.Geometry.SideCount, "  (" & _disk.AllocationSides.ToString(CultureInfo.InvariantCulture) & " used by the DOS)", "")))
            v.Add("Density:") : v.Add(Ui.Lbl(_disk.DensityText))
        ElseIf _disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing Then
            v.Add("Cylinders:") : v.Add(Ui.Lbl(_disk.ActivePartition.CylinderCount.ToString(CultureInfo.InvariantCulture)))
            v.Add("Heads:") : v.Add(Ui.Lbl(_disk.ActivePartition.HeadCount.ToString(CultureInfo.InvariantCulture)))
            v.Add("Location:") : v.Add(Ui.Lbl(_disk.ActivePartition.Location))
        End If
        v.Add("Damaged:") : v.Add(Ui.Lbl(If(_disk.Damaged, "Yes (Check… lists the problems)", "No")))
        ' What "disk check" found beyond damage: signs of copy protection and imaging artifacts.
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette AndAlso _disk.CheckText.Length > 0 Then
            v.Add("Copy protection:") : v.Add(Ui.Lbl(ProtectionSummary()))
            If _disk.ImagingArtifacts.Count > 0 Then
                v.Add("Imaging artifacts:") : v.Add(Ui.Lbl(_disk.ImagingArtifacts.Count.ToString(CultureInfo.InvariantCulture) & " kind(s) that look like damage; reimaging may read those sectors"))
            End If
        End If
        v.Add("Writable:") : v.Add(Ui.Wrapped(If(_disk.IsWritable, "Yes", "No — " & If(_disk.ReadOnlyReason, "")), 420))
        Dim grp1 As GroupBox = Ui.Group(If(_disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing, "Partition " & _disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture) & " information (" & _disk.Partitions.Count.ToString(CultureInfo.InvariantCulture) & " partitions total)", "Volume information"), Ui.Fields(v.ToArray()))

        ' ---- image
        Dim d As New List(Of Object)
        d.Add("Filename:") : d.Add(New TextBox With {.ReadOnly = True, .Text = _disk.Path, .Width = 420})
        d.Add("Format:") : d.Add(Ui.Lbl(_disk.ImageFormat & If(_disk.Description.Length > 0, "  (" & _disk.Description & ")", "")))
        If _disk.PdriveText.Length > 0 Then
            ' NEWDOS/80: the PDRIVE line as trs80-tool gives it, to copy, and where its values come from.
            d.Add("PDRIVE:") : d.Add(New TextBox With {.ReadOnly = True, .Text = _disk.PdriveText, .Width = 420})
            d.Add("") : d.Add(Ui.Wrapped("(" & _disk.PdriveSource & ")", 420))
        End If
        If Not _disk.IsHardDisk AndAlso Not _disk.IsCassette Then
            ' The program's own identification from the boot sector, directory and GAT bytes.
            Dim fromBytes As String = IdentifyFromBytes()
            If fromBytes.Length > 0 Then d.Add("From the bytes:") : d.Add(New TextBox With {.ReadOnly = True, .Text = fromBytes, .Width = 420})
        End If
        If _disk.WriteProtected.HasValue Then
            _wpOriginal = _disk.WriteProtected.Value
            _chkWp.Checked = _wpOriginal
            _wpAvailable = True
        Else
            _chkWp.Enabled = False
            _chkWp.Text &= "  — not available for " & _disk.ImageFormat & " images"
        End If
        _txtDetails = New TextArea With {.ReadOnly = True, .Wrap = False, .Font = Ui.Mono(9.0F), .Size = New Size(540, 130), .Text = String.Join(vbLf, _disk.DetailLines)}
        Dim grp2 As GroupBox = Ui.Group("Disk information", Ui.Down(Nothing, Ui.Fields(d.ToArray()), _chkWp, Ui.Lbl("Details from trs80-tool:"), _txtDetails))

        AddHandler _btnCheck.Click, Async Sub() Await RunCheckAsync()
        Dim btnDmk As New Button With {.Text = "DMK analysis...", .Enabled = Not _disk.IsHardDisk AndAlso Not _disk.IsCassette}
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
            If img.Kind <> "DMK" Then Return "(" & img.Kind & " image - the byte-level identification is for DMK images)"
            Return DosIdentifier.Identify(img).Result
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>One line on what "disk check" found: how many signs, and the schemes it named.</summary>
    Private Function ProtectionSummary() As String
        If _disk.ProtectionSigns.Count = 0 Then Return "No signs found (which doesn't prove there is none)"
        Dim names As New List(Of String)
        For Each sign In _disk.ProtectionSigns
            Dim name As String = If(String.IsNullOrEmpty(sign.Scheme), "", sign.Scheme)
            If name.Length = 0 AndAlso sign.Schemes IsNot Nothing AndAlso sign.Schemes.Count > 0 Then name = String.Join("/", sign.Schemes)
            If name.Length > 0 AndAlso Not names.Contains(name) Then names.Add(name)
        Next
        Dim n As Integer = _disk.ProtectionSigns.Count
        Dim text As String = n.ToString(CultureInfo.InvariantCulture) & If(n = 1, " sign", " signs")
        If _disk.ProtectedSignCount = 0 Then text &= ", none conclusive"
        If names.Count > 0 Then text &= ": " & String.Join("; ", names)
        Return text
    End Function

    Private Async Function RunCheckAsync() As Task
        _btnCheck.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.DiskCheckAsync(_disk.Path, _disk.PartitionArg)
            _txtDetails.Text = Ui.Lf(If(r.Succeeded, r.StdOut, "trs80-tool can't check this disk: " & r.ErrorText))
        Catch ex As Exception
            _txtDetails.Text = ex.Message
        Finally
            _btnCheck.Enabled = True
        End Try
    End Function
End Class
