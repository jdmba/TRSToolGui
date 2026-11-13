' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  SectorMapForm.vb — "Sector map" (Disk menu): trs80-tool's "sectors" command
'  as a window. Everything comes from one "sectors --json --contents IMAGE"
'  call (schema: trs80-tool's sectors.ts): a grid with one row per track
'  (cylinder/side) and one column per sector number, lettered and coloured
'  like the tool's own map (S single, D double, C CRC error, X deleted mark,
'  ? missing, - not on this track); the selected sector's header line and hex
'  dump; and the list of bad sectors with the files that use them (what
'  "sectors --bad" prints). The tool's own text reports are one click away.
'  Floppies only, as in trs80-tool.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class SectorMapForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private _info As Trs80Tool.SectorsInfileJson = Nothing

    Private ReadOnly _grid As New DataGridView()
    Private ReadOnly _details As New TextBox()
    Private ReadOnly _bad As New ListView()
    Private ReadOnly _lblSummary As New Label()
    Private ReadOnly _lblLegend As New Label()
    Private ReadOnly _split As New SplitContainer()
    Private ReadOnly _splitRight As New SplitContainer()
    Private ReadOnly _btnText As New Button()
    Private ReadOnly _btnBadText As New Button()
    Private ReadOnly _btnSave As New Button()
    Private ReadOnly _btnRefresh As New Button()
    Private ReadOnly _btnProtection As New Button()
    Private ReadOnly _btnClose As New Button()
    Private ReadOnly _panelButtons As New Panel()

    Private Shared ReadOnly MonoFont As New Font("Consolas", 9.0F)

    Public Sub New(disk As DiskImage)
        _disk = disk
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Text = "Sector map - " & Path.GetFileName(_disk.Path)
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(1000, 640)
        MinimumSize = New Size(640, 400)
        ShowInTaskbar = True
        KeyPreview = True

        _lblSummary.Dock = DockStyle.Top
        _lblSummary.Height = 22
        _lblSummary.TextAlign = ContentAlignment.MiddleLeft
        _lblSummary.Padding = New Padding(6, 0, 0, 0)
        _lblSummary.Text = "Reading sectors..."

        _lblLegend.Dock = DockStyle.Top
        _lblLegend.Height = 20
        _lblLegend.TextAlign = ContentAlignment.MiddleLeft
        _lblLegend.Padding = New Padding(6, 0, 0, 0)
        _lblLegend.ForeColor = SystemColors.GrayText
        _lblLegend.Text = "Legend (as trs80-tool prints it):  S single density   D double density   C CRC error   X deleted-data mark   ? not found   - not on this track.  Click a sector for its header line and hex dump."

        With _grid
            .Dock = DockStyle.Fill
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToOrderColumns = False
            .RowHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .MultiSelect = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 22
            .RowTemplate.Height = 20
            .Font = MonoFont
            .BackgroundColor = SystemColors.Window
            .BorderStyle = BorderStyle.None
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .GridColor = SystemColors.ControlLight
            .ShowCellToolTips = True
            .EnableHeadersVisualStyles = True
        End With
        AddHandler _grid.SelectionChanged, Sub() ShowSelectedSector()

        With _details
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = MonoFont
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With

        With _bad
            .View = View.Details
            .FullRowSelect = True
            .MultiSelect = False
            .HideSelection = False
            .Dock = DockStyle.Fill
            .Columns.Add("Cyl", 40, HorizontalAlignment.Right)
            .Columns.Add("Side", 40, HorizontalAlignment.Right)
            .Columns.Add("Sector", 50, HorizontalAlignment.Right)
            .Columns.Add("Problem", 220)
            .Columns.Add("File(s) at this sector", 200)
        End With
        AddHandler _bad.SelectedIndexChanged, Sub() SelectFromBadList()

        Dim badHeader As New Label() With {.Dock = DockStyle.Top, .Height = 20, .Text = "Bad sectors (missing or CRC error) and the files that use them — what ""sectors --bad"" lists:", .TextAlign = ContentAlignment.MiddleLeft}
        Dim detHeader As New Label() With {.Dock = DockStyle.Top, .Height = 20, .Text = "Selected sector:", .TextAlign = ContentAlignment.MiddleLeft}

        _splitRight.Orientation = Orientation.Horizontal
        _splitRight.Dock = DockStyle.Fill
        _splitRight.Panel1.Controls.Add(_details)
        _splitRight.Panel1.Controls.Add(detHeader)
        _splitRight.Panel2.Controls.Add(_bad)
        _splitRight.Panel2.Controls.Add(badHeader)

        _split.Orientation = Orientation.Vertical
        _split.Dock = DockStyle.Fill
        _split.Panel1.Controls.Add(_grid)
        _split.Panel2.Controls.Add(_splitRight)

        _panelButtons.Dock = DockStyle.Bottom
        _panelButtons.Height = 40
        _btnText.Text = "Text report..." : _btnText.Bounds = New Rectangle(8, 7, 110, 26)
        _btnBadText.Text = "Bad sectors (text)..." : _btnBadText.Bounds = New Rectangle(124, 7, 140, 26)
        _btnSave.Text = "Save report..." : _btnSave.Bounds = New Rectangle(270, 7, 110, 26)
        _btnRefresh.Text = "Refresh" : _btnRefresh.Bounds = New Rectangle(386, 7, 80, 26)
        _btnProtection.Text = "Protection && artifacts..." : _btnProtection.Bounds = New Rectangle(472, 7, 190, 26) : _btnProtection.Enabled = False
        AddHandler _btnProtection.Click, Sub() ShowProtection()
        _btnClose.Text = "Close" : _btnClose.Bounds = New Rectangle(900, 7, 80, 26) : _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler _btnText.Click, Async Sub() Await ShowTextReportAsync(False)
        AddHandler _btnBadText.Click, Async Sub() Await ShowTextReportAsync(True)
        AddHandler _btnSave.Click, Async Sub() Await SaveReportAsync()
        AddHandler _btnRefresh.Click, Async Sub() Await LoadAsync()
        AddHandler _btnClose.Click, Sub() Close()
        _panelButtons.Controls.AddRange(New Control() {_btnText, _btnBadText, _btnSave, _btnRefresh, _btnProtection, _btnClose})
        _btnClose.Left = _panelButtons.Width - 88

        Controls.Add(_split)
        Controls.Add(_lblLegend)
        Controls.Add(_lblSummary)
        Controls.Add(_panelButtons)
        CancelButton = _btnClose
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Try
            _split.SplitterDistance = Math.Max(200, ClientSize.Width * 55 \ 100)
            _splitRight.SplitterDistance = Math.Max(100, _splitRight.Height * 55 \ 100)
        Catch
        End Try
        Await LoadAsync()
    End Sub

    ' ------------------------------------------------------------------ loading

    Private Async Function LoadAsync() As Task
        UseWaitCursor = True
        _btnRefresh.Enabled = False
        _lblSummary.Text = "Reading sectors with trs80-tool..."
        Try
            Dim caps As EngineCapabilities = Program.Tool.Capabilities
            Dim wantProtection As Boolean = caps IsNot Nothing AndAlso caps.HasTopLevelOption("sectors", "--protection")
            _info = Await Program.Tool.SectorsInfileAsync(_disk.Path, True, False, wantProtection)
        Catch ex As Exception
            UseWaitCursor = False
            _btnRefresh.Enabled = True
            _lblSummary.Text = "trs80-tool could not read the sectors."
            MessageBox.Show(Me, "trs80-tool could not read the sectors of this image:" & vbCrLf & vbCrLf & ex.Message, "Sector map", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        FillGrid()
        FillBadList()
        Dim signs As Integer = If(_info.Protection Is Nothing, 0, _info.Protection.Count)
        Dim artifacts As Integer = If(_info.ImagingArtifacts Is Nothing, 0, _info.ImagingArtifacts.Count)
        _btnProtection.Text = "Protection && artifacts (" & (signs + artifacts).ToString(CultureInfo.InvariantCulture) & ")..."
        _btnProtection.Enabled = _info.Protection IsNot Nothing OrElse artifacts > 0
        UseWaitCursor = False
        _btnRefresh.Enabled = True
    End Function

    ''' <summary>
    ''' What trs80-tool says about copy protection and imaging artifacts on this floppy ("sectors --protection"
    ''' and the artifacts of "sectors"), as text.
    ''' </summary>
    Private Function ProtectionReport() As String
        Dim sb As New StringBuilder()
        sb.AppendLine(Path.GetFileName(_disk.Path) & ": " & If(_info.Description, "") & If(String.IsNullOrEmpty(_info.Os), "", ", " & _info.Os))
        sb.AppendLine()
        sb.AppendLine("Signs of copy protection (trs80-tool sectors --protection)")
        If _info.Protection Is Nothing Then
            sb.AppendLine("    (not reported by this trs80-tool)")
        ElseIf _info.Protection.Count = 0 Then
            sb.AppendLine("    No signs of copy protection (which doesn't prove there's none).")
        Else
            For Each sign In _info.Protection
                Dim verdict As String
                Select Case If(sign.Verdict, "")
                    Case "protected" : verdict = "PROTECTED"
                    Case "dosProtected" : verdict = "DOS PROTECTED"
                    Case "possible" : verdict = "possible"
                    Case "info" : verdict = "not protection"
                    Case Else : verdict = sign.Verdict
                End Select
                sb.AppendLine("    [" & verdict & "] " & sign.Message)
                If Not String.IsNullOrEmpty(sign.Scheme) Then sb.AppendLine("        Scheme: " & sign.Scheme)
                If sign.Schemes IsNot Nothing AndAlso sign.Schemes.Count > 0 Then sb.AppendLine("        Seen with (trs-80.com scheme numbers): " & String.Join(", ", sign.Schemes))
            Next
        End If
        sb.AppendLine()
        sb.AppendLine("Imaging artifacts (bad or missing sectors the imaging likely caused; reimaging may read them)")
        If _info.ImagingArtifacts Is Nothing OrElse _info.ImagingArtifacts.Count = 0 Then
            sb.AppendLine("    None.")
        Else
            For Each a In _info.ImagingArtifacts
                sb.AppendLine("    " & a.Message)
            Next
        End If
        Return sb.ToString()
    End Function

    Private Sub ShowProtection()
        If _info Is Nothing Then Return
        Dim f As New TextReportForm("Copy protection signs and imaging artifacts - " & Path.GetFileName(_disk.Path), ProtectionReport())
        f.Show(Me)
    End Sub

    Private Sub FillGrid()
        _grid.SuspendLayout()
        Try
            _grid.Columns.Clear()
            _grid.Rows.Clear()
            Dim sectors As List(Of Trs80Tool.SectorJson) = _info.Sectors
            If sectors.Count = 0 Then
                _lblSummary.Text = "No sectors reported for this image."
                Return
            End If
            ' Columns: every sector number that occurs anywhere on the disk, ascending (stray numbers get their own column).
            Dim numbers As New SortedSet(Of Integer)
            Dim maxCyl As Integer = 0, maxSide As Integer = 0
            Dim bad As Integer = 0, stray As Integer = 0
            For Each sj In sectors
                numbers.Add(sj.Sector)
                If sj.Cylinder > maxCyl Then maxCyl = sj.Cylinder
                If sj.Side > maxSide Then maxSide = sj.Side
                If sj.IsBad Then bad += 1
                If sj.OutsideUsualNumbers Then stray += 1
            Next
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "cyl", .HeaderText = "Cyl", .Width = 40, .Frozen = True, .SortMode = DataGridViewColumnSortMode.NotSortable, .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight, .BackColor = SystemColors.Control}})
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "side", .HeaderText = "Side", .Width = 40, .Frozen = True, .SortMode = DataGridViewColumnSortMode.NotSortable, .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter, .BackColor = SystemColors.Control}})
            Dim colIndex As New Dictionary(Of Integer, Integer)
            For Each n In numbers
                Dim col As New DataGridViewTextBoxColumn() With {.Name = "s" & n.ToString(CultureInfo.InvariantCulture), .HeaderText = n.ToString(CultureInfo.InvariantCulture), .Width = If(n >= 100, 34, 26), .SortMode = DataGridViewColumnSortMode.NotSortable}
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                colIndex(n) = _grid.Columns.Add(col)
            Next
            ' Rows: cylinder-major, side-minor (the order the tool lists sectors in).
            Dim lookup As New Dictionary(Of String, Trs80Tool.SectorJson)
            For Each sj In sectors
                lookup(sj.Cylinder.ToString(CultureInfo.InvariantCulture) & "/" & sj.Side.ToString(CultureInfo.InvariantCulture) & "/" & sj.Sector.ToString(CultureInfo.InvariantCulture)) = sj
            Next
            For cyl As Integer = 0 To maxCyl
                For side As Integer = 0 To maxSide
                    Dim rowIdx As Integer = _grid.Rows.Add()
                    Dim row As DataGridViewRow = _grid.Rows(rowIdx)
                    row.Cells(0).Value = cyl.ToString(CultureInfo.InvariantCulture)
                    row.Cells(1).Value = side.ToString(CultureInfo.InvariantCulture)
                    For Each n In numbers
                        Dim cell As DataGridViewCell = row.Cells(colIndex(n))
                        Dim sj As Trs80Tool.SectorJson = Nothing
                        If lookup.TryGetValue(cyl.ToString(CultureInfo.InvariantCulture) & "/" & side.ToString(CultureInfo.InvariantCulture) & "/" & n.ToString(CultureInfo.InvariantCulture), sj) Then
                            cell.Value = sj.MapLetter
                            cell.Tag = sj
                            StyleCell(cell, sj)
                            cell.ToolTipText = sj.HeaderText & If(sj.Files IsNot Nothing AndAlso sj.Files.Count > 0, vbCrLf & "File(s): " & String.Join(", ", sj.Files), "")
                        Else
                            cell.Value = "-"
                            cell.Style.ForeColor = SystemColors.GrayText
                            cell.ToolTipText = "Not on this track"
                        End If
                    Next
                Next
            Next
            Dim desc As String = If(_info.Description, "")
            Dim os As String = If(_info.Os, "unknown operating system")
            _lblSummary.Text = desc & If(desc.Length > 0, ", ", "") & os & "  —  " & (maxCyl + 1).ToString(CultureInfo.InvariantCulture) & " cylinders, " & (maxSide + 1).ToString(CultureInfo.InvariantCulture) & " side(s), " &
                sectors.Count.ToString(CultureInfo.InvariantCulture) & " sector positions, " & If(bad = 0, "no bad sectors", bad.ToString(CultureInfo.InvariantCulture) & " bad sector(s)") &
                If(stray > 0, ", " & stray.ToString(CultureInfo.InvariantCulture) & " outside the usual numbers", "")
            _grid.ClearSelection()
        Finally
            _grid.ResumeLayout()
        End Try
    End Sub

    Private Shared Sub StyleCell(cell As DataGridViewCell, sj As Trs80Tool.SectorJson)
        Select Case sj.MapLetter
            Case "?"
                cell.Style.BackColor = Color.Firebrick : cell.Style.ForeColor = Color.White
            Case "C"
                cell.Style.BackColor = Color.Red : cell.Style.ForeColor = Color.White
            Case "X"
                cell.Style.BackColor = Color.Orange : cell.Style.ForeColor = Color.Black
            Case "S"
                cell.Style.BackColor = Color.White : cell.Style.ForeColor = Color.Black
            Case Else
                cell.Style.BackColor = Color.FromArgb(230, 238, 250) : cell.Style.ForeColor = Color.Black
        End Select
        If sj.OutsideUsualNumbers AndAlso Not sj.IsBad Then
            cell.Style.BackColor = Color.LemonChiffon
        End If
        cell.Style.SelectionBackColor = Color.RoyalBlue
        cell.Style.SelectionForeColor = Color.White
    End Sub

    Private Sub FillBadList()
        _bad.BeginUpdate()
        Try
            _bad.Items.Clear()
            For Each sj In _info.Sectors
                If Not sj.IsBad Then Continue For
                Dim problem As String
                If sj.Missing Then
                    problem = "missing (not found on the track)"
                Else
                    problem = "CRC error"
                    Dim parts As New List(Of String)
                    If sj.IdCrc IsNot Nothing AndAlso Not sj.IdCrc.Valid Then parts.Add("ID " & sj.IdCrc.Stored.ToString("X4", CultureInfo.InvariantCulture) & " != " & sj.IdCrc.Computed.ToString("X4", CultureInfo.InvariantCulture))
                    If sj.DataCrc IsNot Nothing AndAlso Not sj.DataCrc.Valid Then parts.Add("data " & sj.DataCrc.Stored.ToString("X4", CultureInfo.InvariantCulture) & " != " & sj.DataCrc.Computed.ToString("X4", CultureInfo.InvariantCulture))
                    If parts.Count > 0 Then problem &= " (" & String.Join(", ", parts) & ")"
                End If
                If sj.ImagingArtifactText.Length > 0 Then problem &= " - imaging artifact: " & sj.ImagingArtifactText
                Dim files As String = If(sj.Files IsNot Nothing AndAlso sj.Files.Count > 0, String.Join(", ", sj.Files), "(no file at this sector)")
                Dim it As New ListViewItem(sj.Cylinder.ToString(CultureInfo.InvariantCulture)) With {.Tag = sj}
                it.SubItems.Add(sj.Side.ToString(CultureInfo.InvariantCulture))
                it.SubItems.Add(sj.Sector.ToString(CultureInfo.InvariantCulture))
                it.SubItems.Add(problem)
                it.SubItems.Add(files)
                _bad.Items.Add(it)
            Next
            If _bad.Items.Count = 0 Then
                _bad.Items.Add(New ListViewItem(New String() {"", "", "", "No bad sectors.", ""}) With {.ForeColor = SystemColors.GrayText})
            End If
        Finally
            _bad.EndUpdate()
        End Try
    End Sub

    ' ------------------------------------------------------------------ selection

    Private Sub ShowSelectedSector()
        If _grid.SelectedCells.Count = 0 Then Return
        Dim cell As DataGridViewCell = _grid.SelectedCells(0)
        Dim sj As Trs80Tool.SectorJson = TryCast(cell.Tag, Trs80Tool.SectorJson)
        If sj Is Nothing Then
            If cell.ColumnIndex >= 2 Then _details.Text = "Cylinder " & _grid.Rows(cell.RowIndex).Cells(0).Value.ToString() & ", side " & _grid.Rows(cell.RowIndex).Cells(1).Value.ToString() & ", sector " & _grid.Columns(cell.ColumnIndex).HeaderText & ": not on this track"
            Return
        End If
        Dim sb As New StringBuilder()
        sb.AppendLine(sj.HeaderText)
        If sj.ImagingArtifactText.Length > 0 Then sb.AppendLine("Imaging artifact: " & sj.ImagingArtifactText & ".")
        If sj.Files IsNot Nothing AndAlso sj.Files.Count > 0 Then
            sb.AppendLine(If(sj.Files.Count = 1, "File", "Files") & " at this sector: " & String.Join(", ", sj.Files))
        Else
            sb.AppendLine("There is no file at this sector.")
        End If
        Dim data As Byte() = sj.GetBytes()
        If data IsNot Nothing Then
            sb.AppendLine()
            sb.Append(TextUtil.HexDump(data))
        End If
        _details.Text = sb.ToString()
        _details.Select(0, 0)
    End Sub

    Private Sub SelectFromBadList()
        If _bad.SelectedItems.Count = 0 Then Return
        Dim sj As Trs80Tool.SectorJson = TryCast(_bad.SelectedItems(0).Tag, Trs80Tool.SectorJson)
        If sj Is Nothing Then Return
        For Each row As DataGridViewRow In _grid.Rows
            If row.Cells(0).Value.ToString() <> sj.Cylinder.ToString(CultureInfo.InvariantCulture) OrElse row.Cells(1).Value.ToString() <> sj.Side.ToString(CultureInfo.InvariantCulture) Then Continue For
            For Each cell As DataGridViewCell In row.Cells
                If cell.Tag Is sj Then
                    _grid.ClearSelection()
                    cell.Selected = True
                    _grid.CurrentCell = cell
                    Return
                End If
            Next
        Next
    End Sub

    ' ------------------------------------------------------------------ the tool's own text

    Private Async Function ShowTextReportAsync(badOnly As Boolean) As Task
        UseWaitCursor = True
        Dim r As ToolResult
        Try
            r = Await Program.Tool.SectorsTextAsync(_disk.Path, badOnly)
        Catch ex As Exception
            UseWaitCursor = False
            MessageBox.Show(Me, ex.Message, "Sector map", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        UseWaitCursor = False
        Dim body As String = If(r.Succeeded, r.StdOut, r.ErrorText)
        If badOnly AndAlso body.Trim().Length = 0 Then body = "trs80-tool reports no bad sectors on this floppy (""sectors --bad"" prints nothing for a clean disk)."
        Dim f As New TextReportForm(If(badOnly, "sectors --bad", "sectors") & " - " & Path.GetFileName(_disk.Path), body)
        f.Show(Me)
    End Function

    Private Async Function SaveReportAsync() As Task
        Using sfd As New SaveFileDialog()
            sfd.Title = "Save sector report"
            sfd.Filter = "Text file (*.txt)|*.txt|All files (*.*)|*.*"
            sfd.FileName = Path.GetFileNameWithoutExtension(_disk.Path) & "_sectors.txt"
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            UseWaitCursor = True
            Try
                Dim all As ToolResult = Await Program.Tool.SectorsTextAsync(_disk.Path, False)
                Dim bad As ToolResult = Await Program.Tool.SectorsTextAsync(_disk.Path, True)
                Dim sb As New StringBuilder()
                sb.AppendLine("trs80-tool sectors """ & _disk.Path & """").AppendLine()
                sb.AppendLine(If(all.Succeeded, all.StdOut, all.ErrorText))
                sb.AppendLine().AppendLine("trs80-tool sectors --bad """ & _disk.Path & """").AppendLine()
                Dim badText As String = If(bad.Succeeded, bad.StdOut, bad.ErrorText)
                sb.AppendLine(If(badText.Trim().Length = 0, "(no bad sectors)", badText))
                If _info IsNot Nothing Then sb.AppendLine().AppendLine(ProtectionReport())
                File.WriteAllText(sfd.FileName, sb.ToString().Replace(vbCr, "").Replace(vbLf, vbCrLf), Encoding.UTF8)
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, "Save sector report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                UseWaitCursor = False
            End Try
        End Using
    End Function

End Class
