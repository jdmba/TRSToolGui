' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  GranuleMapForm.vb — Disk → Granule map.
'
'  One "disk IMAGE --json map [PATTERN...]" run gives every granule of a
'  TRSDOS-family floppy as data: what it holds, which files claim it and
'  where its sectors are. This window draws that as a grid (one row per
'  lump, one column per granule), colours each file's granules, lists the
'  files beside it, and keeps the tool's own text drawing one click away.
'  Nothing here decides anything about the disk: every cell is the tool's.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class GranuleMapForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private _map As Trs80Tool.MapResultJson = Nothing
    Private _mapText As String = ""

    Private ReadOnly _txtFilter As New TextBox() With {.Width = 260}
    Private ReadOnly _btnApply As New Button() With {.Text = UiText.T("btn.show"), .Width = 70, .Height = 25}
    Private ReadOnly _lblSummary As New Label()
    Private ReadOnly _grid As New DataGridView()
    Private ReadOnly _files As New ListView()
    Private ReadOnly _details As New TextBox()
    Private ReadOnly _split As New SplitContainer()
    Private ReadOnly _splitRight As New SplitContainer()
    Private ReadOnly _btnText As New Button() With {.Text = UiText.T("btn.textmap"), .Width = 160, .Height = 26}
    Private ReadOnly _btnRefresh As New Button() With {.Text = UiText.T("btn.refresh"), .Width = 80, .Height = 26}
    Private ReadOnly _btnClose As New Button() With {.Text = UiText.T("btn.close"), .Width = 80, .Height = 26}

    Private Shared ReadOnly MonoFont As New Font(HostPlatform.MonospaceFontName, 9.0F)

    ''' <summary>Colours for the files, by the tool's colorIndex (it repeats past the end).</summary>
    Private Shared ReadOnly FileColors As Color() = {
        Color.FromArgb(174, 214, 241), Color.FromArgb(171, 235, 198), Color.FromArgb(249, 231, 159), Color.FromArgb(245, 183, 177),
        Color.FromArgb(215, 189, 226), Color.FromArgb(250, 215, 160), Color.FromArgb(162, 217, 206), Color.FromArgb(241, 196, 223),
        Color.FromArgb(200, 220, 150), Color.FromArgb(190, 200, 235), Color.FromArgb(235, 205, 185), Color.FromArgb(180, 225, 230)}

    Public Sub New(disk As DiskImage)
        _disk = disk
        Dim shown As String = UiText.T("dlg.granule.title", Path.GetFileName(disk.Path))
        If disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing Then shown = UiText.T("title.partition", shown, disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Text = shown
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(980, 640)
        MinimumSize = New Size(640, 400)
        ShowInTaskbar = True

        Dim top As New Panel() With {.Dock = DockStyle.Top, .Height = 36}
        Dim lblFilter As New Label() With {.Text = UiText.T("dlg.granule.filter"), .AutoSize = True, .Location = New Point(8, 10)}
        top.Controls.Add(lblFilter)
        _txtFilter.Location = New Point(380, 6)
        _btnApply.Location = New Point(648, 5)
        top.Controls.AddRange(New Control() {_txtFilter, _btnApply})
        ' Make room for the language in force (UiFit.vb); in English nothing here moves.
        UiFit.AfterAll(lblFilter, 8, _txtFilter, _btnApply)

        _lblSummary.Dock = DockStyle.Top
        _lblSummary.Height = 40
        _lblSummary.Padding = New Padding(6, 2, 0, 0)
        _lblSummary.Text = UiText.T("dlg.granule.reading")

        ' The summary and its legend are two lines of English and can be three of a longer
        ' language; the label is given the height its text really needs (UiFit.vb).

        With _grid
            .Dock = DockStyle.Fill
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
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
            .GridColor = SystemColors.ControlLight
            .ShowCellToolTips = True
        End With
        AddHandler _grid.SelectionChanged, Sub() ShowSelectedCell()

        With _files
            .View = View.Details
            .FullRowSelect = True
            .HideSelection = False
            .MultiSelect = False
            .Dock = DockStyle.Fill
            .Columns.Add("#", 36, HorizontalAlignment.Right)
            .Columns.Add(UiText.T("dlg.granule.col.file"), 110)
            .Columns.Add(UiText.T("dlg.granule.col.granules"), 70, HorizontalAlignment.Right)
        End With
        AddHandler _files.SelectedIndexChanged, Sub() SelectFirstCellOfFile()

        With _details
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Vertical
            .WordWrap = True
            .Font = MonoFont
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
        End With

        _splitRight.Dock = DockStyle.Fill
        _splitRight.Orientation = Orientation.Horizontal
        UiFit.FitListColumns(_files)
        _splitRight.Panel1.Controls.Add(_files)
        _splitRight.Panel2.Controls.Add(_details)

        _split.Dock = DockStyle.Fill
        _split.Panel1.Controls.Add(_grid)
        _split.Panel2.Controls.Add(_splitRight)

        Dim bottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 40}
        ' The two report buttons are sized to their own text: "trs80-tool's text map..." is
        ' "Carte texte de trs80-tool..." in French.
        UiFit.ButtonRowLeft(8, 8, _btnText, _btnRefresh)
        _btnText.Top = 7 : _btnRefresh.Top = 7
        UiFit.FitButton(_btnClose)
        _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        _btnClose.Location = New Point(bottom.Width - _btnClose.Width - 10, 7)
        bottom.Controls.AddRange(New Control() {_btnText, _btnRefresh, _btnClose})

        Controls.Add(_split)
        Controls.Add(_lblSummary)
        Controls.Add(top)
        Controls.Add(bottom)
        CancelButton = _btnClose
        AcceptButton = _btnApply

        AddHandler _btnApply.Click, Async Sub() Await LoadAsync()
        AddHandler _btnRefresh.Click, Async Sub() Await LoadAsync()
        AddHandler _btnClose.Click, Sub() Close()
        AddHandler _btnText.Click, Sub() ShowTextMap()
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Try
            _split.SplitterDistance = Math.Max(200, _split.Width - 300)
            _splitRight.SplitterDistance = Math.Max(100, _splitRight.Height - 170)
        Catch ex As Exception
            ' A very small window: leave the splitters where they are.
        End Try
        Await LoadAsync()
    End Sub

    Private Function Patterns() As List(Of String)
        Dim pats As New List(Of String)
        If Not String.IsNullOrWhiteSpace(_txtFilter.Text) Then pats.AddRange(_txtFilter.Text.Split(New Char() {" "c, ","c}, StringSplitOptions.RemoveEmptyEntries))
        Return pats
    End Function

    Private Async Function LoadAsync() As Task
        UseWaitCursor = True
        _btnApply.Enabled = False
        _btnRefresh.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.DiskMapAsync(_disk.Path, _disk.PartitionArg, Patterns())
            _map = If(r.Succeeded AndAlso r.Disk IsNot Nothing, r.Disk.ResultOf(Of Trs80Tool.MapResultJson)("map"), Nothing)
            _mapText = If(r.Succeeded, r.StdOut, "")
            If _map Is Nothing OrElse _map.Cells Is Nothing Then
                _grid.Columns.Clear()
                _files.Items.Clear()
                _details.Text = ""
                SetSummary(UiText.T("dlg.granule.failed", r.ErrorText))
                Return
            End If
            FillGrid()
            FillFiles()
            SetSummary(SummaryText())
        Catch ex As Exception
            SetSummary(ex.Message)
        Finally
            UseWaitCursor = False
            _btnApply.Enabled = True
            _btnRefresh.Enabled = True
        End Try
    End Function

    ''' <summary>Sets the summary line and gives the label the height its wrapped text needs.</summary>
    Private Sub SetSummary(text As String)
        _lblSummary.Text = text
        UiFit.FitWrapped(_lblSummary, ClientSize.Width)
    End Sub

    Private Function SummaryText() As String
        Dim t As Trs80Tool.MapTotalsJson = _map.Totals
        If t Is Nothing Then Return ""
        Dim used As Integer = t.GranuleCount - t.FreeGranuleCount - t.NotOnImageGranuleCount - t.PastEndGranuleCount
        Dim sb As New StringBuilder()
        ' The already-formatted numbers are passed, so nothing about number formatting changes here.
        sb.Append(UiText.T("dlg.granule.summary", used.ToString("N0", CultureInfo.InvariantCulture), t.GranuleCount.ToString("N0", CultureInfo.InvariantCulture), t.FreeGranuleCount.ToString("N0", CultureInfo.InvariantCulture), TextUtil.FormatKb(t.FreeByteCount), t.GranuleByteCount.ToString("N0", CultureInfo.InvariantCulture)))
        ' The space before each of the next three sentences is layout, so it stays here.
        If t.NotOnImageGranuleCount > 0 Then sb.Append(" ").Append(UiText.T("dlg.granule.notonimage", t.NotOnImageGranuleCount.ToString("N0", CultureInfo.InvariantCulture)))
        If t.PastEndGranuleCount > 0 Then sb.Append(" ").Append(UiText.T("dlg.granule.pastend", t.PastEndGranuleCount.ToString("N0", CultureInfo.InvariantCulture)))
        If t.LockedOutGranuleCount > 0 Then sb.Append(" ").Append(UiText.T("dlg.granule.lockedout", t.LockedOutGranuleCount.ToString("N0", CultureInfo.InvariantCulture)))
        sb.AppendLine()
        sb.Append(UiText.T("dlg.granule.legend"))
        Return sb.ToString()
    End Function

    ''' <summary>The letter (or file number) and colour of a cell, by the tool's kind.</summary>
    Private Sub StyleCell(cell As DataGridViewCell, c As Trs80Tool.MapCellJson)
        Dim text As String
        Dim back As Color = SystemColors.Window
        Dim fore As Color = SystemColors.WindowText
        Select Case If(c.Kind, "")
            Case "free" : text = "·" : fore = SystemColors.GrayText
            Case "file"
                Dim index As Integer = If(c.File, -1)
                text = If(index >= 0, (index + 1).ToString(CultureInfo.InvariantCulture), "f")
                Dim colorIndex As Integer = index
                If index >= 0 AndAlso _map.Files IsNot Nothing AndAlso index < _map.Files.Count AndAlso _map.Files(index).ColorIndex.HasValue Then colorIndex = _map.Files(index).ColorIndex.Value
                If colorIndex >= 0 Then back = FileColors(colorIndex Mod FileColors.Length)
            Case "otherFile" : text = "-" : back = Color.FromArgb(232, 232, 232)
            Case "directory" : text = "D" : back = Color.FromArgb(210, 210, 210)
            Case "boot" : text = "B" : back = Color.FromArgb(210, 210, 210)
            Case "lockedOut" : text = "x" : fore = SystemColors.GrayText
            Case "shared" : text = "!" : back = Color.FromArgb(240, 128, 128)
            Case "claimedButFree" : text = "?" : back = Color.FromArgb(255, 190, 120)
            Case "unclaimed" : text = "u" : back = Color.FromArgb(255, 235, 150)
            Case "notOnImage" : text = "n" : fore = Color.Firebrick
            Case "pastEnd" : text = "p" : fore = SystemColors.GrayText
            Case Else : text = "*"       ' a kind a later trs80-tool added
        End Select
        cell.Value = text
        cell.Style.BackColor = back
        cell.Style.ForeColor = fore
        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        cell.Tag = c
        cell.ToolTipText = CellText(c)
    End Sub

    Private Sub FillGrid()
        _grid.SuspendLayout()
        _grid.Columns.Clear()
        Dim maxGranule As Integer = 0
        Dim lumps As New SortedDictionary(Of Integer, List(Of Trs80Tool.MapCellJson))
        For Each c In _map.Cells
            If c.Granule > maxGranule Then maxGranule = c.Granule
            Dim list As List(Of Trs80Tool.MapCellJson) = Nothing
            If Not lumps.TryGetValue(c.Lump, list) Then
                list = New List(Of Trs80Tool.MapCellJson)
                lumps(c.Lump) = list
            End If
            list.Add(c)
        Next
        ' trs80-tool 5.2.0 says what a row is on this disk: a cylinder where lumps are cylinders, otherwise a lump.
        Dim rowsAreCylinders As Boolean = _map.Layout IsNot Nothing AndAlso String.Equals(_map.Layout.LumpTerm, "cylinder", StringComparison.OrdinalIgnoreCase)
        Dim colLump As New DataGridViewTextBoxColumn() With {.HeaderText = If(rowsAreCylinders, UiText.T("dlg.granule.col.cylinder"), UiText.T("dlg.granule.col.lump")), .Width = If(rowsAreCylinders, 72, 52), .SortMode = DataGridViewColumnSortMode.NotSortable}
        colLump.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        colLump.DefaultCellStyle.BackColor = SystemColors.Control
        _grid.Columns.Add(colLump)
        For g As Integer = 0 To maxGranule
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {.HeaderText = g.ToString(CultureInfo.InvariantCulture), .Width = 34, .SortMode = DataGridViewColumnSortMode.NotSortable})
        Next
        For Each kv In lumps
            Dim rowIndex As Integer = _grid.Rows.Add()
            Dim row As DataGridViewRow = _grid.Rows(rowIndex)
            row.Cells(0).Value = kv.Key.ToString(CultureInfo.InvariantCulture)
            For Each c In kv.Value
                StyleCell(row.Cells(1 + c.Granule), c)
            Next
        Next
        _grid.ResumeLayout()
        If _grid.Rows.Count > 0 Then _grid.CurrentCell = _grid.Rows(0).Cells(Math.Min(1, _grid.Columns.Count - 1))
        ShowSelectedCell()
    End Sub

    Private Sub FillFiles()
        _files.BeginUpdate()
        _files.Items.Clear()
        If _map.Files IsNot Nothing Then
            For i As Integer = 0 To _map.Files.Count - 1
                Dim f As Trs80Tool.MapFileJson = _map.Files(i)
                Dim item As New ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture)) With {.UseItemStyleForSubItems = False, .Tag = i}
                If f.ColorIndex.HasValue Then item.BackColor = FileColors(f.ColorIndex.Value Mod FileColors.Length)
                item.SubItems.Add(If(f.Name, "").Replace("."c, "/"c))
                item.SubItems.Add(f.GranuleCount.ToString("N0", CultureInfo.InvariantCulture))
                _files.Items.Add(item)
            Next
        End If
        _files.EndUpdate()
    End Sub

    ''' <summary>What a granule is, who claims it and where its sectors are, in words.</summary>
    Private Shared Function CellText(c As Trs80Tool.MapCellJson) As String
        Dim kind As String
        Select Case If(c.Kind, "")
            Case "free" : kind = UiText.T("dlg.granule.kind.free")
            Case "file" : kind = UiText.T("dlg.granule.kind.file")
            Case "otherFile" : kind = UiText.T("dlg.granule.kind.otherfile")
            Case "directory" : kind = UiText.T("dlg.granule.kind.directory")
            Case "boot" : kind = UiText.T("dlg.granule.kind.boot")
            Case "lockedOut" : kind = UiText.T("dlg.granule.kind.lockedout")
            Case "shared" : kind = UiText.T("dlg.granule.kind.shared")
            Case "claimedButFree" : kind = UiText.T("dlg.granule.kind.claimedbutfree")
            Case "unclaimed" : kind = UiText.T("dlg.granule.kind.unclaimed")
            Case "notOnImage" : kind = UiText.T("dlg.granule.kind.notonimage")
            Case "pastEnd" : kind = UiText.T("dlg.granule.kind.pastend")
            Case Else : kind = c.Kind
        End Select
        Dim sb As New StringBuilder()
        sb.Append(UiText.T("dlg.granule.cell", c.Lump.ToString(CultureInfo.InvariantCulture), c.Granule.ToString(CultureInfo.InvariantCulture), kind))
        If c.Owners IsNot Nothing AndAlso c.Owners.Count > 0 Then
            Dim names As New List(Of String)
            For Each o In c.Owners
                names.Add(o.Replace("."c, "/"c))
            Next
            sb.AppendLine().Append(UiText.T("dlg.granule.claimedby", String.Join(", ", names)))
        End If
        If c.Sectors IsNot Nothing AndAlso c.Sectors.Count > 0 Then
            Dim first As Trs80Tool.SectorPositionJson = c.Sectors(0)
            Dim last As Trs80Tool.SectorPositionJson = c.Sectors(c.Sectors.Count - 1)
            sb.AppendLine().Append(UiText.T("dlg.granule.sectors", first.Cylinder.ToString(CultureInfo.InvariantCulture), first.Side.ToString(CultureInfo.InvariantCulture), first.Sector.ToString(CultureInfo.InvariantCulture)))
            If c.Sectors.Count > 1 Then
                ' Two whole alternatives, so a translation can order the words as its grammar needs:
                ' the last sector alone when it is on the same cylinder and side, otherwise all three.
                If last.Cylinder <> first.Cylinder OrElse last.Side <> first.Side Then
                    sb.Append(UiText.T("dlg.granule.sectors.tofull", last.Cylinder.ToString(CultureInfo.InvariantCulture), last.Side.ToString(CultureInfo.InvariantCulture), last.Sector.ToString(CultureInfo.InvariantCulture)))
                Else
                    sb.Append(UiText.T("dlg.granule.sectors.to", last.Sector.ToString(CultureInfo.InvariantCulture)))
                End If
            End If
            sb.Append(UiText.T("dlg.granule.sectors.count", c.Sectors.Count.ToString(CultureInfo.InvariantCulture)))
        End If
        Return sb.ToString()
    End Function

    Private Sub ShowSelectedCell()
        Dim cell As DataGridViewCell = _grid.CurrentCell
        Dim c As Trs80Tool.MapCellJson = If(cell Is Nothing, Nothing, TryCast(cell.Tag, Trs80Tool.MapCellJson))
        _details.Text = If(c Is Nothing, "", CellText(c))
    End Sub

    ''' <summary>Clicking a file in the list moves to its first granule in the grid.</summary>
    Private Sub SelectFirstCellOfFile()
        If _files.SelectedItems.Count = 0 OrElse _map Is Nothing Then Return
        Dim index As Integer = CInt(_files.SelectedItems(0).Tag)
        For Each row As DataGridViewRow In _grid.Rows
            For Each cell As DataGridViewCell In row.Cells
                Dim c As Trs80Tool.MapCellJson = TryCast(cell.Tag, Trs80Tool.MapCellJson)
                If c IsNot Nothing AndAlso c.File.HasValue AndAlso c.File.Value = index Then
                    _grid.CurrentCell = cell
                    Return
                End If
            Next
        Next
    End Sub

    Private Sub ShowTextMap()
        If _mapText.Length = 0 Then Return
        Dim f As New TextReportForm(UiText.T("dlg.granule.texttitle", Path.GetFileName(_disk.Path)), _mapText.Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf))
        f.Show(Me)
    End Sub
End Class
