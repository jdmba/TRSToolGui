' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  GranuleMapForm.vb — Disk → Granule map in the Eto.Forms program: what
'  trs80-tool's "disk map" reports, drawn as a grid of one row per lump and
'  one column per granule. Each granule shows what it holds: a file (a number,
'  listed at the right), the directory, the boot sector, free space, or
'  something wrong.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class GranuleMapForm
    Inherits ToolWindow

    ''' <summary>One row of the map: a lump (a cylinder on most disks) and its granules by number.</summary>
    Private Class LumpRow
        Public Property Lump As Integer
        Public ReadOnly Property Cells As New Dictionary(Of Integer, Trs80Tool.MapCellJson)
    End Class

    ''' <summary>One row of the file list.</summary>
    Private Class FileRow
        Public Property Index As Integer
        Public Property Number As String = ""
        Public Property Name As String = ""
        Public Property Granules As String = ""
        Public Property HasColour As Boolean
        Public Property Colour As Color
    End Class

    Private ReadOnly _disk As DiskImage
    Private _map As Trs80Tool.MapResultJson = Nothing
    Private _mapText As String = ""

    Private ReadOnly _txtFilter As New TextBox With {.Width = 260}
    Private ReadOnly _btnApply As New Button With {.Text = "Show"}
    Private ReadOnly _lblSummary As New Label With {.Text = "Reading the granule map...", .Wrap = WrapMode.Word}
    Private ReadOnly _grid As New GridView With {.GridLines = GridLines.Both, .AllowMultipleSelection = False}
    Private ReadOnly _files As New GridView With {.AllowMultipleSelection = False}
    Private ReadOnly _details As New TextArea With {.ReadOnly = True, .Wrap = True}
    Private ReadOnly _btnRefresh As New Button With {.Text = "Refresh"}
    Private _rows As New List(Of LumpRow)
    Private ReadOnly _columnGranule As New Dictionary(Of GridColumn, Integer)
    Private ReadOnly _monoFont As Font = Ui.Mono(9.0F)

    ''' <summary>Twelve pastel backgrounds, cycled by the colour number trs80-tool gives each file.</summary>
    Private Shared ReadOnly FileColors As Color() = {
        Color.FromArgb(174, 214, 241), Color.FromArgb(171, 235, 198), Color.FromArgb(249, 231, 159), Color.FromArgb(245, 183, 177),
        Color.FromArgb(215, 189, 226), Color.FromArgb(250, 215, 160), Color.FromArgb(162, 217, 206), Color.FromArgb(241, 196, 223),
        Color.FromArgb(200, 220, 150), Color.FromArgb(190, 200, 235), Color.FromArgb(235, 205, 185), Color.FromArgb(180, 225, 230)}

    Public Sub New(disk As DiskImage)
        _disk = disk
        Title = "Granule map - " & Path.GetFileName(disk.Path) & If(disk.IsHardDisk AndAlso disk.ActivePartition IsNot Nothing, " - partition " & disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture), "")
        ClientSize = New Size(1020, 660)
        MinimumSize = New Size(640, 400)
        Padding = New Padding(Ui.Margin)
        _details.Font = _monoFont

        AddHandler _grid.CellClick, Sub(s, e) ShowCell(e.Row, e.GridColumn)
        AddHandler _grid.CellFormatting, AddressOf OnMapCellFormatting

        _files.Columns.Add(New GridColumn With {.HeaderText = "#", .Width = 45, .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Right, .Binding = Binding.Delegate(Of FileRow, String)(Function(r) r.Number)}})
        _files.Columns.Add(New GridColumn With {.HeaderText = "File", .Width = 130, .DataCell = New TextBoxCell With {.Binding = Binding.Delegate(Of FileRow, String)(Function(r) r.Name)}})
        _files.Columns.Add(New GridColumn With {.HeaderText = "Granules", .Width = 80, .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Right, .Binding = Binding.Delegate(Of FileRow, String)(Function(r) r.Granules)}})
        AddHandler _files.CellFormatting, Sub(s, e)
                                              Dim r As FileRow = TryCast(e.Item, FileRow)
                                              If r IsNot Nothing AndAlso r.HasColour AndAlso e.Column Is _files.Columns(0) Then
                                                  e.BackgroundColor = r.Colour
                                                  e.ForegroundColor = Colors.Black
                                              End If
                                          End Sub
        AddHandler _files.SelectionChanged, Sub() SelectFirstCellOfFile()

        Dim right As New Splitter With {.Orientation = Orientation.Vertical, .Panel1 = _files, .Panel2 = _details, .Position = 360}
        Dim split As New Splitter With {.Orientation = Orientation.Horizontal, .Panel1 = _grid, .Panel2 = right, .Position = 700}

        AddHandler _btnApply.Click, Async Sub() Await LoadAsync()
        AddHandler _btnRefresh.Click, Async Sub() Await LoadAsync()
        AddHandler _txtFilter.KeyDown, Async Sub(s, e)
                                           If e.Key = Keys.Enter Then
                                               e.Handled = True
                                               Await LoadAsync()
                                           End If
                                       End Sub
        Content = Ui.Down(split,
                          Ui.Across(New Label With {.Text = "Only these files (TRS-80 name patterns, e.g. *.BAS SYS?.SYS):", .VerticalAlignment = VerticalAlignment.Center}, _txtFilter, _btnApply),
                          _lblSummary,
                          split,
                          Ui.Across(Ui.Btn("trs80-tool's text map...", AddressOf ShowTextMap), _btnRefresh, Nothing, Ui.Btn("Close", Sub() Close())))
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Await LoadAsync()
    End Sub

    Private Function Patterns() As List(Of String)
        Dim pats As New List(Of String)
        If Not String.IsNullOrWhiteSpace(_txtFilter.Text) Then pats.AddRange(_txtFilter.Text.Split(New Char() {" "c, ","c}, StringSplitOptions.RemoveEmptyEntries))
        Return pats
    End Function

    Private Async Function LoadAsync() As Task
        _btnApply.Enabled = False
        _btnRefresh.Enabled = False
        Try
            Dim r As ToolResult = Await Program.Tool.DiskMapAsync(_disk.Path, _disk.PartitionArg, Patterns())
            _map = If(r.Succeeded AndAlso r.Disk IsNot Nothing, r.Disk.ResultOf(Of Trs80Tool.MapResultJson)("map"), Nothing)
            _mapText = If(r.Succeeded, r.StdOut, "")
            If _map Is Nothing OrElse _map.Cells Is Nothing Then
                _grid.DataStore = Nothing
                _grid.Columns.Clear()
                _files.DataStore = Nothing
                _details.Text = ""
                _lblSummary.Text = "trs80-tool can't draw this disk's granules: " & r.ErrorText
                Return
            End If
            FillGrid()
            FillFiles()
            _lblSummary.Text = SummaryText()
        Catch ex As Exception
            _lblSummary.Text = ex.Message
        Finally
            _btnApply.Enabled = True
            _btnRefresh.Enabled = True
        End Try
    End Function

    Private Function SummaryText() As String
        Dim t As Trs80Tool.MapTotalsJson = _map.Totals
        If t Is Nothing Then Return ""
        Dim used As Integer = t.GranuleCount - t.FreeGranuleCount - t.NotOnImageGranuleCount - t.PastEndGranuleCount
        Dim sb As New StringBuilder()
        sb.Append(used.ToString("N0", CultureInfo.InvariantCulture)).Append(" of ").Append(t.GranuleCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" granules used, ")
        sb.Append(t.FreeGranuleCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" free (").Append(TextUtil.FormatKb(t.FreeByteCount)).Append("). A granule is ")
        sb.Append(t.GranuleByteCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" bytes.")
        If t.NotOnImageGranuleCount > 0 Then sb.Append(" ").Append(t.NotOnImageGranuleCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" free granule(s) aren't fully on the image.")
        If t.PastEndGranuleCount > 0 Then sb.Append(" ").Append(t.PastEndGranuleCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" are past the disk's proven end.")
        If t.LockedOutGranuleCount > 0 Then sb.Append(" ").Append(t.LockedOutGranuleCount.ToString("N0", CultureInfo.InvariantCulture)).Append(" locked out.")
        sb.Append(vbLf)
        sb.Append("Number = a file (listed at the right)   · free   B boot   D directory   - another file   x locked out   ! shared by files   ? claimed but free   u used, unclaimed   n not on the image   p past the end")
        Return sb.ToString()
    End Function

    ''' <summary>What a granule shows in its cell: a file's number, or a letter for its kind.</summary>
    Private Function CellLetter(c As Trs80Tool.MapCellJson) As String
        Select Case If(c.Kind, "")
            Case "free" : Return "·"
            Case "file"
                Dim index As Integer = If(c.File, -1)
                Return If(index >= 0, (index + 1).ToString(CultureInfo.InvariantCulture), "f")
            Case "otherFile" : Return "-"
            Case "directory" : Return "D"
            Case "boot" : Return "B"
            Case "lockedOut" : Return "x"
            Case "shared" : Return "!"
            Case "claimedButFree" : Return "?"
            Case "unclaimed" : Return "u"
            Case "notOnImage" : Return "n"
            Case "pastEnd" : Return "p"
            Case Else : Return "*"       ' a kind a later trs80-tool added
        End Select
    End Function

    Private Sub OnMapCellFormatting(sender As Object, e As GridCellFormatEventArgs)
        e.Font = _monoFont
        Dim row As LumpRow = TryCast(e.Item, LumpRow)
        Dim granule As Integer
        If row Is Nothing OrElse e.Column Is Nothing OrElse Not _columnGranule.TryGetValue(e.Column, granule) Then Return
        Dim c As Trs80Tool.MapCellJson = Nothing
        If Not row.Cells.TryGetValue(granule, c) Then Return
        Select Case If(c.Kind, "")
            Case "free" : e.ForegroundColor = Ui.GrayText
            Case "file"
                Dim index As Integer = If(c.File, -1)
                Dim colorIndex As Integer = index
                If index >= 0 AndAlso _map.Files IsNot Nothing AndAlso index < _map.Files.Count AndAlso _map.Files(index).ColorIndex.HasValue Then colorIndex = _map.Files(index).ColorIndex.Value
                If colorIndex >= 0 Then e.BackgroundColor = FileColors(colorIndex Mod FileColors.Length) : e.ForegroundColor = Colors.Black
            Case "otherFile" : e.BackgroundColor = Color.FromArgb(232, 232, 232) : e.ForegroundColor = Colors.Black
            Case "directory", "boot" : e.BackgroundColor = Color.FromArgb(210, 210, 210) : e.ForegroundColor = Colors.Black
            Case "lockedOut", "pastEnd" : e.ForegroundColor = Ui.GrayText
            Case "shared" : e.BackgroundColor = Color.FromArgb(240, 128, 128) : e.ForegroundColor = Colors.Black
            Case "claimedButFree" : e.BackgroundColor = Color.FromArgb(255, 190, 120) : e.ForegroundColor = Colors.Black
            Case "unclaimed" : e.BackgroundColor = Color.FromArgb(255, 235, 150) : e.ForegroundColor = Colors.Black
            Case "notOnImage" : e.ForegroundColor = Colors.Firebrick
        End Select
    End Sub

    Private Sub FillGrid()
        _grid.DataStore = Nothing
        _grid.Columns.Clear()
        _columnGranule.Clear()
        Dim maxGranule As Integer = 0
        Dim lumps As New SortedDictionary(Of Integer, LumpRow)
        For Each c In _map.Cells
            If c.Granule > maxGranule Then maxGranule = c.Granule
            Dim row As LumpRow = Nothing
            If Not lumps.TryGetValue(c.Lump, row) Then
                row = New LumpRow With {.Lump = c.Lump}
                lumps(c.Lump) = row
            End If
            row.Cells(c.Granule) = c
        Next
        ' A lump is a cylinder on a floppy and on trs80-tool's hard-disk partitions; the layout says which word fits.
        Dim rowsAreCylinders As Boolean = _map.Layout IsNot Nothing AndAlso String.Equals(_map.Layout.LumpTerm, "cylinder", StringComparison.OrdinalIgnoreCase)
        _grid.Columns.Add(New GridColumn With {.HeaderText = If(rowsAreCylinders, "Cylinder", "Lump"), .Width = If(rowsAreCylinders, 80, 60),
                                               .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Right, .Binding = Binding.Delegate(Of LumpRow, String)(Function(r) r.Lump.ToString(CultureInfo.InvariantCulture))}})
        For g As Integer = 0 To maxGranule
            Dim granule As Integer = g
            Dim col As New GridColumn With {.HeaderText = granule.ToString(CultureInfo.InvariantCulture), .Width = 38,
                                            .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Center, .Binding = Binding.Delegate(Of LumpRow, String)(
                                                Function(r)
                                                    Dim c As Trs80Tool.MapCellJson = Nothing
                                                    Return If(r.Cells.TryGetValue(granule, c), CellLetter(c), "")
                                                End Function)}}
            _grid.Columns.Add(col)
            _columnGranule(col) = granule
        Next
        _rows = lumps.Values.ToList()
        _grid.DataStore = _rows
        _details.Text = ""
    End Sub

    Private Sub FillFiles()
        Dim rows As New List(Of FileRow)
        If _map.Files IsNot Nothing Then
            For i As Integer = 0 To _map.Files.Count - 1
                Dim f As Trs80Tool.MapFileJson = _map.Files(i)
                Dim row As New FileRow With {.Index = i, .Number = (i + 1).ToString(CultureInfo.InvariantCulture), .Name = If(f.Name, "").Replace("."c, "/"c), .Granules = f.GranuleCount.ToString("N0", CultureInfo.InvariantCulture)}
                If f.ColorIndex.HasValue Then row.HasColour = True : row.Colour = FileColors(f.ColorIndex.Value Mod FileColors.Length)
                rows.Add(row)
            Next
        End If
        _files.DataStore = rows
    End Sub

    Private Shared Function CellText(c As Trs80Tool.MapCellJson) As String
        Dim kind As String
        Select Case If(c.Kind, "")
            Case "free" : kind = "free"
            Case "file" : kind = "a file"
            Case "otherFile" : kind = "another file (not one of those asked for)"
            Case "directory" : kind = "the directory"
            Case "boot" : kind = "the boot sector's granule"
            Case "lockedOut" : kind = "locked out (the lockout table says it can't be used)"
            Case "shared" : kind = "DAMAGE: two or more files claim it"
            Case "claimedButFree" : kind = "DAMAGE: a file claims it, but the allocation table says it is free"
            Case "unclaimed" : kind = "marked used in the allocation table, but no file claims it"
            Case "notOnImage" : kind = "free, but the image doesn't fully hold it, so new files skip it"
            Case "pastEnd" : kind = "free, but past the disk's proven size, so new files never get it"
            Case Else : kind = c.Kind
        End Select
        Dim sb As New StringBuilder()
        sb.Append("Lump ").Append(c.Lump.ToString(CultureInfo.InvariantCulture)).Append(", granule ").Append(c.Granule.ToString(CultureInfo.InvariantCulture)).Append(": ").Append(kind)
        If c.Owners IsNot Nothing AndAlso c.Owners.Count > 0 Then
            Dim names As New List(Of String)
            For Each o In c.Owners
                names.Add(o.Replace("."c, "/"c))
            Next
            sb.Append(vbLf).Append("Claimed by: ").Append(String.Join(", ", names))
        End If
        If c.Sectors IsNot Nothing AndAlso c.Sectors.Count > 0 Then
            Dim first As Trs80Tool.SectorPositionJson = c.Sectors(0)
            Dim last As Trs80Tool.SectorPositionJson = c.Sectors(c.Sectors.Count - 1)
            sb.Append(vbLf).Append("Sectors: cylinder ").Append(first.Cylinder.ToString(CultureInfo.InvariantCulture)).Append(", side ").Append(first.Side.ToString(CultureInfo.InvariantCulture)).Append(", sector ").Append(first.Sector.ToString(CultureInfo.InvariantCulture))
            If c.Sectors.Count > 1 Then
                sb.Append(" to ")
                If last.Cylinder <> first.Cylinder OrElse last.Side <> first.Side Then sb.Append("cylinder ").Append(last.Cylinder.ToString(CultureInfo.InvariantCulture)).Append(", side ").Append(last.Side.ToString(CultureInfo.InvariantCulture)).Append(", sector ")
                sb.Append(last.Sector.ToString(CultureInfo.InvariantCulture))
            End If
            sb.Append(" (").Append(c.Sectors.Count.ToString(CultureInfo.InvariantCulture)).Append(" sectors)")
        End If
        Return sb.ToString()
    End Function

    ''' <summary>A click on the map: describe the granule in that row and column.</summary>
    Private Sub ShowCell(rowIndex As Integer, column As GridColumn)
        If rowIndex < 0 OrElse rowIndex >= _rows.Count OrElse column Is Nothing Then Return
        Dim granule As Integer
        Dim c As Trs80Tool.MapCellJson = Nothing
        If _columnGranule.TryGetValue(column, granule) AndAlso _rows(rowIndex).Cells.TryGetValue(granule, c) Then
            _details.Text = CellText(c)
        Else
            _details.Text = ""
        End If
    End Sub

    ''' <summary>A file picked in the list: go to the row of its first granule and describe that granule.</summary>
    Private Sub SelectFirstCellOfFile()
        Dim picked As FileRow = TryCast(_files.SelectedItem, FileRow)
        If picked Is Nothing OrElse _map Is Nothing Then Return
        For i As Integer = 0 To _rows.Count - 1
            For Each kv In _rows(i).Cells.OrderBy(Function(p) p.Key)
                Dim c As Trs80Tool.MapCellJson = kv.Value
                If c.File.HasValue AndAlso c.File.Value = picked.Index Then
                    Try
                        _grid.SelectRow(i)
                        _grid.ScrollToRow(i)
                    Catch
                    End Try
                    _details.Text = CellText(c)
                    Return
                End If
            Next
        Next
    End Sub

    Private Sub ShowTextMap()
        If _mapText.Length = 0 Then Return
        Dim f As New TextReportForm("Granule map (disk map) - " & Path.GetFileName(_disk.Path), _mapText)
        f.ShowOwned(Me)
    End Sub
End Class
