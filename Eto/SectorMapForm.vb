' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  SectorMapForm.vb — Disk → Sector map in the Eto.Forms program (trs80-tool
'  "sectors"). A grid of every sector of a floppy image (one row per cylinder
'  and side, one column per sector number), coloured by density and status as
'  trs80-tool's own legend does it; the selected sector's header line, files
'  and hex dump; and the list of bad sectors with the files that use them.
'  Everything comes from one "sectors --json --contents" call.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class SectorMapForm
    Inherits ToolWindow

    ''' <summary>One row of the map: a cylinder and side, with its sectors by sector number.</summary>
    Private Class MapRow
        Public Property Cylinder As Integer
        Public Property Side As Integer
        Public ReadOnly Property Sectors As New Dictionary(Of Integer, Trs80Tool.SectorJson)
        Public Function Letter(number As Integer) As String
            Dim sj As Trs80Tool.SectorJson = Nothing
            Return If(Sectors.TryGetValue(number, sj), sj.MapLetter, "-")
        End Function
    End Class

    ''' <summary>One row of the bad-sector list.</summary>
    Private Class BadRow
        Public Property Sector As Trs80Tool.SectorJson
        Public Property Cyl As String = ""
        Public Property Side As String = ""
        Public Property Number As String = ""
        Public Property Problem As String = ""
        Public Property Files As String = ""
    End Class

    Private ReadOnly _disk As DiskImage
    Private _info As Trs80Tool.SectorsInfileJson = Nothing

    Private ReadOnly _grid As New GridView With {.GridLines = GridLines.Horizontal, .AllowMultipleSelection = False}
    Private ReadOnly _details As TextArea
    Private ReadOnly _bad As New GridView With {.AllowMultipleSelection = False}
    Private ReadOnly _lblSummary As New Label With {.Text = "Reading sectors..."}
    Private ReadOnly _btnRefresh As New Button With {.Text = "Refresh"}
    Private ReadOnly _btnProtection As New Button With {.Text = "Protection && artifacts...", .Enabled = False}
    Private _rows As New List(Of MapRow)
    Private ReadOnly _columnNumber As New Dictionary(Of GridColumn, Integer)
    Private ReadOnly _monoFont As Font = Ui.Mono(9.0F)
    Private Shared ReadOnly DoubleBack As Color = Color.FromArgb(230, 238, 250)

    Public Sub New(disk As DiskImage)
        _disk = disk
        _details = Ui.ReportBox("", 9.0F)
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Title = "Sector map - " & Path.GetFileName(_disk.Path)
        ClientSize = New Size(1040, 660)
        MinimumSize = New Size(640, 400)
        Padding = New Padding(Ui.Margin)

        AddHandler _grid.CellClick, Sub(s, e) ShowCell(e.Row, e.GridColumn)
        AddHandler _grid.CellFormatting, AddressOf OnMapCellFormatting

        AddBadColumn("Cyl", 45, Function(r) r.Cyl)
        AddBadColumn("Side", 45, Function(r) r.Side)
        AddBadColumn("Sector", 60, Function(r) r.Number)
        AddBadColumn("Problem", 260, Function(r) r.Problem)
        AddBadColumn("File(s) at this sector", 220, Function(r) r.Files)
        AddHandler _bad.SelectionChanged, Sub() SelectFromBadList()

        Dim right As New Splitter With {
            .Orientation = Orientation.Vertical,
            .Panel1 = Ui.Down(_details, Ui.Lbl("Selected sector:"), _details),
            .Panel2 = Ui.Down(_bad, Ui.Lbl("Bad sectors (missing or CRC error) and the files that use them — what ""sectors --bad"" lists:"), _bad),
            .Position = 330}
        Dim split As New Splitter With {.Orientation = Orientation.Horizontal, .Panel1 = _grid, .Panel2 = right, .Position = 560}

        AddHandler _btnRefresh.Click, Async Sub() Await LoadAsync()
        AddHandler _btnProtection.Click, Sub() ShowProtection()
        Dim buttons As TableLayout = Ui.Across(Ui.Btn("Text report...", Async Sub() Await ShowTextReportAsync(False)),
                                               Ui.Btn("Bad sectors (text)...", Async Sub() Await ShowTextReportAsync(True)),
                                               Ui.Btn("Save report...", Async Sub() Await SaveReportAsync()),
                                               _btnRefresh, _btnProtection, Nothing, Ui.Btn("Close", Sub() Close()))
        Content = Ui.Down(split, _lblSummary,
                          Ui.Lbl("Legend (as trs80-tool prints it):  S single density   D double density   C CRC error   X deleted-data mark   ? not found   - not on this track.  Click a sector for its header line and hex dump.", True),
                          split, buttons)
    End Sub

    Private Sub AddBadColumn(header As String, width As Integer, getter As Func(Of BadRow, String))
        _bad.Columns.Add(New GridColumn With {.HeaderText = header, .Width = width, .Resizable = True, .DataCell = New TextBoxCell With {.Binding = Binding.Delegate(Of BadRow, String)(getter)}})
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Await LoadAsync()
    End Sub

    Private Async Function LoadAsync() As Task
        _btnRefresh.Enabled = False
        _lblSummary.Text = "Reading sectors with trs80-tool..."
        Dim problem As String = Nothing
        Try
            ' --protection (5.1.0) adds trs80-tool's signs of copy protection and its imaging artifacts.
            Dim caps As EngineCapabilities = Program.Tool.Capabilities
            Dim wantProtection As Boolean = caps IsNot Nothing AndAlso caps.HasTopLevelOption("sectors", "--protection")
            _info = Await Program.Tool.SectorsInfileAsync(_disk.Path, True, False, wantProtection)
        Catch ex As Exception
            problem = ex.Message
        End Try
        _btnRefresh.Enabled = True
        If problem IsNot Nothing Then
            _lblSummary.Text = "trs80-tool could not read the sectors."
            Ui.Fail(Me, "trs80-tool could not read the sectors of this image:" & vbLf & vbLf & problem, "Sector map")
            Return
        End If
        FillGrid()
        FillBadList()
        Dim signs As Integer = If(_info.Protection Is Nothing, 0, _info.Protection.Count)
        Dim artifacts As Integer = If(_info.ImagingArtifacts Is Nothing, 0, _info.ImagingArtifacts.Count)
        _btnProtection.Text = "Protection && artifacts (" & (signs + artifacts).ToString(CultureInfo.InvariantCulture) & ")..."
        _btnProtection.Enabled = _info.Protection IsNot Nothing OrElse artifacts > 0
    End Function

    ''' <summary>trs80-tool's signs of copy protection and imaging artifacts, in its own words.</summary>
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
        f.ShowOwned(Me)
    End Sub

    Private Sub FillGrid()
        _grid.DataStore = Nothing
        _grid.Columns.Clear()
        _columnNumber.Clear()
        _rows = New List(Of MapRow)
        Dim sectors As List(Of Trs80Tool.SectorJson) = _info.Sectors
        If sectors.Count = 0 Then
            _lblSummary.Text = "No sectors reported for this image."
            Return
        End If
        ' Columns: every sector number that occurs anywhere, ascending (protected disks may have odd ones).
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
        _grid.Columns.Add(New GridColumn With {.HeaderText = "Cyl", .Width = 45, .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Right, .Binding = Binding.Delegate(Of MapRow, String)(Function(r) r.Cylinder.ToString(CultureInfo.InvariantCulture))}})
        _grid.Columns.Add(New GridColumn With {.HeaderText = "Side", .Width = 45, .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Center, .Binding = Binding.Delegate(Of MapRow, String)(Function(r) r.Side.ToString(CultureInfo.InvariantCulture))}})
        For Each n In numbers
            Dim number As Integer = n
            Dim col As New GridColumn With {.HeaderText = number.ToString(CultureInfo.InvariantCulture), .Width = If(number >= 100, 38, 30),
                                            .DataCell = New TextBoxCell With {.TextAlignment = TextAlignment.Center, .Binding = Binding.Delegate(Of MapRow, String)(Function(r) r.Letter(number))}}
            _grid.Columns.Add(col)
            _columnNumber(col) = number
        Next

        Dim byKey As New Dictionary(Of Integer, MapRow)
        For cyl As Integer = 0 To maxCyl
            For side As Integer = 0 To maxSide
                Dim row As New MapRow With {.Cylinder = cyl, .Side = side}
                _rows.Add(row)
                byKey(cyl * 4 + side) = row
            Next
        Next
        For Each sj In sectors
            Dim row As MapRow = Nothing
            If byKey.TryGetValue(sj.Cylinder * 4 + sj.Side, row) Then row.Sectors(sj.Sector) = sj
        Next
        _grid.DataStore = _rows

        Dim desc As String = If(_info.Description, "")
        Dim os As String = If(_info.Os, "unknown operating system")
        _lblSummary.Text = desc & If(desc.Length > 0, ", ", "") & os & "  —  " & (maxCyl + 1).ToString(CultureInfo.InvariantCulture) & " cylinders, " & (maxSide + 1).ToString(CultureInfo.InvariantCulture) & " side(s), " &
            sectors.Count.ToString(CultureInfo.InvariantCulture) & " sector positions, " & If(bad = 0, "no bad sectors", bad.ToString(CultureInfo.InvariantCulture) & " bad sector(s)") &
            If(stray > 0, ", " & stray.ToString(CultureInfo.InvariantCulture) & " outside the usual numbers", "")
    End Sub

    ''' <summary>The map's colours: bad sectors red, deleted-data marks orange, FM white, MFM pale blue, stray numbers yellow.</summary>
    Private Sub OnMapCellFormatting(sender As Object, e As GridCellFormatEventArgs)
        e.Font = _monoFont
        Dim row As MapRow = TryCast(e.Item, MapRow)
        Dim number As Integer
        If row Is Nothing OrElse e.Column Is Nothing OrElse Not _columnNumber.TryGetValue(e.Column, number) Then Return
        Dim sj As Trs80Tool.SectorJson = Nothing
        If Not row.Sectors.TryGetValue(number, sj) Then
            e.ForegroundColor = Ui.GrayText
            Return
        End If
        Select Case sj.MapLetter
            Case "?"
                e.BackgroundColor = Colors.Firebrick : e.ForegroundColor = Colors.White
            Case "C"
                e.BackgroundColor = Colors.Red : e.ForegroundColor = Colors.White
            Case "X"
                e.BackgroundColor = Colors.Orange : e.ForegroundColor = Colors.Black
            Case "S"
                e.BackgroundColor = Colors.White : e.ForegroundColor = Colors.Black
            Case Else
                e.BackgroundColor = DoubleBack : e.ForegroundColor = Colors.Black
        End Select
        If sj.OutsideUsualNumbers AndAlso Not sj.IsBad Then e.BackgroundColor = Colors.LemonChiffon
    End Sub

    Private Sub FillBadList()
        Dim rows As New List(Of BadRow)
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
            ' trs80-tool says when a bad sector is something the imaging likely caused, so reimaging may read it.
            If sj.ImagingArtifactText.Length > 0 Then problem &= " - imaging artifact: " & sj.ImagingArtifactText
            rows.Add(New BadRow With {
                .Sector = sj,
                .Cyl = sj.Cylinder.ToString(CultureInfo.InvariantCulture),
                .Side = sj.Side.ToString(CultureInfo.InvariantCulture),
                .Number = sj.Sector.ToString(CultureInfo.InvariantCulture),
                .Problem = problem,
                .Files = If(sj.Files IsNot Nothing AndAlso sj.Files.Count > 0, String.Join(", ", sj.Files), "(no file at this sector)")})
        Next
        If rows.Count = 0 Then rows.Add(New BadRow With {.Problem = "No bad sectors."})
        _bad.DataStore = rows
    End Sub

    ''' <summary>A click on the map: show the sector in that row and column.</summary>
    Private Sub ShowCell(rowIndex As Integer, column As GridColumn)
        If rowIndex < 0 OrElse rowIndex >= _rows.Count OrElse column Is Nothing Then Return
        Dim number As Integer
        If Not _columnNumber.TryGetValue(column, number) Then Return
        Dim row As MapRow = _rows(rowIndex)
        Dim sj As Trs80Tool.SectorJson = Nothing
        If row.Sectors.TryGetValue(number, sj) Then
            ShowSector(sj)
        Else
            _details.Text = "Cylinder " & row.Cylinder.ToString(CultureInfo.InvariantCulture) & ", side " & row.Side.ToString(CultureInfo.InvariantCulture) & ", sector " & number.ToString(CultureInfo.InvariantCulture) & ": not on this track"
        End If
    End Sub

    Private Sub ShowSector(sj As Trs80Tool.SectorJson)
        Dim sb As New StringBuilder()
        sb.Append(sj.HeaderText).Append(vbLf)
        If sj.ImagingArtifactText.Length > 0 Then sb.Append("Imaging artifact: " & sj.ImagingArtifactText & ".").Append(vbLf)
        If sj.Files IsNot Nothing AndAlso sj.Files.Count > 0 Then
            sb.Append(If(sj.Files.Count = 1, "File", "Files") & " at this sector: " & String.Join(", ", sj.Files)).Append(vbLf)
        Else
            sb.Append("There is no file at this sector.").Append(vbLf)
        End If
        Dim data As Byte() = sj.GetBytes()
        If data IsNot Nothing Then
            sb.Append(vbLf)
            sb.Append(Ui.Lf(TextUtil.HexDump(data)))
        End If
        _details.Text = sb.ToString()
        Try
            _details.CaretIndex = 0
        Catch
        End Try
    End Sub

    ''' <summary>A bad sector picked in the list: go to its row in the map and show it.</summary>
    Private Sub SelectFromBadList()
        Dim picked As BadRow = TryCast(_bad.SelectedItem, BadRow)
        If picked Is Nothing OrElse picked.Sector Is Nothing Then Return
        Dim sj As Trs80Tool.SectorJson = picked.Sector
        For i As Integer = 0 To _rows.Count - 1
            If _rows(i).Cylinder = sj.Cylinder AndAlso _rows(i).Side = sj.Side Then
                Try
                    _grid.SelectRow(i)
                    _grid.ScrollToRow(i)
                Catch
                End Try
                Exit For
            End If
        Next
        ShowSector(sj)
    End Sub

    Private Async Function ShowTextReportAsync(badOnly As Boolean) As Task
        Dim r As ToolResult
        Try
            r = Await Program.Tool.SectorsTextAsync(_disk.Path, badOnly)
        Catch ex As Exception
            r = Nothing
            Ui.Fail(Me, ex.Message, "Sector map")
        End Try
        If r Is Nothing Then Return
        Dim body As String = If(r.Succeeded, r.StdOut, r.ErrorText)
        If badOnly AndAlso body.Trim().Length = 0 Then body = "trs80-tool reports no bad sectors on this floppy (""sectors --bad"" prints nothing for a clean disk)."
        Dim f As New TextReportForm(If(badOnly, "sectors --bad", "sectors") & " - " & Path.GetFileName(_disk.Path), body)
        f.ShowOwned(Me)
    End Function

    Private Async Function SaveReportAsync() As Task
        Dim target As String = Ui.PickSave(Me, "Save sector report", "Text file (*.txt)|*.txt|All files (*.*)|*.*",
                                           Path.Combine(Path.GetDirectoryName(_disk.Path), Path.GetFileNameWithoutExtension(_disk.Path) & "_sectors.txt"), "txt")
        If target Is Nothing Then Return
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
            File.WriteAllText(target, Ui.HostLines(sb.ToString()), New UTF8Encoding(True))
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, "Save sector report")
        End Try
    End Function
End Class
