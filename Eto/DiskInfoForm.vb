' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskInfoForm.vb — Disk → DMK analysis in the Eto.Forms program. It shows
'  what DmkAnalyzer and DosIdentifier read from the DMK file itself: DOS,
'  model and version from the boot sector, directory and GAT bytes; the
'  geometry track by track; protection signs, damage and imaging artifacts.
'  Read-only; nothing here calls trs80-tool.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class DiskInfoForm
    Inherits ToolWindow

    ''' <summary>A line of the Summary tab: a fact, a finding, or the heading of a group of them.</summary>
    Private Class SummaryRow
        Public Property Item As String = ""
        Public Property Value As String = ""
        Public Property IsHeading As Boolean
        Public Property HasColour As Boolean
        Public Property Colour As Color
    End Class

    ''' <summary>A line of the Tracks tab.</summary>
    Private Class TrackRow
        Public Property Cyl As String = ""
        Public Property Side As String = ""
        Public Property Ids As String = ""
        Public Property Layout As String = ""
        Public Property Bad As String = ""
        Public Property Notes As String = ""
        Public Property HasColour As Boolean
        Public Property Colour As Color
    End Class

    Private ReadOnly _path As String
    Private _analysis As DmkAnalysis = Nothing

    Private ReadOnly _summary As New GridView With {.ShowHeader = True, .GridLines = GridLines.None}
    Private ReadOnly _tracks As New GridView With {.ShowHeader = True, .GridLines = GridLines.Horizontal}
    Private ReadOnly _txtReport As TextArea
    Private ReadOnly _status As New Label With {.VerticalAlignment = VerticalAlignment.Center}
    Private ReadOnly _boldFont As Font = SystemFonts.Bold()
    Private ReadOnly _monoFont As Font = Ui.Mono(9.0F)

    Public Sub New(imagePath As String)
        _path = imagePath
        _txtReport = Ui.ReportBox("", 9.0F)
        BuildUi()
    End Sub

    ''' <summary>Reads and analyses the image on a worker thread, then fills the tabs.</summary>
    Public Async Function LoadAsync() As Task
        _status.Text = "Reading " & Path.GetFileName(_path) & "..."
        Try
            Dim p As String = _path
            _analysis = Await Task.Run(Function() DmkAnalysis.Analyze(p))
            Fill()
            _status.Text = If(_analysis.Dos IsNot Nothing, _analysis.Dos.Result, _analysis.Image.Kind)
        Catch ex As Exception
            _status.Text = "Could not analyse the image: " & ex.Message
            _txtReport.Text = Ui.Lf(ex.ToString())
        End Try
    End Function

    Private Sub BuildUi()
        Title = "DMK analysis - " & Path.GetFileName(_path)
        ClientSize = New Size(980, 640)
        MinimumSize = New Size(640, 400)
        Padding = New Padding(Ui.Margin)

        _summary.Columns.Add(New GridColumn With {.HeaderText = "Item", .Width = 200, .DataCell = New TextBoxCell With {.Binding = Binding.Delegate(Of SummaryRow, String)(Function(r) r.Item)}})
        _summary.Columns.Add(New GridColumn With {.HeaderText = "Value", .Width = 1400, .DataCell = New TextBoxCell With {.Binding = Binding.Delegate(Of SummaryRow, String)(Function(r) r.Value)}})
        AddHandler _summary.CellFormatting, Sub(s, e)
                                                Dim r As SummaryRow = TryCast(e.Item, SummaryRow)
                                                If r Is Nothing Then Return
                                                If r.IsHeading Then e.Font = _boldFont
                                                If r.HasColour Then e.ForegroundColor = r.Colour
                                            End Sub

        AddTrackColumn("Cyl", 50, Function(r) r.Cyl)
        AddTrackColumn("Side", 50, Function(r) r.Side)
        AddTrackColumn("IDs", 50, Function(r) r.Ids)
        AddTrackColumn("Layout (density × count × size, numbering, marks)", 520, Function(r) r.Layout)
        AddTrackColumn("Bad", 50, Function(r) r.Bad)
        AddTrackColumn("Notes", 400, Function(r) r.Notes)
        AddHandler _tracks.CellFormatting, Sub(s, e)
                                               Dim r As TrackRow = TryCast(e.Item, TrackRow)
                                               If r Is Nothing Then Return
                                               e.Font = _monoFont
                                               If r.HasColour Then e.ForegroundColor = r.Colour
                                           End Sub

        Dim tabs As New TabControl()
        tabs.Pages.Add(New TabPage With {.Text = "Summary", .Content = _summary})
        tabs.Pages.Add(New TabPage With {.Text = "Tracks", .Content = _tracks})
        tabs.Pages.Add(New TabPage With {.Text = "Report", .Content = _txtReport})

        Dim bottom As New TableLayout With {.Spacing = New Size(Ui.Gap, Ui.Gap)}
        bottom.Rows.Add(New TableRow(New TableCell(_status, True),
                                     New TableCell(Ui.Btn("Copy report", Sub()
                                                                             If _txtReport.Text.Length > 0 Then Ui.CopyTextToClipboard(_txtReport.Text)
                                                                         End Sub)),
                                     New TableCell(Ui.Btn("Save report...", AddressOf OnSave)),
                                     New TableCell(Ui.Btn("Close", Sub() Close()))))
        Content = Ui.Down(tabs, tabs, bottom)
    End Sub

    Private Sub AddTrackColumn(header As String, width As Integer, getter As Func(Of TrackRow, String))
        _tracks.Columns.Add(New GridColumn With {.HeaderText = header, .Width = width, .Resizable = True, .DataCell = New TextBoxCell With {.Binding = Binding.Delegate(Of TrackRow, String)(getter)}})
    End Sub

    Private Sub Fill()
        Dim a As DmkAnalysis = _analysis

        ' ---- Summary: the facts first, then the findings by section
        Dim rows As New List(Of SummaryRow)
        rows.Add(New SummaryRow With {.Item = "Image and identification", .IsHeading = True})
        For Each kv In a.Facts
            rows.Add(New SummaryRow With {.Item = "   " & kv.Key, .Value = kv.Value})
        Next
        For Each section In New String() {"Image", "Geometry", "DOS identification", "Boot layout", "Protection", "Damage", "Tail tracks"}
            Dim wanted As String = section
            Dim items As List(Of DmkFinding) = a.Findings.Where(Function(f) f.Section = wanted).ToList()
            If items.Count = 0 Then Continue For
            rows.Add(New SummaryRow With {.Item = section, .IsHeading = True})
            For Each f In items
                Dim row As New SummaryRow With {.Value = Ui.Lf(f.Text).Replace(vbLf, " | ").Replace("    ", " ")}
                Select Case f.Level
                    Case "protection" : row.Item = "   Protection" : row.HasColour = True : row.Colour = Colors.DarkRed
                    Case "sign" : row.Item = "   Sign" : row.HasColour = True : row.Colour = Colors.DarkOrange
                    Case "damage" : row.Item = "   Damage" : row.HasColour = True : row.Colour = Colors.Firebrick
                    Case "artifact" : row.Item = "   Imaging" : row.HasColour = True : row.Colour = Colors.DarkSlateBlue
                    Case "warning" : row.Item = "   Note" : row.HasColour = True : row.Colour = Colors.DarkGoldenrod
                End Select
                rows.Add(row)
            Next
        Next
        _summary.DataStore = rows

        ' ---- Tracks
        Dim trackRows As New List(Of TrackRow)
        For Each t In a.Image.Tracks
            Dim tr As DmkTrack = t
            Dim notes As New List(Of String)
            If Not tr.Present Then
                notes.Add("missing from the file")
            Else
                If tr.BadPointers > 0 Then notes.Add(tr.BadPointers.ToString(CultureInfo.InvariantCulture) & " bad pointer(s)")
                Dim wrong As Integer = tr.Good.Where(Function(s) s.IdCylinder <> tr.Cylinder).Count()
                If wrong > 0 Then notes.Add(wrong.ToString(CultureInfo.InvariantCulture) & " ID(s) name another cylinder")
                If tr.Good.Any(Function(s) s.SizeCode > 3) Then notes.Add("size code > 3")
                If tr.Cylinder >= 1 AndAlso tr.HasFm AndAlso tr.HasMfm Then notes.Add("mixed density")
                If a.DosCylinders > 0 AndAlso tr.Cylinder >= a.DosCylinders Then notes.Add("beyond the DOS's " & a.DosCylinders.ToString(CultureInfo.InvariantCulture) & " cylinders")
                If a.Dos IsNot Nothing AndAlso tr.Cylinder = a.Dos.DirectoryCylinder AndAlso tr.Side = a.Dos.DirectorySide Then notes.Add("directory")
            End If
            Dim badCount As Integer = tr.Sectors.Where(Function(s) s.IsBad).Count()
            Dim row As New TrackRow With {
                .Cyl = tr.Cylinder.ToString(CultureInfo.InvariantCulture),
                .Side = tr.Side.ToString(CultureInfo.InvariantCulture),
                .Ids = tr.Sectors.Count.ToString(CultureInfo.InvariantCulture),
                .Layout = StripBad(tr.Summary()),
                .Bad = If(badCount = 0, "", badCount.ToString(CultureInfo.InvariantCulture)),
                .Notes = String.Join("; ", notes)}
            If badCount > 0 Then row.HasColour = True : row.Colour = Colors.Firebrick
            If Not tr.Present OrElse tr.Unformatted Then row.HasColour = True : row.Colour = Ui.GrayText
            trackRows.Add(row)
        Next
        _tracks.DataStore = trackRows

        _txtReport.Text = Ui.Lf(a.ReportText())
    End Sub

    ''' <summary>The track summary without its trailing "N bad" part (the Bad column shows that).</summary>
    Private Shared Function StripBad(s As String) As String
        Dim i As Integer = s.IndexOf("  ", StringComparison.Ordinal)
        Return If(i > 0, s.Substring(0, i), s)
    End Function

    Private Sub OnSave()
        If _analysis Is Nothing Then Return
        Dim target As String = Ui.PickSave(Me, "Save the DMK analysis report", "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                                           Path.Combine(Path.GetDirectoryName(_path), Path.GetFileNameWithoutExtension(_path) & " - DMK analysis.txt"), "txt")
        If target Is Nothing Then Return
        Try
            File.WriteAllText(target, Ui.HostLines(_txtReport.Text), New UTF8Encoding(True))
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, "Save report")
        End Try
    End Sub
End Class
