' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskInfoForm.vb — "DMK analysis": what the GUI itself reads out of a DMK
'  image (DmkAnalyzer.vb + DosIdentifier.vb), beside what trs80-tool reports.
'  Tabs: Summary (label/value rows and the findings), Tracks (one row per
'  track side), Report (plain text; Copy / Save). Opened from Disk → DMK
'  analysis… (Ctrl+D) and from the Disk properties dialog. Read-only.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Public Class DiskInfoForm
    Inherits Form

    Private ReadOnly _path As String
    Private _analysis As DmkAnalysis = Nothing
    Private ReadOnly _tabs As New TabControl()
    Private ReadOnly _lvSummary As New ListView()
    Private ReadOnly _grid As New DataGridView()
    Private ReadOnly _txtReport As New TextBox()
    Private ReadOnly _status As New Label()

    Public Sub New(imagePath As String)
        _path = imagePath
        BuildUi()
    End Sub

    ''' <summary>Runs the analysis off the UI thread (a big DMK is a few MB) and fills the tabs.</summary>
    Public Async Function LoadAsync() As Threading.Tasks.Task
        _status.Text = "Reading " & Path.GetFileName(_path) & "..."
        UseWaitCursor = True
        Try
            Dim p As String = _path
            _analysis = Await Threading.Tasks.Task.Run(Function() DmkAnalysis.Analyze(p))
            Fill()
            _status.Text = If(_analysis.Dos IsNot Nothing, _analysis.Dos.Result, _analysis.Image.Kind)
        Catch ex As Exception
            _status.Text = "Could not analyse the image: " & ex.Message
            _txtReport.Text = ex.ToString()
        Finally
            UseWaitCursor = False
        End Try
    End Function

    Private Sub BuildUi()
        Text = "DMK analysis - " & Path.GetFileName(_path)
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(960, 640)
        MinimumSize = New Size(640, 400)
        ShowInTaskbar = False

        _tabs.Dock = DockStyle.Fill
        ' Summary
        Dim tpSummary As New TabPage("Summary")
        With _lvSummary
            .Dock = DockStyle.Fill
            .View = View.Details
            .FullRowSelect = True
            .HeaderStyle = ColumnHeaderStyle.Nonclickable
            .ShowGroups = True
            .Font = New Font("Segoe UI", 9.0F)
            .Columns.Add("Item", 190)
            .Columns.Add("Value", 720)
        End With
        tpSummary.Controls.Add(_lvSummary)
        ' Tracks
        Dim tpTracks As New TabPage("Tracks")
        With _grid
            .Dock = DockStyle.Fill
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .Font = New Font("Consolas", 9.0F)
            .Columns.Add("cyl", "Cyl") : .Columns(0).Width = 40
            .Columns.Add("side", "Side") : .Columns(1).Width = 40
            .Columns.Add("ids", "IDs") : .Columns(2).Width = 40
            .Columns.Add("layout", "Layout (density × count × size, numbering, marks)") : .Columns(3).Width = 480
            .Columns.Add("bad", "Bad") : .Columns(4).Width = 40
            .Columns.Add("notes", "Notes") : .Columns(5).Width = 300
        End With
        tpTracks.Controls.Add(_grid)
        ' Report
        Dim tpReport As New TabPage("Report")
        With _txtReport
            .Dock = DockStyle.Fill
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font("Consolas", 9.0F)
        End With
        tpReport.Controls.Add(_txtReport)
        _tabs.TabPages.AddRange(New TabPage() {tpSummary, tpTracks, tpReport})

        Dim bottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 38}
        _status.AutoSize = False
        _status.Location = New Point(10, 11)
        _status.Size = New Size(560, 18)
        _status.AutoEllipsis = True
        _status.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Dim btnCopy As New Button() With {.Text = "Copy report", .Size = New Size(100, 26), .Anchor = AnchorStyles.Right Or AnchorStyles.Top}
        Dim btnSave As New Button() With {.Text = "Save report...", .Size = New Size(100, 26), .Anchor = AnchorStyles.Right Or AnchorStyles.Top}
        Dim btnClose As New Button() With {.Text = "Close", .Size = New Size(80, 26), .Anchor = AnchorStyles.Right Or AnchorStyles.Top, .DialogResult = DialogResult.Cancel}
        AddHandler btnCopy.Click, Sub()
                                      If _txtReport.Text.Length > 0 Then Clipboard.SetText(_txtReport.Text)
                                  End Sub
        AddHandler btnSave.Click, AddressOf OnSave
        AddHandler bottom.Resize, Sub()
                                      btnClose.Location = New Point(bottom.Width - 90, 6)
                                      btnSave.Location = New Point(bottom.Width - 200, 6)
                                      btnCopy.Location = New Point(bottom.Width - 310, 6)
                                      _status.Width = Math.Max(100, bottom.Width - 330)
                                  End Sub
        bottom.Controls.AddRange(New Control() {_status, btnCopy, btnSave, btnClose})
        Controls.Add(_tabs)
        Controls.Add(bottom)
        CancelButton = btnClose
    End Sub

    Private Sub Fill()
        Dim a As DmkAnalysis = _analysis
        _lvSummary.BeginUpdate()
        _lvSummary.Items.Clear()
        _lvSummary.Groups.Clear()
        Dim gFacts As New ListViewGroup("Image and identification")
        _lvSummary.Groups.Add(gFacts)
        For Each kv In a.Facts
            Dim it As New ListViewItem(kv.Key) With {.Group = gFacts}
            it.SubItems.Add(kv.Value)
            _lvSummary.Items.Add(it)
        Next
        For Each section In New String() {"Image", "Geometry", "DOS identification", "Boot layout", "Protection", "Damage", "Tail tracks"}
            Dim items As List(Of DmkFinding) = a.Findings.Where(Function(f) f.Section = section).ToList()
            If items.Count = 0 Then Continue For
            Dim g As New ListViewGroup(section)
            _lvSummary.Groups.Add(g)
            For Each f In items
                Dim label As String
                Dim colour As Color = SystemColors.WindowText
                Select Case f.Level
                    Case "protection" : label = "Protection" : colour = Color.DarkRed
                    Case "sign" : label = "Sign" : colour = Color.DarkOrange
                    Case "damage" : label = "Damage" : colour = Color.Firebrick
                    Case "artifact" : label = "Imaging" : colour = Color.DarkSlateBlue
                    Case "warning" : label = "Note" : colour = Color.DarkGoldenrod
                    Case Else : label = ""
                End Select
                Dim it As New ListViewItem(label) With {.Group = g, .ForeColor = colour, .ToolTipText = f.Text}
                it.SubItems.Add(f.Text.Replace(vbCrLf, " | ").Replace("    ", " "))
                _lvSummary.Items.Add(it)
            Next
        Next
        _lvSummary.ShowItemToolTips = True
        _lvSummary.EndUpdate()

        _grid.Rows.Clear()
        For Each t In a.Image.Tracks
            Dim notes As New List(Of String)
            If Not t.Present Then
                notes.Add("missing from the file")
            Else
                If t.BadPointers > 0 Then notes.Add(t.BadPointers.ToString(CultureInfo.InvariantCulture) & " bad pointer(s)")
                Dim wrong As Integer = t.Good.Where(Function(s) s.IdCylinder <> t.Cylinder).Count()
                If wrong > 0 Then notes.Add(wrong.ToString(CultureInfo.InvariantCulture) & " ID(s) name another cylinder")
                If t.Good.Any(Function(s) s.SizeCode > 3) Then notes.Add("size code > 3")
                If t.Cylinder >= 1 AndAlso t.HasFm AndAlso t.HasMfm Then notes.Add("mixed density")
                If a.DosCylinders > 0 AndAlso t.Cylinder >= a.DosCylinders Then notes.Add("beyond the DOS's " & a.DosCylinders.ToString(CultureInfo.InvariantCulture) & " cylinders")
                If a.Dos IsNot Nothing AndAlso t.Cylinder = a.Dos.DirectoryCylinder AndAlso t.Side = a.Dos.DirectorySide Then notes.Add("directory")
            End If
            Dim badCount As Integer = t.Sectors.Where(Function(s) s.IsBad).Count()
            Dim row As Integer = _grid.Rows.Add(t.Cylinder, t.Side, t.Sectors.Count, StripBad(t.Summary()), If(badCount = 0, "", badCount.ToString(CultureInfo.InvariantCulture)), String.Join("; ", notes))
            If badCount > 0 Then _grid.Rows(row).DefaultCellStyle.ForeColor = Color.Firebrick
            If Not t.Present OrElse t.Unformatted Then _grid.Rows(row).DefaultCellStyle.ForeColor = SystemColors.GrayText
        Next
        _txtReport.Text = a.ReportText().Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)
    End Sub

    Private Shared Function StripBad(s As String) As String
        Dim i As Integer = s.IndexOf("  ", StringComparison.Ordinal)
        Return If(i > 0, s.Substring(0, i), s)
    End Function

    Private Sub OnSave(sender As Object, e As EventArgs)
        If _analysis Is Nothing Then Return
        Using dlg As New SaveFileDialog()
            dlg.Title = "Save the DMK analysis report"
            dlg.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            dlg.FileName = Path.GetFileNameWithoutExtension(_path) & " - DMK analysis.txt"
            dlg.InitialDirectory = Path.GetDirectoryName(_path)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                Try
                    File.WriteAllText(dlg.FileName, _txtReport.Text)
                Catch ex As Exception
                    MessageBox.Show(Me, ex.Message, "Save report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub
End Class
