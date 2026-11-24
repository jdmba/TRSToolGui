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
        Icon = Program.WindowIcon
        _path = imagePath
        BuildUi()
    End Sub

    ''' <summary>Runs the analysis off the UI thread (a big DMK is a few MB) and fills the tabs.</summary>
    Public Async Function LoadAsync() As Threading.Tasks.Task
        _status.Text = UiText.T("dlg.dmk.reading", Path.GetFileName(_path))
        UseWaitCursor = True
        Try
            Dim p As String = _path
            _analysis = Await Threading.Tasks.Task.Run(Function() DmkAnalysis.Analyze(p))
            Fill()
            _status.Text = If(_analysis.Dos IsNot Nothing, _analysis.Dos.Result, _analysis.Image.Kind)
        Catch ex As Exception
            _status.Text = UiText.T("dlg.dmk.failed", ex.Message)
            _txtReport.Text = ex.ToString()
        Finally
            UseWaitCursor = False
        End Try
    End Function

    Private Sub BuildUi()
        Text = UiText.T("dlg.dmk.title", Path.GetFileName(_path))
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(960, 640)
        MinimumSize = New Size(640, 400)
        ShowInTaskbar = False

        _tabs.Dock = DockStyle.Fill
        ' Summary
        Dim tpSummary As New TabPage(UiText.T("dlg.dmk.tab.summary"))
        With _lvSummary
            .Dock = DockStyle.Fill
            .View = View.Details
            .FullRowSelect = True
            .HeaderStyle = ColumnHeaderStyle.Nonclickable
            .ShowGroups = True
            .Font = New Font(HostPlatform.UiFontName, 9.0F)
            .Columns.Add(UiText.T("dlg.dmk.col.item"), 190)
            .Columns.Add(UiText.T("dlg.dmk.col.value"), 720)
        End With
        UiFit.FitListColumns(_lvSummary)
        tpSummary.Controls.Add(_lvSummary)
        ' Tracks
        Dim tpTracks As New TabPage(UiText.T("dlg.dmk.tab.tracks"))
        With _grid
            .Dock = DockStyle.Fill
            .ReadOnly = True
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
            .Columns.Add("cyl", UiText.T("dlg.dmk.col.cyl")) : .Columns(0).Width = 40
            .Columns.Add("side", UiText.T("dlg.dmk.col.side")) : .Columns(1).Width = 40
            .Columns.Add("ids", UiText.T("dlg.dmk.col.ids")) : .Columns(2).Width = 40
            .Columns.Add("layout", UiText.T("dlg.dmk.col.layout")) : .Columns(3).Width = 480
            .Columns.Add("bad", UiText.T("dlg.dmk.col.bad")) : .Columns(4).Width = 40
            .Columns.Add("notes", UiText.T("dlg.dmk.col.notes")) : .Columns(5).Width = 300
        End With
        ' A heading is never cut: French's "Défauts" does not fit the 40 pixels "Bad" wants.
        UiFit.FitColumns(_grid)
        tpTracks.Controls.Add(_grid)
        ' Report
        Dim tpReport As New TabPage(UiText.T("dlg.dmk.tab.report"))
        With _txtReport
            .Dock = DockStyle.Fill
            .Multiline = True
            .ReadOnly = True
            .ScrollBars = ScrollBars.Both
            .WordWrap = False
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F)
        End With
        tpReport.Controls.Add(_txtReport)
        _tabs.TabPages.AddRange(New TabPage() {tpSummary, tpTracks, tpReport})

        Dim bottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 38}
        _status.AutoSize = False
        _status.Location = New Point(10, 11)
        _status.Size = New Size(560, 18)
        _status.AutoEllipsis = True
        _status.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Dim btnCopy As New Button() With {.Text = UiText.T("btn.copyreport"), .Size = New Size(100, 26), .Top = 6, .Anchor = AnchorStyles.Right Or AnchorStyles.Top}
        Dim btnSave As New Button() With {.Text = UiText.T("btn.savereport"), .Size = New Size(100, 26), .Top = 6, .Anchor = AnchorStyles.Right Or AnchorStyles.Top}
        Dim btnClose As New Button() With {.Text = UiText.T("btn.close"), .Size = New Size(80, 26), .Top = 6, .Anchor = AnchorStyles.Right Or AnchorStyles.Top, .DialogResult = DialogResult.Cancel}
        ' Each button is as wide as its own text needs (UiFit.vb): the French of "Save report..."
        ' is "Enregistrer le rapport...", which does not fit the 100 pixels English wants.
        UiFit.FitButton(btnCopy) : UiFit.FitButton(btnSave) : UiFit.FitButton(btnClose)
        AddHandler btnCopy.Click, Sub()
                                      If _txtReport.Text.Length > 0 Then Clipboard.SetText(_txtReport.Text)
                                  End Sub
        AddHandler btnSave.Click, AddressOf OnSave
        AddHandler bottom.Resize, Sub()
                                      ' Laid out from the right by their real widths, so a longer
                                      ' language cannot push one button under another.
                                      UiFit.ButtonRow(bottom.Width - 10, 10, btnCopy, btnSave, btnClose)
                                      _status.Width = Math.Max(100, btnCopy.Left - _status.Left - 10)
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
        Dim gFacts As New ListViewGroup(UiText.T("dlg.dmk.facts"))
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
                    Case "protection" : label = UiText.T("dlg.dmk.level.protection") : colour = Color.DarkRed
                    Case "sign" : label = UiText.T("dlg.dmk.level.sign") : colour = Color.DarkOrange
                    Case "damage" : label = UiText.T("dlg.dmk.level.damage") : colour = Color.Firebrick
                    Case "artifact" : label = UiText.T("dlg.dmk.level.artifact") : colour = Color.DarkSlateBlue
                    Case "warning" : label = UiText.T("dlg.dmk.level.note") : colour = Color.DarkGoldenrod
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
                notes.Add(UiText.T("dlg.dmk.note.missing"))
            Else
                If t.BadPointers > 0 Then notes.Add(UiText.T("dlg.dmk.note.badpointers", t.BadPointers.ToString(CultureInfo.InvariantCulture)))
                Dim wrong As Integer = t.Good.Where(Function(s) s.IdCylinder <> t.Cylinder).Count()
                If wrong > 0 Then notes.Add(UiText.T("dlg.dmk.note.wrongids", wrong.ToString(CultureInfo.InvariantCulture)))
                If t.Good.Any(Function(s) s.SizeCode > 3) Then notes.Add(UiText.T("dlg.dmk.note.sizecode"))
                If t.Cylinder >= 1 AndAlso t.HasFm AndAlso t.HasMfm Then notes.Add(UiText.T("dlg.dmk.note.mixed"))
                If a.DosCylinders > 0 AndAlso t.Cylinder >= a.DosCylinders Then notes.Add(UiText.T("dlg.dmk.note.beyonddos", a.DosCylinders.ToString(CultureInfo.InvariantCulture)))
                If a.Dos IsNot Nothing AndAlso t.Cylinder = a.Dos.DirectoryCylinder AndAlso t.Side = a.Dos.DirectorySide Then notes.Add(UiText.T("dlg.dmk.note.directory"))
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
            dlg.Title = UiText.T("dlg.dmk.savetitle")
            dlg.Filter = UiText.T("filter.text")
            dlg.FileName = UiText.T("dlg.dmk.savename", Path.GetFileNameWithoutExtension(_path))
            dlg.InitialDirectory = Path.GetDirectoryName(_path)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                Try
                    File.WriteAllText(dlg.FileName, _txtReport.Text)
                Catch ex As Exception
                    MessageBox.Show(Me, ex.Message, UiText.T("dlg.dmk.savecaption"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub
End Class
