' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  MainForm.vb — the TRSTools-style main window: menu, toolbar, file list,
'  status bar, context menu, drag & drop, clipboard. Every disk operation is
'  delegated to Trs80Tool (the engine adapter); this form only decides *what*
'  to ask for and shows the result. The Disk and Tools menus hold the
'  trs80-tool features TRSTool never had.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class MainForm
    Inherits Form

    Private ReadOnly _menu As New MenuStrip()
    Private ReadOnly _toolbar As New ToolStrip()
    Private ReadOnly _status As New StatusStrip()
    Private ReadOnly _lblStatus As New ToolStripStatusLabel()
    Private ReadOnly _lblEngine As New ToolStripStatusLabel()
    Private ReadOnly _lv As New ListView()
    Private ReadOnly _ctx As New ContextMenuStrip()
    Private ReadOnly _sorter As New ListViewColumnSorter()

    ' menu items that need enabling/disabling
    Private _miClose, _miDiskProps, _miCopy, _miPaste, _miAdd, _miExtract, _miSelectAll, _miInvert As ToolStripMenuItem
    Private _miInvisible, _miSystem, _miPartition, _miRefresh As ToolStripMenuItem
    Private _miRecent, _miSaveAs, _miNewHardDisk As ToolStripMenuItem
    Private _miSectorMap, _miGranuleMap, _miCheckRepair, _miFormat, _miTrim, _miBoot, _miShell, _miDmkInfo, _miRunEmpty, _miCpmFormat As ToolStripMenuItem
    Private _miImport As ToolStripMenuItem
    Private _cxPreview, _cxExtract, _cxDelete, _cxRename, _cxProps, _cxConvert As ToolStripMenuItem
    Private _tbProps, _tbCopy, _tbPaste, _tbAdd, _tbExtract, _tbView, _tbSectors As ToolStripButton

    ''' <summary>A trs80-tool console (disk shell or emulator) that may be writing the open image; mutations wait for it.</summary>
    Private _console As ConsoleRun = Nothing
    Private _consoleWhat As String = ""

    Private _disk As DiskImage = Nothing
    Private _busy As Boolean = False
    Private ReadOnly _startFile As String

    Public Sub New(startFile As String)
        _startFile = startFile
        BuildUi()
    End Sub

    ' =================================================================== UI construction

    Private Sub BuildUi()
        Text = Program.AppName
        StartPosition = FormStartPosition.Manual
        Dim s As Settings = Program.Settings
        Size = New Size(Math.Max(480, s.WindowWidth), Math.Max(300, s.WindowHeight))
        If s.WindowX >= 0 AndAlso s.WindowY >= 0 AndAlso Screen.AllScreens.Any(Function(sc) sc.WorkingArea.Contains(New Point(s.WindowX + 20, s.WindowY + 20))) Then
            Location = New Point(s.WindowX, s.WindowY)
        Else
            StartPosition = FormStartPosition.WindowsDefaultLocation
        End If
        If s.WindowMaximized Then WindowState = FormWindowState.Maximized
        KeyPreview = True
        AllowDrop = True
        Icon = Program.WindowIcon

        ' ---- menu
        Dim mFile As New ToolStripMenuItem(UiText.T("menu.file"))
        mFile.DropDownItems.Add(MakeItem(UiText.T("menu.file.new"), Keys.Control Or Keys.N, AddressOf OnFileNew))
        _miNewHardDisk = MakeItem(UiText.T("menu.file.newhd"), Keys.None, AddressOf OnFileNewHardDisk)
        mFile.DropDownItems.Add(_miNewHardDisk)
        mFile.DropDownItems.Add(MakeItem(UiText.T("menu.file.open"), Keys.Control Or Keys.O, AddressOf OnFileOpen))
        _miRecent = New ToolStripMenuItem(UiText.T("menu.file.recent"))
        mFile.DropDownItems.Add(_miRecent)
        _miClose = MakeItem(UiText.T("menu.file.close"), Keys.None, AddressOf OnFileClose)
        mFile.DropDownItems.Add(_miClose)
        mFile.DropDownItems.Add(New ToolStripSeparator())
        _miDiskProps = MakeItem(UiText.T("menu.file.diskprops"), Keys.None, AddressOf OnDiskProperties)
        mFile.DropDownItems.Add(_miDiskProps)
        _miSaveAs = MakeItem(UiText.T("menu.file.saveas"), Keys.None, AddressOf OnSaveImageAs)
        mFile.DropDownItems.Add(_miSaveAs)
        mFile.DropDownItems.Add(New ToolStripSeparator())
        mFile.DropDownItems.Add(MakeItem(UiText.T("menu.file.exit"), Keys.None, Sub() Close()))

        Dim mEdit As New ToolStripMenuItem(UiText.T("menu.edit"))
        _miCopy = MakeItem(UiText.T("menu.edit.copy"), Keys.Control Or Keys.C, AddressOf OnEditCopy)
        _miPaste = MakeItem(UiText.T("menu.edit.paste"), Keys.Control Or Keys.V, AddressOf OnEditPaste)
        _miAdd = MakeItem(UiText.T("menu.edit.add"), Keys.None, AddressOf OnEditAdd)
        _miExtract = MakeItem(UiText.T("menu.edit.extract"), Keys.None, AddressOf OnEditExtract)
        _miSelectAll = MakeItem(UiText.T("menu.edit.selectall"), Keys.Control Or Keys.A, Sub() SelectAll())
        _miInvert = MakeItem(UiText.T("menu.edit.invert"), Keys.None, Sub() InvertSelection())
        mEdit.DropDownItems.AddRange(New ToolStripItem() {_miCopy, _miPaste, New ToolStripSeparator(), _miAdd, _miExtract, New ToolStripSeparator(), _miSelectAll, _miInvert, New ToolStripSeparator(), MakeItem(UiText.T("menu.edit.prefs"), Keys.None, AddressOf OnPreferences)})

        Dim mView As New ToolStripMenuItem(UiText.T("menu.view"))
        _miInvisible = MakeItem(UiText.T("menu.view.invisible"), Keys.None, AddressOf OnViewInvisible)
        _miInvisible.CheckOnClick = False
        _miSystem = MakeItem(UiText.T("menu.view.system"), Keys.None, AddressOf OnViewSystem)
        _miPartition = New ToolStripMenuItem(UiText.T("menu.view.partition"))
        _miRefresh = MakeItem(UiText.T("menu.view.refresh"), Keys.F5, AddressOf OnRefresh)
        mView.DropDownItems.AddRange(New ToolStripItem() {_miInvisible, _miSystem, _miPartition, New ToolStripSeparator(), _miRefresh})

        ' ---- Disk and Tools menus: trs80-tool features beyond TRSTool
        Dim mDisk As New ToolStripMenuItem(UiText.T("menu.disk"))
        _miSectorMap = MakeItem(UiText.T("menu.disk.sectormap"), Keys.Control Or Keys.M, AddressOf OnSectorMap)
        _miDmkInfo = MakeItem(UiText.T("menu.disk.dmk"), Keys.Control Or Keys.D, AddressOf OnDmkInfo)
        _miCpmFormat = MakeItem(UiText.T("menu.disk.cpm"), Keys.None, AddressOf OnCpmFormat)
        _miGranuleMap = MakeItem(UiText.T("menu.disk.granule"), Keys.None, AddressOf OnGranuleMap)
        _miCheckRepair = MakeItem(UiText.T("menu.disk.checkrepair"), Keys.None, AddressOf OnCheckRepair)
        _miFormat = MakeItem(UiText.T("menu.disk.format"), Keys.None, AddressOf OnFormat)
        _miTrim = MakeItem(UiText.T("menu.disk.trim"), Keys.None, AddressOf OnTrim)
        _miBoot = MakeItem(UiText.T("menu.disk.boot"), Keys.None, AddressOf OnBoot)
        _miShell = MakeItem(UiText.T("menu.disk.shell"), Keys.None, AddressOf OnShell)
        _miImport = MakeItem(UiText.T("menu.disk.import"), Keys.None, AddressOf OnImport)
        mDisk.DropDownItems.AddRange(New ToolStripItem() {_miSectorMap, _miDmkInfo, _miGranuleMap, _miCheckRepair, _miCpmFormat, New ToolStripSeparator(), _miImport, _miFormat, _miTrim, New ToolStripSeparator(), _miBoot, _miShell})

        Dim mTools As New ToolStripMenuItem(UiText.T("menu.tools"))
        mTools.DropDownItems.Add(MakeItem(UiText.T("menu.tools.convert"), Keys.None, Sub() OnConvertFiles(Nothing)))
        mTools.DropDownItems.Add(MakeItem(UiText.T("menu.tools.assemble"), Keys.None, AddressOf OnAssemble))
        mTools.DropDownItems.Add(New ToolStripSeparator())
        mTools.DropDownItems.Add(MakeItem(UiText.T("menu.tools.hexdump"), Keys.None, AddressOf OnHexdumpFile))
        mTools.DropDownItems.Add(MakeItem(UiText.T("menu.tools.disasm"), Keys.None, AddressOf OnDisasmFile))
        mTools.DropDownItems.Add(New ToolStripSeparator())
        _miRunEmpty = MakeItem(UiText.T("menu.tools.runempty"), Keys.None, AddressOf OnRunEmulatorEmpty)
        mTools.DropDownItems.Add(_miRunEmpty)

        Dim mHelp As New ToolStripMenuItem(UiText.T("menu.help"))
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.manual"), Keys.F1, Sub() ShowHelpText("Manual.txt", UiText.T("win.manual", Program.AppName))))
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.changelog"), Keys.None, Sub() ShowHelpText("Changelog.txt", UiText.T("win.changelog", Program.AppName))))
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.license"), Keys.None, Sub() ShowHelpText("License.txt", UiText.T("win.license", Program.AppName))))
        mHelp.DropDownItems.Add(New ToolStripSeparator())
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.engine"), Keys.None, AddressOf OnEngineInfo))
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.toolweb"), Keys.None, Sub() OpenUrl("https://www.my-trs-80.com/tool/")))
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.report"), Keys.None, Sub() OpenUrl(AppInfo.IssuesUrl)))
        mHelp.DropDownItems.Add(New ToolStripSeparator())
        mHelp.DropDownItems.Add(MakeItem(UiText.T("menu.help.about", Program.AppName), Keys.None, AddressOf OnAbout))

        _menu.Items.AddRange(New ToolStripItem() {mFile, mEdit, mView, mDisk, mTools, mHelp})
        MainMenuStrip = _menu

        ' ---- toolbar (text buttons: no bitmaps to maintain)
        _toolbar.GripStyle = ToolStripGripStyle.Hidden
        _toolbar.Items.Add(MakeButton(UiText.T("tb.new"), AddressOf OnFileNew))
        _toolbar.Items.Add(MakeButton(UiText.T("tb.open"), AddressOf OnFileOpen))
        _tbProps = MakeButton(UiText.T("tb.diskprops"), AddressOf OnDiskProperties)
        _toolbar.Items.Add(_tbProps)
        _toolbar.Items.Add(New ToolStripSeparator())
        _tbCopy = MakeButton(UiText.T("tb.copy"), AddressOf OnEditCopy)
        _tbPaste = MakeButton(UiText.T("tb.paste"), AddressOf OnEditPaste)
        _toolbar.Items.Add(_tbCopy)
        _toolbar.Items.Add(_tbPaste)
        _toolbar.Items.Add(New ToolStripSeparator())
        _tbAdd = MakeButton(UiText.T("tb.add"), AddressOf OnEditAdd)
        _tbExtract = MakeButton(UiText.T("tb.extract"), AddressOf OnEditExtract)
        _tbView = MakeButton(UiText.T("tb.view"), AddressOf OnPreview)
        _toolbar.Items.Add(_tbAdd)
        _toolbar.Items.Add(_tbExtract)
        _toolbar.Items.Add(_tbView)
        _toolbar.Items.Add(New ToolStripSeparator())
        _tbSectors = MakeButton(UiText.T("tb.sectors"), AddressOf OnSectorMap)
        _toolbar.Items.Add(_tbSectors)
        _toolbar.Items.Add(New ToolStripSeparator())
        _toolbar.Items.Add(MakeButton(UiText.T("tb.about"), AddressOf OnAbout))

        ' ---- list view
        With _lv
            .Dock = DockStyle.Fill
            .View = View.Details
            .FullRowSelect = True
            .MultiSelect = True
            .HideSelection = False
            .AllowDrop = True
            .ShowItemToolTips = True
            .Sorting = SortOrder.None
            .ListViewItemSorter = _sorter
            .Columns.Add(UiText.T("col.name"), 110)
            .Columns.Add(UiText.T("col.type"), 55)
            .Columns.Add(UiText.T("col.size"), 80, HorizontalAlignment.Right)
            .Columns.Add(UiText.T("col.date"), 130)
            .Columns.Add(UiText.T("col.attributes"), 110)
            .Columns.Add(UiText.T("col.kind"), 200)
        End With
        ' A heading is never cut (UiFit.vb); a width the user has set himself still wins, below.
        UiFit.FitListColumns(_lv)
        If s.ColumnWidths IsNot Nothing AndAlso s.ColumnWidths.Count = _lv.Columns.Count Then
            For i As Integer = 0 To _lv.Columns.Count - 1
                If s.ColumnWidths(i) > 10 Then _lv.Columns(i).Width = s.ColumnWidths(i)
            Next
        End If
        AddHandler _lv.ColumnClick, AddressOf OnColumnClick
        AddHandler _lv.DoubleClick, Sub() OnPreview()
        AddHandler _lv.KeyDown, AddressOf OnListKeyDown
        AddHandler _lv.ItemDrag, AddressOf OnItemDrag
        AddHandler _lv.DragEnter, AddressOf OnDragEnterList
        AddHandler _lv.DragDrop, AddressOf OnDragDropList
        AddHandler _lv.SelectedIndexChanged, Sub() UpdateEnabled()
        AddHandler _lv.MouseUp, AddressOf OnListMouseUp
        AddHandler Me.DragEnter, AddressOf OnDragEnterList
        AddHandler Me.DragDrop, AddressOf OnDragDropList

        ' ---- context menu
        _cxPreview = MakeItem(UiText.T("cx.preview"), Keys.None, AddressOf OnPreview)
        _cxExtract = MakeItem(UiText.T("menu.edit.extract"), Keys.None, AddressOf OnEditExtract)
        _cxDelete = MakeItem(UiText.T("cx.delete"), Keys.None, AddressOf OnDelete)
        _cxRename = MakeItem(UiText.T("cx.rename"), Keys.None, AddressOf OnRename)
        _cxProps = MakeItem(UiText.T("cx.props"), Keys.None, AddressOf OnFileProperties)
        _cxConvert = MakeItem(UiText.T("cx.convert"), Keys.None, AddressOf OnConvertDiskFile)
        _ctx.Items.AddRange(New ToolStripItem() {_cxPreview, New ToolStripSeparator(), _cxExtract, _cxDelete, _cxRename, _cxConvert, New ToolStripSeparator(), _cxProps})

        ' ---- status bar
        _lblStatus.Spring = True
        _lblStatus.TextAlign = ContentAlignment.MiddleLeft
        _lblStatus.Text = UiText.T("status.ready")
        _lblEngine.Text = ""
        _status.Items.AddRange(New ToolStripItem() {_lblStatus, _lblEngine})

        ' WinForms docks the LAST child first (outermost), so the add order below puts the menu above
        ' the toolbar, the status bar at the bottom and the list in the remaining space. Keep this order.
        Controls.Add(_lv)
        Controls.Add(_toolbar)
        Controls.Add(_menu)
        Controls.Add(_status)
        _toolbar.Dock = DockStyle.Top
        _menu.Dock = DockStyle.Top
        _status.Dock = DockStyle.Bottom

        UpdateViewChecks()
        RebuildRecentMenu()
        UpdateEnabled()
    End Sub

    Private Shared Function MakeItem(text As String, shortcut As Keys, handler As Action) As ToolStripMenuItem
        Dim mi As New ToolStripMenuItem(text)
        If shortcut <> Keys.None Then mi.ShortcutKeys = shortcut
        AddHandler mi.Click, Sub() handler()
        Return mi
    End Function

    Private Shared Function MakeButton(text As String, handler As Action) As ToolStripButton
        Dim b As New ToolStripButton(text) With {.DisplayStyle = ToolStripItemDisplayStyle.Text}
        AddHandler b.Click, Sub() handler()
        Return b
    End Function

    Private Shared Sub OpenUrl(url As String)
        Try
            Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
        Catch
        End Try
    End Sub

    ' =================================================================== startup / shutdown

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Await CheckEngineAsync(True)
        If _startFile IsNot Nothing Then Await OpenImageAsync(_startFile)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Dim s As Settings = Program.Settings
        s.WindowMaximized = (WindowState = FormWindowState.Maximized)
        Dim b As Rectangle = If(WindowState = FormWindowState.Normal, Bounds, RestoreBounds)
        s.WindowWidth = b.Width : s.WindowHeight = b.Height : s.WindowX = b.X : s.WindowY = b.Y
        s.ColumnWidths = New List(Of Integer)
        For Each c As ColumnHeader In _lv.Columns
            s.ColumnWidths.Add(c.Width)
        Next
        s.Save()
        MyBase.OnFormClosing(e)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.Escape AndAlso Program.Settings.ExitOnEscape AndAlso Not _busy Then
            Close()
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    ''' <summary>Run the engine probe; show problems. With askIfMissing, offer Preferences when the tool can't be found.</summary>
    Private Async Function CheckEngineAsync(askIfMissing As Boolean) As Task(Of Boolean)
        Dim tool As Trs80Tool = Program.Tool
        If String.IsNullOrEmpty(tool.ToolPath) OrElse Not File.Exists(tool.ToolPath) Then
            _lblEngine.Text = UiText.T("status.enginenotfound")
            If askIfMissing Then
                Dim r = MessageBox.Show(Me, UiText.T("msg.enginemissing.windows") & vbCrLf & vbCrLf &
                    UiText.T("msg.enginemissing.download.windows") & vbCrLf & vbCrLf &
                    UiText.T("msg.enginemissing.ask"), Program.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If r = DialogResult.Yes Then OnPreferences()
            End If
            Return False
        End If
        SetBusy(True, UiText.T("status.checkingengine"))
        Dim caps As EngineCapabilities
        Try
            caps = Await tool.ProbeAsync()
        Finally
            SetBusy(False, Nothing)
        End Try
        If caps.Version Is Nothing Then
            _lblEngine.Text = UiText.T("status.engineunusable")
            MessageBox.Show(Me, UiText.TCrLf("msg.engineunreadable", String.Join(vbCrLf, caps.Problems), tool.ToolPath), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        _lblEngine.Text = UiText.T("status.engineversion", caps.Version.ToString())
        If caps.Version < Trs80Tool.MinimumVersion Then
            MessageBox.Show(Me, UiText.TCrLf("msg.enginetooold", caps.Version.ToString(), Trs80Tool.MinimumVersion.ToString(), "https://www.my-trs-80.com/tool/", tool.ToolPath), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        If caps.Version > Trs80Tool.MaxTestedVersion AndAlso Program.Settings.LastNotedVersion <> caps.Version.ToString() Then
            Program.Settings.LastNotedVersion = caps.Version.ToString()
            Program.Settings.Save()
            Dim missing As New List(Of String)
            For Each p In caps.Problems
                If p.Contains(" missing") Then missing.Add(p)
            Next
            MessageBox.Show(Me, UiText.TCrLf("msg.enginenewer", caps.Version.ToString(), Trs80Tool.MaxTestedVersion.ToString()) & vbCrLf & vbCrLf &
                            If(missing.Count = 0, UiText.T("msg.enginenewer.ok"), UiText.TCrLf("msg.enginenewer.missing", String.Join(vbCrLf, missing))) & vbCrLf & vbCrLf &
                            UiText.T("msg.enginenewer.details"), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        Return True
    End Function

    ' =================================================================== state helpers

    Private Sub SetBusy(busy As Boolean, statusText As String)
        _busy = busy
        UseWaitCursor = busy
        _lv.Enabled = Not busy
        _menu.Enabled = Not busy
        _toolbar.Enabled = Not busy
        If statusText IsNot Nothing Then _lblStatus.Text = statusText
        If Not busy Then UpdateStatus()
        Application.DoEvents()
    End Sub

    Private Sub UpdateStatus()
        If _disk Is Nothing Then
            _lblStatus.Text = UiText.T("status.ready")
            Return
        End If
        Dim parts As New List(Of String)
        If _disk.DosName.Length > 0 Then parts.Add(_disk.DosName) Else parts.Add(_disk.Description)
        parts.Add(UiText.T("status.filesshown", _lv.Items.Count.ToString(CultureInfo.InvariantCulture), _disk.VisibleFiles(True, True).Count.ToString(CultureInfo.InvariantCulture)))
        If _disk.FreeText.Length > 0 Then parts.Add(_disk.FreeText)
        If _disk.WriteProtected.HasValue AndAlso _disk.WriteProtected.Value Then parts.Add(UiText.T("status.writeprotected"))
        If _disk.Damaged Then parts.Add(UiText.T("status.damaged"))
        If _disk.ProtectedSignCount > 0 Then parts.Add(UiText.T("status.protectionsigns"))
        If Not _disk.IsWritable Then parts.Add(UiText.T("status.readonly", If(_disk.ReadOnlyReason, "")))
        If Program.Tool.CpmFormatOf(_disk.Path).Length > 0 Then parts.Add(UiText.T("status.cpmlayout", Program.Tool.CpmFormatOf(_disk.Path)))
        If ConsoleRunning() Then parts.Add(UiText.T("status.consoleopen", _consoleWhat))
        _lblStatus.Text = String.Join("   |   ", parts)
    End Sub

    Private Sub UpdateTitle()
        If _disk Is Nothing Then
            Text = Program.AppName
            Return
        End If
        Dim inner As String = If(_disk.DosName.Length > 0, _disk.DosName, _disk.Description)
        If _disk.VolumeName.Length > 0 Then inner = UiText.T("win.main.volume", inner, _disk.VolumeName)
        If _disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing Then inner = UiText.T("title.partition", inner, _disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture))
        Dim t As String = UiText.T("win.main.open", Program.AppName, Path.GetFileName(_disk.Path), inner)
        Text = t
    End Sub

    Private Sub UpdateViewChecks()
        _miInvisible.Checked = Program.Settings.ShowInvisible
        _miSystem.Checked = Program.Settings.ShowSystem
    End Sub

    Private Sub UpdateEnabled()
        Dim isOpen As Boolean = _disk IsNot Nothing
        Dim writable As Boolean = isOpen AndAlso _disk.IsWritable AndAlso Not ConsoleRunning()
        Dim selCount As Integer = _lv.SelectedItems.Count
        Dim hasDisk As Boolean = EngineHasDisk("add")
        Dim floppy As Boolean = isOpen AndAlso Not _disk.IsHardDisk AndAlso Not _disk.IsCassette
        _miClose.Enabled = isOpen
        _miDiskProps.Enabled = isOpen : _tbProps.Enabled = isOpen
        _miCopy.Enabled = isOpen AndAlso selCount > 0 : _tbCopy.Enabled = _miCopy.Enabled
        _miPaste.Enabled = writable AndAlso hasDisk AndAlso Clipboard.ContainsFileDropList() : _tbPaste.Enabled = _miPaste.Enabled
        _miAdd.Enabled = writable AndAlso hasDisk : _tbAdd.Enabled = _miAdd.Enabled
        _miExtract.Enabled = isOpen AndAlso _lv.Items.Count > 0 : _tbExtract.Enabled = _miExtract.Enabled
        _miSelectAll.Enabled = isOpen
        _miInvert.Enabled = isOpen
        _miRefresh.Enabled = isOpen
        _miPartition.Enabled = isOpen AndAlso _disk.IsHardDisk
        _miNewHardDisk.Enabled = Program.Tool.Capabilities Is Nothing OrElse Program.Tool.Capabilities.CanCreate("HDV")
        _tbView.Enabled = selCount = 1
        _cxPreview.Enabled = selCount = 1
        _cxRename.Enabled = writable AndAlso selCount = 1
        _cxProps.Enabled = selCount = 1
        _cxConvert.Enabled = selCount = 1 AndAlso EngineHasTop("convert")
        _cxDelete.Enabled = writable AndAlso selCount > 0
        _cxExtract.Enabled = isOpen AndAlso _lv.Items.Count > 0
        _miSaveAs.Enabled = floppy AndAlso EngineHasTop("convert")
        _miSectorMap.Enabled = floppy AndAlso EngineHasTop("sectors") : _tbSectors.Enabled = _miSectorMap.Enabled
        _miDmkInfo.Enabled = floppy
        ' "--cpm-format" is trs80-tool 5.2.0's.
        _miCpmFormat.Enabled = floppy AndAlso Program.Tool.Has520 AndAlso Not ConsoleRunning()
        ' trs80-tool 5.2.0 also maps the open partition of a hard disk.
        _miGranuleMap.Enabled = (floppy OrElse (isOpen AndAlso _disk.IsHardDisk AndAlso Program.Tool.Has520)) AndAlso EngineHasDisk("map")
        _miCheckRepair.Enabled = isOpen AndAlso Not _disk.IsCassette AndAlso EngineHasDisk("check")
        _miFormat.Enabled = floppy AndAlso writable AndAlso EngineHasDisk("format")
        _miTrim.Enabled = floppy AndAlso Not ConsoleRunning() AndAlso EngineHasDisk("trim")
        ' "disk import" is trs80-tool 5.3.0's; CanImport tests the version as well as the probe, because
        ' up to 5.2.0 "import" was an alias of "add" (see Trs80Tool.CanImport).
        _miImport.Enabled = isOpen AndAlso writable AndAlso Not ConsoleRunning() AndAlso Program.Tool.CanImport
        _miBoot.Enabled = floppy AndAlso Not ConsoleRunning() AndAlso EngineHasTop("run")
        _miShell.Enabled = isOpen AndAlso Not _disk.IsCassette AndAlso Not ConsoleRunning()
        _miRunEmpty.Enabled = Not ConsoleRunning() AndAlso EngineHasTop("run")
    End Sub

    ''' <summary>True unless the engine probe ran and its disk --help lacks the sub-command.</summary>
    Private Shared Function EngineHasDisk(name As String) As Boolean
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        Return caps Is Nothing OrElse caps.DiskCommands.Count = 0 OrElse caps.HasDisk(name)
    End Function

    Private Shared Function EngineHasTop(name As String) As Boolean
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        Return caps Is Nothing OrElse caps.TopLevelCommands.Count = 0 OrElse caps.HasTopLevel(name)
    End Function

    Private Function ConsoleRunning() As Boolean
        If _console Is Nothing Then Return False
        Try
            If _console.HasExited Then
                _console.Dispose()
                _console = Nothing
                Return False
            End If
            Return True
        Catch
            _console = Nothing
            Return False
        End Try
    End Function

    Private Sub RebuildRecentMenu()
        _miRecent.DropDownItems.Clear()
        If Program.Settings.Mru.Count = 0 Then
            _miRecent.DropDownItems.Add(New ToolStripMenuItem(UiText.T("menu.file.recent.none")) With {.Enabled = False})
            Return
        End If
        Dim n As Integer = 1
        For Each p In Program.Settings.Mru
            Dim pathCopy As String = p
            Dim mi As New ToolStripMenuItem("&" & n.ToString(CultureInfo.InvariantCulture) & " " & p)
            AddHandler mi.Click, Async Sub() Await OpenImageAsync(pathCopy)
            _miRecent.DropDownItems.Add(mi)
            n += 1
        Next
    End Sub

    Private Sub RebuildPartitionMenu()
        _miPartition.DropDownItems.Clear()
        If _disk Is Nothing OrElse Not _disk.IsHardDisk Then Return
        For Each p In _disk.Partitions
            Dim part As DiskPartition = p
            Dim mi As New ToolStripMenuItem(part.Display) With {.Checked = (part Is _disk.ActivePartition)}
            AddHandler mi.Click, Async Sub() Await SelectPartitionAsync(part)
            _miPartition.DropDownItems.Add(mi)
        Next
        ' trs80-tool 5.1.0 makes and erases partitions ("disk partition --add / --delete").
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        If caps IsNot Nothing AndAlso caps.HasDiskOption("partition", "--add") Then
            If _disk.Partitions.Count > 0 Then _miPartition.DropDownItems.Add(New ToolStripSeparator())
            Dim miAdd As New ToolStripMenuItem(UiText.T("menu.view.partadd")) With {.Enabled = Not ConsoleRunning()}
            AddHandler miAdd.Click, Sub() OnPartitionAdd()
            _miPartition.DropDownItems.Add(miAdd)
            Dim miDel As New ToolStripMenuItem(UiText.T("menu.view.partdelete")) With {.Enabled = _disk.ActivePartition IsNot Nothing AndAlso Not ConsoleRunning()}
            AddHandler miDel.Click, Sub() OnPartitionDelete()
            _miPartition.DropDownItems.Add(miDel)
        End If
    End Sub

    ''' <summary>A window with text the tool printed that the user needs to keep (driver settings of a new hard disk, say).</summary>
    Private Sub ShowToolText(title As String, text As String)
        Dim body As String = If(text, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf).Trim()
        If body.Length = 0 Then Return
        Dim f As New TextReportForm(title, body)
        f.Show(Me)
    End Sub

    ''' <summary>File → New hard disk image: "disk NEW.hdv create [DOS] --cylinders N --heads N [--partitions N] ...".</summary>
    Private Async Sub OnFileNewHardDisk()
        If _busy Then Return
        Using dlg As New CreateHardDiskForm()
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Using sfd As New SaveFileDialog()
                sfd.Title = UiText.T("caption.savenewhd")
                sfd.DefaultExt = "hdv"
                sfd.Filter = UiText.T("filter.hdv")
                sfd.OverwritePrompt = True
                sfd.AddExtension = True
                If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
                Dim target As String = sfd.FileName
                If Not target.EndsWith(".hdv", StringComparison.OrdinalIgnoreCase) Then target &= ".hdv"
                SetBusy(True, UiText.T("status.creating", Path.GetFileName(target)))
                Dim r As ToolResult = Nothing
                Dim failure As String = Nothing
                Try
                    If File.Exists(target) Then File.Delete(target)
                    r = Await Program.Tool.DiskCreateAsync(target, dlg.DosName, dlg.OptionWords())
                Catch ex As Exception
                    failure = ex.Message
                End Try
                SetBusy(False, Nothing)
                If failure IsNot Nothing Then
                    MessageBox.Show(Me, failure, UiText.T("caption.diskcreateerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                Await OpenImageAsync(target)
                ' An HDV file has no partition table: the tool prints what the DOS's hard disk driver must be told.
                ShowToolText(UiText.T("win.newhd", Path.GetFileName(target)), r.StdOut)
            End Using
        End Using
    End Sub

    ''' <summary>View → Partition → Add a partition: "disk IMG partition --add DOS ...".</summary>
    Private Async Sub OnPartitionAdd()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsHardDisk OrElse ConsoleRunning() Then Return
        Using dlg As New PartitionAddForm()
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            SetBusy(True, UiText.T("status.addingpartition"))
            Dim r As ToolResult = Nothing
            Dim failure As String = Nothing
            Try
                r = Await Program.Tool.DiskPartitionAddAsync(_disk.Path, dlg.DosName, dlg.OptionWords())
            Catch ex As Exception
                failure = ex.Message
            End Try
            SetBusy(False, Nothing)
            If failure Is Nothing AndAlso Not r.Succeeded Then failure = r.ErrorText
            If failure IsNot Nothing Then
                MessageBox.Show(Me, failure, UiText.T("caption.addpartition"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Await RefreshAsync()
            ShowToolText(UiText.T("win.newpartition", Path.GetFileName(_disk.Path)), r.StdOut)
        End Using
    End Sub

    ''' <summary>View → Partition → Delete this partition: "disk IMG partition --delete N [--force]".</summary>
    Private Async Sub OnPartitionDelete()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsHardDisk OrElse _disk.ActivePartition Is Nothing OrElse ConsoleRunning() Then Return
        Dim part As DiskPartition = _disk.ActivePartition
        Dim fileCount As Integer = _disk.VisibleFiles(True, False).Count
        ' Two whole sentences for the heading, so a translation can order the words itself.
        Dim msg As String = If(part.DiskName.Length > 0,
                UiText.T("msg.delpart.named", part.Number.ToString(CultureInfo.InvariantCulture), part.DiskName, Path.GetFileName(_disk.Path)),
                UiText.T("msg.delpart.head", part.Number.ToString(CultureInfo.InvariantCulture), Path.GetFileName(_disk.Path))) & vbCrLf & vbCrLf &
            If(fileCount = 0, UiText.T("msg.delpart.nofiles"), UiText.T("msg.delpart.files", fileCount.ToString(CultureInfo.InvariantCulture))) & UiText.T("msg.delpart.renumber") & vbCrLf & vbCrLf & UiText.T("msg.delpart.undone")
        If MessageBox.Show(Me, msg, UiText.T("caption.deletepartition"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
        SetBusy(True, UiText.T("status.deletingpartition"))
        Dim r As ToolResult = Nothing
        Dim failure As String = Nothing
        Try
            r = Await Program.Tool.DiskPartitionDeleteAsync(_disk.Path, part.Number, fileCount > 0)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure Is Nothing AndAlso Not r.Succeeded Then failure = r.ErrorText
        If failure IsNot Nothing Then MessageBox.Show(Me, failure, UiText.T("caption.deletepartition"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Await RefreshAsync()
    End Sub

    ' =================================================================== opening / listing

    Private Async Function OpenImageAsync(imagePath As String) As Task
        If Not File.Exists(imagePath) Then
            MessageBox.Show(Me, UiText.TCrLf("msg.filenotfound", imagePath), UiText.T("caption.openerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Program.Settings.RemoveMru(imagePath)
            RebuildRecentMenu()
            Return
        End If
        If Program.Tool.Capabilities Is Nothing OrElse Not Program.Tool.Capabilities.IsUsable Then
            If Not Await CheckEngineAsync(True) Then Return
        End If
        If HostFiles.IsGzipName(imagePath) AndAlso Not Program.Tool.Has520 Then
            MessageBox.Show(Me, UiText.TCrLf("msg.gzippedneeds520", Path.GetFileName(imagePath), Program.Tool.Capabilities.Version.ToString(), "https://www.my-trs-80.com/tool/"), UiText.T("caption.openerror"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        SetBusy(True, UiText.T("status.reading", Path.GetFileName(imagePath)))
        Try
            Dim keepPart As Integer = If(_disk IsNot Nothing AndAlso String.Equals(_disk.Path, imagePath, HostPlatform.PathComparison) AndAlso _disk.ActivePartition IsNot Nothing, _disk.ActivePartition.Number, 0)
            _disk = Await DiskImage.LoadAsync(Program.Tool, Path.GetFullPath(imagePath), keepPart)
            Program.Settings.AddMru(_disk.Path)
            RebuildRecentMenu()
            RebuildPartitionMenu()
            FillList(Nothing)
            UpdateTitle()
            If _disk.Warnings.Count > 0 Then
                _lblStatus.ToolTipText = String.Join(vbCrLf, _disk.Warnings)
            Else
                _lblStatus.ToolTipText = ""
            End If
        Catch ex As ToolException
            _disk = Nothing
            _lv.Items.Clear()
            UpdateTitle()
            MessageBox.Show(Me, ex.Message, UiText.T("caption.openerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            _disk = Nothing
            _lv.Items.Clear()
            UpdateTitle()
            MessageBox.Show(Me, ex.Message, UiText.T("caption.openerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetBusy(False, Nothing)
            UpdateEnabled()
        End Try
    End Function

    ''' <summary>Reload the open image from disk (after a change), keeping the selection.</summary>
    Private Async Function RefreshAsync() As Task
        If _disk Is Nothing Then Return
        Dim selected As List(Of String) = SelectedNames()
        Dim p As String = _disk.Path
        Await OpenImageAsync(p)
        ReselectNames(selected)
    End Function

    Private Async Function SelectPartitionAsync(part As DiskPartition) As Task
        If _disk Is Nothing Then Return
        _disk.ActivePartition = part
        _disk.DosName = part.Dos
        SetBusy(True, UiText.T("status.readingpartition", part.Number.ToString(CultureInfo.InvariantCulture)))
        Try
            Await _disk.RefreshStatusAsync(Program.Tool)
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetBusy(False, Nothing)
        End Try
        RebuildPartitionMenu()
        FillList(Nothing)
        UpdateTitle()
        UpdateEnabled()
    End Function

    ''' <summary>Fill the list from the model, applying the Invisible/System filters.</summary>
    Private Sub FillList(reselect As List(Of String))
        _lv.BeginUpdate()
        Try
            _lv.Items.Clear()
            If _disk IsNot Nothing Then
                For Each f In _disk.VisibleFiles(Program.Settings.ShowInvisible, Program.Settings.ShowSystem)
                    Dim it As New ListViewItem(f.BaseName) With {.Tag = f}
                    it.SubItems.Add(f.Extension)
                    it.SubItems.Add(TextUtil.FormatBytes(f.Size))
                    it.SubItems.Add(f.DateDisplay)
                    Dim attrs As String = f.AttributeLetters
                    If f.Protection.Length > 0 AndAlso Not String.Equals(f.Protection, "FULL", StringComparison.OrdinalIgnoreCase) Then attrs &= If(attrs.Length > 0, " ", "") & f.Protection
                    If f.CpmUser.HasValue AndAlso f.CpmUser.Value <> 0 Then attrs &= UiText.T("file.user", f.CpmUser.Value.ToString(CultureInfo.InvariantCulture))
                    it.SubItems.Add(attrs)
                    Dim kind As String = f.Description
                    If f.Damage.Length > 0 Then kind &= UiText.T("file.damaged", f.Damage)
                    If f.DecodeError.Length > 0 Then kind &= UiText.T("file.decodeerror", f.DecodeError)
                    it.SubItems.Add(kind)
                    ' TRSTool's colours: invisible+system red, system orange, invisible olive.
                    If f.Hidden AndAlso f.IsSystem Then
                        it.ForeColor = Color.Red
                    ElseIf f.IsSystem Then
                        it.ForeColor = Color.DarkOrange
                    ElseIf f.Hidden Then
                        it.ForeColor = Color.Olive
                    End If
                    If f.Damage.Length > 0 Then it.ToolTipText = UiText.T("tip.damaged", f.Damage)
                    _lv.Items.Add(it)
                Next
            End If
            _lv.Sort()
        Finally
            _lv.EndUpdate()
        End Try
        If reselect IsNot Nothing Then ReselectNames(reselect)
        UpdateStatus()
        UpdateEnabled()
    End Sub

    Private Function SelectedFiles() As List(Of DiskFile)
        Dim l As New List(Of DiskFile)
        For Each it As ListViewItem In _lv.SelectedItems
            l.Add(CType(it.Tag, DiskFile))
        Next
        Return l
    End Function

    Private Function AllListedFiles() As List(Of DiskFile)
        Dim l As New List(Of DiskFile)
        For Each it As ListViewItem In _lv.Items
            l.Add(CType(it.Tag, DiskFile))
        Next
        Return l
    End Function

    Private Function SelectedNames() As List(Of String)
        Dim l As New List(Of String)
        For Each f In SelectedFiles()
            l.Add(f.TrsName)
        Next
        Return l
    End Function

    Private Sub ReselectNames(names As List(Of String))
        If names Is Nothing OrElse names.Count = 0 Then Return
        For Each it As ListViewItem In _lv.Items
            Dim f As DiskFile = CType(it.Tag, DiskFile)
            it.Selected = names.Contains(f.TrsName, StringComparer.OrdinalIgnoreCase)
        Next
    End Sub

    Private Sub SelectAll()
        For Each it As ListViewItem In _lv.Items
            it.Selected = True
        Next
    End Sub

    Private Sub InvertSelection()
        For Each it As ListViewItem In _lv.Items
            it.Selected = Not it.Selected
        Next
    End Sub

    Private Sub OnColumnClick(sender As Object, e As ColumnClickEventArgs)
        If e.Column = _sorter.Column Then
            _sorter.Ascending = Not _sorter.Ascending
        Else
            _sorter.Column = e.Column
            _sorter.Ascending = True
        End If
        _lv.Sort()
    End Sub

    Private Sub OnListKeyDown(sender As Object, e As KeyEventArgs)
        If _busy Then Return
        If e.KeyCode = Keys.Enter AndAlso _lv.SelectedItems.Count >= 1 Then
            e.Handled = True
            OnPreview()
        ElseIf e.KeyCode = Keys.Delete AndAlso _lv.SelectedItems.Count > 0 Then
            e.Handled = True
            OnDelete()
        End If
    End Sub

    Private Sub OnListMouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right AndAlso _disk IsNot Nothing Then
            Dim hit As ListViewItem = _lv.GetItemAt(e.X, e.Y)
            If hit IsNot Nothing Then
                If Not hit.Selected Then
                    _lv.SelectedItems.Clear()
                    hit.Selected = True
                End If
                UpdateEnabled()
                _ctx.Show(_lv, e.Location)
            End If
        End If
    End Sub

    ' =================================================================== File menu

    Private Async Sub OnFileOpen()
        If _busy Then Return
        Using dlg As New OpenFileDialog()
            dlg.Title = UiText.T("caption.openvirtualdisk")
            ' trs80-tool 5.2.0 opens gzipped images (game.dmk.gz, big.hdv.gz) like any other.
            Const gzImages As String = "*.dsk.gz;*.dmk.gz;*.jv1.gz;*.jv3.gz;*.hfe.gz;*.imd.gz;*.scp.gz;*.hdv.gz"
            dlg.Filter = UiText.T("filter.openimages", gzImages)
            dlg.CheckFileExists = True
            If dlg.ShowDialog(Me) = DialogResult.OK Then Await OpenImageAsync(dlg.FileName)
        End Using
    End Sub

    Private Sub OnFileClose()
        _disk = Nothing
        _lv.Items.Clear()
        RebuildPartitionMenu()
        UpdateTitle()
        UpdateStatus()
        UpdateEnabled()
    End Sub

    Private Async Sub OnRefresh()
        If _busy Then Return
        Await RefreshAsync()
    End Sub

    Private Async Sub OnFileNew()
        If _busy Then Return
        Using dlg As New CreateFloppyForm()
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Using sfd As New SaveFileDialog()
                sfd.Title = UiText.T("caption.savenewfloppy")
                sfd.DefaultExt = dlg.ImageTypeExtension
                sfd.Filter = dlg.ImageTypeFilter & UiText.T("filter.allfilessuffix")
                sfd.OverwritePrompt = True
                sfd.AddExtension = True
                If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
                Dim target As String = sfd.FileName
                SetBusy(True, UiText.T("status.creating", Path.GetFileName(target)))
                Try
                    If File.Exists(target) Then File.Delete(target)
                    Program.Tool.SetCpmFormat(target, Nothing)     ' a new disk under an old name: forget a layout named for the old one
                    ' A name like new.dmk.gz is a DMK too, for a trs80-tool that writes gzipped files (5.2.0).
                    Dim targetIsDmkName As Boolean = String.Equals(Path.GetExtension(target), ".dmk", StringComparison.OrdinalIgnoreCase) OrElse
                        (Program.Tool.Has520 AndAlso HostFiles.IsGzipName(target) AndAlso HostFiles.ImageExtension(target) = ".dmk")
                    If dlg.ImageTypeExtension = "dmk" AndAlso targetIsDmkName Then
                        Await Program.Tool.DiskCreateAsync(target, dlg.DosName, dlg.CreateOptionWords())
                    Else
                        ' "create" only writes DMK images and insists on a .dmk name: make it in a temp
                        ' folder, then convert (JV1/JV3) or just copy (a DMK under another name).
                        Dim tmpDir As String = Program.NewTempDir()
                        Dim tmpDmk As String = Path.Combine(tmpDir, "new.dmk")
                        Await Program.Tool.DiskCreateAsync(tmpDmk, dlg.DosName, dlg.CreateOptionWords())
                        If dlg.ImageTypeExtension = "dmk" Then
                            File.Copy(tmpDmk, target, True)
                        Else
                            Dim tmpOut As String = Path.Combine(tmpDir, "new." & dlg.ImageTypeExtension)
                            Dim convertError As ToolException = Nothing
                            Try
                                Await Program.Tool.ConvertAsync(tmpDmk, tmpOut, False)
                            Catch ex As ToolException When dlg.ImageTypeExtension = "jv1"
                                convertError = ex
                            End Try
                            If convertError IsNot Nothing Then
                                ' JV1 can't hold some data-address-mark detail; trs80-tool asks for --force.
                                If MessageBox.Show(Me, UiText.TCrLf("msg.jv1force", convertError.Message), UiText.T("dlg.newfloppy.title"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Throw convertError
                                Await Program.Tool.ConvertAsync(tmpDmk, tmpOut, True)
                            End If
                            File.Copy(tmpOut, target, True)
                        End If
                        Try : Directory.Delete(tmpDir, True) : Catch : End Try
                    End If
                Catch ex As Exception
                    SetBusy(False, Nothing)
                    MessageBox.Show(Me, ex.Message, UiText.T("caption.diskcreateerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
                SetBusy(False, Nothing)
                Await OpenImageAsync(target)
            End Using
        End Using
    End Sub

    Private Async Sub OnDiskProperties()
        If _busy OrElse _disk Is Nothing Then Return
        Using dlg As New DiskPropertiesForm(_disk)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim words As List(Of String) = dlg.LabelOptionWords()
            Dim changed As Boolean = False
            Dim failure As String = Nothing
            SetBusy(True, UiText.T("status.updatingdisk"))
            Try
                If words.Count > 0 Then
                    Await Program.Tool.DiskLabelSetAsync(_disk.Path, _disk.PartitionArg, words)
                    changed = True
                End If
                If dlg.WriteProtectChanged Then
                    Await Program.Tool.DiskProtectSetAsync(_disk.Path, dlg.WriteProtectNew)
                    changed = True
                End If
            Catch ex As Exception
                failure = ex.Message
            End Try
            SetBusy(False, Nothing)
            If failure IsNot Nothing Then
                MessageBox.Show(Me, failure, UiText.T("caption.diskpropserror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Await RefreshAsync()
                Return
            End If
            If changed Then Await RefreshAsync()
        End Using
    End Sub

    ''' <summary>File → Save image as: trs80-tool "convert IMG OUT" to DMK/JV1/JV3/HFE/IMD, then open the copy.</summary>
    Private Async Sub OnSaveImageAs()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk OrElse _disk.IsCassette Then Return
        Using sfd As New SaveFileDialog()
            sfd.Title = UiText.T("caption.saveimageastitle")
            sfd.Filter = UiText.T("filter.saveimageas")
            sfd.FileName = Path.GetFileNameWithoutExtension(_disk.Path)
            sfd.OverwritePrompt = True
            sfd.AddExtension = True
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim target As String = sfd.FileName
            Dim ext As String = Path.GetExtension(target).ToLowerInvariant()
            If ext = ".dsk" OrElse ext.Length = 0 Then
                MessageBox.Show(Me, UiText.T("msg.saveas.needext"), UiText.T("caption.saveimageas"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.Equals(Path.GetFullPath(target), _disk.Path, HostPlatform.PathComparison) Then
                MessageBox.Show(Me, UiText.T("msg.saveas.sameimage"), UiText.T("caption.saveimageas"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            SetBusy(True, UiText.T("status.convertingto", Path.GetFileName(target)))
            Dim failure As String = Nothing
            Dim needsForce As ToolException = Nothing
            Try
                If File.Exists(target) Then File.Delete(target)
                Await Program.Tool.ConvertAsync(_disk.Path, target, False)
            Catch ex As ToolException When ex.Message.IndexOf("--force", StringComparison.OrdinalIgnoreCase) >= 0
                needsForce = ex
            Catch ex As Exception
                failure = ex.Message
            End Try
            If needsForce IsNot Nothing Then
                ' e.g. "JV1 can't hold everything on this disk ... Use --force to write it anyway."
                SetBusy(False, Nothing)
                If MessageBox.Show(Me, UiText.TCrLf("msg.saveas.force", needsForce.Message), UiText.T("caption.saveimageas"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
                SetBusy(True, UiText.T("status.convertingto", Path.GetFileName(target)))
                Try
                    Await Program.Tool.ConvertAsync(_disk.Path, target, True)
                Catch ex As Exception
                    failure = ex.Message
                End Try
            End If
            SetBusy(False, Nothing)
            If failure IsNot Nothing Then
                MessageBox.Show(Me, failure, UiText.T("caption.saveimageas"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If MessageBox.Show(Me, UiText.TCrLf("msg.saveas.written", target), UiText.T("caption.saveimageas"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then Await OpenImageAsync(target)
        End Using
    End Sub

    ' =================================================================== Disk menu (trs80-tool features beyond TRSTool)

    Private Sub OnSectorMap()
        If _busy OrElse _disk Is Nothing Then Return
        If _disk.IsHardDisk OrElse _disk.IsCassette Then
            MessageBox.Show(Me, UiText.T("msg.sectormap.floppyonly"), UiText.T("caption.sectormap"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim f As New SectorMapForm(_disk)
        f.Show(Me)
    End Sub

    ''' <summary>Disk -> DMK analysis: the GUI's own reading of the DMK (DmkAnalyzer.vb / DosIdentifier.vb / DiskInfoForm.vb).</summary>
    Private Async Sub OnDmkInfo()
        If _busy OrElse _disk Is Nothing Then Return
        If _disk.IsHardDisk OrElse _disk.IsCassette Then
            MessageBox.Show(Me, UiText.T("msg.dmk.floppyonly"), UiText.T("caption.dmkanalysis"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim f As New DiskInfoForm(_disk.Path)
        f.Show(Me)
        Await f.LoadAsync()
    End Sub

    Private Sub OnGranuleMap()
        If _busy OrElse _disk Is Nothing Then Return
        Dim f As New GranuleMapForm(_disk)
        f.Show(Me)
    End Sub

    Private Async Sub OnCheckRepair()
        If _busy OrElse _disk Is Nothing Then Return
        Dim changed As Boolean
        Using dlg As New CheckRepairForm(_disk)
            dlg.ShowDialog(Me)
            changed = dlg.Changed
        End Using
        If changed Then Await RefreshAsync()
    End Sub

    Private Async Sub OnFormat()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable OrElse _disk.IsHardDisk Then Return
        If HostFiles.ImageExtension(_disk.Path) <> ".dmk" Then
            MessageBox.Show(Me, UiText.T("msg.format.dmkonly"), UiText.T("caption.format"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim cyl As Integer = If(_disk.Geometry IsNot Nothing, _disk.Geometry.CylinderCount, 0)
        Dim sides As Integer = If(_disk.Geometry IsNot Nothing, _disk.Geometry.SideCount, 0)
        Dim density As String = If(_disk.Geometry IsNot Nothing AndAlso _disk.Geometry.DataTrack IsNot Nothing, _disk.Geometry.DataTrack.Density, Nothing)
        Using dlg As New CreateFloppyForm(True, _disk.DosName, cyl, sides, density)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim fileCount As Integer = _disk.VisibleFiles(True, True).Count
            Dim msg As String = UiText.T("msg.format.confirm",
            If(fileCount = 1, UiText.T("msg.format.onefile"), UiText.T("msg.format.manyfiles", fileCount.ToString(CultureInfo.InvariantCulture))),
            Path.GetFileName(_disk.Path),
            If(dlg.IsCpm, UiText.T("dos.cpm.montezuma"), dlg.DosName))
            If MessageBox.Show(Me, msg, UiText.T("dlg.newfloppy.formattitle"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
            SetBusy(True, UiText.T("status.formatting"))
            Dim r As ToolResult = Nothing
            Dim failure As String = Nothing
            Try
                ' A layout named for the old contents doesn't describe the newly formatted disk.
                Program.Tool.SetCpmFormat(_disk.Path, Nothing)
                r = Await Program.Tool.DiskFormatAsync(_disk.Path, Nothing, dlg.DosName, dlg.CreateOptionWords(), True)
            Catch ex As Exception
                failure = ex.Message
            End Try
            SetBusy(False, Nothing)
            If failure Is Nothing AndAlso Not r.Succeeded Then failure = r.ErrorText
            If failure IsNot Nothing Then MessageBox.Show(Me, failure, UiText.T("dlg.newfloppy.formattitle"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Await RefreshAsync()
        End Using
    End Sub
    ''' <summary>
    ''' Disk → Import from another image: hand the window the open disk, let it run "disk import" (and
    ''' "--dry-run" previews) itself, then refresh if anything was saved. The engine does the copying, which
    ''' is why this keeps dates and the system and hidden flags that Copy/Paste and drag cannot.
    ''' </summary>
    Private Async Sub OnImport()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Using dlg As New ImportForm(_disk)
            dlg.ShowDialog(Me)
            If Not dlg.Changed Then Return
            Await RefreshAsync()
            _lblStatus.Text = If(dlg.Copied > 0,
                                 UiText.T("msg.import.done", dlg.Copied.ToString(CultureInfo.InvariantCulture)),
                                 UiText.T("msg.import.nothing"))
        End Using
    End Sub

    Private Async Sub OnTrim()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk Then Return
        Using dlg As New TrimForm(_disk)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            SetBusy(True, UiText.T("status.trimming"))
            Dim r As ToolResult = Nothing
            Dim failure As String = Nothing
            Try
                r = Await Program.Tool.DiskTrimAsync(_disk.Path, dlg.OptionWords())
            Catch ex As Exception
                failure = ex.Message
            End Try
            SetBusy(False, Nothing)
            If failure IsNot Nothing Then
                MessageBox.Show(Me, failure, UiText.T("btn.trim"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Not r.Succeeded Then
                MessageBox.Show(Me, r.ErrorText, UiText.T("btn.trim"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                MessageBox.Show(Me, r.StdOut.Trim(), UiText.T("btn.trim"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Await RefreshAsync()
        End Using
    End Sub

    Private Sub OnBoot()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk OrElse ConsoleRunning() Then Return
        Using dlg As New BootForm(_disk)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            StartConsole(dlg.RunWords(_disk.Path), UiText.T("console.emulator"))
        End Using
    End Sub

    ''' <summary>Tools -> Run the emulator without a disk: "trs80-tool run --model N" in its own console.</summary>
    Private Sub OnRunEmulatorEmpty()
        If _busy OrElse ConsoleRunning() Then Return
        Dim choices As String() = {UiText.T("word.model1.levelii"), UiText.T("word.model3"), UiText.T("word.model4")}
        Dim model As Integer = 3
        Using dlg As New Form() With {.Text = UiText.T("dlg.runmodel.title"), .FormBorderStyle = FormBorderStyle.FixedDialog, .StartPosition = FormStartPosition.CenterParent, .ClientSize = New Size(320, 110), .MaximizeBox = False, .MinimizeBox = False, .ShowInTaskbar = False, .Icon = Program.WindowIcon}
            dlg.Controls.Add(New Label() With {.AutoSize = True, .Location = New Point(12, 14), .Text = UiText.T("dlg.runmodel.machine")})
            Dim cmb As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(12, 36), .Width = 296}
            cmb.Items.AddRange(choices)
            cmb.SelectedIndex = 1
            Dim ok As New Button() With {.Text = UiText.T("btn.run"), .DialogResult = DialogResult.OK, .Location = New Point(142, 74), .Width = 80}
            Dim cancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Location = New Point(228, 74), .Width = 80}
            dlg.Controls.AddRange(New Control() {cmb, ok, cancel})
            ' Make room for the language in force (UiFit.vb); in English nothing here moves.
            UiFit.GrowForm(dlg)
            cmb.Width = Math.Max(cmb.Width, dlg.ClientSize.Width - 24)
            UiFit.ButtonRow(dlg.ClientSize.Width - 12, 6, ok, cancel)
            dlg.AcceptButton = ok : dlg.CancelButton = cancel
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            model = If(cmb.SelectedIndex = 0, 1, If(cmb.SelectedIndex = 1, 3, 4))
        End Using
        StartConsole(New List(Of String) From {"run", "--model", model.ToString(Globalization.CultureInfo.InvariantCulture)}, UiText.T("console.emulator"))
    End Sub

    Private Sub OnShell()
        If _busy OrElse _disk Is Nothing OrElse ConsoleRunning() Then Return
        ' "disk IMAGE [--partition P] [--cpm-format NAME]": the shell reads the image the way this window does.
        Dim words As List(Of String) = Program.Tool.DiskWords(_disk.Path, _disk.PartitionArg)
        StartConsole(words, UiText.T("console.shell"))
    End Sub

    ''' <summary>
    ''' Disk → CP/M layout: name the layout a CP/M floppy is read and written with (trs80-tool's --cpm-format), or
    ''' hand the choice back to trs80-tool. The choice lasts for this image until the program closes.
    ''' </summary>
    Private Async Sub OnCpmFormat()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk OrElse _disk.IsCassette OrElse ConsoleRunning() Then Return
        Dim imagePath As String = _disk.Path
        Dim before As String = Program.Tool.CpmFormatOf(imagePath)
        Dim layouts As String
        SetBusy(True, UiText.T("status.askingcpmlayouts"))
        Try
            layouts = Await Program.Tool.CpmFormatHelpAsync(imagePath)
        Catch ex As Exception
            layouts = ex.Message
        End Try
        SetBusy(False, Nothing)
        Dim chosen As String
        Using dlg As New CpmFormatForm(layouts, before)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            chosen = dlg.Chosen
        End Using
        If String.Equals(chosen, before, StringComparison.OrdinalIgnoreCase) Then Return
        Program.Tool.SetCpmFormat(imagePath, chosen)
        If chosen.Length > 0 Then
            ' Ask trs80-tool whether the disk can be read that way before the window is reloaded with it.
            Dim problem As String = Nothing
            SetBusy(True, UiText.T("status.readingas", Path.GetFileName(imagePath), chosen))
            Try
                Dim info As Trs80Tool.InfoJson = Await Program.Tool.InfoJsonAsync(imagePath)
                problem = info.Error
            Catch ex As Exception
                problem = ex.Message
            End Try
            SetBusy(False, Nothing)
            If problem IsNot Nothing Then
                Program.Tool.SetCpmFormat(imagePath, before)
                ' An unknown name is answered with the whole list again, which the window just showed: keep the first line.
                If problem.Length > 600 AndAlso problem.IndexOf(ControlChars.Lf) > 0 Then problem = problem.Substring(0, problem.IndexOf(ControlChars.Lf)).Trim()
                MessageBox.Show(Me, UiText.TCrLf("msg.cpm.cantread", chosen, problem), UiText.T("dlg.cpm.title"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If
        Await OpenImageAsync(imagePath)
    End Sub

    ''' <summary>Start trs80-tool in its own console window and refresh the image when it exits (it may have written it).</summary>
    Private Sub StartConsole(words As List(Of String), what As String)
        Try
            _console = Program.Tool.StartInConsole(words)
            _consoleWhat = what
            AddHandler _console.Exited, Sub()
                                           Try
                                               BeginInvoke(New Action(Async Sub()
                                                                          _console = Nothing
                                                                          UpdateStatus()
                                                                          UpdateEnabled()
                                                                          If _disk IsNot Nothing Then Await RefreshAsync()
                                                                      End Sub))
                                           Catch
                                           End Try
                                       End Sub
            UpdateStatus()
            UpdateEnabled()
        Catch ex As Exception
            _console = Nothing
            MessageBox.Show(Me, ex.Message, Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =================================================================== Tools menu (host files)

    ''' <summary>Tools → Convert file(s): trs80-tool convert on host files; afterwards offer to add the result to the open disk.</summary>
    Private Async Sub OnConvertFiles(initialInputs As IEnumerable(Of String))
        If _busy Then Return
        Dim produced As String = Nothing
        Using dlg As New ConvertFileForm(initialInputs)
            dlg.ShowDialog(Me)
            produced = dlg.OutputPath
        End Using
        Await OfferAddToDiskAsync(produced)
    End Sub

    Private Async Sub OnAssemble()
        If _busy Then Return
        Dim produced As String = Nothing
        Using dlg As New AssembleForm()
            dlg.ShowDialog(Me)
            produced = dlg.OutputPath
        End Using
        Await OfferAddToDiskAsync(produced)
    End Sub

    ''' <summary>Context menu → Convert to…: extract the selected file to a temp host file, then open the Convert dialog on it.</summary>
    Private Async Sub OnConvertDiskFile()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Dim host As String = Path.Combine(Program.NewTempDir(), TextUtil.HostNameFromTrs(f.TrsName, False))
        SetBusy(True, UiText.T("status.extractingfile", f.SlashName))
        Try
            Await Program.Tool.DiskExtractAsAsync(_disk.Path, _disk.PartitionArg, f.TrsName, host)
        Catch ex As Exception
            SetBusy(False, Nothing)
            MessageBox.Show(Me, UiText.T("msg.convert.cantextract", f.SlashName, ex.Message), UiText.T("caption.convert"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        SetBusy(False, Nothing)
        OnConvertFiles(New String() {host})
    End Sub

    Private Async Function OfferAddToDiskAsync(produced As String) As Task
        If produced Is Nothing OrElse _disk Is Nothing OrElse Not _disk.IsWritable OrElse Not File.Exists(produced) Then Return
        Dim ext As String = Path.GetExtension(produced).ToLowerInvariant()
        If ext = ".dmk" OrElse ext = ".jv1" OrElse ext = ".jv3" OrElse ext = ".hfe" OrElse ext = ".imd" OrElse ext = ".wav" Then Return
        If MessageBox.Show(Me, UiText.T("msg.offeradd", Path.GetFileName(produced), Path.GetFileName(_disk.Path)), Program.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Await AddHostFilesAsync(New String() {produced})
        End If
    End Function

    Private Async Sub OnHexdumpFile()
        If _busy Then Return
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("caption.hexdumpfile")
            ofd.Filter = UiText.T("filter.anyfile")
            If ofd.ShowDialog(Me) <> DialogResult.OK Then Return
            SetBusy(True, UiText.T("status.hexdumping"))
            Dim text As String
            Try
                text = Await Program.Tool.HexdumpAsync(ofd.FileName, False)
            Catch ex As Exception
                text = "(" & ex.Message & ")"
            Finally
                SetBusy(False, Nothing)
            End Try
            Dim f As New TextReportForm(UiText.T("win.hexdump", Path.GetFileName(ofd.FileName)), text)
            f.Show(Me)
        End Using
    End Sub

    Private Async Sub OnDisasmFile()
        If _busy Then Return
        Using ofd As New OpenFileDialog()
            ofd.Title = UiText.T("caption.disasmfile")
            ofd.Filter = UiText.T("filter.trs80programs")
            If ofd.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim words As List(Of String)
            Using opt As New DisasmOptionsForm(Program.Settings.DisasmOptionWords())
                If opt.ShowDialog(Me) <> DialogResult.OK Then Return
                words = opt.OptionWords()
            End Using
            SetBusy(True, UiText.T("status.disassembling"))
            Dim text As String
            Try
                text = Await Program.Tool.DisasmAsync(ofd.FileName, words)
            Catch ex As Exception
                text = "(" & ex.Message & ")"
            Finally
                SetBusy(False, Nothing)
            End Try
            Dim f As New TextReportForm(UiText.T("win.disasm", Path.GetFileName(ofd.FileName)), text)
            f.Show(Me)
        End Using
    End Sub

    ' =================================================================== View menu

    Private Sub OnViewInvisible()
        Program.Settings.ShowInvisible = Not Program.Settings.ShowInvisible
        UpdateViewChecks()
        FillList(SelectedNames())
    End Sub

    Private Sub OnViewSystem()
        Program.Settings.ShowSystem = Not Program.Settings.ShowSystem
        UpdateViewChecks()
        FillList(SelectedNames())
    End Sub

    ' =================================================================== Edit: preferences / about / engine

    Private Async Sub OnPreferences()
        Using dlg As New PreferencesForm()
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                Program.Settings.Save()
                Dim newPath As String = Trs80Tool.Locate(Program.Settings.ToolPath)
                If newPath IsNot Nothing AndAlso Not String.Equals(newPath, Program.Tool.ToolPath, HostPlatform.PathComparison) Then
                    Program.Tool.ToolPath = newPath
                    Program.Tool.Capabilities = Nothing
                    Await CheckEngineAsync(False)
                ElseIf newPath Is Nothing Then
                    Program.Tool.ToolPath = ""
                    Program.Tool.Capabilities = Nothing
                    _lblEngine.Text = UiText.T("status.enginenotfound")
                End If
                FillList(SelectedNames())
            End If
        End Using
    End Sub

    Private Sub OnAbout()
        MessageBox.Show(Me, Program.AppName & " " & Program.ReleaseVersion & vbCrLf &
                        UiText.T("about.releasedate", Program.ReleaseDate) & vbCrLf &
                        UiText.T("about.build", Program.AppVersion) & vbCrLf &
                        UiText.T("about.by", Program.AppAuthors) & vbCrLf &
                        UiText.T("about.license") & vbCrLf &
                        UiText.T("about.download", "https://github.com/jdmba/TRSToolGui") & vbCrLf & vbCrLf &
                        UiText.T("about.what") & vbCrLf & vbCrLf &
                        UiText.T("about.engine", If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, UiText.T("status.engineversion", Program.Tool.Capabilities.VersionText), UiText.T("about.enginenotfound"))) & vbCrLf &
                        UiText.T("about.enginepath", If(Program.Tool.ToolPath, "")) & vbCrLf & vbCrLf &
                        UiText.T("about.toolcredit", "https://www.my-trs-80.com/tool/") & vbCrLf &
                        UiText.T("about.trstoolscredit"),
                        UiText.T("menu.mac.about", Program.AppName), MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>Help > Manual / Changelog / License: an embedded text file in the report window (Copy / Save / Close).</summary>
    Private Sub ShowHelpText(resourceName As String, title As String)
        Dim f As New TextReportForm(title, Program.ReadHelpText(resourceName))
        f.Show(Me)
    End Sub

    Private Async Sub OnEngineInfo()
        If _busy Then Return
        If Program.Tool.Capabilities Is Nothing AndAlso Not String.IsNullOrEmpty(Program.Tool.ToolPath) Then Await CheckEngineAsync(False)
        Dim reprobe As Boolean = False
        Using dlg As New EngineInfoForm(Program.Tool)
            dlg.ShowDialog(Me)
            reprobe = dlg.ReprobeRequested
        End Using
        If reprobe Then Await CheckEngineAsync(True)
    End Sub

    ' =================================================================== viewer / properties / rename / delete

    Private Sub OnPreview()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count = 0 Then Return
        Dim f As DiskFile = sel(0)
        Dim v As New ViewerForm(_disk, f)
        v.Show(Me)
    End Sub

    Private Async Sub OnFileProperties()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Using dlg As New FilePropertiesForm(_disk, f)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim changed As Boolean = False
            Dim failure As String = Nothing
            SetBusy(True, UiText.T("status.updatingfile", f.SlashName))
            Try
                Dim attribWords As List(Of String) = dlg.AttribOptionWords()
                If attribWords.Count > 0 Then
                    Await Program.Tool.DiskAttribAsync(_disk.Path, _disk.PartitionArg, f.TrsName, attribWords)
                    changed = True
                End If
                If dlg.NewName IsNot Nothing AndAlso Not String.Equals(dlg.NewName, f.TrsName, StringComparison.OrdinalIgnoreCase) Then
                    Await Program.Tool.DiskRenameAsync(_disk.Path, _disk.PartitionArg, f.TrsName, dlg.NewName)
                    changed = True
                End If
            Catch ex As Exception
                failure = ex.Message
            End Try
            SetBusy(False, Nothing)
            If failure IsNot Nothing Then
                MessageBox.Show(Me, failure, UiText.T("caption.filepropserror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Await RefreshAsync()
                Return
            End If
            If changed Then Await RefreshAsync()
        End Using
    End Sub

    Private Async Sub OnRename()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Using dlg As New RenameForm(UiText.T("dlg.renamefile.title"), UiText.T("dlg.renamefile.prompt"), f.SlashName, Nothing)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim newName As String = TextUtil.NormalizeTrsName(dlg.ResultName)
            If newName Is Nothing OrElse String.Equals(newName, f.TrsName, StringComparison.OrdinalIgnoreCase) Then Return
            SetBusy(True, UiText.T("status.renaming"))
            Try
                Await Program.Tool.DiskRenameAsync(_disk.Path, _disk.PartitionArg, f.TrsName, newName)
            Catch ex As Exception
                SetBusy(False, Nothing)
                MessageBox.Show(Me, UiText.T("msg.renamefailed", ex.Message), UiText.T("caption.renameerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try
            SetBusy(False, Nothing)
            Await RefreshAsync()
            ReselectNames(New List(Of String) From {newName})
        End Using
    End Sub

    Private Async Sub OnDelete()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count = 0 Then Return
        If Program.Settings.ConfirmDelete Then
            Dim msg As String = If(sel.Count = 1, UiText.TCrLf("msg.delete.one", sel(0).SlashName), UiText.T("msg.delete.many", sel.Count.ToString(CultureInfo.InvariantCulture)))
            If MessageBox.Show(Me, msg, UiText.T("caption.confirmdelete"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return
        End If
        Dim names As New List(Of String)
        For Each f In sel
            names.Add(f.TrsName)
        Next
        SetBusy(True, UiText.T("status.deleting"))
        Dim failure As String = Nothing
        Try
            Await Program.Tool.DiskKillAsync(_disk.Path, _disk.PartitionArg, names)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then MessageBox.Show(Me, UiText.T("msg.deletefailed", failure), UiText.T("caption.deleteerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Await RefreshAsync()
    End Sub

    ' =================================================================== adding files (Add / Paste / drop)

    Private Async Sub OnEditAdd()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Using dlg As New OpenFileDialog()
            dlg.Title = UiText.T("caption.selectfilestoadd")
            dlg.Filter = UiText.T("filter.anyfile")
            dlg.Multiselect = True
            dlg.CheckFileExists = True
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Await AddHostFilesAsync(dlg.FileNames)
        End Using
    End Sub

    Private Async Sub OnEditPaste()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        If Not Clipboard.ContainsFileDropList() Then Return
        Dim files As StringCollection = Clipboard.GetFileDropList()
        Dim paths As New List(Of String)
        For Each p In files
            If File.Exists(p) Then paths.Add(p)
        Next
        If paths.Count > 0 Then Await AddHostFilesAsync(paths)
    End Sub

    Private Sub OnDragEnterList(sender As Object, e As DragEventArgs)
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then
            e.Effect = DragDropEffects.None
            Return
        End If
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Our own drag-out carries a marker; refuse dropping a disk's files back on itself.
            If e.Data.GetDataPresent("TRSToolGui.SelfDrag") Then
                e.Effect = DragDropEffects.None
            Else
                e.Effect = DragDropEffects.Copy
            End If
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Async Sub OnDragDropList(sender As Object, e As DragEventArgs)
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return
        Dim dropped As String() = TryCast(e.Data.GetData(DataFormats.FileDrop), String())
        If dropped Is Nothing Then Return
        Dim paths As New List(Of String)
        For Each p In dropped
            If File.Exists(p) Then
                paths.Add(p)
            ElseIf Directory.Exists(p) Then
                ' A dropped folder: its files (not recursive), like dropping them one by one.
                paths.AddRange(Directory.GetFiles(p))
            End If
        Next
        Activate()
        If paths.Count > 0 Then Await AddHostFilesAsync(paths)
    End Sub

    ''' <summary>
    ''' Store host files on the disk, with TRSTool's checks: name validity (rename dialog), existing
    ''' file (replace dialog), text conversion preference, host date. One all-or-nothing engine call.
    ''' </summary>
    Private Async Function AddHostFilesAsync(hostPaths As IEnumerable(Of String)) As Task
        If _disk Is Nothing OrElse Not _disk.IsWritable Then
            MessageBox.Show(Me, UiText.T("msg.write.cannot") & If(_disk IsNot Nothing AndAlso _disk.ReadOnlyReason IsNot Nothing, vbCrLf & vbCrLf & _disk.ReadOnlyReason, ""), UiText.T("caption.filewriteerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim s As Settings = Program.Settings
        Dim items As New List(Of Trs80Tool.AddItem)
        Dim replaceAll As Boolean? = Nothing     ' Yes to all / No to all
        Dim pendingNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each hostPath In hostPaths
            Dim fi As New FileInfo(hostPath)
            If Not fi.Exists Then Continue For
            If fi.Length > 10L * 1024 * 1024 Then
                MessageBox.Show(Me, UiText.T("msg.add.toolarge", fi.Name), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Continue For
            End If

            ' 0. A file named *.gz. trs80-tool 5.2.0 unzips such a file by itself and stores what is inside.
            '    The preference decides instead: stored as it is (TRSTool's way: a copy under a name the tool
            '    doesn't unzip), or unzipped here first, so that the name, the text test and the conversion
            '    all work on the file inside.
            Dim sourcePath As String = fi.FullName       ' the file trs80-tool is given
            Dim sourceName As String = fi.Name           ' the name the TRS-80 name is made from
            If HostFiles.IsGzipName(fi.Name) Then
                Try
                    If s.UnzipGzOnAdd Then
                        sourcePath = HostFiles.GunzipInto(fi.FullName, Program.NewTempDir())
                        sourceName = Path.GetFileName(sourcePath)
                    ElseIf Program.Tool.Has520 Then
                        sourcePath = HostFiles.CopyUnderPlainName(fi.FullName, Program.NewTempDir())
                    End If
                Catch ex As Exception
                    MessageBox.Show(Me, UiText.T("msg.add.gzfailed", fi.Name, If(s.UnzipGzOnAdd, UiText.T("msg.add.gzunzipfailed"), UiText.T("msg.add.gzcopyfailed")), ex.Message), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Continue For
                End Try
            End If

            ' 1. TRS-80 name
            Dim trsName As String = TextUtil.TrsNameFromHost(sourceName)
            Dim hostUpper As String = sourceName.ToUpperInvariant()
            Dim needsRename As Boolean = trsName.Length = 0 OrElse Not String.Equals(trsName, hostUpper, StringComparison.Ordinal)
            If needsRename Then
                Using dlg As New RenameForm(UiText.T("dlg.renamefile.title"), UiText.T("dlg.renamefile.newprompt"), If(trsName.Length = 0, "", trsName.Replace(".", "/")), UiText.TCrLf("msg.add.invalidname", sourceName))
                    If dlg.ShowDialog(Me) <> DialogResult.OK Then Continue For
                    trsName = TextUtil.NormalizeTrsName(dlg.ResultName)
                    If trsName Is Nothing Then
                        MessageBox.Show(Me, UiText.T("msg.add.badnameskip", dlg.ResultName, fi.Name), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Continue For
                    End If
                End Using
            End If
            If pendingNames.Contains(trsName) Then
                MessageBox.Show(Me, UiText.T("msg.add.duplicate", trsName.Replace(".", "/"), fi.Name), Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Continue For
            End If

            ' 2. existing file?
            Dim existing As DiskFile = _disk.FindFile(trsName)
            Dim overwrite As Boolean = False
            If existing IsNot Nothing Then
                If replaceAll.HasValue Then
                    If Not replaceAll.Value Then Continue For
                    overwrite = True
                Else
                    Using dlg As New ReplaceForm(existing, trsName.Replace(".", "/"), fi.Length, fi.LastWriteTime)
                        Dim r As DialogResult = dlg.ShowDialog(Me)
                        If dlg.ApplyToAll Then replaceAll = (r = DialogResult.Yes)
                        If r <> DialogResult.Yes Then Continue For
                        overwrite = True
                    End Using
                End If
            End If

            ' 3. conversion
            Dim ext As String = Path.GetExtension(sourceName).ToLowerInvariant()
            Dim item As New Trs80Tool.AddItem With {.HostPath = sourcePath, .Overwrite = overwrite, .Binary = True}
            If (ext = ".bas" OrElse ext = ".cas" OrElse ext = ".3bn") AndAlso s.TokenizeBasicOnAdd Then
                item.Binary = False       ' let trs80-tool tokenize / convert cassette programs
            ElseIf s.ConversionToTrs80 = TextConversionIn.CrLfToCr AndAlso Not _disk.IsCpm Then
                Dim data As Byte() = File.ReadAllBytes(sourcePath)
                If TextUtil.LooksLikeText(ext.TrimStart("."c), data) AndAlso Array.IndexOf(data, CByte(10)) >= 0 Then
                    ' TRSTool converted by content, whatever the extension: "add --text" does exactly that
                    ' (line ends become CR and nothing else), so the test is ours and the conversion the tool's.
                    item.Binary = False
                    item.Text = True
                End If
            End If
            ' trs80-tool derives the name from the host name; force ours whenever it differs (or we renamed a temp copy).
            If needsRename OrElse Not String.Equals(item.HostPath, fi.FullName, StringComparison.Ordinal) OrElse Not String.Equals(TextUtil.TrsNameFromHost(Path.GetFileName(item.HostPath)), trsName, StringComparison.Ordinal) Then
                item.AsName = trsName
            End If
            If s.UseHostDateOnAdd AndAlso _disk.IsTrsdosFamily Then item.DateYmd = fi.LastWriteTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            items.Add(item)
            pendingNames.Add(trsName)
        Next

        If items.Count = 0 Then
            Return
        End If
        ' trs80-tool 5.3.0: with more than one file, let the tool skip one it can't add and add the rest,
        ' instead of the whole run adding nothing. Without it 5.3.0 also puts "use --keep-going to add the
        ' others" in the error, which is a command-line flag nobody can type in a window.
        If items.Count > 1 AndAlso Program.Tool.CanKeepGoing Then
            For Each it In items
                it.KeepGoing = True
            Next
        End If

        SetBusy(True, UiText.T("status.adding", items.Count.ToString(CultureInfo.InvariantCulture)))
        Dim result As ToolResult = Nothing
        Dim failure As String = Nothing
        Dim dateNote As String = Nothing
        Dim failureKind As ToolFailure = ToolFailure.None
        Try
            result = Await Program.Tool.DiskAddAsync(_disk.Path, _disk.PartitionArg, items)
        Catch ex As ToolException
            failure = ex.Message
            failureKind = ex.Failure
        Catch ex As Exception
            failure = ex.Message
        End Try
        ' LDOS 5 (unpatched) can't store years after 2011, and other DOSes have their own limits: the tool
        ' refuses the whole add. Retry once without dates, as TRSTool's users had to live with before the 2012 patches.
        If failure IsNot Nothing AndAlso items.Any(Function(it) it.DateYmd IsNot Nothing) AndAlso Regex.IsMatch(failure, "year|date", RegexOptions.IgnoreCase) Then
            For Each it In items
                it.DateYmd = Nothing
            Next
            dateNote = failure
            failure = Nothing
            failureKind = ToolFailure.None
            Try
                result = Await Program.Tool.DiskAddAsync(_disk.Path, _disk.PartitionArg, items)
            Catch ex As ToolException
                failure = ex.Message
                failureKind = ex.Failure
            Catch ex As Exception
                failure = ex.Message
            End Try
        End If
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            ' trs80-tool's exit status says why: TRSTool's own wording for the two cases it had.
            Dim heading As String = UiText.T("msg.store.error")
            If failureKind = ToolFailure.DiskFull Then heading = If(items.Count = 1, UiText.T("msg.store.diskfull.one"), UiText.T("msg.store.diskfull.many"))
            If failureKind = ToolFailure.WriteProtected Then heading = UiText.T("msg.store.writeprotected")
            MessageBox.Show(Me, heading & vbCrLf & vbCrLf & failure & vbCrLf & vbCrLf & UiText.T("msg.store.nothingwritten"), UiText.T("caption.storeerror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Await RefreshAsync()
            Return
        End If
        ' Surface the tool's warnings (e.g. the DOS couldn't store the date), once.
        Dim warn As New List(Of String)
        If dateNote IsNot Nothing Then warn.Add(UiText.T("msg.add.nodates", dateNote))
        ' With --keep-going the run can succeed having left files out; the tool says which and why.
        Dim skippedNote As String = SkippedAddNote(result, items.Count)
        If skippedNote IsNot Nothing Then warn.Add(skippedNote)
        If result.Disk IsNot Nothing Then
            For Each w In result.Disk.AllWarnings()
                If Not warn.Contains(w) Then warn.Add(w)
            Next
        End If
        Await RefreshAsync()
        Dim added As New List(Of String)
        For Each it In items
            added.Add(If(it.AsName, TextUtil.TrsNameFromHost(Path.GetFileName(it.HostPath))))
        Next
        ReselectNames(added)
        If warn.Count > 0 Then MessageBox.Show(Me, String.Join(vbCrLf, warn), UiText.T("caption.toolnotes"), MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Function
    ''' <summary>
    ''' What "add --keep-going" left out, as one block of text, or Nothing when every file went in. The
    ''' "skipped" list is trs80-tool 5.3.0's; an older tool has no such field and this returns Nothing.
    ''' </summary>
    Private Shared Function SkippedAddNote(result As ToolResult, wanted As Integer) As String
        If result Is Nothing OrElse result.Disk Is Nothing OrElse result.Disk.Commands Is Nothing Then Return Nothing
        Dim added As Integer = 0
        Dim lines As New List(Of String)
        For Each c In result.Disk.Commands
            If Not String.Equals(c.Name, "add", StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim res As Trs80Tool.AddResultJson = c.GetResult(Of Trs80Tool.AddResultJson)()
            If res Is Nothing Then Continue For
            If res.Files IsNot Nothing Then added += res.Files.Count
            If res.Skipped Is Nothing Then Continue For
            For Each sk In res.Skipped
                lines.Add("    " & Path.GetFileName(If(sk.Host, "")) & ": " & If(sk.Problem, ""))
            Next
        Next
        If lines.Count = 0 Then Return Nothing
        Return UiText.T("msg.add.partial", added.ToString(CultureInfo.InvariantCulture),
                        wanted.ToString(CultureInfo.InvariantCulture)) & vbCrLf & String.Join(vbCrLf, lines)
    End Function

    ' =================================================================== extracting (Extract / Copy / drag out)

    Private Async Sub OnEditExtract()
        If _busy OrElse _disk Is Nothing Then Return
        Using dlg As New ExtractForm(_lv.SelectedItems.Count)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim files As List(Of DiskFile) = If(dlg.ExtractAll, AllListedFiles(), SelectedFiles())
            If files.Count = 0 Then Return
            Dim targetDir As String = dlg.Directory
            Try
                If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, UiText.T("caption.extracterror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try
            Program.Settings.ExtractDirectory = targetDir
            Program.Settings.LowercaseExtractNames = dlg.Lowercase
            Dim staged As List(Of String) = Await ExportToTempAsync(files, dlg.Options)
            If staged Is Nothing Then Return
            Dim overwriteAll As Boolean? = Nothing
            Dim written As Integer = 0
            For Each src In staged
                Dim dest As String = Path.Combine(targetDir, Path.GetFileName(src))
                If File.Exists(dest) Then
                    Dim doIt As Boolean
                    If overwriteAll.HasValue Then
                        doIt = overwriteAll.Value
                    Else
                        Dim r As DialogResult = MessageBox.Show(Me, UiText.T("dlg.overwrite.exists", Path.GetFileName(dest), targetDir) & vbCrLf & vbCrLf & UiText.T("dlg.overwrite.question") & vbCrLf & vbCrLf & UiText.T("dlg.overwrite.cancelnote"), UiText.T("dlg.overwrite.title"), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
                        If r = DialogResult.Cancel Then Exit For
                        doIt = (r = DialogResult.Yes)
                        If staged.Count > 1 AndAlso ModifierKeys.HasFlag(Keys.Shift) Then overwriteAll = doIt
                    End If
                    If Not doIt Then Continue For
                End If
                Try
                    File.Copy(src, dest, True)
                    written += 1
                Catch ex As Exception
                    MessageBox.Show(Me, UiText.T("msg.extractfailed", ex.Message), UiText.T("caption.extracterror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Next
            _lblStatus.Text = UiText.T("status.extracted", written.ToString(CultureInfo.InvariantCulture), targetDir)
        End Using
    End Sub

    Private Async Sub OnEditCopy()
        If _busy OrElse _disk Is Nothing Then Return
        Dim files As List(Of DiskFile) = SelectedFiles()
        If files.Count = 0 Then Return
        Dim staged As List(Of String) = Await ExportToTempAsync(files, New ExtractOptions With {.Lowercase = Program.Settings.LowercaseExtractNames})
        If staged Is Nothing OrElse staged.Count = 0 Then Return
        Dim sc As New StringCollection()
        sc.AddRange(staged.ToArray())
        Try
            Clipboard.SetFileDropList(sc)
            _lblStatus.Text = UiText.T("status.copied", staged.Count.ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private _dragging As Boolean = False

    Private Async Sub OnItemDrag(sender As Object, e As ItemDragEventArgs)
        If _busy OrElse _disk Is Nothing OrElse _dragging Then Return
        Dim files As List(Of DiskFile) = SelectedFiles()
        If files.Count = 0 Then Return
        _dragging = True
        Try
            Dim staged As List(Of String) = Await ExportToTempAsync(files, New ExtractOptions With {.Lowercase = Program.Settings.LowercaseExtractNames})
            If staged Is Nothing OrElse staged.Count = 0 Then Return
            Dim data As New DataObject()
            data.SetData(DataFormats.FileDrop, True, staged.ToArray())
            data.SetData("TRSToolGui.SelfDrag", True, _disk.Path)
            _lv.DoDragDrop(data, DragDropEffects.Copy)
        Finally
            _dragging = False
        End Try
    End Sub

    ''' <summary>
    ''' Extract files to a fresh temp folder (the engine writes NAME.EXT), apply the CR-conversion
    ''' preference to text files, and return the host paths. Returns Nothing on failure (already reported).
    ''' </summary>
    Private Async Function ExportToTempAsync(files As List(Of DiskFile), options As ExtractOptions) As Task(Of List(Of String))
        Dim folder As String = Program.NewTempDir()
        Dim lowercase As Boolean = options.Lowercase
        Dim names As New List(Of String)
        For Each f In files
            names.Add(If(options.NeedsTopLevelExtract, f.FullName, f.TrsName))
        Next
        SetBusy(True, UiText.T("status.extracting", files.Count.ToString(CultureInfo.InvariantCulture)))
        Try
            If options.NeedsTopLevelExtract Then
                ' --cpm-records / --ucsd-pascal-raw exist only on the top-level extract command.
                Await Program.Tool.ExtractTopLevelAsync(_disk.Path, names, folder, lowercase, options.MarkDamaged, options.CpmRecords, options.UcsdRaw)
            Else
                Await Program.Tool.DiskExtractAsync(_disk.Path, _disk.PartitionArg, names, folder, lowercase, options.MarkDamaged)
            End If
        Catch ex As Exception
            SetBusy(False, Nothing)
            MessageBox.Show(Me, UiText.T("msg.extractfailedmulti", ex.Message), UiText.T("caption.extracterror"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
        Dim result As New List(Of String)
        Try
            For Each f In files
                Dim host As String = Path.Combine(folder, TextUtil.HostNameFromTrs(f.TrsName, lowercase))
                If Not File.Exists(host) Then
                    ' The tool may have added " (2)" or "[DAMAGED]"; take whatever it wrote for this name.
                    Dim candidates As String() = Directory.GetFiles(folder, Path.GetFileNameWithoutExtension(host) & "*")
                    If candidates.Length > 0 Then host = candidates(0) Else Continue For
                End If
                If Program.Settings.ConversionFromTrs80 <> TextConversionOut.None AndAlso Not _disk.IsCpm Then
                    Dim data As Byte() = File.ReadAllBytes(host)
                    If TextUtil.LooksLikeText(f.Extension, data) Then
                        File.WriteAllBytes(host, TextUtil.ConvertFromTrs80(data, Program.Settings.ConversionFromTrs80))
                    End If
                End If
                If f.DateValue.HasValue Then
                    Try : File.SetLastWriteTime(host, f.DateValue.Value) : Catch : End Try
                End If
                result.Add(host)
            Next
        Finally
            SetBusy(False, Nothing)
        End Try
        Return result
    End Function

End Class

''' <summary>Column sorter: Name/Type/Attributes/Kind by text, Size numerically, Date chronologically.</summary>
Public Class ListViewColumnSorter
    Implements IComparer

    Public Property Column As Integer = 0
    Public Property Ascending As Boolean = True

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim a As ListViewItem = CType(x, ListViewItem)
        Dim b As ListViewItem = CType(y, ListViewItem)
        Dim fa As DiskFile = TryCast(a.Tag, DiskFile)
        Dim fb As DiskFile = TryCast(b.Tag, DiskFile)
        Dim r As Integer
        Select Case Column
            Case 2
                r = If(fa IsNot Nothing AndAlso fb IsNot Nothing, fa.Size.CompareTo(fb.Size), 0)
            Case 3
                Dim da As DateTime = If(fa IsNot Nothing AndAlso fa.DateValue.HasValue, fa.DateValue.Value, DateTime.MinValue)
                Dim db As DateTime = If(fb IsNot Nothing AndAlso fb.DateValue.HasValue, fb.DateValue.Value, DateTime.MinValue)
                r = da.CompareTo(db)
            Case Else
                Dim sa As String = If(Column < a.SubItems.Count, a.SubItems(Column).Text, "")
                Dim sb As String = If(Column < b.SubItems.Count, b.SubItems(Column).Text, "")
                r = String.Compare(sa, sb, StringComparison.OrdinalIgnoreCase)
        End Select
        If r = 0 AndAlso fa IsNot Nothing AndAlso fb IsNot Nothing Then r = String.Compare(fa.TrsName, fb.TrsName, StringComparison.OrdinalIgnoreCase)
        Return If(Ascending, r, -r)
    End Function
End Class
