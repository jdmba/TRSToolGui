' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  MainForm.vb — the TRSTools-style main window of the Eto.Forms program:
'  menu, toolbar, file list, status line, context menu, drop target, clipboard.
'  Every disk operation is delegated to Trs80Tool (the engine adapter); this
'  form only decides *what* to ask for and shows the result. It follows the
'  Windows Forms MainForm command for command. On a Mac the menu is laid out
'  the way Mac programs have it (BuildMacMenuBar).
' ============================================================================
Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Eto.Drawing
Imports Eto.Forms

Public Class MainForm
    Inherits Form

    Private ReadOnly _grid As New GridView()
    Private ReadOnly _lblStatus As New Label()
    Private ReadOnly _lblEngine As New Label()
    Private _rows As New List(Of DiskFile)
    Private _sortColumn As Integer = 0
    Private _sortAscending As Boolean = True

    ' menu items that need enabling/disabling
    Private _miClose, _miDiskProps, _miCopy, _miPaste, _miAdd, _miExtract, _miSelectAll, _miInvert As ButtonMenuItem
    Private _miPartition, _miRefresh, _miRecent, _miSaveAs, _miNewHardDisk As ButtonMenuItem
    Private _miInvisible, _miSystem As CheckMenuItem
    Private _miSectorMap, _miGranuleMap, _miCheckRepair, _miFormat, _miTrim, _miBoot, _miShell, _miDmkInfo, _miRunEmpty, _miCpmFormat As ButtonMenuItem
    Private _cxPreview, _cxExtract, _cxDelete, _cxRename, _cxProps, _cxConvert As ButtonMenuItem
    Private _tbProps, _tbCopy, _tbPaste, _tbAdd, _tbExtract, _tbView, _tbSectors As ButtonToolItem
    ' Mac only: the system's own Edit > Copy, Paste and Delete act on the file list through these.
    Private _cmdCopy, _cmdPaste, _cmdDelete As Command
    Private _settingChecks As Boolean = False

    ''' <summary>A trs80-tool console (disk shell or emulator) that may be writing the open image; mutations wait for it.</summary>
    Private _console As ConsoleRun = Nothing
    Private _consoleWhat As String = ""

    Private _disk As DiskImage = Nothing
    Private _busy As Boolean = False

    ' The size asked for at start and what the window then reports as its inside. The two differ by the menu
    ' and toolbar on some systems; the difference is added back when the size is saved, so the window comes
    ' back the size it was left.
    Private _requestedSize As Size
    Private _sizeDelta As Size
    Private _sizeKnown As Boolean = False
    Private ReadOnly _startFile As String

    Public Sub New(startFile As String)
        _startFile = startFile
        BuildUi()
    End Sub

    ' =================================================================== UI construction

    Private Sub BuildUi()
        Title = Program.AppName
        Dim s As Settings = Program.Settings
        _requestedSize = New Size(Math.Max(480, s.WindowWidth), Math.Max(300, s.WindowHeight))
        ClientSize = _requestedSize
        If s.WindowX >= 0 AndAlso s.WindowY >= 0 AndAlso Screen.Screens.Any(Function(sc) sc.WorkingArea.Contains(New PointF(s.WindowX + 20, s.WindowY + 20))) Then
            Location = New Point(s.WindowX, s.WindowY)
        End If
        If s.WindowMaximized Then WindowState = WindowState.Maximized

        Dim ctrl As Keys = Application.Instance.CommonModifier

        ' ---- menu
        Dim mFile As New ButtonMenuItem With {.Text = "&File"}
        mFile.Items.Add(MakeItem("&New...", ctrl Or Keys.N, AddressOf OnFileNew))
        _miNewHardDisk = MakeItem("New &hard disk image...", Keys.None, AddressOf OnFileNewHardDisk)
        mFile.Items.Add(_miNewHardDisk)
        mFile.Items.Add(MakeItem("&Open...", ctrl Or Keys.O, AddressOf OnFileOpen))
        _miRecent = New ButtonMenuItem With {.Text = "Open Recent"}
        mFile.Items.Add(_miRecent)
        ' A Mac's File menu gets the system's own Close (the window), so this one says what it closes.
        _miClose = MakeItem(If(HostPlatform.IsMacOS, "&Close image", "&Close"), Keys.None, AddressOf OnFileClose)
        mFile.Items.Add(_miClose)
        mFile.Items.Add(New SeparatorMenuItem())
        _miDiskProps = MakeItem("Disk &properties...", Keys.None, AddressOf OnDiskProperties)
        mFile.Items.Add(_miDiskProps)
        _miSaveAs = MakeItem("Save &image as... (convert format)", Keys.None, AddressOf OnSaveImageAs)
        mFile.Items.Add(_miSaveAs)
        If Not HostPlatform.IsMacOS Then
            mFile.Items.Add(New SeparatorMenuItem())
            mFile.Items.Add(MakeItem("E&xit", ctrl Or Keys.Q, Sub() Program.ExitProgram(Me)))
        End If

        Dim mEdit As New ButtonMenuItem With {.Text = "&Edit"}
        _miCopy = MakeItem("&Copy", ctrl Or Keys.C, AddressOf OnEditCopy)
        _miPaste = MakeItem("&Paste", ctrl Or Keys.V, AddressOf OnEditPaste)
        _miAdd = MakeItem("&Add files...", Keys.None, AddressOf OnEditAdd)
        _miExtract = MakeItem("&Extract...", Keys.None, AddressOf OnEditExtract)
        _miSelectAll = MakeItem("Select &all", ctrl Or Keys.A, Sub() SelectAllRows())
        _miInvert = MakeItem("&Invert selection", Keys.None, Sub() InvertSelection())
        If HostPlatform.IsMacOS Then
            ' The system puts Undo, Cut, Copy, Paste, Delete and Select All into this menu (BuildMacMenuBar), and
            ' Preferences is in the application menu. What is left goes below them.
            For Each mi As MenuItem In New MenuItem() {New SeparatorMenuItem(), _miAdd, _miExtract, New SeparatorMenuItem(), _miInvert}
                mi.Order = 500
                mEdit.Items.Add(mi)
            Next
        Else
            mEdit.Items.AddRange(New MenuItem() {_miCopy, _miPaste, New SeparatorMenuItem(), _miAdd, _miExtract, New SeparatorMenuItem(), _miSelectAll, _miInvert, New SeparatorMenuItem(), MakeItem("P&references...", Keys.None, AddressOf OnPreferences)})
        End If

        Dim mView As New ButtonMenuItem With {.Text = "&View"}
        _miInvisible = New CheckMenuItem With {.Text = "&Invisible files"}
        AddHandler _miInvisible.CheckedChanged, Sub() OnViewToggle()
        _miSystem = New CheckMenuItem With {.Text = "S&ystem files"}
        AddHandler _miSystem.CheckedChanged, Sub() OnViewToggle()
        _miPartition = New ButtonMenuItem With {.Text = "&Partition"}
        _miRefresh = MakeItem("&Refresh", Keys.F5, AddressOf OnRefresh)
        mView.Items.AddRange(New MenuItem() {_miInvisible, _miSystem, _miPartition, New SeparatorMenuItem(), _miRefresh})

        ' ---- Disk and Tools menus: trs80-tool features beyond TRSTool
        Dim mDisk As New ButtonMenuItem With {.Text = "&Disk"}
        _miSectorMap = MakeItem("&Sector map...", ctrl Or Keys.M, AddressOf OnSectorMap)
        _miDmkInfo = MakeItem("D&MK analysis...", ctrl Or Keys.D, AddressOf OnDmkInfo)
        _miCpmFormat = MakeItem("CP/M &layout...", Keys.None, AddressOf OnCpmFormat)
        _miGranuleMap = MakeItem("&Granule map...", Keys.None, AddressOf OnGranuleMap)
        _miCheckRepair = MakeItem("&Check && repair...", Keys.None, AddressOf OnCheckRepair)
        _miFormat = MakeItem("&Format...", Keys.None, AddressOf OnFormat)
        _miTrim = MakeItem("&Trim...", Keys.None, AddressOf OnTrim)
        _miBoot = MakeItem("&Boot in the emulator...", Keys.None, AddressOf OnBoot)
        _miShell = MakeItem("&Open trs80-tool shell here", Keys.None, AddressOf OnShell)
        mDisk.Items.AddRange(New MenuItem() {_miSectorMap, _miDmkInfo, _miGranuleMap, _miCheckRepair, _miCpmFormat, New SeparatorMenuItem(), _miFormat, _miTrim, New SeparatorMenuItem(), _miBoot, _miShell})

        Dim mTools As New ButtonMenuItem With {.Text = "&Tools"}
        mTools.Items.Add(MakeItem("&Convert file(s)...", Keys.None, Sub() OnConvertFiles(Nothing)))
        mTools.Items.Add(MakeItem("&Assemble...", Keys.None, AddressOf OnAssemble))
        mTools.Items.Add(New SeparatorMenuItem())
        mTools.Items.Add(MakeItem("&Hex dump of a file...", Keys.None, AddressOf OnHexdumpFile))
        mTools.Items.Add(MakeItem("&Disassemble a file...", Keys.None, AddressOf OnDisasmFile))
        mTools.Items.Add(New SeparatorMenuItem())
        _miRunEmpty = MakeItem("&Run the emulator without a disk...", Keys.None, AddressOf OnRunEmulatorEmpty)
        mTools.Items.Add(_miRunEmpty)

        Dim mHelp As New ButtonMenuItem With {.Text = "&Help"}
        mHelp.Items.Add(MakeItem("&Manual...", Keys.F1, Sub() ShowHelpText("Manual.txt", Program.AppName & " Manual")))
        mHelp.Items.Add(MakeItem("&Changelog...", Keys.None, Sub() ShowHelpText("Changelog.txt", Program.AppName & " Changelog")))
        mHelp.Items.Add(MakeItem("&License...", Keys.None, Sub() ShowHelpText("License.txt", Program.AppName & " License")))
        mHelp.Items.Add(New SeparatorMenuItem())
        mHelp.Items.Add(MakeItem("&Engine capabilities...", Keys.None, AddressOf OnEngineInfo))
        mHelp.Items.Add(MakeItem("trs80-tool manual (&web)", Keys.None, Sub() OpenUrl(AppInfo.ToolUrl)))
        If Not HostPlatform.IsMacOS Then
            mHelp.Items.Add(New SeparatorMenuItem())
            mHelp.Items.Add(MakeItem("&About " & Program.AppName & "...", Keys.None, AddressOf OnAbout))
        End If

        Dim bar As MenuBar
        If HostPlatform.IsMacOS Then
            bar = BuildMacMenuBar(ctrl, mHelp)
        Else
            bar = New MenuBar With {.IncludeSystemItems = MenuBarSystemItems.None}
        End If
        bar.Items.AddRange(New MenuItem() {mFile, mEdit, mView, mDisk, mTools, mHelp})
        Menu = bar

        ' ---- toolbar (text buttons: no bitmaps to maintain)
        Dim tb As New ToolBar With {.TextAlign = ToolBarTextAlign.Right}
        tb.Items.Add(MakeButton("New", AddressOf OnFileNew))
        tb.Items.Add(MakeButton("Open", AddressOf OnFileOpen))
        _tbProps = MakeButton("Disk properties", AddressOf OnDiskProperties)
        tb.Items.Add(_tbProps)
        tb.Items.Add(New SeparatorToolItem())
        _tbCopy = MakeButton("Copy", AddressOf OnEditCopy)
        _tbPaste = MakeButton("Paste", AddressOf OnEditPaste)
        tb.Items.Add(_tbCopy)
        tb.Items.Add(_tbPaste)
        tb.Items.Add(New SeparatorToolItem())
        _tbAdd = MakeButton("Add", AddressOf OnEditAdd)
        _tbExtract = MakeButton("Extract", AddressOf OnEditExtract)
        _tbView = MakeButton("View", AddressOf OnPreview)
        tb.Items.Add(_tbAdd)
        tb.Items.Add(_tbExtract)
        tb.Items.Add(_tbView)
        tb.Items.Add(New SeparatorToolItem())
        _tbSectors = MakeButton("Sectors", AddressOf OnSectorMap)
        tb.Items.Add(_tbSectors)
        tb.Items.Add(New SeparatorToolItem())
        tb.Items.Add(MakeButton("About", AddressOf OnAbout))
        ToolBar = tb

        ' ---- file list
        _grid.AllowMultipleSelection = True
        _grid.AllowDrop = True
        _grid.GridLines = GridLines.None
        AddColumn("Name", 130, Function(f) f.BaseName, False)
        AddColumn("Type", 60, Function(f) f.Extension, False)
        AddColumn("Size", 90, Function(f) TextUtil.FormatBytes(f.Size), True)
        AddColumn("Date", 150, Function(f) f.DateDisplay, False)
        AddColumn("Attributes", 120, AddressOf AttributeText, False)
        AddColumn("Kind", 320, AddressOf KindText, False)
        If s.ColumnWidths IsNot Nothing AndAlso s.ColumnWidths.Count = _grid.Columns.Count Then
            For i As Integer = 0 To _grid.Columns.Count - 1
                If s.ColumnWidths(i) > 10 Then _grid.Columns(i).Width = s.ColumnWidths(i)
            Next
        End If
        AddHandler _grid.ColumnHeaderClick, AddressOf OnColumnClick
        AddHandler _grid.CellDoubleClick, Sub() OnPreview()
        AddHandler _grid.KeyDown, AddressOf OnListKeyDown
        AddHandler _grid.SelectionChanged, Sub() UpdateEnabled()
        AddHandler _grid.CellFormatting, AddressOf OnCellFormatting
        AddHandler _grid.MouseDown, AddressOf OnListMouseDown
        AddHandler _grid.DragEnter, AddressOf OnDragOverList
        AddHandler _grid.DragOver, AddressOf OnDragOverList
        AddHandler _grid.DragDrop, AddressOf OnDragDropList
        If HostPlatform.IsMacOS Then MapMacEditCommands()

        ' ---- context menu
        _cxPreview = MakeItem("&Preview", Keys.None, AddressOf OnPreview)
        _cxExtract = MakeItem("&Extract...", Keys.None, AddressOf OnEditExtract)
        _cxDelete = MakeItem("&Delete", Keys.None, AddressOf OnDelete)
        _cxRename = MakeItem("&Rename...", Keys.None, AddressOf OnRename)
        _cxProps = MakeItem("P&roperties...", Keys.None, AddressOf OnFileProperties)
        _cxConvert = MakeItem("Con&vert to...", Keys.None, AddressOf OnConvertDiskFile)
        Dim ctx As New ContextMenu()
        ctx.Items.AddRange(New MenuItem() {_cxPreview, New SeparatorMenuItem(), _cxExtract, _cxDelete, _cxRename, _cxConvert, New SeparatorMenuItem(), _cxProps})
        _grid.ContextMenu = ctx

        ' ---- status line
        _lblStatus.Text = "Ready"
        _lblEngine.Text = ""
        Dim statusRow As New TableLayout With {.Padding = New Padding(6, 3), .Spacing = New Size(12, 0)}
        statusRow.Rows.Add(New TableRow(New TableCell(_lblStatus, True), New TableCell(_lblEngine)))

        Dim root As New TableLayout()
        root.Rows.Add(New TableRow(_grid) With {.ScaleHeight = True})
        root.Rows.Add(New TableRow(statusRow))
        Content = root

        UpdateViewChecks()
        RebuildRecentMenu()
        UpdateEnabled()
    End Sub

    ''' <summary>
    ''' The menu bar of the Mac program. A Mac has one menu bar for the whole program, and it is expected to
    ''' hold: an application menu with About, Preferences and Quit; the system's own Edit items, which are what
    ''' make Cmd+C, Cmd+V and Cmd+A work in every text box; and a Window menu. Eto adds the system items; the
    ''' three of our own go where a Mac user looks for them.
    ''' </summary>
    Private Function BuildMacMenuBar(cmd As Keys, helpMenu As ButtonMenuItem) As MenuBar
        Dim bar As New MenuBar With {.IncludeSystemItems = MenuBarSystemItems.All}
        bar.AboutItem = MakeItem("About " & Program.AppName, Keys.None, AddressOf OnAbout)
        bar.ApplicationItems.Add(MakeItem("Preferences...", cmd Or Keys.Comma, AddressOf OnPreferences))
        bar.QuitItem = MakeItem("Quit " & Program.AppName, cmd Or Keys.Q, Sub() Program.ExitProgram(Me))
        ' Help is the last menu on a Mac, after the Window menu the system adds.
        helpMenu.Order = 1000
        Return bar
    End Function

    ''' <summary>
    ''' Mac only. Edit > Copy, Paste and Delete are the system's items there and go to whatever has the focus.
    ''' When that is the file list they run the same commands the Copy and Paste buttons run.
    ''' </summary>
    Private Sub MapMacEditCommands()
        _cmdCopy = New Command(Sub(sender, e) OnEditCopy())
        _cmdPaste = New Command(Sub(sender, e) OnEditPaste())
        _cmdDelete = New Command(Sub(sender, e) OnDelete())
        Try
            _grid.MapPlatformCommand("copy", _cmdCopy)
            _grid.MapPlatformCommand("paste", _cmdPaste)
            _grid.MapPlatformCommand("delete", _cmdDelete)
        Catch
            ' Without the mapping the toolbar buttons and the context menu still do all three.
        End Try
    End Sub

    ''' <summary>
    ''' Text that goes into a menu item as it is (a file path, a partition name). An ampersand marks the
    ''' shortcut letter in menu text, so one that belongs to the name is doubled. Nothing else needs care:
    ''' Eto itself looks after the underscore, which is GTK's mark.
    ''' </summary>
    Private Shared Function MenuLiteral(text As String) As String
        Return text.Replace("&", "&&")
    End Function

    Private Sub AddColumn(header As String, width As Integer, getter As Func(Of DiskFile, String), rightAligned As Boolean)
        Dim cell As New TextBoxCell With {.Binding = Binding.Delegate(Of DiskFile, String)(getter)}
        If rightAligned Then cell.TextAlignment = TextAlignment.Right
        _grid.Columns.Add(New GridColumn With {.HeaderText = header, .DataCell = cell, .Width = width, .Resizable = True, .Sortable = True, .AutoSize = False, .Editable = False})
    End Sub

    Private Shared Function MakeItem(text As String, shortcut As Keys, handler As Action) As ButtonMenuItem
        Dim mi As New ButtonMenuItem With {.Text = text}
        If shortcut <> Keys.None Then mi.Shortcut = shortcut
        AddHandler mi.Click, Sub() handler()
        Return mi
    End Function

    Private Shared Function MakeButton(text As String, handler As Action) As ButtonToolItem
        Dim b As New ButtonToolItem With {.Text = text}
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
        If WindowState = WindowState.Normal AndAlso ClientSize.Height > 50 Then
            _sizeDelta = New Size(_requestedSize.Width - ClientSize.Width, _requestedSize.Height - ClientSize.Height)
            _sizeKnown = True
        End If
        Try
            Await CheckEngineAsync(True)
            If _startFile IsNot Nothing Then Await OpenImageAsync(_startFile)
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, Program.AppName)
        End Try
    End Sub

    Protected Overrides Sub OnClosing(e As CancelEventArgs)
        Dim s As Settings = Program.Settings
        s.WindowMaximized = (WindowState = WindowState.Maximized)
        If WindowState = WindowState.Normal Then
            Dim inner As Size = ClientSize
            If _sizeKnown AndAlso inner.Width > 100 AndAlso inner.Height > 100 Then
                s.WindowWidth = inner.Width + _sizeDelta.Width : s.WindowHeight = inner.Height + _sizeDelta.Height
            End If
            s.WindowX = Location.X : s.WindowY = Location.Y
        End If
        s.ColumnWidths = New List(Of Integer)
        For Each c As GridColumn In _grid.Columns
            s.ColumnWidths.Add(c.Width)
        Next
        s.Save()
        MyBase.OnClosing(e)
    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.Key = Keys.Escape AndAlso e.Modifiers = Keys.None AndAlso Program.Settings.ExitOnEscape AndAlso Not _busy Then
            e.Handled = True
            Program.ExitProgram(Me)
            Return
        End If
        MyBase.OnKeyDown(e)
    End Sub

    ''' <summary>Run the engine probe; show problems. With askIfMissing, offer Preferences when the tool can't be found.</summary>
    Private Async Function CheckEngineAsync(askIfMissing As Boolean) As Task(Of Boolean)
        Dim tool As Trs80Tool = Program.Tool
        If String.IsNullOrEmpty(tool.ToolPath) OrElse Not File.Exists(tool.ToolPath) Then
            _lblEngine.Text = "trs80-tool: not found"
            If askIfMissing Then
                ' On a Mac this is also the first question of a first start (Program.Main can't ask there).
                Dim macWhere As String = If(HostPlatform.IsMacOS, Program.ToolFoundWhereHint, "")
                If Ui.Ask(Me, If(HostPlatform.IsMacOS, "trs80-tool was not found.", "trs80-tool was not found at the configured location.") & vbLf & vbLf &
                    Program.ToolDownloadHint & ", then tell this program where it is." & macWhere & vbLf & vbLf &
                    "Open Preferences now to set its location?", Program.AppName, MessageBoxType.Warning) Then OnPreferences()
            End If
            Return False
        End If
        SetBusy(True, "Checking trs80-tool...")
        Dim caps As EngineCapabilities
        Try
            caps = Await tool.ProbeAsync()
        Finally
            SetBusy(False, Nothing)
        End Try
        If caps.Version Is Nothing Then
            _lblEngine.Text = "trs80-tool: unusable"
            Dim macHint As String = Program.MacToolHint(tool.ToolPath)
            Ui.Fail(Me, "trs80-tool could not be run or its version could not be read:" & vbLf & vbLf & String.Join(vbLf, caps.Problems) & vbLf & vbLf & "Path: " & tool.ToolPath &
                    If(macHint.Length > 0, vbLf & vbLf & macHint, ""), Program.AppName)
            Return False
        End If
        _lblEngine.Text = "trs80-tool " & caps.Version.ToString()
        If caps.Version < Trs80Tool.MinimumVersion Then
            Ui.Fail(Me, "trs80-tool " & caps.Version.ToString() & " is too old. This program needs version " & Trs80Tool.MinimumVersion.ToString() & " or later, which answers every disk question as data (""disk --json"", ""help --json"")." & vbLf & vbLf & "Please download the current version from " & AppInfo.ToolUrl & " and put it where the old one was:" & vbLf & tool.ToolPath, Program.AppName)
            Return False
        End If
        If caps.Version > Trs80Tool.MaxTestedVersion AndAlso Program.Settings.LastNotedVersion <> caps.Version.ToString() Then
            Program.Settings.LastNotedVersion = caps.Version.ToString()
            Program.Settings.Save()
            Dim missing As New List(Of String)
            For Each p In caps.Problems
                If p.Contains(" missing") Then missing.Add(p)
            Next
            Ui.Info(Me, "trs80-tool " & caps.Version.ToString() & " is newer than the version this program was tested with (" & Trs80Tool.MaxTestedVersion.ToString() & ")." & vbLf & vbLf &
                    If(missing.Count = 0, "It has every command this program uses, so it should work normally.", "It no longer has some commands this program uses:" & vbLf & String.Join(vbLf, missing) & vbLf & vbLf & "Those features are disabled until this program is updated.") & vbLf & vbLf &
                    "Details: Help → Engine capabilities.", Program.AppName)
        End If
        Return True
    End Function

    ' =================================================================== state helpers

    Private Sub SetBusy(busy As Boolean, statusText As String)
        _busy = busy
        If statusText IsNot Nothing Then _lblStatus.Text = statusText
        If Not busy Then UpdateStatus()
    End Sub

    Private Sub UpdateStatus()
        If _disk Is Nothing Then
            _lblStatus.Text = "Ready"
            Return
        End If
        Dim parts As New List(Of String)
        If _disk.DosName.Length > 0 Then parts.Add(_disk.DosName) Else parts.Add(_disk.Description)
        parts.Add(_rows.Count.ToString(CultureInfo.InvariantCulture) & " file(s) shown of " & _disk.VisibleFiles(True, True).Count.ToString(CultureInfo.InvariantCulture))
        If _disk.FreeText.Length > 0 Then parts.Add(_disk.FreeText)
        If _disk.WriteProtected.HasValue AndAlso _disk.WriteProtected.Value Then parts.Add("WRITE-PROTECTED")
        If _disk.Damaged Then parts.Add("DAMAGED")
        If _disk.ProtectedSignCount > 0 Then parts.Add("signs of copy protection")
        If Not _disk.IsWritable Then parts.Add("read-only: " & If(_disk.ReadOnlyReason, ""))
        If Program.Tool.CpmFormatOf(_disk.Path).Length > 0 Then parts.Add("read as CP/M layout " & Program.Tool.CpmFormatOf(_disk.Path))
        If ConsoleRunning() Then parts.Add(_consoleWhat & " is open on this image - changes here wait until it closes")
        _lblStatus.Text = String.Join("   |   ", parts)
    End Sub

    Private Sub UpdateTitle()
        If _disk Is Nothing Then
            Title = Program.AppName
            Return
        End If
        Dim t As String = Program.AppName & " - " & Path.GetFileName(_disk.Path) & " - (" & If(_disk.DosName.Length > 0, _disk.DosName, _disk.Description)
        If _disk.VolumeName.Length > 0 Then t &= " - " & _disk.VolumeName
        If _disk.IsHardDisk AndAlso _disk.ActivePartition IsNot Nothing Then t &= " - partition " & _disk.ActivePartition.Number.ToString(CultureInfo.InvariantCulture)
        Title = t & ")"
    End Sub

    Private Sub UpdateViewChecks()
        _settingChecks = True
        Try
            _miInvisible.Checked = Program.Settings.ShowInvisible
            _miSystem.Checked = Program.Settings.ShowSystem
        Finally
            _settingChecks = False
        End Try
    End Sub

    Private Function SelectedCount() As Integer
        Return _grid.SelectedRows.Count()
    End Function

    Private Sub UpdateEnabled()
        Dim isOpen As Boolean = _disk IsNot Nothing
        Dim writable As Boolean = isOpen AndAlso _disk.IsWritable AndAlso Not ConsoleRunning()
        Dim selCount As Integer = SelectedCount()
        Dim hasDisk As Boolean = EngineHasDisk("add")
        Dim floppy As Boolean = isOpen AndAlso Not _disk.IsHardDisk AndAlso Not _disk.IsCassette
        _miClose.Enabled = isOpen
        _miDiskProps.Enabled = isOpen : _tbProps.Enabled = isOpen
        _miCopy.Enabled = isOpen AndAlso selCount > 0 : _tbCopy.Enabled = _miCopy.Enabled
        ' Asking the clipboard what it holds stops the window while GTK waits for the answer, so Paste is
        ' offered whenever the disk can be written and looks at the clipboard when it is chosen.
        _miPaste.Enabled = writable AndAlso hasDisk : _tbPaste.Enabled = _miPaste.Enabled
        _miAdd.Enabled = writable AndAlso hasDisk : _tbAdd.Enabled = _miAdd.Enabled
        _miExtract.Enabled = isOpen AndAlso _rows.Count > 0 : _tbExtract.Enabled = _miExtract.Enabled
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
        _cxExtract.Enabled = isOpen AndAlso _rows.Count > 0
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
        _miBoot.Enabled = floppy AndAlso Not ConsoleRunning() AndAlso EngineHasTop("run")
        _miShell.Enabled = isOpen AndAlso Not _disk.IsCassette AndAlso Not ConsoleRunning()
        _miRunEmpty.Enabled = Not ConsoleRunning() AndAlso EngineHasTop("run")
        If _cmdCopy IsNot Nothing Then
            _cmdCopy.Enabled = _miCopy.Enabled
            _cmdPaste.Enabled = _miPaste.Enabled
            _cmdDelete.Enabled = _cxDelete.Enabled
        End If
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
        _miRecent.Items.Clear()
        If Program.Settings.Mru.Count = 0 Then
            _miRecent.Items.Add(New ButtonMenuItem With {.Text = "(none)", .Enabled = False})
            Return
        End If
        Dim n As Integer = 1
        For Each p In Program.Settings.Mru
            Dim pathCopy As String = p
            Dim mi As New ButtonMenuItem With {.Text = "&" & n.ToString(CultureInfo.InvariantCulture) & " " & MenuLiteral(p)}
            AddHandler mi.Click, Async Sub()
                                     If _busy Then Return
                                     Await OpenImageAsync(pathCopy)
                                 End Sub
            _miRecent.Items.Add(mi)
            n += 1
        Next
    End Sub

    Private Sub RebuildPartitionMenu()
        _miPartition.Items.Clear()
        If _disk Is Nothing OrElse Not _disk.IsHardDisk Then Return
        For Each p In _disk.Partitions
            Dim part As DiskPartition = p
            Dim mi As New ButtonMenuItem With {.Text = If(part Is _disk.ActivePartition, "● ", "    ") & MenuLiteral(part.Display)}
            AddHandler mi.Click, Async Sub()
                                     If _busy Then Return
                                     Await SelectPartitionAsync(part)
                                 End Sub
            _miPartition.Items.Add(mi)
        Next
        ' trs80-tool 5.1.0 makes and erases partitions ("disk partition --add / --delete").
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        If caps IsNot Nothing AndAlso caps.HasDiskOption("partition", "--add") Then
            If _disk.Partitions.Count > 0 Then _miPartition.Items.Add(New SeparatorMenuItem())
            Dim miAdd As New ButtonMenuItem With {.Text = "&Add a partition...", .Enabled = Not ConsoleRunning()}
            AddHandler miAdd.Click, Sub() OnPartitionAdd()
            _miPartition.Items.Add(miAdd)
            Dim miDel As New ButtonMenuItem With {.Text = "&Delete this partition...", .Enabled = _disk.ActivePartition IsNot Nothing AndAlso Not ConsoleRunning()}
            AddHandler miDel.Click, Sub() OnPartitionDelete()
            _miPartition.Items.Add(miDel)
        End If
    End Sub

    ''' <summary>A window with text the tool printed that the user needs to keep (driver settings of a new hard disk, say).</summary>
    Private Sub ShowToolText(title As String, text As String)
        Dim body As String = Ui.Lf(text).Trim()
        If body.Length = 0 Then Return
        Dim f As New TextReportForm(title, body)
        f.ShowOwned(Me)
    End Sub

    ''' <summary>File → New hard disk image: "disk NEW.hdv create [DOS] --cylinders N --heads N [--partitions N] ...".</summary>
    Private Async Sub OnFileNewHardDisk()
        If _busy Then Return
        Dim dlg As New CreateHardDiskForm()
        If Not dlg.ShowOk(Me) Then Return
        Dim target As String = Ui.PickSave(Me, "Save new hard disk image as", "HDV hard disk image (*.hdv)|*.hdv", Nothing, "hdv")
        If target Is Nothing Then Return
        If Not target.EndsWith(".hdv", StringComparison.OrdinalIgnoreCase) Then target &= ".hdv"
        SetBusy(True, "Creating " & Path.GetFileName(target) & "...")
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
            Ui.Fail(Me, failure, "Disk Image Creation Error")
            Return
        End If
        Await OpenImageAsync(target)
        ' An HDV file has no partition table: the tool prints what the DOS's hard disk driver must be told.
        ShowToolText("New hard disk image - " & Path.GetFileName(target), r.StdOut)
    End Sub

    ''' <summary>View → Partition → Add a partition: "disk IMG partition --add DOS ...".</summary>
    Private Async Sub OnPartitionAdd()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsHardDisk OrElse ConsoleRunning() Then Return
        Dim dlg As New PartitionAddForm()
        If Not dlg.ShowOk(Me) Then Return
        SetBusy(True, "Adding a partition...")
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
            Ui.Fail(Me, failure, "Add a partition")
            Return
        End If
        Await RefreshAsync()
        ShowToolText("New partition - " & Path.GetFileName(_disk.Path), r.StdOut)
    End Sub

    ''' <summary>View → Partition → Delete this partition: "disk IMG partition --delete N [--force]".</summary>
    Private Async Sub OnPartitionDelete()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsHardDisk OrElse _disk.ActivePartition Is Nothing OrElse ConsoleRunning() Then Return
        Dim part As DiskPartition = _disk.ActivePartition
        Dim fileCount As Integer = _disk.VisibleFiles(True, False).Count
        Dim msg As String = "Delete partition " & part.Number.ToString(CultureInfo.InvariantCulture) & If(part.DiskName.Length > 0, " (" & part.DiskName & ")", "") & " of " & Path.GetFileName(_disk.Path) & "?" & vbLf & vbLf &
            If(fileCount = 0, "It has no files.", "Its " & fileCount.ToString(CultureInfo.InvariantCulture) & " file(s) will be erased.") & " The partitions after it get numbers one lower." & vbLf & vbLf & "This cannot be undone."
        If Not Ui.Ask(Me, msg, "Delete partition", MessageBoxType.Warning, True) Then Return
        SetBusy(True, "Deleting the partition...")
        Dim r As ToolResult = Nothing
        Dim failure As String = Nothing
        Try
            r = Await Program.Tool.DiskPartitionDeleteAsync(_disk.Path, part.Number, fileCount > 0)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure Is Nothing AndAlso Not r.Succeeded Then failure = r.ErrorText
        If failure IsNot Nothing Then Ui.Fail(Me, failure, "Delete partition")
        Await RefreshAsync()
    End Sub

    ' =================================================================== opening / listing

    Private Async Function OpenImageAsync(imagePath As String) As Task
        If Not File.Exists(imagePath) Then
            Ui.Fail(Me, "File not found:" & vbLf & imagePath, "Error Opening Virtual Disk")
            Program.Settings.RemoveMru(imagePath)
            RebuildRecentMenu()
            Return
        End If
        If Program.Tool.Capabilities Is Nothing OrElse Not Program.Tool.Capabilities.IsUsable Then
            If Not Await CheckEngineAsync(True) Then Return
        End If
        If HostFiles.IsGzipName(imagePath) AndAlso Not Program.Tool.Has520 Then
            Ui.Info(Me, Path.GetFileName(imagePath) & " is a gzipped image. Opening it needs trs80-tool 5.2.0 or later; this one is " & Program.Tool.Capabilities.Version.ToString() & "." & vbLf & vbLf & "Unzip the file, or get the current trs80-tool from " & AppInfo.ToolUrl & ".", "Error Opening Virtual Disk")
            Return
        End If
        SetBusy(True, "Reading " & Path.GetFileName(imagePath) & "...")
        Dim problem As String = Nothing
        Try
            Dim keepPart As Integer = If(_disk IsNot Nothing AndAlso String.Equals(_disk.Path, imagePath, HostPlatform.PathComparison) AndAlso _disk.ActivePartition IsNot Nothing, _disk.ActivePartition.Number, 0)
            _disk = Await DiskImage.LoadAsync(Program.Tool, Path.GetFullPath(imagePath), keepPart)
            Program.Settings.AddMru(_disk.Path)
            RebuildRecentMenu()
            RebuildPartitionMenu()
            FillList(Nothing)
            UpdateTitle()
            _lblStatus.ToolTip = If(_disk.Warnings.Count > 0, String.Join(vbLf, _disk.Warnings), "")
        Catch ex As Exception
            _disk = Nothing
            _rows = New List(Of DiskFile)
            _grid.DataStore = _rows
            UpdateTitle()
            problem = ex.Message
        End Try
        SetBusy(False, Nothing)
        UpdateEnabled()
        If problem IsNot Nothing Then Ui.Fail(Me, problem, "Error Opening Virtual Disk")
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
        SetBusy(True, "Reading partition " & part.Number.ToString(CultureInfo.InvariantCulture) & "...")
        Dim problem As String = Nothing
        Try
            Await _disk.RefreshStatusAsync(Program.Tool)
        Catch ex As Exception
            problem = ex.Message
        End Try
        SetBusy(False, Nothing)
        If problem IsNot Nothing Then Ui.Fail(Me, problem, Program.AppName)
        RebuildPartitionMenu()
        FillList(Nothing)
        UpdateTitle()
        UpdateEnabled()
    End Function

    Private Shared Function AttributeText(f As DiskFile) As String
        Dim attrs As String = f.AttributeLetters
        If f.Protection.Length > 0 AndAlso Not String.Equals(f.Protection, "FULL", StringComparison.OrdinalIgnoreCase) Then attrs &= If(attrs.Length > 0, " ", "") & f.Protection
        If f.CpmUser.HasValue AndAlso f.CpmUser.Value <> 0 Then attrs &= " user " & f.CpmUser.Value.ToString(CultureInfo.InvariantCulture)
        Return attrs
    End Function

    Private Shared Function KindText(f As DiskFile) As String
        Dim kind As String = f.Description
        If f.Damage.Length > 0 Then kind &= "  [DAMAGED: " & f.Damage & "]"
        If f.DecodeError.Length > 0 Then kind &= "  (" & f.DecodeError & ")"
        Return kind
    End Function

    ''' <summary>TRSTool's colours: invisible+system red, system orange, invisible olive.</summary>
    Private Sub OnCellFormatting(sender As Object, e As GridCellFormatEventArgs)
        Dim f As DiskFile = TryCast(e.Item, DiskFile)
        If f Is Nothing Then Return
        If f.Hidden AndAlso f.IsSystem Then
            e.ForegroundColor = Colors.Red
        ElseIf f.IsSystem Then
            e.ForegroundColor = Colors.DarkOrange
        ElseIf f.Hidden Then
            e.ForegroundColor = Colors.Olive
        End If
    End Sub

    ''' <summary>Fill the list from the model, applying the Invisible/System filters and the sort order.</summary>
    Private Sub FillList(reselect As List(Of String))
        Dim rows As New List(Of DiskFile)
        If _disk IsNot Nothing Then rows.AddRange(_disk.VisibleFiles(Program.Settings.ShowInvisible, Program.Settings.ShowSystem))
        rows.Sort(AddressOf CompareRows)
        _rows = rows
        _grid.DataStore = _rows
        If reselect IsNot Nothing Then ReselectNames(reselect)
        UpdateStatus()
        UpdateEnabled()
    End Sub

    ''' <summary>Name/Type/Attributes/Kind by text, Size numerically, Date chronologically; ties by name.</summary>
    Private Function CompareRows(a As DiskFile, b As DiskFile) As Integer
        Dim r As Integer
        Select Case _sortColumn
            Case 0 : r = String.Compare(a.BaseName, b.BaseName, StringComparison.OrdinalIgnoreCase)
            Case 1 : r = String.Compare(a.Extension, b.Extension, StringComparison.OrdinalIgnoreCase)
            Case 2 : r = a.Size.CompareTo(b.Size)
            Case 3
                Dim da As DateTime = If(a.DateValue.HasValue, a.DateValue.Value, DateTime.MinValue)
                Dim db As DateTime = If(b.DateValue.HasValue, b.DateValue.Value, DateTime.MinValue)
                r = da.CompareTo(db)
            Case 4 : r = String.Compare(AttributeText(a), AttributeText(b), StringComparison.OrdinalIgnoreCase)
            Case Else : r = String.Compare(KindText(a), KindText(b), StringComparison.OrdinalIgnoreCase)
        End Select
        If r = 0 Then r = String.Compare(a.TrsName, b.TrsName, StringComparison.OrdinalIgnoreCase)
        Return If(_sortAscending, r, -r)
    End Function

    Private Function SelectedFiles() As List(Of DiskFile)
        Dim l As New List(Of DiskFile)
        For Each i In _grid.SelectedRows.OrderBy(Function(x) x)
            If i >= 0 AndAlso i < _rows.Count Then l.Add(_rows(i))
        Next
        Return l
    End Function

    Private Function AllListedFiles() As List(Of DiskFile)
        Return New List(Of DiskFile)(_rows)
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
        _grid.UnselectAll()
        Dim first As Integer = -1
        For i As Integer = 0 To _rows.Count - 1
            If names.Contains(_rows(i).TrsName, StringComparer.OrdinalIgnoreCase) Then
                _grid.SelectRow(i)
                If first < 0 Then first = i
            End If
        Next
        If first >= 0 Then
            Try : _grid.ScrollToRow(first) : Catch : End Try
        End If
    End Sub

    Private Sub SelectAllRows()
        If _rows.Count > 0 Then _grid.SelectAll()
    End Sub

    Private Sub InvertSelection()
        Dim was As New HashSet(Of Integer)(_grid.SelectedRows)
        _grid.UnselectAll()
        For i As Integer = 0 To _rows.Count - 1
            If Not was.Contains(i) Then _grid.SelectRow(i)
        Next
    End Sub

    Private Sub OnColumnClick(sender As Object, e As GridColumnEventArgs)
        Dim index As Integer = _grid.Columns.IndexOf(e.Column)
        If index < 0 Then Return
        If index = _sortColumn Then
            _sortAscending = Not _sortAscending
        Else
            _sortColumn = index
            _sortAscending = True
        End If
        FillList(SelectedNames())
    End Sub

    Private Sub OnListKeyDown(sender As Object, e As KeyEventArgs)
        If _busy Then Return
        If e.Key = Keys.Enter AndAlso e.Modifiers = Keys.None AndAlso SelectedCount() >= 1 Then
            e.Handled = True
            OnPreview()
        ElseIf (e.Key = Keys.Delete OrElse (HostPlatform.IsMacOS AndAlso e.Key = Keys.Backspace)) AndAlso e.Modifiers = Keys.None AndAlso SelectedCount() > 0 Then
            ' The key marked "delete" on a Mac keyboard is Backspace.
            e.Handled = True
            OnDelete()
        End If
    End Sub

    ''' <summary>A right click on a file that isn't selected selects it, so the context menu acts on the file under the pointer.</summary>
    Private Sub OnListMouseDown(sender As Object, e As MouseEventArgs)
        If e.Buttons <> MouseButtons.Alternate OrElse _disk Is Nothing Then Return
        Try
            Dim cell As GridCell = _grid.GetCellAt(e.Location)
            If cell IsNot Nothing AndAlso cell.RowIndex >= 0 AndAlso Not _grid.SelectedRows.Contains(cell.RowIndex) Then
                _grid.UnselectAll()
                _grid.SelectRow(cell.RowIndex)
            End If
        Catch
        End Try
    End Sub

    ' =================================================================== File menu

    Private Async Sub OnFileOpen()
        If _busy Then Return
        ' trs80-tool 5.2.0 opens gzipped images (game.dmk.gz, big.hdv.gz) like any other.
        Const gzImages As String = "*.dsk.gz;*.dmk.gz;*.jv1.gz;*.jv3.gz;*.hfe.gz;*.imd.gz;*.scp.gz;*.hdv.gz"
        Dim filter As String = "Disk images (*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.hdv;*.gz)|*.dsk;*.dmk;*.jv1;*.jv3;*.hfe;*.imd;*.scp;*.hdv;*.hdv1;*.hdv2;" & gzImages & "|" &
                     "Cassettes (*.cas;*.wav)|*.cas;*.wav;*.cas.gz;*.wav.gz|DMK (*.dmk)|*.dmk;*.dmk.gz|JV1 (*.jv1)|*.jv1;*.jv1.gz|JV3 (*.jv3)|*.jv3;*.jv3.gz|HDV hard disks (*.hdv)|*.hdv;*.hdv1;*.hdv2;*.hdv.gz|" &
                     "Gzipped images (*.gz)|" & gzImages & ";*.cas.gz;*.wav.gz|All files (*.*)|*.*"
        Dim chosen As String = Ui.PickFile(Me, "Open virtual disk", filter)
        If chosen IsNot Nothing Then Await OpenImageAsync(chosen)
    End Sub

    Private Sub OnFileClose()
        If _busy Then Return
        _disk = Nothing
        _rows = New List(Of DiskFile)
        _grid.DataStore = _rows
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
        Dim dlg As New CreateFloppyForm()
        If Not dlg.ShowOk(Me) Then Return
        Dim target As String = Ui.PickSave(Me, "Save new floppy image as", dlg.ImageTypeFilter & "|All files (*.*)|*.*", Nothing, dlg.ImageTypeExtension)
        If target Is Nothing Then Return
        SetBusy(True, "Creating " & Path.GetFileName(target) & "...")
        Dim failure As String = Nothing
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
                        If Not Ui.Ask(Me, "trs80-tool says the new disk doesn't fit the JV1 format exactly:" & vbLf & vbLf & convertError.Message & vbLf & vbLf & "Write it as JV1 anyway (--force)?", "Create New Floppy Image", MessageBoxType.Warning) Then Throw convertError
                        Await Program.Tool.ConvertAsync(tmpDmk, tmpOut, True)
                    End If
                    File.Copy(tmpOut, target, True)
                End If
                Try : Directory.Delete(tmpDir, True) : Catch : End Try
            End If
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            Ui.Fail(Me, failure, "Disk Image Creation Error")
            Return
        End If
        Await OpenImageAsync(target)
    End Sub

    Private Async Sub OnDiskProperties()
        If _busy OrElse _disk Is Nothing Then Return
        Dim dlg As New DiskPropertiesForm(_disk)
        If Not dlg.ShowOk(Me) Then Return
        Dim words As List(Of String) = dlg.LabelOptionWords()
        Dim changed As Boolean = False
        Dim failure As String = Nothing
        SetBusy(True, "Updating disk...")
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
            Ui.Fail(Me, failure, "Disk Properties Error")
            Await RefreshAsync()
            Return
        End If
        If changed Then Await RefreshAsync()
    End Sub

    ''' <summary>File → Save image as: trs80-tool "convert IMG OUT" to DMK/JV1/JV3/HFE/IMD, then open the copy.</summary>
    Private Async Sub OnSaveImageAs()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk OrElse _disk.IsCassette Then Return
        Dim startName As String = Path.Combine(Path.GetDirectoryName(_disk.Path), Path.GetFileNameWithoutExtension(HostFiles.WithoutGzipSuffix(_disk.Path)))
        Dim target As String = Ui.PickSave(Me, "Save image as (trs80-tool convert)", "DMK floppy image (*.dmk)|*.dmk|JV3 floppy image (*.jv3)|*.jv3|JV1 floppy image (*.jv1)|*.jv1|HFE floppy image (*.hfe)|*.hfe|IMD floppy image (*.imd)|*.imd", startName)
        If target Is Nothing Then Return
        Dim ext As String = Path.GetExtension(target).ToLowerInvariant()
        If ext = ".dsk" OrElse ext.Length = 0 Then
            Ui.Warn(Me, "End the name with .dmk, .jv3, .jv1, .hfe or .imd. trs80-tool needs the extension to know which format to write (.dsk doesn't say).", "Save image as")
            Return
        End If
        If String.Equals(Path.GetFullPath(target), _disk.Path, HostPlatform.PathComparison) Then
            Ui.Warn(Me, "That is the open image itself; choose another name.", "Save image as")
            Return
        End If
        SetBusy(True, "Converting to " & Path.GetFileName(target) & "...")
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
            If Not Ui.Ask(Me, "trs80-tool says the image doesn't fit this format exactly:" & vbLf & vbLf & needsForce.Message & vbLf & vbLf & "Write it anyway (--force)?", "Save image as", MessageBoxType.Warning) Then Return
            SetBusy(True, "Converting to " & Path.GetFileName(target) & "...")
            Try
                Await Program.Tool.ConvertAsync(_disk.Path, target, True)
            Catch ex As Exception
                failure = ex.Message
            End Try
        End If
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            Ui.Fail(Me, failure, "Save image as")
            Return
        End If
        If Ui.Ask(Me, "Written: " & target & vbLf & vbLf & "Open the new image now?", "Save image as") Then Await OpenImageAsync(target)
    End Sub

    ' =================================================================== Disk menu (trs80-tool features beyond TRSTool)

    Private Sub OnSectorMap()
        If _busy OrElse _disk Is Nothing Then Return
        If _disk.IsHardDisk OrElse _disk.IsCassette Then
            Ui.Info(Me, "trs80-tool's sector map works on floppy images only.", "Sector map")
            Return
        End If
        Dim f As New SectorMapForm(_disk)
        f.ShowOwned(Me)
    End Sub

    ''' <summary>Disk -> DMK analysis: the program's own reading of the DMK (DmkAnalyzer, DosIdentifier, DiskInfoForm).</summary>
    Private Async Sub OnDmkInfo()
        If _busy OrElse _disk Is Nothing Then Return
        If _disk.IsHardDisk OrElse _disk.IsCassette Then
            Ui.Info(Me, "The DMK analysis works on floppy images.", "DMK analysis")
            Return
        End If
        Dim f As New DiskInfoForm(_disk.Path)
        f.ShowOwned(Me)
        Await f.LoadAsync()
    End Sub

    Private Sub OnGranuleMap()
        If _busy OrElse _disk Is Nothing Then Return
        Dim f As New GranuleMapForm(_disk)
        f.ShowOwned(Me)
    End Sub

    Private Async Sub OnCheckRepair()
        If _busy OrElse _disk Is Nothing Then Return
        Dim dlg As New CheckRepairForm(_disk)
        dlg.ShowModal(Me)
        If dlg.Changed Then Await RefreshAsync()
    End Sub

    Private Async Sub OnFormat()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable OrElse _disk.IsHardDisk Then Return
        If HostFiles.ImageExtension(_disk.Path) <> ".dmk" Then
            Ui.Info(Me, "trs80-tool's format makes a DMK image, so it only formats images named *.dmk. Use File → Save image as to make a DMK copy first, then format that.", "Format")
            Return
        End If
        Dim cyl As Integer = If(_disk.Geometry IsNot Nothing, _disk.Geometry.CylinderCount, 0)
        Dim sides As Integer = If(_disk.Geometry IsNot Nothing, _disk.Geometry.SideCount, 0)
        Dim density As String = If(_disk.Geometry IsNot Nothing AndAlso _disk.Geometry.DataTrack IsNot Nothing, _disk.Geometry.DataTrack.Density, Nothing)
        Dim dlg As New CreateFloppyForm(True, _disk.DosName, cyl, sides, density)
        If Not dlg.ShowOk(Me) Then Return
        Dim fileCount As Integer = _disk.VisibleFiles(True, True).Count
        Dim msg As String = "Format erases " & If(fileCount = 1, "the 1 file", "all " & fileCount.ToString(CultureInfo.InvariantCulture) & " files") & " on " & Path.GetFileName(_disk.Path) & " and formats it as an empty " & If(dlg.IsCpm, "Montezuma Micro CP/M", dlg.DosName) & " data disk." & vbLf & vbLf & "This cannot be undone. Go ahead?"
        If Not Ui.Ask(Me, msg, "Format Disk Image", MessageBoxType.Warning, True) Then Return
        SetBusy(True, "Formatting...")
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
        If failure IsNot Nothing Then Ui.Fail(Me, failure, "Format Disk Image")
        Await RefreshAsync()
    End Sub

    Private Async Sub OnTrim()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk Then Return
        Dim dlg As New TrimForm(_disk)
        If Not dlg.ShowOk(Me) Then Return
        SetBusy(True, "Trimming...")
        Dim r As ToolResult = Nothing
        Dim failure As String = Nothing
        Try
            r = Await Program.Tool.DiskTrimAsync(_disk.Path, dlg.OptionWords())
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            Ui.Fail(Me, failure, "Trim")
        ElseIf Not r.Succeeded Then
            Ui.Warn(Me, r.ErrorText, "Trim")
        Else
            Ui.Info(Me, r.StdOut.Trim(), "Trim")
        End If
        Await RefreshAsync()
    End Sub

    Private Sub OnBoot()
        If _busy OrElse _disk Is Nothing OrElse _disk.IsHardDisk OrElse ConsoleRunning() Then Return
        Dim dlg As New BootForm(_disk)
        If Not dlg.ShowOk(Me) Then Return
        StartConsole(dlg.RunWords(_disk.Path), "The emulator")
    End Sub

    ''' <summary>Tools -> Run the emulator without a disk: "trs80-tool run --model N" in its own console.</summary>
    Private Sub OnRunEmulatorEmpty()
        If _busy OrElse ConsoleRunning() Then Return
        Dim dlg As New RunModelForm()
        If Not dlg.ShowOk(Me) Then Return
        StartConsole(New List(Of String) From {"run", "--model", dlg.Model.ToString(CultureInfo.InvariantCulture)}, "The emulator")
    End Sub

    Private Sub OnShell()
        If _busy OrElse _disk Is Nothing OrElse ConsoleRunning() Then Return
        ' "disk IMAGE [--partition P] [--cpm-format NAME]": the shell reads the image the way this window does.
        Dim words As List(Of String) = Program.Tool.DiskWords(_disk.Path, _disk.PartitionArg)
        StartConsole(words, "The trs80-tool shell")
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
        SetBusy(True, "Asking trs80-tool for its CP/M layouts...")
        Try
            layouts = Await Program.Tool.CpmFormatHelpAsync(imagePath)
        Catch ex As Exception
            layouts = ex.Message
        End Try
        SetBusy(False, Nothing)
        Dim dlg As New CpmFormatForm(layouts, before)
        If Not dlg.ShowOk(Me) Then Return
        Dim chosen As String = dlg.Chosen
        If String.Equals(chosen, before, StringComparison.OrdinalIgnoreCase) Then Return
        Program.Tool.SetCpmFormat(imagePath, chosen)
        If chosen.Length > 0 Then
            ' Ask trs80-tool whether the disk can be read that way before the window is reloaded with it.
            Dim problem As String = Nothing
            SetBusy(True, "Reading " & Path.GetFileName(imagePath) & " as " & chosen & "...")
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
                Ui.Warn(Me, "trs80-tool can't read this disk as """ & chosen & """:" & vbLf & vbLf & problem & vbLf & vbLf & "The disk is still read as before.", "CP/M layout")
                Return
            End If
        End If
        Await OpenImageAsync(imagePath)
    End Sub

    ''' <summary>Start trs80-tool in a terminal window and refresh the image when it ends (it may have written it).</summary>
    Private Sub StartConsole(words As List(Of String), what As String)
        Try
            _console = Program.Tool.StartInConsole(words)
            _consoleWhat = what
            AddHandler _console.Exited, Sub()
                                           Try
                                               Application.Instance.AsyncInvoke(Async Sub()
                                                                                    ' Dispose removes the run's script folder.
                                                                                    Dim ended As ConsoleRun = _console
                                                                                    _console = Nothing
                                                                                    If ended IsNot Nothing Then ended.Dispose()
                                                                                    UpdateStatus()
                                                                                    UpdateEnabled()
                                                                                    Try
                                                                                        If _disk IsNot Nothing AndAlso Not _busy Then Await RefreshAsync()
                                                                                    Catch
                                                                                    End Try
                                                                                End Sub)
                                           Catch
                                           End Try
                                       End Sub
            UpdateStatus()
            UpdateEnabled()
        Catch ex As Exception
            _console = Nothing
            Ui.Fail(Me, ex.Message, Program.AppName)
        End Try
    End Sub

    ' =================================================================== Tools menu (host files)

    ''' <summary>Tools → Convert file(s): trs80-tool convert on host files; afterwards offer to add the result to the open disk.</summary>
    Private Async Sub OnConvertFiles(initialInputs As IEnumerable(Of String))
        If _busy Then Return
        Dim dlg As New ConvertFileForm(initialInputs)
        dlg.ShowModal(Me)
        Await OfferAddToDiskAsync(dlg.OutputPath)
    End Sub

    Private Async Sub OnAssemble()
        If _busy Then Return
        Dim dlg As New AssembleForm()
        dlg.ShowModal(Me)
        Await OfferAddToDiskAsync(dlg.OutputPath)
    End Sub

    ''' <summary>Context menu → Convert to…: extract the selected file to a temp host file, then open the Convert dialog on it.</summary>
    Private Async Sub OnConvertDiskFile()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Dim host As String = Path.Combine(Program.NewTempDir(), TextUtil.HostNameFromTrs(f.TrsName, False))
        SetBusy(True, "Extracting " & f.SlashName & "...")
        Dim failure As String = Nothing
        Try
            Await Program.Tool.DiskExtractAsAsync(_disk.Path, _disk.PartitionArg, f.TrsName, host)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            Ui.Fail(Me, "Could not extract " & f.SlashName & ": " & failure, "Convert")
            Return
        End If
        OnConvertFiles(New String() {host})
    End Sub

    Private Async Function OfferAddToDiskAsync(produced As String) As Task
        If produced Is Nothing OrElse _disk Is Nothing OrElse Not _disk.IsWritable OrElse Not File.Exists(produced) Then Return
        Dim ext As String = Path.GetExtension(produced).ToLowerInvariant()
        If ext = ".dmk" OrElse ext = ".jv1" OrElse ext = ".jv3" OrElse ext = ".hfe" OrElse ext = ".imd" OrElse ext = ".wav" Then Return
        If Ui.Ask(Me, "Add " & Path.GetFileName(produced) & " to " & Path.GetFileName(_disk.Path) & "?", Program.AppName) Then
            Await AddHostFilesAsync(New String() {produced})
        End If
    End Function

    Private Async Sub OnHexdumpFile()
        If _busy Then Return
        Dim chosen As String = Ui.PickFile(Me, "Hex dump of a file (trs80-tool hexdump)", "All files (*.*)|*.*")
        If chosen Is Nothing Then Return
        SetBusy(True, "Running hexdump...")
        Dim text As String
        Try
            text = Await Program.Tool.HexdumpAsync(chosen, False)
        Catch ex As Exception
            text = "(" & ex.Message & ")"
        End Try
        SetBusy(False, Nothing)
        Dim f As New TextReportForm("hexdump - " & Path.GetFileName(chosen), text)
        f.ShowOwned(Me)
    End Sub

    Private Async Sub OnDisasmFile()
        If _busy Then Return
        Dim chosen As String = Ui.PickFile(Me, "Disassemble a file (trs80-tool disasm)", "TRS-80 programs (*.cmd;*.sys;*.3bn;*.rom;*.bin;*.cas)|*.cmd;*.sys;*.3bn;*.rom;*.bin;*.cas|All files (*.*)|*.*")
        If chosen Is Nothing Then Return
        Dim opt As New DisasmOptionsForm(Program.Settings.DisasmOptionWords())
        If Not opt.ShowOk(Me) Then Return
        Dim words As List(Of String) = opt.OptionWords()
        SetBusy(True, "Disassembling...")
        Dim text As String
        Try
            text = Await Program.Tool.DisasmAsync(chosen, words)
        Catch ex As Exception
            text = "(" & ex.Message & ")"
        End Try
        SetBusy(False, Nothing)
        Dim f As New TextReportForm("disasm - " & Path.GetFileName(chosen), text)
        f.ShowOwned(Me)
    End Sub

    ' =================================================================== View menu

    Private Sub OnViewToggle()
        If _settingChecks Then Return
        Program.Settings.ShowInvisible = _miInvisible.Checked
        Program.Settings.ShowSystem = _miSystem.Checked
        FillList(SelectedNames())
    End Sub

    ' =================================================================== Edit: preferences / about / engine

    Private Async Sub OnPreferences()
        If _busy Then Return
        Dim dlg As New PreferencesForm()
        If Not dlg.ShowOk(Me) Then Return
        Program.Settings.Save()
        Program.Tool.TerminalCommand = If(Program.Settings.TerminalCommand, "")
        Dim newPath As String = Trs80Tool.Locate(Program.Settings.ToolPath)
        If newPath IsNot Nothing AndAlso Not String.Equals(newPath, Program.Tool.ToolPath, HostPlatform.PathComparison) Then
            Program.Tool.ToolPath = newPath
            Program.Tool.Capabilities = Nothing
            Await CheckEngineAsync(False)
        ElseIf newPath Is Nothing Then
            Program.Tool.ToolPath = ""
            Program.Tool.Capabilities = Nothing
            _lblEngine.Text = "trs80-tool: not found"
        End If
        FillList(SelectedNames())
    End Sub

    Private Sub OnAbout()
        Ui.Info(Me, Program.AppName & " " & Program.ReleaseVersion & vbLf &
                "Release date: " & Program.ReleaseDate & vbLf &
                "Build: " & Program.AppVersion & " (Eto.Forms, " & Application.Instance.Platform.ID & ")" & vbLf &
                "By " & Program.AppAuthors & vbLf &
                "Free software under the MIT License (Help > License)." & vbLf &
                "Download: " & AppInfo.DownloadUrl & vbLf & vbLf &
                "A front end for Lawrence Kesteloot's trs80-tool, reproducing the workflow of Matthew Reed's TRSTools (1.14)." & vbLf & vbLf &
                "Engine: " & If(Program.Tool.Capabilities IsNot Nothing AndAlso Program.Tool.Capabilities.Version IsNot Nothing, "trs80-tool " & Program.Tool.Capabilities.VersionText, "trs80-tool not found") & vbLf &
                "Engine path: " & If(Program.Tool.ToolPath, "") & vbLf & vbLf &
                "trs80-tool: " & AppInfo.ToolUrl & "  (MIT licence, © Lawrence Kesteloot)" & vbLf &
                "TRSTools was © 2009-2021 Matthew Reed; TRSToolGui is modelled on it (Help > License has the details).",
                "About " & Program.AppName)
    End Sub

    ''' <summary>Help > Manual / Changelog / License: an embedded text file in the report window (Copy / Save / Close).</summary>
    Private Sub ShowHelpText(resourceName As String, title As String)
        Dim f As New TextReportForm(title, Program.ReadHelpText(resourceName))
        f.ShowOwned(Me)
    End Sub

    Private Async Sub OnEngineInfo()
        If _busy Then Return
        If Program.Tool.Capabilities Is Nothing AndAlso Not String.IsNullOrEmpty(Program.Tool.ToolPath) Then Await CheckEngineAsync(False)
        Dim dlg As New EngineInfoForm(Program.Tool)
        dlg.ShowModal(Me)
        If dlg.ReprobeRequested Then Await CheckEngineAsync(True)
    End Sub

    ' =================================================================== viewer / properties / rename / delete

    Private Sub OnPreview()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count = 0 Then Return
        Dim v As New ViewerForm(_disk, sel(0))
        v.ShowOwned(Me)
    End Sub

    Private Async Sub OnFileProperties()
        If _busy OrElse _disk Is Nothing Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Dim dlg As New FilePropertiesForm(_disk, f)
        If Not dlg.ShowOk(Me) Then Return
        Dim changed As Boolean = False
        Dim failure As String = Nothing
        SetBusy(True, "Updating " & f.SlashName & "...")
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
            Ui.Fail(Me, failure, "File Properties Error")
            Await RefreshAsync()
            Return
        End If
        If changed Then Await RefreshAsync()
    End Sub

    Private Async Sub OnRename()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count <> 1 Then Return
        Dim f As DiskFile = sel(0)
        Dim dlg As New RenameForm("Rename File", "Filename:", f.SlashName, Nothing)
        If Not dlg.ShowOk(Me) Then Return
        Dim newName As String = TextUtil.NormalizeTrsName(dlg.ResultName)
        If newName Is Nothing OrElse String.Equals(newName, f.TrsName, StringComparison.OrdinalIgnoreCase) Then Return
        SetBusy(True, "Renaming...")
        Dim failure As String = Nothing
        Try
            Await Program.Tool.DiskRenameAsync(_disk.Path, _disk.PartitionArg, f.TrsName, newName)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then
            Ui.Fail(Me, "Error renaming file: " & failure, "Rename error")
            Return
        End If
        Await RefreshAsync()
        ReselectNames(New List(Of String) From {newName})
    End Sub

    Private Async Sub OnDelete()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim sel As List(Of DiskFile) = SelectedFiles()
        If sel.Count = 0 Then Return
        If Program.Settings.ConfirmDelete Then
            Dim msg As String = If(sel.Count = 1, "Are you sure you want to delete this file?" & vbLf & vbLf & sel(0).SlashName, "Are you sure you want to delete these " & sel.Count.ToString(CultureInfo.InvariantCulture) & " files?")
            If Not Ui.Ask(Me, msg, "Confirm File Delete", MessageBoxType.Warning, True) Then Return
        End If
        Dim names As New List(Of String)
        For Each f In sel
            names.Add(f.TrsName)
        Next
        SetBusy(True, "Deleting...")
        Dim failure As String = Nothing
        Try
            Await Program.Tool.DiskKillAsync(_disk.Path, _disk.PartitionArg, names)
        Catch ex As Exception
            failure = ex.Message
        End Try
        SetBusy(False, Nothing)
        If failure IsNot Nothing Then Ui.Fail(Me, "Error deleting file: " & failure, "Delete error")
        Await RefreshAsync()
    End Sub

    ' =================================================================== adding files (Add / Paste / drop)

    Private Async Sub OnEditAdd()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim names As String()
        Using dlg As New OpenFileDialog()
            dlg.Title = "Select files to add"
            Ui.SetFilters(dlg, "All files (*.*)|*.*")
            dlg.MultiSelect = True
            dlg.CheckFileExists = True
            If dlg.ShowDialog(Me) <> DialogResult.Ok Then Return
            names = dlg.Filenames.ToArray()
        End Using
        Await AddHostFilesAsync(names)
    End Sub

    Private Async Sub OnEditPaste()
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        Dim paths As New List(Of String)
        For Each p In Ui.FilesOnClipboard()
            If File.Exists(p) Then paths.Add(p)
        Next
        If paths.Count = 0 Then
            _lblStatus.Text = "There are no files on the clipboard to paste."
            Return
        End If
        Await AddHostFilesAsync(paths)
    End Sub

    Private Sub OnDragOverList(sender As Object, e As DragEventArgs)
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable OrElse Not e.Data.ContainsUris Then
            e.Effects = DragEffects.None
        Else
            e.Effects = DragEffects.Copy
        End If
    End Sub

    Private Async Sub OnDragDropList(sender As Object, e As DragEventArgs)
        If _busy OrElse _disk Is Nothing OrElse Not _disk.IsWritable Then Return
        If Not e.Data.ContainsUris Then Return
        Dim paths As New List(Of String)
        For Each u In e.Data.Uris
            If u Is Nothing OrElse Not u.IsFile Then Continue For
            Dim p As String = u.LocalPath
            If File.Exists(p) Then
                paths.Add(p)
            ElseIf Directory.Exists(p) Then
                ' A dropped folder: its files (not recursive), like dropping them one by one.
                paths.AddRange(Directory.GetFiles(p))
            End If
        Next
        If paths.Count = 0 Then Return
        Try
            Await AddHostFilesAsync(paths)
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, Program.AppName)
        End Try
    End Sub

    ''' <summary>
    ''' Store host files on the disk, with TRSTool's checks: name validity (rename dialog), existing
    ''' file (replace dialog), text conversion preference, host date. One all-or-nothing engine call.
    ''' </summary>
    Private Async Function AddHostFilesAsync(hostPaths As IEnumerable(Of String)) As Task
        If _disk Is Nothing OrElse Not _disk.IsWritable Then
            Ui.Fail(Me, "The disk image cannot be written to." & If(_disk IsNot Nothing AndAlso _disk.ReadOnlyReason IsNot Nothing, vbLf & vbLf & _disk.ReadOnlyReason, ""), "File Write Error")
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
                Ui.Info(Me, "Skipping " & fi.Name & ": larger than 10 MB.", Program.AppName)
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
                    Ui.Warn(Me, "Skipping " & fi.Name & ": " & If(s.UnzipGzOnAdd, "it could not be unzipped (", "it could not be copied (") & ex.Message & ").", Program.AppName)
                    Continue For
                End Try
            End If

            ' 1. TRS-80 name
            Dim trsName As String = TextUtil.TrsNameFromHost(sourceName)
            Dim hostUpper As String = sourceName.ToUpperInvariant()
            Dim needsRename As Boolean = trsName.Length = 0 OrElse Not String.Equals(trsName, hostUpper, StringComparison.Ordinal)
            If needsRename Then
                Dim dlg As New RenameForm("Rename File", "New filename:", If(trsName.Length = 0, "", trsName.Replace(".", "/")), sourceName & vbLf & "is not a valid TRS-80 filename.")
                If Not dlg.ShowOk(Me) Then Continue For
                trsName = TextUtil.NormalizeTrsName(dlg.ResultName)
                If trsName Is Nothing Then
                    Ui.Warn(Me, """" & dlg.ResultName & """ is not a valid TRS-80 filename. Skipping " & fi.Name & ".", Program.AppName)
                    Continue For
                End If
            End If
            If pendingNames.Contains(trsName) Then
                Ui.Warn(Me, "Two of the files would both be " & trsName.Replace(".", "/") & " on the disk. Skipping " & fi.Name & ".", Program.AppName)
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
                    Dim dlg As New ReplaceForm(existing, trsName.Replace(".", "/"), fi.Length, fi.LastWriteTime)
                    Dim yes As Boolean = dlg.ShowReplace(Me)
                    If dlg.ApplyToAll Then replaceAll = yes
                    If Not yes Then Continue For
                    overwrite = True
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

        If items.Count = 0 Then Return

        SetBusy(True, "Adding " & items.Count.ToString(CultureInfo.InvariantCulture) & " file(s)...")
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
            Dim heading As String = "Error storing file(s): "
            If failureKind = ToolFailure.DiskFull Then heading = "Not enough free space to store " & If(items.Count = 1, "this file.", "these files.")
            If failureKind = ToolFailure.WriteProtected Then heading = "The image is write-protected (Disk properties has the flag)."
            Ui.Fail(Me, heading & vbLf & vbLf & failure & vbLf & vbLf & "Nothing was written to the disk.", "File Storing Error")
            Await RefreshAsync()
            Return
        End If
        ' Surface the tool's warnings (e.g. the DOS couldn't store the date), once.
        Dim warn As New List(Of String)
        If dateNote IsNot Nothing Then warn.Add("The files were added without dates: " & dateNote)
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
        If warn.Count > 0 Then Ui.Info(Me, String.Join(vbLf, warn), "trs80-tool notes")
    End Function

    ' =================================================================== extracting (Extract / Copy)

    Private Async Sub OnEditExtract()
        If _busy OrElse _disk Is Nothing Then Return
        Dim dlg As New ExtractForm(SelectedCount())
        If Not dlg.ShowOk(Me) Then Return
        Dim files As List(Of DiskFile) = If(dlg.ExtractAll, AllListedFiles(), SelectedFiles())
        If files.Count = 0 Then Return
        Dim targetDir As String = dlg.TargetDirectory
        Try
            If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, "File Extraction Error")
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
                    Dim ask As New OverwriteForm(Path.GetFileName(dest), targetDir, staged.Count > 1)
                    Dim r As DialogResult = ask.ShowChoice(Me)
                    If r = DialogResult.Cancel Then Exit For
                    doIt = (r = DialogResult.Yes)
                    If ask.ApplyToAll Then overwriteAll = doIt
                End If
                If Not doIt Then Continue For
            End If
            Try
                File.Copy(src, dest, True)
                written += 1
            Catch ex As Exception
                Ui.Fail(Me, "Error extracting file: " & ex.Message, "File Extraction Error")
            End Try
        Next
        _lblStatus.Text = written.ToString(CultureInfo.InvariantCulture) & " file(s) extracted to " & targetDir
    End Sub

    Private Async Sub OnEditCopy()
        If _busy OrElse _disk Is Nothing Then Return
        Dim files As List(Of DiskFile) = SelectedFiles()
        If files.Count = 0 Then Return
        Dim staged As List(Of String) = Await ExportToTempAsync(files, New ExtractOptions With {.Lowercase = Program.Settings.LowercaseExtractNames})
        If staged Is Nothing OrElse staged.Count = 0 Then Return
        Try
            Ui.CopyFilesToClipboard(staged)
            _lblStatus.Text = staged.Count.ToString(CultureInfo.InvariantCulture) & " file(s) copied to the clipboard"
        Catch ex As Exception
            Ui.Fail(Me, ex.Message, Program.AppName)
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
        SetBusy(True, "Extracting " & files.Count.ToString(CultureInfo.InvariantCulture) & " file(s)...")
        Dim failure As String = Nothing
        Try
            If options.NeedsTopLevelExtract Then
                ' --cpm-records / --ucsd-pascal-raw exist only on the top-level extract command.
                Await Program.Tool.ExtractTopLevelAsync(_disk.Path, names, folder, lowercase, options.MarkDamaged, options.CpmRecords, options.UcsdRaw)
            Else
                Await Program.Tool.DiskExtractAsync(_disk.Path, _disk.PartitionArg, names, folder, lowercase, options.MarkDamaged)
            End If
        Catch ex As Exception
            failure = ex.Message
        End Try
        If failure IsNot Nothing Then
            SetBusy(False, Nothing)
            Ui.Fail(Me, "Error extracting file(s): " & failure, "File Extraction Error")
            Return Nothing
        End If
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
