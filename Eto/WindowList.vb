' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Eto\WindowList.vb — the main windows that are open, for the Linux and macOS
'  program. The Windows Forms program has its own copy, WindowList.vb in the
'  root folder, which keeps the same rules in that toolkit's idiom; the header
'  there explains WHY the list exists (zebbchi's report of 26 November 2026),
'  what the two orders are for, and why only the first window of a run saves the
'  size and position. The placement arithmetic is shared, in
'  Core\WindowCascade.vb, and is the same on all three systems.
'
'  THE ONE THING THAT IS DIFFERENT HERE, and it is a Mac thing. A Mac has one
'  menu bar for the whole program and expects a Window menu in it; Eto puts one
'  there itself, with Minimize, Zoom and Bring All To Front in it. But Eto never
'  assigns NSApplication.WindowsMenu — read out of Eto 2.11.0's own
'  MenuBarHandler, not assumed — so AppKit does NOT fill that menu with the list
'  of windows, and on the Mac as much as on Linux the list has to be ours.
'
'  So that the Mac ends up with ONE Window menu rather than two, our own menu and
'  Eto's are made to be the same menu. Eto looks its system items up by name
'  (MenuItemCollection.GetSubmenu, matching on the text with the & taken out), so
'  all that is needed is for the name to match — and that is a problem in French,
'  because Eto looks for "Window" and ours says "Fenêtre". Eto's answer to
'  exactly this is Application.LocalizeString, the event it raises for every
'  system-supplied string; Eto\Program.vb answers it for "&Window" with whatever
'  the language table says, so the two names agree and Eto's three items land in
'  our menu. Nothing of Eto's is replaced or worked around.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO
Imports Eto.Drawing
Imports Eto.Forms

''' <summary>The open main windows, their front-to-back order, and where the next one goes.</summary>
Friend Module WindowList

    ''' <summary>Creation order. The first entry is the geometry keeper.</summary>
    Private ReadOnly _open As New List(Of MainForm)

    ''' <summary>Front to back: _front(0) is the frontmost window.</summary>
    Private ReadOnly _front As New List(Of MainForm)

    ''' <summary>The window that owns the saved size and position; see the Windows Forms header.</summary>
    Private _keeper As MainForm

    ''' <summary>True while File > Exit is closing every window, so a second Exit cannot start over it.</summary>
    Private _exiting As Boolean

    ''' <summary>How many main windows are open.</summary>
    Friend ReadOnly Property Count As Integer
        Get
            Return _open.Count
        End Get
    End Property

    ''' <summary>The open windows in creation order, which is the order the Window menu lists them in.</summary>
    Friend Function All() As List(Of MainForm)
        Return New List(Of MainForm)(_open)
    End Function

    ''' <summary>The frontmost window, or Nothing when none is open.</summary>
    Friend Function Frontmost() As MainForm
        If _front.Count = 0 Then Return Nothing
        Return _front(0)
    End Function

    ''' <summary>True for the one window that saves its size and position.</summary>
    Friend Function IsKeeper(f As MainForm) As Boolean
        Return f IsNot Nothing AndAlso ReferenceEquals(f, _keeper)
    End Function

    ''' <summary>
    ''' A new main window joins the list, from its own constructor, so that it is known before it is
    ''' placed and shown. It joins the BACK of the front-to-back order: it is not the frontmost window
    ''' until it really is in front, which is what NoteActivated records.
    ''' </summary>
    Friend Sub Register(f As MainForm)
        If f Is Nothing OrElse _open.Contains(f) Then Return
        _open.Add(f)
        _front.Add(f)
        If _keeper Is Nothing Then _keeper = f
    End Sub

    ''' <summary>A window has closed.</summary>
    Friend Sub Unregister(f As MainForm)
        If f Is Nothing Then Return
        _open.Remove(f)
        _front.Remove(f)
        ' On Linux Eto ends the program itself when the last top-level window goes; on a Mac the program
        ' stays running with no window, which is what a Mac does and what Program.ExitProgram knows.
        RebuildMenus()
    End Sub

    ''' <summary>This window is now in front: it goes to the head of the order and every tick moves with it.</summary>
    Friend Sub NoteActivated(f As MainForm)
        If f Is Nothing OrElse Not _open.Contains(f) Then Return
        If _front.Count > 0 AndAlso ReferenceEquals(_front(0), f) Then Return
        _front.Remove(f)
        _front.Insert(0, f)
        RebuildMenus()
    End Sub

    ''' <summary>
    ''' Every window's Window menu is rebuilt, not only the frontmost one's. On Linux they are all on the
    ''' screen at once and all of their ticks have to agree; on a Mac only the frontmost window's menu bar
    ''' is the one on show, but which window that is changes with every click, so they are all kept right.
    ''' </summary>
    Friend Sub RebuildMenus()
        For Each f As MainForm In _open
            Try
                f.RebuildWindowMenu()
            Catch
                ' A menu that cannot be rebuilt is not a reason to stop: the rest still can be.
            End Try
        Next
    End Sub

    ''' <summary>
    ''' The window showing this image, or Nothing. The parameter is not called "path", because VB does
    ''' not mind about case and a parameter of that name hides System.IO.Path from the whole routine.
    ''' </summary>
    Friend Function FindByPath(imagePath As String) As MainForm
        If String.IsNullOrEmpty(imagePath) Then Return Nothing
        Dim full As String
        Try
            full = Path.GetFullPath(imagePath)
        Catch
            full = imagePath
        End Try
        For Each f As MainForm In _open
            Dim shown As String = f.ImagePath
            If shown IsNot Nothing AndAlso String.Equals(shown, full, HostPlatform.PathComparison) Then Return f
        Next
        Return Nothing
    End Function

    ''' <summary>Bring a window to the front, un-minimizing it first if it needs it.</summary>
    Friend Sub Activate(f As MainForm)
        If f Is Nothing Then Return
        Try
            If f.WindowState = WindowState.Minimized Then f.WindowState = WindowState.Normal
            f.BringToFront()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Open an image, deciding WHICH window gets it. Every decision here is made before anything is
    ''' awaited, which is the point: the Finder hands the program every file of one open in a single
    ''' event, and three overlapping asynchronous opens that each picked their window after awaiting
    ''' would all pick the same one — which is precisely the fault zebbchi saw.
    '''   * already open in a window  -> that window comes forward, and nothing is read again
    '''   * a window with no image    -> that window, claimed now so a second file cannot take it too
    '''   * otherwise                 -> a new window, cascaded from the frontmost one
    ''' </summary>
    Friend Sub OpenPath(imagePath As String)
        If String.IsNullOrEmpty(imagePath) Then Return

        Dim already As MainForm = FindByPath(imagePath)
        If already IsNot Nothing Then
            Activate(already)
            Return
        End If

        Dim target As MainForm = ClaimEmpty()
        If target IsNot Nothing Then
            Activate(target)
            target.BeginOpen(imagePath)
            Return
        End If

        target = NewWindow()
        If target IsNot Nothing Then target.BeginOpen(imagePath)
    End Sub

    ''' <summary>
    ''' A window with no image open and none on its way, claimed so that it cannot be handed a second
    ''' file in the same breath. The frontmost such window is preferred, being the one the user can see.
    ''' </summary>
    Private Function ClaimEmpty() As MainForm
        For Each f As MainForm In _front
            If f.IsFree Then
                f.Claim()
                Return f
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>A new main window, placed by the cascade and shown. It is claimed, so it stays empty for its caller.</summary>
    Friend Function NewWindow() As MainForm
        Dim f As MainForm
        Try
            f = New MainForm(Nothing)
        Catch
            Return Nothing
        End Try
        f.Claim()
        PlaceCascaded(f)
        f.Show()
        Return f
    End Function

    ''' <summary>
    ''' Put this window down and right of the frontmost one, inside the visible frame of the screen that
    ''' window is on. The arithmetic is WindowCascade.Place; what is done here is only the reading of the
    ''' real screens and the real windows, which is the part no harness can do.
    ''' </summary>
    Friend Sub PlaceCascaded(f As MainForm)
        If f Is Nothing Then Return
        Try
            Dim anchor As MainForm = Frontmost()
            If anchor IsNot Nothing AndAlso ReferenceEquals(anchor, f) Then anchor = Nothing

            ' A maximized window has no corner worth stepping from, so it is not an anchor.
            Dim hasAnchor As Boolean = anchor IsNot Nothing AndAlso anchor.WindowState = WindowState.Normal

            Dim area As RectangleF
            If anchor IsNot Nothing Then
                area = Screen.FromPoint(New PointF(anchor.Location.X + 20, anchor.Location.Y + 20)).WorkingArea
            Else
                area = Screen.PrimaryScreen.WorkingArea
            End If

            Dim taken As New List(Of WindowSpot)
            For Each o As MainForm In _open
                If ReferenceEquals(o, f) Then Continue For
                If o.WindowState <> WindowState.Normal Then Continue For
                taken.Add(New WindowSpot(o.Location.X, o.Location.Y))
            Next

            Dim ax As Integer = 0
            Dim ay As Integer = 0
            If hasAnchor Then
                ax = anchor.Location.X
                ay = anchor.Location.Y
            End If

            ' Before a window is on the screen Eto may not have worked its outside size out yet, so the
            ' size the window ASKED for is what the clamp is given when the real one is not there.
            Dim w As Integer = f.Size.Width
            Dim h As Integer = f.Size.Height
            If w < 50 OrElse h < 50 Then
                w = Math.Max(480, Program.Settings.WindowWidth)
                h = Math.Max(300, Program.Settings.WindowHeight)
            End If

            Dim spot As WindowSpot = WindowCascade.Place(hasAnchor, ax, ay, w, h,
                                                         CInt(Math.Floor(area.X)), CInt(Math.Floor(area.Y)),
                                                         CInt(Math.Floor(area.Width)), CInt(Math.Floor(area.Height)),
                                                         WindowCascade.DefaultStep, taken)
            ' A window placed by the cascade is never maximized: the remembered state belongs to the
            ' first window of the run, and a maximized second window would hide the first completely.
            f.WindowState = WindowState.Normal
            f.Location = New Point(spot.X, spot.Y)
        Catch
            ' Any trouble reading the screens, and the window still appears where Eto would have put it.
        End Try
    End Sub

    ''' <summary>
    ''' File > Exit on Linux: the whole program, not this one window, so every window is closed and each
    ''' gets its own closing work done. A Mac asks the application to quit instead (Program.ExitProgram),
    ''' because that is what Cmd+Q means there.
    ''' </summary>
    Friend Sub CloseAll()
        If _exiting Then Return
        _exiting = True
        Try
            For Each f As MainForm In All()
                f.Close()
            Next
        Finally
            _exiting = False
        End Try
    End Sub

End Module
