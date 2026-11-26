' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  WindowList.vb — the main windows that are open. Windows Forms only; the Eto
'  program has its own copy of this, in Eto\WindowList.vb, which keeps the same
'  rules in Eto's idiom. The placement arithmetic both of them use is shared, in
'  Core\WindowCascade.vb.
'
'  WHY THIS EXISTS. Until v1.8.3 the program had exactly one main window and one
'  open image. Asking it to open several — which is what a Mac user does by
'  selecting a few images in the Finder and pressing Return — did not open
'  several: every file was poured into the one window, one after another, and
'  only the last of them survived. zebbchi reported the symptom on 26 November
'  2026: one window on the screen, but the program behaving as though more than
'  one image were open. The diagnosis and both halves of the fault are in
'  Eto\MacOpenDocuments.vb's header; this class is the other half of the repair.
'
'  WHAT IT KEEPS. Two orders over the same windows, because both are needed and
'  neither can be worked out from the other:
'    * creation order, which is the order the Window menu lists them in, so that
'      the numbers beside the entries do not jump about as the user clicks from
'      one window to another;
'    * front-to-back order, whose first entry is the frontmost window. That is
'      what the Window menu puts its tick beside, and what the cascade is
'      anchored on.
'
'  THE GEOMETRY KEEPER. TRSToolGui.cfg holds one window size and one window
'  position, and several windows cannot all sensibly write to them. The FIRST
'  window of a run owns them; the rest save nothing. The reason is that the first
'  window's position is the user's own — it is where they left the program last
'  time — while every other window's position was chosen by the cascade below,
'  and saving one of those would walk the remembered position a little further
'  down the screen at every run. The role is never handed on: if the keeper is
'  closed while other windows are still open, the size and place already saved
'  are the ones the user chose, and they stand.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

''' <summary>The open main windows, their front-to-back order, and where the next one goes.</summary>
Friend Module WindowList

    ''' <summary>Creation order. The first entry is the geometry keeper.</summary>
    Private ReadOnly _open As New List(Of MainForm)

    ''' <summary>Front to back: _front(0) is the frontmost window.</summary>
    Private ReadOnly _front As New List(Of MainForm)

    ''' <summary>The window that owns the saved size and position; see the note above.</summary>
    Private _keeper As MainForm

    ''' <summary>Set by Program.Main: how to end the message loop once the last window has gone.</summary>
    Friend AllClosed As Action

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
        For Each f As MainForm In _front
            If Not f.IsDisposed Then Return f
        Next
        Return Nothing
    End Function

    ''' <summary>True for the one window that saves its size and position. See the note above.</summary>
    Friend Function IsKeeper(f As MainForm) As Boolean
        Return f IsNot Nothing AndAlso ReferenceEquals(f, _keeper)
    End Function

    ''' <summary>
    ''' A new main window joins the list. Called from the window's own constructor, so that a window
    ''' is known before it is placed and shown. It joins the BACK of the front-to-back order: it is
    ''' not the frontmost window until it is really in front, which is what NoteActivated records.
    ''' </summary>
    Friend Sub Register(f As MainForm)
        If f Is Nothing OrElse _open.Contains(f) Then Return
        _open.Add(f)
        _front.Add(f)
        If _keeper Is Nothing Then _keeper = f
    End Sub

    ''' <summary>A window has closed. When it was the last one, the program can end.</summary>
    Friend Sub Unregister(f As MainForm)
        If f Is Nothing Then Return
        _open.Remove(f)
        _front.Remove(f)
        If _open.Count = 0 Then
            Dim done As Action = AllClosed
            If done IsNot Nothing Then done()
            Return
        End If
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
    ''' Every window's Window menu is rebuilt, not only the frontmost one's: on Windows and Linux all of
    ''' them are on the screen at once and all of their ticks have to agree.
    ''' </summary>
    Friend Sub RebuildMenus()
        For Each f As MainForm In _open
            If f.IsDisposed Then Continue For
            Try
                f.RebuildWindowMenu()
            Catch
                ' A menu that cannot be rebuilt is not a reason to stop: the rest still can be.
            End Try
        Next
    End Sub

    ''' <summary>The window showing this image, or Nothing. Two windows never share one image; see OpenPath.</summary>
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
            If f.IsDisposed Then Continue For
            Dim shown As String = f.ImagePath
            If shown IsNot Nothing AndAlso String.Equals(shown, full, HostPlatform.PathComparison) Then Return f
        Next
        Return Nothing
    End Function

    ''' <summary>Bring a window to the front, un-minimizing it first if it needs it.</summary>
    Friend Sub Activate(f As MainForm)
        If f Is Nothing OrElse f.IsDisposed Then Return
        Try
            If f.WindowState = FormWindowState.Minimized Then f.WindowState = FormWindowState.Normal
            f.Activate()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Open an image, deciding WHICH window gets it. Every decision here is made before anything is
    ''' awaited, which is the point: the Finder can hand the program several files at once, and three
    ''' overlapping asynchronous opens that each pick their window after awaiting would all pick the
    ''' same one. Claiming a window is what stops that.
    '''   * already open in a window  -> that window comes forward, and nothing is read again
    '''   * a window with no image    -> that window, claimed now so a second file cannot take it too
    '''   * otherwise                 -> a new window, cascaded from the frontmost one
    ''' </summary>
    Friend Sub OpenPath(path As String)
        If String.IsNullOrEmpty(path) Then Return

        Dim already As MainForm = FindByPath(path)
        If already IsNot Nothing Then
            Activate(already)
            Return
        End If

        Dim target As MainForm = ClaimEmpty()
        If target IsNot Nothing Then
            Activate(target)
            target.BeginOpen(path)
            Return
        End If

        target = NewWindow()
        If target IsNot Nothing Then target.BeginOpen(path)
    End Sub

    ''' <summary>
    ''' A window with no image open and none on its way, claimed so that it cannot be handed a second
    ''' file in the same breath. The frontmost such window is preferred, because that is the one the
    ''' user is looking at.
    ''' </summary>
    Private Function ClaimEmpty() As MainForm
        For Each f As MainForm In _front
            If f.IsDisposed Then Continue For
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
            Dim area As Rectangle
            If anchor IsNot Nothing AndAlso Not anchor.IsDisposed Then
                area = Screen.FromControl(anchor).WorkingArea
            Else
                area = Screen.PrimaryScreen.WorkingArea
            End If

            Dim taken As New List(Of WindowSpot)
            For Each o As MainForm In _open
                If o.IsDisposed OrElse ReferenceEquals(o, f) Then Continue For
                If o.WindowState <> FormWindowState.Normal Then Continue For
                taken.Add(New WindowSpot(o.Left, o.Top))
            Next

            ' A maximized window has no corner worth stepping from, so it is not an anchor.
            Dim hasAnchor As Boolean = anchor IsNot Nothing AndAlso Not anchor.IsDisposed AndAlso
                                       Not ReferenceEquals(anchor, f) AndAlso anchor.WindowState = FormWindowState.Normal
            Dim ax As Integer = 0
            Dim ay As Integer = 0
            If hasAnchor Then
                ax = anchor.Left
                ay = anchor.Top
            End If

            Dim stepBy As Integer = WindowCascade.StepFor(SystemInformation.CaptionHeight + SystemInformation.FixedFrameBorderSize.Height)
            Dim spot As WindowSpot = WindowCascade.Place(hasAnchor, ax, ay, f.Width, f.Height,
                                                         area.X, area.Y, area.Width, area.Height,
                                                         stepBy, taken)
            ' A window placed by the cascade is never maximized: the remembered state belongs to the
            ' first window of the run, and a maximized second window would hide the first completely.
            f.WindowState = FormWindowState.Normal
            f.StartPosition = FormStartPosition.Manual
            f.Location = New Point(spot.X, spot.Y)
        Catch
            ' Any trouble reading the screens and the window still has to appear somewhere sane.
            f.StartPosition = FormStartPosition.WindowsDefaultLocation
        End Try
    End Sub

    ''' <summary>
    ''' File > Exit: the whole program, not this one window. Every window is closed, so each gets its
    ''' own closing work done, and the keeper saves the size and place. A window that refuses to close
    ''' (a question answered No) stops the exit, which is what a user who answered No meant.
    ''' </summary>
    Friend Sub CloseAll()
        If _exiting Then Return
        _exiting = True
        Try
            For Each f As MainForm In All()
                If f.IsDisposed Then Continue For
                f.Close()
                If Not f.IsDisposed Then Exit For
            Next
        Finally
            _exiting = False
        End Try
    End Sub

End Module
