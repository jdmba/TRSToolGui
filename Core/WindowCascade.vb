' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  WindowCascade.vb — where the next main window goes.
'
'  From v1.8.3 more than one disk image can be open at once, each in its own
'  main window, so the program has to decide where a new one lands. Letting the
'  system choose is what caused the fault Pascal Holdry reported in v1.8.2: a
'  modeless Windows Forms window falls back to the system cascade position,
'  which on a desktop with a second monitor to the LEFT of the main one can be
'  a negative X the mouse never reaches. UiFit.ShowOver fixed that for the
'  windows opened FROM the main one; this is the same question for the main
'  windows themselves, and the arithmetic is the same on all three systems, so
'  it lives here rather than twice over in two window layers.
'
'  Two rules, and they are the whole of it.
'
'  1. The cascade is anchored on the FRONTMOST window, not on the last one
'     opened and not on the first. That is what makes a cascade read as one:
'     the new window appears down and right of the window the user was just
'     looking at, wherever on the desktop that happens to be.
'
'  2. The result is clamped to the visible frame — the WORKING area of the
'     screen the anchor is on, so the taskbar, dock or panel is kept clear.
'     Right and bottom are pulled in first and left and top afterwards, which
'     is the order UiFit.ShowOver uses and for the same reason: when a window
'     is larger than the screen the top left corner has to win, because that is
'     the corner carrying the title bar the user needs in order to drag it.
'
'  When a step would put the window past the right or bottom edge the cascade
'  starts again at the top left of the working area, the way every program with
'  a Window menu has done it since the 1980s. A spot another main window is
'  already sitting on is stepped past rather than landed on, so two windows do
'  not hide each other exactly; that search is bounded, and if it runs out the
'  clamp still puts the window somewhere the user can reach.
'
'  Nothing here touches a window, a screen or a toolkit: it is integers in and
'  integers out, which is what lets the console harness check it (there is no
'  other way to check window placement without a Windows desktop, a GTK
'  desktop and a Mac).
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic

''' <summary>The top left corner of a window, in screen coordinates.</summary>
Public Structure WindowSpot

    Public ReadOnly X As Integer
    Public ReadOnly Y As Integer

    Public Sub New(left As Integer, top As Integer)
        X = left
        Y = top
    End Sub

    Public Overrides Function ToString() As String
        Return X.ToString(Globalization.CultureInfo.InvariantCulture) & "," & Y.ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

End Structure

''' <summary>Works out where the next main window goes. Framework-free arithmetic; see the note above.</summary>
Public NotInheritable Class WindowCascade

    Private Sub New()
        ' Nothing is ever made of this class.
    End Sub

    ''' <summary>The step to use when the system will not say how tall a title bar is.</summary>
    Public Const DefaultStep As Integer = 26

    ''' <summary>Smallest and largest step accepted, so an odd system value cannot make the cascade useless.</summary>
    Public Const MinStep As Integer = 8
    Public Const MaxStep As Integer = 96

    ''' <summary>How close two corners have to be before they count as the same spot.</summary>
    Public Const SameSpot As Integer = 2

    ''' <summary>How many times the search will step aside before it settles for an overlap.</summary>
    Private Const MaxTries As Integer = 32

    ''' <summary>
    ''' The cascade step for a title bar of this height. The step wants to be about the height of a
    ''' title bar, so that each window's own title stays visible above the one in front of it; taking
    ''' it from the system is also what makes the cascade follow the screen's scaling instead of
    ''' being 26 pixels of nothing on a 4K display.
    ''' </summary>
    Public Shared Function StepFor(titleBarHeight As Integer) As Integer
        If titleBarHeight < MinStep Then Return DefaultStep
        If titleBarHeight > MaxStep Then Return MaxStep
        Return titleBarHeight
    End Function

    ''' <summary>
    ''' Where to put a window of this size.
    ''' </summary>
    ''' <param name="hasAnchor">
    ''' True when there is a frontmost window to cascade from. False for the first window of a run, and
    ''' also when the frontmost window is maximized — there is no corner worth stepping from then, so
    ''' the cascade starts at the top left of the working area instead.
    ''' </param>
    ''' <param name="anchorX">The frontmost window's left edge. Ignored when hasAnchor is False.</param>
    ''' <param name="anchorY">The frontmost window's top edge. Ignored when hasAnchor is False.</param>
    ''' <param name="width">The new window's full outside width.</param>
    ''' <param name="height">The new window's full outside height.</param>
    ''' <param name="areaX">Left of the visible frame: the working area of the anchor's screen.</param>
    ''' <param name="areaY">Top of the visible frame.</param>
    ''' <param name="areaWidth">Width of the visible frame.</param>
    ''' <param name="areaHeight">Height of the visible frame.</param>
    ''' <param name="stepBy">The cascade step, from StepFor.</param>
    ''' <param name="taken">
    ''' The top left corners of the main windows already open, the new one excluded. A spot one of them
    ''' is on is stepped past. Nothing or empty is fine.
    ''' </param>
    Public Shared Function Place(hasAnchor As Boolean,
                                 anchorX As Integer, anchorY As Integer,
                                 width As Integer, height As Integer,
                                 areaX As Integer, areaY As Integer,
                                 areaWidth As Integer, areaHeight As Integer,
                                 stepBy As Integer,
                                 taken As IEnumerable(Of WindowSpot)) As WindowSpot

        If stepBy < MinStep OrElse stepBy > MaxStep Then stepBy = DefaultStep
        If width < 1 Then width = 1
        If height < 1 Then height = 1

        Dim right As Integer = areaX + Math.Max(1, areaWidth)
        Dim bottom As Integer = areaY + Math.Max(1, areaHeight)

        Dim spots As List(Of WindowSpot) = Nothing
        If taken IsNot Nothing Then
            spots = New List(Of WindowSpot)(taken)
        End If

        Dim x As Integer
        Dim y As Integer
        If hasAnchor Then
            x = anchorX + stepBy
            y = anchorY + stepBy
        Else
            x = areaX
            y = areaY
        End If

        ' Step down and right until the window both fits and is not sitting exactly on another one.
        ' One wrap back to the top left corner is allowed; after that the clamp below has the last word.
        Dim wrapped As Boolean = Not hasAnchor
        For tries As Integer = 0 To MaxTries - 1
            If x + width > right OrElse y + height > bottom Then
                If wrapped Then Exit For
                x = areaX
                y = areaY
                wrapped = True
            End If
            If Not Occupied(x, y, spots) Then Exit For
            x += stepBy
            y += stepBy
        Next

        ' Right and bottom first, then left and top, so the top left corner always wins: it is the
        ' corner that carries the title bar the user needs in order to drag the window.
        If x + width > right Then x = right - width
        If y + height > bottom Then y = bottom - height
        If x < areaX Then x = areaX
        If y < areaY Then y = areaY

        Return New WindowSpot(x, y)
    End Function

    ''' <summary>True when one of the windows already open has its corner at this spot.</summary>
    Private Shared Function Occupied(x As Integer, y As Integer, spots As List(Of WindowSpot)) As Boolean
        If spots Is Nothing Then Return False
        For Each s As WindowSpot In spots
            If Math.Abs(s.X - x) <= SameSpot AndAlso Math.Abs(s.Y - y) <= SameSpot Then Return True
        Next
        Return False
    End Function

End Class
