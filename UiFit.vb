' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  UiFit.vb — the Windows Forms windows are laid out at fixed pixel positions,
'  which is right for the English they were drawn around and wrong for a
'  translation: French, Spanish and Italian run 15 to 30 percent longer, so a
'  label reaches the box beside it, a sentence runs past the window's edge and a
'  button's text is cut. Each window calls the routines here at the end of
'  building itself, to make room for whatever language is in force.
'
'  The rule every routine keeps: it only ever makes something WIDER or TALLER,
'  and only when the text really needs it. Nothing is ever made smaller, and
'  nothing moves up or left — the buttons of one row are the single exception,
'  and they move only when one of them has had to grow. So in English, where
'  every string is the one the window was drawn around, every call here does
'  nothing at all and the window is pixel for pixel what it has always been.
'  That is what makes this safe to apply to every window without being able to
'  look at the result.
'
'  Widths are measured with TextRenderer, which is what Windows Forms draws
'  label and button text with, so the measurement matches what is painted.
'  Call these at the END of BuildUi, once the controls are in their parents and
'  have the font they will really be drawn in.
'
'  The Eto program needs none of this: it lays itself out and sizes to its text.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>Makes room in a fixed-pixel window for a longer language. Windows Forms only.</summary>
Friend Module UiFit

    ''' <summary>A need within this many pixels of what is there already is left alone.</summary>
    Private Const Slack As Integer = 2

    ''' <summary>The width this control's text needs, in the font it will be drawn in.</summary>
    Friend Function NeedWidth(c As Control, text As String) As Integer
        If c Is Nothing OrElse String.IsNullOrEmpty(text) Then Return 0
        Return TextRenderer.MeasureText(text, c.Font).Width
    End Function

    ''' <summary>
    ''' How far right a control reaches. A label, check box or radio button with AutoSize on has
    ''' already been measured by Windows Forms itself, text, glyph and all, and this codebase
    ''' relies on that at build time (PreferredWidth is read in several windows), so its own Right
    ''' is the truth. A control given a fixed box keeps that box: a note meant to wrap is meant to
    ''' be narrow, and FitLabel, FitButton and FitColumn are the way to say that a fixed box must
    ''' not cut its text.
    ''' </summary>
    Friend Function RealRight(c As Control) As Integer
        If c Is Nothing Then Return 0
        Return c.Right
    End Function

    ''' <summary>
    ''' A label with a control beside it on the same row. Moves the control right when the label's
    ''' text reaches it, and gives back how far it moved so the caller can widen what holds it.
    ''' </summary>
    Friend Function After(lbl As Control, ctrl As Control, Optional gap As Integer = 8) As Integer
        If lbl Is Nothing OrElse ctrl Is Nothing Then Return 0
        Dim want As Integer = RealRight(lbl) + gap
        If want <= ctrl.Left + Slack Then Return 0
        Dim moved As Integer = want - ctrl.Left
        ctrl.Left = want
        Return moved
    End Function

    ''' <summary>
    ''' The same, where everything to the right of the label has to travel with the control: the
    ''' control first, then whatever else shares the row.
    ''' </summary>
    Friend Function AfterAll(lbl As Control, gap As Integer, ParamArray row As Control()) As Integer
        If row Is Nothing OrElse row.Length = 0 Then Return 0
        Dim moved As Integer = After(lbl, row(0), gap)
        If moved > 0 Then
            For i As Integer = 1 To row.Length - 1
                If row(i) IsNot Nothing Then row(i).Left += moved
            Next
        End If
        Return moved
    End Function

    ''' <summary>
    ''' Controls that sit side by side in one row of the same parent, given left to right: each is
    ''' pushed right until it clears the one before it. Only ever moves a control right, so a row
    ''' whose English text already fits does not move.
    ''' </summary>
    Friend Sub SpreadRow(gap As Integer, ParamArray row As Control())
        If row Is Nothing Then Return
        For i As Integer = 1 To row.Length - 1
            If row(i) Is Nothing OrElse row(i - 1) Is Nothing Then Continue For
            Dim want As Integer = RealRight(row(i - 1)) + gap
            If row(i).Left < want Then row(i).Left = want
        Next
    End Sub

    ''' <summary>
    ''' A column of label-and-value rows. Every value is moved to clear the widest label, and never
    ''' to the left of where it is now, so the column stays aligned in any language.
    ''' </summary>
    Friend Sub AlignColumn(rows As IEnumerable(Of KeyValuePair(Of Control, Control)), Optional gap As Integer = 10)
        If rows Is Nothing Then Return
        Dim x As Integer = 0
        For Each r As KeyValuePair(Of Control, Control) In rows
            If r.Key Is Nothing OrElse r.Value Is Nothing Then Continue For
            Dim want As Integer = RealRight(r.Key) + gap
            If want > x Then x = want
            If r.Value.Left > x Then x = r.Value.Left
        Next
        For Each r As KeyValuePair(Of Control, Control) In rows
            If r.Value Is Nothing Then Continue For
            If r.Value.Left < x Then r.Value.Left = x
        Next
    End Sub

    ''' <summary>Widens a label that is not auto-sized so its text is not cut.</summary>
    Friend Sub FitLabel(lbl As Label, Optional extra As Integer = 6)
        If lbl Is Nothing OrElse lbl.AutoSize Then Return
        Dim need As Integer = NeedWidth(lbl, lbl.Text) + extra
        If lbl.Width < need Then lbl.Width = need
    End Sub

    ''' <summary>
    ''' Grows a label that wraps its text (AutoSize off) so every line fits in the height it is
    ''' given, at the width it is given. Only ever makes it taller.
    ''' </summary>
    Friend Sub FitWrapped(lbl As Label, width As Integer)
        If lbl Is Nothing OrElse lbl.AutoSize Then Return
        Dim avail As Integer = width - lbl.Padding.Horizontal
        If avail < 40 Then Return
        Dim need As Integer = TextRenderer.MeasureText(lbl.Text, lbl.Font,
                                                       New Size(avail, Integer.MaxValue),
                                                       TextFormatFlags.WordBreak).Height + lbl.Padding.Vertical + 4
        If lbl.Height < need Then lbl.Height = need
    End Sub

    ''' <summary>Widens a button so its own text fits. Never narrows it.</summary>
    Friend Sub FitButton(b As Button, Optional padding As Integer = 22)
        If b Is Nothing Then Return
        Dim need As Integer = NeedWidth(b, b.Text) + padding
        If b.Width < need Then b.Width = need
    End Sub

    ''' <summary>
    ''' A row of buttons along the bottom right. Each is sized to its text, then the row is laid
    ''' out from rightEdge leftwards in the order given. In English nothing grows, so nothing moves.
    ''' </summary>
    Friend Sub ButtonRow(rightEdge As Integer, gap As Integer, ParamArray buttons As Button())
        If buttons Is Nothing Then Return
        For Each b As Button In buttons
            FitButton(b)
        Next
        Dim x As Integer = rightEdge
        For i As Integer = buttons.Length - 1 To 0 Step -1
            If buttons(i) Is Nothing Then Continue For
            x -= buttons(i).Width
            buttons(i).Left = x
            x -= gap
        Next
    End Sub

    ''' <summary>
    ''' A row of buttons along the bottom left, each sized to its text, laid out left to right.
    ''' </summary>
    Friend Sub ButtonRowLeft(leftEdge As Integer, gap As Integer, ParamArray buttons As Button())
        If buttons Is Nothing Then Return
        Dim x As Integer = leftEdge
        For Each b As Button In buttons
            If b Is Nothing Then Continue For
            FitButton(b)
            b.Left = x
            x = b.Right + gap
        Next
    End Sub

    ''' <summary>
    ''' Widens a group box so its caption fits on one line. A caption too long for the box wraps,
    ''' and the second line is painted over whatever sits at the top of the box.
    ''' </summary>
    Friend Sub FitCaption(box As GroupBox, Optional extra As Integer = 18)
        If box Is Nothing Then Return
        Dim need As Integer = NeedWidth(box, box.Text) + extra
        If box.Width < need Then box.Width = need
    End Sub

    ''' <summary>Widens a container so everything inside it fits, with a margin on the right.</summary>
    Friend Sub GrowToFit(box As Control, Optional margin As Integer = 10)
        If box Is Nothing Then Return
        Dim need As Integer = 0
        For Each c As Control In box.Controls
            Dim r As Integer = RealRight(c)
            If r > need Then need = r
        Next
        ' Only when something is really cut: a box whose English content already sits snugly
        ' inside it must not be nudged wider by the margin.
        Dim gb As GroupBox = TryCast(box, GroupBox)
        If need > 0 Then
            If gb IsNot Nothing Then
                If need > gb.ClientSize.Width Then gb.Width += need + margin - gb.ClientSize.Width
            ElseIf need > box.ClientSize.Width Then
                box.Width += need + margin - box.ClientSize.Width
            End If
        End If
        If gb IsNot Nothing Then FitCaption(gb)
    End Sub

    ''' <summary>
    ''' Widens a form so everything on it fits. Call it after the group boxes have been grown, and
    ''' lay the buttons out afterwards so they stay against the new right edge.
    ''' </summary>
    Friend Sub GrowForm(f As Form, Optional margin As Integer = 12)
        If f Is Nothing Then Return
        Dim need As Integer = 0
        For Each c As Control In f.Controls
            If c.Dock <> DockStyle.None Then Continue For
            Dim r As Integer = RealRight(c)
            If r > need Then need = r
        Next
        If need = 0 Then Return
        ' Only when something is really cut (see GrowToFit).
        If need > f.ClientSize.Width Then
            f.ClientSize = New Size(need + margin, f.ClientSize.Height)
        End If
    End Sub

    ''' <summary>Widens a grid column so its heading is not cut. Never narrows it.</summary>
    Friend Sub FitColumn(grid As DataGridView, index As Integer, Optional extra As Integer = 26)
        If grid Is Nothing OrElse index < 0 OrElse index >= grid.Columns.Count Then Return
        Dim col As DataGridViewColumn = grid.Columns(index)
        Dim f As Font = grid.ColumnHeadersDefaultCellStyle.Font
        If f Is Nothing Then f = grid.Font
        If f Is Nothing Then Return
        Dim need As Integer = TextRenderer.MeasureText(col.HeaderText, f).Width + extra
        If col.Width < need Then col.Width = need
    End Sub

    ''' <summary>Widens every column of a grid so none of the headings is cut.</summary>
    Friend Sub FitColumns(grid As DataGridView, Optional extra As Integer = 26)
        If grid Is Nothing Then Return
        For i As Integer = 0 To grid.Columns.Count - 1
            FitColumn(grid, i, extra)
        Next
    End Sub

    ''' <summary>Widens a list-view column so its heading is not cut. Never narrows it.</summary>
    Friend Sub FitListColumn(list As ListView, index As Integer, Optional extra As Integer = 20)
        If list Is Nothing OrElse index < 0 OrElse index >= list.Columns.Count Then Return
        Dim col As ColumnHeader = list.Columns(index)
        Dim need As Integer = TextRenderer.MeasureText(col.Text, list.Font).Width + extra
        If col.Width < need Then col.Width = need
    End Sub

    ''' <summary>Widens every column of a list view so none of the headings is cut.</summary>
    Friend Sub FitListColumns(list As ListView, Optional extra As Integer = 20)
        If list Is Nothing Then Return
        For i As Integer = 0 To list.Columns.Count - 1
            FitListColumn(list, i, extra)
        Next
    End Sub

    ' ================================================================= placing a window

    ''' <summary>
    ''' Show a window over the one that opened it, and make sure it lands on a screen.
    '''
    ''' Every window in this program sets StartPosition = CenterParent, but Windows Forms honours
    ''' that ONLY for ShowDialog. A modeless Show() ignores it and falls back to the cascade
    ''' position, which on a desktop with more than one monitor can put the window partly or
    ''' wholly off the screen - and at a negative X when a monitor sits to the left of the main
    ''' one, where it cannot be reached at all. MainForm never suffered from this because it
    ''' checks its remembered position against the screens before using it; no window opened
    ''' from it did that, which is why the main window looked fine and a report window went
    ''' missing. Reported by Pascal Holdry, 25 November 2026.
    '''
    ''' So the placing is done here instead: centre on the owner, then pull the window back
    ''' inside the working area of the screen the owner is on. A window bigger than that screen
    ''' is left at its top left rather than being shrunk, because its own minimum size is its
    ''' business and a window that runs off the bottom but can be dragged beats one that cannot.
    ''' </summary>
    Friend Sub ShowOver(f As Form, owner As Form)
        If f Is Nothing Then Return
        If owner Is Nothing OrElse owner.IsDisposed Then
            f.Show()
            Return
        End If
        Try
            Dim area As Rectangle = Screen.FromControl(owner).WorkingArea
            Dim w As Integer = f.Width
            Dim h As Integer = f.Height
            Dim x As Integer = owner.Left + (owner.Width - w) \ 2
            Dim y As Integer = owner.Top + (owner.Height - h) \ 2
            ' Right and bottom first, then left and top, so the top left corner always wins:
            ' it is the corner that carries the title bar the user needs in order to drag it.
            If x + w > area.Right Then x = area.Right - w
            If y + h > area.Bottom Then y = area.Bottom - h
            If x < area.Left Then x = area.Left
            If y < area.Top Then y = area.Top
            f.StartPosition = FormStartPosition.Manual
            f.Location = New Point(x, y)
        Catch
            ' Any trouble reading the screens, and the window still has to appear somewhere sane.
            f.StartPosition = FormStartPosition.CenterScreen
        End Try
        f.Show(owner)
    End Sub

    ' ================================================================= reading a report

    ''' <summary>Smallest and largest the reports go: below 6 nothing is readable, above 36 a line fills the window.</summary>
    Private Const MinPoints As Single = 6.0F
    Private Const MaxPoints As Single = 36.0F

    ''' <summary>
    ''' Let Ctrl with the mouse wheel, and Ctrl with + - and 0, change the size of a monospaced
    ''' report box. On a high-resolution screen the 9.5 point the reports are drawn in is very
    ''' small, and copying the text into a word processor to read it is a poor answer. Asked for
    ''' by Pascal Holdry, 25 November 2026.
    '''
    ''' Only the size changes and the family is kept, because every one of these reports lines
    ''' its columns up by character width and a proportional font would ruin all of them.
    ''' </summary>
    Friend Sub EnableTextZoom(box As TextBoxBase)
        If box Is Nothing Then Return
        Dim original As Single = box.Font.Size
        AddHandler box.MouseWheel,
            Sub(sender As Object, e As MouseEventArgs)
                If (Control.ModifierKeys And Keys.Control) = 0 Then Return
                ZoomBy(box, If(e.Delta > 0, 1.0F, -1.0F))
            End Sub
        AddHandler box.KeyDown,
            Sub(sender As Object, e As KeyEventArgs)
                If Not e.Control Then Return
                Select Case e.KeyCode
                    Case Keys.Oemplus, Keys.Add
                        ZoomBy(box, 1.0F) : e.Handled = True : e.SuppressKeyPress = True
                    Case Keys.OemMinus, Keys.Subtract
                        ZoomBy(box, -1.0F) : e.Handled = True : e.SuppressKeyPress = True
                    Case Keys.D0, Keys.NumPad0
                        SetSize(box, original) : e.Handled = True : e.SuppressKeyPress = True
                End Select
            End Sub
    End Sub

    Private Sub ZoomBy(box As TextBoxBase, points As Single)
        SetSize(box, box.Font.Size + points)
    End Sub

    Private Sub SetSize(box As TextBoxBase, points As Single)
        Dim want As Single = Math.Max(MinPoints, Math.Min(MaxPoints, points))
        If Math.Abs(want - box.Font.Size) < 0.01F Then Return
        Dim old As Font = box.Font
        box.Font = New Font(old.FontFamily, want, old.Style)
        old.Dispose()
    End Sub

End Module
