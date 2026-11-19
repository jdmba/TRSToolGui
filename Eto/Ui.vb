' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  Ui.vb — small helpers shared by every window of the Eto.Forms program:
'  message boxes, file-dialog filters, table and stack layouts, the base class
'  of the OK / Cancel dialogs, and the clipboard formats for files.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Eto.Drawing
Imports Eto.Forms

Public Module Ui

    ' ------------------------------------------------------------------ text

    ''' <summary>Text with LF line ends, which is what Eto's labels and text areas take on every platform.</summary>
    Public Function Lf(text As String) As String
        Return If(text, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
    End Function

    ''' <summary>Text with the line ends of the system the program runs on, for a file that is saved.</summary>
    Public Function HostLines(text As String) As String
        Return Lf(text).Replace(vbLf, Environment.NewLine)
    End Function

    Public Function Mono(Optional size As Single = 10.0F) As Font
        Return Fonts.Monospace(size)
    End Function

    Public ReadOnly Property GrayText As Color
        Get
            Return SystemColors.DisabledText
        End Get
    End Property

    ' ------------------------------------------------------------------ message boxes

    Public Sub Info(parent As Control, text As String, caption As String)
        MessageBox.Show(parent, Lf(text), caption, MessageBoxButtons.OK, MessageBoxType.Information)
    End Sub

    Public Sub Warn(parent As Control, text As String, caption As String)
        MessageBox.Show(parent, Lf(text), caption, MessageBoxButtons.OK, MessageBoxType.Warning)
    End Sub

    Public Sub Fail(parent As Control, text As String, caption As String)
        MessageBox.Show(parent, Lf(text), caption, MessageBoxButtons.OK, MessageBoxType.Error)
    End Sub

    ''' <summary>A Yes / No question. With defaultNo, Enter answers No.</summary>
    Public Function Ask(parent As Control, text As String, caption As String, Optional type As MessageBoxType = MessageBoxType.Question, Optional defaultNo As Boolean = False) As Boolean
        Return MessageBox.Show(parent, Lf(text), caption, MessageBoxButtons.YesNo, type, If(defaultNo, MessageBoxDefaultButton.No, MessageBoxDefaultButton.Yes)) = DialogResult.Yes
    End Function

    Public Function AskCancel(parent As Control, text As String, caption As String, Optional type As MessageBoxType = MessageBoxType.Warning) As DialogResult
        Return MessageBox.Show(parent, Lf(text), caption, MessageBoxButtons.YesNoCancel, type)
    End Function

    ' ------------------------------------------------------------------ file dialogs

    ''' <summary>
    ''' Fill a file dialog's filter list from the "Description|*.a;*.b|Description|*.c" form. On a system whose
    ''' file names are case-sensitive each pattern is offered in lower and upper case, so IMAGE.DMK is listed
    ''' under "*.dmk". A Mac's file dialog goes by the last part of a name only, so for a two-part pattern
    ''' such as "*.dmk.gz" the last part (".gz") is offered as well, or GAME.DMK.GZ could not be chosen.
    ''' </summary>
    Public Sub SetFilters(dlg As FileDialog, filter As String)
        dlg.Filters.Clear()
        Dim parts As String() = filter.Split("|"c)
        For i As Integer = 0 To parts.Length - 2 Step 2
            Dim exts As New List(Of String)
            For Each pattern In parts(i + 1).Split(";"c)
                Dim p As String = pattern.Trim()
                If p = "*.*" OrElse p = "*" Then
                    exts.Add("")
                ElseIf p.StartsWith("*", StringComparison.Ordinal) Then
                    Dim ext As String = p.Substring(1)
                    If HostPlatform.PathComparison = StringComparison.Ordinal Then
                        If Not exts.Contains(ext.ToLowerInvariant()) Then exts.Add(ext.ToLowerInvariant())
                        If Not exts.Contains(ext.ToUpperInvariant()) Then exts.Add(ext.ToUpperInvariant())
                    ElseIf Not exts.Contains(ext) Then
                        exts.Add(ext)
                    End If
                    If HostPlatform.IsMacOS AndAlso ext.LastIndexOf("."c) > 0 Then
                        Dim lastPart As String = ext.Substring(ext.LastIndexOf("."c))
                        If Not exts.Contains(lastPart) Then exts.Add(lastPart)
                    End If
                End If
            Next
            dlg.Filters.Add(New FileFilter(parts(i), exts.ToArray()))
        Next
        If dlg.Filters.Count > 0 Then dlg.CurrentFilterIndex = 0
    End Sub

    ''' <summary>Ask for one existing file; Nothing when the user cancels.</summary>
    Public Function PickFile(parent As Control, title As String, filter As String, Optional startName As String = Nothing) As String
        Using dlg As New OpenFileDialog()
            dlg.Title = title
            SetFilters(dlg, filter)
            dlg.CheckFileExists = True
            SetStart(dlg, startName)
            If dlg.ShowDialog(parent) <> DialogResult.Ok Then Return Nothing
            Return dlg.FileName
        End Using
    End Function

    ''' <summary>
    ''' Ask where to save; Nothing when the user cancels. The save dialogs of GTK don't add an extension, so
    ''' defaultExt (without the dot) is added when the name has none.
    ''' </summary>
    Public Function PickSave(parent As Control, title As String, filter As String, Optional startName As String = Nothing, Optional defaultExt As String = Nothing) As String
        Using dlg As New SaveFileDialog()
            dlg.Title = title
            SetFilters(dlg, filter)
            SetStart(dlg, startName)
            If dlg.ShowDialog(parent) <> DialogResult.Ok Then Return Nothing
            Dim name As String = dlg.FileName
            If String.IsNullOrEmpty(name) Then Return Nothing
            If Not String.IsNullOrEmpty(defaultExt) AndAlso Path.GetExtension(name).Length = 0 Then name &= "." & defaultExt
            Return name
        End Using
    End Function

    Private Sub SetStart(dlg As FileDialog, startName As String)
        If String.IsNullOrEmpty(startName) Then Return
        Try
            Dim folder As String = Path.GetDirectoryName(startName)
            If Not String.IsNullOrEmpty(folder) AndAlso Directory.Exists(folder) Then dlg.Directory = New Uri(folder)
            dlg.FileName = If(TypeOf dlg Is SaveFileDialog, Path.GetFileName(startName), startName)
        Catch
        End Try
    End Sub

    Public Function PickFolder(parent As Control, title As String, startFolder As String) As String
        Using dlg As New SelectFolderDialog()
            dlg.Title = title
            Try
                If Not String.IsNullOrEmpty(startFolder) AndAlso Directory.Exists(startFolder) Then dlg.Directory = startFolder
            Catch
            End Try
            If dlg.ShowDialog(parent) <> DialogResult.Ok Then Return Nothing
            Return dlg.Directory
        End Using
    End Function

    ' ------------------------------------------------------------------ layout

    Public Const Gap As Integer = 6
    Public Const Margin As Integer = 10

    ''' <summary>Controls side by side. A Nothing in the list is a gap that takes the spare width.</summary>
    Public Function Across(ParamArray items As Control()) As TableLayout
        Dim cells As New List(Of TableCell)
        Dim hasSpring As Boolean = False
        For Each c In items
            If c Is Nothing Then
                cells.Add(New TableCell(Nothing, True))
                hasSpring = True
            Else
                cells.Add(New TableCell(c))
            End If
        Next
        If Not hasSpring Then cells.Add(New TableCell(Nothing, True))
        Dim t As New TableLayout With {.Spacing = New Size(Gap, Gap)}
        t.Rows.Add(New TableRow(cells.ToArray()))
        Return t
    End Function

    ''' <summary>Controls one under the other, each as wide as the column. The one named as stretch takes the spare height.</summary>
    Public Function Down(stretch As Control, ParamArray items As Control()) As TableLayout
        Dim t As New TableLayout With {.Spacing = New Size(Gap, Gap)}
        For Each c In items
            t.Rows.Add(New TableRow(New TableCell(c, True)) With {.ScaleHeight = (c IsNot Nothing AndAlso c Is stretch)})
        Next
        If stretch Is Nothing Then t.Rows.Add(New TableRow With {.ScaleHeight = True})
        Return t
    End Function

    ''' <summary>A two-column form: a label on the left of each control. Pass label text and control in turn.</summary>
    Public Function Fields(ParamArray labelsAndControls As Object()) As TableLayout
        Dim t As New TableLayout With {.Spacing = New Size(Gap + 2, Gap)}
        For i As Integer = 0 To labelsAndControls.Length - 2 Step 2
            Dim left As Control = TryCast(labelsAndControls(i), Control)
            If left Is Nothing Then left = New Label With {.Text = CStr(labelsAndControls(i)), .VerticalAlignment = VerticalAlignment.Center}
            t.Rows.Add(New TableRow(New TableCell(left), New TableCell(DirectCast(labelsAndControls(i + 1), Control), True)))
        Next
        Return t
    End Function

    Public Function Group(title As String, content As Control) As GroupBox
        Return New GroupBox With {.Text = title, .Content = content, .Padding = New Padding(8)}
    End Function

    Public Function Lbl(value As String, Optional gray As Boolean = False) As Label
        Dim l As New Label With {.Text = Lf(value)}
        If gray Then l.TextColor = GrayText
        Return l
    End Function

    Public Function Wrapped(value As String, width As Integer) As Label
        Return New Label With {.Text = Lf(value), .Wrap = WrapMode.Word, .Width = width}
    End Function

    Public Function Btn(caption As String, handler As Action) As Button
        Dim b As New Button With {.Text = caption}
        AddHandler b.Click, Sub() handler()
        Return b
    End Function

    ''' <summary>A read-only text box for a report: monospaced, no wrapping, LF line ends.</summary>
    Public Function ReportBox(text As String, Optional size As Single = 10.0F) As TextArea
        Return New TextArea With {.ReadOnly = True, .Wrap = False, .Font = Mono(size), .Text = Lf(text)}
    End Function

    Public Function IsOn(box As CheckBox) As Boolean
        Return box.Checked.HasValue AndAlso box.Checked.Value
    End Function

    ' ------------------------------------------------------------------ the clipboard and files

    ''' <summary>The clipboard format GNOME's and several other Linux file managers use for copied files.</summary>
    Private Const GnomeCopiedFiles As String = "x-special/gnome-copied-files"

    ''' <summary>Put files on the clipboard the way a file manager does when files are copied.</summary>
    Public Sub CopyFilesToClipboard(paths As IEnumerable(Of String))
        Dim uris As New List(Of Uri)
        For Each p In paths
            uris.Add(New Uri(p))
        Next
        Dim cb As Clipboard = Clipboard.Instance
        cb.Clear()
        cb.Uris = uris.ToArray()
        If HostPlatform.IsLinux Then
            Dim sb As New StringBuilder("copy")
            For Each u In uris
                sb.Append(vbLf).Append(u.AbsoluteUri)
            Next
            Try
                cb.SetString(sb.ToString(), GnomeCopiedFiles)
            Catch
            End Try
        End If
    End Sub

    ''' <summary>The files on the clipboard, if a file manager (or this program) put any there.</summary>
    Public Function FilesOnClipboard() As List(Of String)
        Dim result As New List(Of String)
        Try
            Dim cb As Clipboard = Clipboard.Instance
            If cb.ContainsUris Then
                For Each u In cb.Uris
                    If u IsNot Nothing AndAlso u.IsFile Then result.Add(u.LocalPath)
                Next
            End If
            If result.Count = 0 AndAlso HostPlatform.IsLinux AndAlso cb.Contains(GnomeCopiedFiles) Then
                For Each line In Lf(cb.GetString(GnomeCopiedFiles)).Split(ControlChars.Lf)
                    If line.StartsWith("file://", StringComparison.Ordinal) Then result.Add(New Uri(line.Trim()).LocalPath)
                Next
            End If
        Catch
        End Try
        Return result
    End Function

    Public Sub CopyTextToClipboard(text As String)
        Try
            Clipboard.Instance.Text = HostLines(text)
        Catch
        End Try
    End Sub

End Module

''' <summary>
''' The base of the dialogs that end with OK and Cancel. A derived class builds its controls, calls Finish with
''' them, and overrides Accept when OK has something to check. ShowOk returns True when the user chose OK.
''' </summary>
Public Class OkCancelDialog
    Inherits Dialog(Of Boolean)

    Protected ReadOnly OkButton As New Button With {.Text = "OK"}
    Protected ReadOnly CancelButtonControl As New Button With {.Text = "Cancel"}

    Public Sub New(title As String)
        Me.Title = title
        Resizable = False
        Padding = New Padding(Ui.Margin)
        AddHandler OkButton.Click, Sub()
                                       If Accept() Then Close(True)
                                   End Sub
        AddHandler CancelButtonControl.Click, Sub() Close(False)
    End Sub

    ''' <summary>Set the dialog's content and its buttons. okText renames the OK button; noCancel leaves Cancel out.</summary>
    Protected Sub Finish(body As Control, Optional okText As String = Nothing, Optional noCancel As Boolean = False)
        If okText IsNot Nothing Then OkButton.Text = okText
        Content = body
        PositiveButtons.Add(OkButton)
        DefaultButton = OkButton
        If noCancel Then
            AbortButton = OkButton
        Else
            NegativeButtons.Add(CancelButtonControl)
            AbortButton = CancelButtonControl
        End If
    End Sub

    ''' <summary>Called when OK is pressed. Return False to keep the dialog open.</summary>
    Protected Overridable Function Accept() As Boolean
        Return True
    End Function

    Public Function ShowOk(parent As Control) As Boolean
        Return ShowModal(parent)
    End Function
End Class

''' <summary>
''' The base of the windows that stay open beside the main window (viewer, reports, maps). ShowOwned keeps the
''' window in front of the one it was opened from, and Esc closes it.
''' </summary>
Public Class ToolWindow
    Inherits Form

    Public Sub ShowOwned(parent As Window)
        Try
            Owner = parent
        Catch
        End Try
        Show()
    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.Key = Keys.Escape AndAlso e.Modifiers = Keys.None Then
            e.Handled = True
            Close()
            Return
        End If
        MyBase.OnKeyDown(e)
    End Sub
End Class
