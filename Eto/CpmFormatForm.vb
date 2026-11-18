' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  CpmFormatForm.vb — Disk → CP/M layout in the Eto.Forms program: shows
'  trs80-tool's own list of CP/M layouts and takes the name of the one this
'  disk is to be read and written with (trs80-tool's --cpm-format), or hands
'  the choice back to trs80-tool.
' ============================================================================
Option Strict On
Option Explicit On

Imports Eto.Drawing
Imports Eto.Forms

Public Class CpmFormatForm
    Inherits Dialog(Of Boolean)

    Private ReadOnly _txtName As New TextBox With {.Width = 160}
    Private _chosen As String = ""

    ''' <summary>The layout name to use; an empty string hands the choice back to trs80-tool.</summary>
    Public ReadOnly Property Chosen As String
        Get
            Return _chosen
        End Get
    End Property

    Public Sub New(layoutsText As String, current As String)
        Title = "CP/M layout"
        Resizable = True
        ClientSize = New Size(880, 520)
        MinimumSize = New Size(560, 320)
        Padding = New Padding(Ui.Margin)
        Dim top As Label = Ui.Wrapped("A CP/M disk doesn't say how it is laid out, so trs80-tool works it out. If it picks the wrong layout for this disk, name the right one from trs80-tool's list below." &
                                      If(String.IsNullOrEmpty(current), "", "  Now in force: " & current & "."), 840)
        Dim list As TextArea = Ui.ReportBox(layoutsText, 9.0F)
        _txtName.Text = If(current, "")
        Dim btnUse As New Button With {.Text = "Use this layout"}
        Dim btnAuto As New Button With {.Text = "Let trs80-tool decide"}
        Dim btnCancel As New Button With {.Text = "Cancel"}
        AddHandler btnUse.Click, Sub()
                                     Dim name As String = _txtName.Text.Trim()
                                     If name.Length = 0 OrElse name.IndexOf(" "c) >= 0 Then
                                         Ui.Info(Me, "Type one layout name from the list, such as mm-data.", Title)
                                         Return
                                     End If
                                     _chosen = name
                                     Close(True)
                                 End Sub
        AddHandler btnAuto.Click, Sub()
                                      _chosen = ""
                                      Close(True)
                                  End Sub
        AddHandler btnCancel.Click, Sub() Close(False)
        Content = Ui.Down(list, top, list, Ui.Across(New Label With {.Text = "Layout name (for example mm-ds-data):", .VerticalAlignment = VerticalAlignment.Center}, _txtName, btnUse, btnAuto, Nothing, btnCancel))
        DefaultButton = btnUse
        AbortButton = btnCancel
    End Sub

    Public Function ShowOk(parent As Control) As Boolean
        Return ShowModal(parent)
    End Function
End Class
