' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  CpmFormatForm.vb — Disk → CP/M layout: name the layout a CP/M floppy is read
'  and written with ("--cpm-format NAME", trs80-tool 5.2.0), or hand the choice
'  back to trs80-tool. The list of layouts is trs80-tool's own text, shown as it
'  prints it; nothing in it is interpreted here.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Public Class CpmFormatForm
    Inherits Form

    Private ReadOnly _txtName As New TextBox()
    Private _chosen As String = ""

    ''' <summary>The layout name the user gave, or "" to let trs80-tool decide. Valid after DialogResult.OK.</summary>
    Public ReadOnly Property Chosen As String
        Get
            Return _chosen
        End Get
    End Property

    ''' <param name="layoutsText">What "disk IMAGE help cpm-format" printed.</param>
    ''' <param name="current">The layout in force for the image, or "" when trs80-tool decides.</param>
    Public Sub New(layoutsText As String, current As String)
        Text = "CP/M layout"
        FormBorderStyle = FormBorderStyle.Sizable
        MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False
        ClientSize = New Size(860, 520)
        MinimumSize = New Size(560, 320)

        Dim lblTop As New Label() With {
            .AutoSize = False, .Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(8, 8, 8, 0),
            .Text = "A CP/M disk doesn't say how it is laid out, so trs80-tool works it out. If it picks the wrong layout for this disk, name the right one from trs80-tool's list below." &
                    If(String.IsNullOrEmpty(current), "", "  Now in force: " & current & ".")}

        Dim txtList As New TextBox() With {
            .Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Both, .WordWrap = False, .Dock = DockStyle.Fill,
            .Font = New Font(HostPlatform.MonospaceFontName, 9.0F),
            .Text = If(layoutsText, "").Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)}

        Dim bottom As New Panel() With {.Dock = DockStyle.Bottom, .Height = 44}
        Dim lblName As New Label() With {.AutoSize = True, .Location = New Point(10, 14), .Text = "Layout name (for example mm-ds-data):"}
        _txtName.Bounds = New Rectangle(250, 10, 150, 23)
        _txtName.Text = If(current, "")
        Dim btnUse As New Button() With {.Text = "Use this layout", .Bounds = New Rectangle(410, 9, 120, 26)}
        Dim btnAuto As New Button() With {.Text = "Let trs80-tool decide", .Bounds = New Rectangle(536, 9, 150, 26)}
        Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(692, 9, 80, 26)}
        AddHandler btnUse.Click, Sub()
                                     Dim name As String = _txtName.Text.Trim()
                                     If name.Length = 0 OrElse name.IndexOf(" "c) >= 0 Then
                                         MessageBox.Show(Me, "Type one layout name from the list, such as mm-data.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                                         Return
                                     End If
                                     _chosen = name
                                     DialogResult = DialogResult.OK
                                 End Sub
        AddHandler btnAuto.Click, Sub()
                                      _chosen = ""
                                      DialogResult = DialogResult.OK
                                  End Sub
        bottom.Controls.AddRange(New Control() {lblName, _txtName, btnUse, btnAuto, btnCancel})

        ' Fill first, so that the docked edges take their space from it.
        Controls.Add(txtList)
        Controls.Add(bottom)
        Controls.Add(lblTop)
        AcceptButton = btnUse
        CancelButton = btnCancel
    End Sub
End Class
