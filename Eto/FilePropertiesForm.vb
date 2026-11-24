' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  FilePropertiesForm.vb — the File Properties dialog of the Eto.Forms program:
'  name, date, size, kind, protection level, attributes and passwords. OK hands
'  the changes back as "attrib" option words and a new name.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text.RegularExpressions
Imports Eto.Drawing
Imports Eto.Forms

Public Class FilePropertiesForm
    Inherits OkCancelDialog

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _file As DiskFile
    Private ReadOnly _txtName As New TextBox With {.Width = 170}
    Private ReadOnly _txtDate As New TextBox With {.Width = 170}
    Private ReadOnly _chkNoDate As New CheckBox With {.Text = UiText.T("dlg.fileprops.nodate")}
    Private ReadOnly _chkInvisible As New CheckBox With {.Text = UiText.T("dlg.fileprops.invisible")}
    Private ReadOnly _chkSystem As New CheckBox With {.Text = UiText.T("dlg.fileprops.system")}
    Private ReadOnly _cmbProtect As New DropDown With {.Width = 230}
    Private ReadOnly _txtPassword As New TextBox With {.Width = 120, .MaxLength = 8}
    Private ReadOnly _chkSetPassword As New CheckBox With {.Text = UiText.T("dlg.fileprops.setupdate")}
    Private ReadOnly _txtAccess As New TextBox With {.Width = 120, .MaxLength = 8}
    Private ReadOnly _chkSetAccess As New CheckBox With {.Text = UiText.T("dlg.fileprops.setaccess")}
    Private ReadOnly _chkCpmReadOnly As New CheckBox With {.Text = UiText.T("dlg.fileprops.cpmro")}
    Private ReadOnly _chkCpmSys As New CheckBox With {.Text = UiText.T("dlg.fileprops.cpmsys")}
    Private _cpmEditable As Boolean
    Private _nameEditable As Boolean
    Private _attribEditable As Boolean
    Private _hasProtect As Boolean

    ' Protection level names by level number 0..7, in the two vocabularies trs80-tool prints.
    Private Shared ReadOnly LdosNames As String() = {"FULL", "KILL", "RENAME", "UNUSED", "WRITE", "READ", "EXEC", "NO_ACCESS"}
    Private Shared ReadOnly Trsdos6Names As String() = {"FULL", "REMOVE", "RENAME", "WRITE", "UPDATE", "READ", "EXEC", "NO_ACCESS"}

    Public Sub New(disk As DiskImage, file As DiskFile)
        MyBase.New(UiText.T("dlg.fileprops.title"))
        _disk = disk
        _file = file
        BuildUi()
    End Sub

    ''' <summary>The new name, or Nothing when it was not changed (or can't be).</summary>
    Public ReadOnly Property NewName As String
        Get
            If Not _nameEditable Then Return Nothing
            Dim n As String = TextUtil.NormalizeTrsName(_txtName.Text.Trim().ToUpperInvariant())
            If n Is Nothing Then Return Nothing
            If String.Equals(n, _file.TrsName, StringComparison.OrdinalIgnoreCase) Then Return Nothing
            Return n
        End Get
    End Property

    ''' <summary>The words for "disk attrib NAME ...": only what was changed.</summary>
    Public Function AttribOptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _cpmEditable Then Return _file.CpmAttribWords(Ui.IsOn(_chkCpmReadOnly), Ui.IsOn(_chkCpmSys))
        If Not _attribEditable Then Return w
        If Ui.IsOn(_chkInvisible) <> _file.Hidden Then w.Add(If(Ui.IsOn(_chkInvisible), "--invisible", "--visible"))
        If Ui.IsOn(_chkSystem) <> _file.IsSystem Then w.Add(If(Ui.IsOn(_chkSystem), "--system", "--no-system"))
        If _hasProtect AndAlso _cmbProtect.SelectedIndex >= 0 AndAlso _cmbProtect.SelectedIndex <> OriginalProtectIndex() Then
            w.Add("--protect") : w.Add(_cmbProtect.SelectedIndex.ToString(CultureInfo.InvariantCulture))
        End If
        If Ui.IsOn(_chkNoDate) Then
            If _file.DateText.Length > 0 Then w.Add("--no-date")
        Else
            Dim t As String = _txtDate.Text.Trim()
            If t.Length > 0 AndAlso Not String.Equals(t, _file.DateText, StringComparison.Ordinal) Then
                Dim ymd As String = TextUtil.ToYmd(t)
                If ymd IsNot Nothing AndAlso Not String.Equals(ymd, _file.DateText, StringComparison.Ordinal) Then
                    w.Add("--date") : w.Add(ymd)
                End If
            End If
        End If
        If Ui.IsOn(_chkSetPassword) Then
            w.Add("--password") : w.Add(_txtPassword.Text.Trim())
        End If
        If Ui.IsOn(_chkSetAccess) Then
            w.Add("--access-password") : w.Add(_txtAccess.Text.Trim())
        End If
        Return w
    End Function

    ''' <summary>
    ''' Whether the disk's DOS names the levels REMOVE / WRITE / UPDATE (TRSDOS 6 and the Model III TRSDOS line)
    ''' or KILL / UNUSED / WRITE (the Model I TRSDOS and LDOS 5 line). The name trs80-tool printed for this
    ''' file decides when it is one only one vocabulary has.
    ''' </summary>
    Private Function UsesTrsdos6Names() As Boolean
        Dim p As String = _file.Protection.Trim().ToUpperInvariant()
        If p = "KILL" OrElse p = "UNUSED" Then Return False
        If p = "REMOVE" OrElse p = "UPDATE" Then Return True
        Dim d As String = _disk.DosName
        If Regex.IsMatch(d, "^(TRSDOS|LS-DOS) 6\.") Then Return True
        If Regex.IsMatch(d, "^TRSDOS (1\.|2\.7|2\.8)") Then Return True
        If Regex.IsMatch(d, "^(SuperDOS|DOS500|RAPIDOS)\b", RegexOptions.IgnoreCase) Then Return True
        Return False
    End Function

    Private Function ProtectNames() As String()
        Return If(UsesTrsdos6Names(), Trsdos6Names, LdosNames)
    End Function

    Private Shared Function ProtectLabel(level As Integer, names As String()) As String
        Dim n As String = names(level)
        Select Case n
            Case "FULL" : n = UiText.T("dlg.fileprops.prot.full")
            Case "UNUSED" : n = UiText.T("dlg.fileprops.prot.unused")
            Case "NO_ACCESS" : n = UiText.T("dlg.fileprops.prot.noaccess")
        End Select
        Return UiText.T("dlg.fileprops.prot.item", level.ToString(CultureInfo.InvariantCulture), n)
    End Function

    Private Function OriginalProtectIndex() As Integer
        Dim p As String = _file.Protection.Trim().ToUpperInvariant()
        Dim names As String() = ProtectNames()
        For i As Integer = 0 To names.Length - 1
            If names(i) = p Then Return i
        Next
        Dim other As String() = If(names Is LdosNames, Trsdos6Names, LdosNames)
        For i As Integer = 0 To other.Length - 1
            If other(i) = p Then Return i
        Next
        Dim n As Integer
        If Integer.TryParse(p, NumberStyles.Integer, CultureInfo.InvariantCulture, n) AndAlso n >= 0 AndAlso n <= 7 Then Return n
        Return 0
    End Function

    Private Sub BuildUi()
        Dim editable As Boolean = _disk.IsWritable AndAlso _disk.IsTrsdosFamily
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        ' trs80-tool 5.2.0 writes CP/M floppies: a file there can be renamed and its R/O and SYS attributes set.
        _cpmEditable = _disk.IsWritable AndAlso _disk.IsCpm AndAlso Not _disk.IsHardDisk AndAlso caps IsNot Nothing AndAlso caps.HasDiskOption("attrib", "--read-only")
        _nameEditable = editable OrElse _cpmEditable
        _attribEditable = editable

        Dim rows As New List(Of Object)
        _txtName.Text = _file.SlashName
        _txtName.Enabled = _nameEditable
        rows.Add(UiText.T("dlg.fileprops.filename")) : rows.Add(Ui.Across(_txtName))
        _txtDate.Text = _file.DateText
        _txtDate.Enabled = editable
        _chkNoDate.Enabled = editable
        rows.Add(UiText.T("dlg.fileprops.date")) : rows.Add(Ui.Across(_txtDate, _chkNoDate))
        rows.Add("") : rows.Add(Ui.Lbl(UiText.T("dlg.fileprops.datehint"), True))
        rows.Add(UiText.T("dlg.fileprops.size")) : rows.Add(Ui.Lbl(UiText.T("dlg.fileprops.sizevalue", TextUtil.FormatBytes(_file.Size))))
        rows.Add(UiText.T("dlg.fileprops.kind")) : rows.Add(Ui.Lbl(If(_file.ContentType.Length > 0, UiText.T("dlg.fileprops.kindvalue", _file.Description, _file.ContentType), _file.Description)))
        If _file.Damage.Length > 0 Then rows.Add(UiText.T("dlg.fileprops.damage")) : rows.Add(Ui.Lbl(_file.Damage))
        If _file.Sha256.Length > 0 Then rows.Add(UiText.T("dlg.fileprops.sha")) : rows.Add(New Label With {.Text = _file.Sha256, .Font = Ui.Mono(8.5F)})

        If _disk.IsTrsdosFamily Then
            _hasProtect = True
            Dim names As String() = ProtectNames()
            For level As Integer = 0 To names.Length - 1
                _cmbProtect.Items.Add(ProtectLabel(level, names))
            Next
            _cmbProtect.SelectedIndex = OriginalProtectIndex()
            _cmbProtect.Enabled = editable
            rows.Add(UiText.T("dlg.fileprops.protection")) : rows.Add(Ui.Across(_cmbProtect))

            _chkInvisible.Checked = _file.Hidden
            _chkSystem.Checked = _file.IsSystem
            _chkInvisible.Enabled = editable : _chkSystem.Enabled = editable
            Dim attrRow As TableLayout = If(_file.HasLdosBits,
                Ui.Across(_chkInvisible, _chkSystem, Ui.Lbl("    " & UiText.T("dlg.fileprops.ldosbits", If(_file.IsCreated, UiText.T("word.yes"), UiText.T("word.no")), If(_file.IsModified, UiText.T("word.yes"), UiText.T("word.no"))))),
                Ui.Across(_chkInvisible, _chkSystem))
            rows.Add(UiText.T("dlg.fileprops.attributes")) : rows.Add(attrRow)

            ' What the passwords are now: trs80-tool names the ones it recognises, else gives the hash.
            Dim updateNow As String = DiskFile.PasswordText(_file.UpdatePassword, _file.UpdatePasswordHash)
            Dim accessNow As String = If(_file.TimeStampsReplaceAccessPassword, "", DiskFile.PasswordText(_file.AccessPassword, _file.AccessPasswordHash))
            If updateNow.Length > 0 OrElse accessNow.Length > 0 Then
                Dim pwText As String = UiText.T("dlg.fileprops.pwnow.update", If(updateNow.Length > 0, updateNow, "?"))
                If Not _file.TimeStampsReplaceAccessPassword Then pwText = UiText.T("dlg.fileprops.pwnow.access", pwText, If(accessNow.Length > 0, accessNow, "?"))
                rows.Add(UiText.T("dlg.fileprops.pwnow"))
                rows.Add(Ui.Lbl(pwText))
            End If

            _chkSetPassword.Enabled = editable
            _txtPassword.Enabled = False
            AddHandler _chkSetPassword.CheckedChanged, Sub() _txtPassword.Enabled = Ui.IsOn(_chkSetPassword)
            rows.Add(_chkSetPassword) : rows.Add(Ui.Across(_txtPassword, Ui.Lbl(UiText.T("dlg.fileprops.blankremove"), True)))

            ' On DOSes that keep a time stamp where the access password would be, there is none to set.
            Dim accessIsTimeStamp As Boolean = _file.TimeStampsReplaceAccessPassword
            _chkSetAccess.Enabled = editable AndAlso Not accessIsTimeStamp
            _txtAccess.Enabled = False
            AddHandler _chkSetAccess.CheckedChanged, Sub() _txtAccess.Enabled = Ui.IsOn(_chkSetAccess)
            rows.Add(_chkSetAccess) : rows.Add(Ui.Across(_txtAccess, Ui.Lbl(If(accessIsTimeStamp, UiText.T("dlg.fileprops.accessistimestamp"), UiText.T("dlg.fileprops.blankremove")), True)))
        Else
            Dim attrs As String = _file.AttributeLetters
            If _file.CpmUser.HasValue Then attrs = UiText.T("dlg.fileprops.cpmuserin", attrs, _file.CpmUser.Value.ToString(CultureInfo.InvariantCulture))
            If _file.CpmAttributes.Length > 0 Then attrs &= "  " & _file.CpmAttributes
            If _cpmEditable Then
                _chkCpmReadOnly.Checked = _file.CpmReadOnly
                _chkCpmSys.Checked = _file.CpmSys
                rows.Add(UiText.T("dlg.fileprops.attributes")) : rows.Add(Ui.Across(_chkCpmReadOnly, _chkCpmSys))
                If _file.CpmUser.HasValue AndAlso _file.CpmUser.Value <> 0 Then rows.Add(UiText.T("dlg.fileprops.cpmuserarea")) : rows.Add(Ui.Lbl(_file.CpmUser.Value.ToString(CultureInfo.InvariantCulture)))
                rows.Add("") : rows.Add(Ui.Lbl(UiText.T("dlg.fileprops.cpmnone"), True))
            Else
                rows.Add(UiText.T("dlg.fileprops.attributes")) : rows.Add(Ui.Lbl(If(attrs.Trim().Length > 0, attrs, UiText.T("dlg.fileprops.attrsnone"))))
                rows.Add("") : rows.Add(Ui.Lbl(If(_disk.IsCpm, UiText.T("dlg.fileprops.cpmneeds520"), UiText.T("dlg.fileprops.nochange")), True))
            End If
        End If

        Finish(Ui.Fields(rows.ToArray()))
        OkButton.Enabled = editable OrElse _cpmEditable
    End Sub
End Class
