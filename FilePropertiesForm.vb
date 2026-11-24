' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  FilePropertiesForm.vb — "File Properties": filename (rename), date, size,
'  kind, protection level, Invisible/System attributes, date change.
'  Applies through "disk rename" and "disk attrib".
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class FilePropertiesForm
    Inherits Form

    Private ReadOnly _disk As DiskImage
    Private ReadOnly _file As DiskFile
    Private ReadOnly _txtName As New TextBox()
    Private ReadOnly _txtDate As New TextBox()
    Private ReadOnly _chkNoDate As New CheckBox()
    Private ReadOnly _chkInvisible As New CheckBox()
    Private ReadOnly _chkSystem As New CheckBox()
    Private ReadOnly _cmbProtect As New ComboBox()
    Private ReadOnly _txtPassword As New TextBox()
    Private ReadOnly _chkSetPassword As New CheckBox()
    Private ReadOnly _txtAccess As New TextBox()
    Private ReadOnly _chkSetAccess As New CheckBox()
    ''' <summary>CP/M's two attributes, which trs80-tool 5.2.0 changes ("attrib --read-only", "--system").</summary>
    Private ReadOnly _chkCpmReadOnly As New CheckBox()
    Private ReadOnly _chkCpmSys As New CheckBox()
    Private _cpmEditable As Boolean

    ''' <summary>
    ''' Protection levels 0–7 exactly as trs80-tool names them in dir --json ("protection") and accepts them for
    ''' "disk attrib --protect" (trs80-base Trsdos.ts, trsdosProtectionLevelToString / parseTrsdosProtectionLevel).
    ''' The DOSes disagree on levels 1, 3 and 4: TRSDOS 2.x, LDOS 5, NEWDOS/80 and the other 2.3-layout DOSes say
    ''' KILL / (unused) / WRITE; TRSDOS 6, LS-DOS 6, TRSDOS 1.3 and 2.7/2.8 say REMOVE / WRITE / UPDATE.
    ''' The GUI always sends the level NUMBER (0-7), which every DOS accepts, so only the display depends on this.
    ''' </summary>
    Private Shared ReadOnly LdosNames As String() = {"FULL", "KILL", "RENAME", "UNUSED", "WRITE", "READ", "EXEC", "NO_ACCESS"}
    Private Shared ReadOnly Trsdos6Names As String() = {"FULL", "REMOVE", "RENAME", "WRITE", "UPDATE", "READ", "EXEC", "NO_ACCESS"}

    Public Sub New(disk As DiskImage, file As DiskFile)
        Icon = Program.WindowIcon
        _disk = disk
        _file = file
        BuildUi()
    End Sub

    ''' <summary>The normalized new NAME.EXT if the user changed it, else Nothing.</summary>
    Public ReadOnly Property NewName As String
        Get
            If Not _txtName.Enabled Then Return Nothing
            Dim n As String = TextUtil.NormalizeTrsName(_txtName.Text)
            If n Is Nothing Then Return Nothing
            If String.Equals(n, _file.TrsName, StringComparison.OrdinalIgnoreCase) Then Return Nothing
            Return n
        End Get
    End Property

    ''' <summary>Option words for "disk attrib" for whatever changed; empty if nothing.</summary>
    Public Function AttribOptionWords() As List(Of String)
        Dim w As New List(Of String)
        If _cpmEditable Then Return _file.CpmAttribWords(_chkCpmReadOnly.Checked, _chkCpmSys.Checked)
        If Not _chkInvisible.Enabled Then Return w
        If _chkInvisible.Checked <> _file.Hidden Then w.Add(If(_chkInvisible.Checked, "--invisible", "--visible"))
        If _chkSystem.Checked <> _file.IsSystem Then w.Add(If(_chkSystem.Checked, "--system", "--no-system"))
        If _cmbProtect.Enabled AndAlso _cmbProtect.SelectedIndex >= 0 AndAlso _cmbProtect.SelectedIndex <> OriginalProtectIndex() Then
            w.Add("--protect") : w.Add(_cmbProtect.SelectedIndex.ToString(CultureInfo.InvariantCulture))
        End If
        If _chkNoDate.Checked Then
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
        If _chkSetPassword.Checked Then
            w.Add("--password") : w.Add(_txtPassword.Text.Trim())
        End If
        If _chkSetAccess.Checked Then
            w.Add("--access-password") : w.Add(_txtAccess.Text.Trim())
        End If
        Return w
    End Function

    ''' <summary>
    ''' True when this disk's DOS numbers the levels like TRSDOS 6 (REMOVE / WRITE / UPDATE for 1 / 3 / 4).
    ''' Mirrors trs80-base's trsdosProtectionNamesFor(): "TRSDOS 6.x" and "LS-DOS 6.x" by name; the TRSDOS 1.3 layout
    ''' (TRSDOS 1.x and the other DOSes trs80-tool recognizes on it) and TRSDOS 2.7/2.8, which descend from Model II
    ''' TRSDOS; everything else uses the LDOS names. The file's own protection string settles it when it is unambiguous.
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

    ''' <summary>Combo text for one level: "4 - WRITE", "0 - FULL (no restriction)", "3 - (unused on this DOS)".</summary>
    Private Shared Function ProtectLabel(level As Integer, names As String()) As String
        Dim n As String = names(level)
        Select Case n
            Case "FULL" : n = UiText.T("dlg.fileprops.prot.full")
            Case "UNUSED" : n = UiText.T("dlg.fileprops.prot.unused")
            Case "NO_ACCESS" : n = UiText.T("dlg.fileprops.prot.noaccess")
        End Select
        Return UiText.T("dlg.fileprops.prot.item", level.ToString(CultureInfo.InvariantCulture), n)
    End Function

    ''' <summary>The file's current level (0-7) from trs80-tool's protection string; 0 when it cannot be read.</summary>
    Private Function OriginalProtectIndex() As Integer
        Dim p As String = _file.Protection.Trim().ToUpperInvariant()
        Dim names As String() = ProtectNames()
        For i As Integer = 0 To names.Length - 1
            If names(i) = p Then Return i
        Next
        ' A name from the other vocabulary (not expected) or a plain number.
        Dim other As String() = If(names Is LdosNames, Trsdos6Names, LdosNames)
        For i As Integer = 0 To other.Length - 1
            If other(i) = p Then Return i
        Next
        Dim n As Integer
        If Integer.TryParse(p, NumberStyles.Integer, CultureInfo.InvariantCulture, n) AndAlso n >= 0 AndAlso n <= 7 Then Return n
        Return 0
    End Function

    ''' <summary>Every label-and-value row, so the value column can clear the widest label (UiFit.vb).</summary>
    Private ReadOnly _rows As New List(Of KeyValuePair(Of Control, Control))()

    Private Sub BuildUi()
        Text = UiText.T("dlg.fileprops.title")
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False : MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ShowInTaskbar = False

        Dim editable As Boolean = _disk.IsWritable AndAlso _disk.IsTrsdosFamily
        ' A CP/M file can be renamed and given R/O and SYS, by a trs80-tool that has "attrib --read-only" (5.2.0).
        Dim caps As EngineCapabilities = Program.Tool.Capabilities
        _cpmEditable = _disk.IsWritable AndAlso _disk.IsCpm AndAlso Not _disk.IsHardDisk AndAlso caps IsNot Nothing AndAlso caps.HasDiskOption("attrib", "--read-only")
        Dim y As Integer = 12
        _rows.Add(New KeyValuePair(Of Control, Control)(AddLabel(12, y + 3, UiText.T("dlg.fileprops.filename")), _txtName))
        _txtName.Bounds = New Rectangle(110, y, 160, 23)
        _txtName.Text = _file.SlashName
        _txtName.Enabled = editable OrElse _cpmEditable
        Controls.Add(_txtName)
        y += 30
        _rows.Add(New KeyValuePair(Of Control, Control)(AddLabel(12, y + 3, UiText.T("dlg.fileprops.date")), _txtDate))
        _txtDate.Bounds = New Rectangle(110, y, 160, 23)
        _txtDate.Text = _file.DateText
        _txtDate.Enabled = editable
        Controls.Add(_txtDate)
        _chkNoDate.Text = UiText.T("dlg.fileprops.nodate") : _chkNoDate.AutoSize = True
        _chkNoDate.Location = New Point(280, y + 2)
        _chkNoDate.Enabled = editable
        Controls.Add(_chkNoDate)
        AddLabel(110, y + 26, UiText.T("dlg.fileprops.datehint"))
        y += 48
        AddStatic(y, UiText.T("dlg.fileprops.size"), UiText.T("dlg.fileprops.sizevalue", TextUtil.FormatBytes(_file.Size))) : y += 24
        AddStatic(y, UiText.T("dlg.fileprops.kind"), If(_file.ContentType.Length > 0, UiText.T("dlg.fileprops.kindvalue", _file.Description, _file.ContentType), _file.Description)) : y += 24
        If _file.Damage.Length > 0 Then AddStatic(y, UiText.T("dlg.fileprops.damage"), _file.Damage) : y += 24
        If _file.Sha256.Length > 0 Then AddStatic(y, UiText.T("dlg.fileprops.sha"), _file.Sha256) : y += 24

        If _disk.IsTrsdosFamily Then
            Dim lblProt As Label = AddLabel(12, y + 3, UiText.T("dlg.fileprops.protection"))
            _cmbProtect.DropDownStyle = ComboBoxStyle.DropDownList
            Dim names As String() = ProtectNames()
            For level As Integer = 0 To names.Length - 1
                _cmbProtect.Items.Add(ProtectLabel(level, names))
            Next
            _cmbProtect.SelectedIndex = OriginalProtectIndex()
            _cmbProtect.Bounds = New Rectangle(110, y, 220, 23)
            _rows.Add(New KeyValuePair(Of Control, Control)(lblProt, _cmbProtect))
            _cmbProtect.Enabled = editable
            Controls.Add(_cmbProtect)
            y += 30
            _rows.Add(New KeyValuePair(Of Control, Control)(AddLabel(12, y + 3, UiText.T("dlg.fileprops.attributes")), _chkInvisible))
            _chkInvisible.Text = UiText.T("dlg.fileprops.invisible") : _chkInvisible.AutoSize = True : _chkInvisible.Checked = _file.Hidden
            _chkSystem.Text = UiText.T("dlg.fileprops.system") : _chkSystem.AutoSize = True : _chkSystem.Checked = _file.IsSystem
            _chkInvisible.Location = New Point(110, y + 2)
            _chkSystem.Location = New Point(200, y + 2)
            _chkInvisible.Enabled = editable : _chkSystem.Enabled = editable
            Controls.Add(_chkInvisible) : Controls.Add(_chkSystem)
            ' The LDOS line keeps two more bits in the entry, which trs80-tool reports but does not change.
            If _file.HasLdosBits Then AddLabel(290, y + 3, UiText.T("dlg.fileprops.ldosbits", If(_file.IsCreated, UiText.T("word.yes"), UiText.T("word.no")), If(_file.IsModified, UiText.T("word.yes"), UiText.T("word.no"))))
            y += 30
            Dim updateNow As String = DiskFile.PasswordText(_file.UpdatePassword, _file.UpdatePasswordHash)
            Dim accessNow As String = If(_file.TimeStampsReplaceAccessPassword, "", DiskFile.PasswordText(_file.AccessPassword, _file.AccessPasswordHash))
            If updateNow.Length > 0 OrElse accessNow.Length > 0 Then
                Dim pwText As String = UiText.T("dlg.fileprops.pwnow.update", If(updateNow.Length > 0, updateNow, "?"))
                If Not _file.TimeStampsReplaceAccessPassword Then pwText = UiText.T("dlg.fileprops.pwnow.access", pwText, If(accessNow.Length > 0, accessNow, "?"))
                AddStatic(y, UiText.T("dlg.fileprops.pwnow"), pwText) : y += 24
            End If
            _chkSetPassword.Text = UiText.T("dlg.fileprops.setupdate") : _chkSetPassword.AutoSize = True
            _chkSetPassword.Location = New Point(12, y + 2)
            _rows.Add(New KeyValuePair(Of Control, Control)(_chkSetPassword, _txtPassword))
            _chkSetPassword.Enabled = editable
            _txtPassword.Bounds = New Rectangle(170, y, 120, 23)
            _txtPassword.MaxLength = 8
            _txtPassword.Enabled = False
            AddHandler _chkSetPassword.CheckedChanged, Sub() _txtPassword.Enabled = _chkSetPassword.Checked
            Controls.Add(_chkSetPassword) : Controls.Add(_txtPassword)
            AddLabel(300, y + 3, UiText.T("dlg.fileprops.blankremove"))
            y += 30
            _chkSetAccess.Text = UiText.T("dlg.fileprops.setaccess") : _chkSetAccess.AutoSize = True
            _chkSetAccess.Location = New Point(12, y + 2)
            _rows.Add(New KeyValuePair(Of Control, Control)(_chkSetAccess, _txtAccess))
            ' LDOS 5.3 / LS-DOS 6.3 keep a time stamp where the access password would go; trs80-tool says so for
            ' every file of such a disk and refuses to set one.
            Dim accessIsTimeStamp As Boolean = _file.TimeStampsReplaceAccessPassword
            _chkSetAccess.Enabled = editable AndAlso Not accessIsTimeStamp
            _txtAccess.Bounds = New Rectangle(170, y, 120, 23)
            _txtAccess.MaxLength = 8
            _txtAccess.Enabled = False
            AddHandler _chkSetAccess.CheckedChanged, Sub() _txtAccess.Enabled = _chkSetAccess.Checked
            Controls.Add(_chkSetAccess) : Controls.Add(_txtAccess)
            AddLabel(300, y + 3, If(accessIsTimeStamp, UiText.T("dlg.fileprops.accessistimestamp"), UiText.T("dlg.fileprops.blankremove")))
            y += 30
        Else
            _chkInvisible.Enabled = False
            Dim attrs As String = _file.AttributeLetters
            If _file.CpmUser.HasValue Then attrs = UiText.T("dlg.fileprops.cpmuserin", attrs, _file.CpmUser.Value.ToString(CultureInfo.InvariantCulture))
            If _file.CpmAttributes.Length > 0 Then attrs &= "  " & _file.CpmAttributes
            If _cpmEditable Then
                _rows.Add(New KeyValuePair(Of Control, Control)(AddLabel(12, y + 3, UiText.T("dlg.fileprops.attributes")), _chkCpmReadOnly))
                _chkCpmReadOnly.Text = UiText.T("dlg.fileprops.cpmro") : _chkCpmReadOnly.AutoSize = True : _chkCpmReadOnly.Checked = _file.CpmReadOnly
                _chkCpmSys.Text = UiText.T("dlg.fileprops.cpmsys") : _chkCpmSys.AutoSize = True : _chkCpmSys.Checked = _file.CpmSys
                _chkCpmReadOnly.Location = New Point(110, y + 2)
                _chkCpmSys.Location = New Point(240, y + 2)
                Controls.Add(_chkCpmReadOnly) : Controls.Add(_chkCpmSys)
                y += 28
                If _file.CpmUser.HasValue AndAlso _file.CpmUser.Value <> 0 Then AddStatic(y, UiText.T("dlg.fileprops.cpmuserarea"), _file.CpmUser.Value.ToString(CultureInfo.InvariantCulture)) : y += 24
                AddLabel(12, y, UiText.T("dlg.fileprops.cpmnone")) : y += 24
            Else
                AddStatic(y, UiText.T("dlg.fileprops.attributes"), If(attrs.Trim().Length > 0, attrs, UiText.T("dlg.fileprops.attrsnone"))) : y += 24
                AddLabel(12, y, If(_disk.IsCpm, UiText.T("dlg.fileprops.cpmneeds520"), UiText.T("dlg.fileprops.nochange"))) : y += 24
            End If
        End If

        Dim btnOk As New Button() With {.Text = UiText.T("btn.ok"), .DialogResult = DialogResult.OK, .Bounds = New Rectangle(300, y + 6, 80, 26)}
        Dim btnCancel As New Button() With {.Text = UiText.T("btn.cancel"), .DialogResult = DialogResult.Cancel, .Bounds = New Rectangle(390, y + 6, 80, 26)}
        btnOk.Enabled = editable OrElse _cpmEditable
        Controls.Add(btnOk) : Controls.Add(btnCancel)

        ' Make room for the language in force (UiFit.vb); in English nothing here moves. The value
        ' column clears the widest label, and the boxes beside the attribute and password rows
        ' follow whatever is to their left.
        UiFit.AlignColumn(_rows, 10)
        UiFit.SpreadRow(10, _txtDate, _chkNoDate)
        UiFit.SpreadRow(12, _chkInvisible, _chkSystem)
        UiFit.SpreadRow(12, _chkCpmReadOnly, _chkCpmSys)
        UiFit.GrowForm(Me)
        AcceptButton = btnOk
        CancelButton = btnCancel
        ClientSize = New Size(Math.Max(484, ClientSize.Width), y + 42)
        UiFit.ButtonRow(ClientSize.Width - 14, 10, btnOk, btnCancel)
    End Sub

    ''' <summary>Adds a label and gives it back, so the layout pass can measure it.</summary>
    Private Function AddLabel(x As Integer, y As Integer, text As String) As Label
        Dim lbl As New Label() With {.AutoSize = True, .Location = New Point(x, y), .Text = text}
        Controls.Add(lbl)
        Return lbl
    End Function

    Private Sub AddStatic(y As Integer, label As String, value As String)
        Dim lbl As Label = AddLabel(12, y, label)
        Dim val As New Label() With {.AutoSize = False, .Location = New Point(110, y), .Size = New Size(360, 18), .Text = value, .AutoEllipsis = True}
        Controls.Add(val)
        _rows.Add(New KeyValuePair(Of Control, Control)(lbl, val))
    End Sub
End Class
