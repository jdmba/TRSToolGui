' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  TextUtil.vb — name rules, text detection and conversions. These mirror what
'  TRSTool did on the Windows side (the engine does the disk side).
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Module TextUtil

    ''' <summary>
    ''' "PASSWORD" or "blank" for the two password hashes every DOS writes by default, else "". The hash is the
    ''' 16-bit value as trs80-tool reports it (low byte first on the disk). The 2.3 family hashes PASSWORD to 42E0
    ''' and blank to 4296; the TRSDOS 1.3 layout hashes PASSWORD to 8FD3 and blank to 5CEF. Used only by the DMK
    ''' analysis, which reads the bytes itself; for an open disk trs80-tool names these passwords.
    ''' </summary>
    Public Function KnownPasswordText(hash As Integer, layout13 As Boolean) As String
        If layout13 Then
            If hash = &H8FD3 Then Return "PASSWORD"
            If hash = &H5CEF Then Return "blank"
        Else
            If hash = &H42E0 Then Return "PASSWORD"
            If hash = &H4296 Then Return "blank"
        End If
        Return ""
    End Function

    ''' <summary>TRSDOS-family name rule (same as trs80-tool's parseTrsdosFilename): NAME 1-8, EXT 0-3, letters/digits, name starts with a letter; "/" or "." separates.</summary>
    Public Function IsValidTrsName(name As String) As Boolean
        If String.IsNullOrWhiteSpace(name) Then Return False
        If name.Contains(":"c) Then Return False
        Dim parts As String() = name.Split(New Char() {"/"c, "."c})
        If parts.Length > 2 Then Return False
        Dim base As String = parts(0)
        Dim ext As String = If(parts.Length = 2, parts(1), "")
        If Not Regex.IsMatch(base, "^[A-Za-z][A-Za-z0-9]{0,7}$") Then Return False
        If ext.Length > 0 AndAlso Not Regex.IsMatch(ext, "^[A-Za-z0-9]{1,3}$") Then Return False
        Return True
    End Function

    ''' <summary>Normalize a user-typed TRS-80 name to upper-case NAME.EXT (dot form, as trs80-tool's JSON uses). Returns Nothing if invalid.</summary>
    Public Function NormalizeTrsName(name As String) As String
        If Not IsValidTrsName(name) Then Return Nothing
        Dim parts As String() = name.Trim().ToUpperInvariant().Split(New Char() {"/"c, "."c})
        If parts.Length = 2 AndAlso parts(1).Length > 0 Then Return parts(0) & "." & parts(1)
        Return parts(0)
    End Function

    ''' <summary>
    ''' The TRS-80 name trs80-tool would make from a host file name (upper-case, drop non-alphanumerics,
    ''' cut to 8+3, drop an extension that doesn't start with a letter). Returns "" if no name can be made.
    ''' </summary>
    Public Function TrsNameFromHost(hostFileName As String) As String
        Dim upper As String = Path.GetFileName(hostFileName).ToUpperInvariant()
        Dim dot As Integer = upper.LastIndexOf("."c)
        Dim hasExt As Boolean = dot > 0
        Dim nm As String = If(hasExt, upper.Substring(0, dot), upper)
        Dim ex As String = If(hasExt, upper.Substring(dot + 1), "")
        nm = Regex.Replace(nm, "[^A-Z0-9]", "")
        ex = Regex.Replace(ex, "[^A-Z0-9]", "")
        If nm.Length > 8 Then nm = nm.Substring(0, 8)
        If ex.Length > 3 Then ex = ex.Substring(0, 3)
        If nm.Length = 0 OrElse Not Char.IsLetter(nm(0)) Then Return ""
        If ex.Length > 0 AndAlso Not Char.IsLetter(ex(0)) Then ex = ""
        Return If(ex.Length = 0, nm, nm & "." & ex)
    End Function

    ''' <summary>Host file name for a TRS-80 NAME.EXT (what trs80-tool's extract writes).</summary>
    Public Function HostNameFromTrs(trsName As String, lowercase As Boolean) As String
        Dim n As String = trsName.Replace("/"c, "."c)
        Return If(lowercase, n.ToLowerInvariant(), n)
    End Function

    ''' <summary>
    ''' TRSTool's text test: ext TXT/DOC/NOW, or more than 96 % of the bytes are printable
    ''' (CR, LF, TAB and Ctrl-Z count as printable; NUL bytes are ignored). Empty = not text.
    ''' </summary>
    Public Function LooksLikeText(extension As String, data As Byte()) As Boolean
        If extension IsNot Nothing Then
            Dim e As String = extension.ToUpperInvariant()
            If e = "TXT" OrElse e = "DOC" OrElse e = "NOW" Then Return True
        End If
        If data Is Nothing OrElse data.Length = 0 Then Return False
        Dim alpha As Long = 0, nonalpha As Long = 0
        For Each b In data
            If b < 32 Then
                If b = 0 Then Continue For
                If b = 13 OrElse b = 10 OrElse b = 9 OrElse b = &H1A Then alpha += 1 Else nonalpha += 1
            ElseIf b > 127 Then
                nonalpha += 1
            Else
                alpha += 1
            End If
        Next
        If alpha = 0 Then Return False
        Return (alpha * 100L \ data.Length) > 96
    End Function

    ''' <summary>Convert a TRS-80 text file's CR line endings for Windows, per the preference.</summary>
    Public Function ConvertFromTrs80(data As Byte(), conversion As TextConversionOut) As Byte()
        If data Is Nothing OrElse conversion = TextConversionOut.None Then Return data
        Dim ms As New MemoryStream(data.Length + data.Length \ 8)
        For Each b In data
            If b = 13 Then
                If conversion = TextConversionOut.CrToCrLf Then
                    ms.WriteByte(13) : ms.WriteByte(10)
                Else
                    ms.WriteByte(10)
                End If
            Else
                ms.WriteByte(b)
            End If
        Next
        Return ms.ToArray()
    End Function

    ''' <summary>Expand tabs to 8-column stops (TRSTool did this in the viewer). Resets at CR/LF.</summary>
    Public Function ExpandTabs(text As String) As String
        Dim sb As New StringBuilder(text.Length + 64)
        Dim col As Integer = 0
        For Each ch In text
            If ch = ControlChars.Cr OrElse ch = ControlChars.Lf Then
                sb.Append(ch) : col = 0
            ElseIf ch = ControlChars.Tab Then
                Do
                    sb.Append(" "c) : col += 1
                Loop While (col And 7) <> 0
            Else
                sb.Append(ch) : col += 1
            End If
        Next
        Return sb.ToString()
    End Function

    ''' <summary>Bytes → viewer text: drop a trailing Ctrl-Z, CR→CRLF, keep LF, show other control bytes as "."; Latin-1 so every byte maps to a char.</summary>
    Public Function BytesToDisplayText(data As Byte(), isCpm As Boolean) As String
        If data Is Nothing Then Return ""
        Dim n As Integer = data.Length
        If isCpm Then
            Dim z As Integer = Array.IndexOf(data, CByte(&H1A))
            If z >= 0 Then n = z
        ElseIf n > 3 AndAlso data(n - 1) = &H1A Then
            n -= 1
        End If
        Dim sb As New StringBuilder(n + n \ 16)
        Dim prev As Byte = 0
        For i As Integer = 0 To n - 1
            Dim b As Byte = data(i)
            If isCpm Then b = CByte(b And &H7F)
            If b = 13 Then
                sb.Append(vbCrLf)
            ElseIf b = 10 Then
                If prev <> 13 Then sb.Append(vbCrLf)
            ElseIf b = 9 Then
                sb.Append(ControlChars.Tab)
            ElseIf b < 32 OrElse b = 127 Then
                sb.Append("."c)
            ElseIf b > 127 Then
                sb.Append(ChrW(b))   ' TRS-80 graphics/special characters: shown as Latin-1 so the byte is visible
            Else
                sb.Append(ChrW(b))
            End If
            prev = b
        Next
        Return ExpandTabs(sb.ToString())
    End Function

    ''' <summary>Simple hex dump (fallback when the engine's annotated dump isn't available).</summary>
    Public Function HexDump(data As Byte()) As String
        If data Is Nothing Then Return ""
        Dim sb As New StringBuilder(data.Length * 5)
        For offset As Integer = 0 To data.Length - 1 Step 16
            sb.Append(offset.ToString("X4", CultureInfo.InvariantCulture)).Append("  ")
            Dim ascii As New StringBuilder(16)
            For i As Integer = 0 To 15
                If offset + i < data.Length Then
                    Dim b As Byte = data(offset + i)
                    sb.Append(b.ToString("X2", CultureInfo.InvariantCulture)).Append(" "c)
                    ascii.Append(If(b >= 32 AndAlso b < 127, ChrW(b), "."c))
                Else
                    sb.Append("   ")
                End If
                If i = 7 Then sb.Append(" "c)
            Next
            sb.Append(" |").Append(ascii.ToString()).Append("|").Append(vbCrLf)
        Next
        Return sb.ToString()
    End Function

    ''' <summary>Parse a user-typed date (several common forms) into YYYY-MM-DD for the engine. Nothing if unparseable.</summary>
    Public Function ToYmd(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return Nothing
        Dim t As String = text.Trim()
        Dim d As DateTime
        Dim formats As String() = {"yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/d/yy", "dd-MMM-yyyy", "d-MMM-yy", "yyyyMMdd"}
        If DateTime.TryParseExact(t, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        If DateTime.TryParse(t, CultureInfo.CurrentCulture, DateTimeStyles.None, d) Then Return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        Return Nothing
    End Function

    Public Function FormatBytes(n As Long) As String
        Return n.ToString("N0", CultureInfo.CurrentCulture)
    End Function

    Public Function FormatKb(bytes As Long) As String
        Return (bytes / 1024.0).ToString("N2", CultureInfo.CurrentCulture) & "K"
    End Function

End Module
