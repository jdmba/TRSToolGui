' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  AsmSource.vb — what kind of assembler source a file is, and where its real
'  content ends.
'
'  Layouts:
'    EDTASM A  D3 + 6-byte name header, 5-digit line numbers with bit 7 set, then SPACE
'    EDTASM B  no header, same line numbers, then TAB
'    TRS-80 text  plain text with CR endings / 1A EOF / NUL padding / high-bit bytes
'    ASCII text   strictly portable 7-bit text
'
'  Slack: a shortened TRS-80 file keeps the tail of an older, longer version
'  after its 1A. Numbered files are cut at a 1A on a record boundary; plain
'  text only when the 1A is in the last sector or followed by nothing but NULs.
'
'  Dialect: DetectDialect hands the statements to AsmDialects.vb, which knows
'  each dialect's telltales.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

''' <summary>Assembler a source file was written for. AsmDialects.Detect returns Edas, Mras, EdtasmPlus, Macro80, Mzal, Zeus, Zen, Alds or Unidentified so far.</summary>
Public Enum AsmDialect
    Unidentified
    Edtasm          ' Radio Shack Editor/Assembler (tape and disk)
    EdtasmPlus      ' Microsoft EDTASM+
    DiskEdtasm      ' Apparat Disk EDTASM
    Edas            ' Misosys EDAS
    Mras            ' MISOSYS MRAS (1985): EDAS 4.3 plus relocation
    Macro80         ' Microsoft M80
    Nedas           ' NEDAS
    Zeus
    Zmac
    Mzal            ' Computer Applications Unlimited M-ZAL (TED editor, MASM assembler)
    Zen             ' Avalon Software ZEN (1979, cassette)
    Alds            ' Tandy Assembly Language Development System (26-2012)
End Enum

Public NotInheritable Class AsmSourceInfo
    ''' <summary>"EDTASM A", "EDTASM B", "TRS-80 text", "ASCII text" (a "?" suffix = mixed delimiters).</summary>
    Public Property Layout As String = ""
    ''' <summary>The 6-byte name in an EDTASM A header, trimmed.</summary>
    Public Property Header As String = ""
    Public Property Lines As Integer
    ''' <summary>Bytes of real content, the 1A included.</summary>
    Public Property Keep As Integer
    ''' <summary>Bytes after the EOF that belong to an older version of the file.</summary>
    Public Property Slack As Integer
    Public Property Notes As New List(Of String)
    Public Property Dialect As AsmDialect = AsmDialect.Unidentified
    ''' <summary>Release the features need, as shown: "IV or later" (EDAS), "disk version" (EDTASM-PLUS); blank if unknown.</summary>
    Public Property DialectVersion As String = ""
    ''' <summary>The telltales that decided the dialect, one short phrase each (AsmDialects.vb).</summary>
    Public Property DialectEvidence As New List(Of String)

    Public ReadOnly Property IsNumbered As Boolean
        Get
            Return Layout.StartsWith("EDTASM", StringComparison.Ordinal)
        End Get
    End Property

    ''' <summary>One line for the viewer: layout, lines, slack, dialect, notes.</summary>
    Public Function Summary() As String
        Dim parts As New List(Of String)
        parts.Add(Layout & If(Header.Length > 0, " """ & Header & """", ""))
        parts.Add(Lines.ToString("N0", CultureInfo.CurrentCulture) & " lines")
        If Slack > 0 Then parts.Add(Slack.ToString("N0", CultureInfo.CurrentCulture) & " slack bytes after EOF hidden")
        If Dialect = AsmDialect.Unidentified Then
            parts.Add("assembler: not yet identified")
        Else
            parts.Add("assembler: " & AsmDialects.DisplayName(Dialect) &
                      If(DialectVersion.Length > 0, " " & DialectVersion, "") &
                      If(DialectEvidence.Count > 0, " (" & String.Join(", ", DialectEvidence) & ")", ""))
        End If
        parts.AddRange(Notes)
        Return String.Join("  •  ", parts)
    End Function
End Class

Public Module AsmSource
    Private Const EofMark As Byte = &H1A
    Private Const HdrMark As Byte = &HD3
    Private Const HdrLen As Integer = 7
    Private Const NumLen As Integer = 5

    ''' <summary>Extensions that mean assembler source on TRS-80 disks.</summary>
    Public Function IsSourceExtension(extension As String) As Boolean
        If extension Is Nothing Then Return False
        Select Case extension.ToUpperInvariant()
            Case "ASM", "SRC", "EDT", "ASC", "MAC", "Z80", "Z"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function IsNumField(data As Byte(), pos As Integer) As Boolean
        If pos < 0 OrElse pos + NumLen > data.Length Then Return False
        For i As Integer = pos To pos + NumLen - 1
            If data(i) < &HB0 OrElse data(i) > &HB9 Then Return False
        Next
        Return True
    End Function

    Private Function NumValue(data As Byte(), pos As Integer) As Integer
        Dim n As Integer = 0
        For i As Integer = pos To pos + NumLen - 1
            n = n * 10 + ((data(i) And &H7F) - &H30)
        Next
        Return n
    End Function

    Private Function CountByte(data As Byte(), upTo As Integer, value As Byte) As Integer
        Dim n As Integer = 0
        For i As Integer = 0 To upTo - 1
            If data(i) = value Then n += 1
        Next
        Return n
    End Function

    ''' <summary>FixASM's Classify, without the renaming: layout, line count, and where the content ends.</summary>
    Public Function Analyze(data As Byte()) As AsmSourceInfo
        Dim v As New AsmSourceInfo With {.Keep = If(data Is Nothing, 0, data.Length)}
        If data Is Nothing OrElse data.Length = 0 Then
            v.Layout = "TRS-80 text"
            v.Notes.Add("empty file")
            Return v
        End If

        ' ---- M-ZAL TED: binary records, FF FF terminator ----
        Dim tedEnd As Integer, tedLines As Integer
        If ScanTed(data, tedEnd, tedLines) Then
            v.Layout = TedLayout
            v.Lines = tedLines
            v.Keep = tedEnd
            v.Slack = data.Length - tedEnd
            v.Dialect = DetectDialect(data, v)
            Return v
        End If

        ' ---- Editor/Assembler: walk CR records until 1A or a break ----
        Dim hdr As Boolean = data.Length > HdrLen AndAlso data(0) = HdrMark
        Dim pos As Integer = If(hdr, HdrLen, 0)
        Dim last As Integer = -1
        Dim records As Integer = 0
        Dim delims As New HashSet(Of Byte)
        Dim cut As Integer = -1
        Dim hitEof As Boolean = False
        While pos < data.Length
            If data(pos) = EofMark Then cut = pos : hitEof = True : Exit While
            Dim endIdx As Integer = Array.IndexOf(data, CByte(&HD), pos)
            If endIdx < 0 Then endIdx = data.Length
            If Not IsNumField(data, pos) OrElse endIdx - pos < NumLen + 1 Then cut = pos : Exit While
            Dim n As Integer = NumValue(data, pos)
            If n <= last Then cut = pos : Exit While
            records += 1
            delims.Add(data(pos + NumLen))
            last = n
            pos = endIdx + 1
        End While

        If records >= 3 Then
            If hdr AndAlso delims.Count = 1 AndAlso delims.Contains(CByte(&H20)) Then
                v.Layout = "EDTASM A"
            ElseIf Not hdr AndAlso delims.Count = 1 AndAlso delims.Contains(CByte(&H9)) Then
                v.Layout = "EDTASM B"
            Else
                v.Layout = If(hdr, "EDTASM A?", "EDTASM B?")
                v.Notes.Add("mixed delimiters")
            End If
            If hdr Then v.Header = Encoding.ASCII.GetString(data, 1, HdrLen - 1).TrimEnd()
            v.Lines = records
            If hitEof Then
                v.Keep = cut + 1
                v.Slack = data.Length - v.Keep
            ElseIf cut >= 0 Then
                v.Notes.Add("line numbers break at offset " & cut.ToString(CultureInfo.InvariantCulture) & " with no 1A")
            Else
                v.Notes.Add("no 1A EOF marker")
            End If
            v.Dialect = DetectDialect(data, v)
            Return v
        End If

        ' ---- plain text: trust the last 1A only if what follows is short or all NULs ----
        Dim eofIdx As Integer = Array.LastIndexOf(data, EofMark)
        Dim validEof As Boolean = False
        If eofIdx >= 0 Then
            validEof = True
            If data.Length - (eofIdx + 1) >= 256 Then
                For i As Integer = eofIdx + 1 To data.Length - 1
                    If data(i) <> 0 Then validEof = False : Exit For
                Next
            End If
        End If
        If validEof Then
            v.Keep = eofIdx + 1
            v.Slack = data.Length - v.Keep
        End If

        Dim sevenBit As Boolean = True
        Dim strayCtrl As Boolean = False
        For i As Integer = 0 To v.Keep - 1
            Dim b As Byte = data(i)
            If b >= &H80 Then sevenBit = False
            If b < &H20 AndAlso b <> &H9 AndAlso b <> &HA AndAlso b <> &HD AndAlso b <> EofMark Then strayCtrl = True
        Next
        If sevenBit AndAlso Not validEof AndAlso Not strayCtrl Then
            v.Layout = "ASCII text"
        Else
            v.Layout = "TRS-80 text"
            If Not sevenBit Then v.Notes.Add("high-bit bytes")
        End If
        Dim lf As Integer = CountByte(data, v.Keep, &HA)
        v.Lines = If(lf > 0, lf, CountByte(data, v.Keep, &HD))
        v.Dialect = DetectDialect(data, v)
        Return v
    End Function

    ''' <summary>
    ''' Assembler-dialect detector: AsmDialects.Detect scores the statements for each dialect's telltales.
    ''' Fills DialectVersion/DialectEvidence.
    ''' </summary>
    Public Function DetectDialect(data As Byte(), info As AsmSourceInfo) As AsmDialect
        Return AsmDialects.Detect(StatementLines(data, info), info)
    End Function

    ''' <summary>The source statements as text: numbered records without their line numbers, or plain-text lines up to the EOF.</summary>
    Public Function StatementLines(data As Byte(), info As AsmSourceInfo) As List(Of String)
        Dim result As New List(Of String)
        Dim limit As Integer = Math.Min(info.Keep, data.Length)
        Dim sb As New StringBuilder()
        If info.Layout = TedLayout Then
            For Each rec In TedRecords(data, limit)
                result.Add(rec.Value)
            Next
        ElseIf info.IsNumbered Then
            Dim pos As Integer = If(data.Length > HdrLen AndAlso data(0) = HdrMark, HdrLen, 0)
            While pos < limit AndAlso data(pos) <> EofMark
                Dim endIdx As Integer = Array.IndexOf(data, CByte(&HD), pos)
                If endIdx < 0 OrElse endIdx > limit Then endIdx = limit
                If Not IsNumField(data, pos) Then Exit While
                sb.Clear()
                For i As Integer = pos + NumLen + 1 To endIdx - 1
                    sb.Append(ChrW(data(i) And &H7F))
                Next
                result.Add(sb.ToString())
                pos = endIdx + 1
            End While
        Else
            For i As Integer = 0 To limit - 1
                Dim b As Byte = data(i)
                If b = &HD OrElse b = &HA OrElse b = EofMark Then
                    If sb.Length > 0 Then result.Add(sb.ToString())
                    sb.Clear()
                    If b = EofMark Then Exit For
                ElseIf b <> 0 Then
                    sb.Append(ChrW(b And &H7F))
                End If
            Next
            If sb.Length > 0 Then result.Add(sb.ToString())
        End If
        Return result
    End Function

    ' ------------------------------------------------------------------ M-ZAL TED layout
    '  The M-ZAL editor (TED) writes a 2-byte header, then one record per line:
    '  line number (16-bit, low byte first), text length (1 byte), text. FF FF
    '  where the next line number would be ends the file; what follows is slack.
    '  (Verified on DEMO, DEMO02, ROMEQU, SECDQ1/ASM from the M-ZAL disk.)

    Public Const TedLayout As String = "M-ZAL TED"

    ''' <summary>True if the whole file walks as TED records to an FF FF terminator (at least 3 lines, ascending numbers, text bytes).</summary>
    Private Function ScanTed(data As Byte(), ByRef endPos As Integer, ByRef lines As Integer) As Boolean
        endPos = 0 : lines = 0
        If data.Length < 2 + 3 * 3 + 2 Then Return False
        Dim pos As Integer = 2
        Dim last As Integer = -1
        While pos + 1 < data.Length
            Dim ln As Integer = data(pos) Or (CInt(data(pos + 1)) << 8)
            If ln = &HFFFF Then
                If lines < 3 Then Return False
                endPos = pos + 2
                Return True
            End If
            If ln <= last OrElse pos + 2 >= data.Length Then Return False
            Dim len As Integer = data(pos + 2)
            If pos + 3 + len > data.Length Then Return False
            For i As Integer = pos + 3 To pos + 2 + len
                Dim b As Byte = data(i)
                If b < &H9 OrElse (b > &HD AndAlso b < &H20) Then Return False
            Next
            last = ln
            lines += 1
            pos += 3 + len
        End While
        Return False
    End Function

    ''' <summary>(line number, text) for each TED record before the terminator.</summary>
    Private Function TedRecords(data As Byte(), limit As Integer) As List(Of KeyValuePair(Of Integer, String))
        Dim recs As New List(Of KeyValuePair(Of Integer, String))
        Dim pos As Integer = 2
        While pos + 2 < limit
            Dim ln As Integer = data(pos) Or (CInt(data(pos + 1)) << 8)
            If ln = &HFFFF Then Exit While
            Dim len As Integer = data(pos + 2)
            If pos + 3 + len > limit Then Exit While
            Dim sb As New StringBuilder(len)
            For i As Integer = pos + 3 To pos + 2 + len
                sb.Append(ChrW(data(i) And &H7F))
            Next
            recs.Add(New KeyValuePair(Of Integer, String)(ln, sb.ToString()))
            pos += 3 + len
        End While
        Return recs
    End Function

    ''' <summary>A TED file as numbered text lines for the viewer ("00010 text"); Nothing for any other layout.</summary>
    Public Function TedListing(data As Byte(), info As AsmSourceInfo) As String
        If info Is Nothing OrElse info.Layout <> TedLayout Then Return Nothing
        Dim sb As New StringBuilder()
        For Each rec In TedRecords(data, Math.Min(info.Keep, data.Length))
            sb.Append(rec.Key.ToString("00000", CultureInfo.InvariantCulture)).Append(" "c).Append(rec.Value.Replace(ControlChars.Tab, "        ")).AppendLine()
        Next
        Return sb.ToString()
    End Function

    ''' <summary>The file's bytes up to and including its EOF, without the slack.</summary>
    Public Function ContentBytes(data As Byte(), info As AsmSourceInfo) As Byte()
        If info Is Nothing OrElse info.Keep >= data.Length Then Return data
        Dim b(info.Keep - 1) As Byte
        Array.Copy(data, b, info.Keep)
        Return b
    End Function
End Module
