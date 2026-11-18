' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  FileKinds.vb — what a TRS-80 file's bytes are: EDTASM source, tokenized
'  BASIC (and which machine's), ASCII BASIC, a /CMD program. The file viewer
'  chooses its views with these tests. They are the tests of the Windows
'  Forms viewer, word for word; a change to one belongs in the other too
'  until both viewers call this module.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Text

Public Module FileKinds

    ' TRSTool's byte tests (Filters\EDTASM.cpp, LevelIIBasic.cpp, CMDFile.cpp)
    Public Function IsEdtasm(b As Byte()) As Boolean
        Try
            Dim pos As Integer = If(b(0) = &HD3, 7, 0)
            For i As Integer = 0 To 4
                Dim ch As Byte = b(pos + i)
                If ch < (&H30 Or &H80) OrElse ch > (&H39 Or &H80) Then Return False
            Next
            Dim c As Byte = b(pos + 5)
            Return c = &H20 OrElse c = &H9
        Catch ex As IndexOutOfRangeException
            Return False
        End Try
    End Function

    ''' <summary>
    ''' FF, then line records: next-line pointer (2 bytes), line number (2), tokens, 00; a 0 pointer ends it.
    ''' The pointers are load addresses (Model 4 BASIC loads well above Model III, so no address limit), but
    ''' consecutive ones differ by exactly the record lengths. Model 4 binary constants can hold 00 bytes, so
    ''' each 00 in the first line is tried as its end and the chain is followed from there (up to 8 lines).
    ''' trs80-tool then picks the token table itself (guessBasicDialect).
    ''' </summary>
    Public Function IsTokenizedBasic(b As Byte()) As Boolean
        If b.Length < 7 OrElse b(0) <> &HFF Then Return False
        Dim ptr0 As Integer = b(1) Or (CInt(b(2)) << 8)
        If ptr0 = 0 Then Return False
        ' TRSTool's rule, and FFFF = the placeholder pointers trs80-tool writes when it tokenizes (BASIC relinks on load).
        If ptr0 <= &H8500 OrElse ptr0 = &HFFFF Then Return True
        Dim lastEnd As Integer = Math.Min(b.Length - 1, 1 + 4 + 300)
        For z As Integer = 5 To lastEnd
            If b(z) <> 0 Then Continue For
            Dim base As Integer = ptr0 - (z + 1)
            If BasicChainHolds(b, z + 1, base, b(3) Or (CInt(b(4)) << 8)) Then Return True
        Next
        Return False
    End Function

    Public Function BasicChainHolds(b As Byte(), startPos As Integer, base As Integer, prevLine As Integer) As Boolean
        Dim p As Integer = startPos
        For n As Integer = 1 To 8
            If p = b.Length Then Return False                    ' no end marker before the bytes run out
            If p + 1 >= b.Length Then Return b(p) = 0           ' a single trailing byte
            Dim ptr As Integer = b(p) Or (CInt(b(p + 1)) << 8)
            If ptr = 0 Then Return True                         ' end of program
            If p + 4 >= b.Length Then Return False
            Dim lineNo As Integer = b(p + 2) Or (CInt(b(p + 3)) << 8)
            If lineNo <= prevLine Then Return False
            Dim nextP As Integer = ptr - base
            If nextP <= p + 4 OrElse nextP > b.Length OrElse b(nextP - 1) <> 0 Then Return False
            prevLine = lineNo
            p = nextP
        Next
        Return True
    End Function

    ''' <summary>
    ''' Model 4 BASIC (MS BASIC-80 5.x) writes FF-prefixed function tokens and binary number constants
    ''' (bytes 0B-1F); Level II / Model III BASIC has neither (tokens stop at FB, numbers stay ASCII).
    ''' Quoted strings and the rest of a REM / ' line are skipped (REM is 93 or FB on Model III, 8F on Model 4).
    ''' </summary>
    Public Function IsModel4Basic(b As Byte()) As Boolean
        Dim p As Integer = 1
        Dim m4 As Integer = 0
        While p + 4 < b.Length
            Dim ptr As Integer = b(p) Or (CInt(b(p + 1)) << 8)
            If ptr = 0 Then Exit While
            p += 4
            Dim inString As Boolean = False
            Dim inRem As Boolean = False
            While p < b.Length AndAlso b(p) <> 0
                Dim c As Byte = b(p)
                If inString Then
                    If c = &H22 Then inString = False
                ElseIf Not inRem Then
                    If c = &H22 Then
                        inString = True
                    ElseIf c = &H93 OrElse c = &HFB OrElse c = &H8F Then
                        inRem = True
                    ElseIf c >= &HFC OrElse (c >= &HB AndAlso c <= &H1F) Then
                        m4 += 1
                    End If
                End If
                p += 1
            End While
            ' A binary constant can contain 00; a later line's chain resynchronises enough for a vote.
            p += 1
        End While
        Return m4 >= 2
    End Function

    ''' <summary>Plain-text BASIC: the first few non-blank lines each start with a line number, ascending.</summary>
    Public Function IsAsciiBasic(b As Byte()) As Boolean
        Dim lines As String() = Encoding.Latin1.GetString(b).Split({ChrW(13), ChrW(10)}, StringSplitOptions.RemoveEmptyEntries)
        Dim checkedLines As Integer = 0
        Dim prev As Integer = -1
        For Each raw In lines
            Dim s As String = raw.TrimStart(" "c)
            If s.Length = 0 OrElse s = ChrW(&H1A) Then Continue For
            Dim n As Integer = 0
            While n < s.Length AndAlso n < 6 AndAlso Char.IsDigit(s(n))
                n += 1
            End While
            If n = 0 OrElse n = 6 OrElse n = s.Length Then Return False
            If Not (s(n) = " "c OrElse s(n) = "'"c OrElse Char.IsLetter(s(n))) Then Return False   ' "10 PRINT", "10PRINT"; not "1. item"
            Dim lineNo As Integer = Integer.Parse(s.Substring(0, n), Globalization.CultureInfo.InvariantCulture)
            If lineNo <= prev OrElse lineNo > 65529 Then Return False
            prev = lineNo
            checkedLines += 1
            If checkedLines = 5 Then Exit For
        Next
        Return checkedLines > 0
    End Function

    Public Function IsCmd(b As Byte()) As Boolean
        If b.Length <= 1 Then Return False
        Dim first As Integer = b(0)
        Return first = 1 OrElse first = 3 OrElse first = 5 OrElse first = 7 OrElse first = &H1F
    End Function

    ''' <summary>TRSTool's EDTASM filter (fallback only): strip D3 header, clear bit 7 of the 5-digit line number, tab after it.</summary>
    Public Function DecodeEdtasmLocally(b As Byte()) As String
        Dim sb As New StringBuilder(b.Length)
        Dim pos As Integer = If(b.Length > 0 AndAlso b(0) = &HD3, 7, 0)
        Dim linePos As Integer = 0
        For i As Integer = pos To b.Length - 1
            Dim ch As Byte = b(i)
            If ch = &HD Then
                sb.Append(vbCrLf) : linePos = -1
            ElseIf i = b.Length - 1 AndAlso ch = &H1A Then
                Continue For
            Else
                If linePos < 5 Then ch = CByte(ch And &H7F)
                If linePos = 5 Then ch = 9
                sb.Append(ChrW(ch))
            End If
            linePos += 1
        Next
        Return TextUtil.ExpandTabs(sb.ToString())
    End Function

End Module
