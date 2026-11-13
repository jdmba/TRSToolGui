' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  PdriveCalculator.vb — NEWDOS/80 PDRIVE line, computed the way TRSTool did
'  (Matthew Reed's FileSystem\NEWDOS80.cpp). Used only by the DMK analysis
'  (DosIdentifier.vb), which reads the image's bytes itself. For an open disk
'  the PDRIVE line comes from trs80-tool (5.1.0 and later reads the same
'  table), so Disk properties does not use this file.
'
'  NEWDOS/80 FORMAT writes the drive-0 PDRIVE table into cylinder 0, sector 2
'  (sector numbers as recorded on the disk; the same number in single and
'  double density). Byte layout (TRSTool's comments):
'     1: number of relative tracks        3: TC (track count)
'     4: SPT (sectors per track)          5: GPL (granules per lump)
'     8: DDSL (directory starting lump)   9: DDGA (directory granules)
'    12: TSR (step rate)                 13: TI letter, first set bit 0..4 → A..E
'    14: flags: bit0 = I, bit1 = J, bit2 = K   15: TD (drive type, A..H)
'  If sector 2 is blank (all 0xE5), TRSTool read cylinder 1 sector 2 (double
'  density) instead, forced TD to E when it was 0, and set the K flag.
'  The DMK analysis hands this module the sector bytes it read from the image.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization

Public Module PdriveCalculator

    ''' <summary>True if the sector looks unformatted/blank (TRSTool's test: first five bytes 0xE5).</summary>
    Public Function IsBlankSector(sector As Byte()) As Boolean
        If sector Is Nothing OrElse sector.Length < 16 Then Return True
        For i As Integer = 0 To 4
            If sector(i) <> &HE5 Then Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' Build the PDRIVE line from the information sector.
    ''' </summary>
    ''' <param name="sector">Cylinder 0 sector 2 (or cylinder 1 sector 2 when the former was blank).</param>
    ''' <param name="fromCylinder1">True when the cylinder-1 fallback was used (applies TRSTool's TD/K adjustments).</param>
    Public Function FromInfoSector(sector As Byte(), fromCylinder1 As Boolean) As String
        If sector Is Nothing OrElse sector.Length < 16 Then Return ""
        Dim b(15) As Byte
        Array.Copy(sector, b, 16)
        If fromCylinder1 Then
            If b(15) = 0 Then b(15) = 4          ' some double-density disks are marked single-density
            b(14) = CByte(b(14) Or &H4)          ' K: cylinder 0 is in the other density
        End If
        Dim spt As Integer = b(4)
        Dim gpl As Integer = b(5)
        Dim tsr As Integer = b(12)
        Dim tc As Integer = b(3)
        Dim td As Integer = b(15)
        Dim ti As Integer = 0
        For i As Integer = 0 To 4
            If (b(13) And (1 << i)) <> 0 Then
                ti = i
                Exit For
            End If
        Next
        Dim ddsl As Integer = b(8)
        Dim ddga As Integer = b(9)
        Dim flagI As Boolean = (b(14) And &H1) <> 0
        Dim flagJ As Boolean = (b(14) And &H2) <> 0
        Dim flagK As Boolean = (b(14) And &H4) <> 0

        Const letters As String = "ABCDEFGH"
        Dim sb As New Text.StringBuilder()
        sb.Append("TI=").Append(letters(ti And 7))
        If flagI Then sb.Append("I"c)
        If flagJ Then sb.Append("J"c)
        If flagK Then sb.Append("K"c)
        sb.Append(",TD=").Append(letters(td And 7))
        sb.Append(",TC=").Append(tc.ToString(CultureInfo.InvariantCulture))
        sb.Append(",SPT=").Append(spt.ToString(CultureInfo.InvariantCulture))
        sb.Append(",TSR=").Append(tsr.ToString(CultureInfo.InvariantCulture))
        sb.Append(",GPL=").Append(gpl.ToString(CultureInfo.InvariantCulture))
        sb.Append(",DDSL=").Append(ddsl.ToString(CultureInfo.InvariantCulture))
        sb.Append(",DDGA=").Append(ddga.ToString(CultureInfo.InvariantCulture))
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Sanity check before trusting the table: TRSTool required byte 0 (directory lump) to equal byte 8 (DDSL);
    ''' TC, SPT and GPL must be in NEWDOS/80's ranges (TC 1–96, SPT 1–72, GPL 1–8). A NEWDOS 2.1 or random
    ''' sector fails this (verified on Matthew Reed's test disks: 1nd, 1nd25, 3nd, 4nd, 4pnd, newdos90 pass; ndplmin fails).
    ''' </summary>
    Public Function LooksPlausible(sector As Byte()) As Boolean
        If sector Is Nothing OrElse sector.Length < 16 Then Return False
        If sector(0) <> sector(8) Then Return False
        If sector(3) < 1 OrElse sector(3) > 96 Then Return False
        If sector(4) < 1 OrElse sector(4) > 72 Then Return False
        If sector(5) < 1 OrElse sector(5) > 8 Then Return False
        Return True
    End Function

End Module
