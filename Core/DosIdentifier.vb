' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DosIdentifier.vb — names the DOS, model and version of a DMK from the boot
'  sector and the directory structure alone: read the boot sector and the
'  directory, match the boot against the known system boots and FORMAT stubs
'  (family/model), then take the version from the directory, family by family.
'
'  Principle: the directory names the FORMATTER, the boot
'  sector names the last SYSTEM installed; when they disagree, say both.
'  Never guess: an unrecognised boot falls through to the directory and, if
'  that is silent too, the result says so.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text

''' <summary>A directory entry of interest (BOOT/SYS, DIR/SYS, …).</summary>
Public Class DosDirEntry
    Public Property Name As String = ""          ' "BOOT/SYS"
    Public Property Attribute As Byte
    Public Property Byte1 As Byte
    Public Property Byte2 As Byte
    Public Property Password As Byte() = New Byte(3) {}   ' bytes 10h-13h: update hash, access hash
    Public Property Byte1A As Byte
    Public Property Sector As Integer
    Public Property Slot As Integer
    Public Property SectorCount As Integer
    Public Property ExtentSectors As Integer
    Public ReadOnly Property PwText As String
        Get
            Return String.Join(" ", Password.Select(Function(b) b.ToString("X2", CultureInfo.InvariantCulture)))
        End Get
    End Property
    Public Function PwIs(hex As String) As Boolean
        Return String.Equals(PwText, hex, StringComparison.OrdinalIgnoreCase)
    End Function
    Public Function UpdatePwIs(hex As String) As Boolean
        Return String.Equals(PwText.Substring(0, 5), hex, StringComparison.OrdinalIgnoreCase)
    End Function
End Class

''' <summary>The outcome of the identification plus the evidence read on the way.</summary>
Public Class DosIdentification
    Public Property Result As String = ""            ' the wording shown to the user
    Public Property Family As String = ""            ' "LDOS 5", "TRSDOS 6", "NEWDOS/80", "MultiDOS", "VTOS 3.0", … ("" = none)
    Public Property Model As String = ""             ' "Model I", "Model III", "Model III or 4", …
    Public Property BootKind As String = "none"      ' "system", "stub", "none"
    Public Property BootSector As Byte() = Nothing
    Public Property BootIsFm As Boolean
    Public Property BootSectorDescription As String = ""
    Public Property DirectoryCylinder As Integer = -1
    Public Property DirectoryByteCylinder As Integer = -1
    Public Property DirectorySide As Integer = 0
    Public Property LumpLayout As Boolean           ' directory found by scan, not where the boot byte said
    Public Property DirectoryDescription As String = ""
    Public Property Layout13 As Boolean             ' TRSDOS 1.x layout (48-byte entries, GAT sector 1)
    Public Property Gat As Byte() = Nothing
    Public Property Hit As Byte() = Nothing
    Public Property GatMark As Integer
    Public Property Entries As New List(Of DosDirEntry)
    Public Property AllEntries As New List(Of DosDirEntry)
    Public Property Bare As Boolean
    Public Property DiskName As String = ""
    Public Property DiskDate As String = ""
    Public Property AutoCommand As String = ""
    Public Property MasterPasswordText As String = ""
    Public Property GatSummary As String = ""
    Public Property HitSummary As String = ""
    Public Property EntriesSummary As String = ""
    Public Property PdriveLine As String = ""
    Public Property PdriveMedia As PdriveMediaResult = Nothing   ' PDRIVE worked out from the media (NEWDOS/80 only)
    Public Property FillText As String = ""
    Public Property Sector2Byte0 As Integer = -1
    Public Property DosCylinders As Integer
    Public Property DosCylinderSource As String = ""
    Public Property LimitedBackupFiles As New List(Of String)
    Public Property HiddenInUseEntries As Integer
    Public Property ShortFiles As New List(Of String)
    Public Property Evidence As New List(Of String)

    Public Function GatByte(i As Integer) As Integer
        If Gat Is Nothing OrElse i >= Gat.Length Then Return -1
        Return Gat(i)
    End Function

    Public Function Entry(name As String) As DosDirEntry
        For Each e In Entries
            If String.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase) Then Return e
        Next
        Return Nothing
    End Function
End Class

Public Module DosIdentifier

    ' ------------------------------------------------------------------ prefix tables (system boots and FORMAT stubs; ?? = any byte)
    ' Pattern tokens: hex byte, or "??" = any byte (the directory-cylinder byte, a density flag, a per-disk value).
    Private Class Prefix
        Public Pattern As String
        Public Family As String
        Public Model As String
        Public Kind As String      ' "system" or "stub"
        Public Note As String
        Public Sub New(pattern As String, family As String, model As String, kind As String, note As String)
            Me.Pattern = pattern : Me.Family = family : Me.Model = model : Me.Kind = kind : Me.Note = note
        End Sub
    End Class

    Private ReadOnly FmPrefixes As Prefix() = {
        New Prefix("00 FE ?? F3 31 FC 41 CD C9 01 3E 01 32 E1 37 3A 02 42 57", "DOSPLUS", "Model I", "system", "DOSPLUS 3.1 boot (byte 10h 3A 02 42 57)"),
        New Prefix("00 FE ?? F3 31 FC 41 CD C9 01 3E 01 32 E1 37 3A D8 42", "TRSDOS 2.3b", "Model I", "system", "TRSDOS 2.3b boot"),
        New Prefix("00 FE ?? F3 31 FC 41 21 E2 42 CD 9A 42 3E 01 32 E1 37", "TRSDOS 2.x", "Model I", "system", "TRSDOS 2.0-2.3 / 3.0, NEWDOS 2.1 / 3.0, NEWDOS+ boot — FORMAT writes it too"),
        New Prefix("00 FE ?? F3 31 E0 41 21 E2 42 CD 9A 42", "VTOS 3.0", "Model I", "system", "VTOS 3.0 boot; FORMAT writes it too"),
        New Prefix("00 FE ?? F3 31 E0 41 FD 21 00 00", "VTOS 4 / LDOS 5", "Model I", "system", "VTOS 4.0 / LDOS 5 single-density boot — same code; GAT CB decides"),
        New Prefix("00 FE 14 F3 31 E0 41 FD 21 1C 42", "VTOS 4 / LDOS 5", "Model I (double density)", "system", "LDOS 5.1.x / 5.3.1 Sole double-density boot, directory at 20"),
        New Prefix("00 FE ?? F3 31 04 42", "NEWDOS/80 1.0", "Model I", "system", "NEWDOS/80 1.0 boot"),
        New Prefix("00 FE ?? 21 11 42 7E 23 CD 33 00", "NEWDOS/80 1.0", "Model I", "stub", "NEWDOS/80 1.0 FORMAT stub ""NO SYSTEM"""),
        New Prefix("00 FE ?? F3 21 EC 37", "NEWDOS/80 2.x", "Model I", "system", "NEWDOS/80 2.0 / 2.5 / 2.74 Model I boot — FORMAT writes it too"),
        New Prefix("F3 FE ?? 31 FC 41 CD C9 01 3A 02 42 57", "DOSPLUS", "Model I", "system", "DOSPLUS 3.2 boot"),
        New Prefix("F3 FE ?? 31 FC 41 36 FF 36 D0 CD C9 01", "DOSPLUS", "Model I (double density)", "system", "DOSPLUS 3.2.1 DD boot"),
        New Prefix("F3 FE ?? 31 FC 41 CD C9 01 36 FF AF 32 ED 37", "DOSPLUS", "Model I", "system", "DOSPLUS 3.3 boot"),
        New Prefix("F3 FE ?? 31 FC 41 CD C9 01 11 02 00 01 00 50", "DOSPLUS", "Model I", "system", "DOSPLUS 3.4 / 3.4D boot"),
        New Prefix("00 FE ?? CD C9 01 11 02 00 21 00 52 CD B2 42", "DOSPLUS", "Model I", "system", "DOSPLUS 3.5 boot"),
        New Prefix("00 FE ?? F3 31 00 42 CD C9 01 01 00 4D", "UltraDOS", "Model I", "system", "UltraDOS 4.1 boot"),
        New Prefix("00 FE ?? F3 11 01 3C", "UltraDOS", "Model I", "system", "UltraDOS 4.2 system boot"),
        New Prefix("00 FE ?? F3 31 00 42 21 4A 42", "UltraDOS", "Model I", "stub", "UltraDOS 4.2 FORMAT boot"),
        New Prefix("36 FF 36 D0 31 FC 41 CD C9 01", "DoubleDOS", "Model I", "system", "Percom DoubleDOS 4.20 / 4.23 / 4.24 system boot"),
        New Prefix("21 10 42 7E CD 33 00 23 FE 0D 20 F7 CD 40 00 76", "DoubleDOS", "Model I", "stub", "DoubleDOS FORMAT stub ""NO SYSTEM"""),
        New Prefix("FE 00 31 FC 41 CD C9 01 01 00 43 11 EC 37", "Percom MicroDOS", "Model I", "system", "Percom MicroDOS 1.00-2.20 boot track; no directory"),
        New Prefix("C3 7C 42 00 43 00 01", "MMSFORTH", "Model I", "system", "MMSFORTH V2.0 system disk — not a DOS"),
        New Prefix("FE ?? 3E 81 32 EE 37", "TRSDOS 2.7/2.8", "Model I (double density)", "system", "TRSDOS 2.7DD / 2.8DD system boot, Tandy doubler"),
        New Prefix("FE ?? 3E FF 32 EC 37", "TRSDOS 2.7/2.8", "Model I (double density)", "system", "TRSDOS 2.7DD / 2.8DD system boot, Percom doubler"),
        New Prefix("FE ?? F3 31 71 40 CD C9 01", "TRSDOS 2.7/2.8", "Model I (double density)", "stub", "TRSDOS 2.7DD / 2.8DD FORMAT stub ""Not a SYSTEM Disk"""),
        New Prefix("FE ?? 31 FF 41 3E A0", "TRSDOS 2.7SD", "Model I", "system", "TRSDOS 2.7SD boot"),
        New Prefix("01 ?? 11 4D 32 20 40", "MultiDOS", "Model I", "system", "MultiDOS 1.0-1.2 Model I system boot"),
        New Prefix("01 ?? 11 4D 3A 00 42 3D 20", "MultiDOS", "Model I", "system", "MultiDOS 1.3 Model I system boot"),
        New Prefix("?? FE ?? 32 20 40 0F C5 F5 7B AD 47", "MultiDOS 1.0/1.2D", "Model I", "stub", "MultiDOS 1.0 / 1.2D Model I FORMAT boot"),
        New Prefix("?? FE 11 3A 00 42", "MultiDOS", "Model I", "system", "MultiDOS 1.3P / 1.4 Model I boot"),
        New Prefix("?? FE 11 ED 56 3A ?? 42", "MultiDOS", "Model I", "system", "MultiDOS 1.5 / 1.6 / Z'DOS 1.0 Model I boot"),
        New Prefix("00 FE 11 ED 56 E5 DD E1 11 00 50 43 48 CD", "MultiDOS", "Model I", "system", "MultiDOS 1.7 / 2.x Model I boot (byte 0F 71/72 = 1.7, 74 = 2.00 or 2.10+)"),
        New Prefix("?? FE 11 CD 0B 00 11 0E 00 19 CD 67 44", "MultiDOS", "", "stub", "MultiDOS FORMAT stub A: 1.0 (M3) to 1.4"),
        New Prefix("?? FE 11 CD 0B 00 11 2E 00 19 7E 23 B7", "MultiDOS", "", "stub", "MultiDOS FORMAT stub B: 1.5, 1.6"),
        New Prefix("?? FE 11 3A 03 00 FE 17 28 24 CD 0B 00 11 2F 00", "MultiDOS", "", "stub", "MultiDOS FORMAT stub C: 1.7.x, 2.0, 2.10"),
        New Prefix("?? FE 11 CD 0B 00 11 10 00 19 7E 23 CD 33 00 B7 20 F8", "MultiDOS", "", "stub", "MultiDOS / Esoteric FORMAT stub D: 4.01 on; also on some 5.11 / 6.x system disks")
    }

    Private ReadOnly MfmPrefixes As Prefix() = {
        New Prefix("FE ?? 3E D0 D3 F0 21 02 00 22 EA 43", "TRSDOS 1.x", "Model III", "system", "TRSDOS 1.2-1.5 boot — FORMAT writes it too"),
        New Prefix("FE ?? 3A D0 00 D3 F0 21 02 00", "TRSDOS 1.x", "Model III", "system", "TRSDOS 1.1 boot"),
        New Prefix("FE ?? AF 32 EB 43", "RapidDOS III", "Model III", "system", "RapidDOS 1.3 Model III system boot"),
        New Prefix("FE ?? 21 10 43 11 00 3C", "RapidDOS III", "Model III", "stub", "RapidDOS 1.3 Model III FORMAT stub"),
        New Prefix("00 FE ?? F3 3A 02 43 57 1E 04", "LDOS 5", "Model III", "system", "LDOS 5.1.0-5.3.1 Model III boot — FORMAT writes it too"),
        New Prefix("00 1E 14 06 00 CD 7F 04", "LDOS 5", "Model 80 (Lobo MAX-80)", "system", "LDOS 5.3.0 MAX-80 boot"),
        New Prefix("00 FE ?? F3 3E 86 D3 84", "TRSDOS 6", "Model 4", "system", "TRSDOS 6 / LS-DOS 6 system boot"),
        New Prefix("00 FE ?? 3E 86 D3 84 21 00 F8", "DOSPLUS IV", "Model 4", "system", "DOSPLUS IV system boot — the TRSDOS 6 boot without the F3"),
        New Prefix("00 FE ?? F3 21 0E 43 CD 1B 02", "TRSDOS 6", "Model 4", "stub", "TRSDOS 6.0.0 / 6.01.00 FORMAT stub ""No system"""),
        New Prefix("00 FE ?? F3 00 00 00 00 00 00 00 00 00 00 00 00 11 87 FB 21 29 43", "TRSDOS 6", "Model 4", "stub", "TRSDOS 6.01.01 / 6.01.02 FORMAT stub ""Can not boot, Model-4 DATA DISK!"""),
        New Prefix("00 FE ?? F3 00 00 00 00 00 00 00 00 00 00 00 00 11 8D FB 21 29 43", "TRSDOS 6", "Model 4", "stub", "TRSDOS 6.02 / LS-DOS 6.3 FORMAT stub ""Cannot boot, DATA DISK!"""),
        New Prefix("00 FE ?? 3E 82 D3 84", "RapidDOS 4", "Model 4", "system", "RapidDOS 1.3 / 1.5 / 2.0 Model 4 system boot"),
        New Prefix("00 FE ?? 21 11 43 11 00 3C", "RapidDOS 4", "Model 4", "stub", "RapidDOS Model 4 FORMAT stub"),
        New Prefix("00 FE ?? F3 11 05 00 D9 31 E0 41 21 FF 51 CD 45 43", "NEWDOS/80 2.x", "Model III", "system", "NEWDOS/80 2.0 / 2.5 / 2.74 Model III boot — FORMAT writes it too"),
        New Prefix("00 FE ?? F3 11 05 00 D9 31 E0 41 21 FF 51 CD 47 43", "NEWDOS/80 3.0", "Model III", "system", "NEWDOS/80 3.0 Model III boot"),
        New Prefix("3E C3 32 49 40 21 ?? 43 22 4A 40", "DOSPLUS", "Model III", "system", "DOSPLUS Model III system boot; byte 6 C9 = 3.3, CE = 3.4, DD = 3.5"),
        New Prefix("06 0C 21 10 43 7E 23 CD 33 00", "DOSPLUS", "", "stub", "DOSPLUS 3.3.0 FORMAT stub ""NO SYSTEM!"""),
        New Prefix("00 FE ?? 06 0C 21 13 43 7E 23 CD 33 00", "DOSPLUS", "", "stub", "DOSPLUS 3.4 FORMAT stub ""NO SYSTEM!"" — Model I and III identical"),
        New Prefix("FE 00 3E D0 D3 F0 3E C3 32 49 40", "Percom OS-80 III", "Model III", "system", "Percom OS-80 III boot; no directory"),
        New Prefix("31 E5 41 21 C8 43 CD 1B 02", "MultiDOS", "Model III", "system", "MultiDOS 1.0 Model III boot"),
        New Prefix("00 FE 11 AF D3 E4 31 E5 41", "MultiDOS", "Model III", "system", "MultiDOS 1.1 / 1.2 Model III boot"),
        New Prefix("00 FE 11 3A 01 43 FE FE C2 00 00", "MultiDOS", "Model III", "system", "MultiDOS 1.3 / 1.6 / Z'DOS Model III boot"),
        New Prefix("00 FE 11 AF 32 10 42 D3 EC 3E D0 D3 F0", "MultiDOS", "Model III", "system", "MultiDOS 1.3 / 1.6 Model III boot"),
        New Prefix("00 FE 11 3E D0 D3 F0 31 E5 41", "MultiDOS", "Model III", "system", "MultiDOS 1.3 / 1.6 Model III boot"),
        New Prefix("00 FE 11 21 10 42 CB F6 7E D3 EC", "MultiDOS", "Model III", "system", "MultiDOS 1.6c / 1.6d Model III boot"),
        New Prefix("F3 ED 56 3A 02 43 FE 56 C2 00 00", "MultiDOS", "Model III", "system", "MultiDOS 1.7 / 1.7.0 / 1.7.1 Model III boot"),
        New Prefix("?? FE 11 ED 56 3A 04 43 FE 56 C2 00 00", "MultiDOS", "Model III", "system", "MultiDOS 1.71 / 2.00 / 2.10 Model III boot"),
        New Prefix("F3 FE 11 3A 03 43 FE 3A C2 00 00", "MultiDOS", "Model III", "system", "MultiDOS 4.01 / 5.00 Model III boot"),
        New Prefix("00 FE 11 3A 03 43 FE 3A C2 00 00", "MultiDOS", "Model III", "system", "MultiDOS 5.10 Model III boot"),
        New Prefix("F3 3E 81 D3 84 3D 21 00 3C 11 01 3C", "MultiDOS", "Model 4", "system", "MultiDOS 1.7.1 Model 4 boot"),
        New Prefix("3E 81 11 01 3C D3 84 3D", "MultiDOS", "Model 4", "system", "MultiDOS 2.0.0 / 2.1.0 / 5.11 / 6.00 / 6.20 Model 4 boot"),
        New Prefix("3E 86 11 01 F8 D3 84 3E 20", "Esoteric", "Model 4", "system", "Esoteric 2.11-3.20 Model 4 boot"),
        New Prefix("?? FE 11 CD 0B 00 11 0E 00 19 CD 67 44", "MultiDOS", "", "stub", "MultiDOS FORMAT stub A: 1.0 (M3) to 1.4"),
        New Prefix("?? FE 11 CD 0B 00 11 2E 00 19 7E 23 B7", "MultiDOS", "", "stub", "MultiDOS FORMAT stub B: 1.5, 1.6"),
        New Prefix("?? FE 11 3A 03 00 FE 17 28 24 CD 0B 00 11 2F 00", "MultiDOS", "", "stub", "MultiDOS FORMAT stub C: 1.7.x, 2.0, 2.10"),
        New Prefix("?? FE 11 CD 0B 00 11 10 00 19 7E 23 CD 33 00 B7 20 F8", "MultiDOS", "", "stub", "MultiDOS / Esoteric FORMAT stub D: 4.01 on")
    }

    Private Function Matches(sector As Byte(), pattern As String) As Boolean
        Dim toks As String() = pattern.Split(" "c)
        If sector Is Nothing OrElse sector.Length < toks.Length Then Return False
        For i As Integer = 0 To toks.Length - 1
            If toks(i) = "??" Then Continue For
            If sector(i) <> Convert.ToByte(toks(i), 16) Then Return False
        Next
        Return True
    End Function

    Private Function AllZeroAfter(sector As Byte(), fromIndex As Integer) As Boolean
        If sector Is Nothing Then Return False
        For i As Integer = fromIndex To sector.Length - 1
            If sector(i) <> 0 Then Return False
        Next
        Return True
    End Function

    Private Function Hex(b As Byte()) As String
        Return String.Join(" ", b.Select(Function(x) x.ToString("X2", CultureInfo.InvariantCulture)))
    End Function

    Private Function Hex(b As Byte(), start As Integer, count As Integer) As String
        Return String.Join(" ", b.Skip(start).Take(count).Select(Function(x) x.ToString("X2", CultureInfo.InvariantCulture)))
    End Function

    Private Function Ascii(b As Byte(), start As Integer, count As Integer) As String
        Dim sb As New StringBuilder()
        For i As Integer = start To Math.Min(b.Length, start + count) - 1
            Dim c As Integer = b(i) And &H7F
            sb.Append(If(c >= 32 AndAlso c < 127, ChrW(c), "."c))
        Next
        Return sb.ToString()
    End Function

    ' ------------------------------------------------------------------ entry point

    Public Function Identify(img As DmkImage) As DosIdentification
        Dim r As New DosIdentification()
        If img.Kind <> "DMK" OrElse img.Tracks.Count = 0 Then
            r.Result = "Not analysed (not a DMK image)."
            Return r
        End If

        ' the boot sector — FM sector 0, else MFM sector 1, else MFM sector 0 (cylinder 0 side 0)
        Dim t0 As DmkTrack = img.Track(0, 0)
        Dim boot As DmkSector = Nothing
        If t0 IsNot Nothing Then
            boot = t0.FindSector(0, False)
            If boot Is Nothing OrElse Not boot.HasData Then boot = t0.FindSector(1, True)
            If boot Is Nothing OrElse Not boot.HasData Then boot = t0.FindSector(0, True)
        End If
        If boot Is Nothing OrElse Not boot.HasData Then
            r.Result = "No readable boot sector (cylinder 0 has neither FM sector 0 nor MFM sector 1 with data); directory not searched."
            r.Evidence.Add("No boot sector — the disk cannot be booted by a TRS-80 ROM; it may be CP/M, a foreign format, or damaged.")
            Return r
        End If
        r.BootSector = boot.Data
        r.BootIsFm = Not boot.Mfm
        Dim b As Byte() = boot.Data
        r.BootSectorDescription = If(boot.Mfm, "MFM sector " & boot.IdSector.ToString(CultureInfo.InvariantCulture), "FM sector 0") & ", mark " & boot.DataMark.ToString("X2", CultureInfo.InvariantCulture) & ": " & Hex(b, 0, 16) & "  " & Ascii(b, 0, 16)
        Dim s2 As Byte() = img.SectorData(0, 0, 2)
        If s2 IsNot Nothing Then r.Sector2Byte0 = s2(0)

        ' prefix match (stubs and system boots)
        Dim table As Prefix() = If(boot.Mfm, MfmPrefixes, FmPrefixes)
        Dim hit As Prefix = Nothing
        For Each p In table
            If Matches(b, p.Pattern) Then hit = p : Exit For
        Next
        ' the all-zero DOSPLUS 3.5 / DOSPLUS IV boot: 00 FE 14 then zeros
        If hit Is Nothing AndAlso b.Length >= 256 AndAlso b(0) = 0 AndAlso b(1) = &HFE AndAlso AllZeroAfter(b, 3) Then
            hit = New Prefix("", "DOSPLUS 3.5/IV", "", "stub", "all-zero boot after 00 FE nn: DOSPLUS 3.5 or DOSPLUS IV FORMAT; GAT CB decides")
        End If
        ' MMSFORTH data disk: 208 × E5 then " MMSFORTH (c)"
        If hit Is Nothing AndAlso Ascii(b, &HD0, 12).Contains("MMSFORTH") Then
            hit = New Prefix("", "MMSFORTH", "Model I", "stub", "MMSFORTH data-disk fill: 208 × E5 + "" MMSFORTH (c) 198x Miller Microcomputer Services"" — no loader, no directory")
        End If
        If hit IsNot Nothing Then
            r.Family = hit.Family
            r.Model = hit.Model
            r.BootKind = hit.Kind
            r.Evidence.Add("Boot sector: " & hit.Note & ".")
        Else
            r.Evidence.Add("Boot sector matches no system boot or FORMAT stub: " & Hex(b, 0, 20) & ". Falling through to the directory.")
        End If
        If r.Model.Length = 0 AndAlso hit IsNot Nothing Then
            ' MultiDOS stub C gives the model from the FDC addresses at bytes 23-24
            If hit.Pattern.Contains("3A 03 00 FE 17") AndAlso b.Length > &H1D Then
                If b(&H17) = &HF0 AndAlso b(&H18) = 0 Then r.Model = "Model III or 4" Else If b(&H17) = &HEC AndAlso b(&H18) = &H37 Then r.Model = "Model I"
            End If
        End If

        ' directory
        ReadDirectory(img, r, b)

        ' version by family (first matching line wins), with the directory-only fallback
        DecideVersion(img, r, b)

        ' extras: PDRIVE, fill, DOS cylinders
        If r.Family.StartsWith("NEWDOS/80", StringComparison.Ordinal) Then
            Dim pd As Byte() = s2
            Dim fromCyl1 As Boolean = False
            If pd Is Nothing OrElse PdriveCalculator.IsBlankSector(pd) Then
                pd = img.SectorData(1, 0, 2) : fromCyl1 = True
            End If
            If pd IsNot Nothing AndAlso PdriveCalculator.LooksPlausible(pd) Then
                r.PdriveLine = PdriveCalculator.FromInfoSector(pd, fromCyl1) & "  (from cylinder " & If(fromCyl1, "1", "0") & " sector 2)"
                Dim tc As Integer = pd(3)
                If tc > 0 Then
                    r.DosCylinders = tc + If((pd(14) And 4) <> 0, 1, 0)
                    r.DosCylinderSource = "NEWDOS/80 PDRIVE TC" & If((pd(14) And 4) <> 0, " + 1 for the K flag (track 0 in the other density)", "")
                End If
            End If
            r.PdriveMedia = PdriveFromMedia.Analyze(img)
            Dim pm As PdriveMediaResult = r.PdriveMedia
            If pm.Failure.Length > 0 Then
                r.Evidence.Add("PDRIVE from the media: not worked out (" & pm.Failure & ").")
            Else
                For Each n In pm.Notes
                    r.Evidence.Add("PDRIVE from the media: " & n)
                Next
                For Each w In pm.Warnings
                    r.Evidence.Add("PDRIVE from the media, warning: " & w)
                Next
                If r.DosCylinders = 0 AndAlso pm.TC > 0 Then
                    r.DosCylinders = pm.TC + If(pm.FlagK, 1, 0)
                    r.DosCylinderSource = "NEWDOS/80 TC worked out from the media and the GAT" & If(pm.FlagK, " + 1 for the K flag (track 0 in the other density)", "")
                End If
            End If
        End If
        If r.DosCylinders = 0 Then ComputeDosCylinders(img, r)
        r.FillText = FormatterFill(img, r)
        Return r
    End Function

    ' ------------------------------------------------------------------ the directory (steps 2–5)

    Private Sub ReadDirectory(img As DmkImage, r As DosIdentification, b As Byte())
        ' step 2: the directory-cylinder byte
        Dim dirByte As Integer = -1
        If b(0) = &HFE Then
            dirByte = b(1)
        ElseIf b(1) = &HFE Then
            dirByte = b(2) And &H7F
        End If
        r.DirectoryByteCylinder = dirByte
        Dim found As Integer = -1
        Dim foundSide As Integer = 0
        ' candidate from the byte first, then scan cylinders 1 up (both sides) for a DIR/SYS or INHALT/SYS entry
        Dim candidates As New List(Of Integer)
        If dirByte >= 1 AndAlso dirByte < img.Cylinders Then candidates.Add(dirByte)
        For c As Integer = 1 To img.Cylinders - 1
            If Not candidates.Contains(c) Then candidates.Add(c)
        Next
        For Each c In candidates
            For side As Integer = 0 To img.Sides - 1
                If CylinderHasDirSys(img, c, side) Then
                    found = c : foundSide = side
                    Exit For
                End If
            Next
            If found >= 0 Then Exit For
        Next
        If found < 0 Then
            ' fallback: the boot byte when byte 0 = FE (TRSDOS 1.x, 2.7/2.8) or when that cylinder has F8/FA marks
            If dirByte >= 1 AndAlso dirByte < img.Cylinders Then
                Dim t As DmkTrack = img.Track(dirByte, 0)
                If b(0) = &HFE OrElse (t IsNot Nothing AndAlso t.Good.Any(Function(s) s.DataMark = &HF8 OrElse s.DataMark = &HFA)) Then found = dirByte
            End If
        End If
        If found < 0 Then
            r.DirectoryDescription = "No directory found: no cylinder holds a DIR/SYS (or INHALT/SYS) entry and the boot byte points nowhere usable."
            r.Evidence.Add("No directory. Percom MicroDOS / OS-80, MMSFORTH, CP/M and custom-loader disks have none a TRS-80 DOS reads.")
            Return
        End If
        r.DirectoryCylinder = found
        r.DirectorySide = foundSide
        r.LumpLayout = (found <> dirByte)
        Dim dt As DmkTrack = img.Track(found, foundSide)
        Dim good As List(Of DmkSector) = dt.Good.Where(Function(s) s.HasData).OrderBy(Function(s) s.IdSector).ToList()
        If good.Count = 0 Then
            r.DirectoryDescription = "Directory cylinder " & found.ToString(CultureInfo.InvariantCulture) & " has no readable sector."
            Return
        End If
        ' step 3: GAT = first sector with F8/FA; if all FB and numbering starts at 1 → sector 1 (TRSDOS 1.x layout)
        Dim gatSec As DmkSector = good.FirstOrDefault(Function(s) s.DataMark = &HF8 OrElse s.DataMark = &HFA)
        Dim layout13 As Boolean = False
        If gatSec Is Nothing Then
            If good(0).IdSector = 1 Then
                gatSec = good(0) : layout13 = True
            Else
                gatSec = good(0)
            End If
        End If
        ' TRSDOS 1.x layout is also recognised by its numbering (1-18) and FB directory marks even when the GAT test picked FB
        If Not layout13 AndAlso good(0).IdSector = 1 AndAlso gatSec.DataMark = &HFB Then layout13 = True
        r.Layout13 = layout13
        r.GatMark = gatSec.DataMark
        r.Gat = gatSec.Data
        Dim hitSec As DmkSector = good.FirstOrDefault(Function(s) s.IdSector = gatSec.IdSector + 1)
        If hitSec IsNot Nothing Then r.Hit = hitSec.Data
        r.DirectoryDescription = "cylinder " & found.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, " side " & foundSide.ToString(CultureInfo.InvariantCulture), "") & If(r.LumpLayout, " (found by scan; the boot byte says " & If(dirByte >= 0, dirByte.ToString(CultureInfo.InvariantCulture), "nothing") & ")", " (as the boot byte says)") & "; GAT = sector " & gatSec.IdSector.ToString(CultureInfo.InvariantCulture) & " (mark " & gatSec.DataMark.ToString("X2", CultureInfo.InvariantCulture) & ")" & If(layout13, "; TRSDOS 1.x layout (48-byte entries)", "; 2.3-family layout (32-byte entries)")
        ' GAT fields
        Dim g As Byte() = r.Gat
        If g.Length >= 256 Then
            r.GatSummary = Hex(g, &HCB, 3) & " / " & Hex(g, &HCE, 2) & If(Ascii(g, &HCB, 3) = "U41" OrElse Ascii(g, &HCB, 3) = "U42", "  (""" & Ascii(g, &HCB, 3) & """)", "")
            r.DiskName = Ascii(g, &HD0, 8).TrimEnd()
            r.DiskDate = Ascii(g, &HD8, 8).TrimEnd()
            Dim auto As String = ""
            For i As Integer = &HE0 To &HFF
                If g(i) = 13 OrElse g(i) = 3 Then Exit For
                auto &= ChrW(g(i) And &H7F)
            Next
            If auto.Trim().Length > 0 AndAlso auto.All(Function(c) c >= " "c AndAlso c <= "~"c) Then r.AutoCommand = auto.Trim()
            Dim pw As Integer = g(&HCE) Or (CInt(g(&HCF)) << 8)
            Dim knownPw As String = TextUtil.KnownPasswordText(pw, layout13)
            r.MasterPasswordText = If(knownPw.Length > 0, knownPw & " (hash " & Hex(g, &HCE, 2) & ")", "hash " & Hex(g, &HCE, 2))
            If g(&HF5) = 3 Then r.Evidence.Add("GAT F5-FF media block: 03 """ & Ascii(g, &HF6, 3) & """ " & Hex(g, &HF9, 7) & " (TRSDOS 6.2 / LS-DOS 6.3; highest cylinder " & g(&HFC).ToString(CultureInfo.InvariantCulture) & ", directory cylinder " & g(&HFF).ToString(CultureInfo.InvariantCulture) & ").")
        End If
        If r.Hit IsNot Nothing Then r.HitSummary = Hex(r.Hit, 0, 2)

        ' step 4: entries
        Dim entrySectors As New List(Of DmkSector)
        For Each s In good
            If s.IdSector > gatSec.IdSector + 1 Then entrySectors.Add(s)
        Next
        For Each s In good
            If s.IdSector < gatSec.IdSector Then entrySectors.Add(s)
        Next
        ' second side of the directory cylinder (double-sided allocation) follows
        If img.Sides = 2 Then
            Dim other As DmkTrack = img.Track(found, 1 - foundSide)
            If other IsNot Nothing AndAlso foundSide = 0 Then
                For Each s In other.Good.Where(Function(x) x.HasData).OrderBy(Function(x) x.IdSector)
                    If s.DataMark = &HF8 OrElse s.DataMark = &HFA OrElse layout13 Then entrySectors.Add(s)
                Next
            End If
        End If
        Dim wanted As String() = {"BOOT/SYS", "DIR/SYS", "SYSFF/SYS", "SYSFE/SYS", "DOS/SYS", "GDOS/SYS", "HDOS/SYS", "INHALT/SYS", "NoDOS/Ind"}
        Dim entrySize As Integer = If(layout13, 48, 32)
        For Each s In entrySectors
            Dim d As Byte() = s.Data
            Dim n As Integer = d.Length \ entrySize
            For slot As Integer = 0 To n - 1
                Dim o As Integer = slot * entrySize
                Dim attr As Byte = d(o)
                Dim nameRaw As String = Ascii(d, o + 5, 8).TrimEnd() & "/" & Ascii(d, o + 13, 3).TrimEnd()
                If nameRaw = "/" Then Continue For
                Dim e As New DosDirEntry With {.Name = nameRaw, .Attribute = attr, .Byte1 = d(o + 1), .Byte2 = d(o + 2), .Byte1A = d(o + &H1A), .Sector = s.IdSector, .Slot = slot}
                Array.Copy(d, o + &H10, e.Password, 0, 4)
                e.SectorCount = d(o + &H14) Or (CInt(d(o + &H15)) << 8)
                ' extents (2.3 family): 4 × (lump, granule/count) at 0x16; count = low 5 bits + 1
                If Not layout13 Then
                    Dim ext As Integer = 0
                    For k As Integer = 0 To 3
                        Dim lump As Byte = d(o + &H16 + k * 2), gc As Byte = d(o + &H17 + k * 2)
                        If lump = &HFF OrElse lump = &HFE Then Exit For
                        ext += (gc And &H1F) + 1
                    Next
                    e.ExtentSectors = ext   ' in granules; compared loosely below
                End If
                Dim inUse As Boolean = (attr And &H10) <> 0
                Dim extended As Boolean = (attr And &H80) <> 0
                If Not extended Then
                    r.AllEntries.Add(e)
                    If inUse AndAlso (attr And &H20) <> 0 AndAlso layout13 Then r.LimitedBackupFiles.Add(e.Name)
                    If Not inUse AndAlso extended = False AndAlso nameRaw.Length > 1 AndAlso (attr And &H80) <> 0 Then r.HiddenInUseEntries += 1
                End If
                If wanted.Contains(nameRaw, StringComparer.OrdinalIgnoreCase) Then r.Entries.Add(e)
            Next
        Next
        ' hidden entries: attribute 8xh with a name (in-use bit clear, bit 7 set)
        r.HiddenInUseEntries = r.AllEntries.Where(Function(e) (e.Attribute And &H90) = &H80 AndAlso e.Name.Length > 1 AndAlso Not e.Name.StartsWith("/", StringComparison.Ordinal)).Count()
        Dim parts As New List(Of String)
        For Each e In r.Entries
            parts.Add(e.Name & " attr " & e.Attribute.ToString("X2", CultureInfo.InvariantCulture) & " pw " & e.PwText & If(e.Name = "BOOT/SYS" OrElse e.Name = "DIR/SYS", " b1-2 " & e.Byte1.ToString("X2", CultureInfo.InvariantCulture) & " " & e.Byte2.ToString("X2", CultureInfo.InvariantCulture), ""))
        Next
        r.EntriesSummary = If(parts.Count = 0, "none of BOOT/SYS, DIR/SYS, SYSFF/SYS, DOS/SYS found", String.Join("; ", parts))
        ' step 5: bare test — two or fewer non-zero HIT bytes (0..DFh; TRSDOS 1.x keeps FF at E0-FF)
        If r.Hit IsNot Nothing Then
            Dim nz As Integer = 0
            For i As Integer = 0 To Math.Min(&HDF, r.Hit.Length - 1)
                If r.Hit(i) <> 0 AndAlso i <> &H1F Then nz += 1
            Next
            r.Bare = nz <= 2
            r.Evidence.Add("HIT: " & nz.ToString(CultureInfo.InvariantCulture) & " non-zero slot byte(s) → " & If(r.Bare, "bare format (only BOOT/SYS and DIR/SYS)", "files present") & ".")
        End If
    End Sub

    Private Function CylinderHasDirSys(img As DmkImage, cyl As Integer, side As Integer) As Boolean
        Dim t As DmkTrack = img.Track(cyl, side)
        If t Is Nothing OrElse Not t.Present Then Return False
        For Each s In t.Sectors
            If Not s.IdCrcOk OrElse Not s.HasData Then Continue For
            Dim d As Byte() = s.Data
            For Each esz In New Integer() {32, 48}
                For o As Integer = 0 To d.Length - esz Step esz
                    If o + 16 > d.Length Then Exit For
                    Dim nm As String = Ascii(d, o + 5, 11)
                    If nm = "DIR     SYS" OrElse nm = "INHALT  SYS" Then Return True
                Next
            Next
        Next
        Return False
    End Function

    ' ------------------------------------------------------------------ version rules

    Private Sub DecideVersion(img As DmkImage, r As DosIdentification, b As Byte())
        Dim cb As Integer = r.GatByte(&HCB), cc As Integer = r.GatByte(&HCC), cd As Integer = r.GatByte(&HCD)
        Dim ce As String = If(r.Gat IsNot Nothing AndAlso r.Gat.Length >= 256, Hex(r.Gat, &HCE, 2), "")
        Dim hit0 As Integer = If(r.Hit IsNot Nothing, r.Hit(0), -1), hit1 As Integer = If(r.Hit IsNot Nothing AndAlso r.Hit.Length > 1, r.Hit(1), -1)
        Dim bootE As DosDirEntry = If(r.Entry("BOOT/SYS"), If(r.Entry("GDOS/SYS"), r.Entry("HDOS/SYS")))
        Dim dirE As DosDirEntry = If(r.Entry("DIR/SYS"), r.Entry("INHALT/SYS"))
        Dim suffix As String = If(r.Bare, " (bare data disk)", If(r.BootKind = "system", " (system disk)", ""))
        Dim density As String = ""
        Dim dataTrack As DmkTrack = img.Track(Math.Min(1, img.Cylinders - 1), 0)
        If dataTrack IsNot Nothing Then density = If(dataTrack.HasMfm AndAlso Not dataTrack.HasFm, "double density", If(dataTrack.HasFm AndAlso Not dataTrack.HasMfm, "single density", ""))
        Dim ver As String = ""
        Dim fam As String = r.Family

        ' an unrecognised boot falls through to the directory
        If fam.Length = 0 AndAlso r.Gat IsNot Nothing Then
            Dim dirCyl As Integer = r.DirectoryCylinder
            If (dirE IsNot Nothing AndAlso dirE.UpdatePwIs("DD 23")) OrElse (ce = "96 42" AndAlso cb = &H60 AndAlso dirCyl = 20) OrElse (r.Sector2Byte0 >= 0 AndAlso Ascii(img.SectorData(0, 0, 2), 0, 3) = "DCT") Then
                fam = "DOSPLUS IV" : r.Model = "Model 4"
                r.Evidence.Add("Directory names DOSPLUS IV (DD 23 passwords / GAT CE-CF 96 42 with CB 60 on cylinder 20 / ""DCT"" in sector 2) — tested before TRSDOS 6.")
            ElseIf ce = "96 42" AndAlso cb >= &H31 AndAlso cb <= &H40 Then
                fam = "DOSPLUS" : r.Evidence.Add("Directory names DOSPLUS: GAT CE-CF 96 42 with CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ".")
            ElseIf cb >= &H60 AndAlso cb <= &H63 AndAlso (dirCyl = 20 OrElse dirCyl = 40) Then
                fam = "TRSDOS 6" : r.Model = "Model 4" : r.Evidence.Add("Directory names TRSDOS 6 / LS-DOS 6: GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & " at cylinder " & dirCyl.ToString(CultureInfo.InvariantCulture) & ".")
            ElseIf cb >= &H50 AndAlso cb <= &H53 Then
                fam = "VTOS 4 / LDOS 5" : r.Evidence.Add("Directory names LDOS 5: GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ".")
            ElseIf cb = &H80 OrElse cb = &H82 Then
                fam = "NEWDOS/80 (directory)" : r.Evidence.Add("Directory names NEWDOS/80: GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ".")
            ElseIf Ascii(r.Gat, &HCB, 2) = "U4" Then
                fam = "UltraDOS" : r.Evidence.Add("Directory names UltraDOS: GAT CB-CD """ & Ascii(r.Gat, &HCB, 3) & """.")
            ElseIf (hit0 = &HC4 AndAlso hit1 = &H94) OrElse (hit0 = &H94 AndAlso hit1 = &HC4) OrElse (hit0 = &H32 AndAlso hit1 = &HBA) OrElse Ascii(r.Gat, &HD0, 32).Contains("MULTIDOS") Then
                fam = "MultiDOS" : r.Evidence.Add("Directory names MultiDOS: HIT " & Hex(r.Hit, 0, 2) & " or ""MULTIDOS"" in the GAT.")
            ElseIf (cb = &H20 OrElse cb = &H21) AndAlso hit1 = &H2C Then
                fam = "TRSDOS 2.x" : r.Model = "Model I" : r.Evidence.Add("Directory names TRSDOS 2.x: GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & " with HIT byte 1 = 2C.")
            ElseIf cb = &H21 AndAlso dirCyl = 10 AndAlso bootE Is Nothing AndAlso dirE IsNot Nothing Then
                fam = "DoubleDOS" : r.Model = "Model I" : r.Evidence.Add("Directory names DoubleDOS: GAT CB 21 on a cylinder-10 directory with no BOOT/SYS.")
            ElseIf cb = &H30 Then
                fam = "VTOS 3.0" : r.Model = "Model I" : r.Evidence.Add("Directory names VTOS 3.0: GAT CB 30.")
            End If
            If fam.Length > 0 Then
                r.Family = fam
                ' the boot is unrecognised: report the directory's family and say so
                r.Result = "Unrecognised boot sector; directory looks like " & VersionFor(img, r, fam, b, cb, cc, cd, ce, hit0, hit1, bootE, dirE, True) & suffix & "."
                Return
            End If
            r.Result = If(r.Gat Is Nothing, "Unrecognised boot sector and no readable directory — not a known TRS-80 DOS format (custom loader, CP/M, or a foreign format).", "Unrecognised boot sector; the directory on cylinder " & r.DirectoryCylinder.ToString(CultureInfo.InvariantCulture) & " names no known family (GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ", HIT " & Hex(r.Hit, 0, 2) & ").")
            Return
        End If
        If fam.Length = 0 Then
            r.Result = "Unrecognised boot sector and no directory — not a TRS-80 DOS disk as imaged (custom loader, CP/M, MicroDOS/OS-80, MMSFORTH data, or a foreign format)."
            Return
        End If

        ver = VersionFor(img, r, fam, b, cb, cc, cd, ce, hit0, hit1, bootE, dirE, False)
        Dim modelText As String = If(r.Model.Length > 0, ", " & r.Model, "")
        Dim densText As String = If(density.Length > 0 AndAlso Not r.Model.Contains("density"), ", " & density, "")
        r.Result = ver & modelText & densText & suffix & "."
        ' formatter vs system disagreement: boot names one family, directory another
        If r.Gat IsNot Nothing Then
            Dim dirFam As String = DirectoryFamily(cb, ce, hit0, hit1, r.DirectoryCylinder, bootE, dirE, r.Gat)
            If dirFam.Length > 0 AndAlso Not SameFamily(dirFam, fam) Then
                r.Result = "Formatted by " & dirFam & " (directory); boot sector from " & ver & modelText & suffix & "."
                r.Evidence.Add("The directory (formatter) and the boot sector (last system installed) name different families; both are reported.")
            End If
        End If
    End Sub

    Private Function SameFamily(a As String, b As String) As Boolean
        Dim ka As String = a.Split(" "c)(0), kb As String = b.Split(" "c)(0)
        If ka = "VTOS" AndAlso (kb = "VTOS" OrElse kb = "LDOS") Then Return True
        If ka = "LDOS" AndAlso (kb = "VTOS" OrElse kb = "LDOS") Then Return True
        If ka = "TRSDOS" AndAlso kb = "TRSDOS" Then Return True
        If (ka = "TRSDOS" AndAlso kb = "NEWDOS") OrElse (ka = "NEWDOS" AndAlso kb = "TRSDOS") Then Return True
        If (ka = "DOSPLUS" AndAlso kb = "DOSPLUS") Then Return True
        If (ka = "MultiDOS" OrElse ka = "Esoteric") AndAlso (kb = "MultiDOS" OrElse kb = "Esoteric") Then Return True
        Return String.Equals(ka, kb, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>What the directory alone says the formatter family is, "" if silent.</summary>
    Private Function DirectoryFamily(cb As Integer, ce As String, hit0 As Integer, hit1 As Integer, dirCyl As Integer, bootE As DosDirEntry, dirE As DosDirEntry, gat As Byte()) As String
        If dirE IsNot Nothing AndAlso dirE.UpdatePwIs("DD 23") AndAlso cb = &H60 Then Return "DOSPLUS IV"
        If ce = "96 42" AndAlso cb >= &H31 AndAlso cb <= &H40 Then Return "DOSPLUS"
        If cb >= &H60 AndAlso cb <= &H63 AndAlso (dirCyl = 20 OrElse dirCyl = 40) Then Return "TRSDOS 6"
        If cb >= &H50 AndAlso cb <= &H53 Then Return "LDOS 5"
        If cb = &H40 Then Return ""   ' VTOS 4 or DOSPLUS 3.4a — ambiguous
        If cb = &H30 Then Return "VTOS 3.0"
        If cb = &H80 OrElse cb = &H82 Then Return "NEWDOS/80"
        If Ascii(gat, &HCB, 2) = "U4" Then Return "UltraDOS"
        If (hit0 = &HC4 AndAlso hit1 = &H94) OrElse (hit0 = &H94 AndAlso hit1 = &HC4) OrElse (hit0 = &H32 AndAlso hit1 = &HBA) Then Return "MultiDOS"
        If (cb = &H20 OrElse cb = &H21) AndAlso hit1 = &H2C AndAlso dirCyl <> 10 Then Return "TRSDOS 2.x"
        Return ""
    End Function

    ''' <summary>The version, family by family. Returns "DOS version(s)" text without model/suffix.</summary>
    Private Function VersionFor(img As DmkImage, r As DosIdentification, fam As String, b As Byte(), cb As Integer, cc As Integer, cd As Integer, ce As String, hit0 As Integer, hit1 As Integer, bootE As DosDirEntry, dirE As DosDirEntry, directoryOnly As Boolean) As String
        Dim ev As List(Of String) = r.Evidence
        Dim s2b0 As Integer = r.Sector2Byte0
        Select Case fam
            Case "TRSDOS 1.x"
                ' sign-on in cylinder 0 sectors 13-15 (system disks)
                For Each n In New Integer() {13, 14, 15}
                    Dim d As Byte() = img.SectorData(0, 0, n)
                    If d Is Nothing Then Continue For
                    Dim txt As String = Ascii(d, 0, 256)
                    Dim m As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(txt, "TRSDOS\s+version\s+(1\.\d+)|01\.(0\d)\.00|operating System (1\.\d)v", Text.RegularExpressions.RegexOptions.IgnoreCase)
                    If m.Success Then
                        Dim v As String = If(m.Groups(1).Success, m.Groups(1).Value, If(m.Groups(2).Success, "1." & CInt(m.Groups(2).Value).ToString(CultureInfo.InvariantCulture), m.Groups(3).Value))
                        ev.Add("Sign-on text in cylinder 0 sector " & n.ToString(CultureInfo.InvariantCulture) & " names TRSDOS " & v & " (system disks only).")
                        Return "TRSDOS " & v
                    End If
                Next
                Dim fe As Integer = b(&HFE)
                ev.Add("Boot byte FE (BCD version) = " & fe.ToString("X2", CultureInfo.InvariantCulture) & ", GAT CE-CF " & ce & ", HIT 0-1 " & Hex(r.Hit, 0, 2) & ".")
                If fe = &H11 Then Return "TRSDOS 1.1"
                If fe = &H12 Then Return "TRSDOS 1.2"
                If fe = &H14 Then Return "TRSDOS 1.4"
                If ce = "F2 92" OrElse (hit0 = &HA2 AndAlso hit1 = &HC4) Then Return "TRSDOS 1.5 (GRL)"
                If r.Hit IsNot Nothing AndAlso r.Hit.Length > &HC0 AndAlso r.Hit(&HC0) = &HFF Then Return "TRSDOS 1.4 (FORMAT)"
                If ce = "EF 5C" Then Return "TRSDOS 1.3.2, 1.3I or 1.4"
                Return If(r.Bare, "TRSDOS 1.3", "TRSDOS 1.3 (or 1.3.2 / 1.3I / 1.4 / 1.5 — only the system boot is present)")
            Case "RapidDOS III"
                Return "RapidDOS 1.3"
            Case "TRSDOS 2.x"
                ev.Add("GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ", C0-C1 " & Hex(r.Gat, &HC0, 2) & ", HIT byte 1 " & If(hit1 >= 0, hit1.ToString("X2", CultureInfo.InvariantCulture), "?") & ", boot byte 7D " & b(&H7D).ToString("X2", CultureInfo.InvariantCulture) & ", BB-BE " & Hex(b, &HBB, 4) & ".")
                If cb = &H20 Then Return "TRSDOS 2.0"
                If cb = &HFC OrElse (bootE IsNot Nothing AndAlso bootE.PwIs("FF FF FF FF")) Then Return "TRSDOS 3.0"
                If hit1 = &HC4 Then Return "NEWDOS+ 2.2 or 3.0 (FORMAT)"
                If b(&HBB) = &HE3 AndAlso b(&HBC) = &HE3 AndAlso b(&HBD) = &HE3 AndAlso b(&HBE) = &HE3 Then Return "NEWDOS+ 2.1"
                If r.GatByte(&HC0) = &H23 OrElse r.GatByte(&HC1) = &HA5 Then Return "NEWDOS 3.0 or NEWDOS+ 3.0"
                If b(&H7D) = &HB6 Then Return "TRSDOS 2.2 or 2.3 (FORMAT)"
                If b(&H7D) = &HE1 AndAlso r.Bare Then Return "TRSDOS 2.1, NEWDOS 2.1 or NEWDOS+++ (FORMAT)"
                Return "TRSDOS 2.1-2.3, NEWDOS 2.1 or NEWDOS+++"
            Case "TRSDOS 2.3b"
                Return "TRSDOS 2.3b"
            Case "TRSDOS 2.7/2.8"
                If r.BootKind = "system" Then
                    Dim f6 As String = Hex(b, &HF6, 3)
                    ev.Add("Boot bytes F6-F8 = " & f6 & " (00 52 08 → 2.7, 01 21 58 → 2.8).")
                    If f6 = "00 52 08" Then Return "TRSDOS 2.7DD"
                    If f6 = "01 21 58" Then Return "TRSDOS 2.8DD"
                End If
                Return "TRSDOS 2.7DD or 2.8DD"
            Case "TRSDOS 2.7SD"
                Return "TRSDOS 2.7SD"
            Case "VTOS 3.0"
                Return "VTOS 3.0, 3.0.0 or 3.0.1"
            Case "VTOS 4 / LDOS 5"
                ev.Add("GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ", cylinder 0 sector 2 byte 0 " & If(s2b0 >= 0, s2b0.ToString("X2", CultureInfo.InvariantCulture), "?") & ", BOOT/SYS pw " & If(bootE IsNot Nothing, bootE.PwText, "—") & ", DIR/SYS pw " & If(dirE IsNot Nothing, dirE.PwText, "—") & ".")
                If cb = &H40 Then r.Family = "VTOS 4.0" : Return "VTOS 4.0, 4.0.1 or 4.0.2"
                r.Family = "LDOS 5"
                If cb = &H50 Then Return "LDOS 5.0.x"
                If cb = &H51 OrElse cb = &H34 Then
                    If bootE IsNot Nothing AndAlso bootE.PwIs("96 42 96 42") Then Return "LDOS 5.1.x (passwords cleared)"
                    If bootE IsNot Nothing AndAlso bootE.PwIs("20 C2 52 C3") Then Return "LDOS 5.1.0 or 5.1.1"
                    If bootE IsNot Nothing AndAlso bootE.Byte1A = &HFF Then Return "LDOS 5.1.4"
                    Return "LDOS 5.1.2 or 5.1.3"
                End If
                If cb = &H53 Then
                    If dirE IsNot Nothing AndAlso dirE.PwIs("AE 01 96 42") Then Return "LDOS 5.3.1"
                    Return "LDOS 5.3.0 or 5.3.1"
                End If
                If bootE IsNot Nothing Then
                    If bootE.PwIs("20 C2 52 C3") Then Return "LDOS 5.0.0-5.1.1 (GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & " is not FORMAT's; from the BOOT/SYS password)"
                    If bootE.PwIs("AE 01 F5 9C") Then Return "LDOS 5.1.2-5.1.4 (from the BOOT/SYS password)"
                    If bootE.PwIs("AE 01 00 00") Then Return "LDOS 5.3.x (from the BOOT/SYS password)"
                    If bootE.PwIs("5F 01 00 00") Then Return "LDOS 5.3.1 (from the BOOT/SYS password)"
                End If
                Return "LDOS 5 (version not readable from GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & " or the passwords)"
            Case "LDOS 5"
                ' Model III boot: same rules
                Return VersionFor(img, r, "VTOS 4 / LDOS 5", b, cb, cc, cd, ce, hit0, hit1, bootE, dirE, directoryOnly)
            Case "TRSDOS 6", "DOSPLUS 3.5/IV", "DOSPLUS IV"
                ' DOSPLUS IV first (shares 60 05 C2 with TRSDOS 6.0.0)
                Dim isDpIV As Boolean = (dirE IsNot Nothing AndAlso dirE.UpdatePwIs("DD 23")) OrElse (ce = "96 42" AndAlso cb = &H60) OrElse (img.SectorData(0, 0, 2) IsNot Nothing AndAlso Ascii(img.SectorData(0, 0, 2), 0, 3) = "DCT") OrElse fam = "DOSPLUS IV"
                If fam = "DOSPLUS 3.5/IV" Then
                    If isDpIV OrElse cb = &H60 Then
                        r.Family = "DOSPLUS IV" : r.Model = "Model 4"
                        Return "DOSPLUS IV 1.0"
                    End If
                    r.Family = "DOSPLUS" : r.Model = "Model I or III"
                    Return VersionFor(img, r, "DOSPLUS", b, cb, cc, cd, ce, hit0, hit1, bootE, dirE, directoryOnly)
                End If
                If isDpIV Then
                    r.Family = "DOSPLUS IV" : r.Model = "Model 4"
                    If cb = &H35 Then Return "Formatted by DOSPLUS 3.5; system DOSPLUS IV"
                    Return "DOSPLUS IV 1.0"
                End If
                r.Family = "TRSDOS 6"
                Dim vb As Integer = If(s2b0 >= &H60 AndAlso s2b0 <= &H63, s2b0, cb)
                ev.Add("Version byte: cylinder 0 sector 2 byte 0 " & If(s2b0 >= 0, s2b0.ToString("X2", CultureInfo.InvariantCulture), "?") & ", GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ", media block " & If(r.Gat IsNot Nothing AndAlso r.Gat(&HF5) = 3, """" & Ascii(r.Gat, &HF6, 3) & """", "none") & ", BOOT/SYS pw " & If(bootE IsNot Nothing, bootE.PwText, "—") & ".")
                If r.BootKind = "stub" Then
                    If Matches(b, "00 FE ?? F3 21 0E 43") Then
                        If vb = &H60 Then Return "TRSDOS 6.0.0"
                        Return "TRSDOS 6.01.00"
                    End If
                    If Matches(b, "00 FE ?? F3 00 00 00 00 00 00 00 00 00 00 00 00 11 87 FB") Then Return "TRSDOS 6.01.01 or 6.01.02"
                Else
                    If b.Length > &H8A Then
                        If b(&H48) = &H51 Then ev.Add("System boot build: byte 48 = 51 → 6.0.0.") : Return "TRSDOS 6.0.0"
                        If b(&H8A) = &HEE Then ev.Add("System boot build: byte 8A = EE → 6.01.00.") : Return "TRSDOS 6.01.00"
                        If b(&H38) = &H7B Then ev.Add("System boot build: byte 38 = 7B → 6.01.01 / 6.01.02.") : Return "TRSDOS 6.01.01 or 6.01.02"
                    End If
                End If
                Select Case vb
                    Case &H60 : Return "TRSDOS 6.0.0"
                    Case &H61 : Return "TRSDOS 6.01.x"
                    Case &H62
                        Dim mfm0 As Byte() = Nothing
                        Dim s0 As DmkSector = If(img.Track(0, 0) IsNot Nothing, img.Track(0, 0).FindSector(0, True), Nothing)
                        If s0 IsNot Nothing AndAlso s0.HasData Then mfm0 = s0.Data
                        If r.BootKind = "system" AndAlso mfm0 IsNot Nothing AndAlso mfm0.Length > &HF1 Then
                            If mfm0(&HF1) = 8 Then Return "TRSDOS 6.02.00"
                            If mfm0(&HF1) = &HA Then Return "TRSDOS 6.02.01"
                        End If
                        Return "TRSDOS 6.02.00 or 6.02.01"
                    Case &H63
                        If (r.Gat IsNot Nothing AndAlso Ascii(r.Gat, &HF6, 3) = "631") OrElse (bootE IsNot Nothing AndAlso bootE.PwIs("F4 71 00 00")) OrElse (bootE IsNot Nothing AndAlso bootE.UpdatePwIs("F4 71")) Then Return "LS-DOS 6.3.1"
                        If r.Gat IsNot Nothing AndAlso Ascii(r.Gat, &HF6, 3) = "LSI" Then Return "LS-DOS 6.3.0"
                        Return "LS-DOS 6.3.0 or 6.3.1"
                    Case Else
                        Return "TRSDOS 6 / LS-DOS 6 (version byte " & vb.ToString("X2", CultureInfo.InvariantCulture) & " not a known value)"
                End Select
            Case "RapidDOS 4"
                If cb = &H62 Then Return "RapidDOS 2.0 (Model 4)"
                Return "RapidDOS 1.3 or 1.5 (Model 4)"
            Case "NEWDOS/80 1.0"
                Return "NEWDOS/80 1.0"
            Case "NEWDOS/80 2.x", "NEWDOS/80 (directory)"
                If b(2) = &H18 Then Return "NEWDOS/80 2.1b (German)"
                If r.Entry("GDOS/SYS") IsNot Nothing Then Return "G-DOS (NEWDOS-80 II 2.1b)"
                If r.Entry("HDOS/SYS") IsNot Nothing Then Return "H-DOS 2.3c (Genie III)"
                If fam = "NEWDOS/80 (directory)" Then Return If(cb = &H80, "NEWDOS/80 1.0", "NEWDOS/80 2.0, 2.5 or 2.74")
                ev.Add("BOOT/SYS pw " & If(bootE IsNot Nothing, bootE.PwText, "—") & ": 56 78 12 34 = Apparat's passwords (a system copy, or 3.0's FORMAT); 60 7F 1F B2 = TRSDOS's (2.x FORMAT) (D4).")
                Return "NEWDOS/80 2.0, 2.5 or 2.74"
            Case "NEWDOS/80 3.0"
                Return "NEWDOS/80 3.0"
            Case "DOSPLUS"
                ' boot names a release, GAT CB names the FORMAT
                Dim bootRel As String = ""
                If r.BootKind = "system" Then
                    If Not r.BootIsFm Then
                        Select Case b(6)
                            Case &HC9 : bootRel = "3.3"
                            Case &HCE : bootRel = "3.4"
                            Case &HDD : bootRel = "3.5"
                        End Select
                    Else
                        If Matches(b, "00 FE ?? F3 31 FC 41 CD C9 01 3E 01 32 E1 37 3A 02 42 57") Then bootRel = "3.1"
                        If Matches(b, "F3 FE ?? 31 FC 41 CD C9 01 3A 02 42 57") Then bootRel = "3.2"
                        If Matches(b, "F3 FE ?? 31 FC 41 36 FF 36 D0 CD C9 01") Then bootRel = "3.2.1"
                        If Matches(b, "F3 FE ?? 31 FC 41 CD C9 01 36 FF AF 32 ED 37") Then bootRel = "3.3"
                        If Matches(b, "F3 FE ?? 31 FC 41 CD C9 01 11 02 00 01 00 50") Then bootRel = "3.4"
                        If Matches(b, "00 FE ?? CD C9 01 11 02 00 21 00 52 CD B2 42") Then bootRel = "3.5"
                    End If
                End If
                Dim fmtRel As String = ""
                Select Case cb
                    Case &H31 : fmtRel = "3.1"
                    Case &H32 : fmtRel = "3.2 or 3.2.1"
                    Case &H33 : fmtRel = "3.3"
                    Case &H34 : fmtRel = "3.4"
                    Case &H35 : fmtRel = "3.5"
                    Case &H40 : fmtRel = "3.4a or 3.4D"
                End Select
                If fmtRel.Length = 0 AndAlso bootE IsNot Nothing Then
                    If bootE.PwIs("7E B2 82 2A") Then fmtRel = "3.1 or 3.2 (from the passwords)"
                    If bootE.PwIs("27 0B 82 2A") Then fmtRel = "3.2.1-3.4 (from the passwords)"
                    If bootE.PwIs("DD 23 DD 23") Then fmtRel = "3.5 (from the passwords)"
                End If
                ev.Add("DOSPLUS: boot release " & If(bootRel.Length > 0, bootRel, "(stub or blank)") & "; GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & " → FORMAT " & If(fmtRel.Length > 0, fmtRel, "?") & "; passwords " & If(bootE IsNot Nothing, bootE.PwText, "—") & " / " & If(dirE IsNot Nothing, dirE.PwText, "—") & ".")
                If r.Model.Length = 0 Then r.Model = "Model I or III"
                If bootRel.Length > 0 AndAlso fmtRel.Length > 0 AndAlso Not fmtRel.StartsWith(bootRel, StringComparison.Ordinal) AndAlso Not (bootRel = "3.2.1" AndAlso fmtRel.StartsWith("3.2", StringComparison.Ordinal)) Then
                    Return "Formatted by DOSPLUS " & fmtRel & "; system DOSPLUS " & bootRel
                End If
                If bootRel.Length > 0 Then Return "DOSPLUS " & bootRel
                If fmtRel.Length > 0 Then Return "DOSPLUS " & fmtRel
                Return "DOSPLUS (version not readable)"
            Case "UltraDOS"
                Dim u As String = Ascii(r.Gat, &HCB, 3)
                If Matches(b, "00 FE ?? F3 31 00 42 CD C9 01 01 00 4D") Then Return "UltraDOS 4.1" & If(u = "U41", "", " (GAT says """ & u & """)")
                If u = "U42" OrElse Matches(b, "00 FE ?? F3 11 01 3C") OrElse Matches(b, "00 FE ?? F3 31 00 42 21 4A 42") Then Return "UltraDOS 4.2"
                If u = "U41" Then Return "UltraDOS 4.1"
                Return "UltraDOS"
            Case "DoubleDOS"
                If r.BootKind = "stub" Then
                    Dim b4f As Integer = b(&H4F)
                    ev.Add("DoubleDOS FORMAT stub: byte 4F = " & b4f.ToString("X2", CultureInfo.InvariantCulture) & " (30 = ""VERSION 1.20"" → 4.20 / 4.23; 31 = ""1.21"" → 4.24).")
                    If b4f = &H31 Then Return "DoubleDOS 4.24"
                    If b4f = &H30 Then Return "DoubleDOS 4.20 or 4.23"
                End If
                Return "DoubleDOS 4.20, 4.23 or 4.24"
            Case "Percom MicroDOS"
                Return "Percom MicroDOS (version not recorded on disk)"
            Case "Percom OS-80 III"
                Return "Percom OS-80 III (version not recorded on disk)"
            Case "MMSFORTH"
                Return If(r.BootKind = "system", "MMSFORTH system disk (not a DOS; banner in sectors 0-1)", "MMSFORTH data disk (no loader, no directory)")
            Case "MultiDOS 1.0/1.2D"
                If r.GatMark = &HFA Then Return "MultiDOS 1.0 (Model I)"
                If r.GatMark = &HF8 AndAlso img.Track(Math.Min(1, img.Cylinders - 1), 0) IsNot Nothing AndAlso img.Track(Math.Min(1, img.Cylinders - 1), 0).HasFm AndAlso Not img.Track(Math.Min(1, img.Cylinders - 1), 0).HasMfm Then Return "MultiDOS 1.2D (Model I)"
                Return "MultiDOS 1.0 or 1.2D (Model I)"
            Case "MultiDOS", "Esoteric"
                Return MultiDosVersion(r, b, cb, dirE, bootE, ev)
            Case Else
                Return fam
        End Select
    End Function

    Private Function MultiDosVersion(r As DosIdentification, b As Byte(), cb As Integer, dirE As DosDirEntry, bootE As DosDirEntry, ev As List(Of String)) As String
        Dim pw As String = If(dirE IsNot Nothing, dirE.PwText, "")
        ' Up to 2.0, DIR/SYS bytes 16-19 are password hashes; from 2.1 on they are the file's date stamp (the author,
        ' Vernon Hester, 16 November 2026): the release date on a system disk, the date set at boot on a bare disk.
        ' Read as a date only where GAT CB is FF (2.1 on): 1.7-2.0's 52 44 96 42 would also decode to a date.
        Dim stamp As String = If(dirE IsNot Nothing AndAlso cb = &HFF, MultiDosDateStamp(dirE.Password), "")
        ev.Add("MultiDOS: GAT CB " & cb.ToString("X2", CultureInfo.InvariantCulture) & ", F4-F5 """ & Ascii(r.Gat, &HF4, 2) & """, DIR/SYS bytes 16-19 " & If(pw.Length > 0, pw, "—") &
               If(stamp.Length > 0, " (as a 2.1-or-later date stamp: " & stamp & ")", "") &
               ", DIR/SYS bytes 1-2 " & If(dirE IsNot Nothing, dirE.Byte1.ToString("X2", CultureInfo.InvariantCulture) & " " & dirE.Byte2.ToString("X2", CultureInfo.InvariantCulture), "—") & ", boot byte FA " & b(&HFA).ToString("X2", CultureInfo.InvariantCulture) & ".")
        If r.BootKind = "stub" Then
            If Matches(b, "?? FE 11 CD 0B 00 11 0E 00 19 CD 67 44") Then
                If cb = &H50 Then Return "MultiDOS 1.0 (Model I)"
                If dirE IsNot Nothing Then
                    Dim b12 As String = dirE.Byte1.ToString("X2", CultureInfo.InvariantCulture) & " " & dirE.Byte2.ToString("X2", CultureInfo.InvariantCulture)
                    If b12 = "01 BA" Then Return "MultiDOS 1.0 (Model III)"
                    If b12 = "03 AA" Then Return "MultiDOS 1.1, 1.2 or 1.3"
                    If b12 = "04 F2" Then Return "MultiDOS 1.4 (or 1.5)"
                End If
                Return "MultiDOS 1.0-1.4"
            End If
            If Matches(b, "?? FE 11 CD 0B 00 11 2E 00 19 7E 23 B7") Then
                If cb = &H15 Then Return "MultiDOS 1.5"
                If cb = &H16 Then Return "MultiDOS 1.6"
                Return "MultiDOS 1.5 or 1.6"
            End If
            If Matches(b, "?? FE 11 3A 03 00 FE 17 28 24") Then
                If cb = &HFF Then Return "MultiDOS 2.10"
                If bootE IsNot Nothing Then Return "MultiDOS 2.0.0 or 2.00"
                Dim yr As String = Ascii(r.Gat, &HF4, 2)
                If yr = "84" Then Return "MultiDOS 1.7 or 1.7.0"
                If yr = "85" Then Return "MultiDOS 1.7.1 or 1.71"
                Return "MultiDOS 1.7.x or 2.0"
            End If
            If Matches(b, "?? FE 11 CD 0B 00 11 10 00 19 7E 23 CD 33 00 B7 20 F8") Then
                ' Stub D. Some 5.11 / 6.x system disks carry this boot too, and their DIR/SYS date stamp names the release.
                Dim release As String = MultiDosReleaseByDirSys(pw)
                If release IsNot Nothing AndAlso pw <> "A8 00 00 00" Then Return release
                ' A bare disk. A8 00 00 00 is the date 00/00/00 that FORMAT writes when no date was set at boot; with a
                ' date set, DIR/SYS carries that date instead, so the stub alone decides. Boot F8-FF are a copy of the
                ' drive's DCT and FA is the cylinder its head was on, so the FA split may follow the drive, not the
                ' release (the author, 16 November 2026): it is reported, but as unverified.
                If pw = "A8 00 00 00" Then
                    ev.Add("DIR/SYS date stamp 00/00/00: the disk was formatted with no date set.")
                ElseIf stamp.Length > 0 AndAlso stamp <> "no date" Then
                    ev.Add("DIR/SYS date stamp " & stamp & ": the date that was set when the disk was formatted (an undated format has A8 00 00 00).")
                End If
                Dim fa As String = b(&HFA).ToString("X2", CultureInfo.InvariantCulture)
                If b(&HFA) = 0 Then Return "MultiDOS 4.01 or later (4.01 or 5.00 if boot byte FA " & fa & " marks the version, which is unverified)"
                Return "MultiDOS 4.01 or later (5.10, 5.11, 6.00, 6.20 or Esoteric 2.11-3.20 if boot byte FA " & fa & " marks the version, which is unverified) — not separable on a bare disk"
            End If
        End If
        ' System disks: the DIR/SYS table (passwords up to 2.0, the release's date stamp from 2.1); a banner would win but is not read here.
        Dim byDirSys As String = MultiDosReleaseByDirSys(pw)
        If byDirSys IsNot Nothing Then Return byDirSys
        ' Model I system boots by prefix
        If Matches(b, "01 ?? 11 4D 32 20 40") Then Return "MultiDOS 1.0-1.2 (Model I)"
        If Matches(b, "01 ?? 11 4D 3A 00 42 3D 20") Then Return "MultiDOS 1.3 (Model I)"
        If Matches(b, "?? FE 11 3A 00 42") Then Return "MultiDOS 1.3P or 1.4 (Model I)"
        If Matches(b, "?? FE 11 ED 56 3A ?? 42") Then Return "MultiDOS 1.5 (Z'DOS 1.0) or 1.6 (Model I)"
        If Matches(b, "00 FE 11 ED 56 E5 DD E1 11 00 50 43 48 CD") Then
            If b(&HF) = &H71 OrElse b(&HF) = &H72 Then Return "MultiDOS 1.7, 1.7.0 or 1.71 (Model I)"
            If b(&HF) = &H74 Then Return If(b(&H1D) = 3, "MultiDOS 2.00 (Model I)", "MultiDOS 2.10, 4.01, 5.00 or 5.10 (Model I; by DIR/SYS password)")
        End If
        Select Case cb
            Case &H50 : Return "MultiDOS 1.0 or 1.2D (Model I)"
            Case &H10 : Return "MultiDOS 1.0 (Model III) to 1.4"
            Case &H15 : Return "MultiDOS 1.5"
            Case &H16 : Return "MultiDOS 1.6"
            Case &H17 : Return "MultiDOS 1.7.0-2.00"
            Case &HFF : Return "MultiDOS 2.10 or later / Esoteric"
        End Select
        Return "MultiDOS (version not readable; banners and GAT CB disagree on system disks)"
    End Function

    ''' <summary>
    ''' The MultiDOS / Esoteric release that DIR/SYS bytes 16-19 name on a system disk, or Nothing. Up to 2.0 the
    ''' bytes are password hashes; from 2.1 on they are the file's date stamp, which on a system disk is the
    ''' release date (confirmed by the author, 16 November 2026). A8 00 00 00 is no date at all.
    ''' </summary>
    Private Function MultiDosReleaseByDirSys(pw As String) As String
        Select Case pw
            Case "FF FF FF FF" : Return "MultiDOS 1.0-1.3 or 1.6 (Model III)"
            Case "89 AB CD EF" : Return "MultiDOS 1.4-1.6"
            Case "9A BC DE F0" : Return "MultiDOS 1.0 or 1.2D (Model I)"
            Case "52 44 96 42" : Return "MultiDOS 1.7-2.0"
            Case "52 44 00 07", "FF FF 00 07", "FF FF 00 14" : Return "MultiDOS 2.10"
            Case "B2 21 00 00" : Return "MultiDOS 4.01, 5.11 or Esoteric 2.11"
            Case "C4 B6 00 00" : Return "MultiDOS 5.00"
            Case "D4 3E 90 00" : Return "MultiDOS 5.10"
            Case "B1 01 00 00" : Return "MultiDOS 5.11 or Esoteric 2.12"
            Case "D2 44 00 00" : Return "MultiDOS 6.00 or Esoteric 3.00"
            Case "D6 63 7B C0" : Return "MultiDOS 6.20 or Esoteric 3.20"
            Case "A8 00 00 00" : Return "MultiDOS 4.01 or later, or Esoteric (DIR/SYS undated: formatted with no date set)"
        End Select
        If pw.StartsWith("BE EB", StringComparison.Ordinal) Then Return "Esoteric 2.12"
        If pw.StartsWith("BD 18", StringComparison.Ordinal) Then Return "Esoteric 2.11"
        Return Nothing
    End Function

    ''' <summary>
    ''' DIR/SYS bytes 16-17 read as a MultiDOS 2.1-and-later date stamp (the author's layout: xYYYYYYM MMMDDDDD,
    ''' the year less 1980, x set on system files): "MM/DD/YYYY", "no date" for 00/00/00 (A8 00), or "" when the
    ''' bytes are not a date (the password hashes of 1.x and 2.0 mostly aren't).
    ''' </summary>
    Friend Function MultiDosDateStamp(bytes As Byte()) As String
        If bytes Is Nothing OrElse bytes.Length < 2 Then Return ""
        Dim year As Integer = 1980 + ((bytes(0) >> 1) And &H3F)
        Dim month As Integer = ((bytes(0) And 1) << 3) Or (bytes(1) >> 5)
        Dim day As Integer = bytes(1) And &H1F
        If month = 0 AndAlso day = 0 Then Return "no date"
        If month < 1 OrElse month > 12 OrElse day < 1 OrElse day > 31 Then Return ""
        Return month.ToString("00", CultureInfo.InvariantCulture) & "/" & day.ToString("00", CultureInfo.InvariantCulture) & "/" & year.ToString(CultureInfo.InvariantCulture)
    End Function

    ' ------------------------------------------------------------------ cylinders the DOS records (media block)

    Private Sub ComputeDosCylinders(img As DmkImage, r As DosIdentification)
        Dim g As Byte() = r.Gat
        If g Is Nothing OrElse g.Length < 256 Then Return
        Dim cb As Integer = g(&HCB), cc As Integer = g(&HCC), cd As Integer = g(&HCD)
        Dim fam As String = r.Family
        ' TRSDOS 6.2+ / LS-DOS media block: FC = highest cylinder
        If g(&HF5) = 3 AndAlso (fam = "TRSDOS 6" OrElse fam = "RapidDOS 4") AndAlso g(&HFC) >= 34 AndAlso g(&HFC) <= 95 Then
            r.DosCylinders = g(&HFC) + 1
            r.DosCylinderSource = "GAT media block FC (highest cylinder " & g(&HFC).ToString(CultureInfo.InvariantCulture) & ")"
            Return
        End If
        ' LDOS line and DOSPLUS: CC = cylinders − 35, trusted when CB is 50–63 / 31–40 and the value is plausible
        If ((cb >= &H50 AndAlso cb <= &H63) OrElse (cb >= &H31 AndAlso cb <= &H40)) AndAlso cc + 35 >= 35 AndAlso cc + 35 <= 96 Then
            r.DosCylinders = cc + 35
            r.DosCylinderSource = "GAT CC (cylinders − 35 = " & cc.ToString(CultureInfo.InvariantCulture) & ")"
            Return
        End If
        ' MultiDOS 2.0+ keeps CC-CD in LDOS form
        If (fam = "MultiDOS" OrElse fam = "Esoteric") AndAlso cd <> &H28 AndAlso cd <> &H23 AndAlso cc + 35 >= 35 AndAlso cc + 35 <= 96 AndAlso (cd And &H40) <> 0 Then
            r.DosCylinders = cc + 35
            r.DosCylinderSource = "GAT CC (LDOS form, MultiDOS 2.0+)"
            Return
        End If
        ' DoubleDOS: CD = cylinders (46 28 = 70 lumps, 40 cylinders)
        If fam = "DoubleDOS" AndAlso cd >= 35 AndAlso cd <= 80 Then
            r.DosCylinders = cd
            r.DosCylinderSource = "GAT CD (DoubleDOS cylinder count)"
            Return
        End If
        If fam = "TRSDOS 1.x" OrElse fam = "RapidDOS III" Then
            r.DosCylinders = 40
            r.DosCylinderSource = "TRSDOS 1.x is always 40 cylinders"
            Return
        End If
        If fam = "TRSDOS 2.x" OrElse fam = "VTOS 3.0" OrElse fam = "TRSDOS 2.3b" OrElse fam = "UltraDOS" Then
            ' one GAT byte per cylinder; cylinders that do not exist read FF, as do full ones — take the last byte below 96 that is not FF, at least 35
            Dim last As Integer = 34
            For i As Integer = 35 To 95
                If g(i) <> &HFF Then last = i
            Next
            r.DosCylinders = last + 1
            r.DosCylinderSource = If(last = 34, "35 cylinders (2.3-family default; GAT bytes 35+ are FF)", "GAT allocation bytes in use up to cylinder " & last.ToString(CultureInfo.InvariantCulture))
            Return
        End If
        If fam = "TRSDOS 2.7/2.8" Then
            r.DosCylinders = If(r.BootSector IsNot Nothing AndAlso r.BootSector(&HFF) = &H28, 40, 35)
            r.DosCylinderSource = "boot byte FF (23 = 35, 28 = 40 cylinders)"
        End If
    End Sub

    ' ------------------------------------------------------------------ formatter fill

    Private Function FormatterFill(img As DmkImage, r As DosIdentification) As String
        ' Sample never-written sectors: the last data cylinder's sectors whose data is one repeating pattern.
        Dim counts As New Dictionary(Of String, Integer)
        For Each t In img.Tracks
            If t.Cylinder < 1 OrElse Not t.Present Then Continue For
            For Each s In t.Good
                If Not s.HasData OrElse s.Data.Length < 128 Then Continue For
                Dim p As String = RepeatPattern(s.Data)
                If p.Length = 0 Then Continue For
                If counts.ContainsKey(p) Then counts(p) += 1 Else counts(p) = 1
            Next
        Next
        If counts.Count = 0 Then Return ""
        Dim best = counts.OrderByDescending(Function(kv) kv.Value).First()
        Dim name As String
        Select Case best.Key
            Case "E5" : name = "E5 (TRSDOS 2.x, NEWDOS 2.1/3.0, VTOS, DOSPLUS 3.3, NEWDOS/80 2.x Model I, RapidDOS, TRSDOS 1.x/2.7/2.8, LDOS and MultiDOS 1.5+ single density)"
            Case "6D B6" : name = "6D B6 (TRSDOS 6 / LS-DOS, LDOS 5.1-5.3 and MultiDOS 1.5+ double density, DoubleDOS 4.24)"
            Case "6C" : name = "6C (DOSPLUS 3.5, DOSPLUS IV)"
            Case "00" : name = "00 (DOSPLUS 3.1, 3.2, 3.4)"
            Case "EC" : name = "EC (NEWDOS/80 1.0)"
            Case "5B" : name = "5B (NEWDOS/80 2.x and 3.0, Model III)"
            Case "F6" : name = "F6 (IBM PC-DOS)"
            Case Else
                name = best.Key & If(best.Key.StartsWith("4D 55 4C 54", StringComparison.Ordinal), " (""MULTIDOS"" repeated: MultiDOS 1.0-1.4)", If(best.Key.StartsWith("55 4C 54 52", StringComparison.Ordinal), " (""ULTRADOS"" repeated: UltraDOS)", ""))
        End Select
        Return name & " — " & best.Value.ToString(CultureInfo.InvariantCulture) & " never-written sector(s); a weak version witness on a lightly used disk"
    End Function

    ''' <summary>Returns the hex of the repeating unit (1, 2 or 8 bytes) if the sector is one pattern, else "".</summary>
    Private Function RepeatPattern(d As Byte()) As String
        For Each period In New Integer() {1, 2, 8}
            Dim ok As Boolean = True
            For i As Integer = period To d.Length - 1
                If d(i) <> d(i - period) Then ok = False : Exit For
            Next
            If ok Then
                If period = 1 AndAlso d(0) = 0 Then Return "00"
                Return Hex(d, 0, period)
            End If
        Next
        ' TRSDOS 1.x filler: E5 to EF then "(c) 1980 Tandy  "
        Dim allE5 As Boolean = True
        For i As Integer = 0 To &HEF
            If d(i) <> &HE5 Then allE5 = False : Exit For
        Next
        If allE5 Then Return "E5"
        Return ""
    End Function
End Module
