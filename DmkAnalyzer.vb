' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DmkAnalyzer.vb — GUI-side analysis of a DMK image: header, IDAM tables,
'  sectors with ID/data CRCs, geometry per track, formatter fill, protection
'  and special-format signs, damage vs imaging artifacts, tail tracks.
'
'  This is the one place the GUI reads disk bytes itself, to show disk
'  information beyond what trs80-tool displays. It never writes. Everything here is read-only
'  identification; disk *editing* still goes through trs80-tool.
'
'  The DOS identification itself is in DosIdentifier.vb.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>One sector found through an IDAM pointer.</summary>
Public Class DmkSector
    Public Property Cylinder As Integer          ' physical track index in the image
    Public Property Side As Integer              ' physical side in the image
    Public Property Mfm As Boolean               ' pointer bit 15
    Public Property Pointer As Integer           ' offset of the FE from the track start (incl. the 128-byte table)
    Public Property PointerIndex As Integer      ' slot in the IDAM table (0..63)
    Public Property IdCylinder As Integer
    Public Property IdSide As Integer
    Public Property IdSector As Integer
    Public Property SizeCode As Integer
    Public Property Size As Integer              ' bytes of data actually read for the CRC test
    Public Property NonIbmSize As Boolean        ' 1771 non-IBM: size = 16 × code (verified by a good CRC)
    Public Property IdCrcOk As Boolean
    Public Property DataMark As Integer          ' 0 = no data mark found
    Public Property DataOffset As Integer        ' logical offset (in track bytes) of the mark
    Public Property Data As Byte() = Nothing
    Public Property DataCrcOk As Boolean
    Public Property StoredCrc As Integer
    Public Property ComputedCrc As Integer
    Public Property SyncNote As String = ""      ' MFM sync bytes other than A1 A1 A1
    Public Property PointerNotAtFe As Boolean
    Public Property PointerPastEnd As Boolean

    Public ReadOnly Property HasData As Boolean
        Get
            Return DataMark <> 0 AndAlso Data IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property IsBad As Boolean
        Get
            Return Not IdCrcOk OrElse Not HasData OrElse Not DataCrcOk
        End Get
    End Property

    Public ReadOnly Property DensityLetter As String
        Get
            Return If(Mfm, "D", "S")
        End Get
    End Property

    ''' <summary>Map letter with trs80-tool's precedence: ? missing data, C CRC error, X deleted mark, S/D density.</summary>
    Public ReadOnly Property MapLetter As String
        Get
            If Not HasData Then Return "?"
            If Not IdCrcOk OrElse Not DataCrcOk Then Return "C"
            If DataMark = &HF8 OrElse DataMark = &HF9 Then Return "X"
            Return DensityLetter
        End Get
    End Property
End Class

''' <summary>One physical track (one side of one cylinder) of the image.</summary>
Public Class DmkTrack
    Public Property Cylinder As Integer
    Public Property Side As Integer
    Public Property Sectors As New List(Of DmkSector)
    Public Property PointerCount As Integer
    Public Property BadPointers As Integer       ' pointers not landing on an FE / past the end
    Public Property Present As Boolean = True    ' False when the file ends before this track
    Public Property Raw As Byte() = Nothing      ' the whole track incl. the 128-byte table (kept for gap/raw checks)

    Public ReadOnly Property Good As IEnumerable(Of DmkSector)
        Get
            Return Sectors.Where(Function(s) s.IdCrcOk)
        End Get
    End Property

    Public ReadOnly Property HasFm As Boolean
        Get
            Return Sectors.Any(Function(s) Not s.Mfm AndAlso s.IdCrcOk)
        End Get
    End Property

    Public ReadOnly Property HasMfm As Boolean
        Get
            Return Sectors.Any(Function(s) s.Mfm AndAlso s.IdCrcOk)
        End Get
    End Property

    ''' <summary>True when no ID on the track has a good CRC (unformatted or noise).</summary>
    Public ReadOnly Property Unformatted As Boolean
        Get
            Return Not Sectors.Any(Function(s) s.IdCrcOk)
        End Get
    End Property

    Public Function FindSector(number As Integer, mfm As Boolean?) As DmkSector
        For Each s In Sectors
            If s.IdSector = number AndAlso s.IdCrcOk AndAlso (Not mfm.HasValue OrElse s.Mfm = mfm.Value) Then Return s
        Next
        Return Nothing
    End Function

    ''' <summary>Short description: "FM 10×256 (0-9) FA/FB" style.</summary>
    Public Function Summary() As String
        If Not Present Then Return "(not in the file)"
        If Sectors.Count = 0 Then Return "no IDs"
        If Unformatted Then Return "unformatted (" & Sectors.Count.ToString(CultureInfo.InvariantCulture) & " IDs, none with a good CRC)"
        Dim parts As New List(Of String)
        For Each mfm In New Boolean() {False, True}
            Dim grp As List(Of DmkSector) = Sectors.Where(Function(s) s.IdCrcOk AndAlso s.Mfm = mfm).ToList()
            If grp.Count = 0 Then Continue For
            Dim sizes As String = String.Join("/", grp.Select(Function(s) s.Size).Distinct().OrderBy(Function(x) x).Select(Function(x) x.ToString(CultureInfo.InvariantCulture)))
            Dim nums As List(Of Integer) = grp.Select(Function(s) s.IdSector).OrderBy(Function(x) x).ToList()
            Dim range As String = If(nums.Count = 1, nums(0).ToString(CultureInfo.InvariantCulture), nums(0).ToString(CultureInfo.InvariantCulture) & "-" & nums(nums.Count - 1).ToString(CultureInfo.InvariantCulture))
            Dim marks As String = String.Join("/", grp.Where(Function(s) s.DataMark <> 0).Select(Function(s) s.DataMark).Distinct().OrderByDescending(Function(x) x).Select(Function(x) x.ToString("X2", CultureInfo.InvariantCulture)))
            parts.Add(If(mfm, "MFM ", "FM ") & grp.Count.ToString(CultureInfo.InvariantCulture) & "×" & sizes & " (" & range & ")" & If(marks.Length > 0, " " & marks, ""))
        Next
        Dim bad As Integer = Sectors.Where(Function(s) s.IsBad).Count()
        Return String.Join(" + ", parts) & If(bad > 0, "  " & bad.ToString(CultureInfo.InvariantCulture) & " bad", "")
    End Function
End Class

''' <summary>A DMK file parsed into tracks and sectors.</summary>
Public Class DmkImage
    Public Property Path As String = ""
    Public Property FileSize As Long
    ''' <summary>"DMK", "JV3", "JV1" or "unknown" — decided from the bytes, not the extension.</summary>
    Public Property Kind As String = "DMK"
    Public Property KindNote As String = ""
    Public Property WriteProtect As Byte
    Public Property HeaderTracks As Integer
    Public Property TrackLength As Integer
    Public Property Options As Byte
    Public Property Quirks As Byte
    Public Property RealDiskField As UInteger
    Public Property SingleSided As Boolean
    Public Property SingleByteFm As Boolean      ' header bit 6 or 7
    Public Property IgnoreDensity As Boolean     ' bit 7
    Public Property Sides As Integer = 2
    Public Property TracksInFile As Integer      ' whole tracks present (per side × cylinders)
    Public Property PartialTrailingBytes As Long ' bytes after the last whole track
    Public Property Cylinders As Integer         ' cylinders present in the file (ceiling)
    Public Property Tracks As New List(Of DmkTrack)
    Public Property Notes As New List(Of String) ' header/file-level observations

    Public Function Track(cylinder As Integer, side As Integer) As DmkTrack
        For Each t In Tracks
            If t.Cylinder = cylinder AndAlso t.Side = side Then Return t
        Next
        Return Nothing
    End Function

    Public Function Sector(cylinder As Integer, side As Integer, number As Integer, mfm As Boolean?) As DmkSector
        Dim t As DmkTrack = Track(cylinder, side)
        If t Is Nothing Then Return Nothing
        Return t.FindSector(number, mfm)
    End Function

    ''' <summary>Data bytes of a good-ID sector (any density), or Nothing.</summary>
    Public Function SectorData(cylinder As Integer, side As Integer, number As Integer) As Byte()
        Dim s As DmkSector = Sector(cylinder, side, number, Nothing)
        If s Is Nothing OrElse Not s.HasData Then Return Nothing
        Return s.Data
    End Function

    Public ReadOnly Property AllSectors As IEnumerable(Of DmkSector)
        Get
            Return Tracks.SelectMany(Function(t) t.Sectors)
        End Get
    End Property

    ' ------------------------------------------------------------------ CRC

    ''' <summary>CRC-16-CCITT (poly 1021), preset FFFF; MFM fields start from the CRC of A1 A1 A1 (CDB4).</summary>
    Public Shared Function Crc16(data As Byte(), start As Integer, count As Integer, mfm As Boolean) As Integer
        Dim crc As Integer = If(mfm, &HCDB4, &HFFFF)
        For i As Integer = start To start + count - 1
            crc = crc Xor (CInt(data(i)) << 8)
            For b As Integer = 0 To 7
                If (crc And &H8000) <> 0 Then
                    crc = ((crc << 1) Xor &H1021) And &HFFFF
                Else
                    crc = (crc << 1) And &HFFFF
                End If
            Next
        Next
        Return crc And &HFFFF
    End Function

    ' ------------------------------------------------------------------ loading

    Public Shared Function Load(path As String) As DmkImage
        Dim img As New DmkImage With {.Path = path}
        Dim bytes As Byte() = File.ReadAllBytes(path)
        img.FileSize = bytes.LongLength
        If bytes.Length < 16 + 128 Then
            img.Kind = "unknown"
            img.KindNote = "The file is too small to be a DMK image."
            Return img
        End If
        img.WriteProtect = bytes(0)
        img.HeaderTracks = bytes(1)
        img.TrackLength = bytes(2) Or (CInt(bytes(3)) << 8)
        img.Options = bytes(4)
        img.Quirks = bytes(5)
        img.RealDiskField = CUInt(bytes(12)) Or (CUInt(bytes(13)) << 8) Or (CUInt(bytes(14)) << 16) Or (CUInt(bytes(15)) << 24)
        img.SingleSided = (img.Options And &H10) <> 0
        img.SingleByteFm = (img.Options And &HC0) <> 0
        img.IgnoreDensity = (img.Options And &H80) <> 0
        img.Sides = If(img.SingleSided, 1, 2)

        ' Is it a DMK at all? (xtrs/MAME/dmk2gw/trs80-tool header checks; JV3 and JV1 named .dmk exist.)
        Dim headerOk As Boolean = (img.WriteProtect = 0 OrElse img.WriteProtect = &HFF) AndAlso img.TrackLength >= &H80 AndAlso img.TrackLength <= &H4000 AndAlso img.RealDiskField = 0
        If img.RealDiskField = &H12345678UI Then
            img.Kind = "DMK (real-disk specification)"
            img.KindNote = "Header bytes 12-15 = 12345678h: Keil's 'real disk' file, not a disk image."
            Return img
        End If
        If Not headerOk Then
            img.Kind = GuessNonDmk(bytes)
            img.KindNote = "The header is not a DMK header (write-protect byte " & img.WriteProtect.ToString("X2", CultureInfo.InvariantCulture) & ", track length " & img.TrackLength.ToString("X4", CultureInfo.InvariantCulture) & "h, bytes 12-15 " & img.RealDiskField.ToString("X8", CultureInfo.InvariantCulture) & "h). trs80-tool reads such a file as JV3 when it can."
            Return img
        End If

        Dim dataBytes As Long = bytes.LongLength - 16
        img.TracksInFile = CInt(dataBytes \ img.TrackLength)
        img.PartialTrailingBytes = dataBytes Mod img.TrackLength
        img.Cylinders = CInt(Math.Ceiling(img.TracksInFile / CDbl(img.Sides)))
        If img.TracksInFile = 0 Then
            img.Kind = GuessNonDmk(bytes)
            img.KindNote = "No whole track fits in the file for the header's track length; probably not a DMK."
            Return img
        End If

        For t As Integer = 0 To img.TracksInFile - 1
            Dim cyl As Integer = t \ img.Sides
            Dim side As Integer = t Mod img.Sides
            Dim offset As Integer = 16 + t * img.TrackLength
            Dim raw(img.TrackLength - 1) As Byte
            Array.Copy(bytes, offset, raw, 0, img.TrackLength)
            img.Tracks.Add(ParseTrack(raw, cyl, side, img.SingleByteFm))
        Next
        ' A missing last side-1 track in a two-sided image: mark it.
        If img.Sides = 2 AndAlso (img.TracksInFile Mod 2) = 1 Then
            img.Tracks.Add(New DmkTrack With {.Cylinder = img.Cylinders - 1, .Side = 1, .Present = False})
        End If

        ' Header-level notes
        If img.HeaderTracks <> img.Cylinders Then
            img.Notes.Add("Header says " & img.HeaderTracks.ToString(CultureInfo.InvariantCulture) & " tracks (cylinders); the file holds " & img.Cylinders.ToString(CultureInfo.InvariantCulture) & If(img.Cylinders > img.HeaderTracks, " — one or more uncounted extra tracks (gw2dmk read past the last cylinder)", " — the file is short; readers should load the tracks that are present"))
        End If
        If img.PartialTrailingBytes > 0 Then img.Notes.Add(img.PartialTrailingBytes.ToString("N0", CultureInfo.InvariantCulture) & " bytes after the last whole track (ignored).")
        If img.Quirks <> 0 Then img.Notes.Add("Header byte 5 (gw2dmk quirks) = " & img.Quirks.ToString("X2", CultureInfo.InvariantCulture) & "h — UDOS/CZ-SDOS quirk flags; not a TRS-80 image?")
        If img.IgnoreDensity Then img.Notes.Add("Header bit 7 (ignore density) set: an old emulator-made image; single-density bytes stored once.")
        If (img.Options And &H20) <> 0 Then img.Notes.Add("Header bit 5 (RX02) set — a gw2dmk extension, not TRS-80.")
        img.Notes.Add(TrackLengthNote(img.TrackLength))
        Return img
    End Function

    Private Shared Function GuessNonDmk(bytes As Byte()) As String
        ' JV1: a bare multiple of 2,560 bytes (10 × 256 per track), 35–80 tracks, and byte 0 is not a DMK WP byte.
        If bytes.LongLength Mod 2560 = 0 AndAlso bytes.LongLength \ 2560 >= 35 AndAlso bytes.LongLength \ 2560 <= 96 Then Return "JV1 (probably)"
        ' JV3: 2,901 sector headers of 3 bytes + a WP byte = 8,704-byte header block; first header looks like track 0 sector 0/1.
        If bytes.Length > 8704 AndAlso bytes(0) = 0 AndAlso (bytes(1) = 0 OrElse bytes(1) = 1) AndAlso bytes(2) < &H40 Then Return "JV3 (probably)"
        Return "unknown"
    End Function

    Private Shared Function TrackLengthNote(len As Integer) As String
        Dim h As String = len.ToString("X4", CultureInfo.InvariantCulture) & "h"
        Select Case len
            Case &H1980 : Return "Track length " & h & ": gw2dmk / cw2dmk reading real 5¼"" media (the most common archive value)."
            Case &H1900 : Return "Track length " & h & ": Keil-style emulator default for 5¼"" MFM."
            Case &HCC0 : Return "Track length " & h & ": emulator default for 5¼"" FM-only."
            Case &HD00 : Return "Track length " & h & ": gw2dmk 5¼"" FM preset."
            Case &H14E0, &H1560 : Return "Track length " & h & ": 8"" FM."
            Case &H2940, &H2A40, &H2980 : Return "Track length " & h & ": 8"" MFM."
            Case &H3180, &H3260 : Return "Track length " & h & ": 3½"" HD (gw2dmk)."
            Case &H80 : Return "Track length 0080h: meaningless — this is what JV3 files named .dmk show."
            Case Else
                If Math.Abs(len - &H1900) <= &H100 Then Return "Track length " & h & ": within 4 % of 1900h (an unknown imager; 18EAh/18EEh are common)."
                If Math.Abs(len - &HCC0) <= &H60 Then Return "Track length " & h & ": within 4 % of 0CC0h (an unknown FM imager)."
                Return "Track length " & h & " (unusual)."
        End Select
    End Function

    ''' <summary>Read logical bytes from a track: FM bytes are stored twice unless the header says once.</summary>
    Private Shared Function LogicalByte(raw As Byte(), physOffset As Integer, logicalIndex As Integer, doubled As Boolean) As Integer
        Dim p As Integer = physOffset + If(doubled, logicalIndex * 2, logicalIndex)
        If p < 0 OrElse p >= raw.Length Then Return -1
        Return raw(p)
    End Function

    Private Shared Function ParseTrack(raw As Byte(), cyl As Integer, side As Integer, singleByteFm As Boolean) As DmkTrack
        Dim t As New DmkTrack With {.Cylinder = cyl, .Side = side, .Raw = raw}
        For i As Integer = 0 To 63
            Dim entry As Integer = raw(i * 2) Or (CInt(raw(i * 2 + 1)) << 8)
            If entry = 0 OrElse entry = &HFFFF Then Exit For
            t.PointerCount += 1
            Dim mfm As Boolean = (entry And &H8000) <> 0
            Dim ptr As Integer = entry And &H3FFF
            Dim s As New DmkSector With {.Cylinder = cyl, .Side = side, .Mfm = mfm, .Pointer = ptr, .PointerIndex = i}
            Dim doubled As Boolean = (Not mfm) AndAlso (Not singleByteFm)
            ' 7-byte ID must fit
            Dim idEnd As Integer = ptr + If(doubled, 14, 7)
            If ptr < 128 OrElse idEnd > raw.Length Then
                s.PointerPastEnd = True
                t.BadPointers += 1
                t.Sectors.Add(s)
                Continue For
            End If
            If raw(ptr) <> &HFE Then
                s.PointerNotAtFe = True
                t.BadPointers += 1
                t.Sectors.Add(s)
                Continue For
            End If
            Dim id(6) As Byte
            For k As Integer = 0 To 6
                id(k) = CByte(LogicalByte(raw, ptr, k, doubled))
            Next
            s.IdCylinder = id(1) : s.IdSide = id(2) : s.IdSector = id(3) : s.SizeCode = id(4)
            Dim idCrc As Integer = Crc16(id, 0, 5, mfm)
            s.IdCrcOk = (idCrc = ((CInt(id(5)) << 8) Or id(6)))
            If mfm Then
                ' sync bytes before the FE (14/15/A9 happen; CRCs still use A1 ×3)
                If ptr >= 131 Then
                    Dim s1 As Byte = raw(ptr - 1), s2 As Byte = raw(ptr - 2), s3 As Byte = raw(ptr - 3)
                    If Not (s1 = &HA1 AndAlso s2 = &HA1 AndAlso s3 = &HA1) Then
                        s.SyncNote = "sync " & s3.ToString("X2", CultureInfo.InvariantCulture) & " " & s2.ToString("X2", CultureInfo.InvariantCulture) & " " & s1.ToString("X2", CultureInfo.InvariantCulture)
                    End If
                End If
            End If
            ' Data mark: FE+7 .. FE+54 logical bytes (trs80-tool's window; controllers use FE+7..FE+36/49), any F8–FB
            For k As Integer = 7 To 54
                Dim b As Integer = LogicalByte(raw, ptr, k, doubled)
                If b < 0 Then Exit For
                If b >= &HF8 AndAlso b <= &HFB Then
                    s.DataMark = b
                    s.DataOffset = k
                    Exit For
                End If
                If b = &HFE AndAlso k > 7 Then Exit For   ' ran into the next ID
            Next
            If s.DataMark <> 0 Then
                ' Size: standard 128 << (code & 3); 1771 non-IBM 16 × code when code > 3 and that CRC checks.
                Dim stdSize As Integer = 128 << (s.SizeCode And 3)
                Dim sizes As New List(Of Integer)
                If s.SizeCode > 3 Then sizes.Add(16 * s.SizeCode)
                sizes.Add(stdSize)
                For Each sz In sizes
                    Dim field(sz) As Byte   ' mark + data
                    field(0) = CByte(s.DataMark)
                    Dim ok As Boolean = True
                    For k As Integer = 1 To sz
                        Dim b As Integer = LogicalByte(raw, ptr, s.DataOffset + k, doubled)
                        If b < 0 Then ok = False : Exit For
                        field(k) = CByte(b)
                    Next
                    If Not ok Then Continue For
                    Dim c1 As Integer = LogicalByte(raw, ptr, s.DataOffset + sz + 1, doubled)
                    Dim c2 As Integer = LogicalByte(raw, ptr, s.DataOffset + sz + 2, doubled)
                    Dim stored As Integer = If(c1 < 0 OrElse c2 < 0, -1, (c1 << 8) Or c2)
                    Dim computed As Integer = Crc16(field, 0, sz + 1, mfm)
                    If stored = computed OrElse sz = sizes(sizes.Count - 1) Then
                        s.Size = sz
                        s.NonIbmSize = (sz <> stdSize)
                        Dim data(sz - 1) As Byte
                        Array.Copy(field, 1, data, 0, sz)
                        s.Data = data
                        s.StoredCrc = stored
                        s.ComputedCrc = computed
                        s.DataCrcOk = (stored = computed)
                        Exit For
                    End If
                Next
                If s.Data Is Nothing Then s.DataMark = 0   ' data runs off the track
            End If
            t.Sectors.Add(s)
        Next
        Return t
    End Function
End Class

''' <summary>One finding for the report: a section, a severity and text.</summary>
Public Class DmkFinding
    Public Property Section As String = ""
    Public Property Level As String = ""   ' "info", "sign", "protection", "damage", "artifact", "warning"
    Public Property Text As String = ""
    Public Sub New(section As String, level As String, text As String)
        Me.Section = section : Me.Level = level : Me.Text = text
    End Sub
End Class

''' <summary>Everything the GUI works out about a DMK, ready for DiskInfoForm.</summary>
Public Class DmkAnalysis
    Public Property Image As DmkImage
    Public Property Dos As DosIdentification = Nothing
    Public Property Facts As New List(Of KeyValuePair(Of String, String))     ' Summary tab rows (label, value), in order
    Public Property Findings As New List(Of DmkFinding)
    Public Property DosCylinders As Integer = 0        ' cylinders the DOS's own records say the disk has (0 = unknown)
    Public Property DosCylinderSource As String = ""

    Public Shared Function Analyze(path As String) As DmkAnalysis
        Dim a As New DmkAnalysis With {.Image = DmkImage.Load(path)}
        a.Run()
        Return a
    End Function

    Private Sub AddFact(label As String, value As String)
        Facts.Add(New KeyValuePair(Of String, String)(label, value))
    End Sub

    Private Sub Note(section As String, level As String, text As String)
        Findings.Add(New DmkFinding(section, level, text))
    End Sub

    Private Sub Run()
        Dim img As DmkImage = Image
        AddFact("File", img.Path)
        AddFact("Size", img.FileSize.ToString("N0", CultureInfo.InvariantCulture) & " bytes")
        If img.Kind <> "DMK" Then
            AddFact("Image format", img.Kind)
            Note("Image", "warning", img.KindNote)
            Note("Image", "info", "The GUI's DMK analysis applies to DMK images only; trs80-tool still reads JV1/JV3 images normally.")
            Return
        End If
        AddFact("Image format", "DMK")
        AddFact("Header", "write-protect " & img.WriteProtect.ToString("X2", CultureInfo.InvariantCulture) & "h (" & If(img.WriteProtect = &HFF, "protected", "writable") & "), " & img.HeaderTracks.ToString(CultureInfo.InvariantCulture) & " tracks, track length " & img.TrackLength.ToString("X4", CultureInfo.InvariantCulture) & "h (" & img.TrackLength.ToString("N0", CultureInfo.InvariantCulture) & "), options " & img.Options.ToString("X2", CultureInfo.InvariantCulture) & "h" & If(img.SingleSided, " single-sided", " two-sided") & If(img.SingleByteFm, ", FM bytes stored once", ", FM bytes stored twice"))
        AddFact("Tracks in file", img.TracksInFile.ToString(CultureInfo.InvariantCulture) & " (" & img.Cylinders.ToString(CultureInfo.InvariantCulture) & " cylinders × " & img.Sides.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 1, " side)", " sides)"))
        For Each n In img.Notes
            Note("Image", "info", n)
        Next

        ' ---- Geometry per track
        Dim allGood As List(Of DmkSector) = img.AllSectors.Where(Function(s) s.IdCrcOk).ToList()
        Dim fmCount As Integer = allGood.Where(Function(s) Not s.Mfm).Count()
        Dim mfmCount As Integer = allGood.Where(Function(s) s.Mfm).Count()
        AddFact("Sectors", allGood.Count.ToString(CultureInfo.InvariantCulture) & " with a good ID (" & fmCount.ToString(CultureInfo.InvariantCulture) & " FM, " & mfmCount.ToString(CultureInfo.InvariantCulture) & " MFM)" & If(img.AllSectors.Count() > allGood.Count, "; " & (img.AllSectors.Count() - allGood.Count).ToString(CultureInfo.InvariantCulture) & " IDs with a bad CRC or a bad pointer", ""))
        Dim side1Used As Boolean = img.Tracks.Any(Function(t) t.Side = 1 AndAlso Not t.Unformatted)
        If img.Sides = 2 AndAlso Not side1Used Then Note("Geometry", "info", "Two-sided image, but side 1 has no formatted track: a single-sided disk in a two-sided image.")
        Dim cyl0 As DmkTrack = img.Track(0, 0)
        If cyl0 IsNot Nothing Then AddFact("Cylinder 0", cyl0.Summary())
        ' Most common data-track layout (cylinders 1..): summary string with its count
        Dim layouts As New Dictionary(Of String, Integer)
        For Each t In img.Tracks
            If t.Cylinder = 0 OrElse Not t.Present Then Continue For
            Dim k As String = t.Summary()
            If k.EndsWith(" bad", StringComparison.Ordinal) Then k = k.Substring(0, k.LastIndexOf("  ", StringComparison.Ordinal))
            If layouts.ContainsKey(k) Then layouts(k) += 1 Else layouts(k) = 1
        Next
        If layouts.Count > 0 Then
            Dim ordered = layouts.OrderByDescending(Function(kv) kv.Value).ToList()
            AddFact("Data tracks", ordered(0).Key & "  (" & ordered(0).Value.ToString(CultureInfo.InvariantCulture) & " tracks)")
            For i As Integer = 1 To Math.Min(ordered.Count - 1, 5)
                AddFact("", ordered(i).Key & "  (" & ordered(i).Value.ToString(CultureInfo.InvariantCulture) & If(ordered(i).Value = 1, " track)", " tracks)"))
            Next
        End If
        Dim dirMarks As String = ""

        ' ---- DOS identification (DosIdentifier.vb)
        Dos = DosIdentifier.Identify(img)
        AddFact("DOS (from the bytes)", Dos.Result)
        If Dos.Model.Length > 0 Then AddFact("Model", Dos.Model)
        If Dos.BootSectorDescription.Length > 0 Then AddFact("Boot sector", Dos.BootSectorDescription)
        If Dos.DirectoryDescription.Length > 0 Then AddFact("Directory", Dos.DirectoryDescription)
        If Dos.DiskName.Length > 0 OrElse Dos.DiskDate.Length > 0 Then AddFact("Pack name / date", Dos.DiskName & "  " & Dos.DiskDate)
        If Dos.AutoCommand.Length > 0 Then AddFact("AUTO command", Dos.AutoCommand)
        If Dos.MasterPasswordText.Length > 0 Then AddFact("Master password", Dos.MasterPasswordText)
        If Dos.GatSummary.Length > 0 Then AddFact("GAT CB-CF", Dos.GatSummary)
        If Dos.HitSummary.Length > 0 Then AddFact("HIT 00-01", Dos.HitSummary)
        If Dos.EntriesSummary.Length > 0 Then AddFact("System entries", Dos.EntriesSummary)
        If Dos.PdriveLine.Length > 0 Then AddFact("PDRIVE (NEWDOS/80)", Dos.PdriveLine)
        If Dos.FillText.Length > 0 Then AddFact("Formatter fill", Dos.FillText)
        For Each e In Dos.Evidence
            Note("DOS identification", "info", e)
        Next
        DosCylinders = Dos.DosCylinders
        DosCylinderSource = Dos.DosCylinderSource
        If DosCylinders > 0 Then AddFact("Cylinders per the DOS", DosCylinders.ToString(CultureInfo.InvariantCulture) & " (" & DosCylinderSource & ")" & If(DosCylinders < img.Cylinders, " — the image has " & (img.Cylinders - DosCylinders).ToString(CultureInfo.InvariantCulture) & " more; see Tail tracks", If(DosCylinders > img.Cylinders, " — the image is short", "")))

        ' ---- Protection and format signs, damage, artifacts, tail tracks
        CheckBootLayout()
        CheckProtectionSigns()
        CheckDosLevelProtection()
        CheckDamageAndArtifacts()
        CheckTailTracks()
    End Sub

    ' ------------------------------------------------------------------ cylinder 0 layout

    Private Sub CheckBootLayout()
        Dim t As DmkTrack = Image.Track(0, 0)
        If t Is Nothing Then Return
        Dim fm0 As DmkSector = t.FindSector(0, False)
        Dim mfm1 As DmkSector = t.FindSector(1, True)
        If fm0 IsNot Nothing AndAlso mfm1 IsNot Nothing Then
            Note("Boot layout", "sign", "Cylinder 0 holds FM sector 0 (what a Model I ROM boots) and MFM sector 1 (what a Model III/4 ROM boots): the dual-boot layout. Whether both paths run is a loader question ([Dual Boot] needs the code traced or a boot test); the layout itself is not protection.")
        ElseIf fm0 IsNot Nothing AndAlso t.HasMfm Then
            Note("Boot layout", "info", "Cylinder 0 is mixed: FM sector 0 plus MFM sectors other than 1.")
        ElseIf fm0 Is Nothing AndAlso mfm1 Is Nothing Then
            If t.Unformatted Then
                Note("Boot layout", "warning", "Cylinder 0 has no readable ID: the disk cannot boot and no DOS can mount it from this image.")
            Else
                Note("Boot layout", "sign", "Cylinder 0 has neither FM sector 0 nor MFM sector 1: no TRS-80 ROM can boot it (CP/M, a flux-era format, or a renumbered boot track).")
            End If
        End If
        If Dos IsNot Nothing AndAlso Dos.BootKind = "none" AndAlso (fm0 IsNot Nothing OrElse mfm1 IsNot Nothing) Then
            Note("Boot layout", "sign", "The boot sector matches no DOS boot or FORMAT stub on record: a custom loader (self-booting program, transfer system, or a format no DOS reads). The classification tags [Self Booting]/[Special Format] need the loader traced; this analysis only reports the sign.")
        End If
    End Sub

    ' ------------------------------------------------------------------ protection signs

    Private Sub CheckProtectionSigns()
        Dim img As DmkImage = Image
        Dim protection As New List(Of String)
        Dim possible As New List(Of String)

        ' VTOS mini-sectors 7C–7F on cylinder 0 (either side)
        For Each side In {0, 1}
            Dim t As DmkTrack = img.Track(0, side)
            If t Is Nothing Then Continue For
            Dim vt As List(Of DmkSector) = t.Sectors.Where(Function(s) s.IdSector >= &H7C AndAlso s.IdSector <= &H7F).ToList()
            If vt.Count > 0 Then
                Dim desc As String = String.Join(", ", vt.Select(Function(s) s.IdSector.ToString(CultureInfo.InvariantCulture) & If(s.IdCrcOk, "", "(bad ID)") & If(s.HasData, If(s.DataCrcOk, " ok", " CRC error"), " no data")))
                Dim isVtos As Boolean = Dos IsNot Nothing AndAlso Dos.Family.StartsWith("VTOS", StringComparison.Ordinal)
                If isVtos AndAlso Dos.Family = "VTOS 4.0" AndAlso vt.Count = 1 AndAlso vt(0).IdSector = &H7C AndAlso vt(0).DataCrcOk Then
                    Note("Protection", "info", "Cylinder 0 side " & side.ToString(CultureInfo.InvariantCulture) & " sector 7Ch (124), a normal E5 sector: VTOS 4.0's FORMAT writes it; VTOS 4.0 is not protected.")
                Else
                    protection.Add("Cylinder 0 side " & side.ToString(CultureInfo.InvariantCulture) & " carries IDs 124-127 (7Ch-7Fh): " & desc & ". VTOS 3.0's non-IBM mini-sectors with deliberate CRC errors — the DOS's own protection, written on every disk VTOS formats: [DOS Copy Protected] (scheme S-101" & If(isVtos, "", ", variant: the same IDs on another DOS's disk") & "). Hands off: never repair or trim.")
                End If
            End If
        Next

        ' Tandy track 40 (cylinder 40, sectors 129–146)
        Dim t40 As DmkTrack = img.Track(40, 0)
        If t40 IsNot Nothing AndAlso t40.Present Then
            Dim tandy As List(Of DmkSector) = t40.Sectors.Where(Function(s) s.IdCrcOk AndAlso s.IdSector >= &H81 AndAlso s.IdSector <= &H92).ToList()
            If tandy.Count >= 12 Then
                Dim f8 As Boolean = tandy.Any(Function(s) s.DataMark = &HF8)
                Dim code As Boolean = False, overlays As Boolean = False
                For Each s In tandy
                    If Not s.HasData Then Continue For
                    Dim filler As Boolean = IsTandyFiller(s.Data)
                    If s.IdSector >= &H90 AndAlso Not filler Then code = True
                    If s.IdSector < &H90 AndAlso Not filler Then overlays = True
                Next
                If code Then
                    protection.Add("Cylinder 40 is Tandy's protection track (sectors 129-146" & If(f8, ", F8 marks", "") & ") with the TRS-80 Protection Utility code in sectors 144-146: scheme S-005 — [DOS Copy Protected] by Ira's ruling of 1 November 2026 (the code is Tandy's duplication tool that sets TRSDOS's limited-backup count). Never trim or alter this track.")
                ElseIf overlays Then
                    protection.Add("Cylinder 40 is Tandy's track layout (sectors 129-146" & If(f8, ", F8 marks", "") & ") with program data in sectors 129-143: Snapp's overlays (scheme S-006, XOR E5) — [Copy Protected]. Never trim or alter this track.")
                ElseIf f8 Then
                    Note("Protection", "sign", "Cylinder 40 has Tandy's layout (sectors 129-146) with F8 marks but filler only: the track of a protected master whose code is gone (a hacked copy, like Demon Seed [Hacked]) — not protected as it stands. Never trim it.")
                Else
                    Note("Protection", "info", "Cylinder 40 has Tandy's layout (sectors 129-146) with FB marks and filler only: TRSDOS 1.1 FORMAT output (it formats cylinder 40 on every disk), not protection. Never trim it.")
                End If
            End If
        End If

        ' Per-track sector-number and density checks on data cylinders
        Dim usualRange As Func(Of DmkSector, Boolean) = Function(s) s.IdSector >= 0 AndAlso s.IdSector <= 18
        Dim renumbered As New List(Of String), dupTracks As New List(Of String), mixedTracks As New List(Of String), farTracks As New List(Of String), bigSize As New List(Of String), otherCyl As New List(Of String), f9Tracks As New List(Of String), faTracks As New List(Of String)
        For Each t In img.Tracks
            If Not t.Present Then Continue For
            Dim good As List(Of DmkSector) = t.Good.ToList()
            If good.Count = 0 Then Continue For
            Dim label As String = "cyl " & t.Cylinder.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, "/s" & t.Side.ToString(CultureInfo.InvariantCulture), "")
            ' duplicates in the same density
            Dim dups = good.GroupBy(Function(s) (s.Mfm, s.IdSector)).Where(Function(g) g.Count() > 1).ToList()
            If dups.Count > 0 Then dupTracks.Add(label & " (" & String.Join(",", dups.Select(Function(g) g.Key.IdSector.ToString(CultureInfo.InvariantCulture))) & ")")
            ' mixed density on a data track
            If t.Cylinder >= 1 AndAlso t.HasFm AndAlso t.HasMfm Then mixedTracks.Add(label)
            ' size code above 3 (without a 1771 non-IBM reading that checked)
            For Each s In good
                If s.SizeCode > 3 Then bigSize.Add(label & " sector " & s.IdSector.ToString(CultureInfo.InvariantCulture) & " size code " & s.SizeCode.ToString(CultureInfo.InvariantCulture) & If(s.NonIbmSize AndAlso s.DataCrcOk, " (1771 non-IBM, " & s.Size.ToString(CultureInfo.InvariantCulture) & " bytes, CRC good)", ""))
            Next
            ' sector numbers far outside the range on a full track / renumbered
            Dim outside As List(Of DmkSector) = good.Where(Function(s) Not usualRange(s)).ToList()
            If outside.Count > 0 Then
                If outside.Count = good.Count AndAlso good.Count >= 5 Then
                    farTracks.Add(label & " (" & String.Join(",", good.Select(Function(s) s.IdSector).Distinct().OrderBy(Function(x) x).Take(8).Select(Function(x) x.ToString(CultureInfo.InvariantCulture))) & If(good.Select(Function(s) s.IdSector).Distinct().Count() > 8, ",…", "") & ")")
                ElseIf good.Count >= 8 Then
                    renumbered.Add(label & " (" & String.Join(",", outside.Select(Function(s) s.IdSector.ToString(CultureInfo.InvariantCulture))) & " among " & good.Count.ToString(CultureInfo.InvariantCulture) & " sectors)")
                End If
            End If
            ' IDs naming another cylinder (ignore side-1 tracks whose IDs name side 0: single-sided formatter, not protection)
            Dim wrongCyl As Integer = good.Where(Function(s) s.IdCylinder <> t.Cylinder).Count()
            If wrongCyl > 0 AndAlso wrongCyl >= good.Count \ 2 AndAlso (DosCylinders = 0 OrElse t.Cylinder < DosCylinders) Then otherCyl.Add(label & " → ID cylinder " & String.Join("/", good.Where(Function(s) s.IdCylinder <> t.Cylinder).Select(Function(s) s.IdCylinder).Distinct().Take(3).Select(Function(x) x.ToString(CultureInfo.InvariantCulture))))
            ' F9 / FA marks
            If good.Any(Function(s) s.DataMark = &HF9) Then f9Tracks.Add(label)
            If good.Any(Function(s) s.DataMark = &HFA) Then faTracks.Add(label)
        Next
        If renumbered.Count > 0 Then protection.Add("Renumbered sector(s) — a full track with one number replaced by an out-of-range one, the most selective protection sign: " & JoinLimited(renumbered, 6) & ". Schemes of this kind: S-026 QuikPro+, S-028 Cdex, S-029 Talking Adventure I, S-030 Test Bank (one unreadable sector), S-032 TanData.")
        If farTracks.Count > 0 Then protection.Add("Whole tracks numbered outside 0-18 — " & farTracks.Count.ToString(CultureInfo.InvariantCulture) & " track(s): " & JoinLimited(farTracks, 5) & ". Renumbered-block schemes (S-010 Asylum, S-012 Blackjack Master 54-63, S-018 Frogger/Zaxxon 100-109, S-022 MyChess steps of 16, S-023 Monty 16-25, S-025 RSM3-I, S-027 SFinks) or Tandy/VTOS tracks. On a damaged track this would be noise; here the IDs have good CRCs.")
        If dupTracks.Count > 0 Then
            If dupTracks.Count >= 2 Then
                protection.Add("Duplicate sector numbers in the same density on " & dupTracks.Count.ToString(CultureInfo.InvariantCulture) & " tracks: " & JoinLimited(dupTracks, 6) & " (protection when on two or more tracks; S-015 CopyCat, S-021 Master Reversi). An 80-track read of a 40-track disk also makes duplicates — check the Tracks tab for odd/even pairs.")
            Else
                possible.Add("One track with a duplicate sector number: " & dupTracks(0) & " — a single duplicate standing in for a missing number is damage.")
            End If
        End If
        If mixedTracks.Count > 0 Then protection.Add("Single and double density on the same data track (cylinder 1 or later): " & JoinLimited(mixedTracks, 8) & ". No 179x DOS formatter writes this (S-024 Super Utility / PowerTool loader track, S-034 HyperZap). On end-of-disk tracks next to IDs naming other cylinders it would be an overwritten old format instead.")
        If bigSize.Count > 0 Then protection.Add("Size code above 3 on a good ID: " & JoinLimited(bigSize, 6) & " (protection; 1771 non-IBM sectors are S-009 Arranger / S-011 Acorn Basketball; a size byte 30h (768) with a good CRC is the S-008 six-ID track 3).")
        If otherCyl.Count > 0 Then possible.Add("IDs naming another cylinder on " & otherCyl.Count.ToString(CultureInfo.InvariantCulture) & " track(s): " & JoinLimited(otherCyl, 6) & " — possible protection (S-003 Big Five, S-017 Decathlon, S-018 Frogger cylinder 1, S-019 CRAS, S-024/S-033 track-register presets), or a re-read / stale track at the end of the image (see Tail tracks).")
        If f9Tracks.Count > 0 Then possible.Add("F9 data marks on " & f9Tracks.Count.ToString(CultureInfo.InvariantCulture) & " track(s): " & JoinLimited(f9Tracks, 8) & " — only a 1771 writes F9 in FM; required by Infocom's Model I loader (S-016) and Decathlon (S-017); in MFM usually a misread FB/F8 (test 2 before reading anything into it).")
        If faTracks.Count > 0 Then
            Dim dirCyl As Integer = If(Dos IsNot Nothing, Dos.DirectoryCylinder, -1)
            Dim offDir As List(Of String) = faTracks.Where(Function(x) Not x.StartsWith("cyl " & dirCyl.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)).ToList()
            If offDir.Count > 0 Then possible.Add("FA data marks outside the directory cylinder on: " & JoinLimited(offDir, 8) & " — FA is normal on Model I TRSDOS 2.x / NEWDOS 2.1 / VTOS / NEWDOS-80 SD directories and on VTOS's cylinder-0 GAT copy; elsewhere it is a 1771-only mark worth a look (S-023 Monty Plays Scrabble).")
        End If

        ' Hidden extra IDs packed on end tracks (S-031 Vivace!), or many more IDs than the layout
        For Each t In img.Tracks
            If Not t.Present OrElse t.Cylinder < 1 Then Continue For
            Dim good As Integer = t.Good.Count()
            If t.HasMfm AndAlso Not t.HasFm AndAlso good >= 30 AndAlso t.Good.All(Function(s) s.Size = 256) Then
                protection.Add("cyl " & t.Cylinder.ToString(CultureInfo.InvariantCulture) & " has " & good.ToString(CultureInfo.InvariantCulture) & " MFM 256-byte IDs with good CRCs on a track built for 18: hidden sectors inside other sectors' data fields (S-031 Vivace! v3.0).")
            End If
        Next

        For Each p In protection
            Note("Protection", "protection", p)
        Next
        For Each p In possible
            Note("Protection", "sign", p)
        Next
        If protection.Count = 0 AndAlso possible.Count = 0 Then
            Note("Protection", "info", "No format-level protection sign: no renumbered or out-of-range sectors, no duplicate numbers, no mixed-density data track, no size code above 3, no VTOS or Tandy track. File-system and timing protection leave no such sign, so a clean scan does not prove an unprotected disk.")
        End If
    End Sub

    Private Shared Function IsTandyFiller(d As Byte()) As Boolean
        If d Is Nothing OrElse d.Length < 256 Then Return False
        For i As Integer = 0 To &HEF
            If d(i) <> &HE5 Then Return False
        Next
        Dim tail As String = Encoding.ASCII.GetString(d, &HF0, 16)
        Return tail.StartsWith("(c) 19", StringComparison.Ordinal) OrElse tail.Trim().Length = 0
    End Function

    Private Shared Function JoinLimited(items As List(Of String), max As Integer) As String
        If items.Count <= max Then Return String.Join("; ", items)
        Return String.Join("; ", items.Take(max)) & "; … (" & (items.Count - max).ToString(CultureInfo.InvariantCulture) & " more)"
    End Function

    ' ------------------------------------------------------------------ DOS-level protection

    Private Sub CheckDosLevelProtection()
        If Dos Is Nothing Then Return
        ' TRSDOS 1.x limited-backup count: boot byte 22h (1.2/1.3) or word FC-FD (1.1), with bit-5 files
        If Dos.Family.StartsWith("TRSDOS 1.", StringComparison.Ordinal) AndAlso Dos.BootSector IsNot Nothing Then
            Dim b22 As Integer = Dos.BootSector(&H22)
            Dim wFC As Integer = Dos.BootSector(&HFC) Or (CInt(Dos.BootSector(&HFD)) << 8)
            Dim bit5Files As List(Of String) = Dos.LimitedBackupFiles
            Dim is11 As Boolean = Dos.BootSector(&HFE) = &H11
            Dim countSet As Boolean = If(is11, wFC <> &HFFFF, b22 <> &HFF)
            Dim countText As String = If(is11, "boot word FCh-FDh = " & wFC.ToString("X4", CultureInfo.InvariantCulture) & "h", "boot byte 22h = " & b22.ToString("X2", CultureInfo.InvariantCulture) & "h")
            If countSet AndAlso bit5Files.Count > 0 Then
                Note("Protection", "protection", "TRSDOS 1.x limited-backup count is set (" & countText & ") and " & bit5Files.Count.ToString(CultureInfo.InvariantCulture) & " file(s) carry the limited-backup bit (entry byte 0 bit 5): " & JoinLimited(bit5Files, 8) & ". Scheme S-102 — [DOS Copy Protected] (Ira's ruling of 30 October 2026). BACKUP decrements the count and, at 0, drops those files from the copy.")
            ElseIf countSet Then
                Note("Protection", "info", "TRSDOS 1.x limited-backup count is set (" & countText & ") but no file carries bit 5: it limits nothing — no tag (Ira's ruling).")
            ElseIf bit5Files.Count > 0 Then
                Note("Protection", "info", bit5Files.Count.ToString(CultureInfo.InvariantCulture) & " file(s) carry the limited-backup bit but the count is FF (no limit): no tag. Files: " & JoinLimited(bit5Files, 6))
            End If
        End If
        ' Relocated / decoy directory
        If Dos.DirectoryByteCylinder >= 0 AndAlso Dos.DirectoryCylinder >= 0 AndAlso Dos.DirectoryByteCylinder <> Dos.DirectoryCylinder AndAlso Not Dos.LumpLayout Then
            Note("Protection", "sign", "The boot sector says the directory is on cylinder " & Dos.DirectoryByteCylinder.ToString(CultureInfo.InvariantCulture) & " but the DIR/SYS entry was found on cylinder " & Dos.DirectoryCylinder.ToString(CultureInfo.InvariantCulture) & ": a relocated directory or a decoy (Super Utility + Copy Utility [Set 3]), or a density variant. Worth a look.")
        End If
        If Dos.HiddenInUseEntries > 0 Then
            Note("Protection", "sign", Dos.HiddenInUseEntries.ToString(CultureInfo.InvariantCulture) & " directory entr" & If(Dos.HiddenInUseEntries = 1, "y has", "ies have") & " a name but the in-use bit clear (attribute 8xh): DIR does not list them while the DOS can still run them — a hidden entry (Clone-III) or ordinary deleted files.")
        End If
        If Dos.ShortFiles.Count > 0 Then
            Note("Protection", "sign", "Files whose extents hold fewer sectors than their sector count claims: " & JoinLimited(Dos.ShortFiles, 6) & " — files deliberately shorter than their directory entries are file-system protection (Stellar Escort, Weerd; S-002 Music Magic inflates BOOT/SYS the other way), or damage to the entry.")
        End If
    End Sub

    ' ------------------------------------------------------------------ damage and imaging artifacts

    Private Sub CheckDamageAndArtifacts()
        Dim img As DmkImage = Image
        Dim bad As List(Of DmkSector) = img.AllSectors.Where(Function(s) s.IsBad).ToList()
        Dim badPointers As Integer = img.Tracks.Sum(Function(t) t.BadPointers)
        Dim noId As Integer = bad.Where(Function(s) Not s.IdCrcOk AndAlso Not s.PointerNotAtFe AndAlso Not s.PointerPastEnd).Count()
        Dim noData As Integer = bad.Where(Function(s) s.IdCrcOk AndAlso Not s.HasData).Count()
        Dim dataCrc As Integer = bad.Where(Function(s) s.IdCrcOk AndAlso s.HasData AndAlso Not s.DataCrcOk).Count()
        AddFact("Bad sectors", If(bad.Count = 0, "none", bad.Count.ToString(CultureInfo.InvariantCulture) & " (" & dataCrc.ToString(CultureInfo.InvariantCulture) & " data CRC errors, " & noData.ToString(CultureInfo.InvariantCulture) & " good IDs without a data field, " & noId.ToString(CultureInfo.InvariantCulture) & " IDs with a bad CRC" & If(badPointers > 0, ", " & badPointers.ToString(CultureInfo.InvariantCulture) & " bad IDAM pointers", "") & ")"))
        If bad.Count = 0 AndAlso badPointers = 0 Then
            Note("Damage", "info", "Every pointed ID and data field has a good CRC: no damage visible at the sector level.")
            Return
        End If

        ' Converter marker: stored XOR computed = 00FF on every data-CRC-error sector → the CRCs were made by a JV3→DMK converter
        Dim crcErrs As List(Of DmkSector) = bad.Where(Function(s) s.IdCrcOk AndAlso s.HasData AndAlso Not s.DataCrcOk).ToList()
        If crcErrs.Count > 0 AndAlso crcErrs.All(Function(s) ((s.StoredCrc Xor s.ComputedCrc) And &HFFFF) = &HFF) Then
            Note("Damage", "artifact", "Every data-CRC-error sector has stored CRC = computed CRC with the low byte inverted (the 00FF marker): the CRCs were written by whatever converted a sector image (JV3's CRC-error flag) to DMK, so no CRC test can apply to this image. The data of those sectors is as the original imager read it.")
        End If

        ' Unformatted tracks (every ID bad) vs damaged tracks
        Dim unformatted As List(Of DmkTrack) = img.Tracks.Where(Function(t) t.Present AndAlso t.Sectors.Count > 0 AndAlso t.Unformatted).ToList()
        If unformatted.Count > 0 Then
            Note("Damage", "artifact", unformatted.Count.ToString(CultureInfo.InvariantCulture) & " track(s) with IDs but none with a good CRC (noise on unformatted tracks or a game's Read-Track nibble tracks, S-008): " & JoinLimited(unformatted.Select(Function(t) "cyl " & t.Cylinder.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, "/s" & t.Side.ToString(CultureInfo.InvariantCulture), "")).ToList(), 10))
        End If
        If badPointers > 0 Then
            Note("Damage", "artifact", badPointers.ToString(CultureInfo.InvariantCulture) & " IDAM pointer(s) do not land on an FE or run past the track end: imager artifacts (FE bytes in data recorded as IDs, stale tables, pointers past the end). MAME would reject the whole image; xtrs skips them; trs80-tool drops them.")
        End If
        ' Garbled ID past the index: bad-ID, no-data sector beyond ~3,125 FM bytes / ~6,250 MFM with the same bytes on several tracks
        Dim lateIds As List(Of DmkSector) = bad.Where(Function(s) Not s.IdCrcOk AndAlso Not s.PointerNotAtFe AndAlso Not s.PointerPastEnd AndAlso s.DataMark = 0 AndAlso s.Pointer - 128 > If(s.Mfm, 6100, If(img.SingleByteFm, 3050, 6100))).ToList()
        If lateIds.Count >= 2 Then
            Note("Damage", "artifact", lateIds.Count.ToString(CultureInfo.InvariantCulture) & " bad-CRC IDs without data sit past one revolution's worth of track (a garbled ID decoded from past the index on a read longer than one revolution; the same bytes end every track of that read). Not damage.")
        End If
        ' Sync bytes other than A1
        Dim syncs As Integer = img.AllSectors.Where(Function(s) s.SyncNote.Length > 0 AndAlso s.IdCrcOk).Count()
        If syncs > 0 Then Note("Damage", "info", syncs.ToString(CultureInfo.InvariantCulture) & " MFM ID(s) are preceded by sync bytes other than A1 A1 A1 (14/15/A9 stored by some imagers). CRCs are computed over A1 ×3 regardless.")
        ' Read that began late (first slot missing, first ID hundreds of bytes in)
        For Each t In img.Tracks
            If Not t.Present OrElse t.Good.Count() = 0 Then Continue For
            Dim first As Integer = t.Good.Min(Function(s) s.Pointer) - 128
            Dim ref As Integer = -1
            Dim prev As DmkTrack = img.Track(t.Cylinder - 1, t.Side)
            If prev IsNot Nothing AndAlso prev.Present AndAlso prev.Good.Count() > 0 Then ref = prev.Good.Min(Function(s) s.Pointer) - 128
            If ref >= 0 AndAlso first - ref > 200 AndAlso t.Good.Count() = prev.Good.Count() - 1 Then
                Note("Damage", "artifact", "cyl " & t.Cylinder.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, "/s" & t.Side.ToString(CultureInfo.InvariantCulture), "") & ": one sector fewer than the neighbouring track and the first ID " & (first - ref).ToString(CultureInfo.InvariantCulture) & " bytes later — a read that began late (cw2dmk bug fixed in 4.6.0). The disk has the sector; the read missed it.")
            End If
        Next
        ' Last sector cut at the same position on every track (read one revolution long)
        ' List the bad sectors themselves (first 20)
        Dim lines As New List(Of String)
        For Each s In bad.Take(40)
            Dim what As String
            If s.PointerPastEnd Then
                what = "pointer past the track end"
            ElseIf s.PointerNotAtFe Then
                what = "pointer not at an FE"
            ElseIf Not s.IdCrcOk Then
                what = "ID CRC error" & If(s.HasData, "", ", no data field")
            ElseIf Not s.HasData Then
                what = "no data mark within the window"
            Else
                what = "data CRC error (stored " & s.StoredCrc.ToString("X4", CultureInfo.InvariantCulture) & ", computed " & s.ComputedCrc.ToString("X4", CultureInfo.InvariantCulture) & ")" & If(IsRound(s.StoredCrc), " — a round stored CRC: a deliberate Write-Track value if the data is clean filler (test 5)", "")
            End If
            lines.Add("cyl " & s.Cylinder.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, "/s" & s.Side.ToString(CultureInfo.InvariantCulture), "") & " " & If(s.Mfm, "MFM", "FM") & " sector " & s.IdSector.ToString(CultureInfo.InvariantCulture) & ": " & what)
        Next
        If lines.Count > 0 Then Note("Damage", "damage", "Bad sectors" & If(bad.Count > 40, " (first 40 of " & bad.Count.ToString(CultureInfo.InvariantCulture) & ")", "") & ":" & vbCrLf & "    " & String.Join(vbCrLf & "    ", lines))
        Note("Damage", "info", "Whether a bad sector is damage or protection follows these tests (the code checks it; corrected bytes reproduce the stored CRC; two reads disagree; separate originals agree with clean values; another copy is good; the sector is in free space; a user copy on another DOS; garble mid-content; physical geometry; bit-loss patterns). trs80-tool's sectors --bad and check (Disk menu) name the files at each bad sector.")
    End Sub

    Private Shared Function IsRound(crc As Integer) As Boolean
        Return crc = 0 OrElse crc = &HFFFF OrElse crc = &HE5E5 OrElse crc = &H4E4E
    End Function

    ' ------------------------------------------------------------------ tail tracks

    Private Sub CheckTailTracks()
        Dim img As DmkImage = Image
        If DosCylinders <= 0 Then
            Note("Tail tracks", "info", "The DOS's own records do not give a disk size here, so no tail track can be judged (Ira's rule: no identified DOS, no trim).")
            Return
        End If
        If img.Cylinders <= DosCylinders Then
            Note("Tail tracks", "info", "No cylinders beyond the " & DosCylinders.ToString(CultureInfo.InvariantCulture) & " the DOS records (" & DosCylinderSource & ").")
            Return
        End If
        Dim lines As New List(Of String)
        For c As Integer = DosCylinders To img.Cylinders - 1
            For side As Integer = 0 To img.Sides - 1
                Dim t As DmkTrack = img.Track(c, side)
                If t Is Nothing OrElse Not t.Present Then Continue For
                Dim label As String = "cyl " & c.ToString(CultureInfo.InvariantCulture) & If(img.Sides = 2, "/s" & side.ToString(CultureInfo.InvariantCulture), "")
                If t.Sectors.Count = 0 OrElse t.Unformatted Then
                    lines.Add(label & ": unformatted (no good IDs)")
                    Continue For
                End If
                Dim good As List(Of DmkSector) = t.Good.ToList()
                Dim idCyls As List(Of Integer) = good.Select(Function(s) s.IdCylinder).Distinct().ToList()
                If idCyls.Count = 1 AndAlso idCyls(0) < c Then
                    ' re-read of an earlier cylinder? compare data
                    Dim src As DmkTrack = img.Track(idCyls(0), side)
                    Dim same As Boolean = src IsNot Nothing AndAlso src.Present AndAlso good.All(Function(s) SameData(s, src.FindSector(s.IdSector, s.Mfm)))
                    lines.Add(label & ": IDs name cylinder " & idCyls(0).ToString(CultureInfo.InvariantCulture) & If(same, " and the data matches it — a re-read", " but the data differs — review"))
                ElseIf idCyls.Count = 1 AndAlso idCyls(0) = c Then
                    Dim sameFormat As Boolean = LayoutMatchesDataTracks(t)
                    lines.Add(label & ": formatted as cylinder " & c.ToString(CultureInfo.InvariantCulture) & If(sameFormat, " in the disk's own format — an unused same-format tail (never trimmed, Ira's decision)", " in another format (" & t.Summary() & ") — an earlier format left on the diskette; review only"))
                Else
                    lines.Add(label & ": " & t.Summary() & " — mixed ID cylinders; review")
                End If
            Next
        Next
        Dim tailCount As Integer = img.Cylinders - DosCylinders
        Note("Tail tracks", If(tailCount >= 7, "warning", "info"), tailCount.ToString(CultureInfo.InvariantCulture) & " cylinder(s) beyond the " & DosCylinders.ToString(CultureInfo.InvariantCulture) & " the DOS records (" & DosCylinderSource & "):" & vbCrLf & "    " & String.Join(vbCrLf & "    ", lines) & vbCrLf & "    Ira's rule: a tail may be trimmed only if the DOS proves the size, the tail is a re-read or unformatted, it is an unbroken run to the last track, it carries no protection sign (never the Tandy track, never VTOS), and both sides qualify; a tail of 7 or more cylinders goes to review. trs80-tool's trim (Disk menu) does not check contents." & If(tailCount >= 7, " This tail is 7 or more cylinders: review.", ""))
    End Sub

    Private Shared Function SameData(a As DmkSector, b As DmkSector) As Boolean
        If a Is Nothing OrElse b Is Nothing OrElse Not a.HasData OrElse Not b.HasData Then Return False
        If a.Data.Length <> b.Data.Length Then Return False
        For i As Integer = 0 To a.Data.Length - 1
            If a.Data(i) <> b.Data(i) Then Return False
        Next
        Return True
    End Function

    Private Function LayoutMatchesDataTracks(t As DmkTrack) As Boolean
        Dim ref As DmkTrack = Image.Track(Math.Max(1, DosCylinders - 1), t.Side)
        If ref Is Nothing OrElse Not ref.Present Then Return False
        Dim a As String = StripBad(ref.Summary()), b As String = StripBad(t.Summary())
        Return String.Equals(a, b, StringComparison.Ordinal)
    End Function

    Private Shared Function StripBad(s As String) As String
        Dim i As Integer = s.IndexOf("  ", StringComparison.Ordinal)
        Return If(i > 0, s.Substring(0, i), s)
    End Function

    ' ------------------------------------------------------------------ report text

    Public Function ReportText() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("DMK analysis — " & IO.Path.GetFileName(Image.Path))
        sb.AppendLine("Made by TRSToolGui from the image's own bytes (not by trs80-tool).")
        sb.AppendLine()
        sb.AppendLine("== Summary")
        For Each kv In Facts
            sb.AppendLine("  " & If(kv.Key.Length > 0, kv.Key.PadRight(22) & " ", New String(" "c, 23)) & kv.Value)
        Next
        For Each section In New String() {"Image", "Geometry", "DOS identification", "Boot layout", "Protection", "Damage", "Tail tracks"}
            Dim items As List(Of DmkFinding) = Findings.Where(Function(f) f.Section = section).ToList()
            If items.Count = 0 Then Continue For
            sb.AppendLine()
            sb.AppendLine("== " & section)
            For Each f In items
                Dim tag As String
                Select Case f.Level
                    Case "protection" : tag = "[PROTECTION] "
                    Case "sign" : tag = "[SIGN] "
                    Case "damage" : tag = "[DAMAGE] "
                    Case "artifact" : tag = "[IMAGING] "
                    Case "warning" : tag = "[!] "
                    Case Else : tag = ""
                End Select
                sb.AppendLine("  " & tag & f.Text)
            Next
        Next
        sb.AppendLine()
        sb.AppendLine("== Tracks")
        For Each t In Image.Tracks
            sb.AppendLine("  cyl " & t.Cylinder.ToString(CultureInfo.InvariantCulture).PadLeft(2) & If(Image.Sides = 2, " s" & t.Side.ToString(CultureInfo.InvariantCulture), "") & "  " & t.Summary())
        Next
        Return sb.ToString()
    End Function
End Class
