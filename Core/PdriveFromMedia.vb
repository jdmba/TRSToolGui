' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  PdriveFromMedia.vb — works out the NEWDOS/80 v2 PDRIVE settings that read a
'  DMK image from the media itself, for the DMK analysis (DosIdentifier.vb).
'  PdriveCalculator.vb reads the PDRIVE table FORMAT left on the disk; this
'  file needs no table, so it also answers for disks that have none, and it
'  gives the settings for each machine and interface that can read the disk.
'
'  Order of the evidence:
'   1. Media: the main format is the most common (density, sectors, first ID)
'      on side 0 from track 1. Track 0 in the other density = flag K, data
'      area from cylinder 1. Cylinders = contiguous main-format tracks. Side 1
'      counts when most cylinders have sectors there. First ID 1 = flag I;
'      cylinder IDs one more than the track = flag J. 8-inch from a track
'      length of 2000H or more, or more sectors per side than 5.25" holds.
'   2. Directory: the boot sector starts 00 FE and byte 2 is DDSL. The
'      directory-marked sectors (F8H; F9H/FAH in single density) give the
'      GAT's linear position, and GPL = position / (5 x DDSL). Without marks
'      every GPL 1-8 is tried and the sector most like a GAT wins. GAT byte
'      CBH is 8nH (82H version 2, 80H version 1). DDGA = the marked run / 5,
'      else (HIT byte 1FH + 10) / 5.
'   3. TC: data cylinders (less track 0 with K), replaced by the GAT's figure
'      when its lockout table (GAT 60H, up to 96 lumps, FFH past the last)
'      disagrees. A lump exists when it holds one whole granule of 5 sectors:
'      lumps = ceil((TC x SPT \ 5) / GPL).
'   4. Cross-checks, never used for the answer: the boot code's model (ports
'      F0H-F4H = Model III, 37E0H-37EFH = Model I); boot bytes FDH-FFH = the
'      formatting drive's PDRIVE bytes +6, +2, +7; the PDRIVE table in sector
'      2 when its byte EFH = A5H.
'   5. Per machine, after the NEWDOS/80 v2.0 manual's PDRIVE page (2-35):
'      Model III TI=A for 5.25", TI=D (Apparat) for 8"; Model I TI=A for an
'      all-single-density 5.25" disk, C (Percom) and E (LNW) for double
'      density, B (Omikron, single density only) and E for 8". Flags I/J/K
'      from the media, H added for 8". TSR = 0 on the Model III, 3 on the
'      Model I: the drive's step rate, not the disk's. The manual allows no
'      TI=D on the Model I (although Model I SYS16 accepts interface code 2),
'      so it is not offered.
'  Not covered: TI flag M (3-sector granules), double step (L), NEWDOS86's
'  SPG=. The 8-inch rules follow the manual; no 8-inch image has been tried.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text

''' <summary>What the media says about a NEWDOS/80 disk's PDRIVE settings.</summary>
Public Class PdriveMediaResult
    Public Property IsNewdos As Boolean
    Public Property Failure As String = ""           ' why no answer ("" = there is one)

    ' media
    Public Property MediaCylinders As Integer
    Public Property Sides As Integer
    Public Property MainDD As Boolean
    Public Property SptSide As Integer
    Public Property FirstId As Integer
    Public Property FirstIdSide1 As Integer
    Public Property Track0DD As Boolean
    Public Property Track0Spt As Integer
    Public Property FlagK As Boolean
    Public Property FlagJ As Boolean
    Public Property EightInch As Boolean

    ' file system
    Public Property DDSL As Integer
    Public Property GPL As Integer
    Public Property DirSectors As Integer
    Public Property DDGA As Integer
    Public Property TC As Integer
    Public Property Lumps As Integer
    Public Property GatCB As Integer = -1

    ' formatter evidence
    Public Property FormatModel As String = "unknown"    ' "Model I", "Model III" or "unknown"
    Public Property BootFD As Integer = -1
    Public Property BootFE As Integer = -1
    Public Property BootFF As Integer = -1
    Public Property TableSlots As New List(Of Byte())    ' sector 2 records when byte EFH = A5H

    Public Property Notes As New List(Of String)
    Public Property Warnings As New List(Of String)

    Public ReadOnly Property SptTotal As Integer
        Get
            Return SptSide * Sides
        End Get
    End Property

    Public ReadOnly Property TdIndex As Integer
        Get
            Return If(EightInch, 1, 0) Or If(Sides = 2, 2, 0) Or If(MainDD, 4, 0)
        End Get
    End Property

    Public ReadOnly Property TdLetter As String
        Get
            Return ChrW(AscW("A"c) + TdIndex).ToString()
        End Get
    End Property

    ''' <summary>The special-condition letters the disk itself calls for, alphabetical as PDRIVE shows them.</summary>
    Public ReadOnly Property FlagLetters As String
        Get
            Dim sb As New StringBuilder
            If FirstId = 1 Then sb.Append("I"c)
            If FlagJ Then sb.Append("J"c)
            If FlagK Then sb.Append("K"c)
            Return sb.ToString()
        End Get
    End Property

    Public Function ParamString(ti As String, tsr As Integer) As String
        Dim ci As CultureInfo = CultureInfo.InvariantCulture
        Return "TI=" & ti & ",TD=" & TdLetter & ",TC=" & TC.ToString(ci) & ",SPT=" & SptTotal.ToString(ci) &
               ",TSR=" & tsr.ToString(ci) & ",GPL=" & GPL.ToString(ci) & ",DDSL=" & DDSL.ToString(ci) & ",DDGA=" & DDGA.ToString(ci)
    End Function
End Class

''' <summary>One machine's PDRIVE line for the disk.</summary>
Public Class PdriveMachineLine
    Public Property Machine As String = ""
    Public Property InterfaceName As String = ""
    Public Property Params As String = ""
End Class

Public Module PdriveFromMedia

    Private ReadOnly SptMaxByTd As Integer() = {10, 17, 20, 34, 18, 30, 36, 60}

    Public Function Analyze(img As DmkImage) As PdriveMediaResult
        Dim r As New PdriveMediaResult
        If img Is Nothing Then
            r.Failure = "no image"
            Return r
        End If
        If Not MeasureMedia(img, r) Then Return r
        If Not ReadFileSystem(img, r) Then Return r
        ReadFormatterEvidence(img, r)
        CrossCheck(r)
        Return r
    End Function

    ' ---------------------------------------------------------------- sector access

    ''' <summary>Sectors with a good ID CRC, one per sector number (the first good one wins).</summary>
    Private Function GoodSectors(t As DmkTrack) As List(Of DmkSector)
        Dim result As New List(Of DmkSector)
        If t Is Nothing OrElse Not t.Present Then Return result
        Dim seen As New HashSet(Of Integer)
        For Each s In t.Sectors
            If s.IdCrcOk AndAlso seen.Add(s.IdSector) Then result.Add(s)
        Next
        Return result
    End Function

    ''' <summary>A sector by number with data: a good ID first, then any.</summary>
    Private Function FindSector(t As DmkTrack, id As Integer) As DmkSector
        If t Is Nothing Then Return Nothing
        For Each s In t.Sectors
            If s.IdCrcOk AndAlso s.IdSector = id AndAlso s.HasData Then Return s
        Next
        For Each s In t.Sectors
            If s.IdSector = id AndAlso s.HasData Then Return s
        Next
        Return Nothing
    End Function

    ' ---------------------------------------------------------------- media

    Private Structure TrackShape
        Public Present As Boolean
        Public DD As Boolean
        Public Spt As Integer
        Public MinId As Integer
        Public AllSize256 As Boolean
        Public CylIdsPlusOne As Boolean
        Public CylIdsMatch As Boolean
    End Structure

    Private Function Shape(t As DmkTrack) As TrackShape
        Dim s As New TrackShape
        Dim good = GoodSectors(t)
        If good.Count = 0 Then Return s
        s.Present = True
        s.DD = good.Where(Function(x) x.Mfm).Count() * 2 > good.Count
        s.Spt = good.Count
        s.MinId = good.Min(Function(x) x.IdSector)
        s.AllSize256 = good.All(Function(x) x.SizeCode = 1)
        s.CylIdsMatch = good.All(Function(x) x.IdCylinder = t.Cylinder)
        s.CylIdsPlusOne = good.All(Function(x) x.IdCylinder = t.Cylinder + 1)
        Return s
    End Function

    Private Function MeasureMedia(img As DmkImage, r As PdriveMediaResult) As Boolean
        Dim maxPhys As Integer = -1
        For Each t In img.Tracks
            If t.Present AndAlso t.Cylinder > maxPhys Then maxPhys = t.Cylinder
        Next
        If maxPhys < 0 Then
            r.Failure = "the image holds no tracks"
            Return False
        End If

        Dim side0(maxPhys) As TrackShape
        Dim side1(maxPhys) As TrackShape
        For t As Integer = 0 To maxPhys
            side0(t) = Shape(img.Track(t, 0))
            side1(t) = Shape(img.Track(t, 1))
        Next
        If Not side0(0).Present Then
            r.Failure = "track 0 side 0 has no readable sectors"
            Return False
        End If

        ' Main format: the most common (density, sectors, first ID) over side 0, tracks 1 up.
        Dim tally As New Dictionary(Of (Boolean, Integer, Integer), Integer)
        For t As Integer = If(maxPhys >= 1, 1, 0) To maxPhys
            If Not side0(t).Present Then Continue For
            Dim key = (side0(t).DD, side0(t).Spt, side0(t).MinId)
            Dim n As Integer = 0
            tally.TryGetValue(key, n)
            tally(key) = n + 1
        Next
        If tally.Count = 0 Then tally((side0(0).DD, side0(0).Spt, side0(0).MinId)) = 1
        Dim main = tally.OrderByDescending(Function(kv) kv.Value).First().Key
        r.MainDD = main.Item1
        r.SptSide = main.Item2
        r.FirstId = main.Item3

        r.Track0DD = side0(0).DD
        r.Track0Spt = side0(0).Spt
        r.FlagK = (maxPhys >= 1) AndAlso (r.Track0DD <> r.MainDD)
        If Not r.FlagK AndAlso r.Track0Spt <> r.SptSide Then
            r.Warnings.Add("Track 0 has " & r.Track0Spt.ToString(CultureInfo.InvariantCulture) & " sectors, the other tracks " & r.SptSide.ToString(CultureInfo.InvariantCulture) & ".")
        End If

        ' Cylinders: tracks from the first data track that have the main format, contiguously.
        Dim first As Integer = If(r.FlagK, 1, 0)
        Dim last As Integer = first - 1
        For t As Integer = first To maxPhys
            Dim s = side0(t)
            If s.Present AndAlso s.DD = r.MainDD AndAlso s.Spt = r.SptSide AndAlso s.MinId = r.FirstId Then
                last = t
            Else
                Exit For
            End If
        Next
        r.MediaCylinders = last + 1
        Dim odd As New List(Of String)
        For t As Integer = last + 1 To maxPhys
            If side0(t).Present Then odd.Add(t.ToString(CultureInfo.InvariantCulture))
        Next
        If odd.Count > 0 Then
            r.Warnings.Add("Tracks past cylinder " & last.ToString(CultureInfo.InvariantCulture) & " hold sectors in another format (tracks " &
                           String.Join(",", odd.Take(12)) & If(odd.Count > 12, ",...", "") & "); leftovers or damage, not counted.")
        End If
        If last < first Then
            r.Failure = "no track after track 0 has a consistent format"
            Return False
        End If
        Dim used As Integer = last - first + 1

        ' Sides: side 1 counts when most used cylinders have sectors there.
        Dim s1 As Integer = 0
        For t As Integer = first To last
            If side1(t).Present Then s1 += 1
        Next
        r.Sides = If(s1 * 2 > used, 2, 1)
        If r.Sides = 1 AndAlso s1 > 0 Then r.Warnings.Add("Side 1 has sectors on only " & s1.ToString(CultureInfo.InvariantCulture) & " cylinders; treated as single sided.")
        If r.Sides = 2 Then
            Dim ids = Enumerable.Range(first, used).Where(Function(t) side1(t).Present).Select(Function(t) side1(t).MinId)
            r.FirstIdSide1 = ids.GroupBy(Function(v) v).OrderByDescending(Function(g) g.Count()).First().Key
            Dim bad = Enumerable.Range(first, used).Where(Function(t) side1(t).Present AndAlso (side1(t).DD <> r.MainDD OrElse side1(t).Spt <> r.SptSide)).ToList()
            If bad.Count > 0 Then r.Warnings.Add("Side 1 differs from side 0 on " & bad.Count.ToString(CultureInfo.InvariantCulture) & " cylinder(s): " & String.Join(",", bad.Take(12)) & ".")
            If r.FlagK AndAlso side1(0).Present AndAlso side1(0).DD <> r.Track0DD Then r.Warnings.Add("Track 0 is in different densities on its two sides.")
        End If

        ' Sector size and cylinder numbering.
        Dim not256 = Enumerable.Range(first, used).Where(Function(t) Not side0(t).AllSize256).ToList()
        If not256.Count > 0 Then r.Warnings.Add(not256.Count.ToString(CultureInfo.InvariantCulture) & " track(s) have sectors that are not 256 bytes (first: track " & not256(0).ToString(CultureInfo.InvariantCulture) & "). NEWDOS/80 uses 256-byte sectors only.")
        Dim plusOne As Integer = Enumerable.Range(first, used).Where(Function(t) side0(t).CylIdsPlusOne).Count()
        If plusOne * 2 > used AndAlso Not r.FlagK Then r.FlagJ = True
        Dim mismatch As Integer = Enumerable.Range(first, used).Where(Function(t) Not side0(t).CylIdsMatch AndAlso Not side0(t).CylIdsPlusOne).Count()
        If mismatch > 0 Then r.Warnings.Add(mismatch.ToString(CultureInfo.InvariantCulture) & " track(s) have cylinder numbers in their sector IDs that don't match the track (double-stepped image, or damage).")

        ' 8-inch: track length or sectors per side past what a 5.25-inch drive holds.
        Dim sdOnly As Boolean = (img.Options And &H40) <> 0
        r.EightInch = (img.TrackLength >= &H2000 AndAlso Not sdOnly) OrElse
                      (r.MainDD AndAlso r.SptSide > 18) OrElse (Not r.MainDD AndAlso r.SptSide > 10)
        Return True
    End Function

    ''' <summary>Linear sector number of the data area to the physical sector.</summary>
    Private Function Locate(img As DmkImage, r As PdriveMediaResult, lin As Integer) As DmkSector
        If lin < 0 OrElse r.SptTotal = 0 Then Return Nothing
        Dim cyl As Integer = If(r.FlagK, 1, 0) + lin \ r.SptTotal
        Dim rest As Integer = lin Mod r.SptTotal
        Dim side As Integer = rest \ r.SptSide
        Dim id As Integer = If(side = 0, r.FirstId, r.FirstIdSide1) + rest Mod r.SptSide
        Return FindSector(img.Track(cyl, side), id)
    End Function

    Private Function LinearOf(r As PdriveMediaResult, cyl As Integer, side As Integer, id As Integer) As Integer
        Dim baseId As Integer = If(side = 0, r.FirstId, r.FirstIdSide1)
        Return (cyl - If(r.FlagK, 1, 0)) * r.SptTotal + side * r.SptSide + (id - baseId)
    End Function

    ''' <summary>F8H in double density; F8H-FAH in single density (the 1771 records four marks; Model I SD writes FAH).</summary>
    Private Function IsDirMark(s As DmkSector) As Boolean
        Return s.DataMark = &HF8 OrElse s.DataMark = &HF9 OrElse s.DataMark = &HFA
    End Function

    ' ---------------------------------------------------------------- file system

    Private Function BootSector(img As DmkImage) As DmkSector
        Dim t0 = img.Track(0, 0)
        Dim good = GoodSectors(t0)
        If good.Count = 0 Then Return Nothing
        Dim minId As Integer = good.Min(Function(x) x.IdSector)
        Dim b = FindSector(t0, minId)
        If IsBootStart(b) Then Return b
        Dim b1 = FindSector(t0, minId + 1)
        If IsBootStart(b1) Then Return b1
        Return b
    End Function

    Private Function IsBootStart(s As DmkSector) As Boolean
        Return s IsNot Nothing AndAlso s.Data IsNot Nothing AndAlso s.Data.Length >= 256 AndAlso s.Data(0) = 0 AndAlso s.Data(1) = &HFE
    End Function

    Private Function IsPrintable(b As Byte) As Boolean
        Return b >= &H20 AndAlso b < &H7F
    End Function

    Private Function GatScore(g As DmkSector) As Integer
        If g Is Nothing OrElse g.Data Is Nothing OrElse g.Data.Length < 256 Then Return -1
        Dim d = g.Data
        Dim score As Integer = 0
        If IsDirMark(g) Then score += 3
        If d(&HCB) >= &H80 AndAlso d(&HCB) <= &H8F Then score += 3
        If Enumerable.Range(&HD0, 8).All(Function(i) IsPrintable(d(i))) Then score += 2
        If d(&HE0) = &HD OrElse Enumerable.Range(&HE0, 4).All(Function(i) IsPrintable(d(i))) Then score += 1
        Return score
    End Function

    Private Function ReadFileSystem(img As DmkImage, r As PdriveMediaResult) As Boolean
        Dim ci As CultureInfo = CultureInfo.InvariantCulture
        Dim boot = BootSector(img)
        If boot Is Nothing OrElse boot.Data Is Nothing OrElse boot.Data.Length < 256 Then
            r.Failure = "the boot sector (track 0) is unreadable"
            Return False
        End If
        If boot.Data(0) <> 0 OrElse boot.Data(1) <> &HFE Then
            r.Failure = "the boot sector starts " & boot.Data(0).ToString("X2", ci) & " " & boot.Data(1).ToString("X2", ci) & ", not 00 FE"
            Return False
        End If
        r.DDSL = boot.Data(2)
        r.BootFD = boot.Data(&HFD) : r.BootFE = boot.Data(&HFE) : r.BootFF = boot.Data(&HFF)

        ' Where is the directory? Every directory-marked sector of the data area, in order.
        Dim marks As New List(Of Integer)
        For Each t In img.Tracks
            If t.Cylinder < If(r.FlagK, 1, 0) OrElse t.Cylinder >= r.MediaCylinders OrElse t.Side >= r.Sides Then Continue For
            For Each s In GoodSectors(t)
                If IsDirMark(s) Then marks.Add(LinearOf(r, t.Cylinder, t.Side, s.IdSector))
            Next
        Next
        marks.Sort()

        Dim dirLin As Integer = -1, runLen As Integer = 0
        If marks.Count > 0 Then
            dirLin = marks(0)
            runLen = 1
            While runLen < marks.Count AndAlso marks(runLen) = dirLin + runLen
                runLen += 1
            End While
            If runLen < marks.Count Then r.Warnings.Add((marks.Count - runLen).ToString(ci) & " directory-marked sector(s) lie outside the directory run.")
        End If

        ' GPL: the directory starts at sector 5 x GPL x DDSL of the data area.
        r.GPL = 0
        If dirLin >= 0 AndAlso r.DDSL > 0 AndAlso dirLin Mod (5 * r.DDSL) = 0 Then
            Dim g As Integer = dirLin \ (5 * r.DDSL)
            If g >= 1 AndAlso g <= 8 Then r.GPL = g
        End If
        If r.GPL = 0 Then
            ' No usable marks: try every GPL and keep the one whose sector looks most like a GAT.
            Dim best As Integer = -1, bestG As Integer = 0
            For g As Integer = 1 To 8
                Dim sc = GatScore(Locate(img, r, 5 * g * r.DDSL))
                If sc > best Then best = sc : bestG = g
            Next
            If best >= 4 Then
                r.GPL = bestG
                If dirLin >= 0 Then
                    r.Warnings.Add("The directory marks start at data sector " & dirLin.ToString(ci) & ", which is not a lump boundary for DDSL=" & r.DDSL.ToString(ci) & "; GPL=" & bestG.ToString(ci) & " chosen from the GAT's contents.")
                Else
                    r.Warnings.Add("No directory-marked sectors found; GPL=" & bestG.ToString(ci) & " chosen from the GAT's contents.")
                End If
                dirLin = 5 * bestG * r.DDSL
                runLen = 0
            Else
                r.Failure = "no NEWDOS/80 directory found at lump " & r.DDSL.ToString(ci) & " (boot byte 2) for any GPL"
                Return False
            End If
        End If

        Dim gat = Locate(img, r, dirLin)
        Dim hit = Locate(img, r, dirLin + 1)
        If gat Is Nothing OrElse gat.Data Is Nothing OrElse gat.Data.Length < 256 Then
            r.Failure = "the GAT sector is unreadable"
            Return False
        End If
        If Not gat.DataCrcOk Then r.Warnings.Add("The GAT sector has a data CRC error.")
        r.GatCB = gat.Data(&HCB)
        If r.GatCB < &H80 OrElse r.GatCB > &H8F Then
            r.Failure = "GAT byte CBH is " & r.GatCB.ToString("X2", ci) & "H; NEWDOS/80 writes 8nH (82H version 2, 80H version 1)"
            Return False
        End If
        r.IsNewdos = True
        If r.GatCB = &H80 Then r.Notes.Add("GAT byte CBH = 80H: a NEWDOS/80 version 1 disk (version 2 can't read version 1 double-sided double-density disks).")

        ' Directory length: the marked run, else HIT byte 1FH (extra sectors beyond 10).
        Dim hitTotal As Integer = -1
        If hit IsNot Nothing AndAlso hit.Data IsNot Nothing AndAlso hit.Data.Length > &H1F Then
            Dim extra As Integer = hit.Data(&H1F)
            If extra Mod 5 = 0 AndAlso extra <= 20 Then hitTotal = 10 + extra
        End If
        If runLen >= 10 Then
            r.DirSectors = runLen
            If hitTotal > 0 AndAlso hitTotal <> runLen Then r.Warnings.Add("HIT byte 1FH says " & hitTotal.ToString(ci) & " directory sectors, " & runLen.ToString(ci) & " carry the directory mark.")
        ElseIf hitTotal > 0 Then
            r.DirSectors = hitTotal
            If runLen > 0 Then r.Warnings.Add("Only " & runLen.ToString(ci) & " directory-marked sectors; length taken from HIT byte 1FH (" & hitTotal.ToString(ci) & ").")
        Else
            r.DirSectors = 10
            r.Warnings.Add("Directory length unknown; assumed 10 sectors.")
        End If
        If r.DirSectors Mod 5 <> 0 Then r.Warnings.Add("Directory length " & r.DirSectors.ToString(ci) & " is not a whole number of granules.")
        r.DDGA = Math.Max(2, Math.Min(6, (r.DirSectors + 4) \ 5))

        ' TC: media cylinders, less track 0 when it is in the other density (flag K).
        Dim tcMedia As Integer = r.MediaCylinders - If(r.FlagK, 1, 0)
        r.TC = tcMedia
        Dim lumpsMedia As Integer = LumpCount(tcMedia, r.SptTotal, r.GPL)
        ' The lockout table at GAT 60H holds one byte per lump while there are at most 96
        ' lumps; FFH past the last lump.
        If lumpsMedia <= 96 Then
            Dim n As Integer = 0
            While n < 96 AndAlso gat.Data(&H60 + n) <> &HFF
                n += 1
            End While
            If (n > 0 AndAlso n < 96) OrElse (n = 96 AndAlso lumpsMedia = 96) Then
                Dim gatLumps As Integer = n
                Dim spt As Integer = r.SptTotal, gpl As Integer = r.GPL
                Dim cands = Enumerable.Range(1, 255).Where(Function(tc) LumpCount(tc, spt, gpl) = gatLumps).ToList()
                If cands.Count > 0 Then
                    Dim tcGat As Integer = cands.OrderBy(Function(tc) Math.Abs(tc - tcMedia)).First()
                    If tcGat < tcMedia Then
                        r.Warnings.Add("The GAT covers " & gatLumps.ToString(ci) & " lumps, i.e. TC=" & tcGat.ToString(ci) & "; the image has " & tcMedia.ToString(ci) & " data cylinders. Using the GAT's TC (the DOS never allocates past it).")
                        r.TC = tcGat
                    ElseIf tcGat > tcMedia Then
                        ' Are the lumps past the image's end free in the allocation map?
                        Dim freeMask As Integer = (&HFF << r.GPL) And &HFF
                        Dim fromLump As Integer = lumpsMedia
                        Dim allFree As Boolean = Enumerable.Range(fromLump, Math.Max(0, gatLumps - fromLump)).All(Function(l) l < &H60 AndAlso gat.Data(l) = freeMask)
                        Dim which As String = If(gatLumps - fromLump = 1, "lump " & fromLump.ToString(ci) & " is", "lumps " & fromLump.ToString(ci) & "-" & (gatLumps - 1).ToString(ci) & " are")
                        r.Warnings.Add("The GAT covers " & gatLumps.ToString(ci) & " lumps, i.e. TC=" & tcGat.ToString(ci) & ", but the image has only " & tcMedia.ToString(ci) &
                                       " data cylinders: the image is missing tracks" & If(allFree, " (" & which & " free in the GAT, so no file data is lost)", " that the GAT shows in use") & ". Using the GAT's TC.")
                        r.TC = tcGat
                    End If
                Else
                    r.Warnings.Add("The GAT's lockout table covers " & gatLumps.ToString(ci) & " lumps, which no TC gives with SPT=" & r.SptTotal.ToString(ci) & ", GPL=" & r.GPL.ToString(ci) & ".")
                End If
            End If
        End If
        r.Lumps = LumpCount(r.TC, r.SptTotal, r.GPL)
        If r.Lumps > 192 Then r.Warnings.Add(r.Lumps.ToString(ci) & " lumps is more than NEWDOS/80's 192.")
        Return True
    End Function

    ''' <summary>Lumps that hold at least one whole granule of 5 sectors.</summary>
    Public Function LumpCount(tc As Integer, spt As Integer, gpl As Integer) As Integer
        If gpl <= 0 Then Return 0
        Dim granules As Integer = (tc * spt) \ 5
        Return (granules + gpl - 1) \ gpl
    End Function

    ' ---------------------------------------------------------------- formatter evidence

    Private Function Contains(d As Byte(), ParamArray pat As Byte()) As Boolean
        For i As Integer = 0 To d.Length - pat.Length
            Dim ok As Boolean = True
            For j As Integer = 0 To pat.Length - 1
                If d(i + j) <> pat(j) Then ok = False : Exit For
            Next
            If ok Then Return True
        Next
        Return False
    End Function

    Private Sub ReadFormatterEvidence(img As DmkImage, r As PdriveMediaResult)
        Dim boot = BootSector(img)
        If boot IsNot Nothing AndAlso boot.Data IsNot Nothing Then
            Dim d = boot.Data
            ' Model III boot code drives the controller through ports F0H-F4H;
            ' Model I boot code through memory 37E1H/37ECH-37EFH.
            Dim m3 As Boolean = Contains(d, &HD3, &HF4) OrElse Contains(d, &HDB, &HF0)
            Dim m1 As Boolean = Contains(d, &HEC, &H37) OrElse Contains(d, &HE1, &H37) OrElse Contains(d, &HEE, &H37)
            If m3 AndAlso Not m1 Then r.FormatModel = "Model III"
            If m1 AndAlso Not m3 Then r.FormatModel = "Model I"
        End If

        Dim t0 = img.Track(0, 0)
        Dim good = GoodSectors(t0)
        If good.Count > 0 Then
            Dim minId As Integer = good.Min(Function(x) x.IdSector)
            Dim s2 = FindSector(t0, minId + 2)
            If s2 IsNot Nothing AndAlso s2.Data IsNot Nothing AndAlso s2.Data.Length >= 256 AndAlso s2.Data(&HEF) = &HA5 Then
                For i As Integer = 0 To 9
                    Dim rec(15) As Byte
                    Array.Copy(s2.Data, i * 16, rec, 0, 16)
                    r.TableSlots.Add(rec)
                Next
            End If
        End If
    End Sub

    ''' <summary>The interface letter of a PDRIVE interface code (byte +2 bits 2-4).</summary>
    Public Function InterfaceLetter(code As Integer, model As String) As String
        If model = "Model III" Then
            Select Case code
                Case 1 : Return "A"
                Case 2 : Return "D"
            End Select
        Else
            Select Case code
                Case 1 : Return "A"
                Case 2 : Return "D"
                Case 3 : Return "B"
                Case 4 : Return "C"
                Case 5 : Return "E"
            End Select
        End If
        Return "?"
    End Function

    ''' <summary>The TI letters stored in a table record (+0DH A-H, +0EH I-P).</summary>
    Public Function TableTi(rec As Byte()) As String
        Dim sb As New StringBuilder
        For i As Integer = 0 To 7
            If ((rec(13) >> i) And 1) <> 0 Then sb.Append(ChrW(AscW("A"c) + i))
        Next
        For i As Integer = 0 To 7
            If ((rec(14) >> i) And 1) <> 0 Then sb.Append(ChrW(AscW("I"c) + i))
        Next
        Return sb.ToString()
    End Function

    Public Function DescribeTableRecord(rec As Byte()) As String
        If ((rec(2) >> 2) And 7) = 0 Then Return "*** ERROR *** (record failed PDRIVE's checks)"
        Dim ci As CultureInfo = CultureInfo.InvariantCulture
        Dim td As String = If(rec(15) <= 7, ChrW(AscW("A"c) + rec(15)).ToString(), "?")
        Return "TI=" & TableTi(rec) & ",TD=" & td & ",TC=" & rec(3).ToString(ci) & ",SPT=" & rec(4).ToString(ci) & ",TSR=" & rec(12).ToString(ci) &
               ",GPL=" & rec(5).ToString(ci) & ",DDSL=" & rec(8).ToString(ci) & ",DDGA=" & rec(9).ToString(ci)
    End Function

    Private Sub CrossCheck(r As PdriveMediaResult)
        Dim ci As CultureInfo = CultureInfo.InvariantCulture
        ' Boot sector bytes FDH-FFH = the formatting drive's PDRIVE +6, +2, +7.
        If r.BootFE >= 0 Then
            Dim code As Integer = (r.BootFE >> 2) And 7
            Dim flags7 As Integer = r.BootFF
            Dim bootDD As Boolean = (flags7 And 1) <> 0
            Dim bootDS As Boolean = (flags7 And &H40) <> 0
            Dim bootK As Boolean = (r.BootFE And &H40) <> 0
            Dim bootM As Boolean = (r.BootFE And &H20) <> 0
            Dim bootTsr As Integer = r.BootFE And 3
            Dim ti As String = InterfaceLetter(code, r.FormatModel) & If(bootK, "K", "") & If(bootM, "M", "")
            r.Notes.Add("Boot sector FDH-FFH = " & r.BootFD.ToString("X2", ci) & " " & r.BootFE.ToString("X2", ci) & " " & r.BootFF.ToString("X2", ci) &
                        ": formatted on a drive set to TI=" & ti & ", TSR=" & bootTsr.ToString(ci) & ", " & If(bootDD, "double", "single") & " density, " &
                        If(bootDS, "two sides", "one side") & If((flags7 And &H80) <> 0, ", 8-inch", "") & ".")
            Dim diffs As New List(Of String)
            If bootDD <> r.MainDD Then diffs.Add("density")
            If bootDS <> (r.Sides = 2) Then diffs.Add("sides")
            If bootK <> r.FlagK Then diffs.Add("flag K")
            If diffs.Count > 0 Then r.Warnings.Add("The boot sector's formatting-drive settings disagree with the media on " & String.Join(", ", diffs) & ".")
            If r.FormatModel = "Model I" AndAlso code = 1 AndAlso bootDD Then r.Warnings.Add("The boot sector claims TI=A with double density, which Model I PDRIVE refuses.")
        End If

        If r.TableSlots.Count > 0 Then
            Dim matchSlots As New List(Of String)
            For i As Integer = 0 To r.TableSlots.Count - 1
                Dim rec = r.TableSlots(i)
                If ((rec(2) >> 2) And 7) = 0 Then Continue For
                If rec(3) = r.TC AndAlso rec(4) = r.SptTotal AndAlso rec(5) = r.GPL AndAlso rec(8) = r.DDSL AndAlso rec(9) = r.DDGA AndAlso rec(15) = r.TdIndex Then matchSlots.Add(i.ToString(ci))
            Next
            r.Notes.Add("PDRIVE table in sector 2 (the formatting system's drives), slot 0: " & DescribeTableRecord(r.TableSlots(0)) & ".")
            If matchSlots.Count > 0 Then
                r.Notes.Add("The table's slot(s) " & String.Join(",", matchSlots) & " match the settings from the media (TD, TC, SPT, GPL, DDSL, DDGA).")
            Else
                r.Notes.Add("No slot in the table matches the settings from the media; the table describes the formatting system's drives, not necessarily this disk.")
            End If
            Dim rec0 = r.TableSlots(0)
            If r.FormatModel = "unknown" AndAlso ((rec0(2) >> 2) And 7) = 1 AndAlso (rec0(7) And 1) <> 0 Then
                r.Notes.Add("Table slot 0 is TI=A with double density: only a Model III table can hold that.")
            End If
        End If

        If r.FormatModel <> "unknown" Then r.Notes.Insert(0, "Boot code is " & r.FormatModel & " NEWDOS/80 (" & If(r.FormatModel = "Model III", "controller on ports F0H-F4H", "controller at 37E0H-37EFH") & ").")
        If r.FormatModel = "Model III" AndAlso r.Track0DD Then r.Notes.Add("Track 0 is double density: a Model I needs a doubler to read this disk at all.")

        If r.SptTotal > SptMaxByTd(r.TdIndex) Then r.Warnings.Add("SPT=" & r.SptTotal.ToString(ci) & " is more than PDRIVE allows for TD=" & r.TdLetter & " (" & SptMaxByTd(r.TdIndex).ToString(ci) & ").")
        If r.GPL < 2 Then r.Warnings.Add("GPL=1 is below the manual's range (2-8), though PDRIVE accepts it.")
    End Sub

    ' ---------------------------------------------------------------- settings per machine

    Private Function SortLetters(s As String) As String
        Dim a = s.ToCharArray()
        Array.Sort(a)
        Return New String(a)
    End Function

    ''' <summary>
    ''' The PDRIVE line for each machine and interface that can read the disk. Interfaces and the
    ''' drive types each supports, from the NEWDOS/80 v2.0 manual's PDRIVE TI= page (2-35):
    '''   Model I   A standard          TD A, C        (5.25-inch single density)
    '''   Model I   B Omikron mapper    TD A, B, C, D  (single density, 5.25- and 8-inch)
    '''   Model I   C Percom doubler    TD A, C, E, G  (5.25-inch, both densities)
    '''   Model I   E LNW               TD A-H
    '''   Model III A standard          TD A, C, E, G  (5.25-inch, both densities)
    '''   Model III D Apparat           TD A-H
    ''' 8-inch also wants flag H (head settle delay).
    ''' </summary>
    Public Function MachineLines(r As PdriveMediaResult) As List(Of PdriveMachineLine)
        Dim lines As New List(Of PdriveMachineLine)
        If r Is Nothing OrElse Not r.IsNewdos OrElse r.Failure.Length > 0 Then Return lines
        Dim anyDD As Boolean = r.MainDD OrElse r.Track0DD
        Dim fl As String = SortLetters(r.FlagLetters & If(r.EightInch, "H", ""))

        If Not r.EightInch Then
            lines.Add(New PdriveMachineLine With {.Machine = "Model III", .InterfaceName = "standard", .Params = r.ParamString("A" & fl, 0)})
        Else
            lines.Add(New PdriveMachineLine With {.Machine = "Model III", .InterfaceName = "Apparat controller", .Params = r.ParamString(SortLetters("D" & fl), 0)})
        End If

        If Not r.EightInch AndAlso Not anyDD Then
            lines.Add(New PdriveMachineLine With {.Machine = "Model I", .InterfaceName = "standard", .Params = r.ParamString("A" & fl, 3)})
        ElseIf Not r.EightInch Then
            lines.Add(New PdriveMachineLine With {.Machine = "Model I", .InterfaceName = "Percom doubler", .Params = r.ParamString("C" & fl, 3)})
            lines.Add(New PdriveMachineLine With {.Machine = "Model I", .InterfaceName = "LNW", .Params = r.ParamString("E" & fl, 3)})
        Else
            If Not anyDD Then lines.Add(New PdriveMachineLine With {.Machine = "Model I", .InterfaceName = "Omikron mapper", .Params = r.ParamString(SortLetters("B" & fl), 3)})
            lines.Add(New PdriveMachineLine With {.Machine = "Model I", .InterfaceName = "LNW", .Params = r.ParamString(SortLetters("E" & fl), 3)})
        End If
        Return lines
    End Function
End Module
