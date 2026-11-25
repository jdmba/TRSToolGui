' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DmkHexDump.vb — a hexadecimal listing of chosen tracks of a DMK image, as
'  plain text to read, to copy, or to hand to a tool that will not take a .dmk.
'
'  trs80-tool's own "hexdump" takes a whole host file and knows nothing about
'  tracks, so this is done here instead, on the DmkImage that DmkAnalyzer.vb
'  has already parsed. Nothing in this file reads the disk again.
'
'  What it prints, per track:
'    * the track's own line: cylinder, side, how many sectors, the density, and
'      whether FM bytes are stored twice in the file (DMK's own convention —
'      header option bits 6 and 7 turn it off, which DmkImage.SingleByteFm has);
'    * one line per IDAM slot: where the pointer lands, the C/H/R/N it names,
'      both CRCs and the data address mark;
'    * the data bytes of each sector, undoubled, which is the disk's real
'      content and the part worth reading or editing;
'    * the raw track bytes exactly as the file holds them, with the 128-byte
'      IDAM table at offset 0, when the spec asks for "raw".
'
'  The spec is what the user types: a list of cylinders ("0", "0-1", "0,2,5-7"),
'  optionally "all", optionally "raw" for the raw track bytes as well, and
'  optionally "side0" or "side1" to take one side only. Anything it cannot read
'  is reported in the text rather than thrown, because it is driven straight
'  from a text box.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>A hexadecimal listing of chosen tracks of a DMK image.</summary>
Public NotInheritable Class DmkHexDump

    Private Sub New()
    End Sub

    ''' <summary>What the track box starts with: the two tracks that carry the boot and the system loader.</summary>
    Public Const DefaultSpec As String = "0-1"

    ''' <summary>Bytes per line of hex, which is what makes the ASCII column line up.</summary>
    Private Const PerLine As Integer = 16

    ''' <summary>What a spec asked for.</summary>
    Private NotInheritable Class Request
        Public ReadOnly Cylinders As New List(Of Integer)
        Public All As Boolean = False
        Public Raw As Boolean = False
        Public OnlySide As Integer = -1          ' -1 = both
        Public ReadOnly Unreadable As New List(Of String)
    End Class

    ''' <summary>
    ''' Read the image at <paramref name="path"/> and list the tracks <paramref name="spec"/> names.
    ''' Never throws for a bad spec; a file that cannot be read is reported as text.
    ''' </summary>
    Public Shared Function Dump(path As String, spec As String) As String
        Dim img As DmkImage
        Try
            img = DmkImage.Load(path)
        Catch ex As Exception
            Return "Could not read " & path & ":" & Environment.NewLine & ex.Message
        End Try
        Return Dump(img, spec)
    End Function

    ''' <summary>List the tracks <paramref name="spec"/> names out of an image already parsed.</summary>
    Public Shared Function Dump(img As DmkImage, spec As String) As String
        Dim sb As New StringBuilder()
        If img Is Nothing Then Return "No image."

        Dim req As Request = ParseSpec(spec)

        sb.AppendLine(Path.GetFileName(img.Path))
        sb.AppendLine(New String("="c, Math.Max(8, Path.GetFileName(img.Path).Length)))
        sb.AppendLine()

        If Not String.Equals(img.Kind, "DMK", StringComparison.OrdinalIgnoreCase) Then
            sb.AppendLine("This is a " & img.Kind & " image, not a DMK one" & If(img.KindNote.Length > 0, " (" & img.KindNote & ")", "") & ".")
            sb.AppendLine("Tracks can only be listed for DMK images, because only those keep the track as the")
            sb.AppendLine("controller saw it. Save the image as DMK first, then come back.")
            Return sb.ToString()
        End If

        sb.AppendLine("Header     " & img.HeaderTracks.ToString(CultureInfo.InvariantCulture) & " tracks, track length " &
                      img.TrackLength.ToString("X4", CultureInfo.InvariantCulture) & "h (" &
                      img.TrackLength.ToString(CultureInfo.InvariantCulture) & " bytes), " &
                      If(img.SingleSided, "single-sided", "two-sided") & ", options " &
                      img.Options.ToString("X2", CultureInfo.InvariantCulture) & "h")
        sb.AppendLine("           write-protect byte " & img.WriteProtect.ToString("X2", CultureInfo.InvariantCulture) & "h, " &
                      img.Cylinders.ToString(CultureInfo.InvariantCulture) & " cylinder(s) in the file, " &
                      img.Sides.ToString(CultureInfo.InvariantCulture) & " side(s)")
        sb.AppendLine("           FM (single density) bytes are stored " & If(img.SingleByteFm, "once", "twice") & " in the file")
        sb.AppendLine("           every offset below counts from the start of the track, so the 128-byte IDAM")
        sb.AppendLine("           table is 0000h-007Fh and the first written byte follows it")
        sb.AppendLine()

        For Each bad In req.Unreadable
            sb.AppendLine("Ignored in the track list: " & bad)
        Next
        If req.Unreadable.Count > 0 Then sb.AppendLine()

        Dim wanted As List(Of Integer) = WantedCylinders(img, req)
        If wanted.Count = 0 Then
            sb.AppendLine("No track in this image matches """ & If(spec, "") & """.")
            sb.AppendLine("Give a cylinder, a range or a list: 0   0-1   0,2,5-7   all")
            sb.AppendLine("Add ""raw"" for the raw track bytes, and ""side0"" or ""side1"" for one side only.")
            Return sb.ToString()
        End If

        For Each cyl In wanted
            For side As Integer = 0 To img.Sides - 1
                If req.OnlySide >= 0 AndAlso side <> req.OnlySide Then Continue For
                AppendTrack(sb, img, cyl, side, req.Raw)
            Next
        Next
        Return sb.ToString()
    End Function

    ' ------------------------------------------------------------------ the spec

    Private Shared Function ParseSpec(spec As String) As Request
        Dim req As New Request()
        If String.IsNullOrWhiteSpace(spec) Then spec = DefaultSpec
        Dim words As String() = spec.Replace(","c, " "c).Split(New Char() {" "c, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries)
        For Each w In words
            Dim word As String = w.Trim().ToLowerInvariant()
            Select Case word
                Case "all", "*"
                    req.All = True
                Case "raw"
                    req.Raw = True
                Case "side0", "side-0", "s0"
                    req.OnlySide = 0
                Case "side1", "side-1", "s1"
                    req.OnlySide = 1
                Case Else
                    Dim dash As Integer = word.IndexOf("-"c, 1)
                    If dash > 0 Then
                        Dim lo, hi As Integer
                        If Integer.TryParse(word.Substring(0, dash), NumberStyles.Integer, CultureInfo.InvariantCulture, lo) AndAlso
                           Integer.TryParse(word.Substring(dash + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, hi) Then
                            If lo > hi Then
                                Dim t As Integer = lo : lo = hi : hi = t
                            End If
                            For c As Integer = lo To hi
                                If c >= 0 AndAlso Not req.Cylinders.Contains(c) Then req.Cylinders.Add(c)
                            Next
                        Else
                            req.Unreadable.Add(w)
                        End If
                    Else
                        Dim one As Integer
                        If Integer.TryParse(word, NumberStyles.Integer, CultureInfo.InvariantCulture, one) AndAlso one >= 0 Then
                            If Not req.Cylinders.Contains(one) Then req.Cylinders.Add(one)
                        Else
                            req.Unreadable.Add(w)
                        End If
                    End If
            End Select
        Next
        ' "raw" or "side1" on their own still means the tracks the box starts with.
        If req.Cylinders.Count = 0 AndAlso Not req.All Then
            For Each c In ParseSpec(DefaultSpec).Cylinders
                req.Cylinders.Add(c)
            Next
        End If
        Return req
    End Function

    Private Shared Function WantedCylinders(img As DmkImage, req As Request) As List(Of Integer)
        Dim present As New List(Of Integer)
        For Each t In img.Tracks
            If Not present.Contains(t.Cylinder) Then present.Add(t.Cylinder)
        Next
        present.Sort()
        If req.All Then Return present
        Dim wanted As New List(Of Integer)
        For Each c In req.Cylinders
            If present.Contains(c) Then wanted.Add(c)
        Next
        wanted.Sort()
        Return wanted
    End Function

    ' ------------------------------------------------------------------ one track

    Private Shared Sub AppendTrack(sb As StringBuilder, img As DmkImage, cyl As Integer, side As Integer, raw As Boolean)
        Dim t As DmkTrack = img.Track(cyl, side)
        Dim head As String = "Cylinder " & cyl.ToString(CultureInfo.InvariantCulture) & ", side " & side.ToString(CultureInfo.InvariantCulture)
        sb.AppendLine(New String("-"c, 78))
        If t Is Nothing OrElse Not t.Present Then
            sb.AppendLine(head & " — not in the file")
            sb.AppendLine()
            Return
        End If

        Dim mfmCount As Integer = 0
        For Each s In t.Sectors
            If s.Mfm Then mfmCount += 1
        Next
        Dim density As String
        If t.Sectors.Count = 0 Then
            density = "no sectors"
        ElseIf mfmCount = t.Sectors.Count Then
            density = "double density (MFM)"
        ElseIf mfmCount = 0 Then
            density = "single density (FM)"
        Else
            density = "mixed: " & mfmCount.ToString(CultureInfo.InvariantCulture) & " MFM, " &
                      (t.Sectors.Count - mfmCount).ToString(CultureInfo.InvariantCulture) & " FM"
        End If

        sb.AppendLine(head & " — " & t.Sectors.Count.ToString(CultureInfo.InvariantCulture) & " sector(s), " & density)
        sb.AppendLine("  IDAM table: " & t.PointerCount.ToString(CultureInfo.InvariantCulture) & " of 64 slots used" &
                      If(t.BadPointers > 0, ", " & t.BadPointers.ToString(CultureInfo.InvariantCulture) & " pointing nowhere useful", ""))

        If t.Sectors.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("  slot  ptr   C   H   R   N  bytes  den  id CRC   DAM  data CRC")
            For Each s In t.Sectors
                sb.AppendLine("  " &
                    s.PointerIndex.ToString(CultureInfo.InvariantCulture).PadLeft(4) & "  " &
                    s.Pointer.ToString("X4", CultureInfo.InvariantCulture) & " " &
                    s.IdCylinder.ToString(CultureInfo.InvariantCulture).PadLeft(3) & " " &
                    s.IdSide.ToString(CultureInfo.InvariantCulture).PadLeft(3) & " " &
                    s.IdSector.ToString(CultureInfo.InvariantCulture).PadLeft(3) & " " &
                    s.SizeCode.ToString(CultureInfo.InvariantCulture).PadLeft(3) & "  " &
                    s.Size.ToString(CultureInfo.InvariantCulture).PadLeft(5) & "  " &
                    If(s.Mfm, "MFM", " FM") & "  " &
                    If(s.IdCrcOk, "ok     ", "BAD    ") & "  " &
                    If(s.DataMark = 0, "--", s.DataMark.ToString("X2", CultureInfo.InvariantCulture)) & "   " &
                    If(s.DataMark = 0, "no data", If(s.DataCrcOk, "ok", "BAD")) &
                    If(s.SyncNote.Length > 0, "   " & s.SyncNote, "") &
                    If(s.PointerNotAtFe, "   pointer not on an FE", "") &
                    If(s.PointerPastEnd, "   pointer past the end of the track", "") &
                    If(s.NonIbmSize, "   non-IBM length", ""))
            Next
        End If

        For Each s In t.Sectors
            If s.Data Is Nothing OrElse s.Data.Length = 0 Then Continue For
            sb.AppendLine()
            sb.AppendLine("  Sector " & s.IdSector.ToString(CultureInfo.InvariantCulture) & " data — " &
                          s.Data.Length.ToString(CultureInfo.InvariantCulture) & " bytes, " &
                          If(s.DataCrcOk, "CRC ok", "CRC BAD (stored " & s.StoredCrc.ToString("X4", CultureInfo.InvariantCulture) &
                             "h, computed " & s.ComputedCrc.ToString("X4", CultureInfo.InvariantCulture) & "h)"))
            AppendHex(sb, s.Data, 0, s.Data.Length, "    ")
        Next

        If raw AndAlso t.Raw IsNot Nothing Then
            sb.AppendLine()
            sb.AppendLine("  Raw track bytes as the file holds them — " & t.Raw.Length.ToString(CultureInfo.InvariantCulture) &
                          " bytes, IDAM table first")
            AppendHex(sb, t.Raw, 0, t.Raw.Length, "    ")
        End If
        sb.AppendLine()
    End Sub

    ' ------------------------------------------------------------------ hex

    ''' <summary>Offset, 16 bytes, then the printable characters — the shape every hex tool uses.</summary>
    Private Shared Sub AppendHex(sb As StringBuilder, data As Byte(), start As Integer, count As Integer, indent As String)
        Dim line As New StringBuilder()
        Dim text As New StringBuilder()
        Dim i As Integer = 0
        While i < count
            line.Clear()
            text.Clear()
            line.Append(indent)
            line.Append((start + i).ToString("X4", CultureInfo.InvariantCulture))
            line.Append("  ")
            For j As Integer = 0 To PerLine - 1
                If i + j < count Then
                    Dim b As Byte = data(start + i + j)
                    line.Append(b.ToString("x2", CultureInfo.InvariantCulture))
                    line.Append(" "c)
                    text.Append(If(b >= 32 AndAlso b < 127, ChrW(b), "."c))
                Else
                    line.Append("   ")
                    text.Append(" "c)
                End If
                If j = 7 Then line.Append(" "c)
            Next
            line.Append(" |")
            line.Append(text)
            line.Append("|"c)
            sb.AppendLine(line.ToString())
            i += PerLine
        End While
    End Sub

End Class
