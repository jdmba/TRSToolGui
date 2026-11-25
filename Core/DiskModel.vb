' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  DiskModel.vb — what the GUI knows about the open image, built from the
'  engine's JSON answers (info, dir, and one "disk info : check" run; on a
'  CP/M disk "disk info : free").
'  Nothing here parses trs80-tool output; everything comes through
'  Trs80Tool's typed wrappers.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Threading.Tasks

''' <summary>One file on the disk, as shown in the list.</summary>
Public Class DiskFile
    ''' <summary>The name as trs80-tool gives it: NAME.EXT, or PARTDIR/NAME.EXT on a hard disk.</summary>
    Public Property FullName As String = ""
    ''' <summary>Partition directory ("HARDA") for hard disks, else "".</summary>
    Public Property PartitionDir As String = ""
    ''' <summary>NAME.EXT without the partition directory.</summary>
    Public Property TrsName As String = ""
    Public Property BaseName As String = ""
    Public Property Extension As String = ""
    Public Property Size As Long
    Public Property DateText As String = ""
    Public Property DateValue As DateTime? = Nothing
    Public Property Hidden As Boolean
    Public Property IsSystem As Boolean
    Public Property Protection As String = ""
    Public Property CpmUser As Integer? = Nothing
    Public Property CpmAttributes As String = ""
    Public Property ContentType As String = ""
    Public Property Description As String = ""
    Public Property DecodeError As String = ""
    Public Property Damage As String = ""
    Public Property Sha256 As String = ""
    ''' <summary>True for a file of a TRSDOS-family disk (the properties below apply only then).</summary>
    Public Property IsTrsdosEntry As Boolean
    ''' <summary>"" = no update password, "PASSWORD", or Nothing when it is some other password (or not known).</summary>
    Public Property UpdatePassword As String = Nothing
    Public Property UpdatePasswordHash As Integer? = Nothing
    ''' <summary>"" = no access password, "PASSWORD", or Nothing when it is some other password (or not known).</summary>
    Public Property AccessPassword As String = Nothing
    Public Property AccessPasswordHash As Integer? = Nothing
    ''' <summary>The disk keeps a time stamp where the access password would go, so none can be set (LDOS 5.3, LS-DOS 6.3).</summary>
    Public Property TimeStampsReplaceAccessPassword As Boolean
    ''' <summary>Directory entry bytes 0 and 1 as stored, when the tool gives them.</summary>
    Public Property EntryByte0 As Integer? = Nothing
    Public Property EntryByte1 As Integer? = Nothing
    ''' <summary>The disk's DOS is of the LDOS line (LDOS 5, TRSDOS 6, LS-DOS 6), where byte 1 has the CREATEd and MOD bits.</summary>
    Public Property HasLdosBits As Boolean

    ''' <summary>The file was made by CREATE (pre-allocated): LDOS line, entry byte 1 bit 7.</summary>
    Public ReadOnly Property IsCreated As Boolean
        Get
            Return HasLdosBits AndAlso EntryByte1.HasValue AndAlso (EntryByte1.Value And &H80) <> 0
        End Get
    End Property

    ''' <summary>The file has changed since the last backup (MOD): LDOS line, entry byte 1 bit 6.</summary>
    Public ReadOnly Property IsModified As Boolean
        Get
            Return HasLdosBits AndAlso EntryByte1.HasValue AndAlso (EntryByte1.Value And &H40) <> 0
        End Get
    End Property

    ''' <summary>
    ''' Attribute letters in TRSTool's style: I invisible, S system, C created and * MOD (LDOS line), plus CP/M
    ''' R (read-only) and * (archive).
    ''' </summary>
    Public ReadOnly Property AttributeLetters As String
        Get
            Dim s As String = ""
            If Hidden Then s &= "I"
            If IsSystem Then s &= "S"
            If IsCreated Then s &= "C"
            If IsModified Then s &= "*"
            If CpmAttributes.Contains("R/O") Then s &= "R"
            If CpmAttributes.Contains("ARC") Then s &= "*"
            Return s
        End Get
    End Property

    ''' <summary>CP/M: the file is read-only (R/O).</summary>
    Public ReadOnly Property CpmReadOnly As Boolean
        Get
            Return CpmAttributes.Contains("R/O")
        End Get
    End Property

    ''' <summary>CP/M: the file has the SYS attribute, which hides it from DIR. trs80-tool reports it as both hidden and system.</summary>
    Public ReadOnly Property CpmSys As Boolean
        Get
            Return CpmAttributes.Contains("SYS") OrElse (CpmUser.HasValue AndAlso (Hidden OrElse IsSystem))
        End Get
    End Property

    ''' <summary>
    ''' The option words of "disk attrib" that give a CP/M file these two attributes; empty when it has them already.
    ''' (--read-only and --no-read-only are trs80-tool 5.2.0's; SYS is its --system.)
    ''' </summary>
    Public Function CpmAttribWords(wantReadOnly As Boolean, wantSys As Boolean) As List(Of String)
        Dim w As New List(Of String)
        If wantReadOnly <> CpmReadOnly Then w.Add(If(wantReadOnly, "--read-only", "--no-read-only"))
        If wantSys <> CpmSys Then w.Add(If(wantSys, "--system", "--no-system"))
        Return w
    End Function

    ''' <summary>"none", "PASSWORD" or "set (hash 0x....)" for a password of this file; "" when not known.</summary>
    Public Shared Function PasswordText(known As String, hash As Integer?) As String
        If known IsNot Nothing Then Return If(known.Length = 0, "none", known)
        If hash.HasValue Then Return "set (hash 0x" & hash.Value.ToString("X4", CultureInfo.InvariantCulture) & ")"
        Return ""
    End Function

    ''' <summary>The TRS-80 name in NAME/EXT form, as DOS commands show it.</summary>
    Public ReadOnly Property SlashName As String
        Get
            If Extension.Length = 0 Then Return BaseName
            Return BaseName & "/" & Extension
        End Get
    End Property

    Public Shared Function FromJson(j As Trs80Tool.FileJson) As DiskFile
        Dim f As New DiskFile With {.FullName = If(j.Name, ""), .Size = j.Size, .ContentType = If(j.ContentType, ""), .Description = If(j.Description, ""), .DecodeError = If(j.Error, ""), .Sha256 = If(j.Sha256, "")}
        Dim slash As Integer = f.FullName.LastIndexOf("/"c)
        If slash >= 0 Then
            f.PartitionDir = f.FullName.Substring(0, slash)
            f.TrsName = f.FullName.Substring(slash + 1)
        Else
            f.TrsName = f.FullName
        End If
        Dim dot As Integer = f.TrsName.LastIndexOf("."c)
        If dot > 0 Then
            f.BaseName = f.TrsName.Substring(0, dot)
            f.Extension = f.TrsName.Substring(dot + 1)
        Else
            f.BaseName = f.TrsName
            f.Extension = ""
        End If
        If j.Date IsNot Nothing Then
            f.DateText = j.Date
            f.DateValue = ParseToolDate(j.Date)
        End If
        If j.Trsdos IsNot Nothing Then
            f.IsTrsdosEntry = True
            f.Hidden = j.Trsdos.Hidden
            f.IsSystem = j.Trsdos.OsFile
            f.Protection = If(j.Trsdos.Protection, "")
            f.UpdatePassword = j.Trsdos.UpdatePassword
            f.UpdatePasswordHash = j.Trsdos.UpdatePasswordHash
            f.AccessPassword = j.Trsdos.AccessPassword
            f.AccessPasswordHash = j.Trsdos.AccessPasswordHash
            f.TimeStampsReplaceAccessPassword = j.Trsdos.TimeStampsReplaceAccessPassword
            f.EntryByte0 = j.Trsdos.EntryByte0
            f.EntryByte1 = j.Trsdos.EntryByte1
        ElseIf j.Cpm IsNot Nothing Then
            f.Hidden = j.Cpm.Hidden
            f.IsSystem = j.Cpm.OsFile
            f.CpmUser = j.Cpm.User
            If j.Cpm.Attributes IsNot Nothing Then f.CpmAttributes = String.Join(",", j.Cpm.Attributes)
        Else
            For Each g In New Trs80Tool.GenericContainerJson() {j.MicroDos, j.UcsdPascal, j.ModelIITrsdos, j.TrsdosII, j.XenixTar}
                If g IsNot Nothing Then
                    f.Hidden = g.Hidden
                    f.IsSystem = g.OsFile
                    Exit For
                End If
            Next
        End If
        If j.Damage IsNot Nothing Then f.Damage = j.Damage.ToString()
        Return f
    End Function

    ''' <summary>"YYYY-MM-DD", "YYYY-MM-DDTHH:MM", "YYYY-MM" (a DOS that records no day), "YYYY", optional trailing Z.</summary>
    Public Shared Function ParseToolDate(text As String) As DateTime?
        If String.IsNullOrWhiteSpace(text) Then Return Nothing
        Dim t As String = text.Trim().TrimEnd("Z"c)
        Dim formats As String() = {"yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd", "yyyy-MM", "yyyy"}
        Dim d As DateTime
        If DateTime.TryParseExact(t, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return d
        Return Nothing
    End Function

    ''' <summary>Date for the list: local short date, with time when the DOS recorded one.</summary>
    Public ReadOnly Property DateDisplay As String
        Get
            If Not DateValue.HasValue Then Return DateText
            If DateText.Contains("T") Then Return DateValue.Value.ToString("d", CultureInfo.CurrentCulture) & " " & DateValue.Value.ToString("t", CultureInfo.CurrentCulture)
            If DateText.Length <= 4 Then Return DateText
            If DateText.Length <= 7 Then Return DateValue.Value.ToString("MMM yyyy", CultureInfo.CurrentCulture)
            Return DateValue.Value.ToString("d", CultureInfo.CurrentCulture)
        End Get
    End Property
End Class

''' <summary>A partition of a hard-disk image.</summary>
Public Class DiskPartition
    Public Property Number As Integer
    Public Property DirName As String = ""
    Public Property DiskName As String = ""
    Public Property Dos As String = ""
    Public Property Location As String = ""
    Public Property FileCount As Integer
    Public Property Writable As Boolean = True
    Public Property ReadOnlyReason As String = ""
    Public Property CylinderCount As Integer
    Public Property HeadCount As Integer
    Public Property Info As Trs80Tool.OsInfoJson
    Public ReadOnly Property Display As String
        Get
            Dim n As String = If(DiskName.Length = 0, "(no name)", DiskName)
            Return Number.ToString(CultureInfo.InvariantCulture) & ": " & n & "  (" & Dos & ", " & FileCount.ToString(CultureInfo.InvariantCulture) & " files" & If(Writable, "", ", read-only") & ")"
        End Get
    End Property
End Class

''' <summary>The open disk image.</summary>
Public Class DiskImage
    Public Property Path As String = ""
    Public Property ImageFormat As String = ""        ' JV1 JV3 DMK HFE IMD SCP HDV CAS WAV
    Public Property IsHardDisk As Boolean
    Public Property IsCassette As Boolean
    Public Property Description As String = ""        ' tool's one-line description
    Public Property DosName As String = ""            ' e.g. "LDOS 5.3.1"
    Public Property VolumeName As String = ""
    Public Property VolumeDate As String = ""
    Public Property AutoCommand As String = ""
    Public Property MasterPassword As String = ""     ' "PASSWORD", "blank", "set (hash ...)" or ""
    ''' <summary>The NEWDOS/80 PDRIVE line as trs80-tool gives it, or "".</summary>
    Public Property PdriveText As String = ""
    ''' <summary>Where trs80-tool got PdriveText: the disk's own drive table, or worked out from the layout.</summary>
    Public Property PdriveSource As String = ""
    Public Property Geometry As Trs80Tool.GeometryInfoJson = Nothing
    Public Property HardDiskGeometry As Trs80Tool.HardDiskGeometryJson = Nothing
    Public Property AllocationSides As Integer
    ''' <summary>Free space in words, e.g. "90 KB free (72 granules of 1,280 bytes each)"; "" when not known.</summary>
    Public Property FreeText As String = ""
    ''' <summary>Bytes new files could be given; -1 when not known.</summary>
    Public Property FreeBytes As Long = -1
    Public Property FreeGranules As Integer
    Public Property GranuleBytes As Integer
    Public Property WriteProtected As Boolean? = Nothing
    ''' <summary>Nothing when writable, else trs80-tool's reason.</summary>
    Public Property ReadOnlyReason As String = Nothing
    ''' <summary>Which DOS's rules a change follows, by trs80-tool's name ("ldos-5"); "" when it has none for the disk.</summary>
    Public Property Rules As String = ""
    ''' <summary>"check" found damage or a doubtful layout.</summary>
    Public Property Damaged As Boolean
    ''' <summary>What "check" printed; "" when it doesn't apply to this disk.</summary>
    Public Property CheckText As String = ""
    ''' <summary>What "check" found, most serious first.</summary>
    Public Property Findings As New List(Of Trs80Tool.FindingJson)
    ''' <summary>Signs of copy protection trs80-tool found (floppies).</summary>
    Public Property ProtectionSigns As New List(Of Trs80Tool.ProtectionSignJson)
    ''' <summary>What the imaging did to the disk that looks like damage (floppies).</summary>
    Public Property ImagingArtifacts As New List(Of Trs80Tool.ImagingArtifactJson)
    Public Property Warnings As New List(Of String)
    Public Property Partitions As New List(Of DiskPartition)
    Public Property ActivePartition As DiskPartition = Nothing
    Public Property AllFiles As New List(Of DiskFile)
    Public Property DetailLines As New List(Of String)  ' from "disk info"

    Public ReadOnly Property IsWritable As Boolean
        Get
            Return ReadOnlyReason Is Nothing AndAlso Not IsCassette
        End Get
    End Property

    Public ReadOnly Property IsCpm As Boolean
        Get
            Return DosName.IndexOf("CP/M", StringComparison.OrdinalIgnoreCase) >= 0
        End Get
    End Property

    ''' <summary>DOS family is TRSDOS-like (the viewer treats BASIC/CMD/EDTASM only on these, as TRSTool did).</summary>
    Public ReadOnly Property IsTrsdosFamily As Boolean
        Get
            If DosName.Length = 0 Then Return False
            Dim d As String = DosName.ToUpperInvariant()
            If d.Contains("CP/M") OrElse d.Contains("PASCAL") OrElse d.Contains("XENIX") OrElse d.Contains("MICRODOS") Then Return False
            Return True
        End Get
    End Property

    ''' <summary>True for the DOSes of the LDOS line, whose directory entries have the CREATEd and MOD bits.</summary>
    Public Shared Function IsLdosLine(dosName As String) As Boolean
        If String.IsNullOrEmpty(dosName) Then Return False
        Dim d As String = dosName.Trim().ToUpperInvariant()
        Return d.StartsWith("LDOS", StringComparison.Ordinal) OrElse d.StartsWith("LS-DOS", StringComparison.Ordinal) OrElse d.StartsWith("TRSDOS 6", StringComparison.Ordinal)
    End Function

    ''' <summary>The --partition argument for mutations, or Nothing for floppies.</summary>
    Public ReadOnly Property PartitionArg As String
        Get
            If Not IsHardDisk OrElse ActivePartition Is Nothing Then Return Nothing
            Return ActivePartition.Number.ToString(CultureInfo.InvariantCulture)
        End Get
    End Property

    ''' <summary>Number of signs of copy protection that say "protected" (a program's own scheme or the DOS's).</summary>
    Public ReadOnly Property ProtectedSignCount As Integer
        Get
            Dim n As Integer = 0
            For Each s In ProtectionSigns
                If String.Equals(s.Verdict, "protected", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(s.Verdict, "dosProtected", StringComparison.OrdinalIgnoreCase) Then n += 1
            Next
            Return n
        End Get
    End Property

    ''' <summary>Files of the active partition (or all files on a floppy), optionally filtered.</summary>
    Public Function VisibleFiles(showInvisible As Boolean, showSystem As Boolean) As List(Of DiskFile)
        Dim result As New List(Of DiskFile)
        For Each f In AllFiles
            If IsHardDisk AndAlso ActivePartition IsNot Nothing AndAlso Not String.Equals(f.PartitionDir, ActivePartition.DirName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If f.Hidden AndAlso Not showInvisible Then Continue For
            If f.IsSystem AndAlso Not showSystem Then Continue For
            result.Add(f)
        Next
        Return result
    End Function

    Public Function FindFile(trsName As String) As DiskFile
        Dim wanted As String = trsName.Replace("/"c, "."c)
        For Each f In AllFiles
            If IsHardDisk AndAlso ActivePartition IsNot Nothing AndAlso Not String.Equals(f.PartitionDir, ActivePartition.DirName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If String.Equals(f.TrsName, wanted, StringComparison.OrdinalIgnoreCase) Then Return f
        Next
        Return Nothing
    End Function

    ''' <summary>Approximate capacity in bytes from the geometry (data-track layout × cylinders × sides used). 0 when unknown.</summary>
    Public ReadOnly Property ApproxCapacityBytes As Long
        Get
            If IsHardDisk Then
                If ActivePartition Is Nothing OrElse HardDiskGeometry Is Nothing Then Return 0
                Return CLng(ActivePartition.CylinderCount) * ActivePartition.HeadCount * HardDiskGeometry.SectorsPerTrack * HardDiskGeometry.SectorSize
            End If
            If Geometry Is Nothing OrElse Geometry.DataTrack Is Nothing Then Return 0
            Dim sides As Integer = If(AllocationSides > 0, AllocationSides, Geometry.SideCount)
            Return CLng(Geometry.CylinderCount) * sides * Geometry.DataTrack.SectorSpan * Geometry.DataTrack.SectorSize
        End Get
    End Property

    Public ReadOnly Property DensityText As String
        Get
            If Geometry Is Nothing Then Return ""
            Dim data As String = If(Geometry.DataTrack IsNot Nothing, Geometry.DataTrack.Density, "")
            Dim boot As String = If(Geometry.BootTrack IsNot Nothing, Geometry.BootTrack.Density, data)
            If data.Length = 0 Then Return ""
            If Geometry.Homogeneous OrElse String.Equals(boot, data, StringComparison.OrdinalIgnoreCase) Then Return Capitalize(data)
            Return "Mixed (boot " & boot & ", data " & data & ")"
        End Get
    End Property

    Private Shared Function Capitalize(s As String) As String
        If String.IsNullOrEmpty(s) Then Return s
        Return Char.ToUpperInvariant(s(0)) & s.Substring(1)
    End Function

    ''' <summary>
    ''' Load everything about an image: "info" and "dir" (which also read cassettes), then one
    ''' "disk info : check" run for the status. Throws ToolException if the image can't be read at all.
    ''' </summary>
    Public Shared Async Function LoadAsync(tool As Trs80Tool, imagePath As String, keepPartitionNumber As Integer) As Task(Of DiskImage)
        Dim img As New DiskImage With {.Path = imagePath}

        Dim infoTask As Task(Of Trs80Tool.InfoJson) = tool.InfoJsonAsync(imagePath)
        Dim dirTask As Task(Of Trs80Tool.FilesInfileJson) = tool.DirJsonAsync(imagePath, True)
        Dim info As Trs80Tool.InfoJson = Await infoTask
        Dim notAnImage As String = Nothing
        If info.Error IsNot Nothing AndAlso info.Floppy Is Nothing AndAlso info.HardDisk Is Nothing AndAlso info.Cassette Is Nothing Then
            notAnImage = info.Error
        ElseIf info.Floppy Is Nothing AndAlso info.HardDisk Is Nothing AndAlso info.Cassette Is Nothing Then
            notAnImage = """" & imagePath & """ is not a floppy, hard disk or cassette image." & If(info.Description Is Nothing, "", " It is: " & info.Description)
        End If
        If notAnImage IsNot Nothing Then
            ' Let the other call finish before giving up, so nothing is left running.
            Try
                Await dirTask
            Catch ex As Exception
            End Try
            Throw New ToolException(notAnImage)
        End If
        img.Description = If(info.Description, "")
        If info.Warnings IsNot Nothing Then img.Warnings.AddRange(info.Warnings)
        If info.Error IsNot Nothing Then img.Warnings.Add(info.Error)

        If info.Floppy IsNot Nothing Then
            img.ImageFormat = If(info.Floppy.Format, "")
            img.Geometry = info.Floppy.Geometry
            If info.Floppy.Os IsNot Nothing Then
                img.DosName = If(info.Floppy.Os.Name, "")
                img.ApplyOsInfo(info.Floppy.Os)
            End If
        ElseIf info.HardDisk IsNot Nothing Then
            img.IsHardDisk = True
            img.ImageFormat = "HDV"
            If info.HardDisk.Partitions IsNot Nothing Then
                For Each p In info.HardDisk.Partitions
                    Dim dp As New DiskPartition With {.Number = p.Number, .CylinderCount = p.CylinderCount, .HeadCount = p.HeadCount, .Info = p.Os}
                    If p.Os IsNot Nothing Then
                        dp.Dos = If(p.Os.Name, "")
                        dp.FileCount = p.Os.FileCount
                        If p.Os.Trsdos IsNot Nothing AndAlso p.Os.Trsdos.DiskName IsNot Nothing Then dp.DiskName = p.Os.Trsdos.DiskName.Trim()
                    End If
                    img.Partitions.Add(dp)
                Next
            End If
        Else
            img.IsCassette = True
            img.ImageFormat = If(HostFiles.ImageExtension(imagePath) = ".wav", "WAV", "CAS")
            img.ReadOnlyReason = UiText.T("disk.readonly.cassette")
        End If

        Dim dirJson As Trs80Tool.FilesInfileJson = Await dirTask
        If dirJson.Warnings IsNot Nothing Then
            For Each w In dirJson.Warnings
                If Not img.Warnings.Contains(w) Then img.Warnings.Add(w)
            Next
        End If
        If dirJson.Error IsNot Nothing AndAlso Not img.Warnings.Contains(dirJson.Error) Then img.Warnings.Add(dirJson.Error)
        For Each fj In dirJson.Files
            img.AllFiles.Add(DiskFile.FromJson(fj))
        Next
        If img.DosName.Length = 0 AndAlso dirJson.Os IsNot Nothing Then img.DosName = dirJson.Os

        If img.IsHardDisk Then
            img.Partitions.Sort(Function(a, b) a.Number.CompareTo(b.Number))
            ' The files of a hard disk are named PARTDIR/NAME.EXT; the directory names come in disk order,
            ' which is the partitions' order.
            Dim dirNames As New List(Of String)
            For Each f In img.AllFiles
                If f.PartitionDir.Length > 0 AndAlso Not dirNames.Contains(f.PartitionDir) Then dirNames.Add(f.PartitionDir)
            Next
            For i As Integer = 0 To img.Partitions.Count - 1
                If i < dirNames.Count Then img.Partitions(i).DirName = dirNames(i) Else img.Partitions(i).DirName = img.Partitions(i).DiskName
            Next
            If img.Partitions.Count > 0 Then
                img.ActivePartition = img.Partitions.Find(Function(x) x.Number = keepPartitionNumber)
                If img.ActivePartition Is Nothing Then img.ActivePartition = img.Partitions(0)
                img.DosName = img.ActivePartition.Dos
                If img.ActivePartition.Info IsNot Nothing Then img.ApplyOsInfo(img.ActivePartition.Info)
            End If
        End If

        If Not img.IsCassette Then Await img.RefreshStatusAsync(tool)
        img.MarkLdosBits()
        Return img
    End Function

    ''' <summary>Tell each file whether its DOS has the LDOS line's CREATEd and MOD bits.</summary>
    Private Sub MarkLdosBits()
        For Each f In AllFiles
            Dim dos As String = DosName
            If IsHardDisk Then
                Dim dirName As String = f.PartitionDir
                Dim p As DiskPartition = Partitions.Find(Function(x) String.Equals(x.DirName, dirName, StringComparison.OrdinalIgnoreCase))
                dos = If(p IsNot Nothing, p.Dos, "")
            End If
            f.HasLdosBits = f.IsTrsdosEntry AndAlso IsLdosLine(dos)
        Next
    End Sub

    Private Sub ApplyOsInfo(os As Trs80Tool.OsInfoJson)
        VolumeName = ""
        VolumeDate = ""
        AutoCommand = ""
        MasterPassword = ""
        PdriveText = ""
        PdriveSource = ""
        If os.Trsdos IsNot Nothing Then
            Dim t As Trs80Tool.TrsdosInfoJson = os.Trsdos
            VolumeName = If(t.DiskName, "").Trim()
            VolumeDate = If(t.DiskDate, "").Trim()
            AutoCommand = If(t.AutoCommand, "")
            AllocationSides = t.AllocationSides
            ' trs80-tool names the two passwords every DOS writes by default (by the hash of the disk's own layout).
            If t.GatPasswordText IsNot Nothing Then
                MasterPassword = If(t.GatPasswordText.Length = 0, "blank", t.GatPasswordText)
            ElseIf t.GatPassword.HasValue Then
                MasterPassword = "set (hash 0x" & t.GatPassword.Value.ToString("X4", CultureInfo.InvariantCulture) & ")"
            End If
            If t.Pdrive IsNot Nothing AndAlso Not String.IsNullOrEmpty(t.Pdrive.Text) Then
                PdriveText = t.Pdrive.Text
                If String.Equals(t.Pdrive.Source, "table", StringComparison.OrdinalIgnoreCase) Then
                    PdriveSource = "from the disk's own drive table" & If(t.Pdrive.TableCylinder.HasValue AndAlso t.Pdrive.TableCylinder.Value > 0, " on cylinder " & t.Pdrive.TableCylinder.Value.ToString(CultureInfo.InvariantCulture), "") & ", read by trs80-tool"
                Else
                    PdriveSource = "worked out by trs80-tool from the layout; the disk has no drive table, so TI's letter and TSR aren't known"
                End If
            End If
            If t.GatError IsNot Nothing AndAlso Not Warnings.Contains("Allocation table: " & t.GatError) Then Warnings.Add("Allocation table: " & t.GatError)
        ElseIf os.Cpm IsNot Nothing Then
            AllocationSides = os.Cpm.AllocationSides
        End If
    End Sub

    ''' <summary>Free space in words, with the granules (CP/M: blocks) new files can't use counted apart.</summary>
    Private Shared Function FreeSpaceText(f As Trs80Tool.FreeResultJson) As String
        Dim kb As Long = f.FreeBytes \ 1024
        If f.IsBlocks Then
            Dim blockText As String = kb.ToString("N0", CultureInfo.InvariantCulture) & " KB free (" &
                f.FreeBlocks.ToString("N0", CultureInfo.InvariantCulture) & If(f.FreeBlocks = 1, " block of ", " blocks of ") &
                f.BlockBytes.ToString("N0", CultureInfo.InvariantCulture) & " bytes each)"
            If f.MissingBlocks > 0 Then blockText &= "; not usable: " & f.MissingBlocks.ToString("N0", CultureInfo.InvariantCulture) & " more that the image doesn't fully hold"
            Return blockText
        End If
        Dim text As String = kb.ToString("N0", CultureInfo.InvariantCulture) & " KB free (" &
            f.FreeGranules.ToString("N0", CultureInfo.InvariantCulture) & If(f.FreeGranules = 1, " granule of ", " granules of ") &
            f.GranuleBytes.ToString("N0", CultureInfo.InvariantCulture) & " bytes each)"
        Dim extra As New List(Of String)
        If f.MissingGranules > 0 Then extra.Add(f.MissingGranules.ToString("N0", CultureInfo.InvariantCulture) & " more that the image doesn't fully hold")
        If f.PastEndGranules > 0 Then extra.Add(f.PastEndGranules.ToString("N0", CultureInfo.InvariantCulture) & " more past the disk's proven end")
        If extra.Count > 0 Then text &= "; not usable: " & String.Join(", ", extra)
        Return text
    End Function

    ''' <summary>
    ''' Re-read the partition-dependent status in one run: free space, write-protect flag, whether and why the
    ''' disk is read-only, the partitions, and what "check" finds (damage, signs of copy protection, imaging artifacts).
    ''' </summary>
    Public Async Function RefreshStatusAsync(tool As Trs80Tool) As Task
        Dim st As Trs80Tool.DiskStatus = Await tool.DiskStatusAsync(Path, PartitionArg, IsCpm AndAlso Not IsHardDisk)
        FreeText = ""
        FreeBytes = -1
        FreeGranules = 0
        GranuleBytes = 0
        Rules = ""
        Damaged = False
        CheckText = ""
        Findings.Clear()
        ProtectionSigns.Clear()
        ImagingArtifacts.Clear()
        DetailLines.Clear()
        For Each w In st.Warnings
            If Not Warnings.Contains(w) Then Warnings.Add(w)
        Next
        If st.Info Is Nothing Then
            WriteProtected = Nothing
            ReadOnlyReason = If(String.IsNullOrWhiteSpace(st.Error), UiText.T("disk.readonly.nodescribe"), st.Error)
            Return
        End If
        Dim info As Trs80Tool.InfoResultJson = st.Info
        For Each line In st.InfoLines
            If line.StartsWith("    ", StringComparison.Ordinal) Then DetailLines.Add(line.Trim())
        Next
        WriteProtected = info.WriteProtected
        Rules = If(info.Rules, "")
        If info.HardDiskGeometry IsNot Nothing Then HardDiskGeometry = info.HardDiskGeometry
        If info.Writable Then
            ReadOnlyReason = Nothing
        ElseIf WriteProtected.GetValueOrDefault() Then
            ' The image's own write-protect flag, which this program reads itself and offers as a
            ' box to clear in Disk properties, so the reason is said in the user's language rather
            ' than passed through from trs80-tool in English.
            ReadOnlyReason = UiText.T("disk.readonly.writeprotected")
        Else
            ReadOnlyReason = If(String.IsNullOrWhiteSpace(info.ReadOnlyReason), UiText.T("disk.readonly.cantwrite"), info.ReadOnlyReason)
        End If
        ' A CP/M disk's free space comes from "free" (in blocks); "info" gives it only for TRSDOS-family disks.
        Dim free As Trs80Tool.FreeResultJson = If(info.Free, st.Free)
        If free IsNot Nothing Then
            FreeBytes = free.FreeBytes
            FreeGranules = free.FreeGranules
            GranuleBytes = free.GranuleBytes
            FreeText = FreeSpaceText(free)
        End If
        If info.Os IsNot Nothing Then
            If Not String.IsNullOrEmpty(info.Os.Name) Then DosName = info.Os.Name
            ApplyOsInfo(info.Os)
        End If

        If IsHardDisk AndAlso info.Partitions IsNot Nothing Then
            For Each pj In info.Partitions
                Dim number As Integer = pj.Number
                Dim dp As DiskPartition = Partitions.Find(Function(x) x.Number = number)
                If dp Is Nothing Then
                    dp = New DiskPartition With {.Number = pj.Number}
                    Partitions.Add(dp)
                End If
                If Not String.IsNullOrEmpty(pj.DiskName) Then dp.DiskName = pj.DiskName.Trim()
                If dp.DirName.Length = 0 Then dp.DirName = dp.DiskName
                If Not String.IsNullOrEmpty(pj.Os) Then dp.Dos = pj.Os
                dp.FileCount = pj.FileCount
                dp.Writable = pj.Writable
                dp.ReadOnlyReason = If(pj.ReadOnlyReason, "")
                If pj.Location IsNot Nothing Then
                    dp.CylinderCount = pj.Location.CylinderCount
                    dp.HeadCount = pj.Location.HeadCount
                    dp.Location = "cylinders " & pj.Location.CylinderOffset.ToString(CultureInfo.InvariantCulture) & "-" &
                        (pj.Location.CylinderOffset + pj.Location.CylinderCount - 1).ToString(CultureInfo.InvariantCulture) & ", heads " &
                        pj.Location.HeadOffset.ToString(CultureInfo.InvariantCulture) & "-" &
                        (pj.Location.HeadOffset + pj.Location.HeadCount - 1).ToString(CultureInfo.InvariantCulture)
                End If
            Next
            Partitions.Sort(Function(a, b) a.Number.CompareTo(b.Number))
        End If
        If IsHardDisk AndAlso Partitions.Count = 0 Then
            ReadOnlyReason = UiText.T("disk.readonly.nopartitions")
        End If

        If st.Check IsNot Nothing Then
            CheckText = String.Join(vbCrLf, st.CheckLines)
            If st.Check.Findings IsNot Nothing Then
                Findings.AddRange(st.Check.Findings)
                For Each f In st.Check.Findings
                    If String.Equals(f.Severity, "damage", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(f.Severity, "layout", StringComparison.OrdinalIgnoreCase) Then Damaged = True
                Next
            End If
            If st.Check.Protection IsNot Nothing Then ProtectionSigns.AddRange(st.Check.Protection)
            If st.Check.ImagingArtifacts IsNot Nothing Then ImagingArtifacts.AddRange(st.Check.ImagingArtifacts)
        End If
        MarkLdosBits()
    End Function
End Class
