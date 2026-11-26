' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  AppInfo.vb — the program's name and version stamps, in one place for every
'  user interface built on this library (the Windows Forms program and the
'  Eto.Forms program for Linux and macOS).
' ============================================================================
Option Strict On
Option Explicit On

Public Module AppInfo

    Public Const AppName As String = "TRSToolGui"
    ''' <summary>Build stamp, date-based (v26MMDD); it moves with every code change.</summary>
    Public Const AppVersion As String = "v261126"
    ''' <summary>The release number shown in Help > About.</summary>
    Public Const ReleaseVersion As String = "v1.8.3"
    ''' <summary>The date of that release, with a matching entry in Help\Changelog.txt.</summary>
    Public Const ReleaseDate As String = "November 26, 2026"
    Public Const AppAuthors As String = "Ira Goldklang and Claude AI"
    Public Const DownloadUrl As String = "https://github.com/jdmba/TRSToolGui"
    Public Const ToolUrl As String = "https://www.my-trs-80.com/tool/"
    ''' <summary>Where a problem report goes. Every public invitation to report names this and no other.</summary>
    Public Const IssuesUrl As String = "https://github.com/jdmba/TRSToolGui/issues"

End Module
