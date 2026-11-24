' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  UiText.vb — the program's user-visible text, kept outside the code so it can
'  be translated. One table per language in Core\Lang\<code>.txt, embedded in
'  TRSToolGui.Core.dll and read on first use. UiText.T("key") gives the text in
'  the language in force; anything a translation has not filled in falls back to
'  English, and a key missing even from English comes back as the key in square
'  brackets, so a gap shows plainly on screen instead of throwing.
'
'  Languages lists only the languages whose table is actually in the build, so
'  the Preferences list can never offer one that would quietly stay English.
'
'  The disk analysis reports are deliberately NOT in these tables. They stay in
'  English: they are forensic findings thick with names that must not be
'  translated, and they change too often to keep translations of them true.
'
'  Table format (Core\Lang\en.txt is the master):
'    key=text            one entry per line, the first "=" separates them
'    # comment           a line whose first non-blank character is "#"
'    {0} {1}             where the program substitutes a value, in any order
'    \n \r \t \\ \s      newline, return, tab, backslash, and a space that has
'                        to survive at the start or end of the text
'  Text is used exactly as the table holds it when nothing is substituted, so
'  braces in plain text need no escaping. Only an entry the program passes
'  values to goes through String.Format, where a literal brace is {{ or }}.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>
''' The program's translated text. Read-only once loaded, so any thread may call T().
''' </summary>
Public NotInheritable Class UiText

    Private Sub New()
    End Sub

    ''' <summary>The language the code is written in, and the one every other table falls back to.</summary>
    Public Const DefaultCode As String = "en"

    ''' <summary>A language the program carries text for.</summary>
    Public NotInheritable Class Language
        ''' <summary>Two-letter code, and the name of the table file in Core\Lang.</summary>
        Public ReadOnly Property Code As String
        ''' <summary>The language's name in that language, for the Preferences list.</summary>
        Public ReadOnly Property NativeName As String
        Public Sub New(code As String, nativeName As String)
            _Code = code
            _NativeName = nativeName
        End Sub
    End Class

    ' Every language the project has a table file for. A code stays listed here once its
    ' translation starts; Languages below hides it until the table is really in the build.
    ' The name here is only a fallback: a table that defines lang.name gives its own name for
    ' itself, which is how an accented name reaches the Preferences list while these source
    ' files stay plain ASCII.
    Private Shared ReadOnly _declared As Language() = {
        New Language("en", "English"),
        New Language("es", "Espanol"),
        New Language("fr", "Francais"),
        New Language("it", "Italiano")}

    ''' <summary>The key a table uses to give its language's own name for itself.</summary>
    Private Const NameKey As String = "lang.name"

    Private Shared ReadOnly _gate As New Object()
    Private Shared _english As Dictionary(Of String, String) = Nothing
    Private Shared _available As List(Of Language) = Nothing
    Private Shared _active As Dictionary(Of String, String) = Nothing
    Private Shared _activeCode As String = DefaultCode
    Private Shared _problem As String = ""

    ''' <summary>
    ''' The languages this build can actually show, in the order the Preferences list shows them.
    ''' English is always first and always present.
    ''' </summary>
    Public Shared ReadOnly Property Languages As IReadOnlyList(Of Language)
        Get
            EnsureLoaded()
            Return _available
        End Get
    End Property

    ''' <summary>True if this build carries a readable table for the code.</summary>
    Public Shared Function IsAvailable(code As String) As Boolean
        If String.IsNullOrEmpty(code) Then Return False
        EnsureLoaded()
        For Each l As Language In _available
            If String.Equals(l.Code, code, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>The code of the language in force.</summary>
    Public Shared ReadOnly Property ActiveCode As String
        Get
            EnsureLoaded()
            Return _activeCode
        End Get
    End Property

    ''' <summary>
    ''' Empty when the tables loaded. Otherwise what went wrong, for Help to show. The program
    ''' keeps working either way: every key then comes back as the key in square brackets.
    ''' </summary>
    Public Shared ReadOnly Property Problem As String
        Get
            EnsureLoaded()
            Return _problem
        End Get
    End Property

    ''' <summary>
    ''' Puts a language in force. An unknown or unreadable code leaves English in force and says
    ''' why in Problem. Call it before the first window is built: text already on screen is not
    ''' revisited.
    ''' </summary>
    Public Shared Sub Use(code As String)
        SyncLock _gate
            LoadEnglishLocked()
            Dim want As String = If(code, "").Trim().ToLowerInvariant()
            If want.Length = 0 Then want = DefaultCode
            If want = DefaultCode Then
                _active = Nothing
                _activeCode = DefaultCode
                Return
            End If
            Dim known As Boolean = False
            For Each l As Language In _available
                If String.Equals(l.Code, want, StringComparison.OrdinalIgnoreCase) Then known = True
            Next
            If Not known Then
                _problem = "This build has no text for language """ & want & """; English is in use."
                _active = Nothing
                _activeCode = DefaultCode
                Return
            End If
            Dim loadProblem As String = ""
            Dim table As Dictionary(Of String, String) = ReadTable(want, loadProblem)
            If table Is Nothing Then
                _problem = loadProblem
                _active = Nothing
                _activeCode = DefaultCode
                Return
            End If
            _active = table
            _activeCode = want
        End SyncLock
    End Sub

    ''' <summary>
    ''' The text for a key. With values given, the text goes through String.Format, so a
    ''' translation may put {0} and {1} in whatever order its grammar needs. With none given the
    ''' text comes back exactly as the table holds it.
    ''' </summary>
    Public Shared Function T(key As String, ParamArray values As Object()) As String
        Dim text As String = Lookup(key)
        If values Is Nothing OrElse values.Length = 0 Then Return text
        Try
            Return String.Format(CultureInfo.CurrentCulture, text, values)
        Catch ex As FormatException
            ' A translation with a mistyped placeholder must not take the window down with it.
            Return text
        End Try
    End Function

    ''' <summary>
    ''' The text for a key with every line end turned into the carriage-return / line-feed pair.
    ''' The tables hold \n, which is what Eto's labels and text areas take; the Windows Forms controls
    ''' want the pair, so the Windows Forms windows ask for their multi-line text through this.
    ''' </summary>
    Public Shared Function TCrLf(key As String, ParamArray values As Object()) As String
        Dim text As String = T(key, values)
        If text.IndexOf(ControlChars.Lf) < 0 Then Return text
        Return text.Replace(ControlChars.CrLf, ControlChars.Lf).Replace(ControlChars.Lf, ControlChars.CrLf)
    End Function

    ''' <summary>True if the key is in the table of the language in force, or in English.</summary>
    Public Shared Function Has(key As String) As Boolean
        EnsureLoaded()
        If String.IsNullOrEmpty(key) Then Return False
        Dim v As String = Nothing
        Dim act As Dictionary(Of String, String) = _active
        If act IsNot Nothing AndAlso act.TryGetValue(key, v) Then Return True
        Dim eng As Dictionary(Of String, String) = _english
        Return eng IsNot Nothing AndAlso eng.TryGetValue(key, v)
    End Function

    ''' <summary>Every key English defines, for the key check that runs before a release.</summary>
    Public Shared Function EnglishKeys() As IReadOnlyCollection(Of String)
        EnsureLoaded()
        Dim eng As Dictionary(Of String, String) = _english
        If eng Is Nothing Then Return New List(Of String)()
        Return New List(Of String)(eng.Keys)
    End Function

    Private Shared Function Lookup(key As String) As String
        EnsureLoaded()
        If String.IsNullOrEmpty(key) Then Return ""
        Dim v As String = Nothing
        Dim act As Dictionary(Of String, String) = _active
        If act IsNot Nothing AndAlso act.TryGetValue(key, v) Then Return v
        Dim eng As Dictionary(Of String, String) = _english
        If eng IsNot Nothing AndAlso eng.TryGetValue(key, v) Then Return v
        Return "[" & key & "]"
    End Function

    Private Shared Sub EnsureLoaded()
        If _english IsNot Nothing Then Return
        SyncLock _gate
            LoadEnglishLocked()
        End SyncLock
    End Sub

    Private Shared Sub LoadEnglishLocked()
        If _english IsNot Nothing Then Return

        Dim problem As String = ""
        Dim table As Dictionary(Of String, String) = ReadTable(DefaultCode, problem)
        If table Is Nothing Then
            _english = New Dictionary(Of String, String)(StringComparer.Ordinal)
            If _problem.Length = 0 Then _problem = problem
        Else
            _english = table
        End If

        ' Work out which tables this build really carries. English stays listed either way: it is
        ' the fallback, and a list with nothing in it would leave Preferences no choice at all.
        Dim found As New List(Of Language)()
        For Each l As Language In _declared
            Dim name As String = l.NativeName
            If l.Code = DefaultCode Then
                name = NativeNameOf(_english, name)
            ElseIf HasTable(l.Code) Then
                ' Read the table once, for its own name for itself. A table that cannot be read,
                ' or that has no lang.name, keeps the name declared above.
                Dim why As String = ""
                Dim t As Dictionary(Of String, String) = ReadTable(l.Code, why)
                If t IsNot Nothing Then name = NativeNameOf(t, name)
            Else
                Continue For
            End If
            found.Add(New Language(l.Code, name))
        Next
        _available = found
    End Sub

    ''' <summary>A table's own name for its language, or the declared name when it gives none.</summary>
    Private Shared Function NativeNameOf(table As Dictionary(Of String, String), declared As String) As String
        If table Is Nothing Then Return declared
        Dim v As String = Nothing
        If table.TryGetValue(NameKey, v) AndAlso v IsNot Nothing AndAlso v.Trim().Length > 0 Then Return v
        Return declared
    End Function

    ''' <summary>True if the embedded table for this code is in the assembly.</summary>
    Private Shared Function HasTable(code As String) As Boolean
        Try
            Using st As Stream = GetType(UiText).Assembly.GetManifestResourceStream("TRSToolGui.Lang." & code & ".txt")
                Return st IsNot Nothing
            End Using
        Catch ex As IOException
            Return False
        End Try
    End Function

    ''' <summary>Reads one embedded table. Returns Nothing, and says why in problem, if it cannot.</summary>
    Private Shared Function ReadTable(code As String, ByRef problem As String) As Dictionary(Of String, String)
        problem = ""
        Dim name As String = "TRSToolGui.Lang." & code & ".txt"
        Try
            Using st As Stream = GetType(UiText).Assembly.GetManifestResourceStream(name)
                If st Is Nothing Then
                    problem = "The built-in text file " & name & " is missing from this build."
                    Return Nothing
                End If
                Using rd As New StreamReader(st, New UTF8Encoding(False), True)
                    Return Parse(rd)
                End Using
            End Using
        Catch ex As IOException
            problem = "The built-in text file " & name & " could not be read: " & ex.Message
            Return Nothing
        End Try
    End Function

    Private Shared Function Parse(rd As StreamReader) As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Dim line As String = rd.ReadLine()
        Do While line IsNot Nothing
            Dim trimmed As String = line.Trim()
            If trimmed.Length > 0 AndAlso Not trimmed.StartsWith("#", StringComparison.Ordinal) Then
                Dim eq As Integer = line.IndexOf("="c)
                If eq > 0 Then
                    Dim key As String = line.Substring(0, eq).Trim()
                    Dim value As String = Unescape(line.Substring(eq + 1).Trim())
                    If key.Length > 0 Then d(key) = value
                End If
            End If
            line = rd.ReadLine()
        Loop
        Return d
    End Function

    ''' <summary>Turns \n \r \t \\ and \s into their characters. An unknown escape is left as it is.</summary>
    Private Shared Function Unescape(s As String) As String
        If s.IndexOf("\"c) < 0 Then Return s
        Dim sb As New StringBuilder(s.Length)
        Dim i As Integer = 0
        Do While i < s.Length
            Dim c As Char = s(i)
            If c = "\"c AndAlso i + 1 < s.Length Then
                Dim n As Char = s(i + 1)
                Select Case n
                    Case "n"c
                        sb.Append(ControlChars.Lf)
                        i += 2
                        Continue Do
                    Case "r"c
                        sb.Append(ControlChars.Cr)
                        i += 2
                        Continue Do
                    Case "t"c
                        sb.Append(ControlChars.Tab)
                        i += 2
                        Continue Do
                    Case "s"c
                        sb.Append(" "c)
                        i += 2
                        Continue Do
                    Case "\"c
                        sb.Append("\"c)
                        i += 2
                        Continue Do
                End Select
            End If
            sb.Append(c)
            i += 1
        Loop
        Return sb.ToString()
    End Function

End Class
