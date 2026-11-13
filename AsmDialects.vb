' ============================================================================
'  TRSToolGui — Copyright (c) 2026 Ira Goldklang.
'  Written by Ira Goldklang and Claude AI.
'  SPDX-License-Identifier: MIT   (full text: Help\License.txt)
'
'  AsmDialects.vb — which assembler a source file was written for, from the
'  telltales unique to each editor/assembler.
'
'  Each statement is split into label / opcode / operand (comment dropped).
'  Every dialect profile collects three kinds of evidence:
'    decisive        unique to that assembler — one is enough
'    beyond EDTASM   past Radio Shack EDTASM / Series I, but shared with
'                    others — enough only in a numbered (EDTASM-layout) file
'    rule-out        something that assembler never accepted
'  A dialect is claimed only when exactly one profile is decided.
'
'  Trained so far: EDAS (MISOSYS) 3.5 / IV / 4.3, Microsoft EDTASM-PLUS
'  (cassette 1979, disk 1981), Microsoft Macro-80, M-ZAL (Release 1 and 2),
'  MISOSYS MRAS (EDAS 4.3 + relocation; claimed only on its own features),
'  Zeus (Cosmopolitan 1983 / Hester 9th ed.), ZEN (Avalon 1979), Tandy ALDS. Radio Shack EDTASM / Series I is the baseline:
'  a file using only its vocabulary is reported as such, never claimed.
' ============================================================================
Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Module AsmDialects

    ''' <summary>One source statement: label, opcode (upper case), operand without the comment, or a *directive.</summary>
    Private Structure Stmt
        Public Label As String
        Public LabelColon As Boolean
        Public LabelPublic As Boolean
        Public Op As String
        Public Operand As String
        Public Directive As String
    End Structure

    ''' <summary>What one dialect profile found.</summary>
    Private NotInheritable Class Verdict
        Public ReadOnly Strong As New List(Of String)
        Public ReadOnly Medium As New List(Of String)
        Public Version As String = ""
        Public RuledOut As Boolean
        Private ReadOnly _seen As New HashSet(Of String)(StringComparer.Ordinal)

        Public Sub AddStrong(phrase As String)
            If _seen.Add(phrase) Then Strong.Add(phrase)
        End Sub
        Public Sub AddMedium(phrase As String)
            If _seen.Add(phrase) Then Medium.Add(phrase)
        End Sub
        ''' <summary>False for dialects whose files never use the numbered layout as evidence (Macro-80).</summary>
        Public MediumDecides As Boolean = True
        Public Function Decided(numbered As Boolean) As Boolean
            Return Not RuledOut AndAlso (Strong.Count > 0 OrElse (MediumDecides AndAlso numbered AndAlso Medium.Count > 0))
        End Function
    End Class

    ' ------------------------------------------------------------------ vocabulary

    ''' <summary>The Z-80 instruction set as these assemblers spell it (no undocumented SLL/SL1).</summary>
    Private ReadOnly Z80Ops As New HashSet(Of String)(StringComparer.Ordinal) From {
        "ADC", "ADD", "AND", "BIT", "CALL", "CCF", "CP", "CPD", "CPDR", "CPI", "CPIR", "CPL", "DAA", "DEC", "DI",
        "DJNZ", "EI", "EX", "EXX", "HALT", "IM", "IN", "INC", "IND", "INDR", "INI", "INIR", "JP", "JR", "LD", "LDD",
        "LDDR", "LDI", "LDIR", "NEG", "NOP", "OR", "OTDR", "OTIR", "OUT", "OUTD", "OUTI", "POP", "PUSH", "RES", "RET",
        "RETI", "RETN", "RL", "RLA", "RLC", "RLCA", "RLD", "RR", "RRA", "RRC", "RRCA", "RRD", "RST", "SBC", "SCF",
        "SET", "SLA", "SRA", "SRL", "SUB", "XOR"}

    ''' <summary>Radio Shack EDTASM / Series I pseudo-ops (Series I manual, and the keyword table in its EDTASM/CMD).
    ''' Every Z-80 assembler here has them, so they prove nothing.</summary>
    Private ReadOnly EdtasmPseudo As New HashSet(Of String)(StringComparer.Ordinal) From {
        "ORG", "EQU", "DEFL", "END", "DEFB", "DEFW", "DEFS", "DEFM"}

    ''' <summary>EDAS pseudo-ops beyond EDTASM, with the first release that has them.
    ''' 3.5 = the keyword table inside EDAS 3.5.2 (EDAS1/CMD, EDAS3/CMD); IV = 5th Ed. manual; 4.3 = its addendum.</summary>
    Private ReadOnly EdasPseudo As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"DB", "3.5"}, {"DW", "3.5"}, {"DS", "3.5"}, {"DM", "3.5"}, {"COM", "3.5"}, {"ERR", "3.5"},
        {"IF", "3.5"}, {"IFEQ", "3.5"}, {"IFGT", "3.5"}, {"IFLT", "3.5"}, {"ENDIF", "3.5"},
        {"PAGE", "3.5"}, {"SPACE", "3.5"}, {"SUBTTL", "3.5"}, {"TITLE", "3.5"},
        {"DC", "IV"}, {"LORG", "IV"}, {"ELSE", "IV"}, {"IFNE", "IV"}, {"IFDEF", "IV"}, {"IFNDEF", "IV"}, {"IFREF", "IV"},
        {"IFEQ$", "IV"}, {"IFGT$", "IV"}, {"IFLT$", "IV"}, {"IFNE$", "IV"}, {"MACRO", "IV"}, {"ENDM", "IV"},
        {"DATE", "4.3"}, {"TIME", "4.3"}, {"DSYM", "4.3"}, {"DX", "4.3"}, {"ENTRY", "4.3"}, {"OPTION", "4.3"},
        {"REF", "4.3"}, {"IF1", "4.3"}, {"IF2", "4.3"}, {"IF3", "4.3"}, {"REPT", "4.3"}, {"IRP", "4.3"},
        {"IRPC", "4.3"}, {"EXITM", "4.3"}}

    ''' <summary>EDAS *directives (column 1). *LIST is EDTASM's too and proves nothing.</summary>
    Private ReadOnly EdasDirectives As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"*GET", "3.5"}, {"*SEARCH", "IV"}, {"*MOD", "IV"}, {"*PREFIX", "IV"}, {"*INCLUDE", "4.3"}}

    ''' <summary>EDTASM-PLUS pseudo-ops beyond EDTASM: cassette 1979 (manual, reference card MS1104) and the
    ''' three the 1981 disk version added (PAGE, TITLE, INCLUDE).</summary>
    Private ReadOnly EdtasmPlusPseudo As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"COND", ""}, {"ENDC", ""}, {"MACRO", ""}, {"ENDM", ""},
        {"PAGE", "disk"}, {"TITLE", "disk"}, {"INCLUDE", "disk"}}

    ''' <summary>
    ''' Pseudo-ops that none of the trained dialects accept: relocating/linking assemblers (Macro-80, the
    ''' Tandy assembler of the TRSDOS sources) and x86 MASM. Rules out EDAS and EDTASM-PLUS.
    ''' (COND/ENDC are EDTASM-PLUS's and INCLUDE its disk version's; EDAS's own rule-outs add them.)
    ''' </summary>
    Private ReadOnly ForeignPseudo As New HashSet(Of String)(StringComparer.Ordinal) From {
        "PSECT", "EXT", "EXTRN", "EXTERNAL", "PUBLIC", "GLOBAL", "HEADER", "EJECT", "NOLOAD", "DEFT", "DEFE",
        "MACLIB", "ASEG", "CSEG", "DSEG", "COMMON", "NAME", "USING", "DROP",
        "IFT", "IFF", "IFE", "IFB", "IFNB", "IFIDN", "IFDIF", ".Z80", ".8080", ".RADIX", ".PHASE", ".DEPHASE",
        ".LIST", ".XLIST", "SEGMENT", "ENDS", "ASSUME", "PROC", "ENDP", "DD", "DQ",
        "BYTES", "WORDS", "IFNOT",
        "ENIF", "TITL", "SBTL", "COMM", "DL", "MESP", "MESV", "SHOW", "WAIT", "GET", "LOAD"}

    Private Const RxOpts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant
    ''' <summary>Dotted operators both EDAS IV and EDTASM-PLUS have.</summary>
    Private ReadOnly DotOpsShared As New Regex("\.(AND|OR|XOR|NOT|MOD)\.", RxOpts)
    ''' <summary>EDAS's compare operators (EDTASM-PLUS spells them .EQU. / .NEQ.).</summary>
    Private ReadOnly DotOpsEdasIV As New Regex("\.(EQ|NE)\.", RxOpts)
    Private ReadOnly DotOpsEdas43 As New Regex("\.(GE|GT|LE|LT|SHL|SHR|HIGH|LOW)\.", RxOpts)
    Private ReadOnly DotOpsEdtasmPlus As New Regex("\.(EQU|NEQ)\.", RxOpts)
    ''' <summary>A parenthesised sub-expression: "(" right after an operator. EDTASM-PLUS allows it; EDAS does not.</summary>
    Private ReadOnly ParenExpr As New Regex("[-+*/&!<]\s*\(", RxOpts)

    ' ------------------------------------------------------------------ entry point

    ''' <summary>Display name of a dialect for the viewer's info line.</summary>
    Public Function DisplayName(d As AsmDialect) As String
        Select Case d
            Case AsmDialect.Edas : Return "EDAS"
            Case AsmDialect.EdtasmPlus : Return "EDTASM-PLUS"
            Case AsmDialect.Edtasm : Return "EDTASM"
            Case AsmDialect.DiskEdtasm : Return "Disk EDTASM"
            Case AsmDialect.Mras : Return "MRAS"
            Case AsmDialect.Macro80 : Return "Macro-80"
            Case AsmDialect.Nedas : Return "NEDAS"
            Case AsmDialect.Zmac : Return "ZMAC"
            Case AsmDialect.Mzal : Return "M-ZAL"
            Case AsmDialect.Zeus : Return "Zeus"
            Case AsmDialect.Zen : Return "ZEN"
            Case AsmDialect.Alds : Return "ALDS"
            Case Else : Return d.ToString()
        End Select
    End Function

    ''' <summary>Fill info.DialectVersion / DialectEvidence and return the dialect, or Unidentified.</summary>
    Public Function Detect(lines As List(Of String), info As AsmSourceInfo) As AsmDialect
        If lines Is Nothing OrElse lines.Count = 0 Then Return AsmDialect.Unidentified
        Dim stmts As List(Of Stmt) = ParseAll(lines)
        If stmts.Count = 0 Then Return AsmDialect.Unidentified
        If IsForeignSource(stmts) Then Return AsmDialect.Unidentified

        Dim numbered As Boolean = info.IsNumbered
        ' A TED file was written by M-ZAL's own editor: only M-ZAL can claim it.
        If info.Layout = AsmSource.TedLayout Then
            Dim mz As Verdict = ScoreMzal(stmts)
            If mz.RuledOut Then
                info.Notes.Add("M-ZAL TED file, but the statements are not M-ZAL syntax")
                Return AsmDialect.Unidentified
            End If
            mz.Strong.Insert(0, "TED file layout")
            Return Claim(AsmDialect.Mzal, mz, info)
        End If
        Dim profiles As New List(Of KeyValuePair(Of AsmDialect, Verdict)) From {
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Edas, ScoreEdas(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Mras, ScoreEdas(stmts, mras:=True)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.EdtasmPlus, ScoreEdtasmPlus(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Macro80, ScoreMacro80(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Mzal, ScoreMzal(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Zeus, ScoreZeus(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Zen, ScoreZen(stmts)),
            New KeyValuePair(Of AsmDialect, Verdict)(AsmDialect.Alds, ScoreAlds(ParseAll(WithoutBlockComments(lines))))}

        Dim decided As List(Of KeyValuePair(Of AsmDialect, Verdict)) = profiles.Where(Function(p) p.Value.Decided(numbered)).ToList()
        ' Several decided: decisive evidence beats shared features.
        If decided.Count > 1 Then
            Dim withStrong = decided.Where(Function(p) p.Value.Strong.Count > 0).ToList()
            If withStrong.Count = 1 Then
                decided = withStrong
            Else
                info.Notes.Add("fits " & String.Join(" and ", decided.Select(Function(p) DisplayName(p.Key))) & " equally")
                Return AsmDialect.Unidentified
            End If
        End If
        If decided.Count = 1 Then Return Claim(decided(0).Key, decided(0).Value, info)

        ' Macro-80 is left out: colon-less EDTASM-style labels always rule it out.
        ' MRAS is claimed only on its own features, so its "ruled out" here means nothing.
        Dim numberedEditors = profiles.Where(Function(p) p.Key = AsmDialect.Edas OrElse p.Key = AsmDialect.EdtasmPlus)
        If numbered AndAlso numberedEditors.All(Function(p) Not p.Value.RuledOut AndAlso p.Value.Strong.Count + p.Value.Medium.Count = 0) Then
            info.Notes.Add("EDTASM vocabulary only (fits Series I EDTASM, EDTASM-PLUS, EDAS and the other numbered editors)")
        End If
        Return AsmDialect.Unidentified
    End Function

    Private Function Claim(d As AsmDialect, v As Verdict, info As AsmSourceInfo) As AsmDialect
        info.DialectVersion = v.Version
        For Each e In v.Strong.Concat(v.Medium)
            If info.DialectEvidence.Count < 4 Then info.DialectEvidence.Add(e)
        Next
        Return d
    End Function

    ' ------------------------------------------------------------------ shared checks

    ''' <summary>x86 registers in an operand (SP left out: the Z-80 has it too).</summary>
    Private ReadOnly X86Regs As New Regex("(?<![\w$@?.])(AX|BX|CX|DX|SI|DI|BP|AL|AH|BL|BH|CL|CH|DL|DH|ES|SS|CS|EAX|EBX|ECX|EDX)(?![\w$@?])", RxOpts)
    ''' <summary>MASM model/segment directives.</summary>
    ''' <summary>x86 mnemonics that are not also Z-80 spellings.</summary>
    Private ReadOnly X86Ops As New HashSet(Of String)(StringComparer.Ordinal) From {
        "MOV", "LEA", "CMP", "INT", "TEST", "SHL", "SHR", "MUL", "DIV", "LOOP", "JMP", "JE", "JNE", "JZ", "JNZ",
        "MOVSB", "MOVSW", "STOSB", "STOSW", "LODSB", "LODSW", "REP", "XCHG", "LES", "LDS"}
    Private ReadOnly X86Directives As New HashSet(Of String)(StringComparer.Ordinal) From {
        ".MODEL", ".DATA", ".CODE", ".STACK", ".CONST", ".186", ".286", ".386", ".486", ".8086", ".8087"}

    ''' <summary>
    ''' Not Z-80 source for any trained assembler: x86 MASM (registers, .MODEL/.DATA), or zasm/C-preprocessor
    ''' "#target"/"#include" lines. Nothing is claimed for such a file.
    ''' </summary>
    Private Function IsForeignSource(stmts As List(Of Stmt)) As Boolean
        Dim x86 As Integer = 0, hashLines As Integer = 0
        For Each s In stmts
            If X86Directives.Contains(s.Op) Then Return True
            If s.Label.StartsWith("#", StringComparison.Ordinal) Then hashLines += 1
            ' Only instructions count: a macro argument such as "DCE DI,'DIR'" (ALDS) is just a symbol.
            If s.Operand.Length > 0 AndAlso (Z80Ops.Contains(s.Op) OrElse X86Ops.Contains(s.Op)) AndAlso
               X86Regs.IsMatch(StripQuoted(s.Operand)) Then x86 += 1
        Next
        Return x86 >= 3 OrElse hashLines >= 2
    End Function

    ''' <summary>Macro names defined in the file (their invocations are not unknown opcodes).</summary>
    Private Function MacroNames(stmts As List(Of Stmt)) As HashSet(Of String)
        Dim macros As New HashSet(Of String)(StringComparer.Ordinal)
        For Each s In stmts
            If s.Op = "MACRO" AndAlso s.Label.Length > 0 Then macros.Add(s.Label.ToUpperInvariant())
        Next
        Return macros
    End Function

    ''' <summary>Opcodes whose label names a value, not an address: Macro-80 writes these without a colon.</summary>
    Private ReadOnly ValueOps As New HashSet(Of String)(StringComparer.Ordinal) From {"EQU", "SET", "DEFL", "ASET", "MACRO"}

    ''' <summary>Code labels (not EQU/SET/DEFL/MACRO names) counted with and without a trailing colon.</summary>
    Private Sub CountCodeLabels(stmts As List(Of Stmt), ByRef codeLabels As Integer, ByRef colons As Integer)
        codeLabels = 0 : colons = 0
        For Each s In stmts
            If s.Label.Length = 0 OrElse ValueOps.Contains(s.Op) Then Continue For
            codeLabels += 1
            If s.LabelColon Then colons += 1
        Next
    End Sub

    ''' <summary>"LABEL:" on at least half the code labels (and at least 3): EDAS and EDTASM-PLUS never allow the colon.</summary>
    Private Function ColonLabels(stmts As List(Of Stmt)) As Boolean
        Dim labels As Integer, colons As Integer
        CountCodeLabels(stmts, labels, colons)
        Return colons >= 3 AndAlso colons * 2 >= labels
    End Function

    ''' <summary>Unknown opcodes on more than 5% of the statements (and at least 3): another assembler's source.</summary>
    Private Function TooManyUnknown(unknown As Integer, instructions As Integer) As Boolean
        Return unknown >= 3 AndAlso unknown * 20 > instructions
    End Function

    ' ------------------------------------------------------------------ EDAS

    ''' <summary>MRAS (MISOSYS 1985) additions to EDAS 4.3: relocation and linking.</summary>
    Private ReadOnly MrasPseudo As New HashSet(Of String)(StringComparer.Ordinal) From {
        "ASEG", "CSEG", "DSEG", "COMMON", "GLOBAL", "PUBLIC", "EXT", "EXTRN", "NAME", "IFABS", "IFEXT", "IFNEXT", "IFREL"}

    ''' <summary>MRAS directives with the shortest abbreviation it accepts (manual: *Get *Inc *List *Mod *RAdix *REquest *Search).</summary>
    Private ReadOnly MrasDirectives As String()() = {
        New String() {"*GET", "*G"}, New String() {"*INCLUDE", "*I"}, New String() {"*LIST", "*L"}, New String() {"*MOD", "*M"},
        New String() {"*RADIX", "*RA"}, New String() {"*REQUEST", "*RE"}, New String() {"*SEARCH", "*S"}}

    ''' <summary>"name#" declares an external in MRAS (not "#parm", not EDTASM-PLUS's "#$YM").</summary>
    Private ReadOnly MrasExternRef As New Regex("[A-Za-z0-9@$_?]#(?![A-Za-z0-9$])", RxOpts)

    ''' <summary>The full MRAS directive an abbreviation stands for, or Nothing.</summary>
    Private Function MrasDirective(d As String) As String
        For Each pair In MrasDirectives
            If pair(0).StartsWith(d, StringComparison.Ordinal) AndAlso d.Length >= pair(1).Length Then Return pair(0)
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' EDAS, or with mras=True its relocating successor MRAS. MRAS accepts everything EDAS 4.3 does, so it is
    ''' claimed only on MRAS-specific evidence (relocation pseudo-ops, *RADIX/*REQUEST, abbreviated directives,
    ''' LABEL: / LABEL:: colons, name# externals); otherwise the file is EDAS's.
    ''' </summary>
    Private Function ScoreEdas(stmts As List(Of Stmt), Optional mras As Boolean = False) As Verdict
        ' MRAS needs decisive evidence: its colons, DB, ... are Macro-80's and ZMAC's habits too.
        Dim r As New Verdict() With {.MediumDecides = Not mras}
        Dim macros As HashSet(Of String) = MacroNames(stmts)
        Dim mrasOnly As Boolean = False
        If ColonLabels(stmts) Then
            If Not mras Then r.RuledOut = True : Return r
            r.AddMedium("LABEL: colons") : mrasOnly = True
        End If

        Dim unknown As Integer = 0, instructions As Integer = 0, parens As Integer = 0
        Dim v As String = ""
        For Each s In stmts
            If s.Directive.Length > 0 Then
                If mras Then
                    Dim full As String = MrasDirective(s.Directive)
                    If full Is Nothing Then r.RuledOut = True : Return r
                    If full = "*RADIX" OrElse full = "*REQUEST" OrElse full <> s.Directive Then
                        r.AddStrong(If(full = s.Directive, full, "abbreviated " & full)) : mrasOnly = True
                    ElseIf full <> "*LIST" Then
                        r.AddStrong(full)
                    End If
                    Continue For
                End If
                Dim dv As String = Nothing
                If EdasDirectives.TryGetValue(s.Directive, dv) Then
                    r.AddStrong(s.Directive) : v = MaxVersion(v, dv)
                ElseIf s.Directive <> "*LIST" Then
                    ' *MLIST, *INCL, *TITLE, ... belong to other assemblers.
                    r.RuledOut = True : Return r
                End If
                Continue For
            End If
            If mras AndAlso s.LabelPublic Then r.AddStrong("LABEL:: public labels") : mrasOnly = True
            ' EDTASM labels are 1-6 letters/digits; EDAS allows 15 and the characters $ _ ?.
            If s.Label.Length > 6 Then r.AddMedium("labels over 6 characters") : v = MaxVersion(v, "3.5")
            If s.Label.IndexOfAny({"$"c, "_"c, "?"c}) >= 0 Then r.AddMedium("$ _ ? in labels") : v = MaxVersion(v, "3.5")
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)

            If mras AndAlso MrasPseudo.Contains(op) Then
                r.AddStrong(op) : mrasOnly = True
                Continue For
            End If
            If ForeignPseudo.Contains(op) OrElse op = "COND" OrElse op = "ENDC" OrElse op = "INCLUDE" Then r.RuledOut = True : Return r
            If DotOpsEdtasmPlus.IsMatch(bare) OrElse bare.Contains("#$YM") Then r.RuledOut = True : Return r
            If mras Then
                ' Macro-80's word operators and 8080 mnemonics separate it from MRAS.
                If M80WordOps.IsMatch(" " & bare & " ") OrElse I8080Ops.Contains(op) Then r.RuledOut = True : Return r
                If MrasExternRef.IsMatch(bare) Then r.AddStrong("name# externals") : mrasOnly = True
            End If
            If ParenExpr.IsMatch(bare) Then parens += 1
            If DotOpsEdas43.IsMatch(bare) Then r.AddStrong(DotName(DotOpsEdas43, bare) & " operator") : v = MaxVersion(v, "4.3")
            If DotOpsEdasIV.IsMatch(bare) Then r.AddStrong(DotName(DotOpsEdasIV, bare) & " operator") : v = MaxVersion(v, "IV")
            If DotOpsShared.IsMatch(bare) Then r.AddMedium(DotName(DotOpsShared, bare) & " operator") : v = MaxVersion(v, "IV")

            If Z80Ops.Contains(op) OrElse EdtasmPseudo.Contains(op) Then
                If op = "DEFM" AndAlso TopLevelCommas(operand) > 0 Then r.AddMedium("DEFM with several operands") : v = MaxVersion(v, "3.5")
                Continue For
            End If
            If macros.Contains(op) Then Continue For

            Dim ver As String = Nothing
            If Not EdasPseudo.TryGetValue(op, ver) Then
                unknown += 1
                Continue For
            End If
            v = MaxVersion(v, ver)
            Select Case op
                Case "TITLE", "SUBTTL", "COM"
                    If operand.Contains("<"c) Then r.AddStrong(op & " '<…>'") Else r.AddMedium(op)
                Case "IFLT", "IFGT", "IFEQ", "IFNE"
                    ' EDAS's two-argument compare (IFxx exp1,exp2) has no counterpart elsewhere.
                    If TopLevelCommas(operand) = 1 Then r.AddStrong(op & " exp1,exp2") Else r.AddMedium(op)
                Case "IFLT$", "IFGT$", "IFEQ$", "IFNE$", "IFREF", "LORG", "DSYM", "DX"
                    r.AddStrong(op)
                Case "MACRO"
                    ' #parm dummies are EDTASM-PLUS's too; a default value (#LEN=255) is EDAS's alone.
                    If Regex.IsMatch(operand, "#[A-Za-z0-9@$_?]+\s*=") Then
                        r.AddStrong("MACRO #parm=default")
                    Else
                        r.AddMedium("MACRO")
                    End If
                Case "DC"
                    ' EDAS DC = count,value; Macro-80's DC is a string with bit 7 set on its last character.
                    If Not operand.TrimStart().StartsWith("'", StringComparison.Ordinal) AndAlso TopLevelCommas(operand) = 1 Then
                        r.AddStrong("DC count,value")
                    Else
                        unknown += 1
                    End If
                Case "IF", "ENDIF", "ELSE", "ENDM"
                    r.AddMedium("IF/ENDIF")
                Case Else
                    r.AddMedium(op)
            End Select
        Next

        ' "No parentheses are allowed" in an EDAS or MRAS expression (5th Ed. §3; MRAS §2, Expressions).
        If parens >= 2 OrElse TooManyUnknown(unknown, instructions) Then r.RuledOut = True
        If mras Then
            If Not mrasOnly Then r.RuledOut = True
        ElseIf v.Length > 0 Then
            r.Version = v & " or later"
        End If
        Return r
    End Function

    ' ------------------------------------------------------------------ EDTASM-PLUS

    Private Function ScoreEdtasmPlus(stmts As List(Of Stmt)) As Verdict
        Dim r As New Verdict()
        Dim macros As HashSet(Of String) = MacroNames(stmts)
        If ColonLabels(stmts) Then r.RuledOut = True : Return r

        Dim unknown As Integer = 0, instructions As Integer = 0, longLabels As Integer = 0
        Dim disk As Boolean = False
        For Each s In stmts
            If s.Directive.Length > 0 Then
                Select Case s.Directive
                    Case "*MLIST" : r.AddStrong("*MLIST")
                    Case "*LIST"
                    Case Else
                        ' *GET, *INCL, ... are other assemblers'; EDTASM-PLUS knows only *LIST and *MLIST.
                        r.RuledOut = True : Return r
                End Select
                Continue For
            End If
            ' Labels are 1-6 characters ("longer than six characters" is an EDTASM-PLUS error).
            If s.Label.Length > 6 Then longLabels += 1
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)

            If ForeignPseudo.Contains(op) Then r.RuledOut = True : Return r
            If DotOpsEdasIV.IsMatch(bare) OrElse DotOpsEdas43.IsMatch(bare) Then r.RuledOut = True : Return r
            If DotOpsEdtasmPlus.IsMatch(bare) Then r.AddStrong(DotName(DotOpsEdtasmPlus, bare) & " operator")
            If bare.Contains("#$YM") Then r.AddStrong("#$YM macro labels")
            If DotOpsShared.IsMatch(bare) Then r.AddMedium(DotName(DotOpsShared, bare) & " operator")
            If ParenExpr.IsMatch(bare) Then r.AddMedium("parenthesised expressions")

            If Z80Ops.Contains(op) OrElse EdtasmPseudo.Contains(op) Then
                ' DEFM takes one 'string' in EDTASM-PLUS; EDAS-style 'text',CR is a giveaway.
                If op = "DEFM" AndAlso TopLevelCommas(operand) > 0 Then r.RuledOut = True : Return r
                Continue For
            End If
            If macros.Contains(op) Then Continue For

            Dim ver As String = Nothing
            If Not EdtasmPlusPseudo.TryGetValue(op, ver) Then
                unknown += 1
                Continue For
            End If
            If ver = "disk" Then disk = True
            Select Case op
                Case "COND", "ENDC" : r.AddStrong("COND/ENDC")
                Case "MACRO", "ENDM" : r.AddMedium("MACRO")
                Case "TITLE"
                    ' EDAS wants TITLE '<...>'; EDTASM-PLUS takes the string as it stands.
                    If operand.Contains("<"c) Then r.RuledOut = True : Return r
                    r.AddMedium(op)
                Case Else : r.AddMedium(op)
            End Select
        Next

        If longLabels >= 2 OrElse TooManyUnknown(unknown, instructions) Then r.RuledOut = True
        r.Version = If(disk, "disk version", "")
        Return r
    End Function

    ' ------------------------------------------------------------------ Macro-80

    ''' <summary>Macro-80 pseudo-ops (manual §2.6–2.8 and the keyword table in M80/CMD), including the
    ''' "Z80 pseudo-ops" it accepts as aliases (COND/ENDC, DEFx, GLOBAL, EXTERNAL).</summary>
    Private ReadOnly M80Pseudo As New HashSet(Of String)(StringComparer.Ordinal) From {
        "ASEG", "CSEG", "DSEG", "COMMON", "DB", "DC", "DS", "DW", "ENTRY", "PUBLIC", "EXT", "EXTRN", "INCLUDE",
        "$INCLUDE", "MACLIB", "NAME", "PAGE", "SET", "SUBTTL", "TITLE", "EJECT", ".COMMENT", ".PRINTX", "PRINTX",
        ".RADIX", ".Z80", ".8080", ".REQUEST", "IF", "IFT", "IFE", "IFF", "IF1", "IF2", "IFDEF", "IFNDEF", "IFB",
        "IFNB", "IFIDN", "IFDIF", "ELSE", "ENDIF", ".LIST", ".XLIST", ".CREF", ".XCREF", ".LALL", ".SALL", ".XALL",
        ".PHASE", ".DEPHASE", "REPT", "IRP", "IRPC", "MACRO", "ENDM", "EXITM", "LOCAL",
        "COND", "ENDC", "GLOBAL", "EXTERNAL"}

    ''' <summary>Macro-80 pseudo-ops no other trained dialect has: one is decisive.</summary>
    Private ReadOnly M80Decisive As New HashSet(Of String)(StringComparer.Ordinal) From {
        "ASEG", "CSEG", "DSEG", "COMMON", "PUBLIC", "EXT", "EXTRN", "EXTERNAL", "NAME", "LOCAL", "$INCLUDE",
        "MACLIB", ".COMMENT", ".PRINTX", "PRINTX", ".RADIX", ".Z80", ".8080", ".REQUEST", "IFT", "IFE", "IFF", "IFB",
        "IFNB", "IFIDN", "IFDIF", ".LIST", ".XLIST", ".CREF", ".XCREF", ".LALL", ".SALL", ".XALL", ".PHASE", ".DEPHASE"}

    ''' <summary>Intel 8080 mnemonics M80 accepts in .8080 mode (only those that are not also Z-80 spellings).</summary>
    Private ReadOnly I8080Ops As New HashSet(Of String)(StringComparer.Ordinal) From {
        "MOV", "MVI", "LXI", "LDA", "STA", "LHLD", "SHLD", "LDAX", "STAX", "XCHG", "ADI", "ACI", "SUI", "SBI", "ANA",
        "XRA", "ORA", "CMP", "ANI", "XRI", "ORI", "INR", "DCR", "INX", "DCX", "DAD", "RAL", "RAR", "CMA", "STC",
        "CMC", "JMP", "JNZ", "JZ", "JNC", "JC", "JPO", "JPE", "JM", "CNZ", "CZ", "CNC", "CC", "CPO", "CPE", "CM",
        "RNZ", "RZ", "RNC", "RC", "RPO", "RPE", "RM", "PCHL", "SPHL", "XTHL", "HLT"}

    ''' <summary>Macro-80's word operators, which must be separated from their operands by blanks (§2.4.1).
    ''' A binary one needs an operand on each side, a unary one an operand after it — "DEFB NOT" or "LD HL,LOW"
    ''' use NOT / LOW as symbol names (ALDS TRSDOS sources), not operators.</summary>
    Private ReadOnly M80WordOps As New Regex("[\w$')]\s+(AND|OR|XOR|MOD|SHR|SHL|EQ|NE|LT|LE|GT|GE)\s+[\w$'(]|(^|[\s,(])(NOT|HIGH|LOW|NUL)\s+[\w$'(]", RxOpts)

    ''' <summary>Pseudo-ops of the Tandy relocating assembler that Macro-80 lacks.</summary>
    Private ReadOnly NotM80 As New HashSet(Of String)(StringComparer.Ordinal) From {
        "PSECT", "HEADER", "NOLOAD", "DEFT", "DEFE", "USING", "DROP", "SEGMENT", "ENDS", "ASSUME", "PROC", "ENDP", "DD", "DQ",
        "BYTES", "WORDS", "IFNOT",
        "ENIF", "TITL", "SBTL", "COMM", "DL", "MESP", "MESV", "SHOW", "WAIT", "GET", "LOAD"}

    Private Function ScoreMacro80(stmts As List(Of Stmt)) As Verdict
        Dim r As New Verdict()
        Dim macros As HashSet(Of String) = MacroNames(stmts)
        ' Every Macro-80 code label ends in ":" (§2.3.1).
        Dim codeLabels As Integer, colons As Integer
        CountCodeLabels(stmts, codeLabels, colons)
        r.MediumDecides = False
        If codeLabels >= 3 AndAlso colons * 2 < codeLabels Then r.RuledOut = True : Return r
        If colons >= 3 Then r.AddMedium("LABEL: colons")

        Dim unknown As Integer = 0, instructions As Integer = 0
        For Each s In stmts
            If s.Directive.Length > 0 Then
                ' *EJECT is a Macro-80 "Z80 pseudo-op"; *LIST, *GET, *MLIST belong to the EDTASM family.
                If s.Directive = "*EJECT" Then r.AddStrong("*EJECT") Else r.RuledOut = True : Return r
                Continue For
            End If
            If s.LabelPublic Then r.AddStrong("LABEL:: public labels")
            If s.Label.StartsWith("$EJECT", StringComparison.OrdinalIgnoreCase) OrElse
               s.Label.StartsWith("$TITLE", StringComparison.OrdinalIgnoreCase) Then r.AddStrong("ISIS $ controls") : Continue For
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)

            If NotM80.Contains(op) Then r.RuledOut = True : Return r
            ' A dot pseudo-op Macro-80 does not know (.TITLE, .MODEL, ...) is another assembler's.
            If op.StartsWith(".", StringComparison.Ordinal) AndAlso Not M80Pseudo.Contains(op) Then r.RuledOut = True : Return r
            If DotOpsShared.IsMatch(bare) OrElse DotOpsEdasIV.IsMatch(bare) OrElse DotOpsEdas43.IsMatch(bare) OrElse
               DotOpsEdtasmPlus.IsMatch(bare) Then r.RuledOut = True : Return r
            If M80WordOps.IsMatch(" " & bare & " ") Then r.AddStrong("word operators (AND, SHR, HIGH ...)")
            If operand.Contains(""""c) Then r.AddMedium("double-quoted strings")

            If I8080Ops.Contains(op) Then r.AddStrong("8080 mnemonics") : Continue For
            If Z80Ops.Contains(op) Then
                ' Z-80 SET b,r is an instruction; "name SET value" (no comma) is Macro-80's redefinable equate.
                If op = "SET" AndAlso s.Label.Length > 0 AndAlso TopLevelCommas(operand) = 0 Then r.AddMedium("SET")
                Continue For
            End If
            If EdtasmPseudo.Contains(op) Then Continue For
            If macros.Contains(op) Then Continue For
            If Not M80Pseudo.Contains(op) Then
                unknown += 1
                Continue For
            End If
            Select Case op
                Case "MACRO"
                    ' "#" is not a symbol character in Macro-80: #parm dummies are EDAS / EDTASM-PLUS.
                    If operand.Contains("#"c) Then r.RuledOut = True : Return r
                    r.AddMedium("MACRO")
                Case "TITLE", "SUBTTL"
                    If operand.Contains("<"c) Then r.RuledOut = True : Return r
                    r.AddMedium(op)
                Case "DC"
                    ' Macro-80's DC is a string with bit 7 set on its last character.
                    If operand.TrimStart().StartsWith("'", StringComparison.Ordinal) OrElse operand.TrimStart().StartsWith("""", StringComparison.Ordinal) Then
                        r.AddStrong("DC 'string'")
                    Else
                        r.RuledOut = True : Return r
                    End If
                Case "IF", "ELSE", "ENDIF", "ENDM" : r.AddMedium("IF/ENDIF")
                Case Else
                    If M80Decisive.Contains(op) Then r.AddStrong(op) Else r.AddMedium(op)
            End Select
        Next

        If TooManyUnknown(unknown, instructions) Then r.RuledOut = True
        Return r
    End Function

    ' ------------------------------------------------------------------ M-ZAL

    ''' <summary>M-ZAL pseudo-ops beyond EDTASM (System Manual §VI, §XI). "R2" = added in Release 2.</summary>
    Private ReadOnly MzalPseudo As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"DB", ""}, {"DW", ""}, {"DS", ""}, {"BYTES", ""}, {"WORDS", ""}, {"ENTRY", ""}, {"EXTRN", ""},
        {"IF", "R2"}, {"IFNOT", "R2"}, {"ENDIF", "R2"}, {"MACRO", "R2"}, {"ENDM", "R2"}}

    ''' <summary>M-ZAL *commands (column 1). *LIST is EDTASM's; *EJECT is Macro-80's too.</summary>
    Private ReadOnly MzalDirectives As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"*INCL", ""}, {"*SPACE", ""}, {"*TITLE", ""}, {"*PAUSE", ""}, {"*COM", ""}, {"*MACLIST", "R2"}}

    ''' <summary>Special symbolic parameters of M-ZAL macros.</summary>
    Private ReadOnly MzalSpecialParms As New Regex("\?(INDEX|PARM#|TIME|DATE)", RxOpts)

    Private Function ScoreMzal(stmts As List(Of Stmt)) As Verdict
        ' MASM reads TED files; an EDTASM-layout file is no evidence for M-ZAL.
        Dim r As New Verdict() With {.MediumDecides = False}
        Dim macros As HashSet(Of String) = MacroNames(stmts)
        If ColonLabels(stmts) Then r.RuledOut = True : Return r

        Dim unknown As Integer = 0, instructions As Integer = 0, longLabels As Integer = 0
        Dim r2 As Boolean = False
        For Each s In stmts
            If s.Directive.Length > 0 Then
                Dim dv As String = Nothing
                If MzalDirectives.TryGetValue(s.Directive, dv) Then
                    r.AddStrong(s.Directive)
                    If dv = "R2" Then r2 = True
                ElseIf s.Directive = "*EJECT" Then
                    r.AddMedium("*EJECT")
                ElseIf s.Directive <> "*LIST" Then
                    r.RuledOut = True : Return r
                End If
                Continue For
            End If
            ' Labels: 1-8 characters; letters, digits and @ # $ % _ (first a letter, $ or @).
            If s.Label.Length > 8 Then longLabels += 1
            If s.Label.Contains("%"c) Then r.AddStrong("% in labels")
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)

            ' Only + - * / & < exist; dotted and word operators are other assemblers'.
            If DotOpsShared.IsMatch(bare) OrElse DotOpsEdasIV.IsMatch(bare) OrElse DotOpsEdas43.IsMatch(bare) OrElse
               DotOpsEdtasmPlus.IsMatch(bare) OrElse M80WordOps.IsMatch(" " & bare & " ") Then r.RuledOut = True : Return r
            If MzalSpecialParms.IsMatch(operand) OrElse MzalSpecialParms.IsMatch(s.Label) Then r.AddStrong("?INDEX/?PARM# macro parameters") : r2 = True

            If Z80Ops.Contains(op) OrElse EdtasmPseudo.Contains(op) Then
                ' DEFM takes one 'string'.
                If op = "DEFM" AndAlso TopLevelCommas(operand) > 0 Then r.RuledOut = True : Return r
                Continue For
            End If
            If macros.Contains(op) Then Continue For
            Dim ver As String = Nothing
            If Not MzalPseudo.TryGetValue(op, ver) Then
                unknown += 1
                Continue For
            End If
            If ver = "R2" Then r2 = True
            Select Case op
                Case "BYTES", "WORDS", "IFNOT" : r.AddStrong(op)
                Case "MACRO"
                    ' Symbolic parameters start with "?" (EDAS / EDTASM-PLUS use "#").
                    If operand.Contains("#"c) Then r.RuledOut = True : Return r
                    If operand.TrimStart().StartsWith("?", StringComparison.Ordinal) Then r.AddStrong("MACRO ?parm") Else r.AddMedium("MACRO")
                Case "IF", "ENDIF", "ENDM" : r.AddMedium("IF/ENDIF")
                Case Else : r.AddMedium(op)
            End Select
        Next

        If longLabels >= 2 OrElse TooManyUnknown(unknown, instructions) Then r.RuledOut = True
        r.Version = If(r2, "Release 2", "")
        Return r
    End Function

    ' ------------------------------------------------------------------ Zeus

    ''' <summary>Zeus pseudo-ops beyond EDTASM (1983 manual §2; "9th" = added by the 9th edition, 2010).</summary>
    Private ReadOnly ZeusPseudo As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {"DB", ""}, {"DL", ""}, {"DM", ""}, {"DS", ""}, {"DW", ""}, {"COMM", ""}, {"ENIF", ""}, {"ERR", ""},
        {"GET", ""}, {"IF", ""}, {"LIST", ""}, {"PAGE", ""}, {"SBTL", ""}, {"TITL", ""},
        {"MESP", "9th"}, {"MESV", "9th"}, {"SHOW", "9th"}, {"WAIT", "9th"}}

    ''' <summary>Zeus pseudo-ops no other known assembler spells this way.</summary>
    Private ReadOnly ZeusDecisive As New HashSet(Of String)(StringComparer.Ordinal) From {
        "DL", "COMM", "ENIF", "SBTL", "TITL", "MESP", "MESV", "SHOW", "WAIT"}

    ''' <summary>Zeus's XOR operator "#" between two terms (MRAS's "name#" is followed by nothing or a comma).</summary>
    Private ReadOnly ZeusXor As New Regex("[A-Za-z0-9$')]#[A-Za-z0-9$'(]", RxOpts)

    Private ReadOnly ZeusLabel As New Regex("^[@A-Z\[\\\]^_`a-z{|}~][@A-Z\[\\\]^_`a-z{|}~0-9:=>?]*$", RegexOptions.CultureInvariant)

    Private Function ScoreZeus(stmts As List(Of Stmt)) As Verdict
        Dim r As New Verdict()
        If ColonLabels(stmts) Then r.RuledOut = True : Return r
        Dim unknown As Integer = 0, instructions As Integer = 0
        Dim ninth As Boolean = False
        For Each s In stmts
            ' Zeus has no * directives; GET and LIST are pseudo-ops.
            If s.Directive.Length > 0 Then r.RuledOut = True : Return r
            ' Label characters: @ A-Z [ \ ] ^ _ ` a-z { | } ~, then also 0-9 : = > ? — no $ % # . (manual §2).
            If s.Label.Length > 0 AndAlso Not ZeusLabel.IsMatch(s.Label) Then r.RuledOut = True : Return r
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)
            If DotOpsShared.IsMatch(bare) OrElse DotOpsEdasIV.IsMatch(bare) OrElse DotOpsEdas43.IsMatch(bare) OrElse
               DotOpsEdtasmPlus.IsMatch(bare) OrElse M80WordOps.IsMatch(" " & bare & " ") Then r.RuledOut = True : Return r
            If ZeusXor.IsMatch(bare) Then r.AddStrong("# XOR operator")

            If Z80Ops.Contains(op) OrElse EdtasmPseudo.Contains(op) Then
                If op = "DEFM" AndAlso TopLevelCommas(operand) > 0 Then r.AddMedium("DEFM with several operands")
                Continue For
            End If
            Dim ver As String = Nothing
            If Not ZeusPseudo.TryGetValue(op, ver) Then
                ' No macros, no ENDIF, no relocation: anything else is another assembler's.
                unknown += 1
                Continue For
            End If
            If ver = "9th" Then ninth = True
            If ZeusDecisive.Contains(op) Then
                r.AddStrong(op)
            ElseIf op = "IF" AndAlso operand.StartsWith("NOT,", StringComparison.OrdinalIgnoreCase) Then
                r.AddStrong("IF NOT,exp")
            ElseIf op = "GET" OrElse op = "LIST" Then
                r.AddStrong(op & " without *")
            Else
                r.AddMedium(op)
            End If
        Next
        If TooManyUnknown(unknown, instructions) Then r.RuledOut = True
        r.Version = If(ninth, "9th edition (2010) features", "")
        Return r
    End Function

    ' ------------------------------------------------------------------ ZEN

    ''' <summary>ZEN's only pseudo-ops beyond the Z-80 set (Avalon Software manual, 1979).</summary>
    Private ReadOnly ZenPseudo As New HashSet(Of String)(StringComparer.Ordinal) From {"ORG", "EQU", "END", "DB", "DW", "DS", "LOAD"}

    Private Function ScoreZen(stmts As List(Of Stmt)) As Verdict
        ' ZEN never decides on shared features: its DB/DW/DS and colons are common.
        Dim r As New Verdict() With {.MediumDecides = False}
        Dim labels As Integer = 0, colons As Integer = 0, equColons As Integer = 0
        For Each s In stmts
            If s.Label.Length = 0 Then Continue For
            labels += 1
            If s.LabelColon Then
                colons += 1
                If s.Op = "EQU" Then equColons += 1
            End If
        Next
        ' "A symbol must be postfixed with a colon when declared" — EQU names included.
        If labels >= 3 AndAlso colons * 4 < labels * 3 Then r.RuledOut = True : Return r
        If equColons >= 2 Then r.AddMedium("NAME: EQU (colon on equates)")

        Dim unknown As Integer = 0, instructions As Integer = 0
        For Each s In stmts
            If s.Directive.Length > 0 Then r.RuledOut = True : Return r
            If s.Op.Length = 0 Then Continue For
            instructions += 1
            If Z80Ops.Contains(s.Op) Then Continue For
            If Not ZenPseudo.Contains(s.Op) Then
                ' DEFB, DEFM, IF, MACRO, ... are not ZEN's.
                unknown += 1
                Continue For
            End If
            Select Case s.Op
                Case "LOAD" : r.AddStrong("LOAD")
                Case "DB", "DW", "DS" : r.AddMedium(s.Op)
            End Select
        Next
        If unknown > 0 Then r.RuledOut = True
        Return r
    End Function

    ' ------------------------------------------------------------------ ALDS

    ''' <summary>Tandy ALDS (26-2012, v3.02 1983) directives and extended mnemonics no other known assembler has.</summary>
    Private ReadOnly AldsDecisive As New HashSet(Of String)(StringComparer.Ordinal) From {
        "PSECT", "HEADER", "DEFT", "DEFE", "DEFR", "NOLOAD", "NOEND", "NOFILL", "ISECT", "LITORG", "ADISP", "APOP",
        "APUSH", "GLINK", "LINK", "OBJ", "PATCH", "QUIT", "STOP", "RESLOC", "SETLOC", "USING", "DROP", "VERSION",
        "IFM", "IFP", "IFNZ", "IFZ", "IFUND", "ENDI", "EXTERN",
        "CPR", "CMPD", "CMPI", "MOVD", "MOVI", "SAVE", "RSTR", "TZ", "SVC"}

    ''' <summary>ALDS directives shared with other assemblers.</summary>
    Private ReadOnly AldsShared As New HashSet(Of String)(StringComparer.Ordinal) From {
        "EXT", "GLOBAL", "PUBLIC", "INCLUDE", "MACRO", "ENDM", "IFDEF", "IFT", "IFF", "IF", "ELSE", "ENDIF",
        "TITLE", "EJECT", "RADIX", "DATE", "TIME", "REF", "FILL", "PRINT"}

    ''' <summary>ALDS operators no other known assembler has.</summary>
    Private ReadOnly AldsDotOps As New Regex("\.(BIT|RR|RL|ABS|RES|SGN|UGT|UGE|ULT|ULE|MSB|LSB)\.", RxOpts)

    ''' <summary>Other assemblers' tokens ALDS never accepts.</summary>
    Private ReadOnly NotAlds As New HashSet(Of String)(StringComparer.Ordinal) From {
        "COND", "ENDC", "BYTES", "WORDS", "IFNOT", "ENIF", "TITL", "SBTL", "COMM", "DL", "LOAD", "ASEG", "CSEG",
        "DSEG", "COMMON", "LORG", "IFLT", "IFGT", "IFEQ", "IFNE", "IFREF", "DSYM", "DX", "SEGMENT", "ENDS", "ASSUME",
        "PROC", "ENDP", "DD", "DQ"}

    ''' <summary>The source lines without ALDS block comments: a line starting with "*" turns them on, the next one off.</summary>
    Private Function WithoutBlockComments(lines As List(Of String)) As List(Of String)
        Dim kept As New List(Of String)
        Dim inBlock As Boolean = False
        For Each line In lines
            If line.StartsWith("*", StringComparison.Ordinal) Then
                inBlock = Not inBlock
            ElseIf Not inBlock Then
                kept.Add(line)
            End If
        Next
        Return kept
    End Function

    ''' <summary>
    ''' Tandy ALDS. Macro calls defined in INCLUDE files are common (the TRSDOS sources use dozens), so unknown
    ''' opcodes never rule it out; only other assemblers' known tokens do. Never decided on shared features.
    ''' </summary>
    Private Function ScoreAlds(stmts As List(Of Stmt)) As Verdict
        Dim r As New Verdict() With {.MediumDecides = False}
        Dim codeLabels As Integer, colons As Integer
        CountCodeLabels(stmts, codeLabels, colons)
        If colons >= 3 Then r.AddMedium("LABEL: colons")
        For Each s In stmts
            ' "*" lines were block comments; anything left as a directive means another assembler.
            If s.Directive.Length > 0 Then r.RuledOut = True : Return r
            If s.Op.Length = 0 Then Continue For
            Dim op As String = s.Op
            Dim operand As String = s.Operand
            Dim bare As String = StripQuoted(operand)
            If NotAlds.Contains(op) OrElse I8080Ops.Contains(op) OrElse (op.StartsWith(".", StringComparison.Ordinal)) Then r.RuledOut = True : Return r
            If DotOpsEdtasmPlus.IsMatch(bare) OrElse M80WordOps.IsMatch(" " & bare & " ") Then r.RuledOut = True : Return r
            If AldsDotOps.IsMatch(bare) Then r.AddStrong(DotName(AldsDotOps, bare) & " operator")
            If DotOpsShared.IsMatch(bare) OrElse DotOpsEdasIV.IsMatch(bare) OrElse DotOpsEdas43.IsMatch(bare) Then r.AddMedium("dotted operators")
            If AldsDecisive.Contains(op) Then
                r.AddStrong(op)
            ElseIf op = "EXT" AndAlso operand.TrimStart().StartsWith("'", StringComparison.Ordinal) Then
                r.AddStrong("EXT 'file/GBL'")
            ElseIf AldsShared.Contains(op) Then
                r.AddMedium(op)
            End If
        Next
        Return r
    End Function

    ' ------------------------------------------------------------------ helpers

    Private Function ParseAll(lines As List(Of String)) As List(Of Stmt)
        Dim stmts As New List(Of Stmt)
        For Each line In lines
            Dim s As New Stmt()
            If TryParse(line, s) Then stmts.Add(s)
        Next
        Return stmts
    End Function

    ''' <summary>Releases in order 3.5 &lt; IV &lt; 4.3; blank sorts lowest.</summary>
    Private Function MaxVersion(a As String, b As String) As String
        Dim order As String() = {"", "3.5", "IV", "4.3"}
        Return If(Array.IndexOf(order, b) > Array.IndexOf(order, a), b, a)
    End Function

    Private Function DotName(rx As Regex, text As String) As String
        Return "." & rx.Match(text).Groups(1).Value.ToUpperInvariant() & "."
    End Function

    ''' <summary>The operand with every quoted string removed, so text inside quotes is never taken for syntax.</summary>
    Private Function StripQuoted(operand As String) As String
        Return Regex.Replace(operand, "'[^']*'?|""[^""]*""?", "''")
    End Function

    ''' <summary>Split one line into label / opcode / operand. False for blank and comment-only lines.</summary>
    Private Function TryParse(line As String, ByRef s As Stmt) As Boolean
        s = New Stmt With {.Label = "", .Op = "", .Operand = "", .Directive = ""}
        If String.IsNullOrWhiteSpace(line) Then Return False
        Dim c0 As Char = line(0)
        If c0 = ";"c Then Return False
        If c0 = "*"c Then
            ' *GET, *LIST, *MLIST, ... — any other "*" line is a comment to the assemblers that allow it.
            Dim word As String = FirstToken(line, 0)
            If word.Length > 1 AndAlso Char.IsLetter(word(1)) Then
                s.Directive = word.ToUpperInvariant()
                Return True
            End If
            Return False
        End If

        Dim pos As Integer = 0
        If Not Char.IsWhiteSpace(c0) Then
            Dim tok As String = FirstToken(line, 0)
            ' ZEN glues the opcode to the label: "NEWLINE:EQU 13".
            Dim inner As Integer = tok.IndexOf(":"c)
            If inner > 0 AndAlso inner < tok.Length - 1 AndAlso tok(inner + 1) <> ":"c Then tok = tok.Substring(0, inner + 1)
            s.LabelColon = tok.EndsWith(":", StringComparison.Ordinal)
            s.LabelPublic = tok.EndsWith("::", StringComparison.Ordinal)
            s.Label = tok.TrimEnd(":"c)
            pos = tok.Length
        End If
        pos = SkipSpace(line, pos)
        If pos >= line.Length OrElse line(pos) = ";"c Then Return s.Label.Length > 0
        Dim op As String = FirstToken(line, pos)
        ' Macro-80 statements need not start in column 1: an indented "NAME:" is still a label.
        If s.Label.Length = 0 AndAlso op.Length > 1 AndAlso op.EndsWith(":", StringComparison.Ordinal) Then
            s.LabelColon = True
            s.LabelPublic = op.EndsWith("::", StringComparison.Ordinal)
            s.Label = op.TrimEnd(":"c)
            pos = SkipSpace(line, pos + op.Length)
            If pos >= line.Length OrElse line(pos) = ";"c Then Return True
            op = FirstToken(line, pos)
        End If
        s.Op = op.ToUpperInvariant()
        pos = SkipSpace(line, pos + op.Length)
        s.Operand = StripComment(line.Substring(Math.Min(pos, line.Length)), s.Op = "EX")
        Return True
    End Function

    Private Function FirstToken(line As String, start As Integer) As String
        Dim e As Integer = start
        ' A quote after the first character ends a label/opcode token too: ZEN writes "MESSAGE:DB'TEXT'".
        While e < line.Length AndAlso Not Char.IsWhiteSpace(line(e)) AndAlso line(e) <> ";"c AndAlso
              Not (e > start AndAlso (line(e) = "'"c OrElse line(e) = """"c))
            e += 1
        End While
        Return line.Substring(start, e - start)
    End Function

    Private Function SkipSpace(line As String, pos As Integer) As Integer
        While pos < line.Length AndAlso Char.IsWhiteSpace(line(pos))
            pos += 1
        End While
        Return pos
    End Function

    ''' <summary>Operand text up to a ";" outside quotes (EX AF,AF' has a lone quote, so EX skips quote tracking).</summary>
    Private Function StripComment(text As String, noQuotes As Boolean) As String
        Dim inQuote As Boolean = False
        For i As Integer = 0 To text.Length - 1
            Dim ch As Char = text(i)
            If ch = "'"c AndAlso Not noQuotes Then
                inQuote = Not inQuote
            ElseIf ch = ";"c AndAlso Not inQuote Then
                Return text.Substring(0, i).TrimEnd()
            End If
        Next
        Return text.TrimEnd()
    End Function

    ''' <summary>
    ''' Commas outside quotes and &lt;…&gt; lists. A "&lt;" opens a list only at the start of an argument;
    ''' anywhere else it is the shift operator.
    ''' </summary>
    Private Function TopLevelCommas(operand As String) As Integer
        Dim n As Integer = 0
        Dim inQuote As Boolean = False
        Dim angle As Integer = 0
        Dim argStart As Boolean = True
        For Each ch In operand
            Select Case ch
                Case "'"c : inQuote = Not inQuote
                Case "<"c : If Not inQuote AndAlso argStart Then angle += 1
                Case ">"c : If Not inQuote AndAlso angle > 0 Then angle -= 1
                Case ","c : If Not inQuote AndAlso angle = 0 Then n += 1
            End Select
            If Not Char.IsWhiteSpace(ch) Then argStart = (ch = ","c AndAlso Not inQuote AndAlso angle = 0)
        Next
        Return n
    End Function
End Module
