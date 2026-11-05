Option Compare Binary
Option Explicit

' ANCHOR (Access edition) - the naval format check for drafted messages
' (NTP 3(J) GENADMIN layout), brought over unchanged from Message
' Generator's reviewed check (modMGFormat) when ANCHOR took over drafting and
' release (Stephen, 2026-11-05). Only the names are ANCHOR's:
' NavalFormatIssues and FormatIssuesOf.
'
' Sources: ACP 121(G) (Oct 2004) paragraphs as cited, NAVEDTRA 14226 (Oct
' 1997) and the NAV26237 sample, as summarised in design-system/
' message-generator-drafting-rules.md. Two warnings come from the DON
' Correspondence Manual, SECNAV M-5216.5. NTP 3(J) (Jul 1997) rules are
' cited as "NTP 3"; where the 2026 NAV26237 sample departs from them the check
' only warns (design-system/message-generator-ntp3-notes.md). This is not a
' USMTF or MIL-STD-6040 validator: the set syntax (MSGID/, SUBJ/, REF/, NARR/,
' RMKS/, the / and // delimiters) is checked only as far as the NAV26237
' sample shows it.
'
' Issues are returned one per line, each ending in vbCrLf. Lines starting
' "Warning:" are advisory: FormatIssuesOf(issues, False) leaves them out, so
' they never stop a submit or a release (modDrafting); the Drafting page
' shows them.
'
' Option Compare Binary: the check compares set keywords and letters by
' case, as it did in Message Generator.
'
' Owner: drafting and release. Depends on: nothing.

' The 69-character limit is NTP 3 Annex A102b. The NAV26237 sample (2026) has
' lines longer than 69, so it is reported only as a warning.
Private Const NAVAL_LINE_LIMIT As Long = 69
' NTP 3 603: a PLA in the FM, TO or INFO line is limited to 55 characters.
Private Const PLA_LINE_LIMIT As Long = 55
' NTP 3 Annex A, figure A-2: the first SUBJ line holds 64 characters of subject.
Private Const SUBJ_FIELD_LIMIT As Long = 64
' NTP 3 Annex A, figure A-2: the MSGID originator field holds 1-20 characters.
Private Const MSGID_ORIG_LIMIT As Long = 20

' Every blocking issue and warning in a drafted message's text, one per
' line; "" when the text has no MSGID/ set (it isn't a naval message).
'   text   the message as it transmits (modDrafting.RenderDraftText)
' Errors: none.
Public Function NavalFormatIssues(ByVal text As String) As String
    Dim lines As Variant, ub As Long, i As Long, j As Long, issues As String
    Dim msgidLine As Long, btOpen As Long, btClose As Long, headerEnd As Long
    Dim textStart As Long, textEnd As Long, key As String
    Dim topClass As Long, precLine As Long, dtgLine As Long, fmLine As Long, toLine As Long, infoLine As Long
    Dim precLetter As String, words As Variant, k As Long, prosign As String
    Dim hasAddress As Boolean, firstText As Long, lastLine As Long
    Dim textClass As String, topValue As String, bottomValue As String
    Dim setName As String, curSet As String, curStart As Long, rank As Long, lastRank As Long
    Dim msgidCount As Long, subjCount As Long, refCount As Long, narrCount As Long, rmksCount As Long
    Dim rmksStart As Long, rmksEnd As Long, expectedLetter As String, isMarker As Boolean
    Dim counters(1 To 4) As Long, level As Long, number As Long, s As String
    Dim groupLine(1 To 4) As Long
    Dim ampnCount As Long, ampnLine As Long, narrLine As Long, pocCount As Long, needsExplain As Boolean
    Dim refType As String, p As Long, isGenadmin As Boolean, subjText As String, msgidOrig As String

    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ub = UBound(lines)
    msgidLine = -1
    For i = 0 To ub
        If Left$(LineKey(lines(i)), 6) = "MSGID/" Then msgidLine = i: Exit For
    Next
    ' Only text that looks like a naval message (it has an MSGID/ set) is checked.
    If msgidLine < 0 Then Exit Function
    isGenadmin = (Left$(LineKey(lines(msgidLine)), 15) = "MSGID/GENADMIN/")

    btOpen = -1: btClose = -1
    For i = 0 To ub
        If LineKey(lines(i)) = "BT" Then
            If i < msgidLine And btOpen < 0 Then btOpen = i
            If i > msgidLine Then btClose = i
        End If
    Next

    ' Heading: classification, precedence, DTG, FM, TO, INFO (lines before BT).
    If btOpen >= 0 Then headerEnd = btOpen Else headerEnd = msgidLine
    topClass = -1: precLine = -1: dtgLine = -1: fmLine = -1: toLine = -1: infoLine = -1
    For i = 0 To headerEnd - 1
        key = LineKey(lines(i))
        If Len(key) = 0 Then
            ' Blank lines are ignored.
        ElseIf Left$(key, 15) = "CLASSIFICATION:" Then
            If topClass < 0 Then topClass = i
        ElseIf FirstWord(key) = "FM" Then
            If fmLine < 0 Then fmLine = i
        ElseIf FirstWord(key) = "TO" Then
            If toLine < 0 Then toLine = i
        ElseIf FirstWord(key) = "INFO" Then
            If infoLine < 0 Then infoLine = i
        ElseIf fmLine < 0 And toLine < 0 And infoLine < 0 Then
            If LooksLikeDtgLine(key) Then
                If dtgLine < 0 Then dtgLine = i
            ElseIf precLine < 0 Then
                precLine = i
            End If
        End If
    Next

    ' Classification line at the top (per NAV26237 sample; ACP 121(G) 325).
    firstText = NextNonBlank(lines, 0)
    If firstText >= 0 Then
        If firstText <> topClass Then AddIssue issues, firstText, "the first line must be the classification line (CLASSIFICATION: ...), per the NAV26237 sample."
    End If

    ' Precedence word (ACP 121(G) 329).
    If precLine < 0 Then
        AddIssue issues, -1, "no precedence line (FLASH, IMMEDIATE, PRIORITY or ROUTINE) before FM."
    Else
        precLetter = PrecedenceLetter(LineKey(lines(precLine)))
        If Len(precLetter) = 0 Then AddIssue issues, precLine, "precedence '" & Trim$(lines(precLine)) & "' is not FLASH, IMMEDIATE, PRIORITY or ROUTINE."
        ' NTP 3 503: PRIORITY is the highest precedence for administrative
        ' messages, except IMMEDIATE for death, serious illness or injury.
        If isGenadmin And (precLetter = "O" Or precLetter = "Z") Then AddWarning issues, precLine, "GENADMIN is administrative; NTP 3 503 allows no higher than PRIORITY, except IMMEDIATE to report a death, serious illness or serious injury."
    End If

    ' Precedence prosign and DTG DDHHMMZ MON YY (ACP 121(G) 329, 317c, 324c).
    If dtgLine < 0 Then
        AddIssue issues, -1, "no date-time group line (DDHHMMZ MON YY, for example 091808Z OCT 26) before FM."
    Else
        words = SplitWords(LineKey(lines(dtgLine)))
        k = 0
        If Len(words(0)) = 1 Then
            prosign = words(0): k = 1
            If InStr(1, "ZOPR", prosign, vbBinaryCompare) = 0 Then
                AddIssue issues, dtgLine, "precedence prosign '" & prosign & "' must be Z, O, P or R."
            ElseIf Len(precLetter) > 0 And prosign <> precLetter Then
                AddIssue issues, dtgLine, "precedence prosign '" & prosign & "' does not match precedence " & LineKey(lines(precLine)) & " (" & precLetter & ")."
            End If
        Else
            AddIssue issues, dtgLine, "the DTG line must start with the precedence prosign (Z, O, P or R), for example R 091808Z OCT 26."
        End If
        If UBound(words) < k + 2 Then
            AddIssue issues, dtgLine, "the DTG must be DDHHMMZ MON YY, for example 091808Z OCT 26."
        Else
            CheckDtg issues, dtgLine, CStr(words(k)), CStr(words(k + 1)), CStr(words(k + 2))
        End If
    End If

    ' Addressing: FM, then TO with at least one address, then optional INFO
    ' (ACP 121(G) 310, 324f/g).
    If fmLine < 0 Then
        AddIssue issues, -1, "no FM (originator) line before BT."
    ElseIf Len(RestOfLine(LineKey(lines(fmLine)))) = 0 Then
        AddIssue issues, fmLine, "the FM line has no originator."
    End If
    If toLine < 0 Then
        AddIssue issues, -1, "no TO (action addressee) line before BT."
    Else
        If fmLine >= 0 And toLine < fmLine Then AddIssue issues, toLine, "TO must come after FM."
        hasAddress = Len(RestOfLine(LineKey(lines(toLine)))) > 0
        If Not hasAddress Then
            For j = toLine + 1 To headerEnd - 1
                key = LineKey(lines(j))
                If Len(key) > 0 Then
                    If FirstWord(key) <> "INFO" And FirstWord(key) <> "FM" And Left$(key, 15) <> "CLASSIFICATION:" Then hasAddress = True
                    Exit For
                End If
            Next
        End If
        If Not hasAddress Then AddIssue issues, toLine, "the TO line has no addressee."
    End If
    ' PLA length (warning only; NTP 3 603). The prosign FM, TO or INFO doesn't count.
    If fmLine >= 0 Then
        For j = fmLine To headerEnd - 1
            key = LineKey(lines(j))
            If FirstWord(key) = "FM" Or FirstWord(key) = "TO" Or FirstWord(key) = "INFO" Then key = RestOfLine(key)
            If Len(key) > PLA_LINE_LIMIT Then AddWarning issues, j, "the address is " & CStr(Len(key)) & " characters; NTP 3 603 limits a PLA line to " & CStr(PLA_LINE_LIMIT) & " (continue office codes on a line indented five spaces)."
        Next
    End If
    If infoLine >= 0 Then
        If toLine >= 0 And infoLine < toLine Then AddIssue issues, infoLine, "INFO must come after TO."
        If Len(RestOfLine(LineKey(lines(infoLine)))) = 0 Then
            j = NextNonBlank(lines, infoLine + 1)
            If j < 0 Or j >= headerEnd Then
                AddIssue issues, infoLine, "the INFO line has no addressee; remove it or add one."
            ElseIf FirstWord(LineKey(lines(j))) = "TO" Or FirstWord(LineKey(lines(j))) = "FM" Then
                AddIssue issues, infoLine, "the INFO line has no addressee; remove it or add one."
            End If
        End If
    End If

    ' Classification is the first text element after BT and matches the
    ' classification lines (ACP 121(G) 325, 330a; top/bottom lines per NAV26237 sample).
    If btOpen < 0 Then
        AddIssue issues, -1, "no BT line between the heading and the text."
    Else
        j = NextNonBlank(lines, btOpen + 1)
        If j >= 0 Then textClass = ClassificationOf(LineKey(lines(j)))
        If Len(textClass) = 0 Then
            AddIssue issues, j, "the first text line after BT must be the classification (UNCLAS, CONFIDENTIAL, SECRET or TOP SECRET)."
        Else
            ' An SSIC, when given, follows the classification as //N plus five
            ' digits// (NTP 3 706b), e.g. UNCLAS //N05216//. NAV26237 has none.
            ' Only //N and a digit is read as an SSIC, so UNCLAS//FOUO passes.
            p = InStr(1, LineKey(lines(j)), "//N", vbBinaryCompare)
            If p > 0 Then
                If Not IsDigitChar(Mid$(LineKey(lines(j)), p + 3, 1)) Then p = 0
            End If
            If p > 0 Then
                If Not (Mid$(LineKey(lines(j)), p, 10) Like "//N#####//") Then AddIssue issues, j, "the SSIC after the classification must be //N and five digits//, for example //N05216// (NTP 3 706)."
            End If
            If topClass >= 0 Then
                topValue = ClassificationOf(Trim$(Mid$(LineKey(lines(topClass)), 16)))
                If topValue <> textClass Then AddIssue issues, j, "the text classification (" & Trim$(lines(j)) & ") does not match the top classification line."
            End If
            lastLine = LastNonBlank(lines)
            If lastLine > msgidLine And Left$(LineKey(lines(lastLine)), 15) = "CLASSIFICATION:" Then
                bottomValue = ClassificationOf(Trim$(Mid$(LineKey(lines(lastLine)), 16)))
                If bottomValue <> textClass Then AddIssue issues, lastLine, "the bottom classification line does not match the text classification (" & Trim$(lines(j)) & ")."
            End If
        End If
    End If
    lastLine = LastNonBlank(lines)
    If lastLine >= 0 Then
        If lastLine <= msgidLine Or Left$(LineKey(lines(lastLine)), 15) <> "CLASSIFICATION:" Then AddIssue issues, lastLine, "the last line must be the classification line (CLASSIFICATION: ...), per the NAV26237 sample."
    End If

    ' Set lines MSGID, SUBJ, REF, AMPN or NARR, POC, AKNLDG, RMKS, DECL in order
    ' (NTP 3 Annex A, figure A-2), each ending with //.
    ' Order: ACP 121(G) 330d (identifier, subject, references, text); set
    ' delimiters and NARR-after-REF per NAV26237 sample; REF lettering ACP 121(G) 319.
    If btOpen >= 0 Then textStart = btOpen + 1 Else textStart = msgidLine
    If btClose >= 0 Then textEnd = btClose - 1 Else textEnd = ub
    curSet = "": curStart = -1: lastRank = 0: rmksStart = -1: rmksEnd = -1
    For i = textStart To textEnd
        key = LineKey(lines(i))
        setName = SetKeyword(key)
        If Len(setName) > 0 Then
            If Len(curSet) > 0 Then
                ReportUnterminated issues, curSet, curStart
                If curSet = "RMKS" And rmksEnd < 0 Then rmksEnd = i - 1
                curSet = ""
            End If
            rank = SetRank(setName)
            If rank < lastRank Then AddIssue issues, i, setName & " is out of order; the order is MSGID, SUBJ, REF, AMPN or NARR, POC, AKNLDG, RMKS, DECL (NTP 3 Annex A)."
            If rank > lastRank Then lastRank = rank
            Select Case setName
                Case "MSGID"
                    msgidCount = msgidCount + 1
                    If msgidCount = 2 Then AddIssue issues, i, "MSGID appears more than once."
                    If isGenadmin And msgidCount = 1 Then
                        p = InStr(16, key, "/", vbBinaryCompare)
                        If p > 16 Then msgidOrig = Mid$(key, 16, p - 16) Else msgidOrig = ""
                        If Len(msgidOrig) > MSGID_ORIG_LIMIT Then AddWarning issues, i, "the MSGID originator '" & msgidOrig & "' is " & CStr(Len(msgidOrig)) & " characters; NTP 3 Annex A allows " & CStr(MSGID_ORIG_LIMIT) & " (use a shorter form, e.g. without the location)."
                    End If
                Case "SUBJ"
                    subjCount = subjCount + 1
                    If subjCount = 2 Then AddIssue issues, i, "SUBJ appears more than once."
                    subjText = Trim$(Mid$(key, 6))
                    If Right$(subjText, 2) = "//" Then subjText = RTrim$(Left$(subjText, Len(subjText) - 2))
                    If Len(subjText) > SUBJ_FIELD_LIMIT Then AddWarning issues, i, "the subject is " & CStr(Len(subjText)) & " characters; NTP 3 Annex A allows " & CStr(SUBJ_FIELD_LIMIT) & " on the SUBJ line."
                Case "NARR"
                    narrCount = narrCount + 1
                    If narrCount = 1 Then narrLine = i
                    If narrCount = 2 Then AddIssue issues, i, "NARR appears more than once."
                Case "AMPN"
                    ampnCount = ampnCount + 1
                    If ampnCount = 1 Then ampnLine = i
                    If ampnCount = 2 Then AddIssue issues, i, "AMPN appears more than once; with several references use one NARR after the last REF (NTP 3 Annex A)."
                Case "POC"
                    pocCount = pocCount + 1
                Case "RMKS"
                    rmksCount = rmksCount + 1
                    If rmksCount = 2 Then AddIssue issues, i, "RMKS appears more than once."
                    If rmksCount = 1 Then rmksStart = i
                Case "REF"
                    refCount = refCount + 1
                    expectedLetter = Chr$(64 + ((refCount - 1) Mod 26) + 1)
                    If Mid$(key, 6, 1) <> "/" Or Not IsUpperLetter(Mid$(key, 5, 1)) Then
                        AddIssue issues, i, "a REF line must start REF/<letter>/, for example REF/" & expectedLetter & "/."
                    ElseIf Mid$(key, 5, 1) <> expectedLetter Then
                        AddIssue issues, i, "REF letter " & Mid$(key, 5, 1) & " is out of order; expected " & expectedLetter & " (references are lettered A, B, C in order)."
                    End If
                    ' Field 2 is a message identifier (GENADMIN, CASREP...) or a
                    ' communication code; the five codes need AMPN or NARR (NTP 3 Annex A).
                    p = InStr(7, key, "/", vbBinaryCompare)
                    If p > 7 Then refType = Mid$(key, 7, p - 7) Else refType = Mid$(key, 7)
                    If InStr(1, " CON DOC LTR MSG TEL ", " " & Trim$(refType) & " ", vbBinaryCompare) > 0 Then needsExplain = True
            End Select
            If EndsWithSetTerminator(lines(i)) Then
                If setName = "RMKS" And rmksStart = i And rmksEnd < 0 Then rmksEnd = i
            Else
                curSet = setName: curStart = i
            End If
        ElseIf Len(curSet) > 0 Then
            If EndsWithSetTerminator(lines(i)) Then
                If curSet = "RMKS" And rmksEnd < 0 Then rmksEnd = i
                curSet = ""
            End If
        End If
    Next
    If Len(curSet) > 0 Then
        ReportUnterminated issues, curSet, curStart
        If curSet = "RMKS" And rmksEnd < 0 Then rmksEnd = textEnd
    End If
    If subjCount = 0 Then AddIssue issues, -1, "no SUBJ/ set line (ACP 121(G) 330d)."
    ' AMPN explains a single reference and NARR several; never both (NTP 3 Annex A).
    If needsExplain And narrCount = 0 And ampnCount = 0 Then AddIssue issues, -1, "a CON, DOC, LTR, MSG or TEL reference needs an AMPN set (one reference) or a NARR set (several) that explains it (NTP 3 Annex A)."
    If ampnCount > 0 And narrCount > 0 Then AddIssue issues, ampnLine, "use AMPN or NARR, not both (NTP 3 Annex A)."
    If ampnCount > 0 And refCount > 1 Then AddIssue issues, ampnLine, "AMPN is for a single reference; with several references use one NARR after the last REF (NTP 3 Annex A)."
    If ampnCount > 0 And refCount = 0 Then AddIssue issues, ampnLine, "AMPN needs a REF set before it."
    If narrCount > 0 And ampnCount = 0 And refCount = 1 Then AddIssue issues, narrLine, "with one reference use AMPN, not NARR (NTP 3 Annex A)."
    ' NTP 3 712 asks for a POC set in every GENADMIN. NAV26237 gives its point
    ' of contact as an RMKS paragraph instead, which also passes.
    If isGenadmin And pocCount = 0 And InStr(1, UCase$(text), "POINT OF CONTACT", vbBinaryCompare) = 0 Then AddWarning issues, -1, "no POC/ set or point of contact paragraph; NTP 3 712 asks for a POC in every GENADMIN message."
    If rmksCount = 0 Then AddIssue issues, -1, "no RMKS/ set line."

    ' RMKS paragraph numbering 1., a., (1), (a) (ACP 121(G) 324j). An
    ' unnumbered single paragraph has no markers and passes. A marker counts only
    ' on the RMKS/ line, an indented line or a line after a blank line, so a
    ' wrapped line that happens to start "(a) and (b)." or "M. SMITH" is prose.
    If rmksStart >= 0 And rmksEnd >= rmksStart Then
        For i = rmksStart To rmksEnd
            s = LTrim$(lines(i))
            isMarker = (i = rmksStart)
            If Not isMarker Then isMarker = (Left$(lines(i), 1) = " ")
            If Not isMarker Then isMarker = (Len(Trim$(lines(i - 1))) = 0)
            If i = rmksStart Then s = LTrim$(Mid$(s, 6))
            If Not isMarker Then
                ' a continuation line of the paragraph above
            ElseIf ParseMarker(s, level, number) Then
                If level > 1 Then
                    If counters(level - 1) = 0 Then AddIssue issues, i, "RMKS paragraph " & MarkerText(level, number) & " has no " & LevelName(level - 1) & " above it; the scheme is 1., a., (1), (a)."
                End If
                If number <> counters(level) + 1 Then AddIssue issues, i, "RMKS paragraph " & MarkerText(level, number) & " is out of sequence; expected " & MarkerText(level, counters(level) + 1)
                If counters(level) = 0 Then groupLine(level) = i
                counters(level) = number
                For j = level + 1 To 4
                    WarnLoneMarker issues, j, counters(j), groupLine(j)
                    counters(j) = 0
                Next
            End If
        Next
        For j = 1 To 4
            WarnLoneMarker issues, j, counters(j), groupLine(j)
        Next
    End If

    ' Social Security numbers (warning only): SECNAV M-5216.5 2-7 says to leave
    ' them out and use the 10-digit DoD ID. A masked XXX-XX-1234 doesn't match.
    For i = 0 To ub
        If HasSsnPattern(lines(i)) Then
            issues = issues & "Warning: naval format line " & CStr(i + 1) & ": looks like a Social Security number; use the 10-digit DoD ID instead (SECNAV M-5216.5 2-7)." & vbCrLf
        End If
    Next

    ' Trailer: BT, # plus 4-digit station serial number, NNNN
    ' (NAVEDTRA 14226 ch. 1; per NAV26237 sample).
    If btClose < 0 Then
        AddIssue issues, -1, "no closing BT line after the text."
    Else
        j = NextNonBlank(lines, btClose + 1)
        If j < 0 Then
            AddIssue issues, btClose, "the closing BT must be followed by # and a 4-digit serial number (for example #0001), then NNNN."
        ElseIf Not IsSerialLine(LineKey(lines(j))) Then
            AddIssue issues, j, "the closing BT must be followed by # and a 4-digit serial number (for example #0001), then NNNN."
        Else
            k = NextNonBlank(lines, j + 1)
            If k < 0 Then
                AddIssue issues, j, "the serial number must be followed by NNNN."
            ElseIf LineKey(lines(k)) <> "NNNN" Then
                AddIssue issues, k, "the serial number must be followed by NNNN."
            End If
        End If
    End If

    ' NTP 3 713e and A102g: & ' $ and ; aren't allowed in GENADMIN messages.
    ' Warning only, once, because the NAV26237 sample uses apostrophes.
    If isGenadmin Then
        For i = textStart To textEnd
            s = FirstCharOf(lines(i), "&'$;")
            If Len(s) > 0 Then
                AddWarning issues, i, "contains " & s & "; NTP 3 713e doesn't allow & ' $ or ; in a GENADMIN message (spell it out or rephrase)."
                Exit For
            End If
        Next
    End If

    ' Line length (warning only; see NAVAL_LINE_LIMIT) and characters: tabs and
    ' anything outside printable ASCII (per NAV26237 sample, which is plain ASCII
    ' with spaces; ACP 127/NAVEDTRA 14226 teletypewriter character set).
    For i = 0 To ub
        If Len(lines(i)) > NAVAL_LINE_LIMIT Then
            issues = issues & "Warning: naval format line " & CStr(i + 1) & ": exceeds the " & CStr(NAVAL_LINE_LIMIT) & "-character line limit of NTP 3 A102b (" & CStr(Len(lines(i))) & " characters)." & vbCrLf
        End If
        If InStr(1, lines(i), vbTab, vbBinaryCompare) > 0 Then AddIssue issues, i, "contains a tab character; use spaces."
        s = FirstNonAsciiCode(lines(i))
        If Len(s) > 0 Then AddIssue issues, i, "contains a character outside printable ASCII (U+" & s & ")."
    Next

    NavalFormatIssues = issues
End Function

' The blocking issues (warnings False) or the "Warning:" lines (warnings
' True) of NavalFormatIssues' text, each line ending in vbCrLf.
'   issues    NavalFormatIssues' result
'   warnings  True for the warnings, False for the blocking issues
' Errors: none.
Public Function FormatIssuesOf(ByVal issues As String, ByVal warnings As Boolean) As String
    Dim parts As Variant, i As Long, result As String, isWarning As Boolean
    parts = Split(Replace(issues, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            isWarning = (Left$(parts(i), 8) = "Warning:")
            If isWarning = warnings Then result = result & parts(i) & vbCrLf
        End If
    Next
    FormatIssuesOf = result
End Function

Private Sub AddIssue(ByRef issues As String, ByVal lineIndex As Long, ByVal message As String)
    If lineIndex >= 0 Then
        issues = issues & "Naval format line " & CStr(lineIndex + 1) & ": " & message & vbCrLf
    Else
        issues = issues & "Naval format: " & message & vbCrLf
    End If
End Sub

Private Sub AddWarning(ByRef issues As String, ByVal lineIndex As Long, ByVal message As String)
    If lineIndex >= 0 Then
        issues = issues & "Warning: naval format line " & CStr(lineIndex + 1) & ": " & message & vbCrLf
    Else
        issues = issues & "Warning: naval format: " & message & vbCrLf
    End If
End Sub

' Returns the first character of lineText that is one of chars, or "".
Private Function FirstCharOf(ByVal lineText As Variant, ByVal chars As String) As String
    Dim i As Long, t As String
    t = CStr(lineText)
    For i = 1 To Len(t)
        If InStr(1, chars, Mid$(t, i, 1), vbBinaryCompare) > 0 Then
            FirstCharOf = Mid$(t, i, 1)
            Exit Function
        End If
    Next
End Function

' Warning for a group of RMKS markers that ended with only its first member:
' a lone 1. (ACP 121(G) 324j: don't number a single paragraph), or a lone a.,
' (1) or (a) (SECNAV M-5216.5 7-12: a 1a needs a 1b). Advisory only, because
' the NAV26237 sample itself has 4.c.(1) with no (2).
Private Sub WarnLoneMarker(ByRef issues As String, ByVal level As Long, ByVal lastNumber As Long, ByVal lineIndex As Long)
    If lastNumber <> 1 Then Exit Sub
    If level = 1 Then
        issues = issues & "Warning: naval format line " & CStr(lineIndex + 1) & ": RMKS has only paragraph 1.; a single paragraph isn't numbered (ACP 121(G) 324j)." & vbCrLf
    Else
        issues = issues & "Warning: naval format line " & CStr(lineIndex + 1) & ": RMKS paragraph " & MarkerText(level, 1) & " has no " & MarkerText(level, 2) & "; split it only if there are at least two (SECNAV M-5216.5 7-12)." & vbCrLf
    End If
End Sub

' True when the line holds ###-##-#### with no digit or hyphen on either side.
Private Function HasSsnPattern(ByVal s As Variant) As Boolean
    Dim t As String, i As Long, before As String, after As String
    t = CStr(s)
    For i = 1 To Len(t) - 10
        If Mid$(t, i, 11) Like "###-##-####" Then
            If i > 1 Then before = Mid$(t, i - 1, 1) Else before = " "
            after = Mid$(t, i + 11, 1)
            If Not (IsDigitChar(before) Or before = "-" Or IsDigitChar(after) Or after = "-") Then
                HasSsnPattern = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub ReportUnterminated(ByRef issues As String, ByVal setName As String, ByVal lineIndex As Long)
    If setName = "RMKS" Then
        ' NTP 3 713f requires it, but the NAV26237 sample's RMKS has no closing
        ' //, so this is advisory only.
        issues = issues & "Warning: naval format line " & CStr(lineIndex + 1) & ": RMKS should end with // after its last paragraph (NTP 3 713f; the NAV26237 sample omits it)." & vbCrLf
    Else
        AddIssue issues, lineIndex, setName & " does not end with //."
    End If
End Sub

Private Function LineKey(ByVal s As Variant) As String
    LineKey = UCase$(Trim$(CStr(s)))
End Function

Private Function FirstWord(ByVal key As String) As String
    Dim p As Long
    p = InStr(1, key, " ", vbBinaryCompare)
    If p = 0 Then FirstWord = key Else FirstWord = Left$(key, p - 1)
End Function

Private Function RestOfLine(ByVal key As String) As String
    Dim p As Long
    p = InStr(1, key, " ", vbBinaryCompare)
    If p > 0 Then RestOfLine = Trim$(Mid$(key, p + 1))
End Function

' Returns the words of s separated by one or more spaces; always at least one
' element (an empty string when s is blank).
Private Function SplitWords(ByVal s As String) As Variant
    s = Trim$(s)
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    If Len(s) = 0 Then SplitWords = Array("") Else SplitWords = Split(s, " ")
End Function

Private Function NextNonBlank(ByVal lines As Variant, ByVal start As Long) As Long
    Dim i As Long
    NextNonBlank = -1
    For i = start To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 Then NextNonBlank = i: Exit Function
    Next
End Function

Private Function LastNonBlank(ByVal lines As Variant) As Long
    Dim i As Long
    LastNonBlank = -1
    For i = UBound(lines) To 0 Step -1
        If Len(Trim$(CStr(lines(i)))) > 0 Then LastNonBlank = i: Exit Function
    Next
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    If Len(ch) = 1 Then IsDigitChar = (ch >= "0" And ch <= "9")
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Function
    Next
    IsDigits = True
End Function

Private Function IsUpperLetter(ByVal ch As String) As Boolean
    If Len(ch) = 1 Then IsUpperLetter = (ch >= "A" And ch <= "Z")
End Function

Private Function IsLetterChar(ByVal ch As String) As Boolean
    If Len(ch) = 1 Then IsLetterChar = ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z"))
End Function

' A DTG line is "<prosign> DDHHMMZ ..." or "DDHHMMZ ...": its date part starts with a digit.
Private Function LooksLikeDtgLine(ByVal key As String) As Boolean
    Dim words As Variant
    words = SplitWords(key)
    If Len(words(0)) = 1 And UBound(words) >= 1 Then
        LooksLikeDtgLine = IsDigitChar(Left$(words(1), 1))
    Else
        LooksLikeDtgLine = IsDigitChar(Left$(words(0), 1))
    End If
End Function

' ACP 121(G) 329: FLASH Z, IMMEDIATE O, PRIORITY P, ROUTINE R.
Private Function PrecedenceLetter(ByVal word As String) As String
    Select Case word
        Case "FLASH": PrecedenceLetter = "Z"
        Case "IMMEDIATE": PrecedenceLetter = "O"
        Case "PRIORITY": PrecedenceLetter = "P"
        Case "ROUTINE": PrecedenceLetter = "R"
    End Select
End Function

' ACP 121(G) 317c, 324c: DDHHMMZ MON YY. Zone letters run A-Z without J
' (the drafting rules: Z unless the area commander sets local time).
Private Sub CheckDtg(ByRef issues As String, ByVal lineIndex As Long, ByVal group As String, ByVal monthText As String, ByVal yearText As String)
    Dim zone As String
    If Len(group) <> 7 Or Not IsDigits(Left$(group, 6)) Then
        AddIssue issues, lineIndex, "the DTG '" & group & "' must be six digits and a zone letter (DDHHMMZ)."
    Else
        If CLng(Mid$(group, 1, 2)) < 1 Or CLng(Mid$(group, 1, 2)) > 31 Then AddIssue issues, lineIndex, "DTG day " & Mid$(group, 1, 2) & " must be 01-31."
        If CLng(Mid$(group, 3, 2)) > 23 Then AddIssue issues, lineIndex, "DTG hour " & Mid$(group, 3, 2) & " must be 00-23."
        If CLng(Mid$(group, 5, 2)) > 59 Then AddIssue issues, lineIndex, "DTG minute " & Mid$(group, 5, 2) & " must be 00-59."
        If Mid$(group, 3, 4) = "0000" Then AddIssue issues, lineIndex, "DTG time 0000 isn't used; use 0001 (NTP 3 504)."
        zone = Mid$(group, 7, 1)
        If Not IsUpperLetter(zone) Or zone = "J" Then AddIssue issues, lineIndex, "DTG zone '" & zone & "' must be a letter A-Z other than J."
    End If
    If Len(monthText) <> 3 Or InStr(1, " JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC ", " " & monthText & " ", vbBinaryCompare) = 0 Then
        AddIssue issues, lineIndex, "DTG month '" & monthText & "' must be a three-letter month (JAN-DEC)."
    End If
    If Len(yearText) <> 2 Or Not IsDigits(yearText) Then AddIssue issues, lineIndex, "DTG year '" & yearText & "' must be two digits."
End Sub

' True when key starts with word followed by the end, a space or a slash.
Private Function StartsWithWord(ByVal key As String, ByVal word As String) As Boolean
    Dim nextChar As String
    If Left$(key, Len(word)) <> word Then Exit Function
    nextChar = Mid$(key, Len(word) + 1, 1)
    StartsWithWord = (Len(nextChar) = 0 Or nextChar = " " Or nextChar = "/")
End Function

' Normalises UNCLAS/UNCLASSIFIED, CONFIDENTIAL, SECRET, TOP SECRET (ACP 121(G) 325).
Private Function ClassificationOf(ByVal key As String) As String
    If StartsWithWord(key, "UNCLAS") Or StartsWithWord(key, "UNCLASSIFIED") Then
        ClassificationOf = "UNCLASSIFIED"
    ElseIf StartsWithWord(key, "TOP SECRET") Then
        ClassificationOf = "TOP SECRET"
    ElseIf StartsWithWord(key, "SECRET") Then
        ClassificationOf = "SECRET"
    ElseIf StartsWithWord(key, "CONFIDENTIAL") Then
        ClassificationOf = "CONFIDENTIAL"
    ElseIf Left$(key, 17) = "T O P S E C R E T" Then
        ClassificationOf = "TOP SECRET"
    ElseIf Left$(key, 11) = "S E C R E T" Then
        ClassificationOf = "SECRET"
    ElseIf Left$(key, 23) = "C O N F I D E N T I A L" Then
        ' NTP 3 702a spells a classified text classification with spaces.
        ClassificationOf = "CONFIDENTIAL"
    End If
End Function

Private Function SetKeyword(ByVal key As String) As String
    If Left$(key, 6) = "MSGID/" Then
        SetKeyword = "MSGID"
    ElseIf Left$(key, 5) = "SUBJ/" Then
        SetKeyword = "SUBJ"
    ElseIf Left$(key, 4) = "REF/" Then
        SetKeyword = "REF"
    ElseIf Left$(key, 5) = "NARR/" Then
        SetKeyword = "NARR"
    ElseIf Left$(key, 5) = "AMPN/" Then
        SetKeyword = "AMPN"
    ElseIf Left$(key, 4) = "POC/" Then
        SetKeyword = "POC"
    ElseIf Left$(key, 7) = "AKNLDG/" Then
        SetKeyword = "AKNLDG"
    ElseIf Left$(key, 5) = "DECL/" Then
        SetKeyword = "DECL"
    ElseIf Left$(key, 5) = "RMKS/" Then
        SetKeyword = "RMKS"
    End If
End Function

Private Function SetRank(ByVal setName As String) As Long
    Select Case setName
        Case "MSGID": SetRank = 1
        Case "SUBJ": SetRank = 2
        Case "REF": SetRank = 3
        Case "NARR", "AMPN": SetRank = 4
        Case "POC": SetRank = 5
        Case "AKNLDG": SetRank = 6
        Case "RMKS": SetRank = 7
        Case "DECL": SetRank = 8
    End Select
End Function

Private Function EndsWithSetTerminator(ByVal s As Variant) As Boolean
    EndsWithSetTerminator = (Right$(RTrim$(CStr(s)), 2) = "//")
End Function

' Paragraph markers at the start of a line: 1. (level 1), a. (2), (1) (3), (a) (4).
' Letters are case-insensitive so ALL CAPS drafts number A., (A) the same way.
Private Function ParseMarker(ByVal s As String, ByRef level As Long, ByRef number As Long) As Boolean
    Dim p As Long, closing As Long, inner As String, ch As String
    level = 0: number = 0
    If Len(s) = 0 Then Exit Function
    ch = Left$(s, 1)
    If IsDigitChar(ch) Then
        p = 1
        Do While p < Len(s)
            If Not IsDigitChar(Mid$(s, p + 1, 1)) Then Exit Do
            p = p + 1
        Loop
        If p <= 3 And Mid$(s, p + 1, 1) = "." And MarkerEnds(s, p + 2) Then
            level = 1: number = CLng(Left$(s, p))
        End If
    ElseIf IsLetterChar(ch) Then
        If Mid$(s, 2, 1) = "." And MarkerEnds(s, 3) Then
            level = 2: number = Asc(LCase$(ch)) - 96
        End If
    ElseIf ch = "(" Then
        closing = InStr(2, s, ")", vbBinaryCompare)
        If closing > 2 Then
            inner = Mid$(s, 2, closing - 2)
            If MarkerEnds(s, closing + 1) Then
                If Len(inner) <= 3 And IsDigits(inner) Then
                    level = 3: number = CLng(inner)
                ElseIf Len(inner) = 1 And IsLetterChar(inner) Then
                    level = 4: number = Asc(LCase$(inner)) - 96
                End If
            End If
        End If
    End If
    ParseMarker = (level > 0)
End Function

Private Function MarkerEnds(ByVal s As String, ByVal position As Long) As Boolean
    MarkerEnds = (position > Len(s))
    If Not MarkerEnds Then MarkerEnds = (Mid$(s, position, 1) = " ")
End Function

Private Function MarkerText(ByVal level As Long, ByVal number As Long) As String
    Select Case level
        Case 1: MarkerText = CStr(number) & "."
        Case 2: MarkerText = LetterOf(number) & "."
        Case 3: MarkerText = "(" & CStr(number) & ")"
        Case 4: MarkerText = "(" & LetterOf(number) & ")"
    End Select
End Function

Private Function LetterOf(ByVal number As Long) As String
    If number >= 1 And number <= 26 Then LetterOf = Chr$(96 + number) Else LetterOf = "?"
End Function

Private Function LevelName(ByVal level As Long) As String
    Select Case level
        Case 1: LevelName = "numbered paragraph (1.)"
        Case 2: LevelName = "subparagraph (a.)"
        Case 3: LevelName = "sub-subparagraph ((1))"
    End Select
End Function

Private Function IsSerialLine(ByVal key As String) As Boolean
    IsSerialLine = (Len(key) = 5 And Left$(key, 1) = "#" And IsDigits(Mid$(key, 2)))
End Function

' Returns the hex code of the first character outside printable ASCII
' (space through ~), ignoring tabs, which are reported separately.
Private Function FirstNonAsciiCode(ByVal s As Variant) As String
    Dim i As Long, code As Long, t As String
    t = CStr(s)
    For i = 1 To Len(t)
        code = AscW(Mid$(t, i, 1)) And &HFFFF&
        If code <> 9 And (code < 32 Or code > 126) Then
            FirstNonAsciiCode = Right$("000" & Hex$(code), 4)
            Exit Function
        End If
    Next
End Function
