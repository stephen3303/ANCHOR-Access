Option Compare Database
Option Explicit

' ANCHOR (Access edition) - Search: the SQL behind the Search page, its
' criteria as text (for saved searches), and saving, listing and deleting
' a person's saved searches (tblSavedSearches).
'
' A search's criteria are a Scripting.Dictionary of text values, all keys
' always present ("" for not set), made by NewSearchCriteria:
'   mode     basic or advanced (basic uses the words only)
'   words    the words typed: every word must match somewhere in the
'            subject, reference, originator or body text; "a phrase" in
'            double quotes matches as one piece. (The web app's natural-
'            language search is this keyword search here: no live AI.)
'   ref      part of the reference                       (advanced)
'   orig     the originator, exactly                      (advanced)
'   msgtype  the message type code (tblMessageTypes)      (advanced)
'   cls      the communications class A-E (tblMessageClasses) (advanced)
'   prec     the precedence                               (advanced)
'   classif  the classification                           (advanced)
'   life     the lifecycle                                (advanced)
'   from, to received on or after / on or before, yyyy-mm-dd (advanced)
'   ack      "Required" or "Not required"                 (advanced)
' Saved, they are one line of key=value pairs joined by ";", with "%", ";",
' "=" and line breaks escaped (SearchCriteriaText, SearchCriteriaFromText).
'
' Running a search is audited by the page (AuditSearchRun): action Search,
' resource "Keyword: <words>", as the web app logs it.
'
' Owner: the Search page. Depends on: modAudit, modGlobals, modMessages,
' modSignIn; the kit's modKitLocks (KitExecute), modPersonnel.

Private Const CRIT_KEYS As String = "mode;words;ref;orig;msgtype;cls;prec;classif;life;from;to;ack"

' A new, empty set of criteria (every key "", mode basic).
' Returns a Scripting.Dictionary. Errors: none.
Public Function NewSearchCriteria() As Object
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                           ' keys ignore case
    For Each k In Split(CRIT_KEYS, ";")
        d(k) = ""
    Next
    d("mode") = "basic"
    Set NewSearchCriteria = d
End Function

' The words typed, split into search terms: at spaces, keeping "a phrase in
' double quotes" as one term. Returns an array of strings (none: an empty
' array, UBound -1). Errors: none.
Public Function SearchTerms(ByVal text As String) As Variant
    Dim out As String, i As Long, ch As String, cur As String, inQuote As Boolean
    text = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch = """" Then
            If inQuote Then
                If Len(Trim$(cur)) > 0 Then out = out & vbTab & Trim$(cur)
                cur = "": inQuote = False
            Else
                If Len(cur) > 0 Then out = out & vbTab & cur
                cur = "": inQuote = True
            End If
        ElseIf ch = " " And Not inQuote Then
            If Len(cur) > 0 Then out = out & vbTab & cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next
    If Len(Trim$(cur)) > 0 Then out = out & vbTab & Trim$(cur)
    If Len(out) = 0 Then SearchTerms = Split("") Else SearchTerms = Split(Mid$(out, 2), vbTab)
End Function

' True when the criteria set anything an advanced search filters by.
Public Function SearchHasFilters(ByVal crit As Object) As Boolean
    Dim k As Variant
    For Each k In Split(CRIT_KEYS, ";")
        If k <> "mode" And k <> "words" Then
            If Len(crit(k)) > 0 Then SearchHasFilters = True: Exit Function
        End If
    Next
End Function

' The WHERE condition for the criteria over MessagesFromSql (alias m),
' without the word WHERE; "True" when nothing is set. Basic mode uses the
' words only. Every text value goes through Q, and words through
' LikeLiteral, so typed wildcards match themselves.
Public Function SearchWhereSql(ByVal crit As Object) As String
    Dim w As String, t As Variant, p As String
    For Each t In SearchTerms(crit("words"))
        p = Q("*" & LikeLiteral(CStr(t)) & "*")
        w = w & " AND (m.Subject LIKE " & p & " OR m.Reference LIKE " & p & " OR m.Originator LIKE " & p & _
            " OR m.BodyText LIKE " & p & ")"
    Next
    If crit("mode") = "advanced" Then
        If Len(crit("ref")) > 0 Then w = w & " AND m.Reference LIKE " & Q("*" & LikeLiteral(crit("ref")) & "*")
        If Len(crit("orig")) > 0 Then w = w & " AND m.Originator=" & Q(crit("orig"))
        If Len(crit("msgtype")) > 0 Then w = w & " AND m.MsgType=" & Q(crit("msgtype"))
        If Len(crit("cls")) > 0 Then w = w & " AND m.MsgClass=" & Q(crit("cls"))
        If Len(crit("prec")) > 0 Then w = w & " AND m.Precedence=" & Q(crit("prec"))
        If Len(crit("classif")) > 0 Then w = w & " AND m.Classification=" & Q(crit("classif"))
        If Len(crit("life")) > 0 Then w = w & " AND m.Lifecycle=" & Q(crit("life"))
        ' Dates as #yyyy-mm-dd#, which the engine reads the same in every
        ' locale; "to" takes the whole day.
        If IsDate(crit("from")) Then w = w & " AND m.ReceivedOn>=#" & Format$(CDate(crit("from")), "yyyy-mm-dd") & "#"
        If IsDate(crit("to")) Then w = w & " AND m.ReceivedOn<#" & Format$(DateAdd("d", 1, CDate(crit("to"))), "yyyy-mm-dd") & "#"
        If crit("ack") = "Required" Then w = w & " AND m.AckRequired"
        If crit("ack") = "Not required" Then w = w & " AND NOT m.AckRequired"
    End If
    If Len(w) = 0 Then SearchWhereSql = "True" Else SearchWhereSql = Mid$(w, 6)
End Function

' The results' record source: the message list's shared columns
' (MessageListColsSql) and classif_txt, the classification with its demo
' marking note (ClassificationMarkingSql), newest first.
Public Function SearchSql(ByVal crit As Object) As String
    SearchSql = "SELECT " & MessageListColsSql() & ", " & ClassificationMarkingSql("m.Classification") & " AS classif_txt FROM " & _
        MessagesFromSql() & " WHERE " & SearchWhereSql(crit) & " ORDER BY m.ReceivedOn DESC"
End Function

' The criteria as one line of text, for tblSavedSearches.Criteria.
Public Function SearchCriteriaText(ByVal crit As Object) As String
    Dim k As Variant, s As String
    For Each k In Split(CRIT_KEYS, ";")
        If Len(crit(k)) > 0 Then s = s & ";" & k & "=" & CritEscape(crit(k))
    Next
    SearchCriteriaText = Mid$(s, 2)
End Function

' Criteria read back from SearchCriteriaText's line. Unknown keys are
' ignored, so a line saved by a later version still opens.
Public Function SearchCriteriaFromText(ByVal s As Variant) As Object
    Dim d As Object, part As Variant, k As String, eq As Long
    Set d = NewSearchCriteria()
    For Each part In Split(Nz(s, ""), ";")
        eq = InStr(part, "=")
        If eq > 1 Then
            k = Left$(part, eq - 1)
            If InStr(";" & CRIT_KEYS & ";", ";" & k & ";") > 0 Then d(k) = CritUnescape(Mid$(part, eq + 1))
        End If
    Next
    If d("mode") <> "advanced" Then d("mode") = "basic"
    Set SearchCriteriaFromText = d
End Function

' The criteria in words, for the saved list and the audit trail:
' "readiness deployment; Type NAVADMIN; Received from 01 Oct 2026".
Public Function SearchSummary(ByVal crit As Object) As String
    Dim s As String
    If Len(crit("words")) > 0 Then s = "; " & crit("words")
    If crit("mode") = "advanced" Then
        If Len(crit("ref")) > 0 Then s = s & "; Reference " & crit("ref")
        If Len(crit("orig")) > 0 Then s = s & "; Originator " & crit("orig")
        If Len(crit("msgtype")) > 0 Then s = s & "; Type " & crit("msgtype")
        If Len(crit("cls")) > 0 Then s = s & "; Class " & crit("cls")
        If Len(crit("prec")) > 0 Then s = s & "; " & crit("prec")
        If Len(crit("classif")) > 0 Then s = s & "; " & ClassificationMarking(crit("classif"))
        If Len(crit("life")) > 0 Then s = s & "; " & crit("life")
        If IsDate(crit("from")) Then s = s & "; Received from " & Format$(CDate(crit("from")), "dd mmm yyyy")
        If IsDate(crit("to")) Then s = s & "; Received to " & Format$(CDate(crit("to")), "dd mmm yyyy")
        If Len(crit("ack")) > 0 Then s = s & "; Ack " & LCase$(crit("ack"))
    End If
    If Len(s) = 0 Then SearchSummary = "Every message" Else SearchSummary = Mid$(s, 3)
End Function

' SearchSummary of a saved line, for the saved list's record source (it is
' called from SQL, so it takes a Variant and never raises).
Public Function SavedSearchSummary(ByVal s As Variant) As String
    On Error Resume Next
    SavedSearchSummary = SearchSummary(SearchCriteriaFromText(s))
End Function

' Audit one search run (AuditTry: a failure is logged, the search goes on):
' action Search, resource "Keyword: <words>", and how many results.
Public Sub AuditSearchRun(ByVal crit As Object, ByVal n As Long)
    Dim kw As String
    kw = Trim$(crit("words"))
    If Len(kw) = 0 Then kw = "(none)"
    AuditTry "Search", "Keyword: " & kw, "Search returned " & n & " " & Plural(n, "result") & ". " & _
        IIf(crit("mode") = "advanced", "Advanced: ", "Basic: ") & SearchSummary(crit) & "."
End Sub

' ---------------------------------------------------------------------------
' Saved searches (each person's own)
' ---------------------------------------------------------------------------

' The signed-in person's saved searches, by name: id, SearchName, summary,
' created_txt and CreatedOn.
Public Function SavedSearchesSql() As String
    SavedSearchesSql = "SELECT SearchID AS id, SearchName, SavedSearchSummary(Criteria) AS summary, CreatedOn, " & _
        "Format(CreatedOn,'dd mmm yyyy') AS created_txt FROM tblSavedSearches WHERE PersonID=" & SignedInPersonID() & _
        " ORDER BY SearchName"
End Function

' The id of the signed-in person's saved search called searchName, 0 for none.
Public Function SavedSearchID(ByVal searchName As String) As Long
    If Not DataLinked() Then Exit Function
    SavedSearchID = Nz(DLookup("SearchID", "tblSavedSearches", "PersonID=" & SignedInPersonID() & _
        " AND SearchName=" & Q(Trim$(searchName))), 0)
End Function

' Save the criteria as the signed-in person's search searchName, replacing
' their search of that name if they have one, and audit it.
' Returns "" when saved, else what to tell them (no name, too long, nobody
' signed in, no data file).
' Errors: ERR_NOT_ALLOWED without the Message Viewer role (RequireRole);
'   the engine's and KitExecute's.
' Side effects: one tblSavedSearches row added or changed; an audit row.
Public Function SaveSearch(ByVal searchName As String, ByVal crit As Object) As String
    Dim pid As Long, id As Long, txt As String
    searchName = Trim$(Replace(Replace(searchName, vbCr, " "), vbLf, " "))
    If Not DataLinked() Then SaveSearch = "This copy has no data file, so searches can't be saved.": Exit Function
    If Len(searchName) = 0 Then SaveSearch = "Type a name for the search.": Exit Function
    If Len(searchName) > 100 Then SaveSearch = "Keep the name to 100 characters or fewer.": Exit Function
    pid = SignedInPersonID()
    If pid = 0 Then SaveSearch = "Nobody is signed in.": Exit Function
    RequireRole "VIEW", "save a search"
    txt = SearchCriteriaText(crit)
    id = SavedSearchID(searchName)
    If id = 0 Then
        KitExecute "INSERT INTO tblSavedSearches (PersonID, SearchName, Criteria, CreatedOn) VALUES (" & pid & ", " & _
            Q(searchName) & ", " & Q(txt) & ", Now())"
    Else
        KitExecute "UPDATE tblSavedSearches SET Criteria=" & Q(txt) & ", CreatedOn=Now() WHERE SearchID=" & id & " AND PersonID=" & pid
    End If
    AuditTry "Save Search", "Saved search (" & searchName & ")", SearchSummary(crit) & "."
End Function

' Delete one of the signed-in person's saved searches, and audit it.
' Returns "" when deleted, else what to tell them (not theirs, or gone).
' Errors: the engine's and KitExecute's.
Public Function DeleteSavedSearch(ByVal searchId As Long) As String
    Dim nm As String
    If Not DataLinked() Then Exit Function
    nm = Nz(DLookup("SearchName", "tblSavedSearches", "SearchID=" & searchId & " AND PersonID=" & SignedInPersonID()), "")
    If Len(nm) = 0 Then DeleteSavedSearch = "That saved search is no longer there.": Exit Function
    If KitExecute("DELETE FROM tblSavedSearches WHERE SearchID=" & searchId & " AND PersonID=" & SignedInPersonID()) = 0 Then
        DeleteSavedSearch = "That saved search is no longer there."
        Exit Function
    End If
    AuditTry "Delete Search", "Saved search (" & nm & ")", "Saved search deleted."
End Function

' ---------------------------------------------------------------------------
' Criteria text escaping
' ---------------------------------------------------------------------------

Private Function CritEscape(ByVal s As String) As String
    s = Replace(s, "%", "%25")
    s = Replace(s, ";", "%3B")
    s = Replace(s, "=", "%3D")
    s = Replace(s, vbCr, "%0D")
    CritEscape = Replace(s, vbLf, "%0A")
End Function

Private Function CritUnescape(ByVal s As String) As String
    s = Replace(s, "%0A", vbLf)
    s = Replace(s, "%0D", vbCr)
    s = Replace(s, "%3D", "=")
    s = Replace(s, "%3B", ";")
    CritUnescape = Replace(s, "%25", "%")
End Function
