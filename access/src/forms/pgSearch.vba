Option Compare Database
Option Explicit

' Search (see modBuildSearch.Build_pgSearch). The segmented switch picks
' Basic (words only), Advanced (words and filters) or Saved (the person's
' saved searches). Search, or Enter in the words or reference box, runs the
' search (modSearch.SearchSql over modMessages' shared list columns) and
' audits it (modSearch.AuditSearchRun: Search, "Keyword: ..."). Switching
' between Basic and Advanced shows the last results again without running
' a new search. Double-click a result, or Open, to read it on the Message
' page (OpenMessage). Save search asks for a name (KitInputBox) and keeps
' the words, and in Advanced the filters, in tblSavedSearches for the
' signed-in person; Saved lists them to run (double-click or Run) or
' delete. With no data file linked the page says so and searches nothing.

Private Const MSG_TITLE As String = "Search"

Private m_mode As String        ' basic, advanced or saved
Private m_sql As String         ' the last search's record source ("" before the first)
Private m_crit As Object        ' the last search's criteria

Public Property Get HelpTopic() As String
    HelpTopic = "search"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_mode = "basic"
    If DataLinked() Then
        Me.cboOrig.RowSource = "SELECT DISTINCT Originator FROM tblMessages WHERE Originator Is Not Null ORDER BY Originator"
        Me.cboType.RowSource = "SELECT TypeCode FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder, TypeCode"
        Me.cboClass.RowSource = "SELECT ClassCode, ClassCode & ' - ' & ClassName AS cls_txt FROM tblMessageClasses ORDER BY ClassCode"
    End If
    ShowMode
    Exit Sub
Err_Handler:
    HandleError "pgSearch.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' The mode switch (Segmented segMode's onChange)
' ---------------------------------------------------------------------------

Public Sub ModeChanged()
    On Error GoTo Err_Handler
    m_mode = SegmentValue(Me, "segMode")
    ShowMode
    Exit Sub
Err_Handler:
    HandleError "pgSearch.ModeChanged", m_mode
End Sub

' Show the mode's controls (the group labels list them), hiding the others
' first since the modes' cards share the same place on the page.
Private Sub ShowMode()
    ShowGroup "grpBasic", False
    ShowGroup "grpAdv", False
    ShowGroup "grpSaved", False
    ShowEmpty Me, "emptyBasic", False
    ShowEmpty Me, "emptyAdv", False
    ShowEmpty Me, "emptySaved", False
    ShowGroup "grpWords", m_mode <> "saved"
    SetVisible Me.btnSave, m_mode <> "saved" And DataLinked()
    ' A card title's accent bar shows in dark palettes only (the kit's
    ' DarkOnly), so the groups, listed in the build palette, leave it out.
    SetVisible Me.barBasic, IsDark() And m_mode = "basic"
    SetVisible Me.barAdv, IsDark() And m_mode = "advanced"
    SetVisible Me.barSaved, IsDark() And m_mode = "saved"
    Select Case m_mode
        Case "saved"
            ShowGroup "grpSaved", True
            LoadSaved
        Case "advanced"
            ShowGroup "grpAdv", True
            ShowResults "Adv"
        Case Else
            ShowGroup "grpBasic", True
            ShowResults "Basic"
    End Select
End Sub

' Show (or hide) every control a group label lists.
Private Sub ShowGroup(ByVal grp As String, ByVal isShown As Boolean)
    Dim nm As Variant
    For Each nm In Split(Nz(Me.Controls(grp).Caption, ""), ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), isShown
    Next
End Sub

' ---------------------------------------------------------------------------
' Running a search
' ---------------------------------------------------------------------------

Private Sub btnSearch_Click()
    On Error GoTo Err_Handler
    RunSearch
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnSearch_Click"
End Sub

' Enter in the words box runs the search.
Private Sub txtWords_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    If KeyCode = vbKeyReturn Then KeyCode = 0: RunSearch
    Exit Sub
Err_Handler:
    HandleError "pgSearch.txtWords_KeyDown"
End Sub

' Enter in the reference box runs the search.
Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    If KeyCode = vbKeyReturn Then KeyCode = 0: RunSearch
    Exit Sub
Err_Handler:
    HandleError "pgSearch.txtRef_KeyDown"
End Sub

Private Sub btnClear_Click()
    Dim nm As Variant
    On Error GoTo Err_Handler
    For Each nm In Array("txtRef", "cboOrig", "cboType", "cboClass", "cboPrec", "cboClassif", "cboLife", "dtFrom", "dtTo", "cboAck")
        Me.Controls(nm).Value = Null
    Next
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnClear_Click"
End Sub

' The criteria the page's boxes hold now, for mode (basic or advanced).
' Returns Nothing, having said why, when a date can't be read.
Private Function CritFromPage(ByVal mode As String) As Object
    Dim crit As Object, d1 As Variant, d2 As Variant
    Set crit = NewSearchCriteria()
    crit("mode") = mode
    crit("words") = Trim$(FieldText(Me, Me.txtWords))
    If mode = "advanced" Then
        d1 = Me.dtFrom: d2 = Me.dtTo
        If Not IsNull(d1) And Not IsDate(d1) Then KitMsgBox "Type the received-from date as a date, such as 01 Oct 2026.", vbExclamation, MSG_TITLE: Exit Function
        If Not IsNull(d2) And Not IsDate(d2) Then KitMsgBox "Type the received-to date as a date, such as 31 Oct 2026.", vbExclamation, MSG_TITLE: Exit Function
        If IsDate(d1) And IsDate(d2) Then
            If CDate(d1) > CDate(d2) Then KitMsgBox "The received-from date is after the received-to date.", vbExclamation, MSG_TITLE: Exit Function
        End If
        crit("ref") = Trim$(FieldText(Me, Me.txtRef))
        crit("orig") = Nz(Me.cboOrig, "")
        crit("msgtype") = Nz(Me.cboType, "")
        crit("cls") = Nz(Me.cboClass, "")
        crit("prec") = Nz(Me.cboPrec, "")
        crit("classif") = Nz(Me.cboClassif, "")
        crit("life") = Nz(Me.cboLife, "")
        If IsDate(d1) Then crit("from") = Format$(CDate(d1), "yyyy-mm-dd")
        If IsDate(d2) Then crit("to") = Format$(CDate(d2), "yyyy-mm-dd")
        crit("ack") = Nz(Me.cboAck, "")
    End If
    Set CritFromPage = crit
End Function

' Put saved criteria in the page's boxes (every filter, even in Basic, so
' switching to Advanced shows them).
Private Sub CritToPage(ByVal crit As Object)
    Me.txtWords = IIf(Len(crit("words")) > 0, crit("words"), Null)
    Me.txtRef = IIf(Len(crit("ref")) > 0, crit("ref"), Null)
    Me.cboOrig = IIf(Len(crit("orig")) > 0, crit("orig"), Null)
    Me.cboType = IIf(Len(crit("msgtype")) > 0, crit("msgtype"), Null)
    Me.cboClass = IIf(Len(crit("cls")) > 0, crit("cls"), Null)
    Me.cboPrec = IIf(Len(crit("prec")) > 0, crit("prec"), Null)
    Me.cboClassif = IIf(Len(crit("classif")) > 0, crit("classif"), Null)
    Me.cboLife = IIf(Len(crit("life")) > 0, crit("life"), Null)
    Me.dtFrom = Null: Me.dtTo = Null
    If IsDate(crit("from")) Then Me.dtFrom = CDate(crit("from"))
    If IsDate(crit("to")) Then Me.dtTo = CDate(crit("to"))
    Me.cboAck = IIf(Len(crit("ack")) > 0, crit("ack"), Null)
End Sub

' Run the search the page's boxes hold in the current mode, show the
' results and audit the run.
Private Sub RunSearch()
    Dim crit As Object
    If m_mode = "saved" Then Exit Sub
    If Not DataLinked() Then ShowResults ResultsCardName(): Exit Sub
    Set crit = CritFromPage(m_mode)
    If crit Is Nothing Then Exit Sub
    RunCriteria crit
End Sub

' Run criteria crit, show the results in the current mode's card and audit it.
Private Sub RunCriteria(ByVal crit As Object)
    Set m_crit = crit
    m_sql = SearchSql(crit)
    ShowResults ResultsCardName()
    AuditSearchRun crit, ResultCount(Me.Controls("sf" & ResultsCardName()))
End Sub

' The Results card the mode shows: Adv or Basic.
Private Function ResultsCardName() As String
    If m_mode = "advanced" Then ResultsCardName = "Adv" Else ResultsCardName = "Basic"
End Function

' How many rows table sf shows.
Private Function ResultCount(sf As Control) As Long
    Dim rs As DAO.Recordset
    Set rs = sf.Form.RecordsetClone
    If rs.RecordCount > 0 Then rs.MoveLast
    ResultCount = rs.RecordCount
End Function

' Show the last search's results in Results card nm (Basic or Adv): its
' count in the card's title, the table, or the empty state saying why there
' is nothing (no data file, no search yet, no match).
Private Sub ShowResults(ByVal nm As String)
    Dim sf As Control, n As Long, terms As Variant, fresh As Boolean
    Set sf = Me.Controls("sf" & nm)
    If Not DataLinked() Then
        ShowEmpty Me, "empty" & nm, True, "No data file linked", "This copy was built without the data file, so there are no messages to search."
        SetVisible sf, False
        SetEnabled Me.btnSearch, False
        SetEnabled Me.Controls("btnOpen" & nm), False
        Exit Sub
    End If
    If Len(m_sql) = 0 Then
        ShowEmpty Me, "empty" & nm, True, "Search ANCHOR's messages", _
            "Type words and click Search. Every word must match somewhere: the subject, reference, originator or text."
        SetVisible sf, False
        SetEnabled Me.Controls("btnOpen" & nm), False
        Me.Controls("lbl" & nm & "Hdr").Caption = "Results"
        Exit Sub
    End If
    If sf.Form.RecordSource <> m_sql Then sf.Form.RecordSource = m_sql: fresh = True
    ' The marks show the words only when there is one term (they mark one
    ' piece of text).
    terms = SearchTerms(m_crit("words"))
    If UBound(terms) = 0 Then GridHighlight sf, CStr(terms(0)) Else GridHighlight sf, ""
    n = ResultCount(sf)
    Me.Controls("lbl" & nm & "Hdr").Caption = "Results (" & n & ")"
    ShowEmpty Me, "empty" & nm, n = 0, "No messages match", _
        IIf(m_mode = "advanced", "Try fewer words or filters, or a wider date range.", "Try fewer words, or other words.")
    SetVisible sf, n > 0
    SetEnabled Me.Controls("btnOpen" & nm), n > 0
    ' New rows: the first is picked (a kept pick could name a row no longer listed).
    If n > 0 And (fresh Or IsNull(sf.Form.txtSel)) Then SelectFirstRow sf
End Sub

' ---------------------------------------------------------------------------
' Results and saved rows
' ---------------------------------------------------------------------------

Private Sub btnOpenBasic_Click()
    On Error GoTo Err_Handler
    OpenMessage Nz(Me.sfBasic.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnOpenBasic_Click"
End Sub

Private Sub btnOpenAdv_Click()
    On Error GoTo Err_Handler
    OpenMessage Nz(Me.sfAdv.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnOpenAdv_Click"
End Sub

' A row clicked (the tables call this): a double-click on a result opens
' it; on a saved search runs it.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If Not dbl Then Exit Sub
    Select Case grid
        Case "sfBasic", "sfAdv": OpenMessage Nz(Me.Controls(grid).Form.txtSel, 0)
        Case "sfSaved": RunSaved
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgSearch.GridClick", grid
End Sub

' ---------------------------------------------------------------------------
' Saved searches
' ---------------------------------------------------------------------------

Private Sub btnSave_Click()
    Dim crit As Object, answer As Variant, nm As String, problem As String
    On Error GoTo Err_Handler
    If m_mode = "saved" Then Exit Sub
    Set crit = CritFromPage(m_mode)
    If crit Is Nothing Then Exit Sub
    If Len(crit("words")) = 0 And Not SearchHasFilters(crit) Then
        KitMsgBox "Type some words" & IIf(m_mode = "advanced", " or set a filter", "") & " first: there is nothing to save yet.", _
            vbInformation, MSG_TITLE
        Exit Sub
    End If
    answer = KitInputBox("Name this search. Saving under a name you already use replaces that search.", "Save search", _
        Left$(crit("words"), 100))
    If IsNull(answer) Then Exit Sub
    nm = Trim$(answer)
    If Len(nm) > 0 And SavedSearchID(nm) <> 0 Then
        If KitMsgBox("You already have a search called """ & ClipText(nm, 60) & """. Replace it?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    End If
    problem = SaveSearch(nm, crit)
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    KitToast "Saved the search """ & ClipText(nm, 60) & """.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnSave_Click"
End Sub

' List the person's saved searches, or say why there are none.
Private Sub LoadSaved()
    Dim n As Long
    If Not DataLinked() Then
        ShowEmpty Me, "emptySaved", True, "No data file linked", "This copy was built without the data file, so it has no saved searches."
        SetVisible Me.sfSaved, False
        SetEnabled Me.btnRun, False
        SetEnabled Me.btnDelete, False
        Exit Sub
    End If
    Me.sfSaved.Form.RecordSource = SavedSearchesSql()
    n = ResultCount(Me.sfSaved)
    Me.lblSavedHdr.Caption = "Saved searches (" & n & ")"
    ShowEmpty Me, "emptySaved", n = 0
    SetVisible Me.sfSaved, n > 0
    SetEnabled Me.btnRun, n > 0
    SetEnabled Me.btnDelete, n > 0
    If n > 0 Then SelectFirstRow Me.sfSaved
End Sub

Private Sub btnRun_Click()
    On Error GoTo Err_Handler
    RunSaved
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnRun_Click"
End Sub

' Run the picked saved search: its words and filters go in the boxes, the
' switch moves to Basic or Advanced (as it was saved), and it runs.
Private Sub RunSaved()
    Dim id As Long, txt As Variant, crit As Object
    id = Nz(Me.sfSaved.Form.txtSel, 0)
    If id = 0 Then Exit Sub
    txt = DLookup("Criteria", "tblSavedSearches", "SearchID=" & id & " AND PersonID=" & SignedInPersonID())
    If IsNull(txt) Then
        KitMsgBox "That saved search is no longer there.", vbInformation, MSG_TITLE
        LoadSaved
        Exit Sub
    End If
    Set crit = SearchCriteriaFromText(txt)
    CritToPage crit
    m_mode = crit("mode")
    SetSegment Me, "segMode", m_mode
    m_sql = ""                              ' so the mode's card shows only this run
    ShowMode
    RunCriteria crit
End Sub

Private Sub btnDelete_Click()
    Dim id As Long, nm As String, problem As String
    On Error GoTo Err_Handler
    id = Nz(Me.sfSaved.Form.txtSel, 0)
    If id = 0 Then Exit Sub
    nm = Nz(DLookup("SearchName", "tblSavedSearches", "SearchID=" & id), "")
    If KitMsgBox("Delete the saved search """ & ClipText(nm, 60) & """?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    problem = DeleteSavedSearch(id)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE Else KitToast "Deleted the saved search.", "success"
    LoadSaved
    Exit Sub
Err_Handler:
    HandleError "pgSearch.btnDelete_Click", id
End Sub
