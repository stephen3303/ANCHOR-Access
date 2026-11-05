Option Compare Database
Option Explicit

' Access UI kit - in-app help: a help window (dlgHelp, built by
' BuildHelpDialog) that opens on the topic for whatever the user is looking
' at, with a topic list and a search box. F1 opens it (BuildHelpKey writes
' the AutoKeys macro); a sidebar link or button opens it with =KitHelp().
'
' The topics are written in a text file (the app's src\help\help.txt) that
' the build loads into the local table zz_help (LoadHelpTopics), so the
' front end needs no files at run time. The format:
'   === key | Title      starts a topic. A key with a dot (detail.roles) is
'                        listed indented under its parent. Page keys match
'                        frmMain's PageKey.
'   # Heading            a heading inside the topic
'   - text               a bullet point
'   **text**             bold
'   Other lines are paragraph text; a blank line ends a paragraph.
'   Lines above the first topic are notes for the author and are ignored.
'
' Which topic F1 opens: the first of these that names a topic in zz_help.
'   1. The active form's HelpTopic, when it has one (a Public Property Get
'      HelpTopic() As String in its code-behind; dlgHelp's is its own topic).
'   2. frmMain's HelpTopic, so the app can name the page and tab in one place.
'   3. The HelpTopic of the page shown in frmMain's sfPage.
'   4. frmMain's PageKey.
'   5. The first topic in the file.
' So most apps only need topic keys that match their page keys; a dialog or
' a page with tabs adds a HelpTopic property to be more precise.
'
' Help an admin edits in the app, with no rebuild (optional): the app keeps
' the topics in a table in its shared data file (help_topic, made with
' HelpSharedTableSql) and links it. At start-up, once the data file is found,
' KitHelpSync adds the build's topics the shared table lacks (all of them the
' first time) and copies the shared table into zz_help, so F1 shows the
' edited text. The app's help page saves a topic with KitHelpSave, which
' writes the shared table and zz_help at once; other people see the edit at
' their next start-up. help.txt only seeds a topic: once it is in the shared
' table the admin's text wins, and a newer build's wording for it does not
' replace it (delete the topic's row to take help.txt's again at the next
' start-up). Without the linked table both only touch zz_help.

Public Const HELP_TABLE As String = "zz_help"
Public Const HELP_SHARED_TABLE As String = "help_topic"
Private Const HELP_KEY_LEN As Long = 32
Private Const HELP_TITLE_LEN As Long = 80

Private m_topic As String       ' the topic dlgHelp shows
Private m_term As String        ' dlgHelp's search text

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step: (re)create zz_help from the help file at path, one row per
' topic, in file order. Raises error 5 for a topic line without "|", an
' empty, long or repeated key, or a long title, naming the line.
Public Sub LoadHelpTopics(ByVal path As String)
    Dim db As DAO.Database, rs As DAO.Recordset, lines As Variant, i As Long, s As String
    Dim key As String, title As String, body As String, n As Long, p As Long, seen As Object
    Set db = CurrentDb
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1                         ' keys match ignoring case, like the primary key
    lines = Split(Replace(ReadTextFile(path), vbCrLf, vbLf), vbLf)
    On Error Resume Next
    db.TableDefs.Delete HELP_TABLE
    On Error GoTo 0
    db.Execute "CREATE TABLE " & HELP_TABLE & " (topic_key TEXT(" & HELP_KEY_LEN & ") CONSTRAINT pk_" & HELP_TABLE & _
        " PRIMARY KEY, title TEXT(" & HELP_TITLE_LEN & "), sort_order LONG, body MEMO, search_text MEMO)", dbFailOnError
    Set rs = db.OpenRecordset(HELP_TABLE, dbOpenDynaset)
    For i = 0 To UBound(lines)
        s = CStr(lines(i))
        If Left$(s, 4) = "=== " Then
            If n > 0 Then AddTopic rs, key, title, n, body
            p = InStr(s, "|")
            If p = 0 Then BadLine i, "has no | between the key and the title", s
            key = Trim$(Mid$(s, 5, p - 5))
            title = Trim$(Mid$(s, p + 1))
            If Len(key) = 0 Then BadLine i, "has no key", s
            If Len(key) > HELP_KEY_LEN Then BadLine i, "has a key over " & HELP_KEY_LEN & " characters", s
            If Len(title) > HELP_TITLE_LEN Then BadLine i, "has a title over " & HELP_TITLE_LEN & " characters", s
            If seen.Exists(key) Then BadLine i, "repeats the key " & key, s
            seen.Add key, True
            n = n + 1
            body = ""
        ElseIf n > 0 Then
            body = body & s & vbLf
        End If
    Next
    If n > 0 Then AddTopic rs, key, title, n, body
    rs.Close
End Sub

Private Sub AddTopic(rs As DAO.Recordset, ByVal key As String, ByVal title As String, ByVal n As Long, ByVal body As String)
    rs.AddNew
    rs!topic_key = key
    rs!title = title
    rs!sort_order = n
    rs!body = body
    rs!search_text = PlainHelpText(body)
    rs.Update
End Sub

' A topic's text as the reader sees it, for search: without the "# " and
' "- " line markers or the ** around bold, so a search for * or # finds only
' topics that show one.
Private Function PlainHelpText(ByVal body As String) As String
    Dim ln As Variant, s As String, out As String
    For Each ln In Split(body, vbLf)
        s = Trim$(CStr(ln))
        If Left$(s, 2) = "# " Or Left$(s, 2) = "- " Then s = Mid$(s, 3)
        out = out & Replace(s, "**", "") & vbLf
    Next
    PlainHelpText = out
End Function

Private Sub BadLine(ByVal i As Long, ByVal problem As String, ByVal s As String)
    Err.Raise 5, "modKitHelp.LoadHelpTopics", "Help file line " & (i + 1) & " " & problem & ": " & s
End Sub

' ---------------------------------------------------------------------------
' Shared topics (help the admin edits in the app; see the top of the module)
' ---------------------------------------------------------------------------

' CREATE TABLE for the shared help table, for the app's data file script:
' zz_help's columns, plus when and by whom a topic was last changed.
Public Function HelpSharedTableSql(Optional ByVal tableName As String = HELP_SHARED_TABLE) As String
    HelpSharedTableSql = "CREATE TABLE " & tableName & " (topic_key TEXT(" & HELP_KEY_LEN & ") CONSTRAINT pk_" & tableName & _
        " PRIMARY KEY, title TEXT(" & HELP_TITLE_LEN & "), sort_order LONG, body MEMO, search_text MEMO, " & _
        "updated_at DATETIME, updated_by TEXT(" & STAMP_USER_LEN & "))"
End Function

' Start-up, once the data file is found: add the topics zz_help has (the
' build's help.txt) that the shared table lacks, then copy the shared table
' into zz_help. Returns how many topics it added to the shared table; does
' nothing and returns 0 when the app has no shared table or no zz_help.
Public Function KitHelpSync(Optional ByVal tableName As String = HELP_SHARED_TABLE) As Long
    Dim db As DAO.Database, n As Long, e As Long, d As String
    Set db = CurrentDb
    If Not HasHelpTable(db, tableName) Or Not HasHelpTable(db, HELP_TABLE) Then Exit Function
    On Error Resume Next
    db.Execute "INSERT INTO " & tableName & " (topic_key, title, sort_order, body, search_text, updated_at, updated_by) " & _
        "SELECT h.topic_key, h.title, h.sort_order, h.body, h.search_text, Now(), 'help.txt' FROM " & HELP_TABLE & _
        " AS h WHERE h.topic_key NOT IN (SELECT topic_key FROM " & tableName & ")", dbFailOnError
    e = Err.Number: d = Err.Description
    If e = 0 Then n = db.RecordsAffected
    On Error GoTo 0
    ' 3022: someone else's start-up added the same topics a moment earlier.
    If e <> 0 And e <> 3022 Then Err.Raise e, "modKitHelp.KitHelpSync", d
    db.Execute "DELETE FROM " & HELP_TABLE, dbFailOnError
    db.Execute "INSERT INTO " & HELP_TABLE & " (topic_key, title, sort_order, body, search_text) " & _
        "SELECT topic_key, title, sort_order, body, search_text FROM " & tableName, dbFailOnError
    KitHelpSync = n
End Function

' Save topic key from the app's help page: its title and its text in the
' help file's markup, in the shared table (stamped with KitUser and the
' time) and in zz_help, so F1 shows it at once. A new key goes after the
' last topic, or a new key with a dot (exams.new) after the last topic under
' its parent, so the help window lists it there. Raises error 5 for an empty or over-32-character key or an
' over-80-character title.
Public Sub KitHelpSave(ByVal key As String, ByVal title As String, ByVal body As String, _
        Optional ByVal tableName As String = HELP_SHARED_TABLE)
    Dim db As DAO.Database
    key = Trim$(key): title = Trim$(title)
    If Len(key) = 0 Or Len(key) > HELP_KEY_LEN Then
        Err.Raise 5, "modKitHelp.KitHelpSave", "A help topic's key must be 1 to " & HELP_KEY_LEN & " characters: " & key
    End If
    If Len(title) > HELP_TITLE_LEN Then
        Err.Raise 5, "modKitHelp.KitHelpSave", "A help topic's title can be at most " & HELP_TITLE_LEN & " characters: " & title
    End If
    body = Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf)
    Set db = CurrentDb
    If HasHelpTable(db, tableName) Then SaveTopicRow db, tableName, key, title, body, True
    If HasHelpTable(db, HELP_TABLE) Then SaveTopicRow db, HELP_TABLE, key, title, body, False
End Sub

Private Sub SaveTopicRow(db As DAO.Database, ByVal tbl As String, ByVal key As String, ByVal title As String, _
        ByVal body As String, ByVal stamp As Boolean)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM " & tbl & " WHERE topic_key=" & SqlText(key), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!topic_key = key
        rs!sort_order = NewSortOrder(db, tbl, key)
    Else
        rs.Edit
    End If
    rs!title = title
    rs!body = body
    rs!search_text = PlainHelpText(body)
    If stamp Then
        rs!updated_at = Now()
        rs!updated_by = KitUser()
    End If
    rs.Update
    rs.Close
End Sub

' sort_order for new topic key in tbl: after the last topic, or for a key
' with a dot, after its parent and the topics already under it (the topics
' after them move down one).
Private Function NewSortOrder(db As DAO.Database, ByVal tbl As String, ByVal key As String) As Long
    Dim dot As Long, parent As String, after As Variant
    after = Null
    dot = InStr(key, ".")
    If dot > 0 Then
        parent = Left$(key, dot - 1)
        after = DMax("sort_order", tbl, "topic_key=" & SqlText(parent) & " OR topic_key LIKE " & _
            SqlText(LikeText(parent) & ".*"))
    End If
    If IsNull(after) Then
        NewSortOrder = Nz(DMax("sort_order", tbl), 0) + 1
    Else
        db.Execute "UPDATE " & tbl & " SET sort_order=sort_order+1 WHERE sort_order>" & after, dbFailOnError
        NewSortOrder = after + 1
    End If
End Function

' True when the database has table nm (local or linked).
Private Function HasHelpTable(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then HasHelpTable = True: Exit Function
    Next
End Function

' Build step: the AutoKeys macro, so F1 opens the kit's help instead of
' Access's. It replaces any AutoKeys macro the database already has; an app
' that also wants the kit's Ctrl+S, Ctrl+N and Ctrl+F calls
' BuildShortcutKeys (modKitKeys) instead, which keeps F1 for help.
Public Sub BuildHelpKey()
    WriteAutoKeys "Version =196611" & vbCrLf & "ColumnsShown =1" & vbCrLf & "Begin" & vbCrLf & _
        "    MacroName =""{F1}""" & vbCrLf & "    Action =""RunCode""" & vbCrLf & _
        "    Argument =""KitHelp()""" & vbCrLf & "End" & vbCrLf
End Sub

' Build dlgHelp, the help window. Call it once with the app's dialogs.
' 900 x 600px: the header with a help icon and the topic's title; a search
' box over the topic list on the left (240px); the topic as rich text on the
' right; an F1 hint and Close (Esc) along the bottom. Not modal, so it can
' stay open beside the page (KitHelp opens it as a dialog when a modal form
' has the focus).
Public Sub BuildHelpDialog()
    Dim c As Control, w As Long, h As Long, pad As Long, rowY As Long, listY As Long
    Dim btnY As Long, paneBottom As Long, bodyX As Long
    w = px(900): h = px(600)
    pad = Sp(SP_DIALOG)
    rowY = pad + px(28) + Sp(16)                 ' under the header row
    btnY = h - pad - px(30)
    paneBottom = btnY - Sp(16)
    listY = rowY + px(28) + Sp(SP_ROW_GAP)
    bodyX = pad + px(240) + Sp(16)
    NewDialog "Help", w, h
    F.Modal = False
    Ev F, "OnLoad"
    Set c = Lbl("lblIcon", ChrW(&HE897), pad, pad, px(22), px(28), 13, False, CLR_PRIMARY)
    c.FontName = "Segoe MDL2 Assets": c.FontWeight = 400
    Lbl "lblHeader", "Help", pad + px(30), pad, w - 2 * pad - px(30), px(28), 13, True
    Set c = Txt("txtSearch", "", pad, rowY, px(240), px(28), 9)
    c.Format = "@;""Search help..."""
    Ev c, "OnChange", "OnKeyDown"
    Set c = Lst("lstTopics", "", pad, listY, px(240), paneBottom - listY, 2, "0;" & px(220))
    c.ColumnHeads = False: c.FontSize = 10
    Ev c, "AfterUpdate"
    Set c = Lbl("lblNoMatch", "No topics match your search.", pad + px(10), listY + px(10), px(220), px(36), 9, False, CLR_MUTED_FG)
    c.Visible = False
    Set c = Txt("txtBody", "", bodyX, rowY, w - pad - bodyX, paneBottom - rowY, 10)
    c.TextFormat = 1                             ' rich text
    c.Locked = True: c.TabStop = False: c.ScrollBars = 2
    c.OnGotFocus = "": c.OnLostFocus = ""        ' read only, so no focus ring
    c.LeftMargin = px(10): c.TopMargin = px(8): c.RightMargin = px(10)
    Lbl "lblTip", "Press F1 on any page for help with that page.", pad, btnY, px(500), px(30), 8, False, CLR_MUTED_FG
    Set c = Btn("btnClose", "Close", w - pad - px(88), btnY, px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgHelp"
End Sub

' ---------------------------------------------------------------------------
' Run time
' ---------------------------------------------------------------------------

' Open help on topic, or on the topic for what the user is looking at (see
' the top of this module). When help is already open it switches topic.
' Use it as =KitHelp() on a sidebar link, or =KitHelp("reports") for a
' fixed topic. Logs and does nothing when the build loaded no topics.
Public Function KitHelp(Optional ByVal topic As String = "")
    On Error GoTo Fail
    If Not HelpLoaded() Then
        LogError "modKitHelp.KitHelp", "no help topics: the build did not run LoadHelpTopics"
        Exit Function
    End If
    If Len(topic) = 0 Then topic = HelpContextTopic()
    If IsFormOpen("dlgHelp") And ModalFormActive() And ActiveFormName() <> "dlgHelp" Then
        ' Help is open beside the page, but a modal dialog has the focus now,
        ' so the open window can't take it: reopen help as a dialog.
        DoCmd.Close acForm, "dlgHelp", acSaveNo
        DoCmd.OpenForm "dlgHelp", WindowMode:=acDialog, OpenArgs:=topic
    ElseIf IsFormOpen("dlgHelp") Then
        HelpShow Forms("dlgHelp"), topic
        Forms("dlgHelp").SetFocus
    ElseIf ModalFormActive() Then
        ' A window opened over a modal dialog must be a dialog too, or it
        ' can't take the focus.
        DoCmd.OpenForm "dlgHelp", WindowMode:=acDialog, OpenArgs:=topic
    Else
        DoCmd.OpenForm "dlgHelp", OpenArgs:=topic
    End If
    Exit Function
Fail:
    HandleError "modKitHelp.KitHelp", "topic " & topic
End Function

' The topic for what the user is looking at (the order is at the top of
' this module).
Public Function HelpContextTopic() As String
    Dim frm As Object, t As String
    On Error Resume Next
    Set frm = Screen.ActiveForm
    On Error GoTo 0
    If Not frm Is Nothing Then
        t = HelpTopicOf(frm)
        If HelpTopicExists(t) Then HelpContextTopic = t: Exit Function
    End If
    If IsFormOpen("frmMain") Then
        Set frm = Forms("frmMain")
        t = HelpTopicOf(frm)
        If HelpTopicExists(t) Then HelpContextTopic = t: Exit Function
        t = HelpTopicOf(PageForm(frm))
        If HelpTopicExists(t) Then HelpContextTopic = t: Exit Function
        t = AskForm(frm, "PageKey")
        If HelpTopicExists(t) Then HelpContextTopic = t: Exit Function
    End If
    HelpContextTopic = FirstTopic()
End Function

' A form's HelpTopic property, or "" when it has none.
Public Function HelpTopicOf(frm As Object) As String
    HelpTopicOf = AskForm(frm, "HelpTopic")
End Function

' True when zz_help has topic key.
Public Function HelpTopicExists(ByVal key As String) As Boolean
    If Len(key) = 0 Then Exit Function
    HelpTopicExists = Not IsNull(DLookup("topic_key", HELP_TABLE, "topic_key=" & SqlText(key)))
End Function

' Row source for the topic list: the key, then the title (indented under its
' parent). term narrows it to topics whose title or text contain it,
' literally and ignoring case. The text searched is search_text, which
' leaves out the help file's markup (PlainHelpText).
Public Function HelpTopicsSql(Optional ByVal term As String = "") As String
    Dim where As String, t As String
    term = Trim$(term)
    If Len(term) > 0 Then
        t = SqlText("*" & LikeText(term) & "*")
        where = " WHERE title LIKE " & t & " OR search_text LIKE " & t
    End If
    HelpTopicsSql = "SELECT topic_key, IIf(InStr(topic_key,'.')>0,'      ','') & title FROM " & HELP_TABLE & _
        where & " ORDER BY sort_order"
End Function

' Heading for a topic: a subtopic's is prefixed with its parent's title.
Public Function HelpTitleFor(ByVal topic As String) As String
    Dim t As String, dot As Long
    t = Nz(DLookup("title", HELP_TABLE, "topic_key=" & SqlText(topic)), "")
    dot = InStr(topic, ".")
    If dot > 0 Then
        t = Nz(DLookup("title", HELP_TABLE, "topic_key=" & SqlText(Left$(topic, dot - 1))), "") & "  >  " & t
    End If
    HelpTitleFor = t
End Function

' The topic as rich text for a text box with TextFormat = rich text: "# "
' lines are headings, "- " lines bullets, **text** is bold, and other lines
' join into paragraphs that end at a blank line. Each match of term is shown
' in the search-match colors (GridMark).
Public Function HelpBodyHtml(ByVal topic As String, Optional ByVal term As String = "") As String
    Dim lines As Variant, ln As Variant, s As String, out As String, para As String, inList As Boolean
    term = Trim$(term)
    lines = Split(Nz(DLookup("body", HELP_TABLE, "topic_key=" & SqlText(topic)), "") & vbLf, vbLf)
    For Each ln In lines
        s = Trim$(CStr(ln))
        If Left$(s, 2) = "- " Or Left$(s, 2) = "# " Or Len(s) = 0 Then
            If Len(para) > 0 Then out = out & "<div>" & Inline(para, term) & "</div><div>&nbsp;</div>": para = ""
        End If
        If Left$(s, 2) = "- " Then
            If Not inList Then out = out & "<ul>": inList = True
            out = out & "<li>" & Inline(Mid$(s, 3), term) & "</li>"
        Else
            If inList Then out = out & "</ul><div>&nbsp;</div>": inList = False
            If Left$(s, 2) = "# " Then
                out = out & "<div><font size=4><strong>" & Inline(Mid$(s, 3), term) & "</strong></font></div>"
            ElseIf Len(s) > 0 Then
                para = para & IIf(Len(para) > 0, " ", "") & s
            End If
        End If
    Next
    ' Drop the trailing spacer.
    If Right$(out, 17) = "<div>&nbsp;</div>" Then out = Left$(out, Len(out) - 17)
    HelpBodyHtml = out
End Function

' dlgHelp's Load: list every topic and show topic (the first one when it is
' empty or unknown).
Public Sub HelpOpen(frm As Form, ByVal topic As String)
    m_term = ""
    frm.Caption = APP_NAME & " Help"
    frm.lstTopics.RowSource = HelpTopicsSql()
    SetVisible frm.lstTopics, True
    SetVisible frm.lblNoMatch, False
    HelpShow frm, topic
    CenterOnApp frm
End Sub

' Show a topic in dlgHelp and select it in the list. An unknown key shows
' the first topic.
Public Sub HelpShow(frm As Form, ByVal topic As String)
    If Not HelpTopicExists(topic) Then topic = FirstTopic()
    m_topic = topic
    frm.lstTopics = topic
    frm.lblHeader.Caption = Replace(HelpTitleFor(topic), "&", "&&")
    frm.txtBody = HelpBodyHtml(topic, m_term)
End Sub

' dlgHelp's search box changed: narrow the list to the topics that contain
' term and mark term in the topic shown. When the shown topic isn't among
' them, the first one that is is shown instead. When none are, the empty
' list is hidden (it would paint over the note) and the note says so.
Public Sub HelpSearch(frm As Form, ByVal term As String)
    Dim n As Long
    m_term = Trim$(term)
    frm.lstTopics.RowSource = HelpTopicsSql(m_term)
    n = frm.lstTopics.ListCount
    SetVisible frm.lstTopics, n > 0
    SetVisible frm.lblNoMatch, n = 0
    If n > 0 And Not Listed(frm.lstTopics, m_topic) Then
        HelpShow frm, frm.lstTopics.ItemData(0)
    Else
        frm.lstTopics = m_topic
        frm.txtBody = HelpBodyHtml(m_topic, m_term)
    End If
End Sub

' Enter in dlgHelp's search box: show the first topic the search found.
Public Sub HelpShowFirstMatch(frm As Form)
    If frm.lstTopics.ListCount > 0 Then HelpShow frm, frm.lstTopics.ItemData(0)
End Sub

' The topic dlgHelp shows ("" before it opens).
Public Function HelpCurrentTopic() As String
    HelpCurrentTopic = m_topic
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' True when list box lst has a row with key.
Private Function Listed(lst As Control, ByVal key As String) As Boolean
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.ItemData(i) = key Then Listed = True: Exit Function
    Next
End Function

Private Function HelpLoaded() As Boolean
    On Error Resume Next
    HelpLoaded = DCount("*", HELP_TABLE) > 0
End Function

Private Function FirstTopic() As String
    FirstTopic = Nz(DLookup("topic_key", HELP_TABLE, "sort_order=" & Nz(DMin("sort_order", HELP_TABLE), 0)), "")
End Function

Private Function ActiveFormName() As String
    On Error Resume Next
    ActiveFormName = Screen.ActiveForm.Name
End Function

Private Function ModalFormActive() As Boolean
    On Error Resume Next
    ModalFormActive = Screen.ActiveForm.Modal
End Function

' The page shown in frm's sfPage, or Nothing.
Private Function PageForm(frm As Object) As Object
    On Error Resume Next
    Set PageForm = frm.sfPage.Form
End Function

' A form's property as text, or "" when the form is Nothing or has no such
' property (438; 2465 for a missing sfPage). Any other error in the form's
' own property is logged and also counts as "".
Private Function AskForm(frm As Object, ByVal member As String) As String
    Dim v As Variant
    If frm Is Nothing Then Exit Function
    On Error GoTo Fail
    v = CallByName(frm, member, VbGet)
    AskForm = Nz(v, "")
    Exit Function
Fail:
    Select Case Err.Number
        Case 438, 2465
        Case Else: LogError "modKitHelp.AskForm", member
    End Select
End Function

' Escape Access LIKE's wildcards, so the text matches literally.
Private Function LikeText(ByVal s As String) As String
    s = Replace(s, "[", "[[]")
    s = Replace(s, "*", "[*]")
    s = Replace(s, "?", "[?]")
    LikeText = Replace(s, "#", "[#]")
End Function

' Text as a quoted SQL literal.
Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function

' Escape HTML and mark term (GridMark), and turn **text** into bold.
' GridMark treats its text as the start of a cell and keeps a leading space
' as &nbsp;, so each part after the first is passed behind a placeholder
' character (never typed, so never matched) that is then dropped.
Private Function Inline(ByVal s As String, ByVal term As String) As String
    Dim parts As Variant, i As Long, out As String, part As String
    parts = Split(s, "**")
    For i = 0 To UBound(parts)
        If i = 0 Then
            part = Nz(GridMark(parts(i), term), "")
        Else
            part = Mid$(Nz(GridMark(ChrW(1) & parts(i), term), ""), 2)
        End If
        If i Mod 2 = 1 Then out = out & "<strong>" & part & "</strong>" Else out = out & part
    Next
    Inline = out
End Function
