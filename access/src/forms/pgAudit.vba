Option Compare Database
Option Explicit

' Audit (see modBuildAudit.Build_pgAudit), for System Admins: the audit
' trail, tblAudit, which the app only ever adds to (modAudit).
'   Tiles     how many entries the trail holds, how many failed or warned,
'             and the chain's status as the last Verify on this page found it.
'   Verify    recomputes every entry's hash from the first (AuditVerify): the
'             alert under the tiles says "All N entries verified", or names
'             the first entry that doesn't fit. The check is audited.
'   Filters   search (resource and details), action, result, user and a day
'             range; typing in the search box waits for a pause (the form's
'             Timer). The list shows the newest AUD_MAX entries that fit,
'             oldest or newest first (segOrder), and the subtitle says when
'             there are more.
'   Entry     the picked entry in full, every column, the hashes included,
'             with Copy.
'   Export    the entries the filters match (not only those shown) to a CSV
'             file, every column; the export is audited.
' Classifications above UNCLASSIFIED carry the demo marking note
' (modMessages.ClassificationMarking) wherever they show.
' With no data file linked, or for someone who isn't a System Admin, the
' page says so and its buttons are off.
'
' Owner: the Audit page's build task. Depends on: modAudit, modErrors,
' modGlobals, modMessages, modSignIn, modUI; the kit's modCsv, modKitFiles,
' modKitUI, modKitUtil.

Private Const MSG_TITLE As String = "Audit"
Private Const AUD_MAX As Long = 1000            ' entries the list shows at most
Private Const SEARCH_WAIT_MS As Long = 300      ' the pause in typing before the list is loaded again
Private Const ALL_KEY As String = "*"           ' the filter boxes' "every" row

Private m_ok As Boolean         ' a data file is linked and the signed-in person is a System Admin
Private m_search As String
Private m_aid As Long           ' the entry shown (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "audit"
End Property

Private Sub Form_Load()
    Dim nm As Variant
    On Error GoTo Err_Handler
    HideAlert Me, "alVerify"
    m_ok = DataLinked()
    If m_ok Then m_ok = IsAnchorAdmin()
    SetSegment Me, "segOrder", "new"
    If Not m_ok Then
        For Each nm In Array("btnVerify", "btnExport", "btnCopy", "cboAction", "cboResult", "cboUser", "dtFrom", "dtTo", _
                "cal_dtFrom", "cal_dtTo", "txtSearch")
            SetEnabled Me.Controls(nm), False
        Next
        Me.segOrder_val.Locked = True       ' SegmentClick ignores a locked value
        For Each nm In Array("tileEntries", "tileFailed", "tileWarnings", "tileChain")
            SetTile CStr(nm), Null, ""
        Next
        ShowHeaders False
        SetVisible Me.sfList, False
        ShowEmpty Me, "emptyList", True, IIf(DataLinked(), "For System Admins only", "No data file linked"), _
            IIf(DataLinked(), "Only a System Admin can read the audit trail.", "This copy was built without the data file, so there is no audit trail.")
        SetVisible Me.emptyList_b, False
        ShowEntry 0
        Exit Sub
    End If
    Me.cboAction.RowSource = "SELECT " & Q(ALL_KEY) & " AS k, '(Every action)' AS nm, 0 AS o FROM tblSettings WHERE SettingKey='SchemaVersion' " & _
        "UNION SELECT Action, Action, 1 FROM tblAudit ORDER BY o, nm"
    Me.cboUser.RowSource = "SELECT " & Q(ALL_KEY) & " AS k, '(Every user)' AS nm, 0 AS o FROM tblSettings WHERE SettingKey='SchemaVersion' " & _
        "UNION SELECT UserName, UserName, 1 FROM tblAudit ORDER BY o, nm"
    Me.cboResult.RowSourceType = "Value List"
    Me.cboResult.RowSource = ALL_KEY & ";(Every result);Success;Success;Warning;Warning;Failed;Failed"
    Me.cboAction = ALL_KEY: Me.cboUser = ALL_KEY: Me.cboResult = ALL_KEY
    SetTile "tileChain", "Not checked", "Verify chain to check it"
    RefreshAll
    Exit Sub
Err_Handler:
    HandleError "pgAudit.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Filters (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    Me.TimerInterval = 0                    ' restart the wait at each keystroke
    Me.TimerInterval = SEARCH_WAIT_MS       ' Form_Timer loads the list once the typing pauses
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgAudit.SearchFor", text
End Sub

Private Sub Form_Timer()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.Form_Timer"
End Sub

' The order's Segmented control (onChange).
Public Sub OrderChanged()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.OrderChanged"
End Sub

Private Sub cboAction_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.cboAction_AfterUpdate"
End Sub

Private Sub cboResult_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.cboResult_AfterUpdate"
End Sub

Private Sub cboUser_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.cboUser_AfterUpdate"
End Sub

Private Sub dtFrom_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.dtFrom_AfterUpdate"
End Sub

Private Sub dtTo_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.dtTo_AfterUpdate"
End Sub

' The calendar buttons: the kit's picker sets the box without its
' AfterUpdate, so the list is loaded again here when the date changed.
Private Sub cal_dtFrom_Click()
    Dim was As Variant
    On Error GoTo Err_Handler
    was = Me.dtFrom
    PickDate Me, "dtFrom"
    If Nz(Me.dtFrom, 0) <> Nz(was, 0) Then RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.cal_dtFrom_Click"
End Sub

Private Sub cal_dtTo_Click()
    Dim was As Variant
    On Error GoTo Err_Handler
    was = Me.dtTo
    PickDate Me, "dtTo"
    If Nz(Me.dtTo, 0) <> Nz(was, 0) Then RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.cal_dtTo_Click"
End Sub

' The empty state's button: every entry again.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    Me.cboAction = ALL_KEY: Me.cboUser = ALL_KEY: Me.cboResult = ALL_KEY
    Me.dtFrom = Null: Me.dtTo = Null
    SetSearch Me, "txtSearch", ""           ' calls SearchFor, which sets m_search
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAudit.emptyList_b_Click"
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    ShowEntry Nz(Me.sfList.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgAudit.GridClick", grid
End Sub

' Verify chain: recompute every entry's hash (modAudit.AuditVerify), show the
' result in the alert and the Chain tile, and audit the check.
Private Sub btnVerify_Click()
    Dim n As Long, problem As String
    On Error GoTo Err_Handler
    RequireAnchorAdmin
    DoCmd.Hourglass True
    problem = AuditVerify(n)
    DoCmd.Hourglass False
    If Len(problem) = 0 And n = 0 Then
        ShowAlert Me, "alVerify", "The audit trail has no entries yet, so there is nothing to verify.", "info"
        SetTile "tileChain", "Empty", "Checked " & Format$(Now(), "hh:nn")
    ElseIf Len(problem) = 0 Then
        ShowAlert Me, "alVerify", "All " & Format$(n, "#,##0") & " " & IIf(n = 1, "entry", "entries") & " verified.", "success"
        SetTile "tileChain", "Intact", "Checked " & Format$(Now(), "hh:nn") & ", " & Format$(n, "#,##0") & IIf(n = 1, " entry", " entries")
    Else
        ShowAlert Me, "alVerify", ClipText(problem, 160), "destructive"
        SetTile "tileChain", "Broken", "Checked " & Format$(Now(), "hh:nn") & ": see the alert"
    End If
    AuditTry "Integrity Check", "Audit chain (" & n & IIf(n = 1, " entry)", " entries)"), _
        IIf(Len(problem) = 0, "Chain intact. " & n & " entries verified.", problem), , IIf(Len(problem) = 0, "Success", "Failed")
    RefreshAll
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "pgAudit.btnVerify_Click"
End Sub

' Export: the entries the filters match, every column, to a CSV file. The
' export is audited, after it was written.
Private Sub btnExport_Click()
    Dim path As String, n As Long, fileName As String
    On Error GoTo Err_Handler
    RequireAnchorAdmin                      ' the trail names every user and computer
    path = KitSaveAs("Export the audit trail", "Audit trail " & Format$(Date, "yyyy-mm-dd") & ".csv", "CSV file|*.csv", "export")
    If Len(path) = 0 Then Exit Sub
    DoCmd.Hourglass True
    n = ExportSqlToCsv(ExportSql(), path)
    DoCmd.Hourglass False
    fileName = Mid$(path, InStrRev(path, "\") + 1)
    AuditLog "Export", "Audit trail (" & n & IIf(n = 1, " entry)", " entries)"), "Exported to " & fileName & _
        IIf(Len(FilterText()) > 0, ", filtered by " & FilterText() & ".", ", every entry.")
    KitToast "Exported " & n & IIf(n = 1, " entry", " entries") & " to " & fileName, "success"
    RefreshAll
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "pgAudit.btnExport_Click"
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If Len(Nz(Me.txtDetail, "")) = 0 Then Exit Sub
    If KitCopyText(Me.txtDetail) Then
        KitToast "Copied", "success"
    Else
        KitMsgBox "Another program is using the clipboard. Try Copy again.", vbInformation, MSG_TITLE
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAudit.btnCopy_Click"
End Sub

' ---------------------------------------------------------------------------
' The tiles, the list and the entry
' ---------------------------------------------------------------------------

' The tiles and the list (after a Verify or an Export added an entry).
Private Sub RefreshAll()
    Dim n As Long
    n = DCount("*", "tblAudit")
    SetTile "tileEntries", n, IIf(n = 0, "Nothing recorded yet", "Since " & Format$(Nz(DMin("LoggedAt", "tblAudit"), Date), "d mmm yyyy"))
    SetTile "tileFailed", DCount("*", "tblAudit", "Result='Failed'"), "Denied sign-ins, failed checks"
    SetTile "tileWarnings", DCount("*", "tblAudit", "Result='Warning'"), "Done, with a warning"
    Me.cboAction.Requery
    Me.cboUser.Requery
    RefreshList
End Sub

Private Sub SetTile(ByVal nm As String, ByVal v As Variant, ByVal cap As String)
    Me.Controls(nm).Value = v
    Me.Controls(nm & "_cap").Caption = Replace(cap, "&", "&&")
End Sub

Private Sub ShowHeaders(ByVal v As Boolean)
    Dim c As Control
    For Each c In Me.Controls
        If c.Name Like "hdrList#*" Then SetVisible c, v
    Next
End Sub

' The list's WHERE (without the word) for the filters; "True" for none.
Private Function WhereSql() As String
    Dim w As String
    If Nz(Me.cboAction, ALL_KEY) <> ALL_KEY Then w = w & " AND a.Action=" & Q(Me.cboAction.Value)
    If Nz(Me.cboResult, ALL_KEY) <> ALL_KEY Then w = w & " AND a.Result=" & Q(Me.cboResult.Value)
    If Nz(Me.cboUser, ALL_KEY) <> ALL_KEY Then w = w & " AND a.UserName=" & Q(Me.cboUser.Value)
    If IsDate(Me.dtFrom) Then w = w & " AND a.LoggedAt>=" & Format$(DateValue(Me.dtFrom), "\#yyyy\-mm\-dd\#")
    If IsDate(Me.dtTo) Then w = w & " AND a.LoggedAt<" & Format$(DateAdd("d", 1, DateValue(Me.dtTo)), "\#yyyy\-mm\-dd\#")
    If Len(m_search) > 0 Then
        w = w & " AND (Nz(a.Resource,'') & ' ' & Nz(a.Details,'')) LIKE " & Q("*" & LikeLiteral(m_search) & "*")
    End If
    If Len(w) = 0 Then WhereSql = "True" Else WhereSql = Mid$(w, 6)
End Function

' The filters in words, for the export's audit row ("" for none).
Private Function FilterText() As String
    Dim s As String
    If Nz(Me.cboAction, ALL_KEY) <> ALL_KEY Then s = s & "; action " & Me.cboAction
    If Nz(Me.cboResult, ALL_KEY) <> ALL_KEY Then s = s & "; result " & Me.cboResult
    If Nz(Me.cboUser, ALL_KEY) <> ALL_KEY Then s = s & "; user " & Me.cboUser
    If IsDate(Me.dtFrom) Then s = s & "; from " & Format$(Me.dtFrom, "d mmm yyyy")
    If IsDate(Me.dtTo) Then s = s & "; to " & Format$(Me.dtTo, "d mmm yyyy")
    If Len(m_search) > 0 Then s = s & "; search """ & m_search & """"
    FilterText = Mid$(s, 3)
End Function

' The list: the newest AUD_MAX entries the filters match, in the order picked.
Private Function ListSql() As String
    ListSql = "SELECT t.* FROM (SELECT TOP " & AUD_MAX & " a.AuditID AS id, a.AuditID, Format(a.LoggedAt,'dd mmm yy hh:nn') AS when_txt, " & _
        "a.LoggedAt, a.UserName, a.Action, a.Resource, a.Result FROM tblAudit AS a WHERE " & WhereSql() & _
        " ORDER BY a.AuditID DESC) AS t ORDER BY t.id" & IIf(SegmentValue(Me, "segOrder") = "old", "", " DESC")
End Function

' The export: every column of every entry the filters match, oldest first.
' The classification carries the demo marking note above UNCLASSIFIED.
Private Function ExportSql() As String
    ExportSql = "SELECT a.AuditID, a.LoggedAt, a.PersonID, a.UserName, a.Action, a.Resource, " & _
        ClassificationMarkingSql("a.Classification") & " AS ClassMarking, a.Result, a.Device, a.Location, a.Details, a.PrevHash, a.RowHash " & _
        "FROM tblAudit AS a WHERE " & WhereSql() & " ORDER BY a.AuditID"
End Function

' The rows the list shows now.
Private Function ShownCount() As Long
    Dim rs As DAO.Recordset
    Set rs = Me.sfList.Form.RecordsetClone
    If rs.BOF And rs.EOF Then Exit Function
    rs.MoveLast
    ShownCount = rs.RecordCount
End Function

' Load the list for the filters, keep the entry shown when still listed (else
' the first), and say how many there are.
Private Sub RefreshList()
    Dim n As Long, total As Long, rs As DAO.Recordset, hasAny As Boolean
    Me.TimerInterval = 0                    ' a search still waiting is loaded now
    If Not m_ok Then Exit Sub
    If IsDate(Me.dtFrom) And IsDate(Me.dtTo) Then
        If DateValue(Me.dtFrom) > DateValue(Me.dtTo) Then
            Me.sfList.Form.RecordSource = "SELECT a.AuditID AS id FROM tblAudit AS a WHERE False"
            Me.lblSubtitle.Caption = "The first date is after the second: swap them to see entries."
            AfterLoad 0, True
            Exit Sub
        End If
    End If
    Me.sfList.Form.RecordSource = ListSql()
    GridHighlight Me.sfList, m_search
    n = ShownCount()
    If n < AUD_MAX Then
        total = n
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblAudit AS a WHERE " & WhereSql(), dbOpenSnapshot)
        total = Nz(rs.Fields(0).Value, 0)
        rs.Close
    End If
    If total > n Then
        Me.lblSubtitle.Caption = "Showing the newest " & Format$(n, "#,##0") & " of " & Format$(total, "#,##0") & _
            " entries that match. Narrow the filters to see older ones."
    Else
        Me.lblSubtitle.Caption = Format$(n, "#,##0") & IIf(n = 1, " entry", " entries") & IIf(WhereSql() = "True", "", " match the filters")
    End If
    hasAny = (n > 0)
    If Not hasAny Then hasAny = Not IsNull(DLookup("AuditID", "tblAudit"))
    AfterLoad n, hasAny
End Sub

' After the list loaded n rows: the empty state (anyAtAll False: the trail
' itself is empty), and the entry shown.
Private Sub AfterLoad(ByVal n As Long, ByVal anyAtAll As Boolean)
    If n = 0 And Not anyAtAll Then
        ShowEmpty Me, "emptyList", True, "The audit trail is empty", "Nothing has been recorded yet."
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No matching entries", "No entry fits the search, the filters and the dates."
    End If
    SetVisible Me.sfList, n > 0
    ShowHeaders n > 0
    SetEnabled Me.btnExport, n > 0
    If GridSelect(Me.sfList, m_aid) Then
        ShowEntry m_aid
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then  ' SelectFirstRow shows the first row's entry itself (GridClick)
        ShowEntry 0
    End If
End Sub

' Show entry id in full (0: the hint instead).
Private Sub ShowEntry(ByVal id As Long)
    Dim s As String
    If id <> 0 Then s = EntryText(id)
    m_aid = id
    Me.txtDetail = IIf(Len(s) > 0, s, Null)
    SetVisible Me.lblDetailEmpty, Len(s) = 0
    SetEnabled Me.btnCopy, Len(s) > 0 And m_ok
End Sub

' Every column of one entry, as text for the detail box and Copy ("" when
' there is no such entry).
Private Function EntryText(ByVal id As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblAudit WHERE AuditID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        s = "Entry #" & rs!AuditID & "  -  " & Nz(rs!Result, "") & vbCrLf & _
            Nz(rs!Action, "") & ": " & Nz(rs!Resource, "") & vbCrLf & vbCrLf & _
            "When: " & Format$(rs!LoggedAt, "dd mmm yyyy hh:nn:ss") & vbCrLf & _
            "User: " & Nz(rs!UserName, "") & IIf(IsNull(rs!PersonID), "", " (person " & rs!PersonID & ")") & vbCrLf & _
            "Classification: " & ClassificationMarking(rs!Classification.Value) & vbCrLf & _
            "Device: " & Nz(rs!Device, "-") & vbCrLf & _
            "Location: " & Nz(rs!Location, "-") & vbCrLf & _
            "Details: " & Nz(rs!Details, "-") & vbCrLf & vbCrLf & _
            "Previous hash: " & Nz(rs!PrevHash, "(not sealed yet)") & vbCrLf & _
            "Row hash: " & Nz(rs!RowHash, "(not sealed yet)")
    End If
    rs.Close
    EntryText = s
End Function
