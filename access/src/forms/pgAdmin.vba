Option Compare Database
Option Explicit

' Admin (see modBuildAdmin.Build_pgAdmin), for Subscription Approvers and
' System Admins. Three tabs (segTab):
'   Approvals & Access  the requests waiting for a decision, oldest first.
'                       Approve makes the picked one Active; Deny asks why
'                       (the requester sees the reason) and makes it Denied
'                       (modSubscriptions). Under it the mandatory alert rules
'                       and the command access groups, read only.
'   Users               System Admins: add a person from the Personnel file
'                       (they start as Message Viewer and Subscription
'                       Requester), and change a person's billet, unit,
'                       community, location, DSN, roles and access; Save
'                       writes them together (modAdmin.AdmSavePerson).
'   Message types       System Admins: add a type, change its name, category
'                       and description, retire or restore it; the
'                       communications classes, read only.
' A Subscription Approver who isn't a System Admin sees the first tab only,
' with no tabs shown. Every write checks the right in the data file at the
' moment of the write and is audited (modAdmin, modSubscriptions).
'
' Unsaved changes to a person or a type are kept until another is picked or
' the page is left, when the admin is asked whether to save them (CanLeave).
' A change to the signed-in person's own roles shows in the sidebar at once.
' With no data file linked every list says so.
'
' Owner: the Admin page's build task. Depends on: modAdmin, modErrors,
' modGlobals, modSignIn, modSubscriptions, modUI; the kit's modKitKeys,
' modKitUI, modKitUtil, modMenu, modPersonnel.

Private Const MSG_TITLE As String = "Admin"
Private Const ROLE_SLOTS As Long = 6            ' chkRole1..6 (modBuildAdmin)

Private m_tab As String         ' "access", "users" or "types"
Private m_ok As Boolean         ' a data file is linked
Private m_admin As Boolean      ' a System Admin (or a UI-only build): Users and Message types show
Private m_pid As Long           ' the open person (0: none)
Private m_userSnap As String    ' the open person as loaded (UserSnap), to tell unsaved changes
Private m_typeId As Long        ' the open message type (0: none, or a new one)
Private m_typeNew As Boolean    ' the type form holds a new type not saved yet
Private m_typeSnap As String    ' the open type as loaded (TypeSnap)
Private m_roleCount As Long     ' role slots in use
Private m_roleCode(1 To 6) As String    ' the role code of each slot

' The help topic for the open tab (KitHelp).
Public Property Get HelpTopic() As String
    Select Case m_tab
        Case "users": HelpTopic = "admin.users"
        Case "types": HelpTopic = "admin.types"
        Case Else: HelpTopic = "admin"
    End Select
End Property

Private Sub Form_Load()
    Dim c As Control
    On Error GoTo Err_Handler
    m_ok = DataLinked()
    m_admin = (Not m_ok) Or IsAnchorAdmin()
    If Not m_admin Then
        ' An approver only: no tabs to pick from.
        For Each c In Me.Controls
            If c.Name Like "segTab_*" Then SetVisible c, False
        Next
    End If
    If m_ok Then
        Me.sfQueue.Form.RecordSource = PendingSubsSql()
        Me.sfRules.Form.RecordSource = AdmRulesSql()
        Me.sfAccess.Form.RecordSource = AdmAccessSql()
        If m_admin Then
            LoadRoleSlots
            Me.sfTypes.Form.RecordSource = AdmTypesSql()
            Me.sfClasses.Form.RecordSource = AdmClassesSql()
            Me.cboTypeCat.RowSource = AdmTypeCategoriesSql()
        End If
    End If
    m_tab = "access"
    SetSegment Me, "segTab", "access"
    ShowGroup
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.Form_Load"
End Sub

' Leaving the page (the shell asks): unsaved changes are saved or dropped
' first, or the page stays.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = LeavePerson()
    If CanLeave Then CanLeave = LeaveType()
    Exit Function
Err_Handler:
    HandleError "pgAdmin.CanLeave"
End Function

' Ctrl+S saves the open person or type; Ctrl+N adds a type, or goes to the
' person picker (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_tab = "users" And m_pid <> 0 Then SaveUser
            If m_tab = "types" And (m_typeId <> 0 Or m_typeNew) Then SaveType
            Shortcut = True
        Case SHORTCUT_NEW
            If m_tab = "types" And m_ok Then NewType: Shortcut = True
            If m_tab = "users" And m_ok Then Me.txtPerson.SetFocus: Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgAdmin.Shortcut", action
    Shortcut = True
End Function

' ---------------------------------------------------------------------------
' Tabs
' ---------------------------------------------------------------------------

' The tabs' Segmented control (onChange).
Public Sub TabChanged()
    On Error GoTo Err_Handler
    m_tab = SegmentValue(Me, "segTab")
    If Not m_admin Then m_tab = "access"
    ShowGroup
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.TabChanged", m_tab
End Sub

' Show the open tab's controls and hide the others', then load its lists.
Private Sub ShowGroup()
    Dim nm As Variant, grp As Variant, keys As Variant, i As Long
    grp = Array("grpA", "grpU", "grpT")
    keys = Array("access", "users", "types")
    ' Hide first, so the open tab's controls never sit over another's.
    For i = 0 To 2
        If keys(i) <> m_tab Then
            For Each nm In Split(Me.Controls(grp(i)).Caption, ";")
                If Len(nm) > 0 Then SetVisible Me.Controls(nm), False
            Next
        End If
    Next
    For Each nm In Array("emptyQueue", "emptyRules", "emptyAccess", "emptyUsers", "emptyTypes")
        ShowEmpty Me, CStr(nm), False
    Next
    For i = 0 To 2
        If keys(i) = m_tab Then
            For Each nm In Split(Me.Controls(grp(i)).Caption, ";")
                If Len(nm) > 0 Then SetVisible Me.Controls(nm), True
            Next
        End If
    Next
    Select Case m_tab
        Case "users": RefreshUsers m_pid
        Case "types": RefreshTypes m_typeId
        Case Else
            RefreshQueue 0
            RefreshReadOnly Me.sfRules, "hdrRules", "emptyRules", "no alert rules"
            RefreshReadOnly Me.sfAccess, "hdrAccess", "emptyAccess", "no access groups"
    End Select
End Sub

' Show or hide a table's column captions (GridHeaders prefix & 0, 1, ...).
Private Sub ShowHeaders(ByVal prefix As String, ByVal v As Boolean)
    Dim c As Control
    For Each c In Me.Controls
        If c.Name Like prefix & "#*" Then SetVisible c, v
    Next
End Sub

' The rows table sf shows now.
Private Function ShownCount(sf As Control) As Long
    Dim rs As DAO.Recordset
    Set rs = sf.Form.RecordsetClone
    If rs.BOF And rs.EOF Then Exit Function
    rs.MoveLast
    ShownCount = rs.RecordCount
End Function

' A read-only table (rules, access groups): reload it, or its empty state.
'   what   the words for "this copy has no data file, so there are <what>"
Private Sub RefreshReadOnly(sf As Control, ByVal hdr As String, ByVal emp As String, ByVal what As String)
    Dim n As Long
    If m_ok Then
        sf.Form.Requery
        n = ShownCount(sf)
        ShowEmpty Me, emp, n = 0
    Else
        ShowEmpty Me, emp, True, "No data file linked", "This copy was built without the data file, so there are " & what & "."
    End If
    SetVisible sf, n > 0
    ShowHeaders hdr, n > 0
End Sub

' Row click in a table (subGrid).
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    id = Nz(Me.Controls(grid).Form.txtSel, 0)
    Select Case grid
        Case "sfUsers"
            If id = m_pid Or id = 0 Then Exit Sub
            If Not LeavePerson() Then
                GridSelect Me.sfUsers, m_pid
                Exit Sub
            End If
            ShowPerson id
        Case "sfTypes"
            If (id = m_typeId And Not m_typeNew) Or id = 0 Then Exit Sub
            If Not LeaveType() Then
                GridSelect Me.sfTypes, m_typeId
                Exit Sub
            End If
            ShowType id
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.GridClick", grid
End Sub

' ---------------------------------------------------------------------------
' Approvals & Access: the queue
' ---------------------------------------------------------------------------

' Reload the queue and pick keepId (0, or no longer listed: the first).
Private Sub RefreshQueue(ByVal keepId As Long)
    Dim n As Long
    If m_ok Then
        Me.sfQueue.Form.Requery
        n = ShownCount(Me.sfQueue)
        ShowEmpty Me, "emptyQueue", n = 0
    Else
        ShowEmpty Me, "emptyQueue", True, "No data file linked", "This copy was built without the data file, so there are no requests."
    End If
    Me.lblQueueHdr.Caption = "Subscription approval queue" & IIf(n > 0, " (" & n & ")", "")
    SetVisible Me.sfQueue, n > 0
    ShowHeaders "hdrQueue", n > 0
    SetVisible Me.lblQueueNote, n > 0
    SetEnabled Me.btnApprove, n > 0
    SetEnabled Me.btnDeny, n > 0
    If n > 0 Then
        If Not GridSelect(Me.sfQueue, keepId) Then SelectFirstRow Me.sfQueue
    End If
End Sub

' The picked request, "Training Pipeline for ETC Nguyen", or "" when none is.
Private Function RequestText(ByVal id As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT s.SubName, ps.PersonName FROM tblPersonSubs AS ps INNER JOIN tblSubscriptions AS s " & _
        "ON ps.SubID=s.SubID WHERE ps.PersonSubID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then RequestText = Nz(rs!SubName, "") & " for " & Nz(rs!PersonName, "")
    rs.Close
End Function

' The picked request's id, or 0 after telling the approver to pick one.
Private Function PickedRequest() As Long
    PickedRequest = Nz(Me.sfQueue.Form.txtSel, 0)
    If PickedRequest = 0 Then KitMsgBox "Pick a request in the queue first.", vbInformation, MSG_TITLE
End Function

Private Sub btnApprove_Click()
    Dim id As Long, what As String, problem As String
    On Error GoTo Err_Handler
    id = PickedRequest()
    If id = 0 Then Exit Sub
    what = RequestText(id)
    problem = ApproveSubscription(id, "")
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        RefreshQueue id
        Exit Sub
    End If
    KitToast "Approved: " & ClipText(what, 80), "success"
    RefreshQueue 0
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnApprove_Click", id
    On Error Resume Next                    ' show the queue as it is now; a second failure was reported
    RefreshQueue id
End Sub

' Deny: the reason is asked first (the requester sees it); Cancel denies nothing.
Private Sub btnDeny_Click()
    Dim id As Long, what As String, reason As Variant, problem As String
    On Error GoTo Err_Handler
    id = PickedRequest()
    If id = 0 Then Exit Sub
    what = RequestText(id)
    reason = KitInputBox("Why is " & ClipText(what, 80) & " denied? The requester sees the reason.", "Deny request")
    If IsNull(reason) Then Exit Sub
    If Len(Trim$(reason)) = 0 Then
        KitMsgBox "Type a reason to deny the request. Nothing was changed.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    problem = DenySubscription(id, Trim$(CStr(reason)))
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        RefreshQueue id
        Exit Sub
    End If
    KitToast "Denied: " & ClipText(what, 80), "success"
    RefreshQueue 0
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnDeny_Click", id
    On Error Resume Next                    ' show the queue as it is now; a second failure was reported
    RefreshQueue id
End Sub

' ---------------------------------------------------------------------------
' Users
' ---------------------------------------------------------------------------

' Name the role check boxes from tblRoles, in its order; spare slots stay hidden.
Private Sub LoadRoleSlots()
    Dim rs As DAO.Recordset, i As Long
    Set rs = CurrentDb.OpenRecordset("SELECT RoleCode, RoleName, Detail FROM tblRoles ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF And i < ROLE_SLOTS
        i = i + 1
        m_roleCode(i) = rs!RoleCode
        Me.Controls("lblRole" & i).Caption = Replace(Nz(rs!RoleName, rs!RoleCode), "&", "&&")
        Me.Controls("lblRole" & i).ControlTipText = Nz(rs!Detail, "")
        Me.Controls("chkRole" & i).ControlTipText = Nz(rs!Detail, "")
        rs.MoveNext
    Loop
    rs.Close
    m_roleCount = i
End Sub

' Reload the users list and open pid (0, or not listed: the first person).
Private Sub RefreshUsers(ByVal pid As Long)
    Dim n As Long
    If Not m_ok Then
        ShowEmpty Me, "emptyUsers", True, "No data file linked", "This copy was built without the data file, so there are no users."
        SetVisible Me.sfUsers, False
        ShowHeaders "hdrUsers", False
        SetEnabled Me.btnAdd, False
        SetEnabled Me.tgShowOff, False
        ShowPerson 0
        Exit Sub
    End If
    Me.sfUsers.Form.RecordSource = AdmUsersSql(IsSwitchOn("tgShowOff"))
    n = ShownCount(Me.sfUsers)
    ShowEmpty Me, "emptyUsers", n = 0
    SetVisible Me.sfUsers, n > 0
    ShowHeaders "hdrUsers", n > 0
    If GridSelect(Me.sfUsers, pid) Then
        ShowPerson pid
    ElseIf IsNull(SelectFirstRow(Me.sfUsers)) Then  ' SelectFirstRow opens the first row itself (GridClick)
        ShowPerson 0
    End If
End Sub

' Show person pid on the right (0: nobody, and the hint instead).
Private Sub ShowPerson(ByVal pid As Long)
    Dim rs As DAO.Recordset, nm As Variant, isOpen As Boolean, s As String, i As Long, codes As String
    On Error GoTo Fail
    m_pid = pid
    isOpen = (pid <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT p.first_name, p.last_name, p.rank, p.username, e.Billet, e.Unit, e.Community, " & _
            "e.Location, e.Dsn, e.IsActive FROM tblPeople AS e INNER JOIN " & PERSON_TABLE & " AS p ON e.PersonID=p.person_id " & _
            "WHERE e.PersonID=" & pid, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_pid = 0
        Else
            Me.lblName.Caption = Replace(Trim$(Nz(rs!rank, "") & " " & Nz(rs!first_name, "") & " " & Nz(rs!last_name, "")), "&", "&&")
            s = "Signs in as " & Nz(rs!username, "?")
            If pid = SignedInPersonID() Then s = s & " (you)"
            Me.lblWho.Caption = Replace(s, "&", "&&")
            Me.txtBillet = rs!Billet: Me.txtUnit = rs!Unit: Me.txtCommunity = rs!Community
            Me.txtLocation = rs!Location: Me.txtDsn = rs!Dsn
            SetSwitch "tgActive", Nz(rs!IsActive, False)
            codes = "," & PersonRoleCodes(pid) & ","
            For i = 1 To m_roleCount
                Me.Controls("chkRole" & i) = (InStr(codes, "," & m_roleCode(i) & ",") > 0)
            Next
        End If
        rs.Close: Set rs = Nothing
    End If
    For Each nm In Array("lblName", "lblWho", "lblBilletCap", "txtBillet", "lblUnitCap", "lblCommunityCap", "txtUnit", "txtCommunity", _
            "lblLocationCap", "lblDsnCap", "txtLocation", "txtDsn", "lblRolesCap", "tgActive", "lbltgActive", "btnSave", "btnUndo")
        SetVisible Me.Controls(nm), isOpen
    Next
    For i = 1 To ROLE_SLOTS
        SetVisible Me.Controls("chkRole" & i), isOpen And i <= m_roleCount
        SetVisible Me.Controls("lblRole" & i), isOpen And i <= m_roleCount
    Next
    SetVisible Me.lblPersonEmpty, Not isOpen
    m_userSnap = IIf(isOpen, UserSnap(), "")
    Exit Sub
Fail:
    If Not rs Is Nothing Then rs.Close
    RaiseAgain "pgAdmin.ShowPerson", pid
End Sub

' The open person's form as one text, to tell whether it changed.
Private Function UserSnap() As String
    UserSnap = Trim$(FieldText(Me, Me.txtBillet)) & "|" & Trim$(FieldText(Me, Me.txtUnit)) & "|" & Trim$(FieldText(Me, Me.txtCommunity)) & _
        "|" & Trim$(FieldText(Me, Me.txtLocation)) & "|" & Trim$(FieldText(Me, Me.txtDsn)) & "|" & RolesTicked() & "|" & IsSwitchOn("tgActive")
End Function

' The ticked roles' codes, ",VIEW,ADMIN," (",," for none).
Private Function RolesTicked() As String
    Dim i As Long, s As String
    s = ","
    For i = 1 To m_roleCount
        If Nz(Me.Controls("chkRole" & i), False) Then s = s & m_roleCode(i) & ","
    Next
    RolesTicked = s
End Function

' True when the open person may be left: no changes, or the admin saved or
' dropped them. Cancel (or a failed save) keeps them open.
Private Function LeavePerson() As Boolean
    If m_pid = 0 Then LeavePerson = True: Exit Function
    If UserSnap() = m_userSnap Then LeavePerson = True: Exit Function
    Select Case KitMsgBox("Save your changes to " & ClipText(Replace(Me.lblName.Caption, "&&", "&"), 60) & "?", _
            vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: LeavePerson = SaveUser()
        Case vbNo: m_userSnap = UserSnap(): LeavePerson = True
    End Select
End Function

' Save the open person (modAdmin.AdmSavePerson). False when nothing was
' saved: the admin is told why.
Private Function SaveUser() As Boolean
    Dim f As Variant, lim As Variant, caps As Variant, i As Long, s As String, problem As String, pid As Long
    ' Control names, not the controls: FieldText takes a Control, which a
    ' Variant holding one can't be passed as.
    f = Array("txtBillet", "txtUnit", "txtCommunity", "txtLocation", "txtDsn")
    lim = Array(ADM_MAX_BILLET, ADM_MAX_UNIT, ADM_MAX_UNIT, ADM_MAX_UNIT, ADM_MAX_DSN)
    caps = Array("billet", "unit", "community", "location", "DSN")
    For i = 0 To 4
        s = Trim$(FieldText(Me, Me.Controls(CStr(f(i)))))
        If Len(s) > lim(i) Then
            KitMsgBox "The " & caps(i) & " can be at most " & lim(i) & " characters. Nothing was saved.", vbInformation, MSG_TITLE
            Me.Controls(CStr(f(i))).SetFocus
            Exit Function
        End If
    Next
    pid = m_pid
    problem = AdmSavePerson(pid, Trim$(FieldText(Me, Me.txtBillet)), Trim$(FieldText(Me, Me.txtUnit)), _
        Trim$(FieldText(Me, Me.txtCommunity)), Trim$(FieldText(Me, Me.txtLocation)), Trim$(FieldText(Me, Me.txtDsn)), _
        RolesTicked(), IsSwitchOn("tgActive"))
    If Len(problem) > 0 Then
        KitMsgBox problem & " Nothing was saved.", vbInformation, MSG_TITLE
        Exit Function
    End If
    m_userSnap = UserSnap()
    KitToast "Saved", "success"
    RefreshUsers pid
    If pid = SignedInPersonID() Then
        ' Their own roles changed: the sidebar follows at once.
        If Len(LoadRoles()) = 0 Then ShowPagesForRoles Forms("frmMain")
    End If
    SaveUser = True
End Function

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveUser
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnSave_Click", m_pid
End Sub

' Undo changes: show the person as last saved.
Private Sub btnUndo_Click()
    On Error GoTo Err_Handler
    ShowPerson m_pid
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnUndo_Click", m_pid
End Sub

' The role captions tick their check box, as a label of a check box would.
Private Sub FlipRole(ByVal i As Long)
    Me.Controls("chkRole" & i) = Not Nz(Me.Controls("chkRole" & i), False)
End Sub

Private Sub lblRole1_Click()
    On Error GoTo Err_Handler
    FlipRole 1
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole1_Click"
End Sub

Private Sub lblRole2_Click()
    On Error GoTo Err_Handler
    FlipRole 2
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole2_Click"
End Sub

Private Sub lblRole3_Click()
    On Error GoTo Err_Handler
    FlipRole 3
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole3_Click"
End Sub

Private Sub lblRole4_Click()
    On Error GoTo Err_Handler
    FlipRole 4
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole4_Click"
End Sub

Private Sub lblRole5_Click()
    On Error GoTo Err_Handler
    FlipRole 5
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole5_Click"
End Sub

Private Sub lblRole6_Click()
    On Error GoTo Err_Handler
    FlipRole 6
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.lblRole6_Click"
End Sub

' Show deactivated people (Toggle's onChange): the list again, once the
' open person's changes are saved or dropped.
Public Sub ShowOffChanged()
    On Error GoTo Err_Handler
    If Not LeavePerson() Then
        SetSwitch "tgShowOff", Not IsSwitchOn("tgShowOff")
        Exit Sub
    End If
    RefreshUsers m_pid
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.ShowOffChanged"
End Sub

Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtPerson, KeyCode
End Sub

' Add person: give the picked person ANCHOR access, then open them so the
' admin can give them their roles.
Private Sub btnAdd_Click()
    Dim pid As Long
    On Error GoTo Err_Handler
    pid = Nz(Me.txtPersonKey, 0)
    If pid = 0 Then
        KitMsgBox "Type part of a name in the box, then pick the person from the list under it.", vbInformation, MSG_TITLE
        Me.txtPerson.SetFocus
        Exit Sub
    End If
    If Not LeavePerson() Then Exit Sub
    If AdmAddPerson(pid) Then
        KitToast "Added. Give them their roles on the right.", "success"
    Else
        KitToast "They're already listed.", "info"
        If Not IsSwitchOn("tgShowOff") And Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Then
            SetSwitch "tgShowOff", True     ' a deactivated person: show them, to reactivate
        End If
    End If
    Me.txtPersonKey = Null
    KitPickShow Me.txtPerson
    RefreshUsers pid
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnAdd_Click", pid
End Sub

Private Function IsSwitchOn(ByVal nm As String) As Boolean
    IsSwitchOn = CBool(Nz(Me.Controls(nm & "_val").Value, False))
End Function

' Set a switch in code (no onChange).
Private Sub SetSwitch(ByVal nm As String, ByVal v As Boolean)
    Me.Controls(nm & "_val").Value = v
    Me.Controls(nm & "_bg").Requery
End Sub

' ---------------------------------------------------------------------------
' Message types
' ---------------------------------------------------------------------------

' Reload the types and open id (0, or not listed: the first). A new type
' being typed stays open.
Private Sub RefreshTypes(ByVal id As Long)
    Dim n As Long
    If Not m_ok Then
        ShowEmpty Me, "emptyTypes", True, "No data file linked", "This copy was built without the data file, so there are no message types."
        SetVisible Me.sfTypes, False
        SetVisible Me.sfClasses, False
        ShowHeaders "hdrTypes", False
        ShowHeaders "hdrClasses", False
        SetEnabled Me.btnAddType, False
        ShowType 0
        Exit Sub
    End If
    Me.sfTypes.Form.Requery
    Me.sfClasses.Form.Requery
    Me.cboTypeCat.Requery
    n = ShownCount(Me.sfTypes)
    ShowEmpty Me, "emptyTypes", n = 0
    SetVisible Me.sfTypes, n > 0
    ShowHeaders "hdrTypes", n > 0
    If m_typeNew Then
        ' A new type being typed (the tab was switched away and back): its
        ' form stays as it is, and no row is picked.
        GridSelect Me.sfTypes, 0
        ShowTypeForm True
        SetEnabled Me.btnRetire, False
    ElseIf GridSelect(Me.sfTypes, id) Then
        ShowType id
    ElseIf IsNull(SelectFirstRow(Me.sfTypes)) Then
        ShowType 0
    End If
End Sub

' Show message type id on the right (0: none, and the hint).
Private Sub ShowType(ByVal id As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, active As Boolean
    On Error GoTo Fail
    m_typeId = id: m_typeNew = False
    isOpen = (id <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT TypeCode, TypeName, Category, Description, IsActive FROM tblMessageTypes WHERE TypeID=" & id, _
            dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_typeId = 0
        Else
            active = Nz(rs!IsActive, False)
            Me.lblTypeHead.Caption = Replace(Nz(rs!TypeCode, ""), "&", "&&")
            Me.lblTypeState.Caption = IIf(active, "Active: offered for new traffic, filters and drafting", _
                "Retired: messages of this type keep it")
            Me.txtTypeCode = rs!TypeCode: Me.txtTypeName = rs!TypeName
            Me.cboTypeCat = rs!Category: Me.txtTypeDesc = rs!Description
            Me.btnRetire.Caption = IIf(active, "Retire", "Restore")
        End If
        rs.Close: Set rs = Nothing
    End If
    Me.txtTypeCode.Locked = True            ' messages name a type by its code
    ShowTypeForm isOpen
    SetEnabled Me.btnRetire, isOpen
    m_typeSnap = IIf(isOpen, TypeSnap(), "")
    Exit Sub
Fail:
    If Not rs Is Nothing Then rs.Close
    RaiseAgain "pgAdmin.ShowType", id
End Sub

' Show the type's fields, or the hint in their place.
Private Sub ShowTypeForm(ByVal isOpen As Boolean)
    Dim nm As Variant
    For Each nm In Array("lblTypeHead", "lblTypeState", "lblTypeCodeCap", "txtTypeCode", "lblTypeNameCap", "txtTypeName", _
            "lblTypeCatCap", "cboTypeCat", "lblTypeDescCap", "txtTypeDesc", "lblTypeHint", "btnSaveType", "btnRetire")
        SetVisible Me.Controls(nm), isOpen
    Next
    SetVisible Me.lblTypeEmpty, Not isOpen
End Sub

' The type form as one text, to tell whether it changed.
Private Function TypeSnap() As String
    TypeSnap = Trim$(FieldText(Me, Me.txtTypeCode)) & "|" & Trim$(FieldText(Me, Me.txtTypeName)) & "|" & _
        Trim$(FieldText(Me, Me.cboTypeCat)) & "|" & Trim$(FieldText(Me, Me.txtTypeDesc))
End Function

' True when the open type may be left (as LeavePerson).
Private Function LeaveType() As Boolean
    Dim nm As String
    If m_typeId = 0 And Not m_typeNew Then LeaveType = True: Exit Function
    If TypeSnap() = m_typeSnap Then LeaveType = True: Exit Function
    nm = IIf(m_typeNew, "the new message type", Replace(Me.lblTypeHead.Caption, "&&", "&"))
    Select Case KitMsgBox("Save your changes to " & nm & "?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: LeaveType = SaveType()
        Case vbNo
            If m_typeNew Then m_typeNew = False: m_typeId = 0
            m_typeSnap = TypeSnap(): LeaveType = True
    End Select
End Function

' Add type: an empty form with the code open for typing.
Private Sub NewType()
    If Not LeaveType() Then Exit Sub
    m_typeId = 0: m_typeNew = True
    Me.txtTypeCode = Null: Me.txtTypeName = Null: Me.cboTypeCat = Null: Me.txtTypeDesc = Null
    Me.lblTypeHead.Caption = "New message type"
    Me.lblTypeState.Caption = "Not saved yet"
    Me.btnRetire.Caption = "Retire"
    ShowTypeForm True
    SetEnabled Me.btnRetire, False
    Me.txtTypeCode.Locked = False
    GridSelect Me.sfTypes, 0
    m_typeSnap = TypeSnap()
    Me.txtTypeCode.SetFocus
End Sub

' Save the open type (modAdmin). False when it can't be saved yet; the admin
' is told what to fix.
Private Function SaveType() As Boolean
    Dim code As String, tName As String, cat As String, descr As String, problem As String, id As Long
    code = UCase$(Trim$(FieldText(Me, Me.txtTypeCode)))
    tName = Trim$(FieldText(Me, Me.txtTypeName))
    cat = Trim$(FieldText(Me, Me.cboTypeCat))
    descr = Trim$(FieldText(Me, Me.txtTypeDesc))
    problem = AdmTypeProblem(IIf(m_typeNew, 0, m_typeId), code, tName, cat, descr)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        Exit Function
    End If
    If m_typeNew Then
        id = AdmAddType(code, tName, cat, descr)
        m_typeNew = False
        KitToast "Added " & ClipText(code, 40), "success"
    Else
        id = m_typeId
        If AdmSaveType(id, tName, cat, descr) Then KitToast "Saved", "success"
    End If
    m_typeSnap = TypeSnap()
    RefreshTypes id
    SaveType = True
End Function

Private Sub btnAddType_Click()
    On Error GoTo Err_Handler
    NewType
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnAddType_Click"
End Sub

Private Sub btnSaveType_Click()
    On Error GoTo Err_Handler
    SaveType
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnSaveType_Click", m_typeId
End Sub

' Retire or Restore the open type. Unsaved changes are saved or dropped first.
Private Sub btnRetire_Click()
    Dim id As Long, active As Boolean
    On Error GoTo Err_Handler
    If m_typeId = 0 Or m_typeNew Then Exit Sub
    If Not LeaveType() Then Exit Sub
    id = m_typeId
    active = Not Nz(DLookup("IsActive", "tblMessageTypes", "TypeID=" & id), False)
    AdmSetTypeActive id, active
    KitToast IIf(active, "Restored", "Retired"), "success"
    RefreshTypes id
    Exit Sub
Err_Handler:
    HandleError "pgAdmin.btnRetire_Click", m_typeId
End Sub
