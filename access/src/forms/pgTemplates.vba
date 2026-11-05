Option Compare Database
Option Explicit

' Templates (see modBuildDrafting.Build_pgTemplates): every template
' version (modTemplates.TemplateListSql), filtered by a chip, the search box
' and the column funnels, and the picked one on the right in four tabs:
' Header (name, description, the draft header and the subject), Text,
' Fields (one row per {{ field }}, with its rules: dlgTemplateField) and
' Preview (as a draft from it would transmit). ANCHOR admins edit an
' Inactive version (Save, Undo), activate it (the name's active version is
' retired), retire an active one, and copy any as a New version. Anyone
' starts a draft from an active one: Start draft asks for each field, then
' opens the draft on the Drafting page (TempVars("DraftingDraftID")).
'
' Owner: drafting and release. Depends on: modDrafting, modErrors,
' modSignIn, modTemplates, modUI; the kit's modKitUI, modKitUtil,
' modGridFilter.

Private Const MSG_TITLE As String = "Templates"

Private m_chip As Long          ' TPL_CHIPS; 0 All
Private m_search As String
Private m_id As Long            ' the template shown (0: none)
Private m_tab As String         ' header, text, fields or preview
Private m_t As Object           ' its values as last read (TemplateValues)
Private m_sig As String         ' the editor's values as loaded
Private m_editable As Boolean   ' an admin, and the version is Inactive

Public Property Get HelpTopic() As String
    HelpTopic = "templates"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "header"
    ChipOn Me, "chip", 0
    SetVisible Me.btnNewTemplate, IsAnchorAdmin() Or Not DataLinked()
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no templates."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnNewTemplate, False
        ShowTemplate 0
        Exit Sub
    End If
    Me.cboTType.RowSource = "SELECT TypeCode, TypeName FROM tblMessageTypes WHERE IsActive ORDER BY SortOrder"
    Me.cboTClass.RowSource = "SELECT ClassCode, ClassName FROM tblMessageClasses ORDER BY ClassCode"
    RefreshList 0
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.Form_Load"
End Sub

Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = OfferSave()
    Exit Function
Err_Handler:
    HandleError "pgTemplates.CanLeave"
    CanLeave = True
End Function

Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_editable Then SaveClicked
            Shortcut = True
        Case SHORTCUT_NEW
            If Me.btnNewTemplate.Visible And Me.btnNewTemplate.Enabled Then btnNewTemplate_Click
            Shortcut = True
        Case SHORTCUT_FIND
            Me.txtSearch.SetFocus
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgTemplates.Shortcut", action
    Shortcut = True
End Function

' ---------------------------------------------------------------------------
' Chips, search and tabs
' ---------------------------------------------------------------------------

Public Function ChipPick(ByVal i As Long)
    On Error GoTo Err_Handler
    If i = m_chip Then Exit Function
    If Not OfferSave() Then Exit Function
    m_chip = i
    ChipOn Me, "chip", i
    RefreshList m_id
    Exit Function
Err_Handler:
    HandleError "pgTemplates.ChipPick", i
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.SearchFor", text
End Sub

Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.ListFiltered", sfName
End Sub

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    m_tab = key
    ShowTab
    Exit Function
Err_Handler:
    HandleError "pgTemplates.TabPick", key
End Function

' ---------------------------------------------------------------------------
' Rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfList"
            id = Nz(Me.sfList.Form.txtSel, 0)
            If id = m_id Then Exit Sub
            If Not OfferSave() Then GridSelect Me.sfList, m_id: Exit Sub
            ShowTemplate id
        Case "sfFields"
            If dbl Then FieldRules
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.GridClick", grid & " " & id
End Sub

Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    m_chip = 0
    ChipOn Me, "chip", 0
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.emptyList_b_Click"
End Sub

' ---------------------------------------------------------------------------
' Buttons
' ---------------------------------------------------------------------------

Private Sub btnNewTemplate_Click()
    Dim id As Long
    On Error GoTo Err_Handler
    If Not OfferSave() Then Exit Sub
    If Not IsAnchorAdmin() Then KitMsgBox "Only an ANCHOR admin adds templates.", vbInformation, MSG_TITLE: Exit Sub
    id = NewTemplate()
    If m_chip <> 0 And m_chip <> 2 Then m_chip = 0: ChipOn Me, "chip", 0
    m_tab = "header"
    RefreshList id
    KitToast "Template added, inactive. Name it, write its text, then Activate it.", "success"
    Me.txtName.SetFocus
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnNewTemplate_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveClicked
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnSave_Click", m_id
End Sub

Private Sub btnUndo_Click()
    On Error GoTo Err_Handler
    ShowTemplate m_id
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnUndo_Click", m_id
End Sub

Private Sub btnActivate_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If DirtyNow() Then
        If Not SaveClicked() Then Exit Sub
    End If
    why = ActivateTemplate(m_id)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    KitToast TemplateLabel(m_id) & " is active: drafts start from it now.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnActivate_Click", m_id
End Sub

Private Sub btnRetire_Click()
    Dim why As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If KitMsgBox("Retire " & TemplateLabel(m_id) & "? No new draft starts from it; drafts already started keep their text.", _
            vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    why = RetireTemplate(m_id)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE Else KitToast TemplateLabel(m_id) & " retired.", "success"
    RefreshList m_id
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnRetire_Click", m_id
End Sub

Private Sub btnNewVersion_Click()
    Dim id As Long, problem As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    If Not OfferSave() Then Exit Sub
    id = NewTemplateVersion(m_id, problem)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If m_chip <> 0 And m_chip <> 2 Then m_chip = 0: ChipOn Me, "chip", 0
    RefreshList id
    KitToast TemplateLabel(id) & " added, inactive, to edit.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnNewVersion_Click", m_id
End Sub

Private Sub btnStart_Click()
    Dim id As Long, problem As String
    On Error GoTo Err_Handler
    If m_id = 0 Then Exit Sub
    id = StartDraftFromTemplate(m_id, problem)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If id = 0 Then Exit Sub                     ' cancelled
    TempVars("DraftingDraftID") = id
    NavTo "drafting"
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnStart_Click", m_id
End Sub

Private Sub btnFieldRules_Click()
    On Error GoTo Err_Handler
    FieldRules
    Exit Sub
Err_Handler:
    HandleError "pgTemplates.btnFieldRules_Click", m_id
End Sub

' Open the picked field's rules (dlgTemplateField); the fields follow the
' saved tags, so unsaved edits are saved first.
Private Sub FieldRules()
    Dim fid As Long, why As String
    If m_id = 0 Then Exit Sub
    why = TemplateEditProblem(m_id)
    If Len(why) > 0 Then KitMsgBox why, vbInformation, MSG_TITLE: Exit Sub
    If DirtyNow() Then
        If KitMsgBox("Save the template first? Its fields follow the tags as saved.", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
        If Not SaveClicked() Then Exit Sub
    End If
    fid = Nz(Me.sfFields.Form.txtSel, 0)
    If fid = 0 Then KitMsgBox "Pick a field in the list first.", vbInformation, MSG_TITLE: Exit Sub
    If RunDialog("dlgTemplateField", "id=" & fid) Then
        Me.sfFields.Form.Requery
        GridSelect Me.sfFields, fid
        KitToast "Field rules saved.", "success"
    End If
End Sub

' ---------------------------------------------------------------------------
' The list
' ---------------------------------------------------------------------------

Private Sub RefreshList(ByVal openId As Long)
    Dim i As Long
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = TemplateListSql(m_chip, m_search)
    SetGridFilter Me.sfList, ""
    GridHighlight Me.sfList, m_search
    For i = 0 To UBound(Split(TPL_CHIPS, "|"))
        StyleChip Me.Controls("chip_" & i), i = m_chip, TemplateCount(i)
    Next
    m_id = openId
    m_sig = ""
    AfterFilter
End Sub

Private Sub AfterFilter()
    Dim n As Long, keep As Long
    n = Me.sfList.Form.RecordsetClone.RecordCount
    If TemplateCount(0) = 0 Then
        ShowEmpty Me, "emptyList", n = 0, "No templates yet", "An ANCHOR admin adds templates with New template."
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No templates match", "No template fits the chip, the search and the column filters."
        Me.emptyList_b.Caption = "Show all"
    End If
    SetVisible Me.sfList, n > 0
    keep = m_id
    If GridSelect(Me.sfList, keep) Then
        ShowTemplate keep
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowTemplate 0
    Else
        ShowTemplate Nz(Me.sfList.Form.txtSel, 0)
    End If
End Sub

' ---------------------------------------------------------------------------
' The template
' ---------------------------------------------------------------------------

Private Sub ShowTemplate(ByVal tid As Long)
    Dim c As Control, st As String, admin As Boolean, used As String
    m_id = tid
    Set m_t = Nothing
    If tid <> 0 Then
        Set m_t = TemplateValues(tid)
        If m_t Is Nothing Then m_id = 0
    End If
    admin = IsAnchorAdmin()
    If m_id <> 0 Then
        st = m_t("status")
        m_editable = (Len(TemplateEditProblem(m_id)) = 0)
        Me.lblTplName.Caption = m_t("name") & " v" & m_t("version")
        ' Format$ of Null raises, and IIf would run both branches: the dates are read apart.
        If IsNull(m_t("last_used_on")) Then used = "no draft started from it yet" Else used = "last draft started " & _
            Format$(m_t("last_used_on"), "d mmm yyyy")
        Me.lblTplSub.Caption = st & " " & ChrW(183) & " changed " & Format$(Nz(m_t("updated_on"), Now()), "d mmm yyyy") & _
            IIf(Len(m_t("updated_by")) > 0, " by " & m_t("updated_by"), "") & " " & ChrW(183) & " " & used
        Me.txtName = m_t("name")
        Me.txtDesc = m_t("description")
        Me.cboTType = m_t("msg_type")
        Me.cboTClass = m_t("msg_class")
        Me.cboTPrec = m_t("precedence")
        Me.cboTClassif = m_t("classification")
        Me.txtTTo = m_t("to_line")
        Me.txtTInfo = m_t("info_line")
        Me.txtTSubject = m_t("subject")
        Me.txtBody = m_t("body")
        Me.sfFields.Form.RecordSource = TemplateFieldsSql(m_id)
        m_sig = Signature(EditorValues())
    Else
        m_editable = False
        m_sig = ""
    End If
    For Each c In Me.Controls
        If c.Name Like "tab_*" Or c.Name Like "ul_*" Or c.Name = "boxTabLine" Or c.Name = "lblTplName" Or c.Name = "lblTplSub" Then
            SetVisible c, m_id <> 0
        End If
    Next
    SetVisible Me.btnSave, m_id <> 0 And m_editable
    SetVisible Me.btnUndo, m_id <> 0 And m_editable
    SetVisible Me.btnActivate, m_id <> 0 And admin And st = "Inactive"
    SetVisible Me.btnRetire, m_id <> 0 And admin And st = "Active"
    SetVisible Me.btnNewVersion, m_id <> 0 And admin And st <> "Inactive"
    SetVisible Me.btnStart, m_id <> 0 And st = "Active"
    SetVisible Me.lblDetailEmpty, m_id = 0
    ShowTab
End Sub

Private Sub ShowTab()
    Dim k As Variant, v As Variant
    If m_id <> 0 Then SelectTab Me, m_tab
    For Each k In Array("header", "text", "fields", "preview")
        ShowGroup "grp" & UCase$(Left$(k, 1)) & Mid$(k, 2), m_id <> 0 And k = m_tab
    Next
    If m_id = 0 Then Exit Sub
    For Each v In Array("txtName", "txtDesc", "cboTType", "cboTClass", "cboTPrec", "cboTClassif", "txtTTo", "txtTInfo", "txtTSubject", "txtBody")
        Me.Controls(v).Locked = Not m_editable
    Next
    Select Case m_tab
        Case "fields"
            SetEnabled Me.btnFieldRules, m_editable
        Case "preview"
            Me.txtTplPreview = TemplatePreviewText(EditorValues())
    End Select
End Sub

Private Sub ShowGroup(ByVal nm As String, ByVal isShown As Boolean)
    Dim v As Variant
    For Each v In Split(Me.Controls(nm).Caption, ";")
        If Len(v) > 0 Then SetVisible Me.Controls(v), isShown
    Next
End Sub

' ---------------------------------------------------------------------------
' Editing
' ---------------------------------------------------------------------------

' The editor's values (TemplateValues' keys), with the template's id.
Private Function EditorValues() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("id") = m_id
    d("name") = Trim$(Nz(Me.txtName, ""))
    d("description") = Trim$(Nz(Me.txtDesc, ""))
    d("msg_type") = Nz(Me.cboTType, "")
    d("msg_class") = Nz(Me.cboTClass, "")
    d("precedence") = Nz(Me.cboTPrec, "")
    d("classification") = Nz(Me.cboTClassif, "")
    d("to_line") = Trim$(Nz(Me.txtTTo, ""))
    d("info_line") = Trim$(Nz(Me.txtTInfo, ""))
    d("subject") = Trim$(Nz(Me.txtTSubject, ""))
    d("body") = Nz(Me.txtBody, "")
    Set EditorValues = d
End Function

Private Function Signature(ByVal d As Object) As String
    Dim k As Variant
    For Each k In d.Keys
        Signature = Signature & k & "=" & d(k) & vbTab
    Next
End Function

Private Function DirtyNow() As Boolean
    If m_id = 0 Or Not m_editable Or Len(m_sig) = 0 Then Exit Function
    DirtyNow = (Signature(EditorValues()) <> m_sig)
End Function

Private Function SaveClicked() As Boolean
    Dim why As String
    why = SaveTemplate(m_id, EditorValues())
    If Len(why) > 0 Then KitMsgBox why, vbExclamation, MSG_TITLE: Exit Function
    KitToast "Template saved.", "success"
    RefreshList m_id
    SaveClicked = True
End Function

Private Function OfferSave() As Boolean
    If Not DirtyNow() Then OfferSave = True: Exit Function
    Select Case KitMsgBox("Save your changes to this template?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: OfferSave = SaveClicked()
        Case vbNo: m_sig = "": OfferSave = True
        Case Else: OfferSave = False
    End Select
End Function
