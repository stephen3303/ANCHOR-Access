Option Compare Database
Option Explicit

' New draft (modBuildDrafting.Build_dlgNewDraft), from the Drafting page:
' a blank draft or one of the active templates, with the picked template's
' description. Start sets TempVars("NewDraftTemplateID") (0 for blank) and
' closes with DialogResult "ok"; the page then starts the draft (asking for
' the template's fields), so no dialog opens over this one. Cancel, or Esc,
' starts nothing.
'
' Owner: drafting and release. Depends on: modErrors, modTemplates; the
' kit's modKitUtil.

Public Property Get HelpTopic() As String
    HelpTopic = "drafting"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset, s As String
    On Error GoTo Err_Handler
    s = "0;""(Blank draft)"""
    Set rs = CurrentDb.OpenRecordset("SELECT TemplateID, TemplateName, VersionNo FROM tblTemplates WHERE Status='Active' ORDER BY TemplateName", _
        dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & ";" & rs!TemplateID & ";""" & Replace(rs!TemplateName & " v" & rs!VersionNo, """", """""") & """"
        rs.MoveNext
    Loop
    rs.Close
    Me.cboTemplate.RowSource = s
    Me.cboTemplate = 0
    ShowDescription
    CenterOnApp Me
    Me.cboTemplate.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgNewDraft.Form_Load"
End Sub

Private Sub cboTemplate_AfterUpdate()
    On Error GoTo Err_Handler
    ShowDescription
    Exit Sub
Err_Handler:
    HandleError "dlgNewDraft.cboTemplate_AfterUpdate"
End Sub

Private Sub ShowDescription()
    Dim tid As Long, n As Long
    tid = Nz(Me.cboTemplate, 0)
    If tid = 0 Then
        Me.lblTplDesc.Caption = "An empty draft with the command's FM line, GENADMIN and ROUTINE. You write the rest."
    Else
        n = DCount("*", "tblTemplateFields", "TemplateID=" & tid)
        Me.lblTplDesc.Caption = Nz(DLookup("Description", "tblTemplates", "TemplateID=" & tid), "") & " " & _
            IIf(n = 0, "No fields to fill in.", "You fill in " & Plural(n, "field") & " next.")
    End If
End Sub

Private Sub btnStart_Click()
    On Error GoTo Err_Handler
    TempVars("NewDraftTemplateID") = Nz(Me.cboTemplate, 0)
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgNewDraft.btnStart_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgNewDraft.btnCancel_Click"
End Sub
