Option Compare Database
Option Explicit

' Field rules (modBuildDrafting.Build_dlgTemplateField), Message
' Generator's dlgFieldRule: opened by the Templates page with RunDialog,
' OpenArgs "id=<FieldID>". Shows the field's rules (required, max length,
' default) and its hint for drafters; Save stores them
' (modTemplates.SaveTemplateField). A problem the person can fix (a bad max
' length, a default or hint too long, a template that can't change any more)
' shows in lblMsg and keeps the dialog open.
'
' Owner: drafting and release. Depends on: modErrors, modTemplates; the
' kit's modKitUtil.

Private m_id As Long

Public Property Get HelpTopic() As String
    HelpTopic = "templates.fields"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset, label As String
    On Error GoTo Err_Handler
    m_id = OpenArgLong(Me, "id")
    Set rs = CurrentDb.OpenRecordset("SELECT TemplateID, FieldName, IsRequired, MaxLen, DefaultText, Hint FROM tblTemplateFields " & _
        "WHERE FieldID=" & m_id, dbOpenSnapshot)
    If Not rs.EOF Then
        label = ClipText(TemplateLabel(rs!TemplateID), 60)      ' lblIntro holds three lines
        Me.lblIntro.Caption = "{{ " & rs!FieldName & " }} in " & label & _
            ". The rules apply to drafts started after the template is activated."
        Me.cboRequired = IIf(rs!IsRequired, 1, 0)
        Me.txtMax = Nz(rs!MaxLen, 0)
        Me.txtDefault = Nz(rs!DefaultText, "")
        Me.txtHint = Nz(rs!Hint, "")
    End If
    rs.Close
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateField.Form_Load", m_id
End Sub

' Moving the focus to Save first commits a value still being typed.
Private Sub btnSave_Click()
    Dim why As String
    On Error GoTo Err_Handler
    Me.btnSave.SetFocus
    why = SaveTemplateField(m_id, CLng(Nz(Me.cboRequired, 1)) = 1, Nz(Me.txtMax, ""), Nz(Me.txtDefault, ""), Nz(Me.txtHint, ""))
    If Len(why) > 0 Then Me.lblMsg.Caption = why: Exit Sub
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateField.btnSave_Click", m_id
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateField.btnCancel_Click"
End Sub
