Option Compare Database
Option Explicit

' A workflow step (modBuildDrafting.Build_dlgRouteStep), from Admin:
' Routing: OpenArgs "wf=<WorkflowID>;step=<WfStepID, 0 to add one>". The
' step's name, whether it reviews or releases, and who holds it (a person
' in Personnel who may open ANCHOR). Save stores it
' (modRouting.SaveWorkflowStep, which keeps an active workflow's chain whole);
' a refusal shows in lblMsg and keeps the dialog open.
'
' Owner: drafting and release. Depends on: modErrors, modRouting; the
' kit's modKitUtil.

Private m_wf As Long
Private m_step As Long

Public Property Get HelpTopic() As String
    HelpTopic = "routing"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset
    On Error GoTo Err_Handler
    m_wf = OpenArgLong(Me, "wf")
    m_step = OpenArgLong(Me, "step")
    Me.cboPerson.RowSource = RosterSql()
    Me.lblSub.Caption = "On " & ClipText(WorkflowName(m_wf), 60) & ". The last step of the chain releases."
    If m_step = 0 Then
        Me.Caption = "Add a step"
        Me.lblHead.Caption = "Add a step"
        Me.cboAction = "review"
    Else
        Me.Caption = "Change a step"
        Set rs = CurrentDb.OpenRecordset("SELECT StepOrder, StepTitle, StepAction, PersonID FROM tblWorkflowSteps WHERE WfStepID=" & m_step, _
            dbOpenSnapshot)
        If Not rs.EOF Then
            Me.lblHead.Caption = "Step " & rs!StepOrder & ": " & ClipText(Nz(rs!StepTitle, ""), 40)
            Me.txtTitle = rs!StepTitle
            Me.cboAction = rs!StepAction
            Me.cboPerson = rs!PersonID
        End If
        rs.Close
    End If
    CenterOnApp Me
    Me.txtTitle.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgRouteStep.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub btnSave_Click()
    Dim why As String
    On Error GoTo Err_Handler
    Me.btnSave.SetFocus
    why = SaveWorkflowStep(m_wf, m_step, Nz(Me.txtTitle, ""), Nz(Me.cboAction, ""), Nz(Me.cboPerson, 0))
    If Len(why) > 0 Then Me.lblMsg.Caption = why: Exit Sub
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRouteStep.btnSave_Click", m_wf & " " & m_step
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRouteStep.btnCancel_Click"
End Sub
