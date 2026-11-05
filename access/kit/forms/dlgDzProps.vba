Option Compare Database
Option Explicit

' Form designer element settings (BuildDesignerDialog, modKitDesigner): the
' rows the selected element's kind uses, checked and applied on OK.
' RunDialog's result: "ok" when the settings were applied.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DesignerPropsLoad Me
    Exit Sub
Err_Handler:
    HandleError "dlgDzProps.Form_Load"
End Sub

Private Sub btnOK_Click()
    On Error GoTo Err_Handler
    If Not DesignerPropsSave(Me) Then Exit Sub
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgDzProps.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgDzProps.btnCancel_Click"
End Sub
