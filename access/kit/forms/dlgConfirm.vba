Option Compare Database
Option Explicit

' Type-to-confirm dialog for KitConfirmTyped (BuildConfirmDialog builds it),
' before a step that can't be undone. Shows the heading in the red band, the
' detail below it, and which word to type; the action button stays disabled
' until the box holds that word exactly (case counts), as GitHub asks for a
' repository's name before deleting it. Esc and Cancel close it unconfirmed;
' Enter does nothing, so a stray key never confirms. A question with an
' option (KitConfirmOption) shows its check box, unticked; one without hides
' it and moves the rows under it up.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.Caption = KitConfirmTitle()
    Me.lblHeading.Caption = KitConfirmHeading()
    Me.txtDetail = KitConfirmDetail()
    Me.lblAsk.Caption = "To confirm, type " & KitConfirmWord() & " in the box below."
    Me.btnConfirm.Caption = KitConfirmAction()
    Me.chkAlso = False
    If Len(KitConfirmOption()) > 0 Then
        Me.lblAlso.Caption = KitConfirmOption()
    Else
        HideOption
    End If
    SetEnabled Me.btnConfirm, False
    CenterOnApp Me                  ' over Access (modKitUtil)
    Me.txtValue.SetFocus
    Exit Sub
Err_Handler:
    ' The dialog still opens; with the button disabled nothing can be confirmed.
    LogError "dlgConfirm.Form_Load", KitConfirmHeading()
End Sub

' No option: the check box goes, and the rows under it move up.
Private Sub HideOption()
    Dim dy As Long, c As Control, askTop As Long
    askTop = Me.lblAsk.Top
    dy = askTop - Me.lblAlso.Top
    SetVisible Me.chkAlso, False
    SetVisible Me.lblAlso, False
    For Each c In Me.Controls
        If c.Top >= askTop Then c.Top = c.Top - dy
    Next
    Me.Section(acDetail).Height = Me.Section(acDetail).Height - dy
    Me.InsideHeight = Me.Section(acDetail).Height
End Sub

' The option's text ticks and unticks it too, like a check box's own label.
Private Sub lblAlso_Click()
    Me.chkAlso = Not Nz(Me.chkAlso, False)
End Sub

' Each key typed: the button is enabled only while the box holds the word.
Private Sub txtValue_Change()
    On Error GoTo Err_Handler
    SetEnabled Me.btnConfirm, KitConfirmMatches(Nz(Me.txtValue.Text, ""))
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.txtValue_Change"
End Sub

Private Sub btnConfirm_Click()
    On Error GoTo Err_Handler
    ' The click moved the focus off the box, so its Value holds what was typed.
    If Not KitConfirmMatches(Nz(Me.txtValue.Value, "")) Then
        SetEnabled Me.btnConfirm, False
        Exit Sub
    End If
    SetKitConfirmOptionChosen Nz(Me.chkAlso, False)
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.btnConfirm_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgConfirm.btnCancel_Click"
End Sub
