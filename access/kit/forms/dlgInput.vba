Option Compare Database
Option Explicit

' Input dialog for KitInputBox (BuildInputDialog builds it). Asks the
' question KitInputBox was given, with its starting text selected so typing
' replaces it, shows typed characters as * for a masked question, shrinks to
' the lines the prompt takes, and hands back the text on OK (SetKitInputAnswer).
' Enter clicks OK and Esc clicks Cancel; Cancel leaves the answer Null.
' A prompt longer than the dialog's 6 lines keeps its full height and scrolls
' when clicked, so no part of it is cut off.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim promptH As Long, inY As Long, dy As Long
    Me.Caption = IIf(Len(KitInputTitle()) > 0, KitInputTitle(), APP_NAME)
    Me.txtPrompt = KitInputPrompt()
    ' The mask goes on before the starting text, so that is never shown plain.
    If KitInputMasked() Then Me.txtValue.InputMask = SECRET_MASK
    If Len(KitInputDefault()) > 0 Then Me.txtValue = KitInputDefault()

    ' Shrink to the prompt: move the input and buttons up under it, then the
    ' dialog. No prompt: the input takes its place at the top.
    If Len(KitInputPrompt()) = 0 Then
        SetVisible Me.txtPrompt, False
        inY = Me.txtPrompt.Top
    Else
        promptH = PromptHeight()
        If promptH > Me.txtPrompt.Height Then promptH = Me.txtPrompt.Height
        inY = Me.txtPrompt.Top + promptH + 180       ' 12px under the prompt
    End If
    dy = Me.txtValue.Top - inY
    If dy > 0 Then
        If promptH > 0 Then Me.txtPrompt.Height = promptH
        Me.txtValue.Top = inY
        Me.btnOK.Top = Me.btnOK.Top - dy
        Me.btnCancel.Top = Me.btnOK.Top
        Me.Section(acDetail).Height = Me.Section(acDetail).Height - dy
        Me.InsideHeight = Me.Section(acDetail).Height
    End If
    CenterOnApp Me                  ' at its final size, over Access (modKitUtil)

    Me.txtValue.SetFocus
    Me.txtValue.SelStart = 0
    Me.txtValue.SelLength = Len(Nz(Me.txtValue.Text, ""))
    Exit Sub
Err_Handler:
    ' Whatever failed, the dialog still shows the input and both buttons.
    LogError "dlgInput.Form_Load", Left$(KitInputPrompt(), 60)
End Sub

' Height the prompt's lines take (twips), measured as dlgMessage measures a
' message: each paragraph wraps to as many 10pt lines as its measured width
' needs, with 10% spare for word breaks.
Private Function PromptHeight() As Long
    Dim para As Variant, lines As Long, w As Long
    w = Me.txtPrompt.Width
    For Each para In Split(Replace(KitInputPrompt(), vbCrLf, vbLf), vbLf)
        If Len(para) = 0 Then
            lines = lines + 1
        Else
            lines = lines - Int(-(MeasureText(CStr(para), 10, 400) * 1.1) / w)
        End If
    Next
    If lines < 1 Then lines = 1
    PromptHeight = lines * CLng(10 * 20 * 1.36) + 30
End Function

' The text in the input. Enter in the input clicks OK before the input's
' Value is updated, and a masked input's Text is only *, so move the focus to
' OK first, which commits what was typed, then read Value.
Private Function EnteredText() As String
    Me.btnOK.SetFocus
    EnteredText = Nz(Me.txtValue.Value, "")
End Function

Private Sub btnOK_Click()
    On Error GoTo Err_Handler
    SetKitInputAnswer EnteredText()
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgInput.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    SetKitInputAnswer Null
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgInput.btnCancel_Click"
End Sub
