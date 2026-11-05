Option Compare Database
Option Explicit

' Message dialog for KitMsgBox (BuildMessageDialog builds it). Shows the
' message KitMsgBox was given with its icon and buttons, shrinks to the lines
' the message takes, and hands back the button chosen (SetKitMsgAnswer).
' Enter picks the default button and Esc the last one (Cancel, No, or OK).
' A message longer than the dialog's 14 lines keeps its full height and
' starts with the focus in the message, which shows its scroll bar, so no
' part of it is cut off.

Private m_answers As Variant    ' what each visible button returns, btn1 first

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim style As Long, n As Long, i As Long, btnY As Long, x As Long, msgH As Long
    style = KitMsgStyle()
    Me.Caption = IIf(Len(KitMsgTitle()) > 0, KitMsgTitle(), APP_NAME)
    ShowIcon style And &H70
    Me.txtMsg = KitMsgText()

    ' Buttons, right-aligned in Windows order.
    Select Case style And 7
        Case vbOKCancel:    SetButtons Array("OK", "Cancel"), Array(vbOK, vbCancel)
        Case vbYesNo:       SetButtons Array("Yes", "No"), Array(vbYes, vbNo)
        Case vbYesNoCancel: SetButtons Array("Yes", "No", "Cancel"), Array(vbYes, vbNo, vbCancel)
        Case Else:          SetButtons Array("OK"), Array(vbOK)
    End Select
    n = UBound(m_answers) + 1
    x = Me.btn3.Left + Me.btn3.Width
    For i = n To 1 Step -1
        x = x - Me("btn" & i).Width
        Me("btn" & i).Left = x
        x = x - 120                 ' 8px between buttons
    Next
    i = ((style And &H300) \ &H100) + 1
    If i > n Then i = 1
    If i > 1 Then SwapLook Me.btn1, Me("btn" & i)   ' the default button looks primary
    Me("btn" & i).Default = True
    Me("btn" & n).Cancel = True
    msgH = MessageHeight()
    If msgH > Me.txtMsg.Height Then
        ' Too long for the dialog: focus the message so it can be scrolled
        ' (Enter still picks the default button, Esc the last).
        Me.txtMsg.SetFocus
        Me.txtMsg.SelStart = 0
        msgH = Me.txtMsg.Height
    Else
        Me("btn" & i).SetFocus
    End If

    ' Shrink to the message: move the buttons up under it, then the dialog.
    btnY = Me.txtMsg.Top + msgH + 300
    If btnY < Me.lblIcon.Top + Me.lblIcon.Height + 300 Then btnY = Me.lblIcon.Top + Me.lblIcon.Height + 300
    If btnY < Me.btn1.Top Then
        Me.txtMsg.Height = btnY - 300 - Me.txtMsg.Top
        For i = 1 To 3
            Me("btn" & i).Top = btnY
        Next
        Me.Section(acDetail).Height = btnY + Me.btn1.Height + 240
        Me.InsideHeight = Me.Section(acDetail).Height
    End If
    CenterOnApp Me                  ' at its final size, over Access (modKitUtil)
    Exit Sub
Err_Handler:
    ' Whatever failed, the dialog still shows the message and an OK button.
    LogError "dlgMessage.Form_Load", Left$(KitMsgText(), 60)
End Sub

' The icon for vbCritical, vbQuestion, vbExclamation or vbInformation (none
' for 0: the message moves left into its place).
Private Sub ShowIcon(ByVal kind As Long)
    Dim glyph As Long, clr As Long
    Select Case kind
        Case vbCritical:    glyph = &HEA39: clr = Th(CLR_ERROR)
        Case vbQuestion:    glyph = &HE9CE: clr = Th(CLR_PRIMARY)
        Case vbExclamation: glyph = &HE7BA: clr = Th(CLR_WARN_TEXT)
        Case vbInformation: glyph = &HE946: clr = Th(CLR_PRIMARY)
    End Select
    If glyph = 0 Then
        Me.lblIcon.Visible = False
        Me.txtMsg.Width = Me.txtMsg.Left + Me.txtMsg.Width - Me.lblIcon.Left
        Me.txtMsg.Left = Me.lblIcon.Left
    Else
        Me.lblIcon.Caption = ChrW(glyph)
        Me.lblIcon.ForeColor = clr
    End If
End Sub

' Swaps the colors of two buttons (the primary look moves to the default).
Private Sub SwapLook(a As Control, b As Control)
    Dim props As Variant, p As Variant, v As Long
    props = Array("BackColor", "ForeColor", "BorderColor", "HoverColor", "PressedColor", _
        "HoverForeColor", "PressedForeColor")
    For Each p In props
        v = a.Properties(CStr(p))
        a.Properties(CStr(p)) = b.Properties(CStr(p))
        b.Properties(CStr(p)) = v
    Next
End Sub

' Captions and answers for the first UBound(caps) + 1 buttons; the rest hide.
Private Sub SetButtons(ByVal caps As Variant, ByVal answers As Variant)
    Dim i As Long
    m_answers = answers
    For i = 1 To 3
        If i - 1 <= UBound(caps) Then
            Me("btn" & i).Caption = caps(i - 1)
        Else
            Me("btn" & i).Visible = False
        End If
    Next
End Sub

' Height the message's lines take (twips): each paragraph wraps to as many
' 10pt lines as its measured width needs, with 10% spare for word breaks.
Private Function MessageHeight() As Long
    Dim para As Variant, lines As Long, w As Long
    w = Me.txtMsg.Width
    For Each para In Split(Replace(KitMsgText(), vbCrLf, vbLf), vbLf)
        If Len(para) = 0 Then
            lines = lines + 1
        Else
            lines = lines - Int(-(MeasureText(CStr(para), 10, 400) * 1.1) / w)
        End If
    Next
    If lines < 1 Then lines = 1
    MessageHeight = lines * CLng(10 * 20 * 1.36) + 30
End Function

Private Sub Answer(ByVal i As Long)
    If Not IsArray(m_answers) Then m_answers = Array(vbOK)
    If i - 1 > UBound(m_answers) Then i = 1
    SetKitMsgAnswer m_answers(i - 1)
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub btn1_Click()
    On Error GoTo Err_Handler
    Answer 1
    Exit Sub
Err_Handler:
    HandleError "dlgMessage.btn1_Click"
End Sub

Private Sub btn2_Click()
    On Error GoTo Err_Handler
    Answer 2
    Exit Sub
Err_Handler:
    HandleError "dlgMessage.btn2_Click"
End Sub

Private Sub btn3_Click()
    On Error GoTo Err_Handler
    Answer 3
    Exit Sub
Err_Handler:
    HandleError "dlgMessage.btn3_Click"
End Sub
