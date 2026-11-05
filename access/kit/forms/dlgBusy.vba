Option Compare Database
Option Explicit

' Busy pop-up for KitBusy (BuildBusy builds it): a status line in the middle
' of the dimmed app, with a short bar sliding back and forth along a track
' under it, like a web app's indeterminate progress bar. KitBusy opens this
' form hidden and calls ShowBusy, which centres the window over Access and
' shows it through Windows without activating it, so the focus stays in the
' control the user was in. The bar moves on the form's Timer, which runs
' only while the busy work lets Access run (DoEvents, as lock retries do);
' during work that doesn't, the message still shows. KitBusyEnd closes the
' form. A VBA reset in the middle of busy work leaves nothing to call it, so
' the first tick after a reset finds no KitBusy waiting (KitBusyDepth) and
' closes the form; its OnClose (=BackdropOnClose()) then clears the dim.
' While it shows, the Access window and the other pop-up forms are disabled,
' so clicks and keys can't start more work while the busy work lets Access
' run (the dim lets clicks through). The windows it disabled are listed in
' the form's Tag, which a VBA reset keeps, and Form_Unload enables them
' again however the form closes.

#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bEnable As Long) As Long
Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetFocusWnd Lib "user32" Alias "SetFocus" (ByVal hwnd As LongPtr) As LongPtr

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const HWND_TOP As Long = 0
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const LINE_H As Long = 272          ' one line of 10pt text (twips)
Private Const BAR_GAP As Long = 180         ' message to track, 12px (twips)
Private Const STEP_MS As Long = 40          ' the bar moves 25 times a second
Private Const STEP_TW As Long = 90          ' 6px a step (twips)

Private m_shown As Boolean      ' the window is showing (ShowBusy placed it)
Private m_dir As Long           ' the bar's direction: 1 right, -1 left

' Show msg (a new busy pop-up, or a new message on the one showing). The
' first call centres the window over Access and shows it without
' activating it, then starts the bar. Errors go back to KitBusy.
Public Sub ShowBusy(ByVal msg As String)
    Me.lblMsg.Caption = Replace(msg, "&", "&&")     ' a label takes & as an access key
    FitHeight msg
    If Not m_shown Then
        SetWindowLongPtr Me.hwnd, GWL_EXSTYLE, GetWindowLongPtr(Me.hwnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE
        CenterOnApp Me
        SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
        m_shown = True
        m_dir = 1
        Me.boxBar.Left = Me.boxTrack.Left
        BlockInput
        Me.TimerInterval = STEP_MS
    End If
    Me.Repaint                          ' paint now: the busy work may not let Access paint
End Sub

' One line for a message that fits the width, two for a longer one (the
' label cuts off anything past two: keep busy messages short), with the
' track 12px under it and the same padding below the track as above the
' message. The section grows before the controls move down, and shrinks
' after they move up, since Access refuses a section smaller than its
' controls.
Private Sub FitHeight(ByVal msg As String)
    Dim lines As Long, y As Long, h As Long
    lines = 1
    If MeasureText(msg, 10, 400) * 1.1 > Me.lblMsg.Width Then lines = 2     ' 10% spare for word breaks, as dlgToast
    y = Me.lblMsg.Top + lines * LINE_H + BAR_GAP
    h = y + Me.boxTrack.Height + Me.lblMsg.Top
    If h > Me.Section(acDetail).Height Then Me.Section(acDetail).Height = h
    Me.lblMsg.Height = lines * LINE_H
    Me.boxTrack.Top = y
    Me.boxBar.Top = y
    Me.ruleLeft.Height = h
    Me.ruleRight.Height = h
    Me.ruleBottom.Top = h               ' paints the last row (BuildBusy)
    Me.Section(acDetail).Height = h
    Me.InsideHeight = h
End Sub

' Slide the bar one step, turning at each end of the track. With no
' KitBusy waiting (a VBA reset), close instead.
Private Sub Form_Timer()
    Dim x As Long
    On Error GoTo Err_Handler
    If KitBusyDepth() = 0 Then
        CloseBusy
        Exit Sub
    End If
    x = Me.boxBar.Left + m_dir * STEP_TW
    If x <= Me.boxTrack.Left Then
        x = Me.boxTrack.Left: m_dir = 1
    ElseIf x + Me.boxBar.Width >= Me.boxTrack.Left + Me.boxTrack.Width Then
        x = Me.boxTrack.Left + Me.boxTrack.Width - Me.boxBar.Width: m_dir = -1
    End If
    Me.boxBar.Left = x
    Exit Sub
Err_Handler:
    ' The bar is decoration: log it and stop moving it.
    LogError "dlgBusy.Form_Timer"
    Me.TimerInterval = 0
End Sub

' Disable the Access window (and with it every form inside it) and each
' pop-up form open now, noting in the Tag the window with the keyboard
' focus and the windows disabled: "focus;hwnd;hwnd...".
Private Sub BlockInput()
    Dim f As Form, list As String
    list = CStr(GetFocus())
    For Each f In Forms
        If f.PopUp And f.Name <> Me.Name Then
            If IsWindowEnabled(f.hwnd) <> 0 Then
                EnableWindow f.hwnd, 0
                list = list & ";" & CStr(f.hwnd)
            End If
        End If
    Next
    If IsWindowEnabled(Application.hWndAccessApp) <> 0 Then
        EnableWindow Application.hWndAccessApp, 0
        list = list & ";" & CStr(Application.hWndAccessApp)
    End If
    Me.Tag = list
End Sub

' Enable the windows BlockInput disabled, and give the keyboard focus back
' when nothing has it.
Private Sub Form_Unload(Cancel As Integer)
    Dim p() As String, i As Long, h As LongPtr
    On Error GoTo Err_Handler
    If Len(Me.Tag) = 0 Then Exit Sub
    p = Split(Me.Tag, ";")
    Me.Tag = ""
    For i = 1 To UBound(p)
        h = CLngPtr(p(i))
        If IsWindow(h) <> 0 Then EnableWindow h, 1
    Next
    h = CLngPtr(p(0))
    If h <> 0 And GetFocus() = 0 Then
        If IsWindow(h) <> 0 Then SetFocusWnd h
    End If
    Exit Sub
Err_Handler:
    LogError "dlgBusy.Form_Unload"
End Sub

' Close the pop-up. Errors are logged only: the pop-up has no message box.
Private Sub CloseBusy()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    LogError "dlgBusy.CloseBusy"
End Sub
