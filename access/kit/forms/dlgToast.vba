Option Compare Database
Option Explicit

' Toast for KitToast (BuildToast builds it): a short message in the
' bottom-right corner of frmMain that fades on its own. KitToast opens this
' form hidden and calls ShowToast, which places and shows the window through
' Windows without activating it, so the focus stays in the control the user
' is working in. A toast while one is showing replaces its message and starts
' its time again. Clicking the toast closes it. When Windows can't fade the
' window, it closes at the end of its time instead.

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

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
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const LWA_ALPHA As Long = 2
Private Const HWND_TOP As Long = 0
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const LINE_H As Long = 272          ' one line of 10pt text (twips)
Private Const MAX_LINES As Long = 3         ' the message box's design height
Private Const MARGIN As Long = 240          ' from frmMain's edges, 16px (twips)
Private Const FADE_STEP_MS As Long = 40     ' fade: 8 steps of 40 ms
Private Const FADE_STEPS As Long = 8

Private m_alpha As Long         ' opacity now, 255 while the toast holds
Private m_fading As Boolean     ' True once the hold time is over
Private m_layered As Boolean    ' the window can fade (WS_EX_LAYERED took)
Private m_styled As Boolean     ' the layered style was tried (once per open)

' Show msg for ms milliseconds, then fade and close. vrnt picks the icon and
' accent color (KitToast lists them). Errors go back to KitToast.
Public Sub ShowToast(ByVal msg As String, ByVal vrnt As String, ByVal ms As Long)
    Me.lblMsg.Caption = Replace(msg, "&", "&&")     ' a label takes & as an access key
    SetVariant vrnt
    FitHeight msg
    MakeLayered
    m_alpha = 255
    m_fading = False
    If m_layered Then SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
    Place
    Me.TimerInterval = ms
End Sub

' Icon and accent bar for the variant; anything unknown shows as info.
Private Sub SetVariant(ByVal vrnt As String)
    Dim glyph As Long, clr As Long
    Select Case LCase$(vrnt)
        Case "success":     glyph = &HE930: clr = Th(CLR_GREEN_600)
        Case "warning":     glyph = &HE7BA: clr = Th(CLR_WARN_TEXT)
        Case "destructive": glyph = &HEA39: clr = Th(CLR_ERROR)
        Case Else:          glyph = &HE946: clr = Th(CLR_PRIMARY)
    End Select
    Me.lblIcon.Caption = ChrW(glyph)
    Me.lblIcon.ForeColor = clr
    Me.boxAccent.BackColor = clr
End Sub

' Grow or shrink the toast to the lines the message takes (1 to 3), with the
' same padding below the message as above it. The section grows before the
' controls move down, and shrinks after they move up, since Access refuses a
' section smaller than its controls.
Private Sub FitHeight(ByVal msg As String)
    Dim msgH As Long, h As Long
    msgH = MessageLines(msg) * LINE_H
    h = Me.lblMsg.Top * 2 + msgH
    If h < Me.lblIcon.Top * 2 + Me.lblIcon.Height Then h = Me.lblIcon.Top * 2 + Me.lblIcon.Height
    If h > Me.Section(acDetail).Height Then Me.Section(acDetail).Height = h
    Me.lblMsg.Height = msgH
    Me.boxAccent.Height = h
    Me.ruleLeft.Height = h
    Me.ruleRight.Height = h
    Me.ruleBottom.Top = h               ' paints the last row (BuildToast)
    Me.Section(acDetail).Height = h
    Me.InsideHeight = h
End Sub

' Lines the message wraps to in lblMsg: each paragraph takes as many 10pt
' lines as its measured width needs, with 10% spare for word breaks (as
' dlgMessage measures). Text past MAX_LINES is cut off: toasts are short.
Private Function MessageLines(ByVal msg As String) As Long
    Dim para As Variant, n As Long
    For Each para In Split(Replace(msg, vbCrLf, vbLf), vbLf)
        If Len(para) = 0 Then
            n = n + 1
        Else
            n = n - Int(-(MeasureText(CStr(para), 10, 400) * 1.1) / Me.lblMsg.Width)
        End If
    Next
    If n < 1 Then n = 1
    If n > MAX_LINES Then n = MAX_LINES
    MessageLines = n
End Function

' Make the window layered so it can fade, and never activated, so even a
' click on it leaves the keyboard with the user's form. A window made layered
' stays invisible until its opacity is set, so if Windows refuses the opacity
' the layered style is taken off again and the toast simply closes when its
' time is up.
Private Sub MakeLayered()
    Dim st As LongPtr
    If m_styled Then Exit Sub
    m_styled = True
    st = GetWindowLongPtr(Me.hwnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE
    SetWindowLongPtr Me.hwnd, GWL_EXSTYLE, st Or WS_EX_LAYERED
    m_layered = (SetLayeredWindowAttributes(Me.hwnd, 0, 255, LWA_ALPHA) <> 0)
    If Not m_layered Then SetWindowLongPtr Me.hwnd, GWL_EXSTYLE, st
End Sub

' Show the window MARGIN in from the bottom-right corner of frmMain's client
' area (of the Access window when frmMain isn't open or is minimized), on
' top but not active. The client area leaves out the window frame, which on
' Windows 10 and 11 is partly invisible. Window rectangles are in screen
' pixels, so the margin is converted with the toast's own pixels-per-twip.
Private Sub Place()
    Dim rcHost As RECT, rcMe As RECT, pt As POINTAPI, host As LongPtr, w As Long, h As Long, m As Long
    host = Application.hWndAccessApp
    If IsFormOpen("frmMain") Then
        If IsIconic(Forms("frmMain").hwnd) = 0 Then host = Forms("frmMain").hwnd
    End If
    GetClientRect host, rcHost
    pt.x = rcHost.Right: pt.y = rcHost.Bottom
    ClientToScreen host, pt
    rcHost.Right = pt.x: rcHost.Bottom = pt.y
    GetWindowRect Me.hwnd, rcMe
    w = rcMe.Right - rcMe.Left
    h = rcMe.Bottom - rcMe.Top
    If Me.WindowWidth > 0 Then m = CLng(MARGIN * w / Me.WindowWidth) Else m = MARGIN \ 15
    SetWindowPos Me.hwnd, HWND_TOP, rcHost.Right - w - m, rcHost.Bottom - h - m, 0, 0, _
        SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

' First tick: the hold time is over, start fading. Then one step per tick,
' closing when the toast is clear (or at once when it can't fade).
Private Sub Form_Timer()
    On Error GoTo Err_Handler
    If Not m_fading Then
        m_fading = True
        Me.TimerInterval = FADE_STEP_MS
    End If
    m_alpha = m_alpha - 255 \ FADE_STEPS
    If m_alpha <= 0 Or Not m_layered Then
        CloseToast
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, CByte(m_alpha), LWA_ALPHA
    End If
    Exit Sub
Err_Handler:
    ' A toast must never interrupt: log it and close.
    LogError "dlgToast.Form_Timer"
    CloseToast
End Sub

Private Sub Detail_Click()
    CloseToast
End Sub

Private Sub lblMsg_Click()
    CloseToast
End Sub

Private Sub lblIcon_Click()
    CloseToast
End Sub

Private Sub boxAccent_Click()
    CloseToast
End Sub

' Close the toast. Errors are logged only: the toast has no message box.
Private Sub CloseToast()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    LogError "dlgToast.CloseToast"
End Sub
