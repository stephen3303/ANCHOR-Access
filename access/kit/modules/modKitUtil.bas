Option Compare Database
Option Explicit

' Access UI kit - small file and form helpers the other kit modules use.

Private Const ERR_DISABLE_FOCUSED As Long = 2164   ' can't disable the focused control
Private Const ERR_HIDE_FOCUSED As Long = 2165      ' can't hide the focused control

' Window placement for CenterOnApp (user32). Rectangles are screen pixels.
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT      ' the monitor less the taskbar
    dwFlags As Long
End Type

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemoryW Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' True when a file exists.
Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(path) > 0) And (Len(Dir$(path, vbNormal)) > 0)
End Function

' True when a folder exists.
Public Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(path) > 0) And ((GetAttr(path) And vbDirectory) = vbDirectory)
End Function

' Create a folder and any missing parents.
Public Sub EnsureFolder(ByVal path As String)
    Dim parts() As String, i As Long, cur As String
    If FolderExists(path) Then Exit Sub
    parts = Split(path, "\")
    cur = parts(0)
    Dim first As Long
    first = 1
    If Left$(path, 2) = Chr$(92) & Chr$(92) Then
        If UBound(parts) < 3 Then Err.Raise 5, , "A UNC folder needs a server and share"
        cur = Chr$(92) & Chr$(92) & parts(2) & Chr$(92) & parts(3)
        If Not FolderExists(cur) Then Err.Raise 76, , "Network share unavailable: " & cur
        first = 4
    End If
    For i = first To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If Not FolderExists(cur) Then MkDir cur
        End If
    Next
End Sub

' True when the form is open in any view.
Public Function IsFormOpen(ByVal name As String) As Boolean
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, name) <> 0)
End Function

' Show or hide a control. Access raises error 2165 when hiding the control that
' has the focus, or a subform, tab control or option group that holds it (a
' button that hides itself, a tab switch that moved the focus onto the page).
' Try first; only on 2165 move the focus away and try again. When nothing else
' on the form can take the focus, disabling c drops it (Access allows that),
' then c is hidden and its Enabled setting put back.
' A rounded pill's image (<name>_bg) goes with it, see PillImageVisible.
Public Sub SetVisible(c As Control, ByVal v As Boolean)
    Dim wasEnabled As Boolean
    PillImageVisible c, v
    If c.Visible = v Then Exit Sub
    On Error GoTo Focused
    c.Visible = v
    Exit Sub
Focused:
    If Err.Number <> ERR_HIDE_FOCUSED Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Park
Park:
    On Error GoTo 0
    If ParkFocus(c) Then
        c.Visible = v
    Else
        wasEnabled = c.Enabled
        c.Enabled = False
        c.Visible = v
        c.Enabled = wasEnabled
    End If
End Sub

' The pill image <name>_bg behind c (RoundCF, PillLabel), if c has one, follows
' c: hiding c hides it, so a card that is closed or cleared leaves no empty
' pill behind. Showing c shows a bound image again (its ControlSource picks
' the color, or none); a painted image stays as the app's paint left it,
' since only a fresh paint knows its color.
Private Sub PillImageVisible(c As Control, ByVal v As Boolean)
    Dim im As Control, src As String
    On Error Resume Next
    Set im = OwnerForm(c).Controls(c.Name & "_bg")
    ' Read through Properties: im.ControlSource raises 438 on an image control.
    src = Nz(im.Properties("ControlSource").Value, "")
    On Error GoTo 0
    If im Is Nothing Then Exit Sub
    If Not v Then
        im.Visible = False
    ElseIf Len(src) > 0 Then
        im.Visible = True
    End If
End Sub

' Enable or disable a control. Access raises error 2164 when disabling the
' control that has the focus (a button that disables itself when clicked), so
' this moves the focus away first, like SetVisible.
Public Sub SetEnabled(c As Control, ByVal v As Boolean)
    If c.Enabled = v Then Exit Sub
    On Error GoTo Focused
    c.Enabled = v
    Exit Sub
Focused:
    If Err.Number <> ERR_DISABLE_FOCUSED Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Park
Park:
    On Error GoTo 0
    ParkFocus c
    c.Enabled = v
End Sub

' Move the focus out of c: to the form's txtFocusSink, or, when the form has
' none or it can't take the focus, to the first tab stop that c doesn't hold.
' False when nothing could take it.
Private Function ParkFocus(c As Control) As Boolean
    Dim frm As Form, o As Control
    Set frm = OwnerForm(c)
    If c.Name <> "txtFocusSink" And HasFocusSink(frm) Then
        If Not Holds(c, frm.Controls("txtFocusSink")) Then
            If TryFocus(frm.Controls("txtFocusSink")) Then ParkFocus = True: Exit Function
        End If
    End If
    For Each o In frm.Controls
        If o.Name <> c.Name Then
            If Not Holds(c, o) Then
                If TryFocus(o) Then ParkFocus = True: Exit Function
            End If
        End If
    Next
End Function

' The form a control sits on (its Parent can be a tab page, an option group or,
' for an attached label, another control). For a subform control on a tab
' page, the host form whose code and controls the kit's table helpers use.
Public Function OwnerForm(c As Control) As Form
    Dim p As Object
    Set p = c.Parent
    Do Until TypeOf p Is Form
        Set p = p.Parent
    Loop
    Set OwnerForm = p
End Function

' True when o sits inside c (on one of its tab pages or in its option group).
' Names are unique on a form, tab pages included, so they identify controls.
Private Function Holds(c As Control, o As Control) As Boolean
    Dim p As Object
    Set p = o.Parent
    Do Until TypeOf p Is Form
        If p.Name = c.Name Then Holds = True: Exit Function
        Set p = p.Parent
    Loop
End Function

Private Function HasFocusSink(frm As Form) As Boolean
    Dim o As Control
    On Error Resume Next
    Set o = frm.Controls("txtFocusSink")
    HasFocusSink = (Err.Number = 0)
End Function

' Give o the focus if it can take it. Labels, lines and images have no TabStop
' (reading it raises an error), and a control on a hidden page refuses SetFocus.
Private Function TryFocus(o As Control) As Boolean
    On Error GoTo Nope
    If o.TabStop And o.Visible And o.Enabled Then
        o.SetFocus
        TryFocus = True
    End If
Nope:
End Function

' Open a modal dialog and report whether it was confirmed. The app behind it
' is dimmed while it is open (DimApp, modBackdrop) unless dimBehind is False.
' The dim is removed when the dialog closes, and also when it fails to open,
' before the error goes back to the caller. A dialog that cancels its own
' open in Form_Open (error 2501) ends as not confirmed, with no error.
Public Function RunDialog(ByVal formName As String, Optional ByVal openArgs As String = "", _
                          Optional ByVal dimBehind As Boolean = True) As Boolean
    Dim dimmed As Boolean, num As Long, desc As String, src As String
    On Error GoTo Fail
    TempVars("DialogResult") = ""
    If dimBehind Then DimApp: dimmed = True
    DoCmd.OpenForm formName, , , , , acDialog, openArgs
    If dimmed Then UndimApp: dimmed = False
    RunDialog = (Nz(TempVars("DialogResult"), "") = "ok")
    Exit Function
Fail:
    ' UndimApp resets Err, so keep the error to pass it on.
    num = Err.Number: desc = Err.Description: src = Err.Source
    If dimmed Then UndimApp
    If num = 2501 Then Exit Function            ' the dialog cancelled its own open (Form_Open): not ok
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modKitUtil.RunDialog", formName
End Function

' Centre a pop-up form over the Access window, on the Access window's monitor,
' and keep it inside that monitor's work area. Call it at the end of every
' dialog's Form_Load, after the dialog has its final size. AutoCenter alone
' centres a pop-up only roughly (dlgMessage sat about 75px high after it
' shrank) and can place it on another monitor. When Access is minimized the
' dialog centres on the monitor Access was on. Positioning never stops a
' dialog opening: a failure is logged and the dialog stays where Access put it.
' With monitors at different scaling, Windows may resize a dialog it moves to
' the other monitor, so it can then sit a few pixels off centre.
Public Sub CenterOnApp(frm As Form)
    Dim app As RECT, dlg As RECT, mi As MONITORINFO, hApp As LongPtr, x As Long, y As Long
    On Error GoTo Fail
    hApp = Application.hWndAccessApp
    If GetWindowRect(frm.hWnd, dlg) = 0 Then Exit Sub
    mi.cbSize = LenB(mi)
    If GetMonitorInfo(MonitorFromWindow(hApp, MONITOR_DEFAULTTONEAREST), mi) = 0 Then Exit Sub
    If IsIconic(hApp) <> 0 Then
        app = mi.rcWork
    ElseIf GetWindowRect(hApp, app) = 0 Then
        Exit Sub
    End If
    x = CenteredStart(app.Left, app.Right, dlg.Right - dlg.Left, mi.rcWork.Left, mi.rcWork.Right)
    y = CenteredStart(app.Top, app.Bottom, dlg.Bottom - dlg.Top, mi.rcWork.Top, mi.rcWork.Bottom)
    SetWindowPos frm.hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    Exit Sub
Fail:
    LogError "modKitUtil.CenterOnApp", frm.Name
End Sub

' Where a window size wide starts when centred on lo..hi, pulled back inside
' the work area workLo..workHi. A window wider than the work area starts at
' workLo, so its title bar and close button stay reachable. Public so the
' kit tests can check the arithmetic.
Public Function CenteredStart(ByVal lo As Long, ByVal hi As Long, ByVal size As Long, _
        ByVal workLo As Long, ByVal workHi As Long) As Long
    Dim p As Long
    p = lo + (hi - lo - size) \ 2
    If p + size > workHi Then p = workHi - size
    If p < workLo Then p = workLo
    CenteredStart = p
End Function

' Put s on the clipboard as text (a Copy button). True when it is there;
' False when another program kept the clipboard (the caller says so).
Public Function KitCopyText(ByVal s As String) As Boolean
    Const CF_UNICODETEXT As Long = 13, GMEM_MOVEABLE As Long = 2, TRIES As Long = 20
    Dim hMem As LongPtr, p As LongPtr, n As LongPtr, i As Long, opened As Boolean
    n = (Len(s) + 1) * 2
    hMem = GlobalAlloc(GMEM_MOVEABLE, n)
    If hMem = 0 Then Exit Function
    p = GlobalLock(hMem)
    If p = 0 Then GlobalFree hMem: Exit Function
    CopyMemoryW p, StrPtr(s & vbNullChar), n
    GlobalUnlock hMem
    ' Another program (a remote desktop's clipboard sync, a clipboard
    ' manager) can hold the clipboard for a moment: try for about a second,
    ' letting Windows run between tries so it can let go.
    For i = 1 To TRIES
        If OpenClipboard(Application.hWndAccessApp) <> 0 Then opened = True: Exit For
        DoEvents
        Sleep 50
    Next
    If Not opened Then GlobalFree hMem: Exit Function
    EmptyClipboard
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then GlobalFree hMem Else KitCopyText = True   ' Windows owns it once set
    CloseClipboard
End Function
