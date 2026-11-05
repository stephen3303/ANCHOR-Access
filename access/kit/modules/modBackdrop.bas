Option Compare Database
Option Explicit

' Access UI kit - a dimmed backdrop behind modal dialogs, like a web modal.
' DimApp lays a see-through layer of the palette's CLR_SIDEBAR color over the
' Access window and UndimApp takes it away. RunDialog, KitMsgBox and KitInputBox
' call both around their dialog, so apps rarely call them directly.
'
' The layer is a plain Windows window, not an Access form, so there is
' nothing to build: a "Static" window that is layered, never takes the focus,
' stays out of the taskbar and is owned by the Access window (it sits above
' Access, below the dialog opened after it, and closes with Access). Its
' color is handed over once with UpdateLayeredWindow, so it never paints.
' A Static window lets mouse clicks fall through to the window under it, so a
' click on the dim reaches Access, which is blocked by the modal dialog, just
' as a click outside an undimmed dialog is.
'
' Nested dialogs (a message from inside a dialog) share the first one's
' layer: a count keeps it until the last UndimApp. A VBA reset (an End, or
' Reset after an error) while a dialog is open clears that count, and the
' code that would call UndimApp never runs. So every dialog built with
' NewDialog also calls BackdropOnClose as it closes: with no DimApp waiting,
' any layer left is an orphan, and it is found by its fixed title and removed.

#If VBA7 Then
Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal exStyle As Long, ByVal className As LongPtr, _
    ByVal windowName As LongPtr, ByVal style As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
    ByVal owner As LongPtr, ByVal menu As LongPtr, ByVal instance As LongPtr, ByVal param As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal parent As LongPtr, ByVal after As LongPtr, _
    ByVal className As LongPtr, ByVal windowName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal cmd As Long) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef r As RectL) As Long
Private Declare PtrSafe Function DwmGetWindowAttribute Lib "dwmapi" (ByVal hwnd As LongPtr, ByVal attr As Long, _
    ByRef r As RectL, ByVal size As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal cmd As Long) As Long
Private Declare PtrSafe Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal dcDst As LongPtr, _
    ByRef ptDst As PointL, ByRef size As SizeL, ByVal dcSrc As LongPtr, ByRef ptSrc As PointL, ByVal key As Long, _
    ByRef blendFn As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal dc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal dc As LongPtr, ByVal w As Long, ByVal h As Long) As LongPtr
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal color As Long) As LongPtr
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal dc As LongPtr, ByRef r As RectL, ByVal brush As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal dc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal obj As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal dc As LongPtr) As Long
#Else
Private Declare Function CreateWindowExW Lib "user32" (ByVal exStyle As Long, ByVal className As Long, _
    ByVal windowName As Long, ByVal style As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
    ByVal owner As Long, ByVal menu As Long, ByVal instance As Long, ByVal param As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowExW Lib "user32" (ByVal parent As Long, ByVal after As Long, _
    ByVal className As Long, ByVal windowName As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDpiForWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef r As RectL) As Long
Private Declare Function DwmGetWindowAttribute Lib "dwmapi" (ByVal hwnd As Long, ByVal attr As Long, _
    ByRef r As RectL, ByVal size As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal dcDst As Long, _
    ByRef ptDst As PointL, ByRef size As SizeL, ByVal dcSrc As Long, ByRef ptSrc As PointL, ByVal key As Long, _
    ByRef blendFn As Long, ByVal flags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal dc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal dc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal dc As Long, ByVal w As Long, ByVal h As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal color As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal dc As Long, ByRef r As RectL, ByVal brush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal dc As Long, ByVal obj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal obj As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal dc As Long) As Long
#End If

Private Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PointL
    x As Long
    y As Long
End Type

Private Type SizeL
    cx As Long
    cy As Long
End Type

' How strongly the app is dimmed: the layer's opacity, 0-255 (115 = 45%).
Private Const BACKDROP_ALPHA As Long = 115
' Title that marks the layer, so it can be found again (RemoveLayers).
Private Const BACKDROP_TITLE As String = "AccessKitBackdrop"

Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80          ' not in the taskbar or Alt+Tab
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_NOACTIVATE As Long = &H8000000     ' never takes the focus
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const ULW_ALPHA As Long = 2
Private Const GW_OWNER As Long = 4
Private Const DWMWA_EXTENDED_FRAME_BOUNDS As Long = 9

' DimApp calls still waiting for their UndimApp.
Private m_depth As Long

' Dim the Access window behind the dialog about to open. Pair every DimApp
' with one UndimApp, on the error path too. Nothing is drawn while Access is
' hidden or minimized. The dim is decoration, so a failure to draw it is
' logged and the dialog still opens.
Public Sub DimApp()
    m_depth = m_depth + 1
    If m_depth > 1 Then Exit Sub        ' a dialog over a dialog: the layer is already there
    On Error GoTo Fail
    RemoveLayers                        ' one left behind by a VBA reset
    ShowLayer Application.hWndAccessApp, Th(CLR_SIDEBAR)
    Exit Sub
Fail:
    LogError "modBackdrop.DimApp"
End Sub

' Remove the dim once the last dialog over it has closed. Never raises, so it
' is safe in an error handler; it does reset Err, so read Err first.
Public Sub UndimApp()
    If m_depth > 0 Then m_depth = m_depth - 1
    If m_depth > 0 Then Exit Sub
    On Error Resume Next
    RemoveLayers
End Sub

' A dialog's OnClose (NewDialog sets it to =BackdropOnClose()). In normal use
' the DimApp that opened the dialog is still waiting for its UndimApp, so this
' does nothing; after a VBA reset nothing is waiting, and the orphaned layer
' is removed so Access isn't left dimmed. A dialog with its own Form_Close
' replaces this; call BackdropOnClose from it to keep the clean-up.
Public Function BackdropOnClose()
    If m_depth > 0 Then Exit Function
    On Error Resume Next
    RemoveLayers
End Function

' True while a DimApp is waiting for its UndimApp.
Public Function AppDimmed() As Boolean
    AppDimmed = (m_depth > 0)
End Function

' Create the layer over the owner window's visible frame, fill it with color
' and show it without taking the focus. GDI objects are freed on every path.
#If VBA7 Then
Private Sub ShowLayer(ByVal owner As LongPtr, ByVal color As Long)
    Dim wnd As LongPtr, scr As LongPtr, mem As LongPtr, bmp As LongPtr, old As LongPtr, brush As LongPtr
#Else
Private Sub ShowLayer(ByVal owner As Long, ByVal color As Long)
    Dim wnd As Long, scr As Long, mem As Long, bmp As Long, old As Long, brush As Long
#End If
    Dim r As RectL, vis As RectL, fill As RectL, sz As SizeL, at As PointL, origin As PointL
    Dim cls As String, ttl As String, blendFn As Long, drawn As Long
    Dim num As Long, desc As String
    If owner = 0 Then Exit Sub
    If IsWindowVisible(owner) = 0 Or IsIconic(owner) <> 0 Then Exit Sub
    ' GetWindowRect includes the invisible resize border (about 7px a side),
    ' which would show as a dim band around the window, or spill onto the next
    ' monitor when Access is maximized. DWM gives the visible frame, but in
    ' physical pixels: when Windows scales Access (DPI virtualization) the two
    ' rectangles differ by much more than a border, and the window rectangle,
    ' which CreateWindowExW also uses, is kept as it is.
    If GetWindowRect(owner, r) = 0 Then Exit Sub
    If DwmGetWindowAttribute(owner, DWMWA_EXTENDED_FRAME_BOUNDS, vis, LenB(vis)) = 0 Then
        If IsBorderInset(r, vis, WindowDpi(owner)) Then r = vis
    End If
    sz.cx = r.Right - r.Left
    sz.cy = r.Bottom - r.Top
    If sz.cx <= 0 Or sz.cy <= 0 Then Exit Sub
    at.x = r.Left
    at.y = r.Top
    fill.Right = sz.cx
    fill.Bottom = sz.cy
    ' BLENDFUNCTION packed in a Long: AC_SRC_OVER, no flags, constant alpha, no per-pixel alpha.
    blendFn = BACKDROP_ALPHA * &H10000

    On Error GoTo Fail
    cls = "Static"
    ttl = BACKDROP_TITLE
    wnd = CreateWindowExW(WS_EX_LAYERED Or WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE, StrPtr(cls), StrPtr(ttl), _
        WS_POPUP, r.Left, r.Top, sz.cx, sz.cy, owner, 0, 0, 0)
    If wnd = 0 Then Err.Raise 5, , "The backdrop window could not be created"
    scr = GetDC(0)
    mem = CreateCompatibleDC(scr)
    bmp = CreateCompatibleBitmap(scr, sz.cx, sz.cy)
    If mem = 0 Or bmp = 0 Then Err.Raise 7, , "No memory for the backdrop (" & sz.cx & "x" & sz.cy & ")"
    old = SelectObject(mem, bmp)
    brush = CreateSolidBrush(color)
    If old = 0 Or brush = 0 Then Err.Raise 5, , "The backdrop could not be painted"
    FillRect mem, fill, brush
    drawn = UpdateLayeredWindow(wnd, scr, at, sz, mem, origin, 0, blendFn, ULW_ALPHA)
    If drawn = 0 Then Err.Raise 5, , "The backdrop could not be drawn"
    ShowWindow wnd, SW_SHOWNOACTIVATE
Done:
    On Error GoTo 0
    ' The layer keeps its own copy of the pixels, so the bitmap can go now.
    If old <> 0 Then SelectObject mem, old
    If brush <> 0 Then DeleteObject brush
    If bmp <> 0 Then DeleteObject bmp
    If mem <> 0 Then DeleteDC mem
    If scr <> 0 Then ReleaseDC 0, scr
    If num <> 0 Then
        If wnd <> 0 Then DestroyWindow wnd
        Err.Raise num, "modBackdrop.ShowLayer", desc
    End If
    Exit Sub
Fail:
    num = Err.Number
    desc = Err.Description
    Resume Done
End Sub

' True when the visible frame vis sits inside the window rectangle r by no
' more than a resize border on every side, so both are in the same pixels.
' The border grows with the window's DPI (8px at 100%, 20px at 250%), so the
' limit does too: 16px at 96 DPI, a border with room to spare.
Private Function IsBorderInset(r As RectL, vis As RectL, ByVal dpi As Long) As Boolean
    Dim maxB As Long
    maxB = 16 * dpi \ 96
    IsBorderInset = InRange(vis.Left - r.Left, maxB) And InRange(vis.Top - r.Top, maxB) And _
        InRange(r.Right - vis.Right, maxB) And InRange(r.Bottom - vis.Bottom, maxB)
End Function

' The DPI Windows reports for the window: 96 when Windows scales Access
' (DPI virtualization), or before Windows 10 1607, which has no GetDpiForWindow.
#If VBA7 Then
Private Function WindowDpi(ByVal wnd As LongPtr) As Long
#Else
Private Function WindowDpi(ByVal wnd As Long) As Long
#End If
    On Error Resume Next
    WindowDpi = GetDpiForWindow(wnd)
    If WindowDpi < 96 Then WindowDpi = 96
End Function

Private Function InRange(ByVal n As Long, ByVal maxN As Long) As Boolean
    InRange = (n >= 0 And n <= maxN)
End Function

' Destroy every backdrop layer this Access window owns: the current one and
' any a VBA reset left behind. Another Access window's layers are left alone.
Private Sub RemoveLayers()
#If VBA7 Then
    Dim wnd As LongPtr, nxt As LongPtr, owner As LongPtr
#Else
    Dim wnd As Long, nxt As Long, owner As Long
#End If
    Dim cls As String, ttl As String
    owner = Application.hWndAccessApp
    cls = "Static"
    ttl = BACKDROP_TITLE
    wnd = FindWindowExW(0, 0, StrPtr(cls), StrPtr(ttl))
    Do While wnd <> 0
        ' Find the next one before this one is gone: a destroyed window can't be searched past.
        nxt = FindWindowExW(0, wnd, StrPtr(cls), StrPtr(ttl))
        If GetWindow(wnd, GW_OWNER) = owner Then DestroyWindow wnd
        wnd = nxt
    Loop
End Sub
