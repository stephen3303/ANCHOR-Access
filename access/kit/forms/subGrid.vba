Option Compare Database
Option Explicit

' Rows of a styled table (modUIKit.BuildGrid; shared by every subGrid* form).
' Clicking a row selects it and tells the parent form, whose
' GridClick(grid, dbl) gets the subform control's name and whether it was a
' double click. A right-click selects the row the same way, then calls the
' parent's GridRightClick(grid) when it has one, which opens the row's
' actions with KitMenuAt (a parent without it ignores right-clicks).

Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_RBUTTON As Long = 2

Private m_right As Boolean      ' the button going down on a row was the right one

Public Function RowClick()
    SelectRow False
End Function

Public Function RowDblClick()
    SelectRow True
End Function

' A cell's MouseDown: an expression handler gets no Button, so the right
' button is read from the keyboard state while its press is handled.
Public Function RowMouseDown()
    m_right = (GetKeyState(VK_RBUTTON) < 0)
End Function

' A cell's MouseUp: after a right press, select the row and let the parent
' open its menu.
Public Function RowMouseUp()
    Dim errNo As Long, errDesc As String, errSrc As String
    If Not m_right Then Exit Function
    m_right = False
    On Error GoTo Err_Handler
    SelectRow False
    If IsNull(Me!id) Then Exit Function
    On Error Resume Next
    Me.Parent.GridRightClick GridName()
    errNo = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    On Error GoTo Err_Handler
    Select Case errNo
        Case 0, 438, 2465           ' 438, 2465: the parent has no GridRightClick
        Case Else: Err.Raise errNo, errSrc, errDesc
    End Select
    Exit Function
Err_Handler:
    HandleError Me.Name & ".RowMouseUp"
End Function

Private Sub SelectRow(ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    Me.txtFocusSink.SetFocus   ' a focused cell is highlighted and painted white
    If IsNull(Me!id) Then Exit Sub
    Me.txtSel = Me!id
    Me.Recalc
    Me.Parent.GridClick GridName(), dbl
    Exit Sub
Err_Handler:
    HandleError Me.Name & ".SelectRow"
End Sub

' Name of the parent's subform control that shows this form.
Private Function GridName() As String
    Dim c As Control
    For Each c In Me.Parent.Controls
        If c.ControlType = acSubform Then
            If Len(c.SourceObject) > 0 Then
              If c.Form Is Me Then
                GridName = c.Name
                Exit Function
              End If
            End If
        End If
    Next
End Function
