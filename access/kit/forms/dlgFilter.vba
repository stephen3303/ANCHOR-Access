Option Compare Database
Option Explicit

' Column filter for a table whose GridHeaders show funnels (GridFilter opens
' it; BuildFilterDialog builds it). Lists the column's values FILTER_ROWS at
' a time, each with a check box glyph: a ticked value shows. Clicking a value
' ticks or clears it; the arrows and the mouse wheel scroll; the keyboard
' (through txtFocusSink) moves the cursor row and Space ticks it. Select all
' ticks every value, or clears them all when every value is ticked. OK hands
' the ticks back (SetGridFilterChoice) and closes; Cancel or Esc closes
' without changing the table. OK needs at least one value ticked.
' The rows are labels the kit draws, not a list box: a list box shows its
' selection in Access's own colors, which no palette can change.

Private m_items As Variant      ' the values' text (GridFilterItems)
Private m_ticked() As Boolean   ' whether each value is ticked
Private m_top As Long           ' value shown in the first row
Private m_cur As Long           ' the cursor row's value (keyboard)

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim checked As Variant, i As Long
    Me.Caption = GridFilterTitle()
    Me.lblHint.Caption = GridFilterHint()
    m_items = GridFilterItems()
    checked = GridFilterChecked()
    ReDim m_ticked(UBound(m_items))      ' GridFilter opens this with at least one value
    For i = 0 To UBound(m_items)
        m_ticked(i) = checked(i)
    Next
    ShowRows
    Me.txtFocusSink.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.Form_Load"
End Sub

' Number of values listed.
Private Function ItemCount() As Long
    ItemCount = UBound(m_items) + 1
End Function

' Draw the rows from m_top: each value's text and tick, the cursor row's fill,
' the arrows (muted at either end), then the count and buttons.
Private Sub ShowRows()
    Dim r As Long, k As Long, has As Boolean
    For r = 0 To FILTER_ROWS - 1
        k = m_top + r
        has = k < ItemCount()
        SetVisible Me("lblVal" & r), has
        SetVisible Me("lblChk" & r), has
        SetVisible Me("boxRow" & r), has And k = m_cur
        If has Then
            Me("lblVal" & r).Caption = Replace(m_items(k), "&", "&&")   ' a label shows "&" as an underline
            Me("lblVal" & r).ControlTipText = Left$(m_items(k), 255)
            Me("lblChk" & r).Caption = ChrW(IIf(m_ticked(k), BOX_TICKED, BOX_EMPTY))
            Me("lblChk" & r).ForeColor = Th(IIf(m_ticked(k), CLR_PRIMARY, CLR_MUTED_FG))
        End If
    Next
    Me.lnkUp.ForeColor = Th(IIf(m_top > 0, CLR_FG, CLR_BORDER))
    Me.lnkDown.ForeColor = Th(IIf(m_top + FILTER_ROWS < ItemCount(), CLR_FG, CLR_BORDER))
    ShowCount
End Sub

' Show how many values are ticked, name Select all for what it will do, and
' enable OK only while a value is ticked.
Private Sub ShowCount()
    Dim n As Long, i As Long, total As Long
    total = ItemCount()
    For i = 0 To total - 1
        If m_ticked(i) Then n = n + 1
    Next
    Me.lblCount.Caption = n & " of " & total & IIf(total = 1, " value", " values") & " shown"
    Me.btnAll.Caption = IIf(n = total, "Clear all", "Select all")
    SetEnabled Me.btnOK, n > 0
End Sub

' Scroll so value first shows in the first row, as far as the list allows.
Private Sub ScrollTo(ByVal first As Long)
    If first > ItemCount() - FILTER_ROWS Then first = ItemCount() - FILTER_ROWS
    If first < 0 Then first = 0
    m_top = first
End Sub

' Move the cursor to value k (kept inside the list) and scroll it into view.
Private Sub MoveCursor(ByVal k As Long)
    If k > ItemCount() - 1 Then k = ItemCount() - 1
    If k < 0 Then k = 0
    m_cur = k
    If m_cur < m_top Then ScrollTo m_cur
    If m_cur >= m_top + FILTER_ROWS Then ScrollTo m_cur - FILTER_ROWS + 1
    ShowRows
End Sub

' A row's click (lblVal, lblChk or boxRow r): tick or clear its value.
Public Function PickRow(ByVal r As Long)
    On Error GoTo Err_Handler
    If m_top + r >= ItemCount() Then Exit Function
    m_cur = m_top + r
    m_ticked(m_cur) = Not m_ticked(m_cur)
    ShowRows
    Me.txtFocusSink.SetFocus
    Exit Function
Err_Handler:
    HandleError "dlgFilter.PickRow"
End Function

' An arrow's click: scroll a row up (-1) or down (1).
Public Function ScrollList(ByVal delta As Long)
    On Error GoTo Err_Handler
    ScrollTo m_top + delta
    ShowRows
    Exit Function
Err_Handler:
    HandleError "dlgFilter.ScrollList"
End Function

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    On Error GoTo Err_Handler
    ScrollTo m_top + IIf(Page, Sgn(Count) * FILTER_ROWS, Count)
    ShowRows
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.Form_MouseWheel"
End Sub

' The list's keys, while txtFocusSink has the focus.
Private Sub txtFocusSink_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Handler
    Select Case KeyCode
        Case vbKeyUp: MoveCursor m_cur - 1
        Case vbKeyDown: MoveCursor m_cur + 1
        Case vbKeyPageUp: MoveCursor m_cur - FILTER_ROWS
        Case vbKeyPageDown: MoveCursor m_cur + FILTER_ROWS
        Case vbKeyHome: MoveCursor 0
        Case vbKeyEnd: MoveCursor ItemCount() - 1
        Case vbKeySpace
            m_ticked(m_cur) = Not m_ticked(m_cur)
            ShowRows
        Case Else
            Exit Sub
    End Select
    KeyCode = 0                 ' handled: the box takes no text, the focus stays
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.txtFocusSink_KeyDown"
End Sub

' Tick every value, or clear them all when every value is ticked. The list
' stays where it is.
Private Sub btnAll_Click()
    On Error GoTo Err_Handler
    Dim i As Long, pick As Boolean
    For i = 0 To ItemCount() - 1
        If Not m_ticked(i) Then pick = True
    Next
    For i = 0 To ItemCount() - 1
        m_ticked(i) = pick
    Next
    ShowRows
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnAll_Click"
End Sub

Private Sub btnOK_Click()
    On Error GoTo Err_Handler
    Dim i As Long, choice As String
    For i = 0 To ItemCount() - 1
        choice = choice & IIf(m_ticked(i), "1", "0")
    Next
    SetGridFilterChoice choice
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgFilter.btnCancel_Click"
End Sub
