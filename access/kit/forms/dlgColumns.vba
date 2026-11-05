Option Compare Database
Option Explicit

' Column chooser for a BuildGrid table (ChooseGridColumns opens it;
' BuildColumnsDialog builds it). Lists the table's columns in screen order
' with whether each shows. Move up and Move down reorder the selected column,
' Hide/Show or a double-click toggles it, Reset goes back to the design.
' Apply hands the layout back (SetGridColumnChoice) and closes; Cancel or Esc
' closes without changing the table. At least one column stays shown.

Private m_caps As Variant       ' captions in design order (GridColumnCaptions)
Private m_order() As Long       ' column indexes in screen order
Private m_shown() As Boolean    ' whether each column in m_order shows

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Dim n As Long
    m_caps = GridColumnCaptions()
    n = UBound(m_caps) + 1          ' at least 1: ChooseGridColumns checked the table
    If Not ParseLayout(GridColumnLayout(), n, m_order, m_shown) Then ParseLayout DefaultLayout(n), n, m_order, m_shown
    ShowList 0
    Me.lstCols.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.Form_Load"
End Sub

' Fill the list from m_order and m_shown and select the row at position k.
Private Sub ShowList(ByVal k As Long)
    Dim i As Long, rows As String
    For i = 0 To UBound(m_order)
        rows = rows & IIf(i > 0, ";", "") & i & ";" & Quoted(m_caps(m_order(i))) & ";" & IIf(m_shown(i), "Shown", "Hidden")
    Next
    Me.lstCols.RowSource = rows
    Me.lstCols = CStr(k)        ' the bound column is the position
    ShowButtons
End Sub

' A value-list item: quoted so a caption may hold ; or ".
Private Function Quoted(ByVal s As String) As String
    Quoted = """" & Replace(s, """", """""") & """"
End Function

' Position of the selected row, or -1.
Private Function SelectedPos() As Long
    SelectedPos = -1
    If Not IsNull(Me.lstCols) Then SelectedPos = CLng(Me.lstCols)
End Function

' Enable the buttons that apply to the selected row, and name the toggle.
Private Sub ShowButtons()
    Dim k As Long
    k = SelectedPos()
    SetEnabled Me.btnUp, k > 0
    SetEnabled Me.btnDown, k >= 0 And k < UBound(m_order)
    SetEnabled Me.btnToggle, k >= 0
    If k >= 0 Then Me.btnToggle.Caption = IIf(m_shown(k), "Hide", "Show")
End Sub

' Swap the selected column with its neighbour delta away (-1 up, 1 down).
Private Sub MoveSelected(ByVal delta As Long)
    Dim k As Long, i As Long, s As Boolean
    k = SelectedPos()
    If k + delta < 0 Or k + delta > UBound(m_order) Or k < 0 Then Exit Sub
    i = m_order(k): m_order(k) = m_order(k + delta): m_order(k + delta) = i
    s = m_shown(k): m_shown(k) = m_shown(k + delta): m_shown(k + delta) = s
    ShowList k + delta
End Sub

' Hide or show the selected column; the last shown column can't be hidden.
Private Sub ToggleSelected()
    Dim k As Long, i As Long, n As Long
    k = SelectedPos()
    If k < 0 Then Exit Sub
    If m_shown(k) Then
        For i = 0 To UBound(m_shown)
            If m_shown(i) Then n = n + 1
        Next
        If n = 1 Then
            KitMsgBox "A table needs at least one column. Show another column before hiding this one.", vbInformation, "Columns"
            Exit Sub
        End If
    End If
    m_shown(k) = Not m_shown(k)
    ShowList k
End Sub

Private Sub lstCols_AfterUpdate()
    On Error GoTo Err_Handler
    ShowButtons
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.lstCols_AfterUpdate"
End Sub

Private Sub lstCols_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler
    ToggleSelected
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.lstCols_DblClick"
End Sub

Private Sub btnUp_Click()
    On Error GoTo Err_Handler
    MoveSelected -1
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnUp_Click"
End Sub

Private Sub btnDown_Click()
    On Error GoTo Err_Handler
    MoveSelected 1
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnDown_Click"
End Sub

Private Sub btnToggle_Click()
    On Error GoTo Err_Handler
    ToggleSelected
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnToggle_Click"
End Sub

Private Sub btnReset_Click()
    On Error GoTo Err_Handler
    ParseLayout DefaultLayout(UBound(m_order) + 1), UBound(m_order) + 1, m_order, m_shown
    ShowList 0
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnReset_Click"
End Sub

Private Sub btnApply_Click()
    On Error GoTo Err_Handler
    Dim k As Long, layout As String
    For k = 0 To UBound(m_order)
        layout = layout & IIf(k > 0, ",", "") & m_order(k) & IIf(m_shown(k), "", "h")
    Next
    SetGridColumnChoice layout
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnApply_Click"
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgColumns.btnCancel_Click"
End Sub
