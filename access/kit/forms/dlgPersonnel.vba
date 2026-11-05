Option Compare Database
Option Explicit

' The Personnel page for directory admins (BuildPersonnelDialogs builds it;
' KitPersonnelManage opens it). Lists everyone in the Personnel file; the
' search box narrows the list as you type. Add person and a double-click
' (or Edit) open dlgPerson; Import adds people from a CSV or Excel file.

' Only an active Personnel admin may open this, whoever opens it: the
' check is against the Personnel file now (PersonnelIsAdminNow). RunDialog
' treats the cancelled open as not ok. The builder wires OnOpen to this, in
' place of the usual =ThemeForm([Form]) hook, so it themes the form too.
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler
    ThemeForm Me
    If Not PersonnelIsAdminNow() Then
        KitMsgBox "Only an active Personnel admin can add or change people.", vbExclamation
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "dlgPersonnel.Form_Open"
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Handler
    LoadPeople ""
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.Form_Load"
End Sub

' F1 opens the help on this page (KitHelp).
Public Property Get HelpTopic() As String
    HelpTopic = "directory"
End Property

' F1 opens this dialog's help even during startup (PersonnelF1).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    PersonnelF1 KeyCode
End Sub

' Fill the list with the people matching term, keeping the selected person
' when they are still in it.
Private Sub LoadPeople(ByVal term As String)
    Dim keep As Variant
    keep = Me.lstPeople.Value
    Me.lstPeople.RowSource = PersonnelPeopleSql(term, PersonnelUiAppCode())
    ' With only the headings row left, the list is hidden (it would paint
    ' over the note, as in dlgHelp) and the note says so.
    SetVisible Me.lstPeople, Me.lstPeople.ListCount > 1
    SetVisible Me.lblNone, Me.lstPeople.ListCount <= 1
    Me.lstPeople.Value = Null
    If IsNull(keep) Then Exit Sub
    ' Restore the pick only when the search still shows that person, so Edit
    ' never opens someone the list no longer shows. Row 0 is the headings.
    Dim i As Long
    For i = 1 To Me.lstPeople.ListCount - 1
        If CStr(Me.lstPeople.ItemData(i)) = CStr(keep) Then
            Me.lstPeople.Value = keep
            Exit For
        End If
    Next i
End Sub

Private Sub txtSearch_Change()
    On Error GoTo Err_Handler
    LoadPeople Nz(Me.txtSearch.Text, "")
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.txtSearch_Change"
End Sub

Private Sub btnAdd_Click()
    On Error GoTo Err_Handler
    If RunDialog(PERSON_FORM, "new") Then LoadPeople Nz(Me.txtSearch, "")
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.btnAdd_Click"
End Sub

Private Sub lstPeople_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler
    EditSelected
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.lstPeople_DblClick"
End Sub

Private Sub btnEdit_Click()
    On Error GoTo Err_Handler
    EditSelected
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.btnEdit_Click"
End Sub

Private Sub EditSelected()
    If IsNull(Me.lstPeople.Value) Then KitMsgBox "Pick a person in the list first.", vbInformation: Exit Sub
    If RunDialog(PERSON_FORM, CStr(Me.lstPeople.Value)) Then LoadPeople Nz(Me.txtSearch, "")
End Sub

' Add people from a CSV or Excel file and say what happened.
Private Sub btnImport_Click()
    On Error GoTo Err_Handler
    Dim path As String, result As String
    path = KitPickFile("Import people", "People list|*.csv;*.xlsx;*.xls|CSV file|*.csv|Excel workbook|*.xlsx;*.xls", "personnel-import")
    If Len(path) = 0 Then Exit Sub
    DoCmd.Hourglass True
    result = PersonnelImport(path, PersonnelUiAppCode())
    DoCmd.Hourglass False
    LoadPeople Nz(Me.txtSearch, "")
    KitMsgBox result, vbInformation, "Import people"
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    If Err.Number = 5 Then KitMsgBox Err.Description, vbExclamation, "Import people": Exit Sub
    HandleError "dlgPersonnel.btnImport_Click", path, "The people couldn't be imported."
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPersonnel.btnClose_Click"
End Sub
