Option Compare Database
Option Explicit

' rptMessage (see Build_rptMessage): the official rendition of one message,
' the one OpenArgs names ("id=12"), as the Message page's Print opens it.
' The classification marking (with its demo annotation above UNCLASSIFIED)
' is the banner at the top and bottom of every page; then the heading
' block, the official text and the integrity block (modMessageView's
' MessagePrintHead and MessagePrintIntegrity). Opened with no id, or for a
' message ANCHOR doesn't hold, it says so and doesn't open.
'
' Depends on: modErrors, modMessages, modMessageView; the kit's modKitUI.

Private Sub Report_Open(Cancel As Integer)
    Dim msgId As Long
    On Error GoTo Err_Handler
    msgId = ArgId(Nz(Me.OpenArgs, ""))
    If msgId = 0 Then
        KitMsgBox "No message was named to print.", vbInformation, "Print"
        Cancel = True
        Exit Sub
    End If
    Me.RecordSource = "SELECT * FROM tblMessages WHERE MessageID=" & msgId
    Me.Caption = MessageRef(msgId)
    Exit Sub
Err_Handler:
    HandleError "rptMessage.Report_Open", Nz(Me.OpenArgs, "")
    Cancel = True
End Sub

' The message is not in the data file (deleted since the page showed it).
Private Sub Report_NoData(Cancel As Integer)
    On Error GoTo Err_Handler
    KitMsgBox "ANCHOR does not hold this message, so there is nothing to print.", vbInformation, "Print"
    Cancel = True
    Exit Sub
Err_Handler:
    HandleError "rptMessage.Report_NoData"
    Cancel = True
End Sub

' The whole number after "id=" in args, 0 when there is none.
Private Function ArgId(ByVal args As String) As Long
    Dim part As Variant, v As String
    For Each part In Split(args, ";")
        If LCase$(Left$(part, 3)) = "id=" Then v = Mid$(part, 4)
    Next
    If Len(v) > 0 And Len(v) <= 9 And Not v Like "*[!0-9]*" Then ArgId = CLng(v)
End Function
