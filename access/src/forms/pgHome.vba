Option Compare Database
Option Explicit

' Home (see Build_pgHome). Four tiles: the signed-in person's unread
' messages (modMessages.UnreadCount), the acknowledgements they still owe
' (PendingAckCount), the action items that still stand and the messages
' whose status changed (modMessageView). Under them the messages that
' require their acknowledgement, the soonest due first, with how long is
' left (AckDueText), and the 8 newest messages; on the right the messages
' no longer ACTIVE, newest first. A click on a row opens its message on the
' Message page (OpenMessage), where it is acknowledged.
'
' Depends on: modMessages, modMessageView, modSignIn; the kit's modKitUI,
' modKitUtil.

Private Const LATEST_N As Long = 8

Public Property Get HelpTopic() As String
    HelpTopic = "home"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        Me.lblSubtitle.Caption = ""
        SetTile "tileUnread", "-", ""
        SetTile "tileAck", "-", ""
        SetTile "tileAction", "-", ""
        SetTile "tileStatus", "-", ""
        ShowHomeList Me.sfAck, "emptyAck", -1
        ShowHomeList Me.sfLatest, "emptyLatest", -1
        ShowHomeList Me.sfStatus, "emptyStatus", -1
        Exit Sub
    End If
    Me.lblSubtitle.Caption = Replace("Welcome back, " & SignedInName() & ".", "&", "&&")
    ShowTiles
    Me.sfAck.Form.RecordSource = HomeAckSql()
    ShowHomeList Me.sfAck, "emptyAck", Me.sfAck.Form.RecordsetClone.RecordCount
    Me.sfLatest.Form.RecordSource = HomeLatestSql(LATEST_N)
    ShowHomeList Me.sfLatest, "emptyLatest", Me.sfLatest.Form.RecordsetClone.RecordCount
    Me.sfStatus.Form.RecordSource = HomeStatusSql()
    ShowHomeList Me.sfStatus, "emptyStatus", Me.sfStatus.Form.RecordsetClone.RecordCount
    Exit Sub
Err_Handler:
    HandleError "pgHome.Form_Load"
End Sub

' A row clicked: open its message.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfAck", "sfLatest", "sfStatus": id = Nz(Me.Controls(grid).Form.txtSel, 0)
    End Select
    If id <> 0 Then OpenMessage id
    Exit Sub
Err_Handler:
    HandleError "pgHome.GridClick", grid & " " & id
End Sub

' The four tiles, and how many of the person's acknowledgements are done.
Private Sub ShowTiles()
    Dim n As Long, total As Long, done As Long
    n = UnreadCount()
    SetTile "tileUnread", n, IIf(n = 1, "message you haven't read", "messages you haven't read")
    n = PendingAckCount()
    SetTile "tileAck", n, "designated traffic waiting on you"
    n = OpenActionCount()
    SetTile "tileAction", n, IIf(n = 1, "message with an action that stands", "messages with an action that stands")
    n = StatusChangeCount()
    SetTile "tileStatus", n, "amended, corrected, superseded, cancelled or expired"
    OwnAckProgress total, done
    Me.lblAckDone.Caption = IIf(total = 0, "", done & " of " & total & " complete")
End Sub

Private Sub SetTile(ByVal nm As String, ByVal v As Variant, ByVal cap As String)
    Me.Controls(nm).Value = v
    Me.Controls(nm & "_cap").Caption = Replace(cap, "&", "&&")
End Sub

' A table with n rows, or its empty state when it has none; n -1: no data
' file (the empty state says so).
Private Sub ShowHomeList(ByVal sf As Control, ByVal emptyNm As String, ByVal n As Long)
    If n < 0 Then
        ShowEmpty Me, emptyNm, True, "No data file linked", "This copy was built without the data file, so there is nothing to show."
    Else
        ShowEmpty Me, emptyNm, n = 0
    End If
    SetVisible sf, n > 0
End Sub
