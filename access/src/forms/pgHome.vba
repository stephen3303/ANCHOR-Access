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
' The tables load from the form's timer once the page is open and sized
' (Form_Load only sets the timer), so no subform is read or requeried while
' the shell is still loading the page, and each table's scroll bar is
' decided at its real height.
'
' Depends on: modMessages, modMessageView, modSignIn; the kit's modKitUI,
' modKitUtil.

Private Const LATEST_N As Long = 8

Public Property Get HelpTopic() As String
    HelpTopic = "home"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    Me.TimerInterval = 1                    ' Form_Timer loads the page
    Exit Sub
Err_Handler:
    HandleError "pgHome.Form_Load"
End Sub

' Once loaded: the tiles and the tables, once.
Private Sub Form_Timer()
    On Error GoTo Err_Handler
    If Me.TimerInterval = 0 Then Exit Sub   ' a tick left over from before
    Me.TimerInterval = 0
    LoadPage
    Exit Sub
Err_Handler:
    HandleError "pgHome.Form_Timer"
End Sub

' The tiles and the three tables, or with no data file why they are empty.
Private Sub LoadPage()
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
    SetTile "tileAction", n, IIf(n = 1, "message with an open action", "messages with an open action")
    n = StatusChangeCount()
    SetTile "tileStatus", n, IIf(n = 1, "message changed status", "messages changed status")
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
