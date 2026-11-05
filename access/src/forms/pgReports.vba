Option Compare Database
Option Explicit

' Reports (see modBuildReports.Build_pgReports): acknowledgement
' accountability and traffic measures over a period, the last 7 or 30 days
' or all traffic.
'
' Any change of period works the accountability table out again
' (modReports.RptRun) and recounts the traffic. The tiles: the messages
' received in the period, the share of the acknowledgements owed that were
' made, how many are still owed, and how many messages are overdue. Picking
' a row lists who acknowledged it and who still owes it under the table
' (RptDetailFill); Open message, or a double-click, reads it. The traffic
' card counts the period's messages by precedence, type (named from
' tblMessageTypes), classification or lifecycle. Export CSV writes the
' accountability table and audits the export.
'
' Read only: nothing here changes the data file. With no data file linked
' the page says so and shows no figures.
'
' Depends on: modMessages, modReports, modSignIn, modUI; the kit's modKitUI.

Private m_period As String      ' the period the figures were worked out for (RPT_PERIOD_*)
Private m_picked As Long        ' the message whose acknowledgements show (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "reports"
End Property

Private Sub Form_Load()
    Dim c As Variant
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        For Each c In Array("tileMsgs", "tileRate", "tileOwed", "tileOverdue")
            Me.Controls(CStr(c)).Value = Null
        Next
        Me.lblPeriod.Caption = "This copy was built without the data file, so there are no figures to show."
        Me.lblDet.Caption = ""
        Me.segPeriod_val.Locked = True      ' SegmentClick ignores a locked value
        Me.segMeasure_val.Locked = True
        SetEnabled Me.btnExport, False
        SetEnabled Me.btnOpen, False
        ShowEmpty Me, "emptyAck", True, "No data file linked", "There are no messages to report on."
        SetVisible Me.sfAck, False
        SetVisible Me.sfDet, False
        ShowDetHeaders False
        ShowEmpty Me, "emptyTraffic", True, "No data file linked", "There is no traffic to count."
        SetVisible Me.sfTraffic, False
        Exit Sub
    End If
    SetSegment Me, "segPeriod", RPT_PERIOD_30
    SetSegment Me, "segMeasure", "prec"
    RefreshAll
    Exit Sub
Err_Handler:
    HandleError "pgReports.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' The period and the measure
' ---------------------------------------------------------------------------

' The period's Segmented control (onChange).
Public Sub PeriodChanged()
    On Error GoTo Err_Handler
    RefreshAll
    Exit Sub
Err_Handler:
    HandleError "pgReports.PeriodChanged"
End Sub

' The traffic card's Segmented control (onChange).
Public Sub MeasureChanged()
    On Error GoTo Err_Handler
    RefreshTraffic
    Exit Sub
Err_Handler:
    HandleError "pgReports.MeasureChanged"
End Sub

' ---------------------------------------------------------------------------
' The figures
' ---------------------------------------------------------------------------

' Work everything out for the period picked: the accountability table, the
' tiles and the traffic.
Private Sub RefreshAll()
    Dim n As Long, nAcked As Long, nOwed As Long, nOverdue As Long
    m_period = SegmentValue(Me, "segPeriod")
    Me.lblPeriod.Caption = RptPeriodText(m_period) & "."
    n = RptRun(m_period)
    RptTotals nAcked, nOwed, nOverdue
    Me.tileMsgs = RptMessageCount(m_period)
    If nOwed > 0 Then Me.tileRate = nAcked / nOwed Else Me.tileRate = "-"   ' nothing owed: a dash, not a blank
    Me.tileOwed = nOwed - nAcked
    Me.tileOverdue = nOverdue
    Me.sfAck.Form.RecordSource = RptAckSql()
    ShowEmpty Me, "emptyAck", n = 0
    SetVisible Me.sfAck, n > 0
    SetVisible Me.sfDet, n > 0
    ShowDetHeaders n > 0
    SetEnabled Me.btnExport, n > 0
    SetEnabled Me.btnOpen, n > 0
    If GridSelect(Me.sfAck, m_picked) Then
        ShowAcks m_picked
    Else
        ShowAcks Nz(SelectFirstRow(Me.sfAck), 0)
    End If
    RefreshTraffic
End Sub

' The detail table's captions, hidden with it so they don't show through
' the accountability table's empty state.
Private Sub ShowDetHeaders(ByVal isShown As Boolean)
    Dim c As Control
    For Each c In Me.Controls
        If c.Name Like "hdrDet#*" Then SetVisible c, isShown
    Next
End Sub

' The acknowledgements of message msgId under the table (0: none).
Private Sub ShowAcks(ByVal msgId As Long)
    m_picked = msgId
    RptDetailFill msgId
    Me.sfDet.Form.RecordSource = RptDetailSql()
    If msgId = 0 Then
        Me.lblDet.Caption = ""
    Else
        Me.lblDet.Caption = UCase$(Replace(RptAckHeading(msgId), "&", "&&"))
    End If
End Sub

' The traffic counts for the period and the measure picked.
Private Sub RefreshTraffic()
    Dim total As Long, n As Long
    total = RptMessageCount(m_period)
    Me.sfTraffic.Form.RecordSource = RptTrafficSql(SegmentValue(Me, "segMeasure"), m_period, total)
    n = Me.sfTraffic.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyTraffic", n = 0
    SetVisible Me.sfTraffic, n > 0
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

' A row picked: an accountability row shows its acknowledgements; a
' double-click on it opens the message.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If grid = "sfAck" Then
        If dbl Then
            OpenMessage Nz(Me.sfAck.Form.txtSel, 0)
        Else
            ShowAcks Nz(Me.sfAck.Form.txtSel, 0)
        End If
    End If
    Exit Sub
Err_Handler:
    HandleError "pgReports.GridClick", grid
End Sub

Private Sub btnOpen_Click()
    On Error GoTo Err_Handler
    If m_picked <> 0 Then OpenMessage m_picked
    Exit Sub
Err_Handler:
    HandleError "pgReports.btnOpen_Click"
End Sub

Private Sub btnExport_Click()
    On Error GoTo Err_Handler
    RptExportCsv m_period
    Exit Sub
Err_Handler:
    HandleError "pgReports.btnExport_Click"
End Sub
