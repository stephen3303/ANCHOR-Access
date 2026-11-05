Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the run-time UI hooks the kit calls on the
' app: ShowTip for the "?" marks, BadgeColors and StatusVariant for the pills
' in BuildGrid tables, and BandLabels/BandBack, which the kit's BandCF needs
' to compile (the app shows no bands).
'
' Words are mapped to pill variants here, never to colors, so every palette
' keeps them readable. Precedence: FLASH destructive, IMMEDIATE warning,
' PRIORITY info, ROUTINE secondary. Lifecycle: ACTIVE success, AMENDED and
' CORRECTED info, SUPERSEDED warning, CANCELLED destructive, EXPIRED
' secondary. Subscriptions: Active, Paused, Pending Approval, Denied.
' Audit results: Success, Warning, Failed. Admin and Profile: Granted
' success; Retired, Deactivated and Not granted secondary. Acknowledgement: Ack due,
' Acknowledged, Overdue. Classification: UNCLASSIFIED success, CUI info,
' CONFIDENTIAL and SECRET destructive. Matching ignores case (Option Compare
' Database), so ACTIVE and Active are one word.
'
' At the end, the helpers every page shares: GridSelect (select a table's
' row by id), OpenArgText and OpenArgLong (a dialog's OpenArgs), and
' ClipText (a name cut short for a one-line alert).
'
' Owner: the app core. Depends on: modErrors, modGlobals; the kit's
' modDarkMode.

' Click on a "?" help mark (TitleTip): show its full explanation.
Public Function ShowTip(frm As Form, ByVal nm As String)
    On Error GoTo Err_Handler
    Notify frm(nm).Tag, vbInformation
    Exit Function
Err_Handler:
    HandleError "modUI.ShowTip", nm
End Function

' Badge variant -> Array(back, fore).
Public Function BadgeColors(ByVal vrnt As String) As Variant
    Select Case vrnt
        Case "success": BadgeColors = Array(Th(CLR_SUCCESS_BG), Th(CLR_SUCCESS_FG))
        Case "warning": BadgeColors = Array(Th(CLR_WARNING_BG), Th(CLR_WARNING_FG))
        Case "info": BadgeColors = Array(Th(CLR_INFO_BG), Th(CLR_INFO_FG))
        Case "destructive": BadgeColors = Array(Th(CLR_DESTRUCTIVE), Th(CLR_WHITE))
        Case "default": BadgeColors = Array(Th(CLR_PRIMARY), Th(CLR_WHITE))
        Case "outline": BadgeColors = Array(Th(CLR_SURFACE), Th(CLR_FG))
        Case Else: BadgeColors = Array(Th(CLR_SECONDARY_BG), Th(CLR_SECONDARY_FG))
    End Select
End Function

' Badge variant for a status (pill columns in BuildGrid), or "" for no pill.
' It runs on every repaint, also for a row deleted since the list was last
' requeried, whose status can't be read: that row gets no pill.
Public Function StatusVariant(ByVal status As Variant) As String
    On Error GoTo Gone
    Select Case Nz(status, "")
        Case "FLASH", "CANCELLED", "Denied", "Failed", "Overdue", "CONFIDENTIAL", "SECRET", "Not sealed", "Broken": StatusVariant = "destructive"
        Case "IMMEDIATE", "SUPERSEDED", "Pending Approval", "Pending", "Warning", "Ack due", "Required", "In Progress", "Monitoring": StatusVariant = "warning"
        Case "PRIORITY", "AMENDED", "CORRECTED", "CUI", "New", "Unread", "Open", "Mandatory", "Real-time": StatusVariant = "info"
        Case "ACTIVE", "Success", "Acknowledged", "Approved", "UNCLASSIFIED", "Done", "Sealed", "Verified", "Read", "Granted": StatusVariant = "success"
        Case "ROUTINE", "EXPIRED", "Paused", "Not Started", "Off", "Daily Digest", "Weekly", "Closed", "Retired", "Deactivated", "Not granted": StatusVariant = "secondary"
    End Select
    Exit Function
Gone:
    StatusVariant = ""
End Function

' Band pill colors (BandCF). The app shows no bands; these keep the kit's
' BandCF compiling.
Public Function BandLabels() As Variant
    BandLabels = Array("Above", "At", "Below")
End Function

Public Function BandBack(ByVal label As String) As Long
    Select Case label
        Case "Above": BandBack = CLR_BAND_ABOVE
        Case "At": BandBack = CLR_BAND_AT
        Case "Below": BandBack = CLR_BAND_BELOW
        Case Else: BandBack = CLR_BAND_NE
    End Select
End Function

' ---------------------------------------------------------------------------
' Shared page helpers
' ---------------------------------------------------------------------------

' Select the row of table sf (a subGrid subform) whose id column is id, as a
' click on it would: the row's Bookmark and txtSel, then a Recalc so the
' selection mark follows. Every page's list and table selects rows this way.
'   sf       the subform control on the page
'   id       the row's id: a number, or text (help topic keys, through Q);
'            Null, 0 or "" selects no row
'   requery  True to reload the table's rows first
' Returns True when the row is listed and now selected. False leaves no row
' selected (txtSel Null).
' Errors: none of its own; a subform without txtSel raises Access's error.
Public Function GridSelect(sf As Control, ByVal id As Variant, Optional ByVal requery As Boolean = False) As Boolean
    Dim g As Form, crit As String
    Set g = sf.Form
    If requery Then g.Requery
    g.txtSel = Null
    If IsNull(id) Then
        crit = ""
    ElseIf VarType(id) = vbString Then
        If Len(id) > 0 Then crit = "id=" & Q(id)
    ElseIf id <> 0 Then
        crit = "id=" & CLng(id)
    End If
    If Len(crit) > 0 Then
        With g.RecordsetClone
            .FindFirst crit
            If Not .NoMatch Then
                g.Bookmark = .Bookmark
                g.txtSel = id
                GridSelect = True
            End If
        End With
    End If
    g.Recalc
End Function

' The text after name= in frm's OpenArgs ("id=3;total=4"), "" when it isn't
' there. The name matches whatever its case.
Public Function OpenArgText(frm As Form, ByVal name As String) As String
    Dim part As Variant
    For Each part In Split(Nz(frm.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then OpenArgText = Mid$(part, Len(name) + 2)
    Next
End Function

' The whole number after name= in frm's OpenArgs, 0 when it isn't there or
' isn't digits only (so "1e3", "&H10" or "1.5" aren't read as numbers).
Public Function OpenArgLong(frm As Form, ByVal name As String) As Long
    Dim v As String
    v = OpenArgText(frm, name)
    If Len(v) > 0 And Len(v) <= 9 And Not v Like "*[!0-9]*" Then OpenArgLong = CLng(v)
End Function

' Text s cut to at most n characters, ending in an ellipsis when cut, so a
' long name can't push an alert's line past what it shows.
Public Function ClipText(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Or n < 2 Then ClipText = s Else ClipText = RTrim$(Left$(s, n - 1)) & ChrW(8230)
End Function
