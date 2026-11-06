Attribute VB_Name = "modWebBridge"
Option Compare Database
Option Explicit

' ANCHOR web bridge (spike): entry points the web server calls over COM in a
' hidden Access instance, so a web action runs the SAME VBA routine the Access
' front end runs (modMessages.Acknowledge, modAudit.AuditVerify), instead of
' a PowerShell copy of it.
'
' Every entry point:
'   - takes the caller's person explicitly (the server has already checked
'     their Windows account and ANCHOR profile) and signs them in for this
'     call only (TempVars, the same ones modSignIn sets);
'   - runs with g_Silent = True, so modErrors.HandleError raises instead of
'     showing a message box;
'   - traps every error and returns one JSON string:
'       {"ok":true|false,"errNumber":n,"source":"...","message":"...","value":...}
'     and never shows a dialog.
' The WebTest* routines exist only to prove what happens when that rule is
' broken (an untrapped error, a MsgBox, an Access prompt, a hang).

Private Const BRIDGE_VERSION As String = "0.1.0"

' ---------------------------------------------------------------- results

Private Function JStr(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1): code = AscW(c)
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 10: out = out & "\n"
            Case 13: out = out & "\r"
            Case 9: out = out & "\t"
            Case 0 To 31: out = out & "\u" & Right$("000" & Hex$(code), 4)
            Case Else: out = out & c
        End Select
    Next
    JStr = """" & out & """"
End Function

Private Function OkJson(Optional ByVal valueJson As String = "null") As String
    OkJson = "{""ok"":true,""errNumber"":0,""source"":"""",""message"":"""",""value"":" & valueJson & "}"
End Function

Private Function FailJson(ByVal num As Long, ByVal src As String, ByVal desc As String) As String
    FailJson = "{""ok"":false,""errNumber"":" & num & ",""source"":" & JStr(src) & ",""message"":" & JStr(desc) & ",""value"":null}"
End Function

' Sign the caller in for this call only, as modSignIn does at sign-in.
' Returns "" when they may use ANCHOR, else why not.
Private Function BeginCall(ByVal personId As Long, ByVal personName As String) As String
    g_Silent = True
    DoCmd.SetWarnings False
    TempVars("PersonID") = personId
    TempVars("PersonName") = personName
    TempVars("Username") = personName
    TempVars("AuditDevice") = Environ$("COMPUTERNAME") & " (web)"    ' modAudit: tells web rows apart
    BeginCall = LoadRoles()
End Function

Private Sub EndCall()
    On Error Resume Next
    TempVars("PersonID") = 0
    TempVars("PersonName") = ""
    TempVars("Username") = ""
    TempVars("AuditDevice") = ""
    ClearRolesForBridge
End Sub

Private Sub ClearRolesForBridge()
    On Error Resume Next
    TempVars("ANRoles") = ""
    TempVars("ANIsAdmin") = False
End Sub

' ---------------------------------------------------------------- entry points

' Is the bridge alive, and how is Access running (full or runtime, trusted).
Public Function WebPing() As String
    On Error GoTo Fail
    WebPing = OkJson("{""bridge"":" & JStr(BRIDGE_VERSION) & ",""runtime"":" & IIf(SysCmd(acSysCmdRuntime), "true", "false") & _
        ",""trusted"":" & IIf(CurrentProject.IsTrusted, "true", "false") & ",""accessVersion"":" & JStr(SysCmd(acSysCmdAccessVer)) & _
        ",""dataLinked"":" & IIf(DataLinked(), "true", "false") & "}")
    Exit Function
Fail:
    WebPing = FailJson(Err.Number, Err.Source, Err.Description)
End Function

' modMessages.Acknowledge, unchanged, for personId. Error 5 = no ack needed.
Public Function WebAcknowledge(ByVal personId As Long, ByVal personName As String, ByVal messageId As Long) As String
    Dim why As String
    On Error GoTo Fail
    why = BeginCall(personId, personName)
    If Len(why) > 0 Then WebAcknowledge = FailJson(70, "modWebBridge.WebAcknowledge", why): EndCall: Exit Function
    Acknowledge messageId
    WebAcknowledge = OkJson("{""ackOn"":" & JStr(Format$(Nz(AckOnOf(messageId), ""), "yyyy-mm-dd hh:nn:ss")) & "}")
    EndCall
    Exit Function
Fail:
    WebAcknowledge = FailJson(Err.Number, Err.Source, Err.Description)
    EndCall
End Function

' modMessages.MarkRead, unchanged (first open audits the view).
Public Function WebMarkRead(ByVal personId As Long, ByVal personName As String, ByVal messageId As Long) As String
    Dim why As String
    On Error GoTo Fail
    why = BeginCall(personId, personName)
    If Len(why) > 0 Then WebMarkRead = FailJson(70, "modWebBridge.WebMarkRead", why): EndCall: Exit Function
    MarkRead messageId
    WebMarkRead = OkJson()
    EndCall
    Exit Function
Fail:
    WebMarkRead = FailJson(Err.Number, Err.Source, Err.Description)
    EndCall
End Function

' modAudit.AuditVerify, unchanged: "" problem means the chain holds.
Public Function WebAuditVerify() As String
    Dim problem As String, n As Long
    On Error GoTo Fail
    g_Silent = True
    problem = AuditVerify(n)
    WebAuditVerify = OkJson("{""ok"":" & IIf(Len(problem) = 0, "true", "false") & ",""checked"":" & n & ",""problem"":" & JStr(problem) & "}")
    Exit Function
Fail:
    WebAuditVerify = FailJson(Err.Number, Err.Source, Err.Description)
End Function

' ---------------------------------------------------------------- spike tests only

' A trapped error: comes back as JSON.
Public Function WebTestTrapped() As String
    On Error GoTo Fail
    g_Silent = True
    Err.Raise 5, "modWebBridge.WebTestTrapped", "A trapped test error, relayed as JSON."
    Exit Function
Fail:
    WebTestTrapped = FailJson(Err.Number, Err.Source, Err.Description)
End Function

' An untrapped run-time error: Access would show the VBA error box.
Public Function WebTestUntrapped() As String
    Dim x As Long
    x = 1 / 0
    WebTestUntrapped = OkJson()
End Function

' A MsgBox a routine forgot to guard.
Public Function WebTestMsgBox() As String
    MsgBox "A message box nobody can see.", vbExclamation, "ANCHOR"
    WebTestMsgBox = OkJson()
End Function

' An Access-generated prompt: a query naming a field that doesn't exist
' makes Access ask "Enter Parameter Value" (SetWarnings doesn't stop it;
' the update-confirmation prompt used first never appears in a hidden
' instance). It changes nothing.
Public Function WebTestAccessDialog() As String
    DoCmd.RunSQL "UPDATE tblSettings SET SettingValue=SettingValue WHERE SettingKey=[No such field]"
    WebTestAccessDialog = OkJson()
End Function

' A hang with no dialog: busy for seconds without yielding.
Public Function WebTestSlow(ByVal seconds As Long) As String
    Dim t As Single
    t = Timer
    Do While Timer - t < seconds
    Loop
    WebTestSlow = OkJson()
End Function
