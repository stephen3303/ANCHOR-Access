Option Compare Database
Option Explicit

' ANCHOR (Access edition) - the audit trail: append-only and hash chained.
'
' Every audited action adds one tblAudit row (AuditLog): when, who, the
' action, what it was done to, the result and the details. Each row is then
' sealed: RowHash = SHA-256 of the previous row's RowHash and this row's own
' columns (RowText), PrevHash = the previous row's RowHash. A row changed or
' removed after sealing breaks the chain from that row on, which Verify on the
' Audit page finds (AuditVerify).
'
' Rows are sealed in AuditID order by AuditSealPending, never as they are
' added, so two people adding rows at the same moment can't both chain onto
' the same previous row: whoever seals next seals both, in order. Sealing is
' deterministic, so two copies sealing at once write the same hashes. A new
' data file's seeded rows are sealed by the first front end to open it
' (modSignIn.AppStartUp).
'
' The app never updates or deletes a sealed row. Nothing stops a person with
' the data file open in Access from doing it; the chain shows that they did.
'
' Owner: the audit trail. Depends on: modErrors, modGlobals, modSignIn; the
' kit's modCrypto (Sha256Hex), modKitLocks (KitExecute), modPersonnel.

Private Const GENESIS As String = "ANCHOR-AUDIT-GENESIS"

' Add one audit row for the signed-in person and seal it. Device is the
' computer's name; Location is the person's unit (the web app's audit shows
' the unit there).
'   action    what was done, a short word or two: View, Acknowledge, Search, Export
'   resource  what it was done to: "Message (COMSUBLANT 078/26)"
'   details   one or two sentences
'   classif   the resource's classification, UNCLASSIFIED by default
'   result    Success, Failed or Warning
' Errors: the engine's and KitExecute's. Callers that only record (a view, a
'   search) use AuditTry, which logs a failure and carries on.
Public Sub AuditLog(ByVal action As String, ByVal resource As String, ByVal details As String, _
        Optional ByVal classif As String = "UNCLASSIFIED", Optional ByVal result As String = "Success")
    Dim pid As Long, who As String, dev As String
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    who = Nz(TempVars("Username"), "")
    If Len(who) = 0 Then who = SignedInName()
    dev = Nz(TempVars("AuditDevice"), "")          ' the web bridge sets "<computer> (web)"
    If Len(dev) = 0 Then dev = Environ$("COMPUTERNAME")
    KitExecute "INSERT INTO tblAudit (LoggedAt, PersonID, UserName, Action, Resource, Classification, Result, Device, Location, Details) " & _
        "VALUES (Now(), " & IIf(pid = 0, "Null", pid) & ", " & Q(Left$(who, 100)) & ", " & Q(Left$(action, 40)) & ", " & _
        Q(Left$(resource, 150)) & ", " & Q(Left$(classif, 20)) & ", " & Q(Left$(result, 10)) & ", " & _
        Q(Left$(dev, 64)) & ", " & Q(Left$(PersonUnit(pid), 64)) & ", " & Q(details) & ")"
    AuditSealPending
End Sub

' AuditLog for an action that only records (a view, a search, a print): a
' failure is logged and the action goes on.
Public Sub AuditTry(ByVal action As String, ByVal resource As String, ByVal details As String, _
        Optional ByVal classif As String = "UNCLASSIFIED", Optional ByVal result As String = "Success")
    On Error GoTo Fail
    AuditLog action, resource, details, classif, result
    Exit Sub
Fail:
    LogError "modAudit.AuditTry", action & " " & resource
End Sub

' Seal every row with no RowHash, in AuditID order, onto the last sealed row
' before it. Returns how many rows were sealed.
' Errors: the engine's.
Public Function AuditSealPending() As Long
    Dim db As DAO.Database, rs As DAO.Recordset, prev As String, firstId As Long, n As Long
    ' Recordsets, not DMin/DLookup: a domain function reads outside the
    ' caller's transaction, so it can't see a row AuditLog added inside one
    ' (Create draft, Acknowledge), and that row would stay unsealed.
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Min(AuditID) AS FirstID FROM tblAudit WHERE RowHash Is Null", dbOpenSnapshot)
    firstId = Nz(rs!FirstID, 0)
    rs.Close
    If firstId = 0 Then Exit Function
    prev = GENESIS
    Set rs = db.OpenRecordset("SELECT TOP 1 RowHash FROM tblAudit WHERE AuditID<" & firstId & " ORDER BY AuditID DESC", dbOpenSnapshot)
    If Not rs.EOF Then prev = Nz(rs!RowHash, GENESIS)
    rs.Close
    Set rs = db.OpenRecordset("SELECT * FROM tblAudit WHERE AuditID>=" & firstId & " ORDER BY AuditID", dbOpenDynaset)
    Do While Not rs.EOF
        If IsNull(rs!RowHash) Then
            rs.Edit
            rs!PrevHash = prev
            rs!RowHash = Sha256Hex(prev & "|" & RowText(rs))
            rs.Update
            n = n + 1
        End If
        prev = rs!RowHash
        rs.MoveNext
    Loop
    rs.Close
    AuditSealPending = n
End Function

' Recompute the whole chain. Returns "" when every row matches, else the
' first row that doesn't and why, for the Audit page to show.
' Errors: the engine's.
Public Function AuditVerify(Optional ByRef nChecked As Long) As String
    Dim rs As DAO.Recordset, prev As String, want As String
    nChecked = 0
    prev = GENESIS
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblAudit ORDER BY AuditID", dbOpenSnapshot)
    Do While Not rs.EOF
        nChecked = nChecked + 1
        If IsNull(rs!RowHash) Then
            AuditVerify = "Entry " & rs!AuditID & " isn't sealed yet."
            Exit Do
        End If
        If Nz(rs!PrevHash, "") <> prev Then
            AuditVerify = "Entry " & rs!AuditID & " doesn't follow the entry before it: an entry was removed or changed before it."
            Exit Do
        End If
        want = Sha256Hex(prev & "|" & RowText(rs))
        If want <> rs!RowHash Then
            AuditVerify = "Entry " & rs!AuditID & " was changed after it was recorded."
            Exit Do
        End If
        prev = rs!RowHash
        rs.MoveNext
    Loop
    rs.Close
End Function

' The columns a row's hash covers, in a fixed order and format.
Private Function RowText(rs As DAO.Recordset) As String
    RowText = rs!AuditID & "|" & Format$(rs!LoggedAt, "yyyy-mm-dd hh:nn:ss") & "|" & Nz(rs!PersonID, "") & "|" & _
        Nz(rs!UserName, "") & "|" & Nz(rs!Action, "") & "|" & Nz(rs!Resource, "") & "|" & Nz(rs!Classification, "") & "|" & _
        Nz(rs!Result, "") & "|" & Nz(rs!Device, "") & "|" & Nz(rs!Location, "") & "|" & Nz(rs!Details, "")
End Function

' The person's unit from their ANCHOR profile, or "".
Public Function PersonUnit(ByVal pid As Long) As String
    If pid = 0 Then Exit Function
    PersonUnit = Nz(DLookup("Unit", "tblPeople", "PersonID=" & pid), "")
End Function
